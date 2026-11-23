/* =========================================================================
   app.js – Catálogo QR · Almacén CFE Zona Campeche (página para GitHub Pages)
   Secciones:
     1. Núcleo: estado, comunicación con la API, catálogo, búsqueda, navegación
     2. Pantallas (VISTAS)
     3. Importación de Excel, plantillas y exportación
     4. Códigos QR, escáner con cámara e impresión de etiquetas
   La dirección del servidor (Apps Script) está en config.js
   ========================================================================= */


/* ######################################################################
   1. NÚCLEO
   ###################################################################### */
/* =========================================================================
   Scripts.html – Núcleo de la app en el navegador
   - Estado global, comunicación con el servidor, caché local
   - Búsqueda y filtros (se hacen aquí, en el teléfono, para que sea rápido)
   - Navegación entre pantallas, avisos y ventanas emergentes
   ========================================================================= */

/** Estado global de la aplicación. */
const E = {
  cat: null,                 // catálogo descargado del servidor
  mat: new Map(),            // serie → material (con existencias y ubicaciones)
  alm: new Map(),            // código → { c, n, u }
  familias: [],              // [{ f, n }] familias con número de materiales
  zonas: [],                 // zonas/bodegas capturadas
  ses: null,                 // sesión { token, usuario, nombre, rol }
  vista: 'inicio',
  params: {},
  filtros: { q: '', fam: [], alm: [], zona: '', exist: '', ubic: '' },
  modoLista: false,
  limite: 40,                // carga progresiva del catálogo
  seleccionando: false,
  seleccion: new Set(),      // series seleccionadas para imprimir etiquetas
  cargandoCatalogo: false,
  errorCatalogo: ''
};

const ROLES = { Consulta: 1, Almacenista: 2, Administrador: 3 };

// -------------------------------------------------------------------------
//  Utilidades
// -------------------------------------------------------------------------

/** Guardado en el teléfono (con protección por si el navegador lo bloquea). */
const Local = {
  get(k) { try { return localStorage.getItem('cfeqr_' + k); } catch (e) { return null; } },
  set(k, v) { try { localStorage.setItem('cfeqr_' + k, v); } catch (e) { /* sin espacio o bloqueado */ } },
  del(k) { try { localStorage.removeItem('cfeqr_' + k); } catch (e) { } }
};

/** Escapa texto para insertarlo en HTML sin riesgos. */
function esc(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Quita acentos y pasa a minúsculas: "Protección" → "proteccion". */
function norm(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
}

/** Igual que en el servidor: "0000158" → "158". */
function normSerie(v) {
  let s = String(v === null || v === undefined ? '' : v).trim().replace(/\s+/g, '');
  if (/^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
  if (/^\d+$/.test(s)) s = s.replace(/^0+(?=\d)/, '');
  return s.toUpperCase();
}

const _fmt = new Intl.NumberFormat('es-MX', { maximumFractionDigits: 3 });
/** 1581 → "1,581" · 135.1 → "135.1" */
function fmt(n) { return _fmt.format(Number(n) || 0); }
function redondear(n) { return Math.round(Number(n) * 1000) / 1000; }

/** "2026-11-21" → "21/11/2026" */
function fmtFecha(iso) {
  if (!iso) return '';
  const p = String(iso).slice(0, 10).split('-');
  return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : iso;
}
function hoyISO() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

const $ = (sel, raiz) => (raiz || document).querySelector(sel);
const $$ = (sel, raiz) => Array.from((raiz || document).querySelectorAll(sel));

/** Ícono por familia de material (se usa cuando no hay foto). */
function iconoFamilia(f) {
  const n = norm(f);
  if (n.includes('herraje')) return 'ti-tool';
  if (n.includes('conductor') || n.includes('cable')) return 'ti-plug';
  if (n.includes('proteccion') || n.includes('seccionamiento')) return 'ti-shield-bolt';
  if (n.includes('conector')) return 'ti-link';
  if (n.includes('transformador')) return 'ti-bolt';
  if (n.includes('poste') || n.includes('estructura')) return 'ti-building-broadcast-tower';
  if (n.includes('aislador')) return 'ti-circle-dot';
  if (n.includes('medicion') || n.includes('acometida')) return 'ti-gauge';
  return 'ti-package';
}

/** "1071 · Proyectos de obra eléctrica" */
function nombreAlmacen(c, soloCodigo) {
  const a = E.alm.get(c);
  if (soloCodigo || !a || !a.n) return c;
  return c + ' · ' + a.n;
}

/** URL de la foto en Drive (miniatura). */
function urlFoto(id, ancho) {
  return 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w' + (ancho || 400);
}

// -------------------------------------------------------------------------
//  Comunicación con el servidor
// -------------------------------------------------------------------------

/** ¿Ya se pegó la URL de la API en config.js? */
function servidorConfigurado() {
  return typeof URL_SERVIDOR === 'string' && /^https:\/\/script\.google\.com\/.+\/exec$/.test(URL_SERVIDOR.trim());
}

/**
 * Llama a una función del servidor (Apps Script) y devuelve una promesa.
 * Ejemplo:  const r = await api('guardarUbicacion', E.ses.token, datos);
 *
 * Se manda como texto plano para evitar la verificación previa de CORS.
 * No se envían cookies, así que NO importa qué cuentas de Google tenga
 * abiertas el usuario en su navegador.
 */
async function api(funcion, ...args) {
  if (!servidorConfigurado()) throw new Error('Falta configurar la dirección del servidor (URL_SERVIDOR en config.js).');
  let respuesta;
  try {
    const r = await fetch(URL_SERVIDOR.trim(), {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ fn: funcion, args }),
      redirect: 'follow',
      credentials: 'omit'
    });
    const texto = await r.text();
    try {
      respuesta = JSON.parse(texto);
    } catch (e) {
      throw new Error('El servidor no respondió correctamente. Revisa que la implementación tenga acceso "Cualquiera".');
    }
  } catch (e) {
    if (e instanceof TypeError) throw new Error('Sin conexión con el servidor. Revisa tu internet e intenta de nuevo.');
    throw e;
  }
  if (!respuesta.ok) {
    const msg = String(respuesta.error || 'Error desconocido').replace(/^(Exception|Error):\s*/i, '');
    if (/sesi[oó]n expir|Inicia sesi[oó]n para/i.test(msg) && E.ses) {
      E.ses = null; Local.del('ses'); pintarUsuario();
    }
    throw new Error(msg);
  }
  return respuesta.data;
}

/** Token de la sesión actual (o null). */
function token() { return E.ses ? E.ses.token : null; }

/** ¿El usuario tiene al menos este rol? */
function tieneRol(min) { return !!(E.ses && (ROLES[E.ses.rol] || 0) >= ROLES[min]); }

// -------------------------------------------------------------------------
//  Catálogo: se descarga una vez, se guarda en el teléfono y se actualiza
// -------------------------------------------------------------------------

/** Prepara el catálogo para buscar rápido. */
function usarCatalogo(c) {
  E.cat = c;
  E.alm = new Map();
  c.alm.forEach(a => E.alm.set(a.c, a));
  E.mat = new Map();
  const conteoFam = {};
  const zonas = new Set();
  c.mats.forEach(m => {
    m.ex = c.ex[m.s] || [];
    m.ub = c.ub[m.s] || [];
    m.total = redondear(m.ex.reduce((t, x) => t + x[1], 0));
    m.nAlm = m.ex.length;
    m.ub.forEach(u => { if (u.z) zonas.add(u.z); });
    m.idx = norm([
      m.s, m.d, m.f,
      m.ex.map(x => x[0] + ' ' + ((E.alm.get(x[0]) || {}).n || '')).join(' '),
      m.ub.map(u => u.t + ' ' + u.ref).join(' '),
      m.dd
    ].join(' '));
    E.mat.set(m.s, m);
    if (m.act) conteoFam[m.f] = (conteoFam[m.f] || 0) + 1;
  });
  E.familias = Object.keys(conteoFam).map(f => ({ f, n: conteoFam[f] })).sort((a, b) => b.n - a.n || a.f.localeCompare(b.f));
  E.zonas = Array.from(zonas).sort();
}

/**
 * 1) Muestra de inmediato el catálogo guardado en el teléfono (si existe).
 * 2) Pide al servidor la versión más reciente y, si cambió, actualiza la pantalla.
 */
async function cargarCatalogo(forzar) {
  if (!E.cat) {
    const guardado = Local.get('catalogo');
    if (guardado) {
      try { usarCatalogo(JSON.parse(guardado)); render(); } catch (e) { Local.del('catalogo'); }
    }
  }
  E.cargandoCatalogo = true;
  try {
    const texto = await api('obtenerCatalogo');
    const nuevo = JSON.parse(texto);
    E.errorCatalogo = '';
    if (forzar || !E.cat || nuevo.v !== E.cat.v || nuevo.generado !== E.cat.generado) {
      usarCatalogo(nuevo);
      Local.set('catalogo', texto);
      refrescarVistaActual();
    }
  } catch (e) {
    E.errorCatalogo = e.message;
    if (!E.cat) render();
    else aviso('Sin conexión: mostrando datos guardados en el teléfono.', 'error');
  } finally {
    E.cargandoCatalogo = false;
  }
}

/** Vuelve a dibujar sin perder lo que se está escribiendo en el buscador. */
function refrescarVistaActual() {
  const v = VISTAS[E.vista];
  if (v && v.refrescar) v.refrescar();
  else render();
}

/** Materiales visibles: los dados de baja solo los ve el Administrador. */
function materialesVisibles() {
  if (!E.cat) return [];
  const admin = tieneRol('Administrador');
  return E.cat.mats.filter(m => m.act || admin);
}

/**
 * Búsqueda inteligente: por serie, nombre, familia, almacén o ubicación.
 * Sin distinguir acentos ni mayúsculas, búsqueda parcial, todas las palabras deben aparecer.
 */
function buscarMateriales(q, lista) {
  const qn = norm(q).trim();
  if (!qn) return lista;
  const toks = qn.split(/\s+/).map(t => /^\d+$/.test(t) ? t.replace(/^0+(?=\d)/, '') : t);
  const qs = norm(normSerie(q));
  const res = [];
  for (const m of lista) {
    if (!toks.every(t => m.idx.includes(t))) continue;
    const s = norm(m.s), d = norm(m.d);
    let p = 5;
    if (s === qs) p = 0;
    else if (s.startsWith(qs)) p = 1;
    else if (d.startsWith(qn)) p = 2;
    else if (d.includes(qn)) p = 3;
    res.push([p, m]);
  }
  res.sort((a, b) => a[0] - b[0] || a[1].d.localeCompare(b[1].d));
  return res.map(x => x[1]);
}

/** Aplica los filtros combinables del catálogo. */
function aplicarFiltros(lista) {
  const f = E.filtros;
  return lista.filter(m => {
    if (f.fam.length && !f.fam.includes(m.f)) return false;
    if (f.alm.length && !m.ex.some(x => f.alm.includes(x[0]))) return false;
    if (f.zona && !m.ub.some(u => u.z === f.zona)) return false;
    if (f.exist === 'con' && !(m.total > 0)) return false;
    if (f.exist === 'sin' && m.total > 0) return false;
    if (f.ubic === 'con' && !m.ub.length) return false;
    if (f.ubic === 'sin' && m.ub.length) return false;
    return true;
  });
}

/** Material exacto por serie (acepta "0000158"). */
function materialPorSerie(serie) {
  return E.mat.get(normSerie(serie)) || null;
}

// -------------------------------------------------------------------------
//  Sesión (usuario + PIN)
// -------------------------------------------------------------------------

function cargarSesionGuardada() {
  const t = Local.get('ses');
  if (t) { try { E.ses = JSON.parse(t); } catch (e) { Local.del('ses'); } }
}

async function verificarSesionGuardada() {
  if (!E.ses) return;
  try {
    const s = await api('verificarSesion', E.ses.token);
    if (!s) { E.ses = null; Local.del('ses'); pintarUsuario(); render(); }
  } catch (e) { /* sin conexión: se conserva hasta que el servidor diga lo contrario */ }
}

function pintarUsuario() {
  const b = $('#btnUsuario');
  if (E.ses) {
    b.innerHTML = '<span class="inicial">' + esc((E.ses.nombre || E.ses.usuario).charAt(0).toUpperCase()) + '</span>';
    b.setAttribute('aria-label', 'Mi cuenta: ' + E.ses.nombre);
  } else {
    b.innerHTML = '<i class="ti ti-login-2"></i>';
    b.setAttribute('aria-label', 'Iniciar sesión');
  }
}

async function cerrarSesion() {
  const t = token();
  E.ses = null; Local.del('ses'); pintarUsuario();
  E.seleccion.clear(); E.seleccionando = false;
  if (t) api('cerrarSesion', t).catch(() => { });
  aviso('Sesión cerrada');
  ir('inicio', {}, true);
}

// -------------------------------------------------------------------------
//  Navegación entre pantallas
//  Se usa el historial del navegador para que el botón "atrás" del celular
//  funcione y para que la URL muestre ?serie=… al ver una ficha.
// -------------------------------------------------------------------------

let _navId = 0;

function ir(vista, params, reemplazar) {
  E.vista = vista;
  E.params = params || {};
  _navId++;
  const estado = { vista, params: E.params };
  const url = vista === 'ficha' && E.params.serie
    ? location.pathname + '?serie=' + encodeURIComponent(E.params.serie)
    : location.pathname;
  try {
    if (reemplazar) history.replaceState(estado, '', url);
    else history.pushState(estado, '', url);
  } catch (e) { /* navegador sin historial */ }
  render();
  window.scrollTo(0, 0);
}

/** Botón ← del encabezado. */
function atras() {
  const id = _navId;
  try { window.history.back(); } catch (e) { }
  // Si el navegador no regresó (por ejemplo, se abrió directo una ficha desde el QR), ir al inicio
  setTimeout(() => { if (id === _navId) ir((VISTAS[E.vista] && VISTAS[E.vista].padre) || 'inicio', {}, true); }, 450);
}

/** Dibuja la pantalla actual. Las pantallas están en Vistas.html. */
function render() {
  const v = VISTAS[E.vista] || VISTAS.inicio;
  const app = $('#app');
  $('#titulo').textContent = typeof v.titulo === 'function' ? v.titulo() : (v.titulo || 'Almacén CFE Zona Campeche');
  $('#btnAtras').hidden = !!v.raiz;
  $$('#nav button').forEach(b => b.classList.toggle('activo', b.dataset.ir === (v.nav || E.vista)));

  if (v.rol && !tieneRol(v.rol)) {
    app.innerHTML = vacioHTML('ti-lock', 'Necesitas iniciar sesión',
      'Esta sección es para el rol ' + esc(v.rol) + (v.rol === 'Almacenista' ? ' o Administrador' : '') + '.',
      '<button class="btn btn-primario" onclick="ir(\'login\', {volver: \'' + esc(E.vista) + '\'})"><i class="ti ti-login-2"></i>Iniciar sesión</button>');
    return;
  }
  if (v.necesitaCatalogo !== false && !E.cat) {
    app.innerHTML = E.errorCatalogo
      ? vacioHTML('ti-wifi-off', 'No se pudo cargar el catálogo', esc(E.errorCatalogo),
        '<button class="btn btn-primario" onclick="reintentarCatalogo()"><i class="ti ti-refresh"></i>Reintentar</button>')
      : '<div class="cargando-pantalla"><div class="spinner"></div><p>Cargando catálogo…</p></div>';
    return;
  }
  v.render(app);
}

function reintentarCatalogo() {
  E.errorCatalogo = '';
  render();
  cargarCatalogo(true);
}

// -------------------------------------------------------------------------
//  Avisos, ventanas emergentes y ayudas de interfaz
// -------------------------------------------------------------------------

let _avisoTimer = null;
function aviso(texto, tipo) {
  const a = $('#aviso');
  a.textContent = texto;
  a.className = 'aviso' + (tipo === 'error' ? ' error' : '');
  a.hidden = false;
  clearTimeout(_avisoTimer);
  _avisoTimer = setTimeout(() => { a.hidden = true; }, tipo === 'error' ? 5000 : 2600);
}

function vacioHTML(icono, titulo, texto, extra) {
  return '<div class="vacio"><i class="ti ' + icono + '"></i><h2>' + esc(titulo) + '</h2>' +
    (texto ? '<p>' + texto + '</p>' : '') + (extra || '') + '</div>';
}

/** Abre una ventana emergente. contenido = HTML. */
function abrirModal(titulo, contenido, alAbrir, clase) {
  const m = $('#modal');
  m.className = 'modal' + (clase ? ' ' + clase : '');
  m.innerHTML = (clase === 'modal-foto' || clase === 'modal-escaner') ? contenido :
    '<div class="modal-caja" role="dialog" aria-modal="true" aria-label="' + esc(titulo) + '">' +
    '<div class="modal-cab"><h2>' + esc(titulo) + '</h2>' +
    '<button class="btn btn-texto btn-icono" onclick="cerrarModal()" aria-label="Cerrar"><i class="ti ti-x"></i></button></div>' +
    contenido + '</div>';
  m.hidden = false;
  m.onclick = (ev) => { if (ev.target === m) cerrarModal(); };
  if (alAbrir) alAbrir(m);
  const primero = m.querySelector('input:not([type=hidden]), select, textarea');
  if (primero && window.innerWidth > 640) primero.focus();
}
function cerrarModal() {
  if (typeof detenerCamara === 'function') detenerCamara();
  const m = $('#modal');
  m.hidden = true; m.innerHTML = '';
}
document.addEventListener('keydown', ev => { if (ev.key === 'Escape' && !$('#modal').hidden) cerrarModal(); });

/** Pide confirmación con una ventana propia (más clara y grande que el "confirm" del navegador). */
function confirmar(titulo, texto, textoBoton, peligro) {
  return new Promise(res => {
    abrirModal(titulo, '<p>' + texto + '</p><div class="modal-pie">' +
      '<button class="btn" id="cfNo">Cancelar</button>' +
      '<button class="btn ' + (peligro ? 'btn-peligro' : 'btn-primario') + '" id="cfSi">' + esc(textoBoton || 'Aceptar') + '</button></div>',
      m => {
        $('#cfNo', m).onclick = () => { cerrarModal(); res(false); };
        $('#cfSi', m).onclick = () => { cerrarModal(); res(true); };
      });
  });
}

/** Desactiva un botón y muestra "cargando" mientras se ejecuta una tarea. */
async function conCarga(boton, tarea) {
  const original = boton ? boton.innerHTML : '';
  if (boton) { boton.disabled = true; boton.innerHTML = '<span class="spinner chico"></span>' + (boton.dataset.cargando || 'Guardando…'); }
  try {
    return await tarea();
  } catch (e) {
    aviso(e.message, 'error');
    throw e;
  } finally {
    if (boton && boton.isConnected) { boton.disabled = false; boton.innerHTML = original; }
  }
}

/** Copia texto al portapapeles (con respaldo para navegadores antiguos). */
function copiar(texto) {
  const respaldo = () => {
    const t = document.createElement('textarea');
    t.value = texto; t.style.position = 'fixed'; t.style.opacity = '0';
    document.body.appendChild(t); t.select();
    try { document.execCommand('copy'); aviso('Copiado: ' + texto); } catch (e) { aviso('No se pudo copiar', 'error'); }
    t.remove();
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(texto).then(() => aviso('Copiado: ' + texto), respaldo);
  } else respaldo();
}

/** Lee los valores de un formulario como objeto { nombre: valor }. */
function leerFormulario(raiz) {
  const o = {};
  $$('[name]', raiz).forEach(el => {
    o[el.name] = el.type === 'checkbox' ? el.checked : el.value;
  });
  return o;
}

/** Recarga el catálogo después de guardar algo. */
async function despuesDeGuardar(mensaje) {
  if (mensaje) aviso(mensaje);
  await cargarCatalogo(true);
}

// -------------------------------------------------------------------------
//  Arranque
// -------------------------------------------------------------------------

function iniciarApp() {
  cargarSesionGuardada();
  pintarUsuario();
  E.modoLista = Local.get('modoLista') === '1';

  // Navegación
  $$('#nav button').forEach(b => b.addEventListener('click', () => {
    if (b.dataset.accion === 'escanear') return abrirEscaner();
    if (b.dataset.ir === 'catalogo') E.limite = 40;
    ir(b.dataset.ir);
  }));
  $('#btnAtras').addEventListener('click', atras);
  $('#btnUsuario').addEventListener('click', () => ir(E.ses ? 'mas' : 'login', E.ses ? {} : { volver: E.vista }));
  $('#inputEscaner').addEventListener('change', alElegirFotoQR);

  // Botón "atrás" del celular o del navegador
  window.addEventListener('popstate', ev => {
    const st = ev.state;
    _navId++;
    const serie = new URLSearchParams(location.search).get('serie');
    if (st && st.vista) { E.vista = st.vista; E.params = st.params || {}; }
    else if (serie) { E.vista = 'ficha'; E.params = { serie: normSerie(serie) }; }
    else { E.vista = 'inicio'; E.params = {}; }
    cerrarModal();
    render();
  });

  if (!servidorConfigurado()) {
    $('#app').innerHTML = vacioHTML('ti-plug-connected-x', 'Falta conectar el servidor',
      'Abre el archivo <b>config.js</b> y pega la URL de Apps Script que termina en <b>/exec</b> (paso 3 de la guía).');
    return;
  }

  // ¿Se abrió desde un QR?  (…github.io/catalogo/?serie=158)
  const serie = new URLSearchParams(location.search).get('serie');
  if (serie) ir('ficha', { serie: normSerie(serie) }, true);
  else ir('inicio', {}, true);
  cargarCatalogo();
  verificarSesionGuardada();
}


/* ######################################################################
   2. PANTALLAS
   ###################################################################### */
/* =========================================================================
   Vistas.html – Pantallas de la aplicación
   Cada pantalla es un objeto en VISTAS con:
     titulo, raiz (sin botón ←), nav (botón resaltado), rol (mínimo),
     render(app) y opcionalmente refrescar()
   ========================================================================= */

const VISTAS = {};

// -------------------------------------------------------------------------
//  Piezas reutilizables
// -------------------------------------------------------------------------

/** Tarjeta de un material (catálogo, búsqueda, reportes). Un material = una tarjeta. */
function tarjetaMaterial(m) {
  const cero = !(m.total > 0);
  const img = m.foto
    ? '<img src="' + esc(urlFoto(m.foto, 200)) + '" alt="" loading="lazy">'
    : '<i class="ti ' + iconoFamilia(m.f) + '"></i>';
  const chk = E.seleccionando
    ? '<input type="checkbox" class="mat-check" data-sel="' + esc(m.s) + '"' + (E.seleccion.has(m.s) ? ' checked' : '') + ' aria-label="Seleccionar ' + esc(m.d) + '">'
    : '';
  const ubic = m.ub.length ? esc(m.ub.map(u => u.t).join(' / ')) : '<span class="texto-3">Sin ubicación</span>';
  return '<div class="mat" data-serie="' + esc(m.s) + '" role="button" tabindex="0">' + chk +
    '<div class="mat-img' + (cero ? ' cero' : '') + '">' + img + '</div>' +
    '<div class="crece">' +
    '<div class="mat-nombre' + (m.act ? '' : ' cero') + '">' + esc(m.d) + (m.act ? '' : ' <span class="insignia peligro">Baja</span>') + '</div>' +
    '<div class="mat-cant num' + (cero ? ' cero' : '') + '">' + fmt(m.total) + ' <small>' + esc(m.u) + '</small></div>' +
    '<div class="mat-meta"><span class="serie">Serie ' + esc(m.s) + '</span> · ' + esc(m.f) + ' · ' + m.nAlm + (m.nAlm === 1 ? ' almacén' : ' almacenes') + '</div>' +
    '<div class="mat-meta"><i class="ti ti-map-pin"></i> ' + ubic + '</div>' +
    '</div></div>';
}

/** Activa los clics en una lista de tarjetas (abrir ficha o seleccionar). */
function activarTarjetas(contenedor) {
  contenedor.addEventListener('click', ev => {
    const chk = ev.target.closest('[data-sel]');
    if (chk) {
      if (chk.checked) E.seleccion.add(chk.dataset.sel); else E.seleccion.delete(chk.dataset.sel);
      pintarBarraSeleccion();
      return;
    }
    const t = ev.target.closest('.mat');
    if (!t) return;
    if (E.seleccionando) {
      const c = $('[data-sel]', t);
      if (c) c.click(); // marca/desmarca y vuelve a entrar a este mismo evento por la casilla
      return;
    }
    ir('ficha', { serie: t.dataset.serie });
  });
  contenedor.addEventListener('keydown', ev => {
    if (ev.key === 'Enter' && ev.target.classList.contains('mat')) ev.target.click();
  });
}

/** Caja de búsqueda grande. */
function buscadorHTML(id, valor, placeholder) {
  return '<form class="buscador" id="' + id + 'Form" role="search" onsubmit="return false">' +
    '<i class="ti ti-search"></i>' +
    '<input id="' + id + '" type="search" inputmode="search" autocomplete="off" enterkeyhint="search" ' +
    'placeholder="' + esc(placeholder || 'Serie, nombre o ubicación') + '" value="' + esc(valor || '') + '">' +
    '<button type="button" class="btn btn-texto btn-icono" id="' + id + 'Limpiar" aria-label="Limpiar"' + (valor ? '' : ' hidden') + '><i class="ti ti-x"></i></button>' +
    '</form>';
}

/**
 * Conecta un buscador: resultados mientras se escribe y, al pulsar Enter,
 * si el texto es un número de serie exacto se abre directo la ficha.
 */
function conectarBuscador(id, alEscribir) {
  const input = $('#' + id);
  const limpiar = $('#' + id + 'Limpiar');
  let t = null;
  input.addEventListener('input', () => {
    limpiar.hidden = !input.value;
    clearTimeout(t);
    t = setTimeout(() => alEscribir(input.value), 120);
  });
  $('#' + id + 'Form').addEventListener('submit', ev => {
    ev.preventDefault();
    const m = materialPorSerie(input.value);
    if (m && (m.act || tieneRol('Administrador'))) { input.blur(); ir('ficha', { serie: m.s }); }
    else alEscribir(input.value);
  });
  limpiar.addEventListener('click', () => { input.value = ''; limpiar.hidden = true; alEscribir(''); input.focus(); });
}

function fechaCorteHTML() {
  return E.cat && E.cat.corte
    ? 'Existencias al <b>' + fmtFecha(E.cat.corte) + '</b>'
    : '<span class="insignia alerta"><i class="ti ti-alert-triangle"></i> Aún no se importa el Excel de SAP</span>';
}

// -------------------------------------------------------------------------
//  INICIO
// -------------------------------------------------------------------------

VISTAS.inicio = {
  titulo: 'Almacén CFE Zona Campeche',
  raiz: true,
  render(app) {
    const total = materialesVisibles().filter(m => m.act).length;
    app.innerHTML =
      '<div class="pila">' +
      buscadorHTML('qInicio', '') +
      '<div id="resInicio"></div>' +
      '<button class="btn btn-primario btn-gigante" onclick="abrirEscaner()"><i class="ti ti-qrcode"></i>Escanear QR</button>' +
      '<div class="fila-sep pequeno texto-2"><span>' + fechaCorteHTML() + '</span><span class="num">' + fmt(total) + ' materiales</span></div>' +
      (tieneRol('Administrador')
        ? '<button class="tarjeta clic fila" style="text-align:left;border-style:dashed" onclick="ir(\'importar\')"><i class="ti ti-file-upload" style="font-size:26px;color:var(--cfe-verde)"></i><div class="crece"><b>Importar Excel del día</b><div class="texto-3">Actualiza las existencias desde SAP</div></div><i class="ti ti-chevron-right texto-3"></i></button>'
        : '') +
      '<div class="sec-titulo">Familias</div>' +
      (E.familias.length
        ? '<div class="familias">' + E.familias.map(x =>
          '<button class="familia" data-fam="' + esc(x.f) + '"><i class="ti ' + iconoFamilia(x.f) + '"></i>' +
          '<span class="nombre">' + esc(x.f) + '</span><span class="cuenta num">' + x.n + ' materiales</span></button>').join('') + '</div>'
        : vacioHTML('ti-packages', 'El catálogo está vacío', 'El administrador debe importar el Excel de SAP para ver los materiales.')) +
      '</div>';

    $$('.familia', app).forEach(b => b.addEventListener('click', () => {
      E.filtros = { q: '', fam: [b.dataset.fam], alm: [], zona: '', exist: '', ubic: '' };
      E.limite = 40;
      ir('catalogo');
    }));
    const res = $('#resInicio');
    activarTarjetas(res);
    conectarBuscador('qInicio', q => {
      if (!q.trim()) { res.innerHTML = ''; return; }
      const lista = buscarMateriales(q, materialesVisibles());
      res.innerHTML = lista.length
        ? '<div class="resultados">' + lista.slice(0, 6).map(tarjetaMaterial).join('') + '</div>' +
        (lista.length > 6 ? '<button class="btn btn-bloque" style="margin-top:10px" id="verTodos">Ver los ' + fmt(lista.length) + ' resultados</button>' : '')
        : vacioHTML('ti-search-off', 'Sin resultados', 'No hay materiales que coincidan con “' + esc(q) + '”.');
      const vt = $('#verTodos');
      if (vt) vt.onclick = () => ir('buscar', { q });
    });
  },
  refrescar() { if (!$('#qInicio') || !$('#qInicio').value) render(); }
};

// -------------------------------------------------------------------------
//  BUSCAR
// -------------------------------------------------------------------------

VISTAS.buscar = {
  titulo: 'Buscar',
  raiz: true,
  render(app) {
    const q = E.params.q || '';
    app.innerHTML = '<div class="pila">' + buscadorHTML('qBuscar', q) +
      '<div class="texto-3">Escribe parte del nombre, el número de serie (con o sin ceros), la familia, el almacén o la ubicación. Pulsa Enter con una serie exacta para abrir su ficha.</div>' +
      '<div id="resBuscar"></div></div>';
    activarTarjetas($('#resBuscar'));
    conectarBuscador('qBuscar', valor => { E.params.q = valor; this.pintar(valor, 50); });
    this.pintar(q, 50);
    const input = $('#qBuscar');
    if (!q) input.focus();
  },
  pintar(q, limite) {
    const res = $('#resBuscar');
    if (!res) return;
    if (!q.trim()) { res.innerHTML = ''; return; }
    const lista = buscarMateriales(q, materialesVisibles());
    res.innerHTML = lista.length
      ? '<div class="texto-3" style="margin-bottom:8px">' + fmt(lista.length) + ' resultados</div><div class="resultados">' +
      lista.slice(0, limite).map(tarjetaMaterial).join('') + '</div>' +
      (lista.length > limite ? '<button class="btn btn-bloque" style="margin-top:10px" id="masBuscar">Mostrar más</button>' : '')
      : vacioHTML('ti-search-off', 'Sin resultados', 'Prueba con otra palabra o revisa el número de serie.');
    const b = $('#masBuscar');
    if (b) b.onclick = () => this.pintar(q, limite + 50);
  },
  refrescar() { const i = $('#qBuscar'); if (i) this.pintar(i.value, 50); else render(); }
};

// -------------------------------------------------------------------------
//  CATÁLOGO (tarjetas, filtros combinables, lista compacta, carga progresiva)
// -------------------------------------------------------------------------

VISTAS.catalogo = {
  titulo: 'Catálogo',
  raiz: true,
  render(app) {
    const f = E.filtros;
    const nFiltros = f.fam.length + f.alm.length + (f.zona ? 1 : 0) + (f.exist ? 1 : 0) + (f.ubic ? 1 : 0);
    app.innerHTML = '<div class="pila">' +
      buscadorHTML('qCat', f.q, 'Buscar dentro del catálogo') +
      '<div class="fila" style="flex-wrap:wrap">' +
      '<button class="btn btn-chico" id="btnFiltros"><i class="ti ti-adjustments-horizontal"></i>Filtros' + (nFiltros ? ' (' + nFiltros + ')' : '') + '</button>' +
      '<button class="btn btn-chico" id="btnModo" aria-label="Cambiar vista"><i class="ti ' + (E.modoLista ? 'ti-layout-grid' : 'ti-list') + '"></i>' + (E.modoLista ? 'Tarjetas' : 'Lista') + '</button>' +
      '<button class="btn btn-chico" id="btnExportar"><i class="ti ti-file-spreadsheet"></i>Exportar</button>' +
      (tieneRol('Administrador') ? '<button class="btn btn-chico" id="btnSelec"><i class="ti ti-checkbox"></i>' + (E.seleccionando ? 'Terminar selección' : 'Seleccionar') + '</button>' : '') +
      '</div>' +
      '<div class="chips" id="chips"></div>' +
      '<div class="texto-3" id="cuentaCat"></div>' +
      '<div id="resCat"></div>' +
      '<div id="barraSel"></div>' +
      '</div>';

    conectarBuscador('qCat', q => { E.filtros.q = q; E.limite = 40; this.pintar(); });
    $('#btnFiltros').onclick = abrirFiltros;
    $('#btnModo').onclick = () => { E.modoLista = !E.modoLista; Local.set('modoLista', E.modoLista ? '1' : ''); render(); };
    $('#btnExportar').onclick = () => exportarMateriales(this.lista(), 'Catalogo');
    const bs = $('#btnSelec');
    if (bs) bs.onclick = () => { E.seleccionando = !E.seleccionando; render(); };
    activarTarjetas($('#resCat'));
    this.pintar();
  },
  /** Lista filtrada actual. */
  lista() {
    return aplicarFiltros(buscarMateriales(E.filtros.q, materialesVisibles()));
  },
  pintar() {
    const res = $('#resCat');
    if (!res) return;
    const lista = this.lista();
    pintarChips();
    $('#cuentaCat').textContent = fmt(lista.length) + (lista.length === 1 ? ' material' : ' materiales') + ' · ' + (E.cat.corte ? 'existencias al ' + fmtFecha(E.cat.corte) : 'sin corte de SAP');
    if (!lista.length) {
      res.innerHTML = vacioHTML('ti-filter-off', 'Ningún material coincide', 'Quita algún filtro para ver más resultados.',
        '<button class="btn" onclick="limpiarFiltros()">Quitar filtros</button>');
      return;
    }
    res.innerHTML = '<div class="resultados' + (E.modoLista ? ' lista' : '') + '">' +
      lista.slice(0, E.limite).map(tarjetaMaterial).join('') + '</div>' +
      (lista.length > E.limite ? '<button class="btn btn-bloque" id="masCat" style="margin-top:10px">Cargar más (' + fmt(lista.length - E.limite) + ' restantes)</button>' : '');
    const mas = $('#masCat');
    if (mas) {
      mas.onclick = () => { E.limite += 60; this.pintar(); };
      // Carga progresiva automática al llegar al final
      if ('IntersectionObserver' in window) {
        const obs = new IntersectionObserver(ent => {
          if (ent[0].isIntersecting) { obs.disconnect(); E.limite += 60; this.pintar(); }
        }, { rootMargin: '300px' });
        obs.observe(mas);
      }
    }
    pintarBarraSeleccion();
  },
  refrescar() { if ($('#resCat')) this.pintar(); else render(); }
};

function pintarChips() {
  const f = E.filtros, c = $('#chips');
  if (!c) return;
  const chips = [];
  f.fam.forEach(x => chips.push(['fam', x, x]));
  f.alm.forEach(x => chips.push(['alm', x, 'Almacén ' + nombreAlmacen(x)]));
  if (f.zona) chips.push(['zona', f.zona, 'Zona ' + f.zona]);
  if (f.exist) chips.push(['exist', f.exist, f.exist === 'con' ? 'Con existencia' : 'Sin existencia']);
  if (f.ubic) chips.push(['ubic', f.ubic, f.ubic === 'con' ? 'Con ubicación' : 'Sin ubicación']);
  c.innerHTML = chips.map(x => '<button class="chip" data-k="' + x[0] + '" data-v="' + esc(x[1]) + '">' + esc(x[2]) + ' <i class="ti ti-x"></i></button>').join('') +
    (chips.length > 1 ? '<button class="chip neutro" onclick="limpiarFiltros()">Quitar todos</button>' : '');
  $$('.chip[data-k]', c).forEach(b => b.onclick = () => {
    const k = b.dataset.k, v = b.dataset.v;
    if (Array.isArray(E.filtros[k])) E.filtros[k] = E.filtros[k].filter(x => x !== v); else E.filtros[k] = '';
    E.limite = 40;
    render();
  });
}

function limpiarFiltros() {
  E.filtros = { q: E.filtros.q, fam: [], alm: [], zona: '', exist: '', ubic: '' };
  E.limite = 40;
  render();
}

function abrirFiltros() {
  const f = E.filtros;
  const opcionesRadio = (nombre, valor, opciones) => '<div class="chips">' + opciones.map(o =>
    '<label class="check" style="min-height:40px"><input type="radio" name="' + nombre + '" value="' + o[0] + '"' + (valor === o[0] ? ' checked' : '') + '>' + o[1] + '</label>').join('') + '</div>';
  const almacenes = Array.from(E.alm.values()).sort((a, b) => a.c.localeCompare(b.c));
  abrirModal('Filtros',
    '<div class="pila">' +
    '<div class="sec-titulo">Familia</div>' +
    E.familias.map(x => '<label class="check"><input type="checkbox" name="fam" value="' + esc(x.f) + '"' + (f.fam.includes(x.f) ? ' checked' : '') + '>' + esc(x.f) + ' <span class="texto-3">(' + x.n + ')</span></label>').join('') +
    '<div class="sec-titulo">Almacén SAP</div>' +
    almacenes.map(a => '<label class="check"><input type="checkbox" name="alm" value="' + esc(a.c) + '"' + (f.alm.includes(a.c) ? ' checked' : '') + '><span><b>' + esc(a.c) + '</b> <span class="texto-2">' + esc(a.n) + '</span></span></label>').join('') +
    '<div class="sec-titulo">Zona / bodega</div>' +
    '<select class="entrada" name="zona"><option value="">Todas</option>' + E.zonas.map(z => '<option' + (f.zona === z ? ' selected' : '') + '>' + esc(z) + '</option>').join('') + '</select>' +
    '<div class="sec-titulo">Existencia</div>' + opcionesRadio('exist', f.exist, [['', 'Todas'], ['con', 'Con existencia'], ['sin', 'Sin existencia']]) +
    '<div class="sec-titulo">Ubicación física</div>' + opcionesRadio('ubic', f.ubic, [['', 'Todas'], ['con', 'Con ubicación'], ['sin', 'Sin ubicación']]) +
    '</div>' +
    '<div class="modal-pie"><button class="btn" id="fLimpiar">Limpiar</button><button class="btn btn-primario" id="fAplicar">Ver resultados</button></div>',
    m => {
      $('#fLimpiar', m).onclick = () => { cerrarModal(); limpiarFiltros(); };
      $('#fAplicar', m).onclick = () => {
        E.filtros.fam = $$('input[name=fam]:checked', m).map(i => i.value);
        E.filtros.alm = $$('input[name=alm]:checked', m).map(i => i.value);
        E.filtros.zona = $('select[name=zona]', m).value;
        E.filtros.exist = ($('input[name=exist]:checked', m) || {}).value || '';
        E.filtros.ubic = ($('input[name=ubic]:checked', m) || {}).value || '';
        E.limite = 40;
        cerrarModal();
        render();
      };
    });
}

/** Barra inferior con la selección para imprimir etiquetas. */
function pintarBarraSeleccion() {
  const b = $('#barraSel');
  if (!b) return;
  if (!E.seleccionando) { b.innerHTML = ''; return; }
  b.innerHTML = '<div class="barra-seleccion"><span class="crece"><b class="num">' + E.seleccion.size + '</b> seleccionados</span>' +
    '<button class="btn btn-chico" style="background:transparent;color:#fff;border-color:#55625B" id="selTodos">Todos los filtrados</button>' +
    '<button class="btn btn-chico btn-primario" id="selImprimir"' + (E.seleccion.size ? '' : ' disabled') + '><i class="ti ti-printer"></i>Etiquetas</button></div>';
  $('#selTodos').onclick = () => { VISTAS.catalogo.lista().forEach(m => E.seleccion.add(m.s)); VISTAS.catalogo.pintar(); };
  $('#selImprimir').onclick = () => ir('etiquetas', { modo: 'seleccion' });
}

// -------------------------------------------------------------------------
//  FICHA DEL MATERIAL
// -------------------------------------------------------------------------

VISTAS.ficha = {
  titulo: 'Ficha del material',
  nav: 'catalogo',
  padre: 'catalogo',
  render(app) {
    const m = materialPorSerie(E.params.serie);
    if (!m) {
      app.innerHTML = vacioHTML('ti-help-hexagon', 'No encontramos ese material',
        'No existe un material con el número de serie <b>' + esc(E.params.serie || '') + '</b> en el catálogo.' +
        (E.cargandoCatalogo ? '<br><span class="texto-3">Actualizando catálogo…</span>' : ''),
        '<button class="btn btn-primario" onclick="ir(\'buscar\')"><i class="ti ti-search"></i>Buscar otro</button>');
      return;
    }
    const puedeUbicar = tieneRol('Almacenista');
    const admin = tieneRol('Administrador');
    const cero = !(m.total > 0);

    // Encabezado
    const foto = m.foto
      ? '<div class="ficha-foto" id="fotoFicha" role="button" aria-label="Ampliar foto"><img src="' + esc(urlFoto(m.foto, 400)) + '" alt="Foto de ' + esc(m.d) + '"></div>'
      : '<div class="ficha-foto" style="cursor:default"><i class="ti ' + iconoFamilia(m.f) + '"></i></div>';
    let h = '<div class="pila">' +
      (m.act ? '' : '<div class="caja-error"><i class="ti ti-archive"></i> Este material está dado de baja en el catálogo.</div>') +
      '<div class="ficha-cab">' + foto + '<div class="crece">' +
      '<div class="ficha-nombre">' + esc(m.d) + '</div>' +
      '<button class="serie-copiable serie" onclick="copiar(\'' + esc(m.s) + '\')" aria-label="Copiar número de serie">Serie ' + esc(m.s) + ' <i class="ti ti-copy"></i></button>' +
      '<div style="margin-top:6px"><span class="insignia verde"><i class="ti ' + iconoFamilia(m.f) + '"></i>' + esc(m.f) + '</span></div>' +
      '</div></div>';

    // Existencia total destacada
    h += '<div class="tarjeta"><div class="texto-2 pequeno">Existencia total</div>' +
      '<div class="total-grande num' + (cero ? ' cero' : '') + '">' + fmt(m.total) + ' <small>' + esc(m.u) + '</small></div>' +
      '<div class="texto-3">' + fechaCorteHTML() + '</div></div>';

    // ¿Dónde está? (ubicación física)
    h += '<div class="sec-titulo">¿Dónde está?</div><div class="tarjeta pila">';
    if (m.ub.length) {
      h += m.ub.map(u =>
        '<div class="ubic-principal"><i class="ti ti-map-pin"></i><div class="crece">' +
        '<div class="ubic-texto">' + esc(u.t) + '</div>' + (u.ref && u.ref !== u.t ? '<div class="texto-2">' + esc(u.ref) + '</div>' : '') + '</div>' +
        (puedeUbicar ? '<button class="btn btn-texto btn-icono" data-editar-ub="' + esc(u.id) + '" aria-label="Editar ubicación"><i class="ti ti-edit"></i></button>' : '') +
        '</div>').join('');
      if (puedeUbicar) h += '<div class="fila"><button class="btn btn-chico" id="btnAgregarUb"><i class="ti ti-plus"></i>Otra ubicación</button><button class="btn btn-chico btn-texto" id="btnHistUb"><i class="ti ti-history"></i>Historial</button></div>';
    } else {
      h += '<div class="fila"><i class="ti ti-map-pin-off texto-3" style="font-size:24px"></i><div class="crece texto-2">Sin ubicación asignada</div>' +
        (puedeUbicar ? '<button class="btn btn-primario btn-chico" id="btnAgregarUb"><i class="ti ti-plus"></i>Asignar</button>' : '') + '</div>';
    }
    h += '</div>';

    // Existencias por almacén SAP
    h += '<div class="sec-titulo">Existencias por almacén</div><div class="tarjeta" style="padding:6px 14px">';
    if (m.ex.length) {
      h += '<table class="tabla"><thead><tr><th>Almacén</th><th class="der">Cantidad</th></tr></thead><tbody>' +
        m.ex.map(x => {
          const a = E.alm.get(x[0]);
          const regs = String(x[2] || '').split('|').filter(Boolean);
          return '<tr><td><b>' + esc(x[0]) + '</b>' + (a && a.n ? '<div class="texto-3">' + esc(a.n) + '</div>' : '') +
            (regs.length > 1
              ? '<details class="desplegable"><summary>' + regs.length + ' registros en SAP <i class="ti ti-chevron-down"></i></summary>' +
              '<div class="texto-2 pequeno num">' + regs.map(r => fmt(r)).join(' + ') + ' = ' + fmt(x[1]) + ' ' + esc(m.u) + '</div></details>'
              : '') +
            '</td><td class="der num' + (x[1] > 0 ? '' : ' cero') + '"><b>' + fmt(x[1]) + '</b> ' + esc(m.u) + '</td></tr>';
        }).join('') +
        '</tbody><tfoot><tr><td>Total</td><td class="der num' + (cero ? ' cero' : '') + '">' + fmt(m.total) + ' ' + esc(m.u) + '</td></tr></tfoot></table>';
    } else {
      h += '<p class="texto-2 cero">Sin existencia en SAP al corte.</p>';
    }
    h += '</div>';

    // Datos
    h += '<div class="sec-titulo">Datos</div><div class="tarjeta"><dl class="datos" style="margin:0">' +
      '<dt>Unidad de medida</dt><dd>' + esc(m.u || '—') + '</dd>' +
      '<dt>Peso</dt><dd>' + (m.p !== '' ? esc(fmt(m.p)) + ' kg' : '<span class="texto-3">Sin capturar</span>') + '</dd>' +
      '<dt>Descripción detallada</dt><dd>' + (m.dd ? esc(m.dd) : '<span class="texto-3">Sin capturar</span>') + '</dd>' +
      '<dt>Notas</dt><dd>' + (m.n ? esc(m.n) : '<span class="texto-3">Sin notas</span>') + '</dd>' +
      '</dl>' +
      (admin ? '<button class="btn btn-chico" id="btnEditarMat"><i class="ti ti-edit"></i>Editar material</button>'
        : puedeUbicar ? '<button class="btn btn-chico" id="btnEditarDet"><i class="ti ti-edit"></i>Editar descripción y notas</button>' : '') +
      '</div>';

    // QR y etiqueta
    const url = urlMaterial(m.s);
    h += '<div class="sec-titulo">Código QR</div><div class="tarjeta"><div class="qr-caja">' + qrSVG(url) +
      '<div class="crece pila" style="gap:8px"><div class="texto-3" style="overflow-wrap:anywhere">' + (url ? esc(url) : 'Falta configurar la URL de la app') + '</div>' +
      (admin ? '<button class="btn btn-chico" id="btnEtiqueta"><i class="ti ti-printer"></i>Imprimir etiqueta</button>' : '') +
      (m.et ? '<span class="texto-3"><i class="ti ti-check"></i> Etiqueta impresa ' + esc(fmtFecha(m.et)) + '</span>' : '') +
      '</div></div></div>';

    if (admin) {
      h += '<button class="btn ' + (m.act ? 'btn-peligro' : '') + '" id="btnBaja"><i class="ti ' + (m.act ? 'ti-archive' : 'ti-archive-off') + '"></i>' + (m.act ? 'Dar de baja del catálogo' : 'Reactivar material') + '</button>';
    }
    h += '</div>';
    app.innerHTML = h;

    // Eventos
    const fotoEl = $('#fotoFicha');
    if (fotoEl) fotoEl.onclick = () => abrirModal('Foto', '<img src="' + esc(urlFoto(m.foto, 1600)) + '" alt="">', null, 'modal-foto');
    const agregar = $('#btnAgregarUb');
    if (agregar) agregar.onclick = () => abrirEditorUbicacion(m, null);
    $$('[data-editar-ub]', app).forEach(b => b.onclick = () => abrirEditorUbicacion(m, m.ub.find(u => u.id === b.dataset.editarUb)));
    const hist = $('#btnHistUb');
    if (hist) hist.onclick = () => verHistorialUbicacion(m);
    const ed = $('#btnEditarDet');
    if (ed) ed.onclick = () => abrirEditorDetalle(m);
    const em = $('#btnEditarMat');
    if (em) em.onclick = () => abrirEditorMaterial(m);
    const et = $('#btnEtiqueta');
    if (et) et.onclick = () => ir('etiquetas', { modo: 'series', series: [m.s] });
    const bj = $('#btnBaja');
    if (bj) bj.onclick = async () => {
      const ok = await confirmar(m.act ? 'Dar de baja' : 'Reactivar',
        m.act ? 'El material dejará de aparecer en el catálogo. No se borra nada y puedes reactivarlo cuando quieras.' : '¿Volver a mostrar este material en el catálogo?',
        m.act ? 'Dar de baja' : 'Reactivar', m.act);
      if (!ok) return;
      await conCarga(bj, () => api('cambiarActivoMaterial', token(), m.s, !m.act));
      await despuesDeGuardar(m.act ? 'Material dado de baja' : 'Material reactivado');
    };
  }
};

/** Texto corto de ubicación (igual que en el servidor): "B1 · P3 · R02 · N4". */
function textoUbicacion(u) {
  const t = v => String(v || '').trim().toUpperCase();
  const p = [];
  if (t(u.zona)) p.push(t(u.zona));
  [['pasillo', 'P'], ['rack', 'R'], ['nivel', 'N'], ['posicion', 'Pos ']].forEach(([k, l]) => {
    const v = t(u[k]);
    if (v) p.push(/^\d+$/.test(v) ? l + v : v);
  });
  return p.join(' · ') || String(u.referencia || '').trim().toUpperCase();
}

/** Valores ya capturados, para sugerirlos al escribir (mantiene la captura ordenada). */
function sugerencias(campo) {
  const s = new Set();
  E.cat.mats.forEach(m => m.ub.forEach(u => { if (u[campo]) s.add(u[campo]); }));
  return Array.from(s).sort().map(v => '<option value="' + esc(v) + '">').join('');
}

function abrirEditorUbicacion(m, u) {
  u = u || {};
  const campo = (nombre, etiqueta, valor, lista, ph) =>
    '<div class="campo"><label for="ub_' + nombre + '">' + etiqueta + '</label>' +
    '<input class="entrada" id="ub_' + nombre + '" name="' + nombre + '" value="' + esc(valor || '') + '" list="dl_' + nombre + '" autocomplete="off" placeholder="' + (ph || '') + '">' +
    '<datalist id="dl_' + nombre + '">' + (lista ? sugerencias(lista) : '') + '</datalist></div>';
  abrirModal(u.id ? 'Editar ubicación' : 'Asignar ubicación',
    '<p class="texto-2" style="margin-top:0">' + esc(m.d) + ' · Serie ' + esc(m.s) + '</p>' +
    '<form id="fUb" class="pila">' +
    '<div class="rejilla-2">' +
    campo('zona', 'Zona / bodega', u.z, 'z', 'B1') + campo('pasillo', 'Pasillo', u.pa, 'pa', '3') +
    campo('rack', 'Rack', u.r, 'r', '02') + campo('nivel', 'Nivel', u.ni, 'ni', '4') +
    campo('posicion', 'Posición', u.po, 'po', '') +
    '</div>' +
    '<div class="campo"><label for="ub_referencia">Referencia</label><input class="entrada" id="ub_referencia" name="referencia" value="' + esc(u.ref || '') + '" placeholder="Junto a la puerta 2"></div>' +
    '<div class="caja-ok">Se verá así: <b id="ubPrevia">—</b></div>' +
    '<div class="texto-3">Todos los campos son opcionales; llena los que apliquen. Si escribes solo el número (3) se agrega la letra (P3).</div>' +
    '</form>' +
    '<div class="modal-pie">' +
    (u.id ? '<button class="btn btn-peligro" id="ubQuitar"><i class="ti ti-trash"></i>Quitar</button>' : '<button class="btn" onclick="cerrarModal()">Cancelar</button>') +
    '<button class="btn btn-primario" id="ubGuardar">Guardar</button></div>',
    mod => {
      const form = $('#fUb', mod);
      const previa = () => { $('#ubPrevia', mod).textContent = textoUbicacion(leerFormulario(form)) || '—'; };
      form.addEventListener('input', previa);
      previa();
      $('#ubGuardar', mod).onclick = async ev => {
        const d = leerFormulario(form);
        if (!textoUbicacion(d)) { aviso('Escribe al menos un dato de la ubicación.', 'error'); return; }
        d.id = u.id || '';
        d.serie = m.s;
        await conCarga(ev.currentTarget, () => api('guardarUbicacion', token(), d));
        cerrarModal();
        await despuesDeGuardar('Ubicación guardada');
      };
      const q = $('#ubQuitar', mod);
      if (q) q.onclick = async ev => {
        const btn = ev.currentTarget;
        await conCarga(btn, () => api('quitarUbicacion', token(), u.id));
        cerrarModal();
        await despuesDeGuardar('Ubicación quitada');
      };
    });
}

async function verHistorialUbicacion(m) {
  abrirModal('Historial de ubicación', '<div id="histUb"><div class="cargando-pantalla" style="padding:20px"><div class="spinner"></div></div></div>');
  try {
    const lista = await api('historialUbicaciones', m.s);
    const c = $('#histUb');
    if (!c) return;
    c.innerHTML = lista.length
      ? '<ul class="lista-simple">' + lista.map(x => '<li><div>' + esc(x.detalle) + '</div><div class="texto-3">' + esc(x.fecha) + ' · ' + esc(x.usuario) + '</div></li>').join('') + '</ul>'
      : '<p class="texto-2">Sin cambios registrados.</p>';
  } catch (e) { const c = $('#histUb'); if (c) c.innerHTML = '<div class="caja-error">' + esc(e.message) + '</div>'; }
}

/** Almacenista: editar descripción detallada y notas. */
function abrirEditorDetalle(m) {
  abrirModal('Descripción y notas',
    '<form id="fDet" class="pila">' +
    '<div class="campo"><label>Descripción detallada</label><textarea class="entrada" name="descripcion_detallada" maxlength="2000">' + esc(m.dd) + '</textarea></div>' +
    '<div class="campo"><label>Notas</label><textarea class="entrada" name="notas" maxlength="2000">' + esc(m.n) + '</textarea></div>' +
    '</form><div class="modal-pie"><button class="btn" onclick="cerrarModal()">Cancelar</button><button class="btn btn-primario" id="detGuardar">Guardar</button></div>',
    mod => {
      $('#detGuardar', mod).onclick = async ev => {
        await conCarga(ev.currentTarget, () => api('guardarDetalleMaterial', token(), m.s, leerFormulario($('#fDet', mod))));
        cerrarModal();
        await despuesDeGuardar('Cambios guardados');
      };
    });
}

/** Administrador: alta o edición completa de un material. */
function abrirEditorMaterial(m) {
  const nuevo = !m;
  m = m || { s: '', d: '', f: '', u: 'PZA', p: '', dd: '', n: '' };
  const familias = E.familias.map(x => '<option value="' + esc(x.f) + '">').join('');
  const unidades = Array.from(new Set(E.cat.mats.map(x => x.u).filter(Boolean))).sort().map(u => '<option value="' + esc(u) + '">').join('');
  abrirModal(nuevo ? 'Nuevo material' : 'Editar material',
    '<form id="fMat" class="pila">' +
    '<div class="campo"><label>Número de serie</label><input class="entrada serie" name="numero_serie" inputmode="numeric" value="' + esc(m.s) + '"' + (nuevo ? '' : ' readonly') + ' required></div>' +
    '<div class="campo"><label>Texto breve (nombre)</label><input class="entrada" name="descripcion" value="' + esc(m.d) + '" required></div>' +
    '<div class="rejilla-2"><div class="campo"><label>Familia</label><input class="entrada" name="familia" list="dlFam" value="' + esc(m.f) + '"><datalist id="dlFam">' + familias + '</datalist></div>' +
    '<div class="campo"><label>Unidad</label><input class="entrada" name="unidad" list="dlUni" value="' + esc(m.u) + '"><datalist id="dlUni">' + unidades + '</datalist></div></div>' +
    '<div class="campo"><label>Peso (kg)</label><input class="entrada" name="peso" inputmode="decimal" value="' + esc(m.p) + '"></div>' +
    '<div class="campo"><label>Descripción detallada</label><textarea class="entrada" name="descripcion_detallada">' + esc(m.dd) + '</textarea></div>' +
    '<div class="campo"><label>Notas</label><textarea class="entrada" name="notas">' + esc(m.n) + '</textarea></div>' +
    '<div class="texto-3">El nombre, la familia y la unidad vienen de SAP: si cambian allá, la siguiente importación los actualiza. Las existencias solo se cargan desde SAP.</div>' +
    '</form><div class="modal-pie"><button class="btn" onclick="cerrarModal()">Cancelar</button><button class="btn btn-primario" id="matGuardar">Guardar</button></div>',
    mod => {
      $('#matGuardar', mod).onclick = async ev => {
        const d = leerFormulario($('#fMat', mod));
        await conCarga(ev.currentTarget, () => api('guardarMaterial', token(), d, nuevo));
        cerrarModal();
        await despuesDeGuardar(nuevo ? 'Material dado de alta' : 'Material actualizado');
        if (nuevo) ir('ficha', { serie: normSerie(d.numero_serie) });
      };
    });
}

// -------------------------------------------------------------------------
//  MÁS (menú)
// -------------------------------------------------------------------------

VISTAS.mas = {
  titulo: 'Más opciones',
  raiz: true,
  render(app) {
    const item = (icono, texto, accion, sub) =>
      '<button onclick="' + accion + '"><i class="ti ' + icono + '"></i><span class="crece">' + texto +
      (sub ? '<div class="texto-3">' + sub + '</div>' : '') + '</span><i class="ti ti-chevron-right"></i></button>';
    let h = '<div class="pila">';
    if (E.ses) {
      h += '<div class="tarjeta fila"><span class="barra-btn" style="background:var(--cfe-verde)"><span class="inicial">' + esc(E.ses.nombre.charAt(0).toUpperCase()) + '</span></span>' +
        '<div class="crece"><b>' + esc(E.ses.nombre) + '</b><div class="texto-3">' + esc(E.ses.usuario) + ' · ' + esc(E.ses.rol) + '</div></div></div>';
    } else {
      h += '<div class="tarjeta pila"><div><b>Estás en modo consulta</b><div class="texto-2 pequeno">Almacenistas y administradores: inicien sesión para capturar ubicaciones, importar y administrar.</div></div>' +
        '<button class="btn btn-primario" onclick="ir(\'login\', {volver: \'mas\'})"><i class="ti ti-login-2"></i>Iniciar sesión</button></div>';
    }
    if (tieneRol('Almacenista')) {
      h += '<div class="sec-titulo">Almacén</div><div class="menu">' +
        item('ti-map-pin-off', 'Materiales sin ubicación', "ir('sinUbicacion')", 'Para avanzar en la captura de ubicaciones') +
        '</div>';
    }
    if (tieneRol('Administrador')) {
      h += '<div class="sec-titulo">Administración</div><div class="menu">' +
        item('ti-file-upload', 'Importar Excel de SAP', "ir('importar')", 'Existencias del día') +
        item('ti-map-2', 'Carga masiva de ubicaciones', "ir('ubicMasiva')", 'Plantilla de Excel') +
        item('ti-printer', 'Imprimir etiquetas QR', "ir('etiquetas', {modo: 'familia'})") +
        item('ti-tags', 'Etiquetas pendientes', "ir('etiquetas', {modo: 'pendientes'})", 'Materiales sin etiqueta impresa') +
        item('ti-packages', 'Materiales', "ir('admMateriales')", 'Alta, edición y baja') +
        item('ti-building-warehouse', 'Almacenes', "ir('admAlmacenes')", 'Nombres y usos') +
        item('ti-users', 'Usuarios', "ir('admUsuarios')") +
        item('ti-history', 'Bitácora', "ir('bitacora')") +
        item('ti-settings', 'Configuración', "ir('config')", 'URL de los códigos QR') +
        '</div>';
    }
    h += '<div class="sec-titulo">General</div><div class="menu">' +
      item('ti-refresh', 'Actualizar catálogo', 'actualizarManual()', E.cat ? 'Versión de datos ' + esc(E.cat.v) + ' · ' + esc(E.cat.generado || '') : '') +
      (E.ses ? item('ti-key', 'Cambiar mi PIN', 'abrirCambioPin(false)') + item('ti-logout', 'Cerrar sesión', 'cerrarSesion()') : '') +
      '</div>' +
      '<p class="texto-3" style="text-align:center">Catálogo QR · Almacén CFE Zona Campeche<br>' + fechaCorteHTML() + '</p></div>';
    app.innerHTML = h;
  }
};

async function actualizarManual() {
  aviso('Actualizando catálogo…');
  await cargarCatalogo(true);
  aviso('Catálogo actualizado');
}

// -------------------------------------------------------------------------
//  INICIAR SESIÓN
// -------------------------------------------------------------------------

VISTAS.login = {
  titulo: 'Iniciar sesión',
  necesitaCatalogo: false,
  render(app) {
    app.innerHTML = '<div class="tarjeta pila" style="max-width:420px;margin:0 auto">' +
      '<div style="text-align:center"><i class="ti ti-user-shield" style="font-size:44px;color:var(--cfe-verde)"></i><h2>Almacenista o administrador</h2>' +
      '<p class="texto-2 pequeno">La consulta del catálogo no necesita sesión.</p></div>' +
      '<form id="fLogin" class="pila">' +
      '<div class="campo"><label for="lgUsuario">Usuario</label><input class="entrada" id="lgUsuario" name="usuario" autocomplete="username" autocapitalize="none" spellcheck="false" required></div>' +
      '<div class="campo"><label for="lgPin">PIN</label><input class="entrada" id="lgPin" name="pin" type="password" inputmode="numeric" autocomplete="current-password" maxlength="8" required></div>' +
      '<div id="lgError"></div>' +
      '<button class="btn btn-primario btn-bloque" id="lgEntrar" data-cargando="Entrando…" type="submit"><i class="ti ti-login-2"></i>Entrar</button>' +
      '</form></div>';
    $('#fLogin').onsubmit = async ev => {
      ev.preventDefault();
      const d = leerFormulario(ev.target);
      $('#lgError').innerHTML = '';
      try {
        const r = await conCarga($('#lgEntrar'), () => api('iniciarSesion', d.usuario, d.pin));
        E.ses = { token: r.token, usuario: r.usuario, nombre: r.nombre, rol: r.rol };
        Local.set('ses', JSON.stringify(E.ses));
        pintarUsuario();
        aviso('Bienvenido, ' + r.nombre);
        const volver = E.params.volver && E.params.volver !== 'login' ? E.params.volver : 'mas';
        ir(volver, {}, true);
        if (r.debeCambiarPin) abrirCambioPin(true);
      } catch (e) {
        $('#lgError').innerHTML = '<div class="caja-error">' + esc(e.message) + '</div>';
      }
    };
    $('#lgUsuario').focus();
  }
};

function abrirCambioPin(obligatorio) {
  abrirModal('Cambiar mi PIN',
    (obligatorio ? '<div class="caja-aviso" style="margin-bottom:12px">Estás usando el PIN inicial. Por seguridad, cámbialo ahora.</div>' : '') +
    '<form id="fPin" class="pila">' +
    '<div class="campo"><label>PIN actual</label><input class="entrada" name="actual" type="password" inputmode="numeric" maxlength="8"></div>' +
    '<div class="campo"><label>PIN nuevo (4 a 8 números)</label><input class="entrada" name="nuevo" type="password" inputmode="numeric" maxlength="8"></div>' +
    '<div class="campo"><label>Repite el PIN nuevo</label><input class="entrada" name="repite" type="password" inputmode="numeric" maxlength="8"></div>' +
    '</form><div class="modal-pie"><button class="btn" onclick="cerrarModal()">' + (obligatorio ? 'Después' : 'Cancelar') + '</button><button class="btn btn-primario" id="pinGuardar">Cambiar PIN</button></div>',
    mod => {
      $('#pinGuardar', mod).onclick = async ev => {
        const d = leerFormulario($('#fPin', mod));
        if (d.nuevo !== d.repite) { aviso('Los PIN nuevos no coinciden.', 'error'); return; }
        await conCarga(ev.currentTarget, () => api('cambiarMiPin', token(), d.actual, d.nuevo));
        cerrarModal();
        aviso('PIN actualizado');
      };
    });
}

// -------------------------------------------------------------------------
//  REPORTE: MATERIALES SIN UBICACIÓN
// -------------------------------------------------------------------------

VISTAS.sinUbicacion = {
  titulo: 'Materiales sin ubicación',
  rol: 'Almacenista',
  padre: 'mas',
  nav: 'mas',
  filtros: { alm: '', fam: '', soloExist: true },
  render(app) {
    const f = this.filtros;
    const almacenes = Array.from(E.alm.values()).sort((a, b) => a.c.localeCompare(b.c));
    app.innerHTML = '<div class="pila">' +
      '<div class="rejilla-2">' +
      '<div class="campo"><label>Almacén SAP</label><select class="entrada" id="suAlm"><option value="">Todos</option>' +
      almacenes.map(a => '<option value="' + esc(a.c) + '"' + (f.alm === a.c ? ' selected' : '') + '>' + esc(a.c + (a.n ? ' · ' + a.n : '')) + '</option>').join('') + '</select></div>' +
      '<div class="campo"><label>Familia</label><select class="entrada" id="suFam"><option value="">Todas</option>' +
      E.familias.map(x => '<option' + (f.fam === x.f ? ' selected' : '') + '>' + esc(x.f) + '</option>').join('') + '</select></div>' +
      '</div>' +
      '<label class="check"><input type="checkbox" id="suExist"' + (f.soloExist ? ' checked' : '') + '>Solo materiales con existencia</label>' +
      '<div id="suAvance"></div>' +
      '<div class="fila-sep"><span class="texto-3" id="suCuenta"></span><button class="btn btn-chico" id="suExportar"><i class="ti ti-file-spreadsheet"></i>Exportar</button></div>' +
      '<div id="suRes"></div></div>';
    const actualizar = () => {
      f.alm = $('#suAlm').value; f.fam = $('#suFam').value; f.soloExist = $('#suExist').checked;
      this.pintar();
    };
    ['#suAlm', '#suFam', '#suExist'].forEach(s => $(s).addEventListener('change', actualizar));
    $('#suExportar').onclick = () => exportarMateriales(this.lista(), 'Sin_ubicacion');
    activarTarjetas($('#suRes'));
    this.pintar();
  },
  base() {
    const f = this.filtros;
    return materialesVisibles().filter(m => m.act &&
      (!f.alm || m.ex.some(x => x[0] === f.alm)) && (!f.fam || m.f === f.fam) && (!f.soloExist || m.total > 0));
  },
  lista() { return this.base().filter(m => !m.ub.length); },
  pintar() {
    const base = this.base(), lista = base.filter(m => !m.ub.length);
    const hechos = base.length - lista.length;
    const pct = base.length ? Math.round(hechos * 100 / base.length) : 0;
    $('#suAvance').innerHTML = '<div class="tarjeta"><div class="fila-sep"><b>Avance de captura</b><span class="num">' + fmt(hechos) + ' de ' + fmt(base.length) + ' (' + pct + '%)</span></div>' +
      '<div style="height:10px;background:var(--cfe-verde-claro);border-radius:6px;margin-top:8px;overflow:hidden"><div style="height:100%;width:' + pct + '%;background:var(--cfe-verde)"></div></div></div>';
    $('#suCuenta').textContent = fmt(lista.length) + ' sin ubicación';
    $('#suRes').innerHTML = lista.length
      ? '<div class="resultados lista">' + lista.slice(0, 300).map(tarjetaMaterial).join('') + '</div>' +
      (lista.length > 300 ? '<p class="texto-3">Se muestran 300. Usa "Exportar" para la lista completa.</p>' : '')
      : vacioHTML('ti-circle-check', '¡Todo ubicado!', 'Todos los materiales de este filtro ya tienen ubicación física.');
  },
  refrescar() { if ($('#suRes')) this.pintar(); else render(); }
};

// -------------------------------------------------------------------------
//  ADMINISTRACIÓN: MATERIALES
// -------------------------------------------------------------------------

VISTAS.admMateriales = {
  titulo: 'Materiales',
  rol: 'Administrador',
  padre: 'mas',
  nav: 'mas',
  q: '', bajas: false,
  render(app) {
    app.innerHTML = '<div class="pila">' +
      '<button class="btn btn-primario" id="amNuevo"><i class="ti ti-plus"></i>Nuevo material</button>' +
      buscadorHTML('amQ', this.q, 'Buscar material') +
      '<label class="check"><input type="checkbox" id="amBajas"' + (this.bajas ? ' checked' : '') + '>Mostrar solo dados de baja</label>' +
      '<div class="texto-3">Abre un material para editarlo o darlo de baja desde su ficha.</div>' +
      '<div id="amRes"></div></div>';
    $('#amNuevo').onclick = () => abrirEditorMaterial(null);
    $('#amBajas').onchange = e => { this.bajas = e.target.checked; this.pintar(); };
    conectarBuscador('amQ', q => { this.q = q; this.pintar(); });
    activarTarjetas($('#amRes'));
    this.pintar();
  },
  pintar() {
    const lista = buscarMateriales(this.q, E.cat.mats).filter(m => this.bajas ? !m.act : true);
    $('#amRes').innerHTML = lista.length
      ? '<div class="texto-3" style="margin-bottom:8px">' + fmt(lista.length) + ' materiales</div><div class="resultados lista">' + lista.slice(0, 200).map(tarjetaMaterial).join('') + '</div>'
      : vacioHTML('ti-packages', 'Sin materiales', this.bajas ? 'No hay materiales dados de baja.' : 'No hay coincidencias.');
  },
  refrescar() { if ($('#amRes')) this.pintar(); else render(); }
};

// -------------------------------------------------------------------------
//  ADMINISTRACIÓN: ALMACENES
// -------------------------------------------------------------------------

VISTAS.admAlmacenes = {
  titulo: 'Almacenes',
  rol: 'Administrador',
  padre: 'mas',
  nav: 'mas',
  render(app) {
    const conteo = {};
    E.cat.mats.forEach(m => m.ex.forEach(x => { conteo[x[0]] = (conteo[x[0]] || 0) + 1; }));
    const lista = Array.from(E.alm.values()).sort((a, b) => a.c.localeCompare(b.c));
    app.innerHTML = '<div class="pila">' +
      '<div class="texto-2 pequeno">Los almacenes de SAP son clasificaciones o destinos del material; físicamente todo está en el mismo lugar. Los códigos nuevos se crean solos al importar el Excel; aquí puedes ponerles nombre y uso.</div>' +
      '<button class="btn" id="aaNuevo"><i class="ti ti-plus"></i>Nuevo almacén</button>' +
      '<div class="menu">' + lista.map(a =>
        '<button data-alm="' + esc(a.c) + '"><b style="min-width:52px">' + esc(a.c) + '</b><span class="crece">' + (a.n ? esc(a.n) : '<span class="insignia alerta">Sin nombre</span>') +
        '<div class="texto-3">' + (conteo[a.c] || 0) + ' materiales' + (a.u ? ' · ' + esc(a.u) : '') + '</div></span><i class="ti ti-edit texto-3"></i></button>').join('') +
      '</div></div>';
    $$('[data-alm]', app).forEach(b => b.onclick = () => this.editar(E.alm.get(b.dataset.alm)));
    $('#aaNuevo').onclick = () => this.editar(null);
  },
  editar(a) {
    const nuevo = !a;
    a = a || { c: '', n: '', u: '' };
    abrirModal(nuevo ? 'Nuevo almacén' : 'Almacén ' + a.c,
      '<form id="fAlm" class="pila">' +
      '<div class="campo"><label>Código</label><input class="entrada" name="codigo" value="' + esc(a.c) + '"' + (nuevo ? '' : ' readonly') + '></div>' +
      '<div class="campo"><label>Nombre</label><input class="entrada" name="nombre" value="' + esc(a.n) + '"></div>' +
      '<div class="campo"><label>Uso / descripción</label><textarea class="entrada" name="uso">' + esc(a.u) + '</textarea></div>' +
      '</form><div class="modal-pie"><button class="btn" onclick="cerrarModal()">Cancelar</button><button class="btn btn-primario" id="almGuardar">Guardar</button></div>',
      mod => {
        $('#almGuardar', mod).onclick = async ev => {
          await conCarga(ev.currentTarget, () => api('guardarAlmacen', token(), leerFormulario($('#fAlm', mod)), nuevo));
          cerrarModal();
          await despuesDeGuardar('Almacén guardado');
        };
      });
  }
};

// -------------------------------------------------------------------------
//  ADMINISTRACIÓN: USUARIOS
// -------------------------------------------------------------------------

VISTAS.admUsuarios = {
  titulo: 'Usuarios',
  rol: 'Administrador',
  padre: 'mas',
  nav: 'mas',
  necesitaCatalogo: false,
  async render(app) {
    app.innerHTML = '<div class="pila"><button class="btn btn-primario" id="auNuevo"><i class="ti ti-user-plus"></i>Nuevo usuario</button>' +
      '<div class="texto-2 pequeno"><b>Almacenista:</b> captura ubicaciones, descripción detallada y notas. <b>Administrador:</b> todo, incluida la importación de SAP. La consulta no necesita usuario.</div>' +
      '<div id="auLista"><div class="esqueleto"></div></div></div>';
    $('#auNuevo').onclick = () => this.editar(null);
    try {
      this.usuarios = await api('listarUsuarios', token());
      const c = $('#auLista');
      if (!c) return;
      c.innerHTML = '<div class="menu">' + this.usuarios.map(u =>
        '<button data-us="' + esc(u.usuario) + '"><i class="ti ' + (u.rol === 'Administrador' ? 'ti-user-shield' : 'ti-user') + '"></i><span class="crece"><b>' + esc(u.nombre) + '</b>' +
        (u.activo ? '' : ' <span class="insignia peligro">Inactivo</span>') +
        '<div class="texto-3">' + esc(u.usuario) + ' · ' + esc(u.rol) + (u.ultimo ? ' · último acceso ' + esc(u.ultimo) : '') + '</div></span><i class="ti ti-edit texto-3"></i></button>').join('') + '</div>';
      $$('[data-us]', c).forEach(b => b.onclick = () => this.editar(this.usuarios.find(u => u.usuario === b.dataset.us)));
    } catch (e) {
      const c = $('#auLista'); if (c) c.innerHTML = '<div class="caja-error">' + esc(e.message) + '</div>';
    }
  },
  editar(u) {
    const nuevo = !u;
    u = u || { usuario: '', nombre: '', rol: 'Almacenista', activo: true };
    abrirModal(nuevo ? 'Nuevo usuario' : 'Editar usuario',
      '<form id="fUs" class="pila">' +
      '<div class="campo"><label>Usuario (para entrar)</label><input class="entrada" name="usuario" autocapitalize="none" value="' + esc(u.usuario) + '"' + (nuevo ? '' : ' readonly') + ' placeholder="jperez"></div>' +
      '<div class="campo"><label>Nombre</label><input class="entrada" name="nombre" value="' + esc(u.nombre) + '" placeholder="Juan Pérez"></div>' +
      '<div class="campo"><label>Rol</label><select class="entrada" name="rol">' + ['Almacenista', 'Administrador'].map(r => '<option' + (u.rol === r ? ' selected' : '') + '>' + r + '</option>').join('') + '</select></div>' +
      (nuevo
        ? '<div class="campo"><label>PIN inicial (4 a 8 números)</label><input class="entrada" name="pin" inputmode="numeric" maxlength="8"></div>'
        : '<label class="check"><input type="checkbox" name="activo"' + (u.activo ? ' checked' : '') + '>Usuario activo</label>' +
        '<button type="button" class="btn btn-chico" id="usPin"><i class="ti ti-key"></i>Asignar PIN nuevo</button>') +
      '</form><div class="modal-pie"><button class="btn" onclick="cerrarModal()">Cancelar</button><button class="btn btn-primario" id="usGuardar">Guardar</button></div>',
      mod => {
        $('#usGuardar', mod).onclick = async ev => {
          const d = leerFormulario($('#fUs', mod));
          if (nuevo) d.activo = true;
          await conCarga(ev.currentTarget, () => api('guardarUsuario', token(), d, nuevo));
          cerrarModal();
          aviso('Usuario guardado');
          render();
        };
        const bp = $('#usPin', mod);
        if (bp) bp.onclick = () => {
          abrirModal('PIN nuevo para ' + u.usuario,
            '<div class="campo"><label>PIN nuevo (4 a 8 números)</label><input class="entrada" id="pinAdm" inputmode="numeric" maxlength="8"></div>' +
            '<div class="modal-pie"><button class="btn" onclick="cerrarModal()">Cancelar</button><button class="btn btn-primario" id="pinAdmOk">Asignar</button></div>',
            m2 => {
              $('#pinAdmOk', m2).onclick = async ev => {
                await conCarga(ev.currentTarget, () => api('restablecerPin', token(), u.usuario, $('#pinAdm', m2).value));
                cerrarModal();
                aviso('PIN asignado. Compártelo con la persona.');
              };
            });
        };
      });
  }
};

// -------------------------------------------------------------------------
//  ADMINISTRACIÓN: BITÁCORA
// -------------------------------------------------------------------------

VISTAS.bitacora = {
  titulo: 'Bitácora',
  rol: 'Administrador',
  padre: 'mas',
  nav: 'mas',
  necesitaCatalogo: false,
  async render(app) {
    app.innerHTML = '<div class="pila">' + buscadorHTML('bqQ', '', 'Filtrar por usuario, acción o serie') +
      '<div id="bqLista"><div class="esqueleto"></div></div></div>';
    try {
      const datos = await api('verBitacora', token(), 500);
      const pintar = q => {
        const qn = norm(q);
        const lista = datos.filter(x => !qn || norm(x.usuario + ' ' + x.accion + ' ' + x.detalle).includes(qn));
        $('#bqLista').innerHTML = lista.length
          ? '<div class="tarjeta"><ul class="lista-simple">' + lista.map(x =>
            '<li><div class="fila-sep"><span class="insignia verde">' + esc(x.accion) + '</span><span class="texto-3">' + esc(x.fecha) + '</span></div>' +
            '<div style="margin-top:4px;overflow-wrap:anywhere">' + esc(x.detalle) + '</div><div class="texto-3">' + esc(x.usuario) + '</div></li>').join('') + '</ul></div>'
          : vacioHTML('ti-history', 'Sin registros', '');
      };
      conectarBuscador('bqQ', pintar);
      pintar('');
    } catch (e) {
      const c = $('#bqLista'); if (c) c.innerHTML = '<div class="caja-error">' + esc(e.message) + '</div>';
    }
  }
};

// -------------------------------------------------------------------------
//  ADMINISTRACIÓN: CONFIGURACIÓN (URL fija de los QR)
// -------------------------------------------------------------------------

VISTAS.config = {
  titulo: 'Configuración',
  rol: 'Administrador',
  padre: 'mas',
  nav: 'mas',
  necesitaCatalogo: false,
  async render(app) {
    app.innerHTML = '<div class="esqueleto"></div>';
    try {
      const info = await api('infoDespliegue', token());
      const sugerida = urlPaginaActual();
      app.innerHTML = '<div class="pila">' +
        '<div class="tarjeta pila"><h2>URL de los códigos QR</h2>' +
        '<p class="texto-2 pequeno" style="margin:0">Es la dirección de esta página (GitHub Pages). Va dentro de TODAS las etiquetas, así que no la cambies después de imprimirlas.</p>' +
        (info.urlFija ? '' : '<div class="caja-aviso"><i class="ti ti-alert-triangle"></i> Todavía no está guardada. Revisa la sugerencia y pulsa Guardar antes de imprimir etiquetas.</div>') +
        '<div class="campo"><label>URL fija de la página</label><input class="entrada" id="cfgUrl" value="' + esc(info.urlFija || sugerida) + '" placeholder="https://usuario.github.io/catalogo/"></div>' +
        (sugerida ? '<div class="texto-3">Dirección con la que abriste la app ahora: ' + esc(sugerida) + '</div>' : '') +
        '<button class="btn btn-primario" id="cfgGuardar">Guardar URL</button></div>' +
        '<div class="tarjeta pila"><h2>Servidor (Apps Script)</h2>' +
        '<div class="texto-3" style="overflow-wrap:anywhere">' + esc(URL_SERVIDOR) + '</div>' +
        '<p class="texto-2 pequeno" style="margin:0">Se cambia en el archivo config.js de GitHub.</p></div>' +
        '<div class="tarjeta"><h2>Hoja de cálculo</h2><p class="texto-2 pequeno">Base de datos de la app. Evita editarla a mano; si lo haces, no cambies los encabezados ni el orden de las columnas.</p>' +
        '<a class="btn btn-chico" href="' + esc(info.hoja) + '" target="_blank" rel="noopener"><i class="ti ti-external-link"></i>Abrir hoja</a></div>' +
        '</div>';
      $('#cfgGuardar').onclick = async ev => {
        await conCarga(ev.currentTarget, () => api('guardarUrlFija', token(), $('#cfgUrl').value));
        await despuesDeGuardar('URL guardada');
        render();
      };
    } catch (e) {
      app.innerHTML = '<div class="caja-error">' + esc(e.message) + '</div>';
    }
  }
};


/* ######################################################################
   3. EXCEL
   ###################################################################### */
/* =========================================================================
   JsImport.html – Lectura de Excel (SheetJS) en el navegador
   - Importación diaria del Excel de SAP (vista previa + confirmación)
   - Carga masiva de ubicaciones con plantilla descargable
   - Exportar listas a Excel
   ========================================================================= */

/** Carga SheetJS solo cuando se necesita (pesa ~1 MB). */
let _sheetjs = null;
function cargarSheetJS() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  if (_sheetjs) return _sheetjs;
  const fuentes = [
    'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
    'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js'
  ];
  _sheetjs = new Promise((resolver, rechazar) => {
    const intentar = i => {
      if (i >= fuentes.length) { _sheetjs = null; return rechazar(new Error('No se pudo cargar el lector de Excel. Revisa tu conexión.')); }
      const s = document.createElement('script');
      s.src = fuentes[i];
      s.onload = () => resolver(window.XLSX);
      s.onerror = () => { s.remove(); intentar(i + 1); };
      document.head.appendChild(s);
    };
    intentar(0);
  });
  return _sheetjs;
}

/** Convierte un encabezado en clave comparable: "NÚMERO DE SERIE" → "numerodeserie". */
function claveEncabezado(t) { return norm(t).replace(/[^a-z0-9]/g, ''); }

/** Columnas del Excel de SAP. Se reconocen por nombre, sin importar acentos, mayúsculas o espacios. */
const COLS_EXISTENCIAS = {
  a: { nombre: 'Almacén', alias: ['almacen', 'alm', 'almacensap', 'codigoalmacen'] },
  s: { nombre: 'NÚMERO DE SERIE', alias: ['numerodeserie', 'nodeserie', 'numdeserie', 'numserie', 'noserie', 'serie', 'material', 'numerodematerial', 'nomaterial'] },
  d: { nombre: 'Texto breve material', alias: ['textobrevematerial', 'textobrevedematerial', 'textobreve', 'descripcion', 'descripciondematerial'] },
  u: { nombre: 'Unidad de medida', alias: ['unidaddemedida', 'unidadmedida', 'unidad', 'um', 'umb', 'unidaddemedidabase'] },
  c: { nombre: 'Inventario', alias: ['inventario', 'existencia', 'existencias', 'cantidad', 'stock', 'libreutilizacion'] },
  f: { nombre: 'Familia', alias: ['familia', 'categoria'] }
};
const OBLIGATORIAS_EXISTENCIAS = ['a', 's', 'd', 'u', 'c', 'f'];

/** Columnas de la plantilla de ubicaciones. */
const COLS_UBICACIONES = {
  s: { nombre: 'NÚMERO DE SERIE', alias: COLS_EXISTENCIAS.s.alias },
  zona: { nombre: 'Zona', alias: ['zona', 'bodega', 'zonabodega'] },
  pasillo: { nombre: 'Pasillo', alias: ['pasillo'] },
  rack: { nombre: 'Rack', alias: ['rack', 'anaquel', 'estante'] },
  nivel: { nombre: 'Nivel', alias: ['nivel'] },
  posicion: { nombre: 'Posición', alias: ['posicion'] },
  referencia: { nombre: 'Referencia', alias: ['referencia', 'referencias'] },
  libre: { nombre: 'Ubicación', alias: ['ubicacion', 'ubicaciontextolibre', 'ubicacionlibre'] }
};

/**
 * Busca la fila de encabezados (puede no estar en la fila 1 ni empezar en la columna A,
 * como pasa en el Excel de SAP, que trae dos columnas vacías al inicio).
 * Devuelve { indice, columnas: { a: 2, s: 3, ... } } o null.
 */
function encontrarEncabezados(filas, definicion, obligatorias) {
  const limite = Math.min(filas.length, 40);
  let mejor = null;
  for (let i = 0; i < limite; i++) {
    const fila = filas[i] || [];
    const columnas = {};
    fila.forEach((celda, j) => {
      const k = claveEncabezado(celda);
      if (!k) return;
      Object.keys(definicion).forEach(campo => {
        if (columnas[campo] === undefined && definicion[campo].alias.includes(k)) columnas[campo] = j;
      });
    });
    const encontradas = obligatorias.filter(c => columnas[c] !== undefined).length;
    if (!mejor || encontradas > mejor.encontradas) mejor = { indice: i, columnas, encontradas };
    if (encontradas === obligatorias.length) break;
  }
  return mejor;
}

/** Lee el archivo y devuelve las filas de la primera hoja que tenga los encabezados. */
async function leerHojasExcel(archivo) {
  const XLSX = await cargarSheetJS();
  const datos = await archivo.arrayBuffer();
  const libro = XLSX.read(datos, { type: 'array', cellDates: false });
  return libro.SheetNames.map(n => ({
    hoja: n,
    filas: XLSX.utils.sheet_to_json(libro.Sheets[n], { header: 1, raw: true, defval: '', blankrows: true })
  }));
}

/** Toma las hojas del libro y arma las filas para el servidor. */
function prepararFilas(hojas, definicion, obligatorias) {
  let mejor = null;
  for (const h of hojas) {
    const enc = encontrarEncabezados(h.filas, definicion, obligatorias);
    if (enc && (!mejor || enc.encontradas > mejor.enc.encontradas)) mejor = { h, enc };
    if (enc && enc.encontradas === obligatorias.length) break;
  }
  const faltan = obligatorias.filter(c => !mejor || mejor.enc.columnas[c] === undefined);
  if (faltan.length) {
    throw new Error('Faltan columnas obligatorias en el Excel: ' + faltan.map(c => '"' + definicion[c].nombre + '"').join(', ') +
      '. Revisa que la fila de encabezados tenga esos nombres.');
  }
  const { h, enc } = mejor;
  const filas = [];
  let ignoradas = 0;
  for (let i = enc.indice + 1; i < h.filas.length; i++) {
    const r = h.filas[i] || [];
    const o = { fila: i + 1 };
    let algo = false;
    Object.keys(enc.columnas).forEach(campo => {
      let v = r[enc.columnas[campo]];
      if (v instanceof Date) v = v.toISOString().slice(0, 10);
      if (v === null || v === undefined) v = '';
      if (typeof v === 'string') v = v.trim();
      o[campo] = v;
      if (v !== '') algo = true;
    });
    if (!algo) continue;                 // fila vacía
    if (o.s === '' && o.a === '') {      // filas de totales o notas al final
      if (definicion === COLS_EXISTENCIAS) { ignoradas++; continue; }
    }
    filas.push(o);
  }
  return { hoja: h.hoja, filaEncabezado: enc.indice + 1, columnas: enc.columnas, filas, ignoradas };
}

// -------------------------------------------------------------------------
//  Piezas de la vista previa
// -------------------------------------------------------------------------

function metrica(valor, etiqueta, peligro) {
  return '<div class="metrica' + (peligro && valor ? ' peligro' : '') + '"><div class="valor num">' + fmt(valor) + '</div><div class="etq">' + etiqueta + '</div></div>';
}

function detalleLista(titulo, items, filaHTML, total) {
  if (!items || !items.length) return '';
  return '<details class="tarjeta"><summary style="cursor:pointer;font-weight:600">' + titulo + ' <span class="texto-3">(' + fmt(total || items.length) + ')</span></summary>' +
    '<ul class="lista-simple desplaza" style="margin-top:8px">' + items.map(x => '<li>' + filaHTML(x) + '</li>').join('') + '</ul>' +
    ((total || 0) > items.length ? '<p class="texto-3">Se muestran los primeros ' + items.length + '.</p>' : '') + '</details>';
}

function pintarImportaciones(contenedor, tipo) {
  api('listarImportaciones', token()).then(lista => {
    lista = lista.filter(x => x.tipo === tipo).slice(0, 10);
    if (!contenedor.isConnected) return;
    contenedor.innerHTML = lista.length
      ? '<div class="sec-titulo">Últimas importaciones</div><div class="tarjeta"><ul class="lista-simple">' + lista.map(x =>
        '<li><div class="fila-sep"><b>' + esc(x.archivo || 'archivo') + '</b><span class="texto-3">' + esc(x.fecha) + '</span></div>' +
        '<div class="texto-3">' + (x.corte ? 'Corte ' + fmtFecha(x.corte) + ' · ' : '') + fmt(x.leidas) + ' filas · ' + fmt(x.nuevos) + ' nuevos · ' + fmt(x.cambios) + ' cambios · ' + fmt(x.errores) + ' errores · ' + esc(x.usuario) + '</div></li>').join('') + '</ul></div>'
      : '';
  }).catch(() => { });
}

// -------------------------------------------------------------------------
//  IMPORTAR EXCEL DE SAP
// -------------------------------------------------------------------------

VISTAS.importar = {
  titulo: 'Importar Excel de SAP',
  rol: 'Administrador',
  padre: 'mas',
  nav: 'mas',
  render(app) {
    this.payload = null;
    app.innerHTML = '<div class="pila">' +
      '<div class="tarjeta pila">' +
      '<div class="texto-2 pequeno">Sube el Excel exportado de SAP (.xls o .xlsx) con las columnas <b>Almacén, NÚMERO DE SERIE, Texto breve material, Unidad de medida, Inventario y Familia</b>. Antes de aplicar verás una vista previa.</div>' +
      '<div class="campo"><label for="impArchivo">Archivo de Excel</label><input class="entrada" type="file" id="impArchivo" accept=".xls,.xlsx,.xlsm,.csv"></div>' +
      '<div class="campo"><label for="impCorte">Fecha de corte (día de los datos de SAP)</label><input class="entrada" type="date" id="impCorte" value="' + hoyISO() + '"></div>' +
      '<button class="btn btn-primario" id="impAnalizar" data-cargando="Analizando…" disabled><i class="ti ti-file-search"></i>Ver vista previa</button>' +
      '</div>' +
      '<div id="impPrevia"></div>' +
      '<div id="impHistorial"></div>' +
      '</div>';
    const archivo = $('#impArchivo'), btn = $('#impAnalizar');
    archivo.onchange = () => { btn.disabled = !archivo.files.length; $('#impPrevia').innerHTML = ''; };
    btn.onclick = () => conCarga(btn, () => this.analizar(archivo.files[0], $('#impCorte').value)).catch(() => { });
    pintarImportaciones($('#impHistorial'), 'existencias');
  },

  async analizar(archivo, corte) {
    if (!corte) throw new Error('Elige la fecha de corte.');
    const previa = $('#impPrevia');
    const hojas = await leerHojasExcel(archivo);
    let datos;
    try {
      datos = prepararFilas(hojas, COLS_EXISTENCIAS, OBLIGATORIAS_EXISTENCIAS);
    } catch (e) {
      previa.innerHTML = '<div class="caja-error"><b>No se puede importar este archivo.</b><br>' + esc(e.message) + '</div>';
      throw e;
    }
    this.payload = { archivo: archivo.name, fechaCorte: corte, filasLeidas: datos.filas.length, filas: datos.filas, errores: [] };
    this.info = datos;
    const r = await api('previsualizarExistencias', token(), this.payload);
    this.pintarPrevia(r);
  },

  pintarPrevia(r, aplicado) {
    const previa = $('#impPrevia');
    const res = r.resumen;
    let h = '<div class="pila">' +
      (aplicado
        ? '<div class="caja-ok"><i class="ti ti-circle-check"></i> <b>Importación aplicada.</b> Existencias al ' + fmtFecha(r.fechaCorte) + '. Las ubicaciones, fotos y notas no se modificaron.</div>'
        : '<div class="sec-titulo">Vista previa · corte ' + fmtFecha(r.fechaCorte) + '</div>' +
        '<div class="texto-3">Hoja "' + esc(this.info.hoja) + '", encabezados en la fila ' + this.info.filaEncabezado + (this.info.ignoradas ? ' · ' + this.info.ignoradas + ' filas de totales ignoradas' : '') + '</div>') +
      (r.advertencias || []).map(a => '<div class="caja-aviso"><i class="ti ti-alert-triangle"></i> ' + esc(a) + '</div>').join('') +
      '<div class="metricas">' +
      metrica(res.filasLeidas, 'Filas leídas') +
      metrica(res.materiales, 'Materiales') +
      metrica(res.nuevos, 'Materiales nuevos') +
      metrica(res.yaNoAparecen, 'Ya no aparecen') +
      metrica(res.cambios, 'Cambios de cantidad') +
      metrica(res.almacenesNuevos, 'Almacenes nuevos') +
      metrica(res.sumadas, 'Registros sumados') +
      metrica(res.errores, 'Filas con error', true) +
      '</div>' +
      detalleLista('Filas con error', r.errores, x => '<b>Fila ' + esc(x.fila) + ':</b> ' + esc(x.motivo), res.errores) +
      detalleLista('Materiales nuevos', r.nuevos, x => '<b class="serie">' + esc(x.s) + '</b> ' + esc(x.d) + ' <span class="texto-3">· ' + esc(x.f) + ' · ' + fmt(x.total) + ' ' + esc(x.u) + '</span>', res.nuevos) +
      detalleLista('Ya no aparecen (quedarán con existencia 0)', r.yaNoAparecen, x => '<b class="serie">' + esc(x.s) + '</b> ' + esc(x.d) + ' <span class="texto-3">· tenía ' + fmt(x.total) + '</span>', res.yaNoAparecen) +
      detalleLista('Cambios de cantidad', r.cambios, x => '<b class="serie">' + esc(x.s) + '</b> ' + esc(x.d) + '<div class="texto-3 num">Almacén ' + esc(x.a) + ': ' + fmt(x.antes) + ' → <b>' + fmt(x.despues) + '</b> (' + (x.dif > 0 ? '+' : '') + fmt(x.dif) + ')</div>', res.cambios) +
      detalleLista('Cambios de datos desde SAP', r.cambiosDatos, x => '<b class="serie">' + esc(x.s) + '</b> ' + esc(x.campo) + ': ' + esc(x.antes) + ' → <b>' + esc(x.despues) + '</b>', res.cambiosDatos) +
      detalleLista('Almacenes nuevos', r.almacenesNuevos, x => '<b>' + esc(x) + '</b> <span class="texto-3">(ponle nombre en Más › Almacenes)</span>') +
      detalleLista('Registros repetidos que se suman', r.sumadas, x => '<b class="serie">' + esc(x.s) + '</b> en ' + esc(x.a) + ': <span class="num">' + x.regs.map(fmt).join(' + ') + ' = <b>' + fmt(x.c) + '</b></span>', res.sumadas);

    if (!aplicado) {
      h += '<div class="caja-ok"><i class="ti ti-shield-check"></i> Solo se reemplazan las existencias. Ubicaciones, fotos, descripción detallada, peso, notas y nombres de almacén no se tocan.</div>' +
        '<div class="fila"><button class="btn crece" id="impCancelar">Cancelar</button>' +
        '<button class="btn btn-primario crece" id="impConfirmar" data-cargando="Aplicando…"' + (res.registros ? '' : ' disabled') + '><i class="ti ti-check"></i>Confirmar importación</button></div>';
    } else {
      h += '<button class="btn btn-primario" onclick="ir(\'catalogo\')"><i class="ti ti-layout-grid"></i>Ver catálogo</button>';
    }
    h += '</div>';
    previa.innerHTML = h;
    previa.scrollIntoView({ behavior: 'smooth', block: 'start' });

    if (!aplicado) {
      $('#impCancelar').onclick = () => { previa.innerHTML = ''; this.payload = null; };
      $('#impConfirmar').onclick = ev => conCarga(ev.currentTarget, async () => {
        const r2 = await api('confirmarExistencias', token(), this.payload);
        this.payload = null;
        this.pintarPrevia(r2, true);
        $('#impArchivo').value = '';
        $('#impAnalizar').disabled = true;
        await cargarCatalogo(true);
        pintarImportaciones($('#impHistorial'), 'existencias');
      }).catch(() => { });
    }
  }
};

// -------------------------------------------------------------------------
//  CARGA MASIVA DE UBICACIONES
// -------------------------------------------------------------------------

VISTAS.ubicMasiva = {
  titulo: 'Carga masiva de ubicaciones',
  rol: 'Administrador',
  padre: 'mas',
  nav: 'mas',
  render(app) {
    this.payload = null;
    app.innerHTML = '<div class="pila">' +
      '<div class="tarjeta pila"><b>1. Descarga la plantilla</b>' +
      '<div class="texto-2 pequeno">Trae todos los materiales con sus almacenes (solo como referencia). Llena Zona, Pasillo, Rack, Nivel, Posición y/o Referencia; o, si prefieres, escribe todo en la columna "Ubicación". Para una segunda ubicación del mismo material, copia la fila. Las filas que dejes vacías no cambian nada.</div>' +
      '<label class="check"><input type="checkbox" id="umSoloSin" checked>Solo materiales sin ubicación</label>' +
      '<button class="btn" id="umPlantilla" data-cargando="Generando…"><i class="ti ti-download"></i>Descargar plantilla</button></div>' +
      '<div class="tarjeta pila"><b>2. Sube la plantilla llena</b>' +
      '<input class="entrada" type="file" id="umArchivo" accept=".xls,.xlsx,.csv">' +
      '<button class="btn btn-primario" id="umAnalizar" data-cargando="Analizando…" disabled><i class="ti ti-file-search"></i>Ver vista previa</button></div>' +
      '<div id="umPrevia"></div><div id="umHistorial"></div></div>';
    $('#umPlantilla').onclick = ev => conCarga(ev.currentTarget, () => descargarPlantillaUbicaciones($('#umSoloSin').checked)).catch(() => { });
    const archivo = $('#umArchivo'), btn = $('#umAnalizar');
    archivo.onchange = () => { btn.disabled = !archivo.files.length; $('#umPrevia').innerHTML = ''; };
    btn.onclick = () => conCarga(btn, () => this.analizar(archivo.files[0])).catch(() => { });
    pintarImportaciones($('#umHistorial'), 'ubicaciones');
  },

  async analizar(archivo) {
    const previa = $('#umPrevia');
    const hojas = await leerHojasExcel(archivo);
    let datos;
    try {
      datos = prepararFilas(hojas, COLS_UBICACIONES, ['s']);
      const tieneAlguna = ['zona', 'pasillo', 'rack', 'nivel', 'posicion', 'referencia', 'libre'].some(c => datos.columnas[c] !== undefined);
      if (!tieneAlguna) throw new Error('No se encontró ninguna columna de ubicación (Zona, Pasillo, Rack, Nivel, Posición, Referencia o Ubicación).');
    } catch (e) {
      previa.innerHTML = '<div class="caja-error"><b>No se puede usar este archivo.</b><br>' + esc(e.message) + '</div>';
      throw e;
    }
    this.payload = { archivo: archivo.name, filasLeidas: datos.filas.length, filas: datos.filas, errores: [] };
    const r = await api('previsualizarUbicaciones', token(), this.payload);
    this.pintarPrevia(r);
  },

  pintarPrevia(r, aplicado) {
    const previa = $('#umPrevia');
    const res = r.resumen;
    const fila = x => '<b class="serie">' + esc(x.s) + '</b> ' + esc(x.d) + '<div class="texto-3">' + (x.antes ? esc(x.antes) + ' → ' : '') + '<b>' + esc(x.despues) + '</b></div>';
    let h = '<div class="pila">' +
      (aplicado ? '<div class="caja-ok"><i class="ti ti-circle-check"></i> <b>Ubicaciones aplicadas.</b></div>' : '<div class="sec-titulo">Vista previa</div>') +
      '<div class="metricas">' +
      metrica(res.filasLeidas, 'Filas leídas') + metrica(res.nuevas, 'Ubicaciones nuevas') + metrica(res.cambios, 'Ubicaciones que cambian') +
      metrica(res.sinCambios, 'Sin cambios') + metrica(res.vacias, 'Filas vacías (se ignoran)') + metrica(res.errores, 'Filas con error', true) +
      '</div>' +
      detalleLista('Filas con error', r.errores, x => '<b>Fila ' + esc(x.fila) + ':</b> ' + esc(x.motivo), res.errores) +
      detalleLista('Ubicaciones nuevas', r.nuevas, fila, res.nuevas) +
      detalleLista('Ubicaciones que cambian (la anterior queda en el historial)', r.cambios, fila, res.cambios);
    if (!aplicado) {
      h += '<div class="fila"><button class="btn crece" id="umCancelar">Cancelar</button>' +
        '<button class="btn btn-primario crece" id="umConfirmar" data-cargando="Aplicando…"' + (res.nuevas + res.cambios ? '' : ' disabled') + '><i class="ti ti-check"></i>Confirmar</button></div>';
    }
    h += '</div>';
    previa.innerHTML = h;
    previa.scrollIntoView({ behavior: 'smooth', block: 'start' });
    if (!aplicado) {
      $('#umCancelar').onclick = () => { previa.innerHTML = ''; this.payload = null; };
      $('#umConfirmar').onclick = ev => conCarga(ev.currentTarget, async () => {
        const r2 = await api('confirmarUbicaciones', token(), this.payload);
        this.payload = null;
        this.pintarPrevia(r2, true);
        $('#umArchivo').value = '';
        $('#umAnalizar').disabled = true;
        await cargarCatalogo(true);
        pintarImportaciones($('#umHistorial'), 'ubicaciones');
      }).catch(() => { });
    }
  }
};

// -------------------------------------------------------------------------
//  Descargas a Excel
// -------------------------------------------------------------------------

/** Guarda un libro de Excel en el dispositivo. */
async function descargarExcel(nombreArchivo, hojas) {
  const XLSX = await cargarSheetJS();
  const libro = XLSX.utils.book_new();
  hojas.forEach(h => {
    const ws = XLSX.utils.aoa_to_sheet(h.filas);
    if (h.anchos) ws['!cols'] = h.anchos.map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(libro, ws, h.nombre.substring(0, 31));
  });
  XLSX.writeFile(libro, nombreArchivo);
}

/** Plantilla para la carga masiva de ubicaciones (un renglón por material o por ubicación ya capturada). */
async function descargarPlantillaUbicaciones(soloSin) {
  const mats = materialesVisibles().filter(m => m.act && (!soloSin || !m.ub.length))
    .sort((a, b) => a.f.localeCompare(b.f) || a.d.localeCompare(b.d));
  const filas = [['NÚMERO DE SERIE', 'Texto breve material', 'Familia', 'Almacenes (referencia)', 'Existencia total', 'Zona', 'Pasillo', 'Rack', 'Nivel', 'Posición', 'Referencia', 'Ubicación']];
  mats.forEach(m => {
    const alms = m.ex.map(x => x[0]).join(', ');
    const base = [m.s, m.d, m.f, alms, m.total];
    if (m.ub.length) m.ub.forEach(u => filas.push(base.concat([u.z, u.pa, u.r, u.ni, u.po, u.ref, ''])));
    else filas.push(base.concat(['', '', '', '', '', '', '']));
  });
  await descargarExcel('Plantilla_ubicaciones_' + hoyISO() + '.xlsx', [{
    nombre: 'Ubicaciones', filas, anchos: [16, 38, 22, 22, 14, 10, 10, 10, 10, 10, 28, 28]
  }, {
    nombre: 'Instrucciones', filas: [
      ['Cómo llenar la plantilla'],
      ['1. No cambies los nombres de las columnas ni la columna NÚMERO DE SERIE.'],
      ['2. Llena los campos que apliquen: Zona/Bodega, Pasillo, Rack, Nivel, Posición, Referencia. Todos son opcionales.'],
      ['3. Si prefieres texto libre, usa solo la columna "Ubicación".'],
      ['4. Si un material está en dos lugares, copia su fila y escribe la segunda ubicación.'],
      ['5. Las filas sin datos de ubicación se ignoran: no borran nada.'],
      ['6. Si un material ya tenía ubicación y escribes otra, la nueva la reemplaza (la anterior queda en el historial).'],
      ['Ejemplo: Zona B1, Pasillo 3, Rack 02, Nivel 4 → se mostrará "B1 · P3 · R02 · N4".']
    ], anchos: [110]
  }]);
  aviso('Plantilla descargada (' + fmt(mats.length) + ' materiales)');
}

/** Exporta una lista de materiales (catálogo filtrado o reporte). */
async function exportarMateriales(lista, nombre) {
  if (!lista.length) { aviso('No hay materiales para exportar.', 'error'); return; }
  const almacenes = Array.from(new Set(lista.flatMap(m => m.ex.map(x => x[0])))).sort();
  const filas = [['NÚMERO DE SERIE', 'Texto breve material', 'Familia', 'Unidad', 'Existencia total']
    .concat(almacenes.map(a => 'Almacén ' + a))
    .concat(['Ubicación', 'Referencia', 'Peso (kg)', 'Descripción detallada', 'Notas', 'Etiqueta impresa', 'Fecha de corte'])];
  lista.forEach(m => {
    const porAlm = {};
    m.ex.forEach(x => { porAlm[x[0]] = x[1]; });
    filas.push([m.s, m.d, m.f, m.u, m.total]
      .concat(almacenes.map(a => porAlm[a] === undefined ? '' : porAlm[a]))
      .concat([m.ub.map(u => u.t).join(' / '), m.ub.map(u => u.ref).filter(Boolean).join(' / '), m.p === '' ? '' : Number(m.p), m.dd, m.n, m.et ? fmtFecha(m.et) : '', E.cat.corte ? fmtFecha(E.cat.corte) : '']));
  });
  try {
    await descargarExcel(nombre + '_' + hoyISO() + '.xlsx', [{ nombre, filas }]);
    aviso('Excel descargado (' + fmt(lista.length) + ' materiales)');
  } catch (e) { aviso(e.message, 'error'); }
}


/* ######################################################################
   4. QR Y ETIQUETAS
   ###################################################################### */
/* =========================================================================
   JsEtiquetas.html – Códigos QR, escáner por foto e impresión de etiquetas
   ========================================================================= */

// -------------------------------------------------------------------------
//  Generar QR
// -------------------------------------------------------------------------

/**
 * URL que va dentro del QR de un material:  <URL fija de la página>?serie=158
 * Al escanearla con la cámara normal del celular se abre directo la ficha.
 */
function urlMaterial(serie) {
  const base = (E.cat && E.cat.url) || urlPaginaActual();
  if (!base) return '';
  return base + (base.indexOf('?') >= 0 ? '&' : '?') + 'serie=' + encodeURIComponent(serie);
}

/** Dirección de esta página sin parámetros: https://usuario.github.io/catalogo/ */
function urlPaginaActual() {
  if (!/^https?:$/.test(location.protocol)) return '';
  return location.origin + location.pathname.replace(/index\.html$/, '');
}

/** Dibuja un QR como SVG (nítido en pantalla y en impresión). */
function qrSVG(texto, estilo) {
  if (!texto || typeof qrcode !== 'function') {
    return '<svg viewBox="0 0 10 10" ' + (estilo || '') + ' role="img" aria-label="QR no disponible"><rect width="10" height="10" fill="#EEF1EF"/><text x="5" y="5.8" font-size="2" text-anchor="middle" fill="#8A948E">Sin URL</text></svg>';
  }
  const qr = qrcode(0, 'M');     // tamaño automático, corrección de errores media (aguanta rayones)
  qr.addData(texto);
  qr.make();
  const n = qr.getModuleCount();
  const margen = 2;               // zona blanca alrededor (necesaria para leerlo)
  const total = n + margen * 2;
  let d = '';
  for (let f = 0; f < n; f++) {
    for (let c = 0; c < n; c++) {
      if (qr.isDark(f, c)) d += 'M' + (c + margen) + ' ' + (f + margen) + 'h1v1h-1z';
    }
  }
  return '<svg viewBox="0 0 ' + total + ' ' + total + '" ' + (estilo || '') + ' shape-rendering="crispEdges" role="img" aria-label="Código QR">' +
    '<rect width="' + total + '" height="' + total + '" fill="#fff"/><path d="' + d + '" fill="#000"/></svg>';
}

// -------------------------------------------------------------------------
//  Escanear QR
//  1) Cámara en vivo (getUserMedia): se apunta al código y se lee solo.
//  2) Respaldo: tomar una foto (<input capture>) y leerla con jsQR,
//     si el navegador no permite la cámara en vivo o se negó el permiso.
// -------------------------------------------------------------------------

let _camara = null; // { stream, activo, detector }

function abrirEscaner() {
  if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia && window.isSecureContext) {
    abrirCamaraEnVivo();
  } else {
    abrirEscanerFoto();
  }
}

/** Respaldo: abre la cámara del celular para tomar una foto. */
function abrirEscanerFoto() {
  cerrarModal();
  const input = $('#inputEscaner');
  input.value = '';
  input.click();
}

async function abrirCamaraEnVivo() {
  abrirModal('Escanear QR',
    '<div class="escaner">' +
    '<video id="video" playsinline muted autoplay></video>' +
    '<div class="escaner-marco"></div>' +
    '<div class="escaner-texto" id="escTexto">Abriendo la cámara…</div>' +
    '<div class="escaner-pie">' +
    '<button class="btn" id="escCerrar"><i class="ti ti-x"></i>Cerrar</button>' +
    '<button class="btn" id="escFoto"><i class="ti ti-camera"></i>Tomar foto</button>' +
    '</div></div>', m => {
      $('#escCerrar', m).onclick = cerrarModal;
      $('#escFoto', m).onclick = abrirEscanerFoto;
    }, 'modal-escaner');

  try {
    const stream = await navigator.mediaDevices.getUserMedia({
      video: { facingMode: { ideal: 'environment' }, width: { ideal: 1280 }, height: { ideal: 720 } },
      audio: false
    });
    const video = $('#video');
    if (!video) { stream.getTracks().forEach(t => t.stop()); return; } // se cerró mientras pedía permiso
    _camara = { stream, activo: true, detector: null };
    if ('BarcodeDetector' in window) {
      try { _camara.detector = new BarcodeDetector({ formats: ['qr_code'] }); } catch (e) { _camara.detector = null; }
    }
    video.srcObject = stream;
    await video.play();
    $('#escTexto').textContent = 'Apunta al código QR de la etiqueta';
    buscarQREnVideo(video);
  } catch (e) {
    detenerCamara();
    const texto = $('#escTexto');
    if (texto) {
      texto.innerHTML = (e && e.name === 'NotAllowedError')
        ? 'No diste permiso para usar la cámara.<br>Puedes permitirlo en la configuración del navegador, o tomar una foto.'
        : 'No se pudo abrir la cámara.<br>Usa “Tomar foto”.';
    }
  }
}

/** Revisa los cuadros del video hasta encontrar un QR. */
function buscarQREnVideo(video) {
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  let ocupado = false;
  const ciclo = async () => {
    if (!_camara || !_camara.activo) return;
    if (!ocupado && video.readyState >= 2 && video.videoWidth) {
      ocupado = true;
      let texto = null;
      try {
        if (_camara.detector) {
          const r = await _camara.detector.detect(video);
          if (r && r.length) texto = r[0].rawValue;
        } else if (typeof jsQR === 'function') {
          const escala = Math.min(1, 640 / Math.max(video.videoWidth, video.videoHeight));
          canvas.width = Math.round(video.videoWidth * escala);
          canvas.height = Math.round(video.videoHeight * escala);
          ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
          const img = ctx.getImageData(0, 0, canvas.width, canvas.height);
          const r = jsQR(img.data, canvas.width, canvas.height, { inversionAttempts: 'dontInvert' });
          if (r && r.data) texto = r.data;
        }
      } catch (e) { /* cuadro ilegible: se intenta con el siguiente */ }
      ocupado = false;
      if (texto) return alLeerQR(texto);
    }
    setTimeout(() => requestAnimationFrame(ciclo), 120);
  };
  requestAnimationFrame(ciclo);
}

function alLeerQR(texto) {
  const serie = interpretarQR(texto);
  if (navigator.vibrate) navigator.vibrate(80);
  cerrarModal();
  if (serie) ir('ficha', { serie: normSerie(serie) });
  else mostrarFalloEscaneo(texto);
}

/** Apaga la cámara (se llama al cerrar la ventana). */
function detenerCamara() {
  if (!_camara) return;
  _camara.activo = false;
  try { _camara.stream.getTracks().forEach(t => t.stop()); } catch (e) { }
  _camara = null;
}

async function alElegirFotoQR(ev) {
  const archivo = ev.target.files && ev.target.files[0];
  if (!archivo) return;
  abrirModal('Leyendo código', '<div class="cargando-pantalla" style="padding:24px"><div class="spinner"></div><p>Buscando el código QR en la foto…</p></div>');
  try {
    const texto = await decodificarQR(archivo);
    const serie = texto ? interpretarQR(texto) : null;
    cerrarModal();
    if (serie) {
      ir('ficha', { serie: normSerie(serie) });
      return;
    }
    mostrarFalloEscaneo(texto);
  } catch (e) {
    cerrarModal();
    mostrarFalloEscaneo(null, e.message);
  }
}

/** Carga la foto respetando la orientación del celular. */
function cargarImagen(archivo) {
  if (window.createImageBitmap) {
    return createImageBitmap(archivo, { imageOrientation: 'from-image' }).catch(() => cargarImagenElemento(archivo));
  }
  return cargarImagenElemento(archivo);
}
function cargarImagenElemento(archivo) {
  return new Promise((res, rej) => {
    const url = URL.createObjectURL(archivo);
    const img = new Image();
    img.onload = () => { URL.revokeObjectURL(url); res(img); };
    img.onerror = () => { URL.revokeObjectURL(url); rej(new Error('No se pudo abrir la foto.')); };
    img.src = url;
  });
}

/** Intenta leer el QR a varios tamaños (las fotos del celular son muy grandes). */
async function decodificarQR(archivo) {
  if (typeof jsQR !== 'function') throw new Error('No se cargó el lector de QR. Revisa tu conexión.');
  const img = await cargarImagen(archivo);
  const ancho = img.width, alto = img.height;
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  for (const lado of [1000, 1400, 700, 1800]) {
    const escala = Math.min(1, lado / Math.max(ancho, alto));
    canvas.width = Math.round(ancho * escala);
    canvas.height = Math.round(alto * escala);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    const datos = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const r = jsQR(datos.data, canvas.width, canvas.height, { inversionAttempts: 'attemptBoth' });
    if (r && r.data) return r.data;
    if (escala === 1) break;
  }
  return null;
}

/** Del texto del QR obtiene el número de serie. Acepta la URL de la app o solo el número. */
function interpretarQR(texto) {
  texto = String(texto || '').trim();
  try {
    const u = new URL(texto);
    const s = u.searchParams.get('serie');
    if (s) return s;
  } catch (e) { /* no es URL */ }
  const m = texto.match(/[?&]serie=([^&#\s]+)/i);
  if (m) return decodeURIComponent(m[1]);
  if (/^[0-9A-Za-z-]{1,30}$/.test(texto)) return texto;
  return null;
}

function mostrarFalloEscaneo(textoLeido, error) {
  abrirModal('No se pudo leer el QR',
    (textoLeido
      ? '<div class="caja-aviso">El código dice <b>' + esc(textoLeido.substring(0, 120)) + '</b>, pero no es una etiqueta de este catálogo.</div>'
      : '<div class="caja-error">' + esc(error || 'No se encontró un código QR en la foto.') + '</div>') +
    '<ul class="texto-2" style="padding-left:20px">' +
    '<li>Acerca el celular a unos 15–20 cm: que el QR ocupe buena parte de la foto.</li>' +
    '<li>Busca buena luz y evita reflejos sobre la etiqueta.</li>' +
    '<li>También puedes usar la cámara normal del celular: el QR abre la ficha directamente.</li></ul>' +
    '<div class="modal-pie"><button class="btn" onclick="cerrarModal();ir(\'buscar\')"><i class="ti ti-search"></i>Buscar por serie</button>' +
    '<button class="btn btn-primario" onclick="cerrarModal();abrirEscaner()"><i class="ti ti-qrcode"></i>Escanear otra vez</button></div>');
}

// -------------------------------------------------------------------------
//  IMPRESIÓN DE ETIQUETAS
// -------------------------------------------------------------------------

const TAMANOS_ETIQUETA = [
  ['5x3', '5 × 3 cm', 5, 3],
  ['4x2.5', '4 × 2.5 cm', 4, 2.5],
  ['6x4', '6 × 4 cm', 6, 4],
  ['7x5', '7 × 5 cm', 7, 5],
  ['10x5', '10 × 5 cm', 10, 5],
  ['otro', 'Personalizado', 0, 0]
];

/** Preferencias de impresión (se recuerdan en el dispositivo). */
const ConfEtiquetas = {
  leer() {
    const def = { tam: '5x3', ancho: 5, alto: 3, copias: 1, sep: 2, margen: 8, bordes: true, ubic: true };
    try { return Object.assign(def, JSON.parse(Local.get('etiquetas') || '{}')); } catch (e) { return def; }
  },
  guardar(c) { Local.set('etiquetas', JSON.stringify(c)); }
};

VISTAS.etiquetas = {
  titulo: 'Imprimir etiquetas',
  rol: 'Administrador',
  padre: 'mas',
  nav: 'mas',
  render(app) {
    const p = E.params;
    this.modo = p.modo || 'familia';
    this.series = p.series || null;
    const c = ConfEtiquetas.leer();
    const almacenes = Array.from(E.alm.values()).sort((a, b) => a.c.localeCompare(b.c));
    const modos = [
      ['seleccion', 'Selección del catálogo (' + E.seleccion.size + ')'],
      ['familia', 'Por familia'],
      ['almacen', 'Por almacén'],
      ['pendientes', 'Pendientes de imprimir']
    ];
    if (this.series) modos.unshift(['series', 'Este material']);

    let avisoUrl = '';
    if (!E.cat.urlFija) avisoUrl = '<div class="caja-aviso"><i class="ti ti-alert-triangle"></i> Aún no guardas la URL fija de la página. Los QR usarán <b>' + esc(urlPaginaActual()) + '</b>. Confírmala en Más › Configuración antes de imprimir.</div>';

    app.innerHTML = '<div class="pila">' + avisoUrl +
      '<div class="tarjeta pila"><b>¿Qué materiales?</b>' +
      '<div class="pila" style="gap:0">' + modos.map(m => '<label class="check"><input type="radio" name="etModo" value="' + m[0] + '"' + (this.modo === m[0] ? ' checked' : '') + '>' + esc(m[1]) + '</label>').join('') + '</div>' +
      '<div class="rejilla-2">' +
      '<div class="campo" id="etFamCampo"><label>Familia</label><select class="entrada" id="etFam"><option value="">Todas</option>' + E.familias.map(x => '<option>' + esc(x.f) + '</option>').join('') + '</select></div>' +
      '<div class="campo" id="etAlmCampo"><label>Almacén</label><select class="entrada" id="etAlm">' + almacenes.map(a => '<option value="' + esc(a.c) + '">' + esc(a.c + (a.n ? ' · ' + a.n : '')) + '</option>').join('') + '</select></div>' +
      '</div>' +
      (this.modo === 'seleccion' ? '' : '<div class="texto-3">Para escoger materiales uno por uno: Catálogo › Seleccionar.</div>') +
      '</div>' +

      '<div class="tarjeta pila"><b>Tamaño y formato</b>' +
      '<div class="rejilla-2">' +
      '<div class="campo"><label>Tamaño</label><select class="entrada" id="etTam">' + TAMANOS_ETIQUETA.map(t => '<option value="' + t[0] + '"' + (c.tam === t[0] ? ' selected' : '') + '>' + t[1] + '</option>').join('') + '</select></div>' +
      '<div class="campo"><label>Copias por material</label><input class="entrada" id="etCopias" type="number" min="1" max="20" value="' + c.copias + '"></div>' +
      '<div class="campo" id="etMedidas"><label>Ancho × alto (cm)</label><div class="fila"><input class="entrada" id="etAncho" type="number" step="0.1" min="2" max="20" value="' + c.ancho + '"><span>×</span><input class="entrada" id="etAlto" type="number" step="0.1" min="1.5" max="15" value="' + c.alto + '"></div></div>' +
      '<div class="campo"><label>Separación (mm)</label><input class="entrada" id="etSep" type="number" min="0" max="20" value="' + c.sep + '"></div>' +
      '<div class="campo"><label>Margen de la hoja (mm)</label><input class="entrada" id="etMargen" type="number" min="0" max="30" value="' + c.margen + '"></div>' +
      '</div>' +
      '<label class="check"><input type="checkbox" id="etUbic"' + (c.ubic ? ' checked' : '') + '>Incluir ubicación física</label>' +
      '<label class="check"><input type="checkbox" id="etBordes"' + (c.bordes ? ' checked' : '') + '>Mostrar líneas de corte</label>' +
      '</div>' +

      '<div class="fila-sep"><b id="etCuenta"></b><button class="btn btn-primario" id="etImprimir"><i class="ti ti-printer"></i>Imprimir</button></div>' +
      '<div class="texto-3">Consejo: imprime desde una computadora, en papel adhesivo, con escala al 100% ("Tamaño real") y sin encabezados ni pies de página.</div>' +
      '<div id="etMarcar"></div>' +
      '<div class="sec-titulo">Vista previa</div><div class="hoja-previa"><div id="etPrevia"></div></div>' +
      '</div>';

    const repintar = () => this.pintar();
    $$('input[name=etModo]', app).forEach(r => r.onchange = () => { this.modo = r.value; this.mostrarCampos(); repintar(); });
    ['#etFam', '#etAlm', '#etTam', '#etCopias', '#etAncho', '#etAlto', '#etSep', '#etMargen', '#etUbic', '#etBordes']
      .forEach(s => $(s).addEventListener('change', () => {
        if (s === '#etTam') {
          const t = TAMANOS_ETIQUETA.find(x => x[0] === $('#etTam').value);
          if (t && t[2]) { $('#etAncho').value = t[2]; $('#etAlto').value = t[3]; }
        }
        this.mostrarCampos();
        repintar();
      }));
    $('#etImprimir').onclick = () => this.imprimir();
    this.mostrarCampos();
    repintar();
  },

  mostrarCampos() {
    $('#etFamCampo').hidden = this.modo !== 'familia' && this.modo !== 'pendientes';
    $('#etAlmCampo').hidden = this.modo !== 'almacen';
    $('#etMedidas').style.opacity = $('#etTam').value === 'otro' ? '1' : '.6';
  },

  config() {
    const c = {
      tam: $('#etTam').value,
      ancho: Math.max(2, Number($('#etAncho').value) || 5),
      alto: Math.max(1.5, Number($('#etAlto').value) || 3),
      copias: Math.min(20, Math.max(1, parseInt($('#etCopias').value, 10) || 1)),
      sep: Math.max(0, Number($('#etSep').value) || 0),
      margen: Math.max(0, Number($('#etMargen').value) || 0),
      ubic: $('#etUbic').checked,
      bordes: $('#etBordes').checked
    };
    ConfEtiquetas.guardar(c);
    return c;
  },

  /** Materiales que se van a imprimir. */
  materiales() {
    const visibles = materialesVisibles().filter(m => m.act);
    const fam = $('#etFam') ? $('#etFam').value : '';
    let lista;
    switch (this.modo) {
      case 'series': lista = (this.series || []).map(materialPorSerie).filter(Boolean); break;
      case 'seleccion': lista = visibles.filter(m => E.seleccion.has(m.s)); break;
      case 'almacen': { const a = $('#etAlm').value; lista = visibles.filter(m => m.ex.some(x => x[0] === a)); break; }
      case 'pendientes': lista = visibles.filter(m => !m.et && (!fam || m.f === fam)); break;
      default: lista = visibles.filter(m => !fam || m.f === fam);
    }
    return lista.sort((a, b) => a.f.localeCompare(b.f) || a.d.localeCompare(b.d));
  },

  /** HTML de UNA etiqueta: QR + serie + nombre + ubicación. */
  etiquetaHTML(m, c) {
    const pad = 1.5; // mm
    const ladoQR = c.alto * 10 - pad * 2; // mm
    const pt = (factor, min, max) => Math.min(max, Math.max(min, c.alto * factor)).toFixed(1) + 'pt';
    const ubic = c.ubic && m.ub.length ? m.ub.map(u => u.t).join(' / ') : '';
    return '<div class="etiqueta' + (c.bordes ? ' borde' : '') + '" style="width:' + c.ancho + 'cm;height:' + c.alto + 'cm;padding:' + pad + 'mm;gap:' + pad + 'mm;margin:' + (c.sep / 2) + 'mm">' +
      qrSVG(urlMaterial(m.s), 'style="width:' + ladoQR + 'mm;height:' + ladoQR + 'mm"') +
      '<div class="et-txt">' +
      '<div class="et-serie" style="font-size:' + pt(3.6, 7, 20) + '">' + esc(m.s) + '</div>' +
      '<div class="et-nombre" style="font-size:' + pt(2.4, 5.5, 12) + '">' + esc(m.d) + '</div>' +
      (ubic ? '<div class="et-ubic" style="font-size:' + pt(2.2, 5.5, 11) + '">' + esc(ubic) + '</div>' : '') +
      (c.alto >= 4 || c.ancho >= 7 ? '<div class="et-marca">CFE · Almacén Zona Campeche</div>' : '') +
      '</div></div>';
  },

  pintar() {
    const c = this.config();
    const lista = this.materiales();
    const n = lista.length * c.copias;
    $('#etCuenta').textContent = fmt(lista.length) + ' materiales · ' + fmt(n) + ' etiquetas';
    $('#etImprimir').disabled = !lista.length;
    const muestra = [];
    lista.slice(0, 24).forEach(m => { for (let i = 0; i < Math.min(c.copias, 2); i++) muestra.push(this.etiquetaHTML(m, c)); });
    $('#etPrevia').innerHTML = lista.length
      ? '<div class="etiquetas">' + muestra.join('') + '</div>' + (lista.length > 24 ? '<p class="texto-3">Vista previa de los primeros 24 materiales.</p>' : '')
      : vacioHTML('ti-tags-off', 'No hay materiales', this.modo === 'seleccion' ? 'Ve al catálogo, pulsa "Seleccionar" y marca los materiales.' : 'Cambia el filtro para elegir materiales.');
  },

  imprimir() {
    const c = this.config();
    const lista = this.materiales();
    if (!lista.length) return;
    const html = [];
    lista.forEach(m => { for (let i = 0; i < c.copias; i++) html.push(this.etiquetaHTML(m, c)); });
    let estilo = $('#estiloPagina');
    if (!estilo) { estilo = document.createElement('style'); estilo.id = 'estiloPagina'; document.head.appendChild(estilo); }
    estilo.textContent = '@media print { @page { margin: ' + c.margen + 'mm; } }';
    $('#impresion').innerHTML = '<div class="etiquetas">' + html.join('') + '</div>';
    setTimeout(() => {
      window.print();
      this.ofrecerMarcar(lista);
    }, 150);
  },

  ofrecerMarcar(lista) {
    const c = $('#etMarcar');
    if (!c) return;
    c.innerHTML = '<div class="caja-ok fila-sep"><span>¿Se imprimieron bien las ' + fmt(lista.length) + ' etiquetas?</span>' +
      '<button class="btn btn-chico btn-primario" id="etMarcarBtn">Marcar como impresas</button></div>';
    $('#etMarcarBtn').onclick = ev => conCarga(ev.currentTarget, async () => {
      const n = await api('marcarEtiquetasImpresas', token(), lista.map(m => m.s));
      c.innerHTML = '';
      if (this.modo === 'seleccion') { E.seleccion.clear(); E.seleccionando = false; }
      await despuesDeGuardar(fmt(n) + ' etiquetas marcadas como impresas');
    }).catch(() => { });
  },

  refrescar() { if ($('#etPrevia')) this.pintar(); else render(); }
};
