/**
 * CLUB DMD — Tarjeta de fidelización digital
 * Backend único para todos los locales (Dany Marketing Digital).
 *
 * Vive dentro del Sheet central "CLUB DMD - Central", que tiene la pestaña
 * "Locales" (slug | sheetId | pin_caja | activo | nombre).
 * Cada local tiene su propio Sheet con las pestañas:
 *   Config, Premios, Novedades, Clientes, Visitas, Cupones
 *
 * Instalación: ver INSTRUCCIONES.md
 */

// Correo de Dany Marketing Digital que recibe los informes mensuales
var CORREO_INFORMES = 'dpizarro.ascui@gmail.com';

var TABS = {
  Config: ['Clave', 'Valor', 'Ayuda'],
  Premios: ['Visita', 'Emoji', 'Premio', 'Vigencia (días)'],
  Novedades: ['Texto', 'Imagen URL', 'Link'],
  Clientes: ['ID', 'Nombre', 'Teléfono', 'Correo', 'Registro', 'Visitas', 'Última visita', 'Tarjetas completadas', 'Visitas totales', 'Acepta promociones'],
  Visitas: ['Fecha', 'ID cliente', 'Nombre', 'Visita N°', 'Nota'],
  Cupones: ['Código', 'ID cliente', 'Nombre', 'Emoji', 'Premio', 'Visita', 'Creado', 'Vence', 'Estado', 'Usado el']
};

var CONFIG_DEFAULT = [
  ['nombre', 'Mi Local', 'Nombre que aparece en la tarjeta'],
  ['club', 'Club de Visitas', 'Nombre del programa'],
  ['logo_url', '', 'Link directo a la imagen del logo (opcional)'],
  ['color', '#FF8A2A', 'Color principal (hex)'],
  ['color2', '#35C6E8', 'Color secundario (hex)'],
  ['visitas_meta', '10', 'Cantidad de casillas de la tarjeta'],
  ['descripcion', 'Junta visitas y gana premios en cada etapa de tu tarjeta.', 'Texto explicativo'],
  ['instagram', '', 'Usuario sin @ (opcional)'],
  ['whatsapp', '', 'Número con código país, ej: 56912345678 (opcional)'],
  ['una_visita_por_dia', 'SI', 'SI = solo se suma 1 visita por día por cliente']
];

var PREMIOS_DEFAULT = [
  [3, '🥤', 'Jugo o café de regalo', 30],
  [6, '🎁', 'Regalo sorpresa', 30],
  [8, '🛝', 'Zona de juegos gratis', 30],
  [10, '⭐', 'Regalo gratis', 30]
];

/* ---------------- Web app ---------------- */

function doGet(e) {
  return handle_(e.parameter || {});
}

function doPost(e) {
  var data = {};
  try { data = JSON.parse(e.postData.contents || '{}'); } catch (err) {}
  return handle_(data);
}

function handle_(p) {
  var out;
  try {
    out = route_(p);
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function route_(p) {
  var local = getLocal_(p.local);
  var ss = SpreadsheetApp.openById(local.sheetId);
  switch (p.action) {
    case 'config': return { ok: true, config: readConfig_(ss) };
    case 'registrar': return withLock_(function () { return registrar_(ss, p); });
    case 'login': return login_(ss, p.ident);
    case 'cliente': return clienteResp_(ss, p.id);
    case 'caja_login': checkPin_(local, p.pin); return { ok: true, config: readConfig_(ss) };
    case 'caja_buscar': checkPin_(local, p.pin); return login_(ss, p.q);
    case 'caja_visita': checkPin_(local, p.pin); return withLock_(function () { return sumarVisita_(ss, p.id, p.forzar); });
    case 'caja_canjear': checkPin_(local, p.pin); return withLock_(function () { return canjear_(ss, p.codigo); });
    default: throw new Error('Acción no válida');
  }
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------------- Locales ---------------- */

function getLocal_(slug) {
  slug = String(slug || '').trim().toLowerCase();
  if (!slug) throw new Error('Falta el local');
  var sh = SpreadsheetApp.getActive().getSheetByName('Locales');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === slug) {
      if (String(rows[i][3]).toUpperCase() === 'NO') throw new Error('Este club no está activo');
      return { slug: slug, sheetId: rows[i][1], pin: String(rows[i][2]) };
    }
  }
  throw new Error('Local no encontrado');
}

function checkPin_(local, pin) {
  if (String(pin || '') !== local.pin) throw new Error('PIN incorrecto');
}

/* ---------------- Lectura ---------------- */

function readConfig_(ss) {
  var cfg = {};
  ss.getSheetByName('Config').getDataRange().getValues().slice(1).forEach(function (r) {
    if (r[0]) cfg[String(r[0]).trim()] = String(r[1]);
  });
  cfg.visitas_meta = parseInt(cfg.visitas_meta, 10) || 10;
  cfg.premios = ss.getSheetByName('Premios').getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0] && r[2]; })
    .map(function (r) { return { visita: Number(r[0]), emoji: String(r[1] || '🎁'), premio: String(r[2]), vigencia: Number(r[3]) || 30 }; });
  var nov = ss.getSheetByName('Novedades');
  cfg.novedades = nov ? nov.getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0] || r[1]; })
    .map(function (r) { return { texto: String(r[0]), imagen: String(r[1]), link: String(r[2]) }; }) : [];
  return cfg;
}

function clientesData_(ss) {
  var sh = ss.getSheetByName('Clientes');
  return { sh: sh, rows: sh.getDataRange().getValues() };
}

function rowToCliente_(r) {
  return {
    id: String(r[0]), nombre: String(r[1]), telefono: String(r[2]), correo: String(r[3]),
    visitas: Number(r[5]) || 0, ultima: r[6] ? fmtDate_(r[6]) : '',
    completadas: Number(r[7]) || 0, totales: Number(r[8]) || 0
  };
}

function findCliente_(ss, q) {
  q = String(q || '').trim().toLowerCase();
  if (!q) return null;
  var tel = normTel_(q);
  var d = clientesData_(ss);
  for (var i = 1; i < d.rows.length; i++) {
    var r = d.rows[i];
    if (String(r[0]).toLowerCase() === q ||
        (tel && normTel_(r[2]) === tel) ||
        String(r[3]).trim().toLowerCase() === q) {
      return { row: i + 1, data: r };
    }
  }
  return null;
}

function cuponesDe_(ss, id) {
  var hoy = new Date();
  return ss.getSheetByName('Cupones').getDataRange().getValues().slice(1)
    .filter(function (r) { return String(r[1]) === String(id); })
    .map(function (r) {
      var estado = String(r[8] || 'disponible');
      if (estado === 'disponible' && r[7] && new Date(r[7]) < hoy) estado = 'vencido';
      return { codigo: String(r[0]), emoji: String(r[3]), premio: String(r[4]), visita: Number(r[5]),
        creado: fmtDate_(r[6]), vence: r[7] ? fmtDate_(r[7]) : '', estado: estado, usado: r[9] ? fmtDate_(r[9]) : '' };
    })
    .reverse();
}

function clienteResp_(ss, id) {
  var f = findCliente_(ss, id);
  if (!f) throw new Error('No encontramos tu tarjeta');
  return { ok: true, cliente: rowToCliente_(f.data), cupones: cuponesDe_(ss, f.data[0]) };
}

function login_(ss, ident) {
  var f = findCliente_(ss, ident);
  if (!f) throw new Error('No hay una tarjeta registrada con ese dato');
  return clienteResp_(ss, f.data[0]);
}

/* ---------------- Escritura ---------------- */

function registrar_(ss, p) {
  var nombre = String(p.nombre || '').trim();
  var tel = normTel_(p.telefono);
  var correo = String(p.correo || '').trim().toLowerCase();
  if (nombre.length < 2) throw new Error('Ingresa tu nombre');
  if (!/^9\d{8}$/.test(tel)) throw new Error('El teléfono debe tener 9 dígitos y empezar con 9');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) throw new Error('Correo no válido');
  if (!p.acepta) throw new Error('Debes aceptar el uso de tus datos para crear la tarjeta');
  if (findCliente_(ss, tel) || findCliente_(ss, correo)) throw new Error('Ya existe una tarjeta con ese teléfono o correo. Usa "Iniciar sesión".');
  var id;
  // Código corto y fácil de dictar en caja (sin 0/O/1/I/L)
  var abc = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789';
  do {
    id = '';
    for (var k = 0; k < 4; k++) id += abc.charAt(Math.floor(Math.random() * abc.length));
  } while (findCliente_(ss, id));
  var shC = ss.getSheetByName('Clientes');
  if (!shC.getRange(1, 10).getValue()) shC.getRange(1, 10).setValue('Acepta promociones').setFontWeight('bold').setBackground('#F72DA4').setFontColor('#fff');
  shC.appendRow([id, nombre, "'" + tel, correo, new Date(), 0, '', 0, 0, p.promos ? 'SI' : 'NO']);
  return clienteResp_(ss, id);
}

function sumarVisita_(ss, id, forzar) {
  var cfg = readConfig_(ss);
  var f = findCliente_(ss, id);
  if (!f) throw new Error('Cliente no encontrado');
  var c = rowToCliente_(f.data);
  var hoy = new Date();
  if (String(cfg.una_visita_por_dia).toUpperCase() !== 'NO' && !forzar && f.data[6] && sameDay_(new Date(f.data[6]), hoy)) {
    return { ok: false, error: 'Este cliente ya sumó una visita hoy', yaHoy: true };
  }
  var visitas = c.visitas >= cfg.visitas_meta ? 1 : c.visitas + 1; // tarjeta llena → empieza otra
  var completadas = c.completadas + (visitas === cfg.visitas_meta ? 1 : 0);
  var sh = ss.getSheetByName('Clientes');
  sh.getRange(f.row, 6, 1, 4).setValues([[visitas, hoy, completadas, c.totales + 1]]);
  ss.getSheetByName('Visitas').appendRow([hoy, c.id, c.nombre, visitas, forzar ? 'forzada' : '']);

  var nuevo = null;
  cfg.premios.forEach(function (pr) {
    if (pr.visita === visitas) {
      var vence = new Date(hoy.getTime() + pr.vigencia * 86400000);
      var codigo = 'P' + Math.random().toString(36).slice(2, 7).toUpperCase();
      ss.getSheetByName('Cupones').appendRow([codigo, c.id, c.nombre, pr.emoji, pr.premio, visitas, hoy, vence, 'disponible', '']);
      nuevo = { codigo: codigo, emoji: pr.emoji, premio: pr.premio };
    }
  });
  var r = clienteResp_(ss, c.id);
  r.premioNuevo = nuevo;
  return r;
}

function canjear_(ss, codigo) {
  codigo = String(codigo || '').trim().toUpperCase();
  var sh = ss.getSheetByName('Cupones');
  var rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).toUpperCase() === codigo) {
      if (rows[i][8] === 'usado') throw new Error('Este premio ya fue canjeado');
      if (rows[i][7] && new Date(rows[i][7]) < new Date()) throw new Error('Este premio está vencido');
      sh.getRange(i + 1, 9, 1, 2).setValues([['usado', new Date()]]);
      var r = clienteResp_(ss, rows[i][1]);
      r.canjeado = { codigo: codigo, emoji: String(rows[i][3]), premio: String(rows[i][4]) };
      return r;
    }
  }
  throw new Error('Código no encontrado');
}

/* ---------------- Utilidades ---------------- */

function normTel_(t) {
  var d = String(t || '').replace(/\D/g, '');
  if (d.length === 11 && d.indexOf('56') === 0) d = d.slice(2);
  return d.length === 9 ? d : '';
}
function sameDay_(a, b) {
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(a, tz, 'yyyyMMdd') === Utilities.formatDate(b, tz, 'yyyyMMdd');
}
function fmtDate_(d) {
  if (!(d instanceof Date)) d = new Date(d);
  if (isNaN(d)) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy');
}

/* ---------------- Menú del Sheet central ---------------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Club DMD')
    .addItem('Preparar Sheet central', 'prepararCentral')
    .addItem('Crear nuevo local…', 'crearLocal')
    .addSeparator()
    .addItem('Enviar informe del mes pasado (prueba)', 'informePrueba')
    .addItem('Activar informe mensual automático', 'activarInformeMensual')
    .addToUi();
}

function prepararCentral() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Locales') || ss.insertSheet('Locales');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['slug', 'sheetId', 'pin_caja', 'activo', 'nombre', 'link tarjeta', 'link caja', 'correo informe']);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#F72DA4').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  try { SpreadsheetApp.getUi().alert('Listo. Ahora usa "Club DMD > Crear nuevo local…"'); } catch (e) { Logger.log('Sheet central listo'); }
}

function crearLocal() {
  var ui = SpreadsheetApp.getUi();
  var n = ui.prompt('Nombre del local', 'Ej: Mundo RDL', ui.ButtonSet.OK_CANCEL);
  if (n.getSelectedButton() !== ui.Button.OK) return;
  var nombre = n.getResponseText().trim();
  var slugDef = nombre.toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  var s = ui.prompt('Slug (link)', 'Se usará en club.danymarketing.cl/<slug>. Deja vacío para usar: ' + slugDef, ui.ButtonSet.OK_CANCEL);
  if (s.getSelectedButton() !== ui.Button.OK) return;
  var slug = s.getResponseText().trim() || slugDef;
  var pin = String(Math.floor(1000 + Math.random() * 9000));

  var nuevo = SpreadsheetApp.create('Club ' + nombre);
  inicializarLocal_(nuevo, nombre);
  // Mover a la misma carpeta del Sheet central
  var central = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  var parents = central.getParents();
  if (parents.hasNext()) DriveApp.getFileById(nuevo.getId()).moveTo(parents.next());

  var base = 'https://club.danymarketing.cl/';
  SpreadsheetApp.getActive().getSheetByName('Locales').appendRow([
    slug, nuevo.getId(), "'" + pin, 'SI', nombre, base + '?l=' + slug, base + 'caja.html?l=' + slug
  ]);
  ui.alert('Local creado ✅\n\nSheet: ' + nuevo.getUrl() + '\nPIN de caja: ' + pin + '\nTarjeta: ' + base + '?l=' + slug);
}

function inicializarLocal_(ss, nombre) {
  Object.keys(TABS).forEach(function (t, i) {
    var sh = i === 0 ? ss.getSheets()[0].setName(t) : ss.insertSheet(t);
    sh.appendRow(TABS[t]);
    sh.getRange(1, 1, 1, TABS[t].length).setFontWeight('bold').setBackground('#F72DA4').setFontColor('#fff');
    sh.setFrozenRows(1);
  });
  var cfg = CONFIG_DEFAULT.map(function (r) { return r.slice(); });
  cfg[0][1] = nombre;
  ss.getSheetByName('Config').getRange(2, 1, cfg.length, 3).setValues(cfg);
  ss.getSheetByName('Premios').getRange(2, 1, PREMIOS_DEFAULT.length, 4).setValues(PREMIOS_DEFAULT);
  ss.getSheetByName('Novedades').appendRow(['¡Bienvenido al club! Suma visitas y gana premios.', '', '']);
}

/* ---------------- Informe mensual ----------------
 * El día 1 de cada mes genera, por cada local activo, un Excel con los
 * clientes y la actividad del mes anterior, lo guarda en Drive (carpeta
 * "Informes Club DMD") y lo envía por correo a Dany Marketing Digital.
 * Si la columna "correo informe" de Locales tiene un correo, también se
 * envía a ese correo (el cliente), con copia a Dany.
 */

function activarInformeMensual() {
  PropertiesService.getScriptProperties().setProperty('CENTRAL_ID', SpreadsheetApp.getActive().getId());
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'informeMensual') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('informeMensual').timeBased().onMonthDay(1).atHour(9).create();
  aviso_('Informe mensual activado ✅\nSe enviará el día 1 de cada mes, cerca de las 9:00.');
}

function informePrueba() {
  PropertiesService.getScriptProperties().setProperty('CENTRAL_ID', SpreadsheetApp.getActive().getId());
  var n = informeMensual(true);
  aviso_('Listo ✅ Se enviaron ' + n + ' informe(s) a tu correo (modo prueba: solo a ti).');
}

function informeMensual(soloAMi) {
  if (soloAMi !== true) soloAMi = false; // el activador pasa un objeto de evento
  var hoy = new Date();
  var desde = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 1);
  var hasta = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  var meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  var mesTxt = meses[desde.getMonth()] + ' ' + desde.getFullYear();
  var yo = CORREO_INFORMES || Session.getEffectiveUser().getEmail();
  var central = SpreadsheetApp.getActive() || SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('CENTRAL_ID'));
  var rows = central.getSheetByName('Locales').getDataRange().getValues();
  var enviados = 0;
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0] || String(r[3]).toUpperCase() === 'NO') continue;
    var nombre = String(r[4] || r[0]);
    var dest = soloAMi ? '' : String(r[7] || '').trim();
    try {
      var info = construirInforme_(SpreadsheetApp.openById(r[1]), nombre, desde, hasta, mesTxt);
      var correo = {
        to: dest || yo,
        name: 'Dany Marketing Digital',
        subject: 'Informe Club ' + nombre + ' · ' + mesTxt,
        htmlBody: htmlInforme_(nombre, mesTxt, info.resumen),
        attachments: [info.xlsx]
      };
      if (dest) correo.cc = yo;
      MailApp.sendEmail(correo);
      enviados++;
    } catch (e) {
      Logger.log('ERROR ' + nombre + ': ' + (e && e.stack || e));
      if (soloAMi) throw e;
      MailApp.sendEmail(yo, 'Error informe Club ' + nombre, String(e && e.stack || e));
    }
  }
  return enviados;
}

function construirInforme_(ss, nombre, desde, hasta, mesTxt) {
  var tz = Session.getScriptTimeZone();
  var f = function (d) { return d ? Utilities.formatDate(new Date(d), tz, 'dd-MM-yyyy') : ''; };
  var enMes = function (d) { d = new Date(d); return d >= desde && d < hasta; };

  var cli = ss.getSheetByName('Clientes').getDataRange().getValues().slice(1).filter(function (c) { return c[0]; });
  var vis = ss.getSheetByName('Visitas').getDataRange().getValues().slice(1).filter(function (v) { return v[0] && enMes(v[0]); });
  var cup = ss.getSheetByName('Cupones').getDataRange().getValues().slice(1).filter(function (c) { return c[0]; });

  var visPorCli = {};
  vis.forEach(function (v) { visPorCli[v[1]] = (visPorCli[v[1]] || 0) + 1; });
  var nuevos = cli.filter(function (c) { return c[4] && enMes(c[4]); });
  var ganados = cup.filter(function (c) { return c[6] && enMes(c[6]); });
  var canjeados = cup.filter(function (c) { return c[8] === 'usado' && c[9] && enMes(c[9]); });
  var promos = cli.filter(function (c) { return String(c[9]).toUpperCase() === 'SI'; }).length;

  var resumen = [
    ['Período', mesTxt],
    ['Clientes nuevos en el mes', nuevos.length],
    ['Total de clientes registrados', cli.length],
    ['Visitas registradas en el mes', vis.length],
    ['Clientes que vinieron en el mes', Object.keys(visPorCli).length],
    ['Premios ganados en el mes', ganados.length],
    ['Premios canjeados en el mes', canjeados.length],
    ['Clientes que aceptan promociones', promos]
  ];

  var premiosMes = ganados.concat(canjeados.filter(function (c) { return ganados.indexOf(c) < 0; }));
  var rep = SpreadsheetApp.create('Informe Club ' + nombre + ' - ' + mesTxt);
  var hojas = [
    ['Resumen', ['Indicador', 'Valor'], resumen],
    ['Clientes nuevos', ['Nombre', 'Teléfono', 'Correo', 'Fecha registro', 'Acepta promociones'],
      nuevos.map(function (c) { return [c[1], String(c[2]), c[3], f(c[4]), c[9] || 'NO']; })],
    ['Todos los clientes', ['Nombre', 'Teléfono', 'Correo', 'Fecha registro', 'Visitas en el mes', 'Visitas en tarjeta actual', 'Visitas totales', 'Última visita', 'Tarjetas completadas', 'Acepta promociones'],
      cli.map(function (c) { return [c[1], String(c[2]), c[3], f(c[4]), visPorCli[c[0]] || 0, c[5], c[8], f(c[6]), c[7], c[9] || 'NO']; })
        .sort(function (a, b) { return b[4] - a[4] || b[6] - a[6]; })],
    ['Visitas del mes', ['Fecha', 'Nombre', 'Visita N°'],
      vis.map(function (v) { return [Utilities.formatDate(new Date(v[0]), tz, 'dd-MM-yyyy HH:mm'), v[2], v[3]]; })],
    ['Premios del mes', ['Cliente', 'Premio', 'Ganado el', 'Estado', 'Canjeado el'],
      premiosMes.map(function (c) { return [c[2], c[3] + ' ' + c[4], f(c[6]), c[8], f(c[9])]; })]
  ];
  hojas.forEach(function (h, i) {
    var sh = i === 0 ? rep.getSheets()[0].setName(h[0]) : rep.insertSheet(h[0]);
    var vacio = h[1].map(function (_, k) { return k === 0 ? 'Sin datos este mes' : ''; });
    var data = [h[1]].concat(h[2].length ? h[2] : [vacio]);
    sh.getRange(1, 1, data.length, h[1].length).setNumberFormat('@').setValues(data);
    sh.getRange(1, 1, 1, h[1].length).setFontWeight('bold').setBackground('#FF8A2A').setFontColor('#fff');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, h[1].length);
  });
  SpreadsheetApp.flush();

  // Guardar en Drive, carpeta "Informes Club DMD"
  var it = DriveApp.getFoldersByName('Informes Club DMD');
  var carpeta = it.hasNext() ? it.next() : DriveApp.createFolder('Informes Club DMD');
  DriveApp.getFileById(rep.getId()).moveTo(carpeta);

  var xlsx = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + rep.getId() + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }
  }).getBlob().setName('Informe Club ' + nombre + ' - ' + mesTxt + '.xlsx');

  return { resumen: resumen, xlsx: xlsx };
}

function htmlInforme_(nombre, mesTxt, resumen) {
  var filas = resumen.slice(1).map(function (r) {
    return '<tr><td style="padding:8px 12px;border-bottom:1px solid #FFE2C2">' + r[0] +
      '</td><td style="padding:8px 12px;border-bottom:1px solid #FFE2C2;font-weight:700;text-align:right">' + r[1] + '</td></tr>';
  }).join('');
  return '<div style="font-family:Arial,sans-serif;max-width:560px;color:#20263A">' +
    '<div style="background:#FF8A2A;color:#fff;padding:18px 20px;border-radius:14px 14px 0 0;border-bottom:4px solid #35C6E8">' +
    '<div style="font-size:13px">Informe mensual · Tarjeta de visitas</div>' +
    '<div style="font-size:22px;font-weight:700">Club ' + nombre + ' · ' + mesTxt + '</div></div>' +
    '<div style="background:#fff;border:1px solid #FFE2C2;border-top:0;border-radius:0 0 14px 14px;padding:16px 8px">' +
    '<table style="width:100%;border-collapse:collapse;font-size:14px">' + filas + '</table>' +
    '<p style="font-size:14px;padding:0 12px">Adjuntamos el Excel con el detalle de clientes nuevos, todos los clientes (nombre, teléfono, correo y visitas), las visitas y los premios del mes.</p>' +
    '<p style="font-size:12px;color:#5d6377;padding:0 12px">Para enviar promociones, contacta solo a los clientes con "Acepta promociones = SI".</p>' +
    '</div><p style="font-size:12px;color:#5d6377;text-align:center">Creado por <b style="color:#F72DA4">Dany Marketing Digital</b></p></div>';
}

function aviso_(txt) {
  try { SpreadsheetApp.getUi().alert(txt); } catch (e) { Logger.log(txt); }
}
