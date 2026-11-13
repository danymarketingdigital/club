/* CLUB DMD — conexión con el backend (Google Apps Script)
 * Pega aquí la URL de la implementación web del Apps Script (termina en /exec).
 * Si queda vacía, o si la página se abre con ?demo=1, funciona en MODO DEMO
 * (datos guardados solo en este navegador, para mostrar y probar). */
const API_URL = 'https://script.google.com/macros/s/AKfycbwGGRgp1m2cIx5tcqqef_WRiCQW8JtwUbLBQdxx8VX9WeU7CfzdFs2rumq7syFPX0aO/exec';

const PARAMS = new URLSearchParams(location.search);
const LOCAL = (PARAMS.get('l') || 'demo').toLowerCase();
const DEMO = !API_URL || PARAMS.get('demo') === '1';

async function api(action, data = {}) {
  if (DEMO) return demoApi(action, data);
  const res = await fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, local: LOCAL, ...data })
  });
  return res.json();
}

/* ---------- Modo demo (misma lógica que Code.gs, en localStorage) ---------- */
const DEMO_CFG = {
  nombre: 'Mundo RDL', club: 'Club RDL', logo_url: '', color: '#F72DA4', color2: '#35C6E8',
  visitas_meta: 10,
  descripcion: 'Suma visitas en tu tarjeta virtual y gana premios en la visita 3, 6, 8 y 10.',
  instagram: '',whatsapp: '', una_visita_por_dia: 'SI',
  premios: [
    { visita: 3, emoji: '🥤', premio: 'Jugo o café de regalo', vigencia: 30 },
    { visita: 6, emoji: '🎁', premio: 'Regalo sorpresa', vigencia: 30 },
    { visita: 8, emoji: '🛝', premio: 'Zona de juegos gratis', vigencia: 30 },
    { visita: 10, emoji: '⭐', premio: 'Regalo gratis', vigencia: 30 }
  ],
  novedades: [{ texto: '¡Bienvenido al Club RDL! Muestra tu teléfono en caja en cada visita.', imagen: '', link: '' }]
};
const DEMO_PIN = '1234';

function demoDb() {
  try { return JSON.parse(localStorage.getItem('clubdmd_demo') || '{"clientes":[],"cupones":[]}'); }
  catch (e) { return { clientes: [], cupones: [] }; }
}
function demoSave(db) { try { localStorage.setItem('clubdmd_demo', JSON.stringify(db)); } catch (e) {} }
const fmt = d => { d = new Date(d); return [d.getDate(), d.getMonth() + 1, d.getFullYear()].map(n => String(n).padStart(2, '0')).join('-'); };
const normTel = t => { let d = String(t || '').replace(/\D/g, ''); if (d.length === 11 && d.startsWith('56')) d = d.slice(2); return d.length === 9 ? d : ''; };

function demoFind(db, q) {
  q = String(q || '').trim().toLowerCase(); if (!q) return null;
  const tel = normTel(q);
  return db.clientes.find(c => c.id.toLowerCase() === q || (tel && c.telefono === tel) || c.correo === q) || null;
}
function demoResp(db, c) {
  const hoy = new Date();
  const cupones = db.cupones.filter(x => x.cliente === c.id).map(x => ({
    ...x, estado: x.estado === 'disponible' && new Date(x.venceISO) < hoy ? 'vencido' : x.estado
  })).reverse();
  return { ok: true, cliente: { ...c, ultima: c.ultimaISO ? fmt(c.ultimaISO) : '' }, cupones };
}

async function demoApi(action, p) {
  await new Promise(r => setTimeout(r, 250));
  const db = demoDb();
  const cfg = DEMO_CFG;
  const needPin = () => { if (String(p.pin) !== DEMO_PIN) throw new Error('PIN incorrecto (en la demo es 1234)'); };
  try {
    switch (action) {
      case 'config': return { ok: true, config: cfg };
      case 'registrar': {
        const nombre = String(p.nombre || '').trim(), tel = normTel(p.telefono), correo = String(p.correo || '').trim().toLowerCase();
        if (nombre.length < 2) throw new Error('Ingresa tu nombre');
        if (!/^9\d{8}$/.test(tel)) throw new Error('El teléfono debe tener 9 dígitos y empezar con 9');
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(correo)) throw new Error('Correo no válido');
        if (demoFind(db, tel) || demoFind(db, correo)) throw new Error('Ya existe una tarjeta con ese teléfono o correo. Usa "Iniciar sesión".');
        const abc = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; let id;
        do { id = Array.from({ length: 4 }, () => abc[Math.floor(Math.random() * abc.length)]).join(''); } while (demoFind(db, id));
        const c = { id, nombre, telefono: tel, correo, visitas: 0, ultimaISO: '', completadas: 0, totales: 0 };
        db.clientes.push(c); demoSave(db); return demoResp(db, c);
      }
      case 'login': case 'cliente': {
        const c = demoFind(db, p.ident || p.id);
        if (!c) throw new Error(action === 'login' ? 'No hay una tarjeta registrada con ese dato' : 'No encontramos tu tarjeta');
        return demoResp(db, c);
      }
      case 'caja_login': needPin(); return { ok: true, config: cfg };
      case 'caja_buscar': {
        needPin(); const c = demoFind(db, p.q);
        if (!c) throw new Error('No hay una tarjeta registrada con ese dato');
        return demoResp(db, c);
      }
      case 'caja_visita': {
        needPin(); const c = demoFind(db, p.id); if (!c) throw new Error('Cliente no encontrado');
        const hoy = new Date();
        if (!p.forzar && c.ultimaISO && fmt(c.ultimaISO) === fmt(hoy)) return { ok: false, error: 'Este cliente ya sumó una visita hoy', yaHoy: true };
        c.visitas = c.visitas >= cfg.visitas_meta ? 1 : c.visitas + 1;
        if (c.visitas === cfg.visitas_meta) c.completadas++;
        c.totales++; c.ultimaISO = hoy.toISOString();
        let nuevo = null;
        cfg.premios.filter(pr => pr.visita === c.visitas).forEach(pr => {
          const codigo = 'P' + Math.random().toString(36).slice(2, 7).toUpperCase();
          const vence = new Date(hoy.getTime() + pr.vigencia * 864e5);
          db.cupones.push({ codigo, cliente: c.id, emoji: pr.emoji, premio: pr.premio, visita: c.visitas, creado: fmt(hoy), vence: fmt(vence), venceISO: vence.toISOString(), estado: 'disponible', usado: '' });
          nuevo = { codigo, emoji: pr.emoji, premio: pr.premio };
        });
        demoSave(db); const r = demoResp(db, c); r.premioNuevo = nuevo; return r;
      }
      case 'caja_canjear': {
        needPin(); const cod = String(p.codigo || '').trim().toUpperCase();
        const x = db.cupones.find(k => k.codigo === cod); if (!x) throw new Error('Código no encontrado');
        if (x.estado === 'usado') throw new Error('Este premio ya fue canjeado');
        if (new Date(x.venceISO) < new Date()) throw new Error('Este premio está vencido');
        x.estado = 'usado'; x.usado = fmt(new Date()); demoSave(db);
        const r = demoResp(db, demoFind(db, x.cliente)); r.canjeado = { codigo: cod, emoji: x.emoji, premio: x.premio }; return r;
      }
      default: throw new Error('Acción no válida');
    }
  } catch (e) { return { ok: false, error: e.message }; }
}

/* ---------- Utilidades de interfaz compartidas ---------- */
function applyTheme(cfg) {
  const r = document.documentElement.style;
  if (cfg.color) r.setProperty('--brand', cfg.color);
  if (cfg.color2) r.setProperty('--brand2', cfg.color2);
  document.title = (cfg.club || 'Club') + ' · ' + (cfg.nombre || '');
}
function esc(s) { return String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function logoHtml(cfg) {
  return cfg.logo_url
    ? `<img class="logo" src="${esc(cfg.logo_url)}" alt="${esc(cfg.nombre)}">`
    : `<div class="logo">${esc((cfg.nombre || 'C').split(/\s+/).map(w => w[0]).join('').slice(0, 3).toUpperCase())}</div>`;
}
