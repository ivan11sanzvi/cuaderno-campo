'use strict';
(function () {
  // ---------- Constantes ----------
  const KEY = 'cuadernoCampo.v1';
  const VERSION = '1.9.1';
  const TIPOS_LABOR = ['Labrado', 'Siembra', 'Abonado', 'Tratamiento', 'Rodillo', 'Cosecha', 'Empacado', 'Otro'];
  const CON_DOSIS = ['Siembra', 'Abonado', 'Tratamiento'];
  const UNIDADES = ['kg/ha', 'L/ha', 'g/ha', 'cc/ha', 'kg', 'L'];
  const CAT = {
    gasto: ['Semilla', 'Abono', 'Fitosanitarios', 'Gasóleo', 'Maquinaria y reparaciones', 'Servicios contratados', 'Seguro', 'Arrendamiento', 'Otros gastos'],
    ingreso: ['Venta de grano', 'Venta de paja', 'Subvención PAC', 'Otros ingresos']
  };
  const CAT_LABOR = { Siembra: 'Semilla', Abonado: 'Abono', Tratamiento: 'Fitosanitarios' };
  const LABOREO = { C: 'Siembra directa con complemento', T: 'Tradicional', D: 'Siembra directa' };
  const CULT_ALIAS = { lentejas: 'Lenteja', anis: 'Anís dulce', 'anis dulce': 'Anís dulce', matalahuga: 'Anís dulce', chopos: 'Chopo', 'barbecho tradicional': 'Barbecho', 'trigo blando': 'Trigo' };
  const SIN_CULTIVO = ['no se solicita', 'nos la quitan'];
  const CULTIVOS_BASE = ['Trigo', 'Cebada', 'Titarro', 'Lenteja', 'Anís dulce', 'Cilantro', 'Esparceta', 'Barbecho'];
  const HUES = { Trigo: 40, Cebada: 55, Titarro: 150, Lenteja: 18, 'Anís dulce': 85, Cilantro: 120, Esparceta: 320, Barbecho: 30 };
  const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const MESES_L = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const TITULOS = { fincas: 'Fincas', cuaderno: 'Cuaderno de campo', cuentas: 'Cuentas', cosechas: 'Cosechas', almacen: 'Almacén', explotacion: 'Explotación', resumen: 'Resumen' };
  // Servicios públicos: SIGPAC (FEGA) y ortofoto PNOA (IGN).
  const SIGPAC = 'https://sigpac-hubcloud.es';
  const PNOA = 'https://tms-pnoa-ma.idee.es/1.0.0/pnoa-ma/{z}/{x}/{y}.jpeg';
  const XLSX_SRC = window.XLSX_SRC || 'vendor/xlsx.full.min.js';
  const USOS = { TA: 'Tierra arable', TH: 'Huerta', PS: 'Pastizal', PR: 'Pasto arbustivo', PA: 'Pasto con arbolado', FO: 'Forestal', OV: 'Olivar', VI: 'Viñedo', FY: 'Frutales', FS: 'Frutos secos', FL: 'Frutos secos y olivar', CI: 'Cítricos', IM: 'Improductivo', CA: 'Viales', ED: 'Edificaciones', ZU: 'Zona urbana', AG: 'Agua', IV: 'Invernadero', ZC: 'Zona concentrada', ZV: 'Zona censurada' };
  const PROVINCIAS = ['', 'Álava', 'Albacete', 'Alicante', 'Almería', 'Ávila', 'Badajoz', 'Baleares', 'Barcelona', 'Burgos', 'Cáceres', 'Cádiz', 'Castellón', 'Ciudad Real', 'Córdoba', 'A Coruña', 'Cuenca', 'Girona', 'Granada', 'Guadalajara', 'Gipuzkoa', 'Huelva', 'Huesca', 'Jaén', 'León', 'Lleida', 'La Rioja', 'Lugo', 'Madrid', 'Málaga', 'Murcia', 'Navarra', 'Ourense', 'Asturias', 'Palencia', 'Las Palmas', 'Pontevedra', 'Salamanca', 'Santa Cruz de Tenerife', 'Cantabria', 'Segovia', 'Sevilla', 'Soria', 'Tarragona', 'Teruel', 'Toledo', 'Valencia', 'Valladolid', 'Bizkaia', 'Zamora', 'Zaragoza', 'Ceuta', 'Melilla'];
  const PROV_ALIAS = { araba: 1, alacant: 3, 'illes balears': 7, 'islas baleares': 7, castello: 12, 'la coruna': 15, coruna: 15, gerona: 17, guipuzcoa: 20, lerida: 25, orense: 32, vizcaya: 48, rioja: 26, navarra: 31, 'valencia valencia': 46 };
  // Recintos reales de ejemplo (datos públicos SIGPAC, simplificados).
  const EJ_GEO = {"47:215:0:0:403:30:1":{"ha":15.8929,"c":[[[[-5.052177,42.078819],[-5.052217,42.07888],[-5.052783,42.080031],[-5.052848,42.080194],[-5.05288,42.0803],[-5.052918,42.080638],[-5.052963,42.081228],[-5.052991,42.081476],[-5.047682,42.082866],[-5.04761,42.082653],[-5.047246,42.081465],[-5.047224,42.081414],[-5.047168,42.081335],[-5.046598,42.080636],[-5.045823,42.079731],[-5.045817,42.079716],[-5.047099,42.079798],[-5.047188,42.079801],[-5.047332,42.079795],[-5.050176,42.079207],[-5.052177,42.078819]]]]},"47:215:0:0:403:69:1":{"ha":13.0009,"c":[[[[-5.053132,42.074777],[-5.056275,42.074923],[-5.056226,42.075195],[-5.056221,42.075321],[-5.05623,42.075435],[-5.056249,42.075548],[-5.056301,42.075718],[-5.056572,42.076429],[-5.048802,42.077866],[-5.049292,42.075284],[-5.053029,42.075723],[-5.053045,42.07564],[-5.053132,42.074777]]]]},"47:215:0:0:403:36:1":{"ha":13.2826,"c":[[[[-5.049995,42.082359],[-5.051586,42.085451],[-5.047842,42.086482],[-5.047402,42.086353],[-5.045896,42.083426],[-5.049995,42.082359]]]]},"47:215:0:0:404:39:1":{"ha":5.0953,"c":[[[[-5.049141,42.09618],[-5.050061,42.09809],[-5.047217,42.097892],[-5.04508,42.097718],[-5.044889,42.097327],[-5.047057,42.096749],[-5.049141,42.09618]]]]},"47:215:0:0:403:8:1":{"ha":3.1385,"c":[[[[-5.04734,42.089678],[-5.048955,42.092867],[-5.04802,42.093132],[-5.046401,42.089936],[-5.04734,42.089678]]]]}};

  // ---------- Utilidades ----------
  const $ = (s, el) => (el || document).querySelector(s);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const norm = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const hoyISO = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 60000).toISOString().slice(0, 10); };
  const campanaDe = iso => { const [y, m] = iso.split('-').map(Number); const a = m >= 9 ? y : y - 1; return a + '/' + String((a + 1) % 100).padStart(2, '0'); };
  const siguienteCampana = c => { const a = parseInt(c, 10) + 1; return a + '/' + String((a + 1) % 100).padStart(2, '0'); };
  const hue = c => { if (c in HUES) return HUES[c]; let h = 0; for (const ch of String(c)) h = (h * 31 + ch.charCodeAt(0)) % 360; return h; };
  const fmtN = (v, d = 2) => (v == null || isNaN(v)) ? '—' : new Intl.NumberFormat('es-ES', { maximumFractionDigits: d, useGrouping: 'always' }).format(v);
  const eur = (v, d = 0) => (v == null || isNaN(v)) ? '—' : new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'EUR', minimumFractionDigits: d, maximumFractionDigits: d, useGrouping: 'always' }).format(v);
  const fmtKg = v => (v == null || isNaN(v)) ? '—' : fmtN(v, 2) + ' kg'; // siempre en kilos exactos
  const fmtFecha = iso => { if (!iso) return ''; const [y, m, d] = iso.split('-'); return `${+d} ${MESES[+m - 1]} ${y}`; };
  const numStr = v => (v == null || v === '' || isNaN(v)) ? '' : String(v).replace('.', ',');
  const sum = (arr, f) => arr.reduce((a, x) => a + (+f(x) || 0), 0);
  const byNombre = (a, b) => a.nombre.localeCompare(b.nombre, 'es', { numeric: true });
  const byFechaDesc = (a, b) => (b.fecha || '').localeCompare(a.fecha || '') || (b.id < a.id ? -1 : 1);

  // Acepta "1.250,5", "1250.5", "1250,5" y, en campos de kilos/euros, "1.250" como mil doscientos cincuenta.
  function parseNum(v, miles) {
    if (v == null) return null;
    let s = String(v).trim().replace(/\s|€/g, '');
    if (!s) return null;
    if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.');
    else if (miles && /^[1-9]\d{0,2}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
    const n = Number(s);
    return isFinite(n) ? n : NaN;
  }

  // ---------- Datos ----------
  function nuevoDB() {
    const c = campanaDe(hoyISO());
    return { v: 1, campanaActiva: c, campanas: [c], cultivos: CULTIVOS_BASE.slice(), fincas: [], siembras: [], labores: [], movimientos: [], cosechas: [], productos: [], documentos: [], ajustes: [], titulares: [], maquinas: [], equipos: [], ultimaCopia: null, ultimoCambio: null, ejemplo: false };
  }
  function normalizar(d) {
    const base = nuevoDB();
    const o = Object.assign(base, d && typeof d === 'object' ? d : {});
    for (const k of ['campanas', 'fincas', 'siembras', 'labores', 'movimientos', 'cosechas', 'productos', 'documentos', 'ajustes', 'titulares', 'maquinas', 'equipos']) if (!Array.isArray(o[k])) o[k] = [];
    if (!Array.isArray(o.cultivos) || !o.cultivos.length) o.cultivos = CULTIVOS_BASE.slice();
    if (!o.campanaActiva) o.campanaActiva = campanaDe(hoyISO());
    if (!o.campanas.includes(o.campanaActiva)) o.campanas.push(o.campanaActiva);
    return o;
  }
  function cargar() {
    try { const raw = localStorage.getItem(KEY); if (raw) return normalizar(JSON.parse(raw)); } catch (e) { /* sin almacenamiento */ }
    return nuevoDB();
  }
  let db = cargar();
  function guardar(cambio = true) {
    if (cambio) db.ultimoCambio = new Date().toISOString();
    try { localStorage.setItem(KEY, JSON.stringify(db)); }
    catch (e) { toast('No se ha podido guardar en el teléfono. Haz una copia de seguridad.'); }
  }
  function listaCampanas() {
    const s = new Set(db.campanas);
    for (const k of ['siembras', 'labores', 'movimientos', 'cosechas']) db[k].forEach(x => x.campana && s.add(x.campana));
    s.add(db.campanaActiva);
    return [...s].sort().reverse();
  }

  const finca = id => db.fincas.find(f => f.id === id);
  const siembraDe = (fid, c) => db.siembras.find(s => s.fincaId === fid && s.campana === c) || null;
  const cultivoDe = (fid, c) => { const s = siembraDe(fid, c); return s ? s.cultivo : null; };

  // ---------- Fincas agrupadas (varias parcelas que tú llamas igual) ----------
  const grupoDe = f => String((f && f.grupo) || '').trim();
  const fincasDeGrupo = g => db.fincas.filter(f => grupoDe(f) === g);
  function listaGrupos() {
    const m = new Map();
    for (const f of db.fincas) { const g = grupoDe(f); if (g) { if (!m.has(g)) m.set(g, []); m.get(g).push(f); } }
    return m;
  }
  const titulares = () => [...new Set(db.fincas.map(f => String(f.titular || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
  const etiquetaLaboreo = l => LABOREO[l] || l || '';
  const nombreBase = n => String(n || '').replace(/\s+\d+\s*-\s*\d+.*$/, '').trim();

  // Parcelas que lindan (a menos de ~5 m), con los contornos de SIGPAC
  const bbCache = new WeakMap();
  function bbox(geo) {
    if (bbCache.has(geo)) return bbCache.get(geo);
    let x0 = 999, y0 = 999, x1 = -999, y1 = -999;
    geo.c.forEach(p => p.forEach(r => r.forEach(([x, y]) => { if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y; })));
    const b = [x0, y0, x1, y1]; bbCache.set(geo, b); return b;
  }
  function distPS(p, a, b, kx, ky) {
    const px = p[0] * kx, py = p[1] * ky, ax = a[0] * kx, ay = a[1] * ky, dx = b[0] * kx - ax, dy = b[1] * ky - ay;
    const L = dx * dx + dy * dy; let t = L ? ((px - ax) * dx + (py - ay) * dy) / L : 0; t = Math.max(0, Math.min(1, t));
    return Math.hypot(px - ax - t * dx, py - ay - t * dy);
  }
  function cerca(A, B, tol) {
    const lat = A.c[0][0][0][1]; const kx = 111320 * Math.cos(lat * Math.PI / 180), ky = 110540;
    const segs = B.c.flatMap(p => p.flatMap(r => r.slice(1).map((q, i) => [r[i], q])));
    for (const p of A.c) for (const r of p) for (const v of r) for (const [a, b] of segs) if (distPS(v, a, b, kx, ky) < tol) return true;
    return false;
  }
  function lindan(fa, fb, tol = 5) {
    const A = fa.geo, B = fb.geo; if (!A || !B || !A.c.length || !B.c.length) return false;
    const a = bbox(A), b = bbox(B), m = tol / 80000;
    if (a[2] + m < b[0] || b[2] + m < a[0] || a[3] + m < b[1] || b[3] + m < a[1]) return false;
    return cerca(A, B, tol) || cerca(B, A, tol);
  }
  const lindantes = fs => { const ids = new Set(fs.map(f => f.id)); return db.fincas.filter(o => !ids.has(o.id) && fs.some(f => lindan(f, o))); };

  // Apuntes hechos a la vez en varias parcelas comparten «lote»
  const miembrosLote = (arr, x) => x.lote ? arr.filter(y => y.lote === x.lote) : [x];
  function agruparLote(arr) {
    const vistos = new Set(); const out = [];
    for (const x of arr) {
      if (!x.lote) { out.push({ x, n: [x] }); continue; }
      if (vistos.has(x.lote)) continue; vistos.add(x.lote);
      out.push({ x, n: arr.filter(y => y.lote === x.lote) });
    }
    return out;
  }
  // Reparte un total entre fincas según su superficie (el redondeo va a la última)
  function repartir(total, fs) {
    if (total == null || total === '' || Number.isNaN(total)) return fs.map(() => null);
    const ha = fs.map(f => +f.superficie || 0); const T = ha.reduce((a, b) => a + b, 0);
    let acc = 0;
    return fs.map((f, i) => {
      if (i === fs.length - 1) return Math.round((total - acc) * 100) / 100;
      const v = Math.round(total * (T ? ha[i] / T : 1 / fs.length) * 100) / 100; acc += v; return v;
    });
  }
  function nombreLote(n) {
    const fs = n.map(x => finca(x.fincaId)).filter(Boolean);
    if (fs.length <= 1) return fs[0] ? fs[0].nombre : 'Finca borrada';
    const gs = new Set(fs.map(grupoDe));
    if (gs.size === 1 && [...gs][0]) { const g = [...gs][0]; return fincasDeGrupo(g).length === fs.length ? g : `${g} (${fs.length} parcelas)`; }
    return `${fs.length} parcelas`;
  }

  // ---------- SIGPAC: referencias y contornos ----------
  // "1, 2, 5-7" -> [1,2,5,6,7]
  function listaRecintos(s) {
    const out = [];
    String(s == null ? '' : s).split(/[,;\s]+/).forEach(t => {
      const m = t.match(/^(\d+)(?:-(\d+))?$/); if (!m) return;
      const a = +m[1], b = m[2] ? +m[2] : a;
      for (let i = a; i <= Math.min(b, a + 300); i++) out.push(i);
    });
    return [...new Set(out)];
  }
  // Todas las referencias SIGPAC de una finca: la principal (con sus recintos) y las de «otras parcelas».
  function refsFinca(f) {
    const pr = +f.provincia || 0, mu = +f.codMunicipio || 0, ag = +f.agregado || 0, zo = +f.zona || 0;
    const out = [];
    const add = r => { if (r.pr && r.mu && r.po && r.pa) out.push(r); };
    if (f.poligono && f.parcela) {
      const base = { pr, mu, ag, zo, po: +f.poligono, pa: +f.parcela };
      const recs = listaRecintos(f.recinto);
      if (recs.length) recs.forEach(re => add(Object.assign({ re }, base))); else add(Object.assign({ re: null }, base));
    }
    String(f.otras || '').split(/[\n;]+/).forEach(l => {
      const n = l.trim().split(/[:\-\/\s]+/).filter(Boolean).map(Number);
      if (!n.length || n.some(isNaN)) return;
      if (n.length >= 6) add({ pr: n[0], mu: n[1], ag: n[2], zo: n[3], po: n[4], pa: n[5], re: n[6] || null });
      else if (n.length >= 2) add({ pr, mu, ag, zo, po: n[0], pa: n[1], re: n[2] || null });
    });
    return out;
  }
  const refTxt = r => [r.pr, r.mu, r.ag, r.zo, r.po, r.pa].join(':') + (r.re ? ':' + r.re : '');
  const refsKey = f => refsFinca(f).map(refTxt).join('|');
  const nombreProv = c => PROVINCIAS[+c] || '';
  function provinciaDe(v) {
    if (v == null || v === '') return '';
    if (/^\d+$/.test(String(v).trim())) return +v;
    const n = norm(v);
    const i = PROVINCIAS.findIndex(p => p && norm(p) === n);
    return i > 0 ? i : (PROV_ALIAS[n] || '');
  }
  function fincaConRecinto(x) {
    return db.fincas.find(f => refsFinca(f).some(r => r.pr === x.provincia && r.mu === x.municipio && r.ag === x.agregado && r.zo === x.zona && r.po === x.poligono && r.pa === x.parcela && (!r.re || r.re === x.recinto)));
  }

  function dp(pts, tol) {
    if (pts.length < 3) return pts;
    const a = pts[0], b = pts[pts.length - 1], dx = b[0] - a[0], dy = b[1] - a[1], L = Math.hypot(dx, dy);
    let dmax = 0, idx = 0;
    for (let i = 1; i < pts.length - 1; i++) {
      const p = pts[i];
      const d = L ? Math.abs(dy * p[0] - dx * p[1] + b[0] * a[1] - b[1] * a[0]) / L : Math.hypot(p[0] - a[0], p[1] - a[1]);
      if (d > dmax) { dmax = d; idx = i; }
    }
    return dmax > tol ? dp(pts.slice(0, idx + 1), tol).slice(0, -1).concat(dp(pts.slice(idx), tol)) : [a, b];
  }
  // Redondea a ~10 cm y quita vértices sobrantes (< 0,4 m) para que quepan muchas fincas en el teléfono.
  function simplificarAnillo(r) {
    const pts = r.map(p => [Math.round(p[0] * 1e6) / 1e6, Math.round(p[1] * 1e6) / 1e6]);
    if (pts.length < 8) return pts;
    let m = 0, dm = 0;
    pts.forEach((p, i) => { const d = Math.hypot(p[0] - pts[0][0], p[1] - pts[0][1]); if (d > dm) { dm = d; m = i; } });
    return dp(pts.slice(0, m + 1), 4e-6).slice(0, -1).concat(dp(pts.slice(m), 4e-6));
  }
  function areaAnillo(r) {
    const R = 6378137, d = Math.PI / 180; let a = 0;
    for (let i = 0; i < r.length - 1; i++) { const [x1, y1] = r[i], [x2, y2] = r[i + 1]; a += (x2 - x1) * d * (2 + Math.sin(y1 * d) + Math.sin(y2 * d)); }
    return Math.abs(a * R * R / 2);
  }
  const areaHa = c => c.reduce((s, poly) => s + poly.reduce((t, r, i) => t + (i ? -1 : 1) * areaAnillo(r), 0), 0) / 1e4;
  // WKT (POLYGON / MULTIPOLYGON) -> coordenadas MultiPolygon
  function wktCoords(w) {
    const polys = []; const rd = /MULTIPOLYGON/i.test(w) ? 3 : 2;
    let depth = 0, poly = null, buf = '';
    for (const ch of String(w)) {
      if (ch === '(') { depth++; if (depth === rd - 1) poly = []; if (depth === rd) buf = ''; }
      else if (ch === ')') { if (depth === rd) poly.push(buf.split(',').map(p => p.trim().split(/\s+/).map(Number))); if (depth === rd - 1) polys.push(poly); depth--; }
      else if (depth === rd) buf += ch;
    }
    return polys;
  }

  async function contornoFinca(f) {
    // Una consulta por parcela; si solo son algunos recintos, se filtran.
    const grupos = new Map();
    for (const r of refsFinca(f)) {
      const k = [r.pr, r.mu, r.ag, r.zo, r.po, r.pa].join(':');
      const g = grupos.get(k) || { r, recs: new Set(), todo: false };
      if (r.re) g.recs.add(r.re); else g.todo = true;
      grupos.set(k, g);
    }
    const polys = []; let ha = 0;
    for (const { r, recs, todo } of grupos.values()) {
      const q = new URLSearchParams({ f: 'json', provincia: r.pr, municipio: r.mu, agregado: r.ag, zona: r.zo, poligono: r.po, parcela: r.pa, limit: '1000' });
      if (!todo && recs.size === 1) q.set('recinto', [...recs][0]);
      const res = await fetch(`${SIGPAC}/ogcapi/collections/recintos/items?${q}`);
      if (!res.ok) throw new Error('SIGPAC ' + res.status);
      const d = await res.json();
      for (const ft of d.features || []) {
        if (!todo && !recs.has(ft.properties && ft.properties.recinto)) continue;
        const g = ft.geometry; if (!g) continue;
        const ps = g.type === 'Polygon' ? [g.coordinates] : g.type === 'MultiPolygon' ? g.coordinates : [];
        ps.forEach(p => { const s = p.map(simplificarAnillo); polys.push(s); ha += areaHa([p]); });
      }
    }
    return { k: refsKey(f), c: polys, ha: Math.round(ha * 100) / 100, t: hoyISO() };
  }
  let buscandoContornos = false;
  const fallidos = new Set(); // referencias que fallaron en esta sesión (no se reintentan solas)
  async function actualizarContornos(opts = {}) {
    db.fincas.forEach(f => { if (f.geo && !refsKey(f)) delete f.geo; });
    if (opts.forzar) fallidos.clear();
    const pend = db.fincas.filter(f => { const k = refsKey(f); return k && !fallidos.has(k) && (opts.forzar || !f.geo || f.geo.k !== k); });
    if (!pend.length || buscandoContornos) { if (opts.forzar && !opts.silencioso && !pend.length) toast('Ninguna finca tiene referencia SIGPAC'); return; }
    if (navigator.onLine === false) { if (!opts.silencioso) toast('Sin conexión: los contornos se buscarán cuando haya cobertura'); return; }
    buscandoContornos = true;
    if (!opts.silencioso || pend.length > 2) toast(`Buscando ${pend.length === 1 ? 'la finca' : pend.length + ' fincas'} en SIGPAC…`);
    let ok = 0, no = 0, fallo = 0;
    let i = 0;
    const trabajador = async () => {
      while (i < pend.length) {
        const f = pend[i++];
        try {
          const g = await contornoFinca(f);
          f.geo = g;
          if (g.c.length) { ok++; if (!(f.superficie > 0) && g.ha) f.superficie = g.ha; } else no++;
        } catch (e) { fallo++; fallidos.add(refsKey(f)); }
      }
    };
    await Promise.all([1, 2, 3, 4].map(trabajador));
    buscandoContornos = false;
    guardar(false); render();
    const partes = [];
    if (ok) partes.push(`${ok} ${ok === 1 ? 'finca situada' : 'fincas situadas'} en el mapa`);
    if (no) partes.push(`${no} no ${no === 1 ? 'aparece' : 'aparecen'} en SIGPAC (revisa la referencia)`);
    if (fallo) partes.push(`${fallo} sin respuesta de SIGPAC; se reintentará`);
    if (partes.length && (!opts.silencioso || no || fallo || ok > 2)) toast(partes.join(' · '));
  }

  // Coste de la maquinaria propia en una labor (amortización del equipo × hectáreas). null = no se puede calcular
  const haLabor = l => l.superficie > 0 ? +l.superficie : +((finca(l.fincaId) || {}).superficie) || 0;
  function costeMaqLabor(l) {
    const e = equipo(l.equipoId); if (!e) return null;
    const c = costeEquipo(e); if (c.porHa == null) return null;
    return Math.round(c.porHa * haLabor(l) * 100) / 100;
  }
  // ---------- Gasóleo ----------
  // Litros = L/ha del equipo × hectáreas. El gasóleo es de los tres: se gasta primero lo que entró antes
  // (el albarán más antiguo, sea de quien sea), y cuando se acaba, el siguiente. Cada litro lleva el precio de su albarán.
  const esGasoleo = p => p && p.unidad === 'L' && (p.tipo === 'Combustible' || /gas[oó]le?o|gas-?oil/i.test(p.nombre));
  function gasoleoAlm() {
    const ps = db.productos.filter(esGasoleo); if (ps.length <= 1) return ps[0] || null;
    const n = p => db.documentos.reduce((a, d) => a + (d.lineas || []).filter(x => x.productoId === p.id).length, 0);
    return ps.sort((a, b) => n(b) - n(a))[0];
  }
  const litrosLabor = l => { const e = equipo(l.equipoId); return e && +e.gasoleo > 0 ? Math.round(+e.gasoleo * haLabor(l) * 100) / 100 : 0; };
  let cacheGas = { k: '', r: null };
  function repartoGasoleo() {
    const g = gasoleoAlm(); if (!g) return null;
    const k = `${db.ultimoCambio}|${db.labores.length}|${db.documentos.length}|${db.ajustes.length}|${db.equipos.length}|${g.id}`;
    if (cacheGas.k === k) return cacheGas.r;
    const ev = [];
    db.documentos.filter(d => d.tipo === 'albaran').forEach(d => (d.lineas || []).forEach(x => { const q = +x.cantidad || 0; if (x.productoId !== g.id || q <= 0) return; const imp = +x.importe || (+x.precio ? +x.precio * q : 0); ev.push({ f: d.fecha || '', o: 0, q, t: d.titular || '', pu: imp > 0 ? imp / q : null }); }));
    db.ajustes.filter(a => a.productoId === g.id).forEach(a => { const q = +a.cantidad || 0; if (q > 0) ev.push({ f: a.fecha || '', o: 0, q, t: a.titular || '', pu: null }); else if (q < 0) ev.push({ f: a.fecha || '', o: 1, q: -q, t: a.titular || '' }); });
    db.labores.forEach(l => { const q = litrosLabor(l); if (q) ev.push({ f: l.fecha || '', o: 2, q, l }); });
    ev.sort((a, b) => a.f.localeCompare(b.f) || a.o - b.o);
    const conPrecio = ev.filter(e => e.o === 0 && e.pu != null);
    const medio = conPrecio.length ? sum(conPrecio, e => e.pu * e.q) / sum(conPrecio, e => e.q) : null;
    const cola = []; const porLabor = new Map(); let ultimo = medio;
    const tomar = (q, soloDe) => {
      const partes = [];
      for (const x of cola) { if (q <= 1e-9) break; if (x.q <= 1e-9 || (soloDe != null && x.t !== soloDe)) continue; const c = Math.min(x.q, q); x.q -= c; q -= c; partes.push({ t: x.t, q: c, pu: x.pu }); }
      return { partes, falta: q > 1e-9 ? q : 0 };
    };
    for (const e of ev) {
      if (e.o === 0) { const pu = e.pu != null ? e.pu : ultimo; cola.push({ t: e.t, q: e.q, pu }); if (e.pu != null) ultimo = e.pu; continue; }
      if (e.o === 1) { tomar(e.q, e.t); continue; }
      const r = tomar(e.q, null);
      if (r.falta) r.partes.push({ t: titularDe(e.l.fincaId), q: r.falta, pu: ultimo, falta: true });
      const m = new Map(); r.partes.forEach(p => { const o = m.get(p.t) || { t: p.t, q: 0, imp: 0, sinPrecio: false }; o.q += p.q; if (p.pu != null) o.imp += p.q * p.pu; else o.sinPrecio = true; m.set(p.t, o); });
      porLabor.set(e.l.id, { partes: [...m.values()], falta: r.falta, coste: r.partes.every(p => p.pu != null) ? sum(r.partes, p => p.q * p.pu) : (medio != null ? sum(r.partes, p => p.q * (p.pu != null ? p.pu : medio)) : null) });
    }
    const r = { g, porLabor, cola: cola.filter(x => x.q > 1e-9), medio, ultimo };
    cacheGas = { k, r }; return r;
  }
  // Precio de los próximos litros que se gastarán (lo más antiguo que queda)
  function precioGasoleo(litros = 1) {
    const R = repartoGasoleo(); if (!R) return null;
    let q = 0, imp = 0;
    for (const x of R.cola) { if (q >= litros) break; const c = Math.min(x.q, litros - q); if (x.pu == null) continue; q += c; imp += c * x.pu; }
    return q > 0 ? imp / q : R.ultimo;
  }
  function costeGasLabor(l) { if (!litrosLabor(l)) return null; const R = repartoGasoleo(); const x = R && R.porLabor.get(l.id); return x && x.coste != null ? Math.round(x.coste * 100) / 100 : null; }
  function statsFinca(fid, c) {
    const f = finca(fid); const ha = (f && f.superficie) || 0;
    const kg = sum(db.cosechas.filter(x => x.fincaId === fid && x.campana === c), x => x.kg);
    const movs = db.movimientos.filter(m => m.fincaId === fid && m.campana === c);
    const ls = db.labores.filter(l => l.fincaId === fid && l.campana === c && l.equipoId);
    const amort = sum(ls, l => costeMaqLabor(l) || 0), gas = sum(ls, l => costeGasLabor(l) || 0), litros = sum(ls, litrosLabor);
    const maq = amort + gas;
    const gastos = sum(movs.filter(m => m.tipo === 'gasto'), m => m.importe) + maq;
    const ingresos = sum(movs.filter(m => m.tipo === 'ingreso'), m => m.importe);
    return { ha, kg, maq, amort, gas, litros, kgHa: ha && kg ? kg / ha : null, gastos, ingresos, margen: ingresos - gastos, margenHa: ha && (gastos || ingresos) ? (ingresos - gastos) / ha : null };
  }
  function statsGrupo(fs, c) {
    let ha = 0, haCos = 0, kg = 0, gastos = 0, ingresos = 0, maq = 0, amort = 0, gas = 0, litros = 0;
    for (const f of fs) { const s = statsFinca(f.id, c); ha += s.ha; kg += s.kg; if (s.kg > 0) haCos += s.ha; gastos += s.gastos; ingresos += s.ingresos; maq += s.maq; amort += s.amort; gas += s.gas; litros += s.litros; }
    return { ha, kg, maq, amort, gas, litros, kgHa: haCos ? kg / haCos : null, gastos, ingresos, margen: ingresos - gastos, margenHa: ha && (gastos || ingresos) ? (ingresos - gastos) / ha : null };
  }
  function statsCampana(c) {
    const cult = {};
    for (const f of db.fincas) {
      const cu = cultivoDe(f.id, c); const st = statsFinca(f.id, c);
      if (!cu && !st.kg && !st.gastos && !st.ingresos) continue;
      const k = cu || 'Sin cultivo';
      const o = cult[k] || (cult[k] = { cultivo: k, ha: 0, haCos: 0, kg: 0, gastos: 0, ingresos: 0 });
      o.ha += st.ha; o.kg += st.kg; if (st.kg > 0) o.haCos += st.ha; o.gastos += st.gastos; o.ingresos += st.ingresos;
    }
    const lista = Object.values(cult).map(o => Object.assign(o, { kgHa: o.haCos ? o.kg / o.haCos : null, margenHa: o.ha && (o.gastos || o.ingresos) ? (o.ingresos - o.gastos) / o.ha : null }))
      .sort((a, b) => b.ha - a.ha);
    const generales = db.movimientos.filter(m => m.campana === c && !finca(m.fincaId));
    const genG = sum(generales.filter(m => m.tipo === 'gasto'), m => m.importe);
    const genI = sum(generales.filter(m => m.tipo === 'ingreso'), m => m.importe);
    const movs = db.movimientos.filter(m => m.campana === c);
    const ingresos = sum(movs.filter(m => m.tipo === 'ingreso'), m => m.importe);
    const lsC = db.labores.filter(l => l.campana === c && l.equipoId && finca(l.fincaId));
    const maq = sum(lsC, l => costeMaqLabor(l) || 0), gas = sum(lsC, l => costeGasLabor(l) || 0);
    // el gasóleo ya está en los gastos con su factura: en las fincas solo se reparte, no se suma otra vez
    const gastos = sum(movs.filter(m => m.tipo === 'gasto'), m => m.importe) + maq;
    const haSemb = sum(lista.filter(o => o.cultivo !== 'Barbecho' && o.cultivo !== 'Sin cultivo'), o => o.ha);
    return { lista, genG, genI, ingresos, gastos, maq, gas, resultado: ingresos - gastos, kg: sum(lista, o => o.kg), haSemb };
  }

  // ---------- Estado de la interfaz y navegación (botón atrás de Android) ----------
  const ui = { tab: 'fincas', fincaId: null, grupo: null, buscar: '', titular: '', abiertos: new Set(), filtroLabor: { finca: '', tipo: '' }, filtroMov: 'todos', alm: { vista: 'inventario', titular: '' }, exp: { vista: 'agricultores', titular: '' }, resumenVista: 'campana', vistaFincas: 'lista', enfocar: null, enfocarGrupo: null, mapaAjustado: false };
  const capas = [];
  const profundidad = () => capas.length + (ui.grupo ? 1 : 0) + (ui.fincaId ? 1 : 0);
  // Si hay un «atrás» en curso (history.go es asíncrono), las ventanas nuevas esperan a que termine;
  // si no, el historial se descuadra y al cerrar se saldría de la app.
  let ignorarPop = 0;
  const pendientes = []; let relojPend = null;
  function soltarPendientes() { clearTimeout(relojPend); relojPend = null; ignorarPop = 0; pendientes.splice(0).forEach(d => { try { history.pushState({ d }, ''); } catch (e) { /* ignorar */ } }); }
  function empujar() {
    const d = profundidad();
    if (ignorarPop > 0) { pendientes.push(d); clearTimeout(relojPend); relojPend = setTimeout(soltarPendientes, 600); return; }
    try { history.pushState({ d }, ''); } catch (e) { /* ignorar */ }
  }
  function deshacerHasta(d) {
    while (profundidad() > d) {
      if (capas.length) { const c = capas.pop(); c.el.remove(); }
      else if (ui.fincaId) { ui.fincaId = null; render(); }
      else { ui.grupo = null; render(); }
    }
  }
  // Cerrar desde la app: se cierra al momento y se ignora el popstate que provoca history.go.
  function volverA(d) {
    const n = profundidad() - d; if (n <= 0) return;
    deshacerHasta(d);
    try { ignorarPop++; history.go(-n); } catch (e) { ignorarPop = 0; }
  }
  window.addEventListener('popstate', e => {
    if (ignorarPop > 0) { ignorarPop--; if (!ignorarPop && pendientes.length) soltarPendientes(); return; }
    deshacerHasta((e.state && e.state.d) || 0); // botón atrás del teléfono
  });
  try { history.replaceState({ d: 0 }, ''); } catch (e) { /* ignorar */ }

  // ---------- Iconos ----------
  const ico = {
    plus: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M12 5v14M5 12h14"/></svg>',
    x: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>',
    back: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M15 5l-7 7 7 7"/></svg>',
    carpeta: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M3 6h6l2 2h10v11H3z"/></svg>',
    camara: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linejoin="round"><path d="M4 7h3l2-3h6l2 3h3v12H4z"/><circle cx="12" cy="13" r="3.5"/></svg>',
    clip: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M20 11l-8 8a5 5 0 0 1-7-7l8-8a3.5 3.5 0 0 1 5 5l-8 8a2 2 0 0 1-3-3l7-7"/></svg>',
    mic: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="3" width="6" height="11" rx="3"/><path d="M5 11a7 7 0 0 0 14 0M12 18v3"/></svg>',
    chev: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>',
    edit: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 20h4L19 9l-4-4L4 16z"/><path d="M13.5 6.5l4 4"/></svg>'
  };

  // ---------- Piezas de interfaz ----------
  const chip = c => c ? `<span class="chip" style="--h:${hue(c)}">${esc(c)}</span>` : '<span class="chip none">Sin cultivo</span>';
  const fechaBox = iso => { if (!iso) return '<div class="date"></div>'; const [y, m, d] = iso.split('-'); return `<div class="date"><b>${+d}</b><span>${MESES[+m - 1]} ${y.slice(2)}</span></div>`; };
  const stat = (k, v, cls = '') => `<div class="stat"><div class="k">${k}</div><div class="v ${cls}">${v}</div></div>`;
  const vacio = (t, p, btns = '') => `<div class="card empty"><h3>${t}</h3><p>${p}</p>${btns ? `<div class="btns">${btns}</div>` : ''}</div>`;
  const signo = v => v > 0 ? 'pos' : v < 0 ? 'neg' : '';
  function barras(items, max) {
    return `<div class="bars">${items.map(i => `<div class="bar"><div class="lbl"><span>${esc(i.label)}</span><b>${i.valor}</b></div><div class="track"><div class="fill" style="--h:${i.h};width:${max ? Math.max(2, i.n / max * 100) : 0}%"></div></div></div>`).join('')}</div>`;
  }
  function filaLabor(l, conFinca = true, n = [l]) {
    const maqs = equipo(l.equipoId) ? equipo(l.equipoId).nombre : [l.tractorId, l.aperoId].map(id => (maquina(id) || {}).nombre).filter(Boolean).join(' + ');
    const det = [l.producto, l.dosis != null ? `${fmtN(l.dosis)} ${l.unidad || ''}` : '', l.motivo, maqs].filter(Boolean).join(' · ');
    const maqC = sum(n, x => (costeMaqLabor(x) || 0) + (costeGasLabor(x) || 0));
    const coste = sum(n, x => x.coste) + maqC;
    const aviso = n.some(x => x.avisos && x.avisos.length);
    return `<button class="row" data-act="editLabor" data-id="${l.id}">${fechaBox(l.fecha)}<div class="main"><div class="t">${esc(l.tipo)}${conFinca ? ' · ' + esc(nombreLote(n)) : ''}</div>${det ? `<div class="s">${esc(det)}</div>` : ''}${aviso ? '<span class="pill neg">⚠ Revisar normativa</span>' : ''}</div>${coste ? `<div class="end small muted">${eur(coste)}${maqC ? '<div class="small">con maquinaria</div>' : ''}</div>` : ''}</button>`;
  }
  const notaMaq = st => st.maq ? `<p class="small muted" style="margin:0">Los gastos incluyen <b>${eur(st.maq)}</b> de maquinaria propia: ${[st.amort ? eur(st.amort) + ' de amortización' : '', st.gas ? `${eur(st.gas)} de gasóleo (${fmtN(st.litros, 0)} L)` : ''].filter(Boolean).join(' y ')}.</p>` : '';
  function filaMov(m, conFinca = true, n = [m]) {
    const sub = [m.categoria, conFinca ? (finca(m.fincaId) ? nombreLote(n) : 'General') : '', m.laborId ? 'desde el cuaderno' : ''].filter(Boolean).join(' · ');
    const imp = sum(n, x => x.importe); const v = m.tipo === 'gasto' ? -imp : imp;
    return `<button class="row" data-act="editMov" data-id="${m.id}">${fechaBox(m.fecha)}<div class="main"><div class="t">${esc(m.concepto || m.categoria)}</div><div class="s">${esc(sub)}</div></div><div class="end big ${signo(v)}">${v > 0 ? '+' : ''}${eur(v, 2)}</div></button>`;
  }
  function filaCosecha(x, conFinca = true, n = [x]) {
    const sub = [x.humedad != null ? `${fmtN(x.humedad, 1)} % humedad` : '', x.notas].filter(Boolean).join(' · ');
    return `<button class="row" data-act="editCosecha" data-id="${x.id}">${fechaBox(x.fecha)}<div class="main"><div class="t">${conFinca ? esc(nombreLote(n)) : 'Cosecha'}</div>${sub ? `<div class="s">${esc(sub)}</div>` : ''}</div><div class="end big">${fmtKg(sum(n, y => y.kg))}</div></button>`;
  }
  function bannerAvisos() {
    if (db.ejemplo) return `<div class="banner"><p><b>Estás viendo datos de ejemplo.</b> Bórralos cuando quieras empezar con los tuyos.</p><button class="btn" data-act="borrarEjemplo">Borrar ejemplo</button></div>`;
    if (!db.fincas.length) return '';
    const dias = db.ultimaCopia ? Math.floor((Date.now() - new Date(db.ultimaCopia)) / 864e5) : null;
    const pendiente = !db.ultimaCopia || (db.ultimoCambio && db.ultimoCambio > db.ultimaCopia);
    if (!pendiente || (dias != null && dias < 7)) return '';
    const txt = dias == null ? 'Todavía no has hecho ninguna copia de seguridad.' : `Tu última copia es de hace ${dias} días y hay cambios sin guardar en ella.`;
    return `<div class="banner"><p>${txt} Si pierdes el móvil, la copia es lo que te permite recuperar los datos.</p><button class="btn" data-act="compartirCopia">Hacer copia</button></div>`;
  }

  // ---------- Vistas ----------
  function vFincas() {
    let h = bannerAvisos();
    const segs = `<div class="segs" role="group" aria-label="Vista">${[['lista', 'Lista'], ['mapa', 'Mapa']].map(([k, t]) => `<button data-act="vistaFincas" data-k="${k}" aria-pressed="${ui.vistaFincas === k}">${t}</button>`).join('')}</div>`;
    if (ui.vistaFincas === 'mapa') return h + segs + vMapaInfo();
    if (!db.fincas.length) {
      return h + vacio('Aún no tienes fincas', 'Da de alta tus fincas con su superficie, tráelas de un Excel o tócalas en el mapa de SIGPAC. Después podrás apuntar labores, gastos y cosechas de cada una.',
        '<button class="btn primary" data-act="nuevaFinca">Añadir finca</button><button class="btn" data-act="importarExcel">Importar desde Excel</button><button class="btn" data-act="vistaFincas" data-k="mapa">Buscar en el mapa</button><button class="btn" data-act="ejemplo">Ver con datos de ejemplo</button>');
    }
    const c = db.campanaActiva;
    h += segs;
    const ha = sum(db.fincas, f => f.superficie);
    const sembrado = sum(db.fincas.filter(f => { const cu = cultivoDe(f.id, c); return cu && cu !== 'Barbecho'; }), f => f.superficie);
    h += `<div class="stats">${stat('Parcelas', db.fincas.length)}${stat('Superficie', fmtN(ha) + ' ha')}${stat('Sembrado ' + c, fmtN(sembrado) + ' ha')}</div>`;
    const tits = titulares();
    h += `<div class="filters"><input id="fBuscar" type="search" placeholder="Buscar nombre, pueblo, polígono-parcela…" value="${esc(ui.buscar)}" autocomplete="off" aria-label="Buscar finca">${tits.length > 1 ? `<select id="fTitular" aria-label="Filtrar por titular"><option value="">Todos los titulares</option>${tits.map(t => `<option${ui.titular === t ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select>` : ''}</div>`;
    h += `<div id="listaFincas">${listaFincasHTML()}</div>`;
    return h;
  }
  function coincide(f, q) {
    if (ui.titular && String(f.titular || '') !== ui.titular) return false;
    if (!q) return true;
    const t = norm([f.nombre, f.grupo, f.municipio, f.titular, `${f.poligono || ''} ${f.parcela || ''}`, cultivoDe(f.id, db.campanaActiva)].join(' '));
    return q.split(' ').every(w => t.includes(w));
  }
  function filaFinca(f, enGrupo) {
    const sig = enGrupo
      ? [(f.poligono || f.parcela) ? `Pol. ${f.poligono || '–'} · Parc. ${f.parcela || '–'}${f.recinto ? ' · Rec. ' + f.recinto : ''}` : '', fmtN(f.superficie) + ' ha', f.titular].filter(Boolean).join(' · ')
      : [f.municipio, (f.poligono || f.parcela) ? `Pol. ${f.poligono || '–'} · Parc. ${f.parcela || '–'}` : '', fmtN(f.superficie) + ' ha', f.titular].filter(Boolean).join(' · ');
    return `<button class="row${enGrupo ? ' sub' : ''}" data-act="verFinca" data-id="${f.id}"><div class="main"><div class="t">${esc(f.nombre)}</div><div class="s">${esc(sig)}</div></div><div class="end">${chip(cultivoDe(f.id, db.campanaActiva))}</div></button>`;
  }
  function chipsCultivos(fs) {
    const c = db.campanaActiva; const m = {};
    fs.forEach(f => { const cu = cultivoDe(f.id, c) || ''; m[cu] = (m[cu] || 0) + (+f.superficie || 0); });
    const ks = Object.keys(m).sort((a, b) => m[b] - m[a]);
    return `<div class="chips end-chips">${ks.slice(0, 2).map(k => chip(k || null)).join('')}${ks.length > 2 ? `<span class="muted small">+${ks.length - 2}</span>` : ''}</div>`;
  }
  function listaFincasHTML() {
    const q = norm(ui.buscar);
    const vis = db.fincas.filter(f => coincide(f, q));
    if (!vis.length) return '<p class="note">Ninguna finca coincide con la búsqueda.</p>';
    const grupos = new Map(); const sueltas = [];
    vis.forEach(f => { const g = grupoDe(f); if (g) { if (!grupos.has(g)) grupos.set(g, []); grupos.get(g).push(f); } else sueltas.push(f); });
    const items = [...[...grupos].map(([g, fs]) => ({ n: g, g, fs })), ...sueltas.map(f => ({ n: f.nombre, f }))]
      .sort((a, b) => a.n.localeCompare(b.n, 'es', { numeric: true }));
    return '<div class="list">' + items.map(it => {
      if (it.f) return filaFinca(it.f, false);
      const todas = fincasDeGrupo(it.g); const abierto = !!(q || ui.titular) || ui.abiertos.has(it.g);
      const munis = [...new Set(todas.map(f => f.municipio).filter(Boolean))]; const tits = [...new Set(todas.map(f => f.titular).filter(Boolean))];
      const sub = [`${todas.length} parcelas`, fmtN(sum(todas, f => f.superficie)) + ' ha', munis.join(', '), tits.join(', ')].filter(Boolean).join(' · ');
      return `<div class="grow${abierto ? ' open' : ''}"><div class="ghead"><button class="row" data-act="verGrupo" data-g="${esc(it.g)}"><div class="main"><div class="t">${esc(it.g)} <span class="gbadge">Agrupada</span></div><div class="s">${esc(sub)}</div></div>${chipsCultivos(todas)}</button>
        <button class="iconbtn gtoggle" data-act="toggleGrupo" data-g="${esc(it.g)}" aria-expanded="${abierto}" aria-label="${abierto ? 'Ocultar' : 'Ver'} parcelas de ${esc(it.g)}">${ico.chev}</button></div>
        ${abierto ? `<div class="gsub">${it.fs.sort(byNombre).map(f => filaFinca(f, true)).join('')}</div>` : ''}</div>`;
    }).join('') + '</div>';
  }

  function vGrupo(g) {
    const fs = fincasDeGrupo(g).sort(byNombre);
    if (!fs.length) { ui.grupo = null; return vFincas(); }
    const c = db.campanaActiva; const st = statsGrupo(fs, c); const ids = new Set(fs.map(f => f.id));
    const munis = [...new Set(fs.map(f => f.municipio).filter(Boolean))]; const tits = [...new Set(fs.map(f => f.titular).filter(Boolean))];
    const porCult = {}; fs.forEach(f => { const cu = cultivoDe(f.id, c) || ''; porCult[cu] = (porCult[cu] || 0) + (+f.superficie || 0); });
    const G = esc(g);
    let h = `<div class="backhdr"><button class="iconbtn" data-act="atras" aria-label="Volver">${ico.back}</button><h2>${G}</h2><button class="iconbtn" data-act="editGrupo" data-g="${G}" aria-label="Cambiar nombre o parcelas">${ico.edit}</button></div>
      <div class="meta"><span class="gbadge">Finca agrupada</span><span><b>${fmtN(st.ha)} ha</b></span><span>${fs.length} parcelas</span>${munis.length ? `<span>${esc(munis.join(', '))}</span>` : ''}</div>
      ${tits.length ? `<div class="meta"><span>Titular${tits.length > 1 ? 'es' : ''}: <b>${esc(tits.join(', '))}</b></span></div>` : ''}`;
    h += `<div class="card pad" style="display:flex;flex-direction:column;gap:12px">
      <div><div class="tag">Campaña ${c}</div><div class="chips">${Object.entries(porCult).sort((a, b) => b[1] - a[1]).map(([cu, ha]) => `<span>${chip(cu || null)} <span class="small muted">${fmtN(ha)} ha</span></span>`).join('')}</div></div>
      <div class="stats">${stat('Rendimiento', st.kgHa ? fmtN(st.kgHa, 0) + ' kg/ha' : '—')}${stat('Gastos', eur(st.gastos))}${stat('Margen', st.margenHa != null ? eur(st.margenHa) + '/ha' : '—', signo(st.margen))}</div>${notaMaq(st)}
      <div class="actions"><button class="btn sm" data-act="nuevaLabor" data-grupo="${G}">${ico.plus} Labor</button><button class="btn sm" data-act="nuevoMov" data-grupo="${G}" data-tipo="gasto">${ico.plus} Gasto</button><button class="btn sm" data-act="nuevoMov" data-grupo="${G}" data-tipo="ingreso">${ico.plus} Ingreso</button><button class="btn sm" data-act="nuevaCosecha" data-grupo="${G}">${ico.plus} Cosecha</button></div>
      <p class="small muted" style="margin:0">Lo que apuntes aquí se guarda en cada parcela. Costes, importes y kilos se reparten según la superficie de cada una.</p></div>`;
    h += `<h2 class="sec">Parcelas</h2><div class="list">${fs.map(f => filaFinca(f, true)).join('')}</div>
      <div class="btns" style="justify-content:flex-start">${fs.some(f => f.geo && f.geo.c.length) ? `<button class="btn sm" data-act="verMapaGrupo" data-g="${G}">Ver en el mapa</button>` : ''}<button class="btn sm" data-act="editGrupo" data-g="${G}">Añadir o quitar parcelas</button></div>`;
    const lab = agruparLote(db.labores.filter(l => ids.has(l.fincaId) && l.campana === c).sort(byFechaDesc));
    h += `<h2 class="sec">Labores · ${c}</h2>` + (lab.length ? `<div class="list">${lab.map(({ x, n }) => filaLabor(x, true, n)).join('')}</div>` : '<p class="note">Sin labores apuntadas en esta campaña.</p>');
    const cos = agruparLote(db.cosechas.filter(x => ids.has(x.fincaId) && x.campana === c).sort(byFechaDesc));
    if (cos.length) h += `<h2 class="sec">Cosecha · ${fmtKg(st.kg)}</h2><div class="list">${cos.map(({ x, n }) => filaCosecha(x, true, n)).join('')}</div>`;
    const movs = agruparLote(db.movimientos.filter(m => ids.has(m.fincaId) && m.campana === c).sort(byFechaDesc));
    h += `<h2 class="sec">Gastos e ingresos · ${c}</h2>` + (movs.length ? `<div class="list">${movs.map(({ x, n }) => filaMov(x, true, n)).join('')}</div>` : '<p class="note">Sin gastos ni ingresos apuntados en esta campaña.</p>');
    h += `<div class="btns" style="justify-content:flex-start;margin-top:6px"><button class="btn sm danger" data-act="deshacerGrupo" data-g="${G}">Deshacer agrupación</button></div>`;
    return h;
  }

  function vDetalle(id) {
    const f = finca(id); if (!f) { ui.fincaId = null; return vFincas(); }
    const c = db.campanaActiva; const s = siembraDe(id, c); const st = statsFinca(id, c);
    const sig = [f.poligono && `Pol. ${f.poligono}`, f.parcela && `Parc. ${f.parcela}`, f.recinto && `Rec. ${f.recinto}`].filter(Boolean).join(' · ');
    const refs = refsFinca(f);
    let h = `<div class="backhdr"><button class="iconbtn" data-act="atras" aria-label="Volver">${ico.back}</button><h2>${esc(f.nombre)}</h2><button class="iconbtn" data-act="editFinca" data-id="${id}" aria-label="Editar finca">${ico.edit}</button></div>
      <div class="meta"><span><b>${fmtN(f.superficie)} ha</b></span>${f.municipio ? `<span>${esc(f.municipio)}</span>` : ''}${sig ? `<span>SIGPAC ${esc(sig)}</span>` : ''}</div>
      ${f.titular ? `<div class="meta"><span>Titular: <b>${esc(f.titular)}</b></span></div>` : ''}`;
    const g = grupoDe(f);
    if (g) {
      const fsG = fincasDeGrupo(g);
      h += `<div class="card pad sp"><div style="min-width:0"><div class="tag">Finca agrupada</div><b>${esc(g)}</b> <span class="small muted">· ${fsG.length} parcelas · ${fmtN(sum(fsG, x => x.superficie))} ha</span></div><button class="btn sm" data-act="verGrupo" data-g="${esc(g)}">Ver finca</button></div>`;
    } else {
      h += `<div class="card pad sp"><span class="small muted">Esta parcela no está agrupada con otras.</span><button class="btn sm" data-act="agrupar" data-id="${id}">Agrupar</button></div>`;
    }
    if (refs.length) {
      const enMapa = f.geo && f.geo.c.length && f.geo.k === refsKey(f);
      const noEnc = f.geo && !f.geo.c.length && f.geo.k === refsKey(f);
      h += `<div class="card pad sigpac"><div class="sp"><div style="min-width:0"><div class="tag">SIGPAC${f.provincia ? ' · ' + esc(nombreProv(f.provincia)) : ''}</div>
        <div class="refs">${refs.slice(0, 4).map(r => `<code>${refTxt(r)}</code>`).join(' ')}${refs.length > 4 ? ` <span class="muted small">y ${refs.length - 4} más</span>` : ''}</div>
        <div class="small ${noEnc ? 'neg' : 'muted'}">${enMapa ? `Superficie SIGPAC: <b>${fmtN(f.geo.ha)} ha</b>` : noEnc ? 'No aparece en SIGPAC: revisa los códigos' : 'Contorno pendiente de buscar'}</div></div>
        ${enMapa ? `<button class="btn sm" data-act="verMapa" data-id="${id}">Ver en el mapa</button>` : `<button class="btn sm" data-act="contornos">Buscar</button>`}</div></div>`;
    } else {
      h += `<div class="card pad sp"><span class="small muted">Sin referencia SIGPAC: no se puede situar en el mapa.</span><button class="btn sm" data-act="editFinca" data-id="${id}">Añadir referencia</button></div>`;
    }
    h += `<div class="card pad" style="display:flex;flex-direction:column;gap:12px">
      <div class="sp"><div><div class="tag">Campaña ${c}</div><div class="chips">${chip(s && s.cultivo)}${s && s.variedad ? `<span class="muted small">${esc(s.variedad)}</span>` : ''}</div>${s && (s.laboreo || s.estiercol) ? `<div class="small muted" style="margin-top:4px">${[s.laboreo ? 'Laboreo: ' + esc(etiquetaLaboreo(s.laboreo)) : '', s.estiercol ? 'Con estiércol' : ''].filter(Boolean).join(' · ')}</div>` : ''}</div>
      <button class="btn sm" data-act="asignar" data-id="${id}">${s ? 'Cambiar cultivo' : 'Asignar cultivo'}</button></div>
      <div class="stats">${stat('Rendimiento', st.kgHa ? fmtN(st.kgHa, 0) + ' kg/ha' : '—')}${stat('Gastos', eur(st.gastos))}${stat('Margen', st.margenHa != null ? eur(st.margenHa) + '/ha' : '—', signo(st.margen))}</div>${notaMaq(st)}
      <div class="actions"><button class="btn sm" data-act="nuevaLabor" data-finca="${id}">${ico.plus} Labor</button><button class="btn sm" data-act="nuevoMov" data-finca="${id}" data-tipo="gasto">${ico.plus} Gasto</button><button class="btn sm" data-act="nuevoMov" data-finca="${id}" data-tipo="ingreso">${ico.plus} Ingreso</button><button class="btn sm" data-act="nuevaCosecha" data-finca="${id}">${ico.plus} Cosecha</button></div></div>`;
    const lab = db.labores.filter(l => l.fincaId === id && l.campana === c).sort(byFechaDesc);
    h += `<h2 class="sec">Labores · ${c}</h2>` + (lab.length ? `<div class="list">${lab.map(l => filaLabor(l, false)).join('')}</div>` : '<p class="note">Sin labores apuntadas en esta campaña.</p>');
    const cos = db.cosechas.filter(x => x.fincaId === id && x.campana === c).sort(byFechaDesc);
    if (cos.length) h += `<h2 class="sec">Cosecha · ${fmtKg(st.kg)}</h2><div class="list">${cos.map(x => filaCosecha(x, false)).join('')}</div>`;
    const movs = db.movimientos.filter(m => m.fincaId === id && m.campana === c).sort(byFechaDesc);
    h += `<h2 class="sec">Gastos e ingresos · ${c}</h2>` + (movs.length ? `<div class="list">${movs.map(m => filaMov(m, false)).join('')}</div>` : '<p class="note">Sin gastos ni ingresos apuntados en esta campaña.</p>');
    // Historial de rotación
    const camps = listaCampanas().filter(cc => siembraDe(id, cc) || db.cosechas.some(x => x.fincaId === id && x.campana === cc) || db.movimientos.some(m => m.fincaId === id && m.campana === cc));
    if (camps.length) {
      h += `<h2 class="sec">Rotación e historial</h2><div class="tablewrap"><table><thead><tr><th>Campaña</th><th>Cultivo</th><th>Laboreo</th><th>kg/ha</th><th>Margen €/ha</th></tr></thead><tbody>${camps.map(cc => {
        const x = statsFinca(id, cc); const sc = siembraDe(id, cc);
        return `<tr><td>${cc}</td><td style="text-align:left">${chip(cultivoDe(id, cc))}</td><td style="text-align:left" class="small">${sc ? esc(etiquetaLaboreo(sc.laboreo)) + (sc.estiercol ? ' · estiércol' : '') : ''}</td><td>${x.kgHa ? fmtN(x.kgHa, 0) : '—'}</td><td class="${signo(x.margen)}">${x.margenHa != null ? fmtN(x.margenHa, 0) : '—'}</td></tr>`;
      }).join('')}</tbody></table></div>`;
    }
    if (f.notas) h += `<h2 class="sec">Notas</h2><div class="card pad" style="white-space:pre-wrap">${esc(f.notas)}</div>`;
    return h;
  }

  function vCuaderno() {
    const c = db.campanaActiva; const fl = ui.filtroLabor;
    if (!db.fincas.length) return vacio('Primero añade tus fincas', 'Las labores se apuntan siempre en una finca.', '<button class="btn primary" data-act="nuevaFinca">Añadir finca</button>');
    const enFiltro = l => { if (!fl.finca) return true; if (fl.finca.startsWith('g:')) return grupoDe(finca(l.fincaId)) === fl.finca.slice(2); return l.fincaId === fl.finca; };
    let lab = db.labores.filter(l => l.campana === c && enFiltro(l) && (!fl.tipo || l.tipo === fl.tipo)).sort(byFechaDesc);
    const gruposF = [...listaGrupos().keys()].sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));
    let h = `<div class="filters"><select id="fLabFinca" aria-label="Filtrar por finca"><option value="">Todas las fincas</option>${gruposF.length ? `<optgroup label="Fincas agrupadas">${gruposF.map(g => `<option value="g:${esc(g)}"${fl.finca === 'g:' + g ? ' selected' : ''}>${esc(g)}</option>`).join('')}</optgroup><optgroup label="Parcelas">` : ''}${db.fincas.slice().sort(byNombre).map(f => `<option value="${f.id}"${fl.finca === f.id ? ' selected' : ''}>${esc(f.nombre)}</option>`).join('')}${gruposF.length ? '</optgroup>' : ''}</select>
      <select id="fLabTipo" aria-label="Filtrar por labor"><option value="">Todas las labores</option>${TIPOS_LABOR.map(t => `<option${fl.tipo === t ? ' selected' : ''}>${t}</option>`).join('')}</select></div>`;
    if (!lab.length) return h + vacio(`Sin labores en ${c}`, 'Apunta cada vez que labres, siembres, abones o trates una finca. Con fecha, producto y dosis tendrás el cuaderno de campo al día.', '<button class="btn primary" data-act="nuevaLabor">Apuntar labor</button>');
    let mes = '';
    let lista = '';
    for (const { x: l, n } of agruparLote(lab)) {
      const m = l.fecha.slice(0, 7);
      if (m !== mes) { if (mes) lista += '</div>'; mes = m; lista += `<h2 class="sec">${MESES_L[+m.slice(5) - 1]} ${m.slice(0, 4)}</h2><div class="list">`; }
      lista += filaLabor(l, true, n);
    }
    return h + lista + '</div>';
  }

  function vCuentas() {
    const c = db.campanaActiva; const st = statsCampana(c);
    const movs = db.movimientos.filter(m => m.campana === c && (ui.filtroMov === 'todos' || m.tipo === ui.filtroMov)).sort(byFechaDesc);
    let h = `<div class="stats">${stat('Ingresos', eur(st.ingresos), 'pos')}${stat('Gastos', eur(st.gastos), 'neg')}${stat('Resultado', eur(st.resultado), signo(st.resultado))}</div>`;
    if (st.maq) h += `<p class="note" style="margin-top:-4px">Los gastos incluyen <b>${eur(st.maq)}</b> de amortización de la maquinaria propia, calculada con los equipos de trabajo. No es dinero que salga ahora, pero sí lo que cuesta hacer las labores.</p>`;
    if (st.gas) h += `<p class="note" style="margin-top:-4px">En las fincas se reparten <b>${eur(st.gas)}</b> de gasóleo gastado por los equipos. Aquí no se suma otra vez: ya cuenta con la factura del gasóleo.</p>`;
    h += `<div class="segs" role="group" aria-label="Mostrar">${[['todos', 'Todo'], ['gasto', 'Gastos'], ['ingreso', 'Ingresos']].map(([k, t]) => `<button data-act="filtroMov" data-k="${k}" aria-pressed="${ui.filtroMov === k}">${t}</button>`).join('')}</div>`;
    const tipo = ui.filtroMov === 'ingreso' ? 'ingreso' : 'gasto';
    const porCat = {};
    db.movimientos.filter(m => m.campana === c && m.tipo === tipo).forEach(m => { porCat[m.categoria] = (porCat[m.categoria] || 0) + (m.importe || 0); });
    if (tipo === 'gasto' && st.maq) porCat['Maquinaria propia (amortización)'] = st.maq;
    const cats = Object.entries(porCat).sort((a, b) => b[1] - a[1]);
    if (cats.length) {
      h += `<h2 class="sec">${tipo === 'gasto' ? 'Gastos' : 'Ingresos'} por tipo</h2><div class="card pad">${barras(cats.map(([k, v]) => ({ label: k, valor: eur(v), n: v, h: tipo === 'gasto' ? 15 : 130 })), cats[0][1])}</div>`;
    }
    if (!movs.length) return h + vacio(`Sin apuntes en ${c}`, 'Apunta semilla, abono, gasóleo, seguros, ventas de grano o la PAC. Puedes asignarlos a una finca o dejarlos como generales.', '<button class="btn primary" data-act="nuevoMov">Apuntar gasto o ingreso</button>');
    return h + `<h2 class="sec">Apuntes</h2><div class="list">${agruparLote(movs).map(({ x, n }) => filaMov(x, true, n)).join('')}</div>`;
  }

  function vCosechas() {
    const c = db.campanaActiva;
    if (!db.fincas.length) return vacio('Primero añade tus fincas', 'La cosecha se apunta por finca para calcular los kilos por hectárea.', '<button class="btn primary" data-act="nuevaFinca">Añadir finca</button>');
    const fs = db.fincas.filter(f => { const cu = cultivoDe(f.id, c); return (cu && cu !== 'Barbecho') || db.cosechas.some(x => x.fincaId === f.id && x.campana === c); }).sort(byNombre);
    const ent = db.cosechas.filter(x => x.campana === c).sort(byFechaDesc);
    const kg = sum(ent, x => x.kg);
    const haCos = sum(fs.filter(f => ent.some(x => x.fincaId === f.id)), f => f.superficie);
    let h = `<div class="stats">${stat('Producción', fmtKg(kg))}${stat('Cosechado', fmtN(haCos) + ' ha')}${stat('Media', haCos ? fmtN(kg / haCos, 0) + ' kg/ha' : '—')}</div>`;
    if (!fs.length) return h + vacio(`Sin cultivos asignados en ${c}`, 'Asigna a cada finca lo que has sembrado esta campaña para ver aquí el seguimiento de la cosecha.', '<button class="btn primary" data-act="irTab" data-tab="fincas">Ir a fincas</button>');
    h += `<h2 class="sec">Por finca</h2><div class="list">${fs.map(f => {
      const st = statsFinca(f.id, c);
      return `<button class="row" data-act="nuevaCosecha" data-finca="${f.id}"><div class="main"><div class="t">${esc(f.nombre)}</div><div class="s">${chip(cultivoDe(f.id, c))} · ${fmtN(f.superficie)} ha</div></div><div class="end">${st.kg ? `<div class="big">${fmtN(st.kgHa, 0)} kg/ha</div><div class="s muted small">${fmtKg(st.kg)}</div>` : '<span class="muted small">Pendiente</span>'}</div></button>`;
    }).join('')}</div><p class="note">Toca una finca para apuntar un remolque o el total cosechado.</p>`;
    if (ent.length) h += `<h2 class="sec">Entradas</h2><div class="list">${agruparLote(ent).map(({ x, n }) => filaCosecha(x, true, n)).join('')}</div>`;
    return h;
  }

  function vResumen() {
    const segsR = `<div class="segs" role="group" aria-label="Apartado">${[['campana', 'Campaña'], ['titulares', 'Titulares'], ['pac', 'PAC']].map(([k, t]) => `<button data-act="resumenVista" data-k="${k}" aria-pressed="${ui.resumenVista === k}">${t}</button>`).join('')}</div>`;
    if (db.fincas.length && ui.resumenVista === 'titulares') return segsR + vTitulares();
    if (db.fincas.length && ui.resumenVista === 'pac') return segsR + vPAC();
    return (db.fincas.length ? segsR : '') + vResumenCampana();
  }
  function vResumenCampana() {
    const c = db.campanaActiva; const st = statsCampana(c);
    if (!db.fincas.length) return vacio('Sin datos todavía', 'El resumen compara rendimientos y rentabilidad por cultivo, por finca y entre campañas.', '<button class="btn primary" data-act="nuevaFinca">Añadir finca</button><button class="btn" data-act="ejemplo">Ver con datos de ejemplo</button>');
    let h = `<div class="stats">${stat('Sembrado', fmtN(st.haSemb) + ' ha')}${stat('Producción', fmtKg(st.kg))}${stat('Resultado', eur(st.resultado), signo(st.resultado))}</div>`;
    if (st.lista.length) {
      const maxHa = Math.max(...st.lista.map(o => o.ha));
      h += `<h2 class="sec">Superficie por cultivo · ${c}</h2><div class="card pad">${barras(st.lista.map(o => ({ label: o.cultivo, valor: fmtN(o.ha) + ' ha', n: o.ha, h: hue(o.cultivo) })), maxHa)}</div>`;
      h += `<h2 class="sec">Por cultivo</h2><div class="tablewrap"><table><thead><tr><th>Cultivo</th><th>ha</th><th>kg/ha</th><th>Ingresos</th><th>Gastos</th><th>Margen €/ha</th></tr></thead><tbody>${st.lista.map(o =>
        `<tr><td>${chip(o.cultivo)}</td><td>${fmtN(o.ha)}</td><td>${o.kgHa ? fmtN(o.kgHa, 0) : '—'}</td><td>${eur(o.ingresos)}</td><td>${eur(o.gastos)}</td><td class="${signo(o.ingresos - o.gastos)}">${o.margenHa != null ? fmtN(o.margenHa, 0) : '—'}</td></tr>`).join('')}</tbody></table></div>`;
    }
    if (st.genG || st.genI) h += `<p class="note">Gastos generales sin finca: <b>${eur(st.genG)}</b>${st.genI ? ` · ingresos generales (PAC, etc.): <b>${eur(st.genI)}</b>` : ''}. No se reparten por cultivo, pero sí cuentan en el resultado total.</p>`;
    const fs = db.fincas.slice().sort(byNombre);
    h += `<h2 class="sec">Por finca · ${c}</h2><div class="tablewrap"><table><thead><tr><th>Finca</th><th>Cultivo</th><th>ha</th><th>kg/ha</th><th>Margen €/ha</th></tr></thead><tbody>${fs.map(f => {
      const x = statsFinca(f.id, c);
      return `<tr><td>${esc(f.nombre)}</td><td style="text-align:left">${chip(cultivoDe(f.id, c))}</td><td>${fmtN(f.superficie)}</td><td>${x.kgHa ? fmtN(x.kgHa, 0) : '—'}</td><td class="${signo(x.margen)}">${x.margenHa != null ? fmtN(x.margenHa, 0) : '—'}</td></tr>`;
    }).join('')}</tbody></table></div>`;
    // Comparativa entre campañas
    const camps = listaCampanas();
    const porCultivo = {};
    for (const cc of camps) for (const o of statsCampana(cc).lista) { if (o.cultivo === 'Sin cultivo') continue; (porCultivo[o.cultivo] = porCultivo[o.cultivo] || []).push(Object.assign({ campana: cc }, o)); }
    const cultivos = Object.keys(porCultivo).sort((a, b) => db.cultivos.indexOf(a) - db.cultivos.indexOf(b));
    if (cultivos.length && camps.length > 1) {
      h += `<h2 class="sec">Comparativa entre campañas</h2><div class="tablewrap"><table><thead><tr><th>Campaña</th><th>ha</th><th>kg/ha</th><th>Producción</th><th>Margen €/ha</th></tr></thead><tbody>${cultivos.map(cu =>
        `<tr class="grp"><td colspan="5">${chip(cu)}</td></tr>` + porCultivo[cu].map(o => `<tr><td>${o.campana}</td><td>${fmtN(o.ha)}</td><td>${o.kgHa ? fmtN(o.kgHa, 0) : '—'}</td><td>${o.kg ? fmtKg(o.kg) : '—'}</td><td class="${signo(o.ingresos - o.gastos)}">${o.margenHa != null ? fmtN(o.margenHa, 0) : '—'}</td></tr>`).join('')).join('')}</tbody></table></div>
        <p class="note">kg/ha se calcula solo con las fincas que ya tienen cosecha apuntada.</p>`;
    }
    return h;
  }

  // ---------- Mapa (ortofoto PNOA + linderos SIGPAC) ----------
  let mapa = null, mapaEl = null, capaFincas = null, capaSel = null, capaSigpac = null, capaPos = null;
  const icoMapa = {
    loc: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><circle cx="12" cy="12" r="4"/><path d="M12 2v3M12 19v3M2 12h3M19 12h3"/></svg>',
    fit: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 9V4h5M20 9V4h-5M4 15v5h5M20 15v5h-5"/></svg>',
    sig: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linejoin="round"><path d="M3 4h8v7H3zM13 4h8v10h-8zM3 13h8v7H3zM13 16h8v4h-8z"/></svg>'
  };
  function vMapaInfo() {
    const conGeo = db.fincas.filter(f => f.geo && f.geo.c.length);
    const sinRef = db.fincas.filter(f => !refsKey(f));
    const noEnc = db.fincas.filter(f => f.geo && !f.geo.c.length && f.geo.k === refsKey(f));
    const pend = db.fincas.filter(f => refsKey(f) && (!f.geo || f.geo.k !== refsKey(f)));
    const cults = [...new Set(conGeo.map(f => cultivoDe(f.id, db.campanaActiva)).filter(Boolean))];
    let h = '<div id="mapSlot"></div>';
    if (cults.length) h += `<div class="leyenda">${cults.map(chip).join('')}</div>`;
    h += `<div class="card pad mapinfo"><p><b>Toca una parcela</b> del mapa para ver su referencia SIGPAC y añadirla a tus fincas.</p>`;
    if (db.fincas.length) h += `<p class="small">${conGeo.length} de ${db.fincas.length} fincas en el mapa.${sinRef.length ? ` ${sinRef.length} sin referencia SIGPAC (${sinRef.slice(0, 3).map(f => esc(f.nombre)).join(', ')}${sinRef.length > 3 ? '…' : ''}).` : ''}${noEnc.length ? ` <span class="neg">${noEnc.length} no aparecen en SIGPAC: revisa su polígono y parcela.</span>` : ''}${pend.length ? ` ${pend.length} pendientes de buscar.` : ''}</p>`;
    h += `<div class="btns" style="justify-content:flex-start">${db.fincas.some(f => refsKey(f)) ? '<button class="btn sm" data-act="contornos">Actualizar contornos</button>' : ''}<button class="btn sm" data-act="importarExcel">Importar Excel</button></div></div>`;
    h += '<p class="note">Ortofoto: PNOA © Instituto Geográfico Nacional. Recintos: SIGPAC © FEGA. El mapa de fondo necesita cobertura; tus fincas se ven siempre.</p>';
    return h;
  }
  function crearMapa() {
    mapaEl = document.createElement('div'); mapaEl.className = 'mapa';
    mapa = L.map(mapaEl, { zoomControl: true, attributionControl: true, maxZoom: 20 }).setView([40.4, -3.7], 6);
    mapa.attributionControl.setPrefix('');
    L.tileLayer(PNOA, { tms: true, maxZoom: 20, maxNativeZoom: 19, attribution: 'PNOA © IGN' }).addTo(mapa);
    capaSigpac = L.tileLayer.wms(SIGPAC + '/wms', { layers: 'recinto', format: 'image/png', transparent: true, minZoom: 14, maxZoom: 20, attribution: 'SIGPAC © FEGA' }).addTo(mapa);
    capaFincas = L.featureGroup().addTo(mapa);
    const ctl = L.control({ position: 'topright' });
    ctl.onAdd = () => {
      const d = L.DomUtil.create('div', 'mapctl');
      d.innerHTML = `<button type="button" data-m="loc" aria-label="Mi ubicación" title="Mi ubicación">${icoMapa.loc}</button><button type="button" data-m="fit" aria-label="Ver todas mis fincas" title="Ver todas mis fincas">${icoMapa.fit}</button><button type="button" data-m="sig" aria-pressed="true" aria-label="Linderos SIGPAC" title="Linderos SIGPAC">${icoMapa.sig}</button>`;
      L.DomEvent.disableClickPropagation(d);
      d.addEventListener('click', e => {
        const b = e.target.closest('[data-m]'); if (!b) return;
        if (b.dataset.m === 'loc') { toast('Buscando tu ubicación…'); mapa.locate({ setView: true, maxZoom: 17, enableHighAccuracy: true }); }
        if (b.dataset.m === 'fit') ajustarMapa(true);
        if (b.dataset.m === 'sig') {
          const on = mapa.hasLayer(capaSigpac); if (on) mapa.removeLayer(capaSigpac); else capaSigpac.addTo(mapa);
          b.setAttribute('aria-pressed', String(!on));
        }
      });
      return d;
    };
    ctl.addTo(mapa);
    mapa.on('click', e => identificar(e.latlng));
    mapa.on('zoomend', () => mapaEl.classList.toggle('etiquetas', mapa.getZoom() >= 14));
    mapa.on('popupclose', () => { if (capaSel) { capaSel.remove(); capaSel = null; } });
    mapa.on('locationfound', e => {
      if (capaPos) capaPos.remove();
      capaPos = L.layerGroup([L.circle(e.latlng, { radius: e.accuracy, weight: 1, color: '#3b82f6', fillOpacity: .12, interactive: false }), L.circleMarker(e.latlng, { radius: 8, weight: 3, color: '#fff', fillColor: '#3b82f6', fillOpacity: 1, interactive: false })]).addTo(mapa);
    });
    mapa.on('locationerror', () => toast('No se ha podido obtener tu ubicación. Activa el GPS y da permiso a la app.'));
    // Clics en botones dentro de los globos del mapa
    mapaEl.addEventListener('click', e => {
      const b = e.target.closest('.leaflet-popup [data-act]');
      if (b && acciones[b.dataset.act]) { e.stopPropagation(); acciones[b.dataset.act](b); }
    });
  }
  function pintarFincas() {
    capaFincas.clearLayers();
    const c = db.campanaActiva;
    db.fincas.forEach(f => {
      if (!f.geo || !f.geo.c.length) return;
      const cu = cultivoDe(f.id, c);
      const col = cu ? `hsl(${hue(cu)} 90% 58%)` : '#ffffff';
      const lyr = L.geoJSON({ type: 'MultiPolygon', coordinates: f.geo.c }, { style: { color: col, weight: 3, fillColor: col, fillOpacity: .28 } });
      lyr.fincaId = f.id;
      if (!grupoDe(f)) lyr.bindTooltip(esc(f.nombre), { permanent: true, direction: 'center', className: 'etq' });
      lyr.on('click', ev => { L.DomEvent.stopPropagation(ev); globoFinca(f, ev.latlng); });
      lyr.addTo(capaFincas);
    });
    // Un solo nombre en el centro de cada finca agrupada
    const porGrupo = new Map();
    capaFincas.getLayers().forEach(l => { const g = grupoDe(finca(l.fincaId)); if (g) { if (!porGrupo.has(g)) porGrupo.set(g, L.latLngBounds([])); porGrupo.get(g).extend(l.getBounds()); } });
    porGrupo.forEach((b, g) => { if (b.isValid()) L.marker(b.getCenter(), { interactive: false, keyboard: false, icon: L.divIcon({ className: 'etq-g', html: `<span>${esc(g)}</span>`, iconSize: null }) }).addTo(capaFincas); });
  }
  function globoFinca(f, ll) {
    const cu = cultivoDe(f.id, db.campanaActiva); const st = statsFinca(f.id, db.campanaActiva);
    L.popup({ maxWidth: 260 }).setLatLng(ll).setContent(`<div class="pop"><div class="t">${esc(f.nombre)}</div><div>${chip(cu)} <span class="muted">${fmtN(f.superficie)} ha</span></div>
      ${st.kgHa ? `<div class="small">${fmtN(st.kgHa, 0)} kg/ha en ${db.campanaActiva}</div>` : ''}${f.geo && f.geo.ha ? `<div class="small muted">SIGPAC: ${fmtN(f.geo.ha)} ha</div>` : ''}
      ${f.titular ? `<div class="small">Titular: <b>${esc(f.titular)}</b></div>` : ''}
      ${grupoDe(f) ? `<div class="small">Finca agrupada: <b>${esc(grupoDe(f))}</b></div>` : ''}
      <div class="popbtns"><button class="btn sm primary" data-act="verFinca" data-id="${f.id}">Abrir parcela</button>${grupoDe(f) ? `<button class="btn sm" data-act="verGrupo" data-g="${esc(grupoDe(f))}">Ver finca</button>` : ''}</div></div>`).openOn(mapa);
  }
  function ajustarMapa(avisar) {
    const b = capaFincas.getBounds();
    if (b.isValid()) mapa.fitBounds(b, { padding: [24, 24], maxZoom: 17 });
    else if (avisar) toast('Todavía no hay fincas situadas en el mapa');
  }
  async function identificar(ll) {
    const pop = L.popup({ maxWidth: 280 }).setLatLng(ll).setContent('<div class="pop">Consultando SIGPAC…</div>').openOn(mapa);
    if (mapa.getZoom() < 13) { pop.setContent('<div class="pop">Acércate más al terreno para elegir una parcela.</div>'); return; }
    try {
      const r = await fetch(`${SIGPAC}/servicioconsultassigpac/query/recinfobypoint/4258/${ll.lng.toFixed(6)}/${ll.lat.toFixed(6)}.json`);
      if (!r.ok) throw new Error();
      const d = await r.json(); const x = Array.isArray(d) ? d[0] : null;
      if (!mapa.hasLayer(pop)) return;
      if (!x) { pop.setContent('<div class="pop">Aquí no hay ningún recinto SIGPAC.</div>'); return; }
      ui.recSel = x;
      if (capaSel) capaSel.remove();
      capaSel = L.geoJSON({ type: 'MultiPolygon', coordinates: wktCoords(x.wkt) }, { style: { color: '#ffd400', weight: 3, dashArray: '6 5', fillColor: '#ffd400', fillOpacity: .15 }, interactive: false }).addTo(mapa);
      const ya = fincaConRecinto(x);
      const ref = refTxt({ pr: x.provincia, mu: x.municipio, ag: x.agregado, zo: x.zona, po: x.poligono, pa: x.parcela, re: x.recinto });
      pop.setContent(`<div class="pop"><div class="tag">Recinto SIGPAC</div><div class="t">${ref}</div>
        <div class="small">${esc(nombreProv(x.provincia))} · Pol. ${x.poligono} · Parc. ${x.parcela} · Rec. ${x.recinto}</div>
        <div class="small"><b>${fmtN(x.superficie)} ha</b> · ${esc(USOS[x.uso_sigpac] || x.uso_sigpac || '')}${x.pendiente_media != null ? ` · pendiente ${fmtN(x.pendiente_media, 1)} %` : ''}</div>
        ${ya ? `<div class="small">Forma parte de <b>${esc(ya.nombre)}</b></div><button class="btn sm primary" data-act="verFinca" data-id="${ya.id}">Abrir ficha</button>`
          : `<div class="popbtns"><button class="btn sm primary" data-act="fincaDesdeRecinto">Crear finca</button>${db.fincas.length ? '<button class="btn sm" data-act="anadirRecinto">Añadir a una finca</button>' : ''}</div>`}</div>`);
    } catch (e) {
      pop.setContent('<div class="pop">No se ha podido consultar SIGPAC. Comprueba que tienes cobertura.</div>');
    }
  }
  function montarMapa() {
    const slot = $('#mapSlot'); if (!slot) return;
    if (!window.L) { slot.outerHTML = '<div class="card pad">No se ha podido cargar el mapa.</div>'; return; }
    if (!mapa) crearMapa();
    slot.replaceWith(mapaEl);
    mapa.invalidateSize();
    pintarFincas();
    if (ui.enfocar) {
      const lyr = capaFincas.getLayers().find(l => l.fincaId === ui.enfocar);
      if (lyr) { mapa.fitBounds(lyr.getBounds(), { padding: [30, 30], maxZoom: 17 }); globoFinca(finca(ui.enfocar), lyr.getBounds().getCenter()); }
      ui.enfocar = null; ui.mapaAjustado = true;
    } else if (ui.enfocarGrupo) {
      const b = L.latLngBounds([]); capaFincas.getLayers().forEach(l => { if (l.fincaId && grupoDe(finca(l.fincaId)) === ui.enfocarGrupo) b.extend(l.getBounds()); });
      if (b.isValid()) mapa.fitBounds(b, { padding: [30, 30], maxZoom: 17 });
      ui.enfocarGrupo = null; ui.mapaAjustado = true;
    } else if (!ui.mapaAjustado) { ajustarMapa(false); ui.mapaAjustado = capaFincas.getLayers().length > 0; }
    setTimeout(() => mapa.invalidateSize(), 50);
    actualizarContornos({ silencioso: true });
  }

  // ---------- Render principal ----------
  function render() {
    $('#campBtn').textContent = db.campanaActiva + ' ▾';
    $('#title').textContent = ui.fincaId ? 'Parcela' : ui.grupo ? 'Finca agrupada' : TITULOS[ui.tab];
    document.querySelectorAll('.tabs button').forEach(b => b.setAttribute('aria-current', b.dataset.tab === ui.tab ? 'page' : 'false'));
    const main = $('#main');
    main.innerHTML = ui.fincaId ? vDetalle(ui.fincaId) : ui.grupo ? vGrupo(ui.grupo) : ({ fincas: vFincas, cuaderno: vCuaderno, cuentas: vCuentas, cosechas: vCosechas, almacen: vAlmacen, explotacion: vExplotacion, resumen: vResumen }[ui.tab])();
    const fab = { fincas: ['nuevaFinca', 'Finca'], cuaderno: ['nuevaLabor', 'Labor'], cuentas: ['nuevoMov', 'Apunte'], cosechas: ['nuevaCosecha', 'Cosecha'], almacen: ui.alm.vista === 'factura' ? ['nuevoDoc" data-k="factura', 'Factura'] : ['nuevoDoc" data-k="albaran', 'Albarán'], explotacion: ui.exp.vista === 'maquinaria' ? ['nuevaMaquina', 'Máquina'] : ['nuevoAgricultor', 'Agricultor'] }[ui.tab];
    const mostrarFab = fab && !ui.fincaId && !ui.grupo && (db.fincas.length || ui.tab === 'almacen' || ui.tab === 'explotacion') && !(ui.tab === 'fincas' && ui.vistaFincas === 'mapa');
    $('#fabSlot').innerHTML = mostrarFab ? `<button class="fab" data-act="${fab[0]}">${ico.plus}${fab[1]}</button>` : '';
    montarMapa();
  }

  // ---------- Hojas (formularios) ----------
  function abrirCapa(html, onMount) {
    const el = document.createElement('div');
    el.className = 'scrim';
    el.innerHTML = html;
    $('#layer').appendChild(el);
    capas.push({ el });
    empujar();
    const d = profundidad();
    el.addEventListener('click', e => { if (e.target === el || e.target.closest('[data-close]')) volverA(d - 1); });
    if (onMount) onMount(el, d);
    return el;
  }
  const cabecera = t => `<header><h3>${t}</h3><button class="iconbtn" data-close aria-label="Cerrar">${ico.x}</button></header>`;

  function confirmar(titulo, texto, ok = 'Borrar') {
    return new Promise(res => {
      let resuelto = false;
      abrirCapa(`<div class="sheet" role="alertdialog" aria-modal="true">${cabecera(titulo)}<div class="body"><p style="margin:0">${texto}</p></div>
        <footer><button class="btn" data-close>Cancelar</button><button class="btn primary" data-ok style="background:var(--neg)">${esc(ok)}</button></footer></div>`, (el, d) => {
        el.querySelector('[data-ok]').addEventListener('click', () => { resuelto = true; res({ ok: true, d }); });
        const obs = new MutationObserver(() => { if (!el.isConnected) { obs.disconnect(); if (!resuelto) res({ ok: false }); } });
        obs.observe($('#layer'), { childList: true });
      });
    });
  }

  // ---------- Selector de varias fincas (para labores, gastos y cosechas) ----------
  function pickerHTML(sel) {
    const s = new Set(sel || []); const c = db.campanaActiva;
    const fila = x => `<label class="pk-f" data-txt="${esc(norm([x.nombre, x.grupo, x.municipio, x.titular, (x.poligono || '') + ' ' + (x.parcela || '')].join(' ')))}" data-cu="${esc(cultivoDe(x.id, c) || '')}"><input type="checkbox" value="${x.id}"${s.has(x.id) ? ' checked' : ''}><span class="pk-n">${esc(x.nombre)}</span><span class="muted small">${fmtN(x.superficie)} ha</span></label>`;
    const items = [...[...listaGrupos()].map(([g, fs]) => ({ n: g, g, fs })), ...db.fincas.filter(x => !grupoDe(x)).map(x => ({ n: x.nombre, x }))]
      .sort((a, b) => a.n.localeCompare(b.n, 'es', { numeric: true }));
    const cults = [...new Set(db.fincas.map(x => cultivoDe(x.id, c)).filter(Boolean))].sort((a, b) => db.cultivos.indexOf(a) - db.cultivos.indexOf(b));
    return `<div class="picker">
      <input type="search" class="pk-buscar" placeholder="Buscar finca, pueblo, polígono-parcela…" aria-label="Buscar finca" autocomplete="off">
      ${cults.length ? `<div class="pk-rapido"><span class="small muted">Marcar todas las de ${c}:</span>${cults.map(cu => `<button type="button" class="chip pk-cu" style="--h:${hue(cu)}" data-cu="${esc(cu)}">${esc(cu)}</button>`).join('')}<button type="button" class="chip none pk-nada">Quitar todas</button></div>` : ''}
      <div class="pk-list">${items.map(it => it.x ? fila(it.x) : `<div class="pk-g"><label class="pk-gh"><input type="checkbox" data-g="${esc(it.g)}"><b class="pk-n">${esc(it.g)}</b><span class="muted small">${it.fs.length} parc. · ${fmtN(sum(it.fs, y => y.superficie))} ha</span></label>${it.fs.sort(byNombre).map(fila).join('')}</div>`).join('')}</div>
      <div class="pk-sum small"></div></div>`;
  }
  function montarPicker(root, onChange) {
    const lista = root.querySelector('.pk-list'), sumEl = root.querySelector('.pk-sum');
    const act = () => {
      root.querySelectorAll('.pk-g').forEach(g => {
        const cs = [...g.querySelectorAll('.pk-f input')]; const n = cs.filter(i => i.checked).length;
        const gh = g.querySelector('.pk-gh input'); gh.checked = n === cs.length; gh.indeterminate = n > 0 && n < cs.length;
      });
      const ids = [...lista.querySelectorAll('.pk-f input:checked')].map(i => i.value);
      sumEl.textContent = ids.length ? `${ids.length} ${ids.length === 1 ? 'parcela elegida' : 'parcelas elegidas'} · ${fmtN(sum(ids.map(finca), f => f.superficie))} ha` : 'Ninguna parcela elegida';
      if (onChange) onChange(ids);
    };
    root.addEventListener('change', e => {
      const g = e.target.closest('.pk-gh');
      if (g) g.parentElement.querySelectorAll('.pk-f').forEach(l => { if (!l.hidden) l.querySelector('input').checked = e.target.checked; });
      act();
    });
    root.addEventListener('click', e => {
      const b = e.target.closest('.pk-cu');
      if (b) { lista.querySelectorAll('.pk-f').forEach(l => { if (l.dataset.cu === b.dataset.cu) l.querySelector('input').checked = true; }); act(); }
      if (e.target.closest('.pk-nada')) { lista.querySelectorAll('input').forEach(i => { i.checked = false; }); act(); }
    });
    root.querySelector('.pk-buscar').addEventListener('input', e => {
      const q = norm(e.target.value);
      lista.querySelectorAll('.pk-f').forEach(l => { l.hidden = !!q && !q.split(' ').every(w => l.dataset.txt.includes(w)); });
      lista.querySelectorAll('.pk-g').forEach(g => { const fs = [...g.querySelectorAll('.pk-f')]; g.hidden = fs.every(l => l.hidden) && !(q && norm(g.querySelector('.pk-gh').textContent).includes(q)); if (!g.hidden && fs.every(l => l.hidden)) fs.forEach(l => { l.hidden = false; }); });
    });
    act();
    const prim = lista.querySelector('.pk-f input:checked');
    if (prim) lista.scrollTop = Math.max(0, prim.closest('.pk-f').offsetTop - lista.offsetTop - 50);
  }
  function idsIniciales(pre) {
    if (pre.grupo) return fincasDeGrupo(pre.grupo).map(f => f.id);
    if (pre.finca) return [pre.finca];
    const fl = ui.filtroLabor.finca;
    if (ui.tab === 'cuaderno' && fl) return fl.startsWith('g:') ? fincasDeGrupo(fl.slice(2)).map(f => f.id) : [fl];
    return [];
  }

  // spec: {titulo, campos, valores, guardar(v) -> string|undefined (error), borrar()}
  function formulario(spec) {
    const vals = Object.assign({}, spec.valores);
    const campoHTML = f => {
      const id = 'f_' + f.k; const v = vals[f.k];
      let inp;
      if (f.type === 'select') inp = `<select id="${id}" name="${f.k}">${f.opts(vals).map(([val, t]) => `<option value="${esc(val)}"${String(v ?? '') === String(val) ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select>`;
      else if (f.type === 'fincas') inp = pickerHTML(v);
      else if (f.type === 'textarea') inp = `<textarea id="${id}" name="${f.k}" rows="3" placeholder="${esc(f.placeholder || '')}">${esc(v)}</textarea>`;
      else if (f.type === 'num') inp = `<input id="${id}" name="${f.k}" type="text" inputmode="decimal" autocomplete="off" value="${esc(numStr(v))}" placeholder="${esc(f.placeholder || '')}">`;
      else inp = `<input id="${id}" name="${f.k}" type="${f.type || 'text'}" value="${esc(v)}" placeholder="${esc(f.placeholder || '')}" autocomplete="off"${f.type === 'text' || !f.type ? ' autocapitalize="sentences"' : ''}${f.lista ? ` list="dl_${f.k}"` : ''}>${f.lista ? `<datalist id="dl_${f.k}">${f.lista().map(o => `<option value="${esc(o)}">`).join('')}</datalist>` : ''}`;
      return `<div class="field" data-k="${f.k}"><label for="${id}">${f.label}${f.req || f.sinOpcional ? '' : ' <span class="muted small">(opcional)</span>'}</label>${inp}${f.hint ? `<div class="hint">${f.hint}</div>` : ''}<div class="err" hidden></div></div>`;
    };
    let cuerpo = ''; const cs = spec.campos;
    for (let i = 0; i < cs.length; i++) {
      if (cs[i].half && cs[i + 1] && cs[i + 1].half) { cuerpo += `<div class="grid2" data-pair="${cs[i].k}">${campoHTML(cs[i])}${campoHTML(cs[i + 1])}</div>`; i++; }
      else cuerpo += campoHTML(cs[i]);
    }
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera(spec.titulo)}<form class="body" id="frm" novalidate>${cuerpo}${spec.extra || ''}</form>
      <footer>${spec.borrar ? '<button class="btn danger" type="button" data-del>Borrar</button>' : ''}<button class="btn primary" type="submit" form="frm">Guardar</button></footer></div>`, (el, d) => {
      const form = el.querySelector('form');
      const leer = () => {
        for (const f of cs) {
          if (f.type === 'fincas') { vals[f.k] = [...form.querySelectorAll(`.field[data-k="${f.k}"] .pk-f input:checked`)].map(i => i.value); continue; }
          const i = form.elements[f.k]; if (i) vals[f.k] = i.value;
        }
        return vals;
      };
      cs.filter(f => f.type === 'fincas').forEach(f => montarPicker(form.querySelector(`.field[data-k="${f.k}"] .picker`)));
      const refrescar = (origen) => {
        leer();
        for (const f of cs) {
          const w = form.querySelector(`.field[data-k="${f.k}"]`);
          if (f.show) w.hidden = !f.show(vals);
          if (f.dyn && origen !== f.k) {
            const sel = form.elements[f.k]; const ops = f.opts(vals); const prev = sel.value;
            sel.innerHTML = ops.map(([val, t]) => `<option value="${esc(val)}">${esc(t)}</option>`).join('');
            sel.value = ops.some(o => String(o[0]) === prev) ? prev : (ops[0] ? ops[0][0] : '');
          }
        }
        form.querySelectorAll('.grid2').forEach(g => { g.hidden = [...g.children].every(c => c.hidden); });
        if (spec.alCambiar) spec.alCambiar(vals, form, origen);
      };
      form.addEventListener('change', e => refrescar(e.target.name));
      form.addEventListener('input', e => {
        if (spec.alEscribir) spec.alEscribir(e.target.name, leer(), (k, v) => { form.elements[k].value = v; });
        if (spec.alCambiar && e.target.name) spec.alCambiar(leer(), form, e.target.name);
      });
      if (spec.alMontar) spec.alMontar(() => refrescar());
      refrescar();
      form.addEventListener('submit', async e => {
        e.preventDefault(); leer();
        const out = {}; let ok = true;
        for (const f of cs) {
          const w = form.querySelector(`.field[data-k="${f.k}"]`); const err = w.querySelector('.err'); err.hidden = true;
          if (w.hidden || (w.parentElement && w.parentElement.hidden)) { out[f.k] = f.type === 'num' ? null : ''; continue; }
          let v = vals[f.k];
          if (f.type === 'fincas') {
            if (f.req && !v.length) { err.textContent = 'Elige al menos una finca'; err.hidden = false; ok = false; continue; }
            out[f.k] = v; continue;
          }
          if (f.type === 'num') {
            v = parseNum(v, f.miles);
            if (Number.isNaN(v)) { err.textContent = 'Escribe un número, por ejemplo 12,5'; err.hidden = false; ok = false; continue; }
          } else v = String(v || '').trim();
          if (f.req && (v === null || v === '')) { err.textContent = 'Este dato es obligatorio'; err.hidden = false; ok = false; continue; }
          out[f.k] = v;
        }
        if (!ok) { const first = form.querySelector('.err:not([hidden])'); if (first) first.parentElement.scrollIntoView({ block: 'center', behavior: 'smooth' }); return; }
        if (spec.antesDeGuardar && !(await spec.antesDeGuardar(out))) return;
        const r = spec.guardar(out);
        if (typeof r === 'string') { toast(r); return; }
        guardar(); volverA(d - 1); render(); toast('Guardado');
        if (spec.despues) spec.despues();
      });
      const del = el.querySelector('[data-del]');
      if (del) del.addEventListener('click', async () => {
        const r = await confirmar('¿Borrar?', spec.textoBorrar || 'Esta acción no se puede deshacer.');
        if (!r.ok) return;
        const destino = spec.borrar() === 'salirDetalle' ? 0 : d - 1;
        guardar(); volverA(destino); render(); toast('Borrado');
      });
    });
  }

  const optsFincas = () => db.fincas.slice().sort(byNombre).map(f => [f.id, `${f.nombre} (${fmtN(f.superficie)} ha)`]);
  const optsCamp = () => listaCampanas().map(c => [c, c]);
  const campoCampana = { k: 'campana', label: 'Campaña', type: 'select', req: 1, opts: optsCamp, hint: 'La campaña va de septiembre a agosto.' };

  function formFinca(id, pre) {
    const f = id ? finca(id) : null;
    formulario({
      titulo: f ? 'Editar finca' : 'Nueva finca',
      valores: f || Object.assign({ agregado: 0, zona: 0 }, pre || {}),
      campos: [
        { k: 'nombre', label: 'Nombre', req: 1, placeholder: 'Ej.: Sarrana 663-180' },
        { k: 'superficie', label: 'Superficie (ha)', type: 'num', req: 1, placeholder: 'Ej.: 12,45' },
        { k: 'titular', label: 'Titular', half: 1, placeholder: 'Ej.: Iván', lista: titulares },
        { k: 'grupo', label: 'Finca agrupada', half: 1, placeholder: 'Ej.: Sarrana', lista: () => [...listaGrupos().keys()] },
        { k: 'municipio', label: 'Municipio', placeholder: 'Nombre del pueblo' },
        { k: 'provincia', label: 'Provincia (código)', type: 'num', half: 1, placeholder: 'Ej.: 47' },
        { k: 'codMunicipio', label: 'Municipio (código)', type: 'num', half: 1, placeholder: 'Ej.: 215' },
        { k: 'poligono', label: 'Polígono', half: 1, placeholder: 'SIGPAC' },
        { k: 'parcela', label: 'Parcela', half: 1, placeholder: 'SIGPAC' },
        { k: 'recinto', label: 'Recinto(s)', placeholder: 'Ej.: 1, 2, 5-7 · vacío = toda la parcela' },
        { k: 'agregado', label: 'Agregado', type: 'num', half: 1, placeholder: '0' },
        { k: 'zona', label: 'Zona', type: 'num', half: 1, placeholder: '0' },
        { k: 'otras', label: 'Otras parcelas de esta finca', type: 'textarea', placeholder: 'Una por línea: polígono-parcela-recinto\nEj.: 403-36-1', hint: 'Los códigos los ves en tu declaración de la PAC o en el visor SIGPAC. Lo más fácil: toca la parcela en el mapa de la app.' },
        { k: 'notas', label: 'Notas', type: 'textarea', placeholder: 'Propiedad o arrendada, tipo de suelo, linderos…' }
      ],
      guardar(v) {
        if (v.superficie <= 0) return 'La superficie tiene que ser mayor que cero';
        if ((v.poligono || v.parcela) && (!v.provincia || !v.codMunicipio)) return 'Para situarla en el mapa faltan los códigos de provincia y municipio';
        if (f) Object.assign(f, v); else db.fincas.push(Object.assign({ id: uid() }, v));
      },
      despues: () => actualizarContornos({ silencioso: true }),
      textoBorrar: 'Se borrarán la finca, sus cultivos, labores y cosechas. Sus gastos e ingresos se mantienen como generales.',
      borrar: f && (() => {
        db.fincas = db.fincas.filter(x => x.id !== f.id);
        db.siembras = db.siembras.filter(x => x.fincaId !== f.id);
        const labIds = new Set(db.labores.filter(x => x.fincaId === f.id).map(x => x.id));
        db.labores = db.labores.filter(x => x.fincaId !== f.id);
        db.cosechas = db.cosechas.filter(x => x.fincaId !== f.id);
        db.movimientos = db.movimientos.filter(m => !labIds.has(m.laborId));
        db.movimientos.forEach(m => { if (m.fincaId === f.id) m.fincaId = ''; });
        return 'salirDetalle';
      })
    });
  }

  // Desde el mapa: recinto tocado -> finca nueva o añadirlo a una existente
  function formAnadirRecinto(x) {
    const ref = refTxt({ pr: x.provincia, mu: x.municipio, ag: x.agregado, zo: x.zona, po: x.poligono, pa: x.parcela, re: x.recinto });
    formulario({
      titulo: 'Añadir recinto a una finca',
      valores: { fincaId: optsFincas()[0][0], sumar: 'si' },
      campos: [
        { k: 'fincaId', label: 'Finca', type: 'select', req: 1, opts: optsFincas, hint: `Recinto ${ref} · ${fmtN(x.superficie)} ha` },
        { k: 'sumar', label: 'Superficie', type: 'select', req: 1, opts: () => [['si', `Sumar ${fmtN(x.superficie)} ha a la finca`], ['no', 'Dejar la superficie como está']] }
      ],
      guardar(v) {
        const f = finca(v.fincaId);
        const mismaParcela = +f.provincia === x.provincia && +f.codMunicipio === x.municipio && +(f.agregado || 0) === x.agregado && +(f.zona || 0) === x.zona && +f.poligono === x.poligono && +f.parcela === x.parcela;
        if (!f.poligono || !f.parcela) Object.assign(f, { provincia: x.provincia, codMunicipio: x.municipio, agregado: x.agregado, zona: x.zona, poligono: String(x.poligono), parcela: String(x.parcela), recinto: String(x.recinto) });
        else if (mismaParcela) { const recs = listaRecintos(f.recinto); if (recs.length && !recs.includes(x.recinto)) f.recinto = [...recs, x.recinto].sort((a, b) => a - b).join(', '); }
        else f.otras = (f.otras ? f.otras.trim() + '\n' : '') + ref;
        if (v.sumar === 'si') f.superficie = Math.round(((+f.superficie || 0) + x.superficie) * 100) / 100;
      },
      despues: () => actualizarContornos({ silencioso: true })
    });
  }

  function formAsignar(fid) {
    const c = db.campanaActiva; const s = siembraDe(fid, c);
    formulario({
      titulo: `Cultivo de ${esc(finca(fid).nombre)}`,
      valores: { campana: c, cultivo: s ? s.cultivo : db.cultivos[0], variedad: s ? s.variedad : '', laboreo: s ? s.laboreo || '' : '', estiercol: s && s.estiercol ? 'si' : 'no' },
      campos: [
        { k: 'cultivo', label: 'Cultivo', type: 'select', req: 1, opts: () => [...db.cultivos.map(x => [x, x]), ['__ninguno', 'Sin cultivo (quitar)']] },
        { k: 'variedad', label: 'Variedad', placeholder: 'Ej.: Marcopolo, Pedrezuela, Castellana…', show: v => v.cultivo !== '__ninguno' && v.cultivo !== 'Barbecho' },
        { k: 'laboreo', label: 'Laboreo', type: 'select', half: 1, opts: () => [['', '—'], ...Object.entries(LABOREO)], show: v => v.cultivo !== '__ninguno' },
        { k: 'estiercol', label: 'Estiércol', type: 'select', half: 1, opts: () => [['no', 'No'], ['si', 'Sí']], show: v => v.cultivo !== '__ninguno' },
        campoCampana
      ],
      guardar(v) {
        db.siembras = db.siembras.filter(x => !(x.fincaId === fid && x.campana === v.campana));
        if (v.cultivo !== '__ninguno') db.siembras.push({ id: uid(), fincaId: fid, campana: v.campana, cultivo: v.cultivo, variedad: v.variedad, laboreo: v.laboreo, estiercol: v.estiercol === 'si' });
      }
    });
  }

  // Productos del almacén que se pueden usar en una labor (según el titular de cada parcela)
  function opcionesAlmacen(v, actual) {
    const tipoP = TIPO_DE_LABOR[v.tipo]; if (!tipoP) return [['', '—']];
    const tits = [...new Set((v.fincas || []).map(titularDe))];
    const ps = db.productos.filter(p => p.tipo === tipoP).map(p => ({ p, ex: tits.map(t => existencias(p.id, t)) }))
      .filter(x => x.ex.some(e => e > 0) || x.p.id === actual);
    const prim = v.tipo === 'Tratamiento' ? ['', ps.length ? 'Elige el producto' : 'No hay fitosanitarios en el almacén'] : ['', 'No descontar del almacén'];
    return [prim, ...ps.sort((a, b) => a.p.nombre.localeCompare(b.p.nombre, 'es')).map(x => [x.p.id, `${x.p.nombre} · quedan ${tits.length > 1 ? tits.map((t, i) => `${t || '¿?'} ${fmtN(x.ex[i], 3)}`).join(', ') : fmtN(x.ex[0] || 0, 3)} ${x.p.unidad}`])];
  }
  // Cantidad de producto que gasta una labor en una parcela
  function gastoEnParcela(p, f, v, fs) {
    const u = String(v.unidad || '').toLowerCase();
    if (!(v.dosis > 0)) return null;
    let q, uq;
    if (u === 'kg' || u === 'l') { const T = sum(fs, x => +x.superficie || 0); q = v.dosis * (T ? (+f.superficie || 0) / T : 1 / fs.length); uq = u === 'kg' ? 'kg' : 'L'; }
    else { const b = baseHa(v.dosis, v.unidad); if (!b) return null; const ha = fs.length === 1 && v.superficie > 0 ? v.superficie : (+f.superficie || 0); q = b.v * ha; uq = b.u; }
    if (p.unidad === uq) return q;
    if (p.unidad === 't' && uq === 'kg') return q / 1000;
    return null;
  }
  function revisarLabor(v) {
    const p = producto(v.prodAlm); const fs = (v.fincas || []).map(finca).filter(Boolean);
    const out = { porCultivo: [], graves: [], notas: [], consumo: [], sinCalculo: false };
    if (!p) return out;
    const reg = regProducto(p);
    const grupos = new Map();
    fs.forEach(f => { const cu = cultivoDe(f.id, v.campana || db.campanaActiva) || ''; if (!grupos.has(cu)) grupos.set(cu, []); grupos.get(cu).push(f); });
    if (p.tipo === 'Fitosanitario') {
      if (!reg) out.notas.push(p.registro ? `No encuentro el nº ${p.registro} en el Registro incluido en la app: compruébalo en la ficha oficial.` : 'Este producto no está enlazado con el Registro oficial: la app no puede comprobar cultivos ni dosis.');
      grupos.forEach((ff, cu) => {
        const r = revisarFito(reg, cu, v.dosis, v.unidad);
        out.porCultivo.push({ cu, fs: ff, r });
        r.graves.forEach(g => { if (!out.graves.includes(g)) out.graves.push(g); });
        r.notas.forEach(g => { if (!out.notas.includes(g)) out.notas.push(g); });
      });
    }
    const porTit = new Map();
    fs.forEach(f => { const q = gastoEnParcela(p, f, v, fs); if (q == null) { out.sinCalculo = true; return; } const t = titularDe(f.id); porTit.set(t, (porTit.get(t) || 0) + q); });
    porTit.forEach((q, t) => { const ex = existencias(p.id, t) + (v._propio && v._propio[t] || 0); out.consumo.push({ t, q, ex }); if (q > ex + 1e-9) out.graves.push(`${t || 'El titular'} solo tiene ${fmtCant(ex, p.unidad)} de «${p.nombre}» y la labor gasta ${fmtCant(q, p.unidad)}.`); });
    if (out.sinCalculo && v.dosis > 0) out.notas.push(`No puedo calcular cuánto se gasta: la dosis va en ${v.unidad} y el producto se mide en ${p.unidad}.`);
    return out;
  }
  function infoLaborHTML(v) {
    const p = producto(v.prodAlm);
    if (!p) return TIPO_DE_LABOR[v.tipo] === 'Fitosanitario' ? '<p class="small muted" style="margin:0">Solo puedes elegir fitosanitarios que estén en el almacén del titular. Si falta alguno, sube primero su albarán en la pestaña Almacén.</p>' : '';
    const reg = regProducto(p); const r = revisarLabor(v);
    let h = `<div class="tag">${esc(p.nombre)}${p.registro ? ' · Nº reg. ' + esc(p.registro) : ''}${reg ? ` · <span class="${reg.e === 'Vigente' ? 'pos' : 'neg'}">${esc(reg.e)}</span>` : ''}</div>`;
    if (p.tipo === 'Fitosanitario' && !FITOS) h += '<p class="small muted">Cargando el Registro oficial…</p>';
    r.porCultivo.forEach(({ cu, fs, r: rr }) => {
      const ok = rr.autorizado; const n = `${fs.length} parcela${fs.length === 1 ? '' : 's'}`;
      h += `<div class="fito-cu ${ok === false ? 'mal' : ok ? 'bien' : ''}"><b>${ok === false ? '✗' : ok ? '✓' : '?'} ${esc(cu || 'Sin cultivo')}</b> <span class="small muted">(${n})</span>${ok ? `<div class="small">Dosis autorizada: ${esc(textoUsos(rr.usos)) || 'ver ficha'}</div>` : ok === false ? '<div class="small">No autorizado en este cultivo</div>' : ''}</div>`;
    });
    if (r.consumo.length) h += `<div class="small">Se gastan: ${r.consumo.map(x => `<b>${fmtCant(x.q, p.unidad)}</b>${x.t ? ' de ' + esc(x.t) : ''} (tiene ${fmtCant(x.ex, p.unidad)})`).join(' · ')}</div>`;
    if (r.graves.length) h += `<div class="aviso-norma"><b>⚠ Puedes estar incumpliendo la normativa</b><ul>${r.graves.map(g => `<li>${esc(g)}</li>`).join('')}</ul></div>`;
    if (r.notas.length) h += `<ul class="small muted notas">${r.notas.map(g => `<li>${esc(g)}</li>`).join('')}</ul>`;
    if (reg) h += `<a class="small" href="${urlFicha(reg.id)}" target="_blank" rel="noopener">Ver ficha oficial del Ministerio</a>`;
    return h;
  }

  // Guarda una labor (una por parcela; con varias comparten «lote»). La usan el formulario y el asistente de voz.
  function registrarLabor(v, ctx = {}) {
    const l = ctx.l || null, miembros = ctx.miembros || [], propio = ctx.propio || {};
    const fs = v.fincas.map(finca).filter(Boolean);
    const p = producto(v.prodAlm);
    const r = p ? revisarLabor(Object.assign({}, v, { _propio: propio })) : null;
    const ids = new Set(miembros.map(x => x.id));
    db.labores = db.labores.filter(x => !ids.has(x.id));
    db.movimientos = db.movimientos.filter(m => !ids.has(m.laborId));
    const lote = fs.length > 1 ? (l && l.lote) || uid() : null;
    const costes = repartir(v.coste, fs);
    if (v.tractorId || v.aperoId) { db.ultMaq = db.ultMaq || {}; db.ultMaq[v.tipo] = { equipoId: v.equipoId || '', tractorId: v.tractorId || '', aperoId: v.aperoId || '' }; }
    if (v.equipoId && !equipo(v.equipoId)) v.equipoId = '';
    if (v.tipo !== 'Tratamiento') delete v.aplicador; else if (v.aplicador) db.ultAplicador = v.aplicador;
    fs.forEach((f, i) => {
      const x = Object.assign({}, v, { id: i === 0 && l ? l.id : uid(), fincaId: f.id, coste: costes[i], superficie: fs.length > 1 ? null : v.superficie });
      delete x.fincas; delete x.prodAlm; delete x._propio; delete x.cultivoSiembra; delete x.variedadSiembra; if (lote) x.lote = lote; else delete x.lote;
      delete x.consumo; delete x.avisos; delete x.productoId;
      if (p) {
        x.productoId = p.id; x.producto = p.nombre; if (p.registro) x.registro = p.registro;
        const q = gastoEnParcela(p, f, v, fs);
        if (q != null) x.consumo = { productoId: p.id, cantidad: Math.round(q * 1000) / 1000, unidad: p.unidad, titular: titularDe(f.id) };
        const cu = cultivoDe(f.id, x.campana) || '';
        const pc = r && r.porCultivo.find(z => z.cu === cu);
        const av = [...(pc ? pc.r.graves : []), ...(r ? r.graves.filter(g => / solo tiene /.test(g) && g.startsWith(titularDe(f.id) || 'El titular')) : [])];
        if (av.length) x.avisos = av;
  }
      db.labores.push(x);
      if (x.coste > 0) db.movimientos.push({ id: uid(), laborId: x.id, tipo: 'gasto', fecha: x.fecha, fincaId: x.fincaId, campana: x.campana, categoria: CAT_LABOR[x.tipo] || 'Servicios contratados', concepto: x.tipo + (x.producto ? ' · ' + x.producto : ''), importe: x.coste, lote: lote || undefined });
      if (x.tipo === 'Siembra' && v.cultivoSiembra) {
        db.siembras = db.siembras.filter(s => !(s.fincaId === x.fincaId && s.campana === x.campana));
        db.siembras.push({ id: uid(), fincaId: x.fincaId, campana: x.campana, cultivo: v.cultivoSiembra, variedad: v.variedadSiembra || '' });
      } else if (x.tipo === 'Siembra' && !siembraDe(x.fincaId, x.campana)) {
        const cu = db.cultivos.find(c => (x.producto || '').toLowerCase().includes(c.toLowerCase()));
        if (cu) db.siembras.push({ id: uid(), fincaId: x.fincaId, campana: x.campana, cultivo: cu, variedad: '' });
  }
    });
      
  }

  function formLabor(id, pre = {}) {
    const l = id ? db.labores.find(x => x.id === id) : null;
    if (!db.fincas.length) { toast('Primero añade una finca'); return; }
    const miembros = l ? miembrosLote(db.labores, l) : [];
    const conDosis = v => CON_DOSIS.includes(v.tipo);
    const trat = v => v.tipo === 'Tratamiento';
    const una = v => (v.fincas || []).length <= 1;
    const hayMaq = () => db.maquinas.length > 0;
    const viejo = l && l.tipo === 'Tratamiento' && !l.productoId; // labor antigua sin producto del almacén
    // lo que ya gastaba esta labor vuelve a estar disponible al editarla
    const propio = {}; miembros.forEach(x => { if (x.consumo) propio[x.consumo.titular] = (propio[x.consumo.titular] || 0) + (+x.consumo.cantidad || 0); });
    const ini = l
      ? Object.assign({}, l, { fincas: miembros.map(x => x.fincaId), prodAlm: l.productoId || '', coste: miembros.some(x => x.coste) ? Math.round(sum(miembros, x => x.coste) * 100) / 100 : null })
      : Object.assign({ fecha: hoyISO(), tipo: 'Labrado', fincas: idsIniciales(pre), unidad: 'kg/ha', campana: db.campanaActiva, prodAlm: '' }, maqPorDefecto((pre.valores || {}).tipo || 'Labrado'), { aplicador: aplicadorPorDefecto() }, pre.valores || {});
    let maqTocada = !!l || !!(pre.valores && (pre.valores.tractorId || pre.valores.aperoId || pre.valores.equipoId));
    if (l && l.tipo === 'Siembra') { const s = siembraDe(l.fincaId, l.campana); if (s) { ini.cultivoSiembra = s.cultivo; ini.variedadSiembra = s.variedad || ''; } }
    let ultimoProd = ini.prodAlm;
    cargarFitos().then(d => { FITOS = d; if (refrescarInfo) refrescarInfo(); }).catch(() => {});
    let refrescarInfo = null;
    formulario({
      titulo: l ? (miembros.length > 1 ? `Editar labor (${miembros.length} parcelas)` : 'Editar labor') : 'Nueva labor',
      valores: ini,
      campos: [
        { k: 'fecha', label: 'Fecha', type: 'date', req: 1, half: 1 },
        { k: 'tipo', label: 'Labor', type: 'select', req: 1, half: 1, opts: () => TIPOS_LABOR.map(t => [t, t]) },
        { k: 'fincas', label: 'Fincas', type: 'fincas', req: 1, hint: 'Marca una o varias. Al marcar una finca agrupada se marcan todas sus parcelas.' },
        { k: 'prodAlm', label: 'Producto del almacén', sinOpcional: 1, type: 'select', dyn: 1, opts: v => opcionesAlmacen(v, l && l.productoId), show: v => !!TIPO_DE_LABOR[v.tipo] && !(viejo && !v.prodAlm && v.tipo === 'Tratamiento' && l.producto) },
        { k: 'producto', label: 'Producto, variedad o apero', placeholder: 'Ej.: cultivador, vertedera', show: v => !(TIPO_DE_LABOR[v.tipo] && v.prodAlm) && (v.tipo !== 'Tratamiento' || viejo) },
        { k: 'cultivoSiembra', label: 'Cultivo sembrado', type: 'select', half: 1, opts: () => [['', 'No cambiar'], ...db.cultivos.map(c => [c, c])], show: v => v.tipo === 'Siembra', hint: 'Se asigna a las parcelas en la campaña.' },
        { k: 'variedadSiembra', label: 'Variedad', half: 1, placeholder: 'Ej.: Marcopolo', show: v => v.tipo === 'Siembra' },
        { k: 'dosis', label: 'Dosis', type: 'num', half: 1, show: conDosis },
        { k: 'unidad', label: 'Unidad', type: 'select', half: 1, opts: () => UNIDADES.map(u => [u, u]), show: conDosis },
        { k: 'registro', label: 'Nº de registro', half: 1, show: v => trat(v) && !v.prodAlm, placeholder: 'Registro del producto' },
        { k: 'motivo', label: 'Plaga o mala hierba', half: 1, show: trat, placeholder: 'Ej.: avena loca' },
        { k: 'superficie', label: 'Superficie trabajada (ha)', type: 'num', placeholder: 'Vacío = toda la finca', show: una },
        { k: 'equipoId', label: 'Equipo de trabajo', type: 'select', opts: () => [['', 'Sin equipo (elige las máquinas)'], ...db.equipos.map(e => [e.id, e.nombre + (e.rendimiento ? ` · ${fmtN(e.rendimiento)} ha/h` : '')])], show: () => db.equipos.length > 0, hint: 'Con equipo, la app calcula el coste de la maquinaria por hectárea.' },
        { k: 'tractorId', label: 'Tractor o cosechadora', type: 'select', half: 1, opts: () => optsMaq(true, l && l.tractorId), show: hayMaq },
        { k: 'aperoId', label: 'Apero', type: 'select', half: 1, opts: () => optsMaq(false, l && l.aperoId), show: hayMaq },
        { k: 'aplicador', label: 'Aplicador', type: 'select', opts: () => optsAplicador(l && l.aplicador), show: trat, hint: 'Quien hace el tratamiento, con su carné (ROPO). Sale en el cuaderno oficial.' },
        { k: 'coste', label: 'Coste total (€)', type: 'num', miles: 1, hint: 'Se apunta también como gasto. Con varias fincas se reparte según la superficie de cada una.' },
        campoCampana,
        { k: 'notas', label: 'Notas', type: 'textarea', placeholder: 'Condiciones del tiempo, quién la hizo…' }
      ],
      extra: '<div id="fitoInfo" class="fito-info" hidden></div>',
      alCambiar(v, form, origen) {
        refrescarInfo = () => alCambiarReal(v, form);
        const ponMaq = d => { ['equipoId', 'tractorId', 'aperoId'].forEach(k => { if (form.elements[k] && d[k] !== undefined) { form.elements[k].value = d[k]; v[k] = d[k]; } }); };
        if (origen === 'tractorId' || origen === 'aperoId' || origen === 'equipoId') maqTocada = true;
        if (origen === 'equipoId' && equipo(v.equipoId)) { const e = equipo(v.equipoId); ponMaq({ tractorId: e.tractorId || '', aperoId: e.aperoId || '' }); }
        if ((origen === 'tractorId' || origen === 'aperoId') && v.equipoId) { const e = equipo(v.equipoId); if (!e || (e.tractorId || '') !== v.tractorId || (e.aperoId || '') !== v.aperoId) ponMaq({ equipoId: '' }); }
        if (origen === 'tipo' && !maqTocada && form.elements.tractorId) ponMaq(maqPorDefecto(v.tipo));
        const hEq = form.querySelector('.field[data-k="equipoId"] .hint');
        if (hEq) {
          const e = equipo(v.equipoId); const ce = e ? costeEquipo(e) : null;
          const fs = (v.fincas || []).map(finca).filter(Boolean); const sup = parseNum(v.superficie);
          const ha = fs.length === 1 && sup > 0 ? sup : sum(fs, f => +f.superficie || 0);
          let txt = !e ? 'Con equipo, la app suma a la finca el coste de la maquinaria.' : ce.porHa == null ? `<span class="neg">Faltan datos para la amortización: ${esc(ce.faltan.join(', ') || 'rendimiento')}.</span>` : `Amortización: ${eur(ce.porHa, 2)}/ha${ha ? ` × ${fmtN(ha)} ha = <b>${eur(ce.porHa * ha, 2)}</b>` : ''}.`;
          if (e && +e.gasoleo > 0 && ha) {
            const q = +e.gasoleo * ha; const g = gasoleoAlm();
            const yaGastado = l ? sum(miembros, litrosLabor) : 0;
            const pr = l && yaGastado ? (sum(miembros, x => costeGasLabor(x) || 0) / yaGastado || precioGasoleo(q)) : precioGasoleo(q);
            txt += `<br>Gasóleo: ${fmtN(e.gasoleo)} L/ha × ${fmtN(ha)} ha = <b>${fmtN(q, 0)} L</b>${pr != null ? ` × ${fmtN(pr, 3)} €/L = <b>${eur(q * pr, 2)}</b>` : ' <span class="neg">(sin precio: falta el albarán del gasóleo con su precio)</span>'}${g ? ', del gasóleo común (primero el más antiguo)' : ''}.`;
            if (g) { const ex = existencias(g.id, null) + yaGastado; if (ex < q) txt += ` <span class="neg">Solo quedan ${fmtN(ex, 0)} L entre todos.</span>`; }
          }
          if (e) txt += '<br>Se suma a los gastos de la finca.';
          hEq.innerHTML = txt;
        }
        if (origen === 'prodAlm' || (v.prodAlm && v.prodAlm !== ultimoProd)) {
          ultimoProd = v.prodAlm;
          const p = producto(v.prodAlm); const reg = regProducto(p);
          if (reg && !(parseNum(v.dosis) > 0)) {
            const u0 = Object.values(reg.u || {}).flat().map(x => UNI_REG[x[3]]).find(Boolean);
            if (u0 && form.elements.unidad) { form.elements.unidad.value = u0; v.unidad = u0; }
          } else if (p && p.tipo !== 'Fitosanitario' && !(parseNum(v.dosis) > 0) && form.elements.unidad) { const u0 = p.unidad === 'L' ? 'L/ha' : 'kg/ha'; form.elements.unidad.value = u0; v.unidad = u0; }
        }
        alCambiarReal(v, form);
      },
      antesDeGuardar(v) {
        if (v.tipo === 'Tratamiento' && !v.prodAlm && !viejo) { toast('Elige el fitosanitario del almacén. Si no aparece, sube primero su albarán.'); return Promise.resolve(false); }
        if (!v.prodAlm) return Promise.resolve(true);
        const r = revisarLabor(Object.assign({}, v, { _propio: propio }));
        if (!r.graves.length) return Promise.resolve(true);
        return confirmar('⚠ Revisa antes de guardar', `<ul class="avisos-lista">${r.graves.map(g => `<li>${esc(g)}</li>`).join('')}</ul>Si lo guardas, la labor quedará marcada para revisar en el cuaderno.`, 'Guardar igualmente').then(x => x.ok);
      },
      guardar(v) { registrarLabor(v, { l, miembros, propio }); },
      textoBorrar: miembros.length > 1 ? `Se borrará esta labor en las ${miembros.length} parcelas y, si tenía coste, también sus gastos. Lo gastado del almacén vuelve al inventario.` : 'Se borrará la labor y, si tenía coste, también su gasto. Lo gastado del almacén vuelve al inventario.',
      borrar: l && (() => { const ids = new Set(miembros.map(x => x.id)); db.labores = db.labores.filter(x => !ids.has(x.id)); db.movimientos = db.movimientos.filter(m => !ids.has(m.laborId)); })
    });
    function alCambiarReal(v, form) {
      const box = form.querySelector('#fitoInfo'); if (!box) return;
      const anc = form.querySelector('[data-pair="dosis"]') || form.querySelector('.field[data-k="prodAlm"]');
      if (anc && box.previousElementSibling !== anc) anc.after(box);
      const html = TIPO_DE_LABOR[v.tipo] ? infoLaborHTML(Object.assign({}, v, { dosis: parseNum(v.dosis), superficie: parseNum(v.superficie), _propio: propio })) : '';
      box.hidden = !html; box.innerHTML = html;
    }
  }

  function registrarMov(v, ctx = {}) {
    const m = ctx.m || null, miembros = ctx.miembros || [];
    const fs = (v.fincas || []).map(finca).filter(Boolean);
    db.movimientos = db.movimientos.filter(x => !miembros.includes(x));
    const base = Object.assign({}, v); delete base.fincas;
    if (!fs.length) { db.movimientos.push(Object.assign({ id: m ? m.id : uid() }, base, { fincaId: '' })); return; }
    const lote = fs.length > 1 ? (m && m.lote) || uid() : null;
    const imp = repartir(v.importe, fs), kg = v.kilos ? repartir(v.kilos, fs) : fs.map(() => null);
    fs.forEach((f, i) => { const x = Object.assign({ id: i === 0 && m ? m.id : uid() }, base, { fincaId: f.id, importe: imp[i], kilos: kg[i] }); if (lote) x.lote = lote; db.movimientos.push(x); });
  }
  function formMov(id, pre = {}) {
    const m = id ? db.movimientos.find(x => x.id === id) : null;
    if (m && m.laborId && db.labores.some(l => l.id === m.laborId)) { formLabor(m.laborId); return; }
    const miembros = m ? miembrosLote(db.movimientos, m) : [];
    const tipo = pre.tipo || 'gasto';
    const venta = v => v.categoria === 'Venta de grano' || v.categoria === 'Venta de paja';
    const ini = m
      ? Object.assign({}, m, { fincas: miembros.map(x => x.fincaId).filter(id2 => finca(id2)), importe: Math.round(sum(miembros, x => x.importe) * 100) / 100, kilos: miembros.some(x => x.kilos) ? sum(miembros, x => x.kilos) : null })
      : Object.assign({ tipo, fecha: hoyISO(), categoria: CAT[tipo][0], fincas: pre.grupo || pre.finca ? idsIniciales(pre) : [], campana: db.campanaActiva }, pre.valores || {});
    formulario({
      titulo: m ? (miembros.length > 1 ? `Editar apunte (${miembros.length} parcelas)` : 'Editar apunte') : (tipo === 'ingreso' ? 'Nuevo ingreso' : 'Nuevo gasto'),
      valores: ini,
      campos: [
        { k: 'tipo', label: 'Tipo', type: 'select', req: 1, half: 1, opts: () => [['gasto', 'Gasto'], ['ingreso', 'Ingreso']] },
        { k: 'fecha', label: 'Fecha', type: 'date', req: 1, half: 1 },
        { k: 'categoria', label: 'Categoría', type: 'select', req: 1, dyn: 1, opts: v => CAT[v.tipo || 'gasto'].map(x => [x, x]) },
        { k: 'concepto', label: 'Concepto', placeholder: 'Ej.: 2.500 kg semilla certificada' },
        { k: 'kilos', label: 'Kilos vendidos', type: 'num', miles: 1, half: 1, show: venta },
        { k: 'precio', label: 'Precio (€/kg)', type: 'num', half: 1, show: venta, placeholder: 'Ej.: 0,23' },
        { k: 'importe', label: 'Importe total (€)', type: 'num', miles: 1, req: 1 },
        { k: 'fincas', label: 'Fincas', type: 'fincas', hint: 'Sin marcar ninguna es un gasto general de la explotación (seguro, gasóleo, PAC…). Con varias, el importe se reparte según la superficie.' },
        campoCampana,
        { k: 'notas', label: 'Notas', type: 'textarea' }
      ],
      alEscribir(k, v, set) {
        if (k !== 'kilos' && k !== 'precio') return;
        const kg = parseNum(v.kilos, true), p = parseNum(v.precio);
        if (kg > 0 && p > 0) set('importe', numStr(Math.round(kg * p * 100) / 100));
      },
      guardar(v) {
        if (v.importe <= 0) return 'El importe tiene que ser mayor que cero';
        registrarMov(v, { m, miembros });
      },
      borrar: m && (() => { db.movimientos = db.movimientos.filter(x => !miembros.includes(x)); })
    });
  }

  function registrarCosecha(v, ctx = {}) {
    const x = ctx.x || null, miembros = ctx.miembros || [];
    const fs = v.fincas.map(finca).filter(Boolean);
    db.cosechas = db.cosechas.filter(c => !miembros.includes(c));
    const lote = fs.length > 1 ? (x && x.lote) || uid() : null;
    const kgs = repartir(v.kg, fs);
    const base = Object.assign({}, v); delete base.fincas;
    fs.forEach((f, i) => { const c = Object.assign({ id: i === 0 && x ? x.id : uid() }, base, { fincaId: f.id, kg: kgs[i] }); if (lote) c.lote = lote; db.cosechas.push(c); });
  }
  function formCosecha(id, pre = {}) {
    const x = id ? db.cosechas.find(c => c.id === id) : null;
    if (!db.fincas.length) { toast('Primero añade una finca'); return; }
    const miembros = x ? miembrosLote(db.cosechas, x) : [];
    const ini = x
      ? Object.assign({}, x, { fincas: miembros.map(y => y.fincaId), kg: Math.round(sum(miembros, y => y.kg)) })
      : Object.assign({ fecha: hoyISO(), fincas: idsIniciales(pre), campana: db.campanaActiva, maquinaId: maqPorDefecto('Cosecha').tractorId }, pre.valores || {});
    formulario({
      titulo: x ? (miembros.length > 1 ? `Editar cosecha (${miembros.length} parcelas)` : 'Editar cosecha') : 'Apuntar cosecha',
      valores: ini,
      campos: [
        { k: 'fincas', label: 'Fincas', type: 'fincas', req: 1, hint: 'Si eliges varias, los kilos se reparten según la superficie de cada una.' },
        { k: 'fecha', label: 'Fecha', type: 'date', req: 1, half: 1 },
        { k: 'kg', label: 'Kilos', type: 'num', miles: 1, req: 1, half: 1, placeholder: 'Ej.: 12.400' },
        { k: 'humedad', label: 'Humedad (%)', type: 'num', placeholder: 'Ej.: 11,5' },
        { k: 'maquinaId', label: 'Cosechadora', type: 'select', opts: () => optsMaq(true, x && x.maquinaId, ['Cosechadora']), show: () => db.maquinas.some(m => m.tipo === 'Cosechadora') },
        campoCampana,
        { k: 'notas', label: 'Destino o notas', type: 'textarea', placeholder: 'Cooperativa, almacén, nº de ticket de báscula…' }
      ],
      guardar(v) {
        if (v.kg <= 0) return 'Los kilos tienen que ser mayores que cero';
        if (v.maquinaId) { db.ultMaq = db.ultMaq || {}; db.ultMaq.Cosecha = Object.assign({}, db.ultMaq.Cosecha, { tractorId: v.maquinaId }); }
        registrarCosecha(v, { x, miembros });
      },
      borrar: x && (() => { db.cosechas = db.cosechas.filter(c => !miembros.includes(c)); })
    });
  }

  // ---------- Agrupar parcelas en una finca ----------
  function hojaAgrupar(o) {
    const base = o.grupo ? fincasDeGrupo(o.grupo) : [finca(o.id)];
    const actual = o.grupo || grupoDe(base[0]) || '';
    const nb = nombreBase(base[0].nombre);
    const sinNombre = !nb || norm(nb) === norm(base[0].municipio) || norm(nb) === 'parcela'; // nombre puesto por la app («Lerma 632-599»)
    const sug = actual || (sinNombre ? '' : nb);
    const vec = lindantes(base);
    const mismo = sug ? db.fincas.filter(x => !base.includes(x) && !vec.includes(x) && norm(nombreBase(x.nombre)) === norm(nombreBase(sug))) : [];
    const sinGeo = base.some(f => !f.geo || !f.geo.c.length);
    const fila = (f, chk) => `<label class="pk-f"><input type="checkbox" value="${f.id}"${chk ? ' checked' : ''}><span class="pk-n">${esc(f.nombre)}${grupoDe(f) && grupoDe(f) !== actual ? ` <span class="small muted">(ahora en ${esc(grupoDe(f))})</span>` : ''}</span><span class="muted small">${fmtN(f.superficie)} ha</span></label>`;
    const caja = (t, fs, chk, vacioTxt) => `<div class="field"><label>${t}</label>${fs.length ? `<div class="picker"><div class="pk-list">${fs.sort(byNombre).map(f => fila(f, chk)).join('')}</div></div>` : `<p class="note" style="margin:0">${vacioTxt}</p>`}</div>`;
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera(o.grupo ? 'Finca agrupada' : 'Agrupar parcelas')}<form class="body" id="frmAg" novalidate>
      <div class="field"><label for="agNombre">Nombre de la finca agrupada</label><input id="agNombre" value="${esc(sug)}" placeholder="Ej.: Sarrana" autocomplete="off" list="dl_grupos"><datalist id="dl_grupos">${[...listaGrupos().keys()].map(g => `<option value="${esc(g)}">`).join('')}</datalist></div>
      ${caja(o.grupo ? 'Parcelas de la finca' : 'Esta parcela', base, true, '')}
      ${caja('Lindan con ellas', vec, false, sinGeo ? 'Para saber qué parcelas lindan, la app necesita sus contornos de SIGPAC (polígono, parcela, provincia y municipio).' : 'Ninguna de tus parcelas linda con estas.')}
      ${mismo.length ? caja('Con el mismo nombre', mismo, false, '') : ''}
      <div class="field"><label for="agBuscar">Añadir otra parcela</label><input id="agBuscar" type="search" placeholder="Buscar por nombre o polígono-parcela…" autocomplete="off"><div id="agRes" class="picker" hidden><div class="pk-list"></div></div><div id="agExtra"></div></div>
      </form><footer><button class="btn" type="button" data-close>Cancelar</button><button class="btn primary" type="submit" form="frmAg">Guardar</button></footer></div>`, (el, d) => {
      const form = el.querySelector('form');
      const mostrados = () => new Set([...form.querySelectorAll('.pk-f input')].map(i => i.value));
      const res = el.querySelector('#agRes'), resL = res.querySelector('.pk-list'), extra = el.querySelector('#agExtra');
      el.querySelector('#agBuscar').addEventListener('input', e => {
        const q = norm(e.target.value); const ya = mostrados();
        const hits = q ? db.fincas.filter(f => !ya.has(f.id) && q.split(' ').every(w => norm([f.nombre, f.grupo, f.municipio, (f.poligono || '') + ' ' + (f.parcela || '')].join(' ')).includes(w))).slice(0, 25) : [];
        res.hidden = !hits.length; resL.innerHTML = hits.map(f => fila(f, false)).join('');
      });
      resL.addEventListener('change', e => {
        const lab = e.target.closest('.pk-f'); if (!lab || !e.target.checked) return;
        if (!extra.querySelector('.picker')) extra.innerHTML = '<div class="picker" style="margin-top:8px"><div class="pk-list"></div></div>';
        extra.querySelector('.pk-list').appendChild(lab);
        if (!resL.children.length) res.hidden = true;
      });
      form.addEventListener('submit', e => {
        e.preventDefault();
        const nombre = el.querySelector('#agNombre').value.trim();
        if (!nombre) { toast('Ponle un nombre a la finca agrupada'); return; }
        const marcadas = new Set([...form.querySelectorAll('.pk-list .pk-f input:checked')].filter(i => !i.closest('#agRes')).map(i => i.value));
        if (marcadas.size < 2) { toast('Marca al menos 2 parcelas para agruparlas'); return; }
        base.forEach(f => { if (!marcadas.has(f.id)) f.grupo = ''; });
        db.fincas.forEach(f => { if (marcadas.has(f.id)) f.grupo = nombre; });
        guardar(); volverA(d - 1);
        if (ui.grupo === o.grupo && o.grupo) ui.grupo = nombre;
        render(); toast(`«${nombre}»: ${marcadas.size} parcelas`);
      });
    });
  }

  // ---------- Campañas y ajustes ----------
  function hojaCampanas() {
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera('Campaña')}<div class="body">
      <p class="note">Elige la campaña con la que trabajas. Todo lo que apuntes irá a ella por defecto.</p>
      <div class="list">${listaCampanas().map(c => { const vacia = c !== db.campanaActiva && !['siembras', 'labores', 'movimientos', 'cosechas'].some(k => db[k].some(x => x.campana === c));
        return `<div class="row camp-row"><button class="row" data-camp="${c}"><div class="main"><div class="t">${c}</div></div>${c === db.campanaActiva ? '<span class="tag">Activa</span>' : ''}</button>${vacia ? `<button class="btn sm" data-quitarc="${c}">Quitar</button>` : ''}</div>`; }).join('')}</div>
      <p class="small muted" style="margin:0">Solo se pueden quitar las campañas que no tienen nada apuntado.</p>
      <button class="btn block" data-nueva>${ico.plus} Añadir campaña ${siguienteCampana(listaCampanas()[0])}</button></div></div>`, (el, d) => {
      el.addEventListener('click', e => {
        const q = e.target.closest('[data-quitarc]');
        if (q) { db.campanas = db.campanas.filter(c => c !== q.dataset.quitarc); guardar(); volverA(d - 1); render(); toast('Campaña ' + q.dataset.quitarc + ' quitada'); return; }
        const b = e.target.closest('[data-camp]');
        if (b) { db.campanaActiva = b.dataset.camp; guardar(false); volverA(d - 1); render(); return; }
        if (e.target.closest('[data-nueva]')) {
          const n = siguienteCampana(listaCampanas()[0]); db.campanas.push(n); db.campanaActiva = n; guardar(); volverA(d - 1); render(); toast('Campaña ' + n + ' creada');
        }
      });
    });
  }

  function hojaAjustes() {
    const contenido = () => {
      const ult = db.ultimaCopia ? new Date(db.ultimaCopia).toLocaleString('es-ES', { dateStyle: 'medium', timeStyle: 'short' }) : 'nunca';
      const usados = new Set(db.siembras.map(s => s.cultivo));
      return `<h2 class="sec" style="margin-top:0">Copia de seguridad</h2>
        <div class="card pad" style="display:flex;flex-direction:column;gap:10px">
          <p style="margin:0">Los datos se guardan solo en este teléfono. Haz una copia de vez en cuando y guárdala en Google Drive, en tu correo o en WhatsApp. Con ese archivo puedes recuperar todo en otro móvil.</p>
          <p class="small muted" style="margin:0">Última copia: <b>${ult}</b></p>
          <button class="btn primary block" data-a="compartir">Compartir copia (Drive, WhatsApp, correo…)</button>
          <button class="btn block" data-a="descargar">Descargar copia en el teléfono</button>
          <button class="btn block" data-a="importar">Importar una copia</button>
        </div>
        <h2 class="sec">Fincas en Excel</h2>
        <div class="card pad" style="display:flex;flex-direction:column;gap:10px">
          <p style="margin:0">Trae tus fincas de una hoja de cálculo: nombre, superficie, polígono, parcela, recinto, cultivo, notas… Vale el listado de la PAC.</p>
          <button class="btn primary block" data-a="importarExcel">Importar Excel de fincas</button>
          <button class="btn block" data-a="plantilla">Descargar plantilla de ejemplo</button>
          ${db.fincas.length ? '<button class="btn block" data-a="exportarExcel">Exportar mis fincas a Excel</button>' : ''}
        </div>
        <h2 class="sec">Cultivos</h2>
        <div class="card pad" style="display:flex;flex-direction:column;gap:10px">
          <div style="display:flex;flex-wrap:wrap;gap:6px">${db.cultivos.map(c => `<span class="chip" style="--h:${hue(c)}">${esc(c)}${usados.has(c) ? '' : ` <button data-quitar="${esc(c)}" aria-label="Quitar ${esc(c)}" style="border:0;background:none;padding:0 0 0 4px;font-weight:700;color:inherit">×</button>`}</span>`).join('')}</div>
          <div style="display:flex;gap:8px"><div class="field" style="flex:1"><input id="nuevoCultivo" placeholder="Añadir cultivo (ej.: Avena, Garbanzo)" aria-label="Nuevo cultivo"></div><button class="btn" data-a="addCultivo">Añadir</button></div>
        </div>
        <h2 class="sec">Datos</h2>
        <div class="card pad" style="display:flex;flex-direction:column;gap:10px">
          <p class="small muted" style="margin:0">${db.fincas.length} fincas · ${db.labores.length} labores · ${db.movimientos.length} apuntes · ${db.cosechas.length} cosechas</p>
          ${!db.fincas.length ? '<button class="btn block" data-a="ejemplo">Cargar datos de ejemplo</button>' : ''}
          <button class="btn danger block" data-a="borrarTodo">Borrar todos los datos</button>
        </div>
        <p class="note">Cuaderno de Campo ${VERSION}. Funciona sin cobertura.</p>`;
    };
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera('Ajustes')}<div class="body" id="ajBody">${contenido()}</div></div>`, (el, d) => {
      const body = el.querySelector('#ajBody');
      const repintar = () => { body.innerHTML = contenido(); };
      el.addEventListener('click', async e => {
        const q = e.target.closest('[data-quitar]');
        if (q) { db.cultivos = db.cultivos.filter(c => c !== q.dataset.quitar); guardar(); repintar(); return; }
        const b = e.target.closest('[data-a]'); if (!b) return;
        const a = b.dataset.a;
        if (a === 'compartir') { await exportar(true); repintar(); }
        else if (a === 'descargar') { await exportar(false); repintar(); }
        else if (a === 'importar') $('#fileIn').click();
        else if (a === 'importarExcel') { cargarXLSX().catch(() => {}); $('#xlsIn').click(); }
        else if (a === 'plantilla') excelFincas(true);
        else if (a === 'exportarExcel') excelFincas(false);
        else if (a === 'addCultivo') {
          const v = el.querySelector('#nuevoCultivo').value.trim();
          if (!v) return;
          if (db.cultivos.some(c => c.toLowerCase() === v.toLowerCase())) { toast('Ese cultivo ya está en la lista'); return; }
          db.cultivos.push(v.charAt(0).toUpperCase() + v.slice(1)); guardar(); repintar();
        }
        else if (a === 'ejemplo') { cargarEjemplo(); volverA(d - 1); }
        else if (a === 'borrarTodo') {
          const r = await confirmar('¿Borrar todos los datos?', 'Se borrarán todas las fincas, labores, cuentas y cosechas de este teléfono. Si no tienes copia de seguridad, no podrás recuperarlas.', 'Borrar todo');
          if (!r.ok) return;
          db = nuevoDB(); guardar(false); archivosVaciar().catch(() => {}); ui.fincaId = null; ui.grupo = null; volverA(0); render(); toast('Datos borrados');
        }
      });
    });
  }

  // ---------- Archivos: compartir o descargar ----------
  async function guardarArchivo(blob, nombre, compartir, titulo) {
    let file;
    try { file = new File([blob], nombre, { type: blob.type }); } catch (e) { file = blob; }
    if (compartir && navigator.canShare && file instanceof File && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: titulo || nombre }); return true; }
      catch (e) { if (e && e.name === 'AbortError') return false; }
    }
    try {
      const url = URL.createObjectURL(file);
      const a = document.createElement('a'); a.href = url; a.download = nombre; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      if (compartir) toast('Este navegador no permite compartir; el archivo se ha descargado');
      return true;
    } catch (e) { toast('No se ha podido guardar el archivo en este navegador'); return false; }
  }

  // ---------- Excel de fincas ----------
  let xlsxP = null;
  function cargarXLSX() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    return xlsxP || (xlsxP = new Promise((res, rej) => {
      const s = document.createElement('script'); s.src = XLSX_SRC;
      s.onload = () => res(window.XLSX); s.onerror = () => { xlsxP = null; rej(new Error('xlsx')); };
      document.head.appendChild(s);
    }));
  }
  // Nombres de columna reconocidos (sin tildes ni mayúsculas). El orden importa.
  const COLS = [
    ['codMunicipio', ['cod municipio', 'codigo municipio', 'cod mun', 'codigo mun', 'municipio codigo', 'municipio cod', 'n municipio', 'num municipio', 'municipio sigpac']],
    ['provincia', ['provincia', 'cod provincia', 'codigo provincia', 'prov']],
    ['municipio', ['municipio', 'termino municipal', 'termino', 'localidad', 'pueblo', 'mun']],
    ['agregado', ['agregado', 'agr']],
    ['zona', ['zona']],
    ['poligono', ['poligono', 'pol']],
    ['otras', ['otras parcelas', 'otras referencias', 'mas parcelas']],
    ['parcela', ['parcela', 'parc', 'par']],
    ['recinto', ['recinto', 'recintos', 'rec']],
    ['supDeclarable', ['superficie declarable pac', 'superficie declarable', 'sup declarable', 'superficie pac', 'declarable']],
    ['superficie', ['superficie', 'superficie ha', 'sup', 'sup ha', 'hectareas', 'ha', 'has', 'superficie declarada', 'superficie sigpac', 'superficie cultivada', 'sup declarada']],
    ['cultivo', ['cultivo', 'producto', 'cultivo declarado', 'siembra', 'especie', 'cultivo campana']],
    ['variedad', ['variedad']],
    ['laboreo', ['laboreo', 'tipo laboreo', 'tipo de laboreo']],
    ['estiercol', ['basura', 'estiercol', 'abonado organico', 'estercolado']],
    ['campana', ['campana', 'ano', 'cosecha', 'ano campana']],
    ['notas', ['notas', 'observaciones', 'obs', 'comentarios', 'nota']],
    ['titular', ['titular', 'propietario', 'titular pac', 'dueno']],
    ['grupo', ['finca agrupada', 'agrupacion', 'grupo de fincas']],
    ['nombre', ['nombre', 'finca', 'paraje', 'denominacion', 'nombre finca', 'nombre de la finca', 'alias']]
  ];
  const ETIQ = { nombre: 'Nombre', superficie: 'Superficie', supDeclarable: 'Superficie declarable', municipio: 'Municipio', codMunicipio: 'Cód. municipio', provincia: 'Provincia', agregado: 'Agregado', zona: 'Zona', poligono: 'Polígono', parcela: 'Parcela', recinto: 'Recinto', otras: 'Otras parcelas', cultivo: 'Cultivo', variedad: 'Variedad', laboreo: 'Laboreo', estiercol: 'Basura (estiércol)', campana: 'Campaña', notas: 'Notas', titular: 'Titular', grupo: 'Finca agrupada' };
  function columna(h) {
    const n = norm(h); if (!n) return null;
    for (const [k, syn] of COLS) if (syn.includes(n)) return k;
    for (const [k, syn] of COLS) if (syn.some(s => s.length > 3 && n.startsWith(s + ' '))) return k;
    return null;
  }
  function campanaTxt(v) {
    const s = String(v == null ? '' : v).trim(); if (!s) return '';
    let m = s.match(/^(\d{4})\s*[\/\-]\s*(\d{2}|\d{4})$/);
    if (m) return m[1] + '/' + String((+m[1] + 1) % 100).padStart(2, '0');
    m = s.match(/^(\d{2})\s*\/\s*(\d{2})$/);
    if (m && (+m[1] + 1) % 100 === +m[2]) { const a = 2000 + +m[1]; return a + '/' + m[2]; }
    m = s.match(/^(\d{4})$/);
    if (m) { const a = +m[1] - 1; return a + '/' + String((a + 1) % 100).padStart(2, '0'); } // año de cosecha
    return '';
  }
  const campanaAnterior = c => { const a = parseInt(c, 10) - 1; return a + '/' + String((a + 1) % 100).padStart(2, '0'); };
  const tituloCaso = s => (s && s === s.toUpperCase() && /[A-ZÁÉÍÓÚÑ]/.test(s))
    ? s.toLowerCase().replace(/(^|[\s(\-])(\S)/g, (m0, a, b) => a + b.toUpperCase()).replace(/ (De|La|Del|Y|El|Los|Las) /g, x => x.toLowerCase()) : s;
  function cultivoNormal(v) {
    const n = norm(v);
    if (!n || SIN_CULTIVO.includes(n) || /^\d+(\.\d+)?$/.test(String(v).trim())) return '';
    if (CULT_ALIAS[n]) return CULT_ALIAS[n];
    const c = db.cultivos.find(x => norm(x) === n); const t = String(v).trim();
    return c || (t.charAt(0).toUpperCase() + t.slice(1).toLowerCase());
  }

  // Lee la hoja: una fila por recinto o por finca. Admite varias campañas en columnas (CULTIVO/LABOREO/BASURA bajo «2025/26»…).
  function analizarHoja(XLSX, wb) {
    const nombreHoja = wb.SheetNames.find(n => norm(n) !== 'instrucciones') || wb.SheetNames[0];
    const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombreHoja], { header: 1, raw: true, defval: '' });
    let cab = -1, cols = {};
    for (let i = 0; i < Math.min(15, filas.length); i++) {
      const m = {}; filas[i].forEach((h, j) => { const k = columna(h); if (k) (m[k] = m[k] || []).push(j); });
      if (Object.keys(m).length >= 2) { cab = i; cols = m; break; }
    }
    if (cab < 0) return { error: 'No encuentro la fila de títulos. La primera fila debe tener títulos como Nombre, Superficie, Polígono, Parcela…' };
    const datos = filas.slice(cab + 1).filter(f => f.some(v => String(v).trim()));
    const esNum = j => { const vs = datos.map(f => f[j]).filter(v => v !== '' && v != null).slice(0, 40); return vs.length > 0 && vs.every(v => typeof v === 'number' || /^\d+$/.test(String(v).trim())); };
    if (cols.municipio && !cols.codMunicipio) {
      const num = cols.municipio.find(esNum);
      if (num != null) { cols.codMunicipio = [num]; cols.municipio = cols.municipio.filter(j => j !== num); if (!cols.municipio.length) delete cols.municipio; }
    }
    const uno = k => cols[k] ? cols[k][0] : -1;
    const arriba = cab > 0 ? filas[cab - 1] : [];
    const campDe = j => { for (const dd of [0, 1, -1, 2, -2]) { const c = campanaTxt(arriba[j + dd]); if (c) return c; } return ''; };
    let bloques = [];
    if ((cols.cultivo || []).length > 1) {
      bloques = cols.cultivo.map(j => ({ campana: campDe(j), cultivo: j, laboreo: -1, estiercol: -1, variedad: -1, campCol: -1 }));
      for (const k of ['laboreo', 'estiercol', 'variedad']) (cols[k] || []).forEach(j => { const b = bloques.filter(x => x.cultivo < j).pop(); if (b && b[k] < 0) b[k] = j; });
      bloques = bloques.filter(b => b.campana);
    }
    if (!bloques.length) bloques = [{ campana: '', cultivo: uno('cultivo'), laboreo: uno('laboreo'), estiercol: uno('estiercol'), variedad: uno('variedad'), campCol: uno('campana') }];
    const usadas = new Set(Object.values(cols).flat());
    const reconocidas = [...new Set(Object.keys(cols).map(k => ETIQ[k]))];
    const ignoradas = filas[cab].map((h, j) => [String(h).trim(), j]).filter(([h, j]) => h && !usadas.has(j)).map(([h]) => h);
    const supCol = cols.supDeclarable ? uno('supDeclarable') : uno('superficie');
    const get = (fila, j) => j >= 0 ? fila[j] : '';
    const txt = (fila, j) => { const v = get(fila, j); return String(v == null ? '' : v).trim(); };
    const ent = (fila, j) => { const v = get(fila, j); if (v === '' || v == null) return null; const n = typeof v === 'number' ? v : parseNum(v, false); return Number.isFinite(n) ? Math.round(n) : null; };
    const registros = []; let vacias = 0;
    for (const fila of datos) {
      const pr = provinciaDe(get(fila, uno('provincia')));
      let mu = ent(fila, uno('codMunicipio')); let muNombre = txt(fila, uno('municipio'));
      if (mu == null && /^\d+$/.test(muNombre)) { mu = +muNombre; muNombre = ''; }
      const m2 = muNombre.match(/^(\d+)\s*[-–]\s*(.+)$/); if (m2) { if (mu == null) mu = +m2[1]; muNombre = m2[2]; } // «197-LERMA»
      const po = ent(fila, uno('poligono')), pa = ent(fila, uno('parcela'));
      const nombre = txt(fila, uno('nombre'));
      if (!nombre && !(po && pa)) { vacias++; continue; }
      const supV = get(fila, supCol); const sup = typeof supV === 'number' ? supV : parseNum(supV, false);
      const camp = {};
      for (const b of bloques) {
        const c = b.campana || campanaTxt(get(fila, b.campCol)) || '';
        const cu = cultivoNormal(txt(fila, b.cultivo)); const lab = txt(fila, b.laboreo).toUpperCase(); const est = txt(fila, b.estiercol);
        if (cu || lab || est) camp[c] = { cultivo: cu, laboreo: lab, estiercol: !!est && !/^(no|n|0)$/i.test(est), variedad: txt(fila, b.variedad) };
      }
      registros.push({ nombre, titular: tituloCaso(txt(fila, uno('titular'))), grupo: txt(fila, uno('grupo')), municipio: tituloCaso(muNombre), pr: pr || '', mu: mu || '',
        ag: ent(fila, uno('agregado')) || 0, zo: ent(fila, uno('zona')) || 0, po, pa, recs: listaRecintos(get(fila, uno('recinto'))),
        sup: Number.isFinite(sup) && sup > 0 ? sup : 0, otras: txt(fila, uno('otras')), notas: txt(fila, uno('notas')), camp });
    }
    const campanas = [...new Set(bloques.map(b => b.campana).filter(Boolean))].sort();
    return { hoja: nombreHoja, registros, reconocidas, ignoradas, vacias, filas: datos.length, campanas, conProvincia: !!cols.provincia };
  }

  // Convierte las filas en parcelas (una por parcela SIGPAC, nombre, titular y cultivos) y les pone su finca agrupada.
  function construirPlan(r, o) {
    const valeCamp = c => !c || !o.multi || o.campanas.has(c);
    const partes = new Map();
    for (const x of r.registros) {
      const pr = x.pr || (o.provPorPueblo || {})[x.mu || norm(x.municipio)] || '';
      const camp = {}; for (const [c, d] of Object.entries(x.camp)) if (valeCamp(c)) camp[c || db.campanaActiva] = d;
      const firma = Object.keys(camp).sort().map(c => `${c}:${camp[c].cultivo}:${camp[c].laboreo}:${camp[c].estiercol ? 1 : 0}`).join(',');
      const key = x.po && x.pa ? [pr, x.mu, x.ag, x.zo, x.po, x.pa, norm(x.nombre), norm(x.titular), firma].join('|') : 'n|' + norm(x.nombre) + '|' + norm(x.titular);
      let p = partes.get(key);
      if (!p) { p = { nombre: x.nombre, titular: x.titular, grupoExcel: x.grupo, municipio: x.municipio, pr, mu: x.mu, ag: x.ag, zo: x.zo, po: x.po, pa: x.pa, recs: [], entera: false, sup: 0, otras: [], notas: [], camp }; partes.set(key, p); }
      if (x.recs.length) p.recs.push(...x.recs); else if (x.po) p.entera = true;
      p.sup += x.sup; if (!p.municipio && x.municipio) p.municipio = x.municipio;
      if (x.otras) p.otras.push(...x.otras.split(/[\n;]+/).map(s => s.trim()).filter(Boolean));
      if (x.notas && !p.notas.includes(x.notas)) p.notas.push(x.notas);
    }
    const lista = [...partes.values()];
    const clavePueblo = p => p.mu || norm(p.municipio);
    const porNombre = new Map();
    lista.forEach(p => { if (p.nombre) { const k = norm(p.nombre) + '|' + clavePueblo(p); if (!porNombre.has(k)) porNombre.set(k, []); porNombre.get(k).push(p); } });
    const pueblosPorNombre = new Map();
    porNombre.forEach((ps, k) => { const n = k.split('|')[0]; if (ps.length > 1) pueblosPorNombre.set(n, (pueblosPorNombre.get(n) || 0) + 1); });
    const corto = p => p.po && p.pa ? `${p.po}-${p.pa}` : '';
    lista.forEach(p => {
      const n = norm(p.nombre);
      const hermanos = p.nombre ? porNombre.get(n + '|' + clavePueblo(p)) : null;
      const varios = !!hermanos && hermanos.length > 1;
      if (p.grupoExcel) p.grupo = p.grupoExcel;
      else if (varios && !o.noAgrupar.has(n)) p.grupo = pueblosPorNombre.get(n) > 1 && p.municipio ? `${p.nombre} (${p.municipio})` : p.nombre;
      else p.grupo = '';
      p.nombreFinal = p.nombre ? (varios || o.noAgrupar.has(n) ? `${p.nombre} ${corto(p)}`.trim() : p.nombre) : `${p.municipio || 'Parcela'} ${corto(p)}`.trim();
    });
    const usados = new Map(); lista.forEach(p => { const k = norm(p.nombreFinal); usados.set(k, (usados.get(k) || 0) + 1); });
    lista.forEach(p => { if (usados.get(norm(p.nombreFinal)) > 1 && p.recs.length) p.nombreFinal += ` rec. ${[...new Set(p.recs)].sort((a, b) => a - b).join(',')}`; });
    const tomadas = new Set();
    return lista.map(p => {
      const recinto = p.entera ? '' : [...new Set(p.recs)].sort((a, b) => a - b).join(', ');
      const datos = { nombre: p.nombreFinal, grupo: p.grupo, titular: p.titular, municipio: p.municipio, provincia: p.pr, codMunicipio: p.mu, agregado: p.ag, zona: p.zo, poligono: p.po ? String(p.po) : '', parcela: p.pa ? String(p.pa) : '', recinto, otras: [...new Set(p.otras)].join('\n'), superficie: Math.round(p.sup * 10000) / 10000 || null, notas: p.notas.join('\n') };
      const mis = listaRecintos(recinto);
      const existe = p.po && p.pa
        ? db.fincas.find(f => !tomadas.has(f.id) && +f.codMunicipio === +p.mu && +f.poligono === p.po && +f.parcela === p.pa && (+f.agregado || 0) === p.ag && (+f.zona || 0) === p.zo && (!mis.length || !listaRecintos(f.recinto).length || listaRecintos(f.recinto).some(x => mis.includes(x))))
        : db.fincas.find(f => !tomadas.has(f.id) && norm(f.nombre) === norm(p.nombreFinal));
      if (existe) tomadas.add(existe.id);
      return { datos, camp: p.camp, existe };
    });
  }

  async function importarExcel(file) {
    let XLSX;
    try { XLSX = await cargarXLSX(); } catch (e) { toast('No se ha podido cargar el lector de Excel. Abre la app una vez con cobertura.'); return; }
    let r;
    try { const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' }); r = analizarHoja(XLSX, wb); }
    catch (e) { toast('No se ha podido leer ese archivo. Usa un Excel (.xlsx, .xls), .ods o .csv'); return; }
    if (r.error) { toast(r.error); return; }
    if (!r.registros.length) { toast('No hay fincas en la hoja: cada fila necesita un nombre o polígono y parcela'); return; }
    let noAgr = 'Cambio'; try { const s = localStorage.getItem(KEY + '.noAgrupar'); if (s != null) noAgr = s; } catch (e) { /* ignorar */ }
    const cuenta = {}; db.fincas.forEach(f => { if (+f.provincia) cuenta[+f.provincia] = (cuenta[+f.provincia] || 0) + 1; });
    const provDef = +Object.keys(cuenta).sort((a, b) => cuenta[b] - cuenta[a])[0] || 0;
    let provGuard = {}; try { provGuard = JSON.parse(localStorage.getItem(KEY + '.provPueblo') || '{}') || {}; } catch (e) { /* ignorar */ }
    const pueblos = [...new Map(r.registros.map(x => [String(x.mu || norm(x.municipio)), x.municipio ? `${x.municipio}${x.mu ? ' (' + x.mu + ')' : ''}` : `Municipio ${x.mu}`])).entries()].sort((a, b) => a[1].localeCompare(b[1], 'es'));
    const multi = r.campanas.length > 1;
    const campDef = new Set(r.campanas.filter(c => c === db.campanaActiva || c === campanaAnterior(db.campanaActiva)));
    if (multi && !campDef.size) campDef.add(r.campanas[r.campanas.length - 1]);
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera('Importar fincas')}<div class="body">
      <p class="small" style="margin:0">Hoja «${esc(r.hoja)}» · <b>Columnas reconocidas:</b> ${r.reconocidas.map(esc).join(', ')}${r.ignoradas.length ? `<br><span class="muted">No usadas: ${r.ignoradas.map(esc).join(', ')}</span>` : ''}</p>
      ${multi ? `<div class="field"><label>Campañas que se importan</label><div class="chkrow">${r.campanas.map(c => `<label class="chk"><input type="checkbox" name="impCamp" value="${c}"${campDef.has(c) ? ' checked' : ''}> ${c}</label>`).join('')}</div><div class="hint">Cultivo, laboreo y estiércol de cada campaña marcada.</div></div>` : ''}
      ${!r.conProvincia ? `<div class="field"><label>Provincia de cada pueblo</label><div class="hint" style="margin-bottom:4px">El Excel no trae la provincia y hace falta para encontrar las parcelas en SIGPAC. Revisa los pueblos de otra provincia.</div>
        <div class="provs">${pueblos.map(([k, n]) => `<label class="prov"><span>${esc(n)}</span><select class="impProv" data-k="${esc(k)}" aria-label="Provincia de ${esc(n)}"><option value="">Sin provincia</option>${PROVINCIAS.map((p, i) => i ? `<option value="${i}"${i === (+provGuard[k] || provDef) ? ' selected' : ''}>${esc(p)}</option>` : '').join('')}</select></label>`).join('')}</div></div>` : ''}
      <div class="field"><label for="impNoAgr">Nombres que no se agrupan</label><input id="impNoAgr" value="${esc(noAgr)}" placeholder="Ej.: Cambio" autocomplete="off"><div class="hint">Las parcelas con el mismo nombre en el mismo pueblo se juntan en una finca agrupada, salvo estos nombres (separados por comas).</div></div>
      <div id="impRes"></div>
      </div><footer><button class="btn" data-close>Cancelar</button><button class="btn primary" data-ok>Importar</button></footer></div>`, (el, d) => {
      const opciones = () => ({
        multi, campanas: new Set([...el.querySelectorAll('[name=impCamp]:checked')].map(i => i.value)),
        provPorPueblo: Object.fromEntries([...el.querySelectorAll('.impProv')].map(s => [s.dataset.k, +s.value || ''])),
        noAgrupar: new Set(el.querySelector('#impNoAgr').value.split(',').map(norm).filter(Boolean))
      });
      let plan = [];
      const pintar = () => {
        const o = opciones(); plan = construirPlan(r, o);
        const nuevas = plan.filter(p => !p.existe).length; const grupos = new Set(plan.map(p => p.datos.grupo).filter(Boolean));
        const ult = multi ? [...o.campanas].sort().pop() : null;
        const cuDe = p => { const ks = Object.keys(p.camp).sort(); const k = ult && p.camp[ult] ? ult : ks[ks.length - 1]; return k && p.camp[k] ? p.camp[k].cultivo : ''; };
        const sinProv = plan.filter(p => p.datos.poligono && !p.datos.provincia).length;
        el.querySelector('#impRes').innerHTML = `<p style="margin:0">${r.filas} filas → <b>${plan.length} parcelas</b> (${fmtN(sum(plan, p => p.datos.superficie))} ha) y <b>${grupos.size} fincas agrupadas</b>. ${nuevas} nuevas${plan.length - nuevas ? `, ${plan.length - nuevas} ya las tenías y se actualizarán` : ''}.</p>
          ${sinProv ? `<p class="small neg" style="margin:8px 0 0">${sinProv} parcelas sin provincia: elígela arriba para poder situarlas en el mapa.</p>` : ''}
          <div class="tablewrap" style="margin-top:10px"><table><thead><tr><th>Parcela</th><th>Finca agrupada</th><th>ha</th><th>Titular</th><th>Cultivo${ult ? ' ' + ult : ''}</th></tr></thead><tbody>${plan.slice(0, 40).map(p => `<tr><td>${esc(p.datos.nombre)}${p.existe ? ' <span class="tag">actualiza</span>' : ''}</td><td style="text-align:left">${esc(p.datos.grupo)}</td><td>${p.datos.superficie ? fmtN(p.datos.superficie) : '—'}</td><td style="text-align:left">${esc(p.datos.titular)}</td><td style="text-align:left">${cuDe(p) ? chip(cuDe(p)) : ''}</td></tr>`).join('')}</tbody></table></div>
          ${plan.length > 40 ? `<p class="note">…y ${plan.length - 40} más.</p>` : ''}`;
        el.querySelector('[data-ok]').textContent = `Importar ${plan.length}`;
      };
      el.querySelector('.body').addEventListener('change', pintar);
      el.querySelector('#impNoAgr').addEventListener('input', pintar);
      pintar();
      el.querySelector('[data-ok]').addEventListener('click', () => {
        try {
          localStorage.setItem(KEY + '.noAgrupar', el.querySelector('#impNoAgr').value);
          localStorage.setItem(KEY + '.provPueblo', JSON.stringify(Object.assign(provGuard, opciones().provPorPueblo)));
        } catch (e) { /* ignorar */ }
        for (const p of plan) {
          const f = p.existe || { id: uid() };
          for (const [k, v] of Object.entries(p.datos)) if (v !== '' && v != null) f[k] = v;
          if (!p.existe) db.fincas.push(f);
          for (const [c, dd] of Object.entries(p.camp)) {
            db.siembras = db.siembras.filter(s => !(s.fincaId === f.id && s.campana === c));
            if (dd.cultivo) {
              if (!db.cultivos.includes(dd.cultivo)) db.cultivos.push(dd.cultivo);
              db.siembras.push({ id: uid(), fincaId: f.id, campana: c, cultivo: dd.cultivo, variedad: dd.variedad || '', laboreo: dd.laboreo || '', estiercol: !!dd.estiercol });
            }
            if (!db.campanas.includes(c)) db.campanas.push(c);
          }
        }
        db.ejemplo = false;
        guardar(); volverA(0); ui.tab = 'fincas'; ui.fincaId = null; ui.grupo = null; render();
        toast(`${plan.length} parcelas importadas`);
        setTimeout(() => actualizarContornos({}), 400);
      });
    });
  }

  const COLS_EXPORT = ['Nombre', 'Finca agrupada', 'Titular', 'Superficie (ha)', 'Municipio', 'Provincia', 'Cód. municipio', 'Agregado', 'Zona', 'Polígono', 'Parcela', 'Recinto', 'Otras parcelas', 'Cultivo', 'Variedad', 'Laboreo', 'Basura', 'Campaña', 'Notas'];
  async function excelFincas(plantilla) {
    let XLSX;
    try { XLSX = await cargarXLSX(); } catch (e) { toast('No se ha podido cargar el generador de Excel. Abre la app una vez con cobertura.'); return; }
    const c = db.campanaActiva;
    const filas = plantilla
      ? [['Sarrana 663-180', 'Sarrana', 'Iván', 2.92, 'Lerma', 9, 197, 0, 0, 663, 180, '1, 2', '', 'Trigo', '', 'C', '', c, ''],
         ['Sarrana 663-181', 'Sarrana', 'Iván', 1.5, 'Lerma', 9, 197, 0, 0, 663, 181, '1', '', 'Trigo', '', 'C', 'X', c, ''],
         ['Consejera', '', 'Tomás', 6.82, 'Lerma', 9, 197, 0, 0, 631, 778, '1', '', 'Esparceta', '', 'D', '', c, 'Arrendada']]
      : db.fincas.slice().sort(byNombre).map(f => { const s = siembraDe(f.id, c); return [f.nombre, grupoDe(f), f.titular || '', f.superficie, f.municipio || '', f.provincia || '', f.codMunicipio || '', f.agregado || 0, f.zona || 0, f.poligono || '', f.parcela || '', f.recinto || '', f.otras || '', s ? s.cultivo : '', s ? s.variedad || '' : '', s ? s.laboreo || '' : '', s && s.estiercol ? 'X' : '', c, f.notas || '']; });
    const ws = XLSX.utils.aoa_to_sheet([COLS_EXPORT, ...filas]);
    ws['!cols'] = COLS_EXPORT.map((h, i) => ({ wch: [20, 16, 10, 12, 16, 9, 12, 9, 6, 9, 9, 9, 16, 12, 12, 8, 7, 9, 24][i] }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Fincas');
    if (plantilla) {
      const ins = XLSX.utils.aoa_to_sheet([['Cómo rellenar la hoja «Fincas»'], [''],
        ['• Una fila por recinto o por parcela. Las filas de la misma parcela (mismo polígono y parcela) se juntan y se suman sus superficies.'],
        ['• Las parcelas con el mismo Nombre en el mismo pueblo se juntan en una «finca agrupada» (por ejemplo, Sarrana = parcelas 180 y 181). Cada parcela se sigue viendo por separado.'],
        ['• Puedes poner la columna «Finca agrupada» para decir tú cómo se agrupan.'],
        ['• Provincia, Cód. municipio, Polígono, Parcela y Recinto son los códigos SIGPAC de tu declaración de la PAC. Si el municipio viene como «197-LERMA» también vale.'],
        ['• Laboreo: C = siembra directa con complemento, T = tradicional, D = siembra directa. Basura: X si se echó estiércol.'],
        ['• Campaña: 2026/27, o el año de cosecha (2027). También vale poner varias campañas en columnas, con el año encima de cada bloque Cultivo / Basura / Laboreo.']]);
      ins['!cols'] = [{ wch: 120 }];
      XLSX.utils.book_append_sheet(wb, ins, 'Instrucciones');
    }
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const ok = await guardarArchivo(blob, plantilla ? 'plantilla-fincas.xlsx' : `fincas-${hoyISO()}.xlsx`, !plantilla, 'Fincas');
    if (ok) toast(plantilla ? 'Plantilla descargada' : 'Excel de fincas creado');
  }
  $('#xlsIn').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; if (f) importarExcel(f); });

  // ---------- Copias de seguridad ----------
  async function exportar(compartir) {
    const fecha = hoyISO();
    let archivos = [];
    try { archivos = await Promise.all((await archivosTodos()).map(async a => ({ id: a.id, nombre: a.nombre, tipo: a.tipo, fecha: a.fecha, data: await blobADataURL(a.blob) }))); } catch (e) { /* sin archivos */ }
    const json = JSON.stringify({ app: 'cuaderno-campo', version: VERSION, exportado: new Date().toISOString(), datos: db, archivos }, null, 1);
    const ok = await guardarArchivo(new Blob([json], { type: 'application/json' }), `cuaderno-campo-${fecha}.json`, compartir, `Copia de seguridad del ${fmtFecha(fecha)}`);
    if (ok) { db.ultimaCopia = new Date().toISOString(); guardar(false); render(); toast('Copia hecha'); }
  }

  $('#fileIn').addEventListener('change', async e => {
    const file = e.target.files[0]; e.target.value = '';
    if (!file) return;
    let datos, o;
    try { o = JSON.parse(await file.text()); datos = o && o.datos ? o.datos : o; if (!datos || !Array.isArray(datos.fincas)) throw new Error(); }
    catch (err) { toast('Ese archivo no es una copia de Cuaderno de Campo'); return; }
    const r = await confirmar('¿Importar esta copia?',
      `La copia tiene <b>${datos.fincas.length} fincas</b>, ${(datos.labores || []).length} labores, ${(datos.movimientos || []).length} apuntes y ${(datos.cosechas || []).length} cosechas.<br><br>Sustituirá a los datos que hay ahora en este teléfono (${db.fincas.length} fincas).`, 'Importar');
    if (!r.ok) return;
    db = normalizar(datos); db.ultimaCopia = new Date().toISOString(); db.ejemplo = false; guardar(false);
    if (o && Array.isArray(o.archivos)) {
      let n = 0;
      for (const a of o.archivos) { try { const blob = await (await fetch(a.data)).blob(); await archivoGuardar({ id: a.id, nombre: a.nombre, tipo: a.tipo, fecha: a.fecha, blob }); n++; } catch (e) { /* sigue */ } }
      if (n) toast(`${n} fotos y PDF recuperados`);
    }
    ui.fincaId = null; ui.grupo = null; volverA(0); render(); toast('Copia importada');
  });

  // ---------- Datos de ejemplo ----------
  function cargarEjemplo() {
    const d = nuevoDB(); d.ejemplo = true;
    const actual = d.campanaActiva; const a = parseInt(actual, 10);
    const C = n => (a - n) + '/' + String((a - n + 1) % 100).padStart(2, '0');
    d.campanas = [C(2), C(1), actual];
    const F = (nombre, ref) => {
      const [pr, mu, ag, zo, po, pa, re] = ref.split(':'); const g = EJ_GEO[ref];
      const f = { id: uid(), nombre, superficie: Math.round(g.ha * 100) / 100, municipio: '', provincia: +pr, codMunicipio: +mu, agregado: +ag, zona: +zo, poligono: po, parcela: pa, recinto: re, otras: '', notas: 'Finca de ejemplo sobre un recinto SIGPAC real.' };
      f.geo = { k: refsKey(f), c: g.c, ha: Math.round(areaHa(g.c) * 100) / 100, t: hoyISO() };
      d.fincas.push(f); return f;
    };
    const fs = [F('Valdecasas', '47:215:0:0:403:30:1'), F('El Llano', '47:215:0:0:403:69:1'), F('Carrascal', '47:215:0:0:403:36:1'), F('La Hoya', '47:215:0:0:404:39:1'), F('Los Majuelos', '47:215:0:0:403:8:1')];
    const rot = [['Trigo', 'Titarro', 'Cebada'], ['Cebada', 'Trigo', 'Lenteja'], ['Esparceta', 'Esparceta', 'Esparceta'], ['Lenteja', 'Cebada', 'Trigo'], ['Barbecho', 'Anís dulce', 'Cilantro']];
    const rinde = { Trigo: [3400, 3900], Cebada: [3100, 3600], Titarro: [900, 1300], Lenteja: [800, 1100], 'Anís dulce': [500, 700], Cilantro: [700, 1000], Esparceta: [0, 0] };
    const precio = { Trigo: .23, Cebada: .2, Titarro: .32, Lenteja: .7, 'Anís dulce': 2.4, Cilantro: .9 };
    let seed = 7; const rnd = () => (seed = (seed * 9301 + 49297) % 233280) / 233280;
    fs.forEach((f, i) => {
      [C(2), C(1), actual].forEach((camp, j) => {
        const cu = rot[i][j]; d.siembras.push({ id: uid(), fincaId: f.id, campana: camp, cultivo: cu, variedad: '' });
        const y0 = parseInt(camp, 10);
        if (cu !== 'Barbecho') {
          d.labores.push({ id: uid(), fecha: `${y0}-10-${String(8 + i).padStart(2, '0')}`, tipo: 'Labrado', fincaId: f.id, producto: 'Cultivador', campana: camp, coste: null });
          const fs2 = `${y0}-${cu === 'Anís dulce' || cu === 'Cilantro' ? '11' : '10'}-${String(20 + i).padStart(2, '0')}`;
          if (fs2 <= hoyISO()) {
            const sem = Math.round(f.superficie * (cu === 'Trigo' || cu === 'Cebada' ? 180 : 100) * .45);
            const l = { id: uid(), fecha: fs2, tipo: 'Siembra', fincaId: f.id, producto: cu, dosis: cu === 'Trigo' || cu === 'Cebada' ? 180 : 100, unidad: 'kg/ha', campana: camp, coste: sem };
            d.labores.push(l);
            d.movimientos.push({ id: uid(), laborId: l.id, tipo: 'gasto', fecha: l.fecha, fincaId: f.id, campana: camp, categoria: 'Semilla', concepto: 'Siembra · ' + cu, importe: sem });
          }
        }
        if (camp === actual || cu === 'Barbecho' || cu === 'Esparceta') return;
        const ab = { id: uid(), fecha: `${y0 + 1}-02-${String(10 + i).padStart(2, '0')}`, tipo: 'Abonado', fincaId: f.id, producto: 'Urea 46 %', dosis: 150, unidad: 'kg/ha', campana: camp, coste: Math.round(f.superficie * 150 * .48) };
        if (cu === 'Trigo' || cu === 'Cebada') { d.labores.push(ab); d.movimientos.push({ id: uid(), laborId: ab.id, tipo: 'gasto', fecha: ab.fecha, fincaId: f.id, campana: camp, categoria: 'Abono', concepto: 'Abonado · Urea 46 %', importe: ab.coste }); }
        if (cu === 'Trigo' || cu === 'Cebada') {
          const t = { id: uid(), fecha: `${y0 + 1}-03-${String(12 + i).padStart(2, '0')}`, tipo: 'Tratamiento', fincaId: f.id, producto: 'Herbicida', dosis: 1, unidad: 'L/ha', registro: 'ES-00000', motivo: 'Avena loca', campana: camp, coste: Math.round(f.superficie * 28) };
          d.labores.push(t); d.movimientos.push({ id: uid(), laborId: t.id, tipo: 'gasto', fecha: t.fecha, fincaId: f.id, campana: camp, categoria: 'Fitosanitarios', concepto: 'Tratamiento · Herbicida', importe: t.coste });
        }
        const [lo, hi] = rinde[cu]; const kgHa = lo + rnd() * (hi - lo); const kg = Math.round(f.superficie * kgHa / 10) * 10;
        const fc = `${y0 + 1}-07-${String(2 + i * 3).padStart(2, '0')}`;
        d.cosechas.push({ id: uid(), fincaId: f.id, campana: camp, fecha: fc, kg, humedad: 11 + Math.round(rnd() * 20) / 10, notas: 'Cooperativa' });
        d.movimientos.push({ id: uid(), tipo: 'ingreso', fecha: `${y0 + 1}-08-${String(10 + i).padStart(2, '0')}`, fincaId: f.id, campana: camp, categoria: 'Venta de grano', concepto: `${fmtN(kg, 0)} kg de ${cu.toLowerCase()}`, kilos: kg, precio: precio[cu], importe: Math.round(kg * precio[cu] * 100) / 100 });
      });
    });
    [C(2), C(1)].forEach(camp => {
      const y0 = parseInt(camp, 10);
      d.movimientos.push({ id: uid(), tipo: 'gasto', fecha: `${y0}-11-15`, fincaId: '', campana: camp, categoria: 'Seguro', concepto: 'Seguro agrario', importe: 820 });
      d.movimientos.push({ id: uid(), tipo: 'gasto', fecha: `${y0 + 1}-06-20`, fincaId: '', campana: camp, categoria: 'Gasóleo', concepto: 'Gasóleo agrícola 1.500 L', importe: 1650 });
      d.movimientos.push({ id: uid(), tipo: 'ingreso', fecha: `${y0 + 1}-12-01`, fincaId: '', campana: camp, categoria: 'Subvención PAC', concepto: 'Pago básico', importe: 5200 });
    });
    d.movimientos.push({ id: uid(), tipo: 'gasto', fecha: hoyISO(), fincaId: '', campana: actual, categoria: 'Seguro', concepto: 'Seguro agrario', importe: 840 });
    db = d; guardar(false); ui.tab = 'fincas'; ui.mapaAjustado = false; render(); toast('Datos de ejemplo cargados');
  }

  // ---------- Archivos (fotos y PDF de albaranes y facturas): se guardan en el teléfono (IndexedDB) ----------
  let idbP = null;
  function idb() {
    return idbP || (idbP = new Promise((res, rej) => {
      const r = indexedDB.open('cuadernoCampoArchivos', 1);
      r.onupgradeneeded = () => r.result.createObjectStore('archivos', { keyPath: 'id' });
      r.onsuccess = () => res(r.result); r.onerror = () => { idbP = null; rej(r.error); };
    }));
  }
  async function idbOp(modo, fn) {
    const d = await idb();
    return new Promise((res, rej) => { const t = d.transaction('archivos', modo); const r = fn(t.objectStore('archivos')); t.oncomplete = () => res(r && r.result); t.onerror = () => rej(t.error); });
  }
  const archivoGuardar = rec => idbOp('readwrite', s => s.put(rec));
  const archivoLeer = id => idbOp('readonly', s => s.get(id));
  const archivoBorrar = id => idbOp('readwrite', s => s.delete(id));
  const archivosTodos = () => idbOp('readonly', s => s.getAll());
  const archivosVaciar = () => idbOp('readwrite', s => s.clear());
  // Las fotos se reducen (máx. 1800 px) para que quepan muchas en el teléfono
  async function prepararArchivo(file) {
    if (file.type && file.type.startsWith('image/')) {
      try {
        const bmp = await createImageBitmap(file);
        const k = Math.min(1, 1800 / Math.max(bmp.width, bmp.height));
        const c = document.createElement('canvas'); c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
        c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
        const blob = await new Promise(r => c.toBlob(r, 'image/jpeg', 0.82));
        if (blob) return { blob, tipo: 'image/jpeg', nombre: (file.name || 'foto').replace(/\.\w+$/, '') + '.jpg' };
      } catch (e) { /* se guarda tal cual */ }
    }
    return { blob: file, tipo: file.type || 'application/octet-stream', nombre: file.name || 'documento' };
  }
  async function verArchivo(id) {
    let a; try { a = await archivoLeer(id); } catch (e) { a = null; }
    if (!a) { toast('Ese archivo no está en este teléfono'); return; }
    const url = URL.createObjectURL(a.blob);
    if (a.tipo.startsWith('image/')) {
      abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera(esc(a.nombre))}<div class="body"><img src="${url}" alt="${esc(a.nombre)}" class="visor-img"></div></div>`);
    } else {
      const w = window.open(url, '_blank');
      if (!w) { const l = document.createElement('a'); l.href = url; l.download = a.nombre; document.body.appendChild(l); l.click(); l.remove(); }
    }
  }
  const blobADataURL = b => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(b); });

  // ---------- Registro oficial de fitosanitarios (MAPA), incluido en la app ----------
  let fitosP = null;
  function cargarFitos() {
    return fitosP || (fitosP = fetch('fitos.json').then(r => { if (!r.ok) throw new Error('fitos'); return r.json(); })
      .then(d => { d.porId = new Map(d.productos.map(p => [p.id, p])); return d; })
      .catch(e => { fitosP = null; throw e; }));
  }
  let FITOS = null; // se rellena al cargar
  const regProducto = p => (p && p.regId && FITOS) ? FITOS.porId.get(p.regId) : null;
  // Con qué cultivos del Registro se compara cada cultivo de la app
  // (también la categoría general: «Herbáceas extensivas» incluye cereales, leguminosas, forrajeras y oleaginosas)
  const HE = 'Herbáceas extensivas', HI = 'Herbáceas intensivas';
  const CULTIVO_REG = {
    trigo: ['Trigo', 'Cereales de invierno / primavera', HE], cebada: ['Cebada', 'Cereales de invierno / primavera', HE],
    avena: ['Avena', 'Cereales de invierno / primavera', HE], 'avena forrajera': ['Avena', 'Cereales de invierno / primavera', HE],
    centeno: ['Centeno', 'Cereales de invierno / primavera', HE], triticale: ['Triticale', 'Cereales de invierno / primavera', HE],
    lenteja: ['Lenteja', 'Leguminosas de grano', HE], garbanzo: ['Garbanzo', 'Leguminosas de grano', HE], guisante: ['Guisante para grano', 'Leguminosas de grano', HE],
    titarro: ['Almorta', 'Leguminosas de grano', HE], yeros: ['Yeros', 'Leguminosas forrajeras', HE], veza: ['Veza', 'Leguminosas forrajeras', HE],
    esparceta: ['Esparceta', 'Leguminosas forrajeras', HE], alfalfa: ['Alfalfa', 'Leguminosas forrajeras', HE],
    'anis dulce': ['Anís, matalauva', HI], anis: ['Anís, matalauva', HI], cilantro: ['Cilantro', HI], girasol: ['Girasol', HE], colza: ['Colza', HE], camelina: ['Camelina', HE],
    barbecho: ['Barbechos', 'Rastrojeras'], chopo: ['Chopos y álamos']
  };
  const regDeCultivo = cu => CULTIVO_REG[norm(cu)] || null;
  const fechaDMY = s => { const m = String(s || '').match(/^(\d{2})-(\d{2})-(\d{4})$/); return m ? `${m[3]}-${m[2]}-${m[1]}` : ''; };
  // Dosis por hectárea pasada a L/ha o kg/ha
  function baseHa(v, u) {
    u = String(u || '').toLowerCase().replace(/\s/g, '');
    if (v == null || Number.isNaN(+v)) return null;
    if (u === 'cc/ha' || u === 'ml/ha') return { v: v / 1000, u: 'L' };
    if (u === 'l/ha') return { v: +v, u: 'L' };
    if (u === 'g/ha') return { v: v / 1000, u: 'kg' };
    if (u === 'kg/ha') return { v: +v, u: 'kg' };
    return null;
  }
  const UNI_REG = { 'cc/ha': 'cc/ha', 'ml/ha': 'cc/ha', 'l/ha': 'L/ha', 'g/ha': 'g/ha', 'kg/ha': 'kg/ha' };
  // Comprueba un fitosanitario en un cultivo: autorización, dosis y estado
  function revisarFito(reg, cultivo, dosis, unidad) {
    const r = { graves: [], notas: [], usos: [], rango: null, autorizado: null };
    if (!reg) return r;
    if (reg.e && reg.e !== 'Vigente') r.graves.push(`«${reg.n}» figura como «${reg.e}» en el Registro.`);
    const cad = fechaDMY(reg.cad); if (cad && cad < hoyISO()) r.graves.push(`La autorización de «${reg.n}» caducó el ${reg.cad}.`);
    const lim = fechaDMY(reg.lv); if (lim && lim < hoyISO()) r.notas.push(`Fecha límite de venta: ${reg.lv}. Revisa hasta cuándo se puede usar.`);
    if (!cultivo) { r.notas.push('La parcela no tiene cultivo asignado en esta campaña: no se puede comprobar si el producto está autorizado.'); return r; }
    const cs = regDeCultivo(cultivo);
    if (!cs) { r.notas.push(`«${cultivo}» no está en el Registro: compruébalo en la ficha oficial.`); return r; }
    const aut = cs.filter(c => (reg.c || []).includes(c) || (reg.u && reg.u[c]));
    r.autorizado = aut.length > 0;
    if (!aut.length) { r.graves.push(`«${reg.n}» no está autorizado en ${cultivo}.`); return r; }
    if (norm(cultivo) === 'titarro') r.notas.push('El titarro no figura en el Registro; se compara con almorta y leguminosas de grano. Confírmalo con tu técnico.');
    r.usos = aut.flatMap(c => ((reg.u && reg.u[c]) || []).map(x => ({ c, ag: x[0], min: x[1], max: x[2], u: x[3] })));
    const bs = r.usos.map(x => ({ a: baseHa(x.min, x.u), b: baseHa(x.max, x.u) })).filter(x => x.a && x.b);
    if (bs.length) {
      const u0 = bs[0].a.u; const mm = bs.filter(x => x.a.u === u0);
      r.rango = { min: Math.min(...mm.map(x => x.a.v)), max: Math.max(...mm.map(x => x.b.v)), u: u0 };
      const d = baseHa(dosis, unidad);
      if (d && d.u === r.rango.u) {
        if (d.v > r.rango.max * 1.0001) r.graves.push(`La dosis (${fmtN(dosis, 4)} ${unidad}) supera la máxima autorizada en ${cultivo} (${fmtN(r.rango.max, 4)} ${r.rango.u}/ha).`);
        else if (d.v < r.rango.min * 0.9999) r.notas.push(`La dosis está por debajo de la mínima de la etiqueta en ${cultivo} (${fmtN(r.rango.min, 4)} ${r.rango.u}/ha).`);
      } else if (d && d.u !== r.rango.u) r.notas.push(`La dosis autorizada va en ${r.rango.u}/ha y has puesto ${unidad}.`);
    }
    return r;
  }
  const textoUsos = usos => [...new Set(usos.map(x => `${x.ag ? x.ag.replace(/,.*$/, '') + ': ' : ''}${fmtN(x.min, 4)}${x.max !== x.min ? '-' + fmtN(x.max, 4) : ''} ${x.u}`))].slice(0, 8).join(' · ');
  const urlFicha = id => `https://servicio.mapa.gob.es/regfiweb/Productos/ExportFichaProductoPdfGet?idProducto=${id}`;
  const LIQUIDOS = /\[(EC|SC|SL|OD|EW|SE|ME|CS|DC|EO|OL|AL|UL|ZC|SU|FS|LS|ES|CF|OF)\]/i;
  const unidadDeFormulado = f => LIQUIDOS.test(f || '') ? 'L' : 'kg';

  // ---------- Almacén: productos, existencias por titular, albaranes y facturas ----------
  const TIPOS_PROD = ['Fitosanitario', 'Abono', 'Semilla', 'Combustible', 'Otro'];
  const UNI_PROD = ['L', 'kg', 't', 'ud'];
  const TIPO_DE_LABOR = { Tratamiento: 'Fitosanitario', Abonado: 'Abono', Siembra: 'Semilla' };
  const producto = id => db.productos.find(p => p.id === id);
  const titularesTodos = () => [...new Set([...titulares(), ...db.titulares.map(t => t.nombre), ...db.maquinas.map(m => m.titular), ...db.documentos.map(d => d.titular), ...db.ajustes.map(a => a.titular)].map(t => String(t || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
  const titularDe = fid => String((finca(fid) || {}).titular || '').trim();
  function movimientosProducto(pid, tit) {
    const out = [];
    db.documentos.filter(d => d.tipo === 'albaran' && (tit == null || d.titular === tit)).forEach(d => (d.lineas || []).forEach(l => { if (l.productoId === pid) out.push({ fecha: d.fecha, cant: +l.cantidad || 0, txt: `Albarán ${d.numero || ''} · ${d.proveedor || ''}`.trim(), titular: d.titular, doc: d.id }); }));
    db.labores.filter(l => l.consumo && l.consumo.productoId === pid && (tit == null || l.consumo.titular === tit)).forEach(l => out.push({ fecha: l.fecha, cant: -(+l.consumo.cantidad || 0), txt: `${l.tipo} · ${(finca(l.fincaId) || {}).nombre || ''}`, titular: l.consumo.titular, labor: l.id }));
    db.ajustes.filter(a => a.productoId === pid && (tit == null || a.titular === tit)).forEach(a => out.push({ fecha: a.fecha, cant: +a.cantidad || 0, txt: `Ajuste${a.motivo ? ' · ' + a.motivo : ''}`, titular: a.titular, ajuste: a.id }));
    const g = gasoleoAlm();
    if (g && g.id === pid) { const R = repartoGasoleo(); db.labores.forEach(l => { const x = R.porLabor.get(l.id); if (!x) return; x.partes.forEach(pt => { if (tit == null || pt.t === tit) out.push({ fecha: l.fecha, cant: -Math.round(pt.q * 1000) / 1000, txt: `Gasóleo · ${l.tipo} · ${(finca(l.fincaId) || {}).nombre || ''}${pt.t ? '' : ''}`, titular: pt.t, labor: l.id }); }); }); }
    return out.sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
  }
  const existencias = (pid, tit) => Math.round(sum(movimientosProducto(pid, tit), m => m.cant) * 1000) / 1000;
  const fmtCant = (v, u) => `${fmtN(v, 3)} ${u || ''}`.trim();
  function productosConMovimiento(tit) {
    const ids = new Set();
    db.documentos.filter(d => d.tipo === 'albaran' && (tit == null || d.titular === tit)).forEach(d => (d.lineas || []).forEach(l => ids.add(l.productoId)));
    db.labores.forEach(l => { if (l.consumo && (tit == null || l.consumo.titular === tit)) ids.add(l.consumo.productoId); });
    db.ajustes.forEach(a => { if (tit == null || a.titular === tit) ids.add(a.productoId); });
    const g = gasoleoAlm(); const R = g && repartoGasoleo(); if (R && [...R.porLabor.values()].some(x => x.partes.some(pt => tit == null || pt.t === tit))) ids.add(g.id);
    return [...ids].map(producto).filter(Boolean).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
  }
  const totalDoc = d => Math.round(sum(d.lineas || [], l => l.importe) * 100) / 100;
  function cotejo(f) {
    const albs = db.documentos.filter(d => (f.albaranes || []).includes(d.id));
    const sumAlb = Math.round(sum(albs, totalDoc) * 100) / 100;
    const base = f.base != null && f.base !== '' ? +f.base : (f.total ? Math.round(f.total / (1 + (+f.iva || 0) / 100) * 100) / 100 : null);
    const prods = new Map();
    albs.forEach(a => (a.lineas || []).forEach(l => { const p = prods.get(l.productoId) || { alb: 0, fac: 0 }; p.alb += +l.cantidad || 0; prods.set(l.productoId, p); }));
    (f.lineas || []).forEach(l => { const p = prods.get(l.productoId) || { alb: 0, fac: 0 }; p.fac += +l.cantidad || 0; prods.set(l.productoId, p); });
    const difProd = (f.lineas || []).length ? [...prods].filter(([, p]) => Math.abs(p.alb - p.fac) > 0.001).map(([pid, p]) => ({ pid, ...p })) : [];
    const conPrecio = albs.some(a => (a.lineas || []).some(l => +l.importe > 0));
    const dif = base != null && albs.length && conPrecio ? Math.round((base - sumAlb) * 100) / 100 : null;
    const hayLineas = (f.lineas || []).length > 0;
    let estado;
    if (!albs.length) estado = 'sin';
    else if (conPrecio) estado = (dif != null && Math.abs(dif) <= 0.5 && !difProd.length) ? 'ok' : (dif == null && !difProd.length ? 'sin' : 'mal');
    else estado = hayLineas ? (difProd.length ? 'mal' : 'ok') : 'sin';
    return { albs, sumAlb, base, dif, difProd, estado, conPrecio, hayLineas };
  }
  const ESTADO_COT = { ok: ['Cuadra', 'pos'], mal: ['No cuadra', 'neg'], sin: ['Sin cotejar', 'muted'] };

  function vAlmacen() {
    const a = ui.alm; const tits = titularesTodos();
    let h = `<div class="segs" role="group" aria-label="Apartado">${[['inventario', 'Inventario'], ['albaran', 'Albaranes'], ['factura', 'Facturas']].map(([k, t]) => `<button data-act="almVista" data-k="${k}" aria-pressed="${a.vista === k}">${t}</button>`).join('')}</div>`;
    if (tits.length) h += `<div class="carpetas" role="group" aria-label="Carpeta del titular">${[['', 'Todos'], ...tits.map(t => [t, t])].map(([k, t]) => `<button data-act="almTitular" data-k="${esc(k)}" aria-pressed="${a.titular === k}">${k ? ico.carpeta : ''}${esc(t)}</button>`).join('')}</div>`;
    const tit = a.titular;
    if (a.vista === 'inventario') {
      const grupos = tit ? [tit] : (tits.length ? tits : ['']);
      let algo = false;
      for (const t of grupos) {
        const ps = productosConMovimiento(t || null); if (!ps.length) continue; algo = true;
        h += `${grupos.length > 1 ? `<h2 class="sec">${esc(t)}</h2>` : ''}<div class="list">${ps.map(p => { const ex = existencias(p.id, t || null); const reg = regProducto(p);
          return `<button class="row" data-act="verProducto" data-id="${p.id}" data-t="${esc(t)}"><div class="main"><div class="t">${esc(p.nombre)}</div><div class="s">${esc([p.tipo, p.registro ? 'Nº reg. ' + p.registro : '', reg && reg.e !== 'Vigente' ? reg.e : ''].filter(Boolean).join(' · '))}</div></div><div class="end big ${ex < 0 ? 'neg' : ex === 0 ? 'muted' : ''}">${fmtCant(ex, p.unidad)}</div></button>`; }).join('')}</div>`;
      }
      const huerfanos = !tit ? db.productos.filter(p => !movimientosProducto(p.id).length) : [];
      if (huerfanos.length) { algo = true; h += `<h2 class="sec">Sin movimientos</h2><div class="list">${huerfanos.map(p => `<button class="row" data-act="verProducto" data-id="${p.id}" data-t=""><div class="main"><div class="t">${esc(p.nombre)}</div><div class="s">${esc(p.tipo)} · sin albaranes ni labores</div></div><div class="end small muted">Tocar para borrar</div></button>`).join('')}</div>`; }
      if (!algo) h += vacio('El almacén está vacío', 'Sube el albarán de lo que compras (fitosanitarios, abono, semilla) y aquí verás lo que le queda a cada titular. Al apuntar un tratamiento se descuenta solo.', '<button class="btn primary" data-act="nuevoDoc" data-k="albaran">Subir albarán</button>');
      return h;
    }
    const docs = db.documentos.filter(d => d.tipo === a.vista && (!tit || d.titular === tit)).sort(byFechaDesc);
    if (!docs.length) return h + vacio(a.vista === 'albaran' ? 'No hay albaranes' : 'No hay facturas', a.vista === 'albaran' ? 'Haz una foto al albarán o sube el PDF, y apunta los productos y cantidades. Se guardan en la carpeta de cada titular.' : 'Sube la factura y marca los albaranes que incluye: la app comprueba si cuadran.', `<button class="btn primary" data-act="nuevoDoc" data-k="${a.vista}">${a.vista === 'albaran' ? 'Subir albarán' : 'Subir factura'}</button>`);
    let mes = ''; let lista = '';
    for (const d of docs) {
      const m = (d.fecha || '').slice(0, 7);
      if (m !== mes) { if (mes) lista += '</div>'; mes = m; lista += `<h2 class="sec">${m ? MESES_L[+m.slice(5) - 1] + ' ' + m.slice(0, 4) : 'Sin fecha'}</h2><div class="list">`; }
      let estado = '';
      if (d.tipo === 'albaran') { const fac = db.documentos.find(f => f.tipo === 'factura' && (f.albaranes || []).includes(d.id)); estado = fac ? `<span class="pill pos">Facturado</span>` : '<span class="pill warn">Sin factura</span>'; }
      else { const c = cotejo(d); estado = `<span class="pill ${ESTADO_COT[c.estado][1]}">${ESTADO_COT[c.estado][0]}</span>`; }
      const importe = d.tipo === 'factura' ? (d.total || totalDoc(d)) : totalDoc(d);
      lista += `<button class="row" data-act="editDoc" data-id="${d.id}">${fechaBox(d.fecha)}<div class="main"><div class="t">${esc(d.proveedor || 'Sin proveedor')}${d.numero ? ' · nº ' + esc(d.numero) : ''}</div><div class="s">${esc([tit ? '' : d.titular, `${(d.lineas || []).length} producto${(d.lineas || []).length === 1 ? '' : 's'}`, (d.archivos || []).length ? `${d.archivos.length} archivo${d.archivos.length === 1 ? '' : 's'}` : 'sin archivo'].filter(Boolean).join(' · '))}</div>${estado}</div><div class="end big">${importe ? eur(importe, 2) : ''}</div></button>`;
    }
    return h + lista + '</div>';
  }

  // Elegir producto: los del almacén, buscar en el Registro oficial o crear uno nuevo
  function hojaElegirProducto(tipoPref, alElegir) {
    cargarFitos().then(d => { FITOS = d; }).catch(() => {});
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera('Elegir producto')}<div class="body">
      <div class="field"><input id="epBuscar" type="search" placeholder="Nombre o nº de registro (p. ej. Accresto, 19549)" autocomplete="off" aria-label="Buscar producto"></div>
      <div id="epRes"></div></div></div>`, (el, d) => {
      const inp = el.querySelector('#epBuscar'), res = el.querySelector('#epRes');
      const pintar = () => {
        const q = norm(inp.value);
        const mios = db.productos.filter(p => !q || norm(p.nombre + ' ' + (p.registro || '')).includes(q)).slice(0, 30);
        let h = mios.length ? `<h2 class="sec">En tu almacén</h2><div class="list">${mios.map(p => `<button class="row" data-ep="mio" data-id="${p.id}"><div class="main"><div class="t">${esc(p.nombre)}</div><div class="s">${esc([p.tipo, p.registro ? 'Nº reg. ' + p.registro : '', p.unidad].filter(Boolean).join(' · '))}</div></div></button>`).join('')}</div>` : '';
        if (q.length >= 2 && FITOS) {
          const yaReg = new Set(db.productos.map(p => p.regId));
          const hits = FITOS.productos.filter(p => !yaReg.has(p.id) && norm(p.n + ' ' + p.r).includes(q)).slice(0, 25);
          h += `<h2 class="sec">Registro oficial de fitosanitarios</h2>${hits.length ? `<div class="list">${hits.map(p => `<button class="row" data-ep="reg" data-id="${p.id}"><div class="main"><div class="t">${esc(p.n)}</div><div class="s">${esc(['Nº reg. ' + p.r, p.f, p.e].filter(Boolean).join(' · '))}</div></div><span class="pill ${p.e === 'Vigente' ? 'pos' : 'neg'}">${esc(p.e)}</span></button>`).join('')}</div>` : '<p class="note">Ningún producto del Registro coincide.</p>'}`;
        } else if (q.length >= 2 && !FITOS) h += '<p class="note">Cargando el Registro de fitosanitarios…</p>';
        if (q.length >= 2) h += `<div class="card pad nuevo-prod"><div class="tag">Crear producto nuevo</div><b>${esc(inp.value.trim())}</b>
          <div class="grid2"><div class="field"><label for="epTipo">Tipo</label><select id="epTipo">${TIPOS_PROD.map(t => `<option${t === (tipoPref && tipoPref !== 'Fitosanitario' ? tipoPref : 'Abono') ? ' selected' : ''}>${t}</option>`).join('')}</select></div>
          <div class="field"><label for="epUni">Se mide en</label><select id="epUni">${UNI_PROD.map(u => `<option>${u}</option>`).join('')}</select></div></div>
          <button class="btn block" data-ep="nuevo" type="button">Crear «${esc(inp.value.trim())}»</button><p class="small muted" style="margin:0">Para fitosanitarios usa mejor el Registro de arriba: así la app comprueba cultivos y dosis.</p></div>`;
        res.innerHTML = h || '<p class="note">Escribe el nombre del producto.</p>';
      };
      inp.addEventListener('input', pintar);
      cargarFitos().then(dd => { FITOS = dd; pintar(); }).catch(() => {});
      pintar(); setTimeout(() => inp.focus(), 50);
      res.addEventListener('click', e => {
        const b = e.target.closest('[data-ep]'); if (!b) return;
        let p;
        if (b.dataset.ep === 'mio') p = producto(b.dataset.id);
        else if (b.dataset.ep === 'reg') {
          const r = FITOS.porId.get(+b.dataset.id);
          p = { id: uid(), nombre: r.n, tipo: 'Fitosanitario', unidad: unidadDeFormulado(r.f), regId: r.id, registro: r.r, formulado: r.f };
          db.productos.push(p);
        } else {
          const nombre = inp.value.trim(); if (!nombre) return;
          p = { id: uid(), nombre, tipo: el.querySelector('#epTipo').value, unidad: el.querySelector('#epUni').value };
          db.productos.push(p);
        }
        guardar(); volverA(d - 1); alElegir(p);
      });
    });
  }

  // Albarán o factura
  function formDocumento(tipo, id, pre) {
    const doc = id ? db.documentos.find(x => x.id === id) : null;
    if (doc) tipo = doc.tipo;
    const tits = titularesTodos();
    const v = doc ? JSON.parse(JSON.stringify(doc)) : { tipo, titular: ui.alm.titular || (tits.length === 1 ? tits[0] : ''), fecha: hoyISO(), proveedor: '', numero: '', lineas: [], archivos: [], albaranes: [], iva: tipo === 'factura' ? 10 : null, notas: '' };
    if (!doc && pre) Object.assign(v, pre);
    const nuevosArch = []; const quitarArch = [];
    const esFac = tipo === 'factura';
    const proveedores = [...new Set(db.documentos.map(x => x.proveedor).filter(Boolean))];
    cargarFitos().then(dd => { FITOS = dd; }).catch(() => {});
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera(doc ? (esFac ? 'Factura' : 'Albarán') : (esFac ? 'Nueva factura' : 'Nuevo albarán'))}<form class="body" id="frmDoc" novalidate>
      <div class="grid2"><div class="field"><label for="dTit">Titular (carpeta)</label><input id="dTit" value="${esc(v.titular)}" list="dl_tit" autocomplete="off" placeholder="Ej.: Ivan"><datalist id="dl_tit">${tits.map(t => `<option value="${esc(t)}">`).join('')}</datalist></div>
      <div class="field"><label for="dFecha">Fecha</label><input id="dFecha" type="date" value="${esc(v.fecha)}"></div></div>
      <div class="grid2"><div class="field"><label for="dProv">Proveedor</label><input id="dProv" value="${esc(v.proveedor)}" list="dl_prov" autocomplete="off" placeholder="Ej.: Cooperativa"><datalist id="dl_prov">${proveedores.map(t => `<option value="${esc(t)}">`).join('')}</datalist></div>
      <div class="field"><label for="dNum">Nº de ${esFac ? 'factura' : 'albarán'}</label><input id="dNum" value="${esc(v.numero)}" autocomplete="off"></div></div>
      <div class="field"><label>Foto o PDF</label><div id="dArch" class="archivos"></div>
        <div class="btns" style="justify-content:flex-start"><label class="btn sm">${ico.camara} Hacer foto<input type="file" accept="image/*" capture="environment" hidden data-sube></label><label class="btn sm">${ico.clip} Subir foto o PDF<input type="file" accept="image/*,application/pdf" multiple hidden data-sube></label></div></div>
      ${esFac ? `<div class="field"><label>Albaranes que incluye</label><div id="dAlbs"></div></div>` : ''}
      <div class="field"><label>${esFac ? 'Productos de la factura (opcional, para comparar cantidades)' : 'Productos'}</label><div id="dLineas"></div><button type="button" class="btn sm" data-addl>${ico.plus} Añadir producto</button></div>
      ${esFac ? `<div class="grid3"><div class="field"><label for="dBase">Base (sin IVA)</label><input id="dBase" inputmode="decimal" value="${esc(numStr(v.base))}" autocomplete="off"></div><div class="field"><label for="dIva">IVA %</label><input id="dIva" inputmode="decimal" value="${esc(numStr(v.iva))}" autocomplete="off"></div><div class="field"><label for="dTotal">Total</label><input id="dTotal" inputmode="decimal" value="${esc(numStr(v.total))}" autocomplete="off"></div></div>
        <div id="dCot" class="card pad cotejo"></div>
        <label class="chk"><input type="checkbox" id="dGasto"${doc && db.movimientos.some(m => m.docId === doc.id) ? ' checked' : ''}> Apuntar el total en Cuentas como gasto</label>` : ''}
      <div class="field"><label for="dNotas">Notas</label><textarea id="dNotas" rows="2">${esc(v.notas)}</textarea></div>
      </form><footer>${doc ? '<button class="btn danger" type="button" data-del>Borrar</button>' : ''}<button class="btn primary" type="submit" form="frmDoc">Guardar</button></footer></div>`, (el, d) => {
      const form = el.querySelector('form');
      const leer = () => {
        v.titular = el.querySelector('#dTit').value.trim(); v.fecha = el.querySelector('#dFecha').value; v.proveedor = el.querySelector('#dProv').value.trim(); v.numero = el.querySelector('#dNum').value.trim(); v.notas = el.querySelector('#dNotas').value.trim();
        el.querySelectorAll('.linea').forEach((row, i) => {
          const l = v.lineas[i]; if (!l) return;
          l.cantidad = parseNum(row.querySelector('[data-c]').value, true); l.precio = parseNum(row.querySelector('[data-p]').value); const imp = parseNum(row.querySelector('[data-i]').value, true);
          l.importe = Number.isFinite(imp) && imp !== null ? imp : null;
        });
        if (esFac) { v.base = parseNum(el.querySelector('#dBase').value, true); v.iva = parseNum(el.querySelector('#dIva').value); v.total = parseNum(el.querySelector('#dTotal').value, true); v.albaranes = [...el.querySelectorAll('[data-alb]:checked')].map(i => i.value); }
      };
      const pintarArch = async () => {
        const cont = el.querySelector('#dArch');
        const items = [];
        for (const aid of v.archivos) { let a = nuevosArch.find(x => x.id === aid); if (!a) { try { a = await archivoLeer(aid); } catch (e) { a = null; } } items.push({ id: aid, a }); }
        cont.innerHTML = items.length ? items.map(({ id: aid, a }) => `<div class="arch">${a && a.tipo.startsWith('image/') ? `<img src="${URL.createObjectURL(a.blob)}" alt="">` : `<span class="arch-pdf">${a ? 'PDF' : '?'}</span>`}<span class="arch-n">${esc(a ? a.nombre : 'No está en este teléfono')}</span><button type="button" class="btn sm" data-ver="${aid}">Ver</button><button type="button" class="iconbtn" data-quitar="${aid}" aria-label="Quitar archivo">${ico.x}</button></div>`).join('') : '<p class="note" style="margin:0">Sin foto ni PDF todavía.</p>';
      };
      const pintarLineas = () => {
        el.querySelector('#dLineas').innerHTML = v.lineas.length ? v.lineas.map((l, i) => { const p = producto(l.productoId) || { nombre: '¿?', unidad: '' };
          return `<div class="linea"><div class="linea-top"><b>${esc(p.nombre)}</b><span class="small muted">${esc([p.tipo, p.registro ? 'Nº reg. ' + p.registro : ''].filter(Boolean).join(' · '))}</span><button type="button" class="iconbtn" data-dell="${i}" aria-label="Quitar producto">${ico.x}</button></div>
            <div class="grid3"><label class="mini">Cantidad (${esc(p.unidad)})<input data-c inputmode="decimal" value="${esc(numStr(l.cantidad))}" autocomplete="off"></label><label class="mini">Precio €/${esc(p.unidad)}<input data-p inputmode="decimal" value="${esc(numStr(l.precio))}" autocomplete="off"></label><label class="mini">Importe €<input data-i inputmode="decimal" value="${esc(numStr(l.importe))}" autocomplete="off"></label></div></div>`; }).join('')
          + `<p class="small" style="margin:6px 2px 0"><b id="dTot">Total productos: ${eur(totalDoc(v), 2)}</b>${esFac ? '' : ' (sin IVA)'}</p>` : '<p class="note" style="margin:0">Añade cada producto con su cantidad.</p>';
      };
      const pintarTotal = () => { const t = el.querySelector('#dTot'); if (t) t.textContent = `Total productos: ${eur(totalDoc(v), 2)}`; };
      const pintarAlbs = () => {
        if (!esFac) return;
        const cont = el.querySelector('#dAlbs');
        const otros = new Set(db.documentos.filter(f => f.tipo === 'factura' && f.id !== (doc && doc.id)).flatMap(f => f.albaranes || []));
        const albs = db.documentos.filter(a => a.tipo === 'albaran' && !otros.has(a.id) && (!v.titular || a.titular === v.titular)).sort(byFechaDesc);
        const mismos = albs.filter(a => !v.proveedor || norm(a.proveedor) === norm(v.proveedor) || v.albaranes.includes(a.id));
        cont.innerHTML = mismos.length ? `<div class="picker"><div class="pk-list">${mismos.map(a => `<label class="pk-f"><input type="checkbox" data-alb value="${a.id}"${v.albaranes.includes(a.id) ? ' checked' : ''}><span class="pk-n">${fmtFecha(a.fecha)} · ${esc(a.proveedor || '')}${a.numero ? ' nº ' + esc(a.numero) : ''}</span><span class="small muted">${eur(totalDoc(a), 2)}</span></label>`).join('')}</div></div>` : `<p class="note" style="margin:0">No hay albaranes sin facturar${v.proveedor ? ' de este proveedor' : ''}${v.titular ? ' en la carpeta de ' + esc(v.titular) : ''}.</p>`;
      };
      const pintarCot = () => {
        if (!esFac) return;
        const c = cotejo(v); const cont = el.querySelector('#dCot');
        if (!c.albs.length) { cont.innerHTML = '<span class="muted small">Marca los albaranes de esta factura para comprobar que cuadra.</span>'; return; }
        cont.innerHTML = `<div class="sp"><b>Comprobación</b><span class="pill ${ESTADO_COT[c.estado][1]}">${ESTADO_COT[c.estado][0]}</span></div>
          <div class="kv small"><dt>Suma de ${c.albs.length} albarán${c.albs.length === 1 ? '' : 'es'}</dt><dd>${c.conPrecio ? eur(c.sumAlb, 2) : 'sin precios'}</dd><dt>Base de la factura</dt><dd>${c.base != null ? eur(c.base, 2) : '—'}</dd>${c.dif != null ? `<dt>Diferencia</dt><dd class="${Math.abs(c.dif) > 0.5 ? 'neg' : 'pos'}">${c.dif > 0 ? '+' : ''}${eur(c.dif, 2)}</dd>` : ''}</div>
          ${c.difProd.length ? `<p class="small neg" style="margin:6px 0 0">Cantidades distintas: ${c.difProd.map(x => { const p = producto(x.pid) || {}; return `${esc(p.nombre || '¿?')} (albaranes ${fmtN(x.alb, 3)} · factura ${fmtN(x.fac, 3)} ${esc(p.unidad || '')})`; }).join('; ')}</p>` : ''}
          ${!c.conPrecio ? `<p class="small muted" style="margin:6px 0 0">Los albaranes no tienen precios: ${c.hayLineas ? 'se comparan solo las cantidades de cada producto.' : 'añade los productos de la factura para comparar cantidades.'}</p>` : c.base == null ? '<p class="small muted" style="margin:6px 0 0">Escribe la base o el total con IVA para comparar.</p>' : ''}`;
      };
      const todo = () => { leer(); pintarCot(); };
      pintarArch(); pintarLineas(); pintarAlbs(); pintarCot();
      form.addEventListener('input', e => {
        const row = e.target.closest('.linea');
        if (row && (e.target.matches('[data-c]') || e.target.matches('[data-p]'))) {
          const c = parseNum(row.querySelector('[data-c]').value, true), p = parseNum(row.querySelector('[data-p]').value);
          if (c > 0 && p > 0) row.querySelector('[data-i]').value = numStr(Math.round(c * p * 100) / 100);
        }
        if (esFac && (e.target.id === 'dBase' || e.target.id === 'dIva')) {
          const b = parseNum(el.querySelector('#dBase').value, true), iv = parseNum(el.querySelector('#dIva').value);
          if (b > 0 && iv >= 0) el.querySelector('#dTotal').value = numStr(Math.round(b * (1 + iv / 100) * 100) / 100);
        }
        todo(); pintarTotal();
      });
      form.addEventListener('change', e => {
        if (e.target.matches('[data-sube]')) {
          const files = [...e.target.files]; e.target.value = '';
          Promise.all(files.map(async f => { const pr = await prepararArchivo(f); const a = { id: uid(), nombre: pr.nombre, tipo: pr.tipo, blob: pr.blob, fecha: new Date().toISOString() }; nuevosArch.push(a); v.archivos.push(a.id); })).then(pintarArch);
          return;
        }
        if (e.target.id === 'dTit' || e.target.id === 'dProv') { leer(); pintarAlbs(); }
        todo(); if (e.target.matches('[data-alb]')) pintarCot();
      });
      form.addEventListener('click', e => {
        const ver = e.target.closest('[data-ver]'); if (ver) { const a = nuevosArch.find(x => x.id === ver.dataset.ver); if (a) { const u = URL.createObjectURL(a.blob); window.open(u, '_blank'); } else verArchivo(ver.dataset.ver); return; }
        const q = e.target.closest('[data-quitar]'); if (q) { v.archivos = v.archivos.filter(x => x !== q.dataset.quitar); quitarArch.push(q.dataset.quitar); pintarArch(); return; }
        const dl = e.target.closest('[data-dell]'); if (dl) { leer(); v.lineas.splice(+dl.dataset.dell, 1); pintarLineas(); pintarCot(); return; }
        if (e.target.closest('[data-addl]')) { leer(); hojaElegirProducto(null, p => { v.lineas.push({ productoId: p.id, cantidad: null, precio: null, importe: null }); pintarLineas(); }); }
      });
      form.addEventListener('submit', async e => {
        e.preventDefault(); leer();
        if (!v.titular) { toast('Escribe el titular: es la carpeta donde se guarda'); return; }
        if (!esFac && !v.lineas.length) { toast('Añade al menos un producto al albarán'); return; }
        if (v.lineas.some(l => !(l.cantidad > 0))) { toast('Pon la cantidad de cada producto'); return; }
        try { for (const a of nuevosArch) if (v.archivos.includes(a.id)) await archivoGuardar(a); for (const x of quitarArch) await archivoBorrar(x); }
        catch (err) { toast('No se ha podido guardar el archivo en el teléfono'); return; }
        v.lineas.forEach(l => { if (l.importe == null && l.cantidad > 0 && l.precio > 0) l.importe = Math.round(l.cantidad * l.precio * 100) / 100; });
        const x = doc || { id: uid() };
        Object.assign(x, v);
        if (!doc) db.documentos.push(x);
        if (esFac) {
          db.movimientos = db.movimientos.filter(m => m.docId !== x.id);
          if (el.querySelector('#dGasto').checked && (x.total || x.base)) {
            const tipos = (x.lineas.length ? x.lineas : db.documentos.filter(a => x.albaranes.includes(a.id)).flatMap(a => a.lineas)).map(l => (producto(l.productoId) || {}).tipo);
            const cat = { Fitosanitario: 'Fitosanitarios', Abono: 'Abono', Semilla: 'Semilla', Combustible: 'Gasóleo' }[tipos.sort((a, b) => tipos.filter(t => t === b).length - tipos.filter(t => t === a).length)[0]] || 'Otros gastos';
            db.movimientos.push({ id: uid(), docId: x.id, tipo: 'gasto', fecha: x.fecha, fincaId: '', campana: campanaDe(x.fecha || hoyISO()), categoria: cat, concepto: `Factura ${x.numero || ''} · ${x.proveedor || ''} (${x.titular})`.replace(/\s+/g, ' '), importe: x.total || x.base });
          }
        }
        guardar(); volverA(d - 1); ui.alm.titular = ui.alm.titular && ui.alm.titular !== x.titular ? x.titular : ui.alm.titular; render(); toast(esFac ? 'Factura guardada' : 'Albarán guardado');
      });
      const del = el.querySelector('[data-del]');
      if (del) del.addEventListener('click', async () => {
        const r = await confirmar('¿Borrar?', esFac ? 'Se borrará la factura y su foto o PDF. Los albaranes no se tocan.' : 'Se borrará el albarán con su foto o PDF, y sus productos dejarán de contar en el inventario.');
        if (!r.ok) return;
        for (const a of doc.archivos || []) { try { await archivoBorrar(a); } catch (er) { /* ignorar */ } }
        db.documentos = db.documentos.filter(x => x !== doc);
        db.documentos.forEach(f => { if (f.albaranes) f.albaranes = f.albaranes.filter(a => a !== doc.id); });
        db.movimientos = db.movimientos.filter(m => m.docId !== doc.id);
        guardar(); volverA(d - 1); render(); toast('Borrado');
      });
    });
  }

  // Ficha de un producto del almacén
  function formAjuste(id) {
    const a = db.ajustes.find(x => x.id === id); if (!a) return;
    const p = producto(a.productoId) || { nombre: '¿?', unidad: '' };
    formulario({
      titulo: `Ajuste de ${esc(p.nombre)}`,
      valores: a,
      campos: [
        { k: 'titular', label: 'Titular', req: 1, lista: titularesTodos },
        { k: 'cantidad', label: `Cantidad (${p.unidad})`, type: 'num', miles: 1, req: 1, hint: 'Positiva si entra en el almacén (por ejemplo semilla de vuestra cosecha); negativa si sale.' },
        { k: 'fecha', label: 'Fecha', type: 'date', req: 1 },
        { k: 'motivo', label: 'Motivo' }
      ],
      guardar(v) { if (!v.cantidad) return 'La cantidad no puede ser cero'; Object.assign(a, v); },
      textoBorrar: 'Se borrará este ajuste y las existencias volverán a calcularse sin él.',
      borrar: () => { db.ajustes = db.ajustes.filter(x => x !== a); }
    });
  }
  function hojaProducto(pid, tit) {
    const p = producto(pid); if (!p) return;
    cargarFitos().then(dd => { FITOS = dd; }).catch(() => {});
    const pintar = () => {
      const reg = regProducto(p);
      const tits = [...new Set(movimientosProducto(pid).map(m => m.titular).filter(Boolean))].sort();
      const movs = movimientosProducto(pid, tit || null);
      const misCult = [...new Set(db.fincas.map(f => cultivoDe(f.id, db.campanaActiva)).filter(Boolean))];
      const autoriz = reg ? misCult.map(cu => { const r = revisarFito(reg, cu); return r.autorizado === null ? null : `<span class="pill ${r.autorizado ? 'pos' : 'neg'}">${r.autorizado ? '✓' : '✗'} ${esc(cu)}</span>`; }).filter(Boolean).join(' ') : '';
      return `<div class="kv"><dt>Tipo</dt><dd>${esc(p.tipo)}</dd><dt>Se mide en</dt><dd>${esc(p.unidad)}</dd>${p.registro ? `<dt>Nº de registro</dt><dd>${esc(p.registro)}</dd>` : ''}${reg ? `<dt>Estado</dt><dd class="${reg.e === 'Vigente' ? 'pos' : 'neg'}">${esc(reg.e)}</dd><dt>Caduca</dt><dd>${esc(reg.cad || '—')}</dd>` : ''}</div>
        ${reg ? `<p class="small" style="margin:0">${esc(reg.f)}</p><div><div class="tag">En tus cultivos de ${db.campanaActiva}</div><div class="chips">${autoriz || '<span class="small muted">Sin cultivos asignados</span>'}</div></div><a class="btn sm" href="${urlFicha(reg.id)}" target="_blank" rel="noopener">Ver ficha oficial (PDF del Ministerio)</a>` : ''}
        <h2 class="sec">Existencias</h2><div class="list">${(tits.length ? tits : [tit || '']).map(t => `<div class="row"><div class="main"><div class="t">${esc(t || 'Sin titular')}</div></div><div class="end big ${existencias(pid, t) < 0 ? 'neg' : ''}">${fmtCant(existencias(pid, t), p.unidad)}</div></div>`).join('')}</div>
        <div class="btns" style="justify-content:flex-start"><button type="button" class="btn sm" data-ajuste>Ajustar existencias</button><button type="button" class="btn sm" data-editp>Cambiar nombre o unidad</button></div>
        <h2 class="sec">Movimientos${tit ? ' · ' + esc(tit) : ''}</h2>${movs.length ? `<p class="small muted" style="margin:0">Toca un movimiento para corregirlo o borrarlo.</p><div class="list">${movs.map(m => `<button type="button" class="row" data-mov="${m.doc ? 'doc' : m.labor ? 'labor' : 'ajuste'}" data-id="${m.doc || m.labor || m.ajuste}">${fechaBox(m.fecha)}<div class="main"><div class="t">${esc(m.txt)}</div><div class="s">${esc(m.titular || '')}</div></div><div class="end big ${m.cant < 0 ? 'neg' : 'pos'}">${m.cant > 0 ? '+' : ''}${fmtCant(m.cant, p.unidad)}</div></button>`).join('')}</div>` : '<p class="note">Sin movimientos.</p>'}
        <div class="btns" style="justify-content:flex-start;margin-top:6px"><button type="button" class="btn sm danger" data-borrarp>Borrar producto</button></div>`;
    };
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera(esc(p.nombre))}<div class="body" id="phBody">${pintar()}</div></div>`, (el, d) => {
      const body = el.querySelector('#phBody');
      cargarFitos().then(() => { body.innerHTML = pintar(); }).catch(() => {});
      el.addEventListener('click', async e => {
        const mv = e.target.closest('[data-mov]');
        if (mv) {
          volverA(d - 1);
          if (mv.dataset.mov === 'doc') formDocumento(null, mv.dataset.id);
          else if (mv.dataset.mov === 'labor') formLabor(mv.dataset.id);
          else formAjuste(mv.dataset.id);
          return;
        }
        if (e.target.closest('[data-borrarp]')) {
          const n = movimientosProducto(pid).length;
          if (n) { toast(`Tiene ${n} movimiento${n === 1 ? '' : 's'}: corrígelos o bórralos antes de borrar el producto`); return; }
          const r = await confirmar('¿Borrar el producto?', `Se quitará «${esc(p.nombre)}» del almacén.`);
          if (!r.ok) return;
          db.productos = db.productos.filter(x => x !== p); guardar(); volverA(d - 1); render(); toast('Producto borrado');
          return;
        }
        if (e.target.closest('[data-ajuste]')) {
          const ts = titularesTodos();
          formulario({
            titulo: `Ajustar ${esc(p.nombre)}`,
            valores: { titular: tit || ts[0] || '', real: null, fecha: hoyISO(), motivo: 'Recuento en la nave' },
            campos: [
              { k: 'titular', label: 'Titular', req: 1, lista: titularesTodos },
              { k: 'real', label: `Cantidad que hay de verdad (${p.unidad})`, type: 'num', req: 1, hint: 'La app apunta la diferencia con lo que calculaba.' },
              { k: 'fecha', label: 'Fecha', type: 'date', req: 1 },
              { k: 'motivo', label: 'Motivo' }
            ],
            guardar(v) { const dif = Math.round((v.real - existencias(pid, v.titular)) * 1000) / 1000; if (!dif) return 'No hay diferencia con lo que calcula la app'; db.ajustes.push({ id: uid(), productoId: pid, titular: v.titular, cantidad: dif, fecha: v.fecha, motivo: v.motivo }); },
            despues: () => { body.innerHTML = pintar(); }
          });
        }
        if (e.target.closest('[data-editp]')) {
          formulario({
            titulo: 'Producto', valores: p,
            campos: [{ k: 'nombre', label: 'Nombre', req: 1 }, { k: 'tipo', label: 'Tipo', type: 'select', half: 1, req: 1, opts: () => TIPOS_PROD.map(t => [t, t]) }, { k: 'unidad', label: 'Se mide en', type: 'select', half: 1, req: 1, opts: () => UNI_PROD.map(u => [u, u]) }],
            guardar(v) { Object.assign(p, v); },
            despues: () => { body.innerHTML = pintar(); }
          });
        }
      });
    });
  }

  // ---------- Resumen por titular y PAC ----------
  const GRUPO_CULT = { trigo: 'Cereales', cebada: 'Cereales', avena: 'Cereales', 'avena forrajera': 'Cereales', centeno: 'Cereales', triticale: 'Cereales',
    titarro: 'Leguminosas', lenteja: 'Leguminosas', garbanzo: 'Leguminosas', guisante: 'Leguminosas', yeros: 'Leguminosas', veza: 'Leguminosas',
    girasol: 'Oleaginosas', colza: 'Oleaginosas', camelina: 'Oleaginosas', barbecho: 'Barbecho',
    esparceta: 'Plurianuales', alfalfa: 'Plurianuales', chopo: 'Plurianuales', 'mezcla forestales': 'Plurianuales', 'anis dulce': 'Aromáticas', cilantro: 'Aromáticas' };
  const grupoCult = cu => cu ? (GRUPO_CULT[norm(cu)] || 'Otros') : 'Sin cultivo';
  const ORDEN_GRUPOS = ['Cereales', 'Leguminosas', 'Oleaginosas', 'Aromáticas', 'Plurianuales', 'Barbecho', 'Otros', 'Sin cultivo'];
  const HUE_GRUPO = { Cereales: 42, Leguminosas: 140, Oleaginosas: 55, Aromáticas: 95, Plurianuales: 320, Barbecho: 28, Otros: 200, 'Sin cultivo': 0 };
  const fincasTitular = t => db.fincas.filter(f => (String(f.titular || '').trim() || 'Sin titular') === t);
  const titularesResumen = () => { const ts = titulares(); return db.fincas.some(f => !String(f.titular || '').trim()) ? [...ts, 'Sin titular'] : ts; };

  function vTitulares() {
    const c = db.campanaActiva; const ts = titularesResumen();
    if (!ts.length) return vacio('Sin titulares', 'Pon el titular en cada parcela (o impórtalo del Excel) para ver aquí el resumen de cada uno.');
    const cults = [...new Set(db.fincas.map(f => cultivoDe(f.id, c) || 'Sin cultivo'))].sort((a, b) => sum(db.fincas.filter(f => (cultivoDe(f.id, c) || 'Sin cultivo') === b), f => f.superficie) - sum(db.fincas.filter(f => (cultivoDe(f.id, c) || 'Sin cultivo') === a), f => f.superficie));
    const haDe = (t, cu) => sum(fincasTitular(t).filter(f => (cultivoDe(f.id, c) || 'Sin cultivo') === cu), f => f.superficie);
    let h = `<div class="tit-cards">${ts.map(t => { const fs = fincasTitular(t); const ha = sum(fs, f => f.superficie); const st = statsGrupo(fs, c);
      const porC = cults.map(cu => ({ cu, ha: haDe(t, cu) })).filter(x => x.ha > 0);
      return `<div class="card pad tit-card"><div class="sp"><h3>${esc(t)}</h3><span class="small muted">${fs.length} parcelas</span></div>
        <div class="big-n">${fmtN(ha)} ha</div>
        ${barras(porC.map(x => ({ label: x.cu, valor: fmtN(x.ha) + ' ha', n: x.ha, h: x.cu === 'Sin cultivo' ? 0 : hue(x.cu) })), Math.max(...porC.map(x => x.ha), 1))}
        <div class="kv small"><dt>Pueblos</dt><dd>${esc([...new Set(fs.map(f => f.municipio).filter(Boolean))].join(', ') || '—')}</dd><dt>Gastos ${c}</dt><dd>${eur(st.gastos)}</dd><dt>Ingresos ${c}</dt><dd>${eur(st.ingresos)}</dd></div>
        <button class="btn sm" data-act="filtrarTitular" data-k="${esc(t)}">Ver sus parcelas</button></div>`; }).join('')}</div>`;
    h += `<h2 class="sec">Hectáreas por cultivo y titular · ${c}</h2><div class="tablewrap"><table><thead><tr><th>Cultivo</th>${ts.map(t => `<th>${esc(t)}</th>`).join('')}<th>Total</th></tr></thead><tbody>
      ${cults.map(cu => `<tr><td>${cu === 'Sin cultivo' ? chip(null) : chip(cu)}</td>${ts.map(t => `<td>${haDe(t, cu) ? fmtN(haDe(t, cu)) : '—'}</td>`).join('')}<td><b>${fmtN(sum(ts, t => haDe(t, cu)))}</b></td></tr>`).join('')}
      <tr class="tot"><td>Total</td>${ts.map(t => `<td>${fmtN(sum(fincasTitular(t), f => f.superficie))}</td>`).join('')}<td>${fmtN(sum(db.fincas, f => f.superficie))}</td></tr></tbody></table></div>
      <p class="note">Los gastos e ingresos son los apuntados en las parcelas de cada titular; los generales (sin parcela) no se reparten.</p>`;
    return h;
  }

  function datosPAC(fs, c) {
    const lab = { D: 0, C: 0, T: 0, '': 0 }; let est = 0; const grp = {};
    fs.forEach(f => { const s = siembraDe(f.id, c); const ha = +f.superficie || 0; const g = grupoCult(s && s.cultivo); grp[g] = (grp[g] || 0) + ha;
      if (s && s.cultivo && norm(s.cultivo) !== 'barbecho') { const k = LABOREO[s.laboreo] ? s.laboreo : ''; lab[k] += ha; }
      if (s && s.estiercol) est += ha; });
    const sembrada = lab.D + lab.C + lab.T + lab[''];
    return { lab, est, grp, sembrada, total: sum(fs, f => +f.superficie || 0) };
  }
  const pct = (a, b) => b ? `${fmtN(a / b * 100, 1)} %` : '—';
  function vPAC() {
    const c = db.campanaActiva; const ant = campanaAnterior(c); const ts = titularesResumen();
    const filas = [...ts.map(t => [t, fincasTitular(t)]), ['Total', db.fincas]];
    let h = `<div class="card pad" style="display:flex;flex-direction:column;gap:10px"><p style="margin:0">Datos para preparar la PAC y el seguro de la campaña <b>${c}</b>, por titular.</p>
      <button class="btn primary" data-act="excelPAC">Descargar Excel PAC y seguro</button>
      <p class="small muted" style="margin:0">Con el formato de tu hoja: titular, nombre, término, polígono, parcela, recinto, superficie y, por campaña, cultivo, basura y laboreo. Una hoja con todo y otra por titular.</p></div>`;
    h += `<h2 class="sec">Laboreo de la superficie sembrada · ${c}</h2><div class="tablewrap"><table><thead><tr><th>Titular</th><th>Siembra directa</th><th>Directa con complemento</th><th>Tradicional</th><th>Sin dato</th><th>Estiércol</th></tr></thead><tbody>
      ${filas.map(([t, fs]) => { const d = datosPAC(fs, c); return `<tr${t === 'Total' ? ' class="tot"' : ''}><td>${esc(t)}</td>${['D', 'C', 'T', ''].map(k => `<td>${fmtN(d.lab[k])} ha<br><span class="small muted">${pct(d.lab[k], d.sembrada)}</span></td>`).join('')}<td>${fmtN(d.est)} ha</td></tr>`; }).join('')}</tbody></table></div>`;
    h += `<h2 class="sec">Superficie por tipo de cultivo · ${c}</h2><div class="tablewrap"><table><thead><tr><th>Titular</th>${ORDEN_GRUPOS.map(g => `<th>${g}</th>`).join('')}</tr></thead><tbody>
      ${filas.map(([t, fs]) => { const d = datosPAC(fs, c); return `<tr${t === 'Total' ? ' class="tot"' : ''}><td>${esc(t)}</td>${ORDEN_GRUPOS.map(g => `<td>${d.grp[g] ? `${fmtN(d.grp[g])}<br><span class="small muted">${pct(d.grp[g], d.total)}</span>` : '—'}</td>`).join('')}</tr>`; }).join('')}</tbody></table></div>`;
    // Rotación
    const repiten = db.fincas.filter(f => { const a = cultivoDe(f.id, c), b = cultivoDe(f.id, ant); return a && b && norm(a) === norm(b) && grupoCult(a) !== 'Plurianuales'; });
    const sinCult = db.fincas.filter(f => !cultivoDe(f.id, c));
    h += `<h2 class="sec">Rotación · ${ant} → ${c}</h2><div class="card pad" style="display:flex;flex-direction:column;gap:8px">
      <p style="margin:0">${repiten.length ? `<b class="neg">${repiten.length} parcelas repiten cultivo</b> (${fmtN(sum(repiten, f => f.superficie))} ha). No cuenta las plurianuales (esparceta, alfalfa…).` : '<b class="pos">Ninguna parcela repite cultivo</b> respecto a la campaña anterior.'}</p>
      ${repiten.length ? `<div class="tablewrap"><table><thead><tr><th>Parcela</th><th>Titular</th><th>Cultivo</th><th>ha</th></tr></thead><tbody>${repiten.sort(byNombre).slice(0, 80).map(f => `<tr><td><button class="lnk" data-act="verFinca" data-id="${f.id}">${esc(f.nombre)}</button></td><td style="text-align:left">${esc(f.titular || '')}</td><td style="text-align:left">${chip(cultivoDe(f.id, c))}</td><td>${fmtN(f.superficie)}</td></tr>`).join('')}</tbody></table></div>${repiten.length > 80 ? `<p class="note">…y ${repiten.length - 80} más.</p>` : ''}` : ''}
      ${sinCult.length ? `<p class="small" style="margin:0"><b>${sinCult.length} parcelas sin cultivo asignado</b> en ${c} (${fmtN(sum(sinCult, f => f.superficie))} ha).</p>` : ''}</div>
      <p class="note">Estos datos te ayudan a revisar los ecorregímenes (rotación, siembra directa…). Los porcentajes mínimos y condiciones cambian según el año: confírmalos con tu técnico o en la normativa de la PAC.</p>`;
    return h;
  }

  async function excelPAC() {
    let XLSX;
    try { XLSX = await cargarXLSX(); } catch (e) { toast('No se ha podido cargar el generador de Excel. Abre la app una vez con cobertura.'); return; }
    const camps = [...new Set(db.siembras.map(s => s.campana))].sort();
    const cab1 = ['', '', '', '', '', '', '', '', '', '']; const cab2 = ['PROPIETARIO', 'NOMBRE', 'TERMINO', 'PROVINCIA', 'MUNICIPIO', 'POLIGONO', 'PARCELA', 'RECINTO', 'SUPERFICIE', 'SUPERFICIE DECLARABLE PAC'];
    camps.forEach(c => { cab1.push(c, '', ''); cab2.push('CULTIVO', 'BASURA', 'LABOREO'); });
    const nombreDe = f => { if (grupoDe(f)) return grupoDe(f); const nb = nombreBase(f.nombre); return (!nb || norm(nb) === norm(f.municipio)) ? '' : nb; };
    const fila = f => { const r = [String(f.titular || '').toUpperCase(), nombreDe(f), String(f.municipio || '').toUpperCase(), f.provincia || '', f.codMunicipio || '', +f.poligono || f.poligono || '', +f.parcela || f.parcela || '', f.recinto || '', f.geo && f.geo.ha ? f.geo.ha : '', f.superficie || ''];
      camps.forEach(c => { const s = siembraDe(f.id, c); r.push(s ? String(s.cultivo).toUpperCase() : '', s && s.estiercol ? 'X' : '', s ? s.laboreo || '' : ''); }); return r; };
    const orden = (a, b) => String(a.municipio || '').localeCompare(String(b.municipio || ''), 'es') || (+a.poligono - +b.poligono) || (+a.parcela - +b.parcela);
    const hoja = fs => { const ws = XLSX.utils.aoa_to_sheet([cab1, cab2, ...fs.slice().sort(orden).map(fila)]); ws['!cols'] = cab2.map((h, i) => ({ wch: [12, 22, 16, 9, 9, 9, 9, 9, 11, 13][i] || 11 })); ws['!merges'] = camps.map((c, i) => ({ s: { r: 0, c: 10 + i * 3 }, e: { r: 0, c: 12 + i * 3 } })); return ws; };
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, hoja(db.fincas), 'PAC');
    titularesResumen().forEach(t => XLSX.utils.book_append_sheet(wb, hoja(fincasTitular(t)), String(t).toUpperCase().slice(0, 31)));
    const c = db.campanaActiva; const ts = titularesResumen();
    const cults = [...new Set(db.fincas.map(f => cultivoDe(f.id, c) || 'Sin cultivo'))].sort();
    const res = [[`Resumen campaña ${c}`], [], ['CULTIVO', ...ts.map(t => t.toUpperCase()), 'TOTAL']];
    cults.forEach(cu => { const v = ts.map(t => Math.round(sum(fincasTitular(t).filter(f => (cultivoDe(f.id, c) || 'Sin cultivo') === cu), f => f.superficie) * 10000) / 10000); res.push([cu.toUpperCase(), ...v, Math.round(sum(v, x => x) * 10000) / 10000]); });
    res.push([], ['LABOREO (ha sembradas)', ...ts.map(t => t.toUpperCase())]);
    [['D', 'Siembra directa'], ['C', 'Directa con complemento'], ['T', 'Tradicional'], ['', 'Sin dato']].forEach(([k, n]) => res.push([n, ...ts.map(t => Math.round(datosPAC(fincasTitular(t), c).lab[k] * 10000) / 10000)]));
    res.push(['Con estiércol', ...ts.map(t => Math.round(datosPAC(fincasTitular(t), c).est * 10000) / 10000)]);
    const wr = XLSX.utils.aoa_to_sheet(res); wr['!cols'] = [{ wch: 26 }, ...ts.map(() => ({ wch: 12 })), { wch: 12 }];
    XLSX.utils.book_append_sheet(wb, wr, 'Resumen');
    const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const ok = await guardarArchivo(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), `PAC_${c.replace('/', '-')}_${hoyISO()}.xlsx`, true, 'PAC y seguro');
    if (ok) toast('Excel de PAC y seguro creado');
  }

  // ---------- Asistente de voz: entiende frases como «acabo de sembrar la Sarrana de trigo a 180 kilos» ----------
  // Todo se hace en el teléfono, sin inteligencia artificial: palabras clave + nombres de tus fincas, cultivos y productos.
  function normVoz(s) {
    let t = String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
    t = t.replace(/(\d)\.(\d{3})(?!\d)/g, '$1$2').replace(/(\d)\.(\d{3})(?!\d)/g, '$1$2'); // 20.000 -> 20000
    t = t.replace(/(\d),(\d{3})(?=\s*(?:kilos?|kg|kilogramos?)\b)/g, '$1$2');             // 47,620 kilos -> 47620
    t = t.replace(/(\d),(\d)/g, '$1.$2');                                                  // 1,5 -> 1.5
    t = t.replace(/€/g, ' euros ').replace(/%/g, ' por ciento ').replace(/º|ª/g, ' ');
    t = t.replace(/[^a-z0-9.\/ ]+/g, ' ').replace(/(^|[^\d])\.|\.(?!\d)/g, '$1 ');
    return cifras(t.replace(/\s+/g, ' ').trim());
  }
  const NUM_PAL = { cero: 0, un: 1, uno: 1, una: 1, dos: 2, tres: 3, cuatro: 4, cinco: 5, seis: 6, siete: 7, ocho: 8, nueve: 9, diez: 10, once: 11, doce: 12, trece: 13, catorce: 14, quince: 15, dieciseis: 16, diecisiete: 17, dieciocho: 18, diecinueve: 19, veinte: 20, veintiun: 21, veintiuno: 21, veintiuna: 21, veintidos: 22, veintitres: 23, veinticuatro: 24, veinticinco: 25, veintiseis: 26, veintisiete: 27, veintiocho: 28, veintinueve: 29, treinta: 30, cuarenta: 40, cincuenta: 50, sesenta: 60, setenta: 70, ochenta: 80, noventa: 90, cien: 100, ciento: 100, doscientos: 200, doscientas: 200, trescientos: 300, trescientas: 300, cuatrocientos: 400, cuatrocientas: 400, quinientos: 500, quinientas: 500, seiscientos: 600, seiscientas: 600, setecientos: 700, setecientas: 700, ochocientos: 800, ochocientas: 800, novecientos: 900, novecientas: 900 };
  // «ciento ochenta y cinco» -> 185 · «dos mil quinientos» -> 2500 · «litro y medio» -> 1.5 · «uno coma cinco» -> 1.5
  function cifras(t) {
    const w = t.split(' '); const out = []; let i = 0;
    const esNum = x => x in NUM_PAL || /^\d+(\.\d+)?$/.test(x);
    while (i < w.length) {
      const x = w[i];
      if ((x in NUM_PAL && !(['un', 'una', 'uno'].includes(x) && !esNum(w[i + 1] || '') && w[i + 1] !== 'mil' && !/^(kilo|litro|euro|hectarea|tonelada|gramo|centimetro|mililitro)/.test(w[i + 1] || ''))) || (x === 'mil' && out.length && /^\d/.test(out[out.length - 1]))) {
        let total = 0, cur = 0, j = i;
        if (x === 'mil') { total = +out.pop() * 1000; j++; }
        while (j < w.length) {
          const y = w[j];
          if (y in NUM_PAL) { cur += NUM_PAL[y]; j++; }
          else if (y === 'mil') { total += (cur || 1) * 1000; cur = 0; j++; }
          else if (y === 'y' && NUM_PAL[w[j + 1]] < 10) j++;
          else break;
        }
        let val = total + cur;
        if (w[j] === 'coma' && w[j + 1] != null && esNum(w[j + 1])) { val = parseFloat(val + '.' + (w[j + 1] in NUM_PAL ? NUM_PAL[w[j + 1]] : w[j + 1])); j += 2; }
        out.push(String(val)); i = j;
      } else if (/^\d+(\.\d+)?$/.test(x) && w[i + 1] === 'mil') { out.push(String(+x * 1000)); i += 2; }
      else if (/^\d+$/.test(x) && w[i + 1] === 'coma' && /^\d+$/.test(w[i + 2] || '')) { out.push(x + '.' + w[i + 2]); i += 3; }
      else if ((x === 'medio' || x === 'media') && !/^\d/.test(out[out.length - 1] || '')) { out.push('0.5'); i++; }
      else if (x === 'y' && (w[i + 1] === 'medio' || w[i + 1] === 'media') && /^\d/.test(out[out.length - 1] || '')) { out[out.length - 1] = String(+out[out.length - 1] + 0.5); i += 2; }
      else { out.push(x); i++; }
    }
    return out.join(' ');
  }
  // Parecido entre palabras (para cuando el dictado escribe «serrana» en vez de «Sarrana»)
  function lev(a, b) {
    if (a === b) return 0; const m = a.length, n = b.length; if (!m || !n) return m || n;
    let prev = Array.from({ length: n + 1 }, (_, k) => k);
    for (let i = 1; i <= m; i++) { const cur = [i]; for (let k = 1; k <= n; k++) cur[k] = Math.min(prev[k] + 1, cur[k - 1] + 1, prev[k - 1] + (a[i - 1] === b[k - 1] ? 0 : 1)); prev = cur; }
    return prev[n];
  }
  const parecido = (a, b) => 1 - lev(a, b) / Math.max(a.length, b.length, 1);
  const PALABRAS_VACIAS = new Set(['la', 'el', 'los', 'las', 'de', 'del', 'en', 'a', 'al', 'y', 'con', 'que', 'he', 'ha', 'hemos', 'acabo', 'acabamos', 'hoy', 'ayer', 'por', 'para', 'finca', 'fincas', 'parcela', 'parcelas', 'toda', 'todas', 'todo']);
  // Busca en la frase los nombres de una lista: [{n: 'sarrana clemente', ...}]
  function buscarNombres(texto, cands, umbral = 0.82) {
    const w = texto.split(' '); const hits = [];
    for (const c of cands) {
      if (!c.n) continue;
      const k = c.n.split(' ').length; let best = 0, pos = -1;
      for (let i = 0; i + k <= w.length; i++) {
        if (PALABRAS_VACIAS.has(w[i]) && !c.n.startsWith(w[i] + ' ')) continue;
        const s = parecido(w.slice(i, i + k).join(' '), c.n); if (s > best) { best = s; pos = i; }
      }
      const min = c.n.length <= 4 ? 1 : c.n.length <= 6 ? 0.83 : umbral;
      if (best >= min) hits.push({ c, score: best, i: pos, k });
    }
    hits.sort((a, b) => b.k - a.k || b.score - a.score || b.c.n.length - a.c.n.length);
    const usados = new Set(); const out = [];
    for (const h of hits) { const r = Array.from({ length: h.k }, (_, x) => x + h.i); if (r.some(x => usados.has(x))) continue; r.forEach(x => usados.add(x)); out.push(h); }
    return out.sort((a, b) => a.i - b.i);
  }
  const quitarTramos = (texto, hits) => { const w = texto.split(' '); hits.forEach(h => { for (let x = h.i; x < h.i + h.k; x++) w[x] = '_'; }); return w.join(' '); };

  const DIAS = ['domingo', 'lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado'];
  const MESES_V = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  function fechaDicha(t) {
    const d = new Date(); const iso = x => new Date(x - x.getTimezoneOffset() * 60000).toISOString().slice(0, 10);
    if (/\banteayer\b/.test(t)) { d.setDate(d.getDate() - 2); return iso(d); }
    if (/\bayer\b/.test(t)) { d.setDate(d.getDate() - 1); return iso(d); }
    let m = t.match(/\b(\d{1,2}) de (enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)(?: de (\d{4}))?\b/);
    if (m) { const mes = m[2] === 'setiembre' ? 8 : MESES_V.indexOf(m[2]); let y = m[3] ? +m[3] : d.getFullYear(); let f = new Date(y, mes, +m[1]); if (!m[3] && f > d) f = new Date(y - 1, mes, +m[1]); return iso(f); }
    m = t.match(/\b(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b/);
    if (m) { let y = m[3] ? (+m[3] < 100 ? 2000 + +m[3] : +m[3]) : d.getFullYear(); let f = new Date(y, +m[2] - 1, +m[1]); if (!m[3] && f > d) f = new Date(y - 1, +m[2] - 1, +m[1]); return iso(f); }
    m = t.match(/\b(lunes|martes|miercoles|jueves|viernes|sabado|domingo)\b/);
    if (m) { let diff = (d.getDay() - DIAS.indexOf(m[1]) + 7) % 7; if (diff === 0) diff = 7; d.setDate(d.getDate() - diff); return iso(d); }
    return hoyISO();
  }

  const VERBOS = [
    ['Siembra', /\b(sembr|siembr)\w*/],
    ['Cosecha', /\b(cosech|segad|trill|recolect)\w*/],
    ['Tratamiento', /\b(sulfat|fumig|herbicid|fungicid|insecticid|pulveriz|tratad|trate|tratar|tratamiento|tratando)\w*/],
    ['Abonado', /\b(abon|cobertera|purin|estercol)\w*/],
    ['Rodillo', /\b(rodillo|rular|rulado|rule|rulao|rulo)\b/],
    ['Empacado', /\b(empac|alpac|empaque)\w*/],
    ['Labrado', /\b(are|aro|arado|arando|arar|aramos|arao|labr\w*|cultivador|grade\w*|gradas?|vertedera|chisel|subsol\w*|binad\w*|alzad\w*|alzar|desbroz\w*|escarific\w*)\b/]
  ];
  const VERBO_NEUTRO = /\b(eche|echado|echar|echamos|echao|aplicad\w*|aplique|aplicar|dado|di|dar|metido|meti)\b/;
  const UNIDAD_RE = '(kilogramos?|kilos?|kg|toneladas?|tn|t|litros?|l|centimetros cubicos|centimetros?|cc|centilitros?|mililitros?|ml|gramos?|gr|g)';
  function cantidadesDichas(t) {
    const out = []; const re = new RegExp(`\\b(\\d+(?:\\.\\d+)?) ${UNIDAD_RE}\\b(?: (?:por|la|por la|cada|/) ?(hectarea|ha))?`, 'g');
    let m;
    while ((m = re.exec(t))) {
      const u = m[2]; let base, q = +m[1];
      if (/^(kilo|kg)/.test(u)) base = 'kg';
      else if (/^(tonelada|tn|t)$|^tonelada/.test(u)) { base = 'kg'; q *= 1000; }
      else if (/^(litro|l)$|^litro/.test(u)) base = 'L';
      else if (/^centilitro/.test(u)) { base = 'cc'; q *= 10; }
      else if (/^(centimetro|cc|mililitro|ml)/.test(u)) base = 'cc';
      else base = 'g';
      out.push({ q, base, ha: !!m[3], i: m.index, txt: m[0] });
    }
    return out;
  }
  const dineroDicho = t => { const m = t.match(/\b(\d+(?:\.\d+)?) euros?\b/); return m ? +m[1] : null; };

  function candidatosFinca() {
    const m = new Map();
    const add1 = (n, fs, tipo, nombre) => { if (!n || n.length < 3) return; const o = m.get(n) || { n, fs: new Set(), tipo, nombre }; fs.forEach(f => o.fs.add(f)); m.set(n, o); };
    // «El Llano» también se encuentra si dices solo «Llano»
    const add = (n, fs, tipo, nombre) => { n = normVoz(n); add1(n, fs, tipo, nombre); const sin = n.replace(/^(el|la|los|las) /, ''); if (sin !== n) add1(sin, fs, tipo, nombre); };
    listaGrupos().forEach((fs, g) => { add(g, fs, 'g', g); const sin = g.replace(/\s*\(.*\)\s*$/, ''); if (sin !== g) add(sin, fs, 'g', g); });
    db.fincas.forEach(f => {
      const b = nombreBase(f.nombre);
      if (!grupoDe(f) && b && norm(b) !== norm(f.municipio || '') && norm(b) !== 'parcela') add(b, [f], 'f', b);
      if (!grupoDe(f) && f.nombre && !/\d+\s*-\s*\d+/.test(f.nombre)) add(f.nombre, [f], 'f', f.nombre);
    });
    return [...m.values()].map(o => Object.assign(o, { fs: [...o.fs] }));
  }
  const candidatosTitular = () => titularesTodos().map(t => ({ n: normVoz(t), t }));
  const candidatosCultivo = () => db.cultivos.map(c => ({ n: normVoz(c), c })).concat([{ n: 'lentejas', c: 'Lenteja' }, { n: 'anis', c: 'Anís dulce' }, { n: 'garbanzos', c: 'Garbanzo' }].filter(x => db.cultivos.includes(x.c)));
  function candidatosProducto(tipo) {
    const out = [];
    db.productos.filter(p => !tipo || p.tipo === tipo).forEach(p => {
      out.push({ n: normVoz(p.nombre), p });
      const w1 = normVoz(p.nombre).split(' ')[0]; if (w1.length >= 5 && w1 !== normVoz(p.nombre) && !/^(semilla|trigo|cebada|avena)$/.test(w1)) out.push({ n: w1, p });
    });
    return out;
  }
  // Parcelas dichas por número: «la parcela 180», «polígono 663 parcela 180»
  function parcelasPorNumero(t, dentro) {
    const out = []; const re = /\b(?:poligono (\d+) )?parcelas? (\d+)(?: (?:recinto|rec) (\d+))?/g; let m;
    while ((m = re.exec(t))) { const po = m[1] ? +m[1] : null, pa = +m[2];
      const fs = (dentro && dentro.length ? dentro : db.fincas).filter(f => +f.parcela === pa && (po == null || +f.poligono === po));
      out.push({ po, pa, fs, i: m.index }); }
    return out;
  }

  function interpretar(texto) {
    const t = normVoz(texto);
    const R = { texto, t, intencion: '', fecha: fechaDicha(t), dudas: [], avisos: [] };
    if (!t) return R;
    if (/\b(nueva finca|nueva parcela|finca nueva|parcela nueva|dar de alta|da de alta|dame de alta|alta (de )?(la |una )?(finca|parcela)|anad\w* (la |una )?(finca|parcela)|crea\w* (la |una )?(finca|parcela))\b/.test(t)) return interpretarFinca(t, R);
    if (/\balbaran/.test(t)) return interpretarDocumento(t, R, 'albaran');
    if (/\b(compra|compro|compre|comprado|compramos|hemos comprado|he comprado)\b/.test(t) && cantidadesDichas(t).length && !VERBOS.some(([, re]) => re.test(t))) return interpretarDocumento(t, R, 'albaran');
    if (/\bfactura/.test(t) && !/\b(sembr|sulfat|arad|labr)/.test(t)) return interpretarDocumento(t, R, 'factura');
    if (/\bsemilla|para sembrar\b/.test(t) && /\b(cosecha|propia|nuestra|guard|apart|reserv|dej)\w*/.test(t) && !/\b(he sembrado|sembre|acabo de sembrar)\b/.test(t)) return interpretarSemilla(t, R);
    const mq = maquinasDichas(t); R.maquinas = mq.ms; R.equipo = mq.equipo || null;
    const tv = mq.t; // sin los nombres de las máquinas («la sulfatadora» no es «he sulfatado»)
    const verbo = VERBOS.map(([tipo, re]) => { const m = tv.match(re); return m ? { tipo, i: m.index } : null; }).filter(Boolean).sort((a, b) => a.i - b.i);
    if (verbo.length && verbo[0].tipo === 'Cosecha' && cantidadesDichas(tv).some(c => c.base === 'kg')) return interpretarCosecha(tv, R);
    if (!verbo.length && /\b(vendid\w*|venta|vend[io]|vendemos|cobrad\w*|cobre|ingreso|me han pagado|me han ingresado|pago de la pac)\b/.test(t)) return interpretarMov(t, R, 'ingreso');
    if (!verbo.length && (/\b(gasto|he pagado|pagad\w*|pague|pagamos|compra|comprado)\b/.test(t) || dineroDicho(t) != null)) return interpretarMov(t, R, 'gasto');
    if (verbo.length || VERBO_NEUTRO.test(t)) return interpretarLabor(tv, R, verbo);
    R.dudas.push('No sé qué quieres apuntar. Empieza con lo que has hecho: «he sembrado…», «he sulfatado…», «albarán de…», «gasto de…».');
    return R;
  }

  // Fincas de la frase: nombres, números de parcela, «las de trigo», «las de Tomás» y lo que va detrás de «excepto»
  function fincasDichas(t, opts = {}) {
    const corte = t.search(/\b(excepto|menos|salvo|sin contar|quitando|pero no|no en)\b/);
    let inc = corte >= 0 ? t.slice(0, corte) : t; const exc = corte >= 0 ? t.slice(corte) : '';
    if (opts.quitar) inc = opts.quitar(inc);
    const nombres = buscarNombres(inc, candidatosFinca());
    let fs = [...new Set(nombres.flatMap(h => h.c.fs))];
    const porNum = parcelasPorNumero(inc, fs.length ? null : null);
    porNum.forEach(p => { if (p.fs.length === 1 || (p.fs.length > 1 && fs.length)) { const dentro = fs.length ? p.fs.filter(f => fs.includes(f)) : p.fs; (dentro.length ? dentro : p.fs).forEach(f => { if (!fs.includes(f)) fs.push(f); }); } else if (p.fs.length > 1) opts.dudas && opts.dudas.push(`Hay ${p.fs.length} parcelas con el número ${p.pa}: di también el polígono o el nombre de la finca.`); else opts.dudas && opts.dudas.push(`No tengo ninguna parcela ${p.po ? p.po + '-' : ''}${p.pa}.`); });
    const titInc = buscarNombres(quitarTramos(inc, nombres), candidatosTitular(), 0.85).map(h => h.c.t);
    const cultInc = opts.cultivoFiltro || null;
    if (!fs.length && (cultInc || titInc.length)) fs = db.fincas.filter(f => (!cultInc || norm(cultivoDe(f.id, db.campanaActiva) || '') === norm(cultInc)) && (!titInc.length || titInc.includes(String(f.titular || '').trim())));
    else if (fs.length && titInc.length && /\b(parte|las de|la de|lo de)\b/.test(inc)) fs = fs.filter(f => titInc.includes(String(f.titular || '').trim()));
    const quitadas = [];
    if (exc) {
      const exN = buscarNombres(exc, candidatosFinca()).flatMap(h => h.c.fs);
      const exT = buscarNombres(exc, candidatosTitular(), 0.85).map(h => h.c.t);
      const exP = parcelasPorNumero(exc, fs).flatMap(p => p.fs);
      fs = fs.filter(f => { const q = exN.includes(f) || exP.includes(f) || exT.includes(String(f.titular || '').trim()); if (q) quitadas.push(f); return !q; });
      if (/\brecinto\b/.test(exc)) opts.dudas && opts.dudas.push('No puedo quitar un recinto suelto de una parcela: quita la parcela entera o ajusta la superficie al revisar.');
      if (!quitadas.length) opts.dudas && opts.dudas.push('No he encontrado lo que dices después de «excepto».');
    }
    return { fs, quitadas, nombres: nombres.map(h => h.c.nombre), titulares: titInc };
  }

  function interpretarLabor(t, R, verbo) {
    R.intencion = 'labor';
    let tipo = verbo.length ? verbo[0].tipo : '';
    if (verbo.length > 1 && verbo[1].tipo !== tipo) R.dudas.push(`Has dicho varias labores; apunto «${tipo}». Dicta las otras por separado.`);
    // producto del almacén
    const tipoProd = TIPO_DE_LABOR[tipo] || null;
    let prodHits = buscarNombres(t, candidatosProducto(tipoProd));
    if (!prodHits.length && !tipoProd) prodHits = buscarNombres(t, candidatosProducto(null));
    let p = prodHits.length ? prodHits[0].c.p : null;
    if (!tipo && p) tipo = { Fitosanitario: 'Tratamiento', Abono: 'Abonado', Semilla: 'Siembra' }[p.tipo] || '';
    if (!tipo && /\b(urea|nitrato|abono|npk|nac|estiercol|basura|fosfato|potasa)\b/.test(t)) tipo = 'Abonado';
    if (!tipo) { R.dudas.push('No sé qué labor es: di «sembrado», «arado», «sulfatado», «abonado», «rodillo», «cosechado»…'); tipo = 'Otro'; }
    let resto = quitarTramos(t, prodHits);
    // cultivo
    const cultHits = buscarNombres(resto, candidatosCultivo(), 0.85);
    const cultivo = cultHits.length ? cultHits[0].c.c : '';
    resto = quitarTramos(resto, cultHits);
    const { fs, quitadas, nombres } = fincasDichas(resto, { dudas: R.dudas, cultivoFiltro: tipo !== 'Siembra' && /\b(las|los|todas?|todos?) (de|del) /.test(t) ? cultivo : null });
    if (!fs.length) R.dudas.push('No sé en qué finca: di su nombre (por ejemplo «la Sarrana») o «parcela 180».');
    // semilla: si no la has dicho, la del cultivo que tengas en el almacén
    if (tipo === 'Siembra' && !p && cultivo) {
      const tits = [...new Set(fs.map(f => titularDe(f.id)))];
      const sem = db.productos.filter(x => x.tipo === 'Semilla' && norm(x.nombre).includes(norm(cultivo))).map(x => ({ x, ex: sum(tits, ti => Math.max(0, existencias(x.id, ti))) })).sort((a, b) => b.ex - a.ex);
      if (sem.length) p = sem[0].x;
      else R.avisos.push(`No hay semilla de ${cultivo.toLowerCase()} en el almacén: la siembra se apunta sin descontar nada. Puedes darla de alta con un albarán o diciendo «guardar … kilos de ${cultivo.toLowerCase()} de la cosecha como semilla».`);
    }
    if (tipo === 'Tratamiento' && !p) {
      const palabras = resto.split(' ').filter(w => w.length >= 5 && w !== '_');
      const reg = FITOS ? FITOS.productos.find(r => palabras.some(w => parecido(normVoz(r.n).split(' ')[0], w) > 0.85)) : null;
      R.bloqueo = reg ? `«${reg.n}» no está en el almacén. Sube primero su albarán para poder usarlo.` : 'No encuentro el fitosanitario en el almacén. Di su nombre tal como está en el almacén, o súbelo con su albarán.';
    }
    // dosis
    const cant = cantidadesDichas(t);
    let dosis = null, unidad = null;
    if (cant.length) {
      const c = cant.find(x => x.ha) || cant[0];
      const total = /\b(en total|total|entre todo|en todas)\b/.test(t);
      if (total) { dosis = c.q; unidad = c.base === 'L' ? 'L' : c.base === 'kg' ? 'kg' : c.base === 'cc' ? 'L' : 'kg'; if (c.base === 'cc') dosis = c.q / 1000; if (c.base === 'g') dosis = c.q / 1000; }
      else { dosis = c.q; unidad = { kg: 'kg/ha', L: 'L/ha', cc: 'cc/ha', g: 'g/ha' }[c.base]; }
      if (!CON_DOSIS.includes(tipo)) { dosis = null; unidad = null; }
    }
    const haM = t.match(/\b(\d+(?:\.\d+)?) (?:hectareas?|ha)\b/);
    const motivo = (t.match(/\bcontra (?:la |el |los |las )?([a-z ]{3,40}?)(?= a \d| con | en | por | y |$|\d)/) || [])[1] || '';
    const variedad = (t.match(/\bvariedad ([a-z0-9]+(?: [a-z0-9]+)?)/) || [])[1] || '';
    const mqs = R.maquinas || []; const dm = R.equipo ? { equipoId: R.equipo.id, tractorId: R.equipo.tractorId || '', aperoId: R.equipo.aperoId || '' } : maqPorDefecto(tipo);
    const tractorV = (mqs.find(m => MOTRICES.includes(m.tipo)) || {}).id || dm.tractorId;
    const aperoV = (mqs.find(m => !MOTRICES.includes(m.tipo)) || {}).id || (mqs.length ? '' : dm.aperoId);
    const equipoV = dm.equipoId && tractorV === dm.tractorId && aperoV === dm.aperoId ? dm.equipoId : '';
    const apero = !p && !mqs.length && !CON_DOSIS.includes(tipo) ? ((t.match(/\bcon (?:el |la |los |las )?([a-z]+(?: [a-z]+)?)/) || [])[1] || '') : '';
    R.labor = {
      fecha: R.fecha, tipo, fincas: fs.map(f => f.id), prodAlm: p ? p.id : '', producto: apero ? apero.charAt(0).toUpperCase() + apero.slice(1) : '',
      dosis, unidad: unidad || (tipo === 'Siembra' || tipo === 'Abonado' ? 'kg/ha' : 'L/ha'), motivo: motivo.trim(), superficie: fs.length === 1 && haM ? +haM[1] : null,
      coste: dineroDicho(t), campana: db.campanaActiva, notas: '', equipoId: equipoV, tractorId: tractorV, aperoId: aperoV, aplicador: tipo === 'Tratamiento' ? aplicadorPorDefecto() : '', cultivoSiembra: tipo === 'Siembra' ? cultivo : '', variedadSiembra: variedad ? variedad.charAt(0).toUpperCase() + variedad.slice(1) : ''
    };
    R.quitadas = quitadas; R.nombresFinca = nombres;
    if (tipo === 'Siembra' && !cultivo) R.dudas.push('No has dicho el cultivo: la siembra se apunta sin cambiar el cultivo de las parcelas.');
    if (CON_DOSIS.includes(tipo) && dosis == null && p) R.dudas.push('No has dicho la dosis: no se descuenta nada del almacén.');
    return R;
  }

  function interpretarCosecha(t, R) {
    R.intencion = 'cosecha';
    const { fs, quitadas, nombres } = fincasDichas(t, { dudas: R.dudas });
    const c = cantidadesDichas(t).find(x => x.base === 'kg');
    const hum = t.match(/\bhumedad (?:de |del )?(\d+(?:\.\d+)?)/);
    const cos = (R.maquinas || []).find(m => MOTRICES.includes(m.tipo));
    R.cosecha = { fecha: R.fecha, fincas: fs.map(f => f.id), kg: c ? c.q : null, humedad: hum ? +hum[1] : null, campana: db.campanaActiva, notas: '', maquinaId: cos ? cos.id : maqPorDefecto('Cosecha').tractorId };
    R.quitadas = quitadas; R.nombresFinca = nombres;
    if (!fs.length) R.dudas.push('No sé de qué finca es la cosecha.');
    return R;
  }

  function interpretarSemilla(t, R) {
    R.intencion = 'semilla';
    const cult = buscarNombres(t, candidatosCultivo(), 0.85); const tit = buscarNombres(t, candidatosTitular(), 0.85);
    const c = cantidadesDichas(t).find(x => x.base === 'kg');
    R.semilla = { cultivo: cult.length ? cult[0].c.c : '', titular: tit.length ? tit[0].c.t : (titularesTodos().length === 1 ? titularesTodos()[0] : ''), kg: c ? c.q : null, fecha: R.fecha };
    if (!R.semilla.cultivo) R.dudas.push('¿De qué cultivo es la semilla?');
    if (!R.semilla.titular) R.dudas.push('¿Para qué titular? Di «para Iván», «de Tomás»…');
    if (!R.semilla.kg) R.dudas.push('¿Cuántos kilos?');
    return R;
  }

  function interpretarMov(t, R, tipo) {
    R.intencion = tipo;
    const cat = tipo === 'ingreso'
      ? (/\b(pac|ayuda|subvencion|pago basico)\b/.test(t) ? 'Subvención PAC' : /\bpaja\b/.test(t) ? 'Venta de paja' : /\b(trigo|cebada|avena|grano|titarro|lenteja|girasol|colza|centeno)\b/.test(t) ? 'Venta de grano' : 'Otros ingresos')
      : (/\b(gasoleo|gasoil|combustible)\b/.test(t) ? 'Gasóleo' : /\bseguro\b/.test(t) ? 'Seguro' : /\b(arriendo|arrendamiento|renta)\b/.test(t) ? 'Arrendamiento' : /\b(reparacion|taller|rueda|recambio|tractor|maquina|averia|itv)\w*/.test(t) ? 'Maquinaria y reparaciones' : /\bsemilla/.test(t) ? 'Semilla' : /\b(abono|urea|nitrato)\b/.test(t) ? 'Abono' : /\b(herbicida|fungicida|insecticida|fitosanitario)\w*/.test(t) ? 'Fitosanitarios' : /\b(cosechadora|empacadora|transporte|servicio)\w*/.test(t) ? 'Servicios contratados' : 'Otros gastos');
    const kg = cantidadesDichas(t).find(x => x.base === 'kg');
    const precio = (t.match(/\ba (\d+(?:\.\d+)?)(?: euros?)?(?: el| por| cada)? (?:kilo|kg)\b/) || [])[1];
    let importe = (t.match(/\b(\d+(?:\.\d+)?) euros?\b(?! (?:el|por|cada) (?:kilo|kg))/) || [])[1];
    importe = importe ? +importe : (kg && precio ? Math.round(kg.q * +precio * 100) / 100 : null);
    const sinNum = t.replace(/\d+(\.\d+)?/g, ' ');
    const { fs } = /\b(finca|parcela)\b/.test(t) || buscarNombres(sinNum, candidatosFinca()).length ? fincasDichas(sinNum, {}) : { fs: [] };
    R.mov = { tipo, fecha: R.fecha, categoria: cat, concepto: R.texto.trim().replace(/^./, x => x.toUpperCase()).slice(0, 80), kilos: kg ? kg.q : null, precio: precio ? +precio : null, importe, fincas: fs.map(f => f.id), campana: db.campanaActiva, notas: '' };
    if (!importe) R.dudas.push('No has dicho el importe en euros.');
    return R;
  }

  const compacto = s => normVoz(s).replace(/[^a-z0-9]/g, '');
  // «número de registro ES00367», «registro 19549», «ES 00367»
  function fitoPorRegistro(t) {
    if (!FITOS) return null;
    const m = t.match(/\bregistro(?: numero| n)? (?:es ?)?(\d{3,6})\b/) || t.match(/\bes ?(\d{5})\b/);
    if (!m) return null;
    const dd = m[1].replace(/^0+/, '');
    return FITOS.productos.find(p => { const r = String(p.r).toUpperCase().replace(/[^A-Z0-9]/g, ''); return r.replace(/^ES0*/, '') === dd; }) || null;
  }
  // Busca un nombre comparando sin espacios ni guiones («U46 de comple» ~ «U-46 D COMPLET»)
  function mejorPorNombre(nombre, lista, clave, umbral = 0.78) {
    const c = compacto(nombre); if (c.length < 3) return null;
    let best = null, bs = 0;
    for (const x of lista) { const k = compacto(clave(x)); const s = Math.max(parecido(c, k), k.startsWith(c) && c.length >= 4 ? 0.9 : 0, c.startsWith(k) && k.length >= 5 ? 0.88 : 0); if (s > bs) { bs = s; best = x; } }
    return bs >= umbral ? best : null;
  }
  function lineasDichas(t) {
    // «20 litros de accresto a 30 euros el litro y 1000 kilos de urea»
    const lineas = []; const cants = cantidadesDichas(t);
    cants.forEach((c, k) => {
      const fin = k + 1 < cants.length ? cants[k + 1].i : t.length;
      const trozo = t.slice(c.i + c.txt.length, fin);
      const nombre = (trozo.match(/^\s*(?:de |del )?(.+?)(?= a \d| por \d| para |,| y \d| con (?:el )?(?:numero|registro|n )| numero de registro| n de registro| registro | registro$| (?:del |de la )?(?:empresa|proveedor)|$)/) || [])[1] || '';
      const reg = fitoPorRegistro(trozo);
      const precio = (trozo.match(/\ba (\d+(?:\.\d+)?)(?: euros?)?(?: el| por| cada)? ?(?:kilo|kg|litro|l|unidad|tonelada)?/) || [])[1];
      const imp = (trozo.match(/\bpor (\d+(?:\.\d+)?) euros?/) || [])[1];
      lineas.push({ q: c.q, base: c.base, nombre: nombre.replace(/\b(y|para|el|la|con)\s*$/, '').trim(), precio: precio ? +precio : null, importe: imp ? +imp : null, reg });
    });
    return lineas;
  }
  // Forma ordenada: «producto ES-00599 cantidad 80 litros precio 6 euros, producto urea cantidad 1000 kilos»
  function lineasEstructuradas(t) {
    return t.split(/\bproductos?\b/).slice(1).map(seg => {
      const nombre = (seg.match(/^\s*(?:de |del |el |la )?(.+?)(?= cantidad| precio| importe| a \d| por \d|,| \d+(?:\.\d+)? (?:kilos?|kg|litros?|l|toneladas?|gramos?|cc|ml)\b|$)/) || [])[1] || '';
      const c = cantidadesDichas(seg)[0];
      const precio = (seg.match(/\bprecio(?: de| unitario)? (\d+(?:\.\d+)?)/) || seg.match(/\ba (\d+(?:\.\d+)?) euros?/) || [])[1];
      const imp = (seg.match(/\b(?:importe|por|total)(?: de)? (\d+(?:\.\d+)?) euros?/) || [])[1];
      return c ? { q: c.q, base: c.base, nombre: nombre.trim(), precio: precio ? +precio : null, importe: imp ? +imp : null, reg: fitoPorRegistro(seg) } : null;
    }).filter(Boolean);
  }
  function productoParaLinea(l) {
    const deRegistro = r => { const ya = db.productos.find(p => p.regId === r.id); return ya ? { p: ya } : { nuevo: { nombre: r.n, tipo: 'Fitosanitario', unidad: unidadDeFormulado(r.f), regId: r.id, registro: r.r, formulado: r.f } }; };
    if (l.reg) return deRegistro(l.reg);
    const n = normVoz(l.nombre); if (!n) return null;
    const mio = buscarNombres(n, candidatosProducto(null), 0.8)[0] || (() => { const p = mejorPorNombre(n, db.productos, p => p.nombre, 0.8); return p ? { c: { p } } : null; })();
    if (mio) return { p: mio.c.p };
    if (FITOS) {
      const vig = FITOS.productos.filter(x => x.e === 'Vigente');
      const r = mejorPorNombre(n, vig, x => x.n) || mejorPorNombre(n, FITOS.productos, x => x.n, 0.85);
      if (r) return deRegistro(r);
    }
    const cu = db.cultivos.find(c => n.includes(norm(c)));
    const tipo = /\b(gasoleo|gasoil|gas oil|combustible)\b/.test(n) ? 'Combustible' : /\b(urea|nitrato|abono|npk|nac|fosfato|potasa|sulfato|complejo)\b/.test(n) ? 'Abono' : (/\bsemilla|para sembrar\b/.test(n) || cu) ? 'Semilla' : 'Otro';
    const nombre = tipo === 'Semilla' && cu && !/sembrar|semilla/.test(n) ? `${cu} para sembrar` : l.nombre.replace(/^./, x => x.toUpperCase());
    return { nuevo: { nombre, tipo, unidad: l.base === 'L' || l.base === 'cc' ? 'L' : 'kg' } };
  }
  function interpretarDocumento(t, R, tipo) {
    R.intencion = tipo;
    const tits = candidatosTitular();
    const tit = buscarNombres(t, tits, 0.85);
    const palabraDoc = tipo === 'albaran' ? 'albaran' : 'factura';
    // proveedor: «a la empresa Terrastar», «proveedor X», «albarán de la cooperativa», o uno que ya hayas usado
    const finProv = '(?= para | numero | n \\d|,| \\d| de \\d| con | por | el dia| y | producto| cantidad| precio| importe| a nombre| nombre| titular| fecha|$)';
    let proveedor = (t.match(new RegExp(`\\b(?:empresa|proveedor|casa comercial|distribuidor|distribuidora|comercial)(?: de| del)? (?:la |el |los )?([a-z0-9 ]+?)${finProv}`)) || [])[1]
      || (t.match(new RegExp(`\\b${palabraDoc}(?: (?:numero|n) \\w+)? (?:de|del) (?:la |el |los )?([a-z ]+?)${finProv}`)) || [])[1]
      || (t.match(new RegExp(`\\b(?:compra|compro|compre|comprado|compramos)\\b.*? a (?:la |el |los )?(?!\\d)([a-z][a-z0-9 ]+?)${finProv}`)) || [])[1] || '';
    proveedor = proveedor.replace(/^(empresa|proveedor) /, '').trim();
    if (proveedor && (tits.some(x => parecido(x.n, proveedor) > 0.85) || /^(numero|albaran|factura)\b/.test(proveedor))) proveedor = '';
    const conocidos = [...new Set(db.documentos.map(x => x.proveedor).filter(Boolean))];
    const prevProv = proveedor ? mejorPorNombre(proveedor, conocidos, x => x, 0.75) : (buscarNombres(t, conocidos.map(x => ({ n: normVoz(x), x })), 0.82)[0] || {}).c?.x;
    if (prevProv) proveedor = prevProv;
    // número: «número de albarán 324», «albarán número 324», «albarán 324» (no el nº de registro)
    const numero = (t.match(new RegExp(`\\b(?:numero|n|num)(?: de ${palabraDoc})? (?!de registro)(\\d[\\w\\/-]*)`)) || t.match(new RegExp(`\\b${palabraDoc}(?: numero| n| num)? (\\d[\\w\\/-]*)`)) || [])[1] || '';
    const pre = { titular: tit.length ? tit[0].c.t : '', fecha: R.fecha, proveedor: prevProv ? proveedor : proveedor.replace(/\b\w/g, x => x.toUpperCase()).trim(), numero: numero.toUpperCase(), lineas: [] };
    R.lineasNuevas = [];
    const desde = t.search(/\d+(?:\.\d+)? (kilo|kg|litro|l |l$|tonelada|gramo|centimetro|cc|ml|mililitro)/);
    const ls = /\bproductos?\b/.test(t) ? lineasEstructuradas(t) : desde >= 0 ? lineasDichas(t.slice(desde)) : [];
    ls.forEach(l => {
      const r = productoParaLinea(l); if (!r) return;
      R.lineasNuevas.push({ l, r });
    });
    if (tipo === 'factura') {
      const base = (t.match(/\bbase(?: imponible)?(?: de)? (\d+(?:\.\d+)?)/) || [])[1];
      const iva = (t.match(/\biva(?: del| de)? (\d+(?:\.\d+)?)/) || [])[1];
      const total = (t.match(/\btotal(?: de)? (\d+(?:\.\d+)?)/) || [])[1];
      Object.assign(pre, { base: base ? +base : null, iva: iva ? +iva : 10, total: total ? +total : (base ? Math.round(+base * (1 + (iva ? +iva : 10) / 100) * 100) / 100 : null) });
    }
    R.doc = pre;
    if (!pre.titular) R.dudas.push('¿En qué carpeta lo guardo? Di «para Iván», «de Tomás»…');
    if (tipo === 'albaran' && !ls.length) R.dudas.push('No he entendido los productos: di «20 litros de Accresto», «1.000 kilos de urea»…');
    return R;
  }
  function interpretarFinca(t, R) {
    R.intencion = 'finca';
    const nombre = (t.match(/\b(?:finca|parcela)(?: nueva)? (?:que se llama |llamada |de nombre )?([a-z][a-z ]{1,40}?)(?= poligono| parcela| recinto| en | de \d| con | titular| para |,|\d|$)/) || [])[1] || '';
    const po = (t.match(/\bpoligono (\d+)/) || [])[1], pa = (t.match(/\bparcela (\d+)/) || [])[1], rec = (t.match(/\brecintos? (\d+(?: y \d+)*)/) || [])[1];
    const ha = (t.match(/\b(\d+(?:\.\d+)?) (?:hectareas?|ha)\b/) || [])[1];
    const pueblos = [...new Set(db.fincas.map(f => f.municipio).filter(Boolean))];
    const pu = buscarNombres(t, pueblos.map(x => ({ n: normVoz(x), x })), 0.85)[0];
    const ref = pu ? db.fincas.find(f => f.municipio === pu.c.x && f.codMunicipio) : null;
    const tit = buscarNombres(t, candidatosTitular(), 0.85)[0];
    const grupo = nombre ? [...listaGrupos().keys()].find(g => norm(g) === norm(nombre)) : '';
    R.finca = { nombre: nombre ? nombre.trim().replace(/\b\w/g, x => x.toUpperCase()) + (po && pa ? ` ${po}-${pa}` : '') : (po && pa && pu ? `${pu.c.x} ${po}-${pa}` : ''), grupo: grupo || '', titular: tit ? tit.c.t : '', municipio: pu ? pu.c.x : '', provincia: ref ? ref.provincia : '', codMunicipio: ref ? ref.codMunicipio : '', agregado: 0, zona: 0, poligono: po || '', parcela: pa || '', recinto: rec ? rec.replace(/ y /g, ', ') : '', superficie: ha ? +ha : null };
    if (!po || !pa) R.dudas.push('Para situarla en el mapa di el polígono y la parcela.');
    return R;
  }

  // ---------- Tarjeta de confirmación ----------
  function tarjetaVoz(R) {
    const fila = (k, v) => v ? `<dt>${k}</dt><dd>${v}</dd>` : '';
    const listaParc = (ids, quit) => { const fs = ids.map(finca).filter(Boolean); if (!fs.length) return '<span class="neg">Ninguna</span>';
      const g = new Set(fs.map(grupoDe)); const nombre = g.size === 1 && [...g][0] && fincasDeGrupo([...g][0]).length === fs.length ? [...g][0] : '';
      return `${nombre ? `<b>${esc(nombre)}</b> · ` : ''}${fs.length} parcela${fs.length === 1 ? '' : 's'} · ${fmtN(sum(fs, f => f.superficie))} ha<div class="voz-parc">${fs.map(f => `<span>${esc(f.nombre)}${f.titular ? ` <i>${esc(f.titular)}</i>` : ''}</span>`).join('')}${(quit || []).map(f => `<span class="tachada">${esc(f.nombre)}</span>`).join('')}</div>`; };
    let h = ''; let puedeGuardar = false; let abrir = '';
    if (R.intencion === 'labor') {
      const v = R.labor; const p = producto(v.prodAlm);
      h += `<dl class="kv">${fila('Labor', esc(v.tipo))}${fila('Fecha', fmtFecha(v.fecha))}${fila('Fincas', listaParc(v.fincas, R.quitadas))}${v.cultivoSiembra ? fila('Cultivo', chip(v.cultivoSiembra) + (v.variedadSiembra ? ' ' + esc(v.variedadSiembra) : '') + ` <span class="small muted">en ${esc(v.campana)}</span>`) : ''}${fila('Producto', p ? esc(p.nombre) : esc(v.producto))}${fila('Dosis', v.dosis != null ? `${fmtN(v.dosis, 4)} ${esc(v.unidad)}` : '')}${fila('Contra', esc(v.motivo))}${fila('Maquinaria', (equipo(v.equipoId) ? `<b>${esc(equipo(v.equipoId).nombre)}</b>: ` : '') + esc([v.tractorId, v.aperoId].map(nombreMaq).filter(Boolean).join(' + ')))}${fila('Aplicador', esc(v.aplicador))}${fila('Coste', v.coste ? eur(v.coste, 2) : '')}</dl>`;
      if (p) h += `<div class="fito-info">${infoLaborHTML(Object.assign({}, v))}</div>`;
      puedeGuardar = v.fincas.length > 0 && !R.bloqueo && v.tipo !== 'Otro';
      abrir = 'labor';
    } else if (R.intencion === 'cosecha') {
      const v = R.cosecha; const fs = v.fincas.map(finca).filter(Boolean); const ha = sum(fs, f => f.superficie);
      h += `<dl class="kv">${fila('Cosecha', v.kg ? fmtKg(v.kg) : '<span class="neg">sin kilos</span>')}${fila('Fecha', fmtFecha(v.fecha))}${fila('Fincas', listaParc(v.fincas, R.quitadas))}${fila('Rendimiento', v.kg && ha ? `${fmtN(v.kg / ha, 0)} kg/ha` : '')}${fila('Humedad', v.humedad != null ? fmtN(v.humedad, 1) + ' %' : '')}${fila('Cosechadora', esc(nombreMaq(v.maquinaId)))}</dl>
        <p class="small muted" style="margin:0">Se apunta la cosecha y la labor «Cosecha» en el cuaderno.</p>`;
      puedeGuardar = fs.length > 0 && v.kg > 0; abrir = 'cosecha';
    } else if (R.intencion === 'gasto' || R.intencion === 'ingreso') {
      const v = R.mov;
      h += `<dl class="kv">${fila(v.tipo === 'gasto' ? 'Gasto' : 'Ingreso', v.importe ? `<b class="${v.tipo === 'gasto' ? 'neg' : 'pos'}">${eur(v.importe, 2)}</b>` : '<span class="neg">sin importe</span>')}${fila('Categoría', esc(v.categoria))}${fila('Fecha', fmtFecha(v.fecha))}${fila('Kilos', v.kilos ? fmtKg(v.kilos) : '')}${fila('Precio', v.precio ? `${fmtN(v.precio, 4)} €/kg` : '')}${fila('Fincas', v.fincas.length ? listaParc(v.fincas) : 'General (sin finca)')}${fila('Concepto', esc(v.concepto))}</dl>`;
      puedeGuardar = v.importe > 0; abrir = 'mov';
    } else if (R.intencion === 'semilla') {
      const v = R.semilla;
      h += `<dl class="kv">${fila('Semilla propia', v.cultivo ? `${esc(v.cultivo)} para sembrar` : '?')}${fila('Kilos', v.kg ? fmtKg(v.kg) : '?')}${fila('Titular', esc(v.titular) || '?')}${fila('Fecha', fmtFecha(v.fecha))}</dl><p class="small muted" style="margin:0">Se suma al almacén del titular como semilla de vuestra cosecha.</p>`;
      puedeGuardar = !!(v.cultivo && v.kg && v.titular);
    } else if (R.intencion === 'albaran' || R.intencion === 'factura') {
      const v = R.doc;
      h += `<dl class="kv">${fila(R.intencion === 'albaran' ? 'Albarán' : 'Factura', esc(v.proveedor || 'sin proveedor') + (v.numero ? ' · nº ' + esc(v.numero) : ''))}${fila('Carpeta', esc(v.titular) || '<span class="neg">?</span>')}${fila('Fecha', fmtFecha(v.fecha))}${R.intencion === 'factura' ? fila('Base', v.base != null ? eur(v.base, 2) : '') + fila('IVA', v.iva != null ? fmtN(v.iva) + ' %' : '') + fila('Total', v.total != null ? eur(v.total, 2) : '') : ''}</dl>
        ${R.lineasNuevas.length ? `<div class="list">${R.lineasNuevas.map(({ l, r }) => { const pn = r.p || r.nuevo; return `<div class="row"><div class="main"><div class="t">${esc(pn.nombre)}</div><div class="s">${esc(pn.tipo)}${r.nuevo ? ' · producto nuevo' : ''}${pn.registro ? ' · Nº reg. ' + esc(pn.registro) : ''}</div></div><div class="end">${fmtN(l.base === 'cc' ? l.q / 1000 : l.q, 3)} ${esc(pn.unidad)}${l.precio ? `<div class="small muted">${fmtN(l.precio, 4)} €</div>` : ''}</div></div>`; }).join('')}</div>` : ''}
        <p class="small muted" style="margin:0">Al abrirlo podrás añadir la foto o el PDF y revisar antes de guardar.</p>`;
      abrir = R.intencion;
    } else if (R.intencion === 'finca') {
      const v = R.finca;
      h += `<dl class="kv">${fila('Finca nueva', esc(v.nombre) || '?')}${fila('Pueblo', esc(v.municipio))}${fila('SIGPAC', v.poligono ? `Pol. ${esc(v.poligono)} · Parc. ${esc(v.parcela)}${v.recinto ? ' · Rec. ' + esc(v.recinto) : ''}` : '')}${fila('Superficie', v.superficie ? fmtN(v.superficie) + ' ha' : '')}${fila('Titular', esc(v.titular))}${fila('Finca agrupada', esc(v.grupo))}</dl><p class="small muted" style="margin:0">Al abrirla podrás completar lo que falte.</p>`;
      abrir = 'finca';
    }
    if (R.bloqueo) h += `<div class="aviso-norma"><b>⚠ ${esc(R.bloqueo)}</b></div>`;
    const notas = [...R.avisos, ...R.dudas];
    if (notas.length) h += `<ul class="small notas voz-dudas">${notas.map(x => `<li>${esc(x)}</li>`).join('')}</ul>`;
    const botones = `${puedeGuardar ? '<button type="button" class="btn primary" data-v="guardar">Guardar</button>' : ''}${abrir ? `<button type="button" class="btn${puedeGuardar ? '' : ' primary'}" data-v="abrir">${puedeGuardar ? 'Corregir' : 'Abrir y revisar'}</button>` : ''}`;
    return { html: h ? `<div class="card pad voz-tarjeta"><div class="tag">Esto es lo que he entendido</div>${h}<div class="btns" style="justify-content:flex-start">${botones}</div></div>` : (notas.length ? `<div class="card pad voz-tarjeta"><ul class="small notas voz-dudas">${notas.map(x => `<li>${esc(x)}</li>`).join('')}</ul></div>` : ''), puedeGuardar };
  }

  // Ejecuta lo entendido
  async function guardarVoz(R) {
    if (R.intencion === 'labor') {
      const v = Object.assign({}, R.labor);
      if (v.prodAlm) {
        const r = revisarLabor(v);
        if (r.graves.length) { const c = await confirmar('⚠ Revisa antes de guardar', `<ul class="avisos-lista">${r.graves.map(g => `<li>${esc(g)}</li>`).join('')}</ul>Si lo guardas, la labor quedará marcada para revisar en el cuaderno.`, 'Guardar igualmente'); if (!c.ok) return false; }
      }
      registrarLabor(v); return `${v.tipo} apuntada en ${v.fincas.length} parcela${v.fincas.length === 1 ? '' : 's'}`;
    }
    if (R.intencion === 'cosecha') {
      const v = R.cosecha; registrarCosecha(v);
      if (v.maquinaId) { db.ultMaq = db.ultMaq || {}; db.ultMaq.Cosecha = Object.assign({}, db.ultMaq.Cosecha, { tractorId: v.maquinaId }); }
      registrarLabor({ fecha: v.fecha, tipo: 'Cosecha', fincas: v.fincas, campana: v.campana, notas: '', producto: '', dosis: null, unidad: '', coste: null, tractorId: v.maquinaId || '' });
      return `Cosecha de ${fmtKg(v.kg)} apuntada`;
    }
    if (R.intencion === 'gasto' || R.intencion === 'ingreso') { registrarMov(Object.assign({}, R.mov)); return R.intencion === 'gasto' ? 'Gasto apuntado' : 'Ingreso apuntado'; }
    if (R.intencion === 'semilla') {
      const v = R.semilla; const nombre = `${v.cultivo} para sembrar`;
      let p = db.productos.find(x => x.tipo === 'Semilla' && norm(x.nombre) === norm(nombre));
      if (!p) { p = { id: uid(), nombre, tipo: 'Semilla', unidad: 'kg' }; db.productos.push(p); }
      db.ajustes.push({ id: uid(), productoId: p.id, titular: v.titular, cantidad: v.kg, fecha: v.fecha, motivo: 'Semilla de la cosecha propia' });
      return `${fmtKg(v.kg)} de ${nombre.toLowerCase()} para ${v.titular}`;
    }
    return false;
  }
  function abrirVoz(R) {
    if (R.intencion === 'labor') formLabor(null, { valores: Object.assign({}, R.labor) });
    else if (R.intencion === 'cosecha') formCosecha(null, { valores: Object.assign({}, R.cosecha) });
    else if (R.intencion === 'gasto' || R.intencion === 'ingreso') formMov(null, { tipo: R.mov.tipo, valores: Object.assign({}, R.mov) });
    else if (R.intencion === 'albaran' || R.intencion === 'factura') {
      const pre = Object.assign({}, R.doc);
      pre.lineas = R.lineasNuevas.map(({ l, r }) => {
        let p = r.p; if (!p) { p = db.productos.find(x => norm(x.nombre) === norm(r.nuevo.nombre)) || Object.assign({ id: uid() }, r.nuevo); if (!db.productos.includes(p)) db.productos.push(p); }
        const q = l.base === 'cc' ? l.q / 1000 : l.base === 'g' && p.unidad === 'kg' ? l.q / 1000 : l.q;
        return { productoId: p.id, cantidad: q, precio: l.precio, importe: l.importe != null ? l.importe : (l.precio ? Math.round(q * l.precio * 100) / 100 : null) };
      });
      guardar(false);
      formDocumento(R.intencion, null, pre);
    } else if (R.intencion === 'finca') formFinca(null, Object.assign({}, R.finca));
  }

  const EJEMPLOS_VOZ = ['Acabo de sembrar la Sarrana de trigo a 180 kilos', 'Ayer aré Consejera con el John Deere y el chísel', 'He sulfatado la Sarrana con Accresto a 200 centímetros contra avena loca', 'Ayer aré Consejera', 'He cosechado Consejera, 14.000 kilos', 'Albarán de la cooperativa para Iván: 20 litros de Accresto a 30 euros el litro', 'Guardar 5.000 kilos de trigo de la cosecha como semilla para Iván', 'Gasto de gasóleo 300 euros', 'Nueva finca Los Prados en Lerma polígono 663 parcela 180, 3,5 hectáreas, titular Iván'];
  function hojaVoz() {
    const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
    cargarFitos().then(d => { FITOS = d; }).catch(() => {});
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera('Asistente')}<div class="body">
      <div class="voz-cab"><button type="button" class="voz-mic" id="vozMic" aria-label="Hablar">${ico.mic}</button><div class="voz-estado" id="vozEstado">Pulsa el micrófono y di lo que has hecho.</div></div>
      <div class="field"><textarea id="vozTxt" rows="3" placeholder="O escríbelo aquí: «He sembrado la Sarrana de trigo a 180 kilos»" aria-label="Lo que has dicho"></textarea></div>
      <button type="button" class="btn block" id="vozEntender">Entender</button>
      <div id="vozRes"></div>
      <details class="voz-ej"><summary>Ejemplos de lo que puedes decir</summary><div class="chips">${EJEMPLOS_VOZ.map(e => `<button type="button" class="chip none" data-ej="${esc(e)}">${esc(e)}</button>`).join('')}</div>
        <p class="small muted">Puedes decir «excepto la parte de Tomás» o «menos la parcela 180», y la fecha («ayer», «el lunes», «el 5 de noviembre»). Sin cobertura, usa el micrófono del teclado del móvil.</p></details>
      </div></div>`, (el, d) => {
      const ta = el.querySelector('#vozTxt'), est = el.querySelector('#vozEstado'), res = el.querySelector('#vozRes'), mic = el.querySelector('#vozMic');
      let R = null, rec = null, escuchando = false;
      const entender = () => {
        if (!ta.value.trim()) { res.innerHTML = ''; return; }
        R = interpretar(ta.value);
        res.innerHTML = tarjetaVoz(R).html;
      };
      el.querySelector('#vozEntender').addEventListener('click', entender);
      ta.addEventListener('keydown', e => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); entender(); } });
      el.addEventListener('click', async e => {
        const ej = e.target.closest('[data-ej]'); if (ej) { ta.value = ej.dataset.ej; entender(); return; }
        const b = e.target.closest('[data-v]'); if (!b || !R) return;
        if (b.dataset.v === 'guardar') {
          const r = await guardarVoz(R); if (!r) return;
          guardar(); volverA(d - 1); render(); toast(r);
        } else if (b.dataset.v === 'abrir') { const RR = R; volverA(d - 1); abrirVoz(RR); }
      });
      mic.addEventListener('click', () => {
        if (escuchando && rec) { rec.stop(); return; }
        if (!SR) { est.textContent = 'Este navegador no deja dictar aquí. Pulsa el micrófono del teclado y habla.'; ta.focus(); return; }
        rec = new SR(); rec.lang = 'es-ES'; rec.interimResults = true; rec.maxAlternatives = 1; rec.continuous = false;
        rec.onstart = () => { escuchando = true; mic.classList.add('on'); est.textContent = 'Te escucho…'; ta.value = ''; res.innerHTML = ''; };
        rec.onresult = e => { let txt = ''; for (const r of e.results) txt += r[0].transcript; ta.value = txt; };
        rec.onerror = e => { est.textContent = e.error === 'not-allowed' || e.error === 'service-not-allowed' ? 'Da permiso al micrófono para poder dictar.' : e.error === 'network' ? 'Para dictar aquí hace falta cobertura. Sin cobertura, usa el micrófono del teclado.' : e.error === 'no-speech' ? 'No te he oído. Pulsa y habla otra vez.' : 'No te he entendido. Prueba otra vez.'; };
        rec.onend = () => { escuchando = false; mic.classList.remove('on'); if (ta.value.trim()) { est.textContent = 'Revisa lo que he entendido:'; entender(); } else if (est.textContent === 'Te escucho…') est.textContent = 'Pulsa el micrófono y di lo que has hecho.'; };
        try { rec.start(); } catch (err) { est.textContent = 'No se ha podido usar el micrófono.'; }
      });
    });
  }

  // ---------- Explotación: datos de cada agricultor (titular) y su maquinaria ----------
  const NIVELES_ROPO = ['Básico', 'Cualificado', 'Fumigador', 'Piloto aplicador'];
  const TIPOS_MAQ = ['Tractor', 'Cosechadora', 'Sembradora', 'Pulverizador', 'Abonadora', 'Cultivador', 'Grada', 'Arado', 'Chísel', 'Rodillo', 'Remolque', 'Empacadora', 'Vehículo', 'Otro'];
  const CON_ITV = ['Tractor', 'Cosechadora', 'Remolque', 'Vehículo'];
  const CON_ITEAF = ['Pulverizador', 'Abonadora'];
  const CON_ANCHURA = ['Cosechadora', 'Sembradora', 'Pulverizador', 'Abonadora', 'Cultivador', 'Grada', 'Arado', 'Chísel', 'Rodillo'];
  const fichaTitular = n => db.titulares.find(t => norm(t.nombre) === norm(n)) || null;
  // Comprueba la letra de control de un DNI, NIE o CIF. null = vacío
  function nifValido(s) {
    const v = String(s || '').toUpperCase().replace(/[\s.\-]/g, ''); if (!v) return null;
    const L = 'TRWAGMYFPDXBNJZSQVHLCKE';
    let m = v.match(/^(\d{8})([A-Z])$/); if (m) return L[+m[1] % 23] === m[2];
    m = v.match(/^([XYZ])(\d{7})([A-Z])$/); if (m) return L[+('XYZ'.indexOf(m[1]) + m[2]) % 23] === m[3];
    m = v.match(/^([ABCDEFGHJNPQRSUVW])(\d{7})([0-9A-J])$/);
    if (m) { let a = 0, b = 0; for (let i = 0; i < 7; i++) { const n = +m[2][i]; if (i % 2 === 0) { const x = n * 2; a += Math.floor(x / 10) + x % 10; } else b += n; } const c = (10 - (a + b) % 10) % 10; return m[3] === String(c) || m[3] === 'JABCDEFGHI'[c]; }
    return false;
  }
  // Fecha de caducidad o vencimiento: vencida, pronto (menos de N días) o al día
  function estadoFecha(iso, dias = 60) {
    if (!iso) return null;
    const hoy = hoyISO(); if (iso < hoy) return 'vencido';
    const lim = new Date(); lim.setDate(lim.getDate() + dias);
    return iso <= new Date(lim - lim.getTimezoneOffset() * 60000).toISOString().slice(0, 10) ? 'pronto' : 'ok';
  }
  const vencidoDe = txt => /^(ITV|Inspecci)/.test(txt) ? 'vencida' : 'vencido';
  const pillFecha = (txt, iso, dias) => { const e = estadoFecha(iso, dias); if (!e) return ''; return `<span class="pill ${e === 'vencido' ? 'neg' : e === 'pronto' ? 'warn' : 'pos'}">${txt} ${e === 'vencido' ? vencidoDe(txt) : e === 'pronto' ? 'vence' : 'hasta'} ${fmtFecha(iso)}</span>`; };
  function avisosMaquina(m) {
    const out = [];
    if (CON_ITV.includes(m.tipo)) out.push(['ITV', m.itv, 30]);
    if (CON_ITEAF.includes(m.tipo)) out.push(['Inspección ITEAF', m.iteaf, 60]);
    if (m.seguroVence) out.push(['Seguro', m.seguroVence, 30]);
    return out;
  }

  // ---------- Maquinaria en las labores ----------
  // Cualquier máquina sirve en cualquier finca: la labor guarda la máquina (y de quién es) y la finca (y de quién es).
  const MOTRICES = ['Tractor', 'Cosechadora', 'Vehículo'];
  const maquina = id => (id && db.maquinas.find(m => m.id === id)) || null;
  const nombreMaq = id => { const m = maquina(id); return m ? m.nombre + (m.titular ? ` (${m.titular})` : '') : ''; };
  function optsMaq(motriz, actual, tipos) {
    const ms = db.maquinas.filter(m => (tipos ? tipos.includes(m.tipo) : MOTRICES.includes(m.tipo) === motriz) && (!m.baja || m.id === actual))
      .sort((a, b) => TIPOS_MAQ.indexOf(a.tipo) - TIPOS_MAQ.indexOf(b.tipo) || a.nombre.localeCompare(b.nombre, 'es', { numeric: true }));
    return [['', motriz ? '—' : 'Ninguno'], ...ms.map(m => [m.id, `${m.nombre}${tipos ? '' : ' · ' + m.tipo.toLowerCase()}${m.titular ? ' · ' + m.titular : ''}`])];
  }
  const APERO_DE = { Tratamiento: ['Pulverizador'], Abonado: ['Abonadora'], Siembra: ['Sembradora'], Rodillo: ['Rodillo'], Empacado: ['Empacadora'] };
  // La última que usaste en ese tipo de labor; si no, la única que haya de ese tipo
  function maqPorDefecto(tipo) {
    const u = (db.ultMaq || {})[tipo] || {};
    const eqOk = id => { const e = equipo(id); return e && [e.tractorId, e.aperoId].filter(Boolean).every(x => { const m = maquina(x); return m && !m.baja; }) ? e : null; };
    let e = u.equipoId ? eqOk(u.equipoId) : null;
    if (!e && !u.tractorId && !u.aperoId) { const es = db.equipos.filter(x => x.labor === tipo && eqOk(x.id)); if (es.length === 1) e = es[0]; }
    if (e) return { equipoId: e.id, tractorId: e.tractorId || '', aperoId: e.aperoId || '' };
    const ok = id => { const m = maquina(id); return m && !m.baja ? id : ''; };
    const unica = tipos => { const ms = db.maquinas.filter(m => !m.baja && tipos.includes(m.tipo)); return ms.length === 1 ? ms[0].id : ''; };
    let tractorId = ok(u.tractorId), aperoId = ok(u.aperoId);
    if (!tractorId && !u.tractorId) tractorId = unica(tipo === 'Cosecha' ? ['Cosechadora'] : ['Tractor']);
    if (!aperoId && !u.aperoId && APERO_DE[tipo]) aperoId = unica(APERO_DE[tipo]);
    return { equipoId: '', tractorId, aperoId };
  }
  const conCarne = () => titularesTodos().filter(t => { const f = fichaTitular(t); return f && f.ropo; });
  function aplicadorPorDefecto() {
    if (db.ultAplicador && titularesTodos().includes(db.ultAplicador)) return db.ultAplicador;
    const c = conCarne(); return c.length === 1 ? c[0] : '';
  }
  function optsAplicador(actual) {
    const ts = titularesTodos(); if (actual && !ts.includes(actual)) ts.push(actual);
    return [['', '—'], ...ts.map(t => { const f = fichaTitular(t); return [t, `${t}${f && f.ropo ? ' · ROPO ' + f.ropo : ' · sin carné apuntado'}${f && estadoFecha(f.ropoCaduca, 0) === 'vencido' ? ' · CADUCADO' : ''}`]; })];
  }
  // Lo que ha trabajado una máquina: hectáreas por labor y por dueño de la tierra
  function trabajoMaquina(id, campana) {
    const vistos = new Map();
    const add = (fid, fecha, tipo, ha) => { const k = fid + '|' + fecha + '|' + tipo; if (!vistos.has(k)) vistos.set(k, { fid, tipo, ha }); };
    db.labores.forEach(l => { if ((l.tractorId === id || l.aperoId === id) && (!campana || l.campana === campana)) add(l.fincaId, l.fecha, l.tipo, l.superficie > 0 ? +l.superficie : +((finca(l.fincaId) || {}).superficie) || 0); });
    db.cosechas.forEach(c => { if (c.maquinaId === id && (!campana || c.campana === campana)) add(c.fincaId, c.fecha, 'Cosecha', +((finca(c.fincaId) || {}).superficie) || 0); });
    const xs = [...vistos.values()];
    const agrupa = f => { const m = new Map(); xs.forEach(x => { const k = f(x); m.set(k, (m.get(k) || 0) + x.ha); }); return [...m.entries()].sort((a, b) => b[1] - a[1]); };
    return { n: xs.length, ha: sum(xs, x => x.ha), porTipo: agrupa(x => x.tipo), porTit: agrupa(x => titularDe(x.fid) || 'Sin titular') };
  }
  // Máquinas dichas por voz: nombre, marca, modelo o «la sulfatadora (de Tomás)» si no hay duda
  const SINON_MAQ = { Pulverizador: ['sulfatadora', 'pulverizador', 'atomizador', 'maquina de sulfatar'], Abonadora: ['abonadora'], Sembradora: ['sembradora'], Cosechadora: ['cosechadora'], Tractor: ['tractor'], Remolque: ['remolque'], Empacadora: ['empacadora', 'empacadora'], Cultivador: ['cultivador'], Grada: ['grada'], Arado: ['arado', 'vertedera'], 'Chísel': ['chisel'], Rodillo: ['rodillo'] };
  function candidatosMaquina() {
    const eqs = []; db.equipos.forEach(e => { const n = normVoz(e.nombre); if (n.length >= 3) { eqs.push({ n, e }); const sin = n.replace(/^equipo (de |del )?/, ''); if (sin !== n && sin.length >= 4) eqs.push({ n: 'equipo de ' + sin, e }); } });
    const m = new Map(); const add = (n, x) => { n = normVoz(n); if (!n || n.length < 3) return; if (!m.has(n)) m.set(n, new Set()); m.get(n).add(x); };
    db.maquinas.filter(x => !x.baja).forEach(x => {
      add(x.nombre, x); add(x.nombre.replace(/^(el|la|los|las) /i, ''), x);
      if (x.marca) { add(x.marca, x); if (x.modelo) add(x.marca + ' ' + x.modelo, x); }
      if (x.modelo && /[a-z]/i.test(x.modelo) && /\d/.test(x.modelo)) add(x.modelo, x);
      (SINON_MAQ[x.tipo] || []).forEach(s => { add(s, x); if (x.titular) add(`${s} de ${x.titular}`, x); });
    });
    return [...m.entries()].filter(([, set]) => set.size === 1).map(([n, set]) => ({ n, m: [...set][0] })).concat(eqs);
  }
  function maquinasDichas(t) {
    const cs = candidatosMaquina(); if (!cs.length) return { ms: [], t };
    const hits = buscarNombres(t, cs, 0.86); const ms = []; let eq = null;
    hits.forEach(h => { if (h.c.e) { if (!eq) eq = h.c.e; return; } if (!ms.includes(h.c.m) && ms.length < 2) ms.push(h.c.m); });
    // dos máquinas del mismo tipo (tractor y tractor): solo la primera
    const out = ms.filter((x, i) => !ms.slice(0, i).some(y => MOTRICES.includes(y.tipo) === MOTRICES.includes(x.tipo)));
    let w = quitarTramos(t, hits).split(' ');
    // «el John Deere de Iván»: el «de Iván» es de la máquina, no de las fincas
    hits.forEach(h => { const e = h.i + h.k; if (w[e] === 'de' && w[e + 1] && candidatosTitular().some(c => parecido(c.n, w[e + 1]) > 0.85)) { w[e] = '_'; w[e + 1] = '_'; } });
    return { ms: out, equipo: eq, t: w.join(' ') };
  }


  // ---------- Equipos de trabajo y coste de la maquinaria ----------
  // Coste por hora de una máquina = (precio de compra − valor al venderla) / horas de amortización.
  // Coste por hectárea de un equipo = (coste/hora del tractor + coste/hora del apero) / rendimiento (ha/h).
  const equipo = id => (id && db.equipos.find(e => e.id === id)) || null;
  const costeHoraMaq = m => m && +m.precioCompra > 0 && +m.horasVida > 0 ? Math.max(0, +m.precioCompra - (+m.valorResidual || 0)) / +m.horasVida : null;
  function costeEquipo(e) {
    const r = { porHora: 0, porHa: null, partes: [], faltan: [] };
    [e.tractorId, e.aperoId].filter(Boolean).forEach(id => {
      const m = maquina(id); if (!m) { r.faltan.push('una máquina que ya no está'); return; }
      const c = costeHoraMaq(m);
      if (c == null) r.faltan.push(`${m.nombre}: ${!(+m.precioCompra > 0) ? 'precio de compra' : 'horas de amortización'}`); else { r.porHora += c; r.partes.push([m.nombre, c]); }
    });
    if (+e.rendimiento > 0 && r.partes.length) r.porHa = r.porHora / +e.rendimiento;
    return r;
  }
  // Labores hechas con cada equipo en la campaña: hectáreas, horas y coste
  function usoEquipos(campana) {
    const out = new Map(); db.equipos.forEach(e => out.set(e.id, { e, ha: 0, n: 0 }));
    db.labores.forEach(l => {
      const o = out.get(l.equipoId); if (!o || (campana && l.campana !== campana)) return;
      o.ha += l.superficie > 0 ? +l.superficie : +((finca(l.fincaId) || {}).superficie) || 0; o.n++;
    });
    db.labores.forEach(l => { const o = out.get(l.equipoId); if (!o || (campana && l.campana !== campana)) return; o.litros = (o.litros || 0) + litrosLabor(l); o.gas = (o.gas || 0) + (costeGasLabor(l) || 0); });
    const pr = precioGasoleo(100);
    out.forEach(o => { const c = costeEquipo(o.e); o.horas = +o.e.rendimiento > 0 ? o.ha / +o.e.rendimiento : null; o.gasHa = +o.e.gasoleo > 0 && pr != null ? +o.e.gasoleo * pr : 0; o.porHa = c.porHa != null ? c.porHa + o.gasHa : null; o.coste = c.porHa != null ? o.ha * c.porHa + (o.gas || 0) : null; o.faltan = c.faltan; });
    return [...out.values()];
  }
  function bloqueEquipos() {
    const us = usoEquipos(db.campanaActiva);
    let h = `<div class="sec-row"><h2 class="sec">Equipos de trabajo</h2><button class="btn sm" data-act="nuevoEquipo">${ico.plus}Nuevo equipo</button></div>`;
    if (!us.length) return h + '<p class="note" style="margin-top:4px">Junta un tractor con su apero (por ejemplo «Equipo de siembra»: tractor + sembradora) y pon cuántas hectáreas hace por hora. Con eso la app calcula el coste por hectárea de cada labor.</p>';
    h += '<div class="list">' + us.map(o => {
      const e = o.e; const maqs = [e.tractorId, e.aperoId].map(id => (maquina(id) || {}).nombre).filter(Boolean).join(' + ');
      const sub = [e.labor, maqs, +e.rendimiento > 0 ? `${fmtN(e.rendimiento)} ha/h` : 'sin rendimiento', +e.gasoleo > 0 ? `${fmtN(e.gasoleo)} L/ha` : ''].filter(Boolean).join(' · ');
      const uso = o.n ? [`${fmtN(o.ha)} ha`, o.horas != null ? fmtN(o.horas, 1) + ' h' : '', o.litros ? fmtN(o.litros, 0) + ' L' : '', o.coste != null ? eur(o.coste) : ''].filter(Boolean).join(' · ') : 'Sin usar esta campaña';
      return `<button class="row" data-act="editEquipo" data-id="${e.id}"><div class="main"><div class="t">${esc(e.nombre)}</div><div class="s">${esc(sub)}</div><div class="s">${esc(uso)}</div>${o.faltan.length || !(+e.rendimiento > 0) ? '<span class="pill warn">Faltan datos para el coste</span>' : ''}</div><div class="end">${o.porHa != null ? `<b>${eur(o.porHa, 2)}</b><div class="small muted">por ha</div>` : ''}</div></button>`;
    }).join('') + '</div>';
    const conCoste = us.filter(o => o.coste != null && o.ha > 0);
    if (conCoste.length) {
      const tot = sum(conCoste, o => o.coste); const ha = sum(db.fincas, f => +f.superficie || 0);
      h += `<div class="card pad small" style="margin-top:8px;display:flex;flex-direction:column;gap:2px"><b>Coste de maquinaria en ${esc(db.campanaActiva)} (amortización y gasóleo)</b><span><b>${eur(tot)}</b> en ${fmtN(sum(conCoste, o => o.ha))} ha trabajadas${ha ? ` · ${eur(tot / ha, 2)} por ha de la explotación (${fmtN(ha)} ha)` : ''}</span><span class="muted">Solo cuenta las labores apuntadas con un equipo. El gasóleo sale al precio de los albaranes del almacén.</span></div>`;
    }
    return h;
  }
  function formEquipo(id) {
    const e = id ? equipo(id) : null;
    const info = v => {
      const c = costeEquipo(v); const r = parseNum(v.rendimiento);
      const filas = c.partes.map(([n, x]) => `<span>${esc(n)}: ${eur(x, 2)} por hora</span>`).join('');
      const pha = r > 0 && c.partes.length ? c.porHora / r : null;
      const gl = parseNum(v.gasoleo); const pr = precioGasoleo(100);
      const gas = gl > 0 ? (pr != null ? `<span>Gasóleo: ${fmtN(gl)} L/ha × ${fmtN(pr, 3)} €/L = <b>${eur(gl * pr, 2)} por hectárea</b></span>` : `<span class="neg">Gasóleo: sin precio. Mete en el almacén un albarán de gasóleo con su precio.</span>`) : '';
      const tot = pha != null && gl > 0 && pr != null ? `<span><b>Total: ${eur(pha + gl * pr, 2)} por hectárea</b></span>` : '';
      return `<b>Coste de amortización</b>${filas}${c.faltan.length ? `<span class="neg">Falta en la ficha de la máquina: ${esc(c.faltan.join(', '))}</span>` : ''}${pha != null ? `<span>Total ${eur(c.porHora, 2)}/h ÷ ${fmtN(r)} ha/h = <b>${eur(pha, 2)} por hectárea</b></span>` : '<span class="muted">Pon el rendimiento para calcular el coste por hectárea.</span>'}${gas}${tot}`;
    };
    formulario({
      titulo: e ? esc(e.nombre) : 'Nuevo equipo de trabajo',
      valores: e ? Object.assign({}, e) : { labor: '' },
      campos: [
        { k: 'nombre', label: 'Nombre', req: 1, placeholder: 'Ej.: Equipo de siembra' },
        { k: 'labor', label: 'Para la labor', type: 'select', opts: () => [['', '—'], ...TIPOS_LABOR.filter(t => t !== 'Otro').map(t => [t, t])], hint: 'Al apuntar esa labor, el equipo saldrá puesto.' },
        { k: 'tractorId', label: 'Tractor', type: 'select', half: 1, opts: () => optsMaq(true, e && e.tractorId) },
        { k: 'aperoId', label: 'Apero', type: 'select', half: 1, opts: () => optsMaq(false, e && e.aperoId) },
        { k: 'rendimiento', label: 'Rendimiento medio (ha/hora)', type: 'num', req: 1, half: 1, placeholder: 'Ej.: 3,5', hint: 'Hectáreas en una hora, contando giros y cargas.' },
        { k: 'gasoleo', label: 'Gasóleo (L/ha)', type: 'num', half: 1, placeholder: 'Ej.: 6', hint: 'Litros que gasta por hectárea.' },
        { k: 'notas', label: 'Notas', type: 'textarea' }
      ],
      extra: `<div class="card pad small" id="eqInfo" style="display:flex;flex-direction:column;gap:2px"></div>`,
      alCambiar(v, form) { const box = form.querySelector('#eqInfo'); if (box) box.innerHTML = info(v); },
      guardar(v) {
        if (!v.tractorId && !v.aperoId) return 'Elige al menos el tractor o el apero';
        if (!(v.rendimiento > 0)) return 'El rendimiento tiene que ser mayor que cero';
        if (e) Object.assign(e, v); else db.equipos.push(Object.assign({ id: uid() }, v));
      },
      textoBorrar: 'Se borrará el equipo. Las labores ya apuntadas conservan sus máquinas.',
      borrar: e && (() => { db.equipos = db.equipos.filter(x => x !== e); db.labores.forEach(l => { if (l.equipoId === e.id) l.equipoId = ''; }); })
    });
  }

  // Cambiar el nombre de un titular lo cambia en todo lo suyo
  function renombrarTitular(viejo, nuevo) {
    if (!viejo || !nuevo || viejo === nuevo) return;
    const es = x => String(x || '').trim() === viejo;
    db.fincas.forEach(f => { if (es(f.titular)) f.titular = nuevo; });
    db.documentos.forEach(d => { if (es(d.titular)) d.titular = nuevo; });
    db.ajustes.forEach(a => { if (es(a.titular)) a.titular = nuevo; });
    db.labores.forEach(l => { if (l.consumo && es(l.consumo.titular)) l.consumo.titular = nuevo; });
    db.maquinas.forEach(m => { if (es(m.titular)) m.titular = nuevo; });
    db.labores.forEach(l => { if (es(l.aplicador)) l.aplicador = nuevo; });
    if (es(db.ultAplicador)) db.ultAplicador = nuevo;
    if (ui.alm.titular === viejo) ui.alm.titular = nuevo;
    if (ui.exp.titular === viejo) ui.exp.titular = nuevo;
    if (ui.titular === viejo) ui.titular = nuevo;
  }

  function vExplotacion() {
    const e = ui.exp;
    let h = `<div class="segs" role="group" aria-label="Apartado">${[['agricultores', 'Agricultores'], ['maquinaria', 'Maquinaria']].map(([k, t]) => `<button data-act="expVista" data-k="${k}" aria-pressed="${e.vista === k}">${t}</button>`).join('')}</div>`;
    return h + (e.vista === 'maquinaria' ? vMaquinaria() : vAgricultores());
  }
  function vAgricultores() {
    const ts = titularesTodos();
    if (!ts.length) return vacio('Aún no hay agricultores', 'Añade los datos de cada titular: NIF, nº de explotación y carné de aplicador (ROPO).', '<button class="btn primary" data-act="nuevoAgricultor">Añadir agricultor</button>');
    const avisos = ts.map(fichaTitular).filter(f => f && ['vencido', 'pronto'].includes(estadoFecha(f.ropoCaduca, 90)));
    let h = avisos.length ? `<div class="banner"><p>${avisos.map(f => `Carné de aplicador de <b>${esc(f.nombre)}</b> ${estadoFecha(f.ropoCaduca, 90) === 'vencido' ? 'caducado el' : 'caduca el'} ${fmtFecha(f.ropoCaduca)}.`).join(' ')}</p></div>` : '';
    h += '<div class="list">' + ts.map(n => {
      const f = fichaTitular(n); const fs = fincasTitular(n); const nm = db.maquinas.filter(m => m.titular === n && !m.baja).length;
      const datos = f ? [f.nif ? `NIF ${f.nif}` : '', f.explotacion ? `Explotación ${f.explotacion}` : '', f.ropo ? `ROPO ${f.ropo}` : ''].filter(Boolean).join(' · ') : '';
      const falta = !f || !f.nif || !f.explotacion || !f.ropo;
      return `<button class="row" data-act="editAgricultor" data-n="${esc(n)}"><div class="main"><div class="t">${esc(f && f.nombreCompleto ? f.nombreCompleto : n)}${f && f.nombreCompleto && norm(f.nombreCompleto) !== norm(n) ? ` <span class="small muted">(${esc(n)})</span>` : ''}</div>
        <div class="s">${esc([`${fs.length} parcela${fs.length === 1 ? '' : 's'} · ${fmtN(sum(fs, x => x.superficie))} ha`, nm ? `${nm} máquina${nm === 1 ? '' : 's'}` : ''].filter(Boolean).join(' · '))}</div>
        ${datos ? `<div class="s">${esc(datos)}</div>` : ''}<div class="chips" style="margin-top:4px">${f ? pillFecha('Carné', f.ropoCaduca, 90) : ''}${falta ? '<span class="pill warn">Faltan datos</span>' : ''}</div></div></button>`;
    }).join('') + '</div><p class="note">Toca un agricultor para completar o corregir sus datos. Se guardan solo en este teléfono y en tus copias de seguridad.</p>';
    return h;
  }
  function vMaquinaria() {
    const tits = titularesTodos(); const tit = ui.exp.titular;
    let h = tits.length ? `<div class="carpetas" role="group" aria-label="Maquinaria de">${[['', 'Todos'], ...tits.map(t => [t, t])].map(([k, t]) => `<button data-act="expTitular" data-k="${esc(k)}" aria-pressed="${tit === k}">${k ? ico.carpeta : ''}${esc(t)}</button>`).join('')}</div>` : '';
    const ms = db.maquinas.filter(m => !tit || m.titular === tit);
    if (!tit && db.maquinas.length) h += bloqueEquipos() + '<h2 class="sec" style="margin-top:18px">Máquinas</h2>';
    if (!ms.length) return h + vacio('Sin maquinaria', 'Da de alta los tractores y aperos de cada titular: matrícula, ROMA, ITV, inspección ITEAF y seguro. Más adelante servirá para calcular costes por hectárea, mantenimiento y gasóleo.', '<button class="btn primary" data-act="nuevaMaquina">Añadir máquina</button>');
    const alertas = ms.filter(m => !m.baja).flatMap(m => avisosMaquina(m).filter(([, iso, dd]) => ['vencido', 'pronto'].includes(estadoFecha(iso, dd))).map(([t, iso, dd]) => `${esc(m.nombre)}: ${t} ${estadoFecha(iso, dd) === 'vencido' ? vencidoDe(t) + ' el' : 'vence el'} ${fmtFecha(iso)}`));
    if (alertas.length) h += `<div class="banner"><p>${alertas.join('<br>')}</p></div>`;
    const orden = (a, b) => TIPOS_MAQ.indexOf(a.tipo) - TIPOS_MAQ.indexOf(b.tipo) || a.nombre.localeCompare(b.nombre, 'es', { numeric: true });
    const fila = m => `<button class="row${m.baja ? ' baja' : ''}" data-act="editMaquina" data-id="${m.id}"><div class="main"><div class="t">${esc(m.nombre)}${m.baja ? ' <span class="pill muted">De baja</span>' : ''}</div>
      <div class="s">${esc([m.tipo, [m.marca, m.modelo].filter(Boolean).join(' '), m.anio, m.potencia ? `${fmtN(m.potencia, 0)} CV` : '', m.anchura ? `${fmtN(m.anchura, 2)} m` : '', m.deposito ? `${fmtN(m.deposito, 0)} L` : '', tit ? '' : m.titular].filter(Boolean).join(' · '))}</div>
      ${m.matricula || m.roma ? `<div class="s">${esc([m.matricula ? 'Matrícula ' + m.matricula : '', m.roma ? 'ROMA ' + m.roma : ''].filter(Boolean).join(' · '))}</div>` : ''}
      ${m.baja ? '' : `<div class="chips" style="margin-top:4px">${avisosMaquina(m).map(([t, iso, dd]) => iso ? pillFecha(t, iso, dd) : `<span class="pill muted">${t}: sin fecha</span>`).join('')}</div>`}</div>
      ${(() => { const tr = trabajoMaquina(m.id, db.campanaActiva); return tr.ha || m.contador ? `<div class="end small muted">${tr.ha ? `<b>${fmtN(tr.ha)} ha</b>` : ''}${tr.ha && m.contador ? '<br>' : ''}${m.contador ? fmtN(m.contador, 0) + ' h' : ''}</div>` : ''; })()}</button>`;
    const tipos = [...new Set(ms.map(m => m.tipo))].sort((a, b) => TIPOS_MAQ.indexOf(a) - TIPOS_MAQ.indexOf(b));
    h += tipos.map(t => `<h2 class="sec">${tit ? '' : '<span class="muted">· </span>'}${esc(t)}</h2><div class="list">${ms.filter(m => m.tipo === t).sort(orden).map(fila).join('')}</div>`).join('');
    return h + `<p class="note">Las hectáreas son las trabajadas en la campaña ${esc(db.campanaActiva)}. Cualquier máquina se puede usar en las fincas de cualquier titular: elígela al apuntar la labor.</p>`;
  }

  function formAgricultor(nombre) {
    const f = nombre ? fichaTitular(nombre) : null;
    const viejo = f ? f.nombre : nombre || '';
    const fs = viejo ? fincasTitular(viejo) : [];
    const maqs = viejo ? db.maquinas.filter(m => m.titular === viejo) : [];
    const resumen = viejo ? `<div class="card pad small" style="display:flex;flex-direction:column;gap:4px"><b>En la app</b><span>${fs.length} parcela${fs.length === 1 ? '' : 's'} · ${fmtN(sum(fs, x => x.superficie))} ha${maqs.length ? ` · ${maqs.length} máquina${maqs.length === 1 ? '' : 's'}` : ''}</span>${fs.length ? `<button type="button" class="btn sm" data-act="filtrarTitular" data-k="${esc(viejo)}">Ver sus parcelas</button>` : ''}</div>` : '';
    formulario({
      titulo: f || nombre ? `Agricultor: ${esc(viejo)}` : 'Nuevo agricultor',
      valores: f ? Object.assign({}, f) : { nombre: nombre || '' },
      campos: [
        { k: 'nombre', label: 'Nombre en la app', req: 1, half: 1, placeholder: 'Ej.: Ivan', hint: 'El que usas en parcelas y albaranes.' },
        { k: 'nombreCompleto', label: 'Nombre completo o razón social', half: 1, placeholder: 'Nombre y apellidos' },
        { k: 'nif', label: 'NIF', half: 1, placeholder: 'Ej.: 12345678Z', hint: 'DNI, NIE o CIF. La app comprueba la letra.' },
        { k: 'explotacion', label: 'Nº de explotación', half: 1, placeholder: 'Registro de explotaciones (REACyL)' },
        { k: 'ropo', label: 'Nº de inscripción ROPO', placeholder: 'Registro Oficial de Productores y Operadores' },
        { k: 'ropoNivel', label: 'Carné de aplicador', type: 'select', half: 1, opts: () => [['', '—'], ...NIVELES_ROPO.map(x => [x, x])] },
        { k: 'ropoCaduca', label: 'Caduca el', type: 'date', half: 1, hint: 'Te avisa 3 meses antes.' },
        { k: 'telefono', label: 'Teléfono', type: 'tel', half: 1 },
        { k: 'email', label: 'Correo', type: 'email', half: 1 },
        { k: 'direccion', label: 'Dirección' },
        { k: 'notas', label: 'Notas', type: 'textarea', placeholder: 'Banco, cuenta, gestoría, cooperativa…' }
      ],
      extra: resumen,
      guardar(v) {
        v.nif = String(v.nif || '').toUpperCase().replace(/[\s.\-]/g, '');
        if (nifValido(v.nif) === false) return 'El NIF no es correcto: revisa los números y la letra';
        if (v.email && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(v.email)) return 'El correo no parece correcto';
        const otro = fichaTitular(v.nombre);
        if (otro && otro !== f) return `Ya hay un agricultor llamado «${otro.nombre}»`;
        if (viejo && viejo !== v.nombre) renombrarTitular(viejo, v.nombre);
        if (f) Object.assign(f, v); else db.titulares.push(Object.assign({ id: uid() }, v));
      },
      textoBorrar: 'Se borrará la ficha de este agricultor (NIF, ROPO…). Sus parcelas, albaranes y máquinas no se tocan.',
      borrar: f && (() => { db.titulares = db.titulares.filter(x => x !== f); })
    });
  }

  function resumenTrabajoMaq(m) {
    const c = trabajoMaquina(m.id, db.campanaActiva), tot = trabajoMaquina(m.id);
    if (!tot.n) return '<div class="card pad small muted">Todavía no se ha usado en ninguna labor. Elígela al apuntar una labor en el cuaderno.</div>';
    const lista = xs => xs.map(([k, ha]) => `${esc(k)} <b>${fmtN(ha)} ha</b>`).join(' · ');
    return `<div class="card pad small" style="display:flex;flex-direction:column;gap:4px"><b>Trabajo en la campaña ${esc(db.campanaActiva)}</b>${c.n ? `<span>${c.n} parcela${c.n === 1 ? '' : 's'} trabajada${c.n === 1 ? '' : 's'} · <b>${fmtN(c.ha)} ha</b></span><span>Por labor: ${lista(c.porTipo)}</span><span>Tierras de: ${lista(c.porTit)}</span>` : '<span class="muted">Nada en esta campaña.</span>'}${tot.n !== c.n ? `<span class="muted">Todas las campañas: ${fmtN(tot.ha)} ha</span>` : ''}</div>`;
  }
  function formMaquina(id) {
    const m = id ? db.maquinas.find(x => x.id === id) : null;
    const ts = titularesTodos();
    formulario({
      titulo: m ? esc(m.nombre) : 'Nueva máquina',
      valores: m ? Object.assign({}, m, { baja: m.baja ? 'si' : 'no' }) : { titular: ui.exp.titular || (ts.length === 1 ? ts[0] : ''), tipo: 'Tractor', baja: 'no' },
      campos: [
        { k: 'titular', label: 'Titular', req: 1, half: 1, lista: titularesTodos, placeholder: 'Ej.: Ivan' },
        { k: 'tipo', label: 'Tipo', type: 'select', req: 1, half: 1, opts: () => TIPOS_MAQ.map(t => [t, t]) },
        { k: 'nombre', label: 'Nombre', req: 1, placeholder: 'Cómo la llamáis: «el John Deere grande», «sulfatadora nueva»…' },
        { k: 'marca', label: 'Marca', half: 1, placeholder: 'Ej.: John Deere' },
        { k: 'modelo', label: 'Modelo', half: 1, placeholder: 'Ej.: 6155R' },
        { k: 'matricula', label: 'Matrícula', half: 1, placeholder: 'Ej.: E1234BCD' },
        { k: 'roma', label: 'Nº ROMA', half: 1, placeholder: 'Registro Oficial de Maquinaria Agrícola' },
        { k: 'anio', label: 'Año', type: 'num', half: 1, placeholder: 'Ej.: 2018' },
        { k: 'potencia', label: 'Potencia (CV)', type: 'num', half: 1, show: v => ['Tractor', 'Cosechadora', 'Vehículo'].includes(v.tipo) },
        { k: 'anchura', label: 'Anchura de trabajo (m)', type: 'num', half: 1, show: v => CON_ANCHURA.includes(v.tipo), placeholder: 'Ej.: 24' },
        { k: 'deposito', label: 'Depósito (L)', type: 'num', miles: 1, half: 1, show: v => ['Pulverizador', 'Abonadora'].includes(v.tipo) },
        { k: 'contador', label: 'Horas de contador', type: 'num', miles: 1, half: 1, show: v => ['Tractor', 'Cosechadora', 'Vehículo'].includes(v.tipo), placeholder: 'Ej.: 4.250' },
        { k: 'itv', label: 'Próxima ITV', type: 'date', half: 1, show: v => CON_ITV.includes(v.tipo) },
        { k: 'iteaf', label: 'Próxima inspección ITEAF', type: 'date', half: 1, show: v => CON_ITEAF.includes(v.tipo), hint: 'Inspección de equipos de aplicación de fitosanitarios.' },
        { k: 'seguroCompania', label: 'Seguro (compañía)', half: 1 },
        { k: 'seguroVence', label: 'Seguro vence el', type: 'date', half: 1 },
        { k: 'fechaCompra', label: 'Fecha de compra', type: 'date', half: 1 },
        { k: 'precioCompra', label: 'Precio de compra (€)', type: 'num', miles: 1, half: 1 },
        { k: 'horasVida', label: 'Horas de amortización', type: 'num', miles: 1, half: 1, placeholder: 'Ej.: 12.000', hint: 'Horas de trabajo en las que la pagas.' },
        { k: 'valorResidual', label: 'Valor al venderla (€)', type: 'num', miles: 1, half: 1, placeholder: '0', hint: 'Lo que esperas sacar al final. Vacío = 0.' },
        { k: 'baja', label: 'Estado', type: 'select', opts: () => [['no', 'En uso'], ['si', 'Vendida o de baja']] },
        { k: 'notas', label: 'Notas', type: 'textarea', placeholder: 'Taller, recambios, averías conocidas…' }
      ],
      extra: '<div class="card pad small" id="maqCoste" hidden></div>' + (m ? resumenTrabajoMaq(m) : ''),
      alCambiar(v, form) {
        const box = form.querySelector('#maqCoste'); if (!box) return;
        const c = costeHoraMaq({ precioCompra: parseNum(v.precioCompra, true), horasVida: parseNum(v.horasVida, true), valorResidual: parseNum(v.valorResidual, true) });
        box.hidden = c == null; if (c != null) box.innerHTML = `Coste de amortización: <b>${eur(c, 2)} por hora</b>`;
      },
      guardar(v) {
        if (v.anio && (v.anio < 1900 || v.anio > new Date().getFullYear() + 1)) return 'El año no parece correcto';
        const x = Object.assign({}, v, { baja: v.baja === 'si', matricula: String(v.matricula || '').toUpperCase().replace(/\s/g, '') });
        if (m) Object.assign(m, x); else db.maquinas.push(Object.assign({ id: uid() }, x));
      },
      textoBorrar: 'Se borrará la máquina. Si la has vendido, mejor márcala «Vendida o de baja» para conservar su historial.',
      borrar: m && (() => { db.maquinas = db.maquinas.filter(x => x !== m); })
    });
  }

  // ---------- Eventos ----------
  const acciones = {
    verFinca: b => { ui.fincaId = b.dataset.id; empujar(); render(); window.scrollTo(0, 0); },
    atras: () => volverA(profundidad() - 1),
    nuevaFinca: () => formFinca(),
    editFinca: b => formFinca(b.dataset.id),
    asignar: b => formAsignar(b.dataset.id),
    nuevaLabor: b => formLabor(null, { finca: b.dataset.finca, grupo: b.dataset.grupo }),
    editLabor: b => formLabor(b.dataset.id),
    nuevoMov: b => formMov(null, { finca: b.dataset.finca, grupo: b.dataset.grupo, tipo: b.dataset.tipo }),
    editMov: b => formMov(b.dataset.id),
    nuevaCosecha: b => formCosecha(null, { finca: b.dataset.finca, grupo: b.dataset.grupo }),
    verGrupo: b => {
      const g = b.dataset.g; if (mapa) mapa.closePopup();
      if (ui.fincaId && ui.grupo === g) { volverA(profundidad() - 1); return; }
      if (ui.fincaId) { ui.fincaId = null; ui.grupo = g; } else { if (!ui.grupo) { ui.grupo = g; empujar(); } else ui.grupo = g; }
      render(); window.scrollTo(0, 0);
    },
    toggleGrupo: b => { const g = b.dataset.g; if (ui.abiertos.has(g)) ui.abiertos.delete(g); else ui.abiertos.add(g); const l = $('#listaFincas'); if (l) l.innerHTML = listaFincasHTML(); },
    agrupar: b => hojaAgrupar({ id: b.dataset.id }),
    editGrupo: b => hojaAgrupar({ grupo: b.dataset.g }),
    deshacerGrupo: async b => {
      const g = b.dataset.g;
      const r = await confirmar('¿Deshacer la agrupación?', `Las parcelas de «${esc(g)}» quedarán sueltas. No se borra nada de lo apuntado en ellas.`, 'Deshacer');
      if (!r.ok) return;
      fincasDeGrupo(g).forEach(f => { f.grupo = ''; });
      guardar(); volverA(0); render(); toast('Agrupación deshecha');
    },
    verMapaGrupo: b => { volverA(0); ui.fincaId = null; ui.grupo = null; ui.tab = 'fincas'; ui.vistaFincas = 'mapa'; ui.enfocarGrupo = b.dataset.g; render(); window.scrollTo(0, 0); },
    editCosecha: b => formCosecha(b.dataset.id),
    filtroMov: b => { ui.filtroMov = b.dataset.k; render(); },
    almVista: b => { ui.alm.vista = b.dataset.k; render(); },
    expVista: b => { ui.exp.vista = b.dataset.k; render(); },
    expTitular: b => { ui.exp.titular = b.dataset.k; render(); },
    nuevoAgricultor: () => formAgricultor(null),
    editAgricultor: b => formAgricultor(b.dataset.n),
    nuevaMaquina: () => formMaquina(null),
    editMaquina: b => formMaquina(b.dataset.id),
    nuevoEquipo: () => formEquipo(null),
    editEquipo: b => formEquipo(b.dataset.id),
    almTitular: b => { ui.alm.titular = b.dataset.k; render(); },
    nuevoDoc: b => formDocumento(b.dataset.k || 'albaran'),
    editDoc: b => formDocumento(null, b.dataset.id),
    verProducto: b => hojaProducto(b.dataset.id, b.dataset.t || ''),
    resumenVista: b => { ui.resumenVista = b.dataset.k; render(); },
    excelPAC: () => excelPAC(),
    filtrarTitular: b => { ui.titular = b.dataset.k === 'Sin titular' ? '' : b.dataset.k; ui.buscar = ''; irTab('fincas'); ui.vistaFincas = 'lista'; render(); },
    vistaFincas: b => { ui.vistaFincas = b.dataset.k; try { localStorage.setItem(KEY + '.vista', ui.vistaFincas); } catch (e) { /* ignorar */ } render(); },
    verMapa: b => { volverA(0); ui.fincaId = null; ui.grupo = null; ui.tab = 'fincas'; ui.vistaFincas = 'mapa'; ui.enfocar = b.dataset.id; render(); window.scrollTo(0, 0); },
    contornos: () => actualizarContornos({ forzar: true }),
    importarExcel: () => { cargarXLSX().catch(() => {}); $('#xlsIn').click(); },
    fincaDesdeRecinto: () => { const x = ui.recSel; if (!x) return; mapa.closePopup(); formFinca(null, { provincia: x.provincia, codMunicipio: x.municipio, agregado: x.agregado, zona: x.zona, poligono: String(x.poligono), parcela: String(x.parcela), recinto: String(x.recinto), superficie: x.superficie }); },
    anadirRecinto: () => { const x = ui.recSel; if (!x) return; mapa.closePopup(); formAnadirRecinto(x); },
    irTab: b => irTab(b.dataset.tab),
    ejemplo: () => cargarEjemplo(),
    compartirCopia: () => exportar(true),
    borrarEjemplo: async () => {
      const r = await confirmar('¿Borrar los datos de ejemplo?', 'Se quedará la app vacía para empezar con tus fincas.');
      if (!r.ok) return;
      db = nuevoDB(); guardar(false); ui.fincaId = null; ui.grupo = null; volverA(0); render();
    }
  };
  function irTab(t) {
    volverA(0);
    ui.fincaId = null; ui.grupo = null; ui.tab = t; render(); window.scrollTo(0, 0);
  }
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-act]');
    if (b && acciones[b.dataset.act]) acciones[b.dataset.act](b);
  });
  document.querySelectorAll('.tabs button').forEach(b => b.addEventListener('click', () => irTab(b.dataset.tab)));
  $('#main').addEventListener('input', e => {
    if (e.target.id === 'fBuscar') { ui.buscar = e.target.value; const l = $('#listaFincas'); if (l) l.innerHTML = listaFincasHTML(); }
  });
  $('#main').addEventListener('change', e => {
    if (e.target.id === 'fTitular') { ui.titular = e.target.value; const l = $('#listaFincas'); if (l) l.innerHTML = listaFincasHTML(); }
    if (e.target.id === 'fLabFinca') { ui.filtroLabor.finca = e.target.value; render(); }
    if (e.target.id === 'fLabTipo') { ui.filtroLabor.tipo = e.target.value; render(); }
  });
  $('#cfgBtn').addEventListener('click', hojaAjustes);
  $('#vozBtn').addEventListener('click', hojaVoz);
  $('#campBtn').addEventListener('click', hojaCampanas);

  let toastT;
  function toast(t) { const el = $('#toast'); el.textContent = t; el.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => { el.hidden = true; }, 2600); }

  // ---------- Arranque ----------
  try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist(); } catch (e) { /* ignorar */ }
  if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
    window.addEventListener('load', () => { navigator.serviceWorker.register('sw.js').catch(() => {}); });
  }
  try { const v = localStorage.getItem(KEY + '.vista'); if (v === 'mapa' || v === 'lista') ui.vistaFincas = v; } catch (e) { /* ignorar */ }
  render();
})();
