'use strict';
(function () {
  // ---------- Constantes ----------
  const KEY = 'cuadernoCampo.v1';
  const VERSION = '1.3.0';
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
  const TITULOS = { fincas: 'Fincas', cuaderno: 'Cuaderno de campo', cuentas: 'Cuentas', cosechas: 'Cosechas', almacen: 'Almacén', resumen: 'Resumen' };
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
  const fmtKg = v => (v == null || isNaN(v)) ? '—' : (v >= 10000 ? fmtN(v / 1000, 1) + ' t' : fmtN(v, 0) + ' kg');
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
    return { v: 1, campanaActiva: c, campanas: [c], cultivos: CULTIVOS_BASE.slice(), fincas: [], siembras: [], labores: [], movimientos: [], cosechas: [], productos: [], documentos: [], ajustes: [], ultimaCopia: null, ultimoCambio: null, ejemplo: false };
  }
  function normalizar(d) {
    const base = nuevoDB();
    const o = Object.assign(base, d && typeof d === 'object' ? d : {});
    for (const k of ['campanas', 'fincas', 'siembras', 'labores', 'movimientos', 'cosechas', 'productos', 'documentos', 'ajustes']) if (!Array.isArray(o[k])) o[k] = [];
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

  function statsFinca(fid, c) {
    const f = finca(fid); const ha = (f && f.superficie) || 0;
    const kg = sum(db.cosechas.filter(x => x.fincaId === fid && x.campana === c), x => x.kg);
    const movs = db.movimientos.filter(m => m.fincaId === fid && m.campana === c);
    const gastos = sum(movs.filter(m => m.tipo === 'gasto'), m => m.importe);
    const ingresos = sum(movs.filter(m => m.tipo === 'ingreso'), m => m.importe);
    return { ha, kg, kgHa: ha && kg ? kg / ha : null, gastos, ingresos, margen: ingresos - gastos, margenHa: ha && (gastos || ingresos) ? (ingresos - gastos) / ha : null };
  }
  function statsGrupo(fs, c) {
    let ha = 0, haCos = 0, kg = 0, gastos = 0, ingresos = 0;
    for (const f of fs) { const s = statsFinca(f.id, c); ha += s.ha; kg += s.kg; if (s.kg > 0) haCos += s.ha; gastos += s.gastos; ingresos += s.ingresos; }
    return { ha, kg, kgHa: haCos ? kg / haCos : null, gastos, ingresos, margen: ingresos - gastos, margenHa: ha && (gastos || ingresos) ? (ingresos - gastos) / ha : null };
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
    const gastos = sum(movs.filter(m => m.tipo === 'gasto'), m => m.importe);
    const haSemb = sum(lista.filter(o => o.cultivo !== 'Barbecho' && o.cultivo !== 'Sin cultivo'), o => o.ha);
    return { lista, genG, genI, ingresos, gastos, resultado: ingresos - gastos, kg: sum(lista, o => o.kg), haSemb };
  }

  // ---------- Estado de la interfaz y navegación (botón atrás de Android) ----------
  const ui = { tab: 'fincas', fincaId: null, grupo: null, buscar: '', titular: '', abiertos: new Set(), filtroLabor: { finca: '', tipo: '' }, filtroMov: 'todos', alm: { vista: 'inventario', titular: '' }, resumenVista: 'campana', vistaFincas: 'lista', enfocar: null, enfocarGrupo: null, mapaAjustado: false };
  const capas = [];
  const profundidad = () => capas.length + (ui.grupo ? 1 : 0) + (ui.fincaId ? 1 : 0);
  function empujar() { try { history.pushState({ d: profundidad() }, ''); } catch (e) { /* ignorar */ } }
  function deshacerHasta(d) {
    while (profundidad() > d) {
      if (capas.length) { const c = capas.pop(); c.el.remove(); }
      else if (ui.fincaId) { ui.fincaId = null; render(); }
      else { ui.grupo = null; render(); }
    }
  }
  // Cerrar desde la app: se cierra al momento y se ignora el popstate que provoca history.go.
  let ignorarPop = 0;
  function volverA(d) {
    const n = profundidad() - d; if (n <= 0) return;
    deshacerHasta(d);
    try { ignorarPop++; history.go(-n); } catch (e) { ignorarPop = 0; }
  }
  window.addEventListener('popstate', e => {
    if (ignorarPop > 0) { ignorarPop--; return; }
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
    const det = [l.producto, l.dosis != null ? `${fmtN(l.dosis)} ${l.unidad || ''}` : '', l.motivo].filter(Boolean).join(' · ');
    const coste = sum(n, x => x.coste);
    const aviso = n.some(x => x.avisos && x.avisos.length);
    return `<button class="row" data-act="editLabor" data-id="${l.id}">${fechaBox(l.fecha)}<div class="main"><div class="t">${esc(l.tipo)}${conFinca ? ' · ' + esc(nombreLote(n)) : ''}</div>${det ? `<div class="s">${esc(det)}</div>` : ''}${aviso ? '<span class="pill neg">⚠ Revisar normativa</span>' : ''}</div>${coste ? `<div class="end small muted">${eur(coste)}</div>` : ''}</button>`;
  }
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
      <div class="stats">${stat('Rendimiento', st.kgHa ? fmtN(st.kgHa, 0) + ' kg/ha' : '—')}${stat('Gastos', eur(st.gastos))}${stat('Margen', st.margenHa != null ? eur(st.margenHa) + '/ha' : '—', signo(st.margen))}</div>
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
      <div class="stats">${stat('Rendimiento', st.kgHa ? fmtN(st.kgHa, 0) + ' kg/ha' : '—')}${stat('Gastos', eur(st.gastos))}${stat('Margen', st.margenHa != null ? eur(st.margenHa) + '/ha' : '—', signo(st.margen))}</div>
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
    h += `<div class="segs" role="group" aria-label="Mostrar">${[['todos', 'Todo'], ['gasto', 'Gastos'], ['ingreso', 'Ingresos']].map(([k, t]) => `<button data-act="filtroMov" data-k="${k}" aria-pressed="${ui.filtroMov === k}">${t}</button>`).join('')}</div>`;
    const tipo = ui.filtroMov === 'ingreso' ? 'ingreso' : 'gasto';
    const porCat = {};
    db.movimientos.filter(m => m.campana === c && m.tipo === tipo).forEach(m => { porCat[m.categoria] = (porCat[m.categoria] || 0) + (m.importe || 0); });
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
    main.innerHTML = ui.fincaId ? vDetalle(ui.fincaId) : ui.grupo ? vGrupo(ui.grupo) : ({ fincas: vFincas, cuaderno: vCuaderno, cuentas: vCuentas, cosechas: vCosechas, almacen: vAlmacen, resumen: vResumen }[ui.tab])();
    const fab = { fincas: ['nuevaFinca', 'Finca'], cuaderno: ['nuevaLabor', 'Labor'], cuentas: ['nuevoMov', 'Apunte'], cosechas: ['nuevaCosecha', 'Cosecha'], almacen: ui.alm.vista === 'factura' ? ['nuevoDoc" data-k="factura', 'Factura'] : ['nuevoDoc" data-k="albaran', 'Albarán'] }[ui.tab];
    const mostrarFab = fab && !ui.fincaId && !ui.grupo && (db.fincas.length || ui.tab === 'almacen') && !(ui.tab === 'fincas' && ui.vistaFincas === 'mapa');
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

  function formLabor(id, pre = {}) {
    const l = id ? db.labores.find(x => x.id === id) : null;
    if (!db.fincas.length) { toast('Primero añade una finca'); return; }
    const miembros = l ? miembrosLote(db.labores, l) : [];
    const conDosis = v => CON_DOSIS.includes(v.tipo);
    const trat = v => v.tipo === 'Tratamiento';
    const una = v => (v.fincas || []).length <= 1;
    const viejo = l && l.tipo === 'Tratamiento' && !l.productoId; // labor antigua sin producto del almacén
    // lo que ya gastaba esta labor vuelve a estar disponible al editarla
    const propio = {}; miembros.forEach(x => { if (x.consumo) propio[x.consumo.titular] = (propio[x.consumo.titular] || 0) + (+x.consumo.cantidad || 0); });
    const ini = l
      ? Object.assign({}, l, { fincas: miembros.map(x => x.fincaId), prodAlm: l.productoId || '', coste: miembros.some(x => x.coste) ? Math.round(sum(miembros, x => x.coste) * 100) / 100 : null })
      : { fecha: hoyISO(), tipo: 'Labrado', fincas: idsIniciales(pre), unidad: 'kg/ha', campana: db.campanaActiva, prodAlm: '' };
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
        { k: 'dosis', label: 'Dosis', type: 'num', half: 1, show: conDosis },
        { k: 'unidad', label: 'Unidad', type: 'select', half: 1, opts: () => UNIDADES.map(u => [u, u]), show: conDosis },
        { k: 'registro', label: 'Nº de registro', half: 1, show: v => trat(v) && !v.prodAlm, placeholder: 'Registro del producto' },
        { k: 'motivo', label: 'Plaga o mala hierba', half: 1, show: trat, placeholder: 'Ej.: avena loca' },
        { k: 'superficie', label: 'Superficie trabajada (ha)', type: 'num', placeholder: 'Vacío = toda la finca', show: una },
        { k: 'coste', label: 'Coste total (€)', type: 'num', miles: 1, hint: 'Se apunta también como gasto. Con varias fincas se reparte según la superficie de cada una.' },
        campoCampana,
        { k: 'notas', label: 'Notas', type: 'textarea', placeholder: 'Maquinaria, condiciones del tiempo, quién la hizo…' }
      ],
      extra: '<div id="fitoInfo" class="fito-info" hidden></div>',
      alCambiar(v, form, origen) {
        refrescarInfo = () => alCambiarReal(v, form);
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
      guardar(v) {
        const fs = v.fincas.map(finca).filter(Boolean);
        const p = producto(v.prodAlm);
        const r = p ? revisarLabor(Object.assign({}, v, { _propio: propio })) : null;
        const ids = new Set(miembros.map(x => x.id));
        db.labores = db.labores.filter(x => !ids.has(x.id));
        db.movimientos = db.movimientos.filter(m => !ids.has(m.laborId));
        const lote = fs.length > 1 ? (l && l.lote) || uid() : null;
        const costes = repartir(v.coste, fs);
        fs.forEach((f, i) => {
          const x = Object.assign({}, v, { id: i === 0 && l ? l.id : uid(), fincaId: f.id, coste: costes[i], superficie: fs.length > 1 ? null : v.superficie });
          delete x.fincas; delete x.prodAlm; delete x._propio; if (lote) x.lote = lote; else delete x.lote;
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
          if (x.tipo === 'Siembra' && !siembraDe(x.fincaId, x.campana)) {
            const cu = db.cultivos.find(c => (x.producto || '').toLowerCase().includes(c.toLowerCase()));
            if (cu) db.siembras.push({ id: uid(), fincaId: x.fincaId, campana: x.campana, cultivo: cu, variedad: '' });
          }
        });
      },
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

  function formMov(id, pre = {}) {
    const m = id ? db.movimientos.find(x => x.id === id) : null;
    if (m && m.laborId && db.labores.some(l => l.id === m.laborId)) { formLabor(m.laborId); return; }
    const miembros = m ? miembrosLote(db.movimientos, m) : [];
    const tipo = pre.tipo || 'gasto';
    const venta = v => v.categoria === 'Venta de grano' || v.categoria === 'Venta de paja';
    const ini = m
      ? Object.assign({}, m, { fincas: miembros.map(x => x.fincaId).filter(id2 => finca(id2)), importe: Math.round(sum(miembros, x => x.importe) * 100) / 100, kilos: miembros.some(x => x.kilos) ? sum(miembros, x => x.kilos) : null })
      : { tipo, fecha: hoyISO(), categoria: CAT[tipo][0], fincas: pre.grupo || pre.finca ? idsIniciales(pre) : [], campana: db.campanaActiva };
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
        const fs = v.fincas.map(finca).filter(Boolean);
        db.movimientos = db.movimientos.filter(x => !miembros.includes(x));
        const base = Object.assign({}, v); delete base.fincas;
        if (!fs.length) { db.movimientos.push(Object.assign({ id: m ? m.id : uid() }, base, { fincaId: '' })); return; }
        const lote = fs.length > 1 ? (m && m.lote) || uid() : null;
        const imp = repartir(v.importe, fs), kg = v.kilos ? repartir(v.kilos, fs) : fs.map(() => null);
        fs.forEach((f, i) => { const x = Object.assign({ id: i === 0 && m ? m.id : uid() }, base, { fincaId: f.id, importe: imp[i], kilos: kg[i] }); if (lote) x.lote = lote; db.movimientos.push(x); });
      },
      borrar: m && (() => { db.movimientos = db.movimientos.filter(x => !miembros.includes(x)); })
    });
  }

  function formCosecha(id, pre = {}) {
    const x = id ? db.cosechas.find(c => c.id === id) : null;
    if (!db.fincas.length) { toast('Primero añade una finca'); return; }
    const miembros = x ? miembrosLote(db.cosechas, x) : [];
    const ini = x
      ? Object.assign({}, x, { fincas: miembros.map(y => y.fincaId), kg: Math.round(sum(miembros, y => y.kg)) })
      : { fecha: hoyISO(), fincas: idsIniciales(pre), campana: db.campanaActiva };
    formulario({
      titulo: x ? (miembros.length > 1 ? `Editar cosecha (${miembros.length} parcelas)` : 'Editar cosecha') : 'Apuntar cosecha',
      valores: ini,
      campos: [
        { k: 'fincas', label: 'Fincas', type: 'fincas', req: 1, hint: 'Si eliges varias, los kilos se reparten según la superficie de cada una.' },
        { k: 'fecha', label: 'Fecha', type: 'date', req: 1, half: 1 },
        { k: 'kg', label: 'Kilos', type: 'num', miles: 1, req: 1, half: 1, placeholder: 'Ej.: 12.400' },
        { k: 'humedad', label: 'Humedad (%)', type: 'num', placeholder: 'Ej.: 11,5' },
        campoCampana,
        { k: 'notas', label: 'Destino o notas', type: 'textarea', placeholder: 'Cooperativa, almacén, nº de ticket de báscula…' }
      ],
      guardar(v) {
        if (v.kg <= 0) return 'Los kilos tienen que ser mayores que cero';
        const fs = v.fincas.map(finca).filter(Boolean);
        db.cosechas = db.cosechas.filter(c => !miembros.includes(c));
        const lote = fs.length > 1 ? (x && x.lote) || uid() : null;
        const kgs = repartir(v.kg, fs);
        const base = Object.assign({}, v); delete base.fincas;
        fs.forEach((f, i) => { const c = Object.assign({ id: i === 0 && x ? x.id : uid() }, base, { fincaId: f.id, kg: kgs[i] }); if (lote) c.lote = lote; db.cosechas.push(c); });
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
      <div class="list">${listaCampanas().map(c => `<button class="row" data-camp="${c}"><div class="main"><div class="t">${c}</div></div>${c === db.campanaActiva ? '<span class="tag">Activa</span>' : ''}</button>`).join('')}</div>
      <button class="btn block" data-nueva>${ico.plus} Añadir campaña ${siguienteCampana(listaCampanas()[0])}</button></div></div>`, (el, d) => {
      el.addEventListener('click', e => {
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
  const CULTIVO_REG = {
    trigo: ['Trigo', 'Cereales de invierno / primavera'], cebada: ['Cebada', 'Cereales de invierno / primavera'],
    avena: ['Avena', 'Cereales de invierno / primavera'], 'avena forrajera': ['Avena', 'Cereales de invierno / primavera'],
    centeno: ['Centeno', 'Cereales de invierno / primavera'], triticale: ['Triticale', 'Cereales de invierno / primavera'],
    lenteja: ['Lenteja', 'Leguminosas de grano'], garbanzo: ['Garbanzo', 'Leguminosas de grano'], guisante: ['Guisante para grano', 'Leguminosas de grano'],
    titarro: ['Almorta', 'Leguminosas de grano'], yeros: ['Yeros', 'Leguminosas forrajeras'], veza: ['Veza', 'Leguminosas forrajeras'],
    esparceta: ['Esparceta', 'Leguminosas forrajeras'], alfalfa: ['Alfalfa', 'Leguminosas forrajeras'],
    'anis dulce': ['Anís, matalauva'], anis: ['Anís, matalauva'], cilantro: ['Cilantro'], girasol: ['Girasol'], colza: ['Colza'], camelina: ['Camelina'],
    barbecho: ['Barbechos'], chopo: ['Chopos y álamos']
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
  const TIPOS_PROD = ['Fitosanitario', 'Abono', 'Semilla', 'Otro'];
  const UNI_PROD = ['L', 'kg', 't', 'ud'];
  const TIPO_DE_LABOR = { Tratamiento: 'Fitosanitario', Abonado: 'Abono', Siembra: 'Semilla' };
  const producto = id => db.productos.find(p => p.id === id);
  const titularesTodos = () => [...new Set([...titulares(), ...db.documentos.map(d => d.titular), ...db.ajustes.map(a => a.titular)].map(t => String(t || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es'));
  const titularDe = fid => String((finca(fid) || {}).titular || '').trim();
  function movimientosProducto(pid, tit) {
    const out = [];
    db.documentos.filter(d => d.tipo === 'albaran' && (tit == null || d.titular === tit)).forEach(d => (d.lineas || []).forEach(l => { if (l.productoId === pid) out.push({ fecha: d.fecha, cant: +l.cantidad || 0, txt: `Albarán ${d.numero || ''} · ${d.proveedor || ''}`.trim(), titular: d.titular, doc: d.id }); }));
    db.labores.filter(l => l.consumo && l.consumo.productoId === pid && (tit == null || l.consumo.titular === tit)).forEach(l => out.push({ fecha: l.fecha, cant: -(+l.consumo.cantidad || 0), txt: `${l.tipo} · ${(finca(l.fincaId) || {}).nombre || ''}`, titular: l.consumo.titular, labor: l.id }));
    db.ajustes.filter(a => a.productoId === pid && (tit == null || a.titular === tit)).forEach(a => out.push({ fecha: a.fecha, cant: +a.cantidad || 0, txt: `Ajuste${a.motivo ? ' · ' + a.motivo : ''}`, titular: a.titular, ajuste: a.id }));
    return out.sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
  }
  const existencias = (pid, tit) => Math.round(sum(movimientosProducto(pid, tit), m => m.cant) * 1000) / 1000;
  const fmtCant = (v, u) => `${fmtN(v, 3)} ${u || ''}`.trim();
  function productosConMovimiento(tit) {
    const ids = new Set();
    db.documentos.filter(d => d.tipo === 'albaran' && (tit == null || d.titular === tit)).forEach(d => (d.lineas || []).forEach(l => ids.add(l.productoId)));
    db.labores.forEach(l => { if (l.consumo && (tit == null || l.consumo.titular === tit)) ids.add(l.consumo.productoId); });
    db.ajustes.forEach(a => { if (tit == null || a.titular === tit) ids.add(a.productoId); });
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
          h += `<h2 class="sec">Registro oficial de fitosanitarios</h2>${hits.length ? `<div class="list">${hits.map(p => `<button class="row" data-ep="reg" data-id="${p.id}"><div class="main"><div class="t">${esc(p.n)}</div><div class="s">${esc(['Nº reg. ' + p.r, p.f, p.e].filter(Boolean).join(' · '))}</div></div><span class="pill ${p.e === 'Vigente' ? 'pos' : 'neg'}">${esc(p.e)}</span></button>`).join('')}</div>` : '<p class="note">Ningún fitosanitario autorizado en tus cultivos coincide.</p>'}`;
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
  function formDocumento(tipo, id) {
    const doc = id ? db.documentos.find(x => x.id === id) : null;
    if (doc) tipo = doc.tipo;
    const tits = titularesTodos();
    const v = doc ? JSON.parse(JSON.stringify(doc)) : { tipo, titular: ui.alm.titular || (tits.length === 1 ? tits[0] : ''), fecha: hoyISO(), proveedor: '', numero: '', lineas: [], archivos: [], albaranes: [], iva: tipo === 'factura' ? 10 : null, notas: '' };
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
            const cat = { Fitosanitario: 'Fitosanitarios', Abono: 'Abono', Semilla: 'Semilla' }[tipos.sort((a, b) => tipos.filter(t => t === b).length - tipos.filter(t => t === a).length)[0]] || 'Otros gastos';
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
        <h2 class="sec">Movimientos${tit ? ' · ' + esc(tit) : ''}</h2>${movs.length ? `<div class="list">${movs.map(m => `<div class="row">${fechaBox(m.fecha)}<div class="main"><div class="t">${esc(m.txt)}</div><div class="s">${esc(m.titular || '')}</div></div><div class="end big ${m.cant < 0 ? 'neg' : 'pos'}">${m.cant > 0 ? '+' : ''}${fmtCant(m.cant, p.unidad)}</div></div>`).join('')}</div>` : '<p class="note">Sin movimientos.</p>'}`;
    };
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera(esc(p.nombre))}<div class="body" id="phBody">${pintar()}</div></div>`, (el, d) => {
      const body = el.querySelector('#phBody');
      cargarFitos().then(() => { body.innerHTML = pintar(); }).catch(() => {});
      el.addEventListener('click', e => {
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
