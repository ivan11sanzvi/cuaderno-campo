'use strict';
(function () {
  // ---------- Constantes ----------
  const KEY = 'cuadernoCampo.v1';
  const VERSION = '1.1.0';
  const TIPOS_LABOR = ['Labrado', 'Siembra', 'Abonado', 'Tratamiento', 'Rodillo', 'Cosecha', 'Empacado', 'Otro'];
  const CON_DOSIS = ['Siembra', 'Abonado', 'Tratamiento'];
  const UNIDADES = ['kg/ha', 'L/ha', 'g/ha', 'cc/ha', 'kg', 'L'];
  const CAT = {
    gasto: ['Semilla', 'Abono', 'Fitosanitarios', 'Gasóleo', 'Maquinaria y reparaciones', 'Servicios contratados', 'Seguro', 'Arrendamiento', 'Otros gastos'],
    ingreso: ['Venta de grano', 'Venta de paja', 'Subvención PAC', 'Otros ingresos']
  };
  const CAT_LABOR = { Siembra: 'Semilla', Abonado: 'Abono', Tratamiento: 'Fitosanitarios' };
  const CULTIVOS_BASE = ['Trigo', 'Cebada', 'Titarro', 'Lenteja', 'Anís dulce', 'Cilantro', 'Esparceta', 'Barbecho'];
  const HUES = { Trigo: 40, Cebada: 55, Titarro: 150, Lenteja: 18, 'Anís dulce': 85, Cilantro: 120, Esparceta: 320, Barbecho: 30 };
  const MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
  const MESES_L = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'];
  const TITULOS = { fincas: 'Fincas', cuaderno: 'Cuaderno de campo', cuentas: 'Cuentas', cosechas: 'Cosechas', resumen: 'Resumen' };
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
  const byNombre = (a, b) => a.nombre.localeCompare(b.nombre, 'es');
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
    return { v: 1, campanaActiva: c, campanas: [c], cultivos: CULTIVOS_BASE.slice(), fincas: [], siembras: [], labores: [], movimientos: [], cosechas: [], ultimaCopia: null, ultimoCambio: null, ejemplo: false };
  }
  function normalizar(d) {
    const base = nuevoDB();
    const o = Object.assign(base, d && typeof d === 'object' ? d : {});
    for (const k of ['campanas', 'fincas', 'siembras', 'labores', 'movimientos', 'cosechas']) if (!Array.isArray(o[k])) o[k] = [];
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
    for (const f of pend) {
      try {
        const g = await contornoFinca(f);
        f.geo = g;
        if (g.c.length) { ok++; if (!(f.superficie > 0) && g.ha) f.superficie = g.ha; } else no++;
      } catch (e) { fallo++; fallidos.add(refsKey(f)); }
    }
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
  const ui = { tab: 'fincas', fincaId: null, filtroLabor: { finca: '', tipo: '' }, filtroMov: 'todos', vistaFincas: 'lista', enfocar: null, mapaAjustado: false };
  const capas = [];
  const profundidad = () => capas.length + (ui.fincaId ? 1 : 0);
  function empujar() { try { history.pushState({ d: profundidad() }, ''); } catch (e) { /* ignorar */ } }
  function deshacerHasta(d) {
    while (profundidad() > d) {
      if (capas.length) { const c = capas.pop(); c.el.remove(); }
      else { ui.fincaId = null; render(); }
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
  function filaLabor(l, conFinca = true) {
    const f = finca(l.fincaId);
    const det = [l.producto, l.dosis != null ? `${fmtN(l.dosis)} ${l.unidad || ''}` : '', l.motivo].filter(Boolean).join(' · ');
    return `<button class="row" data-act="editLabor" data-id="${l.id}">${fechaBox(l.fecha)}<div class="main"><div class="t">${esc(l.tipo)}${conFinca ? ' · ' + esc(f ? f.nombre : 'Finca borrada') : ''}</div>${det ? `<div class="s">${esc(det)}</div>` : ''}</div>${l.coste ? `<div class="end small muted">${eur(l.coste)}</div>` : ''}</button>`;
  }
  function filaMov(m, conFinca = true) {
    const f = finca(m.fincaId);
    const sub = [m.categoria, conFinca ? (f ? f.nombre : 'General') : '', m.laborId ? 'desde el cuaderno' : ''].filter(Boolean).join(' · ');
    const v = m.tipo === 'gasto' ? -m.importe : m.importe;
    return `<button class="row" data-act="editMov" data-id="${m.id}">${fechaBox(m.fecha)}<div class="main"><div class="t">${esc(m.concepto || m.categoria)}</div><div class="s">${esc(sub)}</div></div><div class="end big ${signo(v)}">${v > 0 ? '+' : ''}${eur(v, 2)}</div></button>`;
  }
  function filaCosecha(x, conFinca = true) {
    const f = finca(x.fincaId);
    const sub = [x.humedad != null ? `${fmtN(x.humedad, 1)} % humedad` : '', x.notas].filter(Boolean).join(' · ');
    return `<button class="row" data-act="editCosecha" data-id="${x.id}">${fechaBox(x.fecha)}<div class="main"><div class="t">${conFinca ? esc(f ? f.nombre : 'Finca borrada') : 'Cosecha'}</div>${sub ? `<div class="s">${esc(sub)}</div>` : ''}</div><div class="end big">${fmtKg(x.kg)}</div></button>`;
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
    h += `<div class="stats">${stat('Fincas', db.fincas.length)}${stat('Superficie', fmtN(ha) + ' ha')}${stat('Sembrado ' + c, fmtN(sembrado) + ' ha')}</div>`;
    h += '<div class="list">' + db.fincas.slice().sort(byNombre).map(f => {
      const sig = [f.municipio, (f.poligono || f.parcela) ? `Pol. ${f.poligono || '–'} · Parc. ${f.parcela || '–'}` : '', fmtN(f.superficie) + ' ha'].filter(Boolean).join(' · ');
      return `<button class="row" data-act="verFinca" data-id="${f.id}"><div class="main"><div class="t">${esc(f.nombre)}</div><div class="s">${esc(sig)}</div></div><div class="end">${chip(cultivoDe(f.id, c))}</div></button>`;
    }).join('') + '</div>';
    return h;
  }

  function vDetalle(id) {
    const f = finca(id); if (!f) { ui.fincaId = null; return vFincas(); }
    const c = db.campanaActiva; const s = siembraDe(id, c); const st = statsFinca(id, c);
    const sig = [f.poligono && `Pol. ${f.poligono}`, f.parcela && `Parc. ${f.parcela}`, f.recinto && `Rec. ${f.recinto}`].filter(Boolean).join(' · ');
    const refs = refsFinca(f);
    let h = `<div class="backhdr"><button class="iconbtn" data-act="atras" aria-label="Volver">${ico.back}</button><h2>${esc(f.nombre)}</h2><button class="iconbtn" data-act="editFinca" data-id="${id}" aria-label="Editar finca">${ico.edit}</button></div>
      <div class="meta"><span><b>${fmtN(f.superficie)} ha</b></span>${f.municipio ? `<span>${esc(f.municipio)}</span>` : ''}${sig ? `<span>SIGPAC ${esc(sig)}</span>` : ''}</div>`;
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
      <div class="sp"><div><div class="tag">Campaña ${c}</div><div style="margin-top:4px">${chip(s && s.cultivo)} ${s && s.variedad ? `<span class="muted small">${esc(s.variedad)}</span>` : ''}</div></div>
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
      h += `<h2 class="sec">Rotación e historial</h2><div class="tablewrap"><table><thead><tr><th>Campaña</th><th>Cultivo</th><th>kg/ha</th><th>Margen €/ha</th></tr></thead><tbody>${camps.map(cc => {
        const x = statsFinca(id, cc);
        return `<tr><td>${cc}</td><td style="text-align:left">${chip(cultivoDe(id, cc))}</td><td>${x.kgHa ? fmtN(x.kgHa, 0) : '—'}</td><td class="${signo(x.margen)}">${x.margenHa != null ? fmtN(x.margenHa, 0) : '—'}</td></tr>`;
      }).join('')}</tbody></table></div>`;
    }
    if (f.notas) h += `<h2 class="sec">Notas</h2><div class="card pad" style="white-space:pre-wrap">${esc(f.notas)}</div>`;
    return h;
  }

  function vCuaderno() {
    const c = db.campanaActiva; const fl = ui.filtroLabor;
    if (!db.fincas.length) return vacio('Primero añade tus fincas', 'Las labores se apuntan siempre en una finca.', '<button class="btn primary" data-act="nuevaFinca">Añadir finca</button>');
    let lab = db.labores.filter(l => l.campana === c && (!fl.finca || l.fincaId === fl.finca) && (!fl.tipo || l.tipo === fl.tipo)).sort(byFechaDesc);
    let h = `<div class="filters"><select id="fLabFinca" aria-label="Filtrar por finca"><option value="">Todas las fincas</option>${db.fincas.slice().sort(byNombre).map(f => `<option value="${f.id}"${fl.finca === f.id ? ' selected' : ''}>${esc(f.nombre)}</option>`).join('')}</select>
      <select id="fLabTipo" aria-label="Filtrar por labor"><option value="">Todas las labores</option>${TIPOS_LABOR.map(t => `<option${fl.tipo === t ? ' selected' : ''}>${t}</option>`).join('')}</select></div>`;
    if (!lab.length) return h + vacio(`Sin labores en ${c}`, 'Apunta cada vez que labres, siembres, abones o trates una finca. Con fecha, producto y dosis tendrás el cuaderno de campo al día.', '<button class="btn primary" data-act="nuevaLabor">Apuntar labor</button>');
    let mes = '';
    let lista = '';
    for (const l of lab) {
      const m = l.fecha.slice(0, 7);
      if (m !== mes) { if (mes) lista += '</div>'; mes = m; lista += `<h2 class="sec">${MESES_L[+m.slice(5) - 1]} ${m.slice(0, 4)}</h2><div class="list">`; }
      lista += filaLabor(l);
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
    return h + `<h2 class="sec">Apuntes</h2><div class="list">${movs.map(m => filaMov(m)).join('')}</div>`;
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
    if (ent.length) h += `<h2 class="sec">Entradas</h2><div class="list">${ent.map(x => filaCosecha(x)).join('')}</div>`;
    return h;
  }

  function vResumen() {
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
      lyr.bindTooltip(esc(f.nombre), { permanent: true, direction: 'center', className: 'etq' });
      lyr.on('click', ev => { L.DomEvent.stopPropagation(ev); globoFinca(f, ev.latlng); });
      lyr.addTo(capaFincas);
    });
  }
  function globoFinca(f, ll) {
    const cu = cultivoDe(f.id, db.campanaActiva); const st = statsFinca(f.id, db.campanaActiva);
    L.popup({ maxWidth: 260 }).setLatLng(ll).setContent(`<div class="pop"><div class="t">${esc(f.nombre)}</div><div>${chip(cu)} <span class="muted">${fmtN(f.superficie)} ha</span></div>
      ${st.kgHa ? `<div class="small">${fmtN(st.kgHa, 0)} kg/ha en ${db.campanaActiva}</div>` : ''}${f.geo && f.geo.ha ? `<div class="small muted">SIGPAC: ${fmtN(f.geo.ha)} ha</div>` : ''}
      <button class="btn sm primary" data-act="verFinca" data-id="${f.id}">Abrir ficha</button></div>`).openOn(mapa);
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
    } else if (!ui.mapaAjustado) { ajustarMapa(false); ui.mapaAjustado = capaFincas.getLayers().length > 0; }
    setTimeout(() => mapa.invalidateSize(), 50);
    actualizarContornos({ silencioso: true });
  }

  // ---------- Render principal ----------
  function render() {
    $('#campBtn').textContent = db.campanaActiva + ' ▾';
    $('#title').textContent = ui.fincaId ? 'Finca' : TITULOS[ui.tab];
    document.querySelectorAll('.tabs button').forEach(b => b.setAttribute('aria-current', b.dataset.tab === ui.tab ? 'page' : 'false'));
    const main = $('#main');
    main.innerHTML = ui.fincaId ? vDetalle(ui.fincaId) : ({ fincas: vFincas, cuaderno: vCuaderno, cuentas: vCuentas, cosechas: vCosechas, resumen: vResumen }[ui.tab])();
    const fab = { fincas: ['nuevaFinca', 'Finca'], cuaderno: ['nuevaLabor', 'Labor'], cuentas: ['nuevoMov', 'Apunte'], cosechas: ['nuevaCosecha', 'Cosecha'] }[ui.tab];
    const mostrarFab = fab && !ui.fincaId && db.fincas.length && !(ui.tab === 'fincas' && ui.vistaFincas === 'mapa');
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

  // spec: {titulo, campos, valores, guardar(v) -> string|undefined (error), borrar()}
  function formulario(spec) {
    const vals = Object.assign({}, spec.valores);
    const campoHTML = f => {
      const id = 'f_' + f.k; const v = vals[f.k];
      let inp;
      if (f.type === 'select') inp = `<select id="${id}" name="${f.k}">${f.opts(vals).map(([val, t]) => `<option value="${esc(val)}"${String(v ?? '') === String(val) ? ' selected' : ''}>${esc(t)}</option>`).join('')}</select>`;
      else if (f.type === 'textarea') inp = `<textarea id="${id}" name="${f.k}" rows="3" placeholder="${esc(f.placeholder || '')}">${esc(v)}</textarea>`;
      else if (f.type === 'num') inp = `<input id="${id}" name="${f.k}" type="text" inputmode="decimal" autocomplete="off" value="${esc(numStr(v))}" placeholder="${esc(f.placeholder || '')}">`;
      else inp = `<input id="${id}" name="${f.k}" type="${f.type || 'text'}" value="${esc(v)}" placeholder="${esc(f.placeholder || '')}" autocomplete="off"${f.type === 'text' || !f.type ? ' autocapitalize="sentences"' : ''}>`;
      return `<div class="field" data-k="${f.k}"><label for="${id}">${f.label}${f.req ? '' : ' <span class="muted small">(opcional)</span>'}</label>${inp}${f.hint ? `<div class="hint">${f.hint}</div>` : ''}<div class="err" hidden></div></div>`;
    };
    let cuerpo = ''; const cs = spec.campos;
    for (let i = 0; i < cs.length; i++) {
      if (cs[i].half && cs[i + 1] && cs[i + 1].half) { cuerpo += `<div class="grid2" data-pair="${cs[i].k}">${campoHTML(cs[i])}${campoHTML(cs[i + 1])}</div>`; i++; }
      else cuerpo += campoHTML(cs[i]);
    }
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera(spec.titulo)}<form class="body" id="frm" novalidate>${cuerpo}${spec.extra || ''}</form>
      <footer>${spec.borrar ? '<button class="btn danger" type="button" data-del>Borrar</button>' : ''}<button class="btn primary" type="submit" form="frm">Guardar</button></footer></div>`, (el, d) => {
      const form = el.querySelector('form');
      const leer = () => { for (const f of cs) { const i = form.elements[f.k]; if (i) vals[f.k] = i.value; } return vals; };
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
      };
      form.addEventListener('change', e => refrescar(e.target.name));
      form.addEventListener('input', e => { if (spec.alEscribir) spec.alEscribir(e.target.name, leer(), (k, v) => { form.elements[k].value = v; }); });
      refrescar();
      form.addEventListener('submit', e => {
        e.preventDefault(); leer();
        const out = {}; let ok = true;
        for (const f of cs) {
          const w = form.querySelector(`.field[data-k="${f.k}"]`); const err = w.querySelector('.err'); err.hidden = true;
          if (w.hidden || (w.parentElement && w.parentElement.hidden)) { out[f.k] = f.type === 'num' ? null : ''; continue; }
          let v = vals[f.k];
          if (f.type === 'num') {
            v = parseNum(v, f.miles);
            if (Number.isNaN(v)) { err.textContent = 'Escribe un número, por ejemplo 12,5'; err.hidden = false; ok = false; continue; }
          } else v = String(v || '').trim();
          if (f.req && (v === null || v === '')) { err.textContent = 'Este dato es obligatorio'; err.hidden = false; ok = false; continue; }
          out[f.k] = v;
        }
        if (!ok) { const first = form.querySelector('.err:not([hidden])'); if (first) first.parentElement.scrollIntoView({ block: 'center', behavior: 'smooth' }); return; }
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
        { k: 'nombre', label: 'Nombre', req: 1, placeholder: 'Ej.: El Llano' },
        { k: 'superficie', label: 'Superficie (ha)', type: 'num', req: 1, placeholder: 'Ej.: 12,45' },
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
      valores: { campana: c, cultivo: s ? s.cultivo : db.cultivos[0], variedad: s ? s.variedad : '' },
      campos: [
        { k: 'cultivo', label: 'Cultivo', type: 'select', req: 1, opts: () => [...db.cultivos.map(x => [x, x]), ['__ninguno', 'Sin cultivo (quitar)']] },
        { k: 'variedad', label: 'Variedad', placeholder: 'Ej.: Marcopolo, Pedrezuela, Castellana…', show: v => v.cultivo !== '__ninguno' && v.cultivo !== 'Barbecho' },
        campoCampana
      ],
      guardar(v) {
        db.siembras = db.siembras.filter(x => !(x.fincaId === fid && x.campana === v.campana));
        if (v.cultivo !== '__ninguno') db.siembras.push({ id: uid(), fincaId: fid, campana: v.campana, cultivo: v.cultivo, variedad: v.variedad });
      }
    });
  }

  function formLabor(id, fincaPre) {
    const l = id ? db.labores.find(x => x.id === id) : null;
    if (!db.fincas.length) { toast('Primero añade una finca'); return; }
    const conDosis = v => CON_DOSIS.includes(v.tipo);
    const trat = v => v.tipo === 'Tratamiento';
    formulario({
      titulo: l ? 'Editar labor' : 'Nueva labor',
      valores: l || { fecha: hoyISO(), tipo: 'Labrado', fincaId: fincaPre || (ui.filtroLabor.finca || optsFincas()[0][0]), unidad: 'kg/ha', campana: db.campanaActiva },
      campos: [
        { k: 'fecha', label: 'Fecha', type: 'date', req: 1, half: 1 },
        { k: 'tipo', label: 'Labor', type: 'select', req: 1, half: 1, opts: () => TIPOS_LABOR.map(t => [t, t]) },
        { k: 'fincaId', label: 'Finca', type: 'select', req: 1, opts: optsFincas },
        { k: 'producto', label: 'Producto, variedad o apero', placeholder: 'Ej.: Urea 46 %, cultivador, cebada Pedrezuela' },
        { k: 'dosis', label: 'Dosis', type: 'num', half: 1, show: conDosis },
        { k: 'unidad', label: 'Unidad', type: 'select', half: 1, opts: () => UNIDADES.map(u => [u, u]), show: conDosis },
        { k: 'registro', label: 'Nº de registro', half: 1, show: trat, placeholder: 'Registro del producto' },
        { k: 'motivo', label: 'Plaga o mala hierba', half: 1, show: trat, placeholder: 'Ej.: avena loca' },
        { k: 'superficie', label: 'Superficie trabajada (ha)', type: 'num', placeholder: 'Vacío = toda la finca' },
        { k: 'coste', label: 'Coste (€)', type: 'num', miles: 1, hint: 'Si lo rellenas, se apunta también como gasto de la finca.' },
        campoCampana,
        { k: 'notas', label: 'Notas', type: 'textarea', placeholder: 'Maquinaria, condiciones del tiempo, quién la hizo…' }
      ],
      guardar(v) {
        const x = l || { id: uid() };
        Object.assign(x, v);
        if (!l) db.labores.push(x);
        let m = db.movimientos.find(mm => mm.laborId === x.id);
        if (x.coste > 0) {
          if (!m) { m = { id: uid(), laborId: x.id, tipo: 'gasto' }; db.movimientos.push(m); }
          Object.assign(m, { fecha: x.fecha, fincaId: x.fincaId, campana: x.campana, categoria: CAT_LABOR[x.tipo] || 'Servicios contratados', concepto: x.tipo + (x.producto ? ' · ' + x.producto : ''), importe: x.coste });
        } else if (m) db.movimientos = db.movimientos.filter(mm => mm !== m);
        if (x.tipo === 'Siembra' && !siembraDe(x.fincaId, x.campana)) {
          const cu = db.cultivos.find(c => (x.producto || '').toLowerCase().includes(c.toLowerCase()));
          if (cu) db.siembras.push({ id: uid(), fincaId: x.fincaId, campana: x.campana, cultivo: cu, variedad: '' });
        }
      },
      textoBorrar: 'Se borrará la labor y, si tenía coste, también su gasto.',
      borrar: l && (() => { db.labores = db.labores.filter(x => x !== l); db.movimientos = db.movimientos.filter(m => m.laborId !== l.id); })
    });
  }

  function formMov(id, pre = {}) {
    const m = id ? db.movimientos.find(x => x.id === id) : null;
    if (m && m.laborId && db.labores.some(l => l.id === m.laborId)) { formLabor(m.laborId); return; }
    const tipo = pre.tipo || 'gasto';
    const venta = v => v.categoria === 'Venta de grano' || v.categoria === 'Venta de paja';
    formulario({
      titulo: m ? 'Editar apunte' : (tipo === 'ingreso' ? 'Nuevo ingreso' : 'Nuevo gasto'),
      valores: m || { tipo, fecha: hoyISO(), categoria: CAT[tipo][0], fincaId: pre.finca || '', campana: db.campanaActiva },
      campos: [
        { k: 'tipo', label: 'Tipo', type: 'select', req: 1, half: 1, opts: () => [['gasto', 'Gasto'], ['ingreso', 'Ingreso']] },
        { k: 'fecha', label: 'Fecha', type: 'date', req: 1, half: 1 },
        { k: 'categoria', label: 'Categoría', type: 'select', req: 1, dyn: 1, opts: v => CAT[v.tipo || 'gasto'].map(x => [x, x]) },
        { k: 'concepto', label: 'Concepto', placeholder: 'Ej.: 2.500 kg semilla certificada' },
        { k: 'kilos', label: 'Kilos vendidos', type: 'num', miles: 1, half: 1, show: venta },
        { k: 'precio', label: 'Precio (€/kg)', type: 'num', half: 1, show: venta, placeholder: 'Ej.: 0,23' },
        { k: 'importe', label: 'Importe (€)', type: 'num', miles: 1, req: 1 },
        { k: 'fincaId', label: 'Finca', type: 'select', opts: () => [['', 'General (sin finca)'], ...optsFincas()], hint: 'Déjalo en «General» para gastos de toda la explotación: seguro, gasóleo, PAC…' },
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
        if (m) Object.assign(m, v); else db.movimientos.push(Object.assign({ id: uid() }, v));
      },
      borrar: m && (() => { db.movimientos = db.movimientos.filter(x => x !== m); })
    });
  }

  function formCosecha(id, fincaPre) {
    const x = id ? db.cosechas.find(c => c.id === id) : null;
    if (!db.fincas.length) { toast('Primero añade una finca'); return; }
    formulario({
      titulo: x ? 'Editar cosecha' : 'Apuntar cosecha',
      valores: x || { fecha: hoyISO(), fincaId: fincaPre || optsFincas()[0][0], campana: db.campanaActiva },
      campos: [
        { k: 'fincaId', label: 'Finca', type: 'select', req: 1, opts: optsFincas },
        { k: 'fecha', label: 'Fecha', type: 'date', req: 1, half: 1 },
        { k: 'kg', label: 'Kilos', type: 'num', miles: 1, req: 1, half: 1, placeholder: 'Ej.: 12.400' },
        { k: 'humedad', label: 'Humedad (%)', type: 'num', placeholder: 'Ej.: 11,5' },
        campoCampana,
        { k: 'notas', label: 'Destino o notas', type: 'textarea', placeholder: 'Cooperativa, almacén, nº de ticket de báscula…' }
      ],
      guardar(v) {
        if (v.kg <= 0) return 'Los kilos tienen que ser mayores que cero';
        if (x) Object.assign(x, v); else db.cosechas.push(Object.assign({ id: uid() }, v));
      },
      borrar: x && (() => { db.cosechas = db.cosechas.filter(c => c !== x); })
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
          db = nuevoDB(); guardar(false); ui.fincaId = null; volverA(0); render(); toast('Datos borrados');
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
    ['superficie', ['superficie', 'superficie ha', 'sup', 'sup ha', 'hectareas', 'ha', 'has', 'superficie declarada', 'superficie sigpac', 'superficie cultivada', 'sup declarada']],
    ['cultivo', ['cultivo', 'producto', 'cultivo declarado', 'siembra', 'especie', 'cultivo campana']],
    ['variedad', ['variedad']],
    ['campana', ['campana', 'ano', 'cosecha', 'ano campana']],
    ['notas', ['notas', 'observaciones', 'obs', 'comentarios', 'nota']],
    ['nombre', ['nombre', 'finca', 'paraje', 'denominacion', 'nombre finca', 'nombre de la finca', 'alias']]
  ];
  const ETIQ = { nombre: 'Nombre', superficie: 'Superficie', municipio: 'Municipio', codMunicipio: 'Cód. municipio', provincia: 'Provincia', agregado: 'Agregado', zona: 'Zona', poligono: 'Polígono', parcela: 'Parcela', recinto: 'Recinto', otras: 'Otras parcelas', cultivo: 'Cultivo', variedad: 'Variedad', campana: 'Campaña', notas: 'Notas' };
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
    m = s.match(/^(\d{4})$/);
    if (m) { const a = +m[1] - 1; return a + '/' + String((a + 1) % 100).padStart(2, '0'); } // año de cosecha
    return '';
  }
  const cultivoNormal = v => { const n = norm(v); if (!n) return ''; const c = db.cultivos.find(x => norm(x) === n); const t = String(v).trim(); return c || (t.charAt(0).toUpperCase() + t.slice(1).toLowerCase()); };

  function analizarHoja(XLSX, wb) {
    const nombreHoja = wb.SheetNames.find(n => norm(n) !== 'instrucciones') || wb.SheetNames[0];
    const filas = XLSX.utils.sheet_to_json(wb.Sheets[nombreHoja], { header: 1, raw: true, defval: '' });
    let cab = -1, mapa = {};
    for (let i = 0; i < Math.min(15, filas.length); i++) {
      const m = {}; filas[i].forEach((h, j) => { const k = columna(h); if (k && !(k in m)) m[k] = j; });
      if (Object.keys(m).length >= 2) { cab = i; mapa = m; break; }
    }
    if (cab < 0) return { error: 'No encuentro la fila de títulos. La primera fila debe tener títulos como Nombre, Superficie, Polígono, Parcela…' };
    const reconocidas = Object.keys(mapa).map(k => ETIQ[k]);
    const ignoradas = filas[cab].filter((h, j) => String(h).trim() && !Object.values(mapa).includes(j)).map(String);
    const val = (fila, k) => (k in mapa ? fila[mapa[k]] : '');
    const txt = (fila, k) => String(val(fila, k) == null ? '' : val(fila, k)).trim();
    const ent = (fila, k) => { const v = val(fila, k); if (v === '' || v == null) return null; const n = typeof v === 'number' ? v : parseNum(v, false); return Number.isFinite(n) ? Math.round(n) : null; };
    const grupos = new Map(); let vacias = 0;
    for (const fila of filas.slice(cab + 1)) {
      if (!fila.some(v => String(v).trim())) continue;
      let pr = provinciaDe(val(fila, 'provincia'));
      let mu = ent(fila, 'codMunicipio'); let muNombre = txt(fila, 'municipio');
      if (mu == null && /^\d+$/.test(muNombre)) { mu = +muNombre; muNombre = ''; }
      const ag = ent(fila, 'agregado') || 0, zo = ent(fila, 'zona') || 0, po = ent(fila, 'poligono'), pa = ent(fila, 'parcela');
      const recs = listaRecintos(val(fila, 'recinto'));
      const nombre = txt(fila, 'nombre');
      const supV = val(fila, 'superficie'); const sup = typeof supV === 'number' ? supV : parseNum(supV, false);
      if (!nombre && !(po && pa)) { vacias++; continue; }
      const clave = nombre ? 'n:' + norm(nombre) : `r:${pr}:${mu}:${ag}:${zo}:${po}:${pa}`;
      let g = grupos.get(clave);
      if (!g) { g = { nombreAuto: !nombre, nombre: nombre || `Pol. ${po} Parc. ${pa}`, municipio: muNombre, provincia: pr || '', codMunicipio: mu || '', agregado: ag, zona: zo, poligono: po ? String(po) : '', parcela: pa ? String(pa) : '', recs: [], entera: false, otras: [], superficie: 0, cultivos: {}, variedad: txt(fila, 'variedad'), campana: campanaTxt(val(fila, 'campana')), notas: [], filas: 0 }; grupos.set(clave, g); }
      g.filas++;
      if (!g.municipio && muNombre) g.municipio = muNombre;
      if (Number.isFinite(sup) && sup > 0) g.superficie += sup;
      if (po && pa) {
        if (!g.poligono) Object.assign(g, { provincia: pr || g.provincia, codMunicipio: mu || g.codMunicipio, agregado: ag, zona: zo, poligono: String(po), parcela: String(pa) });
        if (+g.poligono === po && +g.parcela === pa && (+g.codMunicipio || 0) === (mu || +g.codMunicipio || 0)) { if (recs.length) g.recs.push(...recs); else g.entera = true; }
        else if (recs.length) recs.forEach(re => g.otras.push(refTxt({ pr: pr || g.provincia, mu: mu || g.codMunicipio, ag, zo, po, pa, re })));
        else g.otras.push(refTxt({ pr: pr || g.provincia, mu: mu || g.codMunicipio, ag, zo, po, pa, re: null }));
      }
      txt(fila, 'otras').split(/[\n;]+/).map(s => s.trim()).filter(Boolean).forEach(s => g.otras.push(s));
      const cu = cultivoNormal(txt(fila, 'cultivo'));
      if (cu) g.cultivos[cu] = (g.cultivos[cu] || 0) + (Number.isFinite(sup) && sup > 0 ? sup : 1);
      if (!g.variedad) g.variedad = txt(fila, 'variedad');
      if (!g.campana) g.campana = campanaTxt(val(fila, 'campana'));
      const nota = txt(fila, 'notas'); if (nota && !g.notas.includes(nota)) g.notas.push(nota);
    }
    const plan = [...grupos.values()].map(g => {
      const cultivo = Object.entries(g.cultivos).sort((a, b) => b[1] - a[1]).map(x => x[0])[0] || '';
      const datos = { nombre: g.nombre, municipio: g.municipio, provincia: g.provincia, codMunicipio: g.codMunicipio, agregado: g.agregado, zona: g.zona, poligono: g.poligono, parcela: g.parcela,
        recinto: g.entera ? '' : [...new Set(g.recs)].sort((a, b) => a - b).join(', '), otras: [...new Set(g.otras)].join('\n'), superficie: Math.round(g.superficie * 10000) / 10000 || null, notas: g.notas.join('\n') };
      const existe = db.fincas.find(f => norm(f.nombre) === norm(g.nombre)) ||
        (g.poligono ? db.fincas.find(f => +f.provincia === +g.provincia && +f.codMunicipio === +g.codMunicipio && +f.poligono === +g.poligono && +f.parcela === +g.parcela) : null);
      if (existe && g.nombreAuto) datos.nombre = existe.nombre; // no cambiar el nombre si el Excel no trae
      return { datos, cultivo, variedad: g.variedad, campana: g.campana || db.campanaActiva, existe };
    });
    return { hoja: nombreHoja, plan, reconocidas, ignoradas, vacias, filas: filas.length - cab - 1 };
  }

  async function importarExcel(file) {
    let XLSX;
    try { XLSX = await cargarXLSX(); } catch (e) { toast('No se ha podido cargar el lector de Excel. Abre la app una vez con cobertura.'); return; }
    let r;
    try { const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' }); r = analizarHoja(XLSX, wb); }
    catch (e) { toast('No se ha podido leer ese archivo. Usa un Excel (.xlsx, .xls), .ods o .csv'); return; }
    if (r.error) { toast(r.error); return; }
    if (!r.plan.length) { toast('No hay fincas en la hoja: cada fila necesita un nombre o polígono y parcela'); return; }
    const nuevas = r.plan.filter(p => !p.existe).length, act = r.plan.length - nuevas;
    const sinSup = r.plan.filter(p => !p.datos.superficie && !(p.existe && p.existe.superficie > 0)).length;
    const conRef = r.plan.filter(p => p.datos.poligono && p.datos.provincia && p.datos.codMunicipio).length;
    const filasHTML = r.plan.slice(0, 40).map(p => `<tr><td>${esc(p.datos.nombre)}${p.existe ? ' <span class="tag">actualiza</span>' : ''}</td><td>${p.datos.superficie ? fmtN(p.datos.superficie) : '—'}</td><td>${p.datos.poligono ? `${esc(p.datos.poligono)}-${esc(p.datos.parcela)}${p.datos.recinto ? '-' + esc(p.datos.recinto) : ''}${p.datos.otras ? ' +' + p.datos.otras.split('\n').length : ''}` : '—'}</td><td style="text-align:left">${p.cultivo ? chip(p.cultivo) : ''}</td></tr>`).join('');
    abrirCapa(`<div class="sheet" role="dialog" aria-modal="true">${cabecera('Importar fincas')}<div class="body">
      <p style="margin:0">Hoja «${esc(r.hoja)}»: ${r.filas} filas → <b>${r.plan.length} fincas</b> (${nuevas} nuevas${act ? `, ${act} ya existen y se actualizarán` : ''}).</p>
      <div class="small"><b>Columnas reconocidas:</b> ${r.reconocidas.map(esc).join(', ')}${r.ignoradas.length ? `<br><span class="muted">No usadas: ${r.ignoradas.map(esc).join(', ')}</span>` : ''}</div>
      ${sinSup ? `<p class="small neg" style="margin:0">${sinSup} sin superficie: se tomará la de SIGPAC si tienen referencia.</p>` : ''}
      ${r.plan.some(p => p.cultivo) ? `<p class="small muted" style="margin:0">Los cultivos se asignan a la campaña indicada en el Excel o, si no hay, a la ${db.campanaActiva}.</p>` : ''}
      <div class="tablewrap"><table><thead><tr><th>Finca</th><th>ha</th><th>Pol-Parc-Rec</th><th>Cultivo</th></tr></thead><tbody>${filasHTML}</tbody></table></div>
      ${r.plan.length > 40 ? `<p class="note">…y ${r.plan.length - 40} más.</p>` : ''}
      ${conRef ? `<p class="small muted" style="margin:0">${conRef} con referencia SIGPAC completa: se situarán en el mapa automáticamente.</p>` : ''}
      </div><footer><button class="btn" data-close>Cancelar</button><button class="btn primary" data-ok>Importar ${r.plan.length}</button></footer></div>`, (el, d) => {
      el.querySelector('[data-ok]').addEventListener('click', () => {
        for (const p of r.plan) {
          const f = p.existe || { id: uid() };
          for (const [k, v] of Object.entries(p.datos)) if (v !== '' && v != null) f[k] = v;
          if (!p.existe) db.fincas.push(f);
          if (p.cultivo) {
            if (!db.cultivos.includes(p.cultivo)) db.cultivos.push(p.cultivo);
            db.siembras = db.siembras.filter(s => !(s.fincaId === f.id && s.campana === p.campana));
            db.siembras.push({ id: uid(), fincaId: f.id, campana: p.campana, cultivo: p.cultivo, variedad: p.variedad || '' });
            if (!db.campanas.includes(p.campana)) db.campanas.push(p.campana);
          }
        }
        db.ejemplo = false;
        guardar(); volverA(0); ui.tab = 'fincas'; render();
        toast(`${r.plan.length} fincas importadas`);
        setTimeout(() => actualizarContornos({}), 400);
      });
    });
  }

  const COLS_EXPORT = ['Nombre', 'Superficie (ha)', 'Municipio', 'Provincia', 'Cód. municipio', 'Agregado', 'Zona', 'Polígono', 'Parcela', 'Recinto', 'Otras parcelas', 'Cultivo', 'Variedad', 'Campaña', 'Notas'];
  async function excelFincas(plantilla) {
    let XLSX;
    try { XLSX = await cargarXLSX(); } catch (e) { toast('No se ha podido cargar el generador de Excel. Abre la app una vez con cobertura.'); return; }
    const c = db.campanaActiva;
    const filas = plantilla
      ? [['El Llano', 12.45, 'Villalón de Campos', 47, 215, 0, 0, 403, 30, '1', '', 'Trigo', 'Marcopolo', c, 'Arrendada'],
         ['El Llano', 3.2, 'Villalón de Campos', 47, 215, 0, 0, 403, 31, '1, 2', '', 'Trigo', 'Marcopolo', c, ''],
         ['Carrascal', 8.2, 'Villalón de Campos', 47, 215, 0, 0, 404, 39, '', '', 'Lenteja', 'Castellana', c, '']]
      : db.fincas.slice().sort(byNombre).map(f => { const s = siembraDe(f.id, c); return [f.nombre, f.superficie, f.municipio || '', f.provincia || '', f.codMunicipio || '', f.agregado || 0, f.zona || 0, f.poligono || '', f.parcela || '', f.recinto || '', f.otras || '', s ? s.cultivo : '', s ? s.variedad || '' : '', c, f.notas || '']; });
    const ws = XLSX.utils.aoa_to_sheet([COLS_EXPORT, ...filas]);
    ws['!cols'] = COLS_EXPORT.map((h, i) => ({ wch: [18, 14, 20, 10, 14, 9, 6, 9, 9, 10, 18, 12, 14, 10, 24][i] }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Fincas');
    if (plantilla) {
      const ins = XLSX.utils.aoa_to_sheet([['Cómo rellenar la hoja «Fincas»'], [''],
        ['• Una fila por finca, o una fila por recinto: las filas con el mismo nombre se juntan en una sola finca y se suman sus superficies.'],
        ['• Provincia, Cód. municipio, Polígono, Parcela y Recinto son los códigos SIGPAC. Están en tu declaración de la PAC o en el visor SIGPAC.'],
        ['• Si dejas Recinto vacío se toma la parcela entera. Puedes poner varios: 1, 2, 5-7.'],
        ['• Agregado y Zona casi siempre son 0.'],
        ['• Campaña: 2026/27, o el año de cosecha (2027). Si se deja vacía se usa la campaña activa.'],
        ['• Puedes borrar las filas de ejemplo. Los títulos de la primera fila pueden variar un poco (Ha, Pol., Paraje…).']]);
      ins['!cols'] = [{ wch: 110 }];
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
    const json = JSON.stringify({ app: 'cuaderno-campo', version: VERSION, exportado: new Date().toISOString(), datos: db }, null, 1);
    const ok = await guardarArchivo(new Blob([json], { type: 'application/json' }), `cuaderno-campo-${fecha}.json`, compartir, `Copia de seguridad del ${fmtFecha(fecha)}`);
    if (ok) { db.ultimaCopia = new Date().toISOString(); guardar(false); render(); toast('Copia hecha'); }
  }

  $('#fileIn').addEventListener('change', async e => {
    const file = e.target.files[0]; e.target.value = '';
    if (!file) return;
    let datos;
    try { const o = JSON.parse(await file.text()); datos = o && o.datos ? o.datos : o; if (!datos || !Array.isArray(datos.fincas)) throw new Error(); }
    catch (err) { toast('Ese archivo no es una copia de Cuaderno de Campo'); return; }
    const r = await confirmar('¿Importar esta copia?',
      `La copia tiene <b>${datos.fincas.length} fincas</b>, ${(datos.labores || []).length} labores, ${(datos.movimientos || []).length} apuntes y ${(datos.cosechas || []).length} cosechas.<br><br>Sustituirá a los datos que hay ahora en este teléfono (${db.fincas.length} fincas).`, 'Importar');
    if (!r.ok) return;
    db = normalizar(datos); db.ultimaCopia = new Date().toISOString(); db.ejemplo = false; guardar(false);
    ui.fincaId = null; volverA(0); render(); toast('Copia importada');
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

  // ---------- Eventos ----------
  const acciones = {
    verFinca: b => { ui.fincaId = b.dataset.id; empujar(); render(); window.scrollTo(0, 0); },
    atras: () => volverA(profundidad() - 1),
    nuevaFinca: () => formFinca(),
    editFinca: b => formFinca(b.dataset.id),
    asignar: b => formAsignar(b.dataset.id),
    nuevaLabor: b => formLabor(null, b.dataset.finca),
    editLabor: b => formLabor(b.dataset.id),
    nuevoMov: b => formMov(null, { finca: b.dataset.finca, tipo: b.dataset.tipo }),
    editMov: b => formMov(b.dataset.id),
    nuevaCosecha: b => formCosecha(null, b.dataset.finca),
    editCosecha: b => formCosecha(b.dataset.id),
    filtroMov: b => { ui.filtroMov = b.dataset.k; render(); },
    vistaFincas: b => { ui.vistaFincas = b.dataset.k; try { localStorage.setItem(KEY + '.vista', ui.vistaFincas); } catch (e) { /* ignorar */ } render(); },
    verMapa: b => { if (ui.fincaId) volverA(0); ui.fincaId = null; ui.tab = 'fincas'; ui.vistaFincas = 'mapa'; ui.enfocar = b.dataset.id; render(); window.scrollTo(0, 0); },
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
      db = nuevoDB(); guardar(false); ui.fincaId = null; volverA(0); render();
    }
  };
  function irTab(t) {
    if (ui.fincaId) volverA(0);
    ui.fincaId = null; ui.tab = t; render(); window.scrollTo(0, 0);
  }
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-act]');
    if (b && acciones[b.dataset.act]) acciones[b.dataset.act](b);
  });
  document.querySelectorAll('.tabs button').forEach(b => b.addEventListener('click', () => irTab(b.dataset.tab)));
  $('#main').addEventListener('change', e => {
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
