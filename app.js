'use strict';
(function () {
  // ---------- Constantes ----------
  const KEY = 'cuadernoCampo.v1';
  const VERSION = '1.0.0';
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

  // ---------- Utilidades ----------
  const $ = (s, el) => (el || document).querySelector(s);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
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
  const ui = { tab: 'fincas', fincaId: null, filtroLabor: { finca: '', tipo: '' }, filtroMov: 'todos' };
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
    if (!db.fincas.length) {
      return h + vacio('Aún no tienes fincas', 'Da de alta tus fincas con su superficie. Después podrás apuntar labores, gastos y cosechas de cada una.',
        '<button class="btn primary" data-act="nuevaFinca">Añadir finca</button><button class="btn" data-act="ejemplo">Ver con datos de ejemplo</button>');
    }
    const c = db.campanaActiva;
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
    let h = `<div class="backhdr"><button class="iconbtn" data-act="atras" aria-label="Volver">${ico.back}</button><h2>${esc(f.nombre)}</h2><button class="iconbtn" data-act="editFinca" data-id="${id}" aria-label="Editar finca">${ico.edit}</button></div>
      <div class="meta"><span><b>${fmtN(f.superficie)} ha</b></span>${f.municipio ? `<span>${esc(f.municipio)}</span>` : ''}${sig ? `<span>SIGPAC ${esc(sig)}</span>` : ''}</div>`;
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

  // ---------- Render principal ----------
  function render() {
    $('#campBtn').textContent = db.campanaActiva + ' ▾';
    $('#title').textContent = ui.fincaId ? 'Finca' : TITULOS[ui.tab];
    document.querySelectorAll('.tabs button').forEach(b => b.setAttribute('aria-current', b.dataset.tab === ui.tab ? 'page' : 'false'));
    const main = $('#main');
    main.innerHTML = ui.fincaId ? vDetalle(ui.fincaId) : ({ fincas: vFincas, cuaderno: vCuaderno, cuentas: vCuentas, cosechas: vCosechas, resumen: vResumen }[ui.tab])();
    const fab = { fincas: ['nuevaFinca', 'Finca'], cuaderno: ['nuevaLabor', 'Labor'], cuentas: ['nuevoMov', 'Apunte'], cosechas: ['nuevaCosecha', 'Cosecha'] }[ui.tab];
    const mostrarFab = fab && !ui.fincaId && (ui.tab === 'fincas' ? db.fincas.length : db.fincas.length);
    $('#fabSlot').innerHTML = mostrarFab ? `<button class="fab" data-act="${fab[0]}">${ico.plus}${fab[1]}</button>` : '';
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
      else if (f.type === 'textarea') inp = `<textarea id="${id}" name="${f.k}" rows="3">${esc(v)}</textarea>`;
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

  function formFinca(id) {
    const f = id ? finca(id) : null;
    formulario({
      titulo: f ? 'Editar finca' : 'Nueva finca',
      valores: f || {},
      campos: [
        { k: 'nombre', label: 'Nombre', req: 1, placeholder: 'Ej.: El Llano' },
        { k: 'superficie', label: 'Superficie (ha)', type: 'num', req: 1, placeholder: 'Ej.: 12,45' },
        { k: 'municipio', label: 'Municipio' },
        { k: 'poligono', label: 'Polígono', half: 1, placeholder: 'SIGPAC' },
        { k: 'parcela', label: 'Parcela', half: 1, placeholder: 'SIGPAC' },
        { k: 'recinto', label: 'Recinto(s)' },
        { k: 'notas', label: 'Notas', type: 'textarea', placeholder: 'Propiedad o arrendada, tipo de suelo, linderos…' }
      ],
      guardar(v) {
        if (v.superficie <= 0) return 'La superficie tiene que ser mayor que cero';
        if (f) Object.assign(f, v); else db.fincas.push(Object.assign({ id: uid() }, v));
      },
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

  // ---------- Copias de seguridad ----------
  async function exportar(compartir) {
    const fecha = hoyISO();
    const json = JSON.stringify({ app: 'cuaderno-campo', version: VERSION, exportado: new Date().toISOString(), datos: db }, null, 1);
    const nombre = `cuaderno-campo-${fecha}.json`;
    let file;
    try { file = new File([json], nombre, { type: 'application/json' }); } catch (e) { file = new Blob([json], { type: 'application/json' }); }
    const marcar = () => { db.ultimaCopia = new Date().toISOString(); guardar(false); render(); toast('Copia hecha'); };
    if (compartir && navigator.canShare && file instanceof File && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: 'Copia Cuaderno de Campo', text: `Copia de seguridad del ${fmtFecha(fecha)}` }); marcar(); return; }
      catch (e) { if (e && e.name === 'AbortError') return; }
    }
    try {
      const url = URL.createObjectURL(file);
      const a = document.createElement('a'); a.href = url; a.download = nombre; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 10000);
      marcar();
      if (compartir) toast('Este navegador no permite compartir; la copia se ha descargado');
    } catch (e) { toast('No se ha podido crear la copia en este navegador'); }
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
    const F = (nombre, superficie, municipio, poligono, parcela) => { const f = { id: uid(), nombre, superficie, municipio, poligono, parcela, recinto: '', notas: '' }; d.fincas.push(f); return f; };
    const fs = [F('El Llano', 12.4, 'Villanueva', '4', '112'), F('Carrascal', 8.2, 'Villanueva', '6', '35'), F('La Hoya', 5.6, 'Villanueva', '4', '140'), F('Valdecasas', 15, 'Torralba', '2', '18'), F('Los Majuelos', 3.1, 'Torralba', '9', '7')];
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
    db = d; guardar(false); ui.tab = 'fincas'; render(); toast('Datos de ejemplo cargados');
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
  render();
})();
