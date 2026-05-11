"""
activar_live.py
===============
Parchea tablero-macro.html para que el botón "Actualizar" descargue en vivo
TODOS los datos de IDECBA:

  Antes  → sólo calendario + IPCBA rubros + PGB sectorial
  Después → también: IPCBA serie general · Canastas · Empleo (ETOI)
                     Ejes comerciales · Líneas de pobreza e indigencia

Cómo funciona:
  1. Lee el HTML existente.
  2. Localiza la línea del addEventListener original de calRefresh.
  3. Inyecta el bloque JS con todos los nuevos fetchers y re-renderers.
  4. El nuevo handler reemplaza al viejo usando cloneNode() para limpiar listeners.
  5. Escribe el HTML modificado (backup automático en tablero-macro.bak.html).

Uso:
    python activar_live.py
    python activar_live.py --out tablero-macro-live.html   (no sobreescribe)
    python activar_live.py --no-backup
"""

import re
import shutil
import argparse
from pathlib import Path

SCRIPT_DIR = Path(__file__).parent
HTML_IN    = SCRIPT_DIR / "tablero-macro.html"

# ── Bloque JS a inyectar ──────────────────────────────────────────────────────
# Se inserta justo después de la línea original del addEventListener de calRefresh.

JS_PATCH = r"""
// ═══════════════════════════════════════════════════════════════════════════════
// LIVE REFRESH · Todos los indicadores IDECBA · activar_live.py
// ═══════════════════════════════════════════════════════════════════════════════

// ── Helpers ────────────────────────────────────────────────────────────────────

/** Destruye la instancia Chart.js existente en un canvas y crea una nueva. */
function _rek(id, cfg) {
  const el = document.getElementById(id); if (!el) return null;
  const c = Chart.getChart(el); if (c) c.destroy();
  return new Chart(el, cfg);
}

/** Descarga un XLSX vía lista de proxies CORS; devuelve un workbook SheetJS. */
async function _xlsxGet(url) {
  const PS = [
    'https://api.cors.lol/?url=',
    'https://corsproxy.io/?url=',
    'https://api.allorigins.win/raw?url=',
  ];
  let lastErr;
  for (const p of PS) {
    try {
      const ctl = new AbortController();
      const to  = setTimeout(() => ctl.abort(), 20000);
      const r   = await fetch(p + encodeURIComponent(url), {cache:'no-store', signal:ctl.signal});
      clearTimeout(to);
      if (!r.ok) throw new Error('HTTP ' + r.status);
      const buf = await r.arrayBuffer();
      return XLSX.read(new Uint8Array(buf), {type:'array'});
    } catch(e) { lastErr = e; console.warn('[_xlsxGet]', p, e.message); }
  }
  throw lastErr || new Error('Sin descarga: ' + url);
}

/** Busca el URL más reciente de un XLSX en el banco de datos IDECBA (wp-json). */
async function _idecbaUrl(searchTerm, urlPattern) {
  const PS  = ['https://corsproxy.io/?url=', 'https://api.cors.lol/?url='];
  const API = 'https://www.estadisticaciudad.gob.ar/eyc/wp-json/wp/v2/banco_datos'
            + '?search=' + encodeURIComponent(searchTerm) + '&per_page=5';
  for (const p of PS) {
    try {
      const posts = await (await fetch(p + encodeURIComponent(API), {cache:'no-store'})).json();
      for (const post of posts) {
        try {
          const html = await (await fetch(p + encodeURIComponent(post.link), {cache:'no-store'})).text();
          const m = html.match(urlPattern); if (m) return m[1];
        } catch(_) {}
      }
    } catch(_) {}
  }
  return null;
}

/** Parsea una celda de fecha SheetJS (número serial o string) → {y, m}. */
function _parseDate(c) {
  if (typeof c === 'number' && c > 40000) {
    const d = XLSX.SSF.parse_date_code(c); return d ? {y:d.y, m:d.m} : null;
  }
  if (typeof c === 'string') {
    let m = c.trim().match(/^(\d{4})[\/\-](\d{1,2})$/);
    if (m) return {y:+m[1], m:+m[2]};
    m = c.trim().match(/^(\d{1,2})[\/\-](\d{4})$/);
    if (m) return {y:+m[2], m:+m[1]};
  }
  return null;
}

const _MES  = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const _f1   = v => v != null ? +parseFloat(v).toFixed(1) : null;
const _f2   = v => v != null ? +parseFloat(v).toFixed(2) : null;
const _ri   = v => v != null ? Math.round(parseFloat(v)) : null;
const _sl   = (a, n) => (a||[]).slice(Math.max(0, (a||[]).length - n));

function _cacheGet(k, ttl) {
  try { const {ts,d} = JSON.parse(localStorage.getItem(k) || '{}'); if (Date.now()-ts < ttl) return d; } catch(_) {}
  return null;
}
function _cacheSet(k, d) { try { localStorage.setItem(k, JSON.stringify({ts:Date.now(), d})); } catch(_) {} }

// ── Fetch: IPCBA serie general ─────────────────────────────────────────────────

async function _fetchIpcbaGral() {
  const CK = 'caba_ipcba_gral_v2', TTL = 6 * 3600 * 1000;
  const cached = _cacheGet(CK, TTL); if (cached) return cached;

  let url = await _idecbaUrl(
    'IPCBA índice nivel general serie mensual empalme base',
    /href="(https:\/\/www\.estadisticaciudad\.gob\.ar\/eyc\/wp-content\/uploads\/[^"]+IPCBA[^"]*(?:empalme|nivel_general|general_mensual|serie_mensual)[^"]*\.xlsx)"/i
  );
  url = url || 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/IPCBA_base_2021100-Indices_nivel_general.xlsx';

  const wb   = await _xlsxGet(url);
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});

  // Detecta fila de encabezado
  let hi = 0;
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    if (rows[i].some(c => /per[ií]odo|mes|fecha/i.test(String(c||'')))) { hi = i; break; }
  }

  const meses=[], meses_label=[], indice=[], var_mensual=[], var_ia=[];
  for (let i = hi+1; i < rows.length; i++) {
    const row = rows[i];
    const dt  = _parseDate(row[0]); if (!dt) continue;
    const {y, m} = dt;
    meses.push(`${y}-${String(m).padStart(2,'0')}`);
    meses_label.push(`${_MES[m-1]} ${y}`);
    indice.push(_f2(row[1]));
    var_mensual.push(_f1(row[2]));
    var_ia.push(_f1(row[3]));
  }
  if (!meses.length) throw new Error('IPCBA gral: sin filas parseables');

  const d = { meses, meses_label, indice, var_mensual, var_ia,
               ultimo_mes: meses.at(-1), ultimo_valor: indice.at(-1) };
  _cacheSet(CK, d);
  return d;
}

// ── Fetch: Canastas y líneas de pobreza ────────────────────────────────────────

async function _fetchCanastas() {
  const CK = 'caba_canastas_v2', TTL = 6 * 3600 * 1000;
  const cached = _cacheGet(CK, TTL); if (cached) return cached;

  let url = await _idecbaUrl(
    'canastas consumo líneas pobreza indigencia ciudad',
    /href="(https:\/\/www\.estadisticaciudad\.gob\.ar\/eyc\/wp-content\/uploads\/[^"]+[Cc]anasta[^"]*\.xlsx)"/i
  );
  url = url || 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/Canastas_y_Lineas_de_Pobreza.xlsx';

  const wb = await _xlsxGet(url);

  // Buscar la hoja con CA / CBT
  let ws = wb.Sheets[wb.SheetNames[0]];
  for (const n of wb.SheetNames) {
    if (/canasta|CA|CBT|alimentaria/i.test(n)) { ws = wb.Sheets[n]; break; }
  }
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});

  let hi = 0, colCA = 1, colCBT = 2;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    if (rows[i].some(c => /per[ií]odo|mes|fecha/i.test(String(c||'')))) {
      hi = i;
      for (let j = 0; j < rows[i].length; j++) {
        const s = String(rows[i][j]||'').toLowerCase();
        if (s.includes('alimentaria') || s === 'ca') colCA = j;
        else if (s.includes('total') || s.includes('básica') || s.includes('basica') || s === 'cbt') colCBT = j;
      }
      break;
    }
  }

  const meses=[], ca=[], total=[];
  for (let i = hi+1; i < rows.length; i++) {
    const row = rows[i];
    const dt  = _parseDate(row[0]); if (!dt) continue;
    const {y, m} = dt;
    meses.push(`${y}-${String(m).padStart(2,'0')}`);
    ca.push(_ri(row[colCA]));
    total.push(_ri(row[colCBT]));
  }
  if (!meses.length) throw new Error('Canastas: sin filas');
  const d = {meses, ca, total};
  _cacheSet(CK, d);
  return d;
}

// ── Fetch: Empleo ETOI ─────────────────────────────────────────────────────────

async function _fetchEmpleo() {
  const CK = 'caba_empleo_v2', TTL = 24 * 3600 * 1000;
  const cached = _cacheGet(CK, TTL); if (cached) return cached;

  let url = await _idecbaUrl(
    'mercado laboral Ciudad Buenos Aires ETOI tasas actividad empleo desocupación trimestral',
    /href="(https:\/\/www\.estadisticaciudad\.gob\.ar\/eyc\/wp-content\/uploads\/[^"]+(?:ETOI|etoi|laboral|empleo|tasas)[^"]*\.xlsx)"/i
  );
  url = url || 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/ETOI_series_historicas.xlsx';

  const wb   = await _xlsxGet(url);
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});

  let hi = 0, colAct = 1, colEmp = 2, colDesoc = 3;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    if (rows[i].some(c => /trimestre|per[ií]odo/i.test(String(c||'')))) {
      hi = i;
      for (let j = 0; j < rows[i].length; j++) {
        const s = String(rows[i][j]||'').toLowerCase();
        if (s.includes('actividad')) colAct = j;
        else if ((s.includes('empleo') || s.includes('ocup')) && j !== colAct) colEmp = j;
        else if (s.includes('desocup')) colDesoc = j;
      }
      break;
    }
  }

  const trimestres=[], actividad=[], empleo=[], desocupacion=[];
  for (let i = hi+1; i < rows.length; i++) {
    const row = rows[i]; if (!row[0]) continue;
    const s   = String(row[0]).trim();
    if (!/\d{4}/.test(s) || !/[Tt][1-4]/.test(s)) continue;
    trimestres.push(s.replace(/\s+/, '-'));
    actividad.push(_f1(row[colAct]));
    empleo.push(_f1(row[colEmp]));
    desocupacion.push(_f1(row[colDesoc]));
  }
  if (!trimestres.length) throw new Error('Empleo: sin filas');
  const d = {trimestres, actividad, empleo, desocupacion};
  _cacheSet(CK, d);
  return d;
}

// ── Fetch: Ejes comerciales / locales ──────────────────────────────────────────

async function _fetchLocales() {
  const CK = 'caba_locales_v2', TTL = 72 * 3600 * 1000;
  const cached = _cacheGet(CK, TTL); if (cached) return cached;

  let url = await _idecbaUrl(
    'ejes comerciales Ciudad Buenos Aires locales ocupación vacancia cuatrimestre',
    /href="(https:\/\/www\.estadisticaciudad\.gob\.ar\/eyc\/wp-content\/uploads\/[^"]+(?:ejes|locales|vacancia)[^"]*\.xlsx)"/i
  );
  url = url || 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/Ejes_comerciales_series.xlsx';

  const wb   = await _xlsxGet(url);
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});

  // Primera fila con contenido = encabezados de ejes
  let hi = 0;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    if (rows[i].length > 2 && rows[i][0] != null) { hi = i; break; }
  }

  const ejeNames = rows[hi].slice(1).map(c => String(c||'').trim()).filter(Boolean);
  const ejesArr  = {}; ejeNames.forEach(e => ejesArr[e] = []);
  const periodos = [];

  for (let i = hi+1; i < rows.length; i++) {
    const row = rows[i]; if (!row[0]) continue;
    periodos.push(String(row[0]).trim());
    ejeNames.forEach((e, j) => ejesArr[e].push(row[j+1] != null ? _f1(row[j+1]) : null));
  }
  if (!periodos.length) throw new Error('Locales: sin filas');
  const d = {periodos, ejes: ejesArr};
  _cacheSet(CK, d);
  return d;
}

// ── Fetch: Líneas de pobreza ───────────────────────────────────────────────────

async function _fetchPobreza() {
  const CK = 'caba_pobreza_v2', TTL = 6 * 3600 * 1000;
  const cached = _cacheGet(CK, TTL); if (cached) return cached;

  let url = await _idecbaUrl(
    'líneas pobreza indigencia hogares tipo Ciudad Buenos Aires',
    /href="(https:\/\/www\.estadisticaciudad\.gob\.ar\/eyc\/wp-content\/uploads\/[^"]+(?:[Ll]inea|[Pp]obreza|[Cc]anasta)[^"]*\.xlsx)"/i
  );
  url = url || 'https://www.estadisticaciudad.gob.ar/eyc/wp-content/uploads/2024/12/Canastas_y_Lineas_de_Pobreza.xlsx';

  const wb = await _xlsxGet(url);

  // Buscar hoja con LP/LI (suele ser la segunda hoja)
  let ws = null;
  for (const n of wb.SheetNames) {
    if (/linea|lp|li|pobreza|indigencia|hogar/i.test(n)) { ws = wb.Sheets[n]; break; }
  }
  ws = ws || wb.Sheets[wb.SheetNames[Math.min(1, wb.SheetNames.length - 1)]];

  const rows = XLSX.utils.sheet_to_json(ws, {header:1, defval:null});
  let hi = 0;
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    if (rows[i].some(c => /hogar|lp|li|linea|per[ií]odo/i.test(String(c||'')))) { hi = i; break; }
  }

  const hogares_labels = ['Hogar 1a', 'Hogar 2', 'Hogar 3', 'Hogar 4', 'Hogar 5'];
  const periodos=[], lp_hogares=[], li_hogares=[];
  const nH = 5;

  for (let i = hi+1; i < rows.length; i++) {
    const row = rows[i];
    const dt  = _parseDate(row[0]); if (!dt) continue;
    const {y, m} = dt;
    periodos.push(`${y}-${String(m).padStart(2,'0')}`);
    // Columnas 1..nH → LP; nH+1..2*nH → LI
    lp_hogares.push(Array.from({length:nH}, (_, j) => _ri(row[j+1])));
    li_hogares.push(Array.from({length:nH}, (_, j) => _ri(row[j+nH+1])));
  }
  if (!periodos.length) throw new Error('Pobreza: sin filas');
  const d = {periodos, lp_hogares, li_hogares, hogares_labels};
  _cacheSet(CK, d);
  return d;
}

// ── Re-renderers ───────────────────────────────────────────────────────────────

window.reRenderIpcba = function() {
  const ip = MACRO.ipcba; if (!ip || !ip.meses) return;
  const N   = 24;
  const lbl = _sl(ip.meses_label || ip.meses.map(ms => {
    const [y, mm] = ms.split('-'); return _MES[+mm-1] + ' ' + y;
  }), N);
  const SKY  = 'rgba(106,173,228,.85)', NAVY = '#0C2340';
  _rek('chIpcM', {
    type:'bar',
    data:{labels:lbl, datasets:[{label:'Var. mensual %', data:_sl(ip.var_mensual,N), backgroundColor:SKY}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>c.parsed.y.toFixed(1).replace('.',',')+'%'}}},
      scales:{y:{ticks:{callback:v=>v+'%'}},x:{ticks:{font:{size:9.5},maxRotation:45,minRotation:45}}}},
  });
  _rek('chIpcIA', {
    type:'line',
    data:{labels:lbl, datasets:[{label:'i.a. %', data:_sl(ip.var_ia,N),
      borderColor:NAVY, backgroundColor:'rgba(12,35,64,.08)', fill:true, tension:.3, pointRadius:2}]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>c.parsed.y==null?'—':c.parsed.y.toFixed(1).replace('.',',')+'%'}}},
      scales:{y:{ticks:{callback:v=>v+'%'}},x:{ticks:{font:{size:9.5},maxRotation:45,minRotation:45}}}},
  });
  const li = ip.meses.length - 1;
  const elMeta = document.getElementById('metaLast');
  if (elMeta) elMeta.textContent = 'Último dato IPCBA · ' + (ip.meses_label?.[li] || ip.meses[li] || '');
};

window.reRenderCanastas = function() {
  const ca = MACRO.canastas; if (!ca || !ca.meses) return;
  const N   = 36;
  const lbl = _sl(ca.meses, N).map(ms => {
    const [y, mm] = ms.split('-'); return _MES[+mm-1].slice(0,3) + '-' + y.slice(2);
  });
  const caV = _sl(ca.ca, N), ctV = _sl(ca.total, N);
  const fmt$ = v => v == null ? '—' : '$' + (v >= 1e6 ? (v/1e6).toFixed(1)+'M' : v >= 1e3 ? (v/1e3).toFixed(0)+'k' : v);
  const tagCA = document.getElementById('tagCA'); if (tagCA) tagCA.textContent = fmt$(caV.at(-1));
  const tagCT = document.getElementById('tagCT'); if (tagCT) tagCT.textContent = fmt$(ctV.at(-1));
  const scaleY = {y:{ticks:{callback:v=>'$'+(v>=1e6?(v/1e6).toFixed(1)+'M':v>=1e3?(v/1e3).toFixed(0)+'k':v)}},
                  x:{ticks:{font:{size:9},maxRotation:45,minRotation:45}}};
  _rek('chCA', {type:'line', data:{labels:lbl,datasets:[{label:'CA',data:caV,
    borderColor:'#1B6EC2',backgroundColor:'rgba(27,110,194,.1)',fill:true,tension:.25,pointRadius:2}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},
      tooltip:{callbacks:{label:c=>fmt$(c.parsed.y)}}},scales:scaleY}});
  _rek('chCBT', {type:'line', data:{labels:lbl,datasets:[{label:'Canasta total',data:ctV,
    borderColor:'#C06A00',backgroundColor:'rgba(192,106,0,.1)',fill:true,tension:.25,pointRadius:2}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},
      tooltip:{callbacks:{label:c=>fmt$(c.parsed.y)}}},scales:scaleY}});
};

window.reRenderEmpleo = function() {
  const emp = MACRO.empleo; if (!emp || !emp.trimestres) return;
  const tL  = t => t.replace(/^(\d{4})-?T/, '$1 T');
  _rek('chEmp', {type:'line',
    data:{labels:emp.trimestres.map(tL), datasets:[
      {label:'Actividad',data:emp.actividad,borderColor:'#0C2340',
       backgroundColor:'rgba(12,35,64,.05)',fill:false,tension:.25,pointRadius:2},
      {label:'Empleo',data:emp.empleo,borderColor:'#1B6EC2',
       backgroundColor:'rgba(27,110,194,.08)',fill:false,tension:.25,pointRadius:2},
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:c=>c.dataset.label+': '+(c.parsed.y==null?'—':c.parsed.y.toFixed(1).replace('.',',')+'%')}}},
      scales:{y:{ticks:{callback:v=>v+'%'}},x:{ticks:{font:{size:9},autoSkip:true,maxTicksLimit:14,maxRotation:45,minRotation:45}}}},
  });
  const tagEmp = document.getElementById('tagEmp');
  if (tagEmp) tagEmp.textContent = emp.trimestres.at(-1) || '';
};

window.reRenderLocales = function() {
  const loc = MACRO.locales; if (!loc || !loc.periodos) return;
  const ET  = ['Zona Centro','Florida','Microcentro','Recoleta','Palermo Soho','Caballito','Flores','Belgrano'];
  const EC  = ['#0C2340','#1B6EC2','#6AADE4','#0A8A5A','#C06A00','#C0392B','#6B7280','#374151'];
  const ed  = ET.filter(e => loc.ejes[e]);
  _rek('chVac', {type:'line',
    data:{labels:loc.periodos.map(p=>p.replace('-',' ')), datasets:ed.map((e,i) => ({
      label:e, data:loc.ejes[e].map(v=>v==null?null:+(100-v).toFixed(1)),
      borderColor:EC[i%EC.length], backgroundColor:EC[i%EC.length], tension:.25, pointRadius:2, fill:false,
    }))},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'bottom',labels:{boxWidth:10,font:{size:10}}},
        tooltip:{callbacks:{label:c=>c.dataset.label+': '+(c.parsed.y==null?'—':c.parsed.y.toFixed(1)+'%')}}},
      scales:{y:{ticks:{callback:v=>v+'%'}}}},
  });
  const tg  = loc.ejes['Total General'] || [];
  const vl  = tg.at(-1) != null ? +(100 - tg.at(-1)).toFixed(1) : null;
  const tagV = document.getElementById('tagVac');
  if (tagV) tagV.textContent = (loc.periodos.at(-1)||'') + ' · vac. total ' + (vl==null?'—':vl.toFixed(1)+'%');
};

window.reRenderPobreza = function() {
  const p = MACRO.pobreza; if (!p || !p.periodos.length) return;
  const N  = 36, st = Math.max(0, p.periodos.length - N);
  const lbl = p.periodos.slice(st).map(per => {
    const [y, m] = per.split('-'); return _MES[+m-1] + '-' + y.slice(2);
  });
  const lp2 = p.lp_hogares.slice(st).map(a => a&&a[1] ? a[1] : null);
  const li2 = p.li_hogares.slice(st).map(a => a&&a[1] ? a[1] : null);
  const fmtM = v => '$ ' + (v>=1e6?(v/1e6).toFixed(1)+' M':v>=1e3?(v/1e3).toFixed(0)+' k':v);
  _rek('chLP', {type:'line',
    data:{labels:lbl, datasets:[
      {label:'LP',data:lp2,borderColor:'#E84A5F',backgroundColor:'rgba(232,74,95,.1)',tension:.2,pointRadius:1,fill:true},
      {label:'LI',data:li2,borderColor:'#0C2340',backgroundColor:'rgba(12,35,64,.05)',tension:.2,pointRadius:1,fill:true},
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:c=>c.dataset.label+': $ '+(c.parsed.y||0).toLocaleString('es-AR')}}},
      scales:{y:{ticks:{callback:v=>fmtM(v)}},x:{ticks:{font:{size:9},autoSkip:true,maxTicksLimit:12,maxRotation:45,minRotation:45}}}},
  });
  const li  = p.periodos.length - 1;
  const up  = p.periodos.at(-1) || '';
  const tagLP   = document.getElementById('tagLP');   if (tagLP)   tagLP.textContent   = '$ ' + (lp2.at(-1)||0).toLocaleString('es-AR');
  const lpSub   = document.getElementById('lpSub');   if (lpSub)   lpSub.textContent   = `Hogar tipo 2 · pareja con dos hijos · ${p.periodos[st]}–${up}`;
  _rek('chLILP', {type:'bar',
    data:{labels:(p.hogares_labels||[]).slice(0,5), datasets:[
      {label:'LP',data:p.lp_hogares[li],backgroundColor:'#E84A5F',borderRadius:2},
      {label:'LI',data:p.li_hogares[li],backgroundColor:'#0C2340',borderRadius:2},
    ]},
    options:{responsive:true,maintainAspectRatio:false,
      plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:c=>c.dataset.label+': $ '+(c.parsed.y||0).toLocaleString('es-AR')}}},
      scales:{y:{ticks:{callback:v=>'$ '+(v>=1e6?(v/1e6).toFixed(1)+' M':v>=1e3?(v/1e3).toFixed(0)+' k':v)}}}}},
  );
  const tagLILP = document.getElementById('tagLILP'); if (tagLILP) tagLILP.textContent = up;
  const liLpSub = document.getElementById('liLpSub'); if (liLpSub) liLpSub.textContent = `Costo en pesos · ${up} · 5 hogares tipo IDECBA`;
};

// ── Reemplaza el botón Actualizar con handler completo ─────────────────────────
{
  // cloneNode elimina todos los listeners existentes (incluyendo el calRefresh original)
  const oldBtn = document.getElementById('btnUpdCal');
  const newBtn = oldBtn.cloneNode(true);
  oldBtn.replaceWith(newBtn);

  newBtn.addEventListener('click', async function liveRefreshAll() {
    const btn = document.getElementById('btnUpdCal');
    const msg = document.getElementById('updMsg');
    btn.classList.add('loading'); btn.disabled = true;
    msg.className = 'upd-msg';

    const ok = [], fail = [];
    const step = s => { msg.textContent = s; };

    // ① Calendario + IPCBA rubros + PGB (función original calRefresh)
    step('📅 Actualizando calendario IDECBA...');
    try {
      await calRefresh();
      // calRefresh re-habilita el botón al terminar; lo deshabilitamos de nuevo
      btn.classList.add('loading'); btn.disabled = true;
      ok.push('Calendario · IPCBA rubros · PGB');
    } catch(e) { console.warn('[calRefresh]', e); fail.push('Calendario'); }

    // ② IPCBA serie general (índice mensual + variaciones)
    step('📈 Bajando IPCBA · serie mensual general...');
    try {
      localStorage.removeItem('caba_ipcba_gral_v2');
      const d = await _fetchIpcbaGral();
      Object.assign(MACRO.ipcba, d);
      window.MACRO = MACRO;
      window.reRenderIpcba();
      ok.push('IPCBA serie');
    } catch(e) { console.warn('[ipcba-gral]', e); fail.push('IPCBA serie'); }

    // ③ Canastas (CA + CBT)
    step('🛒 Bajando canastas y líneas de consumo...');
    try {
      localStorage.removeItem('caba_canastas_v2');
      const d = await _fetchCanastas();
      MACRO.canastas = d; window.MACRO = MACRO;
      window.reRenderCanastas();
      ok.push('Canastas');
    } catch(e) { console.warn('[canastas]', e); fail.push('Canastas'); }

    // ④ Empleo ETOI (actividad, empleo, desocupación)
    step('👷 Bajando mercado laboral ETOI...');
    try {
      localStorage.removeItem('caba_empleo_v2');
      const d = await _fetchEmpleo();
      MACRO.empleo = d; window.MACRO = MACRO;
      window.reRenderEmpleo();
      ok.push('Empleo');
    } catch(e) { console.warn('[empleo]', e); fail.push('Empleo'); }

    // ⑤ Ejes comerciales / locales (vacancia)
    step('🏪 Bajando ejes comerciales...');
    try {
      localStorage.removeItem('caba_locales_v2');
      const d = await _fetchLocales();
      MACRO.locales = d; window.MACRO = MACRO;
      window.reRenderLocales();
      ok.push('Locales');
    } catch(e) { console.warn('[locales]', e); fail.push('Locales'); }

    // ⑥ Líneas de pobreza e indigencia
    step('📉 Bajando líneas de pobreza e indigencia...');
    try {
      localStorage.removeItem('caba_pobreza_v2');
      const d = await _fetchPobreza();
      MACRO.pobreza = d; window.MACRO = MACRO;
      window.reRenderPobreza();
      ok.push('Pobreza');
    } catch(e) { console.warn('[pobreza]', e); fail.push('Pobreza'); }

    // ── Resultado final ──────────────────────────────────────────────────────
    msg.className = fail.length ? 'upd-msg' : 'upd-msg ok';
    const partes = [`✓ ${ok.length} módulos · ` + ok.join(' · ')];
    if (fail.length) partes.push(`⚠ Sin datos: ${fail.join(', ')}`);
    msg.textContent = partes.join(' | ');
    btn.classList.remove('loading'); btn.disabled = false;
  });
}
// ══════════════════════════════════════════════════════════════════════════════
// FIN LIVE REFRESH
// ══════════════════════════════════════════════════════════════════════════════
"""

# ── Función de parcheo ────────────────────────────────────────────────────────

ANCHOR = "document.getElementById('btnUpdCal').addEventListener('click',calRefresh);"

def parchear(html: str) -> str:
    if ANCHOR not in html:
        raise ValueError(
            f"No se encontró el ancla de inyección en el HTML.\n"
            f"Buscando: {ANCHOR!r}\n"
            "Verificá que el tablero-macro.html sea el original sin modificaciones previas."
        )
    # Marcador de que ya fue parcheado
    if "LIVE REFRESH · Todos los indicadores IDECBA" in html:
        print("⚠  El HTML ya tiene el parche aplicado. Omitiendo.")
        return html

    return html.replace(ANCHOR, ANCHOR + "\n" + JS_PATCH, 1)

# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Activa el live refresh completo en tablero-macro.html")
    parser.add_argument("--out", default=str(HTML_IN),
                        help="Archivo de salida (por defecto sobreescribe el original)")
    parser.add_argument("--no-backup", action="store_true",
                        help="No crea backup antes de sobreescribir")
    args = parser.parse_args()

    if not HTML_IN.exists():
        import sys; sys.exit(f"No se encontró: {HTML_IN}")

    html = HTML_IN.read_text(encoding="utf-8")
    print(f"Leído:     {HTML_IN.name}  ({len(html):,} bytes)")

    # Backup automático
    out = Path(args.out)
    if out == HTML_IN and not args.no_backup:
        bak = HTML_IN.with_suffix(".bak.html")
        shutil.copy2(HTML_IN, bak)
        print(f"Backup:    {bak.name}")

    try:
        html_nuevo = parchear(html)
    except ValueError as e:
        import sys; sys.exit(str(e))

    out.write_text(html_nuevo, encoding="utf-8")
    delta = len(html_nuevo) - len(html)
    print(f"Guardado:  {out.name}  ({len(html_nuevo):,} bytes  +{delta:,} bytes del parche)")
    print()
    print("OK. Abri el tablero con un servidor local:")
    print(f"    cd \"{out.parent}\"")
    print( "    python -m http.server 8080")
    print(f"    http://localhost:8080/{out.name}")
    print()
    print("  El boton 'Actualizar' ahora descarga en vivo:")
    print("    [1] Calendario IDECBA")
    print("    [2] IPCBA rubros + PGB sectorial  (ya existia)")
    print("    [3] IPCBA serie mensual general   (nuevo)")
    print("    [4] Canastas y lineas de consumo  (nuevo)")
    print("    [5] Empleo ETOI                   (nuevo)")
    print("    [6] Ejes comerciales / vacancia   (nuevo)")
    print("    [7] Lineas de pobreza e indigencia(nuevo)")

if __name__ == "__main__":
    main()
