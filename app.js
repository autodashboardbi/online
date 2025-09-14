// app.js — arranged visuals by type + landing page + company theme + improved KPI logic
// State
let rawData = [];           // original normalized array-of-objects
let filteredData = [];      // after slicers applied
let charts = [];            // Chart.js instances to destroy
let theme = { primary: '#2E8B57', accent: '#f0b429' };

// ---------- COMPANY THEME SUPPORT ----------
/*
 Company colours from MLK Community Healthcare logo:
  - orange: #F7941E
  - teal:   #00AFC1
  - green:  #7CC243
*/
let companyPalette = null;

function setCompanyTheme(){
  const primary = '#00AFC1'; // teal as primary
  const accent = '#F7941E';  // orange as accent
  const third = '#7CC243';   // green
  companyPalette = [primary, accent, third];
  setTheme(primary, accent);
  try { updateAllVisuals(); } catch(e){}
}

// ---------- THEME ----------
function setTheme(primary, accent){
  theme.primary = primary;
  theme.accent = accent;
  document.documentElement.style.setProperty('--accent', primary);
  document.documentElement.style.setProperty('--accent-2', accent);
  // re-render visuals to apply colors
  try { updateAllVisuals(); } catch(e) {}
}

// ---------- LANDING ----------
document.getElementById('enterBtn').addEventListener('click', ()=> {
  document.getElementById('landing').style.display = 'none';
  document.getElementById('app').classList.remove('hidden');
});
// allow landing upload to open dashboard and parse file immediately
document.getElementById('fileInputLanding').addEventListener('change', (e)=>{
  const f = e.target.files && e.target.files[0];
  if(!f) return;
  // hide landing and show app
  document.getElementById('landing').style.display = 'none';
  document.getElementById('app').classList.remove('hidden');
  handleFile(f);
});

// ---------- UTIL: Chart lifecycle ----------
function destroyAllCharts(){
  charts.forEach(c => { try { c.destroy(); } catch(e){} });
  charts = [];
}

// ---------- PARSING & NORMALIZATION ----------
function normalizeParsedInput(input){
  if(!input || !input.length) return [];
  let out = [];
  if(Array.isArray(input[0])){ // array-of-arrays - first row is header
    const headers = input[0].map(h => (h||'').toString().trim());
    for(let i=1;i<input.length;i++){
      const row = input[i];
      if(!row || row.every(c=>c===undefined||c==='')) continue;
      const obj = {};
      for(let j=0;j<headers.length;j++) obj[headers[j] || `col${j}`] = row[j];
      out.push(obj);
    }
  } else {
    out = input.map(r=>{
      const o = {};
      Object.keys(r).forEach(k => o[(k||'').toString().trim()] = r[k]);
      return o;
    });
  }
  // convert numeric-ish strings into numbers
  return out.map(row=>{
    const o = {};
    Object.entries(row).forEach(([k,v])=>{
      if(typeof v === 'string'){
        const cleaned = v.replace(/,/g,'').trim();
        if(cleaned !== '' && !isNaN(cleaned)) { o[k]=parseFloat(cleaned); return; }
      }
      o[k]=v;
    });
    return o;
  });
}

// small helper: find key by candidate list (case-insensitive contains)
function findKey(sample, targets){
  if(!sample) return null;
  const keys = Object.keys(sample);
  for(const t of targets){
    const lt = t.toLowerCase();
    const exact = keys.find(k => k.toLowerCase() === lt);
    if(exact) return exact;
    const contains = keys.find(k => k.toLowerCase().includes(lt));
    if(contains) return contains;
  }
  return null;
}

// analyze columns: numeric / categorical / date-like
function analyzeColumns(data){
  if(!data.length) return { numeric:[], categorical:[], dateLike:[] };
  const sample = data[0];
  const keys = Object.keys(sample);
  const numeric = [], cat = [], dateLike = [];
  keys.forEach(k=>{
    const values = data.map(r => r[k]);
    const allNumbers = values.every(v => v === null || v === '' || v === undefined || typeof v === 'number' || (!isNaN(Number(v)) && v !== ''));
    const headerDate = /date|month|day|year|period|timestamp/i.test(k);
    const parseable = values.some(v => {
      if(!v && v !== 0) return false;
      if(typeof v === 'string') { const d = Date.parse(v); return !isNaN(d); }
      return false;
    });
    if(headerDate || parseable) dateLike.push(k);
    else if(allNumbers) numeric.push(k);
    else cat.push(k);
  });
  return { numeric, categorical: cat, dateLike };
}

// ---------- SLICERS ----------
function renderSlicers(data){
  const slicerDiv = document.getElementById('slicers');
  slicerDiv.innerHTML = '';
  const analysis = analyzeColumns(data);
  const categories = analysis.categorical;
  if(!categories.length){ slicerDiv.innerHTML = `<div style="color:var(--muted)">No categorical columns for slicers</div>`; return; }

  categories.forEach(col=>{
    const wrapper = document.createElement('div'); wrapper.className='slicer';
    const label = document.createElement('label'); label.innerText = col;
    const sel = document.createElement('select'); sel.dataset.col = col;
    const vals = Array.from(new Set(data.map(r=>r[col]).filter(v=>v !== undefined && v !== ''))).sort((a,b)=>String(a).localeCompare(String(b)));
    const optAll = document.createElement('option'); optAll.value='__all__'; optAll.innerText = 'All'; sel.appendChild(optAll);
    vals.forEach(v => { const o = document.createElement('option'); o.value = String(v); o.innerText = String(v); sel.appendChild(o); });
    sel.addEventListener('change', applySlicers); // global update
    wrapper.appendChild(label); wrapper.appendChild(sel);
    slicerDiv.appendChild(wrapper);
  });
}

function applySlicers(){
  const slicerDiv = document.getElementById('slicers');
  const selects = slicerDiv.querySelectorAll('select');
  const activeFilters = [];
  selects.forEach(s => {
    const col = s.dataset.col;
    const val = s.value;
    if(val !== '__all__') activeFilters.push({ col, val });
  });
  // filter rawData
  filteredData = rawData.filter(row => {
    return activeFilters.every(f => {
      const cell = row[f.col];
      return String(cell) === f.val;
    });
  });
  if(activeFilters.length === 0) filteredData = rawData.slice();
  updateAllVisuals();
}

// ---------- HELPER: detect identifier-like keys ----------
function isIdentifierKey(keyName){
  if(!keyName) return false;
  return /(^|[^a-z])id(s)?$|id\b|_id\b|\bcode\b|\bstore\b|\bproduct\b|\bsku\b|\bref\b|\bclient\b|\bcustomer\b/i.test(keyName);
}

// ---------- KPIs (improved) ----------
function buildKPIs(data){
  const kpiArea = document.getElementById('kpiArea');
  kpiArea.innerHTML = '';
  if(!data || !data.length) return;
  const columns = analyzeColumns(data);
  const headers = Object.keys(data[0]).map(h=>h.toLowerCase());
  const financeKW = ['sales','revenue','profit','income','price','amount'];
  const healthcareKW = ['admission','admit','birth','lengthofstay','los','visit','emergency','patient'];
  const hasFinance = financeKW.some(k=>headers.some(h=>h.includes(k)));
  const hasHealthcare = healthcareKW.some(k=>headers.some(h=>h.includes(k)));

  const createCard = (title, value, subtitle) => {
    const c = document.createElement('div'); c.className='kpi-card';
    const h = document.createElement('h4'); h.innerText = title;
    const p = document.createElement('p'); p.innerText = value;
    c.appendChild(h); c.appendChild(p);
    if(subtitle){ const s = document.createElement('small'); s.style.color='var(--muted)'; s.innerText=subtitle; c.appendChild(s); }
    kpiArea.appendChild(c);
  };

  // helper to decide whether to show count (unique) or sum for a given column
  function kpiValueForColumn(colKey){
    // if column looks like identifier or categorical -> show counts (unique)
    if(isIdentifierKey(colKey) || columns.categorical.includes(colKey)){
      // show unique count
      return { val: uniqueCount(data, colKey), subtitle: 'unique' };
    }
    // if numeric -> show sum
    if(columns.numeric.includes(colKey)){
      return { val: sumColumn(data, colKey), subtitle: 'sum' };
    }
    // fallback -> sample value or row count
    return { val: data.length, subtitle: 'rows' };
  }

  // domain-specific KPIs
  if(hasFinance){
    const salesK = findKey(data[0], ['sales','amount','units','quantity']);
    const revK = findKey(data[0], ['revenue','income']);
    const profitK = findKey(data[0], ['profit','net profit']);
    const custK = findKey(data[0], ['customer','client']);
    if(salesK){
      const kv = kpiValueForColumn(salesK);
      createCard(`Total ${salesK}`, typeof kv.val === 'number' ? (kv.subtitle==='sum' ? Number(kv.val).toLocaleString() : kv.val) : kv.val, kv.subtitle);
    }
    if(revK){
      const kv = kpiValueForColumn(revK);
      createCard(`Total ${revK}`, kv.subtitle==='sum' ? '$' + Number(kv.val).toLocaleString() : kv.val, kv.subtitle);
    }
    if(profitK){
      const kv = kpiValueForColumn(profitK);
      // prefer average for profit-like fields if numeric
      if(kv.subtitle === 'sum' && columns.numeric.includes(profitK)){
        createCard(`Avg ${profitK}`, '$' + (sumColumn(data, profitK)/data.length).toFixed(2), 'avg');
      } else {
        createCard(`Avg ${profitK}`, kv.val, kv.subtitle);
      }
    }
    if(custK){
      const kv = kpiValueForColumn(custK);
      createCard(`Unique ${custK}`, kv.val, kv.subtitle);
    }
  }

  if(hasHealthcare){
    const admitK = findKey(data[0], ['admission','admit','admissions']);
    const losK = findKey(data[0], ['lengthofstay','los','length']);
    const erK = findKey(data[0], ['emergency','er','er visits','emergency visits']);
    const birthK = findKey(data[0], ['birth','births','deliveries']);
    if(admitK){
      const kv = kpiValueForColumn(admitK);
      createCard(`${admitK}`, kv.val, kv.subtitle);
    }
    if(losK){
      const kv = kpiValueForColumn(losK);
      createCard(`Avg ${losK}`, kv.subtitle === 'sum' ? (sumColumn(data, losK)/data.length).toFixed(2) : kv.val, kv.subtitle);
    }
    if(erK){
      const kv = kpiValueForColumn(erK);
      createCard(`${erK}`, kv.val, kv.subtitle);
    }
    if(birthK){
      const kv = kpiValueForColumn(birthK);
      createCard(`${birthK}`, kv.val, kv.subtitle);
    }
  }

  // fill remaining slots with best numeric / identifier columns
  const existing = kpiArea.children.length;
  if(existing < 4){
    // prefer identifier-like columns first (counts), then numeric totals
    const allKeys = Object.keys(data[0]);
    const idKeys = allKeys.filter(k => isIdentifierKey(k));
    const numeric = columns.numeric;
    const picks = idKeys.concat(numeric).filter((v,i,a)=>a.indexOf(v)===i);
    const needed = Math.max(0, 4 - existing);
    picks.slice(0, needed).forEach(k=>{
      const kv = kpiValueForColumn(k);
      const display = kv.subtitle === 'sum' ? Number(kv.val).toLocaleString() : kv.val;
      createCard(k, display, kv.subtitle);
    });
  }

  // final fallback: first 3 keys sample
  if(kpiArea.children.length === 0){
    Object.keys(data[0]).slice(0,3).forEach(k => {
      const sample = data[0][k];
      createCard(k, typeof sample === 'number' ? sample : String(sample || '—'));
    });
  }
}

// ---------- CHART PLACEMENT: NOW ARRANGED BY TYPE ----------
function clearChartSections(){
  ['timeSeries','barCharts','donutCharts','tableCharts'].forEach(id => {
    const el = document.getElementById(id);
    if(el) el.innerHTML = '';
  });
}

function createChartCardInSection(title, parentId){
  const parent = document.getElementById(parentId);
  const card = document.createElement('div'); card.className='chart-card';
  const h = document.createElement('h4'); h.innerText = title; h.style.margin='0 0 8px 0';
  const canvas = document.createElement('canvas'); canvas.style.width='100%'; canvas.style.height='260px';
  card.appendChild(h); card.appendChild(canvas);
  parent.appendChild(card);
  return canvas;
}

// ---------- CHART BUILDING (uses sections) ----------
function buildCharts(data){
  destroyAllCharts();
  clearChartSections();
  if(!data || !data.length) return;
  const analysis = analyzeColumns(data);
  const dateKey = analysis.dateLike.length ? analysis.dateLike[0] : null;
  const numeric = analysis.numeric;
  const categorical = analysis.categorical;

  // Time series section (lines/areas)
  if(dateKey && numeric.length){
    const labels = aggregateLabelsByKey(data, dateKey, true);
    numeric.slice(0,3).forEach((k,i)=>{
      const canvas = createChartCardInSection(`Time Series — ${k}`, 'timeSeries');
      const ds = labels.map(l => sumByGroup(data, dateKey, l, k));
      charts.push(new Chart(canvas.getContext('2d'), {
        type:'line', data:{ labels, datasets:[{ label:k, data:ds, borderColor: i===0?theme.primary:theme.accent, backgroundColor:'rgba(255,255,255,0.02)', fill:true }]},
        options:{ responsive:true, maintainAspectRatio:false }
      }));
    });
  }

  // Bars section
  numeric.slice(0,4).forEach((k,i)=>{
    // if categorical exists, show by category else row-based
    if(categorical.length){
      const groupBy = categorical[0];
      const labels = Array.from(new Set(data.map(r => r[groupBy]).filter(Boolean)));
      const ds = labels.map(l => sumByGroup(data, groupBy, l, k));
      const canvas = createChartCardInSection(`${k} by ${groupBy}`, 'barCharts');
      charts.push(new Chart(canvas.getContext('2d'), { type:'bar', data:{ labels, datasets:[{ label:k, data:ds, backgroundColor: theme.primary }]}, options:{responsive:true, maintainAspectRatio:false} }));
    } else {
      const labels = data.slice(0,20).map((_,i)=>`R${i+1}`);
      const ds = data.slice(0,20).map(r => Number(r[k]||0));
      const canvas = createChartCardInSection(`${k} (first 20 rows)`, 'barCharts');
      charts.push(new Chart(canvas.getContext('2d'), { type:'bar', data:{ labels, datasets:[{ label:k, data:ds, backgroundColor: theme.primary }]}, options:{responsive:true, maintainAspectRatio:false} }));
    }
  });

  // Donut / Pie section
  if(numeric.length && categorical.length){
    const cat = categorical[0];
    const labels = Array.from(new Set(data.map(r => r[cat]).filter(Boolean)));
    numeric.slice(0,3).forEach((k,i)=>{
      const ds = labels.map(l => sumByGroup(data, cat, l, k));
      const canvas = createChartCardInSection(`Donut: ${k} by ${cat}`, 'donutCharts');
      charts.push(new Chart(canvas.getContext('2d'), { type:'doughnut', data:{ labels, datasets:[{ data:ds, backgroundColor: generatePalette(labels.length) }]}, options:{responsive:true, maintainAspectRatio:false} }));
    });
  } else if(categorical.length){
    const cat = categorical[0];
    const labels = Array.from(new Set(data.map(r => r[cat]).filter(Boolean)));
    const ds = labels.map(l => data.filter(r => r[cat] === l).length);
    const canvas = createChartCardInSection(`Distribution — ${cat}`, 'donutCharts');
    charts.push(new Chart(canvas.getContext('2d'), { type:'doughnut', data:{ labels, datasets:[{ data:ds, backgroundColor: generatePalette(labels.length) }]}, options:{responsive:true, maintainAspectRatio:false} }));
  }

  // Stacked bars: if at least two numeric & a category
  if(numeric.length >= 2 && categorical.length){
    const cat = categorical[0];
    const labels = Array.from(new Set(data.map(r => r[cat]).filter(Boolean)));
    const ds1 = labels.map(l => sumByGroup(data, cat, l, numeric[0]));
    const ds2 = labels.map(l => sumByGroup(data, cat, l, numeric[1]));
    const canvas = createChartCardInSection(`Stacked: ${numeric[0]} & ${numeric[1]} by ${cat}`, 'barCharts');
    charts.push(new Chart(canvas.getContext('2d'), { type:'bar', data:{ labels, datasets:[{ label:numeric[0], data:ds1, backgroundColor:'#4A90E2' },{ label:numeric[1], data:ds2, backgroundColor:'#50E3C2' }]}, options:{responsive:true, maintainAspectRatio:false, scales:{x:{stacked:true}, y:{stacked:true}}} }));
  }

  // Table section: small tables (for example top numeric breakdowns)
  if(numeric.length && categorical.length){
    const cat = categorical[0];
    numeric.slice(0,2).forEach(k=>{
      const labels = Array.from(new Set(data.map(r => r[cat]).filter(Boolean)));
      const rows = labels.map(l => ({ cat: l, val: sumByGroup(data, cat, l, k) }))
        .sort((a,b)=>b.val - a.val).slice(0,10);
      const container = document.getElementById('tableCharts');
      const card = document.createElement('div'); card.className='chart-card';
      const h = document.createElement('h4'); h.innerText = `Top ${k} by ${cat}`; card.appendChild(h);
      const table = document.createElement('table'); table.style.width='100%';
      const thead = document.createElement('thead'); const thr = document.createElement('tr'); ['Category', 'Value'].forEach(t=>{ const th=document.createElement('th'); th.innerText=t; thr.appendChild(th); }); thead.appendChild(thr);
      const tbody = document.createElement('tbody');
      rows.forEach(r => { const tr=document.createElement('tr'); const td1=document.createElement('td'); td1.innerText=r.cat; const td2=document.createElement('td'); td2.innerText = r.val; tr.appendChild(td1); tr.appendChild(td2); tbody.appendChild(tr); });
      table.appendChild(thead); table.appendChild(tbody); card.appendChild(table);
      container.appendChild(card);
    });
  }

  // final table preview below
  renderTablePreview(data);
}

// ---------- helper utilities ----------
function createChartCard(title, type, parentId){
  // kept in case older code uses it - but prefer createChartCardInSection
  return createChartCardInSection(title, parentId);
}

function sumColumn(data, col){
  return data.reduce((s,r)=>s + (isNaN(Number(r[col])) ? 0 : Number(r[col]||0)), 0);
}
function uniqueCount(data, col){ return new Set(data.map(r=>r[col]).filter(v=>v !== undefined && v !== '')).size; }
function sumByGroup(data, groupCol, groupValue, valueCol){
  return data.filter(r => String(r[groupCol]) === String(groupValue)).reduce((s,r)=>s + (isNaN(Number(r[valueCol]))?0:Number(r[valueCol]||0)), 0);
}
function aggregateLabelsByKey(data, key, trySortDates=false){
  let labels = Array.from(new Set(data.map(r => r[key]).filter(Boolean)));
  if(trySortDates){
    const withDates = labels.map(l => ({ raw:l, parsed: Date.parse(l) })).filter(x=>!isNaN(x.parsed));
    if(withDates.length === labels.length){
      labels = withDates.sort((a,b)=>a.parsed - b.parsed).map(x=>x.raw);
    }
  }
  return labels;
}

// ---------- PALETTE: prefer companyPalette when set ----------
function generatePalette(n){
  const base = ['#2E8B57','#f0b429','#4caf50','#2196f3','#ff9800','#f44336','#9c27b0','#009688','#8e44ad','#e67e22'];
  if(Array.isArray(companyPalette) && companyPalette.length){
    const result = [];
    for(let i=0;i<n;i++){
      if(i < companyPalette.length) result.push(companyPalette[i]);
      else result.push(base[(i - companyPalette.length) % base.length]);
    }
    return result;
  }
  return Array.from({length:n}, (_,i)=>base[i % base.length]);
}

// Table preview for first 100 rows
function renderTablePreview(data){
  const tableDiv = document.getElementById('tablePreview');
  tableDiv.innerHTML = '';
  if(!data.length){ tableDiv.innerHTML='<div style="color:var(--muted)">No data</div>'; return; }
  const table = document.createElement('table');
  const thead = document.createElement('thead'), tbody = document.createElement('tbody');
  const keys = Object.keys(data[0]);
  const trh = document.createElement('tr'); keys.forEach(k => { const th = document.createElement('th'); th.innerText = k; trh.appendChild(th); }); thead.appendChild(trh);
  data.slice(0,100).forEach(row => {
    const tr = document.createElement('tr');
    keys.forEach(k => {
      const td = document.createElement('td'); td.innerText = row[k] === undefined || row[k] === null ? '' : String(row[k]); tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(thead); table.appendChild(tbody); tableDiv.appendChild(table);
}

// ---------- top-level update ----------
function updateAllVisuals(){
  buildKPIs(filteredData);
  buildCharts(filteredData);
  // update slicer summary (if present)
  try {
    const selects = Array.from(document.querySelectorAll('#slicers select'));
    const active = selects.filter(s=>s.value && s.value !== '__all__').map(s => `${s.dataset.col}: ${s.value}`);
    const ssum = document.getElementById('slicerSummary');
    if(ssum) ssum.innerText = active.length ? active.join(' • ') : 'No filters';
  } catch(e){}
}

// ---------- POWER BI OPEN (simple) ----------
function openInPowerBI(){
  // update visuals before opening
  try { updateAllVisuals(); } catch(e){}
  // replace this URL with your published Power BI link
  const powerBiUrl = "https://app.powerbi.com/view?r=YOUR_DASHBOARD_LINK_HERE";
  window.open(powerBiUrl, "_blank");
}

// ---------- file handling ----------
function handleFile(f){
  const name = f.name.toLowerCase();
  if(name.endsWith('.csv') || name.endsWith('.txt')){
    Papa.parse(f, { header:true, dynamicTyping:true, skipEmptyLines:true, complete(results){
      rawData = normalizeParsedInput(results.data);
      filteredData = rawData.slice();
      renderSlicers(rawData);
      updateAllVisuals();
    }, error(err){ console.error(err); alert('CSV parse failed'); }});
  } else {
    const reader = new FileReader();
    reader.onload = e => {
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type:'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const aoa = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' });
        rawData = normalizeParsedInput(aoa);
        filteredData = rawData.slice();
        renderSlicers(rawData);
        updateAllVisuals();
      } catch(err){
        console.error(err);
        alert('Excel parse failed');
      }
    };
    reader.readAsArrayBuffer(f);
  }
}

// main input listeners
document.getElementById('fileInput').addEventListener('change', (ev)=>{
  const f = ev.target.files && ev.target.files[0];
  if(!f) return;
  handleFile(f);
});
document.getElementById('fileInputLanding').addEventListener('change', (ev)=>{
  const f = ev.target.files && ev.target.files[0];
  if(!f) return;
  handleFile(f);
});

// ---------- small helper findKey duplicate-free ----------
function findKey(sample, targets){
  if(!sample) return null;
  const keys = Object.keys(sample);
  for(const t of targets){
    const lt = t.toLowerCase();
    const exact = keys.find(k => k.toLowerCase() === lt);
    if(exact) return exact;
    const contains = keys.find(k => k.toLowerCase().includes(lt));
    if(contains) return contains;
  }
  return null;
}
// ---------- EXPORT PACKAGE FOR POWER BI (xlsx + config json) ----------
function openInPowerBI() {
  // Instead of trying to generate a .pbix (not possible), we export a package:
  // 1) dataset.xlsx (current filtered dataset)
  // 2) dashboard-config.json (KPIs + charts description)
  try {
    exportDatasetAsXlsx(filteredData, 'dataset.xlsx');
    const cfg = buildDashboardConfig(filteredData);
    downloadJSON(cfg, 'dashboard-config.json');
    showExportSuccessModal();
  } catch (err) {
    console.error(err);
    alert('Export failed. See console for details.');
  }
}

// Export filteredData (array of objects) as an Excel file using SheetJS (XLSX)
function exportDatasetAsXlsx(data, filename = 'dataset.xlsx') {
  if(!data || !data.length) {
    alert('No data to export.');
    return;
  }
  // convert to worksheet
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Data');
  // write file to user
  XLSX.writeFile(wb, filename);
}

// Build a JSON "dashboard-config" describing detected KPIs/charts
function buildDashboardConfig(data){
  const analysis = analyzeColumns(data);
  const config = {
    exportedAt: new Date().toISOString(),
    rowCount: data.length,
    columns: {
      numeric: analysis.numeric.slice(),
      categorical: analysis.categorical.slice(),
      dateLike: analysis.dateLike.slice()
    },
    kpis: [],
    charts: []
  };

  // Build KPIs similarly to buildKPIs logic (capture intent)
  // We'll capture top 4 KPI items and how they were computed
  const columns = Object.keys(data[0] || {});
  const numeric = analysis.numeric;
  // prefer identifiers -> unique counts; else numeric sum/avg
  const idCandidates = columns.filter(c => isIdentifierKey(c));
  const kpiCandidates = idCandidates.concat(numeric).slice(0,4);

  kpiCandidates.forEach(k => {
    const isId = isIdentifierKey(k) || analysis.categorical.includes(k);
    if(isId){
      config.kpis.push({ key: k, label: `Unique ${k}`, type: 'unique_count', value: uniqueCount(data,k) });
    } else if(numeric.includes(k)){
      config.kpis.push({ key: k, label: `Total ${k}`, type: 'sum', value: sumColumn(data,k) });
    } else {
      config.kpis.push({ key: k, label: k, type: 'sample', sample: data[0][k] });
    }
  });

  // Build charts config: time series (if date & numeric), bars (numeric by first category), donuts
  const dateKey = analysis.dateLike.length ? analysis.dateLike[0] : null;
  if(dateKey && numeric.length){
    numeric.slice(0,2).forEach(k => {
      config.charts.push({ type:'timeseries', dateField: dateKey, metric: k, aggregation: 'sum' });
    });
  }
  if(analysis.categorical.length){
    const cat = analysis.categorical[0];
    numeric.slice(0,3).forEach(k=>{
      config.charts.push({ type:'bar', category: cat, metric: k, aggregation: 'sum' });
      config.charts.push({ type:'donut', category: cat, metric: k, aggregation: 'sum' });
    });
  } else {
    // fallback simple bar charts of first numeric(s)
    numeric.slice(0,3).forEach(k => {
      config.charts.push({ type:'bar', metric: k, aggregation: 'sum', note: 'first-20-rows' });
    });
  }

  return config;
}

// helper: download JSON as file
function downloadJSON(obj, filename='config.json'){
  const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click();
  setTimeout(()=>{ URL.revokeObjectURL(url); a.remove(); }, 1000);
}

// friendly modal with instructions after export
function showExportSuccessModal(){
  const overlay = document.createElement('div');
  overlay.style.position='fixed'; overlay.style.left=0; overlay.style.top=0; overlay.style.right=0; overlay.style.bottom=0;
  overlay.style.background='rgba(0,0,0,0.6)'; overlay.style.zIndex=9999; overlay.style.display='flex';
  overlay.style.alignItems='center'; overlay.style.justifyContent='center';

  const box = document.createElement('div');
  box.style.width='720px'; box.style.maxWidth='96%'; box.style.background='#0f1724';
  box.style.border='1px solid rgba(255,255,255,0.06)'; box.style.color='#e6eef8';
  box.style.padding='18px'; box.style.borderRadius='10px';
  box.innerHTML = `
    <h3 style="margin-top:0">Export ready for Power BI Desktop</h3>
    <p style="color:var(--muted)">We exported <strong>dataset.xlsx</strong> and <strong>dashboard-config.json</strong>. Follow these steps to open and edit in Power BI Desktop:</p>
    <ol style="color:var(--muted)">
      <li>Open Power BI Desktop.</li>
      <li>Home → Get Data → Excel → choose <strong>dataset.xlsx</strong>.</li>
      <li>Option A (recommended): If you have a Power BI template (.pbit) that matches this layout, use <em>File → Import → Power BI Template</em> and then load the dataset.</li>
      <li>Option B: Recreate the visuals using the instructions in <strong>dashboard-config.json</strong> (it lists KPIs and chart hints).</li>
      <li>Once open, File → Save to create your editable <strong>.pbix</strong>.</li>
    </ol>
    <div style="display:flex;gap:8px; justify-content:flex-end; margin-top:12px;">
      <button id="exp_ok" style="padding:8px 12px; border-radius:6px; background:linear-gradient(90deg,var(--accent),var(--accent-2)); color:#001017; border:none">OK</button>
    </div>
  `;
  overlay.appendChild(box);
  document.body.appendChild(overlay);

  document.getElementById('exp_ok').onclick = ()=> overlay.remove();
}
// buildDashboardConfig - extended: dimensions, measures (DAX strings), visuals, dataTypes
function buildDashboardConfig(data){
  const analysis = analyzeColumns(data);
  const columns = Object.keys(data[0] || {});
  const numeric = analysis.numeric.slice();
  const categorical = analysis.categorical.slice();
  const dateLike = analysis.dateLike.slice();
  const rowCount = data.length;

  // helper to sample distinct values for dimensions (max 10)
  function sampleDistinct(col){
    const vals = Array.from(new Set(data.map(r => r[col]).filter(v => v !== undefined && v !== null && v !== '')));
    return vals.slice(0, 10);
  }

  // suggested measures (DAX) generator
  function suggestedMeasures(){
    const measures = [];
    // common finance-like measures
    const salesKey = findKey(data[0], ['sales','amount','units','quantity']);
    const revenueKey = findKey(data[0], ['revenue','income','salesamount','sales amount']);
    const profitKey = findKey(data[0], ['profit','net profit','margin']);
    const customerKey = findKey(data[0], ['customer','client','customerid','clientid']);
    const qtyKey = findKey(data[0], ['quantity','qty','units']);

    if(salesKey){
      measures.push({
        name: `Total ${salesKey}`,
        dax: `Total ${salesKey} = SUM('Data'[${salesKey}])`,
        description: `Total ${salesKey} (use in cards, totals).`,
        usage: ['Cards','Tables','Bar/Line']
      });
    }
    if(revenueKey){
      measures.push({
        name: `Total ${revenueKey}`,
        dax: `Total ${revenueKey} = SUM('Data'[${revenueKey}])`,
        description: `Total revenue / income.`,
        usage: ['Cards','Bar/Line','Donut']
      });
    }
    if(profitKey){
      measures.push({
        name: `Average ${profitKey}`,
        dax: `Avg ${profitKey} = AVERAGE('Data'[${profitKey}])`,
        description: `Average profit per row. Useful as Average KPI.`,
        usage: ['Cards','Line']
      });
      measures.push({
        name: `Total ${profitKey}`,
        dax: `Total ${profitKey} = SUM('Data'[${profitKey}])`,
        description: `Total profit.`,
        usage: ['Cards','Bar','Stacked Bar']
      });
    }
    if(customerKey){
      measures.push({
        name: `Unique ${customerKey}`,
        dax: `Unique ${customerKey} = DISTINCTCOUNT('Data'[${customerKey}])`,
        description: `Unique customer count.`,
        usage: ['Cards','Slicers']
      });
    }
    if(qtyKey){
      measures.push({
        name: `Total ${qtyKey}`,
        dax: `Total ${qtyKey} = SUM('Data'[${qtyKey}])`,
        description: `Total quantity sold.`,
        usage: ['Bars','Tables']
      });
    }

    // Generic: Row count
    measures.push({
      name: 'Row Count',
      dax: 'Row Count = COUNTROWS(\'Data\')',
      description: 'Number of rows in the current filter context.',
      usage: ['Cards','Tables']
    });

    // Example rate or margin if revenue & profit exist
    if(revenueKey && profitKey){
      measures.push({
        name: 'Profit Margin %',
        dax: `Profit Margin % = DIVIDE([Total ${profitKey}], [Total ${revenueKey}], 0)`,
        description: 'Profit margin as percentage. Use FORMAT() or % in visuals.',
        usage: ['Cards','Bar','Line']
      });
    }

    // date-driven measures if date column present
    if(dateLike.length){
      const d = dateLike[0];
      measures.push({
        name: `YoY ${revenueKey || 'Metric'}`,
        dax: `YoY ${revenueKey || 'Metric'} = 
  VAR Cur = [Total ${revenueKey || (numeric[0]||'Metric')}]
  VAR Prev = CALCULATE([Total ${revenueKey || (numeric[0]||'Metric')}], SAMEPERIODLASTYEAR('Data'[${d}]))
  RETURN DIVIDE(Cur - Prev, Prev, 0)`,
        description: 'Year-over-year growth for revenue-like metric (requires date table).',
        usage: ['Line','Cards']
      });
    }

    return measures;
  }

  // suggested visuals capture: type, primary field(s), aggregation
  function suggestedVisuals(){
    const visuals = [];

    // Time series if date + numeric
    if(dateLike.length && numeric.length){
      const d = dateLike[0];
      numeric.slice(0,3).forEach(k => {
        visuals.push({
          type: 'line',
          title: `Time Series — ${k}`,
          x: d,
          y: k,
          aggregation: 'SUM',
          notes: 'Use date hierarchy; consider continuous axis.'
        });
      });
    }

    // Bars: numeric by first categorical
    if(categorical.length && numeric.length){
      const cat = categorical[0];
      numeric.slice(0,4).forEach(k => {
        visuals.push({
          type: 'bar',
          title: `${k} by ${cat}`,
          category: cat,
          value: k,
          aggregation: 'SUM',
          notes: 'Top N by value or sort descending.'
        });
      });

      // Donut/pie for distribution
      visuals.push({
        type: 'donut',
        title: `Distribution of ${categorical[0]}`,
        category: categorical[0],
        metric: null,
        aggregation: 'COUNT',
        notes: 'Use as slicer companion.'
      });
    } else {
      // fallback simple bar for first numeric columns
      numeric.slice(0,3).forEach(k => {
        visuals.push({
          type: 'bar',
          title: `${k} (sample rows)`,
          category: 'Row Index',
          value: k,
          aggregation: 'SUM',
          notes: 'Shows first 20 rows; useful when no categorical dims.'
        });
      });
    }

    // tables: top N by metric
    if(categorical.length && numeric.length){
      const cat = categorical[0];
      numeric.slice(0,2).forEach(k => {
        visuals.push({
          type: 'table',
          title: `Top ${k} by ${cat}`,
          columns: [cat, k],
          notes: 'Sort by metric desc, show top 10.'
        });
      });
    }

    return visuals;
  }

  // Build dimensions list with sample values (useful to create slicers quickly)
  const dims = categorical.map(c => ({ column: c, sampleValues: sampleDistinct(c), distinctCount: new Set(data.map(r=>r[c]).filter(v=>v!==undefined && v!=='')).size }));

  // Compose final config
  const cfg = {
    exportedAt: new Date().toISOString(),
    rowCount,
    columns: columns.map(c => ({
      name: c,
      dataType: (numeric.includes(c) ? 'numeric' : (dateLike.includes(c) ? 'date' : 'text')),
      isIdentifier: isIdentifierKey(c),
      sampleValues: (categorical.includes(c) ? sampleDistinct(c) : undefined)
    })),
    dimensions: dims,
    measures: suggestedMeasures(),
    visuals: suggestedVisuals(),
    notes: 'This config is machine- and human-readable. Copy DAX into Power BI Desktop to create measures; use the visuals array to recreate report pages quickly.'
  };

  return cfg;
}