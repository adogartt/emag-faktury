<!DOCTYPE html>
<html lang="pl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>eMAG — Scalacz Wypłat</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Syne:wght@400;600;800&display=swap');
  :root {
    --bg:#0d0f14;--surface:#161920;--surface2:#1e222d;--border:#2a2f3d;
    --accent:#ff5c1a;--accent2:#ffb800;--green:#00e676;--red:#ff3d57;
    --blue:#4da6ff;--purple:#b57aff;--text:#e8eaf0;--muted:#6b7280;--radius:8px;
  }
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:var(--bg);color:var(--text);font-family:'Syne',sans-serif;min-height:100vh;padding:32px 24px;}
  header{display:flex;align-items:baseline;gap:16px;margin-bottom:36px;}
  header h1{font-size:28px;font-weight:800;letter-spacing:-0.5px;}
  header h1 span{color:var(--accent);}
  header p{font-family:'Space Mono',monospace;font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:1px;}

  .main-nav{display:flex;gap:2px;margin-bottom:28px;border-bottom:1px solid var(--border);padding-bottom:0;}
  .main-tab{padding:10px 22px;font-size:14px;font-weight:600;cursor:pointer;background:none;border:none;color:var(--muted);border-bottom:2px solid transparent;margin-bottom:-1px;transition:all 0.15s;}
  .main-tab.active{color:var(--text);border-bottom-color:var(--accent);}
  .main-tab .bc{display:inline-block;background:var(--surface2);border-radius:10px;padding:1px 7px;font-size:10px;font-family:'Space Mono',monospace;margin-left:6px;}
  .main-tab.active .bc{background:var(--accent);color:#fff;}

  .section{display:none;}
  .section.active{display:block;}

  .drop-panels{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;}
  @media(max-width:700px){.drop-panels{grid-template-columns:1fr;}}
  .drop-label{font-family:'Space Mono',monospace;font-size:10px;text-transform:uppercase;letter-spacing:1px;margin-bottom:8px;color:var(--muted);}
  .drop-zone{border:2px dashed var(--border);border-radius:var(--radius);padding:32px 20px;text-align:center;cursor:pointer;transition:border-color 0.2s,background 0.2s;background:var(--surface);position:relative;}
  .drop-zone:hover,.drop-zone.drag-over{border-color:var(--accent);background:#1a1310;}
  .drop-zone.has-files{border-color:var(--green);border-style:solid;}
  .drop-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%;}
  .drop-zone .dz-icon{font-size:32px;margin-bottom:8px;}
  .drop-zone h3{font-size:15px;font-weight:700;margin-bottom:4px;}
  .drop-zone p{font-family:'Space Mono',monospace;font-size:11px;color:var(--muted);}

  .files-list{display:flex;flex-wrap:wrap;gap:6px;margin-top:10px;}
  .file-tag{background:var(--surface2);border:1px solid var(--border);border-radius:4px;padding:3px 9px;font-family:'Space Mono',monospace;font-size:10px;color:var(--accent2);display:flex;align-items:center;gap:5px;}
  .file-tag button{background:none;border:none;color:var(--muted);cursor:pointer;font-size:12px;padding:0;}
  .file-tag button:hover{color:var(--red);}

  .controls{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:18px;}
  .btn{padding:9px 22px;border-radius:var(--radius);font-family:'Syne',sans-serif;font-weight:600;font-size:13px;cursor:pointer;border:none;transition:opacity 0.15s,transform 0.1s;}
  .btn:active{transform:scale(0.97);}
  .btn-primary{background:var(--accent);color:#fff;}
  .btn-secondary{background:var(--surface2);color:var(--text);border:1px solid var(--border);}
  .btn:disabled{opacity:0.35;cursor:not-allowed;}
  .status-bar{font-family:'Space Mono',monospace;font-size:12px;color:var(--muted);margin-top:12px;min-height:18px;}
  .progress-bar{height:3px;background:var(--border);border-radius:2px;margin-top:10px;overflow:hidden;}
  .progress-fill{height:100%;background:linear-gradient(90deg,var(--accent),var(--accent2));border-radius:2px;transition:width 0.2s;}

  .summary-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:10px;margin-bottom:24px;}
  .card{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:16px 18px;}
  .card-label{font-family:'Space Mono',monospace;font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;}
  .card-value{font-size:21px;font-weight:800;}
  .cv-green{color:var(--green);} .cv-red{color:var(--red);} .cv-orange{color:var(--accent);}
  .cv-yellow{color:var(--accent2);} .cv-blue{color:var(--blue);} .cv-purple{color:var(--purple);}

  .tab-bar{display:flex;gap:3px;margin-bottom:14px;background:var(--surface);padding:3px;border-radius:var(--radius);border:1px solid var(--border);width:fit-content;flex-wrap:wrap;}
  .tab{padding:6px 14px;border-radius:5px;font-size:12px;font-weight:600;cursor:pointer;background:none;border:none;color:var(--muted);transition:all 0.15s;white-space:nowrap;}
  .tab.active{background:var(--accent);color:#fff;}
  .tab .tc{display:inline-block;margin-left:4px;font-family:'Space Mono',monospace;font-size:10px;opacity:0.8;}

  .search-row{display:flex;gap:10px;margin-bottom:14px;align-items:center;flex-wrap:wrap;}
  .search-input{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:7px 13px;color:var(--text);font-family:'Space Mono',monospace;font-size:12px;width:240px;outline:none;}
  .search-input:focus{border-color:var(--accent);}
  .period-filter{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:7px 13px;color:var(--text);font-family:'Space Mono',monospace;font-size:12px;outline:none;cursor:pointer;}
  .row-count{font-family:'Space Mono',monospace;font-size:11px;color:var(--muted);margin-left:auto;}

  .table-wrap{overflow-x:auto;border-radius:var(--radius);border:1px solid var(--border);}
  table{width:100%;border-collapse:collapse;font-size:12px;}
  thead th{background:var(--surface2);padding:9px 13px;text-align:left;font-family:'Space Mono',monospace;font-size:10px;text-transform:uppercase;letter-spacing:1px;color:var(--muted);border-bottom:1px solid var(--border);white-space:nowrap;cursor:pointer;user-select:none;}
  thead th:hover{color:var(--text);}
  thead th.sorted .sa{color:var(--accent);}
  tbody tr{border-bottom:1px solid var(--border);transition:background 0.1s;}
  tbody tr:last-child{border-bottom:none;}
  tbody tr:hover{background:var(--surface);}
  td{padding:8px 13px;white-space:nowrap;font-family:'Space Mono',monospace;font-size:11px;}

  .badge{display:inline-block;padding:2px 7px;border-radius:3px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;}
  .b-cod{background:#1a2a1a;color:var(--green);border:1px solid #2a4a2a;}
  .b-refund{background:#2a1a1a;color:var(--red);border:1px solid #4a2a2a;}
  .b-ok{background:#1a2a1a;color:var(--green);border:1px solid #2a4a2a;}
  .b-diff{background:#2a2010;color:var(--accent2);border:1px solid #4a3a10;}
  .b-wait{background:#1e1a2a;color:var(--purple);border:1px solid #3a2a5a;}
  .b-unknown{background:#1a1a2a;color:var(--blue);border:1px solid #2a2a4a;}

  .val-p{color:var(--green);} .val-n{color:var(--red);} .val-m{color:var(--accent2);}
  .name-diff{color:var(--accent2);}
  .diff-p{color:var(--green);} .diff-n{color:var(--red);} .diff-z{color:var(--muted);}

  .pagination{display:flex;gap:5px;align-items:center;margin-top:12px;justify-content:center;}
  .page-btn{background:var(--surface);border:1px solid var(--border);border-radius:4px;padding:4px 9px;color:var(--text);font-family:'Space Mono',monospace;font-size:10px;cursor:pointer;}
  .page-btn.active{background:var(--accent);border-color:var(--accent);color:#fff;}
  .page-btn:disabled{opacity:0.3;cursor:not-allowed;}

  .loader{display:inline-block;width:14px;height:14px;border:2px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin 0.6s linear infinite;vertical-align:middle;margin-right:6px;}
  @keyframes spin{to{transform:rotate(360deg);}}

  .legend{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px;}
  .legend-item{display:flex;align-items:center;gap:6px;font-family:'Space Mono',monospace;font-size:11px;color:var(--muted);}

  .info-box{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:14px 18px;margin-bottom:20px;font-family:'Space Mono',monospace;font-size:12px;line-height:1.8;color:var(--muted);}
  .info-box strong{color:var(--text);}

  .toolbar{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;margin-bottom:14px;}
</style>
</head>
<body>

<header>
  <h1>eMAG <span>Scalacz</span> Wypłat</h1>
  <p>Scalanie + porównanie z ERP · historia wszystkich zamówień</p>
</header>

<nav class="main-nav">
  <button class="main-tab active" onclick="showMain('upload',this)">📂 Wczytaj pliki</button>
  <button class="main-tab" id="tab-emag" onclick="showMain('emag',this)">📊 Raport eMAG <span class="bc" id="cnt-emag">—</span></button>
  <button class="main-tab" id="tab-compare" onclick="showMain('compare',this)">🔍 Porównanie ERP <span class="bc" id="cnt-compare">—</span></button>
</nav>

<!-- ══════════════════════════ UPLOAD ══════════════════════════ -->
<div class="section active" id="sec-upload">
  <div class="drop-panels">
    <div>
      <div class="drop-label">Pliki wypłat eMAG (DP) — wrzuć całą historię</div>
      <div class="drop-zone" id="dzEmag">
        <input type="file" id="fiEmag" multiple accept=".xlsx,.xls"/>
        <div class="dz-icon">📦</div>
        <h3>Pliki DP z eMAG</h3>
        <p>Wrzuć wszystkie pliki naraz — program sam dopasuje okresy</p>
      </div>
      <div class="files-list" id="flEmag"></div>
    </div>
    <div>
      <div class="drop-label">Eksport z IDEAerp — cała historia zamówień</div>
      <div class="drop-zone" id="dzErp">
        <input type="file" id="fiErp" multiple accept=".xlsx,.xls,.csv"/>
        <div class="dz-icon">🗂️</div>
        <h3>Plik sale_order z ERP</h3>
        <p>Możesz wrzucić pełną historię — program sam odfiltruje co już rozliczone</p>
      </div>
      <div class="files-list" id="flErp"></div>
    </div>
  </div>
  <div class="controls">
    <button class="btn btn-primary" id="processBtn" disabled>▶ Scal i porównaj</button>
    <button class="btn btn-secondary" id="clearBtn" style="display:none">✕ Wyczyść wszystko</button>
  </div>
  <div class="status-bar" id="statusBar"></div>
  <div class="progress-bar" id="progressBar" style="display:none"><div class="progress-fill" id="progressFill" style="width:0%"></div></div>
</div>

<!-- ══════════════════════════ EMAG ══════════════════════════ -->
<div class="section" id="sec-emag">
  <div class="summary-cards" id="eMagCards"></div>
  <div class="toolbar">
    <div class="tab-bar" id="eMagTabBar">
      <button class="tab active" onclick="switchEmagTab('by-client',this)">Klienci</button>
      <button class="tab" onclick="switchEmagTab('all-rows',this)">Wszystkie wiersze</button>
      <button class="tab" onclick="switchEmagTab('by-period',this)">Okresy</button>
    </div>
    <button class="btn btn-secondary" onclick="exportEmag()">⬇ Eksportuj XLSX</button>
  </div>
  <div class="search-row">
    <input class="search-input" id="eMagSearch" placeholder="Szukaj klienta, Order ID..." oninput="applyEmagFilter()"/>
    <select class="period-filter" id="eMagPeriod" onchange="applyEmagFilter()"><option value="">Wszystkie okresy</option></select>
    <span class="row-count" id="eMagCount"></span>
  </div>
  <div class="table-wrap"><table><thead id="eMagHead"></thead><tbody id="eMagBody"></tbody></table></div>
  <div class="pagination" id="eMagPag"></div>
</div>

<!-- ══════════════════════════ PORÓWNANIE ══════════════════════════ -->
<div class="section" id="sec-compare">
  <div class="info-box">
    Matching po <strong>Order ID</strong> (kolumna <strong>Id integracji</strong> w ERP ↔ <strong>Order ID</strong> w eMAG).
    Porównywane: <strong>Suma RON</strong> z ERP vs suma <strong>Fraction value</strong> per zamówienie z eMAG.<br>
    Zamówienia z ERP których nie ma jeszcze w żadnym pliku eMAG → kategoria <strong>⏳ Oczekuje na rozliczenie</strong>.
  </div>
  <div class="summary-cards" id="compareCards"></div>
  <div class="legend">
    <div class="legend-item"><span class="badge b-ok">✓ Zgodne</span></div>
    <div class="legend-item"><span class="badge b-diff">⚠ Różnica kwot</span></div>
    <div class="legend-item"><span class="badge b-wait">⏳ Oczekuje</span> W ERP, jeszcze nie w eMAG</div>
    <div class="legend-item"><span class="badge b-unknown">? Nieznane</span> W eMAG, brak w ERP</div>
  </div>
  <div class="toolbar">
    <div class="tab-bar" id="cmpTabBar">
      <button class="tab active" onclick="switchCmpTab('all',this)">Wszystko <span class="tc" id="tc-all"></span></button>
      <button class="tab" onclick="switchCmpTab('ok',this)">✓ Zgodne <span class="tc" id="tc-ok"></span></button>
      <button class="tab" onclick="switchCmpTab('diff',this)">⚠ Różnice <span class="tc" id="tc-diff"></span></button>
      <button class="tab" onclick="switchCmpTab('wait',this)">⏳ Oczekuje <span class="tc" id="tc-wait"></span></button>
      <button class="tab" onclick="switchCmpTab('unknown',this)">? Nieznane <span class="tc" id="tc-unknown"></span></button>
    </div>
    <button class="btn btn-secondary" onclick="exportCompare()">⬇ Eksportuj raport</button>
  </div>
  <div class="search-row">
    <input class="search-input" id="cmpSearch" placeholder="Szukaj klienta, Order ID..." oninput="applyCmpFilter()"/>
    <select class="period-filter" id="cmpPeriod" onchange="applyCmpFilter()"><option value="">Wszystkie okresy eMAG</option></select>
    <span class="row-count" id="cmpCount"></span>
  </div>
  <div class="table-wrap"><table><thead id="cmpHead"></thead><tbody id="cmpBody"></tbody></table></div>
  <div class="pagination" id="cmpPag"></div>
</div>

<script>
// ─── State ───────────────────────────────────────────────────────────────
let eMagFiles=[], erpFiles=[];
let allEmagRows=[], allErpRows=[], compareRows=[];
let clientSummary=[], periodSummary=[];
let eMagTab='by-client', cmpTab='all';
let eMagPage=1, cmpPage=1;
const PAGE=50;
let eMagFiltered=[], cmpFiltered=[];
let eS=null,eSd=1, cS=null,cSd=1;

// ─── Nav ──────────────────────────────────────────────────────────────────
function showMain(sec, btn) {
  document.querySelectorAll('.section').forEach(s=>s.classList.remove('active'));
  document.querySelectorAll('.main-tab').forEach(t=>t.classList.remove('active'));
  document.getElementById('sec-'+sec).classList.add('active');
  if(btn) btn.classList.add('active');
}

// ─── Drop zones ───────────────────────────────────────────────────────────
function setupDrop(dzId,inputId,store,listId) {
  const dz=document.getElementById(dzId), fi=document.getElementById(inputId);
  dz.addEventListener('dragover',e=>{e.preventDefault();dz.classList.add('drag-over');});
  dz.addEventListener('dragleave',()=>dz.classList.remove('drag-over'));
  dz.addEventListener('drop',e=>{e.preventDefault();dz.classList.remove('drag-over');addFiles(store,[...e.dataTransfer.files],listId,dz);});
  fi.addEventListener('change',()=>addFiles(store,[...fi.files],listId,dz));
}
setupDrop('dzEmag','fiEmag',eMagFiles,'flEmag');
setupDrop('dzErp','fiErp',erpFiles,'flErp');

function addFiles(store,files,listId,dz) {
  files.forEach(f=>{if(!store.find(x=>x.name===f.name))store.push(f);});
  renderList(store,listId,dz); checkReady();
}
window.removeFileFn=function(storeName,name,listId,dzId) {
  const store=storeName==='e'?eMagFiles:erpFiles;
  const i=store.findIndex(f=>f.name===name); if(i>=0)store.splice(i,1);
  renderList(store,listId,document.getElementById(dzId)); checkReady();
};
function renderList(store,listId,dz) {
  const isEmag=store===eMagFiles;
  document.getElementById(listId).innerHTML=store.map(f=>
    `<div class="file-tag">${f.name}<button onclick="removeFileFn('${isEmag?'e':'r'}','${f.name.replace(/'/g,"\\'")}','${listId}','${dz.id}')">×</button></div>`
  ).join('');
  dz.classList.toggle('has-files',store.length>0);
}
function checkReady() {
  document.getElementById('processBtn').disabled=eMagFiles.length===0;
  document.getElementById('clearBtn').style.display=(eMagFiles.length||erpFiles.length)?'':'none';
}
document.getElementById('clearBtn').addEventListener('click',()=>{
  eMagFiles.length=0; erpFiles.length=0;
  allEmagRows=[]; allErpRows=[]; compareRows=[];
  ['flEmag','flErp'].forEach(id=>document.getElementById(id).innerHTML='');
  ['dzEmag','dzErp'].forEach(id=>document.getElementById(id).classList.remove('has-files'));
  checkReady(); setStatus('');
  document.getElementById('cnt-emag').textContent='—';
  document.getElementById('cnt-compare').textContent='—';
});

// ─── Processing ───────────────────────────────────────────────────────────
document.getElementById('processBtn').addEventListener('click',processAll);

async function processAll() {
  allEmagRows=[]; allErpRows=[];
  const pb=document.getElementById('progressBar'), pf=document.getElementById('progressFill');
  pb.style.display=''; let done=0;
  const total=eMagFiles.length+erpFiles.length;

  setStatus('<span class="loader"></span>Wczytuję pliki eMAG...');
  for(const f of eMagFiles) {
    pf.style.width=((done/total)*70)+'%';
    try { (await readFile(f)).forEach(r=>allEmagRows.push({...r,_source:f.name})); }
    catch(e) { setStatus(`❌ Błąd: ${f.name}: ${e.message}`); }
    done++;
  }

  if(erpFiles.length) {
    setStatus('<span class="loader"></span>Wczytuję plik ERP...');
    for(const f of erpFiles) {
      pf.style.width=((done/total)*70)+'%';
      try { (await readFile(f)).forEach(r=>allErpRows.push(r)); }
      catch(e) { setStatus(`❌ Błąd ERP: ${f.name}: ${e.message}`); }
      done++;
    }
  }

  pf.style.width='85%'; setStatus('<span class="loader"></span>Analizuję...'); await sleep(50);
  computeEmagSummaries();
  if(allErpRows.length) computeComparison();
  pf.style.width='100%'; await sleep(200); pb.style.display='none';

  const msg=`✅ ${allEmagRows.length} wierszy eMAG z ${eMagFiles.length} pliku(ów)`+(allErpRows.length?` + ${allErpRows.length} zamówień ERP`:'');
  setStatus(msg);
  document.getElementById('cnt-emag').textContent=allEmagRows.length;
  if(allErpRows.length) document.getElementById('cnt-compare').textContent=compareRows.length;
  renderPeriodFilters();
  applyEmagFilter(); applyCmpFilter();
}

function readFile(file) {
  return new Promise((res,rej)=>{
    const reader=new FileReader();
    const isCsv=file.name.toLowerCase().endsWith('.csv');
    reader.onload=e=>{
      try {
        const wb=XLSX.read(e.target.result,{type:isCsv?'string':'array',cellDates:true});
        res(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''}));
      } catch(err){rej(err);}
    };
    reader.onerror=rej;
    if(isCsv) reader.readAsText(file,'utf-8');
    else reader.readAsArrayBuffer(file);
  });
}

// ─── eMAG summaries ───────────────────────────────────────────────────────
function computeEmagSummaries() {
  const byC={};
  allEmagRows.forEach(r=>{
    const k=(r['Client name']||'Brak').trim();
    if(!byC[k]) byC[k]={client:k,orders:new Set(),totalCOD:0,totalRefund:0,rows:0,periods:new Set()};
    const val=parseFloat(r['Fraction value'])||0;
    const isR=(r['Fraction type']||'').toLowerCase().includes('refund');
    byC[k].orders.add(r['Order ID']); byC[k].rows++;
    byC[k].periods.add(fmtPeriod(r));
    if(isR) byC[k].totalRefund+=val; else byC[k].totalCOD+=val;
  });
  clientSummary=Object.values(byC).map(c=>({
    client:c.client,orders:c.orders.size,rows:c.rows,
    totalCOD:r2(c.totalCOD),totalRefund:r2(c.totalRefund),
    net:r2(c.totalCOD+c.totalRefund),periods:[...c.periods].join(', ')
  })).sort((a,b)=>b.net-a.net);

  const byP={};
  allEmagRows.forEach(r=>{
    const k=fmtPeriod(r);
    if(!byP[k]) byP[k]={period:k,clients:new Set(),totalCOD:0,totalRefund:0,rows:0};
    const val=parseFloat(r['Fraction value'])||0;
    const isR=(r['Fraction type']||'').toLowerCase().includes('refund');
    byP[k].clients.add((r['Client name']||'').trim()); byP[k].rows++;
    if(isR) byP[k].totalRefund+=val; else byP[k].totalCOD+=val;
  });
  periodSummary=Object.values(byP).map(p=>({
    period:p.period,clients:p.clients.size,rows:p.rows,
    totalCOD:r2(p.totalCOD),totalRefund:r2(p.totalRefund),net:r2(p.totalCOD+p.totalRefund)
  })).sort((a,b)=>a.period.localeCompare(b.period));

  const cod=allEmagRows.filter(r=>!(r['Fraction type']||'').toLowerCase().includes('refund')).reduce((s,r)=>s+(parseFloat(r['Fraction value'])||0),0);
  const ref=allEmagRows.filter(r=>(r['Fraction type']||'').toLowerCase().includes('refund')).reduce((s,r)=>s+(parseFloat(r['Fraction value'])||0),0);
  const uc=new Set(allEmagRows.map(r=>(r['Client name']||'').trim())).size;

  document.getElementById('eMagCards').innerHTML=`
    <div class="card"><div class="card-label">COD Suma</div><div class="card-value cv-green">${fmtM(cod)}</div></div>
    <div class="card"><div class="card-label">Zwroty</div><div class="card-value cv-red">${fmtM(ref)}</div></div>
    <div class="card"><div class="card-label">Netto</div><div class="card-value cv-orange">${fmtM(cod+ref)}</div></div>
    <div class="card"><div class="card-label">Klientów</div><div class="card-value cv-yellow">${uc}</div></div>
    <div class="card"><div class="card-label">Wierszy łącznie</div><div class="card-value">${allEmagRows.length}</div></div>
    <div class="card"><div class="card-label">Plików eMAG</div><div class="card-value">${eMagFiles.length}</div></div>
  `;
}

// ─── Comparison ───────────────────────────────────────────────────────────
function computeComparison() {
  // Aggregate eMAG per Order ID
  const emByOid={};
  allEmagRows.forEach(r=>{
    const oid=String(r['Order ID']||'').trim(); if(!oid) return;
    if(!emByOid[oid]) emByOid[oid]={oid,clientEmag:(r['Client name']||'').trim(),total:0,fractions:[],period:fmtPeriod(r),payoutDate:r['Payout date']};
    emByOid[oid].total=r2(emByOid[oid].total+(parseFloat(r['Fraction value'])||0));
    emByOid[oid].fractions.push(r['Fraction type']);
  });

  // ERP by Id integracji
  const erpByOid={};
  allErpRows.forEach(r=>{
    const oid=String(r['Id integracji']||'').trim(); if(!oid) return;
    if(!erpByOid[oid]) erpByOid[oid]={
      oid, numer:r['Numer']||'',
      clientErp:(r['Klient']||'').trim(),
      sumaRon:parseFloat(r['Suma'])||0,
      sumaPln:parseFloat(r['Suma przewalutowana'])||0,
      status:r['Szczegółowy status zamówienia']||'',
      orderDate:r['Data zamówienia']||''
    };
  });

  compareRows=[];
  const allIds=new Set([...Object.keys(emByOid),...Object.keys(erpByOid)]);

  allIds.forEach(oid=>{
    const em=emByOid[oid], erp=erpByOid[oid];
    if(em && erp) {
      const diff=r2(em.total-erp.sumaRon);
      compareRows.push({
        status: Math.abs(diff)>0.05?'diff':'ok',
        oid, clientEmag:em.clientEmag, clientErp:erp.clientErp,
        nameDiff: em.clientEmag.toLowerCase()!==erp.clientErp.toLowerCase(),
        eMagTotal:em.total, erpRon:erp.sumaRon, erpPln:erp.sumaPln,
        diff, numer:erp.numer, erpStatus:erp.status,
        fractions:[...new Set(em.fractions)].join(', '),
        period:em.period, payoutDate:em.payoutDate, orderDate:erp.orderDate
      });
    } else if(erp && !em) {
      // In ERP but not yet in any eMAG file → Awaiting
      compareRows.push({
        status:'wait',
        oid, clientEmag:'—', clientErp:erp.clientErp, nameDiff:false,
        eMagTotal:null, erpRon:erp.sumaRon, erpPln:erp.sumaPln,
        diff:null, numer:erp.numer, erpStatus:erp.status,
        fractions:'—', period:'—', payoutDate:'—', orderDate:erp.orderDate
      });
    } else {
      // In eMAG but not in ERP → Unknown
      compareRows.push({
        status:'unknown',
        oid, clientEmag:em.clientEmag, clientErp:'—', nameDiff:false,
        eMagTotal:em.total, erpRon:null, erpPln:null,
        diff:null, numer:'—', erpStatus:'—',
        fractions:[...new Set(em.fractions)].join(', '),
        period:em.period, payoutDate:em.payoutDate, orderDate:'—'
      });
    }
  });

  // Count by status
  const cnt={ok:0,diff:0,wait:0,unknown:0};
  compareRows.forEach(r=>cnt[r.status]=(cnt[r.status]||0)+1);
  ['ok','diff','wait','unknown'].forEach(s=>{
    const el=document.getElementById('tc-'+s);
    if(el) el.textContent=cnt[s]||0;
  });
  document.getElementById('tc-all').textContent=compareRows.length;

  const sumDiff=r2(compareRows.filter(r=>r.diff!==null).reduce((s,r)=>s+r.diff,0));
  const waitRon=r2(compareRows.filter(r=>r.status==='wait').reduce((s,r)=>s+r.erpRon,0));

  document.getElementById('compareCards').innerHTML=`
    <div class="card"><div class="card-label">Zgodnych</div><div class="card-value cv-green">${cnt.ok||0}</div></div>
    <div class="card"><div class="card-label">Różnice kwot</div><div class="card-value ${cnt.diff?'cv-red':'cv-green'}">${cnt.diff||0}</div></div>
    <div class="card"><div class="card-label">⏳ Oczekuje</div><div class="card-value cv-purple">${cnt.wait||0}</div></div>
    <div class="card"><div class="card-label">Wartość oczekujących</div><div class="card-value cv-purple">${fmtM(waitRon)}</div></div>
    <div class="card"><div class="card-label">? Nieznane</div><div class="card-value ${cnt.unknown?'cv-yellow':'cv-green'}">${cnt.unknown||0}</div></div>
    <div class="card"><div class="card-label">Suma różnic</div><div class="card-value ${Math.abs(sumDiff)>0.05?'cv-red':'cv-green'}">${fmtM(sumDiff)}</div></div>
  `;
}

// ─── eMAG table ───────────────────────────────────────────────────────────
function switchEmagTab(t,btn) {
  eMagTab=t; eMagPage=1; eS=null;
  document.querySelectorAll('#eMagTabBar .tab').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  document.getElementById('eMagSearch').value='';
  applyEmagFilter();
}
function applyEmagFilter() {
  const q=document.getElementById('eMagSearch').value.toLowerCase();
  const period=document.getElementById('eMagPeriod').value;
  let data;
  if(eMagTab==='by-client') {
    data=clientSummary.filter(r=>(!q||r.client.toLowerCase().includes(q))&&(!period||r.periods.includes(period)));
  } else if(eMagTab==='all-rows') {
    data=allEmagRows.filter(r=>{
      const m=!q||Object.values(r).some(v=>String(v).toLowerCase().includes(q));
      return m&&(!period||fmtPeriod(r)===period);
    });
  } else {
    data=periodSummary.filter(r=>!q||r.period.toLowerCase().includes(q));
  }
  eMagFiltered=data; eMagPage=1; renderEmagTable();
}
function renderEmagTable() {
  let cols,rr;
  if(eMagTab==='by-client') {
    cols=[{k:'client',l:'Klient'},{k:'orders',l:'Zamówień'},{k:'totalCOD',l:'COD Suma'},{k:'totalRefund',l:'Zwroty'},{k:'net',l:'Netto'},{k:'periods',l:'Okresy'}];
    rr=r=>`<td>${r.client}</td><td>${r.orders}</td><td class="val-p">${fmtM(r.totalCOD)}</td><td class="${r.totalRefund<0?'val-n':'val-m'}">${fmtM(r.totalRefund)}</td><td class="${r.net>=0?'val-p':'val-n'}">${fmtM(r.net)}</td><td style="color:var(--muted);font-size:10px">${r.periods}</td>`;
  } else if(eMagTab==='all-rows') {
    cols=[{k:'Client name',l:'Klient'},{k:'Order ID',l:'Order ID'},{k:'Fraction type',l:'Typ'},{k:'Fraction value',l:'Wartość'},{k:'Order date',l:'Data zamówienia'},{k:'Courier',l:'Kurier'},{k:'_source',l:'Plik'}];
    rr=r=>{
      const v=parseFloat(r['Fraction value'])||0;
      const isR=(r['Fraction type']||'').toLowerCase().includes('refund');
      return `<td>${r['Client name']||'—'}</td><td>${r['Order ID']||'—'}</td><td><span class="badge ${isR?'b-refund':'b-cod'}">${r['Fraction type']||'—'}</span></td><td class="${v<0?'val-n':'val-p'}">${fmtM(v)}</td><td>${fmtD(r['Order date'])}</td><td style="color:var(--muted)">${r['Courier']||'—'}</td><td style="color:var(--muted);font-size:10px">${r['_source']||''}</td>`;
    };
  } else {
    cols=[{k:'period',l:'Okres'},{k:'clients',l:'Klientów'},{k:'rows',l:'Wierszy'},{k:'totalCOD',l:'COD Suma'},{k:'totalRefund',l:'Zwroty'},{k:'net',l:'Netto'}];
    rr=r=>`<td>${r.period}</td><td>${r.clients}</td><td>${r.rows}</td><td class="val-p">${fmtM(r.totalCOD)}</td><td class="${r.totalRefund<0?'val-n':'val-m'}">${fmtM(r.totalRefund)}</td><td class="${r.net>=0?'val-p':'val-n'}">${fmtM(r.net)}</td>`;
  }
  renderTable('eMagHead','eMagBody','eMagPag','eMagCount',cols,eMagFiltered,eMagPage,eS,eSd,rr,
    (c,d)=>{eS=c;eSd=d;renderEmagTable();},p=>{eMagPage=p;renderEmagTable();});
}

// ─── Compare table ────────────────────────────────────────────────────────
function switchCmpTab(t,btn) {
  cmpTab=t; cmpPage=1; cS=null;
  document.querySelectorAll('#cmpTabBar .tab').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  document.getElementById('cmpSearch').value='';
  applyCmpFilter();
}
function applyCmpFilter() {
  const q=document.getElementById('cmpSearch').value.toLowerCase();
  const period=document.getElementById('cmpPeriod').value;
  let data=compareRows;
  if(cmpTab==='ok') data=data.filter(r=>r.status==='ok');
  else if(cmpTab==='diff') data=data.filter(r=>r.status==='diff');
  else if(cmpTab==='wait') data=data.filter(r=>r.status==='wait');
  else if(cmpTab==='unknown') data=data.filter(r=>r.status==='unknown');
  if(q) data=data.filter(r=>r.oid.includes(q)||r.clientEmag.toLowerCase().includes(q)||r.clientErp.toLowerCase().includes(q));
  if(period) data=data.filter(r=>r.period===period);
  cmpFiltered=data; cmpPage=1; renderCmpTable();
}
function renderCmpTable() {
  const cols=[
    {k:'status',l:'Status'},{k:'oid',l:'Order ID'},{k:'clientEmag',l:'Klient eMAG'},
    {k:'clientErp',l:'Klient ERP'},{k:'eMagTotal',l:'Kwota eMAG'},{k:'erpRon',l:'Kwota ERP (RON)'},
    {k:'erpPln',l:'Kwota ERP (PLN)'},{k:'diff',l:'Różnica'},{k:'period',l:'Okres eMAG'},
    {k:'erpStatus',l:'Status ERP'},{k:'numer',l:'Nr ERP'}
  ];
  const BADGE={ok:'<span class="badge b-ok">✓ Zgodne</span>',diff:'<span class="badge b-diff">⚠ Różnica</span>',wait:'<span class="badge b-wait">⏳ Oczekuje</span>',unknown:'<span class="badge b-unknown">? Nieznane</span>'};
  const rr=r=>{
    const nc=r.nameDiff?'class="name-diff"':'';
    const diffStr=r.diff===null?'—':fmtM(r.diff);
    const diffCls=r.diff===null?'':Math.abs(r.diff)<0.05?'diff-z':r.diff>0?'diff-p':'diff-n';
    return `<td>${BADGE[r.status]||r.status}</td>
      <td>${r.oid}</td>
      <td ${nc}>${r.clientEmag}</td>
      <td ${nc}>${r.clientErp}</td>
      <td class="${r.eMagTotal!==null?'val-p':''}">${r.eMagTotal!==null?fmtM(r.eMagTotal):'—'}</td>
      <td class="${r.erpRon!==null?'val-p':''}">${r.erpRon!==null?fmtM(r.erpRon):'—'}</td>
      <td style="color:var(--muted)">${r.erpPln!==null?r.erpPln.toLocaleString('pl-PL',{minimumFractionDigits:2,maximumFractionDigits:2})+' PLN':'—'}</td>
      <td class="${diffCls}">${diffStr}</td>
      <td style="color:var(--muted);font-size:10px">${r.period}</td>
      <td style="color:var(--muted);font-size:10px">${r.erpStatus}</td>
      <td style="color:var(--muted);font-size:10px">${r.numer}</td>`;
  };
  renderTable('cmpHead','cmpBody','cmpPag','cmpCount',cols,cmpFiltered,cmpPage,cS,cSd,rr,
    (c,d)=>{cS=c;cSd=d;renderCmpTable();},p=>{cmpPage=p;renderCmpTable();});
}

// ─── Generic table renderer ───────────────────────────────────────────────
function renderTable(hId,bId,pId,cntId,cols,data,page,sortCol,sortDir,rowRender,onSort,onPage) {
  let sorted=data;
  if(sortCol!==null) {
    const k=cols[sortCol].k;
    sorted=[...data].sort((a,b)=>{
      const av=a[k],bv=b[k];
      if(av===null||av===undefined) return 1;
      if(bv===null||bv===undefined) return -1;
      if(typeof av==='number'&&typeof bv==='number') return sortDir*(av-bv);
      return sortDir*String(av).localeCompare(String(bv));
    });
  }
  const total=sorted.length, pages=Math.max(1,Math.ceil(total/PAGE));
  const p=Math.min(page,pages);
  const slice=sorted.slice((p-1)*PAGE,p*PAGE);

  const head=document.getElementById(hId);
  head.innerHTML='<tr>'+cols.map((c,i)=>`<th class="${sortCol===i?'sorted':''}" data-i="${i}">${c.l}<span class="sa"> ${sortCol===i?(sortDir===1?'↑':'↓'):'↕'}</span></th>`).join('')+'</tr>';
  head.querySelectorAll('th').forEach(th=>{
    const i=parseInt(th.dataset.i);
    th.onclick=()=>onSort(i,sortCol===i?-sortDir:1);
  });

  document.getElementById(cntId).textContent=`${total} wynik${total===1?'':'ów'}`;

  const tbody=document.getElementById(bId);
  if(!slice.length) {
    tbody.innerHTML=`<tr><td colspan="${cols.length}" style="text-align:center;padding:40px;color:var(--muted);font-family:'Space Mono',monospace">Brak wyników</td></tr>`;
  } else {
    tbody.innerHTML=slice.map(r=>`<tr>${rowRender(r)}</tr>`).join('');
  }

  // Pagination
  const el=document.getElementById(pId);
  if(pages<=1){el.innerHTML='';return;}
  let html=`<button class="page-btn" ${p===1?'disabled':''} data-pg="${p-1}">‹</button>`;
  const range=[];
  for(let i=1;i<=pages;i++){
    if(i===1||i===pages||Math.abs(i-p)<=2) range.push(i);
    else if(range[range.length-1]!=='…') range.push('…');
  }
  range.forEach(n=>{
    if(n==='…') html+=`<span class="page-btn" style="cursor:default">…</span>`;
    else html+=`<button class="page-btn ${n===p?'active':''}" data-pg="${n}">${n}</button>`;
  });
  html+=`<button class="page-btn" ${p===pages?'disabled':''} data-pg="${p+1}">›</button>`;
  el.innerHTML=html;
  el.querySelectorAll('button[data-pg]').forEach(btn=>{
    btn.onclick=()=>onPage(parseInt(btn.dataset.pg));
  });
}

// ─── Exports ──────────────────────────────────────────────────────────────
function exportEmag() {
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Klient','Zamówień','COD Suma','Zwroty','Netto','Okresy'],...clientSummary.map(r=>[r.client,r.orders,r.totalCOD,r.totalRefund,r.net,r.periods])]),'Klienci');
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Okres','Klientów','Wierszy','COD Suma','Zwroty','Netto'],...periodSummary.map(r=>[r.period,r.clients,r.rows,r.totalCOD,r.totalRefund,r.net])]),'Okresy');
  const ac=['Client name','Order ID','Fraction type','Fraction value','Order date','Payout date','Reference period start','Reference period end','Courier','_source'];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([ac,...allEmagRows.map(r=>ac.map(c=>r[c]??''))]),'Wszystkie wiersze');
  XLSX.writeFile(wb,`emag_${new Date().toISOString().slice(0,10)}.xlsx`);
}
function exportCompare() {
  const wb=XLSX.utils.book_new();
  const hdr=['Status','Order ID','Klient eMAG','Klient ERP','Kwota eMAG (RON)','Kwota ERP (RON)','Kwota ERP (PLN)','Różnica','Okres eMAG','Status ERP','Nr ERP'];
  const toRow=r=>[r.status,r.oid,r.clientEmag,r.clientErp,r.eMagTotal??'',r.erpRon??'',r.erpPln??'',r.diff??'',r.period,r.erpStatus,r.numer];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([hdr,...compareRows.map(toRow)]),'Porównanie');
  const diffRows=compareRows.filter(r=>r.status==='diff');
  if(diffRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([hdr,...diffRows.map(toRow)]),'Różnice kwot');
  const waitRows=compareRows.filter(r=>r.status==='wait');
  if(waitRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([hdr,...waitRows.map(toRow)]),'Oczekuje na rozliczenie');
  const unkRows=compareRows.filter(r=>r.status==='unknown');
  if(unkRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([hdr,...unkRows.map(toRow)]),'Nieznane');
  XLSX.writeFile(wb,`emag_porownanie_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ─── Helpers ──────────────────────────────────────────────────────────────
function renderPeriodFilters() {
  const periods=[...new Set(allEmagRows.map(fmtPeriod))].sort();
  const opts='<option value="">Wszystkie okresy</option>'+periods.map(p=>`<option value="${p}">${p}</option>`).join('');
  document.getElementById('eMagPeriod').innerHTML=opts;
  document.getElementById('cmpPeriod').innerHTML=opts;
}
function fmtPeriod(r) {
  const f=d=>{if(!d)return'?';if(d instanceof Date)return d.toISOString().slice(0,10);return String(d).slice(0,10);};
  return `${f(r['Reference period start'])} → ${f(r['Reference period end'])}`;
}
function fmtM(v) {
  if(v===undefined||v===null)return'—';
  return v.toLocaleString('pl-PL',{minimumFractionDigits:2,maximumFractionDigits:2})+' RON';
}
function fmtD(d){if(!d)return'—';if(d instanceof Date)return d.toISOString().slice(0,10);return String(d).slice(0,10);}
function r2(v){return Math.round(v*100)/100;}
function sleep(ms){return new Promise(r=>setTimeout(r,ms));}
function setStatus(html){document.getElementById('statusBar').innerHTML=html;}
</script>
</body>
</html>
