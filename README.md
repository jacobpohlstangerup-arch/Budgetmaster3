# Budget Master

<html lang="da">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>Forskningsbudget</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:ital,wght@0,300;0,400;0,500;0,600;1,400&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@2.44.0/tabler-icons.min.css"/>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Sortable/1.15.2/Sortable.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<style>
:root{
  --bg:#F7F6F3;--surface:#FFFFFF;--surface2:#F0EFE9;
  --text:#1C1B18;--text2:#6B6860;--text3:#A8A69E;
  --accent:#C0392B;--accent-light:#FDECEA;
  --blue:#1A56B0;--blue-light:#EBF2FD;--blue-total:#D6E4F7;
  --green:#1A7A4A;--green-light:#E6F5ED;
  --border:#E2E0D8;--border2:#C8C6BC;
  --radius:8px;--radius-lg:12px;--radius-xl:16px;
  --shadow:0 1px 3px rgba(0,0,0,.08);
  font-family:'DM Sans',sans-serif;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{background:var(--bg);color:var(--text);min-height:100vh;font-size:14px;line-height:1.5;}
.app{max-width:1240px;margin:0 auto;padding:1.5rem 1rem 5rem;}
.app-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:1.5rem;flex-wrap:wrap;gap:12px;}
.app-title{font-size:20px;font-weight:600;letter-spacing:-.3px;}
.app-title small{font-size:12px;font-weight:400;color:var(--text3);display:block;margin-top:2px;}
.header-controls{display:flex;align-items:center;gap:6px;flex-wrap:wrap;}
.pill-group{display:flex;gap:2px;background:var(--surface2);border-radius:var(--radius);padding:3px;}
.pill{padding:4px 10px;font-size:12px;font-weight:500;cursor:pointer;border:none;border-radius:6px;background:transparent;color:var(--text2);transition:all .15s;font-family:inherit;}
.pill.active{background:var(--surface);color:var(--text);box-shadow:var(--shadow);}
.lang-toggle{display:flex;border:1px solid var(--border2);border-radius:var(--radius);overflow:hidden;}
.lang-btn{padding:5px 10px;cursor:pointer;border:none;background:transparent;color:var(--text2);font-family:inherit;font-weight:600;font-size:12px;transition:all .15s;}
.lang-btn.active{background:var(--text);color:var(--bg);}
.icon-btn{display:inline-flex;align-items:center;gap:5px;padding:6px 12px;font-size:12px;font-weight:500;cursor:pointer;border:1px solid var(--border2);border-radius:var(--radius);background:var(--surface);color:var(--text);font-family:inherit;transition:all .15s;white-space:nowrap;}
.icon-btn:hover{background:var(--surface2);}
.icon-btn:disabled{opacity:.4;cursor:not-allowed;}
.icon-btn.primary{background:var(--accent);color:#fff;border-color:var(--accent);}
.icon-btn.primary:hover{background:#a93226;}
.icon-btn.success{background:var(--green);color:#fff;border-color:var(--green);}
.icon-btn.danger{color:var(--accent);border-color:transparent;background:transparent;}
.icon-btn.danger:hover{background:var(--accent-light);}
.icon-btn.ghost{border-color:transparent;background:transparent;color:var(--text2);}
.icon-btn.ghost:hover{background:var(--surface2);color:var(--text);}
.config-bar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:10px 14px;background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);margin-bottom:1rem;font-size:12px;color:var(--text2);}
.config-bar label{font-weight:500;white-space:nowrap;}
.config-bar input[type=text]{font-size:12px;padding:4px 8px;border:1px solid var(--border2);border-radius:var(--radius);background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;}
.config-bar input[type=number]{font-size:12px;padding:4px 8px;border:1px solid var(--border2);border-radius:var(--radius);background:var(--bg);color:var(--text);font-family:'DM Mono',monospace;text-align:right;}
.config-sep{width:1px;height:20px;background:var(--border2);flex-shrink:0;}
.share-bar{display:flex;align-items:center;gap:8px;padding:8px 14px;background:var(--blue-light);border:1px solid #A8C4E8;border-radius:var(--radius-lg);margin-bottom:1rem;font-size:12px;}
.share-bar input{flex:1;font-size:11px;padding:4px 8px;border:1px solid var(--border2);border-radius:var(--radius);background:var(--surface);color:var(--text);font-family:'DM Mono',monospace;}
.summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:1.5rem;}
.metric{background:var(--surface);border:1px solid var(--border);border-radius:var(--radius-lg);padding:12px 14px;}
.metric-label{font-size:10px;color:var(--text3);text-transform:uppercase;letter-spacing:.6px;margin-bottom:3px;font-weight:600;}
.metric-value{font-size:15px;font-weight:600;font-family:'DM Mono',monospace;}
.metric-infl{font-size:10px;color:var(--text3);font-family:'DM Mono',monospace;margin-top:1px;}
.panel{border:1px solid var(--border);border-radius:var(--radius-lg);margin-bottom:1rem;overflow:hidden;background:var(--surface);}
.panel-header{display:flex;align-items:center;justify-content:space-between;padding:9px 14px;background:var(--surface2);cursor:pointer;user-select:none;}
.panel-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;display:flex;align-items:center;gap:7px;color:var(--text2);}
.chevron{font-size:13px;color:var(--text3);transition:transform .2s;display:inline-block;}
.chevron.open{transform:rotate(180deg);}
.rates-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:8px;padding:12px 14px;}
.rate-row{display:flex;flex-direction:column;gap:2px;}
.rate-label{font-size:11px;color:var(--text2);font-weight:500;}
.rate-inp{font-size:12px;padding:4px 8px;border:1px solid var(--border2);border-radius:var(--radius);background:var(--bg);color:var(--text);font-family:'DM Mono',monospace;text-align:right;}
.rate-inp:focus{outline:none;box-shadow:0 0 0 2px var(--blue-light);border-color:var(--blue);}
.sections-wrap{display:flex;flex-direction:column;gap:8px;margin-bottom:10px;}
.section{border:1px solid var(--border);border-radius:var(--radius-xl);overflow:hidden;background:var(--surface);}
.section-header{display:flex;align-items:center;gap:8px;padding:9px 12px;background:var(--surface2);cursor:pointer;user-select:none;}
.section-drag-handle{color:var(--text3);cursor:grab;font-size:16px;flex-shrink:0;}
.section-drag-handle:active{cursor:grabbing;}
.section-dot{width:9px;height:9px;border-radius:2px;flex-shrink:0;}
.section-name-wrap{flex:1;min-width:0;}
.section-name-inp{font-size:13px;font-weight:600;border:none;background:transparent;color:var(--text);font-family:'DM Sans',sans-serif;width:100%;outline:none;}
.section-name-inp:focus{background:var(--surface);border-radius:4px;padding:2px 5px;}
.section-controls{display:flex;align-items:center;gap:3px;flex-shrink:0;}
.section-total-badge{font-size:12px;font-weight:600;font-family:'DM Mono',monospace;white-space:nowrap;margin-right:4px;}
.section-body{border-top:1px solid var(--border);}
.section-desc-row{padding:7px 14px;border-bottom:1px solid var(--border);background:var(--bg);}
.tbl-wrap{overflow-x:auto;}
table{width:100%;border-collapse:collapse;font-size:12px;}
th{font-size:9px;font-weight:700;color:var(--text3);text-align:left;padding:6px 7px;border-bottom:1px solid var(--border);background:var(--surface2);white-space:nowrap;text-transform:uppercase;letter-spacing:.4px;}
th.r{text-align:right;}
th.fte-hdr{background:#EEF4FC;border-left:2px solid var(--blue-total);text-align:right;}
th.bud-hdr{background:#EEF4FC;border-right:2px solid var(--blue-total);text-align:right;}
th.yr-group{background:var(--blue-total);color:var(--blue);text-align:center;font-weight:700;border-left:2px solid #A8C4E8;border-right:2px solid #A8C4E8;font-size:9px;letter-spacing:.5px;}
td{padding:5px 7px;border-bottom:1px solid var(--border);vertical-align:middle;color:var(--text);}
tr:last-child td{border-bottom:none;}
tr.data-row:hover td{background:#FAFAF8;}
tr.hidden-row{display:none;}
td.fte-cell{background:#F7FAFF;border-left:2px solid var(--blue-total);text-align:right;}
td.bud-cell{background:#F7FAFF;border-right:2px solid var(--blue-total);text-align:right;}
.drag-handle-cell{color:var(--text3);cursor:grab;width:16px;font-size:13px;vertical-align:middle;}
.drag-handle-cell:active{cursor:grabbing;}
.num-inp{width:100%;font-size:12px;padding:2px 4px;border:1px solid transparent;border-radius:4px;background:transparent;color:var(--text);text-align:right;font-family:'DM Mono',monospace;}
.num-inp:hover{border-color:var(--border2);background:var(--bg);}
.num-inp:focus{outline:none;border-color:var(--blue);background:var(--surface);}
.text-inp{width:100%;font-size:12px;padding:2px 5px;border:1px solid transparent;border-radius:4px;background:transparent;color:var(--text);font-family:'DM Sans',sans-serif;}
.text-inp:hover{border-color:var(--border2);background:var(--bg);}
.text-inp:focus{outline:none;border-color:var(--blue);background:var(--surface);}
.desc-inp{font-size:11px;color:var(--text3);font-style:italic;}
.desc-inp::placeholder{color:var(--text3);}
.row-total{font-weight:600;text-align:right;font-family:'DM Mono',monospace;white-space:nowrap;}
select.inst-sel{font-size:11px;padding:2px 3px;border:1px solid var(--border2);border-radius:4px;background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;cursor:pointer;}
select.inst-sel:focus{outline:none;border-color:var(--blue);}
.subtotal-row td{background:var(--blue-total)!important;font-weight:700;border-top:1px solid var(--border2);}
.subtotal-row .mc{font-family:'DM Mono',monospace;text-align:right;}
.overhead-row td{background:#EBF5FF!important;font-size:11px;color:var(--blue);}
.overhead-row .mc{font-family:'DM Mono',monospace;text-align:right;}
.section-foot{display:flex;align-items:center;gap:6px;padding:7px 12px;border-top:1px solid var(--border);background:var(--surface2);position:relative;}
.role-menu{position:absolute;bottom:100%;left:0;background:var(--surface);border:1px solid var(--border2);border-radius:var(--radius-lg);box-shadow:0 4px 20px rgba(0,0,0,.15);min-width:290px;z-index:200;padding:4px 0;max-height:300px;overflow-y:auto;}
.role-menu-item{padding:7px 12px;cursor:pointer;font-size:12px;display:flex;align-items:center;justify-content:space-between;gap:8px;}
.role-menu-item:hover{background:var(--surface2);}
.role-menu-item .rn{font-weight:500;}
.role-menu-item .rr{font-family:'DM Mono',monospace;font-size:10px;color:var(--text3);}
.role-menu-hdr{padding:5px 12px 3px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--text3);}
.add-section-btn{display:flex;align-items:center;justify-content:center;gap:8px;padding:11px;border:2px dashed var(--border2);border-radius:var(--radius-xl);background:transparent;color:var(--text3);font-size:13px;font-weight:500;cursor:pointer;font-family:'DM Sans',sans-serif;transition:all .15s;width:100%;margin-bottom:1.25rem;}
.add-section-btn:hover{border-color:var(--blue);color:var(--blue);background:var(--blue-light);}
.cofin-panel{border:1px solid #9DC9AE;border-radius:var(--radius-xl);overflow:hidden;background:var(--surface);margin-bottom:1.25rem;}
.cofin-header{display:flex;align-items:center;gap:8px;padding:9px 12px;background:#E6F5ED;cursor:pointer;user-select:none;}
.cofin-title{font-size:13px;font-weight:600;display:flex;align-items:center;gap:7px;color:#14603A;flex:1;}
.cofin-total-badge{font-size:12px;font-weight:600;font-family:'DM Mono',monospace;color:#14603A;white-space:nowrap;margin-right:4px;}
.cofin-body{border-top:1px solid #9DC9AE;}
.cofin-desc-row{padding:7px 14px;border-bottom:1px solid var(--border);background:#F5FBF7;}
.cofin-panel th{background:#EAF6F0;color:#14603A;}
.cofin-panel th.yr-group{background:#C5E5D3;color:#14603A;border-left:2px solid #9DC9AE;border-right:2px solid #9DC9AE;}
.cofin-panel td.amt-cell{background:#F5FBF7;border-right:2px solid #C5E5D3;text-align:right;}
.cofin-subtotal td{background:#C5E5D3!important;font-weight:700;border-top:1px solid #9DC9AE;color:#14603A;}
.cofin-subtotal .mc{font-family:'DM Mono',monospace;text-align:right;}
.cofin-foot{display:flex;align-items:center;gap:6px;padding:7px 12px;border-top:1px solid #9DC9AE;background:#EAF6F0;}
.grand-footer{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;background:var(--surface);border:1px solid var(--border2);border-radius:var(--radius-xl);margin-top:1.25rem;}
.grand-breakdown{display:flex;flex-direction:column;gap:6px;min-width:280px;}
.gb-row{display:flex;align-items:baseline;justify-content:space-between;gap:24px;font-family:'DM Mono',monospace;}
.gb-label{font-family:'DM Sans',sans-serif;font-size:12px;color:var(--text2);font-weight:500;}
.gb-val{font-size:13px;font-weight:600;text-align:right;white-space:nowrap;}
.gb-row.gb-cofin .gb-label{color:#14603A;}
.gb-row.gb-cofin .gb-val{color:#14603A;}
.gb-divider{height:1px;background:var(--border2);margin:2px 0;}
.gb-row.gb-final .gb-label{font-size:14px;font-weight:700;color:var(--text);}
.gb-row.gb-final .gb-val{font-size:20px;font-weight:700;}
.grand-label{font-size:14px;font-weight:600;color:var(--text2);}
.grand-val{font-size:22px;font-weight:700;font-family:'DM Mono',monospace;text-align:right;}
.grand-sub{font-size:12px;color:var(--text3);font-family:'DM Mono',monospace;text-align:right;margin-top:2px;}
#dl-area{margin-bottom:8px;min-height:8px;}
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.4);display:flex;align-items:center;justify-content:center;z-index:1000;padding:1rem;}
.modal{background:var(--surface);border-radius:var(--radius-xl);padding:22px;max-width:420px;width:100%;box-shadow:0 8px 32px rgba(0,0,0,.18);}
.modal h3{font-size:15px;font-weight:700;margin-bottom:14px;}
.modal label{font-size:11px;color:var(--text2);font-weight:600;text-transform:uppercase;letter-spacing:.4px;display:block;margin-bottom:3px;}
.modal input,.modal textarea,.modal select{width:100%;font-size:13px;padding:7px 9px;border:1px solid var(--border2);border-radius:var(--radius);background:var(--bg);color:var(--text);font-family:'DM Sans',sans-serif;margin-bottom:10px;}
.modal textarea{height:60px;resize:vertical;}
.modal-actions{display:flex;gap:8px;justify-content:flex-end;margin-top:4px;}
.color-grid{display:flex;gap:5px;flex-wrap:wrap;margin-bottom:10px;}
.color-swatch{width:24px;height:24px;border-radius:5px;cursor:pointer;border:3px solid transparent;transition:transform .1s;}
.color-swatch:hover{transform:scale(1.15);}
.color-swatch.selected{border-color:#000;}
.toast{position:fixed;bottom:24px;right:24px;background:var(--text);color:var(--bg);padding:10px 18px;border-radius:var(--radius-lg);font-size:13px;font-weight:500;z-index:2000;animation:slideIn .2s ease;}
@keyframes slideIn{from{transform:translateY(20px);opacity:0;}to{transform:translateY(0);opacity:1;}}
</style>
</head>
<body>
<div class="app">
  <div class="app-header">
    <div>
      <div class="app-title">Forskningsbudget <span id="proj-title-display" style="font-weight:400;color:var(--text3)"></span></div>
      <small id="header-sub">Budgetværktøj til fondsansøgninger</small>
    </div>
    <div class="header-controls">
      <div class="lang-toggle">
        <button class="lang-btn active" id="btn-da" onclick="setLang('da')">DA</button>
        <button class="lang-btn" id="btn-en" onclick="setLang('en')">EN</button>
      </div>
      <div class="pill-group" id="year-btns">
        <button class="pill active" data-y="1">1</button><button class="pill" data-y="2">2</button>
        <button class="pill" data-y="3">3</button><button class="pill" data-y="4">4</button><button class="pill" data-y="5">5</button>
      </div>
      <div class="pill-group" id="inst-btns">
        <button class="pill active" data-inst="all" id="inst-all">Alle</button>
        <button class="pill" data-inst="Rigshospitalet">Rigshospitalet</button>
        <button class="pill" data-inst="DTU">DTU</button>
      </div>
      <div class="pill-group">
        <button class="pill active" id="btn-dkk" onclick="setCurrency('DKK')">DKK</button>
        <button class="pill" id="btn-eur" onclick="setCurrency('EUR')">EUR</button>
      </div>
      <button class="icon-btn success" onclick="exportPDF()"><i class="ti ti-file-type-pdf"></i> <span id="lbl-pdf">PDF</span></button>
      <button class="icon-btn" onclick="exportExcel()"><i class="ti ti-file-spreadsheet"></i> Excel</button>
      <button class="icon-btn" onclick="shareLink()"><i class="ti ti-link"></i> <span id="lbl-share">Del</span></button>
    </div>
  </div>

  <div id="dl-area"></div>
  <div id="share-bar-wrap"></div>

  <div class="config-bar">
    <label id="lbl-project">Projekttitel</label>
    <input type="text" id="proj-title" placeholder="Projekttitel…" style="width:200px" oninput="onProjTitle(this.value)"/>
    <div class="config-sep"></div>
    <label id="lbl-eurrate">1 EUR =</label>
    <input type="number" id="rate-inp" value="7.4728" step="0.0001" style="width:75px" oninput="eurRate=parseFloat(this.value)||7.4728;updateTotals()"/>
    <span style="font-size:11px;color:var(--text3)">DKK</span>
    <div class="config-sep"></div>
    <label id="lbl-overhead">Overhead (%)</label>
    <input type="number" id="overhead-inp" value="0" min="0" max="100" step="0.5" style="width:55px" oninput="overheadPct=parseFloat(this.value)||0;updateTotals()"/>
    <div class="config-sep"></div>
    <label id="lbl-inflation">Lønstigning (%/år)</label>
    <input type="number" id="inflation-inp" value="2" min="0" max="20" step="0.1" style="width:55px" oninput="inflationPct=parseFloat(this.value)||0;recalcAllSalary();updateTotals()"/>
    <span id="inflation-note" style="font-size:11px;color:var(--text3)">Anvendes fra År 2</span>
    <span id="overhead-result" style="font-size:12px;font-weight:600;color:var(--blue);font-family:'DM Mono',monospace;display:none;margin-left:8px;"></span>
  </div>

  <div class="panel" id="rates-panel">
    <div class="panel-header" onclick="togglePanel('rates')">
      <span class="panel-title"><i class="ti ti-coin"></i> <span id="lbl-ratesTitle">Lønsatser — DKK/år ved 1,0 FTE (År 1)</span></span>
      <i class="ti ti-chevron-down chevron open" id="rates-chev"></i>
    </div>
    <div id="rates-body">
      <div class="rates-grid" id="rates-grid"></div>
      <div style="padding:0 14px 10px;font-size:11px;color:var(--text3)" id="lbl-ratesHint">Budget = FTE × lønsats × (1+lønstigning)^(år-1). Redigér budgetfeltet manuelt for at låse.</div>
    </div>
  </div>

  <div class="summary-grid" id="summary-grid"></div>
  <div class="sections-wrap" id="sections-wrap"></div>

  <button class="add-section-btn" onclick="addSection()">
    <i class="ti ti-plus"></i> <span id="lbl-addCat">Tilføj kategori</span>
  </button>

  <!-- CO-FINANCING -->
  <div class="cofin-panel" id="cofin-panel"></div>

  <div class="grand-footer">
    <span class="grand-label" id="lbl-grandTotal">Budgetoversigt</span>
    <div class="grand-breakdown" id="grand-breakdown"></div>
  </div>
</div>
<script>
// ════════════════════════════════════════════════════
//  TRANSLATIONS
// ════════════════════════════════════════════════════
const T={
  da:{
    appSub:'Budgetværktøj til fondsansøgninger',
    project:'Projekttitel',eurRate:'1 EUR =',overhead:'Overhead (%)',
    inflation:'Lønstigning (%/år)',inflationNote:'Anvendes fra År 2',
    ratesTitle:'Lønsatser - DKK/år ved 1,0 FTE (År 1)',
    ratesHint:'Budget = FTE × lønsats × (1+lønstigning)^(år-1). Redigér budgetfeltet manuelt for at låse.',
    addCat:'Tilføj kategori',grandTotal:'Budgetoversigt',
    instAll:'Alle',pdfLbl:'PDF',shareLbl:'Del',
    cofinTitle:'Medfinansiering',
    cofinDesc:'Angiv midler, der allerede er bevilget eller tilsagt fra andre kilder.',
    cofinSource:'Finansieringskilde',cofinStatus:'Status',
    cofinAmount:'Beløb',cofinAddRow:'+ Tilføj finansieringskilde',
    cofinSubtotal:'Medfinansiering i alt',cofinNewSource:'Ny finansieringskilde',
    statusGranted:'Bevilget',statusPending:'Ansøgt',statusExpected:'Forventet',
    totalBudget:'Samlet budget',totalWithOh:'Budget inkl. overhead',
    lessCofin:'- Medfinansiering',amountSought:'Ansøgt beløb',
    cofinNone:'Ingen medfinansiering angivet',
    sub:'Underkategori',desc:'Beskrivelse',inst:'Institution',
    fte:'FTE',budget:'Budget',total:'Total',
    subtotal:'Subtotal',addPerson:'+ Tilføj person',addRow:'+ Tilføj post',
    overheadRow:'Overhead',inclOverhead:'Inkl. overhead',
    save:'Gem',cancel:'Annuller',editCat:'Rediger kategori',
    catName:'Kategorinavn',catDesc:'Beskrivelse (valgfri)',catColor:'Farve',
    hasFte:'Lønkategori (FTE-baseret)',
    delCatConfirm:'Slet denne kategori og alle dens poster?',
    chooseRole:'Vælg rolle',copyLink:'Kopiér link',linkCopied:'Link kopieret!',
    yr:y=>`År ${y}`,
    salCats:{
      consultant_rh:'Speciallæge',pregrad_rh:'Prægraduat stipendiat',
      scientist_rh:'Forsker',phd_rh:'Ph.d.-studerende',
      tech_rh:'Forskningssygeplejerske',projempl_rh:'Projektansat',
      projempl_dtu:'Projektansat',phd_dtu:'Ph.d.-studerende',
      postdoc_dtu:'Postdoc',pregrad_dtu:'Prægraduat stipendiat',
    },
    catLabels:{
      salary:'Løn',operation:'Drift',
      dissemination:'Formidling, uddannelse og træning',
      admin:'Administration',supplement:'Projekttillæg',
    },
    rowLabels:{
      data_mgmt:'Datahåndtering',subcontractor:'Underleverandøromkostninger',
      bench_fee:'Bench fee',infrastructure:'Infrastruktur',
      proj_specific:'Projektspecifikke omkostninger',operating:'Driftsomkostninger',
      equipment_dtu:'Udstyr',equipment_rh:'Udstyr',
      travel:'Rejseudgifter',training:'Uddannelse og kurser',
      conf_rh:'Konferencedeltagelse',conf_dtu:'Konferencedeltagelse',
      collab_rh:'Samarbejdsaktiviteter',collab_dtu:'Samarbejdsaktiviteter',
      publication:'Publikationsomkostninger',
      admin_direct:'Direkte administrative udgifter',
      supplement:'Projekttillæg',
    },
  },
  en:{
    appSub:'Budget tool for grant applications',
    project:'Project title',eurRate:'1 EUR =',overhead:'Overhead (%)',
    inflation:'Salary increase (%/yr)',inflationNote:'Applied from Year 2',
    ratesTitle:'Salary rates - DKK/yr at 1.0 FTE (Year 1)',
    ratesHint:'Budget = FTE × rate × (1+increase)^(year-1). Edit budget field manually to lock.',
    addCat:'Add category',grandTotal:'Budget summary',
    instAll:'All',pdfLbl:'PDF',shareLbl:'Share',
    cofinTitle:'Co-financing',
    cofinDesc:'Enter funds already granted or committed from other sources.',
    cofinSource:'Funding source',cofinStatus:'Status',
    cofinAmount:'Amount',cofinAddRow:'+ Add funding source',
    cofinSubtotal:'Total co-financing',cofinNewSource:'New funding source',
    statusGranted:'Granted',statusPending:'Applied',statusExpected:'Expected',
    totalBudget:'Total budget',totalWithOh:'Budget incl. overhead',
    lessCofin:'- Co-financing',amountSought:'Amount requested',
    cofinNone:'No co-financing entered',
    sub:'Subcategory',desc:'Description',inst:'Institution',
    fte:'FTE',budget:'Budget',total:'Total',
    subtotal:'Subtotal',addPerson:'+ Add person',addRow:'+ Add line',
    overheadRow:'Overhead',inclOverhead:'Incl. overhead',
    save:'Save',cancel:'Cancel',editCat:'Edit category',
    catName:'Category name',catDesc:'Description (optional)',catColor:'Colour',
    hasFte:'Salary category (FTE-based)',
    delCatConfirm:'Delete this category and all its lines?',
    chooseRole:'Choose role',copyLink:'Copy link',linkCopied:'Link copied!',
    yr:y=>`Year ${y}`,
    salCats:{
      consultant_rh:'Consultant, MD',pregrad_rh:'Pre-graduate scholar',
      scientist_rh:'Scientist / researcher',phd_rh:'PhD student',
      tech_rh:'Research nurse',projempl_rh:'Project employee',
      projempl_dtu:'Project employee',phd_dtu:'PhD student',
      postdoc_dtu:'Postdoc',pregrad_dtu:'Pre-graduate scholar',
    },
    catLabels:{
      salary:'Salary',operation:'Operation',
      dissemination:'Dissemination, training & education',
      admin:'Administration',supplement:'Project supplement',
    },
    rowLabels:{
      data_mgmt:'Data management',subcontractor:'Subcontractor costs',
      bench_fee:'Bench fee',infrastructure:'Infrastructure',
      proj_specific:'Project-specific costs',operating:'Operating expenses',
      equipment_dtu:'Equipment',equipment_rh:'Equipment',
      travel:'Travel expenses',training:'Training & courses',
      conf_rh:'Conference participation',conf_dtu:'Conference participation',
      collab_rh:'Collaborative activities',collab_dtu:'Collaborative activities',
      publication:'Publication costs',
      admin_direct:'Direct administrative expenses',
      supplement:'Project supplement',
    },
  }
};

// ════════════════════════════════════════════════════
//  STATE
// ════════════════════════════════════════════════════
let lang='da',currency='DKK',eurRate=7.4728,numYears=1,instFilter='all';
let overheadPct=0,inflationPct=2;

const RATE_DEFS=[
  {key:'consultant_rh',defaultInst:'Rigshospitalet',default:1080000},
  {key:'pregrad_rh',defaultInst:'Rigshospitalet',default:150000},
  {key:'scientist_rh',defaultInst:'Rigshospitalet',default:624000},
  {key:'phd_rh',defaultInst:'Rigshospitalet',default:624000},
  {key:'tech_rh',defaultInst:'Rigshospitalet',default:504000},
  {key:'projempl_rh',defaultInst:'Rigshospitalet',default:550000},
  {key:'projempl_dtu',defaultInst:'DTU',default:550000},
  {key:'phd_dtu',defaultInst:'DTU',default:530000},
  {key:'postdoc_dtu',defaultInst:'DTU',default:650000},
  {key:'pregrad_dtu',defaultInst:'DTU',default:150000},
];
let rates={};
RATE_DEFS.forEach(r=>{rates[r.key]=r.default;});

const CAT_COLORS=['#185FA5','#0F6E56','#854F0B','#5F5E5A','#534AB7','#7B2D8B','#C0392B','#1A7A4A','#B7791F','#2B6CB0'];
const INSTS=['Rigshospitalet','DTU'];

function uid(){return Math.random().toString(36).slice(2,9);}
function mkRow(lk,label,inst,rateKey=''){
  return {id:uid(),labelKey:lk,label,inst:inst||'Rigshospitalet',rateKey,fte:{},budget:{},budgetManual:{},desc:''};
}

let categories=[
  {id:'salary',labelKey:'salary',label:'Løn',color:'#185FA5',hasFte:true,desc:'',rows:[
    mkRow('consultant_rh','Speciallæge','Rigshospitalet','consultant_rh'),
    mkRow('pregrad_rh','Prægraduat stipendiat','Rigshospitalet','pregrad_rh'),
    mkRow('scientist_rh','Forsker','Rigshospitalet','scientist_rh'),
    mkRow('phd_rh','Ph.d.-studerende','Rigshospitalet','phd_rh'),
    mkRow('tech_rh','Forskningssygeplejerske','Rigshospitalet','tech_rh'),
    mkRow('projempl_rh','Projektansat','Rigshospitalet','projempl_rh'),
    mkRow('projempl_dtu','Projektansat','DTU','projempl_dtu'),
    mkRow('phd_dtu','Ph.d.-studerende','DTU','phd_dtu'),
    mkRow('postdoc_dtu','Postdoc','DTU','postdoc_dtu'),
    mkRow('pregrad_dtu','Prægraduat stipendiat','DTU','pregrad_dtu'),
  ]},
  {id:'operation',labelKey:'operation',label:'Drift',color:'#0F6E56',hasFte:false,desc:'',rows:[
    mkRow('data_mgmt','Datahåndtering','DTU'),
    mkRow('subcontractor','Underleverandøromkostninger','Rigshospitalet'),
    mkRow('bench_fee','Bench fee','Rigshospitalet'),
    mkRow('infrastructure','Infrastruktur','DTU'),
    mkRow('proj_specific','Projektspecifikke omkostninger','Rigshospitalet'),
    mkRow('operating','Driftsomkostninger','Rigshospitalet'),
    mkRow('equipment_dtu','Udstyr','DTU'),
    mkRow('equipment_rh','Udstyr','Rigshospitalet'),
  ]},
  {id:'dissemination',labelKey:'dissemination',label:'Formidling, uddannelse og træning',color:'#854F0B',hasFte:false,desc:'',rows:[
    mkRow('travel','Rejseudgifter','Rigshospitalet'),
    mkRow('training','Uddannelse og kurser','Rigshospitalet'),
    mkRow('conf_rh','Konferencedeltagelse','Rigshospitalet'),
    mkRow('conf_dtu','Konferencedeltagelse','DTU'),
    mkRow('collab_rh','Samarbejdsaktiviteter','Rigshospitalet'),
    mkRow('collab_dtu','Samarbejdsaktiviteter','DTU'),
    mkRow('publication','Publikationsomkostninger','Rigshospitalet'),
  ]},
  {id:'admin',labelKey:'admin',label:'Administration',color:'#5F5E5A',hasFte:false,desc:'',rows:[
    mkRow('admin_direct','Direkte administrative udgifter','Rigshospitalet'),
  ]},
  {id:'supplement',labelKey:'supplement',label:'Projekttillæg',color:'#534AB7',hasFte:false,desc:'',rows:[
    mkRow('supplement','Projekttillæg','DTU'),
  ]},
];
let collapsed={};

// ── CO-FINANCING state ──
// Each entry: {id, source, status, amounts:{year:amount}}
let coFin=[];
let coFinCollapsed=false;
const COFIN_STATUS=['granted','pending','expected'];
function mkCoFin(source,status){
  return {id:uid(),source:source||'',status:status||'granted',amounts:{}};
}

// ════════════════════════════════════════════════════
//  INFLATION: compute adjusted salary for year y
// ════════════════════════════════════════════════════
function inflFactor(y){
  return Math.pow(1+(inflationPct||0)/100, y-1);
}
function calcSalaryBudget(row,y){
  const fte=row.fte[y]||0;
  const baseRate=rates[row.rateKey]||0;
  return Math.round(fte*baseRate*inflFactor(y));
}

// ════════════════════════════════════════════════════
//  FORMAT
// ════════════════════════════════════════════════════
function toDKK(n){return Math.round(n);}
function toDisplay(n){return currency==='EUR'?n/eurRate:n;}
function fmtCur(n){
  const v=Math.round(toDisplay(n));
  const s=Math.abs(v).toLocaleString('da-DK');
  return currency==='EUR'?`€\u00a0${s}`:`${s}\u00a0kr.`;
}
function fmtDKKraw(n){return Math.round(n).toLocaleString('da-DK')+'\u00a0kr.';}
function yrs(){return Array.from({length:numYears},(_,i)=>i+1);}
function esc(s){return (s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;');}

// ════════════════════════════════════════════════════
//  LABEL RESOLUTION
// ════════════════════════════════════════════════════
function rowLabel(row){
  if(row.rateKey&&T[lang].salCats[row.rateKey]){
    const da=T.da.salCats[row.rateKey],en=T.en.salCats[row.rateKey];
    if(row.label===da||row.label===en) return T[lang].salCats[row.rateKey];
  }
  if(row.labelKey&&T[lang].rowLabels[row.labelKey]){
    const da=T.da.rowLabels[row.labelKey],en=T.en.rowLabels[row.labelKey];
    if(row.label===da||row.label===en) return T[lang].rowLabels[row.labelKey];
  }
  return row.label;
}
function catLabel(cat){
  if(cat.labelKey&&T[lang].catLabels[cat.labelKey]){
    const da=T.da.catLabels[cat.labelKey],en=T.en.catLabels[cat.labelKey];
    if(cat.label===da||cat.label===en) return T[lang].catLabels[cat.labelKey];
  }
  return cat.label;
}

// ════════════════════════════════════════════════════
//  TOTALS
// ════════════════════════════════════════════════════
function rowBudget(row,y){return row.budget[y]||0;}
function catTotal(cat){
  let t=0;
  cat.rows.forEach(row=>{
    if(instFilter!=='all'&&row.inst!==instFilter)return;
    yrs().forEach(y=>{t+=rowBudget(row,y);});
  });
  return t;
}
function yearGrand(y){
  let t=0;
  categories.forEach(cat=>cat.rows.forEach(row=>{
    if(instFilter!=='all'&&row.inst!==instFilter)return;
    t+=rowBudget(row,y);
  }));
  return t;
}
function grandTotal(){let t=0;yrs().forEach(y=>{t+=yearGrand(y);});return t;}

// ── CO-FINANCING calculations (all in DKK) ──
function coFinYear(y){
  let t=0;
  coFin.forEach(cf=>{t+=cf.amounts[y]||0;});
  return t;
}
function coFinTotal(){
  let t=0;yrs().forEach(y=>{t+=coFinYear(y);});return t;
}
function coFinRowTotal(cf){
  let t=0;yrs().forEach(y=>{t+=cf.amounts[y]||0;});return t;
}
// Budget including overhead
function budgetInclOverhead(){
  const gt=grandTotal();
  return gt+(overheadPct>0?Math.round(gt*overheadPct/100):0);
}
function yearInclOverhead(y){
  const yt=yearGrand(y);
  return yt+(overheadPct>0?Math.round(yt*overheadPct/100):0);
}
// Amount to request from funder = budget (incl overhead) - co-financing
function amountSought(){
  return budgetInclOverhead()-coFinTotal();
}
function yearSought(y){
  return yearInclOverhead(y)-coFinYear(y);
}

// ════════════════════════════════════════════════════
//  RECALC ALL SALARY (when inflation/rates change)
// ════════════════════════════════════════════════════
function recalcAllSalary(){
  categories.forEach(cat=>{
    if(!cat.hasFte)return;
    cat.rows.forEach(row=>{
      if(!row.rateKey)return;
      yrs().forEach(y=>{
        if(row.budgetManual[y])return;
        row.budget[y]=calcSalaryBudget(row,y);
        const inp=document.getElementById(`bi_${row.id}_${y}`);
        if(inp)inp.value=row.budget[y]||'';
      });
    });
  });
}
function recalcSalaryKey(rateKey){
  categories.forEach(cat=>{
    if(!cat.hasFte)return;
    cat.rows.forEach(row=>{
      if(row.rateKey!==rateKey)return;
      yrs().forEach(y=>{
        if(row.budgetManual[y])return;
        row.budget[y]=calcSalaryBudget(row,y);
        const inp=document.getElementById(`bi_${row.id}_${y}`);
        if(inp)inp.value=row.budget[y]||'';
      });
    });
  });
}

// ════════════════════════════════════════════════════
//  UPDATE TOTALS (no re-render)
// ════════════════════════════════════════════════════
function updateTotals(){
  const gt=grandTotal();
  const oh=overheadPct>0?Math.round(gt*overheadPct/100):0;
  const cf=coFinTotal();
  const sought=amountSought();

  // ── Grand footer breakdown ──
  const gb=document.getElementById('grand-breakdown');
  if(gb){
    let h='';
    h+=`<div class="gb-row"><span class="gb-label">${T[lang].totalBudget}</span><span class="gb-val">${fmtCur(gt)}</span></div>`;
    if(overheadPct>0){
      h+=`<div class="gb-row"><span class="gb-label">${T[lang].overheadRow} (${overheadPct}%)</span><span class="gb-val">${fmtCur(oh)}</span></div>`;
      h+=`<div class="gb-row"><span class="gb-label">${T[lang].totalWithOh}</span><span class="gb-val">${fmtCur(gt+oh)}</span></div>`;
    }
    if(cf>0){
      h+=`<div class="gb-row gb-cofin"><span class="gb-label">${T[lang].lessCofin}</span><span class="gb-val">-${fmtCur(cf)}</span></div>`;
    }
    h+=`<div class="gb-divider"></div>`;
    h+=`<div class="gb-row gb-final"><span class="gb-label">${cf>0?T[lang].amountSought:T[lang].totalBudget}</span><span class="gb-val">${fmtCur(sought)}</span></div>`;
    if(currency==='DKK'){
      h+=`<div class="gb-row"><span class="gb-label" style="font-size:11px;color:var(--text3)">≈ EUR</span><span class="gb-val" style="font-size:11px;color:var(--text3);font-weight:400">€ ${Math.round(sought/eurRate).toLocaleString('da-DK')}</span></div>`;
    }
    gb.innerHTML=h;
  }

  if(overheadPct>0){
    document.getElementById('overhead-result').style.display='inline';
    document.getElementById('overhead-result').textContent=T[lang].overheadRow+': '+fmtCur(oh);
  } else {
    document.getElementById('overhead-result').style.display='none';
  }

  // ── Co-financing live cells ──
  coFin.forEach(cfRow=>{
    const rt=document.getElementById('cft_'+cfRow.id);
    if(rt) rt.textContent=fmtCur(coFinRowTotal(cfRow));
  });
  yrs().forEach(y=>{
    const el=document.getElementById('cfsub_'+y);
    if(el) el.textContent=fmtCur(coFinYear(y));
  });
  const cfTot=document.getElementById('cfsub_tot');
  if(cfTot) cfTot.textContent=fmtCur(cf);
  const cfBadge=document.getElementById('cofin-total-badge');
  if(cfBadge) cfBadge.textContent=fmtCur(cf);

  // Summary cards — show year budget and, if co-financing exists, net sought
  const sg=document.getElementById('summary-grid');sg.innerHTML='';
  yrs().forEach(y=>{
    const yt=yearGrand(y);
    const yc=coFinYear(y);
    const m=document.createElement('div');m.className='metric';
    let sub='';
    if(yc>0){
      sub=`<div class="metric-infl" style="color:#14603A">-${fmtCur(yc)} ${lang==='da'?'medfin.':'co-fin.'}</div>`;
    } else if(y>1&&inflationPct>0){
      sub=`<div class="metric-infl">+${inflationPct}%/${lang==='da'?'år':'yr'}</div>`;
    }
    m.innerHTML=`<div class="metric-label">${T[lang].yr(y)}</div><div class="metric-value">${fmtCur(yt)}</div>${sub}`;
    sg.appendChild(m);
  });

  // Per-category
  categories.forEach(cat=>{
    const ct=catTotal(cat);
    const el=document.getElementById('sec-total-'+cat.id);
    if(el) el.textContent=fmtCur(ct);
    cat.rows.forEach(row=>{
      let rowT=0;yrs().forEach(y=>{rowT+=rowBudget(row,y);});
      const rt=document.getElementById('rt_'+row.id);
      if(rt) rt.textContent=fmtCur(rowT);
    });
    yrs().forEach(y=>{
      let yt=0;
      cat.rows.forEach(row=>{
        if(instFilter!=='all'&&row.inst!==instFilter)return;
        yt+=rowBudget(row,y);
      });
      const sc=document.getElementById(`sc_${cat.id}_${y}`);if(sc)sc.textContent=fmtCur(yt);
      const oc=document.getElementById(`oc_${cat.id}_${y}`);if(oc&&overheadPct>0)oc.textContent=fmtCur(Math.round(yt*overheadPct/100));
    });
    const stot=document.getElementById('sc_'+cat.id+'_tot');if(stot)stot.textContent=fmtCur(ct);
    const otot=document.getElementById('oc_'+cat.id+'_tot');if(otot)otot.textContent=fmtCur(Math.round(ct*overheadPct/100));
    const ohRow=document.getElementById('oh_'+cat.id);
    if(ohRow)ohRow.style.display=overheadPct>0?'':'none';
  });

  document.querySelectorAll('.budget-th').forEach(th=>{
    th.textContent=`${T[lang].budget} (${currency})`;
  });
  document.querySelectorAll('.fte-th').forEach(th=>{th.textContent=T[lang].fte;});
  document.querySelectorAll('.yr-group-th').forEach(th=>{
    th.textContent=T[lang].yr(th.dataset.y);
  });
}

// ════════════════════════════════════════════════════
//  LANGUAGE
// ════════════════════════════════════════════════════
function setLang(l){
  lang=l;
  document.getElementById('btn-da').classList.toggle('active',l==='da');
  document.getElementById('btn-en').classList.toggle('active',l==='en');
  const m={
    'header-sub':'appSub','lbl-project':'project','lbl-eurrate':'eurRate',
    'lbl-overhead':'overhead','lbl-inflation':'inflation','inflation-note':'inflationNote',
    'lbl-ratesTitle':'ratesTitle','lbl-ratesHint':'ratesHint',
    'lbl-addCat':'addCat','lbl-grandTotal':'grandTotal',
    'lbl-pdf':'pdfLbl','lbl-share':'shareLbl',
  };
  Object.entries(m).forEach(([id,k])=>{
    const el=document.getElementById(id);if(el&&T[lang][k])el.textContent=T[lang][k];
  });
  document.getElementById('inst-all').textContent=T[lang].instAll;
  document.title=l==='da'?'Forskningsbudget':'Research Budget';
  buildRatesPanel();renderAll();  // renderAll also rebuilds co-financing panel
}

function setCurrency(c){
  currency=c;
  document.getElementById('btn-dkk').classList.toggle('active',c==='DKK');
  document.getElementById('btn-eur').classList.toggle('active',c==='EUR');
  updateTotals();
}

function onProjTitle(v){
  document.getElementById('proj-title-display').textContent=v?'— '+v:'';
}

// ════════════════════════════════════════════════════
//  PANELS
// ════════════════════════════════════════════════════
function togglePanel(id){
  const body=document.getElementById(id+'-body');
  const chev=document.getElementById(id+'-chev');
  if(!body)return;
  const open=body.style.display!=='none';
  body.style.display=open?'none':'block';
  chev.classList.toggle('open',!open);
}

// ════════════════════════════════════════════════════
//  SALARY RATES PANEL
// ════════════════════════════════════════════════════
function buildRatesPanel(){
  const grid=document.getElementById('rates-grid');grid.innerHTML='';
  RATE_DEFS.forEach(r=>{
    const lbl=T[lang].salCats[r.key]||r.key;
    const div=document.createElement('div');div.className='rate-row';
    div.innerHTML=`<span class="rate-label">${lbl} <span style="color:var(--text3);font-weight:400">(${r.defaultInst})</span></span><input class="rate-inp" type="number" min="0" step="1000" value="${rates[r.key]}" data-key="${r.key}"/>`;
    grid.appendChild(div);
  });
}
document.getElementById('rates-grid').addEventListener('input',e=>{
  const el=e.target;if(!el.dataset.key)return;
  rates[el.dataset.key]=parseFloat(el.value)||0;
  recalcSalaryKey(el.dataset.key);
  updateTotals();
});

// ════════════════════════════════════════════════════
//  CO-FINANCING PANEL
// ════════════════════════════════════════════════════
function statusLabel(s){
  return s==='pending'?T[lang].statusPending
       : s==='expected'?T[lang].statusExpected
       : T[lang].statusGranted;
}

function buildCoFinPanel(){
  const panel=document.getElementById('cofin-panel');
  if(!panel) return;
  const yr=yrs();
  const open=!coFinCollapsed;
  const cfTotal=coFinTotal();

  let h=`<div class="cofin-header" onclick="toggleCoFin()">
    <i class="ti ti-coins" style="font-size:16px;color:#14603A"></i>
    <span class="cofin-title">${T[lang].cofinTitle}</span>
    <span class="cofin-total-badge" id="cofin-total-badge">${fmtCur(cfTotal)}</span>
    <i class="ti ti-chevron-down chevron ${open?'open':''}" id="cofin-chev" style="color:#14603A"></i>
  </div>
  <div class="cofin-body" id="cofin-body" style="display:${open?'block':'none'}">
    <div class="cofin-desc-row" style="font-size:11px;color:var(--text2);font-style:italic">
      ${T[lang].cofinDesc}
    </div>`;

  if(coFin.length===0){
    h+=`<div style="padding:20px;text-align:center;color:var(--text3);font-size:12px">
      ${T[lang].cofinNone}
    </div>`;
  } else {
    h+=`<div class="tbl-wrap"><table>
      <thead>
        <tr>
          <th style="width:16px"></th>
          <th style="min-width:180px">${T[lang].cofinSource}</th>
          <th style="width:110px">${T[lang].cofinStatus}</th>`;
    yr.forEach(y=>{
      h+=`<th class="r yr-group" style="min-width:110px">${T[lang].yr(y)}</th>`;
    });
    h+=`<th class="r" style="min-width:100px">${T[lang].total}</th>
          <th style="width:32px"></th>
        </tr>
      </thead>
      <tbody id="cofin-tbody">`;

    coFin.forEach(cf=>{
      h+=`<tr class="data-row" data-cofin-id="${cf.id}">
        <td class="drag-handle-cell"><i class="ti ti-grip-horizontal"></i></td>
        <td><input class="text-inp" type="text" value="${esc(cf.source)}"
             placeholder="${T[lang].cofinNewSource}"
             onchange="onCoFinSource('${cf.id}',this.value)"/></td>
        <td>
          <select class="inst-sel" onchange="onCoFinStatus('${cf.id}',this.value)">
            ${COFIN_STATUS.map(s=>`<option value="${s}"${cf.status===s?' selected':''}>${statusLabel(s)}</option>`).join('')}
          </select>
        </td>`;
      yr.forEach(y=>{
        h+=`<td class="amt-cell"><input id="cfa_${cf.id}_${y}" class="num-inp" type="number"
             min="0" step="1000" placeholder="0" value="${cf.amounts[y]||''}"
             onchange="onCoFinAmount('${cf.id}',${y},this.value)"/></td>`;
      });
      h+=`<td class="row-total" id="cft_${cf.id}">${fmtCur(coFinRowTotal(cf))}</td>
        <td style="text-align:center">
          <button class="icon-btn danger" style="padding:3px 6px"
            onclick="delCoFin('${cf.id}')" title="Slet"><i class="ti ti-trash"></i></button>
        </td>
      </tr>`;
    });

    // Subtotal row
    h+=`<tr class="cofin-subtotal">
      <td colspan="3" style="font-weight:700;font-family:'DM Sans',sans-serif;padding-left:10px">
        ${T[lang].cofinSubtotal}
      </td>`;
    yr.forEach(y=>{
      h+=`<td class="mc" id="cfsub_${y}" style="text-align:right;padding-right:7px"></td>`;
    });
    h+=`<td class="mc" id="cfsub_tot" style="text-align:right;padding-right:7px"></td><td></td></tr>`;
    h+=`</tbody></table></div>`;
  }

  h+=`<div class="cofin-foot">
    <button class="icon-btn ghost" style="font-size:11px;color:#14603A" onclick="addCoFin()">
      <i class="ti ti-plus"></i> ${T[lang].cofinAddRow}
    </button>
  </div></div>`;

  panel.innerHTML=h;

  // Sortable co-financing rows
  const tb=document.getElementById('cofin-tbody');
  if(tb){
    Sortable.create(tb,{handle:'.drag-handle-cell',animation:100,
      onEnd(){
        const ids=[...tb.querySelectorAll('tr.data-row')].map(tr=>tr.dataset.cofinId);
        coFin.sort((a,b)=>ids.indexOf(a.id)-ids.indexOf(b.id));
        updateTotals();
      }
    });
  }
}

window.toggleCoFin=function(){
  coFinCollapsed=!coFinCollapsed;
  const body=document.getElementById('cofin-body');
  const chev=document.getElementById('cofin-chev');
  if(body) body.style.display=coFinCollapsed?'none':'block';
  if(chev) chev.classList.toggle('open',!coFinCollapsed);
};
window.addCoFin=function(e){
  if(e) e.stopPropagation();
  coFin.push(mkCoFin('','granted'));
  buildCoFinPanel();updateTotals();
};
window.delCoFin=function(id){
  coFin=coFin.filter(cf=>cf.id!==id);
  buildCoFinPanel();updateTotals();
};
window.onCoFinSource=function(id,val){
  const cf=coFin.find(c=>c.id===id);if(cf)cf.source=val;
};
window.onCoFinStatus=function(id,val){
  const cf=coFin.find(c=>c.id===id);if(cf)cf.status=val;
};
window.onCoFinAmount=function(id,y,val){
  const cf=coFin.find(c=>c.id===id);
  if(cf){cf.amounts[y]=parseFloat(val)||0;}
  updateTotals();
};

// ════════════════════════════════════════════════════
//  RENDER ALL
// ════════════════════════════════════════════════════
function renderAll(){
  const wrap=document.getElementById('sections-wrap');
  wrap.innerHTML='';
  categories.forEach(cat=>wrap.appendChild(buildSection(cat)));
  initSortableSections();
  buildCoFinPanel();
  updateTotals();
}

// ════════════════════════════════════════════════════
//  BUILD SECTION
// ════════════════════════════════════════════════════
function buildSection(cat){
  const isOpen=collapsed[cat.id]!==true;
  const sec=document.createElement('div');
  sec.className='section';sec.dataset.catId=cat.id;

  const hdr=document.createElement('div');hdr.className='section-header';
  hdr.innerHTML=`
    <i class="ti ti-grip-vertical section-drag-handle"></i>
    <div class="section-dot" style="background:${cat.color}"></div>
    <div class="section-name-wrap">
      <input class="section-name-inp" type="text" value="${esc(catLabel(cat))}"
        onclick="event.stopPropagation()" onchange="renameCat('${cat.id}',this.value)"/>
    </div>
    <div class="section-controls">
      <span class="section-total-badge" id="sec-total-${cat.id}">${fmtCur(catTotal(cat))}</span>
      <button class="icon-btn ghost" style="padding:4px 7px" onclick="event.stopPropagation();editSection('${cat.id}')" title="${T[lang].editCat}"><i class="ti ti-settings-2"></i></button>
      <button class="icon-btn danger" style="padding:4px 7px" onclick="event.stopPropagation();delSection('${cat.id}')" title="Slet"><i class="ti ti-trash"></i></button>
      <i class="ti ti-chevron-down chevron ${isOpen?'open':''}" id="chev-${cat.id}"></i>
    </div>`;
  hdr.addEventListener('click',()=>toggleSec(cat.id));
  sec.appendChild(hdr);

  const body=document.createElement('div');
  body.className='section-body';body.id='sb-'+cat.id;
  body.style.display=isOpen?'block':'none';

  const descRow=document.createElement('div');descRow.className='section-desc-row';
  descRow.innerHTML=`<input class="text-inp desc-inp" type="text" placeholder="${T[lang].catDesc}" value="${esc(cat.desc)}" style="width:100%"/>`;
  descRow.querySelector('input').addEventListener('change',e=>{cat.desc=e.target.value;});
  body.appendChild(descRow);
  body.appendChild(buildTable(cat));

  const foot=document.createElement('div');foot.className='section-foot';
  if(cat.hasFte){
    const apBtn=document.createElement('button');
    apBtn.className='icon-btn ghost';apBtn.style.fontSize='11px';
    apBtn.innerHTML=`<i class="ti ti-user-plus"></i> ${T[lang].addPerson}`;
    apBtn.addEventListener('click',e=>{e.stopPropagation();showRoleMenu(cat.id,apBtn);});
    foot.appendChild(apBtn);
  }
  const arBtn=document.createElement('button');
  arBtn.className='icon-btn ghost';arBtn.style.fontSize='11px';
  arBtn.innerHTML=`<i class="ti ti-plus"></i> ${T[lang].addRow}`;
  arBtn.addEventListener('click',()=>addRow(cat.id));
  foot.appendChild(arBtn);
  body.appendChild(foot);
  sec.appendChild(body);
  return sec;
}

function toggleSec(catId){
  const isOpen=collapsed[catId]!==true;
  collapsed[catId]=isOpen;
  const body=document.getElementById('sb-'+catId);
  const chev=document.getElementById('chev-'+catId);
  if(body)body.style.display=isOpen?'none':'block';
  if(chev)chev.classList.toggle('open',!isOpen);
}

// ════════════════════════════════════════════════════
//  ROLE PICKER
// ════════════════════════════════════════════════════
function showRoleMenu(catId,anchor){
  document.querySelectorAll('.role-menu').forEach(m=>m.remove());
  const menu=document.createElement('div');menu.className='role-menu';
  menu.innerHTML=`<div class="role-menu-hdr">${T[lang].chooseRole}</div>`;
  RATE_DEFS.forEach(r=>{
    const lbl=T[lang].salCats[r.key]||r.key;
    const rateStr=rates[r.key].toLocaleString('da-DK')+' kr./år';
    const item=document.createElement('div');item.className='role-menu-item';
    item.innerHTML=`<span class="rn">${lbl}</span><span style="display:flex;align-items:center;gap:5px"><span style="font-size:10px;padding:1px 5px;border-radius:4px;background:var(--surface2);color:var(--text3)">${r.defaultInst}</span><span class="rr">${rateStr}</span></span>`;
    item.addEventListener('click',()=>{menu.remove();addPersonWithRole(catId,r.key,lbl,r.defaultInst);});
    menu.appendChild(item);
  });
  const div=document.createElement('div');div.style.cssText='height:1px;background:var(--border);margin:3px 0;';
  menu.appendChild(div);
  const custom=document.createElement('div');custom.className='role-menu-item';
  custom.innerHTML=`<span class="rn" style="color:var(--text2)">${lang==='da'?'Tilpas manuelt…':'Custom…'}</span>`;
  custom.addEventListener('click',()=>{menu.remove();addRow(catId);});
  menu.appendChild(custom);
  anchor.closest('.section-foot').appendChild(menu);
  setTimeout(()=>{
    document.addEventListener('click',function h(e){if(!menu.contains(e.target)){menu.remove();document.removeEventListener('click',h);}});
  },10);
}

function addPersonWithRole(catId,rateKey,label,inst){
  const cat=categories.find(c=>c.id===catId);
  const row=mkRow(rateKey,label,inst,rateKey);
  // Pre-calc year 1 budget if FTE not set yet (set to 0)
  cat.rows.push(row);
  renderAll();
}

// ════════════════════════════════════════════════════
//  TABLE with paired FTE+Budget per year
// ════════════════════════════════════════════════════
function buildTable(cat){
  const wrap=document.createElement('div');wrap.className='tbl-wrap';
  const yr=yrs();

  let h=`<table><thead>`;
  // Row 1: year group spans
  h+=`<tr>
    <th style="width:16px;background:var(--surface2);border:none"></th>
    <th style="min-width:140px;background:var(--surface2);border:none"></th>
    <th style="min-width:120px;background:var(--surface2);border:none"></th>
    <th style="width:78px;background:var(--surface2);border:none"></th>`;
  yr.forEach(y=>{
    const span=cat.hasFte?2:1;
    h+=`<th colspan="${span}" class="yr-group yr-group-th" data-y="${y}">${T[lang].yr(y)}</th>`;
  });
  h+=`<th style="background:var(--surface2);border:none"></th></tr>`;

  // Row 2: sub-headers
  h+=`<tr>
    <th style="width:16px"></th>
    <th>${T[lang].sub}</th>
    <th>${T[lang].desc}</th>
    <th>${T[lang].inst}</th>`;
  yr.forEach(y=>{
    if(cat.hasFte){
      h+=`<th class="r fte-hdr fte-th" style="width:55px">${T[lang].fte}</th>`;
      h+=`<th class="r bud-hdr budget-th" data-y="${y}" style="min-width:105px">${T[lang].budget} (${currency})</th>`;
    } else {
      h+=`<th class="r bud-hdr budget-th" data-y="${y}" style="min-width:105px">${T[lang].budget} (${currency})</th>`;
    }
  });
  h+=`<th class="r" style="min-width:95px">${T[lang].total}</th></tr></thead>`;

  h+=`<tbody id="tbody-${cat.id}">`;
  cat.rows.forEach(row=>{h+=buildRowHTML(cat,row);});

  // Subtotal — colspan covers: drag(1) + sub(1) + desc(1) + inst(1) = 4
  h+=`<tr class="subtotal-row">
    <td colspan="4" style="font-weight:700;font-family:'DM Sans',sans-serif;padding-left:10px">${T[lang].subtotal} — ${catLabel(cat)}</td>`;
  yr.forEach(y=>{
    if(cat.hasFte) h+=`<td class="mc"></td>`; // empty FTE cell
    h+=`<td class="mc" id="sc_${cat.id}_${y}" style="text-align:right;padding-right:7px"></td>`;
  });
  h+=`<td class="mc" id="sc_${cat.id}_tot" style="text-align:right;padding-right:7px"></td></tr>`;

  // Overhead row — same structure
  h+=`<tr class="overhead-row" id="oh_${cat.id}" style="${overheadPct>0?'':'display:none'}">
    <td colspan="4" style="padding-left:10px">${T[lang].overheadRow}${overheadPct>0?` (${overheadPct}%)`:''}`;
  h+=`</td>`;
  yr.forEach(y=>{
    if(cat.hasFte) h+=`<td class="mc"></td>`;
    h+=`<td class="mc" id="oc_${cat.id}_${y}" style="text-align:right;padding-right:7px"></td>`;
  });
  h+=`<td class="mc" id="oc_${cat.id}_tot" style="text-align:right;padding-right:7px"></td></tr>`;

  h+=`</tbody></table>`;
  wrap.innerHTML=h;

  const tbody=wrap.querySelector('#tbody-'+cat.id);
  if(tbody){
    Sortable.create(tbody,{handle:'.drag-handle-cell',animation:100,
      onEnd(){
        const ids=[...tbody.querySelectorAll('tr.data-row')].map(tr=>tr.dataset.rowId);
        cat.rows.sort((a,b)=>ids.indexOf(a.id)-ids.indexOf(b.id));
        updateTotals();
      }
    });
  }
  return wrap;
}

function buildRowHTML(cat,row){
  const yr=yrs();
  const hide=instFilter!=='all'&&row.inst!==instFilter;
  let h=`<tr class="data-row${hide?' hidden-row':''}" data-row-id="${row.id}">
    <td class="drag-handle-cell"><i class="ti ti-grip-horizontal"></i></td>
    <td><input class="text-inp" type="text" value="${esc(rowLabel(row))}" onchange="renameRow('${cat.id}','${row.id}',this.value)"/></td>
    <td><input class="text-inp desc-inp" type="text" placeholder="${T[lang].desc}…" value="${esc(row.desc)}" onchange="onRowDesc('${cat.id}','${row.id}',this.value)"/></td>
    <td><select class="inst-sel" onchange="onInstChange('${cat.id}','${row.id}',this.value)">
      ${INSTS.map(i=>`<option value="${i}"${row.inst===i?' selected':''}>${i}</option>`).join('')}
    </select></td>`;
  yr.forEach(y=>{
    if(cat.hasFte){
      h+=`<td class="fte-cell"><input id="fi_${row.id}_${y}" class="num-inp" type="number" min="0" max="10" step="0.1" placeholder="0" value="${row.fte[y]||''}" style="width:50px" onchange="onFte('${cat.id}','${row.id}',${y},this.value)"/></td>`;
    }
    h+=`<td class="bud-cell"><input id="bi_${row.id}_${y}" class="num-inp" type="number" min="0" step="1000" placeholder="0" value="${row.budget[y]||''}" onchange="onBudget('${cat.id}','${row.id}',${y},this.value)"/></td>`;
  });
  let rowT=0;yr.forEach(y=>{rowT+=rowBudget(row,y);});
  h+=`<td class="row-total" id="rt_${row.id}">${fmtCur(rowT)}</td></tr>`;
  return h;
}

// ════════════════════════════════════════════════════
//  INPUT HANDLERS
// ════════════════════════════════════════════════════
window.onFte=function(catId,rowId,y,val){
  const cat=categories.find(c=>c.id===catId);
  const row=cat.rows.find(r=>r.id===rowId);
  row.fte[y]=parseFloat(val)||0;
  if(!row.budgetManual[y]){
    row.budget[y]=calcSalaryBudget(row,y);
    const bi=document.getElementById(`bi_${rowId}_${y}`);
    if(bi)bi.value=row.budget[y]||'';
  }
  updateTotals();
};
window.onBudget=function(catId,rowId,y,val){
  const cat=categories.find(c=>c.id===catId);
  const row=cat.rows.find(r=>r.id===rowId);
  const v=parseFloat(val)||0;
  if(cat.hasFte&&row.rateKey){
    const auto=calcSalaryBudget(row,y);
    row.budgetManual[y]=v!==auto;
  }
  row.budget[y]=v;
  updateTotals();
};
window.onRowDesc=function(catId,rowId,val){
  categories.find(c=>c.id===catId).rows.find(r=>r.id===rowId).desc=val;
};
window.onInstChange=function(catId,rowId,val){
  const cat=categories.find(c=>c.id===catId);
  cat.rows.find(r=>r.id===rowId).inst=val;
  applyInstFilter();
};
window.renameRow=function(catId,rowId,val){
  categories.find(c=>c.id===catId).rows.find(r=>r.id===rowId).label=val;
};
window.renameCat=function(catId,val){
  const cat=categories.find(c=>c.id===catId);cat.label=val;cat.labelKey='';updateTotals();
};

// ════════════════════════════════════════════════════
//  SECTION ACTIONS
// ════════════════════════════════════════════════════
function addRow(catId){
  const cat=categories.find(c=>c.id===catId);
  cat.rows.push(mkRow('',lang==='da'?'Ny post':'New line','Rigshospitalet'));
  renderAll();
}
window.delSection=function(catId){
  if(!confirm(T[lang].delCatConfirm))return;
  categories=categories.filter(c=>c.id!==catId);renderAll();
};
window.addSection=function(){
  const color=CAT_COLORS[categories.length%CAT_COLORS.length];
  categories.push({id:uid(),labelKey:'',label:lang==='da'?'Ny kategori':'New category',
    color,hasFte:false,desc:'',rows:[mkRow('',lang==='da'?'Ny post':'New line','Rigshospitalet')]});
  renderAll();
};

// ════════════════════════════════════════════════════
//  EDIT SECTION MODAL
// ════════════════════════════════════════════════════
window.editSection=function(catId){
  const cat=categories.find(c=>c.id===catId);
  const ov=document.createElement('div');ov.className='modal-overlay';
  ov.innerHTML=`<div class="modal">
    <h3>${T[lang].editCat}</h3>
    <label>${T[lang].catName}</label><input id="m-name" value="${esc(catLabel(cat))}"/>
    <label>${T[lang].catDesc}</label><textarea id="m-desc">${esc(cat.desc)}</textarea>
    <label>${T[lang].catColor}</label>
    <div class="color-grid">${CAT_COLORS.map(c=>`<div class="color-swatch${c===cat.color?' selected':''}" style="background:${c}" data-c="${c}" onclick="pickColor(this)"></div>`).join('')}</div>
    <label style="display:flex;align-items:center;gap:7px;margin-bottom:12px;text-transform:none;letter-spacing:0;font-size:13px">
      <input id="m-fte" type="checkbox" ${cat.hasFte?'checked':''} style="width:auto;margin:0"/>
      <span style="font-weight:400">${T[lang].hasFte}</span>
    </label>
    <div class="modal-actions">
      <button class="icon-btn" onclick="this.closest('.modal-overlay').remove()">${T[lang].cancel}</button>
      <button class="icon-btn primary" onclick="saveSection('${catId}',this)">${T[lang].save}</button>
    </div>
  </div>`;
  document.body.appendChild(ov);
  ov.addEventListener('click',e=>{if(e.target===ov)ov.remove();});
};
window.pickColor=function(el){
  el.closest('.color-grid').querySelectorAll('.color-swatch').forEach(s=>s.classList.remove('selected'));
  el.classList.add('selected');
};
window.saveSection=function(catId,btn){
  const ov=btn.closest('.modal-overlay');
  const cat=categories.find(c=>c.id===catId);
  cat.label=ov.querySelector('#m-name').value;cat.labelKey='';
  cat.desc=ov.querySelector('#m-desc').value;
  cat.hasFte=ov.querySelector('#m-fte').checked;
  const sel=ov.querySelector('.color-swatch.selected');if(sel)cat.color=sel.dataset.c;
  ov.remove();renderAll();
};

function initSortableSections(){
  const wrap=document.getElementById('sections-wrap');
  Sortable.create(wrap,{handle:'.section-drag-handle',animation:150,
    onEnd(){
      const ids=[...wrap.querySelectorAll('.section')].map(el=>el.dataset.catId);
      categories.sort((a,b)=>ids.indexOf(a.id)-ids.indexOf(b.id));
    }
  });
}

function applyInstFilter(){
  categories.forEach(cat=>cat.rows.forEach(row=>{
    const tr=document.querySelector(`tr[data-row-id="${row.id}"]`);
    if(tr)tr.classList.toggle('hidden-row',instFilter!=='all'&&row.inst!==instFilter);
  }));
  updateTotals();
}

// ════════════════════════════════════════════════════
//  STATE SERIALIZATION — compact, minimal JSON
// ════════════════════════════════════════════════════

// Build a compact state object — strip defaults and empty values to minimise size
function buildState(){
  const defaultRates={};
  RATE_DEFS.forEach(r=>{defaultRates[r.key]=r.default;});

  // Only store rates that differ from defaults
  const changedRates={};
  Object.keys(rates).forEach(k=>{
    if(rates[k]!==defaultRates[k]) changedRates[k]=rates[k];
  });

  const st={
    v:2, // version
    l:lang==='en'?'en':undefined,
    c:currency==='EUR'?'EUR':undefined,
    e:eurRate!==7.4728?eurRate:undefined,
    n:numYears>1?numYears:undefined,
    f:instFilter!=='all'?instFilter:undefined,
    o:overheadPct||undefined,
    i:inflationPct!==2?inflationPct:undefined,
    p:document.getElementById('proj-title').value||undefined,
    r:Object.keys(changedRates).length?changedRates:undefined,
    cf:coFin.filter(x=>x.source||Object.keys(x.amounts).some(y=>x.amounts[y]>0))
       .map(x=>({
         id:x.id,
         s:x.source||undefined,
         st:x.status!=='granted'?x.status:undefined,
         a:Object.keys(x.amounts).filter(y=>x.amounts[y]>0).length?
           Object.fromEntries(Object.entries(x.amounts).filter(([,v])=>v>0)):undefined,
       })),
    cats:categories.map(cat=>({
      id:cat.id,
      lk:cat.labelKey||undefined,
      lb:cat.label,
      co:cat.color,
      ft:cat.hasFte?1:undefined,
      ds:cat.desc||undefined,
      rw:cat.rows
        .filter(row=>Object.keys(row.budget).some(y=>row.budget[y]>0)||Object.keys(row.fte).some(y=>row.fte[y]>0)||row.desc)
        .map(row=>({
          id:row.id,
          lk:row.labelKey||undefined,
          lb:row.label,
          in:row.inst!=='Rigshospitalet'?row.inst:undefined,
          rk:row.rateKey||undefined,
          ft:Object.keys(row.fte).length?row.fte:undefined,
          bu:Object.keys(row.budget).filter(y=>row.budget[y]>0).length?
             Object.fromEntries(Object.entries(row.budget).filter(([,v])=>v>0)):undefined,
          bm:Object.keys(row.budgetManual).filter(y=>row.budgetManual[y]).length?
             Object.fromEntries(Object.entries(row.budgetManual).filter(([,v])=>v)):undefined,
          ds:row.desc||undefined,
        }))
    }))
  };
  if(st.cf && st.cf.length===0) delete st.cf;
  return st;
}

// Expand compact state back to full state
function expandState(s){
  if(!s||!s.cats) return false;
  try{
    lang      = s.l||'da';
    currency  = s.c||'DKK';
    eurRate   = s.e||7.4728;
    numYears  = s.n||1;
    instFilter= s.f||'all';
    overheadPct = s.o||0;
    inflationPct= s.i!==undefined?s.i:2;
    if(s.r) Object.assign(rates,s.r);
    if(s.p){
      document.getElementById('proj-title').value=s.p;
      onProjTitle(s.p);
    }
    coFin=(s.cf||[]).map(x=>({
      id:x.id||uid(),
      source:x.s||'',
      status:x.st||'granted',
      amounts:x.a||{},
    }));
    categories=s.cats.map(cat=>({
      id:cat.id,
      labelKey:cat.lk||'',
      label:cat.lb,
      color:cat.co,
      hasFte:!!cat.ft,
      desc:cat.ds||'',
      rows:(cat.rw||[]).map(row=>({
        id:row.id,
        labelKey:row.lk||'',
        label:row.lb,
        inst:row.in||'Rigshospitalet',
        rateKey:row.rk||'',
        fte:row.ft||{},
        budget:row.bu||{},
        budgetManual:row.bm||{},
        desc:row.ds||'',
      }))
    }));
    // Re-add rows that were stripped (empty rows) as blank rows to preserve structure
    // Sync UI
    document.getElementById('rate-inp').value=eurRate;
    document.getElementById('overhead-inp').value=overheadPct;
    document.getElementById('inflation-inp').value=inflationPct;
    document.querySelectorAll('#year-btns .pill').forEach(p=>p.classList.toggle('active',+p.dataset.y===numYears));
    document.querySelectorAll('#inst-btns .pill').forEach(p=>p.classList.toggle('active',p.dataset.inst===instFilter));
    document.getElementById('btn-dkk').classList.toggle('active',currency==='DKK');
    document.getElementById('btn-eur').classList.toggle('active',currency==='EUR');
    document.getElementById('btn-da').classList.toggle('active',lang==='da');
    document.getElementById('btn-en').classList.toggle('active',lang==='en');
    return true;
  } catch(e){console.error('expandState error:',e);return false;}
}

// ════════════════════════════════════════════════════
//  SHARE LINK — two strategies:
//  1. window.storage (Claude artifact env): save JSON, link = #s=KEY (8 chars)
//  2. Fallback (GitHub Pages etc.): compress JSON into URL-safe base62, link = #d=...
// ════════════════════════════════════════════════════

// Base62 helpers for minimal URL encoding
const B62='0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
function b62enc(str){
  // Encode UTF-8 bytes as base62 — ~25% shorter than base64
  const bytes=new TextEncoder().encode(str);
  let n=BigInt(0);
  for(const b of bytes){ n=(n<<8n)|BigInt(b); }
  if(n===0n) return '0';
  let out='';
  while(n>0n){ out=B62[Number(n%62n)]+out; n/=62n; }
  return out;
}
function b62dec(s){
  let n=BigInt(0);
  for(const c of s){ n=n*62n+BigInt(B62.indexOf(c)); }
  const bytes=[];
  while(n>0n){ bytes.unshift(Number(n&255n)); n>>=8n; }
  return new TextDecoder().decode(new Uint8Array(bytes));
}

async function saveToStorage(key,data){
  try{
    if(typeof window.storage==='undefined') return false;
    const res=await window.storage.set('budget:'+key,JSON.stringify(data));
    return !!res;
  } catch(e){ return false; }
}
async function loadFromStorage(key){
  try{
    if(typeof window.storage==='undefined') return null;
    const res=await window.storage.get('budget:'+key);
    return res?JSON.parse(res.value):null;
  } catch(e){ return null; }
}

function shortId(){
  // 8 random base62 chars — 62^8 ≈ 218 trillion combinations
  let id='';
  for(let i=0;i<8;i++) id+=B62[Math.floor(Math.random()*62)];
  return id;
}

function showShareBar(shortUrl,fullUrl){
  const wrap=document.getElementById('share-bar-wrap');
  wrap.innerHTML=`<div class="share-bar">
    <i class="ti ti-link" style="color:var(--blue)"></i>
    <span style="font-weight:600;color:var(--blue);font-size:12px;white-space:nowrap">Del-link:</span>
    <input type="text" value="${shortUrl}" readonly onclick="this.select()" id="share-url" style="min-width:0;flex:1"/>
    <button class="icon-btn" onclick="copyShareLink()"><i class="ti ti-copy"></i> ${T[lang].copyLink}</button>
    ${fullUrl&&fullUrl!==shortUrl?`<span style="font-size:10px;color:var(--text3)">Kopiér hvis link ikke virker</span>`:''}
  </div>`;
}

window.shareLink=async function(){
  const state=buildState();
  const base=window.location.href.split('#')[0];

  // Try storage-based short link first
  const key=shortId();
  const saved=await saveToStorage(key,state);

  if(saved){
    const shortUrl=base+'#s='+key;
    window.history.replaceState(null,'',shortUrl);
    showShareBar(shortUrl,null);
    showToast('Kort link genereret ✓');
    return;
  }

  // Fallback: compact JSON → base62 → URL
  const json=JSON.stringify(state);
  const compressed=b62enc(json);
  const fallbackUrl=base+'#d='+compressed;
  window.history.replaceState(null,'',fallbackUrl);
  showShareBar(fallbackUrl,fallbackUrl);

  if(fallbackUrl.length>2000){
    showToast(lang==='da'?'Link er langt — brug del-knap til at kopiere':'Link is long — use copy button');
  } else {
    showToast(lang==='da'?'Link klar til deling':'Link ready to share');
  }
};

window.copyShareLink=function(){
  const inp=document.getElementById('share-url');
  if(!inp) return;
  inp.select();
  navigator.clipboard.writeText(inp.value).then(()=>{
    showToast(T[lang].linkCopied);
  }).catch(()=>{
    document.execCommand('copy');
    showToast(T[lang].linkCopied);
  });
};

function showToast(msg){
  const t=document.createElement('div');t.className='toast';t.textContent=msg;
  document.body.appendChild(t);
  setTimeout(()=>t.remove(),2500);
}

// Load shared state from URL hash
// Handles three formats:
//   #s=KEY      — short key stored in window.storage (Claude artifact env)
//   #d=BASE62   — compact base62-encoded JSON (GitHub Pages fallback)
//   #state=B64  — legacy long base64 format (backwards compat)
async function loadFromHash(){
  const hash=window.location.hash;

  // Format 1: storage-based short key
  if(hash.startsWith('#s=')){
    const key=hash.slice(3);
    const state=await loadFromStorage(key);
    if(state&&expandState(state)){
      buildRatesPanel();renderAll();
      return true;
    }
    // Key not found in storage — show helpful error
    showStorageError(key);
    return false;
  }

  // Format 2: base62-compressed JSON
  if(hash.startsWith('#d=')){
    try{
      const compressed=hash.slice(3);
      const json=b62dec(compressed);
      const state=JSON.parse(json);
      if(expandState(state)){
        buildRatesPanel();renderAll();
        return true;
      }
    } catch(e){ console.warn('b62 decode failed',e); }
    return false;
  }

  // Format 3: legacy base64 (old links still work)
  if(hash.startsWith('#state=')){
    try{
      const b64=hash.slice(7);
      const s=JSON.parse(decodeURIComponent(atob(b64)));
      // Map old format to new category structure
      if(s.categories) categories=s.categories;
      lang=s.lang||'da'; currency=s.currency||'DKK';
      eurRate=s.eurRate||7.4728; numYears=s.numYears||1;
      instFilter=s.instFilter||'all'; overheadPct=s.overheadPct||0;
      inflationPct=s.inflationPct!==undefined?s.inflationPct:2;
      if(s.rates) Object.assign(rates,s.rates);
      if(s.projTitle){
        document.getElementById('proj-title').value=s.projTitle;
        onProjTitle(s.projTitle);
      }
      document.getElementById('rate-inp').value=eurRate;
      document.getElementById('overhead-inp').value=overheadPct;
      document.getElementById('inflation-inp').value=inflationPct;
      document.querySelectorAll('#year-btns .pill').forEach(p=>p.classList.toggle('active',+p.dataset.y===numYears));
      document.querySelectorAll('#inst-btns .pill').forEach(p=>p.classList.toggle('active',p.dataset.inst===instFilter));
      document.getElementById('btn-dkk').classList.toggle('active',currency==='DKK');
      document.getElementById('btn-eur').classList.toggle('active',currency==='EUR');
      document.getElementById('btn-da').classList.toggle('active',lang==='da');
      document.getElementById('btn-en').classList.toggle('active',lang==='en');
      buildRatesPanel();renderAll();
      return true;
    } catch(e){ console.warn('legacy decode failed',e); }
    return false;
  }

  return false;
}

function showStorageError(key){
  const wrap=document.getElementById('share-bar-wrap');
  wrap.innerHTML=`<div class="share-bar" style="background:#FEF2F2;border-color:#FCA5A5">
    <i class="ti ti-alert-triangle" style="color:#DC2626"></i>
    <span style="color:#DC2626;font-size:12px;font-weight:500">
      ${lang==='da'?'Budget ikke fundet (nøgle: '+key+'). Linket er muligvis udløbet eller genereret i en anden session.':'Budget not found (key: '+key+'). The link may have expired or was created in a different session.'}
    </span>
  </div>`;
}

// ════════════════════════════════════════════════════
//  PDF EXPORT
// ════════════════════════════════════════════════════
function hexToRgb(hex){
  return [parseInt(hex.slice(1,3),16),parseInt(hex.slice(3,5),16),parseInt(hex.slice(5,7),16)];
}

window.exportPDF=function(){
  const {jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
  const yr=yrs();
  const projTitle=document.getElementById('proj-title').value||'';
  const isDa=lang==='da';
  const dateStr=new Date().toLocaleDateString(isDa?'da-DK':'en-GB');
  const isEUR=currency==='EUR';
  const curSym=isEUR?'EUR':'kr.';

  // -- ASCII sanitizer --------------------------------------------------
  // jsPDF's built-in fonts use WinAnsi encoding. A single character outside
  // that table (U+2212 true minus, U+00A0 nbsp, smart quotes, em dash, euro)
  // makes jsPDF encode that WORD as UTF-16BE. Rendered with a Latin font the
  // high byte of U+2212 (0x22) shows as a quote mark and every following ASCII
  // digit gains a leading null byte -- which is the '" 3 0 0 . 0 0 0' artefact.
  // Everything drawn into the PDF must pass through this.
  function pdfTxt(s){
    return String(s==null?'':s)
      .replace(/\u2212/g,'-')          // true minus     -> hyphen
      .replace(/\u00a0|\u202f/g,' ')   // nbsp / narrow  -> space
      .replace(/[\u2010-\u2015]/g,'-') // all dashes     -> hyphen
      .replace(/[\u2018\u2019\u201a]/g,"'")
      .replace(/[\u201c\u201d\u201e]/g,'"')
      .replace(/\u2026/g,'...')
      .replace(/\u00b7|\u2022/g,'-')
      .replace(/\u2248/g,'~')
      .replace(/\u20ac/g,'EUR')
      // final safety net: drop anything still outside printable WinAnsi
      .replace(/[^\n\x20-\x7E\xA1-\xFF]/g,'');
  }

  // -- Number grouping without toLocaleString ---------------------------
  // Browsers with full ICU return U+2212 (not '-') for negative numbers in
  // da-DK, and some locales use U+00A0 as the group separator. Formatting the
  // digits by hand removes that entire class of failure.
  function grp(n){
    const s=String(Math.abs(Math.round(n)));
    let out='';
    for(let i=0;i<s.length;i++){
      if(i>0 && (s.length-i)%3===0) out+='.';
      out+=s[i];
    }
    return out;
  }
  function grpFte(f){
    const r=Math.round(f*10)/10;
    return String(r).replace('.',',');
  }

  // -- Currency formatter for PDF (pure ASCII output) -------------------
  function pdfFmt(dkk){
    const v=isEUR?Math.round(dkk/eurRate):Math.round(dkk);
    const neg=v<0;
    const body=isEUR?('EUR '+grp(v)):(grp(v)+' kr.');
    return (neg?'-':'')+body;
  }

  // -- Cell-level guard --------------------------------------------------
  // Runs on every cell of every table. Even if a string reaches autoTable
  // un-sanitised, its text is cleaned here before line-breaking and drawing.
  function sanitizeCell(d){
    if(Array.isArray(d.cell.text)){
      d.cell.text=d.cell.text.map(t=>pdfTxt(t));
    } else if(typeof d.cell.text==='string'){
      d.cell.text=pdfTxt(d.cell.text);
    }
    d.cell.styles.fontSize=FS;
    d.cell.styles.valign='middle';
  }

  const gt=grandTotal(); // always in DKK
  const oh=overheadPct>0?Math.round(gt*overheadPct/100):0;

  const RED=[192,57,43],BLUE_H=[26,86,176],BLUE_TOT=[214,228,247];
  const DARK=[28,27,24],GREY2=[107,104,96],BLUE_LP=[235,245,255];
  const PAGE_W=297,MARGIN=8,TBL_W=PAGE_W-2*MARGIN; // 281mm exactly

  // ── Page header ──
  doc.setFillColor(...RED);doc.rect(0,0,PAGE_W,20,'F');
  doc.setTextColor(255,255,255);
  doc.setFontSize(13);doc.setFont(undefined,'bold');
  doc.text(pdfTxt(isDa?'Forskningsbudget':'Research Budget'),MARGIN+2,12);
  if(projTitle){doc.setFontSize(9);doc.setFont(undefined,'normal');doc.text(pdfTxt(projTitle),MARGIN+2,18);}
  doc.setFontSize(8);doc.setFont(undefined,'normal');
  doc.text(dateStr,PAGE_W-MARGIN,12,{align:'right'});
  const metaLine=[];
  if(isEUR) metaLine.push(`EUR (1 EUR = ${eurRate} DKK)`);
  if(inflationPct>0) metaLine.push(isDa?`Lønstigning: ${inflationPct}%/år`:`Salary increase: ${inflationPct}%/yr`);
  if(overheadPct>0) metaLine.push(`Overhead: ${overheadPct}%`);
  if(metaLine.length>0) doc.text(pdfTxt(metaLine.join(' | ')),PAGE_W-MARGIN,18,{align:'right'});

  let yPos=26;

  // ── Unified typography ──
  // One base size for every cell so nothing looks visually louder than its neighbours.
  // Emphasis is carried by weight and fill colour only, never by size.
  const FS = 8;          // base font size for all table text
  const ROW_H = 6.5;     // uniform body row height (mm)
  const HEAD_H = 7.5;    // header row height (mm)
  const PAD = [1.8, 2.5];// [vertical, horizontal] cell padding

  // ── Column widths — adaptive, but identical structure for every category ──
  // Institution must never wrap: "Rigshospitalet" needs ~24mm at 8pt.
  const nY = yr.length;
  const INST_C = 25;                       // wide enough for "Rigshospitalet" on one line
  const TOT    = nY >= 4 ? 27 : 30;
  const SUB    = nY >= 4 ? 38 : 44;
  const FTE_W_FTE = nY >= 4 ? 11 : 13;
  const BUD_MAX = 34;                      // stop budget cols becoming absurdly wide

  // Compute budget width for FTE layout, then let description absorb any slack
  function layoutFor(hasFte){
    const fteTotal = hasFte ? nY * FTE_W_FTE : 0;
    let desc = nY >= 4 ? 26 : 34;
    let avail = TBL_W - SUB - INST_C - desc - TOT - fteTotal;
    let bud = Math.floor(avail / nY);
    if (bud > BUD_MAX) {           // give surplus back to the description column
      desc += (bud - BUD_MAX) * nY;
      bud = BUD_MAX;
    }
    const used = SUB + INST_C + desc + TOT + fteTotal + bud * nY;
    const rem = TBL_W - used;      // rounding remainder goes on the last budget column
    return { desc, bud, rem };
  }

  // ── Per category ──
  categories.forEach(cat=>{
    const visRows=cat.rows.filter(row=>{
      if(instFilter!=='all'&&row.inst!==instFilter)return false;
      // Include row if ANY year has a budget amount OR FTE set
      return yr.some(y=>rowBudget(row,y)>0||(cat.hasFte&&(row.fte[y]||0)>0));
    });
    if(!visRows.length)return;

    // Estimate height for page break
    const descLines=cat.desc?Math.ceil(cat.desc.length/80)+1:0;
    const estH=8+descLines*4+(visRows.length+1)*7+(overheadPct>0?6:0);
    if(yPos+estH>200){doc.addPage();yPos=10;}

    // Category colour bar
    doc.setFillColor(...hexToRgb(cat.color));
    doc.rect(MARGIN,yPos,TBL_W,7,'F');
    doc.setTextColor(255,255,255);
    doc.setFontSize(9);doc.setFont(undefined,'bold');
    doc.text(pdfTxt(catLabel(cat).toUpperCase()),MARGIN+3,yPos+5);
    yPos+=7;
    if(cat.desc){
      doc.setTextColor(...GREY2);doc.setFontSize(FS-0.5);doc.setFont(undefined,'italic');
      const lines=doc.splitTextToSize(pdfTxt(cat.desc),TBL_W-6);
      doc.text(lines,MARGIN+3,yPos+4);
      yPos+=lines.length*4+3;
    }

    const L = layoutFor(cat.hasFte);
    const bw = L.bud, brem = L.rem, DESC_W = L.desc;

    // ── Build header ──
    // hasFte: two-row header (year group spans FTE+Budget cols).
    // non-FTE: single-row header ONLY — never use rowSpan on a single-row head because
    // jsPDF-autotable will treat the first body row as a header continuation, hiding it.
    let headArr;
    if(cat.hasFte){
      const hr1=[
        {content:pdfTxt(T[lang].sub),  rowSpan:2,styles:{valign:'middle',halign:'left'}},
        {content:pdfTxt(T[lang].inst), rowSpan:2,styles:{valign:'middle',halign:'center'}},
        {content:pdfTxt(T[lang].desc), rowSpan:2,styles:{valign:'middle',halign:'left'}},
      ];
      yr.forEach(y=>{
        hr1.push({content:pdfTxt(T[lang].yr(y)),colSpan:2,styles:{halign:'center',valign:'middle',fillColor:[195,215,245],textColor:[12,45,100]}});
      });
      hr1.push({content:pdfTxt(T[lang].total),rowSpan:2,styles:{halign:'right',valign:'middle'}});
      const hr2=[];
      yr.forEach(()=>{
        hr2.push({content:pdfTxt(T[lang].fte),styles:{halign:'right',valign:'middle'}});
        hr2.push({content:pdfTxt(T[lang].budget+' ('+curSym+')'),styles:{halign:'right',valign:'middle'}});
      });
      headArr=[hr1,hr2];
    } else {
      // Single flat header row — no rowSpan (would swallow the first body row)
      const sr=[
        {content:pdfTxt(T[lang].sub),  styles:{halign:'left',valign:'middle'}},
        {content:pdfTxt(T[lang].inst), styles:{halign:'center',valign:'middle'}},
        {content:pdfTxt(T[lang].desc), styles:{halign:'left',valign:'middle'}},
      ];
      yr.forEach(y=>{
        sr.push({content:pdfTxt(T[lang].yr(y)+' '+T[lang].budget+' ('+curSym+')'),styles:{halign:'right',valign:'middle',fillColor:[195,215,245],textColor:[12,45,100]}});
      });
      sr.push({content:pdfTxt(T[lang].total),styles:{halign:'right',valign:'middle'}});
      headArr=[sr];
    }

    // ── Body rows ──
    const bodyArr=[];
    visRows.forEach(row=>{
      const cells=[
        {content:pdfTxt(rowLabel(row)),styles:{halign:'left',valign:'middle'}},
        {content:pdfTxt(row.inst),styles:{halign:'center',valign:'middle'}},
        {content:pdfTxt(row.desc||''),styles:{halign:'left',valign:'middle',fontStyle:'italic',textColor:GREY2}},
      ];
      yr.forEach(y=>{
        if(cat.hasFte){
          const f=row.fte[y]||0;
          cells.push({content:f>0?grpFte(f):'-',styles:{halign:'right',valign:'middle'}});
        }
        const b=rowBudget(row,y);
        cells.push({content:b>0?pdfFmt(b):'-',styles:{halign:'right',valign:'middle'}});
      });
      let rowT=0;yr.forEach(y=>{rowT+=rowBudget(row,y);});
      cells.push({content:rowT>0?pdfFmt(rowT):'-',styles:{halign:'right',valign:'middle',fontStyle:'bold'}});
      bodyArr.push(cells);
    });

    // ── Subtotal row ──
    let catT=0;visRows.forEach(row=>{yr.forEach(y=>{catT+=rowBudget(row,y);});});
    const subCells=[
      {content:pdfTxt(`${T[lang].subtotal} - ${catLabel(cat)}`),colSpan:3,
       styles:{fontStyle:'bold',fillColor:BLUE_TOT,textColor:DARK,halign:'left',valign:'middle'}}
    ];
    yr.forEach(y=>{
      if(cat.hasFte) subCells.push({content:'',styles:{fillColor:BLUE_TOT}});
      let yt=0;visRows.forEach(row=>{yt+=rowBudget(row,y);});
      subCells.push({content:pdfFmt(yt),styles:{fontStyle:'bold',fillColor:BLUE_TOT,halign:'right',valign:'middle',textColor:DARK}});
    });
    subCells.push({content:pdfFmt(catT),styles:{fontStyle:'bold',fillColor:BLUE_TOT,halign:'right',valign:'middle',textColor:DARK}});
    bodyArr.push(subCells);

    // ── Overhead row (only if overheadPct > 0) ──
    if(overheadPct>0){
      const catOh=Math.round(catT*overheadPct/100);
      const ohCells=[
        {content:pdfTxt(`${T[lang].overheadRow} (${overheadPct}%)`),colSpan:3,
         styles:{textColor:BLUE_H,fillColor:BLUE_LP,halign:'left',valign:'middle'}}
      ];
      yr.forEach(y=>{
        if(cat.hasFte) ohCells.push({content:'',styles:{fillColor:BLUE_LP}});
        let yt=0;visRows.forEach(row=>{yt+=rowBudget(row,y);});
        ohCells.push({content:pdfFmt(Math.round(yt*overheadPct/100)),
          styles:{halign:'right',valign:'middle',textColor:BLUE_H,fillColor:BLUE_LP}});
      });
      ohCells.push({content:pdfFmt(catOh),
        styles:{halign:'right',valign:'middle',textColor:BLUE_H,fillColor:BLUE_LP}});
      bodyArr.push(ohCells);
    }

    // ── Column styles (sum exactly to TBL_W) ──
    const cs={};let ci=0;
    cs[ci++]={cellWidth:SUB,    halign:'left'};
    cs[ci++]={cellWidth:INST_C, halign:'center'};
    cs[ci++]={cellWidth:DESC_W, halign:'left', fontStyle:'italic', textColor:GREY2};
    yr.forEach((y,yi)=>{
      const isLast=yi===yr.length-1;
      if(cat.hasFte) cs[ci++]={cellWidth:FTE_W_FTE,halign:'right'};
      cs[ci++]={cellWidth:isLast?bw+brem:bw,halign:'right'};
    });
    cs[ci++]={cellWidth:TOT,halign:'right',fontStyle:'bold'};

    doc.autoTable({
      head:headArr,body:bodyArr,
      startY:yPos,margin:{left:MARGIN,right:MARGIN},tableWidth:TBL_W,
      styles:{
        font:'helvetica', fontSize:FS, cellPadding:PAD,
        lineColor:[210,208,200], lineWidth:0.2, textColor:DARK,
        overflow:'linebreak', valign:'middle', minCellHeight:ROW_H,
        cellWidth:'wrap',
      },
      headStyles:{
        fillColor:BLUE_H, textColor:[255,255,255], fontStyle:'bold',
        fontSize:FS, minCellHeight:HEAD_H, valign:'middle',
      },
      bodyStyles:{ fontSize:FS, minCellHeight:ROW_H, valign:'middle' },
      columnStyles:cs,
      didParseCell(d){
        sanitizeCell(d);
        if(d.section==='body'){
          const nSpecial=overheadPct>0?2:1; // subtotal + optional overhead
          if(d.row.index < bodyArr.length - nSpecial && d.row.index%2===1){
            d.cell.styles.fillColor=[250,250,248];
          }
        }
      },
    });
    yPos=doc.lastAutoTable.finalY+5;
  });

  // ── CO-FINANCING SECTION (only if entries with amounts exist) ──
  const cfRows=coFin.filter(cf=>yr.some(y=>(cf.amounts[y]||0)>0));
  if(cfRows.length>0){
    const cfEstH=8+(cfRows.length+1)*7;
    if(yPos+cfEstH>200){doc.addPage();yPos=10;}

    const GREEN=[20,96,58],GREEN_LIGHT=[197,229,211];
    doc.setFillColor(...GREEN);
    doc.rect(MARGIN,yPos,TBL_W,7,'F');
    doc.setTextColor(255,255,255);
    doc.setFontSize(9);doc.setFont(undefined,'bold');
    doc.text(pdfTxt(T[lang].cofinTitle.toUpperCase()),MARGIN+3,yPos+5);
    yPos+=7;

    // Header: Source | Status | Year cols | Total
    const cfHead=[[
      {content:pdfTxt(T[lang].cofinSource),styles:{halign:'left',valign:'middle'}},
      {content:pdfTxt(T[lang].cofinStatus),styles:{halign:'center',valign:'middle'}},
      ...yr.map(y=>({content:pdfTxt(T[lang].yr(y)),styles:{halign:'right',valign:'middle',fillColor:GREEN_LIGHT,textColor:GREEN}})),
      {content:pdfTxt(T[lang].total),styles:{halign:'right',valign:'middle'}},
    ]];

    const cfBody=[];
    cfRows.forEach(cf=>{
      const cells=[
        {content:pdfTxt(cf.source||'-'),styles:{halign:'left',valign:'middle'}},
        {content:pdfTxt(statusLabel(cf.status)),styles:{halign:'center',valign:'middle'}},
      ];
      yr.forEach(y=>{
        const a=cf.amounts[y]||0;
        cells.push({content:a>0?pdfFmt(a):'-',styles:{halign:'right',valign:'middle'}});
      });
      cells.push({content:pdfFmt(coFinRowTotal(cf)),styles:{halign:'right',valign:'middle',fontStyle:'bold'}});
      cfBody.push(cells);
    });

    // Co-financing subtotal
    const cfSub=[
      {content:pdfTxt(T[lang].cofinSubtotal),colSpan:2,styles:{fontStyle:'bold',fillColor:GREEN_LIGHT,textColor:GREEN,halign:'left',valign:'middle'}}
    ];
    yr.forEach(y=>{
      cfSub.push({content:pdfFmt(coFinYear(y)),styles:{fontStyle:'bold',fillColor:GREEN_LIGHT,halign:'right',valign:'middle',textColor:GREEN}});
    });
    cfSub.push({content:pdfFmt(coFinTotal()),styles:{fontStyle:'bold',fillColor:GREEN_LIGHT,halign:'right',valign:'middle',textColor:GREEN}});
    cfBody.push(cfSub);

    // Column widths — status fixed, year cols equal, source absorbs remainder
    const CF_STATUS=26, CF_TOT=TOT;
    let cfValW=Math.floor((TBL_W-CF_STATUS-CF_TOT-70)/nY);
    if(cfValW>BUD_MAX) cfValW=BUD_MAX;
    const cfSrcW=TBL_W-CF_STATUS-CF_TOT-cfValW*nY;
    const cfCS={0:{cellWidth:cfSrcW,halign:'left'},1:{cellWidth:CF_STATUS,halign:'center'}};
    yr.forEach((_,i)=>{cfCS[i+2]={cellWidth:cfValW,halign:'right'};});
    cfCS[nY+2]={cellWidth:CF_TOT,halign:'right',fontStyle:'bold'};

    doc.autoTable({
      head:cfHead,body:cfBody,
      startY:yPos,margin:{left:MARGIN,right:MARGIN},tableWidth:TBL_W,
      styles:{
        font:'helvetica', fontSize:FS, cellPadding:PAD,
        lineColor:[157,201,174], lineWidth:0.2, textColor:DARK,
        overflow:'linebreak', valign:'middle', minCellHeight:ROW_H,
      },
      headStyles:{
        fillColor:GREEN, textColor:[255,255,255], fontStyle:'bold',
        fontSize:FS, minCellHeight:HEAD_H, valign:'middle',
      },
      bodyStyles:{ fontSize:FS, minCellHeight:ROW_H, valign:'middle' },
      columnStyles:cfCS,
      didParseCell(d){
        sanitizeCell(d);
        if(d.section==='body'&&d.row.index<cfBody.length-1&&d.row.index%2===1){
          d.cell.styles.fillColor=[245,251,247];
        }
      },
    });
    yPos=doc.lastAutoTable.finalY+5;
  }

  // ── Grand total summary table ──
  if(yPos>185){doc.addPage();yPos=12;}
  yPos+=4;

  // Header: category name col + one col per year + total col
  const gtHead=[[
    {content:pdfTxt(isDa?'Kategori':'Category'),styles:{fillColor:BLUE_H,textColor:[255,255,255],halign:'left',valign:'middle'}},
    ...yr.map(y=>({content:pdfTxt(T[lang].yr(y)),styles:{fillColor:BLUE_H,textColor:[255,255,255],halign:'right',valign:'middle'}})),
    {content:pdfTxt(T[lang].total),styles:{fillColor:BLUE_H,textColor:[255,255,255],halign:'right',valign:'middle'}},
  ]];

  const gtBody=[];
  categories.forEach(cat=>{
    const visRows=cat.rows.filter(row=>{
      if(instFilter!=='all'&&row.inst!==instFilter)return false;
      return yr.some(y=>rowBudget(row,y)>0);
    });
    if(!visRows.length)return;
    const row=[{content:pdfTxt(catLabel(cat)),styles:{halign:'left',valign:'middle'}}];
    let catT=0;
    yr.forEach(y=>{
      let yt=0;visRows.forEach(r=>{yt+=rowBudget(r,y);});
      row.push({content:pdfFmt(yt),styles:{halign:'right',valign:'middle'}});catT+=yt;
    });
    row.push({content:pdfFmt(catT),styles:{halign:'right',valign:'middle',fontStyle:'bold'}});
    gtBody.push(row);
  });

  // Grand total row — bold only; size stays at FS so row height matches the rest
  const totRow=[{content:pdfTxt(isDa?'SAMLET BUDGET (ALLE ÅR)':'TOTAL BUDGET (ALL YEARS)'),
    styles:{fontStyle:'bold',fillColor:BLUE_TOT,textColor:DARK,halign:'left',valign:'middle'}}];
  yr.forEach(y=>{
    totRow.push({content:pdfFmt(yearGrand(y)),
      styles:{halign:'right',valign:'middle',fontStyle:'bold',fillColor:BLUE_TOT,textColor:DARK}});
  });
  totRow.push({content:pdfFmt(gt),
    styles:{halign:'right',valign:'middle',fontStyle:'bold',fillColor:BLUE_TOT,textColor:DARK}});
  gtBody.push(totRow);

  // Overhead total (only if > 0)
  if(overheadPct>0){
    const ohTotRow=[{content:pdfTxt(`${T[lang].totalWithOh} (${T[lang].overheadRow} ${overheadPct}%)`),
      styles:{fontStyle:'bold',textColor:BLUE_H,fillColor:BLUE_LP,halign:'left',valign:'middle'}}];
    yr.forEach(y=>{
      const yt=yearGrand(y);
      ohTotRow.push({content:pdfFmt(Math.round(yt*(1+overheadPct/100))),
        styles:{halign:'right',valign:'middle',fontStyle:'bold',fillColor:BLUE_LP,textColor:BLUE_H}});
    });
    ohTotRow.push({content:pdfFmt(gt+oh),
      styles:{halign:'right',valign:'middle',fontStyle:'bold',fillColor:BLUE_LP,textColor:BLUE_H}});
    gtBody.push(ohTotRow);
  }

  // -- Co-financing deduction + amount requested --
  // The label and figures use a plain ASCII hyphen. U+2212 (true minus) sits
  // outside WinAnsi, which forces jsPDF into a UTF-16 path and renders as a
  // stray quote mark followed by null-byte "letter spacing".
  const cfT=coFinTotal();
  if(cfT>0){
    const GREEN_D=[20,96,58],GREEN_L=[230,245,237];
    const RED_L=[253,236,234],RED_D=[123,26,18];
    const cfLabel=isDa?'- Medfinansiering':'- Co-financing';

    const cfDedRow=[{content:pdfTxt(cfLabel),
      styles:{fontStyle:'bold',textColor:GREEN_D,fillColor:GREEN_L,halign:'left',valign:'middle'}}];
    yr.forEach(y=>{
      const yc=coFinYear(y);
      cfDedRow.push({content:yc>0?('-'+pdfFmt(yc)):'-',
        styles:{halign:'right',valign:'middle',fontStyle:'bold',fillColor:GREEN_L,textColor:GREEN_D}});
    });
    cfDedRow.push({content:'-'+pdfFmt(cfT),
      styles:{halign:'right',valign:'middle',fontStyle:'bold',fillColor:GREEN_L,textColor:GREEN_D}});
    gtBody.push(cfDedRow);

    // Amount requested -- the headline figure for the funder
    const soughtRow=[{content:pdfTxt(T[lang].amountSought.toUpperCase()),
      styles:{fontStyle:'bold',textColor:RED_D,fillColor:RED_L,halign:'left',valign:'middle'}}];
    yr.forEach(y=>{
      soughtRow.push({content:pdfFmt(yearSought(y)),
        styles:{halign:'right',valign:'middle',fontStyle:'bold',fillColor:RED_L,textColor:RED_D}});
    });
    soughtRow.push({content:pdfFmt(amountSought()),
      styles:{halign:'right',valign:'middle',fontStyle:'bold',fillColor:RED_L,textColor:RED_D}});
    gtBody.push(soughtRow);
  }

  // Column widths: label gets 40% of TBL_W, rest split equally among yr+total cols
  const gtValW=Math.floor(TBL_W*0.60/(yr.length+1));
  const gtLblW=TBL_W-gtValW*(yr.length+1);
  const gtCS={0:{cellWidth:gtLblW,halign:'left',fontStyle:'bold'}};
  yr.forEach((_,i)=>{gtCS[i+1]={cellWidth:gtValW,halign:'right'};});
  gtCS[yr.length+1]={cellWidth:gtValW,halign:'right',fontStyle:'bold'};

  doc.autoTable({
    head:gtHead,body:gtBody,
    startY:yPos,margin:{left:MARGIN,right:MARGIN},tableWidth:TBL_W,
    styles:{
      font:'helvetica', fontSize:FS, cellPadding:PAD,
      lineColor:[192,57,43], lineWidth:0.25, textColor:DARK,
      overflow:'linebreak', valign:'middle', minCellHeight:ROW_H,
    },
    headStyles:{
      fillColor:BLUE_H, textColor:[255,255,255], fontStyle:'bold',
      fontSize:FS, minCellHeight:HEAD_H, valign:'middle',
    },
    bodyStyles:{ fontSize:FS, minCellHeight:ROW_H, valign:'middle' },
    columnStyles:gtCS,
    didParseCell(d){
      sanitizeCell(d);
      if(d.section==='body'){
        let nSpec=1;                      // grand total row
        if(overheadPct>0) nSpec++;        // + overhead row
        if(coFinTotal()>0) nSpec+=2;      // + co-financing and requested rows
        if(d.row.index<gtBody.length-nSpec && d.row.index%2===1){
          d.cell.styles.fillColor=[250,250,248];
        }
      }
    }
  });

  // ── Page numbers (no confidential footer) ──
  const pc=doc.internal.getNumberOfPages();
  for(let i=1;i<=pc;i++){
    doc.setPage(i);
    doc.setFontSize(7);doc.setTextColor(...GREY2);doc.setFont(undefined,'normal');
    doc.text(pdfTxt(`${isDa?'Side':'Page'} ${i} / ${pc}`),PAGE_W-MARGIN,207,{align:'right'});
  }

  const safeName=(projTitle||'budget').replace(/[^a-zA-Z0-9æøåÆØÅ]/g,'_');
  doc.save(isDa?`Forskningsbudget_${safeName}.pdf`:`Research_Budget_${safeName}.pdf`);
};

// ════════════════════════════════════════════════════
//  EXCEL EXPORT — client-side via SheetJS (fixed)
// ════════════════════════════════════════════════════
window.exportExcel=function(){
  const yr=yrs();
  const projTitle=document.getElementById('proj-title').value||'';
  const gt=grandTotal();
  const oh=overheadPct>0?Math.round(gt*overheadPct/100):0;
  const isDa=lang==='da';

  // Convert value to display currency
  function dispVal(dkk){
    if(currency==='EUR') return Math.round(dkk/eurRate);
    return Math.round(dkk);
  }
  const curLabel=currency==='EUR'?'EUR':'DKK';

  const wb=XLSX.utils.book_new();

  // ── BUDGET SHEET ──
  const aoa=[];

  // Title row
  aoa.push([isDa?'Forskningsbudget':'Research Budget',projTitle]);
  aoa.push([isDa?`Dato: ${new Date().toLocaleDateString('da-DK')}`:`Date: ${new Date().toLocaleDateString('en-GB')}`,'']);
  if(inflationPct>0) aoa.push([isDa?`Lønstigning: ${inflationPct}% pr. år`:`Salary increase: ${inflationPct}% p.a.`,'']);
  aoa.push([]);

  // Header row
  const hdr=[isDa?'Kategori':'Category', isDa?'Underkategori':'Subcategory', isDa?'Beskrivelse':'Description', isDa?'Institution':'Institution'];
  yr.forEach(y=>{
    if(true) hdr.push(`${T[lang].yr(y)} ${T[lang].budget} (${curLabel})`);
    // FTE columns for salary
  });
  hdr.push(`${T[lang].total} (${curLabel})`);
  aoa.push(hdr);

  const catSectionRows={}; // track row indices for styling
  let rowIdx=aoa.length; // 0-based

  categories.forEach(cat=>{
    const visRows=cat.rows.filter(row=>{
      if(instFilter!=='all'&&row.inst!==instFilter)return false;
      return yr.some(y=>rowBudget(row,y)>0||(cat.hasFte&&(row.fte[y]||0)>0));
    });
    if(!visRows.length)return;

    catSectionRows[cat.id]={start:rowIdx,end:0};

    visRows.forEach((row,i)=>{
      const r=[i===0?catLabel(cat):'',rowLabel(row),row.desc||'',row.inst];
      yr.forEach(y=>{r.push(dispVal(rowBudget(row,y)));});
      let rowT=0;yr.forEach(y=>{rowT+=rowBudget(row,y);});
      r.push(dispVal(rowT));
      aoa.push(r);rowIdx++;
    });

    // Subtotal row
    const sub=[`${isDa?'Subtotal':'Subtotal'} — ${catLabel(cat)}`,'','',''];
    yr.forEach(y=>{
      let yt=0;visRows.forEach(row=>{yt+=rowBudget(row,y);});
      sub.push(dispVal(yt));
    });
    let catT=0;visRows.forEach(row=>{yr.forEach(y=>{catT+=rowBudget(row,y);});});
    sub.push(dispVal(catT));
    aoa.push(sub);catSectionRows[cat.id].end=rowIdx;rowIdx++;

    if(overheadPct>0){
      const ohRow=[`${isDa?'Overhead':'Overhead'} (${overheadPct}%)`,'','',''];
      yr.forEach(y=>{
        let yt=0;visRows.forEach(row=>{yt+=rowBudget(row,y);});
        ohRow.push(dispVal(Math.round(yt*overheadPct/100)));
      });
      const catOh=Math.round(catT*overheadPct/100);
      ohRow.push(dispVal(catOh));
      aoa.push(ohRow);rowIdx++;
    }
    aoa.push([]);rowIdx++;
  });

  // ── CO-FINANCING SECTION ──
  const cfRowsX=coFin.filter(cf=>yr.some(y=>(cf.amounts[y]||0)>0));
  if(cfRowsX.length>0){
    aoa.push([isDa?'MEDFINANSIERING — ALLEREDE BEVILGEDE MIDLER':'CO-FINANCING — FUNDING ALREADY SECURED','','','']);
    aoa.push([isDa?'Finansieringskilde':'Funding source',isDa?'Status':'Status','','']);
    rowIdx+=2;
    cfRowsX.forEach(cf=>{
      const r=[cf.source||'—',statusLabel(cf.status),'',''];
      yr.forEach(y=>{r.push(dispVal(cf.amounts[y]||0));});
      r.push(dispVal(coFinRowTotal(cf)));
      aoa.push(r);rowIdx++;
    });
    const cfSubX=[isDa?'Medfinansiering i alt':'Total co-financing','','',''];
    yr.forEach(y=>{cfSubX.push(dispVal(coFinYear(y)));});
    cfSubX.push(dispVal(coFinTotal()));
    aoa.push(cfSubX);rowIdx++;
    aoa.push([]);rowIdx++;
  }

  // ── SUMMARY: total budget → overhead → co-financing → amount sought ──
  const gtRow=[isDa?'SAMLET BUDGET':'TOTAL BUDGET','','',''];
  yr.forEach(y=>{gtRow.push(dispVal(yearGrand(y)));});
  gtRow.push(dispVal(gt));
  aoa.push(gtRow);rowIdx++;

  if(overheadPct>0){
    const ohGtRow=[`${isDa?'Budget inkl. overhead':'Budget incl. overhead'} (${overheadPct}%)`,'','',''];
    yr.forEach(y=>{ohGtRow.push(dispVal(Math.round(yearGrand(y)*(1+overheadPct/100))));});
    ohGtRow.push(dispVal(gt+oh));
    aoa.push(ohGtRow);rowIdx++;
  }

  const cfTotX=coFinTotal();
  if(cfTotX>0){
    const cfDedX=[isDa?'- Medfinansiering':'- Co-financing','','',''];
    yr.forEach(y=>{cfDedX.push(-dispVal(coFinYear(y)));});
    cfDedX.push(-dispVal(cfTotX));
    aoa.push(cfDedX);rowIdx++;

    const soughtX=[isDa?'ANSØGT BELØB':'AMOUNT REQUESTED','','',''];
    yr.forEach(y=>{soughtX.push(dispVal(yearSought(y)));});
    soughtX.push(dispVal(amountSought()));
    aoa.push(soughtX);rowIdx++;
  }

  const ws=XLSX.utils.aoa_to_sheet(aoa);

  // Column widths
  const colW=[{wch:28},{wch:32},{wch:30},{wch:16}];
  yr.forEach(()=>colW.push({wch:18}));
  colW.push({wch:18});
  ws['!cols']=colW;

  XLSX.utils.book_append_sheet(wb,ws,isDa?'Budget':'Budget');

  // ── ASSUMPTIONS SHEET ──
  const assumpData=[
    [isDa?'Forudsætninger':'Assumptions',''],
    [],
    [isDa?'EUR/DKK kurs':'EUR/DKK rate',eurRate],
    [isDa?'Valuta':'Currency',currency],
    [isDa?'Overhead (%)':'Overhead (%)',overheadPct],
    [isDa?'Lønstigning (%/år)':'Salary increase (%/yr)',inflationPct],
    [],
    [isDa?'Lønsatser (DKK/år ved 1,0 FTE, År 1)':'Salary rates (DKK/yr at 1.0 FTE, Year 1)',''],
    [isDa?'Rolle':'Role',isDa?'Lønsats (DKK)':'Rate (DKK)',isDa?'Standard institution':'Default institution'],
  ];
  RATE_DEFS.forEach(r=>{
    assumpData.push([T[lang].salCats[r.key]||r.key,rates[r.key],r.defaultInst]);
  });
  if(coFin.length>0){
    assumpData.push([]);
    assumpData.push([isDa?'Medfinansiering':'Co-financing','','']);
    assumpData.push([isDa?'Kilde':'Source',isDa?'Status':'Status',isDa?'Beløb i alt (DKK)':'Total (DKK)']);
    coFin.forEach(cf=>{
      assumpData.push([cf.source||'—',statusLabel(cf.status),coFinRowTotal(cf)]);
    });
  }
  const aws=XLSX.utils.aoa_to_sheet(assumpData);
  aws['!cols']=[{wch:36},{wch:18},{wch:20}];
  XLSX.utils.book_append_sheet(wb,aws,isDa?'Forudsætninger':'Assumptions');

  // Download
  const wbout=XLSX.write(wb,{bookType:'xlsx',type:'base64'});
  const fname2=projTitle?`Budget_${projTitle.replace(/[^a-zA-Z0-9æøåÆØÅ]/g,'_')}.xlsx`:(isDa?'Forskningsbudget.xlsx':'Research_Budget.xlsx');
  const link=document.createElement('a');
  link.href='data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,'+wbout;
  link.download=fname2;
  document.body.appendChild(link);link.click();document.body.removeChild(link);
};

// ════════════════════════════════════════════════════
//  EVENTS
// ════════════════════════════════════════════════════
document.getElementById('year-btns').addEventListener('click',e=>{
  const b=e.target.closest('.pill');if(!b||!b.dataset.y)return;
  numYears=parseInt(b.dataset.y);
  document.querySelectorAll('#year-btns .pill').forEach(p=>p.classList.toggle('active',+p.dataset.y===numYears));
  recalcAllSalary();renderAll();
});
document.getElementById('inst-btns').addEventListener('click',e=>{
  const b=e.target.closest('.pill');if(!b||!b.dataset.inst)return;
  instFilter=b.dataset.inst;
  document.querySelectorAll('#inst-btns .pill').forEach(p=>p.classList.toggle('active',p.dataset.inst===instFilter));
  applyInstFilter();
});

// ════════════════════════════════════════════════════
//  INIT
// ════════════════════════════════════════════════════
// Init — async because loadFromHash may fetch from storage
(async()=>{
  const loaded=await loadFromHash();
  if(!loaded){
    buildRatesPanel();renderAll();
  }
})();
</script>

</body>
</html>
