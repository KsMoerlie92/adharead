/* Large Item Overview — Royal IHC
 * Toont Large Items, Late Items en FAT Overview vanuit de BEDRIJFSBREDE
 * expediting-lijst. Geen moederlijst nodig.
 *
 * Kolommen worden op NAAM gezocht (robuust), met de kolomletter als terugval.
 * Bevestigde layout (Bedrijfsbreed Expediten, 43 kolommen):
 *   H  Technical Coordinator Name   V  Last Expedited     AO Net Weight
 *   O  Qty                          X  Delivery Address   AP Total Net Weight
 *   S  Delivery Status              Y  Delivery Terms     AQ Weight UoM
 *   T  Latest Wanted Receipt Date   AH FAT Location
 *   U  Planned Delivery Date        AI FAT Date Required
 */
(function () {
  'use strict';

  // naam(en) → kolomletter als terugval
  const MAP = {
    po:        [['Purchase Order No'], 'A'],
    orderNo:   [['Order No'], 'B'],
    status:    [['PO Line Status'], 'E'],
    sp:        [['Sub Project ID'], 'F'],
    spDesc:    [['Sub Project Description'], 'G'],
    engineer:  [['Technical Coordinator Name'], 'H'],
    supplier:  [['Supplier Name'], 'J'],
    part:      [['Part No'], 'K'],
    desc:      [['Description'], 'L'],
    qty:       [['Qty'], 'O'],
    delStatus: [['Delivery Status'], 'S'],
    wanted:    [['Latest Wanted Receipt Date'], 'T'],
    planned:   [['Planned Delivery Date'], 'U'],
    lastExp:   [['Last Expedited'], 'V'],
    address:   [['Delivery Address'], 'X'],
    terms:     [['Delivery Terms'], 'Y'],
    fatLoc:    [['FAT Location'], 'AH'],
    fatDate:   [['FAT Date Required'], 'AI'],
    protoDate: [['FAT Supplier Protocol Date'], 'AG'],
    sysNo:     [['Unified Reference Code'], 'M'],
    refDesc:   [['Unified Ref Code Description'], 'N'],
    netW:      [['Net Weight'], 'AO'],
    totW:      [['Total Net Weight'], 'AP'],
    wUoM:      [['Weight UoM'], 'AQ'],
    // BETA-tab: aantal additional requirements. Optioneel (geen waarschuwing als hij ontbreekt).
    addReq:    [['Additional Requirements', 'Additional Requirement', 'Additional requirements'], 'AG'],
  };
  const OPTIONAL = new Set(['addReq']);
  // Kolommen zónder welke een tabblad zinloos is. Overige kolommen mogen
  // ontbreken: die blijven leeg, met een melding erboven.
  const NEEDED = {
    large: ['totW'],
    late:  ['delStatus'],
    fat:   ['fatDate'],
    beta:  ['totW'],
  };

  let HEADERS = [], ROWS = [], C = {}, MISSING = [];
  let TAB = 'large', SPSEL = new Set(), THRESHOLD = 1000;

  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const trim = (v) => String(v == null ? '' : v).trim();
  const flat = (v) => trim(v).replace(/\s*[\r\n]+\s*/g, ', ').replace(/\s{2,}/g, ' ');

  // Zoekt uitsluitend op kolomNAAM. De letter in MAP is documentatie: zodra de
  // Admin-upload kolommen wegfiltert kloppen posities niet meer, en een
  // letter-terugval koos dan stilzwijgend de verkeerde kolom.
  function resolveCols(headers) {
    const low = headers.map(h => trim(h).toLowerCase());
    const cols = {}; const missing = [];
    for (const key in MAP) {
      const names = MAP[key][0];
      let i = -1;
      for (const n of names) { i = low.indexOf(n.toLowerCase()); if (i >= 0) break; }
      if (i < 0 && key === 'addReq') i = low.findIndex(h => /additional\s*req/.test(h));   // ruimere herkenning
      if (i < 0 && !OPTIONAL.has(key)) missing.push(names[0]);
      cols[key] = i;
    }
    return { cols, missing };
  }

  // ── waarden ──────────────────────────────────────────────────────────────
  const val = (r, key) => (C[key] >= 0 && r[C[key]] != null) ? r[C[key]] : '';
  // Robuuste datumconversie — overgenomen uit PO-Matcher/js/renderer.js zodat
  // Large Item Overview en de Late Items-tab in PO-Matcher altijd dezelfde
  // datum tonen voor hetzelfde item. Verankert Excel-seriegetallen op het
  // MIDDEN van de dag (UTC-noon) vóór lokale reconstructie, en reconstrueert
  // ISO-tijdstempels (met 'T') via lokale getters — anders kan een datum in
  // een tijdzone áchter UTC één dag terugvallen (bv. 12/05 wordt 11/05).
  function toDate(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    const s = trim(v); if (!s) return null;
    if (typeof v === 'number' || /^\d+([.,]\d+)?$/.test(s)) {
      const n = parseFloat(s.replace(',', '.'));
      if (!isNaN(n) && n > 1000) {
        const d = new Date(Math.round((n - 25569 + 0.5) * 86400 * 1000));
        return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
      }
    }
    const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) {
      if (s.indexOf('T') > -1) { const dt = new Date(s); if (!isNaN(dt.getTime())) return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate()); }
      return new Date(+iso[1], +iso[2] - 1, +iso[3]);
    }
    const eu = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
    if (eu) return new Date(+eu[3], +eu[2] - 1, +eu[1]);
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  const fmtDate = (v) => { const d = toDate(v); if (!d) return ''; 
    return String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear(); };
  function num(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    let s = trim(v).replace(/[^\d.,-]/g, '');          // eenheden/spaties weg
    if (!s) return null;
    const c = s.lastIndexOf(','), d = s.lastIndexOf('.');
    if (c > -1 && d > -1) {
      // beide aanwezig → de laatste is het decimaalteken
      s = (c > d) ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    } else if (c > -1) {
      s = (s.match(/,/g) || []).length > 1 ? s.replace(/,/g, '') : s.replace(',', '.');
    } else if (d > -1 && (s.match(/\./g) || []).length > 1) {
      s = s.replace(/\./g, '');                        // 1.234.567 → duizendtallen
    }
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  }
  const fmtNum = (n) => n == null ? '' : n.toLocaleString('nl-NL', { maximumFractionDigits: 1 });
  const startOfToday = () => { const d = new Date(); d.setHours(0,0,0,0); return d; };

  // ── data laden ───────────────────────────────────────────────────────────
  async function boot() {
    wireTabs(); injectFatStyle();
    $('lio-tbody').addEventListener('change', (e) => {
      const t = e.target; if (!t.classList || !t.classList.contains('lio-in')) return;
      setFi(t.dataset.k, t.dataset.f, t.value.trim());
      t.classList.toggle('lio-in-dirty', true); t.classList.remove('lio-in-derived');
    });
    loadFatInfo();
    $('lio-upload').addEventListener('click', () => $('lio-file').click());
    $('lio-file').addEventListener('change', onUpload);
    try {
      const raw = window.ExpeditingData && ExpeditingData.loadRaw
        ? await ExpeditingData.loadRaw() : null;
      const meta = window.ExpeditingData && ExpeditingData.meta ? await ExpeditingData.meta() : null;
      if (raw && raw.headers && raw.rows && raw.rows.length) {
        setData(raw.headers, raw.rows,
          (meta && meta.filename ? meta.filename : 'bedrijfsbrede lijst') +
          (meta && meta.uploaded ? ' · ' + fmtDate(meta.uploaded) : ''));
      } else {
        $('lio-src-txt').innerHTML = '⚠ Geen bedrijfsbrede lijst gevonden — laad er een via de knop hiernaast.';
      }
    } catch (e) {
      $('lio-src-txt').innerHTML = '⚠ Kon de bedrijfsbrede lijst niet laden: ' + esc(e.message);
    }
  }

  function onUpload(ev) {
    const f = ev.target.files && ev.target.files[0]; ev.target.value = '';
    if (!f) return;
    $('lio-src-txt').textContent = 'Bezig met inlezen…';
    const rd = new FileReader();
    rd.onload = () => {
      try {
        const wb = XLSX.read(new Uint8Array(rd.result), { type: 'array', cellDates: true });
        const name = pickSheet(wb);
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' });
        const t = window.ExpeditingCore && ExpeditingCore.rawTable ? null : null;
        const hIdx = detectHeader(aoa);
        setData(aoa[hIdx].map(x => trim(x)), aoa.slice(hIdx + 1), f.name);
      } catch (e) {
        $('lio-src-txt').innerHTML = '⚠ Fout bij inlezen: ' + esc(e.message);
      }
    };
    rd.readAsArrayBuffer(f);
  }
  function pickSheet(wb) {
    const rows = (n) => { const ws = wb.Sheets[n]; if (!ws || !ws['!ref']) return 0;
      const r = XLSX.utils.decode_range(ws['!ref']); return r.e.r - r.s.r + 1; };
    return wb.SheetNames.slice().sort((a, b) => rows(b) - rows(a))[0];
  }
  function detectHeader(aoa) {
    let best = 0, score = -1;
    for (let i = 0; i < Math.min(7, aoa.length); i++) {
      const n = (aoa[i] || []).filter(c => trim(c)).length;
      if (n > score) { score = n; best = i; }
    }
    return best;
  }

  function setData(headers, rows, srcLabel) {
    // ExpeditingData.loadRaw() geeft rijen als objecten (op kolomnaam),
    // een eigen upload geeft arrays. Alles hier gelijktrekken naar arrays.
    if (rows.length && !Array.isArray(rows[0])) {
      rows = rows.map(o => headers.map(h => (o[h] == null ? '' : o[h])));
    }
    HEADERS = headers; ROWS = rows;
    const r = resolveCols(headers); C = r.cols; MISSING = r.missing;
    $('lio-src-txt').innerHTML = '📊 Bron: <b>' + esc(srcLabel) + '</b> — ' +
      rows.length.toLocaleString('nl-NL') + ' regels, ' + headers.length + ' kolommen';
    renderWarnings(); buildSubProjectPanel(); render();
  }

  function injectFatStyle() {
    if (document.getElementById('lio-fat-style')) return;
    const st = document.createElement('style'); st.id = 'lio-fat-style';
    st.textContent =
      '.lio-in{font-family:inherit;font-size:.74rem;background:var(--navy,#0A1628);color:var(--white,#F0F4FA);border:1px solid var(--steel,#1e3a6e);border-radius:4px;padding:.2rem .3rem;max-width:100%}' +
      '.lio-in-txt{min-width:170px}' +
      '.lio-in:focus{outline:none;border-color:var(--teal,#00B4D8)}' +
      '.lio-in-dirty{border-color:#E8B923;box-shadow:0 0 0 1px rgba(232,185,35,.5)}' +
      '.lio-in-derived{color:var(--grey,#8FA3BF);font-style:italic}' +
      '.lio-fat-status{margin-left:.6rem;font-size:.72rem;color:var(--grey,#8FA3BF)}' +
      '.lio-fat-dirty{color:#E8B923;font-weight:700}';
    document.head.appendChild(st);
  }

  function renderWarnings() {
    const el = $('lio-warnings');
    if (!MISSING.length) { el.innerHTML = ''; return; }
    el.innerHTML = '<div class="lio-warn">⚠ Deze kolommen zitten niet in de geladen lijst: <b>' +
      esc(MISSING.join(', ')) + '</b> — die velden blijven leeg.<br>' +
      'Oorzaak: de Admin-upload filtert kolommen weg (KEEP_COLS in <code>shared/expediting-data.js</code>). ' +
      'Upload de bedrijfsbrede lijst opnieuw via <b>Admin</b> met de bijgewerkte <code>expediting-data.js</code>, ' +
      'download de nieuwe <code>expediting-data.json</code> en commit die.</div>';
  }

  // ── Sub Project ID selectie — zelfde opmaak als de Itemlijst-Validator ───
  function injectSPStyle() {
    if (document.getElementById('val-sp-style')) return;
    const s = document.createElement('style'); s.id = 'val-sp-style';
    s.textContent =
      '.val-sp-panel{margin:.6rem 0 0;padding:1rem 1.1rem;background:var(--navy-mid,#0F2040);border:1px solid var(--steel,#1e3a6e);border-radius:12px;font-family:var(--mono,monospace)}' +
      '.val-sp-head{display:flex;align-items:center;gap:.5rem;font-weight:700;font-size:.8rem;letter-spacing:.03em;color:var(--white,#F0F4FA)}' +
      '.val-sp-search{width:100%;box-sizing:border-box;padding:.55rem .8rem;border-radius:8px;border:1px solid var(--steel,#1e3a6e);background:var(--navy,#0A1628);color:var(--white,#F0F4FA);font-family:inherit;font-size:.8rem;outline:none;margin-top:.8rem}' +
      '.val-sp-search:focus{border-color:var(--teal,#00B4D8)}' +
      '.val-sp-actions{display:flex;align-items:center;gap:.5rem;margin:.7rem 0 .4rem}' +
      '.val-sp-btn{font-family:inherit;font-size:.7rem;font-weight:700;letter-spacing:.05em;text-transform:uppercase;padding:.4rem .8rem;border-radius:6px;border:1px solid var(--steel,#1e3a6e);background:transparent;color:var(--white,#F0F4FA);cursor:pointer;transition:border-color .15s,background .15s}' +
      '.val-sp-btn:hover{border-color:var(--teal,#00B4D8);background:rgba(0,180,216,.1)}' +
      '.val-sp-count{margin-left:auto;font-size:.72rem;color:var(--teal,#00B4D8)}' +
      '.val-sp-list{max-height:210px;overflow-y:auto;border:1px solid var(--steel,#1e3a6e);border-radius:8px;background:rgba(10,22,40,.4)}' +
      '.val-sp-list::-webkit-scrollbar{width:10px}.val-sp-list::-webkit-scrollbar-thumb{background:var(--steel,#1e3a6e);border-radius:5px}' +
      '.val-sp-item{display:flex;align-items:center;gap:.6rem;padding:.4rem .8rem;cursor:pointer;font-size:.76rem;color:var(--white,#F0F4FA);border-bottom:1px solid rgba(30,58,110,.25)}' +
      '.val-sp-item:hover{background:rgba(0,180,216,.07)}' +
      '.val-sp-item input{accent-color:var(--teal,#00B4D8);width:14px;height:14px;cursor:pointer}' +
      '.val-sp-item .val-sp-n{margin-left:auto;color:var(--grey,#8FA3BF);font-size:.72rem}' +
      '.val-sp-item.hidden{display:none}';
    document.head.appendChild(s);
  }

  function buildSubProjectPanel() {
    const panel = $('lio-sp-panel');
    const counts = new Map();
    for (const r of ROWS) {
      const id = trim(val(r, 'sp')); if (!id) continue;
      counts.set(id, (counts.get(id) || 0) + 1);
    }
    const ids = [...counts.keys()].sort((a, b) => a.localeCompare(b, 'nl', { numeric: true }));
    if (!ids.length) { panel.innerHTML = ''; return; }

    injectSPStyle();
    SPSEL = new Set();                       // standaard: niets geselecteerd (= alle projecten)
    panel.className = 'val-sp-panel';
    panel.innerHTML =
      '<div class="val-sp-head">📋 Bedrijfsbreed Expediten — selecteer Sub Project ID</div>' +
      '<input class="val-sp-search" id="val-sp-search" placeholder="🔍 Zoek Sub Project ID…" autocomplete="off">' +
      '<div class="val-sp-actions">' +
      '<button class="val-sp-btn" id="val-sp-all" type="button">Alles</button>' +
      '<button class="val-sp-btn" id="val-sp-none" type="button">Wis</button>' +
      '<span class="val-sp-count" id="val-sp-count">0 geselecteerd</span>' +
      '</div>' +
      '<div class="val-sp-list" id="val-sp-list">' +
      ids.map(id => '<label class="val-sp-item" data-id="' + esc(id) + '">' +
        '<input type="checkbox" value="' + esc(id) + '">' +
        '<span class="val-sp-id">' + esc(id) + '</span>' +
        '<span class="val-sp-n">' + counts.get(id) + '</span></label>').join('') +
      '</div>';

    const apply = () => {
      const c = $('val-sp-count');
      if (c) c.textContent = SPSEL.size + ' geselecteerd';
      render();
    };
    panel.querySelector('#val-sp-list').addEventListener('change', e => {
      if (e.target.type !== 'checkbox') return;
      if (e.target.checked) SPSEL.add(e.target.value); else SPSEL.delete(e.target.value);
      apply();
    });
    panel.querySelector('#val-sp-search').addEventListener('input', function () {
      const q = this.value.trim().toLowerCase();
      panel.querySelectorAll('.val-sp-item').forEach(it =>
        it.classList.toggle('hidden', !!q && !it.dataset.id.toLowerCase().includes(q)));
    });
    panel.querySelector('#val-sp-all').addEventListener('click', () => {
      panel.querySelectorAll('.val-sp-item:not(.hidden) input').forEach(cb => { cb.checked = true; SPSEL.add(cb.value); });
      apply();
    });
    panel.querySelector('#val-sp-none').addEventListener('click', () => {
      panel.querySelectorAll('.val-sp-item input').forEach(cb => cb.checked = false);
      SPSEL.clear(); apply();
    });
  }

  // ── tabbladen ────────────────────────────────────────────────────────────
  function wireTabs() {
    document.querySelectorAll('.lio-tab').forEach(b =>
      b.addEventListener('click', () => {
        document.querySelectorAll('.lio-tab').forEach(x => x.classList.remove('active'));
        b.classList.add('active'); TAB = b.dataset.tab; render();
      }));
  }

  // ── filters per tabblad ──────────────────────────────────────────────────
  const inScope = (r) => !SPSEL.size || SPSEL.has(trim(val(r, 'sp')));

  function rowsFor(tab) {
    const base = ROWS.filter(inScope);
    if (tab === 'large') {
      return base.filter(r => { const w = num(val(r, 'totW')); return w != null && w >= THRESHOLD; });
    }
    if (tab === 'late') {
      return base.filter(r => trim(val(r, 'delStatus')).toLowerCase() === 'late');
    }
    if (tab === 'beta') return base.filter(r => betaOf(r).cnt > 0);
    return base.filter(r => trim(val(r, 'fatDate')) || trim(val(r, 'fatLoc')));
  }
  function delay(r) {
    const w = toDate(val(r, 'wanted')), p = toDate(val(r, 'planned'));
    return (w && p) ? Math.round((p - w) / 86400000) : null;
  }

  // ── kolomdefinities per tabblad ──────────────────────────────────────────
  const COMMON_TAIL = [
    { h: 'Delivery Address', k: 'addr', f: r => flat(val(r, 'address')) },
    { h: 'Incoterm',         k: 'inco', f: r => flat(val(r, 'terms')) },
    { h: 'Last Expedited',   k: 'lexp', f: r => fmtDate(val(r, 'lastExp')),
      d: r => toDate(val(r, 'lastExp')), cls: r => lastExpCls(r) },
  ];
  function lastExpCls(r) {
    const d = toDate(val(r, 'lastExp'));
    if (!d) return 'lio-bad';                                   // nooit opgevolgd
    const days = Math.round((startOfToday() - d) / 86400000);
    return days > 60 ? 'lio-warn-c' : '';
  }
  // ── FAT: leadtime-regels uit de oude FAT-Overview ────────────────────────
  const LEAD_MIN = 28, LEAD_MAX = 35, INVITE_DAYS = 14;
  const addDays = (d, n) => d ? new Date(d.getTime() + n * 86400000) : null;
  function calcLead(fat, delivery) {
    if (!fat || !delivery) return { key: 'missing', label: 'Data incompleet', gap: null };
    const gap = Math.round((delivery - fat) / 86400000);
    if (gap >= LEAD_MIN && gap <= LEAD_MAX) return { key: 'ok', label: 'OK (' + gap + ' dgn)', gap };
    if (gap < LEAD_MIN) return { key: 'short', label: 'Te kort (' + gap + ' dgn)', gap };
    return { key: 'long', label: 'Te vroeg (' + gap + ' dgn)', gap };
  }
  const leadOf = (r) => calcLead(toDate(val(r, 'fatDate')), toDate(val(r, 'planned')));
  // Protocol goedgekeurd = er staat een datum in AG
  const protoOk = (r) => toDate(val(r, 'protoDate')) ? 'Y' : '';

  // ── FAT-invulling (handmatig aan te vullen, gedeeld via fat-info.json) ──────
  // Gedeelde waarden staan in fat-info.json naast deze pagina (in de repo).
  // Eigen wijzigingen blijven in de browser (DRAFT) tot je ze opslaat en het
  // gedownloade bestand in de repo zet. Regel-sleutel: PO | System No | Part No.
  const FAT_FILE = 'fat-info.json', DRAFT_LS = 'lio-fat-draft-v1';
  const FAT_FIELDS = ['protRecv', 'protAppr', 'ihcPres', 'cliPres', 'locAddr'];
  let COMMITTED = {}, COMMITTED_META = null, DRAFT = {}, INFO_VER = 0;
  try { DRAFT = JSON.parse(localStorage.getItem(DRAFT_LS) || '{}') || {}; } catch (e) { DRAFT = {}; }
  const saveDraft = () => { try { localStorage.setItem(DRAFT_LS, JSON.stringify(DRAFT)); } catch (e) {} };
  const fatKey = (r) => [trim(val(r, 'orderNo')), flat(val(r, 'sysNo')), trim(val(r, 'part'))].join('|');
  function fi(r, f) {
    const k = fatKey(r);
    if (DRAFT[k] && f in DRAFT[k]) return DRAFT[k][f];
    return (COMMITTED[k] && COMMITTED[k][f]) || '';
  }
  const isDirty = (r, f) => { const k = fatKey(r); return !!(DRAFT[k] && f in DRAFT[k] && DRAFT[k][f] !== ((COMMITTED[k] && COMMITTED[k][f]) || '')); };
  function setFi(k, f, v) {
    const base = (COMMITTED[k] && COMMITTED[k][f]) || '';
    if (v === base) { if (DRAFT[k]) { delete DRAFT[k][f]; if (!Object.keys(DRAFT[k]).length) delete DRAFT[k]; } }
    else { (DRAFT[k] = DRAFT[k] || {})[f] = v; }
    INFO_VER++; saveDraft(); updateFatStatus();
  }
  const dirtyCount = () => Object.keys(DRAFT).reduce((n, k) => n + Object.keys(DRAFT[k]).length, 0);
  function updateFatStatus() {
    const el = $('lio-fat-status'); if (!el) return;
    const n = dirtyCount();
    el.textContent = n ? n + ' wijziging(en) nog niet gedeeld — sla op en zet fat-info.json in de repo' :
      (COMMITTED_META && COMMITTED_META.saved ? 'Gedeelde gegevens van ' + fmtDate(COMMITTED_META.saved) : 'Geen gedeelde gegevens geladen');
    el.className = 'lio-fat-status' + (n ? ' lio-fat-dirty' : '');
  }
  async function loadFatInfo() {
    try {
      const resp = await fetch(FAT_FILE + '?v=' + Date.now(), { cache: 'no-store' });
      if (!resp.ok) throw new Error('HTTP ' + resp.status);
      const obj = await resp.json();
      COMMITTED = (obj && obj.items) || {}; COMMITTED_META = (obj && obj.meta) || null;
      // eigen wijzigingen die inmiddels gelijk zijn aan de gedeelde waarde opruimen
      for (const k of Object.keys(DRAFT)) { for (const f of Object.keys(DRAFT[k])) {
        if (DRAFT[k][f] === ((COMMITTED[k] && COMMITTED[k][f]) || '')) delete DRAFT[k][f]; }
        if (!Object.keys(DRAFT[k]).length) delete DRAFT[k]; }
      saveDraft();
    } catch (e) { COMMITTED = {}; COMMITTED_META = null; }   // nog geen bestand = normaal
    INFO_VER++; updateFatStatus(); if (TAB === 'fat' || TAB === 'beta') render();
  }
  function effectiveInfo() {
    const out = {}; const keys = new Set([...Object.keys(COMMITTED), ...Object.keys(DRAFT)]);
    for (const k of keys) {
      const o = {}; for (const f of FAT_FIELDS) {
        const v = (DRAFT[k] && f in DRAFT[k]) ? DRAFT[k][f] : ((COMMITTED[k] && COMMITTED[k][f]) || '');
        if (v) o[f] = v; }
      if (Object.keys(o).length) out[k] = o;
    }
    return out;
  }
  function downloadFatInfo() {
    const payload = { meta: { saved: new Date().toISOString(), version: 1,
        note: 'Handmatig aangevulde FAT-gegevens (Large Item Overview). Zet dit bestand als fat-info.json naast index.html.' },
      items: effectiveInfo() };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob); const a = document.createElement('a');
    a.href = url; a.download = FAT_FILE; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    const el = $('lio-fat-status');
    if (el) { el.textContent = FAT_FILE + ' gedownload (' + Object.keys(payload.items).length + ' regels) — zet het in de repo naast index.html en commit het.'; el.className = 'lio-fat-status'; }
  }
  function importFatInfo(file) {
    const rd = new FileReader();
    rd.onload = () => { try {
      const obj = JSON.parse(rd.result); const items = (obj && obj.items) || {};
      for (const k of Object.keys(items)) for (const f of FAT_FIELDS) {
        if (f in items[k]) { const v = items[k][f] || ''; const base = (COMMITTED[k] && COMMITTED[k][f]) || '';
          if (v !== base) (DRAFT[k] = DRAFT[k] || {})[f] = v; } }
      INFO_VER++; saveDraft(); render();
    } catch (e) { alert('Kon ' + file.name + ' niet lezen: ' + e.message); } };
    rd.readAsText(file);
  }
  function resetDraft() {
    if (!dirtyCount() || !confirm('Alle eigen, nog niet gedeelde wijzigingen wissen?')) return;
    DRAFT = {}; saveDraft(); INFO_VER++; render();
  }
  // protocol goedgekeurd: handmatige keuze wint; anders afgeleid uit FAT Supplier Protocol Date
  const protApprEff = (r) => fi(r, 'protAppr') || (toDate(val(r, 'protoDate')) ? 'Ja' : '');
  function selHtml(r, f, opts, derived) {
    const cur = fi(r, f), shown = cur || derived || '';
    const dirty = isDirty(r, f) ? ' lio-in-dirty' : '', dim = (!cur && derived) ? ' lio-in-derived' : '';
    return '<select class="lio-in' + dirty + dim + '" data-k="' + esc(fatKey(r)) + '" data-f="' + f + '">' +
      ['', ...opts].map(o => '<option value="' + o + '"' + (o === shown ? ' selected' : '') + '>' + (o === '' ? '—' : o) + '</option>').join('') + '</select>';
  }
  const YN = ['Ja', 'Nee'];

  // ── BETA: kritiekheid = gewicht + FAT + additional requirements ──────────
  // Per aanwezige factor punten (requirements weegt het zwaarst, de drie liggen dicht
  // bij elkaar). 2 van de 3 factoren krijgt een bonus, 3 van de 3 een grotere bonus.
  // Daarna: score >= RED rood, >= ORANGE oranje, anders groen.
  const BETA = {
    pts:   { req: 35, weight: 33, fat: 32 },
    bonus: { 1: 0, 2: 8, 3: 20 },
    red: 70, orange: 40,
    weightFull: 10000,          // kg waarbij de gewichtsfactor volledig meetelt
    fatSoonDays: 30,            // FAT binnen zoveel dagen (of verstreken) = urgent
  };
  const NONE_TXT = new Set(['', '0', '-', '--', 'none', 'nvt', 'n/a', 'na', 'geen', 'no', 'nee']);
  function reqCount(r) {
    const v = val(r, 'addReq');
    if (typeof v === 'number') return isFinite(v) && v > 0 ? Math.round(v) : 0;
    const t = trim(v); if (NONE_TXT.has(t.toLowerCase())) return 0;
    if (/^\d+$/.test(t)) return parseInt(t, 10);
    return t.split(/[\r\n;|]+/).map(x => x.trim()).filter(x => x && !NONE_TXT.has(x.toLowerCase())).length;
  }
  const _betaMemo = new WeakMap();
  function betaOf(r) {
    if (_betaMemo.has(r) && _betaMemo.get(r).thr === THRESHOLD && _betaMemo.get(r).ver === INFO_VER) return _betaMemo.get(r);
    const w = num(val(r, 'totW')), n = reqCount(r);
    const fd = toDate(val(r, 'fatDate')), hasFat = !!(fd || trim(val(r, 'fatLoc')));
    const hasW = w != null && w >= THRESHOLD, hasR = n > 0;
    const f = [];
    let pts = 0;
    if (hasR) { pts += BETA.pts.req * (n >= 3 ? 1 : n === 2 ? 0.8 : 0.6); f.push('Req ' + n); }
    if (hasW) {
      const span = Math.max(1, BETA.weightFull - THRESHOLD);
      pts += BETA.pts.weight * (0.6 + 0.4 * Math.min(1, (w - THRESHOLD) / span)); f.push('Gewicht');
    }
    if (hasFat) {
      let i = 0.6;
      if (protApprEff(r) !== 'Ja') i += 0.2;                                       // protocol nog niet goedgekeurd
      if (fd && Math.round((fd - startOfToday()) / 86400000) <= BETA.fatSoonDays) i += 0.2;   // FAT dichtbij / verstreken
      pts += BETA.pts.fat * i; f.push('FAT');
    }
    const cnt = f.length;
    const score = cnt ? Math.round(pts + BETA.bonus[cnt]) : 0;
    const level = !cnt ? 'none' : score >= BETA.red ? 'red' : score >= BETA.orange ? 'orange' : 'green';
    const out = { score, level, cnt, factors: f, n, w, thr: THRESHOLD, ver: INFO_VER };
    _betaMemo.set(r, out); return out;
  }
  const LEVEL = { red: ['Rood', 'lio-bad'], orange: ['Oranje', 'lio-warn-c'], green: ['Groen', 'lio-ok'], none: ['', ''] };

  const DEFS = {
    large: [
      { h: 'PO No',        k: 'po',   f: r => trim(val(r, 'po')) },
      { h: 'Sub Project',  k: 'sp',   f: r => trim(val(r, 'sp')) },
      { h: 'Supplier',     k: 'sup',  f: r => flat(val(r, 'supplier')) },
      { h: 'Part No',      k: 'part', f: r => trim(val(r, 'part')) },
      { h: 'Omschrijving', k: 'desc', f: r => flat(val(r, 'desc')) },
      { h: 'Qty',          k: 'qty',  f: r => fmtNum(num(val(r, 'qty'))), num: true, n: r => num(val(r, 'qty')) },
      { h: 'Net Weight',   k: 'netW', f: r => fmtNum(num(val(r, 'netW'))), num: true, n: r => num(val(r, 'netW')) },
      { h: 'Totaal gewicht', k: 'totW', f: r => fmtNum(num(val(r, 'totW'))), num: true,
        n: r => num(val(r, 'totW')), cls: () => 'lio-strong' },
      { h: 'UoM',          k: 'uom',  f: r => trim(val(r, 'wUoM')) },
      { h: 'Planned',      k: 'pdd',  f: r => fmtDate(val(r, 'planned')), d: r => toDate(val(r, 'planned')) },
      ...COMMON_TAIL,
    ],
    late: [
      { h: 'PO No',        k: 'po',   f: r => trim(val(r, 'po')) },
      { h: 'Sub Project',  k: 'sp',   f: r => trim(val(r, 'sp')) },
      { h: 'Supplier',     k: 'sup',  f: r => flat(val(r, 'supplier')) },
      { h: 'Part No',      k: 'part', f: r => trim(val(r, 'part')) },
      { h: 'Omschrijving', k: 'desc', f: r => flat(val(r, 'desc')) },
      { h: 'Wanted',       k: 'want', f: r => fmtDate(val(r, 'wanted')), d: r => toDate(val(r, 'wanted')) },
      { h: 'Planned',      k: 'pdd',  f: r => fmtDate(val(r, 'planned')), d: r => toDate(val(r, 'planned')) },
      { h: 'Dagen te laat', k: 'off', f: r => { const d = delay(r); return d == null ? '' : String(d); },
        num: true, n: r => delay(r),
        cls: r => { const d = delay(r); return d > 30 ? 'lio-bad' : d > 0 ? 'lio-warn-c' : ''; } },
      ...COMMON_TAIL,
    ],
    beta: [
      { h: 'Score',        k: 'bscore', f: r => String(betaOf(r).score), num: true, n: r => betaOf(r).score,
        cls: r => LEVEL[betaOf(r).level][1] + ' lio-strong', title: 'Kritiekheidsscore (0-100+): gewicht + FAT + additional requirements' },
      { h: 'Risico',       k: 'blevel', f: r => LEVEL[betaOf(r).level][0], n: r => betaOf(r).score,
        cls: r => LEVEL[betaOf(r).level][1] },
      { h: 'Factoren',     k: 'bfac',  f: r => betaOf(r).factors.join(' · ') + ' (' + betaOf(r).cnt + '/3)', n: r => betaOf(r).cnt * 1000 + betaOf(r).score },
      { h: 'PO No',        k: 'po',   f: r => trim(val(r, 'po')) },
      { h: 'Sub Project',  k: 'sp',   f: r => trim(val(r, 'sp')) },
      { h: 'Supplier',     k: 'sup',  f: r => flat(val(r, 'supplier')) },
      { h: 'Part No',      k: 'part', f: r => trim(val(r, 'part')) },
      { h: 'Omschrijving', k: 'desc', f: r => flat(val(r, 'desc')) },
      { h: 'Totaal gewicht', k: 'totW', f: r => fmtNum(num(val(r, 'totW'))), num: true, n: r => num(val(r, 'totW')) },
      { h: 'Add. req.',    k: 'breq', f: r => { const n = reqCount(r); return n ? String(n) : ''; }, num: true, n: r => reqCount(r) },
      { h: 'FAT date',     k: 'fdate', f: r => fmtDate(val(r, 'fatDate')), d: r => toDate(val(r, 'fatDate')),
        cls: r => { const d = toDate(val(r, 'fatDate')); if (!d) return '';
          const x = Math.round((d - startOfToday()) / 86400000); return x < 0 ? 'lio-bad' : x <= 30 ? 'lio-warn-c' : ''; } },
      { h: 'Protocol',     k: 'proto', f: r => protApprEff(r) === 'Ja' ? 'Y' : '', cls: r => protApprEff(r) === 'Ja' ? 'lio-ok' : '' },
      { h: 'Planned',      k: 'pdd',  f: r => fmtDate(val(r, 'planned')), d: r => toDate(val(r, 'planned')) },
      ...COMMON_TAIL,
    ],
    fat: [
      { h: 'System No',        k: 'sysNo',  f: r => flat(val(r, 'sysNo')) },
      { h: 'PO Number',        k: 'po',     f: r => flat(val(r, 'orderNo')) },
      { h: 'Name',             k: 'name',   f: r => flat(val(r, 'refDesc')) },
      { h: 'Leverancier',      k: 'sup',    f: r => flat(val(r, 'supplier')) },
      { h: 'FAT Location',     k: 'floc',   f: r => flat(val(r, 'fatLoc')) },
      { h: 'Client Required',  k: 'creq',   f: () => '', title: 'Nog geen bronkolom in de bedrijfsbrede lijst (Customer Attendance ID)' },
      { h: 'Expected FAT Date', k: 'fdate', f: r => fmtDate(val(r, 'fatDate')), d: r => toDate(val(r, 'fatDate')),
        cls: r => { const d = toDate(val(r, 'fatDate')); if (!d) return '';
          const x = Math.round((d - startOfToday()) / 86400000);
          return x < 0 ? 'lio-bad' : x <= 30 ? 'lio-warn-c' : ''; } },
      // ── handmatig in te vullen (opslaan in fat-info.json) ──
      { h: 'Protocol ontvangen?',  k: 'protRecv', f: r => fi(r, 'protRecv'), edit: true, cls: r => fi(r, 'protRecv') === 'Ja' ? 'lio-ok' : '',
        html: r => selHtml(r, 'protRecv', YN), title: 'Handmatig — is het FAT-protocol ontvangen?' },
      { h: 'Protocol goedgekeurd', k: 'protAppr', f: protApprEff, edit: true, cls: r => protApprEff(r) === 'Ja' ? 'lio-ok' : '',
        html: r => selHtml(r, 'protAppr', YN, toDate(val(r, 'protoDate')) ? 'Ja' : ''),
        title: 'Handmatige keuze wint; leeg = afgeleid uit FAT Supplier Protocol Date (grijs)' },
      { h: 'Aanwezig IHC',     k: 'ihcPres', f: r => fi(r, 'ihcPres'), edit: true, html: r => selHtml(r, 'ihcPres', YN) },
      { h: 'Aanwezig klant',   k: 'cliPres', f: r => fi(r, 'cliPres'), edit: true, html: r => selHtml(r, 'cliPres', YN) },
      { h: 'Locatie / adres',  k: 'locAddr', f: r => fi(r, 'locAddr') || flat(val(r, 'fatLoc')), edit: true,
        html: r => '<input type="text" class="lio-in lio-in-txt' + (isDirty(r, 'locAddr') ? ' lio-in-dirty' : '') + '" data-k="' + esc(fatKey(r)) + '" data-f="locAddr" value="' +
          esc(fi(r, 'locAddr')) + '" placeholder="' + esc(flat(val(r, 'fatLoc'))) + '">',
        title: 'Handmatig adres van de FAT-locatie; leeg = FAT Location uit de lijst (placeholder)' },
      { h: 'Client invite before', k: 'inv', f: r => fmtDate(addDays(toDate(val(r, 'fatDate')), -INVITE_DAYS)),
        d: r => addDays(toDate(val(r, 'fatDate')), -INVITE_DAYS), cls: () => 'lio-dim',
        title: 'FAT-datum − 14 dagen. Zolang Customer Attendance ID ontbreekt, wordt dit voor élke regel getoond.' },
      { h: 'Responsible Name', k: 'resp',   f: r => flat(val(r, 'engineer')) },
      { h: 'Planned Delivery Date (exw)', k: 'pdd', f: r => fmtDate(val(r, 'planned')), d: r => toDate(val(r, 'planned')),
        cls: r => leadOf(r).key === 'short' || leadOf(r).key === 'long' ? 'lio-bad' : '' },
      { h: 'Remarks',          k: 'rem',    f: r => leadOf(r).label,
        cls: r => { const k = leadOf(r).key; return k === 'ok' ? 'lio-ok' : k === 'missing' ? 'lio-dim' : 'lio-warn-c'; } },
      { h: 'Sub Project',      k: 'sp',     f: r => flat(val(r, 'sp')) },
      { h: 'Last Expedited',   k: 'lexp',   f: r => fmtDate(val(r, 'lastExp')), d: r => toDate(val(r, 'lastExp')),
        cls: r => lastExpCls(r) },
    ],
  };

  // ── sorteren (zelfde gedrag als Late Items in de PO-Matcher) ─────────────
  const SORT = { beta: { col: 'bscore', dir: 'desc' }, large: { col: 'totW', dir: 'desc' }, late: { col: 'off', dir: 'desc' }, fat: { col: 'fdate', dir: 'asc' } };
  window.lioSort = function (key) {
    const s = SORT[TAB];
    if (s.col === key) s.dir = (s.dir === 'asc') ? 'desc' : 'asc';
    else { s.col = key; s.dir = 'asc'; }
    render();
  };
  function sortRows(rows, defs) {
    const s = SORT[TAB];
    const def = defs.find(d => d.k === s.col);
    if (!def) return rows;
    const keyOf = (r) => def.n ? def.n(r) : def.d ? (def.d(r) ? def.d(r).getTime() : null) : def.f(r);
    const dir = (s.dir === 'desc') ? -1 : 1;
    return rows.slice().sort((a, b) => {
      const av = keyOf(a), bv = keyOf(b);
      const aB = (av === null || av === undefined || av === '');
      const bB = (bv === null || bv === undefined || bv === '');
      if (aB && bB) return 0;
      if (aB) return 1;                       // lege waarden altijd onderaan
      if (bB) return -1;
      if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * dir;
      return String(av).localeCompare(String(bv), 'nl', { numeric: true }) * dir;
    });
  }

  // ── zoekbalk: combineert vrij over ALLE kolommen van het huidige tabblad ──
  // Elk los woord moet ergens voorkomen (in willekeurig welke kolom) — zo
  // vind je bv. "Alfa Laval YN1320" ook als leverancier en project in
  // verschillende kolommen staan.
  let SEARCHTERM = '';
  window.lioSearch = function (v) { SEARCHTERM = v; render(); };
  function applySearch(rows, defs) {
    const terms = SEARCHTERM.trim().toLowerCase().split(/\s+/).filter(Boolean);
    if (!terms.length) return rows;
    return rows.filter(r => {
      const hay = defs.map(d => String(d.f(r) ?? '')).join(' \u241F ').toLowerCase();
      return terms.every(t => hay.includes(t));
    });
  }

  // ── rendering ────────────────────────────────────────────────────────────
  function render() {
    if (!ROWS.length) return;
    renderTools();
    const miss = NEEDED[TAB].filter(k => C[k] < 0);
    const defs = DEFS[TAB];
    const rows = miss.length ? [] : sortRows(applySearch(rowsFor(TAB), defs), defs);

    const s = SORT[TAB];
    $('lio-thead').innerHTML = '<tr>' + defs.map(d => {
      const on = d.k === s.col;
      const arrow = on ? (s.dir === 'asc' ? ' \u25B2' : ' \u25BC') : '';
      const cls = ['lio-sortable', d.num ? 'lio-num' : '', on ? 'lio-sort-active' : ''].filter(Boolean).join(' ');
      return '<th class="' + cls + '" onclick="lioSort(\'' + d.k + '\')"' +
        ' title="' + esc(d.title || ('Sorteer op ' + d.h)) + '">' + esc(d.h) + arrow + '</th>';
    }).join('') + '</tr>';

    if (miss.length) {
      $('lio-summary').innerHTML = '<span class="lio-bad-txt">Dit overzicht kan niet getoond worden: ' +
        'kolom(men) <b>' + esc(miss.map(k => MAP[k][0][0]).join(', ')) + '</b> ontbreken in de geladen lijst.</span>';
      $('lio-tbody').innerHTML = '';
      return;
    }
    $('lio-summary').innerHTML = summaryFor(TAB, rows);
    $('lio-tbody').innerHTML = rows.length
      ? rows.map(r => '<tr>' + defs.map(d => {
          const cls = [d.num ? 'lio-num' : '', d.cls ? d.cls(r) : ''].filter(Boolean).join(' ');
          const v = d.f(r);
          if (d.html) return '<td' + (cls ? ' class="' + cls + '"' : '') + '>' + d.html(r) + '</td>';
          return '<td' + (cls ? ' class="' + cls + '"' : '') + (v ? ' title="' + esc(v) + '"' : '') +
                 '>' + esc(v) + '</td>';
        }).join('') + '</tr>').join('')
      : '<tr><td colspan="' + defs.length + '" class="lio-empty">Geen regels binnen deze selectie.</td></tr>';
  }

  function renderTools() {
    const el = $('lio-tools');
    el.innerHTML = (TAB === 'large' || TAB === 'beta')
      ? '<label class="lio-thr">Vanaf <input type="number" id="lio-threshold" value="' + THRESHOLD +
        '" min="0" step="100"> kg totaalgewicht</label>'
      : TAB === 'fat'
      ? '<button class="btn-val" id="lio-fat-save" type="button">💾 Opslaan (fat-info.json)</button> ' +
        '<button class="btn-val" id="lio-fat-load" type="button">📂 Laden…</button> ' +
        '<button class="btn-val" id="lio-fat-reset" type="button">↺ Eigen wijzigingen wissen</button> ' +
        '<input type="file" id="lio-fat-file" accept=".json,application/json" style="display:none">' +
        '<span class="lio-fat-status" id="lio-fat-status"></span>' : '';
    if (TAB === 'fat') {
      $('lio-fat-save').addEventListener('click', downloadFatInfo);
      $('lio-fat-load').addEventListener('click', () => $('lio-fat-file').click());
      $('lio-fat-file').addEventListener('change', e => { const f = e.target.files && e.target.files[0]; e.target.value = ''; if (f) importFatInfo(f); });
      $('lio-fat-reset').addEventListener('click', resetDraft);
      updateFatStatus();
    }
    const inp = $('lio-threshold');
    if (inp) inp.addEventListener('change', () => {
      const v = parseFloat(inp.value); THRESHOLD = isNaN(v) ? 0 : v; render();
    });
  }

  function summaryFor(tab, rows) {
    if (tab === 'large') {
      const tot = rows.reduce((s, r) => s + (num(val(r, 'totW')) || 0), 0);
      return '<b>' + rows.length + '</b> large items (≥ ' + fmtNum(THRESHOLD) + ' kg) · ' +
        'totaal <b>' + fmtNum(tot) + '</b> kg';
    }
    if (tab === 'beta') {
      const c = { red: 0, orange: 0, green: 0 }; rows.forEach(r => c[betaOf(r).level]++);
      const c3 = rows.filter(r => betaOf(r).cnt === 3).length;
      return '<b>' + rows.length + '</b> items met minimaal 1 factor · ' +
        '<span class="lio-bad-txt"><b>' + c.red + '</b> rood</span> · <b>' + c.orange + '</b> oranje · <b>' + c.green + '</b> groen · ' +
        '<b>' + c3 + '</b> met alle 3 factoren' +
        (C.addReq < 0 ? ' <br><span class="lio-bad-txt">⚠ Kolom "Additional Requirements" zit niet in de geladen lijst — requirements tellen nu als 0.</span>' : '') +
        '<br><span class="lio-dim">Factoren: gewicht ≥ ' + fmtNum(THRESHOLD) + ' kg · FAT aanwezig · ≥ 1 additional requirement. ' +
        'Requirements wegen het zwaarst; 2 van de 3 weegt minder dan 3 van de 3.</span>';
    }
    if (tab === 'late') {
      const d = rows.map(delay).filter(x => x != null && x > 0);
      const avg = d.length ? Math.round(d.reduce((s, x) => s + x, 0) / d.length) : 0;
      const crit = rows.filter(r => (delay(r) || 0) > 30).length;
      return '<b>' + rows.length + '</b> late items · gemiddeld <b>' + avg + '</b> dagen te laat · ' +
        '<b>' + crit + '</b> kritiek (> 30 dagen)';
    }
    const t = startOfToday();
    const over = rows.filter(r => { const d = toDate(val(r, 'fatDate')); return d && d < t; }).length;
    const soon = rows.filter(r => { const d = toDate(val(r, 'fatDate')); if (!d) return false;
      const x = Math.round((d - t) / 86400000); return x >= 0 && x <= 30; }).length;
    return '<b>' + rows.length + '</b> FAT-items · <b>' + over + '</b> datum verstreken · ' +
      '<b>' + soon + '</b> binnen 30 dagen';
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();