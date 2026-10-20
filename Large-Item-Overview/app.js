/* Large Item Overview — Royal IHC
 * Toont Large Items, Late Items, FAT Overview en BETA vanuit de BEDRIJFSBREDE
 * expediting-lijst. Geen moederlijst nodig. De FAT Overview komt uit fat-list.json
 * (tabblad FAT van de bedrijfsbrede Excel, via de Admin-pagina).
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
    line:      [['Line No'], 'C'],
    sysNo:     [['Unified Reference Code'], 'M'],
    refDesc:   [['Unified Ref Code Description'], 'N'],
    netW:      [['Net Weight'], 'AO'],
    totW:      [['Total Net Weight'], 'AP'],
    wUoM:      [['Weight UoM'], 'AQ'],
    // BETA-tab: aantal additional requirements. Optioneel (geen waarschuwing als hij ontbreekt).
    addReq:    [['Additional Requirements', 'Additional Requirement', 'Additional requirements'], 'AG'],
  };
  const OPTIONAL = new Set(['addReq', 'line']);
  // Kolommen zónder welke een tabblad zinloos is. Overige kolommen mogen
  // ontbreken: die blijven leeg, met een melding erboven.
  const NEEDED = {
    large: ['totW'],
    late:  ['delStatus'],
    fat:   [],
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
      const t = e.target; if (!t.classList || !t.classList.contains('lio-in') || !t.dataset.k) return;
      let v = t.type === 'checkbox' ? (t.checked ? 'Ja' : 'Nee') : t.value.trim();
      const def = t.dataset.def || '';
      if (v === def) v = ''; else if (v === '') v = '-';             // gelijk aan IFS = geen eigen invoer; leeg = bewust leeggemaakt
      setFi(t.dataset.k, t.dataset.f, v);
      t.classList.add('lio-in-dirty');
      refreshRow(t.closest('tr'));
      if ($('lio-ifs-in') && $('lio-ifs-in').value) runIfs();
    });
    $('lio-tbody').addEventListener('click', async (e) => {
      const eb = e.target.closest && e.target.closest('button.lio-exp'); if (eb) { toggleDoc(eb); return; }
      const a = e.target.closest && e.target.closest('a.lio-ifs-x'); if (!a) return;
      e.preventDefault();
      const r = FATBY[a.closest('tr').dataset.k]; if (!r) return;
      const w = window.open('', '_blank');                         // direct openen (anders blokkeert de browser de pop-up)
      try { const url = await ifsKeysUrl(ifsKeys(r)); if (w) w.location.href = url; else window.location.href = url; }
      catch (err) { if (w) w.close(); alert(err.message); }
    });
    loadFatList();
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
      '.lio-fat-dirty{color:#E8B923;font-weight:700}' +
      '.lio-exp{background:none;border:1px solid var(--steel,#1e3a6e);border-radius:5px;color:var(--teal,#00B4D8);cursor:pointer;width:1.6rem;height:1.5rem;font-size:.8rem;transition:transform .2s}' +
      '.lio-exp.lio-open{transform:rotate(90deg)}' +
      '.lio-detail>td{background:var(--mid,#0F2040);padding:0 !important;white-space:normal !important}' +
      '.lio-slide{overflow:hidden;transition:max-height .25s ease}.lio-slide-in{padding:.7rem 1rem}' +
      '.lio-doc-head{font-size:.85rem;margin-bottom:.3rem}.lio-doc-sum{font-size:.78rem;margin:.2rem 0 .4rem}' +
      '.lio-chips{display:flex;flex-wrap:wrap;gap:.3rem;margin-bottom:.5rem}' +
      '.lio-chip{background:var(--navy,#0A1628);border:1px solid var(--steel,#1e3a6e);border-radius:10px;padding:.1rem .55rem;font-size:.7rem}' +
      '.lio-doc-grp{font-size:.74rem;color:var(--grey,#8FA3BF);margin:.5rem 0 .2rem;font-weight:700}' +
      '.lio-doc{border-collapse:collapse;width:100%;max-width:1200px;font-size:.72rem;margin-bottom:.4rem}.lio-doc th{text-align:left;color:var(--grey,#8FA3BF);font-weight:600;padding:.2rem .5rem;border-bottom:1px solid var(--steel,#1e3a6e)}' +
      '.lio-doc td{padding:.2rem .5rem;border-bottom:1px solid rgba(30,58,110,.5)}' +
      '.lio-cb{width:auto;accent-color:#00B4D8}.lio-in[type=date]{min-width:120px}' +
      '.lio-ifs{margin:.6rem 0;border:1px solid var(--steel,#1e3a6e);border-radius:8px;background:var(--mid,#0F2040);padding:.4rem .8rem}' +
      '.lio-ifs summary{cursor:pointer;font-weight:700;font-size:.8rem}.lio-ifs-body{padding:.5rem 0}' +
      '.lio-ifs-cols{display:grid;grid-template-columns:1fr 1fr;gap:.6rem;margin:.5rem 0}' +
      '.lio-ifs textarea{width:100%;min-height:150px;resize:vertical;font-family:\'JetBrains Mono\',monospace;font-size:.68rem;background:var(--navy,#0A1628);color:var(--white,#F0F4FA);border:1px solid var(--steel,#1e3a6e);border-radius:6px;padding:.4rem}';
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
    return fatRowsFor();
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
  // ── FAT: lijst uit fat-list.json + handmatige invulling ──────────────────
  // fat-list.json (shared/) komt van de Admin-pagina (tabblad FAT van de bedrijfsbrede Excel).
  // Handmatige invulling staat in fat-info.json (gedeeld, naast index.html) + de browser (DRAFT) en
  // kan als IFS-tekst (!IFS.COPYOBJECT) teruggeschreven worden. Regel-sleutel:
  // Order No | Line No | Release No | Requirement Id.
  const DRAFT_LS = 'lio-fat-draft-v2';
  const FAT_FIELDS = ['recv', 'appr', 'cliReq', 'ihc', 'loc', 'clsNR', 'cusNR', 'inv', 'note'];
  const FN = {
    order: 'Order No', line: 'Line No', rel: 'Release No', req: 'Requirement Id',
    recv: 'Date Received', due: 'Date Required', want: 'Wanted Delivery Date', vendor: 'Vendor Name',
    sys: 'Design Object ID', proj: 'Project ID', appr: 'Approved', resp: 'Responsible Name',
    desc: 'Part Description', aIhc: 'IHC Attendance ID', aCls: 'Class Attendance ID', aCus: 'Customer Attendance ID',
    addInfo: 'Additional Information', note: 'Note', inv: 'Invitation ID', meet: 'Meeting Point',
    tloc: 'Test Location', cNR: 'Class Not Required', uNR: 'Customer Not Required', type: 'Type Description',
    reqName: 'Requirement', srcType: 'Requirement Source', unif: 'Unified Reference Code', unifDesc: 'Unified Ref Code Description', sysDesc: 'Design Object Description',
  };
  let FATH = [], FATR = [], FATALL = [], ARSYS = new Map(), ARPO = new Map(), VIEWROWS = [], FC = {}, FATMETA = null, FATERR = '', FATIDX = new Map(), FATBY = {}, FATVIEW = [], FATFILTER = 'all';
  let COMMITTED = {}, COMMITTED_META = null, DRAFT = {}, INFO_VER = 0;
  try { DRAFT = JSON.parse(localStorage.getItem(DRAFT_LS) || '{}') || {}; } catch (e) { DRAFT = {}; }
  const saveDraft = () => { try { localStorage.setItem(DRAFT_LS, JSON.stringify(DRAFT)); } catch (e) {} };
  const fg = (r, k) => (FC[k] >= 0 && r[FC[k]] != null) ? r[FC[k]] : '';
  const fkey = (r) => (window.FatData ? FatData.key : (a, b, c, d) => [a, b, c, d].join('|'))(fg(r, 'order'), fg(r, 'line'), fg(r, 'rel'), fg(r, 'req'));
  const keepProj = (sub) => (window.ExpeditingCore && ExpeditingCore.keepProject) ? ExpeditingCore.keepProject(sub) : /^(YN|EN)/i.test(trim(sub));

  function setFat(headers, rows) {
    FATH = headers; FC = {}; FATR = []; FATALL = []; ARSYS = new Map(); ARPO = new Map(); FATBY = {}; FATIDX = new Map(); FATERR = '';
    for (const k in FN) FC[k] = headers.findIndex(h => trim(h).toLowerCase() === FN[k].toLowerCase());
    const miss = ['order', 'line', 'rel', 'req', 'due', 'sys', 'proj'].filter(k => FC[k] < 0).map(k => FN[k]);
    if (miss.length) { FATERR = 'kolom(men) ontbreken in fat-list.json: ' + miss.join(', '); return; }
    FATALL = rows.filter(r => keepProj(r[FC.proj]));        // alleen onze projecten (YN/EN), zoals de expediting-lijst
    // De FAT-tab toont alleen FAT-regels; zit er later ook andere requirements in het blad, dan blijven die buiten beeld
    FATR = FATALL.filter(r => FC.type < 0 || !trim(r[FC.type]) || /fat|factory acceptance/i.test(trim(r[FC.type])));
    // documentatie-index (alle requirements): per systeem (Unified Reference Code / Design Object ID) en per PO-regel
    for (const r of FATALL) {
      const sy = new Set([trim(fg(r, 'unif')), trim(fg(r, 'sys'))].filter(Boolean));
      for (const k of sy) { if (!ARSYS.has(k)) ARSYS.set(k, []); ARSYS.get(k).push(r); }
      const ol = trim(fg(r, 'order')) + '|' + trim(fg(r, 'line'));
      if (!ARPO.has(ol)) ARPO.set(ol, []); ARPO.get(ol).push(r);
    }
    for (const r of FATR) {
      FATBY[fkey(r)] = r;
      const lk = trim(fg(r, 'order')) + '|' + trim(fg(r, 'line'));
      if (!FATIDX.has(lk)) FATIDX.set(lk, []);
      FATIDX.get(lk).push(r);
    }
  }
  async function loadFatList() {
    try {
      if (!window.FatData) throw new Error('fat-data.js niet geladen');
      const o = await FatData.load(); FATMETA = o.meta || null; setFat(o.headers, o.rows);
    } catch (e) { FATR = []; FATERR = e.message; }
    INFO_VER++; if (TAB === 'fat' || TAB === 'beta') render();
  }

  // waarden uit IFS (standaard) en onze invulling (override)
  const pad2 = (n) => String(n).padStart(2, '0');
  const isoOf = (v) => { const d = toDate(v); return d ? d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) : ''; };
  const isTrue = (v) => v === true || /^(true|ja|yes|1)$/i.test(trim(v));
  function dflt(r, f) {
    switch (f) {
      case 'recv':   return isoOf(fg(r, 'recv'));
      case 'appr':   return isTrue(fg(r, 'appr')) ? 'Ja' : 'Nee';
      case 'cliReq': { const a = trim(fg(r, 'aCus')).toUpperCase(); return (a === 'W' || a === 'R') ? 'X' : ''; }
      case 'ihc':    return ((window.FatData && FatData.ihcFromAddInfo(fg(r, 'addInfo'))) || '') || flat(fg(r, 'resp'));
      case 'loc':    return flat(fg(r, 'tloc')) || flat(fg(r, 'meet'));
      case 'clsNR':  return isTrue(fg(r, 'cNR')) ? 'Ja' : 'Nee';
      case 'cusNR':  return isTrue(fg(r, 'uNR')) ? 'Ja' : 'Nee';
      case 'inv':    return trim(fg(r, 'inv'));
      case 'note':   return flat(fg(r, 'note'));
    }
    return '';
  }
  // Overschrijfwaarde: '' = niets ingevuld (→ IFS-waarde), '-' = bewust leeggemaakt
  const ovr = (k, f) => (DRAFT[k] && f in DRAFT[k]) ? DRAFT[k][f] : ((COMMITTED[k] && COMMITTED[k][f]) || '');
  function eff(r, f) { const o = ovr(fkey(r), f); if (o === '-') return ''; return o || dflt(r, f); }
  const isDirty = (r, f) => { const k = fkey(r); return !!(DRAFT[k] && f in DRAFT[k] && DRAFT[k][f] !== ((COMMITTED[k] && COMMITTED[k][f]) || '')); };
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
    el.textContent = (n ? n + ' eigen wijziging(en) — alleen zichtbaar in deze browser, de bron blijft IFS' : '') +
      (FATMETA && FATMETA.uploaded ? (n ? ' · ' : '') + 'Lijst van ' + fmtDate(FATMETA.uploaded) : '');
    el.className = 'lio-fat-status' + (n ? ' lio-fat-dirty' : '');
  }
  // Alles wat afwijkt van IFS (gedeeld + eigen). Waarden die inmiddels gelijk zijn aan IFS vallen weg.
  function effectiveInfo() {
    const out = {}; const keys = new Set([...Object.keys(COMMITTED), ...Object.keys(DRAFT)]);
    for (const k of keys) {
      const o = {}; for (const f of FAT_FIELDS) {
        const v = ovr(k, f); if (!v) continue;
        const r = FATBY[k];
        if (r) { const d = dflt(r, f); if (v === d || (v === '-' && d === '')) continue; }
        o[f] = v; }
      if (Object.keys(o).length) out[k] = o;
    }
    return out;
  }
  function resetDraft() {
    if (!dirtyCount() || !confirm('Alle eigen, nog niet gedeelde wijzigingen wissen?')) return;
    DRAFT = {}; saveDraft(); INFO_VER++; render();
  }

  // Link naar IFS (scherm CCmtPoPartLine). Sleutelvolgorde en bedrijf zijn hier in te stellen.
  const IFS_BASE = 'https://ihc.ifs.cloud:48080/client/runtime/Ifs.Fnd.Explorer.application';
  const IFS_COMPANY = '233';
  const IFS_KEY_ORDER = ['line', 'order', 'rel', 'req'];       // volgorde in key1, gescheiden met ^
  const IFS_FILL_GAPS = true;                                   // ook de nummers 1 t/m het hoogste REQ_ID van elke release meenemen, ook als ze niet in de export staan

  // Alle FAT-regels van één PO in IFS: IFS verwacht de sleutels (gescheiden met \x1e) als gzip + base64 in 'zipParam'.
  async function gzipB64(str) {
    const cs = new CompressionStream('gzip');
    const w = cs.writable.getWriter(); w.write(new TextEncoder().encode(str)); w.close();
    const buf = new Uint8Array(await new Response(cs.readable).arrayBuffer());
    let bin = ''; for (let i = 0; i < buf.length; i += 0x8000) bin += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
    return btoa(bin);
  }
  const hasAllReqs = () => FATALL.length > FATR.length;           // blad bevat ook niet-FAT-requirements: dan echte sleutels gebruiken
  const keyOf = (x) => IFS_KEY_ORDER.map(k => trim(fg(x, k))).join('^');
  // Bij een FAT met REQ_ID n bestaan op die Line/Release ook de requirements 1 t/m n-1 (terugval als het blad alleen FAT bevat)
  function lineKeys(r) {
    const n = Math.max(1, parseInt(trim(fg(r, 'req')), 10) || 1), out = [];
    for (let i = 1; i <= n; i++) out.push(IFS_KEY_ORDER.map(k => k === 'req' ? String(i) : trim(fg(r, k))).join('^'));
    return out;
  }
  // Alle additional requirements van één PO-regel (Order No + Line No, alle releases), gesorteerd op release en requirement
  function ifsKeys(r) {
    const o = trim(fg(r, 'order')), l = trim(fg(r, 'line')), same = (x) => trim(fg(x, 'order')) === o && trim(fg(x, 'line')) === l;
    if (hasAllReqs()) {
      const rows = FATALL.filter(same), byRel = new Map();
      for (const x of rows) { const rl = trim(fg(x, 'rel')); if (!byRel.has(rl)) byRel.set(rl, { r0: x, ids: new Set() }); byRel.get(rl).ids.add(parseInt(trim(fg(x, 'req')), 10) || 0); }
      const out = [];
      for (const [rl, g] of [...byRel.entries()].sort((a, b) => (num(a[0]) || 0) - (num(b[0]) || 0))) {
        if (IFS_FILL_GAPS) for (let i = 1, mx = Math.max(...g.ids); i <= mx; i++) g.ids.add(i);
        for (const id of [...g.ids].sort((a, b) => a - b)) if (id > 0) out.push(IFS_KEY_ORDER.map(k => k === 'req' ? String(id) : trim(fg(g.r0, k))).join('^'));
      }
      return out;
    }
    const set = new Set(); for (const x of FATR) if (same(x)) lineKeys(x).forEach(k => set.add(k));
    return [...set];
  }
  async function ifsKeysUrl(keys) {
    let q;
    if (keys.length <= 25) q = keys.map((k, i) => 'key' + (i + 1) + '=' + encodeURIComponent(k)).join('&');
    else if (window.CompressionStream) q = 'zipParam=' + encodeURIComponent(await gzipB64(keys.join('\x1e')));
    else throw new Error('Deze browser kan de lange IFS-link niet bouwen (CompressionStream ontbreekt) — gebruik een recente Edge of Chrome.');
    return IFS_BASE + '?url=' + encodeURIComponent('ifsapf:tbwCCmtPoPartLine?action=get&' + q + '&COMPANY=' + IFS_COMPANY);
  }

  // afgeleide kolommen
  const attLet = (r, k) => { const v = trim(fg(r, k)); return (v === '' || v === '-') ? '' : v; };
  const needIhc = (r) => !!(attLet(r, 'aIhc') || attLet(r, 'aCls') || attLet(r, 'aCus'));   // letter in IHC/Class/Klant = IHC moet erbij zijn
  const attStr = (r) => [trim(fg(r, 'aIhc')) || '·', trim(fg(r, 'aCls')) || '·', trim(fg(r, 'aCus')) || '·'].join(' / ');
  // Invitation ID aangemaakt = protocol binnen + goedgekeurd + klant uitgenodigd → administratie moet dat laten zien
  const admOf = (r) => !!(eff(r, 'inv') && (!eff(r, 'recv') || eff(r, 'appr') !== 'Ja'));
  const ST = [['Achterstallig', 'lio-bad'], ['Binnen 14 dgn', 'lio-warn-c'], ['Gepland', ''], ['Goedgekeurd', 'lio-ok']];
  function stOf(r) {
    if (eff(r, 'appr') === 'Ja') return 3;
    const d = toDate(fg(r, 'due')); if (!d) return 2;
    const x = Math.round((d - startOfToday()) / 86400000);
    return x < 0 ? 0 : x <= 14 ? 1 : 2;
  }
  const FILTERS = [['all', 'Alle FAT-regels'], ['adm', 'Administratie bijwerken'], ['need', 'IHC erbij vereist'], ['late', 'Achterstallig'],
    ['soon', 'Binnen 14 dagen'], ['norecv', 'Protocol niet ontvangen'], ['open', 'Nog niet goedgekeurd'], ['appr', 'Goedgekeurd']];
  function fatPass(r) {
    switch (FATFILTER) {
      case 'adm': return admOf(r);
      case 'need': return needIhc(r);
      case 'late': return stOf(r) === 0;
      case 'soon': return stOf(r) === 1;
      case 'norecv': return !eff(r, 'recv');
      case 'open': return stOf(r) !== 3;
      case 'appr': return stOf(r) === 3;
    }
    return true;
  }
  const fatRowsFor = () => FATR.filter(r => (!SPSEL.size || SPSEL.has(trim(fg(r, 'proj')))) && fatPass(r));

  // koppeling BETA ↔ FAT-lijst op Order No + Line No
  function fatLink(r) {
    if (!FATIDX.size || C.line < 0) return null;
    const rs = FATIDX.get(trim(val(r, 'orderNo')) + '|' + trim(val(r, 'line')));
    if (!rs) return null;
    let all = true, dueOpen = null, dueAny = null;
    for (const x of rs) {
      const d = toDate(fg(x, 'due'));
      if (eff(x, 'appr') !== 'Ja') { all = false; if (d && (!dueOpen || d < dueOpen)) dueOpen = d; }
      if (d && (!dueAny || d < dueAny)) dueAny = d;
    }
    return { n: rs.length, allAppr: all, due: dueOpen || dueAny };
  }
  const betaFatDate = (r) => { const l = fatLink(r); return (l && l.due) || toDate(val(r, 'fatDate')); };
  const betaProtOk = (r) => { const l = fatLink(r); return l ? l.allAppr : !!toDate(val(r, 'protoDate')); };

  // invoercellen
  const dAttr = (r, f) => ' data-k="' + esc(fkey(r)) + '" data-f="' + f + '" data-def="' + esc(dflt(r, f)) + '"';
  const dCls = (r, f) => isDirty(r, f) ? ' lio-in-dirty' : '';
  const fInput = (r, f, w, ph) => '<input type="text" class="lio-in lio-in-txt' + dCls(r, f) + '"' + dAttr(r, f) + ' value="' + esc(eff(r, f)) + '"' +
    (ph ? ' placeholder="' + esc(ph) + '"' : '') + (w ? ' style="min-width:' + w + 'px"' : '') + '>';
  const fDate = (r, f) => '<input type="date" class="lio-in' + dCls(r, f) + '"' + dAttr(r, f) + ' value="' + esc(eff(r, f)) + '">';
  const fChk = (r, f) => '<input type="checkbox" class="lio-in lio-cb' + dCls(r, f) + '"' + dAttr(r, f) + (eff(r, f) === 'Ja' ? ' checked' : '') + '>';
  function fSel(r, f, opts) {
    const cur = eff(r, f);
    return '<select class="lio-in' + dCls(r, f) + '"' + dAttr(r, f) + '>' +
      opts.map(o => '<option value="' + o + '"' + (o === cur ? ' selected' : '') + '>' + (o === '' ? '—' : o) + '</option>').join('') + '</select>';
  }
  // na een wijziging alleen de afgeleide cellen (status, administratie) van die regel bijwerken
  function refreshRow(tr) {
    const r = tr && FATBY[tr.dataset.k]; if (!r) return;
    tr.querySelectorAll('td[data-live]').forEach(td => {
      const d = DEFS.fat.find(x => x.k === td.dataset.live); if (!d) return;
      td.className = [d.num ? 'lio-num' : '', d.cls ? d.cls(r) : ''].filter(Boolean).join(' ');
      td.innerHTML = d.html ? d.html(r) : esc(d.f(r));
    });
    $('lio-summary').innerHTML = summaryFor('fat', FATVIEW);
  }

  // ── terugschrijven naar IFS (klembord !IFS.COPYOBJECT) ────────────────────
  // Terugschrijven naar IFS staat uit: IFS maakt bij plakken een nieuwe regel (nieuw REQ_ID) aan. Zet op true om het paneel weer te tonen.
  const IFS_WRITEBACK = false;
  function syncIfsPanel() {
    let p = $('lio-ifs');
    if (!IFS_WRITEBACK || TAB !== 'fat') { if (p) p.style.display = 'none'; return; }
    if (!p) {
      p = document.createElement('details'); p.id = 'lio-ifs'; p.className = 'lio-ifs';
      p.innerHTML = '<summary>↩ Terugschrijven naar IFS (klembord)</summary><div class="lio-ifs-body">' +
        '<p class="lio-dim">Kopieer in IFS de FAT-regels (<code>!IFS.COPYOBJECT</code>) en plak ze links. Rechts staat dezelfde tekst met jullie invulling erin verwerkt — ' +
        'alleen de gewijzigde regels. Kopieer die en plak ze terug in IFS. Naam IHC afwijkend van Responsible Name komt in Additional Information (<code>IHC aanwezig: naam</code>).</p>' +
        '<div class="lio-ifs-cols"><textarea id="lio-ifs-in" spellcheck="false" placeholder="Plak hier de IFS.COPYOBJECT-tekst…"></textarea>' +
        '<textarea id="lio-ifs-out" spellcheck="false" readonly placeholder="Bijgewerkte tekst verschijnt hier"></textarea></div>' +
        '<label class="lio-dim" style="margin-right:.6rem"><input type="checkbox" id="lio-ifs-min"> alleen gewijzigde velden (test: werkt de bestaande regel dan bij?)</label>' +
        '<button class="btn-val" id="lio-ifs-run" type="button">↻ Verwerk</button> <button class="btn-val" id="lio-ifs-copy" type="button">📋 Kopieer resultaat</button> ' +
        '<span class="lio-fat-status" id="lio-ifs-msg"></span></div>';
      const wrap = document.querySelector('.lio-table-wrap'); wrap.parentNode.insertBefore(p, wrap);
      $('lio-ifs-in').addEventListener('input', runIfs);
      $('lio-ifs-run').addEventListener('click', runIfs);
      $('lio-ifs-min').addEventListener('change', runIfs);
      $('lio-ifs-copy').addEventListener('click', () => {
        const o = $('lio-ifs-out'); if (!o.value) return;
        const done = () => { $('lio-ifs-msg').textContent = 'Gekopieerd ✓ — plak het terug in IFS'; };
        if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(o.value).then(done, () => { o.select(); document.execCommand('copy'); done(); });
        else { o.select(); document.execCommand('copy'); done(); }
      });
    }
    p.style.display = '';
  }
  function runIfs() {
    const txt = $('lio-ifs-in').value, out = $('lio-ifs-out'), msg = $('lio-ifs-msg');
    if (!txt.trim()) { out.value = ''; msg.textContent = ''; return; }
    if (!/\$RECORD=/.test(txt)) { out.value = ''; msg.textContent = 'Geen $RECORD gevonden — plak de volledige IFS.COPYOBJECT-tekst.'; return; }
    const r = FatData.applyToIfs(txt, effectiveInfo(), { minimal: !!($('lio-ifs-min') && $('lio-ifs-min').checked) });
    out.value = r.text;
    const nf = r.touched.reduce((s, t) => s + t.fields.length, 0);
    msg.textContent = r.changed
      ? r.changed + ' van ' + r.total + ' regel(s) bijgewerkt (' + nf + ' veld(en)); ' + (r.total - r.changed) + ' ongewijzigd en weggelaten'
      : 'Geen wijzigingen: van de ' + r.total + ' geplakte regel(s) is niets aangepast ten opzichte van IFS.';
  }


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
    const lk = fatLink(r);
    const fd = betaFatDate(r), hasFat = !!(fd || lk || trim(val(r, 'fatLoc')));
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
      if (!betaProtOk(r)) i += 0.2;                                       // protocol nog niet goedgekeurd
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
      { h: '',             k: 'exp',  f: r => '', html: r => docsFor(r).length ? '<button type="button" class="lio-exp" title="Toon gevraagde documentatie voor dit systeem">▸</button>' : '<span class="lio-dim" title="Geen requirements gevonden in addreq-list.json">·</span>',
        title: 'Klik op ▸ om te zien welke documentatie (Type Description) gevraagd wordt' },
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
      { h: 'FAT date',     k: 'fdate', f: r => fmtDate(betaFatDate(r)), d: r => betaFatDate(r),
        cls: r => { const d = betaFatDate(r); if (!d) return '';
          const x = Math.round((d - startOfToday()) / 86400000); return x < 0 ? 'lio-bad' : x <= 30 ? 'lio-warn-c' : ''; } },
      { h: 'Protocol',     k: 'proto', f: r => betaProtOk(r) ? 'Y' : '', cls: r => betaProtOk(r) ? 'lio-ok' : '' },
      { h: 'Planned',      k: 'pdd',  f: r => fmtDate(val(r, 'planned')), d: r => toDate(val(r, 'planned')) },
      ...COMMON_TAIL,
    ],
    fat: [
      { h: 'IFS',          k: 'ifs',  f: r => '', html: r => '<a class="lio-ifs-x" href="#" title="Open alle additional requirements van deze PO-regel in IFS">🔗</a>',
        title: 'Opent in IFS alle additional requirements van deze PO-regel' },
      { h: 'Project',      k: 'sp',   f: r => trim(fg(r, 'proj')) },
      { h: 'PO',           k: 'po',   f: r => trim(fg(r, 'order')) },
      { h: 'Regel/Rel.',   k: 'ln',   f: r => trim(fg(r, 'line')) + '/' + trim(fg(r, 'rel')), n: r => (num(fg(r, 'line')) || 0) * 1000 + (num(fg(r, 'rel')) || 0),
        title: 'Line No / Release No' },
      { h: 'System No.',   k: 'sys',  f: r => trim(fg(r, 'sys')) },
      { h: 'Naam',         k: 'name', f: r => flat(fg(r, 'desc')) },
      { h: 'Leverancier',  k: 'sup',  f: r => flat(fg(r, 'vendor')) },
      { h: 'Expected FAT', k: 'due',  f: r => fmtDate(fg(r, 'due')), d: r => toDate(fg(r, 'due')),
        cls: r => { const d = toDate(fg(r, 'due')); if (!d || eff(r, 'appr') === 'Ja') return '';
          const x = Math.round((d - startOfToday()) / 86400000); return x < 0 ? 'lio-bad' : x <= 14 ? 'lio-warn-c' : ''; } },
      { h: 'Status',       k: 'st',   live: true, f: r => ST[stOf(r)][0], n: r => stOf(r), cls: r => ST[stOf(r)][1] },
      { h: 'Protocol ontvangen', k: 'recv', edit: true, f: r => eff(r, 'recv'), d: r => toDate(eff(r, 'recv')), html: r => fDate(r, 'recv'),
        title: 'Datum waarop het protocol binnenkwam (IFS: Date Received)' },
      { h: 'Attendance (IHC/Class/Klant)', k: 'att', f: r => attStr(r), title: 'IHC / Class / Customer Attendance ID uit IFS (H=Hold, W=Witness, R=Review, M=Monitor, A=Approve, E=Execute)' },
      { h: 'IHC erbij?',   k: 'need', f: r => needIhc(r) ? 'Ja, IHC erbij' : '', cls: r => needIhc(r) ? 'lio-warn-c' : '',
        title: 'Er staat een letter in IHC/Class/Klant Attendance → iemand van IHC moet aanwezig zijn' },
      { h: 'Administratie', k: 'adm', live: true, f: r => admOf(r) ? 'Administratie bijwerken' : '', cls: r => admOf(r) ? 'lio-warn-c' : '',
        title: 'Invitation ID aangemaakt, maar Protocol ontvangen of goedgekeurd ontbreekt in de administratie' },
      { h: 'Protocol goedgekeurd', k: 'appr', edit: true, f: r => eff(r, 'appr'), cls: r => eff(r, 'appr') === 'Ja' ? 'lio-ok' : '',
        html: r => fSel(r, 'appr', ['Nee', 'Ja']), title: 'Door Engineers goedgekeurd om uit te voeren (IFS: Approved)' },
      { h: 'Klant aanwezig', k: 'cli', edit: true, f: r => eff(r, 'cliReq'), cls: r => eff(r, 'cliReq') ? 'lio-ok' : '',
        html: r => fSel(r, 'cliReq', ['', 'X']), title: 'X = klant aanwezig (voorgevuld bij Customer Attendance W of R)' },
      { h: 'IHC aanwezig (naam)', k: 'ihc', edit: true, f: r => eff(r, 'ihc'),
        html: r => fInput(r, 'ihc', 130, needIhc(r) ? 'naam invullen' : ''),
        title: 'Standaard Responsible Name; een afwijkende naam komt in Additional Information in IFS' },
      { h: 'Locatie / adres', k: 'loc', edit: true, f: r => eff(r, 'loc'), html: r => fInput(r, 'loc', 160), title: 'IFS: Test Location (leeg → Meeting Point)' },
      { h: 'Class n.v.t.', k: 'cnr', edit: true, f: r => eff(r, 'clsNR') === 'Ja' ? 'Ja' : '', html: r => fChk(r, 'clsNR'), title: 'Class Not Required' },
      { h: 'Klant n.v.t.', k: 'unr', edit: true, f: r => eff(r, 'cusNR') === 'Ja' ? 'Ja' : '', html: r => fChk(r, 'cusNR'), title: 'Customer Not Required' },
      { h: 'Invitation ID', k: 'inv', edit: true, f: r => eff(r, 'inv'), html: r => fInput(r, 'inv', 70) },
      { h: 'Leverdatum',   k: 'want', f: r => fmtDate(fg(r, 'want')), d: r => toDate(fg(r, 'want')), title: 'Wanted Delivery Date' },
      { h: 'Opmerking',    k: 'note', edit: true, f: r => eff(r, 'note'), html: r => fInput(r, 'note', 180), title: 'IFS: Note' },
    ],

  };

  // ── sorteren (zelfde gedrag als Late Items in de PO-Matcher) ─────────────
  const SORT = { beta: { col: 'bscore', dir: 'desc' }, large: { col: 'totW', dir: 'desc' }, late: { col: 'off', dir: 'desc' }, fat: { col: 'adm', dir: 'asc' } };
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
    if (TAB === 'fat') FATVIEW = rows;
    VIEWROWS = rows;
    syncIfsPanel();

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
    if (TAB === 'fat' && !FATR.length) {
      $('lio-summary').innerHTML = '<span class="lio-bad-txt">Geen FAT-lijst geladen' + (FATERR ? ' (' + esc(FATERR) + ')' : '') +
        '. Upload op de Admin-pagina de bedrijfsbrede Excel met een tabblad <b>FAT</b> en zet het gedownloade <code>fat-list.json</code> in <code>shared/</code> in de repo.</span>';
      $('lio-tbody').innerHTML = ''; return;
    }
    $('lio-summary').innerHTML = summaryFor(TAB, rows);
    $('lio-tbody').innerHTML = rows.length
      ? rows.map((r, ri) => '<tr data-ri="' + ri + '"' + (TAB === 'fat' ? ' data-k="' + esc(fkey(r)) + '"' : '') + '>' + defs.map(d => {
          const cls = [d.num ? 'lio-num' : '', d.cls ? d.cls(r) : ''].filter(Boolean).join(' ');
          const v = d.f(r);
          const lv = d.live ? ' data-live="' + d.k + '"' : '';
          if (d.html) return '<td' + lv + (cls ? ' class="' + cls + '"' : '') + '>' + d.html(r) + '</td>';
          return '<td' + lv + (cls ? ' class="' + cls + '"' : '') + (v ? ' title="' + esc(v) + '"' : '') +
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
      ? '<select class="lio-in" id="lio-fat-filter" title="Filter">' +
          FILTERS.map(f => '<option value="' + f[0] + '"' + (f[0] === FATFILTER ? ' selected' : '') + '>' + f[1] + '</option>').join('') + '</select> ' +
        '<button class="btn-val" id="lio-fat-reset" type="button" title="Wist je eigen invulling in deze browser">↺ Eigen wijzigingen wissen</button> ' +
        '<span class="lio-fat-status" id="lio-fat-status"></span>' : '';
    if (TAB === 'fat') {
      $('lio-fat-reset').addEventListener('click', resetDraft);
      updateFatStatus();
      const fsel = $('lio-fat-filter'); if (fsel) fsel.addEventListener('change', () => { FATFILTER = fsel.value; render(); });
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
        (!FATALL.length ? ' <br><span class="lio-dim">Documentatie per systeem (▸) vraagt <code>addreq-list.json</code> in <code>shared/</code> — upload het blad "Add Req" via Admin.</span>' : '') +
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
    const c = (fn) => rows.filter(fn).length;
    return '<b>' + rows.length + '</b> FAT-regels · <span class="lio-bad-txt"><b>' + c(r => stOf(r) === 0) + '</b> achterstallig</span> · <b>' +
      c(r => stOf(r) === 1) + '</b> binnen 14 dgn · <b>' + c(r => !!eff(r, 'recv')) + '</b> protocol ontvangen · <b>' +
      c(r => stOf(r) === 3) + '</b> goedgekeurd · <b>' + c(needIhc) + '</b> met IHC erbij · <b>' + c(admOf) + '</b> administratie bijwerken' +
      '<br><span class="lio-dim">Grijs/ingevuld uit IFS (fat-list.json); gele rand = eigen wijziging. Klik op een kolomkop om te sorteren.</span>';

  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();
})();