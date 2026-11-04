/* ============================================================================
 *  IHC Expedite 2.0 — Bezoekverslag (Field Expediting Checklist V2.1)
 *  ---------------------------------------------------------------------------
 *  Invulpagina voor een leveranciersbezoek. Alle 27 auditpunten van de Field
 *  Expediting Checklist V2.1 worden aangetikt (C / O / NC / N/A); bij een
 *  Observation of Non-conformity komt er automatisch een bevinding (F-01, F-02…)
 *  met eigenaar, deadline en status. Foto's worden verkleind en bewaard.
 *  Uit deze gegevens maakt verslag-docx.js het Word-verslag.
 *
 *  Opslag: alleen in de browser van dit apparaat (IndexedDB) — er is geen server.
 *  Wissel van apparaat? Gebruik "Concept opslaan als bestand" / "laden".
 *
 *  Scoring conform de checklist: C = 1,0 · O = 0,5 · NC = 0,0 · N/A telt niet mee.
 *  Classificatie: ≥95 A · ≥85 B · ≥70 C · anders D.
 * ========================================================================== */
(function () {
  'use strict';

  /* ── de 27 auditpunten (letterlijk uit de checklist) ───────────────────── */
  const POINTS = [
    [1, 'Planning & Delivery', 'Confirm visit scope, objectives, requirements and participants.', 'Method'],
    [2, 'Planning & Delivery', 'Verify progress and delivery feasibility against plan, critical parts and capacity.', 'Method'],
    [3, 'Planning & Delivery', 'Align WRD, Planned and Confirmed Delivery Dates with the schedule and IFS.', 'Method'],
    [4, 'Scope & Configuration', 'Confirm the latest Item List covers the complete supplier scope.', 'Material'],
    [5, 'Scope & Configuration', 'Reconcile PO and engineering scope with the Item List; flag gaps.', 'Material'],
    [6, 'Logistics Readiness', 'Verify mandatory logistics data is complete.', 'Material'],
    [7, 'Logistics Readiness', 'Confirm transport, Incoterm, address and ADN requirements.', 'Method'],
    [8, 'FAT & Quality', 'Confirm FAT scope, protocol, participants, timing, resources and records.', 'Method'],
    [9, 'Documentation', 'Define the document package and transfer route to Document Control.', 'Method'],
    [10, 'Inspection Governance', 'Assign inspection class, matrix, scope, sampling, criteria and evidence.', 'Method'],
    [11, 'Issue Management', 'Register pre-visit issues with owner, due date, status and escalation.', 'Method'],
    [12, 'Visit Readiness', 'Agree agenda, records, photo plan and inspection/loading milestones.', 'Method'],
    [13, 'Progress Verification', 'Verify physical progress and material status against report and baseline.', 'Method'],
    [14, 'Progress Verification', 'Record delays, deviations and delivery or quality risks with evidence.', 'Method'],
    [15, 'FAT & Quality', 'Verify QC and FAT completion, signed results and approvals.', 'Method'],
    [16, 'Scope & Configuration', 'Match physical scope and quantities to approved records.', 'Material'],
    [17, 'Traceability', 'Verify goods and package traceability through required identifiers.', 'Material'],
    [18, 'Packaging & Handling', 'Verify packaging condition, protection, securing, handling and storage.', 'Material'],
    [19, 'Labeling & Identification', 'Verify label durability, position, readability and package identification.', 'Material'],
    [20, 'Inspection Execution', 'Execute assigned inspection scope and sampling; record reclassification requests.', 'Method'],
    [21, 'Evidence', 'Capture retrievable photos of scope, packaging, labels, defects and corrections.', 'Method'],
    [22, 'Documentation', 'Transfer complete visit documents and link to the correct record.', 'Method'],
    [23, 'Issue Management', 'Register each finding once and control ownership, closure and approval.', 'Method'],
    [24, 'Planning Governance', 'Verify IFS change history and register planning mismatches.', 'Method'],
    [25, 'Release Readiness', 'Confirm final documents, labels, logistics terms and dates are consistent.', 'Method'],
    [26, 'Release Readiness', 'Confirm blocking issues are closed and technical, FAT and QA approvals are complete.', 'Method'],
    [27, 'Release Decision', 'Record decision, owner, conditions, open issues and dossier reference.', 'Method']
  ].map(p => ({ no: p[0], area: p[1], text: p[2], cat: p[3] }));

  /* ── configurator: de punten gegroepeerd in stappen ─────────────────────── */
  const GROUPS = [
    { id: 'prep', title: 'Voorbereiding & planning', icon: '🗓️', pts: [1, 2, 3, 12, 24] },
    { id: 'scope', title: 'Scope & configuratie', icon: '🧩', pts: [4, 5, 16] },
    { id: 'logi', title: 'Logistieke gereedheid', icon: '📦', pts: [6, 7, 17, 18, 19] },
    { id: 'fat', title: 'FAT, kwaliteit & inspectie', icon: '🔍', pts: [8, 15, 10, 20] },
    { id: 'prog', title: 'Voortgang', icon: '📈', pts: [13, 14] },
    { id: 'doc', title: 'Documentatie, bewijs & issues', icon: '📎', pts: [9, 22, 21, 11, 23] },
    { id: 'rel', title: 'Vrijgave', icon: '✅', pts: [25, 26, 27] }
  ];
  // Bevestigingen per punt: de onderdelen die de tekst van het auditpunt zelf noemt.
  const CHECKS = {
    1: ['Visit scope', 'Objectives', 'Requirements', 'Participants'],
    2: ['Progress vs plan', 'Delivery feasibility', 'Critical parts', 'Capacity'],
    3: ['WRD', 'Planned Delivery Date', 'Confirmed Delivery Date', 'Aligned with schedule', 'Aligned with IFS'],
    4: ['Latest Item List available', 'Covers the complete supplier scope'],
    5: ['PO scope reconciled with Item List', 'Engineering scope reconciled with Item List', 'Gaps flagged'],
    6: ['Mandatory logistics data is complete'],
    7: ['Packaging', 'Labeling', 'Handling', 'Transport', 'Incoterm', 'Address', 'ADN requirements'],
    8: ['FAT scope', 'Protocol', 'Participants', 'Timing', 'Resources', 'Records'],
    9: ['Document package defined', 'Transfer route to Document Control agreed'],
    10: ['Inspection class', 'Matrix', 'Scope', 'Sampling', 'Criteria', 'Evidence'],
    11: ['Pre-visit issues registered', 'Owner', 'Due date', 'Status', 'Escalation'],
    12: ['Agenda agreed', 'Records', 'Photo plan', 'Inspection/loading milestones'],
    13: ['Physical progress verified vs report and baseline', 'Material status verified vs report and baseline'],
    14: ['Delays recorded', 'Deviations recorded', 'Delivery or quality risks recorded', 'Evidence attached'],
    15: ['QC completed', 'FAT completed', 'Signed results', 'Approvals'],
    16: ['Physical scope matches approved records', 'Quantities match approved records'],
    17: ['Goods traceable', 'Packages traceable', 'Required identifiers present'],
    18: ['Packaging condition', 'Protection', 'Securing', 'Handling', 'Storage'],
    19: ['Label durability', 'Label position', 'Label readability', 'Package identification'],
    20: ['Assigned inspection scope executed', 'Sampling executed', 'Reclassification requests recorded'],
    21: ['Photos of scope', 'Photos of packaging', 'Photos of labels', 'Photos of defects', 'Photos of corrections', 'Photos retrievable'],
    22: ['Visit documents complete', 'Linked to the correct record'],
    23: ['Each finding registered once', 'Ownership controlled', 'Closure controlled', 'Approval controlled'],
    24: ['IFS change history verified', 'Planning mismatches registered'],
    25: ['Final documents consistent', 'Labels consistent', 'Logistics terms consistent', 'Dates consistent'],
    26: ['Blocking issues closed', 'Technical approval complete', 'FAT approval complete', 'QA approval complete'],
    27: ['Decision recorded', 'Owner', 'Conditions', 'Open issues', 'Dossier reference']
  };
  // Welke punten horen bij welk type bezoek (meerdere types = samen). Geen type = alle 27.
  // Een logistieke inspectie volgt pas ná een eventuele FAT aan het eind van het traject:
  // daar horen dus geen planning/voortgang/FAT-voorbereiding bij, wel de FAT-afronding (#15, #26).
  const SCOPE = {
    'Milestone': [1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14, 16, 21, 22, 23, 24],
    'Logistics inspection': [1, 6, 7, 12, 15, 16, 17, 18, 19, 20, 21, 22, 23, 25, 26, 27],
    'FAT': [1, 8, 10, 12, 14, 15, 20, 21, 22, 23],
    'Audit': Array.from({ length: 27 }, (x, i) => i + 1),
    'Follow-up': [1, 11, 14, 21, 22, 23, 26]
  };
  const TYPE_NL = { 'Milestone': 'milestone-bezoek', 'Logistics inspection': 'logistieke inspectie', 'FAT': 'FAT', 'Audit': 'audit', 'Follow-up': 'follow-up-bezoek' };
  // Vinkjes die al ergens anders worden geborgd: niet meer tonen en niet meetellen (geen dubbele borging).
  const REMOVED = {
    3: ['Aligned with IFS'],                                   // → #24 IFS change history
    7: ['Packaging', 'Labeling', 'Handling'],                  // → #18 verpakking, #19 labels
    10: ['Evidence'],                                          // → #21 bewijsfoto's
    14: ['Evidence attached'],                                 // → foto's bij de bevindingen
    17: ['Packages traceable'],                                // → #19 package identification
    21: ['Photos of packaging', 'Photos of labels', 'Photos of defects', 'Photos retrievable'],  // → foto's bij #18/#19/bevindingen; verslag
    25: ['Labels consistent', 'Logistics terms consistent'],   // → #19, #7
    26: ['Blocking issues closed', 'FAT approval complete']    // → bevindingenregister, #15
  };
  const NOTES = {
    3: 'Afstemming met IFS: zie #24.', 7: 'Packaging en handling: zie #18; labeling: zie #19.', 10: 'Bewijs: zie #21.',
    14: 'Bewijs: foto\'s bij de bevindingen.', 17: 'Pakketidentificatie: zie #19.', 21: 'Foto\'s van verpakking, labels en afwijkingen: bij #18, #19 en de bevindingen.',
    25: 'Labels: zie #19; logistieke afspraken: zie #7.', 26: 'Blokkerende issues: zie bevindingenregister; FAT-afronding: zie #15.'
  };
  // Punten die de tool zelf invult (niet aan te tikken): #1 uit Bezoek, #23 uit het bevindingenregister, #27 uit Vrijgave.
  const AUTO_PTS = [1, 23, 27];
  // Rondgang bij een logistieke inspectie (volgorde van de inspectie ter plaatse).
  const ROUTE_LOGI = [
    { id: 'l-doc', title: 'Voorbereiding bezoek', icon: '🗂️', pts: [12, 6, 7, 22, 25] },
    { id: 'l-scope', title: 'Scope & aantallen', icon: '🧩', pts: [16, 17] },
    { id: 'l-pack', title: 'Verpakking', icon: '📦', pts: [18] },
    { id: 'l-label', title: 'Labels', icon: '🏷️', pts: [19] },
    { id: 'l-qual', title: 'Kwaliteit & FAT-afronding', icon: '🔍', pts: [15, 20] },
    { id: 'l-end', title: 'Afronding', icon: '✅', pts: [21, 26] }
  ];
  const VISIT_TYPES = ['Milestone', 'Logistics inspection', 'FAT', 'Audit', 'Follow-up'];
  const STATUSES = [['C', 'C', 'Compliant'], ['O', 'O', 'Observation'], ['NC', 'NC', 'Non-conf.'], ['NA', 'N/A', 'n.v.t.']];
  const PTS = { C: 1, O: 0.5, NC: 0 };
  const ROLES = ['Field Expediter', 'Supplier Representative', 'QA / Technical Approver'];
  const DECISIONS = [['GO', 'GO', 'Zending goedgekeurd'], ['CONDITIONAL', 'CONDITIONAL GO', 'Onder voorwaarden'], ['NOGO', 'NO GO', 'Zending geblokkeerd']];

  /* ── hulpfuncties ───────────────────────────────────────────────────────── */
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const today = () => { const d = new Date(); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0'); };
  const uid = () => Math.random().toString(36).slice(2, 9);
  const pct = (n) => n.toFixed(1).replace('.', ',') + '%';
  let toastT = null;
  function toast(msg) { const t = $('vm-toast'); t.textContent = msg; t.classList.add('show'); clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('show'), 2600); }

  function blank() {
    return {
      v: 1,
      hdr: { supplier: '', date: today(), project: '', po: '', location: '', expediter: '', rep: '', client: '', present: '', status: 'Concept', types: [], goal: '', summary: '' },
      pts: {}, extra: [], photos: [], ref: { pos: [], source: '' },
      rel: { decision: '', conditions: '', action: '', owner: '', ref: '' },
      sign: ROLES.map(() => ({ name: '', date: today(), sig: '' }))
    };
  }
  let D = blank();
  let TAB = 'bezoek';
  let GRP = 'prep';           // gekozen stap in de configurator
  const OPEN = {};            // uitgeklapte auditpunten (alleen UI-staat)
  let SHOWALL = false;        // ook punten buiten de scope van het bezoektype tonen
  let EXP = null, EXPMETA = null, POQ = '';   // bedrijfsbrede expeditinglijst (alleen-lezen)
  function scopeSet() { if (!D.hdr.types.length) return null; const s = new Set(); D.hdr.types.forEach((t) => (SCOPE[t] || []).forEach((n) => s.add(n))); return s; }
  function inScope(no) { const s = scopeSet(); return !s || s.has(+no) || !!(D.pts[no] || {}).st; }
  function dateNl(iso) { if (!iso) return ''; const d = new Date(String(iso).slice(0, 10) + 'T12:00:00'); return isNaN(d) ? String(iso) : d.toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' }); }

  /* ── afgeleide gegevens: score, bevindingen, advies ─────────────────────── */
  function pt(no) { return D.pts[no] || (D.pts[no] = { st: '', checks: [], man: '', remark: '', owner: '', due: '', fstatus: 'Open', approval: '', blocking: false }); }

  // Gevaarlijke stoffen (ADN) in de gekoppelde PO-regels — zelfde trefwoorden als de DG Overview (shared/dg-rules.js).
  function dgHits() {
    const R = window.DG_RULES || {}, out = [];
    (D.ref && D.ref.pos || []).forEach((po) => po.lines.forEach((l) => {
      const desc = String(l.desc || '').toLowerCase(); if (!desc) return;
      for (const cat in R) {
        // alleen categorieën die vervoersplichtig gevaarlijk zijn (IMDG-klasse/UN-nummer of gas onder druk),
        // niet bv. cyberresilience, IHM of werkplaatsmiddelen
        if (!/IMDG|UN \d{4}|H280/.test(R[cat].regulations || '')) continue;
        if ((R[cat].keywords || []).some((k) => desc.indexOf(String(k).toLowerCase()) >= 0)) {
          out.push({ order: po.order, line: l.line + '-' + l.rel, desc: l.desc, cat: cat.replace(/^[^A-Za-z]*\d+\.\s*/, '') }); break;
        }
      }
    }));
    return out;
  }
  // ADN-vinkje alleen tonen als er gevaarlijke stoffen in de PO-regels zitten (of als er nog geen PO is gekoppeld).
  const adnNeeded = () => !(D.ref && D.ref.pos && D.ref.pos.length) || dgHits().length > 0;
  // Indexen (in CHECKS) van de vinkjes die getoond en meegeteld worden
  function visIdx(no) {
    const rm = REMOVED[no] || [];
    return (CHECKS[no] || []).map((c, i) => i).filter((i) => {
      const c = CHECKS[no][i];
      if (rm.indexOf(c) >= 0) return false;
      if (+no === 7 && c === 'ADN requirements' && !adnNeeded()) return false;
      return true;
    });
  }
  // Status volgt uit de getoonde vinkjes: alles bevestigd = C, deels = O; "afwijking"/"n.v.t." overrulen.
  function recompute(no) {
    const d = pt(no), V = visIdx(no), t = V.filter((i) => (d.checks || []).indexOf(i) >= 0).length;
    d.st = d.man || (V.length && t >= V.length ? 'C' : t > 0 ? 'O' : '');
  }
  const unconfirmed = (no) => { const d = D.pts[no] || {}; return visIdx(no).filter((i) => (d.checks || []).indexOf(i) < 0).map((i) => CHECKS[no][i]); };
  const isAuto = (no) => AUTO_PTS.indexOf(+no) >= 0;

  // Automatisch bepaalde punten: [onderdeel, in orde?]
  function autoChecks(no, findings) {
    const H = D.hdr, R = D.rel;
    if (+no === 1) return [['Visit scope (type bezoek)', H.types.length > 0], ['Objectives', !!tv(H.goal) || H.types.length > 0],
      ['Requirements (PO gekoppeld)', !!(D.ref && D.ref.pos && D.ref.pos.length) || !!tv(H.po)], ['Participants', !!tv(H.present)]];
    if (+no === 23) return [['Each finding registered once', true], ['Ownership', findings.every((f) => !!tv(f.owner))],
      ['Closure', true], ['Approval', findings.every((f) => f.fstatus !== 'Closed' || !!tv(f.approval))]];
    if (+no === 27) {
      if (!R.decision) return [];
      return [['Decision recorded', true], ['Owner', !!tv(R.owner)], ['Conditions', R.decision === 'GO' || !!tv(R.conditions)],
        ['Open issues', true], ['Dossier reference', !!tv(R.ref)]];
    }
    return [];
  }
  function autoStatus(no, findings) {
    const c = autoChecks(no, findings); if (!c.length) return { st: '', checks: c };
    const ok = c.filter((x) => x[1]).length;
    return { st: ok === c.length ? 'C' : ok ? 'O' : '', checks: c };
  }

  function derive() {
    let sum = 0, n = 0, total = 0; const cnt = { C: 0, O: 0, NC: 0, NA: 0, open: 0 };
    // status van aangetikte punten opnieuw bepalen (de getoonde vinkjes kunnen wijzigen, bv. ADN na het koppelen van een PO)
    POINTS.forEach((p) => { const d = D.pts[p.no]; if (!isAuto(p.no) && d && Array.isArray(d.checks) && (d.checks.length || d.man)) recompute(p.no); });
    const findings = [];
    POINTS.forEach((p) => {
      if (isAuto(p.no)) return;
      const d = D.pts[p.no]; if (!d || (d.st !== 'O' && d.st !== 'NC')) return;
      findings.push({ kind: 'pt', key: p.no, ptNo: p.no, type: d.st === 'NC' ? 'NCR' : 'OBS', text: d.remark, owner: d.owner, due: d.due, fstatus: d.fstatus || 'Open', approval: d.approval, blocking: !!d.blocking && d.st === 'NC', ref: 'p' + p.no });
    });
    D.extra.forEach((x) => findings.push({ kind: 'x', key: x.id, ptNo: x.point ? +x.point : null, type: x.type, text: x.text, owner: x.owner, due: x.due, fstatus: x.fstatus || 'Open', approval: x.approval, blocking: !!x.blocking && x.type === 'NCR', ref: 'x' + x.id }));
    findings.forEach((f, i) => { f.fid = 'F-' + String(i + 1).padStart(2, '0'); });
    const auto = {};
    POINTS.forEach((p) => {
      if (!inScope(p.no)) return;
      let s;
      if (isAuto(p.no)) { auto[p.no] = autoStatus(p.no, findings); s = auto[p.no].st; }
      else { total++; s = (D.pts[p.no] || {}).st; if (!s) { cnt.open++; return; } cnt[s]++; }
      if (s in PTS) { sum += PTS[s]; n++; }
    });
    const score = n ? Math.round(sum / n * 1000) / 10 : 0;
    const cls = !n ? '–' : score >= 95 ? 'A - Release Ready' : score >= 85 ? 'B - Minor Findings' : score >= 70 ? 'C - Conditional Release' : 'D - No Release';
    const blocking = findings.some((f) => f.blocking && f.fstatus !== 'Closed');
    const suggestion = blocking || cls.startsWith('D') ? 'NOGO' : cls.startsWith('A') ? 'GO' : n ? 'CONDITIONAL' : '';
    return { score, cls, cnt, scored: n, total, findings, blocking, suggestion, auto };
  }

  /* ── opslag (IndexedDB) ─────────────────────────────────────────────────── */
  const DBN = 'ihc-bezoekverslag';
  function idb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open(DBN, 1);
      r.onupgradeneeded = () => r.result.createObjectStore('kv');
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    });
  }
  async function kvSet(k, v) { const db = await idb(); return new Promise((res, rej) => { const tx = db.transaction('kv', 'readwrite'); tx.objectStore('kv').put(v, k); tx.oncomplete = res; tx.onerror = () => rej(tx.error); }); }
  async function kvGet(k) { const db = await idb(); return new Promise((res, rej) => { const q = db.transaction('kv').objectStore('kv').get(k); q.onsuccess = () => res(q.result); q.onerror = () => rej(q.error); }); }
  async function kvDel(k) { const db = await idb(); return new Promise((res) => { const tx = db.transaction('kv', 'readwrite'); tx.objectStore('kv').delete(k); tx.oncomplete = res; }); }
  let saveT = null, saveWarned = false;
  function save() {
    clearTimeout(saveT);
    saveT = setTimeout(() => { kvSet('draft', D).catch(() => { if (!saveWarned) { saveWarned = true; toast('Automatisch bewaren lukt niet in deze browser — sla het concept zelf op als bestand.'); } }); }, 350);
    updateBar();
  }

  /* ── foto's ─────────────────────────────────────────────────────────────── */
  async function compress(file) {
    const bmp = await createImageBitmap(file);
    const max = 1600, k = Math.min(1, max / Math.max(bmp.width, bmp.height));
    const w = Math.round(bmp.width * k), h = Math.round(bmp.height * k);
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    c.getContext('2d').drawImage(bmp, 0, 0, w, h);
    if (bmp.close) bmp.close();
    return { data: c.toDataURL('image/jpeg', 0.78), w, h };
  }
  function pickPhotos(ref, camera) {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = 'image/*'; inp.multiple = !camera;
    if (camera) inp.setAttribute('capture', 'environment');
    inp.onchange = async () => {
      const files = Array.from(inp.files || []); if (!files.length) return;
      toast('Foto verwerken…');
      for (const f of files) {
        try { const c = await compress(f); D.photos.push({ id: uid(), ref, desc: '', data: c.data, w: c.w, h: c.h }); }
        catch (e) { toast('Deze foto kon niet gelezen worden.'); }
      }
      save(); render(); toast(files.length > 1 ? files.length + ' foto\'s toegevoegd' : 'Foto toegevoegd');
    };
    inp.click();
  }
  const photosOf = (ref) => D.photos.filter((p) => p.ref === ref);
  function photoButtons(ref) {
    return '<div class="vm-row"><button type="button" class="vm-btn vm-btn-sm" data-act="ph-cam" data-ref="' + ref + '">📷 Foto maken</button>' +
      '<button type="button" class="vm-btn vm-btn-sm" data-act="ph-gal" data-ref="' + ref + '">🖼️ Uit galerij</button></div>';
  }
  function photoStrip(ref) {
    return '<div class="vm-photos">' + photosOf(ref).map((p) =>
      '<div class="vm-ph"><img src="' + p.data + '" alt=""><button type="button" data-act="ph-del" data-id="' + p.id + '" aria-label="Verwijder foto">×</button>' +
      '<input type="text" data-k="photo:' + p.id + '" value="' + esc(p.desc) + '" placeholder="Omschrijving"></div>').join('') + '</div>' + photoButtons(ref);
  }

  /* ── koppeling met de bedrijfsbrede expeditinglijst (project → leverancier → PO) ── */
  const COLDEF = { orderNo: 'Order No', sub: 'Sub Project ID', subDesc: 'Sub Project Description', supplier: 'Supplier Name', buyer: 'Buyer Name', tc: 'Technical Coordinator Name', part: 'Part No', desc: 'Description', uref: 'Unified Reference Code', qty: 'Qty', status: 'PO Line Status', planned: 'Planned Delivery Date', lwr: 'Latest Wanted Receipt Date', fatReq: 'FAT Date Required' };
  const COL = (k) => (window.ExpeditingCore && ExpeditingCore.COLS && ExpeditingCore.COLS[k]) || COLDEF[k];
  const tv = (v) => String(v == null ? '' : v).trim();
  const capital = (x) => x.charAt(0).toUpperCase() + x.slice(1);
  const trimDot = (x) => tv(x).replace(/[.\s]+$/, '');
  async function loadExp() {
    try {
      if (!window.ExpeditingData) return;
      const raw = await ExpeditingData.loadRaw(); if (raw && Array.isArray(raw.rows)) EXP = raw.rows;
      try { EXPMETA = await ExpeditingData.meta(); } catch (e) { EXPMETA = null; }
    } catch (e) { EXP = null; }
    if (TAB === 'bezoek') renderPO();
  }
  function poMatches() {
    if (!EXP) return [];
    const pj = tv(D.hdr.project).toLowerCase(), sp = tv(D.hdr.supplier).toLowerCase(), q = tv(POQ).toLowerCase();
    if (!pj && !sp && !q) return [];
    const by = new Map();
    EXP.forEach((r) => {
      const sub = tv(r[COL('sub')]), sd = tv(r[COL('subDesc')]), su = tv(r[COL('supplier')]), o = tv(r[COL('orderNo')]);
      if (!o) return;
      if (pj && sub.toLowerCase().indexOf(pj) < 0 && sd.toLowerCase().indexOf(pj) < 0) return;
      if (sp && su.toLowerCase().indexOf(sp) < 0) return;
      if (q && o.toLowerCase().indexOf(q) < 0) return;
      let g = by.get(o); if (!g) { g = { order: o, supplier: su, sub, subDesc: sd, n: 0 }; by.set(o, g); } g.n++;
    });
    return Array.from(by.values()).sort((a, b) => (a.order < b.order ? -1 : 1));
  }
  function lineSnap(r) {
    return { line: tv(r['Line No']), rel: tv(r['Release No']), part: tv(r[COL('part')]), desc: tv(r[COL('desc')]), uref: tv(r[COL('uref')]), qty: tv(r[COL('qty')]),
      uom: tv(r['Purch UoM']), status: tv(r[COL('status')]), planned: tv(r[COL('planned')]), lwr: tv(r[COL('lwr')]), fat: tv(r[COL('fatReq')]) };
  }
  function pickPO(order) {
    const ex = D.ref.pos.findIndex((x) => x.order === order);
    if (ex >= 0) D.ref.pos.splice(ex, 1);
    else {
      const rows = (EXP || []).filter((r) => tv(r[COL('orderNo')]) === order); if (!rows.length) return;
      const r0 = rows[0];
      D.ref.pos.push({ order, supplier: tv(r0[COL('supplier')]), sub: tv(r0[COL('sub')]), subDesc: tv(r0[COL('subDesc')]), buyer: tv(r0[COL('buyer')]), tc: tv(r0[COL('tc')]),
        lines: rows.map(lineSnap).sort((a, b) => (+a.line - +b.line) || (+a.rel - +b.rel)) });
      const up = EXPMETA && (EXPMETA.uploaded || EXPMETA.date); D.ref.source = up ? String(up).slice(0, 10) : '';
    }
    syncPOHeader(); save(); render();
  }
  function syncPOHeader() {
    const P = D.ref.pos; if (!P.length) return;
    D.hdr.po = P.map((x) => x.order).join(', ');
    const sups = Array.from(new Set(P.map((x) => x.supplier))), subs = Array.from(new Set(P.map((x) => x.sub)));
    if (sups.length === 1) D.hdr.supplier = sups[0];
    if (subs.length === 1) D.hdr.project = subs[0];
  }
  function poBox() {
    if (!EXP) return '<div class="vm-empty">' + (window.ExpeditingData ? 'Expeditinglijst laden…' : 'De expeditinglijst is niet beschikbaar — vul de gegevens handmatig in.') + '</div>';
    if (!tv(D.hdr.project) && !tv(D.hdr.supplier) && !tv(POQ)) return '<div class="vm-empty">Kies een project en/of leverancier; de bijbehorende PO\'s verschijnen hier.</div>';
    const m = poMatches();
    if (!m.length) return '<div class="vm-empty">Geen PO\'s gevonden voor deze combinatie.</div>';
    const sel = new Set(D.ref.pos.map((x) => x.order));
    return '<div class="vm-polist">' + m.slice(0, 40).map((g) => '<button type="button" class="vm-po' + (sel.has(g.order) ? ' on' : '') + '" data-act="po-pick" data-o="' + esc(g.order) + '"><span class="box">✓</span><span class="vm-po-main"><b>PO ' + esc(g.order) + '</b><small>' + esc(g.supplier) + ' · ' + esc(g.sub) + (projDesc(g) ? ' – ' + esc(projDesc(g)) : '') + '</small></span><span class="vm-tile-prog">' + g.n + ' regel' + (g.n === 1 ? '' : 's') + '</span></button>').join('') + '</div>' +
      (m.length > 40 ? '<p class="vm-hint">Nog ' + (m.length - 40) + ' PO\'s — verfijn met leverancier of PO-nummer.</p>' : '');
  }
  function dlists() {
    if (!EXP) return '';
    const pj = tv(D.hdr.project).toLowerCase(), projs = new Map(), sups = new Set();
    EXP.forEach((r) => {
      const sub = tv(r[COL('sub')]); if (sub && !projs.has(sub)) projs.set(sub, tv(r[COL('subDesc')]));
      const su = tv(r[COL('supplier')]); if (su && (!pj || sub.toLowerCase().indexOf(pj) >= 0)) sups.add(su);
    });
    return '<datalist id="dl-proj">' + Array.from(projs).sort().map((x) => '<option value="' + esc(x[0]) + '">' + esc(x[1]) + '</option>').join('') + '</datalist>' +
      '<datalist id="dl-sup">' + Array.from(sups).sort().map((v) => '<option value="' + esc(v) + '"></option>').join('') + '</datalist>';
  }
  function poSelected() {
    if (!D.ref.pos.length) return '';
    return '<div class="vm-q">Gekoppelde PO\'s (referentie voor het verslag)</div>' + D.ref.pos.map((x) =>
      '<div class="vm-find"><button type="button" class="vm-btn vm-btn-sm" style="float:right;padding:.25rem .6rem" data-act="po-del" data-o="' + esc(x.order) + '" aria-label="Ontkoppel">✕</button>' +
      '<b class="fid">PO ' + esc(x.order) + '</b> · ' + esc(x.supplier) + '<br><span class="vm-hint">' + esc(x.sub) + (projDesc(x) ? ' – ' + esc(projDesc(x)) : '') + ' · ' + x.lines.length + ' regel(s)' + (x.buyer ? ' · buyer ' + esc(x.buyer) : '') + '</span>' +
      '<div class="vm-hint" style="margin-top:.35rem">' + x.lines.slice(0, 6).map((l) => esc(l.line + '-' + l.rel + '  ' + l.desc)).join('<br>') + (x.lines.length > 6 ? '<br>…' : '') + '</div></div>').join('');
  }
  function renderPO() { const r = $('vm-po-res'); if (r) r.innerHTML = poBox(); const d = $('vm-dl'); if (d) d.innerHTML = dlists(); }

  /* ── samenvatting als lopende tekst ─────────────────────────────────────── */
  // projectomschrijving zonder het herhaalde projectnummer ("YN1320 - Cable layer" → "Cable layer")
  const projDesc = (x) => tv(x.subDesc).replace(new RegExp('^' + tv(x.sub).replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*[-–:]\\s*', 'i'), '');
  const joinNl = (a) => (a.length <= 1 ? (a[0] || '') : a.slice(0, -1).join(', ') + ' en ' + a[a.length - 1]);
  const nWord = (n, one, more) => n + ' ' + (n === 1 ? one : more);
  function narrative(dv) {
    dv = dv || derive();
    const H = D.hdr, out = [], types = H.types.map((t) => TYPE_NL[t] || t);
    let p1 = 'Op ' + (dateNl(H.date) || '[datum]') + ' heeft ' + (H.expediter || 'IHC') + ' ' + (types.length ? 'een ' + joinNl(types) : 'een leveranciersbezoek') +
      ' uitgevoerd bij ' + (H.supplier || 'de leverancier') + (H.location ? ' in ' + H.location : '');
    const sd = D.ref.pos[0] ? projDesc(D.ref.pos[0]) : '';
    if (H.project) p1 += ' voor project ' + H.project + (sd ? ' (' + sd + ')' : '');
    if (D.ref.pos.length) p1 += ', met betrekking tot ' + (D.ref.pos.length === 1 ? 'PO ' : 'de PO\'s ') + joinNl(D.ref.pos.map((x) => x.order + ' (' + nWord(x.lines.length, 'regel', 'regels') + ')'));
    else if (H.po) p1 += ', met betrekking tot ' + H.po;
    p1 += '.'; if (H.present) p1 += ' Aanwezig waren ' + trimDot(H.present) + '.';
    out.push(p1);

    const parts = [];
    routeGroups().forEach((g) => {
      const pts = g.pts.filter((no) => !isAuto(no) && inScope(no) && (D.pts[no] || {}).st); if (!pts.length) return;
      const c = { C: 0, O: 0, NC: 0, NA: 0 }; pts.forEach((no) => { c[D.pts[no].st]++; });
      const areas = Array.from(new Set(pts.map((no) => POINTS[no - 1].area)));
      const res = [];
      if (c.C) res.push(c.C === pts.length ? 'alles was in orde' : c.C + ' in orde');
      if (c.O) res.push(nWord(c.O, 'observatie', 'observaties'));
      if (c.NC) res.push(nWord(c.NC, 'non-conformiteit', 'non-conformiteiten'));
      if (c.NA) res.push(c.NA + ' niet van toepassing');
      parts.push('bij ' + (/^[A-Z]{2}/.test(g.title) ? g.title : g.title.charAt(0).toLowerCase() + g.title.slice(1)) + ' ' + (pts.length === 1 ? 'is 1 punt' : 'zijn ' + pts.length + ' punten') + ' nagelopen (' + areas.join(', ') + '): ' + joinNl(res));
    });
    if (parts.length) {
      let p2 = 'Tijdens het bezoek is de Field Expediting Checklist doorlopen. ' + capital(parts.join('; ')) + '.';
      if (dv.cnt.open) p2 += ' ' + capital(nWord(dv.cnt.open, 'punt', 'punten')) + ' binnen de scope van dit bezoek ' + (dv.cnt.open === 1 ? 'is' : 'zijn') + ' niet beoordeeld.';
      out.push(p2);
    } else out.push('Er zijn nog geen punten van de checklist beoordeeld.');
    if (inScope(7) && D.ref && D.ref.pos && D.ref.pos.length) {
      const dg = dgHits();
      if (dg.length) out.push('De gekoppelde PO-regels bevatten gevaarlijke stoffen (' + joinNl(Array.from(new Set(dg.map((x) => x.cat.toLowerCase())))) + '); daarom zijn ook de ADN-eisen gecontroleerd.');
    }

    if (!dv.findings.length) out.push('Er zijn geen afwijkingen vastgesteld.');
    else {
      const s = ['Er ' + (dv.findings.length === 1 ? 'is één bevinding' : 'zijn ' + dv.findings.length + ' bevindingen') + ' vastgelegd.'];
      dv.findings.forEach((f) => {
        const p = f.ptNo ? POINTS[f.ptNo - 1] : null;
        let t = f.fid + ' is een ' + (f.blocking ? 'blokkerende ' : '') + (f.type === 'NCR' ? 'non-conformiteit' : 'observatie') + (p ? ' bij punt #' + p.no + ' (' + p.area + ')' : '');
        const miss = f.kind === 'pt' && f.type === 'OBS' ? unconfirmed(f.ptNo) : [];
        const txt = trimDot(f.text);
        if (txt) t += ': ' + txt;
        if (miss.length) t += (txt ? '; ' : ': ') + 'niet bevestigd ' + (miss.length === 1 ? 'is ' : 'zijn ') + joinNl(miss);
        t += '.';
        const fol = []; if (f.owner) fol.push('de opvolging ligt bij ' + f.owner); if (f.due) fol.push('uiterlijk ' + dateNl(f.due));
        t += ' ' + capital((fol.length ? fol.join(', ') + '; ' : '') + 'status: ' + (f.fstatus === 'Closed' ? 'gesloten' : 'open')) + '.';
        s.push(t);
      });
      out.push(s.join(' '));
    }

    const r = D.rel;
    let p4 = dv.scored ? 'De auditscore komt uit op ' + pct(dv.score) + ' (classificatie ' + dv.cls + ')' + (dv.blocking ? ', met ten minste één openstaande blokkerende non-conformiteit' : '') + '.' : '';
    const dec = r.decision === 'GO' ? 'Op basis hiervan is de zending vrijgegeven (GO).'
      : r.decision === 'CONDITIONAL' ? 'De zending is voorwaardelijk vrijgegeven (CONDITIONAL GO)' + (r.conditions ? ', onder de volgende voorwaarden: ' + trimDot(r.conditions) : '') + '.'
      : r.decision === 'NOGO' ? 'De zending is niet vrijgegeven (NO GO)' + (r.conditions ? ' vanwege: ' + trimDot(r.conditions) : '') + '.'
      : 'Er is nog geen vrijgavebesluit vastgelegd.';
    p4 += (p4 ? ' ' : '') + dec;
    if (r.action) p4 += ' Vereiste corrigerende actie: ' + trimDot(r.action) + '.';
    if (r.owner) p4 += ' Het besluit is vastgelegd door ' + trimDot(r.owner) + '.';
    if (r.ref) p4 += ' Dossierreferentie: ' + trimDot(r.ref) + '.';
    out.push(p4);
    if (tv(H.summary)) out.push(tv(H.summary));
    return out;
  }
  const narrHtml = () => narrative().map((x) => '<p>' + esc(x).replace(/\n/g, '<br>') + '</p>').join('');
  function updateNarr() { const n = $('vm-narr'); if (n) n.innerHTML = narrHtml(); }

  /* ── tabbladen ──────────────────────────────────────────────────────────── */
  function tabs() {
    const dv = derive();
    const list = [['bezoek', 'Bezoek'], ['checklist', 'Checklist', (dv.total - dv.cnt.open) + '/' + dv.total], ['bevind', 'Bevindingen', String(dv.findings.length)],
      ['fotos', "Foto's", String(D.photos.length)], ['vrijgave', 'Vrijgave'], ['teken', 'Ondertekening']];
    $('vm-tabs').innerHTML = list.map((t) => '<button type="button" class="vm-tab' + (t[0] === TAB ? ' on' : '') + '" data-tab="' + t[0] + '">' + t[1] + (t[2] ? '<small>' + t[2] + '</small>' : '') + '</button>').join('');
  }
  function updateBar() {
    const dv = derive();
    $('vm-bar-info').innerHTML = '<b>' + (dv.scored ? 'Score ' + pct(dv.score) + ' · ' + dv.cls.split(' - ')[0] : 'Nog niet beoordeeld') + '</b><br>' +
      (dv.total - dv.cnt.open) + '/' + dv.total + ' punten · ' + dv.findings.length + ' bevinding' + (dv.findings.length === 1 ? '' : 'en') + ' · ' + D.photos.length + " foto's";
    const t = document.querySelector('.vm-tab[data-tab=checklist] small'); if (t) t.textContent = (dv.total - dv.cnt.open) + '/' + dv.total;
    const b = document.querySelector('.vm-tab[data-tab=bevind] small'); if (b) b.textContent = String(dv.findings.length);
    const f = document.querySelector('.vm-tab[data-tab=fotos] small'); if (f) f.textContent = String(D.photos.length);
  }

  function fld(label, key, val, opts) {
    opts = opts || {};
    return '<div class="' + (opts.full ? 'full' : '') + '"><label class="f">' + label + (opts.area
      ? '<textarea data-k="' + key + '" placeholder="' + esc(opts.ph || '') + '">' + esc(val) + '</textarea>'
      : '<input type="' + (opts.type || 'text') + '" data-k="' + key + '"' + (opts.list ? ' list="' + opts.list + '" autocomplete="off"' : '') + ' value="' + esc(val) + '" placeholder="' + esc(opts.ph || '') + '">') + '</label></div>';
  }

  function viewBezoek() {
    const h = D.hdr;
    return '<section class="vm-card"><h2>Project, leverancier &amp; PO</h2><div class="vm-grid">' +
      fld('Project / Yard no.', 'hdr.project', h.project, { list: 'dl-proj', ph: 'Bijv. YN1320 of projectnaam' }) +
      fld('Supplier / IHC Unit', 'hdr.supplier', h.supplier, { list: 'dl-sup', ph: 'Begin met typen…' }) +
      '<div class="full"><label class="f">PO-nummer zoeken (optioneel)<input type="text" data-ui="poq" value="' + esc(POQ) + '" placeholder="Bijv. 3156018519" inputmode="numeric" autocomplete="off"></label></div>' +
      '</div><div id="vm-dl">' + dlists() + '</div><div id="vm-po-res" style="margin-top:.6rem">' + poBox() + '</div>' + poSelected() + '</section>' +
      '<section class="vm-card"><h2>Type bezoek</h2><p class="vm-hint" style="margin:0 0 .4rem">Het type bepaalt welke checklistpunten van toepassing zijn.</p><div class="vm-chips">' +
      VISIT_TYPES.map((t) => '<button type="button" class="vm-chip' + (h.types.indexOf(t) >= 0 ? ' on' : '') + '" data-act="type" data-t="' + t + '">' + t + ' <small>' + (SCOPE[t] || []).length + '</small></button>').join('') +
      '</div></section>' +
      '<section class="vm-card"><h2>Bezoekgegevens</h2><div class="vm-grid">' +
      fld('Audit date', 'hdr.date', h.date, { type: 'date' }) + fld('PO / Delivery ref.', 'hdr.po', h.po) +
      fld('Location', 'hdr.location', h.location) + fld('Field Expediter', 'hdr.expediter', h.expediter) +
      fld('Supplier representative', 'hdr.rep', h.rep) + fld('Client', 'hdr.client', h.client) +
      fld('Aanwezig (namen)', 'hdr.present', h.present, { full: true, ph: 'Bijv. Kris Moerlie, Jan Jansen (leverancier)' }) +
      fld('Doel van het bezoek (optioneel)', 'hdr.goal', h.goal, { full: true, area: true, ph: 'Komt in het verslag onder "Doel van het bezoek"' }) +
      '<div class="full"><label class="f">Status van het verslag<select data-k="hdr.status"><option' + (h.status === 'Concept' ? ' selected' : '') + '>Concept</option><option' + (h.status === 'Final' ? ' selected' : '') + '>Final</option></select></label></div>' +
      '</div></section>' +
      '<button type="button" class="vm-btn vm-btn-teal" data-tab="checklist" style="width:100%">Naar de checklist →</button>';
  }
  function adnNote() {
    if (!(D.ref && D.ref.pos && D.ref.pos.length)) return '<p class="vm-hint" style="margin:.7rem 0 0">Koppel op het tabblad Bezoek een PO; dan wordt bepaald of er gevaarlijke stoffen (ADN) in de PO-regels zitten.</p>';
    const H = dgHits();
    if (!H.length) return '<p class="vm-hint" style="margin:.7rem 0 0">✓ Geen gevaarlijke stoffen in de gekoppelde PO-regels — ADN is niet van toepassing.</p>';
    return '<p class="vm-warn">⚠️ Gevaarlijke stoffen in de PO-regels: ' + H.slice(0, 5).map((x) => esc(x.line + ' ' + x.desc) + ' (' + esc(x.cat) + ')').join('; ') + (H.length > 5 ? ' … (+' + (H.length - 5) + ')' : '') + '. Bevestig de ADN-eisen.</p>';
  }
  const STATE_ICO = { C: '✓', O: '!', NC: '✕', NA: '–', '': '' };
  const STATE_TXT = { C: 'Compliant', O: 'Observation', NC: 'Non-conformity', NA: 'N/A', '': 'Nog niet beoordeeld' };
  const routeGroups = () => (D.hdr.types.length === 1 && D.hdr.types[0] === 'Logistics inspection' && !SHOWALL ? ROUTE_LOGI : GROUPS);
  const visPts = (g) => g.pts.filter((no) => !isAuto(no) && (SHOWALL || inScope(no)));
  const visGroups = () => routeGroups().filter((g) => visPts(g).length);
  const grpProg = (g) => { let done = 0; visPts(g).forEach((no) => { if ((D.pts[no] || {}).st) done++; }); return done; };

  function tileHtml(p, dv) {
    const d = D.pts[p.no] || {}; const st = d.st || '', open = !!OPEN[p.no];
    const V = visIdx(p.no), ticked = V.filter((i) => (d.checks || []).indexOf(i) >= 0).length;
    let h = '<section class="vm-card vm-tile' + (st ? ' s-' + st : '') + (open ? ' open' : '') + (inScope(p.no) ? '' : ' out') + '" id="pt-' + p.no + '">' +
      '<button type="button" class="vm-tile-head" data-act="tile" data-no="' + p.no + '" aria-expanded="' + open + '">' +
      '<span class="vm-tile-ico">' + STATE_ICO[st] + '</span>' +
      '<span class="vm-tile-main"><span class="vm-pt-top"><span class="vm-pt-no">#' + p.no + '</span><span>' + p.area + '</span><span class="vm-pt-cat">' + p.cat + '</span>' + (inScope(p.no) ? '' : '<span class="vm-pt-cat">buiten scope</span>') + '</span>' +
      '<span class="vm-pt-text">' + esc(p.text) + '</span></span>' +
      '<span class="vm-tile-prog">' + (st === 'NA' ? 'n.v.t.' : st === 'NC' ? 'NC' : ticked + '/' + V.length) + '</span></button>';
    if (open) {
      h += '<div class="vm-tile-body">' + (+p.no === 7 ? adnNote() : '') + '<div class="vm-q">Bevestig ter plaatse</div><div class="vm-chks">' +
        V.map((i) => [CHECKS[p.no][i], i]).map(([c, i]) => '<button type="button" class="vm-chk' + ((d.checks || []).indexOf(i) >= 0 ? ' on' : '') + '" data-act="chk" data-no="' + p.no + '" data-i="' + i + '"><span class="box">✓</span><span>' + esc(c) + '</span></button>').join('') +
        '</div><div class="vm-quick">' +
        '<button type="button" class="vm-btn vm-btn-sm" data-act="all" data-no="' + p.no + '">✓ Alles in orde</button>' +
        '<button type="button" class="vm-btn vm-btn-sm' + (d.man === 'NC' ? ' on-nc' : '') + '" data-act="man" data-no="' + p.no + '" data-m="NC">✕ Afwijking (NC)</button>' +
        '<button type="button" class="vm-btn vm-btn-sm' + (d.man === 'NA' ? ' on-na' : '') + '" data-act="man" data-no="' + p.no + '" data-m="NA">n.v.t.</button></div>' +
        '<p class="vm-hint">Status: <b>' + STATE_TXT[st] + '</b>' + (st === 'O' ? ' — niet alle onderdelen bevestigd; licht toe wat er ontbreekt.' : '') + '</p>' +
        '<div class="vm-more"><label class="f">Opmerking<textarea data-k="pts.' + p.no + '.remark" placeholder="Wat is er gezien of afgesproken?">' + esc(d.remark || '') + '</textarea></label>' +
        '<div class="vm-q" style="margin-top:.7rem">Bewijs (foto)</div>' + photoStrip('p' + p.no);
      if ((st === 'O' || st === 'NC') && !photosOf('p' + p.no).length) h += '<p class="vm-warn">Nog geen bewijs toegevoegd bij deze afwijking.</p>';
      if (st === 'O' || st === 'NC') {
        const f = dv.findings.find((q) => q.kind === 'pt' && q.key === p.no) || {};
        h += '<div class="vm-find"><b class="fid">' + (f.fid || '') + '</b> · ' + (st === 'NC' ? 'NCR' : 'OBS') + ' — opvolging' +
          '<div class="vm-row">' + fld('Owner', 'pts.' + p.no + '.owner', d.owner || '') + fld('Due date', 'pts.' + p.no + '.due', d.due || '', { type: 'date' }) + '</div>' +
          '<div class="vm-row"><div><label class="f">Status<select data-k="pts.' + p.no + '.fstatus"><option' + (d.fstatus !== 'Closed' ? ' selected' : '') + '>Open</option><option' + (d.fstatus === 'Closed' ? ' selected' : '') + '>Closed</option></select></label></div>' +
          fld('Closure approval (naam)', 'pts.' + p.no + '.approval', d.approval || '') + '</div>' +
          (st === 'NC' ? '<label class="vm-check"><input type="checkbox" data-k="pts.' + p.no + '.blocking"' + (d.blocking ? ' checked' : '') + '> Blokkerend voor de vrijgave (blocking NCR)</label>' : '') + '</div>';
      }
      h += '</div></div>';
    }
    return h + '</section>';
  }

  function viewChecklist() {
    const dv = derive(), G = visGroups();
    if (!G.some((g) => g.id === GRP)) GRP = (G[0] || routeGroups()[0]).id;
    const gi = Math.max(0, G.findIndex((g) => g.id === GRP)), g = G[gi], sc = scopeSet();
    const scopeBar = '<div class="vm-scope">' + (sc
      ? '<span>Type bezoek: <b>' + D.hdr.types.join(' + ') + '</b> — ' + sc.size + ' van 27 punten van toepassing.</span>'
      : '<span>Geen type bezoek gekozen — alle 27 punten worden getoond. Kies het type op het tabblad Bezoek om alleen de relevante punten te zien.</span>') +
      (sc ? '<button type="button" class="vm-btn vm-btn-sm" data-act="showall">' + (SHOWALL ? 'Alleen relevante punten' : 'Toon alle 27 punten') + '</button>' : '') + '</div>';
    const step = '<div class="vm-steps">' + G.map((x, i) => {
      const done = grpProg(x), tot = visPts(x).length, full = done === tot;
      return '<button type="button" class="vm-step' + (x.id === GRP ? ' on' : '') + (full ? ' full' : '') + '" data-act="grp" data-g="' + x.id + '"><span class="vm-step-ico">' + (full ? '✓' : x.icon) + '</span><span class="vm-step-t">' + (i + 1) + '. ' + x.title + '</span><small>' + done + '/' + tot + '</small></button>';
    }).join('') + '</div>';
    let h = scopeBar + step + '<div class="vm-gh"><h2>' + g.icon + ' ' + g.title + '</h2><span>Stap ' + (gi + 1) + ' van ' + G.length + ' · ' + grpProg(g) + '/' + visPts(g).length + ' beoordeeld</span></div>' +
      '<p class="vm-hint" style="margin:0 0 .6rem">Tik een punt open, vink aan wat ter plaatse bevestigd is en voeg zo nodig bewijs toe.</p>';
    h += visPts(g).map((no) => tileHtml(POINTS[no - 1], dv)).join('');
    h += gi < G.length - 1
      ? '<button type="button" class="vm-btn vm-btn-teal" data-act="grp" data-g="' + G[gi + 1].id + '" style="width:100%">Volgende: ' + G[gi + 1].title + ' →</button>'
      : '<button type="button" class="vm-btn vm-btn-teal" data-tab="bevind" style="width:100%">Klaar — naar de bevindingen →</button>';
    return h;
  }

  function viewBevind() {
    const dv = derive();
    let h = '<section class="vm-card"><h2>Findings &amp; corrective action register</h2>';
    if (!dv.findings.length) h += '<div class="vm-empty">Nog geen bevindingen. Een Observation (O) of Non-conformity (NC) in de checklist maakt automatisch een bevinding aan.</div>';
    else h += '<table class="vm-tbl"><tr><th>ID</th><th>Punt</th><th>Type</th><th>Owner</th><th>Due</th><th>Status</th></tr>' +
      dv.findings.map((f) => '<tr' + (f.kind === 'pt' ? ' data-act="goto" data-no="' + f.ptNo + '" style="cursor:pointer"' : '') + '><td><b>' + f.fid + '</b></td><td>' + (f.ptNo ? '#' + f.ptNo : '–') + '</td><td>' + f.type + (f.blocking ? ' ⛔' : '') + '</td><td>' + esc(f.owner) + '</td><td>' + esc(f.due) + '</td><td>' + f.fstatus + '</td></tr>').join('') + '</table>' +
      '<p class="vm-hint">Tik op een regel om het bijbehorende checklistpunt te openen. ⛔ = blokkerend.</p>';
    h += '</section>';
    D.extra.forEach((x) => {
      const f = dv.findings.find((q) => q.kind === 'x' && q.key === x.id);
      h += '<section class="vm-card"><h2>' + (f ? f.fid : '') + ' · extra bevinding</h2><div class="vm-grid">' +
        fld('Bevinding / evidence', 'x.' + x.id + '.text', x.text, { full: true, area: true }) +
        '<div><label class="f">Type<select data-k="x.' + x.id + '.type"><option value="OBS"' + (x.type === 'OBS' ? ' selected' : '') + '>OBS</option><option value="NCR"' + (x.type === 'NCR' ? ' selected' : '') + '>NCR</option></select></label></div>' +
        '<div><label class="f">Checklistpunt<select data-k="x.' + x.id + '.point"><option value="">– geen –</option>' + POINTS.map((p) => '<option value="' + p.no + '"' + (String(x.point) === String(p.no) ? ' selected' : '') + '>#' + p.no + ' ' + esc(p.area) + '</option>').join('') + '</select></label></div>' +
        fld('Owner', 'x.' + x.id + '.owner', x.owner) + fld('Due date', 'x.' + x.id + '.due', x.due, { type: 'date' }) +
        '<div><label class="f">Status<select data-k="x.' + x.id + '.fstatus"><option' + (x.fstatus !== 'Closed' ? ' selected' : '') + '>Open</option><option' + (x.fstatus === 'Closed' ? ' selected' : '') + '>Closed</option></select></label></div>' +
        fld('Closure approval (naam)', 'x.' + x.id + '.approval', x.approval) + '</div>' +
        '<label class="vm-check"><input type="checkbox" data-k="x.' + x.id + '.blocking"' + (x.blocking ? ' checked' : '') + '> Blokkerend voor de vrijgave (alleen bij NCR)</label>' +
        photoStrip('x' + x.id) +
        '<div class="vm-row"><button type="button" class="vm-btn vm-btn-sm" data-act="x-del" data-id="' + x.id + '">🗑️ Verwijder deze bevinding</button></div></section>';
    });
    h += '<button type="button" class="vm-btn vm-btn-teal" data-act="x-add" style="width:100%">＋ Extra bevinding (niet aan een punt gekoppeld)</button>';
    return h;
  }

  function viewFotos() {
    const labelOf = (ref) => ref === 'g' ? 'Overig' : ref[0] === 'p' ? 'Punt #' + ref.slice(1) : (derive().findings.find((f) => f.ref === ref) || {}).fid || 'Bevinding';
    let h = '<section class="vm-card"><h2>Foto\'s (' + D.photos.length + ')</h2>';
    if (!D.photos.length) h += '<div class="vm-empty">Nog geen foto\'s. Voeg ze toe bij een checklistpunt of hieronder als overige foto.</div>';
    else h += '<div class="vm-photos">' + D.photos.map((p) => '<div class="vm-ph" style="width:150px"><img style="width:150px;height:150px" src="' + p.data + '" alt=""><button type="button" data-act="ph-del" data-id="' + p.id + '" aria-label="Verwijder foto">×</button><div class="vm-hint" style="margin:.2rem 0 0">' + esc(labelOf(p.ref)) + '</div><input type="text" data-k="photo:' + p.id + '" value="' + esc(p.desc) + '" placeholder="Omschrijving"></div>').join('') + '</div>';
    h += '</section><section class="vm-card"><h2>Overige foto\'s toevoegen</h2>' + photoButtons('g') + '</section>';
    return h;
  }

  function viewVrijgave() {
    const dv = derive(); const r = D.rel;
    return '<section class="vm-card"><h2>Audit score &amp; release gate</h2><div class="vm-score">' +
      '<div class="vm-kpi"><b>' + (dv.scored ? pct(dv.score) : '–') + '</b><span>Audit score</span></div>' +
      '<div class="vm-kpi"><b style="font-size:1.05rem;line-height:1.7">' + dv.cls + '</b><span>Classification</span></div>' +
      '<div class="vm-kpi"><b>' + dv.cnt.C + 'C · ' + dv.cnt.O + 'O · ' + dv.cnt.NC + 'NC</b><span>' + dv.scored + ' punten telt mee · ' + dv.cnt.NA + ' N/A · ' + dv.cnt.open + ' open</span></div>' +
      '<div class="vm-kpi"><b>' + (dv.blocking ? 'Yes' : 'No') + '</b><span>Blocking NCR</span></div></div></section>' +
      '<section class="vm-card"><h2>Shipment release decision</h2><div class="vm-dec">' +
      DECISIONS.map((d) => '<button type="button" data-act="dec" data-d="' + d[0] + '" class="' + (r.decision === d[0] ? 'on' : '') + '">' + d[1] + '<small style="display:block;font-weight:500;font-size:.75rem">' + d[2] + '</small></button>').join('') + '</div>' +
      (dv.suggestion ? '<p class="vm-hint">Advies op basis van de score' + (dv.blocking ? ' en de blokkerende NCR' : '') + ': <b>' + DECISIONS.find((d) => d[0] === dv.suggestion)[1] + '</b>. De beslissing is aan jou.</p>' : '') +
      '<div class="vm-grid" style="margin-top:.8rem">' +
      fld('Conditions / blocking reasons', 'rel.conditions', r.conditions, { full: true, area: true }) +
      fld('Required corrective action', 'rel.action', r.action, { full: true, area: true }) +
      fld('Release decision owner / date', 'rel.owner', r.owner) + fld('Final release reference / dossier', 'rel.ref', r.ref) + '</div></section>' +
      '<section class="vm-card"><h2>Samenvatting voor het verslag</h2><p class="vm-hint" style="margin:0 0 .5rem">Wordt automatisch opgebouwd uit het bezoek, de checklist, de bevindingen en het vrijgavebesluit.</p>' +
      '<div class="vm-narr" id="vm-narr">' + narrHtml() + '</div>' +
      fld('Eigen aanvulling (optioneel, komt achter de samenvatting)', 'hdr.summary', D.hdr.summary, { full: true, area: true }) + '</section>' +
      '<button type="button" class="vm-btn vm-btn-teal" data-tab="teken" style="width:100%">Naar ondertekening →</button>';
  }

  function viewTeken() {
    const dv = derive(), dec = DECISIONS.find((d) => d[0] === D.rel.decision);
    let h = '<section class="vm-card"><h2>Ondertekening</h2>' +
      '<p class="vm-hint" style="margin:0 0 .6rem">Door te ondertekenen bevestigen de bezoeker en de vertegenwoordiger van de leverancier dat zij overeenstemming hebben over de bevindingen, de actiepunten en het vrijgavebesluit in dit verslag.</p>' +
      '<div class="vm-score">' +
      '<div class="vm-kpi"><b>' + (dv.total - dv.cnt.open) + '/' + dv.total + '</b><span>Punten beoordeeld</span></div>' +
      '<div class="vm-kpi"><b>' + dv.findings.length + '</b><span>Bevindingen' + (dv.blocking ? ' · blokkerend' : '') + '</span></div>' +
      '<div class="vm-kpi"><b>' + (dv.scored ? pct(dv.score) : '–') + '</b><span>Audit score</span></div>' +
      '<div class="vm-kpi"><b style="font-size:1.05rem;line-height:1.7">' + (dec ? dec[1] : '–') + '</b><span>Vrijgave</span></div></div>' +
      (dv.cnt.open ? '<p class="vm-warn">Nog ' + nWord(dv.cnt.open, 'punt', 'punten') + ' niet beoordeeld. Rond eerst de checklist af voordat er getekend wordt.</p>' : '') +
      (!dec ? '<p class="vm-warn">Er is nog geen vrijgavebesluit vastgelegd (tabblad Vrijgave).</p>' : '') + '</section>';
    h += ROLES.map((role, i) => {
      const s = D.sign[i];
      return '<section class="vm-card"><h2 style="font-size:1.1rem">' + role + '</h2><div class="vm-row">' + fld('Naam', 'sign.' + i + '.name', s.name) + fld('Datum', 'sign.' + i + '.date', s.date, { type: 'date' }) + '</div>' +
        '<label class="f" style="margin-top:.6rem">Handtekening</label>' +
        '<button type="button" class="vm-sigbox' + (s.sig ? ' has' : '') + '" data-act="sig-open" data-i="' + i + '">' +
        (s.sig ? '<img src="' + s.sig + '" alt="Handtekening ' + esc(role) + '">' : '<span>✍️ Tik hier om te tekenen</span>') + '</button>' +
        (s.sig ? '<div class="vm-row"><button type="button" class="vm-btn vm-btn-sm" data-act="sig-open" data-i="' + i + '">Opnieuw tekenen</button><button type="button" class="vm-btn vm-btn-sm" data-act="sig-clear" data-i="' + i + '">Wissen</button></div>' : '') +
        '</section>';
    }).join('');
    return h;
  }

  const VIEWS = { bezoek: viewBezoek, checklist: viewChecklist, bevind: viewBevind, fotos: viewFotos, vrijgave: viewVrijgave, teken: viewTeken };
  function render() {
    tabs(); $('vm-main').innerHTML = VIEWS[TAB](); updateBar();
  }

  /* ── handtekening: volledig scherm, hele vlak om te tekenen ─────────────── */
  function openSigPad(i) {
    const role = ROLES[i], s = D.sign[i];
    const ov = document.createElement('div');
    ov.className = 'vm-sigpad';
    ov.innerHTML = '<div class="vm-sigpad-top"><div><b>' + esc(role) + (s.name ? ' — ' + esc(s.name) : '') + '</b><small>Teken met je vinger of pen in het witte vlak</small></div></div>' +
      '<div class="vm-sigpad-area"><canvas></canvas><div class="vm-sigpad-line"></div><span class="vm-sigpad-ph">Teken hier</span></div>' +
      '<div class="vm-sigpad-bar"><button type="button" class="vm-btn" data-sp="clear">Wissen</button><button type="button" class="vm-btn" data-sp="cancel">Annuleren</button><button type="button" class="vm-btn vm-btn-red" data-sp="ok">✓ Bevestigen</button></div>';
    document.body.appendChild(ov);
    document.body.style.overflow = 'hidden';
    try { if (ov.requestFullscreen && window.matchMedia('(pointer: coarse)').matches) ov.requestFullscreen().catch(() => {}); } catch (e) { /* niet ondersteund */ }

    const area = ov.querySelector('.vm-sigpad-area'), cv = ov.querySelector('canvas'), ph = ov.querySelector('.vm-sigpad-ph');
    const g = cv.getContext('2d');
    let strokes = [], cur = null, cw = 0, ch = 0;
    function fit() {
      const r = area.getBoundingClientRect(), dpr = window.devicePixelRatio || 1;
      const ow = cw;
      cw = r.width; ch = r.height;
      if (ow && ow !== cw) { const k = cw / ow; strokes.forEach((st) => st.forEach((p) => { p[0] *= k; p[1] *= k; })); }
      cv.width = Math.round(cw * dpr); cv.height = Math.round(ch * dpr);
      cv.style.width = cw + 'px'; cv.style.height = ch + 'px';
      g.setTransform(dpr, 0, 0, dpr, 0, 0);
      redraw();
    }
    function drawStroke(st) {
      g.lineCap = 'round'; g.lineJoin = 'round'; g.strokeStyle = '#0a1628'; g.lineWidth = Math.max(2.5, Math.min(cw, ch) / 120);
      g.beginPath(); g.moveTo(st[0][0], st[0][1]);
      if (st.length === 1) { g.lineTo(st[0][0] + 0.1, st[0][1] + 0.1); }
      for (let k = 1; k < st.length - 1; k++) { const mx = (st[k][0] + st[k + 1][0]) / 2, my = (st[k][1] + st[k + 1][1]) / 2; g.quadraticCurveTo(st[k][0], st[k][1], mx, my); }
      if (st.length > 1) { const L = st[st.length - 1]; g.lineTo(L[0], L[1]); }
      g.stroke();
    }
    function redraw() { g.clearRect(0, 0, cw, ch); strokes.forEach(drawStroke); ph.style.display = strokes.length ? 'none' : ''; }
    const pos = (e) => { const r = cv.getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top]; };
    cv.addEventListener('pointerdown', (e) => { e.preventDefault(); cv.setPointerCapture(e.pointerId); cur = [pos(e)]; strokes.push(cur); ph.style.display = 'none'; redraw(); });
    cv.addEventListener('pointermove', (e) => {
      if (!cur) return; e.preventDefault();
      const evs = e.getCoalescedEvents ? e.getCoalescedEvents() : [e];
      evs.forEach((ev) => cur.push(pos(ev)));
      redraw();
    });
    const end = () => { cur = null; };
    cv.addEventListener('pointerup', end); cv.addEventListener('pointercancel', end);
    const onResize = () => fit();
    window.addEventListener('resize', onResize);
    requestAnimationFrame(fit);

    function close() {
      window.removeEventListener('resize', onResize);
      try { if (document.fullscreenElement) document.exitFullscreen(); } catch (e) { /* */ }
      document.body.style.overflow = '';
      ov.remove();
    }
    // Bijsnijden tot de getekende lijnen (+ marge), zodat de handtekening netjes in het verslag past
    function exportSig() {
      let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
      strokes.forEach((st) => st.forEach((p) => { x0 = Math.min(x0, p[0]); y0 = Math.min(y0, p[1]); x1 = Math.max(x1, p[0]); y1 = Math.max(y1, p[1]); }));
      const m = 12; x0 = Math.max(0, x0 - m); y0 = Math.max(0, y0 - m); x1 = Math.min(cw, x1 + m); y1 = Math.min(ch, y1 + m);
      const w = Math.max(1, x1 - x0), h = Math.max(1, y1 - y0), sc = Math.min(2, 900 / w);
      const out = document.createElement('canvas'); out.width = Math.round(w * sc); out.height = Math.round(h * sc);
      const o = out.getContext('2d'); o.setTransform(sc, 0, 0, sc, -x0 * sc, -y0 * sc);
      const keep = g; strokes.forEach((st) => {
        o.lineCap = 'round'; o.lineJoin = 'round'; o.strokeStyle = '#0a1628'; o.lineWidth = Math.max(2.5, Math.min(cw, ch) / 120);
        o.beginPath(); o.moveTo(st[0][0], st[0][1]); if (st.length === 1) o.lineTo(st[0][0] + 0.1, st[0][1] + 0.1);
        for (let k = 1; k < st.length - 1; k++) { const mx = (st[k][0] + st[k + 1][0]) / 2, my = (st[k][1] + st[k + 1][1]) / 2; o.quadraticCurveTo(st[k][0], st[k][1], mx, my); }
        if (st.length > 1) { const L = st[st.length - 1]; o.lineTo(L[0], L[1]); }
        o.stroke();
      });
      void keep;
      return { data: out.toDataURL('image/png'), w: out.width, h: out.height };
    }
    ov.addEventListener('click', (e) => {
      const a = e.target.dataset && e.target.dataset.sp; if (!a) return;
      if (a === 'clear') { strokes = []; redraw(); }
      else if (a === 'cancel') close();
      else if (a === 'ok') {
        if (!strokes.length) { toast('Er is nog niets getekend.'); return; }
        const r = exportSig(); s.sig = r.data; s.sigW = r.w; s.sigH = r.h;
        save(); close(); render(); toast('Handtekening opgeslagen');
      }
    });
  }

  /* ── invoer en klikken ──────────────────────────────────────────────────── */
  function setByKey(k, val) {
    if (k.indexOf('photo:') === 0) { const p = D.photos.find((x) => x.id === k.slice(6)); if (p) p.desc = val; return; }
    const parts = k.split('.');
    if (parts[0] === 'x') { const x = D.extra.find((q) => q.id === parts[1]); if (x) x[parts[2]] = val; return; }
    if (parts[0] === 'pts') { pt(parts[1])[parts[2]] = val; return; }
    if (parts[0] === 'sign') { D.sign[+parts[1]][parts[2]] = val; return; }
    D[parts[0]][parts[1]] = val;
  }
  const main = $('vm-main');
  main.addEventListener('input', (e) => {
    if (e.target.dataset && e.target.dataset.ui === 'poq') { POQ = e.target.value; renderPO(); return; }
    const k = e.target.dataset && e.target.dataset.k; if (!k || e.target.type === 'checkbox') return;
    setByKey(k, e.target.value); save();
    if (k === 'hdr.project' || k === 'hdr.supplier') renderPO();
    if (TAB === 'vrijgave') updateNarr();
  });
  main.addEventListener('change', (e) => {
    const k = e.target.dataset && e.target.dataset.k; if (!k) return;
    const val = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
    setByKey(k, val); save();
    if (/\.(type|blocking|point|fstatus)$/.test(k) || k.indexOf('.fstatus') > 0) render();
  });
  document.addEventListener('click', (e) => {
    const tb = e.target.closest('[data-tab]'); if (tb) { TAB = tb.dataset.tab; render(); window.scrollTo(0, 0); return; }
    if (!e.target.closest('#vm-menu') && e.target.id !== 'vm-more') $('vm-menu').hidden = true;
    const el = e.target.closest('[data-act]'); if (!el) return;
    const a = el.dataset.act;
    if (a === 'tile') { OPEN[el.dataset.no] = !OPEN[el.dataset.no]; const y = window.scrollY; render(); window.scrollTo(0, y); }
    else if (a === 'chk') {
      const no = el.dataset.no, i = +el.dataset.i, d = pt(no); d.checks = d.checks || [];
      const k = d.checks.indexOf(i); if (k >= 0) d.checks.splice(k, 1); else d.checks.push(i);
      if (d.man === 'NA') d.man = '';
      recompute(no); save(); const y = window.scrollY; render(); window.scrollTo(0, y);
    } else if (a === 'all') {
      const no = el.dataset.no, d = pt(no); d.checks = visIdx(no); d.man = ''; recompute(no); save(); const y = window.scrollY; render(); window.scrollTo(0, y);
    } else if (a === 'man') {
      const no = el.dataset.no, d = pt(no); d.man = d.man === el.dataset.m ? '' : el.dataset.m; recompute(no); save(); const y = window.scrollY; render(); window.scrollTo(0, y);
    } else if (a === 'grp') { GRP = el.dataset.g; render(); window.scrollTo(0, 0); }
    else if (a === 'showall') { SHOWALL = !SHOWALL; render(); }
    else if (a === 'po-pick') { pickPO(el.dataset.o); }
    else if (a === 'po-del') { D.ref.pos = D.ref.pos.filter((x) => x.order !== el.dataset.o); syncPOHeader(); save(); render(); }
    else if (a === 'more') { OPEN[el.dataset.no] = true; render(); }
    else if (a === 'type') { const t = el.dataset.t, l = D.hdr.types, i = l.indexOf(t); if (i >= 0) l.splice(i, 1); else l.push(t); save(); render(); }
    else if (a === 'ph-cam') pickPhotos(el.dataset.ref, true);
    else if (a === 'ph-gal') pickPhotos(el.dataset.ref, false);
    else if (a === 'ph-del') { D.photos = D.photos.filter((p) => p.id !== el.dataset.id); save(); render(); }
    else if (a === 'x-add') { D.extra.push({ id: uid(), text: '', type: 'OBS', point: '', owner: '', due: '', fstatus: 'Open', approval: '', blocking: false }); save(); render(); }
    else if (a === 'x-del') { if (confirm('Deze bevinding en haar foto\'s verwijderen?')) { D.extra = D.extra.filter((x) => x.id !== el.dataset.id); D.photos = D.photos.filter((p) => p.ref !== 'x' + el.dataset.id); save(); render(); } }
    else if (a === 'dec') { D.rel.decision = D.rel.decision === el.dataset.d ? '' : el.dataset.d; save(); render(); }
    else if (a === 'goto') { const no = +el.dataset.no; TAB = 'checklist'; GRP = (visGroups().find((x) => x.pts.indexOf(no) >= 0) || visGroups()[0] || GROUPS[0]).id; OPEN[no] = true; render(); const t = $('pt-' + no); if (t) t.scrollIntoView({ block: 'start' }); }
    else if (a === 'sig-clear') { const sg = D.sign[+el.dataset.i]; sg.sig = ''; sg.sigW = 0; sg.sigH = 0; save(); render(); }
    else if (a === 'sig-open') openSigPad(+el.dataset.i);
  });

  /* ── onderbalk: menu, verslag, concept opslaan/laden ────────────────────── */
  $('vm-more').addEventListener('click', () => { $('vm-menu').hidden = !$('vm-menu').hidden; });
  $('vm-menu').addEventListener('click', async (e) => {
    const a = e.target.dataset && e.target.dataset.act; $('vm-menu').hidden = true; if (!a) return;
    if (a === 'save-json') download(new Blob([JSON.stringify(D)], { type: 'application/json' }), 'Bezoekverslag_concept_' + safeName(D.hdr.supplier) + '_' + D.hdr.date + '.json');
    else if (a === 'load-json') $('vm-load-file').click();
    else if (a === 'new') { if (confirm('Een nieuw verslag beginnen? Het huidige concept op dit apparaat wordt gewist.')) { D = blank(); Object.keys(OPEN).forEach((k) => delete OPEN[k]); await kvDel('draft'); TAB = 'bezoek'; render(); window.scrollTo(0, 0); toast('Nieuw verslag'); } }
  });
  $('vm-load-file').addEventListener('change', async (e) => {
    const f = e.target.files[0]; e.target.value = ''; if (!f) return;
    try {
      const j = JSON.parse(await f.text()); if (!j || !j.hdr || !j.pts) throw new Error('geen verslag');
      if (!confirm('Dit bestand vervangt het huidige concept op dit apparaat. Doorgaan?')) return;
      D = Object.assign(blank(), j); save(); render(); toast('Concept geladen');
    } catch (err) { toast('Dit bestand is geen geldig concept.'); }
  });
  const safeName = (s) => String(s || 'leverancier').replace(/[^\w\-]+/g, '_').replace(/^_+|_+$/g, '') || 'leverancier';
  function download(blob, name) {
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click(); document.body.removeChild(a); setTimeout(() => URL.revokeObjectURL(a.href), 3000);
  }
  $('vm-make').addEventListener('click', async () => {
    if (!window.docx || !window.VerslagDocx) { toast('De Word-bibliotheek is niet geladen (internet nodig). Probeer opnieuw.'); return; }
    const dv = derive();
    if (dv.cnt.open === dv.total && !confirm('Er is nog geen enkel checklistpunt beoordeeld. Toch een verslag maken?')) return;
    if (dv.cnt.open > 0 && dv.cnt.open < dv.total && !confirm(dv.cnt.open + ' punt(en) zijn nog niet beoordeeld en komen als "niet beoordeeld" in het verslag. Doorgaan?')) return;
    const btn = $('vm-make'); btn.disabled = true; toast('Verslag wordt opgemaakt…');
    try {
      const blob = await VerslagDocx.build(D, dv, { POINTS, DECISIONS, ROLES, CHECKS, unconfirmed, inScope, narrative, projDesc, matrixInfo });
      download(blob, 'Verslag_' + safeName(D.hdr.supplier) + '_' + D.hdr.date + '.docx'); toast('Verslag gedownload');
    } catch (err) { console.error(err); toast('Verslag maken mislukt: ' + (err && err.message || err)); }
    btn.disabled = false;
  });

  // Voor de auditmatrix in het verslag: status + toelichting per punt
  function matrixInfo(no, dv) {
    dv = dv || derive();
    if (isAuto(no)) {
      const a = dv.auto[no] || autoStatus(no, dv.findings);
      const miss = a.checks.filter((x) => !x[1]).map((x) => x[0]);
      return { st: a.st, remark: 'Automatisch bepaald' + (+no === 1 ? ' uit de bezoekgegevens' : +no === 23 ? ' uit het bevindingenregister' : ' uit het vrijgavebesluit') + (miss.length ? '. Ontbreekt: ' + miss.join('; ') : '') + '.' };
    }
    const x = D.pts[no] || {}, st = x.st || '', miss = st === 'O' ? unconfirmed(no) : [];
    const parts = [x.remark || '', miss.length ? 'Niet bevestigd: ' + miss.join('; ') : '', NOTES[no] || ''];
    if (+no === 7 && D.ref && D.ref.pos && D.ref.pos.length) { const h = dgHits(); parts.push(h.length ? 'Gevaarlijke stoffen in PO-regels: ' + h.map((q) => q.line + ' (' + q.cat + ')').join(', ') + '.' : 'Geen gevaarlijke stoffen in de PO-regels; ADN n.v.t.'); }
    return { st, remark: parts.filter(Boolean).join('\n') };
  }
  window.VM = { POINTS, derive, get data() { return D; } };

  /* ── start ──────────────────────────────────────────────────────────────── */
  (async function boot() {
    try { const d = await kvGet('draft'); if (d && d.hdr && d.pts) D = Object.assign(blank(), d); } catch (e) { /* geen opslag beschikbaar */ }
    render();
    loadExp();
  })();
})();
