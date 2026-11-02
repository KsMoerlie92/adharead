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
    [7, 'Logistics Readiness', 'Confirm packaging, labeling, handling, transport, Incoterm, address and ADN requirements.', 'Method'],
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
      pts: {}, extra: [], photos: [],
      rel: { decision: '', conditions: '', action: '', owner: '', ref: '' },
      sign: ROLES.map(() => ({ name: '', date: today(), sig: '' }))
    };
  }
  let D = blank();
  let TAB = 'bezoek';
  let GRP = 'prep';           // gekozen stap in de configurator
  const OPEN = {};            // uitgeklapte auditpunten (alleen UI-staat)

  /* ── afgeleide gegevens: score, bevindingen, advies ─────────────────────── */
  function pt(no) { return D.pts[no] || (D.pts[no] = { st: '', checks: [], man: '', remark: '', owner: '', due: '', fstatus: 'Open', approval: '', blocking: false }); }

  // Status volgt uit de vinkjes: alles bevestigd = C, deels = O; "afwijking"/"n.v.t." overrulen.
  function recompute(no) {
    const d = pt(no), n = (CHECKS[no] || []).length, t = (d.checks || []).length;
    d.st = d.man || (n && t >= n ? 'C' : t > 0 ? 'O' : '');
  }
  const unconfirmed = (no) => { const d = D.pts[no] || {}; return (CHECKS[no] || []).filter((c, i) => (d.checks || []).indexOf(i) < 0); };

  function derive() {
    let sum = 0, n = 0; const cnt = { C: 0, O: 0, NC: 0, NA: 0, open: 0 };
    POINTS.forEach((p) => {
      const s = (D.pts[p.no] || {}).st;
      if (!s) { cnt.open++; return; }
      cnt[s]++;
      if (s in PTS) { sum += PTS[s]; n++; }
    });
    const score = n ? Math.round(sum / n * 1000) / 10 : 0;
    const cls = !n ? '–' : score >= 95 ? 'A - Release Ready' : score >= 85 ? 'B - Minor Findings' : score >= 70 ? 'C - Conditional Release' : 'D - No Release';
    const findings = [];
    POINTS.forEach((p) => {
      const d = D.pts[p.no]; if (!d || (d.st !== 'O' && d.st !== 'NC')) return;
      findings.push({ kind: 'pt', key: p.no, ptNo: p.no, type: d.st === 'NC' ? 'NCR' : 'OBS', text: d.remark, owner: d.owner, due: d.due, fstatus: d.fstatus || 'Open', approval: d.approval, blocking: !!d.blocking && d.st === 'NC', ref: 'p' + p.no });
    });
    D.extra.forEach((x) => findings.push({ kind: 'x', key: x.id, ptNo: x.point ? +x.point : null, type: x.type, text: x.text, owner: x.owner, due: x.due, fstatus: x.fstatus || 'Open', approval: x.approval, blocking: !!x.blocking && x.type === 'NCR', ref: 'x' + x.id }));
    findings.forEach((f, i) => { f.fid = 'F-' + String(i + 1).padStart(2, '0'); });
    const blocking = findings.some((f) => f.blocking && f.fstatus !== 'Closed');
    const suggestion = blocking || cls.startsWith('D') ? 'NOGO' : cls.startsWith('A') ? 'GO' : n ? 'CONDITIONAL' : '';
    return { score, cls, cnt, scored: n, findings, blocking, suggestion };
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

  /* ── tabbladen ──────────────────────────────────────────────────────────── */
  function tabs() {
    const dv = derive();
    const list = [['bezoek', 'Bezoek'], ['checklist', 'Checklist', (27 - dv.cnt.open) + '/27'], ['bevind', 'Bevindingen', String(dv.findings.length)],
      ['fotos', "Foto's", String(D.photos.length)], ['vrijgave', 'Vrijgave'], ['teken', 'Ondertekening']];
    $('vm-tabs').innerHTML = list.map((t) => '<button type="button" class="vm-tab' + (t[0] === TAB ? ' on' : '') + '" data-tab="' + t[0] + '">' + t[1] + (t[2] ? '<small>' + t[2] + '</small>' : '') + '</button>').join('');
  }
  function updateBar() {
    const dv = derive();
    $('vm-bar-info').innerHTML = '<b>' + (dv.scored ? 'Score ' + pct(dv.score) + ' · ' + dv.cls.split(' - ')[0] : 'Nog niet beoordeeld') + '</b><br>' +
      (27 - dv.cnt.open) + '/27 punten · ' + dv.findings.length + ' bevinding' + (dv.findings.length === 1 ? '' : 'en') + ' · ' + D.photos.length + " foto's";
    const t = document.querySelector('.vm-tab[data-tab=checklist] small'); if (t) t.textContent = (27 - dv.cnt.open) + '/27';
    const b = document.querySelector('.vm-tab[data-tab=bevind] small'); if (b) b.textContent = String(dv.findings.length);
    const f = document.querySelector('.vm-tab[data-tab=fotos] small'); if (f) f.textContent = String(D.photos.length);
  }

  function fld(label, key, val, opts) {
    opts = opts || {};
    return '<div class="' + (opts.full ? 'full' : '') + '"><label class="f">' + label + (opts.area
      ? '<textarea data-k="' + key + '" placeholder="' + esc(opts.ph || '') + '">' + esc(val) + '</textarea>'
      : '<input type="' + (opts.type || 'text') + '" data-k="' + key + '" value="' + esc(val) + '" placeholder="' + esc(opts.ph || '') + '">') + '</label></div>';
  }

  function viewBezoek() {
    const h = D.hdr;
    return '<section class="vm-card"><h2>Bezoekgegevens</h2><div class="vm-grid">' +
      fld('Supplier / IHC Unit', 'hdr.supplier', h.supplier) + fld('Audit date', 'hdr.date', h.date, { type: 'date' }) +
      fld('Project / Yard no.', 'hdr.project', h.project) + fld('PO / Delivery ref.', 'hdr.po', h.po) +
      fld('Location', 'hdr.location', h.location) + fld('Field Expediter', 'hdr.expediter', h.expediter) +
      fld('Supplier representative', 'hdr.rep', h.rep) + fld('Client', 'hdr.client', h.client) +
      fld('Aanwezig (namen)', 'hdr.present', h.present, { full: true, ph: 'Bijv. Kris Moerlie, Jan Jansen (leverancier)' }) +
      '</div></section>' +
      '<section class="vm-card"><h2>Type bezoek</h2><div class="vm-chips">' +
      VISIT_TYPES.map((t) => '<button type="button" class="vm-chip' + (h.types.indexOf(t) >= 0 ? ' on' : '') + '" data-act="type" data-t="' + t + '">' + t + '</button>').join('') +
      '</div></section>' +
      '<section class="vm-card"><h2>Doel &amp; samenvatting</h2><div class="vm-grid">' +
      fld('Doel van het bezoek (optioneel)', 'hdr.goal', h.goal, { full: true, area: true, ph: 'Komt in het verslag onder "Doel van het bezoek"' }) +
      fld('Samenvatting / conclusie (optioneel)', 'hdr.summary', h.summary, { full: true, area: true, ph: 'Komt onderaan in het verslag' }) +
      '<div class="full"><label class="f">Status van het verslag<select data-k="hdr.status"><option' + (h.status === 'Concept' ? ' selected' : '') + '>Concept</option><option' + (h.status === 'Final' ? ' selected' : '') + '>Final</option></select></label></div>' +
      '</div></section>' +
      '<button type="button" class="vm-btn vm-btn-teal" data-tab="checklist" style="width:100%">Naar de checklist →</button>';
  }

  const STATE_ICO = { C: '✓', O: '!', NC: '✕', NA: '–', '': '' };
  const STATE_TXT = { C: 'Compliant', O: 'Observation', NC: 'Non-conformity', NA: 'N/A', '': 'Nog niet beoordeeld' };
  const grpProg = (g) => { let done = 0; g.pts.forEach((no) => { if ((D.pts[no] || {}).st) done++; }); return done; };

  function tileHtml(p, dv) {
    const d = D.pts[p.no] || {}; const st = d.st || '', open = !!OPEN[p.no];
    const chk = CHECKS[p.no] || [], ticked = (d.checks || []).length;
    let h = '<section class="vm-card vm-tile' + (st ? ' s-' + st : '') + (open ? ' open' : '') + '" id="pt-' + p.no + '">' +
      '<button type="button" class="vm-tile-head" data-act="tile" data-no="' + p.no + '" aria-expanded="' + open + '">' +
      '<span class="vm-tile-ico">' + STATE_ICO[st] + '</span>' +
      '<span class="vm-tile-main"><span class="vm-pt-top"><span class="vm-pt-no">#' + p.no + '</span><span>' + p.area + '</span><span class="vm-pt-cat">' + p.cat + '</span></span>' +
      '<span class="vm-pt-text">' + esc(p.text) + '</span></span>' +
      '<span class="vm-tile-prog">' + (st === 'NA' ? 'n.v.t.' : st === 'NC' ? 'NC' : ticked + '/' + chk.length) + '</span></button>';
    if (open) {
      h += '<div class="vm-tile-body"><div class="vm-q">Bevestig ter plaatse</div><div class="vm-chks">' +
        chk.map((c, i) => '<button type="button" class="vm-chk' + ((d.checks || []).indexOf(i) >= 0 ? ' on' : '') + '" data-act="chk" data-no="' + p.no + '" data-i="' + i + '"><span class="box">✓</span><span>' + esc(c) + '</span></button>').join('') +
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
    const dv = derive();
    const gi = Math.max(0, GROUPS.findIndex((g) => g.id === GRP)), g = GROUPS[gi];
    const step = '<div class="vm-steps">' + GROUPS.map((x, i) => {
      const done = grpProg(x), full = done === x.pts.length;
      return '<button type="button" class="vm-step' + (x.id === GRP ? ' on' : '') + (full ? ' full' : '') + '" data-act="grp" data-g="' + x.id + '"><span class="vm-step-ico">' + (full ? '✓' : x.icon) + '</span><span class="vm-step-t">' + (i + 1) + '. ' + x.title + '</span><small>' + done + '/' + x.pts.length + '</small></button>';
    }).join('') + '</div>';
    let h = step + '<div class="vm-gh"><h2>' + g.icon + ' ' + g.title + '</h2><span>Stap ' + (gi + 1) + ' van ' + GROUPS.length + ' · ' + grpProg(g) + '/' + g.pts.length + ' beoordeeld</span></div>' +
      '<p class="vm-hint" style="margin:0 0 .6rem">Tik een punt open, vink aan wat ter plaatse bevestigd is en voeg zo nodig bewijs toe.</p>';
    h += g.pts.map((no) => tileHtml(POINTS[no - 1], dv)).join('');
    h += gi < GROUPS.length - 1
      ? '<button type="button" class="vm-btn vm-btn-teal" data-act="grp" data-g="' + GROUPS[gi + 1].id + '" style="width:100%">Volgende: ' + GROUPS[gi + 1].title + ' →</button>'
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
      fld('Release decision owner / date', 'rel.owner', r.owner) + fld('Final release reference / dossier', 'rel.ref', r.ref) + '</div></section>';
  }

  function viewTeken() {
    return '<section class="vm-card"><h2>Sign-off</h2><div class="vm-grid">' + ROLES.map((role, i) => {
      const s = D.sign[i];
      return '<div class="full" style="border-top:1px solid var(--v-line);padding-top:.7rem"><b>' + role + '</b><div class="vm-row">' + fld('Naam', 'sign.' + i + '.name', s.name) + fld('Datum', 'sign.' + i + '.date', s.date, { type: 'date' }) + '</div>' +
        '<label class="f">Handtekening (optioneel)</label><canvas class="vm-sig" data-sig="' + i + '"></canvas>' +
        '<div class="vm-row"><button type="button" class="vm-btn vm-btn-sm" data-act="sig-clear" data-i="' + i + '">Wissen</button></div></div>';
    }).join('') + '</div></section>';
  }

  const VIEWS = { bezoek: viewBezoek, checklist: viewChecklist, bevind: viewBevind, fotos: viewFotos, vrijgave: viewVrijgave, teken: viewTeken };
  function render() {
    tabs(); $('vm-main').innerHTML = VIEWS[TAB](); updateBar();
    if (TAB === 'teken') initSigs();
  }

  /* ── handtekeningen ─────────────────────────────────────────────────────── */
  function initSigs() {
    document.querySelectorAll('canvas.vm-sig').forEach((cv) => {
      const i = +cv.dataset.sig, dpr = window.devicePixelRatio || 1;
      cv.width = cv.clientWidth * dpr; cv.height = cv.clientHeight * dpr;
      const g = cv.getContext('2d'); g.scale(dpr, dpr); g.lineWidth = 2.2; g.lineCap = 'round'; g.strokeStyle = '#0a1628';
      if (D.sign[i].sig) { const im = new Image(); im.onload = () => g.drawImage(im, 0, 0, cv.clientWidth, cv.clientHeight); im.src = D.sign[i].sig; }
      let drawing = false;
      const pos = (e) => { const r = cv.getBoundingClientRect(); return [e.clientX - r.left, e.clientY - r.top]; };
      cv.addEventListener('pointerdown', (e) => { drawing = true; cv.setPointerCapture(e.pointerId); const p = pos(e); g.beginPath(); g.moveTo(p[0], p[1]); e.preventDefault(); });
      cv.addEventListener('pointermove', (e) => { if (!drawing) return; const p = pos(e); g.lineTo(p[0], p[1]); g.stroke(); });
      const end = () => { if (!drawing) return; drawing = false; D.sign[i].sig = cv.toDataURL('image/png'); save(); };
      cv.addEventListener('pointerup', end); cv.addEventListener('pointercancel', end);
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
  main.addEventListener('input', (e) => { const k = e.target.dataset && e.target.dataset.k; if (!k || e.target.type === 'checkbox') return; setByKey(k, e.target.value); save(); });
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
      const no = el.dataset.no, d = pt(no); d.checks = (CHECKS[no] || []).map((c, i) => i); d.man = ''; recompute(no); save(); const y = window.scrollY; render(); window.scrollTo(0, y);
    } else if (a === 'man') {
      const no = el.dataset.no, d = pt(no); d.man = d.man === el.dataset.m ? '' : el.dataset.m; recompute(no); save(); const y = window.scrollY; render(); window.scrollTo(0, y);
    } else if (a === 'grp') { GRP = el.dataset.g; render(); window.scrollTo(0, 0); }
    else if (a === 'more') { OPEN[el.dataset.no] = true; render(); }
    else if (a === 'type') { const t = el.dataset.t, l = D.hdr.types, i = l.indexOf(t); if (i >= 0) l.splice(i, 1); else l.push(t); save(); render(); }
    else if (a === 'ph-cam') pickPhotos(el.dataset.ref, true);
    else if (a === 'ph-gal') pickPhotos(el.dataset.ref, false);
    else if (a === 'ph-del') { D.photos = D.photos.filter((p) => p.id !== el.dataset.id); save(); render(); }
    else if (a === 'x-add') { D.extra.push({ id: uid(), text: '', type: 'OBS', point: '', owner: '', due: '', fstatus: 'Open', approval: '', blocking: false }); save(); render(); }
    else if (a === 'x-del') { if (confirm('Deze bevinding en haar foto\'s verwijderen?')) { D.extra = D.extra.filter((x) => x.id !== el.dataset.id); D.photos = D.photos.filter((p) => p.ref !== 'x' + el.dataset.id); save(); render(); } }
    else if (a === 'dec') { D.rel.decision = D.rel.decision === el.dataset.d ? '' : el.dataset.d; save(); render(); }
    else if (a === 'goto') { const no = +el.dataset.no; TAB = 'checklist'; GRP = (GROUPS.find((x) => x.pts.indexOf(no) >= 0) || GROUPS[0]).id; OPEN[no] = true; render(); const t = $('pt-' + no); if (t) t.scrollIntoView({ block: 'start' }); }
    else if (a === 'sig-clear') { D.sign[+el.dataset.i].sig = ''; save(); render(); }
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
    if (dv.cnt.open === 27 && !confirm('Er is nog geen enkel checklistpunt beoordeeld. Toch een verslag maken?')) return;
    if (dv.cnt.open > 0 && dv.cnt.open < 27 && !confirm(dv.cnt.open + ' punt(en) zijn nog niet beoordeeld en komen als "niet beoordeeld" in het verslag. Doorgaan?')) return;
    const btn = $('vm-make'); btn.disabled = true; toast('Verslag wordt opgemaakt…');
    try {
      const blob = await VerslagDocx.build(D, dv, { POINTS, DECISIONS, ROLES, CHECKS, unconfirmed });
      download(blob, 'Verslag_' + safeName(D.hdr.supplier) + '_' + D.hdr.date + '.docx'); toast('Verslag gedownload');
    } catch (err) { console.error(err); toast('Verslag maken mislukt: ' + (err && err.message || err)); }
    btn.disabled = false;
  });

  window.VM = { POINTS, derive, get data() { return D; } };

  /* ── start ──────────────────────────────────────────────────────────────── */
  (async function boot() {
    try { const d = await kvGet('draft'); if (d && d.hdr && d.pts) D = Object.assign(blank(), d); } catch (e) { /* geen opslag beschikbaar */ }
    render();
  })();
})();
