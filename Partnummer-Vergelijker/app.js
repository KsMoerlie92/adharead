// ── Partnummer-vergelijker ─────────────────────────────────────────────────
// Per regel in het gekozen project met een leverancier waarin "IHC" staat:
// zoek hetzelfde Part No in de andere projecten van dezelfde groep
// (shared/project-relations.js) en toon het verschil in leverdatum.
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const esc = ExpeditingCore.esc, dmy = ExpeditingCore.dmy;
  const trim = v => String(v == null ? '' : v).trim();
  const GROUPS = (window.PROJECT_RELATIONS || []).filter(g => g && g.parent && Array.isArray(g.projects));

  let rows = [], result = [], sortKey = 'absDev', sortDir = -1;

  // Datum → lokale middernacht (de lijst bevat ISO-tijden als 22:00Z = 00:00 NL)
  function day(v) { const d = ExpeditingCore.toDate(v); return d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()) : null; }
  const diffDays = (a, b) => (a && b) ? Math.round((a - b) / 86400000) : null;
  const poLine = r => [trim(r['Order No']), trim(r['Line No']), trim(r['Release No'])].filter(Boolean).join('-');
  const proj = r => trim(r['Sub Project ID']).toUpperCase();
  const part = r => trim(r['Part No']);
  // Incoterm (kolom AE 'Delivery Terms', bv. 'DAP20 Kinderdijk') → wat de leverdatum betekent
  const incoterm = r => trim(r['Delivery Terms']);
  function incoKind(t) {
    const c = (String(t).toUpperCase().match(/\b(EXW|FCA|FAS|FOB|CPT|CIP|CFR|CIF|DAP|DPU|DAT|DDP)/) || [])[1];
    if (!c) return { code: '', cls: '', label: '' };
    if (/^D/.test(c)) return { code: c, cls: 'pv-inco-d', label: 'Komt op leverdatum bij ons aan' };
    if (/^C/.test(c)) return { code: c, cls: 'pv-inco-c', label: 'Leverancier regelt hoofdvervoer; leverdatum = verzending' };
    return { code: c, cls: 'pv-inco-ef', label: 'Wij regelen het vervoer; leverdatum = gereed bij leverancier' };
  }
  const incoCell = (t, extra) => { const k = incoKind(t); return `<td class="${extra} ${k.cls}" title="${esc(k.label ? k.label + ' — ' : '')}${esc(t)}">${k.code ? `<span class="pv-inco">${esc(k.code)}</span>` : ''}${esc(k.code ? t.replace(/^\s*[A-Z]{3}\s*(20(10|20)?)?\s*/i, '') : t) || (k.code ? '' : esc(t) || '—')}</td>`; };

  const pid = p => trim(typeof p === 'string' ? p : p && p.id).toUpperCase();
  // Alle projectnummers van een groep: hoofdproject + subprojecten
  function groupProjects(g) {
    const set = new Set([trim(g.parent).toUpperCase()]);
    g.projects.forEach(p => { if (pid(p)) set.add(pid(p)); });
    return set;
  }
  // Info per projectnummer (omschrijving/zekerheid) uit project-relations.js
  function projInfo(g, id) {
    if (id === trim(g.parent).toUpperCase()) return { zekerheid: 'hoofdproject', desc: g.name || '' };
    const p = g.projects.find(x => pid(x) === id);
    return (p && typeof p === 'object') ? p : { zekerheid: '', desc: '' };
  }

  function fillGroups() {
    const sel = $('pv-group');
    if (!GROUPS.length) { sel.innerHTML = '<option value="">— geen groepen gedefinieerd —</option>'; return; }
    sel.innerHTML = GROUPS.map((g, i) => `<option value="${i}">${esc(g.parent)}${g.name ? ' · ' + esc(g.name) : ''}</option>`).join('');
  }
  function fillProjects() {
    const g = GROUPS[+$('pv-group').value]; const sel = $('pv-project');
    if (!g) { sel.innerHTML = ''; return; }
    const present = new Set(rows.map(proj));
    const list = [...groupProjects(g)].filter(p => present.has(p));
    sel.innerHTML = list.map(p => {
      const d = rows.find(r => proj(r) === p);
      return `<option value="${esc(p)}">${esc(p)}${d && d['Sub Project Description'] ? ' · ' + esc(trim(d['Sub Project Description'])) : ''}</option>`;
    }).join('') || '<option value="">— geen projecten uit deze groep in de lijst —</option>';
    const missing = [...groupProjects(g)].filter(p => !present.has(p));
    $('pv-ginfo').innerHTML = `Hoofdproject <b>${esc(g.parent)}</b> · subprojecten: ${g.projects.map(p => {
        const i = projInfo(g, pid(p)); return `<span title="${esc(i.desc || '')}${i.bron ? ' — ' + esc(i.bron) : ''}">${esc(pid(p))}${i.zekerheid && i.zekerheid !== 'bevestigd' ? ' <span class="pv-muted">(' + esc(i.zekerheid) + ')</span>' : ''}</span>`; }).join(', ')}` +
      (missing.length ? `<br><span class="pv-muted">Nog geen regels in de bedrijfsbrede lijst voor: ${missing.map(esc).join(', ')}</span>` : '');
  }

  function compute() {
    const g = GROUPS[+$('pv-group').value], p = $('pv-project').value;
    const basis = $('pv-basis').value, showNo = $('pv-nomatch').checked;
    result = [];
    if (!g || !p) return render();
    const others = groupProjects(g); others.delete(p);
    // index Part No → regels in de andere projecten
    const idx = new Map();
    rows.forEach(r => { if (others.has(proj(r)) && part(r)) { const k = part(r); (idx.get(k) || idx.set(k, []).get(k)).push(r); } });
    rows.filter(r => proj(r) === p && /IHC/i.test(trim(r['Supplier Name'])) && part(r)).forEach(r => {
      const own = basis === 'lwr' ? day(r['Latest Wanted Receipt Date']) : day(r['Planned Delivery Date']);
      const hits = idx.get(part(r)) || [];
      if (!hits.length) {
        if (showNo) result.push({ part: part(r), desc: trim(r['Description']), sup: trim(r['Supplier Name']), po: poLine(r), own, inco: incoterm(r), oInco: '',
          oProj: '', oZek: '', oPo: '', oSup: '', oDate: null, oEst: false, dev: null, absDev: -1 });
        return;
      }
      hits.forEach(h => {
        let oDate = day(h['Planned Delivery Date']), oEst = false;
        if (!oDate) { oDate = day(h['Latest Wanted Receipt Date']); oEst = !!oDate; }
        const dev = diffDays(oDate, own);
        result.push({ part: part(r), desc: trim(r['Description']), sup: trim(r['Supplier Name']), po: poLine(r), own, inco: incoterm(r), oInco: incoterm(h),
          oProj: proj(h), oZek: projInfo(g, proj(h)).zekerheid || '', oPo: poLine(h), oSup: trim(h['Supplier Name']), oDate, oEst, dev, absDev: dev == null ? -1 : Math.abs(dev) });
      });
    });
    render();
  }

  // sec: 0 = hoofdproject, 1 = gevonden in subproject, 2 = vergelijking
  const COLS = [
    ['part', 'Part No', 0], ['desc', 'Omschrijving', 0], ['sup', 'IHC-leverancier', 0], ['po', 'PO-regel', 0], ['inco', 'Incoterm', 0], ['own', 'Leverdatum', 0],
    ['oPo', 'PO-regel', 1], ['oSup', 'Leverancier', 1], ['oInco', 'Incoterm', 1], ['oDate', 'Leverdatum', 1],
    ['dev', 'Afwijking (d)', 2],
  ];
  const SEC_CLS = ['pv-own', 'pv-rel', 'pv-cmp'];
  function render() {
    const q = trim($('pv-search').value).toLowerCase().split(/\s+/).filter(Boolean);
    let list = result.filter(o => !q.length || q.every(w => [o.part, o.desc, o.sup, o.po, o.inco, o.oProj, o.oPo, o.oSup, o.oInco].join(' ').toLowerCase().includes(w)));
    const val = (o, k) => { const v = o[k]; return v instanceof Date ? v.getTime() : (v == null || v === '' ? null : v); };
    list.sort((a, b) => {
      const x = val(a, sortKey), y = val(b, sortKey);
      if (x == null && y == null) return 0; if (x == null) return 1; if (y == null) return -1;
      return (typeof x === 'number' ? x - y : String(x).localeCompare(String(y), 'nl', { numeric: true })) * sortDir;
    });
    const sk = sortKey === 'absDev' ? 'dev' : sortKey;
    const nOwn = COLS.filter(c => c[2] === 0).length, nRel = COLS.filter(c => c[2] === 1).length;
    $('pv-thead').innerHTML =
      `<tr class="pv-band"><th class="pv-own pv-first" colspan="${nOwn}">Hoofdproject ${esc($('pv-project').value)}</th><th class="pv-rel pv-first" colspan="${nRel}">Gevonden in subproject ${esc([...new Set(list.map(o => o.oProj).filter(Boolean))].join(', ') || '—')}</th><th class="pv-cmp pv-first">Δ</th></tr>` +
      '<tr>' + COLS.map(([k, t, sec], i) =>
      `<th data-k="${k}" class="${SEC_CLS[sec]}${i && COLS[i - 1][2] !== sec ? ' pv-first' : ''}">${t}${sk === k ? `<span class="pv-arr">${sortKey === 'absDev' ? '⇅' : (sortDir > 0 ? '▲' : '▼')}</span>` : ''}</th>`).join('') + '</tr>';
    $('pv-tbody').innerHTML = list.length ? list.map(o => `<tr>
      <td class="pv-own pv-strong">${esc(o.part)}</td><td class="pv-own" title="${esc(o.desc)}">${esc(o.desc)}</td><td class="pv-own">${esc(o.sup)}</td><td class="pv-own">${esc(o.po)}</td>
      ${incoCell(o.inco, 'pv-own')}<td class="pv-own pv-strong">${o.own ? dmy(o.own) : '—'}</td>
      <td class="pv-rel pv-first" title="${esc(o.oProj)}${o.oZek ? ' (' + esc(o.oZek) + ')' : ''}">${esc(o.oPo) || '<span class="pv-muted">geen match</span>'}</td><td class="pv-rel">${esc(o.oSup)}</td>
      ${o.oProj ? incoCell(o.oInco, 'pv-rel') : '<td class="pv-rel"></td>'}
      <td class="pv-rel pv-strong">${o.oDate ? dmy(o.oDate) + (o.oEst ? ' <span class="pv-muted" title="Geen Planned Delivery Date; Latest Wanted Receipt Date gebruikt">*</span>' : '') : '—'}</td>
      <td class="pv-cmp pv-first pv-num ${o.dev > 0 ? 'pv-late' : o.dev < 0 ? 'pv-early' : ''}">${o.dev == null ? '—' : (o.dev > 0 ? '+' : '') + o.dev}</td></tr>`).join('')
      : `<tr><td colspan="${COLS.length}" class="pv-empty">${GROUPS.length ? 'Geen IHC-regels met hetzelfde Part No in de bijbehorende projecten. Staan de subprojecten al in de bedrijfsbrede lijst? (zie hierboven)' : 'Nog geen projectrelaties gedefinieerd (shared/project-relations.js).'}</td></tr>`;
    const parts = new Set(result.filter(o => o.oProj).map(o => o.part)).size;
    const late = result.filter(o => o.dev > 0).length;
    $('pv-sum').innerHTML = `<b>${parts}</b> Part No's met match · <b>${result.filter(o => o.oProj).length}</b> vergelijkingen · <b>${late}</b> later geleverd in gerelateerd project` +
      ` · afwijking = leverdatum gerelateerd project − leverdatum ${$('pv-basis').value === 'lwr' ? '(wanted) ' : ''}gekozen project`;
    $('pv-thead').querySelectorAll('th').forEach(th => th.onclick = () => {
      const k = th.dataset.k;
      if (k === 'dev') { if (sortKey === 'absDev') { sortKey = 'dev'; sortDir = -1; } else if (sortDir === -1) sortDir = 1; else { sortKey = 'absDev'; sortDir = -1; } }
      else if (sortKey === k) sortDir = -sortDir; else { sortKey = k; sortDir = 1; }
      render();
    });
    window._pvList = list;
  }

  function exportXlsx() {
    const list = window._pvList || [];
    const aoa = [COLS.map(c => (c[2] === 0 ? 'Hoofdproject · ' : c[2] === 1 ? 'Subproject · ' : '') + c[1])].concat(list.map(o =>
      [o.part, o.desc, o.sup, o.po, o.inco, o.own ? dmy(o.own) : '', o.oPo, o.oSup, o.oInco, o.oDate ? dmy(o.oDate) : '', o.dev == null ? '' : o.dev]));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Vergelijking');
    XLSX.writeFile(wb, `Partnummer-vergelijking_${$('pv-project').value || 'project'}.xlsx`);
  }

  async function init() {
    fillGroups();
    try {
      const raw = await ExpeditingData.loadRaw({ related: true }); const meta = await ExpeditingData.meta();
      rows = raw ? raw.rows : [];
      const rm = raw && raw.relatedMeta;
      $('pv-src').textContent = rows.length ? `Bedrijfsbrede lijst: stand ${meta && meta.uploaded ? dmy(meta.uploaded) : '?'}` +
        (rm ? ` · andere projecten (related-data.json): stand ${rm.uploaded ? dmy(rm.uploaded) : '?'}` : ' · nog geen related-data.json (andere projecten) — lees de uitgebreide lijst in via Admin') : 'Geen bedrijfsbrede lijst gevonden.';
    } catch (e) { $('pv-src').textContent = 'Lijst laden mislukt: ' + e.message; }
    fillProjects(); compute();
    $('pv-group').onchange = () => { fillProjects(); compute(); };
    ['pv-project', 'pv-basis', 'pv-nomatch'].forEach(id => $(id).onchange = compute);
    $('pv-search').oninput = render;
    $('pv-export').onclick = exportXlsx;
  }
  init();
})();
