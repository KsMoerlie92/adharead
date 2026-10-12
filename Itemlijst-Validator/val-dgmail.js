/**
 * val-dgmail.js  –  IHC Expedite 2.0 · Itemlijst-Validator
 *
 * Als in de (door de leverancier ingevulde) itemlijst Dangerous Goods = YES/TRUE
 * en/of Dual use = YES/TRUE staat, komt er een mail-popup om bij de leverancier de
 * bijbehorende documenten op te vragen (DGD/MSDS en/of exportclassificatie/licentie).
 *
 * De tekst is een OPVOLGING op de ontvangen itemlijst: wij zijn al geïnformeerd en
 * vragen alleen de documenten voor de als DG / dual-use gemarkeerde regels.
 *
 * Koppeling: validator.js roept na elke validatie ValDgMail.afterValidation() aan.
 * De knop "📨 DG / Dual use mail" wordt zelf in de toolbar geplaatst.
 * Leest globals uit validator.js: _valRows, COL, _valBool.
 */
(function () {
  if (window.ValDgMail) return;

  const CONTACT = 'Expedite.hol@royalihc.com';
  const LS_NAME = 'ihc_dgmail_name';
  const trim = v => String(v ?? '').trim();

  /* ─── Data uit de itemlijst ─────────────────────────────────────────── */

  function flaggedRows() {
    if (typeof _valRows === 'undefined' || !_valRows || !_valRows.length) return [];
    const out = [];
    for (const r of _valRows) {
      const c = r.cells;
      const dg = _valBool(c[COL.AA]);
      const du = _valBool(c[COL.AB]);
      if (!dg && !du) continue;
      out.push({
        deliv: trim(c[COL.A]), project: trim(c[COL.B]), po: trim(c[COL.C]),
        item: trim(c[COL.D]).replace(/^'/, ''), desc: trim(c[COL.E]), comp: trim(c[COL.H]),
        supplier: trim(c[COL.K]), eccn: trim(c[COL.AC]), dg, du,
      });
    }
    return out;
  }
  const uniq = arr => [...new Set(arr.filter(Boolean))];
  const fmtDate = d => d.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' });
  function defaultDeadline() { const d = new Date(); d.setDate(d.getDate() + 7); return d; }
  function lsGet(k) { try { return localStorage.getItem(k) || ''; } catch (e) { return ''; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }

  /* ─── Mailtekst (Engels, naar de leverancier) ───────────────────────── */

  function buildBody(rows, f) {
    const hasDG = rows.some(r => r.dg), hasDU = rows.some(r => r.du);
    const pos = uniq(rows.map(r => r.po)), projs = uniq(rows.map(r => r.project)), refs = uniq(rows.map(r => r.deliv));
    const poTxt = pos.length ? pos.join(', ') : '[PO number]';
    const prTxt = projs.length ? projs.join(', ') : '[Project / Yard number]';
    const refTxt = refs.length ? ` (delivery reference ${refs.join(', ')})` : '';

    const flagTxt = [hasDG ? 'dangerous goods' : null, hasDU ? 'dual-use' : null].filter(Boolean).join(' and/or ');
    const lines = rows.map(r => {
      const flags = [r.dg ? 'Dangerous goods: YES' : null, r.du ? 'Dual use: YES' : null,
                     r.du && r.eccn ? 'ECCN: ' + r.eccn : null].filter(Boolean).join(' | ');
      const id = [r.item ? 'Item ' + r.item : null, r.comp ? '[' + r.comp + ']' : null].filter(Boolean).join(' ');
      return `  - ${id ? id + ' – ' : ''}${r.desc || '(no description)'}  →  ${flags}`;
    });

    const L = [];
    L.push('Dear Sir / Madam,', '');
    L.push(`Thank you for the completed item list${refTxt} for the shipment under Purchase Order ${poTxt}, project ${prTxt}. We have reviewed the information you provided.`, '');
    L.push(`According to your item list, the following item(s) are classified as ${flagTxt}:`, '');
    L.push(...lines, '');
    L.push('As a follow-up to this information, we kindly request the documentation below for these item(s). We require these documents before we can issue the release for shipment notification.', '');
    if (hasDG) {
      L.push('A. Dangerous goods', '');
      L.push('* Dangerous Goods Declaration (DGD) for the applicable transport mode (IMDG / IATA / ADR)');
      L.push('* Material Safety Data Sheet (MSDS / SDS), current version, in English');
      L.push('* Per item: UN number, proper shipping name, hazard class, packing group and net quantity');
      L.push('* Confirmation that packaging and marking are approved for the dangerous goods concerned', '');
    }
    if (hasDU) {
      L.push((hasDG ? 'B' : 'A') + '. Dual-use / export control', '');
      L.push('* Classification per item under EU Regulation 2021/821, the UK Strategic Export Control Lists and the US Commerce Control List (ECCN), or the statement "item is not controlled"');
      L.push('* Copy of the applicable export licence, including licence number and validity');
      L.push('* Any end-use / end-user statements or licence conditions that apply to this delivery', '');
    }
    L.push('If the classification in your item list is incorrect or has changed, please let us know in writing per item.', '');
    L.push(`Please submit the documents by ${f.deadline || '[deadline date]'}. Shipments without the required documentation will not be released.`, '');
    L.push(`All deliveries must comply with the Royal IHC Logistic Instructions (Logistic Specification ${f.spec || '[Logistic Specification]'}), attached to this message, in particular §2 (export control classification per item) and §3 (DGD / MSDS as part of the advance delivery notification; delivery only after the release for shipment notification). Royal IHC is entitled to refuse delivery of any shipment that does not comply.`, '');
    L.push(`For questions or submissions, contact us at: ${CONTACT}`, '');
    L.push('Kind regards,', '');
    L.push(`${f.name || '[Name]'}`, 'Supply Chain / Expediting', 'Royal IHC', CONTACT);
    return L.join('\r\n');
  }

  function buildSubject(rows) {
    const refs = uniq(rows.map(r => r.deliv)), pos = uniq(rows.map(r => r.po));
    const hasDG = rows.some(r => r.dg), hasDU = rows.some(r => r.du);
    const what = [hasDG ? 'Dangerous goods' : null, hasDU ? 'Dual-use' : null].filter(Boolean).join(' / ');
    return `${what} documentation request – ${refs.join(', ') || (pos.length ? 'PO ' + pos.join(', ') : 'shipment')}`;
  }

  /* ─── .eml (concept in Outlook) ─────────────────────────────────────── */

  function buildEml(to, subject, body) {
    const b = unescape(encodeURIComponent(body));
    const b64 = btoa(b).match(/.{1,76}/g).join('\r\n');
    const subj = /[^\x20-\x7e]/.test(subject)
      ? '=?UTF-8?B?' + btoa(unescape(encodeURIComponent(subject))) + '?=' : subject;
    return [
      to ? `To: ${to}` : 'To: ', `Subject: ${subj}`, 'MIME-Version: 1.0', 'X-Unsent: 1',
      'Content-Type: text/plain; charset=utf-8', 'Content-Transfer-Encoding: base64', '', b64,
    ].join('\r\n');
  }
  function download(content, name, mime) {
    const url = URL.createObjectURL(new Blob([content], { type: mime }));
    const a = Object.assign(document.createElement('a'), { href: url, download: name });
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  /* ─── Popup ─────────────────────────────────────────────────────────── */

  const css = (el, s) => Object.assign(el.style, s);
  const mk = (tag, s, txt) => { const e = document.createElement(tag); if (s) css(e, s); if (txt != null) e.textContent = txt; return e; };

  function openPopup() {
    const rows = flaggedRows();
    if (!rows.length) { alert('Geen regels met Dangerous Goods = YES of Dual use = YES gevonden.'); return; }
    document.getElementById('val-dg-overlay')?.remove();

    const f = {
      deadline: fmtDate(defaultDeadline()),
      name: lsGet(LS_NAME),
      spec: (uniq(rows.map(r => r.project))[0] || '[Project]') + '_2025, rev. 01-07-2025',
    };
    const supplier = uniq(rows.map(r => r.supplier))[0] || '';

    const ov = mk('div', { position: 'fixed', inset: '0', background: 'rgba(0,0,0,.65)', zIndex: '100000',
      display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '16px' });
    ov.id = 'val-dg-overlay';
    const box = mk('div', { background: 'var(--card,#0F2040)', color: 'var(--text,#e8edf5)', border: '1px solid var(--border,#1e3a6e)',
      borderRadius: '10px', width: 'min(860px,100%)', maxHeight: '92vh', overflow: 'auto', padding: '18px 20px',
      fontFamily: 'Barlow, sans-serif', fontSize: '.88rem' });

    const nDG = rows.filter(r => r.dg).length, nDU = rows.filter(r => r.du).length;
    box.append(mk('div', { fontSize: '1.05rem', fontWeight: '700', marginBottom: '4px' }, '📨 Documenten opvragen bij leverancier — Dangerous goods / Dual use'));
    box.append(mk('div', { color: 'var(--muted,#8da2c9)', marginBottom: '12px' },
      `In de ontvangen itemlijst staan ${nDG} regel(s) met Dangerous Goods = YES en ${nDU} regel(s) met Dual use = YES` +
      (supplier ? ` (leverancier: ${supplier})` : '') + '. Controleer de velden en de tekst hieronder.'));

    const inp = (label, val, type) => {
      const w = mk('label', { display: 'flex', flexDirection: 'column', gap: '3px', flex: '1', minWidth: '180px' });
      w.append(mk('span', { fontSize: '.7rem', textTransform: 'uppercase', color: 'var(--muted,#8da2c9)' }, label));
      const i = mk('input', { background: 'rgba(255,255,255,.06)', color: 'inherit', border: '1px solid var(--border,#1e3a6e)', borderRadius: '5px', padding: '6px 8px' });
      i.type = type || 'text'; i.value = val; w.append(i); return [w, i];
    };
    const [wTo, iTo] = inp('Aan (e-mail leverancier)', '');
    iTo.placeholder = 'naam@leverancier.com';
    const [wNm, iNm] = inp('Naam afzender', f.name);
    const [wDl, iDl] = inp('Uiterste datum documenten', f.deadline);
    const [wSp, iSp] = inp('Logistic Specification', f.spec);
    const r1 = mk('div', { display: 'flex', gap: '10px', flexWrap: 'wrap', marginBottom: '10px' }); r1.append(wTo, wNm);
    const r2 = mk('div', { display: 'flex', gap: '10px', flexWrap: 'wrap', marginBottom: '10px' }); r2.append(wDl, wSp);
    box.append(r1, r2);

    const subj = mk('input', { width: '100%', background: 'rgba(255,255,255,.06)', color: 'inherit', border: '1px solid var(--border,#1e3a6e)', borderRadius: '5px', padding: '6px 8px', marginBottom: '10px', boxSizing: 'border-box' });
    subj.value = buildSubject(rows);
    box.append(mk('div', { fontSize: '.7rem', textTransform: 'uppercase', color: 'var(--muted,#8da2c9)', marginBottom: '3px' }, 'Onderwerp'), subj);

    const ta = mk('textarea', { width: '100%', height: '340px', background: 'rgba(255,255,255,.06)', color: 'inherit', border: '1px solid var(--border,#1e3a6e)', borderRadius: '5px', padding: '8px', fontFamily: 'Consolas, monospace', fontSize: '.8rem', lineHeight: '1.45', boxSizing: 'border-box' });
    let touched = false;
    const regen = () => { f.name = iNm.value.trim(); f.deadline = iDl.value.trim(); f.spec = iSp.value.trim(); if (!touched) ta.value = buildBody(rows, f); };
    [iNm, iDl, iSp].forEach(i => i.addEventListener('input', regen));
    ta.addEventListener('input', () => { touched = true; });
    regen();
    box.append(mk('div', { fontSize: '.7rem', textTransform: 'uppercase', color: 'var(--muted,#8da2c9)', marginBottom: '3px' }, 'Mailtekst (aan te passen)'), ta);
    box.append(mk('div', { color: 'var(--amber,#f59e0b)', fontSize: '.78rem', margin: '8px 0' },
      '📎 Voeg de Royal IHC Logistic Instructions zelf als bijlage toe in Outlook voordat je verstuurt.'));

    const bar = mk('div', { display: 'flex', gap: '8px', justifyContent: 'flex-end', flexWrap: 'wrap' });
    const btn = (t, primary, fn) => { const b = mk('button', { cursor: 'pointer', borderRadius: '6px', padding: '8px 14px', fontWeight: '700',
      border: primary ? 'none' : '1px solid var(--border,#1e3a6e)', background: primary ? 'var(--ihc-red,#D91F2C)' : 'transparent', color: '#fff' }, t); b.onclick = fn; return b; };
    const stamp = () => { if (iNm.value.trim()) lsSet(LS_NAME, iNm.value.trim()); };
    bar.append(
      btn('Sluiten', false, () => ov.remove()),
      btn('📋 Kopieer tekst', false, async () => {
        try { await navigator.clipboard.writeText(ta.value); } catch (e) { ta.select(); document.execCommand('copy'); }
      }),
      btn('✉️ Open als concept in Outlook (.eml)', true, () => {
        stamp();
        download(buildEml(iTo.value.trim(), subj.value, ta.value), 'DG-DualUse documentatie aanvraag.eml', 'message/rfc822');
      })
    );
    box.append(bar);
    ov.append(box);
    ov.addEventListener('click', e => { if (e.target === ov) ov.remove(); });
    document.body.appendChild(ov);
  }

  /* ─── Knop in toolbar + automatische popup na validatie ─────────────── */

  let _shownFor = '';
  function ensureButton() {
    const tb = document.querySelector('.val-toolbar');
    if (!tb) return null;
    let b = document.getElementById('btn-val-dgmail');
    if (!b) {
      b = document.createElement('button');
      b.id = 'btn-val-dgmail'; b.className = 'btn-val';
      b.style.cssText = 'background:#b45309;color:#fff;display:none';
      b.title = 'Documenten opvragen bij de leverancier voor regels met Dangerous Goods / Dual use = YES';
      b.onclick = openPopup;
      const anchor = document.getElementById('btn-val-lparts');
      anchor && anchor.parentNode === tb ? anchor.after(b) : tb.appendChild(b);
    }
    return b;
  }

  function afterValidation() {
    const rows = flaggedRows();
    const b = ensureButton();
    if (b) {
      b.style.display = rows.length ? '' : 'none';
      b.textContent = `📨 DG / Dual use mail (${rows.length})`;
    }
    if (!rows.length) { _shownFor = ''; return; }
    const sig = rows.map(r => r.deliv + '|' + r.item + '|' + r.dg + '|' + r.du).join(';');
    if (sig === _shownFor) return;          // niet opnieuw openen bij herhaald valideren zonder wijziging
    _shownFor = sig;
    setTimeout(openPopup, 400);
  }

  window.ValDgMail = { afterValidation, open: openPopup, flaggedRows, buildBody, buildSubject };
})();
