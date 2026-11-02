/* ============================================================================
 *  IHC Expedite 2.0 — Bezoekverslag → Word (.docx)
 *  Bouwt uit de ingevulde checklist een verslag in dezelfde stijl als het
 *  voorbeeldverslag (rode titelbalk, roze labelcellen, rode kopjes, foto's
 *  bij de bevindingen, voettekst met datum). Gebruikt de docx-bibliotheek
 *  (global `docx`, geladen in index.html).
 * ========================================================================== */
(function () {
  'use strict';

  const RED = 'E60012', PINK = 'F7D6DB', PALE = 'FCEBED', GREY = 'F2F2F2', LINE = 'D9D9D9';
  const FONT = 'Calibri';
  const W = 9638;                       // inhoudsbreedte A4 (twips) bij 2 cm marge
  const STATUS_TXT = { C: 'Compliant', O: 'Observation', NC: 'Non-conformity', NA: 'N/A', '': 'Niet beoordeeld' };
  const PTS = { C: '1,0', O: '0,5', NC: '0,0', NA: '', '': '' };

  function dateNl(iso) {
    if (!iso) return '';
    const d = new Date(iso + 'T12:00:00'); if (isNaN(d)) return iso;
    return d.toLocaleDateString('nl-NL', { day: 'numeric', month: 'long', year: 'numeric' });
  }
  function bytes(dataUrl) {
    const bin = atob(dataUrl.split(',')[1]); const a = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) a[i] = bin.charCodeAt(i);
    return a;
  }

  async function build(D, dv, ctx) {
    const d = window.docx;
    const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
      ImageRun, Header, Footer, AlignmentType, PageNumber, VerticalAlign, TableLayoutType } = d;
    const H = D.hdr;

    /* ── bouwstenen ─────────────────────────────────────────────────────── */
    const run = (t, o) => new TextRun(Object.assign({ text: String(t == null ? '' : t), font: FONT, size: 20 }, o || {}));
    const para = (t, o) => {
      o = o || {};
      const lines = String(t == null ? '' : t).split('\n');
      return new Paragraph({
        spacing: { after: o.after == null ? 100 : o.after, before: o.before || 0, line: 264 },
        alignment: o.align, keepNext: o.keepNext, pageBreakBefore: o.pageBreak,
        children: lines.map((l, i) => run(l, Object.assign({ break: i ? 1 : 0 }, o.run || {})))
      });
    };
    const h1 = (t) => new Paragraph({ spacing: { before: 300, after: 110 }, keepNext: true, children: [run(t, { bold: true, color: RED, size: 28 })] });
    const h2 = (t) => new Paragraph({ spacing: { before: 200, after: 60 }, keepNext: true, children: [run(t, { bold: true, size: 22 })] });
    const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
    const thin = { style: BorderStyle.SINGLE, size: 4, color: LINE };
    const cell = (children, o) => {
      o = o || {};
      return new TableCell({
        width: { size: o.w, type: WidthType.DXA }, verticalAlign: o.v || VerticalAlign.CENTER, columnSpan: o.span,
        shading: o.fill ? { type: ShadingType.CLEAR, color: 'auto', fill: o.fill } : undefined,
        margins: { top: o.pad == null ? 70 : o.pad, bottom: o.pad == null ? 70 : o.pad, left: 110, right: 110 },
        borders: o.borders || { top: thin, bottom: thin, left: thin, right: thin },
        children: Array.isArray(children) ? children : [children]
      });
    };
    const table = (rows, widths) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: widths, layout: TableLayoutType.FIXED, rows });
    const cp = (t, o) => para(t, Object.assign({ after: 0 }, o || {}));

    // sleutel/waarde-tabel zoals in het voorbeeldverslag (roze labelkolom)
    const kv = (pairs) => !pairs.some((p) => p[1] != null && p[1] !== '') ? para('–') : table(pairs.filter((p) => p[1] != null && p[1] !== '').map((p) =>
      new TableRow({ cantSplit: true, children: [
        cell(cp(p[0], { run: { bold: true } }), { w: 2700, fill: PINK }),
        cell(cp(p[1]), { w: W - 2700 })
      ] })), [2700, W - 2700]);

    // gewone tabel met kopregel
    const grid = (head, rows, widths, size) => {
      const fs = size || 18;
      return table([
        new TableRow({ tableHeader: true, cantSplit: true, children: head.map((h, i) => cell(cp(h, { run: { bold: true, size: fs } }), { w: widths[i], fill: PINK })) })
      ].concat(rows.map((r) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(cp(c, { run: { size: fs } }), { w: widths[i], v: VerticalAlign.TOP })) }))), widths);
    };

    // afbeeldingen in een raster van twee per rij
    function photoGrid(items) {
      if (!items.length) return [];
      const colW = W / 2, maxW = 300, maxH = 330;
      const cells = items.map((p) => {
        const k = Math.min(maxW / p.w, maxH / p.h, 1);
        const w = Math.round(p.w * k), h = Math.round(p.h * k);
        return cell([
          new Paragraph({ spacing: { after: 40 }, alignment: AlignmentType.CENTER, children: [new ImageRun({ data: bytes(p.data), transformation: { width: w, height: h }, type: 'jpg' })] }),
          para(p.caption, { align: AlignmentType.CENTER, after: 0, run: { size: 16, color: '595959' } })
        ], { w: colW, borders: { top: none, bottom: none, left: none, right: none }, v: VerticalAlign.TOP });
      });
      const rows = [];
      for (let i = 0; i < cells.length; i += 2) {
        const pair = cells.slice(i, i + 2);
        if (pair.length === 1) pair.push(cell(cp(''), { w: colW, borders: { top: none, bottom: none, left: none, right: none } }));
        rows.push(new TableRow({ cantSplit: true, children: pair }));
      }
      return [table(rows, [colW, colW]), para('', { after: 60 })];
    }

    /* ── gegevens voorbereiden ──────────────────────────────────────────── */
    const photoNo = new Map(); D.photos.forEach((p, i) => photoNo.set(p.id, i + 1));
    const labelOfRef = (ref) => ref === 'g' ? 'Overig' : ref[0] === 'p' ? 'Punt #' + ref.slice(1) : ((dv.findings.find((f) => f.ref === ref) || {}).fid || 'Bevinding');
    const capOf = (p) => 'Foto ' + photoNo.get(p.id) + ' — ' + labelOfRef(p.ref) + (p.desc ? ': ' + p.desc : '');
    const decision = ctx.DECISIONS.find((x) => x[0] === D.rel.decision);
    const types = H.types.length ? H.types.join(', ') : '';
    const blocks = [];
    let n = 0; const sec = (t) => (++n) + '. ' + t;

    /* ── titelbalk + gegevens ───────────────────────────────────────────── */
    const title = 'Bezoek ' + (H.supplier || 'leverancier') + (H.location ? ' – ' + H.location : '');
    blocks.push(table([new TableRow({ children: [cell([
      cp('FORMEEL BEZOEKVERSLAG', { run: { bold: true, color: 'FFFFFF', size: 16 }, after: 40 }),
      cp(title, { run: { bold: true, color: 'FFFFFF', size: 40 } })
    ], { w: W, fill: RED, pad: 160, borders: { top: none, bottom: none, left: none, right: none } })] })], [W]));
    blocks.push(para('', { after: 120 }));
    blocks.push(kv([
      ['Onderwerp', 'Field Expediting Checklist V2.1' + (types ? ' — ' + types : '')],
      ['Leverancier / IHC Unit', H.supplier], ['Locatie', H.location], ['Datum', dateNl(H.date)],
      ['Project / Yard no.', H.project], ['PO / Delivery ref.', H.po],
      ['Field Expediter', H.expediter], ['Vertegenwoordiger leverancier', H.rep], ['Opdrachtgever (Client)', H.client],
      ['Aanwezig', H.present], ['Status', H.status]
    ]));

    /* ── 1 doel ─────────────────────────────────────────────────────────── */
    blocks.push(h1(sec('Doel van het bezoek')));
    blocks.push(para(H.goal || ('Leveranciersbezoek' + (types ? ' (' + types + ')' : '') + ' aan ' + (H.supplier || 'de leverancier') + (H.location ? ' te ' + H.location : '') + ', uitgevoerd aan de hand van de Field Expediting Checklist V2.1.')));

    /* ── 2 score + vrijgave ─────────────────────────────────────────────── */
    blocks.push(h1(sec('Auditresultaat')));
    blocks.push(kv([
      ['Audit score', dv.scored ? dv.score.toFixed(1).replace('.', ',') + '%' : 'Niet beoordeeld'],
      ['Classification', dv.cls],
      ['Beoordeelde punten', dv.scored + ' (meetellend) · ' + dv.cnt.C + ' Compliant · ' + dv.cnt.O + ' Observation · ' + dv.cnt.NC + ' Non-conformity · ' + dv.cnt.NA + ' N/A' + (dv.cnt.open ? ' · ' + dv.cnt.open + ' niet beoordeeld' : '')],
      ['Blocking NCR', dv.blocking ? 'Ja' : 'Nee'],
      ['Release decision', decision ? decision[1] + ' — ' + decision[2] : 'Nog niet vastgelegd']
    ]));
    blocks.push(para('Scoring: Compliant = 1,0 · Observation = 0,5 · Non-conformity = 0,0 · N/A telt niet mee. Classification: ≥ 95% A · ≥ 85% B · ≥ 70% C · lager D.', { before: 60, run: { size: 16, color: '595959' } }));

    /* ── 3 bevindingen met foto's ───────────────────────────────────────── */
    blocks.push(h1(sec('Vastgestelde bevindingen')));
    if (!dv.findings.length) blocks.push(para('Er zijn geen Observations of Non-conformities vastgesteld.'));
    dv.findings.forEach((f) => {
      const p = f.ptNo ? ctx.POINTS.find((x) => x.no === f.ptNo) : null;
      blocks.push(h2(f.fid + ' · ' + f.type + (f.blocking ? ' (blokkerend)' : '') + (p ? ' · #' + p.no + ' ' + p.area : '')));
      if (p) blocks.push(para(p.text, { after: 40, run: { italics: true, color: '595959', size: 18 } }));
      blocks.push(para(f.text || '(geen toelichting ingevuld)'));
      const miss = f.ptNo && ctx.unconfirmed ? ctx.unconfirmed(f.ptNo) : [];
      if (miss.length && f.type === 'OBS') blocks.push(para('Niet bevestigd: ' + miss.join('; ') + '.', { run: { size: 18, color: '404040' } }));
      blocks.push(para('Owner: ' + (f.owner || '–') + '   ·   Due date: ' + (f.due ? dateNl(f.due) : '–') + '   ·   Status: ' + f.fstatus + (f.fstatus === 'Closed' && f.approval ? '   ·   Closure approval: ' + f.approval : ''), { run: { size: 18, color: '404040' } }));
      photoGrid(D.photos.filter((x) => x.ref === f.ref).map((x) => Object.assign({ caption: capOf(x) }, x))).forEach((b) => blocks.push(b));
    });

    /* ── 4 actiepunten ──────────────────────────────────────────────────── */
    blocks.push(h1(sec('Actiepunten')));
    if (!dv.findings.length) blocks.push(para('Geen openstaande actiepunten.'));
    else blocks.push(grid(['Nr.', 'Actie / bevinding', 'Verantwoordelijke', 'Deadline', 'Status'],
      dv.findings.map((f) => [f.fid, (f.type === 'NCR' ? 'NCR: ' : 'OBS: ') + (f.text || (f.ptNo ? 'Zie punt #' + f.ptNo : '')), f.owner || '–', f.due ? dateNl(f.due) : '–', f.fstatus]),
      [800, 4338, 1900, 1500, 1100], 19));

    /* ── 5 vrijgavebesluit ──────────────────────────────────────────────── */
    blocks.push(h1(sec('Shipment release decision')));
    const decColor = D.rel.decision === 'GO' ? '1E7B4F' : D.rel.decision === 'NOGO' ? RED : 'B07700';
    blocks.push(table([new TableRow({ children: [cell([
      cp('Besluit', { run: { bold: true, color: RED, size: 18 }, after: 40 }),
      cp(decision ? decision[1] + ' — ' + decision[2] : 'Nog niet vastgelegd', { run: { bold: true, size: 26, color: decision ? decColor : '595959' } })
    ], { w: W, fill: PALE, pad: 130, borders: { top: none, bottom: none, left: none, right: none } })] })], [W]));
    blocks.push(para('', { after: 60 }));
    blocks.push(kv([
      ['Conditions / blocking reasons', D.rel.conditions], ['Required corrective action', D.rel.action],
      ['Decision owner / date', D.rel.owner], ['Final release reference / dossier', D.rel.ref]
    ]));

    /* ── 6 samenvatting (optioneel) ─────────────────────────────────────── */
    if (H.summary) { blocks.push(h1(sec('Samenvatting'))); blocks.push(para(H.summary)); }

    /* ── 7 fotobijlage ──────────────────────────────────────────────────── */
    blocks.push(h1(sec('Fotobijlage')));
    if (!D.photos.length) blocks.push(para('Er zijn geen foto\'s toegevoegd.'));
    else {
      blocks.push(grid(['Foto', 'Auditpunt / bevinding', 'Omschrijving', 'Bestand / dossierreferentie'],
        D.photos.map((p) => [String(photoNo.get(p.id)), labelOfRef(p.ref), p.desc || '', 'IMG-' + (H.date || '').replace(/-/g, '') + '-' + String(photoNo.get(p.id)).padStart(2, '0') + '.jpg']),
        [800, 2300, 4138, 2400], 18));
      const rest = D.photos.filter((p) => !dv.findings.some((f) => f.ref === p.ref));
      if (rest.length) { blocks.push(para('', { after: 60 })); photoGrid(rest.map((x) => Object.assign({ caption: capOf(x) }, x))).forEach((b) => blocks.push(b)); }
    }

    /* ── 8 ondertekening ────────────────────────────────────────────────── */
    blocks.push(h1(sec('Ondertekening')));
    const sigRows = ctx.ROLES.map((role, i) => {
      const s = D.sign[i];
      const sigCell = s.sig
        ? [new Paragraph({ children: [new ImageRun({ data: bytes(s.sig), transformation: { width: 150, height: Math.round(150 * 110 / 340) || 48 }, type: 'png' })] })]
        : [cp('')];
      return new TableRow({ cantSplit: true, height: { value: 700, rule: d.HeightRule.ATLEAST }, children: [
        cell(cp(role, { run: { bold: true } }), { w: 2700, fill: PINK }), cell(cp(s.name), { w: 2400 }), cell(sigCell, { w: 2638 }), cell(cp(dateNl(s.date)), { w: 1900 })
      ] });
    });
    blocks.push(table([new TableRow({ tableHeader: true, children: [
      cell(cp('Rol', { run: { bold: true, size: 18 } }), { w: 2700, fill: GREY }), cell(cp('Naam', { run: { bold: true, size: 18 } }), { w: 2400, fill: GREY }),
      cell(cp('Handtekening', { run: { bold: true, size: 18 } }), { w: 2638, fill: GREY }), cell(cp('Datum', { run: { bold: true, size: 18 } }), { w: 1900, fill: GREY })
    ] })].concat(sigRows), [2700, 2400, 2638, 1900]));

    /* ── bijlage: volledige auditmatrix ─────────────────────────────────── */
    blocks.push(new Paragraph({ pageBreakBefore: true, spacing: { after: 110 }, children: [run('Bijlage A — Audit matrix', { bold: true, color: RED, size: 28 })] }));
    blocks.push(grid(['#', 'Audit area', 'Audit point', 'Cat.', 'Status', 'Pts', 'Evidence / remarks'],
      ctx.POINTS.map((p) => {
        const x = D.pts[p.no] || {}; const st = x.st || '';
        const miss = st === 'O' && ctx.unconfirmed ? ctx.unconfirmed(p.no) : [];
        return [String(p.no), p.area, p.text, p.cat, STATUS_TXT[st], PTS[st], [x.remark || '', miss.length ? 'Niet bevestigd: ' + miss.join('; ') : ''].filter(Boolean).join('\n')];
      }), [450, 1500, 3000, 800, 1050, 450, 2388], 15));

    /* ── document ───────────────────────────────────────────────────────── */
    const doc = new Document({
      creator: H.expediter || 'Royal IHC', title: title,
      styles: { default: { document: { run: { font: FONT, size: 20 } } } },
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134, header: 500, footer: 500 } } },
        headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run('ROYAL IHC', { bold: true, color: RED, size: 16 }), run('  |  BEZOEKVERSLAG', { bold: true, size: 16, color: '404040' })] })] }) },
        footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
          run((H.supplier || 'Bezoek') + (H.location ? ' - ' + H.location : '') + '  |  ' + dateNl(H.date) + '  |  pagina ', { size: 16, color: '7F7F7F' }),
          new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '7F7F7F' })
        ] })] }) },
        children: blocks
      }]
    });
    return Packer.toBlob(doc);
  }

  window.VerslagDocx = { build };
})();
