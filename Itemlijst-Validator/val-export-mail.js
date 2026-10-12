/**
 * val-export-mail.js  –  IHC Expedite 2.0 · Itemlijst-Validator
 *
 * Vervangt de "Export naar Excel" download door een mail-popup die
 * een kant-en-klare .eml klaarzet voor de Power Automate koppeling.
 *
 * Stroom:
 *   1. Gebruiker klikt "Versturen naar Moederlijst"
 *   2. Popup toont de vooraf ingevulde mail ter controle
 *   3. Gebruiker klikt "Mail klaarzetten"
 *   4. .eml downloadt → dubbelklik opent in Outlook als bewerkbaar concept
 *      (X-Unsent:1) met bijlage erin → gebruiker bewerkt evt. en klikt Verzenden
 *   5. Power Automate pikt de mail op via het subject-sleutelwoord "Itemlijst"
 *      en schrijft de bijlage terug naar de Moederlijst
 *
 * Subject-formaat:  "{Kolom A}  Itemlijst  {Kolom K}"
 *   Voorbeeld:      "1321-010 Itemlijst Alfa Laval"
 *
 * Bijlage-bestandsnaam:  "{Kolom A}_Itemlijst_{Kolom K}.xlsx"
 *
 * Ontvanger:  m.wendels@royalihc.com
 *
 * Zelflaadpatroon: validator.js laadt dit script dynamisch,
 * identiek aan val-mailgen.js / val-crossref.js.
 */

(function () {
  if (window.__valExportMailLoaded) return;
  window.__valExportMailLoaded = true;

  /* ─── Configuratie ────────────────────────────────────────────────────── */

  const RECIPIENT  = 'm.wendels@royalihc.com';
  const TRIGGER_KW = 'ITEMLIJST';       // Power Automate trigger-sleutelwoord (was 'Itemlijst')
  const SENDER_LBL = 'IHC Expedite 2.0';

  /* ─── Export-layout (HIER de volgorde aanpassen) ──────────────────────────
   * Volgt SHIPMENT REFERENCE ITEMLIJST – SUPPLIER TEMPLATE V1.2:
   *   rij 1 = groepslabels, rij 2 = kopregel, data vanaf rij 3,
   *   Excel-tabel "Table1" over kolom A t/m Y (zoals in de template);
   *   kolommen erná (compliance / transport) staan als losse cellen naast de tabel.
   *
   * Elke regel = één kolom in de bijlage, in deze volgorde:
   *   k    : interne sleutel = kolomletter in de validator (A…AH)
   *   h    : kopnaam in de bijlage
   *   g    : groepslabel in rij 1 (optioneel)
   *   bool : true → wordt als echte TRUE/FALSE weggeschreven (YES/NO → TRUE/FALSE)
   *   num  : true → echt getal (nooit tekst); fmt = Excel-getalnotatie (bv. '0.00')
   *   text : true → altijd tekst (bv. HS-code, Item)
   *   apos : true → tekst begint met één ' (voorkomt datum-interpretatie bij plakken, bv. '1-1)
   *   { blank:true } = lege kolom (houdt de letters van de template op hun plek)
   * Alle kolommen blijven altijd staan, ook als ze leeg zijn.
   */
  const SHEET_NAME     = 'Sheet1';
  const TABLE_NAME     = 'Table1';
  const HEADER_ROW     = 2;      // kopregel (1-gebaseerd); rij 1 = groepslabels
  const TABLE_LAST_KEY = 'Y';    // laatste kolom binnen de tabel (template: A:Y)
  const EXPORT_LAYOUT = [
    { k:'A',  h:'Delivery ref.',            g:'IHC' },
    { k:'B',  h:'Project',                  g:'IHC' },
    { k:'C',  h:'IHC PO',                   g:'Supplier' },
    { k:'D',  h:'Item',                     g:'Supplier / IHC', text:true, apos:true },
    { k:'E',  h:'Item description',         g:'Supplier' },
    { k:'F',  h:'Quantity',                 g:'Supplier', num:true },
    { k:'G',  h:'Unit of measure',          g:'Supplier' },
    { k:'H',  h:'Component (Mark/Label)',   g:'Supplier' },
    { k:'I',  h:'Supplier article number',  g:'Supplier' },
    { k:'J',  h:'Serial number',            g:'Supplier' },
    { k:'K',  h:'Supplier name',            g:'Supplier' },
    { k:'L',  h:'Manufacturer name',        g:'Supplier' },
    { k:'M',  h:'Material',                 g:'Supplier' },
    { k:'N',  h:'Country of origin',        g:'Supplier' },
    { k:'O',  h:'Hs-code',                  g:'Supplier', text:true },
    { k:'P',  h:'Value pc (EUR)',           g:'Supplier', num:true, fmt:'0.00' },
    { k:'Q',  h:'Value total',              g:'Supplier', num:true, fmt:'0.00' },
    { k:'R',  h:'Collo',                    g:'Supplier' },
    { k:'S',  h:'Type of packaging',        g:'Supplier' },
    { k:'T',  h:'Length cm',                g:'Supplier', num:true },
    { k:'U',  h:'Width cm',                 g:'Supplier', num:true },
    { k:'V',  h:'Height cm',                g:'Supplier', num:true },
    { k:'W',  h:'Volume m3',                g:'IHC', num:true, fmt:'0.0000' },
    { k:'X',  h:'Weight gross kg',          g:'Supplier', num:true },
    { k:'Y',  h:'Weight nett kg',           g:'Supplier', num:true },
    { blank:true },
    { k:'AA', h:'Dangerous Goods yes/no',   g:'Mandatory compliance information', bool:true },
    { k:'AB', h:'Dual use yes/no',          bool:true },
    { blank:true },
    { k:'AC', h:'ECCN code' },
    { k:'AE', h:'Container number',         g:'Optional transport information' },
    { k:'AF', h:'Seal number' },
    { k:'AG', h:'Remarks supplier' },
    { k:'AH', h:'Inspection Level',         g:'IHC' },
  ];

  // Itemlijst-sleutels die het onderwerp bepalen (kolomletters)
  const COL_A = 'A';   // Delivery ref. → eerste deel subject
  const COL_K = 'K';   // Supplier      → derde deel subject

  /* ─── Hulpfuncties ────────────────────────────────────────────────────── */

  const trim = v => String(v ?? '').trim();

  /** Meest voorkomende niet-lege waarde in een kolom */
  function dominant(rows, col) {
    const counts = {};
    for (const r of rows) {
      const v = trim(r[col]);
      if (v) counts[v] = (counts[v] || 0) + 1;
    }
    return Object.entries(counts).sort((a, b) => b[1] - a[1])[0]?.[0] ?? '';
  }

  /** Bouw het e-mailonderwerp */
  function buildSubject(rows) {
    const delivRef = dominant(rows, COL_A);
    const supplier = dominant(rows, COL_K);
    return [delivRef, TRIGGER_KW, supplier].filter(Boolean).join(' ');
  }

  /** Bestandsnaam voor de bijlage */
  function buildFilename(rows) {
    const delivRef = dominant(rows, COL_A).replace(/[/\\:*?"<>|]/g, '-');
    const supplier = dominant(rows, COL_K).replace(/[/\\:*?"<>|]/g, '-');
    return [delivRef, TRIGGER_KW, supplier].filter(Boolean).join(' ') + '.xlsx';
  }

  /** YES/NO/JA/NEE/TRUE/FALSE/1/0 → true/false (leeg → false) */
  function toBool(v) {
    if (v === true || v === 1) return true;
    const t = String(v ?? '').trim().toLowerCase();
    return ['true', 'yes', 'y', 'ja', 'j', '1', 'x', 'waar'].includes(t);
  }

  /** Tekst → getal ("1.234,56", "€ 5" → getal); geen geldig getal → null */
  function toNum(v) {
    const t = String(v ?? '').replace(/[^\d.,\-]/g, '');
    if (!/\d/.test(t)) return null;
    let s = t; const lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
    if (lc > -1 && ld > -1) s = lc > ld ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
    else if (lc > -1) s = s.replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? null : n;
  }

  /** Genereer Excel (base64) via SheetJS — structuur volgens EXPORT_LAYOUT */
  function rowsToXlsxBase64(rows) {
    const cleanCell = v => (v === undefined || v === null) ? '' : (typeof v === 'string' ? (v.trim() === '' ? '' : v) : v);

    const row1 = EXPORT_LAYOUT.map(c => (c.g || ''));
    const row2 = EXPORT_LAYOUT.map(c => c.blank ? '' : c.h);
    const body = rows.map(r => EXPORT_LAYOUT.map(c => {
      if (c.blank) return '';
      if (c.bool) return toBool(r[c.k]);
      let v = cleanCell(r[c.k]);
      if (v === '') return '';
      if (c.num) { const nv = (typeof v === 'number') ? v : toNum(v); return nv === null ? v : nv; }
      if (c.text) v = String(v);
      if (c.apos) v = "'" + String(v).replace(/^'+/, '');
      return v;
    }));
    const wsData = [row1, row2, ...body];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Kolombreedtes op basis van inhoud
    ws['!cols'] = EXPORT_LAYOUT.map((c, i) => {
      if (c.blank) return { wch: 3 };
      const maxLen = Math.max(c.h.length, ...rows.map(r => trim(r[c.k]).length).slice(0, 50));
      return { wch: Math.min(Math.max(maxLen + 2, 8), 40) };
    });

    // Celtype/getalnotatie per kolom: getallen als getal (met formaat), tekstkolommen als tekst
    EXPORT_LAYOUT.forEach((c, ci) => {
      if (c.blank) return;
      for (let ri = 0; ri < rows.length; ri++) {
        const addr = XLSX.utils.encode_cell({ r: ri + HEADER_ROW, c: ci });
        const cell = ws[addr]; if (!cell) continue;
        if (c.num && cell.t === 'n' && c.fmt) cell.z = c.fmt;
        if (c.text && cell.t === 's') cell.z = '@';
      }
    });

    XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);

    // Tabelkolommen = EXPORT_LAYOUT t/m TABLE_LAST_KEY (moeten gelijk zijn aan de kopcellen)
    const last = EXPORT_LAYOUT.findIndex(c => c.k === TABLE_LAST_KEY);
    const tableNames = EXPORT_LAYOUT.slice(0, last + 1).map(c => c.h);

    // Als echte Excel-tabel wegschrijven (nodig voor Power Automate).
    // Lukt dat niet, dan valt hij terug op een gewoon werkblad i.p.v. te crashen.
    try {
      return sheetToTableXlsxBase64(wb, tableNames, rows.length);
    } catch (e) {
      console.warn('Tabel-injectie mislukt, export als gewoon werkblad:', e);
      return XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
    }
  }

  /**
   * SheetJS schrijft geen echte Excel-tabel (ListObject). Die injecteren we hier
   * zelf in de xlsx-zip via XLSX.CFB: table1.xml + relatie + content-type + tableParts.
   * Resultaat: een tabel met de naam TABLE_NAME, zoals Ctrl+T in Excel.
   */
  function sheetToTableXlsxBase64(wb, cols, nRows) {
    const xesc = (s) => String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

    // cols zijn al uniek + niet-leeg gemaakt in rowsToXlsxBase64 (kop == tabelkolom)
    const names = cols;

    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const cfb = XLSX.CFB.read(new Uint8Array(buf), { type: 'array' });
    const find = (p) => {
      const i = cfb.FullPaths.findIndex(f => f.replace(/^Root Entry\//, '') === p);
      return i < 0 ? null : cfb.FileIndex[i];
    };
    const readTxt = (p) => {
      const f = find(p); if (!f) throw new Error('ontbrekend deel: ' + p);
      const c = f.content;
      let s = ''; for (let i = 0; i < c.length; i++) s += String.fromCharCode(c[i]);
      return decodeURIComponent(escape(s));   // UTF-8 → tekst
    };
    const write = (p, s) => {
      const u = unescape(encodeURIComponent(s));            // tekst → UTF-8
      const a = new Uint8Array(u.length);
      for (let i = 0; i < u.length; i++) a[i] = u.charCodeAt(i) & 0xFF;
      XLSX.CFB.utils.cfb_add(cfb, '/' + p, a);
    };

    const lastCol = XLSX.utils.encode_col(names.length - 1);
    const ref = 'A' + HEADER_ROW + ':' + lastCol + Math.max(nRows + HEADER_ROW, HEADER_ROW + 1);   // minimaal 1 (lege) datarij

    // 1. de tabel zelf
    const tcols = names.map((n, i) => '<tableColumn id="' + (i + 1) + '" name="' + xesc(n) + '"/>').join('');
    write('xl/tables/table1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" ' +
      'name="' + TABLE_NAME + '" displayName="' + TABLE_NAME + '" ref="' + ref + '" totalsRowShown="0">' +
      '<autoFilter ref="' + ref + '"/>' +
      '<tableColumns count="' + names.length + '">' + tcols + '</tableColumns>' +
      '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" ' +
      'showRowStripes="1" showColumnStripes="0"/></table>');

    // 2. werkblad koppelen (eigen autoFilter weg — de tabel regelt dat zelf)
    let sheet = readTxt('xl/worksheets/sheet1.xml').replace(/<autoFilter[^>]*\/>/, '');
    sheet = sheet.replace('</worksheet>',
      '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>');
    write('xl/worksheets/sheet1.xml', sheet);

    // 3. relatie werkblad → tabel
    write('xl/worksheets/_rels/sheet1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" ' +
      'Target="../tables/table1.xml"/></Relationships>');

    // 4. content-type registreren
    write('[Content_Types].xml', readTxt('[Content_Types].xml').replace('</Types>',
      '<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>'));

    return XLSX.CFB.write(cfb, { fileType: 'zip', type: 'base64', compression: true });
  }

  /** Bouw de e-mailtekst (body) */
  function buildBody(rows, subject) {
    const delivRef = dominant(rows, COL_A);
    const supplier = dominant(rows, COL_K);
    const collos   = [...new Set(rows.map(r => trim(r['R'])).filter(Boolean))];
    const totalVal = rows.reduce((s, r) => {
      const v = parseFloat(String(r['Q'] ?? '').replace(',', '.'));
      return s + (isNaN(v) ? 0 : v);
    }, 0);

    return [
      `Geachte,`,
      ``,
      `Bijgesloten de gevalideerde itemlijst voor verwerking in de Moederlijst.`,
      ``,
      `Delivery ref. : ${delivRef || '–'}`,
      `Leverancier   : ${supplier || '–'}`,
      `Aantal regels : ${rows.length}`,
      collos.length ? `Collo nummers : ${collos.join(', ')}` : null,
      totalVal > 0  ? `Totaalwaarde  : EUR ${totalVal.toLocaleString('nl-NL', { minimumFractionDigits: 2 })}` : null,
      ``,
      `Onderwerp bevat "${TRIGGER_KW}" als trigger voor Power Automate.`,
      ``,
      `Automatisch gegenereerd door IHC Expedite 2.0.`,
    ].filter(l => l !== null).join('\r\n');
  }

  /** Genereer .eml bestandsinhoud (RFC 2822 + MIME multipart) */
  function buildEml(to, subject, body, xlsxBase64, attachFilename) {
    const boundary = 'IHC_EXPEDITE_' + Date.now().toString(36).toUpperCase();
    // RFC 2045: base64-regels van max 76 tekens
    const b64Lines = (xlsxBase64.match(/.{1,76}/g) || [xlsxBase64]).join('\r\n');

    // Encoded subject voor non-ASCII ondersteuning
    const encSubject = subject;  // ASCII-veilig voor typische delivery refs + namen

    return [
      `To: ${to}`,
      `Subject: ${encSubject}`,
      `Date: ${new Date().toUTCString().replace('GMT', '+0000')}`,
      `MIME-Version: 1.0`,
      `X-Unsent: 1`,
      `Content-Type: multipart/mixed; boundary="${boundary}"`,
      `X-Generator: IHC-Expedite-2.0`,
      ``,
      `--${boundary}`,
      `Content-Type: text/plain; charset=utf-8`,
      `Content-Transfer-Encoding: 7bit`,
      ``,
      body,
      ``,
      `--${boundary}`,
      `Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`,
      `Content-Disposition: attachment; filename="${attachFilename}"`,
      `Content-Transfer-Encoding: base64`,
      ``,
      b64Lines,
      ``,
      `--${boundary}--`,
    ].join('\r\n');
  }

  /** Download een blob als bestand */
  function downloadBlob(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url  = URL.createObjectURL(blob);
    const a    = Object.assign(document.createElement('a'), {
      href: url, download: filename,
    });
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  /* ─── Modal: compose-venster ──────────────────────────────────────────── */

  function css(el, s) { Object.assign(el.style, s); return el; }

  function field(label, value, editable = false) {
    const wrap = css(document.createElement('div'), {
      borderBottom: '1px solid rgba(30,58,110,0.5)',
      padding: '8px 0', display: 'flex', gap: '10px', alignItems: 'flex-start',
    });
    const lbl = css(document.createElement('span'), {
      minWidth: '88px', color: 'var(--ihc-muted,#6b7a99)',
      fontSize: '0.8rem', paddingTop: '2px', flexShrink: '0',
      textTransform: 'uppercase', letterSpacing: '0.04em',
      fontFamily: 'Barlow Condensed, Barlow, sans-serif',
    });
    lbl.textContent = label;

    if (editable) {
      const inp = css(document.createElement(
        value.length > 80 ? 'textarea' : 'input'
      ), {
        flex: '1', background: 'transparent',
        color: 'var(--text,#e8edf5)', border: 'none', outline: 'none',
        fontSize: '0.88rem', fontFamily: 'Barlow, sans-serif',
        resize: 'vertical', lineHeight: '1.5',
      });
      inp.value = value;
      if (inp.tagName === 'TEXTAREA') inp.rows = 4;
      wrap.append(lbl, inp);
      wrap._input = inp;
    } else {
      const val = css(document.createElement('span'), {
        flex: '1', color: 'var(--text,#e8edf5)',
        fontSize: '0.88rem', lineHeight: '1.5', wordBreak: 'break-word',
      });
      val.textContent = value;
      wrap.append(lbl, val);
    }
    return wrap;
  }

  function showComposeModal(rows, onSend) {
    /* ── Data voorbereiden ── */
    const subject  = buildSubject(rows);
    const filename = buildFilename(rows);
    const body     = buildBody(rows, subject);
    const collos   = [...new Set(rows.map(r => trim(r['R'])).filter(Boolean))];

    /* ── Overlay ── */
    const overlay = css(document.createElement('div'), {
      position: 'fixed', inset: '0',
      background: 'rgba(10,22,40,0.90)',
      display: 'flex', alignItems: 'center', justifyContent: 'center',
      zIndex: '9999', padding: '16px',
    });

    /* ── Window ── */
    const win = css(document.createElement('div'), {
      background: 'var(--ihc-mid,#0F2040)',
      border: '1px solid var(--ihc-steel,#1e3a6e)',
      borderRadius: '10px', width: '100%', maxWidth: '580px',
      maxHeight: '92vh', overflow: 'hidden',
      display: 'flex', flexDirection: 'column',
      fontFamily: 'Barlow, sans-serif',
      boxShadow: '0 12px 48px rgba(0,0,0,0.6)',
    });

    /* ── Titelbalk ── */
    const titlebar = css(document.createElement('div'), {
      background: 'var(--ihc-navy,#0A1628)',
      borderBottom: '1px solid var(--ihc-steel,#1e3a6e)',
      padding: '12px 18px', display: 'flex',
      alignItems: 'center', gap: '10px',
    });
    titlebar.innerHTML = `
      <span style="font-size:1.05rem;color:var(--ihc-teal,#00B4D8);">📧</span>
      <span style="font-weight:700;color:var(--text,#e8edf5);font-size:0.95rem;">
        Mail klaarzetten – Moederlijst update
      </span>
      <span style="margin-left:auto;font-size:0.75rem;color:var(--ihc-muted,#6b7a99);">
        via Power Automate
      </span>
    `;

    /* ── Body (velden) ── */
    const body_el = css(document.createElement('div'), {
      padding: '0 18px', overflowY: 'auto', flex: '1',
    });

    const fTo      = field('Aan',       RECIPIENT);
    const fFrom    = field('Van',       'Je eigen Outlook-account');
    const fSubjWrap = field('Onderwerp', subject, true);
    const fSubjInp  = fSubjWrap._input;

    /* Bijlage-badge */
    const attachWrap = css(document.createElement('div'), {
      borderBottom: '1px solid rgba(30,58,110,0.5)',
      padding: '8px 0', display: 'flex', alignItems: 'center', gap: '8px',
    });
    const attachLbl = css(document.createElement('span'), {
      minWidth: '88px', color: 'var(--ihc-muted,#6b7a99)',
      fontSize: '0.8rem', textTransform: 'uppercase',
      letterSpacing: '0.04em',
      fontFamily: 'Barlow Condensed, Barlow, sans-serif',
    });
    attachLbl.textContent = 'Bijlage';
    const attachBadge = css(document.createElement('span'), {
      background: 'rgba(0,180,216,0.12)', border: '1px solid rgba(0,180,216,0.3)',
      borderRadius: '4px', padding: '3px 10px',
      color: 'var(--ihc-teal,#00B4D8)', fontSize: '0.8rem',
      fontFamily: 'JetBrains Mono, monospace',
    });
    attachBadge.textContent = `📎 ${filename}`;
    attachWrap.append(attachLbl, attachBadge);

    /* Info-lijn */
    const infoLine = css(document.createElement('div'), {
      padding: '10px 0 4px',
      fontSize: '0.78rem', color: 'var(--ihc-muted,#6b7a99)',
      lineHeight: '1.6',
    });
    infoLine.innerHTML =
      `<b style="color:#94a3b8">${rows.length} regels</b>` +
      (collos.length ? ` &nbsp;·&nbsp; Collo: <b style="color:#94a3b8">${collos.join(', ')}</b>` : '') +
      `<br>` +
      `Onderwerp bevat <b style="color:var(--ihc-teal,#00B4D8)">"${TRIGGER_KW}"</b> ` +
      `als Power Automate trigger.`;

    /* Body-preview */
    const bodyPreviewWrap = field('Bericht', body);

    body_el.append(fTo, fFrom, fSubjWrap, attachWrap, infoLine, bodyPreviewWrap);

    /* ── Footer ── */
    const footer = css(document.createElement('div'), {
      padding: '14px 18px',
      borderTop: '1px solid var(--ihc-steel,#1e3a6e)',
      display: 'flex', gap: '10px', alignItems: 'center',
    });

    const sendBtn = css(document.createElement('button'), {
      background: 'var(--ihc-teal,#00B4D8)', color: 'var(--ihc-navy,#0A1628)',
      border: 'none', borderRadius: '6px', padding: '10px 22px',
      fontWeight: '700', fontSize: '0.92rem', cursor: 'pointer',
      fontFamily: 'Barlow, sans-serif', display: 'flex', gap: '7px',
      alignItems: 'center',
    });
    sendBtn.innerHTML = '📨 Mail klaarzetten';

    const cancelBtn = css(document.createElement('button'), {
      background: 'transparent', color: '#a0b0c8',
      border: '1px solid var(--ihc-steel,#1e3a6e)',
      borderRadius: '6px', padding: '10px 18px',
      fontSize: '0.88rem', cursor: 'pointer',
      fontFamily: 'Barlow, sans-serif',
    });
    cancelBtn.textContent = 'Annuleren';

    const statusMsg = css(document.createElement('span'), {
      marginLeft: 'auto', fontSize: '0.8rem', color: '#a0b0c8',
    });

    footer.append(sendBtn, cancelBtn, statusMsg);
    win.append(titlebar, body_el, footer);
    overlay.appendChild(win);
    document.body.appendChild(overlay);

    /* ── Events ── */
    const close = () => {
      if (document.body.contains(overlay)) document.body.removeChild(overlay);
    };
    cancelBtn.addEventListener('click', close);
    overlay.addEventListener('click', e => { if (e.target === overlay) close(); });

    sendBtn.addEventListener('click', () => {
      sendBtn.disabled = true;
      cancelBtn.disabled = true;
      sendBtn.innerHTML = '⏳ Genereren…';
      statusMsg.textContent = '';

      try {
        // Gebruik het (evt. door gebruiker bijgewerkte) onderwerp
        const finalSubject = trim(fSubjInp?.value || subject) || subject;

        const xlsxB64  = rowsToXlsxBase64(rows);
        const emlBody  = buildBody(rows, finalSubject);
        const emlText  = buildEml(RECIPIENT, finalSubject, emlBody, xlsxB64, filename);
        const emlFile  = filename.replace('.xlsx', '.eml');

        downloadBlob(emlText, emlFile, 'message/rfc822');

        // Feedback
        sendBtn.innerHTML = '✓ Klaar';
        sendBtn.style.background = '#22c55e';
        statusMsg.style.color = '#4ade80';
        statusMsg.innerHTML =
          `<b>${emlFile}</b> gedownload.<br>` +
          `Open het bestand in Outlook en klik Verzenden.`;

        setTimeout(() => { close(); if (onSend) onSend(); }, 3500);

      } catch (err) {
        sendBtn.disabled = false; cancelBtn.disabled = false;
        sendBtn.innerHTML = '📨 Mail klaarzetten';
        statusMsg.style.color = '#f87171';
        statusMsg.textContent = `Fout: ${err.message}`;
      }
    });
  }

  /* ─── Publieke API ────────────────────────────────────────────────────── */

  window.ValExportMail = {
    /**
     * Open de compose-popup voor de Moederlijst-mail.
     *
     * @param {Object[]} rows       Gevalideerde + verrijkte itemlijst-rijen
     * @param {Function} [onSend]   Optionele callback na succesvol klaarzetten
     */
    open(rows, onSend) {
      if (!rows?.length) {
        alert('Geen rijen beschikbaar voor export.');
        return;
      }
      if (typeof XLSX === 'undefined') {
        alert('SheetJS (XLSX) is niet geladen. Controleer de scriptlading.');
        return;
      }
      showComposeModal(rows, onSend);
    },
  };

})();