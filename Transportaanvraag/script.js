/* Transportaanvraag Shipping – dynamisch formulier
 * Alle vragen, verplichte velden en afhankelijkheden staan declaratief in SECTIONS.
 * Wil je een vraag toevoegen of een regel aanpassen? Pas alleen SECTIONS / CONFIG aan.
 */
(function () {
  'use strict';

  /* =====================================================================
   * 1. CONFIG – drempelwaarden en mailadressen (hier aanpassen)
   * ===================================================================== */
  var CONFIG = {
    shippingMail: 'Shipping@royalihc.com',
    warehouseMail: 'DLPSWarehousing@royalihc.com',
    packagingMail: 'DLPSPackaging@royalihc.com',
    maxRows: 60,
    // Drempels voor zwaar / overmaat. LET OP: indicatief, stem af met Shipping.
    heavyKg: 1000,        // vanaf hier: heftruck / kraan benoemen
    veryHeavyKg: 5000,    // vanaf hier: zwaartepunt/hijspunten + tekening verplicht
    oversize: { l: 600, w: 240, h: 240 }, // cm; boven één van deze waarden = overmaat
    courierKg: 70,        // indicatie maximaal gewicht per colli koerier
    courierCm: 120        // indicatie maximale lengte per colli koerier
  };

  /* =====================================================================
   * 2. Vaste lijsten (afgeleid uit Shipping Transport Formulier.xlsx)
   * ===================================================================== */
  var NVT = 'Non Hazardous / NVT';
  var ENTITIES = ['IHC Smit BV', 'IHC Holland', 'IHC Beavers', 'IHC Mining', 'IHC Defense', 'IHC Merwede Holding'];
  var PKG_TYPES = ['Pallet', 'Box / Crate', 'Non-Pack', 'Package / Parcel', 'Document', '20 ft Container', '40 ft Container'];
  var DG_CLASSES = [
    'CLASS 1 – EXPLOSIVES', 'CLASS 2 – GASES', 'CLASS 3 – FLAMMABLE LIQUIDS', 'CLASS 4 – FLAMMABLE SOLIDS',
    'CLASS 5 – OXIDIZING SUBSTANCES; ORGANIC PEROXIDES', 'CLASS 6 – TOXIC SUBSTANCES; INFECTIOUS SUBSTANCES',
    'CLASS 7 – RADIOACTIVE MATERIAL', 'CLASS 8 – CORROSIVES', 'CLASS 9 – MISCELLANEOUS DANGEROUS GOODS'
  ];
  var SEA_ONLY = ['FAS', 'FOB', 'CFR', 'CIF'];
  var INCOTERMS = [
    { v: 'EXW', l: 'EXW – Ex Works (niet gebruiken, alternatief: FCA)', off: true },
    { v: 'FCA', l: 'FCA – Free Carrier (vrachtvrij tot vervoerder)' },
    { v: 'CPT', l: 'CPT – Carriage Paid To (vrachtvrij tot)' },
    { v: 'CIP', l: 'CIP – Carriage and Insurance Paid To' },
    { v: 'DAT', l: 'DAT – Delivered At Terminal (franco terminal)' },
    { v: 'DAP', l: 'DAP – Delivered At Place (franco ter plaatse)' },
    { v: 'DDP', l: 'DDP – Delivered Duty Paid (niet gebruiken, alternatief: DAP)', off: true },
    { v: 'FAS', l: 'FAS – Free Alongside Ship (alleen zeevracht)' },
    { v: 'FOB', l: 'FOB – Free On Board (alleen zeevracht)' },
    { v: 'CFR', l: 'CFR – Cost and Freight (alleen zeevracht)' },
    { v: 'CIF', l: 'CIF – Cost, Insurance and Freight (alleen zeevracht)' }
  ];
  var EU_OTHER = ['België', 'Bulgarije', 'Cyprus', 'Denemarken', 'Duitsland', 'Estland', 'Finland', 'Frankrijk', 'Griekenland',
    'Hongarije', 'Ierland', 'Italië', 'Kroatië', 'Letland', 'Litouwen', 'Luxemburg', 'Malta', 'Oostenrijk', 'Polen', 'Portugal',
    'Roemenië', 'Slovenië', 'Slowakije', 'Spanje', 'Tsjechië', 'Zweden'];
  var COUNTRY_OPTS = [
    { v: 'Nederland', l: 'Nederland' },
    { group: 'Europese Unie', items: EU_OTHER.map(function (c) { return { v: c, l: c }; }) },
    { group: 'Overig', items: [{ v: 'Verenigd Koninkrijk', l: 'Verenigd Koninkrijk' }, { v: 'OTHER', l: 'Ander land (buiten EU en UK)…' }] }
  ];
  var YN = ['Ja', 'Nee'];

  /* =====================================================================
   * 3. Hulpfuncties
   * ===================================================================== */
  function h(tag, attrs) {
    var e = document.createElement(tag);
    attrs = attrs || {};
    Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v === false || v == null) return;
      if (k === 'class') e.className = v;
      else if (k === 'text') e.textContent = v;
      else if (k === 'html') e.innerHTML = v;
      else e.setAttribute(k, v === true ? '' : v);
    });
    for (var i = 2; i < arguments.length; i++) {
      var kids = [].concat(arguments[i]);
      for (var j = 0; j < kids.length; j++) {
        var c = kids[j];
        if (c == null || c === false) continue;
        e.appendChild(c.nodeType ? c : document.createTextNode(String(c)));
      }
    }
    return e;
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function num(v) {
    var n = parseFloat(String(v == null ? '' : v).replace(',', '.'));
    return isNaN(n) ? 0 : n;
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function todayStr() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function nowLocal() { var d = new Date(); return todayStr() + 'T' + pad(d.getHours()) + ':' + pad(d.getMinutes()); }
  function fmtDate(s) { return /^\d{4}-\d{2}-\d{2}/.test(s || '') ? s.slice(8, 10) + '-' + s.slice(5, 7) + '-' + s.slice(0, 4) : (s || ''); }
  function fmtDT(s) { return /T/.test(s || '') ? fmtDate(s.slice(0, 10)) + ' ' + s.slice(11, 16) : (s || ''); }
  function r2(n) { return (Math.round(n * 100) / 100).toLocaleString('nl-NL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function r3(n) { return (Math.round(n * 1000) / 1000).toLocaleString('nl-NL', { minimumFractionDigits: 3, maximumFractionDigits: 3 }); }
  function isEmpty(v) {
    if (v == null) return true;
    if (Array.isArray(v)) return v.length === 0;
    if (typeof v === 'boolean') return v === false;
    return String(v).trim() === '';
  }
  var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
  var PHONE_RE = /^[+()\d\s./-]{6,}$/;
  function validEmail(v) { return EMAIL_RE.test(String(v).trim()) ? '' : 'Ongeldig e-mailadres.'; }
  function validPhone(v) { return PHONE_RE.test(String(v).trim()) ? '' : 'Ongeldig telefoonnummer (gebruik cijfers, bij voorkeur met landcode).'; }
  function notEU(v) {
    return /^\s*(eu|e\.u\.|europa|europe|europese unie|european union)\s*$/i.test(v) ? 'Land van oorsprong moet een land zijn, geen "EU".' : '';
  }

  function zoneOf(c) {
    if (c === 'Nederland') return 'NL';
    if (c === 'Hongarije') return 'HU';
    if (c === 'Roemenië') return 'RO';
    if (c === 'Verenigd Koninkrijk') return 'UK';
    if (c === 'OTHER') return 'NONEU';
    if (EU_OTHER.indexOf(c) >= 0) return 'EU';
    return '';
  }
  function cname(S, p) {
    var v = S[p];
    if (v === 'OTHER') return String(S[p + 'Other'] || '').trim();
    return v || '';
  }

  /* =====================================================================
   * 4. Afgeleide route/zendingskenmerken (de "schifting")
   * ===================================================================== */
  function derive(S) {
    var X = {};
    X.oz = zoneOf(S.origin);
    X.dz = zoneOf(S.dest);
    X.oName = cname(S, 'origin');
    X.dName = cname(S, 'dest');
    X.routeKnown = !!(X.oz && X.dz && X.oName && X.dName);
    X.domestic = X.routeKnown && X.oName.toLowerCase() === X.dName.toLowerCase();
    X.cross = X.routeKnown && !X.domestic;
    var nonEU = function (z) { return z === 'UK' || z === 'NONEU'; };
    X.customs = X.cross && (nonEU(X.oz) || nonEU(X.dz));
    if (!X.cross) X.direction = X.routeKnown ? 'Binnenlands' : '';
    else if (!X.customs) X.direction = 'Intra-EU';
    else if (nonEU(X.dz) && !nonEU(X.oz)) X.direction = 'Export';
    else if (nonEU(X.oz) && !nonEU(X.dz)) X.direction = 'Import';
    else X.direction = 'Transit';
    X.customsSuggest = { 'Intra-EU': 'EU-goederen', Export: 'Export', Import: 'Import', Transit: 'Transit' }[X.direction] || '';
    X.isUK = X.dz === 'UK';
    X.isROHU = X.dz === 'RO' || X.dz === 'HU';
    X.isDG = S.dg === 'Ja';
    X.urgent = S.urgent === 'Ja';
    X.maxL = num(S.maxL); X.maxW = num(S.maxW); X.maxH = num(S.maxH);
    X.maxKg = num(S.maxKg); X.totalKg = num(S.totalKg);
    X.veryHeavy = X.maxKg >= CONFIG.veryHeavyKg;
    X.heavy = X.maxKg >= CONFIG.heavyKg;
    X.oversize = X.maxL > CONFIG.oversize.l || X.maxW > CONFIG.oversize.w || X.maxH > CONFIG.oversize.h;
    X.liftReq = X.veryHeavy || X.oversize;
    X.mode = S.mode || '';
    X.defence = S.defence === 'Ja';
    X.partsTotal = (S.parts || []).reduce(function (a, r) { return a + num(r.qty) * num(r.val); }, 0);
    return X;
  }

  /* =====================================================================
   * 5. Tabelblokken (colli-specificatie en UK-partlijst)
   * ===================================================================== */
  var colliModel = [];
  var partsModel = [];

  function newColli() { return { qty: '1', type: '', desc: '', l: '', w: '', h: '', net: '', gross: '', stack: '', haz: NVT }; }
  function newPart() { return { pn: '', desc: '', qty: '', coo: '', val: '', hs: '' }; }

  function makeTable(cfg) {
    var model = cfg.model;
    var wrap = h('div', { class: 'tblblock' });
    var table = h('table', { class: 'dt' });
    var thead = h('thead'), tbody = h('tbody'), tfoot = h('tfoot');
    var hr = h('tr');
    hr.appendChild(h('th', { text: '#' }));
    cfg.cols.forEach(function (c) { hr.appendChild(h('th', { text: c.l })); });
    hr.appendChild(h('th', { text: '' }));
    thead.appendChild(hr);
    table.appendChild(thead); table.appendChild(tbody); table.appendChild(tfoot);
    var problemsEl = h('ul', { class: 'tbl-problems' });
    var tools = h('div', { class: 'tbl-tools' });
    var addBtn = h('button', { type: 'button', class: 'btn btn-sm', text: '+ ' + cfg.addLabel });
    tools.appendChild(addBtn);
    var pasteBox = null;
    if (cfg.paste) {
      var pasteBtn = h('button', { type: 'button', class: 'btn btn-sm', text: 'Plak uit Excel' });
      var ta = h('textarea', { placeholder: cfg.paste.hint, 'aria-label': 'Plak rijen uit Excel' });
      var go = h('button', { type: 'button', class: 'btn btn-sm btn-primary', text: 'Verwerk geplakte regels' });
      pasteBox = h('div', { class: 'paste-box', hidden: true }, ta, h('div', { class: 'tbl-tools' }, go));
      pasteBtn.addEventListener('click', function () { pasteBox.hidden = !pasteBox.hidden; });
      go.addEventListener('click', function () {
        var rows = cfg.paste.parse(ta.value);
        if (!rows.length) { alert('Geen regels herkend. Plak kolommen gescheiden door tabs (rechtstreeks uit Excel).'); return; }
        if (model.length === 1 && cfg.isBlank(model[0])) model.length = 0;
        rows.forEach(function (r) { if (model.length < CONFIG.maxRows) model.push(r); });
        ta.value = ''; pasteBox.hidden = true;
        render(); cfg.onChange();
      });
      tools.appendChild(pasteBtn);
    }
    wrap.appendChild(h('div', { class: 'tbl-wrap' }, table));
    wrap.appendChild(tools);
    if (pasteBox) wrap.appendChild(pasteBox);
    wrap.appendChild(problemsEl);

    addBtn.addEventListener('click', function () {
      if (model.length >= CONFIG.maxRows) return;
      model.push(cfg.newRow());
      render(); cfg.onChange();
    });

    function render() {
      tbody.innerHTML = '';
      model.forEach(function (row, r) {
        var tr = h('tr');
        tr.appendChild(h('td', { class: 'nr', text: String(r + 1) }));
        cfg.cols.forEach(function (c) {
          var td = h('td', { class: 'c-' + c.k });
          if (c.t === 'calc') {
            td.className += ' calc';
            td.setAttribute('data-calc', c.k);
            td.textContent = c.calc(row);
          } else if (c.t === 'select') {
            var sel = h('select', { 'data-r': r, 'data-k': c.k, 'aria-label': c.l + ' regel ' + (r + 1) });
            sel.appendChild(h('option', { value: '', text: '— kies —' }));
            c.opts.forEach(function (o) { sel.appendChild(h('option', { value: o, text: o })); });
            sel.value = row[c.k];
            td.appendChild(sel);
          } else {
            var inp = h('input', {
              type: c.t === 'number' ? 'number' : 'text', 'data-r': r, 'data-k': c.k,
              'aria-label': c.l + ' regel ' + (r + 1), min: c.t === 'number' ? '0' : false, step: c.t === 'number' ? 'any' : false
            });
            inp.value = row[c.k];
            td.appendChild(inp);
          }
          tr.appendChild(td);
        });
        var del = h('button', { type: 'button', class: 'row-del', title: 'Regel verwijderen', 'aria-label': 'Verwijder regel ' + (r + 1), text: '×' });
        del.addEventListener('click', function () {
          model.splice(r, 1);
          if (!model.length) model.push(cfg.newRow());
          render(); cfg.onChange();
        });
        tr.appendChild(h('td', {}, del));
        tbody.appendChild(tr);
      });
      renderFoot();
    }

    function renderFoot() {
      tfoot.innerHTML = '';
      var f = cfg.foot(model);
      var tr = h('tr');
      tr.appendChild(h('td', { text: 'Totaal' }));
      cfg.cols.forEach(function (c) { tr.appendChild(h('td', { 'data-foot': c.k, text: f[c.k] == null ? '' : f[c.k], style: f[c.k] != null ? 'text-align:right' : false })); });
      tr.appendChild(h('td'));
      tfoot.appendChild(tr);
    }

    function refreshCalcs() {
      model.forEach(function (row, r) {
        var tr = tbody.children[r];
        if (!tr) return;
        cfg.cols.forEach(function (c) {
          if (c.t !== 'calc') return;
          var td = tr.querySelector('[data-calc="' + c.k + '"]');
          if (td) td.textContent = c.calc(row);
        });
      });
      renderFoot();
    }

    wrap.addEventListener('input', function (e) {
      var t = e.target;
      if (!t.hasAttribute('data-k')) return;
      model[+t.getAttribute('data-r')][t.getAttribute('data-k')] = t.value;
      refreshCalcs();
    });
    wrap.addEventListener('change', function (e) {
      var t = e.target;
      if (!t.hasAttribute('data-k')) return;
      model[+t.getAttribute('data-r')][t.getAttribute('data-k')] = t.value;
      refreshCalcs();
    });
    wrap.addEventListener('focusout', function (e) {
      if (e.target.hasAttribute && e.target.hasAttribute('data-k')) e.target.setAttribute('data-t', '1');
    });

    /* Markeer ongeldige cellen en geef lijst met problemen terug */
    function mark(S, X, show) {
      var out = [];
      if (cfg.sync && cfg.sync(S, X)) render();
      model.forEach(function (row, r) {
        var bad = [];
        cfg.cols.forEach(function (c) {
          if (c.t === 'calc') return;
          var ctrl = tbody.querySelector('[data-r="' + r + '"][data-k="' + c.k + '"]');
          if (!ctrl) return;
          var st = cfg.cellState ? cfg.cellState(row, c.k, S, X) : {};
          ctrl.disabled = !!st.disabled;
          var invalid = !st.disabled && cfg.cellErr(row, c.k, S, X);
          ctrl.classList.toggle('miss', !!invalid && (show || ctrl.getAttribute('data-t') === '1'));
          if (invalid) bad.push(c.l);
        });
        if (bad.length) out.push('Regel ' + (r + 1) + ': ' + bad.join(', '));
      });
      var gen = cfg.general ? cfg.general(S, X) : [];
      out = gen.concat(out);
      problemsEl.innerHTML = '';
      if (show) out.forEach(function (m) { problemsEl.appendChild(h('li', { text: m })); });
      return out;
    }

    render();
    return { el: wrap, render: render, mark: mark, refreshCalcs: refreshCalcs };
  }

  var colliTable, partsTable;

  function buildColliTable(onChange) {
    var cols = [
      { k: 'qty', l: 'Aantal', t: 'number' },
      { k: 'type', l: 'Verpakkingstype', t: 'select', opts: PKG_TYPES },
      { k: 'desc', l: 'Omschrijving / markering', t: 'text' },
      { k: 'l', l: 'L (cm)', t: 'number' },
      { k: 'w', l: 'B (cm)', t: 'number' },
      { k: 'h', l: 'H (cm)', t: 'number' },
      { k: 'cube', l: 'Kubus (m³)', t: 'calc', calc: function (r) { return r3(num(r.l) * num(r.w) * num(r.h) / 1e6); } },
      { k: 'net', l: 'Netto kg / colli', t: 'number' },
      { k: 'gross', l: 'Bruto kg / colli', t: 'number' },
      { k: 'stack', l: 'Stapelbaar', t: 'select', opts: YN },
      { k: 'haz', l: 'Gevaarlijk', t: 'select', opts: [NVT].concat(DG_CLASSES) }
    ];
    var lastDG = null;
    return makeTable({
      model: colliModel, cols: cols, addLabel: 'Colli-regel toevoegen', onChange: onChange,
      newRow: newColli,
      isBlank: function (r) { return !r.type && !r.desc && !r.l && !r.w && !r.h && !r.net && !r.gross; },
      cellErr: function (r, k, S) {
        var n = num(r[k]);
        switch (k) {
          case 'qty': return !(n >= 1 && Math.floor(n) === n);
          case 'type': case 'desc': case 'stack': case 'haz': return !String(r[k]).trim();
          case 'l': case 'w': case 'h': case 'gross': return !(n > 0);
          case 'net': return !(n > 0) || (num(r.gross) > 0 && n > num(r.gross));
        }
        return false;
      },
      cellState: function (r, k, S, X) { return k === 'haz' ? { disabled: !X.isDG } : {}; },
      sync: function (S, X) {
        var changed = false;
        if (!X.isDG) {
          colliModel.forEach(function (r) { if (r.haz !== NVT) { r.haz = NVT; changed = true; } });
        } else if (lastDG === false) {
          colliModel.forEach(function (r) { if (r.haz === NVT) { r.haz = ''; changed = true; } });
        }
        lastDG = X.isDG;
        return changed;
      },
      general: function (S, X) {
        var g = [];
        if (X.isDG && !colliModel.some(function (r) { return r.haz && r.haz !== NVT; })) g.push('Gevaarlijke goederen = Ja: kies bij minimaal één colli de gevarenklasse.');
        return g;
      },
      foot: function (m) {
        var q = 0, cube = 0, net = 0, gross = 0;
        m.forEach(function (r) {
          var n = Math.max(0, Math.floor(num(r.qty)));
          q += n; cube += n * num(r.l) * num(r.w) * num(r.h) / 1e6; net += n * num(r.net); gross += n * num(r.gross);
        });
        return { qty: String(q), cube: r3(cube), net: r2(net), gross: r2(gross) };
      }
    });
  }

  function buildPartsTable(onChange) {
    var cols = [
      { k: 'pn', l: 'Part number', t: 'text' },
      { k: 'desc', l: 'Parts description', t: 'text' },
      { k: 'qty', l: 'Booked quantity', t: 'number' },
      { k: 'coo', l: 'Country of origin', t: 'text' },
      { k: 'val', l: 'Item value (€)', t: 'number' },
      { k: 'total', l: 'Total value (€)', t: 'calc', calc: function (r) { return r2(num(r.qty) * num(r.val)); } },
      { k: 'hs', l: 'HS code (8 cijfers)', t: 'text' }
    ];
    return makeTable({
      model: partsModel, cols: cols, addLabel: 'Part toevoegen', onChange: onChange,
      newRow: newPart,
      isBlank: function (r) { return !r.pn && !r.desc && !r.qty && !r.coo && !r.val && !r.hs; },
      cellErr: function (r, k) {
        var n = num(r[k]), v = String(r[k]).trim();
        switch (k) {
          case 'pn': case 'desc': return !v;
          case 'qty': return !(n >= 1 && Math.floor(n) === n);
          case 'coo': return !v || !!notEU(v);
          case 'val': return !(n > 0);
          case 'hs': return !/^\d{8}$/.test(v.replace(/[\s.]/g, ''));
        }
        return false;
      },
      foot: function (m) {
        var t = m.reduce(function (a, r) { return a + num(r.qty) * num(r.val); }, 0);
        return { total: r2(t) };
      },
      paste: {
        hint: 'Plak hier regels uit Excel. Kolomvolgorde: Part number | Description | Quantity | Country of origin | Item value | (Total value) | HS code',
        parse: function (txt) {
          return txt.split(/\r?\n/).map(function (line) { return line.split('\t'); })
            .filter(function (c) { return c.length >= 6 && c.join('').trim(); })
            .map(function (c) {
              c = c.map(function (x) { return x.trim(); });
              var hs = c.length >= 7 ? c[6] : c[5];
              return { pn: c[0], desc: c[1], qty: c[2], coo: c[3], val: c[4].replace(',', '.'), hs: hs.replace(/[\s.]/g, '') };
            }).filter(function (r) { return !/^part\s*number$/i.test(r.pn); });
        }
      },
      general: function () { return partsModel.length ? [] : ['Voeg minimaal één part toe.']; }
    });
  }

  /* =====================================================================
   * 6. SECTIES EN VRAGEN
   *   req / vis / validate krijgen (S = ingevulde waarden, X = afgeleide route)
   * ===================================================================== */
  var SECTIONS = [
    {
      id: 'main', main: true, title: 'Hoofdvragen',
      intro: 'Op basis van deze antwoorden bepaalt het formulier welke vervolgvragen en documenten nodig zijn.',
      fields: [
        { id: 'origin', label: 'Land van verzending', type: 'select', opts: COUNTRY_OPTS, req: true, def: 'Nederland', hint: 'Waar worden de goederen opgehaald?' },
        { id: 'originOther', label: 'Land van verzending (naam)', type: 'text', req: true, vis: function (S) { return S.origin === 'OTHER'; } },
        { id: 'dest', label: 'Bestemming (land van aflevering)', type: 'select', opts: COUNTRY_OPTS, req: true, hint: 'Bepaalt o.a. Incoterm, douane en UK/RO/HU-regels.' },
        { id: 'destOther', label: 'Bestemming (naam land)', type: 'text', req: true, vis: function (S) { return S.dest === 'OTHER'; } },
        { id: 'maxL', label: 'Grootste colli – lengte (cm)', type: 'number', req: true },
        { id: 'maxW', label: 'Grootste colli – breedte (cm)', type: 'number', req: true },
        { id: 'maxH', label: 'Grootste colli – hoogte (cm)', type: 'number', req: true },
        { id: 'maxKg', label: 'Zwaarste colli – bruto gewicht (kg)', type: 'number', req: true },
        {
          id: 'totalKg', label: 'Totaal bruto gewicht zending (kg)', type: 'number', req: true,
          validate: function (v, S) { return num(v) < num(S.maxKg) ? 'Totaalgewicht kan niet lager zijn dan het zwaarste colli.' : ''; }
        },
        { id: 'dg', label: 'Gevaarlijke goederen', type: 'radio', opts: YN, req: true },
        {
          id: 'pickDate', label: 'Ophaaldatum (datum dat de goederen gereed staan)', type: 'date', req: true,
          validate: function (v) { return v < todayStr() ? 'Ophaaldatum ligt in het verleden.' : ''; }
        },
        { id: 'urgent', label: 'Spoed', type: 'radio', opts: YN, req: true }
      ]
    },
    {
      id: 'plan', title: 'Aanvraag en planning',
      fields: [
        { id: 'requestNo', label: 'Aanvraagnummer', type: 'text', hint: 'Unieke referentie, indien intern beschikbaar.' },
        { id: 'reqDate', label: 'Datum aanvraag', type: 'date', req: true, def: todayStr },
        { id: 'readyDate', label: 'Gewenste ophaaldatum (goederen gereed)', type: 'computed', calc: function (S) { return fmtDate(S.pickDate); } },
        { id: 'pickWindow', label: 'Tijdvenster ophalen', type: 'text', ph: 'bijv. 08:00 - 12:00' },
        {
          id: 'deliverDate', label: 'Gewenste leverdatum / deadline', type: 'date',
          req: function (S, X) { return X.urgent; }, hint: 'Harde deadline of gewenste datum.',
          validate: function (v, S) { return S.pickDate && v < S.pickDate ? 'Leverdatum ligt vóór de ophaaldatum.' : ''; }
        },
        { id: 'priority', label: 'Prioriteit', type: 'select', opts: ['Spoed', 'Kritiek'], req: true, vis: function (S, X) { return X.urgent; } },
        { id: 'urgentReason', label: 'Reden spoed / deadline', type: 'textarea', wide: true, req: true, vis: function (S, X) { return X.urgent; } }
      ]
    },
    {
      id: 'req', title: 'Aanvrager en kosten',
      fields: [
        { id: 'seller', label: 'Seller – IHC entity', type: 'select', opts: ENTITIES, req: true },
        { id: 'name', label: 'Naam aanvrager', type: 'text', req: true },
        { id: 'dept', label: 'Afdeling', type: 'text', req: true },
        { id: 'email', label: 'E-mail', type: 'email', req: true, validate: validEmail },
        { id: 'phone', label: 'Telefoonnummer', type: 'tel', req: true, validate: validPhone },
        { id: 'project', label: 'Project / bouwnummer', type: 'text', req: true, hint: 'Bijvoorbeeld YN-nummer.' },
        { id: 'po', label: 'PO# / referentie', type: 'text', req: true, hint: 'PO, delivery reference of andere herkenbare dossiercode.', ph: 'bijv. 6226000277' },
        { id: 'costRef', label: 'Kostenplaats / WBS / boekingsreferentie', type: 'text', req: true, hint: 'Vereist voor boeking van transportkosten.' },
        { id: 'invEntity', label: 'Invoicing entity / billing department', type: 'select', opts: ENTITIES, req: true }
      ]
    },
    {
      id: 'pickup', title: 'Ophaallocatie',
      fields: [
        { id: 'pkCompany', label: 'Bedrijfsnaam laadlocatie', type: 'text', req: true },
        { id: 'pkStreet', label: 'Straat en huisnummer', type: 'text', req: true },
        { id: 'pkPostcode', label: 'Postcode', type: 'text', req: true },
        { id: 'pkCity', label: 'Plaats', type: 'text', req: true },
        { id: 'pkCountry', label: 'Land', type: 'computed', calc: function (S, X) { return X.oName; } },
        { id: 'pkRef', label: 'Laadreferentie', type: 'text' },
        { id: 'pkContact', label: 'Contactpersoon laadlocatie', type: 'text', req: true, hint: 'Naam en functie indien bekend.' },
        { id: 'pkPhone', label: 'Telefoon laadlocatie', type: 'tel', req: true, validate: validPhone, hint: 'Inclusief landcode.' },
        { id: 'pkEmail', label: 'E-mail laadlocatie', type: 'email', req: true, validate: validEmail },
        { id: 'pkHours', label: 'Openingstijden', type: 'text', req: true, ph: '08:00 - 16:00 (lunch 12:00 - 13:00)', hint: 'Vermeld pauzes en lokale tijd.' },
        { id: 'pkReadyTime', label: 'Goederen gereed vanaf (tijd)', type: 'time' },
        {
          id: 'pkReady', label: 'Goederen gepickt en verpakt, documenten gereed?', type: 'radio', opts: YN, req: true, wide: true,
          validate: function (v) { return v === 'Nee' ? 'Laat eerst picken en verpakken; stuur het formulier daarna naar Shipping.' : ''; }
        },
        {
          id: 'pkNotice', type: 'info', tone: 'warn', wide: true, vis: function (S) { return S.pkReady === 'Nee'; },
          text: function (S, X) {
            var subj = 'Picken/verpakken: ' + (S.project || '') + ' ' + (S.po || '');
            var body = 'Graag picken en verpakken.\nProject: ' + (S.project || '') + '\nPO: ' + (S.po || '') +
              '\nGewenste ophaaldatum: ' + fmtDate(S.pickDate) + '\nBestemming: ' + (X.dName || '') +
              '\nGraag gewichten en afmetingen doorsturen indien nog niet bekend.';
            var href = 'mailto:' + CONFIG.warehouseMail + '?cc=' + encodeURIComponent(CONFIG.packagingMail) +
              '&subject=' + encodeURIComponent(subj) + '&body=' + encodeURIComponent(body);
            return '<b>Nog niet gepickt of verpakt?</b> Shipping kan geen order of shipment opstarten en stuurt geen Warehousing aan. ' +
              'Mail <a href="' + href + '">DLPSWarehousing</a> (met DLPSPackaging in cc) zodat de goederen gepickt en verpakt worden en gewichten en afmetingen worden doorgegeven.';
          }
        },
        { id: 'pkInstr', label: 'Laadinstructies / beperkingen', type: 'textarea', wide: true, hint: 'Bijvoorbeeld heftruck, kraan, zijlading, dock, PBM, toegang of aanmelden.' }
      ]
    },
    {
      id: 'delivery', title: 'Afleverlocatie',
      fields: [
        { id: 'custName', label: 'Klantnaam (customer name)', type: 'text', req: true },
        { id: 'dlCompany', label: 'Bedrijfsnaam loslocatie', type: 'text', req: true },
        { id: 'dlStreet', label: 'Straat en huisnummer', type: 'text', req: true },
        { id: 'dlPostcode', label: 'Postcode', type: 'text', req: true },
        { id: 'dlCity', label: 'Plaats', type: 'text', req: true },
        { id: 'dlCountry', label: 'Land', type: 'computed', calc: function (S, X) { return X.dName; } },
        { id: 'dlRef', label: 'Losreferentie', type: 'text' },
        { id: 'dlContact', label: 'Contactpersoon loslocatie (klant)', type: 'text', req: true, hint: 'Naam en functie indien bekend.' },
        { id: 'dlPhone', label: 'Telefoon loslocatie', type: 'tel', req: true, validate: validPhone, hint: 'Inclusief landcode.' },
        { id: 'dlEmail', label: 'E-mail loslocatie', type: 'email', req: true, validate: validEmail },
        { id: 'dlHours', label: 'Openingstijden / afleveruren', type: 'text', req: true, ph: '08:00 - 17:00', hint: 'Vermeld pauzes en lokale tijd.' },
        { id: 'dlInstr', label: 'Losinstructies / beperkingen', type: 'textarea', wide: true, hint: 'Bijvoorbeeld heftruck, kraan, dock, PBM, toegang of voormelding.' },
        { id: 'defence', label: 'Defence-zending of beveiligde locatie?', type: 'radio', opts: YN, req: true, wide: true }
      ]
    },
    {
      id: 'goods', title: 'Goederen en colli',
      fields: [
        { id: 'goodsDesc', label: 'Omschrijving goederen', type: 'text', wide: true, req: true, hint: 'Duidelijke commerciële omschrijving, niet alleen interne artikelnummers.' },
        { id: 'colli', label: 'Colli-specificatie', type: 'colli', wide: true },
        {
          id: 'liftPoints', label: 'Zwaartepunt / hijspunten', type: 'textarea', wide: true, req: function (S, X) { return X.liftReq; },
          hint: 'Verplicht bij zware, hoge of afwijkende goederen.'
        },
        { id: 'weather', label: 'Temperatuur- of weersgevoelig', type: 'radio', opts: YN, req: true },
        { id: 'weatherNote', label: 'Toelichting beschermingsbehoefte', type: 'text', req: true, vis: function (S) { return S.weather === 'Ja'; } },
        {
          id: 'goodsValue', label: 'Waarde goederen', type: 'number', req: function (S, X) { return X.cross; }, vis: function (S, X) { return X.cross; },
          hint: 'Relevant voor verzekering en douane.'
        },
        { id: 'currency', label: 'Valuta', type: 'select', opts: ['EUR', 'USD', 'GBP', 'Anders'], req: function (S, X) { return X.cross; }, vis: function (S, X) { return X.cross; }, def: 'EUR' },
        {
          id: 'valueNotice', type: 'info', tone: 'info', wide: true, vis: function (S, X) { return X.isROHU; },
          text: function (S, X) {
            return 'Voor ' + (X.dz === 'RO' ? 'Roemenië heeft Shipping de totale goederenwaarde nodig voor de <b>UIT-code</b>.' : 'Hongarije heeft Shipping de totale goederenwaarde nodig voor de <b>EKAER-code</b>.');
          }
        }
      ]
    },
    {
      id: 'transport', title: 'Transportvereisten',
      fields: [
        {
          id: 'mode', label: 'Mode of transport', type: 'select', req: true,
          opts: [
            { v: 'Road Transport', l: 'Road Transport (wegvervoer)' },
            { v: 'Courier', l: 'Courier (koerier)' },
            { v: 'Sea freight', l: 'Sea freight (zeevracht)', dis: function (S, X) { return !X.cross; } },
            { v: 'Air freight', l: 'Air freight (luchtvracht)', dis: function (S, X) { return !X.cross; } }
          ]
        },
        {
          id: 'vehicle', label: 'Voorkeur transportmiddel', type: 'select', req: true, vis: function (S) { return S.mode === 'Road Transport'; },
          opts: ['Gesloten wagen', 'Open wagen', 'Tautliner', 'Bestelbus', 'Dieplader / semi-dieplader', 'Anders']
        },
        {
          id: 'openClosed', label: 'Noodzaak open of gesloten wagen', type: 'text', wide: true, req: true, vis: function (S) { return S.mode === 'Road Transport'; },
          hint: 'Onderbouw op basis van weersbescherming en laad-/losmethode.'
        },
        {
          id: 'incoterm', label: 'Incoterm 2020', type: 'select', req: true, vis: function (S, X) { return X.cross; },
          opts: INCOTERMS
        },
        { id: 'incotermPlace', label: 'Incoterm – plaats (verplicht)', type: 'text', req: true, vis: function (S, X) { return X.cross; }, ph: 'bijv. Zwijndrecht (DAP)' },
        {
          id: 'equipment', label: 'Speciale uitrusting', type: 'checks', wide: true,
          opts: ['Laadklep', 'Kraan', 'Kooiaap', 'Spanmateriaal', 'Dekzeil', 'ADR-uitrusting', 'Anders (zie instructies)']
        },
        { id: 'specialInstr', label: 'Special instructions', type: 'textarea', wide: true, ph: 'bijv. Dry transport, crane needed, oversized transport' }
      ]
    },
    {
      id: 'dg', title: 'Gevaarlijke goederen', vis: function (S, X) { return X.isDG; },
      fields: [
        { id: 'dgUN', label: 'UN-nummer', type: 'text', req: true, ph: 'UN1234', validate: function (v) { return /^(UN)?\s?\d{4}$/i.test(v.trim()) ? '' : 'UN-nummer bestaat uit 4 cijfers (bijv. UN1993).'; } },
        { id: 'dgClass', label: 'Klasse', type: 'select', opts: DG_CLASSES, req: true },
        { id: 'dgPG', label: 'Verpakkingsgroep', type: 'select', opts: ['I', 'II', 'III', 'N.v.t.'], req: true },
        { id: 'dgName', label: 'Officiële vervoersbenaming (proper shipping name)', type: 'text', req: true },
        {
          id: 'dgRegime', label: 'Regelgeving', type: 'computed', hint: 'Afgeleid van de transportwijze.',
          calc: function (S) { return { 'Road Transport': 'ADR', Courier: 'ADR (controleer of koerier ADR accepteert)', 'Sea freight': 'IMDG', 'Air freight': 'IATA' }[S.mode] || 'Kies eerst de transportwijze'; }
        },
        {
          id: 'dgNotice', type: 'info', tone: 'warn', wide: true,
          text: function () { return 'Bij gevaarlijke goederen zijn de ADR/IMDG/IATA-classificatie en het SDS verplicht mee te sturen (zie bijlagencontrole).'; }
        }
      ]
    },
    {
      id: 'customs', title: 'Douane en compliance', vis: function (S, X) { return X.cross; },
      fields: [
        {
          id: 'customsStatus', label: 'Douanestatus', type: 'select', req: true,
          opts: ['EU-goederen', 'Export', 'Import', 'Transit', 'Onbekend'], hint: 'Automatisch voorgesteld op basis van de route; pas aan indien nodig.'
        },
        {
          id: 'hsCodes', label: 'HS-code(s)', type: 'textarea', wide: true, req: function (S, X) { return X.customs; }, vis: function (S, X) { return !X.isUK; },
          hint: 'Per artikel. Verplicht bij zendingen van/naar landen buiten de EU.'
        },
        {
          id: 'coo', label: 'Land van oorsprong', type: 'text', req: function (S, X) { return X.customs; }, vis: function (S, X) { return !X.isUK; },
          validate: notEU, hint: 'Per artikel; een land, geen "EU".'
        },
        { id: 'licence', label: 'Uitvoer- of vergunningseisen', type: 'textarea', wide: true, vis: function (S, X) { return X.customs; }, hint: 'Dual-use, sancties, exportvergunning of andere restricties.' },
        {
          id: 'euNotice', type: 'info', tone: 'info', wide: true, vis: function (S, X) { return X.direction === 'Intra-EU'; },
          text: function () { return 'Zending binnen de EU: handelsfactuur en paklijst zijn nodig; HS-codes en land van oorsprong zijn optioneel.'; }
        }
      ]
    },
    {
      id: 'uk', title: 'Aanvullend voor Verenigd Koninkrijk', vis: function (S, X) { return X.isUK; },
      fields: [
        {
          id: 'ukNotice', type: 'info', tone: 'warn', wide: true,
          text: function () { return 'Voor het VK zijn alle partnummers nodig inclusief waarde, 8-cijferige HS-code en land van oorsprong (een land, geen "EU"). Het netto gewicht moet op de paklijst/pickinglist staan.'; }
        },
        { id: 'parts', label: 'Part details', type: 'parts', wide: true },
        { id: 'ukEmail', label: 'E-mail klant voor voormelding', type: 'email', validate: validEmail, hint: 'Shipping meldt de levering aan; klant verstrekt agent voor pre-entry.' },
        { id: 'ukAgent', label: 'Naam / gegevens douaneagent', type: 'text', hint: 'Indien al bekend.' }
      ]
    },
    {
      id: 'defenceSec', title: 'Defence en beveiligde locaties', vis: function (S) { return S.defence === 'Ja'; },
      fields: [
        { id: 'vog', label: 'VOG chauffeur vereist', type: 'radio', opts: ['Ja', 'Nee', 'Onbekend'], req: true },
        { id: 'notifyDriver', label: 'Chauffeur vooraf aanmelden', type: 'radio', opts: YN, req: true },
        { id: 'notifyTerm', label: 'Aanmeldtermijn en vereiste gegevens', type: 'text', wide: true, req: true, vis: function (S) { return S.notifyDriver === 'Ja'; } },
        { id: 'securityInstr', label: 'Toegangs- en veiligheidsinstructies', type: 'textarea', wide: true, hint: 'Identiteitsbewijs, voertuiggegevens, PBM, screening, escort, enz.' }
      ]
    },
    {
      id: 'attach', title: 'Bijlagencontrole',
      fields: [
        { id: 'attPack', label: 'Paklijst / pickinglist', type: 'select', req: true, opts: [{ v: 'Bijgevoegd' }, { v: 'Volgt' }, { v: 'N.v.t.', dis: function () { return true; } }] },
        {
          id: 'attInv', label: 'Commerciële factuur / pro forma', type: 'select', req: function (S, X) { return X.cross; },
          opts: [{ v: 'Bijgevoegd' }, { v: 'Volgt' }, { v: 'N.v.t.', dis: function (S, X) { return X.cross; } }]
        },
        { id: 'attPhotos', label: 'Foto’s van goederen en verpakking', type: 'select', opts: [{ v: 'Bijgevoegd' }, { v: 'Volgt' }, { v: 'N.v.t.' }] },
        {
          id: 'attSds', label: 'SDS / DG-documentatie', type: 'select', req: true, vis: function (S, X) { return X.isDG; },
          opts: [{ v: 'Bijgevoegd' }, { v: 'Volgt', dis: function () { return true; } }, { v: 'N.v.t.', dis: function () { return true; } }]
        },
        {
          id: 'attCustoms', label: 'Douane- of exportdocumenten', type: 'select', req: true, vis: function (S, X) { return X.customs; },
          opts: [{ v: 'Bijgevoegd' }, { v: 'Volgt' }, { v: 'N.v.t.', dis: function () { return true; } }]
        },
        {
          id: 'attDraw', label: 'Tekeningen / hijsplan / COG', type: 'select', req: function (S, X) { return X.liftReq; },
          opts: [{ v: 'Bijgevoegd' }, { v: 'N.v.t.', dis: function (S, X) { return X.liftReq; } }]
        },
        { id: 'attOther', label: 'Overige bijlagen', type: 'text', wide: true, hint: 'Bestandsnamen en korte toelichting.' }
      ]
    },
    {
      id: 'handover', title: 'Overdracht aan Shipping',
      fields: [
        {
          id: 'confirm', type: 'check', wide: true, req: true, short: 'Controle door aanvrager bevestigen',
          label: 'Controle door aanvrager: adressen en contactgegevens zijn compleet, de goederen zijn daadwerkelijk gereed, aantal/maat/gewicht per colli zijn bekend, kostenreferentie en factuurentiteit zijn ingevuld, bijzondere eisen zijn benoemd en de relevante documenten zijn bijgevoegd.'
        },
        { id: 'completedBy', label: 'Aanvraag compleet verklaard door', type: 'text', req: true, hint: 'Naam aanvrager.' },
        { id: 'handoverAt', label: 'Datum / tijd overdracht aan Shipping', type: 'datetime-local', req: true, def: nowLocal },
        { id: 'remarks', label: 'Aanvullende opmerkingen', type: 'textarea', wide: true, hint: 'Uitzonderingen, afhankelijkheden en openstaande acties.' }
      ]
    }
  ];

  /* =====================================================================
   * 7. Waarschuwingen en "extra vereist" (niet blokkerend)
   * ===================================================================== */
  function warnings(S, X) {
    var w = [];
    if (X.mode === 'Courier') {
      if (X.maxKg > CONFIG.courierKg) w.push('Zwaarste colli (' + X.maxKg + ' kg) is zwaar voor een koerier (indicatie max. ' + CONFIG.courierKg + ' kg). Overweeg Road Transport.');
      if (Math.max(X.maxL, X.maxW, X.maxH) > CONFIG.courierCm) w.push('Grootste afmeting boven ' + CONFIG.courierCm + ' cm: controleer of een koerier dit kan vervoeren.');
      if (X.isDG) w.push('Gevaarlijke goederen via koerier: controleer of de koerier ADR-zendingen accepteert.');
    }
    if (X.cross && X.mode && X.mode !== 'Sea freight' && SEA_ONLY.indexOf(S.incoterm) >= 0) {
      w.push('Incoterm ' + S.incoterm + ' is alleen bedoeld voor zeevracht; kies bijvoorbeeld FCA of DAP.');
    }
    if (X.heavy && !X.liftReq) w.push('Zwaar colli (≥ ' + CONFIG.heavyKg + ' kg): vermeld heftruck/kraan bij laad- en losinstructies.');
    if (X.oversize) w.push('Overmaat (grens: ' + CONFIG.oversize.l + ' x ' + CONFIG.oversize.w + ' x ' + CONFIG.oversize.h + ' cm): Shipping beoordeelt of speciaal transport of ontheffing nodig is.');
    if (S.weather === 'Ja' && S.vehicle === 'Open wagen') w.push('Weersgevoelige goederen op een open wagen: onderbouw de bescherming (dekzeil).');
    if (X.isDG && X.mode === 'Air freight') w.push('Luchtvracht met gevaarlijke goederen vraagt IATA-documentatie (DGD).');
    if (X.isDG && X.mode === 'Sea freight') w.push('Zeevracht met gevaarlijke goederen vraagt IMDG-documentatie (DGD).');
    // vergelijk colli-specificatie met hoofdvragen
    var maxKg = 0, maxL = 0, maxW = 0, maxH = 0, tot = 0;
    colliModel.forEach(function (r) {
      var q = Math.max(0, Math.floor(num(r.qty)));
      maxKg = Math.max(maxKg, num(r.gross)); maxL = Math.max(maxL, num(r.l)); maxW = Math.max(maxW, num(r.w)); maxH = Math.max(maxH, num(r.h));
      tot += q * num(r.gross);
    });
    if (maxKg > X.maxKg || maxL > X.maxL || maxW > X.maxW || maxH > X.maxH) {
      w.push('Colli-specificatie bevat een groter/zwaarder colli dan bij de hoofdvragen is opgegeven. Pas de hoofdvragen aan zodat de juiste vereisten gelden.');
    }
    if (tot > 0 && X.totalKg > 0 && Math.abs(tot - X.totalKg) / X.totalKg > 0.05) {
      w.push('Totaal bruto gewicht colli-specificatie (' + r2(tot) + ' kg) wijkt >5% af van de hoofdvragen (' + r2(X.totalKg) + ' kg).');
    }
    ['attPack', 'attInv', 'attSds', 'attCustoms'].forEach(function (id) {
      if (S[id] === 'Volgt') w.push('Bijlage staat op "Volgt": Shipping kan pas starten zodra deze is ontvangen.');
    });
    return w.filter(function (x, i) { return w.indexOf(x) === i; });
  }

  function extras(S, X) {
    var e = [];
    if (X.domestic) e.push('Binnenlands transport: geen Incoterm, douane of handelsfactuur nodig');
    if (X.cross) e.push('Incoterm + plaats, douanestatus, waarde goederen en handelsfactuur');
    if (X.customs && !X.isUK) e.push('HS-codes, land van oorsprong en douane-/exportdocumenten');
    if (X.isUK) e.push('UK: partlijst met waarde, HS (8 cijfers) en oorsprong; netto gewicht op paklijst');
    if (X.isROHU) e.push('Totale goederenwaarde voor ' + (X.dz === 'RO' ? 'UIT-code (RO)' : 'EKAER-code (HU)'));
    if (X.isDG) e.push('UN-nummer, klasse, verpakkingsgroep, benaming en SDS');
    if (X.liftReq) e.push('Zwaartepunt/hijspunten en tekening/hijsplan');
    if (X.urgent) e.push('Prioriteit, reden spoed en deadline leverdatum');
    if (X.defence) e.push('Defence: VOG en aanmelding chauffeur');
    return e;
  }

  /* =====================================================================
   * 8. Formulier opbouwen
   * ===================================================================== */
  var form, REG = {}, touched = {}, showAll = false;

  function optList(f) {
    var list = [];
    (f.opts || []).forEach(function (o) {
      if (typeof o === 'string') list.push({ v: o, l: o });
      else if (o.group) list.push(o);
      else list.push({ v: o.v, l: o.l || o.v, off: o.off, dis: o.dis });
    });
    return list;
  }

  function buildField(f) {
    var wrap = h('div', { class: 'field' + (f.wide ? ' wide' : ''), 'data-field': f.id });
    var entry = { f: f, wrap: wrap, ctrl: null, msg: h('div', { class: 'msg', 'aria-live': 'polite' }), mark: h('span', { class: 'req', hidden: true, text: ' *' }) };
    var cid = 'f-' + f.id;

    if (f.type === 'info') {
      entry.box = h('div', { class: 'notice ' + (f.tone || 'info'), role: 'note' });
      wrap.appendChild(entry.box);
      REG[f.id] = entry;
      return wrap;
    }
    if (f.type === 'colli' || f.type === 'parts') {
      wrap.appendChild(h('div', { class: 'lbl' }, f.label, entry.mark));
      var t = f.type === 'colli' ? (colliTable = buildColliTable(update)) : (partsTable = buildPartsTable(update));
      entry.table = t;
      wrap.appendChild(t.el);
      REG[f.id] = entry;
      return wrap;
    }

    var ctrl;
    if (f.type === 'radio') {
      wrap.appendChild(h('div', { class: 'lbl', id: cid + '-l' }, f.label, entry.mark));
      var seg = h('div', { class: 'seg', role: 'radiogroup', 'aria-labelledby': cid + '-l' });
      f.opts.forEach(function (o, i) {
        var rid = cid + '-' + i;
        seg.appendChild(h('input', { type: 'radio', name: f.id, id: rid, value: o }));
        seg.appendChild(h('label', { for: rid, text: o }));
      });
      wrap.appendChild(seg);
      ctrl = seg;
    } else if (f.type === 'checks') {
      wrap.appendChild(h('div', { class: 'lbl', id: cid + '-l' }, f.label, entry.mark));
      var box = h('div', { class: 'checks', role: 'group', 'aria-labelledby': cid + '-l' });
      f.opts.forEach(function (o, i) {
        var cb = h('input', { type: 'checkbox', value: o, id: cid + '-' + i });
        box.appendChild(h('label', { for: cid + '-' + i }, cb, o));
      });
      wrap.appendChild(box);
      ctrl = box;
    } else if (f.type === 'check') {
      var cb2 = h('input', { type: 'checkbox', id: cid });
      var lab = h('label', { class: 'check-single', for: cid }, cb2, h('span', {}, f.label, entry.mark));
      wrap.appendChild(lab);
      ctrl = cb2;
    } else {
      wrap.appendChild(h('label', { class: 'lbl', for: cid }, f.label, entry.mark));
      if (f.type === 'select') {
        ctrl = h('select', { id: cid });
        ctrl.appendChild(h('option', { value: '', text: '— kies —' }));
        optList(f).forEach(function (o) {
          if (o.group) {
            var og = h('optgroup', { label: o.group });
            o.items.forEach(function (it) { og.appendChild(h('option', { value: it.v, text: it.l })); });
            ctrl.appendChild(og);
          } else {
            ctrl.appendChild(h('option', { value: o.v, text: o.l, disabled: o.off ? true : false }));
          }
        });
      } else if (f.type === 'textarea') {
        ctrl = h('textarea', { id: cid, placeholder: f.ph || false });
      } else if (f.type === 'computed') {
        ctrl = h('input', { type: 'text', id: cid, readonly: true, tabindex: '-1' });
      } else {
        ctrl = h('input', { type: f.type, id: cid, placeholder: f.ph || false, min: f.type === 'number' ? '0' : false, step: f.type === 'number' ? 'any' : false });
      }
      wrap.appendChild(ctrl);
    }
    if (f.def != null && f.type !== 'radio' && f.type !== 'checks' && f.type !== 'check') {
      ctrl.value = typeof f.def === 'function' ? f.def() : f.def;
    }
    if (f.id === 'customsStatus') ctrl.addEventListener('change', function () { ctrl.setAttribute('data-user', '1'); });
    if (f.type !== 'radio' && f.type !== 'checks') ctrl.setAttribute('data-id', f.id);
    if (f.hint) wrap.appendChild(h('span', { class: 'hint', text: f.hint }));
    wrap.appendChild(entry.msg);
    entry.ctrl = ctrl;
    REG[f.id] = entry;
    return wrap;
  }

  function buildForm() {
    form.innerHTML = '';
    REG = {}; touched = {}; showAll = false;
    colliModel.length = 0; colliModel.push(newColli());
    partsModel.length = 0; partsModel.push(newPart());
    SECTIONS.forEach(function (sec) {
      var card = h('section', { class: 'card', id: 'sec-' + sec.id, 'data-sec': sec.id });
      card.appendChild(h('h2', {}, h('span', { class: 'num', text: sec.main ? '►' : '' }), h('span', { class: 'ttl', text: sec.title })));
      if (sec.intro) card.appendChild(h('p', { class: 'intro', text: sec.intro }));
      var grid = h('div', { class: 'grid' });
      sec.fields.forEach(function (f) { grid.appendChild(buildField(f)); });
      card.appendChild(grid);
      form.appendChild(card);
      sec.el = card;
    });
    // gate-melding na de hoofdvragen
    var gate = h('div', { class: 'notice warn', id: 'gate', text: 'Beantwoord eerst alle hoofdvragen. Daarna verschijnen de overige secties.' });
    form.insertBefore(gate, SECTIONS[1].el);
  }

  /* =====================================================================
   * 9. Waarden verzamelen en evalueren
   * ===================================================================== */
  function collect() {
    var S = {};
    Object.keys(REG).forEach(function (id) {
      var r = REG[id], f = r.f;
      if (f.type === 'info' || f.type === 'colli' || f.type === 'parts') return;
      if (f.type === 'radio') { var c = r.ctrl.querySelector('input:checked'); S[id] = c ? c.value : ''; }
      else if (f.type === 'checks') { S[id] = [].slice.call(r.ctrl.querySelectorAll('input:checked')).map(function (i) { return i.value; }); }
      else if (f.type === 'check') S[id] = r.ctrl.checked;
      else S[id] = r.ctrl.value;
    });
    S.colli = colliModel; S.parts = partsModel;
    return S;
  }

  function isReq(f, S, X) { return typeof f.req === 'function' ? !!f.req(S, X) : !!f.req; }

  function evalField(f, S, X) {
    var out = { vis: !f.vis || !!f.vis(S, X), req: false, missing: false, err: '' };
    if (!out.vis) return out;
    if (f.type === 'info' || f.type === 'computed' || f.type === 'colli' || f.type === 'parts') return out;
    out.req = isReq(f, S, X);
    var v = S[f.id];
    if (isEmpty(v)) { out.missing = out.req; return out; }
    if (f.type === 'number' && !(num(v) > 0)) out.err = 'Voer een getal groter dan 0 in.';
    else if (f.validate) out.err = f.validate(typeof v === 'string' ? v : v, S, X) || '';
    return out;
  }

  function evaluate(S, X) {
    var ev = { secs: [], items: [], fields: {} };
    var mainOK = true;
    SECTIONS.forEach(function (sec) {
      var visible = sec.main || (mainOK && (!sec.vis || !!sec.vis(S, X)));
      var s = { sec: sec, visible: visible, issues: 0 };
      if (visible) {
        sec.fields.forEach(function (f) {
          var e = evalField(f, S, X);
          ev.fields[f.id] = e;
          if (!e.vis) return;
          if (f.type === 'colli' || f.type === 'parts') {
            var tbl = REG[f.id].table;
            var show = showAll;
            var probs = tbl.mark(S, X, show);
            probs.forEach(function (p) { ev.items.push({ sec: sec, id: f.id, label: f.label + ' – ' + p }); s.issues++; });
            return;
          }
          if (e.missing || e.err) {
            ev.items.push({ sec: sec, id: f.id, label: (f.short || f.label || 'Bevestiging') + (e.err ? ' – ' + e.err : '') });
            s.issues++;
          }
        });
      }
      if (sec.main && s.issues > 0) mainOK = false;
      ev.secs.push(s);
    });
    ev.mainOK = mainOK;
    return ev;
  }

  /* =====================================================================
   * 10. update(): afgeleide waarden, zichtbaarheid, validatie, UI
   * ===================================================================== */
  var lastNum = '';

  function applyDerived(S, X) {
    // berekende velden
    Object.keys(REG).forEach(function (id) {
      var r = REG[id], f = r.f;
      if (f.type === 'computed') {
        var val = f.calc(S, X);
        if (r.ctrl.value !== val) r.ctrl.value = val;
      }
      if (f.type === 'select') {
        // dynamische opties uitschakelen
        var list = optList(f), changed = false;
        var opts = r.ctrl.querySelectorAll('option');
        [].forEach.call(opts, function (op) {
          var def = null;
          list.forEach(function (o) { if (o.v === op.value) def = o; });
          if (!def) return;
          var dis = !!def.off || (typeof def.dis === 'function' && !!def.dis(S, X));
          if (op.disabled !== dis) op.disabled = dis;
          if (dis && r.ctrl.value === op.value && !def.off) { r.ctrl.value = ''; changed = true; }
        });
      }
    });
    // douanestatus automatisch voorstellen
    var cs = REG.customsStatus;
    if (cs && X.cross && X.customsSuggest && !cs.ctrl.getAttribute('data-user') && cs.ctrl.value !== X.customsSuggest) cs.ctrl.value = X.customsSuggest;
    // UK: waarde = som van partlijst, in EUR
    var gv = REG.goodsValue, cur = REG.currency;
    if (gv) {
      if (X.isUK) {
        var tv = String(Math.round(X.partsTotal * 100) / 100);
        if (gv.ctrl.value !== tv) gv.ctrl.value = tv;
        gv.ctrl.readOnly = true; cur.ctrl.value = 'EUR'; cur.ctrl.disabled = true;
      } else { gv.ctrl.readOnly = false; cur.ctrl.disabled = false; }
    }
  }

  function update() {
    if (!form) return;
    var S = collect(), X = derive(S);
    applyDerived(S, X);
    S = collect(); X = derive(S);
    var ev = evaluate(S, X);
    var warn = warnings(S, X);

    // zichtbaarheid per sectie en veld; nummering
    var n = 0;
    ev.secs.forEach(function (s) {
      s.sec.el.hidden = !s.visible;
      if (!s.visible) return;
      var numEl = s.sec.el.querySelector('.num');
      if (s.sec.main) { numEl.textContent = '►'; }
      else { n++; s.n = n; numEl.textContent = String(n); }
      s.sec.fields.forEach(function (f) {
        var e = ev.fields[f.id], r = REG[f.id];
        r.wrap.hidden = !e.vis;
        if (!e.vis) return;
        if (f.type === 'info') {
          var html = f.text(S, X);
          if (r.box.innerHTML !== html) r.box.innerHTML = html;
          return;
        }
        if (f.type === 'colli' || f.type === 'parts' || f.type === 'computed') {
          if (f.type === 'colli' || f.type === 'parts') r.mark.hidden = false;
          return;
        }
        r.mark.hidden = !e.req;
        var show = showAll || touched[f.id];
        var msg = '';
        if (show && e.missing) msg = 'Dit veld is verplicht.';
        else if (show && e.err) msg = e.err;
        else if (e.err && f.id === 'pkReady') msg = e.err;
        r.msg.textContent = msg;
        r.wrap.classList.toggle('invalid', !!msg);
      });
    });
    var gate = document.getElementById('gate');
    if (gate) gate.hidden = ev.mainOK;

    // status
    var total = ev.items.length;
    var ok = ev.mainOK && total === 0;
    var st = document.getElementById('status');
    st.className = 'status ' + (ok ? 'status-ok' : 'status-bad');
    st.textContent = ok ? '✅ Formulier compleet' : '❌ Formulier incompleet' + (total ? ' (' + total + ' open)' : '');
    document.getElementById('btnMail').disabled = !ok;

    renderSide(S, X, ev, warn);
    document.title = (ok ? '✅ ' : '') + 'Transportaanvraag Shipping';
    update.last = { S: S, X: X, ev: ev, warn: warn, ok: ok };
  }

  function renderSide(S, X, ev, warn) {
    var chips = document.getElementById('chips');
    chips.innerHTML = '';
    if (!ev.mainOK && !X.routeKnown) {
      chips.appendChild(h('span', { class: 'chip empty', text: 'Nog geen route gekozen' }));
    } else {
      if (X.routeKnown) chips.appendChild(h('span', { class: 'chip', text: X.oName + ' → ' + X.dName }));
      if (X.direction) chips.appendChild(h('span', { class: 'chip ' + (X.customs ? 'amber' : 'green'), text: X.direction }));
      if (X.isUK) chips.appendChild(h('span', { class: 'chip amber', text: 'UK' }));
      if (X.isROHU) chips.appendChild(h('span', { class: 'chip amber', text: X.dz === 'RO' ? 'UIT-code (RO)' : 'EKAER (HU)' }));
      if (S.maxKg) chips.appendChild(h('span', { class: 'chip ' + (X.liftReq ? 'red' : X.heavy ? 'amber' : ''), text: X.oversize ? 'Overmaat' : X.veryHeavy ? 'Zwaar (hijsen)' : X.heavy ? 'Zwaar' : 'Standaard maat' }));
      if (S.dg) chips.appendChild(h('span', { class: 'chip ' + (X.isDG ? 'red' : ''), text: X.isDG ? 'Gevaarlijke goederen' : 'Geen DG' }));
      if (S.urgent) chips.appendChild(h('span', { class: 'chip ' + (X.urgent ? 'red' : ''), text: X.urgent ? 'Spoed' : 'Geen spoed' }));
      if (X.defence) chips.appendChild(h('span', { class: 'chip red', text: 'Defence' }));
    }
    var ul = document.getElementById('extras');
    ul.innerHTML = '';
    var ex = ev.mainOK ? extras(S, X) : [];
    if (!ex.length) ul.appendChild(h('li', { class: 'none', text: ev.mainOK ? 'Alleen de standaardvragen.' : 'Verschijnt na de hoofdvragen.' }));
    ex.forEach(function (t) { ul.appendChild(h('li', { text: t })); });
    var wEl = document.getElementById('warnings');
    wEl.innerHTML = '';
    warn.forEach(function (t) { wEl.appendChild(h('div', { class: 'w', text: t })); });

    var nav = document.getElementById('nav');
    nav.innerHTML = '';
    ev.secs.forEach(function (s) {
      if (!s.visible) return;
      var a = h('a', { href: '#sec-' + s.sec.id },
        h('span', { class: 'dot' + (s.issues ? '' : ' done'), text: s.issues ? '!' : '✓' }),
        h('span', { text: (s.sec.main ? '' : s.n + '. ') + s.sec.title }),
        s.issues ? h('span', { class: 'count', text: String(s.issues) }) : null);
      nav.appendChild(h('li', {}, a));
    });
  }

  /* =====================================================================
   * 11. Samenvatting / export / mail
   * ===================================================================== */
  function fieldValue(f, S) {
    var v = S[f.id];
    if (f.type === 'check') return v ? 'Ja' : '';
    if (f.type === 'checks') return v.join(', ');
    if (f.type === 'date') return fmtDate(v);
    if (f.type === 'datetime-local') return fmtDT(v);
    if (f.type === 'select' && v === 'OTHER') return '';
    if (f.type === 'computed') return f.calc(S, update.last.X);
    return String(v == null ? '' : v).trim();
  }

  function buildSummary() {
    var L = update.last, S = L.S, X = L.X, ev = L.ev;
    var out = [];
    out.push('TRANSPORTAANVRAAG | SHIPPING');
    out.push('Route: ' + X.oName + ' → ' + X.dName + '  |  ' + X.direction + '  |  DG: ' + S.dg + '  |  Spoed: ' + S.urgent);
    out.push('Status: ' + (L.ok ? 'compleet' : 'INCOMPLEET'));
    ev.secs.forEach(function (s) {
      if (!s.visible) return;
      out.push('');
      out.push((s.sec.main ? 'HOOFDVRAGEN' : s.n + '. ' + s.sec.title.toUpperCase()));
      s.sec.fields.forEach(function (f) {
        var e = ev.fields[f.id];
        if (!e || !e.vis || f.type === 'info') return;
        if (f.type === 'colli') {
          S.colli.forEach(function (r, i) {
            out.push('  Colli ' + (i + 1) + ': ' + r.qty + 'x ' + r.type + ' | ' + r.desc + ' | ' + r.l + 'x' + r.w + 'x' + r.h + ' cm | netto ' + r.net + ' kg | bruto ' + r.gross + ' kg | stapelbaar: ' + r.stack + ' | ' + r.haz);
          });
          var q = 0, g = 0, c = 0;
          S.colli.forEach(function (r) { var k = Math.max(0, Math.floor(num(r.qty))); q += k; g += k * num(r.gross); c += k * num(r.l) * num(r.w) * num(r.h) / 1e6; });
          out.push('  Totaal: ' + q + ' colli | ' + r3(c) + ' m³ | ' + r2(g) + ' kg bruto');
          return;
        }
        if (f.type === 'parts') {
          S.parts.forEach(function (r, i) {
            out.push('  Part ' + (i + 1) + ': ' + r.pn + ' | ' + r.desc + ' | qty ' + r.qty + ' | COO ' + r.coo + ' | € ' + r.val + ' | HS ' + r.hs);
          });
          out.push('  Totaal waarde: € ' + r2(X.partsTotal));
          return;
        }
        var val = fieldValue(f, S);
        if (val) out.push('  ' + (f.label || 'Bevestiging').replace(/:$/, '') + ': ' + val);
      });
    });
    return out.join('\n');
  }

  function msg(t) { document.getElementById('actionMsg').textContent = t; }

  function copyText(t) {
    if (navigator.clipboard && navigator.clipboard.writeText) return navigator.clipboard.writeText(t);
    return new Promise(function (res, rej) {
      var ta = h('textarea', { style: 'position:fixed;left:-9999px' });
      ta.value = t; document.body.appendChild(ta); ta.select();
      try { document.execCommand('copy') ? res() : rej(); } catch (e) { rej(e); } document.body.removeChild(ta);
    });
  }
  function download(name, text, mime) {
    var a = h('a', { href: URL.createObjectURL(new Blob([text], { type: mime })), download: name });
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
  }
  function fileBase() {
    var S = update.last.S;
    return 'Transportaanvraag_' + (S.project || 'concept').replace(/[^\w-]+/g, '_') + '_' + todayStr();
  }

  function mailToShipping() {
    var L = update.last;
    if (!L.ok) { msg('Het formulier is nog niet compleet.'); return; }
    var S = L.S, X = L.X;
    var summary = buildSummary();
    var subject = 'Transportaanvraag | ' + S.project + ' | ' + S.po + ' | ' + X.oName + ' → ' + X.dName + (X.urgent ? ' | SPOED' : '');
    var atts = [['Paklijst', S.attPack], ['Factuur', S.attInv], ['Foto’s', S.attPhotos], ['SDS', S.attSds], ['Douane/export', S.attCustoms], ['Tekeningen/hijsplan', S.attDraw]]
      .filter(function (a) { return a[1] === 'Bijgevoegd'; }).map(function (a) { return a[0]; });
    var body = 'Hallo Shipping,\n\nHierbij de transportaanvraag. De volledige samenvatting staat op het klembord: plak deze hieronder met Ctrl+V.\n' +
      (atts.length ? '\nBijlagen om toe te voegen: ' + atts.join(', ') + '.\n' : '') + '\nMet vriendelijke groet,\n' + S.name;
    copyText(summary).then(function () {
      msg('Samenvatting gekopieerd. Plak deze in de geopende e-mail en voeg de bijlagen toe.');
    }, function () { msg('Kopiëren naar klembord mislukt; gebruik “Download samenvatting”.'); });
    window.location.href = 'mailto:' + CONFIG.shippingMail + '?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(body);
  }

  function exportJson() {
    var L = update.last, vals = {};
    Object.keys(L.ev.fields).forEach(function (id) {
      var e = L.ev.fields[id];
      if (e.vis && L.S[id] !== undefined && !isEmpty(L.S[id])) vals[id] = L.S[id];
    });
    return JSON.stringify({ gegenereerd: new Date().toISOString(), compleet: L.ok, route: { van: L.X.oName, naar: L.X.dName, type: L.X.direction }, waarden: vals, colli: L.S.colli, parts: L.X.isUK ? L.S.parts : undefined }, null, 2);
  }

  function showMissing() {
    var L = update.last, box = document.getElementById('missing');
    showAll = true; update(); L = update.last;
    box.innerHTML = '';
    if (L.ok) { box.hidden = true; msg('Alles is ingevuld. U kunt het formulier nu naar Shipping sturen.'); return; }
    box.hidden = false;
    box.appendChild(h('b', { text: 'Nog te corrigeren (' + L.ev.items.length + '):' }));
    var ul = h('ul');
    if (!L.ev.mainOK && !L.ev.items.length) ul.appendChild(h('li', { text: 'Beantwoord de hoofdvragen.' }));
    L.ev.items.forEach(function (it) {
      var a = h('a', { href: '#', text: (it.sec.main ? 'Hoofdvragen' : 'Sectie ' + (L.ev.secs.filter(function (s) { return s.sec === it.sec; })[0].n)) + ': ' + it.label });
      a.addEventListener('click', function (e) {
        e.preventDefault();
        var r = REG[it.id];
        r.wrap.scrollIntoView({ behavior: 'smooth', block: 'center' });
        var focusEl = r.ctrl && (r.ctrl.focus ? r.ctrl : r.ctrl.querySelector('input'));
        if (focusEl && focusEl.focus) setTimeout(function () { focusEl.focus({ preventScroll: true }); }, 350);
      });
      ul.appendChild(h('li', {}, a));
    });
    box.appendChild(ul);
    box.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  }

  /* =====================================================================
   * 12. Init
   * ===================================================================== */
  function init() {
    form = document.getElementById('form');
    buildForm();
    form.addEventListener('input', update);
    form.addEventListener('change', function (e) {
      var w = e.target.closest && e.target.closest('[data-field]');
      if (w) touched[w.getAttribute('data-field')] = true;
      update();
    });
    form.addEventListener('focusout', function (e) {
      var w = e.target.closest && e.target.closest('[data-field]');
      if (w) touched[w.getAttribute('data-field')] = true;
      update();
    });
    form.addEventListener('submit', function (e) { e.preventDefault(); });

    document.getElementById('btnCheck').addEventListener('click', showMissing);
    document.getElementById('btnMail').addEventListener('click', mailToShipping);
    document.getElementById('btnCopy').addEventListener('click', function () {
      copyText(buildSummary()).then(function () { msg('Samenvatting gekopieerd naar het klembord.'); }, function () { msg('Kopiëren mislukt.'); });
    });
    document.getElementById('btnTxt').addEventListener('click', function () { download(fileBase() + '.txt', buildSummary(), 'text/plain;charset=utf-8'); msg('Samenvatting gedownload.'); });
    document.getElementById('btnJson').addEventListener('click', function () { download(fileBase() + '.json', exportJson(), 'application/json'); msg('Gegevens gedownload.'); });
    document.getElementById('btnPrint').addEventListener('click', function () { window.print(); });
    document.getElementById('btnReset').addEventListener('click', function () {
      if (!window.confirm('Alle ingevulde gegevens wissen?')) return;
      buildForm(); document.getElementById('missing').hidden = true; msg(''); update();
    });
    update();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
