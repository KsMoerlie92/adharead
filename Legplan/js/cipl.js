// ── Create CIPL — Commercial Invoice & Packing List per Shipment ─────────────
// Bron: Moederlijst (Sheet2), gefilterd op kolom "Shipment". Opmaak en rijposities volgen
// het voorbeeld CIPL_YN1320_*.xlsx (kop rij 4-48, kolomkoppen rij 54, data vanaf rij 55,
// "Grand Total" direct onder de laatste regel, met totaal Gross én Nett).
//
//   initCIPL()          — vult de shipmentkeuze (aangeroepen door showTab('cipl'))
//   renderCIPL()        — bouwt de CIPL voor het gekozen shipment (nieuw: edits vervallen)
//   CIPL.*              — kolomkeuze, bewerken, export (Excel via ExcelJS, PDF via afdrukvenster)

// ── CIPL-format blad van de Moederlijst (kopgegevens: vessel, afzender, consignee …) ──
function parseCIPLSheet(ws) {
  if (!ws) return [null, null, null];
  const cv = (r, c) => {
    const cell = ws[XLSX.utils.encode_cell({r: r-1, c: c-1})];
    return cell ? String(cell.v == null ? '' : cell.v).trim() : '';
  };
  const meta = {
    docNo: cv(7,10), project: cv(22,2), shipmentNr: cv(23,2), date: cv(25,2),
    packages: cv(40,5), nettWgt: cv(41,5), grossWgt: cv(42,5), vessel: cv(45,5), delivery: cv(48,5), c51: cv(51,3),
  };
  return [[], [], meta];
}

var CIPL = (function () {
  'use strict';

  var LOGO_PNG = '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';
  var LOGO_JPG = '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';

  /* ── kolommen (volgorde = standaard, zoals het voorbeeld A…U) ─────────────── */
  // k = sleutel · h = kop · src = kolomnaam in Moederlijst · i = vaste index als terugval
  // t = type (txt|num) · w = breedte (tekens) · on = standaard zichtbaar · grp = hoort bij collo
  var COLS = [
    { k: 'ref',    h: 'Deliveryref',          src: /^delivery\s*ref/i,            i: 0,  t: 'txt', w: 13, on: 1 },
    { k: 'collo',  h: 'Collo',                src: /^collo/i,                     i: 17, t: 'txt', w: 9,  on: 1 },
    { k: 'pkg',    h: 'Type of packaging',    src: /type\s*of\s*packag/i,         i: 18, t: 'txt', w: 13, on: 1, grp: 1 },
    { k: 'len',    h: 'Length cm',            src: /^length/i,                    i: 19, t: 'num', w: 9,  on: 1, grp: 1 },
    { k: 'wid',    h: 'Width cm',             src: /^width/i,                     i: 20, t: 'num', w: 9,  on: 1, grp: 1 },
    { k: 'hgt',    h: 'Height cm',            src: /^heig/i,                      i: 21, t: 'num', w: 9,  on: 1, grp: 1 },
    { k: 'gross',  h: 'Gross Weight (kg)',    src: /gross\s*weight/i,             i: 23, t: 'num', w: 12, on: 1, grp: 1, tot: 1, fmt: '#,##0.00' },
    { k: 'nett',   h: 'Nett Weight (kg)',     src: /ne?tt?\s*weight/i,            i: 24, t: 'num', w: 12, on: 1, grp: 1, tot: 1, fmt: '#,##0.00' },
    { k: 'mark',   h: 'Mark/Label',           src: /^mark/i,                      i: 7,  t: 'txt', w: 16, on: 1 },
    { k: 'desc',   h: 'Item description',     src: /item\s*description/i,         i: 4,  t: 'txt', w: 60, on: 1 },
    { k: 'codeSup',h: 'Code Supplier',        src: /^code\s*supplier/i,           i: 8,  t: 'txt', w: 16, on: 1 },
    { k: 'hs',     h: 'HSCode',               src: /^hs\s*code/i,                 i: 14, t: 'txt', w: 12, on: 1 },
    { k: 'make',   h: 'Make',                 src: /^make/i,                      i: 11, t: 'txt', w: 18, on: 1 },
    { k: 'mat',    h: 'Material',             src: /^material/i,                  i: 12, t: 'txt', w: 16, on: 1 },
    { k: 'coo',    h: 'Country of Origin',    src: /country\s*of\s*origin/i,      i: 13, t: 'txt', w: 10, on: 1 },
    { k: 'cont',   h: 'Container',            src: /^container/i,                 i: 38, t: 'txt', w: 14, on: 1 },
    { k: 'uom',    h: 'UoM',                  src: /^uom/i,                       i: 6,  t: 'txt', w: 9,  on: 1 },
    { k: 'qty',    h: 'Quantity',             src: /^quantity/i,                  i: 5,  t: 'num', w: 10, on: 1, tot: 1, fmt: '#,##0.##' },
    { k: 'vpp',    h: 'Value p.p.',           calc: 1,                            t: 'num', w: 14, on: 1, val: 1, fmt: '#,##0.00' },
    { k: 'vtot',   h: 'Value Tot.',           calc: 1,                            t: 'num', w: 16, on: 1, val: 1, tot: 1, fmt: '#,##0.00' },
    // extra (standaard uit)
    { k: 'po',     h: 'IHC PO',               src: /^ihc\s*po/i,                  i: 2,  t: 'txt', w: 13, on: 0 },
    { k: 'item',   h: 'Item #',               src: /^item\s*#/i,                  i: 3,  t: 'txt', w: 8,  on: 0 },
    { k: 'sup',    h: 'Name of Supplier',     src: /name\s*of\s*supplier/i,       i: 10, t: 'txt', w: 24, on: 0 },
    { k: 'serial', h: 'Serial number',        src: /^serial/i,                    i: 9,  t: 'txt', w: 14, on: 0 },
    { k: 'vpcEur', h: 'Value pc (EUR)',       src: /value\s*pc/i,                 i: 15, t: 'num', w: 13, on: 0, fmt: '#,##0.00' },
    { k: 'vtotEur',h: 'Value total (EUR)',    src: /value\s*total/i,              i: 16, t: 'num', w: 14, on: 0, tot: 1, fmt: '#,##0.00' },
    { k: 'loc',    h: 'Location',             src: /^location/i,                  i: 26, t: 'txt', w: 12, on: 0 }
  ];
  var COLMAP = {}; COLS.forEach(function (c) { COLMAP[c.k] = c; });

  /* ── state ───────────────────────────────────────────────────────────────── */
  var S = {
    ship: '', rows: [], hdr: null, cfg: null, conv: { mode: 'markup', factor: 1.16, markup: 38 },
    hdrByShip: {}, cfgByShip: {}, src: null, colIdx: null, shipCol: null, ship_meta: null
  };
  var LS_COLS = 'ihc-cipl-cols-v1';

  /* ── helpers ─────────────────────────────────────────────────────────────── */
  function E(s) { return (typeof esc === 'function') ? esc(s) : String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function $(id) { return document.getElementById(id); }
  function clean(s) { return String(s == null ? '' : s).replace('(leeg)', '').replace(/ /g, ' ').trim(); }
  /** "2,083.30" · "1.234,56" · "16.03" · " -   " → getal of null */
  function pNum(v) {
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = clean(v).replace(/[€$\s]/g, '');
    if (!s || /^[-–—]+$/.test(s)) return null;
    var hasC = s.indexOf(',') >= 0, hasD = s.indexOf('.') >= 0;
    if (hasC && hasD) s = s.lastIndexOf('.') > s.lastIndexOf(',') ? s.replace(/,/g, '') : s.replace(/\./g, '').replace(',', '.');
    else if (hasC) s = /^-?\d{1,3}(,\d{3})+$/.test(s) ? s.replace(/,/g, '') : s.replace(',', '.');
    else if (hasD && /^-?\d{1,3}(\.\d{3}){2,}$/.test(s)) s = s.replace(/\./g, '');
    var n = parseFloat(s); return isNaN(n) ? null : n;
  }
  function pDate(v) {
    if (v instanceof Date && !isNaN(v)) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), 12));
    var s = clean(v), m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], 12));
    if ((m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})/.exec(s))) { var y = +m[3]; if (y < 100) y += 2000; return new Date(Date.UTC(y, +m[2] - 1, +m[1], 12)); }
    return null;
  }
  function dmy(d) { return d ? ('0' + d.getUTCDate()).slice(-2) + '/' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '/' + d.getUTCFullYear() : ''; }
  function f2(n) { return (n || 0).toLocaleString('nl-NL', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function ls(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; } return null; }
  function MF() { try { return (typeof moederFile !== 'undefined') ? moederFile : null; } catch (e) { return null; } }

  /* ── valuta ──────────────────────────────────────────────────────────────── */
  function conv(v) {
    v = Number(v) || 0; var c = S.conv;
    if (c.mode === 'eur2usd') return v * (c.factor || 1);
    if (c.mode === 'usd2eur') return v / (c.factor || 1);
    if (c.mode === 'markup') return v * (1 + (c.markup || 0) / 100);
    return v;
  }
  function cur() { return (S.conv.mode === 'markup' || S.conv.mode === 'eur2usd') ? 'USD' : 'EUR'; }
  function colHead(c) { return c.val ? c.h + ' ' + cur() + ' customs' : c.h; }

  /* ── kolomconfig (volgorde + zichtbaar), per shipment, laatste keuze wordt standaard ── */
  function defaultCfg() { return COLS.map(function (c) { return { k: c.k, on: !!c.on }; }); }
  function loadCfg() {
    var base = defaultCfg(), raw = ls(LS_COLS);
    if (!raw) return base;
    try {
      var saved = JSON.parse(raw), out = [], seen = {};
      saved.forEach(function (s) { if (COLMAP[s.k] && !seen[s.k]) { seen[s.k] = 1; out.push({ k: s.k, on: !!s.on }); } });
      base.forEach(function (b) { if (!seen[b.k]) out.push(b); });
      return out;
    } catch (e) { return base; }
  }
  function saveCfg() { S.cfgByShip[S.ship] = S.cfg; ls(LS_COLS, JSON.stringify(S.cfg)); }
  function visCols() { return S.cfg.filter(function (c) { return c.on; }).map(function (c) { return COLMAP[c.k]; }); }

  /* ── Moederlijst → shipmentlijst ─────────────────────────────────────────── */
  function indexCols(headers) {
    var idx = {};
    COLS.forEach(function (c) {
      if (c.calc) return;
      idx[c.k] = headers.find(function (h) { return h && c.src.test(String(h).trim()); }) || headers[c.i];
    });
    idx.proj = headers.find(function (h) { return /^project/i.test(String(h).trim()); }) || headers[1];
    idx.shipment = headers.find(function (h) { return h && /^shipment/i.test(String(h).trim()); }) || headers[31];
    idx.dest = headers.find(function (h) { return /^destination/i.test(String(h).trim()); }) || headers[32];
    idx.etd = headers.find(function (h) { return /^etd\b/i.test(String(h).trim()); }) || headers[35];
    idx.eta = headers.find(function (h) { return /^eta\b/i.test(String(h).trim()); }) || headers[36];
    idx.bl = headers.find(function (h) { return /^bl\s*#|bill\s*of\s*lading|^bl\b/i.test(String(h).trim()); }) || headers[39];
    return idx;
  }

  function shipmentList(data) {
    var s = {}, order = [];
    var shipCol = S.colIdx.shipment;
    data.forEach(function (r) { var k = clean(r[shipCol]); if (k && !s[k]) { s[k] = 1; order.push(k); } });
    return order.sort(function (a, b) { return (parseFloat(a) || 9999) - (parseFloat(b) || 9999) || a.localeCompare(b); });
  }

  function makeRow(src) {
    var row = { c: {}, ov: {}, eur: {} };
    COLS.forEach(function (c) {
      if (c.calc) return;
      var raw = src[S.colIdx[c.k]];
      if (c.t === 'num') {
        var n = pNum(raw);
        if (c.k === 'gross' || c.k === 'nett') { row.c[c.k] = (n && n !== 0) ? n : null; }
        else row.c[c.k] = n;
      } else row.c[c.k] = clean(raw);
    });
    row.eur.pc = pNum(src[S.colIdx.vpcEur]); row.eur.tot = pNum(src[S.colIdx.vtotEur]);
    recalc(row);
    return row;
  }
  // Value p.p. = Value pc (EUR) × factor · Value Tot. = Value total (EUR) × factor (zoals in het voorbeeld)
  function recalc(row) {
    var pc = row.eur.pc, tot = row.eur.tot, q = row.c.qty;
    if (!row.ov.vpp) row.c.vpp = pc != null ? conv(pc) : null;
    if (!row.ov.vtot) row.c.vtot = tot != null && tot !== 0 ? conv(tot) : (pc != null && q != null ? conv(pc * q) : null);
  }

  /* ── kopgegevens ─────────────────────────────────────────────────────────── */
  var DEF_SHIPPER = 'IHC HOLLAND B.V.\nSMITWEG 6\n2961AW KINDERDIJK\nTHE NETHERLANDS\nNL813735646B 01\nPh # T +31 88 015 2871\nE-mail: vri.coolen@royalihc.com';
  var DEF_CONSIGNEE = 'PACIFIC SHIPBUILDING JSC\nLot 6 Tien Lang\nCommune, Hai Phong City\nVietnam\nMrs. Kristyle Le\nPh # +84-225-3943273\nE-mail: business@daiduongship.vn\nTAX Code: 0200751035';

  function cellRange(ws, col, r1, r2) {
    var out = [];
    for (var r = r1; r <= r2; r++) { var c = ws[XLSX.utils.encode_cell({ r: r - 1, c: col - 1 })]; var v = c ? clean(c.v) : ''; if (v) out.push(v); }
    return out.join('\n');
  }
  function sheetParties() {
    var mf = MF(), ws = mf && mf._ciplWs;
    if (!ws || typeof XLSX === 'undefined') return {};
    try { return { shipper: cellRange(ws, 2, 12, 18), consignee: cellRange(ws, 2, 29, 36), notify: cellRange(ws, 5, 29, 36) }; } catch (e) { return {}; }
  }

  function buildHeader(ship, srcRows) {
    var first = srcRows[0] || {}, ix = S.colIdx, mf = MF();
    var proj = clean(first[ix.proj]) || 'YN1320';
    var projNum = (proj.match(/\d+/) || [''])[0];
    var bls = [], conts = [], seenB = {}, seenC = {};
    srcRows.forEach(function (r) {
      var b = clean(r[ix.bl]).replace(/\s+/g, ' '); if (b && !seenB[b]) { seenB[b] = 1; bls.push(b); }
      var c = clean(r[ix.cont]); if (c && !seenC[c]) { seenC[c] = 1; conts.push(c); }
    });
    var blc = bls.map(function (b) { return b.toUpperCase().replace(/[^A-Z0-9]/g, ''); });
    var isAir = /\bair/i.test(conts.join(' ')) || blc.some(function (b) { return /^(AWB|DHL)/.test(b) || /^\d{11}$/.test(b); });
    var destRaw = ''; srcRows.some(function (r) { destRaw = clean(r[ix.dest]); return !!destRaw; });
    var dest = /hanoi/i.test(destRaw) ? 'HANOI' : 'HAIPHONG';
    var carrier = isAir ? 'Not applicable' : '';
    if (!isAir && window.CarrierTrack) {
      var cand = blc.concat(conts.map(function (c) { return c.toUpperCase().replace(/[^A-Z0-9]/g, ''); }));
      for (var i = 0; i < cand.length && !carrier; i++) { var kk = CarrierTrack.detect(cand[i]); if (kk) carrier = CarrierTrack.carriers[kk].name; }
    }
    var etd = null, eta = null;
    srcRows.some(function (r) { etd = pDate(r[ix.etd]); return !!etd; });
    srcRows.some(function (r) { eta = pDate(r[ix.eta]); return !!eta; });
    var parties = sheetParties(), now = new Date();
    return {
      docNo: now.getFullYear() + projNum + '-SH' + ship,
      project: proj, shipNo: 'SH' + ship, date: dmy(new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate(), 12))),
      shipper: parties.shipper || DEF_SHIPPER, consignee: parties.consignee || DEF_CONSIGNEE, notify: parties.notify || parties.consignee || DEF_CONSIGNEE,
      transLabel: isAir ? 'AIR WAYBILL:' : 'BILL OF LADING:', transNo: bls.join(' / '),
      carrier: carrier, vessel: (mf && mf.ciplMeta && mf.ciplMeta.vessel) || '',
      etdPlace: isAir ? 'AMSTERDAM' : 'ROTTERDAM', etd: etd ? dmy(etd) : clean(srcRows.length ? srcRows[0][ix.etd] : ''), etdDate: etd,
      etaPlace: dest, eta: eta ? dmy(eta) : clean(srcRows.length ? srcRows[0][ix.eta] : ''), etaDate: eta,
      terms: 'CIP ' + (dest === 'HANOI' ? 'Hanoi' : 'Haiphong'), pkgOverride: ''
    };
  }

  /* ── totalen en controles ────────────────────────────────────────────────── */
  function totals() {
    var t = { gross: 0, nett: 0, qty: 0, val: 0, collos: {} , n: S.rows.length };
    S.rows.forEach(function (r) {
      t.gross += r.c.gross || 0; t.nett += r.c.nett || 0; t.qty += r.c.qty || 0; t.val += r.c.vtot || 0;
      if (r.c.collo) t.collos[r.c.collo] = 1;
    });
    t.pkgs = Object.keys(t.collos).length;
    return t;
  }
  /** per collo: ontbreken afmetingen/gewicht op álle regels van dat collo? */
  function colloIssues() {
    var by = {};
    S.rows.forEach(function (r) {
      var k = r.c.collo; if (!k) return;
      var o = by[k] || (by[k] = { first: r, gross: false, nett: false, dim: false });
      if (r.c.gross) o.gross = true; if (r.c.nett) o.nett = true; if (r.c.len && r.c.wid && r.c.hgt) o.dim = true;
    });
    var out = { noWeight: [], noDim: [] };
    Object.keys(by).forEach(function (k) { var o = by[k]; if (!o.gross || !o.nett) out.noWeight.push(k); if (!o.dim) out.noDim.push(k); });
    return out;
  }
  function rowWarn(r, key) {
    var v = r.c[key];
    if (key === 'hs') return !v;
    if (key === 'coo') return !v;
    if (key === 'qty') return !(v > 0);
    if (key === 'vtot' || key === 'vpp') return !(v > 0);
    if (key === 'desc') return !v;
    return false;
  }

  /* ── UI: samenvatting, kopformulier, kolommen, tabel ─────────────────────── */
  var FORM = [
    ['docNo', 'Number', 'text'], ['project', 'Project', 'text'], ['shipNo', 'Shipment', 'text'], ['date', 'Date', 'text'],
    ['transLabel', 'Transportdocument', 'select'], ['transNo', 'B/L · AWB nummer', 'text'], ['carrier', 'Carrier', 'text'], ['vessel', 'Vessel', 'text'],
    ['etdPlace', 'ETD vanuit', 'text'], ['etd', 'ETD datum', 'text'], ['etaPlace', 'ETA naar', 'text'], ['eta', 'ETA datum', 'text'],
    ['terms', 'Delivery terms', 'text'], ['pkgOverride', 'Aantal packages (leeg = automatisch)', 'text'],
    ['shipper', "Shipper's name and address", 'area'], ['consignee', 'Consignee', 'area'], ['notify', 'Notify', 'area']
  ];

  function renderSummary() {
    var el = $('cp-summary'); if (!el) return;
    if (!S.rows.length) { el.innerHTML = ''; return; }
    var t = totals(), ci = colloIssues();
    var miss = function (k) { return S.rows.filter(function (r) { return rowWarn(r, k); }).length; };
    var chip = function (cls, txt) { return '<span class="cp-chip ' + cls + '">' + txt + '</span>'; };
    var h = chip('', '<b>' + t.n + '</b> regels') + chip('', '<b>' + t.pkgs + '</b> packages') +
      chip('', 'Gross <b>' + f2(t.gross) + '</b> kg') + chip('', 'Nett <b>' + f2(t.nett) + '</b> kg') +
      chip('', 'Value <b>' + f2(t.val) + '</b> ' + cur());
    if (ci.noWeight.length) h += chip('warn', '⚠ ' + ci.noWeight.length + ' collo zonder gross/nett: ' + E(ci.noWeight.slice(0, 6).join(', ')) + (ci.noWeight.length > 6 ? ' …' : ''));
    if (ci.noDim.length) h += chip('warn', '⚠ ' + ci.noDim.length + ' collo zonder afmetingen');
    [['hs', 'zonder HS-code'], ['coo', 'zonder Country of Origin'], ['vtot', 'zonder waarde']].forEach(function (x) { var n = miss(x[0]); if (n) h += chip('warn', '⚠ ' + n + ' regels ' + x[1]); });
    if (!ci.noWeight.length && !ci.noDim.length && !miss('hs') && !miss('coo') && !miss('vtot')) h += chip('ok', '✓ compleet');
    el.innerHTML = h;
  }

  function renderForm() {
    var el = $('cp-form'); if (!el) return;
    var H = S.hdr; if (!H) { el.innerHTML = ''; return; }
    el.innerHTML = FORM.map(function (f) {
      var id = 'cph-' + f[0], v = E(H[f[0]]), cls = f[2] === 'area' ? ' cp-wide' : '';
      var inp = f[2] === 'area' ? '<textarea id="' + id + '" rows="8" data-h="' + f[0] + '">' + v + '</textarea>'
        : f[2] === 'select' ? '<select id="' + id + '" data-h="' + f[0] + '">' + ['AIR WAYBILL:', 'BILL OF LADING:', 'SEAWAYBILL:'].map(function (o) { return '<option' + (H[f[0]] === o ? ' selected' : '') + '>' + o + '</option>'; }).join('') + '</select>'
        : '<input id="' + id + '" data-h="' + f[0] + '" value="' + v + '">';
      return '<label class="cp-f' + cls + '"><span>' + f[1] + '</span>' + inp + '</label>';
    }).join('');
    el.querySelectorAll('[data-h]').forEach(function (i) {
      i.addEventListener('input', function () { S.hdr[i.dataset.h] = i.value; if (/^(etd|eta)$/.test(i.dataset.h)) { S.hdr[i.dataset.h + 'Date'] = pDate(i.value); } });
    });
  }

  function renderColsBox() {
    var el = $('cp-cols'); if (!el || !S.cfg) return;
    el.innerHTML = S.cfg.map(function (c, i) {
      var d = COLMAP[c.k];
      return '<div class="cp-col' + (c.on ? ' on' : '') + '" data-i="' + i + '" draggable="true">' +
        '<label><input type="checkbox" ' + (c.on ? 'checked' : '') + ' data-on="' + i + '"> ' + E(colHead(d)) + '</label>' +
        '<span class="cp-mv"><button title="Naar links" data-mv="' + i + '" data-d="-1">◀</button><button title="Naar rechts" data-mv="' + i + '" data-d="1">▶</button></span></div>';
    }).join('');
    el.querySelectorAll('[data-on]').forEach(function (b) { b.addEventListener('change', function () { S.cfg[+b.dataset.on].on = b.checked; saveCfg(); renderColsBox(); renderTable(); }); });
    el.querySelectorAll('[data-mv]').forEach(function (b) { b.addEventListener('click', function () { move(+b.dataset.mv, +b.dataset.d); }); });
    // slepen
    var dragI = null;
    el.querySelectorAll('.cp-col').forEach(function (n) {
      n.addEventListener('dragstart', function (e) { dragI = +n.dataset.i; e.dataTransfer.effectAllowed = 'move'; try { e.dataTransfer.setData('text/plain', String(dragI)); } catch (x) { } });
      n.addEventListener('dragover', function (e) { e.preventDefault(); n.classList.add('over'); });
      n.addEventListener('dragleave', function () { n.classList.remove('over'); });
      n.addEventListener('drop', function (e) { e.preventDefault(); var to = +n.dataset.i; if (dragI == null || dragI === to) return; var it = S.cfg.splice(dragI, 1)[0]; S.cfg.splice(to, 0, it); dragI = null; saveCfg(); renderColsBox(); renderTable(); });
    });
  }
  function move(i, d) {
    var j = i + d; if (j < 0 || j >= S.cfg.length) return;
    var t = S.cfg[i]; S.cfg[i] = S.cfg[j]; S.cfg[j] = t; saveCfg(); renderColsBox(); renderTable();
  }

  function inputCell(r, ri, c) {
    var v = r.c[c.k], warn = rowWarn(r, c.k) ? ' warn' : '', num = c.t === 'num';
    var shown = v == null ? '' : (num ? (c.val || c.fmt && /0\.00/.test(c.fmt) ? (+v).toFixed(2) : String(v)) : v);
    return '<td class="cp-td' + (num ? ' r' : '') + warn + (r.ov[c.k] ? ' ov' : '') + '"><input class="cp-in' + (num ? ' r' : '') + '" data-r="' + ri + '" data-k="' + c.k + '" value="' + E(shown) + '" style="width:' + Math.max(6, Math.min(c.w, 46)) + 'ch"></td>';
  }

  function renderTable() {
    var thead = $('cp-thead'), tbody = $('cp-tbody'), tfoot = $('cp-tfoot'); if (!thead) return;
    var cols = visCols();
    if (!S.rows.length) { thead.innerHTML = ''; tbody.innerHTML = '<tr><td class="cp-empty">Kies een shipment om de CIPL op te bouwen.</td></tr>'; tfoot.innerHTML = ''; return; }
    var hasVal = cols.some(function (c) { return c.val; });
    var g1 = '', first = true;
    if (hasVal) g1 = '<tr class="cp-g">' + cols.map(function (c) { var s = c.val && first ? 'Values' : ''; if (c.val) first = false; return '<th>' + s + '</th>'; }).join('') + '</tr>';
    thead.innerHTML = g1 + '<tr>' + cols.map(function (c) { return '<th class="' + (c.t === 'num' ? 'r' : '') + '">' + E(colHead(c)) + '</th>'; }).join('') + '</tr>';
    tbody.innerHTML = S.rows.map(function (r, ri) { return '<tr>' + cols.map(function (c) { return inputCell(r, ri, c); }).join('') + '</tr>'; }).join('');
    renderFoot();
    tbody.querySelectorAll('.cp-in').forEach(function (inp) {
      inp.addEventListener('input', function () { edit(+inp.dataset.r, inp.dataset.k, inp.value, inp); });
    });
  }

  function renderFoot() {
    var tfoot = $('cp-tfoot'); if (!tfoot) return;
    var cols = visCols(), t = totals(), firstDone = false;
    tfoot.innerHTML = '<tr class="cp-tot">' + cols.map(function (c) {
      var s = '';
      if (!firstDone) { s = 'Grand Total'; firstDone = true; }
      else if (c.k === 'gross') s = f2(t.gross); else if (c.k === 'nett') s = f2(t.nett);
      else if (c.k === 'qty') s = t.qty.toLocaleString('nl-NL', { maximumFractionDigits: 2 });
      else if (c.k === 'vtot') s = f2(t.val); else if (c.k === 'vtotEur') s = f2(S.rows.reduce(function (a, r) { return a + (r.c.vtotEur || 0); }, 0));
      return '<td class="' + (c.t === 'num' ? 'r' : '') + '">' + s + '</td>';
    }).join('') + '</tr>';
    renderSummary();
  }

  function edit(ri, k, raw, inp) {
    var r = S.rows[ri], c = COLMAP[k]; if (!r) return;
    if (c.t === 'num') { var n = pNum(raw); r.c[k] = raw.trim() === '' ? null : n; if (c.val) r.ov[k] = true; if (k === 'vpcEur') { r.eur.pc = r.c[k]; } if (k === 'vtotEur') { r.eur.tot = r.c[k]; } }
    else r.c[k] = raw;
    if (k === 'qty' || k === 'vpcEur' || k === 'vtotEur') { recalc(r); refreshRowVals(ri); }
    var td = inp.closest('td'); if (td) td.classList.toggle('warn', rowWarn(r, k));
    renderFoot();
  }
  function refreshRowVals(ri) {
    ['vpp', 'vtot'].forEach(function (k) {
      var inp = document.querySelector('.cp-in[data-r="' + ri + '"][data-k="' + k + '"]');
      if (inp && document.activeElement !== inp) { var v = S.rows[ri].c[k]; inp.value = v == null ? '' : v.toFixed(2); inp.closest('td').classList.toggle('warn', rowWarn(S.rows[ri], k)); }
    });
  }

  /* ── opbouwen ────────────────────────────────────────────────────────────── */
  function build(ship) {
    var mf = MF(), data = (mf && mf.data) || [];
    var srcRows = data.filter(function (r) { return clean(r[S.colIdx.shipment]) === ship; });
    S.ship = ship;
    S.rows = srcRows.map(makeRow);
    S.hdr = buildHeader(ship, srcRows);
    S.hdrByShip[ship] = S.hdr;
    S.cfg = S.cfgByShip[ship] || loadCfg();
    S.cfgByShip[ship] = S.cfg;
  }

  function renderAll() {
    var st = $('cipl-stats');
    if (!S.ship) { if (st) st.textContent = 'Kies een Shipment # om de CIPL te maken.'; }
    else if (st) st.innerHTML = '<span>Shipment <strong>#' + E(S.ship) + '</strong></span><span>' + S.rows.length + ' regels uit de Moederlijst</span>';
    renderForm(); renderColsBox(); renderTable(); renderSummary();
    var box = $('cp-main'); if (box) box.style.display = S.ship ? '' : 'none';
  }

  function onShipChange(keepEdits) {
    var sel = $('cipl-shipment-sel'), ship = sel ? sel.value : '';
    if (!ship) { S.ship = ''; S.rows = []; S.hdr = null; renderAll(); return; }
    if (keepEdits && S.hdrByShip[ship] && S.ship === ship) { renderAll(); return; }
    build(ship); renderAll();
  }

  /* ── valuta-balk ─────────────────────────────────────────────────────────── */
  function onConv() {
    var mode = $('cipl-conv-mode').value, f = parseFloat($('cipl-conv-factor').value), m = parseFloat($('cipl-conv-markup').value);
    S.conv = { mode: mode, factor: isNaN(f) || f <= 0 ? 1.16 : f, markup: isNaN(m) ? 0 : m };
    $('cipl-conv-factor-wrap').style.display = (mode === 'eur2usd' || mode === 'usd2eur') ? '' : 'none';
    $('cipl-conv-markup-wrap').style.display = mode === 'markup' ? '' : 'none';
    var note = $('cipl-conv-note');
    if (note) note.textContent = mode === 'eur2usd' ? 'Value × ' + S.conv.factor + ' → USD' : mode === 'usd2eur' ? 'Value ÷ ' + S.conv.factor + ' → EUR' : mode === 'markup' ? 'Customs value = EUR + ' + S.conv.markup + '% (USD customs)' : 'Originele waarden (EUR)';
    S.rows.forEach(function (r) { r.ov = {}; recalc(r); });
    renderColsBox(); renderTable();
  }

  /* ── Excel (ExcelJS) — opmaak en posities volgens het voorbeeld ──────────── */
  function colLetter(n) { var s = ''; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function stamp() { var d = new Date(), p = function (n) { return ('0' + n).slice(-2); }; return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()); }

  async function exportExcel() {
    if (!S.rows.length) { alert('Kies eerst een shipment.'); return; }
    if (typeof ExcelJS === 'undefined') { alert('ExcelJS is niet geladen (controleer de internetverbinding).'); return; }
    var H = S.hdr, cols = visCols(), t = totals();
    var wb = new ExcelJS.Workbook(); wb.creator = 'IHC Expedite 2.0';
    var ws = wb.addWorksheet('CIPL', { pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
    var FIRST = 2, LAST = FIRST + cols.length - 1; // kolom B = eerste tabelkolom
    var HDR = 54, DATA = 55;
    ws.getColumn(1).width = 4.5;
    cols.forEach(function (c, i) { ws.getColumn(FIRST + i).width = c.w; });
    var A10 = { name: 'Arial', size: 10 }, A10B = { name: 'Arial', size: 10, bold: true };
    var put = function (addr, v, font, extra) { var c = ws.getCell(addr); c.value = v; c.font = font || A10; if (extra) Object.assign(c, extra); return c; };

    // logo
    var img = wb.addImage({ base64: LOGO_PNG, extension: 'png' });
    ws.addImage(img, { tl: { col: 0, row: 0 }, ext: { width: 224, height: 66 } });
    // customs value factor (zoals V1/W1 in het voorbeeld)
    if (S.conv.mode === 'markup') { put(colLetter(LAST + 2) + '1', 'Customs value factor', A10); var wc = put(colLetter(LAST + 3) + '1', (S.conv.markup || 0) / 100, A10); wc.numFmt = '0%'; }

    // titel (gecentreerd over de tabelbreedte)
    var title = function (r, text, size, h) { ws.mergeCells(r, FIRST, r, LAST); var c = ws.getCell(r, FIRST); c.value = text; c.font = { name: 'Arial', size: size, bold: true }; c.alignment = { horizontal: 'center', vertical: 'middle' }; if (h) ws.getRow(r).height = h; };
    title(4, 'NO COMMERCIAL VALUE/PACKING LIST', 22, 28);
    title(7, 'Number: ' + H.docNo, 14, 22.5);
    title(8, 'Value for customs purpose only', 12, 22.5);

    put('B10', "SHIPPER'S NAME AND ADDRESS:", A10B);
    String(H.shipper).split('\n').forEach(function (l, i) { put('B' + (12 + i), l, A10); });
    put('B22', 'PROJECT : ' + H.project, A10B);
    put('B23', 'SHIPMENT ' + H.shipNo, A10B);
    put('B25', 'DATE: ' + H.date, A10B);
    put('B27', 'CONSIGNEE', A10B); put('E27', 'NOTIFY ', A10B);
    String(H.consignee).split('\n').forEach(function (l, i) { put('B' + (29 + i), l, A10); });
    String(H.notify).split('\n').forEach(function (l, i) { put('E' + (29 + i), l, A10); });

    var pk = H.pkgOverride && pNum(H.pkgOverride) != null ? pNum(H.pkgOverride) : t.pkgs;
    var facts = [
      [40, 'NUMBER OF PACKAGES :  ', pk, null], [41, 'TOTAL NETT WEIGHT (KG): ', Math.round(t.nett * 100) / 100, '#,##0.00'],
      [42, 'TOTAL GROSS WEIGHT (KG): ', Math.round(t.gross * 100) / 100, '#,##0.00'], [43, 'TOTAL VALUE ' + cur() + ': ', Math.round(t.val * 100) / 100, '#,##0.00'],
      [44, H.transLabel, H.transNo, null], [45, 'CARRIER:', H.carrier, null],
      [46, 'ETD ' + H.etdPlace, H.etdDate || H.etd, 'dd/mm/yyyy'], [47, 'ETA ' + H.etaPlace, H.etaDate || H.eta, 'dd/mm/yyyy'], [48, 'DELIVERY TERMS: ', H.terms, null]
    ];
    facts.forEach(function (f) {
      put('B' + f[0], f[1], A10B);
      var c = put('E' + f[0], f[2] === '' ? null : f[2], A10, { alignment: { horizontal: 'left' } });
      if (f[3]) c.numFmt = f[3];
      if (f[3] === 'dd/mm/yyyy' && f[2] instanceof Date) c.value = new Date(Date.UTC(f[2].getUTCFullYear(), f[2].getUTCMonth(), f[2].getUTCDate()));
    });
    put('B51', 'Shipment', A10); put('C51', String(S.ship), A10);

    // groepskop "Values"
    var vFirst = cols.findIndex(function (c) { return c.val; });
    if (vFirst >= 0) put(colLetter(FIRST + vFirst) + '53', 'Values', A10B);
    // kolomkoppen
    var GREEN = 'FF95CA82';
    ws.getRow(HDR).height = 26;
    cols.forEach(function (c, i) {
      var cell = ws.getCell(HDR, FIRST + i); cell.value = colHead(c);
      cell.font = A10B; cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
      cell.alignment = { vertical: 'center', horizontal: c.t === 'num' ? 'right' : 'left', wrapText: true };
      cell.border = { bottom: { style: 'thin' } };
    });
    // data
    var money = cur() === 'USD' ? '[$$-409]#,##0.00' : '[$€-2] #,##0.00';
    S.rows.forEach(function (r, ri) {
      var xr = DATA + ri;
      cols.forEach(function (c, i) {
        var v = r.c[c.k]; if (v === '' || v == null) return;
        var cell = ws.getCell(xr, FIRST + i); cell.value = v; cell.font = A10;
        if (c.t === 'num') { cell.numFmt = c.val ? money : (c.fmt || 'General'); cell.alignment = { horizontal: 'right' }; }
        else { cell.numFmt = '@'; }
      });
    });
    // Grand Total (direct onder de laatste regel): Gross + Nett + Quantity + Value als SUM
    var gt = DATA + S.rows.length, GF = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
    cols.forEach(function (c, i) {
      var cell = ws.getCell(gt, FIRST + i); cell.fill = GF; cell.font = A10B; cell.border = { top: { style: 'thin' } };
      if (i === 0) cell.value = 'Grand Total';
      else if (c.tot) {
        var L = colLetter(FIRST + i), sum = c.k === 'gross' ? t.gross : c.k === 'nett' ? t.nett : c.k === 'qty' ? t.qty : c.k === 'vtot' ? t.val : S.rows.reduce(function (a, r) { return a + (r.c[c.k] || 0); }, 0);
        cell.value = { formula: 'SUM(' + L + DATA + ':' + L + (gt - 1) + ')', result: sum };
        cell.numFmt = c.val ? money : (c.fmt || '0.00'); cell.alignment = { horizontal: 'right' };
      }
    });
    var buf = await wb.xlsx.writeBuffer();
    var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = 'CIPL_' + H.project + '_' + H.shipNo + '_' + stamp() + '.xlsx';
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }

  /* ── PDF (afdrukvenster) ─────────────────────────────────────────────────── */
  function exportPDF() {
    if (!S.rows.length) { alert('Kies eerst een shipment.'); return; }
    var H = S.hdr, cols = visCols(), t = totals(), nl = function (s) { return E(s).replace(/\n/g, '<br>'); };
    var pk = H.pkgOverride && pNum(H.pkgOverride) != null ? pNum(H.pkgOverride) : t.pkgs;
    var cell = function (r, c) {
      var v = r.c[c.k]; if (v == null || v === '') return '<td></td>';
      if (c.t === 'num') return '<td class="r">' + (c.val || /0\.00/.test(c.fmt || '') ? f2(v) : E(v)) + '</td>';
      return '<td>' + E(v) + '</td>';
    };
    var body = S.rows.map(function (r, i) { return '<tr class="' + (i % 2 ? 'odd' : 'even') + '">' + cols.map(function (c) { return cell(r, c); }).join('') + '</tr>'; }).join('');
    var first = true;
    var tot = '<tr class="tot-row">' + cols.map(function (c) {
      var s = ''; if (first) { s = 'Grand Total'; first = false; }
      else if (c.k === 'gross') s = f2(t.gross); else if (c.k === 'nett') s = f2(t.nett); else if (c.k === 'qty') s = t.qty.toLocaleString('nl-NL', { maximumFractionDigits: 2 }); else if (c.k === 'vtot') s = f2(t.val);
      return '<td class="' + (c.t === 'num' ? 'r' : '') + '">' + s + '</td>';
    }).join('') + '</tr>';
    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>CIPL — ' + E(H.project) + ' ' + E(H.shipNo) + '</title><style>' +
      '@page{size:A4 landscape;margin:10mm}*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:7pt;color:#000}' +
      '.top{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:6px}.top img{height:38px}.tt{text-align:center;flex:1}' +
      '.tt b{display:block;font-size:12pt;letter-spacing:.06em;border-bottom:2px solid #D91F2C;padding-bottom:3px;margin-bottom:3px}.tt span{display:block;font-size:8pt;font-weight:700;color:#444}' +
      '.grid{display:grid;grid-template-columns:1fr 1fr 1fr;border:1px solid #aaa;margin-bottom:4px}.grid>div{padding:4px 6px;border-right:1px solid #aaa;line-height:1.5;font-size:7.5pt}.grid>div:last-child{border-right:none}' +
      '.grid i{display:block;font-style:normal;font-size:6pt;font-weight:800;text-transform:uppercase;color:#555;border-bottom:1px solid #ddd;margin-bottom:2px}' +
      '.meta{display:flex;border:1px solid #aaa;border-top:none;margin-bottom:5px}.meta div{flex:1;padding:3px 6px;border-right:1px solid #aaa;font-size:7pt}.meta div:last-child{border-right:none}.meta i{display:block;font-style:normal;font-size:5.5pt;font-weight:800;text-transform:uppercase;color:#777}' +
      'table{width:100%;border-collapse:collapse;font-size:6.5pt}th{background:#95CA82;color:#000;padding:3px;text-align:left;font-size:6pt;border:1px solid #6aa456}td{padding:2px 3px;border-bottom:1px solid #e8e8e8;vertical-align:top}.r{text-align:right}tr.even td{background:#f7f8fa}' +
      '.tot-row td{background:#95CA82;font-weight:800;border-top:2px solid #333}.foot{margin-top:4px;font-size:6pt;color:#888;border-top:1px solid #ddd;padding-top:3px}' +
      '</style></head><body>' +
      '<div class="top"><img src="data:image/jpeg;base64,' + LOGO_JPG + '"><div class="tt"><b>NO COMMERCIAL VALUE/PACKING LIST</b><span>Number: ' + E(H.docNo) + '</span><span>Value for customs purpose only</span></div>' +
      '<div style="text-align:right;line-height:1.7"><div><b>Date:</b> ' + E(H.date) + '</div><div><b>Shipment:</b> ' + E(H.shipNo) + '</div></div></div>' +
      '<div class="grid"><div><i>Shipper\'s name and address</i>' + nl(H.shipper) + '<br><br><b>PROJECT : ' + E(H.project) + '</b><br>SHIPMENT ' + E(H.shipNo) + '</div><div><i>Consignee</i>' + nl(H.consignee) + '</div><div><i>Notify</i>' + nl(H.notify) + '</div></div>' +
      '<div class="meta"><div><i>Packages</i><b>' + pk + '</b></div><div><i>Nett weight (kg)</i><b>' + f2(t.nett) + '</b></div><div><i>Gross weight (kg)</i><b>' + f2(t.gross) + '</b></div><div><i>Total value ' + cur() + '</i><b>' + f2(t.val) + '</b></div>' +
      '<div><i>' + E(H.transLabel.replace(':', '')) + '</i><b>' + E(H.transNo || '—') + '</b></div><div><i>Carrier / vessel</i><b>' + E([H.carrier, H.vessel].filter(Boolean).join(' · ') || '—') + '</b></div><div><i>ETD ' + E(H.etdPlace) + ' → ETA ' + E(H.etaPlace) + '</i><b>' + E(H.etd || '—') + ' → ' + E(H.eta || '—') + '</b></div><div><i>Delivery terms</i><b>' + E(H.terms) + '</b></div></div>' +
      '<table><thead><tr>' + cols.map(function (c) { return '<th class="' + (c.t === 'num' ? 'r' : '') + '">' + E(colHead(c)) + '</th>'; }).join('') + '</tr></thead><tbody>' + body + tot + '</tbody></table>' +
      '<div class="foot">Royal IHC — Commercial Invoice &amp; Packing List · ' + E(H.project) + ' · ' + E(H.shipNo) + '</div></body></html>';
    var w = window.open('', '_blank', 'width=1200,height=820');
    if (!w) { alert('Pop-up geblokkeerd — sta pop-ups toe voor deze pagina.'); return; }
    w.document.write(html); w.document.close(); w.focus(); setTimeout(function () { w.print(); }, 500);
  }

  /* ── publiek ─────────────────────────────────────────────────────────────── */
  function init() {
    var mf = MF(), st = $('cipl-stats'), sel = $('cipl-shipment-sel');
    if (!mf || !mf.data) { if (st) st.textContent = 'Laad de Moederlijst om een CIPL te maken.'; return; }
    if (!mf.ciplMeta && mf._ciplWs) { try { mf.ciplMeta = parseCIPLSheet(mf._ciplWs)[2]; } catch (e) { mf.ciplMeta = {}; } }
    if (S.src !== mf.data) {
      S.src = mf.data; S.colIdx = indexCols(Object.keys(mf.data[0] || {})); S.hdrByShip = {}; S.cfgByShip = {}; S.ship = ''; S.rows = []; S.hdr = null;
    }
    var prev = sel.value, list = shipmentList(mf.data);
    sel.innerHTML = '<option value="">— kies shipment —</option>' + list.map(function (v) { return '<option value="' + E(v) + '">' + E(v) + '</option>'; }).join('');
    if (prev && list.indexOf(prev) >= 0) sel.value = prev;
    if (!S.cfg) S.cfg = loadCfg();
    onConv();   // synchroniseert zichtbare velden met S.conv
    if (sel.value) { if (S.ship !== sel.value) build(sel.value); renderAll(); } else renderAll();
  }
  function resetCols() { S.cfg = defaultCfg(); saveCfg(); renderColsBox(); renderTable(); }
  function allCols(on) { S.cfg.forEach(function (c) { c.on = on ? true : !!COLMAP[c.k].on; }); saveCfg(); renderColsBox(); renderTable(); }

  return { init: init, rebuild: function () { onShipChange(false); }, change: function () { onShipChange(true); }, onConv: onConv, exportExcel: exportExcel, exportPDF: exportPDF, resetCols: resetCols, allCols: allCols, _S: S, _pNum: pNum };
})();

// compat: showTab('cipl') roept initCIPL() aan; het shipment-keuzeveld roept renderCIPL() aan
function initCIPL() { CIPL.init(); }
function renderCIPL() { CIPL.rebuild(); }
