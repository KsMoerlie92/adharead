/* ============================================================================
 *  IHC Expedite 2.0 — Shipment tracking (Fase 1: uit de Moederlijst)
 *  ---------------------------------------------------------------------------
 *  Bron: Moederlijst (Sheet2). Per Shipment (kol AF) worden Destination (AG),
 *  Date load (AI), ETD (AJ), ETA (AK), Container (AM) en BL # (AN) gebundeld.
 *  Status, voortgang en kaartpositie volgen uit die gegevens + de datum van vandaag:
 *    Gebundeld   : shipmentnummer, nog geen boeking (geen BL/container/ETD)
 *    Geboekt     : BL en/of container bekend, vertrek (ETD) ligt nog in de toekomst
 *    Onderweg    : ETD verstreken, ETA nog niet
 *    Aangekomen* : ETA verstreken (* afgeleid, geen bevestiging van de rederij)
 *  De positie van het schip is een SCHATTING (lineair tussen ETD en ETA langs een
 *  vereenvoudigde route). Schipnaam en echte positie komen pas met een rederij-API
 *  (Fase 2); tot die tijd opent "Track & Trace" de pagina van de rederij
 *  (CarrierTrack) en kan het schip handmatig worden genoteerd.
 *
 *  Gebruik: ShipTrack.show()  — wordt aangeroepen door showTab('tracking')
 * ========================================================================== */
(function () {
  'use strict';

  /* ── havens en routes ───────────────────────────────────────────────────── */
  var ORIGIN = 'RTM';           // vertrekhaven (aanname: Rotterdam)
  var DEFAULT_DEST = 'HPH';     // bij lege Destination: Hai Phong (aanname, wordt getoond)
  var PORTS = {
    RTM: { n: 'Rotterdam', c: 'NL', ll: [51.95, 4.1] },
    HPH: { n: 'Hai Phong', c: 'VN', ll: [20.85, 106.75] },
    AMS: { n: 'Amsterdam (AMS)', c: 'NL', ll: [52.31, 4.76] },
    HAN: { n: 'Hanoi (HAN)', c: 'VN', ll: [21.22, 105.81] }
  };
  /* Routekeuze: de scheepvaart mijdt nu Rode Zee/Suez/Hormuz en vaart om Afrika (Kaap de Goede Hoop).
     Zet ROUTE_VIA op 'suez' om terug te gaan naar de route door het Suezkanaal. */
  var ROUTE_VIA = 'cape';
  var SUEZ = [[50.2, -3.5], [36.0, -6.2], [37.0, 10.5], [34.0, 26], [31.3, 32.4], [27.5, 34], [20, 38.5],
    [12.6, 43.3], [12.0, 50], [7, 62], [6, 80], [5.5, 95], [1.3, 103.8], [8, 106], [14, 109.5]];
  var CAPE = [[50.2, -3.5], [46.5, -6.5], [43.5, -10.2], [36.8, -10.2], [28, -16.5], [15, -18.5], [5, -12],
    [2, -3], [-5, 7], [-17, 9], [-30, 14], [-36, 19], [-36, 26], [-33, 34], [-30, 45], [-18, 65], [-2, 82],
    [5.5, 95], [1.3, 103.8], [8, 106], [14, 109.5]];
  function seaRoute(dest) { return [PORTS[ORIGIN].ll].concat(ROUTE_VIA === 'suez' ? SUEZ : CAPE, [PORTS[dest].ll]); }
  function airRoute(dest) { return [PORTS.AMS.ll, PORTS[dest].ll]; }
  function destKey(raw) {
    var s = String(raw || '');
    if (/hanoi/i.test(s)) return 'HAN';
    if (/hai\s*phong|haiphong/i.test(s)) return 'HPH';
    return null;
  }

  /* ── hulpfuncties ───────────────────────────────────────────────────────── */
  var DAY = 86400000;
  function clean(s) { return String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
  function trim(v) { return String(v == null ? '' : v).replace('(leeg)', '').trim(); }
  function numPlain(v) { var n = parseFloat(String(v == null ? '' : v).replace(',', '.')); return isNaN(n) ? 0 : n; }
  /** dag-eerst (dd/mm/jjjj, dd-mm-jjjj, dd.mm.jjjj) of ISO; anders null ("tba", "to be advised") */
  function pDate(v) {
    if (v instanceof Date && !isNaN(v)) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), 12));
    var s = trim(v), m;
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(s))) return new Date(Date.UTC(+m[1], +m[2] - 1, +m[3], 12));
    if ((m = /^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})/.exec(s))) {
      var y = +m[3]; if (y < 100) y += 2000;
      return new Date(Date.UTC(y, +m[2] - 1, +m[1], 12));
    }
    return null;
  }
  function fmt(d) { return d ? d.toLocaleDateString('nl-NL', { day: '2-digit', month: 'short', year: 'numeric', timeZone: 'UTC' }).replace('.', '') : '—'; }
  function fmtS(d) { return d ? d.toLocaleDateString('nl-NL', { day: '2-digit', month: 'short', timeZone: 'UTC' }).replace('.', '') : '—'; }
  function isContainerNo(s) { return /^[A-Z]{4}\d{7}$/.test(clean(s)); }
  // moederFile is een top-level `let` (state.js): bereikbaar op naam, niet via window.
  function MF() { try { return (typeof moederFile !== 'undefined') ? moederFile : null; } catch (e) { return null; } }
  function E(s) { return (typeof esc === 'function') ? esc(s) : String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function store(k, v) { try { if (v === undefined) return localStorage.getItem(k); localStorage.setItem(k, v); } catch (e) { return null; } }

  /* ── Moederlijst → zendingen ────────────────────────────────────────────── */
  function colFinder(headers) {
    return function (re, idx) {
      var h = headers.find(function (x) { return x && re.test(String(x).trim()); });
      return h || headers[idx];
    };
  }

  function buildShipments(data) {
    if (!data || !data.length) return [];
    var headers = Object.keys(data[0]);
    var f = colFinder(headers);
    var cShip = headers.find(function (h) { return h && /^shipment/i.test(String(h).trim()); }) || headers[31];
    var cDest = f(/^destination/i, 32), cLoad = f(/date\s*load/i, 34), cEtd = f(/^etd\b/i, 35), cEta = f(/^eta\b/i, 36),
        cCont = f(/^container/i, 38), cBl = f(/^bl\s*#|bill\s*of\s*lading|^bl\b/i, 39), cNote = f(/^notes?$/i, 40);
    var cRef = f(/^delivery\s*ref/i, 0), cPO = headers[2], cItem = headers[3];
    var cDesc = f(/item\s*description/i, 4), cQty = f(/^quantity/i, 5), cUoM = f(/^uom/i, 6), cCoo = f(/country\s*of\s*origin/i, 13),
        cHs = f(/^hs\s*code/i, 14), cCollo = f(/^collo/i, 17), cGross = f(/gross\s*weight/i, 23), cVol = f(/^volume/i, 22),
        cSup = f(/name\s*of\s*supplier|supplier\s*name/i, 10), cLoc = f(/^location/i, 26);

    var map = {}, order = [];
    data.forEach(function (r) {
      var k = trim(r[cShip]);
      if (!k) return;
      if (!map[k]) { map[k] = { no: k, rows: [] }; order.push(k); }
      map[k].rows.push(r);
    });

    var today = new Date();
    var out = order.map(function (k) {
      var s = map[k], rows = s.rows;
      var first = function (c) { for (var i = 0; i < rows.length; i++) { var v = trim(rows[i][c]); if (v) return v; } return ''; };
      var uniq = function (c, fn) { var seen = {}, o = []; rows.forEach(function (r) { var v = trim(r[c]); if (!v) return; if (fn) v = fn(v); if (!seen[v]) { seen[v] = 1; o.push(v); } }); return o; };

      var contRaw = uniq(cCont), conts = [], others = [];
      contRaw.forEach(function (v) { (isContainerNo(v) ? conts : others).push(isContainerNo(v) ? clean(v) : v); });
      conts = conts.filter(function (v, i) { return conts.indexOf(v) === i; });
      var bls = uniq(cBl, function (v) { return v.replace(/\s+/g, ' '); });

      var blClean = bls.map(clean), contAll = contRaw.join(' ');
      var isAir = /\bair|airfreight/i.test(contAll) || blClean.some(function (b) { return /^(AWB|DHL)/.test(b) || /^\d{3}\d{8}$/.test(b); })
                  || /hanoi/i.test(first(cDest));
      var mode = isAir ? 'air' : 'sea';
      var dest = destKey(first(cDest));
      var destAssumed = false;
      if (!dest) { dest = isAir ? 'HAN' : DEFAULT_DEST; destAssumed = true; }
      if (isAir && dest === 'HPH') dest = 'HAN';

      // rederij: B/L-voorvoegsel eerst, dan containernummer
      var carrier = null, carrierKey = null;
      if (window.CarrierTrack && !isAir) {
        var cand = blClean.concat(conts.map(clean));
        for (var i = 0; i < cand.length && !carrierKey; i++) carrierKey = CarrierTrack.detect(cand[i]);
        if (carrierKey) carrier = CarrierTrack.carriers[carrierKey].name;
      }
      var forwarder = blClean.some(function (b) { return /^DSV/.test(b); }) ? 'DSV (forwarder)' : '';

      var etd = pDate(first(cEtd)), eta = pDate(first(cEta)), load = pDate(first(cLoad));
      var etaTxt = first(cEta);
      var etdTxt = first(cEtd);

      var booked = bls.length > 0 || conts.length > 0 || others.length > 0;
      var status;
      if (!booked && !etd) status = 'plan';
      else if (eta && today >= eta) status = 'arrived';
      else if (etd && today >= etd) status = 'transit';
      else status = 'booked';
      var pct = 0;
      if (status === 'arrived') pct = 100;
      else if (status === 'transit' && etd && eta && eta > etd) pct = Math.max(1, Math.min(99, Math.round((today - etd) / (eta - etd) * 100)));
      else if (status === 'transit') pct = null;

      var colloSet = {}, gross = 0, vol = 0, pos = {}, refs = {}, sups = {};
      rows.forEach(function (r) {
        var c = trim(r[cCollo]); if (c) colloSet[c] = 1;
        gross += numPlain(r[cGross]); vol += numPlain(r[cVol]);
        var p = trim(r[cPO]); if (p) pos[p] = 1;
        var rf = trim(r[cRef]); if (rf) refs[rf] = 1;
        var su = trim(r[cSup]); if (su) sups[su] = 1;
      });

      return {
        no: k, rows: rows, mode: mode, dest: dest, destAssumed: destAssumed, conts: conts, others: others, bls: bls,
        carrier: carrier, carrierKey: carrierKey, forwarder: forwarder,
        etd: etd, eta: eta, load: load, etdTxt: etdTxt, etaTxt: etaTxt, status: status, pct: pct,
        note: first(cNote), refs: Object.keys(refs), pos: Object.keys(pos), sups: Object.keys(sups),
        collo: Object.keys(colloSet).length, gross: gross, vol: vol, n: rows.length,
        cols: { ref: cRef, po: cPO, item: cItem, desc: cDesc, qty: cQty, uom: cUoM, coo: cCoo, hs: cHs, collo: cCollo, loc: cLoc, sup: cSup }
      };
    });
    return out;
  }

  /* ── geometrie ──────────────────────────────────────────────────────────── */
  function hav(a, b) {
    var R = 6371, t = function (x) { return x * Math.PI / 180; }, dl = t(b[0] - a[0]), dn = t(b[1] - a[1]);
    var q = Math.sin(dl / 2) * Math.sin(dl / 2) + Math.cos(t(a[0])) * Math.cos(t(b[0])) * Math.sin(dn / 2) * Math.sin(dn / 2);
    return 2 * R * Math.asin(Math.sqrt(q));
  }
  /** hoek (graden, kloksgewijs, 0 = oost) van a naar b zoals die op de kaart (Mercator) getekend wordt */
  function screenAngle(a, b) {
    var pa = L.CRS.EPSG3857.project(L.latLng(a[0], a[1])), pb = L.CRS.EPSG3857.project(L.latLng(b[0], b[1]));
    return Math.atan2(-(pb.y - pa.y), pb.x - pa.x) * 180 / Math.PI;
  }
  function along(pts, f) {
    var seg = [], tot = 0, i;
    for (i = 1; i < pts.length; i++) { var d = hav(pts[i - 1], pts[i]); seg.push(d); tot += d; }
    var want = tot * f; i = 0;
    while (i < seg.length - 1 && want > seg[i]) { want -= seg[i]; i++; }
    var t = seg[i] ? want / seg[i] : 0, a = pts[i], b = pts[i + 1];
    var pos = [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
    return { pos: pos, done: pts.slice(0, i + 1).concat([pos]), todo: [pos].concat(pts.slice(i + 1)), bearing: screenAngle(a, b) };
  }

  /* ── UI-status ──────────────────────────────────────────────────────────── */
  var ST = {
    transit: { l: 'Onderweg', c: 'trk-s-transit', r: 1 },
    booked:  { l: 'Geboekt', c: 'trk-s-booked', r: 2 },
    plan:    { l: 'Gebundeld', c: 'trk-s-plan', r: 3 },
    arrived: { l: 'Aangekomen*', c: 'trk-s-arrived', r: 4 }
  };
  var S = { ships: [], src: null, sel: null, filter: 'active', q: '', tab: 'events', map: null, layer: null };

  /* ── kaart ──────────────────────────────────────────────────────────────── */
  function ensureMap() {
    if (S.map || !window.L) return;
    var el = document.getElementById('trk-map'); if (!el) return;
    S.map = L.map(el, { minZoom: 1, maxZoom: 8, zoomSnap: .5, worldCopyJump: false, attributionControl: true }).setView([30, 60], 2);
    L.tileLayer('https://tile.openstreetmap.org/{z}/{x}/{y}.png', { maxZoom: 19, attribution: '© OpenStreetMap' }).addTo(S.map);
    S.layer = L.layerGroup().addTo(S.map);
  }
  // Beide iconen wijzen standaard naar rechts (oost) en worden gedraaid naar de vaarrichting.
  function icon(sh, deg) {
    if (sh.mode === 'air') return L.divIcon({ className: '', html: '<svg class="trk-plane" width="30" height="30" viewBox="-15 -15 30 30" style="transform:rotate(' + deg + 'deg)"><path d="M14 0 L4 -2 L-2 -12 L-5 -12 L-3 -2 L-9 -1 L-11 -5 L-13 -5 L-12 0 L-13 5 L-11 5 L-9 1 L-3 2 L-5 12 L-2 12 L4 2 Z" fill="#f2a900" stroke="#fff" stroke-width="1.2" stroke-linejoin="round"/></svg>', iconSize: [30, 30], iconAnchor: [15, 15] });
    return L.divIcon({ className: '', html: '<svg class="trk-ship" width="30" height="30" viewBox="-15 -15 30 30" style="transform:rotate(' + deg + 'deg)"><path d="M12 0 L-8 -7 L-4 0 L-8 7 Z" fill="#f2a900" stroke="#fff" stroke-width="1.6" stroke-linejoin="round"/></svg>', iconSize: [30, 30], iconAnchor: [15, 15] });
  }
  function drawMap(fit) {
    ensureMap(); if (!S.map) return;
    S.layer.clearLayers();
    var sh = S.ships.filter(function (x) { return x.no === S.sel; })[0]; if (!sh) return;
    var route = sh.mode === 'air' ? airRoute(sh.dest) : seaRoute(sh.dest);
    var from = sh.mode === 'air' ? 'AMS' : ORIGIN;
    var f = sh.status === 'arrived' ? 1 : sh.status === 'transit' && sh.pct != null ? sh.pct / 100 : 0;
    var a = along(route, Math.min(Math.max(f, 0.0001), 0.9999));
    if (f > 0) L.polyline(a.done, { color: '#00B4C8', weight: 3 }).addTo(S.layer);
    if (f < 1) L.polyline(f > 0 ? a.todo : route, { color: '#8FA3BF', weight: 2, dashArray: '2 7' }).addTo(S.layer);
    [from, sh.dest].forEach(function (p) {
      L.circleMarker(PORTS[p].ll, { radius: 6, color: '#0A1628', weight: 2, fillColor: '#fff', fillOpacity: 1 }).bindTooltip(PORTS[p].n + ', ' + PORTS[p].c).addTo(S.layer);
    });
    var pos = f >= 1 ? PORTS[sh.dest].ll : f <= 0 ? PORTS[from].ll : a.pos;
    var m = L.marker(pos, { icon: icon(sh, a.bearing) }).addTo(S.layer);
    m.bindPopup('<b>Shipment ' + E(sh.no) + '</b><br>' + E(PORTS[from].n) + ' → ' + E(PORTS[sh.dest].n) + '<br>ETA <b>' + E(fmt(sh.eta) !== '—' ? fmt(sh.eta) : (sh.etaTxt || '—')) + '</b><br>' + ST[sh.status].l + (sh.mode !== 'air' ? '<br><span style="color:#6b7a90">Route om Afrika (Kaap de Goede Hoop)</span>' : '') + (sh.pct != null && sh.status === 'transit' ? ' · ' + sh.pct + '% (geschat)' : ''));
    if (fit) S.map.fitBounds(L.latLngBounds(route).pad(.25), { maxZoom: 4 });
  }

  /* ── lijst ──────────────────────────────────────────────────────────────── */
  function visible() {
    var q = S.q.toLowerCase();
    return S.ships.filter(function (s) {
      if (S.filter === 'active' && s.status === 'arrived') return false;
      if (S.filter !== 'active' && S.filter !== 'all' && s.status !== S.filter) return false;
      if (q) {
        var hay = [s.no, s.bls.join(' '), s.conts.join(' '), s.refs.join(' '), s.pos.join(' '), s.sups.join(' ')].join(' ').toLowerCase();
        if (hay.indexOf(q) < 0) return false;
      }
      return true;
    }).sort(function (a, b) {
      if (ST[a.status].r !== ST[b.status].r) return ST[a.status].r - ST[b.status].r;
      var ea = a.eta ? +a.eta : 9e15, eb = b.eta ? +b.eta : 9e15;
      return a.status === 'arrived' ? eb - ea : ea - eb;
    });
  }
  function renderKpis() {
    var c = { transit: 0, booked: 0, plan: 0, arrived: 0 };
    S.ships.forEach(function (s) { c[s.status]++; });
    var el = document.getElementById('trk-kpis'); if (!el) return;
    el.innerHTML = [['transit', 'Onderweg'], ['booked', 'Geboekt'], ['plan', 'Gebundeld (niet geboekt)'], ['arrived', 'Aangekomen*']].map(function (x) {
      return '<div class="trk-kpi ' + ST[x[0]].c + '"><b>' + c[x[0]] + '</b><span>' + x[1] + '</span></div>';
    }).join('');
  }
  function renderFilters() {
    var el = document.getElementById('trk-filters'); if (!el) return;
    var f = [['active', 'Actief'], ['transit', 'Onderweg'], ['booked', 'Geboekt'], ['plan', 'Gebundeld'], ['arrived', 'Aangekomen*'], ['all', 'Alles']];
    el.innerHTML = f.map(function (x) { return '<button class="trk-chip' + (S.filter === x[0] ? ' on' : '') + '" data-f="' + x[0] + '">' + x[1] + '</button>'; }).join('');
    el.querySelectorAll('button').forEach(function (b) { b.addEventListener('click', function () { S.filter = b.dataset.f; renderFilters(); renderList(); }); });
  }
  function renderList() {
    var tb = document.getElementById('trk-tbody'); if (!tb) return;
    var list = visible();
    if (!list.length) { tb.innerHTML = '<tr><td colspan="4" class="trk-empty">Geen zendingen in dit filter</td></tr>'; return; }
    tb.innerHTML = list.map(function (s) {
      var st = ST[s.status];
      var pctTxt = s.status === 'transit' && s.pct != null ? s.pct + '%' : s.status === 'arrived' ? '100%' : '';
      return '<tr class="trk-row' + (s.no === S.sel ? ' sel' : '') + '" tabindex="0" data-no="' + E(s.no) + '">' +
        '<td><span class="trk-no">#' + E(s.no) + '</span><span class="trk-sub">' + (s.mode === 'air' ? '✈ ' : '🚢 ') + E(PORTS[s.dest].n) + (s.destAssumed ? ' (aangenomen)' : '') + '</span>' +
        '<span class="trk-sub trk-mono">' + E(s.bls[0] || s.conts[0] || '—') + (s.conts.length > 1 ? ' +' + (s.conts.length - 1) : '') + '</span></td>' +
        '<td class="trk-mono">' + E(s.eta ? fmtS(s.eta) : (s.etaTxt || '—')) + '</td>' +
        '<td><span class="trk-pill ' + st.c + '"><i></i>' + st.l + '</span><div class="trk-bar"><span style="width:' + (s.pct || 0) + '%"></span></div><span class="trk-sub trk-mono">' + pctTxt + '</span></td>' +
        '<td class="trk-mono trk-r">' + s.collo + '<span class="trk-sub">collo</span></td></tr>';
    }).join('');
    tb.querySelectorAll('tr.trk-row').forEach(function (tr) {
      tr.addEventListener('click', function () { select(tr.dataset.no, true); });
      tr.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); select(tr.dataset.no, true); } });
    });
  }

  /* ── detail ─────────────────────────────────────────────────────────────── */
  function vesselKey(no) { return 'ihc-trk-vessel-' + no; }
  function renderDetail() {
    var el = document.getElementById('trk-detail'); if (!el) return;
    var s = S.ships.filter(function (x) { return x.no === S.sel; })[0];
    if (!s) { el.innerHTML = '<div class="trk-empty">Kies een zending in de lijst.</div>'; return; }
    var st = ST[s.status], from = s.mode === 'air' ? 'AMS' : ORIGIN;
    var mf = MF(), cipl = (mf && mf.ciplMeta && mf.ciplMeta.vessel) || '';
    var vessel = store(vesselKey(s.no)) || '';
    var etaShow = s.eta ? fmt(s.eta) : (s.etaTxt || '—');
    var rest = s.status === 'arrived' ? 'ETA verstreken' : s.eta ? Math.max(0, Math.round((s.eta - new Date()) / DAY)) + ' dagen' : '—';
    var pct = s.pct == null ? 0 : s.pct;

    // events
    var now = new Date(), ev = [];
    ev.push({ d: s.load, t: 'Gebundeld en geladen (Date load)', sub: s.refs.length + ' delivery refs · ' + s.collo + ' collo', done: !!s.load && s.load <= now, nod: !s.load });
    if (s.bls.length || s.conts.length) ev.push({ d: null, t: 'Boeking bij rederij', sub: (s.bls[0] ? 'B/L ' + s.bls[0] : '') + (s.conts.length ? ' · ' + s.conts.length + ' container(s)' : ''), done: true, nod: true });
    ev.push({ d: s.etd, t: 'Vertrek ' + PORTS[from].n + ' (ETD)', sub: s.etd ? '' : (s.etdTxt || 'ETD nog niet bekend'), done: !!s.etd && s.etd <= now });
    ev.push({ d: s.eta, t: 'Aankomst ' + PORTS[s.dest].n + ' (ETA)', sub: s.eta ? (s.status === 'arrived' ? 'ETA verstreken, geen bevestiging uit tracking' : '') : (s.etaTxt || 'ETA nog niet bekend'), done: s.status === 'arrived' });
    var evHtml = '<ul class="trk-ev">' + ev.map(function (e) {
      return '<li class="' + (e.done ? 'done' : 'fut') + '"><span class="trk-evd">' + (e.d ? fmt(e.d) : '—') + '</span><span class="trk-evk"></span><div><b>' + E(e.t) + '</b>' + (e.sub ? '<span class="trk-sub">' + E(e.sub) + '</span>' : '') + '</div></li>';
    }).join('') + '</ul>';

    // items
    var c = s.cols, g = function (r, k) { return k ? trim(r[k]) : ''; };
    var itemRows = s.rows.slice(0, 300).map(function (r) {
      return '<tr><td class="trk-mono">' + E(g(r, c.ref)) + '</td><td class="trk-mono">' + E(g(r, c.po) + g(r, c.item)) + '</td><td class="trk-wrap">' + E(g(r, c.desc)) + '</td><td class="trk-r">' + E(g(r, c.qty)) + '</td><td>' + E(g(r, c.coo)) + '</td><td class="trk-mono">' + E(g(r, c.hs)) + '</td><td class="trk-r">' + E(g(r, c.collo)) + '</td></tr>';
    }).join('');
    var itemsHtml = '<div class="trk-scroll"><table class="trk-t"><thead><tr><th>Delivery ref</th><th>PO / Item</th><th>Omschrijving</th><th class="trk-r">Qty</th><th>COO</th><th>HS-code</th><th class="trk-r">Collo</th></tr></thead><tbody>' + itemRows + '</tbody></table></div>' +
      (s.rows.length > 300 ? '<p class="trk-note">Eerste 300 van ' + s.rows.length + ' regels.</p>' : '');

    // containers / BL
    var canTrack = !!window.CarrierTrack;
    var btn = function (kind, v) { return canTrack ? '<button class="trk-track" data-k="' + kind + '" data-v="' + E(v) + '">Track &amp; Trace ↗</button>' : ''; };
    var cRows = s.bls.map(function (b) { return '<tr><td>B/L</td><td class="trk-mono">' + E(b) + '</td><td>' + E(s.carrier || s.forwarder || (s.mode === 'air' ? 'Luchtvracht' : 'Onbekend')) + '</td><td>' + (s.mode === 'air' ? '' : btn('bl', b)) + '</td></tr>'; }).join('') +
      s.conts.map(function (v) { return '<tr><td>Container</td><td class="trk-mono">' + E(v) + '</td><td>' + E((canTrack && CarrierTrack.detect(v) && CarrierTrack.carriers[CarrierTrack.detect(v)].name) || '—') + '</td><td>' + btn('container', v) + '</td></tr>'; }).join('') +
      s.others.map(function (v) { return '<tr><td>Overig</td><td class="trk-mono">' + E(v) + '</td><td colspan="2">geen geldig containernummer</td></tr>'; }).join('');
    var contHtml = (cRows ? '<div class="trk-scroll"><table class="trk-t"><thead><tr><th>Type</th><th>Nummer</th><th>Rederij</th><th></th></tr></thead><tbody>' + cRows + '</tbody></table></div>' : '<p class="trk-note">Nog geen B/L of container ingevuld in de Moederlijst (kolom AM / AN).</p>') +
      '<p class="trk-note">Schipnaam en echte positie komen uit de trackingpagina van de rederij. Noteer het schip hieronder; dat blijft op deze computer bewaard.</p>';

    var body = S.tab === 'events' ? evHtml : S.tab === 'items' ? itemsHtml : contHtml;
    var tabs = [['events', 'Events'], ['items', 'Items · COO · HS (' + s.n + ')'], ['conts', 'Container / B/L']];

    el.innerHTML =
      '<div class="trk-dhead"><div><h3>Shipment #' + E(s.no) + '</h3><div class="trk-sub">' + E(s.carrier || s.forwarder || (s.mode === 'air' ? 'Luchtvracht' : 'Rederij onbekend')) + ' · ' + s.refs.length + ' delivery refs · ' + s.pos.length + ' POs' + (s.note ? ' · ' + E(s.note) : '') + '</div></div>' +
      '<span class="trk-pill ' + st.c + '"><i></i>' + st.l + (s.status === 'transit' && s.pct != null ? ' · ' + s.pct + '% (geschat)' : '') + '</span></div>' +
      '<dl class="trk-facts"><div><dt>Vertrek (ETD)</dt><dd>' + E(s.etd ? fmt(s.etd) : (s.etdTxt || '—')) + '</dd></div><div><dt>Aankomst (ETA)</dt><dd>' + E(etaShow) + '</dd></div><div><dt>Resterend</dt><dd>' + E(rest) + '</dd></div>' +
      '<div><dt>Bruto / volume</dt><dd>' + Math.round(s.gross).toLocaleString('nl-NL') + ' kg' + (s.vol > 0 ? ' · ' + s.vol.toFixed(1) + ' m³' : '') + '</dd></div>' +
      '<div><dt>Schip</dt><dd><input id="trk-vessel" class="trk-in" placeholder="' + E(cipl || 'bv. METTE MAERSK') + '" value="' + E(vessel) + '"></dd></div></dl>' +
      '<div class="trk-strip"><div class="trk-line"></div><div class="trk-fill" style="width:calc((100% - 16px)*' + (pct / 100) + ')"></div><span class="trk-dot" style="left:8px"></span><span class="trk-dot' + (s.status === 'arrived' ? '' : ' todo') + '" style="left:calc(100% - 8px)"></span><span class="trk-boat" style="left:calc(8px + (100% - 16px)*' + (pct / 100) + ')">' + (s.mode === 'air' ? '✈' : '🚢') + '</span></div>' +
      '<div class="trk-strip-l"><span><b>' + E(PORTS[from].n) + '</b></span><span><b>' + E(PORTS[s.dest].n) + '</b>' + (s.destAssumed ? ' <span class="trk-sub">(aangenomen)</span>' : '') + '</span></div>' +
      '<div class="trk-tabs" role="tablist">' + tabs.map(function (t) { return '<button class="trk-tab" role="tab" aria-selected="' + (S.tab === t[0]) + '" data-t="' + t[0] + '">' + t[1] + '</button>'; }).join('') + '</div>' +
      '<div class="trk-body">' + body + '</div>';

    el.querySelectorAll('.trk-tab').forEach(function (b) { b.addEventListener('click', function () { S.tab = b.dataset.t; renderDetail(); }); });
    el.querySelectorAll('.trk-track').forEach(function (b) {
      b.addEventListener('click', function () { var o = {}; o[b.dataset.k === 'bl' ? 'bl' : 'container'] = b.dataset.v; CarrierTrack.open(o); });
    });
    var vi = document.getElementById('trk-vessel');
    if (vi) vi.addEventListener('change', function () { store(vesselKey(s.no), vi.value.trim()); });
  }

  function select(no, fit) {
    S.sel = no; renderList(); renderDetail();
    ensureMap(); drawMap(fit !== false);
    if (S.map) setTimeout(function () { S.map.invalidateSize(); }, 30);
  }

  /* ── publiek ────────────────────────────────────────────────────────────── */
  function show() {
    var wrap = document.getElementById('track-wrap'); if (!wrap) return;
    var mf0 = MF(), data = mf0 && mf0.data;
    var msg = document.getElementById('trk-msg');
    if (!data) { if (msg) msg.textContent = 'Laad eerst de Moederlijst.'; return; }
    if (S.src !== data) {
      S.src = data; S.ships = buildShipments(data); S.sel = null;
      S.filter = 'active'; S.q = '';
    }
    if (msg) msg.textContent = S.ships.length + ' shipments uit kolom "Shipment" · status afgeleid uit ETD/ETA van vandaag';
    renderKpis(); renderFilters();
    if (!S.sel) {
      var l = visible(); S.sel = l.length ? l[0].no : (S.ships[0] && S.ships[0].no);
    }
    ensureMap();
    renderList(); renderDetail(); drawMap(true);
    if (S.map) setTimeout(function () { S.map.invalidateSize(); drawMap(true); }, 60);
  }
  function search(v) { S.q = String(v || ''); renderList(); }

  window.ShipTrack = { show: show, search: search, build: buildShipments, _state: S };
})();
