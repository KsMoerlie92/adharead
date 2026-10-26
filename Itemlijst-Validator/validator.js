// ── Itemlijst Validator ───────────────────────────────────────────────────
// Validates supplier-filled Itemlijst Excel against Royal IHC rules.
// Writes back IHC-fillable fields and exports corrected file.

// ── Master validation lists (from Itemlijst Master tab) ───────────────────
const VL_UOM  = new Set(['Piece(s)','Bucket(s)','Meter','Set(s)','Kilogram']);
const VL_PKG  = new Set(['Pallet','Case','Crate','Carton','Skid','Loose','Reel','Bundle','Bag']);
// Inspectieniveaus — keuzelijst (dropdown) in de validatortabel
const INSP_OPTIONS = ["Foto's en Steekproef", "Foto's", "Fysiek controleren", "TBD", "Geen Controle", "Volledige Controle"];
const VL_INSP = new Set(INSP_OPTIONS);
const VL_INSP_LC = new Set(INSP_OPTIONS.map(s => s.toLowerCase())); // tolerant matchen op casing
// 250 ISO-2 country codes (abbreviated — full set checked at runtime from file)
// Volledige ISO 3166-1 alpha-2 landcodelijst — gebruikt als vangnet wanneer
// het bestand geen Master-tabblad heeft (of dat tabblad geen landcodes bevat).
// Eerder stond hier een lijst van maar 30 landen, waardoor volkomen geldige
// codes (bv. 'MY' voor Maleisië) onterecht als fout werden gemeld zodra er
// geen Master-tabblad aanwezig was.
const VL_COO_FALLBACK = new Set(['AD','AE','AF','AG','AI','AL','AM','AO','AQ','AR','AS','AT','AU','AW','AX','AZ',
  'BA','BB','BD','BE','BF','BG','BH','BI','BJ','BL','BM','BN','BO','BQ','BR','BS','BT','BV','BW','BY','BZ',
  'CA','CC','CD','CF','CG','CH','CI','CK','CL','CM','CN','CO','CR','CU','CV','CW','CX','CY','CZ',
  'DE','DJ','DK','DM','DO','DZ',
  'EC','EE','EG','EH','ER','ES','ET','EU',
  'FI','FJ','FK','FM','FO','FR',
  'GA','GB','GD','GE','GF','GG','GH','GI','GL','GM','GN','GP','GQ','GR','GS','GT','GU','GW','GY',
  'HK','HM','HN','HR','HT','HU',
  'ID','IE','IL','IM','IN','IO','IQ','IR','IS','IT',
  'JE','JM','JO','JP',
  'KE','KG','KH','KI','KM','KN','KP','KR','KW','KY','KZ',
  'LA','LB','LC','LI','LK','LR','LS','LT','LU','LV','LY',
  'MA','MC','MD','ME','MF','MG','MH','MK','ML','MM','MN','MO','MP','MQ','MR','MS','MT','MU','MV','MW','MX','MY','MZ',
  'NA','NC','NE','NF','NG','NI','NL','NO','NP','NR','NU','NZ',
  'OM',
  'PA','PE','PF','PG','PH','PK','PL','PM','PN','PR','PS','PT','PW','PY',
  'QA',
  'RE','RO','RS','RU','RW',
  'SA','SB','SC','SD','SE','SG','SH','SI','SJ','SK','SL','SM','SN','SO','SR','SS','ST','SV','SX','SY','SZ',
  'TC','TD','TF','TG','TH','TJ','TK','TL','TM','TN','TO','TR','TT','TV','TW','TZ',
  'UA','UG','UM','US','UY','UZ',
  'VA','VC','VE','VG','VI','VN','VU',
  'WF','WS',
  'YE','YT',
  'ZA','ZM','ZW']);

// ── Landnaam / alpha-3 → ISO 3166-1 alpha-2 code ────────────────────────────
// Zet volledige landnamen (NL + EN) en 3-letter codes om naar de 2-letter code.
const _COUNTRY_TO_ISO2 = (function () {
  const m = {};
  const add = (code, ...names) => { names.forEach(n => { m[n.toLowerCase()] = code; }); m[code.toLowerCase()] = code; };
  add('NL','nederland','netherlands','the netherlands','holland','nld');
  add('DE','duitsland','germany','deutschland','deu','ger');
  add('BE','belgie','belgië','belgium','belgique','bel');
  add('FR','frankrijk','france','fra');
  add('GB','verenigd koninkrijk','united kingdom','great britain','groot-brittannie','groot-brittannië','england','engeland','uk','gbr');
  add('IE','ierland','ireland','irl');
  add('IT','italie','italië','italy','italia','ita');
  add('ES','spanje','spain','espana','españa','esp');
  add('PT','portugal','prt');
  add('LU','luxemburg','luxembourg','lux');
  add('CH','zwitserland','switzerland','suisse','schweiz','che');
  add('AT','oostenrijk','austria','osterreich','österreich','aut');
  add('DK','denemarken','denmark','dnk');
  add('SE','zweden','sweden','sverige','swe');
  add('NO','noorwegen','norway','norge','nor');
  add('FI','finland','suomi','fin');
  add('IS','ijsland','iceland','isl');
  add('PL','polen','poland','polska','pol');
  add('CZ','tsjechie','tsjechië','czech republic','czechia','czech','cze');
  add('SK','slowakije','slovakia','svk');
  add('HU','hongarije','hungary','hun');
  add('RO','roemenie','roemenië','romania','rou');
  add('BG','bulgarije','bulgaria','bgr');
  add('GR','griekenland','greece','grc');
  add('HR','kroatie','kroatië','croatia','hrv');
  add('SI','slovenie','slovenië','slovenia','svn');
  add('RS','servie','servië','serbia','srb');
  add('EE','estland','estonia','est');
  add('LV','letland','latvia','lva');
  add('LT','litouwen','lithuania','ltu');
  add('UA','oekraine','oekraïne','ukraine','ukr');
  add('RU','rusland','russia','russian federation','rus');
  add('BY','belarus','wit-rusland','wit rusland','blr');
  add('TR','turkije','turkey','turkiye','türkiye','tur');
  add('US','verenigde staten','united states','united states of america','usa','amerika','america','u.s.a','u.s','united states of america (usa)');
  add('CA','canada','can');
  add('MX','mexico','mex');
  add('BR','brazilie','brazilië','brazil','bra');
  add('AR','argentinie','argentinië','argentina','arg');
  add('CL','chili','chile','chl');
  add('CN','china','volksrepubliek china','p.r. china','prc','chn');
  add('HK','hongkong','hong kong','hkg');
  add('TW','taiwan','twn');
  add('JP','japan','jpn');
  add('KR','zuid-korea','zuid korea','south korea','korea','republic of korea','korea, republic of','kor');
  add('KP','noord-korea','north korea','prk');
  add('IN','india','ind');
  add('PK','pakistan','pak');
  add('BD','bangladesh','bgd');
  add('VN','vietnam','viet nam','vnm');
  add('TH','thailand','tha');
  add('MY','maleisie','maleisië','malaysia','mys');
  add('SG','singapore','sgp');
  add('ID','indonesie','indonesië','indonesia','idn');
  add('PH','filipijnen','philippines','phl');
  add('AE','verenigde arabische emiraten','united arab emirates','uae','emirates','are');
  add('SA','saoedi-arabie','saoedi-arabië','saudi arabia','saudi-arabia','sau');
  add('QA','qatar','qat');
  add('KW','koeweit','kuwait','kwt');
  add('BH','bahrein','bahrain','bhr');
  add('OM','oman','omn');
  add('IL','israel','israël','isr');
  add('EG','egypte','egypt','egy');
  add('ZA','zuid-afrika','zuid afrika','south africa','zaf');
  add('NG','nigeria','nga');
  add('MA','marokko','morocco','mar');
  add('AU','australie','australië','australia','aus');
  add('NZ','nieuw-zeeland','nieuw zeeland','new zealand','nzl');
  return m;
})();

// Geeft de ISO 3166-1 alpha-2 code terug, of null als onbekend (dan onveranderd laten).
function _toCountryCode(v) {
  const s = String(v ?? '').trim();
  if (!s) return null;
  const key = s.toLowerCase().replace(/\.+$/, '').replace(/\s+/g, ' ').trim();
  if (_COUNTRY_TO_ISO2[key]) return _COUNTRY_TO_ISO2[key];
  if (/^[a-z]{2}$/.test(key)) return key.toUpperCase();   // al een 2-letter code
  return null;
}

// Rood-markering voor lege/onjuiste cellen + markering op slim ingevulde headers
(function () {
  if (typeof document === 'undefined' || document.getElementById('val-extra-style')) return;
  const s = document.createElement('style'); s.id = 'val-extra-style';
  s.textContent =
    '.val-cell-err{background:rgba(201,100,66,.30)!important}' +
    '.val-cell-err,.val-cell-err .val-input{color:#ffe0d4!important}' +
    '.val-cell-err .val-input{border-bottom-color:rgba(224,122,82,.6)!important}' +
    '.val-cell-err .val-input:focus{border-bottom-color:#e07a52!important}' +
    '.val-fill-mark{margin-left:4px;font-size:.72em;opacity:.9;cursor:help;vertical-align:middle}';
  (document.head || document.documentElement).appendChild(s);
})();

// ── Column index map (0-based, row 2 = headers) ───────────────────────────
// A–Y volgen de vaste positie in de officiële Itemlijst-template (zie
// kolomspecificatie). Kolom Z is in de template niet in gebruik (leeg) en
// AD idem — beide indices blijven als niet-semantische plek gereserveerd.
// AA t/m AH worden hieronder in _remapColumns() alsnog op naam herkend
// (positie-onafhankelijk), met een virtuele kolom als vangnet wanneer een
// veld niet in het bestand voorkomt.
let COL = {
  A:0, B:1, C:2, D:3, E:4, F:5, G:6, H:7, I:8, J:9,
  K:10,L:11,M:12,N:13,O:14,P:15,Q:16,R:17,S:18,T:19,
  U:20,V:21,W:22,X:23,Y:24,Z:25,
  AA:26,AB:27,AC:28,AD:29,AE:30,AF:31,AG:32,AH:33
};

// State
let _valRows     = [];   // parsed data rows [{cells:[], errors:{}, warnings:{}}]
let _valHeaders  = [];   // header row (detected dynamically)
let _valOwners   = [];   // owner row (row before headers, if present)
let _valCOO      = new Set(); // country codes from Master tab
let _usdRate     = null; // cached EUR/USD rate
let _valWb       = null; // original workbook for write-back
let _virtualCols = {};   // { Z: true, AA: true } when columns are IHC-added (not in file)
let _valHdrIdx  = 1;     // row index (0-based) where headers were found

// ── Remap COL indices from actual header names ──────────────────────────────
// Columns found by name keep their real index.
// Columns missing from the file get virtual indices beyond the last column.
function _remapColumns() {
  // Reset to defaults first
  COL = {
    A:0,  B:1,  C:2,  D:3,  E:4,  F:5,  G:6,  H:7,  I:8,  J:9,
    K:10, L:11, M:12, N:13, O:14, P:15, Q:16, R:17, S:18, T:19,
    U:20, V:21, W:22, X:23, Y:24, Z:25,
    AA:26, AB:27, AC:28, AD:29, AE:30, AF:31, AG:32, AH:33
  };
  _virtualCols = {};

  const hdrs = _valHeaders;
  const find = (pattern) => hdrs.findIndex(h => h && new RegExp(pattern,'i').test(String(h)));

  // Virtuele kolommen altijd ná de hoogste bekende kolompositie laten beginnen
  // (nooit enkel na hdrs.length) — anders kan een kortere/niet-standaard
  // aangeleverde itemlijst (minder kolommen dan de volledige template) een
  // virtuele kolom (bv. Dangerous Goods) laten samenvallen met een vaste
  // positie (bv. Height) en zo stilzwijgend data overschrijven.
  let nextVirtual = Math.max(hdrs.length, ...Object.values(COL)) + 1;

  // Re-detect a field by header name so it works regardless of its position
  // in the actual file; fall back to a virtual (IHC-added) column when the
  // field isn't present at all.
  const remapOne = (key, pattern, label) => {
    const idx = find(pattern);
    if (idx >= 0) {
      COL[key] = idx;
    } else {
      COL[key] = nextVirtual++;
      _virtualCols[key] = true;
      _valHeaders[COL[key]] = label;
    }
  };

  remapOne('AA', 'dangerous',  'Dangerous Goods');   // AA: Dangerous Goods — tickbox
  remapOne('AB', 'dual.?use',  'Dual use');           // AB: Dual use — tickbox
  remapOne('AC', 'eccn',       'ECCN code');          // AC: ECCN code — verplicht als AB aangevinkt
  remapOne('AE', 'container',  'Container number');   // AE: Container number
  remapOne('AF', 'seal',       'Seal number');         // AF: Seal number
  remapOne('AG', 'remarks',    'Remarks supplier');   // AG: Remarks supplier
  remapOne('AH', 'inspection', 'Inspection Level');   // AH: Inspection level — door IHC in te vullen
}

// ── Ja/Nee-waarden (YES/NO, JA/NEE, TRUE/FALSE, 1/0) → echte boolean ─────────
// Leveranciers vullen Dangerous Goods / Dual use soms in als YES/NO i.p.v. TRUE/FALSE.
function _valBool(v) {
  if (v === true || v === 1) return true;
  const t = String(v ?? '').trim().toLowerCase();
  return ['true', 'yes', 'y', 'ja', 'j', '1', 'x', 'waar'].includes(t);
}
function _valNormalizeBoolCols() {
  for (const row of _valRows) {
    for (const k of ['AA', 'AB']) {
      const ci = COL[k], cur = row.cells[ci];
      if (cur === undefined || cur === null || (typeof cur === 'string' && cur.trim() === '')) continue;
      const b = _valBool(cur);
      if (cur !== b) { row.cells[ci] = b; row._edited = true; }
    }
  }
}

// ── Itemlijstregel koppelen aan de Expeditinglijst via PO-nummer (+ regel/release) ──
// Kolom C = Order No, kolom D = "Line-Release" (bv. 1-1). Zo komen de officiële leverancier
// en het Sub Project ID uit Expediting, ook als de leverancier de naam fout spelt of afkort.
let _valPoIdx = null, _valPoIdxFor = null;
function _valPoIndex() {
  const data = (typeof fileData !== 'undefined' && fileData.expediting && Array.isArray(fileData.expediting.data)) ? fileData.expediting.data : [];
  if (_valPoIdx && _valPoIdxFor === data) return _valPoIdx;
  const m = new Map();
  for (const r of data) {
    const ord = String(r['Order No'] ?? '').trim();
    if (!ord) continue;
    if (!m.has(ord)) m.set(ord, []);
    m.get(ord).push(r);
  }
  _valPoIdx = m; _valPoIdxFor = data;
  return m;
}
function _valMatchExpByPo(cells) {
  const idx = _valPoIndex();
  if (!idx.size) return null;
  let ord = String(cells[COL.C] ?? '').replace(/^\s*PO[\s\-_:]*(?=\d)/i, '').trim();
  const mFull = ord.match(/^(\d+)-(\d+)-(\d+)$/);            // volledig PO-nummer "3156010690-1-1"
  let line = null, rel = null;
  if (mFull) { ord = mFull[1]; line = mFull[2]; rel = mFull[3]; }
  const rows = idx.get(ord);
  if (!rows || !rows.length) return null;
  if (line === null) {
    const mi = String(cells[COL.D] ?? '').replace(/^'+/, '').match(/(\d+)\D+(\d+)/);
    if (mi) { line = mi[1]; rel = mi[2]; }
  }
  if (line !== null) {
    const hit = rows.find(r => String(r['Line No'] ?? '').trim() === String(+line) && String(r['Release No'] ?? '').trim() === String(+rel));
    if (hit) return { row: hit, exact: true };
  }
  // Regel niet te bepalen: alleen gebruiken als de hele PO één leverancier en één Sub Project ID heeft
  const sup = new Set(rows.map(r => String(r['Supplier Name'] ?? '').trim()));
  const sp  = new Set(rows.map(r => String(r[_valSubKey] ?? '').trim()));
  return (sup.size === 1 && sp.size === 1) ? { row: rows[0], exact: false } : null;
}

// ── Inspection Level (kolom AH) automatisch invullen vanuit de Admin-regels ─────
// shared/inspection-rules.js (+ inspection-rules.json). Vult ALLEEN lege cellen; wat
// de gebruiker/leverancier al heeft ingevuld blijft staan. Per regel wordt het Sub
// Project ID uit kolom B gebruikt (zelfde normalisatie als de export).
let _valInspRules = null;
async function _valApplyInspectionRules() {
  const R = window.InspectionRules;
  if (!R || typeof R.lookup !== 'function') { console.warn('[validator] InspectionRules niet geladen — shared/inspection-rules.js ontbreekt in de pagina'); return { filled: 0, missing: true }; }
  if (!_valInspRules || !_valInspRules.length) {
    try { _valInspRules = await R.loadRules(); } catch (e) { console.warn('[validator] inspection-rules.json niet te laden:', e); _valInspRules = []; }
  }
  if (!_valInspRules || !_valInspRules.length) return { filled: 0, noRules: true };
  // Veiligheid: alleen vullen als de doelkolom echt de Inspection Level-kolom is
  // (kop met 'inspection') — nooit in een andere kolom schrijven, ook niet bij afwijkende bestanden.
  const _ahHdr = String(_valHeaders[COL.AH] ?? '');
  if (!/inspection/i.test(_ahHdr)) {
    console.warn('[validator] Inspection Level-kolom niet herkend (kop op index ' + COL.AH + ': "' + _ahHdr + '") — niets ingevuld');
    return { filled: 0, badCol: true };
  }
  console.info('[validator] Inspection Level wordt ingevuld in kolom-index ' + COL.AH + ' ("' + _ahHdr + '")');
  let filled = 0, viaExpSupplier = 0, poMatched = 0;
  for (const row of _valRows) {
    const po = _valMatchExpByPo(row.cells);
    const expSup = po ? String(po.row['Supplier Name'] ?? '').trim() : '';
    const expSpid = po ? String(po.row[_valSubKey] ?? '').trim() : '';
    if (po) {
      poMatched++;
      row._expPo = { supplier: expSup, spid: expSpid, exact: po.exact };
      // Project (B) leeg → Sub Project ID uit Expediting
      if (expSpid && String(row.cells[COL.B] ?? '').trim() === '') { row.cells[COL.B] = expSpid; row._edited = true; }
    }
    const cur = row.cells[COL.AH];
    if (cur !== undefined && cur !== null && String(cur).trim() !== '') continue;
    const spid = expSpid || _valNormSubId(row.cells[COL.B]);
    const hRaw = String(row.cells[COL.H] ?? '').trim();
    const codes = (typeof parseHColumn === 'function' ? parseHColumn(hRaw) : []);
    const tryCodes = codes.length ? codes : [hRaw];
    // Leveranciers-kandidaten: eerst de officiële naam uit Expediting (via PO), dan de naam uit de itemlijst
    const sups = [expSup, String(row.cells[COL.K] ?? '').trim()].filter((x, i, a) => x && a.indexOf(x) === i);
    let hit = null, supUsed = '';
    outer:
    for (const sup of (sups.length ? sups : [''])) {
      for (const c of tryCodes) { hit = R.lookup(_valInspRules, spid, String(c).trim(), sup); if (hit) { supUsed = sup; break outer; } }
      hit = R.lookup(_valInspRules, spid, hRaw, sup);
      if (hit) { supUsed = sup; break; }
    }
    if (hit && hit.level) {
      const canon = INSP_OPTIONS.find(o => o.toLowerCase() === String(hit.level).trim().toLowerCase()) || String(hit.level).trim();
      row.cells[COL.AH] = canon; row._edited = true; row._inspRule = hit.rule; filled++;
      if (hit.rule.type === 'supplier' && supUsed && supUsed === expSup && supUsed !== String(row.cells[COL.K] ?? '').trim()) viaExpSupplier++;
    }
  }
  return { filled, viaExpSupplier, poMatched };
}


// ── Hoofdletterongevoelig: dropdown-waarden naar de officiële schrijfwijze ──────
// "piece(s)" → "Piece(s)", "carton " → "Carton", "tr" → "TR", "foto's" → "Foto's".
// Spaties aan de randen worden verwijderd; waarden die niet in de lijst staan blijven
// ongemoeid (die geeft validateRow alsnog als fout).
function _valCanon(list, value) {
  const t = String(value ?? '').replace(/\u00A0/g, ' ').trim().toLowerCase();
  if (!t) return null;
  for (const o of list) if (String(o).toLowerCase() === t) return o;
  // Eenheid: "pcs" (ook pc / pce) betekent Piece(s) — stil omzetten, geen foutmelding
  if (list === VL_UOM && /^(pcs|pc|pce)\.?$/.test(t)) return 'Piece(s)';
  return null;
}
// ── Item (kolom D): vast formaat '-<regel>-<release>, bv. '-1-1, '-1-2 ───────────
// Neemt "'1-1", "1-1", "-1-1" en "'-1-1" over; andere waarden blijven ongemoeid.
function _valItemFmt(v) {
  if (typeof v !== 'string') return v;
  const m = v.replace(/\u00A0/g, ' ').trim().match(/^'?\s*[-\u2013]?\s*(\d+)\s*-\s*(\d+)$/);
  return m ? `'-${+m[1]}-${+m[2]}` : v;
}
function _valNormalizeCase() {
  let n = 0;
  for (const row of _valRows) {
    const fix = (k, fn) => {
      const ci = COL[k], cur = row.cells[ci];
      if (cur === undefined || cur === null || typeof cur !== 'string' || cur.trim() === '') return;
      const nv = fn(cur);
      if (nv !== null && nv !== undefined && nv !== cur) { row.cells[ci] = nv; row._edited = true; n++; }
    };
    fix('G',  v => _valCanon(VL_UOM, v));
    fix('S',  v => _valCanon(VL_PKG, v));
    fix('AH', v => _valCanon(INSP_OPTIONS, v));
    fix('N',  v => (/^[A-Za-z]{2}$/.test(v.trim()) ? v.trim().toUpperCase() : null));
  }
  return n;
}

// ── HS-code normaliseren bij het inladen ────────────────────────────────────
// Een ingeladen HS-code moet altijd 8 cijfers hebben (EU Combined Nomenclature-
// standaard). Bevat de bron 10 of 12 cijfers (bv. een TARIC-code met 2 of 4
// extra nationale/EU-subverdelingscijfers), dan vallen de laatste cijfers weg
// zodat er 8 overblijven. Korter dan 8 blijft ongewijzigd (dat vangt
// validateRow al af als fout). Let op: dit is een ANDER doel dan _toTaric10
// hieronder, die juist naar 10 cijfers opvult/kort t.b.v. de live opzoeking
// bij tariffnumber.com — dat systeem verwacht wél 10 cijfers (TARIC).
function _normalizeHSCode(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return v;
  const digits = s.replace(/[\s.]/g, '');
  if (!/^\d+$/.test(digits)) return v;          // geen zuiver cijferveld — met rust laten
  return digits.length > 8 ? digits.slice(0, 8) : digits;
}


function parseHColumn(value) {
  if (!value || !String(value).trim()) return [];
  const s = String(value).trim();

  // Rule 1: ÷ range  "23205-520V001 ÷ V007"
  if (s.includes('÷')) {
    const [left, right] = s.split('÷').map(p => p.trim());
    const m = right.match(/^([A-Za-z]+)(\d+)$/);
    if (m) {
      const letter = m[1], endNum = parseInt(m[2]), pad = m[2].length;
      const idx = left.lastIndexOf(letter);
      if (idx >= 0) {
        const prefix = left.slice(0, idx);
        const startNum = parseInt(left.slice(idx + letter.length));
        return Array.from({length: endNum - startNum + 1}, (_, i) =>
          `${prefix}${letter}${String(startNum + i).padStart(pad, '0')}`);
      }
    }
    return [left, right];
  }

  // Rules 2-5: split on , & /  then apply prefix-sharing for tokens starting with -
  const tokens = s.split(/\s*[,&/]\s*/);
  const results = [];
  let prevRoot = '';
  for (const token of tokens) {
    const t = token.trim();
    if (!t) continue;
    if (t.startsWith('-')) {
      results.push(prevRoot + t);
    } else {
      results.push(t);
      const di = t.indexOf('-');
      prevRoot = di >= 0 ? t.slice(0, di) : t;
    }
  }
  return results;
}

// ── USD→EUR rate via Frankfurter API ─────────────────────────────────────
async function fetchUSDRate() {
  if (_usdRate) return _usdRate;
  try {
    const r = await fetch('https://api.frankfurter.dev/v2/rate/USD/EUR');
    if (!r.ok) throw new Error('HTTP ' + r.status);
    const d = await r.json();
    _usdRate = d.rate;
    return _usdRate;
  } catch(e) {
    setStatus('Valutakoers ophalen mislukt — EUR/USD niet beschikbaar', true);
    return null;
  }
}

// ── Determine if column P is in USD ───────────────────────────────────────
function isUSD(header) {
  return /usd/i.test(String(header || ''));
}

// ── Validate a single data row ────────────────────────────────────────────
// Parse een getal uit een veld dat ook een valutateken (€, $, £, ¥) en
// duizendtal-/decimaalscheiding kan bevatten. Geeft null als het geen getal is.
function _parseNum(raw) {
  // Verwijder alles wat geen cijfer, komma, punt of min-teken is
  // (valutatekens €/$/£/¥, spaties incl. vaste spaties, letters, verborgen tekens…).
  let s = String(raw == null ? '' : raw).replace(/[^\d.,\-]/g, '');
  if (!s) return null;
  const lastComma = s.lastIndexOf(','), lastDot = s.lastIndexOf('.');
  if (lastComma > -1 && lastDot > -1) {
    // beide aanwezig → de laatste is het decimaalteken, de andere is duizendtalscheiding
    s = lastComma > lastDot ? s.replace(/\./g, '').replace(',', '.')   // europees: 1.234,56
                            : s.replace(/,/g, '');                      // amerikaans: 1,234.56
  } else if (lastComma > -1) {
    s = s.replace(',', '.');   // alleen komma → decimaalteken
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

// Herkent een valuta-aanduiding (symbool of 3-letterige code) in een celwaarde.
// Geen aanduiding gevonden → aangenomen EUR (het standaardformaat van de Itemlijst).
function _detectCurrency(raw) {
  const s = String(raw == null ? '' : raw);
  if (/€/.test(s) || /\bEUR\b/i.test(s)) return 'EUR';
  if (/\$/.test(s) || /\bUSD\b/i.test(s)) return 'USD';
  if (/£/.test(s) || /\bGBP\b/i.test(s)) return 'GBP';
  if (/¥/.test(s) || /\bJPY\b/i.test(s)) return 'JPY';
  if (/\bCHF\b/i.test(s)) return 'CHF';
  return 'EUR';
}

// ── Tekst-i.p.v.-getalnotatie detectie ────────────────────────────────────
// De template-instructie eist voor diverse kolommen "Use number formatting".
// _parseNum/_detectCurrency zijn bewust coulant (strippen valutatekens,
// eenheden e.d. en parsen toch een getal), maar dat betekent dat een cel die
// feitelijk als TEKST is ingevuld (bv. "5 stuks", "1250 kg") zonder melding
// werd geaccepteerd — exact de "Textformat"-meldingen die vanuit de
// Moederlijst terugkwamen. Deze helper herkent of er, ná het wegstrepen van
// een valutasymbool/-code, nog iets anders dan cijfers/scheidingstekens
// overblijft, en zo ja: een waarschuwing i.p.v. stilzwijgend doorlaten.
function _looksLikeCleanNumber(raw) {
  const s = String(raw == null ? '' : raw).trim();
  if (!s) return true; // leeg wordt elders als 'verplicht' afgevangen
  const stripped = s
    .replace(/^[€$£¥]\s*/, '').replace(/\s*[€$£¥]$/, '')
    .replace(/\b(EUR|USD|GBP|JPY|CHF)\b/ig, '')
    .trim();
  return /^-?[\d\s.,]+$/.test(stripped);
}

function _fmtWarnIfTextual(col, label, raw, errors, warnings) {
  if (errors[col]) return;                 // al een fout op deze kolom — geen dubbele melding
  const s = String(raw == null ? '' : raw).trim();
  if (!s || _looksLikeCleanNumber(s)) return;
  warnings[col] = (warnings[col] ? warnings[col] + ' — ' : '') +
    `${label} lijkt als tekst ingevuld i.p.v. getalnotatie (was: '${s}')`;
}

async function validateRow(cells, isUSDPrice, usdRate, coo, expeditingData) {
  const errors   = {};  // col letter → error message
  const warnings = {};  // col letter → warning message
  const computed = {};  // col letter → computed value to write back

  const v  = (col) => cells[COL[col]];
  const vs = (col) => String(v(col) ?? '').trim();
  const vn = (col) => _parseNum(v(col));

  // ── B: Project — IHC-veld, geen harde eis maar wel overtollige spaties melden ──
  const bRaw = v('B');
  if (bRaw !== null && bRaw !== undefined && String(bRaw) !== '') {
    const bStr = String(bRaw);
    if (/^\s|\s$|\s{2,}|\u00A0/.test(bStr)) {
      warnings['B'] = `Project # bevat overtollige spaties (was: '${bStr}')`;
    }
  }

  // ── C: IHC PO — required, not empty ──────────────────────────────────────
  if (!vs('C')) errors['C'] = 'IHC PO is verplicht';

  // ── D: Item — must contain '-', else look up H in Expediting Kol M ───────
  const dVal = vs('D');
  if (dVal && !dVal.includes('-')) {
    warnings['D'] = `Geen '-' in Item — H-waarden worden opgezocht in Expediting`;
  } else if (!dVal) {
    // Check if H can serve as fallback
    const hVal = vs('H');
    if (!hVal) {
      errors['D'] = 'Item # ontbreekt en Mark/Label (H) is ook leeg';
    } else {
      const hCodes = parseHColumn(hVal);
      if (expeditingData && expeditingData.length) {
        const found = hCodes.filter(code =>
          expeditingData.some(row => {
            const mVal = String(Object.values(row)[12] || '').trim();
            return mVal.toLowerCase() === String(code).trim().toLowerCase();
          })
        );
        if (!found.length) {
          warnings['D'] = `Item leeg — H-waarden (${hCodes.join(', ')}) niet gevonden in Expediting Kol M`;
        } else {
          warnings['D'] = `Item leeg — gevonden via H: ${found.join(', ')}`;
        }
      } else {
        warnings['D'] = `Item leeg — H-waarden: ${hCodes.join(', ')} (Expediting niet geladen)`;
      }
    }
  }

  // ── E: Item description — required ───────────────────────────────────────
  if (!vs('E')) errors['E'] = 'Item description is verplicht';

  // ── F: Quantity — required, numeric > 0 ──────────────────────────────────
  const qty = vn('F');
  if (qty === null) errors['F'] = 'Quantity moet een getal zijn';
  else if (qty <= 0)  errors['F'] = 'Quantity moet groter zijn dan 0';
  else computed['F'] = qty;   // schone numerieke waarde — export schrijft dit terug i.p.v. de ruwe tekst
  _fmtWarnIfTextual('F', 'Quantity', v('F'), errors, warnings);

  // ── G: Unit of measure — required, must be in list ────────────────────────
  const uom = vs('G');
  if (!uom)                errors['G'] = 'Unit of measure is verplicht';
  else if (!_valCanon(VL_UOM, uom)) errors['G'] = `'${uom}' niet in toegestane lijst`;

  // ── H: Mark/Label — required ──────────────────────────────────────────────
  if (!vs('H')) errors['H'] = 'Component (Mark/Label) is verplicht';

  // ── I: Supplier article number — verplicht (door leverancier aan te leveren) ──
  if (!vs('I')) errors['I'] = 'Supplier article number is verplicht';

  // ── J: Serial number / (bij Eriks/W&O) ons eigen partnummer ──────────────
  // Bij Eriks/W&O hoort hier ons eigen partnummer te staan, automatisch
  // gevuld vanuit Expediting (kolom K "Part No") via de Unified Reference
  // Code-match op kolom H — zie _fillRowFromExpediting(). Staat het hier
  // na die poging nog leeg, dan kon die koppeling niet gelegd worden.
  // Voor overige leveranciers blijft J optioneel (geen verplichte validatie).
  if (!vs('J') && _isPartNoSupplier(v('K'))) {
    warnings['J'] = 'Partnummer (kolom J) ontbreekt — Unified Reference Code (kolom H) kon niet gekoppeld '
      + 'worden aan de Expediting-lijst, of Part No staat daar leeg. Controleer de koppeling handmatig.';
  }

  // ── K: Supplier — required ────────────────────────────────────────────────
  if (!vs('K')) errors['K'] = 'Supplier is verplicht';

  // ── L: Make — required ───────────────────────────────────────────────────
  if (!vs('L')) errors['L'] = 'Make is verplicht';

  // ── M: Material — required ────────────────────────────────────────────────
  if (!vs('M')) errors['M'] = 'Material is verplicht';

  // ── N: Country of origin — required, must be ISO-2 ────────────────────────
  const cooVal = vs('N');
  const cooSet = coo.size > 0 ? coo : VL_COO_FALLBACK;
  if (!cooVal)               errors['N'] = 'Country of origin is verplicht';
  else if (!cooSet.has(cooVal.toUpperCase())) errors['N'] = `'${cooVal}' is geen geldige landcode`;

  // ── O: HS-code — required + format check; live nomenclature checked async ──
  const hsVal = vs('O');
  let hsClean = '';
  if (!hsVal) {
    errors['O'] = 'HS-code is verplicht';
  } else {
    hsClean = hsVal.replace(/\s+/g,'').trim();
    if (!/^\d{8}$/.test(hsClean)) {
      errors['O'] = `HS-code moet 8 cijfers zijn (EU Combined Nomenclature) — was: '${hsClean}'`;
    }
    // Live nomenclatuur-check via douane.nl wordt async uitgevoerd na validatie
  }
  // Bij het inladen wordt een HS-code met spaties/punten of >8 cijfers al
  // automatisch opgeschoond (zie _normalizeHSCode in handleValFile) — dat
  // gebeurde tot nu toe stilzwijgend, waardoor niemand zag dát de leverancier
  // de instructie "Do not use space or dots" niet volgde. Dit maakt die
  // correctie zichtbaar als waarschuwing i.p.v. onopgemerkt te verdwijnen.
  if (cells._hsOrigRaw !== undefined) {
    warnings['O'] = (warnings['O'] ? warnings['O'] + ' — ' : '') +
      `HS-code automatisch opgeschoond (spaties/punten verwijderd of ingekort tot 8 cijfers): '${cells._hsOrigRaw}' → '${hsClean}'`;
  }

  // ── P: Value pc — required, numeric, alleen EUR toegestaan (USD via toggle) ──
  const pCur = _detectCurrency(v('P'));
  let pVal = vn('P');   // altijd alleen de cijfers — valutatekens worden hier al genegeerd
  if (pVal === null) {
    errors['P'] = 'Value pc is verplicht en moet numeriek zijn';
  } else if (pCur === 'USD' && isUSDPrice && usdRate) {
    const pEur = +(pVal * usdRate).toFixed(2);
    computed['P_EUR'] = pEur;
    warnings['P'] = `USD ${pVal.toLocaleString('nl-NL')} = EUR ${pEur.toLocaleString('nl-NL')} (koers: ${usdRate})`;
    pVal = pEur; // use EUR value for Q cross-check
  } else if (pCur !== 'EUR') {
    errors['P'] = `Waarde staat in ${pCur} — verwacht EUR. Zet de waarde om naar EUR of vink USD-koers aan.`;
  }
  if (pVal !== null && !errors['P']) computed['P'] = pVal;   // schone EUR-waarde zonder valutateken/tekst
  _fmtWarnIfTextual('P', 'Value pc', v('P'), errors, warnings);

  // ── Q: Value total — required, numeric, alleen EUR toegestaan; warn if >1% off from P×F ─
  const qCur = _detectCurrency(v('Q'));
  const qVal = vn('Q');
  if (qVal === null) {
    errors['Q'] = 'Value total is verplicht en moet numeriek zijn';
  } else if (qCur !== 'EUR' && !(qCur === 'USD' && isUSDPrice && usdRate)) {
    errors['Q'] = `Waarde staat in ${qCur} — verwacht EUR. Zet de waarde om naar EUR of vink USD-koers aan.`;
  } else if (pVal !== null && qty !== null) {
    const expected = pVal * qty;
    const diff = Math.abs(qVal - expected);
    const pct  = expected > 0 ? (diff / expected) * 100 : 0;
    if (pct > 1) {
      warnings['Q'] = `Value total (${qVal}) wijkt ${pct.toFixed(1)}% af van Value pc × Qty (${expected.toFixed(2)})`;
    }
  }
  if (qVal !== null && !errors['Q']) computed['Q'] = qVal;   // schone EUR-waarde zonder valutateken/tekst
  _fmtWarnIfTextual('Q', 'Value total', v('Q'), errors, warnings);

  // ── R/S/T/U/V/X/Y: at least one row must have these — checked at sheet level
  // Per-row: validate types
  const t = vn('T'), u = vn('U'), h2 = vn('V');
  if (v('T') !== null && v('T') !== '' && t === null) errors['T'] = 'Length moet numeriek zijn';
  if (v('U') !== null && v('U') !== '' && u === null) errors['U'] = 'Width moet numeriek zijn';
  if (v('V') !== null && v('V') !== '' && h2 === null) errors['V'] = 'Height moet numeriek zijn';
  _fmtWarnIfTextual('T', 'Length', v('T'), errors, warnings);
  _fmtWarnIfTextual('U', 'Width', v('U'), errors, warnings);
  _fmtWarnIfTextual('V', 'Height', v('V'), errors, warnings);
  if (t !== null)  computed['T'] = t;
  if (u !== null)  computed['U'] = u;
  if (h2 !== null) computed['V'] = h2;

  // ── W: Volume — compute if T/U/V present ─────────────────────────────────
  if (t && u && h2) {
    const vol = +(t * u * h2 / 1000000).toFixed(4);
    computed['W'] = vol;
    const wVal = vn('W');
    if (wVal && Math.abs(wVal - vol) > 0.001) {
      warnings['W'] = `Volume ${wVal} klopt niet — berekend: ${vol} m³`;
    }
  }

  // ── X/Y: weight ───────────────────────────────────────────────────────────
  const xVal = vn('X'), yVal = vn('Y');
  if (v('X') !== null && v('X') !== '' && xVal === null) errors['X'] = 'Gross weight moet numeriek zijn';
  if (v('Y') !== null && v('Y') !== '' && yVal === null) errors['Y'] = 'Nett weight moet numeriek zijn';
  if (xVal !== null && yVal !== null && yVal > xVal)
    errors['Y'] = `Nett weight (${yVal}) mag niet groter zijn dan Gross weight (${xVal})`;
  _fmtWarnIfTextual('X', 'Gross weight', v('X'), errors, warnings);
  _fmtWarnIfTextual('Y', 'Nett weight', v('Y'), errors, warnings);
  if (xVal !== null) computed['X'] = xVal;
  if (yVal !== null) computed['Y'] = yVal;

  // ── AA: Dangerous Goods — tickbox, geen tekstvalidatie nodig ──────────────

  // ── AB/AC: Dual use + ECCN code — ECCN verplicht zodra Dual use is aangevinkt ──
  const abVal       = v('AB');
  const hasDualUseCB = _valBool(abVal);
  const eccnVal      = vs('AC');
  if (hasDualUseCB && !eccnVal) {
    errors['AC'] = 'ECCN code is verplicht wanneer Dual use is aangevinkt';
  } else if (!hasDualUseCB && eccnVal) {
    warnings['AB'] = 'ECCN code ingevuld maar Dual use niet aangevinkt';
  }

  // ── AH: Inspection Level — optional but must be valid if filled ───────────
  const ahVal = vs('AH');
  if (ahVal && !VL_INSP_LC.has(String(ahVal).toLowerCase())) errors['AH'] = `'${ahVal}' niet in toegestane lijst`;

  // ── AFS-voormelding: ≥2 maten > 3 m (300 cm) óf bruto > 10.000 kg ─────────
  {
    const reasons = [];
    if ([t, u, h2].filter(d => d != null && d > 300).length >= 2) reasons.push('oversized');
    if (xVal != null && xVal > 10000) reasons.push('>10t');
    if (reasons.length) computed._afs = reasons.join('+');
  }

  return { errors, warnings, computed };
}

// ── Sheet-level validation (min 1 row with R/S/T/U/V/X/Y) ────────────────
function validateSheet(rows) {
  const sheetWarnings = [];
  const checks = { R:false, S:false, T:false, U:false, V:false, X:false, Y:false };
  for (const row of rows) {
    for (const col of Object.keys(checks)) {
      const val = row.cells[COL[col]];
      if (val !== null && val !== undefined && String(val).trim()) checks[col] = true;
    }
  }
  for (const [col, ok] of Object.entries(checks)) {
    if (!ok) sheetWarnings.push(`Kolom ${col}: minimaal 1 rij met een waarde vereist`);
  }
  return sheetWarnings;
}

// ── Collonummer ↔ maatvoering (L×B×H) consistentie ──────────────────────────
// Regel: rijen met hetzelfde collonummer moeten dezelfde maatvoering hebben.
// Staat één collonummer op meerdere regels met afwijkende L×B×H, dan is dat fout
// (elke unieke maatvoering hoort een eigen collonummer te hebben).

// Vindt de kolomletter van het collonummer op basis van de headernaam.
function _findColloLetter() {
  const hdr = L => String(_valHeaders[COL[L]] || '').toLowerCase();
  const letters = Object.keys(COL);
  const isCollo = h => /coll[io]/.test(h) || /(package|pakket|pkg)/.test(h);
  const isCount = h => /(aantal|qty|pcs|stuks|count|#\s*coll)/.test(h);
  const hasNr   = h => /(nr|no|nummer|number|#)/.test(h);
  let L = letters.find(l => { const h = hdr(l); return isCollo(h) && hasNr(h) && !isCount(h); });
  if (L) return L;
  L = letters.find(l => { const h = hdr(l); return isCollo(h) && !isCount(h); });
  return L || null;
}

// Genormaliseerde L×B×H-handtekening van een rij. Een rij zonder eigen
// maatvoering (T/U/V allemaal leeg) geeft null terug — die rij "erft" de
// maatvoering van de eerste regel van het collonummer en is dus nooit in
// conflict, in plaats van als eigen (lege) handtekening mee te tellen.
function _dimSignature(cells) {
  const num = c => { const n = parseFloat(String(c ?? '').replace(',', '.').trim()); return isNaN(n) ? null : n; };
  const t = num(cells[COL.T]), u = num(cells[COL.U]), v = num(cells[COL.V]);
  if (t === null && u === null && v === null) return null;
  return t + '×' + u + '×' + v;
}

function _validateColloConsistency(rows) {
  const cLetter = _findColloLetter();
  if (!cLetter) return { added: 0, warnings: [], collo: null };   // geen collonummer-kolom gevonden
  const cIdx = COL[cLetter];

  // Groepeer rijen per collonummer
  const groups = new Map();
  rows.forEach(row => {
    const collo = String(row.cells[cIdx] ?? '').trim();
    if (!collo) return;
    if (!groups.has(collo)) groups.set(collo, []);
    groups.get(collo).push(row);
  });

  let added = 0; const warnings = [];
  for (const [collo, members] of groups) {
    if (members.length < 2) continue;
    // Alleen rijen mét eigen maatvoering meetellen — regels zonder eigen
    // T/U/V erven de maatvoering van het collonummer en zijn nooit conflicterend.
    const sigs = new Set(members.map(m => _dimSignature(m.cells)).filter(s => s !== null));
    if (sigs.size > 1) {   // zelfde collonummer, verschillende maatvoering → fout
      warnings.push(`Collo ${collo}: ${members.length} regels met verschillende maatvoering (L×B×H) — moet gelijk zijn`);
      for (const row of members) {
        if (!row.errors[cLetter]) { row.errors[cLetter] = `Collo ${collo} heeft afwijkende maatvoering (L×B×H) op meerdere regels`; added++; }
        for (const L of ['T','U','V']) {
          if (!row.errors[L]) { row.errors[L] = `Maatvoering wijkt af binnen collo ${collo}`; added++; }
        }
      }
    }
  }
  return { added, warnings, collo: cLetter };
}

// ── Slim aanvullen vanuit Expediting ────────────────────────────────────────
// Blauwe Boekje nummer = Unified Reference Code. Match: itemlijst-kolom H
// (Mark/Label, kan meerdere codes bevatten) ↔ Expediting "Unified Reference Code".
// Bij een match worden LEGE itemcellen aangevuld vanuit de Expediting-rij.
let _expIndex = null;       // Map: urefCode(lowercase) → eerste Expediting-rij
let _expIndexFor = null;    // databron waarvoor de index is opgebouwd

function _buildExpIndex(expeditingData) {
  if (_expIndex && _expIndexFor === expeditingData) return _expIndex;
  const m = new Map();
  if (Array.isArray(expeditingData)) {
    for (const row of expeditingData) {
      const uref = String(row['Unified Reference Code'] ?? '').trim();
      if (uref && !m.has(uref.toLowerCase())) m.set(uref.toLowerCase(), row);
    }
  }
  _expIndex = m; _expIndexFor = expeditingData;
  return m;
}

// ── Eriks / W&O: ons eigen partnummer in kolom J ───────────────────────────
// Bij deze twee leveranciers hoort in kolom J (Serial number — deze kolom
// wordt later hernoemd naar "Part NR") óns eigen partnummer te staan — niet
// een door de leverancier verzonnen artikelcode. Dat partnummer staat in de
// Expediting-lijst, kolom K ("Part No"), op dezelfde rij als de Unified
// Reference Code waarmee hierboven al gematcht wordt (itemlijst H).
function _isPartNoSupplier(supplierName) {
  const s = String(supplierName ?? '');
  return /\beriks\b/i.test(s) || /\bw\s*&\s*o\b/i.test(s);
}

// Expediting-veldnaam → itemlijst-kolomletter. Alleen LEGE cellen worden gevuld.
const _EXP_FILL_MAP = [
  ['Order No',          'C'],  // IHC PO (basis-PO = Order No)
  // Kolom D (Item) wordt apart samengesteld: '"+Line No+"-"+Release No
  // ('prefix voorkomt datuminterpretatie in Excel, bv. '7-1)
  ['Description',       'E'],  // Omschrijving
  ['Supplier Name',     'K'],  // Supplier
  ['Country of Origin', 'N'],  // Land van oorsprong
  ['Customs Stat No',   'O'],  // HS-code
];

// Kolomletters die automatisch aangevuld kunnen worden (voor de header-markering)
const _FILL_COLS = new Set(_EXP_FILL_MAP.map(x => x[1]));

function _fillRowFromExpediting(cells, expeditingData) {
  if (!expeditingData || !expeditingData.length) return 0;
  const idx = _buildExpIndex(expeditingData);
  const hCodes = parseHColumn(String(cells[COL.H] ?? '').trim());
  if (!hCodes.length) return 0;

  let match = null;
  for (const code of hCodes) {
    const hit = idx.get(String(code).trim().toLowerCase());
    if (hit) { match = hit; break; }
  }
  if (!match) return 0;

  let filled = 0;

  // ── Kolom B (Project): altijd het Sub Project ID uit Expediting (bv. YN1320 / E10000) ──
  {
    const sid = String(match[_valSubKey] ?? match['Sub Project ID'] ?? '').trim();
    if (sid && String(cells[COL.B] ?? '').trim() !== sid) { cells[COL.B] = sid; filled++; }
  }

  // ── Kolom D (Item): '"+Line No+"-"+Release No  (apostrof = tekstprefix Excel) ──
  if (!String(cells[COL.D] ?? '').trim()) {
    const line    = String(match['Line No']    ?? '').trim();
    const release = String(match['Release No'] ?? '').trim();
    if (line || release) {
      cells[COL.D] = line && release ? `'-${line}-${release}` : line || release;
      filled++;
    }
  }

  // ── Overige kolommen via de fill-map ──────────────────────────────────────
  for (const [field, colLetter] of _EXP_FILL_MAP) {
    const ci = COL[colLetter];
    if (ci === undefined) continue;
    if (String(cells[ci] ?? '').trim()) continue;        // niet overschrijven
    const val = match[field];
    if (val === null || val === undefined || String(val).trim() === '') continue;
    cells[ci] = (colLetter === 'O') ? _normalizeHSCode(String(val).trim()) : String(val).trim();
    filled++;
  }

  // ── J: Serial number → bij Eriks/W&O vervangen door ons eigen partnummer
  // (Expediting kolom K "Part No"), via dezelfde H ↔ Unified Reference
  // Code-match hierboven. De leverancier vult hier een serienummer in, maar
  // voor deze twee leveranciers hoort op die plek óns partnummer te staan —
  // de crossreference overschrijft die cel dus bewust (kolom I blijft
  // ongemoeid: dat blijft het eigen artikelnummer van de leverancier).
  if (_isPartNoSupplier(cells[COL.K])) {
    const partNo = String(match['Part No'] ?? '').trim();
    if (partNo && String(cells[COL.J] ?? '').trim() !== partNo) {
      cells[COL.J] = partNo;
      filled++;
    }
  }

  return filled;
}

// ── Centrale expediting-koppeling + Sub Project ID-selectie (Itemlijst-Validator) ──
// Draait alleen op de validatorpagina (herkenbaar aan de .val-toolbar). Vult
// fileData.expediting vanuit de gedeelde datalaag en bouwt een Sub Project ID-keuze.
let _valSubProjects = new Set();   // meervoudige selectie; leeg = alle projecten
let _valSubKey = 'Sub Project ID';

// Vindt de Sub Project ID-kolomnaam, ongeacht spaties/hoofdletters/schrijfwijze
function _valDetectSubKey(rows, headers) {
  const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');
  const hs = (headers && headers.length) ? headers : (rows[0] ? Object.keys(rows[0]) : []);
  return hs.find(h => norm(h) === 'subprojectid')
      || hs.find(h => norm(h).includes('subproject'))
      || hs.find(h => norm(h).includes('project') && norm(h).includes('id'))
      || 'Sub Project ID';
}

function _valExpRows() {
  const all = (typeof fileData !== 'undefined' && fileData.expediting && Array.isArray(fileData.expediting.data))
    ? fileData.expediting.data : [];
  if (_valSubProjects.size === 0) return all;   // geen selectie = alle projecten
  return all.filter(r => _valSubProjects.has(String(r[_valSubKey] ?? '').trim()));
}

// Injecteert dezelfde opmaak als de PO-Matcher Sub Project ID-selector
function _valInjectSPStyle() {
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

function _buildSubProjectSelector(rows) {
  const bar = document.querySelector('.val-toolbar');
  if (!bar || document.getElementById('val-sp-panel')) return;
  _valSubKey = _valDetectSubKey(rows, fileData.expediting && fileData.expediting.headers);

  const counts = new Map();
  rows.forEach(r => { const id = String(r[_valSubKey] ?? '').trim(); if (id) counts.set(id, (counts.get(id) || 0) + 1); });
  const MIN_ROWS = 11;   // projecten met 10 of minder regels weglaten uit de lijst
  const ids = [...counts.keys()]
    .filter(id => counts.get(id) >= MIN_ROWS)
    .sort((a, b) => a.localeCompare(b, 'nl', { numeric: true }));

  _valInjectSPStyle();

  const panel = document.createElement('div');
  panel.className = 'val-sp-panel';
  panel.id = 'val-sp-panel';
  panel.innerHTML =
    `<div class="val-sp-head">📋 Bedrijfsbreed Expediten — selecteer Sub Project ID</div>` +
    `<input class="val-sp-search" id="val-sp-search" placeholder="🔍 Zoek Sub Project ID…" autocomplete="off">` +
    `<div class="val-sp-actions">` +
    `<button class="val-sp-btn" id="val-sp-all" type="button">Alles</button>` +
    `<button class="val-sp-btn" id="val-sp-none" type="button">Wis</button>` +
    `<span class="val-sp-count" id="val-sp-count">0 geselecteerd</span>` +
    `</div>` +
    `<div class="val-sp-list" id="val-sp-list">` +
    ids.map(id => `<label class="val-sp-item" data-id="${esc(id)}"><input type="checkbox" value="${esc(id)}"><span class="val-sp-id">${esc(id)}</span><span class="val-sp-n">${counts.get(id)}</span></label>`).join('') +
    `</div>`;
  bar.parentNode.insertBefore(panel, bar);

  const apply = () => {
    const cnt = document.getElementById('val-sp-count');
    if (cnt) cnt.textContent = _valSubProjects.size + ' geselecteerd';
    // Is er al een itemlijst gevalideerd? Opnieuw valideren met de nieuwe selectie.
    if (typeof _valRows !== 'undefined' && _valRows.length && typeof runValidation === 'function') runValidation();
  };
  panel.querySelector('#val-sp-list').addEventListener('change', e => {
    if (e.target.type !== 'checkbox') return;
    if (e.target.checked) _valSubProjects.add(e.target.value); else _valSubProjects.delete(e.target.value);
    apply();
  });
  panel.querySelector('#val-sp-search').addEventListener('input', function () {
    const q = this.value.trim().toLowerCase();
    panel.querySelectorAll('.val-sp-item').forEach(it =>
      it.classList.toggle('hidden', !!q && !it.dataset.id.toLowerCase().includes(q)));
  });
  panel.querySelector('#val-sp-all').addEventListener('click', () => {
    panel.querySelectorAll('.val-sp-item:not(.hidden) input').forEach(cb => { cb.checked = true; _valSubProjects.add(cb.value); });
    apply();
  });
  panel.querySelector('#val-sp-none').addEventListener('click', () => {
    panel.querySelectorAll('.val-sp-item input').forEach(cb => cb.checked = false);
    _valSubProjects.clear();
    apply();
  });
}

async function _valConnectExpediting() {
  if (!document.querySelector('.val-toolbar')) return;      // alleen op de validatorpagina
  if (typeof fileData === 'undefined' || !window.ExpeditingData) return;
  let raw = null, meta = null;
  try { raw = await window.ExpeditingData.loadRaw(); meta = await window.ExpeditingData.meta(); }
  catch (e) { console.error('Centrale expediting-lijst lezen mislukt:', e); }

  const sh = document.getElementById('val-sheet-warnings');
  if (raw && Array.isArray(raw.rows) && raw.rows.length) {
    fileData.expediting = { data: raw.rows, headers: raw.headers,
      name: (meta && meta.filename) || 'Bedrijfsbreed', source: meta && meta.source };
    _buildSubProjectSelector(raw.rows);
  } else if (sh) {
    sh.innerHTML = '<span style="color:#f59e0b">⚠ Geen centrale expeditinglijst gevonden — stel in via Admin. ' +
      'De lookups/aanvullingen bij het valideren werken pas als de lijst gekoppeld is.</span>';
  }
}

if (document.readyState === 'loading')
// Laadt de mail-template-generator vanzelf op de validatorpagina — ook als de
// index het script niet expliciet inlaadt. Dubbel laden wordt voorkomen.
function _valLoadMailgen() {
  if (!document.querySelector('.val-toolbar')) return;                 // alleen validatorpagina
  if (window.ValMailer || document.querySelector('script[data-valmailgen]')) return;  // al geladen
  var s = document.createElement('script');
  s.src = 'val-mailgen.js';
  s.setAttribute('data-valmailgen', '1');
  s.onerror = function () { console.warn('val-mailgen.js kon niet geladen worden'); };
  document.head.appendChild(s);
}

document.addEventListener('DOMContentLoaded', _valConnectExpediting);
document.addEventListener('DOMContentLoaded', _valLoadMailgen);
if (document.readyState !== 'loading') _valLoadMailgen();

// ── Export → mail naar Wendels (val-export-mail.js) ────────────────────────
// Bestandsnaam én onderwerp: "{Delivery ref} ITEMLIJST {Supplier}"
function _valExportBaseName() {
  const dom = (ci) => {
    const counts = {};
    _valRows.forEach(r => { const v = String(r.cells[ci] ?? '').trim(); if (v) counts[v] = (counts[v] || 0) + 1; });
    return (Object.entries(counts).sort((a, b) => b[1] - a[1])[0] || [''])[0];
  };
  const clean = s => String(s || '').replace(/[/\\:*?"<>|]/g, '-').trim();
  return [clean(dom(COL.A)), 'ITEMLIJST', clean(dom(COL.K))].filter(Boolean).join(' ');
}

function _valLoadExportMail() {
  if (!document.querySelector('.val-toolbar')) return;
  if (window.ValExportMail || document.querySelector('script[data-valexportmail]')) return;
  var s = document.createElement('script');
  s.src = 'val-export-mail.js';
  s.setAttribute('data-valexportmail', '1');
  s.onerror = function () { console.warn('val-export-mail.js kon niet geladen worden'); };
  document.head.appendChild(s);
}

// Normaliseert kolom B naar een bekend Sub Project ID (YN1320 / E10000), nooit EN10000 of kale 1320
function _valNormSubId(raw) {
  const v = String(raw ?? '').replace(/\s+/g, '').toUpperCase();
  const known = [];
  try {
    const all = (typeof fileData !== 'undefined' && fileData.expediting && Array.isArray(fileData.expediting.data)) ? fileData.expediting.data : [];
    const seen = new Set();
    all.forEach(r => { const id = String(r[_valSubKey] ?? '').trim(); if (id && !seen.has(id)) { seen.add(id); known.push(id); } });
  } catch (e) {}
  if (v) {
    const exact = known.find(k => k.toUpperCase() === v);
    if (exact) return exact;
    const digits = v.replace(/\D/g, '');
    if (digits) {
      const cand = known.filter(k => k.replace(/\D/g, '') === digits);
      if (cand.length === 1) return cand[0];
      if (cand.length > 1) {   // voorkeur: zelfde letters-prefix, anders kortste/eerste
        const pre = v.replace(/\d/g, '');
        const p = cand.find(k => k.toUpperCase().replace(/\d/g, '') === pre) ||
                  cand.find(k => k.toUpperCase().startsWith(pre.charAt(0)));
        return p || cand[0];
      }
    }
  }
  if (_valSubProjects && _valSubProjects.size === 1) return [..._valSubProjects][0];
  return String(raw ?? '').trim();
}

// Haalt valutatekens/scheiders uit een bedrag → schoon getal als string (leeg blijft leeg)
function _valCleanMoney(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'number') return isFinite(v) ? v : '';
  let t = String(v).replace(/[\u00A0\s]/g, '').replace(/[^0-9.,\-]/g, '');
  if (!t || !/\d/.test(t)) return '';
  const lc = t.lastIndexOf(','), ld = t.lastIndexOf('.');
  if (lc > -1 && ld > -1) t = lc > ld ? t.replace(/\./g, '').replace(',', '.') : t.replace(/,/g, '');
  else if (lc > -1) t = /,\d{1,2}$/.test(t) && t.indexOf(',') === lc ? t.replace(',', '.') : t.replace(/,/g, '');
  else if ((t.match(/\./g) || []).length > 1) t = t.replace(/\./g, '');
  const n = parseFloat(t);
  return isFinite(n) ? n : '';
}

// Bouwt objecten per rij, gesleuteld op kolomletter (A…AH). De volgorde/kopnamen van de
// bijlage staan in val-export-mail.js (EXPORT_LAYOUT) — zo blijft de export onafhankelijk
// van hoe de leverancier de kolommen heeft genoemd.
function _valMailRows() {
  return _valRows.map(r => {
    const o = {};
    Object.entries(COL).forEach(([L, ci]) => {
      let v = r.cells[ci];
      if (v === undefined || v === null) v = '';
      if (typeof v === 'string' && v.trim() === '') v = '';          // lege velden écht leeg (geen spatie)
      if (L === 'B') v = v === '' ? '' : _valNormSubId(v);
      else if (L === 'P' || L === 'Q') v = _valCleanMoney(v);        // geen €/$/valutatekens
      else if (L === 'AA' || L === 'AB') v = _valBool(v);            // TRUE/FALSE (nooit YES/NO)
      else if (L === 'C' && typeof v === 'string') v = v.replace(/^\s*PO[\s\-_:]*(?=\d)/i, '').trim();   // IHC PO zonder prefix 'PO'
      else if (L === 'O' && v !== '') v = String(v).replace(/[\s.]/g, '');                                // HS-code: geen spaties/punten
      else if (['F', 'T', 'U', 'V', 'X', 'Y'].includes(L) && v !== '') {                                // "Use number formatting" → echte getallen
        const cn = r.computed && r.computed[L];
        if (typeof v === 'number') { /* al een getal */ }
        else if (typeof cn === 'number' && isFinite(cn)) v = cn;
        else { const pn = _parseNum(v); if (pn !== null && _looksLikeCleanNumber(v)) v = pn; }
      }
      else if (L === 'W' && v === '') {                              // Volume uit L×B×H als de bron geen waarde heeft
        const t = _parseNum(r.cells[COL.T]), u = _parseNum(r.cells[COL.U]), h = _parseNum(r.cells[COL.V]);
        if ([t, u, h].every(x => x !== null && x !== undefined && !isNaN(x))) v = Math.round(t * u * h / 1e6 * 1e6) / 1e6;
      }
      o[L] = v;
    });
    return o;
  });
}

function _valExportMail() {
  if (!_valRows || !_valRows.length) { alert('Laad en valideer eerst een Itemlijst.'); return; }
  if (!window.ValExportMail || typeof window.ValExportMail.open !== 'function') {
    alert('Mailmodule (val-export-mail.js) is nog niet geladen. Ververs de pagina en probeer opnieuw.'); return;
  }
  window.ValExportMail.open(_valMailRows());
}

// Koppelt de export-knop aan de mailflow (i.p.v. de directe download)
function _valWireExportMail() {
  if (!document.querySelector('.val-toolbar')) return;
  const btn = document.getElementById('btn-val-export');
  if (!btn || btn.dataset.mailWired) return;
  btn.dataset.mailWired = '1';
  btn.removeAttribute('onclick');
  btn.onclick = _valExportMail;
  if (/export/i.test(btn.textContent)) btn.textContent = '\uD83D\uDCE7 Exporteer & mail';
}

document.addEventListener('DOMContentLoaded', _valLoadExportMail);
document.addEventListener('DOMContentLoaded', _valWireExportMail);
if (document.readyState !== 'loading') { _valLoadExportMail(); _valWireExportMail(); }

// ── Moederlijst uploaden → koppelen aan smart-fill / cross-ref ─────────────
// Zet een geüploade Moederlijst (expediting-lijst) in fileData.expediting,
// exact de koppeling die _valConnectExpediting ook gebruikt, en hervalideert.
function _valPickMoederSheet(wb) {
  const names = (wb.SheetNames || []);
  const rows = (n) => { const ws = wb.Sheets[n]; if (!ws || !ws['!ref']) return 0;
    const r = XLSX.utils.decode_range(ws['!ref']); return (r.e.r - r.s.r + 1); };
  const nonCipl = names.filter(n => !/cipl|shipment/i.test(n));   // CIPL-/shipment-bladen overslaan
  const pool = nonCipl.length ? nonCipl : names;
  return pool.slice().sort((a, b) => rows(b) - rows(a))[0] || names[0];
}

function _valHandleMoederUpload(ev) {
  const file = ev.target.files && ev.target.files[0];
  ev.target.value = '';   // reset zodat hetzelfde bestand opnieuw gekozen kan worden
  if (!file) return;
  const sh = document.getElementById('val-sheet-warnings');
  if (sh) sh.innerHTML = '<span style="color:var(--teal)">\u23F3 Moederlijst inlezen\u2026</span>';
  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const pick = _valPickMoederSheet(wb);
      const raw = XLSX.utils.sheet_to_json(wb.Sheets[pick], { header: 1, raw: true, defval: '' });
      let table = { headers: [], rows: [] };
      if (window.ExpeditingCore && ExpeditingCore.rawTable) table = ExpeditingCore.rawTable(raw);
      if (!table.rows || !table.rows.length) {
        if (sh) sh.innerHTML = '<span style="color:#f59e0b">\u26A0 Geen bruikbare regels in de Moederlijst gevonden.</span>';
        return;
      }
      if (typeof fileData === 'undefined') window.fileData = {};
      fileData.expediting = { data: table.rows, headers: table.headers, name: file.name, source: 'upload' };
      if (typeof _buildSubProjectSelector === 'function') _buildSubProjectSelector(table.rows);
      const heeftItemlijst = (typeof _valRows !== 'undefined' && _valRows && _valRows.length);
      if (sh) sh.innerHTML = '<span style="color:#22c55e">\u2713 Moederlijst gekoppeld: <b>' + esc(file.name) +
        '</b> (' + table.rows.length + ' regels)' +
        (heeftItemlijst ? ' \u2014 opnieuw aan het valideren\u2026' : ' \u2014 upload nu een Itemlijst en klik Valideer.') + '</span>';
      if (heeftItemlijst && typeof runValidation === 'function') runValidation();
    } catch (err) {
      console.error('Moederlijst-upload mislukt:', err);
      if (sh) sh.innerHTML = '<span style="color:#ef4444">Fout bij inlezen Moederlijst: ' + esc(err.message) + '</span>';
    }
  };
  reader.readAsArrayBuffer(file);
}

// ── Teambrede telling van validaties ───────────────────────────────────────
// Stuurt een klein event naar /api/validaties (zie shared/validatie-log.js).
// Vereist geen handeling van de gebruiker en mag nooit iets breken.
function _valLogValidatie() {
  try {
    if (!window.ValidatieLog || !Array.isArray(_valRows) || !_valRows.length) return;
    let fouten = 0, waarschuwingen = 0;
    for (const r of _valRows) {
      if (r && r.errors)   fouten += Object.keys(r.errors).length;
      if (r && r.warnings) waarschuwingen += Object.keys(r.warnings).length;
    }
    const dom = (idx) => {                       // meest voorkomende waarde in een kolom
      const t = {};
      for (const r of _valRows) {
        const v = String((r.cells && r.cells[idx]) || '').trim();
        if (v) t[v] = (t[v] || 0) + 1;
      }
      return Object.keys(t).sort((a, b) => t[b] - t[a])[0] || '';
    };
    ValidatieLog.log({
      tool: 'itemlijst-validator',
      deliveryRef: dom(COL.A),
      supplier: dom(COL.K),
      bestand: (document.getElementById('val-filename')?.textContent || '').trim(),
      regels: _valRows.length,
      fouten, waarschuwingen,
    });
  } catch (e) { /* loggen mag de validatie nooit hinderen */ }
}

// ── AFS/leverancier-mailknop laten oplichten bij fouten of afwijkingen ─────
function _valHighlightMailBtn() {
  const btn = document.getElementById('valmail-launch');
  if (!btn) return;
  let hasErr = false, hasDev = false;
  try { hasErr = Array.isArray(_valRows) && _valRows.some(r => r && r.errors && Object.keys(r.errors).length > 0); } catch (e) {}
  try {
    if (window.ValMailer) {
      const afs = (ValMailer._afsItems && ValMailer._afsItems()) || [];
      const pal = (ValMailer._palletItems && ValMailer._palletItems()) || [];
      hasDev = (afs.length > 0) || (pal.length > 0);
    }
  } catch (e) {}
  if (!document.getElementById('valmail-alert-style')) {
    const st = document.createElement('style'); st.id = 'valmail-alert-style';
    st.textContent =
      '#valmail-launch.valmail-alert{background:#D91F2C!important;color:#fff!important;' +
      'box-shadow:0 0 0 2px rgba(217,31,44,.55),0 0 16px rgba(217,31,44,.6)!important;' +
      'animation:valmailPulse 1.5s ease-in-out infinite}' +
      '@keyframes valmailPulse{0%,100%{box-shadow:0 0 0 2px rgba(217,31,44,.5),0 0 10px rgba(217,31,44,.45)}' +
      '50%{box-shadow:0 0 0 3px rgba(217,31,44,.78),0 0 22px rgba(217,31,44,.85)}}';
    document.head.appendChild(st);
  }
  const on = hasErr || hasDev;
  btn.classList.toggle('valmail-alert', on);
  btn.title = on
    ? 'Let op: ' + [hasErr ? 'fouten in de lijst' : '', hasDev ? 'AFS-voormelding nodig (omvang/gewicht/>20 colli)' : ''].filter(Boolean).join(' + ')
    : 'Mail opstellen voor AFS / leverancier';
}

if (document.readyState !== 'loading') _valConnectExpediting();

// ── Run full validation ────────────────────────────────────────────────────
async function runValidation() {
  const dzEl  = document.getElementById('val-dz');
  const tbEl  = document.getElementById('val-tbody');
  const sumEl = document.getElementById('val-summary');
  const shWEl = document.getElementById('val-sheet-warnings');

  if (!_valRows.length) {
    if (sumEl) sumEl.innerHTML = '<span style="color:var(--muted)">Upload eerst een Itemlijst.</span>';
    return;
  }

  const btn = document.getElementById('btn-val-run');
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Valideren…'; }

  // Get USD rate if needed
  const pHeader  = _valHeaders[COL.P] || '';
  const usdPrice = isUSD(pHeader);
  const usdRate  = usdPrice ? await fetchUSDRate() : null;

  // Get country codes from file's Master tab (already loaded)
  const coo = _valCOO.size > 0 ? _valCOO : VL_COO_FALLBACK;

  // Expediting-data uit de centrale koppeling, gefilterd op de gekozen Sub Project ID
  const expeditingData = (typeof _valExpRows === 'function')
    ? _valExpRows()
    : ((typeof fileData !== 'undefined' && fileData.expediting) ? fileData.expediting.data : null);

  // Slim aanvullen vanuit Expediting (H ↔ Unified Reference Code) — alleen lege cellen
  let filledFields = 0, filledRows = 0;
  if (expeditingData && expeditingData.length) {
    for (const row of _valRows) {
      const n = _fillRowFromExpediting(row.cells, expeditingData);
      if (n) { filledFields += n; filledRows++; row._edited = true; }   // meenemen in export
    }
  }

  _valNormalizeCase();   // dropdown-velden hoofdletterongevoelig naar officiële schrijfwijze

  // Inspection Level (AH) vanuit de Admin-regels (alleen lege cellen)
  const _inspRes = await _valApplyInspectionRules();

  // Landnaam → ISO 3166-1 alpha-2 code (Country of Origin, kolom N)
  for (const row of _valRows) {
    const code = _toCountryCode(row.cells[COL.N]);
    if (code && code !== row.cells[COL.N]) { row.cells[COL.N] = code; row._edited = true; }   // meenemen in export
  }

  // Validate each row
  let totalErrors = 0, totalWarnings = 0;
  for (const row of _valRows) {
    const result = await validateRow(row.cells, usdPrice, usdRate, coo, expeditingData);
    row.errors   = result.errors;
    row.warnings = result.warnings;
    row.computed = result.computed;
    totalErrors   += Object.keys(result.errors).length;
    totalWarnings += Object.keys(result.warnings).length;
  }

  // Collonummer ↔ maatvoering-consistentie (kruis-controle over alle regels)
  const colloCheck = _validateColloConsistency(_valRows);
  totalErrors += colloCheck.added;

  // Sheet-level checks
  const sheetWarnings = validateSheet(_valRows).concat(colloCheck.warnings);

  renderValidationTable(usdPrice, usdRate);

  // Summary
  if (sumEl) {
    const cls = totalErrors === 0 ? 'var(--green)' : 'var(--red)';
    sumEl.innerHTML =
      `<span style="color:${cls};font-weight:700">${totalErrors === 0 ? '✅' : '❌'} ${totalErrors} fout(en)</span>` +
      `<span style="color:var(--amber)">  ⚠️ ${totalWarnings} waarschuwing(en)</span>` +
      `<span style="color:var(--muted)">${_valRows.length} rijen gevalideerd</span>` +
      (_inspRes && _inspRes.filled ? `<span style="color:var(--teal)">  🔍 ${_inspRes.filled} inspectieniveau(s) ingevuld in kolom ‘${esc(String(_valHeaders[COL.AH] ?? ''))}’${_inspRes.viaExpSupplier ? ` (waarvan ${_inspRes.viaExpSupplier} via de leveranciersnaam uit Expediting)` : ''}${_inspRes.poMatched != null ? ` · ${_inspRes.poMatched}/${_valRows.length} regels aan Expediting gekoppeld via PO` : ''}</span>` : '') +
      (filledFields ? `<span style="color:var(--teal)">  🔗 ${filledFields} veld(en) aangevuld vanuit Expediting (${filledRows} rij(en))</span>` : '');
  }

  // Sheet warnings
  if (shWEl) {
    shWEl.innerHTML = sheetWarnings.map(w =>
      `<div style="color:var(--amber);font-size:.72rem">⚠️ ${esc(w)}</div>`
    ).join('');
  }

  if (btn) { btn.disabled = false; btn.textContent = '▶ Valideer'; }

  document.getElementById('btn-val-export')?.removeAttribute('disabled');
  document.getElementById('btn-val-labels')?.removeAttribute('disabled');

  // AFS-voormelding: pop-up tonen als een voorwaarde (oversized/zwaar óf >20 colli) geraakt is
  if (window.ValMailer && typeof window.ValMailer.checkPrenotify === 'function') {
    try { window.ValMailer.checkPrenotify(); } catch (e) {}
  }
  if (window.ValDgMail) { try { window.ValDgMail.afterValidation(); } catch (e) { console.warn('ValDgMail:', e); } }   // DG / Dual use = YES → document-aanvraag-popup
  _valHighlightMailBtn();   // AFS/leverancier-mailknop laten oplichten bij fouten of afwijkingen
  _valLogValidatie();       // teambrede telling (fire-and-forget, faalt stil)

  // Live HS-code check against douane.nl nomenclature (async, updates cells in place)
  checkHSCodesLive();
}

// ── Render validation table ────────────────────────────────────────────────
function renderValidationTable(usdPrice, usdRate) {
  const tbEl = document.getElementById('val-tbody');
  if (!tbEl) return;

  // IHC columns shown first so they can be filled in immediately
  // Supplier columns follow — all editable
  const COLS_SHOW = [
    'A','B',          // IHC: Delivery ref, Project
    'C','D','E','F','G','H',
    'I','J',          // Supplier: Supplier article number (verplicht), Serial number (optioneel)
    'K','L','M','N','O','P','Q','R','S','T','U','V',
    'W',              // IHC: Volume (computed)
    'X','Y',
    'AA','AB','AC',   // Dangerous Goods, Dual use, ECCN code
    'AE','AF','AG',   // Container number, Seal number, Remarks supplier
    'AH'              // IHC: Inspection Level
  ];
  const IHC_COLS = new Set(['A','B','W','AH']); // teal tint = IHC to fill
  const VIRT_COLS = new Set(Object.keys(_virtualCols));  // amber tint = IHC-added column

  const html = _valRows.map((row, ri) => {
    const hasDG  = _valBool(row.cells[COL.AA]);
    const anyErr = Object.keys(row.errors).length > 0;
    const anyWrn = Object.keys(row.warnings).length > 0;
    const rowCls = hasDG ? 'val-row-dg' : anyErr ? 'val-row-err' : anyWrn ? 'val-row-warn' : 'val-row-ok';

    const cells = COLS_SHOW.map(col => {
      const ci  = COL[col];
      const val = row.cells[ci];
      const err = row.errors[col];
      const wrn = row.warnings[col];
      const cmp = row.computed?.[col] ?? row.computed?.[col+'_EUR'];
      const disp = val !== null && val !== undefined ? String(val) : '';

      // Kolombreedte op inhoud: de input schaalt mee via het `size`-attribuut.
      // Omschrijving (E) krijgt een vaste breedte en schuift intern (zoals bij
      // eerdere kolommen toegepast); de overige kolommen worden zo smal als nodig.
      const isDesc = (col === 'E');
      const inSize = isDesc ? 30 : Math.min(Math.max(disp.length, 4), 48);

      let cellCls = IHC_COLS.has(col) ? 'val-cell-ihc' : '';
      if (VIRT_COLS.has(col)) cellCls += ' val-cell-virtual';
      if (err) cellCls += ' val-cell-err';
      else if (wrn) cellCls += ' val-cell-warn';
      if (isDesc) cellCls += ' val-cell-desc';

      const tooltip = err || wrn || (cmp ? `Berekend: ${cmp}` : '');
      const tAttr   = tooltip ? `title="${esc(tooltip)}"` : '';

      // HS-code cell: format indicator + live-check icon (updated async)
      if (col === 'O') {
        const fmtOk  = !err && disp;
        const fmtIcon = !disp ? '' : fmtOk
          ? `<span class="hs-icon" style="color:var(--muted);font-size:.65rem;margin-right:.25rem">⏳</span>`
          : `<span class="hs-icon" style="color:#ef4444;font-weight:700;margin-right:.25rem">✗</span>`;
        return `<td class="val-cell ${cellCls}" ${tAttr}>
          <div style="display:flex;align-items:center;gap:.2rem">
            ${fmtIcon}
            <input class="val-input" data-row="${ri}" data-col="${ci}"
              value="${esc(disp)}" size="${Math.max(disp.length, 10)}"
              oninput="valCellEdit(${ri},${ci},this.value)">
          </div>
        </td>`;
      }

      // Dangerous Goods cell — checkbox, not text input
      if (col === 'AA') {
        return `<td class="val-cell ${hasDG ? 'val-cell-dg' : ''}" ${tAttr} style="text-align:center">
          <input type="checkbox" ${hasDG ? 'checked' : ''}
            onchange="valCellEdit(${ri},${ci},this.checked)">
          ${hasDG ? ' 🔴' : ''}
        </td>`;
      }

      // Dual use cell — checkbox, not text input (ECCN code in AC is required if checked)
      if (col === 'AB') {
        const hasDU = _valBool(val);
        return `<td class="val-cell ${cellCls}" ${tAttr} style="text-align:center">
          <input type="checkbox" ${hasDU ? 'checked' : ''}
            onchange="valCellEdit(${ri},${ci},this.checked)">
          ${hasDU ? ' ⚠️' : ''}
        </td>`;
      }

      // USD value — show EUR conversion below
      if (col === 'P' && usdPrice && cmp) {
        return `<td class="val-cell ${cellCls}" ${tAttr}>
          <input class="val-input" data-row="${ri}" data-col="${ci}"
            value="${esc(disp)}" size="${inSize}" oninput="valCellEdit(${ri},${ci},this.value)">
          <div style="font-size:.6rem;color:var(--teal);margin-top:.1rem">≈ ${Number(cmp).toLocaleString('nl-NL')} EUR</div>
        </td>`;
      }

      // Inspection Level (AH) — keuzelijst (dropdown) met de 6 vaste opties
      if (col === 'AH') {
        const cur = disp;
        const matched = INSP_OPTIONS.find(o => o.toLowerCase() === cur.toLowerCase());
        const optsHtml = INSP_OPTIONS.map(o =>
          `<option value="${esc(o)}" ${matched === o ? 'selected' : ''}>${esc(o)}</option>`
        ).join('');
        // Onbekende bestaande waarde tóch tonen zodat niets stil verdwijnt
        const strayOpt = (cur && !matched)
          ? `<option value="${esc(cur)}" selected>${esc(cur)} (onbekend)</option>` : '';
        return `<td class="val-cell ${cellCls}" ${tAttr}>
          <select class="val-input val-select" data-row="${ri}" data-col="${ci}"
            onchange="valCellEdit(${ri},${ci},this.value)">
            <option value="" ${!cur ? 'selected' : ''}>—</option>
            ${strayOpt}${optsHtml}
          </select>
        </td>`;
      }

      // All other cells — editable input, pre-filled with current value
      return `<td class="val-cell ${cellCls}" ${tAttr}>
        <input class="val-input" data-row="${ri}" data-col="${ci}"
          value="${esc(disp)}" size="${inSize}"
          placeholder="${esc(_valHeaders[ci]||col)}"
          oninput="valCellEdit(${ri},${ci},this.value)">
      </td>`;
    }).join('');

    const rowStatus = anyErr ? '❌' : anyWrn ? '⚠️' : '✅';
    return `<tr class="val-row ${rowCls}">
      <td class="val-cell val-cell-num">${ri+1}</td>
      <td class="val-cell" style="text-align:center">${rowStatus}</td>
      ${cells}
    </tr>`;
  }).join('');

  tbEl.innerHTML = html || '<tr><td colspan="30" style="text-align:center;color:var(--muted);padding:1.5rem">Geen data</td></tr>';
}

// ── Cell edit (IHC write-back) ─────────────────────────────────────────────
function valCellEdit(rowIdx, colIdx, value) {
  if (_valRows[rowIdx]) {
    _valRows[rowIdx].cells[colIdx] = value;
    _valRows[rowIdx]._edited = true;
  }
}

// ── Export corrected Itemlijst ─────────────────────────────────────────────
function exportValidatedItemlijst() {
  if (!_valWb) { alert('Laad eerst een Itemlijst.'); return; }
  const ws = _valWb.Sheets[_valSheetName || _valPickItemSheet(_valWb)];

  // Data rows in worksheet start at: hdrIdx row (0-based) + 1 header row = hdrIdx+1 (0-based)
  // In 1-based row index: _valHdrIdx + 2  (hdrIdx is 0-based, +1 for header, +1 for 1-based)
  const dataRowStart = _valHdrIdx + 1; // 0-based row index of first data row

  // Write virtual column headers into the worksheet if they are IHC-added
  const virtualEntries = Object.entries(_virtualCols);
  if (virtualEntries.length > 0) {
    virtualEntries.forEach(([letter]) => {
      const ci = COL[letter];
      // Write owner row label (if there is one)
      if (_valHdrIdx > 0) {
        const ownerAddr = XLSX.utils.encode_cell({ r: _valHdrIdx - 1, c: ci });
        ws[ownerAddr] = { v: 'IHC', t: 's' };
      }
      // Write column header
      const hdrAddr = XLSX.utils.encode_cell({ r: _valHdrIdx, c: ci });
      ws[hdrAddr] = { v: _valHeaders[ci] || letter, t: 's' };
    });
    // Expand worksheet range to include new columns
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    const maxVirtCol = Math.max(...Object.values(_virtualCols).map
      ? Object.keys(_virtualCols).map(l => COL[l])
      : [range.e.c]);
    if (maxVirtCol > range.e.c) range.e.c = maxVirtCol;
    ws['!ref'] = XLSX.utils.encode_range(range);
  }

  // Kolommen die een schone numerieke waarde horen te zijn (zie validateRow —
  // daar staat voor elk van deze de opgeschoonde waarde al klaar in
  // row.computed, zonder valutateken/eenheid/tekst). Deze worden bij export
  // ALTIJD als echt Excel-getal teruggeschreven, ook als de rij verder niet
  // bewerkt is — anders bleef een origineel "€ 1.250,00"-tekstveld in
  // onbewerkte rijen gewoon staan zoals het was. Twee decimalen voor
  // geldbedragen en gewicht (conform de template-instructie), de rest vrij.
  const NUMERIC_COLS  = new Set(['F','P','Q','T','U','V','X','Y']);
  const TWO_DEC_COLS  = new Set(['P','Q','X','Y']);

  // Write back edited/computed values
  _valRows.forEach((row, ri) => {
    const wsRowIdx = dataRowStart + ri; // 0-based
    // Edited cells (including virtual column values)
    Object.entries(COL).forEach(([colLetter, colIdx]) => {
      const isVirtual   = _virtualCols[colLetter];
      const isNumericCol = NUMERIC_COLS.has(colLetter);
      const cleanNum     = row.computed ? row.computed[colLetter] : undefined;
      const useClean     = isNumericCol && cleanNum !== undefined;
      let val = useClean ? cleanNum : row.cells[colIdx];
      if (val === undefined || val === null || val === '') return;
      // Kolom D: altijd in het vaste formaat '-1-1 wegschrijven (ook bij onbewerkte rijen)
      let fmtD = false;
      if (colLetter === 'D') { const nv = _valItemFmt(val); if (nv !== val) { val = nv; fmtD = true; } }
      // Numerieke kolommen altijd herschrijven (ook als de rij verder niet
      // bewerkt is); overige kolommen zoals voorheen alleen bij edit/virtueel.
      if (!(useClean || row._edited || isVirtual || fmtD)) return;
      const addr = XLSX.utils.encode_cell({ r: wsRowIdx, c: colIdx });
      if (!ws[addr]) ws[addr] = {};
      ws[addr].v = val;
      ws[addr].t = useClean ? 'n' : (typeof val === 'boolean' ? 'b' : typeof val === 'number' ? 'n' : 's');
      if (useClean) {
        ws[addr].z = TWO_DEC_COLS.has(colLetter) ? '0.00' : 'General';
      } else {
        delete ws[addr].z; // geen oude valuta-/tekstopmaak laten hangen op een niet-numerieke herschrijving
      }
    });
    // Computed volume W
    if (row.computed?.W !== undefined) {
      const addr = XLSX.utils.encode_cell({ r: wsRowIdx, c: COL.W });
      ws[addr] = { v: row.computed.W, t: 'n' };
    }
  });

  // Add validation summary to a new sheet
  const summaryData = [
    ['Rij', 'IHC PO', 'Item', 'Fouten', 'Waarschuwingen'],
    ..._valRows.map((row, i) => [
      i+1,
      row.cells[COL.C] || '',
      row.cells[COL.D] || '',
      Object.entries(row.errors).map(([k,v]) => `${k}: ${v}`).join(' | '),
      Object.entries(row.warnings).map(([k,v]) => `${k}: ${v}`).join(' | '),
    ])
  ];
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(_valWb, wsSummary, 'Validatie rapport');

  XLSX.writeFile(_valWb, _valExportBaseName() + '.xlsx');
}


// ── Build dynamic table header from loaded file ─────────────────────────
function buildValHeader() {
  const thEl = document.getElementById('val-thead');
  if (!thEl) return;
  const COLS_SHOW = [
    'A','B',
    'C','D','E','F','G','H',
    'I','J',
    'K','L','M','N','O','P','Q','R','S','T','U','V',
    'W','X','Y',
    'AA','AB','AC','AE','AF','AG','AH'
  ];
  const IHC_OWN   = new Set(['A','B','W','AH']);
  const ownerRow  = _valOwners;
  const headerRow = _valHeaders;
  const VIRT_COLS = new Set(Object.keys(_virtualCols));  // IHC-toegevoegde kolommen (bv. AA/AH indien niet in bestand)

  const ownerMap  = {};
  COLS_SHOW.forEach(col => {
    ownerMap[col] = VIRT_COLS.has(col) ? 'IHC ✚' : (ownerRow[COL[col]] || '');
  });

  const th1 = `<tr>
    <th rowspan="2" style="width:30px">#</th>
    <th rowspan="2">Status</th>
    ${COLS_SHOW.map(col => {
      const owner = ownerMap[col];
      const cls   = /IHC/.test(owner) ? 'col-owner-ihc' : 'col-owner-sup';
      return `<th class="${cls}">${esc(owner||'')}</th>`;
    }).join('')}
  </tr>`;
  const th2 = `<tr>
    ${COLS_SHOW.map(col => {
      const hdr = (headerRow[COL[col]] || col) + (VIRT_COLS.has(col) ? ' ✚' : '');
      const cls = IHC_OWN.has(col) ? 'col-owner-ihc' : VIRT_COLS.has(col) ? 'col-owner-virtual' : '';
      const fill = _FILL_COLS.has(col);
      const mark = fill ? '<span class="val-fill-mark" title="Automatisch aangevuld vanuit de centrale Expediting-lijst">🔗</span>' : '';
      const tip = fill ? 'Automatisch aangevuld vanuit Expediting' : (VIRT_COLS.has(col) ? 'Door IHC in te vullen vóór export' : hdr);
      return `<th class="${cls}" title="${esc(tip)}">${esc(hdr)}${mark}</th>`;
    }).join('')}
  </tr>`;
  thEl.innerHTML = th1 + th2;
}

// ── Itemlijst-blad kiezen ──────────────────────────────────────────────────
// De nieuwe Itemlijst-template bevat naast het invulblad ook een blad
// "Manual" (uitleg voor leveranciers), naast Master/Broncontrole. Niet blindelings
// blad 1 lezen: bladen met een instructie-/hulpnaam worden overgeslagen en de
// rest wordt beoordeeld op de aanwezigheid van de verwachte kolomkoppen in de
// eerste rijen. Het blad met de meeste treffers wint; bij gelijkspel het eerste.
let _valSheetName = null;
const _VAL_SKIP_SHEETS = /^(manual|handleiding|instruct|toelichting|uitleg|readme|master|broncontrole|validatie)/i;
const _VAL_HDR_HINTS = [/delivery\s*ref/i, /^project$/i, /ihc\s*po/i, /^item$/i, /item\s*desc/i, /quantity/i,
  /unit\s*of\s*measure/i, /component|mark/i, /supplier/i, /country\s*of\s*origin/i, /hs[-\s]?code/i,
  /value\s*pc/i, /value\s*total/i, /type\s*of\s*packag/i, /length/i, /width/i, /heig/i, /gross/i, /nett?/i];
function _valPickItemSheet(wb) {
  const names = wb.SheetNames || [];
  const score = (n) => {
    const ws = wb.Sheets[n]; if (!ws || !ws['!ref']) return 0;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false }).slice(0, 7);
    let best = 0;
    rows.forEach(r => {
      const cells = (r || []).map(c => String(c == null ? '' : c).trim()).filter(Boolean);
      const hits = _VAL_HDR_HINTS.filter(re => cells.some(c => re.test(c))).length;
      if (hits > best) best = hits;
    });
    return best;
  };
  const candidates = names.filter(n => !_VAL_SKIP_SHEETS.test(n.trim()));
  const pool = candidates.length ? candidates : names;
  let pick = null, top = -1;
  pool.forEach(n => { const sc = score(n); if (sc > top) { top = sc; pick = n; } });
  return pick || names[0];
}

// ── Load Itemlijst file ────────────────────────────────────────────────────
function handleValFile(fileOrEvent) {
  // Accept a File object directly (new), or fall back to legacy event
  const file = (fileOrEvent instanceof File)
    ? fileOrEvent
    : (fileOrEvent?.target?.files?.[0] ?? fileOrEvent?.dataTransfer?.files?.[0]);
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
   try {
    const wb = XLSX.read(e.target.result, { type: 'array', cellDates: true });
    _valWb = wb;
    const wsName = _valPickItemSheet(wb);
    _valSheetName = wsName;
    const ws = wb.Sheets[wsName];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: false });

    const hasVal = (v) => {
      if (v === null || v === undefined) return false;
      if (typeof v === 'boolean') return v === true;
      const s = String(v).trim();
      if (!s || s === '0') return false;
      if (s.startsWith('=')) return false;
      return true;
    };

    // ── Dynamic header row detection ─────────────────────────────────────
    // Scan rows 0–6: de koprij wordt bepaald op het aantal UNIEKE niet-lege
    // waarden, niet het ruwe aantal niet-lege cellen. Een "eigenaar"-rij
    // (bv. 'Supplier'/'IHC' herhaald over tientallen kolommen) kan evenveel
    // gevulde cellen hebben als de echte koprij, maar heeft veel minder
    // UNIEKE waarden — elke kolom in een echte koprij heeft immers een eigen
    // naam. Dit voorkomt dat de eigenaar-rij de koprij wint bij een
    // gelijke-stand op ruw aantal (zoals eerder gebeurde: 27 gevuld in
    // zowel de eigenaar-rij als de echte koprij, > koos dan de eerste).
    let hdrIdx = 0;
    let hdrMax = -1;
    const scanLimit = Math.min(raw.length, 7);
    for (let i = 0; i < scanLimit; i++) {
      const cells = (raw[i] || []).filter(c => c !== null && c !== undefined && String(c).trim() !== '');
      const uniq = new Set(cells.map(c => String(c).trim().toLowerCase()));
      if (uniq.size > hdrMax) { hdrMax = uniq.size; hdrIdx = i; }
    }
    _valHdrIdx  = hdrIdx;
    _valOwners  = hdrIdx > 0 ? (raw[hdrIdx - 1] || []) : [];
    _valHeaders = raw[hdrIdx] || [];

    // ── Remap COL to actual column positions by header name ───────────────
    // Also add virtual columns for DG / Inspection if missing from the file
    _remapColumns();

    _valRows = (raw.slice(hdrIdx + 1) || [])
      // FIX: accept rows where col A (Delivery ref.) is filled AND either
      // col C (IHC PO) OR col H (Component/Mark) is present.
      // Col C may be empty before cross-referencing with the Expediting list
      // (val-crossref.js fills C and D in afterwards via the modal).
      .filter(r => hasVal(r[0]) && (hasVal(r[2]) || hasVal(r[7])))
      .map(r => ({ cells: r, errors: {}, warnings: {}, computed: {} }));

    // YES/NO → TRUE/FALSE voor Dangerous Goods (AA) en Dual use (AB)
    _valNormalizeBoolCols();
    _valNormalizeCase();

    // HS-code (kolom O) altijd op 10 cijfers brengen — bv. een 12-cijferige
    // TARIC-code uit het bestand verliest de laatste 2 cijfers.
    for (const row of _valRows) {
      const orig = row.cells[COL.O];
      const norm = _normalizeHSCode(orig);
      if (norm !== orig) {
        row.cells[COL.O] = norm;
        row._edited = true;
        // Bewaar het origineel zodat validateRow de stille opschoning
        // (spaties/punten weg, of >8 cijfers ingekort) als waarschuwing kan
        // tonen i.p.v. onopgemerkt te laten verdwijnen.
        row.cells._hsOrigRaw = orig;
      }
    }

    // Load country codes from Master tab
    _valCOO = new Set();
    if (wb.Sheets['Master']) {
      const mRaw = XLSX.utils.sheet_to_json(wb.Sheets['Master'], { header: 1, defval: null });
      mRaw.forEach(r => {
        if (r[0] && /^[A-Z]{2}$/.test(String(r[0]).trim()))
          _valCOO.add(String(r[0]).trim().toUpperCase());
      });
    }

    const fn = document.getElementById('val-filename');
    const dz = document.getElementById('val-dz');
    if (dz) dz.classList.add('loaded');
    buildValHeader();
    document.getElementById('btn-val-run')?.removeAttribute('disabled');
    document.getElementById('btn-val-labels')?.removeAttribute('disabled');

    // ── Val-crossref: vul C (IHC PO) en D (Item) in via Expediting lijst ──
    // Loopt VOOR de samenvatting zodat de gebruiker de complete rijen ziet.
    function _afterCrossref() {
      if (fn) { fn.textContent = `${file.name} — ${_valRows.length} rijen`; fn.style.display = 'block'; }
      const sumEl0 = document.getElementById('val-summary');
      if (sumEl0) sumEl0.innerHTML =
        `<span style="color:var(--muted)">${_valRows.length} rijen geladen — klik Valideer om te starten</span>`;
    }

    if (window.ValCrossref) {
      // Zet _valRows om naar objecten die ValCrossref begrijpt
      const xrefRows = _valRows.map(row => ({
        'Delivery ref.'         : row.cells[COL.A] ?? '',
        'Project'               : row.cells[COL.B] ?? '',
        'IHC PO'                : row.cells[COL.C] ?? '',
        'Item'                  : row.cells[COL.D] ?? '',
        'Item description'      : row.cells[COL.E] ?? '',
        'Quantity'              : row.cells[COL.F] ?? '',
        'Unit of measure'       : row.cells[COL.G] ?? '',
        'Component (Mark/Label)': row.cells[COL.H] ?? '',
        'Supplier'              : row.cells[COL.K] ?? '',
      }));

      ValCrossref.runIfNeeded(xrefRows, function(enriched) {
        // Kopieer aangevulde waarden terug naar de cells-array
        enriched.forEach((obj, i) => {
          if (!_valRows[i]) return;
          const c = _valRows[i].cells;
          // Alleen lege cellen overschrijven; markeer als bewerkt voor export
          const was = { c: c[COL.C], d: c[COL.D] };
          if (!c[COL.C] && obj['IHC PO'])  c[COL.C] = obj['IHC PO'];
          if (!c[COL.D] && obj['Item'])     c[COL.D] = obj['Item'];
          if (!c[COL.E] && obj['Item description']) c[COL.E] = obj['Item description'];
          if (!c[COL.G] && obj['Unit of measure'])  c[COL.G] = obj['Unit of measure'];
          if (!c[COL.K] && obj['Supplier']) c[COL.K] = obj['Supplier'];
          if (c[COL.C] !== was.c || c[COL.D] !== was.d) _valRows[i]._edited = true;
        });
        _afterCrossref();
      });
    } else {
      _afterCrossref();
    }
   } catch (err) {
     console.error('Itemlijst inlezen mislukt:', err);
     const sumErr = document.getElementById('val-summary');
     if (sumErr) sumErr.innerHTML =
       `<span style="color:#ef4444">❌ Inlezen mislukt: ${esc(err.message)}</span>`;
   }
  };
  reader.readAsArrayBuffer(file);
}

// Drag & drop wiring
function valDragOver(e) { e.preventDefault(); e.stopPropagation(); document.getElementById('val-dz')?.classList.add('dz-hover'); }
function valDragLeave(e) { e.preventDefault(); document.getElementById('val-dz')?.classList.remove('dz-hover'); }
function valDrop(e) {
  e.preventDefault(); e.stopPropagation();
  document.getElementById('val-dz')?.classList.remove('dz-hover');
  const f = e.dataTransfer?.files?.[0];
  if (f) handleValFile(f);
}

// ── Merken / Labels Generator ────────────────────────────────────────────
// Kolom A bevat per rij het unieke pakketnummer zoals dat voor verzending
// is vastgesteld (bv. "128-001", "128-002", ...). Rijen met hetzelfde
// pakketnummer in kolom A horen bij hetzelfde pakket/label — er wordt dus
// simpel per unieke waarde van kolom A gegroepeerd (geen losse collo-telling
// meer nodig, kolom A ÍS het pakketnummer).

// ── Bouw labeldata, gegroepeerd per uniek pakketnummer (kolom A) ──────────
function buildLabelData() {
  const labels = {};
  const order  = [];
  const shipment = document.getElementById('cipl-shipment-sel')?.value || '';

  _valRows.forEach(row => {
    const pkgNr = String(row.cells[COL.A] || '').trim();
    if (!pkgNr) return;

    if (!labels[pkgNr]) {
      // Delivery Ref = het gedeelte vóór het eerste streepje van het
      // pakketnummer (bv. "128" uit "128-001"); geen streepje → heel nummer.
      const dashIx = pkgNr.indexOf('-');
      const deliveryRef = dashIx > -1 ? pkgNr.slice(0, dashIx) : pkgNr;

      labels[pkgNr] = {
        packageNr:   pkgNr,
        deliveryRef,
        project:     String(row.cells[COL.B]  || '').trim(),
        pkgType:     String(row.cells[COL.S]  || '').trim(),
        gross:       String(row.cells[COL.X]  || '').trim(),
        nett:        String(row.cells[COL.Y]  || '').trim(),
        length:      String(row.cells[COL.T]  || '').trim(),
        width:       String(row.cells[COL.U]  || '').trim(),
        height:      String(row.cells[COL.V]  || '').trim(),
        remarks:     String(row.cells[COL.AG] || '').trim(),
        items:       [],
        shipment,
      };
      order.push(pkgNr);
    }

    const lb = labels[pkgNr];
    // sommige velden kunnen op een latere rij van hetzelfde pakket pas gevuld zijn
    if (!lb.remarks) lb.remarks = String(row.cells[COL.AG] || '').trim();
    if (!lb.pkgType) lb.pkgType = String(row.cells[COL.S]  || '').trim();

    const hVal = String(row.cells[COL.H] || '').trim();
    if (hVal) {
      parseHColumn(hVal).forEach(code => {
        if (!lb.items.includes(code)) lb.items.push(code);
      });
    }
  });

  // Natuurlijke sortering op pakketnummer: 128-002 vóór 128-010 (niet alfabetisch)
  return order.map(k => labels[k]).sort((a, b) =>
    a.packageNr.localeCompare(b.packageNr, undefined, { numeric: true, sensitivity: 'base' })
  );
}

// ── QR-inhoud: leesbare samenvatting van alle gegevens die ook op het
// label zelf staan — zo kan een scanner (of mens, via elke QR-app) de
// pakketgegevens aflezen zonder het Itemlijst-systeem te hoeven raadplegen.
function _buildQrPayload(lb, consignee) {
  const itemDisp = lb.items.length === 0   ? '—'
    : lb.items.length <= 3 ? lb.items.join(', ')
    : `CIPL #${lb.shipment || '—'}`;
  const measDisp = (lb.length && lb.width && lb.height)
    ? `${lb.length} x ${lb.width} x ${lb.height} CM`
    : '—';
  return [
    `PACKAGE: ${lb.packageNr}`,
    `PROJECT: ${lb.project || '—'}`,
    `SHIPPING MARKS: ${consignee}`,
    `ITEM NUMBER: ${itemDisp}`,
    `TYPE OF PACKAGE: ${lb.pkgType || '—'}`,
    `GROSS WEIGHT: ${lb.gross || '—'} KG`,
    `NETT WEIGHT: ${lb.nett || '—'} KG`,
    `MEASUREMENT: ${measDisp}`,
    `DELIVERY REF: ${lb.deliveryRef || '—'}`,
    `REMARKS: ${lb.remarks || '—'}`,
  ].join('\n');
}

// ── QR-afbeelding (data-URI GIF) opbouwen via de ingebakken qrcode-generator
// library (shared/qrcode.min.js, geen netwerk nodig — werkt ook offline/print).
// ── Build-markering: zo is te zien welke versie van dit bestand de browser echt draait ──
const VAL_BUILD = '2026-10-07-insp-guard-v6';
console.info('[validator.js] build ' + VAL_BUILD);

// ── QR-code bibliotheek (qrcode-generator, MIT) — ingebouwd zodat er geen los bestand nodig is ──
const _QRLIB = (function () {
/* qrcode-generator v2.0.4 — Kazuhiko Arase — MIT License — https://github.com/kazuhikoarase/qrcode-generator */
var qrcode=function(){var t=function(t,r){var e=t,n=g[r],o=null,i=0,a=null,u=[],f={},c=function(t,r){o=function(t){for(var r=new Array(t),e=0;e<t;e+=1){r[e]=new Array(t);for(var n=0;n<t;n+=1)r[e][n]=null}return r}(i=4*e+17),l(0,0),l(i-7,0),l(0,i-7),s(),h(),d(t,r),e>=7&&v(t),null==a&&(a=p(e,n,u)),w(a,r)},l=function(t,r){for(var e=-1;e<=7;e+=1)if(!(t+e<=-1||i<=t+e))for(var n=-1;n<=7;n+=1)r+n<=-1||i<=r+n||(o[t+e][r+n]=0<=e&&e<=6&&(0==n||6==n)||0<=n&&n<=6&&(0==e||6==e)||2<=e&&e<=4&&2<=n&&n<=4)},h=function(){for(var t=8;t<i-8;t+=1)null==o[t][6]&&(o[t][6]=t%2==0);for(var r=8;r<i-8;r+=1)null==o[6][r]&&(o[6][r]=r%2==0)},s=function(){for(var t=B.getPatternPosition(e),r=0;r<t.length;r+=1)for(var n=0;n<t.length;n+=1){var i=t[r],a=t[n];if(null==o[i][a])for(var u=-2;u<=2;u+=1)for(var f=-2;f<=2;f+=1)o[i+u][a+f]=-2==u||2==u||-2==f||2==f||0==u&&0==f}},v=function(t){for(var r=B.getBCHTypeNumber(e),n=0;n<18;n+=1){var a=!t&&1==(r>>n&1);o[Math.floor(n/3)][n%3+i-8-3]=a}for(n=0;n<18;n+=1){a=!t&&1==(r>>n&1);o[n%3+i-8-3][Math.floor(n/3)]=a}},d=function(t,r){for(var e=n<<3|r,a=B.getBCHTypeInfo(e),u=0;u<15;u+=1){var f=!t&&1==(a>>u&1);u<6?o[u][8]=f:u<8?o[u+1][8]=f:o[i-15+u][8]=f}for(u=0;u<15;u+=1){f=!t&&1==(a>>u&1);u<8?o[8][i-u-1]=f:u<9?o[8][15-u-1+1]=f:o[8][15-u-1]=f}o[i-8][8]=!t},w=function(t,r){for(var e=-1,n=i-1,a=7,u=0,f=B.getMaskFunction(r),c=i-1;c>0;c-=2)for(6==c&&(c-=1);;){for(var g=0;g<2;g+=1)if(null==o[n][c-g]){var l=!1;u<t.length&&(l=1==(t[u]>>>a&1)),f(n,c-g)&&(l=!l),o[n][c-g]=l,-1==(a-=1)&&(u+=1,a=7)}if((n+=e)<0||i<=n){n-=e,e=-e;break}}},p=function(t,r,e){for(var n=A.getRSBlocks(t,r),o=b(),i=0;i<e.length;i+=1){var a=e[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var u=0;for(i=0;i<n.length;i+=1)u+=n[i].dataCount;if(o.getLengthInBits()>8*u)throw"code length overflow. ("+o.getLengthInBits()+">"+8*u+")";for(o.getLengthInBits()+4<=8*u&&o.put(0,4);o.getLengthInBits()%8!=0;)o.putBit(!1);for(;!(o.getLengthInBits()>=8*u||(o.put(236,8),o.getLengthInBits()>=8*u));)o.put(17,8);return function(t,r){for(var e=0,n=0,o=0,i=new Array(r.length),a=new Array(r.length),u=0;u<r.length;u+=1){var f=r[u].dataCount,c=r[u].totalCount-f;n=Math.max(n,f),o=Math.max(o,c),i[u]=new Array(f);for(var g=0;g<i[u].length;g+=1)i[u][g]=255&t.getBuffer()[g+e];e+=f;var l=B.getErrorCorrectPolynomial(c),h=k(i[u],l.getLength()-1).mod(l);for(a[u]=new Array(l.getLength()-1),g=0;g<a[u].length;g+=1){var s=g+h.getLength()-a[u].length;a[u][g]=s>=0?h.getAt(s):0}}var v=0;for(g=0;g<r.length;g+=1)v+=r[g].totalCount;var d=new Array(v),w=0;for(g=0;g<n;g+=1)for(u=0;u<r.length;u+=1)g<i[u].length&&(d[w]=i[u][g],w+=1);for(g=0;g<o;g+=1)for(u=0;u<r.length;u+=1)g<a[u].length&&(d[w]=a[u][g],w+=1);return d}(o,n)};f.addData=function(t,r){var e=null;switch(r=r||"Byte"){case"Numeric":e=M(t);break;case"Alphanumeric":e=x(t);break;case"Byte":e=m(t);break;case"Kanji":e=L(t);break;default:throw"mode:"+r}u.push(e),a=null},f.isDark=function(t,r){if(t<0||i<=t||r<0||i<=r)throw t+","+r;return o[t][r]},f.getModuleCount=function(){return i},f.make=function(){if(e<1){for(var t=1;t<40;t++){for(var r=A.getRSBlocks(t,n),o=b(),i=0;i<u.length;i++){var a=u[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var g=0;for(i=0;i<r.length;i++)g+=r[i].dataCount;if(o.getLengthInBits()<=8*g)break}e=t}c(!1,function(){for(var t=0,r=0,e=0;e<8;e+=1){c(!0,e);var n=B.getLostPoint(f);(0==e||t>n)&&(t=n,r=e)}return r}())},f.createTableTag=function(t,r){t=t||2;var e="";e+='<table style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: "+(r=void 0===r?4*t:r)+"px;",e+='">',e+="<tbody>";for(var n=0;n<f.getModuleCount();n+=1){e+="<tr>";for(var o=0;o<f.getModuleCount();o+=1)e+='<td style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: 0px;",e+=" width: "+t+"px;",e+=" height: "+t+"px;",e+=" background-color: ",e+=f.isDark(n,o)?"#000000":"#ffffff",e+=";",e+='"/>';e+="</tr>"}return e+="</tbody>",e+="</table>"},f.createSvgTag=function(t,r,e,n){var o={};"object"==typeof arguments[0]&&(t=(o=arguments[0]).cellSize,r=o.margin,e=o.alt,n=o.title),t=t||2,r=void 0===r?4*t:r,(e="string"==typeof e?{text:e}:e||{}).text=e.text||null,e.id=e.text?e.id||"qrcode-description":null,(n="string"==typeof n?{text:n}:n||{}).text=n.text||null,n.id=n.text?n.id||"qrcode-title":null;var i,a,u,c,g=f.getModuleCount()*t+2*r,l="";for(c="l"+t+",0 0,"+t+" -"+t+",0 0,-"+t+"z ",l+='<svg version="1.1" xmlns="http://www.w3.org/2000/svg"',l+=o.scalable?"":' width="'+g+'px" height="'+g+'px"',l+=' viewBox="0 0 '+g+" "+g+'" ',l+=' preserveAspectRatio="xMinYMin meet"',l+=n.text||e.text?' role="img" aria-labelledby="'+y([n.id,e.id].join(" ").trim())+'"':"",l+=">",l+=n.text?'<title id="'+y(n.id)+'">'+y(n.text)+"</title>":"",l+=e.text?'<description id="'+y(e.id)+'">'+y(e.text)+"</description>":"",l+='<rect width="100%" height="100%" fill="white" cx="0" cy="0"/>',l+='<path d="',a=0;a<f.getModuleCount();a+=1)for(u=a*t+r,i=0;i<f.getModuleCount();i+=1)f.isDark(a,i)&&(l+="M"+(i*t+r)+","+u+c);return l+='" stroke="transparent" fill="black"/>',l+="</svg>"},f.createDataURL=function(t,r){t=t||2,r=void 0===r?4*t:r;var e=f.getModuleCount()*t+2*r,n=r,o=e-r;return T(e,e,function(r,e){if(n<=r&&r<o&&n<=e&&e<o){var i=Math.floor((r-n)/t),a=Math.floor((e-n)/t);return f.isDark(a,i)?0:1}return 1})},f.createImgTag=function(t,r,e){t=t||2,r=void 0===r?4*t:r;var n=f.getModuleCount()*t+2*r,o="";return o+="<img",o+=' src="',o+=f.createDataURL(t,r),o+='"',o+=' width="',o+=n,o+='"',o+=' height="',o+=n,o+='"',e&&(o+=' alt="',o+=y(e),o+='"'),o+="/>"};var y=function(t){for(var r="",e=0;e<t.length;e+=1){var n=t.charAt(e);switch(n){case"<":r+="&lt;";break;case">":r+="&gt;";break;case"&":r+="&amp;";break;case'"':r+="&quot;";break;default:r+=n}}return r};return f.createASCII=function(t,r){if((t=t||1)<2)return function(t){t=void 0===t?2:t;var r,e,n,o,i,a=1*f.getModuleCount()+2*t,u=t,c=a-t,g={"██":"█","█ ":"▀"," █":"▄","  ":" "},l={"██":"▀","█ ":"▀"," █":" ","  ":" "},h="";for(r=0;r<a;r+=2){for(n=Math.floor((r-u)/1),o=Math.floor((r+1-u)/1),e=0;e<a;e+=1)i="█",u<=e&&e<c&&u<=r&&r<c&&f.isDark(n,Math.floor((e-u)/1))&&(i=" "),u<=e&&e<c&&u<=r+1&&r+1<c&&f.isDark(o,Math.floor((e-u)/1))?i+=" ":i+="█",h+=t<1&&r+1>=c?l[i]:g[i];h+="\n"}return a%2&&t>0?h.substring(0,h.length-a-1)+Array(a+1).join("▀"):h.substring(0,h.length-1)}(r);t-=1,r=void 0===r?2*t:r;var e,n,o,i,a=f.getModuleCount()*t+2*r,u=r,c=a-r,g=Array(t+1).join("██"),l=Array(t+1).join("  "),h="",s="";for(e=0;e<a;e+=1){for(o=Math.floor((e-u)/t),s="",n=0;n<a;n+=1)i=1,u<=n&&n<c&&u<=e&&e<c&&f.isDark(o,Math.floor((n-u)/t))&&(i=0),s+=i?g:l;for(o=0;o<t;o+=1)h+=s+"\n"}return h.substring(0,h.length-1)},f.renderTo2dContext=function(t,r){r=r||2;for(var e=f.getModuleCount(),n=0;n<e;n++)for(var o=0;o<e;o++)t.fillStyle=f.isDark(n,o)?"black":"white",t.fillRect(o*r,n*r,r,r)},f};t.stringToBytes=(t.stringToBytesFuncs={default:function(t){for(var r=[],e=0;e<t.length;e+=1){var n=t.charCodeAt(e);r.push(255&n)}return r}}).default,t.createStringToBytes=function(t,r){var e=function(){for(var e=S(t),n=function(){var t=e.read();if(-1==t)throw"eof";return t},o=0,i={};;){var a=e.read();if(-1==a)break;var u=n(),f=n()<<8|n();i[String.fromCharCode(a<<8|u)]=f,o+=1}if(o!=r)throw o+" != "+r;return i}(),n="?".charCodeAt(0);return function(t){for(var r=[],o=0;o<t.length;o+=1){var i=t.charCodeAt(o);if(i<128)r.push(i);else{var a=e[t.charAt(o)];"number"==typeof a?(255&a)==a?r.push(a):(r.push(a>>>8),r.push(255&a)):r.push(n)}}return r}};var r,e,n,o,i,a=1,u=2,f=4,c=8,g={L:1,M:0,Q:3,H:2},l=0,h=1,s=2,v=3,d=4,w=5,p=6,y=7,B=(r=[[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70],[6,26,50,74],[6,30,54,78],[6,30,56,82],[6,30,58,86],[6,34,62,90],[6,28,50,72,94],[6,26,50,74,98],[6,30,54,78,102],[6,28,54,80,106],[6,32,58,84,110],[6,30,58,86,114],[6,34,62,90,118],[6,26,50,74,98,122],[6,30,54,78,102,126],[6,26,52,78,104,130],[6,30,56,82,108,134],[6,34,60,86,112,138],[6,30,58,86,114,142],[6,34,62,90,118,146],[6,30,54,78,102,126,150],[6,24,50,76,102,128,154],[6,28,54,80,106,132,158],[6,32,58,84,110,136,162],[6,26,54,82,110,138,166],[6,30,58,86,114,142,170]],e=1335,n=7973,i=function(t){for(var r=0;0!=t;)r+=1,t>>>=1;return r},(o={}).getBCHTypeInfo=function(t){for(var r=t<<10;i(r)-i(e)>=0;)r^=e<<i(r)-i(e);return 21522^(t<<10|r)},o.getBCHTypeNumber=function(t){for(var r=t<<12;i(r)-i(n)>=0;)r^=n<<i(r)-i(n);return t<<12|r},o.getPatternPosition=function(t){return r[t-1]},o.getMaskFunction=function(t){switch(t){case l:return function(t,r){return(t+r)%2==0};case h:return function(t,r){return t%2==0};case s:return function(t,r){return r%3==0};case v:return function(t,r){return(t+r)%3==0};case d:return function(t,r){return(Math.floor(t/2)+Math.floor(r/3))%2==0};case w:return function(t,r){return t*r%2+t*r%3==0};case p:return function(t,r){return(t*r%2+t*r%3)%2==0};case y:return function(t,r){return(t*r%3+(t+r)%2)%2==0};default:throw"bad maskPattern:"+t}},o.getErrorCorrectPolynomial=function(t){for(var r=k([1],0),e=0;e<t;e+=1)r=r.multiply(k([1,C.gexp(e)],0));return r},o.getLengthInBits=function(t,r){if(1<=r&&r<10)switch(t){case a:return 10;case u:return 9;case f:case c:return 8;default:throw"mode:"+t}else if(r<27)switch(t){case a:return 12;case u:return 11;case f:return 16;case c:return 10;default:throw"mode:"+t}else{if(!(r<41))throw"type:"+r;switch(t){case a:return 14;case u:return 13;case f:return 16;case c:return 12;default:throw"mode:"+t}}},o.getLostPoint=function(t){for(var r=t.getModuleCount(),e=0,n=0;n<r;n+=1)for(var o=0;o<r;o+=1){for(var i=0,a=t.isDark(n,o),u=-1;u<=1;u+=1)if(!(n+u<0||r<=n+u))for(var f=-1;f<=1;f+=1)o+f<0||r<=o+f||0==u&&0==f||a==t.isDark(n+u,o+f)&&(i+=1);i>5&&(e+=3+i-5)}for(n=0;n<r-1;n+=1)for(o=0;o<r-1;o+=1){var c=0;t.isDark(n,o)&&(c+=1),t.isDark(n+1,o)&&(c+=1),t.isDark(n,o+1)&&(c+=1),t.isDark(n+1,o+1)&&(c+=1),0!=c&&4!=c||(e+=3)}for(n=0;n<r;n+=1)for(o=0;o<r-6;o+=1)t.isDark(n,o)&&!t.isDark(n,o+1)&&t.isDark(n,o+2)&&t.isDark(n,o+3)&&t.isDark(n,o+4)&&!t.isDark(n,o+5)&&t.isDark(n,o+6)&&(e+=40);for(o=0;o<r;o+=1)for(n=0;n<r-6;n+=1)t.isDark(n,o)&&!t.isDark(n+1,o)&&t.isDark(n+2,o)&&t.isDark(n+3,o)&&t.isDark(n+4,o)&&!t.isDark(n+5,o)&&t.isDark(n+6,o)&&(e+=40);var g=0;for(o=0;o<r;o+=1)for(n=0;n<r;n+=1)t.isDark(n,o)&&(g+=1);return e+=Math.abs(100*g/r/r-50)/5*10},o),C=function(){for(var t=new Array(256),r=new Array(256),e=0;e<8;e+=1)t[e]=1<<e;for(e=8;e<256;e+=1)t[e]=t[e-4]^t[e-5]^t[e-6]^t[e-8];for(e=0;e<255;e+=1)r[t[e]]=e;var n={glog:function(t){if(t<1)throw"glog("+t+")";return r[t]},gexp:function(r){for(;r<0;)r+=255;for(;r>=256;)r-=255;return t[r]}};return n}();function k(t,r){if(void 0===t.length)throw t.length+"/"+r;var e=function(){for(var e=0;e<t.length&&0==t[e];)e+=1;for(var n=new Array(t.length-e+r),o=0;o<t.length-e;o+=1)n[o]=t[o+e];return n}(),n={getAt:function(t){return e[t]},getLength:function(){return e.length},multiply:function(t){for(var r=new Array(n.getLength()+t.getLength()-1),e=0;e<n.getLength();e+=1)for(var o=0;o<t.getLength();o+=1)r[e+o]^=C.gexp(C.glog(n.getAt(e))+C.glog(t.getAt(o)));return k(r,0)},mod:function(t){if(n.getLength()-t.getLength()<0)return n;for(var r=C.glog(n.getAt(0))-C.glog(t.getAt(0)),e=new Array(n.getLength()),o=0;o<n.getLength();o+=1)e[o]=n.getAt(o);for(o=0;o<t.getLength();o+=1)e[o]^=C.gexp(C.glog(t.getAt(o))+r);return k(e,0).mod(t)}};return n}var A=function(){var t=[[1,26,19],[1,26,16],[1,26,13],[1,26,9],[1,44,34],[1,44,28],[1,44,22],[1,44,16],[1,70,55],[1,70,44],[2,35,17],[2,35,13],[1,100,80],[2,50,32],[2,50,24],[4,25,9],[1,134,108],[2,67,43],[2,33,15,2,34,16],[2,33,11,2,34,12],[2,86,68],[4,43,27],[4,43,19],[4,43,15],[2,98,78],[4,49,31],[2,32,14,4,33,15],[4,39,13,1,40,14],[2,121,97],[2,60,38,2,61,39],[4,40,18,2,41,19],[4,40,14,2,41,15],[2,146,116],[3,58,36,2,59,37],[4,36,16,4,37,17],[4,36,12,4,37,13],[2,86,68,2,87,69],[4,69,43,1,70,44],[6,43,19,2,44,20],[6,43,15,2,44,16],[4,101,81],[1,80,50,4,81,51],[4,50,22,4,51,23],[3,36,12,8,37,13],[2,116,92,2,117,93],[6,58,36,2,59,37],[4,46,20,6,47,21],[7,42,14,4,43,15],[4,133,107],[8,59,37,1,60,38],[8,44,20,4,45,21],[12,33,11,4,34,12],[3,145,115,1,146,116],[4,64,40,5,65,41],[11,36,16,5,37,17],[11,36,12,5,37,13],[5,109,87,1,110,88],[5,65,41,5,66,42],[5,54,24,7,55,25],[11,36,12,7,37,13],[5,122,98,1,123,99],[7,73,45,3,74,46],[15,43,19,2,44,20],[3,45,15,13,46,16],[1,135,107,5,136,108],[10,74,46,1,75,47],[1,50,22,15,51,23],[2,42,14,17,43,15],[5,150,120,1,151,121],[9,69,43,4,70,44],[17,50,22,1,51,23],[2,42,14,19,43,15],[3,141,113,4,142,114],[3,70,44,11,71,45],[17,47,21,4,48,22],[9,39,13,16,40,14],[3,135,107,5,136,108],[3,67,41,13,68,42],[15,54,24,5,55,25],[15,43,15,10,44,16],[4,144,116,4,145,117],[17,68,42],[17,50,22,6,51,23],[19,46,16,6,47,17],[2,139,111,7,140,112],[17,74,46],[7,54,24,16,55,25],[34,37,13],[4,151,121,5,152,122],[4,75,47,14,76,48],[11,54,24,14,55,25],[16,45,15,14,46,16],[6,147,117,4,148,118],[6,73,45,14,74,46],[11,54,24,16,55,25],[30,46,16,2,47,17],[8,132,106,4,133,107],[8,75,47,13,76,48],[7,54,24,22,55,25],[22,45,15,13,46,16],[10,142,114,2,143,115],[19,74,46,4,75,47],[28,50,22,6,51,23],[33,46,16,4,47,17],[8,152,122,4,153,123],[22,73,45,3,74,46],[8,53,23,26,54,24],[12,45,15,28,46,16],[3,147,117,10,148,118],[3,73,45,23,74,46],[4,54,24,31,55,25],[11,45,15,31,46,16],[7,146,116,7,147,117],[21,73,45,7,74,46],[1,53,23,37,54,24],[19,45,15,26,46,16],[5,145,115,10,146,116],[19,75,47,10,76,48],[15,54,24,25,55,25],[23,45,15,25,46,16],[13,145,115,3,146,116],[2,74,46,29,75,47],[42,54,24,1,55,25],[23,45,15,28,46,16],[17,145,115],[10,74,46,23,75,47],[10,54,24,35,55,25],[19,45,15,35,46,16],[17,145,115,1,146,116],[14,74,46,21,75,47],[29,54,24,19,55,25],[11,45,15,46,46,16],[13,145,115,6,146,116],[14,74,46,23,75,47],[44,54,24,7,55,25],[59,46,16,1,47,17],[12,151,121,7,152,122],[12,75,47,26,76,48],[39,54,24,14,55,25],[22,45,15,41,46,16],[6,151,121,14,152,122],[6,75,47,34,76,48],[46,54,24,10,55,25],[2,45,15,64,46,16],[17,152,122,4,153,123],[29,74,46,14,75,47],[49,54,24,10,55,25],[24,45,15,46,46,16],[4,152,122,18,153,123],[13,74,46,32,75,47],[48,54,24,14,55,25],[42,45,15,32,46,16],[20,147,117,4,148,118],[40,75,47,7,76,48],[43,54,24,22,55,25],[10,45,15,67,46,16],[19,148,118,6,149,119],[18,75,47,31,76,48],[34,54,24,34,55,25],[20,45,15,61,46,16]],r=function(t,r){var e={};return e.totalCount=t,e.dataCount=r,e},e={};return e.getRSBlocks=function(e,n){var o=function(r,e){switch(e){case g.L:return t[4*(r-1)+0];case g.M:return t[4*(r-1)+1];case g.Q:return t[4*(r-1)+2];case g.H:return t[4*(r-1)+3];default:return}}(e,n);if(void 0===o)throw"bad rs block @ typeNumber:"+e+"/errorCorrectionLevel:"+n;for(var i=o.length/3,a=[],u=0;u<i;u+=1)for(var f=o[3*u+0],c=o[3*u+1],l=o[3*u+2],h=0;h<f;h+=1)a.push(r(c,l));return a},e}(),b=function(){var t=[],r=0,e={getBuffer:function(){return t},getAt:function(r){var e=Math.floor(r/8);return 1==(t[e]>>>7-r%8&1)},put:function(t,r){for(var n=0;n<r;n+=1)e.putBit(1==(t>>>r-n-1&1))},getLengthInBits:function(){return r},putBit:function(e){var n=Math.floor(r/8);t.length<=n&&t.push(0),e&&(t[n]|=128>>>r%8),r+=1}};return e},M=function(t){var r=a,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+2<r.length;)t.put(o(r.substring(n,n+3)),10),n+=3;n<r.length&&(r.length-n==1?t.put(o(r.substring(n,n+1)),4):r.length-n==2&&t.put(o(r.substring(n,n+2)),7))}},o=function(t){for(var r=0,e=0;e<t.length;e+=1)r=10*r+i(t.charAt(e));return r},i=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);throw"illegal char :"+t};return n},x=function(t){var r=u,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+1<r.length;)t.put(45*o(r.charAt(n))+o(r.charAt(n+1)),11),n+=2;n<r.length&&t.put(o(r.charAt(n)),6)}},o=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);if("A"<=t&&t<="Z")return t.charCodeAt(0)-"A".charCodeAt(0)+10;switch(t){case" ":return 36;case"$":return 37;case"%":return 38;case"*":return 39;case"+":return 40;case"-":return 41;case".":return 42;case"/":return 43;case":":return 44;default:throw"illegal char :"+t}};return n},m=function(r){var e=f,n=t.stringToBytes(r),o={getMode:function(){return e},getLength:function(t){return n.length},write:function(t){for(var r=0;r<n.length;r+=1)t.put(n[r],8)}};return o},L=function(r){var e=c,n=t.stringToBytesFuncs.SJIS;if(!n)throw"sjis not supported.";!function(){var t=n("友");if(2!=t.length||38726!=(t[0]<<8|t[1]))throw"sjis not supported."}();var o=n(r),i={getMode:function(){return e},getLength:function(t){return~~(o.length/2)},write:function(t){for(var r=o,e=0;e+1<r.length;){var n=(255&r[e])<<8|255&r[e+1];if(33088<=n&&n<=40956)n-=33088;else{if(!(57408<=n&&n<=60351))throw"illegal char at "+(e+1)+"/"+n;n-=49472}n=192*(n>>>8&255)+(255&n),t.put(n,13),e+=2}if(e<r.length)throw"illegal char at "+(e+1)}};return i},D=function(){var t=[],r={writeByte:function(r){t.push(255&r)},writeShort:function(t){r.writeByte(t),r.writeByte(t>>>8)},writeBytes:function(t,e,n){e=e||0,n=n||t.length;for(var o=0;o<n;o+=1)r.writeByte(t[o+e])},writeString:function(t){for(var e=0;e<t.length;e+=1)r.writeByte(t.charCodeAt(e))},toByteArray:function(){return t},toString:function(){var r="";r+="[";for(var e=0;e<t.length;e+=1)e>0&&(r+=","),r+=t[e];return r+="]"}};return r},S=function(t){var r=t,e=0,n=0,o=0,i={read:function(){for(;o<8;){if(e>=r.length){if(0==o)return-1;throw"unexpected end of file./"+o}var t=r.charAt(e);if(e+=1,"="==t)return o=0,-1;t.match(/^\s$/)||(n=n<<6|a(t.charCodeAt(0)),o+=6)}var i=n>>>o-8&255;return o-=8,i}},a=function(t){if(65<=t&&t<=90)return t-65;if(97<=t&&t<=122)return t-97+26;if(48<=t&&t<=57)return t-48+52;if(43==t)return 62;if(47==t)return 63;throw"c:"+t};return i},T=function(t,r,e){for(var n=function(t,r){var e=t,n=r,o=new Array(t*r),i={setPixel:function(t,r,n){o[r*e+t]=n},write:function(t){t.writeString("GIF87a"),t.writeShort(e),t.writeShort(n),t.writeByte(128),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(255),t.writeByte(255),t.writeByte(255),t.writeString(","),t.writeShort(0),t.writeShort(0),t.writeShort(e),t.writeShort(n),t.writeByte(0);var r=a(2);t.writeByte(2);for(var o=0;r.length-o>255;)t.writeByte(255),t.writeBytes(r,o,255),o+=255;t.writeByte(r.length-o),t.writeBytes(r,o,r.length-o),t.writeByte(0),t.writeString(";")}},a=function(t){for(var r=1<<t,e=1+(1<<t),n=t+1,i=u(),a=0;a<r;a+=1)i.add(String.fromCharCode(a));i.add(String.fromCharCode(r)),i.add(String.fromCharCode(e));var f,c,g,l=D(),h=(f=l,c=0,g=0,{write:function(t,r){if(t>>>r!=0)throw"length over";for(;c+r>=8;)f.writeByte(255&(t<<c|g)),r-=8-c,t>>>=8-c,g=0,c=0;g|=t<<c,c+=r},flush:function(){c>0&&f.writeByte(g)}});h.write(r,n);var s=0,v=String.fromCharCode(o[s]);for(s+=1;s<o.length;){var d=String.fromCharCode(o[s]);s+=1,i.contains(v+d)?v+=d:(h.write(i.indexOf(v),n),i.size()<4095&&(i.size()==1<<n&&(n+=1),i.add(v+d)),v=d)}return h.write(i.indexOf(v),n),h.write(e,n),h.flush(),l.toByteArray()},u=function(){var t={},r=0,e={add:function(n){if(e.contains(n))throw"dup key:"+n;t[n]=r,r+=1},size:function(){return r},indexOf:function(r){return t[r]},contains:function(r){return void 0!==t[r]}};return e};return i}(t,r),o=0;o<r;o+=1)for(var i=0;i<t;i+=1)n.setPixel(i,o,e(i,o));var a=D();n.write(a);for(var u=function(){var t=0,r=0,e=0,n="",o={},i=function(t){n+=String.fromCharCode(a(63&t))},a=function(t){if(t<0);else{if(t<26)return 65+t;if(t<52)return t-26+97;if(t<62)return t-52+48;if(62==t)return 43;if(63==t)return 47}throw"n:"+t};return o.writeByte=function(n){for(t=t<<8|255&n,r+=8,e+=1;r>=6;)i(t>>>r-6),r-=6},o.flush=function(){if(r>0&&(i(t<<6-r),t=0,r=0),e%3!=0)for(var o=3-e%3,a=0;a<o;a+=1)n+="="},o.toString=function(){return n},o}(),f=a.toByteArray(),c=0;c<f.length;c+=1)u.writeByte(f[c]);return u.flush(),"data:image/gif;base64,"+u};return t}();qrcode.stringToBytesFuncs["UTF-8"]=function(t){return function(t){for(var r=[],e=0;e<t.length;e++){var n=t.charCodeAt(e);n<128?r.push(n):n<2048?r.push(192|n>>6,128|63&n):n<55296||n>=57344?r.push(224|n>>12,128|n>>6&63,128|63&n):(e++,n=65536+((1023&n)<<10|1023&t.charCodeAt(e)),r.push(240|n>>18,128|n>>12&63,128|n>>6&63,128|63&n))}return r}(t)},function(t){"function"==typeof define&&define.amd?define([],t):"object"==typeof exports&&(module.exports=t())}(function(){return qrcode}),function(t){t.stringToBytes=t.stringToBytesFuncs["UTF-8"]}(qrcode);
return qrcode;
})();

let _qrLastError = '';
function _buildQrDataUrl(payload) {
  _qrLastError = '';
  try {
    const qr = _QRLIB(0, 'M');   // altijd de ingebouwde bibliotheek (nooit een eventueel andere global 'qrcode')
    qr.addData(payload);
    qr.make();
    return qr.createDataURL(6, 24);   // 6px per module; marge in PIXELS: 24px = 4 modules (quiet zone voor scanners)
  } catch (e) {
    _qrLastError = String(e && e.message || e);
    console.error('[validator.js] QR-generatie mislukt:', e);
    return '';
  }
}

// ── Generate label HTML (one page per pakket) ──────────────────────────────
function generateMerkenLabels() {
  if (!_valRows.length) {
    alert('Laad en valideer eerst een Itemlijst.'); return;
  }

  const consigneeEl = document.getElementById('val-consignee');
  const consignee   = (consigneeEl?.value || 'PACIFIC SHIPBUILDING JSC').trim().toUpperCase();
  const labels      = buildLabelData();

  if (!labels.length) {
    alert('Geen pakketnummers gevonden in kolom A om labels voor te genereren.'); return;
  }

  // IHC-beeldmerk (rood kroon + "IHC") — rechtsboven naast de tekst "Royal"
  const LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAY8AAACnCAYAAADpEtAIAABp8klEQVR42u2ddYBkV5X/P+fc96qqfVwzcScECUGDb1gcgruzyKILLLI/XBc2WIDgbhvcWRxCEiQJ8RC3kYz2SEtVvXfv+f1xb1V3TybJJJme7pl531BMT0919atX7x0/36+wi9EYGOD7W7cOPnKeyNoLV774/Y985JyDh+Ydca8HnfikoSynoQ4DjAoVKlSosD0EQYNggMdoq9GywKgvGPUtrl153QVrt43+5OWf+Dgbt22++rDDjvwu0BKR1u49zl0EMxv62Ov/49n11etOWF7qAy/78c/k6EOPXJyvXU/ftjF6Wm168dQwQnIesoM/d/jaO3iO7eBN3N7Xur3Pq46rOq7quKrj2vXHJYBiQMAoENpAG2jhKbIMFi3D5szl+i0buMsTHnvjlpqs/MPfz/j2s376nZ8dMbDgOgARGZ/VzsPMHvTDD3zgXvPH/b+v+sZ3V8y7/kb2K4zBrE4oSzIESR7UY5gqaoZYlXtUqFChwo6cjReQrhsRxAQQTAQzIxAwjBp1xvFsBDbN72P0wCVbjnriE8Y2LOzn2Kc++ZtXDwy8+R4ixaxyHmY258fvfN8L2+dd+u7+v1/UM2fVehYq1M3wYpgJYoaiAASEkE6Hit91KU+FChUq7GXew3XNc6dSE21nkBh0O9oIhpkCDhGNNhehScFVvf3cuP8CigMXXrTw4Q/5n5Ne/epvimrBLgzab7MNN7OBL73nXS/e8puzXnPwP1etOGDtMPMQBB8zCxGkSioqVKhQYUYggIjSJLA5V65bMEjzuDtfdMHWG5//oT//+Wy8nxHnMfCtf3/1ae3v//rhB29sscwH8tAG8wjxgEpxCFp1xCtUqFBhppIXMYIaSobzjs2qXLaof8vcZz72f/d7zQs+eeSKQy7Ybc5jPgw8Ie8/7eHMefjdvcNZSWYliqclRssJeRAawVGghOrzq1ChQoWZcR5qMagXAXHkQVEyVvXX+MfSno3/d/mFr/7uceE0zuF290N2ynkshP6X5/NOe5Sb+4iB4OkpmvGHxWNAIUIQoeYdNQMPlfOoUKFChRmCCngzgjPUhJqBGBQohcA/pGD1Q+95wYt/9ZNX9Iqcfnt+h7u1J5z9xz/eff7//vzUpzDvEYvbLYKVFA4yE9SEtsaGTd07wNFSRYyqIV6hQoUKM4Qa0S4XqqhBZrGTUAIKLJVeelZuWnzpPy856cgH3uPMP/zt7Ot3qfNYt2bNk89783t+dv9L1x252JcUeIRALWhcZEGom5KbEIBSBIQ4ilt9fhUqVKgwIxhXh6H0hbgx0hKh7eIwk4ggBOYHhQsub3h1Jz30mU866+dnnHHdbcpudvRNUeWMM35/zDcfedIHDvr5GfV5VjBCiyCGiiMnjo6VyERfXAKOglooJn+3QoUKFSrsdnhMSko8gQASmwkmQm6QW4nYGEvx3P386+ZsOeXLP/zxJz95ipn17Oxv2GGCsBj6XnnXe1983wtWH3CQE8yPxx6GZOR+4odKhVLABaNuAcUoMQIZs6Fw5QC/naeU7b5XoUKFCnsbNAXwNokMytKSoaAoHigp1AE1trmcs5cPcs7+Qyed+sc//nBnfke2o2+++VnP+erib/36gCOlFytbtHDULB5GqYIL8WA0QA5py8PFPY+b90m7FULABLBYYnMIghEwghhCtY9SoUKFvROha4On2uJo8ywuHJLhUfLgWe4hXLee8YGez42ZbdyZJvrUspUI66+77nGNsy98xFG1AfAlLQmUOhHJa9hx6mJ0OFlm3nEooCgjmdLKpMsUUwhsywRDK8dRoUKFfRaW7L0LcTJ2VDzzJOMuV2xc8NPnvfwHNzS3HHbrAfrkFzQb/MTDH3vJcadftHzZeCAzz7Y8YAQGimiQ2ymCn90pG2QoYy4yUtY9ZAEKjW85D5KW/itUqFBhH81OJNrDzAJeAw4Q6+W8uvDL+eHLn1h5zYtE5Gar/NkkxyHfftd7XrT4vGuWLx0vaJjRksiqUgvRcQSEoOBmud3tuLa6Bwmg5nEIGiLJmK9a+hUqVNiHMbli5BDqwWhJSVvGOLSs84hs/vN+9eUvjwCvvKUgPb6YU1v7i9P//dD1TZwFisTaWPfQ8PGJpcTofbYbXgNalIgKhRnj8+Yw3MjxCJmQJg8q91GhQoV9EwrkaR+vLZIyCaHISnJaHHLdeq750GeeeenKq4+4RedhZnLJOec+bPnKdQsXWkmhgbbGpkoeFDeJW36Pcat5znAIhKOPYv9vnULfu1/N2oF+xs3IVKrco0KFCvswAhmBzISWE8ZUURMaZaR8H0Q56vpNc/9xymffYGb5jl6hU7Zy63/+my8csGHzANZGk7OQZIlLwIuhQM3P/gVAL8LmItB73HEc+Jn30jjuKA56yAPI+uZx45s+yNytw9TEVclHhQoV9klEBnRwBg2veHE0xdAAmSglnsUjo7TPv+qFHn4HfHOHmcc//v73e6757LfnLGp5Sg3kJl36kQAUErfHxSC3MCuch6SDNwGvke1e08zXaCjpeeB9OOzrH6J+3J0xy8DqrHjhszngk29j8+KFFBbS3od15tfS3yXJV1WoUKHCXuo8FFrOMIx+g9yUsczRdo7cFAF6CfT/6R986ZnP/08zuwkbiRMR5v7hT8857saxhw95n3YjmKozLhMLdsySgk+OkCM0HYzVAvVQUjNhvRjZox7KkZ/7IO7Qg7AwKYdy0HeXo2kccjCr/vx3erZtJZdAS0ucCHVzKesSqrSkQoUKeyvEBDGdyEIw3EQcDRjeBXrLwPjqTYsuGJDLv/vzn184xXmYWf/AppFTs1//eX5fx8zaTaP82RSLR5nGuOjnTHAGQYSRWp3akx7LEae+C7d8RVLWAg2GzwJtMbK20HPkEfQeeygrz7mQ2satIHH5kaAYhhePVflHhQoV9lbnManCYjuw84LhNKoXWrOUC4stB//q6is/M1mJUM8Bt/Hyq4d6NceZoXtAwB004LVM++KQ+4w+X0OXHsDhb/0PZPFSzIMZUbZRysgGDATnMIO5D3sQhz7nWRTUMOp4zfACXjxeC0wrEpMKFSrsm3BAHjymgSHNWHLu1ftfv+baQyc/R1esWfPiDX87Z17dwGGozX7voRaX/rp9D4ymCu2Vq7nqvz5KuGE94oga6iYULkcto8+DaUBy2Pq933D9Z7+JJM11Vxp1i+Nr1fp5hQoV9mUYECwAgdy3Oarl5v/uXSe/Fp0gJdGzf/F/Cxa7ehYb4VM6HbPYeQhqDp8UC0sNBGsxQMHwD3/ENc/9D5pXXoG4kqBAELLCoaXirMWq736fK/79jQxccwWZtlDa9IVALXF0aXBoqIpWFSpU2Dfh0WQLAyKegW1jjPz53Gdu9f7Irh0+9YXPb9uqG8mQSN27BziPUoVClTy18Uvx+KwExlmBJ/z+t1z0vFcwfuHFkaoxJAZHB2s+822u+fc30792FTVpY9Kk5ktqGG2EplPEMtyO2eorVKhQYe/PPARKiezoQTyZtbnTqA397Tvf+deu83jBK1/3wvp4gRDwYuwJAbeGmCE1Ne6e9AYlM8UHT1ugj5zBM87jime+jvYZ/8BlUUPr2o98nuH/+AAr1m2CTGmL4EJS3MLRFsVLpDCp8o4KFSrsq5BU3YkTWYEaRnbldWw++6Jnm1kGkC2cv2iZeY+IJqrA3bcEOLnLf5ucB4YLgVLA48hN6CkdQWoUOII4elHkwsu49IWv48j/fhPrzr2ADf/zRZY0xxGtUQsZPcEwPGMaX0dNqPk4phsmrUne4kneRbna9q8zE8PCxq6dqpPtPuc99X1M2w26y89NfOfbv+6ecj5u2znZ8buSHdgX6/6346uxM3e0ox0vu9m/7d3hpSbVj7bEM+IILHANzv79Hw69+IYrDwSuzMQnJY6kfaGm036zOzyFBgoVsqCR5VaEQkOM+oMl7t54gUxRCTHIcTgcIi3MSkrqmGYEa5JRopZTiKMHQy+7jKuf90qyMc+Sdkk7i3ok9eAT2aMjt0ApBV58ktbN8J3pZyGdOkNFMPPpAhecZIgZYorvkrcEgnpA0JDd4m1iiaYxo0zcYQ4v0dO7UBLiCMM0fhIBFU8whyEUmeJCSRbiKEIpOWrlLW72CNB2Qh7AmRHwFM7jTGkER0BoT7er0BIMMmqUBoWL/btaEJxp1HCZBUYRIpNp4gRC05kt0z/ekUlHk0DQgtwrdTLKZDQLPEWe01vEa9bPaFnakgiRJD0JCBqNUxbihI+kZ7TVMDEygyztJJSphIL47jk060gbhXQPesDRTmVoczk0GuT1XvJaDXNG6JA6BYMyYEWJtEpsvE3ZbpEh6a6zKEOR7pVMIomgIXgDFQGzGHibYRpluIPFyN1ZDHQtXX8iUUtctldhnZWI3IYaYlDdFqPuxzhu1IYOnrf0CcAHM5JQkhlpL0K7xno6L6HMlJqPJzMAWCD3RtNBaUKWbrewg42LFi4uLkpIaoFKgSBiZGJ47yktAJ4aBeXmjdTIEBxFMDQUeMlpkaXmu5GJx0vALGMEKDtxikm6XOJFbOl85RKdYEkkiwwIGgxFqYV0Y+zEufBKLJWlWYWAQIASxdItMX3RRaQn6MQZdW/kyTWWSFwkulWjaFM+IUtMzJjRFKFthlJOq1F2QbuyyB5wpkCRDM5suRVLSgJmHRI6utdIaZ2g6g7cU2bgIaAUKE2gLtFEFxgms4HRNAZMOinCj3cpeDfhPOMCm5EldVJU8cEIKF6jrXIBMhxthBY1wuJ5NBcMwYFL6T14BfOWL8UtXYzMn0+tr4dao4dGrYE4h9VcCgrBfCC0Cspmm9ZYk2JkDDbeSHHjjWxZuYZN16zErVpH77pNNNcNU6NI7yI5tRQQmICKpxaMYEqJw6Ug00u8z8Xie9wTmPUssRkKGWKOtpbUCPSu2cA7HvXYNkAWnCRDEe9E2w11BiPHTMnwsWzkAC/UCbS9Mr5oIe3Fi3GL59K3cC4hV6zwtDaPUKzfgm4dwdauYsHoOI4AImRm8WIyT6u3BxucE5l1xTOeKbVSqXmlKW1q7ZLWli3d/k4tCDVTlIxNtZzWkkUU8+ZRmzNIY94Q2tNAzBjbvI3ixk3k64YJN66m3h6hoRmiBsEwETJzyb2EJMp782dBiD2mQhVnkPt46tv9A7Tm9iHSMe7TdzMHMlqZo9729K8bJitGIzGmGC74nbgWDBcMMaEUCKK4AI4amwb7Keb10yinN+43cwTN8SHQKD3Z+nX0+xA/AQ1xmTTM5I0I7bnzKPp7YxZkgUwcEizR69zxzCNmNQ5PjokQ1OjdtInebSOoGV5ALMxoucUQguRgMYLPCOQBPEKQGCiWYoiEqFYqdIXoCjxiAe8dpfZRHrKM9pEHsPQ+92DO0YczeOBB1JYsgaFBqLubeOKOo+4wZXTKY53v9aTH5KNd5gO0Sti8heL61WxdvZbRSy9jw9/OQS+9iuzqG+nx4+TBCAhFFgNJZ2kxmbR8jNAThMygKVAAmc1+jkCbVBxUovOzkVEWjBZL1DmyTZs3bxgUWTBBghhw0/i2BGgAo2qMAI0gjHlj80CDobsey7KTHkHvPe9G7yEHwZw+cGDO4o3mBdoBWTvM2vedzPovfYWG1LEUBTigBOaccB+WvestUAMoKZ3gvEPaARoZ63/7Z9a96Z0MFp62xk0RMWHcAq3DDub4z3wIDjsQ66tD7sApXQ350RbFDWsYvvxy1v3wZ2z91Rnk6zdQy4SGD2QIHqV5K7onkkx37sEFIQPqpgybpzj+ztz1g/8Ftcb03uwKQT3mBNk0wtUveRsjF16IEyM3o9cC7VQIuCXEG0EIGkk0B6xG0wIL73tvlr3n9ZDnMF37QxKvDy9Gpo7R8/7Jpa94C70bhmNvrDt9PjO3qgDjeJa++TUMPfJfwLeT5ZKpKeAdnZJXiyUZzcAMJ8LKT3+LTR89lUHNqQVL45czF/NKyipCt/4Qy0JZN2CNe2ZtjJYKPjjUwxiB8TlzaR++gmUPOIHF97knQ8ccDgctI+QZQh7nRC3q96RCBqaxhNQ51ZIiagmBINFxG9F5lamIHPMCTdezQ3oV611ItmwR84H5/Csr2uPY6rVsvuQytp1+Nht//xfcxVeiI8PUURoiCIERJ5hATxmHPX0ywPke0S6RSb2iqIUkQMOEvr7eF/3+4nM+k335ox/47Ev7DngLoyUiNu206wa0tCAoNHx0Vu3DD+Wwt76cRY99ODo49ybJkwXDMoEMfB3ygX78fksQaniLhCwuFV9KjPb8Ibjn0fhYAYqePj0UaN2wllIzcjyhG4zFgYGRuQNw18Moe3piEz29rjgBB1Krkc0dZNGxR7DokSfS/Ms/uO5Dn6L9y9/TizGqHq8B9Qa30q8QDIfiTMhjPhaPYd4A3P0oSm1M+83sOp/Kpk1Yoy9WpcXHGvpOjt51ojgzw2nc4Bcy/Pw5cNcjKDWf1vcQP6MYI4Wt4wTNcER2gZCm52ayTuAx/MH7w50OiSwH6bzrpIiYW71abr0EKamA2PnUWosXADl4S/0AmXFzlOEpxGi6aMBdgBwlJkVKPfEr+WCMZ734ww5i0cMfzH4n3p/Be9wJWTgXuhtpEwXTrJtGRA/SId5wYpPynnTFqOv2mzr3gG1XopXUo7Bu/2Mi/vG1BnLggcw78EDmPfJfWLF1M2PnXMiWX/yJDT/9E+U/L6ePcZx3sU8q0SGagFnoXrGzG2HK2XAIiFGzwOCqjfMPa8x5UfbcU07J3Slfo7z8ukgKuBvGfAIlzmBYFH+v47jTx99D4/i7RoW/AkTBHN2au1jyAqkZBUah8XbxeFxKycvY4sG8QTtAPXZwssSB1dkcd+02hQotBC/xYjMTHIoWQKloyHBhopw35YrSVKttNKg95D4ccdQhrHr7ydz4hW8wTwSxNqiAz2824u1c60Is95RihBA/MC0FSofUbJpvZsECiAoEoV7G6IgQDe/OSWYl/ptO2cWH1OgX2iqpI5VN2yUlxJJnSJ8LBlmIsWMrfU4zrXwpgJQ+JQghXULx2ko91l12HjplsJxYQgzE/kEQ6/ZbZhadQZQ41OBEcF4wc7RMGAZa84YYut/xrHjK45jz4PuhixdBJgQPPtleR1wWpsMIIQHvbDv3KVO4miaxMsUfYRKX38QNCS4aepmUrZBufRLjOCmj8eIoB4foe/AJ9D/o3ix59YvZ8Mc/sOq0H6C/+QdDo6MYDu8CGR7Mp/rPnrJHNsF45S0gvuRQHaSv3kN210c//Od/+dCpLzlC3KBYiDfhNKMWlHGBbYfuz50//m4ax98V8/GisFSUDICTlOsx8cE6ly6KMr0tp0iIswuFBiwYmTcw343GnCkT+rkBsRaFxnQ19h0CXgKZ1al7Be/o9BfFpoaGQWJprPOtEECWzmP5B9/I8PAmmt/9OUOaM663bLE6/jCzONPRcp6gJVYKeQm0HS63VCubrixQ8epi81ZT8x6jlqZF/E4EEgFNvaPY4ASjEMNbp2QUUk43XRZcCeq6JQjrfkZGW6JevUuN9JktIMcjcJYuXKTzv112k5fOTRly8Jqc+sS8zwz3PKBIU1Y1L9SC0hBHy2AYZWTuEINP+FcOffrjGbj3sdDXn2az4s3Sydim3BKSLL9F4y7IFMO8/bsVmdCy0O1iwm750PtOMWLSP6YrWJLTkomst2YZZh7TGrZ8CYuf8VQWPvHhjPzpH6z93Hdo/vR0esY3E1RR56LR2EPcRpBJGZeChMCCIGy5cS3ZQQcd9tcPz18xppoNim93a4bTG3tkbBPh8H97EYP3uBtF+p1jEqhrQSMoatEgBI0fnptIOiFAwzLGgBJPzeJNUbh4heVBQBwFQjapllkSswvEkXmhJ0VlTYvljZaV6erUOIKr6aJxk6fKJ/i/1EALoVl3tOfM5ZA3/DuX/+lC/LrVZKG81Rs1CGSm5CaoJzXZjaCdkFRuMmtmk8pEcgs36fbP2dGEemCCxksJ0YkSEIt3YjtWCm+V6ssnm9TwcUOmqYJ5oe5DfJNucmHp1o9rZ99jt2xlyRBIKnYT0hSc0AiaJvJmGp0CYdbNODW9kckJgdyOz3Hi+wFn6WLVjk217aLumTZHcTgks9jz2GSBTUNz6T/poRzzwqfRd/d7QE89lptTU7kTfDgNqVMo3VIz0ulRdAYOdtQ8kil/6HblFUmBnHV9uiaDKVPyl8kThZI+OLXOxFxcOHYQZ6/zuQyc+BAGT7gXI787ixs++zXW/+p0FjWb9JHFMtYsdxzbfycAuWZsue56/uu4e7YzM8vOfvP7mps/8GnmiODM48mm9cDGCbQPPYR5jz6RQuJBiUKGkZsDEyQIXn3sw6xZz8bT/8qWcy8m27wVfElx9sWxHBViaagtnsxrrCl3o4VJKaxMpGBiLjV5U5koxHKdQ2hccw3Xv/YtmBNsqJ+eA1cw7/73IT/mcJoulr/ykOrHKojz1ExpitBz12OpP/T+tL71bWoS2IF+ypSAyaVlxAxPHkhZXzpeF1I7UXbLBWJp+lw7LtYcmYVb3bUPEsknxSxFiXH3JdDCxMBq0x7ths6os4Al4jYhUE91t5mO8www6YYeGDvwBLvgd8jkBlQq6mo349Bp/xyk68I0DcZH6YRCUpkqeFxaB2ghDOc95I+4P0e88rkMPfB4yHPM8m5+lCOdoL/7Z5h0Lrul18nvfXKZgu1OyBQnIDfpGXXcjle5yT1y05zNJvkkxblYhhSLcUIs4ULZ00f+qIdy6IPvzpbTfsXKD3+F8QsvZAEFgYxCYnnXGTRCdPblrFjr7AwSWZxo7Z5royzbPPppTzsxA0Z+9v3vffnE3v53DI1v2S3rVC1KsqMPwO23FLUUuQK9CJKo0dXHqa/hP57FJW94P7XzL2Vee4yCVmpyZYg0cKaR20qgHhxtAqULqYkuyUHolKYiKY1vx9W/+GlbrBPPvXEd7S99hRyjiWez9nDNkmXs/+qXsPiVz6HsqcX008eYLuTx5NbJoCYsvP/xrP3Wt6hJMVGDu7mikcURxInyWVyGckEmX5s7NPiyE05BbjGS6Az9dBYxHSY+MZzlBNyk2Ry5hfKbkYXoYkoxnOVkVuJpp7KX3q7j2tn32DmGTvhuoinON+reCDjKWSHuFbpOLa613fSNyO38HOmWptI50IlpvvgZuhSITP85kEnZqCWnrmap7wKlBEYDtI88nP1f82IWPPWxMGco3olFQDLtloSmmvCsew/f+hG42xpWT3EQulOGVW/6G9ML+GR0NcSx9VIE3zuHBc99MnPudz+u/tTnuf4rX2fp8Bh1U5oSUjM9ZkWzQUvImDxiEMfeXUdZ1kqOvMtd7qciYg//7If/tG7xwGjQKIY+3RdXjjC0bCnUXDd5yybdYg4wJ/g1G7jo9R9k0d8vYr92QZ2MmvSTSy86bdmRkFMnp0HD9dKvwoLVa1n/9k/S/MVfqHXiK3VTLrluSrvfougOrCJWrLBvoVNkUjHaWtKWGBo5U1QyJOSsr88le+7TOPYHn2PhS55KmDOABMEVCpLtFfqd1pG2NkNKaBRG7uO2vxy2iMM/9EYO+MopbLjTMTQpGSxLHC3aUlIg02jbbuP7kNTz6LwnM9QiIW1ciAbu9cAHnn7A0x+/aVsIKdmcXmQIPhjkik/cKfF8x8vPW2xoNS+4lIF/XsUcSoQCryFRLEy/YW67GAXkoU2v8ww2NzF8xllIKll5mYg5DJ1Y8HJl3KGtJEEq7IPwKQRsWBlL4BIHFzYFYfjAQzjoY+/jqE++Hz3yMMYsowxKsNiMbbqZLzHuCnQsaHAQMggupWHmKFHaWZ3Fj3kkR333U4THPYLNQL3slDXjgMlsyDwCE4M9NiXnMnKXda2w//lPv/+FTX0Ncj/9Xk9RxtZvSjPPcQgqlh5Imz6xBVZs3QZli2202VIraWv0fGrT2wsACMHIQhzzHZGSMUqKYhvgkTCZdG1y5mHYpg2pXKYVM2+FfQ6aCI4646zBlG0m6IPuy+Hf/giLXvwk6Gkgpog4nAiqYGqoGLIX3DSdibeWwKhCy4FmSo6j7nPyEPm76kceykFf/BD9b3otGxrzCCHH4cloM1ukMbo0LkwMX7ju5wyIiJ3w/neeecOSeW0v0595BDzh2tUwvK07cVISxzstuTwBepYvpOitk5Nqhxr3HWu74cTWTBN3mpEljiRdvAA0w3XGUJmY9hEMQsHWSy5PlCkOq9xHhX0Ijsg351UZyzLU6oxbnZ7nP4mjvnEyvfc6FjPDqyIi1EMcKgwSyCRQt06Dfw9HAAnaXUrWCevbtRnOytjbmLeAFe9+E4s//C6GFyxKy4SzQwJ7St9p0kRCHIeeRNn6kMc//tf5ve7+m/W16XUecQJOyK64ltaFV3YHGDvnVi16NDMjO/Jg5JADqOMY9BkSFC+dtG56R93iionhzFH3gvQMsfi4u6V+R2T+ZcoxBNgywuY//wOHUUiorEmFfQoBaGGMOqH0jo09fSz5r1dw6CfeTli2kOAdZomENRhqAfUhjr7bdtOGe3TqEctUeRB6A9RTOd4clC4qn0JAQxxPKjJh6cueyiGfejcb9l/BsGWzgA0gBcU2OfOYmg91nUcoCo57wyu/t+2IFTvBZnRHc1uPbh1m02/PjLPoLtbRpOOmJS7wuflDLHzcQ9mIA8vITfASKNLxTefpbYtROkVxNFH83Y6kdvxdKDB8Yiy0Tv0qzRqPXnQJtfOvIGPPENWqUGGXOg8xJJGcbhjoY9kH3sSCd7yU0NuL+pxSlLbGxVglDsWYCMGUIC5Nhe35N44lhoxYwkpBpkSGbCGkPZMMnyopYiVlaDHnySdyxKc+xNYVh9Oe4b5HZ4Ksw2k1Ofi37Z0HwMF3ufPPLu+VNVvyDATy9DSfrHr8/zsaUcelvT6M1T/+JeWNN6aMQwkyQcZFMAIZSx//CLYcsB9bABcMFbq0E1MyqsmZzU5kP51anqbJe031vM7XprGZXwKbspyFz3sCOn8uwSSNt3qwOGIXLIAvueYHP6O+dROZy6ZuflWosJdAJu1MaKJXR6JdyBDyMmPLnCEO+cQ7WfzKp0NWQ8gpVaOeBROLv2KC17jnYMTe594waNLlU5i0XtNZGHBmqBcsROr8QIg7ZyKUBkOPeAhHf/L9bN1vGc3EfdehkMzSqrRXvxtkM0hkTx3+b+kqrHb4D6c4D1Fd2/eg+556+dwBTHNyDBGjUIdJRpZmne4YG71g4qipI7voAlb97w8ioaEXJCiEArWSLLVmakccweInPYqtaQohN8E0YxLdVfLwE296Z6L+SAEtk5iZrFs8U4zcBzJTxgF316NZ+riH4xEy0Zhme9Dg4k84pXnuxWz5zs+pE/Ch6OqRVKiw98C6WhYdq6hpZD1ISTBYO7CQA09+N0uedRKl5GA5ziLVkCT9nS4bQHImjRTl7oY9xt0Cx+S9j7RVT5RqQBxkgrj4vnMcaI6TOrnUQWHwMffj0E++n62Ll0Z+OSeUGiUjsuQ8gk6/8wgyIRthKmkQqOM8tpepM+M573/vWeF+xxWbPIxJnILIumN3QtPdMaXBqCYe655zA6w59VsUV10LmSVBuAzvlKCJDijL2O+FT0YOWoFYRtM8mS/v0DUmKcrRtBJZSuThL136E4lsSKFkJKtzwIufjVu0sJvWaHKAiMaYoCy49tSvMbRyLSEtrFmVeVTY68pSgpdEBYKnUE+bEhVlTHJW9/ey7OQ3sPQZj2TcQmS89kAJmY9qgdVdsXM2cuixD+PA/34TG+bUGLcRTMdoSTPuyfkaOgv2yG5yBCLym+JOh7x75aIhvNQoCdRDgVJSCCkLuYOXQNqirANzL7+KGz72BXLfgixQSJSGNQuYRjr1/KjDWP5vz2JElEwUZ3dElaBD4WyJzykKzgSZKGAFEUoKgrXJH3wC8056TGJupSsQEFTxGilNtv7yD4x95xfMQ6hLFICpJq0q7I1FqziG6im1JGgZyzBBGXM9rHjbK1n2wpNo5uBxSdExWhmTqpK702fZwApY8JwnMPf/vYTRekbuPbl52kDAzYqTuUP39ax3v/MzC575+LO2WpzmrRHphD3gOiyBd8B050EoEl/eoMDWr3yH0V/9MTGiGuoDksRaJESunEXPfwrhHndFQmBXDINP8NjE06AWo6N6gNxiujk8OMiBr3kRsmAg1TEDEhkMu1w45YZ1XP2ejzN/ZDOmIQkYK16rLcEKexecGVmXrDM2f81gK8LSFz6T5a96Pl4dLmT0WkaGEhwUDgoNkbKkOo07ZZykiHbp4Fe8lLnPfCabrEanGC53WDlsGp2HiKz78d/+8KFr5/bgrEaR3lHcuL7jy2+a9NLbIYqj9G3dzFXv+TisWUNNBBVBRMnDhGC8LF7C/v/5Erb0DlCY7hL2F58qd2KQB+hByFNjaJs5Bp71BPofch984t8y75OoTkACZGXJtR/+Au5v51HD08ZoOiHzVGFWhb0wJO4MtATqwVHzGZsRwiMezJJ3vApqDVzIyEOGhpipt2WCcshZVbbaycIMZR2aCKZ9HPy2t1B/yIm001JlRhs3W50HwAf+/Me/DzzjcedsEqWZZZRJVGlXaO96jc2WNLgWV/jPOo817/ssbrxERCb4cBOnXRmMgcf8C9mTH8PmLr337T+SsJ0zy5JaVinCNmDsiKM48BUvITQSUSOQS9w1keBBjS0//jUbPv01BikJWBSGkbihn4Uqxqqwd6EQo5RARhygKS2jffThHPzRNxIWz4cym0jrZYLwtGbgvFKtPu183arQglIMnzt0zlwW3+lOZCZ4AjpL1vD15oMMWVl7yN1ft/KQJYVZRiGOYH4XRQ6RnSpDaWuM9hdYYN3nT2Pdj3/TDU+Cekx99MQGvl7j4Ne/lPGDV+CtxKzD+Xrbj6ozESEhZUJACxhRZUMDDviPfyc/4jAMj0iU6FUTsuBAoLziKq5654dZPLwZlSKNGzpqZYwJ8js8lVahwiyzaQRMAqUE2jjW9fZyyDteizv8UEqL/UPfIalW8ARcytLj/azVHbGTyK2kPxisWs9Vr3krK7/wOQrXwgiU1HYLv9/tdh4A9z/pyX9q3uPoP2w2xbtaGmX13bX17R+2k9WaqO5nhERQXSuhgdDfHOGyd32I9oVXAlGxrztVrjFLadzpMA59zYsZS7so4NMomXXHbsVCEolJbzGJi3VeK6hiImRAnpTuLHY02CQFcx5+f5Y88eGUAi5I+sCI6oUm2HiLiz5wClxwKT0CmQiFKllQ8kQtOTuqkhUq3LZyCdvd0xP3LEnHxlNIzhYTlv3b05n3mIfggicL0XFERtkAEiYmbzUQHDR3Rc9jhzfWxDdv8Z8nRAdv9t8mvhmmPPGmT7FJj50/3B3bhY7S48T2RoZj6+l/459PfgVjX/wag+PD5LQIGvnCZkPucYvOQ0Rs3hMe+f7zDltajvo6PVEZOqnzTezATJ7d9jsxVaHB4ZP+eBbi3uUYRg3Pwksu5vI3vws/PEyeRIkUqEvU20CERc9/Gn2PfhhtC9TUMe5KSlUcSg2hAdAq0LaHVgllgfg2rmhCEbAQaGlccqoRj1mcRg2Ipcs5+L9eA/MHyATQjBo1aiJYHchL1n3yK/CNnzAPo21QhJyaV9Q8nhJPnDapKrwV9ijnkRQ5NQVAHf2bVFzGieGCUQSlvPtdWPofL4JGDSdKpopTiS1diYM2OUqGpr/HvYY7HC/fxPoGgnm8BWJHMiQ9TiYM+3ZR7mSW2M4tamnWJXLrxTH9YJ4QJqiT4tSxEdKCMCHEPy10tVptkvOJO+WWequdo51UMk8HEgi0rcT7Ih7OWIsbT/kyFzznZXDW6SzUqJzqQp5oXWYHceKtUug+9IlP/OOXfnP6H7es/d1D52wZjxrASV00SFLEs65ufOKl2pncY/JnN/H3Xhzrf/lbVv33R1nypEfRHh5leNMGyuHNFFu3UYy1KVoFPWMlmdajZnkiITRyIDB89kVsednroqaxGd6BzwwNUGsDN6ylxwvbpAYSKBx4r/SRs0D6ueHrv2T8R7+np96gMdCLLhygNn+QwQUL2HbFNVx78iksbY/BJDqFyUuKFSrsiZgc7oQUVIWudGsccszJ2NhocMAbX0K2YjmYp5SooufCVO+wfeayK0IpnyoMGqJMcylKEEVCSIt5gllIQqIa30mzSTE+RlG0od1G29FIS16Dep0sz9HeHqgnETQRgkWVPwzqMXlCgaAWddhVuhbHTXJopUwVlNJkGzr2IbKzdziNHGZQiotiSw78dSu59r2fZNU3/5c5TU8/SjCjEKVUwwUhj8X0GaeQvFXnISLBzF723Yse87slZ23eTxBEZSK7kOh4nU3oXZe3w4hOkJrDEi9sO/nLXP7Z7yClp+Wb+PYYWnp6yQFjWw5LJQPa9ASjFAjSpi6Cv+4axq65Kn6wxF2RdlfL3OhDmSdCgafUQG5GLhoV/a6/lPbHLqLdza9KMpexuVFnVa2Ottr0l+O0nFLzlcGpsPdAJxVRyk7/YkJsAEXZGow5jzmR+Y99EG2iM/EYWbceNL3ZdpmkdcUyJMRNbRPIRCNZqROkKGivWsumcy5i07kX03P5tbSvuY7m+vXIeIveItqosl6j3d8gWzKf2opFyMH7s+ROR9Jz93tQ238prq+Ol05pKjoSNSHXuGoQiPRF0pFBTvYmdF1D3IvpWvlU25cUbYZUqckAZ8rYn/7Glf/1Pooz/8J+WqcRFBGloKRwcb8jT0Z7Nswe7JR4h4hc8dN3vvcz1178z3cftNXHE2UTKWBI3tZsgmHgjkTgraTp2z88Tg+egiItx+TUyEAcTR/iIqHUEPOoKu3EBdNjEstQwcgtVhO9iwJOOQI+6oIIilhk9SwFRnIjk4w5ZcYgGQWBWgg478lHW4RRT47iJUt7HNX4SIW9Bx2m6o7j6DIxmCSFe9iyeD53ec3z8Y0+vBm5QM2SurWbfgHVeohlsbaL7NsaYkkbjLB2A1t+cxabf/pb1vz1bPINw+Rj21ACNTz1FEC6pEYRMMp1UFx9FQVGSWBlrZfW/GU07n0sg49+AEsfel9k//3ARdaIIIIr4z5YdzBAIl9X5K2KPFVBkyDcpBJLSOmXJOqlUlLJv9Vm3ee/w+r3fBh343UscEq9Y5cwRDWqEkJasGbaua12mfMAuOfb3vKF7//6dy9detYVy+emqSuxWM9raSpj2cTy3B2KgMxTw/BibBEhQ9EgyeMqZkLNJPVNor8vfYwmHBonPHynOplqlj5e1kVabxLqCErdCw0Mj9EGSvHUQiQs88Q5dSXQY0oOjHd+ow+V66iwV6ETDoWUgmRpYbYTTQ8jzH/a42kcfwzBjHrSridExbyS2JecTvdRqqI+7mXhLA7SbFrP6u/9nLVfOA099xL6ijH2w1Hv6qIKbRwtCRQSd8s6qqUuZQiNjsxRu0XfmqsJP7iCkR/8hMuOPJzepz6K5c85iezgA6OT0cixp2V0qloT2rGKhkskh5M1PEyhLXEgx6UgWwxqCuXqG7j6XR9l6xe+w1BZ4Ih7dR7f5dGTEHuznRJ5a5aUx3faeSwSWfO/73jXZ/75jyvfdbem0OutK3biLUUrk9K3O3QRa4ETo10KGy2y+faIUpeMpgJO406GqxE05jquXkOyHJdl4DIKdahmmCpOFEKgDJ4gig8l46GNDwU63qLRLPDBsy1LhF/NkkYhlIRIzR4CbdrUKFHJUNrpY6xXFqfCXoOQAjZL5ZbMSEVioY1nfP8VHPHcp0GWx4GqbvJtBNk9o6MeKAj0ZAJFweZf/4nrT/4MxZ/+Sl8oUBfIXaxKlEG7uuqGkJlL7EIBS32a7ct2hqHaxgn0+jb2z0vY/M7LufjbP2PhS5/Osuc+ntbcQVoKDe/AC2qOXBJtZLKoMrk6kyQmBJBQxmBYHaN//hvXvPk98OczmZsC35IcM0fh2ngL1LwmBy54jJLYd1GTGWcgvk2asw9++1s//9f1wy/Y+JXvHBjGPf0+0fTaRLixK96PN4sz4ocfQ37CPdDF88gXDtGY08fAUB/1gTnUeoawvl6sp5eQOZw6MtEoTSaC4pKMlyAWOdwtkBpqHq/tyA5ZeGh5pCgJVsDYGGHLKDK8lfFtw2zdupmtGzexZeNWBlZvpDj9LHrGmnHno+p5VNirylYpGJzEmBpEGccYFs/Cxz+M+jFHdof1nQZQT5zQynYLKW4eAjghbNzANR/+PGOf+BqDWzdRQynUCCEes1el3XVo0h2ewaQrE2up3M6kYZc48JPT1MC2XGh4YciguOxybnztuxn/8W9Z/s7X0nPfu4MYTWfUxHWVRYs0jpwzwaLUmSPIg8eFHJotVn79q6x576kMrFzJoOQE8xQquACNECiYoEIvuxkUIGXsN1tXo3DPcB6LRNZ84NnPfvaq9pbT71sbotYK5CHWOWWSxOIddSCOGtt8YPARj+KuH33j1LxadhyNdC7cQGyQZ8Syku++pkypx2bc9Hi3f+k+YOHkyOzytVzyr09HtzbxmSer7E2Fvc15pOlJTRMxbYxRDTQXzeHoJz8Wy2NvMdNo1toEMsvISrpR964qou2IRUIVysuu5qI3vo/2T3/LPN9CBUrzZOZwlnVYtyjFIq24JUUM69iI7fckQvf9C0KwBnUPPRIIoaSlJaUG5pPT+v3vuPiya1nx/17H4uc+GuvNaGM0Qmqou6hPHlI/o5PRdBQ5ylXDXP2hj7Pt85+jt+2pa04wi93aEOWwMgK1IBTJlpWpJdDZY3NmMSCeYdzmI3jT17721233vdevt0mdtnP4xNHvzKGm7AqZkjwIdZSiHIt1zlQaMwt4KzAr4na5xQKi2ETooMnrm8X5dIePJ7s7qe0x8131jpAirYDFD9HiclAwT4ERLH2CQCklI2MjmCgmWlmbWWT0ul/bTS/qzvLqjDz2sDPp0n3izCWCzwL1BfMf/Ajq97g7AtSJ+xxCRk4NJw7crsw6AlhJgacww8pkBIDy3PO55Pkvpf3jn7HYl9TEUWgcYjGbWPc1DDUjC+BMEoW5dnOqqbuBE/8WeSE8Yh5CLDZpUDLLIBg95AytvpYbXvsGrn/Xh2hs20qGpymBlhjSjr2ikgKzQCgDzkcbOfKXv3HRc15K85OfZ1Er0EscMQ42IbAUEAqUNtp1ckpn1lfAHBKyGc86uJ1HUDz5858+Y9WRKyh9wFTSh5V6Hrtg3kLwKCV+6zCEgixFQzH3VMTihnhH0Cme1KnyiCKa2KoyQOIHZEKw+HGoSaQmCZAFQcuYgUtaRVfvyDs0oJ2zNLKFvGxRtzb1oqys9ix0HrYDx9FxKjPy2IPOoyGUXYenlCgWAu1Gg8WPeQQ0son7KxWrXOeOl13nNwhxFSCOsAri4x5A8+LLOfcVb6L9l3+wSDJq3d5ClFGI9sCS67CbXAedEVrbwWPyv5FWDf2knERDLHcZ0K8w2G6y6oOf5ZpXvx+3bphcjDyAqhDMICTNVWcQCq7/wv9yydNfSf6n0xnU2DxX60y4hSlv36dH2O7anZzHzAbcrqM44JBD3nPgvz/vlOGBHnwgNtkkJHWrO34VFRJoEtjS3EIomyAWS3ydXoa4iTRbum2O7ZjaO0loZOjVSQ+ReLF3NdMVJBPIIgOipbV5SSOL3qWLa3SYkgKHdSV6K1TYWxCwyNEmLhpPFyeD8qMPo/9Bd989V3xqRngyxLKo2lkX/PWrufBV/w/528Us0AaNQCIoTzvduzXLU3pxLDVl5Evf5epXvhu9cSPq4sIxqmRecKqUGzdx8WvexarXvJOh669njpZkt2sTbvbhdlUpRcSvNHvnGV/99pPm//n8pT0+CiupVzyadh9v/8mppSxidPNWKFtQawAeaRUw0iaMtgljoxTtNkW7xPuAL0uCD4TSY94jpU8hhYeyxFyGZBmEgDjFahmSOVzucJnD5Rkuz8jznFqjDn290NdL3mjQqZSGbVsx304NsIp6ZPZEzJNCFrkpD5HMYKzWlbHeA2yFJcMtnQJKUFpkzP2XB6NL5lMAtd1wDHHkNQZ4wQnSHOWGt59M/rs/M+SUuo/HV2KxhIx1ezW7AwVxfN80MGA5W0/7AdfUHAef8laYMyduo2fK6F/O55I3vIfsz39nGQW5xt4LIVZM9nT3cbtbXPuJbPzBu9/9iVXnX/new7aMkSVKgPEO79UdODMtcRg1ei9eyYaXf4DxkVFaW7fgx0Yp2gU23kKaY5TtgrJdYKXHihLzAfMe8wF86PJtxZZInL+2ZPjFKeockmeQxT8lc2T1Gnm9jqs1CD29tPp76J0zh/4l86itWc/c0QBkSW62MtyzxXlM9uU2xRjuuJy1249vTziPYtRCR+zJ0FDSGpjHvAc/GI9jdwwXihkthawURGNZfOPXvs/Wr3+PxQQKHxl9BcFE8GLdRn8nQJju4/QmiW4+UGqLAcvY/I3vc/3gIPt/4A1QM9Z87UesescpzFm1mkE8XtLCX8iTYuk+mnl08ND/9/8+e9r/nf7KBX+5dMmCMjaZxTnM3zHND4fDoejqTWz6yrfpIdBDkVb58q6qX6y6TvRZlInpjHArH05eTkxnTTTNJ2qmGUaNjGbKo4aJAvR9SfMjbp9Ua4KzAbfUZbPt2lYz4Tl2lm165p2wdbXqwFMSKA8/iPzYO+PjCh3T3cWxEPBOUZfYqi+7jCv+51QWl6203utou4kjFkuOYzcGCYbgQqSbz8UI4pmPY9unv8m6vj7K0Y1c9+VvM9gsGBAh4ClN4g6HxFKcsuf3TO+Q8xhU3fCNd7zjlIsvvvq999hW0FO2EQt3uGWuFjcsBWiIUkqBiuEsUo647SbKO4rhEyw8cVTuZm9Y23FmpIk6QBFUlNJiO75hGTWJLMClKKVC1slsKswyTyLpGuhQpHam6mYm0pMZ/N23qyQjQiNNLW3DM/Sge6ELh1KzeDdMAIiQh+gNpNlk1f98kfoVV8cRXXNppDXe251sQ9Nh7S5znBtY2v7ODEyEUkoaGBs+8il6yhaHIIw4Y6saeVAyU5yB72xX7gVVizs2mW3Go9/2ts9879IrXrn+R39asp8HNeOOag2WUoJzmC9xEk943GAVSNRgXotJqe72qS+JxOwWUs8d3OSd6Zh4IUZ+LTRmHmJGJkZJQc0cmRll1feYNRFzN3Qwo1WWNGNFvKsZM1NZogBNIr33nnAmWxoX44IEWo2c5ccfS8gFLS36ZTf9vl9DQEzYduYFbPjuT1ma2FdbkuMC1LBI8cHE5FcJFJP4uKa7DGkmFGSRBN6M0hml8/SVBS4XmkGijLZqZG9P6/g1K+gsEezbzgOYo7px7crrTznrsme/d875l9EnHh+EzIScqNVRpsKSECg0zk67oLdwwynBB5QoXelwkU6gm2OEKQ6qm2EkG9Kpgd7ih89NRzw73zOi6BOAhImxPzGJEpzhjjvICtNTtuo9cAUHfu5kas1map4ru6+gsWMMYfTf5x7xWFUm8uZJ4kuzAw4XYgG3NKW5YAmDx96pc/Xvlq5/CYgTtFUy/Ln/Jd+8CXOB3ENLIGjamGNixLbzte6mj9gmTXiVSGKtMPK4Fk47gKlQQxIHXixSRcVSYW/R+bnDzsPMWLhsv8/oSf/6mjUXX7XwsDKutpgYwUJsDslEVG9iXcN88y8q3Wmm2Nlwk67beLncbHRhO/xyhxHhLf+4pd891bl0urMVM8nswSTmcOrz5rL8iY/cQ1zdLDsyU3ICJW1KMgYOOgRdujSe393BPQJIChLHLrqCLb/5M30oHk8NJUvLdDd376nt7vM1QW8yMX3pYk7hOz3UGI5qsht7Q8Yx5b7bBR/4xvn/csIzrloxzxchp2aKERjPAm0Xh9I6bLtZmP60skKFCrfDGHQFjQRPYMERh+L6GnHhW2Jtf9qPIUTa9K0/+yPtjRupiZD72CrPDDKrRhz3KucBcJ/7nPC7Y974mv9dOzgYOfMtMnR6DV1OGS8kUrKq5FOhwuwrQ0Rqj0IjGV92+MGQZ4hMTCTujsyDka1s/r8/0W8luSqKUogmfqgKe53zEJFw3xc/7yXrjrvTqi0COUrNS5fsC4xSBdPKcVSoMBthBJA4YJDVBhg6aMXUnZndARE2X3wVoxdfSi+G9yWFKF46sqtV5rHXOQ8AdTpiD77Lx65bPEghObWgkWBtUns5VB9/hQqzElEN1DAPrb4e6iuWd4k0Ik3f7qg3G+1zLsZv2RRpCkVoq0ukhVb1GvdW52HBOOmt//W5q47cb9XaTFFqIJG0TEkqW9U5r1BhVucfDmV8cICweFF3Z2r30fEFmmeexwCBUtoEQlQRDXHwJlSh597pPABUdXP9gff8+MqhBp6MQoRCYpcj0iNbtVxXocIsREjD6g7I589B+vu71YIuFcN0Y6zNyGVX00NBSGRVNVPyzt5V5Tz2XudhZrzgHe/67Nq7HLLyxlocy43TVZrkLasLoEKF2ZlzSKIIL6ktnIfWc3LSXqDsHnpJf+Na2LwJJUNMcURhNyMkPqmqdrHXOg8AUd38uK9/7pS1dzsEC0aduNFtEjctrXIeFSrMPuchgkjcf87mDiJJq1xDmVS2pt9wj2/ayNi2zRQ4NLio54Gn0EApHR2FCnut88CMJUuWfWbrccf8al2e0xSlpE1mPgo5VYWrChVmHQTB0iJeT18vOGWS1tpuCfnGtoxQbhuLBI02IYokFjMgqeLOvdx5ACKy5YAnPu6dl65YOLpNQCTOSugu0RqsUKHCLo/5IKnxBfJGPVnriX/bHXdtMdZCm0WXs3ryhJUGqrBzX3AeAPd76EPPXHPo0rdtXjyvQxgVlf8q51GhwuxzHhbppg1F6rWbCGvtlqC/XSIhKomjQlDp0gLtJoaUCrPBeQC87Zc/+eLqo/e/bpvmuNT6qnoeFSrMPkxmH7Ek1TxZjXF3GO7Ylo/CCGYdWv0K+6TzEJHN4f73/Mi18+YyTo7HY1XhskKF2WcMLPGQIOBjscjYvTK+wXUo9KNOT0erA6kWjPc55wHwtLe/9Tp5wsNYbR2dvmpPtEKFWZd5QBJ4BV8UEKZSx+8Oy13rqSP1HE9k5+4sJ3YF3qqPad9yHsBfHvg/7zqzvNedKZOKXyJrjxeqTkQ4ew/bfYUKe6IDiR3Jst2GpKRpprtNv6Y+2E/e15MEEax7RBOGqnIf+5TzEJEb+wYGHnd2vX3tlrwfoUFQRcUwKWk5MHVpEsuQShu8QoXdjtiPVJSM1tYtULYRA2+aagXTf1/2LBgiG+jH4fHqaauRoeQmqPnKNuyDmQcismHhvzzgE+fPzWi5DGdGW2J0UfMhaZPH1LSQKveoUGF3wwBvUelzdHgzIfU9dmcpoLZkAWHOUJz4Sjoi45J06BXKyjTse84D4HlvfftXhp73lL/eoJ6cLC7/KNRCXBvs0CCU1TB3hQoz4jxEopBbc8MmrFXs5hqyYQP9NA5YEfVETEGEZrIHQaCsBD32TechIhv673+fV1170KL2ONATHME6WgExacbS1EeFChV2r/NIuqqCwOZtMLx1NzsvIVBj3vF3oYnQwGHWUf1OSupVy2PfdB4A//roR/+dh9/vl+sG+3HkBBEKiTNYZbo4sqqsWaHCbsfEjoXQs20MblwXI36AMP2Dsp3fz/F3oWj0YAhqQi3ZA2eVhPU+7TxExJ7/sY+8v3jMQ7iRgJCDKAaU6RKqMtMKFWbCeQiEgArkm0fYdv2aaLSlE/FPf0VADIaOPQw75ADGEs+vWtxv1zChs15hH3QeCX8bP/6ot1yzcADBkXtLaWs1iFehwkymHpJ0wrVoMXL9yq6RUHbTzRkgW7KAvgfclS2UZAaFgJe4AV/Zh33ceYhIOPZVL//chqP3v2ZUhBzXHdGNxGxVz6PCbYB17c4ecZyz9vAsZh9iRoYxeuXVWFmk45bddo68KPMf/RBa/b2YgXfxs9XdqGd4q4cpE9laZzfNJPWNOv2Zzve7z7dk5/YOF5jN1C9eLrLhax/44CcuuOiGk4/fMkqjHMUQPL1pgbCsjGKFnYxWjSBCEGiuXsP13/4R2dhWUMFrDcwQmyHXIlCGwP4nPYb+ow4HYhStHWG07fijZhKFKhKMzEocwtqLLiZsHkEWzMVEd1s52WMMnXA/anc/nrE/nUGfb0dTZQ1yAiXFDCdogYBgouTmcAhtPG0XEPPU0EmDP4LX+J7UjDzEUWg/c6Z3z3ceAI974xu++vHv/eTflp279YhDycgIlFZQOqVWMZhU2OlgNSBByDJl66p1rHzvx5k/vJ56mt4p0nNmyHcwQsHCQw6n/05HYGaR629WJtdxkzuLrWrk6tWEq68nWzAHMdltzIhmoINzWPyMJ7D6jL8zx7fxAk0pURF0xlNMpWbQUmiJpxYEUyMzqHlF08RaSYhON/VqShVamuNs7+jdzKjzGBTZ8LInP/6Zq7YVZy+7bBVzrUClxGT2pKcVZj8ivU0gQ1AX6K0rgyg9yUb7GbTVAvQiZCITpYxZugTrguFMUBymQt/GLWw572IW3PPY3cgbFD1r8MbCx5/Ihq98h/Gz/kJuJWXuUXPUwsydPwHqKONqeAk483gHLghijjaOHGXMCsqa0GOBmjeCCVlwmLikFb/nR8czbqE/ddoPzmvd69gfrG304KkBRj1U83gVboPRM8hQxAwJHhSKLNDKPGMulhNKmaEHAY/tEQsKksQSfKrk18sWY3/7B/hWdwdk2gMBK2N2poJbvIDlr3ohm3oGCTjqVpKXxYyfJ48STFECdfNkoSQzA1M2LF/E4MffTvGUx9KyHlzh8BbltxGjZoHevaTrP+POQ0T8whc85QNX7L/ANqpDRaiFaqaiwm25iByYgigS6uSF0iihpxR6vZAFif8+E4+OhFHqsNos1qcQjLZAO+1X1BHW/fVsyg2bCJ3O73QfQ6bURfBpcGbopBPpffIjGUboL5WeGa73BWDUxZJNb4jFUDWHkLGpf4AD3v0fDL3iORzz+ZNZ/p43s3X+UoLlOHEIHpECk72jnzsrakOPfsBDz912r6N+dO3c3tRIqkpWFXa+yFEIFOmSMVWMDMFN8LEmIy4382C7PyeIwG/6ve2ft/OvtWecS6+xca4pWtZrV9H624Vxec92xzFEUngDLIDW6xzy9tfQPOYoxnD4WXAujRIRw8TRkoySOmvrvSx512tZ9Nwn4E2ht5/F//kiDvnxpxl5yP3ZYhlGRkugTeU8dmX2Ub7qK59/65bjDrt62DmKynlU2OloGXIz8lTplBBQPAGjwGin52UY7mYe2XZ/dr7OdvC97Z+3M6/l9pDRTMEiu4MllyeBbGQzG351BlKE3WaQDMMRUAxKyA8+gDu9/y2sXLCAEZlZEWsBGj5OVY2o0szqrK43WPD217DkVc+hrXU0OEIwSh/oue9xHP2dU+h/+6tZPW8uY0EpXV45j13sQC5a8dwn/2zd/AGQbIcfGlRLQhV2GC8j4uPXUmLiMTxBjCCBQgPtmXpIoE1gT6BrE4PMiPTnCM6gH8f6355JuXrDzVqLXXlPSoC2CGqBRLpNaTDwqAdy1Dtex0itTpkm5yZ2JqbuVezKK4sdvGaWflFewqjL2f89r2X5G56Pdxm5V7wYIYu5pwXQefM58K2v4phvnIK/590ZLZkkxW1TbJzs8DdOPcez5VKaVSH+Cc94xre2nHh/rteAuBK0BIxMMhpkFE5o5UW1A1Jhyq0UZKK3EJyCCVmatReTOBZpMjOP7hrZnnAmo1qGpCZ/S5SaNMgvv57hP/wR8Jh5vBntEMAHPJ5WIjTdJdPQEvsJIhpptjPIBYI45v7bs1n6zjexcnCAEYE+PA2aeNem7YyMnAY5SpjykO6jZEKGzk2So9NJ/8WimVdPkZUELRAKXAoACjXaYoyHwPq5czjww29n2X+8iHbWiL01iXO5IoI4QTU6ZUMYeNiDOPq7X6T/NS9nW2OAEQzvDK9NRNrpqHICdUo1TAJIwKvRdkbhAiIlOWXMyirnMQV/m3PiAz92WX9G0wQxwQkYnnbyvnlgVpy4CrOseNVd+Y3G2k0y2moyow/ZgxgTukYLwyROsA34Fqt/+FOstTWa3mC4IJhJ2qD2u27wVKLPECIdOzKJMDHPWPr6F3LY+/4fW+bOZxgo1VEzyAmMucCI6ziH+AgoQRQvSqEOr4rgUUqUEsGDFJgUBClSxhrp4Bs+Zl/xUEJ6RWU0wKajDuGgr36IxS9+KhZysqAEVVCSW9oubzCBYLgVizj4A29kvy//D1uPPYphL6hlZAZBjXFXAgU1E3ITclNqJtSDkAUFc3jcblN23GOch4j4Y5/99P9qPOHEG4YbveSWgxmFeEayEheg4TNcRZ1YocI0FACFUgQvEns2vkTE04/S/MPfGfnrBYhkSEjOWQUNQkYSadoN9szUsfQlz+boL5zM+iPvxKbQQytkcbhNxiilRUuVIjkMxZGZo2aORsipBdd1K4JhYkmALg5dtJ2lJT5FzeEsA8kYzzJGMYoQqP3rw7jrNz/Lgkc/FK/xfKjFxEuCQ8JU+xTMuo6QYPia0PfUR3L3736G3uc/g1E3RJsaBYazJpmMJzLIWJjLTOkxR8NyoEZbarOCwmnWdaaXqY4e/863PemG/Za022TxfKuP0YIBuIp3t0KF6cnfmDyRqwRK8QRazNs8zOovfR9pF4Qsqvp1Og1qyu7qRlqIeyhzHn8id/vhZymfcxIb6v1Y4ekvhboPE8MLNtETCTGnoBTwZHhyAjlqOVnIya1GLdSphbhrZl1OggzzGaOFMnzQ/sz94H9y2Lc+Qv0uR0HpKL0wlkPbxX0jsZvSu6oqIYRJ31UCjtphh3DUx9/N4lPfw8ZDD6EZjEZQcvOUEhh3xmgWGFNLhCxx89/MQeU8dnBxmHH4ihVnj9/32B+t6snJyJFg1H1k+yxRmlJNY1WoMB3GwFmnRg+oMCYlhUA/0Pzx72n+7UJKgZbGiN0khnJ5p+8x3ceogjjwZjSOOJS7fvr9HP6Vkxk54QTWuX5aphCKbq/DKPCUtKWklXlarsSrUcrklrVNjAen0l2Z2vJbMDYsXkr95c/lbj/6Igve8FKac4fwIog6nBNUuh03gvOESfwpnd/hMheZBdL5aoSAlAL9DRa86Ckc871P4x7/WDZqH2M0yMShYpgG2lrQ1oKCMo4JUzIbRodmJTuXiIRV1vrAFZtfep+tP/zVfnXLsXSwhQbG1WiUVNy7FSrs4sJVFjq979g8xwlNCfSWwtCmjdz46a9y4L2PJGR1EEeASGkSOqZ3mu9KCcnMK94CWq8x9ymPZs6D78fwr05nzbe+y/iZf6fYvIVelBqQA7mBlQFVIUgbs0AsasX+SmklhlESGDeh2deLP+wgFj/yXzjkiY+hfucjCXmOB2oBVASLLQ4aZlEsSwNFOrbalLOazoowyV15cCCmBIH82CM58msnc+PnvsP1H/4i81deRyM46i7y/ZkFShUkHfdsICyYtdSOy6V+7lmf+cKn1517wXuWXr8eTTK1YhVhe4UK0wEvhlpsWIeOil8abCzEqJuw+me/Zt5vH8vQvz440oiIYB2xqN1BX9Ipr1nKlURiT2HhHIae9XgGHv8Q7LxLWHvW2Ww68x9suuhKBtdvobF1lNxaWIiTSobQxhhH8LUa5UAvzf4aCw4+kDl3PY7BE45j4N53QZctiqWrArSM+y+deo14kM6AgUbOMpdykO1XC2yKMxFK0e44tFocT6a/lyWvfQEDx9+da9/xIdb/6S8sKkv6TSgoaWN4DUlNcearL7OaF/je//aC3/z8d6f/+9ANv186xzLGpSSjpL+Esup7VKiwaw2zggRDTDAUZ9CwGCc3BUa1pLF5E9d/9Mscc9w9kAWDkByNU3aL99A0ayldhxWLUyKOYGD9g7gT7sOKE+7DfiNt/MZhuOY6tt6wirHhTdjoGK4ZJ6pcT4N6fx/1eUMMLluEW7GMbO4g0j8AuWIEAgoBTAOoxywgmkUHISkDQSkldmOzztjZLZgnQcjMdbORYOC8IQFMCnpOuDNHnvZRrjvlq6z/xNexDetp4KAz6WazQzhvVjsPEfnrq+58548+pCH/fYwXQjAy75KkWDWuW6HCLr3fTLolq0603C1GmRCc0G+w7jd/YPX3v8/yFz8LSyOqEsBclGyaHHWnSde0a8Md3uSb8uPSsdMd/Yz4O7wm8aj+Gln/YjhgMfOAOekQ3M0W7eLPdY5Z0uuipJ5I7PN0ft6049AEt4PjuiV3OjlTMwEnabfeZXjzyLxBDn7zq1hwr/txzbs+xvhZf6GXEpGCQg03MbEwY5j1nefX/voXp48+7N7r11GQh8hoWUpW3ekVKuxqY5CoSeKufvyvTZz0ESD3goowp2xz/Uc+RfvCSxAnmDe8gfi0fTDJqAV2g8Jjx6OkXVF3M4ZNbzkhQOjuJXboLLvN8OggHLWke+q2e63u73Q7aVV1wnl0S2G5Q0TJxeFMsSxn8OH34U7f/iT9r3sho4M91NG0+zbzzOOz3nkctGS/s7jvse/YvP9SJDhaGnBWbZhXqLDbyxSmlBiqMPDPG7j6/aciW0dj3d+BhkioqLF3jLPYrFagQ2xcNSx3zhua5ERdE7D9BznwvS9l3gn3JIR2krytRnV3Cs/4zzd/qfGIh3x/Yy2nZkKonEeFCrsdJZ5MMuoG8xDGv/sLVn/h2+QSEDMsk6kDV2nsV1P2MTsGTGc/DKFAKDWWAhVj3bd+yZoz/o5KnXHVWSEltUc4DxEZX/GiZ7z3moMXjooJFWV7hQq7H7lkqPdxF0I8c8s2qz/4KcZ//1dyVQoxgtJ9dDcOrUo4bnuWB2qGEGj9/hxWv/mTDG5u0kbJPWRWZR47jcPudtdzWyfc9acrGzVU8urqqlBhtxu0gCKMOaXphExg6MbVXPWG99O+6hpy6PZJ2hi+s4xgU0r8FW4tWDaQ0uOsxF92JVe84V00blxDb1p97PFCNgtSuD0nhA/GnV/14g9dfNDC0U2uyjwqVJiBmxCfxlcNh5kxIEp57rlc8vp3EzZsxvkYNQeSSnfKPqTKPm6bYXYQ1m7g6te8G3/OP8g1EAg4i6y+VdnqNuLYu9zjHDvhHj+7XItIeVxdYxUq7Da0JODTtnm9zHBkBDxDTvA/+jXXvuUjsLUg81HCtro/b2/qAWHbOJe/9WTav/wNc52hlNQINIBmZpRaUbLfNphxj1e84IPZIx40stECXiHHI2KJQVPTG6rachUq7Gp4idopkdvVKAgEEdQ8c4FNX/wmV3z4VGi3UPPdRvnEckZsrNMhDzSLG3Ihbr3tNXftpA2+MOkxRfLEYmYW0nmwlNUBWKvk6vd8jG1f+i4D4rq8WZrEr7KgaNXzuO047tjjzul9yD1fd/3y+RTOURNDJOBV6UxoS+U8KlTY5XDBoUEpCBSpcOKDg+BwFljgm6z/n4+w6uRTkHazy0wbAgQfIPi4PSKWyIaSA0kLv7Nha3pXOg8jTph5pi5cmhlmlqSpwILhvSdQQrPFle/+OOMfOZVFoR3DYe+QkDEOjAOZTwJnlfO47XjYq1717U1HHXj5WFCa4sAg955x5ymEpAdWoUKF3epcgAOagXXv/BQ3vPeTZM0x1ErEoj6hmIvCTEmKqRChdOCdgsQNcdkbvIfGnZaQwtmMuPPS2X0RMUw9ieSE4BRxOdl4ydVv/x9G338KAz4gqphZVFWcnW9zz4OIbO078YEfuWrpXJqW4wwyKSldiYl1lY0rVKiwO42JJ9OSId9i+H2f5obXfwDbNkyRlxQSunxZmUVm2sxiMaZI0TkW9orz4FPGISSHuH3dyqK6pZs0haZbx7nyLR+iefKnWRZGIjllCIn7KkzSPK+cxx3GU/7z9d+a+4Kn/HpDniXRUU/Nl9SCUc7KU12hwt6NQmFbVuLUM+TH2Pipr3D5K94GN6zCaeQNCWLdjUEtoZZYfI2w18zzdmzPZKoWiwqy8YFCcBP7L2s3cMkr/4utp3yRHitoJ62RroyyKDoLs4891nmIyJbefznh9dccsd94S6ID6fVJ8F6qolWFCjNhNEOwWLsXzyAl9vWf8M+nv55wzkU4QqQW1wJzAUuqUzmRkXZvCfg6a8zdNRcxCjGaxD6HF6IyoEJ58cVc8eyX0PrqVxmQVtour01olNtEj6RyHrsQ933Qgy4o7nuXH12ROcYlI0cT62U1Jlihwu5GFoSeUnFBEYR+CywA6mf8jQuf9FLWfvk06kUTFU/hSsakiJmIRdbevQWThsuS85AuU3FmhngjU2PLz3/HRU9+Gf7Xv2NOFshCiTNH3fI9woLt0c7DQuCQZz/1vzedeHy5JcsoIGmBVdxXFSrsfqMpQIZJTmaOgKOtJQ016tdfzXWveDNXv+7DyJXrqYXYOA4iE6Ud9rIhe5lqaOtBEC+4LeOs/NAnuPJZr6Hn0itpZA3MMnLLqFka353lZ0JV93ySqAc88IHnHfCkx7x63dwB2q6WNAmiVKVsFwl0+dpmgVM3EbzEG04n3XwTmgC75+KZPFzgVZKac/wX24mLWK3z8wFn8eeCxE/gJvzc0/gewqTPNEqoynbvbnaZWLl5O3O7z4FHutvcnWssqlIYuhua0XGrXAko4+oYFzB1FNKmTwILRluMnPJ1LnjKK9n8/d/R27IoKCSdK22CDGvKlTdVhu8mzmaHV+kdtL+39OO2/ZPi/O2UO2ay7Kxg5BiKMXrupfzjRf/Bhjd/gEXDwwzi8KIEamSWRXoXLWeccF3S/weJ7yje54IXQwS2DQ/v+c4jeM+9n/+sb6455tCLtkgfkKPSMR8TD5ceEGuOM+1AAtByLi1cBUo1DCUn0lojZYpApt3ypp0toZnFGzgjRoWl3vp5yiyjdIZpSQPBxGg7j0fjQthuyAIlLVx1v1ZHkRZGI4n4bJji6VgSmTBMKej23HHNizjhI/GLIl1jGq/6jIDQ7qiST/O7DIhFGg0xIZSQhxz1DocwKG0G/3E21z33lVzz72+mdeGFIGU8PUEQbwRf4gmUFjqbIpA2QyZPLHWYejt7FFPUnG7GC9jO3RK3+PztHZsFT7CAN8NbwELAvMUVlhQg2tobueHkj3Hpk59O4wc/YF6IQwIFgVop5N4oUk9EbeaLVh3tk7aLs3C1kGHiaGlBnis/OOXj7BWqSiKy+fz/+80/rrvkP4+Ze+M4QSduUEvz4y50tWLirPUsUOHKQ/yASoS2g1qqFUv36gxMp2K7muFKIpFOUdJTBOoEBMHfqnTOhBMMEm+UkG6muhcyXNyCTQth05WBCIIrDYdBLkjpaXgjR1AcrTRv72Y8lEuLcOK7an06ydb5XXA91S3EcZ4U9poYJUpLHG0XyLzslutedmiBJd17LcYzJW+Psu1L3+bS3/+FuS98Ksuf9ySy/ZYBinhBfeqFpB0Qm/xGExyWrtAw4T1E4+NmLL/s1EcVpv7spBJGmJxZIFEGV5RShCzE68wkUNIk1waMjLHpV2ew6iOfR888m4WhhalOHDIxa5FJhzs7dl1S/SEdjKQMSpP0bq2nh71Gku/Yhz30v86607J77r9+/RGDQQmmXcGUSJEQMANHSMRtOqMFDTGoWdxJaYvgpUM+oJh0RTAn5HYna3sKU3UTdhQa7dRzAu3cE1BqNY/XSLjWTkW/PEi3hHDzUVpIS2Cdn1MawaWpEgWRdKZt547rthx/OpE+DxQYGYLVAl4MlzKPmHfajCreC6QsDDTEG7DzHpzEx03e+86eh070K4aXgNaiEmBG1MiuAbkp+Ayb4XQ7YATxEDy5Klmm6LXXsfWtH2bzd3/O3OeexLInPZFs2SLEIEggTKp+Srfk3ClHTjxM6Gbq3Wq83L6SoMmkX9atQ0nXYfiUIWQyMVWVhRDPuAiiRj42zpbTz2Tdp75N8X9/pt4eRVx8yTzorO/KdsM96wjwGmJGFoSRssVjX/7avcd5iMgN3//oRz96yXn/PPWeW4SsKLs6wTHpVYIEJM2VCzPfkpJJkacLQhaUcZR2Fg/SM91NKen+BxliDkNTWc9wFlL36Jbfg0t+oaWKWnwEjFIlRY06rWdbEiWNIGBZt4tkzI4mbKoap6/dTQya7aJryVnMrUOKFtVP/j0z3+gTBDWHEVJJJ5CL0W9G6/xL2PS6S1n7hR+z8KmPYdmTHkp+1IGTenBxkjJ0sk2TRG2SEgDtBDlT9ddlR9nQrZ4KpUQohK7zytNDPdRSPy9ozBhVAi6d87BpGxt//2fWff37jP/2LHq3baGuYM5TC4ozx56wCjlZf162u9c9htXr7FVi4Ce9+tXffuc3/velS86+7i6HiiO3SMTmkagvoDEW1VlCwlamaDkzcCWRo1QEr7FwbdTiNAo7TqFvrbx+a8+xNJGeARJyCnExCxFS2zNmJbc0lKcxgYnH7WIzqe2hBfGGsk6lXXbuuG7re+weR9SVNi+UBm0iaaZOGUmYyTplMmoaz4emb00u08v27+s2nK/4OcQn1pPTzgtLhc9AhseTzajhikY4j/eeD9SSKxjVaKoWBGPs4vPZ/M4L2fiZz7PgEQ9h6DEnMnjfu6ML5kfaoRDANCYFEgclROKor5t8ulLkGMzQW2HgnrxDIalGlhmpmZ+yf4mDOCaCJi+mBuoC1h6jddXVrPnpHxn97m8ozv8nWWuUQYXMBbSjfJqOM8PRnuXTVJ0BBk1/ifwAkVOrSI58r3IeIrL5zc985jPWjOn5iy69PhuySA1NMoGeGEn4WfC5mUApEvluOl5dwKxFw3sYKcj62yBuei+RIOAzaHoyFVpqNNLNszMutjsfk3pLWRBKCRRWMFAUMF7E7eLpvFmC4nw8gXlRkmmgJR7BqIUYPNgMRt4CuLJMtBTlRDkyNUad7KrfkiYcfLRs4jyjzmPi6RGHFDaj+wMBYVyiARKUpsS+TB7iUXlgiIIhHyhXrWL089/k+m/+FI48lMGHP4g59z+OweOPhXlDIFl61xOjV9IpM2Cx7aETYc+tVf6S/YiORAXzKV+2Tu4dQKOOCeKxkRbtq69j/Z//zvivTqf513PI16+jEUoGyKLhNY8LsYdYJlnZUpNC4B4wk2xM6LD4blFd8S7jiksvOXuv26UzM/3qi1721SWn/fyZK5rCQBlQ87GWr3EaJA+kwbkZdh4KNR8jpoBh4hk3wx94IPMe/AB8rt1a+fRcHBYnu8pepGwy9sufIxvW0qCO4WhpGwm3bG4cSikFXj1qdRohoyVjlGLIwUfQ88ATUl3LptE0K+oF74AN62n+4rf4chQ16PEujbHOnOMoMGoPvj/5EYfgfRGjVwQJ1r0WdtV5EIvXdcg97fMuovjL2UjuyLyQ+Zml7THoZtIhlYQUqHvIUmEvSxNIYyKgGc6UMhgjBBgYRPdfwbx7HYfe687Mvdth1I84EPr6EVeP3qKIxr9Tb7Gp7e1bPb7Jz4uV/hKxNoyM429Yy7rLrqJ1zsVsOfNcwmXXImvWk1tJpkZNjF4jZkfpVdoYhcZ2c+zjpIBh1teuFC+CShMXoJA6zgIZxpWNOv/TvPYde+Ui9vnnn33n857zsnMOvmx1vtg7GkUZ6ZEVnMV9hDALnEcAain6bGvAS0lmCtQZCSHFK9Nn9jyxuZxbD22EuTRjg5UGJTVGXBsXOvsbN2eylMIVmHhqvk6fKePaxCRQ+F62EchpT6tBahPlUQOKirEIx0jmMZT+Ik/U2DN3twrQpqRNSGOlkmb/ZYrGg9zB82AYWZpBGscYxJijOaMEXIgZuJ9xkxRwdOhIhDbQdqnpbIaXPGqECGAetcQ+awFvJS0yxgkYObWFC7Elixk46jDCMYejBy2n74D96V2ymPqC+dDXQFRAdcflwE5hP4SYcZhhW7fi125kZN0GRtesw65bhf3zKrZedDm6ah22eTOhtZmckh7yxGoRlx2LjqNyggVLU4eaBqVjj8dZwGFpmnE2Zx2xNOhoISa0tYaGEiEwfOjh9H/t5Pftlc7DzPSqP/75q1c985XPXL5uE3PKAp+WXOpJi+DWavm7KxLrHEHoTJEYqU7vuqZmWksJYrHBLYJaZ7lSsbQhIXLrKbZ1F4k0ZlECwTyCS4Zyeg23pdqZqWIhkEFSWotDCDYLiDJ1yjF0FhinRry74orqrCB6OsG3pbFgmRXnQVLtXLoBTNy76p4V0e7SnXbd7MTyhqXGkEsfusdoAqMEQlYjG+gn6+9H+/qQgV565s6hNjhArb8P8gyp5ZgKVpbQLgjjTYqRUcY2b6G1dRs0m+jIGOXIGOXIKK4sqSPUgTxRsMbh85DuUZtYqEvH51N7Lwt0xSF8J9NnYlBmthet4vmPQzSlFDjaNMlYd9idOPb00963V/U8JtUug5n992WPeuAzxj/3HemfdMcU6QKeLdNWU2e7J40WT6y9Te8xpA1xsQ6VvXQr1LqTli3yEklX/Cb+jOvc7tPfb+hMM4cJV6Whcy5nR3E57GCn3KbFNE+9tozZdR52VEKcLGwk223CT5SdZNJ9QncIQxB6gT4USo8ND6PDmwipZNQCRpN+SAlRG8PioEoMkdrkQB2hb8qVL1MenSSl7B6P2+HnKN1+YWd/x25yv+8ZFCxGkACWx5UHMbIQF09XyiiHtMZ1j98wvwVcPPjg+75r7bx+Wq6TWhotJylp3pvfeoUK+w46pb+AEMTRVqUUh9Mada3RJxlDkrFAa8xHWYxjIY55OIZo0Ct1Msmw7tqh68bd+642UNo5I6SSOigOb8J+Jz1q89IVK76811pQEQknPO0pH1h9/BGXbVbBmXanHKxyHBUq7KWOJMrcqoJ0eiZm1CyQh5KadbqdgUBJkECpgbYYbQeFsyk8afsuBDVFBLzGUnZGTlMz/vKHP3zRZdlle7UVFZHmp375kyecy/h40Mito8HSyF0lF1Whwt4GJS7c4uOOV4fmx6dORYnRkkBLQpTBVev2KuLTK8/RgUPI0/Z8HowCyA85hAe/+fWrg/d7fwj+D7Nrh57/1JVr63W8ZDiJi0PVRVKhwl4YMFpspndoMU2EUoRSotJhKR26mDhBngXIfeRjawSh7mVKD2Zfzjwk9YXyEIkum5pzqR8tDrj3PdfBXrYkeDPZx9jG1rYn/PjM887uuXJtfU6zTJMSlfOoUGFvQ4fdamIQxbozJ2aTiFHpUJZI11R2EKqqBHTJUX2awnRswsj/5d5bFixacFony9vrsbBv7kXXr5j30YuaG4NTN8MbHhUqVJhOTKZ76fBgZSbU0iRUh8U48ldFYtKWQjM9qp5HorsXSdNlkaVhW3+dTTX3XdJQpdsXToSFwO+vuOz3/7z+2qfn16yaP9RsRq1gjfQIJlDvLusJNetsfVeoUGHPi5m5iRBc5+vJo7ImU7/uoCpbQZZyMe8CIoExV+f6/ReVj/japz7a29t/wT6TeQCoqv+XT3/8Sb9rFNeO1CLjVW5xy7qRpi+8QG6WFNEqVKiwxweO0JWT2l50SyyRGya9n86jQtqvkUgiW4pjPcbY3Y/aNG/ewu93beo+czLMmFevX7joRc9+68VL5oPWKQn04chDwEug0MBkGcwKFSpU2BdRqtFWjwbBa4M1C+eEI577tG/Q1alk3+sam1n2zJ6hDz3V+l9zTJnR54tIcaCBUiELSm+IetBldQ1VqFBhH0TQQJCSPt9gpeac98jjL3zZb350bNlssc9lHl1vKVJ+s7n1tX8c9J9bWcsw6cGLUIqSBaMeIgNQlb1WqFBhX4UGo+6hwLFuv8UsOelfTy1branP2VdPzsnrbvyfLc95zOornSA06AlKI2QEjHExymoJvUKFCvsoTJRAxvosZ/QR91796Be84Mzta/nZvnpyROTyUbPH/knyH1//uR8s29/nFFLgNTJlqt8Z+ZgKFSpU2PsQVNlgsOF+d+LRn/7IG0Xk/O2fk+3LJ6hP5JxRs8eeIeEnl3/xu0sPKh25LxBxU9gwK1SoUGGfCawBb8bmg5Yz9+XP+nwPnLaj52X7+onqOJA/iP/xys/8YOlh1GhTIAJqkVlyov8R0mKMx3dkPytUqFBhD8PEjovgRXEBeohzpk313FhX/tTa9OV3PeVJLxGRUDmPm3cgZ7fNHvvTtv/htq/9aPkRPqee1MUylAIjJLU6S2z/GVTTWBUqVNgjUYpgAg0f5bYQSVopGRsbOSNPP3Hj8z//8Q/cnOPoZCgVEi5Zt27pB5cf+ObH2LxX3qVU1HlC8NSTClshMQupBYcmps4KFSpU2PMyjxxByK2FSZNSDHM9rPN1Ljh6xfpnXvinJ/aKnH5Lr1FlHpNw9KJFa4BXNw7aD5cPvXLx1TcwN89wRTsmdGJgjkKiPohUe+gVKlTYA9GwGPhucx5cSYayyTznHTSX85b3vvTFt+I4oCra79Apn3r1Za/uf+sr3n/J3Q+/cX29F7MGmTnyYCge0wKTahOkQoUKe6CBwyikpKBFboBvsKGsc97yBRzy7jd87VO//OUvduZ1qrLVLZ1ks7t+/lnP/9iyM//5gAOvWcU8DE9BiRFEUHPVSapQocIegUAg04xgcXtcMVxosKlngLMPXTB8/sLshR/77W9/KiJF5Tx2AbLeHn73la+/79L3fuz5h121bsmS0Sa9lBB8dfoqVKiwR0BEulmHmVEDWmRckikbHnC3s5/y2x+8rlfkT7fpNavTunNYY3bQWw885OX3y+a+fsWqjSwtPYOl72oHVKhQocIs9h6oCBY8mSjrDFYfdQjylEd86ZHvfPO/cTuGRyvncVuykEadH37u049a/eP/e9PgOf88YfkNG1jolYEQgJK2eEyjcLwFA1EsDsJ1HYyaoHS0BTq6ZxMfhnFTzQF2tOze/V7luipU2Mss/Xb3dRyjjdok8cYPOzDikp7hscg3n6xP3EgTMGW8lrPGWtSe86TNvS982reOvc99XyUi5e09ygq32Yk0+O7nPveotT/99Zt6/nrxCYeuWseyUIK18RL3QRwZJYpa/FCDKCXxM82D4jA8YZJYjXSdxUQ2k1zPJOcx2Wd0hKxu4lfi2Hb3T25yKd7889iJ19rRi93e16qOqzqu6rgmnoeQHIVMiRPVOkFmdAYlSpD4NwVyESwYJkpbDEdAzWOU0Y9IjTVZjfOHasPznvuEb9/rHf/5qf0H5l90R11chdsJ16jzvc99/lHX/Pgnb2qcce6hR7bdkgWbmgwFyKKQI7UkIh8wWkQyfJF4IZh1ttclXggyoWgWncytj8NVeUeFCnuhbUl3fhSyCl1Rq459cGYEgVJjEKnBoja7ghejxzvElK0KW/obXNdfG153l8OuPYfRf/vi//3ubMIdnxatnMeuSDLzjJXtYv8/vPXtLx7+7ZkvPmYzi3uuuI65lGjpcXhqqvhQ4iXgXbwcNOS4IGREPeBOKSteKDIp6bw5x9FJT2+mprWzYRO3QAK5o+fsKGza4e+kOq7quKrjuk3HFT1ESLlH51k+1bo7f6/7aPxLjRYiC0YmQjBPEGXU1VnfW2f9ocs21x52wjeOfPWLTz1sxcEX40NKcaicx6yCCmM+rLji9DMe/su3v3t5/2jz+Yu3FCsO3NyW3nWb6FfF+xaCxxEjhE4U0SlfmcRpiIAk/hm5lSyjyj0qVNirglFkio66dWyBTPQ6asFQBElhZ4Ew5hwjgz1slGL48sHatsVPeMSPH/CON566bO7iS/B+Go6zwvQgy3jyN79Z+9JjH/uwP/73R+7xz5/84t+y8y/lbksOXTgULGuuWs0CVTIxWr4gCGSqKEbmjdhqD9xUErdyFhUq7M0woNSYqTgTFEFTGcvjY58jbzAeYIsvKect4MbBBlf22pqDHnniafd6zcs+9bDPvfbaS977gzbl9DHwVc5jN+G1Z57W85H7PoWRzSP3P/3UU+9pm7a9+PwvfLE2ODYuRyw/dPGSWj9jN6wijI3jQ4kg1MRo1GoTM9pmeB8vBhHFOXdLSfKUpHyHDbodPG9nXuv2Pq86ruq4quO65ePq/F0MfFlS+LiU3MLQnl4Gly+jKcaZa6/amB99VLHsPve65tyLL/jef5z29dUP/8J7fviXN358nHL30CZVzmOmoguz/vnzRe9x+H7ZN79z1kvqWa3vdfe7Hwfut+JOxz/4YY/vy3KGr7uBX3/rm2QSB/QWLF3GgUcfhYXA8Lr1XHXJxVQ77hUq7H3wBsc/9CEcdq/jaYbA8PhY+ZOvfemTb/n+90aGli7m26d99fOXaHPTx1/1jqaItGfiGP8/0Wj2pKG9Vz0AAAAASUVORK5CYII=";

  const labelPages = labels.map(lb => {
    const itemDisp = lb.items.length === 0   ? '—'
      : lb.items.length <= 3 ? lb.items.join('<br>')
      : `CIPL #${lb.shipment || '—'}`;

    const measDisp = (lb.length && lb.width && lb.height)
      ? `${lb.length} &nbsp; ${lb.width} &nbsp; ${lb.height} &nbsp; CM`
      : '—';

    const qrDataUrl = _buildQrDataUrl(_buildQrPayload(lb, consignee));
    const qrBlock = qrDataUrl
      ? `<img class="label-qr-img" src="${qrDataUrl}" alt="QR" data-len="${qrDataUrl.length}">`
      : `<div class="label-qr-err">QR niet beschikbaar<br>${esc(_qrLastError || 'onbekende fout')}</div>`;

    return `
<div class="label-page">
  <div class="label-header">
    <div class="label-company">
      <div class="label-company-name">IHC Holland B.V.</div>
      <div class="label-company-sub">Property of IHC Holland B.V.</div>
    </div>
    <div class="label-project">
      <div class="label-project-cap">PROJECT:</div>
      <div class="label-project-val">${esc(lb.project || '—')}</div>
    </div>
    <div class="label-brand">
      <span class="label-brand-royal">Royal</span>
      <img class="label-brand-logo" src="data:image/png;base64,${LOGO_B64}" alt="IHC">
    </div>
  </div>

  <div class="label-pkgnr">${esc(lb.packageNr)}</div>
  <div class="label-pkgnr-cap">PACKAGE NUMBER</div>

  <div class="label-divider"></div>

  <div class="label-body">
    <div class="label-rows">
      <div class="label-row"><span class="lbl">SHIPPING MARKS:</span><span class="val">${esc(consignee)}</span></div>
      <div class="label-row"><span class="lbl">ITEM NUMBER:</span><span class="val item-val">${itemDisp}</span></div>
      <div class="label-row"><span class="lbl">TYPE OF PACKAGE:</span><span class="val">${esc(lb.pkgType || '—')}</span></div>
      <div class="label-row"><span class="lbl">GROSS WEIGHT:</span><span class="val">${esc(lb.gross || '—')} &nbsp; KG</span></div>
      <div class="label-row"><span class="lbl">NETT WEIGHT:</span><span class="val">${esc(lb.nett || '—')} &nbsp; KG</span></div>
      <div class="label-row"><span class="lbl">MEASUREMENT:</span><span class="val">${measDisp}</span></div>
      <div class="label-row"><span class="lbl">DELIVERY REF:</span><span class="val">${esc(lb.deliveryRef || '—')}</span></div>
      <div class="label-row"><span class="lbl">REMARKS:</span><span class="val">${esc(lb.remarks || '—')}</span></div>
    </div>

    <div class="label-qr">
      ${qrBlock}
      <div class="label-qr-cap">SCAN FOR<br>PACKAGE INFO</div>
    </div>
  </div>
</div>`;
  }).join('');

  const now = new Date().toLocaleDateString('nl-NL');
  const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>Labels — ${now} (build ${VAL_BUILD})</title>
<style>
/* Standaard IHC shipping mark: liggend, verhouding 11.69 x 8.27 inch — één label per pagina */
@page { size: 11.69in 8.27in; margin: 0; }
* { box-sizing: border-box; margin: 0; padding: 0; }
body { font-family: Arial, sans-serif; background: #fff; color: #111; }

.label-page {
  width: 11.69in; height: 8.27in;
  padding: .55in .55in .45in .55in;
  display: flex; flex-direction: column;
  page-break-after: always;
  break-after: page;
  overflow: hidden;
}
.label-page:last-child { page-break-after: avoid; break-after: avoid; }

.label-header {
  display: flex; align-items: flex-start; justify-content: space-between;
  margin-bottom: 10mm;
}
.label-company-name { font-size: 16pt; font-weight: 800; }
.label-company-sub   { font-size: 9pt; color: #6b7280; margin-top: 1mm; }

.label-project { text-align: center; }
.label-project-cap {
  font-size: 9pt; font-weight: 700; color: #6b7280;
  letter-spacing: .08em; margin-bottom: 1mm;
}
.label-project-val { font-size: 20pt; font-weight: 800; }

.label-brand { display: flex; align-items: center; gap: 3mm; }
.label-brand-royal { font-size: 14pt; font-weight: 600; color: #374151; }
.label-brand-logo  { height: 9mm; }

.label-pkgnr {
  font-size: 56pt; font-weight: 900; line-height: 1;
  letter-spacing: -.01em;
}
.label-pkgnr-cap {
  font-size: 10pt; font-weight: 700; color: #6b7280;
  letter-spacing: .1em; margin-top: 2mm;
}

.label-divider {
  border-top: 1px solid #d1d5db;
  margin: 7mm 0 7mm 0;
}

.label-body { display: flex; justify-content: space-between; align-items: flex-start; }

.label-rows { flex: 1; }
.label-row {
  display: flex; align-items: baseline;
  padding: 2mm 0;
}
.lbl {
  font-size: 11pt; font-weight: 700;
  white-space: nowrap; flex: 0 0 auto;
  width: 40mm;
}
.val { font-size: 11pt; font-weight: 500; }
.item-val { font-size: 10pt; line-height: 1.4; }

.label-qr {
  width: 38mm; margin-left: 10mm; text-align: center; flex-shrink: 0;
}
.label-qr-img { width: 38mm; height: 38mm; image-rendering: pixelated; }
.label-qr-err { width: 38mm; height: 38mm; border: 1px dashed #c00; color: #c00; font-size: 7pt; padding: 2mm; word-break: break-word; }
.build-tag { position: fixed; right: 4px; bottom: 4px; font: 10px monospace; color: #999; background: #fff; padding: 2px 4px; }
@media print { .build-tag, .qr-diag { display: none; } }
.qr-diag { position: fixed; left: 6px; bottom: 6px; max-width: 70%; white-space: pre-wrap; font: 11px/1.35 monospace; color: #111; background: #fffbe0; border: 1px solid #d4b106; padding: 6px 8px; z-index: 9999; }
.label-qr-cap {
  font-size: 7.5pt; font-weight: 700; color: #6b7280;
  letter-spacing: .06em; margin-top: 2mm; line-height: 1.3;
}
</style>
</head>
<body>
${labelPages}
<div class="build-tag">validator.js build ${VAL_BUILD}</div>
<div id="qr-diag" class="qr-diag"></div>
<script>
(function(){
  function run(){
    var out=[],imgs=document.querySelectorAll('.label-qr-img'),errs=document.querySelectorAll('.label-qr-err');
    out.push('QR-DIAGNOSE (alleen op scherm, wordt niet geprint)');
    out.push('build: ${VAL_BUILD}');
    out.push('labels: '+document.querySelectorAll('.label-page').length+' | QR-afbeeldingen: '+imgs.length+' | foutblokken: '+errs.length);
    errs.forEach(function(e,i){out.push('fout '+(i+1)+': '+e.textContent.replace(/\\s+/g,' '));});
    imgs.forEach(function(im,i){
      var r=im.getBoundingClientRect(),cs=getComputedStyle(im);
      out.push('QR '+(i+1)+': geladen='+im.complete+' natuurlijk='+im.naturalWidth+'x'+im.naturalHeight+' getoond='+Math.round(r.width)+'x'+Math.round(r.height)+' display='+cs.display+' visibility='+cs.visibility+' data-url-lengte='+im.getAttribute('data-len'));
      var pg=im.closest('.label-page');if(pg){var pr=pg.getBoundingClientRect();out.push('   buiten pagina? '+((r.right>pr.right+1||r.bottom>pr.bottom+1)?'JA (afgesneden)':'nee'));}
    });
    out.push('venster: '+innerWidth+'x'+innerHeight+' | browser: '+navigator.userAgent);
    var d=document.getElementById('qr-diag');d.textContent=out.join('\\n');
    try{console.log(out.join('\\n'));}catch(e){}
  }
  window.addEventListener('load',function(){setTimeout(run,150);});
  document.querySelectorAll('.label-qr-img').forEach(function(im){im.addEventListener('error',function(){setTimeout(run,0);});});
})();
</script>
</body></html>`;

  const win = window.open('', '_blank', 'width=1000,height=760');
  if (!win) { alert('Pop-up geblokkeerd — sta pop-ups toe.'); return; }
  win.document.write(html);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 500);
}

// ── HS-code validatie ─────────────────────────────────────────────────────────
// 1. Local check: GN_CODES set (gn_codes_2026.js, 22K valid codes)
// 2. Measures API: tariffnumber.com/api/v1/cnDuties (free, 10 req/min)
//    → Try direct browser fetch first, then GAS proxy (TARIFF_PROXY_URL)
// 3. Expand panel: shows Vietnam/ERGA OMNES export restrictions + footnotes
// 4. Fallback: deep-link to tariffnumber.com page

// EU TARIC error message for invalid codes (matches official text)
const TARIC_INVALID_MSG =
  'The goods code is not or no longer valid in the European Union. ' +
  'Please try entering the first 6 digits and browse the nomenclature ' +
  'until you find a proper description corresponding to your product.';

// Measure types that are relevant for export (Royal IHC ships goods FROM EU)
const EXPORT_TYPES = new Set([
  'Export authorization (Dual use)',
  'Export control on restricted goods and technologies',
  'Export control',
  'Restriction on export',
]);

// Countries always shown (besides Vietnam)
const HIGHLIGHT_ORIGINS = new Set(['VN','1008','1011']); // Vietnam, All third countries, ERGA OMNES

const _hsCache = new Map(); // taric10 → { valid, desc, measures }

function _toTaric10(input) {
  const clean = String(input || '').replace(/\s|\./g, '');
  if (!/^\d{8,10}$/.test(clean)) return null;
  return clean.padEnd(10, '0').slice(0, 10);
}

function _checkGNCodes(t10) {
  if (typeof GN_CODES === 'undefined') return null;
  const valid   = GN_CODES.has(t10);
  const h8      = t10.slice(0, 8);
  const h6      = t10.slice(0, 6);
  const h4      = t10.slice(0, 4);
  const desc    = (typeof GN_DESC_HEADINGS !== 'undefined')
    ? (GN_DESC_HEADINGS[h8] || GN_DESC_HEADINGS[h6] || GN_DESC_HEADINGS[h4] || '') : '';
  return { valid, desc };
}

function _tariffPageLink(t10) {
  return `https://www.tariffnumber.com/2026/${t10.replace(/0+$/,'')}`;
}

// Fetch trade measures from tariffnumber.com API
// Tries: direct → GAS proxy
async function _fetchMeasures(t10) {
  const base = `https://www.tariffnumber.com/api/v1/cnDuties?term=${t10}&lang=en&year=2026`;
  const gasBase = typeof TARIFF_PROXY_URL !== 'undefined' && TARIFF_PROXY_URL
    ? `${TARIFF_PROXY_URL}?code=${encodeURIComponent(t10)}&lang=en&year=2026` : null;

  const candidates = [
    base,
    ...(gasBase ? [gasBase] : []),
  ];

  for (const url of candidates) {
    try {
      const resp = await fetch(url, {
        headers: { 'Accept': 'application/json' },
        signal: AbortSignal.timeout(7000)
      });
      if (!resp.ok) continue;
      const data = await resp.json();
      if (data && Array.isArray(data.duties)) return data.duties;
    } catch { /* try next */ }
  }
  return null; // all failed → show link only
}

// ── Expand panel ──────────────────────────────────────────────────────────────
function toggleValHSMeasures(t10, badgeEl) {
  const td = badgeEl.closest('td');
  const tr = td?.closest('tr');
  if (!tr) return;

  const next = tr.nextElementSibling;
  if (next?.classList.contains('hs-measures-row')) {
    const panel = next.querySelector('.hs-measures-panel');
    if (panel?.classList.contains('open')) {
      panel.classList.remove('open');
      setTimeout(() => { if (next.parentNode) next.remove(); }, 350);
    } else { panel?.classList.add('open'); }
    return;
  }

  const cached    = _hsCache.get(t10);
  const pageLink  = _tariffPageLink(t10);

  const expandRow = document.createElement('tr');
  expandRow.className = 'hs-measures-row';
  const expTd = document.createElement('td');
  expTd.colSpan = 99;

  const renderPanel = (cached) => {
    if (!cached?.valid) {
      return `<div class="hs-measures-inner hs-restricted">
        <div style="color:#ef4444;font-weight:700;margin-bottom:.5rem">
          ✗ Ongeldige GN-code: <code>${esc(t10)}</code>
        </div>
        <div style="font-size:.78rem;color:var(--muted);line-height:1.55">${esc(TARIC_INVALID_MSG)}</div>
        <div style="margin-top:.6rem">
          <a href="${esc(pageLink)}" target="_blank" class="hs-taric-btn">🔍 Zoek in nomenclatuur ↗</a>
        </div>
      </div>`;
    }

    const duties   = ('measures' in (cached || {})) ? cached.measures : undefined;
    const descHtml = cached.desc
      ? `<div style="color:var(--muted);font-size:.68rem;margin:.2rem 0 .5rem">${esc(cached.desc)}</div>` : '';
    const pageBtn  = `<a href="${esc(pageLink)}" target="_blank" class="hs-taric-btn">📋 tariffnumber.com ↗</a>`;

    if (duties === undefined) {
      // Nog niet opgehaald — de async fetch hieronder is net gestart
      return `<div class="hs-measures-inner hs-clean">
        <span style="font-weight:700">✓ <code>${esc(t10)}</code></span> — geldig in EU CN 2026
        ${descHtml}
        <div style="margin-top:.5rem;display:flex;gap:.5rem;align-items:center;flex-wrap:wrap">
          ${pageBtn}
          <span style="color:var(--muted);font-size:.65rem" id="hs-loading-${t10}">⏳ Maatregelen laden…</span>
        </div>
      </div>`;
    }

    if (duties === null) {
      // API onbereikbaar (CORS/netwerk) — NIET stilzwijgend als "geen maatregelen" tonen,
      // dat zou exportcontrole/dual-use-restricties kunnen verbergen.
      return `<div class="hs-measures-inner hs-unknown">
        <span style="font-weight:700">? <code>${esc(t10)}</code></span> — geldig in EU CN 2026
        ${descHtml}
        <div style="margin-top:.5rem;color:#f59e0b;font-size:.7rem">
          ⚠ Maatregelen (export/dual-use) konden niet worden opgehaald (netwerk/CORS).
          Controleer handmatig op ${esc(pageLink)} vóór export.
        </div>
        <div style="margin-top:.5rem">${pageBtn}</div>
      </div>`;
    }

    if (duties.length === 0) {
      return `<div class="hs-measures-inner hs-clean">
        <span style="font-weight:700">✓ <code>${esc(t10)}</code></span> — geldig, geen maatregelen gevonden
        ${descHtml}
        <div style="margin-top:.5rem">${pageBtn}</div>
      </div>`;
    }

    // Filter measures: export measures (all countries) + any measure for Vietnam/ERGA OMNES
    const exportMeasures = duties.filter(d => EXPORT_TYPES.has(d.measure_type));
    const vnMeasures     = duties.filter(d => d.origin_code === 'VN');
    const allThirdMeasures = duties.filter(d => ['1008','1011'].includes(d.origin_code));

    const hasDualUse  = exportMeasures.some(d => /dual.use/i.test(d.measure_type));
    const hasRestrict = exportMeasures.some(d => /restriction|control/i.test(d.measure_type));

    const flagHtml = [
      hasDualUse  ? `<span class="hs-flag hs-flag-warn">⚠️ DUAL USE — Reg. ${exportMeasures.find(d=>/dual.use/i.test(d.measure_type))?.legal_base||''}</span>` : '',
      hasRestrict ? `<span class="hs-flag hs-flag-alert">⛔ EXPORT CONTROL</span>` : '',
    ].filter(Boolean).join(' ');

    const renderDutyRow = (d) => {
      const isExport   = EXPORT_TYPES.has(d.measure_type);
      const isVN       = d.origin_code === 'VN';
      const isErga     = ['1008','1011'].includes(d.origin_code);
      const highlight  = isExport || isVN;
      const cls        = isExport ? 'hs-duty-export' : isVN ? 'hs-duty-vn' : isErga ? 'hs-duty-erga' : 'hs-duty-other';
      return `<div class="hs-duty-row ${cls}">
        <span class="hs-duty-origin">${esc(d.origin||'—')}</span>
        <span class="hs-duty-type">${esc(d.measure_type||'')}</span>
        <span class="hs-duty-reg">${esc(d.legal_base||'')}</span>
        ${d.duty ? `<span class="hs-duty-val">${esc(d.duty.trim())}</span>` : ''}
      </div>`;
    };

    // Show: export measures + VN/ERGA measures. Collapse the rest.
    const shown  = [...new Map([...exportMeasures, ...vnMeasures, ...allThirdMeasures].map(d=>[d.legal_base+d.origin_code,d])).values()];
    const others = duties.filter(d => !shown.includes(d));

    const shownHtml  = shown.map(renderDutyRow).join('');
    const othersHtml = others.length
      ? `<details style="margin-top:.4rem">
          <summary style="cursor:pointer;font-size:.65rem;color:var(--muted);padding:.2rem 0">
            + ${others.length} andere maatregel(en) (tariefpreferenties)
          </summary>
          ${others.map(renderDutyRow).join('')}
        </details>` : '';

    return `<div class="hs-measures-inner ${exportMeasures.length ? 'hs-restricted' : 'hs-clean'}">
      <div style="display:flex;align-items:center;flex-wrap:wrap;gap:.4rem;margin-bottom:.35rem">
        <span style="font-weight:700">✓ <code>${esc(t10)}</code></span>
        <span style="color:var(--muted);font-size:.65rem">— ${duties.length} maatregel(en)</span>
        ${flagHtml}
        <span style="margin-left:auto">${pageBtn}</span>
      </div>
      ${descHtml}
      ${shownHtml}
      ${othersHtml}
    </div>`;
  };

  expTd.innerHTML = `<div class="hs-measures-panel">${renderPanel(cached)}</div>`;
  expandRow.appendChild(expTd);
  tr.insertAdjacentElement('afterend', expandRow);
  requestAnimationFrame(() => expTd.querySelector('.hs-measures-panel')?.classList.add('open'));

  // Async-fetch measures if not cached yet
  if (cached?.valid && !cached.measures) {
    (async () => {
      const duties = await _fetchMeasures(t10);
      const newCached = { ..._hsCache.get(t10), measures: duties };   // null = onbekend, [] = bevestigd leeg
      _hsCache.set(t10, newCached);
      // Update panel content in-place
      const panel = expTd.querySelector('.hs-measures-panel');
      if (panel) panel.innerHTML = renderPanel(newCached);
    })();
  }
}

// ── Bulk HS check after validation ───────────────────────────────────────────
async function checkHSCodesLive() {
  const codeMap = new Map();
  _valRows.forEach((row, ri) => {
    const t10 = _toTaric10(row.cells[COL.O]);
    if (!t10) return;
    if (!codeMap.has(t10)) codeMap.set(t10, []);
    codeMap.get(t10).push(ri);
  });
  if (!codeMap.size) return;

  codeMap.forEach((idxs) => idxs.forEach(ri => _setHSCellState(ri, 'loading', '⏳')));

  for (const [t10, rowIdxs] of codeMap) {
    const local = _checkGNCodes(t10);
    const cached = { valid: local?.valid ?? null, desc: local?.desc || '' };   // measures: (nog) niet gezet = niet opgehaald
    _hsCache.set(t10, cached);

    let stateClass, badgeHtml;
    if (cached.valid === null) {
      // GN_CODES not loaded
      stateClass = 'hs-unknown';
      badgeHtml  = `<span class="hs-badge hs-unknown-badge"
        onclick="toggleValHSMeasures('${esc(t10)}',this)"
        title="GN_CODES niet geladen — klik voor tariffnumber.com">↗</span>`;
    } else if (!cached.valid) {
      stateClass = 'hs-invalid';
      badgeHtml  = `<span class="hs-badge hs-invalid-badge"
        onclick="toggleValHSMeasures('${esc(t10)}',this)"
        title="Niet geldig in EU CN 2026 — klik voor details">✗</span>`;
      rowIdxs.forEach(ri => { _valRows[ri].errors['O'] = TARIC_INVALID_MSG.slice(0,80) + '…'; });
    } else {
      stateClass = 'hs-ok';
      badgeHtml  = `<span class="hs-badge hs-ok-badge"
        onclick="toggleValHSMeasures('${esc(t10)}',this)"
        title="${esc(cached.desc.slice(0,50) || 'Geldig — klik voor maatregelen')}">✓</span>`;
    }

    rowIdxs.forEach(ri => _setHSCellState(ri, stateClass, badgeHtml));
  }

  // Ongeldige HS-codes worden pas hier (async) bekend — de samenvatting, de
  // AFS/leverancier-mailknop en de validatie-log zijn vlak na runValidation()
  // al vastgezet en missen deze fouten dus. Bijwerken zodra het klaar is.
  const sumEl = document.getElementById('val-summary');
  if (sumEl) {
    let totalErrors = 0, totalWarnings = 0;
    for (const r of _valRows) {
      totalErrors   += Object.keys(r.errors || {}).length;
      totalWarnings += Object.keys(r.warnings || {}).length;
    }
    const cls = totalErrors === 0 ? 'var(--green)' : 'var(--red)';
    sumEl.innerHTML =
      `<span style="color:${cls};font-weight:700">${totalErrors === 0 ? '✅' : '❌'} ${totalErrors} fout(en)</span>` +
      `<span style="color:var(--amber)">  ⚠️ ${totalWarnings} waarschuwing(en)</span>` +
      `<span style="color:var(--muted)">${_valRows.length} rijen gevalideerd</span>` +
      `<span style="color:var(--teal)">  🔎 HS-nomenclatuur gecontroleerd</span>`;
  }
  if (typeof _valHighlightMailBtn === 'function') _valHighlightMailBtn();
  // GEEN nieuwe _valLogValidatie()-aanroep hier: die telt elke aanroep als een
  // aparte validatie (nieuw event-id). Die stond al vóór checkHSCodesLive() in
  // runValidation() en blijft dus de bron voor de teller — dat telt correct,
  // alleen mist die ene telling nog de HS-nomenclatuurfouten (bekende beperking).
}

function _setHSCellState(rowIdx, stateClass, badgeHtml) {
  const td = document.querySelector(`[data-row="${rowIdx}"][data-col="${COL.O}"]`)?.closest('td');
  if (!td) return;
  // Alleen de hs-* statusklasse vervangen, de rest (val-cell-err/-warn/-ihc/...) intact laten.
  td.className = td.className
    .split(/\s+/)
    .filter(c => c && !/^hs-/.test(c) && c !== 'loading')
    .concat(stateClass)
    .join(' ');
  const iconEl = td.querySelector('.hs-icon');
  if (iconEl) iconEl.innerHTML = badgeHtml || '';
}