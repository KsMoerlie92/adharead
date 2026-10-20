/* fat-data.js — gedeelde FAT-datalaag (Admin + Large Item Overview)
 *
 * 1) Admin: leest het blad "FAT" uit de bedrijfsbrede Expediten-Excel en maakt fat-list.json.
 * 2) Large Item Overview: laadt fat-list.json (naast expediting-data.json in shared/).
 * 3) IFS-terugschrijven: leest een !IFS.COPYOBJECT-klembord (CCmtPoPartLine), past onze
 *    invulling toe en geeft de bijgewerkte tekst terug om in IFS te plakken.
 *
 * Regel-sleutel (uniek per FAT-regel): Order No | Line No | Release No | Requirement Id.
 */
(function (global) {
  'use strict';
  const FILE = 'fat-list.json';
  const SRC = (document.currentScript && document.currentScript.src) || '';
  const trim = v => String(v == null ? '' : v).trim();
  const pad = n => String(n).padStart(2, '0');
  const low = s => trim(s).toLowerCase();

  // ── Excel → tabel ─────────────────────────────────────────────────────────
  function cellOut(v) {
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      const d = new Date(v.getTime() + 12 * 3600 * 1000);        // +12u: vangt SheetJS-afrondingen af
      return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    }
    if (v === undefined || v === '') return null;
    return v;
  }
  function headerOf(rows) {                       // eerste rij met Order No + Line No
    for (let i = 0; i < Math.min(rows.length, 10); i++) {
      const l = (rows[i] || []).map(low);
      if (l.includes('order no') && l.includes('line no')) return i;
    }
    return -1;
  }
  // Het blad "FAT"; anders elk ander blad (niet "Result") met de kopkolommen Design Object ID + Type Description.
  function findSheet(wb) {
    const names = wb.SheetNames || [];
    const byName = names.find(n => low(n) === 'fat');
    if (byName) return byName;
    for (const n of names) {
      if (low(n) === 'result') continue;
      const ws = wb.Sheets[n]; if (!ws || !ws['!ref']) continue;
      const top = global.XLSX.utils.sheet_to_json(ws, { header: 1, range: 'A1:CZ10', defval: null });
      const hi = headerOf(top); if (hi < 0) continue;
      const l = top[hi].map(low);
      if (l.includes('design object id') && l.includes('type description')) return n;
    }
    return null;
  }
  function tableFromSheet(ws) {
    const raw = global.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const hi = headerOf(raw);
    if (hi < 0) throw new Error('FAT-blad: kolommen "Order No" en "Line No" niet gevonden.');
    const keep = []; const headers = [];
    raw[hi].forEach((h, i) => { if (trim(h)) { keep.push(i); headers.push(trim(h)); } });
    const iOrder = headers.findIndex(h => low(h) === 'order no');
    const rows = [];
    for (let r = hi + 1; r < raw.length; r++) {
      const row = raw[r] || [];
      if (!trim(row[keep[iOrder]])) continue;
      rows.push(keep.map(i => cellOut(row[i])));
    }
    return { headers, rows };
  }
  function buildCommitJSON(table, meta) {
    return JSON.stringify({
      meta: { filename: (meta && meta.filename) || '', uploaded: new Date().toISOString(), rows: table.rows.length, version: 1 },
      headers: table.headers, rows: table.rows
    });
  }

  // ── laden ─────────────────────────────────────────────────────────────────
  async function load() {
    const url = (SRC ? new URL(FILE, SRC).href : '../shared/' + FILE) + '?v=' + Date.now();
    const resp = await fetch(url, { cache: 'no-store' });
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const o = await resp.json();
    if (!o || !Array.isArray(o.headers) || !Array.isArray(o.rows)) throw new Error('fat-list.json heeft een onverwacht formaat');
    return o;
  }

  // ── sleutel ───────────────────────────────────────────────────────────────
  const key = (order, line, rel, req) => [order, line, rel, req].map(trim).join('|');

  // ── IFS-klembord (!IFS.COPYOBJECT) ────────────────────────────────────────
  function parseIfs(text) {
    const eol = /\r\n/.test(text) ? '\r\n' : '\n';
    const head = [], blocks = []; let cur = null;
    for (const ln of String(text).split(/\r?\n/)) {
      if (/^\$RECORD=/.test(ln)) { cur = { start: ln, lines: [] }; blocks.push(cur); continue; }
      if (cur) cur.lines.push(ln); else head.push(ln);
    }
    for (const b of blocks) {
      while (b.lines.length && b.lines[b.lines.length - 1] === '') b.lines.pop();
      b.f = {};
      b.lines.forEach((ln, i) => { const m = /^-\$(\d+):([^=]*)=(.*)$/.exec(ln); if (m) b.f[m[2]] = { i, n: m[1], v: m[3] }; });
      b.key = key((b.f.ORDER_NO || {}).v, (b.f.LINE_NO || {}).v, (b.f.RELEASE_NO || {}).v, (b.f.REQ_ID || {}).v);
    }
    while (head.length && head[head.length - 1] === '') head.pop();
    return { eol, head, blocks };
  }
  const ifsDate = iso => /^\d{4}-\d{2}-\d{2}/.test(iso || '') ? iso.slice(0, 10) + '-00.00.00' : '';
  const flatTxt = s => String(s == null ? '' : s).replace(/\s*[\r\n]+\s*/g, ' ');
  const IHC_MARK = /^IHC aanwezig:\s*(.*)$/i;
  function ihcFromAddInfo(addInfo) {              // → naam uit Additional Information, of ''
    for (const s of String(addInfo || '').split(/\s*\|\s*/)) { const m = IHC_MARK.exec(s.trim()); if (m) return m[1].trim(); }
    return '';
  }
  function addInfoWith(addInfo, name, responsible) {   // zet/verwijdert ons stukje in Additional Information
    const segs = String(addInfo || '').split(/\s*\|\s*/).map(s => s.trim()).filter(s => s && !IHC_MARK.test(s));
    if (name && low(name) !== low(responsible)) segs.push('IHC aanwezig: ' + name);
    return segs.join(' | ');
  }
  // fatField → IFS-veld (ihc wordt apart afgehandeld)
  const IFS_FIELDS = {
    recv: 'DATE_RECEIVED', appr: 'APPROVED_DB', loc: 'CF$_TEST_LOCATION',
    clsNR: 'CF$_CLASS_NOT_REQUIRED_DB', cusNR: 'CF$_CUSTOMER_NOT_REQUIRED_DB', inv: 'CF$_INVITATION_ID', note: 'NOTE'
  };
  const bool = v => v === 'Ja' ? 'TRUE' : v === 'Nee' ? 'FALSE' : null;
  function ifsValue(f, v) {
    if (f === 'recv') return ifsDate(v);
    if (f === 'appr' || f === 'clsNR' || f === 'cusNR') return bool(v);
    return flatTxt(v);
  }
  /* Past onze invulling (info = { sleutel: { veld: waarde } }) toe op het geplakte klembord.
   * Geeft { text, changed, total, unknown[], touched[] } — text bevat ALLEEN de gewijzigde blokken. */
  function applyToIfs(text, info) {
    const p = parseIfs(text);
    const out = []; const touched = []; let n = 0;
    for (const b of p.blocks) {
      const ov = info[b.key]; if (!ov) continue;
      const lines = b.lines.slice(); let diff = false; const what = [];
      const setv = (name, value) => {
        const fld = b.f[name]; if (!fld || value == null) return;
        if (fld.v === value) return;
        lines[fld.i] = lines[fld.i].replace(/=.*$/, () => '=' + value); diff = true; what.push(name);
      };
      for (const f of Object.keys(IFS_FIELDS)) {
        if (!(f in ov)) continue;
        let v = ov[f]; if (v === '-') v = '';
        setv(IFS_FIELDS[f], ifsValue(f, v));
      }
      if ('ihc' in ov) {
        const fld = b.f.ADD_INFO;
        if (fld) setv('ADD_INFO', flatTxt(addInfoWith(fld.v, ov.ihc === '-' ? '' : ov.ihc, (b.f.CF$_C_RESPONSIBLE_NAME || {}).v)));
      }
      if (diff) { out.push(b.start, ...lines); touched.push({ key: b.key, fields: what }); n++; }
    }
    const unknown = p.blocks.filter(b => !info[b.key]).map(b => b.key);
    const text2 = n ? p.head.concat(out).join(p.eol) + p.eol : '';
    return { text: text2, changed: n, total: p.blocks.length, unknown, touched };
  }

  global.FatData = { FILE, findSheet, tableFromSheet, buildCommitJSON, load, key, parseIfs, applyToIfs, ihcFromAddInfo, IFS_FIELDS };
})(typeof window !== 'undefined' ? window : globalThis);
