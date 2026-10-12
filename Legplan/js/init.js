// ── UI wiring: showTab, getLegplanRows ──────
// ── Tab switching ──────────────────────────────────────────────────────────
function showTab(tab, el) {
  document.querySelectorAll('.tab-chip').forEach(c => c.className = 'tab-chip');
  document.getElementById('legplan-wrap').classList.toggle('visible', tab === 'legplan');
  document.getElementById('cipl-wrap').classList.toggle('visible', tab === 'cipl');
  const trk = document.getElementById('track-wrap');
  if (trk) trk.classList.toggle('visible', tab === 'tracking');
  el.classList.add(tab === 'legplan' ? 'active-legplan' : tab === 'tracking' ? 'active-track' : 'active-cipl');
  if (tab === 'cipl') initCIPL();
  if (tab === 'tracking' && window.ShipTrack) ShipTrack.show();
}

// ── getLegplanRows: override from the shared JS (uses allRows) ─────────────
// Filter: BL# leeg · Kol Z niet "not received yet" (tenzij toggle aan)
function getLegplanRows() {
  const inclNRY = document.getElementById('toggle-nry')?.checked ?? false;
  return allRows.filter(r => {
    const alBlank = !r.colAL || r.colAL.trim() === '';
    if (!alBlank) return false;
    const zVal  = (r.colZ||'').trim().toLowerCase();
    const isNRY = zVal === 'not received yet' || zVal === '';
    if (isNRY && !inclNRY) return false;
    return true;
  });
}

// ── Drag-and-drop ─────────────────────────────────────────────────────────
document.querySelectorAll('.drop-zone').forEach(dz => {
  dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('dragover'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('dragover'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('dragover');
    const input = dz.querySelector('input[type=file]');
    if (e.dataTransfer.files.length) { input.files = e.dataTransfer.files; input.dispatchEvent(new Event('change')); }
  });
});

