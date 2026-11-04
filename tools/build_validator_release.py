"""Bouwt het releasepakket 'Itemlijst Validator' voor de repo Upgrade 2.0 uit de testomgeving (adharead).
adharead blijft ongewijzigd; de nette indeling wordt hier gemaakt:
- bestanden die de validator uit PO-Matcher/ leent, gaan naar shared/ resp. de eigen map
- alleen de bestanden die de validator nodig heeft
- geen site-gate (Entra-login komt via Azure)
- SheetJS en lettertypen lokaal (geen externe CDN's)
- menu met alleen de vrijgegeven pagina's, geen Admin-knop
"""
import os, re, shutil, sys
SRC = sys.argv[1]; OUT = sys.argv[2]; NM = sys.argv[3]   # adharead, uitvoermap, node_modules
if os.path.exists(OUT): shutil.rmtree(OUT)
def cp(rel, dst=None):
    d = os.path.join(OUT, dst or rel); os.makedirs(os.path.dirname(d), exist_ok=True); shutil.copy2(os.path.join(SRC, rel), d)
def edit(rel, pairs):
    p = os.path.join(OUT, rel); s = open(p, encoding='utf-8', newline='').read()
    for o, n in pairs:
        assert s.count(o) == 1, (rel, o[:80], s.count(o)); s = s.replace(o, n)
    open(p, 'w', encoding='utf-8', newline='').write(s)

V = 'Itemlijst-Validator/'
for f in ['index.html', 'validator.js', 'val-crossref.js', 'val-manual-match.js', 'val-lparts.js', 'val-export-mail.js', 'val-dgmail.js', 'val-mailgen.js']:
    cp(V + f)
for f in ['nav.css', 'shared.js', 'expediting-core.js', 'expediting-data.js', 'expediting-data.json', 'gn_codes_2026.js',
          'inspection-rules.js', 'inspection-rules.json', 'validatie-log.js', 'validatie-log.json', 'ifs-lparts-columns.js', 'lparts-numbering.js',
          'logo.png', 'logo-dark.png']:
    cp('shared/' + f)
# Leningen uit PO-Matcher/ → eigen plek in de release
cp('PO-Matcher/js/utils.js', 'shared/utils.js')
cp('PO-Matcher/js/state.js', 'shared/state.js')
for f in ['bg0.png', 'bg1.png', 'bg2.jpg']:
    cp('PO-Matcher/img/' + f, 'shared/img/' + f)
cp(V + 'styles.css')                                    # eigen stijl van de validator (achtergronden uit PO-Matcher/img)
edit(V + 'styles.css', [("url('../PO-Matcher/img/%s')" % f, "url('../shared/img/%s')" % f) for f in ['bg0.png', 'bg1.png', 'bg2.jpg']])

# ── SheetJS lokaal ──
os.makedirs(os.path.join(OUT, 'shared/vendor'), exist_ok=True)
shutil.copy2(os.path.join(NM, 'xlsx/dist/xlsx.full.min.js'), os.path.join(OUT, 'shared/vendor/xlsx-0.18.5.full.min.js'))
shutil.copy2(os.path.join(NM, 'xlsx/LICENSE'), os.path.join(OUT, 'shared/vendor/xlsx-LICENSE.txt'))

# ── lettertypen lokaal (latin + latin-ext, alleen woff2) ──
FD = os.path.join(OUT, 'shared/fonts'); os.makedirs(FD, exist_ok=True)
css = ['/* Lokale lettertypen (Fontsource, SIL Open Font License) — vervangt Google Fonts */']
for pkg, weights in [('barlow', [300, 400, 500, 600]), ('barlow-condensed', [300, 400, 600, 700, 800]), ('jetbrains-mono', [400, 700])]:
    base = os.path.join(NM, '@fontsource', pkg)
    for w in weights:
        for sub in ['latin-ext', 'latin']:
            t = open(os.path.join(base, f'{sub}-{w}.css'), encoding='utf-8').read()
            for m in re.findall(r"url\(\./files/([^)]+\.woff2)\)", t):
                shutil.copy2(os.path.join(base, 'files', m), os.path.join(FD, m))
            t = re.sub(r"url\(\./files/([^)]+\.woff2)\) format\('woff2'\), url\(\./files/[^)]+\.woff\) format\('woff'\)", r"url(\1) format('woff2')", t)
            css.append(t.strip())
    lic = os.path.join(base, 'LICENSE')
    if os.path.exists(lic): shutil.copy2(lic, os.path.join(FD, f'{pkg}-LICENSE.txt'))
open(os.path.join(FD, 'fonts.css'), 'w', encoding='utf-8').write('\n'.join(css) + '\n')

# ── pagina aanpassen ──
edit(V + 'index.html', [
    ('<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>', '<script src="../shared/vendor/xlsx-0.18.5.full.min.js"></script>'),
    ('<link href="https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@300;400;600;700;800&family=Barlow:wght@300;400;500;600&family=JetBrains+Mono:wght@400;700&display=swap" rel="stylesheet">',
     '<link rel="stylesheet" href="../shared/fonts/fonts.css">'),
    ('<script src="../shared/site-gate.js"></script>\n', ''),
    ('<script src="../PO-Matcher/js/utils.js"></script>', '<script src="../shared/utils.js"></script>'),
    ('<script src="../PO-Matcher/js/state.js"></script>', '<script src="../shared/state.js"></script>'),
    ("<!-- Load the exact same JS chain as PO-Matcher, from PO-Matcher's js/ folder -->", '<!-- Gedeelde basis uit shared/, pagina-eigen validator.js en styles.css -->'),
])
# link naar een tool die (nog) niet in deze release zit
edit(V + 'val-lparts.js', [(
    """plak deze rechtstreeks in IFS (of in de <a href="../IFS-Migration-Tool/index.html"
          target="_blank" style="color:var(--ihc-teal,#00B4D8)">IFS Migration Tool</a> om nog te bewerken).""",
    "plak deze rechtstreeks in IFS.")])

# ── menu: alleen vrijgegeven pagina's, geen Admin ──
p = os.path.join(OUT, 'shared/shared.js'); s = open(p, encoding='utf-8').read()
a = s.index('  const MENUS = ['); b = s.index('  ];', a) + len('  ];')
s = s[:a] + """  // Vrijgegeven pagina's in deze release (nieuwe pagina's per release toevoegen)
  const MENUS = [
    { label: 'Logistics', items: [
      { href: '../Itemlijst-Validator/', icon: '📋', title: 'Itemlijst Validator', sub: 'Valideer & corrigeer itemlijsten' },
    ]},
  ];""" + s[b:]
o = """      <a class="ihc-nav-btn" id="ihc-admin-btn" href="../Admin/" title="Admin — centrale expediting-upload" aria-label="Admin">⚙️</a>\n"""
assert s.count(o) == 1; s = s.replace(o, '')
open(p, 'w', encoding='utf-8').write(s)

# ── startpagina: direct naar de validator ──
open(os.path.join(OUT, 'index.html'), 'w', encoding='utf-8').write("""<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Royal IHC — Expedite 2.0</title>
<meta http-equiv="refresh" content="0; url=Itemlijst-Validator/">
<style>body{margin:0;background:#0A1628;color:#F0F4FA;font-family:Arial,sans-serif;display:grid;place-items:center;min-height:100vh}a{color:#00B4D8}</style>
</head>
<body>
<p>Doorsturen naar de <a href="Itemlijst-Validator/">Itemlijst Validator</a>…</p>
</body>
</html>
""")
open(os.path.join(OUT, 'RELEASE.md'), 'w', encoding='utf-8').write("""# Upgrade 2.0 — release 1: Itemlijst Validator

Samengesteld uit de testomgeving (adharead) met tools/build_validator_release.py.\nDe validator is hier zelfstandig: hij gebruikt alleen `shared/` en zijn eigen map.

- Geen toegangspoort (site-gate): de toegang loopt via Microsoft Entra ID in Azure.
- Geen externe CDN's: SheetJS 0.18.5 (`shared/vendor/`) en de lettertypen Barlow, Barlow Condensed
  en JetBrains Mono (`shared/fonts/`) staan lokaal.
- Menu toont alleen de vrijgegeven pagina('s); de Admin-pagina zit niet in deze release.

## Data verversen
`shared/expediting-data.json`, `shared/inspection-rules.json` en `shared/validatie-log.json` worden
in de testomgeving via de Admin-pagina gemaakt en hier vervangen.

## Externe verbindingen tijdens gebruik (geen CDN, wel functies)
- `api.frankfurter.dev` — EUR/USD-koers voor waarden in USD (zonder: melding, geen omrekening)
- `www.tariffnumber.com` — HS-code-controle en link naar de tariefpagina (zonder: geen extra controle)
""")
print('klaar')
