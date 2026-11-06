"""Bouwt een releasepakket voor de repo Upgrade 2.0 uit de testomgeving (adharead).

    python3 tools/build_release.py <adharead-map> <uitvoermap> <node_modules-map>

adharead blijft ongewijzigd; de nette indeling wordt hier gemaakt:
- alleen de pagina's uit RELEASED (plus wat ze uit shared/ nodig hebben)
- wat een pagina uit PO-Matcher/ leent, krijgt een eigen plek (shared/ of de eigen map)
- geen site-gate (toegang via Microsoft Entra ID in Azure)
- SheetJS en lettertypen lokaal (geen externe CDN's)
- startpagina en menu tonen alleen de vrijgegeven pagina's, geen Admin
Een nieuwe pagina vrijgeven = een blok toevoegen aan RELEASED.
"""
import os, re, shutil, sys

SRC, OUT, NM = sys.argv[1], sys.argv[2], sys.argv[3]
RELEASE = 'Release 2 — Itemlijst Validator + Itemlijst Template Generator'

# ── vrijgegeven pagina's ─────────────────────────────────────────────────────
RELEASED = {
    'Itemlijst-Validator': {
        'files': ['index.html', 'validator.js', 'styles.css', 'val-crossref.js', 'val-manual-match.js', 'val-lparts.js',
                  'val-export-mail.js', 'val-dgmail.js', 'val-mailgen.js'],
        'styles': 'own',                      # eigen styles.css (achtergronden uit PO-Matcher/img)
        'menu': {'icon': '📋', 'title': 'Itemlijst Validator', 'sub': 'Valideer & corrigeer itemlijsten'},
        'patch': {   # link naar een tool die (nog) niet in de release zit
            'val-lparts.js': [("""plak deze rechtstreeks in IFS (of in de <a href="../IFS-Migration-Tool/index.html"
          target="_blank" style="color:var(--ihc-teal,#00B4D8)">IFS Migration Tool</a> om nog te bewerken).""",
                               "plak deze rechtstreeks in IFS.")],
        },
    },
    'Itemlijst-Template-Generator': {
        'files': ['index.html', 'template-base.xlsx', 'template-generator.js', 'template-xlsx-writer.js',
                  'template-generator-ui.js', 'template-mail.js'],
        'styles': 'po-matcher',               # gebruikt in adharead PO-Matcher/styles.css → eigen kopie
        'menu': {'icon': '📝', 'title': 'Itemlijst Template Generator', 'sub': 'Template per PO opstellen & naar leverancier mailen'},
    },
}
SHARED = ['nav.css', 'shared.js', 'expediting-core.js', 'expediting-data.js', 'expediting-data.json', 'gn_codes_2026.js',
          'inspection-rules.js', 'inspection-rules.json', 'validatie-log.js', 'validatie-log.json',
          'ifs-lparts-columns.js', 'lparts-numbering.js', 'logo.png', 'logo-dark.png']
FONTS = [('barlow', [300, 400, 500, 600]), ('barlow-condensed', [300, 400, 500, 600, 700, 800]), ('jetbrains-mono', [400, 500, 700])]

if os.path.exists(OUT): shutil.rmtree(OUT)
def cp(rel, dst=None):
    d = os.path.join(OUT, dst or rel); os.makedirs(os.path.dirname(d), exist_ok=True); shutil.copy2(os.path.join(SRC, rel), d)
def rd(rel): return open(os.path.join(OUT, rel), encoding='utf-8', newline='').read()
def wr(rel, s): open(os.path.join(OUT, rel), 'w', encoding='utf-8', newline='').write(s)
def edit(rel, pairs, optional=False):
    s = rd(rel)
    for o, n in pairs:
        if optional and s.count(o) == 0: continue
        assert s.count(o) == 1, (rel, o[:80], s.count(o)); s = s.replace(o, n)
    wr(rel, s)

# ── gedeelde basis ──
for f in SHARED: cp('shared/' + f)
cp('PO-Matcher/js/utils.js', 'shared/utils.js')
cp('PO-Matcher/js/state.js', 'shared/state.js')
for f in ['bg0.png', 'bg1.png', 'bg2.jpg']: cp('PO-Matcher/img/' + f, 'shared/img/' + f)
os.makedirs(os.path.join(OUT, 'shared/vendor'), exist_ok=True)
shutil.copy2(os.path.join(NM, 'xlsx/dist/xlsx.full.min.js'), os.path.join(OUT, 'shared/vendor/xlsx-0.18.5.full.min.js'))
shutil.copy2(os.path.join(NM, 'xlsx/LICENSE'), os.path.join(OUT, 'shared/vendor/xlsx-LICENSE.txt'))

# ── lettertypen lokaal (latin + latin-ext, alleen woff2) ──
FD = os.path.join(OUT, 'shared/fonts'); os.makedirs(FD, exist_ok=True)
css = ['/* Lokale lettertypen (Fontsource, SIL Open Font License) — vervangt Google Fonts */']
for pkg, weights in FONTS:
    base = os.path.join(NM, '@fontsource', pkg)
    for w in weights:
        for sub in ['latin-ext', 'latin']:
            t = open(os.path.join(base, f'{sub}-{w}.css'), encoding='utf-8').read()
            for m in re.findall(r"url\(\./files/([^)]+\.woff2)\)", t): shutil.copy2(os.path.join(base, 'files', m), os.path.join(FD, m))
            t = re.sub(r"url\(\./files/([^)]+\.woff2)\) format\('woff2'\), url\(\./files/[^)]+\.woff\) format\('woff'\)", r"url(\1) format('woff2')", t)
            css.append(t.strip())
    if os.path.exists(os.path.join(base, 'LICENSE')): shutil.copy2(os.path.join(base, 'LICENSE'), os.path.join(FD, f'{pkg}-LICENSE.txt'))
open(os.path.join(FD, 'fonts.css'), 'w', encoding='utf-8').write('\n'.join(css) + '\n')

XLSX_CDN = '<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>'
GFONTS = re.compile(r'<link href="https://fonts\.googleapis\.com/css2\?[^"]*" rel="stylesheet">')
IMG = [("url('img/%s')" % f, "url('../shared/img/%s')" % f) for f in ['bg0.png', 'bg1.png', 'bg2.jpg']]
IMG_PM = [("url('../PO-Matcher/img/%s')" % f, "url('../shared/img/%s')" % f) for f in ['bg0.png', 'bg1.png', 'bg2.jpg']]

# ── pagina's ──
for page, cfg in RELEASED.items():
    for f in cfg['files']: cp(f'{page}/{f}')
    if cfg['styles'] == 'po-matcher':
        cp('PO-Matcher/styles.css', f'{page}/styles.css'); edit(f'{page}/styles.css', IMG)
    else:
        edit(f'{page}/styles.css', IMG_PM)
    s = rd(f'{page}/index.html')
    assert s.count(XLSX_CDN) == 1 and len(GFONTS.findall(s)) == 1, page
    s = s.replace(XLSX_CDN, '<script src="../shared/vendor/xlsx-0.18.5.full.min.js"></script>')
    s = GFONTS.sub('<link rel="stylesheet" href="../shared/fonts/fonts.css">', s)
    s = re.sub(r'<link rel="preconnect" href="https://fonts\.(googleapis|gstatic)\.com"[^>]*>\n?', '', s)
    s = s.replace('<script src="../shared/site-gate.js"></script>\n', '')
    s = s.replace('<link rel="stylesheet" href="../PO-Matcher/styles.css">', '<link rel="stylesheet" href="styles.css">')
    s = s.replace('<script src="../PO-Matcher/js/utils.js"></script>', '<script src="../shared/utils.js"></script>')
    s = s.replace('<script src="../PO-Matcher/js/state.js"></script>', '<script src="../shared/state.js"></script>')
    s = s.replace("<!-- Load the exact same JS chain as PO-Matcher, from PO-Matcher's js/ folder -->", '<!-- Gedeelde basis uit shared/, pagina-eigen scripts en styles.css -->')
    assert not re.search(r'(src|href)="\.\./PO-Matcher/', s) and 'site-gate.js' not in s and 'googleapis' not in s and 'cdnjs' not in s, page
    wr(f'{page}/index.html', s)
    for f, pairs in cfg.get('patch', {}).items(): edit(f'{page}/{f}', pairs)

# ── menu (subpagina's): alleen vrijgegeven pagina's, geen Admin ──
s = rd('shared/shared.js')
a = s.index('  const MENUS = ['); b = s.index('  ];', a) + len('  ];')
items = '\n'.join("      { href: '../%s/', icon: '%s', title: '%s', sub: '%s' }," % (p, c['menu']['icon'], c['menu']['title'], c['menu']['sub']) for p, c in RELEASED.items())
s = s[:a] + "  // Vrijgegeven pagina's in deze release\n  const MENUS = [\n    { label: 'Logistics', items: [\n" + items + "\n    ]},\n  ];" + s[b:]
o = """      <a class="ihc-nav-btn" id="ihc-admin-btn" href="../Admin/" title="Admin — centrale expediting-upload" aria-label="Admin">⚙️</a>\n"""
assert s.count(o) == 1; s = s.replace(o, ''); wr('shared/shared.js', s)

# ── startpagina: de suite-startpagina van adharead, ingekort tot de vrijgegeven pagina's ──
cp('index.html'); cp('img/hero.png')
s = rd('index.html'); nl = '\r\n' if '\r\n' in s else '\n'; s = s.replace('\r\n', '\n')
def keep(href): return href.rstrip('/') in RELEASED
s = re.sub(r'\n        <a class="dd-link" href="([^"]+)">.*?</a>', lambda m: m.group(0) if keep(m.group(1)) else '', s, flags=re.S)
s = re.sub(r'\n      <a class="card" href="([^"]+)">.*?\n      </a>', lambda m: m.group(0) if keep(m.group(1)) else '', s, flags=re.S)
# lege dropdowns/secties weglaten
s = re.sub(r'\n    <!-- [A-Z]+ -->\n    <div class="nav-item">\n      <button[^\n]*\n      <div class="dropdown">\n      </div>\n    </div>', '', s)
s = re.sub(r'\n  <!-- [A-Z ]+SECTION -->\n  <section[^>]*>\n    <div class="sec-head">.*?</div>\n    <div class="grid">\n    </div>\n  </section>\n', '\n', s, flags=re.S)
s = re.sub(r'\n    <a class="icon-btn" href="Admin/"[^\n]*', '', s)
s = s.replace('<script src="shared/site-gate.js"></script>\n', '')
s = re.sub(r'<link rel="preconnect" href="https://fonts\.(googleapis|gstatic)\.com"[^>]*>\n', '', s)
s = re.sub(r'<link href="https://fonts\.googleapis\.com/css2\?[^"]*" rel="stylesheet">', '<link rel="stylesheet" href="shared/fonts/fonts.css">', s)
for p in RELEASED:
    m = re.search(r'<a class="card" href="%s/">.*?<img src="(img/[^"]+)"' % re.escape(p), s, re.S)
    if m and os.path.exists(os.path.join(SRC, m.group(1))): cp(m.group(1))
assert 'googleapis' not in s and 'site-gate' not in s and 'Admin/' not in s
s = re.sub(r'<p>Kies hieronder een tool[^<]*</p>', '<p>Kies hieronder een tool, of gebruik de menubalk: stel per PO een itemlijst-template op en valideer de ingevulde itemlijsten van leveranciers.</p>', s)
wr('index.html', s.replace('\n', nl))
# de eerste sectie heet in de release "01"
wr('index.html', re.sub(r'<span class="sec-eyebrow">0\d</span>', '<span class="sec-eyebrow">01</span>', rd('index.html'), count=1))

open(os.path.join(OUT, 'RELEASE.md'), 'w', encoding='utf-8').write(f"""# Upgrade 2.0 — {RELEASE}

Samengesteld uit de testomgeving (adharead) met `tools/build_release.py`.
Elke pagina gebruikt alleen `shared/` en zijn eigen map.

## Pagina's in deze release
""" + ''.join(f"- `{p}/` — {c['menu']['title']}: {c['menu']['sub']}\n" for p, c in RELEASED.items()) + """
## Afspraken
- Geen toegangspoort (site-gate): de toegang loopt via Microsoft Entra ID in Azure.
- Geen externe CDN's: SheetJS 0.18.5 (`shared/vendor/`) en de lettertypen Barlow, Barlow Condensed
  en JetBrains Mono (`shared/fonts/`) staan lokaal.
- Startpagina en menu tonen alleen de vrijgegeven pagina's; de Admin-pagina zit niet in de release.

## Data verversen
`shared/expediting-data.json`, `shared/inspection-rules.json` en `shared/validatie-log.json` worden
in de testomgeving via de Admin-pagina gemaakt en hier vervangen.

## Externe verbindingen tijdens gebruik (geen CDN, wel functies)
- `api.frankfurter.dev` — EUR/USD-koers voor waarden in USD (validator; zonder: melding, geen omrekening)
- `www.tariffnumber.com` — HS-code-controle (validator; zonder: geen extra controle)
- `mailto:` — de mailfuncties openen het mailprogramma van de gebruiker
""")
print('klaar:', ', '.join(RELEASED))
