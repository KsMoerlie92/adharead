# Transportaanvraag Shipping – GitHub Pages formulier

Bestanden: `index.html`, `style.css`, `script.js` (geen externe afhankelijkheden, geen backend).

## Publiceren
1. Maak een repository en upload de 3 bestanden in de root.
2. Settings → Pages → Source: *Deploy from a branch* → `main` / `(root)`.
3. Open de Pages-URL.

## Aanpassen (alles in `script.js`)
- `CONFIG`: mailadressen en drempels (zwaar, overmaat, koerier).
- `SECTIONS`: alle vragen. Per veld: `req` (verplicht), `vis` (zichtbaar), `validate`.
- `derive()`: de schifting (binnenlands / intra-EU / export / import, UK, RO/HU).
