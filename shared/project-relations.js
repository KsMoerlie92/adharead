// ── Projectrelaties — hoofdprojecten en bijbehorende subprojecten ──────────
// Onder elk hoofdproject (Sub Project ID zoals in de Expediting-lijst) hangen
// de subprojectnummers die er relatie mee hebben. Gekoppeld betekent niet
// dubbel: scopes en subprojectcodes blijven behouden.
//
// Gebruik:
//  • Partnummer-vergelijker: Part No van een IHC-leverancier in het hoofdproject
//    opzoeken in de subprojecten (en omgekeerd).
//  • Admin: regels van deze subprojecten gaan mee in expediting-data.json.
//    Overige pagina's tonen nog steeds alleen YN/EN-projecten.
//
// zekerheid: 'bevestigd' | 'sterk' | 'kandidaat' (nog te bevestigen)
// Bron: interne inventarisatie (Sheet1 + Koppelingen), 23-11-2026.
(function (g) {
  g.PROJECT_RELATIONS = [
    { parent: 'EN10000', name: 'USACE / MCHD 6000', projects: [
      { id: 'DE10000',  desc: 'Dredge Installation MCHD 6.000 yrd3, USACE', zekerheid: 'bevestigd', bron: 'Mail; Koppelingen rij 7' } ] },
    { parent: 'EN10002', name: 'Customized Beagle 8 Mk2 / Dutra', projects: [
      { id: 'DE10002',  desc: 'Dredge Installation Customized Beagle 8 Mk2, Eastern Shipbuilding Group', zekerheid: 'sterk', bron: 'Via E10002; Koppelingen rij 8' },
      { id: 'SH000012', desc: 'Beagle 8 Mk2 Super Size - Dutra', zekerheid: 'sterk', bron: 'Indirect via E10002 / Dutra; Koppelingen rijen 9–10' } ] },
    { parent: 'YN1319', name: 'Beagle 8 Mk2 / Vishwa', projects: [
      { id: 'DE1319',   desc: 'Dredge Installation Beagle 8 Mk2, Vishwa Samudra Dredging', zekerheid: 'sterk', bron: 'Koppelingen rij 13' },
      { id: 'SH000013', desc: 'Beagle 8 - standard - Vishwa', zekerheid: 'sterk', bron: 'Koppelingen rijen 14–15' },
      { id: 'SE000406', desc: 'Planned Maintenance System Vishwa 01319', zekerheid: 'sterk', bron: 'Service voor hoofdproject; Koppelingen rij 16' } ] },
    { parent: 'YN1320', name: 'Cable layer / CS Genesis', projects: [
      { id: 'SH000008', desc: 'Tow Winch, HPU & Umbilical Winch OMS Vessel 1 - CS Genesis - YN1320', zekerheid: 'bevestigd', bron: 'YN1320 in omschrijving; Koppelingen rij 5' } ] },
    { parent: 'YN1321', name: 'Cable layer / CS Triumph', projects: [
      { id: 'SH000107', desc: 'Tow Winch, HPU (& Umbilical Winch) OMS Vessel 2 - CS Triumph - YN1321', zekerheid: 'bevestigd', bron: 'YN1321 in omschrijving; Koppelingen rij 6' } ] },
    { parent: 'YN1323', name: 'Easydredge 2700 XL', projects: [
      { id: 'DE1323',   desc: 'ED2700XL Congo scope', zekerheid: 'sterk', bron: 'Koppelingen rij 17' } ] },
    { parent: 'YN1324', name: 'Boskalis / TSHD 9000 m³ — schip 1', projects: [
      { id: 'SMIT1324', desc: 'Dredging installation - 9.000m3 TSHD - Boskalis (1st Vessel)', zekerheid: 'sterk', bron: 'Koppelingen rij 18' } ] },
    { parent: 'YN1325', name: 'Boskalis / TSHD 9000 m³ — schip 2', projects: [
      { id: 'SMIT1325', desc: 'Dredging installation - 9.000m3 TSHD - Boskalis (2nd Vessel)', zekerheid: 'kandidaat', bron: 'Nog te bevestigen: geen koppeling in Koppelingen; YN1325 ontbreekt in Sheet1' } ] },
    { parent: 'YN15061', name: 'QMM / QIT Madagascar', projects: [
      { id: 'DE000010', desc: 'Baggerinstallatie QMM', zekerheid: 'bevestigd', bron: 'QMM-mailketen; Koppelingen rijen 2–3' },
      { id: 'SH000085', desc: 'QMM Mining Dredger Rio Tinto 1100kW', zekerheid: 'bevestigd', bron: 'QMM-mailketen; Koppelingen rijen 2–4' } ] },
  ];
})(window);
