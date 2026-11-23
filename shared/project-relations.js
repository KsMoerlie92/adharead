// ── Projectrelaties — IHC Holland ─────────────────────────────────────────
// Onder het projectnummer van IHC Holland hangen alle projectnummers die met
// dat project te maken hebben. De Partnummer-Vergelijker zoekt een Part No
// van een IHC-leverancier op in de ándere projecten van dezelfde groep.
//
// Aanpassen: voeg een groep toe of zet projectnummers (Sub Project ID zoals in
// de Expediting-lijst) onder 'projects'. Het hoofdnummer zelf telt ook mee
// als het in de Expediting-lijst voorkomt.
(function (g) {
  g.PROJECT_RELATIONS = [
    // VOORBEELD — vervangen door de definitieve indeling
    { parent: 'VOORBEELD', name: 'Voorbeeld (te vervangen)', projects: ['YN1320', 'YN1321'] },
  ];
})(window);
