/* ============================================================
 *  shared/dg-rules.js  —  IHC Expedite 2.0
 *  Categorieën gevaarlijke stoffen (trefwoorden op de omschrijving).
 *  Eén bron voor de DG Overview en het Bezoekverslag (ADN-check).
 * ========================================================== */
window.DG_RULES = {
  "🔴 1. Brandstof & Brandbare vloeistoffen": {
    "risk": "high",
    "keywords": [
      "Biofuel",
      "B100 -",
      "Fuel storage",
      "fuel gas funnel",
      "Emulsifiable Saw",
      "Cutting Oil",
      "Sludge oil mixing tank",
      "SLUDGE BUOY",
      "Deep fat fryer"
    ],
    "action": "SDS/MSDS opvragen bij leverancier. IMDG Klasse 3 controle.",
    "regulations": "IMDG Code Klasse 3, MARPOL Annex I"
 ,
    "color": "#E8383D"
  },
  "🔴 2. Koelmiddelen (F-gas)": {
    "risk": "high",
    "keywords": [
      "Refrigerant R407",
      "Refrigerant 513",
      "R407f",
      "R513a",
      "Condensing unit",
      "Chilled water unit",
      "Blast Chiller",
      "Ice cube maker",
      "Refrigerator",
      "Freezer MBF",
      "Cold cupboard",
      "Cold counter",
      "Drinkwater cooler",
      "Combi cabinet",
      "Cabin Refrigerator"
    ],
    "action": "Koelmiddeltype bevestigen. F-gas registratie vereist. SDS opvragen.",
    "regulations": "EU F-gas Regulation 517/2014, CLP H280"
 ,
    "color": "#48CAE4"
  },
  "🔴 3. Batterijen & Energieopslag": {
    "risk": "high",
    "keywords": [
      "ESS battery",
      "UPS (3kVA)",
      "UPS 1000VA",
      "UPS 1500",
      "Uninterruptable Power Supply",
      "Uninterruptible Power Supply",
      "Li-Ion",
      "EPIRB radio beacon",
      "SART rescue",
      "ATEX UHF",
      "DECT Handset",
      "Motorola R7 NKP",
      "GMDSS emergency power",
      "UPS kits",
      "UPS kit"
    ],
    "action": "Batterijtype en State of Charge (SoC) bevestigen. IMDG Klasse 9 UN 3480/3481.",
    "regulations": "IMDG Code Klasse 9, UN 3480/3481/3536"
 ,
    "color": "#FFEA00"
  },
  "🔴 4. Brandblusmiddelen & Gassen": {
    "risk": "high",
    "keywords": [
      "Fire extinguishing medium cylinder",
      "CO2 Manifold",
      "CO2 cylinders",
      "Fire extinguishing cabinet",
      "Gen. Service Air Receiver",
      "Watermist pump unit"
    ],
    "action": "SDS opvragen. Drukvatveiligheid controleren. CO2 = Klasse 2.2.",
    "regulations": "IMDG Code Klasse 2.2, UN 1013, PED Directive"
 ,
    "color": "#FB8500"
  },
  "🟠 5. SCR/Ureum (uitlaatgas)": {
    "risk": "medium",
    "keywords": [
      "SCR reactor",
      "Urea dosing unit",
      "Urea supply unit",
      "Urea supply pump",
      "Nozzle injector unit",
      "Injection pipe",
      "Soot blowing panel",
      "SCR Control Unit"
    ],
    "action": "Katalysator bevat mogelijk V2O5 (giftig). Ureum = ammoniakvorming bij verhitting.",
    "regulations": "CLP: V2O5 = Carc. 2 / Repr. 2, AFS-40 urea handling"
 ,
    "color": "#A68A3C"
  },
  "🟠 6. Hydraulische systemen": {
    "risk": "medium",
    "keywords": [
      "Hydraulic power unit",
      "Hangar and Side Doors HPU",
      "Hydraulic hatches Forward",
      "Hydraulic Cylinder",
      "Hydraulic hand pump",
      "Hydraulic pipe bending",
      "HiPAP Gate Valve HPU"
    ],
    "action": "Hydraulische olie type bevestigen. Milieugevaarlijk. Lekpreventie controleren.",
    "regulations": "CLP H304/H411, MARPOL Annex I"
 ,
    "color": "#3A86FF"
  },
  "🟠 7. Transformatoren": {
    "risk": "medium",
    "keywords": [
      "440V/110V jointing room transformer",
      "440V/110V cable hanger transformer",
      "Plough Surface Transformer",
      "Jetting Transformer",
      "AVR ABB Unitrol"
    ],
    "action": "Controleer op PCB’s (IHM Table A). Isolatieolie/hars type verifiëren.",
    "regulations": "IHM Appendix 1 (PCBs), EU SRR, Stockholm Convention"
 ,
    "color": "#2A4494"
  },
  "🟠 8. Motoren & Verbranding": {
    "risk": "medium",
    "keywords": [
      "Generator set - Engine - Himsen",
      "Exhaust gas silencer",
      "Incinerator - Atlas",
      "Spark Arrester",
      "prelubricating oil pump"
    ],
    "action": "Smeerolie, koelvloeistof, RCF isolatie controleren. IHM Table B check.",
    "regulations": "IHM Appendix 2 (RCF), MARPOL Annex VI"
 ,
    "color": "#BC6C25"
  },
  "🟡 9. Lassen & Werkplaats": {
    "risk": "low",
    "keywords": [
      "Welding machine",
      "MIG Welding torch",
      "Grinding machine",
      "Metal saw blade",
      "Fuel oil injection nozzle test",
      "Welding screen"
    ],
    "action": "Lasdampen bevatten zware metalen. Persoonlijke beschermingsmiddelen vereist.",
    "regulations": "EU OEL Directive, CLP bijlage VI"
 ,
    "color": "#E8B923"
  },
  "🟡 10. Rubber & Isolatie (IHM)": {
    "risk": "low",
    "keywords": [
      "Rubber Design",
      "Insulation",
      "Steel Expansion Joint",
      "Thermal Insulation"
    ],
    "action": "IHM Material Declaration opvragen. Check op asbest, RCF, HBCDD, cadmium, zware metalen.",
    "regulations": "IHM Table A (asbest) & Table B (cadmium, HBCDD, RCF)"
 ,
    "color": "#8D6E63"
  },
  "🟡 11. Overig gevaarlijk": {
    "risk": "low",
    "keywords": [
      "Searchlight xenon",
      "Gyro master compass",
      "CJC Marine Diesel Purifier",
      "Azimuth thruster",
      "Liquid immersion heater",
      "Crimpin tool"
    ],
    "action": "Individuele beoordeling nodig. Xenon = gas onder druk, Gyro = mogelijk kwik.",
    "regulations": "Diverse – per item beoordelen"
 ,
    "color": "#6C757D"
  },
  "\ud83d\udd35 12. Computer Based Systems (cyberresilience)": {
    "risk": "medium",
    "keywords": [
      "ECDIS", "RADAR", "AIS transponder", "Gyro compass", "GNSS receiver",
      "Dynamic Positioning", "DPS", "Integrated Automation System", "IAS",
      "Engine Control System", "Remote Propulsion control", "Boiler Management system",
      "Ballast water management system", "VDR", "SSAS", "Cargo Control and monitoring",
      "Loading computer", "Inert Gas System", "NAVTEX", "INMARSAT"
    ],
    "action": "Bepaal of dit een Computer Based System (CBS) is onder IACS UR E26/E27. Toets tegen de laag-risico-uitzonderingscriteria (geïsoleerd/geen netwerk, geen bereikbare fysieke poorten, fysiek afgeschermde locatie, niet-geïntegreerd, geen categorie III-veiligheidsfunctie) — bij twijfel is een volledige cyberresilience-beoordeling vereist. Vul het volledige beoordelingsformulier in via de link hieronder.",
    "regulations": "IACS UR E26/E27, IACS Rec. 166 (Cyber Resilience), IACS Rec. 171 (mei 2022)",
    "formLink": "../E27-Assessment/index.html",
    "color": "#B24BF3"
  },
  "🔴 13. Explosieven & Pyrotechniek": {
    "risk": "high",
    "keywords": [
      "Distress flare", "Signal flare", "Hand flare", "Smoke signal",
      "Parachute flare", "Rocket line throwing", "Pyrotechnic",
      "Cartridge actuated", "Explosive bolt", "Igniter cartridge"
    ],
    "action": "SDS opvragen. IMDG Klasse 1 subklasse en compatibiliteitsgroep bevestigen. Reddingsmiddelen (vuurpijlen/fakkels) vallen vaak onder deze klasse.",
    "regulations": "IMDG Code Klasse 1, SOLAS Hoofdstuk III (reddingsmiddelen)",
    "color": "#FF5A1F"
  },
  "🔴 14. Bijtende stoffen (Corrosieven)": {
    "risk": "high",
    "keywords": [
      "Battery acid", "Sulphuric acid", "Sodium hydroxide", "Caustic soda",
      "Hydrochloric acid", "Boiler descaler", "Corrosion inhibitor concentrate",
      "Alkaline cleaner concentrate"
    ],
    "action": "SDS opvragen. pH en verpakkingsgroep (I/II/III) bevestigen. Corrosief voor huid/ogen — PBM vereist.",
    "regulations": "IMDG Code Klasse 8, CLP Skin Corr. 1 / Eye Dam. 1",
    "color": "#14B8A6"
  },
  "🔴 15. Toxische & Milieugevaarlijke stoffen": {
    "risk": "high",
    "keywords": [
      "Pesticide", "Biocide", "Fumigant", "Antifouling paint", "TBT",
      "Marine pollutant", "Methanol", "Toxic reagent", "Environmentally hazardous"
    ],
    "action": "SDS opvragen. Marine Pollutant-markering (IMDG) controleren. Blootstellingsrisico beoordelen.",
    "regulations": "IMDG Code Klasse 6.1 / Marine Pollutant, MARPOL Annex III",
    "color": "#2ECC71"
  },
  "🔴 16. Radioactieve bronnen (meetinstrumenten)": {
    "risk": "high",
    "keywords": [
      "Nucleonic level gauge", "Radioactive source", "Density gauge",
      "Gamma gauge", "Level detector caesium", "Americium source",
      "Check source", "Radiometric gauge"
    ],
    "action": "Stralingsbeschermingsdeskundige (SBD) inschakelen. Vergunning/kennisgeving controleren — nucleonic density-/niveaumeters komen veel voor op baggerinstallaties (bv. slurriedichtheidsmeting).",
    "regulations": "IMDG Code Klasse 7, Besluit basisveiligheidsnormen stralingsbescherming (ANVS)",
    "color": "#FF3EA5"
  },
  "🟠 17. Oxiderende stoffen & Peroxiden": {
    "risk": "medium",
    "keywords": [
      "Sodium hypochlorite", "Calcium hypochlorite", "Hydrogen peroxide",
      "Oxidizing agent", "Chlorine dosing", "Bleach concentrate"
    ],
    "action": "SDS opvragen. Niet samen met brandbare stoffen opslaan (IMDG-segregatietabel).",
    "regulations": "IMDG Code Klasse 5.1 / 5.2",
    "color": "#CDDC39"
  },
  "🟠 18. Verven, coatings & Isocyanaten": {
    "risk": "medium",
    "keywords": [
      "Two-component paint", "Isocyanate", "MDI", "TDI", "Epoxy coating",
      "Polyurethane coating", "Primer concentrate", "Paint thinner"
    ],
    "action": "SDS + ademhalingsbescherming controleren. Isocyanaten zijn sensibiliserend (astma-risico) — ventilatie-eisen naleven.",
    "regulations": "CLP Resp. Sens. 1, REACH Bijlage XVII entry 74 (diisocyanaten)",
    "color": "#CE74D6"
  }
};
