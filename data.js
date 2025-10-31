// DASHBOARD VUOTA - Nessun dato precaricato
// Gli utenti devono caricare i propri dati tramite file Excel

// Dati di esempio commentati - non vengono caricati
const REAL_SITES_DATA_EXAMPLE = [
  {
    "SITO": "ROMA Via Barletta 29",
    "Metallo": "bronze",
    "DOIT": "ROMA",
    "Località": "Roma",
    "lat": 41.9028,
    "long": 12.4964,
    "KPI_CENTRALINA": "SI",
    "KPI_STAZIONE": "NO",
    "KPI_FV": "SI",
    "KPI_ILLUMINAZIONE_BANCHINE": "SI",
    "KPI_LFM": "NO",
    "utenze": [
      {"Nome_Utenza": "Illuminazione Esterna", "Tipologia": "Elettrica", "POD": "IT001E12345678", "Competenza": "ROMA", "Consumo_kWh": 45000, "Uso_Energetico": "Illuminazione"},
      {"Nome_Utenza": "Climatizzazione Uffici", "Tipologia": "Elettrica", "POD": "IT001E12345679", "Competenza": "ROMA", "Consumo_kWh": 62000, "Uso_Energetico": "Climatizzazione"},
      {"Nome_Utenza": "Apparati Tecnologici", "Tipologia": "Elettrica", "POD": "IT001E12345680", "Competenza": "ROMA", "Consumo_kWh": 180000, "Uso_Energetico": "Apparati e Sistemi tecnologici"}
    ],
    "pod_fatturati": [
      {"Nome": "POD Principale", "Codice": "IT001E12345678", "2020": 280000, "2021": 285000, "2022": 287000, "2023": 290000}
    ],
    "pdr_fatturati": [
      {"Nome": "PDR Riscaldamento", "Codice": "00123456789012", "2020": 12000, "2021": 11500, "2022": 11800, "2023": 12200}
    ],
    "teleriscaldamento": null,
    "consumi_uso": {
      "Illuminazione": 45000,
      "Climatizzazione": 62000,
      "Apparati e Sistemi tecnologici": 180000,
      "FM utenze": 0
    }
  },
  {
    "SITO": "MILANO Piazzale Corvetto",
    "Metallo": "gold",
    "DOIT": "MILANO",
    "Località": "Milano",
    "lat": 45.4373,
    "long": 9.2351,
    "KPI_CENTRALINA": "SI",
    "KPI_STAZIONE": "SI",
    "KPI_FV": "SI",
    "KPI_ILLUMINAZIONE_BANCHINE": "SI",
    "KPI_LFM": "SI",
    "utenze": [
      {"Nome_Utenza": "Illuminazione", "Tipologia": "Elettrica", "POD": "IT001E98765432", "Competenza": "MILANO", "Consumo_kWh": 85000, "Uso_Energetico": "Illuminazione"},
      {"Nome_Utenza": "Climatizzazione", "Tipologia": "Elettrica", "POD": "IT001E98765433", "Competenza": "MILANO", "Consumo_kWh": 145000, "Uso_Energetico": "Climatizzazione"},
      {"Nome_Utenza": "Apparati Tecnologici", "Tipologia": "Elettrica", "POD": "IT001E98765434", "Competenza": "MILANO", "Consumo_kWh": 320000, "Uso_Energetico": "Apparati e Sistemi tecnologici"},
      {"Nome_Utenza": "FM", "Tipologia": "Elettrica", "POD": "IT001E98765435", "Competenza": "MILANO", "Consumo_kWh": 25000, "Uso_Energetico": "FM utenze"}
    ],
    "pod_fatturati": [
      {"Nome": "POD 1", "Codice": "IT001E98765432", "2020": 520000, "2021": 540000, "2022": 575000, "2023": 585000},
      {"Nome": "POD 2", "Codice": "IT001E98765436", "2020": 180000, "2021": 185000, "2022": 190000, "2023": 195000}
    ],
    "pdr_fatturati": [
      {"Nome": "PDR Principale", "Codice": "98765432101234", "2020": 28000, "2021": 27500, "2022": 28200, "2023": 29000}
    ],
    "teleriscaldamento": {
      "Codice": "TLR-MI-001",
      "Nome_PDR": "Teleriscaldamento Milano Corvetto",
      "kWht": 156122
    },
    "consumi_uso": {
      "Illuminazione": 85000,
      "Climatizzazione": 145000,
      "Apparati e Sistemi tecnologici": 320000,
      "FM utenze": 25000
    }
  },
  {
    "SITO": "TORINO Corso Giulio Cesare",
    "Metallo": "silver",
    "DOIT": "TORINO",
    "Località": "Torino",
    "lat": 45.0865,
    "long": 7.6688,
    "KPI_CENTRALINA": "SI",
    "KPI_STAZIONE": "NO",
    "KPI_FV": "NO",
    "KPI_ILLUMINAZIONE_BANCHINE": "SI",
    "KPI_LFM": "NO",
    "utenze": [
      {"Nome_Utenza": "Illuminazione", "Tipologia": "Elettrica", "POD": "IT001E55512345", "Competenza": "TORINO", "Consumo_kWh": 58000, "Uso_Energetico": "Illuminazione"},
      {"Nome_Utenza": "Climatizzazione", "Tipologia": "Elettrica", "POD": "IT001E55512346", "Competenza": "TORINO", "Consumo_kWh": 92000, "Uso_Energetico": "Climatizzazione"},
      {"Nome_Utenza": "Apparati", "Tipologia": "Elettrica", "POD": "IT001E55512347", "Competenza": "TORINO", "Consumo_kWh": 210000, "Uso_Energetico": "Apparati e Sistemi tecnologici"}
    ],
    "pod_fatturati": [
      {"Nome": "POD Unico", "Codice": "IT001E55512345", "2020": 340000, "2021": 350000, "2022": 360000, "2023": 365000}
    ],
    "pdr_fatturati": [
      {"Nome": "PDR Gas", "Codice": "55512345678901", "2020": 18000, "2021": 17800, "2022": 18200, "2023": 18500}
    ],
    "teleriscaldamento": null,
    "consumi_uso": {
      "Illuminazione": 58000,
      "Climatizzazione": 92000,
      "Apparati e Sistemi tecnologici": 210000,
      "FM utenze": 0
    }
  }
];

// Function to generate additional sites (reach 125 total)
function generateAdditionalSites() {
  const cities = [
    {name: "ROMA", lat: 41.9, lng: 12.5, count: 19},
    {name: "MILANO", lat: 45.4, lng: 9.2, count: 8},
    {name: "TORINO", lat: 45.0, lng: 7.7, count: 9},
    {name: "BOLOGNA", lat: 44.5, lng: 11.3, count: 12},
    {name: "FIRENZE", lat: 43.8, lng: 11.2, count: 10},
    {name: "GENOVA", lat: 44.4, lng: 8.9, count: 10},
    {name: "NAPOLI", lat: 40.8, lng: 14.2, count: 8},
    {name: "BARI", lat: 41.1, lng: 16.9, count: 9},
    {name: "PALERMO", lat: 38.1, lng: 13.4, count: 6},
    {name: "ANCONA", lat: 43.6, lng: 13.5, count: 10},
    {name: "VENEZIA", lat: 45.4, lng: 12.3, count: 7},
    {name: "VERONA", lat: 45.4, lng: 10.9, count: 5},
    {name: "PADOVA", lat: 45.4, lng: 11.9, count: 4},
    {name: "BRESCIA", lat: 45.5, lng: 10.2, count: 3},
    {name: "TRIESTE", lat: 45.6, lng: 13.8, count: 3}
  ];
  
  const categories = ["bronze", "silver", "gold", "platinum", "officina", "ust"];
  const categoryWeights = [13, 65, 37, 5, 4, 1]; // Target distribution
  
  const additionalSites = [];
  let siteCounter = 4; // Start from 4 since we have 3 base sites
  
  cities.forEach(city => {
    for (let i = 0; i < city.count; i++) {
      // Assign category based on weights
      let catIndex = 1; // Default to silver
      const rand = Math.random() * 125;
      if (rand < 13) catIndex = 0; // bronze
      else if (rand < 78) catIndex = 1; // silver (13 + 65)
      else if (rand < 115) catIndex = 2; // gold (78 + 37)
      else if (rand < 120) catIndex = 3; // platinum (115 + 5)
      else if (rand < 124) catIndex = 4; // officina (120 + 4)
      else catIndex = 5; // ust
      
      const category = categories[catIndex];
      const variation = 0.001 * (Math.random() - 0.5);
      
      // Generate consumption based on category
      const baseConsumption = {
        bronze: {electric: 180000, gas: 10000},
        silver: {electric: 250000, gas: 14000},
        gold: {electric: 380000, gas: 20000},
        platinum: {electric: 520000, gas: 28000},
        officina: {electric: 320000, gas: 16000},
        ust: {electric: 420000, gas: 22000}
      };
      
      const base = baseConsumption[category];
      const electricTotal = Math.round(base.electric * (0.85 + Math.random() * 0.3));
      const gasTotal = Math.round(base.gas * (0.85 + Math.random() * 0.3));
      
      // Distribute consumption across uso energetico
      const illuminazione = Math.round(electricTotal * 0.2);
      const climatizzazione = Math.round(electricTotal * 0.25);
      const apparati = Math.round(electricTotal * 0.50);
      const fm = electricTotal - illuminazione - climatizzazione - apparati;
      
      const site = {
        SITO: `${city.name} Via ${siteCounter}`,
        Metallo: category,
        DOIT: city.name,
        Località: city.name,
        lat: city.lat + variation,
        long: city.lng + variation,
        KPI_CENTRALINA: Math.random() > 0.3 ? "SI" : "NO",
        KPI_STAZIONE: Math.random() > 0.6 ? "SI" : "NO",
        KPI_FV: Math.random() > 0.5 ? "SI" : "NO",
        KPI_ILLUMINAZIONE_BANCHINE: Math.random() > 0.4 ? "SI" : "NO",
        KPI_LFM: Math.random() > 0.7 ? "SI" : "NO",
        utenze: [
          {Nome_Utenza: "Illuminazione", Tipologia: "Elettrica", POD: `IT001E${siteCounter}0001`, Competenza: city.name, Consumo_kWh: illuminazione, Uso_Energetico: "Illuminazione"},
          {Nome_Utenza: "Climatizzazione", Tipologia: "Elettrica", POD: `IT001E${siteCounter}0002`, Competenza: city.name, Consumo_kWh: climatizzazione, Uso_Energetico: "Climatizzazione"},
          {Nome_Utenza: "Apparati Tecnologici", Tipologia: "Elettrica", POD: `IT001E${siteCounter}0003`, Competenza: city.name, Consumo_kWh: apparati, Uso_Energetico: "Apparati e Sistemi tecnologici"},
          {Nome_Utenza: "FM", Tipologia: "Elettrica", POD: `IT001E${siteCounter}0004`, Competenza: city.name, Consumo_kWh: fm, Uso_Energetico: "FM utenze"}
        ],
        pod_fatturati: [
          {Nome: "POD Principale", Codice: `IT001E${siteCounter}0001`, "2020": Math.round(electricTotal * 0.92), "2021": Math.round(electricTotal * 0.95), "2022": electricTotal, "2023": Math.round(electricTotal * 1.03)}
        ],
        pdr_fatturati: [
          {Nome: "PDR Gas", Codice: `${siteCounter}123456789012`, "2020": Math.round(gasTotal * 0.93), "2021": Math.round(gasTotal * 0.96), "2022": gasTotal, "2023": Math.round(gasTotal * 1.02)}
        ],
        teleriscaldamento: (Math.random() > 0.98 && siteCounter % 60 === 0) ? {
          Codice: `TLR-${city.name}-${siteCounter}`,
          Nome_PDR: `Teleriscaldamento ${city.name}`,
          kWht: Math.round(150000 + Math.random() * 20000)
        } : null,
        consumi_uso: {
          Illuminazione: illuminazione,
          Climatizzazione: climatizzazione,
          "Apparati e Sistemi tecnologici": apparati,
          "FM utenze": fm
        }
      };
      
      additionalSites.push(site);
      siteCounter++;
    }
  });
  
  return additionalSites;
}

// NO DATA PRELOADED - Dashboard starts empty
// Users must upload their own Excel file
const ALL_SITES_DATA = [];

console.log('Data.js loaded: Dashboard vuota - nessun dato precaricato. Carica un file Excel per iniziare.');
