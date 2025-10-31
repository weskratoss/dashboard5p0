// Real data sites loaded from JSON
let userSites = [];
let siteIdCounter = 1;
let isLoadingRealData = false;

// Store parsed Excel data
let excelData = {
  modello: [],
  podFatturati: [],
  pdrFatturati: [],
  teleriscaldamento: []
};

// KPI data per category (CORRECT VALUES from image)
const KPI_PER_CATEGORIA = {
  bronze: {
    num_siti: 13,
    kpi_centralina: 2.000,
    kpi_stazione: 0.217,
    kpi_fv: 0.070,
    kpi_illuminazione: 0.065,
    kpi_lfm: 12.249
  },
  silver: {
    num_siti: 65,
    kpi_centralina: 4.738,
    kpi_stazione: 0.105,
    kpi_fv: 0.022,
    kpi_illuminazione: 0.075,
    kpi_lfm: 14.861
  },
  gold: {
    num_siti: 37,
    kpi_centralina: 4.852,
    kpi_stazione: 0.069,
    kpi_fv: 0.020,
    kpi_illuminazione: 0.089,
    kpi_lfm: 30.717
  },
  platinum: {
    num_siti: 5,
    kpi_centralina: 2.815,
    kpi_stazione: 0.070,
    kpi_fv: 0.021,
    kpi_illuminazione: 0.169,
    kpi_lfm: 48.043
  }
};

// TEP Conversion constants
const TEP_CONVERSIONS = {
  ELECTRIC_KWH_PER_TEP: 5347,  // 1 TEP = 5347 kWh
  GAS_SMC_PER_TEP: 1187        // 1 TEP = 1187 Smc
};

// Simulation base values per category
const SIMULATION_BASE = {
  platinum: { elettrico: 1500000, gas: 400000, superficie: 8000, dipendenti: 200 },
  gold: { elettrico: 1200000, gas: 350000, superficie: 6000, dipendenti: 150 },
  silver: { elettrico: 900000, gas: 300000, superficie: 5000, dipendenti: 120 },
  bronze: { elettrico: 700000, gas: 250000, superficie: 4000, dipendenti: 80 },
  officina: { elettrico: 800000, gas: 280000, superficie: 4500, dipendenti: 100 },
  ust: { elettrico: 950000, gas: 320000, superficie: 5500, dipendenti: 130 }
};

// Generate simulated consumption data for a site
function generateSiteConsumption(category) {
  const base = SIMULATION_BASE[category.toLowerCase()] || SIMULATION_BASE.bronze;
  const variation = 0.8 + Math.random() * 0.4; // 80-120% variation
  
  const electricKwh = Math.round(base.elettrico * variation);
  const gasSmc = Math.round(base.gas * variation);
  const superficie = Math.round(base.superficie * variation);
  const dipendenti = Math.round(base.dipendenti * variation);
  const electricTep = kwhToTep(electricKwh);
  const gasTep = smcToTep(gasSmc);
  const costoAnnuo = (electricKwh * ENERGY_COSTS.ELECTRIC_EUR_KWH) + (gasSmc * ENERGY_COSTS.GAS_EUR_SMC);
  
  return {
    elettrico_kwh: electricKwh,
    elettrico_tep: electricTep,
    gas_smc: gasSmc,
    gas_tep: gasTep,
    superficie_mq: superficie,
    numero_dipendenti: dipendenti,
    kwh_per_mq: superficie > 0 ? electricKwh / superficie : 0,
    kwh_per_dipendente: dipendenti > 0 ? electricKwh / dipendenti : 0,
    costo_annuo_euro: costoAnnuo
  };
}

// Generate buildings for a site
function generateBuildings(siteId, numBuildings = null) {
  const count = numBuildings || (2 + Math.floor(Math.random() * 3)); // 2-4 buildings
  const buildings = [];
  
  for (let i = 0; i < count; i++) {
    buildings.push({
      nome: `Edificio ${String.fromCharCode(65 + i)}`,
      pod: `IT001E${siteId}${String(Math.floor(Math.random() * 1000000)).padStart(6, '0')}`,
      pdr: String(Math.floor(Math.random() * 1000000000000)).padStart(12, '0'),
      consumo_elettrico: Math.round(200000 + Math.random() * 300000),
      consumo_gas: Math.round(50000 + Math.random() * 150000),
      superficie_mq: Math.round(1000 + Math.random() * 4000)
    });
  }
  
  return buildings;
}

// Data
const appData = {
  azienda: {
    nome: "EnergyTech Italia S.p.A.",
    settore: "Tecnologia e Servizi Energetici"
  },
  consumi_totali: {
    elettrico_kwh_anno: 12517814,
    elettrico_tep_anno: 2341.09,
    gas_mc_anno: 3811927,
    gas_tep_anno: 3211.4,
    costo_elettrico_euro: 3129453.5,
    costo_gas_euro: 3240138.95,
    distribuzione_elettrico: {
      "Climatizzazione tecnologica": 3754344.2,
      "Climatizzazione civile": 3129453.5,
      "Illuminazione": 2503562.8,
      "FM": 1564726.75,
      "Apparati tecnologici": 1565727.75
    },
    distribuzione_gas: {
      "Climatizzazione tecnologica": 2553846.55,
      "Climatizzazione civile": 1144766.1,
      "Illuminazione": 0,
      "FM": 76238.54,
      "Apparati tecnologici": 37076.81
    }
  },

  storico_mensile: [
    { mese: "Gen", elettrico_kwh: 938352.6, gas_mc: 5336697.8 },
    { mese: "Feb", elettrico_kwh: 834535.12, gas_mc: 4955505.1 },
    { mese: "Mar", elettrico_kwh: 938352.6, gas_mc: 4192910.97 },
    { mese: "Apr", elettrico_kwh: 1043235.1, gas_mc: 3049541.6 },
    { mese: "Mag", elettrico_kwh: 1147970.8, gas_mc: 2286172.35 },
    { mese: "Giu", elettrico_kwh: 1356376.2, gas_mc: 1524803.1 },
    { mese: "Lug", elettrico_kwh: 1460494.3, gas_mc: 1143578.1 },
    { mese: "Ago", elettrico_kwh: 1460494.3, gas_mc: 1143578.1 },
    { mese: "Set", elettrico_kwh: 1356376.2, gas_mc: 1524803.1 },
    { mese: "Ott", elettrico_kwh: 1147970.8, gas_mc: 2286172.35 },
    { mese: "Nov", elettrico_kwh: 1043235.1, gas_mc: 3049541.6 },
    { mese: "Dic", elettrico_kwh: 938352.6, gas_mc: 4192910.97 }
  ],
  kpi: {
    consumo_per_mq: 195.8,
    consumo_per_dipendente: 7511.3,
    efficienza_media: 0.675
  }
};

// Global variables
let map;
let currentView = 'general';
let currentSite = null;
let charts = {};
let allMarkers = [];
let currentDoitFilter = 'all';
let currentCategoryFilter = 'all';
let filteredSites = [];

// Category colors - Extended to 6 categories
const categoryColors = {
  'Bronze': '#CD7F32',
  'Silver': '#C0C0C0',
  'Gold': '#FFD700',
  'Platinum': '#E5E4E2',
  'Officina': '#FF6B35',
  'Ust': '#4ECDC4'
};

// Energy costs (UPDATED)
const ENERGY_COSTS = {
  ELECTRIC_EUR_KWH: 0.21,  // Updated from 0.25
  GAS_EUR_SMC: 0.745       // Updated from 0.85
};

// Chart colors
const chartColors = ['#1FB8CD', '#FFC185', '#B4413C', '#ECEBD5', '#5D878F', '#DB4545', '#D2BA4C', '#964325', '#944454', '#13343B'];

// Convert kWh to TEP
function kwhToTep(kwh) {
  return kwh / TEP_CONVERSIONS.ELECTRIC_KWH_PER_TEP;
}

// Convert Smc to TEP
function smcToTep(smc) {
  return smc / TEP_CONVERSIONS.GAS_SMC_PER_TEP;
}

// Format TEP value
function formatTep(tep) {
  return tep.toFixed(2) + ' TEP';
}

// Update DOIT filter options based on loaded sites
function updateDoitFilter() {
  const doitSelect = document.getElementById('doit-filter');
  const currentValue = doitSelect.value;
  
  // Clear existing options except "all"
  doitSelect.innerHTML = '<option value="all">Tutti i siti</option>';
  
  if (userSites.length === 0) return;
  
  // Get unique DOIT values
  const doitSet = new Set(userSites.map(site => site.DOIT));
  const doitList = Array.from(doitSet).sort();
  
  // Populate dropdown
  doitList.forEach(doit => {
    const option = document.createElement('option');
    option.value = doit;
    option.textContent = doit;
    doitSelect.appendChild(option);
  });
  
  // Restore previous selection if still valid
  if (currentValue !== 'all' && doitList.includes(currentValue)) {
    doitSelect.value = currentValue;
  } else {
    doitSelect.value = 'all';
    currentDoitFilter = 'all';
  }
}

// Apply both DOIT and Category filters
function applyFilters() {
  // Filter sites by both DOIT and Category
  filteredSites = userSites;
  
  if (currentDoitFilter !== 'all') {
    filteredSites = filteredSites.filter(site => site.DOIT === currentDoitFilter);
  }
  
  if (currentCategoryFilter !== 'all') {
    filteredSites = filteredSites.filter(site => site.metallo.toLowerCase() === currentCategoryFilter.toLowerCase());
  }
  
  // Update filter count
  const filterCount = document.getElementById('filter-count');
  const activeFilters = [];
  if (currentDoitFilter !== 'all') activeFilters.push(`DOIT: ${currentDoitFilter}`);
  if (currentCategoryFilter !== 'all') activeFilters.push(`Categoria: ${currentCategoryFilter}`);
  
  if (activeFilters.length > 0) {
    filterCount.textContent = `${filteredSites.length} siti (${activeFilters.join(', ')})`;
  } else {
    filterCount.textContent = '';
  }
  
  // Update map markers
  updateMapMarkers();
  
  // Update statistics and legend
  if (currentView === 'general') {
    updateGeneralStats();
    updateCategoryChart();
    updateUsoEnergeticoCharts();
    updateCompetenzaChart();
    
    // Log filter results
    const kpis = calculateKPIs();
    console.log('=== FILTRI APPLICATI ===');
    console.log(`Filtro DOIT: ${currentDoitFilter}`);
    console.log(`Filtro Categoria: ${currentCategoryFilter}`);
    console.log(`Siti visibili: ${filteredSites.length}`);
    console.log(`Vettore Elettrico: ${formatNumber(Math.round(kpis.electricKwh), 0)} kWh (${formatNumber(kpis.electricTep, 2)} TEP)`);
    console.log(`Vettore Termico: ${formatNumber(Math.round(kpis.gasSmc), 0)} Smc (${formatNumber(kpis.gasTep, 2)} TEP)`);
    console.log(`Costo: €${formatNumber(Math.round(kpis.costoAnnuo), 0)}`);
    console.log('========================');
  }
  updateMapLegend();
}

// Update category filter options
function updateCategoryFilter() {
  const categorySelect = document.getElementById('category-filter');
  if (!categorySelect) return;
  
  const currentValue = categorySelect.value;
  categorySelect.innerHTML = '<option value="all">Tutte le categorie</option>';
  
  if (userSites.length === 0) return;
  
  // Get unique categories from loaded sites
  const categorySet = new Set(userSites.map(site => site.metallo.toLowerCase()));
  const categoryList = Array.from(categorySet).sort();
  
  categoryList.forEach(cat => {
    const option = document.createElement('option');
    option.value = cat;
    const displayName = cat.charAt(0).toUpperCase() + cat.slice(1);
    option.textContent = displayName;
    categorySelect.appendChild(option);
  });
  
  if (currentValue !== 'all' && categoryList.includes(currentValue)) {
    categorySelect.value = currentValue;
  } else {
    categorySelect.value = 'all';
    currentCategoryFilter = 'all';
  }
}

// Update map legend with counts (shows FILTERED sites)
function updateMapLegend() {
  const legendItems = document.querySelector('.legend-items');
  if (!legendItems) return;
  
  legendItems.innerHTML = '';
  
  // Count sites per category in FILTERED view (dynamic)
  const categoryCounts = {};
  Object.keys(categoryColors).forEach(cat => {
    categoryCounts[cat] = filteredSites.filter(s => 
      s.metallo.toLowerCase() === cat.toLowerCase()
    ).length;
  });
  
  // Display legend items
  Object.entries(categoryColors).forEach(([category, color]) => {
    const count = categoryCounts[category] || 0;
    const item = document.createElement('div');
    item.className = 'legend-item';
    item.innerHTML = `
      <span class="legend-marker" style="background-color: ${color};"></span>
      <span>${category.toUpperCase()} (${count})</span>
    `;
    legendItems.appendChild(item);
  });
}

// Update map markers based on filter
function updateMapMarkers() {
  // Remove all existing markers
  allMarkers.forEach(marker => map.removeLayer(marker));
  allMarkers = [];
  
  // Add markers for filtered sites
  filteredSites.forEach(site => {
    const marker = addSiteMarker(site);
    allMarkers.push(marker);
  });
  
  // Update empty overlay
  showEmptyMapOverlay();
}

// Initialize map
function initMap() {
  map = L.map('map').setView([41.8, 12.6], 6);
  
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '© OpenStreetMap contributors',
    maxZoom: 19
  }).addTo(map);

  // Initialize with empty sites
  filteredSites = userSites;
  
  // Show empty map overlay
  showEmptyMapOverlay();
}

// Show empty map overlay when no sites loaded
function showEmptyMapOverlay() {
  if (userSites.length === 0) {
    const mapContainer = document.getElementById('map');
    let overlay = document.getElementById('map-empty-overlay');
    
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'map-empty-overlay';
      overlay.innerHTML = `
        <div style="font-size: 48px; margin-bottom: var(--space-16);">🗺️</div>
        <h3>Mappa Vuota</h3>
        <p>Nessun sito caricato.<br>Carica un file Excel per visualizzare i marker.</p>
      `;
      mapContainer.appendChild(overlay);
    }
  } else {
    const overlay = document.getElementById('map-empty-overlay');
    if (overlay) overlay.remove();
  }
}

// Add marker for a site
function addSiteMarker(site) {
  const categoryName = site.metallo.charAt(0).toUpperCase() + site.metallo.slice(1);
  const color = categoryColors[categoryName];
  
  const icon = L.divIcon({
    className: 'custom-marker',
    html: `<div style="background-color: ${color}; width: 24px; height: 24px; border-radius: 50%; border: 3px solid white; box-shadow: 0 2px 8px rgba(0,0,0,0.3);"></div>`,
    iconSize: [24, 24],
    iconAnchor: [12, 12]
  });

  const marker = L.marker([site.lat, site.lng], { icon })
    .addTo(map);

  const category = site.metallo.charAt(0).toUpperCase() + site.metallo.slice(1);
  
  // Popup content
  const popupContent = `
    <div class="popup-title">${site.nome}</div>
    <div class="popup-doit">📍 DOIT: ${site.DOIT}</div>
    <div class="popup-category" style="background-color: ${color}; color: ${category === 'Platinum' || category === 'Gold' ? '#000' : '#fff'};">
      ${category}
    </div>
    <div class="popup-info"><strong>Coordinate:</strong> ${site.lat.toFixed(5)}, ${site.lng.toFixed(5)}</div>
  `;

  marker.bindPopup(popupContent);

  // Click event
  marker.on('click', () => {
    showSiteDetail(site);
  });
  
  return marker;
}

// Format numbers with Italian locale (comma for decimals, dot for thousands)
function formatNumber(num, decimals = 0) {
  if (num === null || num === undefined || isNaN(num)) {
    return 'N/A';
  }
  return num.toLocaleString('it-IT', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals
  });
}

// Calculate comprehensive KPIs from filtered sites
function calculateKPIs() {
  if (userSites.length === 0) {
    return {
      totalSites: 0,
      electricKwh: 0,
      electricTep: 0,
      gasSmc: 0,
      gasTep: 0,
      totalTep: 0,
      superficie: 0,
      dipendenti: 0,
      kwhPerMq: 0,
      kwhPerDipendente: 0,
      costoAnnuo: 0
    };
  }
  
  // Calculate from FILTERED sites (updates dynamically based on map filters)
  // Source: MODELLO for electric (Consumo [kWh]), PDR Fatturati for gas (2022)
  let totalElectricKwh = 0;
  let totalGasSmc = 0;
  let totalSuperficie = 0;
  let totalDipendenti = 0;
  
  // Use FILTERED sites for totals
  filteredSites.forEach(site => {
    if (site.consumption) {
      totalElectricKwh += site.consumption.elettrico_kwh || 0;
      totalGasSmc += site.consumption.gas_smc || 0;
      totalSuperficie += site.consumption.superficie_mq || 0;
      totalDipendenti += site.consumption.numero_dipendenti || 0;
    }
  });
  
  const electricTep = kwhToTep(totalElectricKwh);
  const gasTep = smcToTep(totalGasSmc);
  const costoElettrico = totalElectricKwh * ENERGY_COSTS.ELECTRIC_EUR_KWH;
  const costoGas = totalGasSmc * ENERGY_COSTS.GAS_EUR_SMC;
  const costoAnnuo = costoElettrico + costoGas;
  
  return {
    totalSites: filteredSites.length,
    electricKwh: totalElectricKwh,
    electricTep: electricTep,
    gasSmc: totalGasSmc,
    gasTep: gasTep,
    totalTep: electricTep + gasTep,
    superficie: totalSuperficie,
    dipendenti: totalDipendenti,
    kwhPerMq: totalSuperficie > 0 ? totalElectricKwh / totalSuperficie : 0,
    kwhPerDipendente: totalDipendenti > 0 ? totalElectricKwh / totalDipendenti : 0,
    costoAnnuo: costoAnnuo
  };
}

// Update general statistics with simplified KPIs (4 cards only)
function updateGeneralStats() {
    // Ottieni i siti visibili in base ai filtri attivi
    const visibleSites = getVisibleSites();
    
    if (visibleSites.length === 0) {
        document.getElementById('total-electric').textContent = '0 kWh';
        document.getElementById('total-gas').textContent = '0 Smc';
        document.getElementById('total-tep').textContent = '0 TEP';
        return;
    }

    let totalElectricKwh = 0;
    let totalGasSmc = 0;
    const consumiPerCompetenza = {};
    const consumiPerUsoEnergetico = {};

    // Itera sui siti filtrati visibili
    visibleSites.forEach(site => {
        // Trova utenze dal MODELLO per questo sito
        const utenzeDelSito = excelData.modello.filter(row => {
            return (row.SITO && site.SITO && 
                   row.SITO.toString().trim().toLowerCase() === site.SITO.toString().trim().toLowerCase()) ||
                   (row.DOIT && site.DOIT && 
                   row.DOIT.toString().trim().toLowerCase() === site.DOIT.toString().trim().toLowerCase());
        });

        // Somma consumi elettrici e aggrega per Competenza e Uso Energetico
        utenzeDelSito.forEach(utenza => {
            if (utenza['Consumo kWh'] && !isNaN(parseFloat(utenza['Consumo kWh']))) {
                const consumo = parseFloat(utenza['Consumo kWh']);
                totalElectricKwh += consumo;

                // Aggregazione per Competenza
                const competenza = utenza.Competenza || 'Non specificata';
                consumiPerCompetenza[competenza] = (consumiPerCompetenza[competenza] || 0) + consumo;

                // Aggregazione per Uso Energetico
                const usoEnergetico = utenza['Uso energetico'] || 'Non specificato';
                consumiPerUsoEnergetico[usoEnergetico] = (consumiPerUsoEnergetico[usoEnergetico] || 0) + consumo;
            }
        });

        // Trova consumi gas PDR 2022 per questo sito
        const pdrDelSito = excelData.pdrFatturati.filter(row => {
            return (row.SITO && site.SITO && 
                   row.SITO.toString().trim().toLowerCase() === site.SITO.toString().trim().toLowerCase());
        });

        pdrDelSito.forEach(pdr => {
            if (pdr['2022'] && !isNaN(parseFloat(pdr['2022']))) {
                totalGasSmc += parseFloat(pdr['2022']);
            }
        });
    });

    // Calcola TEP totali
    const totalElectricTep = totalElectricKwh / TEP_CONVERSIONS.ELECTRIC_KWH_PER_TEP;
    const totalGasTep = totalGasSmc / TEP_CONVERSIONS.GAS_SMC_PER_TEP;
    const totalTep = totalElectricTep + totalGasTep;

    // Aggiorna UI
    document.getElementById('total-electric').textContent = formatNumber(totalElectricKwh) + ' kWh';
    document.getElementById('total-gas').textContent = formatNumber(totalGasSmc) + ' Smc';
    document.getElementById('total-tep').textContent = formatNumber(totalTep, 2) + ' TEP';

    // Aggiorna grafici con i dati aggregati
    updateCompetenzaChart(consumiPerCompetenza);
    updateUsoEnergeticoCharts(consumiPerUsoEnergetico);
}

// Funzione helper per ottenere i siti visibili
function getVisibleSites() {
    const selectedDOIT = document.getElementById('filter-doit').value;
    const selectedCategory = document.getElementById('filter-category').value;

    return userSites.filter(site => {
        const doitMatch = selectedDOIT === 'all' || site.DOIT === selectedDOIT;
        const categoryMatch = selectedCategory === 'all' || site.Metallo.toLowerCase() === selectedCategory.toLowerCase();
        return doitMatch && categoryMatch;
    });
}
  
  // Update 4 main KPI cards with Italian formatting
  // Update total sites count
  document.getElementById('total-sites').textContent = kpis.totalSites;
  
  // Update sites breakdown by category (only 4 main categories) - FROM FILTERED SITES
  const breakdown = document.getElementById('sites-breakdown');
  if (breakdown) {
    const categoryCounts = {};
    ['Bronze', 'Silver', 'Gold', 'Platinum'].forEach(cat => {
      categoryCounts[cat] = filteredSites.filter(s => s.metallo.toLowerCase() === cat.toLowerCase()).length;
    });
    const breakdownHTML = Object.entries(categoryCounts)
      .filter(([_, count]) => count > 0)
      .map(([cat, count]) => `<div style="font-size: var(--font-size-xs); margin-top: var(--space-4);">${cat}: ${count}</div>`)
      .join('');
    breakdown.innerHTML = breakdownHTML;
  }
  
  // Vettore Elettrico (fonte: MODELLO, Consumo [kWh]) - Italian formatting
  document.getElementById('electric-kwh').textContent = formatNumber(Math.round(kpis.electricKwh), 0) + ' kWh';
  document.getElementById('electric-tep').textContent = formatNumber(kpis.electricTep, 2) + ' TEP';
  
  // Vettore Termico (fonte: PDR Fatturati, colonna 2022) - Italian formatting
  document.getElementById('gas-smc').textContent = formatNumber(Math.round(kpis.gasSmc), 0) + ' Smc';
  document.getElementById('gas-tep').textContent = formatNumber(kpis.gasTep, 2) + ' TEP';
  
  // Costo Energetico Totale (Italian formatting)
  document.getElementById('total-cost').textContent = '€' + formatNumber(Math.round(kpis.costoAnnuo), 0);
}

// Display empty dashboard state
function displayEmptyDashboard() {
  // Hide charts
  const chartContainers = document.querySelectorAll('.charts-grid, .chart-card');
  chartContainers.forEach(container => {
    if (container.style) container.style.display = 'none';
  });
  
  // Show empty state in company overview
  const overview = document.getElementById('company-overview');
  if (overview && userSites.length === 0) {
    const emptyState = document.createElement('div');
    emptyState.id = 'empty-dashboard-state';
    emptyState.style.cssText = 'padding: var(--space-32); text-align: center; grid-column: 1 / -1;';
    emptyState.innerHTML = `
      <div style="font-size: 64px; margin-bottom: var(--space-24);">📊</div>
      <h2 style="font-size: var(--font-size-3xl); margin-bottom: var(--space-16);">Nessun sito caricato</h2>
      <p style="font-size: var(--font-size-lg); color: var(--color-text-secondary); margin-bottom: var(--space-24);">Carica un file Excel per iniziare a visualizzare i dati energetici</p>
      <div style="max-width: 600px; margin: 0 auto; text-align: left; background: var(--color-bg-1); padding: var(--space-24); border-radius: var(--radius-lg); border: 1px solid var(--color-border);">
        <h3 style="font-size: var(--font-size-lg); margin-bottom: var(--space-16);">📄 Come caricare i dati:</h3>
        <ol style="padding-left: var(--space-20); margin: 0; line-height: 1.8;">
          <li>Clicca su <strong>"Seleziona File Excel"</strong> nel pannello di gestione sopra</li>
          <li>Scegli un file Excel (.xlsx) dal tuo computer</li>
          <li>Il file verrà elaborato automaticamente</li>
          <li>I siti appariranno sulla mappa e le statistiche verranno calcolate</li>
        </ol>
      </div>
    `;
    
    // Remove existing empty state if present
    const existing = document.getElementById('empty-dashboard-state');
    if (existing) existing.remove();
    
    // Insert after KPI grid
    const kpiGrid = overview.querySelector('.kpi-grid');
    if (kpiGrid && kpiGrid.nextSibling) {
      kpiGrid.parentNode.insertBefore(emptyState, kpiGrid.nextSibling);
    }
  }
}

// Initialize charts
function initCharts() {
  // Sites by category chart (4 main categories) - uses FILTERED sites
  const categoryCtx = document.getElementById('sites-category-chart');
  if (!categoryCtx) return;
  
  const categoryCounts = {
    'Bronze': filteredSites.filter(s => s.metallo === 'bronze').length,
    'Silver': filteredSites.filter(s => s.metallo === 'silver').length,
    'Gold': filteredSites.filter(s => s.metallo === 'gold').length,
    'Platinum': filteredSites.filter(s => s.metallo === 'platinum').length
  };
  
  charts.sitesCategory = new Chart(categoryCtx.getContext('2d'), {
    type: 'doughnut',
    data: {
      labels: Object.keys(categoryCounts),
      datasets: [{
        data: Object.values(categoryCounts),
        backgroundColor: Object.keys(categoryCounts).map(cat => categoryColors[cat]),
        borderWidth: 2,
        borderColor: '#fff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            boxWidth: 12,
            padding: 10,
            font: { size: 11 }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const label = context.label || '';
              const value = context.parsed;
              return `${label}: ${value} siti`;
            }
          }
        }
      }
    }
  });
}

// Update category chart with FILTERED data
function updateCategoryChart() {
  if (!charts.sitesCategory) return;
  
  const categoryCounts = {
    'Platinum': filteredSites.filter(s => s.metallo === 'platinum').length,
    'Gold': filteredSites.filter(s => s.metallo === 'gold').length,
    'Silver': filteredSites.filter(s => s.metallo === 'silver').length,
    'Bronze': filteredSites.filter(s => s.metallo === 'bronze').length
  };
  
  charts.sitesCategory.data.datasets[0].data = Object.values(categoryCounts);
  charts.sitesCategory.update();
}

// Update Uso Energetico charts based on filtered sites
function updateUsoEnergeticoCharts() {
  if (!excelData.modello || excelData.modello.length === 0) return;
  
  // Calculate uso energetico only for filtered sites
  const usoData = {};
  let totalKwh = 0;
  
  filteredSites.forEach(site => {
    const nomeSito = site.nome;
    const utenzeDelSito = excelData.modello.filter(u => u['Stazione'] === nomeSito);
    
    utenzeDelSito.forEach(utenza => {
      const uso = utenza['Uso energetico**']?.trim();
      const kwh = parseFloat(utenza['Consumo [kWh]']) || 0;
      
      if (uso) {
        usoData[uso] = (usoData[uso] || 0) + kwh;
        totalKwh += kwh;
      }
    });
  });
  
  if (totalKwh === 0) return;
  
  const labels = Object.keys(usoData);
  const values = Object.values(usoData);
  const percentages = values.map(v => ((v / totalKwh) * 100).toFixed(1));
  
  // Update radar chart
  if (charts.usoRadar) {
    charts.usoRadar.data.labels = labels;
    charts.usoRadar.data.datasets[0].data = values;
    charts.usoRadar.update();
  }
  
  // Update pie chart
  if (charts.usoPie) {
    charts.usoPie.data.labels = labels;
    charts.usoPie.data.datasets[0].data = values;
    charts.usoPie.update();
  }
}

// Update Competenza chart based on filtered sites
function updateCompetenzaChart() {
  if (!excelData.modello || excelData.modello.length === 0) return;
  
  // Calculate competenza only for filtered sites
  const competenzaData = {};
  let totalKwh = 0;
  
  filteredSites.forEach(site => {
    const nomeSito = site.nome;
    const utenzeDelSito = excelData.modello.filter(u => u['Stazione'] === nomeSito);
    
    utenzeDelSito.forEach(utenza => {
      const competenza = utenza['Competenza ']?.trim();
      const kwh = parseFloat(utenza['Consumo [kWh]']) || 0;
      
      if (competenza) {
        competenzaData[competenza] = (competenzaData[competenza] || 0) + kwh;
        totalKwh += kwh;
      }
    });
  });
  
  if (totalKwh === 0) return;
  
  // Calculate percentages and create data array
  const dataArray = [];
  Object.entries(competenzaData).forEach(([nome, kwh]) => {
    const tep = kwhToTep(kwh);
    const percentuale = ((kwh / totalKwh) * 100).toFixed(1);
    const costo = kwh * ENERGY_COSTS.ELECTRIC_EUR_KWH;
    dataArray.push({ nome, kwh, tep, percentuale, costo });
  });
  
  // Sort by kwh descending
  dataArray.sort((a, b) => b.kwh - a.kwh);
  
  // Update bar chart
  if (charts.competenzaBar) {
    const barColors = ['#1FB8CD', '#FFC185', '#B4413C', '#5D878F', '#D2BA4C', '#964325'];
    charts.competenzaBar.data.labels = dataArray.map(d => d.nome);
    charts.competenzaBar.data.datasets[0].data = dataArray.map(d => d.kwh);
    charts.competenzaBar.data.datasets[0].backgroundColor = barColors.slice(0, dataArray.length);
    charts.competenzaBar.update();
  }
  
  // Update table
  const tbody = document.getElementById('competenza-tbody');
  if (tbody) {
    tbody.innerHTML = '';
    
    dataArray.forEach(data => {
      const row = document.createElement('tr');
      row.innerHTML = `
        <td><strong>${data.nome}</strong></td>
        <td>${formatNumber(Math.round(data.kwh), 0)} kWh</td>
        <td>${formatNumber(data.tep, 2)} TEP</td>
        <td><strong>${data.percentuale}%</strong></td>
        <td>€${formatNumber(Math.round(data.costo), 0)}</td>
      `;
      tbody.appendChild(row);
    });
    
    // Add total row
    const totalRow = document.createElement('tr');
    totalRow.style.borderTop = '2px solid var(--color-border)';
    totalRow.style.fontWeight = 'var(--font-weight-bold)';
    totalRow.innerHTML = `
      <td><strong>TOTALE</strong></td>
      <td>${formatNumber(Math.round(totalKwh), 0)} kWh</td>
      <td>${formatNumber(kwhToTep(totalKwh), 2)} TEP</td>
      <td><strong>100%</strong></td>
      <td>€${formatNumber(Math.round(totalKwh * ENERGY_COSTS.ELECTRIC_EUR_KWH), 0)}</td>
    `;
    tbody.appendChild(totalRow);
  }
}

// Show site detail with all TABs
function showSiteDetail(site) {
  currentView = 'detail';
  currentSite = site;

  // Update view title
  document.getElementById('view-title').textContent = 'Dettaglio Sito';
  document.getElementById('back-btn').style.display = 'inline-flex';

  // Hide general view, show detail view
  document.getElementById('company-overview').style.display = 'none';
  document.getElementById('site-detail').style.display = 'block';

  // Reset to first tab
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  document.querySelector('.tab-btn[data-tab="riepilogo"]').classList.add('active');
  document.getElementById('tab-riepilogo').classList.add('active');

  // Update site info
  document.getElementById('site-name').textContent = site.nome;
  
  const category = site.metallo.charAt(0).toUpperCase() + site.metallo.slice(1);
  const categoryBadge = document.getElementById('site-category');
  categoryBadge.textContent = category;
  categoryBadge.style.backgroundColor = categoryColors[category];
  categoryBadge.style.color = (category === 'Platinum' || category === 'Gold') ? '#000' : '#fff';
  
  const doitBadge = document.getElementById('site-doit-badge');
  doitBadge.textContent = '📍 ' + site.DOIT;

  // Update metadata
  const categoryKPI = KPI_PER_CATEGORIA[site.metallo];
  document.getElementById('site-location').textContent = site.localita || 'N/A';
  document.getElementById('site-coords').textContent = `${site.lat.toFixed(5)}, ${site.lng.toFixed(5)}`;
  
  // Show category average KPIs (Italian formatting)
  if (categoryKPI && !categoryKPI.kpi_note) {
    document.getElementById('site-kpi-centralina').textContent = `Media ${category}: ${formatNumber(categoryKPI.kpi_centralina, 3)}`;
    document.getElementById('site-kpi-stazione').textContent = `Media ${category}: ${formatNumber(categoryKPI.kpi_stazione, 3)}`;
  } else {
    document.getElementById('site-kpi-centralina').textContent = 'N/A';
    document.getElementById('site-kpi-stazione').textContent = 'N/A';
  }
  
  // Get consumption from site data
  const consumption = site.consumption || generateSiteConsumption(site.metallo);
  const electricKwh = consumption.elettrico_kwh || 0;
  const gasSmc = consumption.gas_smc || 0;
  const electricTep = consumption.elettrico_tep || 0;
  const gasTep = consumption.gas_tep || 0;
  const superficie = consumption.superficie_mq || 0;
  const dipendenti = consumption.numero_dipendenti || 0;
  const kwhMq = consumption.kwh_per_mq || 0;
  const kwhDip = consumption.kwh_per_dipendente || 0;
  const costo = consumption.costo_annuo_euro || 0;
  
  // Update consumption values (Italian formatting)
  document.getElementById('site-electric-kwh').textContent = `${formatNumber(Math.round(electricKwh), 0)} kWh`;
  document.getElementById('site-electric-tep').textContent = formatNumber(electricTep, 2) + ' TEP';
  document.getElementById('site-gas-smc').textContent = `${formatNumber(Math.round(gasSmc), 0)} Smc`;
  document.getElementById('site-gas-tep').textContent = formatNumber(gasTep, 2) + ' TEP';
  
  // Update extended KPIs for site (Italian formatting)
  const siteSuperficieEl = document.getElementById('site-superficie');
  const siteDipendentiEl = document.getElementById('site-dipendenti');
  const siteKwhMqEl = document.getElementById('site-kwh-mq');
  const siteKwhDipEl = document.getElementById('site-kwh-dip');
  const siteCostoEl = document.getElementById('site-costo');
  
  if (siteSuperficieEl) siteSuperficieEl.textContent = `${formatNumber(Math.round(superficie), 0)} m²`;
  if (siteDipendentiEl) siteDipendentiEl.textContent = formatNumber(dipendenti, 0);
  if (siteKwhMqEl) siteKwhMqEl.textContent = formatNumber(kwhMq, 2) + ' kWh/m²';
  if (siteKwhDipEl) siteKwhDipEl.textContent = formatNumber(Math.round(kwhDip), 0) + ' kWh';
  if (siteCostoEl) siteCostoEl.textContent = '€' + formatNumber(Math.round(costo), 0);

  // Update TAB contents
  updateTabRiepilogo(site, electricKwh);
  updateTabUtenze(site);
  updateTabPOD(site);
  updateTabPDR(site);
  updateTabTeleriscaldamento(site);
  updateTabAnalisi(site);

  // Zoom map to site
  map.setView([site.lat, site.lng], 13, {
    animate: true,
    duration: 1
  });
}

// Update TAB 1: Riepilogo
function updateTabRiepilogo(site, totalElectric) {
  // Get category KPI
  const categoryKPI = KPI_PER_CATEGORIA[site.metallo];
  const category = site.metallo.charAt(0).toUpperCase() + site.metallo.slice(1);
  
  // Add KPI table at the top if available
  const riepilogoTab = document.getElementById('tab-riepilogo');
  let kpiTable = riepilogoTab.querySelector('.site-kpi-table');
  if (!kpiTable && categoryKPI && !categoryKPI.kpi_note) {
    kpiTable = document.createElement('div');
    kpiTable.className = 'site-kpi-table card';
    kpiTable.style.cssText = 'margin-bottom: var(--space-24);';
    kpiTable.innerHTML = `
      <h3 style="margin-bottom: var(--space-16);">KPI Specifici del Sito</h3>
      <p style="margin-bottom: var(--space-16); color: var(--color-text-secondary);">Valori medi per la categoria <strong>${category}</strong></p>
      <div class="table-container">
        <table class="data-table">
          <thead>
            <tr>
              <th>KPI</th>
              <th>Valore</th>
              <th>Unità</th>
              <th>Media Categoria</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><strong>Centralina</strong></td>
              <td style="text-align: center;">-</td>
              <td>kWh_Centralina/N° treni</td>
              <td style="text-align: center; color: var(--color-primary); font-weight: var(--font-weight-bold);">${categoryKPI.kpi_centralina.toFixed(2)}</td>
            </tr>
            <tr>
              <td><strong>Stazione</strong></td>
              <td style="text-align: center;">-</td>
              <td>TEP_Stazione/(N° pax*10^-3)</td>
              <td style="text-align: center; color: var(--color-primary); font-weight: var(--font-weight-bold);">${categoryKPI.kpi_stazione.toFixed(2)}</td>
            </tr>
            <tr>
              <td><strong>FV</strong></td>
              <td style="text-align: center;">-</td>
              <td>TEP_LFM/mq_FV</td>
              <td style="text-align: center; color: var(--color-primary); font-weight: var(--font-weight-bold);">${categoryKPI.kpi_fv.toFixed(2)}</td>
            </tr>
            <tr>
              <td><strong>Illuminazione</strong></td>
              <td style="text-align: center;">-</td>
              <td>kWh_illuminazione/(lux*mq)</td>
              <td style="text-align: center; color: var(--color-primary); font-weight: var(--font-weight-bold);">${categoryKPI.kpi_illuminazione.toFixed(2)}</td>
            </tr>
            <tr>
              <td><strong>LFM</strong></td>
              <td style="text-align: center;">-</td>
              <td>kWh_LFM/mq_Stazione</td>
              <td style="text-align: center; color: var(--color-primary); font-weight: var(--font-weight-bold);">${categoryKPI.kpi_lfm.toFixed(2)}</td>
            </tr>
          </tbody>
        </table>
      </div>
      <p style="margin-top: var(--space-12); font-size: var(--font-size-sm); color: var(--color-text-secondary); font-style: italic;">ℹ️ I valori specifici del sito non sono disponibili, vengono mostrati i valori medi della categoria</p>
    `;
    riepilogoTab.insertBefore(kpiTable, riepilogoTab.firstChild);
  } else if (kpiTable && (!categoryKPI || categoryKPI.kpi_note)) {
    kpiTable.innerHTML = '<p style="padding: var(--space-16); color: var(--color-text-secondary);">KPI non disponibili per questa categoria</p>';
  }
  
  // Add consumption aggregates section
  let consumiAggregati = riepilogoTab.querySelector('.consumi-aggregati');
  if (!consumiAggregati) {
    consumiAggregati = document.createElement('div');
    consumiAggregati.className = 'consumi-aggregati card';
    consumiAggregati.style.cssText = 'margin-bottom: var(--space-24);';
    
    // Calculate totals from all sources
    let consumoModello = 0;
    if (site.utenze && site.utenze.length > 0) {
      site.utenze.forEach(u => {
        consumoModello += parseFloat(u.consumo_kwh || u.Consumo_kWh || u['Consumo (kWh)'] || u.Consumo || 0);
      });
    }
    
    let consumoPOD2022 = 0;
    if (site.pod_fatturati && site.pod_fatturati.length > 0) {
      site.pod_fatturati.forEach(pod => {
        consumoPOD2022 += parseFloat(pod.consumo_2022 || pod['2022'] || pod['Consumo 2022'] || 0);
      });
    }
    
    let consumoPDR2022 = 0;
    if (site.pdr_fatturati && site.pdr_fatturati.length > 0) {
      site.pdr_fatturati.forEach(pdr => {
        consumoPDR2022 += parseFloat(pdr.consumo_2022 || pdr['2022'] || pdr['Consumo 2022'] || 0);
      });
    }
    
    let teleriscaldamento = 0;
    if (site.teleriscaldamento) {
      teleriscaldamento = parseFloat(site.teleriscaldamento.kwht || site.teleriscaldamento.kWht || site.teleriscaldamento.Consumo || 0);
    }
    
    consumiAggregati.innerHTML = `
      <h3 style="margin-bottom: var(--space-16);">Consumi Aggregati Sito</h3>
      <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: var(--space-16);">
        <div style="padding: var(--space-16); background-color: var(--color-bg-1); border-radius: var(--radius-base);">
          <p style="font-size: var(--font-size-sm); color: var(--color-text-secondary); margin-bottom: var(--space-8);">Consumo utenze (MODELLO)</p>
          <p style="font-size: var(--font-size-xl); font-weight: var(--font-weight-bold);">${consumoModello > 0 ? formatNumber(Math.round(consumoModello), 0) + ' kWh' : 'Non disponibile'}</p>
          ${consumoModello > 0 ? `<p style="font-size: var(--font-size-sm); color: var(--color-primary);">${formatNumber(kwhToTep(consumoModello), 2)} TEP</p>` : ''}
        </div>
        <div style="padding: var(--space-16); background-color: var(--color-bg-2); border-radius: var(--radius-base);">
          <p style="font-size: var(--font-size-sm); color: var(--color-text-secondary); margin-bottom: var(--space-8);">Consumo POD 2022</p>
          <p style="font-size: var(--font-size-xl); font-weight: var(--font-weight-bold);">${consumoPOD2022 > 0 ? formatNumber(Math.round(consumoPOD2022), 0) + ' kWh' : 'Non disponibile'}</p>
          ${consumoPOD2022 > 0 ? `<p style="font-size: var(--font-size-sm); color: var(--color-primary);">${formatNumber(kwhToTep(consumoPOD2022), 2)} TEP</p>` : ''}
        </div>
        <div style="padding: var(--space-16); background-color: var(--color-bg-3); border-radius: var(--radius-base);">
          <p style="font-size: var(--font-size-sm); color: var(--color-text-secondary); margin-bottom: var(--space-8);">Consumo gas PDR 2022</p>
          <p style="font-size: var(--font-size-xl); font-weight: var(--font-weight-bold);">${consumoPDR2022 > 0 ? formatNumber(Math.round(consumoPDR2022), 0) + ' Smc' : 'Non disponibile'}</p>
          ${consumoPDR2022 > 0 ? `<p style="font-size: var(--font-size-sm); color: var(--color-primary);">${formatNumber(smcToTep(consumoPDR2022), 2)} TEP</p>` : ''}
        </div>
        <div style="padding: var(--space-16); background-color: var(--color-bg-4); border-radius: var(--radius-base);">
          <p style="font-size: var(--font-size-sm); color: var(--color-text-secondary); margin-bottom: var(--space-8);">Teleriscaldamento</p>
          <p style="font-size: var(--font-size-xl); font-weight: var(--font-weight-bold);">${teleriscaldamento > 0 ? formatNumber(Math.round(teleriscaldamento), 0) + ' kWht' : 'Non disponibile'}</p>
        </div>
      </div>
    `;
    
    // Insert before KPI grid
    const kpiGrid = riepilogoTab.querySelector('.kpi-grid');
    if (kpiGrid) {
      riepilogoTab.insertBefore(consumiAggregati, kpiGrid);
    }
  }
  
  // Update uso energetico chart
  const siteUsoCtx = document.getElementById('site-uso-chart').getContext('2d');
  if (charts.siteUso) {
    charts.siteUso.destroy();
  }
  
  // Use the distribution percentages from appData
  const distributionPercentages = {
    "Climatizzazione tecnologica": 30,
    "Climatizzazione civile": 25,
    "Illuminazione": 20,
    "FM": 12.5,
    "Apparati tecnologici": 12.5
  };
  
  // Use real data if available, otherwise use percentages
  let labels, values;
  if (site.consumi_uso && Object.keys(site.consumi_uso).length > 0) {
    labels = Object.keys(site.consumi_uso);
    values = Object.values(site.consumi_uso);
  } else {
    labels = Object.keys(distributionPercentages);
    values = Object.entries(distributionPercentages).map(
      ([key, percentage]) => (totalElectric * percentage) / 100
    );
  }

  charts.siteUso = new Chart(siteUsoCtx, {
    type: 'doughnut',
    data: {
      labels: labels,
      datasets: [{
        data: values,
        backgroundColor: chartColors,
        borderWidth: 2,
        borderColor: '#fff'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            boxWidth: 12,
            padding: 10,
            font: { size: 11 }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const label = context.label || '';
              const value = formatNumber(Math.round(context.parsed), 0);
              const tep = formatNumber(kwhToTep(context.parsed), 2);
              return [`${label}: ${value} kWh`, `TEP: ${tep}`];
            }
          }
        }
      }
    }
  });
}

// Update TAB 2: Utenze
function updateTabUtenze(site) {
  const tbody = document.getElementById('utenze-tbody');
  tbody.innerHTML = '';
  
  if (!site.utenze || site.utenze.length === 0) {
    tbody.innerHTML = '<tr><td colspan="6" class="no-data">⚠️ Nessuna utenza registrata nel MODELLO per questo sito</td></tr>';
    document.getElementById('usage-summary').innerHTML = '';
    return;
  }
  
  // Show count
  const cardHeader = tbody.closest('.card').querySelector('h3');
  if (cardHeader) {
    cardHeader.textContent = `Utenze del Sito - Questo sito ha ${site.utenze.length} utenze`;
  }
  
  // Populate table with all possible column names
  site.utenze.forEach(utenza => {
    const row = document.createElement('tr');
    const nomeUtenza = utenza.nome_utenza || utenza.Nome_Utenza || utenza['Nome Utenza'] || utenza.Utenza || 'N/A';
    const tipologia = utenza.tipologia || utenza.Tipologia || 'N/A';
    const pod = utenza.pod || utenza.POD || 'N/A';
    const competenza = utenza.competenza || utenza.Competenza || 'N/A';
    const consumo = parseFloat(utenza.consumo_kwh || utenza.Consumo_kWh || utenza['Consumo (kWh)'] || utenza.Consumo || 0);
    const uso = utenza.uso_energetico || utenza.Uso_Energetico || utenza['Uso Energetico'] || 'N/A';
    
    row.innerHTML = `
      <td>${nomeUtenza}</td>
      <td>${tipologia}</td>
      <td style="font-family: var(--font-family-mono); font-size: var(--font-size-xs);">${pod}</td>
      <td>${competenza}</td>
      <td><strong>${formatNumber(Math.round(consumo), 0)} kWh</strong></td>
      <td>${uso}</td>
    `;
    tbody.appendChild(row);
  });
  
  // Calculate usage summary and total
  const usageSummary = {};
  let totalConsumo = 0;
  site.utenze.forEach(u => {
    const uso = u.uso_energetico || u.Uso_Energetico || u['Uso Energetico'] || 'Altro';
    const consumo = parseFloat(u.consumo_kwh || u.Consumo_kWh || u['Consumo (kWh)'] || u.Consumo || 0);
    usageSummary[uso] = (usageSummary[uso] || 0) + consumo;
    totalConsumo += consumo;
  });
  
  // Display summary
  const summaryDiv = document.getElementById('usage-summary');
  summaryDiv.innerHTML = '<h4 style="margin-bottom: var(--space-16);">Totali per Uso Energetico</h4>';
  Object.entries(usageSummary).forEach(([uso, totale]) => {
    const item = document.createElement('div');
    item.className = 'usage-item';
    item.innerHTML = `
      <span class="usage-label">${uso}</span>
      <span class="usage-value">${formatNumber(Math.round(totale), 0)} kWh</span>
    `;
    summaryDiv.appendChild(item);
  });
  
  // Add grand total
  const totalItem = document.createElement('div');
  totalItem.className = 'usage-item';
  totalItem.style.borderTop = '2px solid var(--color-border)';
  totalItem.style.paddingTop = 'var(--space-12)';
  totalItem.style.marginTop = 'var(--space-12)';
  totalItem.innerHTML = `
    <span class="usage-label" style="font-weight: var(--font-weight-bold); font-size: var(--font-size-lg);">Totale Consumo</span>
    <span class="usage-value" style="font-size: var(--font-size-xl);">${formatNumber(Math.round(totalConsumo), 0)} kWh</span>
  `;
  summaryDiv.appendChild(totalItem);
  
  const tepItem = document.createElement('div');
  tepItem.className = 'usage-item';
  tepItem.innerHTML = `
    <span class="usage-label">Totale TEP</span>
    <span class="usage-value" style="color: var(--color-primary);">${formatNumber(kwhToTep(totalConsumo), 2)} TEP</span>
  `;
  summaryDiv.appendChild(tepItem);
}

// Update TAB 3: POD Fatturati
function updateTabPOD(site) {
  const podList = document.getElementById('pod-list');
  podList.innerHTML = '';
  
  if (!site.pod_fatturati || site.pod_fatturati.length === 0) {
    podList.innerHTML = '<p class="no-data">⚠️ Nessun POD fatturato registrato per questo sito</p>';
    const podTrendCanvas = document.getElementById('pod-trend-chart');
    if (podTrendCanvas) podTrendCanvas.style.display = 'none';
    return;
  }
  
  const podTrendCanvas = document.getElementById('pod-trend-chart');
  if (podTrendCanvas) podTrendCanvas.style.display = 'block';
  
  // Show count
  const headerText = document.createElement('p');
  headerText.style.cssText = 'font-size: var(--font-size-base); font-weight: var(--font-weight-semibold); margin-bottom: var(--space-16); color: var(--color-text);';
  headerText.textContent = `Questo sito ha ${site.pod_fatturati.length} punti di prelievo elettrici`;
  podList.appendChild(headerText);
  
  // Calculate totals
  let total2022 = 0;
  let total2021 = 0;
  
  // Display POD items
  site.pod_fatturati.forEach(pod => {
    const item = document.createElement('div');
    item.className = 'pod-item';
    
    const nome = pod.nome || pod.Nome || pod['Nome POD'] || 'POD';
    const codice = pod.codice || pod.Codice || pod.POD || pod['Codice POD'] || 'N/A';
    const c2020 = parseFloat(pod.consumo_2020 || pod['2020'] || pod['Consumo 2020'] || 0);
    const c2021 = parseFloat(pod.consumo_2021 || pod['2021'] || pod['Consumo 2021'] || 0);
    const c2022 = parseFloat(pod.consumo_2022 || pod['2022'] || pod['Consumo 2022'] || 0);
    const c2023 = parseFloat(pod.consumo_2023 || pod['2023'] || pod['Consumo 2023'] || 0);
    
    total2022 += c2022;
    total2021 += c2021;
    
    item.innerHTML = `
      <div class="pod-header">
        <span class="pod-name"><strong>POD:</strong> ${codice}</span>
      </div>
      <p style="margin-bottom: var(--space-12); color: var(--color-text-secondary);"><strong>Nome:</strong> ${nome}</p>
      <p style="margin-bottom: var(--space-8); font-weight: var(--font-weight-semibold);">Consumi storici:</p>
      <div class="consumption-years">
        <div class="year-data">
          <div class="year-label">2020</div>
          <div class="year-value">${c2020 > 0 ? formatNumber(Math.round(c2020), 0) + ' kWh' : 'Non disponibile'}</div>
          ${c2020 > 0 ? `<div class="kpi-metric-unit">${formatNumber(kwhToTep(c2020), 2)} TEP</div>` : ''}
        </div>
        <div class="year-data">
          <div class="year-label">2021</div>
          <div class="year-value">${c2021 > 0 ? formatNumber(Math.round(c2021), 0) + ' kWh' : 'Non disponibile'}</div>
          ${c2021 > 0 ? `<div class="kpi-metric-unit">${formatNumber(kwhToTep(c2021), 2)} TEP</div>` : ''}
        </div>
        <div class="year-data">
          <div class="year-label">2022</div>
          <div class="year-value" style="font-weight: var(--font-weight-bold); color: var(--color-primary);">${c2022 > 0 ? formatNumber(Math.round(c2022), 0) + ' kWh' : 'Non disponibile'}</div>
          ${c2022 > 0 ? `<div class="kpi-metric-unit">${formatNumber(kwhToTep(c2022), 2)} TEP</div>` : ''}
        </div>
        <div class="year-data">
          <div class="year-label">2023</div>
          <div class="year-value">${c2023 > 0 ? formatNumber(Math.round(c2023), 0) + ' kWh' : 'Non disponibile'}</div>
          ${c2023 > 0 ? `<div class="kpi-metric-unit">${formatNumber(kwhToTep(c2023), 2)} TEP</div>` : ''}
        </div>
      </div>
    `;
    podList.appendChild(item);
  });
  
  // Add totals summary
  if (total2022 > 0) {
    const totalDiv = document.createElement('div');
    totalDiv.style.cssText = 'margin-top: var(--space-24); padding: var(--space-20); background-color: var(--color-bg-1); border-radius: var(--radius-base); border: 2px solid var(--color-primary);';
    const variation = total2021 > 0 ? ((total2022 - total2021) / total2021 * 100) : 0;
    const variationText = variation >= 0 ? `+${variation.toFixed(1)}%` : `${variation.toFixed(1)}%`;
    const variationColor = variation >= 0 ? 'var(--color-error)' : 'var(--color-success)';
    
    totalDiv.innerHTML = `
      <h4 style="margin-bottom: var(--space-16); font-size: var(--font-size-lg);">Totali Sito</h4>
      <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: var(--space-16);">
        <div>
          <p style="color: var(--color-text-secondary); font-size: var(--font-size-sm); margin-bottom: var(--space-4);">Totale 2022</p>
          <p style="font-size: var(--font-size-2xl); font-weight: var(--font-weight-bold);">${formatNumber(Math.round(total2022), 0)} kWh</p>
          <p style="color: var(--color-primary); font-size: var(--font-size-sm);">${formatNumber(kwhToTep(total2022), 2)} TEP</p>
        </div>
        <div>
          <p style="color: var(--color-text-secondary); font-size: var(--font-size-sm); margin-bottom: var(--space-4);">Variazione 2021-2022</p>
          <p style="font-size: var(--font-size-2xl); font-weight: var(--font-weight-bold); color: ${variationColor};">${variationText}</p>
        </div>
      </div>
    `;
    podList.appendChild(totalDiv);
  }
  
  // Create trend chart
  const podTrendCtx = document.getElementById('pod-trend-chart').getContext('2d');
  if (charts.podTrend) charts.podTrend.destroy();
  
  const years = ['2020', '2021', '2022', '2023'];
  const datasets = site.pod_fatturati.slice(0, 5).map((pod, idx) => {
    const nome = pod.nome || pod.Nome || pod['Nome POD'] || `POD ${idx + 1}`;
    return {
      label: nome,
      data: years.map(year => parseFloat(pod[`consumo_${year}`] || pod[year] || pod[`Consumo ${year}`] || 0)),
      borderColor: chartColors[idx],
      backgroundColor: chartColors[idx] + '40',
      tension: 0.3,
      fill: false
    };
  });
  
  charts.podTrend = new Chart(podTrendCtx, {
    type: 'line',
    data: { labels: years, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${formatNumber(Math.round(ctx.parsed.y), 0)} kWh`
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { callback: (val) => formatNumber(val) }
        }
      }
    }
  });
}

// Update TAB 4: PDR Fatturati
function updateTabPDR(site) {
  const pdrList = document.getElementById('pdr-list');
  pdrList.innerHTML = '';
  
  if (!site.pdr_fatturati || site.pdr_fatturati.length === 0) {
    pdrList.innerHTML = '<p class="no-data">⚠️ Nessun PDR fatturato registrato per questo sito</p>';
    const pdrTrendCanvas = document.getElementById('pdr-trend-chart');
    if (pdrTrendCanvas) pdrTrendCanvas.style.display = 'none';
    return;
  }
  
  const pdrTrendCanvas = document.getElementById('pdr-trend-chart');
  if (pdrTrendCanvas) pdrTrendCanvas.style.display = 'block';
  
  // Show count
  const headerText = document.createElement('p');
  headerText.style.cssText = 'font-size: var(--font-size-base); font-weight: var(--font-weight-semibold); margin-bottom: var(--space-16); color: var(--color-text);';
  headerText.textContent = `Questo sito ha ${site.pdr_fatturati.length} punti di riconsegna gas`;
  pdrList.appendChild(headerText);
  
  // Calculate totals
  let total2022 = 0;
  let total2021 = 0;
  
  // Display PDR items
  site.pdr_fatturati.forEach(pdr => {
    const item = document.createElement('div');
    item.className = 'pdr-item';
    
    const nome = pdr.nome || pdr.Nome || pdr['Nome PDR'] || 'PDR';
    const codice = pdr.codice || pdr.Codice || pdr.PDR || pdr['Codice PDR'] || 'N/A';
    const c2020 = parseFloat(pdr.consumo_2020 || pdr['2020'] || pdr['Consumo 2020'] || 0);
    const c2021 = parseFloat(pdr.consumo_2021 || pdr['2021'] || pdr['Consumo 2021'] || 0);
    const c2022 = parseFloat(pdr.consumo_2022 || pdr['2022'] || pdr['Consumo 2022'] || 0);
    const c2023 = parseFloat(pdr.consumo_2023 || pdr['2023'] || pdr['Consumo 2023'] || 0);
    
    total2022 += c2022;
    total2021 += c2021;
    
    item.innerHTML = `
      <div class="pdr-header">
        <span class="pdr-name"><strong>PDR:</strong> ${codice}</span>
      </div>
      <p style="margin-bottom: var(--space-12); color: var(--color-text-secondary);"><strong>Nome:</strong> ${nome}</p>
      <p style="margin-bottom: var(--space-8); font-weight: var(--font-weight-semibold);">Consumi storici:</p>
      <div class="consumption-years">
        <div class="year-data">
          <div class="year-label">2020</div>
          <div class="year-value">${c2020 > 0 ? formatNumber(Math.round(c2020), 0) + ' Smc' : 'Non disponibile'}</div>
          ${c2020 > 0 ? `<div class="kpi-metric-unit">${formatNumber(smcToTep(c2020), 2)} TEP</div>` : ''}
        </div>
        <div class="year-data">
          <div class="year-label">2021</div>
          <div class="year-value">${c2021 > 0 ? formatNumber(Math.round(c2021), 0) + ' Smc' : 'Non disponibile'}</div>
          ${c2021 > 0 ? `<div class="kpi-metric-unit">${formatNumber(smcToTep(c2021), 2)} TEP</div>` : ''}
        </div>
        <div class="year-data">
          <div class="year-label">2022</div>
          <div class="year-value" style="font-weight: var(--font-weight-bold); color: var(--color-primary);">${c2022 > 0 ? formatNumber(Math.round(c2022), 0) + ' Smc' : 'Non disponibile'}</div>
          ${c2022 > 0 ? `<div class="kpi-metric-unit">${formatNumber(smcToTep(c2022), 2)} TEP</div>` : ''}
        </div>
        <div class="year-data">
          <div class="year-label">2023</div>
          <div class="year-value">${c2023 > 0 ? formatNumber(Math.round(c2023), 0) + ' Smc' : 'Non disponibile'}</div>
          ${c2023 > 0 ? `<div class="kpi-metric-unit">${formatNumber(smcToTep(c2023), 2)} TEP</div>` : ''}
        </div>
      </div>
    `;
    pdrList.appendChild(item);
  });
  
  // Add totals summary
  if (total2022 > 0) {
    const totalDiv = document.createElement('div');
    totalDiv.style.cssText = 'margin-top: var(--space-24); padding: var(--space-20); background-color: var(--color-bg-6); border-radius: var(--radius-base); border: 2px solid var(--color-orange-500);';
    const variation = total2021 > 0 ? ((total2022 - total2021) / total2021 * 100) : 0;
    const variationText = variation >= 0 ? `+${variation.toFixed(1)}%` : `${variation.toFixed(1)}%`;
    const variationColor = variation >= 0 ? 'var(--color-error)' : 'var(--color-success)';
    
    totalDiv.innerHTML = `
      <h4 style="margin-bottom: var(--space-16); font-size: var(--font-size-lg);">Totali Sito</h4>
      <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: var(--space-16);">
        <div>
          <p style="color: var(--color-text-secondary); font-size: var(--font-size-sm); margin-bottom: var(--space-4);">Totale 2022</p>
          <p style="font-size: var(--font-size-2xl); font-weight: var(--font-weight-bold);">${formatNumber(Math.round(total2022), 0)} Smc</p>
          <p style="color: var(--color-primary); font-size: var(--font-size-sm);">${formatNumber(smcToTep(total2022), 2)} TEP</p>
        </div>
        <div>
          <p style="color: var(--color-text-secondary); font-size: var(--font-size-sm); margin-bottom: var(--space-4);">Variazione 2021-2022</p>
          <p style="font-size: var(--font-size-2xl); font-weight: var(--font-weight-bold); color: ${variationColor};">${variationText}</p>
        </div>
      </div>
    `;
    pdrList.appendChild(totalDiv);
  }
  
  // Create trend chart
  const pdrTrendCtx = document.getElementById('pdr-trend-chart').getContext('2d');
  if (charts.pdrTrend) charts.pdrTrend.destroy();
  
  const years = ['2020', '2021', '2022', '2023'];
  const datasets = site.pdr_fatturati.slice(0, 5).map((pdr, idx) => {
    const nome = pdr.nome || pdr.Nome || pdr['Nome PDR'] || `PDR ${idx + 1}`;
    return {
      label: nome,
      data: years.map(year => parseFloat(pdr[`consumo_${year}`] || pdr[year] || pdr[`Consumo ${year}`] || 0)),
      backgroundColor: chartColors[idx],
      borderRadius: 6
    };
  });
  
  charts.pdrTrend = new Chart(pdrTrendCtx, {
    type: 'bar',
    data: { labels: years, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom' },
        tooltip: {
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${formatNumber(Math.round(ctx.parsed.y), 0)} Smc`
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          ticks: { callback: (val) => formatNumber(val) }
        }
      }
    }
  });
}

// Update TAB 5: Teleriscaldamento
function updateTabTeleriscaldamento(site) {
  const content = document.getElementById('teleriscaldamento-content');
  
  if (!site.teleriscaldamento) {
    content.innerHTML = '<p class="no-data">⚠️ Nessun dato teleriscaldamento per questo sito</p>';
    return;
  }
  
  const tele = site.teleriscaldamento;
  const codice = tele.codice || tele.Codice || tele['Codice Teleriscaldamento'] || 'N/A';
  const nomePdr = tele.nome_pdr || tele.Nome_PDR || tele['Nome PDR'] || 'N/A';
  const kwht = parseFloat(tele.kwht || tele.kWht || tele['Consumo'] || 0);
  
  content.innerHTML = `
    <p style="font-size: var(--font-size-base); font-weight: var(--font-weight-semibold); margin-bottom: var(--space-16);">Dati teleriscaldamento disponibili</p>
    <div class="kpi-grid" style="grid-template-columns: repeat(2, 1fr);">
      <div class="kpi-card" style="background-color: var(--color-bg-6);">
        <div class="kpi-icon">🏷️</div>
        <div class="kpi-label">Codice Teleriscaldamento</div>
        <div class="kpi-value" style="font-size: var(--font-size-base); font-family: var(--font-family-mono);">${codice}</div>
      </div>
      <div class="kpi-card" style="background-color: var(--color-bg-7);">
        <div class="kpi-icon">📛</div>
        <div class="kpi-label">Nome PDR</div>
        <div class="kpi-value" style="font-size: var(--font-size-lg);">${nomePdr}</div>
      </div>
      <div class="kpi-card" style="background-color: var(--color-bg-8); grid-column: 1 / -1;">
        <div class="kpi-icon">🌡️</div>
        <div class="kpi-label">Consumo Teleriscaldamento</div>
        <div class="kpi-value">${formatNumber(Math.round(kwht), 0)} kWht</div>
        ${kwht > 0 ? `<div class="kpi-subtitle">Equivalente TEP: ${formatNumber(kwht / TEP_CONVERSIONS.ELECTRIC_KWH_PER_TEP, 2)}</div>` : ''}
      </div>
    </div>
  `;
}

// Update TAB 6: Analisi
function updateTabAnalisi(site) {
  // Count KPIs
  const podCount = site.pod_fatturati?.length || 0;
  const pdrCount = site.pdr_fatturati?.length || 0;
  const utenzeCount = site.utenze?.length || 0;
  
  // Calculate total and average consumption from utenze
  let totalConsumoUtenze = 0;
  if (utenzeCount > 0) {
    site.utenze.forEach(u => {
      totalConsumoUtenze += parseFloat(u.consumo_kwh || u.Consumo_kWh || u['Consumo (kWh)'] || u.Consumo || 0);
    });
  }
  const avgConsumo = utenzeCount > 0 ? totalConsumoUtenze / utenzeCount : 0;
  
  document.getElementById('analisi-pod-count').textContent = podCount;
  document.getElementById('analisi-pdr-count').textContent = pdrCount;
  document.getElementById('analisi-utenze-count').textContent = utenzeCount;
  document.getElementById('analisi-avg-consumo').textContent = formatNumber(Math.round(avgConsumo), 0) + ' kWh';
  
  // Get category KPI for comparison
  const categoryKPI = KPI_PER_CATEGORIA[site.metallo];
  
  // Add KPI comparison section if available
  const analisiSection = document.getElementById('tab-analisi');
  let kpiCompSection = analisiSection.querySelector('.kpi-comparison-section');
  if (!kpiCompSection) {
    kpiCompSection = document.createElement('div');
    kpiCompSection.className = 'kpi-comparison-section';
    kpiCompSection.style.cssText = 'margin-bottom: var(--space-32);';
    analisiSection.insertBefore(kpiCompSection, analisiSection.querySelector('.kpi-grid'));
  }
  
  if (categoryKPI && !categoryKPI.kpi_note) {
    kpiCompSection.innerHTML = `
      <div class="card">
        <h3 style="margin-bottom: var(--space-20);">Performance KPI vs Media Categoria</h3>
        <p style="margin-bottom: var(--space-16); color: var(--color-text-secondary);">Confronto tra i valori medi della categoria <strong>${site.metallo.toUpperCase()}</strong> e il sito corrente</p>
        <div class="chart-container" style="height: 300px;">
          <canvas id="site-kpi-comparison-radar"></canvas>
        </div>
        <div style="margin-top: var(--space-24);">
          <table class="data-table" style="width: 100%;">
            <thead>
              <tr>
                <th>KPI</th>
                <th>Media Categoria</th>
                <th>Unità</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>Centralina</td>
                <td><strong>${formatNumber(categoryKPI.kpi_centralina, 3)}</strong></td>
                <td>kWh_Centralina/N° treni</td>
              </tr>
              <tr>
                <td>Stazione</td>
                <td><strong>${formatNumber(categoryKPI.kpi_stazione, 3)}</strong></td>
                <td>TEP_Stazione/(N° pax*10^-3)</td>
              </tr>
              <tr>
                <td>FV</td>
                <td><strong>${formatNumber(categoryKPI.kpi_fv, 3)}</strong></td>
                <td>TEP_LFM/mq_FV</td>
              </tr>
              <tr>
                <td>Illuminazione Banchine</td>
                <td><strong>${formatNumber(categoryKPI.kpi_illuminazione, 3)}</strong></td>
                <td>kWh_illuminazione/(lux*mq)</td>
              </tr>
              <tr>
                <td>LFM</td>
                <td><strong>${formatNumber(categoryKPI.kpi_lfm, 3)}</strong></td>
                <td>kWh_LFM/mq_Stazione</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    `;
    
    // Create radar chart for KPI comparison
    setTimeout(() => {
      const radarCtx = document.getElementById('site-kpi-comparison-radar');
      if (radarCtx) {
        if (charts.siteKpiRadar) charts.siteKpiRadar.destroy();
        
        charts.siteKpiRadar = new Chart(radarCtx, {
          type: 'radar',
          data: {
            labels: ['Centralina', 'Stazione', 'FV', 'Illumin. Banchine', 'LFM'],
            datasets: [{
              label: `Media ${site.metallo.toUpperCase()}`,
              data: [
                categoryKPI.kpi_centralina,
                categoryKPI.kpi_stazione,
                categoryKPI.kpi_fv,
                categoryKPI.kpi_illuminazione,
                categoryKPI.kpi_lfm
              ],
              backgroundColor: categoryColors[site.metallo.charAt(0).toUpperCase() + site.metallo.slice(1)] + '40',
              borderColor: categoryColors[site.metallo.charAt(0).toUpperCase() + site.metallo.slice(1)],
              borderWidth: 2
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
              r: {
                beginAtZero: true,
                max: 6,
                ticks: {
                  stepSize: 1
                }
              }
            },
            plugins: {
              legend: {
                position: 'top'
              }
            }
          }
        });
      }
    }, 100);
  } else {
    kpiCompSection.innerHTML = '<p style="padding: var(--space-16); color: var(--color-text-secondary);">KPI non disponibili per questa categoria</p>';
  }
  
  // Electric trend chart
  const electricTrendCtx = document.getElementById('analisi-electric-trend').getContext('2d');
  if (charts.analisiElectricTrend) charts.analisiElectricTrend.destroy();
  
  if (podCount > 0) {
    const years = ['2020', '2021', '2022', '2023'];
    const totals = years.map(year => 
      site.pod_fatturati.reduce((sum, pod) => 
        sum + parseFloat(pod[`consumo_${year}`] || pod[year] || pod[`Consumo ${year}`] || 0), 0
      )
    );
    
    charts.analisiElectricTrend = new Chart(electricTrendCtx, {
      type: 'line',
      data: {
        labels: years,
        datasets: [{
          label: 'Consumo Elettrico Totale',
          data: totals,
          borderColor: chartColors[0],
          backgroundColor: chartColors[0] + '40',
          tension: 0.3,
          fill: true
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => `${formatNumber(Math.round(ctx.parsed.y))} kWh`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: { callback: (val) => formatNumber(val) }
          }
        }
      }
    });
  }
  
  // Gas trend chart
  const gasTrendCtx = document.getElementById('analisi-gas-trend').getContext('2d');
  if (charts.analisiGasTrend) charts.analisiGasTrend.destroy();
  
  if (pdrCount > 0) {
    const years = ['2020', '2021', '2022', '2023'];
    const totals = years.map(year => 
      site.pdr_fatturati.reduce((sum, pdr) => 
        sum + parseFloat(pdr[`consumo_${year}`] || pdr[year] || pdr[`Consumo ${year}`] || 0), 0
      )
    );
    
    charts.analisiGasTrend = new Chart(gasTrendCtx, {
      type: 'line',
      data: {
        labels: years,
        datasets: [{
          label: 'Consumo Gas Totale',
          data: totals,
          borderColor: chartColors[1],
          backgroundColor: chartColors[1] + '40',
          tension: 0.3,
          fill: true
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => `${formatNumber(Math.round(ctx.parsed.y))} Smc`
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: { callback: (val) => formatNumber(val) }
          }
        }
      }
    });
  }
}

// Update buildings list
function updateBuildingsList(edifici) {
  const buildingsList = document.getElementById('buildings-list');
  buildingsList.innerHTML = '';

  edifici.forEach(edificio => {
    const buildingCard = document.createElement('div');
    buildingCard.className = 'building-card';
    buildingCard.innerHTML = `
      <div class="building-header">
        <div class="building-name">${edificio.nome}</div>
        <div class="building-year">Anno: ${edificio.anno_costruzione}</div>
      </div>
      <div class="building-codes">
        <div class="building-code">
          <div class="code-label">POD (Elettrico)</div>
          <div class="code-value">${edificio.pod}</div>
        </div>
        <div class="building-code">
          <div class="code-label">PDR (Gas)</div>
          <div class="code-value">${edificio.pdr}</div>
        </div>
      </div>
      <div class="building-stats">
        <div class="building-stat">
          <div class="stat-label">⚡ Elettrico</div>
          <div class="stat-value">${formatNumber(edificio.consumo_elettrico_annuo)} kWh</div>
        </div>
        <div class="building-stat">
          <div class="stat-label">🔥 Gas</div>
          <div class="stat-value">${formatNumber(edificio.consumo_gas_annuo)} m³</div>
        </div>
        <div class="building-stat">
          <div class="stat-label">📐 Superficie</div>
          <div class="stat-value">${formatNumber(edificio.superficie_mq)} m²</div>
        </div>
      </div>
    `;
    buildingsList.appendChild(buildingCard);
  });
}

// Show general view
function showGeneralView() {
  currentView = 'general';
  currentSite = null;

  // Update view title
  document.getElementById('view-title').textContent = 'Vista Generale';
  document.getElementById('back-btn').style.display = 'none';

  // Show general view, hide detail view
  document.getElementById('company-overview').style.display = 'block';
  document.getElementById('site-detail').style.display = 'none';

  // Update statistics
  updateGeneralStats();

  // Reset map view
  map.setView([41.8, 12.6], 6, {
    animate: true,
    duration: 1
  });
}

// Display KPI Comparison Chart (grouped bars)
function displayKPIComparisonChart() {
  const section = document.getElementById('kpi-comparison-section');
  if (!section) return;
  
  section.style.display = 'block';
  
  const ctx = document.getElementById('kpi-comparison-chart');
  if (!ctx) return;
  
  if (charts.kpiComparison) {
    charts.kpiComparison.destroy();
  }
  
  // Data from instructions
  const categories = ['Bronze', 'Silver', 'Gold', 'Platinum'];
  const categoryColors = {
    'Bronze': '#CD7F32',
    'Silver': '#C0C0C0',
    'Gold': '#FFD700',
    'Platinum': '#E5E4E2'
  };
  
  const kpiData = {
    'KPI Centralina': [2.000, 4.738, 4.852, 2.815],
    'KPI Stazione': [0.217, 0.105, 0.069, 0.070],
    'KPI FV': [0.070, 0.022, 0.020, 0.021],
    'KPI Illuminazione': [0.065, 0.075, 0.089, 0.169],
    'KPI LFM': [12.249, 14.861, 30.717, 48.043]
  };
  
  const datasets = categories.map((cat, idx) => {
    return {
      label: cat,
      data: Object.values(kpiData).map(values => values[idx]),
      backgroundColor: categoryColors[cat],
      borderRadius: 4,
      borderWidth: 0
    };
  });
  
  charts.kpiComparison = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: Object.keys(kpiData),
      datasets: datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'top',
          labels: {
            boxWidth: 15,
            padding: 15,
            font: { size: 12, weight: 'bold' }
          }
        },
        tooltip: {
          callbacks: {
            label: function(context) {
              const category = context.dataset.label;
              const value = context.parsed.y;
              return `${category}: ${formatNumber(value, 3)}`;
            }
          }
        }
      },
      scales: {
        x: {
          grid: {
            display: false
          },
          ticks: {
            font: { size: 11, weight: '500' }
          }
        },
        y: {
          beginAtZero: true,
          grid: {
            color: 'rgba(0, 0, 0, 0.05)'
          },
          ticks: {
            callback: function(value) {
              return formatNumber(value, 2);
            }
          }
        }
      }
    }
  });
}

// Display KPI by Category section with accordion (ONLY 4 categories) - HORIZONTAL BARS
function displayKPIByCategory() {
  const section = document.getElementById('kpi-by-category-section');
  if (!section) return;
  
  section.style.display = 'block';
  
  const accordionsContainer = document.getElementById('category-accordions');
  accordionsContainer.innerHTML = '';
  
  // ONLY 4 categories: Bronze, Silver, Gold, Platinum (no OFFICINA, no UST)
  const categories = ['bronze', 'silver', 'gold', 'platinum'];
  const categoryLabels = {
    bronze: 'BRONZE',
    silver: 'SILVER', 
    gold: 'GOLD',
    platinum: 'PLATINUM'
  };
  
  categories.forEach(cat => {
    const kpiData = KPI_PER_CATEGORIA[cat];
    const color = categoryColors[categoryLabels[cat]];
    const textColor = (cat === 'platinum' || cat === 'gold') ? '#000' : '#fff';
    
    const accordion = document.createElement('div');
    accordion.className = 'accordion-item';
    accordion.id = `accordion-${cat}`;
    
    let contentHTML = '';
    if (kpiData.kpi_note) {
      contentHTML = `<p style="text-align: center; padding: var(--space-24); color: var(--color-text-secondary); font-style: italic;">${kpiData.kpi_note}</p>`;
    } else {
      contentHTML = `
        <div style="margin-bottom: var(--space-24);">
          <h4 style="margin-bottom: var(--space-16); font-size: var(--font-size-base);">KPI Medi per ${categoryLabels[cat]}</h4>
          <table class="data-table" style="width: 100%;">
            <thead>
              <tr>
                <th>KPI</th>
                <th>Valore Medio</th>
                <th>Unità</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td>Centralina</td>
                <td><strong>${formatNumber(kpiData.kpi_centralina, 3)}</strong></td>
                <td>kWh_Centralina/N° treni</td>
              </tr>
              <tr>
                <td>Stazione</td>
                <td><strong>${formatNumber(kpiData.kpi_stazione, 3)}</strong></td>
                <td>TEP_Stazione/(N° pax*10^-3)</td>
              </tr>
              <tr>
                <td>FV</td>
                <td><strong>${formatNumber(kpiData.kpi_fv, 3)}</strong></td>
                <td>TEP_LFM/mq_FV</td>
              </tr>
              <tr>
                <td>Illuminazione Banchine</td>
                <td><strong>${formatNumber(kpiData.kpi_illuminazione, 3)}</strong></td>
                <td>kWh_illuminazione/(lux*mq)</td>
              </tr>
              <tr>
                <td>LFM</td>
                <td><strong>${formatNumber(kpiData.kpi_lfm, 3)}</strong></td>
                <td>kWh_LFM/mq_Stazione</td>
              </tr>
            </tbody>
          </table>
        </div>
        <div class="chart-container-large" style="height: 350px;">
          <canvas id="bar-${cat}"></canvas>
        </div>
      `;
    }
    
    accordion.innerHTML = `
      <div class="accordion-header">
        <div class="accordion-title">
          <span class="accordion-badge" style="background-color: ${color}; color: ${textColor};">${categoryLabels[cat]}</span>
          <span>${kpiData.num_siti} siti</span>
        </div>
        <span class="accordion-icon">▼</span>
      </div>
      <div class="accordion-content">
        <div class="accordion-body">
          ${contentHTML}
        </div>
      </div>
    `;
    
    accordionsContainer.appendChild(accordion);
    
    // Add click handler
    const header = accordion.querySelector('.accordion-header');
    header.addEventListener('click', () => {
      const wasActive = accordion.classList.contains('active');
      
      // Close all accordions
      document.querySelectorAll('.accordion-item').forEach(item => {
        item.classList.remove('active');
      });
      
      // Open clicked accordion if it wasn't active
      if (!wasActive) {
        accordion.classList.add('active');
        
        // Create horizontal bar chart only when opened
        if (!kpiData.kpi_note) {
          setTimeout(() => createHorizontalBarChart(cat, kpiData), 300);
        }
      }
    });
  });
}

// Display Uso Energetico section with radar chart (DYNAMIC from filtered sites)
function displayUsoEnergeticoSection() {
  const section = document.getElementById('uso-energetico-section');
  if (!section) return;
  
  section.style.display = 'block';
  
  // Calculate from MODELLO for FILTERED sites
  const usoData = {};
  let totalKwh = 0;
  
  if (excelData.modello && excelData.modello.length > 0) {
    filteredSites.forEach(site => {
      const nomeSito = site.nome;
      const utenzeDelSito = excelData.modello.filter(u => u['Stazione'] === nomeSito);
      
      utenzeDelSito.forEach(utenza => {
        const uso = utenza['Uso energetico**']?.trim();
        const kwh = parseFloat(utenza['Consumo [kWh]']) || 0;
        
        if (uso) {
          usoData[uso] = (usoData[uso] || 0) + kwh;
          totalKwh += kwh;
        }
      });
    });
  }
  
  // If no data, use default distribution
  let labels, values, percentages;
  if (totalKwh === 0 || Object.keys(usoData).length === 0) {
    // Fallback to default data
    labels = ['Apparati e Sistemi tecnologici', 'Illuminazione', 'Climatizzazione', 'FM utenze'];
    values = [13557260, 4992435, 4869756, 1500169];
    percentages = [54.4, 20.0, 19.5, 6.0];
  } else {
    labels = Object.keys(usoData);
    values = Object.values(usoData);
    percentages = values.map(v => ((v / totalKwh) * 100).toFixed(1));
  }
  
  // Create radar chart
  const radarCtx = document.getElementById('uso-energetico-radar');
  if (radarCtx) {
    if (charts.usoRadar) charts.usoRadar.destroy();
    
    charts.usoRadar = new Chart(radarCtx, {
      type: 'radar',
      data: {
        labels: labels,
        datasets: [{
          label: 'Consumi Elettrici (kWh)',
          data: values,
          backgroundColor: 'rgba(31, 184, 205, 0.2)',
          borderColor: '#1FB8CD',
          borderWidth: 3,
          pointBackgroundColor: '#1FB8CD',
          pointBorderColor: '#fff',
          pointBorderWidth: 2,
          pointRadius: 5
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          r: {
            beginAtZero: true,
            ticks: {
              callback: function(value) {
                return formatNumber(value / 1000000, 1) + 'M';
              }
            }
          }
        },
        plugins: {
          legend: {
            position: 'top',
            labels: {
              font: { size: 12 }
            }
          },
          tooltip: {
            callbacks: {
              label: function(context) {
                const value = context.parsed.r;
                const percentage = percentages[context.dataIndex];
                return [
                  `Consumo: ${formatNumber(Math.round(value), 0)} kWh`,
                  `TEP: ${formatNumber(kwhToTep(value), 2)}`,
                  `Percentuale: ${percentage}%`
                ];
              }
            }
          }
        }
      }
    });
  }
  
  // Create pie chart
  const pieCtx = document.getElementById('uso-energetico-pie');
  if (pieCtx) {
    if (charts.usoPie) charts.usoPie.destroy();
    
    charts.usoPie = new Chart(pieCtx, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: values,
          backgroundColor: [chartColors[0], chartColors[1], chartColors[2], chartColors[3]],
          borderWidth: 2,
          borderColor: '#fff'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              boxWidth: 12,
              padding: 10,
              font: { size: 11 }
            }
          },
          tooltip: {
            callbacks: {
              label: function(context) {
                const value = context.parsed;
                const percentage = percentages[context.dataIndex];
                return [
                  `${context.label}`,
                  `${formatNumber(Math.round(value), 0)} kWh`,
                  `${percentage}% del totale`,
                  `${formatNumber(kwhToTep(value), 2)} TEP`
                ];
              }
            }
          }
        }
      }
    });
  }
}

// Display Competenza section with bar chart and table (DYNAMIC from filtered sites)
function displayCompetenzaSection() {
  const section = document.getElementById('competenza-section');
  if (!section) return;
  
  section.style.display = 'block';
  
  // Calculate from MODELLO for FILTERED sites
  const competenzaMap = {};
  let totalKwh = 0;
  
  if (excelData.modello && excelData.modello.length > 0) {
    filteredSites.forEach(site => {
      const nomeSito = site.nome;
      const utenzeDelSito = excelData.modello.filter(u => u['Stazione'] === nomeSito);
      
      utenzeDelSito.forEach(utenza => {
        const competenza = utenza['Competenza ']?.trim();
        const kwh = parseFloat(utenza['Consumo [kWh]']) || 0;
        
        if (competenza) {
          competenzaMap[competenza] = (competenzaMap[competenza] || 0) + kwh;
          totalKwh += kwh;
        }
      });
    });
  }
  
  // If no data, use default
  let competenzaData;
  if (totalKwh === 0 || Object.keys(competenzaMap).length === 0) {
    // Fallback to default data
    competenzaData = [
      { nome: 'DOIT', kwh: 14420500.41, tep: 2696.85, percentuale: 57.9, costo: 3028305 },
      { nome: 'RFI', kwh: 5470620.45, tep: 1023.07, percentuale: 22.0, costo: 1148830 },
      { nome: 'DOS', kwh: 3814122.02, tep: 713.38, percentuale: 15.3, costo: 800966 },
      { nome: 'NON RFI', kwh: 960768.29, tep: 179.68, percentuale: 3.9, costo: 201761 },
      { nome: 'DIR. CIRC.', kwh: 233710.32, tep: 43.71, percentuale: 0.9, costo: 49079 },
      { nome: 'RFI - DOI', kwh: 19898.27, tep: 3.72, percentuale: 0.1, costo: 4179 }
    ];
  } else {
    // Create array from calculated data
    competenzaData = [];
    Object.entries(competenzaMap).forEach(([nome, kwh]) => {
      const tep = kwhToTep(kwh);
      const percentuale = ((kwh / totalKwh) * 100).toFixed(1);
      const costo = kwh * ENERGY_COSTS.ELECTRIC_EUR_KWH;
      competenzaData.push({ nome, kwh, tep, percentuale, costo });
    });
    // Sort by kwh descending
    competenzaData.sort((a, b) => b.kwh - a.kwh);
  }
  
  // Create horizontal bar chart
  const barCtx = document.getElementById('competenza-bar-chart');
  if (barCtx) {
    if (charts.competenzaBar) charts.competenzaBar.destroy();
    
    const barColors = ['#1FB8CD', '#FFC185', '#B4413C', '#5D878F', '#D2BA4C', '#964325'];
    
    charts.competenzaBar = new Chart(barCtx, {
      type: 'bar',
      data: {
        labels: competenzaData.map(c => c.nome),
        datasets: [{
          label: 'Consumo Elettrico (kWh)',
          data: competenzaData.map(c => c.kwh),
          backgroundColor: barColors,
          borderRadius: 6
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: false
          },
          tooltip: {
            callbacks: {
              label: function(context) {
                const index = context.dataIndex;
                const data = competenzaData[index];
                return [
                  `Consumo: ${formatNumber(Math.round(data.kwh), 0)} kWh`,
                  `TEP: ${formatNumber(data.tep, 2)}`,
                  `Percentuale: ${data.percentuale}%`,
                  `Costo: €${formatNumber(Math.round(data.costo), 0)}`
                ];
              }
            }
          }
        },
        scales: {
          x: {
            beginAtZero: true,
            ticks: {
              callback: function(value) {
                return formatNumber(value / 1000000, 1) + 'M';
              }
            }
          }
        }
      }
    });
  }
  
  // Populate table
  const tbody = document.getElementById('competenza-tbody');
  if (tbody) {
    tbody.innerHTML = '';
    
    competenzaData.forEach(data => {
      const row = document.createElement('tr');
      row.innerHTML = `
        <td><strong>${data.nome}</strong></td>
        <td>${formatNumber(Math.round(data.kwh), 0)} kWh</td>
        <td>${formatNumber(data.tep, 2)} TEP</td>
        <td><strong>${data.percentuale}%</strong></td>
        <td>€${formatNumber(Math.round(data.costo), 0)}</td>
      `;
      tbody.appendChild(row);
    });
    
    // Add total row
    const totalKwh = competenzaData.reduce((sum, d) => sum + d.kwh, 0);
    const totalTep = competenzaData.reduce((sum, d) => sum + d.tep, 0);
    const totalCosto = competenzaData.reduce((sum, d) => sum + d.costo, 0);
    
    const totalRow = document.createElement('tr');
    totalRow.style.borderTop = '2px solid var(--color-border)';
    totalRow.style.fontWeight = 'var(--font-weight-bold)';
    totalRow.innerHTML = `
      <td><strong>TOTALE</strong></td>
      <td>${formatNumber(Math.round(totalKwh), 0)} kWh</td>
      <td>${formatNumber(totalTep, 2)} TEP</td>
      <td><strong>100%</strong></td>
      <td>€${formatNumber(Math.round(totalCosto), 0)}</td>
    `;
    tbody.appendChild(totalRow);
  }
}

// KPI comparison chart removed - using accordion instead

// Create horizontal bar chart for category in accordion
function createHorizontalBarChart(category, kpiData) {
  const ctx = document.getElementById(`bar-${category}`);
  if (!ctx) return;
  
  const chartKey = `bar_${category}`;
  if (charts[chartKey]) {
    charts[chartKey].destroy();
  }
  
  const barColors = ['#1FB8CD', '#FFC185', '#B4413C', '#5D878F', '#D2BA4C'];
  
  charts[chartKey] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['Centralina', 'Stazione', 'FV', 'Illumin. Banchine', 'LFM'],
      datasets: [{
        label: category.toUpperCase(),
        data: [
          kpiData.kpi_centralina,
          kpiData.kpi_stazione,
          kpiData.kpi_fv,
          kpiData.kpi_illuminazione,
          kpiData.kpi_lfm
        ],
        backgroundColor: barColors,
        borderRadius: 6
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(context) {
              return `Valore: ${formatNumber(context.parsed.x, 3)}`;
            }
          }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          grid: {
            color: 'rgba(0, 0, 0, 0.05)'
          }
        },
        y: {
          grid: {
            display: false
          }
        }
      }
    }
  });
}

// Create radar chart for category
function createRadarChart(category, kpiData) {
  const ctx = document.getElementById(`radar-${category}`);
  if (!ctx) return;
  
  const chartKey = `radar_${category}`;
  if (charts[chartKey]) {
    charts[chartKey].destroy();
  }
  
  charts[chartKey] = new Chart(ctx, {
    type: 'radar',
    data: {
      labels: ['Centralina', 'Stazione', 'FV', 'Illumin. Banchine', 'LFM'],
      datasets: [{
        label: category.toUpperCase(),
        data: [
          kpiData.kpi_centralina,
          kpiData.kpi_stazione,
          kpiData.kpi_fv,
          kpiData.kpi_illuminazione,
          kpiData.kpi_lfm
        ],
        backgroundColor: categoryColors[category.charAt(0).toUpperCase() + category.slice(1)] + '40',
        borderColor: categoryColors[category.charAt(0).toUpperCase() + category.slice(1)],
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        r: {
          beginAtZero: true,
          max: 6
        }
      },
      plugins: {
        legend: { display: false }
      }
    }
  });
}

// Add site to list
function addSite(siteData) {
  const metallo = siteData.metallo.toLowerCase();
  
  // Use custom consumption if provided, otherwise generate
  let consumption;
  if (siteData.customConsumption && Object.keys(siteData.customConsumption).length > 0) {
    // Merge custom data with generated defaults for missing fields
    const generated = generateSiteConsumption(metallo);
    consumption = { ...generated, ...siteData.customConsumption };
  } else {
    consumption = generateSiteConsumption(metallo);
  }
  
  const site = {
    id: siteIdCounter++,
    nome: siteData.nome,
    metallo: metallo,
    DOIT: siteData.DOIT,
    localita: siteData.localita || '',
    lat: parseFloat(siteData.lat),
    lng: parseFloat(siteData.lng),
    consumption: consumption,
    buildings: generateBuildings(siteIdCounter - 1),
    utenze: siteData.utenze || [],
    pod_fatturati: siteData.podFatturati || [],
    pdr_fatturati: siteData.pdrFatturati || [],
    teleriscaldamento: siteData.teleriscaldamento || null
  };
  
  userSites.push(site);
  return site;
}

// Remove site
function removeSite(siteId) {
  userSites = userSites.filter(s => s.id !== siteId);
  refreshDashboard();
}

// Update site
function updateSite(siteId, newData) {
  const site = userSites.find(s => s.id === siteId);
  if (site) {
    site.nome = newData.nome;
    site.metallo = newData.metallo.toLowerCase();
    site.DOIT = newData.DOIT;
    site.lat = parseFloat(newData.lat);
    site.lng = parseFloat(newData.lng);
    site.consumption = generateSiteConsumption(site.metallo);
    refreshDashboard();
  }
}

// Refresh entire dashboard
function refreshDashboard() {
  updateSitesList();
  updateDoitFilter();
  updateCategoryFilter();
  applyFilters();
  updateCategoryChart();
  
  // Remove empty state if sites exist
  if (userSites.length > 0) {
    const emptyState = document.getElementById('empty-dashboard-state');
    if (emptyState) emptyState.remove();
    
    // Show charts again
    const chartContainers = document.querySelectorAll('.charts-grid, .chart-card');
    chartContainers.forEach(container => {
      if (container.style) container.style.display = '';
    });
  }
  
  if (currentView === 'general') {
    updateGeneralStats();
  }
}

// Update sites list UI
function updateSitesList() {
  const sitesCount = document.getElementById('sites-count');
  const sitesList = document.getElementById('sites-list');
  
  sitesCount.textContent = userSites.length;
  
  if (userSites.length === 0) {
    sitesList.innerHTML = '<div class="empty-state">Nessun sito caricato. Aggiungi siti manualmente o carica un file Excel.</div>';
    return;
  }
  
  sitesList.innerHTML = '';
  
  userSites.forEach(site => {
    const siteItem = document.createElement('div');
    siteItem.className = 'site-item';
    
    const category = site.metallo.charAt(0).toUpperCase() + site.metallo.slice(1);
    const categoryColor = categoryColors[category];
    const textColor = (category === 'Platinum' || category === 'Gold') ? '#000' : '#fff';
    
    siteItem.innerHTML = `
      <div class="site-item-info">
        <div class="site-item-name">${site.nome}</div>
        <div class="site-item-meta">
          <span class="site-item-category" style="background-color: ${categoryColor}; color: ${textColor};">${category}</span>
          <span>📍 ${site.DOIT}</span>
          <span>📌 ${site.lat.toFixed(4)}, ${site.lng.toFixed(4)}</span>
        </div>
      </div>
      <div class="site-item-actions">
        <button class="btn-icon edit" onclick="editSite(${site.id})" title="Modifica">✏️</button>
        <button class="btn-icon delete" onclick="deleteSite(${site.id})" title="Elimina">🗑️</button>
      </div>
    `;
    
    sitesList.appendChild(siteItem);
  });
}

// Edit site
window.editSite = function(siteId) {
  const site = userSites.find(s => s.id === siteId);
  if (!site) return;
  
  const newName = prompt('Nome Sito:', site.nome);
  if (newName === null) return;
  
  const newCategory = prompt('Categoria (bronze/silver/gold/platinum):', site.metallo);
  if (newCategory === null) return;
  
  const newDOIT = prompt('DOIT:', site.DOIT);
  if (newDOIT === null) return;
  
  const newLat = prompt('Latitudine:', site.lat);
  if (newLat === null) return;
  
  const newLng = prompt('Longitudine:', site.lng);
  if (newLng === null) return;
  
  updateSite(siteId, {
    nome: newName,
    metallo: newCategory,
    DOIT: newDOIT,
    lat: newLat,
    lng: newLng
  });
};

// Delete site
window.deleteSite = function(siteId) {
  if (confirm('Sei sicuro di voler eliminare questo sito?')) {
    removeSite(siteId);
  }
};

// Parse array data to objects with headers (REMOVED - using direct JSON parsing)

// Match data by site name (flexible matching)
function matchSiteData(siteName, dataArray, siteColumnName) {
  if (!dataArray || dataArray.length === 0) return [];
  
  const normalizedSiteName = siteName.toLowerCase().trim();
  const matched = dataArray.filter(row => {
    const rowSiteName = String(row[siteColumnName] || '').toLowerCase().trim();
    return rowSiteName === normalizedSiteName;
  });
  
  if (matched.length > 0) {
    console.log(`Matched ${matched.length} records for site: ${siteName}`);
  }
  
  return matched;
}

// Parse Excel/CSV file with multiple sheets support
function parseExcelFile(file) {
  const reader = new FileReader();
  
  reader.onload = function(e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      
      console.log('Available sheets:', workbook.SheetNames);
      
      // Check for required sheet: "KPI e coordinate"
      const kpiSheetName = workbook.SheetNames.find(name => 
        name.toLowerCase().includes('kpi') && name.toLowerCase().includes('coordinate')
      );
      
      if (!kpiSheetName) {
        showUploadStatus('Errore: Sheet "KPI e coordinate" non trovato. Sheets disponibili: ' + workbook.SheetNames.join(', '), true);
        return;
      }
      
      console.log('Using KPI sheet:', kpiSheetName);
      
      // Parse all sheets with better error handling
      const sheets = {
        kpi: null,
        modello: null,
        pod: null,
        pdr: null,
        teleriscaldamento: null
      };
      
      // Parse KPI e coordinate sheet
      try {
        sheets.kpi = XLSX.utils.sheet_to_json(workbook.Sheets[kpiSheetName], { defval: '' });
        console.log('KPI sheet parsed:', sheets.kpi.length, 'rows');
        if (sheets.kpi.length > 0) {
          console.log('KPI columns:', Object.keys(sheets.kpi[0]));
        }
      } catch (err) {
        showUploadStatus('Errore parsing KPI sheet: ' + err.message, true);
        return;
      }
      
      // Try to find MODELLO sheet - CRITICAL for electric consumption
      const modelloSheet = workbook.SheetNames.find(name => name.toLowerCase().includes('modello'));
      if (modelloSheet) {
        try {
          excelData.modello = XLSX.utils.sheet_to_json(workbook.Sheets[modelloSheet], { defval: '' });
          console.log('✅ MODELLO sheet found:', excelData.modello.length, 'utenze');
          if (excelData.modello.length > 0) {
            console.log('MODELLO columns:', Object.keys(excelData.modello[0]));
            // Pre-aggregate MODELLO data by station
            const stazioniSet = new Set();
            excelData.modello.forEach(row => {
              const stazione = row['Stazione'];
              if (stazione) stazioniSet.add(stazione);
            });
            console.log(`✅ MODELLO contains data for ${stazioniSet.size} stations`);
          }
        } catch (err) {
          console.warn('Error parsing MODELLO:', err);
        }
      } else {
        console.warn('⚠️ MODELLO sheet not found - electric consumption will be 0');
      }
      
      // Try to find POD Fatturati sheet
      const podSheet = workbook.SheetNames.find(name => 
        name.toLowerCase().includes('pod') && name.toLowerCase().includes('fatturati')
      );
      if (podSheet) {
        try {
          excelData.podFatturati = XLSX.utils.sheet_to_json(workbook.Sheets[podSheet], { defval: '' });
          console.log('POD Fatturati sheet found:', excelData.podFatturati.length, 'rows');
        } catch (err) {
          console.warn('Error parsing POD:', err);
        }
      }
      
      // Try to find PDR Fatturati sheet
      const pdrSheet = workbook.SheetNames.find(name => 
        name.toLowerCase().includes('pdr') && name.toLowerCase().includes('fatturati')
      );
      if (pdrSheet) {
        try {
          excelData.pdrFatturati = XLSX.utils.sheet_to_json(workbook.Sheets[pdrSheet], { defval: '' });
          console.log('PDR Fatturati sheet found:', excelData.pdrFatturati.length, 'rows');
          if (excelData.pdrFatturati.length > 0) {
            console.log('PDR columns:', Object.keys(excelData.pdrFatturati[0]));
          }
        } catch (err) {
          console.warn('Error parsing PDR:', err);
        }
      }
      
      // Try to find Teleriscaldamento sheet
      const teleSheet = workbook.SheetNames.find(name => name.toLowerCase().includes('teleriscaldamento'));
      if (teleSheet) {
        try {
          excelData.teleriscaldamento = XLSX.utils.sheet_to_json(workbook.Sheets[teleSheet], { defval: '' });
          console.log('Teleriscaldamento sheet found:', excelData.teleriscaldamento.length, 'rows');
        } catch (err) {
          console.warn('Error parsing Teleriscaldamento:', err);
        }
      }
      
      const rows = sheets.kpi;
      
      if (!rows || rows.length === 0) {
        showUploadStatus('Sheet KPI e coordinate vuoto', true);
        return;
      }
      
      console.log('Processing', rows.length, 'sites from KPI sheet');
      
      let addedCount = 0;
      
      // Process data rows from KPI sheet
      for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        
        // Extract site name (try multiple column names)
        const siteName = row['SITO'] || row['Sito'] || row['sito'] || row['Nome'] || row['nome'];
        if (!siteName || siteName === '') continue;
        
        // Extract metadata with fallbacks
        const metallo = (row['Metallo'] || row['metallo'] || row['Categoria'] || 'silver').toLowerCase().trim();
        const doit = String(row['DOIT'] || row['doit'] || 'N/A').trim();
        const localita = String(row['Località'] || row['Locality'] || row['localita'] || '').trim();
        
        // Extract coordinates
        let lat = parseFloat(row['lat'] || row['Lat'] || row['Latitudine'] || 0);
        let lng = parseFloat(row['long'] || row['lng'] || row['Long'] || row['Longitudine'] || 0);
        
        if (isNaN(lat) || isNaN(lng) || lat === 0 || lng === 0) {
          console.warn(`Site ${siteName}: invalid coordinates (${lat}, ${lng})`);
          continue;
        }
        
        try {
          const siteData = {
            nome: String(siteName).trim(),
            metallo: metallo,
            DOIT: doit,
            localita: localita,
            lat: lat,
            lng: lng
          };
          
          // customConsumption will be populated from MODELLO and PDR data below
          siteData.customConsumption = {};
          
          // Add optional data from other sheets
          const nomeSito = siteData.nome;
          
          // ✅ CRITICAL: Match MODELLO data for electric consumption
          // This is the PRIMARY and ONLY source for Vettore Elettrico calculations
          if (excelData.modello && excelData.modello.length > 0) {
            const utenzeData = excelData.modello.filter(row => {
              const stazione = String(row['Stazione'] || '').trim();
              return stazione === nomeSito;
            });
            siteData.utenze = utenzeData;
            
            // Calculate total electric from MODELLO - column "Consumo [kWh]"
            if (utenzeData.length > 0) {
              const totalElectric = utenzeData.reduce((sum, u) => {
                const consumo = parseFloat(u['Consumo [kWh]'] || 0);
                return sum + consumo;
              }, 0);
              
              if (totalElectric > 0) {
                if (!siteData.customConsumption) siteData.customConsumption = {};
                siteData.customConsumption.elettrico_kwh = totalElectric;
                siteData.customConsumption.elettrico_tep = kwhToTep(totalElectric);
                console.log(`✅ ${nomeSito}: Electric from MODELLO = ${formatNumber(totalElectric, 0)} kWh (${utenzeData.length} utenze)`);
              } else {
                console.log(`⚠️ ${nomeSito}: No MODELLO data or zero consumption`);
              }
            } else {
              console.log(`⚠️ ${nomeSito}: No utenze in MODELLO`);
            }
          }
          
          // Match POD Fatturati - ONLY for visualization in detail tab, NOT for totals
          if (excelData.podFatturati && excelData.podFatturati.length > 0) {
            const podData = excelData.podFatturati.filter(row => {
              const stazione = String(row['STAZIONE'] || '').trim();
              return stazione === nomeSito;
            });
            siteData.podFatturati = podData;
            if (podData.length > 0) {
              console.log(`   POD Fatturati: ${podData.length} records (visualization only, not used in totals)`);
            }
          }
          
          // ✅ Match PDR Fatturati - PRIMARY source for gas consumption (year 2022)
          if (excelData.pdrFatturati && excelData.pdrFatturati.length > 0) {
            const pdrData = excelData.pdrFatturati.filter(row => {
              const stazione = String(row['STAZIONE'] || '').trim();
              return stazione === nomeSito;
            });
            siteData.pdrFatturati = pdrData;
            
            // Calculate total gas from PDR Fatturati - column 2022
            if (pdrData.length > 0) {
              const totalGas2022 = pdrData.reduce((sum, pdr) => {
                const consumo = parseFloat(pdr['2022'] || 0);
                return sum + consumo;
              }, 0);
              
              if (totalGas2022 > 0) {
                if (!siteData.customConsumption) siteData.customConsumption = {};
                siteData.customConsumption.gas_smc = totalGas2022;
                siteData.customConsumption.gas_tep = smcToTep(totalGas2022);
                console.log(`✅ ${nomeSito}: Gas from PDR 2022 = ${formatNumber(totalGas2022, 0)} Smc (${pdrData.length} PDR)`);
              }
            }
          }
          
          // Calculate cost if we have consumption data
          if (siteData.customConsumption) {
            const elCost = (siteData.customConsumption.elettrico_kwh || 0) * ENERGY_COSTS.ELECTRIC_EUR_KWH;
            const gasCost = (siteData.customConsumption.gas_smc || 0) * ENERGY_COSTS.GAS_EUR_SMC;
            siteData.customConsumption.costo_annuo_euro = elCost + gasCost;
          }
          
          // Match Teleriscaldamento
          if (excelData.teleriscaldamento && excelData.teleriscaldamento.length > 0) {
            const teleData = matchSiteData(nomeSito, excelData.teleriscaldamento, 'STAZIONE');
            siteData.teleriscaldamento = teleData.length > 0 ? teleData[0] : null;
          }
          
          addSite(siteData);
          addedCount++;
        } catch (err) {
          console.error(`Errore processing site ${siteName}:`, err);
        }
      }
      
      // ✅ VALIDATION: Calculate totals from all sites (for verification)
      let totalElectric = 0;
      let totalGas = 0;
      let sitesWithElectric = 0;
      let sitesWithGas = 0;
      
      userSites.forEach(site => {
        if (site.consumption) {
          const elec = site.consumption.elettrico_kwh || 0;
          const gas = site.consumption.gas_smc || 0;
          totalElectric += elec;
          totalGas += gas;
          if (elec > 0) sitesWithElectric++;
          if (gas > 0) sitesWithGas++;
        }
      });
      
      console.log('\n=== ✅ VALIDAZIONE CARICAMENTO ===');
      console.log(`Siti totali caricati: ${userSites.length}`);
      console.log(`Siti con dati MODELLO (elettrico): ${sitesWithElectric}`);
      console.log(`Siti con dati PDR (gas): ${sitesWithGas}`);
      console.log(`\n📊 VETTORE ELETTRICO (fonte: MODELLO):`);
      console.log(`   ${formatNumber(totalElectric, 0)} kWh`);
      console.log(`   ${formatNumber(kwhToTep(totalElectric), 2)} TEP`);
      console.log(`\n🔥 VETTORE TERMICO (fonte: PDR Fatturati 2022):`);
      console.log(`   ${formatNumber(totalGas, 0)} Smc`);
      console.log(`   ${formatNumber(smcToTep(totalGas), 2)} TEP`);
      console.log(`\n💰 COSTO ENERGETICO TOTALE:`);
      console.log(`   €${formatNumber((totalElectric * ENERGY_COSTS.ELECTRIC_EUR_KWH) + (totalGas * ENERGY_COSTS.GAS_EUR_SMC), 0)}`);
      console.log('===================================\n');
      
      refreshDashboard();
      
      // ✅ VERIFICATION: Log calculated totals for dashboard
      console.log('\n=== ✅ VERIFICA TOTALI DASHBOARD ===');
      console.log(`Siti caricati: ${userSites.length}`);
      console.log(`Siti filtrati visibili: ${filteredSites.length}`);
      const kpis = calculateKPIs();
      console.log(`\n📊 Vettore Elettrico (MODELLO):`);
      console.log(`   ${formatNumber(Math.round(kpis.electricKwh), 0)} kWh`);
      console.log(`   ${formatNumber(kpis.electricTep, 2)} TEP`);
      console.log(`\n🔥 Vettore Termico (PDR 2022):`);
      console.log(`   ${formatNumber(Math.round(kpis.gasSmc), 0)} Smc`);
      console.log(`   ${formatNumber(kpis.gasTep, 2)} TEP`);
      console.log(`\n💰 Costo Totale: €${formatNumber(Math.round(kpis.costoAnnuo), 0)}`);
      
      // Check if values match expected
      const expectedElectric = 24919620;
      const expectedGas = 454060;
      const electricMatch = Math.abs(kpis.electricKwh - expectedElectric) < 100;
      const gasMatch = Math.abs(kpis.gasSmc - expectedGas) < 100;
      
      if (electricMatch && gasMatch) {
        console.log(`\n✅ VALORI CORRETTI! Match con valori attesi.`);
      } else {
        if (!electricMatch) {
          console.log(`\n⚠️ Vettore Elettrico differisce: atteso ~${formatNumber(expectedElectric, 0)} kWh`);
        }
        if (!gasMatch) {
          console.log(`\n⚠️ Vettore Termico differisce: atteso ~${formatNumber(expectedGas, 0)} Smc`);
        }
      }
      console.log('====================================\n');
      
      // Show KPI comparison chart
      displayKPIComparisonChart();
      
      // Show KPI section with accordion
      displayKPIByCategory();
      
      // Show uso energetico section
      displayUsoEnergeticoSection();
      
      // Show competenza section
      displayCompetenzaSection();
      
      let statusMsg = `✓ ${addedCount} siti caricati con successo!`;
      
      // Count how many sheets were found
      const sheetsFound = [];
      if (excelData.modello && excelData.modello.length > 0) sheetsFound.push('MODELLO');
      if (excelData.podFatturati && excelData.podFatturati.length > 0) sheetsFound.push('POD Fatturati');
      if (excelData.pdrFatturati && excelData.pdrFatturati.length > 0) sheetsFound.push('PDR Fatturati');
      if (excelData.teleriscaldamento && excelData.teleriscaldamento.length > 0) sheetsFound.push('Teleriscaldamento');
      
      if (sheetsFound.length > 0) {
        statusMsg += ` | Sheets caricati: ${sheetsFound.join(', ')}`;
      }
      
      const warnings = [];
      if (!excelData.modello || excelData.modello.length === 0) warnings.push('MODELLO non trovato');
      if (!excelData.pdrFatturati || excelData.pdrFatturati.length === 0) warnings.push('PDR Fatturati non trovato');
      
      if (warnings.length > 0) {
        statusMsg += ` | Avvisi: ${warnings.join(', ')}`;
      }
      
      showUploadStatus(statusMsg, false);
      
      // Zoom map to show all loaded sites
      if (userSites.length > 0) {
        const bounds = L.latLngBounds(userSites.map(s => [s.lat, s.lng]));
        map.fitBounds(bounds, { padding: [50, 50], maxZoom: 12 });
      }
    } catch (err) {
      showUploadStatus('Errore nel parsing del file: ' + err.message, true);
    }
  };
  
  reader.readAsArrayBuffer(file);
}

// Show upload status
function showUploadStatus(message, isError) {
  const statusEl = document.getElementById('upload-status');
  statusEl.textContent = message;
  statusEl.className = 'upload-status' + (isError ? ' error' : '');
  
  setTimeout(() => {
    statusEl.textContent = '';
  }, 5000);
}

// Export to JSON
function exportJSON() {
  const dataStr = JSON.stringify(userSites, null, 2);
  const dataBlob = new Blob([dataStr], { type: 'application/json' });
  const url = URL.createObjectURL(dataBlob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'siti_energetici.json';
  link.click();
  URL.revokeObjectURL(url);
}

// Export to Excel with all KPI data
function exportExcel() {
  if (userSites.length === 0) {
    alert('Nessun sito da esportare');
    return;
  }
  
  const data = userSites.map(site => {
    const c = site.consumption || {};
    return {
      'Nome': site.nome,
      'Categoria': site.metallo.charAt(0).toUpperCase() + site.metallo.slice(1),
      'DOIT': site.DOIT,
      'Latitudine': site.lat,
      'Longitudine': site.lng,
      'Consumo Elettrico (kWh)': Math.round(c.elettrico_kwh || 0),
      'Consumo Elettrico (TEP)': (c.elettrico_tep || 0).toFixed(2),
      'Consumo Gas (Smc)': Math.round(c.gas_smc || 0),
      'Consumo Gas (TEP)': (c.gas_tep || 0).toFixed(2),
      'Superficie (m²)': Math.round(c.superficie_mq || 0),
      'Numero Dipendenti': c.numero_dipendenti || 0,
      'kWh/m²': (c.kwh_per_mq || 0).toFixed(2),
      'kWh/Dipendente': Math.round(c.kwh_per_dipendente || 0),
      'Costo Annuo (€)': Math.round(c.costo_annuo_euro || 0)
    };
  });
  
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Siti Energetici');
  
  // Add KPI summary sheet
  const kpis = calculateKPIs();
  const kpiData = [
    { 'KPI': 'Totale Siti', 'Valore': kpis.totalSites },
    { 'KPI': 'Consumo Elettrico Totale (kWh)', 'Valore': Math.round(kpis.electricKwh) },
    { 'KPI': 'Consumo Elettrico Totale (TEP)', 'Valore': kpis.electricTep.toFixed(2) },
    { 'KPI': 'Consumo Gas Totale (Smc)', 'Valore': Math.round(kpis.gasSmc) },
    { 'KPI': 'Consumo Gas Totale (TEP)', 'Valore': kpis.gasTep.toFixed(2) },
    { 'KPI': 'TEP Totale', 'Valore': kpis.totalTep.toFixed(2) },
    { 'KPI': 'Superficie Totale (m²)', 'Valore': Math.round(kpis.superficie) },
    { 'KPI': 'Dipendenti Totali', 'Valore': kpis.dipendenti },
    { 'KPI': 'kWh/m² Medio', 'Valore': kpis.kwhPerMq.toFixed(2) },
    { 'KPI': 'kWh/Dipendente Medio', 'Valore': Math.round(kpis.kwhPerDipendente) },
    { 'KPI': 'Costo Energetico Annuo (€)', 'Valore': Math.round(kpis.costoAnnuo) }
  ];
  const wsKpi = XLSX.utils.json_to_sheet(kpiData);
  XLSX.utils.book_append_sheet(wb, wsKpi, 'KPI Summary');
  
  XLSX.writeFile(wb, 'report_energetico_completo.xlsx');
}

// Load real data from data.js
function loadRealData() {
  isLoadingRealData = true;
  try {
    // Check if ALL_SITES_DATA is available from data.js
    if (typeof ALL_SITES_DATA === 'undefined') {
      console.warn('ALL_SITES_DATA non disponibile, usando modalità caricamento manuale');
      isLoadingRealData = false;
      return;
    }
    const data = ALL_SITES_DATA;
    
    // Load sites from JSON
    data.forEach(site => {
      const siteObj = {
        id: siteIdCounter++,
        nome: site.SITO || site.nome,
        metallo: (site.Metallo || site.metallo || 'silver').toLowerCase(),
        DOIT: site.DOIT || 'N/A',
        localita: site.Località || site.località || site.Localita || '',
        lat: parseFloat(site.lat || site.Latitudine || 41.9),
        lng: parseFloat(site.long || site.lng || site.Longitudine || 12.5),
        kpi: {
          centralina: site.KPI_CENTRALINA || site.kpi_centralina || 'N/A',
          stazione: site.KPI_STAZIONE || site.kpi_stazione || 'N/A',
          fv: site.KPI_FV || site.kpi_fv || 'N/A',
          illuminazione_banchine: site.KPI_ILLUMINAZIONE_BANCHINE || 'N/A',
          lfm: site.KPI_LFM || site.kpi_lfm || 'N/A'
        },
        utenze: site.utenze || [],
        pod_fatturati: site.pod_fatturati || [],
        pdr_fatturati: site.pdr_fatturati || [],
        teleriscaldamento: site.teleriscaldamento || null,
        consumi_uso: site.consumi_uso || {},
        consumption: calculateSiteConsumption(site)
      };
      userSites.push(siteObj);
    });
    
    console.log(`✓ Caricati ${userSites.length} siti dal file JSON`);
    filteredSites = userSites;
    isLoadingRealData = false;
  } catch (error) {
    console.error('Errore caricamento dati:', error);
    isLoadingRealData = false;
  }
}

// Calculate site consumption from all data sources
function calculateSiteConsumption(site) {
  let elettrico_kwh = 0;
  let gas_smc = 0;
  
  // Sum from utenze
  if (site.utenze && site.utenze.length > 0) {
    elettrico_kwh = site.utenze.reduce((sum, u) => sum + (parseFloat(u.consumo_kwh || u.Consumo_kWh || 0)), 0);
  }
  
  // Sum from POD fatturati (use 2022 data)
  if (site.pod_fatturati && site.pod_fatturati.length > 0) {
    const pod2022 = site.pod_fatturati.reduce((sum, pod) => {
      return sum + (parseFloat(pod.consumo_2022 || pod['2022'] || 0));
    }, 0);
    if (pod2022 > elettrico_kwh) elettrico_kwh = pod2022;
  }
  
  // Sum from PDR fatturati (use 2022 data)
  if (site.pdr_fatturati && site.pdr_fatturati.length > 0) {
    gas_smc = site.pdr_fatturati.reduce((sum, pdr) => {
      return sum + (parseFloat(pdr.consumo_2022 || pdr['2022'] || 0));
    }, 0);
  }
  
  const elettrico_tep = kwhToTep(elettrico_kwh);
  const gas_tep = smcToTep(gas_smc);
  const costo = (elettrico_kwh * ENERGY_COSTS.ELECTRIC_EUR_KWH) + (gas_smc * ENERGY_COSTS.GAS_EUR_SMC);
  
  return {
    elettrico_kwh,
    elettrico_tep,
    gas_smc,
    gas_tep,
    superficie_mq: 5000,
    numero_dipendenti: 100,
    kwh_per_mq: elettrico_kwh / 5000,
    kwh_per_dipendente: elettrico_kwh / 100,
    costo_annuo_euro: costo
  };
}

// TAB system for site detail
function initTabSystem() {
  const tabBtns = document.querySelectorAll('.tab-btn');
  
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      const tabName = btn.getAttribute('data-tab');
      
      // Remove active from all tabs and contents
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      
      // Add active to clicked tab
      btn.classList.add('active');
      document.getElementById(`tab-${tabName}`).classList.add('active');
    });
  });
}

// Export utenze to CSV
window.exportUtenzeCSV = function() {
  if (!currentSite || !currentSite.utenze || currentSite.utenze.length === 0) {
    alert('Nessuna utenza da esportare');
    return;
  }
  
  const headers = ['Nome Utenza', 'Tipologia', 'POD', 'Competenza', 'Consumo kWh', 'Uso Energetico'];
  const rows = currentSite.utenze.map(u => [
    u.nome_utenza || u.Nome_Utenza || '',
    u.tipologia || u.Tipologia || '',
    u.pod || u.POD || '',
    u.competenza || u.Competenza || '',
    u.consumo_kwh || u.Consumo_kWh || 0,
    u.uso_energetico || u.Uso_Energetico || ''
  ]);
  
  let csv = headers.join(',') + '\n';
  rows.forEach(row => {
    csv += row.map(cell => `"${cell}"`).join(',') + '\n';
  });
  
  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `utenze_${currentSite.nome.replace(/\s+/g, '_')}.csv`;
  link.click();
  URL.revokeObjectURL(url);
};

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  // DO NOT load real data - start empty
  // loadRealData();
  
  initMap();
  initCharts();
  initTabSystem();
  updateGeneralStats();
  updateSitesList();
  applyFilters();
  updateMapLegend();
  
  // Keep management panel OPEN by default when starting empty
  // Dashboard starts empty, so users need to see upload options
  if (userSites.length === 0) {
    const panel = document.getElementById('management-panel');
    const btn = document.getElementById('toggle-management');
    if (panel && btn) {
      panel.classList.remove('collapsed');
      btn.textContent = '▼ Comprimi';
    }
  }

  // Toggle management panel
  document.getElementById('toggle-management').addEventListener('click', () => {
    const panel = document.getElementById('management-panel');
    const btn = document.getElementById('toggle-management');
    panel.classList.toggle('collapsed');
    btn.textContent = panel.classList.contains('collapsed') ? '▲ Espandi' : '▼ Comprimi';
  });

  // Manual site form removed per user request

  // Excel upload
  document.getElementById('upload-btn').addEventListener('click', () => {
    document.getElementById('excel-upload').click();
  });
  
  document.getElementById('excel-upload').addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      showUploadStatus('Caricamento in corso...', false);
      parseExcelFile(file);
    }
    e.target.value = '';
  });

  // Clear all
  document.getElementById('clear-all-btn').addEventListener('click', () => {
    if (confirm('Sei sicuro di voler eliminare tutti i siti?')) {
      userSites = [];
      refreshDashboard();
      showGeneralView();
    }
  });

  // Export buttons
  document.getElementById('export-json-btn').addEventListener('click', exportJSON);
  document.getElementById('export-excel-btn').addEventListener('click', exportExcel);
  
  // Initialize filters on page load
  updateDoitFilter();
  updateCategoryFilter();
  updateMapLegend();

  // DOIT filter - triggers dynamic recalculation
  document.getElementById('doit-filter').addEventListener('change', (e) => {
    currentDoitFilter = e.target.value;
    applyFilters();
  });
  
  // Category filter - triggers dynamic recalculation
  const categoryFilter = document.getElementById('category-filter');
  if (categoryFilter) {
    categoryFilter.addEventListener('change', (e) => {
      currentCategoryFilter = e.target.value;
      applyFilters();
    });
  }
  
  // Show empty state message
  if (userSites.length === 0) {
    showUploadStatus('Dashboard vuota. Carica un file Excel per iniziare.', false);
  }
  
  // Reset filters button - restores all sites view
  const resetFiltersBtn = document.getElementById('reset-filters-btn');
  if (resetFiltersBtn) {
    resetFiltersBtn.addEventListener('click', () => {
      currentDoitFilter = 'all';
      currentCategoryFilter = 'all';
      document.getElementById('doit-filter').value = 'all';
      if (categoryFilter) categoryFilter.value = 'all';
      applyFilters();
      
      console.log('=== FILTRI RESETTATI ===');
      console.log('Visualizzazione ripristinata a tutti i ' + userSites.length + ' siti');
      console.log('========================');
    });
  }

  // Back button
  document.getElementById('back-btn').addEventListener('click', showGeneralView);
});

function updateGeneralView() {
    // Get currently visible sites from map (filtered)
    const visibleSites = getVisibleSites();

    if (visibleSites.length === 0) {
        console.log('No visible sites for General View');
        return;
    }

    console.log('Updating General View with', visibleSites.length, 'visible sites');

    // Calculate totals from MODELLO sheet for visible sites
    let totalElectricKwh = 0;
    let totalGasSmc = 0;
    let totalTeleriscaldamentoGj = 0;

    // Aggregazioni per Competenza e Uso Energetico
    const consumiPerCompetenza = {};
    const consumiPerUsoEnergetico = {};

    // Itera sui siti visibili
    visibleSites.forEach(site => {
        // Trova le utenze di questo sito nel MODELLO
        const utenzeDelSito = excelData.modello.filter(row => {
            // Match per SITO o DOIT
            const siteMatch = row.SITO && site.SITO && 
                             row.SITO.toString().trim().toLowerCase() === site.SITO.toString().trim().toLowerCase();
            const doitMatch = row.DOIT && site.DOIT && 
                            row.DOIT.toString().trim().toLowerCase() === site.DOIT.toString().trim().toLowerCase();
            return siteMatch || doitMatch;
        });

        // Somma i consumi elettrici dal MODELLO
        utenzeDelSito.forEach(utenza => {
            // Vettore Elettrico
            if (utenza['Consumo kWh'] && !isNaN(parseFloat(utenza['Consumo kWh']))) {
                const consumo = parseFloat(utenza['Consumo kWh']);
                totalElectricKwh += consumo;

                // Aggregazione per Competenza
                const competenza = utenza.Competenza || 'Non specificata';
                if (!consumiPerCompetenza[competenza]) {
                    consumiPerCompetenza[competenza] = 0;
                }
                consumiPerCompetenza[competenza] += consumo;

                // Aggregazione per Uso Energetico
                const usoEnergetico = utenza['Uso energetico'] || 'Non specificato';
                if (!consumiPerUsoEnergetico[usoEnergetico]) {
                    consumiPerUsoEnergetico[usoEnergetico] = 0;
                }
                consumiPerUsoEnergetico[usoEnergetico] += consumo;
            }
        });

        // Trova i consumi gas per questo sito dal PDR Fatturati (anno 2022)
        const pdrDelSito = excelData.pdrFatturati.filter(row => {
            const siteMatch = row.SITO && site.SITO && 
                             row.SITO.toString().trim().toLowerCase() === site.SITO.toString().trim().toLowerCase();
            const doitMatch = row.DOIT && site.DOIT && 
                            row.DOIT.toString().trim().toLowerCase() === site.DOIT.toString().trim().toLowerCase();
            return siteMatch || doitMatch;
        });

        pdrDelSito.forEach(pdr => {
            // Somma consumi 2022
            if (pdr['2022'] && !isNaN(parseFloat(pdr['2022']))) {
                totalGasSmc += parseFloat(pdr['2022']);
            }
        });

        // Teleriscaldamento (se presente)
        const telDelSito = excelData.teleriscaldamento.filter(row => {
            const siteMatch = row.SITO && site.SITO && 
                             row.SITO.toString().trim().toLowerCase() === site.SITO.toString().trim().toLowerCase();
            return siteMatch;
        });

        telDelSito.forEach(tel => {
            if (tel.Consumo_GJ && !isNaN(parseFloat(tel.Consumo_GJ))) {
                totalTeleriscaldamentoGj += parseFloat(tel.Consumo_GJ);
            }
        });
    });

    // Conversione in TEP
    const totalElectricTep = totalElectricKwh / TEP_CONVERSIONS.ELECTRIC_KWH_PER_TEP;
    const totalGasTep = totalGasSmc / TEP_CONVERSIONS.GAS_SMC_PER_TEP;
    const totalTeleriscaldamentoTep = totalTeleriscaldamentoGj / 0.0419; // 1 TEP = 0.0419 GJ circa
    const totalTep = totalElectricTep + totalGasTep + totalTeleriscaldamentoTep;

    console.log('Calculated totals:', {
        electricKwh: totalElectricKwh,
        gasSmc: totalGasSmc,
        teleriscaldamentoGj: totalTeleriscaldamentoGj,
        totalTep: totalTep
    });

    // Aggiorna i valori nella UI
    document.getElementById('total-electric').textContent = formatNumber(totalElectricKwh) + ' kWh';
    document.getElementById('total-gas').textContent = formatNumber(totalGasSmc) + ' Smc';
    document.getElementById('total-tep').textContent = formatNumber(totalTep, 2) + ' TEP';

    // Aggiorna i grafici
    updateGeneralCharts(consumiPerCompetenza, consumiPerUsoEnergetico, totalElectricKwh, totalGasSmc);
}

// Helper function to get visible sites based on current filters
function getVisibleSites() {
    const selectedDOIT = document.getElementById('filter-doit').value;
    const selectedCategory = document.getElementById('filter-category').value;

    return userSites.filter(site => {
        const doitMatch = selectedDOIT === 'all' || site.DOIT === selectedDOIT;
        const categoryMatch = selectedCategory === 'all' || site.Metallo.toLowerCase() === selectedCategory.toLowerCase();
        return doitMatch && categoryMatch;
    });
}
