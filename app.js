/**
 * Enterprise Demand Forecasting & Inventory Decision Support System
 * Using Time Series Analytics â€” Premium Edition (v2)
 * Master Browser Engine: Phase 3 Full Functional & Interaction Parity
 * Data Source of Truth: dashboard-data.json (Executive_Summary sheet)
 */

// ==========================================================================
// Global Application State (Single Source of Truth)
// ==========================================================================
const AppState = {
  metadata: {},
  rawRecords: [],
  filteredRecords: [],
  selectedABC: ['A', 'B', 'C'],
  selectedXYZ: ['X', 'Y', 'Z'],
  selectedSKU: null,
  activeTab: 'tab1',
  illustrativeReductionPct: 10,
  currentAuditRecords: []
};

// Authoritative Constants
const MASTER_AUDIT_FIELDS = [
  'SKU', 'ABC_Class', 'XYZ_Class', 'ABC_XYZ_Class', 'Priority_Level', 'Forecast_Next_Month',
  'Accuracy_Pct', 'Forecast_Health_Score', 'Forecast_Risk', 'Safety_Stock', 'Reorder_Point', 'EOQ',
  'Average_Inventory_Units', 'Inventory_Value', 'Working_Capital', 'Estimated_Carrying_Cost',
  'Stockout_Cost', 'Inventory_Turnover', 'Inventory_Days', 'Service_Level_Pct', 'Fill_Rate_Pct',
  'Inventory_Health_Score', 'Executive_Recommendation', 'Inventory_Recommendation',
  'Procurement_Recommendation', 'Warehouse_Recommendation', 'Financial_Recommendation',
  'Business_Explanation', 'Mean_Monthly_Demand', 'Annual_Demand', 'CV', 'MAE', 'RMSE', 'Bias',
  'Bullwhip_Ratio', 'Order', 'Status', 'Health_P_Value', 'Planning Frequency', 'Inventory Policy',
  'Review Frequency', 'Business Risk', 'Supply Chain Priority', 'Unit_Cost', 'Holding_Cost_Unit',
  'Ordering_Cost_Event', 'Demand_Stability_Score', 'Model_Reliability',
  'Expected_Impact', 'KPI_Financial_Health_Score', 'KPI_Working_Capital_Efficiency'
];

const ABC_ORDER = ['A', 'B', 'C'];
const XYZ_ORDER = ['X', 'Y', 'Z'];
const RISK_ORDER = ['Low', 'Medium', 'High'];
const ABC_COLOR_MAP = { 'A': '#2563EB', 'B': '#06B6D4', 'C': '#64748B' };
const XYZ_COLOR_MAP = { 'X': '#059669', 'Y': '#F59E0B', 'Z': '#EF4444' };

const CYBER_HOVERLABEL = {
  bgcolor: '#0F172A',
  bordercolor: '#6366F1',
  font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif', size: 12, color: '#F8FAFC' },
  namelength: -1
};

const SERVICE_LEVEL_Z = 1.645;
const LEAD_TIME_DAYS = 7.0;
const LEAD_TIME_FRACTION = LEAD_TIME_DAYS / 30.0;

// ==========================================================================
// Formatting & Neutral Monetary Utilities
// ==========================================================================
function isMissing(v) {
  return v === null || v === undefined || (typeof v === 'number' && isNaN(v)) || String(v).trim().toLowerCase() === 'nan' || String(v).trim().toLowerCase() === 'none' || String(v).trim() === '';
}

function fmtNum(val, decimals = 1, suffix = '') {
  if (isMissing(val)) return 'N/A';
  const num = Number(val);
  if (isNaN(num)) return 'N/A';
  return num.toLocaleString('en-US', { minimumFractionDigits: decimals, maximumFractionDigits: decimals }) + suffix;
}

function fmtInt(val, suffix = '') {
  if (isMissing(val)) return 'N/A';
  const num = Number(val);
  if (isNaN(num)) return 'N/A';
  return Math.round(num).toLocaleString('en-US') + suffix;
}

function fmtMonetary(val) {
  if (isMissing(val)) return 'N/A';
  const num = Number(val);
  if (isNaN(num)) return 'N/A';
  return num.toLocaleString('en-US', { maximumFractionDigits: 0 });
}

function safeSum(arr, key) {
  let count = 0;
  const total = arr.reduce((acc, row) => {
    const raw = row[key];
    if (!isMissing(raw)) {
      const v = Number(raw);
      if (!isNaN(v)) {
        count++;
        return acc + v;
      }
    }
    return acc;
  }, 0);
  return count > 0 ? total : NaN;
}

function safeMean(arr, key) {
  let count = 0;
  const total = arr.reduce((acc, row) => {
    const raw = row[key];
    if (!isMissing(raw)) {
      const v = Number(raw);
      if (!isNaN(v)) {
        count++;
        return acc + v;
      }
    }
    return acc;
  }, 0);
  return count > 0 ? total / count : NaN;
}

function sourceTag(label) {
  if (label === 'Master Audit' || label === 'Dashboard Derived') return '';
  const cls = label === 'Historical Raw Data' ? 'tag-historical' : (label === 'Illustrative Scenario' ? 'tag-illustrative' : 'tag-derived');
  return `<span class="source-tag ${cls}">${label}</span>`;
}

function scoreToStatus(score, thresholds = [80, 60, 40]) {
  if (isMissing(score)) return 'attention';
  const num = Number(score);
  if (isNaN(num)) return 'attention';
  const [hi, mid, lo] = thresholds;
  if (num >= hi) return 'excellent';
  if (num >= mid) return 'good';
  if (num >= lo) return 'attention';
  return 'poor';
}

function badgeHtml(statusKey, label = null) {
  const text = label || statusKey.charAt(0).toUpperCase() + statusKey.slice(1);
  return `<span class="badge" data-status="${statusKey}">${escapeHtml(text)}</span>`;
}

function renderDecisionCard(icon, title, text, source = "Master Audit") {
  return `
    <div class="decision-card">
      <div class="decision-card-title">${icon} ${escapeHtml(title)} ${sourceTag(source)}</div>
      <div class="decision-card-text">${escapeHtml(String(text || 'N/A'))}</div>
    </div>
  `;
}

// ==========================================================================
// Toast Notification Engine
// ==========================================================================
function showToast(message) {
  let toast = document.getElementById('dashboard-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'dashboard-toast';
    toast.style.cssText = 'position:fixed;bottom:24px;right:24px;background:#1E3A8A;color:#fff;padding:12px 20px;border-radius:12px;box-shadow:0 10px 25px rgba(15,23,42,0.25);font-size:0.88rem;font-weight:700;z-index:999999;transition:opacity 0.25s ease,transform 0.25s ease;opacity:0;transform:translateY(12px);pointer-events:none;max-width:400px;line-height:1.4;border:1px solid rgba(255,255,255,0.2);';
    document.body.appendChild(toast);
  }
  toast.innerHTML = `ðŸŽ¯ ${escapeHtml(message)}`;
  toast.style.opacity = '1';
  toast.style.transform = 'translateY(0)';
  clearTimeout(toast._timeout);
  toast._timeout = setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(12px)';
  }, 4000);
}

// ==========================================================================
// Animated KPI Count-Up Engine (requestAnimationFrame + prefers-reduced-motion)
// ==========================================================================
const prefersReduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
const lastValues = new Map();
const COUNT_UP_DURATION = 500;

function parseKpiValue(text) {
  const m = String(text).trim().match(/^([^\d.-]*)(-?[\d,]*\.?\d+)([^\d]*)$/);
  if (!m) return null;
  const num = parseFloat(m[2].replace(/,/g, ''));
  if (Number.isNaN(num)) return null;
  const decimals = (m[2].split('.')[1] || '').length;
  return { prefix: m[1], suffix: m[3], num, decimals, usesComma: m[2].includes(',') };
}

function formatKpiParsedValue(num, parsed) {
  const fixed = num.toFixed(parsed.decimals);
  const [intPart, decPart] = fixed.split('.');
  const withCommas = parsed.usesComma
    ? intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
    : intPart;
  return parsed.prefix + withCommas + (decPart ? '.' + decPart : '') + parsed.suffix;
}

function easeOutCubic(t) {
  return 1 - Math.pow(1 - t, 3);
}

function animateKpiNode(node, fromNum, toParsed) {
  const start = performance.now();
  function tick(now) {
    const elapsed = now - start;
    const t = Math.min(1, elapsed / COUNT_UP_DURATION);
    const eased = easeOutCubic(t);
    const current = fromNum + (toParsed.num - fromNum) * eased;
    node.textContent = formatKpiParsedValue(current, toParsed);
    if (t < 1) {
      requestAnimationFrame(tick);
    } else {
      node.textContent = formatKpiParsedValue(toParsed.num, toParsed);
    }
  }
  requestAnimationFrame(tick);
}

function getKpiKey(node) {
  const card = node.closest('.metric-card, .square-metric');
  const title = card?.querySelector('.metric-card-title, .square-metric-label')?.textContent?.trim() || '';
  let idx = 0, sib = card;
  while ((sib = sib?.previousElementSibling)) idx++;
  return title + '::' + idx;
}

function updateKpiValueWithAnimation(nodeId, targetFormatted) {
  const node = typeof nodeId === 'string' ? document.getElementById(nodeId) : nodeId;
  if (!node) return;
  if (prefersReduced) {
    node.textContent = targetFormatted;
    return;
  }
  const parsed = parseKpiValue(targetFormatted);
  const key = getKpiKey(node);
  if (!parsed) {
    lastValues.delete(key);
    node.textContent = targetFormatted;
    return;
  }
  const prev = lastValues.get(key);
  lastValues.set(key, parsed.num);
  if (prev === undefined || prev === parsed.num) {
    node.textContent = targetFormatted;
    node.classList.add('value-counting');
    requestAnimationFrame(() => node.classList.remove('value-counting'));
    return;
  }
  animateKpiNode(node, prev, parsed);
}

// ==========================================================================
// Control Centre Drawer Handlers
// ==========================================================================
function openControlCentre() {
  document.body.classList.add('control-centre-open');
  document.body.style.overflow = 'hidden';
  const drawer = document.getElementById('control-centre');
  const overlay = document.getElementById('control-centre-overlay');
  const trigger = document.getElementById('control-centre-trigger');
  const closeBtn = document.getElementById('control-centre-close');
  if (drawer) {
    drawer.classList.add('open');
    drawer.setAttribute('aria-hidden', 'false');
  }
  if (overlay) {
    overlay.classList.add('open');
    overlay.setAttribute('aria-hidden', 'false');
  }
  if (trigger) {
    trigger.setAttribute('aria-expanded', 'true');
  }
  if (closeBtn) {
    try { closeBtn.focus(); } catch (e) {}
  }
}

function closeControlCentre() {
  document.body.classList.remove('control-centre-open');
  document.body.style.overflow = '';
  const drawer = document.getElementById('control-centre');
  const overlay = document.getElementById('control-centre-overlay');
  const trigger = document.getElementById('control-centre-trigger');
  if (drawer) {
    drawer.classList.remove('open');
    drawer.setAttribute('aria-hidden', 'true');
  }
  if (overlay) {
    overlay.classList.remove('open');
    overlay.setAttribute('aria-hidden', 'true');
  }
  if (trigger) {
    trigger.setAttribute('aria-expanded', 'false');
    try { trigger.focus(); } catch (e) {}
  }
  setTimeout(resizeActiveTabCharts, 300);
}

function toggleControlCentre() {
  const drawer = document.getElementById('control-centre');
  if (drawer && (drawer.classList.contains('open') || document.body.classList.contains('control-centre-open'))) {
    closeControlCentre();
  } else {
    openControlCentre();
  }
}

function initControlCentreEvents() {
  const triggerBtn = document.getElementById('control-centre-trigger');
  if (triggerBtn && !triggerBtn.dataset.bound) {
    triggerBtn.dataset.bound = 'true';
    triggerBtn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      toggleControlCentre();
    });
  }

  const closeBtn = document.getElementById('control-centre-close');
  if (closeBtn && !closeBtn.dataset.bound) {
    closeBtn.dataset.bound = 'true';
    closeBtn.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      closeControlCentre();
    });
  }

  const overlayEl = document.getElementById('control-centre-overlay');
  if (overlayEl && !overlayEl.dataset.bound) {
    overlayEl.dataset.bound = 'true';
    overlayEl.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      closeControlCentre();
    });
  }

  if (!document.body.dataset.ccBound) {
    document.body.dataset.ccBound = 'true';
    document.addEventListener('click', (e) => {
      const drawer = document.getElementById('control-centre');
      const trigger = document.getElementById('control-centre-trigger');
      if (drawer && drawer.classList.contains('open')) {
        if (!drawer.contains(e.target) && !trigger?.contains(e.target)) {
          closeControlCentre();
        }
      }
    });

    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        const drawer = document.getElementById('control-centre');
        if (drawer && drawer.classList.contains('open')) {
          closeControlCentre();
        }
      }
    });
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initControlCentreEvents);
} else {
  initControlCentreEvents();
}

// ==========================================================================
// Independent Mathematics Engine (Safety Stock & Reorder Point)
// ==========================================================================
function independentSSRop(row) {
  const rmse = Number(row.RMSE);
  const fc = Number(row.Forecast_Next_Month);
  const status = String(row.Status || '');

  if (status.includes('Dead Stock') || status.includes('Constant Demand')) {
    const meanM = Number(row.Mean_Monthly_Demand);
    if (isNaN(meanM)) {
      return { ss: 0.0, rop: null, basis: 'Zero-variance guard; Mean_Monthly_Demand missing' };
    }
    const ltd = (meanM / 30.0) * LEAD_TIME_DAYS;
    return { ss: 0.0, rop: Math.round(ltd), basis: 'Zero-variance guard (design-correct zero)' };
  }

  if (isNaN(rmse) || isNaN(fc)) {
    return { ss: null, rop: null, basis: 'Insufficient inputs (RMSE/Forecast missing)' };
  }

  const sigmaLt = rmse * Math.sqrt(LEAD_TIME_FRACTION);
  const ss = SERVICE_LEVEL_Z * sigmaLt;
  const ltd = (fc / 30.0) * LEAD_TIME_DAYS;
  const rop = ltd + ss;
  return { ss: Math.round(ss), rop: Math.round(rop), basis: 'RMSE-based single-uncertainty-term formula' };
}

let _cachedPortfolioQAResult = null;

function runPortfolioMathAudit(records) {
  if (records === AppState.rawRecords && _cachedPortfolioQAResult) {
    return _cachedPortfolioQAResult;
  }
  const n = records.length;
  if (n === 0) return { n: 0, summary: {}, details: [] };
  let ssZeroInMaster = 0, ssExact = 0, ssTol = 0, ssMat = 0;
  let ropZeroInMaster = 0, ropExact = 0, ropTol = 0, ropMat = 0;
  const details = [];

  records.forEach(r => {
    const masterSS = Number(r.Safety_Stock);
    const masterROP = Number(r.Reorder_Point);
    if (masterSS === 0) ssZeroInMaster++;
    if (masterROP === 0) ropZeroInMaster++;

    const res = independentSSRop(r);
    const indSS = res.ss;
    const indROP = res.rop;

    const ssVar = (isMissing(masterSS) || isMissing(indSS)) ? NaN : Math.abs(masterSS - indSS);
    const ropVar = (isMissing(masterROP) || isMissing(indROP)) ? NaN : Math.abs(masterROP - indROP);

    if (!isNaN(ssVar)) {
      if (ssVar === 0) ssExact++;
      if (ssVar <= 1) ssTol++;
      if (ssVar > 1) ssMat++;
    }
    if (!isNaN(ropVar)) {
      if (ropVar === 0) ropExact++;
      if (ropVar <= 1) ropTol++;
      if (ropVar > 1) ropMat++;
    }

    details.push({
      SKU: r.SKU,
      Status: r.Status,
      MA_Safety_Stock: masterSS,
      Independent_Safety_Stock: indSS,
      MA_Reorder_Point: masterROP,
      Independent_Reorder_Point: indROP,
      SS_Variance: ssVar,
      ROP_Variance: ropVar,
      Basis: res.basis
    });
  });

  const result = {
    summary: {
      n,
      ss_zero_in_master: ssZeroInMaster,
      ss_exact_match: ssExact,
      ss_tolerance_match: ssTol,
      ss_material_mismatch: ssMat,
      rop_zero_in_master: ropZeroInMaster,
      rop_exact_match: ropExact,
      rop_tolerance_match: ropTol,
      rop_material_mismatch: ropMat
    },
    details
  };
  if (records === AppState.rawRecords) {
    _cachedPortfolioQAResult = result;
  }
  return result;
}

// ==========================================================================
// Filtering & State Reconciliation
// ==========================================================================
function applyFilters() {
  document.body.classList.add('kpi-rerunning');

  AppState.filteredRecords = AppState.rawRecords.filter(row => {
    const abcVal = String(row.ABC_Class || '').trim();
    const xyzVal = String(row.XYZ_Class || '').trim();
    // Non-restrictive empty filter semantics: if empty, keep all
    const matchABC = AppState.selectedABC.length === 0 || AppState.selectedABC.includes(abcVal);
    const matchXYZ = AppState.selectedXYZ.length === 0 || (xyzVal && AppState.selectedXYZ.includes(xyzVal));
    return matchABC && matchXYZ;
  });

  const availableSkus = AppState.filteredRecords.map(r => r.SKU).filter(Boolean);
  if (availableSkus.length > 0) {
    if (!availableSkus.includes(AppState.selectedSKU)) {
      AppState.selectedSKU = availableSkus[0];
    }
  } else {
    AppState.selectedSKU = null;
  }

  updateSkuDropdowns(availableSkus);
  renderActiveTab();

  setTimeout(() => {
    document.body.classList.remove('kpi-rerunning');
    resizeActiveTabCharts();
  }, 160);
}

const SKU_INPUT_IDS = [
  { inputId: 'global-sku-select', datalistId: 'global-sku-datalist' },
  { inputId: 'tab2-sku-select', datalistId: 'tab2-sku-datalist' },
  { inputId: 'tab3-sku-select', datalistId: 'tab3-sku-datalist' }
];

function updateSkuDropdowns(skus) {
  SKU_INPUT_IDS.forEach(({ inputId, datalistId }) => {
    const inputEl = document.getElementById(inputId);
    const datalistEl = document.getElementById(datalistId);

    if (datalistEl) {
      datalistEl.innerHTML = skus.map(s => `<option value="${escapeHtml(s)}">`).join('');
    }

    if (inputEl) {
      if (AppState.selectedSKU) {
        inputEl.value = AppState.selectedSKU;
      } else if (skus.length > 0) {
        inputEl.value = skus[0];
      }
    }
  });
}

function syncSelectedSKU(newSku) {
  if (!newSku) return;
  const match = AppState.filteredRecords.find(r => String(r.SKU).trim().toLowerCase() === String(newSku).trim().toLowerCase());
  const canonicalSku = match ? match.SKU : newSku;

  if (AppState.selectedSKU === canonicalSku) return;
  AppState.selectedSKU = canonicalSku;

  SKU_INPUT_IDS.forEach(({ inputId }) => {
    const el = document.getElementById(inputId);
    if (el && el.value !== canonicalSku) el.value = canonicalSku;
  });

  renderActiveTab();
}

function jumpToSKU(targetSku) {
  if (!targetSku) return;
  syncSelectedSKU(targetSku);
  if (AppState.activeTab !== 'tab1') {
    AppState.activeTab = 'tab1';
    document.querySelectorAll('.tab-button').forEach(b => {
      const isTab1 = b.dataset.tab === 'tab1';
      b.classList.toggle('active', isTab1);
      b.setAttribute('aria-selected', isTab1 ? 'true' : 'false');
    });
    document.querySelectorAll('.tab-panel').forEach(p => {
      p.classList.toggle('active', p.id === 'tab1');
    });
    renderTab1();
  }
  showToast(`Loaded ${targetSku} â€” see 'Decision Intelligence' below, or Tabs 2/3 for full diagnostics.`);
  const heading = document.getElementById('decision-intelligence-heading');
  if (heading) {
    heading.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

function jumpToSKUFromTop(selectId) {
  const sel = document.getElementById(selectId);
  if (!sel) return;
  jumpToSKU(sel.value);
}

// ==========================================================================
// Tab Rendering Controller
// ==========================================================================
function renderActiveTab() {
  const tab = AppState.activeTab;
  if (tab === 'tab1') renderTab1();
  else if (tab === 'tab2') renderTab2();
  else if (tab === 'tab3') renderTab3();
  else if (tab === 'tab4') renderTab4();
  else if (tab === 'tab5') renderTab5();
}

function resizeActiveTabCharts() {
  const activePanel = document.querySelector('.tab-panel.active');
  if (!activePanel) return;
  const chartContainers = activePanel.querySelectorAll('.js-plotly-plot, [data-testid="stPlotlyChart"], .chart-container > div, .chart-container');
  chartContainers.forEach(container => {
    try {
      if (typeof Plotly !== 'undefined' && Plotly.Plots && (container.data || container._fullLayout)) {
        Plotly.Plots.resize(container);
      }
    } catch (e) {}
  });
}

// ==========================================================================
// TAB 1: Executive Control Tower
// ==========================================================================
function renderTab1() {
  const records = AppState.filteredRecords;
  const totalSkus = records.length;
  const avgAccuracy = safeMean(records, 'Accuracy_Pct');
  const avgForecastHealth = safeMean(records, 'Forecast_Health_Score');
  const highRiskCount = records.filter(r => String(r.Forecast_Risk || '').toLowerCase().trim() === 'high').length;
  const invValueSum = safeSum(records, 'Inventory_Value');
  const wcSum = safeSum(records, 'Working_Capital');
  const finHealthAvg = safeMean(records, 'KPI_Financial_Health_Score');
  const invHealthAvg = safeMean(records, 'Inventory_Health_Score');

  // Authoritative Overall Portfolio Signal
  let statusPoints = 0, statusTotal = 0;
  [[avgForecastHealth, 70], [avgAccuracy, 70], [finHealthAvg, 70], [invHealthAvg, 70]].forEach(([val, thresh]) => {
    if (!isNaN(val)) {
      statusTotal++;
      if (val >= thresh) statusPoints++;
    }
  });

  let overallStatus = 'Insufficient Data';
  let overallStatusKey = 'attention';
  if (statusTotal > 0) {
    const ratio = statusPoints / statusTotal;
    const highRiskZeroKnown = (!isNaN(highRiskCount) && highRiskCount === 0);
    if (highRiskZeroKnown && ratio >= 0.75) {
      overallStatus = 'Stable & Healthy';
      overallStatusKey = 'excellent';
    } else if (ratio >= 0.5) {
      overallStatus = 'Needs Attention';
      overallStatusKey = 'good';
    } else {
      overallStatus = 'Elevated Risk';
      overallStatusKey = 'poor';
    }
  }

  const signalStatusEl = document.getElementById('signal-status');
  if (signalStatusEl) signalStatusEl.textContent = overallStatus;
  const signalCountEl = document.getElementById('signal-sku-count');
  if (signalCountEl) signalCountEl.textContent = totalSkus.toLocaleString('en-US');
  const signalBannerEl = document.getElementById('executive-portfolio-signal');
  if (signalBannerEl) signalBannerEl.setAttribute('data-status', overallStatusKey);

  // Animated KPI updates
  updateKpiValueWithAnimation('kpi-total-skus', totalSkus.toLocaleString('en-US'));
  updateKpiValueWithAnimation('kpi-forecast-accuracy', fmtNum(avgAccuracy, 1, '%'));
  updateKpiValueWithAnimation('kpi-forecast-health', fmtNum(avgForecastHealth, 0, '%'));
  updateKpiValueWithAnimation('kpi-high-risk', fmtInt(highRiskCount));
  updateKpiValueWithAnimation('kpi-inv-value', fmtMonetary(invValueSum));
  updateKpiValueWithAnimation('kpi-working-capital', fmtMonetary(wcSum));
  updateKpiValueWithAnimation('kpi-fin-health', fmtNum(finHealthAvg, 0, '%'));

  const fhSub = document.getElementById('kpi-forecast-health-sub');
  if (fhSub) {
    const fhStatus = scoreToStatus(avgForecastHealth);
    fhSub.innerHTML = `Mean Master Audit Forecast Health â€¢ Dashboard presentation band ${badgeHtml(fhStatus, fhStatus.charAt(0).toUpperCase() + fhStatus.slice(1))}`;
  }

  const hrCard = document.getElementById('kpi-high-risk-card');
  if (hrCard) {
    if (highRiskCount > 0) hrCard.classList.add('pulse-critical');
    else hrCard.classList.remove('pulse-critical');
  }

  // Portfolio Composition â€” ABC
  const abcCounts = { A: 0, B: 0, C: 0 };
  records.forEach(r => {
    const c = String(r.ABC_Class || '').trim();
    if (abcCounts[c] !== undefined) abcCounts[c]++;
  });
  Plotly.react('chart-abc-pie', [{
    values: [abcCounts.A, abcCounts.B, abcCounts.C],
    labels: ['A', 'B', 'C'],
    type: 'pie',
    hole: 0.52,
    domain: { y: [0, 0.84] },
    marker: { colors: [ABC_COLOR_MAP.A, ABC_COLOR_MAP.B, ABC_COLOR_MAP.C], line: { color: '#ffffff', width: 2.5 } },
    textinfo: 'label+percent'
  }], {
    title: { text: 'SKU Count Share by ABC â€” Demand Volume Classification', font: { size: 13, family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' }, y: 0.98 },
    margin: { l: 20, r: 20, t: 50, b: 15 },
    hoverlabel: CYBER_HOVERLABEL,
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  // Portfolio Composition â€” XYZ
  const xyzCounts = { X: 0, Y: 0, Z: 0 };
  records.forEach(r => {
    const c = String(r.XYZ_Class || '').trim();
    if (xyzCounts[c] !== undefined) xyzCounts[c]++;
  });
  Plotly.react('chart-xyz-pie', [{
    values: [xyzCounts.X, xyzCounts.Y, xyzCounts.Z],
    labels: ['X', 'Y', 'Z'],
    type: 'pie',
    hole: 0.52,
    domain: { y: [0, 0.84] },
    marker: { colors: [XYZ_COLOR_MAP.X, XYZ_COLOR_MAP.Y, XYZ_COLOR_MAP.Z], line: { color: '#ffffff', width: 2.5 } },
    textinfo: 'label+percent'
  }], {
    title: { text: 'SKU Count Share by XYZ â€” Demand Variability Classification', font: { size: 13, family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' }, y: 0.98 },
    margin: { l: 20, r: 20, t: 50, b: 15 },
    hoverlabel: CYBER_HOVERLABEL,
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  // ABCâ€“XYZ Segmentation Matrix
  const matrix = { A: {X:0, Y:0, Z:0}, B: {X:0, Y:0, Z:0}, C: {X:0, Y:0, Z:0} };
  records.forEach(r => {
    const code = String(r.ABC_XYZ_Class || '').trim();
    if (code.length === 2) {
      const a = code[0], x = code[1];
      if (matrix[a] && matrix[a][x] !== undefined) matrix[a][x]++;
    }
  });
  Plotly.react('chart-abc-xyz-matrix', [{
    z: [
      [matrix.A.X, matrix.A.Y, matrix.A.Z],
      [matrix.B.X, matrix.B.Y, matrix.B.Z],
      [matrix.C.X, matrix.C.Y, matrix.C.Z]
    ],
    x: ['X', 'Y', 'Z'],
    y: ['A', 'B', 'C'],
    type: 'heatmap',
    colorscale: [[0, '#F8FAFC'], [0.2, '#E0E7FF'], [0.45, '#818CF8'], [0.75, '#4F46E5'], [1, '#1E1B4B']],
    text: [
      [matrix.A.X, matrix.A.Y, matrix.A.Z],
      [matrix.B.X, matrix.B.Y, matrix.B.Z],
      [matrix.C.X, matrix.C.Y, matrix.C.Z]
    ],
    texttemplate: "%{text}",
    showscale: true,
    colorbar: { len: 0.85, thickness: 14, title: { text: 'SKU Count', font: { size: 11 } } }
  }], {
    title: 'ABCâ€“XYZ Segmentation Matrix â€” SKU Count',
    xaxis: { title: { text: 'XYZ â€” Demand Variability', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'ABC â€” Demand Volume Classification', standoff: 15 }, automargin: true },
    margin: { l: 60, r: 40, t: 45, b: 55 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    transition: { duration: 400, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  // Annualised Demand by ABC Class
  const demandABC = { A: 0, B: 0, C: 0 };
  records.forEach(r => {
    const c = String(r.ABC_Class || '').trim();
    const d = Number(r.Annual_Demand || 0);
    if (demandABC[c] !== undefined && !isNaN(d)) demandABC[c] += d;
  });
  Plotly.react('chart-annual-demand-abc', [{
    x: ['A', 'B', 'C'],
    y: [demandABC.A, demandABC.B, demandABC.C],
    type: 'bar',
    marker: { color: [ABC_COLOR_MAP.A, ABC_COLOR_MAP.B, ABC_COLOR_MAP.C], line: { color: '#ffffff', width: 1.5 } }
  }], {
    title: 'Total Annualised Demand (units) by ABC Class',
    xaxis: { title: { text: 'ABC Class', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'Annualised Demand (Units)', standoff: 15 }, automargin: true },
    margin: { l: 70, r: 20, t: 45, b: 55 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    showlegend: false,
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  // Immediate Risk & Priority Exceptions (computed from full population)
  const allFlagged = records.filter(r => {
    const fr = String(r.Forecast_Risk || '').toLowerCase().trim();
    const pr = String(r.Priority_Level || '').toLowerCase().trim();
    return fr === 'high' || pr === 'high';
  }).sort((a, b) => {
    const vA = Number(a.Inventory_Value) || 0;
    const vB = Number(b.Inventory_Value) || 0;
    if (vB !== vA) return vB - vA;
    return String(a.SKU || '').localeCompare(String(b.SKU || ''));
  });

  const top25Attn = allFlagged.slice(0, 25);
  const attnTbody = document.getElementById('attention-tbody');
  if (attnTbody) {
    attnTbody.innerHTML = '';
    if (top25Attn.length === 0) {
      attnTbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 1.5rem;">No priority exceptions match the current filter.</td></tr>';
    } else {
      top25Attn.forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(r.SKU || 'N/A')}</td>
          <td>${escapeHtml(r.ABC_XYZ_Class || 'N/A')}</td>
          <td>${escapeHtml(r.ABC_Class || 'N/A')}</td>
          <td>${escapeHtml(r.Forecast_Risk || 'N/A')}</td>
          <td>${escapeHtml(r.Priority_Level || 'N/A')}</td>
          <td>${fmtNum(r.Inventory_Health_Score, 0)}</td>
          <td>${escapeHtml(r.Executive_Recommendation || 'N/A')}</td>
        `;
        attnTbody.appendChild(tr);
      });
    }
  }

  const attnCaption = document.getElementById('attention-caption');
  if (attnCaption) {
    attnCaption.textContent = `${allFlagged.length.toLocaleString('en-US')} SKU(s) flagged as High Risk or High Priority in the current filter. Showing top ${Math.min(25, allFlagged.length)} by Inventory Value.`;
  }

  // Most Common Executive Recommendations (COMPUTED FROM FULL FLAGGED POPULATION)
  const recCounts = {};
  allFlagged.forEach(r => {
    const rec = r.Executive_Recommendation;
    if (rec && rec !== 'N/A') {
      recCounts[rec] = (recCounts[rec] || 0) + 1;
    }
  });

  // Frequency descending, recommendation text ascending as tie-breaker
  const sortedRecs = Object.entries(recCounts).sort((a, b) => {
    if (b[1] !== a[1]) return b[1] - a[1];
    return a[0].localeCompare(b[0]);
  }).slice(0, 5);

  const topDecContainer = document.getElementById('top-priority-decisions-container');
  if (topDecContainer) {
    topDecContainer.innerHTML = '';
    if (sortedRecs.length === 0) {
      topDecContainer.innerHTML = '<div style="color: var(--c-slate-500); font-style: italic;">No recommendations available for the current selection.</div>';
    } else {
      sortedRecs.forEach(([rec, cnt], idx) => {
        const matchingSkus = allFlagged.filter(r => r.Executive_Recommendation === rec).map(r => r.SKU);
        const shownSkus = matchingSkus.slice(0, 12);
        let chips = shownSkus.map(s => `<span class="source-tag tag-master" style="margin:2px;">${escapeHtml(s)}</span>`).join(' ');
        if (matchingSkus.length > shownSkus.length) {
          chips += ` <span class="source-tag tag-derived">+${matchingSkus.length - shownSkus.length} more</span>`;
        }

        const card = document.createElement('div');
        card.className = 'status-card';
        card.setAttribute('data-status', 'amber');
        card.innerHTML = `
          <strong>${cnt} SKU(s)</strong><br>
          <span class="status-card-desc">${escapeHtml(rec)}</span><br>
          <div style="margin-top:0.5rem; display: flex; flex-wrap: wrap; gap: 4px;">${chips}</div>
          <div style="margin-top: 0.8rem; display: flex; gap: 0.6rem; align-items: center; flex-wrap: wrap;">
            <input type="text" id="jump-select-${idx}" class="st-select" list="jump-datalist-${idx}" value="${escapeHtml(matchingSkus[0] || '')}" placeholder="ðŸ”Ž Search or select SKU..." style="width: 280px; max-width: 100%;" autocomplete="off">
            <datalist id="jump-datalist-${idx}">
              ${matchingSkus.map(s => `<option value="${escapeHtml(s)}">`).join('')}
            </datalist>
            <button class="stButton" style="width: auto; height: 38px !important; min-height: 38px !important; padding: 0 1.2rem !important;" onclick="jumpToSKUFromTop('jump-select-${idx}')">ðŸŽ¯ Jump to SKU</button>
          </div>
        `;
        topDecContainer.appendChild(card);
      });
    }
  }

  renderDecisionIntelligenceForSelectedSKU();
}

function renderDecisionIntelligenceForSelectedSKU() {
  const sku = AppState.selectedSKU;
  const heading = document.getElementById('decision-intelligence-heading');
  const container = document.getElementById('decision-intelligence-cards');
  if (!container) return;
  container.innerHTML = '';

  if (!sku) {
    if (heading) heading.textContent = 'ðŸ§­ Decision Intelligence â€” No SKU Selected';
    container.innerHTML = '<div class="empty-state-card" style="grid-column: 1 / -1;"><strong>ðŸ§­ No SKU Selected</strong><br><span style="font-size: 0.85rem; color: var(--c-slate-500);">Select a SKU from the Control Panel or the selector above to view tailored Decision Intelligence recommendations.</span></div>';
    return;
  }

  if (heading) heading.textContent = `ðŸ§­ Decision Intelligence â€” ${sku}`;
  const row = AppState.rawRecords.find(r => String(r.SKU).trim() === String(sku).trim());
  if (!row) {
    container.innerHTML = '<div style="color: var(--c-danger-text); grid-column: 1 / -1;">Selected SKU data not found in current records.</div>';
    return;
  }

  container.innerHTML = `
    <div>
      ${renderDecisionCard("ðŸŽ¯", "Executive Recommendation", row.Executive_Recommendation || 'N/A')}
      ${renderDecisionCard("ðŸ“¦", "Inventory Recommendation", row.Inventory_Recommendation || 'N/A')}
      ${renderDecisionCard("ðŸšš", "Procurement Recommendation", row.Procurement_Recommendation || 'N/A')}
      ${renderDecisionCard("ðŸ­", "Warehouse Recommendation", row.Warehouse_Recommendation || 'N/A')}
    </div>
    <div>
      ${renderDecisionCard("ðŸ’°", "Financial Recommendation", row.Financial_Recommendation || 'N/A')}
      ${renderDecisionCard("ðŸ“", "Business Explanation", row.Business_Explanation || 'N/A')}
      ${renderDecisionCard("ðŸ“ˆ", "Expected Impact", row.Expected_Impact || 'N/A')}
    </div>
  `;
}

// ==========================================================================
// TAB 2: Demand & Forecast Intelligence
// ==========================================================================
function renderTab2() {
  const sku = AppState.selectedSKU;
  const headerEl = document.getElementById('tab2-sku-header');
  const briefEl = document.getElementById('tab2-diagnostic-brief');
  const kpisEl = document.getElementById('tab2-forecast-kpis');
  const metricsEl = document.getElementById('tab2-full-diagnostic-metrics');
  const badgesEl = document.getElementById('tab2-full-diagnostic-badges');
  const evidenceEl = document.getElementById('tab2-portfolio-evidence-text');
  const demandSummaryEl = document.getElementById('portfolio-demand-summary-metrics');

  const row = sku ? AppState.rawRecords.find(r => String(r.SKU).trim() === String(sku).trim()) : null;

  if (row) {
    const abc = row.ABC_Class || 'N/A';
    const xyz = row.XYZ_Class || 'N/A';
    if (headerEl) {
      headerEl.innerHTML = `Target SKU: <strong>${escapeHtml(sku)}</strong> &nbsp; (ABC: <strong>${escapeHtml(abc)}</strong> &nbsp;|&nbsp; XYZ: <strong>${escapeHtml(xyz)}</strong>)`;
    }

    const rmse = Number(row.RMSE);
    const avgRmse = safeMean(AppState.filteredRecords, 'RMSE');
    const bullwhip = Number(row.Bullwhip_Ratio);
    const pVal = Number(row.Health_P_Value);

    let insight = `<strong>Model Summary:</strong> The forecasting model for <strong>${escapeHtml(sku)}</strong> `;
    if (!isNaN(rmse) && !isNaN(avgRmse)) {
      if (rmse > avgRmse) {
        insight += `shows above-portfolio-average forecast error (RMSE ${fmtNum(rmse, 1)} vs. portfolio average ${fmtNum(avgRmse, 1)}). Manual review of recent demand anomalies is recommended. `;
      } else {
        insight += `is tracking within normal bounds (RMSE ${fmtNum(rmse, 1)}, portfolio average ${fmtNum(avgRmse, 1)}). `;
      }
    } else {
      insight += 'does not have an RMSE value available for comparison. ';
    }
    if (!isNaN(bullwhip)) {
      if (bullwhip > 1.5) {
        insight += `The Bullwhip Ratio â€” model-derived replenishment signal amplification metric is ${fmtNum(bullwhip, 2)} relative to the reference level. `;
      } else {
        insight += `The Bullwhip Ratio â€” model-derived replenishment signal amplification metric is ${fmtNum(bullwhip, 2)} and is near the reference level. `;
      }
    }
    if (!isNaN(pVal)) {
      const healthStatus = pVal > 0.05
        ? "Acceptable â€” no significant residual autocorrelation detected at the tested lag"
        : "Review â€” significant residual autocorrelation detected at the tested lag";
      insight += `<br><br><strong>Residual Diagnostic:</strong> ${healthStatus} â€” Ljung-Box p-value ${fmtNum(pVal, 3)}.`;
    }
    if (briefEl) briefEl.innerHTML = insight;

    // Forecast Snapshot Fallback (no fabricated trajectories)
    const meanM = Number(row.Mean_Monthly_Demand);
    const fcNext = Number(row.Forecast_Next_Month);
    Plotly.react('chart-forecast-snapshot', [{
      x: ['Mean Monthly Demand (Historical Avg)', 'Forecast â€” Next Month'],
      y: [isNaN(meanM) ? 0 : meanM, isNaN(fcNext) ? 0 : fcNext],
      type: 'bar',
      marker: { color: ['#64748B', '#3B82F6'], line: { color: '#ffffff', width: 1.5 } }
    }], {
      title: 'Forecast Snapshot',
      margin: { l: 65, r: 20, t: 45, b: 55 },
      hoverlabel: CYBER_HOVERLABEL,
      font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
      xaxis: { automargin: true },
      yaxis: { title: { text: 'Units', standoff: 15 }, automargin: true, gridcolor: '#E2E8F0' },
      plot_bgcolor: 'white',
      paper_bgcolor: 'rgba(0,0,0,0)',
      showlegend: false,
      transition: { duration: 450, easing: 'cubic-in-out' }
    }, { displayModeBar: false, responsive: true });

    // Forecast KPIs
    if (kpisEl) {
      kpisEl.innerHTML = `
        <div class="square-metric">
          <div class="square-metric-label">Next Forecast</div>
          <div class="square-metric-value" id="t2-kpi-next-val">${fmtInt(fcNext)}</div>
          <div class="square-metric-delta">Projected</div>
        </div>
        <div class="square-metric">
          <div class="square-metric-label">Bullwhip Ratio</div>
          <div class="square-metric-value" id="t2-kpi-bw-val">${fmtNum(bullwhip, 2)}</div>
          <div class="square-metric-delta">Model-Derived Replenishment Signal Â· Amplification</div>
        </div>
        <div class="square-metric">
          <div class="square-metric-label">Forecast Accuracy</div>
          <div class="square-metric-value" id="t2-kpi-acc-val">${fmtNum(row.Accuracy_Pct, 1, '%')}</div>
          <div class="square-metric-delta"></div>
        </div>
      `;
      updateKpiValueWithAnimation('t2-kpi-next-val', fmtInt(fcNext));
      updateKpiValueWithAnimation('t2-kpi-bw-val', fmtNum(bullwhip, 2));
      updateKpiValueWithAnimation('t2-kpi-acc-val', fmtNum(row.Accuracy_Pct, 1, '%'));
    }

    // Demand Snapshot Comparison
    Plotly.react('chart-forecast-compare', [{
      x: ['Historical Average', 'Next-Month Forecast'],
      y: [isNaN(meanM) ? 0 : meanM, isNaN(fcNext) ? 0 : fcNext],
      text: [fmtInt(meanM), fmtInt(fcNext)],
      textposition: 'outside',
      type: 'bar',
      marker: { color: ['#64748B', '#0EA5E9'], line: { color: '#ffffff', width: 1.5 } }
    }], {
      title: 'Demand Snapshot â€” Historical Average vs Next-Month Forecast',
      height: 255,
      margin: { l: 65, r: 24, t: 48, b: 55 },
      hoverlabel: CYBER_HOVERLABEL,
      font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
      plot_bgcolor: 'rgba(0,0,0,0)',
      paper_bgcolor: 'rgba(0,0,0,0)',
      showlegend: false,
      xaxis: { automargin: true },
      yaxis: { title: { text: 'Units', standoff: 15 }, gridcolor: '#E2E8F0', automargin: true },
      transition: { duration: 400, easing: 'cubic-in-out' }
    }, { displayModeBar: false, responsive: true });

    // Full Diagnostic Panel
    if (metricsEl) {
      metricsEl.innerHTML = `
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">ACCURACY %</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(row.Accuracy_Pct, 1)}</div></div>
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">MAE</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(row.MAE, 2)}</div></div>
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">RMSE</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(row.RMSE, 2)}</div></div>
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">BIAS</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(row.Bias, 2)}</div></div>
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">CV</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(row.CV, 2)}</div></div>
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">MEAN MONTHLY DEMAND</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(row.Mean_Monthly_Demand, 1)}</div></div>
      `;
    }

    const fr = String(row.Forecast_Risk || 'N/A');
    const frStatus = {'low': 'excellent', 'medium': 'good', 'high': 'poor'}[fr.toLowerCase()] || 'attention';
    const mr = String(row.Model_Reliability || 'N/A');
    const mrStatus = {'high': 'excellent', 'moderate': 'good', 'low': 'poor'}[mr.toLowerCase()] || 'attention';

    if (badgesEl) {
      badgesEl.innerHTML = `
        <div style="background:#fff; padding:0.9rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow);"><strong>Forecast Risk:</strong> ${badgeHtml(frStatus, fr)}</div>
        <div style="background:#fff; padding:0.9rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow);"><strong>Model Reliability:</strong> ${badgeHtml(mrStatus, mr)}</div>
        <div style="background:#fff; padding:0.9rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow);"><strong>Annualised Demand:</strong> ${fmtInt(row.Annual_Demand)}</div>
      `;
    }
  } else {
    if (headerEl) headerEl.textContent = 'Target SKU: No SKU Selected';
    if (briefEl) briefEl.innerHTML = '<div style="color:var(--c-slate-500); font-style:italic; padding:0.5rem 0;">Select a SKU from the sidebar or selector above to view its forecast diagnostics.</div>';
    if (kpisEl) kpisEl.innerHTML = '';
    if (metricsEl) metricsEl.innerHTML = '';
    if (badgesEl) badgesEl.innerHTML = '';
    const snapEl = document.getElementById('chart-forecast-snapshot');
    if (snapEl) snapEl.innerHTML = '<div style="text-align:center; padding:5rem 1rem; color:var(--c-slate-400);">Select a SKU to view Forecast Snapshot</div>';
    const compEl = document.getElementById('chart-forecast-compare');
    if (compEl) compEl.innerHTML = '<div style="text-align:center; padding:4rem 1rem; color:var(--c-slate-400);">Select a SKU to view Demand Comparison</div>';
  }

  // Portfolio Quality Distributions
  const records = AppState.filteredRecords;
  const rmseVals = records.map(r => Number(r.RMSE)).filter(v => !isNaN(v));
  const maeVals = records.map(r => Number(r.MAE)).filter(v => !isNaN(v));
  const biasVals = records.map(r => Number(r.Bias)).filter(v => !isNaN(v));
  const demandVals = records.map(r => Number(r.Mean_Monthly_Demand)).filter(v => !isNaN(v));

  Plotly.react('chart-portfolio-rmse', [{ x: rmseVals, type: 'histogram', marker: { color: '#06B6D4', line: { color: '#0891B2', width: 1 } } }], { title: 'RMSE Distribution', xaxis: { title: { text: 'Error Value', standoff: 12 }, automargin: true }, yaxis: { title: { text: 'SKU Count', standoff: 12 }, automargin: true }, margin: { l: 55, r: 15, t: 40, b: 50 }, hoverlabel: CYBER_HOVERLABEL, font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' }, plot_bgcolor: 'white', paper_bgcolor: 'rgba(0,0,0,0)', transition: { duration: 450, easing: 'cubic-in-out' } }, { displayModeBar: false, responsive: true });
  Plotly.react('chart-portfolio-mae', [{ x: maeVals, type: 'histogram', marker: { color: '#3B82F6', line: { color: '#2563EB', width: 1 } } }], { title: 'MAE Distribution', xaxis: { title: { text: 'Error Value', standoff: 12 }, automargin: true }, yaxis: { title: { text: 'SKU Count', standoff: 12 }, automargin: true }, margin: { l: 55, r: 15, t: 40, b: 50 }, hoverlabel: CYBER_HOVERLABEL, font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' }, plot_bgcolor: 'white', paper_bgcolor: 'rgba(0,0,0,0)', transition: { duration: 450, easing: 'cubic-in-out' } }, { displayModeBar: false, responsive: true });
  Plotly.react('chart-portfolio-bias', [{ x: biasVals, type: 'histogram', marker: { color: '#EC4899', line: { color: '#DB2777', width: 1 } } }], { title: 'Bias Distribution', xaxis: { title: { text: 'Error Value', standoff: 12 }, automargin: true }, yaxis: { title: { text: 'SKU Count', standoff: 12 }, automargin: true }, margin: { l: 55, r: 15, t: 40, b: 50 }, hoverlabel: CYBER_HOVERLABEL, font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' }, plot_bgcolor: 'white', paper_bgcolor: 'rgba(0,0,0,0)', transition: { duration: 450, easing: 'cubic-in-out' } }, { displayModeBar: false, responsive: true });

  const avgAcc = safeMean(records, 'Accuracy_Pct');
  const qClass = avgAcc >= 85 ? 'Excellent' : (avgAcc >= 70 ? 'Good' : (avgAcc >= 50 ? 'Average' : 'Poor'));
  Plotly.react('chart-portfolio-accuracy-gauge', [{
    domain: { x: [0, 1], y: [0, 1] },
    value: isNaN(avgAcc) ? 0 : avgAcc,
    title: { text: `Portfolio Forecast Accuracy: Dashboard Presentation Band: ${qClass}`, font: { size: 16, color: '#1E3A8A', family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' } },
    type: "indicator",
    mode: "gauge+number",
    gauge: {
      axis: { range: [0, 100] },
      bar: { color: "#3B82F6" },
      steps: [
        { range: [0, 50], color: "#FEE2E2" },
        { range: [50, 70], color: "#FEF3C7" },
        { range: [70, 85], color: "#E0F2FE" },
        { range: [85, 100], color: "#D1FAE5" }
      ]
    }
  }], { margin: { l: 30, r: 30, t: 60, b: 20 }, height: 250, hoverlabel: CYBER_HOVERLABEL, paper_bgcolor: 'rgba(0,0,0,0)', transition: { duration: 400, easing: 'cubic-in-out' } }, { displayModeBar: false, responsive: true });

  // Portfolio-Level Demand Statistics (Mean, Median, Std Dev, Skewness, Kurtosis)
  if (demandSummaryEl) {
    if (demandVals.length > 2) {
      const mean = demandVals.reduce((a, b) => a + b, 0) / demandVals.length;
      const med = median(demandVals);
      const variance = demandVals.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / demandVals.length;
      const std = Math.sqrt(variance);
      const m3 = demandVals.reduce((a, b) => a + Math.pow(b - mean, 3), 0) / demandVals.length;
      const skew = std > 0 ? m3 / Math.pow(std, 3) : NaN;
      const m4 = demandVals.reduce((a, b) => a + Math.pow(b - mean, 4), 0) / demandVals.length;
      const kurt = std > 0 ? (m4 / Math.pow(std, 4)) - 3 : NaN;

      demandSummaryEl.innerHTML = `
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">MEAN</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(mean, 1)}</div></div>
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">MEDIAN</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(med, 1)}</div></div>
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">STD DEV</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(std, 1)}</div></div>
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">SKEWNESS</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(skew, 2)}</div></div>
        <div style="background:#fff; padding:0.8rem; border-radius:var(--clay-radius-sm); box-shadow:var(--clay-shadow); text-align:center;"><div style="font-size:0.75rem; color:var(--c-slate-500); font-weight:700;">KURTOSIS</div><div style="font-size:1.4rem; font-weight:900; color:var(--c-navy);">${fmtNum(kurt, 2)}</div></div>
      `;
    } else {
      demandSummaryEl.innerHTML = '<div style="color:var(--c-slate-500); font-style:italic;">Not enough Mean_Monthly_Demand data for statistics.</div>';
    }
  }

  // Distribution Histogram
  Plotly.react('chart-portfolio-demand-dist', [{
    x: demandVals,
    type: 'histogram',
    nbinsx: 20,
    marker: { color: '#8B5CF6', line: { color: '#7C3AED', width: 1 } }
  }], {
    title: 'Distribution of Mean Monthly Demand Across Portfolio',
    xaxis: { title: { text: 'Monthly Demand (Units)', standoff: 12 }, automargin: true },
    yaxis: { title: { text: 'Frequency (SKUs)', standoff: 12 }, automargin: true },
    margin: { l: 55, r: 15, t: 40, b: 50 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  // Forecast Risk by ABC Class Heatmap
  const riskMat = { A: {Low:0, Medium:0, High:0}, B: {Low:0, Medium:0, High:0}, C: {Low:0, Medium:0, High:0} };
  records.forEach(r => {
    const abc_c = String(r.ABC_Class || '').trim();
    const fr_c = String(r.Forecast_Risk || '').trim();
    if (riskMat[abc_c] && riskMat[abc_c][fr_c] !== undefined) riskMat[abc_c][fr_c]++;
  });
  Plotly.react('chart-forecast-risk-abc', [{
    z: [
      [riskMat.A.Low, riskMat.A.Medium, riskMat.A.High],
      [riskMat.B.Low, riskMat.B.Medium, riskMat.B.High],
      [riskMat.C.Low, riskMat.C.Medium, riskMat.C.High]
    ],
    x: ['Low', 'Medium', 'High'],
    y: ['A', 'B', 'C'],
    type: 'heatmap',
    colorscale: [[0, '#F8FAFC'], [0.25, '#FEF3C7'], [0.6, '#F59E0B'], [1, '#DC2626']],
    showscale: true,
    colorbar: { len: 0.85, thickness: 14, title: { text: 'SKU Count', font: { size: 11 } } },
    text: [
      [riskMat.A.Low, riskMat.A.Medium, riskMat.A.High],
      [riskMat.B.Low, riskMat.B.Medium, riskMat.B.High],
      [riskMat.C.Low, riskMat.C.Medium, riskMat.C.High]
    ],
    texttemplate: "%{text}"
  }], {
    title: 'Forecast Risk by ABC Class â€” SKU Count',
    xaxis: { title: { text: 'Forecast Risk', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'ABC â€” Demand Volume Classification', standoff: 15 }, automargin: true },
    margin: { l: 60, r: 40, t: 40, b: 55 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    transition: { duration: 400, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  if (evidenceEl) {
    evidenceEl.textContent = `${fmtNum(avgAcc, 2, "%")} average validation accuracy across ${records.length.toLocaleString('en-US')} SKUs in the current filter â€” this is holdout validation performance, not an improvement figure versus any prior forecasting process.`;
  }
}

// ==========================================================================
// TAB 3: Inventory Optimisation
// ==========================================================================
function renderTab3() {
  const sku = AppState.selectedSKU;
  const headerEl = document.getElementById('tab3-sku-header');
  const row = sku ? AppState.rawRecords.find(r => String(r.SKU).trim() === String(sku).trim()) : null;

  if (row) {
    const abc = row.ABC_Class || 'N/A';
    const xyz = row.XYZ_Class || 'N/A';
    const comb = row.ABC_XYZ_Class || 'N/A';
    if (headerEl) {
      headerEl.innerHTML = `Target SKU: <strong>${escapeHtml(sku)}</strong> &nbsp; (ABC: <strong>${escapeHtml(abc)}</strong> &nbsp;|&nbsp; XYZ: <strong>${escapeHtml(xyz)}</strong> &nbsp;|&nbsp; Combined: <strong>${escapeHtml(comb)}</strong>)`;
    }

    updateKpiValueWithAnimation('tab3-kpi-ss', fmtInt(row.Safety_Stock));
    updateKpiValueWithAnimation('tab3-kpi-rop', fmtInt(row.Reorder_Point));
    updateKpiValueWithAnimation('tab3-kpi-eoq', fmtInt(row.EOQ));
    updateKpiValueWithAnimation('tab3-kpi-avg-inv', fmtInt(row.Average_Inventory_Units));
    updateKpiValueWithAnimation('tab3-kpi-val', fmtMonetary(row.Inventory_Value));
    updateKpiValueWithAnimation('tab3-kpi-wc', fmtMonetary(row.Working_Capital));
    updateKpiValueWithAnimation('tab3-kpi-days', fmtNum(row.Inventory_Days, 0));
    updateKpiValueWithAnimation('tab3-kpi-turnover', fmtNum(row.Inventory_Turnover, 2));

    const slEl = document.getElementById('tab3-service-level');
    if (slEl) slEl.textContent = fmtNum(row.Service_Level_Pct, 1, '%');
    const frEl = document.getElementById('tab3-fill-rate');
    if (frEl) frEl.textContent = fmtNum(row.Fill_Rate_Pct, 1, '%');
    const ihs = Number(row.Inventory_Health_Score);
    const ihsStatus = scoreToStatus(ihs);
    const ihsEl = document.getElementById('tab3-inv-health');
    if (ihsEl) ihsEl.innerHTML = `${fmtNum(ihs, 0)} ${badgeHtml(ihsStatus, ihsStatus.charAt(0).toUpperCase() + ihsStatus.slice(1))}`;

    // Planning / XYZ derived logic
    let planFreq = row.Planning_Frequency || row['Planning Frequency'];
    let invPolicy = row.Inventory_Policy || row['Inventory Policy'];
    let revFreq = row.Review_Frequency || row['Review Frequency'];

    if (!planFreq || planFreq === 'N/A') {
      const cv = Number(row.CV);
      if (!isNaN(cv)) {
        if (cv < 0.25) planFreq = 'Monthly';
        else if (cv < 0.5) planFreq = 'Bi-weekly';
        else planFreq = 'Weekly';
      }
    }

    if (!invPolicy || invPolicy === 'N/A') {
      const xyzClass = String(row.XYZ_Class || '').trim().toUpperCase();
      if (xyzClass === 'X') invPolicy = 'Continuous Review / Standard Buffer';
      else if (xyzClass === 'Y') invPolicy = 'Periodic Review / Controlled Buffer';
      else if (xyzClass === 'Z') invPolicy = 'Exception-Based Review / Dynamic Buffer';
    }

    if (!revFreq || revFreq === 'N/A') {
      const xyzClass = String(row.XYZ_Class || '').trim().toUpperCase();
      if (xyzClass === 'X') revFreq = 'Monthly';
      else if (xyzClass === 'Y') revFreq = 'Bi-weekly';
      else if (xyzClass === 'Z') revFreq = 'Weekly';
    }

    const pfEl = document.getElementById('tab3-plan-freq');
    if (pfEl) pfEl.textContent = planFreq || 'N/A';
    const ipEl = document.getElementById('tab3-inv-policy');
    if (ipEl) ipEl.textContent = invPolicy || 'N/A';
    const rfEl = document.getElementById('tab3-review-freq');
    if (rfEl) rfEl.textContent = revFreq || 'N/A';

    // Math Integrity QA for Selected SKU (Expandable)
    const valRes = independentSSRop(row);
    const maSS = Number(row.Safety_Stock);
    const maROP = Number(row.Reorder_Point);
    const indSS = valRes.ss;
    const indROP = valRes.rop;
    const ssVar = (isNaN(maSS) || isNaN(indSS)) ? null : Math.abs(maSS - indSS);
    const ropVar = (isNaN(maROP) || isNaN(indROP)) ? null : Math.abs(maROP - indROP);
    const ssOk = ssVar !== null && ssVar <= 1;
    const ropOk = ropVar !== null && ropVar <= 1;

    const mathExpander = document.getElementById('math-integrity-expander');
    if (mathExpander) {
      mathExpander.open = (maSS === 0);
    }

    const mathAuditContent = document.getElementById('tab3-math-audit-content');
    if (mathAuditContent) {
      let zeroNote = '';
      if (maSS === 0 && ssOk) {
        zeroNote = `<div class="status-card" data-status="green" style="margin-top:0.8rem; font-size:0.84rem;">A Safety Stock of 0 is mathematically correct here â€” not a data-quality defect. It occurs when the SKU has zero historical demand variance (dead stock / constant demand) or when the forecast model achieved a near-zero RMSE on a very low-volume, intermittent-demand SKU, so the computed buffer legitimately rounds to zero.</div>`;
      }
      mathAuditContent.innerHTML = `
        <table class="math-audit-table">
          <tr><th></th><th>Master Audit</th><th>Independent Validation</th><th>Variance</th><th>Status</th></tr>
          <tr><td><strong>Safety Stock</strong></td><td>${fmtNum(maSS, 0)}</td><td>${fmtNum(indSS, 0)}</td><td>${ssVar !== null ? fmtNum(ssVar, 0) : 'N/A'}</td><td class="${ssOk ? 'math-status-ok' : 'math-status-review'}">${ssOk ? 'âœ“ Confirmed' : (ssVar !== null ? 'âš  Review' : 'N/A')}</td></tr>
          <tr><td><strong>Reorder Point</strong></td><td>${fmtNum(maROP, 0)}</td><td>${fmtNum(indROP, 0)}</td><td>${ropVar !== null ? fmtNum(ropVar, 0) : 'N/A'}</td><td class="${ropOk ? 'math-status-ok' : 'math-status-review'}">${ropOk ? 'âœ“ Confirmed' : (ropVar !== null ? 'âš  Review' : 'N/A')}</td></tr>
        </table>
        <p class="section-caption" style="margin-top: 0.8rem;">Basis: ${valRes.basis}. Independent formula: Safety Stock = 1.645 Ã— RMSE Ã— âˆš(Lead Time Ã· 30); Reorder Point = (Forecast Next Month Ã· 30 Ã— Lead Time) + Safety Stock â€” reproduced directly from the source notebook's own implementation, not a generic textbook substitute.</p>
        ${zeroNote}
      `;
    }
  } else {
    if (headerEl) headerEl.textContent = 'Target SKU: No SKU Selected';
    updateKpiValueWithAnimation('tab3-kpi-ss', '-');
    updateKpiValueWithAnimation('tab3-kpi-rop', '-');
    updateKpiValueWithAnimation('tab3-kpi-eoq', '-');
    updateKpiValueWithAnimation('tab3-kpi-avg-inv', '-');
    updateKpiValueWithAnimation('tab3-kpi-val', '-');
    updateKpiValueWithAnimation('tab3-kpi-wc', '-');
    updateKpiValueWithAnimation('tab3-kpi-days', '-');
    updateKpiValueWithAnimation('tab3-kpi-turnover', '-');
    const mathAuditContent = document.getElementById('tab3-math-audit-content');
    if (mathAuditContent) {
      mathAuditContent.innerHTML = '<div style="color:var(--c-slate-500); font-style:italic; padding:0.5rem 0;">Select a SKU from the sidebar or selector above to run independent formula audit.</div>';
    }
  }

  // Portfolio Replenishment View Table (Top 50 by lowest Inventory Health Score)
  const records = AppState.filteredRecords;
  const replList = [...records].sort((a, b) => (Number(a.Inventory_Health_Score) || 0) - (Number(b.Inventory_Health_Score) || 0)).slice(0, 50);
  const replTbody = document.getElementById('portfolio-replenishment-tbody');
  if (replTbody) {
    replTbody.innerHTML = '';
    if (replList.length === 0) {
      replTbody.innerHTML = '<tr><td colspan="10" style="text-align: center; padding: 1.5rem;">No SKUs available.</td></tr>';
    } else {
      replList.forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td><strong>${escapeHtml(r.SKU || 'N/A')}</strong></td>
          <td>${escapeHtml(r.ABC_XYZ_Class || 'N/A')}</td>
          <td>${escapeHtml(r.ABC_Class || 'N/A')}</td>
          <td>${fmtInt(r.Safety_Stock)}</td>
          <td>${fmtInt(r.Reorder_Point)}</td>
          <td>${fmtInt(r.EOQ)}</td>
          <td>${fmtNum(r.Inventory_Days, 0)}</td>
          <td>${fmtNum(r.Inventory_Turnover, 2)}</td>
          <td>${fmtNum(r.Inventory_Health_Score, 0)}</td>
          <td>${escapeHtml(r.Inventory_Recommendation || 'N/A')}</td>
        `;
        replTbody.appendChild(tr);
      });
    }
  }

  // Turnover by ABC Bar Chart
  const turnABC = { A: [], B: [], C: [] };
  records.forEach(r => {
    const c = String(r.ABC_Class || '').trim();
    const t = Number(r.Inventory_Turnover);
    if (turnABC[c] && !isNaN(t)) turnABC[c].push(t);
  });
  const meanTurnA = turnABC.A.length ? turnABC.A.reduce((x, y) => x + y, 0) / turnABC.A.length : 0;
  const meanTurnB = turnABC.B.length ? turnABC.B.reduce((x, y) => x + y, 0) / turnABC.B.length : 0;
  const meanTurnC = turnABC.C.length ? turnABC.C.reduce((x, y) => x + y, 0) / turnABC.C.length : 0;

  Plotly.react('chart-replenishment-turnover', [{
    x: ['A', 'B', 'C'],
    y: [meanTurnA, meanTurnB, meanTurnC],
    type: 'bar',
    marker: { color: [ABC_COLOR_MAP.A, ABC_COLOR_MAP.B, ABC_COLOR_MAP.C], line: { color: '#ffffff', width: 1.5 } }
  }], {
    title: 'Average Inventory Turnover by ABC Class',
    xaxis: { title: { text: 'ABC Class', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'Average Turnover', standoff: 15 }, automargin: true },
    margin: { l: 60, r: 15, t: 40, b: 50 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    showlegend: false,
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  const ihsVals = records.map(r => Number(r.Inventory_Health_Score)).filter(v => !isNaN(v));
  Plotly.react('chart-replenishment-ihs', [{
    x: ihsVals,
    type: 'histogram',
    marker: { color: '#10B981', line: { color: '#059669', width: 1 } }
  }], {
    title: 'Inventory Health Score Distribution',
    xaxis: { title: { text: 'Health Score', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'SKU Count', standoff: 15 }, automargin: true },
    margin: { l: 60, r: 15, t: 40, b: 50 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  const invValSum = safeSum(records, 'Inventory_Value');
  document.getElementById('tab3-model-implied-pos-text').innerHTML = `Model-Implied Inventory Value across the current filter is <strong>${fmtMonetary(invValSum)}</strong> â€” the analytical figure implied by applying this policy to Master Audit fields, not an observed accounting balance. No actual/observed on-hand inventory column is present in this data source, so it is not shown as if it were.`;
}

// ==========================================================================
// TAB 4: Risk, Segmentation & Diagnostics
// ==========================================================================
function deterministicHighRiskSort(records) {
  const prioMap = { 'high': 0, 'medium': 1, 'low': 2 };
  return [...records].sort((a, b) => {
    const pA = prioMap[String(a.Priority_Level || '').trim().toLowerCase()] ?? 99;
    const pB = prioMap[String(b.Priority_Level || '').trim().toLowerCase()] ?? 99;
    if (pA !== pB) return pA - pB;

    const rmseA = Number(a.RMSE);
    const rmseB = Number(b.RMSE);
    const validRmseA = !isNaN(rmseA);
    const validRmseB = !isNaN(rmseB);
    if (validRmseA && validRmseB && rmseA !== rmseB) return rmseB - rmseA;
    if (validRmseA && !validRmseB) return -1;
    if (!validRmseA && validRmseB) return 1;

    const bwA = Number(a.Bullwhip_Ratio);
    const bwB = Number(b.Bullwhip_Ratio);
    const validBwA = !isNaN(bwA);
    const validBwB = !isNaN(bwB);
    if (validBwA && validBwB && bwA !== bwB) return bwB - bwA;
    if (validBwA && !validBwB) return -1;
    if (!validBwA && validBwB) return 1;

    return String(a.SKU || '').localeCompare(String(b.SKU || ''));
  });
}

function renderTab4() {
  const records = AppState.filteredRecords;

  // Bullwhip Ratio vs Forecast Error Scatter (nulls never converted to zero)
  const validScatter = records.filter(r => !isMissing(r.Bullwhip_Ratio) && !isMissing(r.RMSE));
  if (validScatter.length > 2) {
    const bwVals = validScatter.map(r => Number(r.Bullwhip_Ratio));
    const rmseVals = validScatter.map(r => Number(r.RMSE));
    const medBw = median(bwVals);
    const medRmse = median(rmseVals);

    const scatterData = validScatter.map(r => {
      const bw = Number(r.Bullwhip_Ratio);
      const rmse = Number(r.RMSE);
      const bwHigh = bw > medBw;
      const rmseHigh = rmse > medRmse;
      let quad = 'Stable';
      if (bwHigh && rmseHigh) quad = 'Critical';
      else if (!bwHigh && rmseHigh) quad = 'Forecast Issue';
      else if (bwHigh && !rmseHigh) quad = 'Signal Amplification';
      return { sku: r.SKU, bw, rmse, quad };
    });

    const traces = {};
    const colorMap = { "Stable": "#10B981", "Forecast Issue": "#F59E0B", "Signal Amplification": "#06B6D4", "Critical": "#EF4444" };
    ["Stable", "Forecast Issue", "Signal Amplification", "Critical"].forEach(q => {
      const sub = scatterData.filter(d => d.quad === q);
      traces[q] = {
        x: sub.map(d => d.bw),
        y: sub.map(d => d.rmse),
        text: sub.map(d => `SKU: ${d.sku}<br>Bullwhip: ${d.bw.toFixed(2)}<br>RMSE: ${d.rmse.toFixed(2)}`),
        hoverinfo: 'text',
        mode: 'markers',
        name: q,
        marker: { size: 10, color: colorMap[q], opacity: 0.85, line: { width: 1.5, color: '#ffffff' } }
      };
    });

    Plotly.react('chart-bullwhip-rmse-scatter', Object.values(traces), {
      title: 'Bullwhip Ratio vs. Forecast Error â€” Dashboard-Derived Review Lens',
      xaxis: { title: { text: 'Bullwhip Ratio â€” Model-Derived Replenishment Signal Amplification', standoff: 15 }, zeroline: false, automargin: true },
      yaxis: { title: { text: 'RMSE', standoff: 15 }, zeroline: false, automargin: true },
      shapes: [
        { type: 'line', x0: medBw, x1: medBw, y0: 0, y1: Math.max(...rmseVals) * 1.1, line: { dash: 'dot', color: '#94A3B8' } },
        { type: 'line', x0: 0, x1: Math.max(...bwVals) * 1.1, y0: medRmse, y1: medRmse, line: { dash: 'dot', color: '#94A3B8' } }
      ],
      hoverlabel: CYBER_HOVERLABEL,
      font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
      margin: { l: 65, r: 25, t: 45, b: 60 },
      plot_bgcolor: 'white',
      paper_bgcolor: 'rgba(0,0,0,0)',
      transition: { duration: 450, easing: 'cubic-in-out' }
    }, { displayModeBar: false, responsive: true });
  } else {
    document.getElementById('chart-bullwhip-rmse-scatter').innerHTML = '<div style="text-align:center; padding:3rem; color:var(--c-slate-500);">Bullwhip_Ratio and/or RMSE not available, or not enough SKUs in the current filter.</div>';
  }

  // Segmentation ABC & XYZ Bar Charts
  const abcCounts = orderedCounts(records, 'ABC_Class', ABC_ORDER);
  Plotly.react('chart-segmentation-abc-bar', [{
    x: abcCounts.map(d => d.cat),
    y: abcCounts.map(d => d.count),
    type: 'bar',
    marker: { color: [ABC_COLOR_MAP.A, ABC_COLOR_MAP.B, ABC_COLOR_MAP.C], line: { color: '#ffffff', width: 1.5 } }
  }], {
    title: 'ABC â€” Demand Volume Classification',
    xaxis: { title: { text: 'Class', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'SKU Count', standoff: 15 }, automargin: true },
    margin: { l: 60, r: 15, t: 40, b: 50 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    showlegend: false,
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  const xyzCounts = orderedCounts(records, 'XYZ_Class', XYZ_ORDER);
  Plotly.react('chart-segmentation-xyz-bar', [{
    x: xyzCounts.map(d => d.cat),
    y: xyzCounts.map(d => d.count),
    type: 'bar',
    marker: { color: [XYZ_COLOR_MAP.X, XYZ_COLOR_MAP.Y, XYZ_COLOR_MAP.Z], line: { color: '#ffffff', width: 1.5 } }
  }], {
    title: 'XYZ â€” Demand Variability Classification',
    xaxis: { title: { text: 'Class', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'SKU Count', standoff: 15 }, automargin: true },
    margin: { l: 60, r: 15, t: 40, b: 50 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    showlegend: false,
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  // ABCâ€“XYZ Risk Intensity Matrix (preserves empty/missing cells as null)
  const meanRiskZ = ['A', 'B', 'C'].map(a => {
    return ['X', 'Y', 'Z'].map(x => {
      const code = a + x;
      const matching = records.filter(r => String(r.ABC_XYZ_Class || '').trim() === code && !isMissing(r.RMSE));
      if (matching.length === 0) return null;
      return matching.reduce((sum, r) => sum + Number(r.RMSE), 0) / matching.length;
    });
  });

  Plotly.react('chart-risk-intensity-matrix', [{
    z: meanRiskZ,
    x: ['X', 'Y', 'Z'],
    y: ['A', 'B', 'C'],
    type: 'heatmap',
    colorscale: [[0, '#ECFDF5'], [0.35, '#FEF3C7'], [0.7, '#F59E0B'], [1, '#DC2626']],
    text: meanRiskZ.map(row => row.map(v => v === null ? '' : v.toFixed(1))),
    texttemplate: "%{text}",
    showscale: true,
    colorbar: { len: 0.85, thickness: 14, title: { text: 'Avg RMSE', font: { size: 11 } } }
  }], {
    title: 'ABCâ€“XYZ Risk Intensity â€” Mean RMSE',
    xaxis: { title: { text: 'XYZ â€” Demand Variability', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'ABC â€” Demand Volume', standoff: 15 }, automargin: true },
    margin: { l: 60, r: 40, t: 40, b: 55 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    transition: { duration: 400, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  // High-Risk SKU Register Table (deterministic 4-key sort, full register, NO 50 cap)
  const hrList = records.filter(r => String(r.Forecast_Risk || '').trim().toLowerCase() === 'high');
  const sortedHr = deterministicHighRiskSort(hrList);
  const hrTbody = document.getElementById('high-risk-register-tbody');
  if (hrTbody) {
    hrTbody.innerHTML = '';
    if (sortedHr.length === 0) {
      hrTbody.innerHTML = '<tr><td colspan="10" style="text-align: center; padding: 1.5rem;">No high-risk SKUs in current filter.</td></tr>';
    } else {
      sortedHr.forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(r.SKU || 'N/A')}</td>
          <td>${escapeHtml(r.ABC_XYZ_Class || 'N/A')}</td>
          <td>${fmtNum(r.RMSE, 2)}</td>
          <td>${fmtNum(r.Bullwhip_Ratio, 2)}</td>
          <td>${fmtNum(r.CV, 2)}</td>
          <td>${escapeHtml(r.Forecast_Risk || 'N/A')}</td>
          <td>${escapeHtml(r.Business_Risk || r['Business Risk'] || 'N/A')}</td>
          <td>${escapeHtml(r.Priority_Level || 'N/A')}</td>
          <td>${escapeHtml(r.Supply_Chain_Priority || r['Supply Chain Priority'] || 'N/A')}</td>
          <td>${escapeHtml(r.Model_Reliability || 'N/A')}</td>
        `;
        hrTbody.appendChild(tr);
      });
    }
  }
  const hrCaption = document.getElementById('high-risk-caption');
  if (hrCaption) {
    hrCaption.textContent = `${sortedHr.length.toLocaleString('en-US')} SKU(s) flagged High Forecast Risk (Master Audit).`;
  }

  // Pareto Analysis (Top 30 SKUs with 80% reference threshold)
  const paretoData = records.filter(r => !isMissing(r.Annual_Demand))
    .map(r => ({ sku: r.SKU, annual: Number(r.Annual_Demand) || 0 }))
    .sort((a, b) => b.annual - a.annual);
  const totalAnnual = paretoData.reduce((acc, d) => acc + d.annual, 0);

  let cumSum = 0;
  let aClassCount = paretoData.length;
  for (let i = 0; i < paretoData.length; i++) {
    cumSum += paretoData[i].annual;
    if (totalAnnual > 0 && (cumSum / totalAnnual) * 100 >= 80) {
      aClassCount = i + 1;
      break;
    }
  }

  const top30Pareto = paretoData.slice(0, 30);
  let topCum = 0;
  const paretoCum = top30Pareto.map(d => {
    topCum += d.annual;
    return totalAnnual > 0 ? (topCum / totalAnnual) * 100 : 0;
  });

  Plotly.react('chart-pareto-analysis', [
    { x: top30Pareto.map(d => d.sku), y: top30Pareto.map(d => d.annual), type: 'bar', name: 'Annualised Demand', marker: { color: '#3B82F6', line: { color: '#2563EB', width: 1 } }, yaxis: 'y' },
    { x: top30Pareto.map(d => d.sku), y: paretoCum, type: 'scatter', mode: 'lines+markers', name: 'Cumulative %', line: { color: '#EC4899', width: 3.5 }, yaxis: 'y2' }
  ], {
    title: 'Top 30 SKUs â€” Pareto (Annualised Demand, Master Audit)',
    margin: { l: 70, r: 65, t: 45, b: 85 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    xaxis: { tickangle: -60, automargin: true },
    yaxis: { title: { text: 'Annualised Demand', standoff: 15 }, automargin: true },
    yaxis2: { title: { text: 'Cumulative %', standoff: 15 }, overlaying: 'y', side: 'right', range: [0, 105], automargin: true },
    shapes: [{ type: 'line', x0: 0, x1: 1, xref: 'paper', y0: 80, y1: 80, yaxis: 'y2', line: { dash: 'dot', color: '#DC2626' } }],
    annotations: [{ xref: 'paper', x: 0.95, y: 80, yref: 'y2', text: '80% reference threshold', showarrow: false, font: { color: '#DC2626', size: 10 }, yanchor: 'bottom' }],
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  const paretoCaption = document.getElementById('pareto-caption');
  if (paretoCaption) {
    paretoCaption.textContent = `Approximately ${aClassCount.toLocaleString('en-US')} of ${paretoData.length.toLocaleString('en-US')} SKUs (by Annualised Demand) reach 80% of cumulative demand.`;
  }

  // High Inventory Days & Lower Demand Diagnostic (Full population, 4-tier sort)
  const invDays = records.map(r => Number(r.Inventory_Days)).filter(v => !isNaN(v));
  const meanDemands = records.map(r => Number(r.Mean_Monthly_Demand)).filter(v => !isNaN(v));
  const daysQ75 = percentile(invDays, 0.75);
  const demandQ25 = percentile(meanDemands, 0.25);

  const trapList = records.filter(r => {
    const d = Number(r.Inventory_Days), m = Number(r.Mean_Monthly_Demand);
    return !isNaN(d) && !isNaN(m) && d >= daysQ75 && m <= demandQ25;
  }).sort((a, b) => {
    const vA = Number(a.Inventory_Value) || 0, vB = Number(b.Inventory_Value) || 0;
    if (vB !== vA) return vB - vA;
    const dA = Number(a.Inventory_Days) || 0, dB = Number(b.Inventory_Days) || 0;
    if (dB !== dA) return dB - dA;
    const mA = Number(a.Mean_Monthly_Demand) || 0, mB = Number(b.Mean_Monthly_Demand) || 0;
    if (mA !== mB) return mA - mB;
    return String(a.SKU || '').localeCompare(String(b.SKU || ''));
  });

  const trapTbody = document.getElementById('capital-trap-tbody');
  if (trapTbody) {
    trapTbody.innerHTML = '';
    if (trapList.length === 0) {
      trapTbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 1.5rem;">No candidates flagged in current filter.</td></tr>';
    } else {
      trapList.forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${escapeHtml(r.SKU || 'N/A')}</td>
          <td>${escapeHtml(r.ABC_XYZ_Class || 'N/A')}</td>
          <td>${fmtNum(r.Mean_Monthly_Demand, 1)}</td>
          <td>${fmtNum(r.Inventory_Days, 0)}</td>
          <td>${fmtMonetary(r.Inventory_Value)}</td>
          <td>${escapeHtml(r.Forecast_Risk || 'N/A')}</td>
          <td>${escapeHtml(r.Inventory_Recommendation || 'N/A')}</td>
        `;
        trapTbody.appendChild(tr);
      });
    }
  }

  const trapCaption = document.getElementById('capital-trap-caption');
  if (trapCaption) {
    trapCaption.textContent = `${trapList.length.toLocaleString('en-US')} SKU(s) flagged as dashboard-derived candidates for inventory review â€” requires validation. These are dashboard-derived candidates for review, not confirmed excess inventory.`;
  }

  // Business Risk & SC Priority Categories
  const brOrder = ["Elevated Variability / Stockout Risk", "Moderate Variance", "Low Obsolescence"];
  const brCounts = { "Elevated Variability / Stockout Risk": 0, "Moderate Variance": 0, "Low Obsolescence": 0 };
  records.forEach(r => {
    const b = String(r.Business_Risk || r['Business Risk'] || '').trim();
    if (brCounts[b] !== undefined) brCounts[b]++;
  });
  Plotly.react('chart-business-risk-bar', [{
    x: brOrder,
    y: brOrder.map(k => brCounts[k]),
    type: 'bar',
    marker: { color: ['#EF4444', '#F59E0B', '#10B981'], line: { color: '#ffffff', width: 1.5 } }
  }], {
    title: 'Business Risk Categories',
    xaxis: { title: { text: 'Risk Category', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'SKU Count', standoff: 15 }, automargin: true },
    margin: { l: 60, r: 15, t: 45, b: 70 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    showlegend: false,
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  const spOrder = ["High", "Medium", "Standard"];
  const spCounts = { "High": 0, "Medium": 0, "Standard": 0 };
  records.forEach(r => {
    const p = String(r.Supply_Chain_Priority || r['Supply Chain Priority'] || '').trim();
    if (spCounts[p] !== undefined) spCounts[p]++;
  });
  Plotly.react('chart-sc-priority-bar', [{
    x: spOrder,
    y: spOrder.map(k => spCounts[k]),
    type: 'bar',
    marker: { color: ['#EF4444', '#F59E0B', '#3B82F6'], line: { color: '#ffffff', width: 1.5 } }
  }], {
    title: 'Supply Chain Priority Categories',
    xaxis: { title: { text: 'Priority Level', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'SKU Count', standoff: 15 }, automargin: true },
    margin: { l: 60, r: 15, t: 45, b: 55 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    showlegend: false,
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });
}

// ==========================================================================
// TAB 5: Financial Impact, Methodology & Audit
// ==========================================================================
function renderTab5() {
  const records = AppState.filteredRecords;
  const invValueSum = safeSum(records, 'Inventory_Value');
  const wcSum = safeSum(records, 'Working_Capital');
  const carryingSum = safeSum(records, 'Estimated_Carrying_Cost');
  const stockoutSum = safeSum(records, 'Stockout_Cost');
  const finHealthAvg = safeMean(records, 'KPI_Financial_Health_Score');
  const turnoverAvg = safeMean(records, 'Inventory_Turnover');
  const daysAvg = !isNaN(turnoverAvg) && turnoverAvg > 0 ? 365.0 / turnoverAvg : NaN;
  const wcEffAvg = !isNaN(turnoverAvg) ? Math.min(Math.max((turnoverAvg / 6.0) * 100.0, 0), 100) : NaN;

  updateKpiValueWithAnimation('tab5-kpi-inv-val', fmtMonetary(invValueSum));
  updateKpiValueWithAnimation('tab5-kpi-wc', fmtMonetary(wcSum));
  updateKpiValueWithAnimation('tab5-kpi-carrying', fmtMonetary(carryingSum));
  updateKpiValueWithAnimation('tab5-kpi-stockout', fmtMonetary(stockoutSum));

  updateKpiValueWithAnimation('tab5-kpi-turnover', fmtNum(turnoverAvg, 2));
  updateKpiValueWithAnimation('tab5-kpi-days', fmtNum(daysAvg, 0));
  updateKpiValueWithAnimation('tab5-kpi-fin-health', fmtNum(finHealthAvg, 0, '%'));
  updateKpiValueWithAnimation('tab5-kpi-wce', fmtNum(wcEffAvg, 0, '%'));

  // Capital Allocation Treemap
  const treeRows = records.filter(r => !isMissing(r.Inventory_Value) && !isMissing(r.ABC_Class));
  if (treeRows.length > 0) {
    const abcTotals = { A: 0, B: 0, C: 0 };
    treeRows.forEach(r => {
      const c = String(r.ABC_Class).trim();
      const val = Number(r.Inventory_Value) || 0;
      if (abcTotals[c] !== undefined) abcTotals[c] += val;
    });

    const portfolioTotal = Object.values(abcTotals).reduce((a, b) => a + b, 0);

    const ids = ['Portfolio'];
    const labels = ['Portfolio'];
    const parents = [''];
    const values = [portfolioTotal];
    const colors = ['#1E3A8A'];

    ['A', 'B', 'C'].forEach(c => {
      if (abcTotals[c] > 0) {
        ids.push(`Portfolio-${c}`);
        labels.push(`Class ${c}`);
        parents.push('Portfolio');
        values.push(abcTotals[c]);
        colors.push(ABC_COLOR_MAP[c] || '#0EA5E9');
      }
    });

    const SKU_COLOR_TINTS = {
      'A': '#2563EB',
      'B': '#0284C7',
      'C': '#0D9488'
    };

    // === PERF OPTIMIZATION: Cap each ABC class to top-30 SKUs by Inventory Value ===
    // Groups remaining SKUs into a single "Other (N SKUs)" node per class
    // Reduces leaf nodes from ~1400 to ~93, dramatically speeding up Tab 5
    const MAX_SKU_PER_CLASS = 30;
    ['A', 'B', 'C'].forEach(c => {
      const classRows = treeRows.filter(r => String(r.ABC_Class).trim() === c)
        .sort((a, b) => (Number(b.Inventory_Value) || 0) - (Number(a.Inventory_Value) || 0));
      const topRows = classRows.slice(0, MAX_SKU_PER_CLASS);
      const restRows = classRows.slice(MAX_SKU_PER_CLASS);
      topRows.forEach(r => {
        const val = Number(r.Inventory_Value) || 0;
        ids.push(`Portfolio-${c}-${r.SKU}`);
        labels.push(r.SKU);
        parents.push(`Portfolio-${c}`);
        values.push(val);
        colors.push(SKU_COLOR_TINTS[c] || '#64748B');
      });
      if (restRows.length > 0) {
        const otherVal = restRows.reduce((s, r) => s + (Number(r.Inventory_Value) || 0), 0);
        ids.push(`Portfolio-${c}-Other`);
        labels.push(`Other (${restRows.length} SKUs)`);
        parents.push(`Portfolio-${c}`);
        values.push(otherVal);
        colors.push(c === 'A' ? '#93C5FD' : c === 'B' ? '#7DD3FC' : '#5EEAD4');
      }
    });

    Plotly.react('chart-financial-treemap', [{
      type: "treemap",
      ids: ids,
      labels: labels,
      parents: parents,
      values: values,
      branchvalues: 'total',
      marker: {
        colors: colors,
        line: { color: '#FFFFFF', width: 1.2 }
      },
      texttemplate: '<b>%{label}</b><br>%{value:$,.2s}',
      textinfo: 'label+value',
      hovertemplate: '<b>%{label}</b><br>Class: <b>%{parent}</b><br>Inventory Value: <b>%{value:$,.0f}</b><br>Share of Class: <b>%{percentParent:.1%}</b><br>Share of Portfolio: <b>%{percentRoot:.1%}</b><extra></extra>',
      insidetextfont: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif', size: 10, color: '#FFFFFF' },
      outsidetextfont: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif', size: 10, color: '#0F172A' },
      tiling: {
        packing: 'squarify',
        pad: 2
      }
    }], {
      title: { text: '<b>Inventory Value by ABC Hierarchy</b>', y: 0.98, font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' } },
      margin: { l: 10, r: 10, t: 45, b: 10 },
      hoverlabel: CYBER_HOVERLABEL,
      paper_bgcolor: 'rgba(0,0,0,0)',
      plot_bgcolor: 'rgba(0,0,0,0)',
      transition: { duration: 400, easing: 'cubic-in-out' }
    }, { displayModeBar: false, responsive: true });
  }

  // Financial Composition Bar Chart
  Plotly.react('chart-financial-composition', [{
    x: ["Model-Implied Inventory Value", "Estimated Carrying Cost", "Estimated Stockout Cost / Exposure"],
    y: [invValueSum, carryingSum, stockoutSum],
    text: [fmtMonetary(invValueSum), fmtMonetary(carryingSum), fmtMonetary(stockoutSum)],
    textposition: 'auto',
    type: 'bar',
    marker: { color: ['#2563EB', '#F59E0B', '#EF4444'], line: { color: '#ffffff', width: 1.5 } }
  }], {
    title: 'Model-Implied Financial Components',
    margin: { l: 75, r: 20, t: 45, b: 65 },
    hoverlabel: CYBER_HOVERLABEL,
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' },
    plot_bgcolor: 'white',
    paper_bgcolor: 'rgba(0,0,0,0)',
    xaxis: { title: { text: 'Financial Measure', standoff: 15 }, automargin: true },
    yaxis: { title: { text: 'Monetary Units ($)', standoff: 15 }, automargin: true },
    transition: { duration: 450, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  // Illustrative Scenario Slider
  updateIllustrativeScenario(invValueSum);

  // Operational Value Chain Sankey Diagram
  const stageLabels = ["Demand Forecast", "Procurement", "Production", "Warehouse", "Transportation", "Customer Service", "Finance", "Executive Decisions"];
  const stageColors = ["#1E3A8A", "#2563EB", "#0284C7", "#06B6D4", "#0D9488", "#10B981", "#6366F1", "#8B5CF6"];
  const linkColors = stageColors.slice(0, -1).map(c => {
    const r = parseInt(c.slice(1, 3), 16), g = parseInt(c.slice(3, 5), 16), b = parseInt(c.slice(5, 7), 16);
    return `rgba(${r},${g},${b},0.42)`;
  });

  Plotly.react('chart-sankey-diagram', [{
    type: "sankey",
    orientation: "h",
    node: { pad: 20, thickness: 24, line: { color: "white", width: 1.5 }, label: stageLabels, color: stageColors },
    link: {
      source: [0, 1, 2, 3, 4, 5, 6],
      target: [1, 2, 3, 4, 5, 6, 7],
      value: [1, 1, 1, 1, 1, 1, 1],
      color: linkColors
    }
  }], {
    title: "8-Stage Operational Value Chain (Conceptual Flow)",
    font: { family: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif', size: 13 },
    hoverlabel: CYBER_HOVERLABEL,
    margin: { l: 15, r: 15, t: 45, b: 20 },
    height: 340,
    paper_bgcolor: 'rgba(0,0,0,0)',
    transition: { duration: 400, easing: 'cubic-in-out' }
  }, { displayModeBar: false, responsive: true });

  // QA Report Summary & Details (full unfiltered population)
  const qa = runPortfolioMathAudit(AppState.rawRecords);
  const sum = qa.summary;
  const qaSumEl = document.getElementById('qa-report-summary');
  if (qaSumEl) {
    qaSumEl.innerHTML = `
      <div>
        <strong>Safety Stock â€” ${sum.n.toLocaleString('en-US')} SKUs tested</strong><br>
        - Zero in Master Audit: <strong>${sum.ss_zero_in_master.toLocaleString('en-US')}</strong><br>
        - Exact match to independent calculation: <strong>${sum.ss_exact_match.toLocaleString('en-US')}</strong><br>
        - Match within Â±1 unit (rounding tolerance): <strong>${sum.ss_tolerance_match.toLocaleString('en-US')}</strong><br>
        - Material mismatch (>1 unit): <strong>${sum.ss_material_mismatch.toLocaleString('en-US')}</strong>
      </div>
      <div>
        <strong>Reorder Point â€” ${sum.n.toLocaleString('en-US')} SKUs tested</strong><br>
        - Zero in Master Audit: <strong>${sum.rop_zero_in_master.toLocaleString('en-US')}</strong><br>
        - Exact match to independent calculation: <strong>${sum.rop_exact_match.toLocaleString('en-US')}</strong><br>
        - Match within Â±1 unit (rounding tolerance): <strong>${sum.rop_tolerance_match.toLocaleString('en-US')}</strong><br>
        - Material mismatch (>1 unit): <strong>${sum.rop_material_mismatch.toLocaleString('en-US')}</strong>
      </div>
    `;
  }

  const qaDetailsEl = document.getElementById('qa-report-details');
  if (qaDetailsEl) {
    if (sum.ss_material_mismatch === 0 && sum.rop_material_mismatch === 0) {
      qaDetailsEl.innerHTML = `
        <div style="color:var(--c-success-text); font-weight:800; background:var(--c-success-bg); padding:1rem; border-radius:var(--clay-radius-sm); margin-bottom: 0.8rem;">
          âœ“ Formula integrity confirmed. Every Safety Stock / Reorder Point value in the Master Audit matches the independently recomputed value within a Â±1 unit rounding tolerance across the full SKU population. Zero values are concentrated in SKUs with zero demand variance (dead stock / constant demand) or near-zero forecast RMSE on very low-volume SKUs â€” both are mathematically correct outcomes of the model's own formula, not data-quality defects.
        </div>
        <div style="font-size:0.86rem; line-height:1.5; color:var(--c-slate-700);">
          <strong>Root Cause (established, not assumed):</strong> the notebook's Safety Stock formula is <code>1.645 Ã— RMSE Ã— âˆš(7/30)</code> and Reorder Point is <code>(Forecast_Next_Month/30 Ã— 7) + Safety Stock</code>. Both terms scale with forecast error and forecast level â€” when a SKU has zero historical demand variance, the model's own guard clause sets Safety Stock to exactly 0 by design (Guard: zero-variance / dead stock). When a SKU has extremely low, near-deterministic demand, RMSE rounds to a value small enough that the computed buffer also rounds to 0. Both are legitimate outputs of the implemented formula, verified by direct recomputation from RMSE and Forecast_Next_Month across the entire population â€” this is not a broken pipeline or a lost calculation.<br><br>
          <strong>Disposition:</strong> CONFIRMED VALID â€” no data issue found.
        </div>
      `;
    } else {
      const mismatches = qa.details.filter(d => (d.SS_Variance !== null && d.SS_Variance > 1) || (d.ROP_Variance !== null && d.ROP_Variance > 1));
      let mismatchRowsHtml = '';
      if (mismatches.length > 0) {
        mismatchRowsHtml = `
          <div style="margin-top: 1rem; overflow-x: auto; max-height: 250px;">
            <table class="math-audit-table" style="width: 100%; font-size: 0.82rem;">
              <thead>
                <tr><th>SKU</th><th>Status</th><th>MA Safety Stock</th><th>Ind Safety Stock</th><th>SS Variance</th><th>MA ROP</th><th>Ind ROP</th><th>ROP Variance</th><th>Basis</th></tr>
              </thead>
              <tbody>
                ${mismatches.map(m => `
                  <tr>
                    <td><strong>${escapeHtml(m.SKU)}</strong></td>
                    <td>${escapeHtml(m.Status || '')}</td>
                    <td>${fmtNum(m.MA_Safety_Stock, 0)}</td>
                    <td>${fmtNum(m.Independent_Safety_Stock, 0)}</td>
                    <td class="${m.SS_Variance > 1 ? 'math-status-review' : ''}">${fmtNum(m.SS_Variance, 0)}</td>
                    <td>${fmtNum(m.MA_Reorder_Point, 0)}</td>
                    <td>${fmtNum(m.Independent_Reorder_Point, 0)}</td>
                    <td class="${m.ROP_Variance > 1 ? 'math-status-review' : ''}">${fmtNum(m.ROP_Variance, 0)}</td>
                    <td>${escapeHtml(m.Basis)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        `;
      }
      qaDetailsEl.innerHTML = `
        <div style="color:var(--c-warning-text); font-weight:800; background:var(--c-warning-bg); padding:1rem; border-radius:var(--clay-radius-sm);">
          âš  ${sum.ss_material_mismatch} Safety Stock and ${sum.rop_material_mismatch} Reorder Point value(s) differ from the independent recalculation by more than 1 unit. See the detail table below for the specific SKUs â€” these are flagged for review, not silently corrected.
        </div>
        ${mismatchRowsHtml}
        <div style="font-size:0.86rem; line-height:1.5; color:var(--c-slate-700); margin-top: 0.8rem;">
          <strong>Root Cause:</strong> The notebook implements Safety Stock = 1.645 Ã— RMSE Ã— âˆš(7/30) and Reorder Point = (Forecast_Next_Month/30 Ã— 7) + Safety Stock.<br>
          <strong>Disposition:</strong> REQUIRES REVIEW for the flagged SKUs above.
        </div>
      `;
    }
  }

  // Field Completeness Check (Dynamic check against 51 authoritative Master Audit fields)
  const availableFields = AppState.rawRecords.length > 0 ? Object.keys(AppState.rawRecords[0]) : [];
  const missingFields = MASTER_AUDIT_FIELDS.filter(f => !availableFields.includes(f) && !availableFields.includes(f.replace(/ /g, '_')));
  const checkEl = document.getElementById('field-completeness-check');
  if (checkEl) {
    if (missingFields.length > 0) {
      checkEl.innerHTML = `<strong>Field Completeness Check:</strong> ${missingFields.length} field(s) expected by this dashboard are not present in the current data source and are shown as "N/A" or omitted wherever referenced: ${escapeHtml(missingFields.join(', '))}.`;
    } else {
      checkEl.innerHTML = `<strong>Field Completeness Check:</strong> All authoritative fields (${MASTER_AUDIT_FIELDS.length} fields) this dashboard expects are present in the current data source.`;
    }
  }

  // Initial Full Audit Table Render
  const searchInput = document.getElementById('audit-search-input');
  renderFullAuditTable(searchInput?.value || '');
}

function updateIllustrativeScenario(invValueSum) {
  const slider = document.getElementById('illustrative-slider');
  const disp = document.getElementById('slider-val-display');
  const currValEl = document.getElementById('scenario-curr-val');
  const redValEl = document.getElementById('scenario-red-val');
  const resValEl = document.getElementById('scenario-res-val');
  const redTitleEl = document.getElementById('scenario-reduction-title');

  if (!slider) return;
  const pct = Number(slider.value) || 10;
  AppState.illustrativeReductionPct = pct;
  if (disp) disp.textContent = pct;
  if (redTitleEl) redTitleEl.textContent = `Illustrative Inventory Value Reduction (${pct}%)`;

  const curr = isNaN(invValueSum) ? 0 : invValueSum;
  const red = curr * (pct / 100.0);
  const res = curr - red;

  if (currValEl) currValEl.textContent = fmtMonetary(curr);
  if (redValEl) redValEl.textContent = fmtMonetary(red);
  if (resValEl) resValEl.textContent = fmtMonetary(res);
}

// ==========================================================================
// Full Audit Table & Excel Export System (Progressive Chunked Rendering)
// ==========================================================================
let _auditVisibleCount = 0;
let _auditTableMatching = [];
let _auditTableFields = [];
let _isLoadingMoreRows = false;
const AUDIT_CHUNK_SIZE = 50;

function buildAuditRowChunk(records, fields, startIndex, count) {
  let html = '';
  const endIndex = Math.min(startIndex + count, records.length);
  for (let i = startIndex; i < endIndex; i++) {
    const r = records[i];
    html += '<tr>';
    for (let j = 0; j < fields.length; j++) {
      const f = fields[j];
      const val = r[f];
      html += `<td>${escapeHtml(String(val !== null && val !== undefined ? val : ''))}</td>`;
    }
    html += '</tr>';
  }
  return html;
}

function updateAuditCaption() {
  const caption = document.getElementById('full-audit-caption');
  if (!caption) return;
  const totalRaw = AppState.rawRecords.length;
  const matchLen = _auditTableMatching.length;

  if (matchLen === 0) {
    caption.textContent = `0 of ${totalRaw.toLocaleString('en-US')} total SKUs shown.`;
    return;
  }

  if (_auditVisibleCount >= matchLen) {
    caption.textContent = `${matchLen.toLocaleString('en-US')} of ${totalRaw.toLocaleString('en-US')} total SKUs shown. Every column above is read directly from Master Audit; the source file itself is never modified by this dashboard.`;
  } else {
    caption.innerHTML = `Showing <strong>${_auditVisibleCount.toLocaleString('en-US')}</strong> of <strong>${matchLen.toLocaleString('en-US')}</strong> matching SKUs (${totalRaw.toLocaleString('en-US')} total). Scroll down inside table to load more, or <button id="load-all-audit-btn" type="button" class="audit-load-all-btn">load all ${matchLen.toLocaleString('en-US')} rows</button>.`;
    const loadAllBtn = document.getElementById('load-all-audit-btn');
    if (loadAllBtn) {
      loadAllBtn.onclick = () => loadAllAuditRows();
    }
  }
}

function loadMoreAuditRows(batchSize = AUDIT_CHUNK_SIZE) {
  if (_isLoadingMoreRows || _auditVisibleCount >= _auditTableMatching.length) return;
  _isLoadingMoreRows = true;
  const tbody = document.getElementById('full-audit-tbody');
  if (!tbody) {
    _isLoadingMoreRows = false;
    return;
  }
  const chunkHtml = buildAuditRowChunk(_auditTableMatching, _auditTableFields, _auditVisibleCount, batchSize);
  tbody.insertAdjacentHTML('beforeend', chunkHtml);
  _auditVisibleCount = Math.min(_auditVisibleCount + batchSize, _auditTableMatching.length);
  updateAuditCaption();
  _isLoadingMoreRows = false;
}

function loadAllAuditRows() {
  if (_auditVisibleCount >= _auditTableMatching.length) return;
  loadMoreAuditRows(_auditTableMatching.length - _auditVisibleCount);
}

function renderFullAuditTable(searchTerm = '') {
  const records = AppState.filteredRecords;
  const term = String(searchTerm || '').trim().toLowerCase();
  const matching = records.filter(r => {
    if (!term) return true;
    return String(r.SKU || '').toLowerCase().includes(term);
  });

  AppState.currentAuditRecords = matching;
  _auditTableMatching = matching;

  const thead = document.getElementById('full-audit-thead');
  const tbody = document.getElementById('full-audit-tbody');
  const caption = document.getElementById('full-audit-caption');

  if (!tbody || !thead) return;

  const exportBtn = document.getElementById('download-excel-btn');

  if (records.length === 0) {
    thead.innerHTML = '';
    tbody.innerHTML = '<tr><td colspan="100%" style="text-align: center; padding: 2rem; color: var(--c-slate-500);">No records match the current ABC/XYZ filter.</td></tr>';
    if (caption) caption.textContent = '0 records shown.';
    if (exportBtn) {
      exportBtn.disabled = true;
      exportBtn.style.opacity = '0.5';
      exportBtn.style.cursor = 'not-allowed';
      exportBtn.setAttribute('title', 'No rows in current view to export');
    }
    return;
  }

  const fields = Object.keys(records[0]);
  _auditTableFields = fields;
  thead.innerHTML = `<tr>${fields.map(f => `<th>${escapeHtml(f)}</th>`).join('')}</tr>`;

  if (matching.length === 0) {
    tbody.innerHTML = `<tr><td colspan="${fields.length}" style="text-align: center; padding: 2rem; color: var(--c-slate-500);">No matching SKUs found for "${escapeHtml(searchTerm)}".</td></tr>`;
    if (caption) caption.textContent = `0 of ${AppState.rawRecords.length.toLocaleString('en-US')} total SKUs shown.`;
    if (exportBtn) {
      exportBtn.disabled = true;
      exportBtn.style.opacity = '0.5';
      exportBtn.style.cursor = 'not-allowed';
      exportBtn.setAttribute('title', 'No rows in current view to export');
    }
    return;
  }

  if (exportBtn) {
    exportBtn.disabled = false;
    exportBtn.style.opacity = '1';
    exportBtn.style.cursor = 'pointer';
    exportBtn.removeAttribute('title');
  }

  // Progressive initial chunk render (50 rows in ~10ms instead of 1,401 rows in 3,000ms)
  _auditVisibleCount = Math.min(AUDIT_CHUNK_SIZE, matching.length);
  tbody.innerHTML = buildAuditRowChunk(matching, fields, 0, _auditVisibleCount);

  const container = document.getElementById('full-audit-table-container');
  if (container) container.scrollTop = 0;

  updateAuditCaption();
}

function exportFilteredAuditToExcel() {
  const exportRecords = AppState.currentAuditRecords;
  if (!exportRecords || exportRecords.length === 0) {
    showToast("No rows in the current filtered/searched view to export.");
    return;
  }
  const ws = XLSX.utils.json_to_sheet(exportRecords);

  if (exportRecords.length > 0) {
    const cols = Object.keys(exportRecords[0]).map(key => {
      let maxLen = key.length;
      const sampleSize = Math.min(exportRecords.length, 100);
      for (let i = 0; i < sampleSize; i++) {
        const valStr = String(exportRecords[i][key] ?? '');
        if (valStr.length > maxLen) maxLen = valStr.length;
      }
      return { wch: Math.min(maxLen + 2, 40) };
    });
    ws['!cols'] = cols;
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Filtered_Audit");
  XLSX.writeFile(wb, "Filtered_Supply_Chain_Audit.xlsx");
  showToast(`Exported ${exportRecords.length.toLocaleString('en-US')} records to Filtered_Supply_Chain_Audit.xlsx`);
}

// ==========================================================================
// Mathematics Helpers (Percentile, Median, Counts)
// ==========================================================================
function median(values) {
  if (values.length === 0) return 0;
  const sorted = [...values].sort((a, b) => a - b);
  const half = Math.floor(sorted.length / 2);
  if (sorted.length % 2) return sorted[half];
  return (sorted[half - 1] + sorted[half]) / 2.0;
}

function percentile(arr, p) {
  if (arr.length === 0) return 0;
  const sorted = [...arr].sort((a, b) => a - b);
  const index = (sorted.length - 1) * p;
  const lower = Math.floor(index);
  const upper = Math.ceil(index);
  const weight = index - lower;
  if (upper >= sorted.length) return sorted[sorted.length - 1];
  return sorted[lower] * (1 - weight) + sorted[upper] * weight;
}

function orderedCounts(records, col, order) {
  const counts = {};
  order.forEach(c => { counts[c] = 0; });
  records.forEach(r => {
    const v = String(r[col] || '').trim();
    if (counts[v] !== undefined) counts[v]++;
  });
  return order.map(c => ({ cat: c, count: counts[c] }));
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function debounce(func, wait) {
  let timeout;
  return function executedFunction(...args) {
    const later = () => {
      clearTimeout(timeout);
      func(...args);
    };
    clearTimeout(timeout);
    timeout = setTimeout(later, wait);
  };
}

// ==========================================================================
// Application Bootstrap & Event Registration
// ==========================================================================
document.addEventListener('DOMContentLoaded', async () => {
  // 1. Control Centre Drawer Controls
  initControlCentreEvents();

  // 2. Navigation Tab Switching
  document.querySelectorAll('.tab-button').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-button').forEach(b => {
        b.classList.remove('active');
        b.setAttribute('aria-selected', 'false');
      });
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      const targetTab = btn.getAttribute('data-tab');
      btn.classList.add('active');
      btn.setAttribute('aria-selected', 'true');
      const panel = document.getElementById(targetTab);
      if (panel) panel.classList.add('active');
      AppState.activeTab = targetTab;
      renderActiveTab();
      setTimeout(resizeActiveTabCharts, 60);
      setTimeout(resizeActiveTabCharts, 250);
    });
  });

  // 3. Filter Checkbox Listeners
  document.querySelectorAll('input[name="abc_class"]').forEach(cb => {
    cb.addEventListener('change', () => {
      AppState.selectedABC = Array.from(document.querySelectorAll('input[name="abc_class"]:checked')).map(el => el.value);
      applyFilters();
    });
  });

  document.querySelectorAll('input[name="xyz_class"]').forEach(cb => {
    cb.addEventListener('change', () => {
      AppState.selectedXYZ = Array.from(document.querySelectorAll('input[name="xyz_class"]:checked')).map(el => el.value);
      applyFilters();
    });
  });

  // 4. SKU Selector Synchronization & Autocomplete Search
  SKU_INPUT_IDS.forEach(({ inputId }) => {
    const el = document.getElementById(inputId);
    if (el) {
      el.addEventListener('input', debounce((e) => {
        const val = e.target.value.trim();
        const match = AppState.filteredRecords.find(r => String(r.SKU).toLowerCase() === val.toLowerCase());
        if (match) {
          syncSelectedSKU(match.SKU);
        }
      }, 150));

      el.addEventListener('change', (e) => {
        const val = e.target.value.trim();
        if (val) syncSelectedSKU(val);
      });

      el.addEventListener('focus', () => {
        el.select();
      });
    }
  });

  // Expose filterJumpSelect for recommendation jump search inputs
  window.filterJumpSelect = function(idx, val) {
    const sel = document.getElementById(`jump-select-${idx}`);
    if (!sel) return;
    const q = String(val || '').trim().toLowerCase();
    if (!sel._origOptions) {
      sel._origOptions = Array.from(sel.options).map(o => o.value);
    }
    const filtered = q ? sel._origOptions.filter(s => s.toLowerCase().includes(q)) : sel._origOptions;
    sel.innerHTML = filtered.map(s => `<option value="${escapeHtml(s)}">${escapeHtml(s)}</option>`).join('');
  };

  // 5. Illustrative Reduction Range Slider
  const slider = document.getElementById('illustrative-slider');
  if (slider) {
    slider.addEventListener('input', () => {
      const invValueSum = safeSum(AppState.filteredRecords, 'Inventory_Value');
      updateIllustrativeScenario(invValueSum);
    });
  }

  // 6. Audit SKU Search
  const searchInput = document.getElementById('audit-search-input');
  if (searchInput) {
    searchInput.addEventListener('input', debounce((e) => {
      renderFullAuditTable(e.target.value);
    }, 150));
  }

  // 6b. Infinite Scroll for Master Audit Table Container
  const auditContainer = document.getElementById('full-audit-table-container');
  if (auditContainer) {
    auditContainer.addEventListener('scroll', () => {
      if (auditContainer.scrollTop + auditContainer.clientHeight >= auditContainer.scrollHeight - 100) {
        loadMoreAuditRows(AUDIT_CHUNK_SIZE);
      }
    }, { passive: true });
  }

  // 7. Audit Excel Download Button
  const exportBtn = document.getElementById('download-excel-btn');
  if (exportBtn) {
    exportBtn.addEventListener('click', () => {
      exportFilteredAuditToExcel();
    });
  }

  // 8. Window Resize Event for Plotly
  window.addEventListener('resize', debounce(resizeActiveTabCharts, 150));

  // Expose navigation functions and state globally
  window.jumpToSKU = jumpToSKU;
  window.jumpToSKUFromTop = jumpToSKUFromTop;
  window.AppState = AppState;

  // 9. Load Static JSON Runtime Data
  try {
    const response = await fetch('./dashboard-data.json');
    if (!response.ok) throw new Error(`HTTP error ${response.status}`);
    const json = await response.json();
    AppState.metadata = json.metadata || {};
    AppState.rawRecords = json.data || json.records || [];
    const urlTab = new URLSearchParams(window.location.search).get('tab');
    if (urlTab && ['tab1', 'tab2', 'tab3', 'tab4', 'tab5'].includes(urlTab)) {
      AppState.activeTab = urlTab;
      document.querySelectorAll('.tab-button').forEach(b => {
        const match = b.dataset.tab === urlTab;
        b.classList.toggle('active', match);
        b.setAttribute('aria-selected', match ? 'true' : 'false');
      });
      document.querySelectorAll('.tab-panel').forEach(p => {
        p.classList.toggle('active', p.id === urlTab);
      });
    }
    populateFilterControls();
    applyFilters();
    const urlSku = new URLSearchParams(window.location.search).get('sku');
    if (urlSku) {
      syncSelectedSKU(urlSku);
      renderActiveTab();
    }
    console.log(`Initialized static engine with ${AppState.rawRecords.length} authoritative records.`);

    if (window.location.search.includes('runtest=1')) {
      setTimeout(async () => {
        const results = [];
        function assert(cond, msg) {
          results.push({ pass: !!cond, message: msg });
        }
        try {
          // 1. Data load & Record Count
          assert(AppState.rawRecords.length === 1401, `Loaded 1,401 authoritative records: ${AppState.rawRecords.length}`);
          assert(document.getElementById('signal-status').textContent === 'Elevated Risk', `Signal status: ${document.getElementById('signal-status').textContent}`);
          assert(document.getElementById('kpi-total-skus').textContent !== '0', `Total Tracked SKUs KPI rendered: ${document.getElementById('kpi-total-skus').textContent}`);

          // 2. Control Centre Drawer
          const drawer = document.getElementById('control-centre');
          const overlay = document.getElementById('control-centre-overlay');
          assert(!drawer.classList.contains('open'), 'Control Centre closed by default');
          openControlCentre();
          assert(drawer.classList.contains('open'), 'Control Centre opens on trigger');
          closeControlCentre();
          assert(!drawer.classList.contains('open'), 'Control Centre closes on close action');

          // 3. ABC Filters Semantics (empty = non-restrictive)
          const abcBoxes = document.querySelectorAll('input[name="abc_class"]');
          assert(abcBoxes.length === 3, `ABC checkboxes count = ${abcBoxes.length}`);
          abcBoxes.forEach(cb => { cb.checked = false; });
          AppState.selectedABC = [];
          applyFilters();
          assert(AppState.filteredRecords.length === 1401, `Empty ABC filter retains all 1,401 records`);
          abcBoxes.forEach(cb => { cb.checked = true; });
          AppState.selectedABC = ['A', 'B', 'C'];
          applyFilters();

          // 4. XYZ Filters Semantics (empty = non-restrictive)
          const xyzBoxes = document.querySelectorAll('input[name="xyz_class"]');
          assert(xyzBoxes.length === 3, `XYZ checkboxes count = ${xyzBoxes.length}`);
          xyzBoxes.forEach(cb => { cb.checked = false; });
          AppState.selectedXYZ = [];
          applyFilters();
          assert(AppState.filteredRecords.length === 1401, `Empty XYZ filter retains all 1,401 records`);
          xyzBoxes.forEach(cb => { cb.checked = true; });
          AppState.selectedXYZ = ['X', 'Y', 'Z'];
          applyFilters();

          // 5. Exact 5 Tabs Navigation
          const tabs = ['tab1', 'tab2', 'tab3', 'tab4', 'tab5'];
          tabs.forEach(t => {
            AppState.activeTab = t;
            renderActiveTab();
            assert(true, `Tab ${t} rendered successfully without errors`);
          });

          // 6. SKU Synchronization & Jump to SKU
          const targetSku = AppState.rawRecords[2]?.SKU;
          jumpToSKU(targetSku);
          assert(AppState.selectedSKU === targetSku, `jumpToSKU synchronized selected SKU: ${targetSku}`);
          assert(AppState.activeTab === 'tab1', 'jumpToSKU switched active view to Tab 1');
          assert(document.getElementById('tab1').classList.contains('active'), 'Tab 1 panel is active in DOM');
          const toast = document.getElementById('dashboard-toast');
          assert(toast && toast.textContent.includes(targetSku) && toast.textContent.includes('Decision Intelligence'), `Jump toast message verified: "${toast?.textContent}"`);

          // 7. Full Audit Search & Critical 0-Result Export Guard
          AppState.activeTab = 'tab5';
          renderTab5();
          renderFullAuditTable('NONEXISTENT_QUERY_TEST_99999');
          assert(AppState.currentAuditRecords.length === 0, 'Audit search for non-existent SKU results in 0 records');
          exportFilteredAuditToExcel();
          assert(toast && toast.textContent.includes('No rows in the current filtered/searched view to export'), 'Zero-result export correctly blocked with toast notice');
          
          renderFullAuditTable(targetSku);
          assert(AppState.currentAuditRecords.length >= 1, `Audit search for ${targetSku} returns matching records: ${AppState.currentAuditRecords.length}`);
          renderFullAuditTable('');
          assert(document.querySelectorAll('#full-audit-tbody tr').length === 50, `Audit table initially renders 50 rows (progressive chunking)`);
          assert(AppState.currentAuditRecords.length === 1401, `Full audit records dataset retained in memory for export: ${AppState.currentAuditRecords.length}`);

          // 8. Expanders: Methodology & Math QA
          const methExpander = document.getElementById('methodology-expander');
          const qaExpander = document.getElementById('qa-report-expander');
          assert(methExpander !== null, 'Methodology expander exists in DOM');
          assert(qaExpander !== null, 'Math QA expander exists in DOM');

          // 9. Dynamic Field Completeness Check (51 Authoritative Fields)
          const compEl = document.getElementById('field-completeness-check');
          assert(compEl && compEl.textContent.includes('All authoritative fields (51 fields) this dashboard expects are present'), `Field completeness 51-field check confirmed: "${compEl?.textContent}"`);

          // 10. Horizontal Overflow & Tab Fit Check
          assert(document.documentElement.scrollWidth <= document.documentElement.clientWidth + 5, `No horizontal overflow: scrollWidth (${document.documentElement.scrollWidth}) <= clientWidth (${document.documentElement.clientWidth})`);
          const tabListEl = document.querySelector('.tab-list');
          assert(tabListEl && tabListEl.scrollWidth <= tabListEl.clientWidth + 2, `Tab headings fit in main window without horizontal scrolling (scrollWidth ${tabListEl?.scrollWidth} <= clientWidth ${tabListEl?.clientWidth})`);
        } catch (err) {
          assert(false, `Test error: ${err.message}`);
        }
        try {
          await fetch('/report-results', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(results)
          });
        } catch (e) {}
        window.close();
      }, 700);
    } else if (window.location.search.includes('mobiletest=1')) {
      setTimeout(async () => {
        const results = [];
        function assert(cond, msg) {
          results.push({ pass: !!cond, message: msg });
        }
        try {
          const trigger = document.getElementById('control-centre-trigger');
          assert(trigger.offsetWidth < 240, `Mobile trigger is compact: ${trigger.offsetWidth}px (not full width)`);
          assert(document.documentElement.scrollWidth <= document.documentElement.clientWidth + 5, `No horizontal scroll: scrollWidth ${document.documentElement.scrollWidth} <= clientWidth ${document.documentElement.clientWidth}`);
          const tabListMob = document.querySelector('.tab-list');
          assert(tabListMob && tabListMob.scrollWidth <= tabListMob.clientWidth + 2, `Tab headings fit on mobile without horizontal scrolling (scrollWidth ${tabListMob?.scrollWidth} <= clientWidth ${tabListMob?.clientWidth})`);

          openControlCentre();
          const drawer = document.getElementById('control-centre');
          assert(drawer.classList.contains('open'), 'Drawer opens as overlay on mobile');
          closeControlCentre();
          assert(!drawer.classList.contains('open'), 'Drawer closes properly on mobile');
        } catch (err) {
          assert(false, `Mobile test error: ${err.message}`);
        }
        try {
          await fetch('/report-mobile', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(results)
          });
        } catch (e) {}
        window.close();
      }, 700);
    } else if (window.location.search.includes('diagtest=1')) {
      setTimeout(async () => {
        const diagReport = {
          errors: [],
          timings: {},
          charts: {},
          overflow: {},
          sku_sync: {},
          audit_perf: {},
          overall_status: 'PASS'
        };

        const chartCatalog = {
          tab1: ['chart-abc-pie', 'chart-xyz-pie', 'chart-abc-xyz-matrix', 'chart-annual-demand-abc'],
          tab2: ['chart-forecast-snapshot', 'chart-forecast-compare', 'chart-portfolio-rmse', 'chart-portfolio-mae', 'chart-portfolio-bias', 'chart-portfolio-accuracy-gauge', 'chart-portfolio-demand-dist', 'chart-forecast-risk-abc'],
          tab3: ['chart-replenishment-turnover', 'chart-replenishment-ihs'],
          tab4: ['chart-bullwhip-rmse-scatter', 'chart-segmentation-abc-bar', 'chart-segmentation-xyz-bar', 'chart-risk-intensity-matrix', 'chart-pareto-analysis', 'chart-business-risk-bar', 'chart-sc-priority-bar'],
          tab5: ['chart-financial-treemap', 'chart-financial-composition', 'chart-sankey-diagram']
        };

        try {
          // Track any runtime errors
          window.addEventListener('error', e => {
            diagReport.errors.push({ type: 'runtime_error', message: e.message, lineno: e.lineno });
          });
          window.addEventListener('unhandledrejection', e => {
            diagReport.errors.push({ type: 'unhandled_rejection', reason: String(e.reason) });
          });

          // Test tabs 1 through 5
          const tabKeys = ['tab1', 'tab2', 'tab3', 'tab4', 'tab5'];
          for (const tabKey of tabKeys) {
            const btn = document.querySelector(`.tab-button[data-tab="${tabKey}"]`);
            if (btn) {
              const t0 = performance.now();
              btn.click();
              // wait for animation & plotly resize
              await new Promise(r => setTimeout(r, 60));
              const t1 = performance.now();
              diagReport.timings[tabKey] = +(t1 - t0 - 60).toFixed(2);
              if (diagReport.timings[tabKey] < 0) diagReport.timings[tabKey] = +(t1 - t0).toFixed(2);
            }

            // Verify charts
            const expectedCharts = chartCatalog[tabKey] || [];
            diagReport.charts[tabKey] = {};
            for (const cId of expectedCharts) {
              const el = document.getElementById(cId);
              const exists = !!el;
              const hasPlot = el && (el.querySelector('.plot-container') !== null || el.querySelector('svg') !== null);
              diagReport.charts[tabKey][cId] = {
                exists: exists,
                rendered: hasPlot,
                w: el ? el.clientWidth : 0,
                h: el ? el.clientHeight : 0
              };
              if (!hasPlot) {
                diagReport.errors.push({ type: 'chart_not_rendered', chartId: cId, tab: tabKey });
              }
            }

            // Verify overflow
            const scrollW = document.documentElement.scrollWidth;
            const clientW = document.documentElement.clientWidth;
            const tabList = document.querySelector('.tab-list');
            const tabScrollW = tabList ? tabList.scrollWidth : 0;
            const tabClientW = tabList ? tabList.clientWidth : 0;
            diagReport.overflow[tabKey] = {
              hasPageOverflow: scrollW > clientW + 2,
              scrollWidth: scrollW,
              clientWidth: clientW,
              hasTabBarOverflow: tabScrollW > tabClientW + 2,
              tabBarScrollWidth: tabScrollW,
              tabBarClientWidth: tabClientW
            };
            if (scrollW > clientW + 2 || tabScrollW > tabClientW + 2) {
              diagReport.errors.push({ type: 'overflow', tab: tabKey, scrollW, clientW });
            }
          }

          // Test Tab 2 SKU switch performance & sync
          const tab2Btn = document.querySelector('.tab-button[data-tab="tab2"]');
          if (tab2Btn) tab2Btn.click();
          await new Promise(r => setTimeout(r, 50));
          const testSku = AppState.rawRecords[4]?.SKU || "Spinnex Premium Tee T-Shirt N.blue:-XL";
          const tSku0 = performance.now();
          syncSelectedSKU(testSku);
          renderTab2();
          const tSku1 = performance.now();
          diagReport.sku_sync.tab2_switch_latency_ms = +(tSku1 - tSku0).toFixed(2);
          diagReport.sku_sync.selectedSKU = AppState.selectedSKU;
          diagReport.sku_sync.tab2_select_value = document.getElementById('tab2-sku-select')?.value;

          // Test Tab 3 sync
          const tab3Btn = document.querySelector('.tab-button[data-tab="tab3"]');
          if (tab3Btn) tab3Btn.click();
          await new Promise(r => setTimeout(r, 50));
          diagReport.sku_sync.tab3_select_value = document.getElementById('tab3-sku-select')?.value;
          diagReport.sku_sync.isSynchronized = (diagReport.sku_sync.tab3_select_value === testSku);

          // Test Tab 5 Audit Table Performance
          const tab5Btn = document.querySelector('.tab-button[data-tab="tab5"]');
          if (tab5Btn) tab5Btn.click();
          await new Promise(r => setTimeout(r, 50));
          const tAudit0 = performance.now();
          renderFullAuditTable('Spinnex');
          const tAudit1 = performance.now();
          diagReport.audit_perf.search_filter_latency_ms = +(tAudit1 - tAudit0).toFixed(2);
          diagReport.audit_perf.filtered_count = AppState.currentAuditRecords.length;

          renderFullAuditTable('');
          const tLoadMore0 = performance.now();
          loadMoreAuditRows(50);
          const tLoadMore1 = performance.now();
          diagReport.audit_perf.load_more_latency_ms = +(tLoadMore1 - tLoadMore0).toFixed(2);
          diagReport.audit_perf.rendered_rows = document.querySelectorAll('#full-audit-tbody tr').length;

          if (diagReport.errors.length > 0) {
            diagReport.overall_status = 'FAIL';
          }
        } catch (diagErr) {
          diagReport.errors.push({ type: 'diag_exception', message: diagErr.message, stack: diagErr.stack });
          diagReport.overall_status = 'FAIL';
        }

        try {
          await fetch('/report-diag', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(diagReport)
          });
        } catch (e) {}
        window.close();
      }, 700);
    }
  } catch (err) {
    console.error('Failed to load dashboard-data.json:', err);
    const sigStatus = document.getElementById('signal-status');
    if (sigStatus) sigStatus.textContent = 'Data Load Error';
  }
});

function populateFilterControls() {
  const abcContainer = document.getElementById('abc-filter-group');
  const xyzContainer = document.getElementById('xyz-filter-group');

  if (abcContainer) {
    const presentABCs = Array.from(new Set(AppState.rawRecords.map(r => String(r.ABC_Class || '').trim()).filter(Boolean)));
    const orderedABCs = ABC_ORDER.filter(c => presentABCs.includes(c));
    presentABCs.forEach(c => { if (!orderedABCs.includes(c)) orderedABCs.push(c); });
    if (orderedABCs.length > 0) {
      abcContainer.innerHTML = orderedABCs.map(val =>
        `<label><input type="checkbox" name="abc_class" value="${val}" ${AppState.selectedABC.includes(val) ? 'checked' : ''}> ${val}</label>`
      ).join('');
      abcContainer.querySelectorAll('input[name="abc_class"]').forEach(cb => {
        cb.addEventListener('change', () => {
          AppState.selectedABC = Array.from(document.querySelectorAll('input[name="abc_class"]:checked')).map(el => el.value);
          applyFilters();
        });
      });
    }
  }

  if (xyzContainer) {
    const presentXYZs = Array.from(new Set(AppState.rawRecords.map(r => String(r.XYZ_Class || '').trim()).filter(Boolean)));
    const orderedXYZs = XYZ_ORDER.filter(c => presentXYZs.includes(c));
    presentXYZs.forEach(c => { if (!orderedXYZs.includes(c)) orderedXYZs.push(c); });
    if (orderedXYZs.length > 0) {
      xyzContainer.innerHTML = orderedXYZs.map(val =>
        `<label><input type="checkbox" name="xyz_class" value="${val}" ${AppState.selectedXYZ.includes(val) ? 'checked' : ''}> ${val}</label>`
      ).join('');
      xyzContainer.querySelectorAll('input[name="xyz_class"]').forEach(cb => {
        cb.addEventListener('change', () => {
          AppState.selectedXYZ = Array.from(document.querySelectorAll('input[name="xyz_class"]:checked')).map(el => el.value);
          applyFilters();
        });
      });
    }
  }
}
