// Retail KPI Studio - Store Sales Report Generator
// Client-side Application Logic

document.addEventListener('DOMContentLoaded', () => {
  // Store default / sample template data
  const sampleData = {
    store: 'DEMO STORE',
    date: '24-09-2026',
    month_target: 3000000,
    target: 90000,
    sale: 105000,
    ff: 280,
    bill: 55,
    unit: 130,
    upt: '2.4',
    atv: '1909',
    auv: '808',
    conv: '20%',

    mtd_tgt: 2300000,
    mtd_sale: 2250000,
    mtd_sale_pct: '98%',
    mtd_ff: 6000,
    mtd_bill: 950,
    mtd_qty: 2150,
    mtd_upt: '2.26',
    mtd_atv: '2368',
    mtd_auv: '1047',
    mtd_conv: '16%',

    atvd_sale: 0,
    atvd_qty: 0,
    atvd_contri_pct: '0%',
    mtd_atvd_contri_pct: '1%',

    hotspot_sale: 10000,
    hotspot_pct: '10%',
    mtd_hotspot_pct: '4%',

    self_demo_sale: 4000,
    self_demo_pct: '4%',
    mtd_self_demo_pct: '4%',

    core: 4500,
    core_pct: '4%',
    mtd_core_sale: 100000,
    mtd_core_demo_pct: '4%',

    mii_sale: 12000,
    mii_sale_pct: '11%',
    mtd_mii_sale: 200000,
    mtd_mii_sale_pct: '9%',

    battery: 8,
    battery_attach: '15%',
    mtd_battery_attach: '12%',

    bubble: 6,
    bubble_attach: '11%',
    mtd_bubble_attach: '2%',

    multibill_tgt: '60%',
    multibill: 28,
    multibill_pct: '51%',

    loyalty: 5,
    loyalty_pct: '8%',
    mtd_loyalty: 116,
    mtd_loyalty_pct: '12%',

    redcard: 0,
    mtd_redcard: 1,

    mtd_ly_sale: 2264430,
    ly_sale: 68405,
    lfl_pct: '2%',

    tomorrow_tgt: 96074,
    tomorrow_netplay_tgt: 6239,

    tele_tgt: 100,
    tele_calls: 50,
    tele_sale: 0,
    tele_bills: 0,
    instore_sale: 0,
    instore_bill: 0,
    hd_sale: 0,
    hd_bill: 0,
    mtd_call: 1450,
    mtd_tele_bill: 0,
    mtd_tele_sale: 46834,

    netplay_tgt: 5429,
    netplay_sale: 1600,
    mtd_netplay_tgt: 104639,
    mtd_netplay_sale: 46450
  };

  // Field mapping between HTML ID and State key
  const fieldIds = [
    'store', 'date', 'month_target', 'target', 'sale', 'ff', 'bill', 'unit', 'upt', 'atv', 'auv', 'conv',
    'mtd_tgt', 'mtd_sale', 'mtd_sale_pct', 'mtd_ff', 'mtd_bill', 'mtd_qty', 'mtd_upt', 'mtd_atv', 'mtd_auv', 'mtd_conv',
    'atvd_sale', 'atvd_qty', 'atvd_contri_pct', 'mtd_atvd_contri_pct',
    'hotspot_sale', 'hotspot_pct', 'mtd_hotspot_pct',
    'self_demo_sale', 'self_demo_pct', 'mtd_self_demo_pct',
    'core', 'core_pct', 'mtd_core_sale', 'mtd_core_demo_pct',
    'mii_sale', 'mii_sale_pct', 'mtd_mii_sale', 'mtd_mii_sale_pct',
    'battery', 'battery_attach', 'mtd_battery_attach',
    'bubble', 'bubble_attach', 'mtd_bubble_attach',
    'multibill_tgt', 'multibill', 'multibill_pct',
    'loyalty', 'loyalty_pct', 'mtd_loyalty', 'mtd_loyalty_pct',
    'redcard', 'mtd_redcard',
    'mtd_ly_sale', 'ly_sale', 'lfl_pct',
    'tomorrow_tgt', 'tomorrow_netplay_tgt',
    'tele_tgt', 'tele_calls', 'tele_sale', 'tele_bills', 'instore_sale', 'instore_bill', 'hd_sale', 'hd_bill', 'mtd_call', 'mtd_tele_bill', 'mtd_tele_sale',
    'netplay_tgt', 'netplay_sale', 'mtd_netplay_tgt', 'mtd_netplay_sale'
  ];

  // DOM Elements
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const fileUploadStatus = document.getElementById('fileUploadStatus');
  const uploadedFileName = document.getElementById('uploadedFileName');
  const uploadSummaryText = document.getElementById('uploadSummaryText');
  const btnClearFile = document.getElementById('btnClearFile');
  const btnLoadSample = document.getElementById('btnLoadSample');
  const btnReset = document.getElementById('btnReset');
  const btnToggleTheme = document.getElementById('btnToggleTheme');
  const btnCopyReport = document.getElementById('btnCopyReport');
  const btnDownloadTxt = document.getElementById('btnDownloadTxt');
  const btnDownloadXlsx = document.getElementById('btnDownloadXlsx');
  const reportOutput = document.getElementById('reportOutput');
  const toast = document.getElementById('toast');
  const tabButtons = document.querySelectorAll('.tab-btn');
  const tabContents = document.querySelectorAll('.tab-content');
  const formatRadios = document.querySelectorAll('input[name="formatMode"]');

  // KPI Dashboard Cards
  const cardTodayAch = document.getElementById('cardTodayAch');
  const barTodayAch = document.getElementById('barTodayAch');
  const cardTodaySub = document.getElementById('cardTodaySub');
  const cardMtdAch = document.getElementById('cardMtdAch');
  const barMtdAch = document.getElementById('barMtdAch');
  const cardMtdSub = document.getElementById('cardMtdSub');
  const cardConv = document.getElementById('cardConv');
  const cardConvSub = document.getElementById('cardConvSub');
  const cardAtv = document.getElementById('cardAtv');
  const cardUptSub = document.getElementById('cardUptSub');

  // Excel-style Barcode Filter Elements
  const barcodeFilterCard = document.getElementById('barcodeFilterCard');
  const barcodeFilterToggle = document.getElementById('barcodeFilterToggle');
  const barcodeFilterStatusText = document.getElementById('barcodeFilterStatusText');
  const btnResetBarcodeFilter = document.getElementById('btnResetBarcodeFilter');
  const btnToggleBarcodeDropdown = document.getElementById('btnToggleBarcodeDropdown');
  const barcodeDropdownBody = document.getElementById('barcodeDropdownBody');
  const barcodeSearchInput = document.getElementById('barcodeSearchInput');
  const btnClearBarcodeSearch = document.getElementById('btnClearBarcodeSearch');
  const chkBarcodeSelectAll = document.getElementById('chkBarcodeSelectAll');
  const btnOnlyNonBags = document.getElementById('btnOnlyNonBags');
  const barcodeCheckboxList = document.getElementById('barcodeCheckboxList');
  const barcodeCountSummary = document.getElementById('barcodeCountSummary');
  const btnApplyBarcodeFilter = document.getElementById('btnApplyBarcodeFilter');

  // Barcode filter state
  let currentRawPosRows = [];
  let availableBarcodes = new Map(); // barcode -> { desc, count, qty, net, isCarryBag }
  let selectedBarcodes = new Set();
  let currentFileName = '';

  // Salesman filter DOM elements
  const salesmanFilterCard = document.getElementById('salesmanFilterCard');
  const salesmanFilterToggle = document.getElementById('salesmanFilterToggle');
  const salesmanFilterStatusText = document.getElementById('salesmanFilterStatusText');
  const btnResetSalesmanFilter = document.getElementById('btnResetSalesmanFilter');
  const btnToggleSalesmanDropdown = document.getElementById('btnToggleSalesmanDropdown');
  const salesmanSearchInput = document.getElementById('salesmanSearchInput');
  const btnClearSalesmanSearch = document.getElementById('btnClearSalesmanSearch');
  const chkSalesmanSelectAll = document.getElementById('chkSalesmanSelectAll');
  const salesmanCheckboxList = document.getElementById('salesmanCheckboxList');
  const salesmanCountSummary = document.getElementById('salesmanCountSummary');
  const btnApplySalesmanFilter = document.getElementById('btnApplySalesmanFilter');
  const btnOpenNameMapping = document.getElementById('btnOpenNameMapping');
  const salesmanNameModal = document.getElementById('salesmanNameModal');
  const salesmanNameMappingList = document.getElementById('salesmanNameMappingList');
  const btnCloseNameModal = document.getElementById('btnCloseNameModal');
  const btnCancelNameModal = document.getElementById('btnCancelNameModal');
  const btnSaveNameMapping = document.getElementById('btnSaveNameMapping');

  // Pre-configured Default Salesman Map & Display Order (Mock Defaults)
  const DEFAULT_SALESMAN_NAME_MAP = {
    '50000001': 'Executive 01',
    '50157523': 'Executive 02',
    '50195907': 'Executive 03',
    '50000002': 'Executive 04',
    '50170981': 'Executive 05',
    '50000003': 'Executive 06',
    '50151682': 'Executive 07',
    '50201915': 'Executive 08',
    '50195439': 'Executive 09',
    '50188930': 'Executive 10',
    '50105787': 'Executive 11',
    '50000000': 'General'
  };

  const DEFAULT_SALESMAN_ORDER = [
    '50000001',
    '50157523',
    '50195907',
    '50000002',
    '50170981',
    '50000003',
    '50151682',
    '50201915',
    '50195439',
    '50188930',
    '50105787',
    '50000000'
  ];

  const DEFAULT_SALESMAN_TARGET_MAP = {
    '50000001': 12000,
    '50157523': 8000,
    '50195907': 12000,
    '50000002': 11000,
    '50170981': 7000,
    '50000003': 12000,
    '50151682': 9000,
    '50201915': 10000,
    '50195439': 8500,
    '50188930': 12000,
    '50105787': 9500,
    '50000000': 0
  };


  function getLoadedSalesmanMap() {
    try {
      const saved = localStorage.getItem('salesman_name_map');
      if (saved) {
        return Object.assign({}, DEFAULT_SALESMAN_NAME_MAP, JSON.parse(saved));
      }
    } catch (e) {}
    return Object.assign({}, DEFAULT_SALESMAN_NAME_MAP);
  }

  function getLoadedSalesmanTargetMap() {
    try {
      const saved = localStorage.getItem('salesman_target_map');
      if (saved) {
        return Object.assign({}, DEFAULT_SALESMAN_TARGET_MAP, JSON.parse(saved));
      }
    } catch (e) {}
    return Object.assign({}, DEFAULT_SALESMAN_TARGET_MAP);
  }

  // Salesman filter state
  let availableSalesmen = new Map(); // id -> { bills, units, net }
  let selectedSalesmen = new Set();
  let salesmanNameMap = getLoadedSalesmanMap();
  let salesmanTargetMap = getLoadedSalesmanTargetMap();



  // Helper: Format numbers with Indian rupee / commas
  function formatINR(val) {
    const num = Math.round(Number(val) || 0);
    return '₹' + num.toLocaleString('en-IN');
  }

  function getVal(key) {
    const el = document.getElementById('f_' + key);
    if (!el) return '';
    return el.value;
  }

  function setVal(key, val, isExtracted = false) {
    const el = document.getElementById('f_' + key);
    if (!el) return;
    el.value = val;
    if (isExtracted) {
      el.classList.add('field-extracted');
      el.classList.remove('field-manual');
    }
  }

  function getNum(key) {
    const v = getVal(key);
    const n = parseFloat(String(v).replace(/[^0-9.-]/g, ''));
    return isNaN(n) ? 0 : n;
  }

  // Load Saved or Sample Data
  function loadInitialData() {
    const saved = localStorage.getItem('retail_kpi_store_data');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        populateForm(parsed);
        return;
      } catch (e) {
        console.warn('Error reading saved data:', e);
      }
    }
    populateForm(sampleData);
  }

  function populateForm(data) {
    fieldIds.forEach(id => {
      if (data[id] !== undefined) {
        setVal(id, data[id]);
      }
    });
    recalculateAll();
  }

  // Recalculate Derived KPIs
  function recalculateAll() {
    const sale = getNum('sale');
    const bill = getNum('bill');
    const unit = getNum('unit');
    const ff = getNum('ff');
    const target = getNum('target');

    // Daily KPIs
    if (bill > 0) {
      const upt = unit / bill;
      const atv = Math.round(sale / bill);
      setVal('upt', upt.toFixed(1));
      setVal('atv', atv);
    }
    if (unit > 0) {
      const auv = Math.round(sale / unit);
      setVal('auv', auv);
    }
    if (ff > 0 && bill > 0) {
      const conv = Math.round((bill / ff) * 100);
      setVal('conv', conv + '%');
    }

    // MTD KPIs
    const mtdTgt = getNum('mtd_tgt');
    const mtdSale = getNum('mtd_sale');
    const mtdBill = getNum('mtd_bill');
    const mtdQty = getNum('mtd_qty');
    const mtdFf = getNum('mtd_ff');

    if (mtdTgt > 0) {
      const mtdPct = Math.round((mtdSale / mtdTgt) * 100);
      setVal('mtd_sale_pct', mtdPct + '%');
    }
    if (mtdBill > 0) {
      const mtdUpt = (mtdQty / mtdBill).toFixed(2);
      const mtdAtv = Math.round(mtdSale / mtdBill);
      setVal('mtd_upt', mtdUpt);
      setVal('mtd_atv', mtdAtv);
    }
    if (mtdQty > 0) {
      const mtdAuv = Math.round(mtdSale / mtdQty);
      setVal('mtd_auv', mtdAuv);
    }
    if (mtdFf > 0 && mtdBill > 0) {
      const mtdConv = Math.round((mtdBill / mtdFf) * 100);
      setVal('mtd_conv', mtdConv + '%');
    }

    // Attachments & Category %
    if (bill > 0) {
      const battery = getNum('battery');
      setVal('battery_attach', Math.round((battery / bill) * 100) + '%');

      const bubble = getNum('bubble');
      setVal('bubble_attach', Math.round((bubble / bill) * 100) + '%');

      const multibill = getNum('multibill');
      setVal('multibill_pct', Math.round((multibill / bill) * 100) + '%');

      const loyalty = getNum('loyalty');
      setVal('loyalty_pct', Math.round((loyalty / bill) * 100) + '%');
    }

    if (sale > 0) {
      const hotspot = getNum('hotspot_sale');
      setVal('hotspot_pct', Math.round((hotspot / sale) * 100) + '%');

      const selfDemo = getNum('self_demo_sale');
      setVal('self_demo_pct', Math.round((selfDemo / sale) * 100) + '%');

      const core = getNum('core');
      setVal('core_pct', Math.round((core / sale) * 100) + '%');

      const mii = getNum('mii_sale');
      setVal('mii_sale_pct', Math.round((mii / sale) * 100) + '%');
    }

    // LFL% (Like for like growth)
    const mtdLySale = getNum('mtd_ly_sale');
    if (mtdLySale > 0 && mtdSale > 0) {
      const lfl = Math.round(((mtdSale - mtdLySale) / mtdLySale) * 100);
      setVal('lfl_pct', lfl + '%');
    }

    // Update Dashboard Mini-Cards
    updateDashboardCards();

    // Generate output text
    generateReport();

    // Auto-save to localStorage
    saveCurrentData();
  }

  function updateDashboardCards() {
    const sale = getNum('sale');
    const target = getNum('target');
    const bill = getNum('bill');
    const ff = getNum('ff');
    const mtdSale = getNum('mtd_sale');
    const mtdTgt = getNum('mtd_tgt');

    const todayAchPct = target > 0 ? ((sale / target) * 100).toFixed(1) : 0;
    cardTodayAch.textContent = todayAchPct + '%';
    barTodayAch.style.width = Math.min(todayAchPct, 100) + '%';
    cardTodaySub.textContent = `${formatINR(sale)} / ${formatINR(target)}`;

    const mtdAchPct = mtdTgt > 0 ? ((mtdSale / mtdTgt) * 100).toFixed(1) : 0;
    cardMtdAch.textContent = mtdAchPct + '%';
    barMtdAch.style.width = Math.min(mtdAchPct, 100) + '%';
    cardMtdSub.textContent = `${formatINR(mtdSale)} / ${formatINR(mtdTgt)}`;

    const convPct = ff > 0 ? ((bill / ff) * 100).toFixed(1) : 0;
    cardConv.textContent = convPct + '%';
    cardConvSub.textContent = `${bill} Bills / ${ff} Footfall`;

    cardAtv.textContent = formatINR(getVal('atv') || (bill > 0 ? Math.round(sale / bill) : 0));
    cardUptSub.textContent = `UPT: ${getVal('upt')} | AUV: ${formatINR(getVal('auv'))}`;
  }

  // Generate Exact Formatted Report
  function generateReport() {
    const formatMode = document.querySelector('input[name="formatMode"]:checked')?.value || 'tabs';
    const sep = formatMode === 'tabs' ? '\t' : '  -  ';

    // Build lines strictly adhering to the user's requested template structure:
    const lines = [
      `Store${sep}${getVal('store')}`,
      `Date-${sep}${getVal('date')}`,
      `MONTH TRGT -${sep}${getVal('month_target')}`,
      `TARGET-${sep}${getVal('target')}`,
      `SALE-${sep}${getVal('sale')}`,
      `FF-${sep}${getVal('ff')}`,
      `BILL-${sep}${getVal('bill')}`,
      `UNIT-${sep}${getVal('unit')}`,
      `UPT-${sep}${getVal('upt')}`,
      `ATV-${sep}${getVal('atv')}`,
      `AUV-${sep}${getVal('auv')}`,
      `CONV%-${sep}${getVal('conv')}`,
      ` \t `,
      `MTD TGT-${sep}${getVal('mtd_tgt')}`,
      `MTD SALE-${sep}${getVal('mtd_sale')}`,
      `MTD SALE%-${sep}${getVal('mtd_sale_pct')}`,
      `MTD FF-${sep}${getVal('mtd_ff')}`,
      `MTD BILL-${sep}${getVal('mtd_bill')}`,
      `MTD QTY-${sep}${getVal('mtd_qty')}`,
      `MTD UPT-${sep}${getVal('mtd_upt')}`,
      `MTD ATV-${sep}${getVal('mtd_atv')}`,
      `MTD AUV-${sep}${getVal('mtd_auv')}`,
      `MTD CONV%-${sep}${getVal('mtd_conv')}`,
      ` \t `,
      `ATVD SALE-${sep}${getVal('atvd_sale')}`,
      `ATVD QTY-${sep}${getVal('atvd_qty')}`,
      `ATVD CONTRI%-${sep}${getVal('atvd_contri_pct')}`,
      `MTD ATVD CONTRI%-${sep}${getVal('mtd_atvd_contri_pct')}`,
      ` \t `,
      `HOT SPOT SALE-${sep}${getVal('hotspot_sale')}`,
      `HOT SPOT %-${sep}${getVal('hotspot_pct')}`,
      `MTD HOTSPOT %-${sep}${getVal('mtd_hotspot_pct')}`,
      ` \t `,
      `SELF DEMO SALE-${sep}${getVal('self_demo_sale')}`,
      `SELF DEMO %-${sep}${getVal('self_demo_pct')}`,
      `MTD SELF DEMO %-${sep}${getVal('mtd_self_demo_pct')}`,
      ` \t `,
      `CORE-${sep}${getVal('core')}`,
      `CORE %-${sep}${getVal('core_pct')}`,
      `MTD CORE SALE-${sep}${getVal('mtd_core_sale')}`,
      `MTD CORE DEMO%-${sep}${getVal('mtd_core_demo_pct')}`,
      ` \t `,
      `MII SALE -${sep}${getVal('mii_sale')}`,
      `MII SALE %-${sep}${getVal('mii_sale_pct')}`,
      `MTD MII SALE-${sep}${getVal('mtd_mii_sale')}`,
      `MTD MII SALE %-${sep}${getVal('mtd_mii_sale_pct')}`,
      ` \t `,
      `BATTERY-${sep}${getVal('battery')}`,
      `BATTERY ATTACH%-${sep}${getVal('battery_attach')}`,
      `MTD BATTERY ATTACH%-${sep}${getVal('mtd_battery_attach')}`,
      ` \t `,
      `BUBBLE STKS-${sep}${getVal('bubble')}`,
      `BUBBLE STKS ATTACH%-${sep}${getVal('bubble_attach')}`,
      `MTD BUBBLE STKS ATTACH%-${sep}${getVal('mtd_bubble_attach')}`,
      ` \t `,
      `MULTI BILL TRGT-${sep}${getVal('multibill_tgt')}`,
      `MULTI BILL -${sep}${getVal('multibill')}`,
      `MULTI BILL %-${sep}${getVal('multibill_pct')}`,
      ` \t `,
      `LOYALTY-${sep}${getVal('loyalty')}`,
      `LOYALTY %-${sep}${getVal('loyalty_pct')}`,
      `MTD LOYALTY-${sep}${getVal('mtd_loyalty')}`,
      `MTD LOYALTY%-${sep}${getVal('mtd_loyalty_pct')}`,
      ` \t `,
      `RED CARD-${sep}${getVal('redcard')}`,
      `MTD RED CARD-${sep}${getVal('mtd_redcard')}`,
      ` \t `,
      `MTD LAST YR SALE-${sep}${getVal('mtd_ly_sale')}`,
      `LY SALE-${sep}${getVal('ly_sale')}`,
      `LFL%-${sep}${getVal('lfl_pct')}`,
      ` \t `,
      `TOMORROW TRGT-${sep}${getVal('tomorrow_tgt')}`,
      `NET PLAY TGT-${sep}${getVal('tomorrow_netplay_tgt')}`,
      ` \t `,
      `TELECALLING TRGT-${sep}${getVal('tele_tgt')}`,
      `NO OF CALLS-${sep}${getVal('tele_calls')}`,
      `TELECALLING SALE-${sep}${getVal('tele_sale')}`,
      `TELECALLING BILLS-${sep}${getVal('tele_bills')}`,
      `IN STORE SALE-${sep}${getVal('instore_sale')}`,
      `IN STORE BILL-${sep}${getVal('instore_bill')}`,
      `HOME DELIVERY SALE-${sep}${getVal('hd_sale')}`,
      `HOME DELIVERY BILL-${sep}${getVal('hd_bill')}`,
      `MTD CALL-${sep}${getVal('mtd_call')}`,
      `MTD BILL-${sep}${getVal('mtd_tele_bill')}`,
      `MTD SALE-${sep}${getVal('mtd_tele_sale')}`,
      ` \t `,
      `Net Play Target-${sep}${getVal('netplay_tgt')}`,
      `Net Play Sale-${sep}${getVal('netplay_sale')}`,
      `MTD Net Play Target-${sep}${getVal('mtd_netplay_tgt')}`,
      `MTD Net Play Sale-${sep}${getVal('mtd_netplay_sale')}`
    ];

    reportOutput.textContent = lines.join('\n');
  }

  function saveCurrentData() {
    const data = {};
    fieldIds.forEach(id => {
      data[id] = getVal(id);
    });
    localStorage.setItem('retail_kpi_store_data', JSON.stringify(data));
  }

  // File Handling & Parsing
  dropzone.addEventListener('click', () => fileInput.click());

  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      handleUploadedFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files.length) {
      handleUploadedFile(e.target.files[0]);
    }
  });

  btnClearFile.addEventListener('click', () => {
    fileUploadStatus.classList.add('hidden');
    barcodeFilterCard.classList.add('hidden');
    salesmanFilterCard.classList.add('hidden');
    fileInput.value = '';
    currentRawPosRows = [];
    availableBarcodes.clear();
    selectedBarcodes.clear();
    barcodeCheckboxList.innerHTML = '';
    availableSalesmen.clear();
    selectedSalesmen.clear();
    salesmanCheckboxList.innerHTML = '';
  });

  // Toggle Barcode Filter Dropdown (Expand / Collapse)
  barcodeFilterToggle.addEventListener('click', (e) => {
    // Prevent toggle if clicking on action buttons
    if (e.target.closest('.filter-header-actions')) return;
    barcodeFilterCard.classList.toggle('collapsed');
  });

  btnToggleBarcodeDropdown.addEventListener('click', () => {
    barcodeFilterCard.classList.toggle('collapsed');
  });

  btnApplyBarcodeFilter.addEventListener('click', () => {
    barcodeFilterCard.classList.add('collapsed');
    showToast('Barcode filter applied');
  });

  // Search in Barcode filter (Excel-like instant search)
  barcodeSearchInput.addEventListener('input', (e) => {
    const q = e.target.value.trim().toLowerCase();
    btnClearBarcodeSearch.classList.toggle('hidden', q.length === 0);
    renderBarcodeList(q);
  });

  btnClearBarcodeSearch.addEventListener('click', () => {
    barcodeSearchInput.value = '';
    btnClearBarcodeSearch.classList.add('hidden');
    renderBarcodeList('');
    barcodeSearchInput.focus();
  });

  // Master (Select All) Checkbox
  chkBarcodeSelectAll.addEventListener('change', () => {
    const isChecked = chkBarcodeSelectAll.checked;
    const q = barcodeSearchInput.value.trim().toLowerCase();

    availableBarcodes.forEach((data, bc) => {
      // If searching, only affect items matching the search query (Excel behavior)
      const matchesSearch = !q || bc.toLowerCase().includes(q) || (data.desc || '').toLowerCase().includes(q);
      if (matchesSearch) {
        if (isChecked) {
          selectedBarcodes.add(bc);
        } else {
          selectedBarcodes.delete(bc);
        }
      }
    });

    renderBarcodeList(q);
    filterAndApplyBarcodeData(false);
  });

  // Quick Action: Exclude Carry Bags
  btnOnlyNonBags.addEventListener('click', () => {
    availableBarcodes.forEach((data, bc) => {
      if (data.isCarryBag) {
        selectedBarcodes.delete(bc);
      }
    });
    renderBarcodeList(barcodeSearchInput.value.trim().toLowerCase());
    filterAndApplyBarcodeData(false);
    showToast('Carry bags excluded from calculation');
  });

  // Reset / Show All Barcodes
  btnResetBarcodeFilter.addEventListener('click', () => {
    selectedBarcodes = new Set(availableBarcodes.keys());
    chkBarcodeSelectAll.checked = true;
    barcodeSearchInput.value = '';
    btnClearBarcodeSearch.classList.add('hidden');
    renderBarcodeList('');
    filterAndApplyBarcodeData(false);
    showToast('Reset: All barcodes selected');
  });

  // ===== Salesman Filter Event Handlers =====

  salesmanFilterToggle.addEventListener('click', (e) => {
    if (e.target.closest('.filter-header-actions')) return;
    salesmanFilterCard.classList.toggle('collapsed');
  });

  btnToggleSalesmanDropdown.addEventListener('click', () => {
    salesmanFilterCard.classList.toggle('collapsed');
  });

  btnApplySalesmanFilter.addEventListener('click', () => {
    salesmanFilterCard.classList.add('collapsed');
    showToast('Salesman filter applied');
  });

  salesmanSearchInput.addEventListener('input', (e) => {
    const q = e.target.value.trim().toLowerCase();
    btnClearSalesmanSearch.classList.toggle('hidden', q.length === 0);
    renderSalesmanList(q);
  });

  btnClearSalesmanSearch.addEventListener('click', () => {
    salesmanSearchInput.value = '';
    btnClearSalesmanSearch.classList.add('hidden');
    renderSalesmanList('');
    salesmanSearchInput.focus();
  });

  chkSalesmanSelectAll.addEventListener('change', () => {
    const isChecked = chkSalesmanSelectAll.checked;
    const q = salesmanSearchInput.value.trim().toLowerCase();
    availableSalesmen.forEach((data, id) => {
      const name = (salesmanNameMap[id] || '').toLowerCase();
      const matches = !q || id.toLowerCase().includes(q) || name.includes(q);
      if (matches) {
        if (isChecked) selectedSalesmen.add(id);
        else selectedSalesmen.delete(id);
      }
    });
    renderSalesmanList(q);
    applyBothFilters();
  });

  btnResetSalesmanFilter.addEventListener('click', () => {
    selectedSalesmen = new Set(availableSalesmen.keys());
    chkSalesmanSelectAll.checked = true;
    salesmanSearchInput.value = '';
    btnClearSalesmanSearch.classList.add('hidden');
    renderSalesmanList('');
    applyBothFilters();
    showToast('Reset: All salesmen selected');
  });

  // Name Mapping Modal
  btnOpenNameMapping.addEventListener('click', () => {
    openNameMappingModal();
  });
  btnCloseNameModal.addEventListener('click', () => salesmanNameModal.classList.add('hidden'));
  btnCancelNameModal.addEventListener('click', () => salesmanNameModal.classList.add('hidden'));
  salesmanNameModal.addEventListener('click', (e) => {
    if (e.target === salesmanNameModal) salesmanNameModal.classList.add('hidden');
  });

  btnSaveNameMapping.addEventListener('click', () => {
    // Collect all input values from modal rows
    const rows = salesmanNameMappingList.querySelectorAll('.salesman-name-row');
    rows.forEach(row => {
      const id = row.dataset.salesmanId;
      const nameInput = row.querySelector('.salesman-name-input');
      const targetInput = row.querySelector('.salesman-target-input');
      const name = nameInput ? nameInput.value.trim() : '';
      const tgtVal = targetInput && targetInput.value !== '' ? parseFloat(targetInput.value) : undefined;

      if (name) {
        salesmanNameMap[id] = name;
      } else {
        delete salesmanNameMap[id];
      }

      if (tgtVal !== undefined && !isNaN(tgtVal)) {
        salesmanTargetMap[id] = tgtVal;
      } else {
        delete salesmanTargetMap[id];
      }
    });

    // Persist
    localStorage.setItem('salesman_name_map', JSON.stringify(salesmanNameMap));
    localStorage.setItem('salesman_target_map', JSON.stringify(salesmanTargetMap));
    salesmanNameModal.classList.add('hidden');
    renderSalesmanList(salesmanSearchInput.value.trim().toLowerCase());
    showToast('Salesman names & targets saved!');
  });

  function openNameMappingModal() {
    salesmanNameMappingList.innerHTML = '';
    const allIds = new Set([
      ...DEFAULT_SALESMAN_ORDER,
      ...Object.keys(salesmanNameMap),
      ...availableSalesmen.keys()
    ]);

    const sortedIds = [...allIds].sort((a, b) => {
      const idxA = DEFAULT_SALESMAN_ORDER.indexOf(a);
      const idxB = DEFAULT_SALESMAN_ORDER.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b);
    });

    sortedIds.forEach((id) => {
      const currentName = salesmanNameMap[id] || '';
      const currentTarget = salesmanTargetMap[id] !== undefined
        ? salesmanTargetMap[id]
        : (DEFAULT_SALESMAN_TARGET_MAP[id] !== undefined ? DEFAULT_SALESMAN_TARGET_MAP[id] : '');
      const smData = availableSalesmen.get(id);
      const stats = smData ? ` • ${smData.receipts.size} bills` : '';
      const row = document.createElement('div');
      row.className = 'salesman-name-row';
      row.dataset.salesmanId = id;
      row.innerHTML = `
        <input type="text" class="salesman-name-input" placeholder="Enter name…" value="${currentName}" title="Type a name for ID ${id}">
        <div class="salesman-row-equals">=</div>
        <div class="salesman-id-label" title="${id}${stats}">${id}</div>
        <input type="number" class="salesman-target-input" placeholder="Target ₹" value="${currentTarget}" title="Target for ID ${id}">
      `;
      salesmanNameMappingList.appendChild(row);
    });
    salesmanNameModal.classList.remove('hidden');
    // Focus first empty input
    const firstEmpty = salesmanNameMappingList.querySelector('.salesman-name-input:not([value])') ||
                       salesmanNameMappingList.querySelector('.salesman-name-input');
    if (firstEmpty) setTimeout(() => firstEmpty.focus(), 50);
  }




  function handleUploadedFile(file) {
    uploadedFileName.textContent = file.name;
    uploadSummaryText.textContent = 'Processing file...';
    fileUploadStatus.classList.remove('hidden');

    const reader = new FileReader();

    reader.onload = function(e) {
      const buffer = e.target.result;
      const textDecoder = new TextDecoder('utf-8');
      const textStart = textDecoder.decode(new Uint8Array(buffer.slice(0, 1024)));

      if (textStart.includes('<Table') || textStart.includes('<table') || textStart.includes('<TR>') || textStart.includes('<!DOCTYPE')) {
        parseHtmlPosTable(buffer, file.name);
      } else {
        parseSheetJs(buffer, file.name);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function parseHtmlPosTable(buffer, fileName) {
    try {
      const decoder = new TextDecoder('utf-8');
      const htmlText = decoder.decode(buffer);
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlText, 'text/html');
      const table = doc.querySelector('table');

      if (!table) {
        throw new Error('No table element found in HTML file');
      }

      const rows = Array.from(table.querySelectorAll('tr'));
      if (rows.length < 2) {
        throw new Error('Table does not contain enough rows');
      }

      const headerCells = Array.from(rows[0].querySelectorAll('td, th'));
      const headers = headerCells.map(c => c.textContent.trim().toUpperCase());

      const dataRows = [];
      for (let i = 1; i < rows.length; i++) {
        const cells = Array.from(rows[i].querySelectorAll('td'));
        if (cells.length === headers.length) {
          const rowObj = {};
          headers.forEach((h, idx) => {
            rowObj[h] = cells[idx].textContent.trim();
          });
          dataRows.push(rowObj);
        }
      }

      processPosRows(dataRows, fileName);
    } catch (err) {
      console.error('HTML Table parsing error:', err);
      parseSheetJs(buffer, fileName);
    }
  }

  function parseSheetJs(buffer, fileName) {
    try {
      if (typeof XLSX === 'undefined') {
        throw new Error('SheetJS XLSX library not loaded');
      }
      const data = new Uint8Array(buffer);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

      if (jsonData && jsonData.length > 0) {
        const normalizedRows = jsonData.map(row => {
          const norm = {};
          Object.keys(row).forEach(k => {
            norm[k.trim().toUpperCase()] = String(row[k]).trim();
          });
          return norm;
        });

        if (normalizedRows[0]['RECEIPTNO'] !== undefined || normalizedRows[0]['NET'] !== undefined) {
          processPosRows(normalizedRows, fileName);
        } else {
          processConsolidatedSheet(jsonData, fileName);
        }
      } else {
        throw new Error('No data rows found in worksheet');
      }
    } catch (err) {
      console.error('SheetJS parse error:', err);
      uploadSummaryText.textContent = 'Could not parse file. Please verify file format.';
    }
  }

  // Setup Excel-style Barcode Filter
  function setupBarcodeFilter(rows, fileName) {
    availableBarcodes = new Map();

    rows.forEach(r => {
      const bc = (r['BARCODE'] || r['ARTICLE'] || r['EAN'] || r['SKU'] || '').trim();
      if (!bc) return;
      const desc = (r['ARTICLEDESC'] || r['ITEMDESC'] || r['DESCRIPTION'] || 'Item').trim();
      const qty = parseFloat(r['QTY'] || '1') || 0;
      const net = parseFloat(r['NET'] || r['NETAMOUNT'] || '0') || 0;
      const descUpper = desc.toUpperCase();
      const isCarryBag = descUpper.includes('BAG') && (descUpper.includes('CARRY') || descUpper.includes('PAPER HAM'));

      if (!availableBarcodes.has(bc)) {
        availableBarcodes.set(bc, { desc, count: 0, qty: 0, net: 0, isCarryBag });
      }
      const bObj = availableBarcodes.get(bc);
      bObj.count++;
      bObj.qty += qty;
      bObj.net += net;
    });

    // Default: select ALL barcodes (like Excel autofilter)
    selectedBarcodes = new Set(availableBarcodes.keys());
    chkBarcodeSelectAll.checked = true;
    barcodeSearchInput.value = '';
    btnClearBarcodeSearch.classList.add('hidden');

    renderBarcodeList('');
    barcodeFilterCard.classList.remove('hidden');
    barcodeFilterCard.classList.remove('collapsed');

    filterAndApplyBarcodeData(true);
  }

  // Render Barcode Checkbox Scrollable List
  function renderBarcodeList(query = '') {
    barcodeCheckboxList.innerHTML = '';
    let visibleCount = 0;
    let visibleSelectedCount = 0;

    availableBarcodes.forEach((data, bc) => {
      const descLower = (data.desc || '').toLowerCase();
      const bcLower = bc.toLowerCase();
      const matches = !query || bcLower.includes(query) || descLower.includes(query);

      if (!matches) return;

      visibleCount++;
      const isChecked = selectedBarcodes.has(bc);
      if (isChecked) visibleSelectedCount++;

      const row = document.createElement('label');
      row.className = `excel-item-row ${isChecked ? '' : 'unselected'}`;
      row.dataset.bc = bc;

      row.innerHTML = `
        <div class="excel-item-left">
          <input type="checkbox" class="barcode-chk" value="${bc}" ${isChecked ? 'checked' : ''}>
          <span class="excel-item-barcode">${bc}</span>
          <span class="excel-item-desc" title="${data.desc}">${data.desc}</span>
        </div>
        <span class="excel-item-stats">${data.qty} Qty • ₹${Math.round(data.net).toLocaleString('en-IN')}</span>
      `;

      const chk = row.querySelector('.barcode-chk');
      chk.addEventListener('change', () => {
        if (chk.checked) {
          selectedBarcodes.add(bc);
          row.classList.remove('unselected');
        } else {
          selectedBarcodes.delete(bc);
          row.classList.add('unselected');
        }
        filterAndApplyBarcodeData(false);
      });

      barcodeCheckboxList.appendChild(row);
    });

    barcodeCountSummary.textContent = `${visibleCount} item${visibleCount === 1 ? '' : 's'} displayed`;
    chkBarcodeSelectAll.checked = visibleCount > 0 && visibleSelectedCount === visibleCount;
    chkBarcodeSelectAll.indeterminate = visibleSelectedCount > 0 && visibleSelectedCount < visibleCount;
  }

  // Filter Data by Selected Barcodes & Recalculate
  function filterAndApplyBarcodeData(isInitial = false) {
    const totalCount = availableBarcodes.size;
    const selectedCount = selectedBarcodes.size;

    if (selectedCount === totalCount) {
      barcodeFilterStatusText.textContent = `All Barcodes Selected (${totalCount} SKUs)`;
    } else if (selectedCount === 0) {
      barcodeFilterStatusText.textContent = `0 of ${totalCount} Barcodes Selected (No items)`;
    } else {
      barcodeFilterStatusText.textContent = `Filtered: ${selectedCount} of ${totalCount} SKUs Selected`;
    }

    if (selectedCount === 0) {
      setVal('sale', 0, true);
      setVal('bill', 0, true);
      setVal('unit', 0, true);
      setVal('battery', 0, true);
      setVal('bubble', 0, true);
      setVal('multibill', 0, true);
      setVal('loyalty', 0, true);
      setVal('redcard', 0, true);
      setVal('netplay_sale', 0, true);
      setVal('mii_sale', 0, true);
      uploadSummaryText.textContent = 'No barcodes selected. Please check at least one barcode.';
      recalculateAll();
      return;
    }

    const filteredRows = currentRawPosRows.filter(r => {
      const bc = (r['BARCODE'] || r['ARTICLE'] || r['EAN'] || r['SKU'] || '').trim();
      return selectedBarcodes.has(bc);
    });

    let detectedStore = '';
    let detectedDate = '';
    const receiptsMap = new Map();
    let totalNet = 0;
    let totalMerchUnits = 0;

    let batteryQty = 0;
    let bubbleQty = 0;
    let loyaltyCount = 0;
    let redcardCount = 0;
    let netplayNet = 0;
    let miiNet = 0;

    filteredRows.forEach(r => {
      const storeNo = r['STORENO'] || '';
      if (!detectedStore && storeNo) {
        if (storeNo.toUpperCase().includes('TLX5') || storeNo.toUpperCase().includes('VEGAS')) {
          detectedStore = 'VEGAS';
        } else {
          detectedStore = storeNo;
        }
      }

      const transDate = r['TRANSDATE'] || '';
      if (!detectedDate && transDate) {
        const parts = transDate.split(' ');
        if (parts[0]) detectedDate = parts[0];
      }

      const receiptNo = r['RECEIPTNO'] || r['RECEIPT NO'] || '';
      const desc = (r['ARTICLEDESC'] || r['ITEMDESC'] || r['DESCRIPTION'] || '').toUpperCase();
      const qty = parseFloat(r['QTY'] || '1') || 0;
      const net = parseFloat(r['NET'] || r['NETAMOUNT'] || '0') || 0;

      const isCarryBag = desc.includes('BAG') && (desc.includes('CARRY') || desc.includes('PAPER HAM'));

      totalNet += net;

      if (!isCarryBag) {
        totalMerchUnits += qty;
      }

      if (receiptNo) {
        if (!receiptsMap.has(receiptNo)) {
          receiptsMap.set(receiptNo, { merchCount: 0, items: [] });
        }
        const rcObj = receiptsMap.get(receiptNo);
        if (!isCarryBag) {
          rcObj.merchCount += qty;
        }
        rcObj.items.push(r);
      }

      // Category checks
      if (desc.includes('BATT') || desc.startsWith('AA ') || desc.startsWith('AAA ') || desc.includes('DURACELL')) {
        batteryQty += qty;
      }
      if (desc.includes('BUBBLE') || desc.includes('BUBBLE WAND') || desc.includes('BUBBLE TUBE')) {
        bubbleQty += qty;
      }
      if (desc.includes('HPLUS') || desc.includes('MEMBERSHIP') || desc.includes('LOYALTY')) {
        loyaltyCount += qty;
      }
      if (desc.includes('REDCARD') || desc.includes('RED CARD')) {
        redcardCount += qty;
      }
      if (desc.includes('NETPLAY') || desc.includes('NET PLAY')) {
        netplayNet += net;
      }
      if (desc.includes('MII') || desc.includes('MAKE IN INDIA') || desc.includes('MAKE IT INDIA')) {
        miiNet += net;
      }
    });

    const uniqueBills = receiptsMap.size;
    let multiBillCount = 0;
    receiptsMap.forEach(rc => {
      if (rc.merchCount > 1) {
        multiBillCount++;
      }
    });

    // Populate extracted values into form
    if (detectedStore) setVal('store', detectedStore, true);
    if (detectedDate) setVal('date', detectedDate, true);

    setVal('sale', Math.round(totalNet), true);
    setVal('bill', uniqueBills, true);
    setVal('unit', Math.round(totalMerchUnits), true);

    setVal('battery', Math.round(batteryQty), true);
    setVal('bubble', Math.round(bubbleQty), true);
    setVal('multibill', multiBillCount, true);
    setVal('loyalty', Math.round(loyaltyCount), true);
    setVal('redcard', Math.round(redcardCount), true);
    setVal('netplay_sale', Math.round(netplayNet), true);
    setVal('mii_sale', Math.round(miiNet), true);

    const filterText = selectedCount === totalCount
      ? `All ${totalCount} Barcodes`
      : `${selectedCount} Barcodes Filtered`;

    uploadSummaryText.textContent = `${filterText}: ${uniqueBills} Bills, ${Math.round(totalMerchUnits)} Units, ₹${Math.round(totalNet).toLocaleString('en-IN')}`;

    // Automatically recalculate all formulas
    recalculateAll();

    if (isInitial) {
      showToast(`Loaded ${totalCount} Barcodes from POS file!`);
    }
  }

  // Process POS Receipt Journal Data
  function processPosRows(rows, fileName) {
    if (!rows || rows.length === 0) return;
    currentRawPosRows = rows;
    currentFileName = fileName;
    salesmanNameMap = getLoadedSalesmanMap();
    setupBarcodeFilter(rows, fileName);
    setupSalesmanFilter(rows);
  }

  // Setup Salesman Filter
  function setupSalesmanFilter(rows) {
    availableSalesmen = new Map();

    rows.forEach(r => {
      const code = (r['SALESMANCODE'] || r['SALESMAN CODE'] || r['SALESMANID'] || r['SALESMAN'] || '').trim();
      if (!code || code === '0' || code.toLowerCase() === 'null') return;

      const receiptNo = r['RECEIPTNO'] || r['RECEIPT NO'] || '';
      const qty = parseFloat(r['QTY'] || '1') || 0;
      const net = parseFloat(r['NET'] || r['NETAMOUNT'] || '0') || 0;
      const desc = (r['ARTICLEDESC'] || r['ITEMDESC'] || r['DESCRIPTION'] || '').toUpperCase();
      const isCarryBag = desc.includes('BAG') && (desc.includes('CARRY') || desc.includes('PAPER HAM'));

      if (!availableSalesmen.has(code)) {
        availableSalesmen.set(code, { receipts: new Set(), units: 0, net: 0 });
      }
      const smObj = availableSalesmen.get(code);
      if (receiptNo) smObj.receipts.add(receiptNo);
      if (!isCarryBag) smObj.units += qty;
      smObj.net += net;
    });

    if (availableSalesmen.size === 0) {
      // No SALESMANCODE column found – hide the filter
      salesmanFilterCard.classList.add('hidden');
      return;
    }

    selectedSalesmen = new Set(availableSalesmen.keys());
    chkSalesmanSelectAll.checked = true;
    salesmanSearchInput.value = '';
    btnClearSalesmanSearch.classList.add('hidden');

    renderSalesmanList('');
    salesmanFilterCard.classList.remove('hidden');
    salesmanFilterCard.classList.remove('collapsed');
    updateSalesmanStatusText();
  }

  // Render Salesman Checkbox List
  function renderSalesmanList(query = '') {
    salesmanCheckboxList.innerHTML = '';
    let visibleCount = 0;
    let visibleSelectedCount = 0;

    // Sort by user's preferred order (DEFAULT_SALESMAN_ORDER), then others
    const sortedEntries = [...availableSalesmen.entries()].sort((a, b) => {
      const idxA = DEFAULT_SALESMAN_ORDER.indexOf(a[0]);
      const idxB = DEFAULT_SALESMAN_ORDER.indexOf(b[0]);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a[0].localeCompare(b[0]);
    });

    sortedEntries.forEach(([id, data]) => {
      const name = salesmanNameMap[id] || '';
      const idLower = id.toLowerCase();
      const nameLower = name.toLowerCase();
      const matches = !query || idLower.includes(query) || nameLower.includes(query);
      if (!matches) return;

      visibleCount++;
      const isChecked = selectedSalesmen.has(id);
      if (isChecked) visibleSelectedCount++;

      const bills = data.receipts.size;
      const units = Math.round(data.units);
      const net = Math.round(data.net);

      const row = document.createElement('label');
      row.className = `excel-item-row ${isChecked ? '' : 'unselected'}`;
      row.dataset.smId = id;

      const labelContent = name
        ? `<span class="salesman-item-name">${name}</span> <span class="salesman-item-sep">-</span> <span class="salesman-item-id">${id}</span>`
        : `<span class="salesman-item-id">${id}</span>`;

      row.innerHTML = `
        <div class="excel-item-left">
          <input type="checkbox" class="barcode-chk salesman-chk" value="${id}" ${isChecked ? 'checked' : ''}>
          <span class="salesman-item-num">${visibleCount}.</span>
          <div class="salesman-item-text">
            ${labelContent}
          </div>
        </div>
        <span class="excel-item-stats">${bills} Bills • ${units} Units • ₹${net.toLocaleString('en-IN')}</span>
      `;

      const chk = row.querySelector('.salesman-chk');
      chk.addEventListener('change', () => {
        if (chk.checked) {
          selectedSalesmen.add(id);
          row.classList.remove('unselected');
        } else {
          selectedSalesmen.delete(id);
          row.classList.add('unselected');
        }
        applyBothFilters();
      });

      salesmanCheckboxList.appendChild(row);
    });

    salesmanCountSummary.textContent = `${visibleCount} salesman${visibleCount === 1 ? '' : 's'} displayed`;
    chkSalesmanSelectAll.checked = visibleCount > 0 && visibleSelectedCount === visibleCount;
    chkSalesmanSelectAll.indeterminate = visibleSelectedCount > 0 && visibleSelectedCount < visibleCount;
  }

  function updateSalesmanStatusText() {
    const total = availableSalesmen.size;
    const selected = selectedSalesmen.size;
    if (selected === total) {
      salesmanFilterStatusText.textContent = `All ${total} Salesmen Selected`;
    } else if (selected === 0) {
      salesmanFilterStatusText.textContent = `No Salesmen Selected`;
    } else {
      // Show names if available
      const names = [...selectedSalesmen].map(id => salesmanNameMap[id] || id).join(', ');
      const label = names.length > 40 ? `${selected} of ${total} Selected` : names;
      salesmanFilterStatusText.textContent = `Showing: ${label}`;
    }
  }

  // Apply both barcode + salesman filters and recalculate
  function applyBothFilters() {
    updateSalesmanStatusText();

    // If salesman filter active, pre-filter rows to selected salesmen only
    let rowsForBarcode = currentRawPosRows;

    if (availableSalesmen.size > 0 && selectedSalesmen.size < availableSalesmen.size) {
      rowsForBarcode = currentRawPosRows.filter(r => {
        const code = (r['SALESMANCODE'] || r['SALESMAN CODE'] || r['SALESMANID'] || r['SALESMAN'] || '').trim();
        if (!code || code === '0') {
          // If 'Non' (50000000) is among selected salesmen, include unassigned rows
          return selectedSalesmen.has('50000000');
        }
        return selectedSalesmen.has(code);
      });
    }


    // Re-do the barcode-level filter on the salesman-filtered rows
    // Rebuild availableBarcodes from the salesman-filtered rows
    const prevSelectedBarcodes = new Set(selectedBarcodes);
    availableBarcodes = new Map();
    rowsForBarcode.forEach(r => {
      const bc = (r['BARCODE'] || r['ARTICLE'] || r['EAN'] || r['SKU'] || '').trim();
      if (!bc) return;
      const desc = (r['ARTICLEDESC'] || r['ITEMDESC'] || r['DESCRIPTION'] || 'Item').trim();
      const qty = parseFloat(r['QTY'] || '1') || 0;
      const net = parseFloat(r['NET'] || r['NETAMOUNT'] || '0') || 0;
      const descUpper = desc.toUpperCase();
      const isCarryBag = descUpper.includes('BAG') && (descUpper.includes('CARRY') || descUpper.includes('PAPER HAM'));
      if (!availableBarcodes.has(bc)) {
        availableBarcodes.set(bc, { desc, count: 0, qty: 0, net: 0, isCarryBag });
      }
      const bObj = availableBarcodes.get(bc);
      bObj.count++;
      bObj.qty += qty;
      bObj.net += net;
    });

    // Maintain previous selection where possible; add new barcodes as selected
    selectedBarcodes = new Set();
    availableBarcodes.forEach((data, bc) => {
      // If it was explicitly selected before, keep it selected; if brand-new, select it
      if (prevSelectedBarcodes.has(bc)) {
        selectedBarcodes.add(bc);
      } else if (!prevSelectedBarcodes.has(bc)) {
        // New barcode after filter change – auto-select
        selectedBarcodes.add(bc);
      }
    });

    renderBarcodeList(barcodeSearchInput.value.trim().toLowerCase());
    filterAndApplyBarcodeData(false);
  }



  // Process Consolidated Sheet (if user uploads a key-value or columnar summary report)
  function processConsolidatedSheet(rows, fileName) {
    let matchCount = 0;
    rows.forEach(r => {
      Object.keys(r).forEach(k => {
        const keyNorm = k.toUpperCase().replace(/[^A-Z0-9]/g, '');
        const val = String(r[k]).trim();
        
        fieldIds.forEach(fid => {
          const fidNorm = fid.toUpperCase().replace(/[^A-Z0-9]/g, '');
          if (keyNorm === fidNorm && val !== '') {
            setVal(fid, val, true);
            matchCount++;
          }
        });
      });
    });

    uploadSummaryText.textContent = `Imported ${matchCount} metrics from ${fileName}`;
    recalculateAll();
    showToast(`Loaded ${matchCount} metrics from Excel`);
  }

  // Toast Notification
  function showToast(msg) {
    toast.textContent = msg;
    toast.classList.add('show');
    setTimeout(() => {
      toast.classList.remove('show');
    }, 2800);
  }

  // Copy to Clipboard
  btnCopyReport.addEventListener('click', async () => {
    try {
      const text = reportOutput.textContent;
      await navigator.clipboard.writeText(text);
      showToast('Report copied to clipboard! Ready to paste in WhatsApp.');
    } catch (err) {
      const textarea = document.createElement('textarea');
      textarea.value = reportOutput.textContent;
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
      showToast('Report copied to clipboard!');
    }
  });

  // Download as TXT
  btnDownloadTxt.addEventListener('click', () => {
    const text = reportOutput.textContent;
    const store = getVal('store') || 'STORE';
    const date = getVal('date') || 'REPORT';
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Closing_Report_${store}_${date}.txt`;
    a.click();
    URL.revokeObjectURL(url);
    showToast('Report downloaded as .txt');
  });

  // Download as Excel (.xlsx) with 3 Sheets:
  // Sheet 1: Raw ReceiptJournalReport
  // Sheet 2: Salesman Data (Achievement table matching photo)
  // Sheet 3: Closing Report (Store KPI closing summary)
  btnDownloadXlsx.addEventListener('click', () => {
    if (typeof XLSX === 'undefined') {
      alert('Excel export library is not loaded');
      return;
    }

    const wb = XLSX.utils.book_new();

    // ==========================================
    // SHEET 1: ReceiptJournalReport (Raw POS data)
    // ==========================================
    let ws1;
    if (currentRawPosRows && currentRawPosRows.length > 0) {
      ws1 = XLSX.utils.json_to_sheet(currentRawPosRows);
    } else {
      ws1 = XLSX.utils.aoa_to_sheet([
        ['No ReceiptJournalReport file uploaded yet.'],
        ['Upload a POS ReceiptJournalReport (.xls/.html) to view the raw transaction journal.']
      ]);
    }
    XLSX.utils.book_append_sheet(wb, ws1, 'ReceiptJournalReport');

    // ==========================================
    // SHEET 2: Salesman Data (Matching the photo!)
    // ==========================================
    const allSmIds = new Set([
      ...DEFAULT_SALESMAN_ORDER,
      ...availableSalesmen.keys()
    ]);

    const sortedSmIds = [...allSmIds].sort((a, b) => {
      const idxA = DEFAULT_SALESMAN_ORDER.indexOf(a);
      const idxB = DEFAULT_SALESMAN_ORDER.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b);
    });

    let sumTarget = 0;
    let sumSale = 0;
    let sumBills = 0;
    let sumQty = 0;

    const smRows = [];

    sortedSmIds.forEach(id => {
      const data = availableSalesmen.get(id);
      const bills = data ? data.receipts.size : 0;
      const qty = data ? Math.round(data.units) : 0;
      const sale = data ? Math.round(data.net) : 0;
      const name = salesmanNameMap[id] || id;

      // When a POS file is loaded, only list salesmen active on that day
      if (availableSalesmen.size > 0 && (!data || (bills === 0 && sale === 0 && qty === 0))) {
        return;
      }

      const target = Number(salesmanTargetMap[id] !== undefined
        ? salesmanTargetMap[id]
        : (DEFAULT_SALESMAN_TARGET_MAP[id] || 0));

      const upt = bills > 0 ? Number((qty / bills).toFixed(2)) : 0;
      const atv = bills > 0 ? Math.round(sale / bills) : 0;
      const auv = qty > 0 ? Math.round(sale / qty) : 0;

      sumTarget += target;
      sumSale += sale;
      sumBills += bills;
      sumQty += qty;

      smRows.push([
        name,
        target,
        sale,
        bills,
        qty,
        upt,
        atv,
        auv
      ]);
    });

    // Fallback sample rows if no POS report was loaded yet
    if (smRows.length === 0) {
      const sampleSmRows = [
        ['Executive 01', 12000, 11500, 6, 15, 2.50, 1916, 766],
        ['Executive 02', 8000, 8200, 4, 10, 2.50, 2050, 820],
        ['Executive 03', 12000, 13100, 7, 18, 2.57, 1871, 727],
        ['Executive 04', 11000, 10800, 5, 12, 2.40, 2160, 900],
        ['Executive 05', 7000, 7400, 4, 9, 2.25, 1850, 822],
        ['Executive 06', 12000, 12500, 6, 14, 2.33, 2083, 892],
        ['Executive 07', 9000, 9200, 5, 11, 2.20, 1840, 836]
      ];

      sampleSmRows.forEach(r => {
        sumTarget += r[1];
        sumSale += r[2];
        sumBills += r[3];
        sumQty += r[4];
        smRows.push(r);
      });
    }

    const totalUpt = sumBills > 0 ? Number((sumQty / sumBills).toFixed(2)) : 0;
    const totalAtv = sumBills > 0 ? Math.round(sumSale / sumBills) : 0;
    const totalAuv = sumQty > 0 ? Math.round(sumSale / sumQty) : 0;

    const sheet2Aoa = [
      ['', sumTarget, 'Achievement', '', '', '', '', ''],
      ['Name', 'Target', 'Sale', 'Bills', 'Qty', 'UPT', 'ATV', 'AUV'],
      ...smRows,
      ['Total', sumTarget, sumSale, sumBills, sumQty, totalUpt, totalAtv, totalAuv]
    ];

    const ws2 = XLSX.utils.aoa_to_sheet(sheet2Aoa);

    // Merge 'Achievement' header across columns C to H (index 2 to 7) in Row 1 (index 0)
    ws2['!merges'] = [
      { s: { r: 0, c: 2 }, e: { r: 0, c: 7 } }
    ];

    ws2['!cols'] = [
      { wch: 18 }, // Name
      { wch: 12 }, // Target
      { wch: 12 }, // Sale
      { wch: 10 }, // Bills
      { wch: 10 }, // Qty
      { wch: 10 }, // UPT
      { wch: 12 }, // ATV
      { wch: 12 }  // AUV
    ];

    XLSX.utils.book_append_sheet(wb, ws2, 'Salesman Data');

    // ==========================================
    // SHEET 3: Closing Report (Store KPI Closing Summary)
    // ==========================================
    const reportText = reportOutput.textContent;
    const lines = reportText.split('\n');
    const sheet3Data = [
      ['Metric', 'Value']
    ];

    lines.forEach(l => {
      if (l.trim() === '') {
        sheet3Data.push(['', '']);
      } else {
        const parts = l.split('\t');
        if (parts.length >= 2) {
          sheet3Data.push([parts[0].trim(), parts[1].trim()]);
        } else {
          sheet3Data.push([l.trim(), '']);
        }
      }
    });

    const ws3 = XLSX.utils.aoa_to_sheet(sheet3Data);
    ws3['!cols'] = [
      { wch: 28 },
      { wch: 18 }
    ];

    XLSX.utils.book_append_sheet(wb, ws3, 'Closing Report');

    // Write file
    const store = getVal('store') || 'STORE';
    const date = getVal('date') || 'REPORT';
    const fileName = `Store_Report_${store}_${date}.xlsx`;
    XLSX.writeFile(wb, fileName);
    showToast(`Downloaded Excel with 3 Sheets (${fileName})!`);
  });


  // Reset to default
  btnReset.addEventListener('click', () => {
    if (confirm('Reset all values to sample dataset?')) {
      localStorage.removeItem('retail_kpi_store_data');
      populateForm(sampleData);
      showToast('Form reset to default sample');
    }
  });

  btnLoadSample.addEventListener('click', () => {
    // Populate form with clean mock demo data
    populateForm(sampleData);
    currentFileName = 'Sample_POS_Demo.xls';
    uploadedFileName.textContent = currentFileName;
    fileUploadStatus.classList.remove('hidden');

    currentRawPosRows = [
      { TXNID: '1', STORENO: 'STR01', TRANSACTIONTYPE: 'SALE', RECEIPTNO: 'R1001', TRANSDATE: '24-09-2026 11:29:58', BARCODE: '616985378166', ARTICLEDESC: 'TOY CAR SPEEDSTER', QTY: '1', NET: '899', CASHIER: '50170981', SALESMANCODE: '50000001' },
      { TXNID: '2', STORENO: 'STR01', TRANSACTIONTYPE: 'SALE', RECEIPTNO: 'R1002', TRANSDATE: '24-09-2026 11:33:09', BARCODE: '8904432224961', ARTICLEDESC: 'KIDS BACKPACK BLUE 16IN', QTY: '1', NET: '2999', CASHIER: '50170981', SALESMANCODE: '50188930' },
      { TXNID: '3', STORENO: 'STR01', TRANSACTIONTYPE: 'SALE', RECEIPTNO: 'R1003', TRANSDATE: '24-09-2026 12:00:00', BARCODE: '5000394121171', ARTICLEDESC: 'AA 2 CB BATTERY', QTY: '4', NET: '396', CASHIER: '50170981', SALESMANCODE: '50195907' },
      { TXNID: '4', STORENO: 'STR01', TRANSACTIONTYPE: 'SALE', RECEIPTNO: 'R1004', TRANSDATE: '24-09-2026 12:11:38', BARCODE: '8904027127257', ARTICLEDESC: 'BUBBLE BLASTER TOY', QTY: '5', NET: '396', CASHIER: '50170981', SALESMANCODE: '50000002' },
      { TXNID: '5', STORENO: 'STR01', TRANSACTIONTYPE: 'SALE', RECEIPTNO: 'R1005', TRANSDATE: '24-09-2026 12:33:50', BARCODE: '8905450757387', ARTICLEDESC: 'BOARD GAME SET', QTY: '1', NET: '3599', CASHIER: '50170981', SALESMANCODE: '50170981' },
      { TXNID: '6', STORENO: 'STR01', TRANSACTIONTYPE: 'SALE', RECEIPTNO: 'R1006', TRANSDATE: '24-09-2026 13:01:19', BARCODE: '600552240', ARTICLEDESC: 'LOYALTY MEMBERSHIP PASS', QTY: '2', NET: '1000', CASHIER: '50170981', SALESMANCODE: '50000003' },
      { TXNID: '7', STORENO: 'STR01', TRANSACTIONTYPE: 'SALE', RECEIPTNO: 'R1007', TRANSDATE: '24-09-2026 13:30:45', BARCODE: '494500286', ARTICLEDESC: 'CANVAS TOTE BAG', QTY: '2', NET: '598', CASHIER: '50170981', SALESMANCODE: '50201915' },
      { TXNID: '8', STORENO: 'STR01', TRANSACTIONTYPE: 'SALE', RECEIPTNO: 'R1008', TRANSDATE: '24-09-2026 14:00:00', BARCODE: '310028543', ARTICLEDESC: 'BAG CARRY PAPER HAM', QTY: '10', NET: '0', CASHIER: '50170981', SALESMANCODE: '' }
    ];

    setupBarcodeFilter(currentRawPosRows, currentFileName);
    setupSalesmanFilter(currentRawPosRows);
    showToast('Loaded Demo Sample Data!');
  });


  // Theme Toggle
  btnToggleTheme.addEventListener('click', () => {
    document.body.classList.toggle('theme-light');
    const isLight = document.body.classList.contains('theme-light');
    localStorage.setItem('retail_kpi_theme', isLight ? 'light' : 'dark');
  });

  if (localStorage.getItem('retail_kpi_theme') === 'light') {
    document.body.classList.add('theme-light');
  }

  // Tabs Switching
  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      tabButtons.forEach(b => b.classList.remove('active'));
      tabContents.forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      const targetTab = document.getElementById(btn.dataset.tab);
      if (targetTab) targetTab.classList.add('active');
    });
  });

  // Format Radio Buttons
  formatRadios.forEach(radio => {
    radio.addEventListener('change', generateReport);
  });

  // Reactive Event Listeners for all inputs
  fieldIds.forEach(id => {
    const el = document.getElementById('f_' + id);
    if (el) {
      el.addEventListener('input', () => {
        recalculateAll();
      });
      el.addEventListener('change', () => {
        recalculateAll();
      });
    }
  });

  // Initialize
  loadInitialData();
});
