/**
 * Payroll Merge - Vanilla JS App
 * Compares new Excel (biweekly) with base Excel and generates a merged file.
 */

// Global state
let newData = null;     // Array of objects from new Excel
let baseData = null;    // Array of objects from base Excel
let mergedData = null;  // Merge result
let additions = [];     // New employees (altas)
let removals = [];      // Removed employees (bajas)
let efectivosData = [];  // Records without account (cash payments)

// Expected columns (kept in Spanish to match Excel files)
const COL_NEW = ['TIPOPAGO', 'NUE', 'NUP', 'RFC', 'CURP', 'NOMBRE', 'CATEGORIA', 'PUESTO', 'PROYECTO', 'NOMINA', 'DESDE', 'HASTA', 'LIQUIDO'];
const COL_BASE = ['NUM', 'NE', 'NOMBRE', 'RFC', 'CUENTA', 'BANCO', 'TELEFONO', 'CORREO ELECTRONICO', 'SE ENVIA SOBRE A', 'TIPOPAGO', 'OBSERVACIONES'];
const COL_REMOVALS = ['NUM', 'NOMBRE', 'RFC', 'CUENTA', 'BANCO', 'TELEFONO', 'CORREO ELECTRONICO', 'SE ENVIA SOBRE A', 'TIPOPAGO', 'MOTIVO'];
const COL_MERGED = ['NUM', 'NOMBRE', 'RFC', 'CURP', 'CUENTA', 'BANCO', 'TELEFONO', 'CORREO ELECTRONICO', 'SE ENVIA SOBRE A', 'TIPOPAGO', 'CATEGORIA', 'PUESTO', 'PROYECTO', 'NOMINA', 'DESDE', 'HASTA', 'LIQUIDO'];
const COL_EFECTIVOS = ['NOMBRE', 'PROYECTO', 'MODALIDAD', 'SE ENVIA SOBRE A', 'LIQUIDO', 'TELEFONO', 'OBSERVACIONES'];

// Required columns to detect header row
const REQUIRED_BASE_COLS = ['NOMBRE', 'RFC'];
const REQUIRED_NEW_COLS = ['RFC', 'NOMBRE'];
//const MAX_HEADER_SEARCH_ROWS = 20; // Search headers in first 20 rows

// Project code for Jardin (can be changed if needed)
const JARDIN_PROJECT = '1170141530100000200';

// Flag to control split by project (Jardin vs Otros)
const SPLIT_BY_PROJECT = false;

// DOM Elements - Main Menu (optional, may not exist in nominas.html)
const mainMenu = document.getElementById('mainMenu');
const nominasSection = document.getElementById('nominasSection');
const pensionesSection = document.getElementById('pensionesSection');
const menuOptions = document.querySelectorAll('.menu-option');
const btnBackNominas = document.getElementById('btnBackNominas');
const btnBackPensiones = document.getElementById('btnBackPensiones');

// DOM Elements - Nominas
const fileNewInput = document.getElementById('fileNuevo');
const fileBaseInput = document.getElementById('fileBase');
const fileNewName = document.getElementById('fileNuevoName');
const fileBaseName = document.getElementById('fileBaseName');
const btnMerge = document.getElementById('btnMerge');
const btnDownload = document.getElementById('btnDownload');
const resultsSection = document.getElementById('results');

// Tabs
const tabs = document.querySelectorAll('.tab');
const tabAdditions = document.getElementById('tabAltas');
const tabRemovals = document.getElementById('tabBajas');
const tabFinal = document.getElementById('tabFinal');

// Stats
const countAdditions = document.getElementById('countAltas');
const countRemovals = document.getElementById('countBajas');
const countTotal = document.getElementById('countTotal');

// Tables
const tableAdditions = document.getElementById('tableAltas');
const tableRemovals = document.getElementById('tableBajas');
const tableFinal = document.getElementById('tableFinal');

// ===============================
// Navigation Functions
// ===============================

/**
 * Shows a specific section and hides others
 */
function showSection(sectionName) {
  // Hide all sections
  mainMenu.classList.add('hidden');
  nominasSection.classList.add('hidden');
  pensionesSection.classList.add('hidden');

  // Show selected section
  if (sectionName === 'nominas') {
    nominasSection.classList.remove('hidden');
  } else if (sectionName === 'pensiones') {
    pensionesSection.classList.remove('hidden');
  } else {
    mainMenu.classList.remove('hidden');
  }
}

/**
 * Returns to main menu
 */
function showMainMenu() {
  mainMenu.classList.remove('hidden');
  nominasSection.classList.add('hidden');
  pensionesSection.classList.add('hidden');
}

// ===============================
// Event Listeners - Navigation
// ===============================

// Menu option clicks (only if menu exists)
if (menuOptions.length > 0) {
  menuOptions.forEach(option => {
    option.addEventListener('click', (e) => {
      const section = e.currentTarget.dataset.section;
      if (section) {
        showSection(section);
      }
    });
  });
}

// Back buttons
if (btnBackNominas) {
  btnBackNominas.addEventListener('click', () => {
    window.location.href = 'index.html';
  });
}
if (btnBackPensiones) {
  btnBackPensiones.addEventListener('click', () => {
    window.location.href = 'index.html';
  });
}

// ===============================
// Event Listeners - Nominas
// ===============================

fileNewInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) {
    fileNewName.textContent = file.name;
    readExcel(file, 'new');
  }
});

fileBaseInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) {
    fileBaseName.textContent = file.name;
    readExcel(file, 'base');
  }
});



btnMerge.addEventListener('click', () => {
  if (newData && baseData) {
    performMerge();
  }
});

btnDownload.addEventListener('click', downloadMerged);

tabs.forEach(tab => {
  tab.addEventListener('click', () => {
    tabs.forEach(t => t.classList.remove('active'));
    tab.classList.add('active');
    const target = tab.dataset.tab;
    tabAdditions.classList.toggle('hidden', target !== 'altas');
    tabRemovals.classList.toggle('hidden', target !== 'bajas');
    tabFinal.classList.toggle('hidden', target !== 'final');
  });
});

// ===============================
// Functions
// ===============================

// findHeaderRow is now in utils.js

/**
 * Reads an Excel file and converts it to an array of objects
 */
function readExcel(file, type) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

      // Detect header row automatically
      let requiredCols;
      if (type === 'base') {
        requiredCols = REQUIRED_BASE_COLS;
      } else {
        requiredCols = REQUIRED_NEW_COLS;
      }
      const headerRow = findHeaderRow(firstSheet, requiredCols);

      if (headerRow === -1) {
        const colNames = requiredCols.join(', ');
        alert(`No se encontraron los encabezados (${colNames}) en las primeras ${MAX_HEADER_SEARCH_ROWS} filas del archivo.`);
        return;
      }

      const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
        defval: '',
        range: headerRow
      });

      // Normalize data using utils function
      const normalizedData = normalizeData(jsonData);
      let fileTypeName;
      if (type === 'new') {
        fileTypeName = 'Nuevo';
      } else {
        fileTypeName = 'Base';
      }

      // Validate required columns (warning only, don't block processing)
      const missingCols = validateRequiredColumns(normalizedData, type);
      if (missingCols) {
        const missingColsStr = missingCols.join(', ');
        alert(`⚠️ ADVERTENCIA: El archivo ${fileTypeName} no contiene todos los campos requeridos.\n\nCampos faltantes:\n${missingColsStr}\n\nPuedes continuar, pero algunos procesos podrían no funcionar correctamente.`);
      }

      // Assign to global variables
      if (type === 'new') {
        newData = normalizedData;
        console.log('New Excel loaded:', newData.length, 'rows');
      } else {
        baseData = normalizedData;
        console.log('Base Excel loaded:', baseData.length, 'rows');
      }

      updateMergeButton();
    } catch (err) {
      alert('Error reading file: ' + err.message);
      console.error(err);
    }
  };
  reader.readAsArrayBuffer(file);
}

// normalizeData is now in utils.js

/**
 * Validates that the data contains all required columns
 * @param {Array} data - Normalized data array
 * @param {string} type - Type of file: 'new', 'base', or 'cash'
 * @returns {Array|null} - Array of missing columns or null if all are present
 */
function validateRequiredColumns(data, type) {
  if (!data || data.length === 0) {
    return null; // Empty data, validation will happen elsewhere
  }

  // Get expected columns based on type
  let expectedCols;
  let fileTypeName;
  if (type === 'new') {
    expectedCols = COL_NEW;
    fileTypeName = 'Nuevo';
  } else if (type === 'base') {
    expectedCols = COL_BASE;
    fileTypeName = 'Base';
  } else {
    return null;
  }

  // Get all available keys from the first row (all rows should have same structure after normalization)
  const availableKeys = new Set(Object.keys(data[0]));

  // Find missing columns
  const missingCols = expectedCols.filter(col => !availableKeys.has(col));

  return missingCols.length > 0 ? missingCols : null;
}

/**
 * Enables the merge button if both files are loaded
 */
function updateMergeButton() {
  btnMerge.disabled = !(newData && baseData);
}

/**
 * Performs comparison and merge
 * - Additions: people in new Excel not in base (by RFC)
 * - Removals: people in base not in new Excel (by RFC)
 * - Merge: base + additions, removing removals, updating data from new
 * - Note: A person can have multiple rows with different TIPOPAGO (NORMAL, RETROACTIVO)
 */
function performMerge() {
  // Create map by RFC for base (one entry per person)
  const rfcBase = new Map();
  baseData.forEach(row => {
    const rfc = (row.RFC || '').toUpperCase();
    if (rfc) rfcBase.set(rfc, row);
  });

  // For new data, keep ALL rows (a person can have NORMAL and RETROACTIVO)
  // Group by RFC to know which RFCs exist in new
  const rfcNewSet = new Set();
  newData.forEach(row => {
    const rfc = (row.RFC || '').toUpperCase();
    if (rfc) rfcNewSet.add(rfc);
  });

  // Detect additions: people in new Excel not in base (by RFC)
  additions = [];
  const addedRfcs = new Set();
  newData.forEach(rowNew => {
    const rfc = (rowNew.RFC || '').toUpperCase();
    if (!rfc || addedRfcs.has(rfc)) return;

    const rowBase = rfcBase.get(rfc);
    const isInBase = !!rowBase;

    if (!isInBase) {
      additions.push(rowNew);
      addedRfcs.add(rfc);
      console.log('Addition (new employee):', rowNew.NOMBRE || rfc);
    }
  });

  // Detect removals: people in base but not in new Excel
  removals = [];
  rfcBase.forEach((rowBase, rfc) => {
    if (!rfcNewSet.has(rfc)) {
      removals.push({
        ...rowBase,
        MOTIVO: 'No aparece en nómina nueva'
      });
    }
  });

  // Create merged: iterate through ALL rows in newData
  // Each row in newData becomes a row in mergedData (with base info if available)
  mergedData = [];
  efectivosData = [];
  let num = 1;

  // Process ALL rows from new Excel (includes NORMAL and RETROACTIVO for same person)
  newData.forEach(rowNew => {
    const rfc = (rowNew.RFC || '').toUpperCase();
    const rowBase = rfcBase.get(rfc);

    const cuenta = rowBase && rowBase.CUENTA ? String(rowBase.CUENTA).trim() : '';
    const hasAccount = cuenta.length > 0;

    const mergedRow = {
      NUM: num++,
      NE: rowBase ? rowBase.NE : '',
      NOMBRE: rowNew.NOMBRE || (rowBase ? rowBase.NOMBRE : ''),
      RFC: rfc,
      CURP: rowNew.CURP || '',
      CUENTA: cuenta,
      BANCO: rowBase ? rowBase.BANCO : '',
      TELEFONO: rowBase ? rowBase.TELEFONO : '',
      'CORREO ELECTRONICO': rowBase ? rowBase['CORREO ELECTRONICO'] : '',
      'SE ENVIA SOBRE A': rowBase ? rowBase['SE ENVIA SOBRE A'] : '',
      OBSERVACIONES: rowBase ? rowBase.OBSERVACIONES : '',
      TIPOPAGO: rowNew.TIPOPAGO || '',
      MODALIDAD: rowNew.TIPOPAGO || '', // MODALIDAD is same as TIPOPAGO for nominas
      CATEGORIA: rowNew.CATEGORIA || '',
      PUESTO: rowNew.PUESTO || '',
      PROYECTO: rowNew.PROYECTO || '',
      NOMINA: rowNew.NOMINA || '',
      DESDE: rowNew.DESDE || '',
      HASTA: rowNew.HASTA || '',
      LIQUIDO: rowNew.LIQUIDO || ''
    };

    mergedData.push(mergedRow);

    // If no account, add to efectivos
    if (!hasAccount) {
      efectivosData.push(mergedRow);
    }
  });

  // Display results
  displayResults();
}

/**
 * Displays results in the UI
 */
function displayResults() {
  resultsSection.classList.remove('hidden');

  // Stats
  countAdditions.textContent = additions.length;
  countRemovals.textContent = removals.length;
  countTotal.textContent = mergedData.length;

  // Additions table
  renderTable(tableAdditions, additions, COL_NEW);
  tabAdditions.querySelector('.empty-msg').classList.toggle('hidden', additions.length > 0);

  // Removals table
  renderTable(tableRemovals, removals, COL_REMOVALS);
  tabRemovals.querySelector('.empty-msg').classList.toggle('hidden', removals.length > 0);

  // Final table
  renderTable(tableFinal, mergedData, COL_MERGED);

  // Scroll to results
  resultsSection.scrollIntoView({ behavior: 'smooth' });
}

// renderTable is now in utils.js

/**
 * Downloads the merged Excel file
 */
function downloadMerged() {
  if (!mergedData || mergedData.length === 0) {
    alert('No hay datos para descargar');
    return;
  }

  // Create workbook
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(mergedData, { header: COL_MERGED });

  // Adjust column widths
  const colWidths = COL_MERGED.map(col => ({ wch: Math.max(col.length, 15) }));
  ws['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(wb, ws, 'Nomina Fusionada');

  // Generate filename with date
  const today = new Date();
  const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');
  const fileName = `Nomina_Fusionada_${dateStr}.xls`;

  // Download
  XLSX.writeFile(wb, fileName);
}

// ===============================
// Split Functionality
// ===============================

const btnSplit = document.getElementById('btnSplit');
const btnDownloadAll = document.getElementById('btnDownloadAll');
const btnDownloadEfectivosNominas = document.getElementById('btnDownloadEfectivosNominas');
const splitResultsSection = document.getElementById('splitResults');
const splitTree = document.getElementById('splitTree');

let splitData = null; // Stores the hierarchical split structure

btnSplit.addEventListener('click', performSplit);
btnDownloadAll.addEventListener('click', downloadAllSplitFiles);
if (btnDownloadEfectivosNominas) {
  btnDownloadEfectivosNominas.addEventListener('click', downloadEfectivosNominas);
}

/**
 * Determines if a project code belongs to "Jardin"
 */
function isJardin(proyecto) {
  const proyectoStr = String(proyecto || '').trim();
  return proyectoStr === JARDIN_PROJECT;
}

/**
 * Calculates the payroll period based on current date
 * Returns { quincena: '1a' or '2a', mes: 'Ene', 'Feb', etc., display: '1a Nomina de Ene' }
 */
function getPayrollPeriod() {
  const today = new Date();
  const day = today.getDate();
  const month = today.getMonth(); // 0-11

  const monthNames = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const mes = monthNames[month];

  // First 15 days = 1a quincena, rest = 2a quincena
  const quincena = day <= 15 ? '1a' : '2a';

  return {
    quincena,
    mes,
    display: `${quincena} Nomina de ${mes}`
  };
}

/**
 * Splits the merged data by Proyecto -> Nomina -> TipoPago -> Banco
 * or by Nomina -> TipoPago -> Banco (if SPLIT_BY_PROJECT is false)
 */
function performSplit() {
  if (!mergedData || mergedData.length === 0) {
    alert('No hay datos para separar. Primero realiza la fusión.');
    return;
  }

  // Initialize split structure
  if (SPLIT_BY_PROJECT) {
    splitData = {
      'JARDIN': {},
      'OTROS': {}
    };
  } else {
    splitData = {};
  }

  // Group data hierarchically
  mergedData.forEach(row => {
    // Level 4: Banco - Skip rows without bank info
    const banco = (row.BANCO || '').toUpperCase().trim();
    if (!banco) {
      console.log('Split skipped (no bank):', row.NOMBRE || row.RFC);
      return;
    }

    // Level 1: Proyecto (Jardin vs Otros) - only if flag is enabled
    let projectGroup = null;
    if (SPLIT_BY_PROJECT) {
      projectGroup = isJardin(row.PROYECTO) ? 'JARDIN' : 'OTROS';
    }

    // Level 2: Nomina
    const nomina = (row.NOMINA || 'SIN_NOMINA').toUpperCase().trim();

    // Level 3: TipoPago
    const tipoPago = (row.TIPOPAGO || 'SIN_TIPOPAGO').toUpperCase().trim();

    // Initialize nested structure if needed
    if (SPLIT_BY_PROJECT) {
      if (!splitData[projectGroup][nomina]) {
        splitData[projectGroup][nomina] = {};
      }
      if (!splitData[projectGroup][nomina][tipoPago]) {
        splitData[projectGroup][nomina][tipoPago] = {};
      }
      if (!splitData[projectGroup][nomina][tipoPago][banco]) {
        splitData[projectGroup][nomina][tipoPago][banco] = [];
      }
      splitData[projectGroup][nomina][tipoPago][banco].push(row);
    } else {
      if (!splitData[nomina]) {
        splitData[nomina] = {};
      }
      if (!splitData[nomina][tipoPago]) {
        splitData[nomina][tipoPago] = {};
      }
      if (!splitData[nomina][tipoPago][banco]) {
        splitData[nomina][tipoPago][banco] = [];
      }
      splitData[nomina][tipoPago][banco].push(row);
    }
  });

  // Display results
  displaySplitResults();
}

// calculateTotalAmount and formatCurrency are now in utils.js

/**
 * Displays the split results in a tree structure
 */
function displaySplitResults() {
  splitResultsSection.classList.remove('hidden');
  splitTree.innerHTML = '';

  // Display calculated payroll period
  const payrollPeriod = getPayrollPeriod();
  document.getElementById('payrollPeriodDisplay').textContent = payrollPeriod.display;

  // Update split description based on flag
  const splitDescription = document.getElementById('splitDescription');
  if (SPLIT_BY_PROJECT) {
    splitDescription.textContent = 'Archivos separados por: Proyecto (Jardín/Otros) → Nómina → TipoPago → Banco';
  } else {
    splitDescription.textContent = 'Archivos separados por: Nómina → TipoPago → Banco';
  }

  const projectIcons = { 'JARDIN': '🌳', 'OTROS': '🏢' };

  if (SPLIT_BY_PROJECT) {
    // Structure: Proyecto -> Nomina -> TipoPago -> Banco
    for (const [projectGroup, nominas] of Object.entries(splitData)) {
      // Skip empty project groups
      if (Object.keys(nominas).length === 0) continue;

      const projectDiv = document.createElement('div');
      projectDiv.className = 'split-project';

      const projectName = document.createElement('div');
      projectName.className = 'split-project-name';
      projectName.innerHTML = `${projectIcons[projectGroup] || '📁'} ${projectGroup}`;
      projectDiv.appendChild(projectName);

      for (const [nomina, tipoPagos] of Object.entries(nominas)) {
        const nominaDiv = document.createElement('div');
        nominaDiv.className = 'split-nomina';

        const nominaName = document.createElement('div');
        nominaName.className = 'split-nomina-name';
        nominaName.textContent = `📋 ${nomina}`;
        nominaDiv.appendChild(nominaName);

        for (const [tipoPago, bancos] of Object.entries(tipoPagos)) {
          const tipoPagoDiv = document.createElement('div');
          tipoPagoDiv.className = 'split-tipopago-group';

          const tipoPagoName = document.createElement('div');
          tipoPagoName.className = 'split-tipopago-name';
          tipoPagoName.textContent = `💰 ${tipoPago}`;
          tipoPagoDiv.appendChild(tipoPagoName);

          for (const [banco, rows] of Object.entries(bancos)) {
            const bancoDiv = document.createElement('div');
            bancoDiv.className = 'split-banco';

            const totalAmount = calculateTotalAmount(rows, 'LIQUIDO');
            const formattedAmount = formatCurrency(totalAmount);

            bancoDiv.innerHTML = `
              <span>🏦 ${banco}</span>
              <span class="count">${rows.length} registros</span>
              <span class="amount">${formattedAmount}</span>
              <button class="btn-download-single" data-project="${projectGroup}" data-nomina="${nomina}" data-tipopago="${tipoPago}" data-banco="${banco}">⬇️</button>
            `;

            tipoPagoDiv.appendChild(bancoDiv);
          }

          nominaDiv.appendChild(tipoPagoDiv);
        }

        projectDiv.appendChild(nominaDiv);
      }

      splitTree.appendChild(projectDiv);
    }
  } else {
    // Structure: Nomina -> TipoPago -> Banco (no project level)
    for (const [nomina, tipoPagos] of Object.entries(splitData)) {
      // Skip empty nominas
      if (Object.keys(tipoPagos).length === 0) continue;

      const nominaDiv = document.createElement('div');
      nominaDiv.className = 'split-nomina';

      const nominaName = document.createElement('div');
      nominaName.className = 'split-nomina-name';
      nominaName.textContent = `📋 ${nomina}`;
      nominaDiv.appendChild(nominaName);

      for (const [tipoPago, bancos] of Object.entries(tipoPagos)) {
        const tipoPagoDiv = document.createElement('div');
        tipoPagoDiv.className = 'split-tipopago-group';

        const tipoPagoName = document.createElement('div');
        tipoPagoName.className = 'split-tipopago-name';
        tipoPagoName.textContent = `💰 ${tipoPago}`;
        tipoPagoDiv.appendChild(tipoPagoName);

        for (const [banco, rows] of Object.entries(bancos)) {
          const bancoDiv = document.createElement('div');
          bancoDiv.className = 'split-banco';

          const totalAmount = calculateTotalAmount(rows, 'LIQUIDO');
          const formattedAmount = formatCurrency(totalAmount);

          bancoDiv.innerHTML = `
            <span>🏦 ${banco}</span>
            <span class="count">${rows.length} registros</span>
            <span class="amount">${formattedAmount}</span>
            <button class="btn-download-single" data-nomina="${nomina}" data-tipopago="${tipoPago}" data-banco="${banco}">⬇️</button>
          `;

          tipoPagoDiv.appendChild(bancoDiv);
        }

        nominaDiv.appendChild(tipoPagoDiv);
      }

      splitTree.appendChild(nominaDiv);
    }
  }

  // Add click handlers for individual download buttons
  splitTree.querySelectorAll('.btn-download-single').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const project = e.target.dataset.project || null;
      const nomina = e.target.dataset.nomina;
      const tipoPago = e.target.dataset.tipopago;
      const banco = e.target.dataset.banco;
      downloadSingleSplitFile(project, nomina, tipoPago, banco);
    });
  });

  splitResultsSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * Determines account type for Banamex based on account number length
 * @param {string} cuenta - Account number
 * @returns {string} - 'Tarjeta' if 16 digits, 'Cheque' if 9 or 12 digits
 */
function getBanamexAccountType(cuenta) {
  const cuentaStr = String(cuenta || '').trim();
  // Remove all non-digit characters to count only digits
  const digitsOnly = cuentaStr.replace(/\D/g, '');
  const digitCount = digitsOnly.length;

  if (digitCount === 16) {
    return 'Tarjeta';
  } else {
    return 'Cheque';
  }
}

/**
 * Downloads a single split file
 * For BANAMEX: special format with specific columns
 */
function downloadSingleSplitFile(project, nomina, tipoPago, banco) {
  let rows;
  if (SPLIT_BY_PROJECT) {
    rows = splitData[project]?.[nomina]?.[tipoPago]?.[banco];
  } else {
    rows = splitData[nomina]?.[tipoPago]?.[banco];
  }

  if (!rows || rows.length === 0) {
    alert('No hay datos para este archivo');
    return;
  }

  const wb = XLSX.utils.book_new();
  let ws;
  let fileName;

  const today = new Date();
  const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');

  // Check if BANAMEX - use special format
  if (banco.toUpperCase() === 'BANAMEX') {
    // Get payroll period (calculated automatically)
    const payrollPeriod = getPayrollPeriod();
    const conceptoBancario = payrollPeriod.display;

    // Transform data to Banamex format
    const banamexData = rows.map((row, index) => ({
      'Tipo de Cuenta': getBanamexAccountType(row.CUENTA),
      'Cuenta': row.CUENTA || '',
      'Importe': row.LIQUIDO || 0,
      'Nombre/Razón Social': row.NOMBRE || '',
      'Ref. Num.': index + 1,
      'Ref. AlfN.': conceptoBancario
    }));

    const banamexHeaders = ['Tipo de Cuenta', 'Cuenta', 'Importe', 'Nombre/Razón Social', 'Ref. Num.', 'Ref. AlfN.'];
    ws = XLSX.utils.json_to_sheet(banamexData, { header: banamexHeaders });

    const colWidths = banamexHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
    ws['!cols'] = colWidths;

    fileName = SPLIT_BY_PROJECT
      ? `BANAMEX_${project}_${nomina}_${tipoPago}_${dateStr}.xls`
      : `BANAMEX_${nomina}_${tipoPago}_${dateStr}.xls`;
  } else if (banco.toUpperCase() === 'BANORTE') {
    // Transform data to Banorte format
    const banorteData = rows.map((row) => ({
      'NO. EMPLEADO': row.NE || '',
      'NOMBRE': row.NOMBRE || '',
      'IMPORTE': row.LIQUIDO || 0,
      'NO. BANCO RECEPTOR': '072',
      'TIPO DE CUENTA': '01',
      'CUENTA': row.CUENTA || ''
    }));

    const banorteHeaders = ['NO. EMPLEADO', 'NOMBRE', 'IMPORTE', 'NO. BANCO RECEPTOR', 'TIPO DE CUENTA', 'CUENTA'];
    ws = XLSX.utils.json_to_sheet(banorteData, { header: banorteHeaders });

    const colWidths = banorteHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
    ws['!cols'] = colWidths;

    fileName = SPLIT_BY_PROJECT
      ? `BANORTE_${project}_${nomina}_${tipoPago}_${dateStr}.xls`
      : `BANORTE_${nomina}_${tipoPago}_${dateStr}.xls`;
  } else {
    // Standard format for other banks
    ws = XLSX.utils.json_to_sheet(rows, { header: COL_MERGED });

    const colWidths = COL_MERGED.map(col => ({ wch: Math.max(col.length, 15) }));
    ws['!cols'] = colWidths;

    fileName = SPLIT_BY_PROJECT
      ? `${project}_${nomina}_${tipoPago}_${banco}_${dateStr}.xls`
      : `${nomina}_${tipoPago}_${banco}_${dateStr}.xls`;
  }

  XLSX.utils.book_append_sheet(wb, ws, 'Datos');
  XLSX.writeFile(wb, fileName);
}

/**
 * Downloads all split files at once
 * Applies special Banamex format when banco is BANAMEX
 */
function downloadAllSplitFiles() {
  if (!splitData) {
    alert('No hay datos para descargar');
    return;
  }

  const today = new Date();
  const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');

  // Get payroll period (calculated automatically)
  const payrollPeriod = getPayrollPeriod();
  const conceptoBancario = payrollPeriod.display;

  if (SPLIT_BY_PROJECT) {
    // Structure: Proyecto -> Nomina -> TipoPago -> Banco
    for (const [project, nominas] of Object.entries(splitData)) {
      for (const [nomina, tipoPagos] of Object.entries(nominas)) {
        for (const [tipoPago, bancos] of Object.entries(tipoPagos)) {
          for (const [banco, rows] of Object.entries(bancos)) {
            if (rows.length === 0) continue;

            const wb = XLSX.utils.book_new();
            let ws;
            let fileName;

            // Check if BANAMEX - use special format
            if (banco.toUpperCase() === 'BANAMEX') {
              // Transform data to Banamex format
              const banamexData = rows.map((row, index) => ({
                'Tipo de Cuenta': getBanamexAccountType(row.CUENTA),
                'Cuenta': row.CUENTA || '',
                'Importe': row.LIQUIDO || 0,
                'Nombre/Razón Social': row.NOMBRE || '',
                'Ref. Num.': index + 1,
                'Ref. AlfN.': conceptoBancario
              }));

              const banamexHeaders = ['Tipo de Cuenta', 'Cuenta', 'Importe', 'Nombre/Razón Social', 'Ref. Num.', 'Ref. AlfN.'];
              ws = XLSX.utils.json_to_sheet(banamexData, { header: banamexHeaders });

              const colWidths = banamexHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
              ws['!cols'] = colWidths;

              fileName = `BANAMEX_${project}_${nomina}_${tipoPago}_${dateStr}.xls`;
            } else if (banco.toUpperCase() === 'BANORTE') {
              // Transform data to Banorte format
              const banorteData = rows.map((row) => ({
                'NO. EMPLEADO': row.NE || '',
                'NOMBRE': row.NOMBRE || '',
                'IMPORTE': row.LIQUIDO || 0,
                'NO. BANCO RECEPTOR': '072',
                'TIPO DE CUENTA': '01',
                'CUENTA': row.CUENTA || ''
              }));

              const banorteHeaders = ['NO. EMPLEADO', 'NOMBRE', 'IMPORTE', 'NO. BANCO RECEPTOR', 'TIPO DE CUENTA', 'CUENTA'];
              ws = XLSX.utils.json_to_sheet(banorteData, { header: banorteHeaders });

              const colWidths = banorteHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
              ws['!cols'] = colWidths;

              fileName = `BANORTE_${project}_${nomina}_${tipoPago}_${dateStr}.xls`;
            } else {
              // Standard format for other banks
              ws = XLSX.utils.json_to_sheet(rows, { header: COL_MERGED });

              const colWidths = COL_MERGED.map(col => ({ wch: Math.max(col.length, 15) }));
              ws['!cols'] = colWidths;

              fileName = `${project}_${nomina}_${tipoPago}_${banco}_${dateStr}.xls`;
            }

            XLSX.utils.book_append_sheet(wb, ws, 'Datos');
            XLSX.writeFile(wb, fileName);
          }
        }
      }
    }
  } else {
    // Structure: Nomina -> TipoPago -> Banco (no project level)
    for (const [nomina, tipoPagos] of Object.entries(splitData)) {
      for (const [tipoPago, bancos] of Object.entries(tipoPagos)) {
        for (const [banco, rows] of Object.entries(bancos)) {
          if (rows.length === 0) continue;

          const wb = XLSX.utils.book_new();
          let ws;
          let fileName;

          // Check if BANAMEX - use special format
          if (banco.toUpperCase() === 'BANAMEX') {
            // Transform data to Banamex format
            const banamexData = rows.map((row, index) => ({
              'Tipo de Cuenta': getBanamexAccountType(row.CUENTA),
              'Cuenta': row.CUENTA || '',
              'Importe': row.LIQUIDO || 0,
              'Nombre/Razón Social': row.NOMBRE || '',
              'Ref. Num.': index + 1,
              'Ref. AlfN.': conceptoBancario
            }));

            const banamexHeaders = ['Tipo de Cuenta', 'Cuenta', 'Importe', 'Nombre/Razón Social', 'Ref. Num.', 'Ref. AlfN.'];
            ws = XLSX.utils.json_to_sheet(banamexData, { header: banamexHeaders });

            const colWidths = banamexHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
            ws['!cols'] = colWidths;

            fileName = `BANAMEX_${nomina}_${tipoPago}_${dateStr}.xls`;
          } else if (banco.toUpperCase() === 'BANORTE') {
            // Transform data to Banorte format
            const banorteData = rows.map((row) => ({
              'NO. EMPLEADO': row.NE || '',
              'NOMBRE': row.NOMBRE || '',
              'IMPORTE': row.LIQUIDO || 0,
              'NO. BANCO RECEPTOR': '072',
              'TIPO DE CUENTA': '01',
              'CUENTA': row.CUENTA || ''
            }));

            const banorteHeaders = ['NO. EMPLEADO', 'NOMBRE', 'IMPORTE', 'NO. BANCO RECEPTOR', 'TIPO DE CUENTA', 'CUENTA'];
            ws = XLSX.utils.json_to_sheet(banorteData, { header: banorteHeaders });

            const colWidths = banorteHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
            ws['!cols'] = colWidths;

            fileName = `BANORTE_${nomina}_${tipoPago}_${dateStr}.xls`;
          } else {
            // Standard format for other banks
            ws = XLSX.utils.json_to_sheet(rows, { header: COL_MERGED });

            const colWidths = COL_MERGED.map(col => ({ wch: Math.max(col.length, 15) }));
            ws['!cols'] = colWidths;

            fileName = `${nomina}_${tipoPago}_${banco}_${dateStr}.xls`;
          }

          XLSX.utils.book_append_sheet(wb, ws, 'Datos');
          XLSX.writeFile(wb, fileName);
        }
      }
    }
  }
}

/**
 * Downloads the efectivos (cash payments) file
 * Includes people without bank accounts with OBSERVACIONES column
 */
function downloadEfectivosNominas() {
  if (!efectivosData || efectivosData.length === 0) {
    alert('No hay registros sin cuenta para descargar');
    return;
  }

  const today = new Date();
  const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');
  const fileName = `Nominas_Efectivos_${dateStr}.xls`;

  // Transform data to only include the fields in COL_EFECTIVOS
  const efectivosFiltered = efectivosData.map(row => ({
    'NOMBRE': row.NOMBRE || '',
    'PROYECTO': row.PROYECTO || '',
    'MODALIDAD': row.MODALIDAD || '',
    'SE ENVIA SOBRE A': row['SE ENVIA SOBRE A'] || '',
    'LIQUIDO': row.LIQUIDO || 0,
    'TELEFONO': row.TELEFONO || '',
    'OBSERVACIONES': row.OBSERVACIONES || ''
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(efectivosFiltered, { header: COL_EFECTIVOS });

  // Adjust column widths
  const colWidths = COL_EFECTIVOS.map(col => ({ wch: Math.max(col.length, 15) }));
  ws['!cols'] = colWidths;

  XLSX.utils.book_append_sheet(wb, ws, 'Efectivos');
  XLSX.writeFile(wb, fileName);
}
