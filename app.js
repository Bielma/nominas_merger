/**
 * Payroll Merge - Vanilla JS App
 * Compares new Excel (biweekly) with base Excel and generates a merged file.
 */

// Global state
let newData = null;     // Array of objects from new Excel
let baseData = null;    // Array of objects from base Excel
let cashData = null;    // Array of objects from cash payments Excel (optional)
let mergedData = null;  // Merge result
let additions = [];     // New employees (altas)
let removals = [];      // Removed employees (bajas)

// Expected columns (kept in Spanish to match Excel files)
const COL_NEW = ['TIPOPAGO', 'NUE', 'NUP', 'RFC', 'CURP', 'NOMBRE', 'CATEGORIA', 'PUESTO', 'PROYECTO', 'NOMINA', 'DESDE', 'HASTA', 'LIQUIDO'];
const COL_BASE = ['NUM', 'NE', 'NOMBRE', 'RFC', 'CUENTA', 'BANCO', 'TELEFONO', 'CORREO ELECTRONICO', 'SE ENVIA SOBRE A', 'TIPOPAGO', 'OBSERVACIONES'];
const COL_CASH = ['RFC', 'NOMBRE', 'MODALIDAD', 'MONTO', 'MOTIVO'];
const COL_REMOVALS = ['NUM', 'NOMBRE', 'RFC', 'CUENTA', 'BANCO', 'TELEFONO', 'CORREO ELECTRONICO', 'SE ENVIA SOBRE A', 'TIPOPAGO', 'MOTIVO'];
const COL_MERGED = ['NUM', 'NOMBRE', 'RFC', 'CURP', 'CUENTA', 'BANCO', 'TELEFONO', 'CORREO ELECTRONICO', 'SE ENVIA SOBRE A', 'TIPOPAGO', 'CATEGORIA', 'PUESTO', 'PROYECTO', 'NOMINA', 'DESDE', 'HASTA', 'LIQUIDO'];

// Required columns to detect header row
const REQUIRED_BASE_COLS = ['NOMBRE', 'RFC'];
const REQUIRED_NEW_COLS = ['RFC', 'NOMBRE'];
const REQUIRED_CASH_COLS = ['RFC', 'NOMBRE'];
const MAX_HEADER_SEARCH_ROWS = 20; // Search headers in first 20 rows

// Project code for Jardin (can be changed if needed)
const JARDIN_PROJECT = '1170141530100000200';

// Flag to control split by project (Jardin vs Otros)
const SPLIT_BY_PROJECT = false;

// DOM Elements - Main Menu
const mainMenu = document.getElementById('mainMenu');
const nominasSection = document.getElementById('nominasSection');
const pensionesSection = document.getElementById('pensionesSection');
const menuOptions = document.querySelectorAll('.menu-option');
const btnBackNominas = document.getElementById('btnBackNominas');
const btnBackPensiones = document.getElementById('btnBackPensiones');

// DOM Elements - Nominas
const fileNewInput = document.getElementById('fileNuevo');
const fileBaseInput = document.getElementById('fileBase');
const fileCashInput = document.getElementById('fileCash');
const fileNewName = document.getElementById('fileNuevoName');
const fileBaseName = document.getElementById('fileBaseName');
const fileCashName = document.getElementById('fileCashName');
const btnMerge = document.getElementById('btnMerge');
const btnDownload = document.getElementById('btnDownload');
const btnClearCash = document.getElementById('btnClearCash');
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

// Menu option clicks
if (menuOptions.length > 0) {
	menuOptions.forEach(option => {
		option.addEventListener('click', (e) => {
			const section = e.currentTarget.dataset.section;
			showSection(section);
		});
	});
}

// Back buttons
if (btnBackNominas) {
	btnBackNominas.addEventListener('click', showMainMenu);
}
if (btnBackPensiones) {
	btnBackPensiones.addEventListener('click', showMainMenu);
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

fileCashInput.addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (file) {
    fileCashName.textContent = file.name;
    btnClearCash.classList.remove('hidden');
    readExcel(file, 'cash');
  }
});

btnClearCash.addEventListener('click', () => {
  cashData = null;
  fileCashInput.value = '';
  fileCashName.textContent = 'Sin archivo';
  btnClearCash.classList.add('hidden');
  console.log('Cash Excel cleared');
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

/**
 * Finds the row index where headers are located by searching for required columns
 * @param {object} sheet - XLSX sheet object
 * @param {string[]} requiredCols - Array of column names that must be present
 * @returns {number} - 0-indexed row number where headers were found, or -1 if not found
 */
function findHeaderRow(sheet, requiredCols) {
  const range = XLSX.utils.decode_range(sheet['!ref']);
  const maxRow = Math.min(range.e.r, MAX_HEADER_SEARCH_ROWS);
  
  for (let row = 0; row <= maxRow; row++) {
    const rowValues = [];
    
    // Collect all cell values in this row
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[cellAddress];
      if (cell && cell.v !== undefined) {
        const value = String(cell.v).trim().toUpperCase();
        rowValues.push(value);
      }
    }
    
    // Check if all required columns are present in this row
    const foundAll = requiredCols.every(reqCol => 
      rowValues.some(val => val.includes(reqCol.toUpperCase()))
    );
    
    if (foundAll) {
      console.log(`Headers found at row ${row + 1} (0-indexed: ${row})`);
      return row;
    }
  }
  
  return -1; // Not found
}

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
      } else if (type === 'cash') {
        requiredCols = REQUIRED_CASH_COLS;
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

      // Normalize data
      let normalizedData;
      let fileTypeName;
      if (type === 'new') {
        normalizedData = normalizeData(jsonData);
        fileTypeName = 'Nuevo';
      } else if (type === 'cash') {
        normalizedData = normalizeData(jsonData);
        fileTypeName = 'Efectivo';
      } else {
        normalizedData = normalizeData(jsonData);
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
      } else if (type === 'cash') {
        cashData = normalizedData;
        console.log('Cash Excel loaded:', cashData.length, 'rows');
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

/**
 * Normalizes object keys (trim and uppercase)
 */
function normalizeData(data) {
  return data.map(row => {
    const normalized = {};
    for (const key in row) {
      const cleanKey = key.trim().toUpperCase();
      normalized[cleanKey] = typeof row[key] === 'string' ? row[key].trim() : row[key];
    }
    return normalized;
  });
}

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
  } else if (type === 'cash') {
    expectedCols = COL_CASH;
    fileTypeName = 'Efectivo';
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

  // Build set of RFCs from cash payments (these are not considered additions)
  // Also track removals from cash file (where MOTIVO contains "BAJA")
  const cashRfcs = new Set();
  const cashRemovals = new Map(); // RFC -> row data for people marked as "BAJA"
  
  if (cashData && cashData.length > 0) {
    cashData.forEach(row => {
      const rfc = (row.RFC || '').toUpperCase().trim();
      const motivo = (row.MOTIVO || '').toUpperCase();
      
      if (rfc) {
        cashRfcs.add(rfc);
        
        // Check if MOTIVO contains "BAJA"
        if (motivo.includes('BAJA')) {
          cashRemovals.set(rfc, row);
          console.log('Removal detected from cash (BAJA):', row.NOMBRE || rfc, '-', row.MOTIVO);
        }
      }
    });
    console.log('Cash payments excluded:', cashRfcs.size, 'people');
    console.log('Cash removals (BAJA):', cashRemovals.size, 'people');
  }

  // Detect additions:
  // Case 1: People in new but not in base (and not in cash payments)
  // Case 2: People in new, in base but WITHOUT bank/account info, and not in cash payments
  additions = [];
  const addedRfcs = new Set();
  newData.forEach(rowNew => {
    const rfc = (rowNew.RFC || '').toUpperCase();
    if (!rfc || addedRfcs.has(rfc)) return;
    
    const rowBase = rfcBase.get(rfc);
    const isInBase = !!rowBase;
    const hasBankInfo = rowBase && String(rowBase.CUENTA || '').trim();
    const isInCash = cashRfcs.has(rfc);
    
    // Case 1: Not in base and not in cash
    if (!isInBase && !isInCash) {
      additions.push(rowNew);
      addedRfcs.add(rfc);
      console.log('Addition (new employee):', rowNew.NOMBRE || rfc);
    }
    // Case 2: In base but no bank/account info and not in cash
    else if (isInBase && !hasBankInfo && !isInCash) {
      additions.push(rowNew);
      addedRfcs.add(rfc);
      console.log('Addition (no bank info):', rowNew.NOMBRE || rfc);
    }
    // Excluded: in cash payments
    else if (!isInBase && isInCash) {
      console.log('Excluded from additions (cash payment):', rowNew.NOMBRE || rfc);
    }
  });

  // Detect removals: 
  // 1. People in base but not in new Excel
  // 2. People in cash file with MOTIVO containing "BAJA"
  removals = [];
  
  // Case 1: in base but not in new
  rfcBase.forEach((rowBase, rfc) => {
    if (!rfcNewSet.has(rfc)) {
      // Check if there's a MOTIVO from cash file for this person
      const cashRow = cashRemovals.get(rfc);
      removals.push({
        ...rowBase,
        MOTIVO: cashRow ? cashRow.MOTIVO : 'No aparece en nómina nueva'
      });
    }
  });
  
  // Case 2: marked as BAJA in cash file (add if not already in removals)
  const removalsRfcs = new Set(removals.map(r => (r.RFC || '').toUpperCase()));
  cashRemovals.forEach((cashRow, rfc) => {
    if (!removalsRfcs.has(rfc)) {
      // Try to get full info from base if available, otherwise use cash data
      const baseRow = rfcBase.get(rfc);
      if (baseRow) {
        removals.push({
          ...baseRow,
          MOTIVO: cashRow.MOTIVO || ''
        });
      } else {
        // Create a minimal row from cash data
        removals.push({
          NUM: '',
          NOMBRE: cashRow.NOMBRE || '',
          RFC: rfc,
          CUENTA: '',
          BANCO: '',
          TELEFONO: '',
          'CORREO ELECTRONICO': '',
          'SE ENVIA SOBRE A': '',
          MOTIVO: cashRow.MOTIVO || ''
        });
      }
    }
  });

  // Create merged: iterate through ALL rows in newData
  // Each row in newData becomes a row in mergedData (with base info if available)
  mergedData = [];
  let num = 1;

  // Process ALL rows from new Excel (includes NORMAL and RETROACTIVO for same person)
  newData.forEach(rowNew => {
    const rfc = (rowNew.RFC || '').toUpperCase();
    const rowBase = rfcBase.get(rfc);
    
    mergedData.push({
      NUM: num++,
      NE: rowBase ? rowBase.NE : '',
      NOMBRE: rowNew.NOMBRE || (rowBase ? rowBase.NOMBRE : ''),
      RFC: rfc,
      CURP: rowNew.CURP || '',
      CUENTA: rowBase ? rowBase.CUENTA : '',
      BANCO: rowBase ? rowBase.BANCO : '',
      TELEFONO: rowBase ? rowBase.TELEFONO : '',
      'CORREO ELECTRONICO': rowBase ? rowBase['CORREO ELECTRONICO'] : '',
      'SE ENVIA SOBRE A': rowBase ? rowBase['SE ENVIA SOBRE A'] : '',
      OBSERVACIONES: rowBase ? rowBase.OBSERVACIONES : '',
      TIPOPAGO: rowNew.TIPOPAGO || '',
      CATEGORIA: rowNew.CATEGORIA || '',
      PUESTO: rowNew.PUESTO || '',
      PROYECTO: rowNew.PROYECTO || '',
      NOMINA: rowNew.NOMINA || '',
      DESDE: rowNew.DESDE || '',
      HASTA: rowNew.HASTA || '',
      LIQUIDO: rowNew.LIQUIDO || ''
    });
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

/**
 * Renders a table with the specified data and columns
 */
function renderTable(table, data, columns) {
  const thead = table.querySelector('thead tr');
  const tbody = table.querySelector('tbody');

  // Header
  thead.innerHTML = columns.map(col => `<th>${col}</th>`).join('');

  // Body
  tbody.innerHTML = data.map(row => {
    return '<tr>' + columns.map(col => `<td>${row[col] ?? ''}</td>`).join('') + '</tr>';
  }).join('');
}

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
  const fileName = `Nomina_Fusionada_${dateStr}.xlsx`;

  // Download
  XLSX.writeFile(wb, fileName);
}

// ===============================
// Split Functionality
// ===============================

const btnSplit = document.getElementById('btnSplit');
const btnDownloadAll = document.getElementById('btnDownloadAll');
const splitResultsSection = document.getElementById('splitResults');
const splitTree = document.getElementById('splitTree');

let splitData = null; // Stores the hierarchical split structure

btnSplit.addEventListener('click', performSplit);
btnDownloadAll.addEventListener('click', downloadAllSplitFiles);

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
            
            bancoDiv.innerHTML = `
              <span>🏦 ${banco}</span>
              <span class="count">${rows.length} registros</span>
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
          
          bancoDiv.innerHTML = `
            <span>🏦 ${banco}</span>
            <span class="count">${rows.length} registros</span>
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
      ? `BANAMEX_${project}_${nomina}_${tipoPago}_${dateStr}.xlsx`
      : `BANAMEX_${nomina}_${tipoPago}_${dateStr}.xlsx`;
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
      ? `BANORTE_${project}_${nomina}_${tipoPago}_${dateStr}.xlsx`
      : `BANORTE_${nomina}_${tipoPago}_${dateStr}.xlsx`;
  } else {
    // Standard format for other banks
    ws = XLSX.utils.json_to_sheet(rows, { header: COL_MERGED });
    
    const colWidths = COL_MERGED.map(col => ({ wch: Math.max(col.length, 15) }));
    ws['!cols'] = colWidths;
    
    fileName = SPLIT_BY_PROJECT 
      ? `${project}_${nomina}_${tipoPago}_${banco}_${dateStr}.xlsx`
      : `${nomina}_${tipoPago}_${banco}_${dateStr}.xlsx`;
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
              
              fileName = `BANAMEX_${project}_${nomina}_${tipoPago}_${dateStr}.xlsx`;
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
              
              fileName = `BANORTE_${project}_${nomina}_${tipoPago}_${dateStr}.xlsx`;
            } else {
              // Standard format for other banks
              ws = XLSX.utils.json_to_sheet(rows, { header: COL_MERGED });
              
              const colWidths = COL_MERGED.map(col => ({ wch: Math.max(col.length, 15) }));
              ws['!cols'] = colWidths;
              
              fileName = `${project}_${nomina}_${tipoPago}_${banco}_${dateStr}.xlsx`;
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
            
            fileName = `BANAMEX_${nomina}_${tipoPago}_${dateStr}.xlsx`;
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
            
            fileName = `BANORTE_${nomina}_${tipoPago}_${dateStr}.xlsx`;
          } else {
            // Standard format for other banks
            ws = XLSX.utils.json_to_sheet(rows, { header: COL_MERGED });
            
            const colWidths = COL_MERGED.map(col => ({ wch: Math.max(col.length, 15) }));
            ws['!cols'] = colWidths;
            
            fileName = `${nomina}_${tipoPago}_${banco}_${dateStr}.xlsx`;
          }

          XLSX.utils.book_append_sheet(wb, ws, 'Datos');
          XLSX.writeFile(wb, fileName);
        }
      }
    }
  }
}
