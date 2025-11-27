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
const COL_BASE = ['NUM', 'NOMBRE', 'RFC', 'CUENTA', 'BANCO', 'TELEFONO', 'CORREO ELECTRONICO', 'SE ENVIA SOBRE A'];
const COL_CASH = ['RFC', 'NOMBRE', 'MODALIDAD', 'MONTO'];
const COL_MERGED = ['NUM', 'NOMBRE', 'RFC', 'CURP', 'CUENTA', 'BANCO', 'TELEFONO', 'CORREO ELECTRONICO', 'SE ENVIA SOBRE A', 'CATEGORIA', 'PUESTO', 'PROYECTO', 'NOMINA', 'DESDE', 'HASTA', 'LIQUIDO'];

// Required columns to detect header row
const REQUIRED_BASE_COLS = ['NOMBRE', 'RFC'];
const REQUIRED_NEW_COLS = ['RFC', 'NOMBRE'];
const REQUIRED_CASH_COLS = ['RFC', 'NOMBRE'];
const MAX_HEADER_SEARCH_ROWS = 20; // Search headers in first 20 rows

// DOM Elements
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
// Event Listeners
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

      if (type === 'new') {
        newData = normalizeData(jsonData);
        console.log('New Excel loaded:', newData.length, 'rows');
      } else if (type === 'cash') {
        cashData = normalizeData(jsonData);
        console.log('Cash Excel loaded:', cashData.length, 'rows');
      } else {
        baseData = normalizeData(jsonData);
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
 */
function performMerge() {
  // Create maps by RFC for fast lookup
  const rfcNew = new Map();
  newData.forEach(row => {
    const rfc = (row.RFC || '').toUpperCase();
    if (rfc) rfcNew.set(rfc, row);
  });

  const rfcBase = new Map();
  baseData.forEach(row => {
    const rfc = (row.RFC || '').toUpperCase();
    if (rfc) rfcBase.set(rfc, row);
  });

  // Build set of RFCs from cash payments (these are not considered additions)
  const cashRfcs = new Set();
  if (cashData && cashData.length > 0) {
    cashData.forEach(row => {
      const rfc = (row.RFC || '').toUpperCase().trim();
      if (rfc) cashRfcs.add(rfc);
    });
    console.log('Cash payments excluded:', cashRfcs.size, 'people');
  }

  // Detect additions (in new but not in base, and not in cash payments)
  additions = [];
  rfcNew.forEach((rowNew, rfc) => {
    if (!rfcBase.has(rfc)) {
      // Check if this person is in cash payments (by RFC)
      if (!cashRfcs.has(rfc)) {
        additions.push(rowNew);
      } else {
        console.log('Excluded from additions (cash payment):', rowNew.NOMBRE || rfc);
      }
    }
  });

  // Detect removals (in base but not in new)
  removals = [];
  rfcBase.forEach((rowBase, rfc) => {
    if (!rfcNew.has(rfc)) {
      removals.push(rowBase);
    }
  });

  // Create merged: people in both (updating data) + additions
  mergedData = [];
  let num = 1;

  // First: people in both (update with new data)
  rfcBase.forEach((rowBase, rfc) => {
    if (rfcNew.has(rfc)) {
      const rowNew = rfcNew.get(rfc);
      mergedData.push({
        NUM: num++,
        NOMBRE: rowNew.NOMBRE || rowBase.NOMBRE,
        RFC: rfc,
        CURP: rowNew.CURP || '',
        CUENTA: rowBase.CUENTA || '',
        BANCO: rowBase.BANCO || '',
        TELEFONO: rowBase.TELEFONO || '',
        'CORREO ELECTRONICO': rowBase['CORREO ELECTRONICO'] || '',
        'SE ENVIA SOBRE A': rowBase['SE ENVIA SOBRE A'] || '',
        CATEGORIA: rowNew.CATEGORIA || '',
        PUESTO: rowNew.PUESTO || '',
        PROYECTO: rowNew.PROYECTO || '',
        NOMINA: rowNew.NOMINA || '',
        DESDE: rowNew.DESDE || '',
        HASTA: rowNew.HASTA || '',
        LIQUIDO: rowNew.LIQUIDO || ''
      });
    }
  });

  // Then: add additions
  additions.forEach(rowNew => {
    mergedData.push({
      NUM: num++,
      NOMBRE: rowNew.NOMBRE || '',
      RFC: rowNew.RFC || '',
      CURP: rowNew.CURP || '',
      CUENTA: '',
      BANCO: '',
      TELEFONO: '',
      'CORREO ELECTRONICO': '',
      'SE ENVIA SOBRE A': '',
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
  renderTable(tableRemovals, removals, COL_BASE);
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
