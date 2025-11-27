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
const COL_BASE = ['NUM', 'NOMBRE', 'RFC', 'CUENTA', 'BANCO', 'TELEFONO', 'CORREO ELECTRONICO', 'SE ENVIA SOBRE A', 'TIPOPAGO'];
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

  // Detect additions: people in new but not in base, and not in cash payments
  // We need to track unique RFCs for additions (not all rows)
  additions = [];
  const addedRfcs = new Set();
  newData.forEach(rowNew => {
    const rfc = (rowNew.RFC || '').toUpperCase();
    if (rfc && !rfcBase.has(rfc) && !addedRfcs.has(rfc)) {
      // Check if this person is in cash payments (by RFC)
      if (!cashRfcs.has(rfc)) {
        additions.push(rowNew);
        addedRfcs.add(rfc);
      } else {
        console.log('Excluded from additions (cash payment):', rowNew.NOMBRE || rfc);
      }
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
  // Skip rows without BANCO (SIN_BANCO)
  newData.forEach(rowNew => {
    const rfc = (rowNew.RFC || '').toUpperCase();
    const rowBase = rfcBase.get(rfc);
    const banco = rowBase ? (rowBase.BANCO || '').toUpperCase().trim() : '';
    
    // Skip if no bank info
    if (!banco) {
      console.log('Skipped (no bank):', rowNew.NOMBRE || rfc);
      return;
    }
    
    mergedData.push({
      NUM: num++,
      NOMBRE: rowNew.NOMBRE || (rowBase ? rowBase.NOMBRE : ''),
      RFC: rfc,
      CURP: rowNew.CURP || '',
      CUENTA: rowBase ? rowBase.CUENTA : '',
      BANCO: banco,
      TELEFONO: rowBase ? rowBase.TELEFONO : '',
      'CORREO ELECTRONICO': rowBase ? rowBase['CORREO ELECTRONICO'] : '',
      'SE ENVIA SOBRE A': rowBase ? rowBase['SE ENVIA SOBRE A'] : '',
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
 * Splits the merged data by Proyecto -> Nomina -> TipoPago -> Banco
 */
function performSplit() {
  if (!mergedData || mergedData.length === 0) {
    alert('No hay datos para separar. Primero realiza la fusión.');
    return;
  }

  // Initialize split structure
  splitData = {
    'JARDIN': {},
    'OTROS': {}
  };

  // Group data hierarchically
  mergedData.forEach(row => {
    // Level 1: Proyecto (Jardin vs Otros)
    const projectGroup = isJardin(row.PROYECTO) ? 'JARDIN' : 'OTROS';
    
    // Level 2: Nomina
    const nomina = (row.NOMINA || 'SIN_NOMINA').toUpperCase().trim();
    
    // Level 3: TipoPago
    const tipoPago = (row.TIPOPAGO || 'SIN_TIPOPAGO').toUpperCase().trim();
    
    // Level 4: Banco
    const banco = (row.BANCO || 'SIN_BANCO').toUpperCase().trim();

    // Initialize nested structure if needed
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

  const projectIcons = { 'JARDIN': '🌳', 'OTROS': '🏢' };

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

  // Add click handlers for individual download buttons
  splitTree.querySelectorAll('.btn-download-single').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const project = e.target.dataset.project;
      const nomina = e.target.dataset.nomina;
      const tipoPago = e.target.dataset.tipopago;
      const banco = e.target.dataset.banco;
      downloadSingleSplitFile(project, nomina, tipoPago, banco);
    });
  });

  splitResultsSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * Downloads a single split file
 * For BANAMEX: special format with specific columns
 */
function downloadSingleSplitFile(project, nomina, tipoPago, banco) {
  const rows = splitData[project]?.[nomina]?.[tipoPago]?.[banco];
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
    // Get payroll period from selectors
    const quincena = document.getElementById('selectQuincena').value;
    const mes = document.getElementById('selectMes').value;
    const conceptoBancario = `${quincena} Nomina de ${mes}`;
    
    // Transform data to Banamex format
    const banamexData = rows.map((row, index) => ({
      'Tipo de Cuenta': 'Tarjeta',
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
  
  // Get payroll period for Banamex files
  const quincena = document.getElementById('selectQuincena').value;
  const mes = document.getElementById('selectMes').value;
  const conceptoBancario = `${quincena} Nomina de ${mes}`;

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
              'Tipo de Cuenta': 'Tarjeta',
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
}
