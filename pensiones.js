/**
 * Pensiones Module - Vanilla JS App
 * Compares quincenal Excel with base Excel and generates a merged file.
 */

// Global state
let quincenalData = null;    // Array of objects from quincenal Excel
let basePensionesData = null; // Array of objects from base Excel
let mergedPensionesData = null; // Merge result
let additionsPensiones = [];  // New beneficiaries (altas)
let removalsPensiones = [];   // Removed beneficiaries (bajas)
let efectivosData = [];       // Records without account (cash payments)

// Expected columns
const COL_QUINCENAL = ['PROYECTO', 'RFC', 'NOMBRE', 'BENEFICIARIO', 'FOLIO', 'IMPORTE', 'CVE', 'NOMINA', 'TOTAL DE DESCUENTOS', 'MODALIDAD'];
const COL_BASE_PENSIONES = ['NO.', 'NOMBRE', 'CUENTA', 'NE', 'BANCO'];
const COL_REMOVALS_PENSIONES = ['NO.', 'NOMBRE', 'CUENTA', 'NE', 'BANCO', 'MOTIVO'];

// Required columns to detect header row
const REQUIRED_QUINCENAL_COLS = ['RFC', 'NOMBRE'];
const REQUIRED_BASE_PENSIONES_COLS = ['NOMBRE'];

// Modalidades posibles
const MODALIDADES = {
	'BASE': 'Base',
	'CONTRATO CONFIANZA': 'Contrato confianza',
	'MANDOS MEDIOS': 'Mandos medios',
	'NOMBRAMIENTO CONFIANZA': 'Nombramiento confianza'
};

// DOM Elements
const pensionesSection = document.getElementById('pensionesSection');
const btnBackPensiones = document.getElementById('btnBackPensiones');
const fileQuincenalInput = document.getElementById('fileQuincenal');
const fileBasePensionesInput = document.getElementById('fileBasePensiones');
const fileQuincenalName = document.getElementById('fileQuincenalName');
const fileBasePensionesName = document.getElementById('fileBasePensionesName');
const btnMergePensiones = document.getElementById('btnMergePensiones');
const btnDownloadPensiones = document.getElementById('btnDownloadPensiones');
const resultsPensionesSection = document.getElementById('resultsPensiones');

// Tabs
const tabsPensiones = document.querySelectorAll('#pensionesSection .tab');
const tabAdditionsPensiones = document.getElementById('tabAltasPensiones');
const tabRemovalsPensiones = document.getElementById('tabBajasPensiones');
const tabFinalPensiones = document.getElementById('tabFinalPensiones');

// Stats
const countAdditionsPensiones = document.getElementById('countAltasPensiones');
const countRemovalsPensiones = document.getElementById('countBajasPensiones');
const countTotalPensiones = document.getElementById('countTotalPensiones');

// Tables
const tableAdditionsPensiones = document.getElementById('tableAltasPensiones');
const tableRemovalsPensiones = document.getElementById('tableBajasPensiones');
const tableFinalPensiones = document.getElementById('tableFinalPensiones');

// Split elements
const btnSplitPensiones = document.getElementById('btnSplitPensiones');
const btnDownloadAllPensiones = document.getElementById('btnDownloadAllPensiones');
const btnDownloadEfectivos = document.getElementById('btnDownloadEfectivos');
const splitResultsPensionesSection = document.getElementById('splitResultsPensiones');
const splitTreePensiones = document.getElementById('splitTreePensiones');

let splitPensionesData = null; // Stores the hierarchical split structure

// ===============================
// Event Listeners
// ===============================

if (btnBackPensiones) {
	btnBackPensiones.addEventListener('click', () => {
		window.location.href = 'index.html';
	});
}

fileQuincenalInput.addEventListener('change', (e) => {
	const file = e.target.files[0];
	if (file) {
		fileQuincenalName.textContent = file.name;
		readExcelFile(file, REQUIRED_QUINCENAL_COLS, (data, error) => {
			if (error) {
				alert(error);
				return;
			}
			quincenalData = data;
			console.log('Quincenal Excel loaded:', quincenalData.length, 'rows');
			
			// Validate columns
			const missingCols = validateRequiredColumns(quincenalData, COL_QUINCENAL);
			if (missingCols) {
				const missingColsStr = missingCols.join(', ');
				alert(`⚠️ ADVERTENCIA: El archivo Quincenal no contiene todos los campos requeridos.\n\nCampos faltantes:\n${missingColsStr}\n\nPuedes continuar, pero algunos procesos podrían no funcionar correctamente.`);
			}
			
			updateMergeButtonPensiones();
		});
	}
});

fileBasePensionesInput.addEventListener('change', (e) => {
	const file = e.target.files[0];
	if (file) {
		fileBasePensionesName.textContent = file.name;
		readExcelFile(file, REQUIRED_BASE_PENSIONES_COLS, (data, error) => {
			if (error) {
				alert(error);
				return;
			}
			basePensionesData = data;
			console.log('Base Pensiones Excel loaded:', basePensionesData.length, 'rows');
			
			// Validate columns
			const missingCols = validateRequiredColumns(basePensionesData, COL_BASE_PENSIONES);
			if (missingCols) {
				const missingColsStr = missingCols.join(', ');
				alert(`⚠️ ADVERTENCIA: El archivo Base no contiene todos los campos requeridos.\n\nCampos faltantes:\n${missingColsStr}\n\nPuedes continuar, pero algunos procesos podrían no funcionar correctamente.`);
			}
			
			updateMergeButtonPensiones();
		});
	}
});

btnMergePensiones.addEventListener('click', () => {
	if (quincenalData && basePensionesData) {
		performMergePensiones();
	}
});

btnDownloadPensiones.addEventListener('click', downloadMergedPensiones);

tabsPensiones.forEach(tab => {
	tab.addEventListener('click', () => {
		tabsPensiones.forEach(t => t.classList.remove('active'));
		tab.classList.add('active');
		const target = tab.dataset.tab;
		tabAdditionsPensiones.classList.toggle('hidden', target !== 'altas');
		tabRemovalsPensiones.classList.toggle('hidden', target !== 'bajas');
		tabFinalPensiones.classList.toggle('hidden', target !== 'final');
	});
});

btnSplitPensiones.addEventListener('click', performSplitPensiones);
btnDownloadAllPensiones.addEventListener('click', downloadAllSplitPensionesFiles);
btnDownloadEfectivos.addEventListener('click', downloadEfectivosFile);

// ===============================
// Functions
// ===============================

/**
 * Normalizes a name for comparison (removes accents, converts to uppercase)
 */
function normalizeName(name) {
	if (!name) return '';
	return String(name)
		.normalize('NFD')
		.replace(/[\u0300-\u036f]/g, '')
		.toUpperCase()
		.trim();
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
 * Gets modalidad from row (MODALIDAD field comes directly from quincenal file)
 */
function getModalidad(row) {
	// Get MODALIDAD directly from quincenal data
	if (row.MODALIDAD) {
		const modalidad = String(row.MODALIDAD).toUpperCase().trim();
		// Normalize modalidad name
		for (const [key, value] of Object.entries(MODALIDADES)) {
			if (modalidad.includes(key) || modalidad === key) {
				return value;
			}
		}
		// If not found in MODALIDADES map, return the original value
		return row.MODALIDAD;
	}
	
	// Fallback: try NOMINA field
	if (row.NOMINA) {
		const nomina = String(row.NOMINA).toUpperCase().trim();
		for (const [key, value] of Object.entries(MODALIDADES)) {
			if (nomina.includes(key) || nomina === key) {
				return value;
			}
		}
	}
	
	// Default
	return 'Base';
}

/**
 * Enables the merge button if both files are loaded
 */
function updateMergeButtonPensiones() {
	btnMergePensiones.disabled = !(quincenalData && basePensionesData);
}

/**
 * Performs comparison and merge for pensiones
 * - Additions: people in quincenal (BENEFICIARIO) not in base (NOMBRE)
 * - Removals: people in base (NOMBRE) not in quincenal (BENEFICIARIO)
 * - Merge: base + additions, removing removals, updating data from quincenal
 */
function performMergePensiones() {
	// Create map by normalized NOMBRE for base
	const nombreBase = new Map();
	basePensionesData.forEach(row => {
		const nombre = normalizeName(row.NOMBRE);
		if (nombre) nombreBase.set(nombre, row);
	});

	// For quincenal data, group by normalized BENEFICIARIO
	const beneficiarioQuincenalSet = new Set();
	quincenalData.forEach(row => {
		const beneficiario = normalizeName(row.BENEFICIARIO);
		if (beneficiario) beneficiarioQuincenalSet.add(beneficiario);
	});

	// Detect additions: people in quincenal (BENEFICIARIO) but not in base (NOMBRE)
	additionsPensiones = [];
	const addedBeneficiarios = new Set();
	quincenalData.forEach(rowQuincenal => {
		const beneficiario = normalizeName(rowQuincenal.BENEFICIARIO);
		if (!beneficiario || addedBeneficiarios.has(beneficiario)) return;
		
		const rowBase = nombreBase.get(beneficiario);
		const isInBase = !!rowBase;
		
		if (!isInBase) {
			additionsPensiones.push(rowQuincenal);
			addedBeneficiarios.add(beneficiario);
			console.log('Addition (new beneficiary):', rowQuincenal.BENEFICIARIO || beneficiario);
		}
	});

	// Detect removals: people in base (NOMBRE) but not in quincenal (BENEFICIARIO)
	removalsPensiones = [];
	nombreBase.forEach((rowBase, nombre) => {
		if (!beneficiarioQuincenalSet.has(nombre)) {
			removalsPensiones.push({
				...rowBase,
				MOTIVO: 'No aparece en nómina quincenal'
			});
		}
	});

	// Create merged: iterate through ALL rows in quincenalData
	mergedPensionesData = [];
	efectivosData = [];
	let num = 1;

	// Determine merged columns (all fields from both files)
	const mergedColumns = [
		'NO.', 'NOMBRE', 'RFC', 'BENEFICIARIO', 'CUENTA', 'NE', 'BANCO',
		'PROYECTO', 'FOLIO', 'IMPORTE', 'CVE', 'NOMINA', 'TOTAL DE DESCUENTOS', 'MODALIDAD'
	];

	quincenalData.forEach(rowQuincenal => {
		const beneficiario = normalizeName(rowQuincenal.BENEFICIARIO || '');
		const rowBase = beneficiario ? nombreBase.get(beneficiario) : null;
		
		const cuenta = rowBase && rowBase.CUENTA ? String(rowBase.CUENTA).trim() : '';
		const hasAccount = cuenta.length > 0;
		
		const mergedRow = {
			'NO.': num++,
			'NOMBRE': rowBase ? rowBase.NOMBRE : (rowQuincenal.NOMBRE || ''),
			'RFC': rowQuincenal.RFC || '',
			'BENEFICIARIO': rowQuincenal.BENEFICIARIO || '',
			'CUENTA': cuenta,
			'NE': rowBase ? (rowBase.NE || '') : '',
			'BANCO': rowBase ? (rowBase.BANCO || '') : '',
			'PROYECTO': rowQuincenal.PROYECTO || '',
			'FOLIO': rowQuincenal.FOLIO || '',
			'IMPORTE': rowQuincenal.IMPORTE || 0,
			'CVE': rowQuincenal.CVE || '',
			'NOMINA': rowQuincenal.NOMINA || '',
			'TOTAL DE DESCUENTOS': rowQuincenal['TOTAL DE DESCUENTOS'] || 0,
			'MODALIDAD': getModalidad(rowQuincenal)
		};
		
		mergedPensionesData.push(mergedRow);
		
		// If no account, add to efectivos
		if (!hasAccount) {
			efectivosData.push(mergedRow);
		}
	});

	// Store merged columns for later use
	window.COL_MERGED_PENSIONES = mergedColumns;

	// Display results
	displayResultsPensiones();
}

/**
 * Displays results in the UI
 */
function displayResultsPensiones() {
	resultsPensionesSection.classList.remove('hidden');

	// Stats
	countAdditionsPensiones.textContent = additionsPensiones.length;
	countRemovalsPensiones.textContent = removalsPensiones.length;
	countTotalPensiones.textContent = mergedPensionesData.length;

	// Additions table
	renderTable(tableAdditionsPensiones, additionsPensiones, COL_QUINCENAL);
	tabAdditionsPensiones.querySelector('.empty-msg').classList.toggle('hidden', additionsPensiones.length > 0);

	// Removals table
	renderTable(tableRemovalsPensiones, removalsPensiones, COL_REMOVALS_PENSIONES);
	tabRemovalsPensiones.querySelector('.empty-msg').classList.toggle('hidden', removalsPensiones.length > 0);

	// Final table
	renderTable(tableFinalPensiones, mergedPensionesData, window.COL_MERGED_PENSIONES);

	// Scroll to results
	resultsPensionesSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * Downloads the merged Excel file
 */
function downloadMergedPensiones() {
	if (!mergedPensionesData || mergedPensionesData.length === 0) {
		alert('No hay datos para descargar');
		return;
	}

	const today = new Date();
	const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');
	const fileName = `Pensiones_Fusionadas_${dateStr}.xlsx`;

	downloadExcel(mergedPensionesData, window.COL_MERGED_PENSIONES, 'Pensiones Fusionadas', fileName);
}

/**
 * Splits the merged data by Modalidad -> Banco
 */
function performSplitPensiones() {
	if (!mergedPensionesData || mergedPensionesData.length === 0) {
		alert('No hay datos para separar. Primero realiza la fusión.');
		return;
	}

	// Initialize split structure
	splitPensionesData = {};

	// Group data hierarchically
	mergedPensionesData.forEach(row => {
		// Level 2: Banco - Skip rows without bank info
		const banco = (row.BANCO || '').toUpperCase().trim();
		if (!banco) {
			console.log('Split skipped (no bank):', row.NOMBRE);
			return;
		}
		
		// Level 1: Modalidad
		const modalidad = (row.MODALIDAD || 'Base').toUpperCase().trim();
		
		// Initialize nested structure if needed
		if (!splitPensionesData[modalidad]) {
			splitPensionesData[modalidad] = {};
		}
		if (!splitPensionesData[modalidad][banco]) {
			splitPensionesData[modalidad][banco] = [];
		}

		splitPensionesData[modalidad][banco].push(row);
	});

	// Display results
	displaySplitResultsPensiones();
}

/**
 * Displays the split results in a tree structure
 */
function displaySplitResultsPensiones() {
	splitResultsPensionesSection.classList.remove('hidden');
	splitTreePensiones.innerHTML = '';

	for (const [modalidad, bancos] of Object.entries(splitPensionesData)) {
		// Skip empty modalidades
		if (Object.keys(bancos).length === 0) continue;

		const modalidadDiv = document.createElement('div');
		modalidadDiv.className = 'split-project';

		const modalidadName = document.createElement('div');
		modalidadName.className = 'split-project-name';
		modalidadName.textContent = `📋 ${modalidad}`;
		modalidadDiv.appendChild(modalidadName);

		for (const [banco, rows] of Object.entries(bancos)) {
			const bancoDiv = document.createElement('div');
			bancoDiv.className = 'split-banco';
			
			const totalAmount = calculateTotalAmount(rows, 'IMPORTE');
			const formattedAmount = formatCurrency(totalAmount);
			
			bancoDiv.innerHTML = `
				<span>🏦 ${banco}</span>
				<span class="count">${rows.length} registros</span>
				<span class="amount">${formattedAmount}</span>
				<button class="btn-download-single" data-modalidad="${modalidad}" data-banco="${banco}">⬇️</button>
			`;
			
			modalidadDiv.appendChild(bancoDiv);
		}

		splitTreePensiones.appendChild(modalidadDiv);
	}

	// Add click handlers for individual download buttons
	splitTreePensiones.querySelectorAll('.btn-download-single').forEach(btn => {
		btn.addEventListener('click', (e) => {
			const modalidad = e.target.dataset.modalidad;
			const banco = e.target.dataset.banco;
			downloadSingleSplitPensionesFile(modalidad, banco);
		});
	});

	splitResultsPensionesSection.scrollIntoView({ behavior: 'smooth' });
}

/**
 * Downloads a single split file
 * For BANAMEX and BANORTE: special formats with specific columns
 */
function downloadSingleSplitPensionesFile(modalidad, banco) {
	const rows = splitPensionesData[modalidad]?.[banco];
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
			'Importe': row.IMPORTE || 0,
			'Nombre/Razón Social': row.NOMBRE || row.BENEFICIARIO || '',
			'Ref. Num.': index + 1,
			'Ref. AlfN.': conceptoBancario
		}));
		
		const banamexHeaders = ['Tipo de Cuenta', 'Cuenta', 'Importe', 'Nombre/Razón Social', 'Ref. Num.', 'Ref. AlfN.'];
		ws = XLSX.utils.json_to_sheet(banamexData, { header: banamexHeaders });
		
		const colWidths = banamexHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
		ws['!cols'] = colWidths;
		
		fileName = `BANAMEX_${modalidad}_${dateStr}.xlsx`;
	} else if (banco.toUpperCase() === 'BANORTE') {
		// Transform data to Banorte format
		const banorteData = rows.map((row) => ({
			'NO. EMPLEADO': row.NE || '',
			'NOMBRE': row.NOMBRE || row.BENEFICIARIO || '',
			'IMPORTE': row.IMPORTE || 0,
			'NO. BANCO RECEPTOR': '072',
			'TIPO DE CUENTA': '01',
			'CUENTA': row.CUENTA || ''
		}));
		
		const banorteHeaders = ['NO. EMPLEADO', 'NOMBRE', 'IMPORTE', 'NO. BANCO RECEPTOR', 'TIPO DE CUENTA', 'CUENTA'];
		ws = XLSX.utils.json_to_sheet(banorteData, { header: banorteHeaders });
		
		const colWidths = banorteHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
		ws['!cols'] = colWidths;
		
		fileName = `BANORTE_${modalidad}_${dateStr}.xlsx`;
	} else {
		// Standard format for other banks
		ws = XLSX.utils.json_to_sheet(rows, { header: window.COL_MERGED_PENSIONES });
		
		const colWidths = window.COL_MERGED_PENSIONES.map(col => ({ wch: Math.max(col.length, 15) }));
		ws['!cols'] = colWidths;
		
		fileName = `${modalidad}_${banco}_${dateStr}.xlsx`;
	}

	XLSX.utils.book_append_sheet(wb, ws, 'Datos');
	XLSX.writeFile(wb, fileName);
}

/**
 * Downloads all split files at once
 * Applies special Banamex and Banorte formats when applicable
 */
function downloadAllSplitPensionesFiles() {
	if (!splitPensionesData) {
		alert('No hay datos para descargar');
		return;
	}

	const today = new Date();
	const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');
	
	// Get payroll period (calculated automatically)
	const payrollPeriod = getPayrollPeriod();
	const conceptoBancario = payrollPeriod.display;

	for (const [modalidad, bancos] of Object.entries(splitPensionesData)) {
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
					'Importe': row.IMPORTE || 0,
					'Nombre/Razón Social': row.NOMBRE || row.BENEFICIARIO || '',
					'Ref. Num.': index + 1,
					'Ref. AlfN.': conceptoBancario
				}));
				
				const banamexHeaders = ['Tipo de Cuenta', 'Cuenta', 'Importe', 'Nombre/Razón Social', 'Ref. Num.', 'Ref. AlfN.'];
				ws = XLSX.utils.json_to_sheet(banamexData, { header: banamexHeaders });
				
				const colWidths = banamexHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
				ws['!cols'] = colWidths;
				
				fileName = `BANAMEX_${modalidad}_${dateStr}.xlsx`;
			} else if (banco.toUpperCase() === 'BANORTE') {
				// Transform data to Banorte format
				const banorteData = rows.map((row) => ({
					'NO. EMPLEADO': row.NE || '',
					'NOMBRE': row.NOMBRE || row.BENEFICIARIO || '',
					'IMPORTE': row.IMPORTE || 0,
					'NO. BANCO RECEPTOR': '072',
					'TIPO DE CUENTA': '01',
					'CUENTA': row.CUENTA || ''
				}));
				
				const banorteHeaders = ['NO. EMPLEADO', 'NOMBRE', 'IMPORTE', 'NO. BANCO RECEPTOR', 'TIPO DE CUENTA', 'CUENTA'];
				ws = XLSX.utils.json_to_sheet(banorteData, { header: banorteHeaders });
				
				const colWidths = banorteHeaders.map(col => ({ wch: Math.max(col.length, 20) }));
				ws['!cols'] = colWidths;
				
				fileName = `BANORTE_${modalidad}_${dateStr}.xlsx`;
			} else {
				// Standard format for other banks
				ws = XLSX.utils.json_to_sheet(rows, { header: window.COL_MERGED_PENSIONES });
				
				const colWidths = window.COL_MERGED_PENSIONES.map(col => ({ wch: Math.max(col.length, 15) }));
				ws['!cols'] = colWidths;
				
				fileName = `${modalidad}_${banco}_${dateStr}.xlsx`;
			}

			XLSX.utils.book_append_sheet(wb, ws, 'Datos');
			XLSX.writeFile(wb, fileName);
		}
	}
}

/**
 * Downloads the efectivos (cash payments) file
 */
function downloadEfectivosFile() {
	if (!efectivosData || efectivosData.length === 0) {
		alert('No hay registros sin cuenta para descargar');
		return;
	}

	const today = new Date();
	const dateStr = today.toISOString().slice(0, 10).replace(/-/g, '');
	const fileName = `Pensiones_Efectivos_${dateStr}.xlsx`;

	downloadExcel(efectivosData, window.COL_MERGED_PENSIONES, 'Efectivos', fileName);
}

