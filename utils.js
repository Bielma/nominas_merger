/**
 * Utility functions shared across Nominas and Pensiones modules
 */

const MAX_HEADER_SEARCH_ROWS = 20; // Search headers in first 20 rows

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
 * Normalizes object keys (trim and uppercase)
 * @param {Array} data - Array of objects to normalize
 * @returns {Array} - Normalized array
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
 * Reads an Excel file and converts it to an array of objects
 * @param {File} file - File object from input
 * @param {string[]} requiredCols - Required columns to detect header row
 * @param {Function} callback - Callback function(data, error)
 */
function readExcelFile(file, requiredCols, callback) {
	const reader = new FileReader();
	reader.onload = (e) => {
		try {
			const data = new Uint8Array(e.target.result);
			const workbook = XLSX.read(data, { type: 'array' });
			const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
			
			const headerRow = findHeaderRow(firstSheet, requiredCols);
			
			if (headerRow === -1) {
				const colNames = requiredCols.join(', ');
				callback(null, `No se encontraron los encabezados (${colNames}) en las primeras ${MAX_HEADER_SEARCH_ROWS} filas del archivo.`);
				return;
			}
			
			const jsonData = XLSX.utils.sheet_to_json(firstSheet, { 
				defval: '',
				range: headerRow 
			});

			const normalizedData = normalizeData(jsonData);
			callback(normalizedData, null);
		} catch (err) {
			callback(null, 'Error reading file: ' + err.message);
		}
	};
	reader.readAsArrayBuffer(file);
}

/**
 * Validates that the data contains all required columns
 * @param {Array} data - Normalized data array
 * @param {string[]} expectedCols - Array of expected column names
 * @returns {Array|null} - Array of missing columns or null if all are present
 */
function validateRequiredColumns(data, expectedCols) {
	if (!data || data.length === 0) {
		return null; // Empty data, validation will happen elsewhere
	}

	// Get all available keys from the first row
	const availableKeys = new Set(Object.keys(data[0]));

	// Find missing columns
	const missingCols = expectedCols.filter(col => !availableKeys.has(col));

	return missingCols.length > 0 ? missingCols : null;
}

/**
 * Calculates the sum of a numeric field from an array of rows
 * @param {Array} rows - Array of row objects
 * @param {string} fieldName - Name of the field to sum
 * @returns {number} - Sum of all field values
 */
function calculateTotalAmount(rows, fieldName = 'LIQUIDO') {
	return rows.reduce((total, row) => {
		const value = parseFloat(row[fieldName] || 0);
		return total + (isNaN(value) ? 0 : value);
	}, 0);
}

/**
 * Formats a number as Mexican peso currency
 * @param {number} amount - Amount to format
 * @returns {string} - Formatted currency string
 */
function formatCurrency(amount) {
	return new Intl.NumberFormat('es-MX', {
		style: 'currency',
		currency: 'MXN',
		minimumFractionDigits: 2,
		maximumFractionDigits: 2
	}).format(amount);
}

/**
 * Renders a table with the specified data and columns
 * @param {HTMLElement} table - Table element
 * @param {Array} data - Data array
 * @param {string[]} columns - Column names array
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
 * Downloads data as Excel file
 * @param {Array} data - Data to export
 * @param {string[]} columns - Column headers
 * @param {string} sheetName - Sheet name
 * @param {string} fileName - File name
 */
function downloadExcel(data, columns, sheetName, fileName) {
	if (!data || data.length === 0) {
		alert('No hay datos para descargar');
		return;
	}

	const wb = XLSX.utils.book_new();
	const ws = XLSX.utils.json_to_sheet(data, { header: columns });

	// Adjust column widths
	const colWidths = columns.map(col => ({ wch: Math.max(col.length, 15) }));
	ws['!cols'] = colWidths;

	XLSX.utils.book_append_sheet(wb, ws, sheetName);
	XLSX.writeFile(wb, fileName);
}

