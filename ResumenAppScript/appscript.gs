/** @OnlyCurrentDoc */

const REGISTRO_SOURCE_COLUMNS = 5;
const REGISTRO_TARGET_DATA_COLUMNS = 6;
const SALDOS_DATA_COLUMNS = 3;
const SALDOS_SOURCE_COLUMNS = 2;
const CAPACITY_BUFFER_ROWS = 10;

function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('=RESUMEN=')
        .addItem('Incorporar datos', 'addNewData')
        .addItem('Aplicar reglas', 'processDataByRules')
        .addItem('Generar cuenta de gastos', 'generateExpenseAccountPivot')
        .addToUi();

    syncRegistroContableNamedRange_();
    syncSaldosNamedRanges_();
    syncCuentasAndReglasNamedRanges_();
    setupCurrencyConvertorTable_();
    refreshResumenCsvLinkInCell();
}

function setupCurrencyConvertorTable_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const startDateRange = ss.getRangeByName('StartDate');
    const endDateRange = ss.getRangeByName('EndDate');

    if (!startDateRange || !endDateRange) {
        throw new Error('Named ranges "StartDate" and "EndDate" are required.');
    }

    const startDate = parseNamedDateValue_(startDateRange.getValue(), 'StartDate');
    const endDate = parseNamedDateValue_(endDateRange.getValue(), 'EndDate');
    if (startDate.getTime() > endDate.getTime()) {
        throw new Error('StartDate must be less than or equal to EndDate.');
    }

    const sheetName = 'CurrencyConvertor';
    const headers = [
        'Date',
        'HUF',
        'RON',
        'USD',
        'CHF',
        'EUR2HUF',
        'EUR2RON',
        'EUR2USD',
        'EUR2CHF',
        'HUF2EUR',
        'RON2EUR',
        'USD2EUR',
        'CHF2EUR'
    ];

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        sheet = ss.insertSheet(sheetName);
    }

    const dates = buildDateSeries_(startDate, endDate);
    const tableRows = dates.length + 1;
    const tableCols = headers.length;

    ensureSheetSize_(sheet, tableRows, tableCols);

    sheet.getRange(1, 1, 1, tableCols).setValues([headers]);

    if (dates.length > 0) {
        const isoDateTextValues = dates.map(d => [formatIsoDateText_(d)]);
        sheet.getRange(2, 1, dates.length, 1).setNumberFormat('@');
        sheet.getRange(2, 1, dates.length, 1).setValues(isoDateTextValues);

        sheet
            .getRange(2, 2, dates.length, 4)
            .setFormulaR1C1('=INDEX(GOOGLEFINANCE("CURRENCY:EUR"&R1C,"price",RC1),2,2)');

        sheet.getRange(2, 6, 1, 4).setFormulaR1C1('=RC[-4]');
        if (dates.length > 1) {
            sheet.getRange(3, 6, dates.length - 1, 4).setFormulaR1C1('=IFNA(RC[-4],R[-1]C)');
        }

        sheet.getRange(2, 10, dates.length, 4).setFormulaR1C1('=1/RC[-4]');
    }

    clearRowsBelowTable_(sheet, tableRows, tableCols);
    ss.setNamedRange('CurrenciesConvertorTable', sheet.getRange(1, 1, tableRows, tableCols));
}

function parseNamedDateValue_(value, rangeName) {
    const isDateObject = Object.prototype.toString.call(value) === '[object Date]';
    if (isDateObject && !Number.isNaN(value.getTime())) {
        return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }

    const parsed = new Date(String(value || '').trim());
    if (Number.isNaN(parsed.getTime())) {
        throw new Error(`Invalid date value in named range "${rangeName}".`);
    }

    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function buildDateSeries_(startDate, endDate) {
    const dates = [];
    const cursor = new Date(startDate.getTime());

    while (cursor.getTime() <= endDate.getTime()) {
        dates.push(new Date(cursor.getTime()));
        cursor.setDate(cursor.getDate() + 1);
    }

    return dates;
}

function formatIsoDateText_(date) {
    const year = String(date.getFullYear());
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function ensureSheetSize_(sheet, minRows, minCols) {
    const currentRows = sheet.getMaxRows();
    const currentCols = sheet.getMaxColumns();

    if (currentRows < minRows) {
        sheet.insertRowsAfter(currentRows, minRows - currentRows);
    }
    if (currentCols < minCols) {
        sheet.insertColumnsAfter(currentCols, minCols - currentCols);
    }
}

function clearRowsBelowTable_(sheet, tableRows, tableCols) {
    const maxRows = sheet.getMaxRows();
    if (maxRows > tableRows) {
        sheet.getRange(tableRows + 1, 1, maxRows - tableRows, tableCols).clearContent();
    }
}

/**
 * Processes data based on specific word-matching rules.
 */
function processDataByRules() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Get ranges (ensure these Named Ranges exist in your sheet)
    const dataRange = ss.getRangeByName("EntradaDeDatosRange");
    const rulesRange = ss.getRangeByName("ReglasRange");

    if (!dataRange || !rulesRange) {
        SpreadsheetApp.getUi().alert("Error: Ensure 'EntradaDeDatosRange' and 'ReglasRange' are defined.");
        return;
    }

    const dataValues = dataRange.getValues();
    const rulesValues = rulesRange.getValues();
    const rulesBody = rulesValues.slice(1);

    // Remove header from data
    const dataBody = dataValues.slice(1);

    // 2. Prepare Rules (Clean them once for efficiency)
    const preparedRules = rulesBody
        .filter(row => row[0] !== "") // Skip empty rules
        .map(row => {
            return {
                words: cleanText(row[0]).split(new RegExp('\\s+')).filter(w => w.length > 0),
                valToCol1: row[1],
                valToCol4: row[2]
            };
        });

    // 3. Iterate through Data Rows
    for (let i = 0; i < dataBody.length; i++) {
        let row = dataBody[i];

        // Stop if the second column (index 1) is empty
        if (!row[1]) break;

        let targetText = cleanText(row[4] || ""); // Fifth column (index 4)

        // Check each rule
        for (let rule of preparedRules) {
            if (rule.words.length === 0) continue;

            if (isMatchInOrder(targetText, rule.words)) {
                // Match found: Update Column 1 (index 0) and Column 4 (index 3)
                // Note: +1 to i because of slice(1), +1 again because Sheets is 1-indexed
                dataRange.getCell(i + 2, 1).setValue(rule.valToCol1);
                dataRange.getCell(i + 2, 4).setValue(rule.valToCol4);
                break; // Stop checking other rules for this row if a match is found
            }
        }
    }
}

/**
 * Normalizes text: removes diacritics, removes numbers,
 * replaces non-alphanumeric with spaces, and lowercase.
 */
function cleanText(text) {
    if (typeof text !== 'string') text = text.toString();

    return text
        .normalize("NFD").replace(new RegExp('[\\u0300-\\u036f]', 'g'), "") // Remove diacritics
        .replace(new RegExp('[0-9]', 'g'), "")           // Remove numbers
        .replace(new RegExp('[^a-zA-Z]', 'g'), " ")      // Non-alphanumeric to spaces
        .toLowerCase()
        .trim();
}

/**
 * Checks if the array of ruleWords appears in the targetText in the correct order.
 */
function isMatchInOrder(targetText, ruleWords) {
    const targetWords = targetText.split(new RegExp('\\s+')).filter(w => w.length > 0);

    let currentIdx = 0;
    for (let word of ruleWords) {
        // Find the next word in the sequence
        let foundIdx = targetWords.indexOf(word, currentIdx);
        if (foundIdx === -1) return false; // Word not found or out of order
        currentIdx = foundIdx + 1; // Move pointer to next word
    }
    return true;
}
function addNewData() {

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if(ss.getRangeByName('ValidDataEntryRange').getValue() !== true) {
        throw new Error('Datos invalidos. Corregir antes de incorporar');
    }

    // --- SOURCE: EntradaDeDatos ---
    const srcRangeOrig = ss.getRangeByName('EntradaDeDatosRange');
    const srcRange = srcRangeOrig.offset(1, 0, srcRangeOrig.getNumRows() - 1, srcRangeOrig.getNumColumns());
    if (!srcRange) throw new Error('No se encontró el rango con nombre "EntradaDeDatosRange".');
    const srcNumCols = srcRange.getNumColumns();
    const srcValues = srcRange.getValues();

    // Find how many contiguous non-blank rows from the top
    let rowsToCopy = 0;
    for (let r = 0; r < srcValues.length; r++) {
        if (isBlankRow(srcValues[r])) break; // stop at the first blank row
        rowsToCopy++;
    }
    if (rowsToCopy === 0) {
        SpreadsheetApp.getUi().alert('No hay filas no vacías al fin de "EntradaDeDatos".');
        return;
    }

    const dataToCopy = srcValues
        .slice(0, rowsToCopy)
        .map(row => {
            const paddedRow = new Array(REGISTRO_SOURCE_COLUMNS).fill('');
            const colsToCopy = Math.min(srcNumCols, REGISTRO_SOURCE_COLUMNS);
            for (let c = 0; c < colsToCopy; c++) {
                paddedRow[c] = row[c];
            }
            return paddedRow;
        });

    let dstRangeOrig = ss.getRangeByName('RegistroContableRange');
    let dstRange = dstRangeOrig.offset(1, 0, dstRangeOrig.getNumRows() - 1, dstRangeOrig.getNumColumns());
    if (!dstRange) throw new Error('No se encontró el rango con nombre "RegistroContable".');

    const dstNumCols = dstRange.getNumColumns();

    if(dstNumCols < REGISTRO_TARGET_DATA_COLUMNS) {
        throw new Error('Wrong number of columns');
    }

    const registroSetup = ensureCapacityInNamedRange(
        'RegistroContableRange',
        REGISTRO_TARGET_DATA_COLUMNS,
        rowsToCopy,
        CAPACITY_BUFFER_ROWS
    );
    dstRange = registroSetup.dataRange;
    const firstBlankRowIdx = registroSetup.firstBlankRowIdx;

    // Write transaction values
    dstRange
        .offset(firstBlankRowIdx, 0, rowsToCopy, REGISTRO_SOURCE_COLUMNS)
        .setValues(dataToCopy);

    // Write bank account on the next column
    const accountId = ss.getRangeByName('CuentaEntradaDatosID').getValue();
    dstRange
        .offset(firstBlankRowIdx, REGISTRO_SOURCE_COLUMNS, rowsToCopy, 1)
        .setValue(accountId);

    // Now copy the balance rows
    const balanceEntryRangeOrig = ss.getRangeByName('SaldosEntradaRange');
    const balanceEntryRange =  balanceEntryRangeOrig.offset(1, 0, balanceEntryRangeOrig.getNumRows() - 1, balanceEntryRangeOrig.getNumColumns());

    const balanceValuesToCopy = findValuesToCopy(balanceEntryRange, SALDOS_SOURCE_COLUMNS);

    const saldosSetup = ensureCapacityInNamedRange(
        'SaldosRange',
        SALDOS_DATA_COLUMNS,
        balanceValuesToCopy.length,
        CAPACITY_BUFFER_ROWS
    );
    const balancesRange = saldosSetup.dataRange;
    const firstBlankBalanceRow = saldosSetup.firstBlankRowIdx;

    balancesRange
        .offset(firstBlankBalanceRow, 0, balanceValuesToCopy.length, SALDOS_SOURCE_COLUMNS)
        .setValues(balanceValuesToCopy);

    // Now set balances accountId
    balancesRange
        .offset(firstBlankBalanceRow, SALDOS_SOURCE_COLUMNS, balanceValuesToCopy.length, 1)
        .setValue(accountId);

    // Now clear source ranges
    balanceEntryRange.offset(0, 0, balanceValuesToCopy.length, SALDOS_SOURCE_COLUMNS).clearContent();
    srcRange.offset(0, 0, rowsToCopy, REGISTRO_SOURCE_COLUMNS).clearContent();

    ss.getRangeByName('CuentaEntradaDatosID').clearContent();
}

/**
 * Inserts rows before the last row of a named range using native sheet behavior.
 *
 * @param {string} rangeName Named range to extend.
 * @param {number} rowsToInsert Number of rows to insert.
 */
function insertRowsBeforeLastRowInNamedRange(rangeName, rowsToInsert) {
    if (!Number.isInteger(rowsToInsert) || rowsToInsert < 1) {
        throw new Error('rowsToInsert must be an integer greater than 0.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetRange = ss.getRangeByName(rangeName);
    if (!targetRange) {
        throw new Error(`No se encontró el rango con nombre "${rangeName}".`);
    }

    const sheet = targetRange.getSheet();
    const lastRowInRange = targetRange.getRow() + targetRange.getNumRows() - 1;
    sheet.insertRowsBefore(lastRowInRange, rowsToInsert);
}

/**
 * Ensures there are at least rowsToWrite + bufferRows available from the first writable row.
 * Returns the refreshed writable range and insertion index.
 */
function ensureCapacityInNamedRange(rangeName, relevantColumnsNumber, rowsToWrite, bufferRows) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let tableRange = ss.getRangeByName(rangeName);
    if (!tableRange) {
        throw new Error(`No se encontró el rango con nombre "${rangeName}".`);
    }

    let dataRange = tableRange.offset(1, 0, tableRange.getNumRows() - 1, tableRange.getNumColumns());
    let dataValues = dataRange.getValues();

    let firstBlankRowIdx = findFirstBlankRowIdx(dataValues, relevantColumnsNumber);
    if (firstBlankRowIdx === -1) {
        // Table is full: create one writable row while preserving the previous last-row data.
        createWritableRowAtEnd(rangeName, relevantColumnsNumber);
        tableRange = ss.getRangeByName(rangeName);
        dataRange = tableRange.offset(1, 0, tableRange.getNumRows() - 1, tableRange.getNumColumns());
        dataValues = dataRange.getValues();
        firstBlankRowIdx = findFirstBlankRowIdx(dataValues, relevantColumnsNumber);
        if (firstBlankRowIdx === -1) {
            firstBlankRowIdx = dataValues.length;
        }
    }

    const minRequiredRows = rowsToWrite + bufferRows;
    const remainingRows = dataValues.length - firstBlankRowIdx;
    if (remainingRows < minRequiredRows) {
        insertRowsBeforeLastRowInNamedRange(rangeName, minRequiredRows);
        tableRange = ss.getRangeByName(rangeName);
        dataRange = tableRange.offset(1, 0, tableRange.getNumRows() - 1, tableRange.getNumColumns());
    }

    return {
        dataRange: dataRange,
        firstBlankRowIdx: firstBlankRowIdx
    };
}

/**
 * When a table is fully occupied, insert one row before its last row and move
 * the last row up so the table ends with a writable blank row.
 */
function createWritableRowAtEnd(rangeName, dataColumnsNumber) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetRange = ss.getRangeByName(rangeName);
    if (!targetRange) {
        throw new Error(`No se encontró el rango con nombre "${rangeName}".`);
    }

    const sheet = targetRange.getSheet();
    const startCol = targetRange.getColumn();
    const numCols = targetRange.getNumColumns();
    if (!Number.isInteger(dataColumnsNumber) || dataColumnsNumber < 1 || dataColumnsNumber > numCols) {
        throw new Error(`Invalid dataColumnsNumber for range "${rangeName}".`);
    }
    const lastRowInRange = targetRange.getRow() + targetRange.getNumRows() - 1;

    sheet.insertRowsBefore(lastRowInRange, 1);

    // Old last row shifted down by 1; copy it up and clear the old location.
    const oldLastRowAfterInsert = lastRowInRange + 1;
    const sourceRange = sheet.getRange(oldLastRowAfterInsert, startCol, 1, dataColumnsNumber);
    const destinationRange = sheet.getRange(lastRowInRange, startCol, 1, dataColumnsNumber);
    destinationRange.setValues(sourceRange.getValues());
    sourceRange.clearContent();
}

function findFirstBlankRowIdx(range, relevantColumnsNumber) {
    console.log(range.length);
    // Find first completely blank row within the destination table
    let firstBlankRowIdx = -1;
    for (let r = 0; r < range.length; r++) {
        if (range[r].slice(0, relevantColumnsNumber).every(v => v === '' || v === null)) {
            firstBlankRowIdx = r;
            break;
        }
    }

    return firstBlankRowIdx;
}

function findValuesToCopy(origRange, numRelevantColumns) {
    let rowsToCopy = 0;
    const values = origRange.getValues();
    for (let r = 0; r < values.length; r++) {
        if (isBlankRow(values[r].slice(0, numRelevantColumns))) break; // stop at the first blank row
        rowsToCopy++;
    }

    return origRange.offset(0, 0, rowsToCopy, numRelevantColumns).getValues();
}

function isBlankRow(row) {
    return row.every(v => v === '' || v === null);
}

function buildSheetCsvExportUrl_(spreadsheetId, sheetId) {
    return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv&gid=${sheetId}`;
}

/**
 * Regenerates and writes the Resumen CSV link into Resumen!D1.
 */
function refreshResumenCsvLinkInCell() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Resumen');
    if (!sheet) {
        return;
    }

    setSheetCsvLinkInCell_(sheet, 'D1', 'Download as CSV');
}

function setSheetCsvLinkInCell_(sheet, a1Notation, linkText) {
    const url = buildSheetCsvExportUrl_(SpreadsheetApp.getActiveSpreadsheet().getId(), sheet.getSheetId());
    const formula = `=HYPERLINK("${url}","${linkText}")`;
    sheet.getRange(a1Notation).setFormula(formula);
}

/**
 * Builds a pivot table for expense reporting from RegistroContableRange.
 */
function generateExpenseAccountPivot() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceRange = ss.getRangeByName('RegistroContableRange');
    const yearResumenRange = ss.getRangeByName('YearResumen');
    const divisaResumenRange = ss.getRangeByName('DivisaResumen');

    if (!sourceRange) {
        throw new Error('No se encontró el rango con nombre "RegistroContableRange".');
    }
    if (!yearResumenRange) {
        throw new Error('No se encontró el rango con nombre "YearResumen".');
    }
    if (!divisaResumenRange) {
        throw new Error('No se encontró el rango con nombre "DivisaResumen".');
    }

    const headers = sourceRange.offset(0, 0, 1, sourceRange.getNumColumns()).getValues()[0];
    const divisaResumen = String(divisaResumenRange.getDisplayValue() || '').trim();
    const yearResumen = String(yearResumenRange.getDisplayValue() || '').trim();

    if (!divisaResumen) {
        throw new Error('DivisaResumen está vacío.');
    }
    if (!yearResumen) {
        throw new Error('YearResumen está vacío.');
    }

    const categoriaCol = findColumnIndexByHeader_(headers, 'Categoria');
    const motivoCol = findColumnIndexByHeader_(headers, 'Motivo');
    const anoCol = findColumnIndexByHeader_(headers, 'ano', 'año');
    const mesCol = findColumnIndexByHeader_(headers, 'mes');
    findColumnIndexByHeader_(headers, 'Importe'); // Required field check
    const valueCol = findColumnIndexByHeader_(headers, divisaResumen);

    const targetSheetName = 'Cuenta de gastos';
    let targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) {
        targetSheet = ss.insertSheet(targetSheetName);
    }
    targetSheet.clear();

    targetSheet
        .getRange('A1')
        .setValue('Cuenta de gastos')
        .setFontWeight('bold')
        .setFontSize(16);

    const pivotAnchor = targetSheet.getRange('A3');
    const pivotTable = pivotAnchor.createPivotTable(sourceRange);

    const categoriaGroup = pivotTable.addRowGroup(categoriaCol);
    categoriaGroup.showTotals(true);
    categoriaGroup.sortAscending();

    const motivoGroup = pivotTable.addRowGroup(motivoCol);
    motivoGroup.showTotals(true);
    motivoGroup.sortAscending();

    const anoGroup = pivotTable.addColumnGroup(anoCol);
    anoGroup.showTotals(false);
    anoGroup.sortAscending();

    const mesGroup = pivotTable.addColumnGroup(mesCol);
    mesGroup.sortAscending();

    pivotTable.addPivotValue(valueCol, SpreadsheetApp.PivotTableSummarizeFunction.SUM);

    const yearCriteria = SpreadsheetApp.newFilterCriteria()
        .setVisibleValues([yearResumen])
        .build();
    pivotTable.addFilter(anoCol, yearCriteria);

    SpreadsheetApp.flush();
    formatPivotNumbers_(targetSheet, pivotAnchor.getRow(), pivotAnchor.getColumn());
    setSheetCsvLinkInCell_(targetSheet, 'C1', 'Download as CSV');
    protectSheetFromUserEdits_(targetSheet);
}

function findColumnIndexByHeader_(headers, primaryName) {
    const aliases = Array.prototype.slice.call(arguments, 1);
    const targetNames = [primaryName].concat(aliases).map(normalizeHeader_);

    for (let i = 0; i < headers.length; i++) {
        const normalizedHeader = normalizeHeader_(headers[i]);
        if (targetNames.indexOf(normalizedHeader) !== -1) {
            return i + 1; // 1-based column index for pivot APIs
        }
    }

    throw new Error(`No se encontró la columna "${primaryName}" en RegistroContableRange.`);
}

function normalizeHeader_(value) {
    return String(value || '')
        .normalize('NFD')
        .replace(new RegExp('[\\u0300-\\u036f]', 'g'), '')
        .trim()
        .toLowerCase();
}

function formatPivotNumbers_(sheet, startRow, startCol) {
    const rowCount = sheet.getMaxRows() - startRow + 1;
    const colCount = sheet.getMaxColumns() - startCol + 1;

    // Format a large area to avoid missing numeric cells as pivot dimensions change.
    sheet.getRange(startRow, startCol, rowCount, colCount).setNumberFormat('0');
}

/**
 * Keeps RegistroContableRange aligned with the actual RegistroContable table.
 */
function syncRegistroContableNamedRange_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('RegistroContable');
    if (!sheet) {
        return;
    }

    const anchor = sheet.getRange('A1');
    if (!anchor.getValue()) {
        return;
    }

    ss.setNamedRange('RegistroContableRange', anchor.getDataRegion());
}

function syncSaldosNamedRanges_() {
    syncNamedRangeToAnchorDataRegion_('SaldosRange', 'Saldos', 'A1');
    syncNamedRangeToAnchorDataRegion_('SaldosEntradaRange', 'Entrada de datos', 'A5');
}

function syncCuentasAndReglasNamedRanges_() {
    syncNamedRangeToAnchorDataRegion_('CuentasRange', 'Cuentas', 'A1');
    syncNamedRangeToAnchorDataRegion_('ReglasRange', 'Reglas de categorización', 'A1');
}

function syncNamedRangeToAnchorDataRegion_(rangeName, sheetName, anchorA1) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        return;
    }

    const anchor = sheet.getRange(anchorA1);
    if (!anchor.getValue()) {
        return;
    }

    ss.setNamedRange(rangeName, anchor.getDataRegion());
}

function protectSheetFromUserEdits_(sheet) {
    const description = 'AUTO_PROTECT_CUENTA_GASTOS';
    const protections = sheet
        .getProtections(SpreadsheetApp.ProtectionType.SHEET)
        .filter(p => p.getDescription() === description);

    const protection = protections.length > 0 ? protections[0] : sheet.protect();
    protection.setDescription(description);
    protection.setWarningOnly(true);
}

