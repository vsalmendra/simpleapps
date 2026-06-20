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
        .addToUi();

    refreshResumenCsvLinkInCell();
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

    // Remove header from data
    const dataBody = dataValues.slice(1);

    // 2. Prepare Rules (Clean them once for efficiency)
    const preparedRules = rulesValues
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

