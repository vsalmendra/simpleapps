/** @OnlyCurrentDoc */

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
                words: cleanText(row[0]).split(/\s+/).filter(w => w.length > 0),
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
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove diacritics
        .replace(/[0-9]/g, "")                           // Remove numbers
        .replace(/[^a-zA-Z]/g, " ")                      // Non-alphanumeric to spaces
        .toLowerCase()
        .trim();
}

/**
 * Checks if the array of ruleWords appears in the targetText in the correct order.
 */
function isMatchInOrder(targetText, ruleWords) {
    const targetWords = targetText.split(/\s+/).filter(w => w.length > 0);

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
    const srcValues = srcRange.getValues();
    const srcNumCols = srcRange.getNumColumns();

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

    const dataToCopy = srcValues.slice(0, rowsToCopy);

    const dstRangeOrig = ss.getRangeByName('RegistroContableRange');
    const dstRange = dstRangeOrig.offset(1, 0, dstRangeOrig.getNumRows() - 1, dstRangeOrig.getNumColumns());
    if (!dstRange) throw new Error('No se encontró el rango con nombre "RegistroContable".');

    const dstValues = dstRange.getValues();
    const dstNumCols = dstRange.getNumColumns();

    if(dstNumCols < srcNumCols) {
        throw new Error('Wrong number of columns');
    }

    // Find first completely blank row within the destination table
    let firstBlankRowIdx = findFirstBlankRowIdx(dstValues, srcNumCols);
    if (firstBlankRowIdx === -1) {
        throw new Error('No hay filas vacías en RegistroContable para pegar los datos.');
    }

    // Check space available
    const remainingRows = dstValues.length - firstBlankRowIdx;
    if (rowsToCopy > remainingRows) {
        throw new Error(`No hay suficiente espacio en "RegistroContable". Filas a copiar: ${rowsToCopy}, espacio disponible: ${remainingRows}.`);
    }

    // Write transaction values
    dstRange
        .offset(firstBlankRowIdx, 0, rowsToCopy, srcNumCols)
        .setValues(dataToCopy);

    // Write bank account on the next column
    const accountId = ss.getRangeByName('CuentaEntradaDatosID').getValue();
    dstRange
        .offset(firstBlankRowIdx, srcNumCols, rowsToCopy, 1)
        .setValue(accountId);

    // Now copy the balance rows
    const balanceEntryRangeOrig = ss.getRangeByName('SaldosEntradaRange');
    const balanceEntryRange =  balanceEntryRangeOrig.offset(1, 0, balanceEntryRangeOrig.getNumRows() - 1, balanceEntryRangeOrig.getNumColumns());

    const balanceValuesToCopy = findValuesToCopy(balanceEntryRange, 2);

    const balancesRangeOrig = ss.getRangeByName('SaldosRange');
    const balancesRange = balancesRangeOrig.offset(1, 0, balancesRangeOrig.getNumRows() - 1, balancesRangeOrig.getNumColumns());

    const balancesValues = balancesRange.getValues();

    const firstBlankBalanceRow = findFirstBlankRowIdx(balancesValues, 2);
    if (firstBlankBalanceRow === -1) {
        throw new Error('No hay filas vacías en Balances para pegar los datos.');
    }

    balancesRange
        .offset(firstBlankBalanceRow, 0, balanceValuesToCopy.length, 2)
        .setValues(balanceValuesToCopy);

    // Now set balances accountId
    balancesRange
        .offset(firstBlankBalanceRow, 2, balanceValuesToCopy.length, 1)
        .setValue(accountId);

    // Now clear source ranges
    balanceEntryRange.offset(0, 0, balanceValuesToCopy.length, 2).clearContent();
    srcRange.clearContent();

    ss.getRangeByName('CuentaEntradaDatosID').clearContent();
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
