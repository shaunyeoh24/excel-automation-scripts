/**
 * Inserts a column of formulas into a specified range, using a row-based formula template.
 *
 * Applies the formula dynamically for each row based on start and end row indexes.
 *
 * @param sheet             The ExcelScript worksheet where formulas will be inserted.
 * @param startRowIndex     Row index to start inserting formulas (0-based).
 * @param endRowIndex       Row index to stop inserting formulas (0-based).
 * @param targetColumn      Column letter where formulas will be inserted (e.g., "AA").
 * @param formulaTemplate   Function that generates a formula string based on a row number.
 */

function insertFormulaColumn(
    sheet: ExcelScript.Worksheet,
    formulaTemplate: (row: number) => string,
    targetColumn: string,
    startRowNum: number,
    endRowNum: number
): void {
    const formulaRange = sheet.getRange(`${targetColumn}${startRowNum}:${targetColumn}${endRowNum}`);
    const formulas: string[][] = [];

    for (let row = startRowNum; row <= endRowNum; row++) {
        formulas.push([formulaTemplate(row)]);
    }

    console.log(`Formulas inserted: From ${targetColumn}${startRowNum} to ${targetColumn}${endRowNum}`);
    formulaRange.setFormulas(formulas);
}
