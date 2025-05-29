/**
 * Applies earthwork calculation formulas to an Excel worksheet.
 * 
 * This includes:
 * - Averaging existing levels (Cols D to G → Col M)
 * - Averaging proposed levels (Cols H to K → Col N)
 * - Calculating cut volumes (Col O)
 * - Calculating fill volumes (Col P)
 * 
 * Uses LET and IFS constructs for efficient Excel formula logic.
 *
 * @param workbook - The active ExcelScript workbook object
 */
function applyEarthworkCalculationFormulas(workbook: ExcelScript.Workbook): void {

    const sheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();

    const dataStartRowNum = 14;

    // Determine the range of rows based on the Cell ID column
    const {startCellRowNum, endCellRowNum, rowCount } = getRowRangeFromStartCell(workbook, `C${dataStartRowNum}`);

    // Formula templates for each target column
    const averageExistingLevelFormulaTemplate = (row: number) => `=IF(SUM(D${row}:G${row})=0, "", AVERAGE(D${row}:G${row}))`;
    const averageProposedLevelFormulaTemplate = (row: number) => `=IF(SUM(H${row}:K${row})=0, "", AVERAGE(H${row}:K${row}))`;

    const cutVolumeFormulaTemplate = (row: number) => `
    =LET(
    existing, M${row},
    proposed, N${row},
    cellSizeHori, $J$8,
    cellSizeVert, $J$9,
    coverage, L${row},

    area, cellSizeHori * cellSizeVert,

    final, IFS(
        AND(existing = "", proposed = ""), "",
        existing > proposed, (existing - proposed) * area * coverage,
        TRUE, "-"
    ),

    final
)`.trim();

    const fillVolumeFormulaTemplate = (row: number) => `
    =LET(
    existing, M${row},
    proposed, N${row},
    cellSizeHori, $J$8,
    cellSizeVert, $J$9,
    coverage, L${row},

    area, cellSizeHori * cellSizeVert,

    final, IFS(
        AND(existing = "", proposed = ""), "",
        proposed > existing, (proposed - existing) * area * coverage,
        TRUE, "-"
    ),

    final
)`.trim();

    // Insert each formula into its target column
    insertFormulaColumn(sheet, averageExistingLevelFormulaTemplate, "M", startCellRowNum, endCellRowNum);
    insertFormulaColumn(sheet, averageProposedLevelFormulaTemplate, "N", startCellRowNum, endCellRowNum);
    insertFormulaColumn(sheet, cutVolumeFormulaTemplate, "O", startCellRowNum, endCellRowNum);
    insertFormulaColumn(sheet, fillVolumeFormulaTemplate, "P", startCellRowNum, endCellRowNum);

    console.log("Formula(s) insertion for earthwork calculation completed.")
}
