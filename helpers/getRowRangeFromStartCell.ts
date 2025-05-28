/**
 * Returns the number of rows from a start cell to the last non-empty cell in the same column, within an optional search limit.
 *
 * Assumes data is vertically aligned. Allows gaps between cells.
 *
 * @param workbook          The ExcelScript workbook
 * @param startCell         Optional cell address to start from (default: active cell)
 * @param rowSearchLimit    Optional limit of rows to search (default: 1000)
 * @returns                 Object with rowCount, startCellRowIndex and lastCellRowIndex
 */

function getRowRangeFromStartCell(
    workbook: ExcelScript.Workbook,
    startCell?: string,
    rowSearchLimit?: number
): { startCellRowNum: number; endCellRowNum: number; rowCount: number } {

    let sheet = workbook.getActiveWorksheet();
    let startCellAddress: string = startCell ?? workbook.getActiveCell().getAddress().split("!")[1];
    let startCellRange = sheet.getRange(startCellAddress);
    let startCellRowIndex = startCellRange.getRowIndex();
    let startCellRowNum = startCellRowIndex + 1;

    // Check that start cell is not empty
    let startValue = startCellRange.getValue();
    if (startValue === "" || startValue === null) throw new Error("Start cell empty.")

    // Define how far down to search (defaults to 1000 rows)
    let searchLimit = rowSearchLimit ?? 1000;
    let searchRange = startCellRange.getResizedRange(searchLimit - 1, 0);

    // Identify the portion of the search range that contains data
    let usedRange = searchRange.getUsedRange();

    // Check that search range is not empty
    if (!usedRange) throw new Error("No data found.");

    // Calculate the position of the last populated cell
    let rowCount = usedRange.getRowCount();
    let endCell = sheet.getCell(startCellRowIndex + rowCount - 1, startCellRange.getColumnIndex());
    let endCellRowIndex = endCell.getRowIndex();
    let endCellRowNum = endCellRowIndex + 1;
    let endCellAddress = endCell.getAddressLocal().split("!")[1];

    console.log(`Data row count complete: From ${startCellAddress} to ${endCellAddress}, found ${rowCount} row(s).`);

    return { startCellRowNum, endCellRowNum, rowCount };
}
