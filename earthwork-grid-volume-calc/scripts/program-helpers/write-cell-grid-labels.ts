/**
 * Generates and prints a grid of Excel-style cell labels (e.g., "A1", "B3", ...) 
 * into the sheet, based on specified horizontal and vertical resolution.
 * 
 * @param horizontalRes - Number of columns in the grid (must be a positive integer)
 * @param verticalRes - Number of rows in the grid (must be a positive integer)
 * @param startCellAddress - Top cell where the one-column grid will be printed (e.g., "C14")
 * @param sheet - ExcelScript.Worksheet where the data will be written
 * @returns An object containing:
 *   - labels: array of generated cell labels in column-major order
 *   - columnFlags: array of corresponding column letters for each label
 *   - rowCount: total number of grid cells
 */

function writeCellGridLabels(
    horizontalRes: number,
    verticalRes: number,
    startCellAddress: string,
    sheet: ExcelScript.Worksheet,
): { labels: string[]; columnFlags: string[]; rowCount: number} {

    // Validate input resolutions: must be positive integers
    if (!Number.isInteger(horizontalRes) || horizontalRes <= 0) {
        throw new Error("Horizontal resolution must be a positive integer.");
    }
    if (!Number.isInteger(verticalRes) || verticalRes <= 0) {
        throw new Error("Vertical resolution must be a positive integer.");
    }

    const labels: string[] = [];
    const columnFlags: string[] = [];
    const rowCount: number = horizontalRes * verticalRes;

    // Generate labels in column-major order (top-to-bottom, then next column)
    for (let col = 0; col < horizontalRes; col++) {
        const columnLetter = convertIndexToColumnLabel(col);
        for (let row = 1; row <= verticalRes; row++) {
            labels.push(`${columnLetter}${row}`);
            columnFlags.push(columnLetter);
        }
    }

     // Determine the top cell's range and write vertically downwards
    const labelData = labels.map(label => [label]); // Convert to 2D array (vertical)
    const labelRange = sheet.getRange(startCellAddress).getResizedRange(labelData.length - 1, 0);
    labelRange.setValues(labelData);

    return { labels, columnFlags, rowCount };
}
