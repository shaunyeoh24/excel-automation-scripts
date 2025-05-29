/**
 * Validates that the specified cells contain non-empty positive integers.
 * 
 * @param sheet - The Excel worksheet to validate.
 * @param cells - An array of cell addresses to check (e.g., ["E8", "E9"]).
 * @returns An array of error messages for invalid or empty cells.
 */
function validatePositiveIntegerCells(sheet: ExcelScript.Worksheet, cells: string[]): string[] {
    return validateCellsCore(
        sheet,
        cells,
        (v) => typeof v === "number" && Number.isInteger(v) && v > 0,
        (cell) => `${cell} must be a positive integer.`
    );
}
