/**
 * Generic cell validator that checks emptiness, then checks a custom validation function to each specified cell.
 * 
 * @param sheet - The Excel worksheet to validate.
 * @param cells - Array of cell addresses to validate.
 * @param validateFn - Function that returns true if the cell value is valid.
 * @param errorMsg - Function that returns an error message string for a given cell address.
 * @returns Array of error messages for cells that fail validation.
 */
function validateCellsCore(
    sheet: ExcelScript.Worksheet,
    cells: string[],
    validateFn: (value) => boolean,
    errorMsg: (cell: string) => string
) {

    // Initialize array to store errors
    const errors: string[] = [];

    // Process validation for each cell in range
    for (const cell of cells) {
        const value = sheet.getRange(cell).getValue();

        if (value === "" || value === null) {
            errors.push(`${cell} is empty.`)
        } else if (!validateFn(value)) {
            errors.push(errorMsg(cell))
        }
    }

    // Report validation status in terminal
    if (errors.length === 0) {
        console.log("Validation successful - no errors found...");
    } else {
        throw new Error("Validation error(s):\n" + errors.join("\n"));
    }
}
