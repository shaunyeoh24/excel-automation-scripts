/**
 * Applies formula-based validation checks to a worksheet to verify the continuity of topographic data
 * across adjacent grid cells based on their Cell IDs.
 * 
 * Validations performed:
 * - Existing level vertical (Up-Down) continuity → Col S
 * - Proposed level vertical (Up-Down) continuity → Col T
 * - Existing level horizontal (Left-Right) continuity → Col U
 * - Proposed level horizontal (Left-Right) continuity → Col V
 * 
 * Each validation compares relevant edge cell values between a current cell and its adjacent cell,
 * identified using Cell IDs in Column C (e.g., "A01", "B02", etc.).
 * 
 * @param workbook - The ExcelScript workbook containing the worksheet to validate.
 */
function applyTopographyContinuityValidationFormulas(workbook: ExcelScript.Workbook): void {

  const sheet: ExcelScript.Worksheet = workbook.getActiveWorksheet();

  const dataStartRowNum = 14;

  // Determine the range of rows based on the Cell ID column
  const { startCellRowNum, endCellRowNum, rowCount } = getRowRangeFromStartCell(workbook, `C${dataStartRowNum}`);

  // Validation: Existing Topography Data – Vertical Continuity
  const existingUpDownValidation = (row: number) => `
  =LET(
    currentID, C${row},
    current_Ex3, F${row},
    current_Ex4, G${row},

    colLetter, LEFT(currentID, 1),
    rowNumber, VALUE(MID(currentID, 2, LEN(currentID)-1)),
    downID, colLetter & TEXT(rowNumber + 1, "00"),

    matchRow, XMATCH(downID, C:C, 0),

    below_Ex1, INDEX(D:D, matchRow),
    below_Ex2, INDEX(E:E, matchRow),

    isMatch, AND(current_Ex3 = below_Ex1, current_Ex4 = below_Ex2),
    result, IF(ISNUMBER(matchRow), isMatch, "-"),

    result
)`.trim();

  // Validation: Proposed Topography Data – Vertical Continuity
  const proposedUpDownValidation = (row: number) => `
  =LET(
    currentID, C${row},
    current_Pr3, J${row},
    current_Pr4, K${row},

    colLetter, LEFT(currentID, 1),
    rowNumber, VALUE(MID(currentID, 2, LEN(currentID)-1)),
    downID, colLetter & TEXT(rowNumber + 1, "00"),

    matchRow, XMATCH(downID, C:C, 0),

    below_Pr1, INDEX(H:H, matchRow),
    below_Pr2, INDEX(I:I, matchRow),

    isMatch, AND(current_Pr3 = below_Pr1, current_Pr4 = below_Pr2),
    result, IF(ISNUMBER(matchRow), isMatch, "-"),

    result
)`.trim();

  // Validation: Existing Topography Data – Horizontal Continuity
  const existingLeftRightValidation = (row: number) => `
  =LET(
    currentID, C${row},
    current_Ex1, D${row},
    current_Ex3, F${row},

    colLetter, LEFT(currentID, 1),
    rowNumber, VALUE(MID(currentID, 2, LEN(currentID)-1)),
    leftColLetter, CHAR(CODE(colLetter) - 1),
    leftID, leftColLetter & TEXT(rowNumber, "00"),

    matchRow, XMATCH(leftID, C:C, 0),

    left_Ex2, INDEX(E:E, matchRow),
    left_Ex4, INDEX(G:G, matchRow),

    isMatch, AND(current_Ex1 = left_Ex2, current_Ex3 = left_Ex4),
    result, IF(ISNUMBER(matchRow), isMatch, "-"),

    result
)`.trim();

  // Validation: Proposed Topography data – Horizontal Continuity
  const proposedLeftRightValidation = (row: number) => `
  =LET(
    currentID, C${row},
    current_Pr1, H${row},
    current_Pr3, J${row},

    colLetter, LEFT(currentID, 1),
    rowNumber, VALUE(MID(currentID, 2, LEN(currentID)-1)),
    leftColLetter, CHAR(CODE(colLetter) - 1),
    leftID, leftColLetter & TEXT(rowNumber, "00"),

    matchRow, XMATCH(leftID, C:C, 0),

    left_Pr2, INDEX(I:I, matchRow),
    left_Pr4, INDEX(K:K, matchRow),

    isMatch, AND(current_Pr1 = left_Pr2, current_Pr3 = left_Pr4),
    result, IF(ISNUMBER(matchRow), isMatch, "-"),

    result
)`.trim();

  // Insert each formula into its target column
  insertFormulaColumn(sheet, existingUpDownValidation, "S", startCellRowNum, endCellRowNum);
  insertFormulaColumn(sheet, proposedUpDownValidation, "T", startCellRowNum, endCellRowNum);
  insertFormulaColumn(sheet, existingLeftRightValidation, "U", startCellRowNum, endCellRowNum);
  insertFormulaColumn(sheet, proposedLeftRightValidation, "V", startCellRowNum, endCellRowNum);

  console.log("Formula insertion for topographic continuity validation completed.")
}
