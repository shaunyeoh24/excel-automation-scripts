/**
 * Converts a zero-based column index to Excel-style letters (e.g. 0 -> A)
*/

function convertIndexToColumnLabel(index: number): string {

  if (!Number.isInteger(index) || index < 0) {
    throw new Error("Number is not a valid non-negative integer");
  }

  let output = "";

  while (index >= 0) {
    let remainder = index % 26;
    output = String.fromCharCode(65 + remainder) + output;
    index = Math.floor(index / 26) - 1;
  }

  console.log(output);
  return output;
}
