/**
 * @function findLastRow_
 * @file https://www.reddit.com/r/GoogleAppsScript/comments/10n4n1r/how_to_find_the_last_row_in_a_column_with_google/
 * @author u/IAmMoonie <https://www.reddit.com/user/IAmMoonie/>
 * @version 1.0
 * @param {string} column - The column letter to find the last row of. Must be a single upper-case letter.
 * @param {Object} sheet - The sheet object to find the last row in.
 * @returns {number} - The last row index of the provided column.
 * @throws {Error} - If an invalid column letter is provided, or if the provided column does not exist in the sheet or has no cells with data.
 * @description - Find the last row of a given column in a Google Sheets sheet.
 * @example
 * const sheet = SpreadsheetApp.getActiveSheet();
 * const lastRow = findLastRow_("A", sheet);
 * console.log(lastRow);
 * // Output: 5 (if the last row with data in column A is the 5th row)
 */
function findLastRow_(column, sheet) {
    try {
      if (!column || !column.match(/^[A-Z]+$/)) {
        throw new Error(
          "Invalid column letter provided. Please provide a valid column letter (e.g. A, B, etc.)."
        );
      }
      const values = sheet
        .getRange(`${column}1:${column}${sheet.getLastRow()}`)
        .getValues()
        .flat();
      const lastRowIndex = values
        .map((val, index) => (val !== "" ? index + 1 : ""))
        .filter((x) => x !== "")
        .pop();
      if (lastRowIndex === undefined) {
        throw new Error(
          `Provided column ${column} does not exist in the sheet or has no cells with data.`
        );
      }
      return lastRowIndex;
    } catch (e) {
      console.error(e);
    }
}
