/**
 * Class responsible for dynamically formatting, updating data validation,
 * adjusting bandings/borders, and managing filters across spreadsheet tables.
 * STRICT CANONICAL 100% STATIC EDITION: Zero object instantiation, zero silent errors.
 * 
 * Run `TableFormatter.help()` for full architectural specifications and API contracts.
 */
class TableFormatter {

  /**
   * Returns and logs the architectural specifications, execution contracts, and interface definitions.
   * @returns {string} Formatted architectural specification.
   */
  static help() {
    const doc = `
====================================================================================================
TABLE FORMATTER SPECIFICATION & ARCHITECTURAL CONTRACT (100% STATIC)
====================================================================================================

1. AUTOMATED TABLE GEOMETRY & FORMAT RECONCILIATION:
   Dynamic formatting and structural maintenance engine:
     - Strict Sheet Input: Expects a valid Sheet object directly (no string resolution or silent fallback).
     - Fail-Fast Execution: Zero try/catch blocks; all errors surface immediately without suppression.
     - Header & Boundary Detection: Automatically discovers table boundaries, headers, and last
       data rows using zero-allocation in-memory array scans.
     - Smart Early Exit: Skips write and formatting RPCs entirely (<0.02s) if grid dimensions
       and filter bounds are already synchronized.
     - Incremental Format & Rule Propagation: Clones cell styles, number formats, and data
       validation rules from the baseline row strictly across newly appended rows.
     - Banding & Border Reconciliation: Dynamically expands existing alternating row bandings
       or applies standardized grid borders across expanded dimensions.
     - Filter Range Synchronization: Directly resizes spreadsheet filter ranges to envelope
       newly added data rows without catching or ignoring errors.

2. UNIFIED STATIC EXECUTION INTERFACES:
   All operations are 100% static, stateless, and allocate zero instance wrappers.
   - TableFormatter.format(sheet):
     Formats, updates data validation, resizes bandings, and synchronizes filters for the sheet.
     Throws an error if sheet is missing or invalid.
   - TableFormatter.extend(sheet):
     Direct high-performance alias to TableFormatter.format(sheet).
   - TableFormatter.help():
     Returns formatted architectural specification.
====================================================================================================`.trim();

    Logger.log(doc);
    return doc;
  }

  /**
   * Refreshes formatting, data validation, and filters for the target sheet.
   * Strictly requires a valid Sheet object and throws immediately on invalid inputs or runtime issues.
   * 
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target Sheet instance.
   */
  static format(sheet) {
    if (!sheet || typeof sheet.getDataRange !== "function") {
      throw new Error("[TableFormatter.format] A valid GoogleAppsScript.Spreadsheet.Sheet object must be provided.");
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const numRows = values.length;
    if (numRows === 0) return;

    const numCols = values[0].length;
    if (numCols === 0) return;

    const dataStartRow = dataRange.getRow();
    const dataStartCol = dataRange.getColumn();
    const dataLastCol  = dataStartCol + numCols - 1;

    // 1. Dynamic Header Row & Last Data Row Detection (Fast Inlined Loops)
    let headerIndex = -1;
    for (let r = 0; r < numRows; r++) {
      const row = values[r];
      let hasData = false;
      for (let c = 0; c < numCols; c++) {
        const val = row[c];
        if (val !== "" && val !== null && val !== undefined) {
          hasData = true;
          break;
        }
      }
      if (hasData) {
        headerIndex = r;
        break;
      }
    }

    if (headerIndex === -1 || headerIndex >= numRows - 1) return;

    let lastIndex = -1;
    for (let r = numRows - 1; r >= headerIndex; r--) {
      const row = values[r];
      let hasData = false;
      for (let c = 0; c < numCols; c++) {
        const val = row[c];
        if (val !== "" && val !== null && val !== undefined) {
          hasData = true;
          break;
        }
      }
      if (hasData) {
        lastIndex = r;
        break;
      }
    }

    const headerRow = dataStartRow + headerIndex;
    const lastRow   = dataStartRow + lastIndex;
    const startRow  = headerRow + 1;

    if (lastRow <= headerRow) return;

    // 2. Filter Evaluation
    const filter = sheet.getFilter();
    let filterNeedsUpdate = false;
    let templateRow = startRow;

    if (filter) {
      const fRange = filter.getRange();
      const isExactMatch =
        fRange.getColumn() === dataStartCol &&
        fRange.getLastColumn() === dataLastCol &&
        fRange.getRow() === headerRow &&
        fRange.getLastRow() === lastRow;

      if (isExactMatch) {
        templateRow = lastRow;
      } else {
        filterNeedsUpdate = true;
        const fLastRow = fRange.getLastRow();
        templateRow = fLastRow >= startRow ? Math.min(fLastRow, lastRow) : startRow;
      }
    } else {
      filterNeedsUpdate = true;
      templateRow = Math.max(startRow, lastRow - 1);
    }

    // SMART EARLY EXIT: If row count didn't grow and filter is exact match, skip expensive writes
    if (lastRow <= templateRow && !filterNeedsUpdate) {
      return;
    }

    // 3. Incremental Formatting & Data Validation Copying (Only if new rows exist)
    if (lastRow > templateRow) {
      const numRowsToCopy = lastRow - templateRow;
      const sourceRange   = sheet.getRange(templateRow, dataStartCol, 1, numCols);
      const pasteTarget   = sourceRange.offset(1, 0, numRowsToCopy);

      sourceRange.copyTo(pasteTarget, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

      const rules = sourceRange.getDataValidations()[0];
      let hasValidation = false;
      for (let i = 0; i < rules.length; i++) {
        if (rules[i] !== null) {
          hasValidation = true;
          break;
        }
      }

      if (hasValidation) {
        sourceRange.copyTo(pasteTarget, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
      }
    }

    // 4. Handle Bandings and Borders (Only if dimensions changed)
    const fullTableRange = sheet.getRange(headerRow, dataStartCol, lastRow - headerRow + 1, numCols);

    const bandings = sheet.getBandings();
    let tableBanding = null;
    for (let i = 0; i < bandings.length; i++) {
      const bRange = bandings[i].getRange();
      if (
        bRange.getRow() <= lastRow &&
        bRange.getLastRow() >= headerRow &&
        bRange.getColumn() <= dataLastCol &&
        bRange.getLastColumn() >= dataStartCol
      ) {
        tableBanding = bandings[i];
        break;
      }
    }

    if (tableBanding) {
      if (tableBanding.getRange().getLastRow() !== lastRow) {
        tableBanding.setRange(fullTableRange);
      }
    } else {
      const bodyRange = sheet.getRange(startRow, dataStartCol, lastRow - startRow + 1, numCols);
      bodyRange.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    }

    // 5. Recreate Filter Directly (No Try/Catch - Fail-Fast)
    if (filterNeedsUpdate) {
      if (filter) filter.remove();
      fullTableRange.createFilter();
    }
  }

  /**
   * Static helper alias for format().
   * Strictly requires a valid Sheet object.
   * 
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target Sheet instance.
   */
  static extend(sheet) {
    TableFormatter.format(sheet);
  }
}
