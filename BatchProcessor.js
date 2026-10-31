/**
 * Fully self-contained, high-performance batch data processor for Google Sheets.
 * Operates on 1 API Read + 1 API Write + (1 optional API Delete / Insert for structural mutations).
 * Run `BatchProcessor.help()` for full architectural specifications and API contracts.
 */
class BatchProcessor {

  /**
   * Returns and logs the architectural specifications, execution contracts, and interface definitions.
   * @returns {string} Formatted architectural specification.
   */
  static help() {
    const doc = `
====================================================================================================
BATCHPROCESSOR ARCHITECTURAL SPECIFICATION & EXECUTION CONTRACT
====================================================================================================

1. 1-READ / 1-WRITE BATCH LATENCY ELIMINATION:
   Eliminates spreadsheet RPC overhead by consolidating grid access into a single bulk read
   (getDataRange().getValues()), running all row transformations in V8 heap memory, and 
   committing updates in at most one bulk write (Range.setValues()).

2. IMMUTABLE REFERENCE ISOLATION (.slice()):
   Supplies processFn with an isolated shallow copy (rawData[index].slice()), preserving the
   original values in rawData. This prevents in-place mutation side effects and guarantees 
   deterministic dirty-checking against the baseline sheet state.

3. ZERO-WRITE PASS & DIRTY CHECKING:
   Compares transformed rows against baseline values using type-aware equality (handling Date
   timestamps and NaN). If dimensions and non-formula cells are unchanged, all write, insert,
   delete, and filter RPCs are short-circuited entirely (<0.1s execution).

4. DEFERRED SMART FILTER PROTECTION (_clearFilterCriteria):
   Inspects and clears active filter criteria ONLY when write mutations actually occur.
   Eliminates redundant RPC queries during no-op executions while preventing "operation not
   supported on a filtered row" write exceptions when committing data.

5. ARRAYFORMULA SHIELDING:
   Masks registered formula columns with null prior to write-back. Google Sheets ignores null
   entries during setValues(), preventing formula overwrites and #SPILL! / #REF! expansion errors.

6. BOUNDARY-SAFE GRID RECONCILIATION & AUTO-PADDING:
   - Expands sheet grid boundaries (insertRowsAfter) if processed row counts exceed sheet capacity.
   - Truncates surplus trailing rows (deleteRows) when output shrinks, guarded against grid minimums.
   - Enforces uniform 2D rectangularity: returned rows shorter than numCols are padded with '""',
     while rows exceeding numCols are cleanly truncated.

7. STRICT UNIFORM INTERFACE CONTRACT:
   Target Specification:
     sheetTarget: string (sheet name) | GoogleAppsScript.Spreadsheet.Sheet

   Callback Signature:
     (row: any[], col: (name: string) => number, index: number, rawData: any[][]) => any[][]
     * col(name): Case-insensitive, whitespace-trimmed column index lookup (returns 0-based idx or -1).
     * Cardinality:
       - [row]                -> Standard 1:1 row transformation.
       - [row, ...newRows]    -> 1:N expansion (dynamic row insertion).
       - []                   -> 1:0 reduction (row omission/filtering).

   Methods:
     - BatchProcessor.read(sheetTarget, headerRowNumber = 1): 
         { sheet: Sheet, headers: string[], rawData: any[][], numCols: number, headerRow: number }
     - BatchProcessor.process(sheetTarget, headerRowNumber, processFn, arrayFormulaHeaders = []): void
     - BatchProcessor.help(): string
====================================================================================================`.trim();

    Logger.log(doc);
    return doc;
  }

  /**
   * Normalizes header strings by converting nullish values to empty strings, trimming whitespace, and lowercasing.
   * @private
   * @param {any} header - Raw header value.
   * @returns {string} Normalized lowercase string.
   */
  static _normalizeHeader(header) {
    return String(header ?? "").trim().toLowerCase();
  }

  /**
   * Clears filter criteria ONLY from columns that currently have active criteria hiding rows.
   * Preserves the filter range and dropdown UI buttons while unhiding all rows.
   * @private
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Target Sheet instance.
   */
  static _clearFilterCriteria(sheet) {
    if (!sheet) return;
    const filter = sheet.getFilter();
    if (!filter) return;

    const range = filter.getRange();
    const startCol = range.getColumn();
    const endCol = startCol + range.getNumColumns() - 1;

    for (let col = startCol; col <= endCol; col++) {
      if (filter.getColumnFilterCriteria(col) !== null) {
        filter.removeColumnFilterCriteria(col);
      }
    }
  }

  /**
   * Resolves a sheet target (string name or Sheet object) into a validated Sheet instance.
   * @private
   * @param {string|GoogleAppsScript.Spreadsheet.Sheet} sheetTarget - Sheet name or Sheet instance.
   * @returns {GoogleAppsScript.Spreadsheet.Sheet} Validated Sheet instance.
   * @throws {Error} If the target sheet cannot be found or is invalid.
   */
  static _resolveSheet(sheetTarget) {
    const sheet = typeof sheetTarget === "string"
      ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetTarget)
      : sheetTarget;

    if (!sheet || typeof sheet.getDataRange !== "function") {
      throw new Error(`Invalid sheet target: "${sheetTarget}" was not found.`);
    }
    return sheet;
  }

  /**
   * Low-level read engine: resolves the sheet, reads grid data, and extracts normalized headers and raw rows.
   * Minimizes V8 object allocations by directly reusing sub-arrays where column dimensions match.
   * @private
   * @param {string|GoogleAppsScript.Spreadsheet.Sheet} sheetTarget - Target sheet name or instance.
   * @param {number} [headerRowNumber=1] - 1-based row index containing column headers.
   * @returns {{sheet: GoogleAppsScript.Spreadsheet.Sheet, headers: string[], rawData: any[][], numCols: number, headerRow: number}}
   */
  static read(sheetTarget, headerRowNumber = 1) {
    const sheet = BatchProcessor._resolveSheet(sheetTarget);
    const headerRow = Math.max(1, Math.floor(Number(headerRowNumber) || 1));

    // 1. Single API read across active grid
    const allValues = sheet.getDataRange().getValues();
    if (allValues.length < headerRow) {
      return { sheet, headers: [], rawData: [], numCols: 0, headerRow };
    }

    // 2. Scan header row in reverse to locate rightmost populated column
    const headerRowIndex = headerRow - 1;
    const rawHeaderRow = allValues[headerRowIndex];

    let lastHeaderColIndex = -1;
    for (let c = rawHeaderRow.length - 1; c >= 0; c--) {
      const h = rawHeaderRow[c];
      if (h !== "" && h !== null && h !== undefined && String(h).trim() !== "") {
        lastHeaderColIndex = c;
        break;
      }
    }

    if (lastHeaderColIndex === -1) {
      return { sheet, headers: [], rawData: [], numCols: 0, headerRow };
    }

    const numCols = lastHeaderColIndex + 1;
    const normalizedHeaders = new Array(numCols);
    for (let c = 0; c < numCols; c++) {
      normalizedHeaders[c] = BatchProcessor._normalizeHeader(rawHeaderRow[c]);
    }

    // 3. Scan upwards from bottom to find the last populated data row
    let lastDataRowIndex = -1;
    for (let r = allValues.length - 1; r >= headerRow; r--) {
      const row = allValues[r];
      let hasData = false;

      for (let c = 0; c < numCols; c++) {
        const val = row[c];
        if (val !== "" && val !== null && val !== undefined) {
          hasData = true;
          break;
        }
      }

      if (hasData) {
        lastDataRowIndex = r;
        break;
      }
    }

    if (lastDataRowIndex < headerRow) {
      return { sheet, headers: normalizedHeaders, rawData: [], numCols, headerRow };
    }

    // 4. Extract data rows; reuse source row reference if column width matches numCols
    const dataRowCount = lastDataRowIndex - headerRow + 1;
    const rawData = new Array(dataRowCount);
    for (let i = 0, r = headerRow; r <= lastDataRowIndex; r++, i++) {
      const row = allValues[r];
      rawData[i] = row.length === numCols ? row : row.slice(0, numCols);
    }

    return { sheet, headers: normalizedHeaders, rawData, numCols, headerRow };
  }

  /**
   * Batch process & write: transforms data in-memory and persists mutations to the sheet.
   * Operates strictly on a 1-Read / 1-Write model with smart filter protection and zero-write pass.
   * 
   * @param {string|GoogleAppsScript.Spreadsheet.Sheet} sheetTarget - Sheet name string or Sheet object.
   * @param {number} headerRowNumber - 1-based row number where headers reside.
   * @param {Function} processFn - Strict transform callback: (row: any[], col: (name: string) => number, index: number, rawData: any[][]) => Array<Array<any>>
   * @param {string[]} [arrayFormulaHeaders=[]] - Header names containing ARRAYFORMULAs to be protected.
   */
  static process(sheetTarget, headerRowNumber, processFn, arrayFormulaHeaders = []) {
    // 1. Single API read & sheet resolution
    const { sheet: targetSheet, headers: normalizedHeaders, rawData, numCols, headerRow } =
      BatchProcessor.read(sheetTarget, headerRowNumber);

    if (!rawData || rawData.length === 0) {
      return; // No data rows to process
    }

    const oldNumRows = rawData.length;

    // 2. Build unified header map & memoized column index resolver
    const headerMap = new Map();
    for (let c = 0; c < numCols; c++) {
      const h = normalizedHeaders[c];
      if (h !== "" && !headerMap.has(h)) {
        headerMap.set(h, c);
      }
    }

    const col = (headerName) => {
      let idx = headerMap.get(headerName);
      if (idx !== undefined) return idx;

      const normalized = BatchProcessor._normalizeHeader(headerName);
      idx = headerMap.has(normalized) ? headerMap.get(normalized) : -1;
      headerMap.set(headerName, idx);
      return idx;
    };

    // 3. O(1) typed array lookup for ArrayFormula protected columns
    const isFormulaCol = new Uint8Array(numCols);
    const formulaColIndices = [];
    for (let i = 0; i < arrayFormulaHeaders.length; i++) {
      const idx = col(arrayFormulaHeaders[i]);
      if (idx !== -1) {
        formulaColIndices.push(idx);
        isFormulaCol[idx] = 1;
      }
    }
    const numFormulaCols = formulaColIndices.length;

    // 4. In-memory data transformation loop
    const processedData = [];
    for (let index = 0; index < oldNumRows; index++) {
      // Shallow copy (.slice()) preserves pristine original row in rawData for Zero-Write comparison
      const rowInput = rawData[index].slice();

      // Single-form invocation: strictly passes (row, col, index, rawData)
      // and strictly expects a 2D array return: Array<Array<any>>
      const res = processFn(rowInput, col, index, rawData);
      if (!Array.isArray(res) || res.length === 0) continue;

      const resLen = res.length;
      for (let r = 0; r < resLen; r++) {
        const targetRow = res[r];
        if (!Array.isArray(targetRow)) continue;

        const curLen = targetRow.length;
        if (curLen < numCols) {
          for (let c = curLen; c < numCols; c++) targetRow.push("");
        } else if (curLen > numCols) {
          targetRow.length = numCols;
        }

        // Nullify ArrayFormula columns to prevent write collisions
        for (let f = 0; f < numFormulaCols; f++) {
          targetRow[formulaColIndices[f]] = null;
        }
        processedData.push(targetRow);
      }
    }

    const startRow = headerRow + 1;
    const newNumRows = processedData.length;

    // 5. Zero-Write pass: compare processed data against pristine original raw sheet data
    let dataHasChanged = (newNumRows !== oldNumRows);
    if (!dataHasChanged) {
      for (let r = 0; r < newNumRows; r++) {
        const pRow = processedData[r];
        const oRow = rawData[r];
        for (let c = 0; c < numCols; c++) {
          if (isFormulaCol[c]) continue; // Skip ArrayFormula columns
          const pVal = pRow[c];
          const oVal = oRow[c];
          if (pVal !== oVal) {
            // Guard Date references and NaN comparisons against false-positive writes
            if (pVal instanceof Date && oVal instanceof Date) {
              if (pVal.getTime() === oVal.getTime()) continue;
            } else if (Number.isNaN(pVal) && Number.isNaN(oVal)) {
              continue;
            }
            dataHasChanged = true;
            break;
          }
        }
        if (dataHasChanged) break;
      }
    }

    // 6. Write back only if changes were detected
    if (dataHasChanged) {
      // SMART FILTER PROTECTION: Clear active filter criteria only when batch mutations actually occur
      BatchProcessor._clearFilterCriteria(targetSheet);

      let currentMaxRows = targetSheet.getMaxRows();

      if (newNumRows > 0) {
        const requiredMaxRow = startRow + newNumRows - 1;
        if (requiredMaxRow > currentMaxRows) {
          const rowsToAdd = requiredMaxRow - currentMaxRows;
          targetSheet.insertRowsAfter(currentMaxRows, rowsToAdd);
          currentMaxRows = requiredMaxRow; // Update cached row count locally
        }

        // Single bulk API write pass
        targetSheet.getRange(startRow, 1, newNumRows, numCols).setValues(processedData);
      }

      // Delete extra trailing rows if output row count contracted
      if (newNumRows < oldNumRows) {
        const deleteStartRow = startRow + newNumRows;
        const rowsToDelete = oldNumRows - newNumRows;

        if (deleteStartRow <= currentMaxRows) {
          const safeRowsToDelete = Math.min(rowsToDelete, currentMaxRows - deleteStartRow + 1);
          if (safeRowsToDelete > 0) {
            targetSheet.deleteRows(deleteStartRow, safeRowsToDelete);
          }
        }
      }
    }
  }
}
