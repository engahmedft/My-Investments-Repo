/**
 * Core engine responsible for mapping table schemas, indexing workbook formulas,
 * repairing ArrayFormula expansion collisions, and executing atomic formula restores.
 * STRICT CANONICAL 100% STATIC EDITION: Zero object instantiation, zero aliases, zero legacy fallbacks.
 * 
 * Run `WorkbookRegistry.help()` for full architectural specifications and API contracts.
 */
class WorkbookRegistry {

  /**
   * Prevent instantiation of purely static utility class.
   */
  constructor() {
    throw new Error("WorkbookRegistry is a static utility class and cannot be instantiated. Call methods directly via WorkbookRegistry.<methodName>().");
  }

  /**
   * Returns and logs the architectural specifications, execution contracts, and interface definitions.
   * @returns {string} Formatted architectural specification.
   */
  static help() {
    const doc = `
====================================================================================================
WORKBOOK REGISTRY ARCHITECTURAL SPECIFICATION & OPERATIONAL CONTRACT (100% STATIC)
====================================================================================================

1. CORE SYSTEM REGISTRY & DIAGNOSTIC RESPONSIBILITIES:
   Central static architecture and metadata engine:
     - Component 1: Rapid dynamic array and ArrayFormula repair (zero-write collision unblocking).
     - Component 2: Centralized formula inventory and harvesting ('Formula_Registry').
     - Component 3: Contiguous data table island detection ('Table_Registry' via 1D typed BFS).
     - Component 4: High-fidelity formula recovery and atomic sheet restoration.

2. UNIFIED STATIC EXECUTION INTERFACES:
   All operations are 100% static, stateless, and allocate zero instance wrappers.
   - WorkbookRegistry.fixAllArrayFormulaErrors(target, ss):
     Scans top rows for dynamic arrays (ARRAYFORMULA, MAP, BYROW, QUERY, etc.) or #REF! spill errors.
     Identifies residual static cells blocking expansion in memory and clears them in a single
     RangeList API call per sheet.
   - WorkbookRegistry.listAllFormulas(reportName = "Formula_Registry", ss):
     Extracts all live formulas across non-system sheets and persists a consolidated inventory to
     'Formula_Registry' with timestamps and A1 cell coordinates.
   - WorkbookRegistry.listAllTables(reportName = "Table_Registry", ss):
     Scans active grid matrices to detect independent data islands via O(N) 8-direction BFS
     and writes consolidated table schemas to 'Table_Registry'.
   - WorkbookRegistry.restoreAllFormulas(registrySheetName = "Formula_Registry", ss):
     Reapplies archived formulas from 'Formula_Registry' back to their target cell coordinates
     without per-cell RPC flushes.
   - WorkbookRegistry.runAllRegistries(ss):
     Master runner executing both listAllTables and listAllFormulas in sequence.

3. HIGH-PERFORMANCE ZERO-WRITE PRINCIPLES:
   - 1-API Read per sheet via getDataRange().
   - In-memory V8 evaluation for dirty checking and conflict detection.
   - 1-API Write via RangeList.clearContent() only if blocked formulas actually exist.
   - Eliminates redundant RPCs and per-cell server flushes.
====================================================================================================`.trim();

    Logger.log(doc);
    return doc;
  }

  /**
   * Pre-compiled Regex for dynamic array formula detection.
   * @private
   */
  static get _DYNAMIC_ARRAY_REGEX() {
    if (!WorkbookRegistry._DYNAMIC_ARRAY_REGEX_CACHE) {
      WorkbookRegistry._DYNAMIC_ARRAY_REGEX_CACHE = /^=(?:ARRAYFORMULA\b|MAP\b|BYROW\b|BYCOL\b|SCAN\b|REDUCE\b|QUERY\b|FILTER\b|SORT\b|UNIQUE\b|\{)/i;
    }
    return WorkbookRegistry._DYNAMIC_ARRAY_REGEX_CACHE;
  }

  /**
   * Pre-compiled Regex for R1C1 formula notation detection.
   * @private
   */
  static get _R1C1_REGEX() {
    if (!WorkbookRegistry._R1C1_REGEX_CACHE) {
      WorkbookRegistry._R1C1_REGEX_CACHE = /(?:R\[?-?\d*\]?C\[?-?\d*\]?|RC\[?-?\d*\]?)/i;
    }
    return WorkbookRegistry._R1C1_REGEX_CACHE;
  }

  // 8-Direction BFS Grid Offsets
  static get _DR() { return [-1, -1, -1,  0,  0,  1,  1,  1]; }
  static get _DC() { return [-1,  0,  1, -1,  1, -1,  0,  1]; }

  // Blacklist of internal registry, control, and lookup sheets to ignore
  static get SYSTEM_SHEET_KEYWORDS() {
    if (!WorkbookRegistry._SYSTEM_KEYWORDS_CACHE) {
      WorkbookRegistry._SYSTEM_KEYWORDS_CACHE = new Set([
        "tableregistry",
        "formularegistry",
        "optimizelog",
        "systemcontrol",
        "lookupvalues",
        "lookupcategories",
        "perusersettings",
        "detailslookups",
        "masterlookups"
      ]);
    }
    return WorkbookRegistry._SYSTEM_KEYWORDS_CACHE;
  }

  /**
   * Converts a 1-based column index to an A1-notation letter string.
   * @param {number} col - 1-based column index.
   * @returns {string} Column letter (e.g., "A", "Z", "AA").
   */
  static getColLetter(col) {
    let letter = "";
    let tempCol = col;
    while (tempCol > 0) {
      const temp = (tempCol - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      tempCol = Math.floor((tempCol - temp - 1) / 26);
    }
    return letter;
  }

  /**
   * Normalizes and checks whether a sheet name belongs to internal system sheets.
   * @private
   */
  static _isSystemSheet(sheetName) {
    const normalized = String(sheetName || "").replace(/[\s_-]/g, "").toLowerCase();
    return WorkbookRegistry.SYSTEM_SHEET_KEYWORDS.has(normalized);
  }

  /**
   * HIGH-PERFORMANCE ZERO-WRITE ARRAYFORMULA FIXER
   * 
   * 1. 1-READ (getDataRange): Reads formulas and values in bulk.
   * 2. ZERO-WRITE PASS: Skips write APIs completely if no formulas are blocked.
   * 3. BATCH RANGE CLEARING: Clears all blocked columns in 1 single API write per sheet via RangeList.
   * 
   * @param {string|GoogleAppsScript.Spreadsheet.Sheet} [target=null] - Target sheet or null for all sheets.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()]
   * @returns {{ totalColumnsFixed: number, details: string[] }}
   */
  static fixAllArrayFormulaErrors(target = null, ss = SpreadsheetApp.getActiveSpreadsheet()) {
    let sheetsToProcess = [];

    if (target) {
      const sheet = typeof target === "string" ? ss.getSheetByName(target) : target;
      if (!sheet?.getDataRange) {
        throw new Error(`[WorkbookRegistry.fixAllArrayFormulaErrors] Invalid sheet target: "${target}" was not found.`);
      }
      sheetsToProcess = [sheet];
    } else {
      sheetsToProcess = ss.getSheets().filter(s => !WorkbookRegistry._isSystemSheet(s.getName()));
    }

    let totalColumnsFixed = 0;
    const fixedDetails = [];
    const arrayRegex = WorkbookRegistry._DYNAMIC_ARRAY_REGEX;

    for (let s = 0; s < sheetsToProcess.length; s++) {
      const sheet = sheetsToProcess[s];
      const dataRange = sheet.getDataRange();
      const numRows = dataRange.getNumRows();
      const numCols = dataRange.getNumColumns();

      if (numRows < 2 || numCols < 1) continue;

      const formulas = dataRange.getFormulas();
      const values = dataRange.getValues();
      const maxRows = sheet.getMaxRows();
      const sheetName = sheet.getName();

      const scanDepth = Math.min(numRows, 3);
      const rangesToClear = [];

      for (let c = 0; c < numCols; c++) {
        for (let r = 0; r < scanDepth; r++) {
          const formula = String(formulas[r][c] || "").trim();
          const value = String(values[r][c] || "").trim();

          const isArrayFormula = (formula && arrayRegex.test(formula)) || value === "#REF!";

          if (isArrayFormula && (r + 1 < numRows)) {
            let hasBlockingData = false;
            for (let subR = r + 1; subR < numRows; subR++) {
              const cellVal = values[subR][c];
              if (cellVal !== "" && cellVal !== null && cellVal !== undefined) {
                hasBlockingData = true;
                break;
              }
            }

            if (hasBlockingData) {
              const colLetter = WorkbookRegistry.getColLetter(c + 1);
              const startRow = r + 2;
              rangesToClear.push(`${colLetter}${startRow}:${colLetter}${maxRows}`);
              totalColumnsFixed++;
              break;
            }
          }
        }
      }

      if (rangesToClear.length > 0) {
        sheet.getRangeList(rangesToClear).clearContent();
        fixedDetails.push(`${sheetName} (${rangesToClear.length} col)`);
      }
    }

    if (totalColumnsFixed > 0) {
      SpreadsheetApp.flush();
    }

    const resultMsg = totalColumnsFixed > 0
      ? `✅ Fixed ${totalColumnsFixed} blocked ArrayFormula column(s) in: ${fixedDetails.join(", ")}`
      : `✅ All ArrayFormula columns are clean. No action needed.`;

    Logger.log(`[WorkbookRegistry.fixAllArrayFormulaErrors] ${resultMsg}`);
    ss.toast(resultMsg, "ArrayFormula Optimizer", 5);

    return { totalColumnsFixed, details: fixedDetails };
  }

  /**
   * Scans all sheets for formulas and writes a consolidated report to 'Formula_Registry'.
   * @param {string} [reportName="Formula_Registry"]
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()]
   */
  static listAllFormulas(reportName = "Formula_Registry", ss = SpreadsheetApp.getActiveSpreadsheet()) {
    const sheets = ss.getSheets();
    const nowTimestamp = new Date();
    const allDataToPush = [["Sheet Name", "Cell", "Formula", "Last Update"]];

    for (let s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      const sheetName = sheet.getName();
      if (WorkbookRegistry._isSystemSheet(sheetName)) continue;

      const formulas = sheet.getDataRange().getFormulas();
      const numRows = formulas.length;
      if (numRows === 0) continue;

      const numCols = formulas[0].length;

      for (let r = 0; r < numRows; r++) {
        const row = formulas[r];
        const rowRef = r + 1;

        for (let c = 0; c < numCols; c++) {
          const formula = row[c];
          if (formula) {
            allDataToPush.push([
              sheetName,
              WorkbookRegistry.getColLetter(c + 1) + rowRef,
              `'${formula}`,
              nowTimestamp
            ]);
          }
        }
      }
    }

    WorkbookRegistry._writeReport(
      ss,
      reportName,
      allDataToPush,
      "No formulas found in the entire workbook.",
      "Formula registry updated!"
    );
  }

  /**
   * Detects contiguous data "islands" (tables) across all sheets using 1D BFS in O(N) time.
   * Outputs report strictly to 'Table_Registry'.
   * @param {string} [reportName="Table_Registry"]
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()]
   */
  static listAllTables(reportName = "Table_Registry", ss = SpreadsheetApp.getActiveSpreadsheet()) {
    const sheets = ss.getSheets();
    const nowTimestamp = new Date();
    const results = [["Sheet Name", "Header Row", "Column", "Header Name", "Last Update"]];

    let visited = new Uint8Array(1024);
    let queue = new Int32Array(1024);
    const dr = WorkbookRegistry._DR;
    const dc = WorkbookRegistry._DC;

    for (let s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      const sheetName = sheet.getName();
      if (WorkbookRegistry._isSystemSheet(sheetName)) continue;

      const fullDataRange = sheet.getDataRange();
      const numRows = fullDataRange.getNumRows();
      const numCols = fullDataRange.getNumColumns();
      const maxCells = numRows * numCols;

      if (maxCells === 0) continue;

      if (visited.length < maxCells) {
        visited = new Uint8Array(maxCells);
        queue = new Int32Array(maxCells);
      } else {
        visited.fill(0, 0, maxCells);
      }

      const values = fullDataRange.getValues();

      for (let r = 0; r < numRows; r++) {
        const rowValues = values[r];
        let rowHasData = false;
        for (let c = 0; c < numCols; c++) {
          const v = rowValues[c];
          if (v !== "" && v !== null && v !== undefined) {
            rowHasData = true;
            break;
          }
        }
        if (!rowHasData) continue;

        for (let c = 0; c < numCols; c++) {
          const idx = r * numCols + c;

          if (values[r][c] !== "" && visited[idx] === 0) {
            let minR = r, maxR = r, minC = c, maxC = c;
            let head = 0, tail = 0;

            visited[idx] = 1;
            queue[tail++] = idx;

            while (head < tail) {
              const currIdx = queue[head++];
              const cr = (currIdx / numCols) | 0;
              const cc = currIdx % numCols;

              for (let i = 0; i < 8; i++) {
                const nr = cr + dr[i];
                const nc = cc + dc[i];

                if (nr >= 0 && nr < numRows && nc >= 0 && nc < numCols) {
                  const nIdx = nr * numCols + nc;

                  if (visited[nIdx] === 0 && values[nr][nc] !== "") {
                    visited[nIdx] = 1;
                    queue[tail++] = nIdx;

                    if (nr < minR) minR = nr;
                    else if (nr > maxR) maxR = nr;

                    if (nc < minC) minC = nc;
                    else if (nc > maxC) maxC = nc;
                  }
                }
              }
            }

            for (let i = minC; i <= maxC; i++) {
              results.push([
                sheetName,
                minR + 1,
                WorkbookRegistry.getColLetter(i + 1),
                values[minR][i],
                nowTimestamp
              ]);
            }
          }
        }
      }
    }

    WorkbookRegistry._writeReport(
      ss,
      reportName,
      results,
      "No data tables found.",
      "Table registry updated!"
    );
  }

  /**
   * Master runner executing both table and formula registries in sequence.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()]
   */
  static runAllRegistries(ss = SpreadsheetApp.getActiveSpreadsheet()) {
    WorkbookRegistry.listAllTables("Table_Registry", ss);
    WorkbookRegistry.listAllFormulas("Formula_Registry", ss);
  }

  /**
   * Sanitizes formula string by stripping apostrophe prefixes, cleaning escapes, and standardizing '='.
   * @private
   */
  static _sanitizeFormula(rawFormula) {
    if (!rawFormula) return "";
    let formula = String(rawFormula).trim();

    if (formula.startsWith("'")) {
      formula = formula.substring(1).trim();
    }
    formula = formula.replace(/\\"/g, '"');
    formula = formula.replace(/^=+/, '');
    return "=" + formula.trim();
  }

  /**
   * Fast formula application without per-cell flushes.
   * @private
   */
  static _applyFormulaToRange(range, cleanFormula) {
    const isR1C1 = WorkbookRegistry._R1C1_REGEX.test(cleanFormula);
    if (isR1C1) {
      try {
        range.setFormulaR1C1(cleanFormula);
        return true;
      } catch (e) {
        range.setFormula(cleanFormula);
        return true;
      }
    } else {
      try {
        range.setFormula(cleanFormula);
        return true;
      } catch (e) {
        range.setFormulaR1C1(cleanFormula);
        return true;
      }
    }
  }

  /**
   * Restores formulas recorded in 'Formula_Registry' back to their target cells.
   * @param {string} [registrySheetName="Formula_Registry"]
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()]
   * @returns {{ success: boolean, total: number, applied: number, skipped: number, errors: string[] }}
   */
  static restoreAllFormulas(registrySheetName = "Formula_Registry", ss = SpreadsheetApp.getActiveSpreadsheet()) {
    const registrySheet = ss.getSheetByName(registrySheetName);
    if (!registrySheet) {
      const msg = `Could not find '${registrySheetName}' sheet.`;
      ss.toast(msg, "Error ❌", 5);
      return { success: false, total: 0, applied: 0, skipped: 0, errors: [msg] };
    }

    const data = registrySheet.getDataRange().getValues();
    if (data.length <= 1) {
      ss.toast("No records found in formula registry sheet.", "Finished ℹ️", 4);
      return { success: true, total: 0, applied: 0, skipped: 0, errors: [] };
    }

    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const colSheetName = headers.indexOf("sheet name");
    const colCell = headers.indexOf("cell");
    const colFormula = headers.indexOf("formula");

    if (colSheetName === -1 || colCell === -1 || colFormula === -1) {
      const msg = `Invalid header structure in ${registrySheetName}. Expected 'Sheet Name', 'Cell', and 'Formula'.`;
      ss.toast(msg, "Error ❌", 5);
      return { success: false, total: 0, applied: 0, skipped: 0, errors: [msg] };
    }

    let applied = 0;
    let skipped = 0;
    const errors = [];
    const sheetCache = new Map();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const targetSheetName = String(row[colSheetName] || "").trim();
      const targetCell = String(row[colCell] || "").trim();
      const rawFormula = String(row[colFormula] || "").trim();

      if (!targetSheetName || !targetCell || !rawFormula) {
        skipped++;
        continue;
      }

      const cleanFormula = WorkbookRegistry._sanitizeFormula(rawFormula);

      try {
        if (!sheetCache.has(targetSheetName)) {
          sheetCache.set(targetSheetName, ss.getSheetByName(targetSheetName));
        }
        const targetSheet = sheetCache.get(targetSheetName);

        if (!targetSheet) {
          throw new Error(`Sheet '${targetSheetName}' not found.`);
        }

        const range = targetSheet.getRange(targetCell);
        WorkbookRegistry._applyFormulaToRange(range, cleanFormula);
        applied++;

      } catch (err) {
        const errorMsg = `Row ${i + 1} (${targetSheetName}!${targetCell}): ${err.message}`;
        errors.push(errorMsg);
        Logger.log(`[WorkbookRegistry.restoreAllFormulas Error] ${errorMsg}`);
      }
    }

    // Single flush commit after setting all formulas
    SpreadsheetApp.flush();

    const summaryMsg = `Applied: ${applied} | Skipped: ${skipped} | Errors: ${errors.length}`;
    ss.toast(summaryMsg, errors.length === 0 ? "Success ✅" : "Warning ⚠️", 6);
    Logger.log(`[WorkbookRegistry.restoreAllFormulas] Summary -> ${summaryMsg}`);

    return {
      success: errors.length === 0,
      total: data.length - 1,
      applied,
      skipped,
      errors
    };
  }

  /**
   * Helper writing consolidated registry outputs to the target report sheet.
   * @private
   */
  static _writeReport(ss, reportName, data, emptyMsg, successMsg) {
    const reportSheet = ss.getSheetByName(reportName) || ss.insertSheet(reportName);
    reportSheet.clear();

    const dataLen = data.length;
    if (dataLen > 1) {
      const numCols = data[0].length;
      reportSheet.getRange(1, 1, dataLen, numCols).setValues(data);
      reportSheet.getRange(1, 1, 1, numCols).setFontWeight("bold").setBackground("#3c78d8").setFontColor("#ffffff");
      reportSheet.getRange(2, numCols, dataLen - 1, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
      reportSheet.setFrozenRows(1);
      reportSheet.autoResizeColumns(1, numCols);

      ss.toast(`${successMsg} Found ${dataLen - 1} entries in ${reportName}.`, "Success ✅", 5);
    } else {
      ss.toast(emptyMsg, "Finished ℹ️", 3);
    }
  }
}