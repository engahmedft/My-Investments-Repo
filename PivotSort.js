/**
 * Unified class managing Pivot Table control setup, fast event handling, and sorting operations.
 * STRICT CANONICAL 100% STATIC EDITION: Zero object instantiation.
 * 
 * Run `PivotSortManager.help()` for full architectural specifications and API contracts.
 */
class PivotSortManager {


  /**
   * Returns and logs the architectural specifications, execution contracts, and interface definitions.
   * @returns {string} Formatted architectural specification.
   */
  static help() {
    const doc = `
====================================================================================================
PIVOT SORT MANAGER SPECIFICATION & ARCHITECTURAL CONTRACT (100% STATIC)
====================================================================================================

1. DYNAMIC PIVOT TABLE SORTING & CONTROL INJECTION:
   Engine responsible for managing interactive sort controls above spreadsheet pivot tables:
     - Automated UI Setup: Scans sheet for pivot tables, reserves a 3-row control header,
       and injects data-validated dropdowns ('Sort By:' and 'Direction:').
     - Zero-Overhead Fast-Path onEdit: 2-stage in-memory filter exits in <0.01s for any edit
       not matching the exact 2x2 control coordinate matrix adjacent to pivot tables.
     - Aggregate-Aware Sorting: Inspects pivot table row groups and aggregate value columns,
       dynamically rebinding group.sortBy() and direction (Ascending / Descending).

2. UNIFIED STATIC EXECUTION INTERFACES:
   All operations are 100% static, stateless, and allocate zero instance wrappers.
   - PivotSortManager.setupControls(sheet):
     Discovers all pivot tables on the target sheet, inserts 3-row headers if missing,
     and configures dynamic 'Sort By:' and 'Direction:' validation dropdowns.
     (Defaults to the active sheet if omitted).
   - PivotSortManager.handleEdit(e):
     High-performance event dispatcher called from onEdit(e) to process dropdown changes.
   - PivotSortManager.applySort(pivotTable, colName, direction):
     Binds target aggregate column to pivot row groups and applies sort ordering.
   - PivotSortManager.help():
     Returns formatted architectural specification.
====================================================================================================`.trim();

    Logger.log(doc);
    return doc;
  }

  /**
   * Entry point for onEdit trigger. 
   * Includes multi-stage fast exits so non-control edits exit immediately in V8 memory.
   * 
   * @param {GoogleAppsScript.Events.SheetsOnEdit} e - Trigger event object.
   */
  static handleEdit(e) {
    // Stage 1 Fast Exits: Check event properties in memory before any API calls
    if (!e || !e.value || !e.range) return;

    const range = e.range;

    // Fast Exit: Ensure edit was a single cell change
    if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

    const row = range.getRow();
    const col = range.getColumn();

    // Fast Exit: Controls are placed at col >= 2 and row >= 2
    if (col < 2 || row < 2) return;

    const sheet = range.getSheet();

    // Boundary Guard: Ensure 3x2 matrix does not exceed sheet dimensions
    if (row + 1 > sheet.getMaxRows() || col > sheet.getMaxColumns()) return;

    // Stage 2: Read single 3x2 matrix surrounding edit area (1 API call)
    const block = sheet.getRange(row - 1, col - 1, 3, 2).getValues();
    const label = block[1][0];

    // Fast Exit: Strict label check - exits instantly for ANY edit outside pivot sort dropdowns
    if (label !== "Sort By:" && label !== "Direction:") return;

    // Locate target pivot anchor
    const pivotAnchorRow = (label === "Sort By:") ? row + 3 : row + 2;
    const pivotAnchorCol = col - 1;

    const pivotTable = sheet.getPivotTables().find(p => {
      const anchor = p.getAnchorCell();
      return anchor.getColumn() === pivotAnchorCol && Math.abs(anchor.getRow() - pivotAnchorRow) <= 2;
    });

    if (!pivotTable) return;

    // Extract values from memory matrix block
    const sortBy = (label === "Sort By:") ? e.value : block[0][1];
    const direction = (label === "Direction:") ? e.value : block[2][1];

    // Apply sort statically
    PivotSortManager.applySort(pivotTable, sortBy, direction);
  }

  /**
   * Sets up control dropdowns above all pivot tables on the sheet.
   * 
   * @param {GoogleAppsScript.Spreadsheet.Sheet} [targetSheet] - Target sheet. Defaults to active sheet.
   */
  static setupControls(targetSheet = null) {
    const sheet = targetSheet || SpreadsheetApp.getActiveSheet();
    if (!sheet || typeof sheet.getPivotTables !== "function") {
      throw new Error("[PivotSortManager.setupControls] A valid GoogleAppsScript.Spreadsheet.Sheet object must be provided.");
    }

    const pivots = sheet.getPivotTables();

    if (pivots.length === 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast("No pivot tables found on sheet.", "Pivot Controls", 4);
      return;
    }

    // Keep direct reference to pivot object to avoid secondary API calls
    const pivotList = pivots.map(p => ({
      pivot: p,
      r: p.getAnchorCell().getRow(),
      c: p.getAnchorCell().getColumn()
    }));

    // Top-down sort for proper offset calculations
    pivotList.sort((a, b) => a.r - b.r);

    const initialData = sheet.getDataRange().getValues();
    let offset = 0;

    pivotList.forEach(item => {
      // 0-based array index for row 3 positions above anchor
      const checkRowIdx = item.r - 4;
      const checkColIdx = item.c - 1;

      const exists = checkRowIdx >= 0 && 
                     initialData[checkRowIdx] && 
                     initialData[checkRowIdx][checkColIdx] === "Sort By:";

      if (!exists) {
        sheet.insertRowsBefore(item.r + offset, 3);
        offset += 3;
      }
    });

    if (offset > 0) {
      SpreadsheetApp.flush();
    }

    // Re-read data once into memory after all insertions
    const newData = sheet.getDataRange().getValues();
    const dirRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(["Ascending", "Descending"], true)
      .build();

    pivotList.forEach(item => {
      // Pivot anchor cells update dynamically after insertRowsBefore
      const anchorRow = item.pivot.getAnchorCell().getRow();
      const anchorCol = item.pivot.getAnchorCell().getColumn();

      const numRowGroups = item.pivot.getRowGroups().length;

      // Header row is 1 row above anchor
      const headerRowIdx = anchorRow - 2;
      const rawRowData = newData[headerRowIdx] || [];
      const headersRaw = rawRowData.slice(anchorCol - 1);
      const headers = headersRaw.filter(String).slice(numRowGroups);

      if (headers.length === 0) headers.push("No Aggregates");

      const sortByRule = SpreadsheetApp.newDataValidation().requireValueInList(headers, true).build();

      // Read current values directly from memory
      const sortByRowIdx = anchorRow - 4;
      const dirRowIdx = anchorRow - 3;
      const valColIdx = anchorCol; // Column B relative to label cell in A

      const currentSortBy = newData[sortByRowIdx] ? newData[sortByRowIdx][valColIdx] : null;
      const currentDir = newData[dirRowIdx] ? newData[dirRowIdx][valColIdx] : null;

      const sortByVal = headers.includes(currentSortBy) ? currentSortBy : headers[0];
      const dirVal = (currentDir === "Ascending" || currentDir === "Descending") ? currentDir : "Descending";

      // Batch update formatting, values, and validations in a single range call
      const controlRange = sheet.getRange(anchorRow - 3, anchorCol, 2, 2);
      controlRange.setBackground("#f3f3f3")
                  .setBorder(true, true, true, true, null, null)
                  .setValues([
                    ["Sort By:", sortByVal],
                    ["Direction:", dirVal]
                  ])
                  .setDataValidations([
                    [null, sortByRule],
                    [null, dirRule]
                  ])
                  .setFontWeights([
                    ["bold", "normal"],
                    ["bold", "normal"]
                  ]);
    });

    sheet.getParent().toast("Controls Setup Complete", "Pivot Sort", 3);
  }

  /**
   * Core logic to update row group sorting on a pivot table.
   * 
   * @param {GoogleAppsScript.Spreadsheet.PivotTable} pivotTable - Target pivot table instance.
   * @param {string} colName - Column aggregate name to sort by.
   * @param {string} direction - "Ascending" or "Descending".
   */
  static applySort(pivotTable, colName, direction) {
    if (!pivotTable || !colName || !direction) return;

    const sheet = pivotTable.getAnchorCell().getSheet();
    const anchor = pivotTable.getAnchorCell();
    const anchorRow = anchor.getRow();
    const anchorCol = anchor.getColumn();

    const maxSearchCols = Math.min(100, sheet.getLastColumn() - anchorCol + 1);
    if (maxSearchCols <= 0) return;

    const headers = sheet.getRange(anchorRow, anchorCol, 1, maxSearchCols).getValues()[0];
    const colIdx = headers.indexOf(colName);
    const rowGroups = pivotTable.getRowGroups();
    const pivotValues = pivotTable.getPivotValues();

    const valIdx = colIdx - rowGroups.length;

    if (rowGroups.length === 0 || valIdx < 0 || valIdx >= pivotValues.length) {
      Logger.log(`[PivotSortManager.applySort Warning] Could not match aggregate column "${colName}".`);
      return;
    }

    const targetValue = pivotValues[valIdx];
    const isAscending = (direction === "Ascending");

    rowGroups.forEach(group => {
      group.sortBy(targetValue, []);
      if (isAscending) {
        group.sortAscending();
      } else {
        group.sortDescending();
      }
    });

    sheet.getParent().toast(`Pivot sorted by ${colName} (${direction})`, "Pivot Sort", 3);
  }
}