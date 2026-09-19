/**
 * ============================================================================
 * SYSTEM TRIGGERS & COMMAND DISPATCH ENGINE (Optimized Enterprise Edition)
 * ============================================================================
 * Architecture: Batch-Processed Two-Cell Decoupled Command Dispatching
 * Integrations: Google Sheets UI + AppSheet REST API + Apps Script V8 Engine
 * ============================================================================
 */

// ============================================================================
// 1. GLOBAL SYSTEM CONFIGURATION
// ============================================================================
const SYSTEM_CONFIG = Object.freeze({
  SHEETS: {
    CONTROL: "System_Control",
    LOOKUP: "Lookup_Values"
  },
  LOOKUP: {
    HEADER_ROW: 2,
    CATEGORY_ID: 5, // System Automation Category
    COLS: {
      CATEGORY: "category_id",
      CODE: "item",         // Exact JS function name (Detail Code)
      LABEL: "description",  // Human-readable display label
      ACTIVE: "is_active",
      SEPARATOR: "extra_value_1"
    }
  },
  LOCK_TIMEOUT_MS: 10000,
  IDLE_VALUES: new Set(["", "- Select Action -", "IDLE", "Completed", "Not Running"])
});

// ============================================================================
// 2. COMMAND ENGINE (LOOKUP & MENU GENERATION)
// ============================================================================
const CommandEngine = {
  /**
   * Loads automation commands from Lookup_Values using BatchProcessor.read()
   * @returns {Array<{code: string, label: string, addSeparator: boolean}>}
   */
  loadCommands() {
    const table = BatchProcessor.read(SYSTEM_CONFIG.SHEETS.LOOKUP, SYSTEM_CONFIG.LOOKUP.HEADER_ROW);
    if (!table || table.length <= 1) return [];

    const headers = table[0].map(h => String(h ?? "").trim().toLowerCase());
    const catIdx  = headers.indexOf(SYSTEM_CONFIG.LOOKUP.COLS.CATEGORY);
    const codeIdx = headers.indexOf(SYSTEM_CONFIG.LOOKUP.COLS.CODE);
    const descIdx = headers.indexOf(SYSTEM_CONFIG.LOOKUP.COLS.LABEL);
    const actIdx  = headers.indexOf(SYSTEM_CONFIG.LOOKUP.COLS.ACTIVE);
    const sepIdx  = headers.indexOf(SYSTEM_CONFIG.LOOKUP.COLS.SEPARATOR);

    if (catIdx === -1 || codeIdx === -1 || descIdx === -1) {
      Logger.log(`[CommandEngine] Required columns missing in ${SYSTEM_CONFIG.SHEETS.LOOKUP}`);
      return [];
    }

    const commands = [];
    for (let i = 1; i < table.length; i++) {
      const row = table[i];
      if (Number(row[catIdx]) !== SYSTEM_CONFIG.LOOKUP.CATEGORY_ID) continue;

      const isAct = actIdx === -1 ? true : (row[actIdx] === true || String(row[actIdx]).toLowerCase() === "y" || row[actIdx] === 1);
      if (!isAct) continue;

      const code = String(row[codeIdx] ?? "").trim();
      const label = String(row[descIdx] ?? "").trim();

      if (code) {
        commands.push({
          code,
          label: label || code,
          addSeparator: sepIdx !== -1 && String(row[sepIdx] ?? "").toLowerCase().includes("separator")
        });
      }
    }
    return commands;
  },

  /**
   * Refreshes data validation dropdown on System_Control display cell
   */
  syncValidation(ss = SpreadsheetApp.getActiveSpreadsheet()) {
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.CONTROL);
    if (!sheet) return;

    const commands = this.loadCommands();
    const labels = ["- Select Action -", ...commands.map(c => c.label)];

    const cell = sheet.getRange("B2");
    cell.clearDataValidations();
    cell.setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInList(labels, true)
        .setAllowInvalid(true)
        .build()
    );

    const currentVal = String(cell.getValue() ?? "").trim();
    if (!labels.includes(currentVal)) {
      cell.setValue("- Select Action -");
    }
  }
};

// ============================================================================
// 3. CORE TRIGGER DISPATCHER (BATCH OPTIMIZED)
// ============================================================================
/**
 * Handles external state changes with a 1-Read / 2-Write batch pattern
 */
function handleExternalChange(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.CONTROL);
  if (!sheet) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(SYSTEM_CONFIG.LOCK_TIMEOUT_MS)) return;

  // Single bulk read of header row and data row (Row 1 & 2)
  const numCols = Math.max(sheet.getLastColumn(), 6);
  const buffer = sheet.getRange(1, 1, 2, numCols).getValues();
  const headers = buffer[0].map(h => String(h ?? "").trim().toLowerCase());
  const row = buffer[1];

  // Dynamic header resolution (falls back to standard B2:F2 positions if not found)
  const getCol = (key, fallback) => {
    const idx = headers.findIndex(h => h.includes(key));
    return idx !== -1 ? idx : fallback;
  };

  const idxDisplay = getCol("action", 1); // Col B (index 1)
  const idxCode    = getCol("code", 2);   // Col C (index 2)
  const idxStatus  = getCol("status", 3); // Col D (index 3)
  const idxDate    = getCol("update", 4); // Col E (index 4)
  const idxError   = getCol("error", 5);  // Col F (index 5)

  let actionCode = String(row[idxCode] ?? "").trim();

  // Fallback: If detail code is empty, resolve via Display Label
  if (SYSTEM_CONFIG.IDLE_VALUES.has(actionCode)) {
    const displayLabel = String(row[idxDisplay] ?? "").trim();
    if (SYSTEM_CONFIG.IDLE_VALUES.has(displayLabel)) {
      lock.releaseLock();
      return;
    }
    const match = CommandEngine.loadCommands().find(c => c.label.toLowerCase() === displayLabel.toLowerCase());
    if (!match) {
      lock.releaseLock();
      return;
    }
    actionCode = match.code;
    row[idxCode] = actionCode;
  }

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const getTimestamp = () => Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

  // WRITE 1: Set state to "Running" in a single batch write
  row[idxStatus] = "Running";
  row[idxDate] = getTimestamp();
  row[idxError] = "";
  sheet.getRange(2, 1, 1, numCols).setValues([row]);
  SpreadsheetApp.flush();

  let errorMsg = null;
  try {
    if (typeof globalThis[actionCode] !== "function") {
      throw new Error(`Automation error: Function "${actionCode}" is not defined.`);
    }
    // Execute target routine
    globalThis[actionCode]();
  } catch (err) {
    Logger.log(`[handleExternalChange] ${err.stack || err.message}`);
    errorMsg = err.message || String(err);
  } finally {
    // WRITE 2: Reset controls & record outcome in a single batch write
    row[idxDisplay] = "- Select Action -";
    row[idxCode]    = "IDLE";
    row[idxStatus]  = errorMsg ? "Error" : "Completed";
    row[idxDate]    = getTimestamp();
    row[idxError]   = errorMsg ?? "";

    sheet.getRange(2, 1, 1, numCols).setValues([row]);
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

// ============================================================================
// 4. TRIGGER ENTRYPOINTS
// ============================================================================

/**
 * Builds custom menu and aligns dropdown validation on open
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("🚀 System Control");

    CommandEngine.loadCommands().forEach(cmd => {
      if (typeof globalThis[cmd.code] === "function") {
        menu.addItem(cmd.label, cmd.code);
        if (cmd.addSeparator) menu.addSeparator();
      }
    });

    menu.addToUi();
    CommandEngine.syncValidation();
  } catch (err) {
    Logger.log(`[onOpen] ${err.stack || err.message}`);
  }
}

/**
 * Simple Trigger: Handles UI edits with an instant cell guard
 */
function onEdit(e) {
  if (typeof PivotSortManager !== "undefined" && typeof PivotSortManager.handleEdit === "function") {
    PivotSortManager.handleEdit(e);
  }

  // Pre-lock guard: Exit immediately if edit is not in System_Control Row 2 (Cols B or C)
  if (e && e.range) {
    if (
      e.range.getSheet().getName() !== SYSTEM_CONFIG.SHEETS.CONTROL ||
      e.range.getRow() !== 2 ||
      (e.range.getColumn() !== 2 && e.range.getColumn() !== 3)
    ) {
      return;
    }
  }

  handleExternalChange(e);
}

// ============================================================================
// 5. CONCISE PUBLIC EXECUTION ENDPOINTS
// ============================================================================

function optimizeAll() {
  new PortfolioManager().optimizeAll();
}

function resetAndRecalculateAll() {
  new PortfolioManager().resetAndRecalculateAll();
}

function sortTransactions() {
  new PortfolioManager().sortTransactions();
}

function sortGoogleFinance() {
  new PortfolioManager().sortGoogleFinance();
}

function listAllFormulas() {
  if (typeof WorkbookAuditor !== "undefined" && typeof WorkbookAuditor.listAllFormulas === "function") {
    WorkbookAuditor.listAllFormulas();
  }
}

function listAllTables() {
  if (typeof WorkbookAuditor !== "undefined" && typeof WorkbookAuditor.listAllTables === "function") {
    WorkbookAuditor.listAllTables();
  }
}

function setupPivotSortControls() {
  if (typeof PivotSortManager !== "undefined" && typeof PivotSortManager.setupControls === "function") {
    PivotSortManager.setupControls();
  }
}

function extendFormat(sheet) {
  if (typeof TableFormatter !== "undefined" && typeof TableFormatter.extend === "function") {
    TableFormatter.extend(sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
  }
}

/**
 * Wrapper for WorkbookAuditor formula recovery routine
 */
function restoreAllFormulas() {
  if (typeof WorkbookAuditor !== "undefined" && typeof WorkbookAuditor.restoreAllFormulas === "function") {
    WorkbookAuditor.restoreAllFormulas();
  }
}

/**
 * Wrapper for ArrayFormula repair routine
 */
function fixAllArrayFormulaErrors(target = null) {
  if (typeof WorkbookAuditor !== "undefined" && typeof WorkbookAuditor.fixAllArrayFormulaErrors === "function") {
    WorkbookAuditor.fixAllArrayFormulaErrors(target);
  } else if (typeof FormulaAudit !== "undefined" && typeof FormulaAudit.fixAllErrors === "function") {
    FormulaAudit.fixAllErrors(target);
  }
}