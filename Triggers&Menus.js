/**
 * ============================================================================
 * TRIGGER DISPATCHER & COMMAND ENGINE
 * Strictly matching canonical table schemas
 * ============================================================================
 */
const SYSTEM_CONFIG = Object.freeze({
  SHEETS: {
    CONTROL: "System_Control",
    LOOKUP: "Lookup_Values"
  },
  LOOKUP: {
    HEADER_ROW: 2,
    CATEGORY_ID: 5,
    COLS: {
      CATEGORY: "Category_ID",
      CODE: "Item",
      LABEL: "Description",
      ACTIVE: "Is_Active",
      SEPARATOR: "Extra_Value_1"
    }
  },
  LOCK_TIMEOUT_MS: 10000,
  IDLE_VALUES: new Set(["", "- Select Action -", "IDLE", "Completed", "Not Running"])
});

const CommandEngine = {
  /**
   * Loads active system commands strictly via BatchProcessor.read without range fallback.
   * Filters exclusively for Category_ID = 5 and Is_Active = 'Y'.
   * @returns {Array<{code: string, label: string, addSeparator: boolean}>}
   */
  loadCommands() {
    const readResult = BatchProcessor.read(SYSTEM_CONFIG.SHEETS.LOOKUP, SYSTEM_CONFIG.LOOKUP.HEADER_ROW);
    if (!readResult) return [];

    const headers = (readResult.headers || []).map(h => String(h ?? "").trim().toLowerCase());
    const rows = readResult.rawData || [];

    if (!headers.length || !rows.length) return [];

    const getIdx = (colName) => headers.indexOf(String(colName).trim().toLowerCase());

    const catIdx  = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.CATEGORY);
    const codeIdx = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.CODE);
    const descIdx = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.LABEL);
    const actIdx  = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.ACTIVE);
    const sepIdx  = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.SEPARATOR);

    // Verify all essential columns exist
    if (catIdx === -1 || codeIdx === -1 || descIdx === -1 || actIdx === -1) {
      Logger.log(`[CommandEngine] Required columns missing in ${SYSTEM_CONFIG.SHEETS.LOOKUP}`);
      return [];
    }

    const commands = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];

      // Filter by Category ID (Category 5 = System Actions)
      if (Number(row[catIdx]) !== SYSTEM_CONFIG.LOOKUP.CATEGORY_ID) continue;

      // Filter: Strictly Is_Active === 'Y'
      const isAct = String(row[actIdx] ?? "").trim().toUpperCase() === "Y";
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
   * Synchronizes data validation dropdown on System_Control sheet.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss]
   */
  syncValidation(ss = SpreadsheetApp.getActiveSpreadsheet()) {
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.CONTROL);
    if (!sheet) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const actionColIdx = headers.indexOf("Action");
    if (actionColIdx === -1) return;

    const commands = this.loadCommands();
    const labels = ["- Select Action -", ...commands.map(c => c.label)];

    const cell = sheet.getRange(2, actionColIdx + 1);
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

/**
 * Handles automation dispatch triggered by System_Control dropdown or cell edit.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} [e]
 */
function handleExternalChange(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.CONTROL);
  if (!sheet) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(SYSTEM_CONFIG.LOCK_TIMEOUT_MS)) return;

  const numCols = sheet.getLastColumn();
  const buffer = sheet.getRange(1, 1, 2, numCols).getValues();
  const headers = buffer[0].map(h => String(h ?? "").trim());
  const row = buffer[1];

  const idxAction = headers.indexOf("Action");
  const idxCode   = headers.indexOf("Action Code");
  const idxState  = headers.indexOf("State");
  const idxDate   = headers.indexOf("Last update");
  const idxError  = headers.indexOf("Error");

  if (idxAction === -1 || idxState === -1 || idxDate === -1 || idxError === -1) {
    lock.releaseLock();
    throw new Error(`[System_Control] Missing canonical columns in Row 1.`);
  }

  const displayLabel = String(row[idxAction] ?? "").trim();
  let actionCode = "";

  // Prioritize active selection in dropdown first
  if (!SYSTEM_CONFIG.IDLE_VALUES.has(displayLabel)) {
    const match = CommandEngine.loadCommands().find(c => c.label.toLowerCase() === displayLabel.toLowerCase());
    if (match) {
      actionCode = match.code;
    }
  }

  // Fallback to Action Code cell if Action dropdown is idle or manual code was passed
  if (!actionCode && idxCode !== -1 && !SYSTEM_CONFIG.IDLE_VALUES.has(String(row[idxCode] ?? "").trim())) {
    actionCode = String(row[idxCode]).trim();
  }

  if (!actionCode) {
    lock.releaseLock();
    return;
  }

  if (idxCode !== -1) row[idxCode] = actionCode;

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const getTimestamp = () => Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

  row[idxState] = "Running";
  row[idxDate]  = getTimestamp();
  row[idxError] = "";
  sheet.getRange(2, 1, 1, numCols).setValues([row]);
  SpreadsheetApp.flush();

  let errorMsg = null;
  try {
    if (typeof globalThis[actionCode] !== "function") {
      throw new Error(`Automation error: Function "${actionCode}" is not defined.`);
    }
    globalThis[actionCode]();
  } catch (err) {
    Logger.log(`[handleExternalChange] ${err.stack || err.message}`);
    errorMsg = err.message || String(err);
  } finally {
    row[idxAction] = "- Select Action -";
    if (idxCode !== -1) row[idxCode] = "IDLE";
    row[idxState]  = errorMsg ? "Error" : "Completed";
    row[idxDate]   = getTimestamp();
    row[idxError]  = errorMsg ?? "";

    sheet.getRange(2, 1, 1, numCols).setValues([row]);
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

/**
 * Standard onOpen trigger initializing the UI menu and data validation.
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
 * Standard onEdit trigger capturing System_Control dropdown selections.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (typeof PivotSortManager !== "undefined" && typeof PivotSortManager.handleEdit === "function") {
    PivotSortManager.handleEdit(e);
  }

  if (e && e.range) {
    if (
      e.range.getSheet().getName() !== SYSTEM_CONFIG.SHEETS.CONTROL ||
      e.range.getRow() !== 2
    ) {
      return;
    }
  }

  handleExternalChange(e);
}

// ============================================================================
// PUBLIC EXECUTION ENDPOINTS (CALLING 100% STATIC ENGINES)
// ============================================================================

function optimizeAll() {
  PortfolioManager.optimizeAll();
}

function resetAndRecalculateAll() {
  PortfolioManager.resetAndRecalculateAll();
}

function optimizeTransactionsAndStocks() {
  PortfolioManager.optimizeTransactionsAndStocks();
}

function optimizeGoogleFinance() {
  PortfolioManager.optimizeGoogleFinance();
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

function restoreAllFormulas() {
  if (typeof WorkbookAuditor !== "undefined" && typeof WorkbookAuditor.restoreAllFormulas === "function") {
    WorkbookAuditor.restoreAllFormulas();
  }
}

function fixAllArrayFormulaErrors(target = null) {
  if (typeof WorkbookAuditor !== "undefined" && typeof WorkbookAuditor.fixAllArrayFormulaErrors === "function") {
    WorkbookAuditor.fixAllArrayFormulaErrors(target);
  } else if (typeof FormulaAudit !== "undefined" && typeof FormulaAudit.fixAllErrors === "function") {
    FormulaAudit.fixAllErrors(target);
  }
}