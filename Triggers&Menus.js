/**
 * ============================================================================
 * TRIGGER DISPATCHER & COMMAND ENGINE
 * Handles UI menus, onEdit triggers, and System_Control state machines.
 * STRICT CANONICAL 100% STATIC EDITION: Zero legacy aliases, zero fallbacks.
 * ============================================================================
 */
const SYSTEM_CONFIG = Object.freeze({
  SHEETS: {
    CONTROL: "System_Control",
    LOOKUP: "Lookup_Values"
  },
  LOOKUP: {
    HEADER_ROW: 2,
    CATEGORY_ID: 2,
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
  _isRunning: false,

  /**
   * Executes any command with live status updates (Running / Completed / Error) logged to System_Control.
   * Includes a re-entrancy guard to prevent redundant state writes.
   */
  runWithStatus(actionCode, actionFn, spreadsheet = null) {
    if (this._isRunning) {
      return actionFn();
    }

    const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.CONTROL);

    if (!sheet || sheet.getLastColumn() === 0) {
      return actionFn();
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(SYSTEM_CONFIG.LOCK_TIMEOUT_MS)) {
      throw new Error("Could not obtain lock. Another command is currently executing.");
    }

    this._isRunning = true;
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
      this._isRunning = false;
      lock.releaseLock();
      throw new Error(`[System_Control] Missing canonical columns in Row 1.`);
    }

    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    const getTimestamp = () => Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    // Strictly set State to "Running"
    row[idxState] = "Running";
    row[idxDate]  = getTimestamp();
    row[idxError] = "";
    if (idxCode !== -1) row[idxCode] = actionCode;

    sheet.getRange(2, 1, 1, numCols).setValues([row]);
    SpreadsheetApp.flush();

    let errorOccurred = null;
    let result = null;

    try {
      result = actionFn();
    } catch (err) {
      errorOccurred = err.message || String(err);
      Logger.log(`❌ CommandEngine Execution Error: ${err.stack || err.message}`);
    } finally {
      row[idxAction] = "- Select Action -";
      if (idxCode !== -1) row[idxCode] = "IDLE";
      row[idxState]  = errorOccurred ? "Error" : "Completed";
      row[idxDate]   = getTimestamp();
      row[idxError]  = errorOccurred ? errorOccurred : "";

      sheet.getRange(2, 1, 1, numCols).setValues([row]);
      SpreadsheetApp.flush();

      this._isRunning = false;
      lock.releaseLock();

      if (errorOccurred) {
        throw new Error(errorOccurred);
      }
    }

    return result;
  },

  /**
   * Loads active system commands strictly via BatchProcessor.read without range fallback.
   * Filters exclusively for Category_ID = 5 and Is_Active = 'Y'.
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

    if (catIdx === -1 || codeIdx === -1 || descIdx === -1 || actIdx === -1) {
      Logger.log(`[CommandEngine] Required columns missing in ${SYSTEM_CONFIG.SHEETS.LOOKUP}`);
      return [];
    }

    const commands = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];

      if (Number(row[catIdx]) !== SYSTEM_CONFIG.LOOKUP.CATEGORY_ID) continue;
      if (String(row[actIdx] ?? "").trim().toUpperCase() !== "Y") continue;

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
 */
function handleExternalChange(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SYSTEM_CONFIG.SHEETS.CONTROL);
  if (!sheet) return;

  const numCols = sheet.getLastColumn();
  const buffer = sheet.getRange(1, 1, 2, numCols).getValues();
  const headers = buffer[0].map(h => String(h ?? "").trim());
  const row = buffer[1];

  const idxAction = headers.indexOf("Action");
  const idxCode   = headers.indexOf("Action Code");

  if (idxAction === -1) return;

  const displayLabel = String(row[idxAction] ?? "").trim();
  let actionCode = "";

  if (!SYSTEM_CONFIG.IDLE_VALUES.has(displayLabel)) {
    const match = CommandEngine.loadCommands().find(c => c.label.toLowerCase() === displayLabel.toLowerCase());
    if (match) {
      actionCode = match.code;
    }
  }

  if (!actionCode && idxCode !== -1 && !SYSTEM_CONFIG.IDLE_VALUES.has(String(row[idxCode] ?? "").trim())) {
    actionCode = String(row[idxCode]).trim();
  }

  if (!actionCode) return;

  // Route execution through CommandEngine.runWithStatus
  CommandEngine.runWithStatus(actionCode, () => {
    if (typeof globalThis[actionCode] !== "function") {
      throw new Error(`Automation error: Function "${actionCode}" is not defined.`);
    }
    globalThis[actionCode]();
  }, ss);
}

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
// PUBLIC EXECUTION ENDPOINTS (WRAPPED WITH COMMANDENGINE.RUNWITHSTATUS)
// ============================================================================

function optimizeAll() {
  CommandEngine.runWithStatus("optimizeAll", () => PortfolioManager.optimizeAll());
}

function resetAndRecalculateAll() {
  CommandEngine.runWithStatus("resetAndRecalculateAll", () => PortfolioManager.resetAndRecalculateAll());
}

function optimizeTransactionsAndStocks() {
  CommandEngine.runWithStatus("optimizeTransactionsAndStocks", () => PortfolioManager.optimizeTransactionsAndStocks());
}

function optimizeMarketFeeds() {
  CommandEngine.runWithStatus("optimizeMarketFeeds", () => PortfolioManager.optimizeMarketFeeds());
}

function listAllFormulas() {
  CommandEngine.runWithStatus("listAllFormulas", () => {
    if (typeof WorkbookRegistry !== "undefined") WorkbookRegistry.listAllFormulas("Formula_Registry");
  });
}

function listAllTables() {
  CommandEngine.runWithStatus("listAllTables", () => {
    if (typeof WorkbookRegistry !== "undefined") WorkbookRegistry.listAllTables("Table_Registry");
  });
}

function runAllRegistries() {
  CommandEngine.runWithStatus("runAllRegistries", () => {
    if (typeof WorkbookRegistry !== "undefined") WorkbookRegistry.runAllRegistries();
  });
}

function restoreAllFormulas() {
  CommandEngine.runWithStatus("restoreAllFormulas", () => {
    if (typeof WorkbookRegistry !== "undefined") WorkbookRegistry.restoreAllFormulas("Formula_Registry");
  });
}

function fixAllArrayFormulaErrors(target = null) {
  CommandEngine.runWithStatus("fixAllArrayFormulaErrors", () => {
    if (typeof WorkbookRegistry !== "undefined") WorkbookRegistry.fixAllArrayFormulaErrors(target);
  });
}

function setupPivotSortControls() {
  CommandEngine.runWithStatus("setupPivotSortControls", () => {
    if (typeof PivotSortManager !== "undefined") PivotSortManager.setupControls();
  });
}

function extendFormat(sheet) {
  CommandEngine.runWithStatus("extendFormat", () => {
    if (typeof TableFormatter !== "undefined") {
      TableFormatter.extend(sheet || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
    }
  });
}