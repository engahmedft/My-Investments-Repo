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
  loadCommands() {
    const table = BatchProcessor.read(SYSTEM_CONFIG.SHEETS.LOOKUP, SYSTEM_CONFIG.LOOKUP.HEADER_ROW);
    if (!table || table.length <= 1) return [];

    const headers = table[0];
    const getIdx = (colName) => headers.indexOf(colName.toLowerCase());

    const catIdx  = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.CATEGORY);
    const codeIdx = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.CODE);
    const descIdx = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.LABEL);
    const actIdx  = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.ACTIVE);
    const sepIdx  = getIdx(SYSTEM_CONFIG.LOOKUP.COLS.SEPARATOR);

    if (catIdx === -1 || codeIdx === -1 || descIdx === -1) {
      Logger.log(`[CommandEngine] Required columns missing in ${SYSTEM_CONFIG.SHEETS.LOOKUP}`);
      return [];
    }

    const commands = [];
    for (let i = 1; i < table.length; i++) {
      const row = table[i];
      if (Number(row[catIdx]) !== SYSTEM_CONFIG.LOOKUP.CATEGORY_ID) continue;

      const isAct = (row[actIdx] === true || String(row[actIdx]).toUpperCase() === "Y" || row[actIdx] === 1);
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

  let actionCode = idxCode !== -1 ? String(row[idxCode] ?? "").trim() : "";
  const displayLabel = String(row[idxAction] ?? "").trim();

  if (SYSTEM_CONFIG.IDLE_VALUES.has(actionCode)) {
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
    if (idxCode !== -1) row[idxCode] = actionCode;
  }

  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
  const getTimestamp = () => Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

  row[idxState] = "Running";
  row[idxDate] = getTimestamp();
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

function optimizeGoogleFinance() {
  new PortfolioManager().optimizeGoogleFinance();
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