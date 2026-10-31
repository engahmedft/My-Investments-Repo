/**
 * ============================================================================
 * STANDALONE HISTORICAL DIVIDEND SYNCHRONIZATION & RECONCILIATION ENGINE
 * ============================================================================
 * STRICT 1-ROW-PER-TRANSACTION | STRICT YYYY-MM-DD | VALUES ONLY
 */

function syncHistoricalDividends() {
  HistoricalDividendSync.run();
}

function updateSelectedDividendRow() {
  HistoricalDividendSync.updateActiveRow();
}

function updateAllNoMatchDividends() {
  HistoricalDividendSync.updateAllNoMatch();
}

function repairCorruptedTransactionDates() {
  HistoricalDividendSync.repairTransactionDates();
}

/*function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 Dividend Sync")
    .addItem("1. Sync Historical Dividends", "syncHistoricalDividends")
    .addSeparator()
    .addItem("2. Update Selected Row to Deserve Date", "updateSelectedDividendRow")
    .addItem("3. Update ALL 'No Match' Rows to Deserve Date", "updateAllNoMatchDividends")
    .addSeparator()
    .addItem("🛠️ Repair Corrupted Dates in Transactions", "repairCorruptedTransactionDates")
    .addToUi();
}*/

/**
 * In-Row Action Checkbox: Checking column 12 strictly updates ONLY Transaction Date
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Historical_Dividends") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();

  // Column 12 is "Update Date?" Checkbox
  if (col === 12 && row > 1 && e.value === "TRUE") {
    HistoricalDividendSync.updateRow(sheet, row);
  }
}

class HistoricalDividendSync {

  static run(targetSpreadsheet = null) {
    const ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const transSheet = ss.getSheetByName("Transactions");

    if (!transSheet) {
      ss.toast("Sheet 'Transactions' not found.", "Error", 5);
      return;
    }

    // -------------------------------------------------------------------------
    // 1. Build Metadata Map from Market_Feeds
    // -------------------------------------------------------------------------
    const mfSheet = ss.getSheetByName("Market_Feeds");
    const symbolMetaMap = new Map();

    if (mfSheet && mfSheet.getLastRow() >= 3) {
      const mfLastRow = mfSheet.getLastRow();
      const mfLastCol = mfSheet.getLastColumn();
      const mfData = mfSheet.getRange(2, 1, mfLastRow - 1, mfLastCol).getValues();
      const mfHeaders = mfData[0].map(h => String(h || "").trim());

      const idxMfSymbol = mfHeaders.indexOf("Symbol");
      const idxMfId     = mfHeaders.indexOf("MF_ID");
      const idxMarketId = mfHeaders.indexOf("Market_ID");
      const idxMfName   = mfHeaders.indexOf("Manual_Name");

      for (let i = 1; i < mfData.length; i++) {
        const row = mfData[i];
        const sym = String(row[idxMfSymbol] || "").trim().toUpperCase();
        if (sym) {
          symbolMetaMap.set(sym, {
            mfId: idxMfId !== -1 ? String(row[idxMfId] || "").trim() : "",
            marketId: idxMarketId !== -1 ? String(row[idxMarketId] || "").trim() : "",
            company: idxMfName !== -1 ? String(row[idxMfName] || "").trim() : ""
          });
        }
      }
    }

    // -------------------------------------------------------------------------
    // 2. Discover Buy Dates & Catalog Existing Dividends from Transactions
    // -------------------------------------------------------------------------
    const transLastRow = transSheet.getLastRow();
    if (transLastRow < 3) {
      ss.toast("No transactions found.", "Info", 5);
      return;
    }

    const transData = transSheet.getRange(2, 1, transLastRow - 1, transSheet.getLastColumn()).getValues();
    const tHeaders = transData[0].map(h => String(h || "").trim());

    const idxId    = tHeaders.indexOf("Transaction_ID");
    const idxType  = tHeaders.indexOf("Order Type");
    const idxSym   = tHeaders.indexOf("Symbol");
    const idxComp  = tHeaders.indexOf("Company");
    const idxExec  = tHeaders.indexOf("Execution Date");
    const idxTrans = tHeaders.indexOf("Transaction Date");
    const idxVal   = tHeaders.indexOf("Value");

    if (idxSym === -1) {
      ss.toast("Transactions sheet is missing 'Symbol' column.", "Error", 5);
      return;
    }

    const firstBuyMap = new Map();
    const companyNameMap = new Map();
    const recordedDivsMap = new Map();

    for (let i = 1; i < transData.length; i++) {
      const r = transData[i];
      const sym = String(r[idxSym] || "").trim().toUpperCase();
      if (!sym) continue;

      const company = idxComp !== -1 ? String(r[idxComp] || "").trim() : "";
      if (company && !companyNameMap.has(sym)) {
        companyNameMap.set(sym, company);
      }

      const orderType = idxType !== -1 ? String(r[idxType] || "").trim().toLowerCase() : "";
      const execTime = idxExec !== -1 ? HistoricalDividendSync._parseTime(r[idxExec]) : 0;
      const transTime = idxTrans !== -1 ? HistoricalDividendSync._parseTime(r[idxTrans]) : 0;
      const effectiveTime = transTime || execTime;

      if (orderType === "buy" || orderType === "gift" || orderType === "") {
        if (effectiveTime > 0) {
          const prevMin = firstBuyMap.get(sym);
          if (!prevMin || effectiveTime < prevMin) {
            firstBuyMap.set(sym, effectiveTime);
          }
        }
      }

      if (orderType === "dividend") {
        if (!recordedDivsMap.has(sym)) {
          recordedDivsMap.set(sym, []);
        }
        const perShareVal = idxVal !== -1 ? Math.abs(parseFloat(r[idxVal])) || 0 : 0;
        recordedDivsMap.get(sym).push({
          id: idxId !== -1 ? String(r[idxId] || "").trim() : "",
          transDateStr: HistoricalDividendSync._formatDateKey(r[idxTrans]),
          execDateStr: HistoricalDividendSync._formatDateKey(r[idxExec]),
          timestamp: effectiveTime,
          value: perShareVal,
          isUsed: false
        });
      }
    }

    if (firstBuyMap.size === 0) {
      ss.toast("No valid purchases found in Transactions.", "Warning", 5);
      return;
    }

    // -------------------------------------------------------------------------
    // 3. Batch Request Setup for Yahoo Finance Events
    // -------------------------------------------------------------------------
    const { cookie, crumb } = HistoricalDividendSync._getYahooAuth();
    const userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";
    const headers = { "User-Agent": userAgent };
    if (cookie) headers["Cookie"] = cookie;

    const nowSec = Math.floor(Date.now() / 1000);
    const requests = [];
    const tickerTracker = [];

    for (const [symbol, startMs] of firstBuyMap.entries()) {
      if (HistoricalDividendSync._isGold(symbol)) continue;

      const meta = symbolMetaMap.get(symbol);
      const isSaudi = (meta?.marketId === "302" || /^\d{4}$/.test(symbol));
      const yahooTicker = isSaudi ? `${symbol}.SR` : symbol;
      const companyName = companyNameMap.get(symbol) || meta?.company || symbol;

      const startSec = Math.max(0, Math.floor(startMs / 1000) - (4 * 86400));
      let url = `https://query1.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(yahooTicker)}?period1=${startSec}&period2=${nowSec}&interval=1d&events=div`;
      if (crumb) url += `&crumb=${encodeURIComponent(crumb)}`;

      requests.push({
        url: url,
        headers: headers,
        muteHttpExceptions: true
      });

      tickerTracker.push({
        symbol: symbol,
        yahooTicker: yahooTicker,
        mfId: meta?.mfId || "",
        company: companyName,
        minBuyMs: startMs,
        isSaudi: isSaudi
      });
    }

    // -------------------------------------------------------------------------
    // 4. Fetch & Parse Dividends
    // -------------------------------------------------------------------------
    ss.toast(`Fetching dividend records for ${requests.length} symbols...`, "In Progress", 10);
    const responses = UrlFetchApp.fetchAll(requests);
    const symbolDividendEvents = new Map();

    for (let i = 0; i < responses.length; i++) {
      const resp = responses[i];
      const tracker = tickerTracker[i];

      if (resp.getResponseCode() !== 200) {
        let errorDesc = `HTTP ${resp.getResponseCode()}`;
        try {
          const errJson = JSON.parse(resp.getContentText());
          if (errJson?.chart?.error?.description) {
            errorDesc += ` (${errJson.chart.error.description})`;
          }
        } catch (_) {}
        Logger.log(`[Skipped] ${tracker.yahooTicker}: ${errorDesc}`);
        continue;
      }

      try {
        const json = JSON.parse(resp.getContentText());
        const resultObj = json?.chart?.result?.[0];
        const dividendsObj = resultObj?.events?.dividends;
        if (!dividendsObj || typeof dividendsObj !== "object") continue;

        const eventsList = [];

        for (const divKey of Object.keys(dividendsObj)) {
          const item = dividendsObj[divKey];
          const exDateEpochMs = (item.date || Number(divKey)) * 1000;
          const exDateObj = new Date(exDateEpochMs);
          const divAmount = Number(item.amount) || 0;

          const deserveDateStr = HistoricalDividendSync._getDeserveDate(exDateObj, tracker.isSaudi);
          const deserveDateEndMs = new Date(deserveDateStr).setHours(23, 59, 59, 999);

          if (tracker.minBuyMs <= deserveDateEndMs && divAmount > 0) {
            eventsList.push({
              mfId: tracker.mfId,
              symbol: tracker.symbol,
              company: tracker.company,
              amount: divAmount,
              deserveDate: deserveDateStr,
              deserveDateMs: deserveDateEndMs
            });
          }
        }

        eventsList.sort((a, b) => a.deserveDateMs - b.deserveDateMs);
        symbolDividendEvents.set(tracker.symbol, eventsList);

      } catch (err) {
        Logger.log(`[Parse Error] ${tracker.yahooTicker}: ${err.message}`);
      }
    }

    // -------------------------------------------------------------------------
    // 5. Reconciliation: ZERO GROUPING (1 Row Per Transaction)
    // -------------------------------------------------------------------------
    const collectedRows = [];
    const DAY_MS = 24 * 60 * 60 * 1000;

    for (const [symbol, events] of symbolDividendEvents.entries()) {
      const existingDivs = recordedDivsMap.get(symbol) || [];
      existingDivs.sort((a, b) => a.timestamp - b.timestamp);

      for (const ev of events) {
        const availableCandidates = existingDivs.filter(d => {
          if (d.isUsed || d.timestamp <= 0) return false;
          const diffDays = (d.timestamp - ev.deserveDateMs) / DAY_MS;
          return diffDays >= -2.5 && diffDays <= 32;
        });

        let matchedTransactions = [];

        if (availableCandidates.length > 0) {
          availableCandidates.sort((a, b) => Math.abs(a.timestamp - ev.deserveDateMs) - Math.abs(b.timestamp - ev.deserveDateMs));
          const bestCandidate = availableCandidates[0];

          matchedTransactions = availableCandidates.filter(c => {
            const dateMatches = (c.transDateStr === bestCandidate.transDateStr);
            const closeInTime = Math.abs(c.timestamp - bestCandidate.timestamp) <= (2 * DAY_MS);
            return dateMatches || closeInTime;
          });

          matchedTransactions.forEach(m => m.isUsed = true);
        }

        if (matchedTransactions.length > 0) {
          for (const m of matchedTransactions) {
            const isDateIdentical = (m.transDateStr === ev.deserveDate);
            const isValueIdentical = (m.value > 0 ? Math.abs(m.value - ev.amount) < 0.01 : true);
            const status = (isDateIdentical && isValueIdentical) ? "Same" : "No match";

            collectedRows.push({
              mfId: ev.mfId,
              symbol: ev.symbol,
              company: ev.company,
              amount: ev.amount,
              transVal: m.value > 0 ? m.value : "",
              transId: m.id,
              deserveDate: ev.deserveDate,
              transDate: m.transDateStr,
              execDate: m.execDateStr,
              distributionDate: "",
              matchStatus: status,
              timestamp: ev.deserveDateMs
            });
          }
        } else {
          collectedRows.push({
            mfId: ev.mfId,
            symbol: ev.symbol,
            company: ev.company,
            amount: ev.amount,
            transVal: "",
            transId: "",
            deserveDate: ev.deserveDate,
            transDate: "",
            execDate: "",
            distributionDate: "",
            matchStatus: "Not found",
            timestamp: ev.deserveDateMs
          });
        }
      }
    }

    const collator = new Intl.Collator(undefined, { numeric: true, sensitivity: "base" });
    collectedRows.sort((a, b) => {
      const sComp = collator.compare(a.symbol, b.symbol);
      if (sComp !== 0) return sComp;
      return a.timestamp - b.timestamp;
    });

    // -------------------------------------------------------------------------
    // 6. Reset, Format & Write to 'Historical_Dividends' Sheet
    // -------------------------------------------------------------------------
    let divSheet = ss.getSheetByName("Historical_Dividends");
    if (!divSheet) {
      divSheet = ss.insertSheet("Historical_Dividends");
    } else {
      divSheet.clear();
      divSheet.getDataRange().clearDataValidations();
    }

    const headersList = [
      "MF_ID",              // Col 1
      "Symbol",             // Col 2
      "Company",            // Col 3
      "Value Per Share",    // Col 4 (Yahoo)
      "Transaction Value",  // Col 5 (Transactions Counterpart)
      "Transaction_ID",     // Col 6
      "Deserve Date",       // Col 7 (After Transaction_ID)
      "Transaction Date",   // Col 8
      "Execution Date",     // Col 9 (After Transaction Date)
      "Distribution Date",  // Col 10
      "Match Status",       // Col 11
      "Update Date?"        // Col 12 (Interactive Checkbox)
    ];

    const outputRows = [headersList];

    for (const r of collectedRows) {
      outputRows.push([
        r.mfId || "N/A",
        r.symbol,
        r.company,
        r.amount,
        r.transVal !== "" ? r.transVal : "",
        r.transId || "",
        r.deserveDate,
        r.transDate || "",
        r.execDate || "",
        r.distributionDate,
        r.matchStatus,
        false
      ]);
    }

    const numRows = outputRows.length;
    const numCols = headersList.length;

    divSheet.getRange(1, 1, numRows, numCols).setValues(outputRows);

    // Style Header
    divSheet.getRange(1, 1, 1, numCols)
      .setFontWeight("bold")
      .setBackground("#1c4587")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center");
    divSheet.setFrozenRows(1);

    if (numRows > 1) {
      // Format Values (Cols 4-5)
      divSheet.getRange(2, 4, numRows - 1, 2)
        .setNumberFormat("#,##0.0000")
        .setHorizontalAlignment("right");

      // Format Dates (Cols 7, 8, 9, 10: Deserve, Trans, Exec, Dist)
      divSheet.getRange(2, 7, numRows - 1, 4)
        .setNumberFormat("yyyy-mm-dd")
        .setHorizontalAlignment("center");

      // Alignments
      divSheet.getRange(2, 1, numRows - 1, 2).setHorizontalAlignment("center");
      divSheet.getRange(2, 6, numRows - 1, 1).setHorizontalAlignment("center");
      divSheet.getRange(2, 11, numRows - 1, 1).setHorizontalAlignment("center");
      divSheet.getRange(2, 3, numRows - 1, 1).setHorizontalAlignment("left");

      // Insert Checkbox buttons strictly into Column 12
      divSheet.getRange(2, 12, numRows - 1, 1).insertCheckboxes();
    }

    for (let col = 1; col <= numCols; col++) {
      divSheet.autoResizeColumn(col);
    }

    ss.toast(`Synced ${collectedRows.length} rows. Clean 1-to-1 transactions.`, "Complete", 5);
  }

  // =========================================================================
  // TRANSACTION DATE UPDATE ACTIONS (STRICT YYYY-MM-DD | VALUES ONLY)
  // =========================================================================

  static updateRow(sheet, rowNumber) {
    const ss = sheet.getParent();
    const rowValues = sheet.getRange(rowNumber, 1, 1, 12).getValues()[0];

    const targetId = String(rowValues[5] || "").trim();                     // Col 6: Transaction_ID
    const rawDeserveDate = rowValues[6];                                   // Col 7: Deserve Date
    const cleanDeserveDate = HistoricalDividendSync._formatDateKey(rawDeserveDate); // Strictly yyyy-mm-dd

    if (!targetId || !cleanDeserveDate) {
      ss.toast(`Row ${rowNumber}: No Transaction_ID or Deserve Date found.`, "Warning", 4);
      sheet.getRange(rowNumber, 12).setValue(false);
      return;
    }

    const updatedCount = HistoricalDividendSync._applyUpdateToTransactions(ss, targetId, cleanDeserveDate);

    if (updatedCount > 0) {
      // Update Historical_Dividends display strictly with yyyy-mm-dd
      const transDateCell = sheet.getRange(rowNumber, 8);
      transDateCell.setValue(cleanDeserveDate);
      transDateCell.setNumberFormat("yyyy-mm-dd");

      sheet.getRange(rowNumber, 11).setValue("Same");
      sheet.getRange(rowNumber, 12).setValue(false);

      ss.toast(`Updated Transaction ${targetId} to ${cleanDeserveDate} (Values Only).`, "Success", 3);
    } else {
      ss.toast(`Could not find Transaction_ID ${targetId} in Transactions sheet.`, "Error", 5);
      sheet.getRange(rowNumber, 12).setValue(false);
    }
  }

  static updateActiveRow() {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() !== "Historical_Dividends") {
      SpreadsheetApp.getActiveSpreadsheet().toast("Please select a row in 'Historical_Dividends'.", "Info", 4);
      return;
    }
    const row = sheet.getActiveCell().getRow();
    if (row <= 1) return;
    HistoricalDividendSync.updateRow(sheet, row);
  }

  static updateAllNoMatch() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Historical_Dividends");
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();
    let updatedRows = 0;

    for (let i = 0; i < data.length; i++) {
      const status = String(data[i][10] || "").trim(); // Col 11: Match Status
      if (status === "No match") {
        HistoricalDividendSync.updateRow(sheet, i + 2);
        updatedRows++;
      }
    }

    ss.toast(`Updated ${updatedRows} 'No match' dividend rows to Deserve Dates.`, "Complete", 5);
  }

  /**
   * Strictly updates ONLY the value and format of Transaction Date.
   * Completely preserves fonts, background colors, alignments, and borders.
   */
  static _applyUpdateToTransactions(ss, targetId, deserveDateStr) {
    const transSheet = ss.getSheetByName("Transactions");
    if (!transSheet) return 0;

    const lastRow = transSheet.getLastRow();
    if (lastRow < 3) return 0;

    const headers = transSheet.getRange(2, 1, 1, transSheet.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const idxId = headers.indexOf("Transaction_ID");
    const idxTransDate = headers.indexOf("Transaction Date");
    const idxLastUpdate = headers.indexOf("Last Update");

    if (idxId === -1 || idxTransDate === -1) return 0;

    const transValues = transSheet.getRange(3, 1, lastRow - 2, transSheet.getLastColumn()).getValues();
    const matchId = String(targetId).trim();
    let matchCount = 0;
    const nowTime = new Date();

    for (let r = 0; r < transValues.length; r++) {
      const rowId = String(transValues[r][idxId] || "").trim();
      if (rowId === matchId) {
        const targetRowNumber = r + 3;
        const targetCell = transSheet.getRange(targetRowNumber, idxTransDate + 1);

        // Paste as pure value and enforce yyyy-mm-dd number format (Leaves all styles intact)
        targetCell.setValue(deserveDateStr);
        targetCell.setNumberFormat("yyyy-mm-dd");

        if (idxLastUpdate !== -1) {
          transSheet.getRange(targetRowNumber, idxLastUpdate + 1).setValue(nowTime);
        }
        matchCount++;
        break;
      }
    }

    return matchCount;
  }

  /**
   * One-click repair utility: Scans 'Transactions' sheet and converts any verbose 
   * date strings (e.g. 'Tue Sep 23 2025...') back to clean 'yyyy-mm-dd' values.
   */
  static repairTransactionDates(targetSpreadsheet = null) {
    const ss = targetSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
    const transSheet = ss.getSheetByName("Transactions");
    if (!transSheet) return;

    const lastRow = transSheet.getLastRow();
    if (lastRow < 3) return;

    const headers = transSheet.getRange(2, 1, 1, transSheet.getLastColumn()).getValues()[0].map(h => String(h || "").trim());
    const idxTransDate = headers.indexOf("Transaction Date");
    if (idxTransDate === -1) return;

    const dateRange = transSheet.getRange(3, idxTransDate + 1, lastRow - 2, 1);
    const dateValues = dateRange.getValues();
    let fixedCount = 0;

    for (let i = 0; i < dateValues.length; i++) {
      const val = dateValues[i][0];
      if (val) {
        const cleaned = HistoricalDividendSync._formatDateKey(val);
        if (cleaned && cleaned !== val) {
          dateValues[i][0] = cleaned;
          fixedCount++;
        }
      }
    }

    if (fixedCount > 0) {
      dateRange.setValues(dateValues);
      dateRange.setNumberFormat("yyyy-mm-dd");
      ss.toast(`Repaired ${fixedCount} dates back to yyyy-mm-dd.`, "Complete", 5);
    } else {
      dateRange.setNumberFormat("yyyy-mm-dd");
      ss.toast("Dates are already formatted as yyyy-mm-dd.", "Info", 4);
    }
  }

  // =========================================================================
  // HELPER METHODS
  // =========================================================================

  static _getDeserveDate(exDate, isSaudi) {
    const d = new Date(exDate.getTime());
    const dayOfWeek = d.getDay();

    if (isSaudi) {
      if (dayOfWeek === 0) d.setDate(d.getDate() - 3);      // Sunday -> Thursday
      else if (dayOfWeek === 6) d.setDate(d.getDate() - 2); // Saturday -> Thursday
      else d.setDate(d.getDate() - 1);                      // Mon-Thu -> Previous day
    } else {
      if (dayOfWeek === 1) d.setDate(d.getDate() - 3);      // Monday -> Friday
      else if (dayOfWeek === 0) d.setDate(d.getDate() - 2); // Sunday -> Friday
      else d.setDate(d.getDate() - 1);                      // Tue-Fri -> Previous day
    }

    return HistoricalDividendSync._formatDateKey(d);
  }

  static _parseTime(val) {
    if (!val) return 0;
    if (typeof val === "number") return val;
    if (val instanceof Date) return val.getTime();
    const t = new Date(val).getTime();
    return Number.isNaN(t) ? 0 : t;
  }

  /**
   * Coerces any date representation (Date object, timestamp, or verbose text)
   * into a strictly standardized 'yyyy-mm-dd' string.
   */
  static _formatDateKey(val) {
    if (!val) return "";
    let d;
    if (val instanceof Date) {
      d = val;
    } else {
      const str = String(val).trim();
      const match = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
      if (match) {
        const pad = (n) => (Number(n) < 10 ? "0" + Number(n) : String(n));
        return `${match[1]}-${pad(match[2])}-${pad(match[3])}`;
      }
      d = new Date(str);
    }
    if (Number.isNaN(d.getTime())) return "";
    const pad = (n) => (n < 10 ? "0" + n : String(n));
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }

  static _isGold(sym) {
    const s = String(sym || "").trim().toLowerCase();
    return s.includes("gold") || s.startsWith("gc=") || s.startsWith("xau") || s.includes("ذهب");
  }

  static _getYahooAuth() {
    const cache = CacheService.getScriptCache();
    const cachedCookie = cache.get("YF_COOKIE");
    const cachedCrumb = cache.get("YF_CRUMB");

    if (cachedCookie && cachedCrumb) {
      return { cookie: cachedCookie, crumb: cachedCrumb };
    }

    try {
      const userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";
      const resp1 = UrlFetchApp.fetch("https://fc.yahoo.com/", {
        muteHttpExceptions: true,
        headers: { "User-Agent": userAgent }
      });

      const setCookie = resp1.getAllHeaders()["Set-Cookie"];
      let cookie = "";
      if (setCookie) {
        cookie = Array.isArray(setCookie)
          ? setCookie.map(c => c.split(";")[0]).join("; ")
          : String(setCookie).split(";")[0];
      }

      let crumb = "";
      if (cookie) {
        const resp2 = UrlFetchApp.fetch("https://query2.finance.yahoo.com/v1/test/getcrumb", {
          muteHttpExceptions: true,
          headers: { "User-Agent": userAgent, "Cookie": cookie }
        });
        if (resp2.getResponseCode() === 200) {
          crumb = resp2.getContentText().trim();
        }
      }

      if (cookie && crumb) {
        cache.put("YF_COOKIE", cookie, 21600);
        cache.put("YF_CRUMB", crumb, 21600);
        return { cookie, crumb };
      }
    } catch (e) {
      Logger.log(`Yahoo auth exception: ${e.message}`);
    }

    return { cookie: "", crumb: "" };
  }
}