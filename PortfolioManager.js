/**
 * Core orchestrator and calculation engine for investment portfolio operations.
 * STRICT CANONICAL 100% STATIC EDITION: Zero object instantiation, zero aliases.
 * All methods are pure static functions.
 * 
 * Run `PortfolioManager.help()` for full architectural specifications, fee models, and API contracts.
 */
class PortfolioManager {

  /**
   * Prevent instantiation of purely static utility class.
   */
  constructor() {
    throw new Error("PortfolioManager is a static utility class and cannot be instantiated. Call methods directly via PortfolioManager.<methodName>().");
  }

  /**
   * Returns and logs the engine specifications, fee formulas, pipeline stages, and API contracts.
   * @returns {string} Formatted architectural specification.
   */
  static help() {
    const doc = `
====================================================================================================
PORTFOLIO MANAGER ENGINE SPECIFICATION & ARCHITECTURE (100% STATIC)
====================================================================================================

1. UNIFIED PIPELINE OVERVIEW:
   Master static orchestrator executing the 6-stage investment portfolio synchronization:
     Stage 1: Google Finance optimization, multi-tier sorting (Market_ID first, then missing
              52-week min rows [hasNoMin], then natural alphanumeric GF_Key), parallel Yahoo
              Finance price backfill (syncing received Yahoo price and 52-week low into
              Manual_Price and Manual_Year_Min while leaving manual values untouched if no
              quote returns), and timestamp formatting (Optimize_Date with full time).
     Stage 2: Transaction chronological sort (Execution Date -> Trans Date -> Company -> Order Type
              -> Last Update -> Descending Avg Cost), running average cost basis calculation,
              future dividend quantity synchronization, and cross-platform dividend auto-replication.
     Stage 3: Position synchronization into Stocks sheet (updating average cost and inserting new positions).
     Stage 4: Automated formula and table boundary audits (runAllAudits / listAllTables / listAllFormulas).
     Stage 5: Audit log persistence in Optimize_Log sheet with granular operation records.
     Stage 6: Conditional formatting and table visual styling reconciliation via TableFormatter.

2. UNIFIED STATIC EXECUTION INTERFACES & API CONTRACTS:
   All operations are 100% static, stateless, and allocate zero instance wrappers.
   - PortfolioManager.optimizeAll(options):
     Full 6-stage pipeline execution across all sheets, price backfills, audits, and formatting.
     Options: { reset = false, runAudits = true, format = true, spreadsheet = null }
       * reset: Clears 'Optimize Date' checkpoints to force a full baseline recalculation.
       * runAudits: Executes Stage 4 formula and table boundary audits.
       * format: Executes Stage 6 conditional formatting via TableFormatter.
       * spreadsheet: Optional target Spreadsheet instance (defaults to SpreadsheetApp.getActiveSpreadsheet()).

   - PortfolioManager.optimizeTransactionsAndStocks(options):
     Ultra-fast offline execution strictly running Stage 2 (Transactions) & Stage 3 (Stocks).
     Does not execute Google Finance optimization, external HTTP calls, formula audits, or formatting.
     Options: { reset = false, spreadsheet = null }
       * reset: Clears 'Optimize Date' checkpoints to force a full baseline recalculation.
       * spreadsheet: Optional target Spreadsheet instance.

   - PortfolioManager.resetAndRecalculateAll(options):
     Convenience wrapper invoking optimizeAll({ reset: true }) to force a full baseline recalculation.

   - PortfolioManager.optimizeGoogleFinance(logEntries, spreadsheet):
     Targeted optimization, key computation, multi-tier sorting (Market_ID -> hasNoMin -> GF_Key),
     timestamp formatting (Optimize_Date stamped with full time 'yyyy-mm-dd hh:mm:ss'), parallel Yahoo
     Finance backfill, and recording incoming Yahoo price and 52-week low directly into Manual_Price
     and Manual_Year_Min (leaving manual fields untouched if the API returns no quote).

   - PortfolioManager.runWithStatus(actionNameOrFn, maybeFn, spreadsheet):
     State machine wrapper logging live execution status (Running/Completed/Error) and timestamps
     to the System_Control sheet.

3. PARALLEL HTTP ARCHITECTURE & ASSET NORMALIZATION:
   - Yahoo Finance quotes are requested concurrently via UrlFetchApp.fetchAll().
   - Multi-batch payloads (50 symbols/request) execute simultaneously in parallel rather than sequentially.
   - Authentication cookie and crumb are retrieved and cached for 6 hours (21,600s) via CacheService.
   - Manual Price/Min Recording:
     * When Yahoo returns a quote with price > 0, it is recorded into Manual_Price.
     * When Yahoo returns a 52-week low > 0, it is recorded into Manual_Year_Min.
     * If Yahoo returns no quote, Manual_Price and Manual_Year_Min remain unmodified, while Yahoo columns
       are reset to empty strings.
   - Market Ticker Normalization:
     * Saudi Tadawul (Market 702 or 4-digit numeric symbols): Appends '.SR' (e.g., '1120.SR').
     * Gold Assets (symbols matching /gold/i or 'Gold.'): Mapped to COMEX Gold Futures ('GC=F').
     * Gold Price Conversion: Converted from USD per Troy Ounce to SAR per Gram:
       Formula: Price_SAR_per_Gram = Price_USD_per_Oz * (3.75 / 31.1034768).

4. MARKET & TRANSACTION FEE ENGINE:
   - Saudi Market (Tadawul) ['SA']:
     * CMA Fee: 0.03% (0.0003) of total trade value.
     * Tadawul Tariff: 0.009% + (2 * 0.005%) + 0.001% = 0.02% of total trade value.
     * Brokerage Commissions:
       - Sahm Capital / Sahm2 Capital: 0%.
       - Alrajhi Capital / Alrajhi Wallet: 0.105% (0.00105) of total trade value.
       - Default: 0.1009% (0.001009) of total trade value.
     * VAT: 15% applied to Tadawul tariffs and brokerage commissions.
     * Dividends: 0% tax withholding.
   - US Market ['US']:
     * Brokerage Commissions:
       - Sahm Capital / Sahm2 Capital:
         Base fee: $0.49 (share price <= $5) or $1.99 (share price > $5).
         Formula: Math.max(baseFee, Math.min(qty * $0.015, tradeValue * 1.5%)).
       - Abyan: Math.max($0.95, qty * $0.0075).
       - Alrajhi Capital / Alrajhi Wallet: Math.max($1.99, qty * $0.0199).
     * Brokerage VAT: 15% local VAT applied to US brokerage commissions.
     * Regulatory & Exchange Fees:
       - Sahm Capital:
         * Settlement Fee: Math.max(qty * $0.003, $0.01).
         * CAT Fee: Math.max(qty * $0.000027, $0.01).
         * FINRA TAF (Sales only): Math.min(Math.max(qty * $0.000166, $0.01), $8.30).
       - Non-Sahm / Default (Sales only; excluded for Abyan):
         * SEC Fee: tradeValue * 0.0000278.
         * FINRA TAF: Math.min(Math.max(qty * $0.000195, $0.01), $5.95).
     * Dividends: 30% US withholding tax automatically deducted.

5. DIVIDEND REPLICATION & QUANTITY SYNCHRONIZATION:
   - Future Dividend Synchronization: For dividend rows with execution dates in the future
     (tTrans > todayLimit), automatically reconciles transaction quantity to match active holding quantity.
   - Cross-Platform Replication: For dividend events dated today or forward (tTrans >= todayStartMs),
     automatically replicates identical dividend transactions across all other platforms holding the
     stock, calculating custodian-specific fees and logging new Transaction_ID serials.

6. FORMULA SHIELDING & AUDIT PERSISTENCE:
   - ARRAY_FORMULAS Registry: Protects live calculated columns across Google_Finance_Data,
     Transactions, Stocks, and Platforms by nulling them during batch writebacks.
   - Optimize_Log: Retains a comprehensive audit trail with action classification (INSERT, UPDATE),
     order numbers, ticker symbols, platforms, and calculated parameter diffs.
====================================================================================================`.trim();

    Logger.log(doc);
    return doc;
  }

  /**
   * Pre-compiled Date Regex for fast O(1) date key parsing (yyyy-mm-dd or yyyy/mm/dd).
   * @private
   */
  static get _DATE_REGEX() {
    if (!PortfolioManager._DATE_REGEX_CACHE) {
      PortfolioManager._DATE_REGEX_CACHE = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/;
    }
    return PortfolioManager._DATE_REGEX_CACHE;
  }

  /**
   * Pre-compiled, reusable alphanumeric collator for fast natural sorting.
   * Avoids re-parsing Intl options inside O(N log N) sort comparators.
   * @private
   */
  static get _COLLATOR() {
    if (!PortfolioManager._COLLATOR_CACHE) {
      PortfolioManager._COLLATOR_CACHE = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
    }
    return PortfolioManager._COLLATOR_CACHE;
  }

  /**
   * Canonical registry of ArrayFormula column headers per sheet.
   * Nullified during write-backs to prevent Google Apps Script from overwriting live sheet formulas.
   * @returns {Record<string, string[]>}
   */
  static get ARRAY_FORMULAS() {
    if (!PortfolioManager._ARRAY_FORMULAS_CACHE) {
      PortfolioManager._ARRAY_FORMULAS_CACHE = {
        "Google_Finance_Data": [
          "Google_Name",
          "Google_Min",
          "Google_Price",
          "Year_Min",
          "Current_Price"
        ],
        "Transactions": [
          "Exchange",
          "Company",
          "Value + Fee",
          "Total Value",
          "Total Value + Fee",
          "P/L",
          "Total P/L",
          "Percentage"
        ],
        "Stocks": [
          "Exchange",
          "Company",
          "Combined Name",
          "Field",
          "Quantity",
          "Year Min",
          "Market Value",
          "Liquidity Value",
          "Tot.Mkt.Value",
          "Tot. Liquidity Value",
          "Total Cost",
          "Min + 0.5%",
          "Market +0.5%",
          "Mkt - Min",
          "Mkt - Min %",
          "Cost - Min",
          "Cost - Min%",
          "Threshold Quantity",
          "Quantity To Sell",
          "Safe Sell Price",
          "Quantity Of Min+0.5%",
          "Cash Needed",
          "Extra Quantity",
          "Holding",
          "Selling P/L",
          "Dividend",
          "Dividend %",
          "Fut.Dividend",
          "Fut.Dividend %",
          "Share P/L",
          "Tot. Shares P/L",
          "Shares P/L %",
          "Total Portfolio P/L",
          "Total P/L",
          "Total P/L %",
          "10% Profit",
          "10% Loss"
        ],
        "Platforms": [
          "Active Stocks",
          "Inactive Stocks",
          "Cost Portfolio",
          "Market Portfolio",
          "Liquidity Portfolio",
          "Holding",
          "Total Cash",
          "Total Cost",
          "Tot.Mkt.Value",
          "Tot. Liq. Value",
          "Selling P/L",
          "Dividend",
          "Div %",
          "Fut.Dividend",
          "Fut.Div %",
          "Tot. Shares P/L",
          "Portfolio P/L",
          "Total P/L",
          "Total P/L %"
        ],
        "System_Control": [],
        "Optimize_Log": []
      };
    }
    return PortfolioManager._ARRAY_FORMULAS_CACHE;
  }

  /**
   * Retrieves ArrayFormula column names registered for a specific sheet.
   * @param {string} sheetName - Sheet name to inspect.
   * @returns {string[]} Array of formula column headers.
   */
  static getArrayFormulas(sheetName) {
    return PortfolioManager.ARRAY_FORMULAS[sheetName] || [];
  }

  /**
   * Resolves a target spreadsheet instance.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [spreadsheet]
   * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
   */
  static getSpreadsheet(spreadsheet = null) {
    return spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  }

  /**
   * Retrieves a sheet from a target spreadsheet.
   * @param {string} name
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [spreadsheet]
   * @returns {GoogleAppsScript.Spreadsheet.Sheet|null}
   */
  static getSheet(name, spreadsheet = null) {
    const ss = PortfolioManager.getSpreadsheet(spreadsheet);
    return ss.getSheetByName(name);
  }

  /**
   * Zero-pads a 1- or 2-digit number to two characters.
   * @private
   */
  static _pad2(n) {
    return n < 10 ? '0' + n : String(n);
  }

  /**
   * Formats a Date object into an ISO date key (yyyy-MM-dd).
   * @private
   */
  static _fmtDate(d) {
    return `${d.getFullYear()}-${PortfolioManager._pad2(d.getMonth() + 1)}-${PortfolioManager._pad2(d.getDate())}`;
  }

  /**
   * Coerces arbitrary input into a millisecond epoch timestamp.
   * @param {any} val - Input date or timestamp.
   * @returns {number} Millisecond timestamp or 0 if invalid.
   */
  static getTime(val) {
    if (!val) return 0;
    if (typeof val === "number") return val;
    if (val instanceof Date) return val.getTime();
    const d = new Date(val);
    const t = d.getTime();
    return Number.isNaN(t) ? 0 : t;
  }

  /**
   * Formats a raw date value into a standardized date key (yyyy-MM-dd).
   * @param {any} val - Input date representation.
   * @returns {string} Standardized date key or empty string if invalid.
   */
  static formatDateKey(val) {
    if (!val) return "";
    if (val instanceof Date) {
      return Number.isNaN(val.getTime()) ? "" : PortfolioManager._fmtDate(val);
    }
    const str = String(val).trim();
    const match = str.match(PortfolioManager._DATE_REGEX);
    if (match) {
      return `${match[1]}-${PortfolioManager._pad2(match[2])}-${PortfolioManager._pad2(match[3])}`;
    }
    const d = new Date(str);
    return Number.isNaN(d.getTime()) ? "" : PortfolioManager._fmtDate(d);
  }

  /**
   * Determines the active cost basis, prioritizing realized cost over forward projections.
   * @param {{ normalAvgCost: number|null, futureAvgCost: number|null }} item
   * @returns {number} Effective average cost.
   */
  static getEffectiveCost(item) {
    if (!item) return 0;
    return item.normalAvgCost !== null ? item.normalAvgCost : item.futureAvgCost;
  }

  /**
   * Dynamic transaction fee and tax calculation engine for Saudi and US markets.
   */
  static calculateFee(exchange, platform, orderType, quantity, price) {
    const q = typeof quantity === "number" ? Math.abs(quantity) : (Math.abs(parseFloat(quantity)) || 0);
    const p = typeof price === "number" ? Math.abs(price) : (Math.abs(parseFloat(price)) || 0);

    if (q === 0 || p === 0) return 0;

    const tv = q * p;
    const o = String(orderType || "").trim();
    const ex = String(exchange || "").trim();

    const isSaudi = (ex === "SA");
    const isUs = (ex === "US");
    const isSell = (o === "Sell");
    const isDiv = (o === "Dividend");
    const sign = (isSell || isDiv) ? -1 : 1;
    const round2 = (val) => Math.round(val * 100) / 100;

    // Dividend Withholding Tax: US 30%; Saudi Arabia 0%
    if (isDiv) {
      if (isSaudi) return 0;
      if (isUs) return round2(tv * 0.3) * sign;
      return 0;
    }

    const isSahm = platform === "Sahm Capital" || platform === "Sahm2 Capital";
    const isAlrajhi = platform === "Alrajhi Capital" || platform === "Alrajhi Wallet";

    // Saudi Market: CMA (0.03%) + Tadawul tariffs + platform commission + 15% VAT
    if (isSaudi) {
      const cmaFee = round2(tv * 0.0003);
      const tadFees = round2(tv * 0.00009) + 2 * round2(tv * 0.00005) + round2(tv * 0.00001);
      const brkRate = isSahm ? 0 : (isAlrajhi ? 0.00105 : 0.001009);
      const vatableFees = tadFees + round2(tv * brkRate);
      const vat = round2(vatableFees * 0.15);
      return round2(cmaFee + vatableFees + vat) * sign;
    }

    // US Market: SEC fee + FINRA TAF + CAT fee + platform commission + local VAT
    if (isUs) {
      const isAbyan = (platform === "Abyan");
      let usBrokerFee = 0;
      if (isSahm) {
        const baseFee = p <= 5 ? 0.49 : 1.99;
        usBrokerFee = Math.max(baseFee, Math.min(q * 0.015, tv * 0.015));
      } else if (isAbyan) {
        usBrokerFee = Math.max(0.95, q * 0.0075);
      } else if (isAlrajhi) {
        usBrokerFee = Math.max(1.99, q * 0.0199);
      } else {
        return 0;
      }

      let regFees = 0;
      if (isSahm) {
        const settlement = Math.max(q * 0.003, 0.01);
        const catFee = Math.max(q * 0.000027, 0.01);
        const taf = isSell ? Math.min(Math.max(q * 0.000166, 0.01), 8.30) : 0;
        regFees = settlement + catFee + taf;
      } else {
        regFees = (isSell && !isAbyan) ? (tv * 0.0000278 + Math.min(Math.max(q * 0.000195, 0.01), 5.95)) : 0;
      }

      return round2(usBrokerFee * 1.15 + regFees) * sign;
    }

    return 0;
  }

  // =========================================================================
  // PARALLEL BATCH YAHOO FINANCE API CONNECTOR
  // =========================================================================

  /**
   * Retrieves auth cookie and crumb token for Yahoo Finance API endpoints (cached for 6 hours).
   * @private
   */
  static _getYahooCredentials() {
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
      Logger.log(`[Yahoo API Auth Warning] Could not fetch crumb: ${e.message}`);
    }

    return { cookie: "", crumb: "" };
  }

  /**
   * Converts local symbols to Yahoo Finance compatible ticker format.
   * @private
   */
  static _toYahooTicker(symbol, marketId) {
    const s = String(symbol ?? "").trim();
    if (!s) return "";

    const m = String(marketId ?? "").trim();
    if (s.toLowerCase().startsWith("gold") || s === "Gold.") {
      return "GC=F";
    }

    if (m === "702" || /^\d{4}$/.test(s)) {
      return `${s}.SR`;
    }

    return s;
  }

  /**
   * Fetches multiple symbols in batches of 50 in PARALLEL via UrlFetchApp.fetchAll().
   * @param {string[]} tickers
   * @returns {Map<string, { name: string, price: number, min52: number }>}
   */
  static fetchYahooQuotesBatch(tickers) {
    const results = new Map();
    if (!tickers || tickers.length === 0) return results;

    const uniqueTickers = [...new Set(tickers.filter(t => Boolean(t)))];
    const { cookie, crumb } = PortfolioManager._getYahooCredentials();
    const userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";
    const BATCH_SIZE = 50;

    const requests = [];
    for (let i = 0; i < uniqueTickers.length; i += BATCH_SIZE) {
      const chunk = uniqueTickers.slice(i, i + BATCH_SIZE);
      const symbolParam = chunk.map(encodeURIComponent).join(",");

      let url = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${symbolParam}`;
      if (crumb) url += `&crumb=${encodeURIComponent(crumb)}`;

      const headers = { "User-Agent": userAgent };
      if (cookie) headers["Cookie"] = cookie;

      requests.push({
        url: url,
        headers: headers,
        muteHttpExceptions: true
      });
    }

    if (requests.length === 0) return results;

    try {
      const responses = UrlFetchApp.fetchAll(requests);

      for (const response of responses) {
        if (response.getResponseCode() === 200) {
          const json = JSON.parse(response.getContentText());
          const quotes = json?.quoteResponse?.result || [];

          for (const q of quotes) {
            const returnedTicker = q.symbol;
            let currentPrice = q.regularMarketPrice ?? 0;
            let yearLow = q.fiftyTwoWeekLow ?? 0;
            const companyName = q.shortName || q.longName || "";

            if (returnedTicker === "GC=F") {
              const ozToGramSAR = 3.75 / 31.1034768;
              if (currentPrice > 0) currentPrice = Math.round(currentPrice * ozToGramSAR * 100) / 100;
              if (yearLow > 0) yearLow = Math.round(yearLow * ozToGramSAR * 100) / 100;
            }

            results.set(returnedTicker, {
              name: companyName,
              price: currentPrice,
              min52: yearLow
            });
          }
        } else {
          Logger.log(`[Yahoo API Batch Warning] Status ${response.getResponseCode()}`);
        }
      }
    } catch (err) {
      Logger.log(`[Yahoo API Parallel Batch Error] Failed to fetch quotes: ${err.message}`);
    }

    return results;
  }

  // =========================================================================
  // CORE SPREADSHEET PIPELINES (STATIC)
  // =========================================================================

  /**
   * Applies conditional formatting and table styles across targeted sheets.
   * @private
   */
  static _formatSheets(sheetNames, benchmark, spreadsheet = null) {
    if (typeof TableFormatter === "undefined" || !TableFormatter.extend) return;
    const ss = PortfolioManager.getSpreadsheet(spreadsheet);
    for (const name of sheetNames) {
      const sh = ss.getSheetByName(name);
      if (sh) {
        if (typeof Benchmark !== "undefined") {
          Benchmark.time("Optimize All Pipeline", `Extend Format [${name}]`, () => TableFormatter.extend(sh));
        } else {
          TableFormatter.extend(sh);
        }
      }
    }
  }

  /**
   * Persists an audit trace into the dedicated Optimize_Log sheet.
   * @private
   */
  static _writeLogSheet(logEntries, timestamp, spreadsheet = null) {
    const headers = ["Timestamp", "Sheet Name", "Action", "Order", "Platform", "Symbol", "Stock Name", "Order Type", "Details"];
    const ss = PortfolioManager.getSpreadsheet(spreadsheet);

    let logSheet = ss.getSheetByName("Optimize_Log");
    if (!logSheet) {
      logSheet = ss.insertSheet("Optimize_Log");
    }

    logSheet.clear();

    if (logEntries.length === 0) {
      logEntries.push([
        timestamp, "Portfolio", "UPDATE", "N/A", "N/A", "N/A", "N/A", "N/A",
        "No changes required. Portfolio is up to date."
      ]);
    }

    const fullData = [headers, ...logEntries];
    logSheet.getRange(1, 1, fullData.length, headers.length).setValues(fullData);
    logSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#3c78d8").setFontColor("#ffffff");
    logSheet.getRange(2, 1, logEntries.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
    logSheet.setFrozenRows(1);
    logSheet.setColumnWidths(1, 8, 125);
    logSheet.setColumnWidth(9, 420);
  }

  /**
   * Clears the 'Optimize Date' column in Transactions to force a baseline recalculation.
   * @private
   */
  static _clearTransactionOptimizeDates(spreadsheet = null) {
    const ss = PortfolioManager.getSpreadsheet(spreadsheet);
    const transSheet = ss.getSheetByName("Transactions");
    if (!transSheet) return;

    const headers = transSheet.getRange(2, 1, 1, transSheet.getLastColumn()).getValues()[0];
    const optDateIdx = headers.indexOf("Optimize Date");
    const lastRow = transSheet.getLastRow();

    if (optDateIdx !== -1 && lastRow > 2) {
      transSheet.getRange(3, optDateIdx + 1, lastRow - 2, 1).clearContent();
    }
  }

  /**
   * Executes a pipeline action with live status updates logged to System_Control.
   * Supports (actionFn), (actionName, actionFn), and explicit spreadsheet passing.
   */
  static runWithStatus(actionNameOrFn, maybeFn, spreadsheet = null) {
    let actionName = "Running";
    let actionFn;
    let ss = spreadsheet;

    if (typeof actionNameOrFn === "function") {
      actionFn = actionNameOrFn;
      if (maybeFn && typeof maybeFn.getSheetByName === "function") {
        ss = maybeFn;
      }
    } else {
      actionName = typeof actionNameOrFn === "string" ? actionNameOrFn : "Running";
      actionFn = maybeFn;
    }

    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheet = ss.getSheetByName("System_Control");
    if (!sheet || sheet.getLastColumn() === 0) {
      actionFn();
      return;
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const colMap = new Map();
    headerRow.forEach((h, i) => colMap.set(String(h).trim(), i + 1));

    const actionCol = colMap.get("Action");
    const codeCol   = colMap.get("Action Code");
    const stateCol  = colMap.get("State");
    const dateCol   = colMap.get("Last update");
    const errorCol  = colMap.get("Error");

    if (!actionCol || !stateCol || !dateCol || !errorCol) {
      actionFn();
      return;
    }

    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    const getTimestamp = () => Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    sheet.getRange(2, stateCol).setValue(actionName);
    sheet.getRange(2, dateCol).setValue(getTimestamp());
    sheet.getRange(2, errorCol).clearContent();
    SpreadsheetApp.flush();

    let errorOccurred = null;
    try {
      actionFn();
    } catch (e) {
      errorOccurred = e.message || String(e);
      Logger.log(`❌ Execution Error: ${e.stack || e.message}`);
      throw e;
    } finally {
      sheet.getRange(2, actionCol).setValue("- Select Action -");
      if (codeCol) sheet.getRange(2, codeCol).setValue("IDLE");
      sheet.getRange(2, stateCol).setValue(errorOccurred ? "Error" : "Completed");
      sheet.getRange(2, dateCol).setValue(getTimestamp());

      if (errorOccurred) {
        sheet.getRange(2, errorCol).setValue(errorOccurred);
      } else {
        sheet.getRange(2, errorCol).clearContent();
      }
      SpreadsheetApp.flush();
    }
  }

  /**
   * Optimizes the Google_Finance_Data sheet:
   * 1. Multi-tier sort: Market_ID first, then (No Google Min & No Yahoo Min), then natural alphanumeric GF_Key.
   * 2. Yahoo Finance parallel quotes backfill.
   * 3. Syncs incoming Yahoo Price/Min to Manual Price/Min (leaves manual fields untouched if API fails).
   * 4. Explicit timestamp formatting: Stamped Optimize_Date contains full time formatted as 'yyyy-mm-dd hh:mm:ss'.
   */
  static optimizeGoogleFinance(logEntries = null, spreadsheet = null) {
    const ss = PortfolioManager.getSpreadsheet(spreadsheet);
    const sheet = ss.getSheetByName("Google_Finance_Data");
    if (!sheet) return;

    let sortedData = null;
    let optDateColIdx = -1;
    const nowTimestamp = new Date();

    try {
      BatchProcessor.process(
        sheet,
        2,
        (row, col, index, rawData) => {
          if (index === 0) {
            const requireCol = (name) => {
              const idx = col(name);
              if (idx === -1) throw new Error(`[Google_Finance_Data] Missing canonical column: "${name}"`);
              return idx;
            };

            const keyCol          = requireCol("GF_Key");
            const orderCol        = requireCol("Sort_Order");
            const marketIdCol     = requireCol("Market_ID");
            const symbolCol       = requireCol("Symbol");
            const manualNameCol   = requireCol("Manual_Name");
            const combinedNameCol = requireCol("Combined_Name");
            const manualYearMinCol= requireCol("Manual_Year_Min");
            const googleMinCol    = requireCol("Google_Min");
            const yahooNameCol    = requireCol("Yahoo_Name");
            const yahooMinCol     = requireCol("Yahoo_Min");
            const manualPriceCol  = requireCol("Manual_Price");
            const googlePriceCol  = requireCol("Google_Price");
            const yahooPriceCol   = requireCol("Yahoo_Price");
            const lastUpdateCol   = requireCol("Last_Update");
            const optDateCol      = requireCol("Optimize_Date");

            optDateColIdx = optDateCol;

            const numRows = rawData.length;
            const decorated = new Array(numRows);
            const standaloneLogs = logEntries || [];

            const isPopulated = (val) => {
              if (val === "" || val === null || val === undefined) return false;
              if (typeof val === "string") {
                const s = val.trim().toUpperCase();
                if (s === "" || s.startsWith("#") || s === "N/A" || s === "ERROR!") return false;
              }
              const num = Number(val);
              return !Number.isNaN(num) && num > 0;
            };

            for (let i = 0; i < numRows; i++) {
              const r = rawData[i].slice();
              const marketId = String(r[marketIdCol] ?? "").trim();
              const symbol = String(r[symbolCol] ?? "").trim();
              const manualName = String(r[manualNameCol] ?? "").trim();
              const lastUpdateTime = PortfolioManager.getTime(r[lastUpdateCol]);
              const optDateTime = PortfolioManager.getTime(r[optDateCol]);
              const isDirty = (optDateTime === 0 || lastUpdateTime > optDateTime);

              const computedKey = (marketId && symbol) ? `${marketId} - ${symbol}` : String(r[keyCol] ?? "").trim();
              const computedCombinedName = (symbol && manualName) ? `${symbol} - ${manualName}` : (symbol || String(r[combinedNameCol] ?? "").trim());

              const keyChanged = (r[keyCol] !== computedKey);
              const nameChanged = (r[combinedNameCol] !== computedCombinedName);

              r[keyCol] = computedKey;
              r[combinedNameCol] = computedCombinedName;

              if (isDirty || keyChanged || nameChanged) {
                r[optDateCol] = nowTimestamp;
                if (keyChanged || nameChanged) {
                  standaloneLogs.push([
                    nowTimestamp,
                    "Google_Finance_Data",
                    "UPDATE",
                    r[orderCol] || (i + 1),
                    marketId || "N/A",
                    symbol || "N/A",
                    manualName || "N/A",
                    "Field Computation",
                    `Computed: GF_Key="${computedKey}", Combined_Name="${computedCombinedName}"`
                  ]);
                }
              }

              const hasNoMin = !isPopulated(r[googleMinCol]) && !isPopulated(r[yahooMinCol]);

              decorated[i] = {
                row: r,
                marketId: marketId,
                noMinRank: hasNoMin ? 0 : 1,
                key: computedKey,
                symbol: symbol,
                manualName: manualName,
                oldOrder: Number(r[orderCol]) || 0
              };
            }

            // Multi-Tier Sort: Market_ID -> hasNoMin -> GF_Key
            const collator = PortfolioManager._COLLATOR;
            decorated.sort((a, b) => {
              if (a.marketId !== b.marketId) {
                const mComp = collator.compare(a.marketId, b.marketId);
                if (mComp !== 0) return mComp;
              }
              if (a.noMinRank !== b.noMinRank) {
                return a.noMinRank - b.noMinRank;
              }
              return collator.compare(a.key, b.key);
            });

            sortedData = new Array(numRows);
            const needsYahooFetch = [];

            for (let i = 0; i < numRows; i++) {
              const item = decorated[i];
              const r = item.row;
              const newOrder = i + 1;
              const oldOrder = item.oldOrder;

              r[orderCol] = newOrder;
              sortedData[i] = r;

              if (oldOrder !== newOrder) {
                r[optDateCol] = nowTimestamp;
                standaloneLogs.push([
                  nowTimestamp,
                  "Google_Finance_Data",
                  "UPDATE",
                  newOrder,
                  item.marketId || "N/A",
                  item.symbol || "N/A",
                  item.manualName || "N/A",
                  "Sort Sequence",
                  `Sort_Order changed from ${oldOrder} to ${newOrder}`
                ]);
              }

              if (!isPopulated(r[googlePriceCol])) {
                const yahooTicker = PortfolioManager._toYahooTicker(item.symbol, item.marketId);
                if (yahooTicker) {
                  needsYahooFetch.push({
                    rowIndex: i,
                    ticker: yahooTicker,
                    symbol: item.symbol,
                    marketId: item.marketId
                  });
                }
              }
            }

            if (needsYahooFetch.length > 0) {
              const allYahooTickers = needsYahooFetch.map(item => item.ticker);
              const quotesMap = PortfolioManager.fetchYahooQuotesBatch(allYahooTickers);

              for (const target of needsYahooFetch) {
                const r = sortedData[target.rowIndex];
                const quote = quotesMap.get(target.ticker);

                if (quote && quote.price > 0) {
                  r[yahooNameCol] = quote.name || "";
                  r[yahooMinCol] = quote.min52 > 0 ? quote.min52 : "";
                  r[yahooPriceCol] = quote.price > 0 ? quote.price : "";

                  // Record incoming Yahoo values into Manual fields
                  r[manualPriceCol] = quote.price;
                  if (quote.min52 > 0) {
                    r[manualYearMinCol] = quote.min52;
                  }

                  r[optDateCol] = nowTimestamp;

                  standaloneLogs.push([
                    nowTimestamp,
                    "Google_Finance_Data",
                    "UPDATE",
                    r[orderCol],
                    target.marketId || "N/A",
                    target.symbol,
                    quote.name || "N/A",
                    "Yahoo API Fill",
                    `Yahoo backfill applied: Price=${quote.price}, 52WkLow=${quote.min52} (Manual Price/Min updated)`
                  ]);
                } else {
                  // Original behavior: clear Yahoo fields if quote is missing
                  r[yahooNameCol] = "";
                  r[yahooMinCol] = "";
                  r[yahooPriceCol] = "";
                  // Manual_Price and Manual_Year_Min remain untouched
                }
              }
            }

            if (!logEntries && standaloneLogs.length > 0) {
              PortfolioManager._writeLogSheet(standaloneLogs, nowTimestamp, ss);
            }
          }

          return [sortedData[index]];
        },
        PortfolioManager.getArrayFormulas("Google_Finance_Data")
      );

      // Explicitly enforce date-time number format across the Optimize_Date column
      if (optDateColIdx !== -1) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 3) {
          sheet.getRange(3, optDateColIdx + 1, lastRow - 2, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
        }
      }

      ss.toast("Google Finance Data optimized and synchronized.", "Success", 3);
    } catch (e) {
      Logger.log(`❌ optimizeGoogleFinance Error: ${e.message}`);
      ss.toast(e.message, "Optimization Error", 10);
      throw e;
    }
  }

  // =========================================================================
  // TARGETED TRANSACTIONS & STOCKS ENGINE (STATIC)
  // =========================================================================

  /**
   * Optimizes Transactions and synchronizes Stocks positions completely offline.
   * Strictly executes Stage 2 & Stage 3: does not perform network calls, formula audits, or visual formatting.
   * Safe to call directly from custom menu functions or internally from optimizeAll().
   * 
   * @param {Object} [options={}] - Execution options.
   * @param {boolean} [options.reset=false] - Resets transaction optimize dates to force baseline recalculation.
   * @param {any[][]} [options.logEntries=null] - Optional shared log accumulator when invoked by optimizeAll.
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [options.spreadsheet=null] - Target spreadsheet instance.
   */
  static optimizeTransactionsAndStocks(options = {}) {
    const isMenuOrEvent = Boolean(options && typeof options === "object" && ("authMode" in options || "range" in options));
    const safeOpts = (isMenuOrEvent || typeof options !== "object" || options === null) ? {} : options;

    const reset = safeOpts.reset ?? false;
    const logEntries = safeOpts.logEntries || null;
    const ss = PortfolioManager.getSpreadsheet(safeOpts.spreadsheet);

    const isStandalone = !logEntries;
    const currentLogs = logEntries || [];
    const nowTimestamp = new Date();

    const sessionName = isStandalone ? "Transactions & Stocks Pipeline" : "Optimize All Pipeline";
    if (isStandalone && typeof Benchmark !== "undefined") {
      Benchmark.start(sessionName);
    }
    const timeStep = (label, fn) => (typeof Benchmark !== "undefined" ? Benchmark.time(sessionName, label, fn) : fn());

    const execute = () => {
      if (reset) {
        PortfolioManager._clearTransactionOptimizeDates(ss);
      }

      const transSheet = ss.getSheetByName("Transactions");
      if (!transSheet) return;

      const dirtyKeys = new Set();
      const costDataMap = new Map();

      // Stage 2: Transactions Processing, Sorting & Dividend Auto-Replication
      timeStep("Process & Write Transactions Pass", () => {
        let sortedData = null;
        const newDividendRows = [];

        BatchProcessor.process(
          transSheet,
          2,
          (row, col, index, rawData) => {
            if (index === 0) {
              const requireCol = (name) => {
                const idx = col(name);
                if (idx === -1) throw new Error(`[Transactions] Missing canonical column: "${name}"`);
                return idx;
              };

              const serialIdx     = requireCol("Transaction_ID");
              const execDateIdx   = requireCol("Execution Date");
              const transDateIdx  = requireCol("Transaction Date");
              const companyIdx    = requireCol("Company");
              const orderTypeIdx  = requireCol("Order Type");
              const avgCostIdx    = requireCol("Avg Cost");
              const orderIdx      = requireCol("Order");
              const platformIdx   = requireCol("Platform");
              const symbolIdx     = requireCol("Symbol");
              const qtyIdx        = requireCol("Quantity");
              const priceIdx      = requireCol("Value");
              const exchangeIdx   = requireCol("Exchange");
              const feeIdx        = requireCol("Fee");
              const totalFeeIdx   = requireCol("Total Fee");
              const lastUpdateIdx = requireCol("Last Update");
              const optDateIdx    = requireCol("Optimize Date");

              const numRows = rawData.length;
              const decorated = new Array(numRows);
              const existingDivKeys = new Set();
              let maxSerial = 0;

              for (let i = 0; i < numRows; i++) {
                const r = rawData[i].slice();
                const sNum = Number(r[serialIdx]) || 0;
                if (sNum > maxSerial) maxSerial = sNum;

                const platform = String(r[platformIdx] || "").trim();
                const symbol = String(r[symbolIdx] || "").trim();
                const company = String(r[companyIdx] || "").trim();
                const rawType = String(r[orderTypeIdx] || "").trim();
                const keyLower = (platform && symbol) ? `${platform.toLowerCase()}|${symbol.toLowerCase()}` : "";
                const transDateStr = PortfolioManager.formatDateKey(r[transDateIdx]);

                if (keyLower && rawType === "Dividend" && transDateStr) {
                  existingDivKeys.add(`${keyLower}|${transDateStr}`);
                }

                if (keyLower) {
                  const modTime = PortfolioManager.getTime(r[lastUpdateIdx]);
                  const optTime = PortfolioManager.getTime(r[optDateIdx]);
                  if (modTime > optTime || !r[optDateIdx]) {
                    dirtyKeys.add(keyLower);
                  }
                }

                decorated[i] = {
                  row: r,
                  platform: platform,
                  symbol: symbol,
                  company: company,
                  bLower: platform.toLowerCase(),
                  sLower: symbol.toLowerCase(),
                  keyLower: keyLower,
                  exchange: r[exchangeIdx],
                  rawType: rawType,
                  type: rawType.toLowerCase(),
                  qty: Math.abs(parseFloat(r[qtyIdx])) || 0,
                  price: Math.abs(parseFloat(r[priceIdx])) || 0,
                  tExec: PortfolioManager.getTime(r[execDateIdx]),
                  tTrans: PortfolioManager.getTime(r[transDateIdx]),
                  tUpdate: PortfolioManager.getTime(r[lastUpdateIdx]),
                  transDateStr: transDateStr,
                  comp: company.toLowerCase(),
                  cost: Number(r[avgCostIdx]) || 0
                };
              }

              decorated.sort((a, b) => {
                if (a.tExec !== b.tExec) return a.tExec - b.tExec;
                if (a.tTrans !== b.tTrans) return a.tTrans - b.tTrans;
                if (a.comp !== b.comp) return a.comp < b.comp ? -1 : 1;
                if (a.type !== b.type) return a.type < b.type ? -1 : 1;
                if (a.tUpdate !== b.tUpdate) return a.tUpdate - b.tUpdate;
                return b.cost - a.cost;
              });

              const portfolioMap = new Map();
              const symbolPlatformsMap = new Map();
              const todayLimit = new Date().setHours(23, 59, 59, 999);
              const todayStartMs = new Date().setHours(0, 0, 0, 0);

              sortedData = new Array(numRows);

              for (let i = 0; i < numRows; i++) {
                const item = decorated[i];
                const rowItem = item.row;

                const newOrder = i + 1;
                const oldOrder = Number(rowItem[orderIdx]) || newOrder;
                rowItem[orderIdx] = newOrder;

                if (oldOrder !== newOrder) {
                  currentLogs.push([
                    nowTimestamp,
                    "Transactions",
                    "UPDATE",
                    newOrder,
                    item.platform || "N/A",
                    item.symbol || "N/A",
                    item.company || "N/A",
                    item.rawType || "N/A",
                    `Order changed from ${oldOrder} to ${newOrder}`
                  ]);
                }

                const keyLower = item.keyLower;
                if (!keyLower) {
                  sortedData[i] = rowItem;
                  continue;
                }

                let p = portfolioMap.get(keyLower);
                if (!p) {
                  p = { qty: 0, cost: 0, avg: 0 };
                  portfolioMap.set(keyLower, p);
                }

                const totalFee = PortfolioManager.calculateFee(item.exchange, item.platform, item.rawType, item.qty, item.price);
                const singleFee = item.qty > 0 ? Math.round((totalFee / item.qty) * 10000) / 10000 : 0;
                const totalValFee = Math.round(((item.qty * item.price) + totalFee) * 100) / 100;
                const isExecuted = (item.tExec <= todayLimit);

                if (item.rawType === "Buy" || item.rawType === "Gift") {
                  if (isExecuted) {
                    if (p.qty <= 0.000001) {
                      p.qty = 0;
                      p.cost = 0;
                    }
                    p.qty += item.qty;
                    p.cost += Math.abs(totalValFee);
                    if (p.qty > 0) p.avg = p.cost / p.qty;
                  }

                  const cost = Math.round(p.avg * 10000) / 10000;
                  let mapItem = costDataMap.get(keyLower);
                  if (!mapItem) {
                    mapItem = {
                      normalAvgCost: null,
                      futureAvgCost: null,
                      origPlatform: item.platform,
                      origSymbol: item.symbol,
                      origCompany: item.company,
                      qty: p.qty
                    };
                    costDataMap.set(keyLower, mapItem);
                  }
                  mapItem.futureAvgCost = cost;
                  if (isExecuted) mapItem.normalAvgCost = cost;
                  mapItem.qty = p.qty;

                } else if (item.rawType === "Sell") {
                  if (isExecuted) {
                    if (p.qty > 0) {
                      const sellQty = Math.min(item.qty, p.qty);
                      p.cost -= (sellQty * p.avg);
                      p.qty -= item.qty;
                    }
                    if (p.qty <= 0.000001) {
                      p.qty = 0;
                      p.cost = 0;
                    }
                    let mapItem = costDataMap.get(keyLower);
                    if (mapItem) mapItem.qty = p.qty;
                  }

                } else if (item.rawType === "Dividend") {
                  if (item.tTrans > todayLimit) {
                    const targetQty = Math.max(0, p.qty);
                    const oldQty = Number(rowItem[qtyIdx]) || 0;
                    if (oldQty !== targetQty) {
                      rowItem[qtyIdx] = targetQty;
                      rowItem[lastUpdateIdx] = nowTimestamp;
                      currentLogs.push([
                        nowTimestamp,
                        "Transactions",
                        "UPDATE",
                        newOrder,
                        item.platform,
                        item.symbol,
                        item.company,
                        "Dividend",
                        `Updated Dividend Quantity from ${oldQty} to ${targetQty}`
                      ]);
                    }
                  }

                  if (item.tTrans >= todayStartMs && item.symbol && item.transDateStr) {
                    const targetPlatforms = symbolPlatformsMap.get(item.sLower);
                    if (targetPlatforms) {
                      for (const targetPlatformName of targetPlatforms) {
                        if (targetPlatformName.toLowerCase() === item.bLower) continue;

                        const targetKeyLower = `${targetPlatformName.toLowerCase()}|${item.sLower}`;
                        const targetP = portfolioMap.get(targetKeyLower);

                        if (targetP && targetP.qty > 0) {
                          const targetDivKey = `${targetKeyLower}|${item.transDateStr}`;
                          if (!existingDivKeys.has(targetDivKey)) {
                            existingDivKeys.add(targetDivKey);
                            maxSerial++;

                            const cloneRow = rowItem.slice();
                            const repOrderNum = numRows + newDividendRows.length + 1;

                            cloneRow[serialIdx] = maxSerial;
                            cloneRow[orderIdx] = repOrderNum;
                            cloneRow[platformIdx] = targetPlatformName;
                            cloneRow[qtyIdx] = targetP.qty;

                            const targetTotalFee = PortfolioManager.calculateFee(item.exchange, targetPlatformName, item.rawType, targetP.qty, item.price);
                            const targetSingleFee = targetP.qty > 0 ? Math.round((targetTotalFee / targetP.qty) * 10000) / 10000 : 0;

                            cloneRow[feeIdx] = targetSingleFee;
                            cloneRow[totalFeeIdx] = targetTotalFee;
                            cloneRow[avgCostIdx] = Math.round(targetP.avg * 10000) / 10000;
                            cloneRow[lastUpdateIdx] = nowTimestamp;
                            cloneRow[optDateIdx] = nowTimestamp;

                            newDividendRows.push(cloneRow);
                            dirtyKeys.add(targetKeyLower);

                            currentLogs.push([
                              nowTimestamp,
                              "Transactions",
                              "INSERT",
                              repOrderNum,
                              targetPlatformName,
                              item.symbol,
                              item.company,
                              "Dividend",
                              `Replicated Dividend created (Transaction_ID: ${maxSerial}, Qty: ${targetP.qty})`
                            ]);
                          }
                        }
                      }
                    }
                  }
                }

                let bSet = symbolPlatformsMap.get(item.sLower);
                if (!bSet) {
                  bSet = new Set();
                  symbolPlatformsMap.set(item.sLower, bSet);
                }
                if (p.qty > 0) bSet.add(item.platform);
                else bSet.delete(item.platform);

                const oldFee = Number(rowItem[feeIdx]) || 0;
                const oldTotalFee = Number(rowItem[totalFeeIdx]) || 0;
                const oldAvg = item.cost;
                const oldOptDate = rowItem[optDateIdx];

                rowItem[feeIdx] = singleFee;
                rowItem[totalFeeIdx] = totalFee;

                const newAvg = Math.round(p.avg * 10000) / 10000;
                rowItem[avgCostIdx] = newAvg;

                const feeChanged = Math.abs(oldFee - singleFee) > 0.0001;
                const totalFeeChanged = Math.abs(oldTotalFee - totalFee) > 0.01;
                const avgChanged = Math.abs(oldAvg - newAvg) > 0.0001;
                const optDateMissing = !oldOptDate;

                if (dirtyKeys.has(keyLower) || feeChanged || totalFeeChanged || avgChanged || optDateMissing) {
                  dirtyKeys.add(keyLower);
                  rowItem[optDateIdx] = nowTimestamp;

                  if (feeChanged || totalFeeChanged || avgChanged) {
                    rowItem[lastUpdateIdx] = nowTimestamp;
                    currentLogs.push([
                      nowTimestamp,
                      "Transactions",
                      "UPDATE",
                      newOrder,
                      item.platform,
                      item.symbol,
                      item.company,
                      item.rawType,
                      `Transaction recalculated (Avg Cost: ${oldAvg} -> ${newAvg})`
                    ]);
                  }
                }

                sortedData[i] = rowItem;
              }
            }

            if (index === rawData.length - 1 && newDividendRows.length > 0) {
              return [sortedData[index], ...newDividendRows];
            }

            return [sortedData[index]];
          },
          PortfolioManager.getArrayFormulas("Transactions")
        );
      });

      // Stage 3: Stocks Position Synchronization
      const stocksSheet = ss.getSheetByName("Stocks");
      if (stocksSheet) {
        timeStep("Process & Write Stocks Pass", () => {
          let matchedKeys = new Set();
          let sLastCol = 0;
          let idxPlatform = -1, idxSymbol = -1, idxCompany = -1, idxCost = -1, lastUpdateIdx = -1;
          const formulaIndices = [];

          BatchProcessor.process(
            stocksSheet,
            2,
            (row, col, index, rawData) => {
              if (index === 0) {
                const requireCol = (name) => {
                  const idx = col(name);
                  if (idx === -1) throw new Error(`[Stocks] Missing canonical column: "${name}"`);
                  return idx;
                };

                idxPlatform   = requireCol("Platform");
                idxSymbol     = requireCol("Symbol");
                idxCompany    = requireCol("Company");
                idxCost       = requireCol("Avg Cost");
                lastUpdateIdx = requireCol("Last Update");
                sLastCol      = row.length;

                const formulaHeaders = PortfolioManager.getArrayFormulas("Stocks");
                formulaHeaders.forEach(fh => {
                  const fIdx = col(fh);
                  if (fIdx !== -1) formulaIndices.push(fIdx);
                });
              }

              const platform = String(row[idxPlatform] || "").trim();
              const symbol = String(row[idxSymbol] || "").trim();
              const company = String(row[idxCompany] || "").trim();
              const keyLower = (platform && symbol) ? `${platform.toLowerCase()}|${symbol.toLowerCase()}` : null;

              if (keyLower) {
                matchedKeys.add(keyLower);

                if (dirtyKeys.has(keyLower)) {
                  const item = costDataMap.get(keyLower);
                  if (item) {
                    const targetRaw = PortfolioManager.getEffectiveCost(item);
                    const target2pt = Math.round(targetRaw * 100) / 100;
                    const currentCost = Math.round((Number(row[idxCost]) || 0) * 100) / 100;

                    row[idxPlatform] = item.origPlatform;
                    row[idxSymbol] = item.origSymbol;

                    if (Math.abs(currentCost - target2pt) > 0.011) {
                      row[idxCost] = targetRaw;
                      row[lastUpdateIdx] = nowTimestamp;

                      currentLogs.push([
                        nowTimestamp,
                        "Stocks",
                        "UPDATE",
                        "N/A",
                        item.origPlatform,
                        item.origSymbol,
                        item.origCompany || company,
                        "N/A",
                        `Avg Cost updated from ${currentCost} to ${target2pt}`
                      ]);
                    }
                  }
                }
              }

              if (index === rawData.length - 1) {
                const newStockRows = [];
                for (const [k, item] of costDataMap.entries()) {
                  if (!matchedKeys.has(k)) {
                    const finalCost = PortfolioManager.getEffectiveCost(item);
                    const newRow = new Array(sLastCol).fill("");

                    newRow[idxPlatform] = item.origPlatform;
                    newRow[idxSymbol] = item.origSymbol;
                    newRow[idxCompany] = item.origCompany || "";
                    newRow[idxCost] = Math.round(finalCost * 10000) / 10000;
                    newRow[lastUpdateIdx] = nowTimestamp;

                    formulaIndices.forEach(idx => {
                      if (idx >= 0 && idx < sLastCol) newRow[idx] = null;
                    });

                    newStockRows.push(newRow);

                    currentLogs.push([
                      nowTimestamp,
                      "Stocks",
                      "INSERT",
                      "N/A",
                      item.origPlatform,
                      item.origSymbol,
                      item.origCompany || "",
                      "N/A",
                      `New stock position added with Avg Cost: ${Math.round(finalCost * 10000) / 10000}`
                    ]);
                  }
                }
                return [row, ...newStockRows];
              }

              return [row];
            },
            PortfolioManager.getArrayFormulas("Stocks")
          );
        });
      }

      // Standalone completions: Log commitment & Toast notification only (No audits, No formatting)
      if (isStandalone) {
        timeStep("Write Optimize Log Pass", () => {
          PortfolioManager._writeLogSheet(currentLogs, nowTimestamp, ss);
        });

        const durationLabel = (typeof Benchmark !== "undefined")
          ? Benchmark.report(sessionName)?.totalDurationSec
          : "few seconds";
        ss.toast(`Transactions & Stocks optimized in ${durationLabel}.`, "Success", 5);
      }
    };

    if (isStandalone) {
      const statusTitle = reset ? "Recalculating Transactions & Stocks" : "Optimizing Transactions & Stocks";
      PortfolioManager.runWithStatus(statusTitle, execute, ss);
    } else {
      execute();
    }
  }

  // =========================================================================
  // UNIFIED MASTER PIPELINE (STATIC)
  // =========================================================================

  /**
   * UNIFIED MASTER OPTIMIZATION PIPELINE:
   * Executes full 6-stage portfolio synchronization across all sheets.
   */
  static optimizeAll(options = {}) {
    const isMenuOrEvent = Boolean(options && typeof options === "object" && ("authMode" in options || "range" in options));
    const safeOpts = (isMenuOrEvent || typeof options !== "object" || options === null) ? {} : options;

    const reset = safeOpts.reset ?? false;
    const runAudits = safeOpts.runAudits ?? true;
    const format = safeOpts.format ?? true;
    const ss = PortfolioManager.getSpreadsheet(safeOpts.spreadsheet);

    PortfolioManager.runWithStatus(reset ? "Recalculating All" : "Optimizing All", () => {
      const sessionName = "Optimize All Pipeline";
      if (typeof Benchmark !== "undefined") {
        Benchmark.start(sessionName);
      }
      const timeStep = (label, fn) => (typeof Benchmark !== "undefined" ? Benchmark.time(sessionName, label, fn) : fn());

      if (reset) {
        PortfolioManager._clearTransactionOptimizeDates(ss);
      }

      const nowTimestamp = new Date();
      const logEntries = [];

      // Stage 1: Google Finance Data Optimization (Parallel Yahoo API Fill & Timestamp Format)
      timeStep("Optimize Google Finance Data", () => PortfolioManager.optimizeGoogleFinance(logEntries, ss));

      // Stage 2 & 3: Transactions & Stocks Processing (Direct Delegation - Zero Code Duplication)
      PortfolioManager.optimizeTransactionsAndStocks({
        reset: false,
        logEntries: logEntries,
        spreadsheet: ss
      });

      // Stage 4: Synchronize Formula Audits (Exclusive to optimizeAll)
      if (runAudits) {
        timeStep("Synchronize Audits", () => {
          if (typeof runAllAudits === "function") {
            runAllAudits();
          } else if (typeof WorkbookAuditor !== "undefined" && typeof WorkbookAuditor.runAllAudits === "function") {
            WorkbookAuditor.runAllAudits(ss);
          } else {
            if (typeof listAllTables === "function") listAllTables();
            if (typeof listAllFormulas === "function") listAllFormulas();
          }
        });
      }

      // Stage 5: Commit Audit Log
      timeStep("Write Optimize Log Pass", () => {
        PortfolioManager._writeLogSheet(logEntries, nowTimestamp, ss);
      });

      // Stage 6: Visual Formatting (Exclusive to optimizeAll)
      if (format) {
        PortfolioManager._formatSheets(["Google_Finance_Data", "Transactions", "Stocks"], null, ss);
      }

      const durationLabel = (typeof Benchmark !== "undefined")
        ? Benchmark.report(sessionName)?.totalDurationSec
        : "few seconds";
      ss.toast(`Optimization complete in ${durationLabel}.`, "Optimize All", 5);
    }, ss);
  }

  /**
   * One-line delegate for full recalculations.
   */
  static resetAndRecalculateAll(options = {}) {
    const isMenuOrEvent = Boolean(options && typeof options === "object" && ("authMode" in options || "range" in options));
    const safeOpts = (isMenuOrEvent || typeof options !== "object" || options === null) ? {} : options;
    PortfolioManager.optimizeAll({ ...safeOpts, reset: true });
  }
}