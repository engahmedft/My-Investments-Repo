/**
 * Core orchestrator and calculation engine for investment portfolio operations.
 * STRICT CANONICAL EDITION: Zero aliases, zero workarounds. All column names match
 * the Google Sheet headers character-for-character.
 */
class PortfolioManager {

  /**
   * Static Pre-compiled Date Regex for fast O(1) date key parsing.
   * @private
   */
  static get _DATE_REGEX() {
    if (!PortfolioManager._DATE_REGEX_CACHE) {
      PortfolioManager._DATE_REGEX_CACHE = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/;
    }
    return PortfolioManager._DATE_REGEX_CACHE;
  }

  /**
   * Strict Canonical List of ArrayFormula Headers per Sheet.
   * Nullified during write-backs to protect live sheet formulas from being overwritten.
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

  static getArrayFormulas(sheetName) {
    return PortfolioManager.ARRAY_FORMULAS[sheetName] || [];
  }

  static _pad2(n) {
    return n < 10 ? '0' + n : String(n);
  }

  static _fmtDate(d) {
    return `${d.getFullYear()}-${PortfolioManager._pad2(d.getMonth() + 1)}-${PortfolioManager._pad2(d.getDate())}`;
  }

  static getTime(val) {
    if (!val) return 0;
    if (typeof val === "number") return val;
    if (val instanceof Date) return val.getTime();
    const d = new Date(val);
    return isNaN(d.getTime()) ? 0 : d.getTime();
  }

  static formatDateKey(val) {
    if (!val) return "";
    if (val instanceof Date) {
      return isNaN(val.getTime()) ? "" : PortfolioManager._fmtDate(val);
    }
    const str = String(val).trim();
    const match = str.match(PortfolioManager._DATE_REGEX);
    if (match) {
      return `${match[1]}-${PortfolioManager._pad2(match[2])}-${PortfolioManager._pad2(match[3])}`;
    }
    const d = new Date(str);
    return isNaN(d.getTime()) ? "" : PortfolioManager._fmtDate(d);
  }

  static getEffectiveCost(item) {
    if (!item) return 0;
    return item.normalAvgCost !== null ? item.normalAvgCost : item.futureAvgCost;
  }

  /**
   * Dynamic Transaction Fee Engine for Saudi Arabia and US markets.
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

    // Dividend Tax: US 30% withholding tax, SA 0%
    if (isDiv) {
      if (isSaudi) return 0;
      if (isUs) return round2(tv * 0.3) * sign;
      return 0;
    }

    const isSahm = platform === "Sahm Capital" || platform === "Sahm2 Capital";
    const isAlrajhi = platform === "Alrajhi Capital" || platform === "Alrajhi Wallet";

    // Saudi Market: Tadawul + CMA + platform commission + 15% VAT
    if (isSaudi) {
      const cmaFee = round2(tv * 0.0003);
      const tadFees = round2(tv * 0.00009) + 2 * round2(tv * 0.00005) + round2(tv * 0.00001);
      const brkRate = isSahm ? 0 : (isAlrajhi ? 0.00105 : 0.001009);
      const vatableFees = tadFees + round2(tv * brkRate);
      const vat = round2(vatableFees * 0.15);
      return round2(cmaFee + vatableFees + vat) * sign;
    }

    // US Market: SEC + FINRA TAF + CAT + platform commission + VAT
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
  // HIGH-PERFORMANCE BATCH YAHOO FINANCE API CONNECTOR
  // =========================================================================

  /**
   * Retrieves cookie and crumb for Yahoo Finance API v7/v10 requests.
   * Caches in CacheService for 6 hours to minimize setup overhead.
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
        if (Array.isArray(setCookie)) {
          cookie = setCookie.map(c => c.split(";")[0]).join("; ");
        } else {
          cookie = String(setCookie).split(";")[0];
        }
      }

      let crumb = "";
      if (cookie) {
        const resp2 = UrlFetchApp.fetch("https://query2.finance.yahoo.com/v1/test/getcrumb", {
          muteHttpExceptions: true,
          headers: {
            "User-Agent": userAgent,
            "Cookie": cookie
          }
        });
        if (resp2.getResponseCode() === 200) {
          crumb = resp2.getContentText().trim();
        }
      }

      if (cookie && crumb) {
        cache.put("YF_COOKIE", cookie, 21600); // 6 hours
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
   * Handles Saudi Tadawul (.SR), US stocks, and Gold futures (GC=F).
   * @private
   */
  static _toYahooTicker(symbol, marketId) {
    const s = String(symbol ?? "").trim();
    if (!s) return "";

    const m = String(marketId ?? "").trim();
    if (s.toLowerCase().startsWith("gold") || s === "Gold.") {
      return "GC=F"; // Gold Futures
    }

    // Saudi Market ID (702) or numeric stock codes
    if (m === "702" || /^\d{4}$/.test(s)) {
      return `${s}.SR`;
    }

    return s;
  }

  /**
   * Fetches multiple symbols in batches from Yahoo Finance (max 50 per batch).
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

    for (let i = 0; i < uniqueTickers.length; i += BATCH_SIZE) {
      const chunk = uniqueTickers.slice(i, i + BATCH_SIZE);
      const symbolParam = chunk.map(encodeURIComponent).join(",");

      let url = `https://query1.finance.yahoo.com/v7/finance/quote?symbols=${symbolParam}`;
      if (crumb) url += `&crumb=${encodeURIComponent(crumb)}`;

      const headers = { "User-Agent": userAgent };
      if (cookie) headers["Cookie"] = cookie;

      try {
        const response = UrlFetchApp.fetch(url, {
          muteHttpExceptions: true,
          headers: headers
        });

        if (response.getResponseCode() === 200) {
          const json = JSON.parse(response.getContentText());
          const quotes = json?.quoteResponse?.result || [];

          for (const q of quotes) {
            const returnedTicker = q.symbol;
            let currentPrice = q.regularMarketPrice ?? 0;
            let yearLow = q.fiftyTwoWeekLow ?? 0;
            const companyName = q.shortName || q.longName || "";

            // Convert Gold from USD/oz to SAR/gram
            // 1 Troy Ounce = 31.1034768 grams, 1 USD = 3.75 SAR
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
          Logger.log(`[Yahoo API Batch Warning] Status ${response.getResponseCode()} for chunk: ${symbolParam}`);
        }
      } catch (err) {
        Logger.log(`[Yahoo API Batch Error] Failed to fetch chunk: ${err.message}`);
      }
    }

    return results;
  }

  // =========================================================================
  // CORE SPREADSHEET PIPELINES
  // =========================================================================

  constructor(spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) {
    this.ss = spreadsheet;
  }

  getSheet(name) {
    return this.ss.getSheetByName(name);
  }

  _formatSheets(sheetNames, benchmark) {
    for (const name of sheetNames) {
      const sh = this.getSheet(name);
      if (sh) {
        if (benchmark) {
          benchmark.time(`Extend Format [${name}]`, () => TableFormatter.extend(sh));
        } else {
          TableFormatter.extend(sh);
        }
      }
    }
  }

  _validateHeaders(sheet, headerRowIndex) {
    if (!sheet || sheet.getLastColumn() === 0) return false;
    const headerValues = sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nonEntries = headerValues.filter(h => String(h || "").trim() !== "").length;
    if (nonEntries < 3) {
      Logger.log(`⚠️ Header validation failed for sheet "${sheet.getName()}": Header row empty.`);
      this.ss.toast(`Header row in sheet "${sheet.getName()}" is empty.`, "Header Protection", 8);
      return false;
    }
    return true;
  }

  _writeLogSheet(logEntries, timestamp) {
    const headers = ["Timestamp", "Sheet Name", "Action", "Order", "Platform", "Symbol", "Stock Name", "Order Type", "Details"];

    let logSheet = this.getSheet("Optimize_Log");
    if (!logSheet) {
      logSheet = this.ss.insertSheet("Optimize_Log");
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

  runWithStatus(actionFn) {
    const sheet = this.getSheet("System_Control");
    if (!sheet || !this._validateHeaders(sheet, 1)) {
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
      throw new Error(`[System_Control] Canonical header missing. Required: Action, State, Last update, Error.`);
    }

    const tz = this.ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    const getTimestamp = () => Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    sheet.getRange(2, stateCol).setValue("Running");
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
   * Optimized Google_Finance_Data Processor:
   * 1. Sorts and re-sequences Sort_Order.
   * 2. After sorting, queries Yahoo Finance batch API ONLY for records missing Google_Price.
   * 3. Converts Gold to SAR/gram.
   * 4. Writes all updates back in 1 atomic batch write.
   */
  optimizeGoogleFinance(logEntries = null) {
    const sheet = this.getSheet("Google_Finance_Data");
    if (!sheet || !this._validateHeaders(sheet, 2)) return;

    let sortedData = null;
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
            const googleMinCol    = requireCol("Google_Min");
            const yahooNameCol    = requireCol("Yahoo_Name");
            const yahooMinCol     = requireCol("Yahoo_Min");
            const googlePriceCol  = requireCol("Google_Price");
            const yahooPriceCol   = requireCol("Yahoo_Price");
            const lastUpdateCol   = requireCol("Last_Update");
            const optDateCol      = requireCol("Optimize_Date");

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
              return !isNaN(num) && num > 0;
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

              const hasAutoFeed = (
                isPopulated(r[googleMinCol]) ||
                isPopulated(r[googlePriceCol]) ||
                isPopulated(r[yahooMinCol]) ||
                isPopulated(r[yahooPriceCol])
              );

              decorated[i] = {
                row: r,
                manualRank: hasAutoFeed ? 1 : 0,
                key: computedKey,
                symbol: symbol,
                manualName: manualName,
                marketId: marketId,
                oldOrder: Number(r[orderCol]) || 0
              };
            }

            // Two-Tier Sort: Manual first, then Natural Alphanumeric by GF_Key
            decorated.sort((a, b) => {
              if (a.manualRank !== b.manualRank) return a.manualRank - b.manualRank;
              return a.key.localeCompare(b.key, undefined, { numeric: true, sensitivity: 'base' });
            });

            // Re-sequence Sort_Order
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

              // Check if Yahoo Finance backfill is required (No valid Google_Price)
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

            // POST-SORT YAHOO BATCH ENRICHMENT
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
                    `Yahoo backfill applied: Price=${quote.price}, 52WkLow=${quote.min52}`
                  ]);
                } else {
                  // Safe fallback: API failed or delisted -> write empty
                  r[yahooNameCol] = "";
                  r[yahooMinCol] = "";
                  r[yahooPriceCol] = "";
                }
              }
            }

            if (!logEntries && standaloneLogs.length > 0) {
              this._writeLogSheet(standaloneLogs, nowTimestamp);
            }
          }

          return [sortedData[index]];
        },
        PortfolioManager.getArrayFormulas("Google_Finance_Data")
      );

      this.ss.toast("Google Finance Data optimized, sorted, and enriched.", "Success", 3);
    } catch (e) {
      Logger.log(`❌ optimizeGoogleFinance Error: ${e.message}`);
      this.ss.toast(e.message, "Optimization Error", 10);
      throw e;
    }
  }

  sortGoogleFinance(logEntries = null) {
    return this.optimizeGoogleFinance(logEntries);
  }

  sortTransactions(logEntries = null) {
    const sheet = this.getSheet("Transactions");
    if (!sheet || !this._validateHeaders(sheet, 2)) return;

    let sortedData = null;
    const nowTimestamp = new Date();

    try {
      BatchProcessor.process(
        sheet,
        2,
        (row, col, index, rawData) => {
          if (index === 0) {
            const requireCol = (name) => {
              const idx = col(name);
              if (idx === -1) throw new Error(`[Transactions] Missing canonical column: "${name}"`);
              return idx;
            };

            const execDateIdx   = requireCol("Execution Date");
            const transDateIdx  = requireCol("Transaction Date");
            const companyIdx    = requireCol("Company");
            const orderTypeIdx  = requireCol("Order Type");
            const avgCostIdx    = requireCol("Avg Cost");
            const orderIdx      = requireCol("Order");
            const platformIdx   = requireCol("Platform");
            const symbolIdx     = requireCol("Symbol");
            const lastUpdateIdx = requireCol("Last Update");

            const numRows = rawData.length;
            const decorated = new Array(numRows);

            for (let i = 0; i < numRows; i++) {
              const r = rawData[i].slice();
              decorated[i] = {
                row: r,
                tExec: PortfolioManager.getTime(r[execDateIdx]),
                tTrans: PortfolioManager.getTime(r[transDateIdx]),
                tUpdate: PortfolioManager.getTime(r[lastUpdateIdx]),
                comp: String(r[companyIdx] || "").toLowerCase(),
                type: String(r[orderTypeIdx] || "").toLowerCase(),
                cost: Number(r[avgCostIdx]) || 0
              };
            }

            decorated.sort((a, b) => {
              if (a.tExec !== b.tExec) return a.tExec - b.tExec;
              if (a.tTrans !== b.tTrans) return a.tTrans - b.tTrans;
              const cComp = a.comp.localeCompare(b.comp);
              if (cComp !== 0) return cComp;
              const cType = a.type.localeCompare(b.type);
              if (cType !== 0) return cType;
              if (a.tUpdate !== b.tUpdate) return a.tUpdate - b.tUpdate;
              return b.cost - a.cost;
            });

            sortedData = new Array(numRows);
            const standaloneLogs = logEntries || [];

            for (let i = 0; i < numRows; i++) {
              const r = decorated[i].row;
              const newOrder = i + 1;
              const oldOrder = Number(r[orderIdx]) || newOrder;

              r[orderIdx] = newOrder;
              sortedData[i] = r;

              if (oldOrder !== newOrder) {
                standaloneLogs.push([
                  nowTimestamp,
                  "Transactions",
                  "UPDATE",
                  newOrder,
                  String(r[platformIdx] || "").trim() || "N/A",
                  String(r[symbolIdx] || "").trim() || "N/A",
                  String(r[companyIdx] || "").trim() || "N/A",
                  String(r[orderTypeIdx] || "").trim() || "N/A",
                  `Order changed from ${oldOrder} to ${newOrder}`
                ]);
              }
            }

            if (!logEntries && standaloneLogs.length > 0) {
              this._writeLogSheet(standaloneLogs, nowTimestamp);
            }
          }

          return [sortedData[index]];
        },
        PortfolioManager.getArrayFormulas("Transactions")
      );

      this.ss.toast("Transactions sorted and re-sequenced.", "Success", 3);
    } catch (e) {
      this.ss.toast(e.message, "Sort Error", 10);
      throw e;
    }
  }

  _optimizeAllCore() {
    const benchmark = new Benchmark("Optimize All Pipeline");
    const dirtyKeys = new Set();
    const costDataMap = new Map();
    const nowTimestamp = new Date();
    const logEntries = [];

    // Stage 1: Google Finance Optimization (includes Post-Sort Yahoo API Fill)
    benchmark.time("Optimize Google Finance Data", () => this.optimizeGoogleFinance(logEntries));

    const transSheet = this.getSheet("Transactions");
    if (!transSheet || !this._validateHeaders(transSheet, 2)) return;

    // Stage 2: Transactions Processing
    benchmark.time("Process & Write Transactions Pass", () => {
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
              const cComp = a.comp.localeCompare(b.comp);
              if (cComp !== 0) return cComp;
              const cType = a.type.localeCompare(b.type);
              if (cType !== 0) return cType;
              if (a.tUpdate !== b.tUpdate) return a.tUpdate - b.tUpdate;
              return b.cost - a.cost;
            });

            const portfolioMap = {};
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
                logEntries.push([
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

              let p = portfolioMap[keyLower];
              if (!p) {
                p = { qty: 0, cost: 0, avg: 0 };
                portfolioMap[keyLower] = p;
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
                    logEntries.push([
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

                // Replicate across holding platforms on Deserved Date
                if (item.tTrans >= todayStartMs && item.symbol && item.transDateStr) {
                  const targetPlatforms = symbolPlatformsMap.get(item.sLower);
                  if (targetPlatforms) {
                    for (const targetPlatformName of targetPlatforms) {
                      if (targetPlatformName.toLowerCase() === item.bLower) continue;

                      const targetKeyLower = `${targetPlatformName.toLowerCase()}|${item.sLower}`;
                      const targetP = portfolioMap[targetKeyLower];

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

                          logEntries.push([
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
                  logEntries.push([
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

    // Stage 3: Stocks Processing
    const stocksSheet = this.getSheet("Stocks");
    if (stocksSheet && this._validateHeaders(stocksSheet, 2)) {
      benchmark.time("Process & Write Stocks Pass", () => {
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

                    logEntries.push([
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

                  logEntries.push([
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

    // Stage 4: Synchronize Audits
    benchmark.time("Synchronize Audits", () => {
      if (typeof runAllAudits === "function") {
        runAllAudits();
      } else {
        if (typeof listAllTables === "function") listAllTables();
        if (typeof listAllFormulas === "function") listAllFormulas();
      }
    });

    // Stage 5: Write Audit Log
    benchmark.time("Write Optimize Log Pass", () => {
      this._writeLogSheet(logEntries, nowTimestamp);
    });

    // Stage 6: Formatting
    this._formatSheets(["Google_Finance_Data", "Transactions", "Stocks"], benchmark);

    const reportMetrics = benchmark.report();
    const durationLabel = reportMetrics?.totalDurationSec || "few seconds";
    this.ss.toast(`Optimizations complete in ${durationLabel}.`, "Optimize All", 5);
  }

  _resetAndRecalculateAllCore() {
    const transSheet = this.getSheet("Transactions");
    if (!transSheet) {
      this.ss.toast("Transactions sheet not found.", "Error", 5);
      return;
    }

    const headers = transSheet.getRange(2, 1, 1, transSheet.getLastColumn()).getValues()[0];
    const optDateIdx = headers.indexOf("Optimize Date");
    if (optDateIdx === -1) {
      this.ss.toast("Column 'Optimize Date' missing in Transactions sheet.", "Error", 5);
      return;
    }

    const lastRow = transSheet.getLastRow();
    if (lastRow > 2) {
      transSheet.getRange(3, optDateIdx + 1, lastRow - 2, 1).clearContent();
    }

    this._optimizeAllCore();
  }

  optimizeAll() {
    this.runWithStatus(() => this._optimizeAllCore());
  }

  resetAndRecalculateAll() {
    this.runWithStatus(() => this._resetAndRecalculateAllCore());
  }
}
