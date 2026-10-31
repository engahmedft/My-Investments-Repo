/**
 * Core orchestrator and calculation engine for investment portfolio operations.
 * STRICT CANONICAL 100% STATIC EDITION: Zero object instantiation, zero aliases, zero legacy fallbacks.
 * All methods are pure static functions focused strictly on portfolio management.
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
     Stage 1: Market Feeds optimization (Market_Feeds sheet), multi-tier sorting (Market_ID
              first, then missing 52-week low rows [hasNoMin], then natural alphanumeric MF_Key),
              Spot Gold price into Manual_Market_Price, Yahoo 52-week low into Yahoo_Low52
              and Manual_Low52 (Yahoo_Market_Price strictly kept empty for Gold), parallel Yahoo
              Finance price backfill for equities, API error tracking, and timestamp formatting (yyyy-mm-dd hh:mm:ss).
     Stage 2: Transaction chronological sort (Execution Date -> Trans Date -> Company -> Order Type
              -> Last Update -> Descending Avg Cost), running average cost basis calculation,
              future dividend quantity synchronization, and cross-platform dividend auto-replication.
     Stage 3: Position synchronization into Portfolio_Holdings sheet (updating average cost,
              computing strict PH_Key = MF_ID & " - " & Portfolio_ID, and inserting new positions).
     Stage 4: Automated formula and table registry updates via WorkbookRegistry.runAllRegistries.
     Stage 5: Audit log persistence in Optimize_Log sheet with granular operation and API error records.
     Stage 6: Conditional formatting and table visual styling reconciliation via TableFormatter.

2. UNIFIED STATIC EXECUTION INTERFACES & API CONTRACTS:
   All operations are 100% static, stateless, and allocate zero instance wrappers.
   - PortfolioManager.optimizeAll(options):
     Full 6-stage pipeline execution across all sheets, price backfills, registries, and formatting.
     Options: { reset = false, runRegistries = true, format = true, spreadsheet = null }
       * reset: Clears 'Optimize Date' in Transactions and 'Optimize_Date' in Market_Feeds to force full recalculation.
       * runRegistries: Executes Stage 4 formula and table registry updates via WorkbookRegistry.
       * format: Executes Stage 6 conditional formatting via TableFormatter.
       * spreadsheet: Optional target Spreadsheet instance (defaults to SpreadsheetApp.getActiveSpreadsheet()).

   - PortfolioManager.optimizeTransactionsAndStocks(options):
     Ultra-fast offline execution strictly running Stage 2 (Transactions) & Stage 3 (Portfolio_Holdings).
     Synchronizes PH_Key (MF_ID - Portfolio_ID) in Portfolio_Holdings.
     Does not execute Market Feeds optimization, external HTTP calls, registries, or formatting.
     Options: { reset = false, spreadsheet = null }
       * reset: Clears 'Optimize Date' checkpoints to force a full baseline recalculation.
       * spreadsheet: Optional target Spreadsheet instance.

   - PortfolioManager.resetAndRecalculateAll(options):
     Convenience wrapper invoking optimizeAll({ reset: true }) to force a full baseline recalculation.
     Clears Optimize Dates in both Transactions and Market_Feeds sheets.

   - PortfolioManager.optimizeMarketFeeds(logEntries, spreadsheet):
     Targeted optimization, key computation, multi-tier sorting (Market_ID -> hasNoMin -> MF_Key),
     timestamp formatting, parallel Yahoo Finance backfill for equities, Spot Gold backfill, and API error logging.

3. DATA CONNECTORS & ASSET NORMALIZATION:
   - Spot Gold Hybrid Model (fetchGoldPriceOrg):
     * Current Spot Price: Real-time OTC Spot Gold (XAU/USD) from open endpoint (api.goldprice.dev).
     * 52-Week Low: Institutional 52-week low strictly from Yahoo Finance quote feed ('GC=F' fiftyTwoWeekLow).
     * Concurrently requested via UrlFetchApp.fetchAll() in a single parallel roundtrip.
     * SAMA Currency Conversion: Converted from USD/Troy Ounce to 24K pure gold in SAR/Gram:
       Formula: Price_SAR_per_Gram = Price_USD_per_Oz * (3.75 / 31.1034768).
     * Target Columns for Gold:
       - Manual_Market_Price: Spot price (~499.24 SAR).
       - Yahoo_Market_Price: Strictly EMPTY ("").
       - Yahoo_Low52: Yahoo 52-week low.
       - Manual_Low52: Yahoo 52-week low.
   - Yahoo Finance API:
     * Requested concurrently via UrlFetchApp.fetchAll() in 50-symbol batches for equities.
     * Saudi Tadawul (Market 302 or 4-digit numeric symbols): Appends '.SR' (e.g., '1120.SR').
     * Gold is separated from the general equity batch to preserve OTC spot pricing.
   - API Error Audit:
     * Any failed HTTP status codes or payload parsing errors are captured and committed to Optimize_Log.

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
   - ARRAY_FORMULAS Registry: Protects live calculated columns across sheets by nulling them during writebacks.
   - Optimize_Log: Retains audit logs with action classifications (UPDATE, INSERT, ERROR), ticker symbols,
     platforms, and calculated parameter diffs.
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
   * Pre-compiled regex for Saudi Tadawul 4-digit numeric tickers.
   * @private
   */
  static get _TADAWUL_REGEX() {
    if (!PortfolioManager._TADAWUL_REGEX_CACHE) {
      PortfolioManager._TADAWUL_REGEX_CACHE = /^\d{4}$/;
    }
    return PortfolioManager._TADAWUL_REGEX_CACHE;
  }

  /**
   * Pre-compiled, reusable alphanumeric collator for fast natural sorting.
   * @private
   */
  static get _COLLATOR() {
    if (!PortfolioManager._COLLATOR_CACHE) {
      PortfolioManager._COLLATOR_CACHE = new Intl.Collator(undefined, { numeric: true, sensitivity: 'base' });
    }
    return PortfolioManager._COLLATOR_CACHE;
  }

  /**
   * High-performance 2-decimal point rounder.
   * @private
   */
  static _round2(val) {
    return Math.round(val * 100) / 100;
  }

  /**
   * Canonical registry of ArrayFormula column headers per sheet.
   * Nullified during write-backs to prevent Google Apps Script from overwriting live sheet formulas.
   * @returns {Record<string, string[]>}
   */
  static get ARRAY_FORMULAS() {
    if (!PortfolioManager._ARRAY_FORMULAS_CACHE) {
      PortfolioManager._ARRAY_FORMULAS_CACHE = {
        "Market_Feeds": [
          "Google_Name",
          "Google_Low52",
          "Google_Market_Price",
          "Low52",
          "Market_Price"
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
        "Portfolio_Holdings": [
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
   * @param {string} sheetName
   * @returns {string[]}
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
    const num = Number(n);
    return num < 10 ? '0' + num : String(num);
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
   * @param {any} val
   * @returns {number}
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
   * @param {any} val
   * @returns {string}
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
   * @returns {number}
   */
  static getEffectiveCost(item) {
    if (!item) return 0;
    return item.normalAvgCost !== null ? item.normalAvgCost : item.futureAvgCost;
  }

  /**
   * Helper to detect if a symbol or name represents a Gold asset.
   * Matches 'Gold', 'Gold.', 'GC=F', 'XAU', Arabic 'ذهب', or Market_ID containing 'Gold'.
   * @private
   */
  static _isGoldAsset(symbol, manualName = "", marketId = "") {
    const s = String(symbol ?? "").trim().toLowerCase();
    const m = String(manualName ?? "").trim().toLowerCase();
    const mid = String(marketId ?? "").trim().toLowerCase();
    return s === "gold" || s === "gold." || s.startsWith("gold") || s.includes("gold") ||
           s === "gc=f" || s.startsWith("gc=") || s === "xau" || s.startsWith("xau") ||
           m.includes("gold") || m.includes("ذهب") || mid.includes("gold") || mid === "gold";
  }

  /**
   * Appends an API error row into the log entries collection and writes to Logger.
   * @private
   */
  static _logApiError(logEntries, platform, symbol, details) {
    const errorEntry = [
      new Date(),
      "Market_Feeds",
      "ERROR",
      "N/A",
      platform || "API",
      symbol || "N/A",
      "N/A",
      "API Call",
      details
    ];

    if (Array.isArray(logEntries)) {
      logEntries.push(errorEntry);
    }
    Logger.log(`[API ERROR - ${platform}] ${symbol}: ${details}`);
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
    const round2 = PortfolioManager._round2;

    if (isDiv) {
      if (isSaudi) return 0;
      if (isUs) return round2(tv * 0.3) * sign;
      return 0;
    }

    const isSahm = platform === "Sahm Capital" || platform === "Sahm2 Capital";
    const isAlrajhi = platform === "Alrajhi Capital" || platform === "Alrajhi Wallet";

    if (isSaudi) {
      const cmaFee = round2(tv * 0.0003);
      const tadFees = round2(tv * 0.00009) + 2 * round2(tv * 0.00005) + round2(tv * 0.00001);
      const brkRate = isSahm ? 0 : (isAlrajhi ? 0.00105 : 0.001009);
      const vatableFees = tadFees + round2(tv * brkRate);
      const vat = round2(vatableFees * 0.15);
      return round2(cmaFee + vatableFees + vat) * sign;
    }

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
  // PRECIOUS METALS: SPOT GOLD (OPEN SPOT API) + 52-WEEK LOW (YAHOO)
  // =========================================================================

  /**
   * Fetches real-time spot gold price from open spot endpoint (api.goldprice.dev)
   * and 52-week low strictly from Yahoo Finance ('GC=F') concurrently in a single parallel HTTP roundtrip.
   * Converted to 24K SAR/gram at the official SAMA 3.750 fixed peg.
   * Does not get blocked with 403 Forbidden.
   * @param {Array} [logEntries]
   * @returns {{ price: number, low52: number }}
   */
  static fetchGoldPriceOrg(logEntries = null) {
    let priceSAR = 0;
    let low52SAR = 0;
    const ozToGramSAR = 3.75 / 31.1034768;
    const browserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";

    const { cookie, crumb } = PortfolioManager._getYahooCredentials(logEntries);
    let yfUrl = "https://query1.finance.yahoo.com/v7/finance/quote?symbols=GC=F";
    if (crumb) yfUrl += `&crumb=${encodeURIComponent(crumb)}`;

    const yfHeaders = { "User-Agent": browserAgent };
    if (cookie) yfHeaders["Cookie"] = cookie;

    // Open spot API endpoint built specifically for Google Apps Script without Cloudflare blocks
    const requests = [
      {
        url: "https://api.goldprice.dev/v1/spot/XAU-USD-SPOT",
        muteHttpExceptions: true,
        headers: { "User-Agent": browserAgent, "Accept": "application/json" }
      },
      {
        url: yfUrl,
        muteHttpExceptions: true,
        headers: yfHeaders
      }
    ];

    try {
      const responses = UrlFetchApp.fetchAll(requests);

      // 1. Process Spot Gold from open spot endpoint
      if (responses[0]) {
        const respCode = responses[0].getResponseCode();
        if (respCode === 200) {
          try {
            const json = JSON.parse(responses[0].getContentText());
            // Handles both { price: "4140.80" } and { symbols: [{ price: 4140.80 }] }
            const spotUSD = Number(json?.price || json?.symbols?.[0]?.price) || 0;
            if (spotUSD > 0) {
              priceSAR = Math.round(spotUSD * ozToGramSAR * 100) / 100;
            } else {
              PortfolioManager._logApiError(logEntries, "GoldPrice Spot", "Gold", "Returned 200 but spot price was missing or 0.");
            }
          } catch (e) {
            PortfolioManager._logApiError(logEntries, "GoldPrice Spot", "Gold", `JSON parse error: ${e.message}`);
          }
        } else {
          PortfolioManager._logApiError(logEntries, "GoldPrice Spot", "Gold", `HTTP error ${respCode}: ${responses[0].getContentText().slice(0, 150)}`);
        }
      }

      // 2. Process 52-Week Low strictly from Yahoo Finance ('GC=F')
      if (responses[1]) {
        const respCode = responses[1].getResponseCode();
        if (respCode === 200) {
          try {
            const json = JSON.parse(responses[1].getContentText());
            const quote = json?.quoteResponse?.result?.[0];
            const low52USD = Number(quote?.fiftyTwoWeekLow) || 0;
            if (low52USD > 0) {
              low52SAR = Math.round(low52USD * ozToGramSAR * 100) / 100;
            } else {
              PortfolioManager._logApiError(logEntries, "Yahoo Finance", "GC=F", "Returned 200 but fiftyTwoWeekLow was missing or 0.");
            }
          } catch (e) {
            PortfolioManager._logApiError(logEntries, "Yahoo Finance", "GC=F", `JSON parse error: ${e.message}`);
          }
        } else {
          PortfolioManager._logApiError(logEntries, "Yahoo Finance", "GC=F", `HTTP error ${respCode} while fetching 52-week low.`);
        }
      }
    } catch (err) {
      PortfolioManager._logApiError(logEntries, "GoldPrice/Yahoo", "Gold", `Parallel fetch exception: ${err.message}`);
    }

    return { price: priceSAR, low52: low52SAR };
  }

  // =========================================================================
  // PARALLEL BATCH YAHOO FINANCE API CONNECTOR (EQUITIES ONLY)
  // =========================================================================

  static _getYahooCredentials(logEntries = null) {
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
        } else {
          PortfolioManager._logApiError(logEntries, "Yahoo Auth", "Crumb", `Failed to fetch crumb: HTTP ${resp2.getResponseCode()}`);
        }
      } else {
        PortfolioManager._logApiError(logEntries, "Yahoo Auth", "Cookie", "Failed to retrieve authentication cookie from fc.yahoo.com.");
      }

      if (cookie && crumb) {
        cache.put("YF_COOKIE", cookie, 21600);
        cache.put("YF_CRUMB", crumb, 21600);
        return { cookie, crumb };
      }
    } catch (e) {
      PortfolioManager._logApiError(logEntries, "Yahoo Auth", "Crumb/Cookie", `Auth exception: ${e.message}`);
    }

    return { cookie: "", crumb: "" };
  }

  static _toYahooTicker(symbol, marketId) {
    const s = String(symbol ?? "").trim();
    if (!s) return "";

    const m = String(marketId ?? "").trim();
    if (m === "302" || PortfolioManager._TADAWUL_REGEX.test(s)) {
      return `${s}.SR`;
    }

    return s;
  }

  static fetchYahooQuotesBatch(tickers, logEntries = null) {
    const results = new Map();
    if (!tickers || tickers.length === 0) return results;

    const uniqueTickers = [...new Set(tickers.filter(t => Boolean(t)))];
    const { cookie, crumb } = PortfolioManager._getYahooCredentials(logEntries);
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

      for (let idx = 0; idx < responses.length; idx++) {
        const response = responses[idx];
        const status = response.getResponseCode();

        if (status === 200) {
          try {
            const json = JSON.parse(response.getContentText());
            const quotes = json?.quoteResponse?.result || [];

            for (const q of quotes) {
              const returnedTicker = q.symbol;
              const currentPrice = q.regularMarketPrice ?? 0;
              const yearLow = q.fiftyTwoWeekLow ?? 0;
              const companyName = q.shortName || q.longName || "";

              results.set(returnedTicker, {
                name: companyName,
                price: currentPrice,
                min52: yearLow
              });
            }
          } catch (jsonErr) {
            PortfolioManager._logApiError(logEntries, "Yahoo Finance", `Batch #${idx + 1}`, `JSON parse error: ${jsonErr.message}`);
          }
        } else {
          PortfolioManager._logApiError(logEntries, "Yahoo Finance", `Batch #${idx + 1}`, `HTTP status ${status}: ${response.getContentText().slice(0, 150)}`);
        }
      }
    } catch (err) {
      PortfolioManager._logApiError(logEntries, "Yahoo Finance", "Batch Fetch", `UrlFetchApp error: ${err.message}`);
    }

    return results;
  }

  // =========================================================================
  // CORE SPREADSHEET PIPELINES (STATIC)
  // =========================================================================

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
   * Optimized Audit Log persistence: Preserves formats and avoids repetitive setColumnWidth calls.
   * @private
   */
  static _writeLogSheet(logEntries, timestamp, spreadsheet = null) {
    const headers = ["Timestamp", "Sheet Name", "Action", "Order", "Platform", "Symbol", "Stock Name", "Order Type", "Details"];
    const ss = PortfolioManager.getSpreadsheet(spreadsheet);

    let logSheet = ss.getSheetByName("Optimize_Log");
    const isNew = !logSheet;
    if (isNew) {
      logSheet = ss.insertSheet("Optimize_Log");
      logSheet.setFrozenRows(1);
      logSheet.setColumnWidths(1, 8, 125);
      logSheet.setColumnWidth(9, 420);
    } else {
      logSheet.clearContents();
    }

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
  }

  /**
   * Clears Optimize Date checkpoints in both Transactions and Market_Feeds sheets to force a full baseline recalculation.
   * @private
   */
  static _clearAllOptimizeDates(spreadsheet = null) {
    const ss = PortfolioManager.getSpreadsheet(spreadsheet);

    const sheetsToClear = [
      { name: "Transactions", headerNames: ["Optimize Date", "Optimize_Date"] },
      { name: "Market_Feeds", headerNames: ["Optimize_Date", "Optimize Date"] }
    ];

    for (const item of sheetsToClear) {
      const sh = ss.getSheetByName(item.name);
      if (!sh) continue;

      const lastRow = sh.getLastRow();
      const lastCol = sh.getLastColumn();
      if (lastRow <= 2 || lastCol < 1) continue;

      const headers = sh.getRange(2, 1, 1, lastCol).getValues()[0];
      const optIdx = headers.findIndex(h => {
        const str = String(h || "").trim();
        return item.headerNames.includes(str);
      });

      if (optIdx !== -1) {
        sh.getRange(3, optIdx + 1, lastRow - 2, 1).clearContent();
      }
    }
  }

  /**
   * Optimizes the Market_Feeds sheet:
   * 1. Multi-tier sort: Market_ID first, then (No Google Low52 & No Yahoo Low52), then natural alphanumeric MF_Key.
   * 2. Dedicated Spot Gold fetch (GoldPrice spot into Manual_Market_Price, Yahoo 52-week low into Yahoo_Low52 and Manual_Low52).
   *    Yahoo_Market_Price is strictly left empty for Gold.
   * 3. Yahoo Finance parallel quotes backfill for equities.
   * 4. Comprehensive API error logging into Optimize_Log.
   * 5. Explicit timestamp formatting: Stamped Optimize_Date contains full time formatted as 'yyyy-mm-dd hh:mm:ss'.
   */
  static optimizeMarketFeeds(logEntries = null, spreadsheet = null) {
    const ss = PortfolioManager.getSpreadsheet(spreadsheet);
    const sheet = ss.getSheetByName("Market_Feeds");
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
              if (idx === -1) throw new Error(`[Market_Feeds] Missing canonical column: "${name}"`);
              return idx;
            };

            const keyCol          = requireCol("MF_Key");
            const orderCol        = requireCol("Sort_Order");
            const marketIdCol     = requireCol("Market_ID");
            const symbolCol       = requireCol("Symbol");
            const manualNameCol   = requireCol("Manual_Name");
            const combinedNameCol = requireCol("Combined_Name");
            const manualLow52Col  = requireCol("Manual_Low52");
            const googleLow52Col  = requireCol("Google_Low52");
            const yahooNameCol    = requireCol("Yahoo_Name");
            const yahooLow52Col   = requireCol("Yahoo_Low52");
            const manualPriceCol  = requireCol("Manual_Market_Price");
            const googlePriceCol  = requireCol("Google_Market_Price");
            const yahooPriceCol   = requireCol("Yahoo_Market_Price");
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
                    "Market_Feeds",
                    "UPDATE",
                    r[orderCol] || (i + 1),
                    marketId || "N/A",
                    symbol || "N/A",
                    manualName || "N/A",
                    "Field Computation",
                    `Computed: MF_Key="${computedKey}", Combined_Name="${computedCombinedName}"`
                  ]);
                }
              }

              const hasNoMin = !isPopulated(r[googleLow52Col]) && !isPopulated(r[yahooLow52Col]);

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
            const needsGoldFetch = [];

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
                  "Market_Feeds",
                  "UPDATE",
                  newOrder,
                  item.marketId || "N/A",
                  item.symbol || "N/A",
                  item.manualName || "N/A",
                  "Sort Sequence",
                  `Sort_Order changed from ${oldOrder} to ${newOrder}`
                ]);
              }

              // Separate Gold strictly from regular equities
              if (PortfolioManager._isGoldAsset(item.symbol, item.manualName, item.marketId)) {
                needsGoldFetch.push({
                  rowIndex: i,
                  symbol: item.symbol,
                  marketId: item.marketId,
                  manualName: item.manualName
                });
              } else if (!isPopulated(r[googlePriceCol])) {
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

            // 1. Strict Gold Processing: Spot to Manual, Yahoo to Low52, Yahoo_Market_Price EMPTY
            if (needsGoldFetch.length > 0) {
              const goldQuote = PortfolioManager.fetchGoldPriceOrg(standaloneLogs);
              for (const target of needsGoldFetch) {
                const r = sortedData[target.rowIndex];

                // Spot Price strictly to Manual_Market_Price
                if (goldQuote && goldQuote.price > 0) {
                  r[manualPriceCol] = goldQuote.price;
                }

                // Yahoo_Market_Price is strictly kept EMPTY for Gold
                r[yahooPriceCol] = "";
                r[yahooNameCol] = "";

                // Yahoo 52-week low to Yahoo_Low52 and Manual_Low52
                if (goldQuote && goldQuote.low52 > 0) {
                  r[yahooLow52Col] = goldQuote.low52;
                  r[manualLow52Col] = goldQuote.low52;
                }

                r[optDateCol] = nowTimestamp;

                standaloneLogs.push([
                  nowTimestamp,
                  "Market_Feeds",
                  "UPDATE",
                  r[orderCol],
                  target.marketId || "N/A",
                  target.symbol,
                  target.manualName || "Gold",
                  "Gold Backfill",
                  `Gold synchronized: Spot=${r[manualPriceCol]} SAR/g (Manual_Market_Price), Low52=${r[yahooLow52Col]} SAR/g (Yahoo_Low52 & Manual_Low52), Yahoo_Market_Price=EMPTY`
                ]);
              }
            }

            // 2. Parallel Yahoo Quotes Batch for Regular Equities
            if (needsYahooFetch.length > 0) {
              const allYahooTickers = needsYahooFetch.map(item => item.ticker);
              const quotesMap = PortfolioManager.fetchYahooQuotesBatch(allYahooTickers, standaloneLogs);

              for (const target of needsYahooFetch) {
                const r = sortedData[target.rowIndex];
                const quote = quotesMap.get(target.ticker);

                if (quote && quote.price > 0) {
                  r[yahooNameCol] = quote.name || "";
                  r[yahooLow52Col] = quote.min52 > 0 ? quote.min52 : "";
                  r[yahooPriceCol] = quote.price > 0 ? quote.price : "";

                  r[manualPriceCol] = quote.price;
                  if (quote.min52 > 0) {
                    r[manualLow52Col] = quote.min52;
                  }

                  r[optDateCol] = nowTimestamp;

                  standaloneLogs.push([
                    nowTimestamp,
                    "Market_Feeds",
                    "UPDATE",
                    r[orderCol],
                    target.marketId || "N/A",
                    target.symbol,
                    quote.name || "N/A",
                    "Yahoo API Fill",
                    `Yahoo backfill applied: Price=${quote.price}, 52WkLow=${quote.min52} (Manual Price/Low52 updated)`
                  ]);
                } else {
                  r[yahooNameCol] = "";
                  r[yahooLow52Col] = "";
                  r[yahooPriceCol] = "";
                }
              }
            }

            if (!logEntries && standaloneLogs.length > 0) {
              PortfolioManager._writeLogSheet(standaloneLogs, nowTimestamp, ss);
            }
          }

          return [sortedData[index]];
        },
        PortfolioManager.getArrayFormulas("Market_Feeds")
      );

      if (optDateColIdx !== -1) {
        const lastRow = sheet.getLastRow();
        if (lastRow >= 3) {
          sheet.getRange(3, optDateColIdx + 1, lastRow - 2, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
        }
      }

      ss.toast("Market Feeds optimized and synchronized.", "Success", 3);
    } catch (e) {
      Logger.log(`❌ optimizeMarketFeeds Error: ${e.message}`);
      ss.toast(e.message, "Optimization Error", 10);
      throw e;
    }
  }

  // =========================================================================
  // TARGETED TRANSACTIONS & Portfolio_Holdings ENGINE (STATIC)
  // =========================================================================

  /**
   * Optimizes Transactions and synchronizes Portfolio_Holdings positions completely offline.
   * Strictly executes Stage 2 & Stage 3:
   * - Transactions sort, fee evaluation, dividend auto-replication.
   * - Portfolio_Holdings cost updates and strict PH_Key concatenation: PH_Key = MF_ID & " - " & Portfolio_ID.
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

    const sessionName = isStandalone ? "Transactions & Portfolio_Holdings Pipeline" : "Optimize All Pipeline";
    if (isStandalone && typeof Benchmark !== "undefined") {
      Benchmark.start(sessionName);
    }
    const timeStep = (label, fn) => (typeof Benchmark !== "undefined" ? Benchmark.time(sessionName, label, fn) : fn());

    const execute = () => {
      if (reset) {
        PortfolioManager._clearAllOptimizeDates(ss);
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

      // Stage 3: Portfolio_Holdings Position Synchronization & Strict PH_Key Calculation
      const holdingsSheet = ss.getSheetByName("Portfolio_Holdings");
      if (holdingsSheet) {
        timeStep("Process & Write Portfolio_Holdings Pass", () => {
          let matchedKeys = new Set();
          let sLastCol = 0;
          let idxPlatform = -1, idxSymbol = -1, idxCompany = -1, idxCost = -1, lastUpdateIdx = -1;
          let idxPHKey = -1, idxMFId = -1, idxPortId = -1;
          const formulaIndices = [];

          BatchProcessor.process(
            holdingsSheet,
            2,
            (row, col, index, rawData) => {
              if (index === 0) {
                const requireCol = (name) => {
                  const idx = col(name);
                  if (idx === -1) throw new Error(`[Portfolio_Holdings] Missing canonical column: "${name}"`);
                  return idx;
                };

                idxPlatform   = requireCol("Platform");
                idxSymbol     = requireCol("Symbol");
                idxCompany    = requireCol("Company");
                idxCost       = requireCol("Avg Cost");
                lastUpdateIdx = requireCol("Last Update");

                // Strict canonical column resolution
                idxPHKey      = requireCol("PH_Key");
                idxMFId       = requireCol("MF_ID");
                idxPortId     = requireCol("Portfolio_ID");

                sLastCol      = row.length;

                const formulaHeaders = PortfolioManager.getArrayFormulas("Portfolio_Holdings");
                formulaHeaders.forEach(fh => {
                  const fIdx = col(fh);
                  if (fIdx !== -1) formulaIndices.push(fIdx);
                });
              }

              // Always apply concatenation rule unconditionally: PH_Key = MF_ID & " - " & Portfolio_ID
              const mfIdVal = String(row[idxMFId] ?? "").trim();
              const portIdVal = String(row[idxPortId] ?? "").trim();
              const computedPHKey = `${mfIdVal} - ${portIdVal}`;
              if (row[idxPHKey] !== computedPHKey) {
                row[idxPHKey] = computedPHKey;
                row[lastUpdateIdx] = nowTimestamp;
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
                        "Portfolio_Holdings",
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

                    // Unconditional concatenation on newly inserted rows
                    const newMfId = String(newRow[idxMFId] ?? "").trim();
                    const newPortId = String(newRow[idxPortId] ?? "").trim();
                    newRow[idxPHKey] = `${newMfId} - ${newPortId}`;

                    formulaIndices.forEach(idx => {
                      if (idx >= 0 && idx < sLastCol) newRow[idx] = null;
                    });

                    newStockRows.push(newRow);

                    currentLogs.push([
                      nowTimestamp,
                      "Portfolio_Holdings",
                      "INSERT",
                      "N/A",
                      item.origPlatform,
                      item.origSymbol,
                      item.origCompany || "",
                      "N/A",
                      `New position added with Avg Cost: ${Math.round(finalCost * 10000) / 10000}`
                    ]);
                  }
                }
                return [row, ...newStockRows];
              }

              return [row];
            },
            PortfolioManager.getArrayFormulas("Portfolio_Holdings")
          );
        });
      }

      if (isStandalone) {
        timeStep("Write Optimize Log Pass", () => {
          PortfolioManager._writeLogSheet(currentLogs, nowTimestamp, ss);
        });

        const durationLabel = (typeof Benchmark !== "undefined")
          ? Benchmark.report(sessionName)?.totalDurationSec
          : "few seconds";
        ss.toast(`Transactions & Portfolio_Holdings optimized in ${durationLabel}.`, "Success", 5);
      }
    };

    execute();
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
    const runRegistries = safeOpts.runRegistries ?? true;
    const format = safeOpts.format ?? true;
    const ss = PortfolioManager.getSpreadsheet(safeOpts.spreadsheet);

    const sessionName = "Optimize All Pipeline";
    if (typeof Benchmark !== "undefined") {
      Benchmark.start(sessionName);
    }
    const timeStep = (label, fn) => (typeof Benchmark !== "undefined" ? Benchmark.time(sessionName, label, fn) : fn());

    if (reset) {
      PortfolioManager._clearAllOptimizeDates(ss);
    }

    const nowTimestamp = new Date();
    const logEntries = [];

    // Stage 1: Market Feeds Optimization (Spot Gold, Equities Backfills & Error Capture)
    timeStep("Optimize Market Feeds", () => PortfolioManager.optimizeMarketFeeds(logEntries, ss));

    // Stage 2 & 3: Transactions & Portfolio_Holdings Processing (Includes strict PH_Key calculation)
    PortfolioManager.optimizeTransactionsAndStocks({
      reset: false,
      logEntries: logEntries,
      spreadsheet: ss
    });

    // Stage 4: Synchronize Table and Formula Registries (Exclusive to optimizeAll)
    if (runRegistries) {
      timeStep("Synchronize Registries", () => {
        if (typeof WorkbookRegistry !== "undefined" && typeof WorkbookRegistry.runAllRegistries === "function") {
          WorkbookRegistry.runAllRegistries(ss);
        } else if (typeof runAllRegistries === "function") {
          runAllRegistries();
        } else {
          if (typeof listAllTables === "function") listAllTables();
          if (typeof listAllFormulas === "function") listAllFormulas();
        }
      });
    }

    // Stage 5: Commit Audit Log (Includes UPDATE, INSERT, and all API ERROR entries)
    timeStep("Write Optimize Log Pass", () => {
      PortfolioManager._writeLogSheet(logEntries, nowTimestamp, ss);
    });

    // Stage 6: Visual Formatting (Exclusive to optimizeAll)
    if (format) {
      PortfolioManager._formatSheets(["Market_Feeds", "Transactions", "Portfolio_Holdings"], null, ss);
    }

    const durationLabel = (typeof Benchmark !== "undefined")
      ? Benchmark.report(sessionName)?.totalDurationSec
      : "few seconds";
    ss.toast(`Optimization complete in ${durationLabel}.`, "Optimize All", 5);
  }

  /**
   * One-line delegate for full recalculations.
   * Clears Optimize Dates across Transactions and Market_Feeds sheets.
   */
  static resetAndRecalculateAll(options = {}) {
    const isMenuOrEvent = Boolean(options && typeof options === "object" && ("authMode" in options || "range" in options));
    const safeOpts = (isMenuOrEvent || typeof options !== "object" || options === null) ? {} : options;
    PortfolioManager.optimizeAll({ ...safeOpts, reset: true });
  }
}