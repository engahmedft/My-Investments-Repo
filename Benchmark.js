// =========================================================================
// PERFORMANCE BENCHMARK UTILITY (100% STATIC EDITION)
// =========================================================================

/**
 * High-precision, 100% static benchmark engine for tracking pipeline performance.
 * Eliminates object instantiations and supports named session profiling.
 * 
 * Run `Benchmark.help()` for full architectural specifications and API contracts.
 */
class Benchmark {

  /**
   * Prevent instantiation of purely static utility class.
   */
  constructor() {
    throw new Error("Benchmark is a static utility class and cannot be instantiated. Use Benchmark.start(), Benchmark.time(), and Benchmark.report().");
  }

  /**
   * Returns and logs the architectural specifications, execution contracts, and interface definitions.
   * @returns {string} Formatted architectural specification.
   */
  static help() {
    const doc = `
====================================================================================================
BENCHMARK ENGINE SPECIFICATION & EXECUTION CONTRACT (100% STATIC)
====================================================================================================

1. HIGH-PRECISION PIPELINE PROFILING:
   Lightweight performance instrumentation engine optimized for Google Apps Script (V8):
     - Zero object allocations: pure static execution holding sessions in memory.
     - Multi-session registry: profile concurrent or nested pipelines by name.
     - High-resolution timing via performance.now() with automatic Date.now() fallback.
     - Zero-overhead bypass: disabled sessions execute functions directly with zero closure allocations.
     - Single-event log consolidation: batches report lines to eliminate Cloud Logging RPC spam.
     - Native support for both synchronous routines and asynchronous Promises.

2. UNIFIED STATIC EXECUTION INTERFACES:
   All operations are 100% static:
   - Benchmark.start(pipelineName = "Default", enabled = true, precision = 2):
     Initializes or resets a profiling session and sets it as the active session.
   - Benchmark.time(label, fn):
     Measures execution time of fn() under the currently active session.
   - Benchmark.time(pipelineName, label, fn):
     Measures execution time of fn() under a specific named session.
   - Benchmark.report(pipelineName):
     Prints a single consolidated ASCII execution scorecard and returns structured metrics.
     (Defaults to the active session if pipelineName is omitted).
   - Benchmark.reset(pipelineName):
     Clears recorded steps and restarts the elapsed time baseline for a session.
   - Benchmark.run(label, fn):
     One-off convenience runner to profile an isolated function call without manual session setup.

3. STRUCTURED RETURN CONTRACT:
   Benchmark.report(pipelineName) returns:
     {
       pipelineName: string,
       enabled: boolean,
       totalDurationMs: number,
       totalDurationSec: string,
       steps: Array<{ label: string, durationMs: number, failed: boolean }>
     }
====================================================================================================`.trim();

    Logger.log(doc);
    return doc;
  }

  /**
   * Static session cache storing active pipeline profiles.
   * @private
   */
  static get _sessions() {
    if (!Benchmark._sessionsCache) {
      Benchmark._sessionsCache = new Map();
    }
    return Benchmark._sessionsCache;
  }

  /**
   * Monotonic high-resolution clock with Date.now fallback.
   * @private
   */
  static _now() {
    return (typeof performance !== 'undefined' && performance.now)
      ? performance.now()
      : Date.now();
  }

  /**
   * Formats duration in milliseconds to seconds string.
   * @private
   */
  static _formatMs(ms, precision = 2) {
    return (ms / 1000).toFixed(precision) + 's';
  }

  /**
   * Initializes or resets a named benchmarking session.
   * @param {string} [pipelineName="Default Pipeline"] - Identifier for the session.
   * @param {boolean} [enabled=true] - Toggle tracking and logging.
   * @param {number} [precision=2] - Decimal places for seconds formatting.
   * @returns {Object} Initialized session state.
   */
  static start(pipelineName = "Default Pipeline", enabled = true, precision = 2) {
    const session = {
      pipelineName,
      enabled,
      precision,
      startTime: Benchmark._now(),
      steps: []
    };

    Benchmark._sessions.set(pipelineName, session);
    Benchmark._activeSession = pipelineName;

    if (enabled) {
      console.log(`\n🚀 === STARTING PIPELINE: ${pipelineName} ===`);
    }

    return session;
  }

  /**
   * Records a completed step metric into a target session.
   * @private
   */
  static _recordStep(session, label, elapsedMs, failed, error = null) {
    session.steps.push({ label, durationMs: elapsedMs, failed });
    if (failed) {
      console.error(`❌ [FAILED] ${label} after ${Benchmark._formatMs(elapsedMs, session.precision)} - Error: ${error ? error.message : "Unknown"}`);
    } else {
      console.log(`✅ [DONE]  ${label} (${Benchmark._formatMs(elapsedMs, session.precision)})`);
    }
  }

  /**
   * Times the execution of a synchronous or asynchronous function.
   * Supports both (label, fn) for the active session, or (pipelineName, label, fn) for specific sessions.
   * 
   * @param {string} arg1 - Pipeline name OR step label.
   * @param {string|Function} arg2 - Step label OR function to execute.
   * @param {Function} [arg3] - Function to execute (when pipelineName is passed as arg1).
   * @returns {Promise<*>|*} Result of the executed function.
   */
  static time(arg1, arg2, arg3) {
    let pipelineName, label, fn;

    if (typeof arg3 === 'function') {
      pipelineName = String(arg1);
      label = String(arg2);
      fn = arg3;
    } else if (typeof arg2 === 'function') {
      pipelineName = Benchmark._activeSession || "Default Pipeline";
      label = String(arg1);
      fn = arg2;
    } else {
      throw new TypeError("Benchmark.time requires a function argument.");
    }

    let session = Benchmark._sessions.get(pipelineName);
    if (!session) {
      session = Benchmark.start(pipelineName);
    }

    // FAST PATH: Zero timing & logging overhead if disabled
    if (!session.enabled) {
      return fn();
    }

    console.log(`⏱️  [START] ${label}`);
    const start = Benchmark._now();

    let result;
    try {
      result = fn();
    } catch (err) {
      const elapsedMs = Benchmark._now() - start;
      Benchmark._recordStep(session, label, elapsedMs, true, err);
      throw err;
    }

    // Handle Promises / Async functions
    if (result && typeof result.then === 'function') {
      return result
        .then(res => {
          const elapsedMs = Benchmark._now() - start;
          Benchmark._recordStep(session, label, elapsedMs, false);
          return res;
        })
        .catch(err => {
          const elapsedMs = Benchmark._now() - start;
          Benchmark._recordStep(session, label, elapsedMs, true, err);
          throw err;
        });
    }

    // Synchronous execution (zero closure allocation)
    const elapsedMs = Benchmark._now() - start;
    Benchmark._recordStep(session, label, elapsedMs, false);
    return result;
  }

  /**
   * Resets benchmark step history and baseline start time for a session.
   * @param {string} [pipelineName] - Defaults to active session.
   */
  static reset(pipelineName) {
    const target = pipelineName || Benchmark._activeSession || "Default Pipeline";
    const session = Benchmark._sessions.get(target);
    if (session) {
      session.steps = [];
      session.startTime = Benchmark._now();
      if (session.enabled) {
        console.log(`\n🔄 === RESET PIPELINE: ${session.pipelineName} ===`);
      }
    }
  }

  /**
   * Generates benchmark report and returns structured metric results.
   * Batches output into a single console write to prevent Cloud Logging fragmentation.
   * 
   * @param {string} [pipelineName] - Defaults to active session.
   * @returns {{pipelineName: string, enabled: boolean, totalDurationMs: number, totalDurationSec: string, steps: Array}}
   */
  static report(pipelineName) {
    const target = pipelineName || Benchmark._activeSession || "Default Pipeline";
    const session = Benchmark._sessions.get(target);

    if (!session || !session.enabled) {
      return {
        pipelineName: target,
        enabled: false,
        totalDurationMs: 0,
        totalDurationSec: '0s',
        steps: []
      };
    }

    const totalMs = Benchmark._now() - session.startTime;
    const formattedTotal = Benchmark._formatMs(totalMs, session.precision);

    let maxLabelLen = 38;
    for (let i = 0; i < session.steps.length; i++) {
      if (session.steps[i].label.length > maxLabelLen) {
        maxLabelLen = session.steps[i].label.length;
      }
    }

    const divider = '='.repeat(maxLabelLen + 20);
    const lines = [
      '',
      divider,
      `📊 PERFORMANCE BENCHMARK REPORT: ${session.pipelineName}`,
      `⏱️ Total Execution Time: ${formattedTotal}`,
      '-'.repeat(maxLabelLen + 20)
    ];

    for (let i = 0; i < session.steps.length; i++) {
      const step = session.steps[i];
      const paddedLabel = step.label.padEnd(maxLabelLen);
      const status = step.failed ? '❌ FAILED' : '  ';
      lines.push(` - ${paddedLabel} : ${Benchmark._formatMs(step.durationMs, session.precision)} ${status}`);
    }

    lines.push(divider, '');
    console.log(lines.join('\n'));

    return {
      pipelineName: session.pipelineName,
      enabled: true,
      totalDurationMs: totalMs,
      totalDurationSec: formattedTotal,
      steps: session.steps
    };
  }

  /**
   * Static convenience runner for one-off function benchmarks.
   * @param {string} label - Description of the operation.
   * @param {Function} fn - Function to measure.
   * @returns {*} Result of fn().
   */
  static run(label, fn) {
    Benchmark.start(label, true, 2);
    try {
      const res = Benchmark.time(label, label, fn);
      Benchmark.report(label);
      return res;
    } catch (e) {
      Benchmark.report(label);
      throw e;
    }
  }
}