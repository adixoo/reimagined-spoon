import { BrowserWindow } from "electron";
import { currentActiveBrowsers } from "../../store/currentActiveBrowsers";
import { FormEngine, type Step, type FormEngineOptions, type EngineRunResult } from "./engine";

/**
 * Engine-run outcome returned to the renderer.
 *
 * - status:'success' — successCondition matched mid-loop (instant) OR
 *   loop exited cleanly with no successCondition configured.
 * - status:'unknown' — loop exited but successCondition never matched,
 *   or runTimeout fired. Route to Unknown triage panel.
 * - status:'failed'  — hard failure: infinite-loop guard, max-iterations,
 *   or a required command exhausted its retries.
 */
export type { EngineRunResult };

/**
 * IPC handler for `engine:run`.
 *
 * Resolves the active Puppeteer page for `uid` from `currentActiveBrowsers`,
 * constructs a FormEngine with the caller's steps + options, and streams every
 * engine log line back to all renderer windows via `engine:log` IPC.
 *
 * @param uid     Frontend stable key identifying the browser session
 * @param steps   FormEngine step definitions (already hydrated with row data)
 * @param options Passed through to FormEngine constructor — notably:
 *                  `successCondition` for outcome detection
 *                  `runTimeout`       ms before status:unknown
 *                  `pollInterval`     ms between condition scans (default 1000)
 *                  `debug`            verbose logging
 */
export const engineRun = async (
  _: unknown,
  uid: string,
  steps: Step[],
  options?: FormEngineOptions,
): Promise<EngineRunResult> => {
  const entry = currentActiveBrowsers.get(uid);

  console.log("entry", entry);
  console.log("steps", steps);
  console.log("options", options);

  if (!entry) {
    return {
      status: "failed",
      context: {},
      error: `No active browser found for uid=${uid}`,
    };
  }

  const { page } = entry;

  // Logger — streams each log line to all renderer windows in real time.
  const sendLog = (message: string) => {
    BrowserWindow.getAllWindows().forEach((win) => {
      win.webContents.send("engine:log", uid, message);
    });
  };

  try {
    sendLog(`[${uid}] Engine starting — ${steps.length} steps loaded`);

    const engine = new FormEngine(page, steps, {
      ...options,
      // Default debug:true if no options passed at all
      debug: options && Object.keys(options).length > 0 ? (options.debug ?? false) : true,
      // Pipe all engine log lines through IPC to renderer
      logger: (...args: unknown[]) => {
        const message = args.join(" ");
        sendLog(message);
        if (options?.debug) {
          console.log("[FormEngine]", message);
        }
      },
    });

    const result = await engine.run();

    if (result.status === "success") {
      sendLog(`[${uid}] ✅ Engine finished — success`);
    } else if (result.status === "unknown") {
      sendLog(`[${uid}] ⚠ Engine finished — unknown (reason: ${result.reason ?? "n/a"})`);
    } else {
      sendLog(`[${uid}] ❌ Engine failed: ${result.error ?? "unknown error"}`);
    }

    return result;
  } catch (e: unknown) {
    const message = (e as Error).message;
    sendLog(`[${uid}] ❌ Engine error: ${message}`);
    return { status: "failed", context: {}, error: message };
  }
};
