// eslint-disable-next-line @typescript-eslint/no-require-imports
type Page = import("puppeteer").Page;

// ─────────────────────────────────────────────────────────────────────────────
// TYPES
// ─────────────────────────────────────────────────────────────────────────────

export type EngineStatus = "success" | "unknown" | "failed";

export interface EngineRunResult {
  status: EngineStatus;
  context: Record<string, string>;
  error?: string;
  reason?: string;
}

// ── Condition ────────────────────────────────────────────────────────────────

export interface StepCondition {
  /** Selectors that must be visible + interactable */
  exists?: string[];
  /** Selectors that must be absent or hidden */
  notExists?: string[];
  /** Current URL must include this substring */
  urlContains?: string;
  /** document.body.innerText must include this substring */
  textContains?: string;
}

// ── Commands ─────────────────────────────────────────────────────────────────

interface BaseCommand {
  /** How many times to retry on failure (overrides engine default) */
  retries?: number;
  /** ms between retries (overrides engine default) */
  retryDelay?: number;
  /** If true, failure is non-fatal — engine logs and continues */
  optional?: boolean;
}

/** Type text into an input (clears first) */
export interface TypeCommand extends BaseCommand {
  action: "type";
  selector: string;
  /** Supports {{contextKey}} placeholders */
  value: string;
  /** ms delay between keystrokes — default 30 */
  typeDelay?: number;
}

/** Click an element */
export interface ClickCommand extends BaseCommand {
  action: "click";
  selector: string;
}

/** Select an <option> by value in a <select> */
export interface SelectCommand extends BaseCommand {
  action: "select";
  selector: string;
  /** The <option value="…"> to select. Supports {{contextKey}} */
  value: string;
}

/** Check a checkbox (no-op if already checked) */
export interface CheckCommand extends BaseCommand {
  action: "check";
  selector: string;
}

/** Uncheck a checkbox (no-op if already unchecked) */
export interface UncheckCommand extends BaseCommand {
  action: "uncheck";
  selector: string;
}

/** Set a radio button by clicking it */
export interface RadioCommand extends BaseCommand {
  action: "radio";
  selector: string;
}

/** Focus an element */
export interface FocusCommand extends BaseCommand {
  action: "focus";
  selector: string;
}

/** Hover over an element */
export interface HoverCommand extends BaseCommand {
  action: "hover";
  selector: string;
}

/**
 * Scroll an element into view.
 * NOTE: good place for a screenshot before/after for debugging scroll state.
 */
export interface ScrollToCommand extends BaseCommand {
  action: "scrollTo";
  selector: string;
}

/** Scroll to an absolute pixel position on the page */
export interface ScrollByCommand extends BaseCommand {
  action: "scrollBy";
  /** Pixels to scroll horizontally */
  x?: number;
  /** Pixels to scroll vertically */
  y?: number;
}

/** Press a named keyboard key (Enter, Tab, Escape, ArrowDown, …) */
export interface PressKeyCommand extends BaseCommand {
  action: "pressKey";
  /** Puppeteer KeyInput name e.g. "Enter", "Tab", "Escape", "ArrowDown" */
  key: string;
}

/** Hold modifier + press key (e.g. Ctrl+A, Shift+Tab) */
export interface KeyComboCommand extends BaseCommand {
  action: "keyCombo";
  /** Modifier keys to hold: "Control" | "Shift" | "Alt" | "Meta" */
  modifiers: Array<"Control" | "Shift" | "Alt" | "Meta">;
  /** Key to press while modifiers held */
  key: string;
}

/** Upload a file to a file input */
export interface UploadFileCommand extends BaseCommand {
  action: "uploadFile";
  selector: string;
  /** Absolute path on disk. Supports {{contextKey}} */
  filePath: string;
}

/**
 * Clear an input then type into it using the keyboard (no JS injection).
 * Useful for React-controlled inputs that ignore programmatic .value writes.
 */
export interface ClearAndTypeCommand extends BaseCommand {
  action: "clearAndType";
  selector: string;
  /** Supports {{contextKey}} placeholders */
  value: string;
  typeDelay?: number;
}

/** Wait a fixed number of milliseconds */
export interface WaitCommand extends BaseCommand {
  action: "wait";
  /** Milliseconds to wait */
  ms: number;
}

/** Wait until a selector is visible */
export interface WaitForSelectorCommand extends BaseCommand {
  action: "waitForSelector";
  selector: string;
  /** Timeout ms — default 15 000 */
  timeout?: number;
}

/** Wait for a navigation / page load to complete */
export interface WaitForNavigationCommand extends BaseCommand {
  action: "waitForNavigation";
  waitUntil?: "load" | "domcontentloaded" | "networkidle0" | "networkidle2";
  /** Timeout ms — default 15 000 */
  timeout?: number;
}

/**
 * Read an element's value or text and store it in engine context.
 * NOTE: good place for a screenshot to verify the read value.
 */
export interface ReadValueCommand extends BaseCommand {
  action: "readValue";
  selector: string;
  /** Key to store the value under in this.context */
  saveAs: string;
}

/** Manually set a context key to a literal value (supports {{placeholders}}) */
export interface SetContextCommand extends BaseCommand {
  action: "setContext";
  key: string;
  value: string;
}

/**
 * Take a screenshot and save to disk.
 * Use this command liberally:
 *   - before/after multi-field steps
 *   - after navigation
 *   - on error paths (optional:true so it never blocks)
 *   - before submit
 */
export interface ScreenshotCommand extends BaseCommand {
  action: "screenshot";
  /** File path to save PNG. Defaults to `screenshot_<timestamp>.png` */
  path?: string;
  /** Capture full scrollable page — default false */
  fullPage?: boolean;
}

/**
 * Conditional branch — no JS eval; compares context values only.
 * value1 / value2 support {{contextKey}} placeholders.
 */
export interface IfCommand extends BaseCommand {
  action: "if";
  value1: string;
  operator: "equals" | "notEquals" | "contains" | "notContains";
  value2: string;
  then?: Command[];
  otherwise?: Command[];
}

export type Command =
  | TypeCommand
  | ClickCommand
  | SelectCommand
  | CheckCommand
  | UncheckCommand
  | RadioCommand
  | FocusCommand
  | HoverCommand
  | ScrollToCommand
  | ScrollByCommand
  | PressKeyCommand
  | KeyComboCommand
  | UploadFileCommand
  | ClearAndTypeCommand
  | WaitCommand
  | WaitForSelectorCommand
  | WaitForNavigationCommand
  | ReadValueCommand
  | SetContextCommand
  | ScreenshotCommand
  | IfCommand;

// ── Steps ────────────────────────────────────────────────────────────────────

export interface Step {
  id: string;
  condition: StepCondition;
  commands: Command[];
  /** Higher = evaluated first. Default 0 */
  priority?: number;
  /** If false (default), step runs at most once per engine.run() */
  repeatable?: boolean;
}

// ── Engine options ────────────────────────────────────────────────────────────

export interface FormEngineOptions {
  /**
   * How often (ms) the engine polls all conditions (step + successCondition).
   * Replaces old stepSettleTime fixed-wait — engine fires every `pollInterval`
   * ms after page load instead of sleeping a fixed duration after each step.
   * Default: 1 000
   */
  pollInterval?: number;
  /** Default retry count per command. Default 3 */
  commandRetries?: number;
  /** ms between retries. Default 1 000 */
  retryDelay?: number;
  /** Safety cap on loop iterations. Default 30 */
  maxIterations?: number;
  /**
   * How many consecutive poll ticks a single step may match before the
   * engine treats it as an infinite loop. Default 3.
   */
  consecutiveMatchLimit?: number;
  /** Enable verbose console logging. Default false */
  debug?: boolean;
  /**
   * Checked on EVERY poll tick alongside step conditions — highest priority.
   * Matched → return status:'success' instantly, no further steps run.
   * Not matched after loop exits → status:'unknown'.
   * Omitted → legacy: clean exit = success.
   */
  successCondition?: StepCondition;
  /**
   * Total ms budget for engine.run(). Exceeding it → status:'unknown'.
   * 0 = no timeout (default).
   */
  runTimeout?: number;
  /** Override the internal logger (each call = one log line) */
  logger?: (...args: unknown[]) => void;
}

// ─────────────────────────────────────────────────────────────────────────────
// ENGINE
// ─────────────────────────────────────────────────────────────────────────────

export class FormEngine {
  private page: Page;
  private steps: Step[];
  private options: Required<
    Omit<FormEngineOptions, "successCondition" | "logger">
  > & {
    successCondition: StepCondition | null;
    logger: (...args: unknown[]) => void;
  };

  /** Steps that have completed (non-repeatable guard) */
  private executedSteps = new Set<string>();
  /** Shared state across steps — populated by readValue / setContext */
  private context: Record<string, string> = {};
  /** Per-step consecutive-match counter (infinite-loop guard) */
  private consecutiveMatchCount: Record<string, number> = {};

  private log: (...args: unknown[]) => void;

  constructor(page: Page, steps: Step[], options: FormEngineOptions = {}) {
    this.page = page;
    this.steps = this._normalizePriority(steps);

    const defaultLogger = options.debug
      ? (...args: unknown[]) => console.log("[FormEngine]", ...args)
      : () => {};

    this.log = options.logger ?? defaultLogger;

    this.options = {
      pollInterval: options.pollInterval ?? 1_000,
      commandRetries: options.commandRetries ?? 3,
      retryDelay: options.retryDelay ?? 1_000,
      maxIterations: options.maxIterations ?? 30,
      consecutiveMatchLimit: options.consecutiveMatchLimit ?? 3,
      debug: options.debug ?? false,
      successCondition: options.successCondition ?? null,
      runTimeout: options.runTimeout ?? 0,
      logger: this.log,
    };
  }

  // ─────────────────────────────────────────────
  // SETUP
  // ─────────────────────────────────────────────

  private _normalizePriority(steps: Step[]): Step[] {
    return [...steps].sort((a, b) => (b.priority ?? 0) - (a.priority ?? 0));
  }

  /** Resolve {{key}} placeholders from this.context */
  private _resolveValue(value: string): string {
    return value.replace(/\{\{(\w+)\}\}/g, (_, key: string) => {
      if (!(key in this.context)) {
        throw new Error(
          `Context key "{{${key}}}" not found. Available: ${
            Object.keys(this.context).join(", ") || "(none)"
          }`
        );
      }
      return this.context[key];
    });
  }

  // ─────────────────────────────────────────────
  // PAGE LOAD GUARD
  // ─────────────────────────────────────────────

  /**
   * Wait for the page to be fully loaded before scanning conditions.
   *
   * Strategy:
   *   1. domcontentloaded — DOM is parsed and ready (fast)
   *   2. networkidle2     — no more than 2 open network connections for 500ms
   *
   * Both are raced with a 30s safety timeout so the engine never hangs
   * indefinitely on a broken page load.
   */
  private async _waitForPageLoad(): Promise<void> {
    const LOAD_TIMEOUT = 30_000;

    try {
      // Step 1 — DOM ready
      await this.page.waitForFunction(
        () => document.readyState === "interactive" || document.readyState === "complete",
        { timeout: LOAD_TIMEOUT }
      );
      this.log("  📄 DOMContentLoaded");

      // Step 2 — network quiet
      await this.page.waitForNavigation({
        waitUntil: "networkidle2",
        timeout: LOAD_TIMEOUT,
      }).catch(() => {
        // networkidle2 can time out on pages with long-polling/websockets —
        // treat as non-fatal; DOM is ready so we can still scan conditions.
        this.log("  ⚠ networkidle2 timeout (non-fatal) — scanning anyway");
      });
      this.log("  🌐 networkidle2");
    } catch (e: unknown) {
      // If even domcontentloaded times out, log and continue — better to
      // attempt condition matching than to hard-fail the whole run.
      this.log(`  ⚠ Page load check failed: ${(e as Error).message} — continuing`);
    }
  }

  // ─────────────────────────────────────────────
  // VISIBILITY CHECKS  (no JS injection — uses Puppeteer's built-ins)
  // ─────────────────────────────────────────────

  /**
   * True iff selector resolves to an element that is:
   *   - rendered (not display:none / visibility:hidden / opacity:0)
   *   - has non-zero bounding rect
   *   - not disabled
   *
   * Uses Puppeteer's built-in $eval with standard DOM APIs — no arbitrary
   * JS string injection.
   */
  private async _isVisibleAndInteractable(selector: string): Promise<boolean> {
    try {
      const el = await this.page.$(selector);
      if (!el) return false;

      return await this.page.$eval(selector, (el: Element) => {
        const style = window.getComputedStyle(el as HTMLElement);
        const rect = (el as HTMLElement).getBoundingClientRect();
        return (
          style.display !== "none" &&
          style.visibility !== "hidden" &&
          style.opacity !== "0" &&
          rect.height > 0 &&
          rect.width > 0 &&
          !(el as HTMLInputElement).disabled
        );
      });
    } catch {
      return false;
    }
  }

  /** True iff selector is absent from DOM OR is hidden */
  private async _isAbsent(selector: string): Promise<boolean> {
    const el = await this.page.$(selector);
    if (!el) return true;
    const visible = await this._isVisibleAndInteractable(selector);
    return !visible;
  }

  // ─────────────────────────────────────────────
  // CONDITION EVALUATION
  // ─────────────────────────────────────────────

  /**
   * Evaluate a StepCondition against the live page.
   * No JS string injection — every check uses Puppeteer native APIs or
   * $eval with typed DOM accessors only.
   *
   * Shared by step-matching AND successCondition outcome check.
   */
  private async _evaluateCondition(cond: StepCondition): Promise<boolean> {
    const { exists = [], notExists = [], urlContains, textContains } = cond;

    // URL substring check
    if (urlContains) {
      if (!this.page.url().includes(urlContains)) return false;
    }

    // Body text check — needle passed as serialised arg, no string injection
    if (textContains) {
      try {
        const found = await this.page.$eval(
          "body",
          (body: Element, needle: unknown) =>
            ((body as HTMLElement).innerText ?? "").includes(needle as string),
          textContains
        );
        if (!found) return false;
      } catch {
        return false;
      }
    }

    for (const sel of exists) {
      if (!(await this._isVisibleAndInteractable(sel))) return false;
    }

    for (const sel of notExists) {
      if (!(await this._isAbsent(sel))) return false;
    }

    return true;
  }

  // ─────────────────────────────────────────────
  // COMMAND EXECUTION
  // ─────────────────────────────────────────────

  private async _executeCommand(cmd: Command): Promise<void> {
    switch (cmd.action) {

      // ── Text input ──────────────────────────────────────────────────────

      case "type": {
        const value = this._resolveValue(cmd.value);
        await this.page.waitForSelector(cmd.selector, { visible: true, timeout: 8_000 });
        // Select-all then delete → avoids needing JS to clear .value
        await this.page.click(cmd.selector, { clickCount: 3 });
        await this.page.keyboard.press("Backspace");
        await this.page.type(cmd.selector, value, {
          delay: cmd.typeDelay ?? 30,
        });
        break;
      }

      case "clearAndType": {
        /**
         * Keyboard-only clear + type (no JS .value = '' injection).
         * Works for React-controlled inputs that ignore programmatic value writes.
         * Focus → Ctrl+A → Backspace → type.
         */
        const value = this._resolveValue(cmd.value);
        await this.page.waitForSelector(cmd.selector, { visible: true, timeout: 8_000 });
        await this.page.focus(cmd.selector);
        await this.page.keyboard.down("Control");
        await this.page.keyboard.press("KeyA");
        await this.page.keyboard.up("Control");
        await this.page.keyboard.press("Backspace");
        await this.page.type(cmd.selector, value, {
          delay: cmd.typeDelay ?? 30,
        });
        break;
      }

      // ── Click ───────────────────────────────────────────────────────────

      case "click": {
        await this.page.waitForSelector(cmd.selector, { visible: true, timeout: 8_000 });
        await this.page.click(cmd.selector);
        break;
      }

      // ── Select / Dropdowns ──────────────────────────────────────────────

      case "select": {
        const value = this._resolveValue(cmd.value);
        await this.page.waitForSelector(cmd.selector, { visible: true, timeout: 8_000 });
        await this.page.select(cmd.selector, value);
        break;
      }

      // ── Checkboxes / Radio ──────────────────────────────────────────────

      case "check": {
        await this.page.waitForSelector(cmd.selector, { visible: true, timeout: 8_000 });
        const checked = await this.page.$eval(
          cmd.selector,
          (el: Element) => (el as HTMLInputElement).checked
        );
        if (!checked) await this.page.click(cmd.selector);
        break;
      }

      case "uncheck": {
        await this.page.waitForSelector(cmd.selector, { visible: true, timeout: 8_000 });
        const unchecked = await this.page.$eval(
          cmd.selector,
          (el: Element) => (el as HTMLInputElement).checked
        );
        if (unchecked) await this.page.click(cmd.selector);
        break;
      }

      case "radio": {
        // Click radio only if not already selected
        await this.page.waitForSelector(cmd.selector, { visible: true, timeout: 8_000 });
        const selected = await this.page.$eval(
          cmd.selector,
          (el: Element) => (el as HTMLInputElement).checked
        );
        if (!selected) await this.page.click(cmd.selector);
        break;
      }

      // ── Focus / Hover ───────────────────────────────────────────────────

      case "focus": {
        await this.page.waitForSelector(cmd.selector, { visible: true, timeout: 8_000 });
        await this.page.focus(cmd.selector);
        break;
      }

      case "hover": {
        await this.page.waitForSelector(cmd.selector, { visible: true, timeout: 8_000 });
        await this.page.hover(cmd.selector);
        break;
      }

      // ── Scroll ──────────────────────────────────────────────────────────

      case "scrollTo": {
        /**
         * NOTE: Consider adding a screenshot command before/after scrollTo
         * when debugging layout or visibility issues on long pages.
         */
        await this.page.waitForSelector(cmd.selector, { timeout: 8_000 });
        await this.page.$eval(cmd.selector, (el: Element) =>
          el.scrollIntoView({ behavior: "smooth", block: "center" })
        );
        break;
      }

      case "scrollBy": {
        /**
         * NOTE: Screenshot after scrollBy useful when verifying lazy-loaded
         * content appears after scroll.
         */
        await this.page.mouse.wheel({
          deltaX: cmd.x ?? 0,
          deltaY: cmd.y ?? 0,
        });
        break;
      }

      // ── Keyboard ────────────────────────────────────────────────────────

      case "pressKey": {
        await this.page.keyboard.press(cmd.key as any);
        break;
      }

      case "keyCombo": {
        for (const mod of cmd.modifiers) await this.page.keyboard.down(mod);
        await this.page.keyboard.press(cmd.key as any);
        for (const mod of [...cmd.modifiers].reverse())
          await this.page.keyboard.up(mod);
        break;
      }

      // ── File upload ─────────────────────────────────────────────────────

      case "uploadFile": {
        const filePath = this._resolveValue(cmd.filePath);
        /**
         * NOTE: Screenshot after uploadFile to confirm file name appears
         * in the UI (e.g. file chip / preview).
         */
        const input = await this.page.$(cmd.selector);
        if (!input) throw new Error(`File input not found: ${cmd.selector}`);
        await (input as any).uploadFile(filePath);
        break;
      }

      // ── Wait ────────────────────────────────────────────────────────────

      case "wait": {
        await new Promise((r) => setTimeout(r, cmd.ms));
        break;
      }

      case "waitForSelector": {
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: cmd.timeout ?? 15_000,
        });
        break;
      }

      case "waitForNavigation": {
        /**
         * NOTE: Screenshot immediately after waitForNavigation to capture
         * the landed page state — helpful for debugging redirect chains.
         */
        await this.page.waitForNavigation({
          waitUntil: cmd.waitUntil ?? "networkidle2",
          timeout: cmd.timeout ?? 15_000,
        });
        break;
      }

      // ── Context / state ─────────────────────────────────────────────────

      case "readValue": {
        /**
         * NOTE: Screenshot after readValue useful to confirm which element
         * was read, especially for dynamically rendered text values.
         */
        await this.page.waitForSelector(cmd.selector, { timeout: 8_000 });
        const raw = await this.page.$eval(
          cmd.selector,
          (el: Element) =>
            (el as HTMLInputElement).value ||
            (el as HTMLElement).innerText ||
            el.textContent ||
            ""
        );
        this.context[cmd.saveAs] = raw.trim();
        this.log(`Context saved: ${cmd.saveAs} = "${this.context[cmd.saveAs]}"`);
        break;
      }

      case "setContext": {
        this.context[cmd.key] = this._resolveValue(cmd.value);
        this.log(`Context set: ${cmd.key} = "${this.context[cmd.key]}"`);
        break;
      }

      // ── Screenshot ──────────────────────────────────────────────────────

      case "screenshot": {
        /**
         * Add screenshot commands at key checkpoints:
         *   - After each major step completes
         *   - Before form submission (review_and_submit_step)
         *   - After navigation lands
         *   - On optional error paths { optional: true }
         *   - After file uploads
         *   - After readValue to confirm what was captured
         */
        const path = cmd.path ?? `screenshot_${Date.now()}.png`;
        await this.page.screenshot({ path, fullPage: cmd.fullPage ?? false });
        this.log(`Screenshot saved: ${path}`);
        break;
      }

      // ── Conditional branch ───────────────────────────────────────────────

      case "if": {
        const val1 = this._resolveValue(cmd.value1);
        const val2 = this._resolveValue(cmd.value2);
        let isTrue = false;

        switch (cmd.operator) {
          case "equals":      isTrue = val1 === val2; break;
          case "notEquals":   isTrue = val1 !== val2; break;
          case "contains":    isTrue = val1.includes(val2); break;
          case "notContains": isTrue = !val1.includes(val2); break;
        }

        const branch = isTrue ? cmd.then : cmd.otherwise;
        if (branch) {
          for (const sub of branch) await this._executeCommand(sub);
        }
        break;
      }

      default: {
        // Exhaustiveness guard
        const _: never = cmd;
        throw new Error(`Unknown action: "${(_ as Command).action}"`);
      }
    }
  }

  /** Execute a command with retry + optional-failure support */
  private async _executeWithRetry(cmd: Command): Promise<void> {
    const retries = cmd.retries ?? this.options.commandRetries;
    const delay = cmd.retryDelay ?? this.options.retryDelay;
    let lastError: Error | undefined;

    for (let attempt = 1; attempt <= retries; attempt++) {
      try {
        await this._executeCommand(cmd);
        return;
      } catch (e: unknown) {
        lastError = e as Error;
        this.log(
          `  ⚠ "${cmd.action}" attempt ${attempt}/${retries} failed: ${lastError.message}`
        );
        if (attempt < retries) {
          await new Promise((r) => setTimeout(r, delay));
        }
      }
    }

    if (cmd.optional) {
      this.log(`  ↩ Optional "${cmd.action}" skipped after ${retries} attempts`);
      return;
    }

    const sel = "selector" in cmd ? ` on "${(cmd as any).selector}"` : "";
    throw new Error(
      `Command "${cmd.action}"${sel} failed after ${retries} attempts.\nLast error: ${lastError?.message}`
    );
  }

  // ─────────────────────────────────────────────
  // MAIN ENGINE LOOP
  // ─────────────────────────────────────────────

  async run(): Promise<EngineRunResult> {
    this.log("Engine started — waiting for page load");

    // Always wait for page to be fully loaded before first condition scan
    await this._waitForPageLoad();

    const mainLoop = this._runLoop();

    let loopResult: Awaited<ReturnType<typeof this._runLoop>>;

    if (this.options.runTimeout > 0) {
      let timeoutId: ReturnType<typeof setTimeout>;
      const timeoutPromise = new Promise<{ __timedOut: true }>((resolve) => {
        timeoutId = setTimeout(
          () => resolve({ __timedOut: true }),
          this.options.runTimeout
        );
      });
      const raced = await Promise.race([mainLoop, timeoutPromise]);
      clearTimeout(timeoutId!);

      if (raced != null && "__timedOut" in raced) {
        this.log(`⏱ runTimeout (${this.options.runTimeout}ms) hit — status:unknown`);
        return { status: "unknown", context: this.context, reason: "timeout" };
      }
      loopResult = raced as Awaited<ReturnType<typeof this._runLoop>>;
    } else {
      loopResult = await mainLoop;
    }

    // Hard failure from loop (infinite-loop guard or max-iterations)
    if (loopResult?.status === "failed") {
      return loopResult as EngineRunResult;
    }

    // success was returned inline from the loop (successCondition matched)
    if (loopResult?.status === "success") {
      return loopResult as EngineRunResult;
    }

    // Loop exited with no successCondition match
    if (this.options.successCondition) {
      this.log("⚠ successCondition not matched after loop — status:unknown");
      return {
        status: "unknown",
        context: this.context,
        reason: "successCondition_not_matched",
      };
    }

    // No successCondition configured — legacy: clean exit = success
    this.log("Engine finished. status:success");
    return { status: "success", context: this.context };
  }

  /**
   * Inner polling loop — fires every `pollInterval` ms.
   *
   * Each tick:
   *   1. Check successCondition first — match → return success immediately.
   *   2. Scan steps in priority order — first matching step executes then
   *      loop restarts from top (re-scan after every action).
   *   3. No match → wait pollInterval → retry.
   *   4. N consecutive ticks with no match → exit (clean).
   *
   * Returns:
   *   { status:'success', context }  — successCondition matched mid-loop
   *   { status:'failed',  error }    — infinite-loop guard or max-iterations
   *   undefined                      — clean exit (run() decides outcome)
   */
  private async _runLoop(): Promise<
    | { status: "success"; context: Record<string, string> }
    | { status: "failed"; error: string; context: Record<string, string> }
    | undefined
  > {
    let remaining = this.options.maxIterations;

    while (remaining-- > 0) {

      // ── 1. Check successCondition first (highest priority) ───────────────
      if (this.options.successCondition) {
        const won = await this._evaluateCondition(this.options.successCondition);
        if (won) {
          this.log("✅ successCondition matched — status:success");
          return { status: "success", context: this.context };
        }
      }

      // ── 2. Scan steps ────────────────────────────────────────────────────
      let matched = false;

      for (const step of this.steps) {
        // Skip already-executed non-repeatable steps
        if (!step.repeatable && this.executedSteps.has(step.id)) continue;

        const isMatch = await this._evaluateCondition(step.condition);
        if (!isMatch) continue;

        // ── Infinite-loop guard ────────────────────────────────────────────
        this.consecutiveMatchCount[step.id] =
          (this.consecutiveMatchCount[step.id] ?? 0) + 1;

        if (
          this.consecutiveMatchCount[step.id] >
          this.options.consecutiveMatchLimit
        ) {
          const msg =
            `Infinite loop: step "${step.id}" matched ` +
            `${this.consecutiveMatchCount[step.id]}x consecutively. ` +
            `Ensure commands advance the form state.`;
          this.log(`❌ ${msg}`);
          return { status: "failed", error: msg, context: this.context };
        }

        this.log(
          `✅ Step matched: "${step.id}" (priority: ${step.priority ?? 0})`
        );

        // ── Execute commands ───────────────────────────────────────────────
        try {
          for (const cmd of step.commands) {
            const sel = "selector" in cmd ? ` [${(cmd as any).selector}]` : "";
            const val = "value" in cmd ? ` = "${(cmd as any).value}"` : "";
            this.log(`  → ${cmd.action}${sel}${val}`);
            await this._executeWithRetry(cmd);
          }
        } catch (e: unknown) {
          const msg = (e as Error).message;
          this.log(`❌ Step "${step.id}" aborted: ${msg}`);
          return { status: "failed", error: msg, context: this.context };
        }

        // Mark done
        if (!step.repeatable) this.executedSteps.add(step.id);

        // Reset consecutive counter for this step (commands ran OK)
        this.consecutiveMatchCount[step.id] = 0;
        // Reset all other counters — page state changed
        for (const key of Object.keys(this.consecutiveMatchCount)) {
          if (key !== step.id) this.consecutiveMatchCount[key] = 0;
        }

        matched = true;
        break; // Re-scan from top after every matched step
      }

      // ── 3. Poll wait ─────────────────────────────────────────────────────
      // If a step ran, skip the wait and immediately re-scan (step may have
      // triggered a page change we can react to right away).
      // If nothing matched, wait pollInterval before next tick.
      if (!matched) {
        this.log(
          `⏳ No match this tick — waiting ${this.options.pollInterval}ms`
        );
        await new Promise((r) => setTimeout(r, this.options.pollInterval));
      }
    }

    if (remaining <= 0) {
      const msg = `Engine hit maxIterations (${this.options.maxIterations}). Form may be stuck.`;
      this.log(`❌ ${msg}`);
      return { status: "failed", error: msg, context: this.context };
    }

    this.log("Loop finished cleanly. context:", this.context);
    return undefined;
  }
}
