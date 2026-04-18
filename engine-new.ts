import { Page } from "puppeteer-core";

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

export interface StepCondition {
  exists?: string[];
  notExists?: string[];
  urlContains?: string;
  textContains?: string;
}

interface BaseCommand {
  retries?: number;
  retryDelay?: number;
  optional?: boolean;
}

export interface TypeCommand extends BaseCommand {
  action: "type";
  selector: string;
  value: string;
  typeDelay?: number;
}
export interface ClickCommand extends BaseCommand {
  action: "click";
  selector: string;
}
export interface SelectCommand extends BaseCommand {
  action: "select";
  selector: string;
  value: string;
}
export interface CheckCommand extends BaseCommand {
  action: "check";
  selector: string;
}
export interface UncheckCommand extends BaseCommand {
  action: "uncheck";
  selector: string;
}
export interface RadioCommand extends BaseCommand {
  action: "radio";
  selector: string;
}
export interface FocusCommand extends BaseCommand {
  action: "focus";
  selector: string;
}
export interface HoverCommand extends BaseCommand {
  action: "hover";
  selector: string;
}
export interface ScrollToCommand extends BaseCommand {
  action: "scrollTo";
  selector: string;
}
export interface ScrollByCommand extends BaseCommand {
  action: "scrollBy";
  x?: number;
  y?: number;
}
export interface PressKeyCommand extends BaseCommand {
  action: "pressKey";
  key: string;
}
export interface KeyComboCommand extends BaseCommand {
  action: "keyCombo";
  modifiers: Array<"Control" | "Shift" | "Alt" | "Meta">;
  key: string;
}
export interface UploadFileCommand extends BaseCommand {
  action: "uploadFile";
  selector: string;
  filePath: string;
}
export interface ClearAndTypeCommand extends BaseCommand {
  action: "clearAndType";
  selector: string;
  value: string;
  typeDelay?: number;
}
export interface WaitCommand extends BaseCommand {
  action: "wait";
  ms: number;
}
export interface WaitForSelectorCommand extends BaseCommand {
  action: "waitForSelector";
  selector: string;
  timeout?: number;
}
export interface WaitForNavigationCommand extends BaseCommand {
  action: "waitForNavigation";
  waitUntil?: "load" | "domcontentloaded" | "networkidle0" | "networkidle2";
  timeout?: number;
}
export interface ReadValueCommand extends BaseCommand {
  action: "readValue";
  selector: string;
  saveAs: string;
}
export interface SetContextCommand extends BaseCommand {
  action: "setContext";
  key: string;
  value: string;
}
export interface ScreenshotCommand extends BaseCommand {
  action: "screenshot";
  path?: string;
  fullPage?: boolean;
}
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

export interface Step {
  id: string;
  condition: StepCondition;
  commands: Command[];
  priority?: number;
  repeatable?: boolean;
}

export interface FormEngineOptions {
  pollInterval?: number;
  commandRetries?: number;
  retryDelay?: number;
  maxIterations?: number;
  consecutiveMatchLimit?: number;
  debug?: boolean;
  successCondition?: StepCondition;
  runTimeout?: number;
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

  private executedSteps = new Set<string>();
  private context: Record<string, string> = {};
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

  private _resolveValue(value: string): string {
    return value.replace(/\{\{(\w+)\}\}/g, (_, key: string) => {
      if (!(key in this.context)) {
        throw new Error(
          `Context key "{{${key}}}" not found. Available: ${Object.keys(this.context).join(", ") || "(none)"}`,
        );
      }
      return this.context[key];
    });
  }

  // ─────────────────────────────────────────────
  // PAGE LOAD GUARD
  // ─────────────────────────────────────────────

  private async _waitForPageLoad(): Promise<boolean> {
    const LOAD_TIMEOUT = 30_000;
    const MAX_RETRIES = 50;
    const SLEEP_MS = 1_000;
    const sleep = (ms: number) =>
      new Promise<void>((resolve) => setTimeout(resolve, ms));

    let domReady = false;
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      try {
        await this.page.waitForFunction(
          () =>
            document.readyState === "interactive" ||
            document.readyState === "complete",
          { timeout: LOAD_TIMEOUT },
        );
        domReady = true;
        this.log(`Page Loaded (attempt ${attempt})`);
        break;
      } catch {
        this.log(`Page not loaded — attempt ${attempt}/${MAX_RETRIES}`);
        if (attempt < MAX_RETRIES) await sleep(SLEEP_MS);
      }
    }

    if (!domReady) {
      this.log("  ✖ Page never became ready after 50 attempts — aborting");
      return false;
    }
    return true;
  }

  // ─────────────────────────────────────────────
  // VISIBILITY CHECKS
  // ─────────────────────────────────────────────

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

  private async _isAbsent(selector: string): Promise<boolean> {
    const el = await this.page.$(selector);
    if (!el) return true;
    const visible = await this._isVisibleAndInteractable(selector);
    return !visible;
  }

  // ─────────────────────────────────────────────
  // CONDITION SCORING  ← NEW
  // ─────────────────────────────────────────────

  /**
   * Score condition against live page.
   * Each sub-condition that passes adds 1 point.
   * Returns { score, total, pct } where pct = Math.round(score/total*100).
   * total=0 → pct=0 (empty condition = no signal).
   */
  private async _scoreCondition(
    cond: StepCondition,
  ): Promise<{ score: number }> {
    const { exists = [], notExists = [], urlContains, textContains } = cond;
    let score = 0;

    if (urlContains !== undefined) {
      if (this.page.url().includes(urlContains)) score++;
    }

    if (textContains !== undefined) {
      try {
        const html = await this.page.content();
        if (new RegExp(textContains, "i").test(html)) score++;
      } catch {
        /* miss */
      }
    }

    for (const sel of exists) {
      if (await this._isVisibleAndInteractable(sel)) score++;
    }

    for (const sel of notExists) {
      if (await this._isAbsent(sel)) score++;
    }

    return { score };
  }

  /**
   * Legacy full-match check (score === total, total > 0).
   * Used where we need a hard pass/fail (page-load guard etc.).
   */
  // private async _evaluateCondition(cond: StepCondition): Promise<boolean> {
  //   const { score, total } = await this._scoreCondition(cond);
  //   return total > 0 && score === total;
  // }

  // ─────────────────────────────────────────────
  // COMMAND EXECUTION
  // ─────────────────────────────────────────────

  private async _executeCommand(cmd: Command): Promise<void> {
    switch (cmd.action) {
      case "type": {
        const value = this._resolveValue(cmd.value);
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: 8_000,
        });
        await this.page.click(cmd.selector, { clickCount: 3 });
        await this.page.keyboard.press("Backspace");
        await this.page.type(cmd.selector, value, {
          delay: cmd.typeDelay ?? 30,
        });
        break;
      }
      case "clearAndType": {
        const value = this._resolveValue(cmd.value);
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: 8_000,
        });
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
      case "click": {
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: 8_000,
        });
        this.log(`  → Clicking "${cmd.selector}"`);
        await this.page.$eval(cmd.selector, (element) => element.click());
        break;
      }
      case "select": {
        const value = this._resolveValue(cmd.value);
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: 8_000,
        });
        await this.page.select(cmd.selector, value);
        break;
      }
      case "check": {
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: 8_000,
        });
        const checked = await this.page.$eval(
          cmd.selector,
          (el: Element) => (el as HTMLInputElement).checked,
        );
        if (!checked) await this.page.click(cmd.selector);
        break;
      }
      case "uncheck": {
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: 8_000,
        });
        const unchecked = await this.page.$eval(
          cmd.selector,
          (el: Element) => (el as HTMLInputElement).checked,
        );
        if (unchecked) await this.page.click(cmd.selector);
        break;
      }
      case "radio": {
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: 8_000,
        });
        const selected = await this.page.$eval(
          cmd.selector,
          (el: Element) => (el as HTMLInputElement).checked,
        );
        if (!selected) await this.page.click(cmd.selector);
        break;
      }
      case "focus": {
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: 8_000,
        });
        await this.page.focus(cmd.selector);
        break;
      }
      case "hover": {
        await this.page.waitForSelector(cmd.selector, {
          visible: true,
          timeout: 8_000,
        });
        await this.page.hover(cmd.selector);
        break;
      }
      case "scrollTo": {
        await this.page.waitForSelector(cmd.selector, { timeout: 8_000 });
        await this.page.$eval(cmd.selector, (el: Element) =>
          el.scrollIntoView({ behavior: "smooth", block: "center" }),
        );
        break;
      }
      case "scrollBy": {
        await this.page.mouse.wheel({ deltaX: cmd.x ?? 0, deltaY: cmd.y ?? 0 });
        break;
      }
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
      case "uploadFile": {
        const filePath = this._resolveValue(cmd.filePath);
        const input = await this.page.$(cmd.selector);
        if (!input) throw new Error(`File input not found: ${cmd.selector}`);
        await (input as any).uploadFile(filePath);
        break;
      }
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
        await this.page.waitForNavigation({
          waitUntil: cmd.waitUntil ?? "networkidle2",
          timeout: cmd.timeout ?? 15_000,
        });
        break;
      }
      case "readValue": {
        await this.page.waitForSelector(cmd.selector, { timeout: 8_000 });
        const raw = await this.page.$eval(
          cmd.selector,
          (el: Element) =>
            (el as HTMLInputElement).value ||
            (el as HTMLElement).innerText ||
            el.textContent ||
            "",
        );
        this.context[cmd.saveAs] = raw.trim();
        this.log(
          `Context saved: ${cmd.saveAs} = "${this.context[cmd.saveAs]}"`,
        );
        break;
      }
      case "setContext": {
        this.context[cmd.key] = this._resolveValue(cmd.value);
        this.log(`Context set: ${cmd.key} = "${this.context[cmd.key]}"`);
        break;
      }
      case "screenshot": {
        const path = cmd.path ?? `screenshot_${Date.now()}.png`;
        await this.page.screenshot({ path, fullPage: cmd.fullPage ?? false });
        this.log(`Screenshot saved: ${path}`);
        break;
      }
      case "if": {
        const val1 = this._resolveValue(cmd.value1);
        const val2 = this._resolveValue(cmd.value2);
        let isTrue = false;
        switch (cmd.operator) {
          case "equals":
            isTrue = val1 === val2;
            break;
          case "notEquals":
            isTrue = val1 !== val2;
            break;
          case "contains":
            isTrue = val1.includes(val2);
            break;
          case "notContains":
            isTrue = !val1.includes(val2);
            break;
        }
        const branch = isTrue ? cmd.then : cmd.otherwise;
        if (branch) for (const sub of branch) await this._executeCommand(sub);
        break;
      }
      default: {
        const _: never = cmd;
        throw new Error(`Unknown action: "${(_ as Command).action}"`);
      }
    }
  }

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
          `  ⚠ "${cmd.action}" attempt ${attempt}/${retries} failed: ${lastError.message}`,
        );
        if (attempt < retries) await new Promise((r) => setTimeout(r, delay));
      }
    }

    if (cmd.optional) {
      this.log(
        `  ↩ Optional "${cmd.action}" skipped after ${retries} attempts`,
      );
      return;
    }

    const sel = "selector" in cmd ? ` on "${(cmd as any).selector}"` : "";
    throw new Error(
      `Command "${cmd.action}"${sel} failed after ${retries} attempts.\nLast error: ${lastError?.message}`,
    );
  }

  // ─────────────────────────────────────────────
  // MAIN ENGINE LOOP
  // ─────────────────────────────────────────────

  async run(): Promise<EngineRunResult> {
    this.log("Engine started — waiting for page load");

    const loaded = await this._waitForPageLoad();
    if (!loaded) {
      return {
        status: "failed",
        reason: "Failed to load page",
        context: this.context,
      };
    }

    const mainLoop = await this._runLoop();

    if (mainLoop?.status === "failed") return mainLoop as EngineRunResult;
    if (mainLoop?.status === "success") return mainLoop as EngineRunResult;

    if (this.options.successCondition) {
      this.log("⚠ successCondition not matched after loop — status:unknown");
      return {
        status: "unknown",
        context: this.context,
        reason: "successCondition_not_matched",
      };
    }

    this.log("Engine finished. status:success");
    return { status: "success", context: this.context };
  }

  /**
   * SCORE-BASED LOOP  ← completely rewritten
   *
   * Each of 50 cycles (1 s apart):
   *   1. Score ALL eligible flow steps → sort by pct desc
   *   2. Score successCondition (if set)
   *   3. Top-flow-pct vs success-pct
   *      – success wins (or tie) AND success pct === 100 → return success
   *      – flow wins AND top-step pct === 100              → run commands
   *      – nothing at 100 → wait & continue
   *
   * "Wins" = higher pct. Tie → success wins (conservative).
   */
  private async _runLoop(): Promise<
    | { status: "success"; context: Record<string, string> }
    | { status: "failed"; error: string; context: Record<string, string> }
    | undefined
  > {
    const CYCLES = 50;
    const PAUSE_MS = 1_000;

    let run = true;

    while (run) {
      for (let cycle = 1; cycle <= CYCLES; cycle++) {
        this.log(`── Cycle ${cycle}/${CYCLES} ──`);

        // ── Score successCondition ─────────────────────────────────────────
        let successPct = 0;
        if (this.options.successCondition) {
          const s = await this._scoreCondition(this.options.successCondition);
          successPct = s.score;
          this.log(`  successCondition score: ${s.score}`);
        }

        // ── Score all eligible flow steps ──────────────────────────────────
        type ScoredStep = {
          step: Step;
          score: number;
        };
        const scored: ScoredStep[] = [];

        for (const step of this.steps) {
          if (!step.repeatable && this.executedSteps.has(step.id)) continue;
          const s = await this._scoreCondition(step.condition);
          this.log(`  step "${step.id}" score: ${s.score}`);
          scored.push({ step, ...s });
        }

        // Sort by pct desc (priority already baked in step order; stable sort preserves it)
        scored.so
