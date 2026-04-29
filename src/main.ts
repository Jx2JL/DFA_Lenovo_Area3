import { app, BrowserWindow, screen } from "electron";
import * as path from "path";
import { exec, execSync } from "child_process";
// fs removed — no longer needed for VBScript helper
import {
  uIOhook,
  UiohookMouseEvent,
  UiohookWheelEvent,
  UiohookKeyboardEvent,
} from "uiohook-napi";

// uiohook keycodes for arrow keys
const ARROW_KEY_CODES = new Set([
  0xe048, // Up
  0xe050, // Down
  0xe04b, // Left
  0xe04d, // Right
]);

// Keycodes that indicate scrolling behavior
const SCROLL_KEY_CODES = new Set([
  0xe049, // Page Up
  0xe051, // Page Down
  0xe047, // Home
  0xe04f, // End
]);

// ─── Configuration ───────────────────────────────────────────────────────────

const TOAST_WIDTH = 360;
const TOAST_HEIGHT = 90;
const MARGIN = 16;
const TOOLTIP_DURATION_MS = 3000;
const COOLDOWN_MS = 5 * 60 * 1000;
const POLL_INTERVAL_MS = 2000; // 2s — plenty fast, doesn't lag
const SETTLE_DELAY_MS = 3000;

// ─── Coaching messages ───────────────────────────────────────────────────────

interface CoachTip {
  id: string;
  message: string;
  lastShown: number;
  shownThisSession: boolean;
}

const tips: Record<string, CoachTip> = {
  cellNav: {
    id: "cellNav",
    message:
      "Nudge the TrackPoint to move between cells.\nIt works just like arrow keys.",
    lastShown: 0,
    shownThisSession: false,
  },
  scroll: {
    id: "scroll",
    message:
      "Hold the center button and nudge the TrackPoint.\nYou can scroll in any direction — up, down, or sideways.",
    lastShown: 0,
    shownThisSession: false,
  },
};

// ─── State ───────────────────────────────────────────────────────────────────

let overlayWindow: BrowserWindow | null = null;
let excelIsActive = false;
let workbookIsOpen = false;
let coachingReady = false;
let settleTimer: ReturnType<typeof setTimeout> | null = null;
let scrollTipTimer: ReturnType<typeof setTimeout> | null = null;
let lastCursorX = 0;
let lastCursorY = 0;
let hideTimer: ReturnType<typeof setTimeout> | null = null;
let pollInProgress = false; // prevent overlapping polls

// ─── Windows detection helpers ──────────────────────────────────────────────
// Use lightweight cmd commands instead of PowerShell or VBScript.
// `tasklist` checks if Excel is running. `wmic` gets the window title to
// determine if a workbook is open (title contains " - " when a file is open).

// ─── Active-window detection (cross-platform, async) ────────────────────────

function pollWindowStateAsync(): void {
  if (pollInProgress) return; // skip if last poll hasn't finished
  pollInProgress = true;

  if (process.platform === "darwin") {
    pollMacOS();
  } else if (process.platform === "win32") {
    pollWindows();
  } else {
    pollInProgress = false;
  }
}

function pollMacOS(): void {
  // Check if Excel is frontmost
  exec(
    `osascript -e 'tell application "System Events" to get name of first application process whose frontmost is true'`,
    { encoding: "utf-8", timeout: 2000 },
    (err, stdout) => {
      const name = (stdout || "").trim().toLowerCase();
      const wasActive = excelIsActive;
      excelIsActive =
        name.includes("excel") ||
        name.includes("xlmain") ||
        name.includes(".xlsx");

      if (excelIsActive && !wasActive)
        console.log(`[notify] Excel detected: "${name}"`);
      else if (!excelIsActive && wasActive)
        console.log("[notify] Excel lost focus");

      if (!excelIsActive) {
        updateWorkbookState(false);
        pollInProgress = false;
        return;
      }

      // Check for active workbook
      exec(
        `osascript -e 'tell application "Microsoft Excel" to get name of active workbook'`,
        { encoding: "utf-8", timeout: 2000 },
        (err2, stdout2) => {
          const wb = (stdout2 || "").trim();
          updateWorkbookState(wb.length > 0 && !err2);
          pollInProgress = false;
        }
      );
    }
  );
}

function pollWindows(): void {
  // Step 1: Check if Excel.exe is running at all (fast, lightweight)
  exec(
    `tasklist /FI "IMAGENAME eq EXCEL.EXE" /FO CSV /NH`,
    { encoding: "utf-8", timeout: 3000 },
    (err, stdout) => {
      const output = (stdout || "").trim();
      const wasActive = excelIsActive;

      if (!output.toLowerCase().includes("excel.exe")) {
        // Excel isn't running at all
        excelIsActive = false;
        if (wasActive) console.log("[notify] Excel not running");
        updateWorkbookState(false);
        pollInProgress = false;
        return;
      }

      // Excel is running — now check the window title for a workbook
      exec(
        `wmic process where "name='EXCEL.EXE'" get CommandLine 2>nul & for /f "tokens=*" %a in ('wmic path Win32_Process where "name=\'EXCEL.EXE\'" get ProcessId /value 2^>nul') do @echo %a`,
        { encoding: "utf-8", timeout: 3000, shell: "cmd.exe" },
        () => {
          // Use a simpler approach: get Excel window titles via tasklist /V
          exec(
            `tasklist /FI "IMAGENAME eq EXCEL.EXE" /V /FO CSV /NH`,
            { encoding: "utf-8", timeout: 3000 },
            (err3, stdout3) => {
              const titleOutput = (stdout3 || "").trim();
              excelIsActive = true;

              if (!wasActive) console.log(`[notify] Excel detected`);

              // When a workbook is open, the window title contains " - "
              // e.g. "Book1 - Excel" or "MyFile.xlsx - Excel"
              // When on start screen, it's just "Excel" or "Excel Start"
              const hasWorkbook =
                titleOutput.includes(" - Excel") ||
                titleOutput.includes(".xlsx") ||
                titleOutput.includes(".xls,");

              updateWorkbookState(hasWorkbook);
              pollInProgress = false;
            }
          );
        }
      );
    }
  );
}

function updateWorkbookState(isOpen: boolean): void {
  const hadWorkbook = workbookIsOpen;
  workbookIsOpen = isOpen;

  if (workbookIsOpen && !hadWorkbook) {
    console.log("[notify] Workbook detected — settling for 3s");
    markCoachingReady();
  } else if (!workbookIsOpen && hadWorkbook) {
    console.log("[notify] No workbook — coaching paused");
    markCoachingNotReady();
  }
}

// ─── Coaching readiness (settle delay) ───────────────────────────────────────

function markCoachingReady(): void {
  if (settleTimer) clearTimeout(settleTimer);
  settleTimer = setTimeout(() => {
    coachingReady = true;
    console.log("[notify] Coaching active — watching for cell interactions");
  }, SETTLE_DELAY_MS);
}

function markCoachingNotReady(): void {
  if (settleTimer) clearTimeout(settleTimer);
  settleTimer = null;
  coachingReady = false;
}

function excelSpreadsheetIsActive(): boolean {
  return excelIsActive && workbookIsOpen && coachingReady;
}

// ─── Toast window — pinned to bottom-right ───────────────────────────────────

function createToastWindow(): BrowserWindow {
  const { workArea } = screen.getPrimaryDisplay();
  const x = workArea.x + workArea.width - TOAST_WIDTH - MARGIN;
  const y = workArea.y + workArea.height - TOAST_HEIGHT - MARGIN;

  const win = new BrowserWindow({
    width: TOAST_WIDTH,
    height: TOAST_HEIGHT,
    x,
    y,
    frame: false,
    transparent: true,
    alwaysOnTop: true,
    skipTaskbar: true,
    resizable: false,
    hasShadow: false,
    focusable: false,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  win.setIgnoreMouseEvents(true, { forward: true });
  win.loadFile(path.join(__dirname, "renderer", "index.html"));
  win.hide();

  return win;
}

// ─── Show / hide toast ───────────────────────────────────────────────────────

function showTip(tip: CoachTip): void {
  if (!overlayWindow) return;

  const now = Date.now();
  if (tip.shownThisSession && now - tip.lastShown < COOLDOWN_MS) return;

  tip.lastShown = now;
  tip.shownThisSession = true;

  console.log(`[notify] Showing tip: ${tip.id}`);

  const { workArea } = screen.getPrimaryDisplay();
  const x = workArea.x + workArea.width - TOAST_WIDTH - MARGIN;
  const y = workArea.y + workArea.height - TOAST_HEIGHT - MARGIN;
  overlayWindow.setBounds({ x, y, width: TOAST_WIDTH, height: TOAST_HEIGHT });

  overlayWindow.webContents.send("show-tip", tip.message);
  overlayWindow.showInactive();

  if (hideTimer) clearTimeout(hideTimer);
  hideTimer = setTimeout(() => {
    overlayWindow?.webContents.send("hide-tip");
    setTimeout(() => overlayWindow?.hide(), 400);
  }, TOOLTIP_DURATION_MS);

  // After showing the cellNav tip, automatically queue the scroll tip
  if (tip.id === "cellNav" && !tips.scroll.shownThisSession) {
    if (scrollTipTimer) clearTimeout(scrollTipTimer);
    scrollTipTimer = setTimeout(() => {
      console.log("[notify] Auto-showing scroll tip");
      showTip(tips.scroll);
    }, TOOLTIP_DURATION_MS + 1500); // show 1.5s after cellNav fades out
  }
}

// ─── Input hooks ─────────────────────────────────────────────────────────────

function startInputHooks(): void {
  uIOhook.on("mousemove", (evt: UiohookMouseEvent) => {
    lastCursorX = evt.x;
    lastCursorY = evt.y;
  });

  uIOhook.on("keydown", (evt: UiohookKeyboardEvent) => {
    if (!excelSpreadsheetIsActive()) return;
    if (ARROW_KEY_CODES.has(evt.keycode)) {
      showTip(tips.cellNav);
    }
    // Page Up/Down, Home, End → scroll tip
    if (SCROLL_KEY_CODES.has(evt.keycode)) {
      console.log(`[notify] Scroll key detected: 0x${evt.keycode.toString(16)}`);
      showTip(tips.scroll);
    }
  });

  uIOhook.on("wheel", (_evt: UiohookWheelEvent) => {
    console.log("[notify] Wheel event detected");
    if (!excelSpreadsheetIsActive()) return;
    showTip(tips.scroll);
  });

  // Middle mouse button (TrackPoint scroll button) — show scroll tip
  uIOhook.on("mousedown", (_evt: UiohookMouseEvent) => {
    // button 2 = middle click (the TrackPoint center button)
    if (_evt.button === 2) {
      console.log("[notify] Middle button detected");
      if (!excelSpreadsheetIsActive()) return;
      showTip(tips.scroll);
      return;
    }
    if (!excelSpreadsheetIsActive()) return;
    showTip(tips.cellNav);
  });

  uIOhook.start();
}

// ─── Polling ─────────────────────────────────────────────────────────────────

function startWindowPolling(): void {
  // Run first check immediately
  pollWindowStateAsync();
  // Then poll on interval
  setInterval(() => pollWindowStateAsync(), POLL_INTERVAL_MS);
}

// ─── App lifecycle ───────────────────────────────────────────────────────────

app.whenReady().then(() => {
  // Detection now uses tasklist/wmic — no setup needed
  overlayWindow = createToastWindow();
  startWindowPolling();
  startInputHooks();
  console.log("TrackPoint Notify is running. Open Excel to see coaching toasts.");
});

app.on("window-all-closed", () => {
  uIOhook.stop();
  app.quit();
});

app.on("before-quit", () => {
  uIOhook.stop();
});
