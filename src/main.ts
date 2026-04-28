import { app, BrowserWindow, screen } from "electron";
import * as path from "path";
import { execSync } from "child_process";
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

// ─── Configuration ───────────────────────────────────────────────────────────

const TOAST_WIDTH = 360;
const TOAST_HEIGHT = 90;
const MARGIN = 16; // gap from screen edges
const TOOLTIP_DURATION_MS = 3000;
const COOLDOWN_MS = 5 * 60 * 1000;
const POLL_INTERVAL_MS = 500;
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
let lastCursorX = 0;
let lastCursorY = 0;
let hideTimer: ReturnType<typeof setTimeout> | null = null;

// ─── Active-window detection (cross-platform) ───────────────────────────────

function getActiveWindowName(): string {
  try {
    if (process.platform === "darwin") {
      return execSync(
        `osascript -e 'tell application "System Events" to get name of first application process whose frontmost is true'`,
        { encoding: "utf-8", timeout: 1000 }
      ).trim();
    } else if (process.platform === "win32") {
      return execSync(
        `powershell -Command "(Get-Process | Where-Object {$_.MainWindowHandle -eq (Add-Type -MemberDefinition '[DllImport(\\\"user32.dll\\\")] public static extern IntPtr GetForegroundWindow();' -Name Win32 -Namespace Temp -PassThru)::GetForegroundWindow()}).MainWindowTitle"`,
        { encoding: "utf-8", timeout: 2000 }
      ).trim();
    }
  } catch {}
  return "";
}

function isExcelForeground(): boolean {
  const lower = getActiveWindowName().toLowerCase();
  return (
    lower.includes("excel") ||
    lower.includes("xlmain") ||
    lower.includes(".xlsx") ||
    lower.includes(".xls")
  );
}

function checkWorkbookOpen(): boolean {
  try {
    if (process.platform === "darwin") {
      const wb = execSync(
        `osascript -e 'tell application "Microsoft Excel" to get name of active workbook'`,
        { encoding: "utf-8", timeout: 1000 }
      ).trim();
      return wb.length > 0;
    } else if (process.platform === "win32") {
      const title = execSync(
        `powershell -Command "(Get-Process excel -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowTitle -ne ''}).MainWindowTitle"`,
        { encoding: "utf-8", timeout: 2000 }
      ).trim();
      return title.includes(" - ");
    }
  } catch {}
  return false;
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
  // Position in the bottom-right corner of the primary display
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

  // Re-position in case display layout changed
  const { workArea } = screen.getPrimaryDisplay();
  const x = workArea.x + workArea.width - TOAST_WIDTH - MARGIN;
  const y = workArea.y + workArea.height - TOAST_HEIGHT - MARGIN;
  overlayWindow.setBounds({ x, y, width: TOAST_WIDTH, height: TOAST_HEIGHT });

  overlayWindow.webContents.send("show-tip", tip.message);
  overlayWindow.showInactive();

  if (hideTimer) clearTimeout(hideTimer);
  hideTimer = setTimeout(() => {
    overlayWindow?.webContents.send("hide-tip");
    setTimeout(() => overlayWindow?.hide(), 400); // wait for slide-out
  }, TOOLTIP_DURATION_MS);
}

// ─── Input hooks ─────────────────────────────────────────────────────────────

function startInputHooks(): void {
  uIOhook.on("mousemove", (evt: UiohookMouseEvent) => {
    lastCursorX = evt.x;
    lastCursorY = evt.y;
  });

  uIOhook.on("mousedown", (_evt: UiohookMouseEvent) => {
    if (!excelSpreadsheetIsActive()) return;
    showTip(tips.cellNav);
  });

  uIOhook.on("keydown", (evt: UiohookKeyboardEvent) => {
    if (!excelSpreadsheetIsActive()) return;
    if (ARROW_KEY_CODES.has(evt.keycode)) {
      showTip(tips.cellNav);
    }
  });

  uIOhook.on("wheel", (_evt: UiohookWheelEvent) => {
    if (!excelSpreadsheetIsActive()) return;
    showTip(tips.scroll);
  });

  uIOhook.start();
}

// ─── Polling ─────────────────────────────────────────────────────────────────

function startWindowPolling(): void {
  setInterval(() => {
    const wasActive = excelIsActive;
    const hadWorkbook = workbookIsOpen;
    excelIsActive = isExcelForeground();
    workbookIsOpen = excelIsActive ? checkWorkbookOpen() : false;

    if (excelIsActive && !wasActive)
      console.log(`[notify] Excel detected: "${getActiveWindowName()}"`);
    else if (!excelIsActive && wasActive)
      console.log(`[notify] Excel lost focus`);

    if (workbookIsOpen && !hadWorkbook) {
      console.log("[notify] Workbook detected — settling for 3s");
      markCoachingReady();
    } else if (!workbookIsOpen && hadWorkbook) {
      console.log("[notify] No workbook — coaching paused");
      markCoachingNotReady();
    }
  }, POLL_INTERVAL_MS);
}

// ─── App lifecycle ───────────────────────────────────────────────────────────

app.whenReady().then(() => {
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
