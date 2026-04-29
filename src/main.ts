import { app, BrowserWindow, screen } from "electron";
import * as path from "path";
import { exec, execSync } from "child_process";
import * as fs from "fs";
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
let lastCursorX = 0;
let lastCursorY = 0;
let hideTimer: ReturnType<typeof setTimeout> | null = null;
let pollInProgress = false; // prevent overlapping polls

// Path for the temporary VBScript helper (Windows only)
let vbsHelperPath = "";

// ─── Windows helper script ──────────────────────────────────────────────────
// Instead of spawning PowerShell (very slow on Windows), we write a tiny
// VBScript once at startup and call it with cscript, which is near-instant.

function createWindowsHelper(): void {
  if (process.platform !== "win32") return;

  vbsHelperPath = path.join(app.getPath("temp"), "trackpoint-detect.vbs");

  // This VBScript:
  //   1. Uses Win32 API via Shell to get the foreground window title
  //   2. Checks if Excel is running and has a workbook open
  const vbs = `
On Error Resume Next
Set objShell = CreateObject("WScript.Shell")
Set objExcel = Nothing

' Get foreground window title
strTitle = objShell.AppActivate("zzz_nonexistent_window_zzz")
' AppActivate trick doesn't work for reading — use Excel COM instead

' Check if Excel is running and get active workbook
Set objExcel = GetObject(, "Excel.Application")
If Err.Number = 0 And Not objExcel Is Nothing Then
  If objExcel.Workbooks.Count > 0 Then
    WScript.Echo "EXCEL_WORKBOOK:" & objExcel.ActiveWorkbook.Name
  Else
    WScript.Echo "EXCEL_NO_WORKBOOK"
  End If
  Set objExcel = Nothing
Else
  WScript.Echo "NO_EXCEL"
End If
`.trim();

  fs.writeFileSync(vbsHelperPath, vbs, "utf-8");
}

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
  // Use the lightweight VBScript helper via cscript
  exec(
    `cscript //NoLogo //T:3 "${vbsHelperPath}"`,
    { encoding: "utf-8", timeout: 4000 },
    (err, stdout) => {
      const output = (stdout || "").trim();
      const wasActive = excelIsActive;

      if (output.startsWith("EXCEL_WORKBOOK:")) {
        excelIsActive = true;
        if (!wasActive) console.log(`[notify] Excel detected: "${output}"`);
        updateWorkbookState(true);
      } else if (output === "EXCEL_NO_WORKBOOK") {
        excelIsActive = true;
        if (!wasActive) console.log("[notify] Excel detected (no workbook)");
        updateWorkbookState(false);
      } else {
        excelIsActive = false;
        if (wasActive) console.log("[notify] Excel lost focus");
        updateWorkbookState(false);
      }

      pollInProgress = false;
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
  // Run first check immediately
  pollWindowStateAsync();
  // Then poll on interval
  setInterval(() => pollWindowStateAsync(), POLL_INTERVAL_MS);
}

// ─── App lifecycle ───────────────────────────────────────────────────────────

app.whenReady().then(() => {
  createWindowsHelper(); // write VBScript helper (no-op on macOS)
  overlayWindow = createToastWindow();
  startWindowPolling();
  startInputHooks();
  console.log("TrackPoint Notify is running. Open Excel to see coaching toasts.");
});

app.on("window-all-closed", () => {
  uIOhook.stop();
  // Clean up temp VBScript
  if (vbsHelperPath) try { fs.unlinkSync(vbsHelperPath); } catch {}
  app.quit();
});

app.on("before-quit", () => {
  uIOhook.stop();
});
