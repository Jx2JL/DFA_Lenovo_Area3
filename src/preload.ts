import { contextBridge, ipcRenderer } from "electron";

contextBridge.exposeInMainWorld("coach", {
  onShowTip: (callback: (message: string) => void) => {
    ipcRenderer.on("show-tip", (_event, message: string) => callback(message));
  },
  onHideTip: (callback: () => void) => {
    ipcRenderer.on("hide-tip", () => callback());
  },
});
