const { contextBridge, ipcRenderer } = require("electron");

// Only the auth methods are exposed to the renderer — the renderer never
// sees the client secret, tenant config, or MSAL internals directly.
contextBridge.exposeInMainWorld("workhub", {
  login: () => ipcRenderer.invoke("auth:login"),
  getToken: () => ipcRenderer.invoke("auth:getToken"),
  logout: () => ipcRenderer.invoke("auth:logout"),
  onAuthSuccess: (callback) => {
    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on("auth:success", listener);
    return () => ipcRenderer.removeListener("auth:success", listener);
  },

  checkForUpdates: () => ipcRenderer.invoke("update:check"),
  quitAndInstallUpdate: () => ipcRenderer.invoke("update:quitAndInstall"),
  // Returns an unsubscribe function, same pattern as DOM addEventListener,
  // so a React effect can clean it up on unmount without leaking listeners.
  onUpdateStatus: (callback) => {
    const listener = (_event, payload) => callback(payload);
    ipcRenderer.on("update:status", listener);
    return () => ipcRenderer.removeListener("update:status", listener);
  }
});

