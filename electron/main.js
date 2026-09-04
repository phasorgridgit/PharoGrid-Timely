const { app, BrowserWindow, ipcMain, shell, Menu } = require("electron");
const path = require("path");
const { PublicClientApplication, LogLevel } = require("@azure/msal-node");
const { autoUpdater } = require("electron-updater");

// In dev, .env sits next to package.json (default dotenv behavior).
// In a packaged build, electron-builder copies it into resourcesPath via
// "extraResources" (see package.json) — it is NOT inside the asar, so it
// must be loaded from process.resourcesPath explicitly, or auth silently
// fails for every installed user.
require("dotenv").config({
  path: app.isPackaged
    ? path.join(process.resourcesPath, ".env")
    : path.join(__dirname, "..", ".env")
});

const isDev = process.env.NODE_ENV === "development";

// ---- MSAL (delegated / per-user) setup -----------------------------------
// Each employee authenticates as themselves here. The token this produces is
// scoped to that employee's own Microsoft 365 identity — see the Admin &
// Integrations spec, section 5, for why this must stay delegated and not
// app-only for anything the user personally does.
const msalConfig = {
  auth: {
    clientId: process.env.WORKHUB_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.WORKHUB_TENANT_ID}`
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message) => {
        if (level <= LogLevel.Warning) console.log("[msal]", message);
      },
      logLevel: isDev ? LogLevel.Verbose : LogLevel.Warning
    }
  }
};

const pca = new PublicClientApplication(msalConfig);

const GRAPH_SCOPES = [
  "User.Read",
  "Sites.ReadWrite.All",
  "Files.ReadWrite",
  "offline_access"
];

let cachedAccount = null;

function focusMainWindow() {
  if (!mainWindow || mainWindow.isDestroyed()) return;

  if (mainWindow.isMinimized()) mainWindow.restore();

  // The Microsoft callback finishes in the system browser. Explicitly bring
  // Timely back to the foreground so the user does not have to hunt for the
  // Electron window after authentication succeeds.
  mainWindow.show();
  mainWindow.focus();
  mainWindow.moveTop();

  // Windows can sometimes keep focus on the browser for one event loop turn.
  // Re-focus once more on the next tick for a reliable hand-off.
  setTimeout(() => {
    if (!mainWindow || mainWindow.isDestroyed()) return;
    mainWindow.show();
    mainWindow.focus();
    mainWindow.moveTop();
  }, 100);
}

async function acquireTokenInteractive() {
  // Electron has no browser redirect surface, so MSAL Node opens a loopback
  // server on localhost and drives the system browser through the standard
  // Microsoft sign-in page. Microsoft returns the authorization response to
  // that localhost server, after which we immediately foreground Timely.
  const result = await pca.acquireTokenInteractive({
    scopes: GRAPH_SCOPES,
    openBrowser: async (url) => {
      await shell.openExternal(url);
    },
    successTemplate: `
      <html>
        <head><meta charset="utf-8"><title>Timely — Signed in</title></head>
        <body style="font-family:Segoe UI,Arial,sans-serif;background:#0b1724;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0">
          <div style="text-align:center">
            <h1 style="margin:0 0 10px">Signed in successfully</h1>
            <p style="margin:0;color:#9db0c0">Return to Timely — this tab can be closed.</p>
          </div>
        </body>
      </html>`,
    errorTemplate: `
      <html>
        <head><meta charset="utf-8"><title>Timely — Sign-in failed</title></head>
        <body style="font-family:Segoe UI,Arial,sans-serif;background:#0b1724;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0">
          <div style="text-align:center">
            <h1 style="margin:0 0 10px">Sign-in failed</h1>
            <p style="margin:0;color:#9db0c0">Return to Timely and try again.</p>
          </div>
        </body>
      </html>`
  });

  cachedAccount = result.account;
  focusMainWindow();

  // Also notify the renderer explicitly. This is a second, event-based path
  // back into the UI so the login screen updates even if the browser steals
  // focus while the IPC login promise is resolving.
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send("auth:success", {
      account: result.account?.username || ""
    });
  }

  return result;
}

async function acquireTokenSilent() {
  if (!cachedAccount) {
    const accounts = await pca.getTokenCache().getAllAccounts();
    if (accounts.length === 0) return null;
    cachedAccount = accounts[0];
  }
  try {
    return await pca.acquireTokenSilent({ scopes: GRAPH_SCOPES, account: cachedAccount });
  } catch {
    return null; // caller falls back to interactive
  }
}

ipcMain.handle("auth:login", async () => {
  const existing = await acquireTokenSilent();
  const result = existing || (await acquireTokenInteractive());
  return { accessToken: result.accessToken, account: result.account.username, expiresOn: result.expiresOn };
});

ipcMain.handle("auth:getToken", async () => {
  const result = await acquireTokenSilent();
  if (!result) return null;
  return { accessToken: result.accessToken, expiresOn: result.expiresOn };
});

ipcMain.handle("auth:logout", async () => {
  // Defense in depth, matching the renderer-side fix in App.jsx: a failed
  // cache clear should never surface as a rejected promise back to the
  // renderer, because an unhandled rejection there used to abort the
  // whole sign-out flow before any UI state reset — the Sign out button
  // would appear to do nothing. Signing out of the app itself must always
  // succeed from the person's point of view, even if this particular
  // cleanup step didn't.
  try {
    if (cachedAccount) {
      await pca.getTokenCache().removeAccount(cachedAccount);
    }
  } catch (err) {
    console.error("[auth:logout] failed to clear MSAL token cache:", err);
  } finally {
    cachedAccount = null;
  }
  return true;
});

// ---- Window ----------------------------------------------------------------
// No native File/Edit/View/Window/Help menu bar on Windows/Linux — that's
// just Electron's auto-generated default, not part of this app's design,
// and this is a business app with its own in-app navigation. DevTools stays
// reachable via Ctrl+Shift+I / F12 (registered below) even with the menu
// gone, since removing the menu also removes its usual accelerator.
//
// macOS keeps a minimal menu — Cmd+Q (quit) and Cmd+C/V/X/A (copy/paste/
// select-all) come from standard menu roles there, so fully removing the
// menu on Mac breaks those, unlike on Windows/Linux where they're handled
// by the OS/Chromium directly.
if (process.platform === "darwin") {
  Menu.setApplicationMenu(
    Menu.buildFromTemplate([
      { label: app.name, submenu: [{ role: "quit" }] },
      { label: "Edit", submenu: [{ role: "copy" }, { role: "paste" }, { role: "cut" }, { role: "selectAll" }] }
    ])
  );
} else {
  Menu.setApplicationMenu(null);
}

let mainWindow = null;

function createWindow() {
  const win = new BrowserWindow({
    width: 1360,
    height: 860,
    minWidth: 1100,
    icon: path.join(__dirname, "icon.png"),
    minHeight: 700,
    title: "PhasorGrid Timely",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      webviewTag: false
    }
  });
  mainWindow = win;
  win.on("closed", () => { mainWindow = null; });

  // Block the window from navigating anywhere unexpected, and block it from
  // ever opening new Electron windows — the Microsoft sign-in page already
  // opens in the system browser via shell.openExternal, never in-app.
  const allowedOrigin = isDev ? "http://localhost:5173" : "file://";
  win.webContents.on("will-navigate", (event, url) => {
    if (!url.startsWith(allowedOrigin)) event.preventDefault();
  });
  win.webContents.setWindowOpenHandler(() => ({ action: "deny" }));

  win.webContents.on("before-input-event", (event, input) => {
    const isDevToolsShortcut =
      (input.control || input.meta) && input.shift && input.key.toLowerCase() === "i";
    if (isDevToolsShortcut || input.key === "F12") {
      win.webContents.toggleDevTools();
    }
  });

  if (isDev) {
    win.loadURL("http://localhost:5173");
    win.webContents.openDevTools({ mode: "detach" });
  } else {
    win.loadFile(path.join(__dirname, "..", "dist", "index.html"));
  }
}

app.whenReady().then(() => {
  createWindow();
  initAutoUpdate();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// ---- Auto-update -------------------------------------------------------
// Reads the "publish" config in package.json (GitHub Releases) to check for
// a newer published version, download it in the background, and prompt the
// signed-in employee to restart once it's ready. Skipped entirely in dev —
// there is no update feed to check against when running from source, and
// electron-updater errors out immediately if it tries.
function sendUpdateStatus(status, extra = {}) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send("update:status", { status, ...extra });
  }
}

function initAutoUpdate() {
  if (isDev || !app.isPackaged) return;

  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = false; // we ask first — see update:quitAndInstall below

  autoUpdater.on("checking-for-update", () => sendUpdateStatus("checking"));
  autoUpdater.on("update-available", (info) => sendUpdateStatus("available", { version: info.version }));
  autoUpdater.on("update-not-available", () => sendUpdateStatus("up-to-date"));
  autoUpdater.on("error", (err) => sendUpdateStatus("error", { message: err?.message || String(err) }));
  autoUpdater.on("download-progress", (progress) =>
    sendUpdateStatus("downloading", { percent: Math.round(progress.percent) })
  );
  autoUpdater.on("update-downloaded", (info) => sendUpdateStatus("ready", { version: info.version }));

  // Check once on launch, then periodically — this is a long-running desktop
  // app people often leave open for days, so a launch-only check would miss
  // releases published while it's already running.
  autoUpdater.checkForUpdates().catch(() => {});
  setInterval(() => autoUpdater.checkForUpdates().catch(() => {}), 4 * 60 * 60 * 1000);
}

ipcMain.handle("update:check", async () => {
  if (isDev || !app.isPackaged) return { skipped: true };
  await autoUpdater.checkForUpdates().catch(() => {});
  return { skipped: false };
});

ipcMain.handle("update:quitAndInstall", async () => {
  autoUpdater.quitAndInstall();
});
