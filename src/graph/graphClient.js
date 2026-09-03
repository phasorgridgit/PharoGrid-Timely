import { Client } from "@microsoft/microsoft-graph-client";

// Mock mode: no Microsoft sign-in, no SharePoint, nothing saved anywhere.
// Set VITE_WORKHUB_MOCK_MODE=true in .env to try the app locally before
// SharePoint/Entra ID are set up. See every graph/*.js file for the branch
// that uses in-memory data (mockStore.js) instead of a real Graph call.
export const isMockMode = import.meta.env.VITE_WORKHUB_MOCK_MODE === "true";

// Every Graph call in the renderer goes through this client. The token comes
// from the Electron main process (see electron/main.js) — the renderer never
// handles secrets, only a short-lived access token for the signed-in user.
const authProvider = {
  getAccessToken: async () => {
    let token = await window.workhub.getToken();
    if (!token) {
      token = await window.workhub.login(); // triggers interactive sign-in if needed
    }
    return token.accessToken;
  }
};

export const graphClient = Client.initWithMiddleware({ authProvider });

let cachedSiteId = null;

/**
 * Resolves the WorkHub SharePoint site ID once per session and caches it.
 * Reads WORKHUB_SITE_HOSTNAME / WORKHUB_SITE_PATH from the renderer's env
 * (injected at build time by Vite — see vite.config.js `define`, or read
 * from a small /config.json shipped alongside the app).
 */
export async function getSiteId() {
  if (cachedSiteId) return cachedSiteId;
  const hostname = import.meta.env.VITE_WORKHUB_SITE_HOSTNAME;
  const sitePath = import.meta.env.VITE_WORKHUB_SITE_PATH;
  const site = await graphClient.api(`/sites/${hostname}:${sitePath}`).get();
  cachedSiteId = site.id;
  return cachedSiteId;
}
