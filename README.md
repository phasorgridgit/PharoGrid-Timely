# PhasorGrid Timely

Desktop time-tracking, timesheets, billing and analytics app for PhasorGrid Engineering Ltd.
Electron + React front end. No app database — Microsoft Graph (SharePoint Lists + OneDrive) is the system of record.

## Quick start

```bash
npm install
cp .env.example .env        # fill in tenant/client IDs and site path — see Build Guide
npm run dev
```

See the full **PhasorGrid Timely — Build Guide & Source Code** PDF for:
- Registering the Entra ID app and granting Graph permissions
- Provisioning the SharePoint site (two options below)
- What each file in `src/` does and how to extend it
- Packaging installers with `npm run build:win` / `build:mac` / `build:linux`

## Provisioning the SharePoint site

Two interchangeable options — run either one, once per environment, after the
"WorkHub" SharePoint site itself exists (the backend identifier is retained for compatibility):

- **`scripts/provision-sharepoint.ps1`** (PnP PowerShell) — needs the account
  running it to be a Site Collection Admin on the WorkHub SharePoint site, and needs
  the tenant to allow user consent to the "PnP Management Shell" app. Fine
  for most tenants.
- **`npm run provision`** (`scripts/provision-sharepoint-graph.mjs`) — calls
  Microsoft Graph directly with app-only (client-credentials) auth. Use this
  if PnP is blocked by Conditional Access or admin-consent policy, or if you'd
  rather not grant anyone per-site SharePoint admin rights. Needs a one-time
  `Sites.Manage.All` application-permission grant on the Entra ID app (see
  comments at the top of the script) and these env vars:
  `WORKHUB_TENANT_ID`, `WORKHUB_CLIENT_ID`, `WORKHUB_CLIENT_SECRET`,
  `WORKHUB_SITE_HOSTNAME`, `WORKHUB_SITE_PATH`.

After either one, run **`scripts/seed-sharepoint-data.ps1`** (PnP PowerShell)
to populate starter Clients, Projects, and Team Members as real list items —
edit the arrays at the top of the script to match your org first. It also
seeds one `WorkHub_TimeOffBalances` row per teammate, resolving each person's
real Entra object id via `Get-PnPAzureADUser` so it lines up with what the
app itself writes. Safe to re-run — it looks each row up by Title first and
skips anything already there.

## Packaging for the org

`npm run build:win` (etc.) bundles `dist/**/*` and `electron/**/*` into the
installer, plus `.env` as an extra resource (see `build.extraResources` in
`package.json`) — `electron/main.js` loads it from `process.resourcesPath` in
a packaged app. Without this, the installer authenticates fine in `npm run dev`
but silently fails to sign in for every user once installed, because `.env`
never shipped. For 12 users, drop the installer on a shared drive or Teams
channel with a one-page "first run" note (sign in with your work account,
accept the consent screen — already admin-approved).

## Project layout

```
electron/         Main process: window + MSAL sign-in (per-user token)
src/graph/         Microsoft Graph service layer (one file per SharePoint list)
src/pages/          One React page per app section
src/components/  Shared UI (sidebar, etc.)
scripts/               SharePoint provisioning: PnP PowerShell, or Graph app-only (provision-sharepoint-graph.mjs)
```
