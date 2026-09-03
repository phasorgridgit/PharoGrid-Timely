// PhasorGrid Workora — SharePoint provisioning via Microsoft Graph (app-only)
//
// Replaces scripts/provision-sharepoint.ps1 (PnP PowerShell) for orgs where
// per-user Site Collection Admin rights or PnP's third-party app consent are
// blocked by tenant policy. This script authenticates with client-credentials
// (a service principal, not a signed-in user) and calls the Graph API
// directly, so it only needs one-time admin consent on an app registration —
// no per-site admin rights, no PnP consent prompt, immune to Conditional
// Access policies that target interactive/legacy clients.
//
// One-time setup (a Global Admin does this once):
//   1. Entra ID admin center -> App registrations -> your WorkHub app
//      (the same clientId used in electron/main.js is fine, or a new one
//      dedicated to provisioning — either works).
//   2. API permissions -> Add a permission -> Microsoft Graph ->
//      Application permissions -> Sites.Manage.All -> Add.
//   3. Click "Grant admin consent for <tenant>".
//   4. Certificates & secrets -> New client secret -> copy the value.
//
// Usage:
//   WORKHUB_TENANT_ID=...      \
//   WORKHUB_CLIENT_ID=...      \
//   WORKHUB_CLIENT_SECRET=...  \
//   WORKHUB_SITE_HOSTNAME=phasorgrid.sharepoint.com \
//   WORKHUB_SITE_PATH=/sites/WorkHub \
//   node scripts/provision-sharepoint-graph.mjs
//
// Run once per environment (Dev / Prod), after the "WorkHub" site itself
// exists. Safe to re-run: list/column creation is skipped if it already
// exists.

import { ConfidentialClientApplication } from "@azure/msal-node";

const {
  WORKHUB_TENANT_ID,
  WORKHUB_CLIENT_ID,
  WORKHUB_CLIENT_SECRET,
  WORKHUB_SITE_HOSTNAME,
  WORKHUB_SITE_PATH
} = process.env;

for (const [name, val] of Object.entries({
  WORKHUB_TENANT_ID,
  WORKHUB_CLIENT_ID,
  WORKHUB_CLIENT_SECRET,
  WORKHUB_SITE_HOSTNAME,
  WORKHUB_SITE_PATH
})) {
  if (!val) {
    console.error(`Missing required env var: ${name}`);
    process.exit(1);
  }
}

const cca = new ConfidentialClientApplication({
  auth: {
    clientId: WORKHUB_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${WORKHUB_TENANT_ID}`,
    clientSecret: WORKHUB_CLIENT_SECRET
  }
});

async function getAppToken() {
  const result = await cca.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"]
  });
  return result.accessToken;
}

async function graphFetch(token, method, urlPath, body) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${urlPath}`, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: body ? JSON.stringify(body) : undefined
  });
  if (!res.ok && res.status !== 409) {
    const text = await res.text();
    throw new Error(`Graph ${method} ${urlPath} -> ${res.status}: ${text}`);
  }
  return res.status === 204 ? null : res.json();
}

// --- Column type helpers (Graph "columns" resource shape) -----------------
const text = () => ({ text: {} });
const indexedText = () => ({ text: {}, indexed: true }); // for columns the app filters on — see ensureColumn below
const note = (allowMultiple = true) => ({ text: { allowMultipleLines: allowMultiple } });
const currency = () => ({ currency: { locale: "en-us" } });
const number = () => ({ number: {} });
const boolean = () => ({ boolean: {} });
const dateTime = () => ({ dateTime: { displayAs: "default", format: "dateTime" } });
const choice = (choices) => ({ choice: { choices, displayAs: "dropDownMenu" } });
const multiChoice = (choices) => ({ choice: { choices, displayAs: "checkBoxes", allowMultipleValues: true } });
const url = () => ({ hyperlinkOrPicture: {} }); // kept for reference — see note below; not used by any app-written field
const person = () => ({ personOrGroup: { allowMultipleSelection: false } }); // kept for reference — same caveat
const personMulti = () => ({ personOrGroup: { allowMultipleSelection: true } });

// NOTE on person()/url(): don't use these for any column the app itself
// writes to. `person()`/`personMulti()` need a site-scoped SharePoint user id
// (from the User Information List), not the AAD object-id string Graph `/me`
// returns — writing that raw id 400s. `url()` needs a { Url, Description }
// object, not the plain string the app already has.
//
// EmployeeName/ProjectName/ClientName are plain text(), not lookup() columns
// — time entries reference the employee, project, and client by NAME, never
// by SharePoint item id or Entra object id. TeamMembers (on WorkHub_Projects)
// is the one exception: it's set by admins directly in the SharePoint UI,
// never written by the app, so a real personMulti column is fine there.

async function ensureList(token, siteId, title, template = "genericList") {
  const existing = await graphFetch(
    token,
    "GET",
    `/sites/${siteId}/lists?$filter=displayName eq '${title}'`
  );
  if (existing?.value?.length) {
    console.log(`  list "${title}" already exists, reusing`);
    return existing.value[0].id;
  }
  const created = await graphFetch(token, "POST", `/sites/${siteId}/lists`, {
    displayName: title,
    list: { template }
  });
  console.log(`  created "${title}"`);
  return created.id;
}

async function ensureColumn(token, siteId, listId, name, columnDef) {
  const existing = await graphFetch(
    token,
    "GET",
    `/sites/${siteId}/lists/${listId}/columns?$filter=name eq '${name}'`
  );
  if (existing?.value?.length) {
    const col = existing.value[0];
    // Column creation is skipped when it already exists (idempotent re-runs),
    // but that used to mean a column created before `indexed: true` was
    // added to its definition here just silently stayed un-indexed forever
    // — this is exactly what breaks `fields/EmployeeId eq '...'` filters with
    // "not indexed" errors even after re-running the script. Retrofit it.
    if (columnDef.indexed && !col.indexed) {
      await graphFetch(token, "PATCH", `/sites/${siteId}/lists/${listId}/columns/${col.id}`, {
        indexed: true
      });
      console.log(`    indexed existing column "${name}"`);
    }
    return;
  }
  await graphFetch(token, "POST", `/sites/${siteId}/lists/${listId}/columns`, {
    name,
    ...columnDef
  });
}

async function main() {
  const token = await getAppToken();

  console.log("Resolving site...");
  const site = await graphFetch(
    token,
    "GET",
    `/sites/${WORKHUB_SITE_HOSTNAME}:${WORKHUB_SITE_PATH}`
  );
  const siteId = site.id;
  console.log(`  site id: ${siteId}`);

  console.log("Creating document libraries...");
  await ensureList(token, siteId, "WorkHub Invoices", "documentLibrary");
  await ensureList(token, siteId, "WorkHub Documents", "documentLibrary");

  console.log("Creating WorkHub_Clients...");
  const clientsId = await ensureList(token, siteId, "WorkHub_Clients");
  await ensureColumn(token, siteId, clientsId, "ContactName", text());
  await ensureColumn(token, siteId, clientsId, "ContactEmail", text());
  await ensureColumn(token, siteId, clientsId, "DefaultRate", currency());
  await ensureColumn(token, siteId, clientsId, "BillingAddress", note());

  console.log("Creating WorkHub_Projects...");
  const projectsId = await ensureList(token, siteId, "WorkHub_Projects");
  await ensureColumn(token, siteId, projectsId, "ClientName", text());
  await ensureColumn(token, siteId, projectsId, "HourlyRate", currency());
  await ensureColumn(token, siteId, projectsId, "Budget", currency());
  await ensureColumn(token, siteId, projectsId, "BudgetHours", number());
  await ensureColumn(token, siteId, projectsId, "Status", choice(["Active", "On hold", "Completed"]));
  await ensureColumn(token, siteId, projectsId, "TeamMembers", personMulti());

  console.log("Creating WorkHub_Timesheets (weekly summary/cache only — no approval workflow)...");
  const timesheetsId = await ensureList(token, siteId, "WorkHub_Timesheets");
  await ensureColumn(token, siteId, timesheetsId, "EmployeeName", indexedText());
  await ensureColumn(token, siteId, timesheetsId, "WeekStart", dateTime());
  await ensureColumn(token, siteId, timesheetsId, "WeekEnd", dateTime());
  await ensureColumn(token, siteId, timesheetsId, "TotalHours", number());
  await ensureColumn(token, siteId, timesheetsId, "BillableHours", number());
  await ensureColumn(token, siteId, timesheetsId, "NonBillableHours", number());

  console.log("Creating WorkHub_TimeEntries...");
  const timeEntriesId = await ensureList(token, siteId, "WorkHub_TimeEntries");
  await ensureColumn(token, siteId, timeEntriesId, "EmployeeName", indexedText());
  await ensureColumn(token, siteId, timeEntriesId, "ProjectName", text());
  await ensureColumn(token, siteId, timeEntriesId, "ClientName", text());
  // TaskName is plain text, same NAME-based convention as ProjectName —
  // which specific task (from WorkHub_ProjectTasks) under this entry's
  // project the employee was working on, or "" when no specific task was
  // picked. Never a Lookup id back to WorkHub_ProjectTasks.
  await ensureColumn(token, siteId, timeEntriesId, "TaskName", text());
  await ensureColumn(token, siteId, timeEntriesId, "Date", { ...dateTime(), indexed: true });
  // Hours is a plain decimal number — 8, 7.5, 2.25 — entered directly by the
  // employee. There is no timer, no clock-in/clock-out, and nothing here is
  // ever stored or displayed in seconds.
  await ensureColumn(token, siteId, timeEntriesId, "Hours", number());
  // StartTime/EndTime are plain "HH:mm" TEXT, not a DateTime column and not
  // a revival of the timer — still no running clock, nothing captured
  // automatically. Just two optional fields typed by hand (e.g. "09:00" /
  // "17:30") so Calendar can draw a real time-of-day block instead of just
  // a flat daily total. Text avoids timezone conversion; Date already
  // carries which day it was.
  await ensureColumn(token, siteId, timeEntriesId, "StartTime", text());
  await ensureColumn(token, siteId, timeEntriesId, "EndTime", text());
  await ensureColumn(token, siteId, timeEntriesId, "Billable", boolean());
  await ensureColumn(token, siteId, timeEntriesId, "Description", note());
  await ensureColumn(token, siteId, timeEntriesId, "Tags", multiChoice(["scada", "qa", "call", "internal", "analysis"]));
  // No Status field — with no timer and no approval step, every row in this
  // list is simply a logged, finished entry; a status distinguishing "open"
  // from "completed" or "submitted" from "approved" has nothing left to
  // describe, so it was removed rather than kept unused.

  console.log("Creating WorkHub_ProjectTasks...");
  const projectTasksId = await ensureList(token, siteId, "WorkHub_ProjectTasks");
  // ProjectName is plain text here too, same convention as ProjectName on
  // WorkHub_TimeEntries — tasks reference the project by NAME, never by
  // SharePoint item id, so the app never has to resolve/keep a lookup id.
  await ensureColumn(token, siteId, projectTasksId, "ProjectName", indexedText());
  // The task's own text lives in the list's built-in Title column — no
  // separate column needed.

  console.log("Creating WorkHub_Invoices...");
  const invoicesId = await ensureList(token, siteId, "WorkHub_Invoices");
  await ensureColumn(token, siteId, invoicesId, "ClientName", text());
  await ensureColumn(token, siteId, invoicesId, "PeriodStart", dateTime());
  await ensureColumn(token, siteId, invoicesId, "PeriodEnd", dateTime());
  await ensureColumn(token, siteId, invoicesId, "Amount", currency());
  await ensureColumn(token, siteId, invoicesId, "Status", choice(["Draft", "Sent", "Paid", "Overdue"]));
  await ensureColumn(token, siteId, invoicesId, "DueDate", dateTime());
  await ensureColumn(token, siteId, invoicesId, "FileUrl", text());
  await ensureColumn(token, siteId, invoicesId, "LineItemsJson", note());

  console.log("Creating WorkHub_SyncLog...");
  const syncLogId = await ensureList(token, siteId, "WorkHub_SyncLog");
  await ensureColumn(token, siteId, syncLogId, "Timestamp", dateTime());
  await ensureColumn(token, siteId, syncLogId, "ActionType", choice(["Create", "Update", "Delete", "Sync", "Error"]));
  await ensureColumn(token, siteId, syncLogId, "Status", choice(["Success", "Failed"]));
  await ensureColumn(token, siteId, syncLogId, "Detail", note());

  console.log("Creating WorkHub_Settings...");
  const settingsId = await ensureList(token, siteId, "WorkHub_Settings");
  await ensureColumn(token, siteId, settingsId, "RetentionYears", number());

  console.log("Creating WorkHub_TimeOff...");
  const timeOffId = await ensureList(token, siteId, "WorkHub_TimeOff");
  await ensureColumn(token, siteId, timeOffId, "EmployeeId", indexedText());
  // EmployeeName is written alongside EmployeeId on every new request (see
  // requestTimeOff in src/graph/timeOff.js) — the "Team leave" timeline and
  // the Pending Approvals list group/display by name, and resolving a name
  // from an Entra id on every read would mean an extra Graph call per rower.
  // Storing it directly keeps that read path a single list query, same as
  // WorkHub_TimeEntries.
  await ensureColumn(token, siteId, timeOffId, "EmployeeName", indexedText());
  await ensureColumn(token, siteId, timeOffId, "Type", choice(["Vacation", "Sick leave", "Personal", "Holiday"]));
  await ensureColumn(token, siteId, timeOffId, "StartDate", dateTime());
  await ensureColumn(token, siteId, timeOffId, "EndDate", dateTime());
  await ensureColumn(token, siteId, timeOffId, "Days", number());
  await ensureColumn(token, siteId, timeOffId, "Notes", note());
  await ensureColumn(token, siteId, timeOffId, "Status", choice(["Pending", "Approved", "Rejected"]));

  console.log("Creating WorkHub_TimeOffBalances...");
  const timeOffBalancesId = await ensureList(token, siteId, "WorkHub_TimeOffBalances");
  await ensureColumn(token, siteId, timeOffBalancesId, "EmployeeId", indexedText());
  await ensureColumn(token, siteId, timeOffBalancesId, "AccruedDays", number());
  await ensureColumn(token, siteId, timeOffBalancesId, "UsedDays", number());
  await ensureColumn(token, siteId, timeOffBalancesId, "AvailableDays", number());

  console.log("Creating WorkHub_TeamMembers...");
  const teamMembersId = await ensureList(token, siteId, "WorkHub_TeamMembers");
  await ensureColumn(token, siteId, teamMembersId, "Email", text());
  await ensureColumn(token, siteId, teamMembersId, "Role", choice(["Owner", "Admin", "Member"]));
  await ensureColumn(token, siteId, teamMembersId, "Group", text());
  await ensureColumn(token, siteId, teamMembersId, "BillableRate", currency());

  console.log(
    "\nDone. Verify columns in the SharePoint UI, then grant the WorkHub-Members / " +
      "WorkHub-Admins groups access to this site."
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});