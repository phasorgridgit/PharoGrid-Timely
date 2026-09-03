import { graphClient, getSiteId, isMockMode } from "./graphClient";
import { mockCreateEntry, mockPatchEntry, mockDeleteEntry, mockFilterEntries } from "./mockStore";
import { getCurrentEmployeeName } from "./currentUser";
import { formatDateOnly } from "../lib/dateUtils";

const LIST = "WorkHub_TimeEntries";

// Escapes a single quote for safe use inside an OData string literal
// (`fields/EmployeeName eq '...'`) — SharePoint/Graph OData escapes a quote
// by doubling it, same as SQL. Without this, a name/project containing an
// apostrophe (e.g. "O'Brien") breaks every filtered query silently.
function odataEscape(value) {
  return String(value).replace(/'/g, "''");
}

// Normalizes anything date-like (Date object, "YYYY-MM-DD", ISO string) to
// a plain "YYYY-MM-DD" string — Hours are logged against a calendar date,
// never a specific clock time.
//
// IMPORTANT: a Date object here is treated as already being UTC-anchored
// (see src/lib/dateUtils.js) by every caller in this codebase. It is NOT
// re-derived from local wall-clock fields here, because doing so would
// re-introduce the exact off-by-one-day bug dateUtils.js exists to prevent
// (an entry meant for Monday silently saving as Sunday for anyone east of
// UTC). If a raw, un-anchored `new Date()` is ever passed in directly, run
// it through `localDateToUTC()` first.
function toDateOnly(value) {
  if (value instanceof Date) return formatDateOnly(value);
  return String(value).slice(0, 10);
}

// "HH:mm" (24-hour) validation for the optional clock-time fields — these
// are plain text, typed in manually, never captured by a running timer.
const TIME_RE = /^([01]\d|2[0-3]):[0-5]\d$/;

function validateClockTime(value, label) {
  if (value === undefined || value === null || value === "") return "";
  if (!TIME_RE.test(value)) throw new Error(`${label} must be a time like 09:00 or 17:30 (got "${value}").`);
  return value;
}

/**
 * Hours between two "HH:mm" clock times on the same day (undefined if
 * either is missing/invalid) — used to auto-fill Hours when both StartTime
 * and EndTime are given, without forcing the employee to also type Hours
 * by hand. Rounded to the nearest quarter-hour, same granularity the Hours
 * field itself is entered in.
 */
export function hoursBetween(startTime, endTime) {
  if (!TIME_RE.test(startTime || "") || !TIME_RE.test(endTime || "")) return undefined;
  const [sh, sm] = startTime.split(":").map(Number);
  const [eh, em] = endTime.split(":").map(Number);
  const minutes = eh * 60 + em - (sh * 60 + sm);
  if (minutes <= 0) return undefined;
  return Math.round((minutes / 60) * 4) / 4;
}

function buildFields({ employeeName, projectName, clientName, taskName, date, hours, billable, description, tags, startTime, endTime }) {
  const dateOnly = toDateOnly(date);
  const numericHours = Number(hours);
  if (!Number.isFinite(numericHours) || numericHours <= 0) {
    throw new Error(`Hours must be a positive number (got ${hours}).`);
  }
  const fields = {
    Title: description ? String(description).slice(0, 255) : `${projectName}${taskName ? " — " + taskName : ""}`,
    EmployeeName: employeeName,
    ProjectName: projectName,
    ClientName: clientName || "",
    TaskName: taskName || "",
    Date: `${dateOnly}T00:00:00Z`,
    Hours: numericHours,
    Billable: !!billable,
    Description: description || "",
    StartTime: "",
    EndTime: ""
  };
  // Only include Tags when there's actually something to write — sending an
  // explicit empty array to a multi-choice column has been seen to 400 on
  // some tenants/Graph versions, and there's currently no Tags input in the
  // UI, so this key was always empty on every single save.
  if (tags && tags.length) fields.Tags = tags;
  return fields;
}

/**
 * Fetch the signed-in employee's own entries for a date range (inclusive,
 * "YYYY-MM-DD" strings). EmployeeName is always resolved server-side from
 * the signed-in account (see currentUser.js) — never accepted as a
 * parameter here, so there is no way for one employee's session to
 * accidentally (or otherwise) read or write as another employee.
 */
export async function getMyEntries({ startDate, endDate }) {
  const employeeName = await getCurrentEmployeeName();
  return getEntriesForEmployee(employeeName, { startDate, endDate });
}

/**
 * Fetch a specific employee's entries by NAME. Used by the admin "view any
 * employee's timesheet" picker in Timesheets.jsx — gated separately by
 * isCurrentUserAdmin() before this is ever called with a name other than
 * the caller's own.
 */
export async function getEntriesForEmployee(employeeName, { startDate, endDate }) {
  const start = `${toDateOnly(startDate)}T00:00:00Z`;
  const end = `${toDateOnly(endDate)}T23:59:59Z`;

  if (isMockMode) {
    return mockFilterEntries((f) => f.EmployeeName === employeeName && f.Date >= start && f.Date <= end);
  }

  const siteId = await getSiteId();
  const res = await graphClient
    .api(`/sites/${siteId}/lists/${LIST}/items`)
    .expand("fields")
    .filter(
      `fields/EmployeeName eq '${odataEscape(employeeName)}' and fields/Date ge '${start}' and fields/Date le '${end}'`
    )
    .get();

  return res.value.map((i) => ({ id: i.id, ...i.fields }));
}

/**
 * Create a manual time entry. This is the only way a WorkHub_TimeEntries
 * row gets created — there is no timer/tracking workflow. EmployeeName is
 * resolved here, not passed in, for the same reason as getMyEntries above.
 */
export async function addEntry({ projectName, clientName, taskName, date, hours, billable, description, tags, startTime, endTime }) {
  const employeeName = await getCurrentEmployeeName();
  const fields = buildFields({ employeeName, projectName, clientName, taskName, date, hours, billable, description, tags, startTime, endTime });

  if (isMockMode) return mockCreateEntry(fields);

  const siteId = await getSiteId();
  return graphClient.api(`/sites/${siteId}/lists/${LIST}/items`).post({ fields });
}

/**
 * Edit an existing entry. `patch` may include any of projectName, clientName,
 * taskName, date, hours, billable, description, tags, startTime, endTime —
 * mapped to the real SharePoint internal names below so the Graph payload
 * always matches the actual columns (ProjectName/ClientName/TaskName/Date/
 * Hours/Billable/Description/Tags/StartTime/EndTime), never
 * DurationSeconds/ProjectIdLookupId/ClientIdLookupId. StartTime/EndTime are
 * plain "HH:mm" text, not a timer — nothing here is captured automatically.
 */
export async function updateEntry(id, patch) {
  const next = {};
  if (patch.projectName !== undefined) next.ProjectName = patch.projectName;
  if (patch.clientName !== undefined) next.ClientName = patch.clientName || "";
  if (patch.taskName !== undefined) next.TaskName = patch.taskName || "";
  if (patch.date !== undefined) next.Date = `${toDateOnly(patch.date)}T00:00:00Z`;
  if (patch.hours !== undefined) {
    const numericHours = Number(patch.hours);
    if (!Number.isFinite(numericHours) || numericHours <= 0) {
      throw new Error(`Hours must be a positive number (got ${patch.hours}).`);
    }
    next.Hours = numericHours;
  }
  if (patch.billable !== undefined) next.Billable = !!patch.billable;
  if (patch.description !== undefined) {
    next.Description = patch.description || "";
    next.Title = patch.description ? String(patch.description).slice(0, 255) : next.Title;
  }
  if (patch.tags !== undefined) next.Tags = patch.tags || [];
  if (patch.startTime !== undefined) next.StartTime = validateClockTime(patch.startTime, "Start time");
  if (patch.endTime !== undefined) next.EndTime = validateClockTime(patch.endTime, "End time");

  if (isMockMode) return mockPatchEntry(id, next);

  const siteId = await getSiteId();
  await graphClient.api(`/sites/${siteId}/lists/${LIST}/items/${id}/fields`).patch(next);
  return { id, ...next };
}

/** Delete a single time entry row. */
export async function deleteEntry(id) {
  if (isMockMode) return mockDeleteEntry(id);
  const siteId = await getSiteId();
  await graphClient.api(`/sites/${siteId}/lists/${LIST}/items/${id}`).delete();
}

// Follows @odata.nextLink until every page has been fetched — SharePoint
// list items page at a default of 200 rows, so a plain single .get() would
// silently truncate "all-time, every employee" once the list grows past
// that. Only needed here, since every other query in this file is
// date-bounded (a week at a time) and naturally small.
async function fetchAllPages(request) {
  let res = await request.get();
  let items = res.value || [];
  while (res["@odata.nextLink"]) {
    res = await graphClient.api(res["@odata.nextLink"]).get();
    items = items.concat(res.value || []);
  }
  return items;
}

/**
 * Every time entry ever logged, by every employee — no date bound, no
 * employee filter. Used for org-wide reporting/export (e.g. the "Export to
 * Excel" workbook on Reports, one sheet per employee) rather than for
 * anything shown in the normal day-to-day Timesheets/Dashboard views, which
 * stay scoped to one employee and one date range.
 */
export async function getAllEntries() {
  if (isMockMode) return mockFilterEntries(() => true);

  const siteId = await getSiteId();
  const items = await fetchAllPages(graphClient.api(`/sites/${siteId}/lists/${LIST}/items`).expand("fields").top(200));
  return items.map((i) => ({ id: i.id, ...i.fields }));
}
