import { graphClient, getSiteId, isMockMode } from "./graphClient";
import {
  mockMe,
  mockTimeOffBalance,
  mockTimeOffRequests,
  mockAddTimeOffRequest,
  mockUpdateTimeOffStatus
} from "./mockStore";

const LIST = "WorkHub_TimeOff";

export const TIME_OFF_STATUSES = ["Pending", "Approved", "Rejected"];

/**
 * Canonical shape every caller works with — { id, employeeName, type,
 * startDate, endDate, days, notes, status }.
 *
 * THE BUG THIS FUNCTION FIXES: live SharePoint rows come back with
 * PascalCase field names (Status, Type, StartDate, ...) straight off
 * `fields`, but the UI has always read lowerCamel (`r.status`, `r.type`,
 * ...) — the casing mock data happens to use. In mock mode that accidentally
 * worked; against real SharePoint, `r.status` was always `undefined`, so a
 * request's status appeared frozen (blank / stuck rendering as "pending"
 * styling) no matter what the Status column in SharePoint actually said,
 * and no matter how many times it was edited there directly. Every read
 * path below now goes through this normalizer so both sources agree.
 */
function normalize(raw) {
  return {
    id: raw.id,
    employeeName: raw.EmployeeName ?? raw.employeeName,
    type: raw.Type ?? raw.type,
    startDate: raw.StartDate ?? raw.startDate,
    endDate: raw.EndDate ?? raw.endDate,
    days: raw.Days ?? raw.days,
    notes: raw.Notes ?? raw.notes ?? "",
    status: raw.Status ?? raw.status ?? "Pending"
  };
}

/** Accrued / used / available day balance for the signed-in user. */
export async function getMyBalance() {
  if (isMockMode) return mockTimeOffBalance;

  const siteId = await getSiteId();
  const me = await graphClient.api("/me").get();
  const res = await graphClient
    .api(`/sites/${siteId}/lists/WorkHub_TimeOffBalances/items`)
    .expand("fields")
    .filter(`fields/EmployeeId eq '${me.id}'`)
    .get();
  const fields = res.value[0]?.fields || { AccruedDays: 0, UsedDays: 0, AvailableDays: 0 };
  return { accruedDays: fields.AccruedDays, usedDays: fields.UsedDays, availableDays: fields.AvailableDays };
}

/** Every team member's time-off requests, for the timeline / team calendar view. */
export async function getTeamRequests() {
  if (isMockMode) return mockTimeOffRequests.map(normalize);

  const siteId = await getSiteId();
  const res = await graphClient.api(`/sites/${siteId}/lists/${LIST}/items`).expand("fields").get();
  return res.value.map((i) => normalize({ id: i.id, ...i.fields }));
}

/**
 * The signed-in employee's own requests — filtered server-side by their
 * real Entra object id, not a hardcoded display name.
 */
export async function getMyRequests() {
  if (isMockMode) return mockTimeOffRequests.filter((r) => r.employeeName === mockMe.displayName).map(normalize);

  const siteId = await getSiteId();
  const me = await graphClient.api("/me").get();
  const res = await graphClient
    .api(`/sites/${siteId}/lists/${LIST}/items`)
    .expand("fields")
    .filter(`fields/EmployeeId eq '${me.id}'`)
    .get();
  return res.value.map((i) => normalize({ id: i.id, ...i.fields }));
}

/** Submit a new leave/time-off request for approval. */
export async function requestTimeOff({ type, startDate, endDate, days, notes }) {
  if (isMockMode) {
    return normalize(mockAddTimeOffRequest({ employeeName: mockMe.displayName, type, startDate, endDate, days, notes }));
  }

  const siteId = await getSiteId();
  const me = await graphClient.api("/me").get();
  const item = await graphClient.api(`/sites/${siteId}/lists/${LIST}/items`).post({
    fields: {
      Title: `${type} — ${me.displayName}`,
      EmployeeId: me.id,
      EmployeeName: me.displayName,
      Type: type,
      StartDate: startDate,
      EndDate: endDate,
      Days: days,
      Notes: notes || "",
      Status: "Pending"
    }
  });
  return normalize({ id: item.id, ...item.fields });
}

/**
 * Approve, reject, or reset a request — the piece that was entirely
 * missing before: nothing anywhere in the app could ever move a request's
 * Status off "Pending", by design or by accident, live or in mock mode.
 * This is now the one function that changes it, so every place a status
 * is shown (Team leave timeline, My requests, Pending approvals) reflects
 * the same value once it refetches.
 */
export async function updateRequestStatus(id, status) {
  if (!TIME_OFF_STATUSES.includes(status)) {
    throw new Error(`Status must be one of ${TIME_OFF_STATUSES.join(", ")} (got "${status}").`);
  }

  if (isMockMode) return normalize(mockUpdateTimeOffStatus(id, status));

  const siteId = await getSiteId();
  await graphClient.api(`/sites/${siteId}/lists/${LIST}/items/${id}/fields`).patch({ Status: status });
  return { id, status };
}
