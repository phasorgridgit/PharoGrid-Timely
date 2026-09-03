import { graphClient, getSiteId, isMockMode } from "./graphClient";
import { mockTeamMembers, mockCreateTeamMember } from "./mockStore";

const LIST = "WorkHub_TeamMembers";
const ROLES = ["Owner", "Admin", "Member"];

/** Team roster — name, email, role, group, and billable rate. */
export async function getTeamMembers() {
  if (isMockMode) return mockTeamMembers;

  const siteId = await getSiteId();
  const res = await graphClient
    .api(`/sites/${siteId}/lists/${LIST}/items`)
    .expand("fields")
    .get();
  return res.value.map((i) => ({
    id: i.id,
    Name: i.fields.Title,
    Email: i.fields.Email,
    Role: i.fields.Role,
    Group: i.fields.Group,
    BillableRate: i.fields.BillableRate
  }));
}

/**
 * Adds a new roster row. This is the only way a WorkHub_TeamMembers row
 * gets created from inside the app — until someone signs in with an
 * account whose email matches a row added here, getCurrentEmployee() will
 * keep throwing its "no matching row" error for them (see currentUser.js),
 * so Email must be entered exactly as that person's Microsoft 365
 * email/UPN.
 *
 * Always resolves to the same { id, Name, Email, Role, Group,
 * BillableRate } shape getTeamMembers() returns, in both modes — mock and
 * real diverge only in HOW the row is stored: SharePoint has no "Name"
 * column, so the real path maps Name to the list's built-in Title column
 * (same convention as every other Name-bearing list in this app — see the
 * NOTE in provision-sharepoint.ps1) and maps it back on the way out.
 * Returning the raw Title-keyed record from the mock path used to crash
 * the Team table immediately after adding anyone (m.Name was undefined)
 * because mockTeamMembers' seed rows are Name-keyed, not Title-keyed.
 */
export async function createTeamMember({ name, email, role, group, billableRate }) {
  const trimmedName = String(name || "").trim();
  const trimmedEmail = String(email || "").trim();
  if (!trimmedName) throw new Error("Name is required.");
  if (!trimmedEmail) throw new Error("Email is required.");
  if (!ROLES.includes(role)) throw new Error(`Role must be one of ${ROLES.join(", ")}.`);

  const record = {
    Name: trimmedName.slice(0, 255),
    Email: trimmedEmail,
    Role: role,
    Group: group || "",
    BillableRate: Number(billableRate) || 0
  };

  if (isMockMode) return mockCreateTeamMember(record);

  const siteId = await getSiteId();
  const item = await graphClient.api(`/sites/${siteId}/lists/${LIST}/items`).post({
    fields: {
      Title: record.Name,
      Email: record.Email,
      Role: record.Role,
      Group: record.Group,
      BillableRate: record.BillableRate
    }
  });
  return {
    id: item.id,
    Name: item.fields.Title,
    Email: item.fields.Email,
    Role: item.fields.Role,
    Group: item.fields.Group,
    BillableRate: item.fields.BillableRate
  };
}
