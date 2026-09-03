import { graphClient, isMockMode } from "./graphClient";
import { getTeamMembers } from "./team";
import { mockMe } from "./mockStore";

let cachedEmployee = null; // per-renderer-session cache, keyed implicitly to whoever is signed in this run

/**
 * Resolves the currently signed-in Microsoft account to their configured
 * employee record in WorkHub_TeamMembers, by matching email addresses.
 *
 * This is the ONLY place "who am I" is decided. Every other module asks
 * this one for the current employee's Name rather than reading /me or a
 * global directly — that's what makes multi-employee, simultaneous use
 * safe: each signed-in Microsoft account resolves independently to its own
 * WorkHub_TeamMembers row, and nothing here is hardcoded or shared.
 *
 * Throws a clear, actionable error if the signed-in account has no
 * matching WorkHub_TeamMembers row — this is a real, common setup error
 * (someone signed in with an account nobody added to the roster yet), not
 * something to silently paper over with a fallback name.
 */
export async function getCurrentEmployee() {
  if (cachedEmployee) return cachedEmployee;

  if (isMockMode) {
    cachedEmployee = { name: mockMe.displayName, email: mockMe.email, role: mockMe.role };
    return cachedEmployee;
  }

  const me = await graphClient.api("/me").get();
  const myEmail = (me.mail || me.userPrincipalName || "").toLowerCase();
  if (!myEmail) {
    throw new Error("Microsoft sign-in didn't return an email/UPN for this account — can't resolve an employee record.");
  }

  const roster = await getTeamMembers();
  const match = roster.find((m) => (m.Email || "").toLowerCase() === myEmail);

  if (!match) {
    throw new Error(
      `Signed in as ${me.mail || me.userPrincipalName}, but no matching row exists in WorkHub_TeamMembers. ` +
        `Ask an admin to add "${me.displayName}" with Email = ${me.mail || me.userPrincipalName} to the Team Members list, then sign in again.`
    );
  }

  cachedEmployee = { name: match.Name, email: match.Email, role: match.Role };
  return cachedEmployee;
}

/** Convenience wrapper — most callers only need the Name string. */
export async function getCurrentEmployeeName() {
  return (await getCurrentEmployee()).name;
}

/** True for Owner/Admin roles — used to gate the admin "view any employee's timesheet" picker. */
export async function isCurrentUserAdmin() {
  const { role } = await getCurrentEmployee();
  return role === "Owner" || role === "Admin";
}

/** Clears the cached resolution — call this on sign-out so a different account resolves fresh next time. */
export function resetCurrentEmployeeCache() {
  cachedEmployee = null;
}
