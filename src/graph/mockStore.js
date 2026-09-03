// Mock mode data store — lives only in memory for as long as the app window
// is open. Nothing here ever touches Microsoft Graph, SharePoint, or disk.
// Restarting the app resets everything back to the seed data below.

let nextId = 100;

export const mockMe = { id: "mock-user-1", displayName: "Yash", email: "yash@phasorgrid.local", role: "Admin" };

export const mockClients = [
  { id: 1, Title: "Meralco" },
  { id: 2, Title: "Davao Light & Power" }
];

export const mockProjects = [
  { id: 1, Title: "Grid Modernization", ClientName: "Meralco", HourlyRate: 145, Status: "Active" },
  { id: 2, Title: "Meter Rollout Ph.2", ClientName: "Davao Light & Power", HourlyRate: 110, Status: "Active" }
];

export const mockEntries = [];

export function mockCreateEntry(fields) {
  const item = { id: nextId++, fields: { ...fields } };
  mockEntries.push(item);
  return item;
}

export function mockPatchEntry(id, patch) {
  const item = mockEntries.find((e) => e.id === id);
  if (item) Object.assign(item.fields, patch);
  return item;
}

export function mockDeleteEntry(id) {
  const i = mockEntries.findIndex((e) => e.id === id);
  if (i >= 0) mockEntries.splice(i, 1);
}

export function mockFilterEntries(predicate) {
  return mockEntries.filter((e) => predicate(e.fields)).map((e) => ({ id: e.id, ...e.fields }));
}

// ---- team members ----
// Seed reflects a small engineering org — swap for a real WorkHub_TeamMembers
// SharePoint list (see scripts/provision-sharepoint-graph.mjs) once connected.
let nextTeamId = 200;
export const mockTeamMembers = [
  { id: 1, Name: "Yash", Email: "yash@phasorgrid.local", Role: "Admin", Group: "Protection Engineering", BillableRate: 120 },
  { id: 2, Name: "Dinesh K.", Email: "dinesh@phasorgrid.local", Role: "Admin", Group: "Protection Engineering", BillableRate: 110 },
  { id: 3, Name: "Mahesh N.", Email: "mahesh@phasorgrid.local", Role: "Admin", Group: "Compliance", BillableRate: 110 },
  { id: 4, Name: "Santhosh K.", Email: "santhosh@phasorgrid.local", Role: "Member", Group: "Compliance", BillableRate: 95 },
  { id: 5, Name: "Sneha R.", Email: "sneha@phasorgrid.local", Role: "Member", Group: "QA", BillableRate: 90 },
  { id: 6, Name: "Venkata P.", Email: "venkata@phasorgrid.local", Role: "Owner", Group: "Leadership", BillableRate: 150 }
];
nextTeamId = Math.max(...mockTeamMembers.map((m) => m.id)) + 1;

export function mockCreateTeamMember(fields) {
  const item = { id: nextTeamId++, ...fields };
  mockTeamMembers.push(item);
  return item;
}

// ---- time off ----
export const mockTimeOffPolicies = { Vacation: 20, Sick: 10, Personal: 5 };

export const mockTimeOffBalance = { accruedDays: 20, usedDays: 5, availableDays: 15 };

let nextTORId = 500;
export const mockTimeOffRequests = [
  { id: nextTORId++, employeeName: "Yash", type: "Sick leave", startDate: "2026-08-11", endDate: "2026-08-12", days: 2, status: "Approved" },
  { id: nextTORId++, employeeName: "Dinesh K.", type: "Vacation", startDate: "2026-08-18", endDate: "2026-08-22", days: 5, status: "Approved" },
  { id: nextTORId++, employeeName: "Sneha R.", type: "Personal", startDate: "2026-08-27", endDate: "2026-08-27", days: 1, status: "Pending" }
];

export function mockAddTimeOffRequest(req) {
  const item = { id: nextTORId++, status: "Pending", ...req };
  mockTimeOffRequests.push(item);
  return item;
}

export function mockUpdateTimeOffStatus(id, status) {
  const item = mockTimeOffRequests.find((r) => r.id === id);
  if (!item) throw new Error(`No time-off request with id ${id}.`);
  item.status = status;
  return item;
}

// ---- project tasks ----
let nextTaskId = 700;
export const mockProjectTasks = [
  { id: nextTaskId++, ProjectName: "Grid Modernization", Title: "Kick-off call with substation team" },
  { id: nextTaskId++, ProjectName: "Grid Modernization", Title: "Draft relay coordination study" },
  { id: nextTaskId++, ProjectName: "Meter Rollout Ph.2", Title: "Order AMI meters batch 3" }
];

export function mockCreateProjectTask(fields) {
  const item = { id: nextTaskId++, ...fields };
  mockProjectTasks.push(item);
  return item;
}

export function mockPatchProjectTask(id, patch) {
  const item = mockProjectTasks.find((t) => t.id === id);
  if (item) Object.assign(item, patch);
  return item;
}

export function mockDeleteProjectTask(id) {
  const i = mockProjectTasks.findIndex((t) => t.id === id);
  if (i >= 0) mockProjectTasks.splice(i, 1);
}
