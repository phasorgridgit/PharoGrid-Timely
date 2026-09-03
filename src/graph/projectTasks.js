import { graphClient, getSiteId, isMockMode } from "./graphClient";
import { mockProjectTasks, mockCreateProjectTask, mockPatchProjectTask, mockDeleteProjectTask } from "./mockStore";

const LIST = "WorkHub_ProjectTasks";

// Same OData-escaping rule as timeEntries.js — a project name containing an
// apostrophe (e.g. "O'Brien Substation") would otherwise break the filter.
function odataEscape(value) {
  return String(value).replace(/'/g, "''");
}

/**
 * Fetch every task for a single project, by NAME (ProjectName is plain
 * text here, same convention as ProjectName/ClientName on WorkHub_TimeEntries
 * — never a SharePoint Lookup id).
 */
export async function getProjectTasks(projectName) {
  if (!projectName) return [];

  if (isMockMode) {
    return mockProjectTasks.filter((t) => t.ProjectName === projectName).map((t) => ({ ...t }));
  }

  const siteId = await getSiteId();
  const res = await graphClient
    .api(`/sites/${siteId}/lists/${LIST}/items`)
    .expand("fields")
    .filter(`fields/ProjectName eq '${odataEscape(projectName)}'`)
    .get();
  return res.value.map((i) => ({ id: i.id, ...i.fields }));
}

/**
 * Fetch every task across every project in one call — used where a page
 * needs per-project task counts (e.g. Clients & Projects deciding whether
 * to show a "+ Add Task" prompt) without firing one request per project.
 */
export async function getAllProjectTasks() {
  if (isMockMode) return mockProjectTasks.map((t) => ({ ...t }));

  const siteId = await getSiteId();
  const res = await graphClient
    .api(`/sites/${siteId}/lists/${LIST}/items`)
    .expand("fields")
    .get();
  return res.value.map((i) => ({ id: i.id, ...i.fields }));
}

/** Creates a task under a project. */
export async function createProjectTask({ projectName, title }) {
  const fields = { ProjectName: projectName, Title: String(title).slice(0, 255) };

  if (isMockMode) return mockCreateProjectTask(fields);

  const siteId = await getSiteId();
  const item = await graphClient.api(`/sites/${siteId}/lists/${LIST}/items`).post({ fields });
  return { id: item.id, ...item.fields };
}

/**
 * Edits a task's text. The project association is never changed here —
 * moving a task to a different project isn't a supported workflow, so
 * ProjectName is intentionally not part of `patch`.
 */
export async function updateProjectTask(id, title) {
  const next = { Title: String(title).slice(0, 255) };

  if (isMockMode) return mockPatchProjectTask(id, next);

  const siteId = await getSiteId();
  await graphClient.api(`/sites/${siteId}/lists/${LIST}/items/${id}/fields`).patch(next);
  return { id, ...next };
}

/** Deletes a single task. */
export async function deleteProjectTask(id) {
  if (isMockMode) return mockDeleteProjectTask(id);
  const siteId = await getSiteId();
  await graphClient.api(`/sites/${siteId}/lists/${LIST}/items/${id}`).delete();
}
