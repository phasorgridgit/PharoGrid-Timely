import { graphClient, getSiteId, isMockMode } from "./graphClient";
import { mockProjects, mockClients } from "./mockStore";

export async function getProjects() {
  if (isMockMode) return mockProjects;

  const siteId = await getSiteId();
  const res = await graphClient
    .api(`/sites/${siteId}/lists/WorkHub_Projects/items`)
    .expand("fields")
    .get();
  return res.value.map((i) => ({ id: i.id, ...i.fields }));
}

export async function getClients() {
  if (isMockMode) return mockClients;

  const siteId = await getSiteId();
  const res = await graphClient
    .api(`/sites/${siteId}/lists/WorkHub_Clients/items`)
    .expand("fields")
    .get();
  return res.value.map((i) => ({ id: i.id, ...i.fields }));
}

/** Creates a client. */
export async function createClient({ title }) {
  if (isMockMode) {
    const client = { id: mockClients.reduce((max, c) => Math.max(max, c.id), 0) + 1, Title: title };
    mockClients.push(client);
    return client;
  }

  const siteId = await getSiteId();
  const item = await graphClient.api(`/sites/${siteId}/lists/WorkHub_Clients/items`).post({
    fields: { Title: title }
  });
  return { id: item.id, ...item.fields };
}

/** Creates a project. `clientName` is stored as plain text (WorkHub_Projects.ClientName) — not a Lookup id. */
export async function createProject({ title, clientName, hourlyRate, budget, budgetHours }) {
  if (isMockMode) {
    const project = {
      id: mockProjects.reduce((max, p) => Math.max(max, p.id), 0) + 1,
      Title: title,
      ClientName: clientName || "",
      HourlyRate: hourlyRate || 0,
      Budget: budget || 0,
      BudgetHours: budgetHours || 0,
      Status: "Active"
    };
    mockProjects.push(project);
    return project;
  }

  const siteId = await getSiteId();
  const item = await graphClient.api(`/sites/${siteId}/lists/WorkHub_Projects/items`).post({
    fields: {
      Title: title,
      ClientName: clientName || "",
      HourlyRate: hourlyRate,
      Budget: budget,
      BudgetHours: budgetHours,
      Status: "Active"
    }
  });
  return { id: item.id, ...item.fields };
}
