import { graphClient, getSiteId, isMockMode } from "./graphClient";
import { mockEntries, mockProjects } from "./mockStore";
import { getProjects } from "./projects";

// Strips anything that isn't safe in a SharePoint/OneDrive folder or file
// name — blocks path traversal (../), slashes, and reserved characters
// before a user-derived value like a client name ever reaches a Graph path.
function sanitizePathSegment(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, "").replace(/\.\./g, "").trim().slice(0, 100);
}

function odataEscape(value) {
  return String(value).replace(/'/g, "''");
}

/**
 * Generates an invoice by pulling every billable time entry for a client
 * within a date range, from WorkHub_TimeEntries (the source of truth — no
 * separate "locked" or "completed" state; every logged entry is real).
 *
 * BUG FIXED: this used to compute `amount` from `entry.Rate`, a field that
 * is never actually written anywhere in this app (time entries have no
 * Rate field) — so `e.Rate || 0` always fell back to 0 and every invoice
 * came out to $0.00 regardless of hours logged. It now looks up each
 * entry's rate from its project's real HourlyRate (matched by ProjectName,
 * since entries reference projects by name, not id).
 */
export async function generateInvoice({ clientName, periodStart, periodEnd, pdfBuffer }) {
  if (isMockMode) {
    // Mock mode: computes a total from whatever's in memory right now — no
    // SharePoint record and no file upload, just enough to see the flow work.
    const rateByProject = Object.fromEntries(mockProjects.map((p) => [p.Title, p.HourlyRate || 0]));
    const lineItems = mockEntries
      .map((e) => e.fields)
      .filter((f) => f.Billable && f.ClientName === clientName && f.Date >= periodStart && f.Date <= periodEnd);
    const amount = lineItems.reduce((sum, e) => sum + (e.Hours || 0) * (rateByProject[e.ProjectName] || 0), 0);
    const invoiceNumber = `INV-MOCK-${Math.floor(Math.random() * 9000 + 1000)}`;
    return { invoiceNumber, amount, fileUrl: null };
  }

  const siteId = await getSiteId();

  const [entriesRes, projects] = await Promise.all([
    graphClient
      .api(`/sites/${siteId}/lists/WorkHub_TimeEntries/items`)
      .expand("fields")
      .filter(
        `fields/ClientName eq '${odataEscape(clientName)}' and fields/Billable eq 1 and fields/Date ge '${periodStart}' and fields/Date le '${periodEnd}'`
      )
      .get(),
    getProjects()
  ]);

  const rateByProject = Object.fromEntries(projects.map((p) => [p.Title, p.HourlyRate || 0]));

  const lineItems = entriesRes.value.map((e) => e.fields);
  const amount = lineItems.reduce((sum, e) => sum + (e.Hours || 0) * (rateByProject[e.ProjectName] || 0), 0);

  const invoiceNumber = `INV-${new Date().getFullYear()}-${Math.floor(Math.random() * 9000 + 1000)}`;

  const invoice = await graphClient.api(`/sites/${siteId}/lists/WorkHub_Invoices/items`).post({
    fields: {
      Title: invoiceNumber,
      ClientName: clientName,
      PeriodStart: periodStart,
      PeriodEnd: periodEnd,
      Amount: amount,
      Status: "Draft",
      LineItemsJson: JSON.stringify(lineItems)
    }
  });

  const safeClientName = sanitizePathSegment(clientName);
  const uploadPath = `/Invoices/${safeClientName}/${invoiceNumber}.pdf`;
  const uploaded = await graphClient
    .api(`/sites/${siteId}/drive/root:${uploadPath}:/content`)
    .put(pdfBuffer);

  await graphClient.api(`/sites/${siteId}/lists/WorkHub_Invoices/items/${invoice.id}/fields`).patch({
    FileUrl: uploaded.webUrl
  });

  return { invoiceNumber, amount, fileUrl: uploaded.webUrl };
}
