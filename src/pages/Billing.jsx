// Billing & Invoices is intentionally not wired up to real invoice
// generation yet — the underlying graph/invoices.js module is still there
// and untouched for when this feature is built out, but this page no
// longer presents that flow as usable. Keep this page reachable from the
// nav (it must stay listed), it just shows a clear "still being built"
// state instead of a half-working billing UI.
export default function Billing() {
  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h1>Billing &amp; Invoices</h1>
          <div className="sub">Generate invoices from logged, billable time — this month, by client.</div>
        </div>
      </div>

      <div className="card" style={{ textAlign: "center", padding: "56px 24px" }}>
        <span className="tag tag-pending" style={{ fontSize: 12.5, padding: "6px 14px", marginBottom: 14, display: "inline-block" }}>
          Under Progress
        </span>
        <h2 className="section-title" style={{ marginBottom: 6 }}>Billing &amp; Invoices is being built</h2>
        <p className="hint" style={{ maxWidth: 420, margin: "0 auto" }}>
          This section is still under development. Invoice generation isn't available yet — check back soon.
        </p>
      </div>
    </div>
  );
}
