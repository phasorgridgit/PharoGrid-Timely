import { useEffect, useMemo, useState } from "react";
import { getTeamMembers, createTeamMember } from "../graph/team";

const ROLE_TAG = { Owner: "tag-approved", Admin: "tag-billable", Member: "tag-nonbillable" };
const ROLES = ["Owner", "Admin", "Member"];

export default function Team() {
  const [members, setMembers] = useState([]);
  const [search, setSearch] = useState("");
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ name: "", email: "", role: "Member", group: "", billableRate: "" });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    refresh();
  }, []);

  function refresh() {
    getTeamMembers().then(setMembers);
  }

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return members;
    return members.filter((m) => m.Name.toLowerCase().includes(q) || m.Email.toLowerCase().includes(q));
  }, [members, search]);

  async function handleCreate(e) {
    e.preventDefault();
    if (!form.name.trim() || !form.email.trim()) return;
    setBusy(true);
    setError("");
    try {
      await createTeamMember({
        name: form.name.trim(),
        email: form.email.trim(),
        role: form.role,
        group: form.group.trim(),
        billableRate: form.billableRate ? Number(form.billableRate) : 0
      });
      refresh();
      setShowModal(false);
      setForm({ name: "", email: "", role: "Member", group: "", billableRate: "" });
    } catch (err) {
      setError(err.message || "Couldn't add this team member.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h1>Team</h1>
          <div className="sub">{members.length} members across the org.</div>
        </div>
        <button className="btn btn-gold" onClick={() => setShowModal(true)}>+ Add new member</button>
      </div>

      <div className="filters-bar">
        <input
          className="search-input"
          placeholder="Search by name or email"
          value={search}
          onChange={(e) => setSearch(e.target.value)}
        />
        <div className="filter-spacer" />
      </div>

      <div className="card">
        <table>
          <thead>
            <tr><th>Name</th><th>Email</th><th>Billable rate</th><th>Role</th><th>Group</th></tr>
          </thead>
          <tbody>
            {filtered.map((m) => (
              <tr key={m.id}>
                <td className="name-cell">
                  <span className="avatar-sm">{m.Name.slice(0, 2).toUpperCase()}</span>
                  {m.Name}
                </td>
                <td className="mono">{m.Email}</td>
                <td className="mono">${m.BillableRate}/hr</td>
                <td><span className={`tag ${ROLE_TAG[m.Role] || "tag-nonbillable"}`}>{m.Role}</span></td>
                <td>{m.Group}</td>
              </tr>
            ))}
            {filtered.length === 0 && (
              <tr><td colSpan={5} className="hint" style={{ textAlign: "center", padding: 20 }}>No members match "{search}".</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showModal && (
        <div className="modal-overlay" onClick={() => setShowModal(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <h3>Add new member</h3>
              <button className="modal-close" onClick={() => setShowModal(false)}>×</button>
            </div>
            <form onSubmit={handleCreate}>
              <div className="modal-body">
                <div className="field">
                  <label>Name</label>
                  <input
                    autoFocus
                    value={form.name}
                    onChange={(e) => setForm({ ...form, name: e.target.value })}
                    placeholder="e.g. Priya Nair"
                    required
                  />
                </div>
                <div className="field">
                  <label>Email</label>
                  {/* This must exactly match the person's Microsoft 365 email/UPN —
                      getCurrentEmployee() resolves whoever signs in by matching
                      this field, so a typo here means that person's sign-in will
                      fail with a "no matching row" error until it's corrected. */}
                  <input
                    type="email"
                    value={form.email}
                    onChange={(e) => setForm({ ...form, email: e.target.value })}
                    placeholder="e.g. priya@phasorgrid.local"
                    required
                  />
                </div>
                <div className="field-row">
                  <div className="field">
                    <label>Role</label>
                    <select value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })}>
                      {ROLES.map((r) => <option key={r} value={r}>{r}</option>)}
                    </select>
                  </div>
                  <div className="field">
                    <label>Billable rate ($/hr)</label>
                    <input
                      type="number"
                      min="0"
                      step="1"
                      value={form.billableRate}
                      onChange={(e) => setForm({ ...form, billableRate: e.target.value })}
                      placeholder="0"
                    />
                  </div>
                </div>
                <div className="field">
                  <label>Group (optional)</label>
                  <input
                    value={form.group}
                    onChange={(e) => setForm({ ...form, group: e.target.value })}
                    placeholder="e.g. Protection Engineering"
                  />
                </div>
                {error && <div className="error" style={{ fontSize: 12.5 }}>{error}</div>}
              </div>
              <div className="modal-foot">
                <button type="button" className="btn btn-outline" onClick={() => setShowModal(false)}>Cancel</button>
                <button type="submit" className="btn btn-gold" disabled={busy}>{busy ? "Adding…" : "Add member"}</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
