import { useEffect, useState } from "react";
import { getClients, getProjects, createProject, createClient } from "../graph/projects";
import { getAllProjectTasks } from "../graph/projectTasks";
import { isCurrentUserAdmin } from "../graph/currentUser";
import ProjectTasks from "../components/ProjectTasks";

export default function ClientsProjects() {
  const [clients, setClients] = useState([]);
  const [projects, setProjects] = useState([]);
  const [tasks, setTasks] = useState([]);
  const [canManageClients, setCanManageClients] = useState(false);
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ title: "", clientName: "", hourlyRate: "" });
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const [showClientModal, setShowClientModal] = useState(false);
  const [clientForm, setClientForm] = useState({ title: "" });
  const [clientBusy, setClientBusy] = useState(false);
  const [clientError, setClientError] = useState("");

  // Project this project's Tasks popup is currently open for, or null.
  // Clicking either "+ Add Task" (zero tasks) or "N tasks" (some tasks) on a
  // project's row opens the same popup, scoped to that one project.
  const [taskModalProject, setTaskModalProject] = useState(null);

  useEffect(() => {
    refresh();
    isCurrentUserAdmin().then(setCanManageClients);
  }, []);

  function refresh() {
    getClients().then(setClients);
    getProjects().then(setProjects);
    refreshTasks();
  }

  function refreshTasks() {
    getAllProjectTasks().then(setTasks);
  }

  function taskCountFor(projectTitle) {
    return tasks.filter((t) => t.ProjectName === projectTitle).length;
  }

  async function handleCreate(e) {
    e.preventDefault();
    if (!form.title.trim()) return;
    setBusy(true);
    setError("");
    try {
      await createProject({
        title: form.title.trim(),
        clientName: form.clientName || null,
        hourlyRate: form.hourlyRate ? Number(form.hourlyRate) : 0
      });
      refresh();
      setShowModal(false);
      setForm({ title: "", clientName: "", hourlyRate: "" });
    } catch (err) {
      setError(err.message || "Couldn't create the project.");
    } finally {
      setBusy(false);
    }
  }

  async function handleCreateClient(e) {
    e.preventDefault();
    if (!clientForm.title.trim()) return;
    setClientBusy(true);
    setClientError("");
    try {
      await createClient({ title: clientForm.title.trim() });
      getClients().then(setClients);
      setShowClientModal(false);
      setClientForm({ title: "" });
    } catch (err) {
      setClientError(err.message || "Couldn't add the client.");
    } finally {
      setClientBusy(false);
    }
  }

  function closeTaskModal() {
    setTaskModalProject(null);
    refreshTasks();
  }

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h1>Clients &amp; Projects</h1>
          <div className="sub">{clients.length} clients · {projects.length} projects</div>
        </div>
        <div className="btn-row">
          {canManageClients && (
            <button className="btn btn-outline" onClick={() => setShowClientModal(true)}>+ Add client</button>
          )}
          <button className="btn btn-gold" onClick={() => setShowModal(true)}>+ New project</button>
        </div>
      </div>

      <div className="grid g2b">
        <div className="card">
          <h2 className="section-title">Clients</h2>
          <div className="plist">
            {clients.map((c) => (
              <div className="prow" key={c.id}>
                <div className="pmeta">
                  <div className="pname">{c.Title}</div>
                  <div className="pclient">{projects.filter((p) => p.ClientName === c.Title).length} projects</div>
                </div>
              </div>
            ))}
            {clients.length === 0 && <p className="hint">No clients yet.{canManageClients ? "" : " Ask an admin or owner to add one."}</p>}
          </div>
        </div>

        <div className="card">
          <h2 className="section-title">Projects</h2>
          <div className="plist">
            {projects.map((p) => {
              const count = taskCountFor(p.Title);
              return (
                <div className="prow" key={p.id} style={{ flexWrap: "wrap" }}>
                  <div className="pmeta">
                    <div className="pname">{p.Title}</div>
                    <div className="pclient">{p.ClientName ? `${p.ClientName} · ` : ""}${p.HourlyRate}/hr</div>
                  </div>
                  <span className="tag tag-active">{p.Status || "Active"}</span>

                  {count > 0 ? (
                    <button className="row-select-btn" onClick={() => setTaskModalProject(p.Title)}>
                      {count} task{count === 1 ? "" : "s"}
                    </button>
                  ) : (
                    <button className="row-select-btn" onClick={() => setTaskModalProject(p.Title)}>
                      <span style={{ fontSize: 14 }}>⊕</span> Add Task
                    </button>
                  )}
                </div>
              );
            })}
            {projects.length === 0 && <p className="hint">No projects yet.</p>}
          </div>
        </div>
      </div>

      {taskModalProject && <ProjectTasks projectName={taskModalProject} onClose={closeTaskModal} />}

      {showModal && (
        <div className="modal-overlay" onClick={() => setShowModal(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <h3>New project</h3>
              <button className="modal-close" onClick={() => setShowModal(false)}>×</button>
            </div>
            <form onSubmit={handleCreate}>
              <div className="modal-body">
                <div className="field">
                  <label>Project name</label>
                  <input
                    autoFocus
                    value={form.title}
                    onChange={(e) => setForm({ ...form, title: e.target.value })}
                    placeholder="e.g. Grid Modernization"
                    required
                  />
                </div>
                <div className="field-row">
                  <div className="field">
                    <label>Client</label>
                    <select value={form.clientName} onChange={(e) => setForm({ ...form, clientName: e.target.value })}>
                      <option value="">No client</option>
                      {clients.map((c) => <option key={c.id} value={c.Title}>{c.Title}</option>)}
                    </select>
                  </div>
                  <div className="field">
                    <label>Hourly rate ($)</label>
                    <input
                      type="number"
                      min="0"
                      step="1"
                      value={form.hourlyRate}
                      onChange={(e) => setForm({ ...form, hourlyRate: e.target.value })}
                      placeholder="0"
                    />
                  </div>
                </div>
                {error && <div className="error" style={{ fontSize: 12.5 }}>{error}</div>}
              </div>
              <div className="modal-foot">
                <button type="button" className="btn btn-outline" onClick={() => setShowModal(false)}>Cancel</button>
                <button type="submit" className="btn btn-gold" disabled={busy}>{busy ? "Creating…" : "Create project"}</button>
              </div>
            </form>
          </div>
        </div>
      )}
      {showClientModal && canManageClients && (
        <div className="modal-overlay" onClick={() => setShowClientModal(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <h3>Add client</h3>
              <button className="modal-close" onClick={() => setShowClientModal(false)}>×</button>
            </div>
            <form onSubmit={handleCreateClient}>
              <div className="modal-body">
                <div className="field">
                  <label>Client name</label>
                  <input
                    autoFocus
                    value={clientForm.title}
                    onChange={(e) => setClientForm({ title: e.target.value })}
                    placeholder="e.g. Meralco"
                    required
                  />
                </div>
                {clientError && <div className="error" style={{ fontSize: 12.5 }}>{clientError}</div>}
              </div>
              <div className="modal-foot">
                <button type="button" className="btn btn-outline" onClick={() => setShowClientModal(false)}>Cancel</button>
                <button type="submit" className="btn btn-gold" disabled={clientBusy}>{clientBusy ? "Adding…" : "Add client"}</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}
