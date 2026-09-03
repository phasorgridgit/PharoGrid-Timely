import { useEffect, useMemo, useRef, useState } from "react";
import { createProject } from "../graph/projects";
import { getProjectTasks } from "../graph/projectTasks";

/**
 * Searchable "Select project" popover, matching the Clockify-style flow:
 * search existing projects, pick one, or create a brand new one inline
 * without leaving the picker. After a project is picked, a second step
 * shows that project's tasks so the employee can also say which task
 * they're working on — the row this produces is keyed by (project, task)
 * together, not just the project.
 *
 * Props:
 *  - projects: current project list
 *  - onPick(project, task): called when an existing project (+ optional task, or null) is chosen
 *  - onCreated(project, task): called after a new project is created (already added to the list) (+ optional task, or null)
 *  - onClose(): called when the popover should close
 */
export default function ProjectPicker({ projects, onPick, onCreated, onClose }) {
  const [query, setQuery] = useState("");
  const [creating, setCreating] = useState(false);
  const [newTitle, setNewTitle] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState(null);
  const boxRef = useRef(null);

  // Step 2: once a project is chosen (existing or just created), show its
  // tasks so the employee can pick which one they're logging time against.
  const [taskStepProject, setTaskStepProject] = useState(null); // { project, wasCreated }
  const [tasks, setTasks] = useState([]);
  const [tasksLoading, setTasksLoading] = useState(false);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return projects;
    return projects.filter((p) => p.Title.toLowerCase().includes(q));
  }, [projects, query]);

  useEffect(() => {
    if (!taskStepProject) return;
    setTasksLoading(true);
    getProjectTasks(taskStepProject.project.Title)
      .then(setTasks)
      .finally(() => setTasksLoading(false));
  }, [taskStepProject]);

  function goToTaskStep(project, wasCreated) {
    setTaskStepProject({ project, wasCreated });
  }

  function confirmTask(task) {
    const { project, wasCreated } = taskStepProject;
    if (wasCreated) onCreated(project, task || null);
    else onPick(project, task || null);
  }

  async function handleCreate() {
    const title = (creating ? newTitle : query).trim();
    if (!title) return;
    setBusy(true);
    setError(null);
    try {
      const project = await createProject({ title, clientName: null, hourlyRate: 0 });
      goToTaskStep({ id: project.id, Title: title, HourlyRate: 0, Status: "Active" }, true);
    } catch (err) {
      setError(err.message || "Couldn't create the project.");
    } finally {
      setBusy(false);
    }
  }

  if (taskStepProject) {
    const { project } = taskStepProject;
    return (
      <div className="popover" ref={boxRef} onClick={(e) => e.stopPropagation()}>
        <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
          <button
            className="row-select-btn"
            style={{ fontWeight: 600, fontSize: 12 }}
            onClick={() => setTaskStepProject(null)}
          >
            ‹ Back
          </button>
        </div>
        <div className="hint" style={{ marginBottom: 8, fontWeight: 600, color: "var(--text)" }}>{project.Title}</div>
        <div className="popover-list">
          <div className="popover-item" onClick={() => confirmTask(null)}>
            <span className="pdot" style={{ background: "var(--text-mute)" }} />
            No specific task
          </div>
          {tasksLoading && <div className="hint" style={{ padding: "8px 4px" }}>Loading tasks…</div>}
          {!tasksLoading && tasks.map((t) => (
            <div className="popover-item" key={t.id} onClick={() => confirmTask(t)}>
              <span className="pdot" style={{ background: "var(--teal)" }} />
              {t.Title}
            </div>
          ))}
          {!tasksLoading && tasks.length === 0 && (
            <div className="hint" style={{ padding: "8px 4px" }}>
              No tasks yet for this project — add one from Clients &amp; Projects.
            </div>
          )}
        </div>
      </div>
    );
  }

  return (
    <div className="popover" ref={boxRef} onClick={(e) => e.stopPropagation()}>
      <input
        className="popover-search"
        autoFocus
        placeholder="Search project"
        value={query}
        onChange={(e) => setQuery(e.target.value)}
        onKeyDown={(e) => e.key === "Escape" && onClose()}
      />
      <div className="popover-list">
        {filtered.map((p) => (
          <div className="popover-item" key={p.id} onClick={() => goToTaskStep(p, false)}>
            <span className="pdot" style={{ background: "var(--teal)" }} />
            {p.Title}
          </div>
        ))}
        {filtered.length === 0 && !query && <div className="hint" style={{ padding: "8px 4px" }}>No projects yet.</div>}
        {filtered.length === 0 && query && <div className="hint" style={{ padding: "8px 4px" }}>No matches for "{query}".</div>}
      </div>

      {error && <div className="error" style={{ fontSize: 12, padding: "4px 4px 0" }}>{error}</div>}

      {!creating && (
        <div className="popover-create" onClick={() => { setCreating(true); setNewTitle(query); }}>
          + Create new project{query ? ` "${query}"` : ""}
        </div>
      )}
      {creating && (
        <div className="popover-create" style={{ cursor: "default" }}>
          <input
            autoFocus
            placeholder="New project name"
            value={newTitle}
            onChange={(e) => setNewTitle(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && handleCreate()}
          />
          <button className="btn btn-gold btn-sm" onClick={handleCreate} disabled={busy || !newTitle.trim()}>
            {busy ? "Creating…" : "Create"}
          </button>
        </div>
      )}
    </div>
  );
}
