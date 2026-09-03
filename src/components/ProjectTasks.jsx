import { useEffect, useState } from "react";
import { getProjectTasks, createProjectTask, updateProjectTask, deleteProjectTask } from "../graph/projectTasks";

/**
 * Project Tasks popup — opened from a single project's row in Clients &
 * Projects (either its "+ Add Task" link with zero tasks, or its "N tasks"
 * link once it has some). Scoped to exactly that one project: shows its
 * existing tasks, a box to add a new one, and an Edit option for any task
 * selected from the list. Tasks persist through the same Graph/SharePoint
 * (or mock-store) mechanism as every other list in the app — see
 * graph/projectTasks.js.
 */
export default function ProjectTasks({ projectName, onClose }) {
  const [tasks, setTasks] = useState([]);
  const [selectedTaskId, setSelectedTaskId] = useState(null);
  const [text, setText] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    refresh();
  }, [projectName]);

  async function refresh() {
    try {
      setError("");
      const fetched = await getProjectTasks(projectName);
      setTasks(fetched);
    } catch (err) {
      setError(err.message || "Couldn't load tasks for this project.");
    } finally {
      setLoading(false);
    }
  }

  const selectedTask = tasks.find((t) => t.id === selectedTaskId) || null;

  function handleSelectTask(task) {
    // Selecting the same task again deselects it (and clears the box) so
    // there's an easy way back to "adding a new task".
    if (selectedTaskId === task.id) {
      setSelectedTaskId(null);
      setText("");
      return;
    }
    setSelectedTaskId(task.id);
    setText(task.Title);
  }

  async function handleAdd() {
    const title = text.trim();
    if (!title) return;
    setBusy(true);
    setError("");
    try {
      const created = await createProjectTask({ projectName, title });
      setTasks((t) => [...t, created]);
      setText("");
      setSelectedTaskId(null);
    } catch (err) {
      setError(err.message || "Couldn't add the task.");
    } finally {
      setBusy(false);
    }
  }

  async function handleEdit() {
    const title = text.trim();
    if (!selectedTask || !title) return;
    setBusy(true);
    setError("");
    try {
      await updateProjectTask(selectedTask.id, title);
      setTasks((list) => list.map((t) => (t.id === selectedTask.id ? { ...t, Title: title } : t)));
    } catch (err) {
      setError(err.message || "Couldn't update the task.");
    } finally {
      setBusy(false);
    }
  }

  async function handleRemove() {
    if (!selectedTask) return;
    if (!window.confirm(`Remove "${selectedTask.Title}"? This can't be undone.`)) return;
    setBusy(true);
    setError("");
    try {
      await deleteProjectTask(selectedTask.id);
      setTasks((list) => list.filter((t) => t.id !== selectedTask.id));
      setSelectedTaskId(null);
      setText("");
    } catch (err) {
      setError(err.message || "Couldn't remove the task.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h3>Tasks — {projectName}</h3>
          <button className="modal-close" onClick={onClose}>×</button>
        </div>
        <div className="modal-body">
          <div className="field">
            <label>Task name</label>
            <input
              autoFocus
              value={text}
              onChange={(e) => setText(e.target.value)}
              placeholder="e.g. Site survey"
              onKeyDown={(e) => e.key === "Enter" && (selectedTask ? handleEdit() : handleAdd())}
            />
          </div>

          <div className="btn-row" style={{ marginTop: 10 }}>
            <button className="btn btn-outline btn-sm" onClick={handleEdit} disabled={busy || !selectedTask || !text.trim()}>
              Edit
            </button>
            <button className="btn btn-gold btn-sm" onClick={handleAdd} disabled={busy || !text.trim()}>
              Add Task
            </button>
            <button className="btn btn-outline btn-sm" onClick={handleRemove} disabled={busy || !selectedTask} style={{ color: selectedTask ? "var(--red)" : undefined }}>
              Remove Task
            </button>
          </div>

          {error && <div className="error" style={{ fontSize: 12.5, marginTop: 10 }}>{error}</div>}

          <div className="plist" style={{ marginTop: 14, maxHeight: 260, overflowY: "auto" }}>
            {loading && <p className="hint">Loading tasks…</p>}
            {!loading && tasks.map((task) => (
              <div
                className="prow"
                key={task.id}
                onClick={() => handleSelectTask(task)}
                style={{
                  cursor: "pointer",
                  borderRadius: 8,
                  padding: "10px 10px",
                  background: task.id === selectedTaskId ? "var(--teal-soft)" : undefined
                }}
              >
                <div className="pmeta">
                  <div className="pname" style={{ color: task.id === selectedTaskId ? "var(--teal-deep)" : undefined }}>
                    {task.Title}
                  </div>
                </div>
                {task.id === selectedTaskId && <span className="tag tag-active">Selected — editing</span>}
              </div>
            ))}
            {!loading && tasks.length === 0 && (
              <p className="hint">No tasks yet for {projectName} — add one above.</p>
            )}
          </div>
        </div>
        <div className="modal-foot">
          <button className="btn btn-outline" onClick={onClose}>Done</button>
        </div>
      </div>
    </div>
  );
}
