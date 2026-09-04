import { useEffect, useMemo, useState } from "react";
import { getMyEntries, getEntriesForEmployee, addEntry, updateEntry, deleteEntry } from "../graph/timeEntries";
import { getProjects } from "../graph/projects";
import { getTeamMembers } from "../graph/team";
import { getCurrentEmployee, isCurrentUserAdmin } from "../graph/currentUser";
import ProjectPicker from "../components/ProjectPicker";
import { todayUTC, localDateToUTC, parseDateOnly, formatDateOnly, addDays, weekdayIndex, startOfWeek, isSameDay, formatLongDate, formatShortDate, formatShortDateWithYear } from "../lib/dateUtils";

const DAY_LABELS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

// A timesheet row is a (project, task) pair. TaskName is "" when no task
// was selected for that row — every row still keys/dedupes correctly since
// "" is a stable, consistent value rather than undefined/null sometimes.
function rowKey(projectName, taskName) {
  return `${projectName}\u0000${taskName || ""}`;
}

function rowLabel(projectName, taskName) {
  return taskName ? `${projectName} — ${taskName}` : projectName;
}

// `anchor` is a plain local `new Date()` (e.g. "today", or today shifted by
// whole weeks via the ‹ › buttons). It's converted to a UTC-anchored date
// exactly once here, then every day in the returned range is derived with
// UTC-only math — see src/lib/dateUtils.js for why that matters. Previously
// this built `start`/`end` with local setDate/setHours and then serialized
// them with toISOString(), which silently rolled the date back a day for
// anyone east of UTC (e.g. IST) — a Monday entry would save under Sunday's
// date and appear to "leak" into the previous week.
function getWeekBounds(anchor) {
  const start = startOfWeek(localDateToUTC(anchor));
  const end = addDays(start, 6);
  return { start, end };
}

function toDateStr(d) {
  return formatDateOnly(d);
}

function dayIndexOf(dateStr) {
  // entry.Date is stored as "YYYY-MM-DDT00:00:00Z" — parse as UTC so the
  // weekday doesn't shift depending on the viewer's local timezone.
  return weekdayIndex(parseDateOnly(dateStr));
}

// Builds the label shown in the date-nav: the actual date range for every
// week (so navigating always shows exactly which week you're on, not a
// stale "This week"), plus a relative "This week / Last week / Next week"
// tag when applicable.
function weekLabel(start, end) {
  const range = `${formatShortDate(start)} – ${formatShortDateWithYear(end)}`;
  const thisWeekStart = startOfWeek(todayUTC());
  let relative = null;
  if (isSameDay(start, thisWeekStart)) relative = "This week";
  else if (isSameDay(start, addDays(thisWeekStart, -7))) relative = "Last week";
  else if (isSameDay(start, addDays(thisWeekStart, 7))) relative = "Next week";
  return { range, relative };
}

export default function Timesheets() {
  const [anchor, setAnchor] = useState(new Date());
  const [me, setMe] = useState(null); // { name, email, role }
  const [admin, setAdmin] = useState(false);
  const [roster, setRoster] = useState([]);
  const [viewingEmployee, setViewingEmployee] = useState(null);
  const [entries, setEntries] = useState([]);
  const [projects, setProjects] = useState([]);
  const [addedRows, setAddedRows] = useState([]); // [{ projectName, taskName }] added manually with no entries yet
  const [pickerOpen, setPickerOpen] = useState(false);
  const [modal, setModal] = useState(null); // { projectName, taskName, dayIndex, entry? }
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");

  const { start, end } = getWeekBounds(anchor);
  const weekDays = DAY_LABELS.map((_, i) => addDays(start, i));
  const isOwnTimesheet = !!me && viewingEmployee === me.name;

  useEffect(() => {
    getProjects().then(setProjects);
    (async () => {
      try {
        const employee = await getCurrentEmployee();
        setMe(employee);
        setViewingEmployee(employee.name);
        const isAdmin = await isCurrentUserAdmin();
        setAdmin(isAdmin);
        if (isAdmin) getTeamMembers().then(setRoster);
      } catch (err) {
        setError(err.message || "Couldn't resolve your employee record.");
      }
    })();
  }, []);

  useEffect(() => {
    if (viewingEmployee) refresh();
  }, [start.getTime(), viewingEmployee]);

  async function refresh() {
    try {
      setError("");
      const fetched = isOwnTimesheet
        ? await getMyEntries({ startDate: toDateStr(start), endDate: toDateStr(end) })
        : await getEntriesForEmployee(viewingEmployee, { startDate: toDateStr(start), endDate: toDateStr(end) });
      setEntries(fetched);
    } catch (err) {
      setError(err.message || "Couldn't load this week's entries.");
    }
  }

  const totalHours = entries.reduce((s, e) => s + (e.Hours || 0), 0);

  const clientForProject = (name) => projects.find((p) => p.Title === name)?.ClientName || null;

  // Every distinct (project, task) pair that has entries this week, plus
  // any rows added but not yet saved — each is its own row in the grid.
  const rows = useMemo(() => {
    const byKey = new Map();
    entries.forEach((e) => {
      const key = rowKey(e.ProjectName, e.TaskName);
      if (!byKey.has(key)) byKey.set(key, { projectName: e.ProjectName, taskName: e.TaskName || "" });
    });
    addedRows.forEach((r) => {
      const key = rowKey(r.projectName, r.taskName);
      if (!byKey.has(key)) byKey.set(key, r);
    });
    return Array.from(byKey.values());
  }, [entries, addedRows]);

  function entriesFor(row, dayIndex) {
    return entries.filter(
      (e) => e.ProjectName === row.projectName && (e.TaskName || "") === (row.taskName || "") && dayIndexOf(e.Date) === dayIndex
    );
  }

  const dayTotals = DAY_LABELS.map((_, i) =>
    entries.filter((e) => dayIndexOf(e.Date) === i).reduce((s, e) => s + (e.Hours || 0), 0)
  );

  function openCell(row, dayIndex) {
    if (!isOwnTimesheet) return; // admin browsing someone else's timesheet is read-only
    const dayEntries = entriesFor(row, dayIndex);
    setModal({ projectName: row.projectName, taskName: row.taskName, dayIndex, entry: dayEntries[0] || null });
  }

  function handlePicked(project, task) {
    const row = { projectName: project.Title, taskName: task ? task.Title : "" };
    const key = rowKey(row.projectName, row.taskName);
    setAddedRows((r) => (r.some((x) => rowKey(x.projectName, x.taskName) === key) ? r : [...r, row]));
    setPickerOpen(false);
  }

  function handleCreated(project, task) {
    setProjects((p) => [...p, project]);
    setAddedRows((r) => [...r, { projectName: project.Title, taskName: task ? task.Title : "" }]);
    setPickerOpen(false);
  }

  // Deletes an entire row — every entry logged this week for that
  // (project, task) pair, not just one day's cell. A row with no saved
  // entries yet (added via the picker but nothing logged) is just removed
  // locally with no API calls.
  async function handleDeleteRow(row) {
    const rowEntries = DAY_LABELS.flatMap((_, i) => entriesFor(row, i));
    if (rowEntries.length === 0) {
      setAddedRows((r) => r.filter((x) => rowKey(x.projectName, x.taskName) !== rowKey(row.projectName, row.taskName)));
      return;
    }
    if (!window.confirm(`Delete all ${rowEntries.length} entr${rowEntries.length === 1 ? "y" : "ies"} logged this week for ${rowLabel(row.projectName, row.taskName)}? This can't be undone.`)) {
      return;
    }
    setBusy(true);
    setError("");
    try {
      await Promise.all(rowEntries.map((e) => deleteEntry(e.id)));
      await refresh();
    } catch (err) {
      setError(err.message || "Couldn't delete this row.");
    } finally {
      setBusy(false);
    }
  }

  const { range: weekRange, relative: weekRelative } = weekLabel(start, end);

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h1>Timesheets</h1>
          <div className="sub">Week of {formatLongDate(start)} · {totalHours.toFixed(2)}h logged{!isOwnTimesheet ? ` · viewing ${viewingEmployee}` : ""}</div>
        </div>
        <div className="date-nav">
          <button onClick={() => setAnchor((a) => { const d = new Date(a); d.setDate(d.getDate() - 7); return d; })}>‹</button>
          <div className="date-label" style={{ textAlign: "center" }}>
            <div>{weekRange}</div>
            {weekRelative && (
              <div style={{ fontSize: 11, fontWeight: 600, color: "var(--text-mute)", marginTop: 1 }}>{weekRelative}</div>
            )}
          </div>
          <button onClick={() => setAnchor((a) => { const d = new Date(a); d.setDate(d.getDate() + 7); return d; })}>›</button>
        </div>
      </div>

      {admin && roster.length > 0 && (
        <div className="filters-bar" style={{ marginBottom: 12, display: "flex", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
          <label className="hint" style={{ marginRight: 0 }}>Viewing timesheet for</label>
          <select className="filter-btn" value={viewingEmployee || ""} onChange={(e) => setViewingEmployee(e.target.value)}>
            {roster.map((m) => <option key={m.id} value={m.Name}>{m.Name}{m.Name === me?.name ? " (you)" : ""}</option>)}
          </select>
        </div>
      )}

      {error && <div className="mock-banner" style={{ background: "var(--red-soft)", borderColor: "var(--red)", color: "var(--red)" }}>{error}</div>}

      <div className="card">
        <table className="tsgrid">
          <thead>
            <tr>
              <th style={{ textAlign: "left" }}>Project</th>
              {DAY_LABELS.map((d, i) => (
                <th key={d}>
                  <div>{d}</div>
                  <div style={{ fontSize: 10.5, fontWeight: 500, color: "var(--text-mute)", marginTop: 2 }}>
                    {formatShortDate(weekDays[i])}
                  </div>
                </th>
              ))}
              <th>Total</th>
              {isOwnTimesheet && <th></th>}
            </tr>
          </thead>
          <tbody>
            {rows.map((row) => {
              const key = rowKey(row.projectName, row.taskName);
              const days = DAY_LABELS.map((_, i) => entriesFor(row, i).reduce((s, e) => s + (e.Hours || 0), 0));
              return (
                <tr key={key}>
                  <td className="proj">
                    <span className="pdot" style={{ background: "var(--teal)" }} />
                    <span>
                      {row.projectName}
                      {row.taskName && (
                        <div style={{ fontSize: 11, fontWeight: 500, color: "var(--text-mute)" }}>{row.taskName}</div>
                      )}
                    </span>
                  </td>
                  {days.map((h, i) => {
                    return (
                      <td key={i}>
                        <button
                          className="tscell"
                          style={{ border: "none", cursor: isOwnTimesheet ? "pointer" : "default", width: "100%", display: "flex", flexDirection: "column", gap: 2, height: "auto", padding: "6px 4px" }}
                          onClick={() => openCell(row, i)}
                        >
                          <span>{h ? h.toFixed(2) : "—"}</span>
                        </button>
                      </td>
                    );
                  })}
                  <td className="mono">{days.reduce((a, b) => a + b, 0).toFixed(2)}</td>
                  {isOwnTimesheet && (
                    <td style={{ textAlign: "center" }}>
                      <button
                        className="row-delete-btn"
                        title={`Delete this row's entries`}
                        onClick={() => handleDeleteRow(row)}
                        disabled={busy}
                      >
                        ×
                      </button>
                    </td>
                  )}
                </tr>
              );
            })}

            {isOwnTimesheet && (
              <tr>
                <td className="proj" style={{ position: "relative" }}>
                  <button className="row-select-btn" onClick={() => setPickerOpen((o) => !o)}>
                    <span style={{ fontSize: 15 }}>⊕</span> Select project
                  </button>
                  {pickerOpen && (
                    <ProjectPicker
                      projects={projects}
                      onPick={handlePicked}
                      onCreated={handleCreated}
                      onClose={() => setPickerOpen(false)}
                    />
                  )}
                </td>
                {DAY_LABELS.map((d) => (
                  <td key={d} title="Select a project first.">
                    <div className="tscell" style={{ opacity: 0.4 }}>—</div>
                  </td>
                ))}
                <td className="mono">0.00</td>
                <td></td>
              </tr>
            )}

            {rows.length === 0 && (
              <tr><td colSpan={isOwnTimesheet ? 10 : 9} className="hint" style={{ textAlign: "center", padding: 20 }}>
                {isOwnTimesheet ? "Nothing logged yet this week — select a project above to add hours." : `${viewingEmployee} hasn't logged anything this week.`}
              </td></tr>
            )}
          </tbody>
          <tfoot>
            <tr>
              <td>Daily total</td>
              {dayTotals.map((h, i) => <td key={i} className="mono">{h.toFixed(2)}</td>)}
              <td className="mono">{totalHours.toFixed(2)}</td>
              {isOwnTimesheet && <td></td>}
            </tr>
          </tfoot>
        </table>

      </div>

      {modal && (
        <EditTimeModal
          modal={modal}
          start={start}
          projectName={modal.projectName}
          taskName={modal.taskName}
          clientName={clientForProject(modal.projectName)}
          onClose={() => setModal(null)}
          onSaved={async () => { setModal(null); await refresh(); }}
        />
      )}
    </div>
  );
}

function EditTimeModal({ modal, start, projectName, taskName, clientName, onClose, onSaved }) {
  const { dayIndex, entry } = modal;
  const dayDate = addDays(start, dayIndex);

  const [description, setDescription] = useState(entry?.Title || "");
  const [hours, setHours] = useState(entry?.Hours != null ? String(entry.Hours) : "");
  const [billable, setBillable] = useState(entry ? !!entry.Billable : true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const hoursValue = Number(hours);
  const hoursValid = Number.isFinite(hoursValue) && hoursValue > 0 && hoursValue <= 24;

  async function handleSave() {
    setSaving(true);
    setError("");
    try {
      if (!hoursValid) throw new Error("Hours must be a positive number, up to 24.");
      if (entry) {
        await updateEntry(entry.id, { description, billable, hours: hoursValue, startTime: "", endTime: "" });
      } else {
        await addEntry({ description, projectName, clientName, taskName, date: dayDate, hours: hoursValue, billable });
      }
      await onSaved();
    } catch (err) {
      setError(err.message || "Couldn't save this entry. Please try again.");
    } finally {
      setSaving(false);
    }
  }

  async function handleDelete() {
    if (!entry) return;
    setSaving(true);
    setError("");
    try {
      await deleteEntry(entry.id);
      await onSaved();
    } catch (err) {
      setError(err.message || "Couldn't delete this entry.");
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <h3>Log time</h3>
          <button className="modal-close" onClick={onClose}>×</button>
        </div>
        <div className="modal-body">
          <div className="hint">{formatLongDate(dayDate)} · {projectName}{taskName ? ` — ${taskName}` : ""}</div>
          <div className="field" style={{ maxWidth: 150 }}>
            <label>Hours worked</label>
            <input
              type="number"
              className="mono"
              min="0.01"
              max="24"
              step="0.25"
              placeholder="e.g. 8 or 7.5"
              value={hours}
              onChange={(e) => setHours(e.target.value)}
              autoFocus
            />
          </div>
          <p className="hint" style={{ marginTop: -4, marginBottom: 12 }}>
            Only total hours worked are recorded. Start and end times are not required.
          </p>
          <div className="field">
            <label>Description</label>
            <textarea rows={2} placeholder="What have you worked on?" value={description} onChange={(e) => setDescription(e.target.value)} />
          </div>
          <label className="checkbox-field" style={{ color: "var(--text)" }}>
            <input type="checkbox" checked={billable} onChange={(e) => setBillable(e.target.checked)} /> Billable
          </label>
          {error && <div className="error" style={{ fontSize: 12.5 }}>{error}</div>}
        </div>
        <div className="modal-foot">
          {entry && <button className="btn btn-outline" onClick={handleDelete} disabled={saving} style={{ marginRight: "auto", color: "var(--red)" }}>Delete</button>}
          <button className="btn btn-outline" onClick={onClose} disabled={saving}>Cancel</button>
          <button className="btn btn-gold" onClick={handleSave} disabled={saving || !hoursValid}>{saving ? "Saving…" : "Save"}</button>
        </div>
      </div>
    </div>
  );
}
