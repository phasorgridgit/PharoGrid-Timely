import { useEffect, useMemo, useState } from "react";
import { getMyEntries, getEntriesForEmployee } from "../graph/timeEntries";
import { getCurrentEmployee, isCurrentUserAdmin } from "../graph/currentUser";
import { getTeamMembers } from "../graph/team";
import { getProjects } from "../graph/projects";
import { getAllProjectTasks } from "../graph/projectTasks";
import { todayUTC, startOfWeek, addDays, weekdayIndex, parseDateOnly, formatDateOnly, formatShortDate } from "../lib/dateUtils";

const WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];
const DEFAULT_FILTERS = { team: "all", project: "all", task: "all", status: "all", description: "" };

/** "3h 20m" — the KPI cards use clock-style duration, not decimal hours. */
function formatHM(hoursDecimal) {
  const totalMinutes = Math.round((hoursDecimal || 0) * 60);
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${h}h ${String(m).padStart(2, "0")}m`;
}

/**
 * Dashboard shows different data depending on role. A Member only ever sees
 * their own logged time — getMyEntries resolves EmployeeName server-side
 * from the signed-in account, so there's no Team filter and no way to
 * request anyone else's data. An Admin/Owner gets the "Admin Dashboard"
 * view instead: a Team filter defaulting to the whole org, with Project/
 * Task/Status/Description filters layered on top, and the option to drill
 * into one employee's week at a time.
 */
export default function Dashboard() {
  const [admin, setAdmin] = useState(false);
  const [me, setMe] = useState(null);
  const [roster, setRoster] = useState([]);
  const [projects, setProjects] = useState([]);
  const [allTasks, setAllTasks] = useState([]);

  const [anchor, setAnchor] = useState(todayUTC());
  const [todayEntries, setTodayEntries] = useState([]);
  const [weekEntries, setWeekEntries] = useState([]);

  const [draftFilters, setDraftFilters] = useState(DEFAULT_FILTERS);
  const [appliedFilters, setAppliedFilters] = useState(DEFAULT_FILTERS);

  useEffect(() => {
    (async () => {
      const employee = await getCurrentEmployee();
      setMe(employee);
      const isAdmin = await isCurrentUserAdmin();
      setAdmin(isAdmin);
      if (isAdmin) getTeamMembers().then(setRoster);
      else {
        setDraftFilters((f) => ({ ...f, team: "me" }));
        setAppliedFilters((f) => ({ ...f, team: "me" }));
      }
      getProjects().then(setProjects);
      getAllProjectTasks().then(setAllTasks);
    })();
  }, []);

  const weekStart = startOfWeek(anchor);
  const weekEnd = addDays(weekStart, 6);

  async function fetchForTeam(team, start, end) {
    const range = { startDate: formatDateOnly(start), endDate: formatDateOnly(end) };
    if (!admin || team === "me") return getMyEntries(range);
    if (team === "all") {
      const perEmployee = await Promise.all(roster.map((m) => getEntriesForEmployee(m.Name, range)));
      return perEmployee.flat();
    }
    return getEntriesForEmployee(team, range);
  }

  useEffect(() => {
    if (!me) return;
    (async () => {
      const today = todayUTC();
      const [todayEnt, weekEnt] = await Promise.all([
        fetchForTeam(appliedFilters.team, today, today),
        fetchForTeam(appliedFilters.team, weekStart, weekEnd)
      ]);
      setTodayEntries(todayEnt);
      setWeekEntries(weekEnt);
    })();
  }, [me, admin, appliedFilters.team, weekStart.getTime(), roster.length]);

  function matchesClientFilters(e) {
    if (appliedFilters.project !== "all" && e.ProjectName !== appliedFilters.project) return false;
    if (appliedFilters.task !== "all" && e.TaskName !== appliedFilters.task) return false;
    if (appliedFilters.status === "billable" && !e.Billable) return false;
    if (appliedFilters.status === "nonbillable" && e.Billable) return false;
    if (appliedFilters.description.trim()) {
      const q = appliedFilters.description.trim().toLowerCase();
      const hay = `${e.Title || ""} ${e.Description || ""}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  }

  const filteredToday = useMemo(() => todayEntries.filter(matchesClientFilters), [todayEntries, appliedFilters]);
  const filteredWeek = useMemo(() => weekEntries.filter(matchesClientFilters), [weekEntries, appliedFilters]);

  const todayHours = filteredToday.reduce((s, e) => s + (e.Hours || 0), 0);
  const weekHours = filteredWeek.reduce((s, e) => s + (e.Hours || 0), 0);
  const weekTargetHours = admin && appliedFilters.team === "all" ? (roster.length || 1) * 40 : 40;
  const weekPct = Math.min(100, Math.round((weekHours / weekTargetHours) * 100));

  const byDay = WEEKDAYS.map((_, i) =>
    filteredWeek.filter((e) => weekdayIndex(parseDateOnly(e.Date)) === i).reduce((s, e) => s + (e.Hours || 0), 0)
  );
  const maxDay = Math.max(...byDay, 1);

  const byProjectTask = {};
  filteredWeek.forEach((e) => {
    const project = e.ProjectName || "No project";
    const task = e.TaskName || "—";
    const key = `${project}\u0000${task}`;
    if (!byProjectTask[key]) byProjectTask[key] = { project, task, hours: 0 };
    byProjectTask[key].hours += e.Hours || 0;
  });

  const taskOptions = useMemo(() => {
    const scoped = draftFilters.project === "all" ? allTasks : allTasks.filter((t) => t.ProjectName === draftFilters.project);
    return [...new Set(scoped.map((t) => t.Title))];
  }, [allTasks, draftFilters.project]);

  const activeCount = Object.entries(appliedFilters).filter(([k, v]) => {
    if (k === "team") return admin && v !== "all";
    if (k === "description") return v.trim() !== "";
    return v !== "all";
  }).length;

  function handleApply() {
    setAppliedFilters(draftFilters);
  }

  let chartTitle = "Weekly activity";
  if (admin) {
    if (appliedFilters.team === "all") chartTitle = "Team weekly work";
    else if (appliedFilters.team === "me") chartTitle = `${me?.name || "My"}'s weekly work`;
    else chartTitle = `${appliedFilters.team}'s weekly work`;
  }

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h1>{admin ? "Admin Dashboard" : "Dashboard"}</h1>
          <div className="sub">Signed in as {me?.name || "…"}{me ? ` · ${me.role}` : ""}</div>
        </div>
        <div className="date-nav">
          <button onClick={() => setAnchor(addDays(weekStart, -7))}>‹</button>
          <div className="date-label">{formatShortDate(weekStart)} – {formatShortDate(weekEnd)}</div>
          <button onClick={() => setAnchor(addDays(weekStart, 7))}>›</button>
        </div>
      </div>

      <div className="grid g4" style={{ marginBottom: 16 }}>
        <div className="card kpi">
          <div className="label">Tracked Today</div>
          <div className="value mono">{formatHM(todayHours)}</div>
        </div>
        <div className="card kpi">
          <div className="label">Tracked This Week</div>
          <div className="value mono">{formatHM(weekHours)}</div>
          <div style={{ height: 6, background: "var(--surface)", borderRadius: 99, marginTop: 10, overflow: "hidden" }}>
            <div style={{ width: `${weekPct}%`, height: "100%", background: "linear-gradient(90deg,var(--teal),var(--teal-deep))" }} />
          </div>
        </div>
        <div className="card kpi">
          <div className="label">Entries This Week</div>
          <div className="value">{filteredWeek.length}</div>
        </div>
        <div className="card kpi">
          <div className="label">Weekly Target</div>
          <div className="value">{weekPct}%</div>
          <div className="delta">of {weekTargetHours}h</div>
        </div>
      </div>

      <div className="card" style={{ marginBottom: 16 }}>
        <div className="filters-panel-head">
          <h2 className="section-title" style={{ marginBottom: 0 }}>Filters</h2>
          {activeCount > 0 && <span className="active-badge">{activeCount} active</span>}
        </div>
        <div className="filters-panel-grid">
          {admin && (
            <div className="field">
              <label>Team</label>
              <select value={draftFilters.team} onChange={(e) => setDraftFilters({ ...draftFilters, team: e.target.value })}>
                <option value="all">All employees</option>
                <option value="me">Me</option>
                {roster.filter((m) => m.Name !== me?.name).map((m) => (
                  <option key={m.id} value={m.Name}>{m.Name}</option>
                ))}
              </select>
            </div>
          )}
          <div className="field">
            <label>Project</label>
            <select
              value={draftFilters.project}
              onChange={(e) => setDraftFilters({ ...draftFilters, project: e.target.value, task: "all" })}
            >
              <option value="all">All projects</option>
              {projects.map((p) => <option key={p.id} value={p.Title}>{p.Title}</option>)}
            </select>
          </div>
          <div className="field">
            <label>Task</label>
            <select value={draftFilters.task} onChange={(e) => setDraftFilters({ ...draftFilters, task: e.target.value })}>
              <option value="all">All tasks</option>
              {taskOptions.map((title) => <option key={title} value={title}>{title}</option>)}
            </select>
          </div>
          <div className="field">
            <label>Status</label>
            <select value={draftFilters.status} onChange={(e) => setDraftFilters({ ...draftFilters, status: e.target.value })}>
              <option value="all">All statuses</option>
              <option value="billable">Billable</option>
              <option value="nonbillable">Non-billable</option>
            </select>
          </div>
          <div className="field">
            <label>Description</label>
            <input
              type="text"
              placeholder="Search description…"
              value={draftFilters.description}
              onChange={(e) => setDraftFilters({ ...draftFilters, description: e.target.value })}
            />
          </div>
        </div>
        <div className="btn-row" style={{ marginTop: 12, justifyContent: "flex-end" }}>
          <button className="btn btn-gold" onClick={handleApply}>Apply filter</button>
        </div>
      </div>

      <div className="grid g2">
        <div className="card">
          <h2 className="section-title">{chartTitle}</h2>
          {weekHours === 0 ? (
            <p className="hint">No data to show. Try adjusting the filters.</p>
          ) : (
            <div className="bars">
              {WEEKDAYS.map((day, i) => {
                const heightPx = Math.max(4, Math.round((byDay[i] / maxDay) * 140));
                return (
                  <div className="bar-col" key={day}>
                    <div className="bar-stack" style={{ height: heightPx }}>
                      <div className="bar-b" style={{ height: "100%" }} />
                    </div>
                    <div className="lbl">{day}</div>
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div className="card">
          <h2 className="section-title">By project</h2>
          <table>
            <thead><tr><th>Project</th><th>Task</th><th>Duration</th></tr></thead>
            <tbody>
              {Object.values(byProjectTask)
                .sort((a, b) => b.hours - a.hours)
                .map(({ project, task, hours }) => (
                  <tr key={`${project}\u0000${task}`}>
                    <td>{project}</td>
                    <td>{task}</td>
                    <td className="mono">{hours.toFixed(2)}h</td>
                  </tr>
                ))}
              {Object.keys(byProjectTask).length === 0 && (
                <tr><td colSpan={3} className="hint" style={{ textAlign: "center", padding: 20 }}>No data to show.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
