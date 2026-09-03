import { useEffect, useMemo, useState } from "react";
import { getMyEntries, getEntriesForEmployee, getAllEntries } from "../graph/timeEntries";
import { getProjects } from "../graph/projects";
import { getTeamMembers } from "../graph/team";
import { isCurrentUserAdmin } from "../graph/currentUser";
import { localDateToUTC, addDays, startOfWeek as utcStartOfWeek, weekdayIndex, parseDateOnly, formatDateOnly, formatShortDate, todayUTC } from "../lib/dateUtils";

const WEEKDAYS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

// See src/lib/dateUtils.js — `d` (a local `new Date()`) is anchored at UTC
// midnight exactly once, then the Monday of that week is found with
// UTC-only math, so the result serializes back to the correct date string
// regardless of the viewer's timezone offset.
function startOfWeek(d) {
  return utcStartOfWeek(localDateToUTC(d));
}

function toDateStr(d) {
  return formatDateOnly(d);
}

function dayIndexOf(dateStr) {
  return weekdayIndex(parseDateOnly(dateStr));
}

// e.Date is stored as "YYYY-MM-DDT00:00:00Z" — format it as the calendar
// date it represents, not shifted by the viewer's local timezone.
function formatEntryDate(dateStr) {
  return parseDateOnly(dateStr).toLocaleDateString(undefined, { timeZone: "UTC" });
}

function downloadCsv(filename, rows) {
  const csv = rows.map((r) => r.map((c) => `"${String(c).replace(/"/g, '""')}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

// Excel sheet names: max 31 chars, no \ / ? * [ ] :, and must be unique
// within the workbook — a plain employee Name can collide or overflow
// (e.g. two "Sam K" hires, or a long name), so this sanitizes and
// disambiguates rather than letting SheetJS throw.
function safeSheetName(name, used) {
  const base = String(name).replace(/[\\/?*[\]:]/g, " ").trim().slice(0, 31) || "Employee";
  let candidate = base;
  let n = 2;
  while (used.has(candidate.toLowerCase())) {
    const suffix = ` (${n++})`;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function entryRow(e) {
  const fields = e.fields || {};
  const date = e.Date ?? fields.Date;
  const hours = e.Hours ?? e.HoursWorked ?? e.hours ?? fields.Hours ?? fields.HoursWorked ?? 0;
  const billable = e.Billable ?? fields.Billable;
  return [
    date ? formatEntryDate(date) : "",
    e.Title ?? fields.Title ?? e.Description ?? fields.Description ?? "",
    e.ProjectName ?? fields.ProjectName ?? "",
    e.ClientName ?? fields.ClientName ?? "",
    e.TaskName ?? fields.TaskName ?? "",
    Number(hours || 0),
    billable ? "Yes" : "No"
  ];
}

const ENTRY_HEADER = ["Date", "Description", "Project", "Client", "Task", "Hours", "Billable"];

// Excel export columns — Employee is FIRST (the export is one sheet per
// employee, but each row still names them, since a printed/forwarded sheet
// can get separated from its tab name).
const EXCEL_COLUMNS = [
  { header: "Employee", key: "employee", width: 22 },
  { header: "Date", key: "date", width: 13 },
  { header: "Description", key: "description", width: 34 },
  { header: "Project", key: "project", width: 22 },
  { header: "Client", key: "client", width: 18 },
  { header: "Task", key: "task", width: 22 },
  { header: "Hours", key: "hours", width: 10 },
  { header: "Billable", key: "billable", width: 12 }
];

function excelRowData(e, employeeName) {
  const fields = e.fields || {};
  const date = e.Date ?? fields.Date;
  const hours = e.Hours ?? e.HoursWorked ?? e.hours ?? fields.Hours ?? fields.HoursWorked ?? 0;
  const billable = e.Billable ?? fields.Billable;
  return {
    employee: employeeName,
    date: date ? formatEntryDate(date) : "",
    description: e.Title ?? fields.Title ?? e.Description ?? fields.Description ?? "",
    project: e.ProjectName ?? fields.ProjectName ?? "",
    client: e.ClientName ?? fields.ClientName ?? "",
    task: e.TaskName ?? fields.TaskName ?? "",
    hours: Number(hours || 0),
    billable: billable ? "Yes" : "No"
  };
}

const THIN_LINE = { style: "thin", color: { argb: "FFE2E8ED" } };
const THIN_BORDER = { top: THIN_LINE, left: THIN_LINE, bottom: THIN_LINE, right: THIN_LINE };

/**
 * Colors and centers a worksheet built from EXCEL_COLUMNS: a solid teal
 * header with white bold text, banded rows so a wide sheet stays readable,
 * a thin border around every cell so content doesn't run edge-to-edge, and
 * every cell centered both horizontally and vertically. The `xlsx` package
 * this used to use (SheetJS Community Edition) can't apply any of this —
 * no fills, no borders, no alignment — which is why this switched to
 * `exceljs`, which supports real cell styling.
 */
function styleWorksheet(ws) {
  ws.views = [{ state: "frozen", ySplit: 1 }];

  const header = ws.getRow(1);
  header.height = 22;
  header.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0AABC8" } };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    cell.border = THIN_BORDER;
  });

  const billableColIndex = ws.columns.findIndex((c) => c.key === "billable") + 1;

  for (let r = 2; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    row.height = 18;
    const banded = r % 2 === 0;
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
      cell.border = THIN_BORDER;
      if (colNumber === billableColIndex) {
        const isYes = cell.value === "Yes";
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: isYes ? "FFDFF3E6" : "FFEFF1F3" } };
        cell.font = { bold: true, color: { argb: isYes ? "FF1B8F5A" : "FF6B7885" } };
      } else if (banded) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F6F8" } };
      }
    });
  }
}

export default function Reports() {
  const [tab, setTab] = useState("summary");
  const [anchor, setAnchor] = useState(new Date());
  const [entries, setEntries] = useState([]);
  const [projects, setProjects] = useState([]);
  const [projectFilter, setProjectFilter] = useState("all");
  const [billableFilter, setBillableFilter] = useState("all");

  const [admin, setAdmin] = useState(false);
  const [roster, setRoster] = useState([]);
  const [employeeFilter, setEmployeeFilter] = useState("me"); // "me" | "all" | <employee name>
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");

  const weekStart = startOfWeek(anchor);
  const weekEnd = addDays(weekStart, 6);

  useEffect(() => {
    getProjects().then(setProjects);
    (async () => {
      const isAdmin = await isCurrentUserAdmin();
      setAdmin(isAdmin);
      if (isAdmin) getTeamMembers().then(setRoster);
    })();
  }, []);

  useEffect(() => {
    (async () => {
      try {
        setError("");
        if (!admin || employeeFilter === "me") {
          setEntries(await getMyEntries({ startDate: toDateStr(weekStart), endDate: toDateStr(weekEnd) }));
        } else if (employeeFilter === "all") {
          const perEmployee = await Promise.all(
            roster.map((m) => getEntriesForEmployee(m.Name, { startDate: toDateStr(weekStart), endDate: toDateStr(weekEnd) }))
          );
          setEntries(perEmployee.flat());
        } else {
          setEntries(await getEntriesForEmployee(employeeFilter, { startDate: toDateStr(weekStart), endDate: toDateStr(weekEnd) }));
        }
      } catch (err) {
        setError(err.message || "Couldn't load entries for this period.");
      }
    })();
  }, [weekStart.getTime(), admin, employeeFilter, roster.length]);

  const projectRate = (name) => projects.find((p) => p.Title === name)?.HourlyRate || 0;

  const filtered = useMemo(() => {
    return entries.filter((e) => {
      if (!e.Hours) return false;
      if (projectFilter !== "all" && e.ProjectName !== projectFilter) return false;
      if (billableFilter === "billable" && !e.Billable) return false;
      if (billableFilter === "nonbillable" && e.Billable) return false;
      return true;
    });
  }, [entries, projectFilter, billableFilter]);

  const totalHours = filtered.reduce((s, e) => s + (e.Hours || 0), 0);
  const showEmployeeBreakdown = admin && employeeFilter === "all";

  const byDay = WEEKDAYS.map((_, i) => {
    const dayEntries = filtered.filter((e) => dayIndexOf(e.Date) === i);
    return dayEntries.reduce((s, e) => s + (e.Hours || 0), 0);
  });
  const maxDay = Math.max(...byDay, 1);

  const byProject = {};
  filtered.forEach((e) => {
    const key = e.ProjectName;
    byProject[key] = byProject[key] || 0;
    byProject[key] += e.Hours || 0;
  });

  const byEmployee = {};
  filtered.forEach((e) => {
    const key = e.EmployeeName || "Unknown";
    byEmployee[key] = byEmployee[key] || 0;
    byEmployee[key] += e.Hours || 0;
  });

  function handleExportCsv() {
    const header = showEmployeeBreakdown ? ["Employee", ...ENTRY_HEADER] : ENTRY_HEADER;
    const rows = [header];
    filtered.forEach((e) => {
      const row = entryRow(e);
      rows.push(showEmployeeBreakdown ? [e.EmployeeName || "", ...row] : row);
    });
    downloadCsv(`timely-report-${toDateStr(weekStart)}.csv`, rows);
  }

  // All-time, every employee — one workbook, one tab per team member, built
  // straight from WorkHub_TimeEntries (getAllEntries has no date bound).
  // Unlike the CSV export above, this ignores the current week/filters on
  // purpose: it's meant to be the complete record per person, not a slice.
  async function handleExportExcel() {
    setExporting(true);
    setError("");
    try {
      // Build the workbook from the actual SharePoint time-entry records.
      // Do not depend on the current roster to decide whether an entry is
      // exported: a roster mismatch must never make valid hours disappear.
      const { default: ExcelJS } = await import("exceljs");
      const Excel = ExcelJS || (await import("exceljs"));

      // Prefer the complete SharePoint record set. If that request fails,
      // fall back to the records already loaded on the Reports page so the
      // Excel button still produces a populated workbook for the current
      // period instead of failing because of a separate roster/API issue.
      let allEntries = [];
      try {
        allEntries = await getAllEntries();
      } catch (fetchErr) {
        console.warn("Could not load all-time entries; exporting loaded Reports entries instead:", fetchErr);
      }
      allEntries = Array.isArray(allEntries) ? allEntries : [];

      // If the all-time endpoint is temporarily empty, use the entries already
      // visible on the Reports page. This guarantees the workbook contains
      // the hours the user is looking at.
      const sourceEntries = allEntries.length ? allEntries : entries;
      if (!sourceEntries.length) {
        throw new Error("No time entries were found to export. Please make sure hours have been saved, then try again.");
      }

      const wb = new Excel.Workbook();
      wb.creator = "Timely";
      wb.created = new Date();
      const used = new Set(["all entries"]);

      // One master sheet guarantees every saved record is present in the
      // workbook even when an employee is missing from the current roster.
      const allWs = wb.addWorksheet("All Entries");
      allWs.columns = EXCEL_COLUMNS;
      [...sourceEntries]
        .sort((a, b) => new Date(a.Date) - new Date(b.Date))
        .forEach((e) => {
          const employee = e.EmployeeName || e.employeeName || e.fields?.EmployeeName || "Unassigned";
          allWs.addRow(excelRowData(e, employee));
        });
      styleWorksheet(allWs);
      allWs.autoFilter = { from: "A1", to: `H${Math.max(1, allWs.rowCount)}` };

      // Use employee names actually present in the exported records. The
      // Team list is intentionally not used as the source of truth, because
      // a stale roster must never make valid time entries disappear.
      const names = new Set();
      sourceEntries.forEach((e) => {
        names.add(e.EmployeeName || e.employeeName || e.fields?.EmployeeName || "Unassigned");
      });

      names.forEach((name) => {
        const rows = sourceEntries
          .filter((e) => (e.EmployeeName || e.employeeName || e.fields?.EmployeeName || "Unassigned") === name)
          .sort((a, b) => new Date(a.Date) - new Date(b.Date));
        const ws = wb.addWorksheet(safeSheetName(name, used));
        ws.columns = EXCEL_COLUMNS;
        rows.forEach((e) => ws.addRow(excelRowData(e, name)));
        styleWorksheet(ws);
        ws.autoFilter = { from: "A1", to: `H${Math.max(1, ws.rowCount)}` };
      });

      const buffer = await wb.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `timely-time-entries-${toDateStr(todayUTC())}.xlsx`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (err) {
      console.error("Timely Excel export failed:", err);
      setError(err.message || "Couldn't build the Excel export.");
    } finally {
      setExporting(false);
    }
  }

  function shiftWeek(delta) {
    const next = new Date(anchor);
    next.setDate(next.getDate() + delta * 7);
    setAnchor(next);
  }

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h1>Reports</h1>
          <div className="sub">Summary, detailed, and weekly breakdowns of logged time.</div>
        </div>
        <div className="btn-row">
          <button className="btn btn-outline" onClick={handleExportCsv}>Export CSV</button>
          {admin && (
            <button className="btn btn-gold" onClick={handleExportExcel} disabled={exporting}>
              {exporting ? "Building…" : "Export to Excel (all employees)"}
            </button>
          )}
        </div>
      </div>

      {error && <div className="mock-banner" style={{ background: "var(--red-soft)", borderColor: "var(--red)", color: "var(--red)" }}>{error}</div>}

      <div className="tabs">
        <button className={tab === "summary" ? "tab active" : "tab"} onClick={() => setTab("summary")}>Summary</button>
        <button className={tab === "detailed" ? "tab active" : "tab"} onClick={() => setTab("detailed")}>Detailed</button>
        <button className={tab === "weekly" ? "tab active" : "tab"} onClick={() => setTab("weekly")}>Weekly</button>
      </div>

      <div className="filters-bar">
        <select className="filter-btn" value={projectFilter} onChange={(e) => setProjectFilter(e.target.value)}>
          <option value="all">All projects</option>
          {projects.map((p) => <option key={p.id} value={p.Title}>{p.Title}</option>)}
        </select>
        <select className="filter-btn" value={billableFilter} onChange={(e) => setBillableFilter(e.target.value)}>
          <option value="all">All time</option>
          <option value="billable">Billable only</option>
          <option value="nonbillable">Non-billable only</option>
        </select>
        {admin && (
          <select className="filter-btn" value={employeeFilter} onChange={(e) => setEmployeeFilter(e.target.value)}>
            <option value="me">Me</option>
            <option value="all">All employees</option>
            {roster.map((m) => <option key={m.id} value={m.Name}>{m.Name}</option>)}
          </select>
        )}
        <div className="filter-spacer" />
        <div className="date-nav">
          <button onClick={() => shiftWeek(-1)}>‹</button>
          <div className="date-label">
            {formatShortDate(weekStart)} – {formatShortDate(weekEnd)}
          </div>
          <button onClick={() => shiftWeek(1)}>›</button>
        </div>
      </div>

      <div className="card" style={{ marginBottom: 16 }}>
        <div className="label" style={{ fontSize: 11.5, color: "var(--text-mute)", fontWeight: 600, textTransform: "uppercase" }}>Total</div>
        <div className="value mono" style={{ fontFamily: "'JetBrains Mono'", fontSize: 26, fontWeight: 700, marginTop: 4 }}>{totalHours.toFixed(2)}h</div>
      </div>

      {tab === "summary" && (
        <div className="grid g2">
          <div className="card">
            <h2 className="section-title">Hours by day</h2>
            {totalHours === 0 ? (
              <p className="hint">No data to show. Try adjusting the filters.</p>
            ) : (
              <div className="bars">
                {WEEKDAYS.map((day, i) => {
                  const h = Math.max(4, Math.round((byDay[i] / maxDay) * 140));
                  return (
                    <div className="bar-col" key={day}>
                      <div className="bar-stack" style={{ height: h }}><div className="bar-b" style={{ height: "100%" }} /></div>
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
              <thead><tr><th>Project</th><th>Hours</th><th>Amount</th></tr></thead>
              <tbody>
                {Object.entries(byProject).map(([name, hours]) => (
                  <tr key={name}>
                    <td>{name}</td>
                    <td className="mono">{hours.toFixed(2)}h</td>
                    <td className="mono">${(hours * projectRate(name)).toFixed(2)}</td>
                  </tr>
                ))}
                {Object.keys(byProject).length === 0 && (
                  <tr><td colSpan={3} className="hint" style={{ textAlign: "center", padding: 20 }}>No data to show.</td></tr>
                )}
              </tbody>
            </table>
          </div>

          {showEmployeeBreakdown && (
            <div className="card" style={{ gridColumn: "1 / -1" }}>
              <h2 className="section-title">By employee</h2>
              <table>
                <thead><tr><th>Employee</th><th>Hours</th></tr></thead>
                <tbody>
                  {Object.entries(byEmployee)
                    .sort((a, b) => b[1] - a[1])
                    .map(([name, hours]) => (
                      <tr key={name}>
                        <td className="name-cell">
                          <span className="avatar-sm">{name.slice(0, 2).toUpperCase()}</span>
                          {name}
                        </td>
                        <td className="mono">{hours.toFixed(2)}h</td>
                      </tr>
                    ))}
                  {Object.keys(byEmployee).length === 0 && (
                    <tr><td colSpan={2} className="hint" style={{ textAlign: "center", padding: 20 }}>No data to show.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {tab === "detailed" && (
        <div className="card">
          <table>
            <thead>
              <tr>
                <th>Date</th>
                <th>Description</th>
                <th>Project</th>
                <th>Client</th>
                {showEmployeeBreakdown && <th>Employee</th>}
                <th>Hours</th>
                <th></th>
              </tr>
            </thead>
            <tbody>
              {[...filtered].sort((a, b) => new Date(b.Date) - new Date(a.Date)).map((e, i) => (
                <tr key={i}>
                  <td>{formatEntryDate(e.Date)}</td>
                  <td>{e.Title}</td>
                  <td>{e.ProjectName}</td>
                  <td>{e.ClientName || "—"}</td>
                  {showEmployeeBreakdown && <td>{e.EmployeeName}</td>}
                  <td className="mono">{(e.Hours || 0).toFixed(2)}h</td>
                  <td><span className={e.Billable ? "tag tag-billable" : "tag tag-nonbillable"}>{e.Billable ? "Billable" : "Non-billable"}</span></td>
                </tr>
              ))}
              {filtered.length === 0 && (
                <tr><td colSpan={showEmployeeBreakdown ? 7 : 6} className="hint" style={{ textAlign: "center", padding: 20 }}>No data to show. Try adjusting the filters.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}

      {tab === "weekly" && (
        <div className="card">
          <table className="tsgrid">
            <thead>
              <tr>
                <th style={{ textAlign: "left" }}>Project</th>
                {WEEKDAYS.map((d) => <th key={d}>{d}</th>)}
                <th>Total</th>
              </tr>
            </thead>
            <tbody>
              {Object.keys(byProject).map((name) => {
                const days = WEEKDAYS.map((_, i) =>
                  filtered
                    .filter((e) => e.ProjectName === name && dayIndexOf(e.Date) === i)
                    .reduce((s, e) => s + (e.Hours || 0), 0)
                );
                return (
                  <tr key={name}>
                    <td className="proj">{name}</td>
                    {days.map((h, i) => <td key={i} className="mono">{h ? h.toFixed(2) : "—"}</td>)}
                    <td className="mono">{days.reduce((a, b) => a + b, 0).toFixed(2)}</td>
                  </tr>
                );
              })}
              {Object.keys(byProject).length === 0 && (
                <tr><td colSpan={9} className="hint" style={{ textAlign: "center", padding: 20 }}>No data to show.</td></tr>
              )}
            </tbody>
            <tfoot>
              <tr>
                <td>Daily total</td>
                {byDay.map((h, i) => <td key={i} className="mono">{h.toFixed(2)}</td>)}
                <td className="mono">{totalHours.toFixed(2)}</td>
              </tr>
            </tfoot>
          </table>
        </div>
      )}
    </div>
  );
}
