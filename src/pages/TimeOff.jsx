import { useEffect, useMemo, useState } from "react";
import { getMyBalance, getTeamRequests, getMyRequests, requestTimeOff, updateRequestStatus } from "../graph/timeOff";
import { getTeamMembers } from "../graph/team";
import { getEntriesForEmployee } from "../graph/timeEntries";
import { isCurrentUserAdmin } from "../graph/currentUser";
import { startOfWeek as utcStartOfWeek, parseDateOnly, formatDateOnly, addDays, todayUTC, isSameDay, formatShortDateWithYear } from "../lib/dateUtils";

const TYPE_COLOR = { "Vacation": "#0AABC8", "Sick leave": "#1FAE6E", "Personal": "#F8BA17", "Holiday": "#8B6FE0" };
const STATUS_TAG = { Approved: "tag-approved", Rejected: "tag-rejected", Pending: "tag-pending" };
const MONTH_DAY_LABELS = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"];

// r.startDate/r.endDate are plain "YYYY-MM-DD" strings straight from the
// <input type="date"> in the request form — parse them as UTC-anchored
// (see src/lib/dateUtils.js) so they compare correctly against weekStart
// below, which is anchored the same way.
function normalizeEmployeeName(name) {
  return String(name || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function sameEmployee(employeeA, employeeB) {
  return normalizeEmployeeName(employeeA) === normalizeEmployeeName(employeeB);
}

function dayIndexInWeek(dateStr, weekStart) {
  const d = parseDateOnly(dateStr);
  return Math.round((d - weekStart) / 86400000);
}

// `d` is a plain local Date (just "today"). Anchored at UTC midnight once,
// then the Monday of that week is found with UTC-only math — same
// Monday-start convention as Timesheets/Calendar/Reports/Dashboard (this
// previously used a local, Sunday-start `setDate(getDate() - getDay())`,
// which put "Team leave — this week" a day out of step with every other
// week view in the app, and mixed local time-of-day math with UTC-parsed
// request dates).
function startOfWeek(d) {
  return utcStartOfWeek(localDateToUTC(d));
}

export default function TimeOff() {
  const [balance, setBalance] = useState({ accruedDays: 0, usedDays: 0, availableDays: 0 });
  const [requests, setRequests] = useState([]);
  const [myRequests, setMyRequests] = useState([]);
  const [roster, setRoster] = useState([]);
  const [admin, setAdmin] = useState(false);
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({ type: "Vacation", startDate: "", endDate: "", notes: "" });
  const [busy, setBusy] = useState(false);
  const [actioningId, setActioningId] = useState(null);
  const [error, setError] = useState("");
  const [formError, setFormError] = useState("");

  useEffect(() => {
    refresh();
    isCurrentUserAdmin().then(setAdmin);
    getTeamMembers().then(setRoster);
  }, []);

  function refresh() {
    getMyBalance().then(setBalance);
    getTeamRequests().then(setRequests);
    getMyRequests().then(setMyRequests);
  }


  const pendingRequests = requests
    .filter((r) => r.status === "Pending")
    .sort((a, b) => (a.startDate < b.startDate ? -1 : 1));

  async function submitRequest(e) {
    e.preventDefault();
    if (!form.startDate || !form.endDate) return;
    setFormError("");
    if (form.endDate < form.startDate) {
      setFormError("End date can't be before the start date.");
      return;
    }

    // A standalone one-day leave request cannot be filed on a federal holiday.
    // Multi-day leave is allowed to span a federal holiday; the holiday simply
    // remains part of that approved leave range.
    if (form.startDate === form.endDate) {
      const federalHolidayName = federalHolidayNameForDate(form.startDate);
      if (federalHolidayName) {
        setFormError(`${federalHolidayName} is a federal holiday. You cannot request leave for that date.`);
        return;
      }
    }

    // A person can't have two live requests covering the same day — check
    // for any date-range overlap against their own Pending/Approved
    // requests (a Rejected one doesn't block a new request for those days).
    const overlap = myRequests.find(
      (r) => r.status !== "Rejected" && form.startDate <= r.endDate && r.startDate <= form.endDate
    );
    if (overlap) {
      setFormError(
        `You already have a ${overlap.status.toLowerCase()} "${overlap.type}" request covering ` +
          `${formatShortDateWithYear(parseDateOnly(overlap.startDate))} → ${formatShortDateWithYear(parseDateOnly(overlap.endDate))}. ` +
          `Edit or cancel that one first if this replaces it.`
      );
      return;
    }
    setBusy(true);
    const days = Math.max(1, Math.round((new Date(form.endDate) - new Date(form.startDate)) / 86400000) + 1);
    const created = await requestTimeOff({ ...form, days });
    setRequests((prev) => [...prev, created]);
    setMyRequests((prev) => [...prev, created]);
    setBusy(false);
    setShowModal(false);
    setForm({ type: "Vacation", startDate: "", endDate: "", notes: "" });
  }

  // Approving/rejecting requests is handled here; the rest of the page
  // refreshes from the same request data source after the status changes.
  async function handleDecision(id, status) {
    setActioningId(id);
    setError("");
    try {
      await updateRequestStatus(id, status);
      refresh();
    } catch (err) {
      setError(err.message || `Couldn't set this request to ${status}.`);
    } finally {
      setActioningId(null);
    }
  }

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h1>Time Off</h1>
          <div className="sub">Manage leave balances, holidays, and who's out this week.</div>
        </div>
        <button className="btn btn-gold" onClick={() => { setFormError(""); setShowModal(true); }}>+ Request time off</button>
      </div>

      {error && <div className="mock-banner" style={{ background: "var(--red-soft)", borderColor: "var(--red)", color: "var(--red)" }}>{error}</div>}

      <div className="grid g3" style={{ marginBottom: 16 }}>
        <div className="card kpi">
          <div className="label">Accrued</div>
          <div className="value mono">{balance.accruedDays}d</div>
        </div>
        <div className="card kpi">
          <div className="label">Used</div>
          <div className="value mono">{balance.usedDays}d</div>
        </div>
        <div className="card kpi">
          <div className="label">Available</div>
          <div className="value mono" style={{ color: "var(--green)" }}>{balance.availableDays}d</div>
        </div>
      </div>

      {admin && (
        <div className="card" style={{ marginBottom: 16 }}>
          <h2 className="section-title">Pending approvals {pendingRequests.length > 0 ? `(${pendingRequests.length})` : ""}</h2>
          <table>
            <thead><tr><th>Employee</th><th>Type</th><th>Dates</th><th>Days</th><th>Notes</th><th></th></tr></thead>
            <tbody>
              {pendingRequests.map((r) => (
                <tr key={r.id}>
                  <td className="name-cell">
                    <span className="avatar-sm">{(r.employeeName || "?").slice(0, 2).toUpperCase()}</span>
                    {r.employeeName}
                  </td>
                  <td>{r.type}</td>
                  <td>{formatShortDateWithYear(parseDateOnly(r.startDate))} → {formatShortDateWithYear(parseDateOnly(r.endDate))}</td>
                  <td className="mono">{r.days}</td>
                  <td className="hint">{r.notes || "—"}</td>
                  <td style={{ textAlign: "right", whiteSpace: "nowrap" }}>
                    <button
                      className="btn btn-outline btn-sm"
                      onClick={() => handleDecision(r.id, "Approved")}
                      disabled={actioningId === r.id}
                      style={{ color: "var(--green)", marginRight: 6 }}
                    >
                      {actioningId === r.id ? "…" : "Approve"}
                    </button>
                    <button
                      className="btn btn-outline btn-sm"
                      onClick={() => handleDecision(r.id, "Rejected")}
                      disabled={actioningId === r.id}
                      style={{ color: "var(--red)" }}
                    >
                      {actioningId === r.id ? "…" : "Reject"}
                    </button>
                  </td>
                </tr>
              ))}
              {pendingRequests.length === 0 && (
                <tr><td colSpan={6} className="hint" style={{ textAlign: "center", padding: 20 }}>Nothing waiting on approval.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}


      <EmployeeMonthView roster={roster} teamRequests={requests} />

      <div className="card" style={{ marginTop: 16 }}>
        <h2 className="section-title">My requests</h2>
        <table>
          <thead><tr><th>Type</th><th>Dates</th><th>Days</th><th>Status</th></tr></thead>
          <tbody>
            {myRequests.map((r) => (
              <tr key={r.id}>
                <td>{r.type}</td>
                <td>{formatShortDateWithYear(parseDateOnly(r.startDate))} → {formatShortDateWithYear(parseDateOnly(r.endDate))}</td>
                <td className="mono">{r.days}</td>
                <td><span className={`tag ${STATUS_TAG[r.status] || "tag-pending"}`}>{r.status}</span></td>
              </tr>
            ))}
            {myRequests.length === 0 && (
              <tr><td colSpan={4} className="hint" style={{ textAlign: "center", padding: 20 }}>No time off requests yet.</td></tr>
            )}
          </tbody>
        </table>
      </div>

      {showModal && (
        <div className="modal-overlay" onClick={() => setShowModal(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <h3>Request time off</h3>
              <button className="modal-close" onClick={() => setShowModal(false)}>×</button>
            </div>
            <form onSubmit={submitRequest}>
              <div className="modal-body">
                <div className="field">
                  <label>Type</label>
                  <select value={form.type} onChange={(e) => setForm({ ...form, type: e.target.value })}>
                    {Object.keys(TYPE_COLOR).map((t) => <option key={t} value={t}>{t}</option>)}
                  </select>
                </div>
                <div className="field-row">
                  <div className="field">
                    <label>Start date</label>
                    <input type="date" value={form.startDate} onChange={(e) => setForm({ ...form, startDate: e.target.value })} required />
                  </div>
                  <div className="field">
                    <label>End date</label>
                    <input type="date" value={form.endDate} onChange={(e) => setForm({ ...form, endDate: e.target.value })} required />
                  </div>
                </div>
                <div className="field">
                  <label>Notes (optional)</label>
                  <textarea rows={3} value={form.notes} onChange={(e) => setForm({ ...form, notes: e.target.value })} />
                </div>
                {formError && <div className="error" style={{ fontSize: 12.5 }}>{formError}</div>}
              </div>
              <div className="modal-foot">
                <button type="button" className="btn btn-outline" onClick={() => setShowModal(false)}>Cancel</button>
                <button type="submit" className="btn btn-gold" disabled={busy}>{busy ? "Submitting…" : "Submit request"}</button>
              </div>
            </form>
          </div>
        </div>
      )}
    </div>
  );
}

/**
 * "Filter by employee" month view — pick anyone from the roster and see a
 * calendar of their month: which days they had APPROVED time off (colored
 * by leave type, same legend as the timeline above) versus which days they
 * logged hours worked. Pending/Rejected requests never mark a day as
 * "off" here — only Approved does, per the actual ask.
 */
function nthWeekdayOfMonth(year, month, weekday, nth) {
  const first = new Date(Date.UTC(year, month, 1));
  const offset = (weekday - first.getUTCDay() + 7) % 7;
  return new Date(Date.UTC(year, month, 1 + offset + (nth - 1) * 7));
}

function lastWeekdayOfMonth(year, month, weekday) {
  const last = new Date(Date.UTC(year, month + 1, 0));
  const offset = (last.getUTCDay() - weekday + 7) % 7;
  return new Date(Date.UTC(year, month + 1, last.getUTCDate() - offset));
}

function observeFederalHoliday(date) {
  const day = date.getUTCDay();
  if (day === 6) return addDays(date, -1);
  if (day === 0) return addDays(date, 1);
  return date;
}

function federalHolidaysForYear(year) {
  const list = [];
  const push = (name, actual) => {
    const observed = observeFederalHoliday(actual);
    list.push({
      date: formatDateOnly(observed),
      name: formatDateOnly(observed) === formatDateOnly(actual) ? name : `${name} (Observed)`
    });
  };
  push("New Year's Day", new Date(Date.UTC(year, 0, 1)));
  push("Martin Luther King Jr. Day", nthWeekdayOfMonth(year, 0, 1, 3));
  push("Washington's Birthday", nthWeekdayOfMonth(year, 1, 1, 3));
  push("Memorial Day", lastWeekdayOfMonth(year, 4, 1));
  push("Juneteenth National Independence Day", new Date(Date.UTC(year, 5, 19)));
  push("Independence Day", new Date(Date.UTC(year, 6, 4)));
  push("Labor Day", nthWeekdayOfMonth(year, 8, 1, 1));
  push("Columbus Day", nthWeekdayOfMonth(year, 9, 1, 2));
  push("Veterans Day", new Date(Date.UTC(year, 10, 11)));
  push("Thanksgiving Day", nthWeekdayOfMonth(year, 10, 4, 4));
  push("Christmas Day", new Date(Date.UTC(year, 11, 25)));
  return list;
}

function federalHolidayMapForDays(days) {
  const years = new Set();
  days.forEach((d) => {
    const y = d.getUTCFullYear();
    years.add(y - 1);
    years.add(y);
    years.add(y + 1);
  });
  const map = new Map();
  [...years].forEach((year) => {
    federalHolidaysForYear(year).forEach((holiday) => map.set(holiday.date, holiday.name));
  });
  return map;
}

function federalHolidayNameForDate(dateStr) {
  if (!dateStr) return null;
  const d = parseDateOnly(dateStr);
  const year = d.getUTCFullYear();
  return federalHolidaysForYear(year).find((holiday) => holiday.date === dateStr)?.name || null;
}

function EmployeeMonthView({ roster, teamRequests }) {
  const [employeeName, setEmployeeName] = useState("all");
  const [holidayPopup, setHolidayPopup] = useState(null);
  const [monthAnchor, setMonthAnchor] = useState(() => {
    const t = todayUTC();
    return new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth(), 1));
  });
  const [workedByDate, setWorkedByDate] = useState({});
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState("");


  // Monday of the week containing the 1st, through 41 days later — a full
  // 6-row month grid, same convention as a normal calendar (padded with the
  // trailing days of the previous/next month).
  const gridStart = utcStartOfWeek(monthAnchor);
  const gridDays = Array.from({ length: 42 }, (_, i) => addDays(gridStart, i));

  useEffect(() => {
    if (!employeeName) return;
    setLoading(true);
    setLoadError("");
    const gridEnd = gridDays[41];
    getEntriesForEmployee(employeeName, { startDate: formatDateOnly(gridStart), endDate: formatDateOnly(gridEnd) })
      .then((entries) => {
        const byDate = {};
        entries.forEach((e) => {
          const d = String(e.Date).slice(0, 10);
          byDate[d] = (byDate[d] || 0) + (e.Hours || 0);
        });
        setWorkedByDate(byDate);
      })
      .catch((err) => setLoadError(err.message || "Couldn't load this employee's logged time."))
      .finally(() => setLoading(false));
  }, [employeeName, monthAnchor.getTime()]);

  // Every calendar date covered by APPROVED requests. For "All employees"
  // keep every employee/leave type for the date; for one employee keep the
  // existing single-type behavior.
  const approvedOffByDate = useMemo(() => {
    const map = {};
    teamRequests
      .filter((r) => r.status === "Approved")
      .filter((r) => employeeName === "all" || sameEmployee(r.employeeName, employeeName))
      .forEach((r) => {
        let d = parseDateOnly(r.startDate);
        const last = parseDateOnly(r.endDate);
        while (d <= last) {
          const key = formatDateOnly(d);
          if (employeeName === "all") {
            (map[key] = map[key] || []).push({
              employeeName: r.employeeName || "Unknown",
              type: r.type
            });
          } else {
            map[key] = r.type;
          }
          d = addDays(d, 1);
        }
      });
    return map;
  }, [teamRequests, employeeName]);

  // Approved Holiday requests represent holidays declared for an employee.
  // Where creator metadata is present, it is used to identify Admin/Owner
  // declarations; approved Holiday requests remain the fallback holiday event
  // used by the existing Time Off data model.
  const adminHolidayByDate = useMemo(() => {
    const map = {};
    teamRequests
      .filter((r) => r.status === "Approved" && r.type === "Holiday")
      .filter((r) => employeeName === "all" || sameEmployee(r.employeeName, employeeName))
      .forEach((r) => {
        let d = parseDateOnly(r.startDate);
        const last = parseDateOnly(r.endDate);
        while (d <= last) {
          const key = formatDateOnly(d);
          (map[key] = map[key] || []).push({
            ...r,
            label: "Holiday by Admin/Owner"
          });
          d = addDays(d, 1);
        }
      });
    return map;
  }, [teamRequests, employeeName]);

  const today = todayUTC();
  const federalHolidayMap = useMemo(() => federalHolidayMapForDays(gridDays), [monthAnchor.getTime()]);
  const monthLabel = monthAnchor.toLocaleDateString(undefined, { month: "long", year: "numeric", timeZone: "UTC" });

  function shiftMonth(delta) {
    setMonthAnchor((m) => new Date(Date.UTC(m.getUTCFullYear(), m.getUTCMonth() + delta, 1)));
  }

  return (
    <div className="card" style={{ marginTop: 16 }}>
      <div className="page-head" style={{ marginBottom: 12 }}>
        <div>
          <h2 className="section-title" style={{ marginBottom: 4 }}>Employee calendar</h2>
          <div className="hint">Select an employee, or choose All employees to see who has approved time off on each date.</div>
        </div>
        <div className="date-nav">
          <button onClick={() => shiftMonth(-1)}>‹</button>
          <div className="date-label">{monthLabel}</div>
          <button onClick={() => shiftMonth(1)}>›</button>
        </div>
      </div>

      <div className="field" style={{ maxWidth: 280, marginBottom: 14 }}>
        <label>Employee</label>
        <select value={employeeName} onChange={(e) => setEmployeeName(e.target.value)} disabled={roster.length === 0}>
          <option value="all">All employees</option>
          {roster.length === 0 && <option value="">No team members yet</option>}
          {roster.map((m) => <option key={m.id} value={m.Name}>{m.Name}</option>)}
        </select>
      </div>

      {loadError && <div className="error" style={{ fontSize: 12.5, marginBottom: 10 }}>{loadError}</div>}

      <div
        className="month-grid"
        style={{
          width: "100%",
          maxWidth: "100%",
          minWidth: 0,
          boxSizing: "border-box",
          gridTemplateColumns: "repeat(7, minmax(0, 1fr))",
          overflow: "hidden"
        }}
      >
        {MONTH_DAY_LABELS.map((d) => (
          <div
            className="month-grid-dow"
            key={d}
            style={{ minWidth: 0, boxSizing: "border-box", overflow: "hidden" }}
          >{d}</div>
        ))}
        {gridDays.map((day, i) => {
          const dateStr = formatDateOnly(day);
          const inMonth = day.getUTCMonth() === monthAnchor.getUTCMonth();
          const offType = approvedOffByDate[dateStr];
          const hours = workedByDate[dateStr];
          return (
            <div
              key={i}
              className={isSameDay(day, today) ? "month-grid-cell today" : "month-grid-cell"}
              style={{ opacity: inMonth ? 1 : 0.35, minWidth: 0, maxWidth: "100%", boxSizing: "border-box", overflow: "hidden" }}
            >
              <div className="mgc-daynum">{day.getUTCDate()}</div>
              {adminHolidayByDate[dateStr]?.map((holiday) => (
                <button
                  key={`${holiday.id || holiday.employeeName}-${holiday.startDate}-${dateStr}`}
                  type="button"
                  className="mgc-tag"
                  title="View holiday details"
                  onClick={() => setHolidayPopup({ ...holiday, eventLabel: "Holiday by Admin/Owner" })}
                  style={{
                    width: "100%",
                    minWidth: 0,
                    maxWidth: "100%",
                    boxSizing: "border-box",
                    overflow: "hidden",
                    whiteSpace: "nowrap",
                    textOverflow: "ellipsis",
                    border: "none",
                    background: "#8B6FE0",
                    color: "#fff",
                    cursor: "pointer",
                    textAlign: "left"
                  }}
                >
                  {employeeName === "all" ? `${holiday.employeeName || "Employee"} · Holiday by Admin/Owner` : "Holiday by Admin/Owner"}
                </button>
              ))}

              {federalHolidayMap.get(dateStr) && (
                <button
                  type="button"
                  className="mgc-tag"
                  title="View event details"
                  onClick={() => setHolidayPopup({
                    employeeName: "All employees",
                    type: "Federal holiday",
                    startDate: dateStr,
                    endDate: dateStr,
                    notes: federalHolidayMap.get(dateStr),
                    eventLabel: `Holiday · ${federalHolidayMap.get(dateStr)}`
                  })}
                  style={{
                    width: "100%",
                    minWidth: 0,
                    maxWidth: "100%",
                    boxSizing: "border-box",
                    overflow: "hidden",
                    whiteSpace: "nowrap",
                    textOverflow: "ellipsis",
                    border: "none",
                    background: "#F8BA17",
                    color: "#6F5200",
                    cursor: "pointer",
                    textAlign: "left"
                  }}
                >
                  Holiday · {federalHolidayMap.get(dateStr)}
                </button>
              )}

              {employeeName === "all" ? (
                Array.isArray(offType) && offType.length > 0 ? (
                  <div style={{ display: "flex", flexDirection: "column", gap: 3 }}>
                    {offType.map((leave, idx) => (
                      <button
                        key={`${leave.employeeName}-${leave.type}-${idx}`}
                        type="button"
                        className="mgc-tag"
                        title="View event details"
                        onClick={() => {
                          const request = teamRequests.find((r) =>
                            r.status === "Approved" &&
                            sameEmployee(r.employeeName, leave.employeeName) &&
                            r.type === leave.type &&
                            r.startDate <= dateStr &&
                            r.endDate >= dateStr
                          );
                          setHolidayPopup(request || {
                            employeeName: leave.employeeName,
                            type: leave.type,
                            startDate: dateStr,
                            endDate: dateStr,
                            notes: ""
                          });
                        }}
                        style={{
                          width: "100%",
                          border: "none",
                          background: TYPE_COLOR[leave.type] || "#0AABC8",
                          color: "#fff",
                          cursor: "pointer",
                          textAlign: "left"
                        }}
                      >
                        {leave.employeeName} · {leave.type}
                      </button>
                    ))}
                  </div>
                ) : null
              ) : (
                offType ? (
                  <button
                    type="button"
                    className="mgc-tag"
                    title="View event details"
                    onClick={() => {
                      const request = teamRequests.find((r) =>
                        r.status === "Approved" &&
                        sameEmployee(r.employeeName, employeeName) &&
                        r.type === offType &&
                        r.startDate <= dateStr &&
                        r.endDate >= dateStr
                      );
                      setHolidayPopup(request || {
                        employeeName,
                        type: offType,
                        startDate: dateStr,
                        endDate: dateStr,
                        notes: ""
                      });
                    }}
                    style={{
                      width: "100%",
                      border: "none",
                      background: TYPE_COLOR[offType] || "#0AABC8",
                      color: "#fff",
                      cursor: "pointer",
                      textAlign: "left"
                    }}
                  >
                    {employeeName} · {offType}
                  </button>
                ) : hours > 0 ? (
                  <button
                    type="button"
                    className="mgc-worked"
                    title="View worked time"
                    onClick={() => setHolidayPopup({
                      employeeName,
                      type: "Worked time",
                      startDate: dateStr,
                      endDate: dateStr,
                      notes: `${hours.toFixed(2)}h worked`
                    })}
                    style={{
                      width: "100%",
                      border: "none",
                      background: "transparent",
                      cursor: "pointer",
                      textAlign: "left"
                    }}
                  >
                    {hours.toFixed(2)}h worked
                  </button>
                ) : null
              )}
            </div>
          );
        })}
      </div>

      {holidayPopup && (
        <div className="modal-overlay" onClick={() => setHolidayPopup(null)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <div className="modal-head">
              <h3>{holidayPopup.eventLabel || (holidayPopup.type === "Worked time" ? "Worked time details" : "Event details")}</h3>
              <button className="modal-close" onClick={() => setHolidayPopup(null)}>×</button>
            </div>
            <div className="modal-body">
              <div className="field">
                <label>Employee</label>
                <div>{holidayPopup.employeeName || "—"}</div>
              </div>
              <div className="field">
                <label>Type of holiday</label>
                <div>{holidayPopup.type || "Holiday"}</div>
              </div>
              <div className="field">
                <label>Date range</label>
                <div>{formatShortDateWithYear(parseDateOnly(holidayPopup.startDate))} → {formatShortDateWithYear(parseDateOnly(holidayPopup.endDate))}</div>
              </div>
              <div className="field">
                <label>Reason</label>
                <div>{holidayPopup.notes || "—"}</div>
              </div>
            </div>
            <div className="modal-foot">
              <button className="btn btn-outline" onClick={() => setHolidayPopup(null)}>Close</button>
            </div>
          </div>
        </div>
      )}

      {loading && <p className="hint" style={{ marginTop: 8 }}>Loading…</p>}
    </div>
  );
}
