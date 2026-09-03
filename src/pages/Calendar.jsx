import { useEffect, useMemo, useState } from "react";
import { getMyEntries } from "../graph/timeEntries";
import { todayUTC, localDateToUTC, addDays, startOfWeek as utcStartOfWeek, formatDateOnly, formatShortDateWithYear } from "../lib/dateUtils";

const PALETTE = ["#0AABC8", "#F8BA17", "#1FAE6E", "#E1503F", "#8B6FE0", "#EE9F2E"];

function colorFor(key) {
  let hash = 0;
  for (const ch of String(key)) hash = (hash * 31 + ch.charCodeAt(0)) >>> 0;
  return PALETTE[hash % PALETTE.length];
}

function startOfWeek(d) {
  return utcStartOfWeek(localDateToUTC(d));
}

function toDateStr(d) {
  return formatDateOnly(d);
}

export default function Calendar() {
  const [anchor, setAnchor] = useState(new Date());
  const [entries, setEntries] = useState([]);
  const weekStart = startOfWeek(anchor);
  const days = Array.from({ length: 7 }, (_, i) => addDays(weekStart, i));

  useEffect(() => {
    const end = addDays(weekStart, 6);
    getMyEntries({ startDate: toDateStr(weekStart), endDate: toDateStr(end) }).then(setEntries);
  }, [weekStart.getTime()]);

  const byDay = useMemo(() => days.map((day) => {
    const dayStr = toDateStr(day);
    return entries.filter((e) => String(e.Date).slice(0, 10) === dayStr);
  }), [entries, weekStart.getTime()]);

  const weekTotal = entries.reduce((sum, e) => sum + Number(e.Hours || 0), 0);
  const todayStr = toDateStr(todayUTC());

  function shift(deltaWeeks) {
    const next = new Date(anchor);
    next.setDate(next.getDate() + deltaWeeks * 7);
    setAnchor(next);
  }

  const rangeLabel = `${formatShortDateWithYear(weekStart)} – ${formatShortDateWithYear(days[6])}`;

  return (
    <div className="page">
      <div className="page-head">
        <div>
          <h1>Calendar</h1>
          <div className="sub">Your weekly work summary — hours worked by day and project.</div>
        </div>
        <div className="date-nav">
          <button onClick={() => shift(-1)}>‹</button>
          <div className="date-label">{rangeLabel}</div>
          <button onClick={() => shift(1)}>›</button>
        </div>
      </div>

      <div className="card" style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ display: "grid", gridTemplateColumns: "150px repeat(7, minmax(125px, 1fr))", minWidth: 1050 }}>
          <div className="cal2-corner" style={{ padding: 14, fontWeight: 700 }}>Project</div>
          {days.map((day) => {
            const dayStr = toDateStr(day);
            const isToday = dayStr === todayStr;
            const total = byDay[days.indexOf(day)].reduce((s, e) => s + Number(e.Hours || 0), 0);
            return (
              <div key={dayStr} className={isToday ? "cal2-daylabel today" : "cal2-daylabel"} style={{ minHeight: 78, padding: 12 }}>
                <div className="dname">{day.toLocaleDateString(undefined, { weekday: "short", timeZone: "UTC" })}</div>
                <div className="ddate">{day.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric", timeZone: "UTC" })}</div>
                <div className="mono dtotal">{total ? `${total.toFixed(2)}h` : "—"}</div>
              </div>
            );
          })}

          {Array.from(new Set(entries.map((e) => `${e.ProjectName}\u0000${e.TaskName || ""}`))).map((key) => {
            const [projectName, taskName] = key.split("\u0000");
            return (
              <div key={key} style={{ display: "contents" }}>
                <div className="proj" style={{ padding: 14, borderTop: "1px solid var(--line)" }}>
                  <span className="pdot" style={{ background: colorFor(projectName) }} />
                  <span>{projectName}{taskName ? <div style={{ fontSize: 11, color: "var(--text-mute)" }}>{taskName}</div> : null}</span>
                </div>
                {days.map((day) => {
                  const dayStr = toDateStr(day);
                  const total = entries
                    .filter((e) => String(e.Date).slice(0, 10) === dayStr && e.ProjectName === projectName && (e.TaskName || "") === taskName)
                    .reduce((s, e) => s + Number(e.Hours || 0), 0);
                  return (
                    <div key={dayStr} className="tscell" style={{ margin: 8, borderTop: "1px solid var(--line)", minHeight: 48, display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 600 }}>
                      {total ? `${total.toFixed(2)}h` : "—"}
                    </div>
                  );
                })}
              </div>
            );
          })}

          {entries.length === 0 && (
            <div style={{ gridColumn: "1 / -1", padding: 32, textAlign: "center", color: "var(--text-mute)" }}>
              No hours logged for this week.
            </div>
          )}
        </div>
      </div>

      <div className="card" style={{ marginTop: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span style={{ fontWeight: 600 }}>Week total</span>
        <span className="mono" style={{ fontSize: 20, fontWeight: 700 }}>{weekTotal.toFixed(2)}h</span>
      </div>
    </div>
  );
}
