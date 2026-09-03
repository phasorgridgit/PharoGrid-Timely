// Centralized calendar-date helpers.
//
// Every "day" this app cares about (a timesheet cell, a week boundary, a
// calendar column) is a calendar date, not a moment in time — "Mon Sep 7"
// means Sep 7 no matter what timezone the person happens to be sitting in.
//
// THE BUG THIS FILE EXISTS TO PREVENT: building a Date from local wall-clock
// fields (`new Date()`, `.setDate()`, `.setHours(0,0,0,0)`) and then turning
// it into a string with `.toISOString()` silently shifts the day backwards
// for anyone east of UTC (e.g. IST, UTC+5:30) — local midnight on Sep 7 is
// still Sep 6, 18:30 in UTC, so `.toISOString().slice(0,10)` returns
// "2026-09-06". That's exactly why a Monday entry was landing on the
// *previous* week: the day got silently decremented on its way to storage.
//
// The fix: never mix local wall-clock Date math with UTC-based
// serialization. A calendar date is read off local wall-clock fields
// EXACTLY ONCE (to find out "what day is it right now for this person"),
// then immediately re-anchored at UTC midnight. Every subsequent step —
// adding/subtracting days, finding the start of the week, converting back
// to a "YYYY-MM-DD" string — uses only the getUTC*/setUTC* family, so it's
// safe regardless of the viewer's timezone offset.

/** "Right now", anchored at UTC midnight of today's local calendar date. */
export function todayUTC() {
  return localDateToUTC(new Date());
}

/**
 * Reads the *local* wall-clock year/month/day off `d` (i.e. "what date is
 * this, on the wall calendar") and re-anchors it at UTC midnight so every
 * later step can use UTC-only math safely. Call this exactly once, at the
 * point a real local Date (like `new Date()`) first enters date math.
 */
export function localDateToUTC(d) {
  return new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
}

/**
 * Parses a "YYYY-MM-DD" or "YYYY-MM-DDT00:00:00Z"-style string (exactly how
 * dates are stored) as a UTC-anchored calendar date.
 */
export function parseDateOnly(dateStr) {
  const [y, m, d] = String(dateStr).slice(0, 10).split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}

/** Formats a UTC-anchored Date (see above) back to "YYYY-MM-DD". Safe only for dates produced by this module. */
export function formatDateOnly(d) {
  return d.toISOString().slice(0, 10);
}

/** Adds (or subtracts, with a negative n) whole days to a UTC-anchored Date, returning a new Date. */
export function addDays(d, n) {
  const copy = new Date(d);
  copy.setUTCDate(copy.getUTCDate() + n);
  return copy;
}

/** Monday = 0 ... Sunday = 6, for a UTC-anchored Date. */
export function weekdayIndex(d) {
  return (d.getUTCDay() + 6) % 7;
}

/** The Monday of the week containing a UTC-anchored Date. */
export function startOfWeek(d) {
  return addDays(d, -weekdayIndex(d));
}

/** True if two UTC-anchored Dates are the same calendar day. */
export function isSameDay(a, b) {
  return formatDateOnly(a) === formatDateOnly(b);
}

// --- Display formatting -----------------------------------------------
// A UTC-anchored Date must always be *displayed* in UTC too — calling the
// plain (local-timezone) `toDateString()`/`toLocaleDateString()` on one of
// these would re-introduce the same day-shift bug for anyone west of UTC
// (midnight UTC is still "yesterday evening" there). Passing
// `timeZone: "UTC"` pins the formatter to the calendar date the value
// actually represents, regardless of the viewer's local timezone.

/** e.g. "Mon Sep 07 2026" — for a UTC-anchored Date. */
export function formatLongDate(d) {
  return d.toLocaleDateString(undefined, {
    weekday: "short", month: "short", day: "2-digit", year: "numeric", timeZone: "UTC"
  });
}

/** e.g. "Sep 7" — for a UTC-anchored Date. */
export function formatShortDate(d) {
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric", timeZone: "UTC" });
}

/** e.g. "Sep 7, 2026" — for a UTC-anchored Date. */
export function formatShortDateWithYear(d) {
  return d.toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric", timeZone: "UTC" });
}
