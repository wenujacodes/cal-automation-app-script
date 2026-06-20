# ⚙️ Technical Documentation

> How the parser, sync engine, and supporting systems work under the hood.

---

## Stack

| Layer | Technology |
|-------|------------|
| Runtime | Google Apps Script (V8 engine) |
| File storage | Google Drive |
| Spreadsheet parsing | Google Sheets API via SpreadsheetApp |
| Calendar output | Google Calendar API via CalendarApp |
| State persistence | PropertiesService (key-value store) |
| Scheduling | Apps Script time-based triggers |

No external libraries. No backend server. Entirely serverless.

---

## File Structure

```
📁 NSBM-Timetable-Sync
└── nsbm_calendar_sync.gs    ← entire project in one file
```

---

## Timetable Structure

The NSBM timetable Excel file follows this layout:

```
Row  | Col A | Col B               | Col C     | Col D     | Col E     | Col F     | Col G
-----|-------|---------------------|-----------|-----------|-----------|-----------|----------
  9  |       |                     | Monday    | Tuesday   | Wednesday | Thursday  | Friday
 10  |       | 1  ← week number    | 22-Jun-26 | 23-Jun-26 | 24-Jun-26 | 25-Jun-26 | 26-Jun-26
 11  |       | 09.00 am - 10.00 am | WEB (Lec) | ALGO(Lec) |           |           | WEB (Lec)
 12  |       | 10.00 am - 11.00 am |           |           |           |           |
 ...
 19  |       | 2  ← week number    | formula   | formula   | formula   | formula   | formula
 20  |       | 09.00 am - 10.00 am | POYA DAY  | ALGO(Lec) |           |           | WEB (Lec)
```

**Column mapping:**
- `Col B (index 1)` — time slot strings or week numbers
- `Col C–G (index 2–6)` — Mon to Fri event cells

---

## Core Flow

```
syncTimetable()
    ├── isLocked() → bail if already running
    ├── lock()
    ├── fetchAndConvert() → returns Sheet object
    ├── parse(sheet) → returns events[]
    ├── guard: events.length === 0 → abort
    ├── applyToCalendar(events)
    └── unlock() ← always runs via finally
```

---

## Parser Deep Dive

### Week Detection

Col B contains a small integer (1–20) on week header rows. The check filters out time strings and subject names:

```javascript
const weekNum = row[1];
if (typeof weekNum === 'number' && weekNum >= 1 && weekNum <= 20 &&
    String(row[1]) === String(Math.floor(weekNum))) {
  // this is a week header row
}
```

### ArrayFormula Problem

Week 1 contains real `Date` objects in the date row. Weeks 2–15 use Excel ArrayFormulas (`=C10+7`) to calculate dates. Google Sheets cannot evaluate these — they come back as formula objects, not dates.

**Solution:** Read real dates only from Week 1, then calculate all other weeks by offset:

```javascript
// Week 1 — read real Date objects
if (weekNum === 1) {
  for (const col of DAY_COLS) {
    if (row[col] instanceof Date) baseWeekDates[col] = row[col];
  }
}

// Week N — calculate from base
const offset = (weekNum - 1) * 7;
const d = new Date(baseWeekDates[col]);
d.setDate(d.getDate() + offset);
dates[col] = d;
```

### Merged Cell Handling

Multi-hour events (labs, double lectures) span multiple rows via merged cells. The end time of a merged event is in the last row of the merge, not the first.

`buildMergeMap()` pre-processes all merged ranges into a flat lookup:

```javascript
// key: "rowIndex,colIndex" → { lastRow, lastCol }
map["10,2"] = { lastRow: 12, lastCol: 2 };
```

When building an event, the parser checks if the cell is merged and uses the `lastRow` to find the correct end time slot:

```javascript
const key = r + ',' + col;
const lastRow = merges[key] ? merges[key].lastRow : r;
const endSlot = parseTimeSlot(String(data[lastRow][TIME_COL] || ''));
```

### Time Parsing

Time strings follow the format `09.00 am - 10.00 am` (note: dot separator, not colon).

```javascript
function parseTimeSlot(s) {
  const m = s.match(
    /(\d{1,2})[.:](\d{2})\s*(am|pm)\s*[-–]\s*(\d{1,2})[.:](\d{2})\s*(am|pm)/i
  );
  // handles both . and : separators, both - and – dashes
}
```

### Skip Logic

```javascript
function shouldSkip(s) {
  const u = String(s).toUpperCase();
  return u.includes('POYA') || u.includes('ADCERT');
}
```

Catches: `POYA DAY`, `FULL MOON POYADAY`, `ADCERT 25.3 PYTHON`, `ADCERT 25.3 MATHS`, etc.

Exams (`WEB EXAM`, `ALGO EXAM`) are intentionally **not** skipped.

### Newline Cleanup

Some cells contain newline characters (e.g. `"ADCERT 25.3\nPYTHON"`). These are cleaned before processing:

```javascript
const val = String(row[col] || '').replace(/\n/g, ' ').trim();
```

---

## Calendar Sync

### Full Wipe Strategy

Every sync wipes all existing events and recreates them from scratch. This ensures the calendar always exactly matches the timetable — no stale events, no duplicates.

### Rate Limiting

Google Calendar API throttles bulk operations. The script pauses every 10 API calls:

```javascript
// During delete phase
if (i % 10 === 9) Utilities.sleep(1000);

Utilities.sleep(2000); // buffer between delete and create phases

// During create phase
if (i % 10 === 9) Utilities.sleep(1000);
```

### Colour Mapping

Subjects get consistent colours across syncs using `PropertiesService`:

```javascript
// Stored as JSON: { "WEB (Lec) 1": "BLUE", "ALGO Tute": "RED", ... }
PropertiesService.getScriptProperties().setProperty("COLOR_MAP", JSON.stringify(map));
```

Colour cycle: `BLUE → RED → GREEN → YELLOW → PURPLE → ORANGE → CYAN`

---

## Concurrency Guard

Apps Script triggers can overlap if a sync runs longer than the trigger interval. The lock system prevents this:

```javascript
const PROP_KEY = "SYNC_RUNNING";

function isLocked() {
  return PropertiesService.getScriptProperties().getProperty(PROP_KEY) === "1";
}
```

The `finally` block in `syncTimetable()` guarantees `unlock()` always runs — even on uncaught errors — so the script never gets permanently stuck.

---

## Safe Wipe Guard

The calendar is only wiped if parsing produced at least one event:

```javascript
if (events.length === 0) {
  console.log('⚠ No events parsed — aborting to protect calendar');
  return;
}
// only reaches here if events exist
applyToCalendar(events);
```

This protects against an empty calendar if the Drive file is missing, corrupted, or returns unexpected data.

---

## Utility Functions

| Function | Purpose |
|----------|---------|
| `resetCalendar()` | Wipe all events from the NSBM Lectures calendar |
| `resetColorMap()` | Clear the saved colour assignments — colours reassign on next sync |
| `setupTrigger()` | Register the daily 6AM trigger (run once during setup) |

---

[← Back to README](README.md) · [📋 Setup Guide](SETUP.md) · [⚠️ Limitations & Roadmap](LIMITATIONS.md)
