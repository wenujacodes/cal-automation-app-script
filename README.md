# 📅 NSBM Timetable Sync

> Automatically syncs the NSBM Green University academic timetable to Google Calendar using Google Apps Script.

**By [@wenujacodes](https://github.com/wenujacodes)**

---

## 📌 Overview

NSBM shares timetables as Excel files via SharePoint. This project parses that file and creates colour-coded Google Calendar events for every lecture, tutorial, and lab — automatically, every morning at 6AM.

No manual entry. No copy-pasting. Just open your calendar.

---

## 📖 Table of Contents

- [How It Works](#how-it-works)
- [Quick Setup](#quick-setup)
- [Technical Deep Dive](#technical-deep-dive)
- [Current Limitations](#current-limitations)
- [Roadmap](#roadmap)

---

## How It Works

```
Google Drive (NSBM_Timetable)
         ↓
Google Apps Script reads the file
         ↓
Parser extracts dates, times, and subjects week by week
         ↓
Events created in Google Calendar (colour-coded per subject)
         ↓
Daily 6AM trigger keeps everything in sync
```

The script runs automatically every day. If the timetable file in Drive is updated, the next morning's sync will reflect those changes — wiping the old events and recreating them fresh.

---

## Quick Setup

> For students who just want it working. No coding knowledge needed.

**Step 1 — Download the timetable**

Download the latest timetable Excel file from the NSBM SharePoint (FOC site) and upload it to your Google Drive.

**Step 2 — Convert to Google Sheet**

In Google Drive, right-click the uploaded `.xlsx` file → Open with → Google Sheets. Once open, rename it exactly to:

```
NSBM_Timetable
```

**Step 3 — Set up the script**

1. Go to [script.google.com](https://script.google.com) and create a new project
2. Delete the default code
3. Paste the contents of `nsbm_calendar_sync.gs` from this repo
4. Save with `Ctrl+S`

**Step 4 — Enable Drive API**

1. In the Apps Script editor, click the `+` next to **Services** on the left panel
2. Find **Drive API** and click **Add**

**Step 5 — Authorize and run**

1. Select `setupTrigger` from the function dropdown and click **Run** — authorize when prompted
2. Select `syncTimetable` and click **Run** — this does the first sync

That's it. Check your Google Calendar for a new **NSBM Lectures** calendar.

**Updating the timetable**

When NSBM releases an updated timetable, download the new Excel file, open it in Google Sheets, and replace the content in your existing `NSBM_Timetable` file. The next 6AM sync will automatically pick up the changes.

---

## Technical Deep Dive

> For developers who want to understand or extend the project.

### Architecture

The project is a single Google Apps Script file with no external dependencies. It uses three Google services:

- **DriveApp** — to locate the timetable file in Google Drive
- **SpreadsheetApp** — to read and parse the spreadsheet data
- **CalendarApp** — to manage Google Calendar events

### File Structure

```
nsbm_calendar_sync.gs   ← entire project lives here
```

### Parsing Strategy

The timetable Excel file has a specific structure:

```
Row:  | Col A | Col B              | Col C–G (Mon–Fri)     |
------|-------|--------------------|-----------------------|
      |       | 1 (week number)    | date  date  date ...  |
      |       | 09.00 am-10.00 am  | WEB   ALGO  ...       |
      |       | 10.00 am-11.00 am  |                       |
      |       | 2 (week number)    | (ArrayFormula) ...    |
      |       | 09.00 am-10.00 am  | WEB   ALGO  ...       |
```

**Key challenge — ArrayFormulas:** Week 1 contains real date values. Weeks 2–15 use Excel ArrayFormulas for dates, which openpyxl and Google Sheets cannot evaluate. The parser solves this by reading real dates only from Week 1, then calculating all subsequent week dates by adding `(weekN - 1) * 7` days.

**Key challenge — Merged cells:** Multi-hour events (e.g. labs) span multiple rows via merged cells. The `buildMergeMap()` function pre-processes all merged ranges into a flat lookup map so the parser can find the correct end time for any event.

### Week Detection

```javascript
// Col B contains a small integer (1–20) on week header rows
const weekNum = row[1];
if (typeof weekNum === 'number' && weekNum >= 1 && weekNum <= 20 &&
    String(row[1]) === String(Math.floor(weekNum))) {
  // this is a week header row
}
```

### Date Calculation

```javascript
// Week 1: read real Date objects from the row
if (weekNum === 1) {
  for (const col of DAY_COLS) {
    if (row[col] instanceof Date) baseWeekDates[col] = row[col];
  }
}

// Week N: offset from week 1
const offset = (weekNum - 1) * 7;
const d = new Date(baseWeekDates[col]);
d.setDate(d.getDate() + offset);
```

### Skip Logic

Events are skipped if their cell value contains:
- `POYA` — covers Poya Day, Full Moon Poya Day, etc.
- `ADCERT` — Advanced Certificate course sessions (not relevant to degree students)

Exams (`WEB EXAM`, `ALGO EXAM`, etc.) are **included** intentionally.

### Rate Limiting

Google Calendar API throttles if you create/delete too many events rapidly. The script batches operations with a 1-second pause every 10 API calls and a 2-second buffer between the delete and create phases.

```javascript
if (i % 10 === 9) Utilities.sleep(1000);
```

### Concurrency Guard

A lock system using `PropertiesService` prevents two trigger instances from running simultaneously:

```javascript
function isLocked() {
  return PropertiesService.getScriptProperties().getProperty(PROP_KEY) === "1";
}
```

The `finally` block in `syncTimetable()` ensures the lock is always released even if an error occurs.

### Safe Wipe

The calendar is only wiped **after** the timetable is successfully parsed and at least one event is found. This protects against an empty calendar if the file is missing or the parse fails.

### Colour Mapping

Each unique subject (e.g. `WEB (Lec) 1`, `ALGO Tute`) gets a consistent colour assigned on first sync and persisted via `PropertiesService`. Colours cycle through: BLUE, RED, GREEN, YELLOW, PURPLE, ORANGE, CYAN.

---

## Current Limitations

**Manual file update required**

NSBM's SharePoint blocks automated external access — direct URL fetching, Microsoft Graph API, and Power Automate all return errors or empty responses for student accounts. The only reliable way to get the file is via a browser session. This means when NSBM updates the timetable, you need to manually update the `NSBM_Timetable` file in Google Drive.

**SharePoint sharing links expire**

SharePoint sharing tokens (`?e=xxxxx`) are session-based and expire. Automated fetching using these links does not work reliably outside a browser.

**Format dependency**

The parser is built specifically for the NSBM FOC Year 1 timetable format. If the university restructures the Excel layout, the parser will need to be updated.

**Google Calendar API rate limits**

Syncing a full semester (~80–100 events) takes 2–3 minutes due to intentional sleep delays between API calls. This is by design to avoid throttling errors.

---

## License

MIT — do whatever you want with it, just don't blame me if your calendar breaks 😄
