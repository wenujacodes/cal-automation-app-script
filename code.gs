/**
 * ================================================================
 *  NSBM Timetable → Google Calendar Auto-Sync
 * ================================================================
 *  SETUP:
 *  1. Go to https://script.google.com → New project
 *  2. Delete default code, paste this entire file
 *  3. Save (Ctrl+S)
 *  4. Run setupTrigger() once → authorize when prompted
 *  5. Make sure timetable is in Google Drive named: NSBM_Timetable
 *  6. Run syncTimetable() to do the first sync
 * ================================================================
 */

// ── Constants ─────────────────────────────────────────────────────
const COLORS = ["BLUE", "RED", "GREEN", "YELLOW", "PURPLE", "ORANGE", "CYAN"];
const CALENDAR_NAME = 'NSBM Lectures';
const TIME_COL = 1;            // Col B: "09.00 am - 10.00 am"
const DAY_COLS = [2, 3, 4, 5, 6]; // Cols C–G: Mon → Fri
const PROP_KEY = "SYNC_RUNNING";

// ── Entry Point ───────────────────────────────────────────────────
function syncTimetable() {
  if (isLocked()) {
    console.log('⚠ Sync already running, skipping');
    return;
  }

  lock();

  try {
    console.log('▶ Sync started');

    const sheet = fetchAndConvert();
    if (!sheet) {
      console.log('✖ Could not load sheet');
      return;
    }

    // Parse FIRST before touching the calendar
    const events = parse(sheet);
    console.log('✔ Parsed ' + events.length + ' lecture events');

    if (events.length === 0) {
      console.log('⚠ No events parsed — aborting to protect calendar');
      return;
    }

    // Only wipe + write if we actually got events
    applyToCalendar(events);
    console.log('✔ Done');

  } catch (e) {
    console.log('✖ Sync failed: ' + e.message);
  } finally {
    unlock();
  }
}

// ── Fetch Sheet ───────────────────────────────────────────────────
function fetchAndConvert() {
  const files = DriveApp.getFilesByName('NSBM_Timetable');
  if (!files.hasNext()) {
    console.log('✖ File not found in Drive: NSBM_Timetable');
    return null;
  }
  const file = files.next();
  return SpreadsheetApp.openById(file.getId()).getSheets()[0];
}

// ── Parser ────────────────────────────────────────────────────────
function parse(sheet) {
  const data = sheet.getDataRange().getValues();
  const merges = buildMergeMap(sheet.getDataRange().getMergedRanges());

  const events = [];
  let dates = {};
  let baseWeekDates = null;

  for (let r = 0; r < data.length; r++) {
    const row = data[r];

    // Detect week number rows (col B is a small integer 1-20)
    const weekNum = row[1];
    if (typeof weekNum === 'number' && weekNum >= 1 && weekNum <= 20 &&
      String(row[1]) === String(Math.floor(weekNum))) {

      if (weekNum === 1) {
        // Week 1 has real dates — read them normally
        baseWeekDates = {};
        for (const col of DAY_COLS) {
          const v = row[col];
          if (v instanceof Date) baseWeekDates[col] = v;
        }
        dates = baseWeekDates;
      } else if (baseWeekDates) {
        // All other weeks — calculate from week 1 + 7 days per week
        dates = {};
        const offset = (weekNum - 1) * 7;
        for (const col of DAY_COLS) {
          if (baseWeekDates[col]) {
            const d = new Date(baseWeekDates[col]);
            d.setDate(d.getDate() + offset);
            dates[col] = d;
          }
        }
      }
      continue;
    }

    const time = parseTimeSlot(String(row[TIME_COL] || ''));
    if (!time || !Object.keys(dates).length) continue;

    for (const col of DAY_COLS) {
      const val = String(row[col] || '').replace(/\n/g, ' ').trim();
      if (!val || shouldSkip(val)) continue;

      const eventDate = dates[col];
      if (!eventDate) continue;

      const key = r + ',' + col;
      const lastRow = merges[key] ? merges[key].lastRow : r;
      const endSlot = parseTimeSlot(String(data[lastRow][TIME_COL] || ''));

      events.push({
        title: val,
        start: setTime(eventDate, time.startH, time.startM),
        end: setTime(
          eventDate,
          endSlot ? endSlot.endH : time.endH,
          endSlot ? endSlot.endM : time.endM
        )
      });
    }
  }

  return events;
}

function buildMergeMap(ranges) {
  const map = {};
  for (const mr of ranges) {
    const sr = mr.getRow() - 1;
    const er = mr.getLastRow() - 1;
    const sc = mr.getColumn() - 1;
    const ec = mr.getLastColumn() - 1;
    for (let r = sr; r <= er; r++) {
      for (let c = sc; c <= ec; c++) {
        map[r + ',' + c] = { lastRow: er, lastCol: ec };
      }
    }
  }
  return map;
}

// ── Calendar ──────────────────────────────────────────────────────
function applyToCalendar(events) {
  const cal = getOrCreateCalendar();

  // Wipe existing events in batches
  const existing = cal.getEvents(new Date(2000, 0, 1), new Date(2100, 0, 1));
  for (let i = 0; i < existing.length; i++) {
    existing[i].deleteEvent();
    if (i % 10 === 9) Utilities.sleep(1000);
  }

  Utilities.sleep(2000); // breathe before creating

  const colorMap = getColorMap();
  let colorIndex = Object.keys(colorMap).length;
  let added = 0;

  for (let i = 0; i < events.length; i++) {
    const e = events[i];
    const subject = e.title.trim();

    if (!colorMap[subject]) {
      colorMap[subject] = COLORS[colorIndex % COLORS.length];
      colorIndex++;
    }

    const event = cal.createEvent(e.title, e.start, e.end);
    const colorEnum = CalendarApp.EventColor[colorMap[subject]];
    if (colorEnum) event.setColor(colorEnum);
    event.addPopupReminder(10);
    added++;

    if (i % 10 === 9) Utilities.sleep(1000);
  }

  saveColorMap(colorMap);
  console.log('✔ Fresh sync completed: ' + added + ' events added');
}

// ── Trigger Setup (run once) ──────────────────────────────────────
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncTimetable')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('syncTimetable')
    .timeBased().everyDays(1).atHour(6).create();

  console.log('✔ Daily 6 AM trigger set');
}

// ── Helpers ───────────────────────────────────────────────────────
function parseTimeSlot(s) {
  const m = s.match(/(\d{1,2})[.:](\d{2})\s*(am|pm)\s*[-–]\s*(\d{1,2})[.:](\d{2})\s*(am|pm)/i);
  if (!m) return null;
  return {
    startH: to24(parseInt(m[1]), m[3]),
    startM: parseInt(m[2]),
    endH: to24(parseInt(m[4]), m[6]),
    endM: parseInt(m[5])
  };
}

function to24(h, ampm) {
  const p = ampm.toLowerCase();
  if (p === 'pm' && h !== 12) return h + 12;
  if (p === 'am' && h === 12) return 0;
  return h;
}

function shouldSkip(s) {
  const u = String(s).toUpperCase();
  return u.includes('POYA') || u.includes('ADCERT');
}

function setTime(date, h, m) {
  const d = new Date(date);
  d.setHours(h, m, 0, 0);
  return d;
}

// ── Lock System ───────────────────────────────────────────────────
function isLocked() {
  return PropertiesService.getScriptProperties().getProperty(PROP_KEY) === "1";
}

function lock() {
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, "1");
}

function unlock() {
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, "0");
}

// ── Color Map ─────────────────────────────────────────────────────
function getColorMap() {
  const data = PropertiesService.getScriptProperties().getProperty("COLOR_MAP");
  return data ? JSON.parse(data) : {};
}

function saveColorMap(map) {
  PropertiesService.getScriptProperties().setProperty("COLOR_MAP", JSON.stringify(map));
}

// ── Reset Utilities ───────────────────────────────────────────────
function resetCalendar() {
  const cal = getOrCreateCalendar();
  const events = cal.getEvents(new Date(2000, 0, 1), new Date(2100, 0, 1));
  for (const e of events) e.deleteEvent();
  console.log('✔ Calendar fully reset');
}

function resetColorMap() {
  PropertiesService.getScriptProperties().deleteProperty("COLOR_MAP");
  console.log('✔ COLOR_MAP reset done');
}

function getOrCreateCalendar() {
  const list = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (list.length > 0) return list[0];
  return CalendarApp.createCalendar(CALENDAR_NAME);
}
