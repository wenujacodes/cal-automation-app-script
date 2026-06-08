/**
 * ================================================================
 *  NSBM Timetable → Google Calendar Auto-Sync  (v2 - no Advanced Services needed)
 * ================================================================
 *  SETUP:
 *  1. Go to https://script.google.com → New project
 *  2. Delete default code, paste this entire file
 *  3. Save (Ctrl+S)  ← NO need to add Drive API service this time
 *  4. Run setupTrigger() once → authorize when prompted
 *  5. Run syncTimetable() to do the first sync
 * ================================================================
 */


function resetColorMap() {
  PropertiesService.getScriptProperties().deleteProperty("COLOR_MAP");
  console.log("COLOR_MAP reset done");
}

const COLORS = [
  "BLUE",
  "RED",
  "GREEN",
  "YELLOW",
  "PURPLE",
  "ORANGE",
  "CYAN"
];










function getColorMap() {
  const props = PropertiesService.getScriptProperties();
  const data = props.getProperty("COLOR_MAP");
  return data ? JSON.parse(data) : {};
}


const SHAREPOINT_URL = 'https://nsbm365.sharepoint.com/sites/SOC/_layouts/15/download.aspx?SourceUrl=/sites/SOC/IQDKQEYFJHblQq-gKPQvlQVSAVbQ7DjxC1yep7hTz0oxidw.xlsx';
const CALENDAR_NAME  = 'NSBM Lectures';
const TEMP_FILE_NAME = '_nsbm_tt_temp_';

const SKIP_WORDS = ['INAUGURATION', 'ORIENTATION', 'POYA', 'HOLIDAY'];

const TIME_COL = 1;           // Col B: "09.00 am - 10.00 am"
const DAY_COLS = [2,3,4,5,6]; // Cols C–G: Mon → Fri


// ── Entry Point ──────────────────────────────────────────────────
function syncTimetable() {
  console.log('▶ Sync started');

  const sheet = fetchAndConvert();
  if (!sheet) { console.log('✖ Could not load sheet'); return; }

  const events = parse(sheet);
  console.log('✔ Parsed ' + events.length + ' lecture events');

  applyToCalendar(events);
  cleanup();
  console.log('✔ Done');
}
function fetchAndConvert() {
  const file = DriveApp.getFilesByName("NSBM_Timetable").next();
  return SpreadsheetApp.open(file).getSheets()[0];
}

// ── Parser ───────────────────────────────────────────────────────
function parse(sheet) {
  const data = sheet.getDataRange().getValues();

  console.log("SAMPLE ROWS:");
for (let i = 0; i < Math.min(20, data.length); i++) {
  console.log(JSON.stringify(data[i]));
}

  const merges = buildMergeMap(
    sheet.getDataRange().getMergedRanges()
  );

  const events = [];
  let dates = {};

  for (let r = 0; r < data.length; r++) {
    const row = data[r];

    const detected = extractDates(row);
    if (Object.keys(detected).length >= 3) {
      dates = detected;
      continue;
    }

    const time = parseTimeSlot(String(row[TIME_COL] || ''));
    if (!time || !Object.keys(dates).length) continue;

    for (const col of DAY_COLS) {
      const val = String(row[col] || '').trim();
      if (!val || shouldSkip(val)) continue;

      const eventDate = dates[col];
      if (!eventDate) continue;

      const key = r + ',' + col;
      const lastRow = merges[key] ? merges[key].lastRow : r;
      const endSlot = parseTimeSlot(
        String(data[lastRow][TIME_COL] || '')
      );

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
  console.log("DATES DETECTED:", dates);

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
        map[r + ',' + c] = {
          lastRow: er,
          lastCol: ec
        };
      }
    }
  }

  return map;
}

// ── Calendar ─────────────────────────────────────────────────────
function applyToCalendar(events) {
  const cal = getOrCreateCalendar();

  // FULL CLEAN SLATE
  const now = new Date(2000, 0, 1);
  const far = new Date(2100, 0, 1);

  const existing = cal.getEvents(now, far);
  for (const e of existing) {
    e.deleteEvent();
  }

  const colorMap = getColorMap();
  let colorIndex = Object.keys(colorMap).length;

  const COLORS = ["BLUE","RED","GREEN","YELLOW","PURPLE","ORANGE","CYAN"];

  let added = 0;

  for (const e of events) {
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
  }

  saveColorMap(colorMap);

  console.log("✔ Fresh sync completed: " + added);
}

// ── Cleanup ───────────────────────────────────────────────────────
function cleanup() {
  const files = DriveApp.getFilesByName(TEMP_FILE_NAME);
  while (files.hasNext()) files.next().setTrashed(true);
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





function extractDates(row) {
  const out = {};

  for (const c of DAY_COLS) {
    const v = row[c];
    const d = (v instanceof Date)
      ? v
      : tryParseDate(String(v || ''));

    if (d) out[c] = d;
  }

  return out;
}


function tryParseDate(s) {
  const m = s.trim().match(/^(\d{1,2})-([A-Za-z]+)-(\d{2,4})$/);
  if (!m) return null;

  let yr = parseInt(m[3]);
  if (yr < 100) yr += 2000;

  const d = new Date(m[2] + ' ' + m[1] + ' ' + yr);
  return isNaN(d) ? null : d;
}



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
  const SKIP_WORDS = ['POYA'];
  const u = String(s).toUpperCase();
  return SKIP_WORDS.some(w => u.includes(w));
}

function setTime(date, h, m) {
  const d = new Date(date);
  d.setHours(h, m, 0, 0);
  return d;
}

const PROP_KEY = "SYNC_RUNNING";

function isLocked() {
  return PropertiesService.getScriptProperties().getProperty(PROP_KEY) === "1";
}

function lock() {
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, "1");
}

function unlock() {
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, "0");
}



function resetCalendar() {
  const cal = getOrCreateCalendar();

  const events = cal.getEvents(
    new Date(2000, 0, 1),
    new Date(2100, 0, 1)
  );

  for (const e of events) {
    e.deleteEvent();
  }

  console.log("✔ Calendar fully reset");
}

function saveColorMap(map) {
  PropertiesService.getScriptProperties().setProperty(
    "COLOR_MAP",
    JSON.stringify(map)
  );
}

function getOrCreateCalendar() {
  const list = CalendarApp.getCalendarsByName(CALENDAR_NAME);
  if (list.length > 0) return list[0];
  return CalendarApp.createCalendar(CALENDAR_NAME);
}
