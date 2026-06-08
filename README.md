# NSBM Timetable → Google Calendar Automation

## Overview

This project automatically converts the NSBM Faculty of Computing timetable into a dedicated Google Calendar.

The system:

* Reads timetable data from an Excel-based timetable source.
* Parses lecture schedules automatically.
* Creates calendar events in Google Calendar.
* Assigns consistent colors for each subject.
* Adds reminders to lectures.
* Prevents duplicate event creation.
* Supports timetable synchronization updates.

---

## Features

### Automatic Calendar Creation

If the target calendar does not exist, the system creates it automatically.

Example:

```
NSBM Timetable AUTO
```

---

### Lecture Event Generation

The timetable is parsed and converted into Google Calendar events.

Each event contains:

* Lecture name
* Start time
* End time
* Reminder notification

Example:

```
DBMS (Lec) 1
09:00 AM - 11:00 AM
```

---

### Subject Color Coding

Each module receives a persistent color.

Example:

| Subject                  | Color  |
| ------------------------ | ------ |
| DBMS                     | Blue   |
| Programming Fundamentals | Red    |
| AI                       | Green  |
| Mathematics              | Yellow |

Colors remain consistent between syncs.

---

### Duplicate Prevention

Each lecture generates a unique identifier based on:

```
Lecture Name
Start Time
End Time
```

The system stores created event IDs and avoids recreating existing events.

---

### Event Cleanup

If a lecture is removed from the timetable:

* The event is automatically deleted from Google Calendar.
* Removed events are cleaned during synchronization.

---

### Automatic Synchronization

A time-based trigger can run:

* Every 6 hours
* Daily
* Any custom interval

Example:

```javascript
ScriptApp.newTrigger('syncTimetable')
  .timeBased()
  .everyHours(6)
  .create();
```

---

## Project Architecture

```
Timetable Source
        │
        ▼
 Google Apps Script
        │
        ▼
 Timetable Parser
        │
        ▼
 Event Generator
        │
        ▼
 Google Calendar
```

---

## Main Functions

### syncTimetable()

Main entry point.

Responsible for:

1. Loading timetable data
2. Parsing timetable
3. Creating/updating events
4. Removing outdated events

---

### parse()

Reads timetable structure and extracts:

* Dates
* Time slots
* Lecture names

Returns:

```javascript
[
  {
    title: "DBMS (Lec) 1",
    start: Date,
    end: Date
  }
]
```

---

### applyToCalendar()

Handles:

* Event creation
* Event updates
* Event deletion
* Color assignment
* Reminder creation

---

### buildMergeMap()

Detects merged cells inside the timetable.

Used for:

* Multi-hour lectures
* Accurate end-time calculation

---

### setupTrigger()

Creates automatic synchronization triggers.

Example:

```javascript
every 6 hours
every day
```

---

## Technologies Used

### Language

* JavaScript (Google Apps Script)

---

### Google Services

* Google Calendar API (Apps Script)
* Google Drive API (Apps Script)
* Google Sheets API (Apps Script)
* Properties Service

---

### Data Source

Current options tested:

#### Option 1

Manual Excel Upload

```
Excel File
      ↓
Google Drive
      ↓
Apps Script
      ↓
Google Calendar
```

Advantages:

* Stable
* Simple
* Works reliably

---

#### Option 2

SharePoint Direct Download

```
SharePoint
      ↓
Apps Script
      ↓
Google Calendar
```

Advantages:

* Fully automated

Disadvantages:

* SharePoint authentication and Excel conversion issues
* Less reliable

---

## Current Recommended Workflow

The most reliable workflow currently is:

```
Updated Timetable
        ↓
Download Excel
        ↓
Upload to Google Drive
        ↓
Run Sync
        ↓
Google Calendar Updated
```

This avoids SharePoint conversion problems.

---

## Setup Instructions

### 1. Create Google Apps Script Project

Go to:

https://script.google.com

Create a new project.

---

### 2. Paste Code

Replace the default code with the project source.

Save.

---

### 3. Configure Constants

Update:

```javascript
const CALENDAR_NAME = "NSBM Timetable AUTO";
```

Adjust if required.

---

### 4. Authorize

Run:

```javascript
syncTimetable()
```

Grant permissions when prompted.

---

### 5. Create Automatic Trigger

Run:

```javascript
setupTrigger()
```

This creates automatic synchronization.

---

## Maintenance

### Reset Calendar

```javascript
resetCalendar();
```

Deletes all events.

---

### Reset Color Map

```javascript
resetColorMap();
```

Clears subject color assignments.

---

### Hard Reset

```javascript
hardReset();
```

Deletes:

* Calendar events
* Stored event mappings

Useful during testing.

---

## Known Limitations

### SharePoint Integration

The NSBM timetable is hosted on SharePoint.

Google Apps Script cannot reliably:

* Open remote Excel files directly
* Convert SharePoint Excel files automatically
* Preserve merged cell formatting after conversion

Because of this, full SharePoint automation remains unreliable.

---

### Timetable Format Dependency

The parser assumes:

* Monday–Friday columns
* Fixed timetable structure
* Standard NSBM timetable layout

Major layout changes may require parser updates.

---

## Future Improvements

### Planned Features

* Google Calendar sharing links
* Multi-semester support
* Student-specific timetable generation
* Lecturer extraction
* Classroom extraction
* Email notifications
* Web dashboard
* Mobile app integration

---

## Author

Created by:

Wenuja Liyanamana

Project Goal:

Create a fully automated timetable synchronization system that keeps student calendars updated without manually creating lecture events.
