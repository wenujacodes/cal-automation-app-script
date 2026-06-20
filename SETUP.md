# 📋 Setup Guide

> No coding knowledge needed. Follow these steps and you'll have your timetable in Google Calendar in under 10 minutes.

---

## What You Need

- A Google account (personal Gmail is fine)
- The NSBM timetable Excel file (download from SharePoint)
- About 10 minutes

---

## Step 1 — Download the Timetable

1. Go to the NSBM FOC SharePoint site
2. Find the latest timetable Excel file for your batch
3. Download it to your computer

---

## Step 2 — Upload to Google Drive

1. Go to [drive.google.com](https://drive.google.com)
2. Click **+ New → File upload**
3. Select the downloaded `.xlsx` timetable file

---

## Step 3 — Convert to Google Sheet

1. Once uploaded, double-click the file in Drive to open it
2. At the top click **Open with Google Sheets**
3. It will open as a Google Sheet — now rename it to exactly:

```
NSBM_Timetable
```

> ⚠️ The name must match exactly — no spaces before or after, no file extension.

---

## Step 4 — Set Up the Script

1. Go to [script.google.com](https://script.google.com)
2. Click **New project** (top left)
3. Delete everything in the editor (Ctrl+A → Delete)
4. Copy the full contents of `nsbm_calendar_sync.gs` from this repo and paste it in
5. Press **Ctrl+S** to save — give the project any name you like (e.g. "Timetable Sync")

---

## Step 5 — Enable Drive API

The script needs permission to access your Drive files.

1. In the Apps Script editor, look at the left panel
2. Click the **+** icon next to **Services**
3. Scroll down to find **Drive API**
4. Click **Add**

---

## Step 6 — Run the Setup

1. In the function dropdown at the top (where it says "Select function"), choose **`setupTrigger`**
2. Click **Run**
3. A popup will ask you to authorize — click **Allow**

This sets up the daily 6AM automatic sync.

---

## Step 7 — Do the First Sync

1. In the function dropdown, now choose **`syncTimetable`**
2. Click **Run**
3. Watch the **Execution log** at the bottom — it should say something like:

```
▶ Sync started
✔ Parsed 87 lecture events
✔ Fresh sync completed: 87 events added
✔ Done
```

---

## Step 8 — Check Your Calendar

Open [calendar.google.com](https://calendar.google.com) — you should see a new calendar called **NSBM Lectures** with all your events added and colour-coded by subject.

---

## Updating the Timetable

When NSBM releases an updated timetable:

1. Download the new Excel file from SharePoint
2. Open your existing `NSBM_Timetable` Google Sheet in Drive
3. Delete all the content and paste in the new data (or replace the sheet entirely)
4. The next morning's 6AM sync will automatically update your calendar

---

## Troubleshooting

**"File not found in Drive: NSBM_Timetable"**
The file name doesn't match. Go to Drive and make sure the Google Sheet is named exactly `NSBM_Timetable` with no extra spaces or characters.

**"No events parsed — aborting to protect calendar"**
The script couldn't read any events from the file. Make sure you opened the xlsx file as a Google Sheet (not kept it as xlsx format).

**Sync is taking a long time**
That's normal — the script deliberately pauses between API calls to avoid Google's rate limits. A full sync of ~80 events takes about 2–3 minutes.

**Getting failure emails from Google at 6AM**
Run `syncTimetable` manually once and check the execution log for the specific error. Paste it in the repo issues tab.

---

[← Back to README](README.md) · [⚙️ Technical Docs](CODE_OF_CONDUCT.md) · [⚠️ Limitations & Roadmap](LIMITATIONS.md)
