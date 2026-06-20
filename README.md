# 📅 NSBM Timetable Sync

> Automatically syncs the NSBM Green University academic timetable to Google Calendar using Google Apps Script.

**By [@wenujacodes](https://github.com/wenujacodes)**

---

## 📂 Navigation

| Page | Description |
|------|-------------|
| [📋 Setup Guide](SETUP.md) | Step-by-step setup for students — no coding needed |
| [⚙️ Technical Docs](TECHNICAL.md) | How the parser, sync, and rate limiting works |
| [⚠️ Limitations & Roadmap](LIMITATIONS.md) | Current known issues and future plans |

---

## What is this?

NSBM Green University shares academic timetables as Excel files via SharePoint. Manually checking and entering lectures into your calendar is a pain.

This project reads that timetable file and automatically creates colour-coded Google Calendar events for every lecture, tutorial, lab, and exam — every morning at 6AM, without you doing anything.

---

## How It Works

```
NSBM_Timetable file in Google Drive
             ↓
Google Apps Script reads and parses the file
             ↓
Extracts dates, times, and subjects week by week
             ↓
Creates colour-coded events in Google Calendar
             ↓
Daily 6AM trigger keeps everything up to date
```

---

## Features

- ✅ Colour-coded events per subject (WEB, ALGO, C#, SAD, ITSA...)
- ✅ Covers all 15 weeks including exams
- ✅ Skips Poya days and ADCERT sessions automatically
- ✅ 10-minute popup reminder on every event
- ✅ Safe sync — never wipes calendar if parsing fails
- ✅ Concurrency guard — won't run twice at the same time
- ✅ Rate limit safe — no more failure emails from Google

---

## Quick Start

Head over to the **[Setup Guide](SETUP.md)** to get it running in under 10 minutes.
