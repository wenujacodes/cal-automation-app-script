# ⚠️ Limitations & Roadmap

---

## Current Limitations

### 1. Manual Timetable Update Required

**The biggest limitation.**

NSBM's SharePoint blocks all automated external file access. Every approach was tested and hit a wall:

| Method | Result |
|--------|--------|
| Direct SharePoint download URL | Returns HTML error page |
| Microsoft Graph API (drives/lists) | Returns empty — student accounts restricted |
| SharePoint sharing link via UrlFetchApp | Returns "Something went wrong" HTML |
| Power Automate | Same SharePoint restrictions apply |

The root cause is NSBM IT restricting what student accounts can access programmatically, even when the file appears publicly accessible in a browser (because the browser has session cookies that Apps Script doesn't).

**Current workaround:** When NSBM updates the timetable, manually download the new Excel file and update the `NSBM_Timetable` Google Sheet in Drive. The next 6AM sync picks up the changes automatically.

---

### 2. SharePoint Sharing Links Expire

SharePoint sharing tokens (`?e=xxxxx`) are session-based and time-limited. Even if a direct fetch worked momentarily, these links can't be hardcoded reliably into a script.

---

### 3. Timetable Format Dependency

The parser is built specifically for the NSBM FOC Year 1 timetable format (26.1 batch). If the university restructures the Excel layout — changes column positions, removes week numbers, or changes time string formatting — the parser will break and need to be updated manually.

---

### 4. Sync Duration

A full sync of ~80–100 events takes 2–3 minutes due to intentional `Utilities.sleep()` calls between Google Calendar API requests. This is by design to avoid rate limit errors, but it means the 6AM trigger might not finish until ~6:03AM.

---

### 5. Scope

Currently only supports:
- NSBM FOC Year 1 Semester 2 timetable format
- Single Google account (yours)
- English language calendar events only

---

## Roadmap

### Short Term

- [ ] **Chrome Extension** — runs in your browser when you visit the SharePoint timetable page, grabs the file silently and triggers a sync. Bypasses all API restrictions because it operates inside your authenticated browser session.
- [ ] **Year 1 Semester 1 support** — extend parser to handle S1 timetable format differences

### Medium Term

- [ ] **Web interface** — simple upload page where non-tech students can drop the timetable Excel file and connect their Google Calendar without touching Apps Script
- [ ] **Batch support** — handle multiple batches (26.1, 25.1, etc.) from a single deployment

### Long Term

- [ ] **Microsoft OAuth integration** — proper app registration via Azure AD, allowing the script to authenticate as the user and fetch files directly from SharePoint via Microsoft Graph API. Requires NSBM IT to allow external app registrations — currently unclear if permitted.
- [ ] **Multi-university support** — generalise the parser to handle different timetable formats from different universities

---

## Contributing

Found a bug or want to add something? Open an issue or pull request on GitHub.

If NSBM IT permissions change and Graph API access becomes available for student accounts, that's the highest-impact contribution possible — full automation with zero manual steps.

---

[← Back to README](README.md) · [📋 Setup Guide](SETUP.md) · [⚙️ Technical Docs](TECHNICAL.md)
