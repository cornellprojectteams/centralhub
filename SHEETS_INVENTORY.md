# Project Teams Safety & Operations — Complete Inventory

> **Hub location:** `~/Desktop/Project-Teams-Ops-Hub`  
> **Prepared for:** Carey Okal Manwa  
> **Owner:** Noah Hamm (`nhh5@cornell.edu`)  
> **Last updated:** June 28, 2026 (from downloaded exports + share emails)

---

## Folder Structure

```
Project-Teams-Ops-Hub/
├── downloads/          ← 14 exported files (xlsx, pdf, docx)
├── email-notifications/ ← 18 Google share notification emails (.eml)
├── SHEETS_INVENTORY.md ← This document
├── index.html          ← Cornell-themed documentation hub
└── analyze_files.py    ← Script to re-parse downloads
```

**Not downloaded (still online only):** Summer Contact List 2026, Production Space Status Tracking Emailer (Apps Script)

---

## Executive Summary

Cornell **Project Teams (SPT)** uses a network of Google Forms, Sheets, Docs, and Apps Script to manage:

| Domain | Primary assets |
|--------|----------------|
| **Space stewardship** | Space Status Report form → Space Status Tracking → email notifications |
| **ELL shift reporting** | ELL Shift Summary form → Responses sheet (~119 submissions) |
| **Access & training compliance** | Master Sheet (Qualtrics + Workday), Fabman Admin (API-driven) |
| **Workspace registry** | Project Teams Workspaces (~97 team spaces) |
| **Team directory** | Project Team Names (37 teams) |
| **Vehicle program** | Student Drivers (220 applicants, 368 reservations) |
| **Inventory** | Tool inventory, purchasing, and locks (4 functional tabs) |

---

## System Map

```
SPACE STEWARDSHIP
  Space Status Report (Form) ──auto──▶ Form Responses tab
       │                                    │
       │                                    ├── Pending Actions (derived, read-only)
       │                                    ├── (WIP) Analysis / Dashboard v2
       │                                    └── Team Contacts (37 teams)
       └──▶ Production Space Status Tracking Emailer (Apps Script, online)

ELL OPERATIONS
  ELL Shift Summary (Form) ──auto──▶ Form Responses 1 + Sheet1 (pivot/summary)

FABMAN / ACCESS
  Sign-up Form tab ──trigger──▶ Fabman API (creates members + trainings)
  Dashboard tab ◀──sync──▶ Fabman API (325 members)
  Master Sheet ◀── Qualtrics exports from Canvas Safety & Training Hub

COMPLIANCE & SPACES
  Master Sheet (waivers by location) ←→ Workspaces (team space registry)
  Project Team Names (canonical 37-team list)
```

---

## Forms

### 1. Space Status Report

| | |
|---|---|
| **File** | `downloads/Space Status Report - Google Forms.pdf` |
| **Live form** | https://forms.gle/6DHYsEPonVtdALps5 |
| **Admin edit** | https://docs.google.com/forms/d/1bVzxR-lGNDtll6wh7wsdZGsVuqMkaRavVcGxykLHMFw/edit |
| **Output sheet** | Space Status Tracking → `Form Responses` tab |
| **Who fills it** | SPT student employees & staff only |
| **Frequency** | One submission per issue, as observed |

**Form fields:**

1. Email (auto)
2. NetID *
3. Responsible Team * — dropdown: AutoBoat, Baja, CUAir, Concrete Canoe, CU Sail, ChemE Car, CEV, Combat Robotics, CUAUV, CUBMD, DEBUT, DBF, ESW, EWB, EWH, Formula, Geodata, Hyperloop, iGEM, Mars Rover, Nexus, Rocketry, Seismic Design, Steel Bridge, Operations Team, Unknown/unsure
4. Current Status * — Red (same-day), Orange (2–3 days), Yellow (7–10 days), Ivory (14–21 days), Purple (as time permits)
5. Photo Upload (optional)
6. Issue Description * — Fire code, Egress/access, Work Practices, Chemical storage, Shared space, Housekeeping, Maintenance/facilities
7. Required Actions / Comments *

**Staff-added columns** (in tracking sheet): Email Tracking, Deadline, Cleared

**Manual:** `downloads/SPT Space Status Report Manual.pdf` (rev. 10/20/25, Jatin Mukerji)

**Legacy note:** Manual says issues should *also* be logged in **SPT Student Employee Log** during transition — that sheet was not included in downloads.

---

### 2. ELL Shift Summary

| | |
|---|---|
| **File** | `downloads/ELL Shift Summary.pdf` |
| **Live form** | https://docs.google.com/forms/d/e/1FAIpQLSfUSbqVXJ8si5c-efXQl8k5LAvvhkTSrSVPWlNCaKUEbNGgog/viewform |
| **Output sheet** | `downloads/ELL Shift Summary (Responses).xlsx` |
| **Who fills it** | ELL student employees during/end of shift |
| **Records** | ~119 form responses (Jan–2026) |

**Form fields:**

| Field | Options / notes |
|-------|-----------------|
| NetID * | |
| Shift * | 10-12, 12-2, 2-4, 4-6, 6-8, 8-10 |
| Overall Activity Level * | Intense 80-100%, High 60-80%, Moderate 40-60%, Low 20-40%, Quiet 0-20% |
| Main ELL use * | Computer/Study, Team Meetings, Assembly, Tours, Social, Empty |
| Composites Lab Use * | Sanding, Layup, Paint Booth, Construction, Testing, Cleaning, Empty |
| Auto Lab Use * | Fabrication, Assembly, Testing, Empty |
| Conference Room Use * | Meetings, Interviews, Work Sessions, Empty |
| Supply Status * | Gloves, First Aid, N95, Safety Glasses, Paper Towels, Cleaning Supplies, All stocked |
| Safety & Maintenance Issues * | Blocked exits, Equipment left on, Facility issues, Chemical waste, No issues, Other |
| Most active teams today * | Multi-select team list |
| Productivity notes | Free text |
| Additional Notes | Free text |
| Photos | Up to 5 images |

**Sheet tabs:**
- `Form Responses 1` — raw form data
- `Sheet1` — processed/summary view with date column

**Usage patterns (from data):** Peak shifts 8-10 PM (43), 6-8 PM (28); most common activity "Moderate 40-60%" (37 responses).

---

## Spreadsheets

### 1. Space Status Tracking

| | |
|---|---|
| **File** | `downloads/Space Status Tracking.xlsx` |
| **Link** | https://docs.google.com/spreadsheets/d/1mZrlnA-GiVKB4_Um21aMH9jsSkf6TUJhauDWivnf7-I/edit |
| **Population** | Auto from form + manual staff columns |
| **Records** | ~103 issues |

**Tabs:**

| Tab | Rows | Purpose |
|-----|------|---------|
| Form Responses | ~104 | Master issue log (form + staff fields) |
| Pending Actions | ~52 | Derived view of open items — **do not edit** |
| (WIP) Dashboard v2 | 1 | Placeholder for charts |
| (WIP) Analysis | ~65 | Pivot data for dashboard (#REF errors present — broken formulas) |
| Team Contacts | ~37 | Team name → email mapping for notifications |

**Top issue types:** Housekeeping (30), Fire code (23), Egress/access (16)  
**Top teams:** Baja (21), Formula (20), Operations Team (20)

---

### 2. ELL Shift Summary (Responses)

| | |
|---|---|
| **File** | `downloads/ELL Shift Summary (Responses).xlsx` |
| **Link** | https://docs.google.com/spreadsheets/d/1yazhxA8nwYWDRSdQEN6Cc3UKKnPn6PWF_b3zAe0KoQk/edit |
| **Population** | Auto from ELL Shift Summary form |

---

### 3. Master Sheet — Building Access and Workday Learning

| | |
|---|---|
| **File** | `downloads/Master Sheet - Building Access and Workday Learning.xlsx` |
| **Link** | https://docs.google.com/spreadsheets/d/1k97Ey5aNcm58zaviaXxKvJtTgL1yeSI_fvUkXZFEyM4/edit |
| **Population** | Qualtrics survey exports from **Safety & Training Hub Canvas course** + manual Workday tracking |

**Description (Read Me tab):** *"Compilation of student usage agreement waiver and training completions from Qualtrics surveys embedded in the Safety & Training Hub Canvas course"*

**Tabs:**

| Tab | ~Rows | Content |
|-----|-------|---------|
| Read Me | 1 | Description |
| ELL | 1,216 | Usage agreement sign-ups: DATE, TIME, NAME, netID, CUID, Grad Year, TEAM |
| HVL | 562 | High Voltage Lab access (+ "Approved by Noah") |
| RHODES PENTHOUSE | 366 | Rhodes penthouse cage access |
| FabMan and Machine Shop | 16 | Legacy — points to Fabman Admin sheet |
| Aquatic Center | 213 | Pool/aquatic facility access |
| Risk Waivers | 95 | General risk waivers |
| Private Vehicle Risk Waiver | 36 | Private vehicle use |
| Canoe waiver | 31 | Concrete Canoe / canoe-specific |
| Hollister room B55 | 1 | Hollister access (empty) |
| Tang Room 403 | 19 | Tang Hall room access |
| Workday Trainings | 79 | Name, Team, completed Workday courses (EHS modules) |
| Workday Training Backend | 86 | Raw Workday export data |
| Combined | 0 | Empty — intended merge tab? |

**Sample Workday courses tracked:** Chemical Waste (EHS 2716), Respiratory Protection (EHS 2381), Hot Works (EHS 2398), Compressed Gas Safety (EHS 2335), Electrical Safe Work Practices (EHS 4336)

---

### 4. Project Teams Fabman Admin

| | |
|---|---|
| **File** | `downloads/Project Teams Fabman Admin.xlsx` |
| **Link** | https://docs.google.com/spreadsheets/d/1lDAA4z1YexkjaABSPuXJyEXfSxN2QAsKWVk88TQCRLI/edit |
| **Population** | Fabman API sync (Dashboard) + Google Form (Sign-up) + Qualtrics imports |

**Tabs:**

| Tab | ~Rows | Purpose |
|-----|-------|---------|
| Dashboard | 326 | Live Fabman member records: ID, Name, Team, NetID, Email, State, Apron, Spaces, Lab Equipment, Machines, Mills, Supervisor |
| Actions | 2 | Bulk email list for batch Fabman operations |
| Qualtrics_2946 | 6 | Qualtrics survey export (attestation) |
| Qualtrics_2947 | 6 | Qualtrics survey export (attestation) |
| Sign-up Form | 460 | Form responses → triggers `onFormSubmit()` to create Fabman members via API |

**Sign-up Form fields:** Policy acknowledgment, First/Last Name, Project Team, Graduation Year, NetID, Previously Completed Training (Emerson Apron Status)

**Integration:** Documented in `Fabman API Guide.docx` and `Fabman SU25 Report.docx`. Dashboard has buttons to fetch/update members via Apps Script.

> ⚠️ **Security:** `Fabman API Guide.docx` contains a live Fabman API key. Do not commit to git or share publicly. Rotate if exposed.

---

### 5. Project Teams Workspaces

| | |
|---|---|
| **File** | `downloads/Project Teams Workspaces.xlsx` |
| **Link** | https://docs.google.com/spreadsheets/d/1cKd-5815CafamF2AQpwTIqATjdkD4HgM404vQvwTvuM/edit |
| **Population** | Manual + renewal survey data |
| **Records** | ~97 team space entries |

**Tabs:**

| Tab | ~Rows | Purpose |
|-----|-------|---------|
| Team Spaces Master List | 98 | Main registry |
| Reference ranges | 39 | Dropdown validation lists |
| Renewal survey data | 38 | Annual workspace renewal responses |
| UGs | 34 | Undergraduate contact reference by location |

**Master List columns:** Team, Space Category, Region/Building, Location, Room/area, SPT Access Form, Activity Waiver, Usage Agreement, Safety plan Required, Usage agreement link (Box.com), UG last updated, SP required, Contact, Notes

**Space categories:** Admin, Bench Work, Computing, Fabrication, Facility, Field, Lab, Shop, Storage

---

### 6. Project Team Names

| | |
|---|---|
| **File** | `downloads/Project Team Names.xlsx` |
| **Link** | https://docs.google.com/spreadsheets/d/1Md3bugxLvqHJQ7Htf3tHI-ARXeEQRFBU2zh6ho3XAfQ/edit |
| **Population** | Manual reference |
| **Records** | 37 teams |

**Columns:** Team, Abbr., Official Team Name, web address, notes

**Teams include:** Agua Clara, AppDev, AutoBoat, Baja Racing, CUAir, CUAUV, CEV, Combat Robotics, Concrete Canoe, Cornell Racing (FSAE), DEBUT, EWB, ESW, Hyperloop, iGem, Mars Rover, Rocketry, Steel Bridge, Project Teams Staff, and others.

**Note:** Form dropdown names (e.g. "Formula", "Baja") don't always match official names (e.g. "Cornell Racing", "Baja Racing") — normalization opportunity.

---

### 7. Project Teams Student Drivers

| | |
|---|---|
| **File** | `downloads/Project Teams Student Drivers.xlsx` |
| **Link** | https://docs.google.com/spreadsheets/d/1KwJLTkdhrQ0jD7-75p6IuhYvh3Fq0B13KAXf2Pcvk_M/edit |
| **Population** | Manual |
| **Vehicle** | Toyota Tacoma, plate MBN-4258 |

**Tabs:**

| Tab | ~Rows | Purpose |
|-----|-------|---------|
| Log | 370 | Vehicle reservation calendar |
| Approval | 221 | Driver application & training checklist |
| Reference | 41 | Team abbreviation lookup |
| Process | 26 | Step-by-step approval workflow instructions |
| Sheet7 | 0 | Empty |

**Approval tab fields:** Name, NetID, Grad Year, Team, Justification, Driver safety (RMI2100), Fleet application, Sign release, Trailer safety, Trailer test, etc.

**Process:** Submit rows via email to Paula, Carrie, and Nakeschi for fleet approval; update column P when filed.

**Reservation fields:** Team, Name, NetID, phone, date, start/end time, destination, parking location

---

### 8. Tool Inventory, Purchasing and Locks

| | |
|---|---|
| **File** | `downloads/Tool inventory  purchasing and locks.xlsx` |
| **Link** | https://docs.google.com/spreadsheets/d/1QZf3LbKOsuwsxeno3f5aDTACObMRXaRY1yX6_5Aj_lU/edit |
| **Population** | Manual |

**Tabs:**

| Tab | ~Rows | Purpose |
|-----|-------|---------|
| Team tool inventory | 25 | General tool list (drills, saws, routers, socket sets) |
| Mill room Tooling | 44 | Equipment procurement: supplier links, priority, P/N, price, units, ordered |
| Facility non recurring | 33 | One-time facility purchases (cord reels, festool vac, totes) |
| Locks | 33 | Lock assignments: Team, Equipment, Lock #, Key holder, NetID, date out/returned |

---

### 9. Summer Contact List 2026 *(online only)*

| | |
|---|---|
| **Link** | https://docs.google.com/spreadsheets/d/1UrI5mBI4rqY9NeElobhzdd1lWsAlBDeYa1LZPVWHNnE/edit |
| **Access** | View only |
| **Status** | Not downloaded — export when needed |

---

## Documents

### SPT Space Status Report Manual
- **File:** `downloads/SPT Space Status Report Manual.pdf`
- Operator guide for space status workflow

### Fabman API Guide
- **File:** `downloads/Fabman API Guide.docx`
- Technical guide: UrlFetchApp, GET/PUT/POST/DELETE, member/training management
- Contains Erica's Apps Script functions (`fetchFabmanMembers`, `updateFabmanMembers`, `onFormSubmit`)
- **Contains embedded API credentials — treat as confidential**

### Fabman SU25 Report
- **File:** `downloads/Fabman SU25 Report.docx`
- Summer 2025 summary: Dashboard usage, training course reorganization, supervisor schedule, van booking procedure

### Fabman-Canvas Training Integration
- **File:** `downloads/Fabman-Canvas Training Integration.docx`
- Proposed training hierarchy mapping Fabman courses to Canvas/Qualtrics attestations
- Documents apron color tiers (Blue/Green/Red) and equipment categories

---

## Automations (Online Only)

### Production Space Status Tracking Emailer
- **Link:** https://script.google.com/d/1DpKw0dXIQ5rH6_BQfIQnXpmTbTcPmfH3wq2Dzc6pJUA7lrmeFG7hh5P4/edit
- Sends team notification emails when space issues are logged
- Built by Erica, Swati, Charlie
- Input: Space Status Tracking sheet + Team Contacts tab

### Fabman Admin Scripts (in spreadsheet)
- Embedded in Project Teams Fabman Admin via Extensions → Apps Script
- Functions documented in Fabman API Guide

---

## Consolidation Recommendations

### High priority

| Issue | Recommendation |
|-------|----------------|
| **Master Sheet has 13 location tabs + empty Combined tab** | Build `Combined` tab with Location column; use filter views instead of separate tabs per building |
| **Team name inconsistency** (Form says "Formula", directory says "Cornell Racing") | Single lookup from Project Team Names with alias column for form dropdowns |
| **Fabman data in 3 places** (Fabman Admin Dashboard, Master Sheet FabMan tab, Sign-up Form) | Fabman Admin Dashboard is source of truth; deprecate Master Sheet FabMan tab |
| **Space Status Analysis tab has #REF! errors** | Fix formulas or rebuild dashboard from Form Responses |
| **ELL Responses has duplicate Sheet1 tab** | Consolidate processing into one analysis tab or Google Looker Studio |

### Medium priority

| Issue | Recommendation |
|-------|----------------|
| **Master Sheet + Workspaces overlap** (waivers vs. space registry) | Link via Team + Location keys; one compliance dashboard |
| **Workday Trainings + Workday Backend** (79 vs 86 rows) | Single Workday tab with import script |
| **Summer Contact List + Project Team Names** | Team Directory with seasonal contact columns |
| **Student Drivers Approval + Log** | Could become a form → sheet workflow for applications |
| **Tool inventory 4 tabs** | Keep separate (different purposes: tools, procurement, facility, locks) — already well-organized |

### Keep as-is

- Form → Response sheet pairs (Space Status, ELL Shift)
- Fabman Admin Dashboard ↔ API sync architecture
- Team Contacts tab in Space Status Tracking
- Pending Actions derived view (useful read-only dashboard)

---

## Automation Opportunities

| Area | Suggestion |
|------|------------|
| Space Status | Auto-set Deadline from status color; daily digest of open Red/Orange items |
| Space Status Analysis | Fix broken dashboard; weekly team stewardship report |
| ELL Shifts | Weekly activity summary; flag "No issues" streaks vs. recurring safety flags |
| Master Sheet | Automate Qualtrics → sheet import (currently manual paste?) |
| Workday | Scheduled CSV import to Workday Backend tab |
| Fabman | Already scripted — add training expiry alerts |
| Student Drivers | Reservation conflict detection; approval status notifications |
| Tool purchasing | Low-stock alerts when Ordered column is empty past deadline |

---

## Access Matrix

| Asset | Your access |
|-------|-------------|
| Space Status Report | Edit |
| Space Status Tracking | Edit |
| ELL Shift Summary | Respond |
| ELL Shift Summary Responses | Edit |
| Master Sheet | Edit |
| Fabman Admin | Edit |
| Workspaces | Edit |
| Team Names | View |
| Summer Contact List | View |
| Student Drivers | Edit |
| Tool Inventory | Edit |
| All Docs | View or Edit (per file) |
| Emailer Script | Edit |

---

## Key Contacts

| Name | Email | Role |
|------|-------|------|
| Noah Hamm | nhh5@cornell.edu | Primary owner |
| Erica Jiang | ej289@cornell.edu | Fabman scripts, emailer |
| Swati Sriram | ss4325@cornell.edu | Workflows |
| Charlie Cohen | cmc533@cornell.edu | Workflows |
| Jatin Mukerji | — | Space Status manual (Slack) |

---

## Open Questions for Team

1. What triggers the Space Status Emailer — form submit or time-based?
2. Is Qualtrics data imported manually into Master Sheet tabs?
3. Is the Combined tab in Master Sheet intended to be built?
4. Should SPT Student Employee Log be retired now?
5. Can we get Summer Contact List 2026 exported?
6. Who maintains Workday Training Backend imports?
7. Are Fabman Qualtrics tabs (2946/2947) still active surveys?

---

*Generated from local file analysis. Re-run `python3 analyze_files.py` after downloading updated exports.*
