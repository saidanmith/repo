# Green Shield Environmental Reports - Extraction Guide

**Purpose:** Knowledge base for Admin Automator agent to perform data extraction tasks from Green Shield Asbestos Inspection Reports without requiring explanations.

---

## Report Structure & Field Locations

### Page 1 - Header Information

**Location:** Top of Page 1 (in the report header)

```
Site address: 18 Hobart Walk, (Workstream: Void), St Albans, AL3 6LS
WO: W5809640
Project No: G-25566
Report Issue: 1
Date: Apr 2026
```

**Fields to Extract:**

| Field | Format | Example | Notes |
|-------|--------|---------|-------|
| **Project No** | `G-xxxxx` | `G-25566` | Always starts with G-, followed by 5 digits. May include `(ISSUE x)` after it |
| **WO** | `Wxxxxxxx` | `W5809640` | Work Order number, starts with W, 7 digits follow |
| **UPRN** | `4-5 digits` | `012345` | Unique Product Reference Number (if no WO, look for UPRN) |
| **Workstream** | Plain name | `Void` | From text: `(Workstream: Void)` - extract just the workstream name |
| **Report Issue** | Number | `1` | `Report Issue: 1` or `(ISSUE 2)` in filename |
| **Date (Report)** | `MMM YYYY` | `Apr 2026` | From `Date: Apr 2026` on Page 1 |

---

### Section 2.0 - Details (Page 5-6)

**Location:** "2.0 Details" section (typically page 5-6)

**Fields to Extract:**

| Field | Table Row | Example | Notes |
|-------|-----------|---------|-------|
| **Date of Survey** | "Date of survey:" | `01/04/2026` | Format: DD/MM/YYYY |
| **Date of Report** | "Date of report:" | `02/04/2026` | Format: DD/MM/YYYY |
| **Survey Duration** | "Duration of survey:" | `1 day` | How long the survey took |
| **Lead Surveyor** | "Survey undertaken by:" | `Colin Welsh` | Name of surveyor |
| **Client** | "Client:" | `Morgan Sindall Property Services Limited` | Company/organization |
| **Date of Revisit** | "Date of revisit:" | `n/a` | Only present if ISSUE 2+ |
| **Duration of Revisit** | "Duration of revisit:" | `n/a` | Only present if ISSUE 2+ |
| **Revisit undertaken by** | "Revisit undertaken by:" | `n/a` | Only present if ISSUE 2+ |

---

### Section 4.0 - Executive Summary (Page 10-13)

**Location:** "4.0 Executive Summary" - contains risk categorized tables

**Asbestos Register Table** (Main data table with color-coded risk levels):

| Column | Example | Notes |
|--------|---------|-------|
| **Area** | `Main building` | Which building section |
| **Floor** | `Ground Floor` or `1st Floor` | Location within building |
| **Location** | `005 / Wetroom` | Room code and name |
| **Material** | `Textured coating to single skinned plasterboard ceiling` | ACM description |
| **Asbestos Type** | `Chrysotile` | Type of asbestos |
| **Material Assessment** | `Very Low Risk` | Risk level (color-coded) |
| **Action** | `Monitor and manage` | Recommended action |

**Risk Levels:**
- **HIGH RISK** (Red): > 10 points
- **MEDIUM RISK** (Orange): 7-9 points
- **LOW RISK** (Blue): 5-6 points
- **VERY LOW RISK** (Green): 2-4 points

---

## File Naming Convention

**Expected Format:**
```
G-xxxxx (ISSUE x) address line.pdf
```

**Examples:**
- `G-25566 18 Hobart Walk, St Albans, AL3 6LS.pdf`
- `G-25566 (ISSUE 2) 18 Hobart Walk, St Albans, AL3 6LS.pdf`
- `G-12345 (ISSUE 1) 42 Longlands, Hemel Hempstead, HP2 4DF.pdf`

**Notes:**
- Project number always first
- Issue number may appear after project number with variable spacing
- Address includes full street, town, postal code
- File extension always `.pdf`

---

## Common Extraction Tasks

### Task 1: Lookup & Populate Report Dates

**User Command:** *"Find reports for these job numbers and get their report dates into column B"*

**What the agent does:**
1. User provides: List of job numbers (e.g., `G-25566, G-12345, G-99999`)
2. User provides: Path to Excel file and column to populate
3. Agent locates each PDF by job number
4. Agent extracts: "Date of report:" from Section 2.0
5. Agent inserts date into corresponding Excel row

**Example:**
```
Input: Job No in Column A: G-25566
Output: Column B populated with: 02/04/2026
```

---

### Task 2: Find Survey Dates for All Reports in Folder

**User Command:** *"Get survey dates for all reports in Q:\SHARED FILES\folder and create a list"*

**What the agent does:**
1. User provides: Folder path
2. Agent finds all `.pdf` files matching pattern `G-xxxxx*.pdf`
3. Agent extracts from each: Project No + Date of Survey
4. Agent outputs list or Excel file with two columns: Job No | Survey Date

---

### Task 3: Validate Revisit Data

**User Command:** *"Find all ISSUE 2+ reports and flag any where revisit data is incomplete"*

**What the agent does:**
1. User provides: Folder path
2. Agent finds all PDFs with `(ISSUE 2)`, `(ISSUE 3)`, etc. in filename
3. Agent checks Section 2.0 for:
   - Date of Revisit = `n/a` ?
   - Duration of Revisit = `n/a` ?
   - Revisit undertaken by = `n/a` ?
4. Agent flags reports where ANY of these are `n/a`
5. Agent outputs list: `Job No | Date Revisit | Duration | Revisit By | Flag Status`

---

### Task 4: Extract Full ACM Register

**User Command:** *"Get all asbestos materials and risk levels from these reports into a spreadsheet"*

**What the agent does:**
1. User provides: Job numbers or folder path
2. Agent extracts from Section 4.0 - Executive Summary table:
   - Project No
   - Area
   - Floor
   - Location
   - Material
   - Asbestos Type
   - Material Assessment (Risk Level)
   - Action
3. Agent creates Excel with one row per ACM found
4. Agent color-codes Material Assessment column by risk level

---

### Task 5: Check for High/Medium Risk Materials

**User Command:** *"Find any HIGH or MEDIUM risk asbestos in these reports"*

**What the agent does:**
1. User provides: Job numbers or folder path
2. Agent reads Section 4.0, filters to: Material Assessment = `HIGH RISK` OR `MEDIUM RISK`
3. Agent outputs: Job No | Location | Material | Risk Level | Recommended Action
4. Agent flags as **URGENT** if any HIGH RISK found

---

## Excel Integration Patterns

### Pattern 1: Column Insertion
**User specifies:**
- Source: Folder path or list of job numbers
- Data to extract: (e.g., "Report dates")
- Target file: `C:\path\to\tracker.xlsx`
- Target column: (e.g., "Column B" or "Date Completed")
- Matching on: Job No (Column A assumed default)

**Agent behavior:**
- Opens Excel file
- Matches rows by Job No
- Inserts extracted data into specified column
- Saves file

---

### Pattern 2: New Extract File
**User specifies:**
- Source: Folder path
- Data to extract: (e.g., "All survey dates")
- Output: "Create new Excel file" or "Output to CSV"

**Agent behavior:**
- Creates new file or outputs data
- Names it descriptively: `Job_Numbers_and_Dates_[TIMESTAMP].xlsx`
- Includes headers and data

---

## Field Extraction Rules & Standards

### Date Extraction
- **Format conversion:** `Apr 2026` → `04/2026` or `02/04/2026` (depends on context/user preference)
- **When date is `n/a`:** Extract as literally `n/a` (not blank, not 0)
- **Revisit dates:** Only applicable if Report Issue > 1

### Text Extraction
- **Workstream:** Extract just the name (`Void`), not the full parenthetical
- **Surveyor names:** Extract full name as written
- **Location:** Include both code and name (e.g., `005 / Wetroom`)

### Risk Level Mapping
- HIGH RISK (Red) → `HIGH` or keep full text
- MEDIUM RISK (Orange) → `MEDIUM` or keep full text
- LOW RISK (Blue) → `LOW` or keep full text
- VERY LOW RISK (Green) → `VERY LOW` or `VL` (user preference)

---

## File Path Handling

**Expected format:**
```
Q:\SHARED FILES\25484 42 LONGLANDS, HEMEL HEMPSTEAD, HP2 4DF\G-25484  42 LONGLANDS, HEMEL HEMPSTEAD, HP2 4DF.pdf
```

**Parsing rules:**
- Paths may be network or local
- Project number in filename matches path folder (usually)
- Spaces in paths are standard - handle gracefully

---

## Error Handling & Edge Cases

| Scenario | Handling |
|----------|----------|
| PDF not found for job number | Report: `Job No: G-xxxxx - FILE NOT FOUND` |
| Field is `n/a` | Extract as `n/a` (literal value) |
| ISSUE number with variable spacing | `G-25566(ISSUE 2)` or `G-25566 (ISSUE 2)` - normalize to `(ISSUE 2)` |
| Multiple PDFs with same job number | Use most recent by file modification date, flag to user |
| Excel file locked/in use | Wait or prompt user to close file |
| Corrupted or unreadable PDF | Report: `Job No: G-xxxxx - PDF UNREADABLE` |

---

## Visual Identification Guides

### Page 1 Header (Quick Scan)
```
WO: W5809640
Project No: G-25566
Report Issue: 1
Date: Apr 2026
```
These four lines appear at the top-right of Page 1 in the report header.

### Section 2.0 Location
Appears after title pages. Starts with green banner: `2.0 Details`
Contains the metadata table with all survey details.

### Section 4.0 Location
Color-coded tables with risk levels. Easy to identify by colored header bars:
- RED banner = HIGH RISK section
- ORANGE banner = MEDIUM RISK section  
- BLUE banner = LOW RISK section
- GREEN banner = VERY LOW RISK section

---

## Future Enhancement Opportunities

- **Flag tracking:** Remember which reports have been processed
- **Batch validation:** Cross-check job numbers in Excel against reports found
- **Trend analysis:** Track when revisits are overdue
- **Auto-categorization:** Sort results by risk level or location
- **Report generation:** Create summaries by workstream, location, or risk level

