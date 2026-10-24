# FST 16 Drivetrain Tracker — Architecture & Deployment

This document describes the system architecture, data models, synchronization pipelines, and security controls for the FST 16 Drivetrain Tracker web application.

---

## 1. System Overview

The FST 16 Drivetrain Tracker operates as a client-side Single Page Application (SPA) hosted on GitHub Pages. It replaces the previous server-dependent deployment model to achieve continuous 24/7 uptime without container hibernation or cold-start latency.

### Core System Attributes
- **Hosting:** GitHub Pages static site (served directly from the repository).
- **Execution:** 100% client-side HTML5, CSS3, and ES6 JavaScript.
- **Offline Mode:** Local Python HTTP server on port 8000.
- **Data Persistence:** Dual-layer storage combining local Excel master files and web JSON data files.
- **Synchronization:** Direct GitHub REST API integration for instant cloud commits.

```
+-------------------------------------------------------------------+
|                        Client Web Browser                         |
|  +---------------------+  +-----------------+  +---------------+  |
|  | Viewer Mode (Read)  |  | Editor Mode (PAT|  | SheetJS Export|  |
|  +----------+----------+  +--------+--------+  +-------+-------+  |
+-------------|----------------------|-------------------|----------+
              |                      |                   |
       HTTP GET (JSON)       HTTP PUT (REST API)     Blob Download
              |                      |                   |
+-------------v----------------------v-------------------v----------+
|                     GitHub Repository Remote                      |
|       - data/parts.json            - images/                      |
|       - data/assemblies.json       - images/requirements/         |
|       - data/requirements.json     - web/                         |
|       - data/trash.json            - FST16_*.xlsx                 |
+------------------------------------+------------------------------+
                                     ^
                         git pull / git push
                                     v
+-------------------------------------------------------------------+
|                      Local Engineering Machine                    |
|  +------------------------+      +-----------------------------+  |
|  | FST16_*.xlsx Workbooks | <--> | sync_cloud.py (Export/Import|  |
|  +------------------------+      +-----------------------------+  |
|                                                                   |
|  - 1_ABRIR_DASHBOARD.bat  (Local HTTP server port 8000)           |
|  - 2_ATUALIZAR_ONLINE.bat (Excel -> JSON -> Git push)            |
|  - 3_DESCARREGAR_CLOUD.bat(Git pull -> JSON -> Excel)            |
+-------------------------------------------------------------------+
```

---

## 2. Directory Structure

```
Drivetrain_Tracker/
|-- index.html                 # Top-level redirect with hash preservation
|-- sync_cloud.py              # Bidirectional Excel <-> JSON synchronizer
|-- test_tracker_backend.py    # Verification test suite
|-- ARCHITECTURE.md            # System technical documentation
|-- 1_ABRIR_DASHBOARD.bat      # Launches local web server & browser
|-- 2_ATUALIZAR_ONLINE.bat     # Exports Excel and pushes to remote
|-- 3_DESCARREGAR_CLOUD.bat    # Pulls remote and imports into Excel
|-- FST16_DT_Parts_Export.xlsx # Primary components workbook
|-- FST16_Assemblies_Export_DL.xlsx # Assemblies workbook
|-- FST16_Requirements.xlsx    # Technical requirements workbook
|-- FST16_Trash_Internal.xlsx  # Recycle bin storage workbook
|-- data/
|   |-- parts.json             # Serialized part iterations
|   |-- assemblies.json        # Serialized assembly compositions
|   |-- requirements.json      # Serialized FSG and team rules
|   |-- trash.json             # Serialized deleted iterations
|   `-- meta.json              # Export timestamp and metadata
|-- images/
|   |-- requirements/          # 5 source requirement diagrams
|   `-- *.png                  # CAD schematics and FEA captures
`-- web/
    |-- index.html             # Main portal interface
    |-- styles.css             # Unified dark design system
    |-- github-sync.js         # GitHub REST API client
    |-- app.js                 # Application router and UI logic
    `-- xlsx.full.min.js       # SheetJS standalone client library
```

---

## 3. Two-Tier Access Model

The web application enforces two operating tiers:

### Tier 1: Viewer Mode (Default)
- **Permissions:** Read-only access.
- **Capabilities:**
  - Browse all assemblies, components, and iterations.
  - Review mass delta comparisons against FST15 baselines.
  - Inspect specifications and GD&T tolerance tables.
  - Inspect FEA displacement tables (UP_01 9 LCs, HB_01 7 LCs).
  - Search and filter technical requirements.
  - View full-resolution schematic and requirement diagrams.
  - Export complete `.xlsx` workbooks directly in browser memory.
- **Security:** Requires no credentials or authentication tokens.

### Tier 2: Editor Mode
- **Permissions:** Full read and write capabilities.
- **Activation:**
  - Entering a GitHub Personal Access Token (PAT) via Cloud Settings.
  - Loading URL hash `#setup_token=<TOKEN>`. The application extracts the token into `localStorage` and cleans the URL via `window.history.replaceState`.
- **Additional Capabilities:**
  - `+ New Iteration`: Clones previous iteration values and creates new version records.
  - `Edit Iteration`: Updates mass, description, specifications, and FEA displacements.
  - `Delete Iteration`: Deletes non-sole iterations.
  - `Delete Part`: Moves all iterations of a part to `data/trash.json`.
  - `+ New Part`: Registers a new component in the tracker.
  - `Edit Tolerances`: Synchronizes GD&T table updates across all iterations of a part.
  - `Edit Justification`: Updates component design justifications across all iterations.
  - `Manage Assemblies`: Edits assembly mass baselines and targets.
  - `Recycle Bin`: Restores deleted components or empties trash.
  - Image Upload & Paste: Adds schematics with client-side downscaling.

---

## 4. Multi-Image Schema & Processing

### Data Representation
The `Image Path` field in `data/parts.json` supports both single strings and JSON arrays:
```json
"Image Path": "images/UP_01_v1.png"
```
```json
"Image Path": "[\"images/UP_01_v6.png\", \"images/UP_01_v6_iso.png\"]"
```

### Parsing Pipeline
`web/app.js` provides `parseImagePaths` and `formatImagePaths`:
1. **`parseImagePaths(val)`**: Converts single strings, nulls, or JSON arrays into a unified JavaScript array of normalized paths.
2. **`formatImagePaths(arr)`**: Converts arrays into a single string (1 image), JSON array string (>1 images), or null (0 images).

### Client-Side Canvas Downscaling
To protect the GitHub repository from storage bloat:
1. Uploaded or clipboard-pasted images load into an offscreen HTML5 `Image` element.
2. If dimensions exceed 1920x1080, an offscreen `Canvas` scales down the image while maintaining aspect ratio.
3. The downscaled image renders to PNG Base64 and commits to `images/` via `window.DTCloud.commitBinaryFile`.

---

## 5. Concurrency & Conflict Resolution

When multiple editors submit changes concurrently:
1. The GitHub REST API validates commit requests against the latest file SHA.
2. If an edit collides with a remote update, GitHub returns HTTP 409 Conflict.
3. `web/github-sync.js` catches HTTP 409, fetches the newest remote SHA, and retries the commit.

---

## 6. Client-Side Excel Export (SheetJS)

The web dashboard embeds `xlsx.full.min.js` to execute instant workbook exports without server roundtrips.
- **Trigger:** Left sidebar "Export Excel" button under "General".
- **Execution Time:** ~0 seconds (client-side heap processing).
- **Structure:**
  - Sheet 1: `Parts` (All iterations matching Excel schema).
  - Sheet 2: `Assemblies` (Subassembly targets and compositions).
  - Sheet 3: `Requirements` (Formula Student and team targets).
- **Output Filename:** `FST16_DT_Parts_Export.xlsx`.

---

## 7. Bidirectional Local Synchronization

Synchronization between Excel workbooks and JSON files runs via `sync_cloud.py`:

### Export (`sync_cloud.py export`)
1. Reads `FST16_DT_Parts_Export.xlsx`, `FST16_Assemblies_Export_DL.xlsx`, `FST16_Requirements.xlsx`, and `FST16_Trash_Internal.xlsx`.
2. Sanitizes types (handling Timestamps, NaNs, and JSON string arrays).
3. Writes formatted JSON records to `data/*.json`.

### Import (`sync_cloud.py import`)
1. Reads `data/parts.json`, `data/assemblies.json`, `data/requirements.json`, and `data/trash.json`.
2. Formats lists and dictionaries into valid JSON strings.
3. Formats openpyxl worksheets with standardized column widths, headers, and sorting.

---

## 8. Batch Scripts Reference

| Script | Purpose |
|---|---|
| `1_ABRIR_DASHBOARD.bat` | Starts local HTTP server on port 8000 and opens browser. |
| `2_ATUALIZAR_ONLINE.bat` | Exports Excel files to JSON, commits, and pushes to GitHub main branch. |
| `3_DESCARREGAR_CLOUD.bat` | Runs `git pull origin main` and imports cloud JSON files into local Excel. |
| `Sincronizar_Cloud_Excel.bat` | Runs local export (`sync_cloud.py export`) without git push. |

---

## 9. Verification & Quality Gates

Run the verification test suite before publishing updates:
```powershell
python Drivetrain_Tracker/test_tracker_backend.py
```

The test suite validates:
1. Zero forbidden bureaucratic terms in code and Excel sheets.
2. Zero emojis across all application files.
3. Strict 1 to 3 words length constraint on all UI headers, tabs, buttons, metrics, and inputs.
4. Data integrity and schema compliance for all data tables.
5. Web asset syntax integrity, absence of stray escape sequences, and active cache-busting query strings.

---

## 10. Post-Mortem & Web Deployment Protocol

### Incident Summary (October 2026)
- **Symptom:** Single Page Application frozen indefinitely on the placeholder screen: `Loading Tracker Data...`.
- **Primary Root Cause:** An automated script injection appended an escaped literal backslash-n (`\n`) into `web/app.js` outside of a string literal, producing an `Uncaught SyntaxError: Invalid or unexpected token`. The parser failed immediately upon script load, preventing DOM initialisation.
- **Secondary Propagation Cause:** Client web browsers aggressively cached the older broken script file. Even after repository fixes, browsers continued loading the faulty cached version from disk.

### Mandatory Preventative Protocol
1. **Pre-Commit Verification:** Run `python Drivetrain_Tracker/test_tracker_backend.py`. Gate 6 verifies that no literal unescaped control characters exist in `web/app.js` or `web/github-sync.js`.
2. **Mandatory Cache Busting:** Whenever modifying `web/app.js`, `web/github-sync.js`, or `web/styles.css`, increment the version timestamp query parameter (`?v=YYYYMMDD_XX`) across all asset tags in `web/index.html`.
3. **HTTP Cache Control:** Retain `<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">` in `web/index.html`.
4. **Post-Push Health Verification:** Check the live remote URL (`https://diogoribeir07.github.io/fst16-drivetrain/web/`) using a headless browser check or automated script to verify that `Wheel Assembly` renders and `Loading Tracker Data` clears.

