# BOM Tools Excel Add-in — v3.0 (Lightweight Rewrite)

A complete from-scratch rewrite of the BOM Tools Excel Add-in incorporating every feature from **v2.6**, the **Project-Only Material Lists & Plain-Sheet Name Repair** branch, and **Offline Windows LTS** compatibility — while making both the add-in bundle and generated Excel BOM sheets dramatically lighter and faster.

## What Makes v3.0 Lightweight for Excel

- **Lean 60-Row Expandable BOM Table (`bom_Take_off`)**: Starts at 60 rows (`A5:M65`) instead of 300 pre-allocated rows (80% fewer BOM formulas) and expands automatically as you type.
- **40-Material Purchase List (`pl_Take_off`)**: Upgraded capacity from 25 to 40 materials (`O5:AB48`) with O(1) `N(O5)+1` S.N. numbering.
- **Zero Volatile Print Pagination While Editing**: Removed volatile `OFFSET(...)` Print Area, Print Titles, and `fitToPage` from live editing sheets (print layout is applied cleanly when flattening for sharing).
- **96% Smaller Conditional Formatting**: Consolidated to 9 lean rules on active table rows (~1,280 cells vs 31,920 cells) with zero cross-sheet `INDEX(WtUnit, ...)` in CF.
- **Pre-Cached Formula Values & Clean View Setup**: Every formula cell in `template.xlsx` ships with a cached `<v>` value, `TableStyleNone`, and visible `freezeRows(4)` (`ySplit="4"`).
- **Zero Typing Lag / Direct2D Fix**: Every `Excel.run` call uses `{ delayForCellEdit: true }`; `BOM Blank` is temporarily unhidden & activated before `.copy()`; Row 6 scalar formulas are re-committed; and `finalizeSheetView()` runs after background refreshes.

## Complete Feature Set

- **Home-First Task Pane (`taskpane.html`, `app.js`)**: Visible `v3.0` / `v3.0 · Offline` badge, workbook auto-detection, 4 quick-action tiles, and calm live validation hints.
- **Zero-Reconnect Built-in Storage (`storage.js`)**: Automatically persists Global Master and isolated Project-Only lists across OPFS, IndexedDB, `localStorage`, and Offline Windows LTS API (`/api/health`, `/api/master`, `/api/projects`).
- **Project-Only Material Lists**: Isolate a workbook's materials with 4 start modes (Copy Global Master, Pick specific materials, Used in workbook only, Blank), export/import `.json` project files, manage saved projects on this PC, and protect Global Master from accidental overwrite.
- **Selective Two-Way Convert (`convert.js`)**:
  - **Flatten for Sharing**: Pick specific sheets (`All`/`None`) — partial flattening preserves `Master Unique Materials`, hidden `BOM Blank`, and named ranges. Choose **Auto-growing Excel 365** (`.formulas` dynamic spill), **Auto-growing Classic Excel** (non-CSE 2019/2021), or **Fixed Simplest**, plus optional **Compiled Purchase List**.
  - **Restore Automated**: Selective per-sheet restore with 3-way Materials picker (Global Master, Saved Project-Only list auto-matched via `BomProjectList`, or Project `.json` file).
  - **Upgrade / Repair Purchase-List Names**: Repairs broken or blank Purchase-List names on already-flattened plain sheets in place without touching BOM rows, Stock, or PO values.
- **Troubleshooting & Repair**: Strips legacy heavy print names, 2000-row CF/DV bloat, upgrades slow `COUNTA` S.N. & `AD2` formulas, and resets `freezeRows(4)` on selected BOM sheets.

## Files Included

| File | Role |
| --- | --- |
| `taskpane.html` | Lightweight Home-first task pane UI |
| `app.js` | Task pane controller, live guards, toast & session log |
| `storage.js` | Zero-reconnect built-in storage, Project-Only lists, Offline LTS bridge & GitHub sync |
| `sheets.js` | Workbook ops, template import, lag-free sheet creation, Materials CRUD & repair |
| `convert.js` | Selective Flatten (3 modes + Compiled), 3-way Restore & Plain-Sheet repair |
| `template.xlsx` | Lightweight 2-sheet Excel template (`BOM Blank` + `Master Unique Materials`) |
| `template-data.js` | Embedded base64 payload of `template.xlsx` |
| `manifest.xml` | Office Add-in manifest (replace `https://YOUR-HOST` with your HTTPS host) |
| `materials.json` | Optional hosted team materials file |
| `assets/icon-{16,32,64,80}.png` | Ribbon command icons |

## Hosting & Sideloading

1. Replace `https://YOUR-HOST` in `manifest.xml` with your HTTPS URL (GitHub Pages, Netlify, or `https://localhost:34873` for Offline Windows LTS).
2. Host all files in this folder over HTTPS.
3. Sideload `manifest.xml` in Excel via **Insert → Add-ins → My Add-ins → Upload My Add-in** (Web) or a Shared Folder Catalog (Windows Desktop).
