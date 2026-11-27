# BOM Tools Excel Add-in — v3.0 (Excel 365 Lightweight Edition)

A complete from-scratch rewrite of the BOM Tools Excel Add-in targeting **Microsoft Excel 365** with a **pixel-identical template**, ultra-lightweight native Excel 365 formulas, **zero-hook input-lag elimination**, bulletproof built-in materials saving, and isolated Project-Only material lists.

## What Was Fixed & Engineered for Speed

- **100% Visually Identical Template (`template.xlsx`)**: Preserves verbatim `xl/styles.xml` and `xl/theme/theme1.xml`, exact column widths, row heights, merges, cell styles, and conditional-formatting `dxfId` styles from the original template while trimming `sheet1.xml` from 653 KB to 214 KB.
- **Root-Cause Fix for Persistent v1.8+ Input Lag**: Removed the unconditional `Worksheet.onChanged` COM/WebView2 event hook (introduced in v1.8) that caused workbook-wide typing lag whenever the add-in was open. Zero Office.js event listeners are attached; direct edits typed on `Master Unique Materials` are detected on-demand when visiting the Materials tab, refreshing, saving, or exporting. Also eliminated `freezePanes`, `pageLayout`, `showGridlines`, `fullRebuild`, and trailing `Excel.run` calls on live BOM sheets.
- **Pure Excel 365 Lightweight Formulas**: Replaced `LOOKUP(2,1/...)`, 43× O(N²) `SMALL+COUNTIF`, `SUMPRODUCT`, and repeated `INDEX(WtUnit)` calls with native Excel 365 `XMATCH(..., 0, -1)`, `LET`, `XLOOKUP`, and a single hidden `AL6` `SORT(UNIQUE(FILTER(...)))` spill cell + O(1) `INDEX($AL$6#, ROW()-5)` lookup. Removed all Classic Excel 2019/2021 compatibility workarounds.
- **Lean 60-Row Expandable BOM + 40-Material Purchase List**: `bom_Take_off` starts at 60 rows (`A5:M65`) and expands automatically; `pl_Take_off` holds 40 materials (`O5:AB48`); conditional formatting & dropdowns are clamped to active table rows (95% fewer CF cells).
- **Bulletproof Materials Saving & Project-Only Isolation**:
  - Newest-timestamp reconciliation across OPFS, IndexedDB, `localStorage`, and Offline Windows LTS API (`/api/health`, `/api/master`, `/api/projects`) prevents stale-layer overwrites across Excel WebView2 restarts.
  - Project-Only mode isolates workbook materials (`BomProjectList` property + orange `A3` banner that clears cleanly when switching back), blocks Global Master reload/import/GitHub sync while active, requires 2-click confirmation before switching back to Global Master, and allows loading Project `.json` files even inside already-flattened workbooks for Restore.

## Complete Feature Set

- **Home-First Task Pane (`taskpane.html`, `app.js`)**: Visible `v3.0` / `v3.0 · Offline` badge, workbook auto-detection, 4 quick-action tiles, and calm live requirement guards.
- **Materials & Project-Only Lists (`storage.js`, `sheets.js`)**: Unobstructed Master list with BOM-ready validation icons, auto-collapsing change toast + Session log, audit columns (`Last Action`, `Last Modified`), isolated Project-Only lists (Copy Global, Pick specific, Used in workbook, Blank), saved PC project slots, JSON Backup & Transfer, and optional GitHub Sync.
- **Selective Two-Way Convert (`convert.js`)**:
  - **Flatten for Sharing**: Selective per-sheet checklist (partial flatten keeps Master, hidden BOM Blank & named ranges), **Auto-growing · Excel 365** (`LET` + `XLOOKUP` + dynamic spill, zero helper col M) or **Fixed · Print-Ready**, plus optional **Compiled Purchase List**.
  - **Restore Automated**: Selective per-sheet restore with 3-way Materials picker (Global Master, auto-matched Saved Project-Only list, or Project `.json` file) preserving Stock/PO values.
  - **Upgrade / Repair Purchase-List Names**: Repairs blank/broken Purchase-List names on flattened plain sheets in place with Excel 365 dynamic spill formulas.
- **Troubleshooting & Repair**: Purges legacy volatile print names, upgrades older v1.x/v2.x sheets to native Excel 365 `LET`/`XMATCH`/`AL6#` formulas, and clamps active-table dropdowns.

## Files Included

| File | Role |
| --- | --- |
| `taskpane.html` | Lightweight Home-first Excel 365 task pane UI |
| `app.js` | Zero-hook task pane controller, live guards, toast & session log |
| `storage.js` | Newest-timestamp built-in storage, Project-Only lists, Offline LTS bridge & GitHub sync |
| `sheets.js` | Lag-free Excel 365 sheet engine, template import, Materials CRUD & on-demand edit sync |
| `convert.js` | Selective Flatten (Excel 365 / Fixed + Compiled), 3-way Restore & Plain-Sheet repair |
| `template.xlsx` | Visually identical 26 KB Excel 365 template (`BOM Blank` + `Master Unique Materials`) |
| `template-data.js` | Embedded base64 payload of `template.xlsx` |
| `manifest.xml` | Office Add-in manifest (replace `https://YOUR-HOST` with your HTTPS host) |
| `materials.json` | Optional hosted team materials file |
| `assets/icon-{16,32,64,80}.png` | Ribbon command icons |

## Hosting & Sideloading

1. Replace `https://YOUR-HOST` in `manifest.xml` with your HTTPS URL (GitHub Pages, Netlify, or `https://localhost:34873` for Offline Windows LTS).
2. Host all files in this folder over HTTPS.
3. Sideload `manifest.xml` in Excel via **Insert → Add-ins → My Add-ins → Upload My Add-in** (Web) or a Shared Folder Catalog (Windows Desktop).
