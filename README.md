# BOM Tools Excel Add-in — 3.0v patch2 (Excel 365 Lightweight Edition)

A complete from-scratch rewrite of the BOM Tools Excel Add-in targeting **Microsoft Excel 365** with a **pixel-identical template**, ultra-lightweight native Excel 365 formulas, **zero-hook input-lag elimination**, full v1.8–v3.0 materials migration, and isolated Project-Only material lists with full access to previously saved materials.

## What's Fixed & Engineered in 3.0v patch2

- **Bulletproof Template Import (Zero `InternalError`)**: `template.xlsx` uses 100% clean, validated standard OOXML (zero `xl/metadata.xml`, zero `cm="1"`, proper `r:id` namespaces) so `insertWorksheetsFromBase64()` succeeds reliably on every Excel build. Immediately upon import, `sheets.js` upgrades `BOM Blank` via native Excel `.formulas` to pure lightweight Excel 365 formulas (`LET`, `XMATCH`, `AL6` `SORT(UNIQUE(FILTER(...)))`, `AB4` `ROWS($AL$6#)`, O(1) `INDEX($AL$6#,...)`).
- **100% Visually Identical Template (`template.xlsx`)**: Preserves verbatim `xl/styles.xml` and `xl/theme/theme1.xml`, exact column widths, row heights, merges, cell styles, and conditional-formatting `dxfId` styles while trimming to 60 expandable BOM rows (`A5:M65`), 40-material Purchase List (`O5:AB48`), and zero volatile Print Area / Print Titles / `fitToPage`.
- **Root-Cause Fix for Persistent v1.8+ Input Lag**: Zero Office.js `Worksheet.onChanged` event listeners are registered; direct edits typed on `Master Unique Materials` are detected on-demand when visiting the Materials tab, refreshing, saving, or exporting. Never touches `freezePanes`, `pageLayout`, `showGridlines`, or `fullRebuild` on live BOM sheets.
- **Full Access to Previously Saved Materials (Global & Project-Only)**:
  - Automatically reads & reconciles ALL storage layers including legacy v1.8–v2.6 stores (`BomToolsLocalDB` IndexedDB, OPFS `master-materials.json` / `project-*.json`, `localStorage` `bomLocalMasterJson` / `bomProjectJson:*`), v3 layers, and Offline Windows LTS (`/api/master`, `/api/projects`), mirroring saves cleanly.
  - Project-Only lists auto-save to isolated per-project slots (never touching Global Master) while letting you **pick from all previously saved materials** (Global Master + Workbook + other saved Projects) both when creating a Project-Only list and anytime while a Project-Only list is active (`+ Add Saved`).
  - Restore Automated detects `BomProjectList` tags and shows a clear status banner indicating whether the project's saved copy exists on this PC or needs its `.json` file loaded.

## Files Included

| File | Role |
| --- | --- |
| `taskpane.html` | Lightweight Home-first Excel 365 task pane UI (`3.0v patch2`) |
| `app.js` | Zero-hook task pane controller, saved-library pickers, live guards, toast & session log |
| `storage.js` | Multi-layer built-in storage with v1.8–v3.0 migration, Project-Only lists & unified material pool |
| `sheets.js` | Lag-free Excel 365 sheet engine, bulletproof template import, bulk add & on-demand edit sync |
| `convert.js` | Selective Flatten (Excel 365 / Fixed + Compiled), 3-way Restore & Plain-Sheet repair |
| `template.xlsx` | Visually identical 27 KB template (`BOM Blank` + `Master Unique Materials`) |
| `template-data.js` | Embedded base64 payload of `template.xlsx` |
| `manifest.xml` | Office Add-in manifest v3.0.2.0 (replace `https://YOUR-HOST` with your HTTPS host) |
| `materials.json` | Optional hosted team materials file |
| `assets/icon-{16,32,64,80}.png` | Ribbon command icons |

## Hosting & Sideloading

1. Replace `https://YOUR-HOST` in `manifest.xml` with your HTTPS URL (GitHub Pages, Netlify, or `https://localhost:34873` for Offline Windows LTS).
2. Host all files in this folder over HTTPS.
3. Sideload `manifest.xml` in Excel via **Insert → Add-ins → My Add-ins → Upload My Add-in** (Web) or a Shared Folder Catalog (Windows Desktop).
