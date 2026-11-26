# BOM Tools - Excel Add-in

Adds a **BOM Tools** button on Excel's Home tab with four clean tabs:
1. **Home** — opens here by default with workbook status and quick-action cards for every common task, plus a collapsed *Troubleshooting & Repair* drawer at the bottom.
2. **Add Sheet** — uncluttered form (just Sheet name + Add BOM Sheet) with project name, copy source, clear data, and local sync tucked inside a *More options* dropdown, and *Troubleshooting & Repair* at the bottom.
3. **Materials** — unobstructed Master list (with clean BOM-ready status icons for Name, Primary Unit, and Weight) and streamlined Add/Edit/Remove editor with calm primary-requirement guards, aesthetic auto-collapsing change notifications with a compact **Session changes** log, and zero-setup **Built-in Add-in Storage + JSON Backup & Transfer** tucked neatly at the bottom.
4. **Convert** — two-way converter: **Flatten for Sharing** (Auto-growing, Fixed, and Compiled Purchase List) and **Restore Automated** (converts flattened BOM sheets right back into full automated BOM sheets with `Master Unique Materials`, hidden `BOM Blank`, dropdowns, formulas, and preserved Stock/PO values).

## 1. Host the files
Host this folder on any HTTPS static host (GitHub Pages or Netlify Drop), then replace `https://YOUR-HOST` in `manifest.xml` with your base URL (no trailing slash):

    sed -i 's#https://YOUR-HOST#https://<you>.github.io/<repo>#g' manifest.xml

## 2. Install the manifest
- **Excel on the web**: Insert > Add-ins > Manage My Add-ins > Upload My Add-in > pick `manifest.xml`.
- **Excel for Windows**: place `manifest.xml` in a shared folder registered under File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs (tick *Show in Menu*), restart Excel, then Insert > My Add-ins > Shared Folder > BOM Tools.
- **Microsoft 365 organization-wide**: Admin center > Settings > Integrated apps > Upload custom apps.

## 3. Built-in Computer Storage + JSON Backup & Transfer (Zero Setup, Never Reconnects)
1. **Works automatically out of the box** — every material addition, edit, or removal (from either the task pane or directly on the `Master Unique Materials` worksheet) saves silently inside Excel's built-in add-in storage (OPFS + IndexedDB + local storage) and **never asks to reconnect** after restarting Excel.
2. Whenever you create a **new BOM sheet**, add a BOM to a **non-BOM workbook**, or **restore a converted BOM**, your latest saved materials list loads automatically.
3. **Backup or move to another computer**: open **Materials > Backup & Transfer (JSON)** and click **Export JSON backup** to save `master-materials.json`, or **Import JSON...** on another computer to load it immediately.

## Project-only materials (one-time list for a single job) — v1.9
Use this when a job needs custom materials that must **not** end up in the shared master list.
1. **Materials > Project-only materials**: type a project name, choose *copy the whole master list* or *only the materials I pick*, then **Start project-only list**.
2. From then on, everything you add, edit or remove (in the task pane or directly on the `Master Unique Materials` sheet) changes **only this workbook's list**. It auto-saves in a separate slot on this computer and is never written to the shared master storage or GitHub. The Master sheet shows an orange *PROJECT-ONLY list* note in A3.
3. **Save project list to file** writes `project-materials-<name>.json`. **Convert > Flatten** offers to save this file automatically first, because the Master sheet is deleted when converting to plain Excel.
4. To bring it back: **Convert > Restore Automated > Project materials list > Load project list file...** (or **Materials > Project-only materials > Load project list file...** in a normal workbook). Restore then uses that list instead of the shared master.
5. **Switch back to shared master list** reloads the shared list into the workbook (project-only materials leave the sheet; the saved project file keeps them).
Project files carry `"type": "project"`, so the normal *Import JSON* in Backup & Transfer refuses them and your master list cannot be overwritten by accident.

## Plain sheets: names missing in the purchase list — fixed in v1.9
The auto-growing purchase list used a hidden helper column (M) that only existed on rows present at conversion time. A row inserted or typed later was counted but had no name. The purchase list now ranks names straight from the Materials column, so any row works. Sheets converted earlier can be fixed with **Convert > Restore Automated > Repair purchase-list names** (no re-flatten needed).
Note: for a bar material (kg/m) added after conversion, copy a neighbouring row (or its `L` cell) so the length-based quantity applies; other rows are counted per piece.

## What was updated in v1.1 – v1.8
- **Built-in Add-in Storage (Zero Reconnects), JSON Backup/Transfer & Calm Requirement Guards (v1.8)**: replaced external hard-drive file linking (which forced a reconnect prompt every Excel restart) with always-on **Built-in Add-in Storage** (OPFS + IndexedDB + `localStorage` + persistent storage request) paired with 1-click **Export JSON backup** / **Import JSON...** for backups or transferring between computers; added live primary-requirement guards across Materials, Add Sheet, Convert (Flatten & Restore), and Repair so actions cannot run until required inputs are valid (e.g. Material name, Primary unit, positive Weight/Sheet Size, valid non-duplicate Sheet name, copy confirmation, or at least one sheet selected), shown with a calm, non-annoying inline hint that disappears as soon as requirements are fulfilled.
- **Cost Removed, Aesthetic Auto-Collapsing Change Toast & Audit Metadata (v1.7)**: removed the unused **Unit cost** field and false Cost warning icons from Materials (validation now checks only BOM-required Name, Primary Unit, and Weight/Sheet Size); change notifications appear in an aesthetic card for ~5.5 seconds (pauses on hover) before collapsing into a compact **Session changes** button listing every modification made during the session; every add, modification, or removal stamps the timestamp and action type (`added` / `modified` / `removed`) into both the saved JSON (`lastChange` + per-material `lastAction`/`lastModified`) and the `Master Unique Materials` worksheet (Row 3 summary + `Last Action` / `Last Modified` columns).
- **Materials "More options" & Top Auto-Saved Feed (v1.6)**: optional fields (**Category**, **Remarks**) are tucked inside a collapsible **More options** dropdown so adding a material only asks for Name, Unit, and Weight, and changed materials bubble to the top of the list.
- **High-speed batched conversion & repair (v1.5)**: both **Flatten for Sharing** and **Restore Automated** (plus **Repair**) now switch Excel to Manual calculation mode with screen-update suspension, write BOM blocks and Purchase Lists as 2D range arrays instead of cell-by-cell calls, batch reads/writes across all sheets into a handful of `context.sync()` round-trips, tile Compiled Purchase List formats in a single `copyFrom`, and consolidate conditional formatting rules.
- **Per-sheet selection & copy option on Restore (v1.4)**: both **Restore Automated** and **Troubleshooting & Repair** let you pick exactly which detected sheets to convert or repair (with quick `All` / `None` toggles), and **Restore Automated** now includes **Step 1 · Open a copy in a new Excel window** just like Flatten for Sharing.
- **Home-first UI & streamlined panels (v1.3)**: opens to a **Home** tab with quick options; optional Add Sheet fields, Repair tools, and plain-English Local Master Save are neatly tucked into collapsible dropdowns so everyday editing stays fast and uncluttered.
- **Restore Automated BOM (v1.3)**: converts flattened/plain BOM sheets back into full automated BOM sheets with `Master Unique Materials`, hidden `BOM Blank`, `bom_*`/`pl_*` tables, live dropdowns, formulas, and preserved Stock/PO values.
- **Local computer master file sync (v1.2)**: uses the File System Access API + IndexedDB handle persistence + local computer storage cache so materials sync automatically across workbooks without a GitHub token; supports adding, editing, and removing materials (`deleteMaterial` + removal sync).
- **Material validation status icons (v1.2)**: displays per-row validation icons confirming whether Material Name, Category, Unit, and Cost are populated.
- **No Formulas or README sheets created**: removed `Formulas` and `README` from the embedded template (`template-data.js` / `template.xlsx`) and added an import filter in `sync.js` so they are never left in project workbooks.
- **Working Materials dropdown (`C6:C2000`) & Primary Unit dropdown (`C5:C300+`)**: `tools.js` removes any shadowing worksheet-scoped names after `insertWorksheetsFromBase64`, recreates all 7 workbook-level named ranges (`Materials`, `LinWt`, `AreaWt`, `PieceWt`, `WtUnit`, `SheetL`, `SheetW`) using structured references with automatic A1-range fallback on Excel builds that reject structured references in `wb.names.add()`, and explicitly re-applies Data Validation on every imported/new sheet.
- **Formulas, layout & print setup preserved**: suspends calculation during template import, restores calculation and runs full recalculation after named ranges are ready, verifies/rewrites `M4`, `AB4`, `AD2`, `O2` and table references after sheet copy, restores `PO Weight` total/footer formulas when clearing a copied sheet, and enforces hidden Row 5, hidden Columns M & AB, freeze panes (rows 1–5), gridlines off, and Landscape / repeat rows `3:4`.
- **Converter skips hidden `BOM Blank`**: `convert.js` skips the hidden `BOM Blank` sheet during conversion/compilation and removes it cleanly alongside `Master Unique Materials`.
