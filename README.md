# BOM Tools - Excel Add-in

Adds a **BOM Tools** button on Excel's Home tab with three tabs:
1. **Add sheet** — imports the cleaned template (`Master Unique Materials` + hidden `BOM Blank`, without `Formulas` or `README` sheets), configures all 7 named ranges and dropdowns, and creates new take-off sheets with auto-renamed `bom_*` / `pl_*` tables. Includes a one-click **Repair dropdowns & formulas** tool for older workbooks.
2. **Materials** — view, search, add, or edit materials in `tblMat`, with optional GitHub Pages sync and publisher auto-sync.
3. **Convert** — creates a copy of the workbook and flattens active BOM sheets into either an **Auto-growing purchase list** or a **Fixed purchase list**, with optional **Compiled Purchase List** sheet.

## 1. Host the files
Host this folder on any HTTPS static host (GitHub Pages or Netlify Drop), then replace `https://YOUR-HOST` in `manifest.xml` with your base URL (no trailing slash):

    sed -i 's#https://YOUR-HOST#https://<you>.github.io/<repo>#g' manifest.xml

## 2. Install the manifest
- **Excel on the web**: Insert > Add-ins > Manage My Add-ins > Upload My Add-in > pick `manifest.xml`.
- **Excel for Windows**: place `manifest.xml` in a shared folder registered under File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs (tick *Show in Menu*), restart Excel, then Insert > My Add-ins > Shared Folder > BOM Tools.
- **Microsoft 365 organization-wide**: Admin center > Settings > Integrated apps > Upload custom apps.

## What was fixed in v1.1
- **No Formulas or README sheets created**: removed `Formulas` and `README` from the embedded template (`template-data.js` / `template.xlsx`) and added an import filter in `sync.js` so they are never left in project workbooks.
- **Working Materials dropdown (`C6:C2000`) & Primary Unit dropdown (`C5:C300+`)**: `tools.js` removes any shadowing worksheet-scoped names after `insertWorksheetsFromBase64`, recreates all 7 workbook-level named ranges (`Materials`, `LinWt`, `AreaWt`, `PieceWt`, `WtUnit`, `SheetL`, `SheetW`) using structured references with automatic A1-range fallback on Excel builds that reject structured references in `wb.names.add()`, and explicitly re-applies Data Validation on every imported/new sheet.
- **Formulas, layout & print setup preserved**: suspends calculation during template import, restores calculation and runs full recalculation after named ranges are ready, verifies/rewrites `M4`, `AB4`, `AD2`, `O2` and table references after sheet copy, restores `PO Weight` total/footer formulas when clearing a copied sheet, and enforces hidden Row 5, hidden Columns M & AB, freeze panes (rows 1–5), gridlines off, and Landscape / repeat rows `3:4`.
- **Converter skips hidden `BOM Blank`**: `convert.js` skips the hidden `BOM Blank` sheet during conversion/compilation and removes it cleanly alongside `Master Unique Materials`.
