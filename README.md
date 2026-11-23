# BOM Tools - Excel add-in (ribbon button + side panel)

One panel, three tabs:

| Tab | What it does |
|---|---|
| **Add sheet** | Adds a take-off (BOM) sheet + purchase list to ANY workbook, from the built-in template (`template.xlsx`) or by copying a sheet already in the workbook. If the workbook has no master list, one is added too, filled with your latest materials. |
| **Materials** | Add / edit materials in the workbook's master list. Shows a "Latest master list" card when the shared list has materials this workbook lacks. |
| **Convert** | The original converter (automated workbook -> plain workbook). |

## Files to host (same folder on GitHub Pages)
`taskpane.html`, `taskpane.js`, `tools.js`, `sync.js`, `convert.js`, `template.xlsx`  (+ `assets/` icons as before)
`manifest.xml` is unchanged (optionally rename the labels to "BOM Tools").

## Updating the template
`template.xlsx` is your `.xltx` saved as a normal workbook. To change the template: edit the `.xltx`, **Save As .xlsx**
(keep the sheet names `Take-off` and `Master Unique Materials`, table names `bom_*`, `pl_*`, `tblMat`), upload it over `template.xlsx`,
and bump the `?v=` in `tools.js` (`TEMPLATE_URL`) so browsers fetch the new copy.

## How the shared master list works
* The add-in keeps ONE "latest master list" in this browser's local storage (per computer, shared by all workbooks).
* Whenever the panel opens on a workbook that has a master list, that workbook's materials are folded into the latest list
  (new materials are added; a material edited in that workbook replaces the older version).
* **Brand-new workbook** (nothing typed in any take-off sheet): its master list is refreshed automatically when the panel opens
  (new materials added; stale ones refreshed only if the workbook was synced before and has not changed them).
* **Workbook with items already entered**: never changed automatically. The Materials tab shows what is missing / different and an
  *Update* button (adds missing materials; a checkbox also overwrites differing values).
* **Add sheet in a workbook with no master list**: the master list sheet is inserted and filled with the latest list.
* Each workbook stores, inside the file, a record of what each material looked like when it was last in sync (workbook setting `bomMasterBase`).

## Limits
* Per computer: another PC has its own latest list. It rebuilds itself from any workbook you open with the panel.
* Clearing Office's web cache / browser data resets the latest list (workbooks keep their own master sheets).
* Deleting a material is not synced: it comes back from the latest list. Renaming is not supported (add the new name instead).
* Inserting the built-in template needs Excel 2021 / Microsoft 365 (ExcelApi 1.13). Copying an existing BOM sheet works on older versions.
* The panel must be open for the sync to run (click the ribbon button once per workbook; the pin icon on the panel keeps it open).
