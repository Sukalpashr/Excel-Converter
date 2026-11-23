# BOM Converter - Excel add-in (ribbon button + side panel)

Adds a **BOM Converter** button on Excel's Home tab. It opens a side panel with:
1. **Open a copy of this workbook** - opens a copy in a new window, so the original is never changed.
2. Two options: **Auto-growing purchase list** (your FlattenBOM.bas) or **Fixed purchase list** (your FlattenBOM1.bas), then **Convert**.

Everything runs inside Excel. No data is sent anywhere. The only thing hosted online is this static web page.

## 1. Host the files (once)
An Office add-in's panel is a web page, so it needs an https address. Any free static host works:
- **GitHub Pages**: put this folder in a repo, enable Pages. Your address is `https://<you>.github.io/<repo>`.
- **Netlify Drop**: drag this folder onto https://app.netlify.com/drop.

Then replace `https://YOUR-HOST` in `manifest.xml` with your address (no trailing slash):

    sed -i 's#https://YOUR-HOST#https://<you>.github.io/<repo>#g' manifest.xml

## 2. Install the manifest
- **Excel on the web**: Insert > Add-ins > Manage My Add-ins > Upload My Add-in > pick `manifest.xml`.
- **Excel for Windows (office / company)**: put `manifest.xml` in a shared network folder, then
  File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs, add the folder path, tick *Show in Menu*.
  Restart Excel, then Insert > My Add-ins > Shared Folder > BOM Converter.
- **Whole company**: a Microsoft 365 admin can push it under Admin center > Settings > Integrated apps > Upload custom apps.
  Staff then get the button with no install steps.

## Notes
- Needs ExcelApi 1.9 (Excel 2019 or Microsoft 365).
- The converters are ports of the VBA macros. Formula text for the growing list was checked against `flatten_bom.py`; it matches.
- Not tested inside Excel yet: try it on a copy of a real BOM first and compare with the VBA result.

## Shared materials list and template (new)

Put these two files in the repo root next to `taskpane.html`:
- `template.xlsx` - your automated template (must contain `BOM Blank`, `Master Unique Materials`, `Formulas`). The Add sheet tab imports it into any workbook that has no BOM template.
- `materials.json` - the shared master list (starts empty).

**Everyone** (no setup): Materials tab > *Update this workbook*, or leave "Update the master list from GitHub first" ticked on Add sheet. It adds missing materials and updates changed ones; it never deletes.

**Publisher computer(s) only**: create a fine-grained token (GitHub > Settings > Developer settings > Fine-grained tokens), repository access = this repo only, permission Contents = Read and write. Paste it under Materials > Publisher settings. After that, adding/editing a material in the add-in publishes it automatically, and *Publish to GitHub* sends anything you typed straight into the Master sheet.

Notes
- Changes reach other computers about a minute after publishing (Pages rebuild).
- To delete a material for everyone, edit `materials.json` on GitHub.
- A public repo makes `materials.json` and `template.xlsx` public.
- Template import needs ExcelApi 1.13 (Microsoft 365 / Excel 2021). Test it first on an empty workbook: check dropdowns and named ranges on the new sheet.
