/* tools.js - "Add BOM sheet" and "Materials" pages of the BOM Tools task pane.
 * Runs inside Excel (Office.js). Nothing leaves the workbook.
 */
(function (root) {
  "use strict";

  var MASTER = "Master Unique Materials";
  var BLANK = "BOM Blank";
  var TITLE_PREFIX = "Bill of Materials for ";
  var UNITS = ["kg/m", "kg/m²", "kg/each"];
  var BOM_INPUT = ["Description", "Materials", "Length", "Width", "Thickness", "Qty", "Set", "Remarks"];
  var PL_INPUT = ["Stock", "PO Weight", "PO Qty", "PO Rate", "Remarks"];

  function isBlank(v) { return v === "" || v === null || v === undefined; }
  function isFormula(f) { return typeof f === "string" && f.charAt(0) === "="; }
  function lc(s) { return String(s).trim().toLowerCase(); }
  function noop() {}

  /* ---------------- pure helpers (unit-tested in node) ---------------- */
  function safeId(s) {
    return String(s).trim().replace(/[^A-Za-z0-9_]+/g, "_").replace(/^_+|_+$/g, "") || "Sheet";
  }

  function checkSheetName(name, existing) {
    var n = String(name || "").trim();
    if (!n) return "Type a name for the new sheet.";
    if (n.length > 31) return "Sheet names can be 31 characters at most.";
    if (/[:\\\/\?\*\[\]]/.test(n)) return "Sheet names cannot contain : \\ / ? * [ ]";
    if (/^'|'$/.test(n)) return "A sheet name cannot start or end with an apostrophe.";
    if (n.toLowerCase() === "history") return "\"History\" is reserved by Excel.";
    if (existing.some(function (e) { return lc(e) === lc(n); })) return "A sheet with that name already exists.";
    return "";
  }

  function uniqueName(base, taken) {
    var set = {};
    taken.forEach(function (t) { set[String(t).toLowerCase()] = 1; });
    var n = base, i = 2;
    while (set[n.toLowerCase()]) n = base + "_" + (i++);
    return n;
  }

  function num(v) {
    if (v === null || v === undefined || v === "") return null;
    var n = Number(String(v).replace(/,/g, ""));
    return isFinite(n) ? n : NaN;
  }
  function pos(v) { return typeof v === "number" && isFinite(v) && v > 0; }

  function normalizeMaterial(raw) {
    var r = {
      material: String(raw.material || "").trim(),
      category: String(raw.category || "").trim(),
      unit: String(raw.unit || "").trim(),
      remarks: String(raw.remarks || "").trim(),
      lin: num(raw.lin), area: num(raw.area), piece: num(raw.piece), sl: num(raw.sl), sw: num(raw.sw)
    };
    // keep only the fields that belong to the chosen unit (the master list never mixes them)
    if (r.unit !== "kg/m") r.lin = null;
    if (r.unit !== "kg/m²") { r.area = null; r.sl = null; r.sw = null; }
    if (r.unit !== "kg/each") r.piece = null;
    return r;
  }

  function validateMaterial(r, existingNames, mode) {
    var e = [];
    if (!r.material) e.push("Material name is required.");
    else if (r.material.charAt(0) === "=") e.push("Material name cannot start with \"=\".");
    else if (mode === "add" && existingNames.some(function (n) { return lc(n) === lc(r.material); }))
      e.push("\"" + r.material + "\" already exists in the master list.");
    if (UNITS.indexOf(r.unit) < 0) e.push("Pick a primary unit.");
    else if (r.unit === "kg/m" && !pos(r.lin)) e.push("Linear weight (kg/m) must be a number above 0.");
    else if (r.unit === "kg/each" && !pos(r.piece)) e.push("Piece weight (kg/each) must be a number above 0.");
    else if (r.unit === "kg/m²") {
      if (!pos(r.area)) e.push("Area weight (kg/m²) must be a number above 0.");
      if (!pos(r.sl) || !pos(r.sw)) e.push("Sheet length and width (mm) are needed for plates and sheets.");
    }
    return e;
  }

  /* ---------------- Excel helpers ---------------- */
  function firstTable(items, prefix) {
    for (var i = 0; i < items.length; i++) if (items[i].name.indexOf(prefix) === 0) return items[i];
    return null;
  }

  /* ================= Add BOM sheet ================= */
  async function inspectSheets() {
    var out = { bomSheets: [], allSheets: [], hasBlank: false, project: "" };
    await Excel.run(async function (context) {
      var wb = context.workbook;
      wb.worksheets.load("items/name,items/visibility,items/position");
      await context.sync();
      var sheets = wb.worksheets.items;
      sheets.forEach(function (ws) { ws.tables.load("items/name"); });
      await context.sync();
      out.allSheets = sheets.map(function (ws) { return ws.name; });
      var bom = sheets.filter(function (ws) { return !!firstTable(ws.tables.items, "bom_"); });
      var a1 = bom.map(function (ws) { return ws.getRange("A1").load("values"); });
      await context.sync();
      bom.forEach(function (ws, i) {
        out.bomSheets.push({ name: ws.name, hidden: ws.visibility !== "Visible" });
        if (ws.name === BLANK) out.hasBlank = true;
        var t = String(a1[i].values[0][0] || "");
        if (!out.project && ws.name !== BLANK && t.indexOf(TITLE_PREFIX) === 0 && t.indexOf("[") < 0)
          out.project = t.substring(TITLE_PREFIX.length).trim();
      });
    });
    return out;
  }

  async function clearInputs(context, table, names) {
    table.columns.load("items/name");
    await context.sync();
    var work = [];
    table.columns.items.forEach(function (c) {
      if (names.indexOf(c.name) >= 0) { var r = c.getDataBodyRange(); r.load("formulas"); work.push(r); }
    });
    await context.sync();
    work.forEach(function (r) {
      var changed = false;
      var out = r.formulas.map(function (row) {
        var v = row[0];
        if (!isFormula(v) && v !== "") { changed = true; return [""]; }
        return [v];
      });
      if (changed) r.formulas = out;          // formulas are written back unchanged, typed values are blanked
    });
    await context.sync();
  }

  async function addBomSheet(opts, log) {
    log = log || noop;
    var result = { name: "", warnings: [] };
    await Excel.run(async function (context) {
      var wb = context.workbook;
      wb.worksheets.load("items/name,items/position");
      wb.tables.load("items/name");
      await context.sync();

      var newName = String(opts.name || "").trim();
      var err = checkSheetName(newName, wb.worksheets.items.map(function (w) { return w.name; }));
      if (err) throw new Error(err);

      var src = wb.worksheets.getItemOrNullObject(opts.source);
      await context.sync();
      if (src.isNullObject) throw new Error("Source sheet \"" + opts.source + "\" was not found.");

      wb.worksheets.items.forEach(function (ws) { ws.tables.load("items/name"); });
      await context.sync();
      var srcSheet = wb.worksheets.items.filter(function (w) { return w.name === opts.source; })[0];
      var srcBom = firstTable(srcSheet.tables.items, "bom_");
      if (!srcBom) throw new Error("The source sheet has no BOM table (bom_...).");

      // put the new sheet after the last real BOM sheet (not after the hidden blank one)
      var rel = src, lastPos = -1;
      wb.worksheets.items.forEach(function (ws) {
        if (ws.name !== BLANK && firstTable(ws.tables.items, "bom_") && ws.position > lastPos) { rel = ws; lastPos = ws.position; }
      });

      log("Copying \"" + opts.source + "\" ...");
      var copy = src.copy(Excel.WorksheetPositionType.after, rel);
      copy.name = newName;
      copy.visibility = Excel.SheetVisibility.visible;
      copy.tables.load("items/name");
      await context.sync();

      var bomT = firstTable(copy.tables.items, "bom_"), plT = firstTable(copy.tables.items, "pl_");
      if (!bomT) { copy.delete(); await context.sync(); throw new Error("The copy has no BOM table. It was removed."); }

      // safety check: the copy must calculate from ITS OWN tables, not from the source sheet
      var f1 = copy.getRange("A6").load("formulas"), f2 = copy.getRange("M4").load("formulas");
      await context.sync();
      if ((f1.formulas[0][0] + " " + f2.formulas[0][0]).indexOf(srcBom.name + "[") >= 0) {
        copy.delete(); await context.sync();
        throw new Error("Excel left the new sheet's formulas pointing at \"" + opts.source + "\". The copy was removed; nothing was changed.");
      }

      // tidy table names to the template convention (bom_<sheet>, pl_<sheet>)
      var taken = wb.tables.items.map(function (t) { return t.name; });
      var id = safeId(newName);
      var bomName = uniqueName("bom_" + id, taken);
      var plName = uniqueName("pl_" + id, taken.concat([bomName]));
      log("Naming tables " + bomName + (plT ? " / " + plName : "") + " ...");
      bomT.name = bomName;
      if (plT) plT.name = plName;
      await context.sync();

      copy.getRange("A2").values = [[newName]];
      if (opts.project && String(opts.project).trim()) copy.getRange("A1").values = [[TITLE_PREFIX + String(opts.project).trim()]];
      await context.sync();

      if (opts.clear) {
        log("Clearing entered data ...");
        await clearInputs(context, bomT, BOM_INPUT);
        if (plT) await clearInputs(context, plT, PL_INPUT);
      }

      var f3 = copy.getRange("A6").load("formulas");
      await context.sync();
      if (String(f3.formulas[0][0]).indexOf(bomName + "[") < 0)
        result.warnings.push("Could not confirm the formulas follow the renamed table. Check the S.N. column on the new sheet.");

      copy.activate();
      copy.getRange("B6").select();
      await context.sync();
      result.name = newName;
    });
    return result;
  }

  /* ================= Materials master ================= */
  var HEAD = { material: "Material", category: "Category", unit: "Primary Unit", lin: "Linear Weight", area: "Area Weight",
               piece: "Piece Weight", sl: "Sheet Length", sw: "Sheet Width", remarks: "Remarks" };
  var EDITABLE = ["category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"];

  async function openMaster(context) {
    var ws = context.workbook.worksheets.getItemOrNullObject(MASTER);
    await context.sync();
    if (ws.isNullObject) throw new Error("Sheet \"" + MASTER + "\" not found. Open the automated template or project workbook (not a converted one).");
    var t = ws.tables.getItemOrNullObject("tblMat");
    await context.sync();
    if (t.isNullObject) throw new Error("Table \"tblMat\" not found on the Master sheet.");
    var hdr = t.getHeaderRowRange().load("values");
    await context.sync();
    var headers = hdr.values[0].map(function (s) { return String(s).trim(); });
    var idx = {};
    Object.keys(HEAD).forEach(function (k) {
      idx[k] = headers.findIndex(function (h) { return h.indexOf(HEAD[k]) === 0; });
    });
    ["material", "unit", "lin", "area", "piece", "sl", "sw"].forEach(function (k) {
      if (idx[k] < 0) throw new Error("Master table has no column starting with \"" + HEAD[k] + "\".");
    });
    return { t: t, idx: idx, ncols: headers.length };
  }

  function rowToObj(r, i, idx) {
    var g = function (k) { return idx[k] >= 0 ? r[idx[k]] : ""; };
    return { row: i, material: String(g("material")), category: String(g("category")), unit: String(g("unit")),
             lin: g("lin"), area: g("area"), piece: g("piece"), sl: g("sl"), sw: g("sw"), remarks: String(g("remarks")) };
  }

  async function listMaterials() {
    var out = { rows: [], categories: [] };
    await Excel.run(async function (context) {
      var m = await openMaster(context);
      var body = m.t.getDataBodyRange().load("values");
      await context.sync();
      var cats = {};
      body.values.forEach(function (r, i) {
        var o = rowToObj(r, i, m.idx);
        if (!o.material.trim()) return;
        out.rows.push(o);
        if (o.category) cats[o.category] = 1;
      });
      out.categories = Object.keys(cats).sort();
    });
    return out;
  }

  async function addMaterial(raw) {
    var r = normalizeMaterial(raw);
    await Excel.run(async function (context) {
      var m = await openMaster(context);
      var cur = m.t.getDataBodyRange().load("values");
      await context.sync();
      var names = cur.values.map(function (row) { return String(row[m.idx.material]); }).filter(function (s) { return s.trim() !== ""; });
      var errs = validateMaterial(r, names, "add");
      if (errs.length) throw new Error(errs.join(" "));

      m.t.rows.add();
      await context.sync();
      var body = m.t.getDataBodyRange().load("rowCount");
      await context.sync();
      var n = body.rowCount;
      var last = body.getRow(n - 1);
      var vals = { material: r.material, category: r.category, unit: r.unit, lin: r.lin, area: r.area, piece: r.piece, sl: r.sl, sw: r.sw, remarks: r.remarks };
      var inputIdx = {};
      Object.keys(vals).forEach(function (k) {
        if (m.idx[k] < 0) return;
        inputIdx[m.idx[k]] = 1;
        last.getCell(0, m.idx[k]).values = [[vals[k] === null ? "" : vals[k]]];
      });
      await context.sync();

      // make sure calculated columns (sheet area / weight) got their formulas; copy them from the row above if not
      if (n > 1) {
        var prev = body.getRow(n - 2);
        last.load("formulas"); prev.load("formulas");
        await context.sync();
        for (var c = 0; c < m.ncols; c++) {
          if (inputIdx[c]) continue;
          if (isFormula(prev.formulas[0][c]) && !isFormula(last.formulas[0][c]))
            last.getCell(0, c).copyFrom(prev.getCell(0, c), Excel.RangeCopyType.formulas);
        }
        await context.sync();
      }
    });
    return r;
  }

  async function updateMaterial(raw) {
    var r = normalizeMaterial(raw);
    await Excel.run(async function (context) {
      var m = await openMaster(context);
      var body = m.t.getDataBodyRange().load("values");
      await context.sync();
      var at = -1;
      body.values.forEach(function (row, i) { if (at < 0 && lc(row[m.idx.material]) === lc(r.material)) at = i; });
      if (at < 0) throw new Error("\"" + r.material + "\" is no longer in the master list.");
      var errs = validateMaterial(r, [], "edit");
      if (errs.length) throw new Error(errs.join(" "));
      var row = body.getRow(at);
      EDITABLE.forEach(function (k) {
        if (m.idx[k] < 0) return;
        row.getCell(0, m.idx[k]).values = [[r[k] === null ? "" : r[k]]];
      });
      await context.sync();
    });
    return r;
  }

  var api = {
    UNITS: UNITS, BLANK: BLANK,
    inspectSheets: inspectSheets, addBomSheet: addBomSheet,
    listMaterials: listMaterials, addMaterial: addMaterial, updateMaterial: updateMaterial,
    // pure helpers
    checkSheetName: checkSheetName, safeId: safeId, uniqueName: uniqueName,
    normalizeMaterial: normalizeMaterial, validateMaterial: validateMaterial
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomTools = api;
})(typeof window !== "undefined" ? window : this);
