/* tools.js - "Add sheet" and "Materials" pages of the BOM Tools task pane.
 * Runs inside Excel (Office.js). Nothing leaves the workbook; the only network call is
 * loading template.xlsx from the same site as this page.
 */
(function (root) {
  "use strict";

  var Sync = (typeof module !== "undefined" && module.exports && typeof require === "function") ? require("./sync.js") : root.BomSync;

  var MASTER = "Master Unique Materials";      // sheet name inside template.xlsx
  var TEMPLATE_BOM = "Take-off";               // BOM sheet name inside template.xlsx
  var TEMPLATE_URL = "template.xlsx?v=1";
  var TEMPLATE_SRC = "__template__";           // value of the "Copy from" option for the built-in template
  var BLANK = "BOM Blank";
  var TITLE_PREFIX = "Bill of Materials for ";
  var UNITS = ["kg/m", "kg/m²", "kg/each"];
  var BOM_INPUT = ["Description", "Materials", "Length", "Width", "Thickness", "Qty", "Set", "Remarks"];
  var PL_INPUT = ["Stock", "PO Weight", "PO Qty", "PO Rate", "Remarks"];
  // workbook-level names the BOM sheets calculate with
  var NAME_COLS = { Materials: "[Material]", LinWt: "[Linear Weight (kg/m)]", AreaWt: "[Area Weight (kg/m²)]",
                    PieceWt: "[Piece Weight (kg/each)]", WtUnit: "[Primary Unit]", SheetL: "[Sheet Length (mm)]", SheetW: "[Sheet Width (mm)]" };

  function isBlank(v) { return v === "" || v === null || v === undefined; }
  function isFormula(f) { return typeof f === "string" && f.charAt(0) === "="; }
  function lc(s) { return String(s).trim().toLowerCase(); }
  function noop() {}
  function cellVal(v) { return v === null || v === undefined ? "" : v; }

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
      material: String(raw.material || "").trim(), category: String(raw.category || "").trim(),
      unit: String(raw.unit || "").trim(), remarks: String(raw.remarks || "").trim(),
      lin: num(raw.lin), area: num(raw.area), piece: num(raw.piece), sl: num(raw.sl), sw: num(raw.sw)
    };
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
  function supportsInsert() {
    try { return !!Office.context.requirements.isSetSupported("ExcelApi", "1.13"); } catch (e) { return false; }
  }

  var tplCache = null;
  async function getTemplateB64() {
    if (tplCache) return tplCache;
    var resp = await fetch(TEMPLATE_URL, { cache: "no-cache" });
    if (!resp.ok) throw new Error("Could not load the built-in template (template.xlsx, HTTP " + resp.status + "). Is it uploaded next to taskpane.html?");
    var u = new Uint8Array(await resp.arrayBuffer()), bin = "";
    for (var k = 0; k < u.length; k += 0x8000) bin += String.fromCharCode.apply(null, u.subarray(k, k + 0x8000));
    tplCache = btoa(bin);
    return tplCache;
  }

  /* ================= Master table access ================= */
  var HEAD = { material: "Material", category: "Category", unit: "Primary Unit", lin: "Linear Weight", area: "Area Weight",
               piece: "Piece Weight", sl: "Sheet Length", sw: "Sheet Width", remarks: "Remarks" };
  var EDITABLE = ["category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"];
  var ALL_FIELDS = ["material"].concat(EDITABLE);

  async function openMaster(context, soft) {
    var t = context.workbook.tables.getItemOrNullObject("tblMat");
    await context.sync();
    if (t.isNullObject) {
      if (soft) return null;
      throw new Error("This workbook has no master list (table \"tblMat\"). Add the master list sheet first.");
    }
    var hdr = t.getHeaderRowRange().load("values");
    await context.sync();
    var headers = hdr.values[0].map(function (s) { return String(s).trim(); });
    var idx = {};
    Object.keys(HEAD).forEach(function (k) { idx[k] = headers.findIndex(function (h) { return h.indexOf(HEAD[k]) === 0; }); });
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
  async function readRows(context, m) {
    var body = m.t.getDataBodyRange().load("values");
    await context.sync();
    var out = [];
    body.values.forEach(function (r, i) { var o = rowToObj(r, i, m.idx); if (o.material.trim()) out.push(o); });
    return out;
  }

  // write a Sync.plan() into the master table: refresh some rows in place, append the rest at the bottom
  async function applyPlan(context, m, plan) {
    var body = m.t.getDataBodyRange();
    plan.updates.forEach(function (u) {
      var row = body.getRow(u.row);
      EDITABLE.forEach(function (k) { if (m.idx[k] >= 0) row.getCell(0, m.idx[k]).values = [[cellVal(u.rec[k])]]; });
    });
    if (plan.updates.length) await context.sync();
    var n = plan.appends.length;
    if (!n) return;
    for (var i = 0; i < n; i++) m.t.rows.add();
    await context.sync();
    body = m.t.getDataBodyRange().load("rowCount");
    await context.sync();
    var first = body.rowCount - n, inputIdx = {};
    ALL_FIELDS.forEach(function (k) { if (m.idx[k] >= 0) inputIdx[m.idx[k]] = 1; });
    plan.appends.forEach(function (rec, j) {
      var row = body.getRow(first + j);
      ALL_FIELDS.forEach(function (k) { if (m.idx[k] >= 0) row.getCell(0, m.idx[k]).values = [[cellVal(rec[k])]]; });
    });
    await context.sync();
    if (first > 0) {                         // make sure calculated columns (sheet area / weight) have their formulas
      var prev = body.getRow(first - 1).load("formulas");
      var news = plan.appends.map(function (_, j) { return body.getRow(first + j).load("formulas"); });
      await context.sync();
      news.forEach(function (nr, j) {
        for (var c = 0; c < m.ncols; c++) {
          if (inputIdx[c]) continue;
          if (isFormula(prev.formulas[0][c]) && !isFormula(nr.formulas[0][c]))
            body.getRow(first + j).getCell(0, c).copyFrom(body.getRow(first - 1).getCell(0, c), Excel.RangeCopyType.formulas);
        }
      });
      await context.sync();
    }
  }

  // a workbook is "fresh" while no take-off sheet has any material typed in
  async function isFresh(context) {
    var tables = context.workbook.tables.load("items/name");
    await context.sync();
    var boms = tables.items.filter(function (t) { return t.name.indexOf("bom_") === 0; });
    boms.forEach(function (t) { t.columns.load("items/name"); });
    await context.sync();
    var ranges = [];
    boms.forEach(function (t) {
      t.columns.items.forEach(function (c) { if (c.name === "Materials") ranges.push(c.getDataBodyRange().load("values")); });
    });
    await context.sync();
    return !ranges.some(function (r) { return r.values.some(function (row) { return !isBlank(row[0]) && String(row[0]).trim() !== ""; }); });
  }

  /* One pass of the shared-master logic for the open workbook.
   *   always : fold the workbook's materials into the shared list (never changes the workbook)
   *   opts.mode "all"/"new" : write the shared list into this workbook's master (explicit / just inserted)
   *   opts.auto : if the workbook is fresh (empty take-off), refresh it; otherwise only report what is pending */
  async function syncWorkbook(context, opts) {
    opts = opts || {};
    var store = Sync.loadStore();
    var out = { has: false, rows: 0, pendingNew: 0, pendingDiff: 0, conflicts: [], applied: { added: 0, updated: 0 },
                fresh: false, storeCount: store.rows.length, storeAt: store.updatedAt, persistent: Sync.persistent() };
    var m = await openMaster(context, true);
    if (!m) return out;
    out.has = true;
    var rows = await readRows(context, m);
    var base = await Sync.readBase(context);
    var cap = Sync.capture(store, rows, base);
    if (cap.added || cap.updated) Sync.saveStore(store);
    out.conflicts = cap.conflicts;
    var mode = opts.mode || null;
    if (!mode && opts.auto) { out.fresh = await isFresh(context); if (out.fresh) mode = "auto"; }
    if (mode) {
      var plan = Sync.plan(store, rows, base, mode);
      if (plan.appends.length || plan.updates.length) {
        await applyPlan(context, m, plan);
        out.applied = { added: plan.appends.length, updated: plan.updates.length };
        rows = await readRows(context, m);
      }
    }
    var nb = Sync.newBase(store, rows, base);
    if (JSON.stringify(nb) !== JSON.stringify(base)) await Sync.writeBase(context, nb);
    var p = Sync.pending(store, rows);
    out.pendingNew = p.newCount; out.pendingDiff = p.diffCount;
    out.rows = rows.length; out.storeCount = store.rows.length; out.storeAt = store.updatedAt; out.persistent = Sync.persistent();
    return out;
  }

  async function masterStatus(opts) {
    var res;
    await Excel.run(async function (context) { res = await syncWorkbook(context, opts); });
    return res;
  }
  async function applyLatest(mode) { return masterStatus({ mode: mode === "all" ? "all" : "new" }); }

  /* ================= sheets: inspect ================= */
  async function inspectSheets() {
    var out = { bomSheets: [], allSheets: [], hasBlank: false, project: "", canTemplate: supportsInsert() };
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

  /* ================= sheets: add ================= */
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
      if (changed) r.formulas = out;
    });
    await context.sync();
  }

  // rename tables to the template convention, set the titles, optionally clear typed data, verify formulas
  async function finishBomSheet(context, ws, bomT, plT, newName, opts, takenTables, result, log) {
    var id = safeId(newName);
    var bomName = uniqueName("bom_" + id, takenTables);
    var plName = uniqueName("pl_" + id, takenTables.concat([bomName]));
    log("Naming tables " + bomName + (plT ? " / " + plName : "") + " ...");
    bomT.name = bomName;
    if (plT) plT.name = plName;
    await context.sync();
    ws.getRange("A2").values = [[newName]];
    if (opts.project && String(opts.project).trim()) ws.getRange("A1").values = [[TITLE_PREFIX + String(opts.project).trim()]];
    await context.sync();
    if (opts.clear) {
      log("Clearing entered data ...");
      await clearInputs(context, bomT, BOM_INPUT);
      if (plT) await clearInputs(context, plT, PL_INPUT);
    }
    var f = ws.getRange("A6").load("formulas");
    await context.sync();
    if (String(f.formulas[0][0]).indexOf(bomName + "[") < 0)
      result.warnings.push("Could not confirm the formulas follow the renamed table. Check the S.N. column on the new sheet.");
    ws.activate();
    ws.getRange("B6").select();
    await context.sync();
    result.name = newName;
  }

  // make sure the 7 workbook-level names exist and no sheet-level copy shadows them
  async function fixNames(context, wb, sheets, result) {
    sheets.forEach(function (ws) { ws.names.load("items/name"); });
    wb.names.load("items/name,items/formula");
    await context.sync();
    sheets.forEach(function (ws) {
      ws.names.items.forEach(function (n) { if (NAME_COLS[n.name.split("!").pop()]) n.delete(); });
    });
    var have = {};
    wb.names.items.forEach(function (n) { have[n.name.toLowerCase()] = n; });
    Object.keys(NAME_COLS).forEach(function (k) {
      var it = have[k.toLowerCase()];
      if (!it) wb.names.add(k, "=tblMat" + NAME_COLS[k]);
      else if (String(it.formula).toLowerCase().indexOf("tblmat[") < 0)
        result.warnings.push("The name \"" + k + "\" already exists in this workbook and does not point at the master list (" + it.formula + "). Fix it in Formulas > Name Manager.");
    });
    await context.sync();
  }

  async function insertFromTemplate(context, wb, names, log) {
    var b64 = await getTemplateB64();
    wb.worksheets.load("items/id");
    await context.sync();
    var before = {};
    wb.worksheets.items.forEach(function (w) { before[w.id] = 1; });
    log("Inserting " + names.join(" + ") + " from the built-in template ...");
    try {
      wb.worksheets.insertWorksheetsFromBase64(b64, { sheetNamesToInsert: names, positionType: Excel.WorksheetPositionType.end, importPageLayoutSettings: true });
      await context.sync();
    } catch (e) {            // some builds do not know the page-layout option
      wb.worksheets.insertWorksheetsFromBase64(b64, { sheetNamesToInsert: names, positionType: Excel.WorksheetPositionType.end });
      await context.sync();
    }
    wb.worksheets.load("items/id,items/name,items/position");
    await context.sync();
    var added = wb.worksheets.items.filter(function (w) { return !before[w.id]; });
    added.forEach(function (w) { w.tables.load("items/name"); });
    await context.sync();
    return added;
  }
  async function rollback(context, sheets) {
    try { sheets.forEach(function (s) { s.delete(); }); await context.sync(); } catch (e) { /* best effort */ }
  }

  async function addFromTemplate(context, wb, opts, log, result) {
    if (!supportsInsert()) throw new Error("The built-in template needs Excel 2021 or Microsoft 365 (ExcelApi 1.13). Copy from an existing BOM sheet instead.");
    var newName = String(opts.name || "").trim();
    wb.worksheets.load("items/name,items/position");
    wb.tables.load("items/name");
    await context.sync();
    var err = checkSheetName(newName, wb.worksheets.items.map(function (w) { return w.name; }));
    if (err) throw new Error(err);
    var taken = wb.tables.items.map(function (t) { return t.name; });
    var hasMaster = taken.some(function (n) { return n.toLowerCase() === "tblmat"; });

    wb.worksheets.items.forEach(function (ws) { ws.tables.load("items/name"); });
    await context.sync();
    var lastPos = -1;
    wb.worksheets.items.forEach(function (ws) {
      if (ws.name !== BLANK && firstTable(ws.tables.items, "bom_") && ws.position > lastPos) lastPos = ws.position;
    });

    var added = await insertFromTemplate(context, wb, hasMaster ? [TEMPLATE_BOM] : [TEMPLATE_BOM, MASTER], log);
    var bomWs = null, masterWs = null;
    added.forEach(function (w) {
      if (firstTable(w.tables.items, "bom_")) bomWs = w;
      else if (firstTable(w.tables.items, "tblMat")) masterWs = w;
    });
    if (!bomWs) { await rollback(context, added); throw new Error("The template did not contain a BOM sheet. Nothing was added."); }
    var bomT = firstTable(bomWs.tables.items, "bom_"), plT = firstTable(bomWs.tables.items, "pl_");

    // the new sheet must calculate from its own tables
    var chk = bomWs.getRange("A6").load("formulas");
    await context.sync();
    if (String(chk.formulas[0][0]).indexOf(bomT.name + "[") < 0) {
      await rollback(context, added);
      throw new Error("The inserted sheet's formulas did not link to its own table. Nothing was added.");
    }

    if (!hasMaster && !masterWs) { await rollback(context, added); throw new Error("The template did not contain the master list. Nothing was added."); }
    await fixNames(context, wb, added, result);

    bomWs.name = newName;
    if (lastPos >= 0) bomWs.position = lastPos + 1;
    await context.sync();

    if (!hasMaster) {
      log("Filling the new master list with your latest materials ...");
      var s = await syncWorkbook(context, { mode: "all" });
      result.master = { added: s.applied.added, updated: s.applied.updated, total: s.rows };
    }
    await finishBomSheet(context, bomWs, bomT, plT, newName, opts, taken, result, log);
  }

  async function addFromSheet(context, wb, opts, log, result) {
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
    var f1 = copy.getRange("A6").load("formulas"), f2 = copy.getRange("M4").load("formulas");
    await context.sync();
    if ((f1.formulas[0][0] + " " + f2.formulas[0][0]).indexOf(srcBom.name + "[") >= 0) {
      copy.delete(); await context.sync();
      throw new Error("Excel left the new sheet's formulas pointing at \"" + opts.source + "\". The copy was removed; nothing was changed.");
    }
    await finishBomSheet(context, copy, bomT, plT, newName, opts, wb.tables.items.map(function (t) { return t.name; }), result, log);
  }

  async function addBomSheet(opts, log) {
    log = log || noop;
    var result = { name: "", warnings: [], master: null };
    await Excel.run(async function (context) {
      var wb = context.workbook;
      if (opts.source === TEMPLATE_SRC) await addFromTemplate(context, wb, opts, log, result);
      else await addFromSheet(context, wb, opts, log, result);
    });
    return result;
  }

  // master list sheet only (for a workbook that has none), filled with the latest materials
  async function addMasterSheet(log) {
    log = log || noop;
    var result = { warnings: [], master: null };
    await Excel.run(async function (context) {
      var wb = context.workbook;
      if (!supportsInsert()) throw new Error("Needs Excel 2021 or Microsoft 365 (ExcelApi 1.13).");
      var m0 = await openMaster(context, true);
      if (m0) throw new Error("This workbook already has a master list.");
      wb.worksheets.load("items/name");
      await context.sync();
      if (wb.worksheets.items.some(function (w) { return lc(w.name) === lc(MASTER); }))
        throw new Error("A sheet called \"" + MASTER + "\" already exists but has no master table. Rename or delete it first.");
      var added = await insertFromTemplate(context, wb, [MASTER], log);
      if (!added.some(function (w) { return firstTable(w.tables.items, "tblMat"); })) {
        await rollback(context, added);
        throw new Error("The template did not contain the master list. Nothing was added.");
      }
      await fixNames(context, wb, added, result);
      var s = await syncWorkbook(context, { mode: "all" });
      result.master = { added: s.applied.added, updated: s.applied.updated, total: s.rows };
    });
    return result;
  }

  /* ================= Materials page ================= */
  async function listMaterials() {
    var out = { rows: [], categories: [], none: false };
    await Excel.run(async function (context) {
      var m = await openMaster(context, true);
      if (!m) { out.none = true; return; }
      var cats = {};
      out.rows = await readRows(context, m);
      out.rows.forEach(function (o) { if (o.category) cats[o.category] = 1; });
      out.categories = Object.keys(cats).sort();
    });
    return out;
  }

  async function publish(context, rec) {          // an explicit add/edit becomes the latest version everywhere
    var store = Sync.loadStore();
    Sync.upsert(store, rec);
    Sync.saveStore(store);
    var base = (await Sync.readBase(context)) || { v: 1, h: {} };
    base.h[lc(rec.material)] = Sync.hash(rec);
    await Sync.writeBase(context, base);
  }

  async function addMaterial(raw) {
    var r = normalizeMaterial(raw);
    await Excel.run(async function (context) {
      var m = await openMaster(context, false);
      var rows = await readRows(context, m);
      var errs = validateMaterial(r, rows.map(function (x) { return x.material; }), "add");
      if (errs.length) throw new Error(errs.join(" "));
      await applyPlan(context, m, { updates: [], appends: [r] });
      await publish(context, r);
    });
    return r;
  }

  async function updateMaterial(raw) {
    var r = normalizeMaterial(raw);
    await Excel.run(async function (context) {
      var m = await openMaster(context, false);
      var rows = await readRows(context, m);
      var at = -1;
      rows.forEach(function (x) { if (at < 0 && lc(x.material) === lc(r.material)) at = x.row; });
      if (at < 0) throw new Error("\"" + r.material + "\" is no longer in the master list.");
      var errs = validateMaterial(r, [], "edit");
      if (errs.length) throw new Error(errs.join(" "));
      await applyPlan(context, m, { updates: [{ row: at, rec: r }], appends: [] });
      await publish(context, r);
    });
    return r;
  }

  var api = {
    UNITS: UNITS, BLANK: BLANK, TEMPLATE_SRC: TEMPLATE_SRC,
    inspectSheets: inspectSheets, addBomSheet: addBomSheet, addMasterSheet: addMasterSheet,
    masterStatus: masterStatus, applyLatest: applyLatest,
    listMaterials: listMaterials, addMaterial: addMaterial, updateMaterial: updateMaterial,
    checkSheetName: checkSheetName, safeId: safeId, uniqueName: uniqueName,
    normalizeMaterial: normalizeMaterial, validateMaterial: validateMaterial,
    _internal: { syncWorkbook: syncWorkbook, openMaster: openMaster, applyPlan: applyPlan }
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomTools = api;
})(typeof window !== "undefined" ? window : this);
