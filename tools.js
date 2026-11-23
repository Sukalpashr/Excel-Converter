/* tools.js - "Add BOM sheet", "Materials" pages, and workbook repair helpers for BOM Tools.
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

  var HEAD = {
    material: "Material",
    category: "Category",
    unit: "Primary Unit",
    lin: "Linear Weight",
    area: "Area Weight",
    piece: "Piece Weight",
    sl: "Sheet Length",
    sw: "Sheet Width",
    remarks: "Remarks",
    lastAction: "Last Action",
    lastModified: "Last Modified"
  };
  var EDITABLE = ["category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"];

  /* Workbook-level named ranges mapped to their Master table column key */
  var MASTER_NAMES = [
    { name: "Materials", key: "material" },
    { name: "LinWt",     key: "lin" },
    { name: "AreaWt",    key: "area" },
    { name: "PieceWt",   key: "piece" },
    { name: "WtUnit",    key: "unit" },
    { name: "SheetL",    key: "sl" },
    { name: "SheetW",    key: "sw" }
  ];

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

  function formatSheetTimestamp(iso) {
    var d = iso ? new Date(iso) : new Date();
    if (isNaN(d.getTime())) d = new Date();
    var p = function (n) { return (n < 10 ? "0" : "") + n; };
    return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()) + " " +
      p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds());
  }

  function normalizeActionType(t) {
    var s = lc(t || "");
    if (s === "add" || s === "added") return "added";
    if (s === "edit" || s === "edited" || s === "mod" || s === "modified" || s === "updated") return "modified";
    if (s === "remove" || s === "removed" || s === "deleted") return "removed";
    return "";
  }

  function actionDisplay(t) {
    var k = normalizeActionType(t);
    if (k === "added") return "Added";
    if (k === "modified") return "Modified";
    if (k === "removed") return "Removed";
    return "";
  }

  function normalizeMaterial(raw) {
    var r = {
      material: String(raw.material || "").trim(),
      category: String(raw.category || "").trim(),
      unit: String(raw.unit || "").trim(),
      remarks: String(raw.remarks || "").trim(),
      lin: num(raw.lin), area: num(raw.area), piece: num(raw.piece), sl: num(raw.sl), sw: num(raw.sw),
      lastAction: normalizeActionType(raw.lastAction || ""),
      lastModified: raw.lastModified ? String(raw.lastModified).trim() : ""
    };
    // keep only the fields that belong to the chosen unit (the master list never mixes them)
    if (r.unit !== "kg/m") r.lin = null;
    if (r.unit !== "kg/m²") { r.area = null; r.sl = null; r.sw = null; }
    if (r.unit !== "kg/each") r.piece = null;
    return r;
  }

  function getMaterialValidationStatus(raw) {
    var r = normalizeMaterial(raw || {});
    var missing = [];
    if (!r.material) missing.push("Material Name");
    if (UNITS.indexOf(r.unit) < 0) {
      missing.push("Primary Unit");
    } else if (r.unit === "kg/m") {
      if (!pos(r.lin)) missing.push("Linear Weight");
    } else if (r.unit === "kg/each") {
      if (!pos(r.piece)) missing.push("Piece Weight");
    } else if (r.unit === "kg/m²") {
      if (!pos(r.area)) missing.push("Area Weight");
      if (!pos(r.sl) || !pos(r.sw)) missing.push("Sheet Size");
    }
    return { valid: missing.length === 0, missing: missing };
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

  function replaceTableRef(formula, oldName, newName) {
    if (!isFormula(formula) || !oldName || oldName === newName) return formula;
    var esc = oldName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return formula.replace(new RegExp("\\b" + esc + "\\[", "gi"), newName + "[");
  }

  function canonicalM4(bomName) {
    return '=IFERROR(LOOKUP(2,1/(((' + bomName + '[Description]<>"")+(' + bomName + '[Materials]<>"")+(' + bomName + '[Length]<>"")+(' + bomName + '[Qty]<>""))>0),ROW(' + bomName + '[Description])),5)';
  }
  function canonicalAB4(plName) {
    return '=COUNT(' + plName + '[Mat Row])';
  }
  function canonicalO2() {
    return '=SUBSTITUTE($A$1,"Bill of Materials","Remaining Purchase List")&IF($A$2<>""," – "&$A$2,"")';
  }
  function canonicalAD2(bomName, plName) {
    return '=IF(IFERROR(INDEX(' + bomName + '[Materials],MATCH(1,INDEX((' + bomName + '[Materials]<>"")*(COUNTIF(' + plName + '[Materials],' + bomName + '[Materials])=0),0),0)),"")<>"","⚠ Not on purchase list (not in Master, or list full): "&IFERROR(INDEX(' + bomName + '[Materials],MATCH(1,INDEX((' + bomName + '[Materials]<>"")*(COUNTIF(' + plName + '[Materials],' + bomName + '[Materials])=0),0),0)),""),IF(ABS(SUMIF(' + plName + '[Mat Row],">0",' + plName + '[Consumable Wt. kg])-SUMIF(' + bomName + '[Mat Row],">0",' + bomName + '[Total Weight (kg)]))>0.01,"⚠ Purchase list total does not match BOM total – check material names","✔ Purchase list covers every BOM material"))';
  }
  function canonicalPoWeightFormula(rowNum) {
    var prev = rowNum - 1;
    return '=IF(ROW()=$AB$4+6,IF(COUNT(W$5:W' + prev + ')=0,"",SUM(W$5:W' + prev + ')),IF(ROW()=$AB$4+8,"Production",""))';
  }

  /* ---------------- Master table & Named Ranges + Validation ---------------- */
  async function openMaster(context) {
    var ws = context.workbook.worksheets.getItemOrNullObject(MASTER);
    await context.sync();
    if (ws.isNullObject) throw new Error("Sheet \"" + MASTER + "\" not found. Open the automated template or project workbook (not a converted one).");
    var t = ws.tables.getItemOrNullObject("tblMat");
    await context.sync();
    if (t.isNullObject) {
      ws.tables.load("items/name");
      await context.sync();
      if (ws.tables.items.length === 1) {
        t = ws.tables.items[0];
        try { t.name = "tblMat"; await context.sync(); } catch (e) { /* keep existing name */ }
      } else {
        throw new Error("Table \"tblMat\" not found on the Master sheet.");
      }
    }
    t.load("name");
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
    return { ws: ws, t: t, tableName: t.name, headers: headers, idx: idx, ncols: headers.length };
  }

  function rowToObj(r, i, idx) {
    var g = function (k) { return idx[k] >= 0 ? r[idx[k]] : ""; };
    return {
      row: i,
      material: String(g("material")),
      category: String(g("category")),
      unit: String(g("unit")),
      lin: g("lin"),
      area: g("area"),
      piece: g("piece"),
      sl: g("sl"),
      sw: g("sw"),
      remarks: String(g("remarks")),
      lastAction: normalizeActionType(g("lastAction")),
      lastModified: String(g("lastModified") || "").trim()
    };
  }

  /* Ensures "Last Action" and "Last Modified" audit columns exist at the right of tblMat without affecting any named range. */
  async function ensureAuditColumns(context, m) {
    var needAction = m.idx.lastAction < 0;
    var needMod = m.idx.lastModified < 0;
    if (!needAction && !needMod) return m;
    try {
      if (needAction) m.t.columns.add(null, null, HEAD.lastAction);
      if (needMod) m.t.columns.add(null, null, HEAD.lastModified);
      await context.sync();
      var hdr = m.t.getHeaderRowRange().load("values");
      await context.sync();
      m.headers = hdr.values[0].map(function (s) { return String(s).trim(); });
      Object.keys(HEAD).forEach(function (k) {
        m.idx[k] = m.headers.findIndex(function (h) { return h.indexOf(HEAD[k]) === 0; });
      });
      m.ncols = m.headers.length;
      try {
        if (m.idx.lastAction >= 0) m.t.columns.getItem(m.headers[m.idx.lastAction]).getRange().format.columnWidth = 90;
        if (m.idx.lastModified >= 0) m.t.columns.getItem(m.headers[m.idx.lastModified]).getRange().format.columnWidth = 135;
        await context.sync();
      } catch (eW) { /* ignore column width errors */ }
    } catch (e) { /* non-fatal if sheet is protected */ }
    return m;
  }

  function buildBannerSummary(entries, stampText) {
    if (!entries || !entries.length) return "";
    if (entries.length === 1) {
      var e = entries[0];
      return "Last change: " + (actionDisplay(e.type) || "Modified") + " \"" + e.material + "\" · " + stampText;
    }
    var counts = { added: 0, modified: 0, removed: 0 };
    entries.forEach(function (e) {
      var k = normalizeActionType(e.type) || "modified";
      if (counts[k] !== undefined) counts[k]++;
    });
    var parts = [];
    if (counts.added) parts.push(counts.added + " added");
    if (counts.modified) parts.push(counts.modified + " modified");
    if (counts.removed) parts.push(counts.removed + " removed");
    var names = entries.slice(0, 3).map(function (e) { return e.material; }).join(", ");
    if (entries.length > 3) names += " +" + (entries.length - 3) + " more";
    return "Last change: " + parts.join(", ") + " (" + names + ") · " + stampText;
  }

  async function writeMasterBanner(context, ws, entries, timeIso) {
    var stamp = formatSheetTimestamp(timeIso);
    var text = buildBannerSummary(entries, stamp);
    if (!text) return;
    try {
      var cell = ws.getRange("A3");
      cell.values = [[text]];
      cell.format.font.size = 9;
      cell.format.font.italic = true;
      cell.format.font.color = "#4B5563";
      await context.sync();
    } catch (e) { /* non-fatal */ }
  }

  /* Stamps Last Action & Last Modified onto matching rows of tblMat + updates row 3 summary on Master Unique Materials. */
  async function stampMasterChanges(entries, timeIso) {
    if (!entries || !entries.length) return;
    var stamp = formatSheetTimestamp(timeIso);
    await Excel.run(async function (context) {
      var m = await openMaster(context);
      var rowEntries = entries.filter(function (e) { return normalizeActionType(e.type) !== "removed" && e.material; });
      if (rowEntries.length) {
        await ensureAuditColumns(context, m);
        if (m.idx.lastAction >= 0 || m.idx.lastModified >= 0) {
          var body = m.t.getDataBodyRange().load("values");
          await context.sync();
          var rowMap = {};
          body.values.forEach(function (row, i) {
            var n = lc(row[m.idx.material] || "");
            if (n && rowMap[n] === undefined) rowMap[n] = i;
          });
          var touched = false;
          rowEntries.forEach(function (e) {
            var at = rowMap[lc(e.material)];
            if (at === undefined) return;
            var rStamp = e.time ? formatSheetTimestamp(e.time) : stamp;
            if (m.idx.lastAction >= 0) body.getCell(at, m.idx.lastAction).values = [[actionDisplay(e.type) || "Modified"]];
            if (m.idx.lastModified >= 0) body.getCell(at, m.idx.lastModified).values = [[rStamp]];
            touched = true;
          });
          if (touched) await context.sync();
        }
      }
      await writeMasterBanner(context, m.ws, entries, timeIso);
    });
  }

  /* Ensures all 7 workbook-level named ranges point cleanly to tblMat columns (no #REF!, no sheet-level shadows),
   * and ensures the Primary Unit dropdown on the Master sheet is intact.
   * Uses structured table references first, falling back to exact Range objects on Excel builds that reject structured refs in names.add.
   */
  async function ensureMasterNames(context, wb, log) {
    log = log || noop;
    var m = await openMaster(context);
    var nameSet = {};
    MASTER_NAMES.forEach(function (d) { nameSet[d.name.toLowerCase()] = true; });

    // 1. Delete any worksheet-scoped shadow names that insertWorksheetsFromBase64 may have created
    wb.worksheets.load("items/name");
    await context.sync();
    wb.worksheets.items.forEach(function (ws) { ws.names.load("items/name"); });
    await context.sync();
    var removedShadow = false;
    wb.worksheets.items.forEach(function (ws) {
      ws.names.items.forEach(function (nm) {
        if (nameSet[nm.name.toLowerCase()]) { nm.delete(); removedShadow = true; }
      });
    });
    if (removedShadow) await context.sync();

    // 2. Load current column body ranges and addresses from tblMat
    var colRanges = {};
    MASTER_NAMES.forEach(function (d) {
      var colHeader = m.headers[m.idx[d.key]];
      var rg = m.t.columns.getItem(colHeader).getDataBodyRange();
      rg.load(["address", "rowCount"]);
      colRanges[d.name] = { header: colHeader, range: rg };
    });
    wb.names.load("items/name,items/formula");
    await context.sync();

    var existing = {};
    wb.names.items.forEach(function (nm) { existing[nm.name.toLowerCase()] = nm; });

    // 3. Verify or recreate each of the 7 workbook-level named ranges (batched for speed)
    var usedRangeFallback = false;
    var toCreate = [];
    var deletedAny = false;
    for (var i = 0; i < MASTER_NAMES.length; i++) {
      var def = MASTER_NAMES[i];
      var info = colRanges[def.name];
      var structFormula = "=" + m.tableName + "[" + info.header + "]";
      var cur = existing[def.name.toLowerCase()];

      var needsCreate = true;
      if (cur) {
        var curF = String(cur.formula || "");
        var isBroken = !curF || curF.indexOf("#REF!") >= 0 || curF.indexOf("#NAME?") >= 0 || /\[\d+\]/.test(curF);
        var matchesStruct = lc(curF) === lc(structFormula);
        var matchesAddr = !isBroken && curF.indexOf("!") >= 0 && curF.replace(/\$/g, "").toLowerCase().indexOf(info.range.address.replace(/\$/g, "").toLowerCase()) >= 0;
        if (!isBroken && (matchesStruct || matchesAddr)) {
          needsCreate = false;
        } else {
          try { cur.delete(); deletedAny = true; } catch (e) { /* ignore */ }
        }
      }
      if (needsCreate) toCreate.push({ def: def, info: info, structFormula: structFormula });
    }
    if (deletedAny) {
      try { await context.sync(); } catch (e) { /* ignore */ }
    }

    if (toCreate.length) {
      var batchOk = false;
      try {
        toCreate.forEach(function (item) { wb.names.add(item.def.name, item.structFormula); });
        await context.sync();
        batchOk = true;
      } catch (eBatch) {
        batchOk = false;
      }

      if (batchOk) {
        // Verify none of the batch-added names became #REF! in a single sync
        try {
          var checks = toCreate.map(function (item) {
            return { item: item, nm: wb.names.getItem(item.def.name).load("formula") };
          });
          await context.sync();
          var bad = checks.filter(function (c) { return String(c.nm.formula || "").indexOf("#REF!") >= 0; });
          if (bad.length) {
            bad.forEach(function (c) { c.nm.delete(); });
            await context.sync();
            bad.forEach(function (c) { wb.names.add(c.item.def.name, c.item.info.range); });
            await context.sync();
            usedRangeFallback = true;
          }
        } catch (eChk) { /* ignore */ }
      } else {
        // Per-name fallback for Excel builds that reject structured refs in wb.names.add()
        for (var j = 0; j < toCreate.length; j++) {
          var it = toCreate[j];
          try {
            wb.names.add(it.def.name, it.info.range);
            await context.sync();
            usedRangeFallback = true;
          } catch (e2) {
            log("WARNING: could not create named range " + it.def.name + " (" + (e2.message || e2) + ").");
          }
        }
      }
    }

    // 4. Ensure Primary Unit dropdown validation on Master Unique Materials (C5:C300+)
    var matRows = colRanges.Materials.range.rowCount || 55;
    var dvEnd = Math.max(300, 4 + matRows + 100);
    try {
      var unitDv = m.ws.getRange("C5:C" + dvEnd).dataValidation;
      unitDv.clear();
      unitDv.rule = {
        list: {
          inCellDropDown: true,
          source: "kg/m,kg/m²,kg/each"
        }
      };
      unitDv.ignoreBlanks = true;
      unitDv.prompt = {
        showPrompt: true,
        title: "Primary Unit",
        message: "kg/m = by length, kg/m² = by area, kg/each = per piece"
      };
      unitDv.errorAlert = {
        showAlert: true,
        style: Excel.DataValidationAlertStyle.stop,
        title: "Pick a unit",
        message: "Choose kg/m, kg/m² or kg/each."
      };
      await context.sync();
    } catch (e) { /* non-fatal on older hosts */ }

    return {
      master: m,
      matRangeAddress: colRanges.Materials.range.address,
      usedRangeFallback: usedRangeFallback
    };
  }

  /* Explicitly sets the Materials dropdown on C6:C2000 of a BOM sheet so it works reliably after template import & sheet copy. */
  async function applyBomDataValidation(context, ws, matInfo, log) {
    log = log || noop;
    var target = ws.getRange("C6:C2000");
    var alertCfg = {
      showAlert: true,
      style: Excel.DataValidationAlertStyle.stop,
      title: "Not in the materials list",
      message: "Pick a material from the dropdown (or add it to the Master Unique Materials sheet first)."
    };

    try {
      target.dataValidation.clear();
      target.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "=Materials"
        }
      };
      target.dataValidation.ignoreBlanks = true;
      target.dataValidation.prompt = { showPrompt: false, title: "", message: "" };
      target.dataValidation.errorAlert = alertCfg;
      await context.sync();
      return;
    } catch (e1) {
      // Fallback: bind directly to the Master Unique Materials A1 range address
      try {
        target.dataValidation.clear();
        target.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "=" + matInfo.matRangeAddress
          }
        };
        target.dataValidation.ignoreBlanks = true;
        target.dataValidation.prompt = { showPrompt: false, title: "", message: "" };
        target.dataValidation.errorAlert = alertCfg;
        await context.sync();
      } catch (e2) {
        log("WARNING: could not set Materials dropdown on " + ws.name + " (" + (e2.message || e2) + ").");
      }
    }
  }

  /* Restores/verifies layout, hidden helper rows/cols, header formulas, print setup, and dropdown on a BOM sheet. */
  async function repairBomSheet(context, wb, ws, matInfo, opts, log) {
    opts = opts || {};
    log = log || noop;

    ws.tables.load("items/name");
    await context.sync();
    var bomT = firstTable(ws.tables.items, "bom_");
    var plT = firstTable(ws.tables.items, "pl_");
    if (!bomT) return;

    var bomName = bomT.name;
    var plName = plT ? plT.name : "";

    // 1. Load top-level formulas + optional stale table bodies in ONE sync
    var rO2 = ws.getRange("O2").load("formulas");
    var rM4 = ws.getRange("M4").load("formulas");
    var rAB4 = ws.getRange("AB4").load("formulas");
    var rAD2 = ws.getRange("AD2").load("formulas");
    var rAE1 = ws.getRange("AE1").load("values");
    var bomBody = (opts.srcBomName && opts.srcBomName !== bomName) ? bomT.getDataBodyRange().load("formulas") : null;
    var plBody = (plT && ((opts.srcBomName && opts.srcBomName !== bomName) || (opts.srcPlName && opts.srcPlName !== plName)))
      ? plT.getDataBodyRange().load("formulas")
      : null;
    await context.sync();

    var fO2 = String(rO2.formulas[0][0] || "");
    if (!isFormula(fO2) || fO2.indexOf("#REF!") >= 0) rO2.formulas = [[canonicalO2()]];

    var fM4 = String(rM4.formulas[0][0] || "");
    if (!isFormula(fM4) || fM4.indexOf(bomName + "[") < 0 || fM4.indexOf("#REF!") >= 0) {
      rM4.formulas = [[canonicalM4(bomName)]];
    }

    if (plName) {
      var fAB4 = String(rAB4.formulas[0][0] || "");
      if (!isFormula(fAB4) || fAB4.indexOf(plName + "[") < 0 || fAB4.indexOf("#REF!") >= 0) {
        rAB4.formulas = [[canonicalAB4(plName)]];
      }
      var fAD2 = String(rAD2.formulas[0][0] || "");
      if (!isFormula(fAD2) || fAD2.indexOf(bomName + "[") < 0 || fAD2.indexOf(plName + "[") < 0 || fAD2.indexOf("#REF!") >= 0) {
        rAD2.formulas = [[canonicalAD2(bomName, plName)]];
      }
    }

    if (isBlank(rAE1.values[0][0])) {
      rAE1.values = [[6000]];
    }

    if (bomBody) {
      var bomChanged = false;
      var bomOut = bomBody.formulas.map(function (row) {
        return row.map(function (cell) {
          var fixed = replaceTableRef(cell, opts.srcBomName, bomName);
          if (fixed !== cell) bomChanged = true;
          return fixed;
        });
      });
      if (bomChanged) bomBody.formulas = bomOut;
    }
    if (plBody) {
      var plChanged = false;
      var plOut = plBody.formulas.map(function (row) {
        return row.map(function (cell) {
          var fixed = cell;
          if (opts.srcBomName && opts.srcBomName !== bomName) fixed = replaceTableRef(fixed, opts.srcBomName, bomName);
          if (opts.srcPlName && opts.srcPlName !== plName) fixed = replaceTableRef(fixed, opts.srcPlName, plName);
          if (fixed !== cell) plChanged = true;
          return fixed;
        });
      });
      if (plChanged) plBody.formulas = plOut;
    }

    // 2. Enforce hidden table header row (row 5), hidden helper columns (M & AB), gridlines off, freeze panes in same sync
    try {
      ws.getRange("5:5").rowHidden = true;
      ws.getRange("M:M").columnHidden = true;
      ws.getRange("AB:AB").columnHidden = true;
      ws.showGridlines = false;
      ws.freezePanes.unfreeze();
      ws.freezePanes.freezeRows(5);
    } catch (e) { /* non-fatal */ }
    await context.sync();

    // 3. Re-apply Materials dropdown on C6:C2000
    await applyBomDataValidation(context, ws, matInfo, log);

    // 4. Enforce landscape print setup & repeating header rows 3:4
    try {
      ws.pageLayout.orientation = Excel.PageOrientation.landscape;
      ws.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0 };
      ws.pageLayout.setPrintTitleRows("3:4");
      await context.sync();
    } catch (e) { /* non-fatal on hosts without full pageLayout support */ }
  }

  /* Repairs named ranges, dropdowns, formatting, and formulas across selected (or all) BOM sheets. */
  async function repairWorkbook(log, opts) {
    log = log || noop;
    opts = opts || {};
    var selectedSet = null;
    if (Array.isArray(opts.sheets)) {
      if (!opts.sheets.length) throw new Error("Select at least one BOM sheet to repair.");
      selectedSet = {};
      opts.sheets.forEach(function (n) { selectedSet[lc(n)] = true; });
    }
    var count = 0;
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var prevCalcMode = null;
      try {
        wb.application.load("calculationMode");
        await context.sync();
        prevCalcMode = wb.application.calculationMode;
        if (prevCalcMode !== Excel.CalculationMode.manual) {
          wb.application.calculationMode = Excel.CalculationMode.manual;
          await context.sync();
        }
      } catch (e) { prevCalcMode = null; }

      try {
        log("Checking Master sheet and workbook named ranges ...");
        var matInfo = await ensureMasterNames(context, wb, log);
        wb.worksheets.load("items/name,items/visibility");
        await context.sync();
        var sheets = wb.worksheets.items;
        sheets.forEach(function (ws) { ws.tables.load("items/name"); });
        await context.sync();

        for (var i = 0; i < sheets.length; i++) {
          var ws = sheets[i];
          if (!firstTable(ws.tables.items, "bom_")) continue;
          if (selectedSet && !selectedSet[lc(ws.name)]) continue;
          log("Verifying dropdowns, layout and formulas on \"" + ws.name + "\" ...");
          await repairBomSheet(context, wb, ws, matInfo, {}, log);
          count++;
        }
      } finally {
        try {
          if (prevCalcMode !== null) wb.application.calculationMode = prevCalcMode;
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        } catch (e) { /* ignore */ }
      }
    });
    log("Repaired " + count + " BOM sheet(s) and Master named ranges.");
    return { sheetsRepaired: count };
  }

  /* ================= Add BOM sheet ================= */
  async function inspectSheets() {
    var out = { bomSheets: [], allSheets: [], hasBlank: false, hasMaster: false, needsTemplate: false, project: "" };
    await Excel.run(async function (context) {
      var wb = context.workbook;
      wb.worksheets.load("items/name,items/visibility,items/position");
      await context.sync();
      var sheets = wb.worksheets.items;
      sheets.forEach(function (ws) { ws.tables.load("items/name"); });
      await context.sync();
      out.allSheets = sheets.map(function (ws) { return ws.name; });
      out.hasMaster = out.allSheets.some(function (n) { return lc(n) === lc(MASTER); });
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
    out.needsTemplate = !out.hasMaster && out.bomSheets.length === 0;
    return out;
  }

  async function clearInputs(context, table, names, isPlTable) {
    table.columns.load("items/name");
    await context.sync();
    var work = [];
    table.columns.items.forEach(function (c) {
      if (names.indexOf(c.name) >= 0) {
        var r = c.getDataBodyRange();
        r.load(["formulas", "rowIndex"]);
        work.push({ colName: c.name, range: r });
      }
    });
    await context.sync();
    work.forEach(function (item) {
      var r = item.range;
      var changed = false;
      var startRow1 = (r.rowIndex || 5) + 1; // 1-based Excel row index of first body row (row 6)
      var out = r.formulas.map(function (row, idx) {
        var v = row[0];
        if (isPlTable && item.colName === "PO Weight") {
          // Restore the Grand Total / Production footer formula that users overwrite when typing PO weights
          var f = canonicalPoWeightFormula(startRow1 + idx);
          if (v !== f) { changed = true; return [f]; }
          return [v];
        }
        if (!isFormula(v) && v !== "") { changed = true; return [""]; }
        return [v];
      });
      if (changed) r.formulas = out;
    });
    await context.sync();
  }

  async function addBomSheet(opts, log) {
    log = log || noop;
    var result = { name: "", warnings: [] };
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var matInfo = await ensureMasterNames(context, wb, log);

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
      var srcPl = firstTable(srcSheet.tables.items, "pl_");
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

      // tidy table names to the template convention (bom_<sheet>, pl_<sheet>)
      wb.tables.load("items/name");
      await context.sync();
      var taken = wb.tables.items.map(function (t) { return t.name; });
      var id = safeId(newName);
      var bomName = uniqueName("bom_" + id, taken);
      var plName = uniqueName("pl_" + id, taken.concat([bomName]));
      log("Naming tables " + bomName + (plT ? " / " + plName : "") + " ...");
      bomT.name = bomName;
      if (plT) plT.name = plName;
      await context.sync();

      copy.getRange("A2").values = [[newName]];
      if (opts.project && String(opts.project).trim()) {
        copy.getRange("A1").values = [[TITLE_PREFIX + String(opts.project).trim()]];
      }
      await context.sync();

      if (opts.clear) {
        log("Clearing entered data ...");
        await clearInputs(context, bomT, BOM_INPUT, false);
        if (plT) await clearInputs(context, plT, PL_INPUT, true);
      }

      // Ensure dropdowns, table formulas, hidden helper row/cols, freeze panes, and print setup are all working
      await repairBomSheet(context, wb, copy, matInfo, {
        srcBomName: srcBom.name,
        srcPlName: srcPl ? srcPl.name : ""
      }, log);

      wb.application.calculate(Excel.CalculationType.full);
      await context.sync();

      var f3 = copy.getRange("A6").load("formulas");
      var checkM4 = copy.getRange("M4").load("values");
      await context.sync();
      if (String(f3.formulas[0][0]).indexOf(bomName + "[") < 0 && String(f3.formulas[0][0]).indexOf("[@") < 0) {
        result.warnings.push("Could not confirm the S.N. formula follows the renamed table. Check column A on the new sheet.");
      }
      if (typeof checkM4.values[0][0] === "string" && checkM4.values[0][0].charAt(0) === "#") {
        result.warnings.push("Formula check cell M4 returned " + checkM4.values[0][0] + ". Try clicking Repair dropdowns & formulas.");
      }

      copy.activate();
      copy.getRange("B6").select();
      await context.sync();
      result.name = newName;
    });
    return result;
  }

  /* ================= Materials master ================= */
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

  async function refreshAllBomDropdownsIfRangeBased(context, wb, matInfo) {
    if (!matInfo || !matInfo.usedRangeFallback) return;
    wb.worksheets.load("items/name");
    await context.sync();
    wb.worksheets.items.forEach(function (ws) { ws.tables.load("items/name"); });
    await context.sync();
    for (var i = 0; i < wb.worksheets.items.length; i++) {
      var ws = wb.worksheets.items[i];
      if (firstTable(ws.tables.items, "bom_")) {
        await applyBomDataValidation(context, ws, matInfo, noop);
      }
    }
  }

  async function addMaterial(raw) {
    var r = normalizeMaterial(raw);
    var nowIso = new Date().toISOString();
    var stamp = formatSheetTimestamp(nowIso);
    r.lastAction = "added";
    r.lastModified = nowIso;
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var m = await openMaster(context);
      await ensureAuditColumns(context, m);
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
      var vals = {
        material: r.material, category: r.category, unit: r.unit,
        lin: r.lin, area: r.area, piece: r.piece, sl: r.sl, sw: r.sw,
        remarks: r.remarks, lastAction: "Added", lastModified: stamp
      };
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

      await writeMasterBanner(context, m.ws, [{ material: r.material, type: "added", time: nowIso }], nowIso);

      // Keep named ranges & dropdowns synced with the newly added row
      var matInfo = await ensureMasterNames(context, wb, noop);
      await refreshAllBomDropdownsIfRangeBased(context, wb, matInfo);
      wb.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
    return r;
  }

  async function updateMaterial(raw) {
    var r = normalizeMaterial(raw);
    var nowIso = new Date().toISOString();
    var stamp = formatSheetTimestamp(nowIso);
    r.lastAction = "modified";
    r.lastModified = nowIso;
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var m = await openMaster(context);
      await ensureAuditColumns(context, m);
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
      if (m.idx.lastAction >= 0) row.getCell(0, m.idx.lastAction).values = [["Modified"]];
      if (m.idx.lastModified >= 0) row.getCell(0, m.idx.lastModified).values = [[stamp]];
      await context.sync();
      await writeMasterBanner(context, m.ws, [{ material: r.material, type: "modified", time: nowIso }], nowIso);
      wb.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
    return r;
  }

  async function deleteMaterial(materialName) {
    var targetName = String(materialName || "").trim();
    if (!targetName) throw new Error("Choose a material to remove.");
    var nowIso = new Date().toISOString();
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var m = await openMaster(context);
      var body = m.t.getDataBodyRange().load(["values", "rowCount"]);
      await context.sync();
      var at = -1;
      body.values.forEach(function (row, i) {
        if (at < 0 && lc(row[m.idx.material]) === lc(targetName)) at = i;
      });
      if (at < 0) throw new Error("\"" + targetName + "\" was not found in the master list.");

      if (body.rowCount <= 1) {
        // Keep table valid by clearing the single remaining row instead of deleting the only row
        var row = body.getRow(0);
        Object.keys(HEAD).forEach(function (k) {
          if (m.idx[k] >= 0) row.getCell(0, m.idx[k]).values = [[""]];
        });
      } else {
        m.t.rows.getItemAt(at).delete();
      }
      await context.sync();

      await writeMasterBanner(context, m.ws, [{ material: targetName, type: "removed", time: nowIso }], nowIso);

      var matInfo = await ensureMasterNames(context, wb, noop);
      await refreshAllBomDropdownsIfRangeBased(context, wb, matInfo);
      wb.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
    return targetName;
  }

  var api = {
    UNITS: UNITS, BLANK: BLANK, MASTER: MASTER,
    _openMaster: openMaster, _rowToObj: rowToObj,
    _ensureAuditColumns: ensureAuditColumns,
    _writeMasterBanner: writeMasterBanner,
    _ensureMasterNames: ensureMasterNames,
    _applyBomDataValidation: applyBomDataValidation,
    _repairBomSheet: repairBomSheet,
    _refreshAllBomDropdownsIfRangeBased: refreshAllBomDropdownsIfRangeBased,
    _replaceTableRef: replaceTableRef,
    _canonical: { O2: canonicalO2, M4: canonicalM4, AB4: canonicalAB4, AD2: canonicalAD2 },
    inspectSheets: inspectSheets, addBomSheet: addBomSheet, repairWorkbook: repairWorkbook,
    listMaterials: listMaterials, addMaterial: addMaterial, updateMaterial: updateMaterial,
    deleteMaterial: deleteMaterial, stampMasterChanges: stampMasterChanges,
    // pure helpers
    checkSheetName: checkSheetName, safeId: safeId, uniqueName: uniqueName,
    normalizeMaterial: normalizeMaterial, validateMaterial: validateMaterial,
    getMaterialValidationStatus: getMaterialValidationStatus,
    formatSheetTimestamp: formatSheetTimestamp, actionDisplay: actionDisplay
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomTools = api;
})(typeof window !== "undefined" ? window : this);
