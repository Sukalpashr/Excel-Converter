/* sheets.js - Lightweight Excel Workbook, Template Import, Lag-Free Sheet Creation, Materials & Repair Engine (3.0v patch2) */
(() => {
  "use strict";

  const MASTER = "Master Unique Materials";
  const BLANK = "BOM Blank";
  const PROP_PROJECT = "BomProjectList";
  const TITLE_PREFIX = "Bill of Materials for ";
  const UNITS = ["kg/m", "kg/m²", "kg/each"];
  const BOM_INPUT_COLS = ["Description", "Materials", "Length", "Width", "Thickness", "Qty", "Set", "Remarks"];
  const PL_INPUT_COLS = ["Stock", "PO Weight", "PO Qty", "PO Rate", "Remarks"];

  const MASTER_HEADERS = {
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

  const NAMED_RANGES = [
    { name: "Materials", key: "material" },
    { name: "LinWt", key: "lin" },
    { name: "AreaWt", key: "area" },
    { name: "PieceWt", key: "piece" },
    { name: "WtUnit", key: "unit" },
    { name: "SheetL", key: "sl" },
    { name: "SheetW", key: "sw" }
  ];

  const lc = (s) => String(s ?? "").trim().toLowerCase();
  const isFormula = (v) => typeof v === "string" && v.startsWith("=");
  const isBlank = (v) => v === "" || v === null || v === undefined;
  const noop = () => {};

  /* Centralized Excel.run with delayForCellEdit so background add-in activity NEVER interrupts active cell typing */
  function runExcel(fn) {
    return Excel.run({ delayForCellEdit: true }, fn);
  }

  function pauseExcel(context) {
    try { context.application.suspendApiCalculationUntilNextSync(); } catch { /* ignore */ }
    try {
      if (Office?.context?.requirements?.isSetSupported("ExcelApi", "1.10") &&
          typeof context.application.suspendScreenUpdatingUntilNextSync === "function") {
        context.application.suspendScreenUpdatingUntilNextSync();
      }
    } catch { /* ignore */ }
  }

  function safeId(s) {
    return String(s || "").trim().replace(/[^A-Za-z0-9_]+/g, "_").replace(/^_+|_+$/g, "") || "Sheet";
  }

  function uniqueName(base, taken) {
    const used = new Set(taken.map(lc));
    let name = base;
    let i = 2;
    while (used.has(lc(name))) name = `${base}_${i++}`;
    return name;
  }

  function checkSheetName(name, existing) {
    const n = String(name || "").trim();
    if (!n) return "Enter a sheet name to continue.";
    if (n.length > 31) return "Sheet names can be at most 31 characters.";
    if (/[:\\/?*\[\]]/.test(n)) return "Sheet names cannot contain : \\ / ? * [ ]";
    if (/^'|'$/.test(n)) return "Sheet name cannot start or end with an apostrophe.";
    if (lc(n) === "history") return '"History" is reserved by Excel.';
    if (existing.some((e) => lc(e) === lc(n))) return `"${n}" already exists in this workbook.`;
    return "";
  }

  function formatStamp(iso) {
    const d = iso ? new Date(iso) : new Date();
    const t = Number.isNaN(d.getTime()) ? new Date() : d;
    const p = (n) => (n < 10 ? "0" : "") + n;
    return `${t.getFullYear()}-${p(t.getMonth() + 1)}-${p(t.getDate())} ${p(t.getHours())}:${p(t.getMinutes())}:${p(t.getSeconds())}`;
  }

  function actionLabel(t) {
    const s = lc(t);
    if (s === "added" || s === "add") return "Added";
    if (s === "modified" || s === "edited") return "Modified";
    if (s === "removed" || s === "deleted") return "Removed";
    return "";
  }

  function materialStatus(raw) {
    const m = window.BomStore.normalizeMaterial(raw || {});
    const missing = [];
    if (!m.material) missing.push("Material Name");
    if (!UNITS.includes(m.unit)) {
      missing.push("Primary Unit");
    } else if (m.unit === "kg/m") {
      if (!(m.lin > 0)) missing.push("Linear Weight");
    } else if (m.unit === "kg/each") {
      if (!(m.piece > 0)) missing.push("Piece Weight");
    } else if (m.unit === "kg/m²") {
      if (!(m.area > 0)) missing.push("Area Weight");
      if (!(m.sl > 0) || !(m.sw > 0)) missing.push("Sheet Size");
    }
    return { valid: missing.length === 0, missing };
  }

  function validateMaterial(raw, existingNames, mode) {
    const m = window.BomStore.normalizeMaterial(raw || {});
    const issues = [];
    if (!m.material) issues.push("Material name is required.");
    else if (m.material.startsWith("=")) issues.push("Material name cannot start with '='.");
    else if (mode === "add" && existingNames.some((n) => lc(n) === lc(m.material))) {
      issues.push(`"${m.material}" already exists. Click it above to edit.`);
    }
    if (!UNITS.includes(m.unit)) issues.push("Choose a Primary Unit.");
    else if (m.unit === "kg/m" && !(m.lin > 0)) issues.push("Linear weight (kg/m) must be greater than 0.");
    else if (m.unit === "kg/each" && !(m.piece > 0)) issues.push("Piece weight (kg/each) must be greater than 0.");
    else if (m.unit === "kg/m²") {
      if (!(m.area > 0)) issues.push("Area weight (kg/m²) must be greater than 0.");
      if (!(m.sl > 0) || !(m.sw > 0)) issues.push("Sheet length and width (mm) must be greater than 0.");
    }
    return issues;
  }

  function firstTable(items, prefix) {
    return items.find((t) => t.name.startsWith(prefix)) || null;
  }

  function replaceTableRef(formula, oldName, newName) {
    if (!isFormula(formula) || !oldName || oldName === newName) return formula;
    const esc = oldName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    return formula.replace(new RegExp(`\\b${esc}\\[`, "gi"), `${newName}[`);
  }

  const canonical = {
    O2: () => '=SUBSTITUTE($A$1,"Bill of Materials","Remaining Purchase List")&IF($A$2<>""," – "&$A$2,"")',
    M4: (bom) => `=IFNA(XMATCH(TRUE,(${bom}[Description]<>"")+(${bom}[Materials]<>"")+(${bom}[Length]<>"")+(${bom}[Qty]<>"")>0,0,-1)+5,5)`,
    AL6: (bom) => `=LET(m,FILTER(${bom}[Mat Row],ISNUMBER(${bom}[Mat Row]),""),IF(INDEX(m,1)="","",SORT(UNIQUE(m))))`,
    AB4: () => '=IF(AL6="",0,ROWS($AL$6#))',
    AD2: (bom, pl) =>
      `=IF($M$4=5,"✔ Purchase list covers every BOM material",IF(COUNTIF(${bom}[Materials],"<>")>COUNT(${bom}[Mat Row]),"⚠ Material not in Master – check red cells in column C",IF(ABS(SUMIF(${pl}[Mat Row],">0",${pl}[Consumable Wt. kg])-SUMIF(${bom}[Mat Row],">0",${bom}[Total Weight (kg)]))>0.01,"⚠ Purchase list total does not match BOM total – check material names or capacity","✔ Purchase list covers every BOM material")))`,
    bomSn: (bom, r) => `=IF(${bom}[[#This Row],[Materials]]="","",MAX(A$5:A${r - 1})+1)`,
    bomI: (bom) => `=IF(ROW()=$M$4+1,"Total:",IF(${bom}[[#This Row],[Mat Row]]="","",IF(INDEX(WtUnit,${bom}[[#This Row],[Mat Row]])="kg/m",INDEX(LinWt,${bom}[[#This Row],[Mat Row]]),"")))`,
    bomJ: (bom, r) => `=IF(ROW()=$M$4+1,SUM(J$5:J${r - 1}),LET(r,${bom}[[#This Row],[Mat Row]],IF(r="","",LET(u,INDEX(WtUnit,r),q,${bom}[[#This Row],[Qty]]*IF(${bom}[[#This Row],[Set]]="",1,${bom}[[#This Row],[Set]]),IF(u="kg/m",q*${bom}[[#This Row],[Length]]/1000*INDEX(LinWt,r),IF(u="kg/m²",q*${bom}[[#This Row],[Length]]*${bom}[[#This Row],[Width]]/10^6*INDEX(AreaWt,r),q*INDEX(PieceWt,r)))))))`,
    bomL: (bom) => `=IF(OR(${bom}[[#This Row],[Mat Row]]="",${bom}[[#This Row],[Unit weight kg/m]]="",NOT(ISNUMBER(${bom}[[#This Row],[Length]]))),"",${bom}[[#This Row],[Length]]*${bom}[[#This Row],[Qty]]*IF(${bom}[[#This Row],[Set]]="",1,${bom}[[#This Row],[Set]]))`,
    bomM: (bom) => `=IF(${bom}[[#This Row],[Materials]]="","",IFNA(XMATCH(${bom}[[#This Row],[Materials]],Materials,0),""))`,
    plO: (pl, r) => `=IF(${pl}[[#This Row],[Mat Row]]="","",N(O${r - 1})+1)`,
    plP: (pl) => `=IF(ROW()=$AB$4+6,"Grand Total Weight in KG :",IF(ROW()=$AB$4+8,"Requested by:",IF(${pl}[[#This Row],[Mat Row]]="","",INDEX(Materials,${pl}[[#This Row],[Mat Row]]))))`,
    plQ: (pl) => `=IF(ROW()=$AB$4+8,"Procurement",LET(r,${pl}[[#This Row],[Mat Row]],IF(r="","",LET(u,INDEX(WtUnit,r),IF(u="kg/m",$AE$1,IF(u="kg/m²",LET(l,INDEX(SheetL,r),w,INDEX(SheetW,r),IF(AND(ISNUMBER(l),ISNUMBER(w),l*w>0),l&" x "&w,"Add size in Master")),""))))))`,
    plR: (bom, pl) => `=IF(${pl}[[#This Row],[Mat Row]]="","",LET(r,${pl}[[#This Row],[Mat Row]],u,INDEX(WtUnit,r),IF(u="kg/m",IFERROR(ROUNDUP(ROUND(SUMIF(${bom}[Materials],${pl}[[#This Row],[Materials]],${bom}[Total Length mm])/${pl}[[#This Row],[Length (mm)]],6),0),""),IF(u="kg/m²",IFERROR(ROUNDUP(ROUND(${pl}[[#This Row],[Consumable Wt. kg]]/(INDEX(AreaWt,r)*INDEX(SheetL,r)*INDEX(SheetW,r)/10^6),6),0),""),IFERROR(ROUNDUP(ROUND(${pl}[[#This Row],[Consumable Wt. kg]]/INDEX(PieceWt,r),6),0),"")))))`,
    plS: (bom, pl, r) => `=IF(ROW()=$AB$4+6,SUM(S$5:S${r - 1}),IF(${pl}[[#This Row],[Mat Row]]="","",SUMIF(${bom}[Materials],${pl}[[#This Row],[Materials]],${bom}[Total Weight (kg)])))`,
    plT: (pl) => `=IF(ROW()=$AB$4+8,"Store Section:",LET(r,${pl}[[#This Row],[Mat Row]],IF(r="","",IF(INDEX(WtUnit,r)="kg/m",INDEX(LinWt,r),""))))`,
    plU: (pl, r) => `=IF(ROW()=$AB$4+6,SUM(U$5:U${r - 1}),LET(r,${pl}[[#This Row],[Mat Row]],IF(r="","",IFERROR(LET(u,INDEX(WtUnit,r),q,${pl}[[#This Row],[Qty (nos)]],IF(u="kg/m",${pl}[[#This Row],[Length (mm)]]/1000*q*${pl}[[#This Row],[Weight Kg/m]],IF(u="kg/m²",q*INDEX(AreaWt,r)*INDEX(SheetL,r)*INDEX(SheetW,r)/10^6,q*INDEX(PieceWt,r)))),${pl}[[#This Row],[Consumable Wt. kg]]))))`,
    plZ: (pl, r) => `=IF(ROW()=$AB$4+6,IF(COUNT(Z$5:Z${r - 1})=0,"",SUM(Z$5:Z${r - 1})),IF(ROW()=$AB$4+8,"Engineering",IF(${pl}[[#This Row],[PO Rate]]="","",${pl}[[#This Row],[PO Weight]]*${pl}[[#This Row],[PO Rate]])))`,
    plAB: () => '=IF(ROW()-5>$AB$4,"",INDEX($AL$6#,ROW()-5))',
    poWeight: (r) => `=IF(ROW()=$AB$4+6,IF(COUNT(W$5:W${r - 1})=0,"",SUM(W$5:W${r - 1})),IF(ROW()=$AB$4+8,"Production",""))`
  };

  /* ---------------- Project-Only Workbook Custom Property ---------------- */
  async function readProjectProp(context, wb) {
    try {
      const prop = wb.properties.custom.getItemOrNullObject(PROP_PROJECT);
      prop.load("value");
      await context.sync();
      if (prop.isNullObject || !prop.value) return null;
      const parsed = JSON.parse(String(prop.value));
      const projectName = String(parsed?.projectName || parsed?.name || "").trim();
      if (projectName) {
        return {
          id: String(parsed?.id || "").trim(),
          projectName,
          updated: String(parsed?.updated || parsed?.created || "")
        };
      }
    } catch { /* ignore */ }
    return null;
  }

  async function writeProjectProp(context, wb, meta) {
    try {
      const existing = wb.properties.custom.getItemOrNullObject(PROP_PROJECT);
      await context.sync();
      if (!existing.isNullObject) {
        existing.delete();
        await context.sync();
      }
      const name = String(meta?.projectName || meta?.name || "").trim();
      if (name) {
        wb.properties.custom.add(PROP_PROJECT, JSON.stringify({
          id: meta.id || "",
          projectName: name,
          name,
          updated: meta.updated || new Date().toISOString()
        }));
        await context.sync();
      }
    } catch { /* non-fatal */ }
  }

  async function getProjectMeta() {
    let out = null;
    await runExcel(async (context) => {
      out = await readProjectProp(context, context.workbook);
    });
    return out;
  }

  async function setProjectMeta(meta) {
    await runExcel(async (context) => {
      await writeProjectProp(context, context.workbook, meta);
    });
  }

  /* ---------------- Master Sheet Access & Named Ranges ---------------- */
  async function openMaster(context) {
    const ws = context.workbook.worksheets.getItemOrNullObject(MASTER);
    await context.sync();
    if (ws.isNullObject) throw new Error(`Sheet "${MASTER}" was not found.`);
    let t = ws.tables.getItemOrNullObject("tblMat");
    await context.sync();
    if (t.isNullObject) {
      ws.tables.load("items/name");
      await context.sync();
      if (ws.tables.items.length === 1) {
        t = ws.tables.items[0];
        try { t.name = "tblMat"; await context.sync(); } catch { /* keep */ }
      } else {
        throw new Error('Table "tblMat" was not found on Master Unique Materials.');
      }
    }
    t.load("name");
    const hdr = t.getHeaderRowRange().load("values");
    await context.sync();
    const headers = hdr.values[0].map((s) => String(s).trim());
    const idx = {};
    for (const [k, prefix] of Object.entries(MASTER_HEADERS)) {
      idx[k] = headers.findIndex((h) => h.startsWith(prefix));
    }
    for (const req of ["material", "unit", "lin", "area", "piece", "sl", "sw"]) {
      if (idx[req] < 0) throw new Error(`Master table is missing column "${MASTER_HEADERS[req]}".`);
    }
    return { ws, t, tableName: t.name, headers, idx, ncols: headers.length };
  }

  function rowToMaterial(row, rowIdx, idx) {
    const get = (k) => (idx[k] >= 0 ? row[idx[k]] : "");
    return {
      row: rowIdx,
      ...window.BomStore.normalizeMaterial({
        material: get("material"),
        category: get("category"),
        unit: get("unit"),
        lin: get("lin"),
        area: get("area"),
        piece: get("piece"),
        sl: get("sl"),
        sw: get("sw"),
        remarks: get("remarks"),
        lastAction: get("lastAction"),
        lastModified: get("lastModified")
      })
    };
  }

  async function ensureAuditColumns(context, m) {
    if (m.idx.lastAction >= 0 && m.idx.lastModified >= 0) return m;
    try {
      if (m.idx.lastAction < 0) m.t.columns.add(null, null, MASTER_HEADERS.lastAction);
      if (m.idx.lastModified < 0) m.t.columns.add(null, null, MASTER_HEADERS.lastModified);
      await context.sync();
      const hdr = m.t.getHeaderRowRange().load("values");
      await context.sync();
      m.headers = hdr.values[0].map((s) => String(s).trim());
      for (const [k, prefix] of Object.entries(MASTER_HEADERS)) {
        m.idx[k] = m.headers.findIndex((h) => h.startsWith(prefix));
      }
      m.ncols = m.headers.length;
    } catch { /* ignore if protected */ }
    return m;
  }

  async function writeMasterBanner(context, ws, text, isProjectOnly) {
    try {
      const cell = ws.getRange("A3");
      if (!text) {
        if (!isProjectOnly) {
          cell.values = [[""]];
          cell.format.font.bold = false;
          cell.format.font.italic = false;
          cell.format.font.color = "#475569";
          cell.format.fill.clear();
          await context.sync();
        }
        return;
      }
      cell.values = [[text]];
      cell.format.font.size = 9.5;
      cell.format.font.bold = Boolean(isProjectOnly);
      cell.format.font.italic = !isProjectOnly;
      cell.format.font.color = isProjectOnly ? "#B45309" : "#475569";
      if (isProjectOnly) cell.format.fill.color = "#FEF3C7";
      else cell.format.fill.clear();
      await context.sync();
    } catch { /* non-fatal */ }
  }

  function formatBannerSummary(entries, timeIso, projectMeta) {
    const stamp = formatStamp(timeIso);
    const prefix = projectMeta ? `PROJECT-ONLY list (${projectMeta.projectName}) · ` : "";
    if (!entries || !entries.length) {
      return projectMeta ? `PROJECT-ONLY list (${projectMeta.projectName}) — edits stay in this workbook only · ${stamp}` : "";
    }
    if (entries.length === 1) {
      const e = entries[0];
      return `${prefix}Last change: ${actionLabel(e.type) || "Modified"} "${e.material}" · ${stamp}`;
    }
    const counts = { added: 0, modified: 0, removed: 0 };
    for (const e of entries) counts[e.type] = (counts[e.type] || 0) + 1;
    const parts = Object.entries(counts).filter(([, v]) => v > 0).map(([k, v]) => `${v} ${k}`);
    return `${prefix}Last change: ${parts.join(", ")} · ${stamp}`;
  }

  async function ensureMasterNames(context, wb, log = noop) {
    const m = await openMaster(context);
    const wantedNames = new Set(NAMED_RANGES.map((d) => lc(d.name)));

    // 1. Delete worksheet-scoped shadow names + any legacy volatile Print_Area / Print_Titles
    wb.worksheets.load("items/name");
    await context.sync();
    wb.worksheets.items.forEach((ws) => ws.names.load("items/name"));
    await context.sync();
    let delShadow = false;
    wb.worksheets.items.forEach((ws) => {
      ws.names.items.forEach((nm) => {
        if (wantedNames.has(lc(nm.name))) { nm.delete(); delShadow = true; }
      });
    });
    if (delShadow) await context.sync();

    // 2. Load table columns + workbook defined names
    const colInfo = {};
    for (const def of NAMED_RANGES) {
      const header = m.headers[m.idx[def.key]];
      const range = m.t.columns.getItem(header).getDataBodyRange();
      range.load(["address", "rowCount"]);
      colInfo[def.name] = { header, range };
    }
    wb.names.load("items/name,items/formula");
    await context.sync();

    const existing = new Map();
    wb.names.items.forEach((nm) => existing.set(lc(nm.name), nm));

    const toCreate = [];
    let removedAny = false;
    for (const def of NAMED_RANGES) {
      const info = colInfo[def.name];
      const structFormula = `=${m.tableName}[${info.header}]`;
      const cur = existing.get(lc(def.name));
      let ok = false;
      if (cur) {
        const f = String(cur.formula || "");
        const broken = !f || f.includes("#REF!") || f.includes("#NAME?") || /\[\d+\]/.test(f);
        if (!broken && (lc(f) === lc(structFormula) || lc(f.replace(/\$/g, "")).includes(lc(info.range.address.replace(/\$/g, ""))))) {
          ok = true;
        } else {
          try { cur.delete(); removedAny = true; } catch { /* ignore */ }
        }
      }
      if (!ok) toCreate.push({ def, info, structFormula });
    }
    if (removedAny) { try { await context.sync(); } catch { /* ignore */ } }

    let usedRangeFallback = false;
    if (toCreate.length) {
      let batchOk = false;
      try {
        toCreate.forEach((item) => wb.names.add(item.def.name, item.structFormula));
        await context.sync();
        batchOk = true;
      } catch {
        batchOk = false;
      }
      if (batchOk) {
        try {
          const checks = toCreate.map((item) => ({ item, nm: wb.names.getItem(item.def.name).load("formula") }));
          await context.sync();
          const bad = checks.filter((c) => String(c.nm.formula || "").includes("#REF!"));
          if (bad.length) {
            bad.forEach((c) => c.nm.delete());
            await context.sync();
            bad.forEach((c) => wb.names.add(c.item.def.name, c.item.info.range));
            await context.sync();
            usedRangeFallback = true;
          }
        } catch { /* ignore */ }
      } else {
        for (const item of toCreate) {
          try {
            wb.names.add(item.def.name, item.info.range);
            await context.sync();
            usedRangeFallback = true;
          } catch (e) {
            log(`Warning: could not define ${item.def.name} (${e.message || e}).`);
          }
        }
      }
    }

    // 3. Clamp Primary Unit DV on Master to active range (max 300 rows — never 2000!)
    const dvEnd = Math.max(200, Math.min(400, 4 + (colInfo.Materials.range.rowCount || 55) + 80));
    try {
      const dv = m.ws.getRange(`C5:C${dvEnd}`).dataValidation;
      dv.clear();
      dv.rule = { list: { inCellDropDown: true, source: "kg/m,kg/m²,kg/each" } };
      dv.ignoreBlanks = true;
      dv.errorAlert = {
        showAlert: true,
        style: Excel.DataValidationAlertStyle.stop,
        title: "Pick a unit",
        message: "Choose kg/m, kg/m² or kg/each."
      };
      await context.sync();
    } catch { /* non-fatal */ }

    return { master: m, matRangeAddress: colInfo.Materials.range.address, usedRangeFallback };
  }

  /* Bind Materials dropdown strictly to bomT[Materials] data body (never bloated C6:C2000!) */
  async function applyBomValidation(context, ws, bomT, matInfo, log = noop) {
    let target = null;
    try {
      target = bomT ? bomT.columns.getItem("Materials").getDataBodyRange() : ws.getRange("C6:C65");
    } catch {
      target = ws.getRange("C6:C65");
    }
    const alertCfg = {
      showAlert: true,
      style: Excel.DataValidationAlertStyle.stop,
      title: "Not in the materials list",
      message: "Pick a material from the dropdown (or add it in the Materials tab first)."
    };
    const applyWith = async (source) => {
      target.dataValidation.clear();
      target.dataValidation.rule = { list: { inCellDropDown: true, source } };
      target.dataValidation.ignoreBlanks = true;
      target.dataValidation.prompt = { showPrompt: false, title: "", message: "" };
      target.dataValidation.errorAlert = alertCfg;
      await context.sync();
    };
    try {
      await applyWith("=Materials");
    } catch {
      try {
        await applyWith(`=${matInfo.matRangeAddress}`);
      } catch (e2) {
        log(`Warning: could not set Materials dropdown on ${ws.name} (${e2.message || e2}).`);
      }
    }
  }

  /* Purge volatile Print_Area / Print_Titles defined names on legacy sheets (NEVER touches ws.pageLayout or ws.freezePanes!) */
  async function purgeLegacyPrintNames(context, ws) {
    try {
      ws.names.load("items/name");
      await context.sync();
      let del = false;
      ws.names.items.forEach((nm) => {
        if (lc(nm.name).includes("print_area") || lc(nm.name).includes("print_titles")) {
          nm.delete();
          del = true;
        }
      });
      if (del) await context.sync();
    } catch { /* ignore */ }
  }

  /* Verify & upgrade formulas, table references, and active-table DV on a BOM sheet (never corrupts freezePanes/pageLayout/styles!) */
  async function repairBomSheet(context, wb, ws, matInfo, opts = {}, log = noop) {
    ws.tables.load("items/name");
    await context.sync();
    const bomT = firstTable(ws.tables.items, "bom_");
    const plT = firstTable(ws.tables.items, "pl_");
    if (!bomT) return;

    const bomName = bomT.name;
    const plName = plT ? plT.name : "";
    const bomBody = bomT.getDataBodyRange().load(["formulas", "rowCount"]);
    const plBody = plT ? plT.getDataBodyRange().load(["formulas", "rowCount"]) : null;
    const rO2 = ws.getRange("O2").load("formulas");
    const rAE1 = ws.getRange("AE1").load("values");
    await context.sync();

    pauseExcel(context);
    if (!isFormula(rO2.formulas[0][0]) || String(rO2.formulas[0][0]).includes("#REF!")) {
      rO2.formulas = [[canonical.O2()]];
    }
    // Always write native Excel 365 M4, AL6 (dynamic spill), AB4 & AD2 via .formulas so XLDAPR metadata is registered cleanly
    ws.getRange("M4").formulas = [[canonical.M4(bomName)]];
    ws.getRange("AL6").formulas = [[canonical.AL6(bomName)]];
    if (plName) {
      ws.getRange("AB4").formulas = [[canonical.AB4()]];
      ws.getRange("AD2").formulas = [[canonical.AD2(bomName, plName)]];
    }
    if (isBlank(rAE1.values[0][0])) rAE1.values = [[6000]];

    // Rewrite stale table refs (and upgrade formulas to pure Excel 365 LET/XMATCH/AL6# when requested)
    let bomChanged = false;
    const bomOut = bomBody.formulas.map((row, idx) => {
      const rNum = 6 + idx;
      return row.map((cell, cIdx) => {
        let fixed = cell;
        if (opts.srcBomName && opts.srcBomName !== bomName) fixed = replaceTableRef(fixed, opts.srcBomName, bomName);
        if (opts.origCopiedBomName && opts.origCopiedBomName !== bomName) fixed = replaceTableRef(fixed, opts.origCopiedBomName, bomName);
        if (opts.upgrade365) {
          if (cIdx === 0) fixed = canonical.bomSn(bomName, rNum);
          else if (cIdx === 8) fixed = canonical.bomI(bomName);
          else if (cIdx === 9) fixed = canonical.bomJ(bomName, rNum);
          else if (cIdx === 11) fixed = canonical.bomL(bomName);
          else if (cIdx === 12) fixed = canonical.bomM(bomName);
        }
        if (fixed !== cell) bomChanged = true;
        return fixed;
      });
    });
    if (bomChanged) bomBody.formulas = bomOut;

    if (plBody) {
      let plChanged = false;
      const plOut = plBody.formulas.map((row, idx) => {
        const rNum = 6 + idx;
        return row.map((cell, cIdx) => {
          let fixed = cell;
          if (opts.srcBomName && opts.srcBomName !== bomName) fixed = replaceTableRef(fixed, opts.srcBomName, bomName);
          if (opts.origCopiedBomName && opts.origCopiedBomName !== bomName) fixed = replaceTableRef(fixed, opts.origCopiedBomName, bomName);
          if (opts.srcPlName && opts.srcPlName !== plName) fixed = replaceTableRef(fixed, opts.srcPlName, plName);
          if (opts.origCopiedPlName && opts.origCopiedPlName !== plName) fixed = replaceTableRef(fixed, opts.origCopiedPlName, plName);
          if (opts.upgrade365) {
            if (cIdx === 0) fixed = canonical.plO(plName, rNum);
            else if (cIdx === 1) fixed = canonical.plP(plName);
            else if (cIdx === 2) fixed = canonical.plQ(plName);
            else if (cIdx === 3) fixed = canonical.plR(bomName, plName);
            else if (cIdx === 4) fixed = canonical.plS(bomName, plName, rNum);
            else if (cIdx === 5) fixed = canonical.plT(plName);
            else if (cIdx === 6) fixed = canonical.plU(plName, rNum);
            else if (cIdx === 11) fixed = canonical.plZ(plName, rNum);
            else if (cIdx === 13) fixed = canonical.plAB();
          }
          if (fixed !== cell) plChanged = true;
          return fixed;
        });
      });
      if (plChanged) plBody.formulas = plOut;
    }

    // Keep helper rows/columns hidden without touching freezePanes or pageLayout
    ws.getRange("5:5").rowHidden = true;
    ws.getRange("M:M").columnHidden = true;
    ws.getRange("AB:AB").columnHidden = true;
    ws.getRange("AL:AL").columnHidden = true;
    await context.sync();

    if (opts.purgePrintNames) {
      await purgeLegacyPrintNames(context, ws);
    }

    if (opts.ensureValidation || matInfo?.usedRangeFallback) {
      await applyBomValidation(context, ws, bomT, matInfo, log);
    }
  }

  /* ---------------- Template Import (Visually Identical 2-Sheet Excel 365 Template) ---------------- */
  async function importTemplate(log = noop) {
    if (!Office?.context?.requirements?.isSetSupported("ExcelApi", "1.13")) {
      throw new Error("Template import requires Microsoft 365 or Excel on the Web (ExcelApi 1.13).");
    }
    let b64 = window.BOM_TEMPLATE_B64;
    if (!b64) {
      log("Fetching lightweight template.xlsx ...");
      const res = await fetch(new URL("template.xlsx", location.href).href + "?t=" + Date.now(), { cache: "no-store" });
      if (!res.ok) throw new Error(`Could not load template.xlsx (${res.status}).`);
      const u = new Uint8Array(await res.arrayBuffer());
      let bin = "";
      for (let i = 0; i < u.length; i += 0x8000) bin += String.fromCharCode(...u.subarray(i, i + 0x8000));
      b64 = btoa(bin);
    }

    await runExcel(async (context) => {
      const wb = context.workbook;
      let prevCalc = null;
      try {
        wb.application.load("calculationMode");
        await context.sync();
        prevCalc = wb.application.calculationMode;
        if (prevCalc !== Excel.CalculationMode.manual) {
          wb.application.calculationMode = Excel.CalculationMode.manual;
          await context.sync();
        }
      } catch { prevCalc = null; }

      try {
        wb.worksheets.load("items/name");
        await context.sync();
        const before = new Set(wb.worksheets.items.map((w) => w.name));

        log("Importing lightweight Master & BOM Blank sheets ...");
        wb.insertWorksheetsFromBase64(b64, { positionType: Excel.WorksheetPositionType.end });
        await context.sync();

        wb.worksheets.load("items/name");
        await context.sync();
        const added = wb.worksheets.items.filter((w) => !before.has(w.name));
        added.forEach((w) => w.tables.load("items/name"));
        await context.sync();

        const master = added.find((w) => lc(w.name) === lc(MASTER)) ||
                       added.find((w) => w.tables.items.some((t) => t.name === "tblMat"));
        const blank = added.find((w) => lc(w.name) === lc(BLANK)) ||
                      added.find((w) => w.tables.items.some((t) => t.name.startsWith("bom_")));

        if (!master || !blank) {
          added.forEach((w) => { try { w.delete(); } catch { /* ignore */ } });
          await context.sync();
          throw new Error("Template did not contain Master Unique Materials and BOM Blank.");
        }

        added.forEach((w) => {
          if (w !== master && w !== blank && /^(README|Formulas)$/i.test(w.name.trim())) w.delete();
        });
        if (master.name !== MASTER) master.name = MASTER;
        if (blank.name !== BLANK) blank.name = BLANK;
        await context.sync();

        const matInfo = await ensureMasterNames(context, wb, log);
        // Upgrade BOM Blank via native Excel .formulas to lightweight Excel 365 LET/XMATCH/AL6# formulas (zero OOXML metadata errors!)
        await repairBomSheet(context, wb, blank, matInfo, { upgrade365: true, ensureValidation: true }, log);

        blank.visibility = Excel.SheetVisibility.hidden;
        await context.sync();
      } finally {
        try {
          if (prevCalc !== null) wb.application.calculationMode = prevCalc;
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        } catch { /* ignore */ }
      }
    });
    log("Lightweight Excel 365 template imported (Master Unique Materials + hidden BOM Blank).");
  }

  /* ---------------- Inspect Workbook ---------------- */
  async function inspectWorkbook() {
    const out = {
      hasMaster: false,
      hasBlank: false,
      needsTemplate: false,
      projectTitle: "",
      projectMode: null,
      allSheets: [],
      bomSheets: [],
      convertedSheets: []
    };
    const SKIP = new Set(["master unique materials", "bom blank", "compiled purchase list", "formulas", "readme"]);
    await runExcel(async (context) => {
      const wb = context.workbook;
      wb.worksheets.load("items/name,items/visibility,items/position");
      await context.sync();
      const sheets = wb.worksheets.items;
      sheets.forEach((ws) => ws.tables.load("items/name"));
      await context.sync();

      out.allSheets = sheets.map((ws) => ws.name);
      out.hasMaster = sheets.some((ws) => lc(ws.name) === lc(MASTER));
      out.hasBlank = sheets.some((ws) => lc(ws.name) === lc(BLANK));
      out.projectMode = await readProjectProp(context, wb);

      const bomWs = sheets.filter((ws) => Boolean(firstTable(ws.tables.items, "bom_")));
      const candConv = sheets.filter((ws) =>
        ws.visibility === "Visible" && !SKIP.has(lc(ws.name)) && !firstTable(ws.tables.items, "bom_")
      );

      const a1Loads = bomWs.map((ws) => ws.getRange("A1").load("values"));
      const convLoads = candConv.map((ws) => ws.getRange("A1:P4").load("values"));
      await context.sync();

      bomWs.forEach((ws, idx) => {
        const hidden = ws.visibility !== "Visible";
        out.bomSheets.push({ name: ws.name, hidden });
        const t = String(a1Loads[idx].values[0][0] || "");
        if (!out.projectTitle && !hidden && t.startsWith(TITLE_PREFIX) && !t.includes("[")) {
          out.projectTitle = t.slice(TITLE_PREFIX.length).trim();
        }
      });

      candConv.forEach((ws, idx) => {
        const v = convLoads[idx].values;
        const a1 = lc(v[0]?.[0]);
        const o2 = lc(v[1]?.[14]);
        const b3 = lc(v[2]?.[1]), b4 = lc(v[3]?.[1]);
        const c3 = lc(v[2]?.[2]), c4 = lc(v[3]?.[2]);
        const isBom = a1.includes("bill of materials") || o2.includes("purchase list") ||
          ((b3.includes("description") || b4.includes("description")) && (c3.includes("material") || c4.includes("material")));
        if (isBom) out.convertedSheets.push(ws.name);
      });
    });
    out.needsTemplate = !out.hasMaster && out.bomSheets.length === 0 && out.convertedSheets.length === 0;
    return out;
  }

  /* ---------------- Clear Copied Sheet Inputs (preserving formulas & restoring PO Weight footer) ---------------- */
  async function clearTableInputs(context, table, colNames, isPl) {
    table.columns.load("items/name");
    await context.sync();
    const targets = [];
    for (const col of table.columns.items) {
      if (colNames.includes(col.name)) {
        const rng = col.getDataBodyRange().load(["formulas", "rowIndex"]);
        targets.push({ name: col.name, rng });
      }
    }
    await context.sync();
    for (const item of targets) {
      const startRow = (item.rng.rowIndex || 5) + 1;
      let dirty = false;
      const out = item.rng.formulas.map((row, idx) => {
        const v = row[0];
        if (isPl && item.name === "PO Weight") {
          const f = canonical.poWeight(startRow + idx);
          if (v !== f) { dirty = true; return [f]; }
          return [v];
        }
        if (!isFormula(v) && v !== "") { dirty = true; return [""]; }
        return [v];
      });
      if (dirty) item.rng.formulas = out;
    }
    await context.sync();
  }

  /* ---------------- Add Lightweight BOM Sheet (All v2.3–v2.6 Lag & Glitch Fixes Included) ---------------- */
  async function addBomSheet(opts, log = noop) {
    const newName = String(opts.name || "").trim();
    const warnings = [];
    await runExcel(async (context) => {
      const wb = context.workbook;
      let prevCalc = null;
      let unhidBlank = false;
      let blankRef = null;

      try {
        wb.application.load("calculationMode");
        await context.sync();
        prevCalc = wb.application.calculationMode;
        if (prevCalc !== Excel.CalculationMode.manual) {
          wb.application.calculationMode = Excel.CalculationMode.manual;
          await context.sync();
        }
      } catch { prevCalc = null; }

      try {
        const matInfo = await ensureMasterNames(context, wb, log);
        wb.worksheets.load("items/name,items/position,items/visibility");
        wb.tables.load("items/name");
        await context.sync();

        const err = checkSheetName(newName, wb.worksheets.items.map((w) => w.name));
        if (err) throw new Error(err);

        const src = wb.worksheets.items.find((w) => w.name === opts.source);
        if (!src) throw new Error(`Source sheet "${opts.source}" was not found.`);

        wb.worksheets.items.forEach((ws) => ws.tables.load("items/name"));
        await context.sync();
        const srcBom = firstTable(src.tables.items, "bom_");
        const srcPl = firstTable(src.tables.items, "pl_");
        if (!srcBom) throw new Error("Source sheet has no BOM table (bom_...).");

        // Place new sheet after the last visible BOM sheet
        let anchor = src;
        let lastPos = -1;
        for (const ws of wb.worksheets.items) {
          if (ws.name !== BLANK && firstTable(ws.tables.items, "bom_") && ws.position > lastPos) {
            anchor = ws;
            lastPos = ws.position;
          }
        }

        // Temporarily unhide BOM Blank (without activating it, avoiding any tab-switch repaint)
        if (src.visibility !== "Visible") {
          blankRef = src;
          src.visibility = Excel.SheetVisibility.visible;
          unhidBlank = true;
          await context.sync();
        }

        log(`Creating "${newName}" from "${opts.source}" ...`);
        pauseExcel(context);
        const copy = src.copy(Excel.WorksheetPositionType.after, anchor);
        copy.name = newName;
        copy.visibility = Excel.SheetVisibility.visible;
        copy.tables.load("items/name");
        await context.sync();

        const bomT = firstTable(copy.tables.items, "bom_");
        const plT = firstTable(copy.tables.items, "pl_");
        if (!bomT) {
          copy.delete();
          await context.sync();
          throw new Error("Copied sheet was missing its BOM table and was removed.");
        }
        const origCopiedBomName = bomT.name;
        const origCopiedPlName = plT ? plT.name : "";
        const al6Check = copy.getRange("AL6").load("formulas");

        wb.tables.load("items/name");
        await context.sync();
        const needs365Upgrade = !isFormula(al6Check.formulas[0][0]);
        const taken = wb.tables.items.map((t) => t.name);
        const id = safeId(newName);
        const bomName = uniqueName(`bom_${id}`, taken);
        const plName = uniqueName(`pl_${id}`, taken.concat([bomName]));

        pauseExcel(context);
        bomT.name = bomName;
        if (plT) plT.name = plName;
        copy.getRange("A2").values = [[newName]];
        if (opts.project && String(opts.project).trim()) {
          copy.getRange("A1").values = [[TITLE_PREFIX + String(opts.project).trim()]];
        }
        await context.sync();

        if (opts.clear) {
          log("Clearing entered items ...");
          await clearTableInputs(context, bomT, BOM_INPUT_COLS, false);
          if (plT) await clearTableInputs(context, plT, PL_INPUT_COLS, true);
        }

        await repairBomSheet(context, wb, copy, matInfo, {
          srcBomName: srcBom.name,
          srcPlName: srcPl ? srcPl.name : "",
          origCopiedBomName,
          origCopiedPlName,
          upgrade365: needs365Upgrade,
          ensureValidation: opts.source !== BLANK
        }, log);

        if (unhidBlank && blankRef) {
          blankRef.visibility = Excel.SheetVisibility.hidden;
          unhidBlank = false;
        }

        // Single clean activation & B6 selection — ZERO freezePanes, ZERO pageLayout, ZERO fullRebuild, ZERO tab flipping!
        copy.activate();
        copy.getRange("B6").select();
        await context.sync();
      } finally {
        if (unhidBlank && blankRef) {
          try { blankRef.visibility = Excel.SheetVisibility.hidden; } catch { /* ignore */ }
        }
        try {
          if (prevCalc !== null) wb.application.calculationMode = prevCalc;
          await context.sync();
        } catch { /* ignore */ }
      }
    });
    return { name: newName, warnings };
  }

  /* ---------------- One-Click Repair & Excel 365 Upgrade ---------------- */
  async function repairWorkbook(opts = {}, log = noop) {
    const selected = Array.isArray(opts.sheets) ? new Set(opts.sheets.map(lc)) : null;
    if (selected && selected.size === 0) throw new Error("Select at least one BOM sheet to repair.");

    let repaired = 0;
    await runExcel(async (context) => {
      const wb = context.workbook;
      let prevCalc = null;
      try {
        wb.application.load("calculationMode");
        await context.sync();
        prevCalc = wb.application.calculationMode;
        if (prevCalc !== Excel.CalculationMode.manual) {
          wb.application.calculationMode = Excel.CalculationMode.manual;
          await context.sync();
        }
      } catch { prevCalc = null; }

      try {
        log("Verifying Master sheet & workbook named ranges ...");
        const matInfo = await ensureMasterNames(context, wb, log);

        // Purge workbook-level volatile Print_Area / Print_Titles from legacy v1.x/v2.x workbooks
        wb.names.load("items/name");
        wb.worksheets.load("items/name,items/visibility");
        await context.sync();
        let purged = false;
        wb.names.items.forEach((nm) => {
          if (lc(nm.name).includes("print_area") || lc(nm.name).includes("print_titles")) {
            try { nm.delete(); purged = true; } catch { /* ignore */ }
          }
        });
        if (purged) await context.sync();

        const sheets = wb.worksheets.items;
        sheets.forEach((ws) => ws.tables.load("items/name"));
        await context.sync();

        for (const ws of sheets) {
          if (!firstTable(ws.tables.items, "bom_")) continue;
          if (selected && !selected.has(lc(ws.name))) continue;
          log(`Upgrading & repairing "${ws.name}" to Excel 365 formulas ...`);
          await repairBomSheet(context, wb, ws, matInfo, {
            upgrade365: true,
            purgePrintNames: true,
            ensureValidation: true
          }, log);
          repaired++;
        }
      } finally {
        try {
          if (prevCalc !== null) wb.application.calculationMode = prevCalc;
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        } catch { /* ignore */ }
      }
    });

    log(`Repaired & upgraded ${repaired} BOM sheet(s).`);
    return { sheetsRepaired: repaired };
  }

  /* ---------------- Materials CRUD & Bulk Apply ---------------- */
  async function listMaterials() {
    const out = { rows: [], categories: [], projectMode: null };
    await runExcel(async (context) => {
      const m = await openMaster(context);
      const body = m.t.getDataBodyRange().load("values");
      out.projectMode = await readProjectProp(context, context.workbook);
      await context.sync();
      const cats = new Set();
      body.values.forEach((r, i) => {
        const obj = rowToMaterial(r, i, m.idx);
        if (!obj.material) return;
        out.rows.push(obj);
        if (obj.category) cats.add(obj.category);
      });
      out.categories = [...cats].sort();
    });
    return out;
  }

  async function refreshRangeBasedDropdowns(context, wb, matInfo) {
    if (!matInfo?.usedRangeFallback) return;
    wb.worksheets.load("items/name");
    await context.sync();
    wb.worksheets.items.forEach((ws) => ws.tables.load("items/name"));
    await context.sync();
    for (const ws of wb.worksheets.items) {
      const bomT = firstTable(ws.tables.items, "bom_");
      if (bomT) await applyBomValidation(context, ws, bomT, matInfo);
    }
  }

  async function addMaterial(raw) {
    const item = window.BomStore.normalizeMaterial(raw);
    const nowIso = new Date().toISOString();
    item.lastAction = "added";
    item.lastModified = nowIso;

    await runExcel(async (context) => {
      const wb = context.workbook;
      const m = await openMaster(context);
      await ensureAuditColumns(context, m);
      const projectMeta = await readProjectProp(context, wb);
      const cur = m.t.getDataBodyRange().load("values");
      await context.sync();
      const names = cur.values.map((r) => String(r[m.idx.material] || "").trim()).filter(Boolean);
      const errs = validateMaterial(item, names, "add");
      if (errs.length) throw new Error(errs[0]);

      // Fill first blank row if present (e.g. blank Project-Only sheet), else append row
      const blankIdx = cur.values.findIndex((r) => !String(r[m.idx.material] || "").trim());
      let targetRow;
      if (blankIdx >= 0) {
        targetRow = cur.getRow(blankIdx);
      } else {
        m.t.rows.add();
        await context.sync();
        const body = m.t.getDataBodyRange().load("rowCount");
        await context.sync();
        targetRow = body.getRow(body.rowCount - 1);
      }

      const vals = {
        material: item.material, category: item.category, unit: item.unit,
        lin: item.lin, area: item.area, piece: item.piece, sl: item.sl, sw: item.sw,
        remarks: item.remarks, lastAction: "Added", lastModified: formatStamp(nowIso)
      };
      for (const [k, v] of Object.entries(vals)) {
        if (m.idx[k] >= 0) targetRow.getCell(0, m.idx[k]).values = [[v === null ? "" : v]];
      }
      await context.sync();

      await writeMasterBanner(
        context, m.ws,
        formatBannerSummary([{ material: item.material, type: "added" }], nowIso, projectMeta),
        Boolean(projectMeta)
      );
      const matInfo = await ensureMasterNames(context, wb);
      await refreshRangeBasedDropdowns(context, wb, matInfo);
      wb.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
    return item;
  }

  async function updateMaterial(originalNameOrRaw, maybeRaw) {
    const hasOrig = typeof originalNameOrRaw === "string" && maybeRaw && typeof maybeRaw === "object";
    const lookupName = hasOrig ? originalNameOrRaw.trim() : "";
    const item = window.BomStore.normalizeMaterial(hasOrig ? maybeRaw : originalNameOrRaw);
    const nowIso = new Date().toISOString();
    item.lastAction = "modified";
    item.lastModified = nowIso;

    await runExcel(async (context) => {
      const wb = context.workbook;
      const m = await openMaster(context);
      await ensureAuditColumns(context, m);
      const projectMeta = await readProjectProp(context, wb);
      const body = m.t.getDataBodyRange().load("values");
      await context.sync();
      const targetKey = lc(lookupName || item.material);
      const at = body.values.findIndex((r) => lc(r[m.idx.material]) === targetKey);
      if (at < 0) throw new Error(`"${lookupName || item.material}" is no longer in the sheet.`);
      const renamed = lc(item.material) !== targetKey;
      const otherNames = body.values
        .map((r, i) => (i === at ? "" : String(r[m.idx.material] || "").trim()))
        .filter(Boolean);
      const errs = validateMaterial(item, otherNames, renamed ? "add" : "edit");
      if (errs.length) throw new Error(errs[0]);

      const row = body.getRow(at);
      for (const k of ["material", "category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"]) {
        if (m.idx[k] >= 0) row.getCell(0, m.idx[k]).values = [[item[k] === null ? "" : item[k]]];
      }
      if (m.idx.lastAction >= 0) row.getCell(0, m.idx.lastAction).values = [["Modified"]];
      if (m.idx.lastModified >= 0) row.getCell(0, m.idx.lastModified).values = [[formatStamp(nowIso)]];
      await context.sync();

      await writeMasterBanner(
        context, m.ws,
        formatBannerSummary([{ material: item.material, type: "modified" }], nowIso, projectMeta),
        Boolean(projectMeta)
      );
      wb.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
    return item;
  }

  async function deleteMaterial(materialName) {
    const target = String(materialName || "").trim();
    const nowIso = new Date().toISOString();
    await runExcel(async (context) => {
      const wb = context.workbook;
      const m = await openMaster(context);
      const projectMeta = await readProjectProp(context, wb);
      const body = m.t.getDataBodyRange().load(["values", "rowCount"]);
      await context.sync();
      const at = body.values.findIndex((r) => lc(r[m.idx.material]) === lc(target));
      if (at < 0) throw new Error(`"${target}" was not found.`);

      if (body.rowCount <= 1) {
        const row = body.getRow(0);
        Object.keys(MASTER_HEADERS).forEach((k) => {
          if (m.idx[k] >= 0) row.getCell(0, m.idx[k]).values = [[""]];
        });
      } else {
        m.t.rows.getItemAt(at).delete();
      }
      await context.sync();

      await writeMasterBanner(
        context, m.ws,
        formatBannerSummary([{ material: target, type: "removed" }], nowIso, projectMeta),
        Boolean(projectMeta)
      );
      const matInfo = await ensureMasterNames(context, wb);
      await refreshRangeBasedDropdowns(context, wb, matInfo);
      wb.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
    return target;
  }

  /* Append multiple previously saved materials into the active Master sheet (skips duplicates by name) */
  async function addMaterialsBulk(rawItems) {
    const clean = (Array.isArray(rawItems) ? rawItems : [])
      .map((r) => window.BomStore.normalizeMaterial(r))
      .filter((m) => m.material);
    if (!clean.length) return { added: [], skipped: [] };

    const nowIso = new Date().toISOString();
    const added = [];
    const skipped = [];

    await runExcel(async (context) => {
      const wb = context.workbook;
      const m = await openMaster(context);
      await ensureAuditColumns(context, m);
      const projectMeta = await readProjectProp(context, wb);

      let prevCalc = null;
      try {
        wb.application.load("calculationMode");
        await context.sync();
        prevCalc = wb.application.calculationMode;
        if (prevCalc !== Excel.CalculationMode.manual) {
          wb.application.calculationMode = Excel.CalculationMode.manual;
          await context.sync();
        }
      } catch { prevCalc = null; }

      try {
        const body = m.t.getDataBodyRange().load("values");
        await context.sync();
        const existing = new Set();
        const blanks = [];
        body.values.forEach((row, i) => {
          const name = String(row[m.idx.material] || "").trim();
          if (!name) blanks.push(i);
          else existing.add(lc(name));
        });

        const toAdd = [];
        for (const item of clean) {
          const key = lc(item.material);
          if (existing.has(key)) {
            skipped.push(item.material);
            continue;
          }
          existing.add(key);
          toAdd.push({ ...item, lastAction: "added", lastModified: nowIso });
        }

        if (toAdd.length) {
          const writeInto = (range, idx, item) => {
            for (const k of ["material", "category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"]) {
              if (m.idx[k] >= 0) range.getCell(idx, m.idx[k]).values = [[item[k] === null ? "" : item[k]]];
            }
            if (m.idx.lastAction >= 0) range.getCell(idx, m.idx.lastAction).values = [["Added"]];
            if (m.idx.lastModified >= 0) range.getCell(idx, m.idx.lastModified).values = [[formatStamp(nowIso)]];
          };

          const useBlank = toAdd.slice(0, blanks.length);
          const extra = toAdd.slice(blanks.length);
          pauseExcel(context);
          useBlank.forEach((item, k) => writeInto(body, blanks[k], item));
          await context.sync();

          if (extra.length) {
            for (let k = 0; k < extra.length; k++) m.t.rows.add();
            await context.sync();
            const b2 = m.t.getDataBodyRange().load("rowCount");
            await context.sync();
            const first = b2.rowCount - extra.length;
            pauseExcel(context);
            extra.forEach((item, j) => writeInto(b2, first + j, item));
            await context.sync();
          }

          toAdd.forEach((it) => added.push(it));
          await writeMasterBanner(
            context, m.ws,
            formatBannerSummary(toAdd.map((it) => ({ material: it.material, type: "added" })), nowIso, projectMeta),
            Boolean(projectMeta)
          );
          const matInfo = await ensureMasterNames(context, wb);
          await refreshRangeBasedDropdowns(context, wb, matInfo);
        }
      } finally {
        try {
          if (prevCalc !== null) wb.application.calculationMode = prevCalc;
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        } catch { /* ignore */ }
      }
    });
    return { added, skipped };
  }

  async function stampAudit(entries, timeIso, projectMeta) {
    if (!entries || !entries.length) return;
    await runExcel(async (context) => {
      const m = await openMaster(context);
      const rowEntries = entries.filter((e) => e.type !== "removed" && e.material);
      if (rowEntries.length) {
        await ensureAuditColumns(context, m);
        const body = m.t.getDataBodyRange().load("values");
        await context.sync();
        const rowMap = new Map();
        body.values.forEach((r, i) => {
          const k = lc(r[m.idx.material]);
          if (k && !rowMap.has(k)) rowMap.set(k, i);
        });
        let touched = false;
        for (const e of rowEntries) {
          const at = rowMap.get(lc(e.material));
          if (at === undefined) continue;
          if (m.idx.lastAction >= 0) body.getCell(at, m.idx.lastAction).values = [[actionLabel(e.type) || "Modified"]];
          if (m.idx.lastModified >= 0) body.getCell(at, m.idx.lastModified).values = [[formatStamp(e.time || timeIso)]];
          touched = true;
        }
        if (touched) await context.sync();
      }
      await writeMasterBanner(
        context, m.ws,
        formatBannerSummary(entries, timeIso, projectMeta),
        Boolean(projectMeta)
      );
    });
  }

  /* Replace or sync the Master Unique Materials table from a given list (batched for speed) */
  async function applyMaterialList(list, opts = {}) {
    const syncRemovals = Boolean(opts.syncRemovals);
    const replaceAll = Boolean(opts.replaceAll);
    const projectMeta = opts.projectMeta === undefined ? undefined : opts.projectMeta;
    const normalized = list.map(window.BomStore.normalizeMaterial).filter((m) => m.material);
    const res = { added: 0, updated: 0, removed: 0 };

    await runExcel(async (context) => {
      const wb = context.workbook;
      const m = await openMaster(context);
      await ensureAuditColumns(context, m);
      if (projectMeta !== undefined) {
        await writeProjectProp(context, wb, projectMeta);
      }
      const activeProject = projectMeta !== undefined ? projectMeta : await readProjectProp(context, wb);

      let prevCalc = null;
      try {
        wb.application.load("calculationMode");
        await context.sync();
        prevCalc = wb.application.calculationMode;
        if (prevCalc !== Excel.CalculationMode.manual) {
          wb.application.calculationMode = Excel.CalculationMode.manual;
          await context.sync();
        }
      } catch { prevCalc = null; }

      try {
        const body = m.t.getDataBodyRange().load(["values", "formulas", "rowCount"]);
        await context.sync();

        if (replaceAll) {
          // Fast full replacement (used for switching Project-Only <-> Global Master)
          const targetCount = Math.max(1, normalized.length);
          if (body.rowCount > targetCount) {
            for (let i = body.rowCount - 1; i >= targetCount; i--) m.t.rows.getItemAt(i).delete();
            await context.sync();
          } else if (body.rowCount < targetCount) {
            for (let i = body.rowCount; i < targetCount; i++) m.t.rows.add();
            await context.sync();
          }
          const b2 = m.t.getDataBodyRange().load("rowCount");
          await context.sync();

          const grid = [];
          for (let i = 0; i < targetCount; i++) {
            const rNum = 5 + i;
            const item = normalized[i];
            const row = new Array(m.ncols).fill("");
            if (item) {
              if (m.idx.material >= 0) row[m.idx.material] = item.material;
              if (m.idx.category >= 0) row[m.idx.category] = item.category;
              if (m.idx.unit >= 0) row[m.idx.unit] = item.unit;
              if (m.idx.lin >= 0) row[m.idx.lin] = item.lin ?? "";
              if (m.idx.area >= 0) row[m.idx.area] = item.area ?? "";
              if (m.idx.piece >= 0) row[m.idx.piece] = item.piece ?? "";
              if (m.idx.sl >= 0) row[m.idx.sl] = item.sl ?? "";
              if (m.idx.sw >= 0) row[m.idx.sw] = item.sw ?? "";
              if (m.idx.remarks >= 0) row[m.idx.remarks] = item.remarks;
              if (m.idx.lastAction >= 0 && item.lastAction) row[m.idx.lastAction] = actionLabel(item.lastAction);
              if (m.idx.lastModified >= 0 && item.lastModified) row[m.idx.lastModified] = formatStamp(item.lastModified);
            }
            // Preserve calculated Sheet Area (col 8) & Sheet Weight (col 9) formulas
            const cArea = m.headers.findIndex((h) => h.startsWith("Sheet Area"));
            const cWt = m.headers.findIndex((h) => h.startsWith("Sheet Weight"));
            if (cArea >= 0) row[cArea] = `=IF(AND(ISNUMBER(G${rNum}),ISNUMBER(H${rNum})),G${rNum}*H${rNum}/10^6,"")`;
            if (cWt >= 0) row[cWt] = `=IF(AND(ISNUMBER(I${rNum}),ISNUMBER(E${rNum})),I${rNum}*E${rNum},"")`;
            grid.push(row);
          }
          pauseExcel(context);
          b2.formulas = grid;
          await context.sync();
          res.added = normalized.length;
        } else {
          // Incremental sync (adds missing, updates modified, optionally removes deleted)
          const byName = new Map();
          const blanks = [];
          body.values.forEach((row, i) => {
            const name = String(row[m.idx.material] || "").trim();
            if (!name) blanks.push(i);
            else if (!byName.has(lc(name))) byName.set(lc(name), i);
          });
          const wanted = new Map(normalized.map((item) => [lc(item.material), item]));
          const toAdd = [];

          pauseExcel(context);
          for (const item of normalized) {
            const at = byName.get(lc(item.material));
            if (at === undefined) { toAdd.push(item); continue; }
            const curObj = rowToMaterial(body.values[at], at, m.idx);
            const ch = window.BomStore.changedFields(curObj, item);
            let did = false;
            for (const k of ch) {
              if (m.idx[k] < 0 || isFormula(body.formulas[at][m.idx[k]])) continue;
              body.getCell(at, m.idx[k]).values = [[item[k] ?? ""]];
              did = true;
            }
            if (did) {
              if (m.idx.lastAction >= 0 && item.lastAction) body.getCell(at, m.idx.lastAction).values = [[actionLabel(item.lastAction)]];
              if (m.idx.lastModified >= 0 && item.lastModified) body.getCell(at, m.idx.lastModified).values = [[formatStamp(item.lastModified)]];
              res.updated++;
            }
          }

          const writeInto = (range, idx, item) => {
            for (const k of ["material", "category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"]) {
              if (m.idx[k] >= 0) range.getCell(idx, m.idx[k]).values = [[item[k] ?? ""]];
            }
            if (m.idx.lastAction >= 0 && item.lastAction) range.getCell(idx, m.idx.lastAction).values = [[actionLabel(item.lastAction)]];
            if (m.idx.lastModified >= 0 && item.lastModified) range.getCell(idx, m.idx.lastModified).values = [[formatStamp(item.lastModified)]];
          };

          const useBlank = toAdd.slice(0, blanks.length);
          const extra = toAdd.slice(blanks.length);
          useBlank.forEach((item, k) => writeInto(body, blanks[k], item));
          await context.sync();

          if (extra.length) {
            for (let k = 0; k < extra.length; k++) m.t.rows.add();
            await context.sync();
            const b2 = m.t.getDataBodyRange().load("rowCount");
            await context.sync();
            const first = b2.rowCount - extra.length;
            pauseExcel(context);
            extra.forEach((item, j) => writeInto(b2, first + j, item));
            await context.sync();
          }

          if (syncRemovals && wanted.size > 0) {
            const b3 = m.t.getDataBodyRange().load(["values", "rowCount"]);
            await context.sync();
            const delIdx = [];
            b3.values.forEach((row, i) => {
              const name = String(row[m.idx.material] || "").trim();
              if (name && !wanted.has(lc(name))) delIdx.push(i);
            });
            if (delIdx.length && delIdx.length < b3.rowCount) {
              for (let i = delIdx.length - 1; i >= 0; i--) m.t.rows.getItemAt(delIdx[i]).delete();
              await context.sync();
              res.removed = delIdx.length;
            }
          }
          res.added = toAdd.length;
        }

        await writeMasterBanner(
          context, m.ws,
          formatBannerSummary(opts.lastChange?.items || [], opts.lastChange?.time, activeProject),
          Boolean(activeProject)
        );
        const matInfo = await ensureMasterNames(context, wb);
        await refreshRangeBasedDropdowns(context, wb, matInfo);
      } finally {
        try {
          if (prevCalc !== null) wb.application.calculationMode = prevCalc;
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        } catch { /* ignore */ }
      }
    });
    return res;
  }

  /* Collect all material names currently referenced across visible BOM sheets */
  async function getUsedMaterialsInWorkbook() {
    const used = new Set();
    await runExcel(async (context) => {
      const wb = context.workbook;
      wb.worksheets.load("items/name,items/visibility");
      await context.sync();
      wb.worksheets.items.forEach((ws) => ws.tables.load("items/name"));
      await context.sync();
      const ranges = [];
      for (const ws of wb.worksheets.items) {
        if (ws.name === BLANK && ws.visibility !== "Visible") continue;
        const bomT = firstTable(ws.tables.items, "bom_");
        if (!bomT) continue;
        try {
          ranges.push(bomT.columns.getItem("Materials").getDataBodyRange().load("values"));
        } catch { /* ignore */ }
      }
      await context.sync();
      for (const r of ranges) {
        for (const row of r.values) {
          const name = String(row[0] || "").trim();
          if (name) used.add(lc(name));
        }
      }
    });
    return used;
  }

  /* ---------------- Zero-Hook On-Demand Master Sheet Edit Detection ----------------
   * IMPORTANT: We intentionally NEVER register ws.onChanged.add(...)!
   * Attaching Worksheet.onChanged in Excel Desktop installs a persistent COM/WebView2 event hook
   * that causes workbook-wide typing lag whenever the add-in pane is open (present since v1.8).
   * Instead, direct edits typed on Master Unique Materials are detected on-demand whenever the
   * user switches to Materials, clicks Refresh, saves, exports, or isolates a project list.
   */
  const snapState = { rows: null, busy: false };

  function setMasterSnap(list) {
    if (Array.isArray(list)) snapState.rows = list.map(window.BomStore.normalizeMaterial);
  }

  async function syncDirectMasterEdits(onEvent) {
    if (snapState.busy) return null;
    snapState.busy = true;
    try {
      const cur = await listMaterials();
      if (!snapState.rows) {
        snapState.rows = cur.rows;
        return { cur, detected: false };
      }
      const d = window.BomStore.diffLists(cur.rows, snapState.rows);
      if (!d.added.length && !d.modified.length && !d.removed.length) {
        return { cur, detected: false };
      }

      const nowIso = new Date().toISOString();
      const items = [];
      d.added.forEach((m) => items.push({ type: "added", material: { ...m, lastAction: "added", lastModified: nowIso }, fields: [], time: nowIso }));
      d.modified.forEach((c) => items.push({ type: "modified", material: { ...c.next, lastAction: "modified", lastModified: nowIso }, fields: c.fields, time: nowIso }));
      d.removed.forEach((m) => items.push({ type: "removed", material: m, fields: [], time: nowIso }));

      const lastChange = window.BomStore.buildLastChange("sheet", items, nowIso);
      const stampEntries = items.map((it) => ({ material: it.material.material, type: it.type, time: nowIso }));

      try { await stampAudit(stampEntries, nowIso, cur.projectMode); } catch { /* non-fatal */ }

      // Update the rows with fresh audit timestamps before persisting
      const stampedMap = new Map(items.filter((i) => i.type !== "removed").map((i) => [lc(i.material.material), i.material]));
      const updatedRows = cur.rows.map((r) => stampedMap.get(lc(r.material)) || r);

      if (cur.projectMode) {
        await window.BomStore.saveProjectList(cur.projectMode, updatedRows, lastChange);
      } else {
        await window.BomStore.writeGlobalMaster(updatedRows, lastChange);
      }
      snapState.rows = updatedRows;
      const evt = { source: "sheet", time: nowIso, projectMode: cur.projectMode, items };
      if (typeof onEvent === "function") onEvent(evt);
      return { cur: { ...cur, rows: updatedRows }, detected: true, evt };
    } catch {
      return null;
    } finally {
      snapState.busy = false;
    }
  }

  window.BomSheets = {
    MASTER, BLANK, PROP_PROJECT, UNITS,
    runExcel, pauseExcel, safeId, uniqueName, checkSheetName,
    formatStamp, actionLabel, materialStatus, validateMaterial, canonical,
    readProjectProp, writeProjectProp, getProjectMeta, setProjectMeta,
    openMaster, ensureMasterNames, applyBomValidation, repairBomSheet,
    importTemplate, inspectWorkbook, addBomSheet, repairWorkbook,
    listMaterials, addMaterial, addMaterialsBulk, updateMaterial, deleteMaterial,
    applyMaterialList, stampAudit, getUsedMaterialsInWorkbook,
    setMasterSnap, syncDirectMasterEdits
  };
})();
