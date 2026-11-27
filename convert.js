/* convert.js - Lightweight Selective Two-Way Converter (Flatten, Restore Automated & Plain-Sheet Repair) (3.0v patch2) */
(() => {
  "use strict";

  const COMPILED = "Compiled Purchase List";
  const PL_DEFAULT_CAP = 40;
  const EXTRA_EDIT_ROWS = 12;
  const lc = (s) => String(s ?? "").trim().toLowerCase();
  const isBlank = (v) => v === "" || v === null || v === undefined;
  const toNum = (v) => {
    if (isBlank(v)) return null;
    const n = Number(String(v).replace(/,/g, ""));
    return Number.isFinite(n) ? n : null;
  };
  const noop = () => {};

  function qSheet(name) {
    return `'${String(name).replace(/'/g, "''")}'`;
  }

  function firstTable(items, prefix) {
    return items.find((t) => t.name.startsWith(prefix)) || null;
  }

  /* Build fast lookup map of Master Unique Materials */
  async function loadMasterMap(context) {
    const map = new Map();
    const list = [];
    try {
      const m = await window.BomSheets.openMaster(context);
      const body = m.t.getDataBodyRange().load("values");
      await context.sync();
      for (const r of body.values) {
        const name = String(r[m.idx.material] || "").trim();
        if (!name) continue;
        const unit = String(r[m.idx.unit] || "").trim();
        const lin = toNum(r[m.idx.lin]);
        const area = toNum(r[m.idx.area]);
        const piece = toNum(r[m.idx.piece]);
        const sl = toNum(r[m.idx.sl]);
        const sw = toNum(r[m.idx.sw]);
        const sheetWt = unit === "kg/m²" && area > 0 && sl > 0 && sw > 0 ? (area * sl * sw) / 1e6 : null;
        const sizeLabel = unit === "kg/m²" && sl > 0 && sw > 0 ? `${sl} x ${sw}` : "";
        const entry = { name, unit, lin, area, piece, sl, sw, sheetWt, sizeLabel };
        map.set(lc(name), entry);
        list.push(entry);
      }
    } catch { /* ignore if Master absent */ }
    return { map, list };
  }

  /* Apply clean print layout ONLY on flattened/compiled sharing sheets (never on live editing sheets) */
  function applySharePrintLayout(ws, lastCol, lastRow) {
    try {
      ws.pageLayout.orientation = Excel.PageOrientation.landscape;
      ws.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0 };
      ws.pageLayout.setPrintArea(`A1:${lastCol}${lastRow}`);
      ws.pageLayout.setPrintTitleRows("3:4");
    } catch { /* non-fatal on older Excel builds */ }
  }

  function applyBorders(range, style = "Continuous", color = "#CBD5E1") {
    try {
      ["EdgeLeft", "EdgeRight", "EdgeTop", "EdgeBottom", "InsideHorizontal", "InsideVertical"].forEach((edge) => {
        const b = range.format.borders.getItem(edge);
        b.style = style;
        b.color = color;
      });
    } catch { /* ignore */ }
  }

  function applyBoxBorders(range, color = "#64748B") {
    try {
      ["EdgeLeft", "EdgeRight", "EdgeTop", "EdgeBottom"].forEach((edge) => {
        const b = range.format.borders.getItem(edge);
        b.style = "Continuous";
        b.color = color;
      });
    } catch { /* ignore */ }
  }

  /* Build BOM row formula array for flattened sheet (rows 5..bomEnd + total row bomTot) */
  function buildFlatBomGrid(rows, bomEnd, masterMap) {
    const grid = [];
    for (let i = 0; i < rows.length; i++) {
      const r = 5 + i;
      const item = rows[i];
      const matName = String(item.mat || "").trim();
      const meta = matName ? masterMap.get(lc(matName)) : null;

      const snF = `=IF(C${r}="","",MAX(A$4:A${r - 1})+1)`;
      let unitWtCell = "";
      let totalWtF = "";
      let totalLenF = "";

      if (meta) {
        if (meta.unit === "kg/m" && meta.lin > 0) {
          unitWtCell = meta.lin;
          totalWtF = `=IF(C${r}="","",IF(ISNUMBER(I${r}),N(G${r})*IF(H${r}="",1,H${r})*N(D${r})/1000*I${r},0))`;
          totalLenF = `=IF(AND(C${r}<>"",ISNUMBER(D${r}),ISNUMBER(I${r})),D${r}*N(G${r})*IF(H${r}="",1,H${r}),"")`;
        } else if (meta.unit === "kg/m²" && meta.area > 0) {
          unitWtCell = "";
          totalWtF = `=IF(C${r}="","",N(G${r})*IF(H${r}="",1,H${r})*N(D${r})*N(E${r})/1000000*${meta.area})`;
          totalLenF = "";
        } else if (meta.unit === "kg/each" && meta.piece > 0) {
          unitWtCell = "";
          totalWtF = `=IF(C${r}="","",N(G${r})*IF(H${r}="",1,H${r})*${meta.piece})`;
          totalLenF = "";
        }
      } else if (matName) {
        // Unlisted / custom material — preserve cached weight so nothing is lost
        unitWtCell = item.cachedUnitWt ?? "";
        totalWtF = isBlank(item.cachedTotalWt)
          ? `=IF(C${r}="","",IF(ISNUMBER(I${r}),N(G${r})*IF(H${r}="",1,H${r})*IF(ISNUMBER(D${r}),D${r}/1000,1)*I${r},""))`
          : item.cachedTotalWt;
        totalLenF = isBlank(item.cachedTotalLen) ? "" : item.cachedTotalLen;
      } else {
        // Blank editable row for recipients to add extra items in auto-growing mode
        unitWtCell = "";
        totalWtF = `=IF(C${r}="","",IF(ISNUMBER(I${r}),N(G${r})*IF(H${r}="",1,H${r})*IF(ISNUMBER(D${r}),D${r}/1000,1)*I${r},""))`;
        totalLenF = `=IF(AND(C${r}<>"",ISNUMBER(D${r}),ISNUMBER(I${r})),D${r}*N(G${r})*IF(H${r}="",1,H${r}),"")`;
      }

      grid.push([
        snF,
        item.desc ?? "",
        matName,
        item.len ?? "",
        item.wid ?? "",
        item.thk ?? "",
        item.qty ?? "",
        item.set ?? "",
        unitWtCell,
        totalWtF,
        item.rem ?? "",
        totalLenF
      ]);
    }

    // Total row at bomEnd + 1
    grid.push([
      "", "", "", "", "", "", "", "",
      "Total:",
      `=SUM(J5:J${bomEnd})`,
      "", ""
    ]);
    return grid;
  }

  /* Build helper lookup rows (AC5:AF<n>) for non-linear materials so auto-growing PL works without Master */
  function buildHiddenLookup(masterList) {
    const out = [];
    for (const m of masterList) {
      if (m.unit === "kg/m²" && m.sheetWt > 0) {
        out.push([m.name, "kg/m²", m.sheetWt, m.sizeLabel]);
      } else if (m.unit === "kg/each" && m.piece > 0) {
        out.push([m.name, "kg/each", m.piece, ""]);
      } else if (m.unit === "kg/m" && m.lin > 0) {
        out.push([m.name, "kg/m", m.lin, ""]);
      }
    }
    return out;
  }

  /* Build Purchase List grid (O5:AA<plSig>) for Excel 365 auto-growing or fixed print-ready mode */
  function buildFlatPlGrid(opts) {
    const { mode, bomEnd, uniqueMats, plInputs, masterMap, lookupEnd } = opts;
    const grid = [];

    if (mode === "fixed") {
      const count = Math.max(1, uniqueMats.length);
      for (let i = 0; i < count; i++) {
        const r = 5 + i;
        const name = uniqueMats[i] || "";
        const meta = name ? masterMap.get(lc(name)) : null;
        const inp = name ? (plInputs.get(lc(name)) || {}) : {};
        if (!name) {
          grid.push(["", "", "", "", "", "", "", "", "", "", "", "", ""]);
          continue;
        }
        let lenCell = "";
        let qtyF = "";
        let wtCell = "";
        let totWtF = `=S${r}`;

        if (meta && meta.unit === "kg/m" && meta.lin > 0) {
          lenCell = "=$AE$1";
          qtyF = `=IFERROR(ROUNDUP(ROUND(SUMIF($C$5:$C$${bomEnd},P${r},$L$5:$L$${bomEnd})/Q${r},6),0),"")`;
          wtCell = meta.lin;
          totWtF = `=IFERROR(Q${r}/1000*R${r}*T${r},S${r})`;
        } else if (meta && meta.unit === "kg/m²" && meta.sheetWt > 0) {
          lenCell = meta.sizeLabel;
          qtyF = `=IFERROR(ROUNDUP(ROUND(S${r}/${meta.sheetWt},6),0),"")`;
          totWtF = `=IFERROR(R${r}*${meta.sheetWt},S${r})`;
        } else if (meta && meta.unit === "kg/each" && meta.piece > 0) {
          qtyF = `=IFERROR(ROUNDUP(ROUND(S${r}/${meta.piece},6),0),"")`;
          totWtF = `=IFERROR(R${r}*${meta.piece},S${r})`;
        } else {
          qtyF = `=IFERROR(ROUNDUP(SUMIF($C$5:$C$${bomEnd},P${r},$G$5:$G$${bomEnd}),0),"")`;
        }

        grid.push([
          i + 1,
          name,
          lenCell,
          qtyF,
          `=SUMIF($C$5:$C$${bomEnd},P${r},$J$5:$J$${bomEnd})`,
          wtCell,
          totWtF,
          inp.stock ?? "",
          inp.poWt ?? "",
          inp.poQty ?? "",
          inp.poRate ?? "",
          `=IF(Y${r}="","",W${r}*Y${r})`,
          inp.rem ?? ""
        ]);
      }

      const plTot = 5 + count;
      grid.push([
        "", "Grand Total Weight in KG :", "", "",
        `=SUM(S5:S${plTot - 1})`, "",
        `=SUM(U5:U${plTot - 1})`, "",
        `=IF(COUNT(W5:W${plTot - 1})=0,"",SUM(W5:W${plTot - 1}))`,
        "", "",
        `=IF(COUNT(Z5:Z${plTot - 1})=0,"",SUM(Z5:Z${plTot - 1}))`,
        ""
      ]);
      grid.push(["", "", "", "", "", "", "", "", "", "", "", "", ""]);
      grid.push(["", "Requested by:", "Procurement", "", "", "Store Section:", "", "", "Production", "", "", "Engineering", ""]);
      return { grid, plDataEnd: plTot - 1, plTot, plSig: plTot + 2, cap: count };
    }

    // Pure Excel 365 auto-growing mode (capacity >= 40, LET + XLOOKUP, zero helper col M)
    const cap = Math.max(PL_DEFAULT_CAP, uniqueMats.length + 10);
    const hasLookup = lookupEnd >= 5;
    const lkMat = `$AC$5:$AC$${lookupEnd}`;
    const lkUnit = `$AD$5:$AD$${lookupEnd}`;
    const lkWt = `$AE$5:$AE$${lookupEnd}`;
    const lkSz = `$AF$5:$AF$${lookupEnd}`;
    const bomMat = `$C$5:$C$${bomEnd}`;
    const bomWt = `$J$5:$J$${bomEnd}`;
    const bomLen = `$L$5:$L$${bomEnd}`;
    const bomUnitWt = `$I$5:$I$${bomEnd}`;
    const bomQty = `$G$5:$G$${bomEnd}`;

    // Map existing PL inputs to sorted positions so Stock/PO stay aligned with their material
    const orderedNames = [...uniqueMats].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));

    for (let i = 0; i < cap; i++) {
      const r = 5 + i;
      const matAtRow = orderedNames[i] || "";
      const inp = matAtRow ? (plInputs.get(lc(matAtRow)) || {}) : {};

      const snF = `=IF(P${r}="","",N(O${r - 1})+1)`;
      const matF = `=IF(ROW()-4>$AE$2,"",INDEX($AB$5#,ROW()-4))`;

      const lenF = hasLookup
        ? `=IF(P${r}="","",LET(u,XLOOKUP(P${r},${lkMat},${lkUnit},""),IF(u="kg/m",$AE$1,IF(u="kg/m²",XLOOKUP(P${r},${lkMat},${lkSz},""),IF(ISNUMBER(XLOOKUP(P${r},${bomMat},${bomUnitWt},"")),$AE$1,"")))))`
        : `=IF(P${r}="","",IF(ISNUMBER(XLOOKUP(P${r},${bomMat},${bomUnitWt},"")),$AE$1,""))`;
      const consWtF = `=IF(P${r}="","",SUMIF(${bomMat},P${r},${bomWt}))`;
      const linWtF = hasLookup
        ? `=IF(P${r}="","",LET(u,XLOOKUP(P${r},${lkMat},${lkUnit},""),IF(u="kg/m",XLOOKUP(P${r},${lkMat},${lkWt},""),IF(u="",LET(b,XLOOKUP(P${r},${bomMat},${bomUnitWt},""),IF(ISNUMBER(b),b,"")),""))))`
        : `=IF(P${r}="","",LET(b,XLOOKUP(P${r},${bomMat},${bomUnitWt},""),IF(ISNUMBER(b),b,"")))`;
      const qtyF = hasLookup
        ? `=IF(P${r}="","",IF(ISNUMBER(Q${r}),IFERROR(ROUNDUP(ROUND(SUMIF(${bomMat},P${r},${bomLen})/Q${r},6),0),""),LET(w,XLOOKUP(P${r},${lkMat},${lkWt},""),IF(ISNUMBER(w),IFERROR(ROUNDUP(ROUND(S${r}/w,6),0),""),IFERROR(ROUNDUP(SUMIF(${bomMat},P${r},${bomQty}),0),"")))))`
        : `=IF(P${r}="","",IF(ISNUMBER(Q${r}),IFERROR(ROUNDUP(ROUND(SUMIF(${bomMat},P${r},${bomLen})/Q${r},6),0),""),IFERROR(ROUNDUP(SUMIF(${bomMat},P${r},${bomQty}),0),"")))`;
      const totWtF = hasLookup
        ? `=IF(P${r}="","",IF(AND(ISNUMBER(Q${r}),ISNUMBER(T${r})),IFERROR(Q${r}/1000*R${r}*T${r},S${r}),LET(w,XLOOKUP(P${r},${lkMat},${lkWt},""),IF(ISNUMBER(w),IFERROR(R${r}*w,S${r}),S${r}))))`
        : `=IF(P${r}="","",IF(AND(ISNUMBER(Q${r}),ISNUMBER(T${r})),IFERROR(Q${r}/1000*R${r}*T${r},S${r}),S${r}))`;
      const amtF = `=IF(OR(P${r}="",Y${r}=""),"",W${r}*Y${r})`;

      grid.push([
        snF, matF, lenF, qtyF, consWtF, linWtF, totWtF,
        inp.stock ?? "", inp.poWt ?? "", inp.poQty ?? "", inp.poRate ?? "",
        amtF, inp.rem ?? ""
      ]);
    }

    const plTot = 5 + cap;
    grid.push([
      "", "Grand Total Weight in KG :", "", "",
      `=SUM(S5:S${plTot - 1})`, "",
      `=SUM(U5:U${plTot - 1})`, "",
      `=IF(COUNT(W5:W${plTot - 1})=0,"",SUM(W5:W${plTot - 1}))`,
      "", "",
      `=IF(COUNT(Z5:Z${plTot - 1})=0,"",SUM(Z5:Z${plTot - 1}))`,
      ""
    ]);
    grid.push(["", "", "", "", "", "", "", "", "", "", "", "", ""]);
    grid.push(["", "Requested by:", "Procurement", "", "", "Store Section:", "", "", "Production", "", "", "Engineering", ""]);
    return { grid, plDataEnd: plTot - 1, plTot, plSig: plTot + 2, cap };
  }

  /* ---------------- 1. FLATTEN WORKBOOK (Selective Sheets + Excel 365 / Fixed + Compiled Sheet) ---------------- */
  async function flattenWorkbook(opts = {}, log = noop) {
    const mode = opts.mode === "fixed" ? "fixed" : "dynamic365"; // 'dynamic365' | 'fixed'
    const makeCompiled = Boolean(opts.compiled);
    const wanted = Array.isArray(opts.sheets) && opts.sheets.length ? new Set(opts.sheets.map(lc)) : null;

    let result = { converted: 0, isPartial: false, compiled: false, mode };

    await window.BomSheets.runExcel(async (context) => {
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
        const projectMeta = await window.BomSheets.readProjectProp(context, wb);
        const { map: masterMap, list: masterList } = await loadMasterMap(context);

        wb.worksheets.load("items/name,items/visibility");
        await context.sync();
        const allWs = wb.worksheets.items;
        allWs.forEach((ws) => ws.tables.load("items/name"));
        await context.sync();

        const visibleBomWs = allWs.filter(
          (ws) => ws.visibility === "Visible" && ws.name !== window.BomSheets.BLANK && Boolean(firstTable(ws.tables.items, "bom_"))
        );
        if (!visibleBomWs.length) throw new Error("No visible automated BOM sheets were found.");

        const targetWs = wanted ? visibleBomWs.filter((ws) => wanted.has(lc(ws.name))) : visibleBomWs;
        if (!targetWs.length) throw new Error("Select at least one BOM sheet to convert.");

        const isPartial = targetWs.length < visibleBomWs.length;
        result.isPartial = isPartial;
        log(
          isPartial
            ? `Flattening ${targetWs.length} of ${visibleBomWs.length} sheets (Master sheet & named ranges will be kept)...`
            : `Flattening all ${targetWs.length} BOM sheet(s) in ${mode} mode...`
        );

        // Batch-read all selected sheets
        const snapshots = targetWs.map((ws) => {
          const bomT = firstTable(ws.tables.items, "bom_");
          const plT = firstTable(ws.tables.items, "pl_");
          return {
            ws,
            name: ws.name,
            bomT,
            plT,
            bomBody: bomT.getDataBodyRange().load("values"),
            plBody: plT ? plT.getDataBodyRange().load("values") : null,
            rA1: ws.getRange("A1:A2").load("values"),
            rO2: ws.getRange("O2").load("values"),
            rAE1: ws.getRange("AE1").load("values")
          };
        });
        await context.sync();

        const lookupRows = buildHiddenLookup(masterList);
        const compiledSummary = [];
        const allUniqueAcross = new Map(); // lc(name) -> original name

        for (const snap of snapshots) {
          const { ws, bomT, plT } = snap;
          log(`Converting "${snap.name}"...`);

          const stdLen = toNum(snap.rAE1.values[0][0]) || 6000;
          const o2Title = String(snap.rO2.values[0][0] || "").trim();

          // Extract non-empty BOM rows up to last used row
          const rawBom = snap.bomBody.values;
          let lastUsed = -1;
          for (let i = rawBom.length - 1; i >= 0; i--) {
            const r = rawBom[i];
            if (!isBlank(r[1]) || !isBlank(r[2]) || !isBlank(r[3]) || !isBlank(r[6])) {
              lastUsed = i;
              break;
            }
          }
          const bomItems = [];
          const sheetUnique = [];
          const seenMat = new Set();
          for (let i = 0; i <= lastUsed; i++) {
            const r = rawBom[i];
            const mat = String(r[2] || "").trim();
            if (mat && !seenMat.has(lc(mat))) {
              seenMat.add(lc(mat));
              sheetUnique.push(mat);
              if (!allUniqueAcross.has(lc(mat))) allUniqueAcross.set(lc(mat), mat);
            }
            bomItems.push({
              desc: r[1],
              mat,
              len: r[3],
              wid: r[4],
              thk: r[5],
              qty: r[6],
              set: r[7],
              cachedUnitWt: r[8],
              cachedTotalWt: r[9],
              rem: r[10],
              cachedTotalLen: r[11]
            });
          }
          const extraRows = mode === "fixed" ? 0 : EXTRA_EDIT_ROWS;
          for (let k = 0; k < extraRows; k++) bomItems.push({ mat: "" });
          if (bomItems.length === 0) bomItems.push({ mat: "" });

          // Extract existing PL user inputs keyed by material name
          const plInputs = new Map();
          if (snap.plBody) {
            for (const r of snap.plBody.values) {
              const name = String(r[1] || "").trim();
              if (!name || lc(name).startsWith("grand total") || lc(name).startsWith("requested by")) continue;
              if (!plInputs.has(lc(name))) {
                plInputs.set(lc(name), {
                  stock: r[7],
                  poWt: r[8],
                  poQty: r[9],
                  poRate: r[10],
                  rem: r[12]
                });
              }
            }
          }

          window.BomSheets.pauseExcel(context);

          // Convert Excel Tables to plain ranges & clear old CF/DV/helpers
          try { bomT.convertToRange(); } catch { /* ignore */ }
          try { if (plT) plT.convertToRange(); } catch { /* ignore */ }
          ws.getRange().conditionalFormats.clearAll();
          try { ws.getRange("C5:C500").dataValidation.clear(); } catch { /* ignore */ }
          ws.getRange("5:5").rowHidden = false;
          ws.getRange("5:5").delete(Excel.DeleteShiftDirection.up);

          // Clear stale cells below row 5 across A:AL (including AL6 dynamic spill helper)
          const clearEnd = Math.max(120, rawBom.length + 25);
          const wipe = ws.getRange(`A5:AL${clearEnd}`);
          wipe.clear(Excel.ClearApplyTo.contents);
          wipe.format.fill.clear();
          wipe.format.borders.getItem("EdgeTop").style = "None";
          wipe.format.borders.getItem("EdgeBottom").style = "None";
          wipe.format.borders.getItem("EdgeLeft").style = "None";
          wipe.format.borders.getItem("EdgeRight").style = "None";
          wipe.format.borders.getItem("InsideHorizontal").style = "None";
          wipe.format.borders.getItem("InsideVertical").style = "None";

          // Clear old helper cells M4, AB4, AD2
          ws.getRange("M4").clear(Excel.ClearApplyTo.contents);
          ws.getRange("AB4").clear(Excel.ClearApplyTo.contents);
          ws.getRange("AD2").clear(Excel.ClearApplyTo.all);
          ws.getRange("AE1").values = [[stdLen]];
          if (o2Title) ws.getRange("O2").values = [[o2Title]];

          // Write hidden compact lookup table in AC5:AF<lookupEnd> for auto-growing mode
          let lookupEnd = 4;
          if (mode !== "fixed" && lookupRows.length) {
            lookupEnd = 4 + lookupRows.length;
            ws.getRange(`AC5:AF${lookupEnd}`).values = lookupRows;
          }

          // Write BOM section A5:L<bomTot>
          const bomEnd = 4 + bomItems.length;
          const bomTot = bomEnd + 1;
          const bomGrid = buildFlatBomGrid(bomItems, bomEnd, masterMap);
          ws.getRange(`A5:L${bomTot}`).formulas = bomGrid;

          // Write Purchase List O5:AA<plSig>
          const plOut = buildFlatPlGrid({
            mode,
            bomEnd,
            uniqueMats: sheetUnique,
            plInputs,
            masterMap,
            lookupEnd
          });

          // Write Excel 365 dynamic spill formula via .formulas
          if (mode === "dynamic365") {
            ws.getRange("AB5").formulas = [
              [`=LET(m,FILTER($C$5:$C$${bomEnd},($C$5:$C$${bomEnd}<>"")*($C$5:$C$${bomEnd}<>"Materials"),""),IF(INDEX(m,1)="","",SORT(UNIQUE(m))))`]
            ];
            ws.getRange("AE2").formulas = [[`=IF($AB$5="",0,ROWS($AB$5#))`]];
          } else {
            ws.getRange("AE2").values = [[sheetUnique.length]];
          }

          ws.getRange(`O5:AA${plOut.plSig}`).formulas = plOut.grid;

          // Clean, lightweight formatting & number formats
          const bomRange = ws.getRange(`A5:L${bomEnd}`);
          bomRange.format.font.name = "Calibri";
          bomRange.format.font.size = 10;
          bomRange.format.font.bold = false;
          applyBorders(bomRange);

          ws.getRange(`I5:J${bomTot}`).numberFormat = [["#,##0.00"]];
          ws.getRange(`D5:F${bomEnd}`).numberFormat = [["#,##0"]];
          ws.getRange(`L5:L${bomEnd}`).numberFormat = [["#,##0"]];

          const bomTotRng = ws.getRange(`A${bomTot}:L${bomTot}`);
          bomTotRng.format.font.bold = true;
          bomTotRng.format.fill.color = "#F4F6F5";
          applyBoxBorders(bomTotRng);

          // Section-heading highlight (lean single CF rule on active BOM rows only)
          const cfHead = bomRange.conditionalFormats.add("Custom");
          cfHead.custom.rule.formula = '=AND($C5="",$B5<>"")';
          cfHead.custom.format.font.bold = true;
          cfHead.custom.format.fill.color = "#F1F5F9";

          const plDataRng = ws.getRange(`O5:AA${plOut.plDataEnd}`);
          plDataRng.format.font.name = "Calibri";
          plDataRng.format.font.size = 10;
          plDataRng.format.font.bold = false;
          if (mode === "fixed") {
            applyBorders(plDataRng);
          } else {
            // Only draw borders on filled PL rows so empty capacity rows stay clean
            const cfPlRow = plDataRng.conditionalFormats.add("Custom");
            cfPlRow.custom.rule.formula = '=$P5<>""';
            ["EdgeLeft", "EdgeRight", "EdgeTop", "EdgeBottom"].forEach((e) => {
              cfPlRow.custom.format.borders.getItem(e).style = "Continuous";
            });
          }

          ws.getRange(`S5:U${plOut.plTot}`).numberFormat = [["#,##0.00"]];
          ws.getRange(`W5:W${plOut.plTot}`).numberFormat = [["#,##0.00"]];
          ws.getRange(`Y5:Z${plOut.plTot}`).numberFormat = [["#,##0.00"]];

          const plTotRng = ws.getRange(`O${plOut.plTot}:AA${plOut.plTot}`);
          plTotRng.format.font.bold = true;
          plTotRng.format.fill.color = "#F4F6F5";
          applyBoxBorders(plTotRng);

          const plSigRng = ws.getRange(`O${plOut.plSig}:AA${plOut.plSig}`);
          plSigRng.format.font.bold = true;
          plSigRng.format.font.size = 9.5;
          plSigRng.format.font.color = "#475569";

          // Hide helper columns M & AB:AF; apply print layout ONLY for print-ready Fixed mode
          ws.getRange("M:M").columnHidden = true;
          ws.getRange("AB:AF").columnHidden = true;
          ws.getRange("AL:AL").columnHidden = true;
          if (mode === "fixed") {
            applySharePrintLayout(ws, "AA", Math.max(bomTot, plOut.plSig));
          }

          await context.sync();
          compiledSummary.push({
            name: snap.name,
            plDataEnd: plOut.plDataEnd,
            bomEnd
          });
          result.converted++;
        }

        // Optional Compiled Purchase List across all converted sheets
        if (makeCompiled && compiledSummary.length) {
          log("Building Compiled Purchase List...");
          const existingComp = wb.worksheets.getItemOrNullObject(COMPILED);
          await context.sync();
          if (!existingComp.isNullObject) {
            existingComp.delete();
            await context.sync();
          }
          const compWs = wb.worksheets.add(COMPILED);
          compWs.position = 0;
          const mats = [...allUniqueAcross.values()].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "base" }));
          const sheetCount = compiledSummary.length;
          const totalCols = 8 + sheetCount; // A..H + per-sheet weight cols

          const colLetter = (idx) => {
            let s = "";
            let n = idx + 1;
            while (n > 0) {
              const rem = (n - 1) % 26;
              s = String.fromCharCode(65 + rem) + s;
              n = Math.floor((n - 1) / 26);
            }
            return s;
          };
          const lastCol = colLetter(totalCols - 1);

          window.BomSheets.pauseExcel(context);
          compWs.getRange("A1").values = [["Compiled Purchase List"]];
          compWs.getRange("A1").format.font.bold = true;
          compWs.getRange("A1").format.font.size = 14;
          compWs.getRange("A2").values = [[`Combined across ${sheetCount} BOM sheet(s): ${compiledSummary.map((s) => s.name).join(", ")}`]];
          compWs.getRange("A2").format.font.size = 10;
          compWs.getRange("A2").format.font.color = "#475569";

          const headerRow = [
            "S.N.", "Materials", "Primary Unit", "Length / Size",
            "Consumable Wt. (kg)", "Weight kg/m", "Required Qty (Nos)", "Total Order Wt. (kg)",
            ...compiledSummary.map((s) => `${s.name} (kg)`)
          ];
          const hdrRng = compWs.getRange(`A4:${lastCol}4`);
          hdrRng.values = [headerRow];
          hdrRng.format.font.bold = true;
          hdrRng.format.font.size = 10;
          hdrRng.format.fill.color = "#1D6F42";
          hdrRng.format.font.color = "#FFFFFF";

          const dataRows = [];
          const firstSheetCol = colLetter(8);
          const lastSheetCol = colLetter(8 + sheetCount - 1);

          mats.forEach((matName, idx) => {
            const r = 5 + idx;
            const meta = masterMap.get(lc(matName));
            const perSheet = compiledSummary.map(
              (s) => `=SUMIF(${qSheet(s.name)}!$P$5:$P$${s.plDataEnd},$B${r},${qSheet(s.name)}!$S$5:$S$${s.plDataEnd})`
            );
            const consF = `=SUM(${firstSheetCol}${r}:${lastSheetCol}${r})`;
            let unit = meta ? meta.unit : "";
            let sizeCell = "";
            let wtCol = "";
            let reqQtyF = "";
            let ordWtF = `=E${r}`;

            if (meta && meta.unit === "kg/m" && meta.lin > 0) {
              sizeCell = 6000;
              wtCol = meta.lin;
              reqQtyF = `=IFERROR(ROUNDUP(ROUND((E${r}/F${r})/(D${r}/1000),6),0),"")`;
              ordWtF = `=IFERROR(G${r}*(D${r}/1000)*F${r},E${r})`;
            } else if (meta && meta.unit === "kg/m²" && meta.sheetWt > 0) {
              sizeCell = meta.sizeLabel;
              reqQtyF = `=IFERROR(ROUNDUP(ROUND(E${r}/${meta.sheetWt},6),0),"")`;
              ordWtF = `=IFERROR(G${r}*${meta.sheetWt},E${r})`;
            } else if (meta && meta.unit === "kg/each" && meta.piece > 0) {
              reqQtyF = `=IFERROR(ROUNDUP(ROUND(E${r}/${meta.piece},6),0),"")`;
              ordWtF = `=IFERROR(G${r}*${meta.piece},E${r})`;
            }

            dataRows.push([idx + 1, matName, unit, sizeCell, consF, wtCol, reqQtyF, ordWtF, ...perSheet]);
          });

          const lastDataRow = Math.max(5, 4 + dataRows.length);
          const totRow = lastDataRow + 1;
          if (dataRows.length) {
            compWs.getRange(`A5:${lastCol}${lastDataRow}`).formulas = dataRows;
          }

          const totArr = [
            "", "Grand Total", "", "",
            `=SUM(E5:E${lastDataRow})`, "", "", `=SUM(H5:H${lastDataRow})`,
            ...compiledSummary.map((_, idx) => {
              const c = colLetter(8 + idx);
              return `=SUM(${c}5:${c}${lastDataRow})`;
            })
          ];
          const totRng = compWs.getRange(`A${totRow}:${lastCol}${totRow}`);
          totRng.formulas = [totArr];
          totRng.format.font.bold = true;
          totRng.format.fill.color = "#F4F6F5";

          applyBorders(compWs.getRange(`A4:${lastCol}${totRow}`));
          compWs.getRange(`E5:F${totRow}`).numberFormat = [["#,##0.00"]];
          compWs.getRange(`H5:${lastCol}${totRow}`).numberFormat = [["#,##0.00"]];
          compWs.showGridlines = false;
          try {
            compWs.getRange(`A4:${lastCol}${totRow}`).format.autofitColumns();
            compWs.getRange("B:B").format.columnWidth = 175;
          } catch { /* ignore */ }
          applySharePrintLayout(compWs, lastCol, totRow);
          compWs.activate();
          await context.sync();
          result.compiled = true;
        }

        // Only remove Master, hidden BOM Blank, and defined names if ALL automated BOM sheets were flattened!
        if (!isPartial) {
          if (projectMeta) {
            // Re-stamp project link custom property so future Restore remembers the Project-Only list
            await window.BomSheets.writeProjectProp(context, wb, projectMeta);
          }
          log("Removing Master Unique Materials, hidden BOM Blank & named ranges...");
          wb.names.load("items/name");
          await context.sync();
          wb.names.items.forEach((nm) => {
            if (!lc(nm.name).includes("print_area") && !lc(nm.name).includes("print_titles")) {
              try { nm.delete(); } catch { /* ignore */ }
            }
          });
          const toDrop = [window.BomSheets.MASTER, window.BomSheets.BLANK, "Formulas", "README"];
          for (const sName of toDrop) {
            const ws = wb.worksheets.getItemOrNullObject(sName);
            await context.sync();
            if (!ws.isNullObject) {
              try { ws.delete(); await context.sync(); } catch { /* ignore */ }
            }
          }
        }
      } finally {
        try {
          if (prevCalc !== null) wb.application.calculationMode = prevCalc;
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        } catch { /* ignore */ }
      }
    });

    log(`Conversion complete (${result.converted} sheet(s)).`);
    return result;
  }

  /* ---------------- 2. RESTORE AUTOMATED BOM (3-Way Material Source + Selective Sheets) ---------------- */
  async function restoreWorkbook(opts = {}, log = noop) {
    const selected = Array.isArray(opts.sheets) ? opts.sheets : [];
    if (!selected.length) throw new Error("Select at least one converted sheet to restore.");

    // Resolve chosen material source (global / project-saved / project-file)
    let chosenList = null;
    let projectMeta = null;
    if (opts.materialSource === "project-file" && opts.projectData) {
      chosenList = opts.projectData.materials || [];
      const saved = await window.BomStore.saveProjectList(
        { id: opts.projectData.id, projectName: opts.projectData.projectName || "Imported Project" },
        chosenList,
        opts.projectData.lastChange || null
      );
      projectMeta = { id: saved.id, projectName: saved.projectName, updated: saved.updated };
    } else if (opts.materialSource === "project-saved" && opts.projectId) {
      const found = await window.BomStore.getSavedProject(opts.projectId);
      if (!found) throw new Error("Selected saved project list was not found.");
      chosenList = found.materials || [];
      projectMeta = { id: found.id, projectName: found.projectName, updated: found.updated };
    } else {
      const globalData = await window.BomStore.readGlobalMaster();
      if (globalData.materials.length) chosenList = globalData.materials;
      projectMeta = null;
    }

    // Ensure Master & hidden BOM Blank exist
    const info = await window.BomSheets.inspectWorkbook();
    if (!info.hasMaster || !info.hasBlank) {
      log("Importing lightweight Master & BOM Blank template...");
      await window.BomSheets.importTemplate(log);
    }

    // Apply chosen materials list to Master
    if (chosenList && chosenList.length) {
      log(projectMeta ? `Applying Project-Only list "${projectMeta.projectName}" (${chosenList.length} items)...` : `Applying Global Master (${chosenList.length} items)...`);
      await window.BomSheets.applyMaterialList(chosenList, {
        replaceAll: true,
        projectMeta
      });
    } else if (projectMeta === null) {
      await window.BomSheets.runExcel(async (context) => {
        await window.BomSheets.writeProjectProp(context, context.workbook, null);
      });
    }

    let restored = 0;
    let firstRestored = null;

    await window.BomSheets.runExcel(async (context) => {
      const wb = context.workbook;
      let prevCalc = null;
      let unhidBlank = false;
      let blankWs = null;

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
        const matInfo = await window.BomSheets.ensureMasterNames(context, wb, log);
        blankWs = wb.worksheets.getItem(window.BomSheets.BLANK);
        blankWs.load("visibility");
        await context.sync();
        if (blankWs.visibility !== "Visible") {
          blankWs.visibility = Excel.SheetVisibility.visible;
          unhidBlank = true;
          await context.sync();
        }

        const restoredSheets = [];
        for (const sheetName of selected) {
          const oldWs = wb.worksheets.getItemOrNullObject(sheetName);
          await context.sync();
          if (oldWs.isNullObject) continue;

          log(`Restoring "${sheetName}"...`);
          const used = oldWs.getUsedRangeOrNullObject(true);
          used.load("rowCount");
          const headRng = oldWs.getRange("A1:AE6").load("values");
          await context.sync();
          const maxRow = used.isNullObject ? 80 : Math.min(600, Math.max(60, used.rowCount + 5));

          const bomScan = oldWs.getRange(`A5:K${maxRow}`).load("values");
          const plScan = oldWs.getRange(`O5:AA${Math.min(120, maxRow + 45)}`).load("values");
          await context.sync();

          const titleA1 = String(headRng.values[0]?.[0] || "").trim();
          const titleA2 = String(headRng.values[1]?.[0] || "").trim() || sheetName;
          const stdLen = toNum(headRng.values[0]?.[30]) || 6000;

          // Detect whether data starts at row 5 (v3.0 plain) or row 6 (legacy plain)
          const r5b = lc(bomScan.values[0]?.[1]);
          const r5c = lc(bomScan.values[0]?.[2]);
          const startIdx = (r5b === "description" || r5c === "materials") ? 1 : 0;

          const bomRows = [];
          for (let i = startIdx; i < bomScan.values.length; i++) {
            const r = bomScan.values[i];
            if (lc(r[8]).startsWith("total") || lc(r[1]) === "total" || lc(r[1]) === "total:") break;
            if (!isBlank(r[1]) || !isBlank(r[2]) || !isBlank(r[3]) || !isBlank(r[6])) {
              bomRows.push({
                desc: r[1], mat: String(r[2] || "").trim(),
                len: r[3], wid: r[4], thk: r[5], qty: r[6], set: r[7], rem: r[10]
              });
            }
          }

          // Collect PL user inputs keyed by material name
          const plInputs = new Map();
          for (let i = startIdx; i < plScan.values.length; i++) {
            const r = plScan.values[i];
            const mName = String(r[1] || "").trim();
            if (!mName || lc(mName).startsWith("grand total") || lc(mName).startsWith("requested by")) continue;
            if (!plInputs.has(lc(mName))) {
              plInputs.set(lc(mName), {
                stock: r[7], poWt: r[8], poQty: r[9], poRate: r[10], rem: r[12]
              });
            }
          }

          // Rename old sheet temporarily and clone clean BOM Blank in its exact position
          const tmpName = window.BomSheets.uniqueName(`__old_${window.BomSheets.safeId(sheetName).slice(0, 18)}`, [sheetName]);
          oldWs.name = tmpName;
          const copy = blankWs.copy(Excel.WorksheetPositionType.before, oldWs);
          copy.name = sheetName;
          copy.visibility = Excel.SheetVisibility.visible;
          copy.tables.load("items/name");
          wb.tables.load("items/name");
          await context.sync();

          const bomT = firstTable(copy.tables.items, "bom_");
          const plT = firstTable(copy.tables.items, "pl_");
          const origCopiedBomName = bomT ? bomT.name : "";
          const origCopiedPlName = plT ? plT.name : "";
          const taken = wb.tables.items.map((t) => t.name);
          const id = window.BomSheets.safeId(sheetName);
          const bomName = window.BomSheets.uniqueName(`bom_${id}`, taken);
          const plName = window.BomSheets.uniqueName(`pl_${id}`, taken.concat([bomName]));

          window.BomSheets.pauseExcel(context);
          bomT.name = bomName;
          if (plT) plT.name = plName;
          if (titleA1) copy.getRange("A1").values = [[titleA1]];
          copy.getRange("A2").values = [[titleA2]];
          copy.getRange("AE1").values = [[stdLen]];
          await context.sync();

          // Expand bom_ table ONLY if restored BOM exceeds the lean 60-row capacity
          const bomBody = bomT.getDataBodyRange().load("rowCount");
          await context.sync();
          if (bomRows.length > bomBody.rowCount) {
            const addN = bomRows.length - bomBody.rowCount + 5;
            for (let k = 0; k < addN; k++) bomT.rows.add();
            await context.sync();
          }

          // Write BOM inputs in 2 fast 2D blocks (B..H and K)
          if (bomRows.length) {
            const endR = 5 + bomRows.length;
            const bh = bomRows.map((x) => [x.desc, x.mat, x.len, x.wid, x.thk, x.qty, x.set]);
            const remCol = bomRows.map((x) => [x.rem]);
            window.BomSheets.pauseExcel(context);
            copy.getRange(`B6:H${endR}`).values = bh;
            copy.getRange(`K6:K${endR}`).values = remCol;
            await context.sync();
          }

          // Wire native Excel 365 formulas (M4, AL6, AB4, AD2, table refs), active-table DV & delete temp sheet
          await window.BomSheets.repairBomSheet(context, wb, copy, matInfo, {
            srcBomName: "bom_Take_off",
            srcPlName: "pl_Take_off",
            origCopiedBomName,
            origCopiedPlName,
            upgrade365: true,
            ensureValidation: true
          }, log);
          oldWs.delete();
          await context.sync();

          restoredSheets.push({ ws: copy, plT, plInputs });
          if (!firstRestored) firstRestored = copy;
          restored++;
        }

        if (unhidBlank && blankWs) {
          blankWs.visibility = Excel.SheetVisibility.hidden;
          unhidBlank = false;
          await context.sync();
        }

        // Recalculate once so Purchase Lists populate their Materials column, then write back Stock/PO inputs
        wb.application.calculate(Excel.CalculationType.full);
        await context.sync();

        for (const item of restoredSheets) {
          if (!item.plT || !item.plInputs.size) continue;
          const plBody = item.plT.getDataBodyRange().load("values");
          await context.sync();
          let touched = false;
          window.BomSheets.pauseExcel(context);
          plBody.values.forEach((row, idx) => {
            const mName = String(row[1] || "").trim();
            const saved = mName ? item.plInputs.get(lc(mName)) : null;
            if (!saved) return;
            if (!isBlank(saved.stock)) { plBody.getCell(idx, 7).values = [[saved.stock]]; touched = true; }
            if (!isBlank(saved.poWt)) { plBody.getCell(idx, 8).values = [[saved.poWt]]; touched = true; }
            if (!isBlank(saved.poQty)) { plBody.getCell(idx, 9).values = [[saved.poQty]]; touched = true; }
            if (!isBlank(saved.poRate)) { plBody.getCell(idx, 10).values = [[saved.poRate]]; touched = true; }
            if (!isBlank(saved.rem)) { plBody.getCell(idx, 12).values = [[saved.rem]]; touched = true; }
          });
          if (touched) await context.sync();
        }

        // Optionally remove Compiled Purchase List
        if (opts.removeCompiled) {
          const comp = wb.worksheets.getItemOrNullObject(COMPILED);
          await context.sync();
          if (!comp.isNullObject) {
            try { comp.delete(); await context.sync(); } catch { /* ignore */ }
          }
        }

        if (firstRestored) {
          firstRestored.activate();
          firstRestored.getRange("B6").select();
          await context.sync();
        }
      } finally {
        if (unhidBlank && blankWs) {
          try { blankWs.visibility = Excel.SheetVisibility.hidden; } catch { /* ignore */ }
        }
        try {
          if (prevCalc !== null) wb.application.calculationMode = prevCalc;
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        } catch { /* ignore */ }
      }
    });

    log(`Restored ${restored} automated BOM sheet(s).`);
    return { restored };
  }

  /* ---------------- 3. REPAIR / UPGRADE PLAIN-SHEET PURCHASE-LIST NAMES (Excel 365) ---------------- */
  async function repairPlainSheets(opts = {}, log = noop) {
    const selected = Array.isArray(opts.sheets) ? opts.sheets : [];
    if (!selected.length) throw new Error("Select at least one converted plain sheet to repair.");

    let fixedCount = 0;
    await window.BomSheets.runExcel(async (context) => {
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
        for (const name of selected) {
          const ws = wb.worksheets.getItemOrNullObject(name);
          await context.sync();
          if (ws.isNullObject) continue;

          log(`Repairing Purchase-List names on "${name}" (Excel 365)...`);
          const used = ws.getUsedRangeOrNullObject(true);
          used.load("rowCount");
          await context.sync();
          const maxRow = used.isNullObject ? 120 : Math.min(500, Math.max(80, used.rowCount + 5));

          const colCI = ws.getRange(`B5:I${maxRow}`).load("values");
          const colP = ws.getRange(`P5:P${Math.min(120, maxRow + 45)}`).load("values");
          await context.sync();

          // Detect start row (5 vs 6) and BOM end row (just above Total:)
          const startRow = (lc(colCI.values[0]?.[0]) === "description" || lc(colCI.values[0]?.[1]) === "materials") ? 6 : 5;
          let bomEnd = startRow + 20;
          for (let i = startRow - 5; i < colCI.values.length; i++) {
            const b = lc(colCI.values[i]?.[0]);
            const colI = lc(colCI.values[i]?.[7]);
            if (colI.startsWith("total") || b === "total" || b === "total:") {
              bomEnd = 5 + i - 1;
              break;
            }
          }

          // Detect PL total row (above "Grand Total Weight in KG :")
          let plDataEnd = startRow + PL_DEFAULT_CAP - 1;
          for (let i = startRow - 5; i < colP.values.length; i++) {
            if (lc(colP.values[i]?.[0]).startsWith("grand total")) {
              plDataEnd = 5 + i - 1;
              break;
            }
          }
          const plCount = Math.max(1, plDataEnd - startRow + 1);
          const bomMat = `$C$${startRow}:$C$${bomEnd}`;

          window.BomSheets.pauseExcel(context);
          ws.getRange(`AB${startRow}`).formulas = [
            [`=LET(m,FILTER(${bomMat},(${bomMat}<>"")*(${bomMat}<>"Materials"),""),IF(INDEX(m,1)="","",SORT(UNIQUE(m))))`]
          ];
          ws.getRange("AE2").formulas = [[`=IF($AB$${startRow}="",0,ROWS($AB$${startRow}#))`]];

          const snCol = [];
          const matCol = [];
          for (let i = 0; i < plCount; i++) {
            const r = startRow + i;
            snCol.push([`=IF(P${r}="","",N(O${r - 1})+1)`]);
            matCol.push([`=IF(ROW()-${startRow - 1}>$AE$2,"",INDEX($AB$${startRow}#,ROW()-${startRow - 1}))`]);
          }

          ws.getRange(`O${startRow}:O${plDataEnd}`).formulas = snCol;
          ws.getRange(`P${startRow}:P${plDataEnd}`).formulas = matCol;
          ws.getRange("AB:AB").columnHidden = true;
          await context.sync();
          fixedCount++;
        }
      } finally {
        try {
          if (prevCalc !== null) wb.application.calculationMode = prevCalc;
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        } catch { /* ignore */ }
      }
    });

    log(`Repaired Purchase-List name formulas on ${fixedCount} plain sheet(s).`);
    return { fixedCount };
  }

  window.BomConvert = {
    COMPILED,
    flattenWorkbook,
    restoreWorkbook,
    repairPlainSheets
  };
})();
