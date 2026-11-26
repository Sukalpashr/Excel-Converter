/* convert.js - high-performance two-way BOM converter for Excel (Office.js).
 *   mode "growing" = FlattenBOM.bas  (formula-driven purchase list that grows by itself)
 *   mode "fixed"   = FlattenBOM1.bas (purchase list written once, fixed rows)
 *   restore        = rebuilds converted (plain) BOM sheets back into automated BOM sheets
 * Runs inside Excel (desktop / web / Mac). Nothing leaves the workbook.
 */
(function (root) {
  "use strict";

  var CAP = 40;
  var PL_LAST = 4 + CAP + 3;   // last purchase-list row carrying formulas (growing mode)
  var BOTTOM = 320;
  var KGM2 = "kg/m²";
  var MASTER = "Master Unique Materials";
  var BLANK = "BOM Blank";
  var COMPILED = "Compiled Purchase List";
  var COLS_PO = [22, 23, 24, 25, 27];      // V W X Y AA : typed stock / PO inputs
  var PO_OFFSETS = [7, 8, 9, 10, 12];      // 0-based offsets inside O..AA (O=0 .. AA=12)
  var PAD = 4;                             // empty rows on top of Compiled Purchase List
  var BLOCK = 13;                          // purchase list is O..AA = 13 columns
  var STRIDE = BLOCK + 1;                  // one gap column between lists

  function colLetter(n) {
    var s = "";
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function isBlank(v) { return v === "" || v === null || v === undefined; }
  function isFormula(f) { return typeof f === "string" && f.charAt(0) === "="; }

  /* ---------- Excel calculation & screen-update suspension helpers ---------- */
  function canSuspendScreen() {
    try {
      return typeof Office !== "undefined" &&
        Office.context && Office.context.requirements &&
        Office.context.requirements.isSetSupported("ExcelApi", "1.10");
    } catch (e) { return false; }
  }
  function pauseExcel(context) {
    try { context.application.suspendApiCalculationUntilNextSync(); } catch (e) { /* ignore */ }
    try {
      if (canSuspendScreen() && typeof context.application.suspendScreenUpdatingUntilNextSync === "function") {
        context.application.suspendScreenUpdatingUntilNextSync();
      }
    } catch (e) { /* ignore */ }
  }
  function setManualCalc(wb, prevCalcMode) {
    try {
      if (prevCalcMode !== Excel.CalculationMode.manual) {
        wb.application.calculationMode = Excel.CalculationMode.manual;
      }
    } catch (e) { /* ignore */ }
  }
  async function restoreCalcAndRecalc(context, wb, prevCalcMode) {
    try {
      if (prevCalcMode !== null) wb.application.calculationMode = prevCalcMode;
      wb.application.calculate(Excel.CalculationType.full);
      await context.sync();
    } catch (e) { /* ignore */ }
  }

  /* ---------- formulas of the growing purchase list (zero hidden helper columns — works for ANY custom/unlisted material) ---------- */
  function plFormulas(r, tot) {
    var rng = function (c) { return "$" + c + "$5:$" + c + "$" + tot; };
    var C = rng("C"), G = rng("G"), H = rng("H"), I = rng("I"), J = rng("J"), L = rng("L");
    var k = "(ROW()-4)", n = "$AE$2";
    var bar = "SUMIF(" + C + ",P" + r + "," + L + ")>0";
    var gt = k + "=" + n + "+1", rq = k + "=" + n + "+3", item = k + "<=" + n, after = k + ">" + n;
    var sumAbove = function (c) { return r > 5 ? "SUM(" + c + "$5:" + c + (r - 1) + ")" : '""'; };
    var cntAbove = function (c) {
      return r > 5 ? 'IF(COUNT(' + c + '$5:' + c + (r - 1) + ')=0,"",SUM(' + c + '$5:' + c + (r - 1) + '))' : '""';
    };
    // Direct unique-name lookup from $C$5:$C$<tot> with no hidden Column M or AB helper:
    // Any new row inserted/typed with any custom material name immediately appears in Column P.
    var matLookup = r === 5
      ? 'IFERROR(INDEX(' + C + ',MATCH(1,INDEX((' + C + '<>"")*1,0),0)),"")'
      : 'IFERROR(INDEX(' + C + ',MATCH(1,INDEX((' + C + '<>"")*(COUNTIF($P$5:P' + (r - 1) + ',' + C + '&"")=0),0),0)),"")';
    // Qty formula handles numeric Set, blank Set (defaults to 1 via SUMIFS), and ignores text safely
    var pieceQty = 'SUMPRODUCT(--(' + C + '=P' + r + '),' + G + ',' + H + ')+SUMIFS(' + G + ',' + C + ',P' + r + ',' + H + ',"")';

    var f = {};
    f.O = '=IF(' + item + ',' + k + ',"")';
    f.P = '=IF(' + item + ',' + matLookup + ',' +
          'IF(' + gt + ',"Grand Total Weight in KG :",IF(' + rq + ',"Requested by:","")))';
    f.Q = '=IF(' + rq + ',"Procurement",IF(' + after + ',"",IF(' + bar + ',$AE$1,"")))';
    f.R = '=IF(' + after + ',"",IF(' + bar + ',ROUNDUP(SUMIF(' + C + ',P' + r + ',' + L + ')/Q' + r + ',0),' +
          pieceQty + '))';
    f.S = '=IF(' + gt + ',' + sumAbove("S") + ',IF(' + after + ',"",SUMIF(' + C + ',P' + r + ',' + J + ')))';
    f.T = '=IF(' + rq + ',"Store Section:",IF(' + after + ',"",IF(' + bar + ',INDEX(' + I + ',MATCH(P' + r + ',' + C + ',0)),"")))';
    f.U = '=IF(' + gt + ',' + sumAbove("U") + ',IF(' + after + ',"",IF(' + bar + ',Q' + r + '/1000*R' + r + '*T' + r + ',S' + r + ')))';
    f.W = '=IF(' + gt + ',' + cntAbove("W") + ',IF(' + rq + ',"Production",""))';
    f.Z = '=IF(' + gt + ',' + cntAbove("Z") + ',IF(' + rq + ',"Engineering",IF(' + after + ',"",' +
          'IF(Y' + r + '="","",W' + r + '*Y' + r + '))))';
    return f;
  }

  /* ---------- small Office.js helpers ---------- */
  function setBorders(range, on, outlineOnly) {
    var edges = outlineOnly
      ? ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"]
      : ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical", "InsideHorizontal"];
    for (var i = 0; i < edges.length; i++) {
      var b = range.format.borders.getItem(edges[i]);
      b.style = on ? "Continuous" : "None";
      if (on) b.weight = "Thin";
    }
  }

  /* Sets borders on a single row (outline + inside vertical) in 5 calls instead of 4 * ncols */
  function setSingleRowBorders(range) {
    var edges = ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideVertical"];
    for (var i = 0; i < edges.length; i++) {
      var b = range.format.borders.getItem(edges[i]);
      b.style = "Continuous";
      b.weight = "Thin";
    }
  }

  function parseMasterFromRanges(hdrValues, bodyValues) {
    var headers = hdrValues[0].map(function (s) { return String(s).trim(); });
    var col = function (prefix) {
      var i = headers.findIndex(function (h) { return h.indexOf(prefix) === 0; });
      if (i < 0) throw new Error('Master table has no column starting with "' + prefix + '".');
      return i;
    };
    var cM = col("Material"), cLin = col("Linear Weight"), cAr = col("Area Weight"), cPc = col("Piece Weight");
    var cU = col("Primary Unit"), cSL = col("Sheet Length"), cSW = col("Sheet Width");
    var m = {};
    for (var i = 0; i < bodyValues.length; i++) {
      var row = bodyValues[i];
      var name = row[cM];
      if (isBlank(name)) continue;
      var sName = String(name);
      var key = sName.trim();
      var info = { order: i, unit: row[cU], lin: row[cLin], area: row[cAr], piece: row[cPc], sl: row[cSL], sw: row[cSW] };
      m[sName] = info;
      if (key !== sName) m[key] = info;
    }
    return m;
  }

  /* ---------- Master table ---------- */
  async function readMaster(context, wb) {
    var ws = wb.worksheets.getItemOrNullObject(MASTER);
    await context.sync();
    if (ws.isNullObject) throw new Error('Sheet "' + MASTER + '" not found - is this the original, unconverted BOM workbook?');
    var t = ws.tables.getItemOrNullObject("tblMat");
    await context.sync();
    if (t.isNullObject) throw new Error('Table "tblMat" not found on the Master sheet.');
    var hdr = t.getHeaderRowRange().load("values");
    var body = t.getDataBodyRange().load("values");
    await context.sync();
    return parseMasterFromRanges(hdr.values, body.values);
  }

  /* Find active BOM sheets (skipping the hidden "BOM Blank" template sheet when real take-off sheets exist). */
  function filterBomSheets(sheets) {
    var allBom = sheets.filter(function (ws) {
      return ws.tables.items.some(function (t) { return t.name.indexOf("bom_") === 0; });
    });
    var realBom = allBom.filter(function (ws) {
      return !(ws.name === BLANK && ws.visibility !== "Visible");
    });
    return realBom.length ? realBom : allBom;
  }

  /* Parse a pre-loaded A1:AK<BOTTOM> range into sheet snapshot S */
  function parseSheetRange(vals, frms) {
    var val = function (r, c) { return vals[r - 1][c - 1]; };
    var frm = function (r, c) { return frms[r - 1][c - 1]; };
    var ae1 = val(1, 31);
    var stdLen = (typeof ae1 === "number") ? ae1 : 6000;
    var lastOld = 5;
    for (var r = 6; r <= 305; r++) {
      if (!isBlank(val(r, 2)) || !isBlank(val(r, 3)) || !isBlank(val(r, 4)) || !isBlank(val(r, 7))) lastOld = r;
    }
    return { val: val, frm: frm, stdLen: stdLen, lastOld: lastOld, title: val(2, 15) };
  }

  function readPo(S, master) {   // typed stock / PO values per material (growing + fixed both put them back)
    var po = {};
    for (var r = 6; r <= 60; r++) {
      var mat = String(S.val(r, 16)).trim();
      if (!mat || mat.indexOf("Grand Total") === 0 || mat === "Requested by:") continue;
      if (master[mat] || !isFormula(S.frm(r, 16)) || S.val(r, 15) !== "") {
        var entry = COLS_PO.map(function (c) {
          return (isFormula(S.frm(r, c)) || S.val(r, c) === "") ? undefined : S.val(r, c);
        });
        if (entry.some(function (x) { return x !== undefined; })) {
          po[mat] = entry;
        }
      }
    }
    return po;
  }

  /* Queue table stripping & row-5 removal for a sheet (tables already loaded on ws) */
  function queueStripTables(ws, title) {
    ws.getRange("O2").values = [[title]];               // title formula -> text
    ws.tables.items.forEach(function (t) { t.convertToRange(); });
    ws.getRange().conditionalFormats.clearAll();
    ws.getRange().dataValidation.clear();
    ws.getRange("AD2:AK2").unmerge();
    ws.getRange("5:5").delete("Up");                    // hidden header row
  }

  /* Batch-write BOM block in 4–5 range writes instead of ~600 per-cell writes */
  function writeBomBlockBatched(ws, S, master, o) {
    var first = o.first, last = o.last, tot = o.tot, missing = [];
    if (last < first) return missing;

    var count = last - first + 1;
    var colA = new Array(count);
    var colH = new Array(count);
    var colIJ = new Array(count);
    var colL = new Array(count);
    var colM = o.helperM ? new Array(count) : null;
    var hChanged = false;
    var boldRows = [];

    for (var idx = 0; idx < count; idx++) {
      var r = first + idx;
      var oldRow = r + 1;   // row index in OLD layout (before row 5 was deleted)
      var rawMat = S.val(oldRow, 3);
      var rawDesc = S.val(oldRow, 2);
      var mat = isBlank(rawMat) ? "" : String(rawMat);
      var desc = isBlank(rawDesc) ? "" : String(rawDesc);

      var origHF = S.frm(oldRow, 8);
      var origHV = S.val(oldRow, 8);
      colH[idx] = [isFormula(origHF) ? origHF : (isBlank(origHV) ? "" : origHV)];
      colA[idx] = [""];
      colIJ[idx] = ["", ""];
      colL[idx] = [""];
      if (colM) {
        colM[idx] = ['=IF(C' + r + '="","",COUNTIF($C$' + first + ':$C$' + tot + ',"<"&C' + r + '))'];
      }

      if (mat === "") {
        if (desc !== "") boldRows.push(r);
        continue;
      }

      colA[idx] = ['=IF(C' + r + '="","",COUNTA($C$' + first + ':C' + r + '))'];
      var info = master[mat] || master[mat.trim()];
      if (!info) {
        missing.push("row " + r + ": '" + mat + "' is not in the Master - weight left blank");
        continue;
      }
      if (isBlank(origHV) && !isBlank(S.val(oldRow, 7))) {
        colH[idx] = [1];
        hChanged = true;
      }
      if (info.unit === "kg/m") {
        colIJ[idx] = [info.lin, "=I" + r + "*G" + r + "*H" + r + "*D" + r + "/1000"];
        colL[idx] = ['=IF(ISNUMBER(D' + r + '),D' + r + '*G' + r + '*H' + r + ',"")'];
      } else if (info.unit === KGM2) {
        colIJ[idx] = [info.area, "=I" + r + "*G" + r + "*H" + r + "*D" + r + "*E" + r + "/10^6"];
      } else {
        colIJ[idx] = [info.piece, "=I" + r + "*G" + r + "*H" + r];
      }
    }

    ws.getRange("A" + first + ":A" + last).formulas = colA;
    if (hChanged) ws.getRange("H" + first + ":H" + last).formulas = colH;
    ws.getRange("I" + first + ":J" + last).formulas = colIJ;
    ws.getRange("L" + first + ":L" + last).formulas = colL;
    if (colM) ws.getRange("M" + first + ":M" + last).formulas = colM;

    // Group contiguous section-header rows so bold styling uses minimal API calls
    var bi = 0;
    while (bi < boldRows.length) {
      var startR = boldRows[bi], endR = startR;
      while (bi + 1 < boldRows.length && boldRows[bi + 1] === endR + 1) { bi++; endR = boldRows[bi]; }
      ws.getRange("A" + startR + ":L" + endR).format.font.bold = true;
      bi++;
    }

    return missing;
  }

  /* Consolidated grow-with-list conditional formatting (10 rules instead of 14 — identical visual output) */
  function addGrowCf(ws, c0, r1, r2, nAddr, rowOffset) {
    var off = rowOffset;
    var L = function (c) { return colLetter(c0 + c); };
    var gtF = "=ROW()-" + off + "=" + nAddr + "+1";

    // 1. Data rows: full grid borders on all 13 columns
    var cf = ws.getRange(L(0) + r1 + ":" + L(BLOCK - 1) + r2).conditionalFormats.add("Custom");
    cf.custom.rule.formula = "=ROW()-" + off + "<=" + nAddr;
    ["EdgeLeft", "EdgeRight", "EdgeTop", "EdgeBottom"].forEach(function (e) {
      cf.custom.format.borders.getItem(e).style = "Continuous";
    });
    cf.stopIfTrue = false;

    // 2. Grand Total row: Top & Bottom across all 13 columns
    var cfTB = ws.getRange(L(0) + r1 + ":" + L(BLOCK - 1) + r2).conditionalFormats.add("Custom");
    cfTB.custom.rule.formula = gtF;
    cfTB.custom.format.borders.getItem("EdgeTop").style = "Continuous";
    cfTB.custom.format.borders.getItem("EdgeBottom").style = "Continuous";
    cfTB.stopIfTrue = false;

    // 3. Grand Total row: Left on col 0 (O), Right on col 3 (R), Left+Right on cols 4..12 (S..AA)
    var cfL = ws.getRange(L(0) + r1 + ":" + L(0) + r2).conditionalFormats.add("Custom");
    cfL.custom.rule.formula = gtF;
    cfL.custom.format.borders.getItem("EdgeLeft").style = "Continuous";
    cfL.stopIfTrue = false;

    var cfR = ws.getRange(L(3) + r1 + ":" + L(3) + r2).conditionalFormats.add("Custom");
    cfR.custom.rule.formula = gtF;
    cfR.custom.format.borders.getItem("EdgeRight").style = "Continuous";
    cfR.stopIfTrue = false;

    var cfLR = ws.getRange(L(4) + r1 + ":" + L(BLOCK - 1) + r2).conditionalFormats.add("Custom");
    cfLR.custom.rule.formula = gtF;
    cfLR.custom.format.borders.getItem("EdgeLeft").style = "Continuous";
    cfLR.custom.format.borders.getItem("EdgeRight").style = "Continuous";
    cfLR.stopIfTrue = false;

    // 4. Grand Total row: Bold on cols 1 (P), 4 (S), 6 (U), 8 (W), 11 (Z)
    var boldCols = [1, 4, 6, 8, 11];
    for (var b = 0; b < boldCols.length; b++) {
      var bc = boldCols[b];
      var cfB = ws.getRange(L(bc) + r1 + ":" + L(bc) + r2).conditionalFormats.add("Custom");
      cfB.custom.rule.formula = gtF;
      cfB.custom.format.font.bold = true;
      cfB.stopIfTrue = false;
    }
  }

  /* ---------- MODE 1: queue growing purchase list writes for one sheet (no sync inside) ---------- */
  function queueConvertSheetGrowing(ws, S, master) {
    var first = 5, last = S.lastOld - 1, tot = Math.max(first, last + 1);

    ws.getRange("M:M").clear();
    ws.getRange("AB:AC").clear();
    ws.getRange("AF:AK").clear();
    ws.getRange("AD2:AE" + BOTTOM).clear();
    ws.getRange("M:M").columnHidden = true;
    ws.getRange("AB:AB").columnHidden = true;
    ws.getRange("AC:AC").format.columnWidth = 14;
    if (tot + 1 <= BOTTOM) ws.getRange("A" + (tot + 1) + ":L" + BOTTOM).clear();
    ws.getRange("A" + tot + ":L" + tot).clear("Contents");
    ws.getRange("O" + (PL_LAST + 1) + ":AA" + BOTTOM).clear();
    ws.getRange("O5:AA" + PL_LAST).clear("Contents");
    setBorders(ws.getRange("O5:AA" + PL_LAST), false);

    if (last >= first) setBorders(ws.getRange("A" + first + ":L" + last), true);
    var missing = writeBomBlockBatched(ws, S, master, { first: first, last: last, tot: tot, helperM: false });

    setBorders(ws.getRange("A" + tot + ":L" + tot), true, true);
    ws.getRange("I" + tot + ":J" + tot).formulas = [[
      "Total:",
      tot > first ? "=SUM(J$" + first + ":INDEX(J:J,ROW()-1))" : 0
    ]];
    ws.getRange("I" + tot + ":J" + tot).format.font.bold = true;

    // Batch entire purchase list O5:AA<PL_LAST> in 1 write (no hidden AB helper needed)
    var plRows = [];
    for (var r = 5; r <= PL_LAST; r++) {
      var f = plFormulas(r, tot);
      plRows.push([f.O, f.P, f.Q, f.R, f.S, f.T, f.U, "", f.W, "", "", f.Z, ""]);
    }
    ws.getRange("O5:AA" + PL_LAST).formulas = plRows;

    ws.getRange("AD2").copyFrom("AD1", "Formats");
    ws.getRange("AE2").copyFrom("AE1", "Formats");
    ws.getRange("AE2").format.fill.clear();
    ws.getRange("AE1").values = [[S.stdLen]];
    ws.getRange("AD2:AE3").formulas = [
      [
        "Materials in purchase list:",
        '=ROUND(SUMPRODUCT(($C$5:$C$' + tot + '<>"")/COUNTIF($C$5:$C$' + tot + ',$C$5:$C$' + tot + '&"")),0)'
      ],
      [
        '=IF(AE2>' + CAP + ',"More than ' + CAP + ' materials - copy the last purchase-list row down","")',
        ""
      ]
    ];
    ws.getRange("AD3").format.font.bold = true;
    ws.getRange("AD3").format.font.color = "#C00000";

    addGrowCf(ws, 15, 5, PL_LAST, "$AE$2", 4);

    ws.pageLayout.setPrintArea("A1:AA" + Math.max(tot, PL_LAST));
    ws.freezePanes.unfreeze();
    ws.freezePanes.freezeRows(4);
    return missing;
  }

  /* ---------- MODE 2: queue fixed purchase list writes for one sheet (no sync inside) ---------- */
  function queueConvertSheetFixed(ws, S, po, master) {
    // materials used in the BOM, in Master order (plus any unlisted custom materials at the end)
    var used = {};
    for (var r0 = 6; r0 <= S.lastOld; r0++) {
      var m0 = S.val(r0, 3);
      var mKey = isBlank(m0) ? "" : String(m0).trim();
      if (!mKey || used[mKey] !== undefined) continue;
      used[mKey] = master[mKey] ? master[mKey].order : (1000000 + r0);
    }
    var names = Object.keys(used).sort(function (a, b) { return used[a] - used[b]; });
    var n = names.length;

    var first = 5, last = S.lastOld - 1, tot = Math.max(first, last + 1), gt = 5 + n, rq = 7 + n;
    ws.getRange("M:M").clear();
    ws.getRange("AB:AK").clear();
    ws.getRange("M:M").format.columnWidth = 14; ws.getRange("M:M").columnHidden = false;
    ws.getRange("AB:AB").format.columnWidth = 14; ws.getRange("AB:AB").columnHidden = false;
    if (tot + 1 <= BOTTOM) ws.getRange("A" + (tot + 1) + ":L" + BOTTOM).clear();
    ws.getRange("A" + tot + ":L" + tot).clear("Contents");
    if (rq + 1 <= BOTTOM) ws.getRange("O" + (rq + 1) + ":AA" + BOTTOM).clear();
    ws.getRange("O5:AA" + rq).clear("Contents");

    if (last >= first) setBorders(ws.getRange("A" + first + ":L" + last), true);
    var missing = writeBomBlockBatched(ws, S, master, { first: first, last: last, tot: tot, helperM: false });
    setBorders(ws.getRange("A" + tot + ":L" + tot), true, true);
    ws.getRange("I" + tot + ":J" + tot).formulas = [[
      "Total:",
      last >= first ? "=SUM(J" + first + ":J" + last + ")" : 0
    ]];
    ws.getRange("I" + tot + ":J" + tot).format.font.bold = true;

    // Build entire purchase list O5:AA<rq> in ONE 2D array write
    var Cr = "$C$" + first + ":$C$" + Math.max(first, last);
    var Jr = "$J$" + first + ":$J$" + Math.max(first, last);
    var Lr = "$L$" + first + ":$L$" + Math.max(first, last);
    var Ir = "$I$" + first + ":$I$" + Math.max(first, last);
    var Gr = "$G$" + first + ":$G$" + Math.max(first, last);
    var Hr = "$H$" + first + ":$H$" + Math.max(first, last);

    var plGrid = [];
    for (var i = 1; i <= n; i++) {
      var r = 4 + i, mat = names[i - 1], info = master[mat];
      var row = new Array(BLOCK).fill("");
      row[0] = '=IF(P' + r + '="","",COUNTA(P$' + first + ':P' + r + '))';
      row[1] = mat;
      row[4] = "=SUMIF(" + Cr + ",P" + r + "," + Jr + ")";
      if (info && info.unit === "kg/m") {
        row[2] = S.stdLen;
        row[3] = "=ROUNDUP(SUMIF(" + Cr + ",P" + r + "," + Lr + ")/Q" + r + ",0)";
        row[5] = "=INDEX(" + Ir + ",MATCH(P" + r + "," + Cr + ",0))";
        row[6] = "=Q" + r + "/1000*R" + r + "*T" + r;
      } else if (info && info.unit === KGM2) {
        row[2] = info.sl + " x " + info.sw;
        row[3] = "=ROUNDUP(S" + r + "/(" + info.area + "*" + info.sl + "*" + info.sw + "/10^6),0)";
        row[6] = "=R" + r + "*" + info.area + "*" + info.sl + "*" + info.sw + "/10^6";
      } else if (info) {
        row[3] = "=SUMPRODUCT(--(" + Cr + "=P" + r + ")," + Gr + "," + Hr + ")+SUMIFS(" + Gr + "," + Cr + ",P" + r + "," + Hr + ',"")';
        row[6] = "=R" + r + "*" + info.piece;
      } else {
        // Custom / unlisted material not in Master
        row[3] = "=SUMPRODUCT(--(" + Cr + "=P" + r + ")," + Gr + "," + Hr + ")+SUMIFS(" + Gr + "," + Cr + ",P" + r + "," + Hr + ',"")';
        row[6] = "=S" + r;
      }
      var pv = po[mat];
      if (pv) {
        for (var j = 0; j < pv.length; j++) {
          if (pv[j] !== undefined) row[PO_OFFSETS[j]] = pv[j];
        }
      }
      row[11] = '=IF(Y' + r + '="","",W' + r + '*Y' + r + ')';
      plGrid.push(row);
    }

    // Grand Total row (gt), blank spacer row, Requested by row (rq)
    var gtRow = new Array(BLOCK).fill("");
    gtRow[1] = "Grand Total Weight in KG :";
    gtRow[4] = n > 0 ? "=SUM(S5:S" + (gt - 1) + ")" : 0;
    gtRow[6] = n > 0 ? "=SUM(U5:U" + (gt - 1) + ")" : 0;
    gtRow[8] = n > 0 ? '=IF(COUNT(W5:W' + (gt - 1) + ')=0,"",SUM(W5:W' + (gt - 1) + '))' : "";
    gtRow[11] = n > 0 ? '=IF(COUNT(Z5:Z' + (gt - 1) + ')=0,"",SUM(Z5:Z' + (gt - 1) + '))' : "";
    plGrid.push(gtRow);
    plGrid.push(new Array(BLOCK).fill(""));

    var rqRow = new Array(BLOCK).fill("");
    rqRow[1] = "Requested by:"; rqRow[2] = "Procurement";
    rqRow[5] = "Store Section:"; rqRow[8] = "Production";
    rqRow[11] = "Engineering";
    plGrid.push(rqRow);

    ws.getRange("O5:AA" + rq).formulas = plGrid;

    if (n > 0) setBorders(ws.getRange("O5:AA" + (4 + n)), true);
    setBorders(ws.getRange("O" + gt + ":R" + gt), true, true);
    setSingleRowBorders(ws.getRange("S" + gt + ":AA" + gt));
    ["P", "S", "U", "W", "Z"].forEach(function (L) { ws.getRange(L + gt).format.font.bold = true; });

    ws.pageLayout.setPrintArea("A1:AA" + Math.max(tot, rq));
    ws.freezePanes.unfreeze();
    ws.freezePanes.freezeRows(4);
    return missing;
  }

  /* ---------- EXTRA: compiled purchase list sheet ---------- */
  var num = function (x) { return typeof x === "number" ? x : 0; };

  /* Batch-read purchase list & BOM length data across ALL sheets in ONE sync */
  async function readAllPurchaseData(context, sheets) {
    var loads = sheets.map(function (ws) {
      return {
        ws: ws,
        pl: ws.getRange("O1:AA80").load("values"),
        bom: ws.getRange("A5:L" + BOTTOM).load("values")
      };
    });
    await context.sync();

    return loads.map(function (entry) {
      var ws = entry.ws;
      var v = entry.pl.values, items = [], gt = 0, rq = 0, i;
      for (i = 4; i < v.length; i++) {
        var nm = String(v[i][1]);
        if (nm.indexOf("Grand Total") === 0) { gt = i + 1; break; }
        if (nm !== "") items.push({ mat: nm, Q: v[i][2], R: v[i][3], S: v[i][4], T: v[i][5], U: v[i][6] });
      }
      if (!gt) throw new Error(ws.name + ": purchase list has no Grand Total row.");
      for (i = gt; i < Math.min(v.length, gt + 6); i++) {
        if (String(v[i][1]) === "Requested by:") { rq = i + 1; break; }
      }
      if (!rq) rq = gt + 2;
      var L = {}, tot = 0;
      var bomVals = entry.bom.values;
      for (var idx = 0; idx < bomVals.length; idx++) {
        var row = bomVals[idx];
        var mat = String(row[2]);
        if (String(row[8]) === "Total:") tot = 5 + idx;
        if (mat !== "" && typeof row[11] === "number") L[mat] = (L[mat] || 0) + row[11];
      }
      if (!tot) tot = 5 + bomVals.length;
      return { name: ws.name, vals: v, items: items, gt: gt, rq: rq, L: L, tot: tot, ws: ws };
    });
  }

  function compileMaterials(datas, master) {
    var acc = {};
    datas.forEach(function (d) {
      d.items.forEach(function (it) {
        var a = acc[it.mat] || (acc[it.mat] = { mat: it.mat, Q: it.Q, T: it.T, sumS: 0, sumR: 0, sumU: 0, sumL: 0, sw: null });
        a.sumS += num(it.S); a.sumR += num(it.R); a.sumU += num(it.U);
        if (a.sw === null && num(it.R) > 0) a.sw = num(it.U) / num(it.R);
      });
      Object.keys(d.L).forEach(function (m) { if (acc[m]) acc[m].sumL += d.L[m]; });
    });
    var r3 = function (x) { return Math.round(x * 1000) / 1000; };
    var rows = Object.keys(acc).map(function (k) {
      var a = acc[k], R, U;
      if (typeof a.Q === "number" && a.sumL > 0) { R = Math.ceil(a.sumL / a.Q - 1e-9); U = a.Q / 1000 * R * num(a.T); }
      else if (typeof a.Q === "string" && a.Q !== "" && a.sw) { R = Math.ceil(a.sumS / a.sw - 1e-9); U = R * a.sw; }
      else { R = a.sumR; U = a.sumU; }
      return { mat: a.mat, Q: a.Q, R: R, S: r3(a.sumS), T: a.T, U: r3(U) };
    });
    var ord = function (m) { return master && master[m] ? master[m].order : 1e6; };
    rows.sort(function (a, b) { return ord(a.mat) - ord(b.mat); });
    return rows;
  }

  async function step(context, name, fn) {
    try { pauseExcel(context); fn(); await context.sync(); }
    catch (e) {
      var loc = e && e.debugInfo && e.debugInfo.errorLocation ? " [" + e.debugInfo.errorLocation + "]" : "";
      throw new Error(name + ": " + ((e && e.message) || e) + loc);
    }
  }
  async function soft(context, log, name, fn) {
    try { pauseExcel(context); fn(); await context.sync(); }
    catch (e) { log("Note: skipped '" + name + "' (" + ((e && e.message) || e) + ")"); }
  }
  var qs = function (name) { return "'" + String(name).replace(/'/g, "''") + "'"; };

  function blockBorders(cs, c0, n) {
    var L = function (c) { return colLetter(c0 + c); };
    if (n > 0) setBorders(cs.getRange(L(0) + (PAD + 5) + ":" + L(BLOCK - 1) + (PAD + 4 + n)), true);
    var gr = PAD + 5 + n;
    setBorders(cs.getRange(L(0) + gr + ":" + L(3) + gr), true, true);
    setSingleRowBorders(cs.getRange(L(4) + gr + ":" + L(BLOCK - 1) + gr));
  }

  async function growingBody(context, cs, datas, master, ctx) {
    var nL = datas.length, src = datas[0].ws;
    var capMain = Math.min(nL * CAP, 150), mainLast = PAD + 4 + capMain + 3;
    var lastListEnd = ctx.lastListEnd;
    var hs = lastListEnd + 3;
    var HN = colLetter(hs), HQ = colLetter(hs + 1), HT = colLetter(hs + 2), HS = colLetter(hs + 3), HR = colLetter(hs + 4),
        HU = colLetter(hs + 5), HL = colLetter(hs + 6), HW = colLetter(hs + 7), HK = colLetter(hs + 8),
        OT1 = colLetter(hs + 10), OT2 = colLetter(hs + 11), NC = colLetter(hs + 12);
    var hEnd = 4 + nL * CAP;
    var rg = function (c) { return "$" + c + "$5:$" + c + "$" + hEnd; };
    var rHN = rg(HN), rHQ = rg(HQ), rHT = rg(HT), rHS = rg(HS), rHR = rg(HR), rHU = rg(HU), rHL = rg(HL), rHW = rg(HW), rHK = rg(HK);
    var nMain = "$" + NC + "$5";

    await step(context, "compiled formulas & side lists", function () {
      // 1. Helper block
      var rows = [];
      datas.forEach(function (d, s) {
        var q = qs(d.name);
        for (var i = 1; i <= CAP; i++) {
          var hr = 4 + s * CAP + i, sr = 4 + i, nm = HN + hr;
          var key = '=IF(' + nm + '="","",IF(COUNTIF(' + HN + '$5:' + nm + ',' + nm + ')>1,"",IFERROR(VLOOKUP(' + nm + ',$' + OT1 + '$5:$' + OT2 + '$600,2,0),1000000)+ROW()/100000))';
          rows.push([
            '=IF(ISNUMBER(' + q + '!O' + sr + '),' + q + '!P' + sr + ',"")',
            '=IF(' + nm + '="","",' + q + '!Q' + sr + ')',
            '=IF(' + nm + '="","",' + q + '!T' + sr + ')',
            '=IF(' + nm + '="",0,N(' + q + '!S' + sr + '))',
            '=IF(' + nm + '="",0,N(' + q + '!R' + sr + '))',
            '=IF(' + nm + '="",0,N(' + q + '!U' + sr + '))',
            '=IF(' + nm + '="",0,SUMIF(' + q + '!$C$5:$C$' + d.tot + ',' + nm + ',' + q + '!$L$5:$L$' + d.tot + '))',
            '=IF(' + nm + '="","",IFERROR(' + q + '!U' + sr + '/' + q + '!R' + sr + ',""))',
            key
          ]);
        }
      });
      cs.getRange(HN + "5:" + HK + hEnd).formulas = rows;
      var ord = Object.keys(master).map(function (k) { return [k, master[k].order]; });
      if (ord.length) cs.getRange(OT1 + "5:" + OT2 + (4 + ord.length)).values = ord;
      var ncCol = [["=COUNT(" + rHK + ")"]];
      datas.forEach(function (d) { ncCol.push(["=" + qs(d.name) + "!$AE$2"]); });
      cs.getRange(NC + "5:" + NC + (5 + datas.length)).formulas = ncCol;

      // 2. Compiled list formulas
      var k = "(ROW()-" + (PAD + 4) + ")", n = nMain, out = [];
      for (var r = PAD + 5; r <= mainLast; r++) {
        var B = "B" + r, C = "C" + r, D = "D" + r, F = "F" + r;
        var bar = 'AND(ISNUMBER(' + C + '),SUMIF(' + rHN + ',' + B + ',' + rHL + ')>0)';
        var sht = 'AND(ISTEXT(' + C + '),' + C + '<>"")';
        var sw = 'INDEX(' + rHW + ',MATCH(' + B + ',' + rHN + ',0))';
        var above = function (c) { return r > PAD + 5 ? 'SUM(' + c + '$' + (PAD + 5) + ':' + c + (r - 1) + ')' : '""'; };
        var row = new Array(BLOCK).fill("");
        row[0] = '=IF(' + k + '<=' + n + ',' + k + ',"")';
        row[1] = '=IF(' + k + '<=' + n + ',INDEX(' + rHN + ',MATCH(SMALL(' + rHK + ',' + k + '),' + rHK + ',0)),IF(' + k + '=' + n + '+1,"Grand Total Weight in KG :",IF(' + k + '=' + n + '+3,"Requested by:","")))';
        row[2] = '=IF(' + k + '=' + n + '+3,"Procurement",IF(' + k + '>' + n + ',"",IFERROR(INDEX(' + rHQ + ',MATCH(' + B + ',' + rHN + ',0)),"")))';
        row[3] = '=IF(' + k + '>' + n + ',"",IF(' + bar + ',ROUNDUP(ROUND(SUMIF(' + rHN + ',' + B + ',' + rHL + ')/' + C + ',6),0),IF(' + sht + ',ROUNDUP(ROUND(SUMIF(' + rHN + ',' + B + ',' + rHS + ')/' + sw + ',6),0),SUMIF(' + rHN + ',' + B + ',' + rHR + '))))';
        row[4] = '=IF(' + k + '=' + n + '+1,' + above("E") + ',IF(' + k + '>' + n + ',"",SUMIF(' + rHN + ',' + B + ',' + rHS + ')))';
        row[5] = '=IF(' + k + '=' + n + '+3,"Store Section:",IF(' + k + '>' + n + ',"",IFERROR(INDEX(' + rHT + ',MATCH(' + B + ',' + rHN + ',0)),"")))';
        row[6] = '=IF(' + k + '=' + n + '+1,' + above("G") + ',IF(' + k + '>' + n + ',"",IF(' + bar + ',' + C + '/1000*' + D + '*' + F + ',IF(' + sht + ',' + D + '*' + sw + ',SUMIF(' + rHN + ',' + B + ',' + rHU + ')))))';
        row[8] = '=IF(' + k + '=' + n + '+3,"Production","")';
        row[11] = '=IF(' + k + '=' + n + '+3,"Engineering","")';
        out.push(row);
      }
      cs.getRange("A" + (PAD + 5) + ":M" + mainLast).formulas = out;

      // 3. Side lists
      datas.forEach(function (d, idx) {
        var c0 = 1 + STRIDE * (idx + 1), q = qs(d.name), sOut = [];
        for (var sr = 5; sr <= PL_LAST; sr++) {
          var sRow = [];
          for (var c = 0; c < BLOCK; c++) { var a = q + "!" + colLetter(15 + c) + sr; sRow.push('=IF(' + a + '="","",' + a + ')'); }
          sOut.push(sRow);
        }
        cs.getRange(colLetter(c0) + (PAD + 5) + ":" + colLetter(c0 + BLOCK - 1) + (PAD + PL_LAST)).formulas = sOut;
      });
    });

    await step(context, "formats & conditional formatting", function () {
      // Single tiled copyFrom instead of 120+ per-row copyFrom calls
      cs.getRange("A" + (PAD + 5) + ":M" + mainLast).copyFrom(src.getRange("O5:AA5"), "Formats");
      datas.forEach(function (d, idx) {
        var c0 = 1 + STRIDE * (idx + 1);
        cs.getRange(colLetter(c0) + (PAD + 5) + ":" + colLetter(c0 + BLOCK - 1) + (PAD + PL_LAST)).copyFrom(d.ws.getRange("O5:AA" + PL_LAST), "Formats");
      });
      cs.getRange().conditionalFormats.clearAll();
      addGrowCf(cs, 1, PAD + 5, mainLast, nMain, PAD + 4);
      datas.forEach(function (d, idx) {
        addGrowCf(cs, 1 + STRIDE * (idx + 1), PAD + 5, PAD + PL_LAST, "$" + NC + "$" + (6 + idx), PAD + 4);
      });
      cs.getRange(HN + ":" + colLetter(hs + 12)).columnHidden = true;
    });
    return { lastRow: mainLast };
  }

  async function staticBody(context, cs, datas, master) {
    var rows = compileMaterials(datas, master), n = rows.length, first = datas[0], src = first.ws;
    var gr = PAD + 5 + n, rr = gr + 2;

    await step(context, "compiled list values & formats", function () {
      if (n > 0) cs.getRange("A" + (PAD + 5) + ":M" + (PAD + 4 + n)).copyFrom(src.getRange("O5:AA5"), "Formats");
      cs.getRange("A" + gr + ":M" + gr).copyFrom(src.getRange("O" + first.gt + ":AA" + first.gt), "Formats");
      cs.getRange("A" + rr + ":M" + rr).copyFrom(src.getRange("O" + first.rq + ":AA" + first.rq), "Formats");

      var vals = rows.map(function (r, i) { return [i + 1, r.mat, r.Q, r.R, r.S, r.T, r.U, "", "", "", "", "", ""]; });
      if (n > 0) cs.getRange("A" + (PAD + 5) + ":M" + (PAD + 4 + n)).values = vals;
      var tot = function (k) { return Math.round(rows.reduce(function (a, r) { return a + num(r[k]); }, 0) * 1000) / 1000; };
      cs.getRange("A" + gr + ":M" + gr).values = [["", "Grand Total Weight in KG :", "", "", tot("S"), "", tot("U"), "", "", "", "", "", ""]];
      cs.getRange("A" + rr + ":M" + rr).values = [first.vals[first.rq - 1].map(function (x) { return x === null ? "" : x; })];

      datas.forEach(function (d, k) {
        var c0 = 1 + STRIDE * (k + 1), a = colLetter(c0), b = colLetter(c0 + BLOCK - 1);
        cs.getRange(a + (PAD + 5) + ":" + b + (PAD + d.rq)).copyFrom(d.ws.getRange("O5:AA" + d.rq), "Formats");
        cs.getRange(a + (PAD + 5) + ":" + b + (PAD + d.rq)).values =
          d.vals.slice(4, d.rq).map(function (r) { return r.map(function (x) { return x === null ? "" : x; }); });
      });

      cs.getRange().conditionalFormats.clearAll();
      blockBorders(cs, 1, n);
      datas.forEach(function (d, k) { blockBorders(cs, 1 + STRIDE * (k + 1), d.items.length); });
    });
    return { lastRow: Math.max(rr, PAD + Math.max.apply(null, datas.map(function (d) { return d.rq; }))) };
  }

  async function buildCompiled(context, wb, sheets, master, log, growing) {
    var datas = await readAllPurchaseData(context, sheets);
    var nL = datas.length, first = datas[0], src = first.ws;

    var old = wb.worksheets.getItemOrNullObject(COMPILED);
    var srcCols = [];
    for (var c = 0; c < BLOCK; c++) {
      var cr = src.getRange(colLetter(15 + c) + ":" + colLetter(15 + c));
      cr.format.load("columnWidth");
      srcCols.push(cr);
    }
    await context.sync();
    var widths = srcCols.map(function (r) { return r.format.columnWidth; });

    pauseExcel(context);
    if (!old.isNullObject) old.delete();
    var cs = wb.worksheets.add(COMPILED);
    var lastListEnd = 1 + STRIDE * nL + BLOCK - 1;

    await step(context, "column widths & headers", function () {
      var setW = function (c0) {
        for (var col = 0; col < BLOCK; col++) cs.getRange(colLetter(c0 + col) + ":" + colLetter(c0 + col)).format.columnWidth = widths[col];
      };
      setW(1);
      for (var k = 0; k <= nL; k++) {
        if (k > 0) setW(1 + STRIDE * k);
        cs.getRange(colLetter(1 + STRIDE * k + BLOCK) + ":" + colLetter(1 + STRIDE * k + BLOCK)).format.columnWidth = 18;
      }

      cs.getRange("A" + (PAD + 1) + ":M" + (PAD + 4)).copyFrom(src.getRange("O1:AA4"), "All");
      datas.forEach(function (d, k2) {
        var c0 = 1 + STRIDE * (k2 + 1);
        cs.getRange(colLetter(c0) + (PAD + 1) + ":" + colLetter(c0 + BLOCK - 1) + (PAD + 4)).copyFrom(d.ws.getRange("O1:AA4"), "All");
        var lab = cs.getRange(colLetter(c0) + PAD);
        lab.values = [[d.name]];
        lab.format.font.bold = true;
      });
      var title = String(first.vals[1][0] || "");
      cs.getRange("A" + (PAD + 2)).values = [[title.indexOf("Remaining Purchase List") >= 0
        ? title.replace("Remaining Purchase List", "Compiled Purchase List (all sheets)") : "Compiled Purchase List (all sheets)"]];
    });

    var res = growing
      ? await growingBody(context, cs, datas, master, { lastListEnd: lastListEnd })
      : await staticBody(context, cs, datas, master);

    await soft(context, log, "page layout", function () {
      cs.pageLayout.orientation = "Landscape";
      cs.pageLayout.setPrintArea("A1:" + colLetter(lastListEnd) + res.lastRow);
      for (var p = 0; p < nL; p++) cs.verticalPageBreaks.add(cs.getRange(colLetter(1 + STRIDE * (p + 1)) + "1"));
      cs.activate();
    });
    log("Compiled sheet created (" + (growing ? "grows with the BOM" : "values only") + ") from " + nL + " sheet(s).");
  }

  /* ---------- workbook level cleanup (batched in 2 syncs instead of 9) ---------- */
  async function cleanWorkbook(context, wb, convertedNames) {
    wb.names.load("items/name");
    var drop = ["Master Unique Materials", "Formulas", "README"];
    if (!convertedNames || convertedNames.indexOf(BLANK) < 0) drop.push(BLANK);
    var dropRefs = drop.map(function (name) { return wb.worksheets.getItemOrNullObject(name); });
    await context.sync();

    pauseExcel(context);
    wb.names.items.forEach(function (nm) { if (nm.name.indexOf("Print_") < 0) nm.delete(); });
    try { await context.sync(); } catch (e) { /* some names are protected - ignore */ }

    pauseExcel(context);
    var deletedAny = false;
    dropRefs.forEach(function (sh) {
      if (!sh.isNullObject) { sh.delete(); deletedAny = true; }
    });
    if (deletedAny) await context.sync();
  }

  /* ---------- detect converted (flattened) BOM sheets ---------- */
  var SKIP_SHEETS = { "master unique materials": 1, "bom blank": 1, "compiled purchase list": 1, "formulas": 1, "readme": 1 };

  async function findConvertedSheets(context, wb, preloadedSheets) {
    var sheets = preloadedSheets;
    if (!sheets) {
      wb.worksheets.load("items/name,items/visibility,items/position");
      await context.sync();
      sheets = wb.worksheets.items;
      sheets.forEach(function (ws) { ws.tables.load("items/name"); });
      await context.sync();
    }

    var candidates = sheets.filter(function (ws) {
      var low = String(ws.name || "").trim().toLowerCase();
      if (SKIP_SHEETS[low]) return false;
      if (ws.visibility !== "Visible") return false;
      var hasBomTable = ws.tables.items.some(function (t) { return t.name.indexOf("bom_") === 0; });
      return !hasBomTable;
    });
    if (!candidates.length) return [];

    var hdrs = candidates.map(function (ws) {
      return ws.getRange("A1:P4").load("values");
    });
    await context.sync();

    var found = [];
    candidates.forEach(function (ws, idx) {
      var v = hdrs[idx].values;
      var a1 = String((v[0] && v[0][0]) || "").toLowerCase();
      var o2 = String((v[1] && v[1][14]) || "").toLowerCase();
      var b3 = String((v[2] && v[2][1]) || "").toLowerCase();
      var b4 = String((v[3] && v[3][1]) || "").toLowerCase();
      var c3 = String((v[2] && v[2][2]) || "").toLowerCase();
      var c4 = String((v[3] && v[3][2]) || "").toLowerCase();
      var looksLikeBom =
        a1.indexOf("bill of materials") >= 0 ||
        o2.indexOf("purchase list") >= 0 ||
        ((b3.indexOf("description") >= 0 || b4.indexOf("description") >= 0) &&
         (c3.indexOf("material") >= 0 || c4.indexOf("material") >= 0));
      if (looksLikeBom) found.push(ws);
    });
    return found;
  }

  /* Parse pre-loaded A1:AE<BOTTOM> values & formulas from a converted BOM sheet */
  function parseConvertedSheetData(sheetName, vals, frms) {
    var val = function (r, c) { return vals[r - 1][c - 1]; };
    var frm = function (r, c) { return frms[r - 1][c - 1]; };

    var title = isBlank(val(1, 1)) ? "" : String(val(1, 1));
    var subtitle = isBlank(val(2, 1)) ? sheetName : String(val(2, 1));

    var stdLen = 6000;
    if (typeof val(1, 31) === "number" && val(1, 31) > 0) {
      stdLen = val(1, 31);
    } else {
      for (var rq = 5; rq <= 45; rq++) {
        var qv = val(rq, 17);
        if (typeof qv === "number" && qv > 0) { stdLen = qv; break; }
      }
    }

    var stopRow = BOTTOM;
    for (var r = 5; r <= BOTTOM; r++) {
      var iv = String(val(r, 9) || "").trim().toLowerCase();
      if (iv === "total:" || iv === "total") { stopRow = r - 1; break; }
    }
    var lastDataRow = 4;
    for (var r2 = 5; r2 <= stopRow; r2++) {
      if (!isBlank(val(r2, 2)) || !isBlank(val(r2, 3)) || !isBlank(val(r2, 4)) || !isBlank(val(r2, 7)) || !isBlank(val(r2, 11))) {
        lastDataRow = r2;
      }
    }

    var bomRows = [];
    var usedMaterials = {};
    for (var r3 = 5; r3 <= lastDataRow; r3++) {
      var matVal = isBlank(val(r3, 3)) ? "" : String(val(r3, 3)).trim();
      if (matVal) usedMaterials[matVal] = true;
      bomRows.push({
        desc: isBlank(val(r3, 2)) ? "" : val(r3, 2),
        mat: matVal,
        len: isBlank(val(r3, 4)) ? "" : val(r3, 4),
        wid: isBlank(val(r3, 5)) ? "" : val(r3, 5),
        thk: isBlank(val(r3, 6)) ? "" : val(r3, 6),
        qty: isBlank(val(r3, 7)) ? "" : val(r3, 7),
        set: isBlank(val(r3, 8)) ? "" : val(r3, 8),
        rem: isBlank(val(r3, 11)) ? "" : val(r3, 11)
      });
    }

    var po = {};
    for (var rp = 5; rp <= 55; rp++) {
      var pName = isBlank(val(rp, 16)) ? "" : String(val(rp, 16)).trim();
      if (!pName) continue;
      if (pName.toLowerCase().indexOf("grand total") === 0 || pName.toLowerCase().indexOf("requested by") === 0) break;
      var entry = COLS_PO.map(function (c) {
        var f = frm(rp, c), v = val(rp, c);
        return (isFormula(f) || isBlank(v)) ? undefined : v;
      });
      if (entry.some(function (x) { return x !== undefined; })) {
        po[pName.toLowerCase()] = entry;
      }
    }

    return {
      name: sheetName,
      title: title,
      subtitle: subtitle,
      stdLen: stdLen,
      bomRows: bomRows,
      usedMaterials: Object.keys(usedMaterials),
      po: po
    };
  }

  /* Restore converted (plain) BOM sheets back into full automated BOM sheets — fully batched under Manual calc mode. */
  async function restoreAutomatedWorkbook(log, opts) {
    log = log || function () {};
    opts = opts || {};
    var T = root.BomTools;
    var S = root.BomSync;
    if (!T || !S) throw new Error("Required modules (tools.js / sync.js) are not loaded.");

    var warnings = [];
    var extracted = [];

    var selectedSet = null;
    if (Array.isArray(opts.sheets)) {
      if (!opts.sheets.length) throw new Error("Select at least one converted sheet to restore.");
      selectedSet = {};
      opts.sheets.forEach(function (n) { selectedSet[String(n).trim().toLowerCase()] = true; });
    }

    // 1. Inspect and batch-read all selected converted BOM sheets in ONE sync
    var hasMaster = false, hasBlank = false;
    await Excel.run(async function (context) {
      var wb = context.workbook;
      wb.worksheets.load("items/name,items/visibility,items/position");
      await context.sync();
      var allSheets = wb.worksheets.items;
      allSheets.forEach(function (ws) { ws.tables.load("items/name"); });
      await context.sync();

      hasMaster = allSheets.some(function (w) { return w.name.toLowerCase() === MASTER.toLowerCase(); });
      hasBlank = allSheets.some(function (w) { return w.name.toLowerCase() === BLANK.toLowerCase(); });

      var convSheets = await findConvertedSheets(context, wb, allSheets);
      if (selectedSet) {
        convSheets = convSheets.filter(function (ws) { return !!selectedSet[String(ws.name).trim().toLowerCase()]; });
      }
      if (!convSheets.length) {
        throw new Error(selectedSet
          ? "None of the selected converted sheets were found."
          : "No converted BOM sheets found to restore. (Automated BOM sheets with bom_ tables are already automated.)");
      }

      log("Reading " + convSheets.length + " converted sheet(s) ...");
      var ranges = convSheets.map(function (ws) {
        return ws.getRange("A1:AE" + BOTTOM).load(["values", "formulas"]);
      });
      await context.sync();
      convSheets.forEach(function (ws, idx) {
        extracted.push(parseConvertedSheetData(ws.name, ranges[idx].values, ranges[idx].formulas));
      });
    });

    // 2. Ensure Master Unique Materials and hidden BOM Blank exist
    if (!hasMaster || !hasBlank) {
      if (hasMaster && !hasBlank) {
        var existingMaster = [];
        try { existingMaster = await S.workbookList(); } catch (e) { existingMaster = []; }
        await Excel.run(async function (context) {
          var mWs = context.workbook.worksheets.getItemOrNullObject(MASTER);
          await context.sync();
          if (!mWs.isNullObject) { mWs.delete(); await context.sync(); }
        });
        await S.importTemplate(log);
        if (existingMaster.length) {
          await S.applyList(existingMaster, log, { syncRemovals: false });
        }
      } else {
        await S.importTemplate(log);
      }
    }

    // 3. Apply either a saved Project-Only JSON materials list OR sync the latest Global Master list
    if (opts.projectMaterials && Array.isArray(opts.projectMaterials.list) && opts.projectMaterials.list.length) {
      var pm = opts.projectMaterials;
      var projTitle = pm.projectName || "Project Custom Materials";
      log("Applying saved Project-Only materials (\"" + projTitle + "\", " + pm.list.length + " items) ...");
      await S.applyList(pm.list, log, { syncRemovals: true, lastChange: pm.lastChange || null });
      if (T.setProjectMaterialMode) {
        await T.setProjectMaterialMode(projTitle);
      }
      log("Workbook set to Project-Only Materials mode (Global Master list untouched).");
    } else if (opts.syncLocal !== false) {
      try {
        await S.pullLocal(log, { allowPrompt: true, syncRemovals: true });
      } catch (e) {
        log("Note: skipped local master file sync (" + (e.message || e) + ").");
      }
    }

    // 4. Rebuild ALL selected converted sheets in batched syncs under Manual calculation mode
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var prevCalcMode = null;
      try {
        wb.application.load("calculationMode");
        wb.worksheets.load("items/name");
        wb.tables.load("items/name");
        await context.sync();
        prevCalcMode = wb.application.calculationMode;
        setManualCalc(wb, prevCalcMode);
      } catch (e) {
        prevCalcMode = null;
        wb.worksheets.load("items/name");
        wb.tables.load("items/name");
        await context.sync();
      }

      try {
        var matInfo = await T._ensureMasterNames(context, wb, log);
        var masterLookup = await readMaster(context, wb);

        var blankWs = wb.worksheets.getItem(BLANK);
        blankWs.tables.load("items/name");
        wb.worksheets.load("items/name");
        wb.tables.load("items/name");
        await context.sync();

        var srcBom = blankWs.tables.items.filter(function (t) { return t.name.indexOf("bom_") === 0; })[0];
        var srcPl = blankWs.tables.items.filter(function (t) { return t.name.indexOf("pl_") === 0; })[0];
        if (!srcBom) throw new Error("Hidden \"BOM Blank\" sheet has no bom_ table.");

        var existingNames = wb.worksheets.items.map(function (w) { return w.name; });
        var takenTables = wb.tables.items.map(function (t) { return t.name; });

        // Batch A: Rename all original sheets to temp names & copy BOM Blank for each
        log("Rebuilding " + extracted.length + " sheet(s) from template ...");
        pauseExcel(context);
        var jobs = extracted.map(function (data, idx) {
          data.usedMaterials.forEach(function (matName) {
            if (!masterLookup[matName] && !masterLookup[matName.trim()]) {
              warnings.push(data.name + ": \"" + matName + "\" is not in Master Unique Materials (add it in the Materials tab so weights calculate).");
            }
          });
          var oldWs = wb.worksheets.getItem(data.name);
          var tmpName = T.uniqueName("__tmp_" + idx, existingNames).slice(0, 31);
          existingNames.push(tmpName);
          oldWs.name = tmpName;

          var restored = blankWs.copy(Excel.WorksheetPositionType.after, oldWs);
          restored.name = data.name;
          restored.visibility = Excel.SheetVisibility.visible;
          restored.tables.load("items/name");
          return { data: data, oldWs: oldWs, restored: restored };
        });
        await context.sync();

        // Batch B: Rename tables, set titles/AE1, and load bomBody rowCount for all sheets
        pauseExcel(context);
        jobs.forEach(function (job) {
          job.bomT = job.restored.tables.items.filter(function (t) { return t.name.indexOf("bom_") === 0; })[0];
          job.plT = job.restored.tables.items.filter(function (t) { return t.name.indexOf("pl_") === 0; })[0];
          var id = T.safeId(job.data.name);
          job.bomName = T.uniqueName("bom_" + id, takenTables);
          takenTables.push(job.bomName);
          job.plName = T.uniqueName("pl_" + id, takenTables);
          takenTables.push(job.plName);

          if (job.bomT) job.bomT.name = job.bomName;
          if (job.plT) job.plT.name = job.plName;
          if (job.data.title) job.restored.getRange("A1").values = [[job.data.title]];
          job.restored.getRange("A2").values = [[job.data.subtitle || job.data.name]];
          job.restored.getRange("AE1").values = [[job.data.stdLen || 6000]];

          if (job.bomT && job.data.bomRows.length) {
            job.bomBodyMeta = job.bomT.getDataBodyRange().load("rowCount");
          }
        });
        await context.sync();

        // Batch C (rare): Expand BOM table if any sheet has > 300 rows
        var expandedAny = false;
        pauseExcel(context);
        jobs.forEach(function (job) {
          if (job.bomBodyMeta && job.data.bomRows.length > job.bomBodyMeta.rowCount) {
            var extraRows = job.data.bomRows.length - job.bomBodyMeta.rowCount;
            for (var er = 0; er < extraRows; er++) job.bomT.rows.add();
            expandedAny = true;
          }
        });
        if (expandedAny) await context.sync();

        // Batch D: Write BOM input rows (B6:H & K6:K) + load formulas to verify across ALL sheets
        log("Restoring BOM rows, formulas, and dropdowns ...");
        pauseExcel(context);
        jobs.forEach(function (job) {
          if (job.bomT && job.data.bomRows.length) {
            var endRow = 5 + job.data.bomRows.length;
            var colsBtoH = job.data.bomRows.map(function (r) {
              return [r.desc, r.mat, r.len, r.wid, r.thk, r.qty, r.set];
            });
            var colK = job.data.bomRows.map(function (r) { return [r.rem]; });
            job.restored.getRange("B6:H" + endRow).values = colsBtoH;
            job.restored.getRange("K6:K" + endRow).values = colK;
          }
          // Queue header formula checks in the same sync
          job.rO2 = job.restored.getRange("O2").load("formulas");
          job.rM4 = job.restored.getRange("M4").load("formulas");
          job.rAB4 = job.restored.getRange("AB4").load("formulas");
          job.rAD2 = job.restored.getRange("AD2").load("formulas");
          if (job.bomT && srcBom.name !== job.bomName) {
            job.bomFormulas = job.bomT.getDataBodyRange().load("formulas");
          }
          if (job.plT && (srcBom.name !== job.bomName || (srcPl && srcPl.name !== job.plName))) {
            job.plFormulas = job.plT.getDataBodyRange().load("formulas");
          }
        });
        var compWs = opts.removeCompiled ? wb.worksheets.getItemOrNullObject(COMPILED) : null;
        await context.sync();

        // Batch E: Fix any stale formulas, enforce layout/hidden rows & delete temp old sheets
        pauseExcel(context);
        jobs.forEach(function (job) {
          var ws = job.restored;
          var bomName = job.bomName, plName = job.plT ? job.plName : "";
          var fO2 = String(job.rO2.formulas[0][0] || "");
          if (!isFormula(fO2) || fO2.indexOf("#REF!") >= 0) job.rO2.formulas = [[T._canonical.O2()]];

          var fM4 = String(job.rM4.formulas[0][0] || "");
          if (!isFormula(fM4) || fM4.indexOf(bomName + "[") < 0 || fM4.indexOf("#REF!") >= 0) {
            job.rM4.formulas = [[T._canonical.M4(bomName)]];
          }
          if (plName) {
            var fAB4 = String(job.rAB4.formulas[0][0] || "");
            if (!isFormula(fAB4) || fAB4.indexOf(plName + "[") < 0 || fAB4.indexOf("#REF!") >= 0) {
              job.rAB4.formulas = [[T._canonical.AB4(plName)]];
            }
            var fAD2 = String(job.rAD2.formulas[0][0] || "");
            if (!isFormula(fAD2) || fAD2.indexOf(bomName + "[") < 0 || fAD2.indexOf(plName + "[") < 0 || fAD2.indexOf("#REF!") >= 0) {
              job.rAD2.formulas = [[T._canonical.AD2(bomName, plName)]];
            }
          }

          if (job.bomFormulas) {
            var bomChanged = false;
            var bomOut = job.bomFormulas.formulas.map(function (row) {
              return row.map(function (cell) {
                var fixed = T._replaceTableRef(cell, srcBom.name, bomName);
                if (fixed !== cell) bomChanged = true;
                return fixed;
              });
            });
            if (bomChanged) job.bomFormulas.formulas = bomOut;
          }
          if (job.plFormulas) {
            var plChanged = false;
            var plOut = job.plFormulas.formulas.map(function (row) {
              return row.map(function (cell) {
                var fixed = T._replaceTableRef(cell, srcBom.name, bomName);
                if (srcPl && srcPl.name !== plName) fixed = T._replaceTableRef(fixed, srcPl.name, plName);
                if (fixed !== cell) plChanged = true;
                return fixed;
              });
            });
            if (plChanged) job.plFormulas.formulas = plOut;
          }

          ws.getRange("5:5").rowHidden = true;
          ws.getRange("M:M").columnHidden = true;
          ws.getRange("AB:AB").columnHidden = true;
          ws.showGridlines = false;
          ws.freezePanes.unfreeze();
          ws.freezePanes.freezeRows(5);

          job.oldWs.delete();
        });
        if (compWs && !compWs.isNullObject) compWs.delete();
        await context.sync();

        // Batch E2: Apply Materials dropdown across all restored sheets in 1 sync (with fallback per sheet)
        var dvFallbackJobs = [];
        try {
          pauseExcel(context);
          jobs.forEach(function (job) {
            var target = job.restored.getRange("C6:C2000");
            target.dataValidation.clear();
            target.dataValidation.rule = { list: { inCellDropDown: true, source: "=Materials" } };
            target.dataValidation.ignoreBlanks = true;
            target.dataValidation.prompt = { showPrompt: false, title: "", message: "" };
            target.dataValidation.errorAlert = {
              showAlert: true,
              style: Excel.DataValidationAlertStyle.stop,
              title: "Not in the materials list",
              message: "Pick a material from the dropdown (or add it to the Master Unique Materials sheet first)."
            };
          });
          await context.sync();
        } catch (batchErr) {
          dvFallbackJobs = jobs;
        }
        if (dvFallbackJobs.length) {
          for (var df = 0; df < dvFallbackJobs.length; df++) {
            await T._applyBomDataValidation(context, dvFallbackJobs[df].restored, matInfo, log);
          }
        }

        // Batch F: Print setup for all restored sheets (non-fatal)
        try {
          pauseExcel(context);
          jobs.forEach(function (job) {
            job.restored.pageLayout.orientation = Excel.PageOrientation.landscape;
            job.restored.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0 };
            job.restored.pageLayout.setPrintTitleRows("3:4");
          });
          await context.sync();
        } catch (e) { /* non-fatal on hosts without full pageLayout support */ }

        // Batch G (ONLY if any sheet had typed PO/Stock inputs): 1 recalc + batch read + batch write
        var poJobs = jobs.filter(function (job) { return job.plT && Object.keys(job.data.po).length > 0; });
        if (poJobs.length) {
          log("Restoring saved Stock / PO inputs ...");
          wb.application.calculate(Excel.CalculationType.full);
          poJobs.forEach(function (job) {
            job.plBody = job.plT.getDataBodyRange().load(["values", "rowIndex"]);
          });
          await context.sync();

          pauseExcel(context);
          var wrotePo = false;
          poJobs.forEach(function (job) {
            var plStart = (job.plBody.rowIndex || 5) + 1;
            job.plBody.values.forEach(function (row, idx) {
              var matKey = String(row[1] || "").trim().toLowerCase();
              var pv = matKey ? job.data.po[matKey] : null;
              if (!pv) return;
              var excelRow = plStart + idx;
              for (var j = 0; j < pv.length; j++) {
                if (pv[j] !== undefined) {
                  job.restored.getRange(colLetter(COLS_PO[j]) + excelRow).values = [[pv[j]]];
                  wrotePo = true;
                }
              }
            });
          });
          if (wrotePo) await context.sync();
        }

        if (jobs.length) jobs[0].restored.activate();
      } finally {
        await restoreCalcAndRecalc(context, wb, prevCalcMode);
      }
    });

    log("Restored " + extracted.length + " sheet(s) back to automated BOM format.");
    return { restoredCount: extracted.length, warnings: warnings };
  }

  async function inspectWorkbook() {
    var out = { master: false, bomSheets: [], convertedSheets: [], masterMode: { isProject: false, projectName: "" } };
    await Excel.run(async function (context) {
      var wb = context.workbook;
      wb.worksheets.load("items/name,items/visibility,items/position");
      await context.sync();
      var allSheets = wb.worksheets.items;
      var masterWs = allSheets.filter(function (w) { return w.name.toLowerCase() === MASTER.toLowerCase(); })[0];
      out.master = !!masterWs;
      allSheets.forEach(function (ws) { ws.tables.load("items/name"); });
      var metaCell = masterWs ? masterWs.getRange("Z1").load("values") : null;
      await context.sync();
      if (metaCell && root.BomTools && root.BomTools.parseProjectMeta) {
        out.masterMode = root.BomTools.parseProjectMeta(metaCell.values[0][0]);
      }
      var sheets = filterBomSheets(allSheets);
      out.bomSheets = sheets.map(function (s) { return s.name; });
      var conv = await findConvertedSheets(context, wb, allSheets);
      out.convertedSheets = conv.map(function (s) { return s.name; });
    });
    return out;
  }

  /* High-speed batched Flatten conversion across all BOM sheets under Manual calculation mode */
  async function convertWorkbook(mode, log, opts) {
    log = log || function () {};
    opts = opts || {};
    var warnings = [];
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var prevCalcMode = null;

      // Sync 1: Load calculationMode + all worksheets
      try { wb.application.load("calculationMode"); } catch (e) { /* ignore */ }
      wb.worksheets.load("items/name,items/visibility");
      await context.sync();

      try {
        prevCalcMode = wb.application.calculationMode;
        setManualCalc(wb, prevCalcMode);
      } catch (e) { prevCalcMode = null; }

      try {
        var allSheets = wb.worksheets.items;
        var masterWs = allSheets.filter(function (w) { return w.name.toLowerCase() === MASTER.toLowerCase(); })[0];
        if (!masterWs) throw new Error('Sheet "' + MASTER + '" not found - is this the original, unconverted BOM workbook?');

        // Sync 2: Load tables on every sheet
        allSheets.forEach(function (ws) { ws.tables.load("items/name"); });
        await context.sync();

        var tblMat = masterWs.tables.items.filter(function (t) { return t.name === "tblMat"; })[0];
        if (!tblMat) throw new Error('Table "tblMat" not found on the Master sheet.');

        var sheets = filterBomSheets(allSheets);
        if (!sheets.length) throw new Error("No BOM sheet (table named bom_...) found in this workbook.");
        var convertedNames = sheets.map(function (s) { return s.name; });

        // Sync 3: Batch-load Master table AND A1:AK<BOTTOM> for ALL BOM sheets in ONE sync
        log("Reading Master list and " + sheets.length + " BOM sheet(s) ...");
        var matHdr = tblMat.getHeaderRowRange().load("values");
        var matBody = tblMat.getDataBodyRange().load("values");
        var sheetRanges = sheets.map(function (ws) {
          return ws.getRange("A1:AK" + BOTTOM).load(["values", "formulas"]);
        });
        await context.sync();

        var master = parseMasterFromRanges(matHdr.values, matBody.values);
        var snapshots = sheets.map(function (ws, idx) {
          var S = parseSheetRange(sheetRanges[idx].values, sheetRanges[idx].formulas);
          return { ws: ws, S: S, po: readPo(S, master) };
        });

        // Sync 4: Strip tables, clear CF/DV, unmerge AD2:AK2, delete row 5 across ALL sheets in ONE sync
        log("Flattening " + sheets.length + " sheet(s) (" + (mode === "growing" ? "auto-growing" : "fixed") + ") ...");
        pauseExcel(context);
        snapshots.forEach(function (snap) {
          queueStripTables(snap.ws, snap.S.title);
        });
        await context.sync();

        // Sync 5: Write BOM blocks, purchase lists, borders, CF, print areas & freeze panes across ALL sheets in ONE sync
        pauseExcel(context);
        snapshots.forEach(function (snap) {
          var w = mode === "growing"
            ? queueConvertSheetGrowing(snap.ws, snap.S, master)
            : queueConvertSheetFixed(snap.ws, snap.S, snap.po, master);
          w.forEach(function (x) { warnings.push(snap.ws.name + ": " + x); });
        });
        await context.sync();

        // Sync 6 (only when needed): Recalculate ONCE for growing-mode PO restoration and/or Compiled sheet
        var growingPoSnaps = mode === "growing"
          ? snapshots.filter(function (snap) { return Object.keys(snap.po).length > 0; })
          : [];
        var needsMidCalc = growingPoSnaps.length > 0 || !!opts.compiled;

        if (needsMidCalc) {
          wb.application.calculate(Excel.CalculationType.full);
          if (growingPoSnaps.length) {
            growingPoSnaps.forEach(function (snap) {
              snap.ae2 = snap.ws.getRange("AE2").load("values");
              snap.plNames = snap.ws.getRange("P5:P" + PL_LAST).load("values");
            });
          }
          await context.sync();

          if (growingPoSnaps.length) {
            pauseExcel(context);
            var wrotePo = false;
            growingPoSnaps.forEach(function (snap) {
              var nn = Math.min(CAP, Math.round(Number(snap.ae2.values[0][0])) || 0);
              for (var i = 0; i < nn; i++) {
                var pv = snap.po[String(snap.plNames.values[i][0]).trim()];
                if (!pv) continue;
                for (var j = 0; j < pv.length; j++) {
                  if (pv[j] !== undefined) {
                    snap.ws.getRange(colLetter(COLS_PO[j]) + (5 + i)).values = [[pv[j]]];
                    wrotePo = true;
                  }
                }
              }
            });
            if (wrotePo) await context.sync();
          }
        }

        if (opts.compiled) {
          log("Building the compiled purchase list sheet ...");
          try { await buildCompiled(context, wb, sheets, master, log, mode === "growing"); }
          catch (e) { warnings.push("Compiled sheet failed: " + (e.message || e)); }
        }

        log("Removing template sheets and named ranges ...");
        await cleanWorkbook(context, wb, convertedNames);
      } finally {
        await restoreCalcAndRecalc(context, wb, prevCalcMode);
      }
    });
    return warnings;
  }

  var api = {
    convertWorkbook: convertWorkbook,
    restoreAutomatedWorkbook: restoreAutomatedWorkbook,
    inspectWorkbook: inspectWorkbook,
    plFormulas: plFormulas,
    compileMaterials: compileMaterials,
    CAP: CAP,
    PL_LAST: PL_LAST
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomConvert = api;
})(typeof window !== "undefined" ? window : this);
