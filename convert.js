/* convert.js - both BOM converters, ported from the VBA macros to Office.js.
 *   mode "growing" = FlattenBOM.bas  (formula-driven purchase list that grows by itself)
 *   mode "fixed"   = FlattenBOM1.bas (purchase list written once, fixed rows)
 * Runs inside Excel (desktop / web / Mac). Nothing leaves the workbook.
 */
(function (root) {
  "use strict";

  var CAP = 25;
  var PL_LAST = 4 + CAP + 3;   // last purchase-list row carrying formulas (growing mode)
  var BOTTOM = 320;
  var KGM2 = "kg/m²";
  var MASTER = "Master Unique Materials";
  var COLS_PO = [22, 23, 24, 25, 27];      // V W X Y AA : typed stock / PO inputs

  function colLetter(n) {
    var s = "";
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function isBlank(v) { return v === "" || v === null || v === undefined; }
  function isFormula(f) { return typeof f === "string" && f.charAt(0) === "="; }

  /* ---------- formulas of the growing purchase list (same as the VBA PLF / Python pl_formulas) ---------- */
  function plFormulas(r, tot) {
    var rng = function (c) { return "$" + c + "$5:$" + c + "$" + tot; };
    var C = rng("C"), G = rng("G"), H = rng("H"), I = rng("I"), J = rng("J"), L = rng("L"), M = rng("M");
    var k = "(ROW()-4)", n = "$AE$2";
    var prev = r === 5 ? "-1" : "AB" + (r - 1);
    var bar = "SUMIF(" + C + ",P" + r + "," + L + ")>0";
    var gt = k + "=" + n + "+1", rq = k + "=" + n + "+3", item = k + "<=" + n, after = k + ">" + n;
    var sumAbove = function (c) { return r > 5 ? "SUM(" + c + "$5:" + c + (r - 1) + ")" : '""'; };
    var cntAbove = function (c) {
      return r > 5 ? 'IF(COUNT(' + c + '$5:' + c + (r - 1) + ')=0,"",SUM(' + c + '$5:' + c + (r - 1) + '))' : '""';
    };
    var f = {};
    f.O = '=IF(' + item + ',' + k + ',"")';
    f.AB = '=IF(' + item + ',IFERROR(SMALL(' + M + ',COUNTIF(' + M + ',"<="&' + prev + ')+1),""),"")';
    f.P = '=IF(' + item + ',IFERROR(INDEX(' + C + ',MATCH(AB' + r + ',' + M + ',0)),""),' +
          'IF(' + gt + ',"Grand Total Weight in KG :",IF(' + rq + ',"Requested by:","")))';
    f.Q = '=IF(' + rq + ',"Procurement",IF(' + after + ',"",IF(' + bar + ',$AE$1,"")))';
    f.R = '=IF(' + after + ',"",IF(' + bar + ',ROUNDUP(SUMIF(' + C + ',P' + r + ',' + L + ')/Q' + r + ',0),' +
          'SUMPRODUCT((' + C + '=P' + r + ')*' + G + '*' + H + ')))';
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
    edges.forEach(function (e) {
      var b = range.format.borders.getItem(e);
      b.style = on ? "Continuous" : "None";
      if (on) b.weight = "Thin";
    });
  }
  function put(ws, addr, v) { ws.getRange(addr).formulas = [[v]]; }   // .formulas accepts plain values too

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
    var headers = hdr.values[0].map(function (s) { return String(s).trim(); });
    var col = function (prefix) {
      var i = headers.findIndex(function (h) { return h.indexOf(prefix) === 0; });
      if (i < 0) throw new Error('Master table has no column starting with "' + prefix + '".');
      return i;
    };
    var cM = col("Material"), cLin = col("Linear Weight"), cAr = col("Area Weight"), cPc = col("Piece Weight");
    var cU = col("Primary Unit"), cSL = col("Sheet Length"), cSW = col("Sheet Width");
    var m = {};
    body.values.forEach(function (row, i) {
      var name = row[cM];
      if (isBlank(name)) return;
      m[String(name)] = { order: i, unit: row[cU], lin: row[cLin], area: row[cAr], piece: row[cPc], sl: row[cSL], sw: row[cSW] };
    });
    return m;
  }

  async function findBomSheets(context, wb) {
    wb.worksheets.load("items/name");
    await context.sync();
    var sheets = wb.worksheets.items;
    sheets.forEach(function (ws) { ws.tables.load("items/name"); });
    await context.sync();
    return sheets.filter(function (ws) {
      return ws.tables.items.some(function (t) { return t.name.indexOf("bom_") === 0; });
    });
  }

  /* ---------- steps shared by both modes ---------- */
  async function readSheet(context, ws) {
    var rg = ws.getRange("A1:AK" + BOTTOM);
    rg.load(["values", "formulas"]);
    await context.sync();
    var val = function (r, c) { return rg.values[r - 1][c - 1]; };
    var frm = function (r, c) { return rg.formulas[r - 1][c - 1]; };
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
    for (var r = 6; r <= 40; r++) {
      var mat = String(S.val(r, 16));
      if (mat !== "" && master[mat]) {
        po[mat] = COLS_PO.map(function (c) {
          return (isFormula(S.frm(r, c)) || S.val(r, c) === "") ? undefined : S.val(r, c);
        });
      }
    }
    return po;
  }

  async function stripTables(context, ws) {
    ws.tables.load("items/name");
    await context.sync();
    ws.tables.items.forEach(function (t) { t.convertToRange(); });
    ws.getRange().conditionalFormats.clearAll();
    ws.getRange().dataValidation.clear();
    ws.getRange("AD2:AK2").unmerge();
    ws.getRange("5:5").delete("Up");                 // hidden header row
    await context.sync();
  }

  /* BOM block: identical in both modes except column M helper (growing only) and row-5-hidden handling */
  function writeBomBlock(ws, S, master, o) {
    var first = o.first, last = o.last, tot = o.tot, missing = [];
    var kgm2 = KGM2;
    for (var r = first; r <= last; r++) {
      var oldRow = r + 1;                                   // row index in the OLD layout (before row 5 was deleted)
      var mat = S.val(oldRow, 3), desc = S.val(oldRow, 2);
      mat = isBlank(mat) ? "" : String(mat);
      desc = isBlank(desc) ? "" : String(desc);
      if (o.helperM) put(ws, "M" + r, '=IF(C' + r + '="","",COUNTIF($C$' + first + ':$C$' + tot + ',"<"&C' + r + '))');
      if (mat === "") {
        if (desc !== "") ws.getRange("A" + r + ":L" + r).format.font.bold = true;
        continue;
      }
      put(ws, "A" + r, '=IF(C' + r + '="","",COUNTA($C$' + first + ':C' + r + '))');
      var info = master[mat];
      if (!info) { missing.push("row " + (r) + ": '" + mat + "' is not in the Master - weight left blank"); continue; }
      if (isBlank(S.val(oldRow, 8)) && !isBlank(S.val(oldRow, 7))) put(ws, "H" + r, 1);
      if (info.unit === "kg/m") {
        put(ws, "I" + r, info.lin);
        put(ws, "J" + r, "=I" + r + "*G" + r + "*H" + r + "*D" + r + "/1000");
        put(ws, "L" + r, '=IF(ISNUMBER(D' + r + '),D' + r + '*G' + r + '*H' + r + ',"")');
      } else if (info.unit === kgm2) {
        put(ws, "I" + r, info.area);
        put(ws, "J" + r, "=I" + r + "*G" + r + "*H" + r + "*D" + r + "*E" + r + "/10^6");
      } else {
        put(ws, "I" + r, info.piece);
        put(ws, "J" + r, "=I" + r + "*G" + r + "*H" + r);
      }
    }
    return missing;
  }

  /* ---------- MODE 1: growing purchase list (FlattenBOM.bas) ---------- */
  async function convertSheetGrowing(context, ws, master, log) {
    var S = await readSheet(context, ws);
    var po = readPo(S, master);
    ws.getRange("O2").values = [[S.title]];               // title formula -> text
    await context.sync();

    await stripTables(context, ws);

    var first = 5, last = S.lastOld - 1, tot = last + 1;
    ws.getRange("M:M").clear();
    ws.getRange("AB:AC").clear();
    ws.getRange("AF:AK").clear();
    ws.getRange("AD2:AE" + BOTTOM).clear();
    ws.getRange("M:M").columnHidden = true;
    ws.getRange("AB:AB").columnHidden = true;
    ws.getRange("AC:AC").format.columnWidth = 14;
    if (tot + 1 <= BOTTOM) ws.getRange("A" + (tot + 1) + ":L" + BOTTOM).clear();
    ws.getRange("A" + tot + ":L" + tot).clear("Contents");
    setBorders(ws.getRange("A" + first + ":L" + last), false);
    ws.getRange("O" + (PL_LAST + 1) + ":AA" + BOTTOM).clear();
    ws.getRange("O5:AA" + PL_LAST).clear("Contents");
    setBorders(ws.getRange("O5:AA" + PL_LAST), false);
    ws.getRange("A" + first + ":A" + last).clear("Contents");
    ws.getRange("I" + first + ":J" + last).clear("Contents");
    ws.getRange("L" + first + ":L" + last).clear("Contents");
    await context.sync();

    context.application.suspendApiCalculationUntilNextSync();
    setBorders(ws.getRange("A" + first + ":L" + last), true);
    var missing = writeBomBlock(ws, S, master, { first: first, last: last, tot: tot, helperM: true });
    setBorders(ws.getRange("A" + tot + ":L" + tot), true, true);
    put(ws, "I" + tot, "Total:");
    put(ws, "J" + tot, "=SUM(J$" + first + ":INDEX(J:J,ROW()-1))");
    ws.getRange("I" + tot + ":J" + tot).format.font.bold = true;

    for (var r = 5; r <= PL_LAST; r++) {
      var f = plFormulas(r, tot);
      Object.keys(f).forEach(function (c) { put(ws, c + r, f[c]); });
    }
    ws.getRange("AD2").copyFrom("AD1", "Formats");
    ws.getRange("AE2").copyFrom("AE1", "Formats");
    ws.getRange("AE2").format.fill.clear();
    put(ws, "AE1", S.stdLen);
    put(ws, "AD2", "Materials in purchase list:");
    put(ws, "AE2", '=ROUND(SUMPRODUCT(($C$5:$C$' + tot + '<>"")/COUNTIF($C$5:$C$' + tot + ',$C$5:$C$' + tot + '&"")),0)');
    put(ws, "AD3", '=IF(AE2>' + CAP + ',"More than ' + CAP + ' materials - copy the last purchase-list row down","")');
    ws.getRange("AD3").format.font.bold = true;
    ws.getRange("AD3").format.font.color = "#C00000";

    // conditional format: borders + bold follow the list length
    var cf = ws.getRange("O5:AA" + PL_LAST).conditionalFormats.add("Custom");
    cf.custom.rule.formula = "=ROW()-4<=$AE$2";
    ["EdgeLeft", "EdgeRight", "EdgeTop", "EdgeBottom"].forEach(function (e) { cf.custom.format.borders.getItem(e).style = "Continuous"; });
    cf.stopIfTrue = false;
    for (var c = 15; c <= 27; c++) {
      var L = colLetter(c);
      var cf2 = ws.getRange(L + "5:" + L + PL_LAST).conditionalFormats.add("Custom");
      cf2.custom.rule.formula = "=ROW()-4=$AE$2+1";
      cf2.custom.format.borders.getItem("EdgeTop").style = "Continuous";
      cf2.custom.format.borders.getItem("EdgeBottom").style = "Continuous";
      if (c !== 16 && c !== 17 && c !== 18) cf2.custom.format.borders.getItem("EdgeLeft").style = "Continuous";
      if (c !== 15 && c !== 16 && c !== 17) cf2.custom.format.borders.getItem("EdgeRight").style = "Continuous";
      if ([16, 19, 21, 23, 26].indexOf(c) >= 0) cf2.custom.format.font.bold = true;
      cf2.stopIfTrue = false;
    }

    ws.pageLayout.setPrintArea("A1:AA" + Math.max(tot, PL_LAST));
    ws.freezePanes.unfreeze();
    ws.freezePanes.freezeRows(4);
    await context.sync();

    // put back typed stock / PO values next to their material
    context.workbook.application.calculate("Full");
    var ae2 = ws.getRange("AE2").load("values");
    await context.sync();
    var nn = Math.round(Number(ae2.values[0][0])) || 0;
    if (nn > 0) {
      var names = ws.getRange("P5:P" + (4 + nn)).load("values");
      await context.sync();
      names.values.forEach(function (row, i) {
        var pv = po[String(row[0])];
        if (!pv) return;
        pv.forEach(function (v, j) { if (v !== undefined) ws.getRange(colLetter(COLS_PO[j]) + (5 + i)).values = [[v]]; });
      });
      await context.sync();
    }
    return missing;
  }

  /* ---------- MODE 2: fixed purchase list (FlattenBOM1.bas) ---------- */
  async function convertSheetFixed(context, ws, master, log) {
    var S = await readSheet(context, ws);
    var po = readPo(S, master);
    ws.getRange("O2").values = [[S.title]];
    await context.sync();

    // materials used in the BOM, in Master order
    var used = {};
    for (var r0 = 6; r0 <= S.lastOld; r0++) {
      var m0 = S.val(r0, 3);
      if (!isBlank(m0) && master[String(m0)]) used[String(m0)] = master[String(m0)].order;
    }
    var names = Object.keys(used).sort(function (a, b) { return used[a] - used[b]; });
    var n = names.length;

    await stripTables(context, ws);

    var first = 5, last = S.lastOld - 1, tot = last + 1, gt = 5 + n, rq = 7 + n;
    ws.getRange("M:M").clear();
    ws.getRange("AB:AK").clear();
    ws.getRange("M:M").format.columnWidth = 14; ws.getRange("M:M").columnHidden = false;
    ws.getRange("AB:AB").format.columnWidth = 14; ws.getRange("AB:AB").columnHidden = false;
    if (tot + 1 <= BOTTOM) ws.getRange("A" + (tot + 1) + ":L" + BOTTOM).clear();
    ws.getRange("A" + tot + ":L" + tot).clear("Contents");
    setBorders(ws.getRange("A" + first + ":L" + last), false);
    if (rq + 1 <= BOTTOM) ws.getRange("O" + (rq + 1) + ":AA" + BOTTOM).clear();
    ws.getRange("O5:AA" + rq).clear("Contents");
    ws.getRange("A" + first + ":A" + last).clear("Contents");
    ws.getRange("I" + first + ":J" + last).clear("Contents");
    ws.getRange("L" + first + ":L" + last).clear("Contents");
    await context.sync();

    context.application.suspendApiCalculationUntilNextSync();
    setBorders(ws.getRange("A" + first + ":L" + last), true);
    var missing = writeBomBlock(ws, S, master, { first: first, last: last, tot: tot, helperM: false });
    setBorders(ws.getRange("A" + tot + ":L" + tot), true, true);
    put(ws, "I" + tot, "Total:");
    put(ws, "J" + tot, "=SUM(J" + first + ":J" + last + ")");
    ws.getRange("I" + tot + ":J" + tot).format.font.bold = true;

    var Cr = "$C$" + first + ":$C$" + last, Jr = "$J$" + first + ":$J$" + last, Lr = "$L$" + first + ":$L$" + last;
    var Ir = "$I$" + first + ":$I$" + last, Gr = "$G$" + first + ":$G$" + last, Hr = "$H$" + first + ":$H$" + last;
    for (var i = 1; i <= n; i++) {
      var r = 4 + i, mat = names[i - 1], info = master[mat];
      setBorders(ws.getRange("O" + r + ":AA" + r), true);
      put(ws, "O" + r, '=IF(P' + r + '="","",COUNTA(P$' + first + ':P' + r + '))');
      put(ws, "P" + r, mat);
      put(ws, "S" + r, "=SUMIF(" + Cr + ",P" + r + "," + Jr + ")");
      if (info.unit === "kg/m") {
        put(ws, "Q" + r, S.stdLen);
        put(ws, "R" + r, "=ROUNDUP(SUMIF(" + Cr + ",P" + r + "," + Lr + ")/Q" + r + ",0)");
        put(ws, "T" + r, "=INDEX(" + Ir + ",MATCH(P" + r + "," + Cr + ",0))");
        put(ws, "U" + r, "=Q" + r + "/1000*R" + r + "*T" + r);
      } else if (info.unit === KGM2) {
        put(ws, "Q" + r, info.sl + " x " + info.sw);
        put(ws, "R" + r, "=ROUNDUP(S" + r + "/(" + info.area + "*" + info.sl + "*" + info.sw + "/10^6),0)");
        put(ws, "U" + r, "=R" + r + "*" + info.area + "*" + info.sl + "*" + info.sw + "/10^6");
      } else {
        put(ws, "R" + r, "=SUMPRODUCT((" + Cr + "=P" + r + ")*" + Gr + "*" + Hr + ")");
        put(ws, "U" + r, "=R" + r + "*" + info.piece);
      }
      var pv = po[mat];
      if (pv) pv.forEach(function (v, j) { if (v !== undefined) put(ws, colLetter(COLS_PO[j]) + r, v); });
      put(ws, "Z" + r, '=IF(Y' + r + '="","",W' + r + '*Y' + r + ')');
    }

    setBorders(ws.getRange("O" + gt + ":R" + gt), true, true);
    for (var c = 19; c <= 27; c++) setBorders(ws.getRange(colLetter(c) + gt), true, true);
    put(ws, "P" + gt, "Grand Total Weight in KG :");
    put(ws, "S" + gt, "=SUM(S5:S" + (gt - 1) + ")");
    put(ws, "U" + gt, "=SUM(U5:U" + (gt - 1) + ")");
    put(ws, "W" + gt, '=IF(COUNT(W5:W' + (gt - 1) + ')=0,"",SUM(W5:W' + (gt - 1) + '))');
    put(ws, "Z" + gt, '=IF(COUNT(Z5:Z' + (gt - 1) + ')=0,"",SUM(Z5:Z' + (gt - 1) + '))');
    ["P", "S", "U", "W", "Z"].forEach(function (L) { ws.getRange(L + gt).format.font.bold = true; });
    put(ws, "P" + rq, "Requested by:"); put(ws, "Q" + rq, "Procurement");
    put(ws, "T" + rq, "Store Section:"); put(ws, "W" + rq, "Production");
    put(ws, "Z" + rq, "Engineering");

    ws.pageLayout.setPrintArea("A1:AA" + Math.max(tot, rq));
    ws.freezePanes.unfreeze();
    ws.freezePanes.freezeRows(4);
    await context.sync();
    return missing;
  }

  /* ---------- workbook level ---------- */
  async function cleanWorkbook(context, wb) {
    wb.names.load("items/name");
    await context.sync();
    wb.names.items.forEach(function (nm) { if (nm.name.indexOf("Print_") < 0) nm.delete(); });
    try { await context.sync(); } catch (e) { /* some names are protected - ignore */ }
    var drop = ["Master Unique Materials", "Formulas", "README"];
    for (var i = 0; i < drop.length; i++) {
      var sh = wb.worksheets.getItemOrNullObject(drop[i]);
      await context.sync();
      if (!sh.isNullObject) { sh.delete(); await context.sync(); }
    }
  }

  async function inspectWorkbook() {
    var out = { master: false, bomSheets: [] };
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var m = wb.worksheets.getItemOrNullObject(MASTER);
      await context.sync();
      out.master = !m.isNullObject;
      var sheets = await findBomSheets(context, wb);
      out.bomSheets = sheets.map(function (s) { return s.name; });
    });
    return out;
  }

  async function convertWorkbook(mode, log) {
    log = log || function () {};
    var warnings = [];
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var master = await readMaster(context, wb);
      var sheets = await findBomSheets(context, wb);
      if (!sheets.length) throw new Error("No BOM sheet (table named bom_...) found in this workbook.");
      for (var i = 0; i < sheets.length; i++) {
        var ws = sheets[i];
        log("Converting " + ws.name + " ...");
        var w = mode === "growing"
          ? await convertSheetGrowing(context, ws, master, log)
          : await convertSheetFixed(context, ws, master, log);
        w.forEach(function (x) { warnings.push(ws.name + ": " + x); });
      }
      log("Removing Master / Formulas / README sheets and named ranges ...");
      await cleanWorkbook(context, wb);
    });
    return warnings;
  }

  var api = { convertWorkbook: convertWorkbook, inspectWorkbook: inspectWorkbook, plFormulas: plFormulas, CAP: CAP, PL_LAST: PL_LAST };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomConvert = api;
})(typeof window !== "undefined" ? window : this);
