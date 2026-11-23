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
  var BLANK = "BOM Blank";
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
      var key = String(name).trim();
      var info = { order: i, unit: row[cU], lin: row[cLin], area: row[cAr], piece: row[cPc], sl: row[cSL], sw: row[cSW] };
      m[String(name)] = info;
      if (key !== String(name)) m[key] = info;
    });
    return m;
  }

  /* Find active BOM sheets (skipping the hidden "BOM Blank" template sheet when real take-off sheets exist). */
  async function findBomSheets(context, wb) {
    wb.worksheets.load("items/name,items/visibility");
    await context.sync();
    var sheets = wb.worksheets.items;
    sheets.forEach(function (ws) { ws.tables.load("items/name"); });
    await context.sync();
    var allBom = sheets.filter(function (ws) {
      return ws.tables.items.some(function (t) { return t.name.indexOf("bom_") === 0; });
    });
    var realBom = allBom.filter(function (ws) {
      return !(ws.name === BLANK && ws.visibility !== "Visible");
    });
    return realBom.length ? realBom : allBom;
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
      var mat = String(S.val(r, 16)).trim();
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
      var info = master[mat] || master[mat.trim()];
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
        var pv = po[String(row[0]).trim()];
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
      var mKey = isBlank(m0) ? "" : String(m0).trim();
      if (mKey && master[mKey]) used[mKey] = master[mKey].order;
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

  /* ---------- EXTRA: compiled purchase list sheet ---------- */
  var COMPILED = "Compiled Purchase List";
  var PAD = 4;                  // empty rows on top for company name etc.
  var BLOCK = 13;               // purchase list is O..AA = 13 columns
  var STRIDE = BLOCK + 1;       // one gap column between lists
  var num = function (x) { return typeof x === "number" ? x : 0; };

  async function readPurchaseData(context, ws) {
    var pl = ws.getRange("O1:AA60").load("values");
    var bom = ws.getRange("A5:L" + BOTTOM).load("values");
    await context.sync();
    var v = pl.values, items = [], gt = 0, rq = 0, i;
    for (i = 4; i < v.length; i++) {
      var nm = String(v[i][1]);
      if (nm.indexOf("Grand Total") === 0) { gt = i + 1; break; }
      if (nm !== "") items.push({ mat: nm, Q: v[i][2], R: v[i][3], S: v[i][4], T: v[i][5], U: v[i][6] });
    }
    if (!gt) throw new Error(ws.name + ": purchase list has no Grand Total row.");
    for (i = gt; i < Math.min(v.length, gt + 6); i++) { if (String(v[i][1]) === "Requested by:") { rq = i + 1; break; } }
    if (!rq) rq = gt + 2;
    var L = {}, tot = 0;
    bom.values.forEach(function (row, idx) {
      var mat = String(row[2]);
      if (String(row[8]) === "Total:") tot = 5 + idx;
      if (mat !== "" && typeof row[11] === "number") L[mat] = (L[mat] || 0) + row[11];
    });
    if (!tot) tot = 5 + bom.values.length;
    return { name: ws.name, vals: v, items: items, gt: gt, rq: rq, L: L, tot: tot, ws: ws };
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
    try { fn(); await context.sync(); }
    catch (e) {
      var loc = e && e.debugInfo && e.debugInfo.errorLocation ? " [" + e.debugInfo.errorLocation + "]" : "";
      throw new Error(name + ": " + ((e && e.message) || e) + loc);
    }
  }
  async function soft(context, log, name, fn) {
    try { fn(); await context.sync(); }
    catch (e) { log("Note: skipped '" + name + "' (" + ((e && e.message) || e) + ")"); }
  }
  var qs = function (name) { return "'" + String(name).replace(/'/g, "''") + "'"; };

  function blockBorders(cs, c0, n) {
    var L = function (c) { return colLetter(c0 + c); };
    if (n > 0) setBorders(cs.getRange(L(0) + (PAD + 5) + ":" + L(BLOCK - 1) + (PAD + 4 + n)), true);
    var gr = PAD + 5 + n;
    setBorders(cs.getRange(L(0) + gr + ":" + L(3) + gr), true, true);
    for (var c = 4; c < BLOCK; c++) setBorders(cs.getRange(L(c) + gr), true, true);
  }

  function addGrowCf(cs, c0, r1, r2, nAddr) {
    var off = PAD + 4, L = function (c) { return colLetter(c0 + c); };
    var cf = cs.getRange(L(0) + r1 + ":" + L(BLOCK - 1) + r2).conditionalFormats.add("Custom");
    cf.custom.rule.formula = "=ROW()-" + off + "<=" + nAddr;
    ["EdgeLeft", "EdgeRight", "EdgeTop", "EdgeBottom"].forEach(function (e) { cf.custom.format.borders.getItem(e).style = "Continuous"; });
    cf.stopIfTrue = false;
    for (var c = 0; c < BLOCK; c++) {
      var sc = 15 + c;
      var cf2 = cs.getRange(L(c) + r1 + ":" + L(c) + r2).conditionalFormats.add("Custom");
      cf2.custom.rule.formula = "=ROW()-" + off + "=" + nAddr + "+1";
      cf2.custom.format.borders.getItem("EdgeTop").style = "Continuous";
      cf2.custom.format.borders.getItem("EdgeBottom").style = "Continuous";
      if (sc !== 16 && sc !== 17 && sc !== 18) cf2.custom.format.borders.getItem("EdgeLeft").style = "Continuous";
      if (sc !== 15 && sc !== 16 && sc !== 17) cf2.custom.format.borders.getItem("EdgeRight").style = "Continuous";
      if ([16, 19, 21, 23, 26].indexOf(sc) >= 0) cf2.custom.format.font.bold = true;
      cf2.stopIfTrue = false;
    }
  }

  async function growingBody(context, cs, datas, master, ctx) {
    var nL = datas.length, src = datas[0].ws;
    var capMain = Math.min(nL * CAP, 120), mainLast = PAD + 4 + capMain + 3;
    var lastListEnd = ctx.lastListEnd;
    var hs = lastListEnd + 3;
    var HN = colLetter(hs), HQ = colLetter(hs + 1), HT = colLetter(hs + 2), HS = colLetter(hs + 3), HR = colLetter(hs + 4),
        HU = colLetter(hs + 5), HL = colLetter(hs + 6), HW = colLetter(hs + 7), HK = colLetter(hs + 8),
        OT1 = colLetter(hs + 10), OT2 = colLetter(hs + 11), NC = colLetter(hs + 12);
    var hEnd = 4 + nL * CAP;
    var rg = function (c) { return "$" + c + "$5:$" + c + "$" + hEnd; };
    var rHN = rg(HN), rHQ = rg(HQ), rHT = rg(HT), rHS = rg(HS), rHR = rg(HR), rHU = rg(HU), rHL = rg(HL), rHW = rg(HW), rHK = rg(HK);
    var nMain = "$" + NC + "$5";

    await step(context, "helper block", function () {
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
            key]);
        }
      });
      cs.getRange(HN + "5:" + HK + hEnd).formulas = rows;
      var ord = Object.keys(master).map(function (k) { return [k, master[k].order]; });
      if (ord.length) cs.getRange(OT1 + "5:" + OT2 + (4 + ord.length)).values = ord;
      cs.getRange(NC + "5").formulas = [["=COUNT(" + rHK + ")"]];
      datas.forEach(function (d, s) { cs.getRange(NC + (6 + s)).formulas = [["=" + qs(d.name) + "!$AE$2"]]; });
    });

    await step(context, "compiled list formulas", function () {
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
    });

    await step(context, "side lists", function () {
      datas.forEach(function (d, k) {
        var c0 = 1 + STRIDE * (k + 1), q = qs(d.name), out = [];
        for (var r = 5; r <= PL_LAST; r++) {
          var row = [];
          for (var c = 0; c < BLOCK; c++) { var a = q + "!" + colLetter(15 + c) + r; row.push('=IF(' + a + '="","",' + a + ')'); }
          out.push(row);
        }
        cs.getRange(colLetter(c0) + (PAD + 5) + ":" + colLetter(c0 + BLOCK - 1) + (PAD + PL_LAST)).formulas = out;
      });
    });

    await step(context, "formats", function () {
      for (var r = PAD + 5; r <= mainLast; r++) cs.getRange("A" + r + ":M" + r).copyFrom(src.getRange("O5:AA5"), "Formats");
      datas.forEach(function (d, k) {
        var c0 = 1 + STRIDE * (k + 1);
        cs.getRange(colLetter(c0) + (PAD + 5) + ":" + colLetter(c0 + BLOCK - 1) + (PAD + PL_LAST)).copyFrom(d.ws.getRange("O5:AA" + PL_LAST), "Formats");
      });
    });

    await step(context, "grow-with-list formatting", function () {
      cs.getRange().conditionalFormats.clearAll();
      addGrowCf(cs, 1, PAD + 5, mainLast, nMain);
      datas.forEach(function (d, k) { addGrowCf(cs, 1 + STRIDE * (k + 1), PAD + 5, PAD + PL_LAST, "$" + NC + "$" + (6 + k)); });
      cs.getRange(HN + ":" + colLetter(hs + 12)).columnHidden = true;
    });
    return { lastRow: mainLast };
  }

  async function staticBody(context, cs, datas, master, ctx) {
    var rows = compileMaterials(datas, master), n = rows.length, first = datas[0], src = first.ws;
    var gr = PAD + 5 + n, rr = gr + 2;
    await step(context, "compiled list formats", function () {
      for (var i = 0; i < n; i++) cs.getRange("A" + (PAD + 5 + i) + ":M" + (PAD + 5 + i)).copyFrom(src.getRange("O5:AA5"), "Formats");
      cs.getRange("A" + gr + ":M" + gr).copyFrom(src.getRange("O" + first.gt + ":AA" + first.gt), "Formats");
      cs.getRange("A" + rr + ":M" + rr).copyFrom(src.getRange("O" + first.rq + ":AA" + first.rq), "Formats");
    });
    await step(context, "compiled list values", function () {
      var vals = rows.map(function (r, i) { return [i + 1, r.mat, r.Q, r.R, r.S, r.T, r.U, "", "", "", "", "", ""]; });
      if (n > 0) cs.getRange("A" + (PAD + 5) + ":M" + (PAD + 4 + n)).values = vals;
      var tot = function (k) { return Math.round(rows.reduce(function (a, r) { return a + num(r[k]); }, 0) * 1000) / 1000; };
      cs.getRange("A" + gr + ":M" + gr).values = [["", "Grand Total Weight in KG :", "", "", tot("S"), "", tot("U"), "", "", "", "", "", ""]];
      cs.getRange("A" + rr + ":M" + rr).values = [first.vals[first.rq - 1].map(function (x) { return x === null ? "" : x; })];
    });
    await step(context, "side lists", function () {
      datas.forEach(function (d, k) {
        var c0 = 1 + STRIDE * (k + 1), a = colLetter(c0), b = colLetter(c0 + BLOCK - 1);
        cs.getRange(a + (PAD + 5) + ":" + b + (PAD + d.rq)).copyFrom(d.ws.getRange("O5:AA" + d.rq), "Formats");
        cs.getRange(a + (PAD + 5) + ":" + b + (PAD + d.rq)).values =
          d.vals.slice(4, d.rq).map(function (r) { return r.map(function (x) { return x === null ? "" : x; }); });
      });
    });
    await step(context, "borders", function () {
      cs.getRange().conditionalFormats.clearAll();
      blockBorders(cs, 1, n);
      datas.forEach(function (d, k) { blockBorders(cs, 1 + STRIDE * (k + 1), d.items.length); });
    });
    return { lastRow: Math.max(rr, PAD + Math.max.apply(null, datas.map(function (d) { return d.rq; }))) };
  }

  async function buildCompiled(context, wb, sheets, master, log, growing) {
    var datas = [];
    for (var s = 0; s < sheets.length; s++) datas.push(await readPurchaseData(context, sheets[s]));
    var nL = datas.length, first = datas[0], src = first.ws;

    var old = wb.worksheets.getItemOrNullObject(COMPILED);
    await context.sync();
    if (!old.isNullObject) { old.delete(); await context.sync(); }
    var cs = wb.worksheets.add(COMPILED);
    await context.sync();
    context.application.suspendApiCalculationUntilNextSync();

    var srcCols = [];
    for (var c = 0; c < BLOCK; c++) {
      var cr = src.getRange(colLetter(15 + c) + ":" + colLetter(15 + c));
      cr.format.load("columnWidth");
      srcCols.push(cr);
    }
    await context.sync();
    var widths = srcCols.map(function (r) { return r.format.columnWidth; });
    var lastListEnd = 1 + STRIDE * nL + BLOCK - 1;
    await step(context, "column widths", function () {
      var setW = function (c0) { for (var c = 0; c < BLOCK; c++) cs.getRange(colLetter(c0 + c) + ":" + colLetter(c0 + c)).format.columnWidth = widths[c]; };
      setW(1);
      for (var k = 0; k <= nL; k++) { if (k > 0) setW(1 + STRIDE * k); cs.getRange(colLetter(1 + STRIDE * k + BLOCK) + ":" + colLetter(1 + STRIDE * k + BLOCK)).format.columnWidth = 18; }
    });

    await step(context, "header rows", function () {
      cs.getRange("A" + (PAD + 1) + ":M" + (PAD + 4)).copyFrom(src.getRange("O1:AA4"), "All");
      datas.forEach(function (d, k) {
        var c0 = 1 + STRIDE * (k + 1);
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
      : await staticBody(context, cs, datas, master, { lastListEnd: lastListEnd });

    await soft(context, log, "landscape", function () { cs.pageLayout.orientation = "Landscape"; });
    await soft(context, log, "print area", function () { cs.pageLayout.setPrintArea("A1:" + colLetter(lastListEnd) + res.lastRow); });
    await soft(context, log, "page breaks", function () {
      for (var p = 0; p < nL; p++) cs.verticalPageBreaks.add(cs.getRange(colLetter(1 + STRIDE * (p + 1)) + "1"));
    });
    await soft(context, log, "activate", function () { cs.activate(); });
    log("Compiled sheet created (" + (growing ? "grows with the BOM" : "values only") + ") from " + nL + " sheet(s).");
  }

  /* ---------- workbook level ---------- */
  async function cleanWorkbook(context, wb, convertedNames) {
    wb.names.load("items/name");
    await context.sync();
    wb.names.items.forEach(function (nm) { if (nm.name.indexOf("Print_") < 0) nm.delete(); });
    try { await context.sync(); } catch (e) { /* some names are protected - ignore */ }
    var drop = ["Master Unique Materials", "Formulas", "README"];
    if (!convertedNames || convertedNames.indexOf(BLANK) < 0) drop.push(BLANK);
    for (var i = 0; i < drop.length; i++) {
      var sh = wb.worksheets.getItemOrNullObject(drop[i]);
      await context.sync();
      if (!sh.isNullObject) { sh.delete(); await context.sync(); }
    }
  }

  /* ---------- detect converted (flattened) BOM sheets ---------- */
  var SKIP_SHEETS = { "master unique materials": 1, "bom blank": 1, "compiled purchase list": 1, "formulas": 1, "readme": 1 };

  async function findConvertedSheets(context, wb) {
    wb.worksheets.load("items/name,items/visibility,items/position");
    await context.sync();
    var sheets = wb.worksheets.items;
    sheets.forEach(function (ws) { ws.tables.load("items/name"); });
    await context.sync();

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

  /* Extract all user-entered BOM rows, title metadata, bar length, and typed PO inputs from a converted BOM sheet. */
  async function extractConvertedSheetData(context, ws) {
    var rg = ws.getRange("A1:AE" + BOTTOM);
    rg.load(["values", "formulas"]);
    await context.sync();

    var vals = rg.values;
    var frms = rg.formulas;
    var val = function (r, c) { return vals[r - 1][c - 1]; };
    var frm = function (r, c) { return frms[r - 1][c - 1]; };

    var title = isBlank(val(1, 1)) ? "" : String(val(1, 1));
    var subtitle = isBlank(val(2, 1)) ? ws.name : String(val(2, 1));

    // Standard bar length: AE1 if present (growing mode), otherwise first numeric Q value in purchase list (fixed mode), else 6000
    var stdLen = 6000;
    if (typeof val(1, 31) === "number" && val(1, 31) > 0) {
      stdLen = val(1, 31);
    } else {
      for (var rq = 5; rq <= 45; rq++) {
        var qv = val(rq, 17);
        if (typeof qv === "number" && qv > 0) { stdLen = qv; break; }
      }
    }

    // Locate BOM data rows (start at row 5 in converted sheets, stop right before "Total:" in col I)
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

    // Extract typed Stock / PO inputs (V, W, X, Y, AA) per material from the converted purchase list (rows 5..55)
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
      name: ws.name,
      title: title,
      subtitle: subtitle,
      stdLen: stdLen,
      bomRows: bomRows,
      usedMaterials: Object.keys(usedMaterials),
      po: po
    };
  }

  /* Restore converted (plain) BOM sheets back into full automated BOM sheets with Master Unique Materials,
   * hidden BOM Blank, bom_* / pl_* tables, working dropdowns, formulas, and preserved Stock/PO inputs.
   */
  async function restoreAutomatedWorkbook(log, opts) {
    log = log || function () {};
    opts = opts || {};
    var T = root.BomTools;
    var S = root.BomSync;
    if (!T || !S) throw new Error("Required modules (tools.js / sync.js) are not loaded.");

    var warnings = [];
    var extracted = [];

    // 1. Inspect and extract all data from converted BOM sheets first
    var hasMaster = false, hasBlank = false;
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var convSheets = await findConvertedSheets(context, wb);
      if (!convSheets.length) {
        throw new Error("No converted BOM sheets found to restore. (Automated BOM sheets with bom_ tables are already automated.)");
      }
      var mWs = wb.worksheets.getItemOrNullObject(MASTER);
      var bWs = wb.worksheets.getItemOrNullObject(BLANK);
      await context.sync();
      hasMaster = !mWs.isNullObject;
      hasBlank = !bWs.isNullObject;

      for (var i = 0; i < convSheets.length; i++) {
        log("Reading converted sheet \"" + convSheets[i].name + "\" ...");
        extracted.push(await extractConvertedSheetData(context, convSheets[i]));
      }
    });

    // 2. Ensure Master Unique Materials and hidden BOM Blank exist (import cleaned template if missing)
    if (!hasMaster || !hasBlank) {
      if (hasMaster && !hasBlank) {
        // Backup existing master list before re-importing clean template so nothing is lost
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

    // 3. Optionally sync the latest local computer master list before restoring formulas
    if (opts.syncLocal !== false) {
      try {
        await S.pullLocal(log, { allowPrompt: true, syncRemovals: true });
      } catch (e) {
        log("Note: skipped local master file sync (" + (e.message || e) + ").");
      }
    }

    // 4. Rebuild each converted sheet using BOM Blank and copy all entered data + PO inputs back in
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var matInfo = await T._ensureMasterNames(context, wb, log);
      var masterLookup = await readMaster(context, wb);

      var blankWs = wb.worksheets.getItem(BLANK);
      blankWs.tables.load("items/name");
      await context.sync();
      var srcBom = blankWs.tables.items.filter(function (t) { return t.name.indexOf("bom_") === 0; })[0];
      var srcPl = blankWs.tables.items.filter(function (t) { return t.name.indexOf("pl_") === 0; })[0];
      if (!srcBom) throw new Error("Hidden \"BOM Blank\" sheet has no bom_ table.");

      var firstRestored = null;
      for (var i = 0; i < extracted.length; i++) {
        var data = extracted[i];
        log("Restoring \"" + data.name + "\" to an automated BOM sheet ...");

        // Warn if any material used on the converted sheet isn't in Master Unique Materials
        data.usedMaterials.forEach(function (matName) {
          if (!masterLookup[matName] && !masterLookup[matName.trim()]) {
            warnings.push(data.name + ": \"" + matName + "\" is not in Master Unique Materials (add it in the Materials tab so weights calculate).");
          }
        });

        var oldWs = wb.worksheets.getItem(data.name);
        wb.worksheets.load("items/name");
        await context.sync();
        var existingNames = wb.worksheets.items.map(function (w) { return w.name; });
        var tmpName = T.uniqueName("__tmp_" + i, existingNames).slice(0, 31);
        oldWs.name = tmpName;
        await context.sync();

        var restored = blankWs.copy(Excel.WorksheetPositionType.after, oldWs);
        restored.name = data.name;
        restored.visibility = Excel.SheetVisibility.visible;
        restored.tables.load("items/name");
        wb.tables.load("items/name");
        await context.sync();

        var bomT = restored.tables.items.filter(function (t) { return t.name.indexOf("bom_") === 0; })[0];
        var plT = restored.tables.items.filter(function (t) { return t.name.indexOf("pl_") === 0; })[0];
        var taken = wb.tables.items.map(function (t) { return t.name; });
        var id = T.safeId(data.name);
        var bomName = T.uniqueName("bom_" + id, taken);
        var plName = T.uniqueName("pl_" + id, taken.concat([bomName]));
        if (bomT) bomT.name = bomName;
        if (plT) plT.name = plName;
        await context.sync();

        if (data.title) restored.getRange("A1").values = [[data.title]];
        restored.getRange("A2").values = [[data.subtitle || data.name]];
        restored.getRange("AE1").values = [[data.stdLen || 6000]];
        await context.sync();

        // Populate BOM table input columns (B, C, D, E, F, G, H, K)
        if (bomT && data.bomRows.length) {
          var bomBody = bomT.getDataBodyRange().load("rowCount");
          await context.sync();
          var extraRows = data.bomRows.length - bomBody.rowCount;
          if (extraRows > 0) {
            for (var er = 0; er < extraRows; er++) bomT.rows.add();
            await context.sync();
          }
          var rowCount = data.bomRows.length;
          var endRow = 5 + rowCount; // data starts on row 6
          var colsBtoH = data.bomRows.map(function (r) {
            return [r.desc, r.mat, r.len, r.wid, r.thk, r.qty, r.set];
          });
          var colK = data.bomRows.map(function (r) { return [r.rem]; });
          restored.getRange("B6:H" + endRow).values = colsBtoH;
          restored.getRange("K6:K" + endRow).values = colK;
          await context.sync();
        }

        // Repair formulas, dropdowns (C6:C2000), hidden helper row/cols, freeze panes, and print setup
        await T._repairBomSheet(context, wb, restored, matInfo, {
          srcBomName: srcBom.name,
          srcPlName: srcPl ? srcPl.name : ""
        }, log);

        // Recalculate so automated Purchase List populates column P, then restore typed PO/Stock inputs
        wb.application.calculate(Excel.CalculationType.full);
        await context.sync();

        if (plT && Object.keys(data.po).length) {
          var plBody = plT.getDataBodyRange().load(["values", "rowIndex"]);
          await context.sync();
          var plStart = (plBody.rowIndex || 5) + 1; // 1-based Excel row (row 6)
          plBody.values.forEach(function (row, idx) {
            var matKey = String(row[1] || "").trim().toLowerCase(); // col P is 2nd column of pl_ (O=0, P=1)
            var pv = matKey ? data.po[matKey] : null;
            if (!pv) return;
            var excelRow = plStart + idx;
            pv.forEach(function (valToPut, j) {
              if (valToPut !== undefined) {
                restored.getRange(colLetter(COLS_PO[j]) + excelRow).values = [[valToPut]];
              }
            });
          });
          await context.sync();
        }

        oldWs.delete();
        await context.sync();
        if (!firstRestored) firstRestored = restored;
      }

      if (opts.removeCompiled) {
        var compWs = wb.worksheets.getItemOrNullObject(COMPILED);
        await context.sync();
        if (!compWs.isNullObject) { compWs.delete(); await context.sync(); }
      }

      wb.application.calculate(Excel.CalculationType.full);
      if (firstRestored) firstRestored.activate();
      await context.sync();
    });

    log("Restored " + extracted.length + " sheet(s) back to automated BOM format.");
    return { restoredCount: extracted.length, warnings: warnings };
  }

  async function inspectWorkbook() {
    var out = { master: false, bomSheets: [], convertedSheets: [] };
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var m = wb.worksheets.getItemOrNullObject(MASTER);
      await context.sync();
      out.master = !m.isNullObject;
      var sheets = await findBomSheets(context, wb);
      out.bomSheets = sheets.map(function (s) { return s.name; });
      var conv = await findConvertedSheets(context, wb);
      out.convertedSheets = conv.map(function (s) { return s.name; });
    });
    return out;
  }

  async function convertWorkbook(mode, log, opts) {
    log = log || function () {};
    opts = opts || {};
    var warnings = [];
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var master = await readMaster(context, wb);
      var sheets = await findBomSheets(context, wb);
      if (!sheets.length) throw new Error("No BOM sheet (table named bom_...) found in this workbook.");
      var convertedNames = sheets.map(function (s) { return s.name; });
      for (var i = 0; i < sheets.length; i++) {
        var ws = sheets[i];
        log("Converting " + ws.name + " ...");
        var w = mode === "growing"
          ? await convertSheetGrowing(context, ws, master, log)
          : await convertSheetFixed(context, ws, master, log);
        w.forEach(function (x) { warnings.push(ws.name + ": " + x); });
      }
      if (opts.compiled) {
        log("Building the compiled purchase list sheet ...");
        try { await buildCompiled(context, wb, sheets, master, log, mode === "growing"); }
        catch (e) { warnings.push("Compiled sheet failed: " + (e.message || e)); }
      }
      log("Removing template sheets and named ranges ...");
      await cleanWorkbook(context, wb, convertedNames);
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
