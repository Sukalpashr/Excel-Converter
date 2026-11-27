/* sheets.js — window.BomSheets
 * ----------------------------------------------------------------------------
 * Template import, lag-free sheet creation, the canonical Excel 365 formula
 * engine, Materials CRUD with audit stamping, and per-sheet repair.
 *
 * Performance contract (CRITICAL — see AGENTS rules):
 *   - NEVER register Worksheet.onChanged listeners. Direct edits on Master are
 *     detected on-demand via snapshot diffing (readMaterials) when the user
 *     visits Materials, refreshes, saves, or exports.
 *   - Every Excel.run is wrapped with { delayForCellEdit: true }.
 *   - Heavy ops switch to Excel.CalculationMode.manual and call
 *     suspendApiCalculationUntilNextSync() / suspendScreenUpdatingUntilNextSync();
 *     calculationMode is restored to automatic in `finally`.
 *   - Writes are batched as 2D arrays via .values / .formulas.
 *   - freezePanes / pageLayout / showGridlines are NEVER called on live editing
 *     sheets after creation.
 *
 * Excel 365 formulas injected at runtime (NOT baked into template.xlsx):
 *   Standalone cells: O2 (PL title), M4 (reverse XMATCH last-used row),
 *     AL6 (hidden SORT(UNIQUE(FILTER)) spill), AB4 (ROWS of spill),
 *     AD2 (status comparing BOM vs PL totals, flags missing Master materials).
 *   BOM columns A/I/J/L/M: structured-ref formulas with single-LET weight
 *     handling for kg/m · kg/m² · kg/each and a dynamic Total row at $M$4+1.
 *   PL columns O/P/Q/R/S/T/U/W/Z/AB: O(1) INDEX($AL$6#, ROW()-5) for Mat Row,
 *     grand-total + signature rows at $AB$4+6 / $AB$4+8.
 */
(function (root) {
  'use strict';

  var META = (root.BOM_TEMPLATE_META) || {};
  var BOM_TABLE = (META.bomTable && META.bomTable.name) || 'bom_Take_off';
  var PL_TABLE  = (META.plTable  && META.plTable.name)  || 'pl_Take_off';
  var MAT_TABLE = (META.matTable && META.matTable.name) || 'tblMat';
  var MASTER_SHEET = 'Master Unique Materials';
  var TEMPLATE_SHEET = 'BOM Blank';

  var NAMED_RANGES = [
    { name: 'Materials', refersTo: MAT_TABLE + '[Material]' },
    { name: 'LinWt',     refersTo: MAT_TABLE + '[Linear Weight (kg/m)]' },
    { name: 'AreaWt',    refersTo: MAT_TABLE + '[Area Weight (kg/m²)]' },
    { name: 'PieceWt',   refersTo: MAT_TABLE + '[Piece Weight (kg/each)]' },
    { name: 'WtUnit',    refersTo: MAT_TABLE + '[Primary Unit]' },
    { name: 'SheetL',    refersTo: MAT_TABLE + '[Sheet Length (mm)]' },
    { name: 'SheetW',    refersTo: MAT_TABLE + '[Sheet Width (mm)]' }
  ];

  // Reserved Excel sheet names (case-insensitive)
  var RESERVED = /^(history|logo|print_area|print_titles|sheet1|sheet2|sheet3|sheets)$/i;
  // Excel sheet-name rules: no : \ / ? * [ ]  , max 31 chars, cannot start/end with '
  function validateSheetName(name, existing) {
    if (!name) return 'Name is required';
    if (name.length > 31) return 'Max 31 characters';
    if (/[\\/?*[\]:]/.test(name)) return 'Cannot contain : \\ / ? * [ ]';
    if (name.length > 0 && name[0] === "'") return 'Cannot start with an apostrophe';
    if (name[name.length - 1] === "'") return 'Cannot end with an apostrophe';
    if (RESERVED.test(name)) return 'Reserved name';
    var ex = (existing || []).map(function (s) { return s.toLowerCase(); });
    if (ex.indexOf(name.toLowerCase()) >= 0) return 'A sheet with this name already exists';
    return null;
  }

  // ---- helpers ----
  function run(opts, body) {
    if (typeof opts === 'function') { body = opts; opts = {}; }
    opts = opts || {};
    var ctxOpts = { delayForCellEdit: true };
    return Excel.run(ctxOpts, function (ctx) {
      var app = ctx.workbook.application;
      var prevMode = null;
      var p = Promise.resolve();
      if (opts.heavy) {
        p = p.then(function () {
          prevMode = Excel.CalculationMode;       // capture enum reference
          return ctx.sync().then(function () {
            app.calculationMode = 'Manual';
            app.suspendApiCalculationUntilNextSync();
            app.suspendScreenUpdatingUntilNextSync();
          });
        });
      }
      return p.then(function () { return body(ctx); }).then(function (r) {
        if (opts.heavy) {
          // restore automatic calc on the same context (suspend* auto-end with the batch)
          app.calculationMode = 'Automatic';
        }
        return ctx.sync().then(function () { return r; });
      }, function (err) {
        if (opts.heavy) {
          try { app.calculationMode = 'Automatic'; } catch (e) {}
        }
        throw err;
      });
    });
  }

  function sheetExists(ctx, name) {
    var sh = ctx.workbook.worksheets.getItemOrNullObject(name);
    return ctx.sync().then(function () { return !sh.isNullObject; });
  }
  function listSheetNames(ctx) {
    var sh = ctx.workbook.worksheets.load('name');
    return ctx.sync().then(function () { return sh.items.map(function (s) { return s.name; }); });
  }
  function isBomSheetName(name) {
    // BOM sheets are any sheet that is NOT the Master, NOT BOM Blank, NOT flattened.
    // The repair/convert functions decide their own inclusion lists.
    return name !== MASTER_SHEET && name !== TEMPLATE_SHEET;
  }

  // ============================================================
  // TEMPLATE IMPORT
  // ============================================================

  function ensureTemplate() {
    return run(function (ctx) {
      var ws = ctx.workbook.worksheets;
      var bom = ws.getItemOrNullObject(TEMPLATE_SHEET);
      var mas = ws.getItemOrNullObject(MASTER_SHEET);
      return ctx.sync().then(function () {
        if (bom.isNullObject || mas.isNullObject) {
          return importTemplateIntoWorkbook(ctx).then(function () {
            return { imported: true };
          });
        }
        return { imported: false };
      });
    });
  }

  // Inserts the two template sheets from window.BOM_TEMPLATE_B64, then injects
  // the canonical Excel 365 formulas and recreates named ranges.
  function importTemplateIntoWorkbook(ctx) {
    var b64 = root.BOM_TEMPLATE_B64;
    if (!b64) return Promise.reject(new Error('BOM_TEMPLATE_B64 missing'));
    var sheets = ctx.workbook.worksheets;
    // insertWorksheetsFromBase64 accepts a base64 string of an .xlsx package.
    sheets.insertWorksheetsFromBase64(b64, { sheetNamesToInsert: [TEMPLATE_SHEET, MASTER_SHEET] });
    return ctx.sync().then(function () {
      // After import, fetch the two new sheets and inject formulas / recreate names.
      return injectFormulasIntoSheet(ctx, TEMPLATE_SHEET).then(function () {
        return recreateNamedRanges(ctx);
      });
    });
  }

  function importTemplate() { return ensureTemplate(); }

  // ============================================================
  // NAMED RANGES (workbook-scoped, A1 fallback)
  // ============================================================

  function recreateNamedRanges(ctx) {
    var wb = ctx.workbook;
    // Purge any sheet-scoped shadows first: get all names, delete ones whose
    // scope is a worksheet and whose name is one of ours.
    var names = wb.names.load('name,scope,worksheet');
    return ctx.sync().then(function () {
      var ours = {};
      NAMED_RANGES.forEach(function (n) { ours[n.name.toLowerCase()] = true; });
      names.items.forEach(function (nm) {
        if (ours[nm.name.toLowerCase()] && nm.scope === 'Worksheet') {
          nm.delete();
        }
      });
      // Add workbook-scoped names. Use try/catch per name with A1 fallback.
      NAMED_RANGES.forEach(function (n) {
        try {
          wb.names.add(n.name, n.refersTo);
        } catch (e) {
          // Fall back to an absolute A1 range on the Master sheet.
          var fb = a1FallbackForNamedRange(n.name);
          if (fb) {
            try { wb.names.add(n.name, "'" + MASTER_SHEET + "'!" + fb); } catch (e2) {}
          }
        }
      });
      return ctx.sync();
    });
  }

  // A1 fallback if structured references aren't accepted for some reason.
  // tblMat lives at A4:M4 header + data rows 5..N. We use the column letter.
  function a1FallbackForNamedRange(name) {
    // Column letters in tblMat: A=Material, C=Primary Unit, D=Lin, E=Area, F=Piece,
    // G=SheetL, H=SheetW. Use generous end row 999.
    var map = {
      materials: 'A5:A999', linwt: 'D5:D999', areawt: 'E5:E999',
      piecewt: 'F5:F999', wtunit: 'C5:C999', sheetl: 'G5:G999', sheetw: 'H5:H999'
    };
    return map[name.toLowerCase()] || null;
  }

  function purgeVolatileNames(ctx) {
    // Delete workbook- or sheet-scoped Print_Area / Print_Titles names.
    var names = ctx.workbook.names.load('name,scope');
    return ctx.sync().then(function () {
      names.items.forEach(function (nm) {
        var ln = nm.name.toLowerCase();
        if (ln === 'print_area' || ln === 'print_titles' ||
            ln.indexOf('!print_area') >= 0 || ln.indexOf('!print_titles') >= 0) {
          try { nm.delete(); } catch (e) {}
        }
      });
      return ctx.sync();
    });
  }

  // ============================================================
  // FORMULA ENGINE (canonical Excel 365 formulas)
  // ============================================================

  // Build the 2D formula array for the BOM table data rows (A6:M65).
  function bomFormulas() {
    var out = [];
    for (var r = 6; r <= 65; r++) {
      var prev = r - 1;
      // A: S.N.
      var A = '=IF(' + BOM_TABLE + '[@Materials]="","",MAX(A$5:A' + prev + ')+1)';
      // B-H: user inputs (Description, Materials, Length, Width, Thickness, Qty, Set)
      var B = '', C = '', D = '', E = '', F = '', G = '', H = '';
      // I: Unit weight kg/m  (single-LET; only kg/m materials show weight here)
      var I = '=IF(ROW()=$M$4+1,"Total:",IF(' + BOM_TABLE + '[@Mat Row]="","",LET(u,INDEX(WtUnit,' + BOM_TABLE + '[@Mat Row]),IF(u="kg/m",INDEX(LinWt,' + BOM_TABLE + '[@Mat Row]),""))))';
      // J: Total Weight (kg) (single-LET handling all three units + Total row sum)
      var J = '=IF(ROW()=$M$4+1,SUM(J$5:J' + prev + '),IF(' + BOM_TABLE + '[@Mat Row]="","",LET(u,INDEX(WtUnit,' + BOM_TABLE + '[@Mat Row]),q,' + BOM_TABLE + '[@Qty],s,IF(' + BOM_TABLE + '[@Set]="",1,' + BOM_TABLE + '[@Set]),IF(u="kg/m",q*s*' + BOM_TABLE + '[@Length]/1000*INDEX(LinWt,' + BOM_TABLE + '[@Mat Row]),IF(u="kg/m²",q*s*' + BOM_TABLE + '[@Length]*' + BOM_TABLE + '[@Width]/10^6*INDEX(AreaWt,' + BOM_TABLE + '[@Mat Row]),q*s*INDEX(PieceWt,' + BOM_TABLE + '[@Mat Row]))))))';
      // K: Remarks (user input)
      var K = '';
      // L: Total Length mm
      var L = '=IF(OR(' + BOM_TABLE + '[@Mat Row]="",' + BOM_TABLE + '[@[Unit weight kg/m]]="",NOT(ISNUMBER(' + BOM_TABLE + '[@Length]))),"",' + BOM_TABLE + '[@Length]*' + BOM_TABLE + '[@Qty]*IF(' + BOM_TABLE + '[@Set]="",1,' + BOM_TABLE + '[@Set]))';
      // M: Mat Row (XMATCH against Materials)
      var M = '=IFNA(XMATCH(' + BOM_TABLE + '[@Materials],Materials,0),"")';
      out.push([A, B, C, D, E, F, G, H, I, J, K, L, M]);
    }
    return out;
  }

  // Build the 2D formula array for the Purchase List table data rows (O6:AB48).
  // Columns O..AB map to S(15) N(14) -> O=15,P=16,Q=17,R=18,S=19,T=20,U=21,V=22,W=23,X=24,Y=25,Z=26,AA=27,AB=28
  // table column order in pl_Take_off:
  //   SN(O), Materials(P), Length(mm)(Q), Qty(nos)(R), Consumable Wt kg(S),
  //   Weight Kg/m(T), Total Weight(kg)(U), Stock(V), PO Weight(W), PO Qty(X),
  //   PO Rate(Y), Amount(Z), Remarks(AA), Mat Row(AB)
  function plFormulas() {
    var out = [];
    for (var r = 6; r <= 48; r++) {
      var prev = r - 1;
      // O: SN
      var O = '=IF(' + PL_TABLE + '[@Mat Row]="","",N(O' + prev + ')+1)';
      // P: Materials (Grand Total + Requested by + INDEX)
      var P = '=IF(ROW()=$AB$4+6,"Grand Total Weight in KG :",IF(ROW()=$AB$4+8,"Requested by:",IF(' + PL_TABLE + '[@Mat Row]="","",INDEX(Materials,' + PL_TABLE + '[@Mat Row]))))';
      // Q: Length (mm) — LET with unit dispatch + sheet-size lookup for kg/m²
      var Q = '=IF(ROW()=$AB$4+8,"Procurement",IF(' + PL_TABLE + '[@Mat Row]="","",LET(u,INDEX(WtUnit,' + PL_TABLE + '[@Mat Row]),IF(u="kg/m",$AE$1,IF(u="kg/m²",IFERROR(IF(INDEX(SheetL,' + PL_TABLE + '[@Mat Row])*INDEX(SheetW,' + PL_TABLE + '[@Mat Row])>0,INDEX(SheetL,' + PL_TABLE + '[@Mat Row])&" x "&INDEX(SheetW,' + PL_TABLE + '[@Mat Row]),"Add size in Master"),"Add size in Master"),"")))))';
      // R: Qty (nos) — ROUNDUP per unit
      var R = '=IF(' + PL_TABLE + '[@Mat Row]="","",LET(u,INDEX(WtUnit,' + PL_TABLE + '[@Mat Row]),IF(u="kg/m",IFERROR(ROUNDUP(ROUND(SUMIF(' + BOM_TABLE + '[Materials],' + PL_TABLE + '[@Materials],' + BOM_TABLE + '[Total Length mm])/' + PL_TABLE + '[@[Length (mm)]],6),0),""),IF(u="kg/m²",IFERROR(ROUNDUP(ROUND(' + PL_TABLE + '[@[Consumable Wt. kg]]/(INDEX(AreaWt,' + PL_TABLE + '[@Mat Row])*INDEX(SheetL,' + PL_TABLE + '[@Mat Row])*INDEX(SheetW,' + PL_TABLE + '[@Mat Row])/10^6),6),0),""),IFERROR(ROUNDUP(ROUND(' + PL_TABLE + '[@[Consumable Wt. kg]]/INDEX(PieceWt,' + PL_TABLE + '[@Mat Row]),6),0),""))))';
      // S: Consumable Wt. kg (Total row + SUMIF)
      var S = '=IF(ROW()=$AB$4+6,SUM(S$5:S' + prev + '),IF(' + PL_TABLE + '[@Mat Row]="","",SUMIF(' + BOM_TABLE + '[Materials],' + PL_TABLE + '[@Materials],' + BOM_TABLE + '[Total Weight (kg)])))';
      // T: Weight Kg/m (signature + kg/m only)
      var T = '=IF(ROW()=$AB$4+8,"Store Section:",IF(' + PL_TABLE + '[@Mat Row]="","",IF(INDEX(WtUnit,' + PL_TABLE + '[@Mat Row])="kg/m",INDEX(LinWt,' + PL_TABLE + '[@Mat Row]),"")))';
      // U: Total Weight (kg) (Total row + LET per unit)
      var U = '=IF(ROW()=$AB$4+6,SUM(U$5:U' + prev + '),IF(' + PL_TABLE + '[@Mat Row]="","",LET(u,INDEX(WtUnit,' + PL_TABLE + '[@Mat Row]),IFERROR(IF(u="kg/m",' + PL_TABLE + '[@[Length (mm)]]/1000*' + PL_TABLE + '[@[Qty (nos)]]*' + PL_TABLE + '[@[Weight Kg/m]],IF(u="kg/m²",' + PL_TABLE + '[@[Qty (nos)]]*INDEX(AreaWt,' + PL_TABLE + '[@Mat Row])*INDEX(SheetL,' + PL_TABLE + '[@Mat Row])*INDEX(SheetW,' + PL_TABLE + '[@Mat Row])/10^6,' + PL_TABLE + '[@[Qty (nos)]]*INDEX(PieceWt,' + PL_TABLE + '[@Mat Row]))),' + PL_TABLE + '[@[Consumable Wt. kg]]))))';
      // V: Stock (user input)
      var V = '';
      // W: PO Weight footer (Total row + Production signature)
      var W = '=IF(ROW()=$AB$4+6,IF(COUNT(W$5:W' + prev + ')=0,"",SUM(W$5:W' + prev + ')),IF(ROW()=$AB$4+8,"Production",""))';
      // X: PO Qty (user input)
      var X = '';
      // Y: PO Rate (user input)
      var Y = '';
      // Z: Amount (Total row + Engineering signature + per-row amount)
      var Z = '=IF(ROW()=$AB$4+6,IF(COUNT(Z$5:Z' + prev + ')=0,"",SUM(Z$5:Z' + prev + ')),IF(ROW()=$AB$4+8,"Engineering",IF(' + PL_TABLE + '[@[PO Rate]]="","",' + PL_TABLE + '[@[PO Weight]]*' + PL_TABLE + '[@[PO Rate]])))';
      // AA: Remarks (user input)
      var AA = '';
      // AB: Mat Row — O(1) INDEX into the AL6# spill
      var AB = '=IFERROR(INDEX($AL$6#,ROW()-5),"")';
      out.push([O, P, Q, R, S, T, U, V, W, X, Y, Z, AA, AB]);
    }
    return out;
  }

  // Standalone cells on every BOM sheet
  var STANDALONE_FORMULAS = {
    'O2':  '=SUBSTITUTE($A$1,"Bill of Materials","Remaining Purchase List")&IF($A$2<>""," – "&$A$2,"")',
    'M4':  '=IFNA(XMATCH(TRUE,(' + BOM_TABLE + '[Description]<>"")+(' + BOM_TABLE + '[Materials]<>"")+(' + BOM_TABLE + '[Length]<>"")+(' + BOM_TABLE + '[Qty]<>"")>0,0,-1)+5,5)',
    'AL6': '=LET(m,FILTER(' + BOM_TABLE + '[Mat Row],ISNUMBER(' + BOM_TABLE + '[Mat Row]),""),IF(INDEX(m,1)="","",SORT(UNIQUE(m))))',
    'AB4': '=IF(AL6="",0,ROWS($AL$6#))',
    'AD2': '=IF($M$4=5,"ℹ BOM is empty",IF(SUMPRODUCT((' + BOM_TABLE + '[Materials]<>"")*(' + BOM_TABLE + '[Mat Row]=""))>0,"⚠ "&SUMPRODUCT((' + BOM_TABLE + '[Materials]<>"")*(' + BOM_TABLE + '[Mat Row]=""))&" BOM row(s) use a material missing from Master",IF(COUNT(' + BOM_TABLE + '[Mat Row])>$AB$4,"⚠ "&(COUNT(' + BOM_TABLE + '[Mat Row])-$AB$4)&" material(s) missing from Purchase List","✔ Purchase list covers every BOM material")))'
  };

  // Inject every formula into a sheet (used after import / on add-sheet / on repair).
  function injectFormulasIntoSheet(ctx, sheetName) {
    var sh = ctx.workbook.worksheets.getItem(sheetName);
    var rngBom = sh.getRange('A6:M65');
    var rngPl  = sh.getRange('O6:AB48');
    rngBom.formulas = bomFormulas();
    rngPl.formulas = plFormulas();
    // standalone cells
    var keys = Object.keys(STANDALONE_FORMULAS);
    keys.forEach(function (addr) {
      sh.getRange(addr).formulas = [[STANDALONE_FORMULAS[addr]]];
    });
    // hide column AL so the spill is invisible (AB already hidden in template)
    sh.getRange('AL:AL').columnHidden = true;
    return ctx.sync();
  }

  function injectFormulas(sheetName) {
    return run({ heavy: true }, function (ctx) {
      return injectFormulasIntoSheet(ctx, sheetName);
    });
  }

  // Reapply Data Validation on the active table ranges.
  function reapplyDataValidation(ctx, sheetName) {
    var sh = ctx.workbook.worksheets.getItem(sheetName);
    // BOM: Materials dropdown on C6:C65 (list = Materials named range)
    var bomMat = sh.getRange('C6:C65').dataValidation;
    bomMat.rule = { list: { in: '=Materials' } };
    bomMat.prompt = { show: true, title: 'Material', message: 'Pick a material from the Master Unique Materials sheet.' };
    bomMat.errorAlert = { show: true, style: 'Stop', title: 'Not in the materials list',
      message: 'Pick a material from the dropdown (or add it to the Master Unique Materials sheet first).' };
    return ctx.sync();
  }
  function reapplyMasterDataValidation(ctx) {
    var sh = ctx.workbook.worksheets.getItem(MASTER_SHEET);
    var dv = sh.getRange('C5:C300').dataValidation;
    dv.rule = { list: { in: '"kg/m,kg/m²,kg/each"' } };
    dv.prompt = { show: true, title: 'Primary Unit',
      message: 'kg/m = by length, kg/m² = by area, kg/each = per piece' };
    dv.errorAlert = { show: true, style: 'Stop', title: 'Pick a unit',
      message: 'Choose kg/m, kg/m² or kg/each.' };
    return ctx.sync();
  }

  // ============================================================
  // SHEET CREATION
  // ============================================================

  function addSheet(name, opts) {
    opts = opts || {};
    return run({ heavy: true }, function (ctx) {
      var ws = ctx.workbook.worksheets;
      // Ensure template is present (import if missing)
      var bom = ws.getItemOrNullObject(TEMPLATE_SHEET);
      var mas = ws.getItemOrNullObject(MASTER_SHEET);
      return ctx.sync().then(function () {
        if (bom.isNullObject || mas.isNullObject) {
          return importTemplateIntoWorkbook(ctx);
        }
      }).then(function () {
        // Decide what to copy from
        var copyFrom = opts.copyFrom;
        var sourceSheet = null;
        if (copyFrom && copyFrom !== '__blank__') {
          sourceSheet = ws.getItem(copyFrom);
        } else if (!copyFrom) {
          // default: blank BOM from template
          sourceSheet = ws.getItem(TEMPLATE_SHEET);
        }
        var newSheet;
        if (sourceSheet) {
          newSheet = sourceSheet.copy(Excel.WorksheetPositionType.after, ws.getActiveWorksheet());
          newSheet.name = name;
        } else {
          newSheet = ws.add(name);
          // copy from BOM Blank template structure (typed headers + styles + tables)
          var tmpl = ws.getItem(TEMPLATE_SHEET);
          var rngAll = tmpl.getUsedRange();
          newSheet.getUsedRange().copyFrom(rngAll, Excel.RangeCopyType.all);
        }
        return ctx.sync().then(function () {
          // Rename tables bom_* / pl_* on the new sheet so they don't collide with template tables.
          return renameTablesOnSheet(ctx, name).then(function () {
            return injectFormulasIntoSheet(ctx, name);
          }).then(function () {
            return reapplyDataValidation(ctx, name);
          });
        }).then(function () {
          // Optional: project title in A1
          if (opts.projectTitle) {
            newSheet.getRange('A1').values = [[opts.projectTitle]];
          }
          // Optional: clear entered BOM/PO inputs while preserving formulas/PO Weight footers
          if (opts.clearInputs) {
            clearBomAndPoInputs(ctx, name);
          }
          // Optional: pre-sync Master from saved Master (Project-Only if a project is active)
          if (opts.preSync) {
            return preSyncMaster(ctx);
          }
        });
      }).then(function () {
        return { name: name, created: true };
      });
    });
  }

  function renameTablesOnSheet(ctx, sheetName) {
    var sh = ctx.workbook.worksheets.getItem(sheetName);
    var tables = sh.tables.load('name');
    return ctx.sync().then(function () {
      var renames = [];
      tables.items.forEach(function (t) {
        var n = t.name;
        if (n === BOM_TABLE || n.indexOf(BOM_TABLE + '_') === 0) {
          var suffix = '_' + sanitizeName(sheetName).slice(0, 16);
          var newName = BOM_TABLE + suffix;
          renames.push(renameTableSafely(ctx, sh, t, newName));
        } else if (n === PL_TABLE || n.indexOf(PL_TABLE + '_') === 0) {
          var suffix2 = '_' + sanitizeName(sheetName).slice(0, 16);
          var newName2 = PL_TABLE + suffix2;
          renames.push(renameTableSafely(ctx, sh, t, newName2));
        }
      });
      return Promise.all(renames).then(function () {});
    });
  }
  function renameTableSafely(ctx, sheet, table, newName) {
    // Use the canonical rewrite pattern: drop and re-add to fully purge stale refs.
    return new Promise(function (resolve) {
      try {
        table.name = newName;
        ctx.sync().then(resolve, function () {
          // If collision (rare), append a counter.
          try {
            table.name = newName + '_' + Math.floor(Math.random() * 9999);
            ctx.sync().then(resolve, resolve);
          } catch (e) { resolve(); }
        });
      } catch (e) { resolve(); }
    });
  }
  function sanitizeName(s) { return String(s || '').replace(/[^A-Za-z0-9_]/g, '_'); }

  // Clear user-typed BOM (B-H, K) and PO (V, X, Y, AA) input cells. Formulas in
  // A, I, J, L, M and O, P, Q, R, S, T, U, W, Z, AB are preserved.
  function clearBomAndPoInputs(ctx, sheetName) {
    var sh = ctx.workbook.worksheets.getItem(sheetName);
    var blanksBom = makeBlanks(60, 7); // B6:H65 (7 cols)
    var blanksK = makeBlanks(60, 1);
    var blanksV = makeBlanks(43, 1);
    var blanksXY = makeBlanks(43, 2); // X6:Y65 — but PL is rows 6-48; use 43 rows
    var blanksAA = makeBlanks(43, 1);
    sh.getRange('B6:H65').values = blanksBom;
    sh.getRange('K6:K65').values = blanksK;
    sh.getRange('V6:V48').values = blanksV;
    sh.getRange('X6:Y48').values = blanksXY;
    sh.getRange('AA6:AA48').values = blanksAA;
    // (W = PO Weight has formulas; do NOT clear)
    return ctx.sync();
  }
  function makeBlanks(rows, cols) {
    var out = [];
    for (var i = 0; i < rows; i++) {
      var row = []; for (var j = 0; j < cols; j++) row.push(null);
      out.push(row);
    }
    return out;
  }

  // Pre-sync Master Unique Materials from saved Master slot (Project-Only if active).
  function preSyncMaster(ctx) {
    if (!root.BomStore) return Promise.resolve();
    return root.BomStore.getMaster().then(function (materials) {
      if (!materials || !materials.length) return;
      return writeMaterialsToMaster(ctx, materials).then(function () {
        return reapplyMasterDataValidation(ctx);
      });
    });
  }

  // ============================================================
  // MATERIALS READ / WRITE
  // ============================================================

  // Map a tblMat row (array of 13 columns) -> material object.
  // Cols: A Material, B Category, C Primary Unit, D Lin, E Area, F Piece,
  //       G SheetL, H SheetW, I SheetArea(formula), J SheetWt(formula),
  //       K Remarks, L Last Action, M Last Modified
  function rowToMat(row) {
    if (!row) return null;
    var name = row[0];
    if (name === '' || name === null || name === undefined) return null;
    return {
      name: String(name),
      category: row[1] || '',
      unit: row[2] || '',
      linear: num(row[3]), area: num(row[4]), piece: num(row[5]),
      sheetL: num(row[6]), sheetW: num(row[7]),
      remarks: row[10] || '',
      lastAction: row[11] || '',
      lastModified: row[12] || ''
    };
  }
  function num(v) { return (v === '' || v === null || v === undefined) ? null : Number(v); }
  function matToRow(m) {
    return [
      m.name || '', m.category || '', m.unit || '',
      m.linear === null ? null : m.linear,
      m.area === null ? null : m.area,
      m.piece === null ? null : m.piece,
      m.sheetL === null ? null : m.sheetL,
      m.sheetW === null ? null : m.sheetW,
      null, null, // I, J are formulas (let Excel recompute)
      m.remarks || '', m.lastAction || '', m.lastModified || ''
    ];
  }

  // Read Master Unique Materials A4:M{last} into material objects.
  // On-demand snapshot — no onChanged listener.
  function readMaterials() {
    return run(function (ctx) {
      var sh = ctx.workbook.worksheets.getItem(MASTER_SHEET);
      // tblMat starts at A4 with header row 4 + data 5+. We read A5:M304 (generous).
      var rng = sh.getRange('A5:M304');
      rng.load('values');
      return ctx.sync().then(function () {
        var vals = rng.values;
        var out = [];
        for (var i = 0; i < vals.length; i++) {
          var m = rowToMat(vals[i]);
          if (m) out.push(m);
        }
        // Update Row 3 summary banner with material count
        sh.getRange('A3').values = [['Last snapshot: ' + out.length + ' materials · ' + new Date().toLocaleString()]];
        return ctx.sync().then(function () { return out; });
      });
    });
  }

  // Snapshot diff vs last-known. Returns { added, modified, removed, all }.
  function snapshotDiff() {
    return Promise.all([
      readMaterials(),
      root.BomStore ? root.BomStore.getMaster() : Promise.resolve([])
    ]).then(function (r) {
      var current = r[0] || [];
      var known = r[1] || [];
      var byNameCur = {}; current.forEach(function (m) { byNameCur[m.name] = m; });
      var byNameKn = {}; known.forEach(function (m) { byNameKn[m.name] = m; });
      var added = [], modified = [], removed = [];
      current.forEach(function (m) {
        if (!byNameKn[m.name]) added.push(m);
        else if (JSON.stringify(m) !== JSON.stringify(byNameKn[m.name])) modified.push(m);
      });
      known.forEach(function (m) {
        if (!byNameCur[m.name]) removed.push(m);
      });
      return { added: added, modified: modified, removed: removed, all: current };
    });
  }

  // Rewrite Master Unique Materials tblMat from a materials array.
  function writeMaterialsToMaster(ctx, materials) {
    var sh = ctx.workbook.worksheets.getItem(MASTER_SHEET);
    var tbl = sh.tables.getItem(MAT_TABLE);
    // Resize tblMat data range to fit materials (header at row 4, data from row 5).
    var rowCount = Math.max(materials.length, 1);
    var newRef = 'A4:M' + (4 + rowCount);
    tbl.resize(newRef);
    // Build values 2D array (rowCount rows × 13 cols).
    var rows = materials.map(matToRow);
    while (rows.length < rowCount) {
      rows.push(['', '', '', null, null, null, null, null, null, null, '', '', '']);
    }
    var dataRange = sh.getRange('A5:M' + (4 + rowCount));
    dataRange.values = rows;
    // Re-inject the I/J formulas for the new range (Sheet Area, Sheet Weight)
    var formulas = [];
    for (var r = 0; r < rowCount; r++) {
      var rowNum = 5 + r;
      formulas.push([
        '=IF(AND(ISNUMBER(G' + rowNum + '),ISNUMBER(H' + rowNum + ')),G' + rowNum + '*H' + rowNum + '/10^6,"")',
        '=IF(AND(ISNUMBER(I' + rowNum + '),ISNUMBER(E' + rowNum + ')),I' + rowNum + '*E' + rowNum + ',"")'
      ]);
    }
    sh.getRange('I5:J' + (4 + rowCount)).formulas = formulas;
    // Update Row 3 summary banner
    sh.getRange('A3').values = [['Last saved: ' + materials.length + ' materials · ' + new Date().toLocaleString()]];
    return ctx.sync();
  }

  function writeMaterials(materials) {
    return run({ heavy: true }, function (ctx) {
      return writeMaterialsToMaster(ctx, materials).then(function () {
        return reapplyMasterDataValidation(ctx);
      });
    });
  }

  // ============================================================
  // MATERIALS CRUD + AUDIT STAMPING
  // ============================================================

  function addMaterial(mat) {
    return run({ heavy: true }, function (ctx) {
      return readMaterialsInto(ctx).then(function (cur) {
        // Reject duplicate name (case-insensitive)
        var exists = cur.some(function (m) { return m.name.toLowerCase() === mat.name.toLowerCase(); });
        if (exists) throw new Error('A material named "' + mat.name + '" already exists');
        mat.lastAction = 'Added';
        mat.lastModified = new Date().toISOString();
        cur.push(mat);
        return writeMaterialsToMaster(ctx, cur).then(function () {
          stampAuditRow(ctx, cur.length, mat); // L/M on the new row
          return ctx.sync().then(function () {
            if (root.BomStore) root.BomStore.recordAudit({ action: 'Added', name: mat.name, at: mat.lastModified });
            if (root.BomStore) root.BomStore.saveMaster(cur);
            return mat;
          });
        });
      });
    });
  }

  function updateMaterial(name, mat) {
    return run({ heavy: true }, function (ctx) {
      return readMaterialsInto(ctx).then(function (cur) {
        var idx = cur.findIndex(function (m) { return m.name === name; });
        if (idx < 0) throw new Error('Material not found: ' + name);
        // If renaming, ensure new name is unique
        if (mat.name && mat.name.toLowerCase() !== name.toLowerCase()) {
          var clash = cur.some(function (m, i) { return i !== idx && m.name.toLowerCase() === mat.name.toLowerCase(); });
          if (clash) throw new Error('A material named "' + mat.name + '" already exists');
        }
        mat.lastAction = 'Modified';
        mat.lastModified = new Date().toISOString();
        cur[idx] = Object.assign({}, cur[idx], mat);
        return writeMaterialsToMaster(ctx, cur).then(function () {
          stampAuditRow(ctx, idx + 1, cur[idx]);
          return ctx.sync().then(function () {
            if (root.BomStore) root.BomStore.recordAudit({ action: 'Modified', name: cur[idx].name, at: cur[idx].lastModified });
            if (root.BomStore) root.BomStore.saveMaster(cur);
            return cur[idx];
          });
        });
      });
    });
  }

  function removeMaterial(name) {
    return run({ heavy: true }, function (ctx) {
      return readMaterialsInto(ctx).then(function (cur) {
        var idx = cur.findIndex(function (m) { return m.name === name; });
        if (idx < 0) throw new Error('Material not found: ' + name);
        var removed = cur.splice(idx, 1)[0];
        return writeMaterialsToMaster(ctx, cur).then(function () {
          return ctx.sync().then(function () {
            if (root.BomStore) root.BomStore.recordAudit({ action: 'Removed', name: removed.name, at: new Date().toISOString() });
            if (root.BomStore) root.BomStore.saveMaster(cur);
            return removed;
          });
        });
      });
    });
  }

  // Stamp L/M columns on a given data row (1-based index among data rows; row in sheet = idx + 4)
  function stampAuditRow(ctx, dataIdx, mat) {
    var sh = ctx.workbook.worksheets.getItem(MASTER_SHEET);
    var rowNum = 4 + dataIdx;
    sh.getRange('L' + rowNum + ':M' + rowNum).values = [[mat.lastAction || '', mat.lastModified || '']];
    // Row 3 summary banner
    sh.getRange('A3').values = [['Last action: ' + mat.lastAction + ' "' + mat.name + '" · ' + new Date().toLocaleString()]];
  }

  function readMaterialsInto(ctx) {
    var sh = ctx.workbook.worksheets.getItem(MASTER_SHEET);
    var rng = sh.getRange('A5:M304');
    rng.load('values');
    return ctx.sync().then(function () {
      var out = [];
      rng.values.forEach(function (row) {
        var m = rowToMat(row);
        if (m) out.push(m);
      });
      return out;
    });
  }

  // ============================================================
  // WORKBOOK STATUS
  // ============================================================

  function getWorkbookStatus() {
    return run(function (ctx) {
      var ws = ctx.workbook.worksheets.load('name');
      var wbProps = ctx.workbook.properties;
      wbProps.load('items');
      return ctx.sync().then(function () {
        var names = ws.items.map(function (s) { return s.name; });
        var bomSheets = names.filter(function (n) {
          return n !== MASTER_SHEET && n !== TEMPLATE_SHEET;
        });
        var hasBomBlank = names.indexOf(TEMPLATE_SHEET) >= 0;
        var hasMaster = names.indexOf(MASTER_SHEET) >= 0;
        // Try to read BomProjectList custom property
        var bomProjectList = null;
        try {
          var prop = ctx.workbook.properties.getItemOrNullObject('BomProjectList');
          return ctx.sync().then(function () {
            if (!prop.isNullObject) bomProjectList = prop.value;
            return finishStatus(bomSheets, hasBomBlank, hasMaster, bomProjectList);
          });
        } catch (e) {
          return finishStatus(bomSheets, hasBomBlank, hasMaster, null);
        }
      });
    });
  }
  function finishStatus(bomSheets, hasBomBlank, hasMaster, bomProjectList) {
    var activeProject = root.BomStore ? root.BomStore.getActiveProject() : null;
    return {
      bomCount: bomSheets.length,
      bomSheets: bomSheets,
      hasBomBlank: hasBomBlank,
      hasMaster: hasMaster,
      needsImport: !hasBomBlank || !hasMaster,
      activeProject: activeProject,
      bomProjectList: bomProjectList,
      isProjectOnly: !!activeProject || !!bomProjectList
    };
  }

  // ============================================================
  // REPAIR (per-sheet)
  // ============================================================

  function repairSheet(sheetName) {
    return run({ heavy: true }, function (ctx) {
      // 1) recreate workbook named ranges (purges shadowing sheet-scoped names)
      return recreateNamedRanges(ctx).then(function () {
        // 2) reapply active-table Data Validation
        if (sheetName === MASTER_SHEET) {
          return reapplyMasterDataValidation(ctx);
        }
        return reapplyDataValidation(ctx, sheetName).then(function () {
          // 3) upgrade to Excel 365 formulas
          return injectFormulasIntoSheet(ctx, sheetName);
        });
      }).then(function () {
        // 4) purge volatile Print_Area / Print_Titles names
        return purgeVolatileNames(ctx);
      }).then(function () {
        return { repaired: true, sheet: sheetName };
      });
    });
  }

  function repairAll(sheetNames) {
    var arr = sheetNames.slice();
    var chain = Promise.resolve();
    arr.forEach(function (n) {
      chain = chain.then(function () { return repairSheet(n); });
    });
    return chain;
  }

  // ============================================================
  // PUBLIC API
  // ============================================================
  root.BomSheets = {
    // constants
    BOM_TABLE: BOM_TABLE, PL_TABLE: PL_TABLE, MAT_TABLE: MAT_TABLE,
    MASTER_SHEET: MASTER_SHEET, TEMPLATE_SHEET: TEMPLATE_SHEET,
    NAMED_RANGES: NAMED_RANGES, STANDALONE_FORMULAS: STANDALONE_FORMULAS,
    // template + sheets
    ensureTemplate: ensureTemplate,
    importTemplate: importTemplate,
    importTemplateIntoWorkbook: importTemplateIntoWorkbook,
    addSheet: addSheet,
    listSheetNames: listSheetNames,
    validateSheetName: validateSheetName,
    isBomSheetName: isBomSheetName,
    // formulas + named ranges
    injectFormulas: injectFormulas,
    injectFormulasIntoSheet: injectFormulasIntoSheet,
    bomFormulas: bomFormulas,
    plFormulas: plFormulas,
    recreateNamedRanges: recreateNamedRanges,
    purgeVolatileNames: purgeVolatileNames,
    reapplyDataValidation: reapplyDataValidation,
    reapplyMasterDataValidation: reapplyMasterDataValidation,
    // materials
    readMaterials: readMaterials,
    snapshotDiff: snapshotDiff,
    writeMaterials: writeMaterials,
    writeMaterialsToMaster: writeMaterialsToMaster,
    addMaterial: addMaterial,
    updateMaterial: updateMaterial,
    removeMaterial: removeMaterial,
    // status + repair
    getWorkbookStatus: getWorkbookStatus,
    repairSheet: repairSheet,
    repairAll: repairAll,
    // helpers exposed for convert.js
    _run: run,
    _clearBomAndPoInputs: clearBomAndPoInputs,
    _renameTablesOnSheet: renameTablesOnSheet
  };
})(typeof window !== 'undefined' ? window : this);
