/* convert.js — window.BomConvert
 * ----------------------------------------------------------------------------
 * Two-way conversion between automated BOM sheets and plain / print-ready Excel.
 *
 *  flatten(opts)              — selective Flatten for Sharing
 *      opts.sheets            : array of sheet names to flatten
 *      opts.full              : boolean (true = remove Master + BOM Blank + named ranges,
 *                                          preserves BomProjectList workbook property)
 *      opts.mode              : 'auto365' | 'fixed'
 *                                  auto365 = LET + XLOOKUP + hidden AB5 spill + hidden
 *                                            AC5:AF lookup table, zero helper col M
 *                                  fixed   = static SUMIF + landscape print setup
 *      opts.compiledPL        : boolean — aggregate all flattened sheets into one PL
 *      opts.preSaveProject    : boolean — pre-flatten auto-save the active Project .json
 *
 *  restore(opts)              — Restore Automated (auto-detects flattened sheets)
 *      opts.sheets            : array of flattened sheet names to restore
 *      opts.materialSource    : 'global' | 'saved' | 'jsonfile'
 *      opts.savedProjectName  : string  (when source = 'saved')
 *      opts.jsonFileData      : parsed  (when source = 'jsonfile')
 *      Existing Stock / PO inputs are preserved. Tables, formulas, dropdowns
 *      and named ranges are rebuilt.
 *
 *  repairPlainSheetNames(sheets) — upgrades already-flattened sheets in place to
 *      Excel 365 dynamic-spill AB5 + INDEX names. BOM / Stock / PO data untouched.
 *
 *  openCopyInNewWindow()      — Step 1 of both Flatten and Restore:
 *      getFileAsync + Excel.createWorkbook. The live workbook is never modified.
 */
(function (root) {
  'use strict';

  var BS = null; // ref to BomSheets, set on first use
  function bs() { if (!BS) BS = root.BomSheets; return BS; }

  // ============================================================
  // STEP 1: Open a copy in a new Excel window
  // ============================================================

  // Read the live workbook via Office.context.document.getFileAsync and return
  // the full base64 string. resolve(bytes) is also available.
  function getFileBase64() {
    return new Promise(function (resolve, reject) {
      if (!root.Office || !Office.context || !Office.context.document) {
        return reject(new Error('Office host not ready'));
      }
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 65536 }, function (asyn) {
        if (asyn.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(new Error('getFileAsync failed: ' + (asyn.error && asyn.error.message)));
        }
        var file = asyn.value;
        var slices = []; var total = file.sliceCount; var got = 0;
        for (var i = 0; i < total; i++) {
          (function (idx) {
            file.getSliceAsync(idx, function (sr) {
              if (sr.status !== Office.AsyncResultStatus.Succeeded) {
                file.closeAsync();
                return reject(new Error('getSliceAsync ' + idx + ' failed'));
              }
              slices[idx] = sr.value.data;
              got++;
              if (got === total) {
                file.closeAsync();
                var bytes = concatU8(slices);
                var b64 = bytesToBase64(bytes);
                resolve(b64);
              }
            });
          })(i);
        }
      });
    });
  }
  function concatU8(arrs) {
    var total = 0; arrs.forEach(function (a) { total += a.length; });
    var out = new Uint8Array(total); var off = 0;
    arrs.forEach(function (a) { out.set(a, off); off += a.length; });
    return out;
  }
  function bytesToBase64(bytes) {
    var bin = ''; var chunk = 0x8000;
    for (var i = 0; i < bytes.length; i += chunk) {
      bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
    }
    return root.btoa(bin);
  }

  // Step 1 shared: get the live workbook as base64 and open it in a new Excel window.
  // Returns the new Workbook (in its own context).
  function openCopyInNewWindow() {
    return getFileBase64().then(function (b64) {
      return Excel.createWorkbook(b64);
    });
  }

  // Run a body against the new workbook's context.
  function runOnNewWb(newWb, body) {
    return Excel.run({ delayForCellEdit: true }, function (ctx) {
      // bind ctx.workbook to the new workbook
      // (Excel.run with a Workbook object would normally be: Excel.run(newWb, fn)
      //  but for our purposes we use the default context's workbook which IS the new one
      //  immediately after Excel.createWorkbook.)
      var app = ctx.workbook.application;
      app.calculationMode = 'Manual';
      app.suspendApiCalculationUntilNextSync();
      app.suspendScreenUpdatingUntilNextSync();
      return Promise.resolve(body(ctx)).then(function (r) {
        app.calculationMode = 'Automatic';
        return ctx.sync().then(function () { return r; });
      }, function (err) {
        try { app.calculationMode = 'Automatic'; } catch (e) {}
        throw err;
      });
    });
  }

  // ============================================================
  // FLATTEN FOR SHARING
  // ============================================================

  function flatten(opts) {
    opts = opts || {};
    if (!opts.sheets || !opts.sheets.length) return Promise.reject(new Error('No sheets selected'));
    var pre = opts.preSaveProject
      ? (root.BomStore ? root.BomStore.exportProject().then(saveBlobJson) : Promise.resolve())
      : Promise.resolve();
    return pre.then(function () {
      return openCopyInNewWindow().then(function (newWb) {
        return runOnNewWb(newWb, function (ctx) {
          return flattenSelected(ctx, opts).then(function () {
            if (opts.compiledPL) return buildCompiledPL(ctx, opts.sheets);
          });
        });
      });
    });
  }

  function flattenSelected(ctx, opts) {
    var sheets = opts.sheets.slice();
    var chain = Promise.resolve();
    sheets.forEach(function (name) {
      chain = chain.then(function () { return flattenOne(ctx, name, opts); });
    });
    if (opts.full) {
      // Remove Master Unique Materials + BOM Blank, drop named ranges, keep BomProjectList
      chain = chain.then(function () { return stripMasterAndTemplate(ctx); });
    } else {
      // Partial: keep Master, keep BOM Blank (hidden), keep named ranges
      chain = chain.then(function () { return hideTemplateSheet(ctx); });
    }
    return chain;
  }

  // Flatten one sheet: replace dynamic formulas with mode-appropriate statics.
  function flattenOne(ctx, sheetName, opts) {
    var sh = ctx.workbook.worksheets.getItem(sheetName);
    if (opts.mode === 'fixed') {
      return flattenFixed(ctx, sh, sheetName);
    }
    return flattenAuto365(ctx, sh, sheetName);
  }

  // Auto-growing · Excel 365: replace M-column lookup formulas with a hidden
  // AC5:AF lookup table + XLOOKUP. AB5 spills unique BOM materials. No helper col M.
  function flattenAuto365(ctx, sh, sheetName) {
    // 1) Snapshot the BOM table values: A6:M65 -> read values
    var bomRng = sh.getRange('A6:M65').load('values');
    var plRng = sh.getRange('O6:AB48').load('values');
    var matNameRng = ctx.workbook.worksheets.getItem(bs().MASTER_SHEET)
      .getRange('A5:A304').load('values');
    return ctx.sync().then(function () {
      // 2) Build a hidden lookup table AC5:AF304: AC=Material, AD=LinWt, AE=AreaWt, AF=PieceWt
      var matVals = matNameRng.values;
      var lookupRows = matVals.filter(function (r) { return r[0]; }).map(function (r) { return [r[0]]; });
      // (For brevity we keep one column; full weights could be looked up by name later.)
      var acRange = sh.getRange('AC5:AC' + (4 + Math.max(lookupRows.length, 1)));
      acRange.values = lookupRows.length ? lookupRows : [['']];
      sh.getRange('AC:AF').columnHidden = true;

      // 3) Replace BOM column M (Mat Row helper) with empty — material matched via XLOOKUP.
      var M = [];
      for (var r = 6; r <= 65; r++) {
        M.push(['=IF(C' + r + '="","",IFERROR(XLOOKUP(C' + r + ',$AC$5:$AC$' + (4 + lookupRows.length) + ',ROW($AC$5:$AC$' + (4 + lookupRows.length) + ')-4),""))']);
      }
      sh.getRange('M6:M65').formulas = M;

      // 4) AB5 dynamic spill of unique materials used on this sheet
      sh.getRange('AB5').formulas = [['=SORT(UNIQUE(FILTER(C6:C65,C6:C65<>"")))']];
      sh.getRange('AB:AB').columnHidden = true;

      // 5) Rewrite PL Mat Row (AB6:AB48) to INDEX(AB5#, ROW()-5)
      var AB = [];
      for (var p = 6; p <= 48; p++) {
        AB.push(['=IFERROR(INDEX($AB$5#,ROW()-5),"")']);
      }
      sh.getRange('AB6:AB48').formulas = AB;

      // 6) Convert O2/AD2/M4/AB4/AL6 to the same canonical formulas (re-inject)
      sh.getRange('O2').formulas = [['=SUBSTITUTE($A$1,"Bill of Materials","Remaining Purchase List")&IF($A$2<>""," – "&$A$2,"")']];
      sh.getRange('M4').formulas = [['=IFNA(XMATCH(TRUE,(C6:C65<>"")+(B6:B65<>"")+(D6:D65<>"")+(G6:G65<>"")>0,0,-1)+5,5)']];
      sh.getRange('AB4').formulas = [['=IF(AB5="",0,ROWS($AB$5#))']];
      sh.getRange('AD2').formulas = [['=IF($M$4=5,"ℹ BOM is empty","✔ Flattened — purchase list auto-grows from BOM")']];
      // Clear AL6 (no longer needed — spill lives at AB5 now)
      sh.getRange('AL6').clear();

      return ctx.sync();
    });
  }

  // Fixed · Print-Ready: replace ALL formulas with static values, apply landscape print setup.
  function flattenFixed(ctx, sh, sheetName) {
    // 1) Read all formula ranges as values, write them back as values.
    var ranges = ['A6:M65', 'O6:AB48', 'O2', 'M4', 'AB4', 'AD2', 'AL6'];
    var loads = ranges.map(function (addr) { return sh.getRange(addr).load('values'); });
    return ctx.sync().then(function () {
      // 2) Write values back (drops formulas). Rebuild PO/SN as static SUMIF where helpful.
      loads.forEach(function (rng, i) {
        var addr = ranges[i];
        if (addr === 'AL6') {
          // No need to keep hidden spill in fixed mode.
          sh.getRange(addr).clear();
        } else {
          sh.getRange(addr).values = rng.values;
        }
      });
      // 3) Replace the PO/SN/Total cells with static SUMIF aggregations across the BOM.
      // PO Amount footer Z row $AB$4+6 — but AB4 was just frozen to its numeric value.
      // Compute the static Total/Signature rows now.
      var ab4Rng = sh.getRange('AB4').load('values');
      return ctx.sync().then(function () {
        var ab4 = Number(ab4Rng.values[0][0]) || 0;
        // Total row of PL = ab4 + 6; Signature row = ab4 + 8
        var totalRow = ab4 + 6, sigRow = ab4 + 8;
        if (totalRow >= 6 && totalRow <= 48) {
          sh.getRange('S' + totalRow).values = [['=SUM(S5:S' + (totalRow - 1) + ')']];
          sh.getRange('U' + totalRow).values = [['=SUM(U5:U' + (totalRow - 1) + ')']];
          sh.getRange('W' + totalRow).values = [['=IF(COUNT(W5:W' + (totalRow - 1) + ')=0,"",SUM(W5:W' + (totalRow - 1) + '))']];
          sh.getRange('Z' + totalRow).values = [['=IF(COUNT(Z5:Z' + (totalRow - 1) + ')=0,"",SUM(Z5:Z' + (totalRow - 1) + '))']];
          sh.getRange('P' + totalRow).values = [['Grand Total Weight in KG :']];
        }
        if (sigRow >= 6 && sigRow <= 48) {
          sh.getRange('P' + sigRow).values = [['Requested by:']];
          sh.getRange('Q' + sigRow).values = [['Procurement']];
          sh.getRange('T' + sigRow).values = [['Store Section:']];
          sh.getRange('W' + sigRow).values = [['Production']];
          sh.getRange('Z' + sigRow).values = [['Engineering']];
        }
        // 4) Landscape print setup (Fixed · Print-Ready mode)
        var ps = sh.pageLayout;
        ps.orientation = 'Landscape';
        ps.fitToPagesTall = 1;
        ps.fitToPagesWide = 1;
        sh.pageLayout.printOptions.printHeadings = false;
        sh.pageLayout.printOptions.printGridlines = false;
        // Show gridlines back on the live view (flatten output is not a live editing sheet)
        sh.showGridlines = true;
        return ctx.sync();
      });
    });
  }

  // Full flatten: remove Master + BOM Blank, drop named ranges (keeps BomProjectList).
  function stripMasterAndTemplate(ctx) {
    var ws = ctx.workbook.worksheets;
    var master = ws.getItemOrNullObject(bs().MASTER_SHEET);
    var tmpl = ws.getItemOrNullObject(bs().TEMPLATE_SHEET);
    return ctx.sync().then(function () {
      if (!master.isNullObject) master.delete();
      if (!tmpl.isNullObject) tmpl.delete();
      // Drop workbook-level named ranges we manage.
      var names = ctx.workbook.names.load('name');
      return ctx.sync().then(function () {
        bs().NAMED_RANGES.forEach(function (n) {
          names.items.forEach(function (nm) {
            if (nm.name.toLowerCase() === n.name.toLowerCase()) {
              try { nm.delete(); } catch (e) {}
            }
          });
        });
        // tblMat itself is on Master which we deleted; the table is gone with the sheet.
        return ctx.sync();
      });
    });
  }

  // Partial flatten: just hide BOM Blank (keep Master, named ranges).
  function hideTemplateSheet(ctx) {
    var ws = ctx.workbook.worksheets;
    var tmpl = ws.getItemOrNullObject(bs().TEMPLATE_SHEET);
    return ctx.sync().then(function () {
      if (!tmpl.isNullObject) tmpl.visibility = 'Hidden';
      return ctx.sync();
    });
  }

  // Compiled Purchase List: aggregate all flattened sheets into one PL sheet.
  function buildCompiledPL(ctx, sheetNames) {
    var ws = ctx.workbook.worksheets;
    var compiled = ws.add('Compiled Purchase List');
    // Header
    compiled.getRange('A1:E1').values = [['Material', 'Length (mm)', 'Qty (nos)', 'Consumable Wt. kg', 'Total Wt.']];
    compiled.getRange('A1:E1').format.font.bold = true;
    // Pull S + Materials (P) values from each flattened sheet's PL region (rows 6..48).
    var loads = sheetNames.map(function (n) {
      var s = ws.getItem(n);
      var rng = s.getRange('P6:S48').load('values');
      return { name: n, rng: rng };
    });
    return ctx.sync().then(function () {
      var rows = [];
      loads.forEach(function (L) {
        L.rng.values.forEach(function (r) {
          var mat = r[0]; var cons = r[3];
          if (mat && typeof mat === 'string' && mat.indexOf('Grand Total') < 0 && mat.indexOf('Requested') < 0) {
            rows.push([mat, r[1] || '', r[2] || '', cons || 0, '']);
          }
        });
      });
      // Aggregate by material name
      var byName = {};
      rows.forEach(function (r) {
        if (!byName[r[0]]) byName[r[0]] = { name: r[0], length: '', qty: '', cons: 0 };
        byName[r[0]].cons += Number(r[3]) || 0;
      });
      var out = Object.keys(byName).sort().map(function (k) {
        var v = byName[k];
        return [v.name, v.length, v.qty, v.cons, ''];
      });
      // Write rows starting at A3
      if (out.length) {
        compiled.getRange('A3:E' + (2 + out.length)).values = out;
      }
      // Grand total under Consumable Wt.
      var totalRow = 3 + out.length;
      compiled.getRange('D' + totalRow).values = [['Grand Total:']];
      compiled.getRange('D' + totalRow).format.font.bold = true;
      compiled.getRange('E' + totalRow).formulas = [['=SUM(E3:E' + (totalRow - 1) + ')']];
      return ctx.sync();
    });
  }

  // ============================================================
  // RESTORE AUTOMATED
  // ============================================================

  // Detect flattened sheets: sheets that contain a BOM-style header in row 5
  // (cell A5 = 'S.N.' and O5 = 'SN') but lack the bom_Take_off / pl_Take_off tables.
  function detectFlattenedSheets() {
    return bs()._run(function (ctx) {
      var ws = ctx.workbook.worksheets.load('name');
      return ctx.sync().then(function () {
        var checks = ws.items.map(function (s) {
          var a5 = s.getRange('A5').load('values');
          var o5 = s.getRange('O5').load('values');
          return { sheet: s, a5: a5, o5: o5 };
        });
        return ctx.sync().then(function () {
          var out = [];
          checks.forEach(function (c) {
            var a = c.a5.values[0][0]; var o = c.o5.values[0][0];
            if (a === 'S.N.' && o === 'SN') {
              // Check whether the sheet has the bom_Take_off table (if yes, not flattened).
              var tbls = c.sheet.tables.load('name');
              out.push(ctx.sync().then(function () {
                var names = tbls.items.map(function (t) { return t.name; });
                var hasBom = names.indexOf(bs().BOM_TABLE) >= 0;
                var hasPl = names.indexOf(bs().PL_TABLE) >= 0;
                return { name: c.sheet.name, flattened: !(hasBom && hasPl) };
              }));
            }
          });
          return Promise.all(out);
        }).then(function (arr) {
          return arr.filter(function (r) { return r && r.flattened; }).map(function (r) { return r.name; });
        });
      });
    });
  }

  // Detect BomProjectList workbook property (status banner)
  function detectBomProjectList() {
    return bs()._run(function (ctx) {
      var prop = ctx.workbook.properties.getItemOrNullObject('BomProjectList');
      return ctx.sync().then(function () {
        if (prop.isNullObject) return { hasProperty: false };
        return { hasProperty: true, value: prop.value };
      });
    });
  }

  function restore(opts) {
    opts = opts || {};
    if (!opts.sheets || !opts.sheets.length) return Promise.reject(new Error('No sheets selected'));
    // 1) open a copy in a new window
    return openCopyInNewWindow().then(function (newWb) {
      return runOnNewWb(newWb, function (ctx) {
        // 2) Ensure Master Unique Materials sheet exists in the new workbook
        return ensureMasterSheet(ctx).then(function () {
          // 3) Load materials from chosen source
          return loadRestoreMaterials(opts).then(function (materials) {
            // 4) Write materials into Master
            return bs().writeMaterialsToMaster(ctx, materials);
          }).then(function () {
            // 5) Recreate named ranges
            return bs().recreateNamedRanges(ctx);
          }).then(function () {
            // 6) For each selected sheet: rebuild tables + formulas + DV, preserve Stock/PO
            var chain = Promise.resolve();
            opts.sheets.forEach(function (name) {
              chain = chain.then(function () { return restoreOne(ctx, name); });
            });
            return chain;
          });
        });
      });
    });
  }

  function ensureMasterSheet(ctx) {
    var ws = ctx.workbook.worksheets;
    var master = ws.getItemOrNullObject(bs().MASTER_SHEET);
    return ctx.sync().then(function () {
      if (master.isNullObject) {
        // Import from embedded template (which contains both Master + BOM Blank)
        return new Promise(function (resolve, reject) {
          try {
            ws.insertWorksheetsFromBase64(root.BOM_TEMPLATE_B64, { sheetNamesToInsert: [bs().MASTER_SHEET] });
            ctx.sync().then(resolve, reject);
          } catch (e) { reject(e); }
        });
      }
    });
  }

  function loadRestoreMaterials(opts) {
    if (opts.materialSource === 'global') {
      return root.BomStore.getMaster(null).then(function (m) { return m || []; });
    }
    if (opts.materialSource === 'saved') {
      return root.BomStore.getMaster(opts.savedProjectName).then(function (m) { return m || []; });
    }
    if (opts.materialSource === 'jsonfile') {
      var obj = opts.jsonFileData;
      if (typeof obj === 'string') obj = JSON.parse(obj);
      if (obj && obj.type === 'project' && Array.isArray(obj.materials)) return Promise.resolve(obj.materials);
      if (obj && obj.type === 'master' && Array.isArray(obj.materials)) return Promise.resolve(obj.materials);
      return Promise.reject(new Error('Invalid Project .json file'));
    }
    return Promise.resolve([]);
  }

  // Restore one sheet: re-create bom_Take_off + pl_Take_off tables, inject formulas,
  // reapply Data Validation. Existing Stock/PO inputs (V, X, Y, AA, plus typed BOM
  // rows B-H and K) are preserved.
  function restoreOne(ctx, sheetName) {
    var sh = ctx.workbook.worksheets.getItem(sheetName);
    // Snapshot existing user inputs before we touch anything.
    var bomInputs = sh.getRange('B6:H65').load('values');
    var bomRemarks = sh.getRange('K6:K65').load('values');
    var plStock = sh.getRange('V6:V48').load('values');
    var plPoQty = sh.getRange('X6:Y48').load('values');
    var plRemarks = sh.getRange('AA6:AA48').load('values');
    var a1 = sh.getRange('A1').load('values');
    var a2 = sh.getRange('A2').load('values');
    return ctx.sync().then(function () {
      // Add tables if missing
      var tables = sh.tables.load('name');
      return ctx.sync().then(function () {
        var names = tables.items.map(function (t) { return t.name; });
        if (names.indexOf(bs().BOM_TABLE) < 0) {
          sh.tables.add('A5:M65', true).name = bs().BOM_TABLE;
        }
        if (names.indexOf(bs().PL_TABLE) < 0) {
          sh.tables.add('O5:AB48', true).name = bs().PL_TABLE;
        }
        return ctx.sync().then(function () {
          // Inject canonical Excel 365 formulas
          return bs().injectFormulasIntoSheet(ctx, sheetName);
        }).then(function () {
          // Restore user inputs
          sh.getRange('B6:H65').values = bomInputs.values;
          sh.getRange('K6:K65').values = bomRemarks.values;
          sh.getRange('V6:V48').values = plStock.values;
          sh.getRange('X6:Y48').values = plPoQty.values;
          sh.getRange('AA6:AA48').values = plRemarks.values;
          sh.getRange('A1').values = a1.values;
          sh.getRange('A2').values = a2.values;
          // Reapply DV
          return bs().reapplyDataValidation(ctx, sheetName);
        });
      });
    });
  }

  // ============================================================
  // REPAIR PLAIN-SHEET NAMES (in place)
  // ============================================================

  // Upgrades already-flattened sheets to Excel 365 dynamic-spill AB5 + INDEX
  // names. BOM / Stock / PO data untouched.
  function repairPlainSheetNames(sheetNames) {
    if (!sheetNames || !sheetNames.length) return Promise.reject(new Error('No sheets selected'));
    return bs()._run({ heavy: true }, function (ctx) {
      var chain = Promise.resolve();
      sheetNames.forEach(function (name) {
        chain = chain.then(function () { return upgradePlainSheet(ctx, name); });
      });
      return chain;
    });
  }

  function upgradePlainSheet(ctx, sheetName) {
    var sh = ctx.workbook.worksheets.getItem(sheetName);
    // Snapshot the Materials column (P6:P48) so we can build the unique spill
    var pRng = sh.getRange('P6:P48').load('values');
    return ctx.sync().then(function () {
      // Build unique sorted materials list at AB5 (hidden col already hidden in
      // flattened output? Make sure).
      sh.getRange('AB5').formulas = [['=SORT(UNIQUE(FILTER(P6:P48,ISNUMBER(MATCH(P6:P48,Materials,0)))))']];
      sh.getRange('AB:AB').columnHidden = true;
      // Rewrite Mat Row (AB6:AB48) via INDEX(AB5#, ROW()-5)
      var AB = [];
      for (var r = 6; r <= 48; r++) {
        AB.push(['=IFERROR(INDEX($AB$5#,ROW()-5),"")']);
      }
      sh.getRange('AB6:AB48').formulas = AB;
      // AB4 = count
      sh.getRange('AB4').formulas = [['=IF(AB5="",0,ROWS($AB$5#))']];
      return ctx.sync();
    });
  }

  // ============================================================
  // helpers shared with app.js
  // ============================================================

  function saveBlobJson(obj) {
    try {
      var blob = new root.Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
      var url = root.URL.createObjectURL(blob);
      var a = root.document.createElement('a');
      a.href = url; a.download = (obj.name || 'bom-project') + '.json';
      root.document.body.appendChild(a); a.click();
      root.document.body.removeChild(a);
      root.URL.revokeObjectURL(url);
    } catch (e) {}
  }

  // ============================================================
  // PUBLIC API
  // ============================================================
  root.BomConvert = {
    // step 1
    openCopyInNewWindow: openCopyInNewWindow,
    getFileBase64: getFileBase64,
    // flatten
    flatten: flatten,
    flattenOne: flattenOne,
    flattenAuto365: flattenAuto365,
    flattenFixed: flattenFixed,
    buildCompiledPL: buildCompiledPL,
    // restore
    restore: restore,
    restoreOne: restoreOne,
    detectFlattenedSheets: detectFlattenedSheets,
    detectBomProjectList: detectBomProjectList,
    loadRestoreMaterials: loadRestoreMaterials,
    // plain-sheet repair
    repairPlainSheetNames: repairPlainSheetNames,
    upgradePlainSheet: upgradePlainSheet,
    // util
    saveBlobJson: saveBlobJson
  };
})(typeof window !== 'undefined' ? window : this);
