/* sync.js - shared materials list on GitHub + template import for BOM Tools.
 * Reads materials.json / template.xlsx from the same GitHub Pages site the add-in is served from.
 * Publishing (writing materials.json) uses the GitHub API and needs a token, kept only in this computer's add-in storage.
 */
(function (root) {
  "use strict";

  var CFG = {
    owner: "", repo: "",            // leave empty: detected from the page address (<owner>.github.io/<repo>/)
    branch: "main",
    list: "materials.json",
    template: "template.xlsx"
  };
  var TOKEN_KEY = "bomGhToken";
  var FIELDS = ["category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"];
  var INPUT_KEYS = ["material"].concat(FIELDS);

  function noop() {}
  function lc(s) { return String(s).trim().toLowerCase(); }
  function isFormula(f) { return typeof f === "string" && f.charAt(0) === "="; }
  function isNil(v) { return v === null || v === undefined || (typeof v === "number" && isNaN(v)); }

  /* ---------------- pure helpers (node-testable) ---------------- */
  function sameVal(a, b) {
    if (isNil(a) || a === "") a = null;
    if (isNil(b) || b === "") b = null;
    if (a === null || b === null) return a === b;
    if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) < 1e-9;
    return String(a).trim() === String(b).trim();
  }

  function changedFields(a, b) {
    return FIELDS.filter(function (k) { return !sameVal(a[k], b[k]); });
  }

  /* wb / remote: arrays of normalized materials. */
  function diff(wb, remote) {
    var w = new Map(), r = new Map();
    wb.forEach(function (m) { if (!w.has(lc(m.material))) w.set(lc(m.material), m); });
    remote.forEach(function (m) { if (!r.has(lc(m.material))) r.set(lc(m.material), m); });
    var d = { onlyWb: [], onlyRemote: [], changed: [] };
    w.forEach(function (m, k) {
      if (!r.has(k)) { d.onlyWb.push(m); return; }
      var f = changedFields(m, r.get(k));
      if (f.length) d.changed.push({ material: m.material, fields: f, wb: m, remote: r.get(k) });
    });
    r.forEach(function (m, k) { if (!w.has(k)) d.onlyRemote.push(m); });
    return d;
  }

  function clean(m) {
    var o = { material: m.material, category: m.category || "", unit: m.unit };
    ["lin", "area", "piece", "sl", "sw"].forEach(function (k) { o[k] = isNil(m[k]) ? null : m[k]; });
    o.remarks = m.remarks || "";
    return o;
  }

  function serialize(list) {
    return '{\n "version": 1,\n "updated": "' + new Date().toISOString() + '",\n "materials": [\n' +
      list.map(function (m) { return "  " + JSON.stringify(clean(m)); }).join(",\n") + "\n ]\n}\n";
  }

  function parseList(text) {
    var j = JSON.parse(text);
    var arr = Array.isArray(j) ? j : (j.materials || []);
    return arr.map(function (m) { return root.BomTools.normalizeMaterial(m); }).filter(function (m) { return m.material; });
  }

  function toB64(str) {
    var b = new TextEncoder().encode(str), s = "";
    for (var i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
    return btoa(s);
  }
  function fromB64(b64) {
    var s = atob(String(b64).replace(/\s/g, "")), u = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) u[i] = s.charCodeAt(i);
    return new TextDecoder().decode(u);
  }
  function bufToB64(buf) {
    var u = new Uint8Array(buf), s = "";
    for (var i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
    return btoa(s);
  }

  /* ---------------- GitHub ---------------- */
  function repoInfo() {
    var o = CFG.owner, r = CFG.repo;
    if ((!o || !r) && typeof location !== "undefined" && /\.github\.io$/i.test(location.hostname)) {
      var host = location.hostname.split(".")[0];
      var seg = location.pathname.split("/").filter(Boolean)[0];
      o = o || host;
      r = r || (seg && seg.indexOf(".") < 0 ? seg : host + ".github.io");
    }
    return { owner: o, repo: r, branch: CFG.branch };
  }

  function getToken() { try { return localStorage.getItem(TOKEN_KEY) || ""; } catch (e) { return ""; } }
  function setToken(t) {
    try { if (t) localStorage.setItem(TOKEN_KEY, t.trim()); else localStorage.removeItem(TOKEN_KEY); return true; }
    catch (e) { return false; }
  }
  function hasToken() { return !!getToken(); }

  function apiUrl() {
    var ri = repoInfo();
    if (!ri.owner || !ri.repo) throw new Error("Could not work out the GitHub repository. Set owner and repo at the top of sync.js.");
    return "https://api.github.com/repos/" + ri.owner + "/" + ri.repo + "/contents/" + CFG.list;
  }
  function apiHeaders() {
    return { "Accept": "application/vnd.github+json", "Authorization": "Bearer " + getToken() };
  }

  /* what everyone reads: the file as served by GitHub Pages */
  async function fetchPages() {
    var url = new URL(CFG.list, location.href).href + "?t=" + Date.now();
    var res = await fetch(url, { cache: "no-store" });
    if (res.status === 404) return { list: [], missing: true };
    if (!res.ok) throw new Error("GitHub Pages answered " + res.status + " for " + CFG.list + ".");
    return { list: parseList(await res.text()), missing: false };
  }

  /* what the publisher reads before writing: the file straight from the repository (always the latest commit) */
  async function fetchApi() {
    var ri = repoInfo();
    var res = await fetch(apiUrl() + "?ref=" + encodeURIComponent(ri.branch) + "&t=" + Date.now(), { headers: apiHeaders(), cache: "no-store" });
    if (res.status === 404) return { list: [], sha: null };
    if (res.status === 401 || res.status === 403) throw new Error("GitHub rejected the token (needs Contents: Read and write on " + ri.owner + "/" + ri.repo + ").");
    if (!res.ok) throw new Error("GitHub API answered " + res.status + ".");
    var j = await res.json();
    return { list: parseList(fromB64(j.content)), sha: j.sha };
  }

  async function putApi(text, sha, message) {
    var ri = repoInfo();
    var body = { message: message, content: toB64(text), branch: ri.branch };
    if (sha) body.sha = sha;
    var h = apiHeaders(); h["Content-Type"] = "application/json";
    return fetch(apiUrl(), { method: "PUT", headers: h, body: JSON.stringify(body) });
  }

  /* ---------------- workbook <-> list ---------------- */
  async function workbookList() {
    var res = await root.BomTools.listMaterials();
    return res.rows.map(function (r) { return root.BomTools.normalizeMaterial(r); });
  }

  async function compare() {
    var wb = await workbookList();
    var rem = await fetchPages();
    var d = diff(wb, rem.list);
    d.remoteMissing = rem.missing;
    d.remoteCount = rem.list.length;
    d.wbCount = wb.length;
    return d;
  }

  function writeRow(body, i, r, idx) {
    INPUT_KEYS.forEach(function (k) {
      if (idx[k] < 0) return;
      var v = r[k];
      body.getCell(i, idx[k]).values = [[isNil(v) ? "" : v]];
    });
  }

  /* bring the workbook's master list up to date with `list` (adds missing, updates different; never deletes) */
  async function applyList(list, log) {
    log = log || noop;
    var T = root.BomTools, res = { added: 0, updated: 0, skipped: [] };
    await Excel.run(async function (context) {
      var m = await T._openMaster(context);
      var idx = m.idx;
      var body = m.t.getDataBodyRange().load("values,formulas");
      await context.sync();

      var byName = new Map(), blanks = [];
      body.values.forEach(function (row, i) {
        var n = String(row[idx.material]).trim();
        if (!n) blanks.push(i); else if (!byName.has(n.toLowerCase())) byName.set(n.toLowerCase(), i);
      });

      var toAdd = [];
      list.forEach(function (r) {
        var i = byName.get(r.material.toLowerCase());
        if (i === undefined) { toAdd.push(r); return; }
        var cur = T.normalizeMaterial(T._rowToObj(body.values[i], i, idx));
        var ch = changedFields(cur, r), did = false;
        ch.forEach(function (k) {
          if (idx[k] < 0) return;
          if (isFormula(body.formulas[i][idx[k]])) { res.skipped.push(r.material + " (" + k + " is a formula here)"); return; }
          body.getCell(i, idx[k]).values = [[isNil(r[k]) ? "" : r[k]]];
          did = true;
        });
        if (did) res.updated++;
      });

      // use empty rows of the table first, then add rows
      var useBlank = toAdd.slice(0, blanks.length), extra = toAdd.slice(blanks.length);
      useBlank.forEach(function (r, k) { writeRow(body, blanks[k], r, idx); });
      await context.sync();

      if (extra.length) {
        for (var k = 0; k < extra.length; k++) m.t.rows.add();
        await context.sync();
        var b2 = m.t.getDataBodyRange().load("rowCount");
        await context.sync();
        var n = b2.rowCount, first = n - extra.length;
        extra.forEach(function (r, j) { writeRow(b2, first + j, r, idx); });
        await context.sync();

        // calculated columns (sheet area / weight): copy from the row above if Excel did not fill them
        if (first >= 1) {
          var f = m.t.getDataBodyRange().load("formulas");
          await context.sync();
          var isInput = {};
          INPUT_KEYS.forEach(function (key) { if (idx[key] >= 0) isInput[idx[key]] = 1; });
          for (var j2 = 0; j2 < extra.length; j2++)
            for (var c = 0; c < m.ncols; c++) {
              if (isInput[c]) continue;
              if (isFormula(f.formulas[first - 1][c]) && !isFormula(f.formulas[first + j2][c]))
                b2.getCell(first + j2, c).copyFrom(b2.getCell(first - 1, c), Excel.RangeCopyType.formulas);
            }
          await context.sync();
        }
      }
      res.added = toAdd.length;
    });
    return res;
  }

  /* ---------------- pull / push ---------------- */
  async function pull(log) {
    log = log || noop;
    log("Reading the shared list from GitHub ...");
    var rem = await fetchPages();
    if (!rem.list.length) { log("The shared list is empty or not published yet. Nothing to update."); return { added: 0, updated: 0, skipped: [] }; }
    var r = await applyList(rem.list, log);
    log("Master list updated: " + r.added + " added, " + r.updated + " changed.");
    r.skipped.forEach(function (s) { log("Skipped " + s); });
    return r;
  }

  /* opts.only: publish just this material name (used right after a save) */
  async function push(log, opts) {
    log = log || noop;
    opts = opts || {};
    if (!hasToken()) throw new Error("No GitHub token on this computer. Add one under Publisher settings.");
    var wb = await workbookList();
    for (var attempt = 0; attempt < 2; attempt++) {
      var cur = await fetchApi();
      var d = diff(wb, cur.list);
      var only = opts.only ? [].concat(opts.only).map(lc) : null;
      var add = d.onlyWb.filter(function (m) { return !only || only.indexOf(lc(m.material)) >= 0; });
      var chg = d.changed.filter(function (c) { return !only || only.indexOf(lc(c.material)) >= 0; });
      if (!add.length && !chg.length) { log("GitHub already has this."); return { added: 0, changed: 0 }; }
      var swap = new Map();
      chg.forEach(function (c) { swap.set(lc(c.material), c.wb); });
      var merged = cur.list.map(function (m) { return swap.get(lc(m.material)) || m; }).concat(add);
      var msg = "Materials: " + add.length + " added, " + chg.length + " changed";
      var res = await putApi(serialize(merged), cur.sha, msg);
      if (res.ok) {
        log("Published to GitHub (" + add.length + " added, " + chg.length + " changed). Other computers see it in about a minute.");
        return { added: add.length, changed: chg.length };
      }
      if ((res.status === 409 || res.status === 422) && attempt === 0) continue;   // someone else published meanwhile: re-read and retry
      if (res.status === 401 || res.status === 403 || res.status === 404)
        throw new Error("GitHub refused the write (" + res.status + "). Check the token has Contents: Read and write on this repository.");
      throw new Error("GitHub answered " + res.status + " while publishing.");
    }
    throw new Error("Could not publish after retrying. Try again.");
  }

  /* ---------------- template import ---------------- */
  async function importTemplate(log) {
    log = log || noop;
    var T = root.BomTools;
    if (!Office.context.requirements.isSetSupported("ExcelApi", "1.13"))
      throw new Error("Importing the template needs Microsoft 365 or Excel 2021 (ExcelApi 1.13). Use a workbook that already has the template instead.");
    var b64 = root.BOM_TEMPLATE_B64;                       // embedded in template-data.js
    if (!b64) {
      log("Downloading the template ...");
      var res = await fetch(new URL(CFG.template, location.href).href + "?t=" + Date.now(), { cache: "no-store" });
      if (!res.ok) throw new Error("No embedded template and " + CFG.template + " was not found (" + res.status + ").");
      b64 = bufToB64(await res.arrayBuffer());
    }

    await Excel.run(async function (context) {
      var wb = context.workbook;
      wb.worksheets.load("items/name");
      await context.sync();
      var before = wb.worksheets.items.map(function (w) { return w.name; });

      log("Importing the master list and the blank BOM sheet ...");
      wb.insertWorksheetsFromBase64(b64, { positionType: Excel.WorksheetPositionType.end });
      await context.sync();
      wb.worksheets.load("items/name");
      await context.sync();
      var added = wb.worksheets.items.filter(function (w) { return before.indexOf(w.name) < 0; });

      var undo = async function (msg) {
        added.forEach(function (w) { w.delete(); });
        try { await context.sync(); } catch (e) { /* ignore */ }
        throw new Error(msg);
      };

      added.forEach(function (w) { w.tables.load("items/name"); });
      await context.sync();
      var hasT = function (w, p) { return w.tables.items.some(function (t) { return t.name.indexOf(p) === 0; }); };
      var bomWs = added.filter(function (w) { return hasT(w, "bom_"); });
      var master = added.filter(function (w) { return lc(w.name) === lc(T.MASTER); })[0] ||
                   added.filter(function (w) { return w.tables.items.some(function (t) { return t.name === "tblMat"; }); })[0];
      var blank = added.filter(function (w) { return lc(w.name) === lc(T.BLANK); })[0] ||
                  bomWs.filter(function (w) { return /blank|template|empty/i.test(w.name); })[0] ||
                  (bomWs.length === 1 ? bomWs[0] : null);
      if (!blank || !master)
        await undo("Could not find the " + (!master ? "master list (a sheet with table tblMat)" : "blank BOM sheet (a sheet with a bom_ table)") +
          " in template.xlsx. Sheets found in it: " + (added.map(function (w) { return "\"" + w.name + "\""; }).join(", ") || "none") + ". The import was undone.");
      if (master.name !== T.MASTER) master.name = T.MASTER;
      if (blank.name !== T.BLANK) blank.name = T.BLANK;
      await context.sync();

      var rng = blank.getRange("A1:Z80").load("formulas");
      await context.sync();
      var broken = rng.formulas.some(function (row) {
        return row.some(function (f) { return typeof f === "string" && (f.indexOf("#REF!") >= 0 || /\[\d+\]/.test(f)); });
      });
      if (broken)
        await undo("The imported blank sheet has broken formulas (#REF! or a link to another file). The import was undone; see the README note on named ranges.");

      // the template's workbook-level names (used by the BOM formulas and dropdowns); recreate any that did not come across
      var NAMES = { Materials: "tblMat[Material]", LinWt: "tblMat[Linear Weight (kg/m)]", AreaWt: "tblMat[Area Weight (kg/m²)]",
                    PieceWt: "tblMat[Piece Weight (kg/each)]", WtUnit: "tblMat[Primary Unit]", SheetL: "tblMat[Sheet Length (mm)]", SheetW: "tblMat[Sheet Width (mm)]" };
      try {
        wb.names.load("items/name");
        await context.sync();
        var have = wb.names.items.map(function (n) { return n.name.toLowerCase(); });
        Object.keys(NAMES).forEach(function (k) { if (have.indexOf(k.toLowerCase()) < 0) wb.names.add(k, "=" + NAMES[k]); });
        await context.sync();
      } catch (e) { log("WARNING: could not create the named ranges (" + (e.message || e) + ")."); }

      added.forEach(function (w) { if (/^README/i.test(w.name)) w.delete(); });   // the template's help sheet is not needed in project files
      blank.visibility = Excel.SheetVisibility.hidden;
      await context.sync();
      await T._openMaster(context);                                                // confirms tblMat is there
    });
    log("Template imported.");
  }

  /* ---------------- auto-publish when the Master sheet changes (while the task pane is open) ---------------- */
  var watcher = { on: false, snap: null, timer: null, busy: false };

  async function watch(log) {
    log = log || noop;
    if (watcher.on || !hasToken()) return;
    var T = root.BomTools;
    try {
      await Excel.run(async function (context) {
        var ws = context.workbook.worksheets.getItemOrNullObject(T.MASTER);
        await context.sync();
        if (ws.isNullObject) return;
        watcher.snap = await workbookList();
        ws.onChanged.add(function () {
          clearTimeout(watcher.timer);
          watcher.timer = setTimeout(async function () {
            if (watcher.busy || !hasToken()) return;
            watcher.busy = true;
            try {
              var cur = await workbookList();
              var d = diff(cur, watcher.snap);
              var names = d.onlyWb.map(function (m) { return m.material; }).concat(d.changed.map(function (c) { return c.material; }));
              if (names.length) {
                await push(log, { only: names });
                watcher.snap = cur;
              }
            } catch (e) { log("Auto-publish failed: " + (e.message || e)); }
            watcher.busy = false;
          }, 2500);
        });
        await context.sync();
        watcher.on = true;
      });
    } catch (e) { log("Could not watch the Master sheet: " + (e.message || e)); }
  }

  var api = {
    watch: watch, CFG: CFG, hasToken: hasToken, getToken: getToken, setToken: setToken,
    compare: compare, pull: pull, push: push, importTemplate: importTemplate, repoInfo: repoInfo,
    // pure helpers
    diff: diff, changedFields: changedFields, serialize: serialize, parseList: parseList
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomSync = api;
})(typeof window !== "undefined" ? window : this);
