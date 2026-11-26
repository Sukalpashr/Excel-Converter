/* sync.js - always-on Built-in Add-in Storage (OPFS + IndexedDB + localStorage, zero reconnects)
 * + JSON backup/transfer + optional GitHub sync + clean template import for BOM Tools.
 *
 * Built-in Computer Storage flow:
 * 1. Materials are saved automatically inside Excel's persistent add-in storage on this computer
 *    (Origin Private File System + IndexedDB + localStorage) — no file picker, no reconnect prompts, no token.
 * 2. Every add, edit, or removal updates built-in storage automatically (including edits made directly on the Master sheet).
 * 3. Whenever you create a new BOM sheet, add a BOM to a non-BOM workbook, or restore a converted BOM,
 *    the workbook's Master Unique Materials table syncs automatically from built-in storage.
 * 4. Export / Import JSON anytime to back up your list or transfer it to another computer.
 *
 * Project-only lists (one-time materials for a single job):
 * - A workbook can be switched to its own project list (flag stored in the workbook's custom properties).
 * - While the flag is on, EVERY save (panel edits, edits typed on the Master sheet, imports) goes to a separate
 *   per-project storage slot and a downloadable project file; the shared master list is never read or written,
 *   and nothing is published to GitHub.
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
  var LOCAL_CACHE_KEY = "bomLocalMasterJson";
  var LOCAL_META_KEY = "bomLocalMasterMeta";
  var IDB_NAME = "BomToolsLocalDB";
  var IDB_STORE = "kv";
  var IDB_HANDLE_KEY = "localMasterHandle";
  var IDB_JSON_KEY = "localMasterJson";
  var OPFS_FILE = "master-materials.json";
  var PROJECT_PROP = "BomProjectList";            // workbook custom property: {"name": "...", "created": "..."}
  var PROJECT_MARK = "PROJECT-ONLY list";         // prefix of the note written to A3 of the Master sheet
  var projectState = { loaded: false, name: "" };

  function slug(s) {
    var t = String(s || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 60);
    return t || "project";
  }
  /* Which storage slot is in use: the shared master, or the separate slot of this workbook's project list */
  function targetFor(name) {
    if (!name) return { project: false, name: "", cache: LOCAL_CACHE_KEY, meta: LOCAL_META_KEY, idb: IDB_JSON_KEY, opfs: OPFS_FILE };
    var k = slug(name);
    return { project: true, name: name, cache: "bomProjectJson:" + k, meta: "bomProjectMeta:" + k, idb: "projectJson:" + k, opfs: "project-" + k + ".json" };
  }
  function currentTarget() { return targetFor(projectState.name); }

  var FIELDS = ["category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"];
  var INPUT_KEYS = ["material"].concat(FIELDS);
  var AUDIT_KEYS = ["lastAction", "lastModified"];

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
    if (m.lastAction) o.lastAction = m.lastAction;
    if (m.lastModified) o.lastModified = m.lastModified;
    return o;
  }

  function buildLastChangeMeta(info) {
    if (!info || !info.items || !info.items.length) return null;
    var time = info.time || new Date().toISOString();
    var items = info.items.map(function (it) {
      var name = typeof it.material === "string" ? it.material : (it.material && it.material.material) || "";
      var type = it.type === "edited" ? "modified" : (it.type || "modified");
      var entry = { material: name, type: type, time: it.time || time };
      if (it.fields && it.fields.length) entry.fields = it.fields.slice();
      return entry;
    }).filter(function (e) { return e.material; });
    if (!items.length) return null;
    var types = {};
    items.forEach(function (e) { types[e.type] = 1; });
    var typeList = Object.keys(types);
    var summaryType = typeList.length === 1 ? typeList[0] : "mixed";
    return {
      time: time,
      type: summaryType,
      source: info.source || "panel",
      items: items
    };
  }

  function serialize(list, lastChange, extra) {
    var nowIso = new Date().toISOString();
    var lcMeta = lastChange && lastChange.items ? buildLastChangeMeta(lastChange) : (lastChange && lastChange.time ? lastChange : null);
    var out = '{\n "version": 1,\n';
    if (extra && extra.type) out += ' "type": ' + JSON.stringify(extra.type) + ',\n';
    if (extra && extra.project) out += ' "project": ' + JSON.stringify(extra.project) + ',\n';
    out += ' "updated": "' + nowIso + '",\n';
    if (lcMeta) {
      out += ' "lastChange": ' + JSON.stringify(lcMeta) + ',\n';
    }
    out += ' "materials": [\n' +
      list.map(function (m) { return "  " + JSON.stringify(clean(m)); }).join(",\n") + "\n ]\n}\n";
    return out;
  }

  function parsePayload(text) {
    var j = JSON.parse(text);
    var arr = Array.isArray(j) ? j : (j.materials || []);
    var list = arr.map(function (m) { return root.BomTools.normalizeMaterial(m); }).filter(function (m) { return m.material; });
    return {
      list: list,
      updated: (j && !Array.isArray(j) && j.updated) || "",
      type: (j && !Array.isArray(j) && j.type) || "",
      project: (j && !Array.isArray(j) && j.project) || "",
      lastChange: (j && !Array.isArray(j) && j.lastChange) || null
    };
  }

  function parseList(text) {
    return parsePayload(text).list;
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

  /* ---------------- Built-in Add-in Storage (OPFS + IndexedDB + localStorage — Zero Reconnects) ---------------- */
  function supportsOpfs() {
    return typeof navigator !== "undefined" &&
      navigator.storage &&
      typeof navigator.storage.getDirectory === "function";
  }

  function supportsSavePicker() {
    return typeof window !== "undefined" && typeof window.showSaveFilePicker === "function";
  }

  function requestPersistentStorage() {
    try {
      if (typeof navigator !== "undefined" && navigator.storage && typeof navigator.storage.persist === "function") {
        navigator.storage.persist().catch(noop);
      }
    } catch (e) { /* ignore */ }
  }

  async function opfsWriteJson(json, file) {
    if (!supportsOpfs()) return false;
    try {
      var dir = await navigator.storage.getDirectory();
      var fh = await dir.getFileHandle(file || OPFS_FILE, { create: true });
      var w = await fh.createWritable();
      await w.write(json);
      await w.close();
      return true;
    } catch (e) { return false; }
  }

  async function opfsReadJson(file) {
    if (!supportsOpfs()) return null;
    try {
      var dir = await navigator.storage.getDirectory();
      var fh = await dir.getFileHandle(file || OPFS_FILE);
      var f = await fh.getFile();
      return await f.text();
    } catch (e) { return null; }
  }

  function openLocalDb() {
    return new Promise(function (resolve) {
      if (typeof indexedDB === "undefined") return resolve(null);
      try {
        var req = indexedDB.open(IDB_NAME, 1);
        req.onupgradeneeded = function (ev) {
          var db = ev.target.result;
          if (!db.objectStoreNames.contains(IDB_STORE)) db.createObjectStore(IDB_STORE);
        };
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { resolve(null); };
      } catch (e) { resolve(null); }
    });
  }

  async function idbGet(key) {
    var db = await openLocalDb();
    if (!db) return null;
    return new Promise(function (resolve) {
      try {
        var tx = db.transaction(IDB_STORE, "readonly");
        var req = tx.objectStore(IDB_STORE).get(key);
        req.onsuccess = function () { resolve(req.result === undefined ? null : req.result); };
        req.onerror = function () { resolve(null); };
      } catch (e) { resolve(null); }
    });
  }

  async function idbSet(key, val) {
    var db = await openLocalDb();
    if (!db) return false;
    return new Promise(function (resolve) {
      try {
        var tx = db.transaction(IDB_STORE, "readwrite");
        tx.objectStore(IDB_STORE).put(val, key);
        tx.oncomplete = function () { resolve(true); };
        tx.onerror = function () { resolve(false); };
      } catch (e) { resolve(false); }
    });
  }

  async function idbDel(key) {
    var db = await openLocalDb();
    if (!db) return false;
    return new Promise(function (resolve) {
      try {
        var tx = db.transaction(IDB_STORE, "readwrite");
        tx.objectStore(IDB_STORE).delete(key);
        tx.oncomplete = function () { resolve(true); };
        tx.onerror = function () { resolve(false); };
      } catch (e) { resolve(false); }
    });
  }

  function getCachedMeta(t) {
    t = t || currentTarget();
    try {
      var raw = localStorage.getItem(t.meta);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  }

  async function saveLocalCache(list, lastChange, t) {
    t = t || currentTarget();
    requestPersistentStorage();
    var existingMeta = getCachedMeta(t);
    var lcMeta = lastChange ? (buildLastChangeMeta(lastChange) || lastChange) : (existingMeta && existingMeta.lastChange) || null;
    var json = serialize(list, lcMeta, t.project ? { type: "project", project: t.name } : null);
    var meta = {
      count: list.length,
      updated: new Date().toISOString(),
      lastChange: lcMeta
    };
    try {
      localStorage.setItem(t.cache, json);
      localStorage.setItem(t.meta, JSON.stringify(meta));
    } catch (e) { /* storage may be restricted */ }
    await Promise.all([
      idbSet(t.idb, json),
      opfsWriteJson(json, t.opfs)
    ]);
    // Clean up any legacy external file handle from earlier versions so Excel never asks to reconnect
    try { await idbDel(IDB_HANDLE_KEY); } catch (e) { /* ignore */ }
    return { json: json, meta: meta };
  }

  async function readLocalCache(t) {
    t = t || currentTarget();
    var json = await opfsReadJson(t.opfs);
    if (!json) json = await idbGet(t.idb);
    if (!json) {
      try { json = localStorage.getItem(t.cache); } catch (e) { json = null; }
    }
    if (!json) return null;
    try {
      var parsed = parsePayload(json);
      return { list: parsed.list, updated: parsed.updated, lastChange: parsed.lastChange, meta: getCachedMeta(t) };
    } catch (e) { return null; }
  }

  async function getLocalFileStatus() {
    var cache = await readLocalCache();
    var count = cache && cache.list ? cache.list.length : 0;
    return {
      ready: true,
      hasCache: count > 0,
      cachedCount: count,
      cachedUpdated: (cache && (cache.updated || (cache.meta && cache.meta.updated))) || "",
      lastChange: (cache && cache.lastChange) || (cache && cache.meta && cache.meta.lastChange) || null,
      opfsSupported: supportsOpfs(),
      project: projectState.name
    };
  }

  /* Always saves automatically to Built-in Add-in Storage (OPFS + IndexedDB + localStorage) with zero prompts */
  async function writeLocalMaster(list, log, opts) {
    log = log || noop;
    opts = opts || {};
    var normalized = list.map(function (m) { return root.BomTools.normalizeMaterial(m); }).filter(function (m) { return m.material; });
    var saved = await saveLocalCache(normalized, opts.lastChange || null);
    return {
      savedToBuiltIn: true,
      count: normalized.length,
      json: saved.json
    };
  }

  /* Reads the authoritative Built-in Add-in Storage list (auto-connected, never prompts) */
  async function readLocalMaster() {
    var c = await readLocalCache();
    if (c && c.list && c.list.length) {
      return { list: c.list, lastChange: c.lastChange || null, source: projectState.name ? "project" : "builtin" };
    }
    return { list: [], lastChange: null, source: "none" };
  }

  /* Seeds Built-in Storage once from the open workbook if storage is currently empty */
  async function ensureSeededFromWorkbook() {
    var c = await readLocalCache();
    if (c && c.list && c.list.length) return { seeded: false, count: c.list.length };
    try {
      var wb = await workbookList();
      if (wb.length) {
        await writeLocalMaster(wb, noop, {});
        return { seeded: true, count: wb.length };
      }
    } catch (e) { /* workbook may not have Master sheet yet */ }
    return { seeded: false, count: 0 };
  }

  /* Save current workbook's Master Unique Materials into Built-in Storage */
  async function saveWorkbookToLocal(log, opts) {
    log = log || noop;
    var wb = await workbookList();
    var res = await writeLocalMaster(wb, log, opts);
    log("Saved " + res.count + " materials to this computer's built-in storage.");
    return res;
  }

  /* Sync current workbook's Master Unique Materials from Built-in Storage */
  async function pullLocal(log, opts) {
    log = log || noop;
    opts = opts || {};
    var local = await readLocalMaster();
    if (!local.list.length) {
      return { added: 0, updated: 0, removed: 0, skipped: [], source: "none", empty: true };
    }
    log("Syncing master list from built-in computer storage (" + local.list.length + " materials) ...");
    var r = await applyList(local.list, log, { syncRemovals: opts.syncRemovals !== false, lastChange: local.lastChange });
    log("Master list synced: " + r.added + " added, " + r.updated + " changed" + (r.removed ? ", " + r.removed + " removed" : "") + ".");
    r.skipped.forEach(function (s) { log("Skipped " + s); });
    r.source = "builtin";
    return r;
  }

  /* Build JSON text for backup/transfer export (prefers live workbook list + latest lastChange) */
  async function buildExportJson() {
    var cache = await readLocalCache();
    var lastChange = (cache && cache.lastChange) || (cache && cache.meta && cache.meta.lastChange) || null;
    var list = [];
    try {
      list = await workbookList();
    } catch (e) {
      list = (cache && cache.list) || [];
    }
    if (!list.length && cache && cache.list) list = cache.list;
    var asProject = projectState.name ? { type: "project", project: projectState.name } : null;
    return {
      json: serialize(list, lastChange, asProject), count: list.length,
      fileName: asProject ? "project-materials-" + slug(projectState.name) + ".json" : "master-materials.json"
    };
  }

  /* Import a JSON backup/transfer file into Built-in Storage and apply to the workbook (if Master sheet exists) */
  async function importLocalJsonText(text, fileName, log) {
    log = log || noop;
    var parsed = parsePayload(text);
    if (!parsed.list.length) throw new Error("No materials found in that JSON file.");
    if (parsed.type === "project" && !projectState.name) {
      throw new Error("That is a project materials file. Use Materials > Project-only materials > Load project list file, so your shared master list stays untouched.");
    }
    await saveLocalCache(parsed.list, parsed.lastChange);
    try {
      var r = await applyList(parsed.list, log, { syncRemovals: true, lastChange: parsed.lastChange });
      log("Imported \"" + (fileName || "master-materials.json") + "\" (" + parsed.list.length + " materials): " +
        r.added + " added, " + r.updated + " changed, " + r.removed + " removed.");
      return r;
    } catch (e) {
      log("Saved " + parsed.list.length + " materials from \"" + (fileName || "master-materials.json") + "\" to this computer's built-in storage. They will load automatically into your next BOM.");
      return { added: 0, updated: 0, removed: 0, skipped: [], savedOnly: true };
    }
  }

  /* ---------------- GitHub (optional) ---------------- */
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

  async function compareLocal() {
    var wb = await workbookList();
    var local = await readLocalMaster();
    var d = diff(wb, local.list);
    d.localSource = local.source;
    d.localCount = local.list.length;
    d.wbCount = wb.length;
    return d;
  }

  function writeRow(body, i, r, idx) {
    var T = root.BomTools;
    INPUT_KEYS.forEach(function (k) {
      if (idx[k] < 0) return;
      var v = r[k];
      body.getCell(i, idx[k]).values = [[isNil(v) ? "" : v]];
    });
    if (idx.lastAction >= 0 && r.lastAction) {
      body.getCell(i, idx.lastAction).values = [[T.actionDisplay(r.lastAction) || r.lastAction]];
    }
    if (idx.lastModified >= 0 && r.lastModified) {
      body.getCell(i, idx.lastModified).values = [[T.formatSheetTimestamp(r.lastModified)]];
    }
  }

  /* Bring the workbook's master list up to date with `list`.
   * Adds missing materials, updates changed fields, and when opts.syncRemovals is true, removes rows deleted from the local master file.
   */
  async function applyList(list, log, opts) {
    log = log || noop;
    opts = opts || {};
    var syncRemovals = !!opts.syncRemovals;
    var T = root.BomTools, res = { added: 0, updated: 0, removed: 0, skipped: [] };
    suppressWatch(3500);
    await Excel.run(async function (context) {
      var wb = context.workbook;
      var m = await T._openMaster(context);
      var hasAuditMeta = list.some(function (r) { return r.lastAction || r.lastModified; });
      if (hasAuditMeta && T._ensureAuditColumns) await T._ensureAuditColumns(context, m);
      var idx = m.idx;
      var body = m.t.getDataBodyRange().load(["values", "formulas", "rowCount"]);
      await context.sync();

      var byName = new Map(), blanks = [];
      body.values.forEach(function (row, i) {
        var n = String(row[idx.material]).trim();
        if (!n) blanks.push(i); else if (!byName.has(n.toLowerCase())) byName.set(n.toLowerCase(), i);
      });

      var wanted = new Map();
      list.forEach(function (r) { wanted.set(lc(r.material), r); });

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
        if (did) {
          if (idx.lastAction >= 0 && r.lastAction) body.getCell(i, idx.lastAction).values = [[T.actionDisplay(r.lastAction) || r.lastAction]];
          if (idx.lastModified >= 0 && r.lastModified) body.getCell(i, idx.lastModified).values = [[T.formatSheetTimestamp(r.lastModified)]];
          res.updated++;
        }
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
          INPUT_KEYS.concat(AUDIT_KEYS).forEach(function (key) { if (idx[key] >= 0) isInput[idx[key]] = 1; });
          for (var j2 = 0; j2 < extra.length; j2++)
            for (var c = 0; c < m.ncols; c++) {
              if (isInput[c]) continue;
              if (isFormula(f.formulas[first - 1][c]) && !isFormula(f.formulas[first + j2][c]))
                b2.getCell(first + j2, c).copyFrom(b2.getCell(first - 1, c), Excel.RangeCopyType.formulas);
            }
          await context.sync();
        }
      }

      // Remove materials deleted from the authoritative local master file (iterate bottom-up)
      if (syncRemovals && wanted.size > 0) {
        var curBody = m.t.getDataBodyRange().load(["values", "rowCount"]);
        await context.sync();
        var toDelete = [];
        curBody.values.forEach(function (row, i) {
          var n = String(row[idx.material]).trim();
          if (n && !wanted.has(lc(n))) toDelete.push(i);
        });
        if (toDelete.length && toDelete.length < curBody.rowCount) {
          for (var dIdx = toDelete.length - 1; dIdx >= 0; dIdx--) {
            m.t.rows.getItemAt(toDelete[dIdx]).delete();
          }
          await context.sync();
          res.removed = toDelete.length;
        }
      }

      if (opts.lastChange && opts.lastChange.items && opts.lastChange.items.length && T._writeMasterBanner) {
        await T._writeMasterBanner(context, m.ws, opts.lastChange.items, opts.lastChange.time);
      }

      if (toAdd.length || res.updated || res.removed) {
        var matInfo = await T._ensureMasterNames(context, wb, noop);
        await T._refreshAllBomDropdownsIfRangeBased(context, wb, matInfo);
        wb.application.calculate(Excel.CalculationType.full);
        await context.sync();
      }
      res.added = toAdd.length;
    });
    return res;
  }

  /* ---------------- pull / push (GitHub) ---------------- */
  async function pull(log) {
    log = log || noop;
    if (projectState.name) { log("This workbook uses a project-only list, so the shared GitHub list is not applied."); return { added: 0, updated: 0, removed: 0, skipped: [] }; }
    log("Reading the shared list from GitHub ...");
    var rem = await fetchPages();
    if (!rem.list.length) { log("The shared list is empty or not published yet. Nothing to update."); return { added: 0, updated: 0, removed: 0, skipped: [] }; }
    var r = await applyList(rem.list, log, { syncRemovals: false });
    log("Master list updated: " + r.added + " added, " + r.updated + " changed.");
    r.skipped.forEach(function (s) { log("Skipped " + s); });
    return r;
  }

  /* opts.only: publish just this material name (used right after a save) */
  async function push(log, opts) {
    log = log || noop;
    opts = opts || {};
    if (projectState.name) { log("Project-only list: not published to the shared GitHub master."); return { added: 0, changed: 0 }; }
    if (!hasToken()) throw new Error("No GitHub token on this computer. Add one under GitHub settings.");
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

  /* ---------------- template import (BOM Blank + Master Unique Materials only; never Formulas/README) ---------------- */
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
      var prevCalcMode = null;
      try {
        wb.application.load("calculationMode");
        await context.sync();
        prevCalcMode = wb.application.calculationMode;
        wb.application.calculationMode = Excel.CalculationMode.manual;
        await context.sync();
      } catch (e) { prevCalcMode = null; }

      var added = [];
      try {
        wb.worksheets.load("items/name");
        await context.sync();
        var before = wb.worksheets.items.map(function (w) { return w.name; });

        log("Importing the master list and blank BOM sheet ...");
        wb.insertWorksheetsFromBase64(b64, { positionType: Excel.WorksheetPositionType.end });
        await context.sync();
        wb.worksheets.load("items/name");
        await context.sync();
        added = wb.worksheets.items.filter(function (w) { return before.indexOf(w.name) < 0; });

        var undo = async function (msg) {
          added.forEach(function (w) { try { w.delete(); } catch (e) {} });
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
                    bomWs.filter(function (w) { return /blank|template|empty|take-?off/i.test(w.name); })[0] ||
                    (bomWs.length === 1 ? bomWs[0] : null);
        if (!blank || !master) {
          await undo("Could not find the " + (!master ? "master list (a sheet with table tblMat)" : "blank BOM sheet (a sheet with a bom_ table)") +
            " in template.xlsx. Sheets found: " + (added.map(function (w) { return "\"" + w.name + "\""; }).join(", ") || "none") + ". The import was undone.");
        }

        // Never leave Formulas or README sheets in the workbook (even if an older external template.xlsx is used)
        added.forEach(function (w) {
          if (w !== master && w !== blank && /^(README|Formulas)$/i.test(w.name.trim())) {
            w.delete();
          }
        });
        if (master.name !== T.MASTER) master.name = T.MASTER;
        if (blank.name !== T.BLANK) blank.name = T.BLANK;
        await context.sync();

        // Create / repair all 7 workbook-level named ranges FIRST so formulas & dropdowns resolve cleanly
        log("Setting up named ranges and dropdowns ...");
        var matInfo = await T._ensureMasterNames(context, wb, log);

        // Repair BOM Blank formulas, dropdown (C6:C2000), hidden row 5 / cols M & AB, freeze panes, and print setup
        await T._repairBomSheet(context, wb, blank, matInfo, {}, log);

        // Verify no broken external links or #REF! remain on BOM Blank
        var rng = blank.getRange("A1:AB35").load("formulas");
        await context.sync();
        var broken = rng.formulas.some(function (row) {
          return row.some(function (f) { return typeof f === "string" && (f.indexOf("#REF!") >= 0 || /\[\d+\]/.test(f)); });
        });
        if (broken) {
          await undo("The imported blank sheet has broken formulas (#REF! or an external link). The import was undone.");
        }

        blank.visibility = Excel.SheetVisibility.hidden;
        await context.sync();
        await T._openMaster(context);
      } finally {
        try {
          if (prevCalcMode !== null) wb.application.calculationMode = prevCalcMode;
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        } catch (e) { /* ignore */ }
      }
    });
    log("Template imported (Master Unique Materials + hidden BOM Blank).");
  }

  /* ---------------- project-only materials list ---------------- */
  function getProject() { return { active: !!projectState.name, name: projectState.name }; }

  async function loadProjectState() {
    try {
      await Excel.run(async function (context) {
        var item = context.workbook.properties.custom.getItemOrNullObject(PROJECT_PROP);
        item.load("value");
        await context.sync();
        var name = "";
        if (!item.isNullObject) {
          try { name = String((JSON.parse(item.value) || {}).name || "").trim(); } catch (e) { name = String(item.value || "").trim(); }
        }
        projectState = { loaded: true, name: name };
      });
    } catch (e) { projectState.loaded = true; }
    return getProject();
  }

  async function setProjectProperty(name) {
    var payload = JSON.stringify({ name: name, created: new Date().toISOString() });
    await Excel.run(async function (context) {
      var props = context.workbook.properties.custom;
      var item = props.getItemOrNullObject(PROJECT_PROP);
      item.load("key");
      await context.sync();
      if (item.isNullObject) props.add(PROJECT_PROP, payload); else item.value = payload;
      await context.sync();
    });
  }

  async function clearProjectProperty() {
    await Excel.run(async function (context) {
      var item = context.workbook.properties.custom.getItemOrNullObject(PROJECT_PROP);
      item.load("key");
      await context.sync();
      if (!item.isNullObject) { item.delete(); await context.sync(); }
    });
  }

  /* Writes / removes the visible note on the Master sheet (A3) */
  async function markMasterSheet(name) {
    try {
      await Excel.run(async function (context) {
        var ws = context.workbook.worksheets.getItemOrNullObject(root.BomTools.MASTER);
        await context.sync();
        if (ws.isNullObject) return;
        var cell = ws.getRange("A3");
        if (name) {
          cell.values = [[PROJECT_MARK + ": " + name + " - changes here do not update the shared master list."]];
          cell.format.font.size = 9; cell.format.font.italic = true; cell.format.font.bold = true; cell.format.font.color = "#B45309";
        } else {
          cell.load("values");
          await context.sync();
          if (String(cell.values[0][0] || "").indexOf(PROJECT_MARK) === 0) { cell.values = [[""]]; cell.format.font.bold = false; }
        }
        await context.sync();
      });
    } catch (e) { /* non-fatal */ }
  }

  /* Switch this workbook to its own project list.
   * opts.keep: array of material names to copy from the current list (omit = copy everything). */
  async function startProject(name, opts) {
    opts = opts || {};
    var T = root.BomTools;
    name = String(name || "").trim();
    if (!name) throw new Error("Give the project list a name.");
    if (projectState.name) throw new Error("This workbook already uses the project list \"" + projectState.name + "\".");
    await workbookList();   // throws if there is no Master sheet
    var keep = Array.isArray(opts.keep) ? opts.keep.map(lc) : null;
    var prev = projectState.name, list;
    suppressWatch(9000);
    projectState.name = name; projectState.loaded = true;   // from here every save goes to the project slot, never the shared master
    try {
      if (keep) {
        await Excel.run(async function (context) {
          var wb = context.workbook;
          var m = await T._openMaster(context);
          var body = m.t.getDataBodyRange().load(["values", "rowCount"]);
          await context.sync();
          var del = [], kept = 0;
          body.values.forEach(function (row, i) {
            var n = String(row[m.idx.material]).trim();
            if (!n) return;
            if (keep.indexOf(lc(n)) >= 0) kept++; else del.push(i);
          });
          if (!kept) throw new Error("None of the picked materials were found in this workbook's list.");
          for (var k = del.length - 1; k >= 0; k--) m.t.rows.getItemAt(del[k]).delete();
          if (del.length) await context.sync();
          var matInfo = await T._ensureMasterNames(context, wb, noop);
          await T._refreshAllBomDropdownsIfRangeBased(context, wb, matInfo);
          wb.application.calculate(Excel.CalculationType.full);
          await context.sync();
        });
      }
      list = await workbookList();
      await saveLocalCache(list, null, targetFor(name));
      await setProjectProperty(name);
      await markMasterSheet(name);
      updateWatchSnap(list);
    } catch (e) {
      projectState.name = prev;
      throw e;
    }
    suppressWatch(3000);
    return { name: name, count: list.length };
  }

  /* Back to the shared master list. The project's saved list stays on this computer (and in any exported file). */
  async function stopProject(log) {
    log = log || noop;
    if (!projectState.name) return { stopped: false };
    var name = projectState.name;
    suppressWatch(9000);
    await clearProjectProperty();
    projectState.name = "";
    await markMasterSheet("");
    var res = { stopped: true, name: name, restored: false, removed: 0, added: 0 };
    try {
      var local = await readLocalMaster();
      if (local.list.length) {
        var r = await applyList(local.list, log, { syncRemovals: true, lastChange: local.lastChange });
        res.restored = true; res.removed = r.removed; res.added = r.added;
        updateWatchSnap(await workbookList());
      }
    } catch (e) { log("Note: could not reload the shared master list into this workbook (" + (e.message || e) + ")."); }
    suppressWatch(3000);
    return res;
  }

  /* JSON text for the separate project file */
  async function buildProjectExport() {
    if (!projectState.name) throw new Error("This workbook is not using a project list.");
    var name = projectState.name, list = [], cache = await readLocalCache(currentTarget());
    try { list = await workbookList(); } catch (e) { list = []; }
    if (list.length) { await saveLocalCache(list, cache && cache.lastChange, currentTarget()); }
    else if (cache && cache.list) list = cache.list;
    if (!list.length) throw new Error("The project list is empty.");
    return {
      name: name, count: list.length,
      fileName: "project-materials-" + slug(name) + ".json",
      json: serialize(list, cache && cache.lastChange, { type: "project", project: name })
    };
  }

  /* Load a saved project file. Works even in a plain/converted workbook (no Master sheet):
   * the list is kept for the next Restore Automated. */
  async function importProjectJson(text, fileName, log) {
    log = log || noop;
    var parsed = parsePayload(text);
    if (!parsed.list.length) throw new Error("No materials found in that file.");
    var useName = projectState.name ||
      String(parsed.project || "").trim() ||
      String(fileName || "").replace(/\.json$/i, "").replace(/^project-materials-?/i, "").replace(/[-_]+/g, " ").trim() || "Project";
    var prev = projectState.name, applied = true, r = null;
    suppressWatch(9000);
    projectState.name = useName; projectState.loaded = true;
    try {
      await saveLocalCache(parsed.list, parsed.lastChange, targetFor(useName));
      await setProjectProperty(useName);
      try {
        r = await applyList(parsed.list, log, { syncRemovals: true, lastChange: parsed.lastChange });
      } catch (e) {
        if (!/not found/i.test(String(e && e.message))) throw e;
        applied = false;   // converted workbook: no Master sheet yet
      }
      if (applied) { await markMasterSheet(useName); updateWatchSnap(await workbookList()); }
    } catch (e) {
      projectState.name = prev;
      throw e;
    }
    suppressWatch(3000);
    return { name: useName, count: parsed.list.length, applied: applied, result: r };
  }

  /* ---------------- auto-save when the Master sheet changes (local file + optional GitHub) ---------------- */
  var watcher = { on: false, snap: null, timer: null, busy: false, onSaved: null, ignoreUntil: 0 };

  function suppressWatch(ms) {
    watcher.ignoreUntil = Math.max(watcher.ignoreUntil, Date.now() + (ms || 2500));
  }

  function updateWatchSnap(list) {
    if (Array.isArray(list)) watcher.snap = list;
  }

  async function watch(log, onSaved) {
    log = log || noop;
    if (typeof onSaved === "function") watcher.onSaved = onSaved;
    if (watcher.on) return;
    var T = root.BomTools;
    try {
      await Excel.run(async function (context) {
        var ws = context.workbook.worksheets.getItemOrNullObject(T.MASTER);
        await context.sync();
        if (ws.isNullObject) return;
        watcher.snap = await workbookList();
        ws.onChanged.add(function () {
          if (Date.now() < watcher.ignoreUntil) return;
          clearTimeout(watcher.timer);
          watcher.timer = setTimeout(async function () {
            if (watcher.busy || Date.now() < watcher.ignoreUntil) return;
            watcher.busy = true;
            try {
              var cur = await workbookList();
              var d = diff(cur, watcher.snap || []);
              var hasChanges = d.onlyWb.length > 0 || d.onlyRemote.length > 0 || d.changed.length > 0;
              if (hasChanges) {
                var nowIso = new Date().toISOString();
                var items = [];
                var stampEntries = [];
                var byKey = new Map();
                cur.forEach(function (m) { byKey.set(lc(m.material), m); });

                d.onlyWb.forEach(function (m) {
                  m.lastAction = "added";
                  m.lastModified = nowIso;
                  items.push({ type: "added", material: m, fields: [], time: nowIso });
                  stampEntries.push({ material: m.material, type: "added", time: nowIso });
                });
                d.changed.forEach(function (c) {
                  var ref = byKey.get(lc(c.material)) || c.wb;
                  ref.lastAction = "modified";
                  ref.lastModified = nowIso;
                  items.push({ type: "modified", material: ref, fields: c.fields || [], time: nowIso });
                  stampEntries.push({ material: ref.material, type: "modified", time: nowIso });
                });
                d.onlyRemote.forEach(function (m) {
                  items.push({ type: "removed", material: m, fields: [], time: nowIso });
                  stampEntries.push({ material: m.material, type: "removed", time: nowIso });
                });

                var lastChange = buildLastChangeMeta({ time: nowIso, source: "sheet", items: items });

                // Stamp Last Action / Last Modified + Row 3 banner on the Master sheet without re-triggering watcher
                suppressWatch(3000);
                try {
                  if (T.stampMasterChanges) await T.stampMasterChanges(stampEntries, nowIso);
                } catch (eSt) { /* non-fatal */ }

                // Always save additions, modifications, and removals (with lastChange info) to local computer file + cache
                var writeRes = await writeLocalMaster(cur, log, { allowPrompt: false, lastChange: lastChange });
                if (hasToken()) {
                  var names = d.onlyWb.map(function (m) { return m.material; }).concat(d.changed.map(function (c) { return c.material; }));
                  if (names.length) {
                    try { await push(log, { only: names }); } catch (e) { log("GitHub auto-publish skipped: " + (e.message || e)); }
                  }
                }

                watcher.snap = cur;
                if (typeof watcher.onSaved === "function") {
                  watcher.onSaved({
                    source: "sheet",
                    time: nowIso,
                    savedToBuiltIn: true,
                    items: items
                  });
                }
              }
            } catch (e) { log("Auto-save failed: " + (e.message || e)); }
            watcher.busy = false;
          }, 1400);
        });
        await context.sync();
        watcher.on = true;
      });
    } catch (e) { log("Could not watch the Master sheet: " + (e.message || e)); }
  }

  var api = {
    watch: watch, updateWatchSnap: updateWatchSnap, suppressWatch: suppressWatch, CFG: CFG,
    // built-in computer storage (always auto-connected, zero reconnects) + JSON backup/transfer
    supportsOpfs: supportsOpfs,
    supportsSavePicker: supportsSavePicker,
    getLocalFileStatus: getLocalFileStatus,
    ensureSeededFromWorkbook: ensureSeededFromWorkbook,
    readLocalMaster: readLocalMaster,
    writeLocalMaster: writeLocalMaster,
    saveWorkbookToLocal: saveWorkbookToLocal,
    pullLocal: pullLocal,
    compareLocal: compareLocal,
    buildExportJson: buildExportJson,
    importLocalJsonText: importLocalJsonText,
    // GitHub optional API
    hasToken: hasToken, getToken: getToken, setToken: setToken,
    compare: compare, pull: pull, push: push, importTemplate: importTemplate, repoInfo: repoInfo,
    // pure helpers
    diff: diff, changedFields: changedFields, serialize: serialize, parseList: parseList, parsePayload: parsePayload,
    buildLastChangeMeta: buildLastChangeMeta,
    // project-only lists
    getProject: getProject, loadProjectState: loadProjectState, startProject: startProject, stopProject: stopProject,
    buildProjectExport: buildProjectExport, importProjectJson: importProjectJson,
    workbookList: workbookList, applyList: applyList
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomSync = api;
})(typeof window !== "undefined" ? window : this);
