/* sync.js - local computer master file (no token required) + optional GitHub sync + template import for BOM Tools.
 *
 * Local Master File flow:
 * 1. Pick or create a local JSON file (e.g. master-materials.json) on your computer via the File System Access API
 *    (supported in Excel for Windows WebView2 and Excel on the Web in Edge/Chrome).
 * 2. The file handle is remembered in IndexedDB and a backup snapshot is stored in local computer storage.
 * 3. Every add, edit, or removal updates the local file automatically (including edits made directly on the Master sheet).
 * 4. Whenever you create a new BOM sheet or add a BOM to a non-BOM workbook, the workbook's Master Unique Materials
 *    table is updated automatically from your local file (additions, edits, and removals) before the sheet is added.
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

  function serialize(list, lastChange) {
    var nowIso = new Date().toISOString();
    var lcMeta = lastChange && lastChange.items ? buildLastChangeMeta(lastChange) : (lastChange && lastChange.time ? lastChange : null);
    var out = '{\n "version": 1,\n "updated": "' + nowIso + '",\n';
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

  /* ---------------- IndexedDB & Local Computer Master File ---------------- */
  function supportsLocalFileApi() {
    return typeof window !== "undefined" &&
      typeof window.showOpenFilePicker === "function" &&
      typeof window.showSaveFilePicker === "function";
  }

  function openLocalDb() {
    return new Promise(function (resolve, reject) {
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

  function getCachedMeta() {
    try {
      var raw = localStorage.getItem(LOCAL_META_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  }

  async function saveLocalCache(list, fileName, lastChange) {
    var existingMeta = getCachedMeta();
    var lcMeta = lastChange ? buildLastChangeMeta(lastChange) || lastChange : (existingMeta && existingMeta.lastChange) || null;
    var json = serialize(list, lcMeta);
    var meta = {
      count: list.length,
      updated: new Date().toISOString(),
      fileName: fileName || (existingMeta && existingMeta.fileName) || "",
      lastChange: lcMeta
    };
    try {
      localStorage.setItem(LOCAL_CACHE_KEY, json);
      localStorage.setItem(LOCAL_META_KEY, JSON.stringify(meta));
    } catch (e) { /* storage may be restricted */ }
    await idbSet(IDB_JSON_KEY, json);
    return { json: json, meta: meta };
  }

  async function readLocalCache() {
    var json = null;
    try { json = localStorage.getItem(LOCAL_CACHE_KEY); } catch (e) { json = null; }
    if (!json) json = await idbGet(IDB_JSON_KEY);
    if (!json) return null;
    try {
      var parsed = parsePayload(json);
      return { list: parsed.list, lastChange: parsed.lastChange, meta: getCachedMeta() };
    } catch (e) { return null; }
  }

  async function getLocalFileHandle() {
    if (!supportsLocalFileApi()) return null;
    try { return await idbGet(IDB_HANDLE_KEY); } catch (e) { return null; }
  }

  async function queryHandlePermission(handle, mode) {
    if (!handle || typeof handle.queryPermission !== "function") return "prompt";
    try { return await handle.queryPermission({ mode: mode || "readwrite" }); }
    catch (e) { return "prompt"; }
  }

  async function requestHandlePermission(handle, mode) {
    if (!handle || typeof handle.requestPermission !== "function") return "denied";
    try { return await handle.requestPermission({ mode: mode || "readwrite" }); }
    catch (e) { return "denied"; }
  }

  async function getLocalFileStatus() {
    var supported = supportsLocalFileApi();
    var handle = await getLocalFileHandle();
    var cache = await readLocalCache();
    var perm = handle ? await queryHandlePermission(handle, "readwrite") : "none";
    return {
      supported: supported,
      linked: !!handle,
      fileName: handle ? handle.name : ((cache && cache.meta && cache.meta.fileName) || ""),
      permission: perm,
      hasCache: !!(cache && cache.list && cache.list.length),
      cachedCount: cache && cache.list ? cache.list.length : 0,
      cachedUpdated: (cache && cache.meta && cache.meta.updated) || ""
    };
  }

  async function requestLocalPermission() {
    var handle = await getLocalFileHandle();
    if (!handle) return false;
    var p = await requestHandlePermission(handle, "readwrite");
    return p === "granted";
  }

  async function unlinkLocalFile() {
    await idbDel(IDB_HANDLE_KEY);
    return true;
  }

  /* Save a list directly to the linked local file on disk + local computer cache */
  async function writeLocalMaster(list, log, opts) {
    log = log || noop;
    opts = opts || {};
    var normalized = list.map(function (m) { return root.BomTools.normalizeMaterial(m); }).filter(function (m) { return m.material; });
    var handle = await getLocalFileHandle();
    var saved = await saveLocalCache(normalized, handle ? handle.name : "", opts.lastChange || null);

    if (!handle) {
      return { savedToFile: false, savedToCache: true, fileName: "", needsPermission: false, count: normalized.length };
    }

    var perm = await queryHandlePermission(handle, "readwrite");
    if (perm !== "granted" && opts.allowPrompt) {
      perm = await requestHandlePermission(handle, "readwrite");
    }
    if (perm !== "granted") {
      log("Saved in local add-in cache (" + normalized.length + " materials). Click \"Reconnect local file\" to write directly to " + handle.name + ".");
      return { savedToFile: false, savedToCache: true, fileName: handle.name, needsPermission: true, count: normalized.length };
    }

    try {
      var writable = await handle.createWritable();
      await writable.write(saved.json);
      await writable.close();
      log("Saved " + normalized.length + " materials to local file \"" + handle.name + "\".");
      return { savedToFile: true, savedToCache: true, fileName: handle.name, needsPermission: false, count: normalized.length };
    } catch (e) {
      log("Saved in local cache, but writing to \"" + handle.name + "\" failed: " + (e.message || e));
      return { savedToFile: false, savedToCache: true, fileName: handle.name, needsPermission: false, error: e.message || String(e), count: normalized.length };
    }
  }

  /* Read the authoritative local master list (prefers linked disk file, falls back to computer cache) */
  async function readLocalMaster(opts) {
    opts = opts || {};
    var handle = await getLocalFileHandle();
    if (handle) {
      var perm = await queryHandlePermission(handle, "read");
      if (perm !== "granted" && opts.allowPrompt) {
        perm = await requestHandlePermission(handle, "readwrite");
      }
      if (perm === "granted") {
        try {
          var file = await handle.getFile();
          var text = await file.text();
          var parsed = parsePayload(text);
          await saveLocalCache(parsed.list, handle.name, parsed.lastChange);
          return { list: parsed.list, lastChange: parsed.lastChange, source: "file", fileName: handle.name, needsPermission: false };
        } catch (e) {
          // fall through to cache if file temporarily unreadable
        }
      } else {
        var cached = await readLocalCache();
        return {
          list: cached ? cached.list : [],
          lastChange: cached ? cached.lastChange : null,
          source: cached ? "cache" : "none",
          fileName: handle.name,
          needsPermission: true
        };
      }
    }
    var c = await readLocalCache();
    if (c && c.list) {
      return { list: c.list, lastChange: c.lastChange || null, source: "cache", fileName: (c.meta && c.meta.fileName) || "", needsPermission: false };
    }
    return { list: [], lastChange: null, source: "none", fileName: "", needsPermission: false };
  }

  /* Create/link a new local master-materials.json file on disk from the current workbook */
  async function linkNewLocalFile(log) {
    log = log || noop;
    if (!supportsLocalFileApi()) {
      throw new Error("This version of Excel does not allow direct file saving. Use Download JSON / Load JSON below instead.");
    }
    var handle = await window.showSaveFilePicker({
      suggestedName: "master-materials.json",
      types: [{ description: "BOM Master Materials JSON", accept: { "application/json": [".json"] } }]
    });
    await idbSet(IDB_HANDLE_KEY, handle);
    var wb = await workbookList();
    var res = await writeLocalMaster(wb, log, { allowPrompt: true });
    return { fileName: handle.name, count: res.count };
  }

  /* Link an existing master-materials.json file on disk and optionally sync it into the workbook */
  async function linkExistingLocalFile(log, opts) {
    log = log || noop;
    opts = opts || {};
    if (!supportsLocalFileApi()) {
      throw new Error("This version of Excel does not allow direct file linking. Use Load JSON below instead.");
    }
    var picks = await window.showOpenFilePicker({
      multiple: false,
      types: [{ description: "BOM Master Materials JSON", accept: { "application/json": [".json"] } }]
    });
    var handle = picks && picks[0];
    if (!handle) return null;
    await requestHandlePermission(handle, "readwrite");
    await idbSet(IDB_HANDLE_KEY, handle);

    var file = await handle.getFile();
    var text = await file.text();
    var parsed = parsePayload(text);
    await saveLocalCache(parsed.list, handle.name, parsed.lastChange);
    log("Linked local master file \"" + handle.name + "\" (" + parsed.list.length + " materials).");
    if (opts.applyToWorkbook !== false && parsed.list.length) {
      var r = await applyList(parsed.list, log, { syncRemovals: true, lastChange: parsed.lastChange });
      log("Workbook updated from \"" + handle.name + "\": " + r.added + " added, " + r.updated + " changed, " + r.removed + " removed.");
    }
    return { fileName: handle.name, count: parsed.list.length };
  }

  /* Save current workbook's Master Unique Materials into the local file & cache */
  async function saveWorkbookToLocal(log, opts) {
    log = log || noop;
    var wb = await workbookList();
    return writeLocalMaster(wb, log, opts);
  }

  /* Sync current workbook's Master Unique Materials from the local file (or local computer cache) */
  async function pullLocal(log, opts) {
    log = log || noop;
    opts = opts || {};
    var local = await readLocalMaster({ allowPrompt: !!opts.allowPrompt });
    if (!local.list.length) {
      if (local.needsPermission) {
        log("Local file \"" + local.fileName + "\" needs permission. Open Materials tab and click Reconnect local file.");
      }
      return { added: 0, updated: 0, removed: 0, skipped: [], source: local.source, fileName: local.fileName, empty: true };
    }
    var label = local.source === "file" ? ("local file \"" + local.fileName + "\"") : "local computer master list";
    log("Syncing master list from " + label + " (" + local.list.length + " materials) ...");
    var r = await applyList(local.list, log, { syncRemovals: opts.syncRemovals !== false, lastChange: local.lastChange });
    log("Master list synced from " + label + ": " + r.added + " added, " + r.updated + " changed" + (r.removed ? ", " + r.removed + " removed" : "") + ".");
    r.skipped.forEach(function (s) { log("Skipped " + s); });
    r.source = local.source;
    r.fileName = local.fileName;
    return r;
  }

  /* Fallback: import raw JSON text (e.g. from <input type="file">) into local computer cache & workbook */
  async function importLocalJsonText(text, fileName, log) {
    log = log || noop;
    var parsed = parsePayload(text);
    if (!parsed.list.length) throw new Error("No materials found in that JSON file.");
    await saveLocalCache(parsed.list, fileName || "master-materials.json", parsed.lastChange);
    var r = await applyList(parsed.list, log, { syncRemovals: true, lastChange: parsed.lastChange });
    log("Loaded \"" + (fileName || "master-materials.json") + "\" into this computer and workbook: " +
      r.added + " added, " + r.updated + " changed, " + r.removed + " removed.");
    return r;
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
    var local = await readLocalMaster({ allowPrompt: false });
    var d = diff(wb, local.list);
    d.localSource = local.source;
    d.localFileName = local.fileName;
    d.needsPermission = local.needsPermission;
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
                    savedToFile: !!(writeRes && writeRes.savedToFile),
                    fileName: (writeRes && writeRes.fileName) || "",
                    needsPermission: !!(writeRes && writeRes.needsPermission),
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
    // local computer master file API (no token required)
    supportsLocalFileApi: supportsLocalFileApi,
    getLocalFileStatus: getLocalFileStatus,
    requestLocalPermission: requestLocalPermission,
    unlinkLocalFile: unlinkLocalFile,
    linkNewLocalFile: linkNewLocalFile,
    linkExistingLocalFile: linkExistingLocalFile,
    readLocalMaster: readLocalMaster,
    writeLocalMaster: writeLocalMaster,
    saveWorkbookToLocal: saveWorkbookToLocal,
    pullLocal: pullLocal,
    compareLocal: compareLocal,
    importLocalJsonText: importLocalJsonText,
    // GitHub optional API
    hasToken: hasToken, getToken: getToken, setToken: setToken,
    compare: compare, pull: pull, push: push, importTemplate: importTemplate, repoInfo: repoInfo,
    // pure helpers
    diff: diff, changedFields: changedFields, serialize: serialize, parseList: parseList, parsePayload: parsePayload,
    buildLastChangeMeta: buildLastChangeMeta,
    workbookList: workbookList, applyList: applyList
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomSync = api;
})(typeof window !== "undefined" ? window : this);
