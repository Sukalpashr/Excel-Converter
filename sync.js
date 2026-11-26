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
  var PROJECTS_INDEX_KEY = "bomProjectsIndex";
  var IDB_PROJECTS_INDEX_KEY = "projectsIndex";
  var OPFS_FILE = "master-materials.json";
  var projectState = { loaded: false, name: "" };

  function slug(s) {
    var t = String(s || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "").slice(0, 60);
    return t || "project";
  }

  /* Returns the isolated storage slot for either Global Master (name="") or a named Project-Only list. */
  function targetFor(name) {
    var cleanName = String(name || "").trim();
    if (!cleanName) {
      return {
        project: false,
        name: "",
        slug: "",
        cache: LOCAL_CACHE_KEY,
        meta: LOCAL_META_KEY,
        idb: IDB_JSON_KEY,
        opfs: OPFS_FILE,
        offlineEndpoint: "./local-master.json"
      };
    }
    var k = slug(cleanName);
    return {
      project: true,
      name: cleanName,
      slug: k,
      cache: "bomProjectJson:" + k,
      meta: "bomProjectMeta:" + k,
      idb: "projectJson:" + k,
      opfs: "project-" + k + ".json",
      offlineEndpoint: "./local-project-" + k + ".json"
    };
  }

  function globalTarget() { return targetFor(""); }
  function getProject() { return { active: !!projectState.name, isProject: !!projectState.name, name: projectState.name, projectName: projectState.name }; }

  async function loadProjectState() {
    try {
      if (root.BomTools && root.BomTools.getMasterMode) {
        var mode = await root.BomTools.getMasterMode();
        projectState = { loaded: true, name: (mode && mode.isProject) ? (mode.projectName || mode.name || "") : "" };
      } else {
        projectState.loaded = true;
      }
    } catch (e) {
      projectState.loaded = true;
    }
    return getProject();
  }

  var FIELDS = ["category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"];
  var INPUT_KEYS = ["material"].concat(FIELDS);
  var AUDIT_KEYS = ["lastAction", "lastModified"];

  function noop() {}
  function runExcel(fn) {
    if (root.BomTools && typeof root.BomTools.runExcel === "function") return root.BomTools.runExcel(fn);
    try {
      if (typeof Office !== "undefined" && Office.context && Office.context.requirements &&
          Office.context.requirements.isSetSupported("ExcelApi", "1.9")) {
        return Excel.run({ delayForCellEdit: true }, fn);
      }
    } catch (e) { /* fall through */ }
    return Excel.run(fn);
  }
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

  function serialize(list, lastChange, extraMeta) {
    var nowIso = new Date().toISOString();
    var lcMeta = lastChange && lastChange.items ? buildLastChangeMeta(lastChange) : (lastChange && lastChange.time ? lastChange : null);
    var out = '{\n "version": 1,\n';
    var isProj = !!(extraMeta && (extraMeta.scope === "project" || extraMeta.type === "project"));
    if (isProj) {
      var pName = String(extraMeta.projectName || extraMeta.project || "Project Custom Materials");
      out += ' "type": "project",\n';
      out += ' "scope": "project",\n';
      out += ' "project": ' + JSON.stringify(pName) + ',\n';
      out += ' "projectName": ' + JSON.stringify(pName) + ',\n';
    }
    out += ' "updated": "' + nowIso + '",\n';
    if (lcMeta) {
      out += ' "lastChange": ' + JSON.stringify(lcMeta) + ',\n';
    }
    out += ' "materials": [\n' +
      list.map(function (m) { return "  " + JSON.stringify(clean(m)); }).join(",\n") + "\n ]\n}\n";
    return out;
  }

  function serializeProjectList(list, projectName, lastChange) {
    return serialize(list, lastChange, { type: "project", scope: "project", project: projectName, projectName: projectName });
  }

  function parsePayload(text) {
    var j = JSON.parse(text);
    var arr = Array.isArray(j) ? j : (j.materials || []);
    var list = arr.map(function (m) { return root.BomTools.normalizeMaterial(m); }).filter(function (m) { return m.material; });
    var isProj = !!(j && !Array.isArray(j) && (j.type === "project" || j.scope === "project"));
    var pName = (j && !Array.isArray(j) && (j.project || j.projectName)) ? String(j.project || j.projectName).trim() : "";
    return {
      list: list,
      isProject: isProj,
      type: isProj ? "project" : "master",
      scope: isProj ? "project" : "master",
      project: pName,
      projectName: pName,
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

  async function opfsDeleteFile(file) {
    if (!supportsOpfs() || !file) return false;
    try {
      var dir = await navigator.storage.getDirectory();
      if (typeof dir.removeEntry === "function") {
        await dir.removeEntry(file);
        return true;
      }
    } catch (e) { /* ignore */ }
    return false;
  }

  async function opfsListProjectSlugs() {
    var out = [];
    if (!supportsOpfs()) return out;
    try {
      var dir = await navigator.storage.getDirectory();
      if (typeof dir.entries === "function") {
        for await (var pair of dir.entries()) {
          var fname = pair && pair[0];
          var m = fname && /^project-([a-z0-9-]+)\.json$/i.exec(fname);
          if (m) out.push(m[1].toLowerCase());
        }
      }
    } catch (e) { /* ignore */ }
    return out;
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

  async function idbGetAllKeys() {
    var db = await openLocalDb();
    if (!db) return [];
    return new Promise(function (resolve) {
      try {
        var tx = db.transaction(IDB_STORE, "readonly");
        var store = tx.objectStore(IDB_STORE);
        if (typeof store.getAllKeys === "function") {
          var req = store.getAllKeys();
          req.onsuccess = function () { resolve(Array.isArray(req.result) ? req.result : []); };
          req.onerror = function () { resolve([]); };
          return;
        }
        var keys = [];
        var cur = store.openCursor();
        cur.onsuccess = function (ev) {
          var c = ev.target.result;
          if (c) { keys.push(c.key); c.continue(); } else resolve(keys);
        };
        cur.onerror = function () { resolve(keys); };
      } catch (e) { resolve([]); }
    });
  }

  function getCachedMeta(t) {
    t = t || globalTarget();
    try {
      var raw = localStorage.getItem(t.meta);
      return raw ? JSON.parse(raw) : null;
    } catch (e) { return null; }
  }

  function isOfflineLts() {
    return !!(root && root.BOM_OFFLINE_LTS);
  }

  async function offlineServerReadJson(t) {
    t = t || globalTarget();
    if (!isOfflineLts() || typeof fetch !== "function" || !t.offlineEndpoint) return null;
    try {
      var res = await fetch(t.offlineEndpoint + "?t=" + Date.now(), { cache: "no-store" });
      if (!res.ok) return null;
      return await res.text();
    } catch (e) { return null; }
  }

  async function offlineServerWriteJson(json, t) {
    t = t || globalTarget();
    if (!isOfflineLts() || typeof fetch !== "function" || !t.offlineEndpoint) return false;
    try {
      var res = await fetch(t.offlineEndpoint, {
        method: "POST",
        headers: { "Content-Type": "application/json; charset=utf-8" },
        body: json
      });
      return !!res.ok;
    } catch (e) { return false; }
  }

  async function offlineServerDeleteProject(t) {
    if (!isOfflineLts() || typeof fetch !== "function" || !t || !t.project || !t.offlineEndpoint) return false;
    try {
      var res = await fetch(t.offlineEndpoint, { method: "DELETE" });
      return !!res.ok;
    } catch (e) { return false; }
  }

  async function offlineServerListProjects() {
    if (!isOfflineLts() || typeof fetch !== "function") return [];
    try {
      var res = await fetch("./local-projects.json?t=" + Date.now(), { cache: "no-store" });
      if (!res.ok) return [];
      var j = await res.json();
      return (j && Array.isArray(j.projects)) ? j.projects : [];
    } catch (e) { return []; }
  }

  async function readProjectIndex() {
    var map = {};
    try {
      var lsRaw = localStorage.getItem(PROJECTS_INDEX_KEY);
      if (lsRaw) {
        var parsedLs = JSON.parse(lsRaw);
        if (parsedLs && typeof parsedLs === "object") {
          Object.keys(parsedLs).forEach(function (k) { map[k] = parsedLs[k]; });
        }
      }
    } catch (e) { /* ignore */ }
    try {
      var idbRaw = await idbGet(IDB_PROJECTS_INDEX_KEY);
      var parsedIdb = typeof idbRaw === "string" ? JSON.parse(idbRaw) : idbRaw;
      if (parsedIdb && typeof parsedIdb === "object") {
        Object.keys(parsedIdb).forEach(function (k) {
          var prev = map[k];
          var cur = parsedIdb[k];
          if (!prev || String((cur && cur.updated) || "") > String((prev && prev.updated) || "")) {
            map[k] = cur;
          }
        });
      }
    } catch (e) { /* ignore */ }
    return map;
  }

  async function writeProjectIndex(map) {
    var json = JSON.stringify(map || {});
    try { localStorage.setItem(PROJECTS_INDEX_KEY, json); } catch (e) { /* ignore */ }
    try { await idbSet(IDB_PROJECTS_INDEX_KEY, json); } catch (e) { /* ignore */ }
  }

  async function upsertProjectIndex(t, count, updated, lastChange) {
    if (!t || !t.project || !t.slug) return;
    var map = await readProjectIndex();
    map[t.slug] = {
      slug: t.slug,
      name: t.name || t.slug,
      count: count || 0,
      updated: updated || new Date().toISOString(),
      lastChange: lastChange || null
    };
    await writeProjectIndex(map);
  }

  async function removeFromProjectIndex(slugKey) {
    if (!slugKey) return;
    var map = await readProjectIndex();
    if (map[slugKey]) {
      delete map[slugKey];
      await writeProjectIndex(map);
    }
  }

  function targetBySlug(slugStr, displayName) {
    var k = slug(slugStr || displayName || "");
    return {
      project: true,
      name: String(displayName || slugStr || k).trim() || k,
      slug: k,
      cache: "bomProjectJson:" + k,
      meta: "bomProjectMeta:" + k,
      idb: "projectJson:" + k,
      opfs: "project-" + k + ".json",
      offlineEndpoint: "./local-project-" + k + ".json"
    };
  }

  async function saveLocalCache(list, lastChange, t) {
    t = t || globalTarget();
    requestPersistentStorage();
    var existingMeta = getCachedMeta(t);
    var lcMeta = lastChange ? (buildLastChangeMeta(lastChange) || lastChange) : (existingMeta && existingMeta.lastChange) || null;
    var json = t.project
      ? serializeProjectList(list, t.name, lcMeta)
      : serialize(list, lcMeta);
    var meta = {
      name: t.name || "",
      slug: t.slug || "",
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
      opfsWriteJson(json, t.opfs),
      offlineServerWriteJson(json, t),
      t.project ? upsertProjectIndex(t, list.length, meta.updated, lcMeta) : Promise.resolve()
    ]);
    // Clean up any legacy external file handle from earlier versions so Excel never asks to reconnect
    if (!t.project) {
      try { await idbDel(IDB_HANDLE_KEY); } catch (e) { /* ignore */ }
    }
    return { json: json, meta: meta };
  }

  async function readLocalCache(t) {
    t = t || globalTarget();
    // In Offline LTS mode, check durable PC disk file first (master-materials.json or projects/project-<slug>.json)
    var json = await offlineServerReadJson(t);
    if (!json) json = await opfsReadJson(t.opfs);
    if (!json) json = await idbGet(t.idb);
    if (!json) {
      try { json = localStorage.getItem(t.cache); } catch (e) { json = null; }
    }
    if (!json) return null;
    try {
      var parsed = parsePayload(json);
      return {
        list: parsed.list,
        projectName: parsed.projectName || t.name || "",
        slug: t.slug || "",
        updated: parsed.updated,
        lastChange: parsed.lastChange,
        meta: getCachedMeta(t)
      };
    } catch (e) { return null; }
  }

  async function saveProjectSlot(projectName, list, lastChange) {
    var cleanName = String(projectName || "").trim();
    if (!cleanName) throw new Error("Project list name is required.");
    var normalized = (list || []).map(function (m) { return root.BomTools.normalizeMaterial(m); }).filter(function (m) { return m.material; });
    var t = targetFor(cleanName);
    var saved = await saveLocalCache(normalized, lastChange || null, t);
    projectState = { loaded: true, name: cleanName };
    return { savedToProjectSlot: true, projectName: cleanName, slug: t.slug, count: normalized.length, json: saved.json };
  }

  async function readProjectSlotBySlug(slugOrName) {
    var raw = String(slugOrName || "").trim();
    if (!raw) return { name: "", slug: "", list: [], lastChange: null, updated: "", source: "none" };
    var t = targetBySlug(raw, raw);
    var c = await readLocalCache(t);
    if (c && c.list && c.list.length) {
      var resolvedName = c.projectName || (c.meta && c.meta.name) || raw;
      return {
        name: resolvedName,
        projectName: resolvedName,
        slug: t.slug,
        list: c.list,
        lastChange: c.lastChange || null,
        updated: c.updated || (c.meta && c.meta.updated) || "",
        source: "project"
      };
    }
    return { name: raw, projectName: raw, slug: t.slug, list: [], lastChange: null, updated: "", source: "none" };
  }

  async function readProjectSlot(projectName) {
    return readProjectSlotBySlug(projectName);
  }

  /* Discovers and returns every Project-Only material list saved on this PC (across Offline LTS disk, OPFS, IDB, localStorage & index), newest first. */
  async function listProjectSlots() {
    var bySlug = {};
    var mergeEntry = function (sKey, entry) {
      if (!sKey) return;
      var cur = bySlug[sKey];
      var nextUpdated = String((entry && entry.updated) || "");
      var curUpdated = String((cur && cur.updated) || "");
      if (!cur) {
        bySlug[sKey] = {
          slug: sKey,
          name: (entry && entry.name) || "",
          count: (entry && typeof entry.count === "number") ? entry.count : -1,
          updated: nextUpdated,
          lastChange: (entry && entry.lastChange) || null
        };
      } else {
        if (nextUpdated && (!curUpdated || nextUpdated > curUpdated)) {
          cur.updated = nextUpdated;
          if (entry.name) cur.name = entry.name;
          if (typeof entry.count === "number" && entry.count >= 0) cur.count = entry.count;
          if (entry.lastChange) cur.lastChange = entry.lastChange;
        } else {
          if (!cur.name && entry && entry.name) cur.name = entry.name;
          if (cur.count < 0 && entry && typeof entry.count === "number") cur.count = entry.count;
        }
      }
    };

    // 1. Offline LTS disk folder (C:\ProgramData\BOMTools\data\projects\)
    var diskItems = await offlineServerListProjects();
    diskItems.forEach(function (it) {
      if (it && it.slug) mergeEntry(String(it.slug).toLowerCase(), it);
    });

    // 2. Stored index
    var idxMap = await readProjectIndex();
    Object.keys(idxMap).forEach(function (k) { mergeEntry(k, idxMap[k]); });

    // 3. Scan localStorage keys (discovers any project saved before v2.2!)
    try {
      for (var i = 0; i < localStorage.length; i++) {
        var lk = localStorage.key(i);
        if (lk && lk.indexOf("bomProjectJson:") === 0) {
          var s1 = lk.slice("bomProjectJson:".length);
          if (s1) mergeEntry(s1, { slug: s1 });
        }
      }
    } catch (e) { /* ignore */ }

    // 4. Scan IndexedDB keys
    var idbKeys = await idbGetAllKeys();
    idbKeys.forEach(function (k) {
      if (typeof k === "string" && k.indexOf("projectJson:") === 0) {
        var s2 = k.slice("projectJson:".length);
        if (s2) mergeEntry(s2, { slug: s2 });
      }
    });

    // 5. Scan OPFS directory
    var opfsSlugs = await opfsListProjectSlugs();
    opfsSlugs.forEach(function (s3) { mergeEntry(s3, { slug: s3 }); });

    // Hydrate any entries missing name/count (and verify data actually exists)
    var slugs = Object.keys(bySlug);
    var results = [];
    for (var j = 0; j < slugs.length; j++) {
      var sk = slugs[j];
      var ent = bySlug[sk];
      if (!ent.name || ent.count < 0) {
        var slot = await readProjectSlotBySlug(sk);
        if (slot && slot.list && slot.list.length) {
          ent.name = slot.name || sk;
          ent.count = slot.list.length;
          if (!ent.updated && slot.updated) ent.updated = slot.updated;
          if (!ent.lastChange && slot.lastChange) ent.lastChange = slot.lastChange;
        }
      }
      if (ent.count > 0) {
        results.push({
          slug: ent.slug,
          name: ent.name || ent.slug,
          count: ent.count,
          updated: ent.updated || "",
          lastChange: ent.lastChange || null
        });
      }
    }

    results.sort(function (a, b) {
      var cmp = String(b.updated || "").localeCompare(String(a.updated || ""));
      return cmp !== 0 ? cmp : String(a.name || "").localeCompare(String(b.name || ""));
    });
    return results;
  }

  async function deleteProjectSlot(slugOrName) {
    var raw = String(slugOrName || "").trim();
    if (!raw) return false;
    var t = targetBySlug(raw, raw);
    try { localStorage.removeItem(t.cache); } catch (e) { /* ignore */ }
    try { localStorage.removeItem(t.meta); } catch (e) { /* ignore */ }
    await Promise.all([
      idbDel(t.idb),
      opfsDeleteFile(t.opfs),
      offlineServerDeleteProject(t),
      removeFromProjectIndex(t.slug)
    ]);
    return true;
  }

  /* Loads a saved Project-Only list from this PC into the current workbook (or links it ready for Restore on converted workbooks). */
  async function applySavedProjectSlot(slugOrName, log) {
    log = log || noop;
    var slot = await readProjectSlotBySlug(slugOrName);
    if (!slot || !slot.list || !slot.list.length) {
      throw new Error("That saved project list could not be found or is empty.");
    }
    var pName = slot.name || String(slugOrName);
    suppressWatch(6000);
    var applied = true;
    var r = { added: 0, updated: 0, removed: 0, skipped: [] };
    try {
      r = await applyList(slot.list, log, { syncRemovals: true, lastChange: slot.lastChange });
    } catch (e) {
      if (!/not found/i.test(String(e && e.message))) throw e;
      applied = false;
    }
    if (root.BomTools && root.BomTools.setProjectMaterialMode) {
      await root.BomTools.setProjectMaterialMode(pName);
    }
    projectState = { loaded: true, name: pName };
    if (applied) {
      try {
        var curList = await workbookList();
        updateWatchSnap(curList);
      } catch (e2) { /* ignore */ }
      log("Loaded saved Project-Only list \"" + pName + "\" (" + slot.list.length + " materials) into this workbook. Global Master was not changed.");
    } else {
      log("Selected saved Project-Only list \"" + pName + "\" (" + slot.list.length + " materials) for this converted workbook. Open Convert > Restore Automated to rebuild sheets with it.");
    }
    r.projectName = pName;
    r.slug = slot.slug;
    r.count = slot.list.length;
    r.applied = applied;
    return r;
  }

  async function getLocalFileStatus() {
    var cache = await readLocalCache(globalTarget());
    var count = cache && cache.list ? cache.list.length : 0;
    var proj = await loadProjectState();
    var pCount = 0;
    var pUpdated = "";
    if (proj.active) {
      var pc = await readLocalCache(targetFor(proj.name));
      if (pc && pc.list) {
        pCount = pc.list.length;
        pUpdated = pc.updated || (pc.meta && pc.meta.updated) || "";
      }
    }
    return {
      ready: true,
      hasCache: count > 0,
      cachedCount: count,
      cachedUpdated: (cache && (cache.updated || (cache.meta && cache.meta.updated))) || "",
      lastChange: (cache && cache.lastChange) || (cache && cache.meta && cache.meta.lastChange) || null,
      opfsSupported: supportsOpfs(),
      project: proj.name,
      hasProjectCache: pCount > 0,
      projectCount: pCount,
      projectUpdated: pUpdated
    };
  }

  /* Saves automatically to Built-in Add-in Storage (or to the isolated project slot when opts.projectName / project mode is active). */
  async function writeLocalMaster(list, log, opts) {
    log = log || noop;
    opts = opts || {};
    var normalized = list.map(function (m) { return root.BomTools.normalizeMaterial(m); }).filter(function (m) { return m.material; });
    var t = opts.forceGlobal
      ? globalTarget()
      : (opts.projectName ? targetFor(opts.projectName) : targetFor(projectState.name));
    var saved = await saveLocalCache(normalized, opts.lastChange || null, t);
    return {
      savedToBuiltIn: !t.project,
      savedToProjectSlot: !!t.project,
      projectName: t.name,
      count: normalized.length,
      json: saved.json
    };
  }

  /* Reads the authoritative Global Master list from Built-in Add-in Storage (auto-connected, never prompts) */
  async function readLocalMaster() {
    var c = await readLocalCache(globalTarget());
    if (c && c.list && c.list.length) {
      return { list: c.list, lastChange: c.lastChange || null, source: "builtin" };
    }
    return { list: [], lastChange: null, source: "none" };
  }

  /* Seeds Built-in Storage once from the open workbook if storage is currently empty (isolated per project when in Project-Only mode) */
  async function ensureSeededFromWorkbook() {
    var proj = await loadProjectState();
    if (proj.active) {
      var pc = await readLocalCache(targetFor(proj.name));
      if (pc && pc.list && pc.list.length) return { seeded: false, count: pc.list.length, projectMode: true };
      try {
        var pwb = await workbookList();
        if (pwb.length) {
          await saveProjectSlot(proj.name, pwb, null);
          return { seeded: true, count: pwb.length, projectMode: true };
        }
      } catch (e) { /* converted workbook has no Master sheet yet */ }
      return { seeded: false, count: 0, projectMode: true };
    }
    var c = await readLocalCache(globalTarget());
    if (c && c.list && c.list.length) return { seeded: false, count: c.list.length };
    try {
      var wb = await workbookList();
      if (wb.length) {
        await writeLocalMaster(wb, noop, { forceGlobal: true });
        return { seeded: true, count: wb.length };
      }
    } catch (e) { /* workbook may not have Master sheet yet */ }
    return { seeded: false, count: 0 };
  }

  /* Save current workbook's Master Unique Materials into Built-in Storage */
  async function saveWorkbookToLocal(log, opts) {
    log = log || noop;
    opts = opts || {};
    var proj = await loadProjectState();
    if (proj.active && !opts.forceGlobal) {
      var pwb = await workbookList();
      var pres = await saveProjectSlot(proj.name, pwb, opts.lastChange || null);
      log("Saved " + pres.count + " materials to this project's isolated storage (\"" + proj.name + "\").");
      return pres;
    }
    var wb = await workbookList();
    var res = await writeLocalMaster(wb, log, { forceGlobal: true, lastChange: opts.lastChange || null });
    log("Saved " + res.count + " materials to this computer's Global Master storage.");
    return res;
  }

  /* Sync current workbook's Master Unique Materials from Built-in Storage.
   * When workbook is in Project-Only mode (unless opts.force / opts.forceGlobal), pulls ONLY from that project's isolated slot!
   */
  async function pullLocal(log, opts) {
    log = log || noop;
    opts = opts || {};
    var forceGlobal = !!(opts.force || opts.forceGlobal);
    var proj = forceGlobal ? { active: false, name: "" } : await loadProjectState();
    if (proj.active) {
      var pSlot = await readProjectSlot(proj.name);
      if (!pSlot.list.length) {
        log("Project-Only list \"" + proj.name + "\" has no saved copy on this computer yet (Global Master was not applied).");
        return { added: 0, updated: 0, removed: 0, skipped: [], source: "project", empty: true, projectMode: true, projectName: proj.name };
      }
      log("Syncing Project-Only list \"" + proj.name + "\" from this computer (" + pSlot.list.length + " materials) ...");
      var pr = await applyList(pSlot.list, log, { syncRemovals: opts.syncRemovals !== false, lastChange: pSlot.lastChange });
      if (root.BomTools && root.BomTools.setProjectMaterialMode) {
        await root.BomTools.setProjectMaterialMode(proj.name);
      }
      log("Project list synced: " + pr.added + " added, " + pr.updated + " changed" + (pr.removed ? ", " + pr.removed + " removed" : "") + ".");
      pr.skipped.forEach(function (s) { log("Skipped " + s); });
      pr.source = "project";
      pr.projectMode = true;
      pr.projectName = proj.name;
      return pr;
    }

    var local = await readLocalMaster();
    if (!local.list.length) {
      return { added: 0, updated: 0, removed: 0, skipped: [], source: "none", empty: true };
    }
    log("Syncing Global Master list from built-in computer storage (" + local.list.length + " materials) ...");
    var r = await applyList(local.list, log, { syncRemovals: opts.syncRemovals !== false, lastChange: local.lastChange });
    log("Master list synced: " + r.added + " added, " + r.updated + " changed" + (r.removed ? ", " + r.removed + " removed" : "") + ".");
    r.skipped.forEach(function (s) { log("Skipped " + s); });
    r.source = "builtin";
    return r;
  }

  /* Build JSON text for Global Master backup/transfer export */
  async function buildExportJson() {
    var cache = await readLocalCache(globalTarget());
    var lastChange = (cache && cache.lastChange) || (cache && cache.meta && cache.meta.lastChange) || null;
    var list = [];
    var proj = await loadProjectState();
    try {
      if (!proj.active) list = await workbookList();
    } catch (e) {
      list = (cache && cache.list) || [];
    }
    if (!list.length && cache && cache.list) list = cache.list;
    return { json: serialize(list, lastChange), count: list.length, fileName: "master-materials.json" };
  }

  function makeProjectFileSlug(name) {
    return slug(name) + "-materials.json";
  }

  /* Build separate reusable JSON file for the current workbook's Project-Only material list.
   * Also syncs the live list to the per-project storage slot (and falls back to the slot if the sheet was already flattened). */
  async function buildProjectExportJson(projectName, lastChange) {
    var proj = await loadProjectState();
    var pName = String(projectName || proj.name || "").trim() || "Project Custom Materials";
    var cache = await readLocalCache(targetFor(pName));
    var list = [];
    try {
      list = await workbookList();
    } catch (e) {
      list = [];
    }
    var useLastChange = lastChange || (cache && cache.lastChange) || null;
    if (list.length) {
      await saveLocalCache(list, useLastChange, targetFor(pName));
    } else if (cache && cache.list && cache.list.length) {
      list = cache.list;
    }
    return {
      json: serializeProjectList(list, pName, useLastChange),
      count: list.length,
      projectName: pName,
      suggestedName: makeProjectFileSlug(pName)
    };
  }

  /* Import a separate Project Materials JSON into THIS workbook ONLY (switches to Project-Only mode; never modifies Global Master).
   * Works even on an already-converted workbook (no Master sheet yet): saves to the per-project slot + sets BomProjectList so Restore uses it automatically. */
  async function importProjectJsonText(text, fileName, log) {
    log = log || noop;
    var parsed = parsePayload(text);
    if (!parsed.list.length) throw new Error("No materials found in that JSON file.");
    var fallbackName = String(fileName || "Project Materials")
      .replace(/\.json$/i, "")
      .replace(/^project-materials-?/i, "")
      .replace(/[-_]+materials$/i, "")
      .replace(/[-_]+/g, " ")
      .trim() || "Project Custom Materials";
    var pName = parsed.projectName || fallbackName;

    suppressWatch(6000);
    await saveProjectSlot(pName, parsed.list, parsed.lastChange);

    var applied = true;
    var r = { added: 0, updated: 0, removed: 0, skipped: [] };
    try {
      r = await applyList(parsed.list, log, { syncRemovals: true, lastChange: parsed.lastChange });
    } catch (e) {
      if (!/not found/i.test(String(e && e.message))) throw e;
      applied = false; // Converted workbook has no Master sheet yet — saved ready for Restore
    }

    if (root.BomTools && root.BomTools.setProjectMaterialMode) {
      await root.BomTools.setProjectMaterialMode(pName);
    }
    if (applied) {
      try {
        var curList = await workbookList();
        updateWatchSnap(curList);
      } catch (e) { /* ignore */ }
      log("Loaded Project-Only list \"" + pName + "\" (" + parsed.list.length + " materials) into this workbook only. Global Master was not changed.");
    } else {
      log("Saved Project-Only list \"" + pName + "\" (" + parsed.list.length + " materials) for this converted workbook. Click Restore Automated to rebuild sheets with these materials.");
    }
    r.projectName = pName;
    r.count = parsed.list.length;
    r.applied = applied;
    return r;
  }

  /* Import a Global Master JSON backup/transfer file into Global Master Built-in Storage.
   * Guardrail: refuses a Project-Only JSON file so it can never accidentally overwrite Global Master! */
  async function importLocalJsonText(text, fileName, log) {
    log = log || noop;
    var parsed = parsePayload(text);
    if (!parsed.list.length) throw new Error("No materials found in that JSON file.");
    if (parsed.isProject) {
      throw new Error("That is a Project-Only materials file (\"" + (parsed.projectName || fileName) + "\"). Use \"Load .json...\" in the Project-Only section at the top of Materials so your Global Master list stays untouched.");
    }
    await saveLocalCache(parsed.list, parsed.lastChange, globalTarget());
    try {
      var proj = await loadProjectState();
      if (proj.active) {
        log("Saved " + parsed.list.length + " materials to your computer's Global Master storage (this workbook stayed on its Project-Only list \"" + proj.name + "\").");
        return { added: 0, updated: 0, removed: 0, skipped: [], savedOnly: true };
      }
      var r = await applyList(parsed.list, log, { syncRemovals: true, lastChange: parsed.lastChange });
      log("Imported \"" + (fileName || "master-materials.json") + "\" (" + parsed.list.length + " materials): " +
        r.added + " added, " + r.updated + " changed, " + r.removed + " removed.");
      return r;
    } catch (e) {
      log("Saved " + parsed.list.length + " materials from \"" + (fileName || "master-materials.json") + "\" to this computer's Global Master storage. They will load automatically into your next BOM.");
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
    await runExcel(async function (context) {
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
    var proj = await loadProjectState();
    if (proj.active) { log("Project-Only list active (\"" + proj.name + "\") — shared GitHub list not applied."); return { added: 0, updated: 0, removed: 0, skipped: [] }; }
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
    var proj = await loadProjectState();
    if (proj.active) { log("Project-Only list (\"" + proj.name + "\"): not published to shared GitHub master."); return { added: 0, changed: 0 }; }
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

    await runExcel(async function (context) {
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
      await runExcel(async function (context) {
        var ws = context.workbook.worksheets.getItemOrNullObject(T.MASTER);
        await context.sync();
        if (ws.isNullObject) return;
        watcher.snap = await workbookList();
        ws.onChanged.add(function (ev) {
          if (Date.now() < watcher.ignoreUntil) return;
          if (ev && ev.source && String(ev.source) !== "Local") return;
          if (ev && ev.changeType && ev.changeType !== "RangeEdited" && ev.changeType !== "RowInserted" && ev.changeType !== "RowDeleted") return;
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

                // Check Project-Only isolation: save to dedicated project slot (NEVER overwrite Global Master or GitHub!)
                var proj = await loadProjectState();
                var isProj = !!proj.active;

                if (isProj) {
                  await saveProjectSlot(proj.name, cur, lastChange);
                } else {
                  await writeLocalMaster(cur, log, { forceGlobal: true, lastChange: lastChange });
                  if (hasToken()) {
                    var names = d.onlyWb.map(function (m) { return m.material; }).concat(d.changed.map(function (c) { return c.material; }));
                    if (names.length) {
                      try { await push(log, { only: names }); } catch (e) { log("GitHub auto-publish skipped: " + (e.message || e)); }
                    }
                  }
                }

                watcher.snap = cur;
                if (typeof watcher.onSaved === "function") {
                  watcher.onSaved({
                    source: "sheet",
                    time: nowIso,
                    savedToBuiltIn: !isProj,
                    projectOnly: isProj,
                    projectName: isProj ? proj.name : "",
                    items: items
                  });
                }
              }
            } catch (e) { log("Auto-save failed: " + (e.message || e)); }
            watcher.busy = false;
          }, 1800);
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
    buildProjectExportJson: buildProjectExportJson,
    importLocalJsonText: importLocalJsonText,
    importProjectJsonText: importProjectJsonText,
    // project-only isolated storage slots
    getProject: getProject,
    loadProjectState: loadProjectState,
    saveProjectSlot: saveProjectSlot,
    readProjectSlot: readProjectSlot,
    readProjectSlotBySlug: readProjectSlotBySlug,
    listProjectSlots: listProjectSlots,
    deleteProjectSlot: deleteProjectSlot,
    applySavedProjectSlot: applySavedProjectSlot,
    targetFor: targetFor,
    slug: slug,
    // GitHub optional API
    hasToken: hasToken, getToken: getToken, setToken: setToken,
    compare: compare, pull: pull, push: push, importTemplate: importTemplate, repoInfo: repoInfo,
    // pure helpers
    diff: diff, changedFields: changedFields, serialize: serialize, serializeProjectList: serializeProjectList,
    parseList: parseList, parsePayload: parsePayload,
    buildLastChangeMeta: buildLastChangeMeta,
    workbookList: workbookList, applyList: applyList
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomSync = api;
})(typeof window !== "undefined" ? window : this);
