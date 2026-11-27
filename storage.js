/* storage.js - Built-in Storage, Full v1.8-v3.0 Migration, Project-Only Lists & Saved Library Pool (3.0v patch2) */
(() => {
  "use strict";

  const VERSION = "3.0v patch2";
  const DB_NAME = "BomToolsV3DB";
  const DB_STORE = "kv";
  const LEGACY_DB_NAME = "BomToolsLocalDB";
  const LEGACY_DB_STORE = "kv";
  const LEGACY_IDB_GLOBAL = "localMasterJson";
  const LEGACY_LS_GLOBAL = "bomLocalMasterJson";
  const LEGACY_OPFS_GLOBAL = "master-materials.json";

  const LS_GLOBAL = "bomTools_globalMaster_v3";
  const LS_PROJECTS = "bomTools_projectLists_v3";
  const LS_TOKEN = "bomTools_ghToken_v3";
  const LEGACY_LS_TOKEN = "bomGhToken";
  const OPFS_GLOBAL = "bom-global-master.json";
  const OPFS_PROJECTS = "bom-project-lists.json";

  const UNITS = ["kg/m", "kg/m²", "kg/each"];
  const COMPARE_FIELDS = ["category", "unit", "lin", "area", "piece", "sl", "sw", "remarks"];

  const lc = (s) => String(s ?? "").trim().toLowerCase();
  const isNil = (v) => v === null || v === undefined || (typeof v === "number" && Number.isNaN(v));
  const toNum = (v) => {
    if (v === null || v === undefined || v === "") return null;
    const n = Number(String(v).replace(/,/g, ""));
    return Number.isFinite(n) ? n : null;
  };
  const slugify = (name) =>
    String(name || "project").trim().toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-+|-+$/g, "") || "project";

  function normAction(v) {
    const s = lc(v);
    if (s === "add" || s === "added") return "added";
    if (s === "edit" || s === "edited" || s === "mod" || s === "modified" || s === "updated") return "modified";
    if (s === "remove" || s === "removed" || s === "deleted") return "removed";
    return "";
  }

  function normalizeMaterial(raw) {
    const unit = String(raw?.unit ?? "").trim();
    return {
      material: String(raw?.material ?? "").trim(),
      category: String(raw?.category ?? "").trim(),
      unit,
      lin: unit === "kg/m" ? toNum(raw?.lin) : null,
      area: unit === "kg/m²" ? toNum(raw?.area) : null,
      piece: unit === "kg/each" ? toNum(raw?.piece) : null,
      sl: unit === "kg/m²" ? toNum(raw?.sl) : null,
      sw: unit === "kg/m²" ? toNum(raw?.sw) : null,
      remarks: String(raw?.remarks ?? "").trim(),
      lastAction: normAction(raw?.lastAction),
      lastModified: raw?.lastModified ? String(raw.lastModified).trim() : ""
    };
  }

  function cleanForJson(m) {
    const n = normalizeMaterial(m);
    const o = {
      material: n.material,
      category: n.category,
      unit: n.unit,
      lin: n.lin,
      area: n.area,
      piece: n.piece,
      sl: n.sl,
      sw: n.sw,
      remarks: n.remarks
    };
    if (n.lastAction) o.lastAction = n.lastAction;
    if (n.lastModified) o.lastModified = n.lastModified;
    return o;
  }

  function sameVal(a, b) {
    const x = isNil(a) || a === "" ? null : a;
    const y = isNil(b) || b === "" ? null : b;
    if (x === null || y === null) return x === y;
    if (typeof x === "number" && typeof y === "number") return Math.abs(x - y) < 1e-9;
    return String(x).trim() === String(y).trim();
  }

  function changedFields(a, b) {
    return COMPARE_FIELDS.filter((k) => !sameVal(a[k], b[k]));
  }

  function diffLists(curr, prev) {
    const mapA = new Map();
    const mapB = new Map();
    for (const m of curr) if (m.material && !mapA.has(lc(m.material))) mapA.set(lc(m.material), m);
    for (const m of prev) if (m.material && !mapB.has(lc(m.material))) mapB.set(lc(m.material), m);
    const added = [];
    const modified = [];
    const removed = [];
    for (const [k, m] of mapA) {
      if (!mapB.has(k)) {
        added.push(m);
      } else {
        const fields = changedFields(mapB.get(k), m);
        if (fields.length) modified.push({ material: m.material, fields, next: m, prev: mapB.get(k) });
      }
    }
    for (const [k, m] of mapB) if (!mapA.has(k)) removed.push(m);
    return { added, modified, removed };
  }

  function buildLastChange(source, items, timeIso) {
    if (!items || !items.length) return null;
    const time = timeIso || new Date().toISOString();
    const cleanItems = items
      .map((it) => ({
        material: typeof it.material === "string" ? it.material : it.material?.material || "",
        type: normAction(it.type) || "modified",
        time: it.time || time,
        ...(it.fields && it.fields.length ? { fields: [...it.fields] } : {})
      }))
      .filter((x) => x.material);
    if (!cleanItems.length) return null;
    const kinds = new Set(cleanItems.map((x) => x.type));
    return {
      time,
      type: kinds.size === 1 ? [...kinds][0] : "mixed",
      source: source || "panel",
      items: cleanItems
    };
  }

  function serializeMaster(list, lastChange, updatedIso) {
    const payload = {
      version: 3,
      type: "master",
      updated: updatedIso || new Date().toISOString(),
      ...(lastChange ? { lastChange } : {}),
      materials: list.map(cleanForJson)
    };
    return JSON.stringify(payload, null, 2) + "\n";
  }

  function serializeProject(projectName, list, lastChange, id, updatedIso) {
    const cleanName = String(projectName || "Project").trim();
    const payload = {
      version: 3,
      type: "project",
      id: id || slugify(cleanName),
      project: cleanName,
      projectName: cleanName,
      updated: updatedIso || new Date().toISOString(),
      ...(lastChange ? { lastChange } : {}),
      materials: list.map(cleanForJson)
    };
    return JSON.stringify(payload, null, 2) + "\n";
  }

  function parseJsonText(text) {
    const raw = JSON.parse(text);
    const isObj = raw && typeof raw === "object" && !Array.isArray(raw);
    const arr = Array.isArray(raw) ? raw : Array.isArray(raw?.materials) ? raw.materials : [];
    const materials = arr.map(normalizeMaterial).filter((m) => m.material);
    return {
      type: isObj && raw.type === "project" ? "project" : "master",
      id: isObj ? String(raw.id || "") : "",
      projectName: isObj ? String(raw.projectName || raw.project || raw.name || "").trim() : "",
      updated: isObj ? String(raw.updated || "") : "",
      lastChange: isObj ? raw.lastChange || null : null,
      materials
    };
  }

  /* ---------------- Multi-Layer Storage (OPFS + IndexedDB v3 & legacy v1.8-v2.6 + localStorage + Offline LTS) ---------------- */
  let offlineDetected = null;

  function isOfflineHost() {
    try {
      const h = location.hostname;
      return h === "localhost" || h === "127.0.0.1" || location.search.includes("offline=1");
    } catch {
      return false;
    }
  }

  async function detectOfflineApi() {
    if (offlineDetected !== null) return offlineDetected;
    if (!isOfflineHost()) {
      offlineDetected = false;
      return false;
    }
    try {
      const r = await fetch("/api/health", { cache: "no-store" });
      offlineDetected = r.ok;
    } catch {
      offlineDetected = false;
    }
    return offlineDetected;
  }

  function openNamedDb(dbName, storeName) {
    return new Promise((resolve) => {
      if (typeof indexedDB === "undefined") return resolve(null);
      try {
        const req = indexedDB.open(dbName, 1);
        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains(storeName)) db.createObjectStore(storeName);
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });
  }

  async function idbGetFrom(dbName, storeName, key) {
    const db = await openNamedDb(dbName, storeName);
    if (!db) return null;
    return new Promise((resolve) => {
      try {
        if (!db.objectStoreNames.contains(storeName)) return resolve(null);
        const req = db.transaction(storeName, "readonly").objectStore(storeName).get(key);
        req.onsuccess = () => resolve(req.result ?? null);
        req.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });
  }

  async function idbSetIn(dbName, storeName, key, val) {
    const db = await openNamedDb(dbName, storeName);
    if (!db) return;
    await new Promise((resolve) => {
      try {
        if (!db.objectStoreNames.contains(storeName)) return resolve();
        const tx = db.transaction(storeName, "readwrite");
        tx.objectStore(storeName).put(val, key);
        tx.oncomplete = () => resolve();
        tx.onerror = () => resolve();
      } catch {
        resolve();
      }
    });
  }

  async function idbDelIn(dbName, storeName, key) {
    const db = await openNamedDb(dbName, storeName);
    if (!db) return;
    await new Promise((resolve) => {
      try {
        if (!db.objectStoreNames.contains(storeName)) return resolve();
        const tx = db.transaction(storeName, "readwrite");
        tx.objectStore(storeName).delete(key);
        tx.oncomplete = () => resolve();
        tx.onerror = () => resolve();
      } catch {
        resolve();
      }
    });
  }

  async function idbEntriesFrom(dbName, storeName, prefix) {
    const db = await openNamedDb(dbName, storeName);
    if (!db) return [];
    return new Promise((resolve) => {
      try {
        if (!db.objectStoreNames.contains(storeName)) return resolve([]);
        const out = [];
        const tx = db.transaction(storeName, "readonly");
        const store = tx.objectStore(storeName);
        const req = store.openCursor();
        req.onsuccess = (ev) => {
          const cur = ev.target.result;
          if (!cur) return resolve(out);
          const k = String(cur.key || "");
          if (!prefix || k.toLowerCase().startsWith(prefix.toLowerCase())) {
            out.push({ key: k, value: cur.value });
          }
          cur.continue();
        };
        req.onerror = () => resolve(out);
      } catch {
        resolve([]);
      }
    });
  }

  async function opfsRead(file) {
    try {
      if (!navigator?.storage?.getDirectory) return null;
      const dir = await navigator.storage.getDirectory();
      const fh = await dir.getFileHandle(file);
      return await (await fh.getFile()).text();
    } catch {
      return null;
    }
  }

  async function opfsWrite(file, text) {
    try {
      if (!navigator?.storage?.getDirectory) return;
      if (navigator.storage.persist) navigator.storage.persist().catch(() => {});
      const dir = await navigator.storage.getDirectory();
      const fh = await dir.getFileHandle(file, { create: true });
      const w = await fh.createWritable();
      await w.write(text);
      await w.close();
    } catch { /* non-fatal */ }
  }

  async function opfsRemove(file) {
    try {
      if (!navigator?.storage?.getDirectory) return;
      const dir = await navigator.storage.getDirectory();
      await dir.removeEntry(file);
    } catch { /* ignore */ }
  }

  async function opfsListProjectFiles() {
    const out = [];
    try {
      if (!navigator?.storage?.getDirectory) return out;
      const dir = await navigator.storage.getDirectory();
      if (typeof dir.entries === "function") {
        for await (const [name, handle] of dir.entries()) {
          if (handle.kind === "file" && /^project-.+\.json$/i.test(name)) {
            try {
              const text = await (await handle.getFile()).text();
              out.push({ fileName: name, text });
            } catch { /* ignore */ }
          }
        }
      }
    } catch { /* ignore */ }
    return out;
  }

  /* ---- Global Master Read / Write (Reconciles ALL v3 + legacy v1.8-v2.6 layers by newest timestamp) ---- */
  function pickNewestMaster(candidates) {
    let best = null;
    for (const c of candidates) {
      if (!c || !Array.isArray(c.materials) || !c.materials.length) continue;
      if (!best) { best = c; continue; }
      const tsA = String(c.updated || "");
      const tsB = String(best.updated || "");
      if (tsA && tsB) {
        if (tsA > tsB) best = c;
      } else if (tsA && !tsB) {
        best = c;
      } else if (!tsA && !tsB && c.materials.length > best.materials.length) {
        best = c;
      }
    }
    return best || { type: "master", updated: "", lastChange: null, materials: [] };
  }

  async function readGlobalMaster() {
    const candidates = [];
    const tryParse = (txt) => {
      if (!txt || typeof txt !== "string") return;
      try {
        const parsed = parseJsonText(txt);
        if (parsed.type !== "project" && parsed.materials.length) candidates.push(parsed);
      } catch { /* ignore corrupt layer */ }
    };

    if (await detectOfflineApi()) {
      try {
        const res = await fetch("/api/master", { cache: "no-store" });
        if (res.ok) tryParse(await res.text());
      } catch { /* fall through */ }
    }

    // Read ALL v3 + legacy v1.8-v2.6 layers in parallel so previously saved materials are ALWAYS discovered!
    const [opfsV3, idbV3, opfsLegacy, idbLegacy] = await Promise.all([
      opfsRead(OPFS_GLOBAL),
      idbGetFrom(DB_NAME, DB_STORE, LS_GLOBAL),
      opfsRead(LEGACY_OPFS_GLOBAL),
      idbGetFrom(LEGACY_DB_NAME, LEGACY_DB_STORE, LEGACY_IDB_GLOBAL)
    ]);

    tryParse(opfsV3);
    tryParse(idbV3);
    tryParse(opfsLegacy);
    tryParse(idbLegacy);
    try {
      tryParse(localStorage.getItem(LS_GLOBAL));
      tryParse(localStorage.getItem(LEGACY_LS_GLOBAL));
    } catch { /* ignore */ }

    const best = pickNewestMaster(candidates);

    // Silently sync v3 layers if we migrated newer data from a legacy v1.8-v2.6 store
    if (best.materials.length && (!opfsV3 || !idbV3)) {
      const json = serializeMaster(best.materials, best.lastChange || null, best.updated || new Date().toISOString());
      try { localStorage.setItem(LS_GLOBAL, json); } catch { /* ignore */ }
      Promise.all([
        idbSetIn(DB_NAME, DB_STORE, LS_GLOBAL, json),
        opfsWrite(OPFS_GLOBAL, json)
      ]).catch(() => {});
    }
    return best;
  }

  async function writeGlobalMaster(list, lastChange) {
    const normalized = list.map(normalizeMaterial).filter((m) => m.material);
    const nowIso = new Date().toISOString();
    const json = serializeMaster(normalized, lastChange || null, nowIso);
    try {
      localStorage.setItem(LS_GLOBAL, json);
      localStorage.setItem(LEGACY_LS_GLOBAL, json);
    } catch { /* ignore */ }
    await Promise.all([
      idbSetIn(DB_NAME, DB_STORE, LS_GLOBAL, json),
      opfsWrite(OPFS_GLOBAL, json),
      idbSetIn(LEGACY_DB_NAME, LEGACY_DB_STORE, LEGACY_IDB_GLOBAL, json),
      opfsWrite(LEGACY_OPFS_GLOBAL, json)
    ]);
    if (await detectOfflineApi()) {
      try {
        await fetch("/api/master", { method: "POST", headers: { "Content-Type": "application/json" }, body: json });
      } catch { /* non-fatal */ }
    }
    return { count: normalized.length, updated: nowIso, json };
  }

  /* ---- Saved Project-Only Material Lists (Full v1.9-v3.0 Migration & Per-Project Auto-Save) ---- */
  async function readProjectMap() {
    const bySlug = new Map(); // slugify(projectName) -> entry

    const absorbEntry = (entry) => {
      if (!entry || !entry.projectName) return;
      const pName = String(entry.projectName).trim();
      if (!pName) return;
      const key = slugify(pName);
      const mats = Array.isArray(entry.materials)
        ? entry.materials.map(normalizeMaterial).filter((m) => m.material)
        : [];
      const cand = {
        id: String(entry.id || key).trim() || key,
        projectName: pName,
        updated: String(entry.updated || ""),
        lastChange: entry.lastChange || null,
        materials: mats
      };
      const prev = bySlug.get(key);
      if (!prev) {
        bySlug.set(key, cand);
        return;
      }
      // Prefer newest updated timestamp (or richer list when timestamps tie/absent)
      if (cand.updated && prev.updated) {
        if (cand.updated > prev.updated) bySlug.set(key, cand);
      } else if (cand.updated && !prev.updated) {
        bySlug.set(key, cand);
      } else if (!cand.updated && !prev.updated && cand.materials.length > prev.materials.length) {
        bySlug.set(key, cand);
      }
    };

    const absorbMapJson = (txt) => {
      if (!txt || typeof txt !== "string") return;
      try {
        const obj = JSON.parse(txt);
        if (obj && typeof obj === "object" && !Array.isArray(obj)) {
          Object.values(obj).forEach(absorbEntry);
        }
      } catch { /* ignore */ }
    };

    const absorbSingleProjectJson = (txt, fallbackName) => {
      if (!txt || typeof txt !== "string") return;
      try {
        const p = parseJsonText(txt);
        const name = p.projectName || String(fallbackName || "").replace(/[-_]+/g, " ").trim();
        if (!name) return;
        absorbEntry({
          id: p.id || slugify(name),
          projectName: name,
          updated: p.updated,
          lastChange: p.lastChange,
          materials: p.materials
        });
      } catch { /* ignore */ }
    };

    // 1. v3 & legacy localStorage keys (bomTools_projectLists_v3 + bomProjectJson:<slug> + bomProject:<slug>)
    try {
      absorbMapJson(localStorage.getItem(LS_PROJECTS));
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (!k) continue;
        if (/^bomProjectJson:/i.test(k)) {
          absorbSingleProjectJson(localStorage.getItem(k), k.replace(/^bomProjectJson:/i, ""));
        } else if (/^bomProject[:_]/i.test(k) && !/^bomProjectMeta:/i.test(k)) {
          absorbSingleProjectJson(localStorage.getItem(k), k.replace(/^bomProject[:_]/i, ""));
        }
      }
    } catch { /* ignore */ }

    // 2. v3 IndexedDB + v3 OPFS + legacy BomToolsLocalDB (projectJson:*) + legacy OPFS (project-*.json)
    const [idbV3, opfsV3, legacyIdbEntries, legacyOpfsFiles] = await Promise.all([
      idbGetFrom(DB_NAME, DB_STORE, LS_PROJECTS),
      opfsRead(OPFS_PROJECTS),
      idbEntriesFrom(LEGACY_DB_NAME, LEGACY_DB_STORE, "projectJson:"),
      opfsListProjectFiles()
    ]);

    absorbMapJson(idbV3);
    absorbMapJson(opfsV3);

    for (const row of legacyIdbEntries) {
      if (typeof row.value === "string") {
        absorbSingleProjectJson(row.value, row.key.replace(/^projectJson:/i, ""));
      }
    }
    for (const file of legacyOpfsFiles) {
      const slugName = file.fileName.replace(/^project-/i, "").replace(/\.json$/i, "");
      absorbSingleProjectJson(file.text, slugName);
    }

    // 3. Offline LTS server API (/api/projects)
    if (await detectOfflineApi()) {
      try {
        const r = await fetch("/api/projects", { cache: "no-store" });
        if (r.ok) {
          const data = await r.json();
          if (Array.isArray(data)) data.forEach(absorbEntry);
        }
      } catch { /* ignore */ }
    }

    const out = {};
    for (const entry of bySlug.values()) out[entry.id] = entry;
    return out;
  }

  async function persistProjectMap(map) {
    const json = JSON.stringify(map);
    try { localStorage.setItem(LS_PROJECTS, json); } catch { /* ignore */ }
    await Promise.all([
      idbSetIn(DB_NAME, DB_STORE, LS_PROJECTS, json),
      opfsWrite(OPFS_PROJECTS, json)
    ]);
  }

  async function listSavedProjects() {
    const map = await readProjectMap();
    return Object.values(map).sort((a, b) => String(b.updated).localeCompare(String(a.updated)));
  }

  async function getSavedProject(idOrName) {
    const map = await readProjectMap();
    if (map[idOrName]) return map[idOrName];
    const target = lc(idOrName);
    const targetSlug = slugify(idOrName);
    return Object.values(map).find(
      (p) => lc(p.projectName) === target || slugify(p.projectName) === targetSlug || lc(p.id) === target
    ) || null;
  }

  async function saveProjectList(meta, materials, lastChange) {
    const map = await readProjectMap();
    const projectName = String(meta?.projectName || meta?.name || "Project").trim();
    const s = slugify(projectName);

    // Reuse existing entry ID if same project slug already exists (prevents duplicates!)
    let existingId = meta?.id || "";
    if (!existingId || !map[existingId]) {
      const byName = Object.values(map).find((p) => slugify(p.projectName) === s);
      if (byName) existingId = byName.id;
    }
    const id = existingId || s;
    const nowIso = new Date().toISOString();
    const cleanMats = materials.map(normalizeMaterial).filter((m) => m.material);

    const entry = {
      id,
      projectName,
      updated: nowIso,
      lastChange: lastChange || null,
      materials: cleanMats
    };
    map[id] = entry;
    await persistProjectMap(map);

    // Mirror to legacy v1.9-v2.6 per-project slots so project auto-save is rock-solid across versions
    const singleJson = serializeProject(projectName, cleanMats, entry.lastChange, id, nowIso);
    try {
      localStorage.setItem(`bomProjectJson:${s}`, singleJson);
    } catch { /* ignore */ }
    await Promise.all([
      idbSetIn(LEGACY_DB_NAME, LEGACY_DB_STORE, `projectJson:${s}`, singleJson),
      opfsWrite(`project-${s}.json`, singleJson)
    ]);

    if (await detectOfflineApi()) {
      try {
        await fetch(`/api/projects/${encodeURIComponent(id)}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: singleJson
        });
      } catch { /* non-fatal */ }
    }
    return entry;
  }

  async function deleteSavedProject(id) {
    const map = await readProjectMap();
    const target = map[id];
    if (!target) return false;
    const s = slugify(target.projectName);
    delete map[id];
    await persistProjectMap(map);

    // Also remove matching legacy per-project keys so it never reappears
    try {
      localStorage.removeItem(`bomProjectJson:${s}`);
      localStorage.removeItem(`bomProjectMeta:${s}`);
      localStorage.removeItem(`bomProject:${s}`);
    } catch { /* ignore */ }
    await Promise.all([
      idbDelIn(LEGACY_DB_NAME, LEGACY_DB_STORE, `projectJson:${s}`),
      opfsRemove(`project-${s}.json`)
    ]);

    if (await detectOfflineApi()) {
      try { await fetch(`/api/projects/${encodeURIComponent(id)}`, { method: "DELETE" }); } catch { /* ignore */ }
    }
    return true;
  }

  /* ---- Unified Pool of ALL Previously Created Materials (Global Master + Workbook + All Saved Projects) ---- */
  async function getAvailableMaterialPool(workbookRows = []) {
    const [globalData, projects] = await Promise.all([readGlobalMaster(), listSavedProjects()]);
    const pool = new Map(); // lc(name) -> { ...material, sources: [], inGlobal: boolean }

    const add = (raw, sourceLabel, isGlobal = false) => {
      const m = normalizeMaterial(raw);
      if (!m.material) return;
      const k = lc(m.material);
      const prev = pool.get(k);
      if (!prev) {
        pool.set(k, { ...m, sources: [sourceLabel], inGlobal: isGlobal });
      } else {
        if (!prev.sources.includes(sourceLabel)) prev.sources.push(sourceLabel);
        if (isGlobal) prev.inGlobal = true;
      }
    };

    // Priority order: Global Master first, then current workbook rows, then saved Project-Only lists (newest first)
    for (const m of globalData.materials) add(m, "Global Master", true);
    for (const m of workbookRows || []) add(m, "This Workbook", false);
    for (const p of projects) {
      for (const m of p.materials || []) add(m, `Project: ${p.projectName}`, false);
    }

    return [...pool.values()].sort((a, b) =>
      a.material.localeCompare(b.material, undefined, { sensitivity: "base" })
    );
  }

  /* ---------------- Portable JSON file save/download helper ---------------- */
  async function saveJsonFile(suggestedName, jsonText) {
    if (typeof window !== "undefined" && typeof window.showSaveFilePicker === "function") {
      try {
        const fh = await window.showSaveFilePicker({
          suggestedName,
          types: [{ description: "BOM Materials JSON", accept: { "application/json": [".json"] } }]
        });
        const w = await fh.createWritable();
        await w.write(jsonText);
        await w.close();
        return { fileName: fh.name, cancelled: false };
      } catch (err) {
        if (err && err.name === "AbortError") return { fileName: "", cancelled: true };
      }
    }
    const blob = new Blob([jsonText], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = suggestedName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
    return { fileName: suggestedName, cancelled: false };
  }

  /* ---------------- Optional GitHub Sync (Global Master only — never Project-Only!) ---------------- */
  function getToken() {
    try { return localStorage.getItem(LS_TOKEN) || localStorage.getItem(LEGACY_LS_TOKEN) || ""; }
    catch { return ""; }
  }
  function setToken(t) {
    try {
      if (t && t.trim()) localStorage.setItem(LS_TOKEN, t.trim());
      else {
        localStorage.removeItem(LS_TOKEN);
        localStorage.removeItem(LEGACY_LS_TOKEN);
      }
      return true;
    } catch { return false; }
  }
  function detectRepo() {
    try {
      if (/\.github\.io$/i.test(location.hostname)) {
        const owner = location.hostname.split(".")[0];
        const seg = location.pathname.split("/").filter(Boolean)[0];
        const repo = seg && !seg.includes(".") ? seg : `${owner}.github.io`;
        return { owner, repo, branch: "main" };
      }
    } catch { /* ignore */ }
    return { owner: "", repo: "", branch: "main" };
  }

  async function fetchGitHubPages() {
    const url = new URL("materials.json", location.href).href + "?t=" + Date.now();
    const res = await fetch(url, { cache: "no-store" });
    if (res.status === 404) return { missing: true, materials: [] };
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return { missing: false, materials: parseJsonText(await res.text()).materials };
  }

  async function pushGitHub(materials, message) {
    const tok = getToken();
    if (!tok) throw new Error("No GitHub token saved.");
    const { owner, repo, branch } = detectRepo();
    if (!owner || !repo) throw new Error("GitHub repository could not be detected from host URL.");
    const api = `https://api.github.com/repos/${owner}/${repo}/contents/materials.json`;
    const headers = { Accept: "application/vnd.github+json", Authorization: `Bearer ${tok}` };
    let sha = null;
    const getRes = await fetch(`${api}?ref=${encodeURIComponent(branch)}&t=${Date.now()}`, { headers, cache: "no-store" });
    if (getRes.ok) sha = (await getRes.json()).sha;
    const json = serializeMaster(materials, null);
    const bytes = new TextEncoder().encode(json);
    let bin = "";
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    const putRes = await fetch(api, {
      method: "PUT",
      headers: { ...headers, "Content-Type": "application/json" },
      body: JSON.stringify({ message: message || "Update materials.json", content: btoa(bin), branch, ...(sha ? { sha } : {}) })
    });
    if (!putRes.ok) throw new Error(`GitHub rejected publish (${putRes.status}).`);
    return true;
  }

  window.BomStore = {
    VERSION,
    UNITS,
    slugify,
    normalizeMaterial,
    changedFields,
    diffLists,
    buildLastChange,
    serializeMaster,
    serializeProject,
    parseJsonText,
    detectOfflineApi,
    readGlobalMaster,
    writeGlobalMaster,
    listSavedProjects,
    getSavedProject,
    saveProjectList,
    deleteSavedProject,
    getAvailableMaterialPool,
    saveJsonFile,
    getToken,
    setToken,
    hasToken: () => Boolean(getToken()),
    fetchGitHubPages,
    pushGitHub
  };
})();
