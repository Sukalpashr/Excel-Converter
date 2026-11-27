/* storage.js - Lightweight Built-in Storage, Project-Only Lists, Offline LTS Bridge & Optional GitHub Sync (v3.0) */
(() => {
  "use strict";

  const VERSION = "3.0";
  const DB_NAME = "BomToolsV3DB";
  const DB_STORE = "kv";
  const LS_GLOBAL = "bomTools_globalMaster_v3";
  const LS_PROJECTS = "bomTools_projectLists_v3";
  const LS_TOKEN = "bomTools_ghToken_v3";
  const LEGACY_LS_GLOBAL = "bomLocalMasterJson";
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
    const out = {
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
    return out;
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

  function serializeMaster(list, lastChange) {
    const payload = {
      version: 3,
      type: "master",
      updated: new Date().toISOString(),
      ...(lastChange ? { lastChange } : {}),
      materials: list.map(cleanForJson)
    };
    return JSON.stringify(payload, null, 2) + "\n";
  }

  function serializeProject(projectName, list, lastChange, id) {
    const payload = {
      version: 3,
      type: "project",
      id: id || `${slugify(projectName)}-${Date.now().toString(36)}`,
      projectName: String(projectName || "Project").trim(),
      updated: new Date().toISOString(),
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
      projectName: isObj ? String(raw.projectName || raw.project || "") : "",
      updated: isObj ? String(raw.updated || "") : "",
      lastChange: isObj ? raw.lastChange || null : null,
      materials
    };
  }

  /* ---------------- Multi-layer zero-reconnect storage (OPFS + IndexedDB + localStorage + Offline LTS) ---------------- */
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

  function openDb() {
    return new Promise((resolve) => {
      if (typeof indexedDB === "undefined") return resolve(null);
      try {
        const req = indexedDB.open(DB_NAME, 1);
        req.onupgradeneeded = () => {
          const db = req.result;
          if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE);
        };
        req.onsuccess = () => resolve(req.result);
        req.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });
  }

  async function idbGet(key) {
    const db = await openDb();
    if (!db) return null;
    return new Promise((resolve) => {
      try {
        const req = db.transaction(DB_STORE, "readonly").objectStore(DB_STORE).get(key);
        req.onsuccess = () => resolve(req.result ?? null);
        req.onerror = () => resolve(null);
      } catch {
        resolve(null);
      }
    });
  }

  async function idbSet(key, val) {
    const db = await openDb();
    if (!db) return;
    await new Promise((resolve) => {
      try {
        const tx = db.transaction(DB_STORE, "readwrite");
        tx.objectStore(DB_STORE).put(val, key);
        tx.oncomplete = () => resolve();
        tx.onerror = () => resolve();
      } catch {
        resolve();
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
    } catch {
      /* non-fatal */
    }
  }

  /* ---- Global Master Read / Write ---- */
  async function readGlobalMaster() {
    if (await detectOfflineApi()) {
      try {
        const res = await fetch("/api/master", { cache: "no-store" });
        if (res.ok) {
          const parsed = parseJsonText(await res.text());
          if (parsed.materials.length) return parsed;
        }
      } catch { /* fall through */ }
    }
    let raw = await opfsRead(OPFS_GLOBAL);
    if (!raw) raw = await idbGet(LS_GLOBAL);
    if (!raw) {
      try { raw = localStorage.getItem(LS_GLOBAL) || localStorage.getItem(LEGACY_LS_GLOBAL); } catch { raw = null; }
    }
    if (!raw) return { type: "master", updated: "", lastChange: null, materials: [] };
    try {
      return parseJsonText(raw);
    } catch {
      return { type: "master", updated: "", lastChange: null, materials: [] };
    }
  }

  async function writeGlobalMaster(list, lastChange) {
    const normalized = list.map(normalizeMaterial).filter((m) => m.material);
    const json = serializeMaster(normalized, lastChange || null);
    try { localStorage.setItem(LS_GLOBAL, json); } catch { /* ignore */ }
    await Promise.all([idbSet(LS_GLOBAL, json), opfsWrite(OPFS_GLOBAL, json)]);
    if (await detectOfflineApi()) {
      try {
        await fetch("/api/master", { method: "POST", headers: { "Content-Type": "application/json" }, body: json });
      } catch { /* non-fatal */ }
    }
    return { count: normalized.length, json };
  }

  /* ---- Saved Project-Only Material Lists (Isolated per-project slots on this PC) ---- */
  async function readProjectMap() {
    const merged = {};
    const absorb = (txt) => {
      if (!txt) return;
      try {
        const obj = JSON.parse(txt);
        if (obj && typeof obj === "object") {
          for (const [id, entry] of Object.entries(obj)) {
            if (!entry || !entry.projectName) continue;
            const prev = merged[id];
            if (!prev || String(entry.updated || "") >= String(prev.updated || "")) {
              merged[id] = {
                id,
                projectName: String(entry.projectName),
                updated: String(entry.updated || ""),
                lastChange: entry.lastChange || null,
                materials: Array.isArray(entry.materials)
                  ? entry.materials.map(normalizeMaterial).filter((m) => m.material)
                  : []
              };
            }
          }
        }
      } catch { /* ignore */ }
    };

    try { absorb(localStorage.getItem(LS_PROJECTS)); } catch { /* ignore */ }
    absorb(await idbGet(LS_PROJECTS));
    absorb(await opfsRead(OPFS_PROJECTS));

    if (await detectOfflineApi()) {
      try {
        const r = await fetch("/api/projects", { cache: "no-store" });
        if (r.ok) {
          const data = await r.json();
          if (Array.isArray(data)) {
            for (const p of data) {
              if (p && p.id && p.projectName) {
                merged[p.id] = {
                  id: p.id,
                  projectName: p.projectName,
                  updated: p.updated || "",
                  lastChange: p.lastChange || null,
                  materials: (p.materials || []).map(normalizeMaterial).filter((m) => m.material)
                };
              }
            }
          }
        }
      } catch { /* ignore */ }
    }
    return merged;
  }

  async function persistProjectMap(map) {
    const json = JSON.stringify(map);
    try { localStorage.setItem(LS_PROJECTS, json); } catch { /* ignore */ }
    await Promise.all([idbSet(LS_PROJECTS, json), opfsWrite(OPFS_PROJECTS, json)]);
  }

  async function listSavedProjects() {
    const map = await readProjectMap();
    return Object.values(map).sort((a, b) => String(b.updated).localeCompare(String(a.updated)));
  }

  async function getSavedProject(idOrName) {
    const map = await readProjectMap();
    if (map[idOrName]) return map[idOrName];
    const target = lc(idOrName);
    return Object.values(map).find((p) => lc(p.projectName) === target) || null;
  }

  async function saveProjectList(meta, materials, lastChange) {
    const map = await readProjectMap();
    const id = meta.id || `${slugify(meta.projectName)}-${Date.now().toString(36)}`;
    const entry = {
      id,
      projectName: String(meta.projectName || "Project").trim(),
      updated: new Date().toISOString(),
      lastChange: lastChange || null,
      materials: materials.map(normalizeMaterial).filter((m) => m.material)
    };
    map[id] = entry;
    await persistProjectMap(map);
    if (await detectOfflineApi()) {
      try {
        await fetch(`/api/projects/${encodeURIComponent(id)}`, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: serializeProject(entry.projectName, entry.materials, entry.lastChange, id)
        });
      } catch { /* non-fatal */ }
    }
    return entry;
  }

  async function deleteSavedProject(id) {
    const map = await readProjectMap();
    if (!map[id]) return false;
    delete map[id];
    await persistProjectMap(map);
    if (await detectOfflineApi()) {
      try { await fetch(`/api/projects/${encodeURIComponent(id)}`, { method: "DELETE" }); } catch { /* ignore */ }
    }
    return true;
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
  function getToken() { try { return localStorage.getItem(LS_TOKEN) || ""; } catch { return ""; } }
  function setToken(t) {
    try {
      if (t && t.trim()) localStorage.setItem(LS_TOKEN, t.trim());
      else localStorage.removeItem(LS_TOKEN);
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
    saveJsonFile,
    getToken,
    setToken,
    hasToken: () => Boolean(getToken()),
    fetchGitHubPages,
    pushGitHub
  };
})();
