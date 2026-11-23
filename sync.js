/* sync.js - keeps ONE "latest master materials list" on this computer and reconciles it with
 * the master sheet (table tblMat) of whichever workbook is open.
 *
 *  - store  : browser localStorage of the add-in's origin (shared by all workbooks on this PC)
 *  - base   : per-workbook record (workbook settings, saved inside the file) of what each material
 *             looked like the last time that workbook was in sync. It tells "edited here" from "stale".
 * Pure logic first (unit-tested in node), Excel I/O at the bottom.
 */
(function (root) {
  "use strict";

  var STORE_KEY = "bomMasterStore.v1";
  var BASE_KEY = "bomMasterBase";
  var mem = null;                         // fallback when localStorage is unavailable (or in tests)
  var persistentOk = true;

  function lc(s) { return String(s).trim().toLowerCase(); }
  function str(v) { return v === null || v === undefined ? "" : String(v).trim(); }
  function n6(v) {
    if (v === null || v === undefined || v === "") return null;
    var n = typeof v === "number" ? v : Number(String(v).replace(/,/g, ""));
    return isFinite(n) ? Math.round(n * 1e6) / 1e6 : null;
  }

  function canon(r) {
    return { material: str(r.material), category: str(r.category), unit: str(r.unit),
             lin: n6(r.lin), area: n6(r.area), piece: n6(r.piece), sl: n6(r.sl), sw: n6(r.sw), remarks: str(r.remarks) };
  }
  function hash(r) {
    var c = canon(r);
    var s = JSON.stringify([c.material, c.category, c.unit, c.lin, c.area, c.piece, c.sl, c.sw, c.remarks]);
    var h = 5381;
    for (var i = 0; i < s.length; i++) h = ((h << 5) + h + s.charCodeAt(i)) | 0;
    return (h >>> 0).toString(36);
  }

  /* Fold a workbook's materials into the store. Never touches the workbook.
   *  new name                              -> added to the store
   *  same name, same values                -> nothing
   *  same name, different values:
   *      workbook unchanged since last sync -> store is newer, skip
   *      store unchanged since last sync    -> edited in this workbook, store takes the workbook's version
   *      both changed / never synced        -> skip (reported as conflict only when both changed) */
  function capture(store, rows, base) {
    var idx = {};
    store.rows.forEach(function (r, i) { idx[lc(r.material)] = i; });
    var bh = (base && base.h) || {};
    var res = { added: 0, updated: 0, skipped: 0, conflicts: [] };
    var seen = {};
    rows.forEach(function (raw) {
      var w = canon(raw), k = lc(w.material);
      if (!k || seen[k]) return;
      seen[k] = 1;
      var i = idx[k];
      if (i === undefined) { store.rows.push(w); idx[k] = store.rows.length - 1; res.added++; return; }
      var s = store.rows[i], hw = hash(w), hs = hash(s);
      if (hw === hs) return;
      var b = bh[k];
      if (b === undefined || hw === b) { res.skipped++; return; }
      if (hs === b) { store.rows[i] = w; res.updated++; return; }
      res.conflicts.push(w.material);
    });
    if (res.added || res.updated) store.updatedAt = Date.now();
    return res;
  }

  /* What to write into the workbook. rows = [{row: bodyIndex, ...fields}]
   *  mode "new"  : append materials the workbook does not have
   *  mode "auto" : "new" + refresh materials the workbook has not changed since its last sync
   *  mode "all"  : "new" + overwrite every differing material with the store's version */
  function plan(store, rows, base, mode) {
    var wIdx = {};
    rows.forEach(function (r) { var k = lc(r.material); if (k && wIdx[k] === undefined) wIdx[k] = r; });
    var bh = (base && base.h) || {};
    var out = { appends: [], updates: [] };
    store.rows.forEach(function (s) {
      var k = lc(s.material), w = wIdx[k];
      if (!w) { out.appends.push(canon(s)); return; }
      var hw = hash(w);
      if (hw === hash(s)) return;
      if (mode === "all" || (mode === "auto" && bh[k] !== undefined && bh[k] === hw)) out.updates.push({ row: w.row, rec: canon(s) });
    });
    return out;
  }

  function pending(store, rows) {
    var wIdx = {};
    rows.forEach(function (r) { var k = lc(r.material); if (k && wIdx[k] === undefined) wIdx[k] = r; });
    var n = 0, d = 0;
    store.rows.forEach(function (s) {
      var w = wIdx[lc(s.material)];
      if (!w) n++; else if (hash(w) !== hash(s)) d++;
    });
    return { newCount: n, diffCount: d };
  }

  function newBase(store, rows, old) {
    var sIdx = {};
    store.rows.forEach(function (s) { sIdx[lc(s.material)] = s; });
    var oh = (old && old.h) || {}, h = {};
    rows.forEach(function (r) {
      var k = lc(r.material); if (!k) return;
      var s = sIdx[k];
      if (s && hash(s) === hash(r)) h[k] = hash(r);
      else if (oh[k] !== undefined) h[k] = oh[k];
    });
    return { v: 1, h: h };
  }

  function upsert(store, rec) {          // explicit add/edit made through the add-in
    var c = canon(rec), k = lc(c.material), at = -1;
    store.rows.forEach(function (r, i) { if (at < 0 && lc(r.material) === k) at = i; });
    if (at < 0) store.rows.push(c); else store.rows[at] = c;
    store.updatedAt = Date.now();
  }

  /* ---------------- store (localStorage) ---------------- */
  function emptyStore() { return { v: 1, rows: [], updatedAt: 0 }; }
  function loadStore() {
    try {
      if (typeof localStorage !== "undefined") {
        var s = JSON.parse(localStorage.getItem(STORE_KEY));
        if (s && Array.isArray(s.rows)) return s;
        return emptyStore();
      }
    } catch (e) { persistentOk = false; }
    return mem ? JSON.parse(JSON.stringify(mem)) : emptyStore();
  }
  function saveStore(s) {
    try {
      if (typeof localStorage !== "undefined") { localStorage.setItem(STORE_KEY, JSON.stringify(s)); return true; }
    } catch (e) { persistentOk = false; }
    mem = JSON.parse(JSON.stringify(s));
    if (typeof localStorage === "undefined") persistentOk = false;
    return false;
  }
  function persistent() { return persistentOk && typeof localStorage !== "undefined"; }
  function resetMemory() { mem = null; persistentOk = true; }

  /* ---------------- per-workbook base (workbook settings) ---------------- */
  async function readBase(context) {
    var it = context.workbook.settings.getItemOrNullObject(BASE_KEY);
    it.load("value");
    await context.sync();
    if (it.isNullObject) return null;
    try { var b = JSON.parse(it.value); return b && b.h ? b : null; } catch (e) { return null; }
  }
  async function writeBase(context, obj) {
    var it = context.workbook.settings.getItemOrNullObject(BASE_KEY);
    await context.sync();
    var v = JSON.stringify(obj);
    if (it.isNullObject) context.workbook.settings.add(BASE_KEY, v); else it.value = v;
    await context.sync();
  }

  var api = { canon: canon, hash: hash, capture: capture, plan: plan, pending: pending, newBase: newBase, upsert: upsert,
              loadStore: loadStore, saveStore: saveStore, persistent: persistent, resetMemory: resetMemory,
              readBase: readBase, writeBase: writeBase, lc: lc };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.BomSync = api;
})(typeof window !== "undefined" ? window : this);
