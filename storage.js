/* storage.js — window.BomStore
 * ----------------------------------------------------------------------------
 * Built-in storage for the BOM Tools add-in. No backend required.
 *
 * Layers (written in parallel, read with first-wins):
 *   1. IndexedDB  — primary, async, large
 *   2. OPFS       — Origin Private File System, async, large (when supported)
 *   3. localStorage — sync, small (~5 MB), always available fallback
 *
 * Project-Only isolation:
 *   - The GLOBAL master lives under key "bom_global_master".
 *   - Each project lives under key "bom_project:<name>".
 *   - Switching the active scope only changes which slot BomStore.getMaster()
 *     returns. The Global slot and GitHub remote are NEVER touched by project
 *     operations unless the user explicitly pushes Global.
 *
 * JSON shapes:
 *   Master  : { type:"master",  version, exportedAt, materials:[...] }
 *   Project : { type:"project",  version, exportedAt, name, bomProjectList,
 *               materials:[...], sheets:[{name, bom:[], po:[]}] }
 *
 *   Importing a "project"-typed file via importMaster() is REJECTED — only
 *   importProject() accepts project files.
 *
 * Offline API (optional): when hosted on localhost and /api/health responds,
 *   BomStore also pulls/pushes from /api/master and /api/projects. Detection
 *   is best-effort; absence is silent.
 */
(function (root) {
  'use strict';

  var DB_NAME = 'bom_tools';
  var DB_STORE = 'kv';
  var DB_VERSION = 1;
  var LS_PREFIX = 'bom_tools:';
  var KEY_GLOBAL = 'bom_global_master';
  var KEY_PROJECT_PREFIX = 'bom_project:';
  var KEY_ACTIVE_PROJECT = 'bom_active_project';
  var KEY_AUDIT = 'bom_audit_log';

  var DEFAULT_MATERIALS = []; // populated lazily from the embedded Master sheet

  // ---------- small utils ----------
  function nowISO() { return new Date().toISOString(); }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function promise(fn) { return new Promise(fn); }
  function logErr(label, err) { console.warn('[BomStore] ' + label + ':', err && err.message || err); }

  // ---------- listeners ----------
  var listeners = {};
  function on(evt, cb) {
    (listeners[evt] = listeners[evt] || []).push(cb);
    return function () { off(evt, cb); };
  }
  function off(evt, cb) {
    var arr = listeners[evt]; if (!arr) return;
    var i = arr.indexOf(cb); if (i >= 0) arr.splice(i, 1);
  }
  function emit(evt, payload) {
    var arr = listeners[evt]; if (!arr) return;
    arr.slice().forEach(function (cb) { try { cb(payload); } catch (e) { logErr('listener', e); } });
  }

  // ---------- IndexedDB ----------
  var _db = null;
  function idbOpen() {
    return promise(function (resolve, reject) {
      if (_db) return resolve(_db);
      if (!root.indexedDB) return reject(new Error('no-idb'));
      var req = root.indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE);
      };
      req.onsuccess = function () { _db = req.result; resolve(_db); };
      req.onerror = function () { reject(req.error); };
    });
  }
  function idbGet(key) {
    return idbOpen().then(function (db) {
      return promise(function (resolve, reject) {
        var tx = db.transaction(DB_STORE, 'readonly');
        var req = tx.objectStore(DB_STORE).get(key);
        req.onsuccess = function () { resolve(req.result); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }
  function idbSet(key, val) {
    return idbOpen().then(function (db) {
      return promise(function (resolve, reject) {
        var tx = db.transaction(DB_STORE, 'readwrite');
        tx.objectStore(DB_STORE).put(val, key);
        tx.oncomplete = function () { resolve(true); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }
  function idbKeys(prefix) {
    return idbOpen().then(function (db) {
      return promise(function (resolve, reject) {
        var tx = db.transaction(DB_STORE, 'readonly');
        var store = tx.objectStore(DB_STORE);
        var out = [];
        var cur = store.openKeyCursor ? store.openKeyCursor() : store.openCursor();
        cur.onsuccess = function () {
          var c = cur.result;
          if (!c) return resolve(out);
          var k = c.key;
          if (typeof k === 'string' && (!prefix || k.indexOf(prefix) === 0)) out.push(k);
          c.continue();
        };
        cur.onerror = function () { reject(cur.error); };
      });
    });
  }

  // ---------- OPFS ----------
  function opfsRoot() {
    if (!root.navigator || !navigator.storage || !navigator.storage.getDirectory) return Promise.resolve(null);
    return navigator.storage.getDirectory().then(function (d) { return d; }, function () { return null; });
  }
  function opfsRead(name) {
    return opfsRoot().then(function (root) {
      if (!root) return null;
      return root.getFileHandle(name, { create: false }).then(function (fh) {
        return fh.getFile().then(function (f) { return f.text(); });
      }, function () { return null; });
    });
  }
  function opfsWrite(name, text) {
    return opfsRoot().then(function (root) {
      if (!root) return false;
      return root.getFileHandle(name, { create: true }).then(function (fh) {
        return fh.createWritable().then(function (w) {
          return w.write(text).then(function () { return w.close(); });
        }).then(function () { return true; });
      }, function () { return false; });
    });
  }

  // ---------- localStorage ----------
  function lsGet(key) {
    try { var v = root.localStorage.getItem(LS_PREFIX + key); return v ? JSON.parse(v) : null; }
    catch (e) { return null; }
  }
  function lsSet(key, val) {
    try { root.localStorage.setItem(LS_PREFIX + key, JSON.stringify(val)); return true; }
    catch (e) { logErr('localStorage.set', e); return false; }
  }
  function lsKeys(prefix) {
    var out = [];
    try {
      for (var i = 0; i < root.localStorage.length; i++) {
        var full = root.localStorage.key(i);
        if (full && full.indexOf(LS_PREFIX) === 0) {
          var k = full.slice(LS_PREFIX.length);
          if (!prefix || k.indexOf(prefix) === 0) out.push(k);
        }
      }
    } catch (e) {}
    return out;
  }

  // ---------- unified get/set (parallel write, first-wins read) ----------
  function persist(key, val) {
    // localStorage is sync — write first so it is always current even if IDB/OPFS lag.
    lsSet(key, val);
    var p1 = idbSet(key, val).catch(function (e) { logErr('idb.set', e); });
    var p2 = opfsWrite(opfsName(key), JSON.stringify(val)).catch(function (e) { logErr('opfs.set', e); });
    return Promise.all([p1, p2]).then(function () { return true; });
  }
  function load(key) {
    return idbGet(key).then(function (idbVal) {
      if (idbVal !== undefined && idbVal !== null) return idbVal;
      return opfsRead(opfsName(key)).then(function (txt) {
        if (txt) { try { return JSON.parse(txt); } catch (e) {} }
        return lsGet(key);
      });
    });
  }
  function opfsName(key) { return key.replace(/[^a-z0-9_-]/gi, '_') + '.json'; }

  // ---------- keys ----------
  function projectKey(name) { return KEY_PROJECT_PREFIX + name; }
  function globalKey() { return KEY_GLOBAL; }

  // ---------- ACTIVE SCOPE ----------
  function getActiveProject() {
    return lsGet(KEY_ACTIVE_PROJECT) || null;
  }
  function setActiveProject(name) {
    if (name) lsSet(KEY_ACTIVE_PROJECT, name);
    else try { root.localStorage.removeItem(LS_PREFIX + KEY_ACTIVE_PROJECT); } catch (e) {}
    emit('scope', { project: name || null });
  }
  function activeScopeLabel() {
    var p = getActiveProject();
    return p ? ('Project: ' + p) : 'Global Master';
  }

  // ---------- MATERIALS GET / SAVE ----------
  function getMaster(projectName) {
    var p = projectName !== undefined ? projectName : getActiveProject();
    var key = p ? projectKey(p) : globalKey();
    return load(key).then(function (rec) {
      if (rec && rec.materials) return rec.materials;
      return null;
    });
  }
  function saveMaster(materials, projectName, opts) {
    opts = opts || {};
    var p = projectName !== undefined ? projectName : getActiveProject();
    var key = p ? projectKey(p) : globalKey();
    var rec = {
      type: p ? 'project' : 'master',
      version: 1,
      name: p || null,
      materials: materials,
      savedAt: nowISO()
    };
    return persist(key, rec).then(function () {
      emit('saved', { project: p, materials: materials });
      if (!opts.silent) emit('change', { scope: p ? 'project' : 'global', project: p, materials: materials });
      return true;
    });
  }

  // ---------- PROJECTS ----------
  function listProjects() {
    var fromIdb = idbKeys(KEY_PROJECT_PREFIX).then(function (ks) {
      return ks.map(function (k) { return k.slice(KEY_PROJECT_PREFIX.length); });
    }, function () { return []; });
    var fromLs = lsKeys(KEY_PROJECT_PREFIX).map(function (k) { return k.slice(KEY_PROJECT_PREFIX.length); });
    var seen = {}; var out = [];
    fromLs.forEach(function (n) { if (!seen[n]) { seen[n] = 1; out.push(n); } });
    return fromIdb.then(function (names) {
      names.forEach(function (n) { if (!seen[n]) { seen[n] = 1; out.push(n); } });
      return out.slice().sort();
    });
  }
  function listSavedProjects() {
    return listProjects().then(function (names) {
      return Promise.all(names.map(function (n) {
        return load(projectKey(n)).then(function (rec) {
          return {
            name: n,
            materialsCount: (rec && rec.materials) ? rec.materials.length : 0,
            savedAt: rec && rec.savedAt,
            bomProjectList: rec && rec.bomProjectList
          };
        });
      }));
    });
  }
  function deleteProject(name) {
    var key = projectKey(name);
    try { root.localStorage.removeItem(LS_PREFIX + key); } catch (e) {}
    return idbOpen().then(function (db) {
      return promise(function (resolve, reject) {
        var tx = db.transaction(DB_STORE, 'readwrite');
        tx.objectStore(DB_STORE).delete(key);
        tx.oncomplete = function () { resolve(true); };
        tx.onerror = function () { reject(tx.error); };
      });
    }).then(function () {
      opfsRoot().then(function (root) {
        if (!root) return;
        root.removeEntry(opfsName(key), { recursive: false }).catch(function () {});
      });
      if (getActiveProject() === name) setActiveProject(null);
      emit('change', { scope: 'project', deleted: name });
      return true;
    });
  }

  function newProject(name, source, opts) {
    opts = opts || {};
    if (!name) return Promise.reject(new Error('Project name required'));
    if (/[\\/?*[\]:]/.test(name) || name.length > 64) return Promise.reject(new Error('Invalid project name'));
    return collectProjectMaterials(name, source, opts).then(function (materials) {
      var rec = {
        type: 'project', version: 1, name: name,
        materials: materials, savedAt: nowISO(),
        bomProjectList: opts.bomProjectList || name
      };
      return persist(projectKey(name), rec).then(function () {
        emit('change', { scope: 'project', created: name });
        return rec;
      });
    });
  }

  // Source options for new project:
  //   'copy-global'    -> clone of Global Master
  //   'unified-pool'   -> opts.pick list of material names (from getUnifiedPool)
  //   'workbook-only'  -> opts.workbookMaterials (array)
  //   'blank'          -> []
  function collectProjectMaterials(name, source, opts) {
    opts = opts || {};
    if (source === 'blank') return Promise.resolve([]);
    if (source === 'copy-global') return getMaster(null).then(function (m) { return m || []; });
    if (source === 'workbook-only') return Promise.resolve(opts.workbookMaterials || []);
    if (source === 'unified-pool') {
      return getUnifiedPool().then(function (pool) {
        var pick = opts.pick || [];
        var byName = {}; pool.forEach(function (m) { byName[m.name] = m; });
        return pick.map(function (n) { return byName[n]; }).filter(Boolean);
      });
    }
    return Promise.resolve([]);
  }

  // Unified pool: Global Master + current workbook + all saved projects, de-duped by name.
  // `workbookMaterials` is optional; if omitted, only Global + projects are pooled.
  function getUnifiedPool(workbookMaterials) {
    var globalP = getMaster(null).then(function (m) { return m || []; });
    var projectsP = listProjects().then(function (names) {
      return Promise.all(names.map(function (n) {
        return load(projectKey(n)).then(function (rec) { return (rec && rec.materials) || []; });
      })).then(function (arrs) {
        var flat = []; arrs.forEach(function (a) { a.forEach(function (m) { flat.push(m); }); });
        return flat;
      });
    });
    return Promise.all([globalP, projectsP]).then(function (r) {
      var pool = []; var seen = {};
      function add(arr) {
        (arr || []).forEach(function (m) {
          if (m && m.name && !seen[m.name]) { seen[m.name] = 1; pool.push(clone(m)); }
        });
      }
      add(r[0]); add(r[1]); add(workbookMaterials);
      pool.sort(function (a, b) { return a.name.localeCompare(b.name); });
      return pool;
    });
  }

  // ---------- EXPORT / IMPORT ----------
  function exportMaster(projectName) {
    var p = projectName !== undefined ? projectName : getActiveProject();
    return getMaster(p).then(function (materials) {
      return {
        type: p ? 'project' : 'master',
        version: 1,
        name: p || undefined,
        exportedAt: nowISO(),
        materials: materials || []
      };
    });
  }
  function importMaster(jsonOrStr, opts) {
    opts = opts || {};
    var obj = (typeof jsonOrStr === 'string') ? JSON.parse(jsonOrStr) : jsonOrStr;
    if (!obj || obj.type === 'project') {
      throw new Error('Refused: this file is a Project export. Use Import Project .json instead.');
    }
    if (!Array.isArray(obj.materials)) throw new Error('Invalid master file: no materials[]');
    var projectName = opts.projectName !== undefined ? opts.projectName : getActiveProject();
    return saveMaster(obj.materials, projectName);
  }

  function exportProject(name) {
    var n = name || getActiveProject();
    if (!n) return Promise.reject(new Error('No active project to export'));
    return load(projectKey(n)).then(function (rec) {
      if (!rec) throw new Error('Project not found: ' + n);
      var out = clone(rec);
      out.type = 'project'; out.exportedAt = nowISO();
      return out;
    });
  }
  function importProject(jsonOrStr, opts) {
    opts = opts || {};
    var obj = (typeof jsonOrStr === 'string') ? JSON.parse(jsonOrStr) : jsonOrStr;
    if (!obj || obj.type !== 'project') throw new Error('Not a project file (type:"project" required)');
    if (!obj.name) throw new Error('Project file missing name');
    var name = opts.rename || obj.name;
    var rec = {
      type: 'project', version: 1, name: name,
      materials: obj.materials || [],
      sheets: obj.sheets || [],
      bomProjectList: obj.bomProjectList || name,
      savedAt: nowISO(), importedAt: nowISO()
    };
    return persist(projectKey(name), rec).then(function () {
      emit('change', { scope: 'project', imported: name });
      return rec;
    });
  }

  // ---------- OFFLINE LOCAL API (auto-detect) ----------
  var _offline = null;
  function detectOfflineApi() {
    if (_offline !== null) return Promise.resolve(_offline);
    var host = root.location && root.location.hostname;
    if (!host || (host !== 'localhost' && host !== '127.0.0.1')) {
      _offline = { available: false };
      return Promise.resolve(_offline);
    }
    return fetch('/api/health', { cache: 'no-store' }).then(function (r) {
      if (!r.ok) { _offline = { available: false }; return _offline; }
      return r.json().then(function (j) {
        _offline = { available: true, endpoints: j && j.endpoints, version: j && j.version };
        emit('offline', _offline);
        return _offline;
      }, function () { _offline = { available: true }; return _offline; });
    }, function () { _offline = { available: false }; return _offline; });
  }

  function offlineGetMaster() {
    if (!_offline || !_offline.available) return Promise.resolve(null);
    return fetch('/api/master').then(function (r) { return r.ok ? r.json() : null; }, function () { return null; });
  }
  function offlinePutMaster(rec) {
    if (!_offline || !_offline.available) return Promise.resolve(false);
    return fetch('/api/master', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(rec) })
      .then(function (r) { return r.ok; }, function () { return false; });
  }
  function offlineGetProjects() {
    if (!_offline || !_offline.available) return Promise.resolve([]);
    return fetch('/api/projects').then(function (r) { return r.ok ? r.json() : []; }, function () { return []; });
  }

  // ---------- GITHUB (optional, via GitHub Pages raw + REST API) ----------
  // opts: { token, owner, repo, branch, pathMaster, pathProject }
  function githubPull(opts) {
    opts = opts || {};
    var base = opts.rawBase || ('https://raw.githubusercontent.com/' + opts.owner + '/' + opts.repo + '/' + (opts.branch || 'main') + '/');
    var path = opts.path || 'bom/master.json';
    var url = base + path + '?t=' + Date.now();
    return fetch(url).then(function (r) {
      if (!r.ok) throw new Error('GitHub GET failed: ' + r.status);
      return r.json();
    });
  }
  function githubPush(opts, rec) {
    opts = opts || {};
    if (!opts.token || !opts.owner || !opts.repo) return Promise.reject(new Error('githubPush needs token, owner, repo'));
    var path = opts.path || 'bom/master.json';
    var url = 'https://api.github.com/repos/' + opts.owner + '/' + opts.repo + '/contents/' + path;
    // 1. GET current sha (if file exists)
    return fetch(url, { headers: { Authorization: 'Bearer ' + opts.token, Accept: 'application/vnd.github+json' } })
      .then(function (r) { return r.ok ? r.json() : null; }, function () { return null; })
      .then(function (meta) {
        var body = {
          message: 'BOM Tools: push ' + (rec.type || 'master') + ' ' + nowISO(),
          content: b64unicode(JSON.stringify(rec, null, 2)),
          branch: opts.branch || 'main'
        };
        if (meta && meta.sha) body.sha = meta.sha;
        return fetch(url, {
          method: 'PUT',
          headers: { Authorization: 'Bearer ' + opts.token, Accept: 'application/vnd.github+json', 'Content-Type': 'application/json' },
          body: JSON.stringify(body)
        });
      })
      .then(function (r) {
        if (!r.ok) return r.text().then(function (t) { throw new Error('GitHub PUT failed: ' + r.status + ' ' + t); });
        return r.json();
      });
  }
  function b64unicode(str) {
    // UTF-8 safe base64 for browsers
    var bytes = new TextEncoder().encode(str);
    var bin = ''; for (var i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    return root.btoa(bin);
  }

  // ---------- AUDIT LOG (session changes) ----------
  function getAuditLog() { return lsGet(KEY_AUDIT) || []; }
  function recordAudit(entry) {
    var log = getAuditLog();
    entry.at = nowISO();
    log.unshift(entry);
    if (log.length > 200) log.length = 200;
    lsSet(KEY_AUDIT, log);
    emit('audit', entry);
    return entry;
  }
  function clearAuditLog() { lsSet(KEY_AUDIT, []); emit('audit-cleared', {}); }

  // ---------- BOM PROJECT LIST (workbook custom property) ----------
  // Stored locally so the add-in can recall it between sessions; the live
  // workbook property is written/read by BomSheets via Office.js.
  function rememberBomProjectList(name) { lsSet('bom_project_list', name || null); }
  function recallBomProjectList() { return lsGet('bom_project_list') || null; }

  // ---------- STATUS (for UI) ----------
  function status() {
    return Promise.all([
      idbOpen().then(function () { return 'ok'; }, function () { return 'off'; }),
      opfsRoot().then(function (d) { return d ? 'ok' : 'off'; }),
      Promise.resolve(lsKeys('').length > 0 ? 'ok' : 'off'),
      detectOfflineApi().then(function (o) { return o.available ? 'ok' : 'off'; })
    ]).then(function (r) {
      return { indexedDB: r[0], opfs: r[1], localStorage: r[2], offlineApi: r[3],
               activeProject: getActiveProject(), scope: activeScopeLabel() };
    });
  }

  // ---------- PUBLIC API ----------
  root.BomStore = {
    // events
    on: on, off: off, emit: emit,
    // scope
    getActiveProject: getActiveProject,
    setActiveProject: setActiveProject,
    activeScopeLabel: activeScopeLabel,
    // materials
    getMaster: getMaster,
    saveMaster: saveMaster,
    // projects
    listProjects: listProjects,
    listSavedProjects: listSavedProjects,
    deleteProject: deleteProject,
    newProject: newProject,
    getUnifiedPool: getUnifiedPool,
    // export / import
    exportMaster: exportMaster,
    importMaster: importMaster,
    exportProject: exportProject,
    importProject: importProject,
    // offline + github
    detectOfflineApi: detectOfflineApi,
    offlineGetMaster: offlineGetMaster,
    offlinePutMaster: offlinePutMaster,
    offlineGetProjects: offlineGetProjects,
    githubPull: githubPull,
    githubPush: githubPush,
    // audit + workbook project list
    getAuditLog: getAuditLog,
    recordAudit: recordAudit,
    clearAuditLog: clearAuditLog,
    rememberBomProjectList: rememberBomProjectList,
    recallBomProjectList: recallBomProjectList,
    // status
    status: status,
    // constants
    KEY_GLOBAL: KEY_GLOBAL,
    KEY_PROJECT_PREFIX: KEY_PROJECT_PREFIX,
    DEFAULT_MATERIALS: DEFAULT_MATERIALS
  };
})(typeof window !== 'undefined' ? window : this);
