/* taskpane.js - UI controller for BOM Tools task pane */
"use strict";

var $ = function (id) { return document.getElementById(id); };
var state = {
  ok: false,
  canRestore: false,
  busy: false,
  confirmDelete: false,
  convertSubmode: "flatten",
  activeTab: "home"
};

function makeLog(id) {
  return function (msg, clear) {
    var el = $(id);
    if (!el) return;
    if (clear) el.textContent = "";
    if (msg) el.textContent += (el.textContent ? "\n" : "") + msg;
    el.scrollTop = el.scrollHeight;
  };
}

var log = makeLog("log");
var restLog = makeLog("rest-log");
var shLog = makeLog("sh-log");
var matLog = makeLog("mat-log");
var localLog = makeLog("local-log");
var syncLog = makeLog("sync-log");
var homeRepairLog = makeLog("home-repair-log");
var projLog = makeLog("proj-log");
var restProjLog = makeLog("rest-proj-log");
var fixLog = makeLog("fixnames-log");

function setDot(id, kind) {
  var el = $(id);
  if (el) el.className = "dot" + (kind ? " " + kind : "");
}

var REQ_ICON_INFO = '<svg class="req-hint-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>';
var REQ_ICON_ISSUE = '<svg class="req-hint-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>';

function setReqHint(id, msg, isIssue) {
  var el = $(id);
  if (!el) return;
  if (!msg) {
    el.hidden = true;
    el.innerHTML = "";
    el.className = "req-hint";
    return;
  }
  el.hidden = false;
  el.className = "req-hint" + (isIssue ? " issue" : "");
  el.innerHTML = (isIssue ? REQ_ICON_ISSUE : REQ_ICON_INFO) + '<span class="req-hint-text"></span>';
  el.lastChild.textContent = msg;
}

function setFieldIssue(id, hasIssue) {
  var el = $(id);
  if (!el) return;
  if (hasIssue) el.classList.add("field-issue");
  else el.classList.remove("field-issue");
}

function updateButtons() {
  var isCopy = $("iscopy") ? $("iscopy").checked : false;
  var restPicked = getPickedSheets("rest-sheets").length;
  var shRepPicked = getPickedSheets("sh-repair-sheets").length;
  var homeRepPicked = getPickedSheets("home-repair-sheets").length;

  if ($("convert")) $("convert").disabled = state.busy || !state.ok || !isCopy;
  if ($("copy")) $("copy").disabled = state.busy;
  if ($("rest-copy")) $("rest-copy").disabled = state.busy;
  if ($("restore-btn")) $("restore-btn").disabled = state.busy || !state.canRestore || !restPicked;
  if ($("sh-repair")) $("sh-repair").disabled = state.busy || !shRepPicked;
  if ($("home-repair-btn")) $("home-repair-btn").disabled = state.busy || !homeRepPicked;

  if ($("fixnames-btn")) $("fixnames-btn").disabled = state.busy || !state.canRestore;
  updateProjectButtons();

  if (!state.ok) {
    setReqHint("conv-req-hint", "Open an automated BOM workbook to flatten.", false);
  } else if (!isCopy) {
    setReqHint("conv-req-hint", "Tick the copy confirmation box above to enable conversion.", false);
  } else {
    setReqHint("conv-req-hint", "", false);
  }

  if (!state.canRestore) {
    setReqHint("rest-req-hint", "No converted BOM sheets found to restore.", false);
  } else if (!restPicked) {
    setReqHint("rest-req-hint", "Select at least one sheet above to restore.", true);
  } else {
    setReqHint("rest-req-hint", "", false);
  }

  var homeRepBox = $("home-repair-sheets");
  if (homeRepBox && !homeRepBox.hidden && !homeRepPicked) {
    setReqHint("home-repair-hint", "Select at least one BOM sheet to repair.", true);
  } else {
    setReqHint("home-repair-hint", "", false);
  }

  var shRepBox = $("sh-repair-sheets");
  if (shRepBox && !shRepBox.hidden && !shRepPicked) {
    setReqHint("sh-repair-hint", "Select at least one BOM sheet to repair.", true);
  } else {
    setReqHint("sh-repair-hint", "", false);
  }
}

/* Reusable compact checklist for selecting detected sheets (used by Repair & Restore) */
function getPickedSheets(containerId) {
  var el = $(containerId);
  if (!el || el.hidden) return [];
  var boxes = el.querySelectorAll('input[type="checkbox"]:checked');
  var out = [];
  Array.prototype.forEach.call(boxes, function (b) { out.push(b.value); });
  return out;
}

function renderSheetPicker(containerId, items, titleText) {
  var el = $(containerId);
  if (!el) return;
  el.innerHTML = "";
  if (!items || !items.length) {
    el.hidden = true;
    updateButtons();
    return;
  }
  el.hidden = false;

  var hdr = document.createElement("div");
  hdr.className = "sheet-pick-hdr";
  var title = document.createElement("span");
  title.textContent = titleText || ("Select sheets (" + items.length + ")");
  hdr.appendChild(title);

  if (items.length > 1) {
    var toggle = document.createElement("button");
    toggle.type = "button";
    toggle.className = "link-btn";
    toggle.textContent = "None";
    toggle.onclick = function () {
      var boxes = el.querySelectorAll('input[type="checkbox"]');
      var allOn = Array.prototype.every.call(boxes, function (b) { return b.checked; });
      Array.prototype.forEach.call(boxes, function (b) { b.checked = !allOn; });
      toggle.textContent = allOn ? "All" : "None";
      updateButtons();
    };
    hdr.appendChild(toggle);
  }
  el.appendChild(hdr);

  var list = document.createElement("div");
  list.className = "sheet-pick-list";
  items.forEach(function (item) {
    var name = typeof item === "string" ? item : item.name;
    var tag = typeof item === "string" ? "" : (item.tag || "");
    var row = document.createElement("label");
    row.className = "sheet-pick-item";
    var cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = name;
    cb.checked = true;
    cb.onchange = function () {
      var boxes = el.querySelectorAll('input[type="checkbox"]');
      var toggleBtn = hdr.querySelector("button");
      if (toggleBtn) {
        var allOn = Array.prototype.every.call(boxes, function (b) { return b.checked; });
        toggleBtn.textContent = allOn ? "None" : "All";
      }
      updateButtons();
    };
    var txt = document.createElement("span");
    txt.textContent = name;
    row.appendChild(cb);
    row.appendChild(txt);
    if (tag) {
      var sm = document.createElement("small");
      sm.textContent = tag;
      row.appendChild(sm);
    }
    list.appendChild(row);
  });
  el.appendChild(list);
  updateButtons();
}

/* =====================================================================
 * Navigation (Always opens to Home tab)
 * ===================================================================== */
var TABS = ["home", "sheet", "mat", "convert"];

function setConvertSubmode(mode) {
  state.convertSubmode = mode === "restore" ? "restore" : "flatten";
  $("cmode-flatten").setAttribute("aria-pressed", state.convertSubmode === "flatten" ? "true" : "false");
  $("cmode-restore").setAttribute("aria-pressed", state.convertSubmode === "restore" ? "true" : "false");
  $("panel-flatten").hidden = state.convertSubmode !== "flatten";
  $("panel-restore").hidden = state.convertSubmode !== "restore";
}

function showTab(name, opts) {
  opts = opts || {};
  state.activeTab = name;
  TABS.forEach(function (t) {
    $("tab-" + t).setAttribute("aria-selected", t === name ? "true" : "false");
    $("page-" + t).hidden = (t !== name);
  });
  if (name === "home") { refreshHome(); refreshLocalFileStatus(); }
  if (name === "sheet") { refreshSheets(); refreshLocalFileStatus(); }
  if (name === "mat") {
    refreshMaterials();
    refreshLocalFileStatus();
    refreshSync();
    renderSessionChanges();
    refreshProject();
    BomSync.watch(localLog, handleAutoSaveEvent);
  }
  if (name === "convert") {
    refreshProject();
    if (opts.submode) setConvertSubmode(opts.submode);
    check(!opts.submode);
  }
}

/* =====================================================================
 * Home Overview & Quick Options
 * ===================================================================== */
async function refreshHome() {
  var st = $("home-wb-status");
  st.className = "status";
  st.textContent = "Checking workbook...";
  setDot("home-dot", "");
  $("qa-add").classList.remove("highlight");
  $("qa-flatten").classList.remove("highlight");
  $("qa-restore").classList.remove("highlight");

  try {
    var info = await BomConvert.inspectWorkbook();
    var shInfo = null;
    try { shInfo = await BomTools.inspectSheets(); } catch (e) { shInfo = null; }
    var repairItems = (shInfo && shInfo.hasMaster && shInfo.bomSheets.length)
      ? shInfo.bomSheets.map(function (s) { return { name: s.name, tag: s.hidden ? "template" : "" }; })
      : [];
    renderSheetPicker("home-repair-sheets", repairItems, "BOM sheets to repair");
    if ($("home-repair-drawer")) $("home-repair-drawer").hidden = !repairItems.length;

    var hasAuto = info.master && info.bomSheets.length > 0;
    var hasConv = info.convertedSheets && info.convertedSheets.length > 0;

    if (hasAuto && !hasConv) {
      setDot("home-dot", "ok");
      st.className = "status ok";
      st.textContent = "Automated BOM ready · " + info.bomSheets.length + " sheet(s): " + info.bomSheets.join(", ");
      $("qa-add").classList.add("highlight");
    } else if (hasConv && !hasAuto) {
      setDot("home-dot", "info");
      st.className = "status";
      st.textContent = "Converted BOM detected (" + info.convertedSheets.join(", ") + "). You can restore it to automated format.";
      $("qa-restore").classList.add("highlight");
    } else if (hasAuto && hasConv) {
      setDot("home-dot", "ok");
      st.className = "status ok";
      st.textContent = info.bomSheets.length + " automated + " + info.convertedSheets.length + " converted BOM sheet(s) found.";
      $("qa-restore").classList.add("highlight");
    } else {
      setDot("home-dot", "info");
      st.className = "status";
      st.textContent = "Regular workbook (no BOM template yet). Click Add a BOM Sheet to set it up automatically.";
      $("qa-add").classList.add("highlight");
    }
  } catch (e) {
    setDot("home-dot", "warn");
    st.className = "status warn";
    st.textContent = "Could not inspect workbook: " + (e.message || e);
  }
  updateButtons();
}

/* =====================================================================
 * Built-in Computer Storage (Always Auto-Connected) + JSON Backup / Transfer
 * ===================================================================== */
async function refreshLocalFileStatus() {
  var st = $("local-status");
  var tag = $("local-summary-tag");
  var homeSt = $("home-local-status");
  var shSt = $("sh-local-status");
  var pullBtn = $("local-pull"), saveBtn = $("local-save");

  try {
    if (BomSync.ensureSeededFromWorkbook) {
      await BomSync.ensureSeededFromWorkbook();
    }
    var info = await BomSync.getLocalFileStatus();
    if (info.project) {
      if (st) {
        st.className = "status ok";
        st.textContent = "Project list \"" + info.project + "\" auto-saves in its own slot (" + info.cachedCount + " materials). The shared master list is untouched.";
      }
      if (tag) { tag.className = "sum-meta ok"; tag.textContent = "Project list"; }
      if (homeSt) homeSt.textContent = "This workbook uses the project-only list \"" + info.project + "\" (" + info.cachedCount + " items).";
      if (shSt) shSt.textContent = "Project-only list \"" + info.project + "\" in use (" + info.cachedCount + " items).";
    } else if (info.hasCache) {
      if (st) {
        st.className = "status ok";
        st.textContent = "Auto-saving in built-in storage (" + info.cachedCount + " materials) — never asks to reconnect.";
      }
      if (tag) { tag.className = "sum-meta ok"; tag.textContent = info.cachedCount + " saved"; }
      if (homeSt) homeSt.textContent = "Master list auto-saves on this computer (" + info.cachedCount + " items).";
      if (shSt) shSt.textContent = "Auto-connected to built-in storage (" + info.cachedCount + " items).";
    } else {
      if (st) {
        st.className = "status";
        st.textContent = "Built-in storage is ready. Open a BOM workbook or import a JSON backup to populate it.";
      }
      if (tag) { tag.className = "sum-meta"; tag.textContent = "Auto-saved"; }
      if (homeSt) homeSt.textContent = "Master list will auto-save on this computer once a BOM is opened.";
      if (shSt) shSt.textContent = "Built-in storage ready — materials auto-save on this computer.";
    }

    if (pullBtn) pullBtn.disabled = state.busy || !info.hasCache || !mat.hasMaster;
    if (saveBtn) saveBtn.disabled = state.busy || !mat.rows.length;
  } catch (e) {
    if (st) { st.className = "status warn"; st.textContent = "Could not check built-in storage: " + (e.message || e); }
  }
}

async function localPull() {
  state.busy = true; localLog("", true);
  try {
    await BomSync.pullLocal(localLog, { syncRemovals: true });
    var list = await BomSync.workbookList();
    BomSync.updateWatchSnap(list);
  } catch (e) {
    localLog("Reload failed: " + (e.message || e));
  }
  state.busy = false;
  await refreshMaterials();
  await refreshLocalFileStatus();
}

async function localSave() {
  state.busy = true; localLog("", true);
  try {
    await BomSync.saveWorkbookToLocal(localLog, {});
  } catch (e) {
    localLog("Save failed: " + (e.message || e));
  }
  state.busy = false;
  await refreshLocalFileStatus();
}

/* Saves text as a file (Save As dialog where available, otherwise a normal download). */
async function saveTextFile(suggestedName, text, description) {
  if (BomSync.supportsSavePicker && BomSync.supportsSavePicker()) {
    try {
      var handle = await window.showSaveFilePicker({
        suggestedName: suggestedName,
        types: [{ description: description || "JSON", accept: { "application/json": [".json"] } }]
      });
      var writable = await handle.createWritable();
      await writable.write(text);
      await writable.close();
      return { saved: true, cancelled: false, name: handle.name };
    } catch (err) {
      if (err && err.name === "AbortError") return { saved: false, cancelled: true, name: suggestedName };
    }
  }
  var blob = new Blob([text], { type: "application/json;charset=utf-8" });
  var url = URL.createObjectURL(blob);
  var a = document.createElement("a");
  a.href = url; a.download = suggestedName;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  return { saved: true, cancelled: false, name: suggestedName };
}

async function localExportDownload() {
  localLog("", true);
  try {
    var exp = await BomSync.buildExportJson();
    if (!exp.count) {
      localLog("No materials to export yet. Open a BOM workbook or add materials first.");
      return;
    }
    if (BomSync.supportsSavePicker && BomSync.supportsSavePicker()) {
      try {
        var handle = await window.showSaveFilePicker({
          suggestedName: exp.fileName || "master-materials.json",
          types: [{ description: "BOM Master Materials JSON", accept: { "application/json": [".json"] } }]
        });
        var writable = await handle.createWritable();
        await writable.write(exp.json);
        await writable.close();
        localLog("Exported \"" + handle.name + "\" (" + exp.count + " materials) for backup or transfer.");
        return;
      } catch (err) {
        if (err && err.name === "AbortError") return;
      }
    }
    var blob = new Blob([exp.json], { type: "application/json;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = exp.fileName || "master-materials.json";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    localLog("Exported " + (exp.fileName || "master-materials.json") + " (" + exp.count + " materials) for backup or transfer.");
  } catch (e) {
    localLog("Export failed: " + (e.message || e));
  }
}

async function localImportFile(ev) {
  var file = ev.target && ev.target.files && ev.target.files[0];
  if (!file) return;
  state.busy = true; localLog("", true);
  try {
    var text = await file.text();
    var res = await BomSync.importLocalJsonText(text, file.name, localLog);
    if (!res || !res.savedOnly) {
      try {
        var list = await BomSync.workbookList();
        BomSync.updateWatchSnap(list);
      } catch (e2) { /* ignore if workbook has no Master sheet */ }
    }
  } catch (e) {
    localLog("Import failed: " + (e.message || e));
  }
  ev.target.value = "";
  state.busy = false;
  await refreshMaterials();
  await refreshLocalFileStatus();
}

/* =====================================================================
 * Add Sheet
 * ===================================================================== */
var sheetInfo = null;

function setSheetClearDefault() {
  $("sh-clear").checked = $("sh-source").value !== BomTools.BLANK;
}

async function refreshSheets() {
  var st = $("sh-status");
  st.className = "status"; st.textContent = "Checking workbook...";
  setDot("sh-dot", "");
  try {
    sheetInfo = await BomTools.inspectSheets();
    var sel = $("sh-source"), keep = sel.value;
    sel.innerHTML = "";
    if (sheetInfo.needsTemplate) {
      var o0 = document.createElement("option");
      o0.value = BomTools.BLANK; o0.textContent = "BOM Blank (auto-imported)";
      sel.appendChild(o0);
    }
    sheetInfo.bomSheets.forEach(function (s) {
      var o = document.createElement("option");
      o.value = s.name; o.textContent = s.name + (s.hidden ? " (blank template)" : "");
      sel.appendChild(o);
    });
    var names = sheetInfo.bomSheets.map(function (s) { return s.name; });
    sel.value = names.indexOf(keep) >= 0 ? keep : (sheetInfo.hasBlank || sheetInfo.needsTemplate ? BomTools.BLANK : names[0]);
    setSheetClearDefault();
    if (!$("sh-project").value && sheetInfo.project) $("sh-project").value = sheetInfo.project;

    var hasRepairTarget = sheetInfo.hasMaster && names.length > 0;
    if ($("sh-repair-drawer")) $("sh-repair-drawer").hidden = !hasRepairTarget;
    var repairItems = hasRepairTarget
      ? sheetInfo.bomSheets.map(function (s) { return { name: s.name, tag: s.hidden ? "template" : "" }; })
      : [];
    renderSheetPicker("sh-repair-sheets", repairItems, "BOM sheets to repair");

    if (sheetInfo.needsTemplate) {
      setDot("sh-dot", "info");
      st.className = "status";
      st.textContent = "New workbook — template & your saved materials list will load automatically.";
    } else if (names.length) {
      setDot("sh-dot", "ok");
      st.className = "status ok";
      var visibleCount = sheetInfo.bomSheets.filter(function (s) { return !s.hidden; }).length;
      st.textContent = "Ready · " + visibleCount + " active BOM sheet(s) in this workbook.";
    } else {
      setDot("sh-dot", "warn");
      st.className = "status warn";
      st.textContent = "No BOM tables found. If this is a converted workbook, use Convert > Restore Automated first.";
    }
  } catch (e) {
    sheetInfo = null;
    if ($("sh-repair-drawer")) $("sh-repair-drawer").hidden = true;
    setDot("sh-dot", "warn");
    st.className = "status warn";
    st.textContent = "Could not read workbook: " + (e.message || e);
  }
  updateSheetButton();
}

function updateSheetButton() {
  var addBtn = $("sh-add");
  setFieldIssue("sh-name", false);

  if (!sheetInfo || (!sheetInfo.bomSheets.length && !sheetInfo.needsTemplate)) {
    if (addBtn) addBtn.disabled = true;
    setReqHint("sh-req-hint", "Open a BOM workbook or restore converted sheets first.", false);
    updateButtons();
    return;
  }

  var raw = $("sh-name") ? $("sh-name").value : "";
  if (!raw.trim()) {
    if (addBtn) addBtn.disabled = true;
    setReqHint("sh-req-hint", "Enter a sheet name to continue.", false);
    updateButtons();
    return;
  }

  var nameErr = BomTools.checkSheetName(raw, sheetInfo.allSheets);
  if (nameErr) {
    if (addBtn) addBtn.disabled = true;
    setFieldIssue("sh-name", true);
    setReqHint("sh-req-hint", nameErr, true);
    updateButtons();
    return;
  }

  setReqHint("sh-req-hint", "", false);
  if (addBtn) addBtn.disabled = state.busy;
  updateButtons();
}

async function addSheet() {
  updateSheetButton();
  if ($("sh-add").disabled) return;
  state.busy = true; updateSheetButton(); updateButtons();
  shLog("", true);
  try {
    var source = $("sh-source").value, clear = $("sh-clear").checked;
    if (sheetInfo && sheetInfo.needsTemplate) {
      await BomSync.importTemplate(shLog);
      source = BomTools.BLANK; clear = false;
    }
    if ($("sh-sync").checked) {
      try {
        var localRes = await BomSync.pullLocal(shLog, { syncRemovals: true });
        if (localRes && localRes.empty && BomSync.hasToken()) {
          await BomSync.pull(shLog);
        }
      } catch (e) {
        shLog("Note: could not load saved materials (" + (e.message || e) + "). Using workbook list.");
      }
    }
    var r = await BomTools.addBomSheet({
      name: $("sh-name").value, project: $("sh-project").value,
      source: source, clear: clear
    }, shLog);
    shLog("Added sheet \"" + r.name + "\".");
    r.warnings.forEach(function (w) { shLog("WARNING: " + w); });
    $("sh-name").value = "";
  } catch (e) {
    shLog("Could not add sheet: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; updateButtons();
  await refreshSheets();
  await refreshLocalFileStatus();
}

async function repairWorkbook(logger, pickerId) {
  logger = logger || shLog;
  var picked = pickerId ? getPickedSheets(pickerId) : null;
  if (picked && !picked.length) {
    logger("Select at least one BOM sheet to repair.", true);
    return;
  }
  state.busy = true; updateSheetButton(); updateButtons();
  logger("", true);
  try {
    await BomTools.repairWorkbook(logger, picked ? { sheets: picked } : {});
  } catch (e) {
    logger("Repair failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; updateButtons();
  await refreshSheets();
  if (state.activeTab === "home") await refreshHome();
}

/* =====================================================================
 * Materials + Aesthetic Change Notification & Session History
 * ===================================================================== */
var TOAST_DURATION_MS = 5500;
var mat = {
  rows: [],
  hasMaster: false,
  mode: "add",
  selected: null,
  sessionLog: [],
  recentMap: {},
  toastTimer: null,
  toastRemaining: 0,
  toastStartedAt: 0,
  toastVisible: false,
  sessionOpen: false
};

var FIELD_LABELS = {
  category: "Category",
  unit: "Primary Unit",
  lin: "Linear Weight",
  area: "Area Weight",
  piece: "Piece Weight",
  sl: "Sheet Length",
  sw: "Sheet Width",
  remarks: "Remarks"
};

var TOAST_ICONS = {
  added: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>',
  modified: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/></svg>',
  removed: '<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><polyline points="3 6 5 6 21 6"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/><path d="M10 11v6"/><path d="M14 11v6"/></svg>'
};

function normalizeChangeType(t) {
  if (t === "added") return "added";
  if (t === "removed") return "removed";
  return "modified";
}

function formatClock(iso) {
  try {
    var d = iso ? new Date(iso) : new Date();
    return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  } catch (e) { return ""; }
}

function summarizeMaterialInfo(m) {
  if (!m) return "";
  var parts = [];
  if (m.unit === "kg/m" && fmt(m.lin) !== "") parts.push(fmt(m.lin) + " kg/m");
  else if (m.unit === "kg/m²") {
    var a = fmt(m.area) !== "" ? (fmt(m.area) + " kg/m²") : "kg/m²";
    if (fmt(m.sl) !== "" && fmt(m.sw) !== "") a += " (" + fmt(m.sl) + "×" + fmt(m.sw) + " mm)";
    parts.push(a);
  } else if (m.unit === "kg/each" && fmt(m.piece) !== "") parts.push(fmt(m.piece) + " kg/each");
  else if (m.unit) parts.push(m.unit);

  if (m.category) parts.push(m.category);
  if (m.remarks) parts.push(m.remarks);
  return parts.join(" · ") || "No weight or details yet";
}

function updateMoreOptionsMeta() {
  var meta = $("m-more-meta");
  if (!meta) return;
  var parts = [];
  var cat = $("m-cat") ? $("m-cat").value.trim() : "";
  var rem = $("m-rem") ? $("m-rem").value.trim() : "";
  if (cat) parts.push(cat);
  if (rem) parts.push("Remarks");
  meta.textContent = parts.length ? parts.join(" · ") : "Category, remarks";
  meta.className = "sum-meta" + (parts.length ? " ok" : "");
}

function openMaterialInEditor(materialName) {
  var found = mat.rows.filter(function (r) {
    return r.material.toLowerCase() === String(materialName || "").toLowerCase();
  })[0];
  if (found) {
    setMatMode("edit", found);
    matLog("", true);
  }
}

function clearToastTimer() {
  if (mat.toastTimer) {
    clearTimeout(mat.toastTimer);
    mat.toastTimer = null;
  }
}

function collapseToast(immediate) {
  clearToastTimer();
  var toast = $("mat-change-toast");
  if (!toast || toast.hidden) {
    mat.toastVisible = false;
    renderSessionChanges();
    return;
  }
  if (immediate) {
    toast.classList.remove("closing");
    toast.hidden = true;
    mat.toastVisible = false;
    renderSessionChanges();
    return;
  }
  toast.classList.add("closing");
  setTimeout(function () {
    toast.classList.remove("closing");
    toast.hidden = true;
    mat.toastVisible = false;
    renderSessionChanges();
  }, 220);
}

function scheduleToastCollapse(ms) {
  clearToastTimer();
  mat.toastRemaining = ms;
  mat.toastStartedAt = Date.now();
  mat.toastTimer = setTimeout(function () {
    collapseToast(false);
  }, ms);
}

function pauseToastTimer() {
  if (!mat.toastVisible || !mat.toastTimer) return;
  var elapsed = Date.now() - mat.toastStartedAt;
  mat.toastRemaining = Math.max(400, mat.toastRemaining - elapsed);
  clearToastTimer();
}

function resumeToastTimer() {
  if (!mat.toastVisible || mat.toastTimer) return;
  scheduleToastCollapse(mat.toastRemaining || 2000);
}

function showChangeToast(evt) {
  var toast = $("mat-change-toast");
  var iconEl = $("mat-toast-icon");
  var titleEl = $("mat-toast-title");
  var subEl = $("mat-toast-sub");
  var itemsEl = $("mat-toast-items");
  var whereEl = $("mat-toast-where");
  var hintEl = $("mat-toast-hint");
  var barEl = $("mat-toast-bar");
  if (!toast || !itemsEl) return;

  var items = evt.items || [];
  if (!items.length) return;

  var firstType = normalizeChangeType(items[0].type);
  var allSame = items.every(function (it) { return normalizeChangeType(it.type) === firstType; });
  var kind = allSame ? firstType : "added";

  toast.className = "change-toast kind-" + kind;
  if (iconEl) iconEl.innerHTML = TOAST_ICONS[kind] || TOAST_ICONS.added;

  if (titleEl) {
    if (items.length === 1) {
      titleEl.textContent = firstType === "added" ? "Material Added"
        : firstType === "removed" ? "Material Removed" : "Material Modified";
    } else {
      titleEl.textContent = items.length + " Changes Saved";
    }
  }
  if (subEl) {
    var srcLabel = evt.source === "sheet" ? "Detected on Master sheet" : "Saved from Materials panel";
    var clk = formatClock(evt.time);
    subEl.textContent = srcLabel + (clk ? " · " + clk : "");
  }

  itemsEl.innerHTML = "";
  var anyClickable = false;
  items.slice(0, 3).forEach(function (it) {
    var m = it.material || {};
    var cType = normalizeChangeType(it.type);
    var clickable = cType !== "removed";
    if (clickable) anyClickable = true;

    var box = document.createElement("div");
    box.className = "toast-item" + (clickable ? " clickable" : "");

    var top = document.createElement("div");
    top.className = "toast-item-row";

    var nm = document.createElement("span");
    nm.className = "toast-name";
    nm.textContent = m.material || "(unnamed)";

    var badge = document.createElement("span");
    badge.className = "toast-badge " + cType;
    badge.textContent = cType === "added" ? "Added" : cType === "removed" ? "Removed" : "Modified";

    top.appendChild(nm);
    top.appendChild(badge);
    box.appendChild(top);

    var infoLine = document.createElement("div");
    infoLine.className = "toast-info";
    infoLine.textContent = summarizeMaterialInfo(m);
    box.appendChild(infoLine);

    if (cType === "modified" && it.fields && it.fields.length) {
      var diffLine = document.createElement("div");
      diffLine.className = "toast-diff";
      diffLine.textContent = "Changed: " + it.fields.map(function (k) { return FIELD_LABELS[k] || k; }).join(", ");
      box.appendChild(diffLine);
    }

    if (clickable) {
      box.title = "Click to edit in panel";
      box.onclick = function () { openMaterialInEditor(m.material); };
    }
    itemsEl.appendChild(box);
  });

  if (items.length > 3) {
    var extra = document.createElement("div");
    extra.className = "toast-info";
    extra.textContent = "+ " + (items.length - 3) + " more item(s) — view all in Session changes";
    itemsEl.appendChild(extra);
  }

  if (whereEl) {
    whereEl.textContent = "Saved to Master sheet & this computer";
  }
  if (hintEl) hintEl.textContent = anyClickable ? "Click to edit" : "";

  // Restart CSS progress bar animation cleanly
  if (barEl) {
    barEl.style.animation = "none";
    void barEl.offsetWidth;
    barEl.style.animation = "";
  }

  mat.toastVisible = true;
  mat.sessionOpen = false;
  toast.hidden = false;
  renderSessionChanges();
  scheduleToastCollapse(TOAST_DURATION_MS);
}

function renderSessionChanges() {
  var wrap = $("mat-session-wrap");
  var btn = $("mat-changes-btn");
  var panel = $("mat-changes-panel");
  var listEl = $("mat-changes-list");
  var countEl = $("mat-session-count");
  var previewEl = $("mat-session-preview");
  var dotEl = $("mat-session-dot");
  if (!wrap || !btn || !panel || !listEl) return;

  var count = mat.sessionLog.length;
  if (!count || mat.toastVisible) {
    wrap.hidden = true;
    panel.hidden = true;
    btn.setAttribute("aria-expanded", "false");
    return;
  }

  wrap.hidden = false;
  if (countEl) countEl.textContent = String(count);

  var latest = mat.sessionLog[0];
  if (dotEl) dotEl.className = "session-dot " + latest.type;
  if (previewEl) {
    var verb = latest.type === "added" ? "Added" : latest.type === "removed" ? "Removed" : "Modified";
    previewEl.textContent = verb + " " + (latest.material && latest.material.material ? latest.material.material : "");
  }

  panel.hidden = !mat.sessionOpen;
  btn.setAttribute("aria-expanded", mat.sessionOpen ? "true" : "false");

  if (!mat.sessionOpen) return;

  listEl.innerHTML = "";
  mat.sessionLog.forEach(function (entry) {
    var m = entry.material || {};
    var clickable = entry.type !== "removed";
    var row = document.createElement("div");
    row.className = "session-row" + (clickable ? " clickable" : "");

    var top = document.createElement("div");
    top.className = "session-row-top";

    var nm = document.createElement("span");
    nm.className = "toast-name";
    nm.textContent = m.material || "(unnamed)";

    var badge = document.createElement("span");
    badge.className = "toast-badge " + entry.type;
    badge.textContent = entry.type === "added" ? "Added" : entry.type === "removed" ? "Removed" : "Modified";

    top.appendChild(nm);
    top.appendChild(badge);
    row.appendChild(top);

    var info = document.createElement("div");
    info.className = "toast-info";
    info.textContent = summarizeMaterialInfo(m);
    row.appendChild(info);

    if (entry.type === "modified" && entry.fields && entry.fields.length) {
      var diff = document.createElement("div");
      diff.className = "toast-diff";
      diff.textContent = "Changed: " + entry.fields.map(function (k) { return FIELD_LABELS[k] || k; }).join(", ");
      row.appendChild(diff);
    }

    var meta = document.createElement("div");
    meta.className = "session-row-time";
    var srcText = entry.source === "sheet" ? "Master sheet" : "Materials panel";
    meta.textContent = formatClock(entry.time) + " · " + srcText;
    row.appendChild(meta);

    if (clickable) {
      row.title = "Click to edit in panel";
      row.onclick = function () {
        mat.sessionOpen = false;
        renderSessionChanges();
        openMaterialInEditor(m.material);
      };
    }
    listEl.appendChild(row);
  });
}

function recordAutoSave(evt) {
  if (!evt || !evt.items || !evt.items.length) return;
  var nowIso = evt.time || new Date().toISOString();
  var normalizedItems = evt.items.map(function (it) {
    return {
      type: normalizeChangeType(it.type),
      material: it.material || {},
      fields: it.fields || [],
      time: it.time || nowIso,
      source: evt.source || "panel"
    };
  });

  // Prepend to session log (newest first, cap at 50 entries)
  for (var i = normalizedItems.length - 1; i >= 0; i--) {
    mat.sessionLog.unshift(normalizedItems[i]);
  }
  if (mat.sessionLog.length > 50) mat.sessionLog.length = 50;

  // Track recent items for top-bubbling in the Master Materials list
  var map = {};
  var order = 0;
  mat.sessionLog.forEach(function (entry) {
    var name = entry.material && entry.material.material;
    if (!name) return;
    var key = name.toLowerCase();
    if (map[key]) return;
    if (entry.type === "removed") {
      map[key] = { removed: true };
    } else {
      map[key] = { type: entry.type, order: order++ };
    }
  });
  Object.keys(map).forEach(function (k) { if (map[k].removed) delete map[k]; });
  mat.recentMap = map;

  var normalizedEvt = {
    source: evt.source || "panel",
    time: nowIso,
    savedToBuiltIn: true,
    items: normalizedItems
  };

  if (state.activeTab === "mat") {
    showChangeToast(normalizedEvt);
  } else {
    mat.toastVisible = false;
    renderSessionChanges();
  }
}

function handleAutoSaveEvent(evt) {
  if (evt && evt.items && evt.items.length) recordAutoSave(evt);
  refreshMaterials(mat.mode === "edit" && mat.selected ? mat.selected.material : null);
  refreshLocalFileStatus();
}

function unitFields() {
  var u = $("m-unit").value;
  var show = { "f-lin": u === "kg/m", "f-area": u === "kg/m²", "f-piece": u === "kg/each" };
  Object.keys(show).forEach(function (cls) {
    Array.prototype.forEach.call(document.getElementsByClassName(cls), function (el) { el.hidden = !show[cls]; });
  });
}

function fmt(v) { return (v === "" || v === null || v === undefined) ? "" : String(v); }

function resetDeleteConfirm() {
  state.confirmDelete = false;
  if ($("m-delete")) $("m-delete").textContent = "Remove";
}

function updateMatValidation() {
  var saveBtn = $("m-save");
  ["m-name", "m-unit", "m-lin", "m-area", "m-sl", "m-sw", "m-piece"].forEach(function (id) {
    setFieldIssue(id, false);
  });

  if (!mat.hasMaster) {
    if (saveBtn) saveBtn.disabled = true;
    setReqHint("mat-req-hint", "Open a workbook with a Master sheet (or add a BOM sheet first).", false);
    return { ready: false };
  }

  var missing = [];
  var issues = [];

  var rawName = $("m-name") ? $("m-name").value.trim() : "";
  if (!rawName) {
    missing.push("Material name");
  } else if (rawName.charAt(0) === "=") {
    issues.push("Material name cannot start with '='.");
    setFieldIssue("m-name", true);
  } else if (mat.mode === "add") {
    var exists = mat.rows.some(function (r) {
      return r.material.toLowerCase() === rawName.toLowerCase();
    });
    if (exists) {
      issues.push("\"" + rawName + "\" already exists. Click it above to edit.");
      setFieldIssue("m-name", true);
    }
  }

  var u = $("m-unit") ? $("m-unit").value : "";
  var checkPosField = function (id, label) {
    var val = $(id) ? $(id).value.trim() : "";
    if (!val) {
      missing.push(label);
      return;
    }
    var n = Number(val.replace(/,/g, ""));
    if (!isFinite(n) || n <= 0) {
      issues.push(label + " must be greater than 0.");
      setFieldIssue(id, true);
    }
  };

  if (!u) {
    missing.push("Primary unit");
  } else if (u === "kg/m") {
    checkPosField("m-lin", "Linear weight (kg/m)");
  } else if (u === "kg/each") {
    checkPosField("m-piece", "Piece weight (kg/each)");
  } else if (u === "kg/m²") {
    checkPosField("m-area", "Area weight (kg/m²)");
    checkPosField("m-sl", "Sheet length (mm)");
    checkPosField("m-sw", "Sheet width (mm)");
  }

  if (issues.length) {
    if (saveBtn) saveBtn.disabled = true;
    setReqHint("mat-req-hint", issues[0], true);
    return { ready: false };
  }
  if (missing.length) {
    if (saveBtn) saveBtn.disabled = true;
    var prefix = mat.mode === "edit" ? "Required to save: " : "Required to add: ";
    setReqHint("mat-req-hint", prefix + missing.join(", "), false);
    return { ready: false };
  }

  setReqHint("mat-req-hint", "", false);
  if (saveBtn) saveBtn.disabled = state.busy;
  return { ready: true };
}

function setMatMode(mode, row) {
  mat.mode = mode; mat.selected = row || null;
  resetDeleteConfirm();
  $("mat-mode").textContent = mode === "edit" ? "Edit Material" : "Add New Material";
  $("m-save").textContent = mode === "edit" ? "Save Changes" : "Add Material";
  $("m-name").readOnly = (mode === "edit");
  $("mat-name-label").textContent = mode === "edit" ? "Material name (locked)" : "Material name";
  if ($("m-delete")) $("m-delete").hidden = (mode !== "edit");
  if ($("m-new")) $("m-new").hidden = (mode !== "edit");
  var r = row || { material: "", category: "", unit: "", lin: "", area: "", piece: "", sl: "", sw: "", remarks: "" };
  $("m-name").value = r.material; $("m-cat").value = r.category; $("m-unit").value = r.unit;
  $("m-lin").value = fmt(r.lin); $("m-area").value = fmt(r.area); $("m-piece").value = fmt(r.piece);
  $("m-sl").value = fmt(r.sl); $("m-sw").value = fmt(r.sw); $("m-rem").value = r.remarks;
  if ($("m-more") && mode === "add") $("m-more").open = false;
  unitFields();
  updateMoreOptionsMeta();
  renderMatList();
  updateMatValidation();
}

function renderMatList() {
  var box = $("mat-list"), q = $("mat-search").value.trim().toLowerCase();
  box.innerHTML = "";
  var shown = mat.rows.filter(function (r) { return !q || (r.material + " " + r.category).toLowerCase().indexOf(q) >= 0; });

  // Bubble recently changed materials to the very top of the list
  var rec = mat.recentMap || {};
  if (Object.keys(rec).length) {
    shown = shown.slice().sort(function (a, b) {
      var ra = rec[a.material.toLowerCase()];
      var rb = rec[b.material.toLowerCase()];
      if (ra && !rb) return -1;
      if (!ra && rb) return 1;
      if (ra && rb) return ra.order - rb.order;
      return a.row - b.row;
    });
  }

  shown.slice(0, 150).forEach(function (r) {
    var recInfo = rec[r.material.toLowerCase()];
    var d = document.createElement("div");
    d.className = "mrow" + (mat.selected && mat.selected.row === r.row ? " sel" : "") + (recInfo ? " recent" : "");

    var left = document.createElement("span");
    left.className = "mleft";

    var v = BomTools.getMaterialValidationStatus(r);
    var icon = document.createElement("span");
    icon.className = "vicon " + (v.valid ? "ok" : "inc");
    icon.textContent = v.valid ? "✓" : "!";
    icon.title = v.valid ? "Ready for BOM (Name, Unit, Weight)" : ("Missing: " + v.missing.join(", "));

    var a = document.createElement("span");
    a.className = "mname";
    a.textContent = r.material;
    left.appendChild(icon);
    left.appendChild(a);

    if (recInfo) {
      var tag = document.createElement("span");
      tag.className = "mtag" + (recInfo.type === "modified" ? " modified" : "");
      tag.textContent = recInfo.type === "added" ? "Added" : "Modified";
      left.appendChild(tag);
    }

    var b = document.createElement("small");
    var w = r.unit === "kg/m" ? r.lin : r.unit === "kg/m²" ? r.area : r.piece;
    b.textContent = fmt(w) + " " + r.unit;
    d.appendChild(left); d.appendChild(b);
    d.onclick = function () { setMatMode("edit", r); matLog("", true); };
    box.appendChild(d);
  });
  if (!shown.length) {
    var e = document.createElement("div"); e.className = "mrow"; e.textContent = "No materials match."; e.style.cursor = "default"; box.appendChild(e);
  }
}

async function refreshMaterials(keepName) {
  var st = $("mat-status");
  st.textContent = "Loading...";
  try {
    var res = await BomTools.listMaterials();
    mat.rows = res.rows;
    mat.hasMaster = true;
    var dl = $("m-cats"); dl.innerHTML = "";
    res.categories.forEach(function (c) { var o = document.createElement("option"); o.value = c; dl.appendChild(o); });
    var incompleteCount = res.rows.filter(function (r) { return !BomTools.getMaterialValidationStatus(r).valid; }).length;
    st.textContent = res.rows.length + " items" + (incompleteCount ? (" · " + incompleteCount + " incomplete") : " · all ready") +
      (state.project && state.project.active ? " · project list" : "");
    renderProjectPicker();
    var again = keepName && res.rows.filter(function (r) { return r.material.toLowerCase() === keepName.toLowerCase(); })[0];
    if (again) setMatMode("edit", again);
    else if (mat.mode === "edit" && mat.selected) setMatMode("add");
    else { renderMatList(); updateMatValidation(); }
  } catch (e) {
    mat.rows = [];
    mat.hasMaster = false;
    st.textContent = "No Master sheet in workbook";
    renderMatList();
    updateMatValidation();
  }
}

function readMatForm() {
  return { material: $("m-name").value, category: $("m-cat").value, unit: $("m-unit").value,
           lin: $("m-lin").value, area: $("m-area").value, piece: $("m-piece").value,
           sl: $("m-sl").value, sw: $("m-sw").value, remarks: $("m-rem").value };
}

async function saveMaterial() {
  if (!updateMatValidation().ready) return;
  var raw = readMatForm();
  matLog("", true);
  var norm = BomTools.normalizeMaterial(raw);
  var pre = BomTools.validateMaterial(norm,
    mat.rows.map(function (r) { return r.material; }), mat.mode);
  if (pre.length) {
    setReqHint("mat-req-hint", pre[0], true);
    return;
  }
  state.busy = true; $("m-save").disabled = true;
  try {
    var wasEdit = mat.mode === "edit";
    var prevObj = wasEdit && mat.selected ? BomTools.normalizeMaterial(mat.selected) : null;
    BomSync.suppressWatch(3500);
    var r = wasEdit ? await BomTools.updateMaterial(raw) : await BomTools.addMaterial(raw);
    var nowIso = r.lastModified || new Date().toISOString();
    var changeType = wasEdit ? "modified" : "added";
    var diffFields = prevObj ? BomSync.changedFields(prevObj, r) : [];
    var changeItems = [{ type: changeType, material: r, fields: diffFields, time: nowIso }];
    var lastChange = BomSync.buildLastChangeMeta({ time: nowIso, source: "panel", items: changeItems });

    var list = await BomSync.workbookList();
    list.forEach(function (item) {
      if (item.material.toLowerCase() === r.material.toLowerCase()) {
        item.lastAction = changeType;
        item.lastModified = nowIso;
      }
    });
    BomSync.updateWatchSnap(list);
    await BomSync.writeLocalMaster(list, matLog, { lastChange: lastChange });

    recordAutoSave({
      source: "panel",
      time: nowIso,
      savedToBuiltIn: true,
      items: changeItems
    });
    if (BomSync.hasToken()) {
      try { await BomSync.push(matLog, { only: r.material }); }
      catch (e) { matLog("Saved on this computer; GitHub publish skipped (" + (e.message || e) + ")."); }
    }
    var keep = wasEdit ? r.material : null;
    state.busy = false;
    await refreshMaterials(keep);
    await refreshLocalFileStatus();
    if (!wasEdit) setMatMode("add");
  } catch (e) {
    matLog("Failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false;
  updateMatValidation();
}

async function removeMaterial() {
  if (mat.mode !== "edit" || !mat.selected) return;
  if (!state.confirmDelete) {
    state.confirmDelete = true;
    $("m-delete").textContent = "Confirm";
    matLog("Click \"Confirm\" to permanently remove \"" + mat.selected.material + "\" from this workbook and your saved list.", true);
    return;
  }
  var removedObj = BomTools.normalizeMaterial(mat.selected);
  var target = mat.selected.material;
  state.busy = true; $("m-save").disabled = true; $("m-delete").disabled = true;
  matLog("", true);
  try {
    BomSync.suppressWatch(3500);
    await BomTools.deleteMaterial(target);
    var nowIso = new Date().toISOString();
    var changeItems = [{ type: "removed", material: removedObj, fields: [], time: nowIso }];
    var lastChange = BomSync.buildLastChangeMeta({ time: nowIso, source: "panel", items: changeItems });
    var list = await BomSync.workbookList();
    BomSync.updateWatchSnap(list);
    await BomSync.writeLocalMaster(list, matLog, { lastChange: lastChange });
    recordAutoSave({
      source: "panel",
      time: nowIso,
      savedToBuiltIn: true,
      items: changeItems
    });
    state.busy = false;
    setMatMode("add");
    await refreshMaterials();
    await refreshLocalFileStatus();
  } catch (e) {
    matLog("Remove failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false;
  if ($("m-delete")) $("m-delete").disabled = false;
  resetDeleteConfirm();
  updateMatValidation();
}

/* =====================================================================
 * Convert (Flatten <-> Restore Automated)
 * ===================================================================== */
async function check(autoPickSubmode) {
  var st = $("status"), rst = $("rest-status");
  st.className = "status"; st.textContent = "Checking...";
  rst.className = "status"; rst.textContent = "Checking...";
  setDot("conv-dot", ""); setDot("rest-dot", "");
  try {
    var info = await BomConvert.inspectWorkbook();
    state.ok = info.master && info.bomSheets.length > 0;
    state.canRestore = info.convertedSheets && info.convertedSheets.length > 0;

    if (state.ok) {
      setDot("conv-dot", "ok");
      st.className = "status ok";
      st.textContent = "Ready · " + info.bomSheets.length + " automated BOM sheet(s): " + info.bomSheets.join(", ");
    } else {
      setDot("conv-dot", "warn");
      st.className = "status warn";
      st.textContent = state.canRestore
        ? "This workbook looks already converted. Switch to \"Restore Automated\" above to convert it back."
        : "No automated BOM sheets found in this workbook.";
    }

    if (state.canRestore) {
      setDot("rest-dot", "ok");
      rst.className = "status ok";
      rst.textContent = "Ready · " + info.convertedSheets.length + " converted sheet(s) found.";
      renderSheetPicker("rest-sheets", info.convertedSheets, "Converted sheets to restore");
    } else {
      setDot("rest-dot", "info");
      rst.className = "status";
      rst.textContent = state.ok
        ? "All BOM sheets in this workbook are already automated."
        : "No converted BOM sheets found to restore.";
      renderSheetPicker("rest-sheets", [], "");
    }

    if (autoPickSubmode && !state.ok && state.canRestore) {
      setConvertSubmode("restore");
    }
  } catch (e) {
    state.ok = false; state.canRestore = false;
    setDot("conv-dot", "warn"); setDot("rest-dot", "warn");
    st.className = "status warn"; st.textContent = "Could not read workbook: " + (e.message || e);
    rst.className = "status warn"; rst.textContent = "Could not read workbook: " + (e.message || e);
    renderSheetPicker("rest-sheets", [], "");
  }
  updateButtons();
}

function getFileBase64() {
  return new Promise(function (resolve, reject) {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, function (r) {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(r.error);
      var file = r.value, count = file.sliceCount, slices = [], i = 0;
      (function next() {
        file.getSliceAsync(i, function (s) {
          if (s.status !== Office.AsyncResultStatus.Succeeded) { file.closeAsync(); return reject(s.error); }
          slices.push(s.value.data);
          if (++i < count) return next();
          file.closeAsync();
          var total = 0; slices.forEach(function (a) { total += a.length; });
          var u = new Uint8Array(total), o = 0;
          slices.forEach(function (a) { u.set(a, o); o += a.length; });
          var bin = "";
          for (var k = 0; k < u.length; k += 0x8000) bin += String.fromCharCode.apply(null, u.subarray(k, k + 0x8000));
          resolve(btoa(bin));
        });
      })();
    });
  });
}

async function openCopy(logger, tipText) {
  logger = logger || log;
  state.busy = true; updateButtons();
  logger("Reading this workbook...", true);
  try {
    var b64 = await getFileBase64();
    await Excel.run(async function (context) {
      Excel.createWorkbook(b64).open();
      await context.sync();
    });
    logger(tipText || "Opened a copy in a new window. Open BOM Tools there to convert it safely.");
  } catch (e) {
    logger("Could not open a copy automatically: " + (e.message || JSON.stringify(e)));
    logger("Tip: use File > Save As a new name, then convert that copy.");
  }
  state.busy = false; updateButtons();
}

async function convert() {
  var mode = document.querySelector('input[name="mode"]:checked').value;
  state.busy = true; updateButtons();
  log("", true);
  try {
    try { await BomSync.loadProjectState(); } catch (e0) { /* ignore */ }
    if (BomSync.getProject().active && $("conv-save-project").checked) {
      try {
        var pexp = await BomSync.buildProjectExport();
        var psv = await saveTextFile(pexp.fileName, pexp.json, "BOM project materials JSON");
        if (psv.cancelled) log("Project list file was not saved (cancelled). It is still kept in this computer's built-in storage.");
        else log("Saved project list \"" + pexp.name + "\" (" + pexp.count + " materials) as " + psv.name + ".");
      } catch (e1) {
        log("Could not save the project list file (" + (e1.message || e1) + "). It is still kept in built-in storage. Continuing ...");
      }
    }
    var warnings = await BomConvert.convertWorkbook(mode, log, { compiled: $("compiled").checked });
    log("Done! Review the sheets, then use File > Save As to save.");
    warnings.forEach(function (w) { log("WARNING: " + w); });
  } catch (e) {
    log("Conversion failed: " + (e.message || JSON.stringify(e)));
    log("Close without saving and try again on a fresh copy.");
  }
  state.busy = false;
  await check(false);
}

async function restoreAutomated() {
  var picked = getPickedSheets("rest-sheets");
  if (!picked.length) {
    restLog("Select at least one converted sheet to restore.", true);
    return;
  }
  state.busy = true; updateButtons();
  restLog("", true);
  try {
    var res = await BomConvert.restoreAutomatedWorkbook(restLog, {
      sheets: picked,
      syncLocal: $("rest-sync").checked,
      removeCompiled: $("rest-drop-compiled").checked
    });
    restLog("Done! Restored " + res.restoredCount + " sheet(s) back to automated BOM format.");
    res.warnings.forEach(function (w) { restLog("NOTE: " + w); });
  } catch (e) {
    restLog("Restore failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false;
  await check(false);
  await refreshHome();
}

/* =====================================================================
 * Project-only materials list (one-time materials for a single job)
 * ===================================================================== */
function projectCopyMode() {
  var r = document.querySelector('input[name="projcopy"]:checked');
  return r ? r.value : "all";
}

function renderProjectPicker() {
  var active = state.project && state.project.active;
  if (!active && projectCopyMode() === "pick" && mat.rows.length) {
    renderSheetPicker("proj-pick", mat.rows.map(function (r) { return { name: r.material, tag: r.category }; }),
      "Materials to copy (" + mat.rows.length + ")");
  } else {
    renderSheetPicker("proj-pick", [], "");
  }
}

function updateProjectButtons() {
  var start = $("proj-start");
  if (!start) return;
  var active = state.project && state.project.active;
  var name = ($("proj-name").value || "").trim();
  var pick = projectCopyMode() === "pick";
  var picked = getPickedSheets("proj-pick").length;
  start.disabled = state.busy || active || !mat.hasMaster || !name || (pick && !picked);
  if ($("proj-save-btn")) $("proj-save-btn").disabled = state.busy || !active;
  if ($("proj-stop")) $("proj-stop").disabled = state.busy || !active;
  if ($("proj-load-btn")) $("proj-load-btn").disabled = state.busy;
  if ($("rest-proj-load-btn")) $("rest-proj-load-btn").disabled = state.busy;

  var drawerOff = $("proj-off") && $("proj-off").hidden;
  if (drawerOff) { setReqHint("proj-req-hint", "", false); return; }
  if (!mat.hasMaster) setReqHint("proj-req-hint", "Open an automated BOM workbook (with the Master Unique Materials sheet) to start a project list.", false);
  else if (!name) setReqHint("proj-req-hint", "Type a name for the project list.", false);
  else if (pick && !picked) setReqHint("proj-req-hint", "Select at least one material to copy.", true);
  else setReqHint("proj-req-hint", "", false);
}

async function refreshProject() {
  var info;
  try { info = await BomSync.loadProjectState(); } catch (e) { info = { active: false, name: "" }; }
  state.project = info;
  var st = $("proj-status"), tag = $("proj-summary-tag");
  if (!st) return;
  $("proj-off").hidden = info.active;
  $("proj-on").hidden = !info.active;
  $("conv-proj-row").hidden = !info.active;
  if (info.active) {
    st.className = "status ok";
    st.textContent = "Project list \"" + info.name + "\" is active. Edits stay in this workbook and its own saved copy; the shared master list is untouched.";
    tag.className = "sum-meta ok"; tag.textContent = "Active";
  } else {
    st.className = "status";
    st.textContent = "Using the shared master list.";
    tag.className = "sum-meta"; tag.textContent = "Off";
  }

  // Restore tab: say which list will be used and whether a saved copy exists on this computer
  var rs = $("rest-proj-status");
  if (rs) {
    if (info.active) {
      var ls = null;
      try { ls = await BomSync.getLocalFileStatus(); } catch (e) { ls = null; }
      if (ls && ls.hasCache) {
        rs.className = "status ok";
        rs.textContent = "Project list \"" + info.name + "\" will be used (" + ls.cachedCount + " materials saved on this computer).";
      } else {
        rs.className = "status warn";
        rs.textContent = "Project list \"" + info.name + "\" has no saved copy on this computer. Load its file below before restoring.";
      }
    } else {
      rs.className = "status";
      rs.textContent = "Shared master list in use. Load a project list file to use that instead.";
    }
  }
  renderProjectPicker();
  updateButtons();
}

async function projectStart() {
  updateProjectButtons();
  if ($("proj-start").disabled) return;
  var name = $("proj-name").value.trim();
  var keep = projectCopyMode() === "pick" ? getPickedSheets("proj-pick") : null;
  if (keep && keep.length >= mat.rows.length) keep = null;      // everything picked = full copy
  state.busy = true; updateButtons(); projLog("", true);
  try {
    if (keep) projLog("Removing " + (mat.rows.length - keep.length) + " unpicked material(s) from this workbook's list only ...");
    var r = await BomSync.startProject(name, { keep: keep });
    projLog("Project list \"" + r.name + "\" started with " + r.count + " materials. From now on changes stay in this workbook and its own saved copy.");
    projLog("Tip: click \"Save project list to file\" to keep a copy you can reload later.");
  } catch (e) {
    projLog("Could not start the project list: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false;
  await refreshMaterials();
  await refreshProject();
  await refreshLocalFileStatus();
}

async function projectSave() {
  state.busy = true; updateButtons(); projLog("", true);
  try {
    var exp = await BomSync.buildProjectExport();
    var res = await saveTextFile(exp.fileName, exp.json, "BOM project materials JSON");
    projLog(res.cancelled ? "Save cancelled." : "Saved \"" + res.name + "\" (" + exp.count + " materials).");
  } catch (e) {
    projLog("Could not save the project list: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; updateButtons();
}

async function projectLoadFile(ev, logger) {
  var file = ev.target && ev.target.files && ev.target.files[0];
  if (!file) return;
  state.busy = true; updateButtons(); logger("", true);
  try {
    var res = await BomSync.importProjectJson(await file.text(), file.name, logger);
    if (res.applied) {
      logger("Loaded project list \"" + res.name + "\" (" + res.count + " materials) into this workbook. The shared master list is untouched.");
    } else {
      logger("Loaded project list \"" + res.name + "\" (" + res.count + " materials). It will be used when you restore the converted sheets.");
    }
  } catch (e) {
    logger("Could not load the file: " + (e.message || JSON.stringify(e)));
  }
  ev.target.value = "";
  state.busy = false;
  await refreshMaterials();
  await refreshProject();
  await refreshLocalFileStatus();
}

async function projectStop() {
  if (!state.confirmStop) {
    state.confirmStop = true;
    $("proj-stop").textContent = "Confirm: switch back to shared master list";
    projLog("This reloads the shared master list into this workbook. Materials that exist only in the project list are removed from the sheet (your saved project file keeps them), and BOM rows using them lose their weights. Click again to confirm.", true);
    setTimeout(function () {
      if (state.confirmStop) { state.confirmStop = false; $("proj-stop").textContent = "Switch back to shared master list"; }
    }, 8000);
    return;
  }
  state.confirmStop = false;
  $("proj-stop").textContent = "Switch back to shared master list";
  state.busy = true; updateButtons(); projLog("", true);
  try {
    var r = await BomSync.stopProject(projLog);
    projLog(r.restored
      ? "Back on the shared master list (" + r.added + " added, " + r.removed + " project-only removed). Project list \"" + r.name + "\" is still saved on this computer."
      : "Back on the shared master list. Project list \"" + r.name + "\" is still saved on this computer.");
  } catch (e) {
    projLog("Could not switch back: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false;
  await refreshMaterials();
  await refreshProject();
  await refreshLocalFileStatus();
}

async function fixNames() {
  state.busy = true; updateButtons(); fixLog("", true);
  try {
    var r = await BomConvert.repairGrowingNames(fixLog);
    fixLog(r.fixed.length
      ? "Done. Every material in the BOM now shows its name in the purchase list."
      : "Nothing needed repairing.");
  } catch (e) {
    fixLog("Repair failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; updateButtons();
}

/* =====================================================================
 * Optional: GitHub Sync
 * ===================================================================== */
function nameList(arr, pick) {
  var n = arr.map(pick);
  return n.slice(0, 8).join(", ") + (n.length > 8 ? " and " + (n.length - 8) + " more" : "");
}

async function refreshSync() {
  var st = $("sync-status"), dt = $("sync-detail");
  if (!st) return;
  st.className = "status"; st.textContent = "Checking GitHub...";
  dt.textContent = "";
  $("sync-pull").disabled = true; $("sync-push").disabled = true;
  $("sync-push").hidden = !BomSync.hasToken();
  try {
    var d = await BomSync.compare();
    var lines = [];
    if (d.onlyRemote.length) lines.push("New on GitHub: " + nameList(d.onlyRemote, function (m) { return m.material; }));
    if (d.changed.length) lines.push("Different values: " + nameList(d.changed, function (c) { return c.material; }));
    if (d.onlyWb.length) lines.push("Only in this workbook: " + nameList(d.onlyWb, function (m) { return m.material; }));
    dt.textContent = lines.join("\n");
    dt.style.whiteSpace = "pre-wrap";
    var canPull = d.onlyRemote.length + d.changed.length > 0;
    var canPush = d.onlyWb.length + d.changed.length > 0;
    if (d.remoteMissing || !d.remoteCount) {
      st.className = "status";
      st.textContent = "No GitHub materials.json published (built-in computer storage is already active above).";
    } else if (!canPull && !canPush) {
      st.className = "status ok";
      st.textContent = "In sync with GitHub (" + d.remoteCount + " materials).";
    } else {
      st.className = "status warn";
      st.textContent = (canPull ? "GitHub has updates. " : "") + (canPush ? "Workbook has unpublished changes." : "");
    }
    $("sync-pull").disabled = !canPull || state.busy;
    $("sync-push").disabled = !canPush || state.busy || !BomSync.hasToken();
  } catch (e) {
    st.className = "status";
    st.textContent = "GitHub sync not configured (not needed when saving locally).";
  }
}

async function syncPull() {
  state.busy = true; syncLog("", true);
  try { await BomSync.pull(syncLog); } catch (e) { syncLog("Failed: " + (e.message || JSON.stringify(e))); }
  state.busy = false;
  await refreshMaterials(); await refreshSync();
}

async function syncPush() {
  state.busy = true; syncLog("", true);
  try { await BomSync.push(syncLog); } catch (e) { syncLog("Failed: " + (e.message || JSON.stringify(e))); }
  state.busy = false;
  await refreshSync();
}

/* =====================================================================
 * Boot
 * ===================================================================== */
Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    $("home-wb-status").textContent = "This add-in runs inside Microsoft Excel.";
    return;
  }

  // Header refresh
  $("hdr-refresh").onclick = function () { showTab(state.activeTab); };

  // Home quick options
  $("qa-add").onclick = function () { showTab("sheet"); };
  $("qa-mat").onclick = function () { showTab("mat"); };
  $("qa-flatten").onclick = function () { showTab("convert", { submode: "flatten" }); };
  $("qa-restore").onclick = function () { showTab("convert", { submode: "restore" }); };
  $("home-repair-btn").onclick = function () { repairWorkbook(homeRepairLog, "home-repair-sheets"); };

  // Add sheet
  $("sh-add").onclick = addSheet;
  $("sh-repair").onclick = function () { repairWorkbook(shLog, "sh-repair-sheets"); };
  $("sh-name").oninput = updateSheetButton;
  $("sh-source").onchange = setSheetClearDefault;

  // Materials (live primary-requirement validation)
  $("mat-search").oninput = renderMatList;
  $("m-name").oninput = updateMatValidation;
  $("m-unit").onchange = function () { unitFields(); updateMatValidation(); };
  $("m-lin").oninput = updateMatValidation;
  $("m-area").oninput = updateMatValidation;
  $("m-sl").oninput = updateMatValidation;
  $("m-sw").oninput = updateMatValidation;
  $("m-piece").oninput = updateMatValidation;
  $("m-cat").oninput = updateMoreOptionsMeta;
  $("m-rem").oninput = updateMoreOptionsMeta;
  $("m-save").onclick = saveMaterial;
  $("m-new").onclick = function () { matLog("", true); setMatMode("add"); };
  $("m-delete").onclick = removeMaterial;
  $("mat-toast-close").onclick = function () { collapseToast(false); };
  $("mat-change-toast").onmouseenter = pauseToastTimer;
  $("mat-change-toast").onmouseleave = resumeToastTimer;
  $("mat-changes-btn").onclick = function () {
    mat.sessionOpen = !mat.sessionOpen;
    renderSessionChanges();
  };
  $("mat-session-clear").onclick = function () {
    mat.sessionLog = [];
    mat.recentMap = {};
    mat.sessionOpen = false;
    renderSessionChanges();
    renderMatList();
  };

  // Built-in storage + JSON backup / transfer controls
  $("local-pull").onclick = localPull;
  $("local-save").onclick = localSave;
  $("local-export-btn").onclick = localExportDownload;
  $("local-import-btn").onclick = function () { $("local-import-input").click(); };
  $("local-import-input").onchange = localImportFile;

  // Project-only materials list + repair of plain sheets
  $("proj-start").onclick = projectStart;
  $("proj-save-btn").onclick = projectSave;
  $("proj-stop").onclick = projectStop;
  $("proj-name").oninput = updateButtons;
  Array.prototype.forEach.call(document.getElementsByName("projcopy"), function (r) {
    r.onchange = function () { renderProjectPicker(); updateButtons(); };
  });
  $("proj-load-btn").onclick = function () { $("proj-load-input").click(); };
  $("proj-load-input").onchange = function (ev) { projectLoadFile(ev, projLog); };
  $("rest-proj-load-btn").onclick = function () { $("rest-proj-load-input").click(); };
  $("rest-proj-load-input").onchange = function (ev) { projectLoadFile(ev, restProjLog); };
  $("fixnames-btn").onclick = fixNames;

  // Optional GitHub sync
  $("sync-pull").onclick = syncPull;
  $("sync-push").onclick = syncPush;
  $("gh-save").onclick = function () {
    var ok = BomSync.setToken($("gh-token").value);
    $("gh-token").value = "";
    syncLog(ok ? "Token saved on this computer." : "Could not store token.", true);
    refreshSync();
  };
  $("gh-clear").onclick = function () { BomSync.setToken(""); syncLog("Token removed.", true); refreshSync(); };

  // Convert (two-way)
  $("cmode-flatten").onclick = function () { setConvertSubmode("flatten"); };
  $("cmode-restore").onclick = function () { setConvertSubmode("restore"); };
  $("copy").onclick = function () {
    openCopy(log, "Opened a copy in a new window. Open BOM Tools there, tick the confirmation box, and convert.");
  };
  $("rest-copy").onclick = function () {
    openCopy(restLog, "Opened a copy in a new window. Open BOM Tools there to restore your sheets while keeping the original file untouched.");
  };
  $("convert").onclick = convert;
  $("iscopy").onchange = updateButtons;
  $("restore-btn").onclick = restoreAutomated;

  // Top nav tabs — always opens to Home tab!
  TABS.forEach(function (t) { $("tab-" + t).onclick = function () { showTab(t); }; });
  BomSync.loadProjectState().then(function () { refreshProject(); }, function () {});
  showTab("home");
  BomSync.watch(localLog, handleAutoSaveEvent);
});
