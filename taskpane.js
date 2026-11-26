/* taskpane.js - UI controller for BOM Tools task pane */
"use strict";

var ADDIN_VERSION = "2.4";
var $ = function (id) { return document.getElementById(id); };
var state = {
  ok: false,
  canRestore: false,
  busy: false,
  confirmDelete: false,
  confirmExitProject: false,
  confirmDeleteSavedProj: "",
  convertSubmode: "flatten",
  activeTab: "home",
  projectMode: { isProject: false, projectName: "" },
  savedProjects: [],
  restMatSrc: "global",
  restMatSrcUserTouched: false,
  restProjectData: null,
  growingConvertedSheets: [],
  autoBomCount: 0,
  projPickSet: null
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
var repairNamesLog = makeLog("rest-repair-names-log");
var shLog = makeLog("sh-log");
var matLog = makeLog("mat-log");
var projLog = makeLog("proj-log");
var localLog = makeLog("local-log");
var syncLog = makeLog("sync-log");
var homeRepairLog = makeLog("home-repair-log");

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

function getRestMatSrc() {
  var checked = document.querySelector('input[name="rest-mat-src"]:checked');
  return checked ? checked.value : (state.restMatSrc || "global");
}

function setRestMatSrc(val, fromUser) {
  var v = (val === "saved-project" || val === "project-file") ? val : "global";
  state.restMatSrc = v;
  if (fromUser) state.restMatSrcUserTouched = true;
  var map = { global: "rest-src-global", "saved-project": "rest-src-saved", "project-file": "rest-src-file" };
  Object.keys(map).forEach(function (k) {
    var r = $(map[k]);
    if (r) r.checked = (k === v);
  });
  if ($("rest-saved-proj-box")) $("rest-saved-proj-box").hidden = (v !== "saved-project");
  if ($("rest-proj-box")) $("rest-proj-box").hidden = (v !== "project-file");
  updateButtons();
}

function updateButtons() {
  var isCopy = $("iscopy") ? $("iscopy").checked : false;
  var convPicked = getPickedSheets("conv-sheets").length;
  var restPicked = getPickedSheets("rest-sheets").length;
  var shRepPicked = getPickedSheets("sh-repair-sheets").length;
  var homeRepPicked = getPickedSheets("home-repair-sheets").length;

  var restSrc = getRestMatSrc();
  var selectedSavedSlug = $("rest-proj-saved-select") ? $("rest-proj-saved-select").value : "";
  var needSavedProj = restSrc === "saved-project" && !selectedSavedSlug;
  var needRestProjFile = restSrc === "project-file" && !state.restProjectData;
  var isPartialFlatten = state.ok && state.autoBomCount > 0 && convPicked > 0 && convPicked < state.autoBomCount;

  if ($("convert")) {
    $("convert").disabled = state.busy || !state.ok || !isCopy || !convPicked;
    $("convert").textContent = isPartialFlatten
      ? ("Flatten " + convPicked + " Selected Sheet" + (convPicked === 1 ? "" : "s"))
      : "Convert to Plain Excel";
  }
  if ($("conv-partial-note")) {
    if (isPartialFlatten) {
      var remain = state.autoBomCount - convPicked;
      $("conv-partial-note").hidden = false;
      $("conv-partial-note").textContent =
        "Partial flatten (" + convPicked + " of " + state.autoBomCount + " selected): Master Unique Materials & BOM Blank will be kept for the remaining " +
        remain + " automated sheet" + (remain === 1 ? "" : "s") + ".";
    } else {
      $("conv-partial-note").hidden = true;
      $("conv-partial-note").textContent = "";
    }
  }
  if ($("conv-proj-desc") && state.ok && state.projectMode && state.projectMode.isProject) {
    $("conv-proj-desc").textContent = isPartialFlatten
      ? "Partial flatten keeps your Project-Only materials sheet in this workbook for the remaining automated sheet(s). You can still save a .json backup anytime."
      : "Flattening all sheets removes the materials sheet from this copy (a snapshot is also kept in this PC's project storage so you can select and restore it anytime). Save a .json file if you need it on another computer.";
  }

  if ($("copy")) $("copy").disabled = state.busy;
  if ($("rest-copy")) $("rest-copy").disabled = state.busy;
  if ($("restore-btn")) $("restore-btn").disabled = state.busy || !state.canRestore || !restPicked || needSavedProj || needRestProjFile;
  if ($("rest-repair-names-btn")) $("rest-repair-names-btn").disabled = state.busy || !state.growingConvertedSheets.length || !restPicked;
  if ($("sh-repair")) $("sh-repair").disabled = state.busy || !shRepPicked;
  if ($("home-repair-btn")) $("home-repair-btn").disabled = state.busy || !homeRepPicked;

  var hasSavedPick = $("proj-saved-select") && !!$("proj-saved-select").value;
  if ($("proj-saved-load-btn")) $("proj-saved-load-btn").disabled = state.busy || !hasSavedPick;
  if ($("proj-saved-export-btn")) $("proj-saved-export-btn").disabled = state.busy || !hasSavedPick;
  if ($("proj-saved-delete-btn")) $("proj-saved-delete-btn").disabled = state.busy || !hasSavedPick;

  if (!state.ok) {
    setReqHint("conv-req-hint", "Open an automated BOM workbook to flatten.", false);
  } else if (!convPicked) {
    setReqHint("conv-req-hint", "Select at least one BOM sheet above to flatten.", true);
  } else if (!isCopy) {
    setReqHint("conv-req-hint", "Tick the copy confirmation box above to enable conversion.", false);
  } else {
    setReqHint("conv-req-hint", "", false);
  }

  if (!state.canRestore) {
    setReqHint("rest-req-hint", "No converted BOM sheets found to restore.", false);
  } else if (!restPicked) {
    setReqHint("rest-req-hint", "Select at least one sheet above to restore.", true);
  } else if (needSavedProj) {
    setReqHint("rest-req-hint", "No saved Project-Only list selected above. Choose Global Master or load a Project .json file instead.", false);
  } else if (needRestProjFile) {
    setReqHint("rest-req-hint", "Choose your saved Project Materials .json file above to continue.", false);
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
    if (!window.BOM_OFFLINE_LTS) refreshSync();
    renderSessionChanges();
    BomSync.watch(localLog, handleAutoSaveEvent);
  }
  if (name === "convert") {
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

    if (info.masterMode) state.projectMode = info.masterMode;
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
    var isProj = !!(state.projectMode && state.projectMode.isProject);
    var pName = isProj ? (state.projectMode.projectName || "Project Custom Materials") : "";
    var pCountText = info.hasProjectCache ? (" (" + info.projectCount + " saved)") : "";
    if (info.hasCache) {
      if (st) {
        st.className = "status ok";
        st.textContent = "Auto-saving Global Master in built-in storage (" + info.cachedCount + " materials) — never asks to reconnect.";
      }
      if (tag) { tag.className = "sum-meta ok"; tag.textContent = info.cachedCount + " saved"; }
      if (homeSt) {
        homeSt.textContent = isProj
          ? ("Project-Only list active (\"" + pName + "\"" + pCountText + ") · Global Master untouched (" + info.cachedCount + " items).")
          : ("Master list auto-saves on this computer (" + info.cachedCount + " items).");
      }
      if (shSt) {
        shSt.textContent = isProj
          ? ("Project-Only list active (\"" + pName + "\") — new sheets will keep your project materials.")
          : ("Auto-connected to built-in storage (" + info.cachedCount + " items).");
      }
    } else {
      if (st) {
        st.className = "status";
        st.textContent = "Built-in storage is ready. Open a BOM workbook or import a JSON backup to populate it.";
      }
      if (tag) { tag.className = "sum-meta"; tag.textContent = "Auto-saved"; }
      if (homeSt) {
        homeSt.textContent = isProj
          ? ("Project-Only list active (\"" + pName + "\"" + pCountText + ") · Global Master untouched.")
          : "Master list will auto-save on this computer once a BOM is opened.";
      }
      if (shSt) {
        shSt.textContent = isProj
          ? ("Project-Only list active (\"" + pName + "\") — new sheets keep your project materials.")
          : "Built-in storage ready — materials auto-save on this computer.";
      }
    }

    if (pullBtn) pullBtn.disabled = state.busy || !info.hasCache || !mat.hasMaster || isProj;
    if (saveBtn) saveBtn.disabled = state.busy || !mat.rows.length || isProj;
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
          suggestedName: "master-materials.json",
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
    a.href = url; a.download = "master-materials.json";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    localLog("Exported master-materials.json (" + exp.count + " materials) for backup or transfer.");
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
    if (sheetInfo.masterMode) state.projectMode = sheetInfo.masterMode;
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
    whereEl.textContent = evt.projectOnly
      ? "Saved to Project list on this computer (Global Master untouched)"
      : "Saved to Master sheet & this computer";
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
    savedToBuiltIn: evt.projectOnly ? false : true,
    projectOnly: !!evt.projectOnly,
    projectName: evt.projectName || "",
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

function resetExitProjectConfirm() {
  state.confirmExitProject = false;
  if ($("proj-exit-btn")) $("proj-exit-btn").textContent = "Switch back to Global Master";
}

function initProjPicker() {
  state.projPickSet = {};
  mat.rows.forEach(function (r) {
    if (r.material) state.projPickSet[r.material.toLowerCase()] = true;
  });
  if ($("proj-pick-search")) $("proj-pick-search").value = "";
}

function getProjPickedNames() {
  if (!state.projPickSet) return [];
  return mat.rows
    .filter(function (r) { return r.material && !!state.projPickSet[r.material.toLowerCase()]; })
    .map(function (r) { return r.material; });
}

function renderProjPicker() {
  var wrap = $("proj-pick-wrap");
  var listEl = $("proj-pick-list");
  var countEl = $("proj-pick-count");
  var toggleBtn = $("proj-pick-toggle");
  var srcSel = $("proj-source-select");
  if (!wrap || !listEl || !srcSel) return;

  var isPick = srcSel.value === "pick";
  wrap.hidden = !isPick;
  if (!isPick) return;

  if (!state.projPickSet) initProjPicker();

  var q = $("proj-pick-search") ? $("proj-pick-search").value.trim().toLowerCase() : "";
  var filtered = mat.rows.filter(function (r) {
    return !q || (r.material + " " + (r.category || "")).toLowerCase().indexOf(q) >= 0;
  });

  var totalSel = getProjPickedNames().length;
  if (countEl) countEl.textContent = totalSel + " of " + mat.rows.length + " selected";

  var allFilteredOn = filtered.length > 0 && filtered.every(function (r) {
    return !!state.projPickSet[r.material.toLowerCase()];
  });
  if (toggleBtn) toggleBtn.textContent = allFilteredOn ? "None" : "All";

  listEl.innerHTML = "";
  filtered.forEach(function (r) {
    var key = r.material.toLowerCase();
    var row = document.createElement("label");
    row.className = "proj-pick-item";

    var cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !!state.projPickSet[key];
    cb.onchange = function () {
      if (cb.checked) state.projPickSet[key] = true;
      else delete state.projPickSet[key];
      var selNow = getProjPickedNames().length;
      if (countEl) countEl.textContent = selNow + " of " + mat.rows.length + " selected";
      var allNow = filtered.length > 0 && filtered.every(function (x) { return !!state.projPickSet[x.material.toLowerCase()]; });
      if (toggleBtn) toggleBtn.textContent = allNow ? "None" : "All";
    };

    var nm = document.createElement("span");
    nm.className = "pname";
    nm.textContent = r.material;

    var meta = document.createElement("small");
    meta.textContent = r.unit || "";

    row.appendChild(cb);
    row.appendChild(nm);
    row.appendChild(meta);
    listEl.appendChild(row);
  });

  if (!filtered.length) {
    var empty = document.createElement("div");
    empty.className = "proj-pick-item";
    empty.style.cursor = "default";
    empty.textContent = "No materials match filter.";
    listEl.appendChild(empty);
  }
}

function toggleProjPickerFiltered() {
  if (!state.projPickSet) initProjPicker();
  var q = $("proj-pick-search") ? $("proj-pick-search").value.trim().toLowerCase() : "";
  var filtered = mat.rows.filter(function (r) {
    return !q || (r.material + " " + (r.category || "")).toLowerCase().indexOf(q) >= 0;
  });
  var allOn = filtered.length > 0 && filtered.every(function (r) {
    return !!state.projPickSet[r.material.toLowerCase()];
  });
  filtered.forEach(function (r) {
    var key = r.material.toLowerCase();
    if (allOn) delete state.projPickSet[key];
    else state.projPickSet[key] = true;
  });
  renderProjPicker();
}

function renderProjectScopeUI() {
  var card = $("proj-scope-card");
  var title = $("proj-scope-title");
  var sub = $("proj-scope-sub");
  var openBtn = $("proj-open-setup");
  var setupPanel = $("proj-setup-panel");
  var noMasterPanel = $("proj-nomaster-actions");
  var activePanel = $("proj-active-actions");
  var listTitle = $("mat-list-title");
  if (!card) return;

  var isProj = !!(state.projectMode && state.projectMode.isProject);
  card.className = "scope-card" + (isProj ? " is-project" : "");

  if (isProj) {
    var pName = state.projectMode.projectName || "Project Custom Materials";
    if (title) title.textContent = "Project-Only: " + pName;
    if (sub) {
      sub.textContent = mat.hasMaster
        ? "Auto-saves separately on this PC — Global Master is untouched."
        : "Saved for this converted workbook — Restore Automated will rebuild using this project list.";
    }
    if (openBtn) openBtn.hidden = true;
    if (setupPanel) setupPanel.hidden = true;
    if (noMasterPanel) noMasterPanel.hidden = true;
    if (activePanel) activePanel.hidden = false;
    if (listTitle) listTitle.textContent = "Project Materials";
  } else {
    if (title) title.textContent = "Global Master List";
    if (sub) sub.textContent = mat.hasMaster
      ? "Edits auto-save to this computer and apply to future BOMs."
      : "Converted or plain workbook (no Master sheet). Load a Project .json below to prepare for Restore.";
    if (openBtn) {
      openBtn.hidden = !mat.hasMaster;
      openBtn.disabled = state.busy;
    }
    if (noMasterPanel) noMasterPanel.hidden = mat.hasMaster;
    if (activePanel) activePanel.hidden = true;
    if (listTitle) listTitle.textContent = "Master Materials";
    resetExitProjectConfirm();
  }
  renderSavedProjectsUI();
}

function formatShortDate(iso) {
  if (!iso) return "";
  try {
    var d = new Date(iso);
    if (isNaN(d.getTime())) return "";
    var now = new Date();
    if (d.toDateString() === now.toDateString()) {
      return d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    }
    return d.toLocaleDateString([], { month: "short", day: "numeric" });
  } catch (e) { return ""; }
}

function formatProjectOptionLabel(p) {
  var parts = [p.count + " item" + (p.count === 1 ? "" : "s")];
  var dt = formatShortDate(p.updated);
  if (dt) parts.push(dt);
  return p.name + " (" + parts.join(" · ") + ")";
}

function resetDeleteSavedProjConfirm() {
  state.confirmDeleteSavedProj = "";
  if ($("proj-saved-delete-btn")) $("proj-saved-delete-btn").textContent = "Delete";
}

function renderSavedProjectsUI(preferredNameOrSlug) {
  var list = state.savedProjects || [];
  var prefSlug = preferredNameOrSlug && BomSync.slug ? BomSync.slug(preferredNameOrSlug) : "";

  // 1. Materials Scope Card: Saved Project Lists drawer
  var drawer = $("proj-saved-drawer");
  var countEl = $("proj-saved-count");
  var matSel = $("proj-saved-select");
  var matNote = $("proj-saved-note");
  if (drawer && matSel) {
    var keepMat = matSel.value;
    matSel.innerHTML = "";
    list.forEach(function (p) {
      var o = document.createElement("option");
      o.value = p.slug;
      o.textContent = formatProjectOptionLabel(p);
      matSel.appendChild(o);
    });
    if (list.length) {
      drawer.hidden = false;
      if (countEl) countEl.textContent = list.length + " saved";
      var hasKeep = list.some(function (p) { return p.slug === keepMat; });
      var hasPref = prefSlug && list.some(function (p) { return p.slug === prefSlug; });
      matSel.value = hasPref ? prefSlug : (hasKeep ? keepMat : list[0].slug);
      if (!mat.hasMaster) drawer.open = true;
      if (matNote) {
        matNote.textContent = mat.hasMaster
          ? "Load switches this workbook's materials sheet to the chosen project list (Global Master stays untouched)."
          : "Load links the chosen project list to this converted workbook — or pick it directly in Convert > Restore Automated.";
      }
    } else {
      drawer.hidden = true;
      if (countEl) countEl.textContent = "0 saved";
      if (matNote) matNote.textContent = "";
    }
    resetDeleteSavedProjConfirm();
  }

  // 2. Convert > Restore Automated: Saved Project-Only selector
  var restSel = $("rest-proj-saved-select");
  var restNote = $("rest-saved-proj-note");
  if (restSel) {
    var keepRest = restSel.value;
    restSel.innerHTML = "";
    if (list.length) {
      list.forEach(function (p) {
        var o = document.createElement("option");
        o.value = p.slug;
        o.textContent = formatProjectOptionLabel(p);
        restSel.appendChild(o);
      });
      var restHasPref = prefSlug && list.some(function (p) { return p.slug === prefSlug; });
      var restHasKeep = list.some(function (p) { return p.slug === keepRest; });
      restSel.value = restHasPref ? prefSlug : (restHasKeep ? keepRest : list[0].slug);
      restSel.disabled = false;
      if (restNote) {
        restNote.textContent = list.length + " saved project list" + (list.length === 1 ? "" : "s") + " found on this computer. Restore will rebuild sheets with the selected list.";
      }
    } else {
      var emptyOpt = document.createElement("option");
      emptyOpt.value = "";
      emptyOpt.textContent = "(No project lists saved on this PC yet)";
      restSel.appendChild(emptyOpt);
      restSel.disabled = true;
      if (restNote) {
        restNote.textContent = "No project lists saved in this add-in yet. Use a Project .json file below, or create a Project-Only list in Materials.";
      }
    }
  }
  updateButtons();
}

async function refreshSavedProjects(preferredNameOrSlug) {
  try {
    if (BomSync.listProjectSlots) {
      state.savedProjects = await BomSync.listProjectSlots();
    }
  } catch (e) {
    state.savedProjects = [];
  }
  renderSavedProjectsUI(preferredNameOrSlug);
  return state.savedProjects;
}

async function loadSavedProjectIntoWorkbook() {
  var sel = $("proj-saved-select");
  var slugKey = sel ? sel.value : "";
  if (!slugKey || state.busy) return;
  state.busy = true; updateButtons();
  projLog("", true);
  try {
    await BomSync.applySavedProjectSlot(slugKey, projLog);
  } catch (e) {
    projLog("Could not load saved project list: " + (e.message || e));
  }
  state.busy = false;
  await refreshMaterials();
  await refreshLocalFileStatus();
  await refreshSavedProjects(slugKey);
}

async function exportSavedProjectFile() {
  var sel = $("proj-saved-select");
  var slugKey = sel ? sel.value : "";
  if (!slugKey) return;
  projLog("", true);
  try {
    var slot = await BomSync.readProjectSlotBySlug(slugKey);
    if (!slot || !slot.list || !slot.list.length) {
      projLog("That saved project list is empty or missing.");
      return;
    }
    var pName = slot.name || slugKey;
    var json = BomSync.serializeProjectList(slot.list, pName, slot.lastChange || null);
    var fileName = (BomSync.slug ? BomSync.slug(pName) : slugKey) + "-materials.json";
    if (BomSync.supportsSavePicker && BomSync.supportsSavePicker()) {
      try {
        var handle = await window.showSaveFilePicker({
          suggestedName: fileName,
          types: [{ description: "Project Materials JSON", accept: { "application/json": [".json"] } }]
        });
        var writable = await handle.createWritable();
        await writable.write(json);
        await writable.close();
        projLog("Exported \"" + handle.name + "\" (" + slot.list.length + " materials).");
        return;
      } catch (err) {
        if (err && err.name === "AbortError") return;
      }
    }
    var blob = new Blob([json], { type: "application/json;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = fileName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    projLog("Exported " + fileName + " (" + slot.list.length + " materials).");
  } catch (e) {
    projLog("Export failed: " + (e.message || e));
  }
}

async function deleteSavedProjectSlot() {
  var sel = $("proj-saved-select");
  var slugKey = sel ? sel.value : "";
  if (!slugKey || state.busy) return;
  var found = (state.savedProjects || []).filter(function (p) { return p.slug === slugKey; })[0];
  var label = (found && found.name) || slugKey;
  if (state.confirmDeleteSavedProj !== slugKey) {
    state.confirmDeleteSavedProj = slugKey;
    if ($("proj-saved-delete-btn")) $("proj-saved-delete-btn").textContent = "Confirm";
    projLog("Click \"Confirm\" to remove saved project list \"" + label + "\" from this PC (Global Master and open workbooks are untouched).", true);
    return;
  }
  state.busy = true; updateButtons();
  projLog("", true);
  try {
    await BomSync.deleteProjectSlot(slugKey);
    projLog("Removed saved project list \"" + label + "\" from this computer.");
  } catch (e) {
    projLog("Could not delete saved project list: " + (e.message || e));
  }
  resetDeleteSavedProjConfirm();
  state.busy = false;
  await refreshSavedProjects();
}

async function createProjectList() {
  if (!mat.hasMaster || state.busy) return;
  var pName = ($("proj-name-input") ? $("proj-name-input").value.trim() : "") || "Project Custom Materials";
  var source = $("proj-source-select") ? $("proj-source-select").value : "clone";
  var keepNames = null;
  if (source === "pick") {
    keepNames = getProjPickedNames();
    if (!keepNames.length) {
      projLog("Select at least one material from the list above (or choose \"Blank one-time list\").", true);
      return;
    }
  }
  state.busy = true; projLog("", true);
  try {
    BomSync.suppressWatch(4000);
    await BomTools.createProjectMaterialList({ projectName: pName, source: source, keepNames: keepNames }, projLog);
    var list = await BomSync.workbookList();
    BomSync.updateWatchSnap(list);
    if (BomSync.saveProjectSlot) {
      await BomSync.saveProjectSlot(pName, list, null);
    }
    if ($("proj-setup-panel")) $("proj-setup-panel").hidden = true;
  } catch (e) {
    projLog("Could not create Project-Only list: " + (e.message || e));
  }
  state.busy = false;
  await refreshMaterials();
  await refreshLocalFileStatus();
  await refreshSavedProjects(pName);
}

async function saveProjectListFile(logger) {
  logger = logger || projLog;
  logger("", true);
  try {
    var pName = (state.projectMode && state.projectMode.projectName) || "Project Custom Materials";
    var exp = await BomSync.buildProjectExportJson(pName);
    if (!exp.count) {
      logger("Add at least one material to the project list before saving.");
      return false;
    }
    if (BomSync.supportsSavePicker && BomSync.supportsSavePicker()) {
      try {
        var handle = await window.showSaveFilePicker({
          suggestedName: exp.suggestedName,
          types: [{ description: "Project Materials JSON", accept: { "application/json": [".json"] } }]
        });
        var writable = await handle.createWritable();
        await writable.write(exp.json);
        await writable.close();
        logger("Saved \"" + handle.name + "\" (" + exp.count + " project materials). You can reload it anytime even after flattening.");
        return true;
      } catch (err) {
        if (err && err.name === "AbortError") return false;
      }
    }
    var blob = new Blob([exp.json], { type: "application/json;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = exp.suggestedName;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    logger("Saved " + exp.suggestedName + " (" + exp.count + " project materials) for reuse.");
    return true;
  } catch (e) {
    logger("Save failed: " + (e.message || e));
    return false;
  }
}

async function loadProjectListFile(ev) {
  var file = ev.target && ev.target.files && ev.target.files[0];
  if (!file) return;
  state.busy = true; projLog("", true);
  try {
    var text = await file.text();
    await BomSync.importProjectJsonText(text, file.name, projLog);
    if ($("proj-setup-panel")) $("proj-setup-panel").hidden = true;
  } catch (e) {
    projLog("Could not load Project JSON: " + (e.message || e));
  }
  ev.target.value = "";
  state.busy = false;
  await refreshMaterials();
  await refreshLocalFileStatus();
  await refreshSavedProjects();
}

async function exitProjectMode() {
  if (!state.projectMode || !state.projectMode.isProject || state.busy) return;
  if (!state.confirmExitProject) {
    state.confirmExitProject = true;
    if ($("proj-exit-btn")) $("proj-exit-btn").textContent = "Confirm switch to Global Master";
    projLog("Click \"Confirm switch to Global Master\" to reconnect this workbook to your Global Master list (save your Project .json first if you want to keep it).", true);
    return;
  }
  state.busy = true; projLog("", true);
  try {
    BomSync.suppressWatch(4000);
    await BomTools.exitProjectMaterialMode();
    if (mat.hasMaster) {
      await BomSync.pullLocal(projLog, { force: true, syncRemovals: true });
      var list = await BomSync.workbookList();
      BomSync.updateWatchSnap(list);
    }
    projLog("Switched back to Global Master list.");
  } catch (e) {
    projLog("Switch failed: " + (e.message || e));
  }
  state.busy = false;
  resetExitProjectConfirm();
  await refreshMaterials();
  await refreshLocalFileStatus();
}

async function refreshMaterials(keepName) {
  var st = $("mat-status");
  st.textContent = "Loading...";
  try {
    var res = await BomTools.listMaterials();
    mat.rows = res.rows;
    mat.hasMaster = true;
    state.projectMode = res.masterMode || { isProject: false, projectName: "" };
    renderProjectScopeUI();
    var dl = $("m-cats"); dl.innerHTML = "";
    res.categories.forEach(function (c) { var o = document.createElement("option"); o.value = c; dl.appendChild(o); });
    var incompleteCount = res.rows.filter(function (r) { return !BomTools.getMaterialValidationStatus(r).valid; }).length;
    st.textContent = res.rows.length + " items" + (incompleteCount ? (" · " + incompleteCount + " incomplete") : " · all ready");
    var again = keepName && res.rows.filter(function (r) { return r.material.toLowerCase() === keepName.toLowerCase(); })[0];
    if (again) setMatMode("edit", again);
    else if (mat.mode === "edit" && mat.selected) setMatMode("add");
    else { renderMatList(); updateMatValidation(); }
    if ($("proj-setup-panel") && !$("proj-setup-panel").hidden && $("proj-source-select") && $("proj-source-select").value === "pick") {
      renderProjPicker();
    }
  } catch (e) {
    mat.rows = [];
    mat.hasMaster = false;
    try {
      state.projectMode = await BomTools.getMasterMode();
    } catch (e2) {
      state.projectMode = { isProject: false, projectName: "" };
    }
    renderProjectScopeUI();
    st.textContent = "No Master sheet in workbook";
    renderMatList();
    updateMatValidation();
  }
  await refreshSavedProjects(state.projectMode && state.projectMode.isProject ? state.projectMode.projectName : "");
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

    var isProj = !!(state.projectMode && state.projectMode.isProject);
    var list = await BomSync.workbookList();
    list.forEach(function (item) {
      if (item.material.toLowerCase() === r.material.toLowerCase()) {
        item.lastAction = changeType;
        item.lastModified = nowIso;
      }
    });
    BomSync.updateWatchSnap(list);
    if (isProj) {
      if (BomSync.saveProjectSlot) {
        await BomSync.saveProjectSlot(state.projectMode.projectName, list, lastChange);
      }
    } else {
      await BomSync.writeLocalMaster(list, matLog, { lastChange: lastChange });
    }

    recordAutoSave({
      source: "panel",
      time: nowIso,
      savedToBuiltIn: !isProj,
      projectOnly: isProj,
      projectName: isProj ? state.projectMode.projectName : "",
      items: changeItems
    });
    if (!isProj && BomSync.hasToken()) {
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
  var isProj = !!(state.projectMode && state.projectMode.isProject);
  if (!state.confirmDelete) {
    state.confirmDelete = true;
    $("m-delete").textContent = "Confirm";
    matLog(isProj
      ? ("Click \"Confirm\" to remove \"" + mat.selected.material + "\" from this project list (Global Master stays untouched).")
      : ("Click \"Confirm\" to permanently remove \"" + mat.selected.material + "\" from this workbook and your saved list."), true);
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
    if (isProj) {
      if (BomSync.saveProjectSlot) {
        await BomSync.saveProjectSlot(state.projectMode.projectName, list, lastChange);
      }
    } else {
      await BomSync.writeLocalMaster(list, matLog, { lastChange: lastChange });
    }
    recordAutoSave({
      source: "panel",
      time: nowIso,
      savedToBuiltIn: !isProj,
      projectOnly: isProj,
      projectName: isProj ? state.projectMode.projectName : "",
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
    state.autoBomCount = info.bomSheets ? info.bomSheets.length : 0;
    state.canRestore = info.convertedSheets && info.convertedSheets.length > 0;
    state.growingConvertedSheets = info.growingConvertedSheets || [];
    state.projectMode = info.masterMode || { isProject: false, projectName: "" };

    var projCallout = $("conv-proj-callout");
    if (projCallout) {
      var showProjBackup = state.ok && !!(state.projectMode && state.projectMode.isProject);
      projCallout.hidden = !showProjBackup;
      if (showProjBackup && $("conv-proj-title")) {
        $("conv-proj-title").textContent = "Project-Only List Active: " + (state.projectMode.projectName || "Project Custom Materials");
      }
    }

    var preferredProj = (state.projectMode && state.projectMode.isProject) ? state.projectMode.projectName : "";
    var savedList = await refreshSavedProjects(preferredProj);

    var restProjDet = $("rest-proj-detected");
    if (restProjDet) {
      var showRestProj = state.canRestore && !!(state.projectMode && state.projectMode.isProject);
      restProjDet.hidden = !showRestProj;
      if (showRestProj) {
        var dpName = state.projectMode.projectName || "Project Custom Materials";
        if ($("rest-proj-detected-title")) {
          $("rest-proj-detected-title").textContent = "Project-Only Workbook Detected: \"" + dpName + "\"";
        }
        var dpSlug = BomSync.slug ? BomSync.slug(dpName) : "";
        var matchedSlot = savedList.filter(function (p) { return p.slug === dpSlug; })[0] || null;
        if ($("rest-proj-detected-desc")) {
          if (matchedSlot) {
            $("rest-proj-detected-desc").textContent =
              "Found \"" + matchedSlot.name + "\" (" + matchedSlot.count + " material(s)) saved on this computer and pre-selected it below. You can also switch to a different saved list, Global Master, or a .json file.";
          } else {
            $("rest-proj-detected-desc").textContent =
              "This workbook used Project-Only list \"" + dpName + "\", which wasn't matched automatically. Pick any saved project list on this PC or load a .json file below.";
          }
        }
        if (!state.restMatSrcUserTouched) {
          if (matchedSlot) {
            setRestMatSrc("saved-project", false);
            if ($("rest-proj-saved-select")) $("rest-proj-saved-select").value = matchedSlot.slug;
          } else if (savedList.length) {
            setRestMatSrc("saved-project", false);
          } else {
            setRestMatSrc("project-file", false);
          }
        }
      } else if (state.canRestore && !state.restMatSrcUserTouched) {
        setRestMatSrc("global", false);
      }
    }

    if ($("rest-repair-names-drawer")) {
      $("rest-repair-names-drawer").hidden = !state.growingConvertedSheets.length;
    }

    if (state.ok) {
      setDot("conv-dot", "ok");
      st.className = "status ok";
      st.textContent = "Ready · " + info.bomSheets.length + " automated BOM sheet(s): " + info.bomSheets.join(", ");
      renderSheetPicker("conv-sheets", info.bomSheets, "Automated BOM sheets to flatten");
    } else {
      setDot("conv-dot", "warn");
      st.className = "status warn";
      st.textContent = state.canRestore
        ? "This workbook looks already converted. Switch to \"Restore Automated\" above to convert it back."
        : "No automated BOM sheets found in this workbook.";
      renderSheetPicker("conv-sheets", [], "");
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
    state.ok = false; state.autoBomCount = 0; state.canRestore = false; state.growingConvertedSheets = [];
    if ($("conv-proj-callout")) $("conv-proj-callout").hidden = true;
    if ($("rest-proj-detected")) $("rest-proj-detected").hidden = true;
    if ($("rest-repair-names-drawer")) $("rest-repair-names-drawer").hidden = true;
    setDot("conv-dot", "warn"); setDot("rest-dot", "warn");
    st.className = "status warn"; st.textContent = "Could not read workbook: " + (e.message || e);
    rst.className = "status warn"; rst.textContent = "Could not read workbook: " + (e.message || e);
    renderSheetPicker("conv-sheets", [], "");
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
    await (BomTools.runExcel || Excel.run)(async function (context) {
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
  var picked = getPickedSheets("conv-sheets");
  if (!picked.length) {
    log("Select at least one automated BOM sheet to flatten.", true);
    return;
  }
  var mode = document.querySelector('input[name="mode"]:checked').value;
  var isProj = !!(state.projectMode && state.projectMode.isProject);
  var isFullFlatten = state.autoBomCount > 0 && picked.length >= state.autoBomCount;
  var autoSaveProj = isProj && isFullFlatten && $("conv-save-proj") && $("conv-save-proj").checked;
  state.busy = true; updateButtons();
  log("", true);
  try {
    if (autoSaveProj) {
      log("Saving project materials (.json) before converting...");
      var savedOk = await saveProjectListFile(log);
      if (!savedOk) {
        log("Conversion paused — save your project materials file first (or uncheck the auto-save box above).");
        state.busy = false;
        updateButtons();
        return;
      }
    }
    var warnings = await BomConvert.convertWorkbook(mode, log, {
      compiled: $("compiled").checked,
      sheets: picked
    });
    log("Done! Flattened " + picked.length + " sheet(s). Review the sheets, then use File > Save As to save.");
    warnings.forEach(function (w) { log("WARNING: " + w); });
  } catch (e) {
    log("Conversion failed: " + (e.message || JSON.stringify(e)));
    log("Close without saving and try again on a fresh copy.");
  }
  state.busy = false;
  await check(false);
}

async function repairGrowingNames() {
  var picked = getPickedSheets("rest-sheets");
  if (!picked.length) {
    repairNamesLog("Select at least one converted sheet above first.", true);
    return;
  }
  state.busy = true; updateButtons();
  repairNamesLog("", true);
  try {
    var r = await BomConvert.repairGrowingNames(repairNamesLog, { sheets: picked });
    if (r.fixed.length) {
      repairNamesLog("Done! Upgraded " + r.fixed.length + " sheet(s) to Excel 365 dynamic purchase lists.");
    } else {
      repairNamesLog("No auto-growing converted sheets were upgraded (fixed sheets were skipped).");
    }
  } catch (e) {
    repairNamesLog("Upgrade failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false;
  await check(false);
}

async function handleRestoreProjectFile(ev) {
  var file = ev.target && ev.target.files && ev.target.files[0];
  if (!file) return;
  try {
    var text = await file.text();
    var parsed = BomSync.parsePayload(text);
    if (!parsed.list.length) throw new Error("No materials found in that JSON file.");
    var fallbackName = String(file.name || "Project Materials")
      .replace(/\.json$/i, "")
      .replace(/[-_]+materials$/i, "")
      .replace(/[-_]+/g, " ")
      .trim() || "Project Custom Materials";
    state.restProjectData = {
      fileName: file.name,
      projectName: parsed.projectName || fallbackName,
      list: parsed.list,
      lastChange: parsed.lastChange || null
    };
    if ($("rest-proj-status")) {
      $("rest-proj-status").textContent = "Ready: \"" + state.restProjectData.projectName + "\" (" + parsed.list.length + " materials from " + file.name + ").";
      $("rest-proj-status").className = "status ok";
    }
    if ($("rest-proj-clear-btn")) $("rest-proj-clear-btn").hidden = false;
  } catch (e) {
    state.restProjectData = null;
    if ($("rest-proj-status")) {
      $("rest-proj-status").textContent = "Could not read file: " + (e.message || e);
      $("rest-proj-status").className = "status warn";
    }
    if ($("rest-proj-clear-btn")) $("rest-proj-clear-btn").hidden = true;
  }
  ev.target.value = "";
  updateButtons();
}

async function restoreAutomated() {
  var picked = getPickedSheets("rest-sheets");
  if (!picked.length) {
    restLog("Select at least one converted sheet to restore.", true);
    return;
  }
  var src = getRestMatSrc();
  var projData = null;
  if (src === "project-file") {
    if (!state.restProjectData) {
      restLog("Choose your saved Project Materials .json file first.", true);
      return;
    }
    projData = state.restProjectData;
  } else if (src === "saved-project") {
    var selSlug = $("rest-proj-saved-select") ? $("rest-proj-saved-select").value : "";
    if (!selSlug) {
      restLog("Select a saved Project-Only list on this PC first (or switch to Global Master / .json file).", true);
      return;
    }
    var slot = await BomSync.readProjectSlotBySlug(selSlug);
    if (!slot || !slot.list || !slot.list.length) {
      restLog("Could not load that saved project list (it may have been deleted). Click Refresh and pick another list.", true);
      return;
    }
    projData = {
      projectName: slot.name || selSlug,
      list: slot.list,
      lastChange: slot.lastChange || null
    };
  }
  state.busy = true; updateButtons();
  restLog("", true);
  try {
    var res = await BomConvert.restoreAutomatedWorkbook(restLog, {
      sheets: picked,
      materialSource: src === "global" ? "global" : (src === "saved-project" ? "project-slot" : "project-file"),
      syncLocal: src === "global",
      projectMaterials: projData,
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
 * Optional: GitHub Sync
 * ===================================================================== */
function nameList(arr, pick) {
  var n = arr.map(pick);
  return n.slice(0, 8).join(", ") + (n.length > 8 ? " and " + (n.length - 8) + " more" : "");
}

async function refreshSync() {
  if (window.BOM_OFFLINE_LTS) return;
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
  if ($("hdr-version")) {
    $("hdr-version").textContent = "v" + ADDIN_VERSION + (window.BOM_OFFLINE_LTS ? " · Offline" : "");
    $("hdr-version").title = "BOM Tools Add-in v" + ADDIN_VERSION + (window.BOM_OFFLINE_LTS ? " (Offline Windows LTS)" : "");
  }
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

  // Project-Only / One-Time Material List controls
  $("proj-open-setup").onclick = function () {
    projLog("", true);
    var panel = $("proj-setup-panel");
    if (panel) panel.hidden = !panel.hidden;
    if (panel && !panel.hidden) {
      if ($("proj-name-input") && !$("proj-name-input").value && sheetInfo && sheetInfo.project) {
        $("proj-name-input").value = sheetInfo.project + " Materials";
      }
      initProjPicker();
      renderProjPicker();
    }
  };
  $("proj-source-select").onchange = function () {
    if ($("proj-source-select").value === "pick" && !state.projPickSet) initProjPicker();
    renderProjPicker();
  };
  $("proj-pick-search").oninput = renderProjPicker;
  $("proj-pick-toggle").onclick = toggleProjPickerFiltered;
  $("proj-cancel-setup-btn").onclick = function () {
    if ($("proj-setup-panel")) $("proj-setup-panel").hidden = true;
    projLog("", true);
  };
  $("proj-create-btn").onclick = createProjectList;
  $("proj-export-btn").onclick = function () { saveProjectListFile(projLog); };
  $("proj-load-btn").onclick = function () { $("proj-import-input").click(); };
  $("proj-load-setup-btn").onclick = function () { $("proj-import-input").click(); };
  $("proj-load-nomaster-btn").onclick = function () { $("proj-import-input").click(); };
  $("proj-import-input").onchange = loadProjectListFile;
  $("proj-exit-btn").onclick = exitProjectMode;
  $("proj-saved-select").onchange = function () { resetDeleteSavedProjConfirm(); updateButtons(); };
  $("proj-saved-load-btn").onclick = loadSavedProjectIntoWorkbook;
  $("proj-saved-export-btn").onclick = exportSavedProjectFile;
  $("proj-saved-delete-btn").onclick = deleteSavedProjectSlot;

  // Built-in storage + JSON backup / transfer controls
  $("local-pull").onclick = localPull;
  $("local-save").onclick = localSave;
  $("local-export-btn").onclick = localExportDownload;
  $("local-import-btn").onclick = function () { $("local-import-input").click(); };
  $("local-import-input").onchange = localImportFile;

  // Optional GitHub sync (hidden automatically in Offline LTS edition)
  if (window.BOM_OFFLINE_LTS && $("github-drawer")) {
    $("github-drawer").hidden = true;
  }
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
  $("conv-proj-save-btn").onclick = function () { saveProjectListFile(log); };
  $("copy").onclick = function () {
    openCopy(log, "Opened a copy in a new window. Open BOM Tools there, tick the confirmation box, and convert.");
  };
  $("rest-copy").onclick = function () {
    openCopy(restLog, "Opened a copy in a new window. Open BOM Tools there to restore your sheets while keeping the original file untouched.");
  };
  $("convert").onclick = convert;
  $("iscopy").onchange = updateButtons;
  $("rest-src-global").onchange = function () { setRestMatSrc("global", true); };
  $("rest-src-saved").onchange = function () { setRestMatSrc("saved-project", true); refreshSavedProjects(); };
  $("rest-src-file").onchange = function () { setRestMatSrc("project-file", true); };
  $("rest-proj-saved-select").onchange = updateButtons;
  $("rest-proj-refresh-btn").onclick = function () { refreshSavedProjects(); };
  $("rest-proj-pick-btn").onclick = function () { $("rest-proj-input").click(); };
  $("rest-proj-input").onchange = handleRestoreProjectFile;
  $("rest-proj-clear-btn").onclick = function () {
    state.restProjectData = null;
    if ($("rest-proj-status")) {
      $("rest-proj-status").textContent = "No project materials file chosen yet.";
      $("rest-proj-status").className = "note";
    }
    if ($("rest-proj-clear-btn")) $("rest-proj-clear-btn").hidden = true;
    updateButtons();
  };
  $("restore-btn").onclick = restoreAutomated;
  $("rest-repair-names-btn").onclick = repairGrowingNames;

  // Top nav tabs — always opens to Home tab!
  TABS.forEach(function (t) { $("tab-" + t).onclick = function () { showTab(t); }; });
  showTab("home");
  BomSync.watch(localLog, handleAutoSaveEvent);
});
