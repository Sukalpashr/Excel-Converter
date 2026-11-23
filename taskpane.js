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

function setDot(id, kind) {
  var el = $(id);
  if (el) el.className = "dot" + (kind ? " " + kind : "");
}

function updateButtons() {
  if ($("convert")) $("convert").disabled = state.busy || !state.ok || !$("iscopy").checked;
  if ($("copy")) $("copy").disabled = state.busy;
  if ($("rest-copy")) $("rest-copy").disabled = state.busy;
  if ($("restore-btn")) $("restore-btn").disabled = state.busy || !state.canRestore || !getPickedSheets("rest-sheets").length;
  if ($("sh-repair")) $("sh-repair").disabled = state.busy || !getPickedSheets("sh-repair-sheets").length;
  if ($("home-repair-btn")) $("home-repair-btn").disabled = state.busy || !getPickedSheets("home-repair-sheets").length;
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
    BomSync.watch(localLog, function () { refreshMaterials(); refreshLocalFileStatus(); });
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
 * Local Master File (This Computer — Layman Friendly, No Token Needed)
 * ===================================================================== */
async function refreshLocalFileStatus() {
  var st = $("local-status");
  var tag = $("local-summary-tag");
  var homeSt = $("home-local-status");
  var homeReRow = $("home-reconnect-row");
  var shSt = $("sh-local-status"), shRe = $("sh-reconnect");
  var reBtn = $("local-reconnect"), pullBtn = $("local-pull"), saveBtn = $("local-save"), unBtn = $("local-unlink");
  var createBtn = $("local-create"), openBtn = $("local-open"), fb = $("local-fallback");

  try {
    var info = await BomSync.getLocalFileStatus();
    if (!info.supported) {
      if (createBtn) createBtn.hidden = true;
      if (openBtn) openBtn.hidden = true;
      if (fb) fb.hidden = false;
    } else {
      if (createBtn) createBtn.hidden = false;
      if (openBtn) openBtn.hidden = false;
      if (fb) fb.hidden = true;
    }

    var needsPerm = info.linked && info.permission !== "granted";
    if (reBtn) reBtn.hidden = !needsPerm;
    if (shRe) shRe.hidden = !needsPerm;
    if (homeReRow) homeReRow.hidden = !needsPerm;
    if (unBtn) unBtn.hidden = !info.linked;

    if (info.linked) {
      if (needsPerm) {
        if (st) { st.className = "status warn"; st.textContent = "Linked to \"" + info.fileName + "\" — click Reconnect once after restarting Excel."; }
        if (tag) { tag.className = "sum-meta warn"; tag.textContent = "Reconnect"; }
        if (homeSt) homeSt.textContent = "Saved materials file (\"" + info.fileName + "\") needs 1-click reconnect.";
        if (shSt) shSt.textContent = "Saved file \"" + info.fileName + "\" needs reconnect to read latest edits.";
      } else {
        if (st) { st.className = "status ok"; st.textContent = "Auto-saving to \"" + info.fileName + "\" (" + info.cachedCount + " materials)."; }
        if (tag) { tag.className = "sum-meta ok"; tag.textContent = info.fileName; }
        if (homeSt) homeSt.textContent = "Master list auto-saves to \"" + info.fileName + "\" (" + info.cachedCount + " items).";
        if (shSt) shSt.textContent = "Using saved file \"" + info.fileName + "\" (" + info.cachedCount + " items).";
      }
    } else if (info.hasCache) {
      if (st) { st.className = "status ok"; st.textContent = "Remembered on this computer (" + info.cachedCount + " materials). Pick a folder below if you also want a visible file on disk."; }
      if (tag) { tag.className = "sum-meta ok"; tag.textContent = info.cachedCount + " saved"; }
      if (homeSt) homeSt.textContent = "Master list saved on this computer (" + info.cachedCount + " items).";
      if (shSt) shSt.textContent = "Using saved materials on this computer (" + info.cachedCount + " items).";
    } else {
      if (st) { st.className = "status"; st.textContent = "Not linked to a file yet. Click \"Choose where to save...\" to start."; }
      if (tag) { tag.className = "sum-meta"; tag.textContent = "Optional"; }
      if (homeSt) homeSt.textContent = "Tip: You can save your Master list to a file on your computer under Materials.";
      if (shSt) shSt.textContent = "No saved file on this computer yet (optional — set up under Materials).";
    }

    if (pullBtn) pullBtn.disabled = state.busy || (!info.linked && !info.hasCache);
    if (saveBtn) saveBtn.disabled = state.busy || !mat.rows.length;
  } catch (e) {
    if (st) { st.className = "status warn"; st.textContent = "Could not check saved file: " + (e.message || e); }
  }
}

async function localReconnect() {
  state.busy = true;
  try {
    var ok = await BomSync.requestLocalPermission();
    localLog(ok ? "Saved file reconnected." : "Permission was not granted.", true);
    if (state.activeTab === "sheet") shLog(ok ? "Saved file reconnected." : "Permission was not granted.", true);
  } catch (e) {
    localLog("Could not reconnect: " + (e.message || e), true);
  }
  state.busy = false;
  await refreshLocalFileStatus();
}

async function localCreateFile() {
  state.busy = true; localLog("", true);
  try {
    var res = await BomSync.linkNewLocalFile(localLog);
    if (res) localLog("Done! Your materials will now auto-save to \"" + res.fileName + "\" (" + res.count + " items).");
  } catch (e) {
    if (e && e.name !== "AbortError") localLog("Could not create file: " + (e.message || e));
  }
  state.busy = false;
  await refreshLocalFileStatus();
}

async function localOpenFile() {
  state.busy = true; localLog("", true);
  try {
    var res = await BomSync.linkExistingLocalFile(localLog, { applyToWorkbook: true });
    if (res) {
      var list = await BomSync.workbookList();
      BomSync.updateWatchSnap(list);
    }
  } catch (e) {
    if (e && e.name !== "AbortError") localLog("Could not open file: " + (e.message || e));
  }
  state.busy = false;
  await refreshMaterials();
  await refreshLocalFileStatus();
}

async function localPull() {
  state.busy = true; localLog("", true);
  try {
    await BomSync.pullLocal(localLog, { allowPrompt: true, syncRemovals: true });
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
    var res = await BomSync.saveWorkbookToLocal(localLog, { allowPrompt: true });
    if (!res.savedToFile && !res.needsPermission) {
      localLog("Saved " + res.count + " materials on this computer. Click \"Choose where to save...\" to also store a visible file on disk.");
    }
  } catch (e) {
    localLog("Save failed: " + (e.message || e));
  }
  state.busy = false;
  await refreshLocalFileStatus();
}

async function localUnlink() {
  await BomSync.unlinkLocalFile();
  localLog("Disconnected local file (materials remain remembered in this computer's cache).", true);
  await refreshLocalFileStatus();
}

async function localExportDownload() {
  try {
    var list = await BomSync.workbookList();
    var json = BomSync.serialize(list);
    var blob = new Blob([json], { type: "application/json;charset=utf-8" });
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = "master-materials.json";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    localLog("Downloaded master-materials.json (" + list.length + " materials).", true);
  } catch (e) {
    localLog("Download failed: " + (e.message || e), true);
  }
}

async function localImportFile(ev) {
  var file = ev.target && ev.target.files && ev.target.files[0];
  if (!file) return;
  state.busy = true; localLog("", true);
  try {
    var text = await file.text();
    await BomSync.importLocalJsonText(text, file.name, localLog);
    var list = await BomSync.workbookList();
    BomSync.updateWatchSnap(list);
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
  $("sh-add").disabled = state.busy || !sheetInfo || (!sheetInfo.bomSheets.length && !sheetInfo.needsTemplate) || !$("sh-name").value.trim();
  updateButtons();
}

async function addSheet() {
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
        var localRes = await BomSync.pullLocal(shLog, { allowPrompt: true, syncRemovals: true });
        if (localRes && localRes.empty && BomSync.hasToken()) {
          await BomSync.pull(shLog);
        }
      } catch (e) {
        shLog("Note: could not load local materials file (" + (e.message || e) + "). Using workbook list.");
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
 * Materials
 * ===================================================================== */
var mat = { rows: [], mode: "add", selected: null };

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

function setMatMode(mode, row) {
  mat.mode = mode; mat.selected = row || null;
  resetDeleteConfirm();
  $("mat-mode").textContent = mode === "edit" ? "Edit Material" : "Add New Material";
  $("m-save").textContent = mode === "edit" ? "Save Changes" : "Add Material";
  $("m-name").readOnly = (mode === "edit");
  $("mat-name-label").textContent = mode === "edit" ? "Material name (locked)" : "Material name";
  if ($("m-delete")) $("m-delete").hidden = (mode !== "edit");
  if ($("m-new")) $("m-new").hidden = (mode !== "edit");
  var r = row || { material: "", category: "", unit: "", lin: "", area: "", piece: "", sl: "", sw: "", cost: "", remarks: "" };
  $("m-name").value = r.material; $("m-cat").value = r.category; $("m-unit").value = r.unit;
  $("m-cost").value = fmt(r.cost);
  $("m-lin").value = fmt(r.lin); $("m-area").value = fmt(r.area); $("m-piece").value = fmt(r.piece);
  $("m-sl").value = fmt(r.sl); $("m-sw").value = fmt(r.sw); $("m-rem").value = r.remarks;
  unitFields();
  renderMatList();
}

function renderMatList() {
  var box = $("mat-list"), q = $("mat-search").value.trim().toLowerCase();
  box.innerHTML = "";
  var shown = mat.rows.filter(function (r) { return !q || (r.material + " " + r.category).toLowerCase().indexOf(q) >= 0; });
  shown.slice(0, 150).forEach(function (r) {
    var d = document.createElement("div");
    d.className = "mrow" + (mat.selected && mat.selected.row === r.row ? " sel" : "");

    var left = document.createElement("span");
    left.className = "mleft";

    var v = BomTools.getMaterialValidationStatus(r);
    var icon = document.createElement("span");
    icon.className = "vicon " + (v.valid ? "ok" : "inc");
    icon.textContent = v.valid ? "✓" : "!";
    icon.title = v.valid ? "All required fields populated (Name, Category, Unit, Cost)" : ("Missing: " + v.missing.join(", "));

    var a = document.createElement("span");
    a.className = "mname";
    a.textContent = r.material;
    left.appendChild(icon);
    left.appendChild(a);

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
    var dl = $("m-cats"); dl.innerHTML = "";
    res.categories.forEach(function (c) { var o = document.createElement("option"); o.value = c; dl.appendChild(o); });
    var validCount = res.rows.filter(function (r) { return BomTools.getMaterialValidationStatus(r).valid; }).length;
    st.textContent = res.rows.length + " items · " + validCount + " complete";
    var again = keepName && res.rows.filter(function (r) { return r.material.toLowerCase() === keepName.toLowerCase(); })[0];
    if (again) setMatMode("edit", again);
    else if (mat.mode === "edit" && mat.selected) setMatMode("add");
    else renderMatList();
    $("m-save").disabled = false;
  } catch (e) {
    mat.rows = [];
    st.textContent = "No Master sheet in workbook";
    renderMatList();
    $("m-save").disabled = true;
  }
}

function readMatForm() {
  return { material: $("m-name").value, category: $("m-cat").value, unit: $("m-unit").value,
           cost: $("m-cost").value,
           lin: $("m-lin").value, area: $("m-area").value, piece: $("m-piece").value,
           sl: $("m-sl").value, sw: $("m-sw").value, remarks: $("m-rem").value };
}

async function saveMaterial() {
  var raw = readMatForm();
  matLog("", true);
  var pre = BomTools.validateMaterial(BomTools.normalizeMaterial(raw),
    mat.rows.map(function (r) { return r.material; }), mat.mode);
  if (pre.length) { pre.forEach(function (m) { matLog(m); }); return; }
  state.busy = true; $("m-save").disabled = true;
  try {
    var wasEdit = mat.mode === "edit";
    var r = wasEdit ? await BomTools.updateMaterial(raw) : await BomTools.addMaterial(raw);
    matLog((wasEdit ? "Saved \"" : "Added \"") + r.material + "\". All BOM sheets updated.");
    var list = await BomSync.workbookList();
    BomSync.updateWatchSnap(list);
    await BomSync.writeLocalMaster(list, matLog, { allowPrompt: true });
    if (BomSync.hasToken()) {
      try { await BomSync.push(matLog, { only: r.material }); }
      catch (e) { matLog("Saved locally; GitHub publish skipped (" + (e.message || e) + ")."); }
    }
    var keep = wasEdit ? r.material : null;
    state.busy = false;
    await refreshMaterials(keep);
    await refreshLocalFileStatus();
    if (!wasEdit) setMatMode("add");
  } catch (e) {
    matLog("Failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; $("m-save").disabled = false;
}

async function removeMaterial() {
  if (mat.mode !== "edit" || !mat.selected) return;
  if (!state.confirmDelete) {
    state.confirmDelete = true;
    $("m-delete").textContent = "Confirm";
    matLog("Click \"Confirm\" to permanently remove \"" + mat.selected.material + "\" from this workbook and your saved list.", true);
    return;
  }
  var target = mat.selected.material;
  state.busy = true; $("m-save").disabled = true; $("m-delete").disabled = true;
  matLog("", true);
  try {
    await BomTools.deleteMaterial(target);
    matLog("Removed \"" + target + "\".");
    var list = await BomSync.workbookList();
    BomSync.updateWatchSnap(list);
    await BomSync.writeLocalMaster(list, matLog, { allowPrompt: true });
    state.busy = false;
    setMatMode("add");
    await refreshMaterials();
    await refreshLocalFileStatus();
  } catch (e) {
    matLog("Remove failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; $("m-save").disabled = false; $("m-delete").disabled = false;
  resetDeleteConfirm();
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
      st.textContent = "No GitHub materials.json published (use Save Master List on This Computer above instead).";
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
  $("home-reconnect").onclick = localReconnect;
  $("home-repair-btn").onclick = function () { repairWorkbook(homeRepairLog, "home-repair-sheets"); };

  // Add sheet
  $("sh-add").onclick = addSheet;
  $("sh-repair").onclick = function () { repairWorkbook(shLog, "sh-repair-sheets"); };
  $("sh-reconnect").onclick = localReconnect;
  $("sh-name").oninput = updateSheetButton;
  $("sh-source").onchange = setSheetClearDefault;

  // Materials
  $("mat-search").oninput = renderMatList;
  $("m-unit").onchange = unitFields;
  $("m-save").onclick = saveMaterial;
  $("m-new").onclick = function () { matLog("", true); setMatMode("add"); };
  $("m-delete").onclick = removeMaterial;

  // Local master file controls
  $("local-reconnect").onclick = localReconnect;
  $("local-create").onclick = localCreateFile;
  $("local-open").onclick = localOpenFile;
  $("local-pull").onclick = localPull;
  $("local-save").onclick = localSave;
  $("local-unlink").onclick = localUnlink;
  $("local-export-btn").onclick = localExportDownload;
  $("local-import-btn").onclick = function () { $("local-import-input").click(); };
  $("local-import-input").onchange = localImportFile;

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
  showTab("home");
  BomSync.watch(localLog, function () { refreshMaterials(); refreshLocalFileStatus(); });
});
