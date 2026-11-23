/* taskpane.js - UI glue for the BOM Tools task pane */
"use strict";

var $ = function (id) { return document.getElementById(id); };
var state = { ok: false, busy: false, confirmDelete: false };

function log(msg) {
  var el = $("log");
  el.textContent += (el.textContent ? "\n" : "") + msg;
  el.scrollTop = el.scrollHeight;
}

function updateButtons() {
  $("convert").disabled = state.busy || !state.ok || !$("iscopy").checked;
  $("copy").disabled = state.busy;
  if ($("sh-repair")) $("sh-repair").disabled = state.busy;
}

async function check() {
  var st = $("status");
  st.className = "status";
  st.textContent = "Checking...";
  try {
    var info = await BomConvert.inspectWorkbook();
    state.ok = info.master && info.bomSheets.length > 0;
    if (state.ok) {
      st.className = "status ok";
      st.textContent = "Ready: " + info.bomSheets.length + " BOM sheet(s) found (" + info.bomSheets.join(", ") + ") and the Master sheet.";
    } else {
      st.className = "status warn";
      st.textContent = info.bomSheets.length
        ? "BOM sheets found, but the 'Master Unique Materials' sheet is missing. This looks already converted."
        : "No BOM sheets (tables named bom_...) found. Open the original automated BOM workbook.";
    }
  } catch (e) {
    state.ok = false;
    st.className = "status warn";
    st.textContent = "Could not read the workbook: " + (e.message || e);
  }
  updateButtons();
}

/* ---- Step 1: open a copy of the current file in a new Excel window ---- */
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

async function openCopy() {
  state.busy = true; updateButtons();
  $("log").textContent = "";
  try {
    log("Reading this workbook ...");
    var b64 = await getFileBase64();
    await Excel.run(async function (context) {
      Excel.createWorkbook(b64).open();
      await context.sync();
    });
    log("A copy was opened in a new window. Open BOM Tools there, tick the box and convert.");
  } catch (e) {
    log("Could not open a copy: " + (e.message || JSON.stringify(e)));
    log("Alternative: File > Save As a new name, then open that file and convert it.");
  }
  state.busy = false; updateButtons();
}

/* ---- Step 2: convert ---- */
async function convert() {
  var mode = document.querySelector('input[name="mode"]:checked').value;
  state.busy = true; updateButtons();
  $("log").textContent = "";
  try {
    var warnings = await BomConvert.convertWorkbook(mode, log, { compiled: $("compiled").checked });
    log("Done. Review the sheets, then use File > Save As to name the new file.");
    warnings.forEach(function (w) { log("WARNING: " + w); });
  } catch (e) {
    log("Conversion failed: " + (e.message || JSON.stringify(e)));
    log("Close this workbook without saving and retry on a fresh copy.");
  }
  state.busy = false;
  await check();
}


/* =====================================================================
 * Tabs & Logging
 * ===================================================================== */
var TABS = ["sheet", "mat", "convert"];

function showTab(name) {
  TABS.forEach(function (t) {
    $("tab-" + t).setAttribute("aria-selected", t === name ? "true" : "false");
    $("page-" + t).hidden = (t !== name);
  });
  try { localStorage.setItem("bomTab", name); } catch (e) { /* storage may be blocked */ }
  if (name === "convert") check();
  if (name === "sheet") { refreshSheets(); refreshLocalFileStatus(); }
  if (name === "mat") {
    refreshMaterials();
    refreshLocalFileStatus();
    refreshSync();
    BomSync.watch(localLog, function () { refreshMaterials(); refreshLocalFileStatus(); });
  }
}

function makeLog(id) {
  return function (msg, clear) {
    var el = $(id);
    if (!el) return;
    if (clear) el.textContent = "";
    if (msg) el.textContent += (el.textContent ? "\n" : "") + msg;
    el.scrollTop = el.scrollHeight;
  };
}
var shLog = makeLog("sh-log"), matLog = makeLog("mat-log"), localLog = makeLog("local-log"), syncLog = makeLog("sync-log");

/* =====================================================================
 * Local Master File (This Computer - No GitHub Token Needed)
 * ===================================================================== */
async function refreshLocalFileStatus() {
  var st = $("local-status"), dt = $("local-detail");
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
    if (unBtn) unBtn.hidden = !info.linked;

    if (info.linked) {
      if (needsPerm) {
        st.className = "status warn";
        st.textContent = "Linked to \"" + info.fileName + "\" (click Reconnect to allow file access after restart).";
        if (shSt) shSt.textContent = "Local master file: \"" + info.fileName + "\" — click Reconnect local file before adding a sheet to read the latest disk file.";
      } else {
        st.className = "status ok";
        st.textContent = "Auto-saving to local file \"" + info.fileName + "\" (" + info.cachedCount + " materials).";
        if (shSt) shSt.textContent = "Local master file active: \"" + info.fileName + "\" (" + info.cachedCount + " materials).";
      }
      dt.textContent = "Every add, edit, or removal automatically updates this file on your computer.";
    } else if (info.hasCache) {
      st.className = "status ok";
      st.textContent = "Saved on this computer (" + info.cachedCount + " materials in local add-in storage).";
      dt.textContent = info.supported
        ? "Link a .json file on your PC (Documents / OneDrive) to keep an editable file on disk too."
        : "Works across all workbooks on this computer. Use Download JSON to back up to a file.";
      if (shSt) shSt.textContent = "Local computer master list ready (" + info.cachedCount + " materials).";
    } else {
      st.className = "status";
      st.textContent = "No local master file linked yet.";
      dt.textContent = info.supported
        ? "Click \"Save new file...\" to save this workbook's Master list to your computer, or \"Link existing file...\" to pick one."
        : "Edit any material or click Import JSON file to store your master list on this computer.";
      if (shSt) shSt.textContent = "No local master file saved yet. Link one in the Materials tab to auto-update new BOMs.";
    }

    if (pullBtn) pullBtn.disabled = state.busy || (!info.linked && !info.hasCache);
    if (saveBtn) saveBtn.disabled = state.busy || !mat.rows.length;
  } catch (e) {
    if (st) { st.className = "status warn"; st.textContent = "Could not check local master file: " + (e.message || e); }
  }
}

async function localReconnect() {
  state.busy = true;
  try {
    var ok = await BomSync.requestLocalPermission();
    localLog(ok ? "Local file reconnected." : "Permission was not granted.", true);
    if ($("sh-log") && !$("page-sheet").hidden) shLog(ok ? "Local file reconnected." : "Permission was not granted.", true);
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
    if (res) localLog("Linked \"" + res.fileName + "\" (" + res.count + " materials).");
  } catch (e) {
    if (e && e.name !== "AbortError") localLog("Failed: " + (e.message || e));
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
    if (e && e.name !== "AbortError") localLog("Failed: " + (e.message || e));
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
    localLog("Failed: " + (e.message || e));
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
      localLog("Saved " + res.count + " materials to local computer storage. Click \"Save new file...\" to also write to a .json file on disk.");
    }
  } catch (e) {
    localLog("Failed: " + (e.message || e));
  }
  state.busy = false;
  await refreshLocalFileStatus();
}

async function localUnlink() {
  await BomSync.unlinkLocalFile();
  localLog("Unlinked local disk file (local computer cache still kept).", true);
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
 * Add sheet
 * ===================================================================== */
var sheetInfo = null;

function setSheetClearDefault() {
  $("sh-clear").checked = $("sh-source").value !== BomTools.BLANK;
}

async function refreshSheets() {
  var st = $("sh-status");
  st.className = "status"; st.textContent = "Checking...";
  try {
    sheetInfo = await BomTools.inspectSheets();
    var sel = $("sh-source"), keep = sel.value;
    sel.innerHTML = "";
    if (sheetInfo.needsTemplate) {
      var o0 = document.createElement("option");
      o0.value = BomTools.BLANK; o0.textContent = "BOM Blank (imported from the template)";
      sel.appendChild(o0);
    }
    sheetInfo.bomSheets.forEach(function (s) {
      var o = document.createElement("option");
      o.value = s.name; o.textContent = s.name + (s.hidden ? " (hidden)" : "");
      sel.appendChild(o);
    });
    var names = sheetInfo.bomSheets.map(function (s) { return s.name; });
    sel.value = names.indexOf(keep) >= 0 ? keep : (sheetInfo.hasBlank || sheetInfo.needsTemplate ? BomTools.BLANK : names[0]);
    setSheetClearDefault();
    if (!$("sh-project").value && sheetInfo.project) $("sh-project").value = sheetInfo.project;
    if ($("sh-repair")) $("sh-repair").hidden = !(sheetInfo.hasMaster && names.length > 0);
    if (sheetInfo.needsTemplate) {
      st.className = "status ok";
      st.textContent = "No BOM template in this workbook yet. Adding a sheet will import Master Unique Materials + hidden BOM Blank and sync your local master list.";
    } else if (names.length) {
      st.className = "status ok";
      st.textContent = names.length + " BOM sheet(s) found" + (sheetInfo.hasBlank ? " incl. the blank sheet." : ". No \"BOM Blank\" sheet yet; you can copy any sheet and clear its data.");
    } else {
      st.className = "status warn";
      st.textContent = "No BOM sheets (tables named bom_...) found. Use this on the automated template or project workbook, not a converted one.";
    }
  } catch (e) {
    sheetInfo = null;
    if ($("sh-repair")) $("sh-repair").hidden = true;
    st.className = "status warn";
    st.textContent = "Could not read the workbook: " + (e.message || e);
  }
  updateSheetButton();
}

function updateSheetButton() {
  $("sh-add").disabled = state.busy || !sheetInfo || (!sheetInfo.bomSheets.length && !sheetInfo.needsTemplate) || !$("sh-name").value.trim();
  if ($("sh-repair")) $("sh-repair").disabled = state.busy;
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
        shLog("WARNING: could not sync local master list (" + (e.message || e) + "). Continuing with the list in this workbook.");
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
    shLog("Could not add the sheet: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; updateButtons();
  await refreshSheets();
  await refreshLocalFileStatus();
}

async function repairWorkbook() {
  state.busy = true; updateSheetButton(); updateButtons();
  shLog("", true);
  try {
    await BomTools.repairWorkbook(shLog);
  } catch (e) {
    shLog("Repair failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; updateButtons();
  await refreshSheets();
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
  $("mat-mode").textContent = mode === "edit" ? "Edit material" : "New material";
  $("m-save").textContent = mode === "edit" ? "Save changes" : "Add material";
  $("m-name").readOnly = (mode === "edit");
  $("mat-name-label").textContent = mode === "edit" ? "Material name (locked)" : "Material name";
  if ($("m-delete")) $("m-delete").hidden = (mode !== "edit");
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
    var e = document.createElement("div"); e.className = "mrow"; e.textContent = "No match."; e.style.cursor = "default"; box.appendChild(e);
  }
}

async function refreshMaterials(keepName) {
  var st = $("mat-status");
  st.className = "status"; st.textContent = "Checking...";
  try {
    var res = await BomTools.listMaterials();
    mat.rows = res.rows;
    var dl = $("m-cats"); dl.innerHTML = "";
    res.categories.forEach(function (c) { var o = document.createElement("option"); o.value = c; dl.appendChild(o); });
    var validCount = res.rows.filter(function (r) { return BomTools.getMaterialValidationStatus(r).valid; }).length;
    st.className = "status ok";
    st.textContent = res.rows.length + " materials (" + validCount + " complete).";
    var again = keepName && res.rows.filter(function (r) { return r.material.toLowerCase() === keepName.toLowerCase(); })[0];
    if (again) setMatMode("edit", again);
    else if (mat.mode === "edit" && mat.selected) setMatMode("add");
    else renderMatList();
    $("m-save").disabled = false;
  } catch (e) {
    mat.rows = [];
    st.className = "status warn";
    st.textContent = "Could not read the master list: " + (e.message || e);
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
    matLog((wasEdit ? "Saved \"" : "Added \"") + r.material + "\". Take-off sheets update automatically.");
    var list = await BomSync.workbookList();
    BomSync.updateWatchSnap(list);
    await BomSync.writeLocalMaster(list, matLog, { allowPrompt: true });
    if (BomSync.hasToken()) {
      try { await BomSync.push(matLog, { only: r.material }); }
      catch (e) { matLog("Saved locally, but GitHub publish failed: " + (e.message || e)); }
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
    $("m-delete").textContent = "Confirm remove";
    matLog("Click \"Confirm remove\" again to delete \"" + mat.selected.material + "\" from this workbook and local master file.", true);
    return;
  }
  var target = mat.selected.material;
  state.busy = true; $("m-save").disabled = true; $("m-delete").disabled = true;
  matLog("", true);
  try {
    await BomTools.deleteMaterial(target);
    matLog("Removed \"" + target + "\" from the master list.");
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
 * Optional: Shared list (GitHub)
 * ===================================================================== */
function nameList(arr, pick) {
  var n = arr.map(pick);
  return n.slice(0, 8).join(", ") + (n.length > 8 ? " and " + (n.length - 8) + " more" : "");
}

async function refreshSync() {
  var st = $("sync-status"), dt = $("sync-detail");
  if (!st) return;
  st.className = "status"; st.textContent = "Checking GitHub ...";
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
      st.textContent = "No GitHub materials.json published (optional — use Local master file above instead).";
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
    st.textContent = "GitHub sync unavailable (" + (e.message || e) + "). Local master file works offline.";
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

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    $("status").textContent = "This add-in works in Excel only.";
    return;
  }
  // convert page
  $("refresh").onclick = check;
  $("copy").onclick = openCopy;
  $("convert").onclick = convert;
  $("iscopy").onchange = updateButtons;
  // add sheet page
  $("sh-add").onclick = addSheet;
  if ($("sh-repair")) $("sh-repair").onclick = repairWorkbook;
  if ($("sh-reconnect")) $("sh-reconnect").onclick = localReconnect;
  $("sh-name").oninput = updateSheetButton;
  $("sh-source").onchange = setSheetClearDefault;
  // local master file controls
  $("local-reconnect").onclick = localReconnect;
  $("local-create").onclick = localCreateFile;
  $("local-open").onclick = localOpenFile;
  $("local-pull").onclick = localPull;
  $("local-save").onclick = localSave;
  $("local-unlink").onclick = localUnlink;
  $("local-export-btn").onclick = localExportDownload;
  $("local-import-btn").onclick = function () { $("local-import-input").click(); };
  $("local-import-input").onchange = localImportFile;
  // materials page
  $("mat-search").oninput = renderMatList;
  $("m-unit").onchange = unitFields;
  $("m-save").onclick = saveMaterial;
  $("m-new").onclick = function () { matLog("", true); setMatMode("add"); };
  $("m-delete").onclick = removeMaterial;
  // optional GitHub sync
  $("sync-pull").onclick = syncPull;
  $("sync-push").onclick = syncPush;
  $("gh-save").onclick = function () {
    var ok = BomSync.setToken($("gh-token").value);
    $("gh-token").value = "";
    syncLog(ok ? "Token saved on this computer." : "Could not store the token (storage is blocked).", true);
    refreshSync();
  };
  $("gh-clear").onclick = function () { BomSync.setToken(""); syncLog("Token removed.", true); refreshSync(); };
  // tabs
  TABS.forEach(function (t) { $("tab-" + t).onclick = function () { showTab(t); }; });
  var first = "convert";
  try { var s = localStorage.getItem("bomTab"); if (TABS.indexOf(s) >= 0) first = s; } catch (e) { /* ignore */ }
  showTab(first);
  BomSync.watch(localLog, function () { refreshMaterials(); refreshLocalFileStatus(); });
});
