/* taskpane.js - UI glue for the BOM Tools task pane */
"use strict";

var $ = function (id) { return document.getElementById(id); };
var state = { ok: false, busy: false };

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
 * Tabs
 * ===================================================================== */
var TABS = ["sheet", "mat", "convert"];

function showTab(name) {
  TABS.forEach(function (t) {
    $("tab-" + t).setAttribute("aria-selected", t === name ? "true" : "false");
    $("page-" + t).hidden = (t !== name);
  });
  try { localStorage.setItem("bomTab", name); } catch (e) { /* storage may be blocked */ }
  if (name === "convert") check();
  if (name === "sheet") refreshSheets();
  if (name === "mat") { refreshMaterials(); refreshSync(); BomSync.watch(syncLog); }
}

function makeLog(id) {
  return function (msg, clear) {
    var el = $(id);
    if (clear) el.textContent = "";
    el.textContent += (el.textContent ? "\n" : "") + msg;
    el.scrollTop = el.scrollHeight;
  };
}
var shLog = makeLog("sh-log"), matLog = makeLog("mat-log");

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
      st.textContent = "No BOM template in this workbook yet. Adding a sheet will import the master list and blank BOM sheet first.";
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
      try { await BomSync.pull(shLog); }
      catch (e) { shLog("WARNING: could not update the master list from GitHub (" + (e.message || e) + "). Continuing with the list in this workbook."); }
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

function setMatMode(mode, row) {
  mat.mode = mode; mat.selected = row || null;
  $("mat-mode").textContent = mode === "edit" ? "Edit material" : "New material";
  $("m-save").textContent = mode === "edit" ? "Save changes" : "Add material";
  $("m-name").readOnly = (mode === "edit");
  $("mat-name-label").textContent = mode === "edit" ? "Material name (locked)" : "Material name";
  var r = row || { material: "", category: "", unit: "", lin: "", area: "", piece: "", sl: "", sw: "", remarks: "" };
  $("m-name").value = r.material; $("m-cat").value = r.category; $("m-unit").value = r.unit;
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
    if (mat.selected && mat.selected.row === r.row) d.className = "sel";
    var a = document.createElement("span"); a.textContent = r.material;
    var b = document.createElement("small");
    var w = r.unit === "kg/m" ? r.lin : r.unit === "kg/m²" ? r.area : r.piece;
    b.textContent = fmt(w) + " " + r.unit;
    d.appendChild(a); d.appendChild(b);
    d.onclick = function () { setMatMode("edit", r); matLog("", true); };
    box.appendChild(d);
  });
  if (!shown.length) {
    var e = document.createElement("div"); e.textContent = "No match."; e.style.cursor = "default"; box.appendChild(e);
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
    st.className = "status ok";
    st.textContent = res.rows.length + " materials in the master list.";
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
    if (BomSync.hasToken()) {
      try { await BomSync.push(matLog, { only: r.material }); }
      catch (e) { matLog("Saved here, but publishing to GitHub failed: " + (e.message || e)); }
    }
    var keep = wasEdit ? r.material : null;
    state.busy = false;
    await refreshMaterials(keep);
    if (!wasEdit) setMatMode("add");
  } catch (e) {
    matLog("Failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; $("m-save").disabled = false;
}

/* =====================================================================
 * Shared list (GitHub)
 * ===================================================================== */
var syncLog = makeLog("sync-log");

function nameList(arr, pick) {
  var n = arr.map(pick);
  return n.slice(0, 8).join(", ") + (n.length > 8 ? " and " + (n.length - 8) + " more" : "");
}

async function refreshSync() {
  var st = $("sync-status"), dt = $("sync-detail");
  st.className = "status"; st.textContent = "Comparing with GitHub ...";
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
      st.className = "status warn";
      st.textContent = "No shared list on GitHub yet." + (BomSync.hasToken() ? " Publish to create it from this workbook." : "");
    } else if (!canPull && !canPush) {
      st.className = "status ok";
      st.textContent = "In sync with GitHub (" + d.remoteCount + " materials).";
    } else {
      st.className = "status warn";
      st.textContent = (canPull ? "GitHub has updates for this workbook. " : "") + (canPush ? "This workbook has changes GitHub does not." : "");
    }
    $("sync-pull").disabled = !canPull || state.busy;
    $("sync-push").disabled = !canPush || state.busy || !BomSync.hasToken();
  } catch (e) {
    st.className = "status warn";
    st.textContent = "Could not compare with GitHub: " + (e.message || e);
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
  $("sh-name").oninput = updateSheetButton;
  $("sh-source").onchange = setSheetClearDefault;
  // materials page
  $("mat-search").oninput = renderMatList;
  $("m-unit").onchange = unitFields;
  $("m-save").onclick = saveMaterial;
  $("m-new").onclick = function () { matLog("", true); setMatMode("add"); };
  $("sync-pull").onclick = syncPull;
  $("sync-push").onclick = syncPush;
  $("gh-save").onclick = function () {
    var ok = BomSync.setToken($("gh-token").value);
    $("gh-token").value = "";
    syncLog(ok ? "Token saved on this computer." : "Could not store the token (storage is blocked).", true);
    refreshSync();
    if (ok) BomSync.watch(syncLog);
  };
  $("gh-clear").onclick = function () { BomSync.setToken(""); syncLog("Token removed.", true); refreshSync(); };
  // tabs
  TABS.forEach(function (t) { $("tab-" + t).onclick = function () { showTab(t); }; });
  var first = "convert";
  try { var s = localStorage.getItem("bomTab"); if (TABS.indexOf(s) >= 0) first = s; } catch (e) { /* ignore */ }
  showTab(first);
  if (BomSync.hasToken()) BomSync.watch(syncLog);
});
