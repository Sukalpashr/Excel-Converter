/* taskpane.js - UI glue for the BOM Converter task pane */
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
    log("A copy was opened in a new window. Open BOM Converter there, tick the box and convert.");
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
var autoNote = "";            // what the silent start-up sync did, shown on the Add sheet / Materials pages

function showTab(name) {
  TABS.forEach(function (t) {
    $("tab-" + t).setAttribute("aria-selected", t === name ? "true" : "false");
    $("page-" + t).hidden = (t !== name);
  });
  try { localStorage.setItem("bomTab", name); } catch (e) { /* storage may be blocked */ }
  if (name === "convert") check();
  if (name === "sheet") refreshSheets();
  if (name === "mat") refreshMaterials();
}

function makeLog(id) {
  return function (msg, clear) {
    var el = $(id);
    if (clear) el.textContent = "";
    if (msg) el.textContent += (el.textContent ? "\n" : "") + msg;
    el.scrollTop = el.scrollHeight;
  };
}
var shLog = makeLog("sh-log"), matLog = makeLog("mat-log"), syncLog = makeLog("sync-log");

function plural(n, w) { return n + " " + w + (n === 1 ? "" : "s"); }
function when(ts) { return ts ? new Date(ts).toLocaleString() : "never"; }

function storeLine(st) {
  if (!st) return "";
  var s = "Latest list on this computer: " + plural(st.storeCount, "material") + ", updated " + when(st.storeAt) + ".";
  if (!st.persistent) s += " (This browser blocks local storage, so the list is only kept while this panel stays open.)";
  return s;
}

/* =====================================================================
 * Add sheet
 * ===================================================================== */
var sheetInfo = null;

function setSheetClearDefault() {
  var v = $("sh-source").value;
  $("sh-clear").checked = v !== BomTools.BLANK && v !== BomTools.TEMPLATE_SRC;
  $("sh-clear").disabled = v === BomTools.TEMPLATE_SRC;
}

async function refreshSheets() {
  var st = $("sh-status"), ms = $("sh-master");
  st.className = "status"; st.textContent = "Checking...";
  try {
    sheetInfo = await BomTools.inspectSheets();
    var ms0 = await BomTools.masterStatus({});
    var sel = $("sh-source"), keep = sel.value;
    sel.innerHTML = "";
    var add = function (v, label, disabled) {
      var o = document.createElement("option"); o.value = v; o.textContent = label; o.disabled = !!disabled; sel.appendChild(o);
    };
    add(BomTools.TEMPLATE_SRC, "Built-in template (blank)" + (sheetInfo.canTemplate ? "" : " - needs Excel 2021 / 365"), !sheetInfo.canTemplate);
    sheetInfo.bomSheets.forEach(function (s) { add(s.name, s.name + (s.hidden ? " (hidden)" : "")); });
    var names = sheetInfo.bomSheets.map(function (s) { return s.name; });
    var wanted = [keep, sheetInfo.hasBlank ? BomTools.BLANK : "", sheetInfo.canTemplate ? BomTools.TEMPLATE_SRC : "", names[0]];
    for (var i = 0; i < wanted.length; i++) {
      var opt = wanted[i] && Array.prototype.filter.call(sel.options, function (o) { return o.value === wanted[i] && !o.disabled; })[0];
      if (opt) { sel.value = wanted[i]; break; }
    }
    setSheetClearDefault();
    if (!$("sh-project").value && sheetInfo.project) $("sh-project").value = sheetInfo.project;

    st.className = "status ok";
    st.textContent = names.length
      ? plural(names.length, "BOM sheet") + " in this workbook."
      : "No BOM sheets in this workbook yet.";
    if (ms0.has) {
      var extra = ms0.pendingNew ? " " + plural(ms0.pendingNew, "newer material") + " from the latest list " + (ms0.pendingNew === 1 ? "is" : "are") + " not in it yet (see the Materials tab)." : "";
      ms.className = "status ok";
      ms.textContent = "Master list: " + plural(ms0.rows, "material") + " in this workbook." + extra + (autoNote ? " " + autoNote : "");
    } else {
      ms.className = "status";
      ms.textContent = "No master list in this workbook. Adding a sheet from the built-in template will add one too, filled with your latest materials (" + plural(ms0.storeCount, "material") + " saved on this computer).";
    }
    $("sh-note").textContent = sheetInfo.canTemplate
      ? "Tip: \"Built-in template\" always works. To copy a sheet you have already set up, pick it in the list instead."
      : "This Excel version cannot insert the built-in template. Copy an existing BOM sheet instead, or use Excel 2021 / Microsoft 365.";
  } catch (e) {
    sheetInfo = null;
    st.className = "status warn";
    st.textContent = "Could not read the workbook: " + (e.message || e);
  }
  updateSheetButton();
}

function updateSheetButton() {
  $("sh-add").disabled = state.busy || !sheetInfo || !$("sh-source").value || !$("sh-name").value.trim();
}

async function addSheet() {
  state.busy = true; updateSheetButton(); updateButtons();
  shLog("", true);
  try {
    var r = await BomTools.addBomSheet({
      name: $("sh-name").value, project: $("sh-project").value,
      source: $("sh-source").value, clear: $("sh-clear").checked
    }, shLog);
    shLog("Added sheet \"" + r.name + "\".");
    if (r.master) shLog("Master list added with " + plural(r.master.total, "material") + " (latest list applied).");
    r.warnings.forEach(function (w) { shLog("WARNING: " + w); });
    $("sh-name").value = "";
  } catch (e) {
    shLog("Could not add the sheet: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; updateButtons();
  await refreshSheets();
}

/* =====================================================================
 * Materials
 * ===================================================================== */
var mat = { rows: [], mode: "add", selected: null, none: false };

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
    var e = document.createElement("div"); e.textContent = mat.none ? "No master list in this workbook." : "No match."; e.style.cursor = "default"; box.appendChild(e);
  }
}

// the "latest master list" card: only shown when there is something to do
function renderSync(st) {
  var card = $("mat-sync");
  card.hidden = true;
  $("sync-apply").hidden = true; $("sync-addmaster").hidden = true; $("sync-allrow").hidden = true;
  if (!st.has) {
    card.hidden = false;
    $("sync-title").textContent = "No master list here";
    $("sync-text").className = "status";
    $("sync-text").textContent = "This workbook has no master list, so the BOM sheets cannot look up weights. Add it from the built-in template; it is filled with your latest materials (" + plural(st.storeCount, "material") + ").";
    $("sync-addmaster").hidden = false;
  } else if (st.pendingNew || st.pendingDiff) {
    card.hidden = false;
    $("sync-title").textContent = "Latest master list";
    $("sync-text").className = "status";
    var parts = [];
    if (st.pendingNew) parts.push(plural(st.pendingNew, "material") + " from the latest list " + (st.pendingNew === 1 ? "is" : "are") + " missing here");
    if (st.pendingDiff) parts.push(plural(st.pendingDiff, "material") + " differ" + (st.pendingDiff === 1 ? "s" : "") + " from the latest list");
    $("sync-text").textContent = parts.join("; ") + ". Nothing is changed in this workbook until you update.";
    $("sync-apply").hidden = false;
    $("sync-allrow").hidden = !st.pendingDiff;
    $("sync-all").checked = false;
  }
  var line = storeLine(st);
  if (autoNote) line = autoNote + " " + line;
  if (st.conflicts && st.conflicts.length) line += " Edited in two places, kept as they are: " + st.conflicts.join(", ") + ".";
  $("mat-store").textContent = line;
}

async function refreshMaterials(keepName) {
  var st = $("mat-status");
  st.className = "status"; st.textContent = "Checking...";
  try {
    var s = await BomTools.masterStatus({});
    renderSync(s);
    var res = await BomTools.listMaterials();
    mat.none = !!res.none;
    mat.rows = res.rows;
    var dl = $("m-cats"); dl.innerHTML = "";
    res.categories.forEach(function (c) { var o = document.createElement("option"); o.value = c; dl.appendChild(o); });
    if (res.none) {
      st.className = "status warn";
      st.textContent = "No master list in this workbook.";
      $("m-save").disabled = true;
      renderMatList();
      return;
    }
    st.className = "status ok";
    st.textContent = plural(res.rows.length, "material") + " in this workbook's master list.";
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

async function syncApply() {
  state.busy = true; $("sync-apply").disabled = true;
  syncLog("", true);
  try {
    var s = await BomTools.applyLatest($("sync-all").checked ? "all" : "new");
    autoNote = "";
    state.busy = false;
    await refreshMaterials();
    $("sync-title").textContent = "Master list updated"; $("sync-text").textContent = "";
    syncLog("Added " + plural(s.applied.added, "material") + (s.applied.updated ? ", updated " + s.applied.updated : "") + ".");
    $("mat-sync").hidden = false;
  } catch (e) {
    syncLog("Failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; $("sync-apply").disabled = false;
}

async function addMasterSheet() {
  state.busy = true; $("sync-addmaster").disabled = true;
  syncLog("", true);
  try {
    var r = await BomTools.addMasterSheet(syncLog);
    r.warnings.forEach(function (w) { syncLog("WARNING: " + w); });
    autoNote = "";
    state.busy = false;
    await refreshMaterials();
    $("sync-title").textContent = "Master list added"; $("sync-text").textContent = "";
    syncLog("Master list sheet added with " + plural(r.master.total, "material") + ".");
    $("mat-sync").hidden = false;
  } catch (e) {
    syncLog("Could not add the master list: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; $("sync-addmaster").disabled = false;
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
    matLog((wasEdit ? "Saved \"" : "Added \"") + r.material + "\". Take-off sheets update automatically, and the latest list on this computer now has it.");
    var keep = wasEdit ? r.material : null;
    state.busy = false;
    await refreshMaterials(keep);
    if (!wasEdit) setMatMode("add");
  } catch (e) {
    matLog("Failed: " + (e.message || JSON.stringify(e)));
  }
  state.busy = false; $("m-save").disabled = mat.none;
}

Office.onReady(async function (info) {
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
  $("sh-name").oninput = updateSheetButton;
  $("sh-source").onchange = function () { setSheetClearDefault(); updateSheetButton(); };
  // materials page
  $("mat-search").oninput = renderMatList;
  $("m-unit").onchange = unitFields;
  $("m-save").onclick = saveMaterial;
  $("m-new").onclick = function () { matLog("", true); setMatMode("add"); };
  $("sync-apply").onclick = syncApply;
  $("sync-addmaster").onclick = addMasterSheet;
  // tabs
  TABS.forEach(function (t) { $("tab-" + t).onclick = function () { showTab(t); }; });
  var first = "convert";
  try { var s = localStorage.getItem("bomTab"); if (TABS.indexOf(s) >= 0) first = s; } catch (e) { /* ignore */ }

  // silent start-up sync: remember this workbook's materials; refresh the master of a brand-new (empty) workbook
  try {
    var r = await BomTools.masterStatus({ auto: true });
    if (r.applied.added || r.applied.updated) {
      autoNote = "This new workbook's master list was refreshed from your latest list (" +
        (r.applied.added ? plural(r.applied.added, "material") + " added" : "") +
        (r.applied.added && r.applied.updated ? ", " : "") +
        (r.applied.updated ? plural(r.applied.updated, "material") + " updated" : "") + ").";
    }
  } catch (e) { /* no master here, or storage blocked: nothing to do */ }
  showTab(first);
});
