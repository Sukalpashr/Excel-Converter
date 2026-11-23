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
    var warnings = await BomConvert.convertWorkbook(mode, log);
    log("Done. Review the sheets, then use File > Save As to name the new file.");
    warnings.forEach(function (w) { log("WARNING: " + w); });
  } catch (e) {
    log("Conversion failed: " + (e.message || JSON.stringify(e)));
    log("Close this workbook without saving and retry on a fresh copy.");
  }
  state.busy = false;
  await check();
}

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    $("status").textContent = "This add-in works in Excel only.";
    return;
  }
  $("refresh").onclick = check;
  $("copy").onclick = openCopy;
  $("convert").onclick = convert;
  $("iscopy").onchange = updateButtons;
  check();
});
