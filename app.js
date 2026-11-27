/* app.js - Lightweight BOM Tools 3.0v patch2 Task Pane Controller */
(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const lc = (s) => String(s ?? "").trim().toLowerCase();

  const state = {
    tab: "home",
    convMode: "flatten",
    wb: {
      hasMaster: false, hasBlank: false, needsTemplate: false,
      projectTitle: "", projectMode: null,
      allSheets: [], bomSheets: [], convertedSheets: []
    },
    materials: [],
    categories: [],
    savedPool: [],
    projPickSelected: new Set(),
    projAddSelected: new Set(),
    globalCount: 0,
    globalUpdated: "",
    savedProjects: [],
    editingName: null,
    deleteArm: false,
    switchGlobalArm: false,
    switchGlobalTimer: null,
    pickedRestProj: null,
    recentNames: new Map(), // lc(name) -> { type, time }
    sessionEvents: [],
    toastTimer: null,
    toastPaused: false,
    busy: false
  };

  /* ---------------- UI Helpers ---------------- */
  function showGlobalMsg(text, kind = "ok") {
    const el = $("globalMsg");
    if (!text) { el.classList.add("hidden"); return; }
    el.textContent = text;
    el.className = `hint ${kind}`;
    clearTimeout(showGlobalMsg._t);
    showGlobalMsg._t = setTimeout(() => el.classList.add("hidden"), 4200);
  }

  function appendConvLog(line) {
    const box = $("convLog");
    box.classList.remove("hidden");
    box.textContent += (box.textContent ? "\n" : "") + line;
    box.scrollTop = box.scrollHeight;
  }

  function clearConvLog() {
    const box = $("convLog");
    box.textContent = "";
    box.classList.add("hidden");
  }

  function setBusy(flag) {
    state.busy = flag;
    $("btnRefresh").disabled = flag;
    updateAllGuards();
  }

  function switchTab(tab, sub) {
    state.tab = tab;
    document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === tab));
    document.querySelectorAll(".panel").forEach((p) => p.classList.toggle("hidden", p.id !== `tab-${tab}`));
    if (tab === "convert" && (sub === "flatten" || sub === "restore")) {
      setConvMode(sub);
    }
    updateAllGuards();
    // On-demand direct-edit detection ONLY when user visits Materials (zero ws.onChanged hooks!)
    if (tab === "mat" && state.wb.hasMaster && !state.busy && typeof Excel !== "undefined") {
      checkDirectMasterEdits();
    }
  }

  function setConvMode(mode) {
    state.convMode = mode;
    $("segFlatten").classList.toggle("active", mode === "flatten");
    $("segRestore").classList.toggle("active", mode === "restore");
    $("viewFlatten").classList.toggle("hidden", mode !== "flatten");
    $("viewRestore").classList.toggle("hidden", mode !== "restore");
    updateConvertGuards();
  }

  function renderChecklist(containerId, items, checkedDefault = true) {
    const box = $(containerId);
    const prevChecked = new Set(
      [...box.querySelectorAll('input[type="checkbox"]:checked')].map((i) => i.value)
    );
    const hadItems = box.querySelectorAll('input[type="checkbox"]').length > 0;
    box.innerHTML = "";
    if (!items.length) {
      box.innerHTML = '<div class="sub">None detected.</div>';
      return;
    }
    for (const item of items) {
      const val = typeof item === "string" ? item : item.value;
      const label = typeof item === "string" ? item : item.label;
      const checked = hadItems ? prevChecked.has(val) : checkedDefault;
      const row = document.createElement("label");
      row.className = "check-row";
      row.innerHTML = `<input type="checkbox" value="${escHtml(val)}" ${checked ? "checked" : ""}/><span>${escHtml(label)}</span>`;
      box.appendChild(row);
    }
  }

  function getChecked(containerId) {
    return [...$(containerId).querySelectorAll('input[type="checkbox"]:checked')].map((i) => i.value);
  }

  function escHtml(s) {
    return String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }

  /* ---------------- Auto-collapsing Change Toast & Session Log ---------------- */
  function recordChangeEvent(evt) {
    if (!evt || !evt.items || !evt.items.length) return;
    const nowTime = new Date(evt.time || Date.now()).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit", second: "2-digit" });
    evt.items.forEach((it) => {
      const name = typeof it.material === "string" ? it.material : it.material?.material || "";
      if (!name) return;
      state.recentNames.set(lc(name), { type: it.type, time: nowTime });
      state.sessionEvents.unshift({
        material: name,
        type: it.type,
        fields: it.fields || [],
        source: evt.source || "panel",
        time: nowTime
      });
    });
    if (state.sessionEvents.length > 40) state.sessionEvents.length = 40;
    renderSessionLog();
    showChangeToast(evt, nowTime);
  }

  function showChangeToast(evt, timeStr) {
    const toast = $("changeToast");
    const first = evt.items[0];
    const kind = first.type === "added" ? "Added" : first.type === "removed" ? "Removed" : "Modified";
    $("toastTitle").textContent = evt.items.length === 1 ? `Material ${kind}` : `${evt.items.length} Materials Updated`;
    $("toastSub").textContent = `${evt.source === "sheet" ? "Detected on Master sheet" : "Updated in panel"} · ${timeStr}`;

    const itemsBox = $("toastItems");
    itemsBox.innerHTML = "";
    evt.items.slice(0, 3).forEach((it) => {
      const name = typeof it.material === "string" ? it.material : it.material?.material || "";
      const div = document.createElement("div");
      div.className = "toast-item";
      const lbl = it.type === "added" ? "Added" : it.type === "removed" ? "Removed" : "Modified";
      const extra = it.fields && it.fields.length ? ` · Changed: ${it.fields.join(", ")}` : "";
      div.innerHTML = `<strong>${escHtml(name)}</strong> <span style="color:var(--blue);font-weight:700;">${lbl}</span><span class="sub">${escHtml(extra)}</span>`;
      if (it.type !== "removed") {
        div.addEventListener("click", () => {
          const found = state.materials.find((m) => lc(m.material) === lc(name));
          if (found) loadMaterialIntoForm(found);
          collapseToast();
        });
      }
      itemsBox.appendChild(div);
    });

    toast.classList.remove("hidden");
    $("sessionWrap").classList.add("hidden");
    startToastTimer();
  }

  function startToastTimer() {
    clearTimeout(state.toastTimer);
    const bar = $("toastBarFill");
    bar.style.transition = "none";
    bar.style.transform = "scaleX(1)";
    requestAnimationFrame(() => {
      bar.style.transition = "transform 5.5s linear";
      bar.style.transform = "scaleX(0)";
    });
    state.toastTimer = setTimeout(() => {
      if (!state.toastPaused) collapseToast();
    }, 5500);
  }

  function collapseToast() {
    clearTimeout(state.toastTimer);
    $("changeToast").classList.add("hidden");
    if (state.sessionEvents.length) $("sessionWrap").classList.remove("hidden");
  }

  function renderSessionLog() {
    const n = state.sessionEvents.length;
    if (!n) { $("sessionWrap").classList.add("hidden"); return; }
    $("sessionCount").textContent = String(n);
    const top = state.sessionEvents[0];
    $("sessionPreview").textContent = `${window.BomSheets.actionLabel(top.type)} ${top.material}`;
    const log = $("sessionLog");
    log.innerHTML = state.sessionEvents.map((e) =>
      `<div class="session-row"><span><strong>${escHtml(e.material)}</strong> · ${escHtml(window.BomSheets.actionLabel(e.type))}</span><span class="sub">${escHtml(e.time)}</span></div>`
    ).join("");
  }

  /* ---------------- Zero-Hook On-Demand Direct Sheet Edit Detection ---------------- */
  async function checkDirectMasterEdits() {
    if (!state.wb.hasMaster || typeof Excel === "undefined") return false;
    const res = await window.BomSheets.syncDirectMasterEdits((evt) => {
      recordChangeEvent(evt);
    });
    if (!res || !res.detected) return false;
    state.materials = res.cur.rows;
    state.categories = res.cur.categories;
    state.wb.projectMode = res.cur.projectMode;
    if (!res.cur.projectMode) {
      state.globalCount = res.cur.rows.length;
      state.globalUpdated = res.evt.time;
    } else {
      state.savedProjects = await window.BomStore.listSavedProjects();
    }
    state.savedPool = await window.BomStore.getAvailableMaterialPool(state.materials);
    renderMaterials();
    renderHome();
    renderSavedProjects();
    updateAllGuards();
    return true;
  }

  /* ---------------- Data Refresh ---------------- */
  async function refreshAll() {
    try {
      const [offline, globalMaster, savedProjs] = await Promise.all([
        window.BomStore.detectOfflineApi(),
        window.BomStore.readGlobalMaster(),
        window.BomStore.listSavedProjects()
      ]);
      state.globalCount = globalMaster.materials.length;
      state.globalUpdated = globalMaster.updated || "";
      state.savedProjects = savedProjs;

      const verTag = $("verTag");
      verTag.textContent = offline ? `${window.BomStore.VERSION} · Offline` : window.BomStore.VERSION;
      verTag.classList.toggle("offline", offline);

      if (typeof Excel !== "undefined") {
        state.wb = await window.BomSheets.inspectWorkbook();
        if (state.wb.hasMaster) {
          // Detect & persist any direct edits typed on Master sheet before refreshing view
          const direct = await window.BomSheets.syncDirectMasterEdits((evt) => recordChangeEvent(evt));
          const matData = direct?.cur || (await window.BomSheets.listMaterials());
          state.materials = matData.rows;
          state.categories = matData.categories;
          state.wb.projectMode = matData.projectMode;
          window.BomSheets.setMasterSnap(state.materials);

          if (direct?.detected) {
            if (!matData.projectMode) {
              state.globalCount = state.materials.length;
              state.globalUpdated = direct.evt.time;
            } else {
              state.savedProjects = await window.BomStore.listSavedProjects();
            }
          } else if (!state.wb.projectMode && state.globalCount === 0 && state.materials.length > 0) {
            // Seed global storage automatically on first run if empty & not Project-Only
            const w = await window.BomStore.writeGlobalMaster(state.materials, null);
            state.globalCount = w.count;
            state.globalUpdated = w.updated;
          }
        } else {
          state.materials = [];
          state.categories = [];
        }
      }
      state.savedPool = await window.BomStore.getAvailableMaterialPool(state.materials);
      renderAll();
    } catch (err) {
      showGlobalMsg(err?.message || String(err), "err");
    }
  }

  /* ---------------- Renderers ---------------- */
  function renderAll() {
    renderHome();
    renderAddSheet();
    renderMaterials();
    renderSavedProjects();
    renderConvert();
    updateAllGuards();
  }

  function renderHome() {
    const visibleBoms = state.wb.bomSheets.filter((s) => !s.hidden);
    const convs = state.wb.convertedSheets;
    const isProj = Boolean(state.wb.projectMode);

    $("homeModeBadge").classList.toggle("hidden", !isProj);
    $("btnQuickImport").classList.toggle("hidden", !state.wb.needsTemplate);

    const st = $("homeWbStatus");
    if (visibleBoms.length) {
      st.className = "status-line";
      const preview = visibleBoms.slice(0, 3).map((s) => s.name).join(", ") + (visibleBoms.length > 3 ? "…" : "");
      st.textContent = `Automated BOM ready · ${visibleBoms.length} sheet(s) (${preview})`;
    } else if (state.wb.hasMaster) {
      st.className = "status-line";
      st.textContent = "Template ready · add your first BOM sheet.";
    } else if (convs.length) {
      st.className = "status-line warn";
      st.textContent = `${convs.length} converted plain BOM sheet(s) detected.`;
    } else {
      st.className = "status-line muted";
      st.textContent = "Blank workbook · import the lightweight template to begin.";
    }

    $("homeStoreStatus").textContent = isProj
      ? `Active Project-Only list: "${state.wb.projectMode.projectName}" (${state.materials.length} items). Global Master has ${state.globalCount} saved.`
      : `Global Master auto-saves on this computer (${state.globalCount} items).`;

    const names = visibleBoms.map((s) => s.name);
    $("homeRepairCount").textContent = `BOM sheets (${names.length})`;
    renderChecklist("homeRepairList", names, true);
  }

  function renderAddSheet() {
    const visibleBoms = state.wb.bomSheets.filter((s) => !s.hidden);
    const st = $("sheetWbStatus");
    if (state.wb.needsTemplate) {
      st.className = "status-line muted";
      st.textContent = "Adding your first sheet will import the lightweight template automatically.";
    } else {
      st.className = "status-line";
      st.textContent = `Ready · ${visibleBoms.length} active BOM sheet(s)`;
    }

    if (state.wb.projectTitle && !$("inpProjectTitle").value) {
      $("inpProjectTitle").value = state.wb.projectTitle;
    }

    const sel = $("selCopySource");
    const prev = sel.value;
    const opts = [];
    if (state.wb.hasBlank) opts.push({ v: window.BomSheets.BLANK, t: "BOM Blank (recommended)" });
    visibleBoms.forEach((s) => opts.push({ v: s.name, t: s.name }));
    if (!opts.length) opts.push({ v: window.BomSheets.BLANK, t: "BOM Blank (auto-import)" });
    sel.innerHTML = opts.map((o) => `<option value="${escHtml(o.v)}">${escHtml(o.t)}</option>`).join("");
    if (prev && opts.some((o) => o.v === prev)) sel.value = prev;

    const names = visibleBoms.map((s) => s.name);
    $("sheetRepairCount").textContent = `BOM sheets (${names.length})`;
    renderChecklist("sheetRepairList", names, true);
  }

  function renderMaterials() {
    const isProj = Boolean(state.wb.projectMode);
    $("projectBanner").classList.toggle("hidden", !isProj);
    $("projActiveBox").classList.toggle("hidden", !isProj);
    $("projCreateBox").classList.toggle("hidden", isProj);
    if (isProj) {
      $("projectBannerTitle").textContent = `Project-Only: ${state.wb.projectMode.projectName}`;
      $("projActiveDesc").textContent = `Active: "${state.wb.projectMode.projectName}" (${state.materials.length} items). Edits save to this project's slot only — Global Master is untouched.`;
    }
    $("matListHeading").textContent = isProj ? `Project Materials (${state.wb.projectMode.projectName})` : "Master Materials";
    $("backupMeta").textContent = `${state.globalCount} saved ▸`;
    $("backupProjectGuard").classList.toggle("hidden", !isProj);
    const stampTxt = state.globalUpdated ? ` · last saved ${window.BomSheets.formatStamp(state.globalUpdated)}` : "";
    $("backupStatusLine").textContent = `${state.globalCount} Global Master material(s) saved on this computer${stampTxt}.`;
    $("btnPullSaved").disabled = isProj || state.busy || !state.wb.hasMaster;
    $("btnPushSaved").disabled = isProj || state.busy || !state.wb.hasMaster;
    $("btnImportGlobal").disabled = isProj || state.busy;

    $("ghProjectGuard").classList.toggle("hidden", !isProj);
    $("btnGhPull").disabled = isProj || state.busy;
    $("btnGhPush").disabled = isProj || state.busy;

    const q = lc($("inpMatSearch").value);
    const validCount = state.materials.filter((m) => window.BomSheets.materialStatus(m).valid).length;
    const badCount = state.materials.length - validCount;
    $("matListMeta").textContent = badCount
      ? `${state.materials.length} items · ${badCount} need info`
      : `${state.materials.length} items · all ready`;

    // Category suggestions
    $("catSuggestions").innerHTML = state.categories.map((c) => `<option value="${escHtml(c)}"></option>`).join("");

    // Bubble recently changed materials to top
    const filtered = state.materials
      .filter((m) => !q || lc(m.material).includes(q) || lc(m.category).includes(q) || lc(m.unit).includes(q))
      .sort((a, b) => {
        const ra = state.recentNames.has(lc(a.material)) ? 1 : 0;
        const rb = state.recentNames.has(lc(b.material)) ? 1 : 0;
        if (ra !== rb) return rb - ra;
        return a.material.localeCompare(b.material, undefined, { sensitivity: "base" });
      });

    const box = $("matList");
    box.innerHTML = "";
    if (!filtered.length) {
      box.innerHTML = `<div style="padding:10px;" class="sub">${state.wb.hasMaster ? "No matching materials." : "Import template first to manage materials."}</div>`;
    } else {
      for (const m of filtered) {
        const st = window.BomSheets.materialStatus(m);
        const rec = state.recentNames.get(lc(m.material));
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "mat-row" +
          (state.editingName && lc(state.editingName) === lc(m.material) ? " selected" : "") +
          (rec ? " recent" : "");
        const wtStr = m.unit === "kg/m" ? `${m.lin ?? "?"} kg/m`
          : m.unit === "kg/m²" ? `${m.area ?? "?"} kg/m²`
          : m.unit === "kg/each" ? `${m.piece ?? "?"} kg/ea` : (m.unit || "—");
        const recTag = rec ? ` <span class="tag-recent">${escHtml(window.BomSheets.actionLabel(rec.type))}</span>` : "";
        btn.innerHTML = `
          <span class="mat-left">
            <span class="mat-state ${st.valid ? "ok" : "bad"}" title="${st.valid ? "BOM ready" : "Missing: " + escHtml(st.missing.join(", "))}">${st.valid ? "✓" : "!"}</span>
            <span class="mat-name">${escHtml(m.material)}${recTag}</span>
          </span>
          <span class="mat-right">${escHtml(wtStr)}</span>`;
        btn.addEventListener("click", () => loadMaterialIntoForm(m));
        box.appendChild(btn);
      }
    }

    // Update Pick-specific-materials list AND active Project-Only Add-from-Saved list
    renderProjPickList();
    renderProjAddSavedList();
  }

  function formatPoolLabel(m) {
    const src = m.inGlobal ? "Global" : (m.sources?.[0] || "Saved");
    return `${m.material} (${m.unit || "—"} · ${src})`;
  }

  function renderTrackedChecklist(containerId, items, selectedSet) {
    const box = $(containerId);
    box.innerHTML = "";
    if (!items.length) {
      box.innerHTML = '<div class="sub">No matching saved materials.</div>';
      return;
    }
    for (const item of items) {
      const key = lc(item.value);
      const checked = selectedSet.has(key);
      const row = document.createElement("label");
      row.className = "check-row";
      row.innerHTML = `<input type="checkbox" value="${escHtml(item.value)}" ${checked ? "checked" : ""}/><span>${escHtml(item.label)}</span>`;
      box.appendChild(row);
    }
  }

  function renderProjPickList() {
    const q = lc($("inpProjPickSearch").value);
    const pool = state.savedPool.length ? state.savedPool : state.materials;
    const items = pool
      .filter((m) => !q || lc(m.material).includes(q) || lc(m.category).includes(q) || (m.sources || []).some((s) => lc(s).includes(q)))
      .map((m) => ({ value: m.material, label: formatPoolLabel(m) }));
    renderTrackedChecklist("projPickList", items, state.projPickSelected);
    updateProjPickCount();
  }

  function renderProjAddSavedList() {
    const existing = new Set(state.materials.map((m) => lc(m.material)));
    const available = state.savedPool.filter((m) => !existing.has(lc(m.material)));
    // Prune any previously checked items that are now already in this project
    for (const k of [...state.projAddSelected]) {
      if (existing.has(k)) state.projAddSelected.delete(k);
    }
    $("projAddSavedMeta").textContent = `${available.length} available ▸`;
    const q = lc($("inpProjAddSearch").value);
    const items = available
      .filter((m) => !q || lc(m.material).includes(q) || lc(m.category).includes(q) || (m.sources || []).some((s) => lc(s).includes(q)))
      .map((m) => ({ value: m.material, label: formatPoolLabel(m) }));
    renderTrackedChecklist("projAddSavedList", items, state.projAddSelected);
    updateProjPickCount();
  }

  function updateProjPickCount() {
    $("projPickCount").textContent = `Selected (${state.projPickSelected.size})`;
    $("projAddSavedCount").textContent = `Selected (${state.projAddSelected.size})`;
    $("btnAddSavedToProj").disabled = state.busy || !state.wb.projectMode || state.projAddSelected.size === 0;
  }

  function renderSavedProjects() {
    const isProj = Boolean(state.wb.projectMode);
    $("savedProjCount").textContent = String(state.savedProjects.length);
    $("projDrawerMeta").textContent = isProj
      ? `Active: ${state.wb.projectMode.projectName} ▸`
      : state.savedProjects.length ? `${state.savedProjects.length} saved ▸` : "Isolate workbook ▸";
    const list = $("savedProjList");
    list.innerHTML = "";
    if (!state.savedProjects.length) {
      list.innerHTML = '<div class="sub">No saved project lists yet.</div>';
    } else {
      for (const p of state.savedProjects) {
        const div = document.createElement("div");
        div.className = "proj-item";
        div.innerHTML = `
          <div class="proj-top">
            <span class="proj-name">${escHtml(p.projectName)}</span>
            <span class="proj-meta">${p.materials.length} items</span>
          </div>
          <div class="row-btns">
            <button type="button" class="btn btn-secondary btn-sm" data-pload="${escHtml(p.id)}">Load</button>
            <button type="button" class="btn btn-secondary btn-sm" data-pexp="${escHtml(p.id)}">Export JSON</button>
            <button type="button" class="btn btn-danger btn-sm" data-pdel="${escHtml(p.id)}">Delete</button>
          </div>`;
        list.appendChild(div);
      }
    }

    // Populate Restore dropdown & show clear status banner when workbook has a BomProjectList tag
    const sel = $("selRestProj");
    const restBanner = $("restProjBanner");
    if (!state.savedProjects.length) {
      sel.innerHTML = '<option value="">No saved project lists on this PC</option>';
    } else {
      sel.innerHTML = state.savedProjects
        .map((p) => `<option value="${escHtml(p.id)}">${escHtml(p.projectName)} (${p.materials.length} items)</option>`)
        .join("");
    }

    if (state.wb.projectMode) {
      const match = state.savedProjects.find(
        (p) => p.id === state.wb.projectMode.id ||
               lc(p.projectName) === lc(state.wb.projectMode.projectName) ||
               window.BomStore.slugify(p.projectName) === window.BomStore.slugify(state.wb.projectMode.projectName)
      );
      if (match) {
        sel.value = match.id;
        restBanner.className = "hint ok";
        restBanner.textContent = `Project list "${match.projectName}" detected (${match.materials.length} materials saved on this computer) — pre-selected below.`;
        const rad = document.querySelector('input[name="restMatSrc"][value="project-saved"]');
        if (rad && !state.pickedRestProj) {
          rad.checked = true;
          document.querySelectorAll('[data-rsrc]').forEach((c) => c.classList.toggle("selected", c.dataset.rsrc === "project-saved"));
        }
      } else {
        restBanner.className = "hint warn";
        restBanner.textContent = `This workbook used Project-Only list "${state.wb.projectMode.projectName}", which has no saved copy on this computer. Load its .json file below before restoring.`;
      }
    } else {
      restBanner.className = "hint hidden";
    }
  }

  function renderConvert() {
    const visibleBoms = state.wb.bomSheets.filter((s) => !s.hidden).map((s) => s.name);
    const convs = state.wb.convertedSheets;

    // Flatten section
    const fst = $("flatWbStatus");
    if (visibleBoms.length) {
      fst.className = "status-line";
      fst.textContent = `Ready · ${visibleBoms.length} automated BOM sheet(s)`;
    } else {
      fst.className = "status-line muted";
      fst.textContent = "No automated BOM sheets found to flatten.";
    }
    $("flatSheetCount").textContent = `Sheets to flatten (${visibleBoms.length})`;
    renderChecklist("flatSheetList", visibleBoms, true);
    $("wrapAutoSaveProj").classList.toggle("hidden", !state.wb.projectMode);

    // Restore section
    const rst = $("restWbStatus");
    if (convs.length) {
      rst.className = "status-line";
      rst.textContent = `Ready · ${convs.length} converted plain sheet(s) detected`;
    } else {
      rst.className = "status-line muted";
      rst.textContent = "No converted plain BOM sheets detected in this workbook.";
    }
    $("restSheetCount").textContent = `Converted sheets (${convs.length})`;
    renderChecklist("restSheetList", convs, true);
    $("restGlobalDesc").textContent = `Use auto-saved Global Master (${state.globalCount} items).`;
  }

  /* ---------------- Material Form Helpers ---------------- */
  function updateUnitFields() {
    const u = $("selMatUnit").value;
    $("grpLin").classList.toggle("hidden", u !== "kg/m");
    $("grpArea").classList.toggle("hidden", u !== "kg/m²");
    $("grpPiece").classList.toggle("hidden", u !== "kg/each");
  }

  function readMaterialForm() {
    const unit = $("selMatUnit").value;
    return {
      material: $("inpMatName").value.trim(),
      unit,
      lin: unit === "kg/m" ? $("inpMatLin").value : null,
      area: unit === "kg/m²" ? $("inpMatArea").value : null,
      sl: unit === "kg/m²" ? $("inpMatSl").value : null,
      sw: unit === "kg/m²" ? $("inpMatSw").value : null,
      piece: unit === "kg/each" ? $("inpMatPiece").value : null,
      category: $("inpMatCat").value.trim(),
      remarks: $("inpMatRem").value.trim()
    };
  }

  function clearMaterialForm() {
    state.editingName = null;
    state.deleteArm = false;
    $("matFormTitle").textContent = "Add Material";
    $("btnMatNew").classList.add("hidden");
    $("btnMatSave").textContent = "Add Material";
    $("btnMatDelete").classList.add("hidden");
    $("btnMatDelete").textContent = "Remove";
    ["inpMatName", "inpMatLin", "inpMatArea", "inpMatSl", "inpMatSw", "inpMatPiece", "inpMatCat", "inpMatRem"]
      .forEach((id) => { $(id).value = ""; });
    $("selMatUnit").value = "";
    updateUnitFields();
    renderMaterials();
    updateMaterialGuard();
  }

  function loadMaterialIntoForm(m) {
    state.editingName = m.material;
    state.deleteArm = false;
    $("matFormTitle").textContent = `Edit Material`;
    $("btnMatNew").classList.remove("hidden");
    $("btnMatSave").textContent = "Save Changes";
    $("btnMatDelete").classList.remove("hidden");
    $("btnMatDelete").textContent = "Remove";
    $("inpMatName").value = m.material || "";
    $("selMatUnit").value = m.unit || "";
    $("inpMatLin").value = m.lin ?? "";
    $("inpMatArea").value = m.area ?? "";
    $("inpMatSl").value = m.sl ?? "";
    $("inpMatSw").value = m.sw ?? "";
    $("inpMatPiece").value = m.piece ?? "";
    $("inpMatCat").value = m.category || "";
    $("inpMatRem").value = m.remarks || "";
    updateUnitFields();
    renderMaterials();
    updateMaterialGuard();
  }

  /* ---------------- Calm Live Primary-Requirement Guards ---------------- */
  function updateAllGuards() {
    updateSheetGuard();
    updateMaterialGuard();
    updateProjectGuard();
    updateConvertGuards();
  }

  function updateProjectGuard() {
    const isProj = Boolean(state.wb.projectMode);
    const name = $("inpNewProjName").value.trim();
    const mode = $("selProjStartMode").value;
    const picked = state.projPickSelected.size;
    const hint = $("projHint");

    let msg = "";
    let warn = false;
    if (!isProj) {
      if (!state.wb.hasMaster) {
        msg = "Open an automated BOM workbook (with Master Unique Materials) to start a new project list.";
      } else if (!name) {
        msg = "Enter a name for this project list.";
      } else if (mode === "pick" && picked === 0) {
        msg = "Select at least one material to include.";
        warn = true;
      }
    }
    if (msg) {
      hint.textContent = msg;
      hint.className = warn ? "hint warn" : "hint";
    } else {
      hint.className = "hint hidden";
    }
    $("btnCreateProj").disabled = state.busy || isProj || Boolean(msg);
    $("btnBannerAddSaved").disabled = state.busy || !isProj;
    $("btnBannerExportProj").disabled = state.busy || !isProj;
    $("btnBannerSwitchGlobal").disabled = state.busy || !isProj;
    $("btnDrawerExportProj").disabled = state.busy || !isProj;
    $("btnDrawerSwitchGlobal").disabled = state.busy || !isProj;
    $("btnAddSavedToProj").disabled = state.busy || !isProj || state.projAddSelected.size === 0;
    $("btnPullSaved").disabled = state.busy || isProj || !state.wb.hasMaster;
    $("btnPushSaved").disabled = state.busy || isProj || !state.wb.hasMaster;
    $("btnImportGlobal").disabled = state.busy || isProj;
    $("btnGhPull").disabled = state.busy || isProj;
    $("btnGhPush").disabled = state.busy || isProj;
  }

  function updateSheetGuard() {
    const val = $("inpSheetName").value;
    const err = window.BomSheets.checkSheetName(val, state.wb.allSheets);
    const hint = $("sheetHint");
    const isEmpty = !val.trim();
    if (err) {
      hint.textContent = err;
      hint.className = isEmpty ? "hint" : "hint warn";
    } else {
      hint.className = "hint hidden";
    }
    $("btnAddSheet").disabled = state.busy || Boolean(err);
    $("btnHomeRepair").disabled = state.busy || getChecked("homeRepairList").length === 0;
    $("btnSheetRepair").disabled = state.busy || getChecked("sheetRepairList").length === 0;
  }

  function updateMaterialGuard() {
    const raw = readMaterialForm();
    const names = state.materials.map((m) => m.material);
    const mode = state.editingName ? "edit" : "add";
    const issues = state.wb.hasMaster
      ? window.BomSheets.validateMaterial(raw, names, mode)
      : ["Import the BOM template first to manage materials."];

    const hint = $("matHint");
    if (issues.length) {
      const neutral = !raw.material && !raw.unit;
      hint.textContent = neutral ? "Required: Material name, Primary unit, and Weight / Sheet Size." : issues[0];
      hint.className = neutral ? "hint" : "hint warn";
    } else {
      hint.className = "hint hidden";
    }
    $("btnMatSave").disabled = state.busy || issues.length > 0;
    $("btnMatDelete").disabled = state.busy || !state.editingName;
  }

  function updateConvertGuards() {
    // Flatten guard
    const totalBom = state.wb.bomSheets.filter((s) => !s.hidden).length;
    const flatSel = getChecked("flatSheetList");
    const confirmed = $("chkCopyConfirmed").checked;
    const partial = flatSel.length > 0 && flatSel.length < totalBom;
    $("partialFlattenNote").classList.toggle("hidden", !partial);

    let flatMsg = "";
    if (!totalBom) flatMsg = "No automated BOM sheets to flatten.";
    else if (!flatSel.length) flatMsg = "Select at least one BOM sheet to convert.";
    else if (!confirmed) flatMsg = "Confirm you are working on a copy (or have a backup) to enable conversion.";
    const fh = $("flatHint");
    if (flatMsg) { fh.textContent = flatMsg; fh.className = "hint"; } else { fh.className = "hint hidden"; }
    $("btnFlatten").disabled = state.busy || Boolean(flatMsg);

    // Restore guard
    const restSel = getChecked("restSheetList");
    const rSrc = document.querySelector('input[name="restMatSrc"]:checked')?.value || "global";
    let restMsg = "";
    if (!state.wb.convertedSheets.length) restMsg = "No converted plain BOM sheets detected.";
    else if (!restSel.length) restMsg = "Select at least one converted sheet to restore.";
    else if (rSrc === "project-saved" && !$("selRestProj").value) restMsg = "Choose a saved project list on this PC.";
    else if (rSrc === "project-file" && !state.pickedRestProj) restMsg = "Choose a project .json file first.";
    const rh = $("restHint");
    if (restMsg) { rh.textContent = restMsg; rh.className = "hint"; } else { rh.className = "hint hidden"; }
    $("btnRestore").disabled = state.busy || Boolean(restMsg);
    $("btnRepairPlain").disabled = state.busy || !restSel.length;
  }

  /* ---------------- Persist Current Workbook Materials ---------------- */
  async function persistAfterMaterialChange(items) {
    const nowIso = new Date().toISOString();
    const lastChange = window.BomStore.buildLastChange("panel", items, nowIso);
    const fresh = await window.BomSheets.listMaterials();
    state.materials = fresh.rows;
    state.categories = fresh.categories;
    window.BomSheets.setMasterSnap(fresh.rows);

    if (fresh.projectMode) {
      const saved = await window.BomStore.saveProjectList(fresh.projectMode, fresh.rows, lastChange);
      state.wb.projectMode = { id: saved.id, projectName: saved.projectName, updated: saved.updated };
      state.savedProjects = await window.BomStore.listSavedProjects();
    } else {
      const w = await window.BomStore.writeGlobalMaster(fresh.rows, lastChange);
      state.globalCount = w.count;
      state.globalUpdated = w.updated;
      // Optional GitHub auto-publish when a publisher token is saved (never in Project-Only mode!)
      if (window.BomStore.hasToken()) {
        window.BomStore.pushGitHub(fresh.rows, `Update materials.json (${fresh.rows.length} items)`).catch(() => {});
      }
    }
    state.savedPool = await window.BomStore.getAvailableMaterialPool(fresh.rows);
    recordChangeEvent({ source: "panel", time: nowIso, items });
  }

  /* ---------------- Open Copy Helper ---------------- */
  async function openWorkbookCopy() {
    try {
      if (!Office?.context?.document?.getFileAsync) {
        showGlobalMsg("Use File → Save a Copy in Excel, then open that copy before converting.", "warn");
        return;
      }
      setBusy(true);
      await new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, async (res) => {
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            return reject(new Error(res.error?.message || "Could not read workbook file."));
          }
          const file = res.value;
          try {
            const chunks = [];
            for (let i = 0; i < file.sliceCount; i++) {
              const slice = await new Promise((res2, rej2) => {
                file.getSliceAsync(i, (sr) => sr.status === Office.AsyncResultStatus.Succeeded ? res2(sr.value.data) : rej2(sr.error));
              });
              chunks.push(new Uint8Array(slice));
            }
            const total = chunks.reduce((a, b) => a + b.length, 0);
            const merged = new Uint8Array(total);
            let off = 0;
            for (const c of chunks) { merged.set(c, off); off += c.length; }
            let bin = "";
            for (let i = 0; i < merged.length; i += 0x8000) bin += String.fromCharCode(...merged.subarray(i, i + 0x8000));
            await Excel.createWorkbook(btoa(bin));
            resolve();
          } catch (e) {
            reject(e);
          } finally {
            try { file.closeAsync(); } catch { /* ignore */ }
          }
        });
      });
      showGlobalMsg("Opened a fresh copy in a new Excel window.", "ok");
    } catch (err) {
      showGlobalMsg(`Could not open copy automatically (${err?.message || err}). Use File → Save a Copy in Excel.`, "warn");
    } finally {
      setBusy(false);
    }
  }

  /* ---------------- Event Wiring ---------------- */
  function wireEvents() {
    // Tabs & jumps
    document.querySelectorAll(".tab").forEach((b) => b.addEventListener("click", () => switchTab(b.dataset.tab)));
    document.querySelectorAll("[data-jump]").forEach((b) => {
      b.addEventListener("click", () => {
        const [tab, sub] = b.dataset.jump.split(":");
        switchTab(tab, sub);
      });
    });
    $("btnRefresh").addEventListener("click", () => refreshAll());

    // All / None checklist buttons (updates tracked sets for searchable saved-material pickers!)
    document.querySelectorAll("[data-checkall]").forEach((b) => {
      b.addEventListener("click", () => {
        const val = b.dataset.val === "1";
        const targetId = b.dataset.checkall;
        $(targetId).querySelectorAll('input[type="checkbox"]').forEach((c) => {
          c.checked = val;
          if (targetId === "projPickList") {
            if (val) state.projPickSelected.add(lc(c.value));
            else state.projPickSelected.delete(lc(c.value));
          } else if (targetId === "projAddSavedList") {
            if (val) state.projAddSelected.add(lc(c.value));
            else state.projAddSelected.delete(lc(c.value));
          }
        });
        updateProjPickCount();
        updateAllGuards();
      });
    });

    // Dynamic checklist changes
    ["homeRepairList", "sheetRepairList", "flatSheetList", "restSheetList"].forEach((id) => {
      $(id).addEventListener("change", () => {
        updateAllGuards();
      });
    });
    $("projPickList").addEventListener("change", (e) => {
      const cb = e.target;
      if (cb && cb.type === "checkbox") {
        if (cb.checked) state.projPickSelected.add(lc(cb.value));
        else state.projPickSelected.delete(lc(cb.value));
      }
      updateProjPickCount();
      updateAllGuards();
    });
    $("projAddSavedList").addEventListener("change", (e) => {
      const cb = e.target;
      if (cb && cb.type === "checkbox") {
        if (cb.checked) state.projAddSelected.add(lc(cb.value));
        else state.projAddSelected.delete(lc(cb.value));
      }
      updateProjPickCount();
      updateAllGuards();
    });

    // Quick Template Import
    $("btnQuickImport").addEventListener("click", async () => {
      setBusy(true);
      try {
        await window.BomSheets.importTemplate();
        const global = await window.BomStore.readGlobalMaster();
        if (global.materials.length) {
          await window.BomSheets.applyMaterialList(global.materials, {
            syncRemovals: true,
            projectMeta: null,
            lastChange: global.lastChange
          });
        }
        await refreshAll();
        showGlobalMsg("Lightweight Excel 365 template imported (Master Unique Materials + hidden BOM Blank).", "ok");
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });

    // Add BOM Sheet (zero trailing Excel.run calls after activation so typing in B6 is 100% lag-free!)
    $("inpSheetName").addEventListener("input", updateSheetGuard);
    $("btnAddSheet").addEventListener("click", async () => {
      const name = $("inpSheetName").value.trim();
      const wasTemplateNeeded = state.wb.needsTemplate;
      setBusy(true);
      try {
        if (wasTemplateNeeded) {
          await window.BomSheets.importTemplate();
          const global = await window.BomStore.readGlobalMaster();
          if (global.materials.length) {
            await window.BomSheets.applyMaterialList(global.materials, {
              syncRemovals: true,
              projectMeta: null,
              lastChange: global.lastChange
            });
          }
        } else if ($("chkSyncBeforeAdd").checked && !state.wb.projectMode) {
          const global = await window.BomStore.readGlobalMaster();
          if (global.materials.length) {
            // Only touch Master sheet if Global Master actually differs from current workbook Master
            const diff = window.BomStore.diffLists(global.materials, state.materials);
            if (diff.added.length || diff.modified.length || diff.removed.length) {
              await window.BomSheets.applyMaterialList(global.materials, {
                syncRemovals: true,
                lastChange: global.lastChange
              });
              state.materials = global.materials;
              window.BomSheets.setMasterSnap(global.materials);
            }
          }
        }
        const source = $("selCopySource").value || window.BomSheets.BLANK;
        const res = await window.BomSheets.addBomSheet({
          name,
          source,
          project: $("inpProjectTitle").value.trim(),
          clear: $("chkClearInputs").checked
        });
        $("inpSheetName").value = "";
        if (wasTemplateNeeded) {
          await refreshAll();
        } else {
          // Update task pane sheet state in memory — ZERO trailing Excel.run COM bursts while user types in B6!
          if (!state.wb.allSheets.includes(res.name)) state.wb.allSheets.push(res.name);
          if (!state.wb.bomSheets.some((s) => lc(s.name) === lc(res.name))) {
            state.wb.bomSheets.push({ name: res.name, hidden: false });
          }
          renderHome();
          renderAddSheet();
          renderConvert();
          updateAllGuards();
        }
        showGlobalMsg(res.warnings.length ? `Added "${res.name}". ${res.warnings[0]}` : `Added lightweight BOM sheet "${res.name}".`, res.warnings.length ? "warn" : "ok");
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });

    // Repair buttons (Home & Add Sheet)
    async function runRepairFrom(listId) {
      const sheets = getChecked(listId);
      if (!sheets.length) return;
      setBusy(true);
      try {
        const res = await window.BomSheets.repairWorkbook({ sheets });
        await refreshAll();
        showGlobalMsg(`Repaired & optimized ${res.sheetsRepaired} BOM sheet(s).`, "ok");
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    }
    $("btnHomeRepair").addEventListener("click", () => runRepairFrom("homeRepairList"));
    $("btnSheetRepair").addEventListener("click", () => runRepairFrom("sheetRepairList"));

    // Toast hover pause & session toggle
    $("changeToast").addEventListener("mouseenter", () => { state.toastPaused = true; clearTimeout(state.toastTimer); });
    $("changeToast").addEventListener("mouseleave", () => { state.toastPaused = false; startToastTimer(); });
    $("btnCloseToast").addEventListener("click", collapseToast);
    $("btnSessionToggle").addEventListener("click", () => {
      const hidden = $("sessionLog").classList.toggle("hidden");
      $("sessionCaret").textContent = hidden ? "▸" : "▾";
    });

    // Materials search & form
    $("inpMatSearch").addEventListener("input", renderMaterials);
    $("selMatUnit").addEventListener("change", () => { updateUnitFields(); updateMaterialGuard(); });
    ["inpMatName", "inpMatLin", "inpMatArea", "inpMatSl", "inpMatSw", "inpMatPiece", "inpMatCat", "inpMatRem"]
      .forEach((id) => $(id).addEventListener("input", updateMaterialGuard));
    $("btnMatNew").addEventListener("click", clearMaterialForm);

    $("btnMatSave").addEventListener("click", async () => {
      const item = readMaterialForm();
      setBusy(true);
      try {
        if (state.editingName) {
          const prev = state.materials.find((m) => lc(m.material) === lc(state.editingName));
          const updated = await window.BomSheets.updateMaterial(state.editingName, item);
          const fields = prev ? window.BomStore.changedFields(prev, updated) : [];
          await persistAfterMaterialChange([{ type: "modified", material: updated, fields }]);
          clearMaterialForm();
          showGlobalMsg(`Updated "${updated.material}".`, "ok");
        } else {
          const added = await window.BomSheets.addMaterial(item);
          await persistAfterMaterialChange([{ type: "added", material: added, fields: [] }]);
          clearMaterialForm();
          showGlobalMsg(`Added "${added.material}".`, "ok");
        }
        renderAll();
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });

    $("btnMatDelete").addEventListener("click", async () => {
      if (!state.editingName) return;
      if (!state.deleteArm) {
        state.deleteArm = true;
        $("btnMatDelete").textContent = "Confirm Remove";
        return;
      }
      const target = state.editingName;
      setBusy(true);
      try {
        await window.BomSheets.deleteMaterial(target);
        await persistAfterMaterialChange([{ type: "removed", material: { material: target }, fields: [] }]);
        clearMaterialForm();
        renderAll();
        showGlobalMsg(`Removed "${target}".`, "ok");
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });

    // Project-Only Material Lists
    $("inpNewProjName").addEventListener("input", updateProjectGuard);
    $("selProjStartMode").addEventListener("change", () => {
      $("projPickBox").classList.toggle("hidden", $("selProjStartMode").value !== "pick");
      if ($("selProjStartMode").value === "pick") renderProjPickList();
      updateProjectGuard();
    });
    $("inpProjPickSearch").addEventListener("input", renderProjPickList);
    $("inpProjAddSearch").addEventListener("input", renderProjAddSavedList);

    $("btnBannerAddSaved").addEventListener("click", () => {
      $("projectDrawer").open = true;
      $("projAddSavedDrawer").open = true;
      renderProjAddSavedList();
      $("inpProjAddSearch").focus();
    });

    $("btnAddSavedToProj").addEventListener("click", async () => {
      if (!state.wb.projectMode || !state.projAddSelected.size) return;
      setBusy(true);
      try {
        await checkDirectMasterEdits();
        const chosen = state.savedPool.filter((m) => state.projAddSelected.has(lc(m.material)));
        if (!chosen.length) throw new Error("Select at least one saved material to add.");
        const res = await window.BomSheets.addMaterialsBulk(chosen);
        state.projAddSelected.clear();
        if (res.added.length) {
          await persistAfterMaterialChange(
            res.added.map((m) => ({ type: "added", material: m, fields: [] }))
          );
          renderAll();
          showGlobalMsg(
            `Added ${res.added.length} saved material(s) to Project-Only list "${state.wb.projectMode.projectName}". Global Master is untouched.`,
            "ok"
          );
        } else {
          showGlobalMsg("All selected materials are already in this project list.", "warn");
        }
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });

    $("btnCreateProj").addEventListener("click", async () => {
      const projectName = $("inpNewProjName").value.trim();
      if (!projectName) return;
      const startMode = $("selProjStartMode").value;
      setBusy(true);
      try {
        // Save any unsaved direct edits on Master sheet first before isolating
        await checkDirectMasterEdits();
        let baseList = [];
        if (startMode === "copy") {
          const g = await window.BomStore.readGlobalMaster();
          baseList = g.materials.length ? g.materials : state.materials;
        } else if (startMode === "pick") {
          const pool = state.savedPool.length ? state.savedPool : state.materials;
          baseList = pool.filter((m) => state.projPickSelected.has(lc(m.material)));
          if (!baseList.length) throw new Error("Select at least one material to include.");
        } else if (startMode === "used") {
          const used = await window.BomSheets.getUsedMaterialsInWorkbook();
          const pool = state.savedPool.length ? state.savedPool : state.materials;
          baseList = pool.filter((m) => used.has(lc(m.material)));
        } else {
          baseList = [];
        }
        const saved = await window.BomStore.saveProjectList({ projectName }, baseList, null);
        const meta = { id: saved.id, projectName: saved.projectName, updated: saved.updated };
        await window.BomSheets.applyMaterialList(baseList, { replaceAll: true, projectMeta: meta });
        $("inpNewProjName").value = "";
        state.projPickSelected.clear();
        state.projAddSelected.clear();
        await refreshAll();
        showGlobalMsg(`Started Project-Only list "${saved.projectName}" (${baseList.length} items). Global Master is untouched.`, "ok");
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });

    async function exportActiveProjectJson() {
      if (!state.wb.projectMode) return;
      setBusy(true);
      try {
        await checkDirectMasterEdits();
        const json = window.BomStore.serializeProject(state.wb.projectMode.projectName, state.materials, null, state.wb.projectMode.id);
        const res = await window.BomStore.saveJsonFile(`project-materials-${window.BomStore.slugify(state.wb.projectMode.projectName)}.json`, json);
        if (!res.cancelled) showGlobalMsg(`Saved project file ${res.fileName} (${state.materials.length} items).`, "ok");
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    }
    $("btnBannerExportProj").addEventListener("click", exportActiveProjectJson);
    $("btnDrawerExportProj").addEventListener("click", exportActiveProjectJson);

    function resetSwitchGlobalArm() {
      state.switchGlobalArm = false;
      clearTimeout(state.switchGlobalTimer);
      $("btnBannerSwitchGlobal").textContent = "Switch to Global";
      $("btnDrawerSwitchGlobal").textContent = "Switch to Global Master";
    }

    async function handleSwitchToGlobal() {
      if (!state.wb.projectMode) return;
      if (!state.switchGlobalArm) {
        state.switchGlobalArm = true;
        $("btnBannerSwitchGlobal").textContent = "Confirm: Switch to Global";
        $("btnDrawerSwitchGlobal").textContent = "Confirm: Switch to Global";
        showGlobalMsg("Click again to confirm switching back to Global Master (your saved project copy remains safe on this PC).", "warn");
        clearTimeout(state.switchGlobalTimer);
        state.switchGlobalTimer = setTimeout(resetSwitchGlobalArm, 8000);
        return;
      }
      resetSwitchGlobalArm();
      setBusy(true);
      try {
        // Save latest project edits to the project's PC slot before leaving Project-Only mode
        await checkDirectMasterEdits();
        if (state.wb.projectMode) {
          await window.BomStore.saveProjectList(state.wb.projectMode, state.materials, null);
        }
        const g = await window.BomStore.readGlobalMaster();
        if (!g.materials.length) throw new Error("Saved Global Master is empty.");
        await window.BomSheets.applyMaterialList(g.materials, {
          replaceAll: true,
          projectMeta: null,
          lastChange: g.lastChange
        });
        await refreshAll();
        showGlobalMsg(`Switched workbook back to Global Master (${g.materials.length} items). Project list remains saved on this PC.`, "ok");
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    }
    $("btnBannerSwitchGlobal").addEventListener("click", handleSwitchToGlobal);
    $("btnDrawerSwitchGlobal").addEventListener("click", handleSwitchToGlobal);

    $("btnLoadProjFile").addEventListener("click", () => $("fileLoadProj").click());
    $("fileLoadProj").addEventListener("change", async (e) => {
      const f = e.target.files?.[0];
      e.target.value = "";
      if (!f) return;
      setBusy(true);
      try {
        const parsed = window.BomStore.parseJsonText(await f.text());
        if (!parsed.materials.length) throw new Error("No materials found in that JSON file.");
        const projectName = parsed.projectName || f.name.replace(/\.json$/i, "").replace(/^project-materials-?/i, "").replace(/[-_]+/g, " ").trim() || "Project";
        const saved = await window.BomStore.saveProjectList({ id: parsed.id, projectName }, parsed.materials, parsed.lastChange);
        const meta = { id: saved.id, projectName: saved.projectName, updated: saved.updated };
        if (state.wb.hasMaster) {
          await window.BomSheets.applyMaterialList(parsed.materials, {
            replaceAll: true,
            projectMeta: meta,
            lastChange: parsed.lastChange
          });
          await refreshAll();
          showGlobalMsg(`Loaded Project-Only list "${saved.projectName}" (${parsed.materials.length} items). Global Master is untouched.`, "ok");
        } else {
          // Converted/plain workbook (no Master sheet yet): tag workbook & save on PC so Restore uses it automatically
          await window.BomSheets.setProjectMeta(meta);
          await refreshAll();
          showGlobalMsg(`Saved Project-Only list "${saved.projectName}" (${parsed.materials.length} items). It is now pre-selected for Convert → Restore Automated.`, "ok");
        }
      } catch (err) {
        showGlobalMsg(`Could not load project file: ${err?.message || err}`, "err");
      } finally {
        setBusy(false);
      }
    });

    // Saved projects delegation (Load / Export / Delete)
    $("savedProjList").addEventListener("click", async (e) => {
      const btn = e.target.closest("button");
      if (!btn) return;
      const loadId = btn.dataset.pload;
      const expId = btn.dataset.pexp;
      const delId = btn.dataset.pdel;
      if (loadId) {
        setBusy(true);
        try {
          const p = await window.BomStore.getSavedProject(loadId);
          if (!p) throw new Error("Saved project list not found.");
          const meta = { id: p.id, projectName: p.projectName, updated: p.updated };
          if (state.wb.hasMaster) {
            await window.BomSheets.applyMaterialList(p.materials, {
              replaceAll: true,
              projectMeta: meta,
              lastChange: p.lastChange
            });
            await refreshAll();
            showGlobalMsg(`Loaded Project-Only list "${p.projectName}" (${p.materials.length} items).`, "ok");
          } else {
            await window.BomSheets.setProjectMeta(meta);
            await refreshAll();
            showGlobalMsg(`Linked Project-Only list "${p.projectName}" (${p.materials.length} items) for Convert → Restore Automated.`, "ok");
          }
        } catch (err) { showGlobalMsg(err?.message || String(err), "err"); }
        finally { setBusy(false); }
      } else if (expId) {
        const p = await window.BomStore.getSavedProject(expId);
        if (!p) return;
        const json = window.BomStore.serializeProject(p.projectName, p.materials, p.lastChange, p.id);
        const res = await window.BomStore.saveJsonFile(`project-materials-${window.BomStore.slugify(p.projectName)}.json`, json);
        if (!res.cancelled) showGlobalMsg(`Exported ${res.fileName}.`, "ok");
      } else if (delId) {
        await window.BomStore.deleteSavedProject(delId);
        state.savedProjects = await window.BomStore.listSavedProjects();
        state.savedPool = await window.BomStore.getAvailableMaterialPool(state.materials);
        renderSavedProjects();
        renderProjPickList();
        renderProjAddSavedList();
        updateAllGuards();
        showGlobalMsg("Deleted saved project list from this PC.", "ok");
      }
    });

    // Backup & Transfer (Global Master JSON — strictly rejects Project-Only files!)
    $("btnExportGlobal").addEventListener("click", async () => {
      if (!state.wb.projectMode) await checkDirectMasterEdits();
      const g = await window.BomStore.readGlobalMaster();
      const list = g.materials.length ? g.materials : (!state.wb.projectMode ? state.materials : []);
      if (!list.length) {
        showGlobalMsg("No Global Master materials saved yet.", "warn");
        return;
      }
      const res = await window.BomStore.saveJsonFile("master-materials.json", window.BomStore.serializeMaster(list, g.lastChange));
      if (!res.cancelled) showGlobalMsg(`Exported ${res.fileName} (${list.length} materials).`, "ok");
    });

    $("btnImportGlobal").addEventListener("click", () => $("fileImportGlobal").click());
    $("fileImportGlobal").addEventListener("change", async (e) => {
      const f = e.target.files?.[0];
      e.target.value = "";
      if (!f) return;
      if (state.wb.projectMode) {
        showGlobalMsg("Switch back to Global Master first before importing a Global Master JSON.", "warn");
        return;
      }
      setBusy(true);
      try {
        const parsed = window.BomStore.parseJsonText(await f.text());
        if (parsed.type === "project") {
          throw new Error("This is a Project-Only file. Use 'Project-Only Material List → Load Project .json File' instead so your Global Master stays safe.");
        }
        if (!parsed.materials.length) throw new Error("JSON file contained no materials.");
        await window.BomStore.writeGlobalMaster(parsed.materials, parsed.lastChange);
        if (state.wb.hasMaster && !state.wb.projectMode) {
          await window.BomSheets.applyMaterialList(parsed.materials, {
            replaceAll: true,
            projectMeta: null,
            lastChange: parsed.lastChange
          });
        }
        await refreshAll();
        showGlobalMsg(`Imported ${parsed.materials.length} Global Master materials.`, "ok");
      } catch (err) {
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });

    $("btnPullSaved").addEventListener("click", async () => {
      if (state.wb.projectMode) {
        showGlobalMsg("Project-Only list is active. Switch to Global Master first.", "warn");
        return;
      }
      setBusy(true);
      try {
        const g = await window.BomStore.readGlobalMaster();
        if (!g.materials.length) throw new Error("No Global Master saved on this computer yet.");
        await window.BomSheets.applyMaterialList(g.materials, {
          syncRemovals: true,
          projectMeta: null,
          lastChange: g.lastChange
        });
        await refreshAll();
        showGlobalMsg(`Loaded ${g.materials.length} Global Master materials into workbook.`, "ok");
      } catch (err) { showGlobalMsg(err?.message || String(err), "err"); }
      finally { setBusy(false); }
    });

    $("btnPushSaved").addEventListener("click", async () => {
      if (state.wb.projectMode) {
        showGlobalMsg("Workbook is in Project-Only mode. Use Save Project .json File instead.", "warn");
        return;
      }
      setBusy(true);
      try {
        const detected = await checkDirectMasterEdits();
        if (!detected) {
          const fresh = await window.BomSheets.listMaterials();
          await window.BomStore.writeGlobalMaster(fresh.rows, null);
          window.BomSheets.setMasterSnap(fresh.rows);
        }
        await refreshAll();
        showGlobalMsg(`Saved ${state.materials.length} materials to Global Master on this computer.`, "ok");
      } catch (err) { showGlobalMsg(err?.message || String(err), "err"); }
      finally { setBusy(false); }
    });

    // Optional GitHub Sync
    $("inpGhToken").value = window.BomStore.getToken();
    $("btnGhSaveTok").addEventListener("click", () => {
      window.BomStore.setToken($("inpGhToken").value);
      showGlobalMsg("GitHub token saved on this computer.", "ok");
    });
    $("btnGhClearTok").addEventListener("click", () => {
      $("inpGhToken").value = "";
      window.BomStore.setToken("");
      showGlobalMsg("GitHub token cleared.", "ok");
    });
    $("btnGhPull").addEventListener("click", async () => {
      if (state.wb.projectMode) return;
      setBusy(true);
      try {
        const r = await window.BomStore.fetchGitHubPages();
        if (r.missing || !r.materials.length) throw new Error("No materials found in hosted materials.json.");
        await window.BomStore.writeGlobalMaster(r.materials, null);
        await window.BomSheets.applyMaterialList(r.materials, { syncRemovals: false, projectMeta: null });
        await refreshAll();
        showGlobalMsg(`Pulled ${r.materials.length} materials from GitHub.`, "ok");
      } catch (err) { showGlobalMsg(err?.message || String(err), "err"); }
      finally { setBusy(false); }
    });
    $("btnGhPush").addEventListener("click", async () => {
      if (state.wb.projectMode) return;
      setBusy(true);
      try {
        await checkDirectMasterEdits();
        const fresh = await window.BomSheets.listMaterials();
        await window.BomStore.pushGitHub(fresh.rows, `Update materials.json (${fresh.rows.length} items)`);
        showGlobalMsg(`Published ${fresh.rows.length} materials to GitHub.`, "ok");
      } catch (err) { showGlobalMsg(err?.message || String(err), "err"); }
      finally { setBusy(false); }
    });

    // Convert tab controls
    $("segFlatten").addEventListener("click", () => setConvMode("flatten"));
    $("segRestore").addEventListener("click", () => setConvMode("restore"));
    $("btnOpenCopy1").addEventListener("click", openWorkbookCopy);
    $("btnOpenCopy2").addEventListener("click", openWorkbookCopy);
    $("chkCopyConfirmed").addEventListener("change", updateConvertGuards);
    $("selRestProj").addEventListener("change", updateConvertGuards);

    document.querySelectorAll('input[name="flatMode"]').forEach((r) => {
      r.addEventListener("change", () => {
        document.querySelectorAll("#modeGroup .radio-card").forEach((c) => {
          c.classList.toggle("selected", c.querySelector("input").checked);
        });
      });
    });

    document.querySelectorAll('input[name="restMatSrc"]').forEach((r) => {
      r.addEventListener("change", () => {
        document.querySelectorAll('[data-rsrc]').forEach((c) => {
          c.classList.toggle("selected", c.querySelector("input").checked);
        });
        updateConvertGuards();
      });
    });

    $("btnPickRestProj").addEventListener("click", () => $("fileRestProj").click());
    $("fileRestProj").addEventListener("change", async (e) => {
      const f = e.target.files?.[0];
      e.target.value = "";
      if (!f) return;
      try {
        const parsed = window.BomStore.parseJsonText(await f.text());
        state.pickedRestProj = {
          id: parsed.id,
          projectName: parsed.projectName || f.name.replace(/\.json$/i, ""),
          materials: parsed.materials,
          lastChange: parsed.lastChange
        };
        $("restFileDesc").textContent = `${state.pickedRestProj.projectName} (${parsed.materials.length} items)`;
        document.querySelector('input[name="restMatSrc"][value="project-file"]').checked = true;
        document.querySelectorAll('[data-rsrc]').forEach((c) => c.classList.toggle("selected", c.dataset.rsrc === "project-file"));
        updateConvertGuards();
      } catch (err) {
        showGlobalMsg(`Invalid JSON file: ${err?.message || err}`, "err");
      }
    });

    // Execute Flatten
    $("btnFlatten").addEventListener("click", async () => {
      const sheets = getChecked("flatSheetList");
      const mode = document.querySelector('input[name="flatMode"]:checked')?.value || "dynamic365";
      const compiled = $("chkCompiled").checked;
      clearConvLog();
      setBusy(true);
      try {
        if (state.wb.projectMode) {
          await checkDirectMasterEdits();
          // Always ensure latest project list is persisted in PC storage before Master sheet is removed
          await window.BomStore.saveProjectList(state.wb.projectMode, state.materials, null);
          if ($("chkAutoSaveProj").checked) {
            appendConvLog(`Auto-saving project materials JSON (${state.wb.projectMode.projectName})...`);
            const json = window.BomStore.serializeProject(state.wb.projectMode.projectName, state.materials, null, state.wb.projectMode.id);
            await window.BomStore.saveJsonFile(`project-materials-${window.BomStore.slugify(state.wb.projectMode.projectName)}.json`, json);
          }
        }
        const res = await window.BomConvert.flattenWorkbook({ sheets, mode, compiled }, appendConvLog);
        $("chkCopyConfirmed").checked = false;
        await refreshAll();
        showGlobalMsg(
          `Flattened ${res.converted} sheet(s)${res.isPartial ? " (Master kept for remaining automated sheets)" : ""}.`,
          "ok"
        );
      } catch (err) {
        appendConvLog(`Error: ${err?.message || err}`);
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });

    // Execute Restore Automated
    $("btnRestore").addEventListener("click", async () => {
      const sheets = getChecked("restSheetList");
      const materialSource = document.querySelector('input[name="restMatSrc"]:checked')?.value || "global";
      clearConvLog();
      setBusy(true);
      try {
        const res = await window.BomConvert.restoreWorkbook({
          sheets,
          materialSource,
          projectId: $("selRestProj").value,
          projectData: state.pickedRestProj,
          removeCompiled: $("chkRemoveCompiled").checked
        }, appendConvLog);
        await refreshAll();
        showGlobalMsg(`Restored ${res.restored} automated BOM sheet(s).`, "ok");
      } catch (err) {
        appendConvLog(`Error: ${err?.message || err}`);
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });

    // Execute Plain-Sheet Purchase-List Name Repair (Excel 365)
    $("btnRepairPlain").addEventListener("click", async () => {
      const sheets = getChecked("restSheetList");
      clearConvLog();
      setBusy(true);
      try {
        const res = await window.BomConvert.repairPlainSheets({ sheets }, appendConvLog);
        showGlobalMsg(`Repaired Purchase-List name formulas on ${res.fixedCount} plain sheet(s).`, "ok");
      } catch (err) {
        appendConvLog(`Error: ${err?.message || err}`);
        showGlobalMsg(err?.message || String(err), "err");
      } finally {
        setBusy(false);
      }
    });
  }

  /* ---------------- Boot ---------------- */
  function init() {
    wireEvents();
    if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
      Office.onReady(() => refreshAll());
    } else {
      refreshAll();
    }
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
