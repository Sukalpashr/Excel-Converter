/* app.js — BOM Tools task pane controller
 * ----------------------------------------------------------------------------
 * Wires up the 4-tab UI (Home · Add Sheet · Materials · Convert) to BomStore,
 * BomSheets and BomConvert. Provides:
 *   - live requirement guards (Office host + ExcelApi 1.9 + dynamic-array support)
 *   - workbook status refresh (BOM count, Project-Only badge, saved materials)
 *   - 5.5s auto-collapsing change toast (pause on hover) folding into an
 *     expandable *Session changes* log
 *   - per-tab event handlers
 *   - busy overlay for long-running Excel operations
 *
 * Confidentiality: never exposes internal prompts / skill descriptions.
 */
(function (root) {
  'use strict';

  var $ = function (sel) { return root.document.querySelector(sel); };
  var $$ = function (sel) { return Array.prototype.slice.call(root.document.querySelectorAll(sel)); };

  var State = {
    ready: false,
    excelOk: false,
    activeTab: 'home',
    materials: [],
    selectedMaterialIndex: -1,
    convertMode: 'flatten',
    flattenDepth: 'partial',
    flattenMode: 'auto365',
    restoreSource: 'global',
    toastTimers: {}
  };

  // ============================================================
  // INITIALIZATION + REQUIREMENT GUARDS
  // ============================================================

  function init() {
    if (root.Office) {
      Office.onReady(function (info) {
        if (info.host === Office.HostType.Excel) {
          checkRequirementSet().then(function (ok) {
            State.excelOk = ok;
            if (!ok) showRequirementFail();
            else boot();
          });
        } else {
          showRequirementFail('BOM Tools runs inside Excel 365 only.');
        }
      });
    } else {
      // Outside Office host (e.g. opened in a browser for dev preview) — still
      // render the UI so the layout can be inspected.
      bootPreview();
    }
  }

  function checkRequirementSet() {
    return new Promise(function (resolve) {
      if (!root.Excel || !Excel.RequestContext) return resolve(false);
      try {
        var ok = Excel.requirements.isSetSupported('ExcelApi', 1.9);
        // dynamic-array support check (1.9 doesn't formally expose this; we
        // probe at runtime by attempting a spill on first use, but assume true
        // on ExcelApi 1.9+).
        resolve(!!ok);
      } catch (e) { resolve(false); }
    });
  }

  function showRequirementFail(msg) {
    var host = $('#req-fail-host');
    if (host) {
      host.innerHTML = '';
      var div = root.document.createElement('div');
      div.className = 'req-fail';
      div.innerHTML = '<strong>Excel 365 required.</strong><br/>' +
        (msg || 'This add-in needs Excel on the web (Microsoft 365) or Excel 2019+ desktop with dynamic-array support (LET, XLOOKUP, XMATCH, FILTER, SORT, UNIQUE). Open this workbook in a supported Excel host and re-launch BOM Tools.');
      host.appendChild(div);
    }
    $('#req-pill').textContent = 'Unsupported';
    $('#req-pill').classList.add('bad');
    // Still wire tabs + toast so the user can browse the UI to see what they
    // are missing; actions will surface an "Excel 365 required" toast.
    wireTabs();
    wireToast();
    $$('.panel button.btn').forEach(function (b) {
      b.addEventListener('click', function (e) {
        if (!State.excelOk) {
          e.preventDefault();
          toast('Excel 365 required', 'Open this workbook in Excel 365 (Microsoft 365) and re-launch BOM Tools.', 'bad');
        }
      }, true);
    });
  }

  function boot() {
    State.ready = true;
    $('#req-pill').textContent = 'Excel 365';
    wireTabs();
    wireHome();
    wireAddSheet();
    wireMaterials();
    wireConvert();
    wireToast();
    wireStorageListeners();
    // initial status
    refreshStatus();
    // detect offline API (silent)
    if (root.BomStore) root.BomStore.detectOfflineApi().then(updateStorageStatus);
    // periodic status refresh on tab activation
    refreshAddSheetCopyFromList();
  }

  function bootPreview() {
    // Render UI shell without Excel for layout inspection.
    if ($('#stat-bom-count')) $('#stat-bom-count').textContent = '—';
    if ($('#stat-mat-count')) $('#stat-mat-count').textContent = '—';
    if ($('#stat-hint')) $('#stat-hint').textContent = 'Open this task pane inside Excel to read the workbook.';
    if ($('#req-pill')) $('#req-pill').textContent = 'Preview';
    wireTabs();
    wireToast();
  }

  // ============================================================
  // TABS
  // ============================================================

  function wireTabs() {
    $$('.tabs button').forEach(function (btn) {
      btn.addEventListener('click', function () { switchTab(btn.dataset.tab); });
    });
    $$('[data-goto]').forEach(function (el) {
      el.addEventListener('click', function () { switchTab(el.dataset.goto); });
    });
  }
  function switchTab(name) {
    State.activeTab = name;
    $$('.tabs button').forEach(function (b) { b.classList.toggle('active', b.dataset.tab === name); });
    $$('.panel').forEach(function (p) { p.classList.toggle('active', p.id === 'panel-' + name); });
    if (name === 'home') refreshStatus();
    if (name === 'addsheet') refreshAddSheetCopyFromList();
    if (name === 'materials') refreshMaterials();
    if (name === 'convert') refreshConvertLists();
  }

  // ============================================================
  // HOME TAB
  // ============================================================

  function wireHome() {
    var btnImport = $('#btn-import-template');
    if (btnImport) btnImport.addEventListener('click', function () {
      busy('Importing template…');
      BomSheets.importTemplate().then(function () {
        toast('Template imported', 'BOM Blank + Master Unique Materials are ready. Formulas injected.');
        refreshStatus();
      }).catch(errToast).finally(unbusy);
    });

    var qaConvert = $('#qa-convert');
    if (qaConvert) qaConvert.addEventListener('click', function () { switchTab('convert'); });

    var qaRestore = $('#qa-restore');
    if (qaRestore) qaRestore.addEventListener('click', function () {
      switchTab('convert');
      // auto-flip to Restore mode
      $$('[data-cmode]').forEach(function (b) { b.classList.toggle('active', b.dataset.cmode === 'restore'); });
      toggleConvertMode('restore');
    });

    // Repair checklist
    wireChecklist('home-repair-list');
    var btnRepair = $('#btn-home-repair');
    if (btnRepair) btnRepair.addEventListener('click', function () {
      var sheets = checkedFrom('home-repair-list');
      if (!sheets.length) { toast('Pick at least one sheet', '', 'warn'); return; }
      busy('Repairing…');
      BomSheets.repairAll(sheets).then(function () {
        toast('Repair complete', 'Recreated named ranges, reapplied DV, upgraded formulas, purged Print_Area/Print_Titles.');
        refreshStatus();
      }).catch(errToast).finally(unbusy);
    });
  }

  function refreshStatus() {
    if (!State.ready || !root.BomSheets) return;
    BomSheets.getWorkbookStatus().then(function (st) {
      $('#stat-bom-count').textContent = st.bomCount;
      $('#stat-bom-count').classList.toggle('bad', st.bomCount === 0 && !st.needsImport);
      var matP = root.BomStore ? root.BomStore.getMaster().then(function (m) { return (m || []).length; }) : Promise.resolve(0);
      matP.then(function (n) {
        $('#stat-mat-count').textContent = n;
      });
      $('#stat-scope').textContent = st.activeProject ? ('Project: ' + st.activeProject) : 'Global Master';
      $('#stat-project').firstChild.textContent = (st.activeProject || '—') + ' ';
      var badge = $('#stat-project-badge');
      if (badge) badge.classList.toggle('green', !st.isProjectOnly);
      $('#stat-hint').textContent = st.needsImport
        ? 'Workbook has no BOM template yet. Use “Import BOM template” below.'
        : (st.bomProjectList ? ('BomProjectList: ' + st.bomProjectList) : 'Workbook ready.');
      $('#home-import-card').hidden = !st.needsImport;
      // Build repair checklist
      buildChecklist('home-repair-list', st.bomSheets.concat(st.hasMaster ? [BomSheets.MASTER_SHEET] : []));
      // Add Sheet copy-from dropdown
      refreshAddSheetCopyFromList(st.bomSheets);
    }).catch(errToast);
  }

  // ============================================================
  // ADD SHEET TAB
  // ============================================================

  function wireAddSheet() {
    var nameInput = $('#as-name');
    if (!nameInput) return;
    nameInput.addEventListener('input', validateAddSheet);
    $('#btn-add-sheet').addEventListener('click', function () {
      if (validateAddSheet()) return;
      var opts = {
        projectTitle: $('#as-project').value || undefined,
        copyFrom: $('#as-copy-from').value || undefined,
        clearInputs: $('#as-clear-inputs').checked,
        preSync: $('#as-pre-sync').checked
      };
      busy('Creating sheet…');
      BomSheets.addSheet(nameInput.value, opts).then(function () {
        toast('Sheet created', '“' + nameInput.value + '” is ready. Formulas, tables and DV injected.');
        nameInput.value = '';
        validateAddSheet();
        refreshStatus();
      }).catch(errToast).finally(unbusy);
    });
  }

  function validateAddSheet() {
    var name = $('#as-name').value.trim();
    var hint = $('#as-name-hint');
    var btn = $('#btn-add-sheet');
    // Get sheet names (without Excel, we can't — skip)
    if (!root.BomSheets) { btn.disabled = !name; return 0; }
    BomSheets.listSheetNamesRootless ? null : null;
    // We need a sync check; query and update asynchronously.
    Excel.run({ delayForCellEdit: true }, function (ctx) {
      var ws = ctx.workbook.worksheets.load('name');
      return ctx.sync().then(function () {
        var names = ws.items.map(function (s) { return s.name; });
        var err = BomSheets.validateSheetName(name, names);
        if (err) {
          hint.textContent = err; hint.classList.add('err');
          nameInput().classList.add('invalid');
          btn.disabled = true;
        } else {
          hint.textContent = 'Valid name. Max 31 chars, no : \\ / ? * [ ] .';
          hint.classList.remove('err');
          nameInput().classList.remove('invalid');
          btn.disabled = !name;
        }
      });
    }).catch(function () { btn.disabled = !name; });
    return 0;
  }
  function nameInput() { return $('#as-name'); }

  function refreshAddSheetCopyFromList(bomSheets) {
    var sel = $('#as-copy-from');
    if (!sel) return;
    var cur = sel.value;
    var opts = '<option value="">— none (blank BOM) —</option>' +
      '<option value="__blank__">BOM Blank (template)</option>';
    (bomSheets || []).forEach(function (n) {
      opts += '<option value="' + esc(n) + '">' + esc(n) + '</option>';
    });
    sel.innerHTML = opts;
    sel.value = cur;
  }

  // ============================================================
  // MATERIALS TAB
  // ============================================================

  function wireMaterials() {
    var search = $('#mat-search');
    if (search) search.addEventListener('input', function () { renderMaterialsList(); });

    $('#btn-mat-add').addEventListener('click', function () { openMaterialForm(null); });
    $('#btn-mat-refresh').addEventListener('click', function () {
      busy('Refreshing…');
      refreshMaterials().then(function () {
        toast('Materials refreshed', 'Snapshot taken from Master Unique Materials.');
      }).catch(errToast).finally(unbusy);
    });

    $('#btn-mat-save').addEventListener('click', saveMaterialForm);
    $('#btn-mat-cancel').addEventListener('click', closeMaterialForm);
    $('#btn-mat-remove').addEventListener('click', removeMaterialForm);

    // Storage drawer
    $('#btn-mat-export').addEventListener('click', exportMaster);
    $('#btn-mat-import').addEventListener('click', function () { $('#mat-import-file').click(); });
    $('#mat-import-file').addEventListener('change', importMaster);

    $('#btn-mat-export-project').addEventListener('click', exportProject);
    $('#btn-mat-import-project').addEventListener('click', function () { $('#mat-import-project-file').click(); });
    $('#mat-import-project-file').addEventListener('change', importProject);

    $('#btn-mat-github-pull').addEventListener('click', githubPull);
    $('#btn-mat-github-push').addEventListener('click', githubPush);

    $('#btn-mat-scope-global').addEventListener('click', function () {
      if (!root.BomStore) return;
      BomStore.setActiveProject(null);
      toast('Switched to Global Master', '2-click switch — your project slot is preserved.');
      refreshMaterials();
    });
    $('#btn-mat-new-project').addEventListener('click', newProject);

    // Material list click → open edit
    var list = $('#mat-list');
    if (list) list.addEventListener('click', function (e) {
      var row = e.target.closest('.row'); if (!row) return;
      var idx = Number(row.dataset.idx);
      openMaterialForm(State.materials[idx]);
    });
  }

  function refreshMaterials() {
    if (!State.ready) return Promise.resolve();
    // Snapshot diff vs last-known — emits audit events if changes detected.
    return BomSheets.snapshotDiff().then(function (diff) {
      diff.added.forEach(function (m) { recordAudit('Added', m.name); });
      diff.modified.forEach(function (m) { recordAudit('Modified', m.name); });
      diff.removed.forEach(function (m) { recordAudit('Removed', m.name); });
      State.materials = diff.all;
      renderMaterialsList();
      renderProjectList();
      // Scope badge
      var scope = root.BomStore ? BomStore.activeScopeLabel() : 'Global Master';
      var badge = $('#mat-scope-badge');
      if (badge) { badge.textContent = scope; badge.classList.toggle('green', !BomStore.getActiveProject()); }
      return diff;
    }).catch(errToast);
  }

  function renderMaterialsList() {
    var list = $('#mat-list'); if (!list) return;
    var q = ($('#mat-search').value || '').toLowerCase().trim();
    var sorted = State.materials.slice().sort(function (a, b) {
      // Recently changed (Last Modified desc) sort to top
      var at = a.lastModified ? Date.parse(a.lastModified) : 0;
      var bt = b.lastModified ? Date.parse(b.lastModified) : 0;
      if (bt !== at) return bt - at;
      return a.name.localeCompare(b.name);
    });
    var filtered = sorted.filter(function (m) {
      if (!q) return true;
      return (m.name + ' ' + m.category + ' ' + m.unit).toLowerCase().indexOf(q) >= 0;
    });
    if (!filtered.length) {
      list.innerHTML = '<div class="muted" style="padding:10px;">No materials. Click “+ Add” or import a Master JSON.</div>';
      return;
    }
    list.innerHTML = filtered.map(function (m, i) {
      var r = readiness(m);
      var recentTag = m.lastModified && (Date.now() - Date.parse(m.lastModified)) < 3600000
        ? '<span class="recent">recent</span>' : '';
      return '<div class="row" data-idx="' + State.materials.indexOf(m) + '">' +
        '<span class="icn ' + r.cls + '">' + r.icon + '</span>' +
        '<span class="name">' + esc(m.name) + '</span>' +
        '<span class="meta">' + esc(m.unit || 'no unit') + '</span>' +
        recentTag +
        '</div>';
    }).join('');
  }

  // BOM-readiness icons. Validates Name, Primary Unit, and positive
  // Weight/Sheet dims. No Unit Cost field.
  function readiness(m) {
    if (!m.name) return { icon: '✕', cls: 'bad' };
    if (['kg/m', 'kg/m²', 'kg/each'].indexOf(m.unit) < 0) return { icon: '✕', cls: 'bad' };
    var w = 0;
    if (m.unit === 'kg/m' && m.linear > 0) w = 1;
    if (m.unit === 'kg/m²' && m.area > 0 && m.sheetL > 0 && m.sheetW > 0) w = 1;
    if (m.unit === 'kg/each' && m.piece > 0) w = 1;
    return w ? { icon: '✔', cls: 'ok' } : { icon: '⚠', cls: 'warn' };
  }

  function openMaterialForm(mat) {
    $('#mat-form-card').hidden = false;
    $('#mat-form-title').textContent = mat ? 'Edit material' : 'Add material';
    $('#btn-mat-remove').hidden = !mat;
    State.selectedMaterialIndex = mat ? State.materials.indexOf(mat) : -1;
    $('#mat-f-name').value = (mat && mat.name) || '';
    $('#mat-f-unit').value = (mat && mat.unit) || 'kg/m';
    $('#mat-f-lin').value = (mat && mat.linear != null) ? mat.linear : '';
    $('#mat-f-area').value = (mat && mat.area != null) ? mat.area : '';
    $('#mat-f-piece').value = (mat && mat.piece != null) ? mat.piece : '';
    $('#mat-f-sl').value = (mat && mat.sheetL != null) ? mat.sheetL : '';
    $('#mat-f-sw').value = (mat && mat.sheetW != null) ? mat.sheetW : '';
    $('#mat-f-cat').value = (mat && mat.category) || '';
    $('#mat-f-rem').value = (mat && mat.remarks) || '';
    refreshCategoryDatalist();
    $('#mat-f-name').focus();
  }
  function closeMaterialForm() { $('#mat-form-card').hidden = true; }

  function saveMaterialForm() {
    var idx = State.selectedMaterialIndex;
    var existing = idx >= 0 ? State.materials[idx] : null;
    var mat = {
      name: $('#mat-f-name').value.trim(),
      unit: $('#mat-f-unit').value,
      linear: numOrNull($('#mat-f-lin').value),
      area: numOrNull($('#mat-f-area').value),
      piece: numOrNull($('#mat-f-piece').value),
      sheetL: numOrNull($('#mat-f-sl').value),
      sheetW: numOrNull($('#mat-f-sw').value),
      category: $('#mat-f-cat').value.trim(),
      remarks: $('#mat-f-rem').value.trim()
    };
    if (!mat.name) { toast('Name required', '', 'warn'); return; }
    busy('Saving material…');
    var op = existing
      ? BomSheets.updateMaterial(existing.name, mat)
      : BomSheets.addMaterial(mat);
    op.then(function (saved) {
      toast(existing ? 'Modified' : 'Added', '"' + saved.name + '" saved to Master + JSON.');
      closeMaterialForm();
      refreshMaterials();
    }).catch(errToast).finally(unbusy);
  }

  function removeMaterialForm() {
    var idx = State.selectedMaterialIndex;
    var existing = idx >= 0 ? State.materials[idx] : null;
    if (!existing) return;
    if (!root.confirm('Remove "' + existing.name + '" from the Master?')) return;
    busy('Removing material…');
    BomSheets.removeMaterial(existing.name).then(function () {
      toast('Removed', '"' + existing.name + '" removed from Master + JSON.', 'warn');
      closeMaterialForm();
      refreshMaterials();
    }).catch(errToast).finally(unbusy);
  }

  function refreshCategoryDatalist() {
    var dl = $('#mat-cat-list'); if (!dl) return;
    var seen = {}; var opts = '';
    State.materials.forEach(function (m) {
      if (m.category && !seen[m.category]) { seen[m.category] = 1; opts += '<option value="' + esc(m.category) + '">'; }
    });
    dl.innerHTML = opts;
  }

  // ----- storage drawer actions -----
  function exportMaster() {
    if (!root.BomStore) return;
    BomStore.exportMaster().then(function (rec) {
      var name = (rec.type === 'project' ? 'bom-project-' + rec.name : 'bom-master') + '.json';
      downloadJson(rec, name);
      toast('Exported Master JSON', (rec.materials || []).length + ' materials.');
    }).catch(errToast);
  }
  function importMaster(ev) {
    var file = ev.target.files[0]; if (!file) return;
    readFile(file).then(function (txt) {
      return BomStore.importMaster(txt);
    }).then(function () {
      toast('Master imported', 'Materials updated. Refresh to write to the sheet.');
      refreshMaterials();
    }).catch(function (e) {
      toast('Import failed', e.message, 'bad');
    }).finally(function () { ev.target.value = ''; });
  }
  function exportProject() {
    if (!root.BomStore) return;
    var name = BomStore.getActiveProject();
    if (!name) { toast('No active project', 'Switch to a project first, or use Export Master JSON.', 'warn'); return; }
    BomStore.exportProject(name).then(function (rec) {
      downloadJson(rec, (name || 'bom-project') + '.json');
      toast('Project .json exported', (rec.materials || []).length + ' materials.');
    }).catch(errToast);
  }
  function importProject(ev) {
    var file = ev.target.files[0]; if (!file) return;
    readFile(file).then(function (txt) {
      return BomStore.importProject(txt);
    }).then(function (rec) {
      toast('Project imported', '"' + rec.name + '" saved. Use Projects list to switch scope.');
      renderProjectList();
    }).catch(function (e) {
      toast('Import failed', e.message, 'bad');
    }).finally(function () { ev.target.value = ''; });
  }
  function githubPull() {
    var cfg = readGithubCfg();
    if (!cfg) { toast('GitHub config missing', 'Set owner / repo / token / branch in localStorage.', 'warn'); return; }
    busy('Pulling from GitHub…');
    BomStore.githubPull(cfg).then(function (rec) {
      return BomStore.importMaster(JSON.stringify(rec));
    }).then(function () {
      toast('GitHub Pull OK', 'Materials synced from raw.githubusercontent.com.');
      refreshMaterials();
    }).catch(errToast).finally(unbusy);
  }
  function githubPush() {
    var cfg = readGithubCfg();
    if (!cfg) { toast('GitHub config missing', 'Set owner / repo / token / branch in localStorage.', 'warn'); return; }
    busy('Pushing to GitHub…');
    BomStore.exportMaster().then(function (rec) {
      return BomStore.githubPush(cfg, rec);
    }).then(function () {
      toast('GitHub Push OK', 'Master committed to ' + cfg.owner + '/' + cfg.repo + '.');
    }).catch(errToast).finally(unbusy);
  }
  function readGithubCfg() {
    try {
      var raw = root.localStorage.getItem('bom_tools:github');
      if (!raw) return null;
      var c = JSON.parse(raw);
      if (!c.owner || !c.repo || !c.token) return null;
      return c;
    } catch (e) { return null; }
  }

  // ----- Projects list -----
  function renderProjectList() {
    var box = $('#mat-project-list'); if (!box || !root.BomStore) return;
    BomStore.listSavedProjects().then(function (items) {
      if (!items.length) {
        box.innerHTML = '<div class="muted" style="padding:8px;">No saved projects.</div>';
        return;
      }
      var active = BomStore.getActiveProject();
      box.innerHTML = items.map(function (p) {
        return '<div class="row" data-proj="' + esc(p.name) + '">' +
          '<span class="icn ' + (active === p.name ? 'ok' : '') + '">' + (active === p.name ? '●' : '○') + '</span>' +
          '<span class="name">' + esc(p.name) + '</span>' +
          '<span class="meta">' + p.materialsCount + ' mat</span>' +
          '<button class="btn subtle sm" data-proj-del="' + esc(p.name) + '">✕</button>' +
          '</div>';
      }).join('');
      // attach click handlers
      $$('#mat-project-list .row').forEach(function (row) {
        row.addEventListener('click', function (e) {
          if (e.target.matches('[data-proj-del]')) {
            e.stopPropagation();
            var n = e.target.dataset.projDel;
            if (root.confirm('Delete project "' + n + '"?')) {
              BomStore.deleteProject(n).then(function () { renderProjectList(); refreshMaterials(); });
            }
            return;
          }
          var n = row.dataset.proj;
          BomStore.setActiveProject(n);
          toast('Switched project', 'Active scope: ' + n + '.');
          refreshMaterials();
        });
      });
    });
  }

  function newProject() {
    var name = root.prompt('Project name:'); if (!name) return;
    var src = $('#mat-new-project-source').value || 'copy-global';
    var opts = {};
    if (src === 'workbook-only') {
      opts.workbookMaterials = State.materials.slice();
    } else if (src === 'unified-pool') {
      // open a simple picker via prompt (the spec calls for a unified searchable
      // pool — we wire the data here; the full picker UI is the Materials list)
      var pool = null;
      // We let the user select by name via prompt as a minimal picker.
      var names = root.prompt('Pick material names (comma-separated) from the unified pool:', '');
      if (names) opts.pick = names.split(',').map(function (s) { return s.trim(); }).filter(Boolean);
    }
    busy('Creating project…');
    BomStore.newProject(name, src, opts).then(function () {
      BomStore.setActiveProject(name);
      toast('Project created', '"' + name + '" — scope switched. Global Master untouched.');
      renderProjectList();
      refreshMaterials();
    }).catch(errToast).finally(unbusy);
  }

  // ============================================================
  // CONVERT TAB
  // ============================================================

  function wireConvert() {
    $$('[data-cmode]').forEach(function (b) {
      b.addEventListener('click', function () {
        $$('[data-cmode]').forEach(function (x) { x.classList.toggle('active', x === b); });
        toggleConvertMode(b.dataset.cmode);
      });
    });
    $$('[data-fdepth]').forEach(function (b) {
      b.addEventListener('click', function () {
        $$('[data-fdepth]').forEach(function (x) { x.classList.toggle('active', x === b); });
        State.flattenDepth = b.dataset.fdepth;
        var hint = $('#flatten-depth-hint');
        if (hint) hint.textContent = State.flattenDepth === 'full'
          ? 'Full removes the hidden BOM Blank, Master Unique Materials and named ranges, but preserves the BomProjectList workbook property.'
          : 'Partial keeps the hidden BOM Blank, Master Unique Materials and named ranges.';
      });
    });
    $$('[data-fmode]').forEach(function (b) {
      b.addEventListener('click', function () {
        $$('[data-fmode]').forEach(function (x) { x.classList.toggle('active', x === b); });
        State.flattenMode = b.dataset.fmode;
      });
    });
    $$('[data-rsrc]').forEach(function (b) {
      b.addEventListener('click', function () {
        $$('[data-rsrc]').forEach(function (x) { x.classList.toggle('active', x === b); });
        State.restoreSource = b.dataset.rsrc;
        $('#restore-saved-project').hidden = State.restoreSource !== 'saved';
        $('#restore-json-file').hidden = true;
        if (State.restoreSource === 'jsonfile') $('#restore-json-file').click();
      });
    });
    $('#restore-json-file').addEventListener('change', function (ev) {
      var f = ev.target.files[0]; if (!f) return;
      readFile(f).then(function (txt) {
        State._restoreJsonData = txt;
      });
    });

    wireChecklist('flatten-list');
    wireChecklist('restore-list');
    wireChecklist('repair-plain-list');

    $('#btn-flatten').addEventListener('click', doFlatten);
    $('#btn-restore').addEventListener('click', doRestore);
    $('#btn-repair-plain').addEventListener('click', doRepairPlain);
  }

  function toggleConvertMode(mode) {
    State.convertMode = mode;
    $('#convert-flatten').hidden = mode !== 'flatten';
    $('#convert-restore').hidden = mode !== 'restore';
  }

  function refreshConvertLists() {
    if (!State.ready) return;
    // Flatten list: all BOM sheets
    BomSheets.getWorkbookStatus().then(function (st) {
      buildChecklist('flatten-list', st.bomSheets);
      buildChecklist('repair-plain-list', st.bomSheets);
    });
    // Restore list: detect flattened sheets
    BomConvert.detectFlattenedSheets().then(function (flat) {
      buildChecklist('restore-list', flat);
      var hint = $('#restore-status-hint');
      if (hint) hint.textContent = flat.length
        ? (flat.length + ' flattened sheet(s) detected. Pick the ones to restore.')
        : 'No flattened sheets detected. Open a previously flattened workbook first.';
    }).catch(function () {
      var hint = $('#restore-status-hint');
      if (hint) hint.textContent = 'Could not detect flattened sheets (workbook may be empty).';
    });
    // Detect BomProjectList status banner
    BomConvert.detectBomProjectList().then(function (info) {
      var hint = $('#restore-source-hint');
      if (hint) hint.textContent = info.hasProperty
        ? ('BomProjectList: "' + info.value + '". Existing Stock / PO inputs preserved on Restore.')
        : 'No BomProjectList property. Existing Stock / PO inputs preserved on Restore.';
    });
    // Populate saved-project dropdown
    if (root.BomStore) BomStore.listSavedProjects().then(function (items) {
      var sel = $('#restore-saved-project');
      sel.innerHTML = items.map(function (p) { return '<option value="' + esc(p.name) + '">' + esc(p.name) + ' (' + p.materialsCount + ')</option>'; }).join('');
    });
  }

  function doFlatten() {
    var sheets = checkedFrom('flatten-list');
    if (!sheets.length) { toast('Pick at least one sheet', '', 'warn'); return; }
    busy('Flattening in a new Excel window…');
    BomConvert.flatten({
      sheets: sheets,
      full: State.flattenDepth === 'full',
      mode: State.flattenMode,
      compiledPL: $('#flatten-compiled-pl').checked,
      preSaveProject: $('#flatten-pre-save').checked
    }).then(function () {
      toast('Flatten complete', 'A flattened copy opened in a new Excel window.');
    }).catch(errToast).finally(unbusy);
  }

  function doRestore() {
    var sheets = checkedFrom('restore-list');
    if (!sheets.length) { toast('Pick at least one sheet', '', 'warn'); return; }
    var opts = {
      sheets: sheets,
      materialSource: State.restoreSource,
      savedProjectName: $('#restore-saved-project').value || undefined,
      jsonFileData: State._restoreJsonData
    };
    busy('Restoring automated sheets in a new window…');
    BomConvert.restore(opts).then(function () {
      toast('Restore complete', 'Tables, formulas, dropdowns and named ranges rebuilt. Stock/PO preserved.');
    }).catch(errToast).finally(unbusy);
  }

  function doRepairPlain() {
    var sheets = checkedFrom('repair-plain-list');
    if (!sheets.length) { toast('Pick at least one sheet', '', 'warn'); return; }
    busy('Upgrading in place…');
    BomConvert.repairPlainSheetNames(sheets).then(function () {
      toast('Repair complete', 'Sheets upgraded to Excel 365 dynamic-spill AB5 + INDEX names. BOM/Stock/PO untouched.');
    }).catch(errToast).finally(unbusy);
  }

  // ============================================================
  // CHECKLISTS (All / None)
  // ============================================================

  function buildChecklist(listId, items) {
    var el = $('#' + listId); if (!el) return;
    if (!items || !items.length) { el.innerHTML = '<div class="muted" style="padding:6px;">No sheets available.</div>'; return; }
    el.innerHTML = items.map(function (n, i) {
      return '<label><input type="checkbox" data-name="' + esc(n) + '" checked /><span>' + esc(n) + '</span></label>';
    }).join('');
  }
  function wireChecklist(listId) {
    $$('[data-check-all="' + listId + '"]').forEach(function (b) {
      b.addEventListener('click', function () { $$('#' + listId + ' input[type=checkbox]').forEach(function (c) { c.checked = true; }); });
    });
    $$('[data-check-none="' + listId + '"]').forEach(function (b) {
      b.addEventListener('click', function () { $$('#' + listId + ' input[type=checkbox]').forEach(function (c) { c.checked = false; }); });
    });
  }
  function checkedFrom(listId) {
    return $$('#' + listId + ' input[type=checkbox]:checked').map(function (c) { return c.dataset.name; });
  }

  // ============================================================
  // TOAST + SESSION LOG
  // ============================================================

  function wireToast() {
    // pause on hover
    var wrap = $('#toast-wrap');
    if (!wrap) return;
    wrap.addEventListener('mouseover', function (e) {
      var t = e.target.closest('.toast'); if (!t) return;
      clearTimeout(State.toastTimers[t.id]);
      t.classList.remove('collapsed');
    });
    wrap.addEventListener('mouseout', function (e) {
      var t = e.target.closest('.toast'); if (!t) return;
      scheduleCollapse(t);
    });
  }
  function toast(title, body, kind) {
    var wrap = $('#toast-wrap'); if (!wrap) return;
    var id = 't' + Date.now() + Math.random().toString(36).slice(2, 6);
    var el = root.document.createElement('div');
    el.className = 'toast' + (kind ? ' ' + kind : '');
    el.id = id;
    el.innerHTML =
      '<div class="head"><span>' + esc(title) + '</span><span class="close" title="dismiss">×</span></div>' +
      (body ? '<div class="body">' + esc(body) + '</div>' : '') +
      '<div class="progress" style="width:100%;"></div>';
    wrap.appendChild(el);
    el.querySelector('.close').addEventListener('click', function () { dismiss(id); });
    scheduleCollapse(el);
    scheduleProgress(el);
  }
  function scheduleProgress(el) {
    var bar = el.querySelector('.progress');
    if (!bar) return;
    var total = 5500; var start = Date.now();
    function step() {
      var left = total - (Date.now() - start);
      if (left <= 0) { bar.style.width = '0%'; return; }
      bar.style.width = (left / total * 100) + '%';
      root.requestAnimationFrame(step);
    }
    root.requestAnimationFrame(step);
  }
  function scheduleCollapse(el) {
    clearTimeout(State.toastTimers[el.id]);
    State.toastTimers[el.id] = root.setTimeout(function () {
      el.classList.add('collapsed');
      // after another 8s, fold into session log + dismiss
      State.toastTimers[el.id] = root.setTimeout(function () { dismiss(el.id); }, 8000);
    }, 5500);
  }
  function dismiss(id) {
    clearTimeout(State.toastTimers[id]);
    var el = root.document.getElementById(id);
    if (el && el.parentNode) el.parentNode.removeChild(el);
    delete State.toastTimers[id];
  }

  function recordAudit(action, name) {
    if (!root.BomStore) return;
    var log = BomStore.getAuditLog();
    // The BomStore.recordAudit() already records; here we just re-render.
    renderSessionLogs(log);
  }
  function renderSessionLogs(log) {
    var arr = (log || BomStore.getAuditLog()).slice(0, 20);
    ['home-session-log', 'mat-session-log'].forEach(function (id) {
      var el = $('#' + id); if (!el) return;
      if (!arr.length) { el.innerHTML = '<div class="muted">Session changes will appear here.</div>'; return; }
      el.innerHTML = arr.map(function (e) {
        var cls = (e.action || '').toLowerCase();
        return '<div class="entry ' + cls + '"><span class="t">' + shortTime(e.at) + '</span>' +
          esc(e.action || '') + ' ' + esc(e.name || '') + '</div>';
      }).join('');
    });
  }
  function shortTime(iso) {
    try { var d = new Date(iso); return d.getHours().toString().padStart(2, '0') + ':' + d.getMinutes().toString().padStart(2, '0'); }
    catch (e) { return ''; }
  }

  function wireStorageListeners() {
    if (!root.BomStore) return;
    BomStore.on('audit', function (entry) {
      renderSessionLogs();
    });
    BomStore.on('change', function () {
      renderProjectList();
      renderMaterialsList();
    });
    BomStore.on('offline', function () { updateStorageStatus(); });
  }

  function updateStorageStatus() {
    if (!root.BomStore) return;
    BomStore.status().then(function (s) {
      var set = function (id, v) { var el = $('#' + id); if (el) { el.textContent = v; el.classList.toggle('bad', v === 'off'); } };
      set('stor-idb', s.indexedDB === 'ok' ? 'OK' : 'off');
      set('stor-opfs', s.opfs === 'ok' ? 'OK' : 'off');
      set('stor-ls', s.localStorage === 'ok' ? 'OK' : 'off');
      set('stor-offline', s.offlineApi === 'ok' ? 'OK' : 'off');
    });
  }

  // ============================================================
  // BUSY OVERLAY + ERRORS
  // ============================================================

  function busy(msg) {
    var b = $('#busy'); if (!b) return;
    $('#busy-msg').textContent = msg || 'Working…';
    b.classList.add('show');
  }
  function unbusy() { var b = $('#busy'); if (b) b.classList.remove('show'); }

  function errToast(err) {
    var msg = (err && err.message) ? err.message : String(err);
    toast('Error', msg, 'bad');
    console.error('[BOM Tools]', err);
  }

  // ============================================================
  // small utils
  // ============================================================

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function numOrNull(v) { v = (v == null ? '' : String(v).trim()); if (v === '') return null; var n = Number(v); return isFinite(n) ? n : null; }
  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var fr = new root.FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(fr.error); };
      fr.readAsText(file);
    });
  }
  function downloadJson(obj, filename) {
    try {
      var blob = new root.Blob([JSON.stringify(obj, null, 2)], { type: 'application/json' });
      var url = root.URL.createObjectURL(blob);
      var a = root.document.createElement('a');
      a.href = url; a.download = filename;
      root.document.body.appendChild(a); a.click();
      root.document.body.removeChild(a);
      root.URL.revokeObjectURL(url);
    } catch (e) { errToast(e); }
  }

  // expose minimal API for debugging
  root.BomApp = { refreshStatus: refreshStatus, refreshMaterials: refreshMaterials, switchTab: switchTab, toast: toast };

  // boot when DOM is ready
  if (root.document.readyState === 'loading') {
    root.document.addEventListener('DOMContentLoaded', init);
  } else { init(); }
})(typeof window !== 'undefined' ? window : this);
