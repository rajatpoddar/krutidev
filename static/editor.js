'use strict';

// ── Config from server ─────────────────────────────────
const CFG = window.EDITOR_CONFIG || { docId: null, access: 'edit', title: 'Untitled' };

// ── State ──────────────────────────────────────────────
let hot          = null;
let sheetsData   = {};
let sheetNames   = [];
let activeSheet  = null;
let currentFilename = CFG.title + '.xlsx';
let fontMode     = 'krutidev';
let editedCells  = new Set();
let unicodeCache = {};
let isModified   = false;
let docId        = CFG.docId;
let fileId       = null;
let shareAccess  = 'edit';
let autoSaveTimer = null;
let lastSaved    = null;
let mergeCells   = {};   // { sheetName: [ {row,col,rowspan,colspan}, ... ] }
let cellAligns   = {};   // { sheetName: { "row,col": "left"|"center"|"right" } }
let cellFontSizes = {};  // { sheetName: { "row,col": sizeNumber } }
let cellBold     = {};   // { sheetName: { "row,col": true } }
let cellItalic   = {};   // { sheetName: { "row,col": true } }
let cellUnderline= {};   // { sheetName: { "row,col": true } }
let cellBgColors = {};   // { sheetName: { "row,col": "#rrggbb" } }

// Save last known selection so ribbon buttons work after focus leaves grid
let _lastSel     = null;

// ── Socket.IO ──────────────────────────────────────────
const socket = io({ transports: ['websocket', 'polling'] });
let myName = 'User' + Math.floor(Math.random() * 900 + 100);

socket.on('connect', () => {
  if (docId) {
    socket.emit('join', { doc_id: docId, name: myName });
  }
});

socket.on('users_update', ({ users }) => {
  renderUsersBar(users);
});

socket.on('cell_change', ({ sheet, row, col, value }) => {
  // Apply remote change to local data without triggering our own emit
  if (sheetsData[sheet]) {
    while (row >= sheetsData[sheet].rows.length) sheetsData[sheet].rows.push([]);
    while (col >= sheetsData[sheet].rows[row].length) sheetsData[sheet].rows[row].push('');
    sheetsData[sheet].rows[row][col] = value;
  }
  if (sheet === activeSheet && hot) {
    hot.setDataAtCell(row, col, value, 'remote');
  }
});

socket.on('title_change', ({ title }) => {
  document.getElementById('titleInput').value = title;
  document.title = `${title} — Krutidev Editor`;
});

socket.on('cursor_move', ({ row, col, user }) => {
  if (!hot) return;
  // Briefly highlight the remote cursor cell
  const td = hot.getCell(row, col);
  if (td) {
    td.classList.add('remote-cursor');
    setTimeout(() => td && td.classList.remove('remote-cursor'), 1500);
  }
});

// ── DOM refs ───────────────────────────────────────────
const fileInput        = document.getElementById('fileInput');
const emptyState       = document.getElementById('emptyState');
const gridContainer    = document.getElementById('gridContainer');
const loadingOverlay   = document.getElementById('loadingOverlay');
const loadingText      = document.getElementById('loadingText');
const sheetTabsBar     = document.getElementById('sheetTabsBar');
const sheetTabs        = document.getElementById('sheetTabs');
const formulaBar       = document.getElementById('formulaBar');
const formulaInput     = document.getElementById('formulaInput');
const cellRef          = document.getElementById('cellRef');
const statusBar        = document.getElementById('statusBar');
const statusCells      = document.getElementById('statusCells');
const statusRows       = document.getElementById('statusRows');
const statusCols       = document.getElementById('statusCols');
const statusFont       = document.getElementById('statusFont');
const statusEdited     = document.getElementById('statusEdited');
const searchBar        = document.getElementById('searchBar');
const searchInput      = document.getElementById('searchInput');
const replaceInput     = document.getElementById('replaceInput');
const dropOverlay      = document.getElementById('dropOverlay');
const titleInput       = document.getElementById('titleInput');
const autosaveIndicator = document.getElementById('autosaveIndicator');

// ── Toast ──────────────────────────────────────────────
function toast(msg, type = 'info', ms = 3000) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = `toast show ${type}`;
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.className = 'toast'; }, ms);
}

// ── Upload ─────────────────────────────────────────────
function triggerUpload() { fileInput.click(); }
fileInput.addEventListener('change', () => { if (fileInput.files[0]) loadFile(fileInput.files[0]); });

document.addEventListener('dragover', e => { e.preventDefault(); dropOverlay.classList.add('active'); });
document.addEventListener('dragleave', e => { if (!e.relatedTarget) dropOverlay.classList.remove('active'); });
document.addEventListener('drop', e => {
  e.preventDefault(); dropOverlay.classList.remove('active');
  const f = e.dataTransfer.files[0]; if (f) loadFile(f);
});

async function loadFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) { toast('Only .xlsx and .xls supported', 'error'); return; }
  showLoading('Parsing Excel file...');
  const fd = new FormData(); fd.append('file', file);
  try {
    const res  = await fetch('/api/load', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.error) { toast(data.error, 'error'); hideLoading(); return; }
    sheetsData      = data.sheets;
    sheetNames      = data.sheet_names;
    currentFilename = data.filename;
    fileId          = data.file_id;
    titleInput.value = data.filename.replace(/\.[^.]+$/, '');
    document.title  = `${titleInput.value} — Krutidev Editor`;
    editedCells = new Set(); unicodeCache = {}; isModified = false;
    mergeCells = {}; cellAligns = {}; cellFontSizes = {};
    cellBold = {}; cellItalic = {}; cellUnderline = {}; cellBgColors = {};
    _lastSel = null;
    renderSheetTabs(); switchSheet(sheetNames[0]); showGrid(); enableButtons();
    toast(`Loaded: ${currentFilename}`, 'success');
  } catch (e) { toast('Network error', 'error'); }
  finally { hideLoading(); fileInput.value = ''; }
}

// ── Load shared doc ────────────────────────────────────
async function loadSharedDoc(id, pwd) {
  showLoading('Loading shared document...');
  try {
    const url = `/api/doc/${id}` + (pwd ? `?pwd=${encodeURIComponent(pwd)}` : '');
    const res  = await fetch(url);
    const data = await res.json();

    if (res.status === 403 && data.protected) {
      hideLoading();
      document.getElementById('pwdModal').style.display = 'flex';
      return;
    }
    if (data.error) { toast(data.error, 'error'); hideLoading(); return; }

    // Normalise sheets — DB may return {} or [] if doc was saved empty
    const rawSheets = data.sheets;
    const sheets = (rawSheets && !Array.isArray(rawSheets) && typeof rawSheets === 'object')
      ? rawSheets : {};
    const names  = Array.isArray(data.sheet_names) && data.sheet_names.length
      ? data.sheet_names : Object.keys(sheets);

    // If doc has no sheets at all, create a blank one so the editor isn't empty
    if (!names.length || !Object.keys(sheets).length) {
      const blankName = 'Sheet1';
      const ROWS = 50, COLS = 26;
      sheets[blankName] = {
        rows: Array.from({ length: ROWS }, () => Array(COLS).fill('')),
        col_widths: {}, row_count: ROWS, col_count: COLS
      };
      names.push(blankName);
    }

    sheetsData      = sheets;
    sheetNames      = names;
    currentFilename = (data.title || 'Untitled') + '.xlsx';
    docId           = id;          // ensure docId is set for autosave
    shareAccess     = data.access || 'edit';
    titleInput.value = data.title || 'Untitled';
    document.title  = `${titleInput.value} — Krutidev Editor`;
    editedCells = new Set(); unicodeCache = {}; isModified = false;
    mergeCells = {}; cellAligns = {}; cellFontSizes = {};
    cellBold = {}; cellItalic = {}; cellUnderline = {}; cellBgColors = {};
    _lastSel = null; fileId = null;

    // Reflect access in CFG so initHot uses correct readOnly state
    CFG.access = data.access || 'edit';

    renderSheetTabs(); switchSheet(sheetNames[0]); showGrid(); enableButtons();

    // Disable write buttons if view-only
    if (data.access === 'view') {
      ['btnSave','btnConvert','ribAddRow','ribDelRow','ribAddCol','ribMerge',
       'ribAlignLeft','ribAlignCenter','ribAlignRight','ribFontSize',
       'ribBold','ribItalic','ribUnderline','ribBgColor','ribBgColorClear'].forEach(bid => {
        const el = document.getElementById(bid);
        if (el) { el.disabled = true; el.title = 'View-only mode'; }
      });
    }

    toast('Document loaded', 'success');
  } catch (e) {
    console.error('loadSharedDoc error:', e);
    toast('Failed to load document', 'error');
  }
  finally { hideLoading(); }
}

function submitPassword() {
  const pwd = document.getElementById('pwdInput').value;
  document.getElementById('pwdModal').style.display = 'none';
  loadSharedDoc(docId, pwd);
}

// ── Sheet tabs ─────────────────────────────────────────
function renderSheetTabs() {
  sheetTabs.innerHTML = '';
  sheetNames.forEach(name => {
    const tab = document.createElement('button');
    tab.className = 'sheet-tab' + (name === activeSheet ? ' active' : '');
    tab.textContent = name;
    tab.onclick = () => switchSheet(name);
    sheetTabs.appendChild(tab);
  });
}

function switchSheet(name) {
  if (hot && activeSheet) sheetsData[activeSheet].rows = hot.getData();
  activeSheet = name;
  renderSheetTabs();
  const sheet = sheetsData[name];
  const cols  = sheet.col_count || (sheet.rows[0] ? sheet.rows[0].length : 26);
  if (hot) { hot.destroy(); hot = null; }

  // Use merge_list from file (server-parsed) if mergeCells not yet populated for this sheet
  if (!mergeCells[name] && sheet.merge_list && sheet.merge_list.length) {
    mergeCells[name] = sheet.merge_list;
  }
  if (!mergeCells[name]) mergeCells[name] = [];

  initHot(
    sheet.rows, cols,
    sheet.col_widths  || {},
    mergeCells[name],
    sheet.row_heights || {},
    sheet.krutidev_cells || []
  );
  updateStatus();
}

// ── Handsontable ───────────────────────────────────────
// krutidevCellSet: Set of "row,col" strings for cells that use a Krutidev-like font
let _krutidevCellSet = new Set();

function initHot(data, numCols, colWidths, merges, rowHeights, krutidevCells) {
  const container = document.getElementById('hot');
  const colWidthArr = Array.from({ length: numCols }, (_, i) => colWidths[colIndexToLetter(i)] || 100);
  const isReadOnly  = CFG.access === 'view';

  // Build per-row height array (Handsontable uses 0-based row index)
  // rowHeights from server is { "1": px, "2": px, ... } (1-based)
  const rowHeightArr = [];
  const numRows = data.length;
  for (let r = 0; r < numRows; r++) {
    const h = rowHeights[String(r + 1)];
    rowHeightArr.push(h ? Math.max(h, 18) : 24);
  }

  // Build krutidev cell set for this sheet
  _krutidevCellSet = new Set(krutidevCells || []);

  hot = new Handsontable(container, {
    data, rowHeaders: true, colHeaders: true,
    contextMenu: !isReadOnly ? {
      items: {
        'row_above': { name: 'Insert row above' },
        'row_below': { name: 'Insert row below' },
        'col_left':  { name: 'Insert column left' },
        'col_right': { name: 'Insert column right' },
        'hsep1': '---------',
        'remove_row':    { name: 'Remove row' },
        'remove_col':    { name: 'Remove column' },
        'hsep2': '---------',
        'mergeCells':    { name: 'Merge cells' },
        'hsep3': '---------',
        'copy':  { name: 'Copy' },
        'cut':   { name: 'Cut' },
      }
    } : false,
    manualColumnResize: true, manualRowResize: true,
    copyPaste: true, undo: !isReadOnly,
    colWidths: colWidthArr,
    rowHeights: rowHeightArr,
    stretchH: 'none', wordWrap: true,
    readOnly: isReadOnly,
    mergeCells: merges && merges.length ? merges : [],
    licenseKey: 'non-commercial-and-evaluation',
    cells(row, col) { return { renderer: cellRenderer }; },

    afterChange(changes, source) {
      if (!changes || source === 'loadData' || source === 'remote') return;
      changes.forEach(([row, col, oldVal, newVal]) => {
        if (oldVal === newVal) return;
        editedCells.add(`${row},${col}`);
        delete unicodeCache[`${activeSheet}:${row},${col}`];
        isModified = true;
        if (docId) {
          socket.emit('cell_change', { doc_id: docId, sheet: activeSheet, row, col, value: newVal });
        }
      });
      updateStatus(); updateStatusEdited(); scheduleAutoSave();
    },

    afterMergeCells(cellRange, mergeParent) {
      if (!mergeCells[activeSheet]) mergeCells[activeSheet] = [];
      mergeCells[activeSheet].push({
        row: mergeParent.row, col: mergeParent.col,
        rowspan: mergeParent.rowspan, colspan: mergeParent.colspan
      });
      isModified = true; updateStatusEdited();
      updateMergeBtn();
    },

    afterUnmergeCells(cellRange) {
      if (mergeCells[activeSheet]) {
        const sel = hot.getSelected();
        if (sel) {
          const [r1, c1] = sel[0];
          mergeCells[activeSheet] = mergeCells[activeSheet].filter(
            m => !(m.row === r1 && m.col === c1)
          );
        }
      }
      isModified = true; updateStatusEdited();
      updateMergeBtn();
    },

    afterSelectionEnd(r1, c1, r2, c2) {
      _lastSel = [r1, c1, r2, c2];
      const cl = colIndexToLetter(c1);
      cellRef.textContent = `${cl}${r1 + 1}`;
      formulaInput.value = String(hot.getDataAtCell(r1, c1) ?? '');
      statusCells.textContent = `${cl}${r1 + 1}`;
      if (docId) socket.emit('cursor_move', { doc_id: docId, row: r1, col: c1 });
      updateMergeBtn();
      updateAlignBtns(r1, c1);
      updateFontSizeInput(r1, c1);
      updateFormatBtns(r1, c1);
    },

    afterDeselect() {
      // Don't clear _lastSel — keep it so ribbon buttons still work
      cellRef.textContent = ''; formulaInput.value = ''; statusCells.textContent = 'Ready';
      updateMergeBtn();
      updateAlignBtns(-1, -1);
    },
  });

  formulaInput.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
      const sel = hot.getSelected();
      if (sel?.[0]) hot.setDataAtCell(sel[0][0], sel[0][1], formulaInput.value);
    }
  });
  applyFontClass();
}

// ── Cell renderer ──────────────────────────────────────
function cellRenderer(instance, td, row, col, prop, value, cellProperties) {
  Handsontable.renderers.TextRenderer.apply(this, arguments);

  const key = `${row},${col}`;

  // Edited highlight
  if (editedCells.has(key)) td.classList.add('cell-edited');

  // Text alignment
  const align = (cellAligns[activeSheet] || {})[key];
  if (align) td.style.textAlign = align;

  // Background color
  const bg = (cellBgColors[activeSheet] || {})[key];
  if (bg) td.style.backgroundColor = bg;

  // Bold / Italic / Underline
  const isBold      = (cellBold[activeSheet]      || {})[key];
  const isItalic    = (cellItalic[activeSheet]     || {})[key];
  const isUnderline = (cellUnderline[activeSheet]  || {})[key];
  td.style.fontWeight     = isBold      ? 'bold'      : '';
  td.style.fontStyle      = isItalic    ? 'italic'    : '';
  td.style.textDecoration = isUnderline ? 'underline' : '';

  // Custom font size — must override CSS class rules
  const customSize = (cellFontSizes[activeSheet] || {})[key];
  const isKrutidevCell = _krutidevCellSet.has(key);

  if (fontMode === 'unicode' && value && typeof value === 'string') {
    const k = `${activeSheet}:${row},${col}`;
    if (!unicodeCache[k]) unicodeCache[k] = clientKrutidevToUnicode(value);
    td.textContent = unicodeCache[k];
    td.style.fontFamily = "'Noto Sans Devanagari','Mangal','Arial Unicode MS',sans-serif";
    td.style.fontSize   = (customSize || 13) + 'px';
  } else if (isKrutidevCell) {
    td.style.fontFamily = "'KrutiDev','Kruti Dev 010',serif";
    td.style.fontSize   = (customSize || 14) + 'px';
  } else {
    td.style.fontFamily = '';
    td.style.fontSize   = customSize ? customSize + 'px' : '';
  }

  // Clip overflow — prevent merged cell text bleeding
  td.style.overflow   = 'hidden';
  td.style.whiteSpace = 'pre-wrap';
}

// ── Font mode ──────────────────────────────────────────
function setFont(mode) {
  fontMode = mode;
  document.getElementById('btnKrutidev').classList.toggle('active', mode === 'krutidev');
  document.getElementById('btnUnicode').classList.toggle('active', mode === 'unicode');
  applyFontClass();
  if (hot) hot.render();
  statusFont.textContent = `Font: ${mode === 'krutidev' ? 'Krutidev' : 'Unicode Preview'}`;
}
function applyFontClass() {
  const a = document.getElementById('gridContainer');
  a.classList.toggle('font-krutidev', fontMode === 'krutidev');
  a.classList.toggle('font-unicode',  fontMode === 'unicode');
}

// ── Add / Delete rows & cols ───────────────────────────
function addRow() {
  if (!hot) return;
  hot.alter('insert_row_below', hot.countRows() - 1);
  isModified = true; updateStatus(); toast('Row added', 'success', 1500);
}
function deleteSelectedRows() {
  if (!hot) return;
  const sel = hot.getSelected();
  if (!sel) { toast('Select rows first', 'warning'); return; }
  const rows = new Set();
  sel.forEach(([r1,,r2]) => { for (let r = Math.min(r1,r2); r <= Math.max(r1,r2); r++) rows.add(r); });
  [...rows].sort((a,b) => b-a).forEach(r => hot.alter('remove_row', r));
  isModified = true; updateStatus(); toast(`${rows.size} row(s) deleted`, 'success', 1500);
}
function addCol() {
  if (!hot) return;
  hot.alter('insert_col_end');
  isModified = true; updateStatus(); toast('Column added', 'success', 1500);
}

// ── Merge cells ────────────────────────────────────────
function toggleMergeCells() {
  if (!hot) return;
  // Use current selection or fall back to last known selection
  const sel = hot.getSelected() || (_lastSel ? [[..._lastSel]] : null);
  if (!sel || sel.length === 0) { toast('Pehle cells select karein', 'warning'); return; }
  const [r1, c1, r2, c2] = sel[0];
  const rMin = Math.min(r1, r2), rMax = Math.max(r1, r2);
  const cMin = Math.min(c1, c2), cMax = Math.max(c1, c2);
  if (rMin === rMax && cMin === cMax) { toast('Merge ke liye multiple cells select karein', 'warning'); return; }

  // Re-focus grid so Handsontable plugin works correctly
  hot.selectCell(rMin, cMin, rMax, cMax);

  // Check if already merged
  const existing = (mergeCells[activeSheet] || []).find(
    m => m.row === rMin && m.col === cMin
  );
  if (existing) {
    hot.getPlugin('mergeCells').unmerge(rMin, cMin, rMax, cMax);
    toast('Cells unmerged', 'success', 1500);
  } else {
    hot.getPlugin('mergeCells').merge(rMin, cMin, rMax, cMax);
    toast(`${rMax - rMin + 1}×${cMax - cMin + 1} cells merged`, 'success', 1500);
  }
  isModified = true; updateStatusEdited(); updateMergeBtn();
}

function updateMergeBtn() {
  const btn = document.getElementById('ribMerge');
  if (!btn || !hot) return;
  const sel = hot.getSelected();
  if (!sel) { btn.classList.remove('active'); return; }
  const [r1, c1, r2, c2] = sel[0];
  const rMin = Math.min(r1, r2), cMin = Math.min(c1, c2);
  const isMerged = (mergeCells[activeSheet] || []).some(
    m => m.row === rMin && m.col === cMin && (m.rowspan > 1 || m.colspan > 1)
  );
  btn.classList.toggle('active', isMerged);
  btn.textContent = isMerged ? '⊞ Unmerge' : '⊞ Merge';
}

// ── Text alignment ─────────────────────────────────────
function setAlign(align) {
  if (!hot) return;
  // Use current or last known selection
  const sel = hot.getSelected() || (_lastSel ? [[..._lastSel]] : null);
  if (!sel) { toast('Pehle cells select karein', 'warning'); return; }
  if (!cellAligns[activeSheet]) cellAligns[activeSheet] = {};
  sel.forEach(([r1, c1, r2, c2]) => {
    const rMin = Math.min(r1, r2), rMax = Math.max(r1, r2);
    const cMin = Math.min(c1, c2), cMax = Math.max(c1, c2);
    for (let r = rMin; r <= rMax; r++) {
      for (let c = cMin; c <= cMax; c++) {
        cellAligns[activeSheet][`${r},${c}`] = align;
      }
    }
  });
  hot.render();
  isModified = true; updateStatusEdited();
  const ref = sel[0];
  updateAlignBtns(ref[0], ref[1]);
}

function updateAlignBtns(row, col) {
  const align = row >= 0 ? ((cellAligns[activeSheet] || {})[`${row},${col}`] || '') : '';
  ['left','center','right'].forEach(a => {
    const btn = document.getElementById(`ribAlign${a.charAt(0).toUpperCase() + a.slice(1)}`);
    if (btn) btn.classList.toggle('active', align === a);
  });
}

// ── Font size ──────────────────────────────────────────
function setFontSize(size) {
  if (!hot) return;
  const sz = parseInt(size, 10);
  if (!sz || sz < 6 || sz > 96) return;
  const sel = hot.getSelected() || (_lastSel ? [[..._lastSel]] : null);
  if (!sel) return;
  if (!cellFontSizes[activeSheet]) cellFontSizes[activeSheet] = {};
  sel.forEach(([r1, c1, r2, c2]) => {
    const rMin = Math.min(r1, r2), rMax = Math.max(r1, r2);
    const cMin = Math.min(c1, c2), cMax = Math.max(c1, c2);
    for (let r = rMin; r <= rMax; r++) {
      for (let c = cMin; c <= cMax; c++) {
        cellFontSizes[activeSheet][`${r},${c}`] = sz;
      }
    }
  });
  hot.render();
  isModified = true; updateStatusEdited();
}

function updateFontSizeInput(row, col) {
  const inp = document.getElementById('ribFontSize');
  if (!inp) return;
  const sz = (cellFontSizes[activeSheet] || {})[`${row},${col}`];
  inp.value = sz || '';
  inp.placeholder = fontMode === 'krutidev' ? '14' : '13';
}

// ── Bold / Italic / Underline ──────────────────────────
function _applyStyleToSelection(stateObj, btnId) {
  if (!hot) return;
  const sel = hot.getSelected() || (_lastSel ? [[..._lastSel]] : null);
  if (!sel) return;
  if (!stateObj[activeSheet]) stateObj[activeSheet] = {};
  // Toggle: if ALL selected cells already have it, remove; else apply
  let allOn = true;
  sel.forEach(([r1,c1,r2,c2]) => {
    for (let r = Math.min(r1,r2); r <= Math.max(r1,r2); r++)
      for (let c = Math.min(c1,c2); c <= Math.max(c1,c2); c++)
        if (!stateObj[activeSheet][`${r},${c}`]) allOn = false;
  });
  sel.forEach(([r1,c1,r2,c2]) => {
    for (let r = Math.min(r1,r2); r <= Math.max(r1,r2); r++)
      for (let c = Math.min(c1,c2); c <= Math.max(c1,c2); c++)
        stateObj[activeSheet][`${r},${c}`] = !allOn;
  });
  hot.render();
  isModified = true; updateStatusEdited();
  const btn = document.getElementById(btnId);
  if (btn) btn.classList.toggle('active', !allOn);
}

function toggleBold()      { _applyStyleToSelection(cellBold,      'ribBold'); }
function toggleItalic()    { _applyStyleToSelection(cellItalic,    'ribItalic'); }
function toggleUnderline() { _applyStyleToSelection(cellUnderline, 'ribUnderline'); }

function updateFormatBtns(row, col) {
  const key = `${row},${col}`;
  const b = document.getElementById('ribBold');
  const i = document.getElementById('ribItalic');
  const u = document.getElementById('ribUnderline');
  if (b) b.classList.toggle('active', !!(cellBold[activeSheet]      || {})[key]);
  if (i) i.classList.toggle('active', !!(cellItalic[activeSheet]    || {})[key]);
  if (u) u.classList.toggle('active', !!(cellUnderline[activeSheet] || {})[key]);
}

// ── Cell background color ──────────────────────────────
function setCellBgColor(color) {
  if (!hot) return;
  const sel = hot.getSelected() || (_lastSel ? [[..._lastSel]] : null);
  if (!sel) return;
  if (!cellBgColors[activeSheet]) cellBgColors[activeSheet] = {};
  sel.forEach(([r1,c1,r2,c2]) => {
    for (let r = Math.min(r1,r2); r <= Math.max(r1,r2); r++)
      for (let c = Math.min(c1,c2); c <= Math.max(c1,c2); c++)
        cellBgColors[activeSheet][`${r},${c}`] = color === 'none' ? '' : color;
  });
  hot.render();
  isModified = true; updateStatusEdited();
}
function openNewSheetModal() {
  document.getElementById('newSheetModal').style.display = 'flex';
}
function closeNewSheetModal(e) {
  if (!e || e.target === document.getElementById('newSheetModal'))
    document.getElementById('newSheetModal').style.display = 'none';
}

function createNewSheet(fontType) {
  document.getElementById('newSheetModal').style.display = 'none';

  // Build a blank 50×26 grid
  const ROWS = 50, COLS = 26;
  const rows = Array.from({ length: ROWS }, () => Array(COLS).fill(''));
  const colWidths = {};
  for (let i = 0; i < COLS; i++) colWidths[colIndexToLetter(i)] = 100;

  // Pick a unique sheet name
  let baseName = 'Sheet', idx = sheetNames.length + 1;
  while (sheetNames.includes(baseName + idx)) idx++;
  const name = baseName + idx;

  sheetsData[name] = { rows, col_widths: colWidths, row_count: ROWS, col_count: COLS };
  sheetNames.push(name);
  mergeCells[name]    = [];
  cellAligns[name]    = {};
  cellFontSizes[name] = {};
  cellBold[name]      = {};
  cellItalic[name]    = {};
  cellUnderline[name] = {};
  cellBgColors[name]  = {};

  // Set font mode based on choice
  if (fontType === 'krutidev') setFont('krutidev');
  else if (fontType === 'unicode') setFont('unicode');
  else setFont('krutidev'); // blank/english default

  showGrid(); enableButtons();
  switchSheet(name);
  isModified = true; updateStatusEdited();
  toast(`New sheet "${name}" created`, 'success', 2000);
}

// ── Save / Download ────────────────────────────────────
async function saveFile(convertMode) {
  if (!hot || !activeSheet) return;
  sheetsData[activeSheet].rows = hot.getData();
  showLoading(convertMode ? 'Converting...' : 'Preparing download...');
  const payload = {
    filename: currentFilename,
    convert: convertMode,
    file_id: fileId,
    sheets: {}
  };
  sheetNames.forEach(n => {
    payload.sheets[n] = {
      rows: sheetsData[n].rows,
      merge_cells: mergeCells[n] || []
    };
  });
  try {
    const res = await fetch('/api/save', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!res.ok) { const e = await res.json(); toast(e.error || 'Save failed', 'error'); return; }
    const blob = await res.blob();
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    const name = currentFilename.replace(/\.[^.]+$/, '');
    a.href = url; a.download = convertMode ? `${name}_unicode.xlsx` : `${name}_edited.xlsx`;
    a.click(); URL.revokeObjectURL(url);
    isModified = false; updateStatusEdited();
    toast(convertMode ? 'Converted & downloaded!' : 'Downloaded!', 'success');
  } catch (e) { toast('Download failed', 'error'); }
  finally { hideLoading(); }
}

// ── Auto-save to DB ────────────────────────────────────
function scheduleAutoSave() {
  clearTimeout(autoSaveTimer);
  autosaveIndicator.textContent = 'Saving...';
  autoSaveTimer = setTimeout(autoSave, 2500);
}

async function autoSave() {
  if (!docId || !isModified) return;
  if (hot && activeSheet) sheetsData[activeSheet].rows = hot.getData();
  try {
    const res = await fetch('/api/share', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        doc_id: docId, title: titleInput.value,
        sheets: sheetsData, sheet_names: sheetNames,
        access: shareAccess, file_id: fileId
      })
    });
    if (res.ok) {
      isModified = false; lastSaved = new Date();
      autosaveIndicator.textContent = `Saved ${lastSaved.toLocaleTimeString()}`;
      updateStatusEdited();
    }
  } catch (e) { autosaveIndicator.textContent = 'Save failed'; }
}

// ── Title change ───────────────────────────────────────
function onTitleChange() {
  const title = titleInput.value.trim() || 'Untitled';
  document.title = `${title} — Krutidev Editor`;
  currentFilename = title + '.xlsx';
  if (docId) socket.emit('title_change', { doc_id: docId, title });
  scheduleAutoSave();
}

// ── Share modal ────────────────────────────────────────
function openShareModal() {
  document.getElementById('shareModal').style.display = 'flex';
  document.getElementById('shareLinkBox').style.display = 'none';
}
function closeShareModal(e) {
  if (!e || e.target === document.getElementById('shareModal'))
    document.getElementById('shareModal').style.display = 'none';
}
function setAccess(mode) {
  shareAccess = mode;
  document.getElementById('accEdit').classList.toggle('active', mode === 'edit');
  document.getElementById('accView').classList.toggle('active', mode === 'view');
}

async function generateLink() {
  if (hot && activeSheet) sheetsData[activeSheet].rows = hot.getData();
  const pwd = document.getElementById('sharePwd').value;
  const btn = document.getElementById('btnGenLink');
  btn.textContent = 'Generating...'; btn.disabled = true;
  try {
    const res = await fetch('/api/share', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        doc_id: docId, title: titleInput.value,
        sheets: sheetsData, sheet_names: sheetNames,
        access: shareAccess, password: pwd || null,
        file_id: fileId
      })
    });
    const data = await res.json();
    if (data.error) { toast(data.error, 'error'); return; }
    docId = data.doc_id;
    // Join the socket room now
    socket.emit('join', { doc_id: docId, name: myName });
    document.getElementById('shareLink').value = data.url;
    document.getElementById('shareLinkBox').style.display = 'block';
    // Update browser URL without reload
    history.replaceState({}, '', `/sheet/${docId}`);
    toast('Share link generated!', 'success');
  } catch (e) { toast('Failed to generate link', 'error'); }
  finally { btn.textContent = '🔗 Generate Link'; btn.disabled = false; }
}

function copyLink() {
  const link = document.getElementById('shareLink').value;
  navigator.clipboard.writeText(link).then(() => toast('Link copied!', 'success'));
}

// ── Search & Replace ───────────────────────────────────
let searchMatches = [], searchIdx = 0;
function toggleSearch() {
  const v = searchBar.style.display !== 'none';
  searchBar.style.display = v ? 'none' : 'flex';
  if (!v) searchInput.focus();
}
function doSearch() {
  if (!hot) return;
  const q = searchInput.value; if (!q) return;
  searchMatches = [];
  hot.getData().forEach((row, r) => row.forEach((cell, c) => {
    if (cell && String(cell).includes(q)) searchMatches.push([r, c]);
  }));
  if (!searchMatches.length) { toast('No matches', 'warning'); return; }
  searchIdx = 0; jumpToMatch();
  toast(`${searchMatches.length} match(es)`, 'success', 2000);
}
function jumpToMatch() {
  if (!searchMatches.length) return;
  const [r, c] = searchMatches[searchIdx % searchMatches.length];
  hot.selectCell(r, c); hot.scrollViewportTo(r, c);
}
function doReplace() {
  if (!hot || !searchMatches.length) return;
  const [r, c] = searchMatches[searchIdx % searchMatches.length];
  const old = hot.getDataAtCell(r, c) || '';
  hot.setDataAtCell(r, c, old.replace(searchInput.value, replaceInput.value));
  searchIdx++;
  if (searchIdx < searchMatches.length) jumpToMatch();
  else toast('Done', 'success', 1500);
}
function doReplaceAll() {
  if (!hot) return;
  const q = searchInput.value, r = replaceInput.value; if (!q) return;
  const changes = [];
  hot.getData().forEach((row, ri) => row.forEach((cell, ci) => {
    if (cell && String(cell).includes(q)) changes.push([ri, ci, String(cell).replaceAll(q, r)]);
  }));
  if (!changes.length) { toast('No matches', 'warning'); return; }
  changes.forEach(([ri, ci, v]) => hot.setDataAtCell(ri, ci, v));
  changes.forEach(([ri, ci]) => editedCells.add(`${ri},${ci}`));
  toast(`Replaced ${changes.length} cell(s)`, 'success');
}

// ── Keyboard shortcuts ─────────────────────────────────
document.addEventListener('keydown', e => {
  if ((e.ctrlKey || e.metaKey) && e.key === 'f') { e.preventDefault(); toggleSearch(); }
  if ((e.ctrlKey || e.metaKey) && e.key === 's') { e.preventDefault(); if (hot) saveFile(false); }
  if ((e.ctrlKey || e.metaKey) && e.key === 'b') { e.preventDefault(); if (hot) toggleBold(); }
  if ((e.ctrlKey || e.metaKey) && e.key === 'i') { e.preventDefault(); if (hot) toggleItalic(); }
  if ((e.ctrlKey || e.metaKey) && e.key === 'u') { e.preventDefault(); if (hot) toggleUnderline(); }
  if (e.key === 'Escape') searchBar.style.display = 'none';
});

// ── Users bar ──────────────────────────────────────────
function renderUsersBar(users) {
  const bar = document.getElementById('usersBar');
  bar.innerHTML = users.slice(0, 6).map(u =>
    `<div class="user-avatar" style="background:${u.color}" title="${u.name}">${u.name[0].toUpperCase()}</div>`
  ).join('');
}

// ── Status ─────────────────────────────────────────────
function updateStatus() {
  if (!hot) return;
  statusRows.textContent = `${hot.countRows()} rows`;
  statusCols.textContent = `${hot.countCols()} cols`;
}
function updateStatusEdited() {
  statusEdited.textContent = isModified ? '● Unsaved' : '';
}

// ── UI helpers ─────────────────────────────────────────
function showGrid() {
  emptyState.style.display    = 'none';
  gridContainer.style.display = 'block';
  sheetTabsBar.style.display  = 'flex';
  formulaBar.style.display    = 'flex';
  statusBar.style.display     = 'flex';
  document.getElementById('ribbon').style.display = 'flex';
}
function showLoading(msg) { loadingText.textContent = msg || 'Loading...'; loadingOverlay.style.display = 'flex'; }
function hideLoading()    { loadingOverlay.style.display = 'none'; }
function enableButtons()  {
  ['btnSave','btnConvert','btnShare'].forEach(id => {
    const el = document.getElementById(id);
    if (el && CFG.access !== 'view') el.disabled = false;
  });
  // Always enable share
  const share = document.getElementById('btnShare');
  if (share) share.disabled = false;
  // Ribbon buttons
  ['ribAddRow','ribDelRow','ribAddCol','ribMerge',
   'ribAlignLeft','ribAlignCenter','ribAlignRight',
   'ribFontSize','ribBold','ribItalic','ribUnderline','ribBgColor','ribBgColorClear'].forEach(id => {
    const el = document.getElementById(id);
    if (el && CFG.access !== 'view') el.disabled = false;
  });
}
function colIndexToLetter(idx) {
  let s = ''; idx++;
  while (idx > 0) { idx--; s = String.fromCharCode(65 + (idx % 26)) + s; idx = Math.floor(idx / 26); }
  return s;
}

// ── Client-side Krutidev→Unicode preview ──────────────
const KD_SIMPLE = [
  ['vks','ओ'],['vkS','औ'],['vk','आ'],['v','अ'],['bZ','ई'],['b','इ'],
  ['m','उ'],['Å','ऊ'],[',s','ऐ'],[',','ए'],
  ['d','क'],['[k','ख'],['x','ग'],['?k','घ'],['p','च'],['N','छ'],
  ['t','ज'],['>','झ'],['V','ट'],['B','ठ'],['M','ड'],['<','ढ'],
  ['.k','ण'],['r','त'],['Fk','थ'],['n','द'],['/k','ध'],['u','न'],
  ['i','प'],['Q','फ'],['c','ब'],['Hk','भ'],['e','म'],[';','य'],
  ['j','र'],['y','ल'],['o','व'],["'k",'श'],['"k','ष'],['l','स'],['g','ह'],
  ['K','ज्ञ'],['=','त्र'],['J','श्र'],
  ['k','ा'],['h','ी'],['q','ु'],['w','ू'],['s','े'],['S','ै'],
  ['ks','ो'],['kS','ौ'],['a','ं'],['%',':'],['~','्'],
  ['A','।'],['-','.'],
];
function clientKrutidevToUnicode(text) {
  if (!text) return text;
  let out = '', i = 0;
  while (i < text.length) {
    if (text[i] === 'f' && i + 1 < text.length) {
      let j = i + 1;
      while (j < text.length && text[j] === 'a') j++;
      const rest = text.slice(j);
      let matched = false;
      for (const [kd, uni] of KD_SIMPLE) {
        if (rest.startsWith(kd) && uni[0] >= '\u0915' && uni[0] <= '\u0939') {
          out += uni + 'ि'; if (j > i + 1) out += 'ं';
          i = j + kd.length; matched = true; break;
        }
      }
      if (!matched) { out += text[i]; i++; }
      continue;
    }
    let matched = false;
    for (const [kd, uni] of KD_SIMPLE) {
      if (text.startsWith(kd, i)) { out += uni; i += kd.length; matched = true; break; }
    }
    if (!matched) { out += text[i]; i++; }
  }
  return out.replace(/(.?)Z/g, (_, ch) => ch ? 'र्' + ch : '');
}

// ── Init ───────────────────────────────────────────────
window.addEventListener('beforeunload', e => { if (isModified) { e.preventDefault(); e.returnValue = ''; } });

// Hide ribbon initially (shown when grid is active)
document.getElementById('ribbon').style.display = 'none';

// Check URL params for new sheet preset
const _urlParams = new URLSearchParams(window.location.search);
const _newPreset = _urlParams.get('new');
if (_newPreset) {
  // Auto-open new sheet modal or create directly
  setTimeout(() => openNewSheetModal(), 100);
} else if (docId) {
  loadSharedDoc(docId);
}
