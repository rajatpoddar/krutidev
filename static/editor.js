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
let shareAccess  = 'edit';
let autoSaveTimer = null;
let lastSaved    = null;

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
    titleInput.value = data.filename.replace(/\.[^.]+$/, '');
    document.title  = `${titleInput.value} — Krutidev Editor`;
    editedCells = new Set(); unicodeCache = {}; isModified = false;
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

    sheetsData      = data.sheets;
    sheetNames      = data.sheet_names;
    currentFilename = data.title + '.xlsx';
    titleInput.value = data.title;
    document.title  = `${data.title} — Krutidev Editor`;
    editedCells = new Set(); unicodeCache = {}; isModified = false;

    // Disable editing if view-only
    if (data.access === 'view') {
      ['btnAddRow','btnDelRow','btnSave','btnConvert'].forEach(id => {
        const el = document.getElementById(id);
        if (el) { el.disabled = true; el.title = 'View-only mode'; }
      });
    }

    renderSheetTabs(); switchSheet(sheetNames[0]); showGrid(); enableButtons();
    toast('Document loaded', 'success');
  } catch (e) { toast('Failed to load document', 'error'); }
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
  initHot(sheet.rows, cols, sheet.col_widths || {});
  updateStatus();
}

// ── Handsontable ───────────────────────────────────────
function initHot(data, numCols, colWidths) {
  const container = document.getElementById('hot');
  const colWidthArr = Array.from({ length: numCols }, (_, i) => colWidths[colIndexToLetter(i)] || 100);
  const isReadOnly  = CFG.access === 'view';

  hot = new Handsontable(container, {
    data, rowHeaders: true, colHeaders: true,
    contextMenu: !isReadOnly, manualColumnResize: true, manualRowResize: true,
    copyPaste: true, undo: !isReadOnly,
    colWidths: colWidthArr, rowHeights: 24,
    stretchH: 'none', wordWrap: false,
    readOnly: isReadOnly,
    licenseKey: 'non-commercial-and-evaluation',
    cells() { return { renderer: cellRenderer }; },

    afterChange(changes, source) {
      if (!changes || source === 'loadData' || source === 'remote') return;
      changes.forEach(([row, col, oldVal, newVal]) => {
        if (oldVal === newVal) return;
        editedCells.add(`${row},${col}`);
        delete unicodeCache[`${activeSheet}:${row},${col}`];
        isModified = true;
        // Emit to collaborators
        if (docId) {
          socket.emit('cell_change', { doc_id: docId, sheet: activeSheet, row, col, value: newVal });
        }
      });
      updateStatus(); updateStatusEdited(); scheduleAutoSave();
    },

    afterSelectionEnd(r1, c1) {
      const cl = colIndexToLetter(c1);
      cellRef.textContent = `${cl}${r1 + 1}`;
      formulaInput.value = String(hot.getDataAtCell(r1, c1) ?? '');
      statusCells.textContent = `${cl}${r1 + 1}`;
      if (docId) socket.emit('cursor_move', { doc_id: docId, row: r1, col: c1 });
    },

    afterDeselect() { cellRef.textContent = ''; formulaInput.value = ''; statusCells.textContent = 'Ready'; },
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
  if (editedCells.has(`${row},${col}`)) td.classList.add('cell-edited');
  if (fontMode === 'unicode' && value && typeof value === 'string') {
    const k = `${activeSheet}:${row},${col}`;
    if (!unicodeCache[k]) unicodeCache[k] = clientKrutidevToUnicode(value);
    td.textContent = unicodeCache[k];
  }
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

// ── Add / Delete rows ──────────────────────────────────
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

// ── Save / Download ────────────────────────────────────
async function saveFile(convertMode) {
  if (!hot || !activeSheet) return;
  sheetsData[activeSheet].rows = hot.getData();
  showLoading(convertMode ? 'Converting...' : 'Preparing download...');
  const payload = { filename: currentFilename, convert: convertMode, sheets: {} };
  sheetNames.forEach(n => { payload.sheets[n] = { rows: sheetsData[n].rows }; });
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
        sheets: sheetsData, sheet_names: sheetNames, access: shareAccess
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
        access: shareAccess, password: pwd || null
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
}
function showLoading(msg) { loadingText.textContent = msg || 'Loading...'; loadingOverlay.style.display = 'flex'; }
function hideLoading()    { loadingOverlay.style.display = 'none'; }
function enableButtons()  {
  ['btnAddRow','btnDelRow','btnSave','btnConvert','btnShare'].forEach(id => {
    const el = document.getElementById(id);
    if (el && CFG.access !== 'view') el.disabled = false;
  });
  const share = document.getElementById('btnShare');
  if (share) share.disabled = false; // share always enabled
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

if (docId) {
  loadSharedDoc(docId);
}
