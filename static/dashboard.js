'use strict';

// ── Toast ──────────────────────────────────────────────
function toast(msg, type = 'info', ms = 3000) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = `toast show ${type}`;
  clearTimeout(el._t);
  el._t = setTimeout(() => { el.className = 'toast'; }, ms);
}

// ── Theme ──────────────────────────────────────────────
function toggleTheme() {
  const html = document.documentElement;
  const dark = html.getAttribute('data-theme') === 'dark';
  html.setAttribute('data-theme', dark ? 'light' : 'dark');
  document.getElementById('themeLabel').textContent = dark ? 'Light Mode' : 'Dark Mode';
  localStorage.setItem('theme', dark ? 'light' : 'dark');
}
(function initTheme() {
  const t = localStorage.getItem('theme') || 'light';
  document.documentElement.setAttribute('data-theme', t);
  const lbl = document.getElementById('themeLabel');
  if (lbl) lbl.textContent = t === 'dark' ? 'Dark Mode' : 'Light Mode';
})();

// ── Sidebar toggle ─────────────────────────────────────
function toggleSidebar() {
  const sb = document.getElementById('sidebar');
  const mw = document.querySelector('.main-wrap');
  if (window.innerWidth <= 768) {
    sb.classList.toggle('open');
  } else {
    sb.classList.toggle('collapsed');
    mw.classList.toggle('full');
  }
}

// ── Section switching ──────────────────────────────────
function showSection(name) {
  document.querySelectorAll('section[id^="section-"]').forEach(s => s.style.display = 'none');
  const el = document.getElementById(`section-${name}`);
  if (el) el.style.display = 'block';

  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const navMap = { 'dashboard': 0, 'my-sheets': 3 };
  const navItems = document.querySelectorAll('.nav-item');
  if (name === 'dashboard') navItems[0].classList.add('active');
  if (name === 'my-sheets') navItems[3].classList.add('active');

  const titles = { 'dashboard': 'Dashboard', 'my-sheets': 'My Sheets' };
  document.getElementById('pageTitle').textContent = titles[name] || 'Dashboard';

  if (name === 'my-sheets') loadSheetsTable();
}

// ── Load docs from API ─────────────────────────────────
let allDocs = [];

async function loadDocs() {
  try {
    const res = await fetch('/api/docs');
    const data = await res.json();
    allDocs = data.docs || [];
    renderStats();
    renderRecentSheets();
    document.getElementById('sheetCount').textContent = allDocs.length;
  } catch (e) {
    console.error(e);
  }
}

function renderStats() {
  document.getElementById('statTotal').textContent = allDocs.length;
  const today = new Date().toISOString().slice(0, 10);
  const recent = allDocs.filter(d => d.updated_at && d.updated_at.startsWith(today)).length;
  document.getElementById('statRecent').textContent = recent;
  document.getElementById('statShared').textContent = allDocs.length;
}

function renderRecentSheets() {
  const container = document.getElementById('recentSheets');
  const recent = allDocs.slice(0, 6);
  if (!recent.length) {
    container.innerHTML = '<p style="color:var(--text-muted);font-size:14px;padding:20px 0;">No sheets yet. Upload an Excel file to get started.</p>';
    return;
  }
  container.innerHTML = '';
  container.className = 'sheets-grid';
  recent.forEach(doc => {
    const card = document.createElement('div');
    card.className = 'sheet-card';
    card.onclick = () => window.location.href = `/sheet/${doc.id}`;
    card.innerHTML = `
      <div class="sheet-card-icon">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/>
        </svg>
      </div>
      <div class="sheet-card-title" title="${doc.title}">${doc.title}</div>
      <div class="sheet-card-meta">${formatDate(doc.updated_at)}</div>
      <span class="sheet-card-badge ${doc.access === 'edit' ? 'badge-edit' : 'badge-view'}">
        ${doc.access === 'edit' ? '✏️ Edit' : '👁 View'}
      </span>`;
    container.appendChild(card);
  });
}

function loadSheetsTable() {
  const tbody = document.getElementById('sheetsTableBody');
  if (!allDocs.length) {
    tbody.innerHTML = '<tr><td colspan="5" class="empty-row">No sheets found. Upload an Excel file to get started.</td></tr>';
    return;
  }
  tbody.innerHTML = allDocs.map(doc => `
    <tr>
      <td><input type="checkbox" class="row-check" data-id="${doc.id}" onchange="onRowCheck()"/></td>
      <td><a href="/sheet/${doc.id}" style="color:var(--primary);font-weight:600;text-decoration:none;">${doc.title}</a></td>
      <td><span class="sheet-card-badge ${doc.access === 'edit' ? 'badge-edit' : 'badge-view'}">${doc.access === 'edit' ? '✏️ Edit' : '👁 View'}</span></td>
      <td style="color:var(--text-muted)">${formatDate(doc.updated_at)}</td>
      <td>
        <div class="action-btns">
          <button class="action-btn" onclick="window.location.href='/sheet/${doc.id}'">Open</button>
          <button class="action-btn" onclick="copyShareLink('${doc.id}')">Copy Link</button>
          <button class="action-btn" onclick="downloadDoc('${doc.id}','${doc.title}')">↓ Excel</button>
          <button class="action-btn del" onclick="deleteDoc('${doc.id}')">Delete</button>
        </div>
      </td>
    </tr>`).join('');
}

function onRowCheck() {
  const any = document.querySelectorAll('.row-check:checked').length > 0;
  document.getElementById('btnDeleteSelected').style.display = any ? 'inline-flex' : 'none';
}

function toggleSelectAll(cb) {
  document.querySelectorAll('.row-check').forEach(c => c.checked = cb.checked);
  onRowCheck();
}

async function deleteSelected() {
  const ids = [...document.querySelectorAll('.row-check:checked')].map(c => c.dataset.id);
  if (!ids.length) return;
  if (!confirm(`Delete ${ids.length} sheet(s)?`)) return;
  await Promise.all(ids.map(id => fetch(`/api/doc/${id}`, { method: 'DELETE' })));
  toast(`${ids.length} sheet(s) deleted`, 'success');
  await loadDocs();
  loadSheetsTable();
}

async function deleteDoc(id) {
  if (!confirm('Delete this sheet?')) return;
  await fetch(`/api/doc/${id}`, { method: 'DELETE' });
  toast('Sheet deleted', 'success');
  await loadDocs();
  loadSheetsTable();
}

function copyShareLink(id) {
  const url = `${location.origin}/sheet/${id}`;
  navigator.clipboard.writeText(url).then(() => toast('Link copied!', 'success'));
}

function downloadDoc(id, title) {
  window.location.href = `/api/doc/${id}/download`;
}

// ── Upload modal ───────────────────────────────────────
function triggerUpload() {
  document.getElementById('uploadModal').style.display = 'flex';
}

function closeUploadModal(e) {
  if (!e || e.target === document.getElementById('uploadModal')) {
    document.getElementById('uploadModal').style.display = 'none';
  }
}

const dashDropZone = document.getElementById('dashDropZone');
const dashFileInput = document.getElementById('dashFileInput');

dashDropZone.addEventListener('click', () => dashFileInput.click());
dashDropZone.addEventListener('dragover', e => { e.preventDefault(); dashDropZone.classList.add('drag-over'); });
dashDropZone.addEventListener('dragleave', () => dashDropZone.classList.remove('drag-over'));
dashDropZone.addEventListener('drop', e => {
  e.preventDefault();
  dashDropZone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) openInEditor(f);
});
dashFileInput.addEventListener('change', () => {
  if (dashFileInput.files[0]) openInEditor(dashFileInput.files[0]);
});

async function openInEditor(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  if (!['xlsx', 'xls'].includes(ext)) { toast('Only .xlsx and .xls supported', 'error'); return; }
  closeUploadModal();
  toast('Loading file...', 'info', 10000);

  const fd = new FormData();
  fd.append('file', file);
  try {
    const res = await fetch('/api/load', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.error) { toast(data.error, 'error'); return; }

    // Share it to get a doc_id, then redirect to editor
    const shareRes = await fetch('/api/share', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        title: data.filename.replace(/\.[^.]+$/, ''),
        sheets: data.sheets,
        sheet_names: data.sheet_names,
        access: 'edit'
      })
    });
    const shareData = await shareRes.json();
    window.location.href = `/sheet/${shareData.doc_id}`;
  } catch (e) {
    toast('Failed to open file', 'error');
  }
}

// ── Helpers ────────────────────────────────────────────
function formatDate(dt) {
  if (!dt) return '—';
  const d = new Date(dt.replace(' ', 'T') + 'Z');
  return d.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
}

// ── Init ───────────────────────────────────────────────
loadDocs();
