// ════════════════════════════════════════════════════════
// DRIVE SYNC — File System Access API
// Auto-save/load project to shared Google Drive folder
// ════════════════════════════════════════════════════════

let DRIVE_DIR = null;     // FileSystemDirectoryHandle
let DRIVE_LINKED = false;
let _driveSaveTimer = null;

// ─── LINK FOLDER ───
async function linkDriveFolder() {
  try {
    DRIVE_DIR = await window.showDirectoryPicker({ mode: 'readwrite' });
    DRIVE_LINKED = true;
    // Save handle for next session
    if ('idb' in window || typeof indexedDB !== 'undefined') {
      saveDirHandle(DRIVE_DIR);
    }
    updateDriveUI();
    showDriveToast('Đã liên kết: ' + DRIVE_DIR.name);

    // Auto-load latest project if exists
    await autoLoadFromDrive();
  } catch (e) {
    if (e.name !== 'AbortError') {
      alert('Lỗi liên kết thư mục: ' + e.message);
    }
  }
}

async function unlinkDriveFolder() {
  DRIVE_DIR = null;
  DRIVE_LINKED = false;
  clearDirHandle();
  updateDriveUI();
  showDriveToast('Đã ngắt liên kết thư mục');
}

// ─── SAVE PROJECT TO DRIVE ───
async function saveProjectToDrive() {
  if (!DRIVE_DIR || !DRIVE_LINKED) return;

  try {
    // Build project data (reuse existing saveProject logic)
    const projectData = buildProjectData();
    if (!projectData) return;

    const fileName = getProjectFileName();
    const fileHandle = await DRIVE_DIR.getFileHandle(fileName, { create: true });
    const writable = await fileHandle.createWritable();
    await writable.write(JSON.stringify(projectData, null, 2));
    await writable.close();

    updateDriveStatus('saved');
    console.log('[Drive] Saved:', fileName);
  } catch (e) {
    console.error('[Drive] Save error:', e);
    updateDriveStatus('error');
  }
}

// ─── AUTO SAVE (debounced) ───
function triggerDriveAutoSave() {
  if (!DRIVE_LINKED) return;
  clearTimeout(_driveSaveTimer);
  updateDriveStatus('saving');
  _driveSaveTimer = setTimeout(async () => {
    await saveProjectToDrive();
    // Also save PPTX template if available
    await saveTemplateToDrive();
  }, 2000);
}

// ─── SAVE PPTX TEMPLATE ───
async function saveTemplateToDrive() {
  if (!DRIVE_DIR || !PPTX_DATA) return;
  try {
    const fh = await DRIVE_DIR.getFileHandle('_template.pptx', { create: true });
    const w = await fh.createWritable();
    await w.write(PPTX_DATA);
    await w.close();
  } catch (e) {
    console.error('[Drive] Template save error:', e);
  }
}

// ─── LOAD FROM DRIVE ───
async function autoLoadFromDrive() {
  if (!DRIVE_DIR) return;

  try {
    // Find latest .json project file
    let latestFile = null;
    let latestTime = 0;

    for await (const [name, handle] of DRIVE_DIR) {
      if (handle.kind === 'file' && name.endsWith('.json') && !name.startsWith('_')) {
        const file = await handle.getFile();
        if (file.lastModified > latestTime) {
          latestTime = file.lastModified;
          latestFile = file;
        }
      }
    }

    if (latestFile) {
      const text = await latestFile.text();
      const data = JSON.parse(text);
      applyProjectData(data);
      showDriveToast('Đã tải: ' + latestFile.name);
    }

    // Load template if exists
    try {
      const tplHandle = await DRIVE_DIR.getFileHandle('_template.pptx');
      const tplFile = await tplHandle.getFile();
      PPTX_DATA = await tplFile.arrayBuffer();
      // Trigger parse
      const blob = new Blob([PPTX_DATA], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
      const fakeFile = new File([blob], tplFile.name, { type: blob.type });
      const dt = new DataTransfer();
      dt.items.add(fakeFile);
      const input = document.getElementById('pptx-upload');
      if (input) { input.files = dt.files; input.dispatchEvent(new Event('change')); }
      showDriveToast('Đã tải template từ thư mục');
    } catch (_) { /* no template yet */ }

  } catch (e) {
    console.error('[Drive] Load error:', e);
  }
}

// ─── LIST PROJECT FILES ───
async function showDriveProjects() {
  if (!DRIVE_DIR) { alert('Chưa liên kết thư mục.'); return; }

  const files = [];
  for await (const [name, handle] of DRIVE_DIR) {
    if (handle.kind === 'file' && name.endsWith('.json') && !name.startsWith('_')) {
      const file = await handle.getFile();
      files.push({ name, size: file.size, modified: file.lastModified });
    }
  }
  files.sort((a, b) => b.modified - a.modified);

  const existing = document.getElementById('drive-projects-dialog');
  if (existing) existing.remove();

  const overlay = document.createElement('div');
  overlay.id = 'drive-projects-dialog';
  overlay.className = 'room-dialog-overlay';
  overlay.innerHTML = `
    <div class="room-dialog" style="max-width:520px">
      <h3 style="margin:0 0 12px">Dự án trong thư mục</h3>
      <p style="font-size:11px;color:var(--text-3);margin:0 0 8px">${DRIVE_DIR.name}</p>
      ${files.length === 0 ? '<p style="color:var(--text-3);text-align:center;margin:20px 0">Chưa có dự án nào</p>' : `
        <div style="max-height:300px;overflow-y:auto">
          <table style="width:100%;border-collapse:collapse;font-size:13px">
            ${files.map(f => `
              <tr style="border-bottom:1px solid var(--border);cursor:pointer" onclick="loadDriveProject('${f.name}')">
                <td style="padding:8px;font-weight:600">${f.name.replace('.json','')}</td>
                <td style="padding:8px;font-size:11px;color:var(--text-3)">${new Date(f.modified).toLocaleString('vi-VN')}</td>
                <td style="padding:8px;font-size:11px;color:var(--text-3)">${(f.size / 1024).toFixed(1)} KB</td>
              </tr>
            `).join('')}
          </table>
        </div>
      `}
      <div style="text-align:right;margin-top:12px">
        <button class="btn btn-ghost" onclick="document.getElementById('drive-projects-dialog').remove()">Đóng</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
}

async function loadDriveProject(fileName) {
  if (!DRIVE_DIR) return;
  try {
    const fh = await DRIVE_DIR.getFileHandle(fileName);
    const file = await fh.getFile();
    const text = await file.text();
    const data = JSON.parse(text);
    applyProjectData(data);
    showDriveToast('Đã mở: ' + fileName);
    document.getElementById('drive-projects-dialog')?.remove();
  } catch (e) {
    alert('Lỗi mở file: ' + e.message);
  }
}

// ─── BUILD/APPLY PROJECT DATA ───
function buildProjectData() {
  const form = document.getElementById('main-form');
  if (!form) return null;
  const inputs = {};
  form.querySelectorAll('input, select').forEach(el => {
    if (el.id) inputs[el.id] = el.value;
  });

  // Slide edits
  const slideEdits = {};
  if (typeof SLIDES !== 'undefined' && SLIDES) {
    SLIDES.forEach((s, i) => {
      slideEdits[`s${i}`] = { slideNum: s.slideNum, title: s.title };
      const noteEl = document.getElementById(`note-${s.slideNum}`);
      if (noteEl) slideEdits[`s${i}`].note = noteEl.value;
      if (s.textBlocks) {
        s.textBlocks.forEach((tb, bi) => {
          const el = document.getElementById(`tblock-${s.slideNum}-${bi}`);
          if (el) slideEdits[`s${i}`][`tb${bi}`] = el.value;
        });
      }
    });
  }

  return {
    _version: 2,
    _savedAt: new Date().toISOString(),
    _savedBy: typeof syncGetUser === 'function' ? syncGetUser().name : 'local',
    inputs,
    slideEdits,
  };
}

function applyProjectData(data) {
  if (!data) return;
  // Apply inputs
  if (data.inputs) {
    const form = document.getElementById('main-form');
    if (form) {
      Object.keys(data.inputs).forEach(key => {
        const el = document.getElementById(key);
        if (el) el.value = data.inputs[key];
      });
      // Trigger recalc
      const calcBtn = document.getElementById('btn-calc');
      if (calcBtn) calcBtn.click();
    }
  }
}

function getProjectFileName() {
  const tenDA = document.getElementById('tenDA')?.value?.trim();
  const name = tenDA
    ? tenDA.replace(/[^a-zA-Z0-9\u00C0-\u024F\u1E00-\u1EFF\s]/g, '').replace(/\s+/g, '_').substring(0, 40)
    : 'du_an';
  return name + '.json';
}

// ─── UI HELPERS ───
function updateDriveUI() {
  const btn = document.getElementById('drive-link-btn');
  const status = document.getElementById('drive-status');
  if (btn) {
    btn.textContent = DRIVE_LINKED ? '📁 ' + DRIVE_DIR.name : '📁 Liên kết thư mục';
    btn.title = DRIVE_LINKED ? 'Click để xem dự án trong ' + DRIVE_DIR.name : 'Liên kết thư mục Google Drive';
    btn.onclick = DRIVE_LINKED ? showDriveProjects : linkDriveFolder;
  }
  if (status) {
    status.style.display = DRIVE_LINKED ? 'inline' : 'none';
  }
}

function updateDriveStatus(state) {
  const el = document.getElementById('drive-status');
  if (!el) return;
  if (state === 'saving') { el.textContent = '💾'; el.title = 'Đang lưu...'; }
  else if (state === 'saved') { el.textContent = '✓'; el.title = 'Đã lưu'; el.style.color = '#22c55e'; setTimeout(() => { el.textContent = ''; }, 2000); }
  else if (state === 'error') { el.textContent = '!'; el.title = 'Lỗi lưu'; el.style.color = '#ef4444'; }
}

function showDriveToast(msg) {
  if (typeof showSyncToast === 'function') { showSyncToast(msg); return; }
  // Fallback
  let container = document.getElementById('sync-toast-container');
  if (!container) {
    container = document.createElement('div');
    container.id = 'sync-toast-container';
    document.body.appendChild(container);
  }
  const toast = document.createElement('div');
  toast.className = 'sync-toast';
  toast.textContent = msg;
  container.appendChild(toast);
  requestAnimationFrame(() => toast.classList.add('show'));
  setTimeout(() => { toast.classList.remove('show'); setTimeout(() => toast.remove(), 300); }, 3000);
}

// ─── PERSIST DIR HANDLE (IndexedDB) ───
function saveDirHandle(handle) {
  const req = indexedDB.open('c2_drive', 1);
  req.onupgradeneeded = e => e.target.result.createObjectStore('handles');
  req.onsuccess = e => {
    const db = e.target.result;
    const tx = db.transaction('handles', 'readwrite');
    tx.objectStore('handles').put(handle, 'dir');
  };
}

function clearDirHandle() {
  const req = indexedDB.open('c2_drive', 1);
  req.onupgradeneeded = e => e.target.result.createObjectStore('handles');
  req.onsuccess = e => {
    const db = e.target.result;
    const tx = db.transaction('handles', 'readwrite');
    tx.objectStore('handles').delete('dir');
  };
}

async function restoreDirHandle() {
  return new Promise(resolve => {
    const req = indexedDB.open('c2_drive', 1);
    req.onupgradeneeded = e => e.target.result.createObjectStore('handles');
    req.onsuccess = async e => {
      const db = e.target.result;
      const tx = db.transaction('handles', 'readonly');
      const getReq = tx.objectStore('handles').get('dir');
      getReq.onsuccess = async () => {
        const handle = getReq.result;
        if (handle) {
          // Verify permission
          const perm = await handle.queryPermission({ mode: 'readwrite' });
          if (perm === 'granted') {
            DRIVE_DIR = handle;
            DRIVE_LINKED = true;
            updateDriveUI();
            resolve(true);
            return;
          }
          // Try request (needs user gesture, so just store for later)
          DRIVE_DIR = handle;
        }
        resolve(false);
      };
      getReq.onerror = () => resolve(false);
    };
    req.onerror = () => resolve(false);
  });
}

// ─── AUTO-SAVE HOOKS ───
function hookDriveAutoSave() {
  // Listen for form changes
  const form = document.getElementById('main-form');
  if (form) {
    form.addEventListener('change', triggerDriveAutoSave);
    form.addEventListener('input', triggerDriveAutoSave);
  }
  // Listen for slide edits
  document.addEventListener('input', e => {
    if (e.target.matches('.tb-edit, .content-input')) {
      triggerDriveAutoSave();
    }
  });
}

// ─── INIT ON LOAD ───
document.addEventListener('DOMContentLoaded', async () => {
  // Check File System Access API support
  if (!('showDirectoryPicker' in window)) {
    const btn = document.getElementById('drive-link-btn');
    if (btn) { btn.style.display = 'none'; }
    return;
  }
  // Try restore previous link
  await restoreDirHandle();
  hookDriveAutoSave();
});
