// ════════════════════════════════════════════════════════
// SYNC MODULE — Firebase Realtime Database
// Level 1: Shared state, auto sync, room-based
// Designed for easy swap to Supabase/Yjs later
// ════════════════════════════════════════════════════════

const FIREBASE_CONFIG = {
  apiKey: "AIzaSyBRB0XjASEk0qviOspNGZ-Wz2v25lWyR6c",
  authDomain: "concept-2-8c5cb.firebaseapp.com",
  databaseURL: "https://concept-2-8c5cb-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "concept-2-8c5cb",
  storageBucket: "concept-2-8c5cb.firebasestorage.app",
  messagingSenderId: "414992594341",
  appId: "1:414992594341:web:2b2fd85256bbfed4ce644c",
};

let SYNC_DB = null;
let SYNC_STORAGE = null;
let SYNC_ROOM = null;
let SYNC_USER = null;
let SYNC_ACTIVE = false;
let SYNC_LISTENERS = []; // track listeners for cleanup

// ─── INIT ───
function syncInit() {
  if (SYNC_DB) return;
  firebase.initializeApp(FIREBASE_CONFIG);
  SYNC_DB = firebase.database();
  SYNC_STORAGE = firebase.storage();
}

// ─── USER IDENTITY (simple, no auth) ───
function syncGetUser() {
  if (SYNC_USER) return SYNC_USER;
  let stored = localStorage.getItem('c2_user');
  if (stored) {
    SYNC_USER = JSON.parse(stored);
  } else {
    const names = ['An','Bình','Chi','Dũng','Giang','Hà','Khoa','Linh','Minh','Nam','Phúc','Quân','Tâm','Vân'];
    const colors = ['#0F766E','#2563EB','#D97706','#DC2626','#7C3AED','#059669','#DB2777','#EA580C'];
    SYNC_USER = {
      id: 'u_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      name: names[Math.floor(Math.random() * names.length)] + ' ' + Math.floor(Math.random() * 99),
      color: colors[Math.floor(Math.random() * colors.length)],
    };
    localStorage.setItem('c2_user', JSON.stringify(SYNC_USER));
  }
  return SYNC_USER;
}

// ─── ROOM (project session) ───
function syncJoinRoom(roomId) {
  if (!SYNC_DB) syncInit();
  SYNC_ROOM = roomId;
  SYNC_ACTIVE = true;
  const user = syncGetUser();

  // Set presence
  const presRef = SYNC_DB.ref(`rooms/${roomId}/presence/${user.id}`);
  presRef.set({ name: user.name, color: user.color, joinedAt: Date.now(), page: 'p1' });
  presRef.onDisconnect().remove();
  saveRoomToHistory(roomId);

  // Watch presence
  syncWatch('presence', (data) => {
    renderOnlineUsers(data || {});
  });

  // Watch inputs (P1 form)
  syncWatch('inputs', (data) => {
    if (data && data._by !== user.id) {
      applyRemoteInputs(data);
    }
  });

  // Watch slide order + edits
  syncWatch('slideEdits', (data) => {
    if (data && data._by !== user.id) {
      applyRemoteSlideEdits(data);
    }
  });

  // Watch template PPTX
  syncWatch('template', (data) => {
    if (data && data._by !== user.id) {
      if (data.builtinKey) {
        loadBuiltinTemplate(data.builtinKey, true);
        showSyncToast(`${data._name || 'Ai đó'} đã chọn ${data.builtinKey === '16x9' ? 'Template 16:9' : 'Template A3'}`);
      } else if (data.url) {
        downloadSharedTemplate(data);
      }
    }
  });

  // Watch images
  syncWatch('images', (data) => {
    if (!data) return;
    Object.keys(data).forEach(key => {
      if (key.startsWith('_')) return;
      const img = data[key];
      if (img._by !== user.id && img.url) {
        // Apply remote image
        if (typeof applyRemoteImage === 'function') {
          applyRemoteImage(img.si, img.ii, img.url, img.name);
        }
      }
    });
  });

  updateSyncUI('connected');
  console.log(`[Sync] Joined room: ${roomId} as ${user.name}`);
}

// ─── SYNC: IMAGE UPLOAD (high-quality base64 via Realtime DB) ───
async function syncUploadImage(si, ii, file) {
  if (!SYNC_ACTIVE || !SYNC_ROOM) return;
  const slot = SLIDES[si]?.imgSlots?.[ii];
  const code = slot?.code || `s${si}_i${ii}`;

  try {
    let dataUrl;

    // If file is already small (<800KB), send original as-is
    if (file.size < 800 * 1024) {
      dataUrl = await fileToDataUrl(file);
    } else {
      // Compress: 1920px wide, JPEG 85% → ~300-500KB (good for PPTX export)
      dataUrl = await compressImageForSync(file, 1920, 0.85);
    }

    // Firebase Realtime DB limit: ~10MB per write. Base64 ~1.37x file size.
    // 1920px JPEG 85% ≈ 300-500KB → base64 ≈ 400-700KB → safe ✅
    syncWrite(`images/${code}`, {
      si, ii,
      url: dataUrl,
      name: file.name,
    });

    const sizeKB = Math.round(dataUrl.length / 1024);
    console.log(`[Sync] Image synced: ${code} (${sizeKB}KB)`);
    showSyncToast(`Ảnh ${file.name} đã đồng bộ (${sizeKB}KB)`);
  } catch (e) {
    console.error('[Sync] Image sync error:', e);
    showSyncToast('Lỗi sync ảnh: ' + e.message);
  }
}

function fileToDataUrl(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = reject;
    r.readAsDataURL(file);
  });
}

// Compress image: resize to maxWidth + JPEG quality
function compressImageForSync(file, maxWidth, quality) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      let w = img.naturalWidth, h = img.naturalHeight;
      // Only downscale, never upscale
      if (w > maxWidth) { h = Math.round(h * maxWidth / w); w = maxWidth; }
      canvas.width = w; canvas.height = h;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0, w, h);
      resolve(canvas.toDataURL('image/jpeg', quality));
    };
    img.onerror = reject;
    img.src = URL.createObjectURL(file);
  });
}

function syncLeaveRoom() {
  if (!SYNC_ACTIVE || !SYNC_ROOM) return;
  const user = syncGetUser();
  SYNC_DB.ref(`rooms/${SYNC_ROOM}/presence/${user.id}`).remove();
  SYNC_LISTENERS.forEach(ref => ref.off());
  SYNC_LISTENERS = [];
  SYNC_ACTIVE = false;
  updateSyncUI('disconnected');
}

// ─── CORE SYNC FUNCTIONS ───

// Write data to a path under current room
function syncWrite(path, data) {
  if (!SYNC_ACTIVE || !SYNC_ROOM) return;
  const user = syncGetUser();
  const payload = { ...data, _by: user.id, _at: Date.now(), _name: user.name };
  SYNC_DB.ref(`rooms/${SYNC_ROOM}/${path}`).set(payload);
}

// Watch a path for changes
function syncWatch(path, callback) {
  if (!SYNC_DB || !SYNC_ROOM) return;
  const ref = SYNC_DB.ref(`rooms/${SYNC_ROOM}/${path}`);
  ref.on('value', snap => callback(snap.val()));
  SYNC_LISTENERS.push(ref);
}

// ─── SYNC: P1 INPUTS ───
let _isRemoteUpdate = false;

function syncInputs() {
  if (!SYNC_ACTIVE || _isRemoteUpdate) return;
  const form = document.getElementById('main-form');
  if (!form) return;
  const data = {};
  form.querySelectorAll('input, select').forEach(el => {
    if (el.id) data[el.id] = el.value;
  });
  syncWrite('inputs', data);
}

function applyRemoteInputs(data) {
  const form = document.getElementById('main-form');
  if (!form) return;
  _isRemoteUpdate = true;
  Object.keys(data).forEach(key => {
    if (key.startsWith('_')) return;
    const el = document.getElementById(key);
    if (el && el.value !== data[key]) {
      el.value = data[key];
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }
  });
  setTimeout(() => { _isRemoteUpdate = false; }, 100);
  showSyncToast(`${data._name || 'Ai đó'} vừa cập nhật thông tin dự án`);
}

// ─── SYNC: SLIDE EDITS (text, notes, order) ───
function syncSlideEdits() {
  if (!SYNC_ACTIVE || _isRemoteUpdate || !SLIDES || SLIDES.length === 0) return;
  const edits = {};
  SLIDES.forEach((s, i) => {
    edits[`s${i}`] = {
      slideNum: s.slideNum,
      title: s.title,
      note: document.getElementById(`note-${s.slideNum}`)?.value || '',
    };
    // Text blocks
    if (s.textBlocks) {
      s.textBlocks.forEach((tb, bi) => {
        const el = document.getElementById(`tblock-${s.slideNum}-${bi}`);
        if (el) edits[`s${i}`][`tb${bi}`] = el.value;
      });
    }
  });
  syncWrite('slideEdits', edits);
}

function applyRemoteSlideEdits(data) {
  if (!SLIDES || SLIDES.length === 0) return;
  _isRemoteUpdate = true;
  let changed = false;
  Object.keys(data).forEach(key => {
    if (key.startsWith('_') || !key.startsWith('s')) return;
    const idx = parseInt(key.slice(1));
    const slide = SLIDES[idx];
    if (!slide) return;

    // Note
    if (data[key].note !== undefined) {
      const noteEl = document.getElementById(`note-${slide.slideNum}`);
      if (noteEl && noteEl.value !== data[key].note && noteEl !== document.activeElement) {
        noteEl.value = data[key].note;
        changed = true;
      }
    }
    // Text blocks
    Object.keys(data[key]).forEach(tbKey => {
      if (!tbKey.startsWith('tb')) return;
      const bi = parseInt(tbKey.slice(2));
      const el = document.getElementById(`tblock-${slide.slideNum}-${bi}`);
      if (el && el.value !== data[key][tbKey] && el !== document.activeElement) {
        el.value = data[key][tbKey];
        changed = true;
      }
    });
  });
  if (changed) showSyncToast(`${data._name || 'Ai đó'} vừa sửa nội dung slide`);
  setTimeout(() => { _isRemoteUpdate = false; }, 100);
}

// ─── SYNC: PPTX TEMPLATE ───
async function syncUploadTemplate(file) {
  if (!SYNC_ACTIVE || !SYNC_ROOM || !SYNC_STORAGE) return;
  const user = syncGetUser();
  showSyncToast('Đang upload template cho team...');

  try {
    const ref = SYNC_STORAGE.ref(`rooms/${SYNC_ROOM}/template.pptx`);
    await ref.put(file);
    const url = await ref.getDownloadURL();

    // Broadcast URL to room
    syncWrite('template', {
      url: url,
      fileName: file.name,
      fileSize: file.size,
      uploadedAt: Date.now(),
    });
    showSyncToast('Template đã upload cho team ✓');
  } catch (e) {
    console.error('[Sync] Template upload error:', e);
    showSyncToast('Lỗi upload template: ' + e.message);
  }
}

async function downloadSharedTemplate(data) {
  // Skip if we already have this template
  if (PPTX_DATA && data.uploadedAt && window._lastTemplateAt === data.uploadedAt) return;
  window._lastTemplateAt = data.uploadedAt;

  showSyncToast(`${data._name || 'Ai đó'} đã chia sẻ template "${data.fileName}". Đang tải...`);

  try {
    const response = await fetch(data.url);
    const buffer = await response.arrayBuffer();
    PPTX_DATA = buffer;

    // Parse and build table (reuse existing loadPptx logic)
    if (typeof parsePptxFromBuffer === 'function') {
      await parsePptxFromBuffer(buffer, data.fileName);
    } else {
      // Fallback: trigger the existing load flow
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
      const fakeFile = new File([blob], data.fileName, { type: blob.type });
      const dt = new DataTransfer();
      dt.items.add(fakeFile);
      const input = document.getElementById('pptx-upload');
      if (input) {
        input.files = dt.files;
        input.dispatchEvent(new Event('change'));
      }
    }
    showSyncToast('Template đã tải xong ✓ ' + data.fileName);
  } catch (e) {
    console.error('[Sync] Template download error:', e);
    showSyncToast('Lỗi tải template: ' + e.message);
  }
}

// ─── UI: ONLINE USERS ───
const PAGE_LABELS = { p1: 'Đầu vào', p2: 'Quy mô', p3: 'Trình bày', p4: 'Xem slide' };

function renderOnlineUsers(presence) {
  let bar = document.getElementById('sync-users');
  if (!bar) return;
  const myId = SYNC_USER ? SYNC_USER.id : '';
  const users = Object.entries(presence).map(([id, u]) => ({ ...u, id }));
  bar.innerHTML = users.map(u => {
    const isMe = u.id === myId;
    const pageBadge = u.page ? PAGE_LABELS[u.page] || u.page : '';
    return `<span class="sync-avatar${isMe ? ' is-me' : ''}" style="background:${u.color}" title="${u.name}${pageBadge ? ' — ' + pageBadge : ''}">
      ${u.name.charAt(0)}
      ${pageBadge ? '<span class="sync-page-badge">' + pageBadge + '</span>' : ''}
    </span>`;
  }).join('');
  bar.title = users.map(u => `${u.name}${u.page ? ' (' + (PAGE_LABELS[u.page] || u.page) + ')' : ''}`).join(', ');
}

// ─── SYNC: PAGE PRESENCE ───
function syncPresencePage(pageId) {
  if (!SYNC_ACTIVE || !SYNC_ROOM) return;
  const user = syncGetUser();
  SYNC_DB.ref(`rooms/${SYNC_ROOM}/presence/${user.id}/page`).set(pageId);
}

// ─── UI: SYNC STATUS ───
function updateSyncUI(status) {
  const dot = document.getElementById('sync-dot');
  const label = document.getElementById('sync-label');
  if (!dot || !label) return;
  if (status === 'connected') {
    dot.className = 'sync-dot online';
    label.textContent = 'Đang đồng bộ';
  } else {
    dot.className = 'sync-dot offline';
    label.textContent = 'Offline';
  }
}

// ─── UI: TOAST NOTIFICATION ───
function showSyncToast(msg) {
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
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 300);
  }, 3000);
}

// ─── ROOM UI (join/create) ───
function showRoomDialog() {
  const existing = document.getElementById('room-dialog');
  if (existing) existing.remove();

  const user = syncGetUser();
  const dialog = document.createElement('div');
  dialog.id = 'room-dialog';
  dialog.className = 'room-dialog-overlay';
  dialog.innerHTML = `
    <div class="room-dialog">
      <h3>Làm việc nhóm</h3>
      <p style="font-size:13px;color:var(--text-2);margin:4px 0 12px">Tạo hoặc tham gia phòng để đồng bộ dữ liệu với team</p>
      <div style="display:flex;gap:8px;align-items:center;margin-bottom:12px">
        <span class="sync-avatar" style="background:${user.color}">${user.name.charAt(0)}</span>
        <input id="room-username" class="room-input" value="${user.name}" placeholder="Tên của bạn" style="flex:1">
      </div>
      <input id="room-id-input" class="room-input" placeholder="Mã phòng (VD: du-an-thpt-01)" value="" style="width:100%;margin-bottom:12px">
      <div style="display:flex;gap:8px;justify-content:flex-end">
        <button class="btn btn-ghost" onclick="document.getElementById('room-dialog').remove()">Hủy</button>
        <button class="btn btn-primary" onclick="doJoinRoom()">Tham gia / Tạo phòng</button>
      </div>
      <div style="margin-top:10px;text-align:center">
        <a onclick="document.getElementById('room-dialog').remove();showRoomManager()" style="cursor:pointer;color:var(--accent);font-size:12px">Quản lý phòng đã tham gia →</a>
      </div>
    </div>
  `;
  document.body.appendChild(dialog);
  document.getElementById('room-id-input').focus();
}

function doJoinRoom() {
  const roomId = document.getElementById('room-id-input').value.trim();
  const userName = document.getElementById('room-username').value.trim();
  if (!roomId) { alert('Nhập mã phòng'); return; }

  // Update username if changed
  const user = syncGetUser();
  if (userName && userName !== user.name) {
    user.name = userName;
    localStorage.setItem('c2_user', JSON.stringify(user));
  }

  syncJoinRoom(roomId);
  document.getElementById('room-dialog').remove();

  // Show room info in header
  const roomInfo = document.getElementById('sync-room-info');
  if (roomInfo) roomInfo.textContent = roomId;
}

// ─── ROOM MANAGEMENT ───
function getJoinedRooms() {
  try { return JSON.parse(localStorage.getItem('c2_rooms') || '[]'); }
  catch { return []; }
}

function saveRoomToHistory(roomId) {
  if (!roomId) return;
  let rooms = getJoinedRooms().filter(r => r.id !== roomId);
  rooms.unshift({ id: roomId, joinedAt: new Date().toISOString() });
  if (rooms.length > 50) rooms = rooms.slice(0, 50);
  localStorage.setItem('c2_rooms', JSON.stringify(rooms));
}

function backupRoom(roomId) {
  if (!SYNC_DB) syncInit();
  SYNC_DB.ref('rooms/' + roomId).once('value').then(snap => {
    const data = snap.val();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `room_${roomId}_backup.json`;
    a.click();
    URL.revokeObjectURL(a.href);
  }).catch(e => alert('Lỗi backup: ' + e.message));
}

function deleteRoom(roomId) {
  if (!confirm('Xóa phòng "' + roomId + '"?')) return;
  if (!confirm('KHÔNG thể hoàn tác! Dữ liệu sẽ mất vĩnh viễn.')) return;
  if (!SYNC_DB) syncInit();
  SYNC_DB.ref('rooms/' + roomId).remove().then(() => {
    let rooms = getJoinedRooms().filter(r => r.id !== roomId);
    localStorage.setItem('c2_rooms', JSON.stringify(rooms));
    if (SYNC_ROOM === roomId) { syncLeaveRoom(); SYNC_ROOM = null; }
    alert('Đã xóa phòng.');
    showRoomManager(); // refresh
  }).catch(e => alert('Lỗi xóa: ' + e.message));
}

function showRoomManager() {
  const existing = document.getElementById('room-manager');
  if (existing) existing.remove();
  const rooms = getJoinedRooms();
  const overlay = document.createElement('div');
  overlay.id = 'room-manager';
  overlay.className = 'room-dialog-overlay';
  overlay.innerHTML = `
    <div class="room-dialog" style="max-width:520px">
      <h3 style="margin:0 0 12px">Quản lý phòng</h3>
      ${rooms.length === 0 ? '<p style="color:var(--text-3);text-align:center;margin:20px 0">Chưa có phòng nào</p>' : `
        <div style="max-height:300px;overflow-y:auto">
          <table style="width:100%;border-collapse:collapse;font-size:13px">
            <thead><tr style="border-bottom:1px solid var(--border)">
              <th style="padding:6px;text-align:left">Mã phòng</th>
              <th style="padding:6px;text-align:left">Ngày</th>
              <th style="padding:6px;text-align:right">Thao tác</th>
            </tr></thead>
            <tbody>${rooms.map(r => `
              <tr style="border-bottom:1px solid var(--border)">
                <td style="padding:6px;font-weight:600">${r.id}</td>
                <td style="padding:6px;font-size:11px;color:var(--text-3)">${r.joinedAt ? new Date(r.joinedAt).toLocaleDateString('vi-VN') : '—'}</td>
                <td style="padding:6px;text-align:right">
                  <button class="btn btn-ghost" onclick="backupRoom('${r.id}')" style="font-size:11px;padding:3px 8px">Backup</button>
                  <button class="btn btn-ghost" onclick="deleteRoom('${r.id}')" style="font-size:11px;padding:3px 8px;color:var(--red)">Xóa</button>
                </td>
              </tr>
            `).join('')}</tbody>
          </table>
        </div>
      `}
      <div style="text-align:right;margin-top:12px">
        <button class="btn btn-ghost" onclick="document.getElementById('room-manager').remove()">Đóng</button>
      </div>
    </div>
  `;
  document.body.appendChild(overlay);
}

// ─── AUTO-SYNC HOOKS ───
// Debounced sync on input changes
let _syncInputTimer = null;
function hookSyncInputs() {
  const form = document.getElementById('main-form');
  if (!form) return;
  form.addEventListener('input', () => {
    if (!SYNC_ACTIVE) return;
    clearTimeout(_syncInputTimer);
    _syncInputTimer = setTimeout(syncInputs, 500);
  });
  form.addEventListener('change', () => {
    if (!SYNC_ACTIVE) return;
    clearTimeout(_syncInputTimer);
    _syncInputTimer = setTimeout(syncInputs, 200);
  });
}

let _syncSlideTimer = null;
function hookSyncSlides() {
  document.addEventListener('input', (e) => {
    if (!SYNC_ACTIVE) return;
    if (e.target.matches('.tb-edit, .content-input')) {
      clearTimeout(_syncSlideTimer);
      _syncSlideTimer = setTimeout(syncSlideEdits, 600);
    }
  });
}

// Init hooks on page load
document.addEventListener('DOMContentLoaded', () => {
  hookSyncInputs();
  hookSyncSlides();
});
