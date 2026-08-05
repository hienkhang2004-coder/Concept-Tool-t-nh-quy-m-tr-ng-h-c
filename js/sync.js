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
let SYNC_ROOM = null;
let SYNC_USER = null;
let SYNC_ACTIVE = false;
let SYNC_LISTENERS = []; // track listeners for cleanup

// ─── INIT ───
function syncInit() {
  if (SYNC_DB) return;
  firebase.initializeApp(FIREBASE_CONFIG);
  SYNC_DB = firebase.database();
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
  presRef.set({ name: user.name, color: user.color, joinedAt: Date.now() });
  presRef.onDisconnect().remove();

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

  updateSyncUI('connected');
  console.log(`[Sync] Joined room: ${roomId} as ${user.name}`);
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
function syncInputs() {
  if (!SYNC_ACTIVE) return;
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
  Object.keys(data).forEach(key => {
    if (key.startsWith('_')) return;
    const el = document.getElementById(key);
    if (el && el.value !== data[key]) {
      el.value = data[key];
      el.dispatchEvent(new Event('change', { bubbles: true }));
    }
  });
  showSyncToast(`${data._name || 'Ai đó'} vừa cập nhật thông tin dự án`);
}

// ─── SYNC: SLIDE EDITS (text, notes, order) ───
function syncSlideEdits() {
  if (!SYNC_ACTIVE || !SLIDES || SLIDES.length === 0) return;
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
  if (!SLIDES) return;
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
}

// ─── UI: ONLINE USERS ───
function renderOnlineUsers(presence) {
  let bar = document.getElementById('sync-users');
  if (!bar) return;
  const users = Object.values(presence);
  bar.innerHTML = users.map(u =>
    `<span class="sync-avatar" style="background:${u.color}" title="${u.name}">${u.name.charAt(0)}</span>`
  ).join('');
  bar.title = users.map(u => u.name).join(', ') + ' đang online';
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
