// ════════════════════════════════════════════════════════
// NAVIGATION & THEME
// ════════════════════════════════════════════════════════

function goPage(id, btn) {
  // Hide all pages
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  // Remove active from all step items
  document.querySelectorAll('.step-item').forEach(t => t.classList.remove('active'));

  // Show target page
  document.getElementById(id).classList.add('active');

  // Activate the correct step
  if (btn && btn.classList) {
    btn.classList.add('active');
  } else {
    const step = document.querySelector(`.step-item[data-page="${id}"]`);
    if (step) step.classList.add('active');
  }

  // Mark previous steps as completed
  const steps = document.querySelectorAll('.step-item');
  let found = false;
  steps.forEach(s => {
    if (s.classList.contains('active')) {
      found = true;
      s.classList.remove('completed');
    } else if (!found) {
      s.classList.add('completed');
    } else {
      s.classList.remove('completed');
    }
  });
  if (id === 'p2' && typeof renderCharts === 'function') {
    requestAnimationFrame(() => renderCharts());
  }
  if (id === 'p4' && typeof initEditor === 'function') {
    requestAnimationFrame(() => initEditor());
  }
  // Sync current page to collab
  if (typeof syncPresencePage === 'function') syncPresencePage(id);
}

function switchTab(e, id) {
  const par = e.currentTarget.closest('.table-card, .p2-wrap');
  (par || document).querySelectorAll('.tb').forEach(b => b.classList.remove('active'));
  (par || document).querySelectorAll('.pane').forEach(p => p.classList.remove('active'));
  e.currentTarget.classList.add('active');
  document.getElementById(id).classList.add('active');
}

function toggleTheme() {
  const html = document.documentElement;
  const currentTheme = html.getAttribute('data-theme');
  const themeBtn = document.getElementById('theme-btn');

  if (currentTheme === 'dark') {
    html.setAttribute('data-theme', 'light');
    if (themeBtn) themeBtn.textContent = '🌙';
    localStorage.setItem('c2_theme', 'light');
  } else {
    html.setAttribute('data-theme', 'dark');
    if (themeBtn) themeBtn.textContent = '☀️';
    localStorage.setItem('c2_theme', 'dark');
  }
}

// ════════════════════════════════════════════════════════
// SAVE / LOAD DỰ ÁN
// ════════════════════════════════════════════════════════

const FORM_FIELDS = ['tenDA','diaDiem','chuDT','dvTK','capTruong','chuanQG','dtDat','loaiDT','pctBT','dtXD','dtSan','soTang'];
const STORAGE_KEY = 'c2_project';
const RECENT_KEY = 'c2_recent';

// Lấy data từ form
function getFormData() {
  const data = {};
  FORM_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el) data[id] = el.value;
  });
  data._savedAt = new Date().toISOString();
  data._name = data.tenDA || 'Chưa đặt tên';
  return data;
}

// Điền data vào form
function setFormData(data) {
  FORM_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el && data[id] !== undefined) {
      el.value = data[id];
    }
  });
  // Trigger tính toán lại
  if (typeof syncAndCalc === 'function') syncAndCalc();
}

// Auto-save vào localStorage (gọi sau mỗi lần tính)
function autoSave() {
  try {
    const data = getFormData();
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    // Hiện badge "đã lưu"
    showSaveStatus('Đã lưu tự động');
  } catch(e) {
    console.warn('Auto-save lỗi:', e);
  }
}

// Auto-load từ localStorage
function autoLoad() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) {
      const data = JSON.parse(raw);
      setFormData(data);
      return true;
    }
  } catch(e) {
    console.warn('Auto-load lỗi:', e);
  }
  return false;
}

// Lưu ra file JSON
function saveProject() {
  const data = getFormData();
  const json = JSON.stringify(data, null, 2);
  const blob = new Blob([json], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const name = (data.tenDA || 'DuAn').replace(/[^a-zA-Z0-9\u00C0-\u024F\u1E00-\u1EFF\s]/g, '').replace(/\s+/g, '_');
  a.download = `${name}_${new Date().toISOString().slice(0,10)}.json`;
  a.click();
  URL.revokeObjectURL(a.href);

  // Thêm vào danh sách gần đây
  addToRecent(data);
  showSaveStatus('Đã tải file');
}

// Mở file JSON
function openProject() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.json';
  input.onchange = e => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = evt => {
      try {
        const data = JSON.parse(evt.target.result);
        setFormData(data);
        addToRecent(data);
        showSaveStatus('Đã mở: ' + (data._name || file.name));
      } catch(err) {
        alert('File không hợp lệ!');
      }
    };
    reader.readAsText(file);
  };
  input.click();
}

// Tạo dự án mới (reset form)
function newProject() {
  if (!confirm('Tạo dự án mới? Dữ liệu chưa lưu sẽ mất.')) return;
  FORM_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    if (el.tagName === 'SELECT') el.selectedIndex = 0;
    else el.value = '';
  });
  localStorage.removeItem(STORAGE_KEY);
  if (typeof syncAndCalc === 'function') syncAndCalc();
  showSaveStatus('Đã tạo mới');
}

// Danh sách dự án gần đây
function addToRecent(data) {
  try {
    let recent = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]');
    // Xóa trùng (theo tên)
    recent = recent.filter(r => r._name !== data._name);
    // Thêm vào đầu
    recent.unshift({ _name: data._name, _savedAt: data._savedAt, data });
    // Giới hạn 5
    if (recent.length > 5) recent.pop();
    localStorage.setItem(RECENT_KEY, JSON.stringify(recent));
    renderRecent();
  } catch(e) {}
}

function renderRecent() {
  const el = document.getElementById('recent-list');
  if (!el) return;
  let recent;
  try { recent = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]'); } catch(e) { return; }
  if (!recent.length) {
    el.innerHTML = '<div style="color:var(--text-3);font-size:12px;padding:8px 0">Chưa có dự án nào</div>';
    return;
  }
  el.innerHTML = recent.map((r, i) => {
    const date = r._savedAt ? new Date(r._savedAt).toLocaleDateString('vi-VN') : '';
    return `<div class="recent-item" onclick="loadRecent(${i})" title="Mở ${r._name}">
      <span class="recent-name">${r._name}</span>
      <span class="recent-date">${date}</span>
    </div>`;
  }).join('');
}

function loadRecent(idx) {
  try {
    const recent = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]');
    if (recent[idx] && recent[idx].data) {
      setFormData(recent[idx].data);
      showSaveStatus('Đã mở: ' + recent[idx]._name);
    }
  } catch(e) {}
}

// Hiện trạng thái save (nhỏ gọn, tự ẩn)
function showSaveStatus(msg) {
  let el = document.getElementById('save-status');
  if (!el) {
    el = document.createElement('div');
    el.id = 'save-status';
    el.style.cssText = 'position:fixed;bottom:20px;right:20px;background:var(--card);color:var(--text-2);padding:8px 16px;border-radius:99px;font-size:12px;font-weight:600;font-family:var(--font);box-shadow:0 4px 16px rgba(0,0,0,0.1);border:1px solid var(--card-border);z-index:9999;transition:opacity 0.3s;pointer-events:none;';
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.style.opacity = '1';
  clearTimeout(el._timer);
  el._timer = setTimeout(() => { el.style.opacity = '0'; }, 2000);
}

// ════════════════════════════════════════════════════════
// INIT
// ════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  // Theme
  const savedTheme = localStorage.getItem('c2_theme');
  if (savedTheme === 'dark' || (!savedTheme && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
    document.documentElement.setAttribute('data-theme', 'dark');
    const themeBtn = document.getElementById('theme-btn');
    if (themeBtn) themeBtn.textContent = '☀️';
  }

  // Auto-load dự án đã lưu
  const loaded = autoLoad();

  // Nếu không có data cũ, chạy calc mặc định
  if (!loaded && typeof syncAndCalc === 'function') {
    syncAndCalc();
  }

  // Render danh sách gần đây
  renderRecent();

  // Default page
  goPage('p1', document.querySelector('.step-item[data-page="p1"]'));

  // Auto-save khi nhập liệu (debounced)
  let saveTimer = null;
  FORM_FIELDS.forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.addEventListener('input', () => {
        clearTimeout(saveTimer);
        saveTimer = setTimeout(autoSave, 1000);
      });
      el.addEventListener('change', () => {
        clearTimeout(saveTimer);
        saveTimer = setTimeout(autoSave, 500);
      });
    }
  });
});
