// ════════════════════════════════════════════════════════
// EXCEL EXPORT — Concept²
// Style theo file tham chiếu: Times New Roman 13.5,
// thin borders, minimal color, professional layout
// ════════════════════════════════════════════════════════

const XC = {
  white: 'FFFFFFFF',
  yellow: 'FFFFFF99',  // title bg
  gray:  'FFF2F2F2',
  black: 'FF000000',
  blue:  'FF0070C0',
  red:   'FFFF0000',
  green: 'FF00B050',
  lightBlue: 'FFDCE6F1', // group header
  lightGreen: 'FFE2EFDA', // summary
};

// ─── FONT (Times New Roman 13.5 = mẫu) ───
function ft(sz=13.5, bold=false, color=XC.black, name='Times New Roman') {
  return { name, size: sz, bold, color: { argb: color } };
}

// ─── FILL ───
function fl(argb) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb } };
}

// ─── BORDERS ───
const bThin = { style: 'thin', color: { argb: 'FF000000' } };
const bMed  = { style: 'medium', color: { argb: XC.black } };

const BDR_THIN = { top: bThin, bottom: bThin, left: bThin, right: bThin };
const BDR_MED  = { top: bMed, bottom: bMed, left: bMed, right: bMed };

// ─── ALIGNMENT ───
const AC = { horizontal: 'center', vertical: 'center', wrapText: true };
const AL = { horizontal: 'left', vertical: 'center', wrapText: true };
const AR = { horizontal: 'right', vertical: 'center', wrapText: false };

// ─── HELPERS ───

// Title row — yellow bg, font 16 bold, medium border, merged
function titleRow(ws, r, text, colCount) {
  ws.mergeCells(r, 1, r, colCount);
  const c = ws.getCell(r, 1);
  c.value = text;
  c.font = ft(16, true);
  c.fill = fl(XC.yellow);
  c.alignment = AC;
  c.border = BDR_MED;
  ws.getRow(r).height = 47;
  return r + 1;
}

// Info rows — merged left/right, thin border
function infoRow(ws, r, leftText, rightText, colCount) {
  const mid = Math.ceil(colCount / 2);
  ws.mergeCells(r, 1, r, mid);
  const c1 = ws.getCell(r, 1);
  c1.value = leftText;
  c1.font = ft(13.5, true);
  c1.alignment = AC;
  c1.border = BDR_THIN;

  if (mid + 1 <= colCount) {
    ws.mergeCells(r, mid + 1, r, colCount);
    const c2 = ws.getCell(r, mid + 1);
    c2.value = rightText || '';
    c2.font = ft(13.5, true);
    c2.alignment = AC;
    c2.border = BDR_THIN;
  }
  ws.getRow(r).height = 21;
  return r + 1;
}

// Header row — bold, thin border, wrap text
function hdrRow(ws, r, cols) {
  cols.forEach((h, i) => {
    const c = ws.getCell(r, i + 1);
    c.value = h;
    c.font = ft(13.5, true);
    c.alignment = AC;
    c.border = BDR_THIN;
  });
  ws.getRow(r).height = 35;
  return r + 1;
}

// Data row — normal or bold, thin border
function dRow(ws, r, vals, opts = {}) {
  const bold = opts.bold !== undefined ? opts.bold : true;
  const bg = opts.bg || null;
  const fc = opts.fc || XC.black;
  const sz = opts.sz || 13.5;
  vals.forEach((v, i) => {
    const c = ws.getCell(r, i + 1);
    c.value = v ?? '';
    c.font = ft(sz, bold, fc);
    if (bg) c.fill = fl(bg);
    c.alignment = i === 0 ? AC : (i === 1 ? AL : AR);
    c.border = BDR_THIN;
  });
  ws.getRow(r).height = 20;
  return r + 1;
}

// Group row — bold, light blue bg, merged
function grpRow(ws, r, label, colCount) {
  ws.mergeCells(r, 1, r, colCount);
  const c = ws.getCell(r, 1);
  c.value = label;
  c.font = ft(13.5, true);
  c.fill = fl(XC.lightBlue);
  c.alignment = AL;
  c.border = BDR_THIN;
  ws.getRow(r).height = 22;
  return r + 1;
}

// Summary/total row — bold, green bg
function sumRow(ws, r, vals) {
  vals.forEach((v, i) => {
    const c = ws.getCell(r, i + 1);
    c.value = v ?? '';
    c.font = ft(13.5, true);
    c.fill = fl(XC.lightGreen);
    c.alignment = i === 0 ? AC : (i === 1 ? AL : AR);
    c.border = BDR_THIN;
  });
  ws.getRow(r).height = 24;
  return r + 1;
}

// ════════════════════════════════════════════════════════
// MAIN EXPORT
// ════════════════════════════════════════════════════════
async function exportExcel() {
  if (!G) { alert('Vui lòng tính toán quy mô trước.'); return; }
  const d = G, ci = d.ci, std = d.std;
  const projTitle = 'BẢNG TÍNH QUY MÔ ' + (d.tenDA || std.name).toUpperCase();

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Concept²';
  wb.created = new Date();

  // ═══════════════════════════════════════════════
  // SHEET 1: TỔNG HỢP NVTK
  // ═══════════════════════════════════════════════
  const ws0 = wb.addWorksheet('TỔNG HỢP NVTK', { views: [{ showGridLines: false }] });
  ws0.columns = [
    { width: 7 }, { width: 54 }, { width: 16 }, { width: 14 }, { width: 18 }, { width: 22 }
  ];

  let r = 1;
  r = titleRow(ws0, r, 'TỔNG HỢP NHIỆM VỤ THIẾT KẾ', 6);

  // Project info
  r = infoRow(ws0, r, 'Dự án: ' + (d.tenDA || '—'), 'Địa điểm: ' + (d.diaDiem || '—'), 6);
  r = infoRow(ws0, r, 'Chủ đầu tư: ' + (d.chuDT || '—'), 'ĐVTK: ' + (d.dvTK || '—'), 6);
  r = infoRow(ws0, r, ci.label + ' - ' + std.name, 'DT đất: ' + nf(d.dt) + ' m² - ' + d.soTang + ' tầng', 6);

  r++; // gap
  r = grpRow(ws0, r, 'A. QUY MÔ HỌC SINH VÀ LỚP HỌC', 6);
  r = hdrRow(ws0, r, ['STT', 'Chỉ tiêu', 'Kết quả', 'Đơn vị', '', '']);
  [
    [1, 'Diện tích đất', nf(d.dt), 'm²', '', ''],
    [2, 'Tiêu chuẩn đất/HS', nf(d.dtPerHS), 'm²/HS', '', ''],
    [3, 'Quy mô học sinh tối đa', nf(d.totalHS), 'HS', '', ''],
    [4, 'Sĩ số tối đa/lớp', std.siso, 'HS/lớp', '', ''],
    [5, 'Tổng số lớp', nf(d.totalLop), 'lớp', '', ''],
    [6, 'Số khối lớp', std.numK, 'khối', '', ''],
    [7, 'HS bán trú (' + Math.round(d.pctBT*100) + '%)', nf(d.hsBT), 'HS', '', ''],
  ].forEach(row => { r = dRow(ws0, r, row); });

  r++; // gap
  r = grpRow(ws0, r, 'B. CƠ CẤU NHÂN SỰ', 6);
  r = hdrRow(ws0, r, ['', 'Chức danh', 'Số lượng', 'Đơn vị', '', '']);
  r = dRow(ws0, r, ['', 'Cán bộ quản lý', nf(d.qlTotal), 'người', '', '']);
  r = dRow(ws0, r, ['', 'Giáo viên', nf(d.gvTotal), 'người', '', '']);
  r = dRow(ws0, r, ['', 'Nhân viên HC & phục vụ', nf(d.nvTotal), 'người', '', '']);
  r = sumRow(ws0, r, ['', 'TỔNG CỘNG CBGV', nf(d.totalCBGV), 'người', '', '']);

  r++; // gap
  r = grpRow(ws0, r, 'C. TỔNG HỢP DIỆN TÍCH SÀN', 6);
  r = hdrRow(ws0, r, ['', 'Hạng mục', 'Số phòng', 'DT (m²)', '', '']);

  let currentGroup = '', groupCount = 0, groupArea = 0;
  const groupSummaries = [];
  d.phongRows.forEach(row => {
    if (row.isGroup) {
      if (currentGroup) groupSummaries.push({ name: currentGroup, count: groupCount, area: groupArea });
      currentGroup = row.ten; groupCount = 0; groupArea = 0;
    } else if (!row.skipped) {
      const n = typeof row.so === 'number' ? row.so : (parseInt(row.so) || 0);
      groupCount += n; groupArea += row.dtTotal || 0;
    }
  });
  if (currentGroup) groupSummaries.push({ name: currentGroup, count: groupCount, area: groupArea });
  groupSummaries.forEach(g => { r = dRow(ws0, r, ['', g.name, nf(g.count), nf(g.area), '', '']); });
  r = dRow(ws0, r, ['', 'Hành lang, cầu thang, kỹ thuật (40%)', '', nf(d.dtHL), '', ''], { bg: XC.lightGreen });
  r = sumRow(ws0, r, ['', 'TỔNG DT SÀN YÊU CẦU', nf(d.totalPhong) + ' phòng', nf(d.totalDTPhong) + ' m²', '', '']);

  r++; // gap
  r = grpRow(ws0, r, 'D. CƠ CẤU SỬ DỤNG ĐẤT', 6);
  r = hdrRow(ws0, r, ['STT', 'Chỉ tiêu', 'Giá trị', 'Đơn vị', 'Tỷ lệ', '']);
  const dtPct = (v) => d.dt > 0 ? (v / d.dt * 100).toFixed(1) + '%' : '—';
  [
    [1, 'Diện tích khu đất', nf(d.dt), 'm²', '100%', ''],
    [2, 'DT xây dựng (footprint)', nf(d.dtXD), 'm²', d.matDo + '%', ''],
    [3, 'DT sàn nổi thực tế', nf(d.dtSan), 'm²', '', ''],
    [4, 'DT sàn yêu cầu', nf(d.totalDTPhong), 'm²', '', ''],
    [5, 'Số tầng', d.soTang, 'tầng', '', ''],
    [6, 'Sân chơi, TDTT', nf(d.sanChoi), 'm²', dtPct(d.sanChoi), ''],
    [7, 'Cây xanh', nf(d.cayXanh), 'm²', dtPct(d.cayXanh), ''],
    [8, 'Giao thông nội bộ', nf(d.dtGT), 'm²', '~12%', ''],
    [9, 'Bãi đỗ xe', nf(d.totalXe), 'm²', dtPct(d.totalXe), ''],
  ].forEach(row => { r = dRow(ws0, r, row); });

  // ═══════════════════════════════════════════════
  // SHEET 2: QUY MÔ PHÒNG (y hệt file mẫu)
  // ═══════════════════════════════════════════════
  const ws1 = wb.addWorksheet('QUY MÔ', { views: [{ showGridLines: false }] });
  ws1.columns = [
    { width: 6.7 }, { width: 54 }, { width: 11.3 }, { width: 17.4 },
    { width: 17 }, { width: 14.6 }, { width: 20.4 }, { width: 22.7 }, { width: 30 }
  ];

  r = 1;
  // Row 1: empty (match file mẫu starts data at row 2)
  // Row 2: Title
  ws1.mergeCells(r, 1, r, 9);
  const tc = ws1.getCell(r, 1);
  tc.value = projTitle;
  tc.font = ft(16, true);
  tc.fill = fl(XC.yellow);
  tc.alignment = AC;
  tc.border = BDR_MED;
  ws1.getRow(r).height = 47;
  r++;

  // Info rows (DT đất, DT mở đường, DT xây dựng, quy mô HS, số lớp, sĩ số)
  r = infoRow(ws1, r, 'Diện tích đất', nf(d.dt) + ' m²', 9);
  r = infoRow(ws1, r, 'Quy mô học sinh', nf(d.totalHS) + ' HS', 9);
  r = infoRow(ws1, r, 'Số lớp học', nf(d.totalLop) + ' lớp', 9);
  r = infoRow(ws1, r, 'Quy mô học sinh một lớp', std.siso + ' HS/lớp', 9);

  // Header row (9 columns like file mẫu)
  r = hdrRow(ws1, r, [
    'STT', 'Khối phòng chức năng chính',
    'Số phòng\n(phòng)', 'Biên chế\n(người)',
    'Tiêu chuẩn\nm²/người', 'DT/phòng\n(m²)',
    'DT yêu cầu XD\n(m²)', 'DT thiết kế\n(m²)',
    'Ghi chú'
  ]);

  // Data rows from phongRows
  let stt = 0;
  d.phongRows.forEach(row => {
    if (row.isGroup) {
      r = grpRow(ws1, r, row.ten, 9);
      stt = 0;
    } else {
      stt++;
      if (row.skipped) {
        r = dRow(ws1, r, [stt, row.ten, '—', '', '', '—', '—', '', row.ref || ''], { bold: false, fc: 'FF999999' });
      } else {
        r = dRow(ws1, r, [stt, row.ten, row.so, '', '', nf(row.dtDon), nf(row.dtTotal), '', row.ref || '']);
      }
    }
  });

  // K2 row + total
  r = dRow(ws1, r, ['', 'Hành lang, cầu thang, kỹ thuật (K₂=40%)', '', '', '', '', nf(d.dtHL), '', 'TCVN 4319:2012'], { bg: XC.lightGreen });
  r = sumRow(ws1, r, ['', 'TỔNG DIỆN TÍCH SÀN YÊU CẦU', nf(d.totalPhong), '', '', '', nf(d.totalDTPhong), '', '']);

  // ═══════════════════════════════════════════════
  // SHEET 3: CƠ CẤU SỬ DỤNG ĐẤT (like sheet 2 of mẫu)
  // ═══════════════════════════════════════════════
  const ws2 = wb.addWorksheet('CƠ CẤU SỬ DỤNG ĐẤT', { views: [{ showGridLines: false }] });
  ws2.columns = [
    { width: 7 }, { width: 40 }, { width: 16 }, { width: 12 }
  ];

  r = 1;
  r = titleRow(ws2, r, 'CHỈ TIÊU SỬ DỤNG ĐẤT', 4);
  r = hdrRow(ws2, r, ['TT', 'Nội dung', 'Chỉ tiêu', 'Đơn vị']);

  const matdoOk = parseFloat(d.matDo) <= 40;
  [
    ['A', 'Diện tích khu đất', nf(d.dt), 'm²'],
    [1, 'Diện tích xây dựng', nf(d.dtXD), 'm²'],
    [2, 'Diện tích sàn xây dựng nổi', nf(d.dtSan), 'm²'],
    [3, 'Mật độ xây dựng', d.matDo, '%'],
    [4, 'Tầng cao', d.soTang, 'tầng'],
    [5, 'Hệ số sử dụng đất', d.dt > 0 ? (d.dtSan / d.dt).toFixed(2) : '—', 'lần'],
    [6, 'Sân chơi, TDTT (' + ci.sanChoi + 'm²/HS)', nf(d.sanChoi), 'm²'],
    [7, 'Cây xanh (' + ci.cayXanh + 'm²/HS)', nf(d.cayXanh), 'm²'],
    [8, 'Giao thông nội bộ (~12%)', nf(d.dtGT), 'm²'],
    [9, 'Bãi đỗ xe', nf(d.totalXe), 'm²'],
  ].forEach(row => {
    const fc = (row[0] === 3 && !matdoOk) ? XC.red : XC.black;
    r = dRow(ws2, r, row, { fc });
  });

  // ═══════════════════════════════════════════════
  // SHEET 4: CBGV
  // ═══════════════════════════════════════════════
  const ws3 = wb.addWorksheet('CBGV', { views: [{ showGridLines: false }] });
  ws3.columns = [
    { width: 7 }, { width: 44 }, { width: 14 }, { width: 12 }, { width: 24 }, { width: 24 }
  ];

  r = 1;
  r = titleRow(ws3, r, 'BẢNG TÍNH BIÊN CHẾ CÁN BỘ GIÁO VIÊN', 6);
  r = infoRow(ws3, r, ci.label + ' - ' + std.name, nf(d.totalHS) + ' HS - ' + nf(d.totalLop) + ' lớp', 6);
  r = hdrRow(ws3, r, ['STT', 'Chức danh', 'Số lượng', 'Đơn vị', 'Ghi chú', 'Căn cứ']);

  r = grpRow(ws3, r, 'A. CÁN BỘ QUẢN LÝ', 6);
  d.qlRows.forEach((row, i) => {
    r = dRow(ws3, r, [i + 1, row.ten, nf(row.soCalc), 'người', row.note || '', row.ref || '']);
  });
  r = sumRow(ws3, r, ['', 'Tổng quản lý', nf(d.qlTotal), 'người', '', '']);

  r = grpRow(ws3, r, 'B. GIÁO VIÊN', 6);
  d.gvRows.forEach((row, i) => {
    r = dRow(ws3, r, [i + 1, row.ten, nf(row.so), 'người', row.note || '', row.ref || '']);
  });
  r = sumRow(ws3, r, ['', 'Tổng giáo viên', nf(d.gvTotal), 'người', '', '']);

  r = grpRow(ws3, r, 'C. NHÂN VIÊN HC VÀ PHỤC VỤ', 6);
  d.nvRows.forEach((row, i) => {
    r = dRow(ws3, r, [i + 1, row.ten, nf(row.soCalc), 'người', row.note || '', row.ref || '']);
  });
  r = sumRow(ws3, r, ['', 'Tổng nhân viên', nf(d.nvTotal), 'người', '', '']);
  r = sumRow(ws3, r, ['', 'TỔNG CỘNG TOÀN TRƯỜNG', nf(d.totalCBGV), 'người', '', '']);

  // ═══════════════════════════════════════════════
  // SHEET 5: BÃI XE
  // ═══════════════════════════════════════════════
  const ws4 = wb.addWorksheet('BÃI XE', { views: [{ showGridLines: false }] });
  ws4.columns = [
    { width: 7 }, { width: 40 }, { width: 12 }, { width: 10 }, { width: 14 }, { width: 14 }, { width: 24 }
  ];

  r = 1;
  r = titleRow(ws4, r, 'BẢNG TÍNH BÃI ĐỖ XE', 7);
  r = hdrRow(ws4, r, ['STT', 'Đối tượng', 'Số xe', 'm²/xe', 'DT thuần', 'DT×1.3', 'Căn cứ']);

  r = grpRow(ws4, r, 'A. ĐỂ XE HỌC SINH — ' + std.xeHS.t, 7);
  r = dRow(ws4, r, [1, 'HS đi xe (' + (std.xeHS.pct * 100).toFixed(0) + '%)', nf(d.soXeHS), std.xeHS.dp || '–', nf(Math.round((d.soXeHS || 0) * (std.xeHS.dp || 0))), nf(d.dtXeHS), 'TCVN 8794']);

  r = grpRow(ws4, r, 'B. ĐỂ XE MÁY CBGV', 7);
  r = dRow(ws4, r, [2, 'CBGV xe máy (90%)', nf(d.soXeGV), '2.5', nf(Math.round(d.soXeGV * 2.5)), nf(d.dtXeGV), 'TCVN 8794']);

  r = grpRow(ws4, r, 'C. ĐÓN TRẢ HỌC SINH', 7);
  r = dRow(ws4, r, [3, 'Sảnh đón trả', '–', '–', '150', nf(150), 'TT13/2020']);

  r = sumRow(ws4, r, ['', 'TỔNG BÃI ĐỖ XE', '', '', '', nf(d.totalXe) + ' m²', '']);

  // ═══════════════════════════════════════════════
  // SHEET 6: PHÁP LÝ
  // ═══════════════════════════════════════════════
  const ws5 = wb.addWorksheet('PHÁP LÝ', { views: [{ showGridLines: false }] });
  ws5.columns = [
    { width: 7 }, { width: 28 }, { width: 12 }, { width: 38 }, { width: 34 }, { width: 14 }
  ];

  r = 1;
  r = titleRow(ws5, r, 'CĂN CỨ PHÁP LÝ', 6);
  r = hdrRow(ws5, r, ['STT', 'Số hiệu', 'Ban hành', 'Tên văn bản', 'Nội dung áp dụng', 'Tình trạng']);

  PHAP_LY.forEach((row, i) => {
    r = dRow(ws5, r, [i + 1, ...row], { bold: false });
  });

  // ─── DOWNLOAD ───
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  const fn = (d.tenDA || 'TruongHoc').replace(/[^a-zA-Z0-9\u00C0-\u024F\u1E00-\u1EFF\s]/g, '').replace(/\s+/g, '_');
  a.download = `QuyMo_${fn}_${new Date().toISOString().slice(0, 10)}.xlsx`;
  a.click();
  URL.revokeObjectURL(a.href);
}
