// ════════════════════════════════════════════════════════
// PRESENTATION MODULE
// ════════════════════════════════════════════════════════

let SLIDES = []; // [{slideNum, title, section, isDivider, imgSlots:[{code,thumbUrl,realUrl,realName}], autoData, note}]
let PPTX_DATA = null; // ArrayBuffer of original pptx
let IMG_MAP = {}; // code -> {dataUrl, file}

// NOTE: DEFAULT_TEMPLATE and AUTO_FIELD_LABELS are defined in data.js
// NOTE: getAutoValue() is defined in calc.js

function buildPresentTable(template){
  SLIDES=template||DEFAULT_TEMPLATE;
  document.getElementById('p3-empty').style.display='none';
  document.getElementById('p3-content').style.display='block';
  document.getElementById('export-bar').style.display='flex';
  document.getElementById('img-folder-btn').style.display='flex';
  document.getElementById('btn-insert-data').style.display='inline-flex';

  const totalImgs=SLIDES.reduce((a,s)=>a+s.imgSlots.length,0);
  document.getElementById('slide-count-info').textContent=SLIDES.length+' slides';
  document.getElementById('img-done-info').textContent='0/'+totalImgs+' ảnh đã upload';
  document.getElementById('pptx-info').textContent=SLIDES.length+' slides · '+totalImgs+' ảnh';
  document.getElementById('export-info').textContent=`${SLIDES.length} slides · ${totalImgs} ảnh cần ghép`;

  const table=document.getElementById('slide-table');
  // Clear all tbodies except thead
  table.querySelectorAll('tbody').forEach(tb => tb.remove());

  SLIDES.forEach((slide,si)=>{
    const tbdy = document.createElement('tbody');
    tbdy.className = 'slide-group' + (slide.isDataSlide ? ' data-slide-group' : '');
    tbdy.dataset.si = si;
    tbdy.draggable = !slide.isDivider;

    if(slide.isDivider){
      const tr=document.createElement('tr');
      tr.className='slide-divider';
      tr.innerHTML=`<td colspan="5"><span class="divider-label">── ${slide.title} ──</span></td>`;
      tbdy.appendChild(tr);
      table.appendChild(tbdy);
      return;
    }

    // Drag events
    tbdy.addEventListener('dragstart', e => {
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', si);
      tbdy.classList.add('dragging');
      requestAnimationFrame(() => tbdy.style.opacity = '0.4');
    });
    tbdy.addEventListener('dragend', () => {
      tbdy.classList.remove('dragging');
      tbdy.style.opacity = '';
      document.querySelectorAll('.slide-group.drag-over').forEach(el => el.classList.remove('drag-over'));
    });
    tbdy.addEventListener('dragover', e => {
      e.preventDefault();
      e.dataTransfer.dropEffect = 'move';
      const dragging = table.querySelector('.dragging');
      if (dragging && dragging !== tbdy) {
        tbdy.classList.add('drag-over');
      }
    });
    tbdy.addEventListener('dragleave', () => tbdy.classList.remove('drag-over'));
    tbdy.addEventListener('drop', e => {
      e.preventDefault();
      tbdy.classList.remove('drag-over');
      const fromIdx = parseInt(e.dataTransfer.getData('text/plain'));
      const toIdx = parseInt(tbdy.dataset.si);
      if (fromIdx !== toIdx && !isNaN(fromIdx) && !isNaN(toIdx)) {
        reorderSlide(fromIdx, toIdx);
      }
    });

    const hasImgs=slide.imgSlots.length>0;
    const rowspan=Math.max(1,slide.imgSlots.length);

    // Data slide badge
    const dataBadge = slide.isDataSlide ? '<span class="data-badge">QUY MÔ</span>' : '';
    const dragHandle = '<span class="drag-handle" title="Kéo để đổi vị trí">⠿</span>';

    slide.imgSlots.forEach((slot,ii)=>{
      const tr=document.createElement('tr');
      tr.id=`row-${slide.slideNum}-${ii}`;
      tr.dataset.slideIdx=si;
      tr.dataset.slotIdx=ii;

      // Col 1: slide name (rowspan)
      if(ii===0){
        const td1=document.createElement('td');
        td1.rowSpan=rowspan||1;
        td1.className='slide-name';
        td1.innerHTML=`${dragHandle}<div class="slide-num">Slide ${slide.slideNum} ${dataBadge}<button class="btn-del-slide" onclick="deleteSlide(${si})" title="Xóa slide này">✕</button></div><div class="slide-title-text">${slide.title}</div>${slide.section?`<div class="slide-section">${slide.section}</div>`:''}`;
        tr.appendChild(td1);
      }

      if(hasImgs){
        // Col 2: thumb + dimensions + drag zone
        const td2=document.createElement('td');
        const dimHtml = slot.pxW && slot.pxH
          ? `<div class="img-dim" title="Kích thước ảnh trong PPTX">
               <span class="dim-screen">${slot.pxW} × ${slot.pxH} px</span>
               <span class="dim-print">PTS 300dpi: ${slot.px300W} × ${slot.px300H}</span>
               <span class="dim-mm">${slot.mmW} × ${slot.mmH} mm</span>
             </div>`
          : '';
        td2.innerHTML=`<div class="img-slot drop-zone" data-si="${si}" data-ii="${ii}">
          <div class="thumb-wrap" id="thumb-${slide.slideNum}-${ii}" onclick="previewImage(${si},${ii})">
            <span class="thumb-placeholder">Kéo ảnh thả vào đây</span>
          </div>
          <div style="font-size:10px;color:var(--text-3);max-width:130px">${slot.label}</div>
          ${dimHtml}
          <span class="quality-badge" id="qbadge-${slide.slideNum}-${ii}"></span>
        </div>`;
        tr.appendChild(td2);

        // Col 3: code
        const td3=document.createElement('td');
        td3.innerHTML=`<span class="code-tag" title="Click để copy" onclick="navigator.clipboard.writeText('${slot.code}')">${slot.code}</span>`;
        tr.appendChild(td3);

        // Col 4: upload + move
        const td4=document.createElement('td');
        td4.innerHTML=`<div class="slot-upload">
          <label class="upload-btn">Chọn ảnh<input type="file" accept="image/*" onchange="loadSlotImage(this,${si},${ii})"></label>
          <div class="move-btns">
            <button class="move-btn" onclick="moveSlot(${si},${ii},-1)" title="Lên" ${ii===0?'disabled':''}>▲</button>
            <button class="move-btn" onclick="moveSlot(${si},${ii},1)" title="Xuống" ${ii===slide.imgSlots.length-1?'disabled':''}>▼</button>
          </div>
          <span class="slot-status s-empty" id="stat-${slide.slideNum}-${ii}">Chưa có</span>
        </div>`;
        tr.appendChild(td4);
      } else if(ii===0){
        // No images placeholder
        const td2=document.createElement('td');td2.innerHTML=`<div class="img-slot"><span style="font-size:10px;color:var(--txt3)">Không có ảnh thay thế</span></div>`;tr.appendChild(td2);
        const td3=document.createElement('td');td3.innerHTML='–';tr.appendChild(td3);
        const td4=document.createElement('td');td4.innerHTML='–';tr.appendChild(td4);
      }

      // Col 5: content (rowspan) — text blocks + notes
      if(ii===0){
        const td5=document.createElement('td');
        td5.rowSpan=rowspan||1;
        td5.className='content-cell';

        // Auto tags
        const tagsHtml=(slide.autoFields||[]).map(f=>`<span class="auto-tag">${AUTO_FIELD_LABELS[f]||f}: ${getAutoValue(f)}</span>`).join('');

        // Editable text blocks from PPTX (collapsible)
        let blocksHtml = '';
        if (slide.textBlocks && slide.textBlocks.length > 0) {
          blocksHtml = `<details class="text-blocks">
            <summary class="tb-toggle">Sửa nội dung <span class="tb-count">${slide.textBlocks.length}</span></summary>
            <div class="tb-list">
            ${slide.textBlocks.map((tb, bi) => `
              <div class="text-block">
                <textarea class="tb-edit" id="tblock-${slide.slideNum}-${bi}"
                  rows="${Math.min(Math.max(tb.text.split('\n').length, 1), 4)}"
                  placeholder="(trống)">${tb.text}</textarea>
              </div>
            `).join('')}
            </div>
          </details>`;
        }

        td5.innerHTML=`${tagsHtml?`<div class="content-tags">${tagsHtml}</div>`:''}
          ${blocksHtml}
          <textarea class="content-input" id="note-${slide.slideNum}" placeholder="Ghi chú..." rows="1">${slide.note||''}</textarea>`;
        tr.appendChild(td5);
      }

      tbdy.appendChild(tr);
    });

    // Slide with no images but needs row
    if(!hasImgs){
      const tr=document.createElement('tr');
      const td1=document.createElement('td');
      td1.className='slide-name';
      td1.innerHTML=`${dragHandle}<div class="slide-num">Slide ${slide.slideNum} ${dataBadge}<button class="btn-del-slide" onclick="deleteSlide(${si})" title="Xóa slide này">✕</button></div><div class="slide-title-text">${slide.title}</div>${slide.section?`<div class="slide-section">${slide.section}</div>`:''}`;
      tr.appendChild(td1);
      const td2=document.createElement('td');td2.innerHTML='<span style="font-size:10px;color:var(--text-3)">Không có ảnh</span>';tr.appendChild(td2);
      const td3=document.createElement('td');td3.innerHTML='–';tr.appendChild(td3);
      const td4=document.createElement('td');td4.innerHTML='–';tr.appendChild(td4);
      const td5=document.createElement('td');td5.className='content-cell';
      const tagsHtml=(slide.autoFields||[]).map(f=>`<span class="auto-tag">${AUTO_FIELD_LABELS[f]||f}: ${getAutoValue(f)}</span>`).join('');

      let blocksHtml = '';
      if (slide.textBlocks && slide.textBlocks.length > 0) {
        blocksHtml = `<details class="text-blocks">
          <summary class="tb-toggle">Sửa nội dung <span class="tb-count">${slide.textBlocks.length}</span></summary>
          <div class="tb-list">
          ${slide.textBlocks.map((tb, bi) => `
            <div class="text-block">
              <textarea class="tb-edit" id="tblock-${slide.slideNum}-${bi}"
                rows="${Math.min(Math.max(tb.text.split('\n').length, 1), 4)}"
                placeholder="(trống)">${tb.text}</textarea>
            </div>
          `).join('')}
          </div>
        </details>`;
      }

      td5.innerHTML=`${tagsHtml?`<div class="content-tags">${tagsHtml}</div>`:''}
        ${blocksHtml}
        <textarea class="content-input" id="note-${slide.slideNum}" placeholder="Ghi chú..." rows="1">${slide.note||''}</textarea>`;
      tr.appendChild(td5);
      tbdy.appendChild(tr);
    }

    table.appendChild(tbdy);
  });
}

// ─── REORDER SLIDES ───
function reorderSlide(fromIdx, toIdx) {
  const moved = SLIDES.splice(fromIdx, 1)[0];
  SLIDES.splice(toIdx > fromIdx ? toIdx - 1 : toIdx, 0, moved);
  // Re-number
  let num = 0;
  SLIDES.forEach(s => { if (!s.isDivider) { num++; s.slideNum = num; } });
  buildPresentTable(SLIDES);
}

function loadSlotImage(input,si,ii){
  const file = input.files ? input.files[0] : input;
  if(!file) return;
  const reader=new FileReader();
  reader.onload=e=>{
    const slot=SLIDES[si].imgSlots[ii];
    slot.realUrl=e.target.result;
    slot.realName=file.name;
    IMG_MAP[slot.code]={dataUrl:e.target.result,file};

    // Update thumb (bigger preview)
    const thumb=document.getElementById(`thumb-${SLIDES[si].slideNum}-${ii}`);
    if(thumb){
      thumb.innerHTML=`<img src="${e.target.result}" alt="${file.name}" title="Click để xem lớn">`;
      thumb.classList.add('thumb-replaced');
    }

    // Update status
    const stat=document.getElementById(`stat-${SLIDES[si].slideNum}-${ii}`);
    if(stat){stat.className='slot-status s-done';stat.textContent='\u2713 '+file.name;}

    // ─── Quality check: compare image resolution vs slot 300dpi ───
    checkImageQuality(e.target.result, si, ii);

    updateImgCount();
  };
  reader.readAsDataURL(file);
}

// Kiểm tra chất lượng ảnh upload vs kích thước slot
function checkImageQuality(dataUrl, si, ii) {
  const slot = SLIDES[si].imgSlots[ii];
  const badge = document.getElementById(`qbadge-${SLIDES[si].slideNum}-${ii}`);
  if (!badge || !slot.px300W) return;

  const img = new Image();
  img.onload = () => {
    const needW = slot.px300W;
    const needH = slot.px300H;
    const ratio = Math.min(img.naturalWidth / needW, img.naturalHeight / needH);

    if (ratio >= 0.95) {
      badge.className = 'quality-badge q-good';
      badge.textContent = `\u2713 ${img.naturalWidth}×${img.naturalHeight} — Du 300dpi`;
    } else if (ratio >= 0.5) {
      badge.className = 'quality-badge q-warn';
      badge.textContent = `\u26A0 ${img.naturalWidth}×${img.naturalHeight} — Thieu ${Math.round((1-ratio)*100)}%`;
    } else {
      badge.className = 'quality-badge q-bad';
      badge.textContent = `\u2717 ${img.naturalWidth}×${img.naturalHeight} — Qua nho! Can ${needW}×${needH}`;
    }
  };
  img.src = dataUrl;
}

// Preview lightbox
function previewImage(si, ii) {
  const slot = SLIDES[si].imgSlots[ii];
  if (!slot.realUrl) return;

  let lb = document.getElementById('img-lightbox');
  if (!lb) {
    lb = document.createElement('div');
    lb.id = 'img-lightbox';
    lb.innerHTML = `<div class="lb-overlay" onclick="closeLightbox()"></div>
      <div class="lb-content">
        <img id="lb-img" src="" alt="">
        <div class="lb-info" id="lb-info"></div>
        <button class="lb-close" onclick="closeLightbox()">\u2715</button>
      </div>`;
    document.body.appendChild(lb);
  }
  document.getElementById('lb-img').src = slot.realUrl;
  const info = [`${slot.code} — ${slot.label}`];
  if (slot.pxW) info.push(`Slot: ${slot.px300W}×${slot.px300H} px (300dpi) | ${slot.mmW}×${slot.mmH} mm`);
  if (slot.realName) info.push(`File: ${slot.realName}`);
  document.getElementById('lb-info').innerHTML = info.join('<br>');
  lb.classList.add('lb-active');
}

function closeLightbox() {
  const lb = document.getElementById('img-lightbox');
  if (lb) lb.classList.remove('lb-active');
}

// ─── DRAG & DROP ───
function initDragDrop() {
  document.querySelectorAll('.drop-zone').forEach(zone => {
    zone.addEventListener('dragover', e => {
      e.preventDefault();
      zone.classList.add('drag-over');
    });
    zone.addEventListener('dragleave', e => {
      e.preventDefault();
      zone.classList.remove('drag-over');
    });
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('drag-over');
      const si = parseInt(zone.dataset.si);
      const ii = parseInt(zone.dataset.ii);
      const file = e.dataTransfer.files[0];
      if (file && file.type.startsWith('image/')) {
        loadSlotImage({ files: [file] }, si, ii);
      }
    });
  });
}

function loadImageFolder(input){
  Array.from(input.files).forEach(file=>{
    const baseName=file.name.replace(/\.[^/.]+$/,'');
    // Find matching slot
    let found=false;
    SLIDES.forEach((slide,si)=>{
      slide.imgSlots.forEach((slot,ii)=>{
        if(slot.code===baseName||file.name.startsWith(slot.code)){
          const reader=new FileReader();
          reader.onload=e=>{
            slot.realUrl=e.target.result;slot.realName=file.name;
            IMG_MAP[slot.code]={dataUrl:e.target.result,file};
            const thumb=document.getElementById(`thumb-${slide.slideNum}-${ii}`);
            if(thumb){thumb.innerHTML=`<img src="${e.target.result}">`;thumb.classList.add('thumb-replaced');}
            const stat=document.getElementById(`stat-${slide.slideNum}-${ii}`);
            if(stat){stat.className='slot-status s-done';stat.textContent='✓ Đã có';}
            updateImgCount();
          };
          reader.readAsDataURL(file);
          found=true;
        }
      });
    });
    if(!found) console.log('Không khớp mã:',file.name);
  });
}

function updateImgCount(){
  const total=SLIDES.reduce((a,s)=>a+s.imgSlots.length,0);
  const done=SLIDES.reduce((a,s)=>a+s.imgSlots.filter(sl=>sl.realUrl).length,0);
  document.getElementById('img-done-info').textContent=done+'/'+total+' ảnh đã upload';
  document.getElementById('prog-bar').style.width=(total>0?done/total*100:0)+'%';
  document.getElementById('export-info').textContent=done+'/'+total+' ảnh đã ghép · '+SLIDES.length+' slides';
}

function moveSlot(si,ii,dir){
  const slots=SLIDES[si].imgSlots;
  const newIdx=ii+dir;
  if(newIdx<0||newIdx>=slots.length) return;
  [slots[ii],slots[newIdx]]=[slots[newIdx],slots[ii]];
  buildPresentTable(SLIDES);
  initDragDrop();
}

function resetPresent(){
  SLIDES=[];IMG_MAP={};PPTX_DATA=null;
  document.getElementById('p3-empty').style.display='block';
  document.getElementById('p3-content').style.display='none';
  document.getElementById('export-bar').style.display='none';
  document.getElementById('img-folder-btn').style.display='none';
  document.getElementById('pptx-info').textContent='Chưa có file mẫu';
  document.getElementById('pptx-upload').value='';
}

// ─── EMU → Pixel converters ───
function emuToPx96(emu) { return Math.round(emu * 96 / 914400); }
function emuToPx300(emu) { return Math.round(emu * 300 / 914400); }
function emuToMm(emu) { return Math.round(emu / 36000); }

// Load .pptx via JSZip — đọc slide titles, image placeholders + kích thước
async function loadPptx(input){
  if(!input.files[0]) return;
  const file=input.files[0];
  document.getElementById('pptx-info').textContent='Đang đọc '+file.name+'…';
  try{
    const buf=await file.arrayBuffer();
    PPTX_DATA=buf;
    const zip=await JSZip.loadAsync(buf);
    const slideFiles=Object.keys(zip.files).filter(k=>k.match(/^ppt\/slides\/slide\d+\.xml$/)).sort((a,b)=>{
      const na=parseInt(a.match(/\d+/)[0]),nb=parseInt(b.match(/\d+/)[0]);return na-nb;
    });
    const slides=[];
    for(let i=0;i<slideFiles.length;i++){
      const sf=slideFiles[i];
      const xml=await zip.files[sf].async('string');
      // Extract title
      let title=`Slide ${i+1}`;
      const tMatch=xml.match(/<p:sp>.*?<p:ph type="title".*?<a:t>(.*?)<\/a:t>/s);
      if(tMatch) title=tMatch[1].replace(/<[^>]+>/g,'').trim();
      else{
        const bodies=xml.matchAll(/<a:t>([\s\S]{1,80}?)<\/a:t>/g);
        const texts=[...bodies].map(m=>m[1].trim()).filter(t=>t.length>2);
        if(texts[0]) title=texts[0];
      }

      // ─── Extract images + dimensions ───
      // Find all <p:pic> elements (actual images)
      const picDims = [];
      const picRegex = /<p:pic>[\s\S]*?<\/p:pic>/g;
      let picMatch;
      while ((picMatch = picRegex.exec(xml)) !== null) {
        const picXml = picMatch[0];
        const extMatch = picXml.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
        if (extMatch) {
          picDims.push({ cx: parseInt(extMatch[1]), cy: parseInt(extMatch[2]) });
        } else {
          picDims.push(null);
        }
      }

      // Also check <p:sp> with blipFill (shapes containing images)
      const spRegex = /<p:sp>[\s\S]*?<\/p:sp>/g;
      let spMatch;
      while ((spMatch = spRegex.exec(xml)) !== null) {
        const spXml = spMatch[0];
        if (spXml.includes('<a:blip ') || spXml.includes('type="pic"')) {
          const extMatch = spXml.match(/<a:ext\s+cx="(\d+)"\s+cy="(\d+)"/);
          if (extMatch) {
            picDims.push({ cx: parseInt(extMatch[1]), cy: parseInt(extMatch[2]) });
          }
        }
      }

      const imgCount = Math.max(picDims.length, title.toLowerCase().includes('bìa')||title.toLowerCase().includes('cover')?1:0);
      // Detect dividers
      const blipCount = (xml.match(/<a:blip /g)||[]).length;
      const isDivider = title.length<40&&(xml.match(/<p:sp>/g)||[]).length<=3&&blipCount===0;

      const imgSlots=[];
      const cnt = imgCount || 0;
      for(let j=0;j<cnt;j++){
        const code=`s${String(i+1).padStart(2,'0')}_img${String(j+1).padStart(2,'0')}`;
        const dim = picDims[j];
        const slot = {
          code,
          label: `Ảnh ${j+1}/${cnt} — ${title}`,
        };
        if (dim) {
          slot.pxW = emuToPx96(dim.cx);
          slot.pxH = emuToPx96(dim.cy);
          slot.px300W = emuToPx300(dim.cx);
          slot.px300H = emuToPx300(dim.cy);
          slot.mmW = emuToMm(dim.cx);
          slot.mmH = emuToMm(dim.cy);
        }
        imgSlots.push(slot);
      }
      // ─── Extract text blocks from shapes ───
      const textBlocks = [];
      const shapeRegex = /<p:sp>([\s\S]*?)<\/p:sp>/g;
      let shMatch;
      while ((shMatch = shapeRegex.exec(xml)) !== null) {
        const spXml = shMatch[1];
        // Skip shapes that are images
        if (spXml.includes('<a:blip ') || spXml.includes('type="pic"')) continue;

        // Extract all <a:t> text from this shape
        const tMatches = [...spXml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)];
        if (tMatches.length === 0) continue;

        // Get shape name for labeling
        const nameMatch = spXml.match(/<p:cNvPr[^>]*name="([^"]*)"/)
                       || spXml.match(/<a:cNvPr[^>]*name="([^"]*)"/)
                       || spXml.match(/name="([^"]*)"/); 
        const shapeName = nameMatch ? nameMatch[1] : '';

        // Combine paragraphs: group <a:t> by <a:p>
        const paragraphs = [];
        const pRegex = /<a:p[^>]*>([\s\S]*?)<\/a:p>/g;
        let pMatch;
        while ((pMatch = pRegex.exec(spXml)) !== null) {
          const pTexts = [...pMatch[1].matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g)];
          const pText = pTexts.map(m => m[1]).join('');
          if (pText.length > 0) paragraphs.push(pText);
        }

        const fullText = paragraphs.join('\n');
        if (fullText.trim().length > 0) {
          textBlocks.push({
            shapeName,
            text: fullText,
            originalText: fullText,
          });
        }
      }

      slides.push({slideNum:i+1, title, section:'', isDivider:isDivider&&imgCount===0, imgSlots, textBlocks, autoFields:[], note:''});
    }
    SLIDES=slides;
    document.getElementById('pptx-info').textContent=file.name+' — '+slides.length+' slides';
    buildPresentTable(SLIDES);
    // Init drag & drop sau khi render bảng
    initDragDrop();
  } catch(e){
    console.error(e);
    document.getElementById('pptx-info').textContent='Lỗi đọc file: '+e.message;
    buildPresentTable(DEFAULT_TEMPLATE);
    initDragDrop();
  }
}

// ════════════════════════════════════════════════════════
// DATA SLIDE GENERATOR (PptxGenJS)
// Tạo slide "TỔNG QUAN DỰ ÁN – QUY MÔ" y hệt file mẫu
// ════════════════════════════════════════════════════════

function generateDataPptx(format) {
  if (!G) { alert('Chưa có dữ liệu. Vui lòng tính toán quy mô trước.'); return null; }
  const d = G, ci = d.ci, std = d.std;
  const isA3 = format === 'A3';

  const pptx = new PptxGenJS();
  if (isA3) {
    pptx.defineLayout({ name: 'A3_LANDSCAPE', width: 16.54, height: 11.69 });
    pptx.layout = 'A3_LANDSCAPE';
  } else {
    pptx.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 inches
  }
  pptx.author = 'Concept²';
  const SW = isA3 ? 16.54 : 13.33;
  const SH = isA3 ? 11.69 : 7.5;

  // ─── Colors ───
  const C = {
    black: '000000', white: 'FFFFFF', gray: 'F2F2F2',
    yellow: 'FFFF99', lightBlue: 'DCE6F1', lightGreen: 'E2EFDA',
    headerDark: '333333', teal: '0F766E', textGray: '666666',
  };

  // ─── SLIDE 1: QUY MÔ ───
  const slide = pptx.addSlide();
  slide.background = { color: C.white };

  // Title bar
  slide.addText('TỔNG QUAN DỰ ÁN\nQUY MÔ', {
    x: 0.3, y: 0.15, w: 3.5, h: 0.7,
    fontSize: isA3 ? 18 : 14, fontFace: 'Times New Roman', bold: true,
    color: C.black, valign: 'middle',
    border: { pt: 1, color: C.black, type: 'solid' },
  });

  // ──── LEFT TABLE: BẢNG TÍNH QUY MÔ ────
  const leftW = SW * 0.6;
  const leftX = 0.3, leftY = 1.0;

  // Info header rows
  const infoRows = [
    [{ text: 'BẢNG TÍNH QUY MÔ ' + std.name.toUpperCase(), options: { colspan: 9, align: 'center', bold: true, fontSize: 11, fill: { color: C.yellow }, border: { pt: 1, color: C.black } } }],
    [
      { text: 'Diện tích đất', options: { colspan: 4, align: 'center', bold: true, fontSize: 9, border: { pt: 0.5, color: '999999' } } },
      { text: nf(d.dt) + ' m²', options: { colspan: 5, align: 'center', bold: true, fontSize: 9, border: { pt: 0.5, color: '999999' } } },
    ],
    [
      { text: 'Quy mô học sinh', options: { colspan: 4, align: 'center', bold: true, fontSize: 9, border: { pt: 0.5, color: '999999' } } },
      { text: nf(d.totalHS) + ' HS', options: { colspan: 5, align: 'center', bold: true, fontSize: 9, border: { pt: 0.5, color: '999999' } } },
    ],
    [
      { text: 'Số lớp học', options: { colspan: 4, align: 'center', bold: true, fontSize: 9, border: { pt: 0.5, color: '999999' } } },
      { text: nf(d.totalLop) + ' lớp', options: { colspan: 5, align: 'center', bold: true, fontSize: 9, border: { pt: 0.5, color: '999999' } } },
    ],
    [
      { text: 'Quy mô học sinh một lớp', options: { colspan: 4, align: 'center', bold: true, fontSize: 9, border: { pt: 0.5, color: '999999' } } },
      { text: std.siso + ' HS/lớp', options: { colspan: 5, align: 'center', bold: true, fontSize: 9, border: { pt: 0.5, color: '999999' } } },
    ],
  ];

  // Header columns
  const colHeader = [
    { text: 'STT', options: { align: 'center', bold: true, fontSize: 8, fill: { color: C.white }, border: { pt: 0.5, color: '999999' } } },
    { text: 'Khối phòng chức năng chính', options: { align: 'center', bold: true, fontSize: 8, fill: { color: C.white }, border: { pt: 0.5, color: '999999' } } },
    { text: 'Số phòng\n(phòng)', options: { align: 'center', bold: true, fontSize: 7, fill: { color: C.white }, border: { pt: 0.5, color: '999999' } } },
    { text: 'Biên chế\n(người)', options: { align: 'center', bold: true, fontSize: 7, fill: { color: C.white }, border: { pt: 0.5, color: '999999' } } },
    { text: 'Tiêu chuẩn\nm²/người', options: { align: 'center', bold: true, fontSize: 7, fill: { color: C.white }, border: { pt: 0.5, color: '999999' } } },
    { text: 'DT/phòng\n(m²)', options: { align: 'center', bold: true, fontSize: 7, fill: { color: C.white }, border: { pt: 0.5, color: '999999' } } },
    { text: 'DT yêu cầu XD\n(m²)', options: { align: 'center', bold: true, fontSize: 7, fill: { color: C.white }, border: { pt: 0.5, color: '999999' } } },
    { text: 'DT thiết kế\n(m²)', options: { align: 'center', bold: true, fontSize: 7, fill: { color: C.white }, border: { pt: 0.5, color: '999999' } } },
    { text: 'Ghi chú', options: { align: 'center', bold: true, fontSize: 7, fill: { color: C.white }, border: { pt: 0.5, color: '999999' } } },
  ];

  // Data rows
  const dataRows = [];
  let stt = 0;
  const bdr = { pt: 0.5, color: '999999' };
  const cellOpt = (v, opts = {}) => ({ text: String(v ?? ''), options: { fontSize: 7, fontFace: 'Times New Roman', align: opts.align || 'center', bold: !!opts.bold, fill: opts.fill ? { color: opts.fill } : undefined, border: bdr, color: opts.color || C.black } });

  d.phongRows.forEach(row => {
    if (row.isGroup) {
      stt = 0;
      dataRows.push([{ text: row.ten, options: { colspan: 9, align: 'left', bold: true, fontSize: 8, fontFace: 'Times New Roman', fill: { color: C.lightBlue }, border: bdr } }]);
    } else if (!row.skipped) {
      stt++;
      dataRows.push([
        cellOpt(stt),
        cellOpt(row.ten, { align: 'left' }),
        cellOpt(row.so),
        cellOpt(''),
        cellOpt(''),
        cellOpt(row.dtDon ? nf(row.dtDon) : ''),
        cellOpt(row.dtTotal ? nf(row.dtTotal) : ''),
        cellOpt(''),
        cellOpt(row.ref || '', { align: 'left' }),
      ]);
    }
  });

  // K2 + Total
  dataRows.push([{ text: 'Hành lang, cầu thang, kỹ thuật (K₂=40%)', options: { colspan: 6, align: 'left', bold: true, fontSize: 7, fill: { color: C.lightGreen }, border: bdr } }, cellOpt(nf(d.dtHL), { bold: true, fill: C.lightGreen }), cellOpt('', { fill: C.lightGreen }), cellOpt('TCVN 4319:2012', { align: 'left', fill: C.lightGreen })]);
  dataRows.push([{ text: 'TỔNG DIỆN TÍCH SÀN YÊU CẦU', options: { colspan: 2, align: 'left', bold: true, fontSize: 8, fill: { color: C.lightGreen }, border: bdr } }, cellOpt(nf(d.totalPhong), { bold: true, fill: C.lightGreen }), cellOpt('', { fill: C.lightGreen }), cellOpt('', { fill: C.lightGreen }), cellOpt('', { fill: C.lightGreen }), cellOpt(nf(d.totalDTPhong), { bold: true, fill: C.lightGreen }), cellOpt('', { fill: C.lightGreen }), cellOpt('', { fill: C.lightGreen })]);

  const allLeftRows = [...infoRows, colHeader, ...dataRows];
  const colW = [0.4, 2.8, 0.65, 0.65, 0.7, 0.7, 0.8, 0.7, 1.0]; // total ~8.4

  slide.addTable(allLeftRows, {
    x: leftX, y: leftY, w: leftW,
    colW: colW,
    fontSize: 8, fontFace: 'Times New Roman',
    autoPage: false,
    rowH: 0.22,
  });

  // ──── RIGHT TABLE: CHỈ TIÊU SỬ DỤNG ĐẤT ────
  const rightW = SW - leftW - leftX - 0.3 - 0.2;
  const rightX = leftX + leftW + 0.2, rightY = 1.0;
  const rbdr = { pt: 0.5, color: '999999' };
  const rfs = isA3 ? 10 : 8;
  const rc = (v, opts = {}) => ({ text: String(v ?? ''), options: { fontSize: rfs, fontFace: 'Times New Roman', align: opts.align || 'center', bold: !!opts.bold, border: rbdr, fill: opts.fill ? { color: opts.fill } : undefined } });

  const rightRows = [
    [{ text: 'CHỈ TIÊU SỬ DỤNG ĐẤT', options: { colspan: 4, align: 'center', bold: true, fontSize: 11, fill: { color: C.yellow }, border: { pt: 1, color: C.black } } }],
    [rc('TT', { bold: true }), rc('Nội dung', { bold: true, align: 'left' }), rc('Chỉ tiêu', { bold: true }), rc('Đơn vị', { bold: true })],
    [rc('A', { bold: true }), rc('Diện tích khu đất', { align: 'left', bold: true }), rc(nf(d.dt), { bold: true }), rc('m²')],
    [rc('1'), rc('Diện tích xây dựng', { align: 'left' }), rc(nf(d.dtXD)), rc('m²')],
    [rc('2'), rc('Diện tích sàn xây dựng nổi', { align: 'left' }), rc(nf(d.dtSan)), rc('m²')],
    [rc('3'), rc('Mật độ xây dựng', { align: 'left', bold: true }), rc(d.matDo, { bold: true }), rc('%')],
    [rc('4'), rc('Tầng cao', { align: 'left' }), rc(d.soTang), rc('tầng')],
    [rc('5'), rc('Hệ số sử dụng đất', { align: 'left' }), rc(d.dt > 0 ? (d.dtSan / d.dt).toFixed(2) : '—'), rc('lần')],
    [rc('6'), rc('Sân chơi, TDTT (' + ci.sanChoi + 'm²/HS)', { align: 'left' }), rc(nf(d.sanChoi)), rc('m²')],
    [rc('7'), rc('Cây xanh (' + ci.cayXanh + 'm²/HS)', { align: 'left' }), rc(nf(d.cayXanh)), rc('m²')],
    [rc('8'), rc('Giao thông nội bộ (~12%)', { align: 'left' }), rc(nf(d.dtGT)), rc('m²')],
    [rc('9'), rc('Bãi đỗ xe', { align: 'left' }), rc(nf(d.totalXe)), rc('m²')],
  ];

  const rcW = isA3 ? [0.5, 2.5, 1.4, 0.9] : [0.4, 1.8, 1.1, 0.7];
  slide.addTable(rightRows, {
    x: rightX, y: rightY, w: rightW,
    colW: rcW,
    fontSize: rfs, fontFace: 'Times New Roman',
    autoPage: false,
    rowH: isA3 ? 0.32 : 0.26,
  });

  // Footer
  const footerY = SH - 0.4;
  slide.addText('* Số liệu tự động điền từ Concept² — Bảng quy mô HS', {
    x: 0.3, y: footerY, w: 6, h: 0.2,
    fontSize: 6, fontFace: 'Times New Roman', italic: true, color: C.textGray,
  });
  slide.addText((d.tenDA || '') + '  |  ' + (d.diaDiem || ''), {
    x: 0.3, y: footerY + 0.2, w: 6, h: 0.2,
    fontSize: 7, fontFace: 'Times New Roman', bold: true, color: C.teal,
  });
  slide.addText('Concept² — ' + (d.dvTK || ''), {
    x: SW - 5, y: footerY + 0.2, w: 4.5, h: 0.2,
    fontSize: 7, fontFace: 'Times New Roman', color: C.textGray, align: 'right',
  });

  return pptx;
}

// Standalone export: chỉ data slide
async function exportDataSlide(format) {
  const pptx = generateDataPptx(format || '16x9');
  if (!pptx) return;
  const fn = G?.tenDA ? G.tenDA.replace(/[^a-zA-Z0-9\u00C0-\u024F\u1E00-\u1EFF\s]/g, '').replace(/\s+/g, '_') : 'TruongHoc';
  const suffix = format === 'A3' ? '_A3' : '_16x9';
  await pptx.writeFile({ fileName: `QuyMo_${fn}${suffix}_${new Date().toISOString().slice(0, 10)}.pptx` });
}

// Insert data slide into template PPTX
async function insertDataSlide() {
  if (!G) { alert('Chưa có dữ liệu quy mô.'); return; }
  if (!PPTX_DATA) { alert('Chưa có file template.'); return; }

  // Detect template format from PPTX_DATA
  const tmpZip = await JSZip.loadAsync(PPTX_DATA);
  const presXml = await tmpZip.file('ppt/presentation.xml').async('string');
  const szMatch = presXml.match(/sldSz\s+cx="(\d+)"\s+cy="(\d+)"/);
  let format = '16x9';
  if (szMatch) {
    const cxIn = parseInt(szMatch[1]) / 914400;
    if (cxIn > 15) format = 'A3';
  }

  // Generate data PPTX matching template format
  const dataPptx = generateDataPptx(format);
  const dataBuffer = await dataPptx.write({ outputType: 'arraybuffer' });
  const dataZip = await JSZip.loadAsync(dataBuffer);

  // Find how many slides template has
  const existSlides = Object.keys(tmpZip.files)
    .filter(k => k.match(/^ppt\/slides\/slide\d+\.xml$/))
    .sort((a,b) => parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0]));
  const newNum = existSlides.length + 1;

  // Copy data slide into template as last slide
  const dataSlideXml = await dataZip.file('ppt/slides/slide1.xml').async('string');
  tmpZip.file(`ppt/slides/slide${newNum}.xml`, dataSlideXml);

  // Copy rels
  const dataRelsFile = dataZip.file('ppt/slides/_rels/slide1.xml.rels');
  if (dataRelsFile) {
    const dataRels = await dataRelsFile.async('string');
    tmpZip.file(`ppt/slides/_rels/slide${newNum}.xml.rels`, dataRels);
  }

  // Add slide reference to presentation.xml
  const newRid = `rIdDataSlide${newNum}`;
  let updatedPres = presXml.replace(
    '</p:sldIdLst>',
    `<p:sldId id="${300 + newNum}" r:id="${newRid}"/></p:sldIdLst>`
  );
  tmpZip.file('ppt/presentation.xml', updatedPres);

  // Add relationship in presentation.xml.rels
  const presRelsFile = tmpZip.file('ppt/_rels/presentation.xml.rels');
  if (presRelsFile) {
    let presRels = await presRelsFile.async('string');
    presRels = presRels.replace(
      '</Relationships>',
      `<Relationship Id="${newRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${newNum}.xml"/></Relationships>`
    );
    tmpZip.file('ppt/_rels/presentation.xml.rels', presRels);
  }

  // Add content type
  const ctFile = tmpZip.file('[Content_Types].xml');
  if (ctFile) {
    let ct = await ctFile.async('string');
    if (!ct.includes(`/ppt/slides/slide${newNum}.xml`)) {
      ct = ct.replace(
        '</Types>',
        `<Override PartName="/ppt/slides/slide${newNum}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`
      );
      tmpZip.file('[Content_Types].xml', ct);
    }
  }

  // Save back
  PPTX_DATA = await tmpZip.generateAsync({ type: 'arraybuffer' });

  // Add to SLIDES array and rebuild table
  SLIDES.push({
    slideNum: newNum,
    title: 'TỔNG QUAN DỰ ÁN – QUY MÔ',
    section: 'Data slide (auto)',
    isDivider: false,
    isDataSlide: true,
    imgSlots: [],
    textBlocks: [],
    autoFields: [],
    note: '',
  });
  buildPresentTable(SLIDES);
  alert('Đã chèn slide Quy mô vào cuối template (slide ' + newNum + ')');
}

// Delete a slide from SLIDES array
function deleteSlide(slideIdx) {
  if (!confirm('Xóa slide "' + (SLIDES[slideIdx]?.title || slideIdx) + '"?')) return;
  SLIDES.splice(slideIdx, 1);
  // Re-number slides
  let num = 0;
  SLIDES.forEach(s => { if (!s.isDivider) { num++; s.slideNum = num; } });
  buildPresentTable(SLIDES);
}

async function exportPptx(){
  if(!PPTX_DATA){alert('Vui lòng upload file .pptx mẫu trước khi xuất.');return;}
  document.getElementById('export-info').textContent='Đang xử lý…';
  try{
    const zip=await JSZip.loadAsync(PPTX_DATA);
    const slideFiles=Object.keys(zip.files).filter(k=>k.match(/^ppt\/slides\/slide\d+\.xml$/)).sort((a,b)=>{
      return parseInt(a.match(/\d+/)[0])-parseInt(b.match(/\d+/)[0]);
    });

    // ─── BUILD TAG MAP ───
    const tags = {};
    if(G){
      const d = G, ci = d.ci, std = d.std;
      // Thông tin dự án
      tags['tenDA'] = d.tenDA || '';
      tags['diaDiem'] = d.diaDiem || '';
      tags['chuDT'] = d.chuDT || '';
      tags['dvTK'] = d.dvTK || '';
      // Quy mô
      tags['totalHS'] = nf(d.totalHS);
      tags['totalLop'] = nf(d.totalLop);
      tags['totalCBGV'] = nf(d.totalCBGV);
      tags['siSo'] = String(std.siso);
      tags['numK'] = String(std.numK);
      tags['kLop'] = std.kLop.join(', ');
      tags['lopPerK'] = String(Math.round(d.totalLop / std.numK));
      tags['hsBT'] = nf(d.hsBT);
      tags['pctBT'] = String(Math.round(d.pctBT * 100));
      // Đất & xây dựng
      tags['dt'] = nf(d.dt);
      tags['dtXD'] = nf(d.dtXD);
      tags['dtSan'] = nf(d.dtSan);
      tags['soTang'] = String(d.soTang);
      tags['matDo'] = String(d.matDo);
      tags['dtPerHS'] = nf(d.dtPerHS);
      // Phòng
      tags['totalPhong'] = nf(d.totalPhong);
      tags['totalDTPhong'] = nf(d.totalDTPhong);
      tags['dtHL'] = nf(d.dtHL);
      // Bãi xe
      tags['totalXe'] = nf(d.totalXe);
      tags['dtXeHS'] = nf(d.dtXeHS);
      tags['dtXeGV'] = nf(d.dtXeGV);
      tags['soXeHS'] = nf(d.soXeHS);
      tags['soXeGV'] = nf(d.soXeGV);
      // Cơ cấu đất
      tags['sanChoi'] = nf(d.sanChoi);
      tags['cayXanh'] = nf(d.cayXanh);
      tags['dtGT'] = nf(d.dtGT);
      // CBGV chi tiết
      tags['qlTotal'] = nf(d.qlTotal);
      tags['gvTotal'] = nf(d.gvTotal);
      tags['nvTotal'] = nf(d.nvTotal);
      // Chuẩn QG
      tags['chuanQG'] = ci.label || '';
      tags['capTruong'] = std.name || '';
      tags['tcvn'] = std.tcvn || '';
    }

    let processed=0;
    for(let i=0;i<slideFiles.length&&i<SLIDES.length;i++){
      const slide=SLIDES[i];
      let xml=await zip.files[slideFiles[i]].async('string');

      // ─── TEXT REPLACEMENT ───
      // Vấn đề: PowerPoint XML thường tách "{{tenDA}}" thành nhiều <a:r>:
      //   <a:r>...<a:t>{{</a:t></a:r><a:r>...<a:t>tenDA</a:t></a:r><a:r>...<a:t>}}</a:t></a:r>
      // Giải pháp: Tìm và ghép lại text trong mỗi <a:p> paragraph, replace tags, rồi ghi lại

      xml = replaceTagsInXml(xml, tags);

      // ─── TEXT BLOCK EDITS ───
      // Apply user edits from textarea back into XML
      if (slide.textBlocks) {
        slide.textBlocks.forEach((tb, bi) => {
          const el = document.getElementById(`tblock-${slide.slideNum}-${bi}`);
          if (!el) return;
          const newText = el.value;
          if (newText === tb.originalText) return; // no change

          // Replace original text in XML paragraph by paragraph
          const origLines = tb.originalText.split('\n');
          const newLines = newText.split('\n');

          origLines.forEach((origLine, li) => {
            if (origLine.trim().length === 0) return;
            const replacement = li < newLines.length ? newLines[li] : '';
            // Escape for XML
            const escapedOrig = origLine.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const regex = new RegExp('(<a:t[^>]*>)' + escapedOrig.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') + '(</a:t>)', 'g');
            xml = xml.replace(regex, '$1' + escapeXml(replacement) + '$2');
          });

          // If user added new lines, append to last paragraph
          if (newLines.length > origLines.length) {
            const extraLines = newLines.slice(origLines.length).filter(l => l.trim());
            if (extraLines.length > 0) {
              const extraXml = extraLines.map(l => `<a:p><a:r><a:rPr lang="vi-VN" dirty="0"/><a:t>${escapeXml(l)}</a:t></a:r></a:p>`).join('');
              // Insert before </p:txBody>
              xml = xml.replace(/<\/p:txBody>/, extraXml + '</p:txBody>');
            }
          }
        });
      }

      // ─── IMAGE REPLACEMENT ───
      const blipMatches=[...xml.matchAll(/<a:blip r:embed="([^"]+)"/g)];
      for(let j=0;j<blipMatches.length&&j<slide.imgSlots.length;j++){
        const slot=slide.imgSlots[j];
        if(!slot.realUrl) continue;
        const rId=blipMatches[j][1];
        const relFile=slideFiles[i].replace('slides/slide','slides/_rels/slide').replace('.xml','.xml.rels');
        if(zip.files[relFile]){
          let relsXml=await zip.files[relFile].async('string');
          const relMatch=relsXml.match(new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`));
          if(relMatch){
            const imgPath='ppt/'+relMatch[1].replace('../','');
            const b64=slot.realUrl.split(',')[1];
            zip.file(imgPath,b64,{base64:true});
          }
        }
      }

      zip.file(slideFiles[i],xml);

      // ─── SPEAKER NOTES ───
      // Ghi notes từ textarea vào notesSlide XML
      const noteText = document.getElementById(`note-${slide.slideNum}`)?.value?.trim();
      if (noteText) {
        const noteSlideFile = `ppt/notesSlides/notesSlide${i+1}.xml`;
        if (zip.files[noteSlideFile]) {
          let noteXml = await zip.files[noteSlideFile].async('string');
          // Replace the body text in notes
          noteXml = noteXml.replace(
            /(<a:p>\s*<a:r>\s*<a:rPr[^>]*\/>\s*<a:t>)(.*?)(<\/a:t>)/s,
            `$1${escapeXml(noteText)}$3`
          );
          zip.file(noteSlideFile, noteXml);
        } else {
          // Create a simple notes slide if it doesn't exist
          const noteXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
    <p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="vi-VN" dirty="0"/><a:t>${escapeXml(noteText)}</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:notes>`;
          zip.file(noteSlideFile, noteXml);
        }
      }

      processed++;
      document.getElementById('prog-bar').style.width=(processed/slideFiles.length*100)+'%';
    }

    const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE'});
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    const fn=G?.tenDA?G.tenDA.replace(/[^a-zA-Z0-9\u00C0-\u024F\u1E00-\u1EFF\s]/g,'').replace(/\s+/g,'_'):'TruongHoc';
    a.download=`Concept_${fn}_${new Date().toISOString().slice(0,10)}.pptx`;
    a.click();
    document.getElementById('export-info').textContent='Xuất thành công! '+processed+' slides đã xử lý.';
    document.getElementById('prog-bar').style.width='100%';
  } catch(e){
    console.error(e);
    document.getElementById('export-info').textContent='Lỗi xuất file: '+e.message;
  }
}

// ─── XML TAG REPLACEMENT ENGINE ───
// Xử lý việc PowerPoint tách text tag thành nhiều <a:r> elements
function replaceTagsInXml(xml, tags) {
  // Bước 1: Thử replace trực tiếp (cho trường hợp tag không bị tách)
  Object.entries(tags).forEach(([key, val]) => {
    const tag = '{{' + key + '}}';
    xml = xml.split(tag).join(escapeXml(val));
  });

  // Bước 2: Xử lý tags bị tách trong XML
  // Tìm mỗi paragraph <a:p>...</a:p>
  xml = xml.replace(/<a:p\b[^>]*>[\s\S]*?<\/a:p>/g, (paragraph) => {
    // Ghép tất cả <a:t> text trong paragraph lại
    const textParts = [];
    const runRegex = /<a:r>([\s\S]*?)<\/a:r>/g;
    let runMatch;
    while ((runMatch = runRegex.exec(paragraph)) !== null) {
      const tMatch = runMatch[1].match(/<a:t[^>]*>([\s\S]*?)<\/a:t>/);
      if (tMatch) textParts.push({ fullRun: runMatch[0], text: tMatch[1] });
    }

    // Ghép text lại thành 1 chuỗi
    const fullText = textParts.map(p => p.text).join('');

    // Kiểm tra có tag nào trong chuỗi ghép không
    const tagPattern = /\{\{(\w+)\}\}/g;
    if (!tagPattern.test(fullText)) return paragraph; // Không có tag → giữ nguyên

    // Có tag bị tách → cần xử lý
    // Thay thế tags trong fullText
    let newText = fullText;
    Object.entries(tags).forEach(([key, val]) => {
      const tag = '{{' + key + '}}';
      newText = newText.split(tag).join(escapeXml(val));
    });

    // Nếu text không đổi → giữ nguyên
    if (newText === fullText) return paragraph;

    // Gộp tất cả text vào run đầu tiên, xóa text ở các run còn lại
    if (textParts.length > 0) {
      let isFirst = true;
      textParts.forEach(part => {
        if (isFirst) {
          // Thay text của run đầu tiên bằng newText
          const newRun = part.fullRun.replace(
            /<a:t[^>]*>[\s\S]*?<\/a:t>/,
            '<a:t>' + newText + '</a:t>'
          );
          paragraph = paragraph.replace(part.fullRun, newRun);
          isFirst = false;
        } else {
          // Xóa text của các run sau (giữ run nhưng text rỗng)
          const emptyRun = part.fullRun.replace(
            /<a:t[^>]*>[\s\S]*?<\/a:t>/,
            '<a:t></a:t>'
          );
          paragraph = paragraph.replace(part.fullRun, emptyRun);
        }
      });
    }

    return paragraph;
  });

  return xml;
}

function escapeXml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function downloadCodeList(){
  const lines=['MÃ FILE ẢNH CẦN ĐẶT TÊN\n=====================\n'];
  SLIDES.forEach(s=>{
    if(s.isDivider||!s.imgSlots.length) return;
    lines.push(`[Slide ${s.slideNum}] ${s.title}`);
    s.imgSlots.forEach(sl=>lines.push(`  → ${sl.code}.jpg  (${sl.label})`));
    lines.push('');
  });
  const blob=new Blob([lines.join('\n')],{type:'text/plain;charset=utf-8'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='ma_file_anh.txt';
  a.click();
}
