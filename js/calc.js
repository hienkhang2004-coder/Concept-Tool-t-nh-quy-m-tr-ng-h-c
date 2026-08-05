let G=null;

const nf=v=>(typeof v==='number'&&!isNaN(v))?v.toLocaleString('vi-VN',{maximumFractionDigits:0}):'–';
function th(cols){return'<tr>'+cols.map(c=>`<th>${c}</th>`).join('')+'</tr>'}
function td(cells,cls=''){return`<tr class="${cls}">`+cells.map((c,i)=>`<td class="${i>=2?'t-num':''}">${c??'–'}</td>`).join('')+'</tr>'}
function animNum(id,v){const el=document.getElementById(id);if(!el)return;el.style.transition='opacity .15s';el.style.opacity='0';setTimeout(()=>{el.textContent=nf(v);el.style.opacity='1'},150)}

function syncAndCalc(){
  const s=(id)=>document.getElementById(id)?.value||'';
  document.getElementById('s-ten').textContent=s('tenDA')||'–';
  document.getElementById('s-dia').textContent=s('diaDiem')||'–';
  document.getElementById('s-chu').textContent=s('chuDT')||'–';
  document.getElementById('s-dv').textContent=s('dvTK')||'–';
  calc();
}

function calc(){
  const cap=document.getElementById('capTruong').value;
  const cq=document.getElementById('chuanQG').value;
  const dt=parseFloat(document.getElementById('dtDat').value)||0;
  const loai=document.getElementById('loaiDT').value;
  const pctBT=parseFloat(document.getElementById('pctBT').value)/100||0;
  const dtXD=parseFloat(document.getElementById('dtXD').value)||0;
  const dtSan=parseFloat(document.getElementById('dtSan').value)||0;
  const soTang=parseInt(document.getElementById('soTang').value)||4;
  const std=STD[cap], ci=CHUAN[cq];
  const bm=cq==='m1'?'full':cq==='m2'?'basic':'min';

  const badge=document.getElementById('lv-badge');
  badge.className='lv-badge '+ci.badge;
  badge.querySelector('.dot').className='dot '+ci.dot;
  document.getElementById('lv-txt').textContent=ci.label;

  const dtPerHS=std.dtPC[loai];
  const totalHS=Math.floor(dt/dtPerHS/std.siso)*std.siso;
  const totalLop=totalHS/std.siso;
  const hsBT=Math.round(totalHS*pctBT);

  const matDo=dt>0?dtXD/dt*100:0;
  const arcFull=214;
  const arcOffset=arcFull-arcFull*Math.min(matDo,100)/100;
  const garc=document.getElementById('g-arc');
  const gtxt=document.getElementById('g-txt');
  if(garc){garc.style.strokeDashoffset=arcOffset;const c=matDo<=40?'#0F766E':matDo<=45?'#EA580C':'#DC2626';garc.style.stroke=c;gtxt.setAttribute('fill',c);gtxt.textContent=matDo.toFixed(1)+'%';}
  document.getElementById('g-num').textContent=matDo.toFixed(1)+'%';
  document.getElementById('g-sub').textContent=matDo<=40?'✓ Đạt QCVN (≤40%)':matDo<=45?'⚠ Gần ngưỡng':'✗ Vượt ngưỡng 40%';

  const loaiN={db:'Đặc biệt',1:'Loại I',2:'Loại II-III',nt:'Nông thôn'};
  document.getElementById('s-cap').textContent=std.name;
  document.getElementById('s-chuan').textContent=ci.label;
  document.getElementById('s-dt').textContent=nf(dt)+' m²';
  document.getElementById('s-xd').textContent=nf(dtXD)+' m² ('+matDo.toFixed(1)+'%)';
  document.getElementById('s-san').textContent=nf(dtSan)+' m²';
  document.getElementById('s-tang').textContent=soTang+' tầng';
  document.getElementById('s-matdo').textContent=matDo.toFixed(1)+'%';
  document.getElementById('s-bantru').textContent=Math.round(pctBT*100)+'% = '+nf(hsBT)+' HS';

  let qlTotal=0,gvTotal=0,nvTotal=0;
  const qlRows=[],gvRows=[],nvRows=[];
  qlRows.push({ten:'Hiệu trưởng',soCalc:1,ref:std.nv[0].ref}); qlTotal+=1;
  const pht=std.nv[1]; const phtS=pht.fn?pht.fn(totalHS,totalLop):pht.so;
  qlRows.push({ten:'Phó Hiệu trưởng',soCalc:phtS,note:pht.note||'',ref:pht.ref}); qlTotal+=parseInt(phtS)||0;
  if(cap==='mn'){
    const s=Math.round(totalLop*std.gvPL);
    gvRows.push({ten:'Giáo viên phụ trách nhóm/lớp',so:s,note:std.gvPL+' GV/lớp',ref:'TT52/2020'}); gvTotal+=s;
  } else {
    const gvCN=Math.round(totalLop*std.gvPL*ci.gvR);
    gvRows.push({ten:'GV chủ nhiệm & các môn chính',so:gvCN,note:std.gvPL+' GV/lớp',ref:'TT16/2017'}); gvTotal+=gvCN;
    std.gvBM.forEach(b=>{const s=Math.max(1,Math.round(totalLop*b.r));gvRows.push({ten:'GV bộ môn: '+b.n,so:s,note:(b.r*100).toFixed(0)+'% số lớp',ref:b.ref});gvTotal+=s;});
  }
  std.nv.slice(2).forEach(nv=>{const s=nv.fn?nv.fn(totalHS,totalLop,hsBT):nv.so;nvRows.push({...nv,soCalc:s});nvTotal+=parseInt(s)||0;});
  const totalCBGV=qlTotal+gvTotal+nvTotal;

  const phongRows=[];let totalPhong=0,totalDTPhong=0;
  std.phong.forEach(nh=>{
    phongRows.push({isGroup:true,ten:nh.g});
    nh.items.forEach(p=>{
      const soRaw=p.fs(totalLop,bm,totalHS,totalCBGV,totalLop);
      const soN=typeof soRaw==='number'?soRaw:0;
      const factor=p.ch?p.ch[cq]:1.0;
      let dtDon=p.fn?p.fn(totalHS,totalLop,totalCBGV,hsBT):(p.fx?Math.round(p.fx*(factor||1)):(p.m2pp?Math.round(p.m2pp*(factor||1)*p.siso):0));
      const dtTot=dtDon*(soRaw==='-'?1:soN);
      const skipped=soN===0&&soRaw!=='-';
      phongRows.push({ten:p.n,so:soRaw,dtDon,dtTotal:dtTot,ref:p.ref,skipped});
      if(!skipped&&soN>0)totalPhong+=soN;
      totalDTPhong+=dtTot;
    });
  });
  const dtHL=Math.round(totalDTPhong*.40);
  totalDTPhong+=dtHL;

  const soXeHS=cap==='mn'?0:Math.round(totalHS*std.xeHS.pct);
  const dtXeHS=Math.round(soXeHS*std.xeHS.dp*1.3);
  const soXeGV=Math.round(totalCBGV*std.xeGV.pct);
  const dtXeGV=Math.round(soXeGV*std.xeGV.dp*1.3);
  const totalXe=dtXeHS+dtXeGV+150;

  const sanChoi=Math.round(totalHS*ci.sanChoi);
  const cayXanh=Math.round(totalHS*ci.cayXanh);
  const dtGT=Math.round(dt*.12);

  G={cap,cq,bm,dt,loai,pctBT,hsBT,dtPerHS,std,ci,
    totalHS,totalLop,totalCBGV,qlRows,qlTotal,gvRows,gvTotal,nvRows,nvTotal,
    phongRows,totalPhong,totalDTPhong,dtHL,soXeHS,dtXeHS,soXeGV,dtXeGV,totalXe,
    sanChoi,cayXanh,dtXD,dtSan,soTang,matDo:matDo.toFixed(1),dtGT,
    tenDA:document.getElementById('tenDA').value,
    diaDiem:document.getElementById('diaDiem').value,
    chuDT:document.getElementById('chuDT').value,
    dvTK:document.getElementById('dvTK').value};

  ['hs1','lop1','cbgv1','san1','xe1'].forEach((id,i)=>{
    animNum('m-'+id,[totalHS,totalLop,totalCBGV,totalDTPhong,totalXe][i]);
  });
  ['hs2','lop2','cbgv2','san2','xe2'].forEach((id,i)=>{
    animNum('m-'+id,[totalHS,totalLop,totalCBGV,totalDTPhong,totalXe][i]);
  });
  document.querySelectorAll('.mc').forEach(c=>c.classList.add('lit'));

  document.getElementById('info-note').innerHTML=`<b>${ci.label}</b> · ${std.name}<br>${nf(totalHS)} học sinh · ${nf(totalLop)} lớp · ${dtPerHS}m²/HS · Đô thị ${loai==='db'?'Đặc biệt':loai==='nt'?'Nông thôn':'Loại '+loai}`;

  renderTables();
}

function renderTables(){
  const d=G;const ci=d.ci;
  const loaiN={db:'Đặc biệt (HN,HCM)',1:'Loại I',2:'Loại II-III',nt:'Nông thôn'};

  let h=`<thead>${th(['STT','Chỉ tiêu','Kết quả','Đơn vị','Ghi chú','Căn cứ'])}</thead><tbody>`;
  [['1','Diện tích đất',nf(d.dt),'m²','',''],['2','Tiêu chuẩn đất/HS',nf(d.dtPerHS),'m²/HS',loaiN[d.loai],'TCVN 8793/8794:2011'],['3','Quy mô học sinh tối đa',`<span class="v-bold">${nf(d.totalHS)}</span>`,'HS','Bội số '+d.std.siso,''],['4','Sĩ số tối đa/lớp',nf(d.std.siso),'HS/lớp','Tối đa QĐ','TT13/2020 Đ.6'],['5','Tổng số lớp',`<span class="v-bold">${nf(d.totalLop)}</span>`,'lớp','','TT13/2020 Đ.6'],['6','Số khối lớp',nf(d.std.numK),'khối',d.std.kLop.join(' · '),'Luật GD 2019'],['7','Số lớp mỗi khối',nf(Math.round(d.totalLop/d.std.numK)),'lớp/khối','',''],['8',`HS bán trú (${Math.round(d.pctBT*100)}%)`,`<span class="v-bold">${nf(d.hsBT)}</span>`,'HS','','TT13/2020 Đ.4'],['9','Chuẩn Quốc gia',`<span class="tag-ok">${ci.label}</span>`,'','','TT17/2018']].forEach(r=>h+=td(r));
  h+='</tbody>'; document.getElementById('tbl-hs').innerHTML=h;

  let c=`<thead>${th(['STT','Chức danh','Số lượng','Đơn vị','Ghi chú','Căn cứ'])}</thead><tbody>`;
  c+=`<tr class="t-grp"><td colspan="6">A. CÁN BỘ QUẢN LÝ</td></tr>`;
  d.qlRows.forEach((r,i)=>c+=td([i+1,r.ten,`<span class="v-bold">${nf(r.soCalc)}</span>`,'người',r.note||'',r.ref||'']));
  c+=td(['','Tổng quản lý',`<span class="v-bold">${nf(d.qlTotal)}</span>`,'người','',''],'t-sum');
  c+=`<tr class="t-grp"><td colspan="6">B. GIÁO VIÊN</td></tr>`;
  d.gvRows.forEach((r,i)=>c+=td([i+1,r.ten,`<span class="v-bold">${nf(r.so)}</span>`,'người',r.note||'',r.ref||'']));
  c+=td(['','Tổng giáo viên',`<span class="v-bold">${nf(d.gvTotal)}</span>`,'người','',''],'t-sum');
  c+=`<tr class="t-grp"><td colspan="6">C. NHÂN VIÊN HC & PHỤC VỤ</td></tr>`;
  d.nvRows.forEach((r,i)=>c+=td([i+1,r.ten,`<span class="v-bold">${nf(r.soCalc)}</span>`,'người',r.note||'',r.ref||'']));
  c+=td(['','Tổng nhân viên',`<span class="v-bold">${nf(d.nvTotal)}</span>`,'người','',''],'t-sum');
  c+=td(['','TỔNG CỘNG TOÀN TRƯỜNG',`<span class="v-bold">${nf(d.totalCBGV)}</span>`,'người','',''],'t-sum');
  c+='</tbody>'; document.getElementById('tbl-cbgv').innerHTML=c;

  let p=`<thead>${th(['STT','Loại phòng','Số phòng','DT/phòng (m²)','Tổng DT (m²)','Trạng thái','Căn cứ'])}</thead><tbody>`;
  let stt=0;
  d.phongRows.forEach(r=>{
    if(r.isGroup){p+=`<tr class="t-grp"><td colspan="7">${r.ten}</td></tr>`;stt=0;}
    else{stt++;p+=`<tr class="${r.skipped?'t-skip':''}"><td>${stt}</td><td>${r.ten}</td><td class="t-num">${r.skipped?'—':r.so}</td><td class="t-num">${r.skipped?'—':nf(r.dtDon)}</td><td class="t-num">${r.skipped?'—':`<span class="v-bold">${nf(r.dtTotal)}</span>`}</td><td>${r.skipped?`<span class="tag-skip">Không bắt buộc</span>`:'<span class="tag-ok">Áp dụng</span>'}</td><td style="font-size:9px;color:var(--txt2)">${r.ref}</td></tr>`;}
  });
  p+=td(['','Hành lang, CThang, kỹ thuật (K2=40%)','–','–',`<span class="v-bold">${nf(d.dtHL)}</span>`,'','TCVN 4319:2012'],'t-sum');
  p+=`<tr class="t-sum"><td></td><td>TỔNG DIỆN TÍCH SÀN YÊU CẦU</td><td class="t-num"><span class="v-bold">${nf(d.totalPhong)} phòng</span></td><td></td><td class="t-num"><span class="v-bold">${nf(d.totalDTPhong)} m²</span></td><td></td><td></td></tr>`;
  p+='</tbody>'; document.getElementById('tbl-phong').innerHTML=p;

  let x=`<thead>${th(['STT','Đối tượng','Số xe','m²/xe','DT thuần','DT×1.3','Căn cứ'])}</thead><tbody>`;
  x+=`<tr class="t-grp"><td colspan="7">A. ĐỂ XE HỌC SINH — ${d.std.xeHS.t}</td></tr>`;
  x+=td(['1',`HS đi xe (${(d.std.xeHS.pct*100).toFixed(0)}%)`,nf(d.soXeHS),d.soXeHS>0?d.std.xeHS.dp:'–',d.soXeHS>0?nf(Math.round(d.soXeHS*d.std.xeHS.dp)):'0',`<span class="v-bold">${nf(d.dtXeHS)}</span>`,'TCVN 8794 Phụ lục A']);
  x+=`<tr class="t-grp"><td colspan="7">B. ĐỂ XE MÁY CBGV (không tính ô tô)</td></tr>`;
  x+=td(['2',`CBGV xe máy (90%)`,nf(d.soXeGV),'2.5',nf(Math.round(d.soXeGV*2.5)),`<span class="v-bold">${nf(d.dtXeGV)}</span>`,'TCVN 8794 Phụ lục A']);
  x+=`<tr class="t-grp"><td colspan="7">C. ĐÓN TRẢ HỌC SINH</td></tr>`;
  x+=td(['3','Sảnh đón trả','–','–','150','<span class="v-bold">150</span>','TT13/2020']);
  x+=`<tr class="t-sum"><td colspan="5"></td><td class="t-num"><span class="v-bold">${nf(d.totalXe)} m²</span></td><td style="font-size:9px;color:var(--txt2)">Tổng bãi đỗ xe</td></tr>`;
  x+='</tbody>'; document.getElementById('tbl-xe').innerHTML=x;

  let dt2=`<thead>${th(['STT','Chỉ tiêu','Giá trị','Đơn vị','Hệ số/Tỷ lệ','Yêu cầu','Căn cứ'])}</thead><tbody>`;
  [['1','Diện tích đất xây dựng',nf(d.dt),'m²','100%','',''],['2','DT xây dựng (footprint)',nf(d.dtXD),'m²',d.matDo+'%','<span class="tag-ok">≤ 40%</span>','QCVN 01:2021/BXD'],['3','DT sàn nổi thực tế',nf(d.dtSan),'m²','','','Hồ sơ thiết kế'],['4','DT sàn yêu cầu (tính toán)',nf(d.totalDTPhong),'m²','','','Bảng phòng chức năng'],['5','Số tầng xây dựng',d.soTang,'tầng','','','Hồ sơ TK'],['6',`Sân chơi, TDTT (${ci.sanChoi}m²/HS)`,nf(d.sanChoi),'m²',ci.sanChoi+'m²/HS','','TT13/2020 Đ.8'],['7',`Cây xanh (${ci.cayXanh}m²/HS)`,nf(d.cayXanh),'m²',ci.cayXanh+'m²/HS','','TT13/2020 Đ.8'],['8','Giao thông nội bộ (~12%)',nf(d.dtGT),'m²','~12%','','Ước tính'],['9','Bãi đỗ xe (không tính ô tô)',nf(d.totalXe),'m²','','','Bảng bãi xe']].forEach(r=>dt2+=td(r));
  dt2+='</tbody>'; document.getElementById('tbl-dat').innerHTML=dt2;

  let pl=`<thead>${th(['STT','Số hiệu','Ban hành','Tên văn bản','Nội dung áp dụng','Tình trạng'])}</thead><tbody>`;
  PHAP_LY.forEach((r,i)=>pl+=td([i+1,...r])); pl+='</tbody>';
  document.getElementById('tbl-phap').innerHTML=pl;

  document.getElementById('note-legal').innerHTML=`<b>${ci.label}</b> · ${d.std.name} · Tiêu chuẩn ${d.dtPerHS}m²/HS · Sĩ số ${d.std.siso}HS/lớp · Bán trú ${Math.round(d.pctBT*100)}%=${nf(d.hsBT)}HS &nbsp;·&nbsp; <span style="color:var(--txt3)">${d.std.tcvn}</span>`;
  renderCharts();
}

function drawDonut(canvasId, legendId, data, colors) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  if (!canvas.dataset.ow) { canvas.dataset.ow = canvas.width; canvas.dataset.oh = canvas.height; }
  const w = +canvas.dataset.ow, h = +canvas.dataset.oh;
  canvas.width = w * dpr; canvas.height = h * dpr;
  canvas.style.width = w + 'px'; canvas.style.height = h + 'px';
  ctx.scale(dpr, dpr);
  const cx = w/2, cy = h/2, oR = Math.min(cx,cy)-8, iR = oR*0.6;
  const total = data.reduce((a,d) => a+d.value, 0);
  if (total === 0) { ctx.fillStyle='#ccc'; ctx.font='12px Arial'; ctx.textAlign='center'; ctx.fillText('Chưa có dữ liệu',cx,cy); return; }
  let ang = -Math.PI/2;
  data.forEach((item,i) => {
    const sl = (item.value/total)*Math.PI*2;
    ctx.beginPath();
    ctx.moveTo(cx+iR*Math.cos(ang), cy+iR*Math.sin(ang));
    ctx.arc(cx,cy,oR,ang,ang+sl); ctx.arc(cx,cy,iR,ang+sl,ang,true);
    ctx.closePath(); ctx.fillStyle=colors[i%colors.length]; ctx.fill();
    if (item.value/total > 0.06) {
      const mid=ang+sl/2, lr=(oR+iR)/2;
      ctx.fillStyle='#fff'; ctx.font='bold 9px Arial'; ctx.textAlign='center'; ctx.textBaseline='middle';
      ctx.fillText(Math.round(item.value/total*100)+'%', cx+lr*Math.cos(mid), cy+lr*Math.sin(mid));
    }
    ang += sl;
  });
  ctx.fillStyle=getComputedStyle(document.documentElement).getPropertyValue('--text-1').trim()||'#1A1D26';
  ctx.font='bold 15px Arial'; ctx.textAlign='center'; ctx.textBaseline='middle';
  ctx.fillText(nf(total), cx, cy-5);
  ctx.font='9px Arial'; ctx.fillStyle=getComputedStyle(document.documentElement).getPropertyValue('--text-3').trim()||'#9CA3B8';
  ctx.fillText(canvasId==='chart-staff'?'người':'m²', cx, cy+10);
  const el=document.getElementById(legendId);
  if(el) el.innerHTML=data.filter(d=>d.value>0).map((d,i)=>`<div class="cl-item"><span class="cl-dot" style="background:${colors[i]}"></span><span class="cl-label">${d.label}</span><span class="cl-val">${nf(d.value)}</span></div>`).join('');
}

function renderCharts() {
  if (!G) return; const d=G;
  const landData=[{label:'Xây dựng',value:d.dtXD||0},{label:'Sân chơi',value:d.sanChoi||0},{label:'Cây xanh',value:d.cayXanh||0},{label:'Giao thông',value:d.dtGT||0},{label:'Bãi xe',value:d.totalXe||0}];
  const rem=d.dt-landData.reduce((a,d)=>a+d.value,0); if(rem>0)landData.push({label:'Còn lại',value:rem});
  drawDonut('chart-land','legend-land',landData,['#0F766E','#16A34A','#65A30D','#EA580C','#2563EB','#CBD5E1']);
  drawDonut('chart-staff','legend-staff',[{label:'Quản lý',value:d.qlTotal||0},{label:'Giáo viên',value:d.gvTotal||0},{label:'Nhân viên',value:d.nvTotal||0}],['#0F766E','#2563EB','#64748B']);
  drawDonut('chart-floor','legend-floor',[{label:'Phòng CN',value:(d.totalDTPhong-d.dtHL)||0},{label:'Hành lang/CThang',value:d.dtHL||0}],['#2563EB','#CBD5E1']);
}

function getAutoValue(field){
  if(!G) return '–';
  const map={tenDA:G.tenDA,cap:G.std.name,diaDiem:G.diaDiem,dt:nf(G.dt)+' m²',
    dtXD:nf(G.dtXD)+' m²',dtSan:nf(G.dtSan)+' m²',soTang:G.soTang+' tầng',
    matDo:G.matDo+'%',totalHS:nf(G.totalHS)+' học sinh',totalLop:nf(G.totalLop)+' lớp',
    totalCBGV:nf(G.totalCBGV)+' người',loaiDT:G.loai==='db'?'Đặc biệt':G.loai==='nt'?'Nông thôn':'Loại '+G.loai};
  return map[field]||'–';
}
