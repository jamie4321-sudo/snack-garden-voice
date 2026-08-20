// 스낵 큐레이션 제안서 — 웹(카카오페이 클린)과 통일된 PPT 생성
const pptxgen = require('pptxgenjs');
const path = require('path');
const OUT = process.argv[2] || path.join(process.env.USERPROFILE || '.', 'Desktop', '카카오페이_제안서_웹통일_2026.pptx');

const P = new pptxgen();
P.defineLayout({ name: 'W', width: 13.333, height: 7.5 });
P.layout = 'W';
// DEBUG: cap slides via MAX / skip via SKIP (comma list) to bisect corruptor
let __S = 0; const __MAX = +(process.env.MAX || 99); const __SKIP = (process.env.SKIP||'').split(',').filter(Boolean);
const __add = P.addSlide.bind(P);
P.addSlide = function(){ __S++; if(__S>__MAX || __SKIP.includes(String(__S))){ return new Proxy({},{get:()=>()=>{}, set:()=>true}); } return __add(); };
const W = 13.333, H = 7.5, M = 0.62, CW = W - M * 2;

const C = {
  yel:'FEE500', yel6:'F5DA00', y100:'FFF3BF', y50:'FFFBE0',
  ink:'191919', sub:'4B4B4B', ter:'8C8C8C', white:'FFFFFF', paper:'F7F7F4', line:'ECECE8',
  infoBg:'E9F0FF', infoFg:'3768FF', succBg:'E4F8EE', succFg:'0FB267', purpBg:'F3ECFF', purpFg:'7C4DFF', amber:'D8920E'
};
const F = { disp:'Wanted Sans ExtraBold', semi:'Wanted Sans SemiBold', med:'Wanted Sans Medium', reg:'Wanted Sans' };
const shadowFx = () => ({ type:'outer', color:'8A8A8A', opacity:0.28, blur:10, offset:3, angle:90 });
const TINTS = [ {bg:'FFF3BF',fg:'191919'}, {bg:'E9F0FF',fg:'3768FF'}, {bg:'E4F8EE',fg:'0FB267'}, {bg:'F3ECFF',fg:'7C4DFF'} ];

function chrome(s, n, dark){
  const col = dark ? 'B0B0B0' : C.ter;
  s.addText('Linkagelab Confidential', { x:M, y:H-0.5, w:4, h:0.3, fontFace:F.med, fontSize:9, color:col, align:'left' });
  s.addText(String(n).padStart(2,'0'), { x:W-M-0.8, y:H-0.5, w:0.8, h:0.3, fontFace:F.med, fontSize:10, color:col, align:'right' });
}
function eyebrow(s, t, x, y, onYel){
  const bg = onYel ? C.white : C.y100;
  s.addShape(P.ShapeType.roundRect, { x, y, w:0.16+t.length*0.098, h:0.34, rectRadius:0.17, fill:{color:bg}, line:{type:'none'} });
  s.addText(t, { x:x+0.02, y, w:0.12+t.length*0.098, h:0.34, fontFace:F.semi, fontSize:10.5, color:C.ink, align:'center', valign:'middle', charSpacing:0.5 });
}
function title(s, t, y, opt){ opt=opt||{};
  s.addText(t, { x:M, y, w:opt.w||CW, h:opt.h||0.9, fontFace:F.disp, fontSize:opt.size||31, color:opt.color||C.ink, align:'left', valign:'middle', lineSpacingMultiple:1.02 });
}
function lede(s, t, y, opt){ opt=opt||{};
  s.addText(t, { x:M, y, w:opt.w||CW-1, h:opt.h||0.5, fontFace:F.med, fontSize:opt.size||12.5, color:opt.color||C.sub, align:'left', valign:'top', lineSpacingMultiple:1.15 });
}
// rounded card with emoji tile
function card(s, o){
  s.addShape(P.ShapeType.roundRect, { x:o.x, y:o.y, w:o.w, h:o.h, rectRadius:0.14, fill:{color:o.fill||C.white}, line:{color:C.line,width:1}, shadow:shadowFx() });
  const t = o.tint || TINTS[0];
  if(o.emoji!==undefined){
    s.addShape(P.ShapeType.roundRect, { x:o.x+0.28, y:o.y+0.28, w:0.62, h:0.62, rectRadius:0.13, fill:{color:t.bg}, line:{type:'none'} });
    s.addText(o.emoji, { x:o.x+0.28, y:o.y+0.28, w:0.62, h:0.62, fontSize:20, align:'center', valign:'middle' });
  }
  const ty = o.emoji!==undefined ? o.y+1.06 : o.y+0.30;
  s.addText(o.title, { x:o.x+0.30, y:ty, w:o.w-0.6, h:0.4, fontFace:F.disp, fontSize:o.tsize||14.5, color:o.tcolor||C.ink, align:'left', valign:'middle' });
  if(o.desc) s.addText(o.desc, { x:o.x+0.30, y:ty+0.42, w:o.w-0.56, h:o.h-(ty-o.y)-0.5, fontFace:F.med, fontSize:o.dsize||10.5, color:o.dcolor||C.sub, align:'left', valign:'top', lineSpacingMultiple:1.12 });
}
function callout(s, o){ // dark bar with yellow tag
  const bg=o.bg||C.ink, fg=o.fg||C.white;
  s.addShape(P.ShapeType.roundRect, { x:M, y:o.y, w:CW, h:o.h||0.72, rectRadius:0.12, fill:{color:bg}, line:{type:'none'}, shadow:shadowFx() });
  s.addShape(P.ShapeType.roundRect, { x:M+0.28, y:o.y+(((o.h||0.72)-0.34)/2), w:0.16+o.tag.length*0.11, h:0.34, rectRadius:0.17, fill:{color:C.yel}, line:{type:'none'} });
  s.addText(o.tag, { x:M+0.30, y:o.y+(((o.h||0.72)-0.34)/2), w:0.12+o.tag.length*0.11, h:0.34, fontFace:F.semi, fontSize:10.5, color:C.ink, align:'center', valign:'middle' });
  s.addText(o.text, { x:M+0.28+(0.16+o.tag.length*0.11)+0.22, y:o.y, w:CW-(0.16+o.tag.length*0.11)-0.9, h:o.h||0.72, fontFace:F.semi, fontSize:o.size||12.5, color:fg, align:'left', valign:'middle', lineSpacingMultiple:1.1 });
}
function pills(s, arr, y){ // process pills with arrows
  let x=M; const gap=0.16, ah=0.5;
  arr.forEach((p,i)=>{
    const w = 0.5 + p.replace(/\|.*/,'').length*0.13 + (p.includes('|')?0.34:0);
    s.addShape(P.ShapeType.roundRect, { x, y, w, h:ah, rectRadius:0.25, fill:{color:C.white}, line:{color:C.line,width:1}, shadow:shadowFx() });
    const idx = p.includes('|')?p.split('|')[0]:''; const lbl = p.includes('|')?p.split('|')[1]:p;
    if(idx) s.addText([{text:idx+' ',options:{color:'B79E00',bold:true}},{text:lbl,options:{color:C.ink}}], { x:x+0.1, y, w:w-0.2, h:ah, fontFace:F.semi, fontSize:11.5, align:'center', valign:'middle' });
    else s.addText(lbl, { x:x+0.1, y, w:w-0.2, h:ah, fontFace:F.semi, fontSize:11.5, color:C.ink, align:'center', valign:'middle' });
    x += w;
    if(i<arr.length-1){ s.addText('→', { x, y, w:0.34, h:ah, fontFace:F.semi, fontSize:13, color:C.ter, align:'center', valign:'middle' }); x+=0.34; }
    x += gap;
  });
}
function sectionBase(n, eb, ttl, ld, bg){
  const s = P.addSlide(); s.background = { color: bg||C.white };
  eyebrow(s, eb, M, 0.62);
  title(s, ttl, 1.12);
  if(ld) lede(s, ld, 2.02);
  chrome(s, n, false);
  return s;
}

/* ───────── 1 · COVER ───────── */
(() => {
  const s = P.addSlide(); s.background = { color: C.yel };
  // rings
  s.addShape(P.ShapeType.ellipse, { x:9.7, y:4.6, w:5.2, h:5.2, fill:{type:'none'}, line:{color:'FFFFFF',width:34,transparency:74} });
  s.addShape(P.ShapeType.ellipse, { x:-1.2, y:0.5, w:1.9, h:1.9, fill:{type:'none'}, line:{color:'FFFFFF',width:18,transparency:78} });
  eyebrow(s, 'Linkagelab × kakaopay', M, 0.72, true);
  s.addText('SNACK CURATION · 2026', { x:M, y:1.5, w:8, h:0.4, fontFace:F.semi, fontSize:12.5, color:'6A5F1E', charSpacing:1, align:'left' });
  s.addText('매일이 이벤트가\n되는 스낵마켓', { x:M-0.02, y:1.95, w:8.2, h:2.5, fontFace:F.disp, fontSize:57, color:C.ink, align:'left', lineSpacingMultiple:0.98 });
  s.addText('구성원 취향과 시즌감을 담은 큐레이션. 실시간 소통으로 빈틈없이 운영합니다.', { x:M, y:4.55, w:7.2, h:0.7, fontFace:F.med, fontSize:14, color:'2E2A18', align:'left', lineSpacingMultiple:1.15 });
  s.addText('카카오페이 판교 아지트 · 10층~15층', { x:M, y:5.25, w:7, h:0.35, fontFace:F.semi, fontSize:12, color:'6A5F1E', align:'left' });
  // stats card
  const sy=5.95, sw=6.8, sh=1.0; const cells=[['3','운영 지역'],['28','관리 스팟'],['일 2회','정기 진열'],['6+','시즌 이벤트']];
  s.addShape(P.ShapeType.roundRect, { x:M, y:sy, w:sw, h:sh, rectRadius:0.12, fill:{color:C.white}, line:{type:'none'}, shadow:shadowFx() });
  cells.forEach((c,i)=>{ const cx=M+i*(sw/4);
    if(i>0) s.addShape(P.ShapeType.line,{x:cx,y:sy+0.2,w:0,h:sh-0.4,line:{color:C.line,width:1}});
    s.addText(c[0], { x:cx, y:sy+0.16, w:sw/4, h:0.46, fontFace:F.disp, fontSize:23, color:C.ink, align:'center' });
    s.addText(c[1], { x:cx, y:sy+0.62, w:sw/4, h:0.3, fontFace:F.med, fontSize:10.5, color:C.ter, align:'center' });
  });
  // floating live card
  const fx=9.1, fw=3.6, fy=1.7;
  s.addShape(P.ShapeType.roundRect, { x:fx, y:fy, w:fw, h:3.6, rectRadius:0.18, fill:{color:C.white}, line:{type:'none'}, shadow:{type:'outer',color:'6A5A00',opacity:0.35,blur:18,offset:8,angle:90}, rotate:2 });
  s.addText([{text:'🍫 ',options:{}},{text:'실시간 접수 현황',options:{}}], { x:fx+0.3, y:fy+0.28, w:2.4, h:0.34, fontFace:F.disp, fontSize:13, color:C.ink, align:'left', valign:'middle', rotate:2 });
  s.addText([{text:'● ',options:{color:C.succFg}},{text:'LIVE',options:{color:C.succFg}}], { x:fx+fw-1.1, y:fy+0.28, w:0.9, h:0.34, fontFace:F.semi, fontSize:10, align:'right', valign:'middle', rotate:2 });
  const rows=[['조치완료',C.succBg,C.succFg,'11F · 바나나 우유 재고 소진'],['확인중','FFF1D6','C88A00','14F · 냉장고 온도 이상'],['접수',C.y100,'8A7400','10F · 제로 탄산음료 요청']];
  rows.forEach((r,i)=>{ const ry=fy+0.82+i*0.86;
    s.addShape(P.ShapeType.roundRect, { x:fx+0.28, y:ry, w:fw-0.56, h:0.68, rectRadius:0.1, fill:{color:C.paper}, line:{type:'none'}, rotate:2 });
    s.addShape(P.ShapeType.roundRect, { x:fx+0.44, y:ry+0.14, w:0.9, h:0.28, rectRadius:0.14, fill:{color:r[1]}, line:{type:'none'}, rotate:2 });
    s.addText(r[0], { x:fx+0.44, y:ry+0.14, w:0.9, h:0.28, fontFace:F.semi, fontSize:8.5, color:r[2], align:'center', valign:'middle', rotate:2 });
    s.addText(r[3], { x:fx+0.44, y:ry+0.4, w:fw-0.9, h:0.24, fontFace:F.med, fontSize:9.5, color:C.sub, align:'left', valign:'middle', rotate:2 });
  });
})();

/* ───────── 2 · SERVICE ───────── */
(() => {
  const s = sectionBase(2, '01 · SERVICE', '스낵 큐레이션 서비스', '구성원 취향, 위생, 데이터, ESG까지 — 스낵바 운영의 모든 축을 한 팀이 책임집니다.');
  const items=[['✨','맞춤형 큐레이션','구성원 선호도와 계절감을 반영한 시즌별 이벤트·맞춤 간식 구성'],
    ['🧼','위생 관리','정기 클렌징과 설비 점검으로 안정적인 이용 환경 유지'],
    ['📊','데이터 재고관리','발주 데이터·월간 재고 현황 투명 공유 기반 운영 최적화'],
    ['🌱','ESG 친환경 운영','친환경 부자재·장애인 크루 운영으로 ESG 가치 실현']];
  const cw=(CW-0.54)/4, cy=2.9, ch=2.5;
  if(!process.env.NOCARD) items.forEach((it,i)=> card(s,{x:M+i*(cw+0.18),y:cy,w:cw,h:ch,emoji:it[0],title:it[1],desc:it[2],tint:TINTS[i%4]}));
  if(!process.env.NOPILLS) pills(s, ['01|수요 조사','02|소싱 & 큐레이션','03|입고 · 진열','04|재고 관리'], 5.75);
})();

/* ───────── 3 · EVENT ───────── */
(() => {
  const s = sectionBase(3, '02 · EVENT ARCHIVE', '이벤트, 순간들', '단순한 간식이 아니라, 매 분기 새로운 이야기를 만듭니다. 다음 시즌은 카카오페이와 함께 기획합니다.', C.paper);
  const ev=[['🧧','구정 한과','전통 한과·유과 · 새해 인사 카드','25 Q1'],['🍬','추억의 불량식품','쫀드기·아폴로·짱구 · 오브제 20종','25 Q2'],
    ['🥗','저칼로리 간식','그릭요거트·프로틴바 · 200kcal 이하','25 Q3'],['🎄','크리스마스','홀리데이 스낵 박스 · 시즌 데코','25 Q4'],
    ['🤍','화이트데이','수제 마시멜로우 · 미니 캔디 부케','26 Q1'],['♻️','ESG 다회용컵','일회용 저감 · 다회용컵·리유저블백','26 Q2']];
  const cw=(CW-0.36)/3, ch=1.75, gx=0.18, gy=0.18, y0=2.75;
  ev.forEach((e,i)=>{ const r=Math.floor(i/3), c=i%3; const x=M+c*(cw+gx), y=y0+r*(ch+gy);
    card(s,{x,y,w:cw,h:ch,emoji:e[0],title:e[1],desc:e[2],tint:TINTS[i%4]});
    s.addText(e[3], { x:x+cw-1.0, y:y+0.3, w:0.8, h:0.3, fontFace:F.semi, fontSize:10, color:C.ter, align:'right' });
  });
})();

/* ───────── 4 · REAL-TIME ───────── */
(() => {
  const s = sectionBase(4, '03 · REAL-TIME RESPONSE', '빠른 소통, 즉각 대응', 'QR로 접수 → 총무 동시 확인 → 현장 즉각 대응. 실제 웹페이지로 구현했습니다.');
  // left steps
  const steps=[['1','QR 코드 접수','크루가 QR을 스캔해 불편·요청을 남깁니다 (10~15F)'],
    ['2','총무 담당자 동시 확인','접수 즉시 관리자 대시보드에 알림 — 실시간 파악'],
    ['3','현장 즉각 조치','확인부터 조치·완료까지 상태를 투명하게 기록']];
  const sx=M, sw=6.5, sy=2.85, sh=0.92;
  steps.forEach((st,i)=>{ const y=sy+i*(sh+0.16);
    s.addShape(P.ShapeType.roundRect, { x:sx, y, w:sw, h:sh, rectRadius:0.12, fill:{color:C.white}, line:{color:C.line,width:1}, shadow:shadowFx() });
    s.addShape(P.ShapeType.roundRect, { x:sx+0.24, y:y+0.22, w:0.48, h:0.48, rectRadius:0.12, fill:{color:C.yel}, line:{type:'none'} });
    s.addText(st[0], { x:sx+0.24, y:y+0.22, w:0.48, h:0.48, fontFace:F.disp, fontSize:16, color:C.ink, align:'center', valign:'middle' });
    s.addText(st[1], { x:sx+0.9, y:y+0.16, w:sw-1.1, h:0.34, fontFace:F.disp, fontSize:13.5, color:C.ink, align:'left', valign:'middle' });
    s.addText(st[2], { x:sx+0.9, y:y+0.48, w:sw-1.1, h:0.34, fontFace:F.med, fontSize:10.5, color:C.sub, align:'left', valign:'middle' });
  });
  // QR card
  const qy=sy+3*(sh+0.16)+0.02;
  s.addShape(P.ShapeType.roundRect, { x:sx, y:qy, w:sw, h:1.35, rectRadius:0.12, fill:{color:C.y50}, line:{color:C.y100,width:1} });
  s.addImage({ path:'qr.png', x:sx+0.18, y:qy+0.16, w:1.03, h:1.03 });
  s.addText('QR을 스캔하면 실제 접수 화면이 열립니다', { x:sx+1.4, y:qy+0.28, w:sw-1.6, h:0.4, fontFace:F.disp, fontSize:13, color:C.ink, align:'left', valign:'middle' });
  s.addText('jamie4321-sudo.github.io/snack-garden-voice', { x:sx+1.4, y:qy+0.72, w:sw-1.6, h:0.34, fontFace:F.med, fontSize:10, color:C.ter, align:'left', valign:'middle' });
  // right : phone-ish live board mock (dark)
  const px=9.5, pw=3.2, py=2.85, ph=3.9;
  s.addShape(P.ShapeType.roundRect, { x:px, y:py, w:pw, h:ph, rectRadius:0.22, fill:{color:'0E0E10'}, line:{type:'none'}, shadow:shadowFx() });
  s.addText([{text:'● ',options:{color:'C6FF2E'}},{text:'LIVE 접수 대시보드',options:{color:'F4F4F5'}}], { x:px+0.3, y:py+0.28, w:pw-0.6, h:0.34, fontFace:F.semi, fontSize:11.5, align:'left', valign:'middle' });
  const rr=[['조치완료','16241C','5BBE82','11F · 바나나 우유 재고 소진'],['확인중','2A2413','D8A54A','14F · 냉장고 온도 이상'],['접수','242424','9CA3AF','10F · 제로 탄산음료 요청']];
  rr.forEach((r,i)=>{ const ry=py+0.78+i*0.72;
    s.addShape(P.ShapeType.roundRect, { x:px+0.24, y:ry, w:pw-0.48, h:0.6, rectRadius:0.09, fill:{color:'1C1C1F'}, line:{type:'none'} });
    s.addShape(P.ShapeType.roundRect, { x:px+0.4, y:ry+0.12, w:0.85, h:0.26, rectRadius:0.13, fill:{color:r[1]}, line:{type:'none'} });
    s.addText(r[0], { x:px+0.4, y:ry+0.12, w:0.85, h:0.26, fontFace:F.semi, fontSize:8, color:r[2], align:'center', valign:'middle' });
    s.addText(r[3], { x:px+0.4, y:ry+0.36, w:pw-0.7, h:0.22, fontFace:F.med, fontSize:8.5, color:'CFCFCF', align:'left', valign:'middle' });
  });
  s.addText('총무·링키지랩 동시 확인 → 즉시 조치', { x:px+0.24, y:py+ph-0.5, w:pw-0.48, h:0.34, fontFace:F.med, fontSize:9, color:'8A8A8A', align:'center', valign:'middle' });
})();

/* ───────── 5 · DASHBOARD ───────── */
(() => {
  const s = sectionBase(5, '04 · DASHBOARD', '담당자를 위한, 편한 관리', '복잡한 보고 없이, 대시보드 하나로 매입·재고·예산·설비 이슈까지 한눈에.');
  // dark panel
  const px=M, pw=7.1, py=2.85, ph=3.85;
  s.addShape(P.ShapeType.roundRect, { x:px, y:py, w:pw, h:ph, rectRadius:0.16, fill:{color:'161616'}, line:{type:'none'}, shadow:shadowFx() });
  s.addText('📊  스낵바 운영 대시보드', { x:px+0.34, y:py+0.28, w:4, h:0.34, fontFace:F.disp, fontSize:13.5, color:'F4F4F5', align:'left', valign:'middle' });
  s.addText([{text:'● ',options:{color:'C6FF2E'}},{text:'LIVE',options:{color:'C6FF2E'}}], { x:px+pw-1.1, y:py+0.28, w:0.8, h:0.34, fontFace:F.semi, fontSize:10, align:'right', valign:'middle' });
  const kp=[['오늘 총 매입','1,261,300원','▲ 8.2%'],['미조치 이슈','3건','확인중·대기'],['예산 사용률','72%','잔여 10,019,525원']];
  const kw=(pw-0.68-0.2*2)/3;
  kp.forEach((k,i)=>{ const kx=px+0.34+i*(kw+0.2), ky=py+0.78;
    s.addShape(P.ShapeType.roundRect, { x:kx, y:ky, w:kw, h:1.1, rectRadius:0.1, fill:{color:'1F1F1F'}, line:{color:'2C2C2C',width:1} });
    s.addText(k[0], { x:kx+0.16, y:ky+0.14, w:kw-0.3, h:0.26, fontFace:F.med, fontSize:9.5, color:'9A9A9A', align:'left' });
    s.addText(k[1], { x:kx+0.16, y:ky+0.4, w:kw-0.3, h:0.4, fontFace:F.disp, fontSize:16, color:(i===2?'FEE500':'F4F4F5'), align:'left', valign:'middle' });
    s.addText(k[2], { x:kx+0.16, y:ky+0.82, w:kw-0.3, h:0.24, fontFace:F.med, fontSize:8.5, color:(i===0?'5BBE82':'9A9A9A'), align:'left' });
  });
  // budget bar
  const by=py+2.1, bw=pw-0.68;
  s.addText('월 예산 사용 현황', { x:px+0.34, y:by, w:3, h:0.28, fontFace:F.semi, fontSize:10.5, color:'CFCFCF', align:'left' });
  s.addShape(P.ShapeType.roundRect, { x:px+0.34, y:by+0.34, w:bw, h:0.24, rectRadius:0.12, fill:{color:'1F1F1F'}, line:{color:'2C2C2C',width:1} });
  s.addShape(P.ShapeType.roundRect, { x:px+0.34, y:by+0.34, w:bw*0.72, h:0.24, rectRadius:0.12, fill:{color:C.yel}, line:{type:'none'} });
  s.addText('72%', { x:px+pw-1.1, y:by, w:0.8, h:0.28, fontFace:F.disp, fontSize:11, color:'FEE500', align:'right' });
  // inventory
  const inv=[['정상 재고','148','5BBE82'],['부족 재고','12','D8A54A'],['발주 진행중','7','C6FF2E']];
  inv.forEach((v,i)=>{ const ix=px+0.34+i*(kw+0.2), iy=by+0.78;
    s.addShape(P.ShapeType.roundRect, { x:ix, y:iy, w:kw, h:0.6, rectRadius:0.09, fill:{color:'1F1F1F'}, line:{color:'2C2C2C',width:1} });
    s.addText(v[1], { x:ix+0.14, y:iy, w:0.9, h:0.6, fontFace:F.disp, fontSize:17, color:v[2], align:'left', valign:'middle' });
    s.addText(v[0], { x:ix+0.9, y:iy, w:kw-1, h:0.6, fontFace:F.med, fontSize:9.5, color:'B5B5B5', align:'left', valign:'middle' });
  });
  // right features
  const fx=8.05, fw=CW-(fx-M), fy=2.85;
  const feat=[['📊','실시간 대시보드','매입·재고·예산·설비 이슈를 한 화면에서 매일 확인'],
    ['📱','스낵바 넘버스앱','층별·품목별 진열 수량을 앱에서 바로 조회'],
    ['🔄','자동 발주 프로세스','재고 소진 전 선제 보충으로 재고 공백 제로'],
    ['💰','예산 초과 없는 관리','월 사용률·잔여 예산을 실시간으로 조절']];
  const fh=0.9;
  feat.forEach((f,i)=>{ const y=fy+i*(fh+0.13);
    s.addShape(P.ShapeType.roundRect, { x:fx, y, w:fw, h:fh, rectRadius:0.11, fill:{color:C.white}, line:{color:C.line,width:1}, shadow:shadowFx() });
    s.addShape(P.ShapeType.roundRect, { x:fx+0.2, y:y+0.2, w:0.5, h:0.5, rectRadius:0.12, fill:{color:TINTS[i%4].bg}, line:{type:'none'} });
    s.addText(f[0], { x:fx+0.2, y:y+0.2, w:0.5, h:0.5, fontSize:16, align:'center', valign:'middle' });
    s.addText(f[1], { x:fx+0.85, y:y+0.14, w:fw-1, h:0.32, fontFace:F.disp, fontSize:12.5, color:C.ink, align:'left', valign:'middle' });
    s.addText(f[2], { x:fx+0.85, y:y+0.44, w:fw-1.05, h:0.36, fontFace:F.med, fontSize:9.5, color:C.sub, align:'left', valign:'top', lineSpacingMultiple:1.05 });
  });
})();

/* ───────── 6 · TRANSPARENCY ───────── */
(() => {
  const s = sectionBase(6, '05 · TRANSPARENCY', '더 투명하게, 데이터로 운영합니다', '비용 지출·재고 현황·매입 현황까지, 운영 데이터를 기반으로 예산 맞춤형 운영 구조를 설계합니다.', C.paper);
  const items=[['💳','비용 지출','월별 집행 내역을 카테고리·스팟별 리포트로 정기 제공'],
    ['📦','재고 현황','월별 재고 현황 및 주요 소진 품목 공유'],
    ['🧾','매입 현황','원두·부자재 매입 단가·수량·거래처를 투명하게 관리']];
  const cw=(CW-0.44)/3, cy=2.9, ch=2.4;
  items.forEach((it,i)=> card(s,{x:M+i*(cw+0.22),y:cy,w:cw,h:ch,emoji:it[0],title:it[1],desc:it[2],tint:TINTS[i%4]}));
  pills(s, ['Q1|분기 운영 리포트','Q2|위생 점검','Q3|ESG 그린','Q4|고객사 맞춤'], 5.65);
})();

/* ───────── 7 · COST ───────── */
(() => {
  const s = sectionBase(7, '06 · COST PROPOSAL', '운영비 제안', '카카오페이 1~5월 평균 집행 기준 · 부가세 미포함', C.y50);
  const rows=[['신선식품','16,198,711'],['과자','9,497,342'],['음료','7,727,260'],['젤리·초콜렛·바','3,680,540'],['라면·즉석식품','2,477,670'],['냉동식품','2,442,180'],['원물간식','2,257,508'],['마트 소계','44,281,211','sum'],['무료음료','15,312,375'],['고정 인건비','9,424,000'],['총합계','69,017,586','sum'],['(−) 장애인 고용 감면 (월, 추정)','△ 4,860,000','minus'],['실질 부담 (감면 반영·월)','64,157,586원','total']];
  const tx=M, tw=6.9, ty=2.6, rh=0.29;
  s.addShape(P.ShapeType.roundRect, { x:tx, y:ty, w:tw, h:rh*rows.length+0.12, rectRadius:0.1, fill:{color:C.white}, line:{color:C.line,width:1}, shadow:shadowFx() });
  rows.forEach((r,i)=>{ const y=ty+0.06+i*rh;
    let bg=null, kc=C.ink, vc=C.ink, fF=F.med, ff2=F.semi;
    if(r[2]==='sum'){bg='F3F3EF';fF=F.semi;}
    if(r[2]==='minus'){kc='B23B1A';vc='B23B1A';}
    if(r[2]==='total'){bg=C.ink;kc='FFFFFF';vc='FEE500';fF=F.disp;ff2=F.disp;}
    if(bg) s.addShape(P.ShapeType.rect, { x:tx+0.06, y, w:tw-0.12, h:rh, fill:{color:bg}, line:{type:'none'} });
    s.addText(r[0], { x:tx+0.28, y, w:tw-2.4, h:rh, fontFace:fF, fontSize:11, color:kc, align:'left', valign:'middle' });
    s.addText(r[1], { x:tx+tw-2.3, y, w:2.0, h:rh, fontFace:ff2, fontSize:11.5, color:vc, align:'right', valign:'middle' });
  });
  // right points
  const rx=7.8, rw=CW-(rx-M);
  const pts=[['01','카카오 연계 소싱 단가 안정화','품목 구성·용량·발주 구조 최적화로 품질은 유지하며 비용 효율화',false],
    ['02','상주 운영으로 관리 강화','8h 2명(비장애1·중증1) + 4h 중증1 상주 운영',false],
    ['EFFECT','예측 가능한 집행 구조','인건비·운영 실비가 커버되는 구조로 월 지출·재고를 안정적으로 측정·집행',true]];
  let py=2.6; const ph=1.28;
  pts.forEach(p=>{ card(s,{x:rx,y:py,w:rw,h:ph,title:p[1],desc:p[2],fill:(p[3]?C.ink:C.white),tcolor:(p[3]?C.white:C.ink),dcolor:(p[3]?'CFCFCF':C.sub),tsize:13});
    s.addText(p[0], { x:rx+0.3, y:py+0.16, w:rw-0.6, h:0.26, fontFace:F.semi, fontSize:10, color:(p[3]?'FEE500':'B79E00'), align:'left' });
    py+=ph+0.16;
  });
})();

/* ───────── generic 3-card + callout slides (8~15) ───────── */
function threeCardSlide(n, eb, ttl, ld, items, co, bg){
  const s = sectionBase(n, eb, ttl, ld, bg);
  const cw=(CW-0.44)/3, cy=2.55, ch=2.35;
  items.forEach((it,i)=> card(s,{x:M+i*(cw+0.22),y:cy,w:cw,h:ch,emoji:it[0],title:it[1],desc:it[2],tint:TINTS[i%4]}));
  if(co) callout(s, {y:5.35, tag:co[0], text:co[1]});
  return s;
}

/* 8 · EXPECTED EFFECTS */
threeCardSlide(8, '07 · EXPECTED EFFECTS', '서비스 기대 효과', '',
 [['📈','운영 품질 향상','상주 운영 기반 빠른 소통으로 진열·위생·품질 관리 안정화'],
  ['😊','구성원 만족도 향상','맞춤형 큐레이션·안정적 운영으로 이용 만족도·서비스 체감 향상'],
  ['💰','안정적 단가 운영','카카오 운영 연계 구조 기반 소싱 단가·운영 비용 안정화']],
 ['NEW','복지 서비스 통합 시너지 — 총무 · 헬스키퍼 · 카페사업 + 스낵마켓 = 복지 시너지']);

/* 9 · OPERATIONS PLAN */
(() => {
  const s = sectionBase(9, '08 · OPERATIONS PLAN', '운영 기본안 & 논의 사항', '', C.paper);
  const items=[['👤','운영 형태','장애인 크루 상주 / 전담 담당자 배치 · 8h 비장애1 + 8h 중증1 + 4h 중증1'],
    ['🗓️','운영 주기','주 5일 정기 운영 · 재고 보충 일 1~2회(오전·오후)'],
    ['✅','품질 관리','정기 점검·위생 운영 · 신선도 확인 후 스팟별 세팅'],
    ['📦','재고 관리','데이터 기반 재고 최적화 · 재고 소진율 월간 보고']];
  const cw=(CW-0.54)/4, cy=2.5, ch=2.3;
  items.forEach((it,i)=> card(s,{x:M+i*(cw+0.18),y:cy,w:cw,h:ch,emoji:it[0],title:it[1],desc:it[2],tint:TINTS[i%4]}));
  callout(s,{y:5.2,tag:'추가 논의',text:'식자재·오피스 용품 재보충 기준 · 원두 발주 범위 · 캡슐머신 운영 구조 · 리포트 기준',h:1.1,size:12});
})();

/* 10 · PROACTIVE CARE */
threeCardSlide(10, '09 · PROACTIVE CARE', '이슈, 터지기 전에 먼저 공유합니다', '',
 [['①','품절 사전 방지','재고 소진 전 선제 보충으로 진열 공백·품절 원천 차단'],
  ['②','변경사항 사전 공유','미입고·단종·변경 발생 즉시 담당자 소통방에 선제 통보'],
  ['③','원인·조치 투명 기록','이슈 발생부터 조치까지 리포트로 투명하게 공유']],
 ['차별화','사후 확인(현행) → 사전 공유(링키지랩) · 이용자 불편 발생 전에 대응']);

/* 11 · FRESH & CONVENIENCE */
threeCardSlide(11, '10 · FRESH & CONVENIENCE', '신선·간편식 중심 운영 역량', '',
 [['🍱','선호 품목 큐레이션','과일·다이어트·도시락·삼각김밥 선호 라인업 집중 구성'],
  ['🔄','선입선출·신선도 관리','FIFO 원칙·유통기한 관리로 신선식품 품절·폐기 최소화'],
  ['🗓️','정기 진열 운영','K-마트 일 2회 + 라운지 일 1회 정기 진열로 진열 공백 제로']],
 ['진열 기준','일 2회(K-마트) + 1회(라운지) 진열 기준 → 상주 8h 2명 + 4h 1명 동선 설계'], C.paper);

/* 12 · PARTNERSHIP VALUE */
threeCardSlide(12, '11 · PARTNERSHIP VALUE', '링키지랩과의 협업 효과', '',
 [['🧑‍🦽','장애인 일자리 창출','장애 크루 고용 기반 협업 모델로 ESG 사회적 가치 실현'],
  ['🔍','투명한 운영 공유','외주 대비 상세한 운영 자료 공유·예산 집행 투명 협의'],
  ['🤝','공동체 소통·디벨롭','공동체 간 소통으로 방향성·개선점을 함께 발전']],
 ['ESG 효과','장애 크루 4명 고용 기반 협업 모델 · 대외 ESG 보고 · 대내 인식 개선']);

/* 13 · SETTLEMENT */
threeCardSlide(13, '12 · SETTLEMENT', '정산 구조', '',
 [['🧾','매입원가 청구','계약 인원 + 계약 단가 기준으로 매입원가를 투명하게 청구'],
  ['📊','안정적 비용 구조','인건비·운영 실비용이 커버되는 구조로 안정적 운영 가능'],
  ['📅','예측 가능한 집행','월 지출·재고를 안정적으로 측정하고 계획적으로 집행']],
 ['정산 원칙','실비 기반 투명 정산 · 운영(링키지랩)과 납품(소싱) 역할 분리'], C.paper);

/* 14 · WORKFORCE */
threeCardSlide(14, '13 · WORKFORCE', '인력 구성 및 추정 인건비', '',
 [['👤','상주 전담 인력','8h 비장애 1명 상주 배치 · 운영 총괄 및 담당자 소통'],
  ['🧑‍🦽','장애 크루 고용','8h 중증 1명 + 4h 중증 1명 · 진열·정리·위생 업무'],
  ['🛡️','운영 안정성','상주 체계 기반으로 응대·품질·위생 실시간 관리']],
 ['추정 인건비','9,424,000원 / 월 (VAT별도·2026 기준) · 8h 2명(비장애1·중증1) + 4h 중증1']);

/* 15 · ESG IMPACT */
threeCardSlide(15, '14 · ESG IMPACT', '장애인 고용 · 사회적 가치', '',
 [['🧑‍🦽','장애인 일자리','스낵바 운영에 장애 크루 고용 · 지속가능한 일자리 창출'],
  ['📉','고용부담금 절감','장애인 고용 확대에 따른 고용부담금 감면 효과 발생'],
  ['📑','ESG 성과 반영','ESG 보고서·대내 인식개선 등 정성·정량 성과로 활용']],
 ['참고','카카오게임즈 사례 — 장애 크루 4명 고용 시 연 1.45억 감면 효과 · 2024 링키지랩 운영'], C.paper);

/* ───────── 16 · CLOSING ───────── */
(() => {
  const s = P.addSlide(); s.background = { color: C.yel };
  s.addShape(P.ShapeType.ellipse, { x:-1.6, y:-2.6, w:5, h:5, fill:{type:'none'}, line:{color:'FFFFFF',width:34,transparency:74} });
  s.addText('맛있는 하루를\n링키지랩과 함께', { x:0, y:2.1, w:W, h:2.5, fontFace:F.disp, fontSize:52, color:C.ink, align:'center', lineSpacingMultiple:1.0 });
  s.addText('링키지랩 복지큐레이션부 · 스낵&가든사업팀', { x:0, y:4.7, w:W, h:0.4, fontFace:F.semi, fontSize:13, color:'5A5320', align:'center' });
  s.addText('THANK YOU · LINKAGELAB × KAKAOPAY', { x:0, y:5.9, w:W, h:0.4, fontFace:F.semi, fontSize:12, color:'6A5F1E', align:'center', charSpacing:2 });
})();

P.writeFile({ fileName: OUT }).then(f => console.log('SAVED', f)).catch(e => { console.error(e); process.exit(1); });
