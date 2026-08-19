// 카카오 옐로 + 블랙 라운드 모듈 + 중앙 말풍선 로고 QR 생성 (+ 디코딩 자동검증)
const QRCode = require('qrcode');
const sharp = require('sharp');
const { PNG } = require('pngjs');
const jsQR = require('jsqr');
const fs = require('fs');

const URL = process.argv[2] || 'https://jamie4321-sudo.github.io/snack-garden-voice/';
const OUT = process.argv[3] || 'qr.png';
const RF = parseFloat(process.env.RF || '0.5');    // 데이터 모듈 라운드(반지름=RF*셀/2 → 원=1.0)
const LOGO = parseFloat(process.env.LOGO || '0.20'); // 중앙 로고 비율
const SHAPE = process.env.SHAPE || 'dot';           // 'dot' | 'square' | 'soft'

const YELLOW = '#FEE500';
const INK = '#191919';

const qr = QRCode.create(URL, { errorCorrectionLevel: 'H' });
const N = qr.modules.size;
const data = qr.modules.data;
const bit = (r, c) => (r < 0 || c < 0 || r >= N || c >= N) ? 0 : data[r * N + c];

// 파인더 패턴 영역(3개 코너 7x7)인지 판별 → 크리스프하게 통짜로 렌더
function inFinder(r, c) {
  const f = (br, bc) => r >= br && r < br + 7 && c >= bc && c < bc + 7;
  return f(0, 0) || f(0, N - 7) || f(N - 7, 0);
}

const M = 4, U = 40, S = N + M * 2, PX = S * U;
const rData = (U / 2) * Math.max(0, Math.min(1, RF)); // 데이터 모듈 반지름

let cells = '';
for (let r = 0; r < N; r++) {
  for (let c = 0; c < N; c++) {
    if (!bit(r, c)) continue;
    if (inFinder(r, c)) continue; // 파인더는 아래에서 통짜 렌더
    const x = (c + M) * U, y = (r + M) * U;
    if (SHAPE === 'dot') {
      cells += `<circle cx="${x + U/2}" cy="${y + U/2}" r="${rData}"/>`;
    } else if (SHAPE === 'soft') {
      cells += `<rect x="${x}" y="${y}" width="${U}" height="${U}" rx="${U*0.28}" ry="${U*0.28}"/>`;
    } else { // square (edge overlap로 연결성 보장 → 스캔 안정)
      cells += `<rect x="${x-0.5}" y="${y-0.5}" width="${U+1}" height="${U+1}"/>`;
    }
  }
}

// 파인더 패턴: 바깥 7x7 라운드 사각 + 중앙 3x3 라운드 사각 (크리스프)
function finder(br, bc) {
  const x = (bc + M) * U, y = (br + M) * U;
  const oR = U * 1.6, iR = U * 0.9;
  return `
    <rect x="${x}" y="${y}" width="${U*7}" height="${U*7}" rx="${oR}" ry="${oR}" fill="${INK}"/>
    <rect x="${x+U}" y="${y+U}" width="${U*5}" height="${U*5}" rx="${U*1.1}" ry="${U*1.1}" fill="${YELLOW}"/>
    <rect x="${x+U*2}" y="${y+U*2}" width="${U*3}" height="${U*3}" rx="${iR}" ry="${iR}" fill="${INK}"/>`;
}
const finders = finder(0, 0) + finder(0, N - 7) + finder(N - 7, 0);

// 중앙 로고
const logoModules = Math.round(N * LOGO);
const lStart = Math.floor((N - logoModules) / 2) + M;
const lx = lStart * U, ly = lStart * U, lw = logoModules * U;
const badgeR = lw * 0.26;
const pad = lw * 0.16;
const bScale = (lw - pad * 2) / 120;
const bx = lx + pad, by = ly + pad + lw * 0.02;

const svg = `<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="${PX}" height="${PX}" viewBox="0 0 ${PX} ${PX}">
  <rect width="${PX}" height="${PX}" rx="${U*1.6}" ry="${U*1.6}" fill="${YELLOW}"/>
  <g fill="${INK}">${cells}</g>
  ${finders}
  <rect x="${lx}" y="${ly}" width="${lw}" height="${lw}" rx="${badgeR}" ry="${badgeR}" fill="${YELLOW}"/>
  <rect x="${lx+U*0.18}" y="${ly+U*0.18}" width="${lw-U*0.36}" height="${lw-U*0.36}" rx="${badgeR*0.9}" ry="${badgeR*0.9}" fill="#ffffff"/>
  <g transform="translate(${bx} ${by}) scale(${bScale})">
    <path d="M60 6 C26 6 8 25 8 48 C8 64 18 78 35 85 C33 94 28 100 21 104 C35 103 49 98 58 89 C58.7 89 59.3 89 60 89 C94 89 112 69 112 48 C112 25 94 6 60 6 Z" fill="${INK}"/>
  </g>
</svg>`;

fs.writeFileSync(OUT.replace(/\.png$/, '.svg'), svg);
sharp(Buffer.from(svg)).png().toFile(OUT).then(() => {
  const png = PNG.sync.read(fs.readFileSync(OUT));
  const code = jsQR(new Uint8ClampedArray(png.data), png.width, png.height);
  console.log('OK', OUT, PX + 'px', 'modules=' + N, 'RF=' + RF, 'LOGO=' + LOGO,
    '| decode:', code ? 'PASS → ' + code.data : 'FAIL');
}).catch(e => { console.error(e); process.exit(1); });
