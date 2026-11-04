// 博丽灵梦 (Hakurei Reimu) — fan illustration drawn entirely in code.
//
// Every shape below is a hand-written SVG path (Bezier curves, gradients and
// filters); no bitmap assets are used. Running this script writes
// `reimu.svg` and rasterises it to `reimu.png` with headless Chromium.
//
//   node art/reimu/reimu.mjs
//
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const here = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(import.meta.url);

const W = 1000;
const H = 1400;

// ---------------------------------------------------------------------------
// helpers
// ---------------------------------------------------------------------------
function mulberry32(a) {
  return function () {
    a |= 0; a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const rand = mulberry32(20261104);
const rr = (a, b) => a + (b - a) * rand();
const f = (n) => Number(n.toFixed(2));

const mirror = (svg) => `<g transform="translate(${W} 0) scale(-1 1)">${svg}</g>`;

const C = {
  hair: '#5b3224', hairDark: '#3b1c14', hairDeep: '#2a120c', hairLight: '#9a5e46', hairLine: '#221009',
  skin: '#ffece2', skinShadow: '#f7c7b3', skinDeep: '#eba792', skinLine: '#b46a58',
  red: '#d9232d', redShadow: '#a5121d', redDeep: '#760a13', redLight: '#ff6a5e', redLine: '#4a070b',
  white: '#ffffff', whiteShadow: '#e1e5f4', whiteDeep: '#bfc6e2', whiteLine: '#6c7397',
  yellow: '#ffd84f', yellowShadow: '#eba72c', yellowLine: '#8a5310',
  lash: '#2b0f0d',
};

// A ruffled lace edge: scallops + pleat ticks centred on path `d`.
// Whatever is painted on top afterwards hides the inner half.
function frill(d, { scal = 32, gap = 19, band = 22, line = C.whiteLine, fill = C.white, tick = C.whiteDeep } = {}) {
  return `
  <path d="${d}" fill="none" stroke="${line}" stroke-width="${scal + 4.5}" stroke-dasharray="0.01 ${gap}" stroke-linecap="round"/>
  <path d="${d}" fill="none" stroke="${fill}" stroke-width="${scal}" stroke-dasharray="0.01 ${gap}" stroke-linecap="round"/>
  <path d="${d}" fill="none" stroke="${fill}" stroke-width="${band}" stroke-linejoin="round"/>
  <path d="${d}" fill="none" stroke="${tick}" stroke-width="${scal - 8}" stroke-dasharray="1.4 ${gap - 1.4}" stroke-dashoffset="${f(gap / 2 + 0.7)}" opacity=".9"/>`;
}

// tapered wedge from (x1,y1) [wide end width w1] to (x2,y2) [width w2], bowed by `bend`
function wedge(x1, y1, x2, y2, w1, w2 = 0, bend = 0) {
  const dx = x2 - x1, dy = y2 - y1, L = Math.hypot(dx, dy);
  const nx = -dy / L, ny = dx / L;
  const mx = (x1 + x2) / 2 + nx * bend, my = (y1 + y2) / 2 + ny * bend;
  const a = [x1 + nx * w1 / 2, y1 + ny * w1 / 2];
  const b = [x1 - nx * w1 / 2, y1 - ny * w1 / 2];
  const c = [x2 + nx * w2 / 2, y2 + ny * w2 / 2];
  const d = [x2 - nx * w2 / 2, y2 - ny * w2 / 2];
  return `M ${f(a[0])} ${f(a[1])} Q ${f(mx + nx * (w1 + w2) / 4)} ${f(my + ny * (w1 + w2) / 4)} ${f(c[0])} ${f(c[1])} ` +
    `L ${f(d[0])} ${f(d[1])} Q ${f(mx - nx * (w1 + w2) / 4)} ${f(my - ny * (w1 + w2) / 4)} ${f(b[0])} ${f(b[1])} Z`;
}

// ---------------------------------------------------------------------------
// shared geometry
// ---------------------------------------------------------------------------
const FACE = 'M 357 430 C 354 540, 369 606, 410 653 C 441 689, 474 709, 500 713 C 526 709, 559 689, 590 653 C 631 606, 646 540, 643 430 Z';
const NECK = 'M 468 640 L 467 700 C 466 738, 463 762, 457 792 L 543 792 C 537 762, 534 738, 533 700 L 532 640 Z';
const BODY_SKIN = 'M 466 742 L 534 742 C 548 785, 610 806, 668 820 C 708 831, 724 868, 724 915 L 728 1010 L 636 1010 L 636 890 L 364 890 L 364 1010 L 272 1010 L 276 915 C 276 868, 292 831, 332 820 C 390 806, 452 785, 466 742 Z';
const TOP_EDGE = 'M 360 892 C 392 852, 428 818, 453 786 C 472 799, 528 799, 547 786 C 572 818, 608 852, 640 892';
const TOP = `${TOP_EDGE} C 648 965, 634 1062, 616 1150 L 384 1150 C 366 1062, 352 965, 360 892 Z`;
const SKIRT = 'M 384 1146 C 360 1240, 330 1330, 296 1430 L 704 1430 C 670 1330, 640 1240, 616 1146 Z';

const CAP_TOP = 'M 326 494 C 298 384, 350 236, 500 230 C 650 236, 702 384, 674 494';
// fringe, right -> left; every strand sweeps a little toward the viewer's left
const BANGS =
  ' Q 678 522 656 548 Q 648 502 636 466' +
  ' Q 628 500 610 522 Q 604 486 590 458' +
  ' Q 586 502 565 532 Q 558 484 546 456' +
  ' Q 542 502 521 548 Q 514 496 506 462' +
  ' Q 500 508 485 558 Q 476 500 466 458' +
  ' Q 460 496 443 524 Q 434 484 422 458' +
  ' Q 414 494 395 522 Q 386 486 374 462' +
  ' Q 366 502 346 542 Q 328 518 326 494 Z';
const CAP = CAP_TOP + BANGS;
const BANG_NOTCHES = [[636, 466], [590, 458], [546, 456], [506, 462], [466, 458], [422, 458], [374, 462]];

const SIDE_HAIR = 'M 338 440 C 312 520, 306 650, 326 766 C 340 712, 350 664, 362 622 C 368 662, 380 694, 398 716 C 390 652, 386 572, 398 476 Z';
const SIDELOCK_TOP = 'M 333 540 C 325 640, 318 720, 314 792 L 362 794 C 364 720, 370 640, 377 560 Z';
const SIDELOCK_END = 'M 316 878 C 298 922, 290 970, 299 1022 C 307 992, 315 975, 326 962 C 329 988, 337 1006, 352 1026 C 356 984, 364 932, 361 878 Z';

const BOW_LOOP = 'M 492 300 C 452 250, 382 160, 300 150 C 244 145, 220 186, 228 236 C 234 288, 262 332, 302 347 C 372 366, 442 332, 492 312 Z';

const BACK_HAIR = 'M 500 230 C 356 228, 278 330, 278 472 C 276 560, 270 640, 256 720 Q 246 744, 230 762 Q 254 754, 262 744 ' +
  'C 258 800, 248 860, 224 914 Q 256 898, 272 874 Q 268 940, 290 998 Q 304 946, 320 914 Q 330 974, 354 1014 Q 354 952, 364 900 ' +
  'L 636 900 Q 646 952, 646 1014 Q 670 974, 680 914 Q 696 946, 710 998 Q 732 940, 728 874 Q 744 898, 776 914 ' +
  'C 752 860, 742 800, 738 744 Q 746 754, 770 762 Q 754 744, 744 720 C 730 640, 724 560, 722 472 C 722 330, 644 228, 500 230 Z';

const TILT = 'rotate(-4 500 720)';

// gohei (purification rod) geometry
const FIST = { x: 712, y: 1336 };
const STICK_ANGLE = 13.1; // degrees from vertical, leaning right
const tan = Math.tan((STICK_ANGLE * Math.PI) / 180);
const STICK_TOP = { x: FIST.x + (FIST.y - 596) * tan, y: 596 };
const STICK_BOT = { x: FIST.x - (1440 - FIST.y) * tan, y: 1440 };

// ---------------------------------------------------------------------------
// defs
// ---------------------------------------------------------------------------
function defs() {
  return `<defs>
  <linearGradient id="sky" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0" stop-color="#5fa8e6"/>
    <stop offset=".45" stop-color="#9fd0f2"/>
    <stop offset=".75" stop-color="#e6f1fb"/>
    <stop offset="1" stop-color="#fde3ee"/>
  </linearGradient>
  <radialGradient id="sun" cx="0.5" cy="0.5" r="0.5">
    <stop offset="0" stop-color="#fffbe8" stop-opacity="1"/>
    <stop offset=".35" stop-color="#fff4d6" stop-opacity=".55"/>
    <stop offset="1" stop-color="#fff4d6" stop-opacity="0"/>
  </radialGradient>
  <radialGradient id="aura" cx="0.5" cy="0.5" r="0.5">
    <stop offset="0" stop-color="#ffffff" stop-opacity=".75"/>
    <stop offset=".6" stop-color="#fff6fb" stop-opacity=".25"/>
    <stop offset="1" stop-color="#ffffff" stop-opacity="0"/>
  </radialGradient>
  <linearGradient id="hairG" gradientUnits="userSpaceOnUse" x1="0" y1="220" x2="0" y2="560">
    <stop offset="0" stop-color="#6d3b2a"/>
    <stop offset=".55" stop-color="${C.hair}"/>
    <stop offset="1" stop-color="#4e2a1e"/>
  </linearGradient>
  <linearGradient id="backHairG" gradientUnits="userSpaceOnUse" x1="0" y1="250" x2="0" y2="1000">
    <stop offset="0" stop-color="#4a271c"/>
    <stop offset="1" stop-color="#2e150e"/>
  </linearGradient>
  <linearGradient id="lockG" gradientUnits="userSpaceOnUse" x1="0" y1="540" x2="0" y2="1030">
    <stop offset="0" stop-color="#4f2a1e"/>
    <stop offset=".5" stop-color="${C.hair}"/>
    <stop offset="1" stop-color="#6b3b2b"/>
  </linearGradient>
  <linearGradient id="skinG" gradientUnits="userSpaceOnUse" x1="0" y1="430" x2="0" y2="720">
    <stop offset="0" stop-color="#ffe7db"/>
    <stop offset="1" stop-color="${C.skin}"/>
  </linearGradient>
  <linearGradient id="redG" gradientUnits="userSpaceOnUse" x1="0" y1="780" x2="0" y2="1430">
    <stop offset="0" stop-color="#e62e35"/>
    <stop offset="1" stop-color="#c41a26"/>
  </linearGradient>
  <linearGradient id="bowG" gradientUnits="userSpaceOnUse" x1="0" y1="150" x2="0" y2="360">
    <stop offset="0" stop-color="#f03a3c"/>
    <stop offset="1" stop-color="#c2182a"/>
  </linearGradient>
  <linearGradient id="sleeveG" x1="0" y1="0" x2="1" y2="0">
    <stop offset="0" stop-color="#eef1fb"/>
    <stop offset=".45" stop-color="#ffffff"/>
    <stop offset="1" stop-color="#e3e7f6"/>
  </linearGradient>
  <linearGradient id="irisG" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0" stop-color="#3a0709"/>
    <stop offset=".38" stop-color="#8d1a1f"/>
    <stop offset=".72" stop-color="#cf3f31"/>
    <stop offset="1" stop-color="#ff9a62"/>
  </linearGradient>
  <linearGradient id="scleraG" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0" stop-color="#c7cbe3"/>
    <stop offset=".45" stop-color="#ffffff"/>
  </linearGradient>
  <linearGradient id="woodG" x1="0" y1="0" x2="1" y2="0">
    <stop offset="0" stop-color="#b9864e"/>
    <stop offset=".45" stop-color="#f0cf96"/>
    <stop offset="1" stop-color="#a4733f"/>
  </linearGradient>
  <linearGradient id="tubeG" x1="0" y1="0" x2="1" y2="0">
    <stop offset="0" stop-color="#d4d9ee"/>
    <stop offset=".35" stop-color="#ffffff"/>
    <stop offset="1" stop-color="#c3cae4"/>
  </linearGradient>
  <linearGradient id="yellowG" x1="0" y1="0" x2="0" y2="1">
    <stop offset="0" stop-color="#ffe682"/>
    <stop offset="1" stop-color="${C.yellow}"/>
  </linearGradient>
  <radialGradient id="orbShade" cx=".38" cy=".34" r=".75">
    <stop offset="0" stop-color="#fff" stop-opacity="0"/>
    <stop offset=".6" stop-color="#2b1630" stop-opacity=".08"/>
    <stop offset="1" stop-color="#2b1630" stop-opacity=".42"/>
  </radialGradient>
  <radialGradient id="orbGlow" cx=".5" cy=".5" r=".5">
    <stop offset="0" stop-color="#ffd6e2" stop-opacity=".9"/>
    <stop offset=".5" stop-color="#ffb3c9" stop-opacity=".35"/>
    <stop offset="1" stop-color="#ffb3c9" stop-opacity="0"/>
  </radialGradient>
  <radialGradient id="petalG" cx=".5" cy=".9" r=".9">
    <stop offset="0" stop-color="#ff8fb3"/>
    <stop offset=".45" stop-color="#ffc6d8"/>
    <stop offset="1" stop-color="#fff0f5"/>
  </radialGradient>
  <radialGradient id="vignette" cx=".5" cy=".45" r=".75">
    <stop offset=".6" stop-color="#3a2050" stop-opacity="0"/>
    <stop offset="1" stop-color="#3a2050" stop-opacity=".28"/>
  </radialGradient>
  <linearGradient id="lightWash" x1="1" y1="0" x2="0" y2="1">
    <stop offset="0" stop-color="#fff6d8" stop-opacity=".55"/>
    <stop offset=".5" stop-color="#fff6d8" stop-opacity="0"/>
  </linearGradient>
  <filter id="b1" x="-20%" y="-20%" width="140%" height="140%"><feGaussianBlur stdDeviation="1.2"/></filter>
  <filter id="b3" x="-30%" y="-30%" width="160%" height="160%"><feGaussianBlur stdDeviation="3"/></filter>
  <filter id="b6" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="6"/></filter>
  <filter id="b12" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="12"/></filter>
  <filter id="b25" x="-50%" y="-50%" width="200%" height="200%"><feGaussianBlur stdDeviation="25"/></filter>
  <clipPath id="clipFace"><path d="${FACE}"/></clipPath>
  <clipPath id="clipNeck"><path d="${NECK}"/></clipPath>
  <clipPath id="clipBody"><path d="${BODY_SKIN}"/></clipPath>
  <clipPath id="clipTop"><path d="${TOP}"/></clipPath>
  <clipPath id="clipSkirt"><path d="${SKIRT}"/></clipPath>
  <clipPath id="clipCap"><path d="${CAP}"/></clipPath>
  <clipPath id="clipBack"><path d="${BACK_HAIR}"/></clipPath>
  <clipPath id="clipSide"><path d="${SIDE_HAIR}"/></clipPath>
  <clipPath id="clipBow"><path d="${BOW_LOOP}"/></clipPath>
  </defs>`;
}

// ---------------------------------------------------------------------------
// background: spring sky, shrine torii, distant sakura
// ---------------------------------------------------------------------------
function cloud(x, y, s, op = 0.9) {
  const puffs = [[0, 0, 60], [55, -18, 48], [105, 4, 52], [-50, 10, 40], [40, 22, 50], [150, 18, 34]];
  return `<g opacity="${op}" filter="url(#b3)">${puffs
    .map(([dx, dy, r]) => `<circle cx="${f(x + dx * s)}" cy="${f(y + dy * s)}" r="${f(r * s)}" fill="#fff"/>`)
    .join('')}</g>`;
}

function background() {
  let s = `<rect width="${W}" height="${H}" fill="url(#sky)"/>`;
  s += `<circle cx="780" cy="230" r="460" fill="url(#sun)"/>`;
  s += cloud(70, 420, 1.1, 0.85) + cloud(760, 520, 0.9, 0.8) + cloud(-40, 700, 1.3, 0.7) + cloud(820, 760, 1.2, 0.75);

  // distant hills
  s += `<path d="M0 930 C 120 880, 230 862, 340 892 C 450 922, 540 850, 660 842 C 780 834, 880 884, 1000 862 L1000 1400 L0 1400Z" fill="#a9c9e6" opacity=".85"/>`;
  s += `<path d="M0 990 C 150 950, 260 960, 380 985 C 520 1012, 640 952, 780 958 C 880 962, 940 985, 1000 975 L1000 1400 L0 1400Z" fill="#c5dcef"/>`;
  // sakura trees in the haze
  let blobs = '';
  for (let i = 0; i < 46; i++) {
    const x = rr(-20, 1020), y = rr(950, 1060), r = rr(22, 50);
    const col = ['#f8c9d8', '#f4b6ca', '#fbd9e4', '#efaec4'][i % 4];
    blobs += `<circle cx="${f(x)}" cy="${f(y)}" r="${f(r)}" fill="${col}"/>`;
  }
  s += `<g filter="url(#b6)" opacity=".9">${blobs}</g>`;
  s += `<rect x="0" y="1040" width="${W}" height="400" fill="#f2d3df"/>`;

  // torii gate (slightly hazy: it stands behind Reimu)
  const red = '#e45a43', redS = '#b63b2e';
  let t = '';
  // pillars
  t += `<path d="M 168 120 L 212 120 L 222 1440 L 158 1440 Z" fill="${red}"/>`;
  t += `<path d="M 168 120 L 182 120 L 176 1440 L 158 1440 Z" fill="${redS}" opacity=".55"/>`;
  t += `<path d="M 788 120 L 832 120 L 842 1440 L 778 1440 Z" fill="${red}"/>`;
  t += `<path d="M 818 120 L 832 120 L 842 1440 L 826 1440 Z" fill="${redS}" opacity=".55"/>`;
  // nuki (lower beam)
  t += `<rect x="96" y="238" width="808" height="30" fill="${red}"/>`;
  t += `<rect x="96" y="260" width="808" height="8" fill="${redS}" opacity=".6"/>`;
  // gakuzuka (centre strut)
  t += `<rect x="484" y="128" width="32" height="112" fill="${red}"/>`;
  // shimaki + kasagi (top lintel, curving up at the ends)
  t += `<path d="M 40 96 Q 500 140 960 96 L 955 126 Q 500 164 45 126 Z" fill="${red}"/>`;
  t += `<path d="M 45 118 Q 500 156 955 118 L 955 126 Q 500 164 45 126 Z" fill="${redS}" opacity=".6"/>`;
  t += `<path d="M -10 58 Q 500 122 1010 58 L 996 98 Q 500 146 4 98 Z" fill="#2c2026"/>`;
  t += `<path d="M -10 58 Q 500 122 1010 58 L 1008 66 Q 500 128 -8 66 Z" fill="#5b4a55"/>`;
  s += `<g opacity=".78" filter="url(#b1)">${t}</g>`;
  // atmospheric wash so the gate sits back
  s += `<rect width="${W}" height="${H}" fill="#eaf4fc" opacity=".18"/>`;
  return s;
}

// ---------------------------------------------------------------------------
// cherry blossoms
// ---------------------------------------------------------------------------
function blossom(cx, cy, s, rot, op = 1) {
  const petal = `M 0 0 C ${-s * 0.5} ${-s * 0.22}, ${-s * 0.58} ${-s * 0.86}, ${-s * 0.17} ${-s} L 0 ${-s * 0.86} L ${s * 0.17} ${-s} C ${s * 0.58} ${-s * 0.86}, ${s * 0.5} ${-s * 0.22}, 0 0 Z`;
  let g = '';
  for (let i = 0; i < 5; i++) {
    g += `<path d="${petal}" transform="rotate(${i * 72})" fill="url(#petalG)" stroke="#e98aa8" stroke-width="1.1"/>`;
  }
  g += `<circle r="${f(s * 0.22)}" fill="#ff7aa2" opacity=".85"/>`;
  for (let i = 0; i < 7; i++) {
    const a = (i / 7) * Math.PI * 2 + 0.3, l = s * 0.42;
    g += `<line x1="0" y1="0" x2="${f(Math.cos(a) * l)}" y2="${f(Math.sin(a) * l)}" stroke="#d9466f" stroke-width="1"/>` +
      `<circle cx="${f(Math.cos(a) * l)}" cy="${f(Math.sin(a) * l)}" r="1.8" fill="#ffd36b"/>`;
  }
  return `<g transform="translate(${f(cx)} ${f(cy)}) rotate(${f(rot)})" opacity="${op}">${g}</g>`;
}

function petal(cx, cy, s, rot, op, blur) {
  const d = `M 0 ${-s} C ${s * 0.85} ${-s * 0.55}, ${s * 0.6} ${s * 0.75}, 0 ${s} C ${-s * 0.6} ${s * 0.75}, ${-s * 0.85} ${-s * 0.55}, 0 ${-s} Z`;
  return `<g transform="translate(${f(cx)} ${f(cy)}) rotate(${f(rot)}) scale(1 ${f(rr(0.55, 1))})" opacity="${f(op)}" ${blur ? `filter="url(#${blur})"` : ''}>` +
    `<path d="${d}" fill="#ffc3d5"/><path d="M 0 ${-s} L ${-s * 0.18} ${-s * 0.62} L 0 ${-s * 0.72} L ${s * 0.18} ${-s * 0.62} Z" fill="#ffe8f0"/>` +
    `<path d="M 0 ${s} C ${s * 0.3} ${s * 0.4}, ${s * 0.3} ${-s * 0.2}, 0 ${-s * 0.5}" stroke="#f593b2" stroke-width="1" fill="none"/></g>`;
}

function branches() {
  let s = '';
  const bark = '#4f3530';
  // top-left branch
  s += `<path d="M -40 40 C 60 70, 150 70, 250 38 C 300 22, 330 0, 345 -20" fill="none" stroke="${bark}" stroke-width="16" stroke-linecap="round"/>`;
  s += `<path d="M 120 64 C 150 100, 170 130, 214 150" fill="none" stroke="${bark}" stroke-width="8" stroke-linecap="round"/>`;
  s += `<path d="M 230 44 C 262 70, 280 84, 312 88" fill="none" stroke="${bark}" stroke-width="6" stroke-linecap="round"/>`;
  s += `<path d="M -40 40 C 60 70, 150 70, 250 38" fill="none" stroke="#7a5650" stroke-width="4" stroke-linecap="round" opacity=".6" transform="translate(0 -5)"/>`;
  // top-right branch
  s += `<path d="M 1040 190 C 960 170, 900 120, 850 60 C 830 36, 815 10, 810 -20" fill="none" stroke="${bark}" stroke-width="15" stroke-linecap="round"/>`;
  s += `<path d="M 905 130 C 880 150, 860 175, 820 186" fill="none" stroke="${bark}" stroke-width="7" stroke-linecap="round"/>`;
  s += `<path d="M 1040 190 C 960 170, 900 120, 850 60" fill="none" stroke="#7a5650" stroke-width="4" stroke-linecap="round" opacity=".6" transform="translate(0 -5)"/>`;

  const spots = [
    [20, 58, 26], [70, 80, 22], [112, 52, 28], [160, 82, 24], [205, 40, 26], [214, 150, 24], [180, 128, 20],
    [262, 18, 22], [300, 92, 24], [326, 58, 18], [58, 30, 18], [140, 22, 16],
    [990, 168, 26], [946, 136, 24], [930, 182, 20], [880, 96, 26], [846, 46, 22], [820, 186, 24], [862, 160, 20],
    [970, 214, 18], [820, 10, 18], [905, 60, 18],
  ];
  for (const [x, y, sz] of spots) s += blossom(x, y, sz, rr(0, 72));
  // a few buds/leaves
  for (const [x, y] of [[96, 100], [240, 70], [960, 120], [870, 200]]) {
    s += `<ellipse cx="${x}" cy="${y}" rx="8" ry="5" fill="#ff9cbc" transform="rotate(${f(rr(0, 180))} ${x} ${y})"/>`;
  }
  return s;
}

// ---------------------------------------------------------------------------
// character
// ---------------------------------------------------------------------------
function backHair() {
  let s = `<path d="${BACK_HAIR}" fill="url(#backHairG)" stroke="${C.hairLine}" stroke-width="3" stroke-linejoin="round"/>`;
  s += `<g clip-path="url(#clipBack)">`;
  // soft sheen on the outer volume, then darker strand separations
  for (const [x1, y1, x2, y2, w, b] of [[300, 400, 252, 880, 16, 14], [700, 400, 748, 880, 16, -14]]) {
    s += `<path d="${wedge(x1, y1, x2, y2, w, 0, b)}" fill="#6f3e2d" opacity=".55"/>`;
  }
  for (const [x1, y1, x2, y2, w, b] of [[316, 520, 268, 900, 7, 8], [338, 600, 304, 990, 7, -6], [372, 650, 352, 1010, 6, 4],
    [684, 520, 732, 900, 7, -8], [662, 600, 696, 990, 7, 6], [628, 650, 648, 1010, 6, -4]]) {
    s += `<path d="${wedge(x1, y1, x2, y2, w, 0, b)}" fill="${C.hairDeep}" opacity=".75"/>`;
  }
  s += `</g>`;
  s += `<mask id="rimBack"><path d="${BACK_HAIR}" fill="#fff"/><path d="${BACK_HAIR}" transform="translate(-8 4)" fill="#000"/></mask>`;
  s += `<rect x="200" y="200" width="620" height="860" fill="#ffd6bf" opacity=".6" mask="url(#rimBack)"/>`;
  const fly = (d) => `<path d="${d}" fill="${C.hair}" stroke="${C.hairLine}" stroke-width="1.8" stroke-linejoin="round"/>`;
  s += fly('M 716 420 Q 752 436 776 476 Q 748 452 712 438 Z');
  s += fly('M 722 560 Q 760 590 772 640 Q 748 600 720 580 Z');
  s += fly('M 284 430 Q 250 450 232 496 Q 258 466 288 450 Z');
  return s;
}

function bow() {
  const loop = (flip) => {
    let s = '';
    s += frill(BOW_LOOP, { scal: 34, gap: 20, band: 24 });
    s += `<path d="${BOW_LOOP}" fill="url(#bowG)" stroke="${C.redLine}" stroke-width="3" stroke-linejoin="round"/>`;
    s += `<g clip-path="url(#clipBow)">`;
    // inner folds
    s += `<path d="M 492 312 C 430 300, 360 300, 302 347 C 340 346, 420 330, 492 312 Z" fill="${C.redShadow}"/>`;
    s += `<path d="M 300 150 C 250 150, 232 190, 238 236 C 244 280, 268 318, 300 338 C 278 296, 270 220, 300 150 Z" fill="${C.redShadow}" opacity=".85"/>`;
    s += `<path d="${wedge(486, 300, 330, 196, 14, 2, -10)}" fill="${C.redShadow}" opacity=".9"/>`;
    s += `<path d="${wedge(486, 306, 316, 268, 10, 1, 6)}" fill="${C.redShadow}" opacity=".8"/>`;
    s += `<path d="M 300 160 C 360 172, 420 228, 470 280" stroke="${C.redLight}" stroke-width="6" fill="none" stroke-linecap="round" opacity=".55"/>`;
    s += `<path d="M 262 178 C 252 200, 250 224, 254 246" stroke="${C.redLight}" stroke-width="4" fill="none" stroke-linecap="round" opacity=".5"/>`;
    s += `</g>`;
    return flip ? mirror(s) : s;
  };
  // tails hanging behind the head
  let tails = '';
  tails += `<path d="M 470 300 C 440 360, 410 420, 392 470 L 420 486 L 432 460 C 452 410, 478 360, 492 312 Z" fill="${C.redShadow}" stroke="${C.redLine}" stroke-width="3" stroke-linejoin="round"/>`;
  tails += mirror(`<path d="M 470 300 C 440 360, 410 420, 392 470 L 420 486 L 432 460 C 452 410, 478 360, 492 312 Z" fill="${C.redShadow}" stroke="${C.redLine}" stroke-width="3" stroke-linejoin="round"/>`);
  const knot = `<path d="M 462 262 C 470 240, 530 240, 538 262 C 542 290, 534 318, 500 320 C 466 318, 458 290, 462 262 Z" fill="${C.red}" stroke="${C.redLine}" stroke-width="3"/>`;
  return `<g transform="rotate(-5 500 290)">${tails}${loop(false)}${loop(true)}${knot}</g>`;
}

function body() {
  let s = '';
  // skin: neck, shoulders, upper arms
  s += `<path d="${BODY_SKIN}" fill="${C.skin}" stroke="${C.skinLine}" stroke-width="2.6" stroke-linejoin="round"/>`;
  s += `<g clip-path="url(#clipBody)">`;
  // hair shadow falling on the shoulders + armpit shading
  s += `<path d="M 300 830 C 340 812, 400 806, 440 796 L 470 760 L 470 860 L 300 900 Z" fill="${C.skinShadow}" opacity=".85" filter="url(#b1)"/>`;
  s += `<path d="M 700 830 C 660 812, 600 806, 560 796 L 530 760 L 530 860 L 700 900 Z" fill="${C.skinShadow}" opacity=".85" filter="url(#b1)"/>`;
  s += `<path d="M 364 890 C 356 930, 360 970, 368 1010 L 340 1010 C 340 960, 345 920, 364 890 Z" fill="${C.skinDeep}" opacity=".6"/>`;
  s += `<path d="M 636 890 C 644 930, 640 970, 632 1010 L 660 1010 C 660 960, 655 920, 636 890 Z" fill="${C.skinDeep}" opacity=".6"/>`;
  // collarbone hints
  s += `<path d="M 444 806 Q 410 812 382 816" stroke="${C.skinLine}" stroke-width="1.8" fill="none" opacity=".45" stroke-linecap="round"/>`;
  s += `<path d="M 556 806 Q 590 812 618 816" stroke="${C.skinLine}" stroke-width="1.8" fill="none" opacity=".45" stroke-linecap="round"/>`;
  s += `</g>`;

  // skirt + hem frill
  s += `<path d="${SKIRT}" fill="url(#redG)" stroke="${C.redLine}" stroke-width="3"/>`;
  s += `<g clip-path="url(#clipSkirt)">`;
  for (const [x1, x2, w, b] of [[420, 360, 30, -6], [470, 450, 26, 4], [530, 552, 26, -4], [585, 640, 30, 6]]) {
    s += `<path d="${wedge(x1, 1150, x2, 1440, 4, w, b)}" fill="${C.redShadow}" opacity=".85"/>`;
  }
  s += `</g>`;
  s += frill('M 376 1146 Q 500 1166 624 1146', { scal: 40, gap: 22, band: 30 });

  // red top
  s += frill(TOP_EDGE, { scal: 30, gap: 17, band: 20 });
  s += `<path d="${TOP}" fill="url(#redG)" stroke="${C.redLine}" stroke-width="3" stroke-linejoin="round"/>`;
  s += `<g clip-path="url(#clipTop)">`;
  s += `<path d="M 360 892 C 380 960, 390 1060, 404 1150 L 370 1150 C 356 1060, 350 960, 360 892 Z" fill="${C.redShadow}"/>`;
  s += `<path d="M 640 892 C 625 960, 616 1060, 600 1150 L 630 1150 C 644 1060, 650 960, 640 892 Z" fill="${C.redShadow}"/>`;
  s += `<path d="M 450 800 C 470 840, 530 840, 550 800 L 550 790 L 450 790 Z" fill="${C.redShadow}" opacity=".9"/>`;
  s += `<path d="M 360 1004 Q 440 1046 500 1020 Q 560 1046 640 1004 L 650 1170 L 350 1170 Z" fill="${C.redShadow}" opacity=".5" filter="url(#b6)"/>`;
  s += `<ellipse cx="446" cy="958" rx="40" ry="28" fill="${C.redLight}" opacity=".35" filter="url(#b6)"/>`;
  s += `<ellipse cx="556" cy="958" rx="40" ry="28" fill="${C.redLight}" opacity=".25" filter="url(#b6)"/>`;
  s += `<path d="${wedge(370, 912, 450, 1000, 8, 0, 8)}" fill="${C.redShadow}" opacity=".85"/>`;
  s += `<path d="${wedge(630, 912, 552, 1000, 8, 0, -8)}" fill="${C.redShadow}" opacity=".85"/>`;
  s += `<path d="${wedge(436, 1150, 452, 1060, 12, 0, 3)}" fill="${C.redShadow}" opacity=".8"/>`;
  s += `<path d="${wedge(566, 1150, 548, 1070, 10, 0, -3)}" fill="${C.redShadow}" opacity=".8"/>`;
  s += `<path d="${wedge(500, 1150, 502, 1086, 6, 0, 2)}" fill="${C.redShadow}" opacity=".6"/>`;
  s += `</g>`;

  // yellow ascot
  const wing = 'M 494 801 C 478 784, 452 780, 446 795 C 440 811, 452 830, 471 827 C 481 825, 490 819, 494 814 Z';
  s += `<path d="M 494 812 C 487 846, 478 880, 465 914 L 482 905 L 492 920 C 497 886, 501 852, 503 816 Z" fill="url(#yellowG)" stroke="${C.yellowLine}" stroke-width="2.5" stroke-linejoin="round"/>`;
  s += `<path d="M 502 814 C 509 850, 517 886, 530 922 L 514 913 L 506 927 C 503 892, 500 854, 497 816 Z" fill="url(#yellowG)" stroke="${C.yellowLine}" stroke-width="2.5" stroke-linejoin="round"/>`;
  s += `<path d="M 497 820 C 494 860, 490 890, 486 902" stroke="${C.yellowShadow}" stroke-width="5" fill="none" opacity=".8"/>`;
  s += `<path d="${wing}" fill="url(#yellowG)" stroke="${C.yellowLine}" stroke-width="2.5" stroke-linejoin="round"/>`;
  s += mirror(`<path d="${wing}" fill="url(#yellowG)" stroke="${C.yellowLine}" stroke-width="2.5" stroke-linejoin="round"/>`);
  s += `<path d="M 492 806 C 478 800, 462 800, 456 806 C 466 812, 480 814, 492 812 Z" fill="${C.yellowShadow}" opacity=".85"/>`;
  s += `<path d="M 508 806 C 522 800, 538 800, 544 806 C 534 812, 520 814, 508 812 Z" fill="${C.yellowShadow}" opacity=".85"/>`;
  s += `<ellipse cx="500" cy="809" rx="11" ry="12.5" fill="url(#yellowG)" stroke="${C.yellowLine}" stroke-width="2.5"/>`;
  s += `<path d="M 494 802 Q 500 798 505 803" stroke="#fff6c8" stroke-width="2.5" fill="none" stroke-linecap="round"/>`;
  return s;
}

function sleeveLeft() {
  const shape = 'M 276 985 C 238 1080, 198 1230, 158 1440 L 446 1440 C 426 1270, 400 1110, 370 985 C 345 1002, 300 1002, 276 985 Z';
  let s = frill('M 278 984 C 302 1002, 344 1002, 368 984', { scal: 26, gap: 15, band: 16 });
  s += `<path d="${shape}" fill="url(#sleeveG)" stroke="${C.whiteLine}" stroke-width="2.8" stroke-linejoin="round"/>`;
  s += `<clipPath id="clipSL"><path d="${shape}"/></clipPath><g clip-path="url(#clipSL)">`;
  s += `<path d="${wedge(372, 990, 446, 1440, 4, 80, -10)}" fill="${C.whiteShadow}"/>`;
  s += `<path d="${wedge(312, 1004, 238, 1440, 3, 26, 10)}" fill="${C.whiteShadow}"/>`;
  s += `<path d="${wedge(338, 1006, 330, 1440, 2, 18, -6)}" fill="${C.whiteShadow}" opacity=".8"/>`;
  s += `<path d="${wedge(300, 1000, 180, 1440, 2, 40, 14)}" fill="${C.whiteShadow}" opacity=".7"/>`;
  s += `<path d="${wedge(380, 1004, 418, 1440, 2, 24, -4)}" fill="${C.whiteDeep}" opacity=".6"/>`;
  s += `</g>`;
  // red ribbon band around the top + little bow
  s += `<path d="M 274 984 C 300 1004, 346 1004, 372 984" stroke="${C.redLine}" stroke-width="15" fill="none" stroke-linecap="round"/>`;
  s += `<path d="M 274 984 C 300 1004, 346 1004, 372 984" stroke="${C.red}" stroke-width="10" fill="none" stroke-linecap="round"/>`;
  s += `<path d="M 288 996 C 270 990, 256 1000, 262 1012 C 270 1020, 284 1010, 290 1000 Z" fill="${C.red}" stroke="${C.redLine}" stroke-width="2"/>`;
  s += `<path d="M 292 998 C 300 1016, 296 1030, 286 1040 L 296 1042 C 304 1030, 306 1014, 296 998 Z" fill="${C.redShadow}" stroke="${C.redLine}" stroke-width="2"/>`;
  return s;
}

function sleeveRight() {
  const shape = 'M 724 985 C 764 1075, 806 1190, 846 1300 C 820 1334, 732 1338, 650 1312 C 648 1200, 646 1080, 630 985 C 655 1002, 700 1002, 724 985 Z';
  const cuff = 'M 846 1300 C 820 1334, 732 1338, 650 1312';
  let s = '';
  // dark sleeve interior seen through the cuff
  s += `<path d="M 652 1300 C 700 1290, 800 1284, 846 1298 C 820 1334, 732 1338, 652 1312 Z" fill="${C.whiteDeep}"/>`;
  s += frill(cuff, { scal: 30, gap: 18, band: 20 });
  s += frill('M 722 984 C 698 1002, 656 1002, 632 984', { scal: 26, gap: 15, band: 16 });
  s += `<path d="${shape}" fill="url(#sleeveG)" stroke="${C.whiteLine}" stroke-width="2.8" stroke-linejoin="round"/>`;
  s += `<clipPath id="clipSR"><path d="${shape}"/></clipPath><g clip-path="url(#clipSR)">`;
  s += `<path d="${wedge(632, 990, 652, 1320, 4, 70, 6)}" fill="${C.whiteShadow}"/>`;
  s += `<path d="${wedge(690, 1004, 760, 1330, 3, 26, -8)}" fill="${C.whiteShadow}"/>`;
  s += `<path d="${wedge(708, 1004, 820, 1320, 2, 30, -12)}" fill="${C.whiteShadow}" opacity=".7"/>`;
  s += `<path d="${wedge(660, 1004, 690, 1330, 2, 18, 4)}" fill="${C.whiteDeep}" opacity=".55"/>`;
  s += `</g>`;
  s += `<path d="M 726 984 C 700 1004, 654 1004, 628 984" stroke="${C.redLine}" stroke-width="15" fill="none" stroke-linecap="round"/>`;
  s += `<path d="M 726 984 C 700 1004, 654 1004, 628 984" stroke="${C.red}" stroke-width="10" fill="none" stroke-linecap="round"/>`;
  s += mirror(`<path d="M 288 996 C 270 990, 256 1000, 262 1012 C 270 1020, 284 1010, 290 1000 Z" fill="${C.red}" stroke="${C.redLine}" stroke-width="2"/>` +
    `<path d="M 292 998 C 300 1016, 296 1030, 286 1040 L 296 1042 C 304 1030, 306 1014, 296 998 Z" fill="${C.redShadow}" stroke="${C.redLine}" stroke-width="2"/>`);
  return s;
}

function neck() {
  let s = `<path d="${NECK}" fill="${C.skin}" stroke="${C.skinLine}" stroke-width="2.4"/>`;
  s += `<g clip-path="url(#clipNeck)"><g transform="${TILT}"><path d="${FACE}" transform="translate(0 34)" fill="${C.skinShadow}"/>` +
    `<path d="${FACE}" transform="translate(0 20)" fill="${C.skinDeep}" opacity=".35"/></g></g>`;
  return s;
}

function eyeShapes() {
  // defined around (0,0) for the viewer-left eye; outer corner on the left
  return {
    sclera: 'M -50 6 C -40 -24, -18 -33, 8 -33 C 28 -33, 42 -24, 48 -12 C 47 14, 38 34, 12 40 C -18 44, -42 32, -50 6 Z',
    lash: 'M -58 12 C -50 -22, -22 -42, 6 -42 C 30 -42, 46 -31, 54 -14 L 50 -7 C 40 -24, 26 -31, 6 -31 C -20 -31, -40 -17, -49 9 Z',
    wing: 'M -47 -6 L -68 2 L -52 13 Z M -38 -20 L -61 -25 L -44 -11 Z M -24 -34 L -40 -47 L -16 -38 Z',
    lower: 'M -42 25 Q -14 47 24 41 Q -12 43 -42 25 Z',
    crease: 'M -42 -40 Q -10 -56 32 -45',
  };
}

function eye(cx, cy, flip) {
  const e = eyeShapes();
  const id = flip ? 'R' : 'L';
  let s = '';
  s += `<clipPath id="sc${id}"><path d="${e.sclera}"/></clipPath>`;
  s += `<path d="${e.sclera}" fill="url(#scleraG)"/>`;
  s += `<g clip-path="url(#sc${id})">`;
  s += `<ellipse cx="-2" cy="5" rx="31" ry="39" fill="url(#irisG)" stroke="#3a0709" stroke-width="2.5"/>`;
  s += `<ellipse cx="-2" cy="4" rx="13" ry="20" fill="#1f0405"/>`;
  s += `<path d="M -28 18 C -20 40, 16 40, 25 18 C 14 30, -16 30, -28 18 Z" fill="#ffc48a" opacity=".75" filter="url(#b1)"/>`;
  s += `<path d="M -40 -33 C -10 -18, 20 -18, 50 -30 L 50 -40 L -40 -40 Z" fill="#2a0607" opacity=".55"/>`;
  s += `</g>`;
  s += `<path d="${e.lash}" fill="${C.lash}"/>`;
  s += `<path d="${e.wing}" fill="${C.lash}"/>`;
  s += `<path d="${e.lower}" fill="${C.lash}" opacity=".85"/>`;
  s += `<path d="${e.crease}" stroke="${C.lash}" stroke-width="2" fill="none" opacity=".55" stroke-linecap="round"/>`;
  const tr = flip ? `translate(${cx} ${cy}) scale(-1.05 1.05)` : `translate(${cx} ${cy}) scale(1.05)`;
  return `<g transform="${tr}">${s}</g>`;
}

function eyeHighlights(ix, iy) {
  return `<ellipse cx="${ix - 13}" cy="${iy - 13}" rx="9" ry="12.5" fill="#fff" transform="rotate(-18 ${ix - 13} ${iy - 13})"/>` +
    `<circle cx="${ix + 13}" cy="${iy + 20}" r="4.5" fill="#fff" opacity=".95"/>` +
    `<circle cx="${ix - 2}" cy="${iy - 24}" r="2.5" fill="#fff" opacity=".8"/>`;
}

function face() {
  let s = '';
  s += `<path d="${FACE}" fill="url(#skinG)" stroke="${C.skinLine}" stroke-width="2.6"/>`;
  s += `<g clip-path="url(#clipFace)">`;
  // shadow cast by the fringe and side hair
  s += `<path d="${CAP}" transform="translate(6 18)" fill="${C.skinShadow}"/>`;
  s += `<path d="${SIDE_HAIR}" transform="translate(12 6)" fill="${C.skinShadow}"/>`;
  s += mirror(`<path d="${SIDE_HAIR}" transform="translate(12 6)" fill="${C.skinShadow}"/>`);
  s += `<path d="M 350 430 C 348 540, 362 610, 404 658 C 436 694, 472 716, 500 720 L 500 704 C 474 700, 446 684, 422 654 C 386 610, 372 540, 374 430 Z" fill="${C.skinShadow}" opacity=".7" filter="url(#b3)"/>`;
  // blush
  s += `<ellipse cx="408" cy="628" rx="36" ry="15" fill="#ff8f9a" opacity=".45" filter="url(#b6)"/>`;
  s += `<ellipse cx="592" cy="628" rx="36" ry="15" fill="#ff8f9a" opacity=".45" filter="url(#b6)"/>`;
  s += `</g>`;
  for (const x of [392, 404, 416, 584, 596, 608]) {
    s += `<path d="M ${x + 5} 618 L ${x - 3} 634" stroke="#f0707f" stroke-width="2" stroke-linecap="round" opacity=".55"/>`;
  }
  // eyes
  s += eye(432, 566, false) + eye(568, 566, true);
  s += eyeHighlights(430, 571) + eyeHighlights(566, 571);
  // nose + mouth
  s += `<ellipse cx="503" cy="616" rx="2.5" ry="5" fill="#fff" opacity=".7"/>`;
  s += `<path d="M 499 624 Q 503 631 497 635" stroke="${C.skinLine}" stroke-width="2.2" fill="none" stroke-linecap="round" opacity=".8"/>`;
  s += `<path d="M 487 665 Q 500 682 513 665 Q 500 670 487 665 Z" fill="#a8333f" stroke="#7a1f29" stroke-width="1.6" stroke-linejoin="round"/>`;
  s += `<path d="M 493 673.5 Q 500 680 507 673.5 Q 500 676 493 673.5 Z" fill="#ef8b93"/>`;
  s += `<path d="M 484 663 Q 487 661 490 664" stroke="#a8333f" stroke-width="1.4" fill="none" stroke-linecap="round" opacity=".7"/>`;
  return s;
}

function sidelock() {
  let s = '';
  s += `<path d="${SIDELOCK_TOP}" fill="url(#lockG)" stroke="${C.hairLine}" stroke-width="2.6" stroke-linejoin="round"/>`;
  s += `<path d="${wedge(350, 580, 340, 790, 5, 3, -4)}" fill="${C.hairDeep}" opacity=".7"/>`;
  s += `<path d="${wedge(368, 600, 358, 790, 6, 4, 2)}" fill="${C.hairLight}" opacity=".45"/>`;
  s += `<path d="${SIDELOCK_END}" fill="url(#lockG)" stroke="${C.hairLine}" stroke-width="2.6" stroke-linejoin="round"/>`;
  s += `<path d="${wedge(338, 884, 326, 1000, 6, 0, 4)}" fill="${C.hairDeep}" opacity=".7"/>`;
  s += `<path d="${wedge(350, 884, 352, 1000, 5, 0, -3)}" fill="${C.hairLight}" opacity=".45"/>`;
  // hair tube (white with red bands)
  const tube = 'M 309 790 Q 337 802 367 790 L 368 882 Q 337 894 308 882 Z';
  s += `<path d="${tube}" fill="url(#tubeG)" stroke="${C.whiteLine}" stroke-width="2.6" stroke-linejoin="round"/>`;
  s += `<path d="M 309 797 Q 337 809 367 797 L 367 811 Q 337 823 309 811 Z" fill="${C.red}" stroke="${C.redLine}" stroke-width="1.6"/>`;
  s += `<path d="M 308 862 Q 337 874 368 862 L 368 874 Q 337 886 308 874 Z" fill="${C.red}" stroke="${C.redLine}" stroke-width="1.6"/>`;
  s += `<path d="M 322 820 L 322 858" stroke="#ffffff" stroke-width="5" stroke-linecap="round" opacity=".9"/>`;
  s += `<path d="M 314 830 Q 337 840 362 830" stroke="${C.whiteDeep}" stroke-width="1.6" fill="none" opacity=".8"/>`;
  return s;
}

function sideHair() {
  let s = `<path d="${SIDE_HAIR}" fill="url(#lockG)" stroke="${C.hairLine}" stroke-width="2.6" stroke-linejoin="round"/>`;
  s += `<g clip-path="url(#clipSide)">`;
  s += `<path d="${wedge(352, 470, 340, 735, 10, 0, -6)}" fill="${C.hairDeep}" opacity=".55"/>`;
  s += `<path d="${wedge(380, 480, 386, 700, 6, 0, 4)}" fill="${C.hairLight}" opacity=".45"/>`;
  s += `</g>`;
  return s;
}

function angelRing() {
  // glossy "angel ring": a thin band following the crown with jagged spikes
  // pointing up and down the strands (classic cel-shaded hair highlight)
  const top = [], bot = [], coreT = [], coreB = [];
  const N = 30;
  for (let i = 0; i <= N; i++) {
    const t = 0.05 + (0.9 * i) / N;
    const x = (1 - t) * (1 - t) * 340 + 2 * (1 - t) * t * 500 + t * t * 660;
    const y = (1 - t) * (1 - t) * 396 + 2 * (1 - t) * t * 284 + t * t * 396;
    const dx = x - 500, dy = y - 120, L = Math.hypot(dx, dy), ux = dx / L, uy = dy / L;
    const k = Math.pow(Math.sin(Math.PI * (i / N)), 0.6);
    const band = 7 * k;
    const up = i % 2 === 0 ? rr(8, 22) * k : 0;
    const dn = i % 2 === 1 ? rr(14, 34) * k : 0;
    top.push([x - ux * (band + up), y - uy * (band + up)]);
    bot.push([x + ux * (band + dn), y + uy * (band + dn)]);
    coreT.push([x - ux * band * 0.4, y - uy * band * 0.4]);
    coreB.push([x + ux * (band * 0.4 + dn * 0.45), y + uy * (band * 0.4 + dn * 0.45)]);
  }
  const poly = (A, B) => 'M ' + A.map((p) => p.map(f).join(' ')).join(' L ') + ' L ' + B.slice().reverse().map((p) => p.map(f).join(' ')).join(' L ') + ' Z';
  return `<path d="${poly(top, bot)}" fill="${C.hairLight}" opacity=".8" stroke-linejoin="round"/>` +
    `<path d="${poly(coreT.slice(6, N - 5), coreB.slice(6, N - 5))}" fill="#e0a98c" opacity=".65"/>`;
}

function frontHair() {
  let s = '';
  // second layer of fringe peeking out of the notches
  const back = CAP_TOP +
    ' Q 668 532 640 554 Q 630 506 614 476 Q 608 512 588 536 Q 578 494 568 470 Q 562 512 542 540' +
    ' Q 532 500 524 474 Q 514 518 502 554 Q 488 508 482 474 Q 472 512 454 536 Q 446 496 440 474' +
    ' Q 428 510 410 532 Q 402 498 394 474 Q 382 512 362 540 Q 334 518 326 494 Z';
  s += `<path d="${back}" fill="${C.hairDark}" stroke="${C.hairLine}" stroke-width="2.4" stroke-linejoin="round"/>`;
  s += `<path d="${CAP}" fill="url(#hairG)" stroke="${C.hairLine}" stroke-width="3" stroke-linejoin="round"/>`;
  s += `<g clip-path="url(#clipCap)">`;
  // cel shadow on the shaded (viewer-left) side of the crown and a rim on the right
  s += `<path d="M 286 300 C 346 262, 408 248, 446 256 C 394 300, 370 384, 370 540 L 286 540 Z" fill="${C.hairDark}" opacity=".5"/>`;
  s += `<path d="M 714 340 C 696 300, 676 280, 652 266 C 672 330, 678 410, 672 540 L 714 540 Z" fill="${C.hairDark}" opacity=".35"/>`;
  s += angelRing();
  // tapered strand separations running into every notch of the fringe
  for (const [nx, ny] of BANG_NOTCHES) {
    s += `<path d="${wedge(500 + (nx - 500) * 0.6, 300, nx, ny + 2, 0.5, 5, (nx - 500) * 0.05)}" fill="${C.hairDeep}" opacity=".6"/>`;
  }
  // lighter streaks down the fringe
  for (const [x1, y1, x2, y2] of [[622, 430, 648, 528], [574, 430, 594, 508], [530, 430, 516, 530], [484, 440, 484, 540], [442, 432, 438, 512], [398, 438, 390, 508]]) {
    s += `<path d="${wedge(x1, y1, x2, y2, 7, 0, 3)}" fill="${C.hairLight}" opacity=".45"/>`;
  }
  // warm light bounced from the skin onto the fringe tips
  s += `<path d="${CAP}" fill="none" stroke="#c98667" stroke-width="6" opacity=".3" transform="translate(0 -3)"/>`;
  s += `</g>`;
  s += `<mask id="rimCap"><path d="${CAP}" fill="#fff"/><path d="${CAP}" transform="translate(-6 5)" fill="#000"/></mask>`;
  s += `<rect x="300" y="200" width="420" height="380" fill="#ffd6bf" opacity=".5" mask="url(#rimCap)"/>`;
  // a couple of loose strands over the fringe
  s += `<path d="M 534 424 Q 526 500 500 566 Q 520 506 520 430 Z" fill="${C.hair}" stroke="${C.hairLine}" stroke-width="2" stroke-linejoin="round"/>`;
  // see-through eyebrows drawn over the fringe
  s += `<path d="M 392 506 Q 430 486 468 497 Q 430 494 392 506 Z" fill="${C.hairDeep}" stroke="${C.hairDeep}" stroke-width="2" stroke-linejoin="round" opacity=".32"/>`;
  s += `<path d="M 608 506 Q 570 486 532 497 Q 570 494 608 506 Z" fill="${C.hairDeep}" stroke="${C.hairDeep}" stroke-width="2" stroke-linejoin="round" opacity=".32"/>`;
  return s;
}

// ---------------------------------------------------------------------------
// props
// ---------------------------------------------------------------------------
function shide(x, y, dir) {
  // zig-zag paper streamer (shide): horizontal folds, alternating offset,
  // drifting outward so the pair hangs in a gentle "A" shape
  let s = '';
  const w = 32, h = 42, a = 21, drift = 11, n = 5;
  const cx = (i) => x + dir * (i * drift + (i % 2 ? a : 0));
  const cy = (i) => y + i * h;
  for (let i = 0; i < n; i++) {
    const x0 = cx(i), x1 = cx(i + 1) - (i === n - 1 ? dir * a * 0.5 : 0), y0 = cy(i), y1 = cy(i + 1);
    const d = `M ${f(x0 - w / 2)} ${f(y0)} L ${f(x0 + w / 2)} ${f(y0)} L ${f(x1 + w / 2)} ${f(y1)} L ${f(x1 - w / 2)} ${f(y1)} Z`;
    const front = i % 2 === 0;
    s += `<path d="${d}" fill="${front ? '#ffffff' : C.whiteShadow}" stroke="${C.whiteLine}" stroke-width="2" stroke-linejoin="round"/>`;
    if (!front) s += `<path d="M ${f(x0 - w / 2 + 3)} ${f(y0 + 3)} L ${f(x0 + w / 2 - 3)} ${f(y0 + 3)}" stroke="${C.whiteDeep}" stroke-width="4" opacity=".8"/>`;
  }
  return s;
}

function gohei() {
  const { x: tx, y: ty } = STICK_TOP, { x: bx, y: by } = STICK_BOT;
  let s = '';
  s += `<line x1="${f(bx)}" y1="${by}" x2="${f(tx)}" y2="${ty}" stroke="#5a3a1c" stroke-width="16" stroke-linecap="round"/>`;
  s += `<line x1="${f(bx)}" y1="${by}" x2="${f(tx)}" y2="${ty}" stroke="#e2b779" stroke-width="11" stroke-linecap="round"/>`;
  s += `<line x1="${f(bx - 2)}" y1="${by}" x2="${f(tx - 2)}" y2="${ty}" stroke="#fbe6bd" stroke-width="3" stroke-linecap="round"/>`;
  s += `<line x1="${f(bx + 3)}" y1="${by}" x2="${f(tx + 3)}" y2="${ty}" stroke="#a87741" stroke-width="3" stroke-linecap="round" opacity=".8"/>`;
  // golden cap + streamers
  s += `<g transform="rotate(${STICK_ANGLE} ${f(tx)} ${ty})">`;
  s += `<rect x="${f(tx - 10)}" y="${ty - 22}" width="20" height="26" rx="4" fill="#f2c94c" stroke="#8a5a10" stroke-width="2"/>`;
  s += `<rect x="${f(tx - 6)}" y="${ty - 18}" width="5" height="18" rx="2" fill="#fff3b8"/>`;
  s += `</g>`;
  s += `<g transform="rotate(4 ${f(tx)} ${ty})">${shide(tx + 17, ty - 6, 1)}</g>`;
  s += `<g transform="rotate(-2 ${f(tx)} ${ty})">${shide(tx - 17, ty - 6, -1)}</g>`;
  return s;
}

function fist() {
  // left hand (viewer's right) gripping the gohei: back of the hand faces us,
  // the fingers curl round the rod on the left, the thumb hooks over the top
  const { x, y } = FIST;
  const L = C.skinLine;
  let s = '';
  // wrist + back of the hand
  const back = 'M -2 -34 C 10 -42, 30 -44, 40 -36 C 46 -20, 46 6, 40 22 C 34 32, 20 34, 6 32 L -4 30 Z';
  s += `<clipPath id="clipHandBack"><path d="${back}"/></clipPath>`;
  s += `<path d="${back}" fill="${C.skin}" stroke="${L}" stroke-width="2.4" stroke-linejoin="round"/>`;
  s += `<g clip-path="url(#clipHandBack)">`;
  s += `<rect x="-10" y="-50" width="70" height="22" fill="${C.skinDeep}" opacity=".75" filter="url(#b3)"/>`;
  s += `<path d="M 30 -40 C 44 -20, 44 10, 34 30 L 50 30 L 50 -40 Z" fill="${C.skinShadow}"/>`;
  s += `</g>`;
  // knuckle ridge
  s += `<path d="M 8 -22 C 12 -10, 12 8, 8 22" stroke="${L}" stroke-width="1.6" fill="none" opacity=".5"/>`;
  // four curled fingers stacked along the rod (index on top, pinky at the bottom)
  const fingers = [[-27, 15, 10], [-12, 15, 12], [3, 14, 11], [17, 12, 8]];
  for (const [fy, fh, tip] of fingers) {
    const y0 = fy, y1 = fy + fh, xl = -18 - tip * 0.6;
    const d = `M 12 ${y0} L ${f(xl + 6)} ${y0} C ${f(xl - 3)} ${y0}, ${f(xl - 3)} ${y1}, ${f(xl + 6)} ${y1} L 12 ${y1} Z`;
    s += `<path d="${d}" fill="${C.skin}" stroke="${L}" stroke-width="2.2" stroke-linejoin="round"/>`;
    s += `<path d="M ${f(xl + 4)} ${y1 - 3} Q ${f(xl + 12)} ${y1 - 1} 10 ${y1 - 2}" stroke="${C.skinShadow}" stroke-width="4" fill="none" opacity=".9"/>`;
    s += `<path d="M -6 ${y0 + 3} L -6 ${y1 - 3}" stroke="${L}" stroke-width="1.4" opacity=".55"/>`;
  }
  // thumb hooking over the index finger
  s += `<path d="M 30 -36 C 18 -44, -2 -42, -16 -34 C -22 -30, -20 -24, -13 -23 C -2 -24, 10 -26, 20 -22 Z" fill="${C.skin}" stroke="${L}" stroke-width="2.2" stroke-linejoin="round"/>`;
  s += `<path d="M -15 -33 C -19 -31, -19 -26, -14 -25 C -10 -26, -9 -31, -15 -33 Z" fill="#ffe2da" stroke="${L}" stroke-width="1.2"/>`;
  s += `<path d="M 22 -38 Q 6 -41 -6 -37" stroke="#fff" stroke-width="2.4" fill="none" opacity=".7" stroke-linecap="round"/>`;
  return `<g transform="translate(${x} ${y}) rotate(${STICK_ANGLE})">${s}</g>`;
}

function orb(cx, cy, r, rot) {
  const yin = `M 0 ${-r} A ${r} ${r} 0 0 1 0 ${r} A ${r / 2} ${r / 2} 0 0 1 0 0 A ${r / 2} ${r / 2} 0 0 0 0 ${-r} Z`;
  let s = '';
  s += `<circle cx="${cx}" cy="${cy}" r="${r * 2}" fill="url(#orbGlow)"/>`;
  s += `<g transform="translate(${cx} ${cy}) rotate(${rot})">`;
  s += `<circle r="${r}" fill="#fff"/>`;
  s += `<path d="${yin}" fill="${C.red}"/>`;
  s += `<circle cy="${-r / 2}" r="${r * 0.15}" fill="#fff"/><circle cy="${r / 2}" r="${r * 0.15}" fill="${C.red}"/>`;
  s += `</g>`;
  s += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="url(#orbShade)" stroke="${C.redLine}" stroke-width="2.6"/>`;
  s += `<ellipse cx="${cx - r * 0.36}" cy="${cy - r * 0.42}" rx="${r * 0.3}" ry="${r * 0.17}" fill="#fff" opacity=".85" transform="rotate(-35 ${cx - r * 0.36} ${cy - r * 0.42})"/>`;
  s += `<circle cx="${cx + r * 0.42}" cy="${cy + r * 0.45}" r="${r * 0.07}" fill="#fff" opacity=".7"/>`;
  return s;
}

function ofuda(cx, cy, rot, sc = 1) {
  let s = '';
  s += `<rect x="-22" y="-60" width="44" height="120" rx="2" fill="#000" opacity=".12" transform="translate(5 7)" filter="url(#b3)"/>`;
  s += `<rect x="-22" y="-60" width="44" height="120" rx="2" fill="#fffaf0" stroke="#b9a98a" stroke-width="1.6"/>`;
  s += `<rect x="-17" y="-55" width="34" height="110" fill="none" stroke="${C.red}" stroke-width="2"/>`;
  // abstract seal glyphs
  s += `<circle cx="0" cy="-34" r="9" fill="none" stroke="${C.red}" stroke-width="2.6"/>`;
  s += `<path d="M -6 -34 L 6 -34 M 0 -40 L 0 -28" stroke="${C.red}" stroke-width="2.2"/>`;
  s += `<path d="M 0 -18 L 0 44 M -9 -10 L 9 -10 M -8 2 L 8 0 M -10 14 Q 0 8 10 16 M -7 28 L 7 26 M -9 38 L 0 32 L 9 38" stroke="${C.red}" stroke-width="2.6" fill="none" stroke-linecap="round"/>`;
  return `<g transform="translate(${cx} ${cy}) rotate(${rot}) scale(${sc})">${s}</g>`;
}

function sparkle(x, y, s, op = 1) {
  return `<path d="M ${x} ${y - s} Q ${x + s * 0.15} ${y - s * 0.15} ${x + s} ${y} Q ${x + s * 0.15} ${y + s * 0.15} ${x} ${y + s} Q ${x - s * 0.15} ${y + s * 0.15} ${x - s} ${y} Q ${x - s * 0.15} ${y - s * 0.15} ${x} ${y - s} Z" fill="#fff" opacity="${op}"/>`;
}

// ---------------------------------------------------------------------------
// compose
// ---------------------------------------------------------------------------
function petals(n, opts) {
  let s = '';
  for (let i = 0; i < n; i++) {
    const x = rr(0, W), y = rr(0, H);
    if (opts.avoid && x > 280 && x < 720 && y > 220 && y < 1000) continue;
    s += petal(x, y, rr(opts.min, opts.max), rr(0, 360), rr(opts.opMin, 1), opts.blur);
  }
  return s;
}

function build() {
  const parts = [];
  parts.push(defs());
  parts.push(background());
  parts.push(petals(26, { min: 5, max: 9, opMin: 0.45, blur: 'b1' }));
  parts.push(`<ellipse cx="500" cy="620" rx="420" ry="560" fill="url(#aura)"/>`);
  parts.push(branches());

  parts.push(`<g transform="${TILT}">${backHair()}${bow()}</g>`);
  parts.push(body());
  parts.push(sleeveLeft());
  parts.push(sleeveRight());
  parts.push(neck());
  parts.push(`<g transform="${TILT}">${face()}${sidelock()}${mirror(sidelock())}${sideHair()}${mirror(sideHair())}${frontHair()}</g>`);

  parts.push(gohei());
  parts.push(fist());

  parts.push(orb(118, 760, 52, -30));
  parts.push(orb(892, 1050, 38, 40));
  parts.push(ofuda(100, 1120, -18, 0.95) + ofuda(905, 420, 16, 0.8) + ofuda(140, 450, 10, 0.7));

  parts.push(petals(30, { min: 8, max: 15, opMin: 0.7, avoid: true }));
  parts.push(petal(80, 1300, 26, 30, 0.75, 'b3') + petal(940, 1290, 30, -40, 0.7, 'b3') + petal(930, 300, 22, 70, 0.8, 'b3'));
  parts.push(sparkle(170, 700, 14) + sparkle(84, 830, 8, 0.8) + sparkle(930, 1000, 10) + sparkle(860, 1110, 7, 0.8) + sparkle(960, 560, 9, 0.8));

  parts.push(`<rect width="${W}" height="${H}" fill="url(#lightWash)" style="mix-blend-mode:screen"/>`);
  parts.push(`<rect width="${W}" height="${H}" fill="url(#vignette)"/>`);

  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}">\n${parts.join('\n')}\n</svg>\n`;
}

async function main() {
  const svg = build();
  const svgPath = path.join(here, 'reimu.svg');
  const pngPath = path.join(here, 'reimu.png');
  fs.writeFileSync(svgPath, svg);
  const scale = Number(process.env.SCALE || 2);
  const { chromium } = require('playwright');
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: W, height: H }, deviceScaleFactor: scale });
  await page.setContent(`<!doctype html><html><body style="margin:0;background:#fff">${svg}</body></html>`);
  await page.screenshot({ path: process.env.OUT || pngPath, clip: { x: 0, y: 0, width: W, height: H } });
  await browser.close();
  console.log('wrote', svgPath, 'and', process.env.OUT || pngPath);
}

main().catch((e) => { console.error(e); process.exit(1); });
