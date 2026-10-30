'use strict';
// ------------------------------------------------------------
//  Script + storyboard. Each scene: beats (captions with timing)
//  and draw(t, S, ctx) for the glowing visual layer, plus over()
//  for crisp overlays (formulas, labels). S.b[i]/S.e[i] are the
//  start/end of beat i in scene-local seconds.
// ------------------------------------------------------------

function defineTex() {
  const T = COL.text, Gd = COL.gold, Bl = '#8fb4dc';
  defTex('cont', '\\left(2^{\\aleph_0}\\right)^{n} = 2^{\\aleph_0}', T, 36);
  defTex('aleph0', '\\aleph_0', Gd, 28);
  defTex('c2aleph0', '2^{\\aleph_0}', Gd, 28);
  defTex('n3', 'N = b^{\\,n} = 3^{4}', T, 30);
  defTex('bek', 'I \\le \\frac{2\\pi R E}{\\hbar c \\ln 2}', T, 32);
  defTex('mkg', 'm \\approx 1.5\\,\\mathrm{kg}', COL.dim, 26);
  defTex('r01', 'R \\approx 0.1\\,\\mathrm{m}', COL.dim, 26);
  defTex('brainI', 'I \\approx 4\\times 10^{42}\\ \\text{bits}', T, 28);
  defTex('brainN', 'N \\le 2^{4\\times 10^{42}}', Gd, 36);
  defTex('dt', '\\Delta t = 10^{-30}\\,\\mathrm{s}', Gd, 26);
  defTex('Neps', 'N(\\varepsilon) < \\infty', Gd, 32);
  defTex('life', '16 \\times 3600 \\times 365 \\times 80 \\;\\approx\\; 1.7\\times 10^{9}', T, 30);
  defTex('lifeN', '2^{1.7\\times 10^{9}} \\approx 10^{5\\times 10^{8}}', Gd, 34);
  defTex('big10', '10^{5\\times 10^{8}}', Gd, 26);
  defTex('not3', '3 \\notin \\{0,\\,2,\\,4,\\,6,\\,\\dots\\}', T, 30);
  defTex('p0', 'P(X = x) = 0', T, 32);
  defTex('nform', 'N \\;\\approx\\; {\\color{#8fb4dc}{b}}^{\\,{\\color{#e9c27b}{T}}/\\Delta t}', T, 44);
  defTex('nb', 'b', Bl, 34); defTex('nT', 'T', Gd, 34); defTex('ndt', '\\Delta t', COL.dim, 34);
  defTex('half', '2^{1.7\\times 10^{5}} \\approx 10^{51000}', Gd, 34);
  defTex('gg80', '10^{51000} \\gg 10^{80}', T, 26);
  defTex('pT', 'p(T)', T, 26);
  defTex('pTd', 'p(T \\mid \\text{D})', Gd, 26);
  defTex('bij', 'f(x) = \\frac{x}{1-x}', T, 32);
  defTex('teps', 'N(\\varepsilon) \\;\\approx\\; \\frac{T}{\\varepsilon}', T, 34);
}

// ---------- shared scene state ----------
let TREE_A, TREE_K3, TREE_P, TREE_DENSE, TREE_COVER, PATH_A, PATH_COVER;
let BLOB, COVER1, COVER2, DAG, LIFELINES, CURVES;
const V_ZOOM = 0.314159265358979, DIG = '31415926535897932384626';

function prepScenes() {
  TREE_A = makeTree({ seed: 11, x0: 230, y0: 470, x1: 1790, yTop: 60, yBot: 880, depth: 9, yj: 0.45 });
  PATH_A = pathOf(TREE_A, 4, 0.55);
  TREE_K3 = makeTree({ seed: 5, k: 3, depth: 4, x0: 420, y0: 430, x1: 1300, yTop: 100, yBot: 760, yj: 0.25, xj: 0 });
  TREE_P = makeTree({ seed: 23, x0: 200, y0: 410, x1: 1760, yTop: 50, yBot: 760, depth: 9, yj: 0.45 });
  const R = rng(77);
  for (const n of TREE_P.nodes) n.drop = n.parent ? (n.parent.drop || (n.d >= 2 && R() < 0.3)) : false;
  TREE_DENSE = makeTree({ seed: 37, x0: 200, y0: 410, x1: 1820, yTop: 40, yBot: 790, depth: 11, p3: 0.05, xp: 0.85, yj: 0.5 });
  TREE_COVER = makeTree({ seed: 3, x0: 1130, y0: 540, x1: 1890, yTop: 40, yBot: 1040, depth: 8, yj: 0.45 });
  PATH_COVER = pathOf(TREE_COVER, 9, 0.3);

  // a smooth blob filled with "continuum" dots, plus two finite covers
  const bc = [880, 420];
  const br = th => 250 * (1 + 0.16 * Math.sin(3 * th + 0.5) + 0.09 * Math.cos(5 * th) + 0.06 * Math.sin(2 * th + 1.3));
  const inside = (x, y) => { const dx = (x - bc[0]) / 1.45, dy = y - bc[1]; return Math.hypot(dx, dy) < br(Math.atan2(dy, dx)); };
  const R2 = rng(5), pts = [];
  while (pts.length < 4200) { const x = bc[0] + (R2() - 0.5) * 900, y = bc[1] + (R2() - 0.5) * 640; if (inside(x, y)) pts.push([x, y]); }
  const outline = [];
  for (let i = 0; i <= 200; i++) { const th = (i / 200) * Math.PI * 2, r = br(th); outline.push([bc[0] + Math.cos(th) * r * 1.45, bc[1] + Math.sin(th) * r]); }
  BLOB = { pts, outline };
  const cover = eps => {
    const out = [], s = eps * Math.sqrt(3);
    for (let row = 0, y = bc[1] - 340; y < bc[1] + 340; row++, y += s * 0.866) {
      for (let x = bc[0] - 470 + (row % 2) * s / 2; x < bc[0] + 470; x += s) {
        if (pts.some(p => (p[0] - x) ** 2 + (p[1] - y) ** 2 < eps * eps)) out.push([x, y]);
      }
    }
    return out.sort((a, b) => a[0] - b[0]);
  };
  COVER1 = { eps: 60, c: cover(60) }; COVER2 = { eps: 27, c: cover(27) };

  // DAG of many paths converging to few end states
  const sizes = [1, 4, 7, 9, 7, 4, 2];
  const layers = sizes.map((n, i) => Array.from({ length: n }, (_, j) => ({ x: 260 + i * 233, y: 430 + (j - (n - 1) / 2) * 78, out: [], i, j })));
  for (let i = 0; i < layers.length - 1; i++) {
    const A = layers[i], Bn = layers[i + 1];
    A.forEach((a, j) => {
      const r = A.length > 1 ? j / (A.length - 1) : 0.5;
      Bn.forEach((b, k) => { const rb = Bn.length > 1 ? k / (Bn.length - 1) : 0.5; if (Math.abs(rb - r) <= 0.36) a.out.push(b); });
    });
  }
  const cnt = new Map(); layers[0][0].c = 1;
  for (const Ly of layers) for (const n of Ly) { n.c = n.c || 0; for (const m of n.out) m.c = (m.c || 0) + n.c; }
  const total = layers[layers.length - 1].reduce((s, n) => s + n.c, 0);
  const Rw = rng(13), walkers = [];
  for (let k = 0; k < 90; k++) { const path = [layers[0][0]]; let n = path[0]; while (n.out.length) { n = n.out[Math.floor(Rw() * n.out.length)]; path.push(n); } walkers.push({ path, off: Rw() * 6, sp: 0.8 + Rw() * 0.5 }); }
  DAG = { layers, total, walkers };

  // life lines for the "everyone's T is finite" scene
  const Rl = rng(21);
  LIFELINES = Array.from({ length: 13 }, (_, i) => ({ y: 170 + i * 50, len: i === 6 ? 150 : 520 + Rl() * 900, ph: Rl() * 6 }));
  // faint alternative life paths (probability-zero scene)
  const Rc = rng(31);
  CURVES = Array.from({ length: 46 }, () => ({ a: [Rc(), Rc(), Rc()].map(v => 40 + v * 90), f: [Rc(), Rc(), Rc()].map((v, i) => (i + 1) * (0.6 + v * 0.8)), p: [Rc(), Rc(), Rc()].map(v => v * 6.28), y: 430 + (Rc() - 0.5) * 160 }));
}
const curveY = (c, u) => c.y + c.a[0] * Math.sin(c.f[0] * u * 3 + c.p[0]) * 0.8 + c.a[1] * Math.sin(c.f[1] * u * 3 + c.p[1]) * 0.6 + c.a[2] * Math.sin(c.f[2] * u * 3 + c.p[2]) * 0.4;

// ---------- reusable visuals ----------
function drawZoomLine(ctx, z, a, y = 430) {
  const w = Math.pow(10, -z), c = lerp(0.5, V_ZOOM, smooth(z)), X = v => 960 + ((v - c) / w) * 1400;
  seg(ctx, 190, y, 1730, y, 2, WHITE, 0.55 * a);
  for (let m = Math.max(0, Math.floor(z)); m <= Math.floor(z) + 3; m++) {
    const s = Math.pow(10, -m), px = (s / w) * 1400; if (px < 5) continue;
    const al = smooth((px - 5) / 60) * a, h = clamp(5 + Math.log10(px) * 9, 5, 34);
    const lab = smooth((px - 260) / 200) * a;
    let k0 = Math.ceil((c - w * 0.62) / s), k1 = Math.floor((c + w * 0.62) / s);
    for (let k = k0; k <= k1; k++) {
      const v = k * s, x = X(v); if (x < 190 || x > 1730) continue;
      seg(ctx, x, y - h / 2, x, y + h / 2, 1.5, WHITE, 0.7 * al);
      if (lab > 0.01) text(ctx, m === 0 ? String(Math.round(v)) : v.toFixed(m), x, y + 44, { size: 22, f: SANS, w: 300, c: COL.dim, a: lab });
    }
  }
  const px = X(V_ZOOM);
  light(ctx, px, y, 5, GOLD, a);
  return px;
}
function drawLogAxis(ctx, o) {
  const x0 = 180, x1 = 1740, y = o.y || 610, R = o.R || 90, X = e => x0 + (e / R) * (x1 - x0), a = o.a ?? 1;
  seg(ctx, x0, y, x1 + 20, y, 2, WHITE, 0.5 * a);
  const step = R <= 100 ? 10 : Math.pow(10, Math.floor(Math.log10(R / 4)));
  for (let e = 0; e <= R + 1e-9; e += step) seg(ctx, X(e), y - 8, X(e), y + 8, 1.5, WHITE, 0.45 * a * (R <= 100 ? 1 : 0.6));
  return X;
}

// ============================================================
const SCENES = [
  // ---------------- intro ----------------
  {
    id: 'intro', mood: 'wonder', fin: 1.5, pre: 0.4,
    beats: [
      { d: 2.0 },
      { text: '每一个瞬间，你都站在一个分岔口。' },
      { text: '说出口，或者沉默。转身，或者留下。' },
      { text: '博尔赫斯写过一座「小径分岔的花园」，' },
      { text: '在那里，时间同时通向无数个未来。' },
      { text: '如果把一生所有的分岔都画出来，' },
      { text: '它们，是无穷无尽的吗？', extra: 0.8 },
    ],
    draw(t, S, ctx) {
      const g = 9 * eio(P(t, S.b[1] - 0.3, S.e[5] - S.b[1] + 0.3));
      const dim = 1 - 0.35 * smooth(P(t, S.b[6], 1.5));
      const pulse = 0.85 + 0.15 * Math.sin(t * 2.2);
      light(ctx, 230, 470, 5, GOLD, smooth(P(t, 0.4, 1.4)) * pulse);
      drawTree(ctx, TREE_A, { g, a: dim, tips: 1 });
    },
  },
  {
    id: 'title', dur: 6.5, fin: 1.0, fout: 1.0,
    draw(t, S, ctx) { drawTree(ctx, TREE_A, { a: 0.28 * smooth(P(t, 0, 1.5)) }); },
    over(t, S, ctx) {
      const u = eout(P(t, 0.4, 1.6)), v = eout(P(t, 1.3, 1.6));
      text(ctx, '人生的分岔，可以数得清吗？', 960, 480 + (1 - u) * 14, { size: 88, w: 700, a: u, sh: 30, ls: 4 });
      seg(ctx, 960 - 90 * v, 572, 960 + 90 * v, 572, 1.5, GOLD, 0.6 * v);
      text(ctx, '关于无穷、有限，与最后一程', 960, 630, { size: 36, f: SANS, w: 300, c: COL.dim, a: v, ls: 10, sh: 20 });
    },
    accents: [{ at: 0.4, kind: 'bell' }],
  },

  // ---------------- chapter 1 ----------------
  { id: 'c1', card: { num: '壹', title: '同一个问题，三个答案' }, dur: 4.2, mood: 'think', accents: [{ at: 0.2, kind: 'bell' }] },
  {
    id: 's1a',
    beats: [
      { text: '第一种答案：*连续*。' },
      { text: '如果时间是连续的，\n每个选择都可以精确到无穷多位小数——' },
      { text: '几点出门，走多快，说话时停顿多久。' },
      { text: '一生做 n 次选择，每次都有实数那么多种取法：' },
      { text: '可能的人生，和实数一样多——*不可数无穷*。', extra: 0.6 },
    ],
    draw(t, S, ctx) {
      const z = 9.3 * eio(P(t, S.b[0] + 1.0, S.e[2] - S.b[0] - 1.0));
      const a = smooth(P(t, 0.2, 1)) * (1 - smooth(P(t, S.b[3] - 0.3, 0.8)));
      if (a > 0.003) S._px = drawZoomLine(ctx, z, a);
      S._z = z;
    },
    over(t, S, ctx) {
      const a = smooth(P(t, S.b[0] + 1.0, 0.8)) * (1 - smooth(P(t, S.b[3] - 0.3, 0.8)));
      const nd = Math.max(1, Math.floor(S._z || 0) + 1);
      if (a > 0.003) text(ctx, '0.' + DIG.slice(0, nd) + '...', S._px || 960, 300, { size: 56, c: COL.gold, a, sh: 20 });
      const b = smooth(P(t, S.b[3] + 0.5, 1)) * (1 - smooth(P(t, S.dur - 0.8, 0.6)));
      tex(ctx, 'cont', 960, 420, { a: b, s: 1.05 });
      const c = smooth(P(t, S.b[4], 1));
      texLabel(ctx, 'aleph0', '：自然数的个数', 760, 590, { a: c * b, size: 30 });
      texLabel(ctx, 'c2aleph0', '：实数的个数', 1160, 590, { a: c * b, size: 30 });
    },
  },
  {
    id: 's1b',
    beats: [
      { text: '第二种答案：*离散*。' },
      { text: '把人生看成一棵决策树：每一步，只有有限个选项。' },
      { text: '有限步，有限个分支，路径的数量就是有限的。', extra: 1.0 },
    ],
    draw(t, S, ctx) {
      const g = 4 * eio(P(t, S.b[0] + 0.2, S.e[1] - S.b[0] - 0.4));
      drawTree(ctx, TREE_K3, { g, tips: 1, w: 3, wm: 1.2, da: () => 0.8 });
      light(ctx, 420, 430, 5, GOLD, smooth(P(t, 0.2, 0.8)));
      const leaves = TREE_K3.byD[4].map(e => e.c).sort((a, b) => a.y - b.y);
      const k = Math.floor(81 * P(t, S.b[2] + 0.3, 2.6));
      for (let i = 0; i < k; i++) glow(ctx, leaves[i].x, leaves[i].y, 14, GOLD, 0.9);
      S._k = k;
    },
    over(t, S, ctx) {
      const a = smooth(P(t, S.b[2], 0.8));
      tex(ctx, 'n3', 1600, 300, { a });
      text(ctx, String(S._k || 0), 1600, 440, { size: 120, w: 700, c: COL.gold, a });
      text(ctx, '条路径', 1600, 540, { size: 34, f: SANS, w: 300, c: COL.dim, a });
    },
  },
  {
    id: 's1c',
    beats: [
      { text: '第三种答案，来自*物理*。' },
      { text: '一块有限大小、有限能量的物质，\n能装下的信息是有上限的。' },
      { text: '这叫 Bekenstein 界。' },
      { text: '代入一颗人脑：约 1.5 千克，半径约 0.1 米，' },
      { text: '它能处于的可区分状态，最多约有这么多种：', extra: 0.8 },
      { text: '严格地说：*有限*。' },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.3, 1.2)), cx = 960, cy = 520, r = 200;
      const flash = 1 + 0.6 * env(t, S.b[5], S.b[5] + 1.6, 0.3, 1.2);
      glow(ctx, cx, cy, r * 1.9, BLUE, 0.12 * a * flash);
      circle(ctx, cx, cy, r, 2, BLUE, 0.65 * a * flash);
      ctx.save(); ctx.globalAlpha = a;
      const cell = 13, tick = Math.floor(t * 7);
      for (let gy = cy - r; gy < cy + r; gy += cell) for (let gx = cx - r; gx < cx + r; gx += cell) {
        if ((gx - cx + 6) ** 2 + (gy - cy + 6) ** 2 > (r - 10) ** 2) continue;
        const id = gx * 7.13 + gy * 3.71, on = hash(id + Math.floor(tick + hash(id) * 9) * 0.37) > 0.55;
        ctx.fillStyle = on ? rgba(BLUE, 0.75 * flash) : rgba(BLUE, 0.12);
        ctx.fillRect(gx, gy, 4, 4);
      }
      ctx.restore();
      const u = smooth(P(t, S.b[3], 1));
      seg(ctx, cx, cy, cx, cy + r * u, 2, GOLD, 0.8 * u, [6, 6]);
    },
    over(t, S, ctx) {
      tex(ctx, 'bek', 960, 175, { a: smooth(P(t, S.b[1] + 0.6, 1)), sh: 12 });
      const u = smooth(P(t, S.b[3] + 0.5, 1));
      tex(ctx, 'r01', 960, 762, { a: u, sh: 10 });
      tex(ctx, 'mkg', 560, 520, { a: u, sh: 10 });
      const v = smooth(P(t, S.b[4] + 0.2, 1));
      tex(ctx, 'brainI', 1490, 470, { a: v, sh: 10 });
      tex(ctx, 'brainN', 1490, 560, { a: v, sh: 10 });
    },
  },
  {
    id: 's1d',
    beats: [
      { text: '同一个问题，三个答案。', extra: 1.0 },
      { text: '到底哪一个，才是对的？', extra: 0.6 },
    ],
    draw(t, S, ctx) {
      const xs = [480, 960, 1440];
      xs.forEach((x, i) => {
        const a = smooth(P(t, 0.3 + i * 0.5, 1));
        if (i === 0) { seg(ctx, x - 110, 330, x + 110, 330, 2, WHITE, 0.5 * a); for (let k = 0; k <= 40; k++) seg(ctx, x - 110 + k * 5.5, 322, x - 110 + k * 5.5, 338, 1, WHITE, 0.35 * a); }
        if (i === 1) { const yy = (d, k) => (d === 0 ? 330 : 330 + ((k + 0.5) / 2 ** d - 0.5) * 150); for (let d = 0; d < 3; d++) for (let k = 0; k < 2 ** d; k++) { const x0 = x - 105 + d * 70; for (const c of [2 * k, 2 * k + 1]) seg(ctx, x0, yy(d, k), x0 + 70, yy(d + 1, c), 1.6, GOLD, 0.75 * a); } }
        if (i === 2) { circle(ctx, x, 330, 70, 2, BLUE, 0.7 * a); glow(ctx, x, 330, 130, BLUE, 0.15 * a); }
      });
    },
    over(t, S, ctx) {
      const xs = [480, 960, 1440], h = ['连续', '离散', '物理'], r = ['不可数无穷', '有限', '严格有限'];
      xs.forEach((x, i) => {
        const a = smooth(P(t, 0.3 + i * 0.5, 1));
        text(ctx, h[i], x, 490, { size: 58, w: 700, c: COL.gold, a });
        text(ctx, r[i], x, 570, { size: 40, a, ls: 4 });
      });
    },
  },

  // ---------------- chapter 2 ----------------
  { id: 'c2', card: { num: '贰', title: '无穷在哪里塌缩' }, dur: 4.2, accents: [{ at: 0.2, kind: 'bell' }] },
  {
    id: 's2a',
    beats: [
      { text: '关键在一个词：*可区分*。' },
      { text: '十点整出门，和晚 10⁻³⁰ 秒出门——' },
      { text: '在数学上，是两个不同的点；', extra: 0.4 },
      { text: '在你的人生里，没有任何不同。', extra: 0.6 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1)), y = 560;
      seg(ctx, 200, y, 1720, y, 2, WHITE, 0.5 * a);
      for (let m = -4; m <= 4; m++) { const x = 960 + m * 190; seg(ctx, x, y - 14, x, y + 14, 1.5, WHITE, 0.5 * a); }
      for (let m = -40; m <= 40; m++) { const x = 960 + m * 19; if (x > 200 && x < 1720) seg(ctx, x, y - 5, x, y + 5, 1, WHITE, 0.25 * a); }
      const lens = env(t, S.b[1] + 0.3, S.e[2] + 0.2, 0.8, 0.8);
      const merge = smooth(P(t, S.b[3] + 0.3, 1.4));
      light(ctx, 960, y, 5 + 2 * merge * (1 + Math.sin(t * 3)) * 0.5, GOLD, a);
      if (lens > 0.003) {
        const lx = 960, ly = 300, R = 170;
        seg(ctx, 960, y - 10, lx, ly + R, 1.2, GOLD, 0.5 * lens, [4, 6]);
        circle(ctx, lx, ly, R, 2, GOLD, 0.7 * lens, 'rgba(8,10,18,0.9)');
        ctx.save(); ctx.beginPath(); ctx.arc(lx, ly, R - 2, 0, 7); ctx.clip();
        seg(ctx, lx - R, ly, lx + R, ly, 1.5, WHITE, 0.45 * lens);
        for (let m = -12; m <= 12; m++) seg(ctx, lx + m * 14, ly - 6, lx + m * 14, ly + 6, 1, WHITE, 0.25 * lens);
        const sep = 110 * smooth(P(t, S.b[2], 1.2));
        light(ctx, lx - sep / 2, ly, 5, GOLD, lens);
        light(ctx, lx + sep / 2, ly, 5, BLUE, lens);
        ctx.restore();
      }
    },
    over(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1)), y = 560;
      ['09:58', '09:59', '10:00', '10:01', '10:02'].forEach((s, i) => text(ctx, s, 960 + (i - 2) * 380, y + 46, { size: 26, f: SANS, w: 300, c: i === 2 ? COL.gold : COL.dim, a }));
      const lens = env(t, S.b[1] + 0.3, S.e[2] + 0.2, 0.8, 0.8);
      tex(ctx, 'dt', 960, 372, { a: lens * smooth(P(t, S.b[2], 1)), s: 0.72 });
    },
  },
  {
    id: 's2b',
    beats: [
      { text: '所以，规定一个最小分辨率 *ε*：\n差别小于 ε 的，算同一个选择。' },
      { text: '人生的参数是有界的：\n时间有限，力气有限，能去的地方也有限。' },
      { text: '而有界闭集有一个性质：无论 ε 多小，' },
      { text: '它都能被*有限个*半径为 ε 的小圆盖住。', extra: 0.8 },
      { text: '数学上，这叫*紧致*。', extra: 1.4 },
      { text: '于是，「不可数无穷」在可区分性面前，\n塌缩成了一个有限的数。', extra: 0.8 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1.2));
      const fadeDots = 1 - smooth(P(t, S.b[5], 1.6));
      ctx.save(); ctx.fillStyle = rgba(BLUE, 0.55 * a * fadeDots);
      for (const p of BLOB.pts) ctx.fillRect(p[0] - 1, p[1] - 1, 2, 2);
      ctx.restore();
      ctx.save(); ctx.beginPath(); BLOB.outline.forEach((p, i) => (i ? ctx.lineTo(p[0], p[1]) : ctx.moveTo(p[0], p[1])));
      ctx.strokeStyle = rgba(BLUE, 0.6 * a * fadeDots); ctx.lineWidth = 2; ctx.stroke(); ctx.restore();
      const c1 = env(t, S.b[3], S.b[4] + 1.0, 0.4, 1.0), c2 = smooth(P(t, S.b[4] + 0.6, 1.2));
      const collapse = smooth(P(t, S.b[5] + 0.4, 1.8));
      const drawCover = (C, alpha, t0) => {
        C.c.forEach((p, i) => {
          const u = smooth(P(t, t0 + (i / C.c.length) * 1.4, 0.35)); if (u <= 0) return;
          const r = C.eps * u * (1 - collapse);
          if (r > 0.5) circle(ctx, p[0], p[1], r, 1.2, GOLD, 0.6 * alpha);
          if (collapse > 0) light(ctx, p[0], p[1], 2.2, GOLD, alpha * collapse * 0.9);
        });
      };
      if (c1 > 0.003) drawCover(COVER1, c1, S.b[3]);
      if (c2 > 0.003) drawCover(COVER2, c2, S.b[4] + 0.6);
      // epsilon gauge
      const ga = env(t, S.b[0] + 0.5, S.b[5] + 0.6, 0.8, 0.8);
      const eps = lerp(COVER1.eps, COVER2.eps, smooth(P(t, S.b[4] + 0.4, 1.2)));
      circle(ctx, 1620, 640, eps, 1.5, GOLD, 0.8 * ga);
      seg(ctx, 1620, 640, 1620 + eps, 640, 1.5, GOLD, 0.8 * ga);
    },
    over(t, S, ctx) {
      const ga = env(t, S.b[0] + 0.5, S.b[5] + 0.6, 0.8, 0.8);
      text(ctx, 'ε', 1620 + 18, 616, { size: 30, c: COL.gold, a: ga, f: SANS });
      const k1 = Math.round(COVER1.c.length * P(t, S.b[3], 1.75)), k2 = Math.round(COVER2.c.length * P(t, S.b[4] + 0.6, 1.75));
      const n = t < S.b[4] + 0.6 ? k1 : Math.max(k2, 0);
      const ca = smooth(P(t, S.b[3], 0.6));
      text(ctx, `N(ε) = ${n}`, 1620, 300, { size: 46, c: COL.gold, a: ca, f: SERIF });
      tex(ctx, 'Neps', 1620, 380, { a: smooth(P(t, S.b[4], 1)), s: 0.8 });
      text(ctx, '紧致 · compact', 1620, 455, { size: 28, f: SANS, w: 300, c: COL.dim, a: smooth(P(t, S.b[4], 1)) });
    },
  },

  // ---------------- chapter 3 ----------------
  { id: 'c3', card: { num: '叁', title: '有限，但是' }, dur: 4.2, accents: [{ at: 0.2, kind: 'bell' }] },
  {
    id: 's3a',
    beats: [
      { text: '有限，不代表小。' },
      { text: '只做「是」与「否」的选择。\n每多一次，可能性就翻一倍。', extra: 0.6 },
      { text: '63 次，比地球上所有海滩的沙粒还多。', extra: 0.6 },
      { text: '80 次，比可观测宇宙中的恒星还多。', extra: 0.6 },
      { text: '266 次，比可观测宇宙中的原子还多。', extra: 1.4 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.3, 1));
      const X = drawLogAxis(ctx, { a });
      const n = S._n = keys(t, [[S.b[1], 0], [S.e[1], 20], [S.b[2] + 0.2, 20], [S.b[2] + 2.4, 63], [S.b[3] + 0.2, 63], [S.b[3] + 1.6, 80], [S.b[4] + 0.2, 80], [S.b[4] + 3.4, 266]]);
      const e = n * Math.LOG10E * Math.LN2;
      const cx = X(e);
      seg(ctx, X(0), 610, cx, 610, 3, GOLD, 0.9 * a);
      light(ctx, cx, 610, 6, GOLD, a);
      [[18.9, true], [24, false], [80, true]].forEach(([me, up], i) => {
        const lit = e >= me ? 1 : 0.25, y2 = up ? 560 : 660;
        seg(ctx, X(me), 610, X(me), y2, 1.5, lit > 0.5 ? GOLD : WHITE, 0.6 * a * lit + 0.1);
        if (e >= me) glow(ctx, X(me), 610, 40, GOLD, 0.6 * a);
      });
    },
    over(t, S, ctx) {
      const a = smooth(P(t, 0.3, 1)), X = e => 180 + (e / 90) * 1560;
      for (let e = 0; e <= 90; e += 10) rich(ctx, [['10', String(e)]], X(e), 650, { size: 22, f: SANS, w: 300, c: COL.dim, a: a * (e === 20 || e === 80 ? 0 : 1) });
      const n = Math.floor(S._n || 0), e = n * Math.LOG10E * Math.LN2;
      const mk = [
        { e: 18.9, label: '地球所有海滩的沙粒', v: ['约 ', ['10', '19']], up: true },
        { e: 24, label: '可观测宇宙的恒星', v: ['约 ', ['10', '22'], ' ~ ', ['10', '24']], up: false },
        { e: 80, label: '可观测宇宙的原子', v: ['约 ', ['10', '80']], up: true },
      ];
      mk.forEach(m => {
        const lit = e >= m.e, c = lit ? COL.gold : COL.dim, al = a * (lit ? 1 : 0.6);
        const y = m.up ? 500 : 720;
        text(ctx, m.label, X(m.e), y, { size: 28, c, a: al });
        rich(ctx, m.v, X(m.e), y + (m.up ? 40 : 40), { size: 24, f: SANS, w: 300, c, a: al });
      });
      const b = smooth(P(t, S.b[1], 0.8));
      text(ctx, `n = ${n}`, 960, 230, { size: 84, w: 700, c: COL.gold, a: b });
      const ex = Math.floor(e + 1e-9), man = Math.pow(10, e - ex);
      if (n < 20) rich(ctx, [['2', String(n)], ' = ' + Math.pow(2, n).toLocaleString('en-US')], 960, 330, { size: 50, a: b });
      else rich(ctx, [['2', String(n)], ` ≈ ${man.toFixed(1)} × `, ['10', String(ex)]], 960, 330, { size: 50, a: b });
    },
  },
  {
    id: 's3b',
    beats: [
      { text: '而一个人清醒时，如果每秒做一次选择，' },
      { text: '一生大约是 *17 亿*次。', extra: 0.6 },
      { text: '可能的人生，约有这么多种——', extra: 1.6 },
      { text: '这个数写出来，有*五亿位*。', extra: 1.0 },
      { text: '有限。但在任何意义上，它都等于无穷。', extra: 1.0 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0, 0.8));
      const u = eio(P(t, S.b[2] + 0.3, 3.6));
      const R = 90 * Math.pow(5.06e8 / 90, u);
      const X = drawLogAxis(ctx, { a, R });
      const cx = lerp(X(80.07), 1740, u);
      seg(ctx, 180, 610, cx, 610, 3, GOLD, 0.9 * a);
      light(ctx, cx, 610, 6 + 3 * u, GOLD, a);
      [18.9, 24, 80].forEach(me => { seg(ctx, X(me), 610, X(me), 560, 1.5, GOLD, 0.5 * a); glow(ctx, X(me), 610, 30, GOLD, 0.4 * a); });
      // a rain of digits for "five hundred million digits"
      const r = smooth(P(t, S.b[3], 1.5)) * (1 - smooth(P(t, S.dur - 1.2, 1)));
      if (r > 0.003) {
        ctx.save(); ctx.font = `300 22px ${SANS}`; ctx.fillStyle = rgba(GOLD, 0.055 * r); ctx.textBaseline = 'middle';
        const off = (t * 26) % 34;
        for (let row = 0; row < 34; row++) {
          const idx = row + Math.floor((t * 26) / 34);
          let s = ''; for (let k = 0; k < 200; k++) s += Math.floor(hash(idx * 157 + k) * 10);
          ctx.fillText(s, -10, row * 34 - off + 17);
        }
        ctx.restore();
      }
    },
    over(t, S, ctx) {
      const a = smooth(P(t, 0, 0.8));
      const u = eio(P(t, S.b[2] + 0.3, 3.6));
      const la = 1 - smooth(u * 6);
      const X = e => 180 + (e / 90) * 1560;
      for (let e = 0; e <= 90; e += 10) rich(ctx, [['10', String(e)]], X(e), 650, { size: 22, f: SANS, w: 300, c: COL.dim, a: a * la });
      rich(ctx, ['原子 ', ['10', '80']], lerp(X(80), 180 + 80 / 5.06e8 * 1560, u), 520, { size: 24, f: SANS, w: 300, c: COL.gold, a });
      tex(ctx, 'big10', 1700, 680, { a: smooth(P(t, S.b[2] + 3.2, 1)) });
      const f1 = env(t, S.b[0] + 0.6, S.b[2] + 0.4, 0.8, 0.6);
      tex(ctx, 'life', 960, 250, { a: f1, sh: 12 });
      tex(ctx, 'lifeN', 960, 250, { a: smooth(P(t, S.b[2] + 0.4, 1)), sh: 12, s: 1.1 });
    },
  },

  // ---------------- chapter 4 ----------------
  { id: 'c4', card: { num: '肆', title: '无穷的另一面' }, dur: 4.2, accents: [{ at: 0.2, kind: 'bell' }] },
  {
    id: 's4a',
    beats: [
      { text: '但「无穷」，不等于「一切皆有可能」。', extra: 0.6 },
      { text: '偶数有无穷多个，可没有一个是 3。', extra: 1.0 },
      { text: '选择的多少，和选择的边界，是两回事。', extra: 0.8 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1)), y = 500;
      seg(ctx, 140, y, 1760, y, 2, WHITE, 0.45 * a);
      for (let k = 0; k <= 14; k++) {
        const x = 180 + k * 110, u = smooth(P(t, 0.4 + k * 0.12, 0.6)) * a;
        if (k % 2 === 0) light(ctx, x, y, 6, GOLD, u);
        else if (k === 3) {
          const h = smooth(P(t, S.b[1] + 0.4, 0.8));
          circle(ctx, x, y, 9 + 4 * h * (0.5 + 0.5 * Math.sin(t * 4)), 2, mixc(WHITE, RED, h), u * (0.4 + 0.6 * h));
        } else circle(ctx, x, y, 7, 1.5, WHITE, u * 0.3);
      }
    },
    over(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1)), y = 500;
      for (let k = 0; k <= 14; k++) {
        const x = 180 + k * 110, u = smooth(P(t, 0.4 + k * 0.12, 0.6)) * a;
        const hot = k === 3 ? smooth(P(t, S.b[1] + 0.4, 0.8)) : 0;
        text(ctx, String(k), x, y + 50, { size: 30, f: SANS, w: 300, c: k % 2 === 0 ? COL.gold : (hot > 0 ? '#d67a68' : COL.dim), a: u * (k % 2 === 0 ? 1 : 0.5 + 0.5 * hot) });
      }
      text(ctx, '…', 1790, y, { size: 40, c: COL.dim, a });
      tex(ctx, 'not3', 960, 320, { a: smooth(P(t, S.b[1] + 0.8, 1)) });
    },
  },
  {
    id: 's4b',
    beats: [
      { text: '再看一个奇怪的事实。' },
      { text: '在 [0, 1] 里随机取一个数，\n取到任何一个特定的数，概率都是零。', extra: 0.6 },
      { text: '但最后，总会取到某一个。', extra: 1.0 },
      { text: '在连续模型里，你正在过的这一生，就是这样一个事件：' },
      { text: '概率为零，却真实地发生了。', extra: 1.6 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1)) * (1 - smooth(P(t, S.b[3], 1.0)));
      const y = 560, x0 = 560, x1 = 1360;
      if (a > 0.003) {
        ctx.save(); ctx.fillStyle = rgba(BLUE, 0.08 * a); ctx.fillRect(x0, y - 150, x1 - x0, 150); ctx.restore();
        seg(ctx, x0, y - 150, x1, y - 150, 1.5, BLUE, 0.6 * a);
        seg(ctx, x0, y - 150, x0, y, 1, BLUE, 0.3 * a); seg(ctx, x1, y - 150, x1, y, 1, BLUE, 0.3 * a);
        seg(ctx, x0 - 60, y, x1 + 60, y, 2, WHITE, 0.5 * a);
        const xv = lerp(x0, x1, 0.71393);
        const f = P(t, S.b[2], 1.1), yy = f < 1 ? lerp(140, y, f * f) : y;
        if (f > 0) light(ctx, xv, yy, 6, GOLD, a);
        const rip = P(t, S.b[2] + 1.1, 1.4);
        if (rip > 0 && rip < 1) circle(ctx, xv, y, 10 + 60 * rip, 1.5, GOLD, a * (1 - rip));
      }
      const b = smooth(P(t, S.b[3], 1.2));
      if (b > 0.003) {
        const u = P(t, S.b[3] + 0.4, S.e[4] - S.b[3] - 1.4);
        ctx.save(); ctx.lineWidth = 1.2; ctx.strokeStyle = rgba(WHITE, 0.09 * b);
        for (const c of CURVES) { ctx.beginPath(); for (let i = 0; i <= 80; i++) { const v = i / 80, x = 160 + v * 1600, yy = curveY(c, v); i ? ctx.lineTo(x, yy) : ctx.moveTo(x, yy); } ctx.stroke(); }
        const c = CURVES[7]; ctx.beginPath(); ctx.lineWidth = 3; ctx.strokeStyle = rgba(GOLD, 0.95 * b);
        const n = Math.floor(160 * u); let px = 160, py = curveY(c, 0);
        for (let i = 0; i <= n; i++) { const v = i / 160; px = 160 + v * 1600; py = curveY(c, v); i ? ctx.lineTo(px, py) : ctx.moveTo(px, py); }
        ctx.stroke(); ctx.restore();
        light(ctx, px, py, 6, GOLD, b);
      }
    },
    over(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1)) * (1 - smooth(P(t, S.b[3], 1.0))), y = 560;
      text(ctx, '0', 560, y + 44, { size: 30, f: SANS, w: 300, c: COL.dim, a });
      text(ctx, '1', 1360, y + 44, { size: 30, f: SANS, w: 300, c: COL.dim, a });
      tex(ctx, 'p0', 960, 240, { a: a * smooth(P(t, S.b[1] + 0.8, 1)) });
      text(ctx, '0.71393…', lerp(560, 1360, 0.71393), y + 44, { size: 30, c: COL.gold, a: a * smooth(P(t, S.b[2] + 1.1, 0.8)) });
      tex(ctx, 'p0', 960, 200, { a: smooth(P(t, S.b[4], 1)) * 0.9, s: 0.9 });
    },
  },
  {
    id: 's4c',
    beats: [
      { text: '还有：很多不同的路，会通往同一个地方。', extra: 0.6 },
      { text: '无论走哪条街去上班，最后都到了公司。', extra: 0.6 },
      { text: '路径的数量，\n远远多于你最终可能成为的人的数量。', extra: 1.4 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1.2));
      ctx.save(); ctx.lineWidth = 1.3;
      for (const Ly of DAG.layers) for (const n of Ly) for (const m of n.out) {
        ctx.strokeStyle = rgba(WHITE, 0.16 * a); ctx.beginPath(); ctx.moveTo(n.x, n.y);
        ctx.bezierCurveTo(n.x + 110, n.y, m.x - 110, m.y, m.x, m.y); ctx.stroke();
      }
      ctx.restore();
      for (const Ly of DAG.layers) for (const n of Ly) dot(ctx, n.x, n.y, 3.5, WHITE, 0.5 * a);
      const end = DAG.layers[DAG.layers.length - 1];
      const fa = smooth(P(t, S.b[0] + 0.8, 1.5)) * a;
      for (const w of DAG.walkers) {
        const T = (t * 0.55 * w.sp + w.off) % 7.5; if (T > 6) continue;
        const i = Math.floor(T), f = T - i, n = w.path[i], m = w.path[i + 1];
        const fe = smooth(f), u = 1 - fe;
        const x = u * u * u * n.x + 3 * u * u * fe * (n.x + 110) + 3 * u * fe * fe * (m.x - 110) + fe * fe * fe * m.x;
        const y = u * u * u * n.y + 3 * u * u * fe * n.y + 3 * u * fe * fe * m.y + fe * fe * fe * m.y;
        light(ctx, x, y, 2.5, GOLD, fa * 0.8);
      }
      for (const n of end) light(ctx, n.x, n.y, 8, GOLD, fa * (0.8 + 0.2 * Math.sin(t * 2 + n.j)));
    },
    over(t, S, ctx) {
      const a = smooth(P(t, S.b[2], 1));
      text(ctx, `不同的路：${DAG.total} 条`, 560, 140, { size: 40, c: COL.text, a });
      text(ctx, `终点：${DAG.layers[DAG.layers.length - 1].length} 个`, 1360, 140, { size: 40, c: COL.gold, a });
    },
  },

  // ---------------- chapter 5 ----------------
  { id: 'c5', card: { num: '伍', title: '坍缩' }, dur: 4.6, mood: 'heavy', accents: [{ at: 0.2, kind: 'low' }] },
  {
    id: 's5a',
    beats: [
      { text: '现在，想一个更沉重的情形。', extra: 0.6 },
      { text: '一个人，被诊断为绝症晚期。', extra: 1.0 },
      { text: '这个人未来的选择，会在那一瞬间坍缩吗？', extra: 1.0 },
      { text: '可能的人生路径，大约有这么多条：', extra: 0.8 },
      { text: 'b：每一步能选的选项。T：剩下的时间。\nΔt：两次选择之间的间隔。', extra: 1.2 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.3, 1.5)) * (1 - smooth(P(t, S.b[3], 1)));
      const shrink = 1 - 0.45 * smooth(P(t, S.b[2] + 0.4, 2.2));
      const hb = Math.max(0, Math.sin(t * 5.2)) ** 12 + 0.6 * Math.max(0, Math.sin(t * 5.2 - 0.9)) ** 12;
      for (let k = 0; k < 40; k++) {
        const th = (k / 40) * Math.PI * 2 + hash(k) * 0.1, r0 = 30, r1 = (180 + hash(k + 9) * 260) * shrink;
        const fl = 0.5 + 0.5 * Math.sin(t * 1.3 + k);
        seg(ctx, 960 + Math.cos(th) * r0, 430 + Math.sin(th) * r0, 960 + Math.cos(th) * r1, 430 + Math.sin(th) * r1, 1.2, GOLD, a * 0.25 * fl * (1 - 0.5 * smooth(P(t, S.b[1] + 0.5, 1.5))));
      }
      light(ctx, 960, 430, 6 + 3 * hb, GOLD, a);
    },
    over(t, S, ctx) {
      const a = smooth(P(t, S.b[3] + 0.4, 1));
      tex(ctx, 'nform', 960, 420, { a, s: 1.2, sh: 20 });
      const b = smooth(P(t, S.b[4], 1));
      const items = [['nb', '每一步的选项数'], ['nT', '剩下的时间'], ['ndt', '选择的间隔']];
      items.forEach(([k, l], i) => texLabel(ctx, k, '　' + l, 560 + i * 400, 620, { a: b * smooth(P(t, S.b[4] + i * 0.4, 0.8)), size: 32, s: 0.9 }));
    },
  },
  {
    id: 's5b',
    beats: [
      { text: '确诊，让两个参数同时变小。' },
      { text: 'T 从几十年，变成几个月。\n它在指数上——这是指数级的缩减。', extra: 1.6 },
      { text: '疾病、疼痛和治疗，也让 b 变小。', extra: 1.0 },
      { text: '远处的分支——二十年后的事业，看着孩子长大——\n被整片剪掉了。', extra: 2.0 },
      { text: '这种坍缩，是真实的。', extra: 1.6 },
    ],
    accents: [{ at: S => S.b[1] + 0.3, kind: 'low' }],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0, 1)), ay = 800;
      const xT = lerp(1760, 430, eio(P(t, S.b[1] + 0.3, 3.4)));
      const bs = smooth(P(t, S.b[2] + 0.2, 1.8));
      const gone = 1 - smooth(P(t, S.b[1] + 3.5, 2.2));
      S._xT = xT;
      ctx.save(); ctx.beginPath(); ctx.rect(0, 0, xT, H); ctx.clip();
      drawTree(ctx, TREE_P, { a, per: e => { const k = e.c.drop ? 1 - bs : 1; return k > 0 ? { a: k } : null; } });
      ctx.restore();
      ctx.save(); ctx.beginPath(); ctx.rect(xT, 0, W - xT, H); ctx.clip();
      drawTree(ctx, TREE_P, {
        a: a * gone, per: e => {
          const f = clamp(0.2 + (e.c.x - xT) / 380);
          return f >= 1 ? null : { a: (1 - f) * 0.8, dy: f * f * 90, c: mixc(GOLD, RED, Math.min(1, f * 2.5)) };
        },
      });
      ctx.restore();
      light(ctx, 200, 410, 5, GOLD, a);
      seg(ctx, 200, ay, 1760, ay, 1.5, WHITE, 0.35 * a);
      const la = smooth(P(t, 0.3, 1));
      seg(ctx, xT, 50, xT, ay, 2, RED, 0.7 * la, [8, 8]);
      glow(ctx, xT, 50, 30, RED, 0.5 * la);
    },
    over(t, S, ctx) {
      const a = smooth(P(t, 0, 1));
      text(ctx, '现在', 200, 830, { size: 26, f: SANS, w: 300, c: COL.dim, a });
      text(ctx, '几十年后', 1760, 830, { size: 26, f: SANS, w: 300, c: COL.dim, a: a * (1 - smooth(P(t, S.b[1] + 1, 1.5)) * 0.6) });
      text(ctx, 'T', (S._xT || 1760) + 22, 60, { size: 40, c: '#d67a68', a: smooth(P(t, 0.3, 1)), f: SERIF, w: 700 });
      const g = [['二十年后的事业', 1250, 240], ['看着孩子长大', 1500, 560], ['一次迟来的重逢', 900, 140], ['老去', 1650, 390]];
      g.forEach(([s, x, y], i) => {
        const u = P(t, S.b[3] + 0.2 + i * 0.5, 4.5);
        const al = Math.sin(Math.PI * clamp(u)) * 0.75;
        text(ctx, s, x, y + 70 * u * u, { size: 34, c: COL.dim, a: al });
      });
    },
  },
  {
    id: 's5c',
    beats: [
      { text: '但它不会塌缩成「几乎没有」。', extra: 1.6 },
      { text: '就算只剩半年，只算清醒时每分钟一次选择——', extra: 0.4 },
      { text: '可能的路径，仍然有这么多：', extra: 1.6 },
      { text: '仍然远远多于宇宙中的原子。', extra: 0.8 },
      { text: '被剪掉的，是远处的分支。\n眼前这一段，依然多得数不完。', extra: 2.2 },
    ],
    draw(t, S, ctx) {
      const z = eio(P(t, S.b[0] + 0.3, 3.4));
      const s = lerp(1, (1760 - 200) / (430 - 200), z);
      const a0 = 1 - smooth(P(t, S.b[0] + 2.2, 1.6));
      ctx.save(); ctx.translate(200, 410); ctx.scale(s, s); ctx.translate(-200, -410);
      ctx.beginPath(); ctx.rect(0, 0, 430, H); ctx.clip();
      drawTree(ctx, TREE_P, { a: a0, scale: s, per: e => (e.c.drop ? null : { a: 1 }) });
      ctx.restore();
      const g = 11 * eio(P(t, S.b[0] + 0.7, 5.0));
      const warm = 0.75 + 0.25 * smooth(P(t, S.b[4], 2));
      drawTree(ctx, TREE_DENSE, { g, a: warm * smooth(P(t, S.b[0] + 0.6, 1.2)), tips: 0.6, w: 3.2, wm: 0.5, da: d => Math.max(0.22, 0.92 - 0.07 * d) });
      light(ctx, 200, 410, 6, GOLD, 1);
      const band = smooth(P(t, S.b[2], 1)) * (1 - smooth(P(t, S.b[4], 1)));
      if (band > 0.003) {
        ctx.save(); ctx.globalCompositeOperation = 'destination-out'; ctx.translate(960, 210); ctx.scale(1, 0.32);
        const rg = ctx.createRadialGradient(0, 0, 0, 0, 0, 640); rg.addColorStop(0, `rgba(0,0,0,${0.92 * band})`); rg.addColorStop(0.6, `rgba(0,0,0,${0.7 * band})`); rg.addColorStop(1, 'rgba(0,0,0,0)');
        ctx.fillStyle = rg; ctx.fillRect(-700, -700, 1400, 1400); ctx.restore();
      }
    },
    over(t, S, ctx) {
      const a = smooth(P(t, S.b[2] + 0.3, 1)) * (1 - smooth(P(t, S.b[4], 1)));
      tex(ctx, 'half', 960, 170, { a, s: 1.1, sh: 22 });
      tex(ctx, 'gg80', 960, 265, { a: a * smooth(P(t, S.b[3], 1)), sh: 16 });
    },
  },

  // ---------------- chapter 6 ----------------
  { id: 'c6', card: { num: '陆', title: '你知道了什么' }, dur: 4.2, accents: [{ at: 0.2, kind: 'bell' }] },
  {
    id: 's6a',
    beats: [
      { text: '从概率的角度看，确诊更像一次「条件化」。', extra: 0.4 },
      { text: '每个人的 T，本来就是有限的。\n这一点，从来没有变过。', extra: 0.8 },
      { text: '健康的人觉得未来无限，\n只是因为不知道 T 有多大，就默认它很大。', extra: 0.8 },
      { text: '确诊，把一条模糊而漫长的分布，\n变成了一条尖锐而短的分布。', extra: 1.8 },
      { text: '改变的不只是世界，还有——*你知道了什么*。', extra: 1.2 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1.2)), ox = 260, oy = 760, X = yr => ox + yr * 20.5;
      seg(ctx, ox, oy, 1720, oy, 2, WHITE, 0.5 * a); seg(ctx, ox, oy, ox, 170, 2, WHITE, 0.5 * a);
      for (let yr = 0; yr <= 70; yr += 10) seg(ctx, X(yr), oy - 8, X(yr), oy + 8, 1.5, WHITE, 0.4 * a);
      const m = eio(P(t, S.b[3] + 0.3, 3.2));
      const f1 = yr => 0.9 * Math.exp(-((yr - 40) ** 2) / (2 * 12 ** 2));
      const f2 = yr => 1.75 * Math.exp(-((yr - 0.8) ** 2) / (2 * 0.45 ** 2));
      const g = smooth(P(t, 0.6, 1.6));
      ctx.save(); ctx.beginPath(); ctx.moveTo(ox, oy);
      for (let i = 0; i <= 700; i++) { const yr = (i / 700) * 70; ctx.lineTo(X(yr), oy - 300 * lerp(f1(yr), f2(yr), m) * g); }
      ctx.lineTo(X(70), oy); ctx.closePath();
      const gr = ctx.createLinearGradient(0, 160, 0, oy); gr.addColorStop(0, rgba(m > 0.5 ? GOLD : BLUE, 0.35 * a)); gr.addColorStop(1, rgba(m > 0.5 ? GOLD : BLUE, 0.03 * a));
      ctx.fillStyle = gr; ctx.fill();
      ctx.strokeStyle = rgba(mixc(BLUE, GOLD, m), 0.9 * a); ctx.lineWidth = 2.5; ctx.stroke(); ctx.restore();
      // fog over the unknown far future
      const fog = env(t, S.b[2], S.b[3] + 1.2, 1, 1.4);
      if (fog > 0.003) { ctx.save(); ctx.translate(X(58), 470); ctx.scale(1.35, 0.8); const fg = ctx.createRadialGradient(0, 0, 0, 0, 0, 480); fg.addColorStop(0, `rgba(150,165,190,${0.15 * fog})`); fg.addColorStop(1, 'rgba(150,165,190,0)'); ctx.fillStyle = fg; ctx.fillRect(-600, -600, 1200, 1200); ctx.restore(); }
    },
    over(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1.2)), ox = 260, oy = 760, X = yr => ox + yr * 20.5;
      for (let yr = 0; yr <= 70; yr += 10) text(ctx, `${yr}`, X(yr), oy + 34, { size: 24, f: SANS, w: 300, c: COL.dim, a });
      text(ctx, '剩余时间 T（年）', 1720, oy + 76, { size: 26, f: SANS, w: 300, c: COL.dim, a, al: 'right' });
      text(ctx, '可能性', ox - 20, 150, { size: 26, f: SANS, w: 300, c: COL.dim, a, al: 'left' });
      const m = eio(P(t, S.b[3] + 0.3, 3.2));
      tex(ctx, 'pT', X(40), 440, { a: a * (1 - m) });
      tex(ctx, 'pTd', X(6.5), 230, { a: a * m });
      text(ctx, 'D：诊断', X(6.5), 290, { size: 24, f: SANS, w: 300, c: COL.dim, a: a * m });
    },
  },
  {
    id: 's6b',
    beats: [
      { text: '这不是说，两者没有区别。', extra: 0.4 },
      { text: '半年，就是比四十年少。这种失去，是真的。', extra: 1.6 },
      { text: '只是，确诊的人看清的，', extra: 0.4 },
      { text: '是所有人迟早都要面对、\n只是暂时还看不见的东西。', extra: 2.4 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1.4)), x0 = 300;
      const fog = 1 - smooth(P(t, S.b[3] + 0.2, 3.5));
      LIFELINES.forEach((l, i) => {
        const gold = i === 6, c = gold ? GOLD : WHITE;
        const hl = gold ? 1 + 0.5 * env(t, S.b[1], S.e[1], 0.5, 0.8) : 1;
        const al = (gold ? 0.9 : 0.4) * a * Math.min(1, hl);
        if (gold) seg(ctx, x0, l.y, x0 + l.len, l.y, 2.4, c, al);
        else { const lg = ctx.createLinearGradient(640, 0, 980, 0); lg.addColorStop(0, rgba(c, al)); lg.addColorStop(1, rgba(c, al * (1 - fog))); ctx.save(); ctx.strokeStyle = lg; ctx.lineWidth = 1.4; ctx.beginPath(); ctx.moveTo(x0, l.y); ctx.lineTo(x0 + l.len, l.y); ctx.stroke(); ctx.restore(); }
        const endVis = gold ? 1 : 1 - fog;
        light(ctx, x0 + l.len, l.y, gold ? 4 : 3, c, a * endVis * (gold ? hl : 0.7));
        dot(ctx, x0, l.y, 3, gold ? COL.gold : '#cfd5df', a * 0.9);
      });
      seg(ctx, x0, 120, x0, 830, 1.2, WHITE, 0.35 * a, [5, 7]);
    },
    over(t, S, ctx) {
      text(ctx, '现在', 300, 100, { size: 26, f: SANS, w: 300, c: COL.dim, a: smooth(P(t, 0.2, 1.4)) });
    },
  },

  // ---------------- chapter 7 ----------------
  { id: 'c7', card: { num: '柒', title: '不会坍缩的维度' }, dur: 4.2, mood: 'warm', accents: [{ at: 0.2, kind: 'bell' }] },
  {
    id: 's7a',
    beats: [
      { text: '外部的选项变少了。\n但有一个维度，几乎不受影响。', extra: 0.6 },
      { text: '心理学家弗兰克尔，在纳粹集中营里得出过一个结论：' },
      { text: '人的一切都可以被夺走，唯独一样不能——\n在任何处境下，选择自己态度的自由。', style: 'quote', pos: 'center', sub: '—— 维克多·弗兰克尔《活出生命的意义》（大意）', extra: 2.6 },
      { text: '「能去做什么」，变少了。', extra: 1.2 },
      { text: '「怎样对待正在发生的事」「对谁说什么话」，\n几乎没有变。', extra: 1.4 },
      { text: '而很多人回头看时，最在意的，恰恰是后者。', extra: 1.4 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1.4)) * (1 - 0.85 * env(t, S.b[2] - 0.3, S.e[2] + 0.3, 0.8, 0.8));
      const O = [820, 560];
      const sx = lerp(1, 0.22, eio(P(t, S.b[3] + 0.3, 2.2)));
      const yz = 1 + 0.6 * smooth(P(t, S.b[4] + 0.3, 1.5));
      const vx = [640 * sx, 0], vy = [0, -370], vz = [-300, 170];
      const pt = (u, v, w) => [O[0] + vx[0] * u + vy[0] * v + vz[0] * w, O[1] + vx[1] * u + vy[1] * v + vz[1] * w];
      const E = (p, q, c, al, wd = 1.2) => { const A = pt(...p), B = pt(...q); seg(ctx, A[0], A[1], B[0], B[1], wd, c, al); };
      ctx.save(); ctx.globalAlpha = 0.07 * a; ctx.fillStyle = rgba(GOLD, 1); ctx.beginPath();
      [[0, 0, 0], [1, 0, 0], [1, 1, 0], [0, 1, 0]].forEach((p, i) => { const q = pt(...p); i ? ctx.lineTo(q[0], q[1]) : ctx.moveTo(q[0], q[1]); }); ctx.fill();
      ctx.beginPath(); [[0, 0, 0], [0, 1, 0], [0, 1, 1], [0, 0, 1]].forEach((p, i) => { const q = pt(...p); i ? ctx.lineTo(q[0], q[1]) : ctx.moveTo(q[0], q[1]); }); ctx.fill(); ctx.restore();
      const box = [[[1, 0, 0], [1, 1, 0]], [[0, 1, 0], [1, 1, 0]], [[1, 0, 0], [1, 0, 1]], [[0, 0, 1], [1, 0, 1]], [[0, 1, 0], [0, 1, 1]], [[0, 0, 1], [0, 1, 1]], [[1, 1, 0], [1, 1, 1]], [[0, 1, 1], [1, 1, 1]], [[1, 0, 1], [1, 1, 1]]];
      box.forEach(([p, q]) => E(p, q, GOLD, 0.25 * a));
      const xc = mixc(GOLD, RED, smooth(P(t, S.b[3] + 0.3, 2)));
      E([0, 0, 0], [1, 0, 0], xc, 0.9 * a, 3);
      E([0, 0, 0], [0, 1, 0], GOLD, Math.min(1, 0.6 * yz) * a, 2 + yz);
      E([0, 0, 0], [0, 0, 1], GOLD, Math.min(1, 0.6 * yz) * a, 2 + yz);
      [[1, 0, 0], [0, 1, 0], [0, 0, 1]].forEach((p, i) => { const q = pt(...p); light(ctx, q[0], q[1], 4, i ? GOLD : xc, a * (i ? Math.min(1, 0.6 * yz) : 0.8)); });
      light(ctx, O[0], O[1], 5, GOLD, a);
      S._lab = [pt(1, 0, 0), pt(0, 1, 0), pt(0, 0, 1)]; S._a = a;
    },
    over(t, S, ctx) {
      if (!S._lab) return; const a = S._a;
      const [X, Y, Z] = S._lab;
      text(ctx, '能去做什么', X[0] + 40, X[1] + 4, { size: 34, al: 'left', a });
      text(ctx, '怎样对待正在发生的事', Y[0], Y[1] - 46, { size: 34, a, c: COL.gold });
      text(ctx, '对谁说什么话', Z[0] - 20, Z[1] + 46, { size: 34, a, c: COL.gold });
    },
  },

  // ---------------- chapter 8 ----------------
  { id: 'c8', card: { num: '捌', title: '长度与深度' }, dur: 4.2, accents: [{ at: 0.2, kind: 'bell' }] },
  {
    id: 's8a',
    beats: [
      { text: '再回到数学。' },
      { text: '一条长度为 1 的线段，和一条无限长的射线——', extra: 0.6 },
      { text: '它们上面的点，一样多。', extra: 1.8 },
      { text: '无穷有两个方向：\n一个向远处延伸，一个向内部细分。', extra: 1.8 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, S.b[1] - 0.2, 1.2)), y1 = 260, y2 = 620, s0 = 560, s1 = 1360, r0 = 160, unit = 160;
      seg(ctx, s0, y1, s1, y1, 2.5, GOLD, 0.8 * a);
      seg(ctx, r0, y2, 1800, y2, 2.5, BLUE, 0.8 * a);
      ctx.save(); ctx.globalAlpha = 0.8 * a; ctx.fillStyle = rgba(BLUE, 1); ctx.beginPath(); ctx.moveTo(1815, y2); ctx.lineTo(1795, y2 - 9); ctx.lineTo(1795, y2 + 9); ctx.fill(); ctx.restore();
      for (let k = 0; k <= 10; k++) seg(ctx, r0 + k * unit, y2 - 8, r0 + k * unit, y2 + 8, 1.5, BLUE, 0.5 * a);
      const b = smooth(P(t, S.b[2] - 0.6, 1.4));
      for (let k = 0; k < 28; k++) {
        const s = k / 28, xs = lerp(s0, s1, s), r = s / (1 - s), xr = r0 + r * unit;
        const fade = xr > 1800 ? 0 : 1;
        seg(ctx, xs, y1, Math.min(xr, 1800), y2, 1, WHITE, 0.18 * b * a * (fade || 0.5));
      }
      const m = P(t, S.b[2] - 0.4, 5.5), s = 0.97 * eio(m), xs = lerp(s0, s1, s), xr = r0 + (s / (1 - s)) * unit;
      if (m > 0) {
        seg(ctx, xs, y1, Math.min(xr, 1830), y2, 1.8, GOLD, 0.7 * a);
        light(ctx, xs, y1, 5, GOLD, a); if (xr < 1830) light(ctx, xr, y2, 5, BLUE, a);
      }
      const d = smooth(P(t, S.b[3] + 0.6, 2));
      for (let lv = 1; lv <= 6; lv++) {
        const n = 2 ** lv, al = d * smooth(P(t, S.b[3] + 0.6 + lv * 0.35, 0.6));
        for (let k = 1; k < n; k++) if (k % 2) seg(ctx, lerp(s0, s1, k / n), y1 - 22 / lv, lerp(s0, s1, k / n), y1 + 22 / lv, 1.2, GOLD, 0.6 * al);
      }
    },
    over(t, S, ctx) {
      const a = smooth(P(t, S.b[1] - 0.2, 1.2));
      text(ctx, '0', 560, 310, { size: 26, f: SANS, w: 300, c: COL.dim, a }); text(ctx, '1', 1360, 310, { size: 26, f: SANS, w: 300, c: COL.dim, a });
      for (let k = 0; k <= 10; k++) text(ctx, String(k), 160 + k * 160, 664, { size: 24, f: SANS, w: 300, c: COL.dim, a });
      text(ctx, '长度为 1', 960, 200, { size: 30, c: COL.gold, a });
      text(ctx, '无限长', 960, 730, { size: 30, c: '#8fb4dc', a });
      tex(ctx, 'bij', 1620, 380, { a: smooth(P(t, S.b[2], 1)) });
      const d = smooth(P(t, S.b[3] + 0.2, 1));
      text(ctx, '向远处延伸 →', 1620, 560, { size: 32, c: '#8fb4dc', a: d });
      text(ctx, '向内部细分', 960, 140, { size: 32, c: COL.gold, a: d });
    },
  },
  {
    id: 's8b',
    beats: [
      { text: '还记得 ε 吗？\n一段时间里可区分的瞬间，大约有 T / ε 个。', extra: 0.4 },
      { text: 'T，是时间的长度。ε，是你对时间的分辨率。', extra: 0.8 },
      { text: '注意得越细，同样长的时间里，\n被真正经历的瞬间就越多。', extra: 2.2 },
      { text: '这不是安慰。时间少了，就是少了。', extra: 1.4 },
      { text: '只是提醒：长度，不是生命唯一的维度。', extra: 1.6 },
    ],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1.2)), y = 470;
      const len = lerp(1000, 520, eio(P(t, S.b[3] + 0.3, 2.2)));
      const x0 = 960 - len / 2, x1 = 960 + len / 2;
      const per = Math.pow(2, lerp(3, 9, eio(P(t, S.b[2] - 0.2, 4.5))));
      const n = Math.round(per * len / 1000);
      S._n = n;
      seg(ctx, x0, y, x1, y, 3, GOLD, 0.85 * a);
      const dens = clamp(n / 200);
      ctx.save(); ctx.strokeStyle = rgba(GOLD, (0.75 - 0.4 * dens) * a); ctx.lineWidth = 1.2; ctx.beginPath();
      for (let k = 0; k <= n; k++) { const x = lerp(x0, x1, k / n); ctx.moveTo(x, y - 14); ctx.lineTo(x, y + 14); }
      ctx.stroke(); ctx.restore();
      if (dens > 0.2) glow(ctx, 960, y, len * 0.55, GOLD, 0.12 * dens * a);
      light(ctx, x0, y, 4, GOLD, a); light(ctx, x1, y, 4, GOLD, a);
      const depth = smooth(P(t, S.b[4] + 0.3, 2));
      if (depth > 0) for (let k = 0; k <= 40; k++) { const x = lerp(x0, x1, k / 40), h = 160 * depth * (0.6 + 0.4 * Math.sin(k * 1.7 + t)); seg(ctx, x, y, x, y - h, 1, GOLD, 0.18 * depth * a); }
    },
    over(t, S, ctx) {
      const a = smooth(P(t, 0.2, 1.2));
      tex(ctx, 'teps', 960, 220, { a });
      text(ctx, `T / ε ≈ ${S._n || 8}`, 960, 560, { size: 40, c: COL.gold, a: a * smooth(P(t, S.b[1], 1)) });
    },
  },
  {
    id: 's8c', pre: 0.9,
    beats: [
      { text: '如果把永恒理解为无时间性，\n而不是时间的无限延续，\n那么，活在当下的人就永恒地活着。', style: 'quote', pos: 'center', sub: '—— 维特根斯坦《逻辑哲学论》6.4311', extra: 3.8 },
    ],
    post: 1.2,
    draw(t, S, ctx) {
      for (let k = 0; k < 5; k++) {
        const u = ((t * 0.12 + k / 5) % 1);
        circle(ctx, 960, 520, 40 + u * 700, 1.2, GOLD, 0.18 * Math.sin(Math.PI * u) * smooth(P(t, 0.3, 1.5)));
      }
    },
  },

  // ---------------- chapter 9 ----------------
  { id: 'c9', card: { num: '玖', title: '如何面对' }, dur: 4.2, accents: [{ at: 0.2, kind: 'bell' }] },
  {
    id: 's9a',
    beats: [
      { text: '那么，该怎样面对？', pos: 'center', extra: 0.6 },
      { text: '首先：没有正确的方式。', pos: 'center', extra: 1.0 },
      { text: '恐惧、愤怒、否认、讨价还价，都很正常。\n它们不按顺序来，也不一定会走到「接受」。', extra: 1.6 },
      { text: '「你要积极乐观」——\n这种要求本身，有时就是额外的负担。', extra: 1.6 },
    ],
    draw(t, S, ctx) {},
    over(t, S, ctx) {
      const words = ['恐惧', '愤怒', '否认', '讨价还价', '悲伤', '不甘', '平静', '释然', '想念'];
      const a = smooth(P(t, S.b[2], 1.5)) * (1 - 0.55 * smooth(P(t, S.b[3], 1.2)));
      words.forEach((w, i) => {
        const bx = 260 + (i % 5) * 350 + (Math.floor(i / 5) % 2) * 170, by = 230 + Math.floor(i / 5) * 280;
        const x = bx + 60 * Math.sin(t * 0.21 + i * 1.7), y = by + 40 * Math.sin(t * 0.17 + i * 2.3);
        text(ctx, w, x, y, { size: 40 + (i % 3) * 8, c: i % 4 === 0 ? COL.gold : COL.dim, a: a * smooth(P(t, S.b[2] + i * 0.25, 1)) * (0.5 + 0.2 * Math.sin(t * 0.5 + i)) });
      });
    },
  },
  {
    id: 's9b',
    beats: [
      { text: '在这个前提下，\n陪伴过许多临终者的人，留下了一些相似的经验。', extra: 0.6 },
      { text: '弄清自己最在乎什么，再让治疗为它服务。', sub: '—— 阿图·葛文德《最好的告别》', extra: 1.6 },
      { text: '了解安宁疗护：它不是放弃治疗，\n而是让剩下的日子，少一些痛苦。', sub: '研究：晚期肺癌患者尽早接受姑息治疗，生活质量更好，中位生存期还多了约 2.7 个月（Temel 等，NEJM，2010）', extra: 2.4 },
      { text: '为家人写下你的意愿，\n让他们在最难的时候，不必替你猜。', extra: 1.4 },
    ],
    draw(t, S, ctx) {
      const xs = [560, 960, 1360], y = 400;
      const a = smooth(P(t, 0.2, 1.2));
      seg(ctx, 360, y, 1560, y, 1.2, WHITE, 0.2 * a, [3, 9]);
      xs.forEach((x, i) => {
        const u = smooth(P(t, S.b[i + 1], 1.2));
        circle(ctx, x, y, 22, 1.5, WHITE, 0.35 * a);
        light(ctx, x, y, 7, GOLD, u * (0.9 + 0.1 * Math.sin(t * 2 + i)));
        glow(ctx, x, y, 200, GOLD, 0.12 * u);
      });
    },
    over(t, S, ctx) {
      const xs = [560, 960, 1360], labels = ['最在乎什么', '安宁疗护', '写下意愿'];
      xs.forEach((x, i) => text(ctx, labels[i], x, 500, { size: 38, c: COL.text, a: smooth(P(t, S.b[i + 1] + 0.3, 1)) }));
    },
  },
  {
    id: 's9c',
    beats: [
      { text: '把该说的话，说出来。', extra: 0.4 },
      { d: 8.2 },
      { text: '姑息医学医生 Ira Byock 说，\n这是告别之前，最重要的四句话。', sub: 'Ira Byock, The Four Things That Matter Most (2004)', extra: 1.0 },
      { text: '允许自己，同时悲伤，和快乐。', extra: 1.6 },
      { text: '它们并不矛盾——生活，还在继续。', extra: 2.0 },
    ],
    draw(t, S, ctx) {
      const C = [960, 450], pos = [[560, 250], [1360, 250], [560, 650], [1360, 650]];
      const fade = 1 - 0.7 * smooth(P(t, S.b[3], 1.5));
      const split = smooth(P(t, S.b[3] + 0.4, 2));
      pos.forEach(([x, y], i) => {
        const u = smooth(P(t, S.b[1] + i * 1.9, 1.2)) * fade;
        seg(ctx, C[0], C[1], lerp(C[0], x, u), lerp(C[1], y, u), 1.5, GOLD, 0.5 * u);
        light(ctx, x, y, 5, GOLD, u);
      });
      const th = t * 0.8, r = 46 * split;
      light(ctx, C[0] + Math.cos(th) * r, C[1] + Math.sin(th) * r * 0.6, 7, GOLD, 1);
      if (split > 0) light(ctx, C[0] - Math.cos(th) * r, C[1] - Math.sin(th) * r * 0.6, 7, BLUE, split);
    },
    over(t, S, ctx) {
      const pos = [[560, 250], [1360, 250], [560, 650], [1360, 650]], words = ['请原谅我', '我原谅你', '谢谢你', '我爱你'];
      const fade = 1 - 0.7 * smooth(P(t, S.b[3], 1.5));
      pos.forEach(([x, y], i) => {
        const u = smooth(P(t, S.b[1] + i * 1.9 + 0.3, 1.2)) * fade;
        text(ctx, words[i], x + (x < 960 ? -40 : 40), y, { size: 54, c: COL.quote, a: u, al: x < 960 ? 'right' : 'left', ls: 6 });
      });
    },
  },

  // ---------------- chapter 10 ----------------
  { id: 'c10', card: { num: '拾', title: '正在走的这一条' }, dur: 4.4, mood: 'end', accents: [{ at: 0.2, kind: 'bell' }] },
  {
    id: 's10a',
    beats: [
      { text: '神经外科医生保罗·卡拉尼什，三十六岁时确诊肺癌晚期。', extra: 0.4 },
      { text: '他在《当呼吸化为空气》里，引用了贝克特的一句话：', extra: 0.2 },
      { text: '我无法继续。我会继续。', style: 'quote', pos: 'center', sub: '—— 塞缪尔·贝克特，引自《当呼吸化为空气》', extra: 2.6 },
      { text: '书的最后，他写给只有几个月大的女儿——', extra: 0.4 },
      { text: '希望她记得：她曾让一个将死之人最后的日子，\n充满了喜悦。', extra: 2.6 },
    ],
    draw(t, S, ctx) {
      const q = 1 - 0.8 * env(t, S.b[2] - 0.3, S.e[2] + 0.3, 0.8, 0.8);
      const a = smooth(P(t, 0.3, 1.5)) * q;
      const fl = 0.9 + 0.1 * Math.sin(t * 3.1) * Math.sin(t * 1.7);
      light(ctx, 900, 470, 7 * fl, GOLD, a);
      glow(ctx, 900, 470, 360, GOLD, 0.12 * a);
      const d = smooth(P(t, S.b[3] + 0.5, 2)) * q;
      light(ctx, 1040, 500, 4, [255, 220, 170], d);
      glow(ctx, 970, 485, 260, GOLD, 0.1 * d);
    },
  },
  {
    id: 's10b',
    beats: [
      { text: '回到开头的问题。', extra: 0.4 },
      { text: '人生的分岔，可以数得清吗？', extra: 1.2 },
      { text: '数学能算出，还剩多少条路；', extra: 0.6 },
      { text: '却量不出，一条路的价值。', extra: 1.4 },
      { text: '一个人的一生是否完整，', extra: 0.4 },
      { text: '从来不取决于，有多少条路没有走，', extra: 0.6 },
      { text: '而在于，正在走的这一条。', extra: 3.4 },
    ],
    post: 1.0,
    accents: [{ at: S => S.b[6], kind: 'bell' }],
    draw(t, S, ctx) {
      const a = smooth(P(t, 0, 1.6));
      const p = P(t, S.b[2], S.e[6] - S.b[2] - 2.0) * PATH_A.length;
      const rest = 1 - 0.75 * smooth(P(t, S.b[6], 2.5));
      const onPath = new Set(PATH_A);
      drawTree(ctx, TREE_A, { a: a * 0.45 * rest, col: GREY, per: e => (onPath.has(e) ? null : { a: 1 }) });
      ctx.save(); ctx.lineCap = 'round';
      PATH_A.forEach((e, i) => {
        const f = clamp(p - i); if (f <= 0) { ctx.beginPath(); edgePath(ctx, e, 1); ctx.strokeStyle = rgba(GREY, 0.45 * a * rest); ctx.lineWidth = 1.5; ctx.stroke(); return; }
        ctx.beginPath(); edgePath(ctx, e, f); ctx.strokeStyle = rgba(GOLD, 0.95 * a); ctx.lineWidth = 3.4; ctx.stroke();
      });
      ctx.restore();
      const i = Math.min(PATH_A.length - 1, Math.floor(p)), f = p >= PATH_A.length ? 1 : p - i;
      const [hx, hy] = p > 0 ? edgePoint(PATH_A[i], f) : [230, 470];
      light(ctx, hx, hy, 6 + 2 * smooth(P(t, S.b[6], 2)), GOLD, a);
      light(ctx, 230, 470, 4, GOLD, a * 0.8);
    },
  },
  {
    id: 'end', mood: 'end', fout: 2.0,
    dur: 11,
    draw(t, S, ctx) {
      const a = smooth(P(t, 0, 1.5));
      ctx.save(); ctx.lineCap = 'round'; ctx.strokeStyle = rgba(GOLD, 0.12 * a); ctx.lineWidth = 2;
      ctx.beginPath(); PATH_A.forEach(e => edgePath(ctx, e, 1)); ctx.stroke(); ctx.restore();
    },
    over(t, S, ctx) {
      const a = smooth(P(t, 0.3, 1.5));
      text(ctx, '人生的分岔，可以数得清吗？', 960, 250, { size: 60, w: 700, a, ls: 3 });
      seg(ctx, 900, 315, 1020, 315, 1.5, GOLD, 0.6 * a);
      const refs = [
        'J. D. Bekenstein, Phys. Rev. D 23, 287 (1981)', '豪尔赫·路易斯·博尔赫斯《小径分岔的花园》',
        '维克多·弗兰克尔《活出生命的意义》', '维特根斯坦《逻辑哲学论》',
        '阿图·葛文德《最好的告别》', 'J. S. Temel et al., NEJM 363:733–742 (2010)',
        'Ira Byock, The Four Things That Matter Most (2004)', '保罗·卡拉尼什《当呼吸化为空气》',
      ];
      const b = smooth(P(t, 1.2, 1.5));
      text(ctx, '参考', 960, 385, { size: 28, f: SANS, w: 300, c: COL.gold, a: b, ls: 8 });
      refs.forEach((r, i) => text(ctx, r, i % 2 ? 1010 : 910, 450 + Math.floor(i / 2) * 50, { size: 25, f: SANS, w: 300, c: COL.dim, a: b, al: i % 2 ? 'left' : 'right' }));
      text(ctx, '谢谢观看', 960, 820, { size: 40, c: COL.gold, a: smooth(P(t, 2.5, 1.5)), ls: 12 });
    },
  },
];

// ---------- cover (B站封面) ----------
function renderCover() {
  for (const c of [M, L]) { c.setTransform(1, 0, 0, 1, 0, 0); c.globalAlpha = 1; c.globalCompositeOperation = 'source-over'; c.filter = 'none'; }
  M.drawImage(BG, 0, 0); L.clearRect(0, 0, W, H);
  drawDust(L, 3, 1);
  drawTree(L, TREE_COVER, { a: 0.75, w: 3.4 });
  L.save(); L.lineCap = 'round'; L.strokeStyle = rgba([255, 226, 170], 1); L.lineWidth = 5; L.beginPath(); PATH_COVER.forEach(e => edgePath(L, e, 1)); L.stroke(); L.restore();
  const last = PATH_COVER[PATH_COVER.length - 1].c;
  light(L, last.x, last.y, 10, GOLD, 1); light(L, 1130, 540, 9, GOLD, 1);
  M.drawImage(Lc, 0, 0);
  SM.clearRect(0, 0, SW, SH); SM.filter = 'blur(3px)'; SM.drawImage(Lc, 0, 0, SW, SH); SM.filter = 'none';
  M.globalCompositeOperation = 'lighter'; M.drawImage(SMc, 0, 0, W, H); M.globalCompositeOperation = 'source-over';
  const gr = M.createLinearGradient(0, 0, 1150, 0); gr.addColorStop(0, 'rgba(3,4,9,0.85)'); gr.addColorStop(1, 'rgba(3,4,9,0)');
  M.fillStyle = gr; M.fillRect(0, 0, 1150, H);
  text(M, '人生的分岔', 110, 400, { size: 136, w: 900, al: 'left', sh: 30 });
  text(M, '可以数得清吗？', 110, 580, { size: 136, w: 900, al: 'left', c: COL.gold, sh: 30 });
  seg(M, 116, 715, 300, 715, 3, GOLD, 0.8);
  text(M, '从无穷，到有限，再到最后一程', 112, 790, { size: 46, f: SANS, w: 300, al: 'left', c: '#c9cfda', ls: 4 });
}
window.renderCover = renderCover;
