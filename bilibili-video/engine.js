'use strict';
// ------------------------------------------------------------
//  Engine: timeline, captions, math, trees, glow, compositing.
//  Everything is a pure function of time t, so any frame can be
//  rendered independently (parallel rendering, exact replays).
// ------------------------------------------------------------
const W = 1920, H = 1080, FPS = 30;
const SERIF = '"Noto Serif SC"', SANS = '"Noto Sans SC"';
const COL = {
  text: '#ece6dc', dim: '#8e96a6', gold: '#e9c27b', quote: '#f3e6c8',
  goldRGB: [233, 194, 123], blueRGB: [143, 180, 220], redRGB: [214, 122, 104],
  whiteRGB: [236, 230, 220], greyRGB: [120, 132, 150],
};
const GOLD = COL.goldRGB, BLUE = COL.blueRGB, RED = COL.redRGB, WHITE = COL.whiteRGB, GREY = COL.greyRGB;

// ---------- math utils ----------
const clamp = (x, a = 0, b = 1) => (x < a ? a : x > b ? b : x);
const lerp = (a, b, u) => a + (b - a) * u;
const smooth = x => { x = clamp(x); return x * x * (3 - 2 * x); };
const eio = x => { x = clamp(x); return x < 0.5 ? 4 * x * x * x : 1 - Math.pow(-2 * x + 2, 3) / 2; };
const eout = x => 1 - Math.pow(1 - clamp(x), 3);
const ein = x => Math.pow(clamp(x), 3);
const P = (t, a, d) => clamp((t - a) / d);
const env = (t, a, b, fi = 0.6, fo = 0.6) => smooth((t - a) / fi) * (1 - smooth((t - (b - fo)) / fo));
const rgba = (c, a) => `rgba(${c[0]},${c[1]},${c[2]},${clamp(a)})`;
const mixc = (c1, c2, u) => [0, 1, 2].map(i => Math.round(lerp(c1[i], c2[i], clamp(u))));
function rng(seed) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const hash = n => { const x = Math.sin(n * 127.1 + 311.7) * 43758.5453; return x - Math.floor(x); };
// piecewise keyframes [[t, v], ...] with eased segments
function keys(t, ks, ease = eio) {
  if (t <= ks[0][0]) return ks[0][1];
  for (let i = 1; i < ks.length; i++) {
    if (t <= ks[i][0]) {
      const [t0, v0] = ks[i - 1], [t1, v1] = ks[i];
      return lerp(v0, v1, ease((t - t0) / Math.max(1e-6, t1 - t0)));
    }
  }
  return ks[ks.length - 1][1];
}
function mkCanvas(w, h) { const c = document.createElement('canvas'); c.width = w; c.height = h; return c; }

// ---------- drawing primitives ----------
const SPR = {};
function glowSprite(rgb) {
  const k = rgb.join(',');
  if (SPR[k]) return SPR[k];
  const c = mkCanvas(256, 256), g = c.getContext('2d');
  const gr = g.createRadialGradient(128, 128, 0, 128, 128, 128);
  gr.addColorStop(0, `rgba(${k},1)`);
  gr.addColorStop(0.12, `rgba(${k},0.6)`);
  gr.addColorStop(0.35, `rgba(${k},0.18)`);
  gr.addColorStop(0.7, `rgba(${k},0.04)`);
  gr.addColorStop(1, `rgba(${k},0)`);
  g.fillStyle = gr; g.fillRect(0, 0, 256, 256);
  return (SPR[k] = c);
}
function glow(ctx, x, y, r, rgb, a = 1) {
  if (a <= 0.003 || r <= 0) return;
  ctx.save(); ctx.globalAlpha *= clamp(a); ctx.globalCompositeOperation = 'lighter';
  ctx.drawImage(glowSprite(rgb), x - r, y - r, 2 * r, 2 * r); ctx.restore();
}
function dot(ctx, x, y, r, col, a = 1) {
  if (a <= 0.003) return;
  ctx.save(); ctx.globalAlpha *= clamp(a); ctx.fillStyle = Array.isArray(col) ? rgba(col, 1) : col;
  ctx.beginPath(); ctx.arc(x, y, r, 0, Math.PI * 2); ctx.fill(); ctx.restore();
}
function light(ctx, x, y, r, rgb, a = 1) { // a bright point with halo
  glow(ctx, x, y, r * 9, rgb, a * 0.55);
  glow(ctx, x, y, r * 3.2, rgb, a * 0.9);
  dot(ctx, x, y, r, mixc(rgb, [255, 255, 255], 0.6), a);
}
function seg(ctx, x1, y1, x2, y2, w, col, a = 1, dash) {
  if (a <= 0.003) return;
  ctx.save(); ctx.globalAlpha *= clamp(a); ctx.strokeStyle = Array.isArray(col) ? rgba(col, 1) : col;
  ctx.lineWidth = w; ctx.lineCap = 'round'; if (dash) ctx.setLineDash(dash);
  ctx.beginPath(); ctx.moveTo(x1, y1); ctx.lineTo(x2, y2); ctx.stroke(); ctx.restore();
}
function circle(ctx, x, y, r, w, col, a = 1, fill) {
  if (a <= 0.003) return;
  ctx.save(); ctx.globalAlpha *= clamp(a); ctx.beginPath(); ctx.arc(x, y, r, 0, Math.PI * 2);
  if (fill) { ctx.fillStyle = fill; ctx.fill(); }
  if (w) { ctx.strokeStyle = Array.isArray(col) ? rgba(col, 1) : col; ctx.lineWidth = w; ctx.stroke(); }
  ctx.restore();
}
function text(ctx, s, x, y, o = {}) {
  const a = o.a ?? 1; if (a <= 0.003) return 0;
  const size = o.size || 40;
  ctx.save(); ctx.globalAlpha *= clamp(a);
  ctx.font = `${o.w || 400} ${size}px ${o.f || SERIF}`;
  ctx.fillStyle = Array.isArray(o.c) ? rgba(o.c, 1) : (o.c || COL.text);
  ctx.textAlign = o.al || 'center'; ctx.textBaseline = 'middle';
  if (o.ls) ctx.letterSpacing = o.ls + 'px';
  if (o.sh) { ctx.shadowColor = 'rgba(0,0,0,0.9)'; ctx.shadowBlur = o.sh; }
  ctx.fillText(s, x, y);
  const w = ctx.measureText(s).width;
  ctx.restore();
  return w;
}
function measure(s, size, w = 400, f = SERIF) {
  M.save(); M.font = `${w} ${size}px ${f}`; const r = M.measureText(s).width; M.restore(); return r;
}
// rich number: parts = ['text', ['base','exp'], ...]; centered at x unless o.al
function rich(ctx, parts, x, y, o = {}) {
  const size = o.size || 40, es = size * 0.6, gap = size * 0.05;
  const widths = parts.map(p => (typeof p === 'string' ? measure(p, size, o.w, o.f)
    : measure(p[0], size, o.w, o.f) + gap + measure(p[1], es, o.w, o.f)));
  const tw = widths.reduce((s, v) => s + v, 0);
  let cx = o.al === 'left' ? x : o.al === 'right' ? x - tw : x - tw / 2;
  parts.forEach((p, i) => {
    if (typeof p === 'string') text(ctx, p, cx, y, { ...o, al: 'left' });
    else {
      const wb = measure(p[0], size, o.w, o.f);
      text(ctx, p[0], cx, y, { ...o, al: 'left' });
      text(ctx, p[1], cx + wb + gap, y - size * 0.42, { ...o, size: es, al: 'left' });
    }
    cx += widths[i];
  });
  return tw;
}

// ---------- TeX (MathJax -> SVG -> Image) ----------
const TEX = {};
function defTex(key, tex, color = COL.text, ex = 30) { TEX[key] = { tex, color, ex }; }
async function prepTex() {
  await MathJax.startup.promise;
  for (const k of Object.keys(TEX)) {
    const o = TEX[k];
    const node = MathJax.tex2svg(o.tex, { display: true });
    const svg = node.querySelector('svg');
    const w = parseFloat(svg.getAttribute('width')) * o.ex;
    const h = parseFloat(svg.getAttribute('height')) * o.ex;
    svg.setAttribute('width', w); svg.setAttribute('height', h);
    svg.setAttribute('color', o.color); svg.setAttribute('style', `color:${o.color}`);
    const s = new XMLSerializer().serializeToString(svg);
    const img = new Image();
    img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(s);
    await img.decode();
    o.img = img; o.w = w; o.h = h;
  }
}
function tex(ctx, key, x, y, o = {}) {
  const T = TEX[key]; const a = o.a ?? 1; const s = o.s || 1;
  if (!T) throw new Error('tex ' + key);
  if (a <= 0.003) return T.w * s;
  const x0 = o.al === 'left' ? x : o.al === 'right' ? x - T.w * s : x - (T.w * s) / 2;
  ctx.save(); ctx.globalAlpha *= clamp(a);
  if (o.sh) { ctx.shadowColor = 'rgba(0,0,0,0.85)'; ctx.shadowBlur = o.sh; }
  ctx.drawImage(T.img, x0, y - (T.h * s) / 2, T.w * s, T.h * s);
  ctx.restore();
  return T.w * s;
}
// tex followed by Chinese text, centered as a unit
function texLabel(ctx, key, label, x, y, o = {}) {
  const s = o.s || 1, size = o.size || 32;
  const tw = TEX[key].w * s, lw = measure(label, size, 400, o.f || SANS);
  const x0 = x - (tw + lw + 6) / 2;
  tex(ctx, key, x0, y, { s, a: o.a, al: 'left' });
  text(ctx, label, x0 + tw + 6, y + 2, { size, f: o.f || SANS, w: 300, c: o.c || COL.dim, al: 'left', a: o.a });
}

// ---------- branching trees ----------
function bez(x0, y0, x1, y1, n) {
  const pts = new Float32Array((n + 1) * 2), dx = x1 - x0;
  const c1x = x0 + dx * 0.5, c1y = y0, c2x = x0 + dx * 0.5, c2y = y1;
  for (let i = 0; i <= n; i++) {
    const t = i / n, u = 1 - t;
    pts[2 * i] = u * u * u * x0 + 3 * u * u * t * c1x + 3 * u * t * t * c2x + t * t * t * x1;
    pts[2 * i + 1] = u * u * u * y0 + 3 * u * u * t * c1y + 3 * u * t * t * c2y + t * t * t * y1;
  }
  return pts;
}
function makeTree(o) {
  const R = rng(o.seed), D = o.depth;
  const root = { x: o.x0, y: o.y0, d: 0, kids: [], parent: null, r: R() };
  const nodes = [root], edges = [];
  const xAt = d => o.x0 + (o.x1 - o.x0) * Math.pow(d / D, o.xp ?? 1);
  (function grow(n, lo, hi) {
    if (n.d >= D) return;
    const k = o.k || (R() < (o.p3 ?? 0.12) ? 3 : 2);
    for (let i = 0; i < k; i++) {
      const a = lo + ((hi - lo) * i) / k, b = lo + ((hi - lo) * (i + 1)) / k;
      const c = {
        x: xAt(n.d + 1) + (R() - 0.5) * (o.xj ?? 0.3) * (o.x1 - o.x0) / D,
        y: (a + b) / 2 + (R() - 0.5) * (b - a) * (o.yj ?? 0.6),
        d: n.d + 1, kids: [], parent: n, r: R(),
      };
      n.kids.push(c); nodes.push(c);
      const e = { p: n, c, d: c.d, r: c.r, pts: bez(n.x, n.y, c.x, c.y, o.seg || 12) };
      e.mx = (n.x + c.x) / 2; e.my = (n.y + c.y) / 2; c.edge = e; edges.push(e);
      grow(c, a, b);
    }
  })(root, o.yTop, o.yBot);
  const byD = [];
  for (const e of edges) (byD[e.d] = byD[e.d] || []).push(e);
  return { root, nodes, edges, byD, D };
}
function edgePath(ctx, e, f, dy = 0) {
  const p = e.pts, n = p.length / 2 - 1, m = f * n, k = Math.floor(m);
  ctx.moveTo(p[0], p[1] + dy);
  for (let i = 1; i <= k; i++) ctx.lineTo(p[2 * i], p[2 * i + 1] + dy);
  if (k < n) { const fr = m - k; ctx.lineTo(lerp(p[2 * k], p[2 * k + 2], fr), lerp(p[2 * k + 1], p[2 * k + 3], fr) + dy); }
}
function edgePoint(e, f) {
  const p = e.pts, n = p.length / 2 - 1, m = clamp(f) * n, k = Math.min(n - 1, Math.floor(m)), fr = m - k;
  return [lerp(p[2 * k], p[2 * k + 2], fr), lerp(p[2 * k + 1], p[2 * k + 3], fr)];
}
const treeW = (o, d) => Math.max(o.wm ?? 0.6, (o.w ?? 2.6) * Math.pow(0.8, d - 1));
const treeA = d => 0.92 - 0.05 * d;
function drawTree(ctx, tr, o = {}) {
  const g = o.g ?? 99, A = o.a ?? 1, col = o.col || GOLD, sc = o.scale || 1;
  if (A <= 0.003) return;
  ctx.save(); ctx.lineCap = 'round'; ctx.lineJoin = 'round';
  if (!o.per) {
    for (let d = 1; d <= tr.D; d++) {
      const f = clamp(g - (d - 1)); if (f <= 0) break;
      ctx.beginPath();
      for (const e of tr.byD[d]) edgePath(ctx, e, f);
      ctx.strokeStyle = rgba(col, A * (o.da ? o.da(d) : treeA(d)));
      ctx.lineWidth = treeW(o, d) / sc; ctx.stroke();
    }
  } else {
    for (const e of tr.edges) {
      const f = clamp(g - (e.d - 1)); if (f <= 0) continue;
      const s = o.per(e); if (!s || s.a <= 0.003) continue;
      ctx.beginPath(); edgePath(ctx, e, f, s.dy || 0);
      ctx.strokeStyle = rgba(s.c || col, A * s.a * treeA(e.d));
      ctx.lineWidth = (s.w ?? treeW(o, e.d)) / sc; ctx.stroke();
    }
  }
  if (o.tips && g < tr.D) {
    const d = Math.floor(g) + 1, f = g - Math.floor(g);
    if (d <= tr.D && f > 0) for (const e of tr.byD[d]) {
      const [x, y] = edgePoint(e, f);
      glow(ctx, x, y, 22 * Math.pow(0.82, d), col, o.tips * A * (d > 7 ? 0.5 : 1));
    }
  }
  ctx.restore();
}
function pathOf(tr, seed, bias = 0.5) { // root -> leaf, staying near the vertical middle
  const R = rng(seed); const out = []; let n = tr.root;
  while (n.kids.length) {
    const ks = n.kids.slice().sort((a, b) => Math.abs(a.y - 470) - Math.abs(b.y - 470));
    n = R() < bias ? ks[0] : ks[Math.floor(R() * ks.length)];
    out.push(n.edge);
  }
  return out;
}

// ---------- background ----------
let M, L, Lc, SM, SMc, BG, TL;
const SW = 480, SH = 270;
function makeBG() {
  const c = mkCanvas(W, H), g = c.getContext('2d');
  const gr = g.createRadialGradient(960, 470, 0, 960, 540, 1250);
  gr.addColorStop(0, '#121726'); gr.addColorStop(0.5, '#0a0d17'); gr.addColorStop(1, '#030409');
  g.fillStyle = gr; g.fillRect(0, 0, W, H);
  const id = g.getImageData(0, 0, W, H), d = id.data, R = rng(99);
  for (let i = 0; i < d.length; i += 4) { const n = (R() - 0.5) * 5; d[i] += n; d[i + 1] += n; d[i + 2] += n; }
  g.putImageData(id, 0, 0);
  return c;
}
const DUST = Array.from({ length: 70 }, (_, i) => ({
  x: hash(i) * W, y: hash(i + 100) * H, r: 0.6 + hash(i + 200) * 1.5, s: 0.25 + hash(i + 300), p: hash(i + 400) * 6.28,
}));
function drawDust(ctx, T, a) {
  if (a <= 0) return;
  for (const d of DUST) {
    const x = (((d.x + T * 7 * d.s + 30 * Math.sin(T * 0.1 * d.s + d.p)) % W) + W) % W;
    const y = (((d.y - T * 4 * d.s) % H) + H) % H;
    dot(ctx, x, y, d.r, '#cfd8e6', a * (0.12 + 0.12 * Math.sin(T * 0.6 * d.s + d.p)));
  }
}

// ---------- captions ----------
const CAP = {
  body: { size: 50, w: 400, c: COL.text, lh: 1.62 },
  quote: { size: 54, w: 400, c: COL.quote, lh: 1.8 },
  big: { size: 70, w: 700, c: COL.text, lh: 1.5 },
};
const LS = 3;
function layoutCaption(b) {
  if (b._lay) return b._lay;
  const st = CAP[b.style || 'body']; const size = b.size || st.size;
  M.save(); M.font = `${st.w} ${size}px ${SERIF}`;
  let n = 0;
  const lines = b.text.split('\n').map(line => {
    const gl = []; let gold = false, x = 0;
    for (const ch of line) {
      if (ch === '*') { gold = !gold; continue; }
      const w = M.measureText(ch).width; gl.push({ ch, x, gold }); x += w + LS; n++;
    }
    return { gl, w: x - LS };
  });
  M.restore();
  return (b._lay = { lines, size, st, n });
}
function captionY(b) {
  const lay = layoutCaption(b), lh = lay.size * lay.st.lh, bh = (lay.lines.length - 1) * lh;
  if (b.y != null) return b.y - bh / 2;
  if (b.pos === 'center') return 540 - bh / 2 - (b.sub ? 24 : 0);
  if (b.pos === 'top') return 150;
  return 952 - bh - (b.sub ? 40 : 0);
}
function drawCaption(ctx, b, t, b0, b1) {
  const lay = layoutCaption(b), { size, st } = lay, lh = size * st.lh;
  const y0 = captionY(b);
  const fo = 1 - smooth((t - (b1 - 0.5)) / 0.5);
  const stg = Math.min(0.04, 0.9 / Math.max(1, lay.n));
  let i = 0;
  ctx.save();
  ctx.font = `${st.w} ${size}px ${SERIF}`; ctx.textBaseline = 'middle'; ctx.textAlign = 'left';
  ctx.shadowColor = 'rgba(0,0,0,0.9)'; ctx.shadowBlur = 16;
  lay.lines.forEach((ln, li) => {
    const x0 = 960 - ln.w / 2, y = y0 + li * lh;
    for (const g of ln.gl) {
      const u = smooth((t - b0 - 0.05 - i * stg) / 0.55); i++;
      const a = u * fo; if (a <= 0.003) continue;
      ctx.globalAlpha = a; ctx.fillStyle = g.gold ? COL.gold : st.c;
      ctx.fillText(g.ch, x0 + g.x, y + (1 - u) * 12);
    }
  });
  ctx.restore();
  const lastY = y0 + (lay.lines.length - 1) * lh;
  if (b.style === 'quote') {
    const u = eout((t - b0) / 1.2) * fo;
    seg(ctx, 960 - 40 * u, y0 - lh * 0.75, 960 + 40 * u, y0 - lh * 0.75, 2, GOLD, 0.8 * u);
  }
  if (b.sub) {
    const a = smooth((t - b0 - 0.8) / 0.7) * fo;
    text(ctx, b.sub, 960, lastY + lh * 0.5 + 34, { size: 27, f: SANS, w: 300, c: COL.dim, a, sh: 10, ls: 1 });
  }
}

// ---------- timeline ----------
const READ_BASE = 1.1, READ_CPS = 5.5;
function textLen(s) { return [...s.replace(/[*\n ]/g, '')].length; }
function beatDur(b) { return b.d != null ? b.d : READ_BASE + (b.text ? textLen(b.text) / READ_CPS : 0) + (b.extra || 0); }
function buildTimeline() {
  let T = 0, mood = 'wonder';
  const scenes = [];
  for (const def of SCENES) {
    const S = { id: def.id, def, start: T, b: [], e: [], beats: def.beats || [] };
    let t = def.pre ?? 0.7;
    for (const b of S.beats) { S.b.push(t); t += beatDur(b); S.e.push(t); t += b.gap ?? 0.15; }
    t += def.post ?? 0.4;
    S.dur = def.dur ?? t;
    mood = def.mood || mood; S.mood = mood;
    scenes.push(S); T += S.dur;
  }
  return { scenes, duration: T };
}
function sceneAt(T) {
  const sc = TL.scenes;
  for (let i = 0; i < sc.length; i++) if (T < sc[i].start + sc[i].dur) return sc[i];
  return sc[sc.length - 1];
}
function exportTimeline() {
  return {
    fps: FPS, duration: TL.duration,
    scenes: TL.scenes.map(S => ({ id: S.id, start: S.start, dur: S.dur, mood: S.mood, card: !!S.def.card, chapter: S.def.card ? S.def.card.num + '　' + S.def.card.title : null })),
    beats: TL.scenes.flatMap(S => S.beats.map((b, i) => ({ start: S.start + S.b[i], end: S.start + S.e[i], text: b.text || null, sub: b.sub || null })).filter(x => x.text)),
    accents: TL.scenes.flatMap(S => (S.def.accents || []).map(a => ({ t: S.start + (typeof a.at === 'function' ? a.at(S) : a.at), kind: a.kind }))),
  };
}

// ---------- chapter card ----------
function drawCard(t, S, ctx) {
  const c = S.def.card;
  const u1 = eout(P(t, 0.25, 1.3)), u2 = eout(P(t, 0.7, 1.4));
  text(ctx, c.num, 960, 430 + (1 - u1) * 14, { size: 150, w: 900, c: COL.gold, a: u1 });
  seg(ctx, 960 - 120 * u2, 535, 960 + 120 * u2, 535, 1.5, GOLD, 0.55 * u2);
  text(ctx, c.title, 960, 610 + (1 - u2) * 10, { size: 58, w: 400, a: u2, ls: 8 });
}

// ---------- frame ----------
function renderFrame(T) {
  const S = sceneAt(T), D = S.def, t = T - S.start;
  for (const c of [M, L]) {
    c.setTransform(1, 0, 0, 1, 0, 0); c.globalAlpha = 1; c.globalCompositeOperation = 'source-over'; c.filter = 'none';
  }
  M.drawImage(BG, 0, 0);
  L.clearRect(0, 0, W, H);
  drawDust(L, T, D.dust ?? 1);
  if (D.card) glow(L, 960, 440, 380, GOLD, 0.16 * eout(P(t, 0.25, 1.3)));
  if (D.draw) D.draw(t, S, L);
  if (!D.noFade) { // keep the caption band readable
    L.save(); L.globalCompositeOperation = 'destination-out';
    const fg = L.createLinearGradient(0, 840, 0, 1010); fg.addColorStop(0, 'rgba(0,0,0,0)'); fg.addColorStop(1, 'rgba(0,0,0,0.8)');
    L.fillStyle = fg; L.fillRect(0, 840, W, H - 840); L.restore();
  }
  M.drawImage(Lc, 0, 0);
  // bloom: blurred low-res copy of the visual layer, added on top
  SM.clearRect(0, 0, SW, SH); SM.filter = 'blur(2.5px)'; SM.drawImage(Lc, 0, 0, SW, SH); SM.filter = 'none';
  M.globalCompositeOperation = 'lighter'; M.globalAlpha = D.bloom ?? 0.85;
  M.drawImage(SMc, 0, 0, W, H);
  M.globalAlpha = 1; M.globalCompositeOperation = 'source-over';
  if (D.card) drawCard(t, S, M);
  if (D.over) D.over(t, S, M);
  S.beats.forEach((b, k) => { if (b.text && t >= S.b[k] && t <= S.e[k]) drawCaption(M, b, t, S.b[k], S.e[k]); });
  const fi = D.fin ?? 0.6, fo = D.fout ?? 0.6;
  const a = Math.min(smooth(t / fi), smooth((S.dur - t) / fo));
  if (a < 1) { M.fillStyle = `rgba(0,0,0,${1 - a})`; M.fillRect(0, 0, W, H); }
}

async function boot() {
  const cv = document.getElementById('c');
  M = cv.getContext('2d');
  Lc = mkCanvas(W, H); L = Lc.getContext('2d');
  SMc = mkCanvas(SW, SH); SM = SMc.getContext('2d');
  M.imageSmoothingQuality = 'high';
  await Promise.all([
    document.fonts.load(`400 50px ${SERIF}`, '人生'), document.fonts.load(`700 50px ${SERIF}`, '人生'),
    document.fonts.load(`900 50px ${SERIF}`, '人生'), document.fonts.load(`300 50px ${SANS}`, '人生'),
  ]);
  await document.fonts.ready;
  BG = makeBG();
  defineTex();
  await prepTex();
  prepScenes();
  TL = buildTimeline();
  window.TL = TL;
  window.READY = true;
}
window.renderFrame = renderFrame;
window.exportTimeline = exportTimeline;
