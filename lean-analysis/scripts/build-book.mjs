#!/usr/bin/env node
// 把 Analysis/ 下的章节 .lean 文件转换为一本可交互的网页教材 book/index.html。
//
// 用法：node scripts/build-book.mjs [--fragment <path>]
//
// 章节文件中 `/-! ... -/` 注释里的内容按 Markdown 渲染为正文，其余部分作为 Lean 代码。
// 含有 `sorry` 的代码块被视为练习，其参考答案从 Analysis/Solutions/ 中同名定理读取。
// 生成器会检查每道练习在章节文件与答案文件中的陈述完全一致。
//
// --fragment 额外输出一个不带 <html>/<head>/<body> 外壳的版本（用于嵌入其他页面）。

import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const REPO_URL = 'https://github.com/kaoxi998533/pie-slang/tree/ccr-3f61a4ac-j9t5cv/lean-analysis';

const CHAPTERS = [
  { id: 'ch00', file: 'Ch00_Welcome', summary: '如何使用本书；norm_num、ring、linarith 等第一批策略。' },
  { id: 'ch01', file: 'Ch01_Logic', summary: '蕴含、合取、析取、量词与否定：ε-N 语言的骨架。' },
  { id: 'ch02', file: 'Ch02_Reals', summary: '有序域中的不等式，绝对值与三角不等式，伯努利不等式。' },
  { id: 'ch03', file: 'Ch03_Completeness', summary: '确界原理、上确界的 ε 刻画与阿基米德性质。' },
  { id: 'ch04', file: 'Ch04_SeqLimit', summary: 'ε-N 定义，1/(n+1) → 0，极限的唯一性。' },
  { id: 'ch05', file: 'Ch05_LimitLaws', summary: '和、积的极限，收敛数列有界，夹逼定理与保序性。' },
  { id: 'ch06', file: 'Ch06_Monotone', summary: '单调收敛定理，柯西数列与柯西收敛准则。' },
  { id: 'ch07', file: 'Ch07_Continuity', summary: 'ε-δ 连续性，连续函数的运算，海涅定理。' },
  { id: 'ch08', file: 'Ch08_IVT', summary: '介值定理、最值定理与不动点定理。' },
  { id: 'ch09', file: 'Ch09_Derivative', summary: '导数的线性逼近定义，可导必连续，拉格朗日中值定理。' },
];

// ---------------------------------------------------------------------------
// Lean 文件切分

/** 把 Lean 源码切分为正文（`/-! -/`）与代码段，记录每段在源码中的位置。 */
function splitLean(src) {
  const segs = [];
  let i = 0;
  while (i < src.length) {
    let p = src.indexOf('/-!', i);
    while (p > 0 && src[p - 1] !== '\n') p = src.indexOf('/-!', p + 3);
    if (p < 0) {
      segs.push({ kind: 'code', text: src.slice(i), end: src.length });
      break;
    }
    if (p > i) segs.push({ kind: 'code', text: src.slice(i, p), end: p });
    let depth = 1;
    let j = p + 3;
    while (j < src.length && depth > 0) {
      if (src.startsWith('/-', j)) { depth++; j += 2; }
      else if (src.startsWith('-/', j)) { depth--; j += 2; }
      else j++;
    }
    segs.push({ kind: 'prose', text: src.slice(p + 3, j - 2), end: j });
    i = j;
  }
  return segs;
}

const HIDDEN_LINE = /^(import\s|namespace\s|end\s+Ch\d+|open\s)/;

function cleanCode(text) {
  const lines = text.split('\n').filter((l) => !HIDDEN_LINE.test(l));
  while (lines.length && !lines[0].trim()) lines.shift();
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
  return lines.join('\n');
}

const DECL_NAME = /^(?:theorem|lemma)\s+([^\s({[:]+)/gm;

/** 从答案文件中提取各个定理的完整文本。 */
function parseSolutions(src) {
  const map = new Map();
  const lines = src.split('\n');
  for (let k = 0; k < lines.length; k++) {
    const m = /^(?:theorem|lemma)\s+([^\s({[:]+)/.exec(lines[k]);
    if (!m) continue;
    const body = [lines[k]];
    let j = k + 1;
    while (j < lines.length && (lines[j].trim() === '' || /^\s/.test(lines[j]))) body.push(lines[j++]);
    while (body.length && !body[body.length - 1].trim()) body.pop();
    map.set(m[1], body.join('\n'));
    k = j - 1;
  }
  return map;
}

/** 定理陈述（`:=` 之前的部分），用于核对练习与答案是否一致。 */
function statementOf(declText) {
  const idx = declText.indexOf(':=');
  return (idx < 0 ? declText : declText.slice(0, idx)).replace(/\s+/g, ' ').trim();
}

function declTextIn(code, name) {
  const lines = code.split('\n');
  const start = lines.findIndex((l) => new RegExp(`^(?:theorem|lemma)\\s+${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?![^\\s({[:])`).test(l));
  if (start < 0) return null;
  const body = [lines[start]];
  for (let j = start + 1; j < lines.length && (lines[j].trim() === '' || /^\s/.test(lines[j])); j++) body.push(lines[j]);
  return body.join('\n');
}

// ---------------------------------------------------------------------------
// 一个够用的 Markdown 渲染器（标题、段落、列表、引用、表格、代码块、公式）

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

const CJK = /[⺀-鿿＀-￯　-〿]/;
function joinLines(lines) {
  let out = '';
  for (const raw of lines) {
    const l = raw.trim();
    if (!out) { out = l; continue; }
    const a = out[out.length - 1];
    const b = l[0];
    out += (CJK.test(a) || CJK.test(b)) ? l : ' ' + l;
  }
  return out;
}

// 公式先在构建时转换为 Unicode/HTML 的近似形式，页面加载 KaTeX 后再替换为排版好的 MathML。
const TEX_SYMBOLS = {
  varepsilon: 'ε', epsilon: 'ϵ', delta: 'δ', eta: 'η', alpha: 'α', beta: 'β', lambda: 'λ', mu: 'μ', pi: 'π',
  le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠', to: '→', rightarrow: '→', Rightarrow: '⇒', iff: '⇔',
  mapsto: '↦', forall: '∀', exists: '∃', in: '∈', notin: '∉', subseteq: '⊆', subset: '⊂', lnot: '¬', neg: '¬',
  land: '∧', lor: '∨', cdot: '·', times: '×', pm: '±', circ: '∘', infty: '∞', equiv: '≡', mid: ' | ',
  ldots: '…', cdots: '⋯', lvert: '|', rvert: '|', sum: '∑', sup: 'sup', inf: 'inf', lim: 'lim', max: 'max',
  min: 'min', quad: '  ',
};
function texToHtml(tex) {
  const SUB0 = '\u0001', SUB1 = '\u0002', SUP0 = '\u0003', SUP1 = '\u0004';
  const wrap = (x) => (/^[\w.!αβδεηλμπ]+$/u.test(x.trim()) ? x.trim() : `(${x.trim()})`);
  let s = tex.replace(/\\\{/g, '\u0005').replace(/\\\}/g, '\u0006');
  s = s.replace(/\\begin\{cases\}([\s\S]*?)\\end\{cases\}/g, (_, body) =>
    '{ ' + body.split('\\\\').map((r) => r.replace(/&/g, ' ').trim()).join('；') + ' }');
  s = s.replace(/\\(?:text|mathrm|operatorname)\{([^{}]*)\}/g, '$1');
  s = s.replace(/\\mathbb\{(\w)\}/g, (_, c) => ({ N: 'ℕ', R: 'ℝ', Q: 'ℚ', Z: 'ℤ', C: 'ℂ' })[c] || c);
  s = s.replace(/\\(left|right|displaystyle|limits)\b/g, '');
  s = s.replace(/\\([A-Za-z]+)/g, (m, name) => (name in TEX_SYMBOLS ? TEX_SYMBOLS[name] : m));
  for (let k = 0; k < 3; k++) {
    s = s.replace(/\\frac\{([^{}]*)\}\{([^{}]*)\}/g, (_, x, y) => `${wrap(x)}/${wrap(y)}`);
    s = s.replace(/\\sqrt\{([^{}]*)\}/g, (_, x) => `√${wrap(x)}`);
  }
  s = s.replace(/\\[,;: ]/g, ' ').replace(/\\!/g, '');
  for (let k = 0; k < 2; k++) {
    s = s.replace(/_\{([^{}]*)\}/g, `${SUB0}$1${SUB1}`).replace(/\^\{([^{}]*)\}/g, `${SUP0}$1${SUP1}`);
  }
  s = s.replace(/_([^\s{}\\])/gu, `${SUB0}$1${SUB1}`).replace(/\^([^\s{}\\])/gu, `${SUP0}$1${SUP1}`);
  s = s.replace(/[{}]/g, '').replace(/\s+/g, ' ').trim().replace(/-/g, '−');
  s = s.replace(/\u0005/g, '{').replace(/\u0006/g, '}');
  s = esc(s).replace(/(^|[^A-Za-z&])([A-Za-z])(?![A-Za-z;])/g, '$1<i>$2</i>');
  return s.split(SUB0).join('<sub>').split(SUB1).join('</sub>').split(SUP0).join('<sup>').split(SUP1).join('</sup>');
}
const mathSpan = (tex, display) =>
  `<span class="math"${display ? ' data-display="1"' : ''} data-tex="${esc(tex)}">${texToHtml(tex)}</span>`;

function inline(text) {
  const slots = [];
  const put = (html) => `\u0000${slots.push(html) - 1}\u0000`;
  let s = text;
  s = s.replace(/`([^`]+)`/g, (_, c) => put(`<code>${esc(c)}</code>`));
  s = s.replace(/\$\$([\s\S]+?)\$\$/g, (_, t) => put(mathSpan(t.trim(), true)));
  s = s.replace(/\$([^$\n]+?)\$/g, (_, t) => put(mathSpan(t, false)));
  s = esc(s);
  s = s.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, (_, t, u) => put(`<a href="${u}" target="_blank" rel="noopener">${t}</a>`));
  s = s.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  s = s.replace(/(^|[^*\w])\*([^*\s][^*]*?)\*(?!\*)/g, '$1<em>$2</em>');
  for (let k = 0; k < 3; k++) s = s.replace(/\u0000(\d+)\u0000/g, (_, n) => slots[+n]);
  return s;
}

function dedent(text) {
  const lines = text.replace(/\t/g, '  ').split('\n');
  const ind = Math.min(...lines.filter((l) => l.trim()).map((l) => l.match(/^ */)[0].length));
  return lines.map((l) => l.slice(Number.isFinite(ind) ? ind : 0));
}

const LIST_ITEM = /^(\s*)([*-]|\d+\.)\s+(.*)$/;
const CALLOUTS = [
  [/^<strong>(定义|Definition)/, 'def'],
  [/^<strong>(定理|公理|推论|引理)/, 'thm'],
  [/^<strong>(例)/, 'example'],
  [/^<strong>(提示|注)/, 'note'],
];

function md(text, ctx) {
  const lines = Array.isArray(text) ? text : dedent(text);
  const out = [];
  let i = 0;
  const isBlank = (l) => !l || !l.trim();
  const startsBlock = (l) => /^(#{1,6}\s|```|>|\|)/.test(l) || LIST_ITEM.test(l) || /^\$\$/.test(l.trim()) || /^<!--/.test(l.trim());
  let skipWebNote = false;
  while (i < lines.length) {
    const line = lines[i];
    if (isBlank(line)) { i++; continue; }
    let m;
    if ((m = /^<!--\s*widget:([\w-]+)\s*-->$/.exec(line.trim()))) {
      out.push(`<div class="widget" data-widget="${m[1]}"></div>`);
      skipWebNote = true;
      i++;
      continue;
    }
    if ((m = /^```(\w*)\s*$/.exec(line))) {
      const lang = m[1] || 'text';
      const body = [];
      i++;
      while (i < lines.length && !/^```\s*$/.test(lines[i])) body.push(lines[i++]);
      i++;
      out.push(`<pre class="snippet${lang === 'lean' ? ' lean-src' : ''}"><code>${esc(body.join('\n'))}</code></pre>`);
      continue;
    }
    if ((m = /^(#{1,6})\s+(.*)$/.exec(line))) {
      const level = m[1].length;
      if (level === 1) { ctx.title = m[2].trim(); i++; continue; }
      const id = `${ctx.id}-s${++ctx.headingCount}`;
      out.push(`<h${level} id="${id}">${inline(m[2])}</h${level}>`);
      i++;
      continue;
    }
    if (line.startsWith('>')) {
      const body = [];
      while (i < lines.length && lines[i].startsWith('>')) body.push(lines[i++].replace(/^> ?/, ''));
      const inner = md(body, ctx);
      const first = inner.replace(/^<p>/, '');
      const kind = (CALLOUTS.find(([re]) => re.test(first)) || [null, 'plain'])[1];
      out.push(`<blockquote class="callout callout-${kind}">${inner}</blockquote>`);
      continue;
    }
    if (line.startsWith('|') && i + 1 < lines.length && /^\|?\s*:?-{3,}/.test(lines[i + 1])) {
      const rows = [];
      while (i < lines.length && lines[i].startsWith('|')) rows.push(lines[i++]);
      const cells = (r) => r.trim().replace(/^\||\|$/g, '').split(/(?<!\\)\|/).map((c) => inline(c.trim().replace(/\\\|/g, '|')));
      const head = cells(rows[0]);
      const bodyRows = rows.slice(2).map(cells);
      out.push(`<div class="table-wrap"><table><thead><tr>${head.map((c) => `<th>${c}</th>`).join('')}</tr></thead><tbody>${bodyRows.map((r) => `<tr>${r.map((c) => `<td>${c}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`);
      continue;
    }
    if (/^\$\$/.test(line.trim())) {
      const body = [line.trim().slice(2)];
      if (!/\$\$$/.test(body[0]) || body[0] === '') {
        i++;
        while (i < lines.length && !/\$\$\s*$/.test(lines[i])) body.push(lines[i++]);
        if (i < lines.length) body.push(lines[i]);
      }
      i++;
      const tex = body.join('\n').replace(/\$\$\s*$/, '').trim();
      out.push(`<div class="math-block">${mathSpan(tex, true)}</div>`);
      continue;
    }
    if (LIST_ITEM.test(line)) {
      const ordered = /^\s*\d+\./.test(line);
      const items = [];
      while (i < lines.length) {
        const l = lines[i];
        const im = LIST_ITEM.exec(l);
        if (im && im[1].length === 0) { items.push([im[3]]); i++; continue; }
        if (!isBlank(l) && /^\s+/.test(l) && items.length) { items[items.length - 1].push(l); i++; continue; }
        if (isBlank(l) && i + 1 < lines.length && LIST_ITEM.test(lines[i + 1])) { i++; continue; }
        break;
      }
      const tag = ordered ? 'ol' : 'ul';
      out.push(`<${tag}>${items.map((it) => `<li>${inline(joinLines(it))}</li>`).join('')}</${tag}>`);
      continue;
    }
    const para = [];
    while (i < lines.length && !isBlank(lines[i]) && !(para.length && startsBlock(lines[i]))) para.push(lines[i++]);
    const joined = joinLines(para);
    if (skipWebNote && joined.startsWith('（网页版')) { skipWebNote = false; continue; }
    skipWebNote = false;
    out.push(`<p>${inline(joined)}</p>`);
  }
  return out.join('\n');
}

// ---------------------------------------------------------------------------
// 组装各章

function buildChapter(meta) {
  const source = readFileSync(join(ROOT, 'Analysis', `${meta.file}.lean`), 'utf8');
  const solSrc = readFileSync(join(ROOT, 'Analysis', 'Solutions', `${meta.file}.lean`), 'utf8');
  const solutions = parseSolutions(solSrc);
  const ctx = { id: meta.id, title: '', headingCount: 0 };
  const blocks = [];
  let exCount = 0;
  for (const seg of splitLean(source)) {
    if (seg.kind === 'prose') {
      const html = md(seg.text, ctx);
      if (html.trim()) blocks.push({ t: 'p', html });
      continue;
    }
    const code = cleanCode(seg.text);
    if (!code) continue;
    const block = { t: 'c', code, upto: seg.end };
    if (/\bsorry\b/.test(code)) {
      const names = [...code.matchAll(DECL_NAME)].map((m) => m[1]);
      if (!names.length) throw new Error(`${meta.file}: 含 sorry 的代码块没有具名定理`);
      const sols = names.map((n) => {
        const sol = solutions.get(n);
        if (!sol) throw new Error(`${meta.file}: 找不到练习 ${n} 的参考答案`);
        const a = statementOf(declTextIn(code, n));
        const b = statementOf(sol);
        if (a !== b) throw new Error(`${meta.file}: 练习 ${n} 的陈述与答案不一致\n  章节: ${a}\n  答案: ${b}`);
        return sol;
      });
      exCount++;
      block.ex = { id: `${meta.id}:${names[0]}`, name: names[0], solution: sols.join('\n\n') };
    }
    blocks.push(block);
  }
  const tm = /^第\s*(\d+)\s*章\s*(.*)$/.exec(ctx.title);
  return {
    id: meta.id,
    num: tm ? Number(tm[1]) : 0,
    title: tm ? tm[2].trim() : ctx.title,
    summary: meta.summary,
    file: `Analysis/${meta.file}.lean`,
    exercises: exCount,
    source,
    blocks,
  };
}

const chapters = CHAPTERS.map(buildChapter);
const totalEx = chapters.reduce((s, c) => s + c.exercises, 0);
const data = JSON.stringify({ chapters, repo: REPO_URL }).replace(/</g, '\\u003c');

const template = readFileSync(join(ROOT, 'scripts', 'book-template.html'), 'utf8')
  .replace('/*__BOOK_DATA__*/', () => data)
  .replace(/__TOTAL_EXERCISES__/g, String(totalEx))
  .replace(/__CHAPTER_COUNT__/g, String(chapters.length));
const [head, body] = template.split('<!--HEAD-END-->');
const fragment = head + body;

const page = `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
${head.trim()}
</head>
<body>
${body.trim()}
</body>
</html>
`;

mkdirSync(join(ROOT, 'book'), { recursive: true });
writeFileSync(join(ROOT, 'book', 'index.html'), page);
const fi = process.argv.indexOf('--fragment');
if (fi > 0 && process.argv[fi + 1]) writeFileSync(process.argv[fi + 1], fragment);
console.log(`已生成 book/index.html：${chapters.length} 章，${totalEx} 道练习。`);
