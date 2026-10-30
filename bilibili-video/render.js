// Render the whole video: node render.js out_dir [workers]
// Exports timeline.json, renders frames in parallel chunks (each piped
// into its own x264 encoder), then concatenates the chunks.
const { chromium } = require('playwright');
const { spawn } = require('child_process');
const path = require('path'), fs = require('fs');

async function openPage(browser) {
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });
  page.on('pageerror', e => console.error('pageerror:', e.message));
  await page.goto('file://' + path.join(__dirname, 'index.html'));
  await page.waitForFunction(() => window.READY || window.BOOT_ERROR, null, { timeout: 180000 });
  const err = await page.evaluate(() => window.BOOT_ERROR);
  if (err) throw new Error(err);
  return page;
}

async function renderChunk(browser, out, idx, f0, f1, fps) {
  const page = await openPage(browser);
  const file = path.join(out, `chunk_${String(idx).padStart(2, '0')}.mp4`);
  const ff = spawn('ffmpeg', ['-y', '-loglevel', 'error', '-f', 'image2pipe', '-framerate', String(fps), '-c:v', 'mjpeg', '-i', '-',
    '-c:v', 'libx264', '-preset', 'slow', '-crf', '17', '-tune', 'film', '-x264-params', 'aq-mode=3',
    '-pix_fmt', 'yuv420p', '-color_primaries', 'bt709', '-color_trc', 'bt709', '-colorspace', 'bt709', '-r', String(fps), file]);
  ff.stderr.on('data', d => process.stderr.write(d));
  const done = new Promise((res, rej) => ff.on('close', c => (c ? rej(new Error('ffmpeg ' + c)) : res())));
  const t0 = Date.now();
  for (let f = f0; f < f1; f++) {
    await page.evaluate(T => renderFrame(T), f / fps);
    const buf = await page.screenshot({ type: 'jpeg', quality: 96 });
    if (!ff.stdin.write(buf)) await new Promise(r => ff.stdin.once('drain', r));
    if ((f - f0) % 300 === 0) console.log(`chunk ${idx}: ${f - f0}/${f1 - f0} frames, ${((Date.now() - t0) / 1000).toFixed(0)}s`);
  }
  ff.stdin.end(); await done; await page.close();
  return file;
}

(async () => {
  const out = process.argv[2] || 'out'; const workers = +(process.argv[3] || 4);
  fs.mkdirSync(out, { recursive: true });
  const browser = await chromium.launch({ args: ['--allow-file-access-from-files'] });
  const p0 = await openPage(browser);
  const tl = await p0.evaluate(() => exportTimeline());
  await p0.evaluate(() => renderCover());
  await p0.screenshot({ path: path.join(out, 'cover.png') });
  await p0.close();
  fs.writeFileSync(path.join(out, 'timeline.json'), JSON.stringify(tl, null, 1));
  const total = Math.ceil(tl.duration * tl.fps);
  const only = process.env.ONLY_SECONDS ? Math.min(total, Math.ceil(+process.env.ONLY_SECONDS * tl.fps)) : total;
  console.log(`duration ${tl.duration.toFixed(2)}s, rendering ${only} frames with ${workers} workers`);
  const per = Math.ceil(only / workers), jobs = [];
  for (let i = 0; i < workers; i++) { const a = i * per, b = Math.min(only, a + per); if (a < b) jobs.push(renderChunk(browser, out, i, a, b, tl.fps)); }
  const files = await Promise.all(jobs);
  await browser.close();
  fs.writeFileSync(path.join(out, 'chunks.txt'), files.map(f => `file '${path.basename(f)}'`).join('\n') + '\n');
  console.log('frames done');
})().catch(e => { console.error(e); process.exit(1); });
