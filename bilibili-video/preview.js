// Render still frames at given times: node preview.js out_dir t1 t2 ... | or "scene:id:offset"
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');
(async () => {
  const [out, ...ts] = process.argv.slice(2);
  fs.mkdirSync(out, { recursive: true });
  const browser = await chromium.launch({ args: ['--allow-file-access-from-files'] });
  const page = await browser.newPage({ viewport: { width: 1920, height: 1080 } });
  page.on('console', m => { if (m.type() === 'error' || m.type() === 'warning') console.log('console:', m.text()); });
  page.on('pageerror', e => console.log('pageerror:', e.message));
  await page.goto('file://' + path.join(__dirname, 'index.html'));
  await page.waitForFunction(() => window.READY || window.BOOT_ERROR, null, { timeout: 120000 });
  const err = await page.evaluate(() => window.BOOT_ERROR);
  if (err) { console.log('BOOT_ERROR', err); process.exit(1); }
  const info = await page.evaluate(() => ({ d: TL.duration, s: TL.scenes.map(s => [s.id, +s.start.toFixed(2), +s.dur.toFixed(2)]) }));
  console.log('duration', info.d.toFixed(1)); if (!ts.length) console.log(JSON.stringify(info.s));
  for (const spec of ts) {
    let T = parseFloat(spec);
    if (spec.startsWith('s:')) { const [, id, off] = spec.split(':'); T = await page.evaluate(([id, off]) => { const S = TL.scenes.find(s => s.id === id); return S.start + (off.startsWith('b') ? S.b[+off.slice(1).split('+')[0]] + (+(off.split('+')[1] || 0)) : +off); }, [id, off]); }
    const t0 = Date.now();
    if (spec === 'cover') await page.evaluate(() => renderCover()); else await page.evaluate(T => renderFrame(T), T);
    const name = spec.replace(/[:+]/g, '_') + '.jpg';
    await page.screenshot({ path: path.join(out, name), type: 'jpeg', quality: 90 });
    console.log(name, T.toFixed(2), (Date.now() - t0) + 'ms');
  }
  await browser.close();
})();
