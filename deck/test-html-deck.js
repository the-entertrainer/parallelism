/**
 * Browser test for parallelism_deck.html.
 *
 * Serve the repo root first:  python3 -m http.server 8123
 * Then:                       node deck/test-html-deck.js
 *
 * Checks structure, keyboard navigation, every interaction, computed
 * visibility (class checks alone can be beaten by Tailwind's cascade),
 * overflow inside the 16:9 frame, and stage scaling at a smaller viewport.
 * Screenshots land in deck/build/html-qa/.
 *
 * The Tailwind CDN bundle is served from /tmp/tw.js so the suite runs without
 * outbound network:  curl -sL https://cdn.tailwindcss.com -o /tmp/tw.js
 */
const { chromium } = require('playwright');
const fs = require('fs');
const OUT = '/home/user/parallelism/deck/build/html-qa';
fs.mkdirSync(OUT, { recursive: true });

(async () => {
  const b = await chromium.launch({
    executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
    args: ['--no-sandbox'],
  });
  const ctx = await b.newContext({ viewport: { width: 1600, height: 900 } });
  // serve the Tailwind CDN bundle from a local cache (sandbox has no direct CDN access for the browser)
  await ctx.route('https://cdn.tailwindcss.com/**', (r) =>
    r.fulfill({ status: 200, contentType: 'application/javascript', body: fs.readFileSync('/tmp/tw.js', 'utf8') }));
  await ctx.route('https://cdn.tailwindcss.com', (r) =>
    r.fulfill({ status: 200, contentType: 'application/javascript', body: fs.readFileSync('/tmp/tw.js', 'utf8') }));

  const page = await ctx.newPage();
  const errors = [];
  page.on('console', (m) => { if (m.type() === 'error') errors.push('CONSOLE: ' + m.text()); });
  page.on('pageerror', (e) => errors.push('PAGEERROR: ' + e.message));

  await page.goto('http://127.0.0.1:8123/parallelism_deck.html', { waitUntil: 'networkidle' });
  await page.waitForTimeout(900);


  // --- computed-visibility checks (class-based checks can be beaten by the cascade)
  const visible = (sel) => page.$eval(sel, (e) => {
    const st = getComputedStyle(e);
    return st.display !== 'none' && st.visibility !== 'hidden' && e.getBoundingClientRect().height > 0;
  });
  const check = (label, cond, extra = '') =>
    console.log(`${cond ? 'PASS' : 'FAIL'}  ${label}${extra ? '  → ' + extra : ''}`);

  // --- structure
  const total = await page.$$eval('.slide', (n) => n.length);
  check('7 slides present', total === 7, `found ${total}`);
  check('counter total = 07', (await page.textContent('#counter-total')) === '07');

  check('hook answer hidden on load', !(await visible('#hook-reveal')));
  // --- slide 1 interaction: wrong then right
  await page.click('.slide[data-slide="1"] .hook-opt[data-correct="false"]');
  await page.waitForTimeout(250);
  let fb = await page.textContent('#hook-feedback');
  check('wrong hook pick gives rose feedback', /Keep looking/.test(fb), fb.slice(0, 40));
  await page.click('.slide[data-slide="1"] .hook-opt[data-correct="true"]');
  await page.waitForTimeout(350);
  check('correct hook pick reveals repair',
    await page.$eval('#hook-reveal', (e) => !e.classList.contains('hidden-soft')));

  // --- keyboard navigation across the whole deck
  for (let i = 2; i <= 7; i++) {
    await page.keyboard.press('ArrowRight');
    await page.waitForTimeout(220);
    const cur = await page.textContent('#counter-current');
    const active = await page.$eval('.slide.is-active', (e) => e.dataset.slide);
    check(`ArrowRight → slide ${i}`, cur === String(i).padStart(2, '0') && active === String(i), `counter ${cur}, active ${active}`);
  }
  await page.keyboard.press('ArrowLeft');
  await page.waitForTimeout(220);
  check('ArrowLeft goes back', (await page.textContent('#counter-current')) === '06');

  // --- progress bar
  const w = await page.$eval('#progress', (e) => e.style.width);
  check('progress bar tracks position', w.startsWith('85.7'), w);

  // --- reveal interaction on a breakdown slide
  await page.keyboard.press('Home');
  await page.waitForTimeout(250);
  await page.keyboard.press('ArrowRight'); await page.keyboard.press('ArrowRight');
  await page.waitForTimeout(300);
  check('on slide 3', (await page.$eval('.slide.is-active', (e) => e.dataset.slide)) === '3');
  check('after-card hidden before reveal (computed)', !(await visible('.slide[data-slide="3"] .after-card')));
  check('placeholder visible before reveal (computed)', await visible('.slide[data-slide="3"] .before-card'));
  await page.keyboard.press('Enter');
  await page.waitForTimeout(300);
  check('Enter reveals the fix',
    await page.$eval('.slide[data-slide="3"] .after-card', (e) => !e.classList.contains('hidden-soft')));
  check('placeholder hidden after reveal (computed)', !(await visible('.slide[data-slide="3"] .before-card')));
  check('after-card visible after reveal (computed)', await visible('.slide[data-slide="3"] .after-card'));
  await page.keyboard.press('Enter');
  await page.waitForTimeout(250);
  check('Enter toggles back',
    await page.$eval('.slide[data-slide="3"] .after-card', (e) => e.classList.contains('hidden-soft')));
  await page.keyboard.press('Enter');
  await page.waitForTimeout(250);

  // --- presenter notes
  await page.keyboard.press('n');
  await page.waitForTimeout(250);
  check('N opens notes', await page.$eval('#notes-panel', (e) => !e.hidden));
  const notes = (await page.textContent('#notes-text')).trim();
  check('notes are slide-specific', notes.length > 60, notes.slice(0, 45) + '…');
  await page.keyboard.press('n');
  await page.waitForTimeout(250);
  check('N closes notes', await page.$eval('#notes-panel', (e) => e.hidden));

  // --- quiz
  await page.keyboard.press('End');
  await page.waitForTimeout(350);
  await page.click('.quiz-opt[data-q="0"][data-ok="false"]');
  await page.waitForTimeout(200);
  check('wrong quiz answer does not score', (await page.textContent('#quiz-score')) === '0');
  for (const q of [0, 1, 2]) {
    await page.click(`.quiz-opt[data-q="${q}"][data-ok="true"]`);
    await page.waitForTimeout(200);
  }
  check('score reaches 3', (await page.textContent('#quiz-score')) === '3');
  check('completion message shows',
    await page.$eval('#quiz-done', (e) => !e.classList.contains('hidden-soft')));
  await page.click('#quiz-reset');
  await page.waitForTimeout(250);
  check('reset clears score', (await page.textContent('#quiz-score')) === '0');

  // --- overflow / layout audit on every slide, plus screenshots
  await page.keyboard.press('Home');
  await page.waitForTimeout(700);
  for (let i = 1; i <= 7; i++) {
    if (i > 1) { await page.keyboard.press('ArrowRight'); }
    await page.waitForTimeout(700);
    const info = await page.evaluate(() => {
      const s = document.querySelector('.slide.is-active');
      const r = s.getBoundingClientRect();
      const bad = [];
      s.querySelectorAll('*').forEach((el) => {
        if (!el.offsetParent && el.tagName !== 'BODY') return;
        const b = el.getBoundingClientRect();
        if (b.width === 0 || b.height === 0) return;
        if (b.top < r.top - 1 || b.bottom > r.bottom + 1 || b.left < r.left - 1 || b.right > r.right + 1) {
          bad.push(el.tagName + '.' + (el.className || '').toString().slice(0, 42));
        }
      });
      const oflow = [];
      s.querySelectorAll('p,h1,h2,button,div').forEach((el) => {
        if (el.scrollHeight > el.clientHeight + 2 && getComputedStyle(el).overflow !== 'visible') {
          oflow.push(el.tagName + '.' + (el.className || '').toString().slice(0, 42));
        }
      });
      return { out: bad.slice(0, 6), oflow: oflow.slice(0, 6) };
    });
    check(`slide ${i}: nothing outside the 16:9 frame`, info.out.length === 0, info.out.join(' | '));
    check(`slide ${i}: no clipped text`, info.oflow.length === 0, info.oflow.join(' | '));
    await page.screenshot({ path: `${OUT}/slide-0${i}.png` });
  }

  // notes drawer screenshot
  await page.keyboard.press('Home'); await page.waitForTimeout(250);
  await page.keyboard.press('n'); await page.waitForTimeout(300);
  await page.screenshot({ path: `${OUT}/notes-open.png` });

  // --- smaller viewport (scaling)
  const p2 = await ctx.newPage();
  await p2.setViewportSize({ width: 1280, height: 720 });
  await p2.goto('http://127.0.0.1:8123/parallelism_deck.html', { waitUntil: 'networkidle' });
  await p2.waitForTimeout(700);
  const t = await p2.$eval('#stage', (e) => e.style.transform);
  check('stage scales to 1280×720', /scale\(0\.8/.test(t), t);
  const box = await p2.$eval('#stage', (e) => { const r = e.getBoundingClientRect(); return { l: r.left, t: r.top, r: r.right, b: r.bottom, w: r.width, h: r.height }; });
  check('stage fits inside 1280×720', box.l >= -1 && box.t >= -1 && box.r <= 1281 && box.b <= 721,
    `l${box.l.toFixed(0)} t${box.t.toFixed(0)} r${box.r.toFixed(0)} b${box.b.toFixed(0)}`);
  check('stage is centred', Math.abs((box.l + box.r) / 2 - 640) < 2 && Math.abs((box.t + box.b) / 2 - 360) < 2,
    `centre ${((box.l + box.r) / 2).toFixed(0)},${((box.t + box.b) / 2).toFixed(0)}`);
  await p2.screenshot({ path: `${OUT}/small-viewport.png` });

  console.log(errors.length ? 'JS ERRORS:\n' + errors.join('\n') : 'PASS  no console or page errors');
  await b.close();
})().catch((e) => { console.error('HARNESS ERROR', e); process.exit(1); });
