const { launch } = require('./browser');
const fs = require('fs');

const MOD = process.argv[2];
const DATA = JSON.parse(fs.readFileSync(`findings_${MOD}.json`, 'utf8'));
const OUT = `shots_annotated/${MOD}`;
fs.mkdirSync(OUT, { recursive: true });

const STYLE = `
.cqa-mark { background: #FFF176 !important; box-shadow: 0 0 0 2px #D50000; border-radius: 2px; }
.cqa-badge { display:inline-block; background:#D50000; color:#fff; font-weight:700;
  font-family: Arial, Helvetica, sans-serif; font-size:13px; line-height:18px; min-width:18px;
  text-align:center; border-radius:9px; padding:0 5px; margin-right:4px; vertical-align:middle; }
.cqa-box { outline: 3px solid #D50000 !important; outline-offset: 2px; }
`;

// ---- in-page helpers -------------------------------------------------------
const PAGE_FNS = () => {
  const norm = (s) => s.replace(/[‘’]/g, "'").replace(/[“”]/g, '"')
                       .replace(/[–—]/g, '-').replace(/\s+/g, ' ');
  window.__cqa = {
    norm,
    highlight(needleRaw, n) {
      const needle = norm(needleRaw).trim();
      // In a quiz, previously answered cards stay in the DOM stack, so search only the
      // card the learner is currently looking at.
      const root = document.querySelector('.quiz__card--active')
        || document.querySelector('.blocks-lesson') || document.body;
      const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      const nodes = [];
      while (walker.nextNode()) nodes.push(walker.currentNode);
      // 1) single text node contains the needle
      for (const tn of nodes) {
        if (!tn.nodeValue || !tn.nodeValue.trim()) continue;
        const hay = norm(tn.nodeValue);
        const i = hay.indexOf(needle);
        if (i === -1) continue;
        if (tn.parentElement && tn.parentElement.closest('.cqa-mark')) continue;
        // skip matches inside elements that are not actually on screen (hidden quiz
        // cards, collapsed panels, the reverse face of a flashcard)
        const pe = tn.parentElement;
        if (pe) {
          const pr = pe.getBoundingClientRect();
          if (pr.width === 0 && pr.height === 0) continue;
          const cs = getComputedStyle(pe);
          if (cs.visibility === 'hidden' || cs.display === 'none') continue;
        }
        const range = document.createRange();
        range.setStart(tn, Math.min(i, tn.nodeValue.length));
        range.setEnd(tn, Math.min(i + needle.length, tn.nodeValue.length));
        const mark = document.createElement('mark');
        mark.className = 'cqa-mark';
        const badge = document.createElement('span');
        badge.className = 'cqa-badge';
        badge.textContent = n;
        badge.setAttribute('data-idx', n);
        try {
          range.surroundContents(mark);
          mark.insertBefore(badge, mark.firstChild);
          return true;
        } catch (e) { /* fall through */ }
      }
      // 2) fall back: smallest element whose text contains the needle -> outline it
      let best = null;
      const all = root.querySelectorAll('*');
      for (const el of all) {
        const t = norm(el.innerText || el.textContent || '');
        if (t.includes(needle)) {
          if (!best || (el.innerText || '').length < (best.innerText || '').length) best = el;
        }
      }
      if (best) {
        best.classList.add('cqa-box');
        const badge = document.createElement('span');
        badge.className = 'cqa-badge';
        badge.textContent = n;
        badge.setAttribute('data-idx', n);
        best.insertBefore(badge, best.firstChild);
        return true;
      }
      return false;
    },
    boxBySelector(sel, nth, n) {
      const root = document.querySelector('.blocks-lesson') || document.body;
      const els = root.querySelectorAll(sel);
      const el = els[nth];
      if (!el) return false;
      el.classList.add('cqa-box');
      const badge = document.createElement('span');
      badge.className = 'cqa-badge';
      badge.textContent = n;
      badge.setAttribute('data-idx', n);
      el.insertBefore(badge, el.firstChild);
      return true;
    },
    bounds() {
      const els = document.querySelectorAll('.cqa-mark, .cqa-box');
      if (!els.length) return null;
      let top = 1e9, bottom = -1e9, left = 1e9, right = -1e9;
      els.forEach(e => {
        const r = e.getBoundingClientRect();
        if (r.height === 0 && r.width === 0) return;
        top = Math.min(top, r.top); bottom = Math.max(bottom, r.bottom);
        left = Math.min(left, r.left); right = Math.max(right, r.right);
      });
      if (top === 1e9) return null;
      return { top, bottom, left, right };
    },
    scroller() { return document.querySelector('.page-wrap') || document.scrollingElement; },
    renumber() {
      const sc = document.querySelector('.page-wrap') || document.scrollingElement;
      const badges = Array.from(document.querySelectorAll('.cqa-badge'))
        .filter(el => { const r = el.getBoundingClientRect(); return r.width || r.height; });
      badges.sort((a, b) => {
        const ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
        return (ra.top - rb.top) || (ra.left - rb.left);
      });
      const order = [];
      badges.forEach((el, i) => {
        order.push(parseInt(el.getAttribute('data-idx'), 10));
        el.textContent = String(i + 1);
      });
      return order; // order[displayNumber-1] = original item number
    },
    offsets() {
      const sc = document.querySelector('.page-wrap') || document.scrollingElement;
      const els = document.querySelectorAll('.cqa-mark, .cqa-box');
      if (!els.length) return null;
      let min = 1e9, max = -1e9;
      els.forEach(e => {
        const r = e.getBoundingClientRect();
        if (!r.height && !r.width) return;
        min = Math.min(min, r.top + sc.scrollTop);
        max = Math.max(max, r.bottom + sc.scrollTop);
      });
      return min === 1e9 ? null : { min, max };
    },
    firstMarkOffset() {
      const e = document.querySelector('.cqa-mark, .cqa-box');
      if (!e) return 0;
      const sc = document.querySelector('.page-wrap') || document.scrollingElement;
      return e.getBoundingClientRect().top + sc.scrollTop;
    }
  };
};

async function scrollThrough(p) {
  await p.evaluate(async () => {
    const c = document.querySelector('.page-wrap') || document.scrollingElement;
    for (let y = 0; y < c.scrollHeight + 1500; y += 400) { c.scrollTop = y; await new Promise(r => setTimeout(r, 70)); }
    c.scrollTop = 0;
  });
}

async function runPrep(p, prep) {
  for (const step of prep || []) {
    if (step.type === 'accordionAll') {
      for (const h of await p.$$('.blocks-accordion__header')) {
        try { if (await h.getAttribute('aria-expanded') !== 'true') { await h.click({ timeout: 2500 }); await p.waitForTimeout(300); } } catch (e) {}
      }
    } else if (step.type === 'tab') {
      const tabs = await p.$$('[role="tab"], .blocks-tabs__label, .tabs__label');
      if (tabs[step.index]) { try { await tabs[step.index].click({ timeout: 3000 }); await p.waitForTimeout(900); } catch (e) {} }
    } else if (step.type === 'carousel') {
      for (let i = 0; i < step.index; i++) {
        try { await p.click('.carousel-controls-next', { timeout: 3000 }); await p.waitForTimeout(700); } catch (e) {}
      }
    } else if (step.type === 'flip') {
      try {
        const btns = await p.$$('.carousel-slide .flashcard-side-flip__btn');
        for (const b of btns) { if (await b.isVisible()) { await b.click({ timeout: 2500 }); break; } }
        await p.waitForTimeout(1100);
      } catch (e) {}
    } else if (step.type === 'quizStart') {
      try { await p.click('.quiz-start__start-button', { timeout: 8000 }); await p.waitForTimeout(2200); } catch (e) {}
    } else if (step.type === 'quizFind') {
      let hit = false;
      for (let i = 0; i < (step.max || 8); i++) {
        const txt = await p.evaluate(() => {
          const c = document.querySelector('.quiz__card--active');
          return c ? c.innerText : '';
        });
        if (txt && txt.toLowerCase().includes(step.text.toLowerCase())) { hit = true; break; }
        try {
          const opts = await p.$$('.quiz__card--active .quiz-multiple-choice-option');
          if (opts.length) { await opts[0].click({ timeout: 2500 }); await p.waitForTimeout(400); }
          await p.click('.quiz__card--active .quiz-card__submit', { timeout: 5000 });
          await p.waitForTimeout(1400);
          await p.click('.quiz__card--active .quiz-card__button--next', { timeout: 5000 });
          await p.waitForTimeout(1700);
        } catch (e) { break; }
      }
      if (!hit) console.log('  !! quizFind missed:', step.text);
    } else if (step.type === 'quizSubmit') {
      try {
        const opts = await p.$$('.quiz__card--active .quiz-multiple-choice-option');
        if (opts.length) { await opts[0].click({ timeout: 2500 }); await p.waitForTimeout(400); }
        await p.click('.quiz__card--active .quiz-card__submit', { timeout: 5000 });
        await p.waitForTimeout(1600);
      } catch (e) {}
    } else if (step.type === 'quizAnswer') {
      for (let i = 0; i < step.n; i++) {
        try {
          const opts = await p.$$('.quiz__card--active .quiz-multiple-choice-option');
          if (opts.length) { await opts[0].click({ timeout: 2500 }); await p.waitForTimeout(500); }
          await p.click('.quiz__card--active .quiz-card__submit', { timeout: 5000 });
          await p.waitForTimeout(1500);
          await p.click('.quiz__card--active .quiz-card__button--next', { timeout: 5000 });
          await p.waitForTimeout(1800);
        } catch (e) {}
      }
    }
  }
}

(async () => {
  const b = await launch();
  const p = await b.newPage({ viewport: { width: 1400, height: 940 } });
  const report = [];

  const only = process.argv[3] ? process.argv[3].split(',') : null;
  for (const sc of DATA.screens) {
    if (only && !only.includes(sc.id)) continue;
    await p.goto('about:blank');
    await p.goto(DATA.base + sc.href, { waitUntil: 'domcontentloaded', timeout: 120000 });
    await p.waitForTimeout(6500);
    await scrollThrough(p);
    // unlock block-level continue gates without leaving the lesson.
    // The last CONTINUE on a page is the lesson-footer button, which navigates away,
    // so count the safe clicks first, then replay exactly that many after a reload.
    let safeClicks = 0;
    for (let i = 0; i < 12; i++) {
      const btn = await p.$('.continue-btn, .blocks-continue__button');
      if (!btn) break;
      try { await btn.click({ timeout: 2000 }); } catch (e) { break; }
      await p.waitForTimeout(1100);
      await scrollThrough(p);
      if (await p.evaluate(() => location.hash) !== sc.href) break;
      safeClicks++;
    }
    await p.goto('about:blank');
    await p.goto(DATA.base + sc.href, { waitUntil: 'domcontentloaded', timeout: 120000 });
    await p.waitForTimeout(6000);
    await scrollThrough(p);
    for (let i = 0; i < safeClicks; i++) {
      const btn = await p.$('.continue-btn, .blocks-continue__button');
      if (!btn) break;
      try { await btn.click({ timeout: 2000 }); } catch (e) { break; }
      await p.waitForTimeout(1100);
      await scrollThrough(p);
    }
    try { const t = await p.$('button[aria-label="Close navigation menu"]'); if (t) await t.click({ timeout: 2000 }); } catch (e) {}
    await p.waitForTimeout(800);

    await runPrep(p, sc.prep);
    await p.waitForTimeout(600);

    await p.addStyleTag({ content: STYLE });
    await p.evaluate(PAGE_FNS);

    const found = [];
    for (let i = 0; i < sc.items.length; i++) {
      const it = sc.items[i];
      const ok = it.box
        ? await p.evaluate(([sel, nth, n]) => window.__cqa.boxBySelector(sel, nth, n), [it.box.selector, it.box.nth, i + 1])
        : await p.evaluate(([q, n]) => window.__cqa.highlight(q, n), [it.quote, i + 1]);
      found.push(ok);
      if (!ok) console.log(`  !! NOT FOUND [${sc.id} #${i + 1}]: ${sc.items[i].quote.slice(0, 60)}`);
    }

    const order = await p.evaluate(() => window.__cqa.renumber());

    // Grow the viewport so every highlight on this screen fits in one capture,
    // then scroll so the first highlight sits near the top.
    let offs = await p.evaluate(() => window.__cqa.offsets());
    if (offs) {
      const span = offs.max - offs.min;
      const vh = Math.min(3000, Math.max(940, Math.round(span) + 260));
      if (vh !== 940) {
        await p.setViewportSize({ width: 1400, height: vh });
        await p.waitForTimeout(1200);
        await scrollThrough(p);
        await p.waitForTimeout(600);
        offs = await p.evaluate(() => window.__cqa.offsets());
      }
    }
    const VH = p.viewportSize().height;
    if (offs) {
      await p.evaluate(y => { const c = document.querySelector('.page-wrap') || document.scrollingElement; c.scrollTop = Math.max(0, y - 110); }, offs.min);
      await p.waitForTimeout(900);
    }

    const bnd = await p.evaluate(() => window.__cqa.bounds());
    let clip = null;
    if (bnd) {
      const top = Math.max(0, bnd.top - 70);
      const bottom = Math.min(VH, bnd.bottom + 70);
      if (bottom - top > 80) clip = { x: 0, y: top, width: 1400, height: bottom - top };
    }
    const file = `${OUT}/${sc.id.replace(/\./g, '_')}.png`;
    await p.screenshot({ path: file, clip: clip || undefined });
    report.push({ id: sc.id, file, found, order });
    if (p.viewportSize().height !== 940) await p.setViewportSize({ width: 1400, height: 940 });
    console.log(sc.id, '->', file, 'highlighted', found.filter(Boolean).length + '/' + found.length);
  }

  const prev = fs.existsSync(`data/${MOD}_shots.json`) ? JSON.parse(fs.readFileSync(`data/${MOD}_shots.json`,'utf8')) : [];
  const byId = Object.fromEntries(prev.map(r => [r.id, r]));
  report.forEach(r => byId[r.id] = r);
  fs.writeFileSync(`data/${MOD}_shots.json`, JSON.stringify(DATA.screens.map(s => byId[s.id]).filter(Boolean), null, 1));
  await b.close();
})();
