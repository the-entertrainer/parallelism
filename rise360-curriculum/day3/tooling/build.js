/**
 * Day 3 — "Blocks that answer back"
 * Articulate Rise 360, 5-day beginner curriculum. Interactive blocks only:
 * Accordion, Tabs, Labeled Graphic, Process, Scenario, Sorting Activity, Timeline,
 * Flashcard grid, Flashcard stack, Button, Button stack, Storyline.
 * Question types and question banks are assessment — Day 4.
 *
 * Closes with a bonus: what Rise can and cannot run, and a real swipe-card
 * interaction brought in through the Embed block.
 *
 * Same design system as Day 2 — the block card motif, the ember accent sampled
 * from the Rise thumbnails, dark title/section/close, light content.
 */
const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

const IMG = path.join(__dirname, '..', '..', 'day2', 'assets', 'block-screenshots');
const OUT = path.join(__dirname, '..', 'build', 'rise360-day3-interactive.pptx');
const META = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'assets', 'rise_blocks_metadata.json'), 'utf8'));

const C = {
  ink: '12161F', inkSoft: '1C2230', paper: 'FFFFFF', ground: 'F4F6F8',
  ember: 'ED6A2A', emberDim: 'F6C7AC', slate: '5A6675', slateLt: 'A9B3C0',
  line: 'E2E6EB', ok: '2E7D5B', code: '0E1420',
};
const HEAD = 'Cambria', BODY = 'Calibri', MONO = 'Courier New';
const W = 13.333, H = 7.5, M = 0.62, CW = W - M * 2, FOOT_Y = 6.98;

const pres = new PptxGenJS();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'Rise 360 curriculum';
pres.title = 'Rise 360 Day 3 — Blocks that answer back';

// ---------------------------------------------------------------- helpers
const shadow = (o = {}) => ({
  type: 'outer', color: o.color || '9AA6B4',
  blur: o.blur === undefined ? 14 : o.blur,
  offset: o.offset === undefined ? 3 : o.offset,
  angle: 90, opacity: o.opacity === undefined ? 0.32 : o.opacity,
});

function card(slide, x, y, w, h, o = {}) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: o.radius === undefined ? 0.09 : o.radius,
    fill: { color: o.fill || C.paper },
    line: { color: o.line || C.line, width: o.lineW === undefined ? 0.75 : o.lineW },
    shadow: o.flat ? undefined : shadow(o.shadow || {}),
  });
}

function pill(slide, x, y, w, text, o = {}) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h: 0.3, rectRadius: 0.15,
    fill: { color: o.fill || C.ember }, line: { type: 'none' },
  });
  slide.addText(text, {
    x, y, w, h: 0.3, align: 'center', valign: 'middle', margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: 'FFFFFF', charSpacing: 0.6,
  });
}

let slideNo = 0;
function footer(slide, dark) {
  slideNo += 1;
  slide.addText('Rise 360  ·  Day 3  ·  Blocks that answer back', {
    x: M, y: FOOT_Y, w: CW * 0.7, h: 0.28, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 9.5, color: dark ? C.slateLt : C.slate,
  });
  slide.addText(String(slideNo), {
    x: W - M - 1.2, y: FOOT_Y, w: 1.2, h: 0.28, margin: 0, align: 'right', valign: 'middle',
    fontFace: BODY, fontSize: 9.5, color: dark ? C.slateLt : C.slate,
  });
}

function contentSlide(title, standfirst, o = {}) {
  const s = pres.addSlide();
  s.background = { color: o.ground || C.ground };
  s.addText(title, {
    x: M, y: 0.46, w: CW, h: 0.62, margin: 0, valign: 'middle',
    fontFace: HEAD, fontSize: o.titleSize || 30, bold: true, color: C.ink,
  });
  if (standfirst) {
    s.addText(standfirst, {
      x: M, y: 1.1, w: o.sfW || CW * 0.8, h: 0.5, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 14, color: C.slate,
    });
  }
  footer(s, false);
  return s;
}

function darkSlide() {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  return s;
}

function stackGlyph(slide, x, y, o = {}) {
  const w = o.w || 2.1, hh = o.h || 0.52, gap = o.gap || 0.2;
  const tones = o.tones || [C.inkSoft, C.inkSoft, C.ember];
  for (let i = 0; i < 3; i++) {
    slide.addShape(pres.ShapeType.roundRect, {
      x: x + (o.stagger ? i * 0.14 : 0), y: y + i * (hh + gap), w, h: hh, rectRadius: 0.08,
      fill: { color: tones[i] },
      line: { color: i === 2 ? C.ember : '2A3242', width: 1 },
    });
  }
}

function sectionSlide(kicker, title, blurb) {
  const s = darkSlide();
  s.addText(kicker, {
    x: M, y: 2.35, w: CW * 0.6, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 12, bold: true, color: C.ember, charSpacing: 2.2,
  });
  s.addText(title, {
    x: M, y: 2.72, w: CW * 0.62, h: 1.15, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 40, bold: true, color: 'FFFFFF',
  });
  if (blurb) {
    s.addText(blurb, {
      x: M, y: 4.0, w: CW * 0.52, h: 0.9, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 14.5, color: C.slateLt, lineSpacing: 22,
    });
  }
  stackGlyph(s, W - M - 2.5, 2.55, { stagger: true });
  footer(s, true);
  return s;
}

const block = (name) => META.find((b) => b.block_name === name);
const imgOf = (name) => path.join(IMG, path.basename(block(name).local_image_path));

/** Grid of block screenshots that shrinks as one to fit the band. */
function blockGrid(slide, names, o = {}) {
  const top = o.top === undefined ? 1.85 : o.top;
  const bottom = o.bottom === undefined ? 6.1 : o.bottom;
  const cols = o.cols;
  const rows = [];
  for (let i = 0; i < names.length; i += cols) rows.push(names.slice(i, i + cols));

  const gapX0 = 0.24, gapY0 = 0.24, pad0 = 0.12, capH0 = 0.34;
  const cardW0 = (CW - gapX0 * (cols - 1)) / cols;
  const imgW0 = cardW0 - pad0 * 2;
  const imgH0 = imgW0 / (294 / 204);
  const cardH0 = pad0 + imgH0 + 0.05 + capH0 + pad0 * 0.4;
  const totalH0 = rows.length * cardH0 + (rows.length - 1) * gapY0;
  const k = Math.min(1, (bottom - top) / totalH0);

  const cardW = cardW0 * k, cardH = cardH0 * k, gapX = gapX0 * k + (1 - k) * 0.05;
  const gapY = gapY0 * k, pad = pad0 * k, imgW = imgW0 * k, imgH = imgH0 * k, capH = capH0 * k;
  const startY = top + Math.max(0, ((bottom - top) - (rows.length * cardH + (rows.length - 1) * gapY)) / 2);

  rows.forEach((row, r) => {
    const rowW = row.length * cardW + (row.length - 1) * gapX;
    const x0 = M + (CW - rowW) / 2;
    row.forEach((nm, c) => {
      const x = x0 + c * (cardW + gapX), y = startY + r * (cardH + gapY);
      card(slide, x, y, cardW, cardH);
      slide.addImage({ path: imgOf(nm), x: x + pad, y: y + pad, w: imgW, h: imgH });
      slide.addText(nm, {
        x: x + pad * 0.4, y: y + pad + imgH + 0.03 * k, w: cardW - pad * 0.8, h: capH, margin: 0,
        align: 'center', valign: 'middle',
        fontFace: BODY, fontSize: Math.max(8.5, 11 * k), bold: true, color: C.ink,
      });
    });
  });
}

/**
 * A family slide: one or two blocks, each with a screenshot, what it is for,
 * and a concrete example of the job it does.
 */
function jobSlide(o) {
  const s = contentSlide(o.title, o.standfirst);
  const n = o.blocks.length;
  const gap = 0.5;
  const cw = (CW - gap * (n - 1)) / n;
  const top = 1.78, cardH = 4.34;

  o.blocks.forEach((b, i) => {
    const x = M + i * (cw + gap);
    card(s, x, top, cw, cardH);

    const imgW = Math.min(cw - 0.9, 3.1);
    const imgH = imgW / (294 / 204);
    s.addImage({ path: imgOf(b.name), x: x + (cw - imgW) / 2, y: top + 0.2, w: imgW, h: imgH });

    let y = top + 0.2 + imgH + 0.14;
    s.addText(b.name, {
      x: x + 0.34, y, w: cw - 0.68, h: 0.34, margin: 0, valign: 'middle', align: 'center',
      fontFace: HEAD, fontSize: 18, bold: true, color: C.ink,
    });
    y += 0.42;
    s.addText([
      { text: 'Use it when  ', options: { bold: true, color: C.ember } },
      { text: b.use, options: { color: C.slate } },
    ], {
      x: x + 0.34, y, w: cw - 0.68, h: 0.62, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12.5, lineSpacing: 17,
    });
    y += 0.66;
    s.addText([
      { text: 'Example  ', options: { bold: true, color: C.ink } },
      { text: b.eg, options: { color: C.slate } },
    ], {
      x: x + 0.34, y, w: cw - 0.68, h: 0.68, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12.5, lineSpacing: 17,
    });
  });

  card(s, M, 6.3, CW, 0.56, { fill: 'FFF4EE', line: C.emberDim, flat: true });
  s.addShape(pres.ShapeType.ellipse, {
    x: M + 0.22, y: 6.46, w: 0.22, h: 0.22, fill: { color: C.ember }, line: { type: 'none' },
  });
  s.addText(o.note, {
    x: M + 0.56, y: 6.3, w: CW - 0.84, h: 0.56, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12, color: '7A3F1C',
  });
  s.addNotes(o.notes);
  return s;
}

// ================================================================ SLIDES

// ---- 1. Title -------------------------------------------------------------
{
  const s = darkSlide();
  s.addText('DAY 3 OF 5', {
    x: M, y: 1.62, w: 4, h: 0.32, margin: 0,
    fontFace: BODY, fontSize: 12.5, bold: true, color: C.ember, charSpacing: 2.6,
  });
  s.addText('Blocks that\nanswer back', {
    x: M, y: 2.04, w: 7.6, h: 2.0, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 50, bold: true, color: 'FFFFFF', lineSpacing: 56,
  });
  s.addText('The twelve interactive blocks, when each one earns its place — and a bonus interaction Rise cannot build on its own.',
    { x: M, y: 4.18, w: 6.9, h: 0.9, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 15.5, color: C.slateLt, lineSpacing: 24 });

  const gx = 8.55, gy = 2.0, gw = 3.8;
  [['Read', C.inkSoft], ['Read', C.inkSoft], ['Do', C.ember], ['Read', C.inkSoft]].forEach((r, i) => {
    const y = gy + i * 0.86;
    s.addShape(pres.ShapeType.roundRect, {
      x: gx + (i === 2 ? 0.22 : 0), y, w: gw, h: 0.66, rectRadius: 0.08,
      fill: { color: r[1] }, line: { color: i === 2 ? C.ember : '2A3242', width: 1 },
    });
    s.addText(r[0], {
      x: gx + (i === 2 ? 0.22 : 0), y, w: gw, h: 0.66, margin: 0, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 12.5, bold: true, color: i === 2 ? 'FFFFFF' : C.slateLt,
    });
  });
  s.addText('one lean-forward moment per screenful', {
    x: gx, y: 5.55, w: gw, h: 0.3, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 11, color: C.slateLt, italic: true,
  });
  footer(s, true);
  s.addNotes(
    'Day 2 was everything a learner reads. Today is everything a learner operates.\n\n' +
    'Open with the honest framing: interactions are not automatically better. A course made entirely of ' +
    'accordions is as tiring as a course made entirely of paragraphs. What changes today is that you get a ' +
    'second gear — and the skill is knowing when to change up.\n\n' +
    'Twelve blocks, grouped into six jobs. Nobody memorises twelve; they remember the six jobs.\n\n' +
    'The bonus at the end is the part people will remember: a Tinder-style swipe deck, which Rise cannot ' +
    'build natively, brought in through the Embed block. Do not skip it — it is also where the honest limits ' +
    'of the tool get taught.'
  );
}

// ---- 2. From reading to doing --------------------------------------------
{
  const s = contentSlide('From reading to doing',
    'You can already build a readable lesson. Today it starts asking things of the learner.');
  const colW = (CW - 0.9) / 2;

  card(s, M, 1.85, colW, 3.9);
  s.addText('YESTERDAY YOU BUILT', {
    x: M + 0.4, y: 2.12, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  s.addText([
    { text: 'A lesson made of static blocks', options: { bullet: true, breakLine: true } },
    { text: 'Consistent padding and backgrounds', options: { bullet: true, breakLine: true } },
    { text: 'No more than half of it text', options: { bullet: true } },
  ], {
    x: M + 0.4, y: 2.5, w: colW - 0.8, h: 1.6, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 14, color: C.ink, lineSpacing: 21, paraSpaceAfter: 10,
  });
  s.addText('Open it now. Today you will change part of it.', {
    x: M + 0.4, y: 4.9, w: colW - 0.8, h: 0.6, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13, italic: true, color: C.slate,
  });

  card(s, M + colW + 0.9, 1.85, colW, 3.9, { fill: C.ink, line: C.ink });
  s.addText('BY TONIGHT YOU WILL', {
    x: M + colW + 1.3, y: 2.12, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Name the twelve interactive blocks by the job they do', options: { bullet: true, breakLine: true } },
    { text: 'Turn a text-heavy block into an interaction', options: { bullet: true, breakLine: true } },
    { text: 'Say why an interaction earns its place — or drop it', options: { bullet: true, breakLine: true } },
    { text: 'Know what Rise cannot build, and what to do about it', options: { bullet: true } },
  ], {
    x: M + colW + 1.3, y: 2.5, w: colW - 0.8, h: 3.0, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 14, color: 'FFFFFF', lineSpacing: 21, paraSpaceAfter: 10,
  });

  card(s, M, 5.96, CW, 0.6, { fill: 'FFF4EE', line: C.emberDim, flat: true });
  s.addText('Assessment blocks — the four question types and question banks — are tomorrow. Today is interaction, not scoring.', {
    x: M + 0.38, y: 5.96, w: CW - 0.76, h: 0.6, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, color: '7A3F1C',
  });
  s.addNotes(
    'Get them to open yesterday’s lesson before you go any further — today only works on real content.\n\n' +
    'The line that matters is the third one on the right: saying why an interaction earns its place, or dropping it. ' +
    'That is the difference between a designer and someone who clicks every button in the block library.\n\n' +
    'Head off the assessment question now. Somebody will ask about knowledge checks within ten minutes. ' +
    'Point at the orange strip: interaction today, scoring tomorrow.'
  );
}

// ---- 3. Section one -------------------------------------------------------
sectionSlide('SECTION ONE', 'Twelve blocks,\nsix jobs',
  'Learn the job, not the block list. The job tells you which two or three to choose between.')
  .addNotes(
    'Six jobs: reveal, explore, sequence, recall, decide, act. Plus the Storyline block, which is a door to ' +
    'another tool rather than a job of its own.\n\n' +
    'Pace: six slides, roughly three minutes each. Demo the block live as you talk about it — these are all ' +
    'faster to show than to describe.'
  );

// ---- 4. The twelve at a glance -------------------------------------------
{
  const s = contentSlide('The twelve interactive blocks',
    'Grouped by the job they do. Every one of them costs the learner a click, so every one needs a reason.');
  blockGrid(s, [
    'Accordion', 'Tabs', 'Labeled Graphic', 'Process', 'Timeline', 'Sorting Activity',
    'Scenario', 'Flashcard grid', 'Flashcard stack', 'Button', 'Button stack', 'Storyline',
  ], { cols: 6, top: 1.85, bottom: 6.2 });
  s.addText('12 interactive blocks  ·  question types and question banks are Day 4', {
    x: M, y: 6.3, w: CW, h: 0.3, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 12, color: C.slate, italic: true,
  });
  s.addNotes(
    'Thirty seconds. Do not walk the twelve — the next six slides do that.\n\n' +
    'The line to land: every interactive block costs the learner a click. A click is a small tax. ' +
    'You charge it when what is behind the click is worth more than the tax — and not otherwise.\n\n' +
    'If anyone asks why Quote carousel and Gallery carousel are not here: they are filed as static content ' +
    'blocks and we covered them yesterday, even though they do take a click.'
  );
}

// ---- 5. Reveal ------------------------------------------------------------
jobSlide({
  title: 'Job 1 — Reveal',
  standfirst: 'Hide secondary text so the page stays short, and let the learner pull only what they need.',
  blocks: [
    { name: 'Accordion',
      use: 'items are independent and a learner will read some but not all of them.',
      eg: 'Twelve refund-policy FAQs. Nobody reads twelve; everyone reads the two that apply to them.' },
    { name: 'Tabs',
      use: 'items are the same shape and the learner is comparing them.',
      eg: 'Three pricing plans, each with the same four rows — switch in place, compare like for like.' },
  ],
  note: 'Never hide anything a learner must not miss. Warnings, safety steps and deadlines belong in a Statement or Note block, in the open.',
  notes:
    'These two look interchangeable and are not. The test is comparison.\n\n' +
    'Accordion is for independent items — FAQs, a reference list, troubleshooting symptoms. You open one, ' +
    'read it, close it. Nobody compares FAQ 3 with FAQ 9.\n\n' +
    'Tabs are for parallel items you want to hold against each other. Three plans, three regions, three ' +
    'customer types — same shape of content each time. The content switches in place, which is what makes ' +
    'comparison possible.\n\n' +
    'Rise converts one to the other, so the cost of getting it wrong is a dropdown click. Show that.\n\n' +
    'The orange rule is the one to labour. Every cohort has someone who puts the safety warning in an ' +
    'accordion because the page looked long. A click is a filter, and a warning must not be filtered.'
});

// ---- 6. Explore -----------------------------------------------------------
{
  const s = contentSlide('Job 2 — Explore',
    'One image, several things worth naming. Markers beat a list underneath the picture.');

  const imgW = 4.3, imgH = imgW / (294 / 204);
  card(s, M, 1.85, imgW + 0.5, imgH + 0.9);
  s.addImage({ path: imgOf('Labeled Graphic'), x: M + 0.25, y: 2.1, w: imgW, h: imgH });
  s.addText('Labeled Graphic', {
    x: M + 0.25, y: 2.1 + imgH + 0.1, w: imgW, h: 0.4, margin: 0, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 19, bold: true, color: C.ink,
  });

  const rx = M + imgW + 0.5 + 0.5, rw = CW - imgW - 1.0;
  const pts = [
    ['Use it when', 'the parts only make sense in place — on the hardware, on the screen, on the map.'],
    ['Example', 'The card reader with five markers: display, chip slot, contactless zone, power button, USB-C port.'],
    ['Watch for', 'markers crowding on a phone. Rise reflows the image, so keep them apart and test on mobile.'],
    ['Sizing', 'medium by default (1100px); small (760px) and full width are the alternatives.'],
  ];
  let y = 1.88;
  pts.forEach((pt) => {
    s.addText(pt[0], {
      x: rx, y, w: rw, h: 0.28, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 11.5, bold: true, color: C.ember, charSpacing: 0.8,
    });
    s.addText(pt[1], {
      x: rx, y: y + 0.27, w: rw, h: 0.76, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 18,
    });
    y += 1.08;
  });

  card(s, M, 6.3, CW, 0.56, { fill: 'FFF4EE', line: C.emberDim, flat: true });
  s.addShape(pres.ShapeType.ellipse, {
    x: M + 0.22, y: 6.46, w: 0.22, h: 0.22, fill: { color: C.ember }, line: { type: 'none' },
  });
  s.addText('A labeled graphic with two markers is a caption wearing a costume. Below about four, just caption the image.', {
    x: M + 0.56, y: 6.3, w: CW - 0.84, h: 0.56, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12, color: '7A3F1C',
  });
  s.addNotes(
    'The single most persuasive block in Rise for anything physical, and the one that most often gets used ' +
    'for the wrong thing.\n\n' +
    'It earns its place when position carries meaning. Naming the parts of a reader, walking a screen, ' +
    'pointing at a floor plan. If you could rewrite it as a bulleted list with no loss, it should be a list.\n\n' +
    'Demo the mobile crop. Markers that sit comfortably apart on a laptop can collide on a phone, and ' +
    'nobody notices until a learner complains. Preview it on mobile portrait in front of them.\n\n' +
    'The image sizes are a real setting worth knowing: medium at 1100px is the default, with small at 760px ' +
    'and full width available.'
  );
}

// ---- 7. Sequence ----------------------------------------------------------
jobSlide({
  title: 'Job 3 — Sequence',
  standfirst: 'Order matters and you want the learner to feel the steps, one at a time.',
  blocks: [
    { name: 'Process',
      use: 'the steps are a procedure someone will carry out.',
      eg: 'Pairing a card reader: six steps, one screen each, so nobody skims past step four.' },
    { name: 'Timeline',
      use: 'the sequence is history or a roadmap, not something to perform.',
      eg: 'How the product line changed: first reader in 2012, acquisition, rebrand, current model.' },
  ],
  note: 'A numbered list is often the better answer. Reach for these when the steps carry enough detail that showing them all at once overwhelms.',
  notes:
    'Both put one thing on screen at a time. The difference is whether the learner will do it or just know it.\n\n' +
    'Process is a procedure. It paces someone through steps they will repeat later — which is exactly why it ' +
    'beats a numbered list for anything with detail per step.\n\n' +
    'Timeline is chronology. Company history, a rollout plan, what changed and when. Nobody performs a timeline.\n\n' +
    'Be honest about the orange note. A six-item numbered list is faster to build, faster to scan and easier ' +
    'to reread than a Process block, and for simple steps it is the better choice. The block earns its place ' +
    'when each step needs a paragraph and a picture — that is when showing all six at once becomes a wall.'
});

// ---- 8. Recall ------------------------------------------------------------
jobSlide({
  title: 'Job 4 — Recall',
  standfirst: 'Term on one side, meaning on the other. The learner tries before they are told.',
  blocks: [
    { name: 'Flashcard grid',
      use: 'the set is small and the learner should see how much there is.',
      eg: 'Six product names to learn, laid out at once so the scope is obvious.' },
    { name: 'Flashcard stack',
      use: 'the set is longer and you want one card at a time, in order.',
      eg: 'Ten support FAQs worked through as a run, like revision cards.' },
  ],
  note: 'The front must be answerable. A card whose front is a statement teaches nothing — the learner flips it without thinking.',
  notes:
    'Flashcards are the cheapest retrieval practice in the tool, and the most commonly wasted.\n\n' +
    'The waste looks like this: the front says "Refund window" and the back says "30 days". The learner flips ' +
    'it, reads it, learns nothing, because they were never asked anything. Make the front a question they ' +
    'could actually attempt: "How long does a customer have to request a refund?"\n\n' +
    'Grid versus stack is about set size and whether scope is useful information. A grid shows six at once and ' +
    'says "this is all of it". A stack gives you one at a time and is better past about eight.\n\n' +
    'Quick exercise if time allows: take any flashcard from a draft and rewrite the front as a question. ' +
    'The difference is immediate.'
});

// ---- 9. Decide ------------------------------------------------------------
jobSlide({
  title: 'Job 5 — Decide',
  standfirst: 'The learner commits to an answer and finds out what it costs. This is where interaction earns most.',
  blocks: [
    { name: 'Sorting Activity',
      use: 'the skill is telling categories apart.',
      eg: 'Twelve emails dragged into Phishing or Safe — the judgement is the learning.' },
    { name: 'Scenario',
      use: 'the skill is choosing a response, and the response has consequences.',
      eg: 'An angry refund call: three replies, each leading somewhere different.' },
  ],
  note: 'Scenario is the most expensive block in the tool to write well. Budget for the branching, not for the styling.',
  notes:
    'If you only add one interaction to a lesson, make it one of these two. Everything else on the list ' +
    'reorganises content; these two make the learner practise the actual skill.\n\n' +
    'Sorting is for discrimination — the moment where someone has to say "this one, not that one". ' +
    'Do and Don’t piles, compliant versus non-compliant, phishing versus legitimate. The dragging is not the ' +
    'point; the judgement is.\n\n' +
    'Scenario is role-play with consequences. It uses Content Library characters, dialogue, responses and ' +
    'feedback. Warn them properly: this is the block that takes real writing time. The branching has to be ' +
    'plausible, every branch needs feedback that teaches rather than scolds, and a badly written scenario is ' +
    'worse than no scenario because it feels like a trick.\n\n' +
    'Note the difference from a quiz: neither of these is scored, and neither reports to an LMS. They are ' +
    'practice. Scoring is tomorrow.'
});

// ---- 10. Act --------------------------------------------------------------
jobSlide({
  title: 'Job 6 — Act',
  standfirst: 'Send the learner somewhere, or give them a choice of routes through the lesson.',
  blocks: [
    { name: 'Button',
      use: 'there is one clear next action.',
      eg: 'Download the job aid, or open the policy in the knowledge base.' },
    { name: 'Button stack',
      use: 'there are two to four routes and the learner picks their own.',
      eg: 'Choose your region, or your role, and jump to the lesson that fits.' },
  ],
  note: 'Label a button with what happens when it is pressed. “Learn more” tells the learner nothing; “Open the refund policy” tells them everything.',
  notes:
    'The least glamorous pair and the most immediately useful.\n\n' +
    'The whole lesson here is labelling. "Click here", "Learn more" and "Read on" are the three worst button ' +
    'labels in e-learning and they are everywhere. A label should describe the destination or the action, so ' +
    'the learner can decide whether to spend the click.\n\n' +
    'Button stacks are how you branch a course without restricting navigation — offer three routes, let ' +
    'people take theirs. Pair it with the Previous and Next toggles under Theme when you build a branched ' +
    'path, or learners will simply scroll past your carefully designed fork.\n\n' +
    'Mention that a button can open a file or an external link, which makes it the natural partner to the ' +
    'Attachment block from yesterday.'
});

// ---- 11. Storyline --------------------------------------------------------
{
  const s = contentSlide('The advanced one — Storyline block',
    'A door out of Rise. It embeds a published Storyline 360 interaction inside a Rise lesson.');

  const imgW = 3.7, imgH = imgW / (294 / 204);
  card(s, M, 1.9, imgW + 0.5, imgH + 0.85);
  s.addImage({ path: imgOf('Storyline'), x: M + 0.25, y: 2.12, w: imgW, h: imgH });
  s.addText('Storyline', {
    x: M + 0.25, y: 2.12 + imgH + 0.08, w: imgW, h: 0.38, margin: 0, align: 'center', valign: 'middle',
    fontFace: HEAD, fontSize: 18, bold: true, color: C.ink,
  });

  const rx = M + imgW + 1.0, rw = CW - imgW - 1.0;
  card(s, rx, 1.9, rw, 1.72, { fill: C.ink, line: C.ink });
  s.addText('Reach for it when Rise genuinely cannot', {
    x: rx + 0.36, y: 2.08, w: rw - 0.72, h: 0.36, margin: 0,
    fontFace: HEAD, fontSize: 17, bold: true, color: 'FFFFFF',
  });
  s.addText('Free-form drag and drop, a simulated interface, custom logic, anything with its own rules. Build it in Storyline 360, publish it, drop it in here.',
    { x: rx + 0.36, y: 2.5, w: rw - 0.72, h: 0.96, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slateLt, lineSpacing: 19 });

  card(s, rx, 3.78, rw, 1.72);
  s.addText('The costs, stated plainly', {
    x: rx + 0.36, y: 3.96, w: rw - 0.72, h: 0.36, margin: 0,
    fontFace: HEAD, fontSize: 17, bold: true, color: C.ink,
  });
  s.addText('A second tool, a second skill set, and a fixed-size interaction sitting inside a responsive page. It is the one place a Rise lesson stops reflowing.',
    { x: rx + 0.36, y: 4.38, w: rw - 0.72, h: 0.96, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 19 });

  card(s, M, 5.72, CW, 0.98, { fill: C.paper });
  s.addText([
    { text: 'Tracking, for later.  ', options: { bold: true, color: C.ink } },
    { text: 'When you publish to an LMS you can choose to track “Storyline block viewed” as the completion trigger. Worth knowing today; we set it on Day 5.', options: { color: C.slate } },
  ], {
    x: M + 0.38, y: 5.72, w: CW - 0.76, h: 0.98, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 13, lineSpacing: 18,
  });
  s.addNotes(
    'Keep this short unless your group already uses Storyline. Most beginners will not touch it for months.\n\n' +
    'The reason it is on the slide is boundary-setting: it is the answer to "can Rise do X?" when X is ' +
    'genuinely beyond block-based authoring. Knowing the door exists stops people trying to fake a simulation ' +
    'with fourteen accordions.\n\n' +
    'Be straight about the cost. It is a second product with its own learning curve, and the embedded ' +
    'interaction is fixed-size inside an otherwise responsive lesson — on a phone, that is visible.\n\n' +
    'This slide also sets up the bonus. Storyline is one way out of Rise’s limits. The Embed block is the other, ' +
    'and it is the one you can use this afternoon without buying anything.'
  );
}

// ---- 12. Section two ------------------------------------------------------
sectionSlide('SECTION TWO', 'Choosing well',
  'Which interaction for which job — and the four rules that stop a lesson turning into a fairground.')
  .addNotes(
    'Short section, two slides. This is the judgement half of the day, and it is what separates today from ' +
    'a feature tour.'
  );

// ---- 13. Decision table ---------------------------------------------------
{
  const s = contentSlide('Which interaction for which job',
    'Start from what the learner has to do. The block falls out of that.');

  const rows = [
    ['Read some, skip the rest', 'Accordion', 'Independent items, read on demand'],
    ['Compare like with like', 'Tabs', 'Same shape of content, switched in place'],
    ['Find it on the picture', 'Labeled Graphic', 'Position carries the meaning'],
    ['Follow the steps', 'Process', 'A procedure they will perform'],
    ['See what changed when', 'Timeline', 'Chronology, not a task'],
    ['Try to remember it', 'Flashcards', 'Retrieval practice — front must ask something'],
    ['Tell these apart', 'Sorting Activity', 'The judgement is the learning'],
    ['Choose, and live with it', 'Scenario', 'Decisions with consequences'],
    ['Go somewhere next', 'Button / stack', 'Label it with what happens'],
  ];
  const rowH = 0.47, top = 1.86;
  const cA = 3.8, cB = 2.9, cC = CW - cA - cB;

  // header
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: top, w: CW, h: 0.44, rectRadius: 0.06,
    fill: { color: C.ink }, line: { type: 'none' },
  });
  [['The learner has to…', M + 0.3, cA], ['Reach for', M + cA + 0.3, cB], ['Because', M + cA + cB + 0.3, cC - 0.6]]
    .forEach(([t, x, w]) => s.addText(t, {
      x, y: top, w, h: 0.44, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 11.5, bold: true, color: 'FFFFFF', charSpacing: 0.8,
    }));

  rows.forEach((r, i) => {
    const y = top + 0.44 + i * rowH;
    if (i % 2 === 0) {
      s.addShape(pres.ShapeType.rect, {
        x: M, y, w: CW, h: rowH, fill: { color: 'FFFFFF' }, line: { color: C.line, width: 0.5 },
      });
    } else {
      s.addShape(pres.ShapeType.rect, {
        x: M, y, w: CW, h: rowH, fill: { color: 'EDF1F5' }, line: { color: C.line, width: 0.5 },
      });
    }
    s.addText(r[0], {
      x: M + 0.3, y, w: cA, h: rowH, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12.5, color: C.ink,
    });
    s.addText(r[1], {
      x: M + cA + 0.3, y, w: cB, h: rowH, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12.5, bold: true, color: C.ember,
    });
    s.addText(r[2], {
      x: M + cA + cB + 0.3, y, w: cC - 0.6, h: rowH, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12.5, color: C.slate,
    });
  });
  s.addNotes(
    'Hand this out. It is the artefact people keep from today, exactly as the checklist was yesterday.\n\n' +
    'Do not read it aloud. Instead run it backwards: give the room three of your own lesson topics and ask ' +
    'which row applies. Arguments are good here — "is this comparison or is it reference?" is the exact ' +
    'question you want them asking on their own later.\n\n' +
    'The left column is deliberately written as learner behaviour, not as content type. That is the habit: ' +
    'start from what the learner has to do.'
  );
}

// ---- 14. Four rules -------------------------------------------------------
{
  const s = contentSlide('Four rules for interactions',
    'Each one prevents a specific way that interactive lessons go wrong.');
  const rules = [
    ['Lead in', 'Every interaction needs a line above it saying what it is and what to do with it.', 'Prevents: a bare accordion nobody opens'],
    ['One per screenful', 'Space them out. Two interactions back to back read as a quiz nobody agreed to take.', 'Prevents: interaction fatigue'],
    ['Never hide the critical', 'Warnings, deadlines, safety steps stay in the open, in a Statement or Note.', 'Prevents: the click that gets skipped'],
    ['Check it on a phone', 'Markers crowd, drag targets shrink, tab labels wrap. Preview mobile portrait before you call it done.', 'Prevents: the defect you ship'],
  ];
  const gap = 0.3, w = (CW - gap * 3) / 4;
  rules.forEach((r, i) => {
    const x = M + i * (w + gap);
    card(s, x, 1.9, w, 3.72);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.28, y: 2.16, w: 0.44, h: 0.44, rectRadius: 0.1,
      fill: { color: C.ember }, line: { type: 'none' },
    });
    s.addText(String(i + 1), {
      x: x + 0.28, y: 2.16, w: 0.44, h: 0.44, margin: 0, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 15, bold: true, color: 'FFFFFF',
    });
    s.addText(r[0], {
      x: x + 0.28, y: 2.74, w: w - 0.56, h: 0.74, margin: 0, valign: 'top',
      fontFace: HEAD, fontSize: 18, bold: true, color: C.ink,
    });
    s.addText(r[1], {
      x: x + 0.28, y: 3.52, w: w - 0.56, h: 1.34, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12.5, color: C.slate, lineSpacing: 18,
    });
    s.addText(r[2], {
      x: x + 0.28, y: 4.94, w: w - 0.56, h: 0.56, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 11.5, bold: true, color: C.ok, lineSpacing: 15,
    });
  });
  card(s, M, 5.72, CW, 0.72, { fill: C.paper });
  s.addText([
    { text: 'The test for any interaction:  ', options: { bold: true, color: C.ink } },
    { text: 'if you removed it and pasted the content back in as text, would the lesson be worse? If not, remove it.', options: { color: C.slate } },
  ], {
    x: M + 0.38, y: 5.72, w: CW - 0.76, h: 0.72, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 13, lineSpacing: 18,
  });
  s.addNotes(
    'Four rules, four failure modes. Name the failure mode as you go — people remember the symptom, not the rule.\n\n' +
    'Rule 1 is the one from yesterday, restated: an interaction with no lead line gets scrolled past. ' +
    'Rise’s own guidance asks for a line that connects the interaction to the content around it.\n\n' +
    'Rule 4 is the one they will skip. Make them do it during the lab, not after.\n\n' +
    'The bottom line is the whole day compressed into one sentence. If you are short of time and have to cut ' +
    'a slide, cut something else and keep this.'
  );
}

// ---- 15. Section three ----------------------------------------------------
sectionSlide('BONUS', 'Past the block list',
  'What Rise cannot build on its own — and the twenty minutes it takes to get around that.')
  .addNotes(
    'This is the part of the day people talk about afterwards. It is also the part where the tool’s limits get ' +
    'taught honestly, which builds more trust than pretending Rise does everything.\n\n' +
    'Three slides: what Rise can and cannot run, what we built, and how it works. Keep it demo-led.'
  );

// ---- 16. What Rise can and cannot run ------------------------------------
{
  const s = contentSlide('Code in Rise: what actually runs',
    'The most common misunderstanding on this course, and it is worth ten minutes.');

  const cols = [
    { t: 'Code snippet', sub: 'Static block', body: 'Displays code as monospaced text. It is a picture of code. Nothing executes — not one line.', tone: 'no' },
    { t: 'Embed', sub: 'Static block', body: 'Pulls in a page you host at a public HTTPS address. Your HTML, CSS and JavaScript run inside it.', tone: 'yes' },
    { t: 'Storyline', sub: 'Interactive block', body: 'Runs a published Storyline 360 interaction. Powerful, but a second tool and a fixed size.', tone: 'ok' },
  ];
  const gap = 0.4, w = (CW - gap * 2) / 3;
  cols.forEach((c, i) => {
    const x = M + i * (w + gap);
    const dark = c.tone === 'yes';
    card(s, x, 1.9, w, 2.7, dark ? { fill: C.ink, line: C.ink } : {});
    s.addText(c.sub, {
      x: x + 0.34, y: 2.1, w: w - 0.68, h: 0.28, margin: 0,
      fontFace: BODY, fontSize: 10.5, bold: true, charSpacing: 1.4,
      color: dark ? C.ember : C.slate,
    });
    s.addText(c.t, {
      x: x + 0.34, y: 2.42, w: w - 0.68, h: 0.44, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 22, bold: true, color: dark ? 'FFFFFF' : C.ink,
    });
    s.addText(c.body, {
      x: x + 0.34, y: 2.94, w: w - 0.68, h: 1.4, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: dark ? C.slateLt : C.slate, lineSpacing: 19,
    });
  });

  card(s, M, 4.86, CW, 1.6, { fill: C.paper });
  s.addText('So the recipe is three steps', {
    x: M + 0.4, y: 5.04, w: CW - 0.8, h: 0.36, margin: 0,
    fontFace: HEAD, fontSize: 18, bold: true, color: C.ink,
  });
  const steps = [
    ['1', 'Build one self-contained HTML file'],
    ['2', 'Host it at a public HTTPS address'],
    ['3', 'Paste the address into an Embed block'],
  ];
  const sw = (CW - 0.8) / 3;
  steps.forEach((st, i) => {
    const x = M + 0.4 + i * sw;
    s.addShape(pres.ShapeType.ellipse, {
      x, y: 5.52, w: 0.32, h: 0.32, fill: { color: C.ember }, line: { type: 'none' },
    });
    s.addText(st[0], {
      x, y: 5.52, w: 0.32, h: 0.32, margin: 0, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 12, bold: true, color: 'FFFFFF',
    });
    s.addText(st[1], {
      x: x + 0.44, y: 5.46, w: sw - 0.6, h: 0.72, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 18,
    });
  });
  s.addNotes(
    'Ask the question first: "If I paste JavaScript into a Code snippet block, does it run?" Let them answer. ' +
    'Most rooms say yes. It does not. The Code snippet block displays code as text — it is a picture of code.\n\n' +
    'The Embed block is the one that runs things, because it loads a page you host and that page executes ' +
    'normally. This is not a hack; it is what the block is for.\n\n' +
    'Two warnings to give with it. The embedded page depends on your hosting staying up, and corporate ' +
    'networks sometimes block third-party frames — so never put anything essential behind an embed alone. ' +
    'And it will not be tracked by your LMS; only quizzes report scores.\n\n' +
    'Hosting does not have to be a project. Any static host works. If your organisation already has an ' +
    'intranet or a documentation site, that is usually the path of least resistance.'
  );
}

// ---- 17. The swipe deck ---------------------------------------------------
{
  const s = contentSlide('Bonus build: a swipe deck',
    'Card sorting with a phone gesture. Rise has no block for this — so we made one and embedded it.');

  // three states, left to right
  const shots = [
    ['../assets/swipe-screenshots/swipe_rest.png', 'The stack', 'Three cards deep, so there is visibly more to come.'],
    ['../assets/swipe-screenshots/swipe_drag.png', 'Mid-swipe', 'The card tilts as it travels and the verdict fades in past halfway.'],
    ['../assets/swipe-screenshots/swipe_done.png', 'The result', 'Score, and feedback on the ones they missed.'],
  ];
  const gap = 0.42, w = (CW - gap * 2) / 3;
  shots.forEach((sh, i) => {
    const x = M + i * (w + gap);
    card(s, x, 1.82, w, 4.05);
    const src = path.join(__dirname, sh[0]);
    const dim = pngSize(src);
    const maxW = w - 0.6, maxH = 2.55;
    let iw = maxW, ih = iw * dim.h / dim.w;
    if (ih > maxH) { ih = maxH; iw = ih * dim.w / dim.h; }
    s.addImage({ path: src, x: x + (w - iw) / 2, y: 2.02, w: iw, h: ih });
    s.addText(sh[1], {
      x: x + 0.3, y: 2.02 + maxH + 0.12, w: w - 0.6, h: 0.34, margin: 0, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 17, bold: true, color: C.ink,
    });
    s.addText(sh[2], {
      x: x + 0.3, y: 2.02 + maxH + 0.5, w: w - 0.6, h: 0.72, margin: 0, align: 'center', valign: 'top',
      fontFace: BODY, fontSize: 12, color: C.slate, lineSpacing: 17,
    });
  });

  card(s, M, 6.06, CW, 0.78, { fill: C.ink, line: C.ink });
  s.addText([
    { text: 'Why it works as learning, not decoration:  ', options: { bold: true, color: C.ember } },
    { text: 'every card forces a judgement before the answer appears — the same job as a Sorting Activity, in a gesture people already know.', options: { color: 'FFFFFF' } },
  ], {
    x: M + 0.38, y: 6.06, w: CW - 0.76, h: 0.78, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, lineSpacing: 18,
  });
  s.addNotes(
    'Run it live, on a phone if you can get it on the screen — the gesture is the point and it does not land ' +
    'from a screenshot.\n\n' +
    'The content of the deck is deliberately self-referential: each card is a scenario and the learner swipes ' +
    'right for "an interactive block earns its place here" or left for "a static block does the job". ' +
    'So the bonus doubles as revision of the whole day.\n\n' +
    'Then ask the harder question: is this better than a Sorting Activity? Honest answer — usually not, for a ' +
    'desktop audience. It is better when your learners are on phones, because swiping is native there and ' +
    'dragging into a bucket is fiddly. Say that. It stops the room concluding that custom is always better.\n\n' +
    'The file ships with this deck. They can change the six cards to their own content without touching the ' +
    'JavaScript — the cards are an array at the top.'
  );
}

// ---- 18. How it works (four ideas) --------------------------------------
{
  const s = contentSlide('How the swipe works',
    'One hundred and sixty-five lines, all of it on the next six slides. Four ideas carry the whole thing.');
  const ideas = [
    ['Pointer events', 'One code path covers mouse, touch and pen. You do not write three versions, and setPointerCapture keeps the card following the finger even when it leaves the element.'],
    ['Rotation from travel', 'The card leans by the distance it has moved, divided by eighteen. That single expression is most of what makes it feel like a physical card.'],
    ['One threshold', 'Past 96 pixels it commits and flies off; short of it, the card snaps back. There is no other state to track.'],
    ['Buttons share the path', 'The two buttons and the arrow keys call the same commit function as a drag, so the whole thing works without a pointer at all.'],
  ];
  const gap = 0.3, w = (CW - gap * 3) / 4;
  ideas.forEach((it, i) => {
    const x = M + i * (w + gap);
    card(s, x, 1.9, w, 3.66);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.28, y: 2.16, w: 0.44, h: 0.44, rectRadius: 0.1,
      fill: { color: C.ember }, line: { type: 'none' },
    });
    s.addText(String(i + 1), {
      x: x + 0.28, y: 2.16, w: 0.44, h: 0.44, margin: 0, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 15, bold: true, color: 'FFFFFF',
    });
    s.addText(it[0], {
      x: x + 0.28, y: 2.74, w: w - 0.56, h: 0.72, margin: 0, valign: 'top',
      fontFace: HEAD, fontSize: 17, bold: true, color: C.ink,
    });
    s.addText(it[1], {
      x: x + 0.28, y: 3.5, w: w - 0.56, h: 1.9, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12, color: C.slate, lineSpacing: 17,
    });
  });
  card(s, M, 5.76, CW, 0.68, { fill: 'FFF4EE', line: C.emberDim, flat: true });
  s.addText('Accessibility is not optional here: a drag-only interaction excludes keyboard and screen-reader users. The buttons and arrow keys are the fix, and they cost four lines.', {
    x: M + 0.38, y: 5.76, w: CW - 0.76, h: 0.68, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, color: '7A3F1C',
  });
  s.addNotes(
    'Do not teach JavaScript. Teach the shape of the thing, so a non-coder can brief a developer and a ' +
    'semi-coder can edit the cards.\n\n' +
    'Idea 2 is the one that surprises people. The physicality comes from a single division — the card leans ' +
    'by how far it has been thrown. Take that line out and it becomes a sliding box.\n\n' +
    'Idea 3 is why the code is short. There is no drag state machine, no velocity tracking, no gesture ' +
    'library. One number decides everything.\n\n' +
    'The orange line is the one to insist on. A swipe-only interaction locks out anyone who cannot drag. ' +
    'In this build the buttons run through exactly the same commit path, so there is no second behaviour ' +
    'to maintain and no way for the two to drift apart.\n\n' +
    'The next six slides are the complete file. Tell them not to copy it off the screen — the file ships ' +
    'with the deck.'
  );
}

// ---- 19-24. The complete file, straight from disk -------------------------
{
  const SRC = fs.readFileSync(path.join(__dirname, '..', 'bonus', 'swipe-deck.html'), 'utf8').split('\n');
  const chunks = [
    [1, 29,   'Document head and the styling'],
    [30, 57,  'The rest of the styling, and the markup'],
    [58, 87,  'The six cards, and the state'],
    [88, 117, 'Drawing the stack, and picking up a card'],
    [118, 144, 'Committing the swipe'],
    [145, 166, 'Buttons, keys and the result screen'],
  ];
  const LINE_H = 0.1567;                 // 11.3pt line spacing in inches
  chunks.forEach((ch, i) => {
    const [a, b, sub] = ch;
    const s = contentSlide(`swipe-deck.html  —  ${i + 1} of ${chunks.length}`, sub, { titleSize: 26 });
    const seg = SRC.slice(a - 1, b);
    const cardH = Math.min(5.06, 0.36 + seg.length * LINE_H);   // the card fits its own code
    card(s, M, 1.66, CW, cardH, { fill: C.code, line: '2A3242' });
    s.addText(seg.join('\n'), {
      x: M + 0.3, y: 1.84, w: CW - 0.6, h: cardH - 0.36, margin: 0, valign: 'top',
      fontFace: MONO, fontSize: 9.5, color: 'D7E1EC', lineSpacing: 11.3,
    });
    s.addText(`lines ${a}–${Math.min(b, SRC.length)} of ${SRC.length}`, {
      x: W - M - 2.4, y: 6.76, w: 2.4, h: 0.24, margin: 0, align: 'right', valign: 'middle',
      fontFace: BODY, fontSize: 9.5, color: C.slate,
    });
    s.addNotes(
      'This is the file as it ships, not a paraphrase — the deck is generated from it, so the two cannot ' +
      'drift apart.\n\n' +
      (i === 0
        ? 'Nothing here is unusual. It is one HTML file with a style block and a script block, which is exactly ' +
          'what the Embed block needs: one address that serves one self-contained page.'
        : i === 2
        ? 'This is the only part most people will ever edit. Six objects, each with the scenario text, the ' +
          'answer, and the feedback line. Swap those for your own subject and you have a new activity.'
        : i === 3
        ? 'render() draws at most three cards so the stack has visible depth, and attaches the drag handlers ' +
          'only to the top one. Everything below it is decoration until its turn comes.'
        : i === 4
        ? 'commit() is where the answer is scored, the card is thrown off screen, and the feedback is announced ' +
          'to screen readers through the live region. Note that it does not care whether the swipe came from ' +
          'a finger or the keyboard.'
        : i === 5
        ? 'The last piece: the buttons and the arrow keys, then the result screen. show() flips between the ' +
          'deck and the result rather than rebuilding either.'
        : 'The styling is ordinary CSS. The two rules worth pointing at are touch-action:none, which stops the ' +
          'browser scrolling the page while you drag a card, and the prefers-reduced-motion query at the end.') +
      '\n\nDo not read the code aloud. Put it on screen, point at the two or three lines that matter, and move on.'
    );
  });
}

// ---- 19. Lab --------------------------------------------------------------
{
  const s = contentSlide('Your build: earn one interaction',
    'Back into yesterday’s lesson. One change, made for a reason you can say out loud.');

  const colW = (CW - 0.5) / 2;
  card(s, M, 1.9, colW, 3.6, { fill: C.ink, line: C.ink });
  s.addText('THE BRIEF', {
    x: M + 0.4, y: 2.12, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Find your most text-heavy block', options: { bullet: true, breakLine: true } },
    { text: 'Turn it into the interaction its job calls for', options: { bullet: true, breakLine: true } },
    { text: 'Write the lead line above it', options: { bullet: true, breakLine: true } },
    { text: 'Add one decide block — Sorting or Scenario', options: { bullet: true, breakLine: true } },
    { text: 'Preview on mobile portrait before you stop', options: { bullet: true } },
  ], {
    x: M + 0.4, y: 2.5, w: colW - 0.8, h: 2.8, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: 'FFFFFF', lineSpacing: 20, paraSpaceAfter: 8,
  });

  const rx = M + colW + 0.5;
  card(s, rx, 1.9, colW, 1.7);
  s.addText('Say it out loud', {
    x: rx + 0.4, y: 2.1, w: colW - 0.8, h: 0.36, margin: 0,
    fontFace: HEAD, fontSize: 18, bold: true, color: C.ink,
  });
  s.addText('For each interaction you added, finish this sentence for your neighbour: “This is a ___ because the learner has to ___.” If you cannot, take it out.',
    { x: rx + 0.4, y: 2.5, w: colW - 0.8, h: 1.0, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 19 });

  card(s, rx, 3.8, colW, 1.7);
  s.addText('If you finish early', {
    x: rx + 0.4, y: 4.0, w: colW - 0.8, h: 0.36, margin: 0,
    fontFace: HEAD, fontSize: 18, bold: true, color: C.ink,
  });
  s.addText('Open the swipe deck file, replace the six cards with scenarios from your own subject, and open it in a browser. No build step, no account.',
    { x: rx + 0.4, y: 4.4, w: colW - 0.8, h: 1.0, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 19 });

  card(s, M, 5.72, CW, 0.72, { fill: 'FFF4EE', line: C.emberDim, flat: true });
  s.addText('One interaction added well beats four added quickly. If a block does not survive the sentence test, that is a result, not a failure.', {
    x: M + 0.38, y: 5.72, w: CW - 0.76, h: 0.72, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, color: '7A3F1C',
  });
  s.addNotes(
    'Forty-five minutes is enough. Walk the room.\n\n' +
    'The sentence test in the top-right is the assessment for the whole day, and it is verbal on purpose. ' +
    '"This is an accordion because the learner has to read some and skip the rest." Anyone who cannot finish ' +
    'the sentence has added decoration, and taking it out is the right answer — say so, warmly, before they ' +
    'start, so removing something does not feel like failing.\n\n' +
    'Insist on the mobile preview. It is the rule they skip and the one that catches real defects.\n\n' +
    'The early-finisher task is genuinely doable by a non-coder: the cards are a plain array of text at the ' +
    'top of the file, and it opens straight from disk in any browser.'
  );
}

// ---- 20. Close ------------------------------------------------------------
{
  const s = darkSlide();
  s.addText('DAY 3 COMPLETE', {
    x: M, y: 2.35, w: 6.4, h: 0.32, margin: 0,
    fontFace: BODY, fontSize: 12.5, bold: true, color: C.ember, charSpacing: 2.6,
  });
  s.addText('Twelve blocks.\nOne good reason each.', {
    x: M, y: 2.76, w: 7.6, h: 1.7, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 38, bold: true, color: 'FFFFFF', lineSpacing: 46,
  });
  s.addText('Interaction is not the goal. Getting the learner to do the thinking is — and now you have twelve ways to ask.',
    { x: M, y: 4.6, w: 6.9, h: 0.9, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 14.5, color: C.slateLt, lineSpacing: 22 });
  stackGlyph(s, W - M - 3.4, 2.5, { w: 3.4, h: 0.62, gap: 0.22, stagger: true });
  s.addText('Day 4  ›  Assess, theme, navigate', {
    x: W - M - 3.4, y: 5.06, w: 3.4, h: 0.3, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 12, color: C.slateLt, italic: true,
  });
  footer(s, true);
  s.addNotes(
    'Close on the contract from slide 2, then two housekeeping items.\n\n' +
    'One: tomorrow is assessment, theming and navigation — the four question types, knowledge checks versus ' +
    'scored quizzes, brand colours and fonts, and free versus restricted navigation. Their lesson now has ' +
    'something worth assessing, which is why it is in this order.\n\n' +
    'Two: send them the swipe deck file. The people who will actually use it are the ones who open it that ' +
    'evening, and they will not if they have to ask.\n\n' +
    'If the room struggled today, it was almost certainly with Scenario. That is normal — it is a writing ' +
    'problem, not a Rise problem. Offer to look at branching drafts tomorrow morning.'
  );
}

// PNG header reader used for the swipe screenshots (they are not 294x204)
function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}

pres.writeFile({ fileName: OUT }).then(() => {
  console.log('wrote', OUT, '| slides:', slideNo);
});
