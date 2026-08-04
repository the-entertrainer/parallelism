/**
 * Day 2 — "Blocks that carry your content"
 * Articulate Rise 360, 5-day beginner curriculum. Static (non-advanced) blocks only:
 * Text, Statement, Quote, List, Image, Gallery, Multimedia, Chart, Divider.
 * Interactive and assessment blocks are Day 3 and Day 4.
 *
 * 13.333 x 7.5in. Motif: the block card — a white rounded card with a soft shadow,
 * the same object the learner stacks in a Rise lesson. Every content slide is built
 * out of them, so the deck is assembled the way the tool is.
 */
const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

const IMG = path.join(__dirname, '..', 'assets', 'block-screenshots');
const OUT = path.join(__dirname, '..', 'build', 'rise360-day2-blocks.pptx');
const META = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'assets', 'rise_blocks_metadata.json'), 'utf8'));

// ---------------------------------------------------------------- palette
// Sampled from the Rise block thumbnails themselves: the product's own ember accent
// on paper-white cards. The deck and the screenshots inside it share one palette.
const C = {
  ink: '12161F',        // dark ground for title / section / close
  inkSoft: '1C2230',    // raised panels on dark
  paper: 'FFFFFF',
  ground: 'F4F6F8',     // light slide ground
  ember: 'ED6A2A',      // accent — matches the orange inside every Rise thumbnail
  emberDim: 'F6C7AC',
  slate: '5A6675',      // secondary text on light
  slateLt: 'A9B3C0',    // secondary text on dark
  line: 'E2E6EB',
  ok: '2E7D5B',
};

const HEAD = 'Cambria';   // safe-list serif, renders true-to-width in QA
const BODY = 'Calibri';   // safe-list sans

const W = 13.333, H = 7.5;
const M = 0.62;                 // slide margin
const CW = W - M * 2;           // 12.093 content width
const FOOT_Y = 6.98;

const pres = new PptxGenJS();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'Rise 360 curriculum';
pres.title = 'Rise 360 Day 2 — Blocks that carry your content';

// ---------------------------------------------------------------- helpers
const shadow = (o = {}) => ({
  type: 'outer',
  color: o.color || '9AA6B4',
  blur: o.blur === undefined ? 14 : o.blur,
  offset: o.offset === undefined ? 3 : o.offset,
  angle: 90,
  opacity: o.opacity === undefined ? 0.32 : o.opacity,
});

/** The motif: a white block card. Everything on a light slide sits in one. */
function card(slide, x, y, w, h, o = {}) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h,
    rectRadius: o.radius === undefined ? 0.09 : o.radius,
    fill: { color: o.fill || C.paper },
    line: { color: o.line || C.line, width: o.lineW === undefined ? 0.75 : o.lineW },
    shadow: o.flat ? undefined : shadow(o.shadow || {}),
  });
}

/** Small ember pill used for counts and labels. */
function pill(slide, x, y, w, text, o = {}) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h: 0.3, rectRadius: 0.15,
    fill: { color: o.fill || C.ember }, line: { type: 'none' },
  });
  slide.addText(text, {
    x, y, w, h: 0.3, align: 'center', valign: 'middle', margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: o.color || 'FFFFFF',
    charSpacing: 0.6,
  });
}

let slideNo = 0;
const TOTAL_PLACEHOLDER = { n: 0 };

function footer(slide, dark) {
  slideNo += 1;
  const n = slideNo;
  slide.addText('Rise 360  ·  Day 2  ·  Blocks that carry your content', {
    x: M, y: FOOT_Y, w: CW * 0.7, h: 0.28, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 9.5, color: dark ? C.slateLt : C.slate,
  });
  slide.addText(String(n), {
    x: W - M - 1.2, y: FOOT_Y, w: 1.2, h: 0.28, margin: 0, align: 'right', valign: 'middle',
    fontFace: BODY, fontSize: 9.5, color: dark ? C.slateLt : C.slate,
  });
  return n;
}

/** Light content slide with a title and optional standfirst. */
function contentSlide(title, standfirst, o = {}) {
  const s = pres.addSlide();
  s.background = { color: o.ground || C.ground };
  s.addText(title, {
    x: M, y: 0.46, w: CW, h: 0.62, margin: 0, valign: 'middle',
    fontFace: HEAD, fontSize: o.titleSize || 30, bold: true, color: C.ink,
  });
  if (standfirst) {
    s.addText(standfirst, {
      x: M, y: 1.1, w: o.sfW || CW * 0.82, h: 0.42, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 14, color: C.slate, italic: !!o.sfItalic,
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

/** The stack glyph — three offset block cards. Used on the dark section slides. */
function stackGlyph(slide, x, y, o = {}) {
  const w = o.w || 2.1, hh = o.h || 0.52, gap = o.gap || 0.2;
  const tones = o.tones || [C.inkSoft, C.inkSoft, C.ember];
  for (let i = 0; i < 3; i++) {
    slide.addShape(pres.ShapeType.roundRect, {
      x: x + (o.stagger ? i * 0.14 : 0), y: y + i * (hh + gap), w, h: hh,
      rectRadius: 0.08,
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

/** Grid of block screenshots, each in a card with its official name beneath.
 *  Cards are sized from the column count, then scaled down as one if the rows
 *  would not fit the vertical space — so no grid ever runs into the note or footer. */
function blockGrid(slide, blocks, o = {}) {
  const top = o.top === undefined ? 1.82 : o.top;
  const bottom = o.bottom === undefined ? 6.02 : o.bottom;
  const cols = o.cols;
  const rows = [];
  for (let i = 0; i < blocks.length; i += cols) rows.push(blocks.slice(i, i + cols));

  const gapX0 = o.gapX === undefined ? 0.26 : o.gapX;
  const gapY0 = o.gapY === undefined ? 0.24 : o.gapY;
  const pad0 = 0.13;
  const capH0 = o.blurbs ? 0.28 : 0.34;
  const blurbH0 = o.blurbs ? 0.66 : 0;

  const cardW0 = (CW - gapX0 * (cols - 1)) / cols;
  const imgW0 = cardW0 - pad0 * 2;
  const imgH0 = imgW0 / (294 / 204);          // every asset is 294 x 204
  const cardH0 = pad0 + imgH0 + 0.06 + capH0 + blurbH0 + pad0 * 0.4;

  const totalH0 = rows.length * cardH0 + (rows.length - 1) * gapY0;
  const availH = bottom - top;
  const k = Math.min(1, availH / totalH0);    // shrink to fit, never enlarge past the column width

  const cardW = cardW0 * k, cardH = cardH0 * k;
  const gapX = gapX0 * k + (1 - k) * 0.06, gapY = gapY0 * k;
  const pad = pad0 * k, imgW = imgW0 * k, imgH = imgH0 * k;
  const capH = capH0 * k, blurbH = blurbH0 * k;

  const totalH = rows.length * cardH + (rows.length - 1) * gapY;
  const startY = top + Math.max(0, (availH - totalH) / 2);

  rows.forEach((row, r) => {
    const rowW = row.length * cardW + (row.length - 1) * gapX;
    const x0 = M + (CW - rowW) / 2;           // centre a short final row
    row.forEach((b, c) => {
      const x = x0 + c * (cardW + gapX);
      const y = startY + r * (cardH + gapY);
      card(slide, x, y, cardW, cardH);
      slide.addImage({
        path: path.join(IMG, path.basename(b.local_image_path)),
        x: x + pad, y: y + pad, w: imgW, h: imgH,
      });
      slide.addText(b.block_name, {
        x: x + pad * 0.4, y: y + pad + imgH + 0.04 * k, w: cardW - pad * 0.8, h: capH, margin: 0,
        align: 'center', valign: 'middle',
        fontFace: BODY, fontSize: Math.max(9, 11.5 * k), bold: true, color: C.ink,
      });
      if (o.blurbs && o.blurbs[b.block_name]) {
        slide.addText(o.blurbs[b.block_name], {
          x: x + pad, y: y + pad + imgH + 0.04 * k + capH, w: cardW - pad * 2, h: blurbH, margin: 0,
          align: 'center', valign: 'top',
          fontFace: BODY, fontSize: Math.max(8.5, 10.5 * k), color: C.slate, lineSpacing: 14 * k,
        });
      }
    });
  });
}

const byCategory = (name) => META.filter((b) => b.category === name);
const catDesc = (name) => byCategory(name)[0].category_description;

/** A family slide: official purpose line, the screenshots, and a coaching note. */
function familySlide(category, opts) {
  const blocks = byCategory(category);
  const s = contentSlide(opts.title || category, null);
  // official purpose, quoted
  s.addText(opts.purpose, {
    x: M, y: 1.10, w: CW * 0.78, h: 0.56, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: C.slate, italic: true,
  });
  pill(s, W - M - 1.35, 1.12, 1.35, `${blocks.length} BLOCKS`);

  const oneRow = blocks.length <= opts.cols;
  blockGrid(s, blocks, {
    cols: opts.cols,
    top: oneRow ? 1.90 : 1.82,
    bottom: oneRow ? 5.20 : 6.02,
    blurbs: opts.blurbs,
  });

  const ny = oneRow ? 5.50 : 6.16;
  card(s, M, ny, CW, 0.62, { fill: 'FFF4EE', line: C.emberDim, flat: true, radius: 0.08 });
  s.addShape(pres.ShapeType.ellipse, {
    x: M + 0.22, y: ny + 0.19, w: 0.24, h: 0.24,
    fill: { color: C.ember }, line: { type: 'none' },
  });
  s.addText(opts.note, {
    x: M + 0.58, y: ny, w: CW - 0.86, h: 0.62, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, color: '7A3F1C',
  });
  s.addNotes(opts.notes);
  return s;
}

// ================================================================ SLIDES

// ---- 1. Title -------------------------------------------------------------
{
  const s = darkSlide();
  s.addText('DAY 2 OF 5', {
    x: M, y: 1.62, w: 4, h: 0.32, margin: 0,
    fontFace: BODY, fontSize: 12.5, bold: true, color: C.ember, charSpacing: 2.6,
  });
  s.addText('Blocks that carry\nyour content', {
    x: M, y: 2.04, w: 7.6, h: 2.0, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 50, bold: true, color: 'FFFFFF', lineSpacing: 56,
  });
  s.addText(
    'The nine static block families in Articulate Rise 360 — and the layout rules that keep a lesson readable.',
    { x: M, y: 4.18, w: 6.9, h: 0.9, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 15.5, color: C.slateLt, lineSpacing: 24 });

  // motif: a stack of cards, the shape of a Rise lesson
  const gx = 8.55, gy = 1.75, gw = 3.8;
  const heights = [0.58, 1.05, 0.58, 0.82, 0.44];
  const fills = [C.inkSoft, C.inkSoft, C.ember, C.inkSoft, C.inkSoft];
  let yy = gy;
  heights.forEach((hh, i) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: gx, y: yy, w: gw, h: hh, rectRadius: 0.08,
      fill: { color: fills[i] },
      line: { color: i === 2 ? C.ember : '2A3242', width: 1 },
    });
    yy += hh + 0.16;
  });
  s.addText('one lesson  =  a stack of blocks', {
    x: gx, y: yy + 0.06, w: gw, h: 0.3, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 11, color: C.slateLt, italic: true,
  });
  footer(s, true);
  s.addNotes(
    'Welcome to Day 2. Yesterday was the mental model — Course > Sections > Lessons > Blocks — and finding your ' +
    'way around. Nobody has built a lesson yet, so today starts by creating one, and then we live inside it.\n\n' +
    'Framing to say out loud: Rise gives you roughly 42 static blocks. Nobody memorises 42 things. ' +
    'What you memorise is nine families and what each family is FOR. Once you know the family, you pick the ' +
    'member by eye from the block library.\n\n' +
    'Today is deliberately not about interactivity — accordions, tabs, labeled graphics, scenarios and the rest are Day 3. ' +
    'If someone asks about them, park it on the whiteboard and point at tomorrow.\n\n' +
    'Deliverable by the end of the day: one visually varied lesson that obeys the anti-clutter rules.'
  );
}

// ---- 2. Where Day 2 sits --------------------------------------------------
{
  const s = contentSlide('Where today sits', 'Five days, one build workflow. Today is the second step: filling a lesson.');
  const days = [
    { n: '1', t: 'Orientation', d: 'The web-page model, the hierarchy, the dashboard, Course vs Microlearning' },
    { n: '2', t: 'Static blocks & layout', d: 'Create a lesson, the nine content families, padding, backgrounds, anti-clutter' },
    { n: '3', t: 'Interactive blocks', d: 'Accordion, tabs, labeled graphic, process, sorting, scenario' },
    { n: '4', t: 'Assessment & theme', d: 'Knowledge checks, quizzes, colours, fonts, free vs restricted navigation' },
    { n: '5', t: 'Preview & publish', d: 'Multi-device preview, Quick Share, Reach 360, LMS export' },
  ];
  const gap = 0.24;
  const w = (CW - gap * 4) / 5;
  days.forEach((d, i) => {
    const x = M + i * (w + gap);
    const on = i === 1;
    card(s, x, 2.05, w, 3.7, { fill: on ? C.ink : C.paper, line: on ? C.ink : C.line });
    s.addShape(pres.ShapeType.ellipse, {
      x: x + 0.28, y: 2.34, w: 0.5, h: 0.5,
      fill: { color: on ? C.ember : C.ground }, line: { color: on ? C.ember : C.line, width: 1 },
    });
    s.addText(d.n, {
      x: x + 0.28, y: 2.34, w: 0.5, h: 0.5, margin: 0, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 17, bold: true, color: on ? 'FFFFFF' : C.slate,
    });
    s.addText(d.t, {
      x: x + 0.26, y: 3.0, w: w - 0.52, h: 0.84, margin: 0, valign: 'top',
      fontFace: HEAD, fontSize: 15, bold: true, color: on ? 'FFFFFF' : C.ink,
    });
    s.addText(d.d, {
      x: x + 0.26, y: 3.92, w: w - 0.52, h: 1.6, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 11.5, color: on ? C.slateLt : C.slate, lineSpacing: 16,
    });
  });
  s.addText('You are here', {
    x: M + (w + gap) * 1, y: 5.9, w, h: 0.3, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 11.5, bold: true, color: C.ember,
  });
  s.addNotes(
    'Thirty seconds on this slide, no more. The point is orientation, not detail.\n\n' +
    'Say: the five days follow the order you actually build in. You outlined yesterday. Today you fill a lesson with ' +
    'content blocks. Tomorrow you add the interactions. Day 4 you assess and brand it. Day 5 you ship it.\n\n' +
    'If your group is publishing to a corporate LMS, flag now that Day 5 will spend extra time on SCORM version ' +
    'and tracking choices.'
  );
}

// ---- 3. Recap -> today ----------------------------------------------------
{
  const s = contentSlide('From yesterday to today', 'Yesterday was the map. Today you build on it — starting from nothing.');
  const colW = (CW - 0.9) / 2;
  // left
  card(s, M, 1.9, colW, 4.05);
  s.addText('WHAT YOU ALREADY KNOW', {
    x: M + 0.4, y: 2.18, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Why Rise is responsive, not slide-based', options: { bullet: true, breakLine: true } },
    { text: 'The hierarchy: Course › Sections › Lessons › Blocks', options: { bullet: true, breakLine: true } },
    { text: 'Your way around the dashboard', options: { bullet: true, breakLine: true } },
    { text: 'The difference between a Course and Microlearning', options: { bullet: true } },
  ], {
    x: M + 0.4, y: 2.58, w: colW - 0.8, h: 3.1, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 14, color: C.ink, lineSpacing: 21, paraSpaceAfter: 10,
  });
  // right
  card(s, M + colW + 0.9, 1.9, colW, 4.05, { fill: C.ink, line: C.ink });
  s.addText('BY TONIGHT YOU WILL', {
    x: M + colW + 1.3, y: 2.18, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Create a lesson and open the block editor', options: { bullet: true, breakLine: true } },
    { text: 'Name the nine static families and what each is for', options: { bullet: true, breakLine: true } },
    { text: 'Add, reorder, swap and delete blocks fluently', options: { bullet: true, breakLine: true } },
    { text: 'Set padding and backgrounds on purpose, not by accident', options: { bullet: true, breakLine: true } },
    { text: 'Build one lesson that passes the anti-clutter rules', options: { bullet: true } },
  ], {
    x: M + colW + 1.3, y: 2.58, w: colW - 0.8, h: 3.1, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 14, color: 'FFFFFF', lineSpacing: 21, paraSpaceAfter: 10,
  });
  card(s, M, 6.12, CW, 0.62, { fill: 'FFF4EE', line: C.emberDim, flat: true });
  s.addShape(pres.ShapeType.ellipse, {
    x: M + 0.22, y: 6.31, w: 0.24, h: 0.24,
    fill: { color: C.ember }, line: { type: 'none' },
  });
  s.addText('Nobody has a lesson yet — that is the first thing we do today, before a single block goes in.', {
    x: M + 0.58, y: 6.12, w: CW - 0.86, h: 0.62, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, color: '7A3F1C',
  });
  s.addNotes(
    'Say the orange line out loud. Day 1 ended with the concepts and the tour, not with anything built, ' +
    'so start from the assumption that everyone is looking at an empty account. That is fine — creating the ' +
    'lesson takes about five minutes and it is the next thing we do.\n\n' +
    'Quick verbal check before you move on — ask two people to answer, do not lecture:\n' +
    '  1. Why does Rise not have slide dimensions?\n' +
    '  2. What is the difference between a Section and a Lesson?\n\n' +
    'If more than a couple of people are shaky on the hierarchy, spend five minutes re-drawing ' +
    'Course > Sections > Lessons > Blocks on the whiteboard before starting. Everything today assumes it.\n\n' +
    'The right-hand column is the contract for the day. Come back to it at the close.'
  );
}

// ---- 4. Section 1 ---------------------------------------------------------
sectionSlide('SECTION ONE', 'The stack',
  'What a lesson is, how to create one from scratch, and the three ways to put a block into it.')
  .addNotes(
    'Three ideas in this section: a lesson is a vertical stack, here is how you create one, and here are the ' +
    'three routes for adding to it.\n\n' +
    'This is the section that changed because Day 1 did not end with a built lesson. Everyone leaves this ' +
    'section with an empty lesson open in the block editor — do not move on to the families until they do, ' +
    'or the rest of the day is a spectator sport.'
  );

// ---- 5. A lesson is a stack ----------------------------------------------
{
  const s = contentSlide('A lesson is a vertical stack',
    'No slides, no canvas, no fixed size. Blocks sit one on top of the next and the learner scrolls.');

  // The phone-ish frame showing a stacked lesson
  const fx = M, fy = 1.95, fw = 3.5, fh = 4.35;
  card(s, fx, fy, fw, fh, { fill: C.paper });
  const rows = [
    { h: 0.5, c: C.ink, label: 'Heading' },
    { h: 0.72, c: C.ground, label: 'Paragraph' },
    { h: 0.92, c: C.emberDim, label: 'Image' },
    { h: 0.5, c: C.ground, label: 'Statement' },
    { h: 0.62, c: C.ground, label: 'List' },
  ];
  let ry = fy + 0.22;
  rows.forEach((r) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: fx + 0.26, y: ry, w: fw - 0.52, h: r.h, rectRadius: 0.06,
      fill: { color: r.c }, line: { color: C.line, width: 0.75 },
    });
    s.addText(r.label, {
      x: fx + 0.26, y: ry, w: fw - 0.52, h: r.h, margin: 0, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 11.5, bold: true, color: r.c === C.ink ? 'FFFFFF' : C.slate,
    });
    ry += r.h + 0.12;
  });
  s.addText('scroll  ↓', {
    x: fx, y: fy + fh - 0.36, w: fw, h: 0.3, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 11, color: C.slate, italic: true,
  });

  // Right: what follows from it
  const rx = M + fw + 0.62, rw = CW - fw - 0.62;
  const pts = [
    ['One column, top to bottom', 'Blocks never sit side by side. Two-column looks come from inside a block — Two column text, Image & text, a grid gallery — never from placing two blocks next to each other.'],
    ['Order is the only layout you control', 'Move a block with the up and down arrows; duplicate and delete appear on hover. There is no free positioning — to move something, you move its block.'],
    ['Every block reflows on its own', 'You do not design a phone version. Rise re-stacks the same blocks narrower. Preview it — do not assume it.'],
  ];
  let py = 2.0;
  pts.forEach((p, i) => {
    s.addShape(pres.ShapeType.ellipse, {
      x: rx, y: py + 0.04, w: 0.34, h: 0.34,
      fill: { color: C.ink }, line: { type: 'none' },
    });
    s.addText(String(i + 1), {
      x: rx, y: py + 0.04, w: 0.34, h: 0.34, margin: 0, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 12, bold: true, color: 'FFFFFF',
    });
    s.addText(p[0], {
      x: rx + 0.5, y: py, w: rw - 0.5, h: 0.34, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 17, bold: true, color: C.ink,
    });
    s.addText(p[1], {
      x: rx + 0.5, y: py + 0.38, w: rw - 0.5, h: 0.95, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 18,
    });
    py += 1.48;
  });
  s.addNotes(
    'This is the slide that unteaches PowerPoint. Say it plainly: there is no canvas and there are no slide dimensions. ' +
    'You are writing a web page.\n\n' +
    'The trap to name early is the two-column one. Somebody always tries to place two blocks side by side to get columns. ' +
    'You cannot. Side-by-side comes from a block that is itself two-column: Two column text, Image & text, or a grid gallery. ' +
    'Naming it now saves you a support question every single cohort.\n\n' +
    'Second trap: people assume a desktop layout is the layout. It is one of five previews. We check all five on Day 5, ' +
    'but get them previewing on mobile today, while their lessons are still small.'
  );
}

// ---- 5b. Create the lesson ------------------------------------------------
{
  const s = contentSlide('First: create the lesson you will fill today',
    'Five minutes, once. Everything after this happens inside the block editor.');

  const steps = [
    ['Open your course', 'From the dashboard, open the course you want this lesson in. No course yet? Create › Course › Blank.'],
    ['Name the lesson', 'Type the lesson title straight into the outline. Titles run to 100 characters and save themselves.'],
    ['Choose Lesson', 'Press Enter and Rise offers Lesson or Quiz. Take Lesson. Shift + Enter would give you a Section instead.'],
    ['Add Content', 'Click Add Content on the new lesson, then Blank Lesson. Templates are there; start blank today.'],
    ['You are in the editor', 'The block editor opens on an empty stack. This is where the rest of the day happens.'],
  ];
  const gap = 0.26;
  const w = (CW - gap * 4) / 5;
  steps.forEach((st, i) => {
    const x = M + i * (w + gap);
    card(s, x, 2.0, w, 3.04);
    s.addShape(pres.ShapeType.ellipse, {
      x: x + 0.26, y: 2.26, w: 0.46, h: 0.46,
      fill: { color: C.ember }, line: { type: 'none' },
    });
    s.addText(String(i + 1), {
      x: x + 0.26, y: 2.26, w: 0.46, h: 0.46, margin: 0, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 14, bold: true, color: 'FFFFFF',
    });
    s.addText(st[0], {
      x: x + 0.26, y: 2.86, w: w - 0.52, h: 0.78, margin: 0, valign: 'top',
      fontFace: HEAD, fontSize: 15.5, bold: true, color: C.ink,
    });
    s.addText(st[1], {
      x: x + 0.26, y: 3.68, w: w - 0.52, h: 1.24, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 11.5, color: C.slate, lineSpacing: 16,
    });
  });

  const colW = (CW - 0.4) / 2;
  card(s, M, 5.3, colW, 1.28, { fill: C.ink, line: C.ink });
  s.addText('Enter vs Shift + Enter', {
    x: M + 0.36, y: 5.46, w: colW - 0.72, h: 0.32, margin: 0,
    fontFace: HEAD, fontSize: 15.5, bold: true, color: 'FFFFFF',
  });
  s.addText('Enter makes a Lesson — a place blocks live. Shift + Enter makes a Section — a text-only label that groups lessons and holds nothing. Everyone does this by accident once.',
    { x: M + 0.36, y: 5.8, w: colW - 0.72, h: 0.7, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12, color: C.slateLt, lineSpacing: 16 });

  card(s, M + colW + 0.4, 5.3, colW, 1.28);
  s.addText('Course, not Microlearning', {
    x: M + colW + 0.76, y: 5.46, w: colW - 0.72, h: 0.32, margin: 0,
    fontFace: HEAD, fontSize: 15.5, bold: true, color: C.ink,
  });
  s.addText('Microlearning is a single lesson and does not support quizzes. For a course you will grow — and assess on Day 4 — start from Course.',
    { x: M + colW + 0.76, y: 5.8, w: colW - 0.72, h: 0.7, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12, color: C.slate, lineSpacing: 16 });

  s.addNotes(
    'Do this together, screens up, everybody following. It is the one moment today where you want the whole ' +
    'room in lockstep, because everything else depends on having an empty lesson open.\n\n' +
    'Walk the five steps on your own screen first, then give them five minutes and walk the room. ' +
    'Do not move on until every person has the block editor open. Ask them to say "in" when they are there.\n\n' +
    'The two boxes at the bottom are the two traps.\n\n' +
    'Shift + Enter is the classic one: someone presses it, gets a Section, and cannot work out why there is ' +
    'no Add Content button. A Section is a label, not a container — it holds no content at all. Show the ' +
    'difference on screen rather than describing it.\n\n' +
    'The Course versus Microlearning point matters because Microlearning cannot hold a quiz, and we add ' +
    'a quiz on Day 4. Anyone who starts a Microlearning today will have to rebuild.\n\n' +
    'Tell them to pick a real topic from their own work now, not later — the lesson they create here is the ' +
    'one they will build in this afternoon\u2019s lab and carry into Day 3.'
  );
}

// ---- 6. Three ways to add a block ----------------------------------------
{
  const s = contentSlide('Three ways to add a block',
    'All three insert the same thing. Which one you reach for is a habit worth choosing deliberately.');
  const ways = [
    { k: 'A', t: 'The blocks shortcut bar', d: 'The row of common block types under the block you are on. Fastest for the block you already know you want.',
      when: 'Use when: drafting fast' },
    { k: 'B', t: 'All Blocks', d: 'Opens the full block library sidebar, grouped by family. Browse when you are not sure what exists or what it looks like.',
      when: 'Use when: choosing' },
    { k: 'C', t: 'The insert icon  ( + )', d: 'Appears between two existing blocks. The only route that puts a block in the middle of a finished stack.',
      when: 'Use when: revising' },
  ];
  const gap = 0.4, w = (CW - gap * 2) / 3;
  ways.forEach((wy, i) => {
    const x = M + i * (w + gap);
    card(s, x, 2.0, w, 3.5);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.34, y: 2.32, w: 0.46, h: 0.46, rectRadius: 0.1,
      fill: { color: C.ember }, line: { type: 'none' },
    });
    s.addText(wy.k, {
      x: x + 0.34, y: 2.32, w: 0.46, h: 0.46, margin: 0, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 16, bold: true, color: 'FFFFFF',
    });
    s.addText(wy.t, {
      x: x + 0.34, y: 2.94, w: w - 0.68, h: 0.6, margin: 0, valign: 'top',
      fontFace: HEAD, fontSize: 18, bold: true, color: C.ink,
    });
    s.addText(wy.d, {
      x: x + 0.34, y: 3.58, w: w - 0.68, h: 1.35, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 19,
    });
    s.addText(wy.when, {
      x: x + 0.34, y: 4.96, w: w - 0.68, h: 0.32, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12, bold: true, color: C.ember,
    });
  });
  card(s, M, 5.78, CW, 0.82, { fill: C.paper });
  s.addText([
    { text: 'Swap without retyping.  ', options: { bold: true, color: C.ink } },
    { text: 'Every block has a dropdown in its upper-left corner. It swaps the block for another member of the same family and keeps your content. Try Quote A, then Quote C, then Quote on image — same words, three looks.', options: { color: C.slate } },
  ], {
    x: M + 0.34, y: 5.78, w: CW - 0.68, h: 0.82, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 13, lineSpacing: 18,
  });
  s.addNotes(
    'Demo all three live, in this order, on a throwaway lesson. Ninety seconds.\n\n' +
    'The swap dropdown at the bottom is the single most useful thing on this slide and the most missed. ' +
    'Show it deliberately: type a quote once, then cycle it through Quote A, C and On image. People relax visibly ' +
    'when they see that choosing the wrong variant costs nothing.\n\n' +
    'Also show, quickly: up/down arrows to reorder, and the duplicate and delete icons that appear on hover. ' +
    'Then stop — resist demonstrating any interactive block today.'
  );
}

// ---- 7. Families at a glance ---------------------------------------------
{
  const s = contentSlide('The nine static families',
    'Learn the families, not the 42 members. The family tells you where to look; the library shows you the rest.');
  const fams = [
    ['Text', 7, 'Carry the argument'],
    ['Statement', 5, 'Make one line land'],
    ['Quote', 6, 'Give a voice a face'],
    ['List', 3, 'Sequence or itemise'],
    ['Image', 5, 'Show, don’t tell'],
    ['Gallery', 4, 'Several images at once'],
    ['Multimedia', 5, 'Audio, video, files, embeds'],
    ['Chart', 3, 'Make a number visual'],
    ['Divider', 4, 'Pace and separate'],
  ];
  const cols = 5, gap = 0.24;
  const w = (CW - gap * (cols - 1)) / cols;
  const hgt = 1.78;
  fams.forEach((f, i) => {
    const r = Math.floor(i / cols), c = i % cols;
    const rowLen = r === 0 ? 5 : 4;
    const rowW = rowLen * w + (rowLen - 1) * gap;
    const x0 = M + (CW - rowW) / 2;
    const x = x0 + c * (w + gap);
    const y = 2.02 + r * (hgt + 0.32);
    card(s, x, y, w, hgt);
    s.addText(f[0], {
      x: x + 0.24, y: y + 0.26, w: w - 0.48, h: 0.42, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 19, bold: true, color: C.ink,
    });
    s.addText(f[2], {
      x: x + 0.24, y: y + 0.70, w: w - 0.48, h: 0.52, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12, color: C.slate, lineSpacing: 16,
    });
    s.addText(`${f[1]} blocks`, {
      x: x + 0.24, y: y + hgt - 0.46, w: w - 0.48, h: 0.3, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 11.5, bold: true, color: C.ember,
    });
  });
  s.addText('42 static blocks in total  ·  interactive and assessment blocks come on Days 3 and 4', {
    x: M, y: 6.06, w: CW, h: 0.34, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 12.5, color: C.slate, italic: true,
  });
  s.addNotes(
    'Have the group read the nine names aloud once. It sounds silly; it works. The families are the vocabulary ' +
    'for the rest of the day and for every conversation you will have about a lesson afterwards.\n\n' +
    'The right-hand number is how many members each family has. Nobody needs to memorise members. ' +
    'The skill is: "I need this line to land" -> Statement family -> open the library -> pick by eye.\n\n' +
    'Note the curriculum outline lists Chart separately from the static families; we teach it here because ' +
    'it behaves exactly like the other content blocks — you fill it in, it is not interactive.\n\n' +
    'Next nine slides are one per family, same shape each time: what it is for, what the members look like, ' +
    'and one thing to watch.'
  );
}

// ---- 8. Section 2 ---------------------------------------------------------
sectionSlide('SECTION TWO', 'The nine families', 'One slide each: the official purpose, every member, and the mistake to avoid.')
  .addNotes(
    'Pace warning. Nine slides at roughly four minutes each is about 35 minutes. Do not let Text blocks eat fifteen ' +
    'of them — it is the family everyone is most comfortable with and least needs.\n\n' +
    'Spend your time on the families they will under-use: Statement, Divider, Chart.'
  );

// ---- 9-17. Family slides --------------------------------------------------
familySlide('Text Blocks', {
  title: 'Text blocks',
  purpose: '“Tell your story with these text block layouts.”  Seven layouts — the workhorses of any lesson, and the ones to ration.',
  cols: 4,
  note: 'Keep paragraphs to two to four sentences. If a paragraph needs more, it usually wants to be a list, a table, or a second block.',
  notes:
    'Four things worth saying here.\n\n' +
    '1. Paragraph vs Paragraph with heading vs with subheading — the difference is only the heading level above the text. ' +
    'Use heading blocks to create scannable structure; people skim before they read.\n\n' +
    '2. Two column is where two-column layout actually comes from. Point back to the stack slide.\n\n' +
    '3. Table is plain and unstyled by design. It is for genuine tabular data. It is not a layout tool, and it does not ' +
    'scroll gracefully on a phone — if a table has more than about four columns, rethink it.\n\n' +
    '4. This is the family the anti-clutter rule is aimed at. Everyone over-reaches for Paragraph. ' +
    'The discipline is: write the paragraph, then ask what it would look like as a list, a statement, or an image with a caption.',
});

familySlide('Statement Blocks', {
  blurbs: {
    'Statement A': 'Narrow, centred, hairline rules — quiet and formal',
    'Statement B': 'Centred under a short accent rule — a wider measure',
    'Statement C': 'Full-width tinted panel — the loudest of the four',
    'Statement D': 'Left-aligned and compact — emphasis mid-flow',
    'Note': 'A boxed aside for a caveat or a heads-up',
  },

  title: 'Statement blocks',
  purpose: '“Make important information stand out… four uniquely-styled statement blocks and a useful note block.”',
  cols: 5,
  note: 'A statement earns its weight by being rare. Two or three per lesson read as emphasis; one every screen reads as shouting.',
  notes:
    'The most under-used family in a beginner’s first course, and the fastest way to make a text-heavy lesson breathe.\n\n' +
    'Statement A through D are four visual treatments of the same thing: one short, important line. Pick by look — ' +
    'there is no functional difference, so do not let people agonise.\n\n' +
    'Note is the odd one out and worth calling out explicitly: it is styled as an aside — a caveat, a heads-up, ' +
    'a "check with your manager first". Teach it as the block for the sentence that would otherwise sit in brackets.\n\n' +
    'Exercise if you have a minute: take a dense paragraph from someone’s draft and pull one sentence out of it into ' +
    'a Statement. The paragraph gets shorter and the point gets louder. That is the whole lesson of this family.',
});

familySlide('Quote Blocks', {
  title: 'Quote blocks',
  purpose: '“Highlight quotes in your story using these eye-catching quote blocks.”  Four styles, plus quote on image and a carousel.',
  cols: 3,
  note: 'Attribute every quote. An unattributed quote in a training module reads as decoration and learners discount it.',
  notes:
    'Quotes do a specific job: they bring a human voice into the lesson. A line from a real customer, a real technician, ' +
    'a real manager carries weight that the same sentence in a paragraph does not.\n\n' +
    'Quote on image is the one to demo — it is the highest-impact block in this family and it is where cover-image ' +
    'cropping bites. Remind them: the image crops differently on every device, so the subject needs to survive edge cropping.\n\n' +
    'Quote carousel holds several quotes in one block. Useful for a "what our learners said" moment. ' +
    'Note that it does require a click to advance, so it sits at the border of static and interactive — but it is filed ' +
    'as a quote block, and we teach it here.\n\n' +
    'Warn against invented quotes. If the quote is fabricated for illustration, label it as an example.',
});

familySlide('List Blocks', {
  blurbs: {
    'Numbered list': 'Numbered markers — use when order matters',
    'Checkbox list': 'Square markers — things the learner will tick off',
    'Bulleted list': 'Plain dots — an unordered set',
  },

  title: 'List blocks',
  purpose: '“Make your point with lists. Choose from three styles.”  The choice of marker is a meaning, not a decoration.',
  cols: 3,
  note: 'Numbered means order matters. Checkbox means the learner will do these. Bulleted means neither. Pick on meaning, not on looks.',
  notes:
    'Small family, easy to teach, and the one where you can plant a genuinely useful habit.\n\n' +
    'The marker carries meaning. Numbered list = a sequence; doing step three before step two breaks something. ' +
    'Checkbox list = things the learner will actually tick off or verify. Bulleted = an unordered set.\n\n' +
    'Beginners pick these by appearance and end up numbering things that have no order, which quietly tells the learner ' +
    'a lie about the task.\n\n' +
    'Second habit worth planting: parallel phrasing. Every item in a list should start the same way — all verbs, or all ' +
    'noun phrases. A list that mixes "Check the reader" with "Battery life" is harder to scan than it looks. ' +
    'This comes up again in the editorial review on Day 5.',
});

familySlide('Image Blocks', {
  blurbs: {
    'Image centered': 'Sits inside the content column — the safe default',
    'Image full width': 'Breaks the margins edge to edge — use sparingly',
    'Image & text': 'Picture on one side, words on the other',
    'Text on image': 'Words laid over the picture — watch contrast',
    'Banner': 'Wide and short — good as a section opener',
  },

  title: 'Image blocks',
  purpose: '“Make pictures pop with these stunning image blocks.”  Five ways to place a single image.',
  cols: 5,
  note: 'Images crop differently on every device. Keep the subject central and preview on mobile before you commit.',
  notes:
    'The five members differ in how much room the image gets and whether text sits with it.\n\n' +
    'Image centered — sits in the content column with margins either side. The safe default.\n' +
    'Image full width — breaks the margins, edge to edge. Use it for a moment, not for every picture.\n' +
    'Image & text — the honest two-column layout: picture one side, words the other.\n' +
    'Text on image — words laid over the picture. Contrast is on you; a busy photo makes the text unreadable. ' +
    'Rise gives you an overlay for exactly this reason.\n' +
    'Banner — a wide, short strip. Good as a section opener inside a long lesson.\n\n' +
    'The cropping point matters more than any of the above. Rise reflows, so the same image is cropped differently on ' +
    'desktop, tablet and phone. Anything critical near an edge will be lost. Say it once here and prove it in preview.',
});

familySlide('Gallery Blocks', {
  blurbs: {
    'Carousel': 'One at a time, learner advances — good for sequence',
    'Two column grid': 'Two images side by side',
    'Three column grid': 'Three peers — comparison at a glance',
    'Four column grid': 'Four peers — keep each one visually simple',
  },

  title: 'Gallery blocks',
  purpose: '“Showcase multiple images in gallery blocks, including carousels and grids.”',
  cols: 4,
  note: 'Every gallery image needs a caption that says something. “Image 1” is not a caption.',
  notes:
    'Use a gallery when several images are peers — four product shots, four workspace setups, four before-and-afters. ' +
    'If one image matters more than the others, it wants an Image block, not a slot in a grid.\n\n' +
    'Carousel shows one at a time and needs a click; the grids show everything at once. ' +
    'Grids are better when you want comparison, the carousel is better when you want sequence.\n\n' +
    'Choose the column count for the number of images, not the other way round. A three-column grid with four images ' +
    'leaves an orphan on the second row and it looks like a mistake.\n\n' +
    'Captions: this is where galleries fall down. A grid of four uncaptioned photos teaches nothing. ' +
    'Every caption should tell the learner what to notice.',
});

familySlide('Multimedia Blocks', {
  blurbs: {
    'Audio': 'Upload or link a clip — introduce it first',
    'Video': 'Upload or link — say what to watch for',
    'Embed': 'Pulls in an external page or player',
    'Attachment': 'A downloadable job aid, checklist or form',
    'Code snippet': 'Monospaced — displays code, never runs it',
  },

  title: 'Multimedia blocks',
  purpose: '“Create media-rich content… audio clips, videos, website embeds, file attachments, and code snippets (text only).”',
  cols: 5,
  note: 'Always write a lead line above a video or audio block that says what the learner should watch or listen for.',
  notes:
    'Five very different things filed under one family.\n\n' +
    'Audio and Video — upload or link. Always introduce them. A bare video with no lead-in gets skipped; ' +
    '"Watch the two minutes below and note where the technician confirms the code" gets watched.\n\n' +
    'Embed — pulls in an external page or player. Two warnings: it depends on that external service staying up, ' +
    'and it may be blocked on a corporate network. Never put anything essential behind an embed alone.\n\n' +
    'Attachment — a downloadable file: a job aid, a checklist, a form. This is the block people forget exists and then ' +
    'email the PDF separately.\n\n' +
    'Code snippet — text only, monospaced. It displays code; it does not run it.\n\n' +
    'Accessibility, briefly: captions and transcripts are your responsibility, not the block’s.',
});

familySlide('Chart Blocks', {
  blurbs: {
    'Bar': 'Compares things',
    'Line': 'Shows change over time',
    'Pie': 'Parts of one whole — only a few slices',
  },

  title: 'Chart blocks',
  purpose: '“Create gorgeous data visualizations… simple to build bar, line, and pie charts your learners will love.”',
  cols: 3,
  note: 'Say the finding in a heading above the chart. A chart on its own makes the learner guess what you wanted them to see.',
  notes:
    'Three chart types, built by typing values into the block — no spreadsheet, no image editor.\n\n' +
    'Match the type to the question. Bar compares things. Line shows change over time. ' +
    'Pie shows parts of one whole — and only works with a handful of slices; past five or six it becomes unreadable.\n\n' +
    'The habit to teach is the headline. Do not write "Q3 results" above a chart. Write the finding: ' +
    '"Escalations fell by a third after the new script." Then the chart is evidence for a claim rather than a puzzle.\n\n' +
    'Keep the data small. These blocks are for making one point visual, not for reporting.\n\n' +
    'A note on the curriculum: charts are sometimes listed apart from the other static families. ' +
    'They belong here — you fill one in exactly like any other content block.',
});

familySlide('Divider Blocks', {
  blurbs: {
    'Continue': 'Hides what follows until the learner clicks',
    'Divider': 'A plain rule between sections',
    'Numbered Divider': 'A rule with a step number — good for procedures',
    'Spacer': 'Pure white space when padding is not enough',
  },

  title: 'Divider blocks',
  purpose: '“Organize block lessons into logical sections… and use continue blocks to encourage learners to master content before moving on.”',
  cols: 4,
  note: 'Continue is the one block here that changes behaviour: it stops the lesson until the learner clicks. Use it to gate, not to decorate.',
  notes:
    'The family that looks like nothing and does the most for readability.\n\n' +
    'Divider — a rule between sections. Numbered divider — the same, with a step number, which is excellent for ' +
    'procedures. Spacer — pure white space, for when two blocks are crowding each other and padding is not enough.\n\n' +
    'Continue is different in kind, so separate it clearly. It hides everything below it until the learner clicks. ' +
    'That is a real behaviour change: it forces a pause. Use it at a genuine decision or comprehension point ' +
    '— before an assessment, or before "now that you have read the policy". Sprinkling Continue blocks through a lesson ' +
    'just makes people click.\n\n' +
    'Tie back to pacing: a long lesson with no dividers reads as an undifferentiated wall. ' +
    'Dividers are how you show structure without adding words.',
});

// ---- 18. Section 3 --------------------------------------------------------
sectionSlide('SECTION THREE', 'Layout & polish', 'Padding, backgrounds, animation — the three settings that decide whether a good lesson looks good.')
  .addNotes(
    'The families were about choosing. This section is about finishing.\n\n' +
    'These three settings are where beginner lessons visibly differ from experienced ones — not in the content, ' +
    'in the consistency.'
  );

// ---- 19. Padding ----------------------------------------------------------
{
  const s = contentSlide('Padding: the space above and below a block',
    'Open the design icon in the block’s upper-right corner. Padding is the first thing in the panel and the first thing people get wrong.');

  const colW = (CW - 0.5) / 2;
  // left — the numbers
  card(s, M, 1.95, colW, 3.45);
  s.addText('THE NUMBERS', {
    x: M + 0.38, y: 2.2, w: colW - 0.76, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  const nums = [
    ['30px', 'the default — Articulate calls this “plenty of breathing room”'],
    ['200px', 'the current maximum in the Format menu'],
    ['0px', 'no padding at all, for blocks meant to touch'],
    ['S / M / L', 'presets; exact pixel values still live under “more”'],
  ];
  let ny = 2.6;
  nums.forEach((n) => {
    s.addText(n[0], {
      x: M + 0.38, y: ny, w: 1.5, h: 0.34, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 17, bold: true, color: C.ember,
    });
    s.addText(n[1], {
      x: M + 1.95, y: ny, w: colW - 2.33, h: 0.44, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12.5, color: C.slate, lineSpacing: 16,
    });
    ny += 0.56;
  });

  // right — the rule
  const rx = M + colW + 0.5;
  card(s, rx, 1.95, colW, 3.45, { fill: C.ink, line: C.ink });
  s.addText('THE RULE THAT MATTERS', {
    x: rx + 0.38, y: 2.2, w: colW - 0.76, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText('Consistency beats any particular value.', {
    x: rx + 0.38, y: 2.56, w: colW - 0.76, h: 0.92, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 21, bold: true, color: 'FFFFFF',
  });
  s.addText(
    'A course where every block sits at 30px looks deliberate. A course where padding wanders between 20, 45 and 60 ' +
    'looks broken, even though no single block is wrong.\n\nPick a value. Use it everywhere. Change it only when you ' +
    'mean something by it.',
    { x: rx + 0.38, y: 3.5, w: colW - 0.76, h: 1.8, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slateLt, lineSpacing: 19 });

  card(s, M, 5.58, CW, 1.16, { fill: 'FFF4EE', line: C.emberDim, flat: true });
  s.addText([
    { text: 'Documentation lag, worth knowing.  ', options: { bold: true, color: '7A3F1C' } },
    { text: 'Older Articulate help pages describe a 0–100px range from the legacy sidebar Settings tab. The current Format menu goes to 200px. When a help article and the interface disagree, the interface is right — Rise ships faster than its documentation. Teach people to check the live UI rather than trust a screenshot in an article.', options: { color: '7A3F1C' } },
  ], {
    x: M + 0.38, y: 5.58, w: CW - 0.76, h: 1.16, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, lineSpacing: 18,
  });
  s.addNotes(
    'Demo, do not describe. Put three blocks on a lesson, set the middle one to 100px, and let them see the lesson ' +
    'go lopsided. Then set all three to 30 and watch it settle.\n\n' +
    'The consistency point is the one to labour. Beginners tune padding block by block, chasing a look on one screen, ' +
    'and end up with a course that has no rhythm. Get them to decide their value now and write it on a sticky note.\n\n' +
    'The documentation-lag box is a genuine teaching moment beyond padding: two Articulate documentation sets exist, ' +
    'the older help.rise.com and the newer support site, and the newer one reflects current settings. ' +
    'Trust the live interface over any article.'
  );
}

// ---- 20. Backgrounds ------------------------------------------------------
{
  const s = contentSlide('Backgrounds and contrast',
    'The same panel fills a block’s white space — with a colour, a theme tint, or an image.');

  const items = [
    ['Colour', 'Type a hex code, pick manually, or take a theme colour or a tint of one.'],
    ['Image', 'A photo behind the block’s content. Handy behind a statement; risky behind a paragraph.'],
    ['Contrast', 'Auto, Light or Dark, set per block. Auto guesses from the background; override it when the guess is wrong.'],
  ];
  const gap = 0.4, w = (CW - gap * 2) / 3;
  items.forEach((it, i) => {
    const x = M + i * (w + gap);
    card(s, x, 1.95, w, 1.9);
    s.addText(it[0], {
      x: x + 0.32, y: 2.16, w: w - 0.64, h: 0.4, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 19, bold: true, color: C.ink,
    });
    s.addText(it[1], {
      x: x + 0.32, y: 2.6, w: w - 0.64, h: 1.05, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12.5, color: C.slate, lineSpacing: 18,
    });
  });

  // the banding demo — two stacks side by side
  const dy = 4.12;
  const dW = (CW - 0.5) / 2;
  card(s, M, dy, dW, 2.6, { fill: C.paper });
  s.addText('Banding — what to avoid', {
    x: M + 0.3, y: dy + 0.16, w: dW - 0.6, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 12.5, bold: true, color: 'B23B23',
  });
  ['E8EDF3', 'FFFFFF', 'E8EDF3', 'FFFFFF'].forEach((cc, i) => {
    s.addShape(pres.ShapeType.rect, {
      x: M + 0.3, y: dy + 0.56 + i * 0.4, w: dW - 0.6, h: 0.36,
      fill: { color: cc }, line: { color: C.line, width: 0.75 },
    });
  });
  s.addText('Alternating fills stripe the page and flatten the hierarchy.', {
    x: M + 0.3, y: dy + 2.14, w: dW - 0.6, h: 0.34, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 11, color: C.slate, italic: true,
  });

  const rx2 = M + dW + 0.5;
  card(s, rx2, dy, dW, 2.6, { fill: C.paper });
  s.addText('Grouping — what to do instead', {
    x: rx2 + 0.3, y: dy + 0.16, w: dW - 0.6, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 12.5, bold: true, color: C.ok,
  });
  ['FFFFFF', 'E8EDF3', 'E8EDF3', 'FFFFFF'].forEach((cc, i) => {
    s.addShape(pres.ShapeType.rect, {
      x: rx2 + 0.3, y: dy + 0.56 + i * 0.4, w: dW - 0.6, h: 0.36,
      fill: { color: cc }, line: { color: C.line, width: 0.75 },
    });
  });
  s.addText('One fill across blocks that belong together reads as a section.', {
    x: rx2 + 0.3, y: dy + 2.14, w: dW - 0.6, h: 0.34, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 11, color: C.slate, italic: true,
  });
  s.addNotes(
    'Backgrounds are where enthusiastic beginners do the most damage, so lead with the failure mode.\n\n' +
    'The banded look on the left happens when someone colours every other block to "add interest". ' +
    'The result is a stripey page where nothing stands out because everything does. ' +
    'Show the two stacks and let them see it.\n\n' +
    'The fix on the right: use a background to group blocks that belong together, so the fill means "this is one section". ' +
    'The other fix, if you want visual transitions without stripes, is full-width imagery between sections.\n\n' +
    'Contrast: Auto is right most of the time and wrong on busy images. If text becomes hard to read over a background ' +
    'image, that is what Light and Dark are for. Check it on mobile, where the crop changes what is behind the text.'
  );
}

// ---- 21. Animations -------------------------------------------------------
{
  const s = contentSlide('Block entrance animations',
    'On by default for non-text blocks. One toggle, course-wide, and worth a deliberate decision.');

  card(s, M, 2.05, CW * 0.52, 2.85);
  s.addText('Where it lives', {
    x: M + 0.4, y: 2.3, w: CW * 0.52 - 0.8, h: 0.4, margin: 0, valign: 'middle',
    fontFace: HEAD, fontSize: 20, bold: true, color: C.ink,
  });
  s.addText('Theme  ›  Blocks  ›  Block Entrance Animations', {
    x: M + 0.4, y: 2.76, w: CW * 0.52 - 0.8, h: 0.42, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 15, bold: true, color: C.ember,
  });
  s.addText(
    'Not in a block’s own settings — this is a Theme-level switch that applies to the whole course. ' +
    'Non-text blocks fade in as the learner scrolls to them; text blocks do not animate.',
    { x: M + 0.4, y: 3.3, w: CW * 0.52 - 0.8, h: 1.05, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 19 });

  const rx = M + CW * 0.52 + 0.5;
  const rw = CW - CW * 0.52 - 0.5;
  card(s, rx, 2.05, rw, 2.85, { fill: C.ink, line: C.ink });
  s.addText('Turn it off when…', {
    x: rx + 0.4, y: 2.3, w: rw - 0.8, h: 0.4, margin: 0, valign: 'middle',
    fontFace: HEAD, fontSize: 20, bold: true, color: 'FFFFFF',
  });
  s.addText([
    { text: 'the course is reference material people scan rather than read', options: { bullet: true, breakLine: true } },
    { text: 'your audience includes people sensitive to motion', options: { bullet: true, breakLine: true } },
    { text: 'the lesson is long and the repeated fade becomes a drag', options: { bullet: true } },
  ], {
    x: rx + 0.4, y: 2.78, w: rw - 0.8, h: 1.95, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13, color: C.slateLt, lineSpacing: 19, paraSpaceAfter: 8,
  });

  card(s, M, 5.16, CW, 1.45, { fill: C.paper });
  s.addText('Why a facilitator should mention it at all', {
    x: M + 0.4, y: 5.3, w: CW - 0.8, h: 0.34, margin: 0,
    fontFace: HEAD, fontSize: 17, bold: true, color: C.ink,
  });
  s.addText(
    'Because it is the most common "why does my course feel slow?" question, and because nobody finds it — ' +
    'people look in the block settings, where it is not. Show them the Theme menu once and it stops being a mystery. ' +
    'It also previews the Day 4 point that all global styling lives under Theme.',
    { x: M + 0.4, y: 5.68, w: CW - 0.8, h: 0.85, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 19 });
  s.addNotes(
    'Two minutes. This is a small setting with an outsized number of support questions attached to it.\n\n' +
    'Show the toggle live. The reason it matters is location, not function: it is under Theme, not under the block, ' +
    'and that is exactly the confusion that comes back on Day 4 with navigation — course navigation lives under ' +
    'Theme too, while each quiz has its own Settings panel. Plant the pattern now: block-level settings are on the block, ' +
    'course-level settings are under Theme.\n\n' +
    'Do not turn this into an accessibility lecture, but do name motion sensitivity as a real reason to switch it off.'
  );
}

// ---- 22. Section 4 --------------------------------------------------------
sectionSlide('SECTION FOUR', 'Keeping it readable', 'Four rules of thumb, and the five mistakes every first Rise course makes.')
  .addNotes(
    'Change of register here. Everything so far has been "what the tool does". This section is judgement.\n\n' +
    'Be explicit that these are practitioner conventions, not settings the software enforces. ' +
    'That honesty buys you credibility for the rest of the day.'
  );

// ---- 23. The rules --------------------------------------------------------
{
  const s = contentSlide('Four rules of thumb',
    'None of these is enforced by Rise. All four are what separates a first course from a fifth.');
  const rules = [
    ['≤ 50%', 'text blocks', 'No more than about half a lesson should be text blocks. The rest earns its place as images, media, dividers and — from tomorrow — interactions.'],
    ['1–2+', 'images per lesson', 'At minimum. A lesson with no picture in it is a document, and people read documents somewhere other than an LMS.'],
    ['≤ 20', 'blocks per lesson', 'Past roughly twenty blocks, ask whether this is really two or three shorter lessons. Rise scrolls forever; attention does not.'],
    ['0', 'stripes', 'Never alternate background colours block by block. Group with one fill, or transition with full-width imagery.'],
  ];
  const gap = 0.3, w = (CW - gap * 3) / 4;
  rules.forEach((r, i) => {
    const x = M + i * (w + gap);
    card(s, x, 2.0, w, 3.35);
    s.addText(r[0], {
      x: x + 0.26, y: 2.28, w: w - 0.52, h: 0.72, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 38, bold: true, color: C.ember,
    });
    s.addText(r[1], {
      x: x + 0.26, y: 3.02, w: w - 0.52, h: 0.34, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, bold: true, color: C.ink,
    });
    s.addText(r[2], {
      x: x + 0.26, y: 3.46, w: w - 0.52, h: 1.7, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12.5, color: C.slate, lineSpacing: 18,
    });
  });
  card(s, M, 5.62, CW, 0.78, { fill: 'FFF4EE', line: C.emberDim, flat: true });
  s.addText([
    { text: 'Say this out loud to your group:  ', options: { bold: true, color: '7A3F1C' } },
    { text: 'these figures come from practitioner guidance in the Articulate community, not from Articulate’s own documentation. They are good defaults to design against, not standards to audit against.', options: { color: '7A3F1C' } },
  ], {
    x: M + 0.38, y: 5.62, w: CW - 0.76, h: 0.78, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, lineSpacing: 18,
  });
  s.addNotes(
    'The credibility move on this slide is the orange box. Do not present community heuristics as vendor rules — ' +
    'somebody in the room will go looking for them in the Articulate documentation, not find them, and then discount ' +
    'everything else you said today.\n\n' +
    'The 50% rule is the useful one and the one to enforce in the lab. It is easy to check: count the blocks, ' +
    'count the text blocks.\n\n' +
    'The 20-block threshold usually lands with people who have already written a monster lesson. ' +
    'Ask who has one. Then ask where it would naturally split.\n\n' +
    'The stripe rule connects straight back to the banding demo two slides ago.'
  );
}

// ---- 24. Common mistakes --------------------------------------------------
{
  const s = contentSlide('The five mistakes every first course makes',
    'Recognise them in your own draft before someone else does.');
  const mistakes = [
    ['Too much text', 'Paragraphs that run past four sentences, and every idea delivered as prose.', 'Cut to need-to-know. Pull one line into a Statement.'],
    ['Not enough visuals', 'A lesson with no image in it at all.', 'One or two image blocks, minimum. Caption them.'],
    ['No block variety', 'Paragraph, paragraph, paragraph, list, paragraph.', 'Vary the family. Aim for no more than half text.'],
    ['No pacing', 'One unbroken run of blocks with no visual structure.', 'Dividers between sections; Continue at real checkpoints.'],
    ['Endless lessons', 'One lesson carrying an entire topic.', 'Split it. Two or three short lessons beat one long one.'],
  ];
  const rowH = 0.82, gapY = 0.14;
  mistakes.forEach((m, i) => {
    const y = 1.9 + i * (rowH + gapY);
    card(s, M, y, CW, rowH);
    s.addShape(pres.ShapeType.ellipse, {
      x: M + 0.28, y: y + rowH / 2 - 0.17, w: 0.34, h: 0.34,
      fill: { color: C.ember }, line: { type: 'none' },
    });
    s.addText(String(i + 1), {
      x: M + 0.28, y: y + rowH / 2 - 0.17, w: 0.34, h: 0.34, margin: 0, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 12, bold: true, color: 'FFFFFF',
    });
    s.addText(m[0], {
      x: M + 0.78, y, w: 2.55, h: rowH, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 16, bold: true, color: C.ink,
    });
    s.addText(m[1], {
      x: M + 3.4, y, w: 4.5, h: rowH, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12.5, color: C.slate, lineSpacing: 17,
    });
    s.addText(m[2], {
      x: M + 8.05, y, w: CW - 8.35, h: rowH, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12.5, bold: true, color: C.ok, lineSpacing: 17,
    });
  });
  s.addNotes(
    'Read the middle column, not the left — the symptom is what people recognise in their own work; the label is not.\n\n' +
    'Best way to run this slide: have someone volunteer a lesson they have already drafted, put it on the screen, ' +
    'and let the group name which of the five it has. It is much more effective than agreeing with the slide in the abstract.\n\n' +
    'Be kind about it. Everyone’s first Rise course has at least three of these, including yours. ' +
    'The point is that they are all cheap to fix once you can see them — which is why we spent the morning ' +
    'learning what else is available.'
  );
}

// ---- 25. Section 5 --------------------------------------------------------
sectionSlide('SECTION FIVE', 'Build it', 'The rest of today is hands on. One lesson, built to a brief.')
  .addNotes(
    'Stop talking. The remaining time is theirs.\n\n' +
    'Have the brief slide up on the screen for the whole lab so nobody has to ask what the requirements were.'
  );

// ---- 26. Lab brief --------------------------------------------------------
{
  const s = contentSlide('Your build: one varied lesson',
    'Fill out the lesson you created this morning. A real topic from your own work — you will want to keep this.');

  const colW = (CW - 0.5) / 2;
  card(s, M, 1.95, colW, 4.05, { fill: C.ink, line: C.ink });
  s.addText('THE BRIEF', {
    x: M + 0.4, y: 2.2, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Eight to fourteen blocks', options: { bullet: true, breakLine: true } },
    { text: 'No more than half of them text', options: { bullet: true, breakLine: true } },
    { text: 'Two or more image or gallery blocks, captioned', options: { bullet: true, breakLine: true } },
    { text: 'One statement or quote carrying the key point', options: { bullet: true, breakLine: true } },
    { text: 'A divider, and a Continue where a pause helps', options: { bullet: true, breakLine: true } },
    { text: 'One padding value throughout', options: { bullet: true, breakLine: true } },
    { text: 'No two adjacent blocks with different fills', options: { bullet: true } },
  ], {
    x: M + 0.4, y: 2.58, w: colW - 0.8, h: 3.3, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13, color: 'FFFFFF', lineSpacing: 19, paraSpaceAfter: 7,
  });

  const rx = M + colW + 0.5;
  card(s, rx, 1.95, colW, 1.92);
  s.addText('How you will know it works', {
    x: rx + 0.4, y: 2.15, w: colW - 0.8, h: 0.36, margin: 0,
    fontFace: HEAD, fontSize: 18, bold: true, color: C.ink,
  });
  s.addText(
    'Swap laptops with the person next to you. They read your lesson on a phone preview and tell you, without ' +
    'scrolling back, what the main point was. If they cannot, the layout is doing the wrong job.',
    { x: rx + 0.4, y: 2.56, w: colW - 0.8, h: 0.95, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 19 });

  card(s, rx, 4.08, colW, 1.92);
  s.addText('If you finish early', {
    x: rx + 0.4, y: 4.28, w: colW - 0.8, h: 0.36, margin: 0,
    fontFace: HEAD, fontSize: 18, bold: true, color: C.ink,
  });
  s.addText(
    'Rebuild your busiest text block three ways using the swap dropdown, and keep the version that reads fastest. ' +
    'Then try one chart block — most people never open that family unprompted.',
    { x: rx + 0.4, y: 4.69, w: colW - 0.8, h: 1.15, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: C.slate, lineSpacing: 19 });

  card(s, M, 6.12, CW, 0.7, { fill: 'FFF4EE', line: C.emberDim, flat: true });
  s.addText('Do not add accordions, tabs, flashcards or knowledge checks today — they are tomorrow, and the brief is deliberately about what static blocks alone can do.', {
    x: M + 0.38, y: 6.12, w: CW - 0.76, h: 0.7, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, color: '7A3F1C',
  });
  s.addNotes(
    'Give them a real block of time — an hour is not too much. Walk the room; do not narrate from the front.\n\n' +
    'They already have the empty lesson from this morning, so nobody should be starting from the dashboard. ' +
    'If anyone is, sort it out before the lab proper begins.\n\n' +
    'Insist on a real topic. People build a throwaway lesson about coffee, learn nothing they can reuse, and delete it. ' +
    'A lesson from their actual work becomes the seed of a real course and they will care whether the layout is any good.\n\n' +
    'The peer test in the top-right corner is the part to protect. It is the only feedback in the day that comes from ' +
    'someone who does not already know what the lesson is supposed to say — and reading it on a phone surfaces ' +
    'cropping and padding problems immediately.\n\n' +
    'Enforce the no-interactions rule gently but firmly. Someone always sneaks in an accordion. The constraint is the ' +
    'lesson: static blocks alone can carry a good lesson, and if you skip straight to interactions you never learn that.'
  );
}

// ---- 27. Checklist --------------------------------------------------------
{
  const s = contentSlide('Before you call it done',
    'Run this against your lesson. It is also the checklist you will use on every course after today.');
  const groups = [
    ['Content', ['Every paragraph is four sentences or fewer', 'Every image and gallery slot has a real caption', 'Lists use the marker that matches their meaning', 'Any chart has the finding stated above it']],
    ['Layout', ['One padding value, used throughout', 'No two adjacent blocks with different fills', 'Dividers mark the sections a reader would expect', 'Continue blocks sit at real checkpoints only']],
    ['Check', ['Previewed on mobile portrait, not just desktop', 'Nothing important sits near an image edge', 'Under twenty blocks, or split into two lessons', 'No more than half the blocks are text']],
  ];
  const gap = 0.4, w = (CW - gap * 2) / 3;
  groups.forEach((g, i) => {
    const x = M + i * (w + gap);
    card(s, x, 2.05, w, 4.1);
    s.addText(g[0], {
      x: x + 0.34, y: 2.3, w: w - 0.68, h: 0.4, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 20, bold: true, color: C.ink,
    });
    g[1].forEach((line, j) => {
      const y = 2.9 + j * 0.82;
      s.addShape(pres.ShapeType.roundRect, {
        x: x + 0.34, y: y + 0.04, w: 0.26, h: 0.26, rectRadius: 0.05,
        fill: { color: C.ground }, line: { color: C.slateLt, width: 1 },
      });
      s.addText(line, {
        x: x + 0.72, y, w: w - 1.06, h: 0.7, margin: 0, valign: 'top',
        fontFace: BODY, fontSize: 12.5, color: C.slate, lineSpacing: 17,
      });
    });
  });
  s.addNotes(
    'Hand this out on paper if you can. It is the artefact people keep from today.\n\n' +
    'Walk it once, quickly — do not read all twelve lines aloud. Point out that the Check column is the one people skip: ' +
    'previewing on mobile takes ten seconds and catches the majority of layout problems while they are still cheap to fix.\n\n' +
    'Tell them this checklist grows. Day 3 adds interaction questions, Day 4 adds assessment and branding, ' +
    'Day 5 adds the publishing checks.'
  );
}

// ---- 28. Tomorrow ---------------------------------------------------------
{
  const s = contentSlide('Tomorrow: the blocks that answer back',
    'Everything today was content the learner reads. Day 3 is content the learner operates.');

  const colW = (CW - 0.5) / 2;
  card(s, M, 2.05, colW, 3.35);
  s.addText('WHAT YOU CAN DO NOW', {
    x: M + 0.4, y: 2.3, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Nine families, forty-two static blocks', options: { bullet: true, breakLine: true } },
    { text: 'Add, reorder, swap, duplicate, delete', options: { bullet: true, breakLine: true } },
    { text: 'Padding and backgrounds set on purpose', options: { bullet: true, breakLine: true } },
    { text: 'A lesson that passes the anti-clutter rules', options: { bullet: true } },
  ], {
    x: M + 0.4, y: 2.7, w: colW - 0.8, h: 2.4, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: C.ink, lineSpacing: 21, paraSpaceAfter: 9,
  });

  const rx = M + colW + 0.5;
  card(s, rx, 2.05, colW, 3.35, { fill: C.ink, line: C.ink });
  s.addText('WHAT DAY 3 ADDS', {
    x: rx + 0.4, y: 2.3, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Accordion and Tabs — hide secondary text', options: { bullet: true, breakLine: true } },
    { text: 'Labeled Graphic — an image to explore', options: { bullet: true, breakLine: true } },
    { text: 'Process and Timeline — sequence with pacing', options: { bullet: true, breakLine: true } },
    { text: 'Sorting and Scenario — decisions, not reading', options: { bullet: true, breakLine: true } },
    { text: 'Flashcards, Buttons and Button stacks', options: { bullet: true } },
  ], {
    x: rx + 0.4, y: 2.7, w: colW - 0.8, h: 2.6, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: 'FFFFFF', lineSpacing: 21, paraSpaceAfter: 9,
  });

  card(s, M, 5.66, CW, 0.78, { fill: C.paper });
  s.addText([
    { text: 'Bring your lesson with you.  ', options: { bold: true, color: C.ink } },
    { text: 'Tomorrow starts by taking the most text-heavy block in the lesson you just built and turning it into an interaction. That works much better on real content than on a blank page.', options: { color: C.slate } },
  ], {
    x: M + 0.38, y: 5.66, w: CW - 0.76, h: 0.78, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 13, lineSpacing: 18,
  });
  s.addNotes(
    'Close the loop back to the contract you set on slide 3. Ask the room to say what they can now do — ' +
    'do not read the left column to them.\n\n' +
    'The reason to preview Day 3 is motivational, not informational: the accordion is the block that solves the ' +
    'problem they have all just felt, which is having more to say than a lesson can comfortably hold.\n\n' +
    'The homework line matters. If they arrive tomorrow with a real lesson, Day 3 has real content to transform and ' +
    'the exercises are far better. If they arrive empty-handed we spend the first half hour inventing filler.'
  );
}

// ---- 29. Close ------------------------------------------------------------
{
  const s = darkSlide();
  s.addText('DAY 2 COMPLETE', {
    x: M, y: 2.35, w: 6.4, h: 0.32, margin: 0,
    fontFace: BODY, fontSize: 12.5, bold: true, color: C.ember, charSpacing: 2.6,
  });
  s.addText('Nine families.\nOne readable lesson.', {
    x: M, y: 2.76, w: 7.4, h: 1.7, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 40, bold: true, color: 'FFFFFF', lineSpacing: 48,
  });
  s.addText(
    'The tool was never the hard part. Choosing which block carries which idea — that is the craft, and you have started it.',
    { x: M, y: 4.6, w: 6.9, h: 0.9, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 14.5, color: C.slateLt, lineSpacing: 22 });
  stackGlyph(s, W - M - 3.4, 2.5, { w: 3.4, h: 0.62, gap: 0.22, stagger: true });
  s.addText('Day 3  ›  Interactive blocks', {
    x: W - M - 3.4, y: 5.06, w: 3.4, h: 0.3, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 12, color: C.slateLt, italic: true,
  });
  footer(s, true);
  s.addNotes(
    'Two things before people leave.\n\n' +
    'One: remind them the deliverable is the lesson, and to bring it tomorrow.\n\n' +
    'Two: collect one sentence each on what they found hardest today. That tells you whether Day 3 needs to open with ' +
    'a recap of block mechanics or can go straight into interactions.\n\n' +
    'If the room struggled with the tool rather than the choices — hunting for settings, losing blocks — slow Day 3 down ' +
    'and re-demo the block toolbar first. If they struggled with the choices, that is exactly right, and Day 3 will help.'
  );
}

pres.writeFile({ fileName: OUT }).then(() => {
  console.log('wrote', OUT, '| slides:', slideNo);
});
