/**
 * Day 4 — "Decide, then ship"
 * Articulate Rise 360, 5-day beginner curriculum.
 *
 * Two blocks in depth — Button / Button stack and Scenario — then a live demo of a
 * complete course (two lessons and a quiz), then the four ways out of Rise, then the
 * capstone hand-over.
 *
 * The demo run sheet lives in ../practice/demo-script.md and the hand-out in
 * ../practice/capstone-brief.md. Slides 17 and 27-28 are the rails for those two files;
 * the detail stays in the files, not on the slides.
 *
 * Same design system as Days 2 and 3 — the block card motif, the ember accent sampled
 * from the Rise thumbnails, dark title/section/close, light content.
 */
const path = require('path');
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

const IMG = path.join(__dirname, '..', '..', 'day2', 'assets', 'block-screenshots');
const OUT = path.join(__dirname, '..', 'build', 'rise360-day4-decide-and-ship.pptx');
const META = JSON.parse(fs.readFileSync(
  path.join(__dirname, '..', '..', 'day3', 'assets', 'rise_blocks_metadata.json'), 'utf8'));

const C = {
  ink: '12161F', inkSoft: '1C2230', paper: 'FFFFFF', ground: 'F4F6F8',
  ember: 'ED6A2A', emberDim: 'F6C7AC', emberPale: 'FFF4EE', emberInk: '7A3F1C',
  slate: '5A6675', slateLt: 'A9B3C0', line: 'E2E6EB', ok: '2E7D5B', bad: 'B4441F',
};
const HEAD = 'Cambria', BODY = 'Calibri';
const W = 13.333, M = 0.62, CW = W - M * 2, FOOT_Y = 6.98;

const pres = new PptxGenJS();
pres.layout = 'LAYOUT_WIDE';
pres.author = 'Rise 360 curriculum';
pres.title = 'Rise 360 Day 4 — Decide, then ship';

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
  const h = o.h || 0.3;
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: h / 2,
    fill: { color: o.fill || C.ember }, line: o.line ? { color: o.line, width: 1 } : { type: 'none' },
  });
  slide.addText(text, {
    x, y, w, h, align: o.align || 'center', valign: 'middle', margin: 0,
    fontFace: BODY, fontSize: o.size || 11, bold: o.bold !== false,
    color: o.color || 'FFFFFF', charSpacing: o.charSpacing === undefined ? 0.6 : o.charSpacing,
  });
}

let slideNo = 0;
function footer(slide, dark) {
  slideNo += 1;
  slide.addText('Rise 360  ·  Day 4  ·  Decide, then ship', {
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
      x: M, y: 1.1, w: o.sfW || CW * 0.82, h: 0.5, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 14, color: C.slate,
    });
  }
  footer(s, false);
  return s;
}

const darkSlide = () => {
  const s = pres.addSlide();
  s.background = { color: C.ink };
  return s;
};

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
      x: M, y: 4.06, w: CW * 0.54, h: 0.9, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 14.5, color: C.slateLt, lineSpacing: 22,
    });
  }
  stackGlyph(s, W - M - 2.5, 2.55, { stagger: true });
  footer(s, true);
  return s;
}

/** The pale orange strip that carries the one thing not to forget. */
function strip(slide, y, text, o = {}) {
  const h = o.h || 0.6;
  card(slide, M, y, CW, h, { fill: C.emberPale, line: C.emberDim, flat: true });
  if (o.dot !== false) {
    slide.addShape(pres.ShapeType.ellipse, {
      x: M + 0.24, y: y + (h - 0.22) / 2, w: 0.22, h: 0.22,
      fill: { color: C.ember }, line: { type: 'none' },
    });
  }
  slide.addText(text, {
    x: M + (o.dot === false ? 0.38 : 0.6), y, w: CW - (o.dot === false ? 0.76 : 0.88), h,
    margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: o.size || 12.5, color: C.emberInk, lineSpacing: 17,
  });
}

const block = (name) => META.find((b) => b.block_name === name);
const imgOf = (name) => path.join(IMG, path.basename(block(name).local_image_path));

/** Zebra table with a dark header. cols = [{label, w, style}]. */
function table(slide, cols, rows, o = {}) {
  const top = o.top === undefined ? 1.86 : o.top;
  const rowH = o.rowH === undefined ? 0.47 : o.rowH;
  const headH = 0.44, pad = 0.3;

  slide.addShape(pres.ShapeType.roundRect, {
    x: M, y: top, w: CW, h: headH, rectRadius: 0.06,
    fill: { color: C.ink }, line: { type: 'none' },
  });
  let x = M;
  cols.forEach((c) => {
    slide.addText(c.label, {
      x: x + pad, y: top, w: c.w - pad * 1.4, h: headH, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 11.5, bold: true, color: 'FFFFFF', charSpacing: 0.8,
    });
    x += c.w;
  });

  rows.forEach((r, i) => {
    const y = top + headH + i * rowH;
    slide.addShape(pres.ShapeType.rect, {
      x: M, y, w: CW, h: rowH,
      fill: { color: i % 2 === 0 ? 'FFFFFF' : 'EDF1F5' },
      line: { color: C.line, width: 0.5 },
    });
    let cx = M;
    cols.forEach((c, j) => {
      slide.addText(r[j], {
        x: cx + pad, y, w: c.w - pad * 1.4, h: rowH, margin: 0, valign: 'middle',
        fontFace: BODY, fontSize: o.size || 12.5,
        bold: !!c.bold, color: c.color || C.ink, lineSpacing: 15,
      });
      cx += c.w;
    });
  });
  return top + headH + rows.length * rowH;
}

/** Two blocks side by side, each with its screenshot, its job and an example. */
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

  strip(s, 6.3, o.note, { h: 0.56, size: 12 });
  s.addNotes(o.notes);
  return s;
}

// ================================================================ SLIDES

// ---- 1. Title -------------------------------------------------------------
{
  const s = darkSlide();
  s.addText('DAY 4 OF 5', {
    x: M, y: 1.62, w: 4, h: 0.32, margin: 0,
    fontFace: BODY, fontSize: 12.5, bold: true, color: C.ember, charSpacing: 2.6,
  });
  s.addText('Decide,\nthen ship', {
    x: M, y: 2.04, w: 7.6, h: 2.0, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 50, bold: true, color: 'FFFFFF', lineSpacing: 56,
  });
  s.addText('Two blocks in depth, a whole course built in front of you, and the four ways out of Rise.',
    { x: M, y: 4.18, w: 6.9, h: 0.9, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 15.5, color: C.slateLt, lineSpacing: 24 });

  const gx = 8.55, gy = 2.0, gw = 3.8;
  [['Decide', C.inkSoft], ['Build', C.inkSoft], ['Ship', C.ember], ['Hand over', C.inkSoft]].forEach((r, i) => {
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
  s.addText('nothing counts until it has a link', {
    x: gx, y: 5.55, w: gw, h: 0.3, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 11, color: C.slateLt, italic: true,
  });
  footer(s, true);
  s.addNotes(
    'Today is the day the week turns into something real. Three days of blocks, and this afternoon they ' +
    'watch a course go from a blank page to a link they can open on their own phones.\n\n' +
    'Two blocks get proper attention this morning. Buttons, because Day 3 only taught the labelling and there ' +
    'is more to them. Scenario, because it is the most expensive block in Rise to write well and the one ' +
    'people get most wrong.\n\n' +
    'Set the expectation for the afternoon now: they watch the demo, they do not build along. Building along ' +
    'means everyone falls behind at a different point and you spend the session rescuing people. They get ' +
    'their turn in the capstone.\n\n' +
    'Say the word "Monday" early. The brief goes out at the end of today and it is due Monday morning.'
  );
}

// ---- 2. Yesterday to today ------------------------------------------------
{
  const s = contentSlide('From choosing blocks to shipping a course',
    'Three days of parts. Today they become a thing with a link on it.');
  const colW = (CW - 0.9) / 2;

  card(s, M, 1.85, colW, 3.9);
  s.addText('BY NOW YOU CAN', {
    x: M + 0.4, y: 2.12, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Build a readable lesson out of static blocks', options: { bullet: true, breakLine: true } },
    { text: 'Name the twelve interactive blocks by the job they do', options: { bullet: true, breakLine: true } },
    { text: 'Say why an interaction earns its place — or drop it', options: { bullet: true } },
  ], {
    x: M + 0.4, y: 2.5, w: colW - 0.8, h: 2.2, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 14, color: C.ink, lineSpacing: 21, paraSpaceAfter: 10,
  });
  s.addText('What you have never done is finish one.', {
    x: M + 0.4, y: 4.9, w: colW - 0.8, h: 0.6, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13, italic: true, color: C.slate,
  });

  card(s, M + colW + 0.9, 1.85, colW, 3.9, { fill: C.ink, line: C.ink });
  s.addText('BY TONIGHT YOU WILL', {
    x: M + colW + 1.3, y: 2.12, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Build a scenario that branches, and forgives', options: { bullet: true, breakLine: true } },
    { text: 'Send a learner somewhere without taking navigation away', options: { bullet: true, breakLine: true } },
    { text: 'Have watched a course go from blank to a live link', options: { bullet: true, breakLine: true } },
    { text: 'Know which of the four exports your organisation needs', options: { bullet: true, breakLine: true } },
    { text: 'Be holding a brief that is due Monday morning', options: { bullet: true } },
  ], {
    x: M + colW + 1.3, y: 2.5, w: colW - 0.8, h: 3.0, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 14, color: 'FFFFFF', lineSpacing: 21, paraSpaceAfter: 10,
  });

  strip(s, 5.96, 'Theme, navigation and accessibility are tomorrow. Today is decision, then delivery.', { dot: false });
  s.addNotes(
    'The line on the left that does the work is the italic one: they have never finished a course. Everything ' +
    'so far has been parts on a bench.\n\n' +
    'Read the right-hand list slowly. Four of the five are things they will be asked to do at work within a ' +
    'month of this training.\n\n' +
    'Head off the theming question — somebody always wants to talk about brand colours and fonts. Point at the ' +
    'orange strip. Tomorrow.'
  );
}

// ---- 3. Today's shape -----------------------------------------------------
{
  const s = contentSlide('The shape of today',
    'Two blocks in depth this morning. This afternoon, the whole thing end to end.');
  const bands = [
    ['45 min', 'Buttons', 'The smallest block in Rise, and the one with the biggest labelling problem.'],
    ['60 min', 'Scenario', 'Role-play with consequences. The most expensive block in the tool to write well.'],
    ['45 min', 'Watch me build one', 'A blank course to a live link. Nothing skipped, including the bits that go wrong.'],
    ['30 min', 'Ship it', 'Preview like a learner, then the four ways out and what each one can tell you.'],
    ['15 min', 'Your turn', 'The capstone brief. Same shape, your subject, due Monday morning.'],
  ];
  const top = 1.84, h = 0.8, gap = 0.12;
  bands.forEach((b, i) => {
    const y = top + i * (h + gap);
    const last = i === bands.length - 1;
    card(s, M, y, CW, h, last ? { fill: C.ink, line: C.ink } : {});
    pill(s, M + 0.3, y + (h - 0.3) / 2, 0.94, b[0], { fill: last ? C.ember : C.ink, size: 10.5 });
    s.addText(b[1], {
      x: M + 1.42, y, w: 3.1, h, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 17, bold: true, color: last ? 'FFFFFF' : C.ink,
    });
    s.addText(b[2], {
      x: M + 4.6, y, w: CW - 4.9, h, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12.5, color: last ? C.slateLt : C.slate,
    });
  });
  s.addText('Breaks fall between the bands. If you run late, take it out of the morning — the demo cannot be rushed.', {
    x: M, y: 6.5, w: CW, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, italic: true, color: C.slate,
  });
  s.addNotes(
    'Show the timings honestly. People plan their attention around them.\n\n' +
    'The one band that must not be compressed is the demo. If the morning overruns, cut a scenario example, ' +
    'not the demo — watching a whole course get finished is the thing they cannot get from a slide or a video ' +
    'later.\n\n' +
    'The last band is only fifteen minutes because the brief is written down. Hand it out, walk the checklist, ' +
    'take questions, stop.'
  );
}

// ---- 4. Section one -------------------------------------------------------
sectionSlide('SECTION ONE', 'Buttons that say\nwhat they do',
  'Two blocks, five destinations and one habit that separates a usable course from a guessing game.')
  .addNotes(
    'Day 3 gave them one slide on buttons and it was all about labelling. That lesson stands — repeat it once ' +
    'and move on.\n\n' +
    'What is new today: where a button can actually send someone, what you can change about how it looks, and ' +
    'the thing buttons are genuinely good for — offering routes without restricting navigation.\n\n' +
    'Forty-five minutes including a live look at the settings panel. Do not spend it all on slides.'
  );

// ---- 5. Button and Button stack -------------------------------------------
jobSlide({
  title: 'Two blocks, one idea',
  standfirst: 'A button is a promise about what happens next. A stack is two or three, offered side by side.',
  blocks: [
    { name: 'Button',
      use: 'one clear next action sits at the end of a run of content.',
      eg: 'Open the one-page feedback template — a file the learner is about to need.' },
    { name: 'Button stack',
      use: 'more than one next move is legitimate and the learner should pick.',
      eg: '“Take the check” beside “Read the template again”, at the end of a lesson.' },
  ],
  note: 'Day 3 taught you to label them. Today: where they can send people, and what you can change about how they look.',
  notes:
    'Keep this short — the blocks themselves are simple and the room already met them.\n\n' +
    'The framing sentence is the one to say out loud: a button is a promise about what happens next. Every ' +
    'labelling failure is a broken promise, and every learner who clicks one and lands somewhere unexpected ' +
    'trusts the course a little less.\n\n' +
    'The stack is the more useful of the two and it gets under-used. Most people reach for a stack only for ' +
    'external links, when its best job is inside the course.',
});

// ---- 6. Five destinations -------------------------------------------------
{
  const s = contentSlide('Where a button can send someone',
    'Five destinations. Pick from the Destination drop-down on the block.');
  table(s, [
    { label: 'Destination', w: 3.5, bold: true },
    { label: 'What happens', w: 4.3, color: C.slate },
    { label: 'Reach for it when', w: CW - 7.8, color: C.slate },
  ], [
    ['Link to a webpage', 'Opens a URL in a new tab', 'A policy, a form, a template that lives somewhere else'],
    ['Link to a relative URL', 'Opens a page on your own subdomain, new tab', 'You host the target yourself'],
    ['Send an email', 'Opens a draft; several addresses allowed', 'The next step needs a human, not a page'],
    ['Navigate to Another Lesson', 'Moves inside the course, no reload', 'Offering routes through your own content'],
    ['Exit the course (LMS only)', 'Closes the course for a learner in an LMS', 'A deliberate, tracked ending'],
  ], { rowH: 0.62 });

  strip(s, 5.5, 'Four of the five take the learner out of your course. Only “Navigate to Another Lesson” keeps them in it — and the learner cannot tell which is which unless your label says so.');
  s.addText('Labels move. If the drop-down on your screen reads differently, the screen is right — Articulate ships changes to Rise every few weeks.', {
    x: M, y: 6.3, w: CW, h: 0.4, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 11, italic: true, color: C.slate,
  });
  s.addNotes(
    'Open the block settings live while this is up. Reading a table aloud is the worst use of five minutes ' +
    'you have today.\n\n' +
    '“Send an email” surprises people every time — it opens the learner’s mail client with a draft, and it ' +
    'takes several addresses separated by commas. It is the right answer for “request access”, “contact your ' +
    'manager”, “report a problem”.\n\n' +
    '“Exit the course” only does anything in an LMS. In Quick Share or a preview it does nothing at all, which ' +
    'is exactly the kind of thing that gets reported as a bug the week after a launch.\n\n' +
    'The italic line at the bottom is worth saying out loud once and then living by all day. When the interface ' +
    'disagrees with the deck, believe the interface — and let them watch you do that calmly.'
  );
}

// ---- 7. Settings ----------------------------------------------------------
{
  const s = contentSlide('What you can change — and what you cannot',
    'Everything here lives in the block settings. None of it rescues a bad label.');
  const colW = (CW - 0.9) / 2;

  card(s, M, 1.85, colW, 3.55);
  s.addText('YOU CAN CHANGE', {
    x: M + 0.4, y: 2.1, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Button colour, by picker or hex code', options: { bullet: true, breakLine: true } },
    { text: 'Alignment, and the corner radius', options: { bullet: true, breakLine: true } },
    { text: 'Button width', options: { bullet: true, breakLine: true } },
    { text: 'Text contrast — auto, light or dark', options: { bullet: true, breakLine: true } },
    { text: 'Hide the description text, which centres the buttons', options: { bullet: true, breakLine: true } },
    { text: 'Spacing between buttons — stacks only', options: { bullet: true } },
  ], {
    x: M + 0.4, y: 2.48, w: colW - 0.8, h: 2.7, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: C.ink, lineSpacing: 20, paraSpaceAfter: 8,
  });

  card(s, M + colW + 0.9, 1.85, colW, 3.55, { fill: C.ink, line: C.ink });
  s.addText('YOU CANNOT CHANGE', {
    x: M + colW + 1.3, y: 2.1, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText('The size of the button text.', {
    x: M + colW + 1.3, y: 2.55, w: colW - 0.8, h: 0.5, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 22, bold: true, color: 'FFFFFF',
  });
  s.addText('There is no setting for it, and there is no workaround worth having. This is a real limit of the tool, not something you have failed to find.\n\nIf a label is too small to read comfortably, the label is too long. Shorten it.', {
    x: M + colW + 1.3, y: 3.2, w: colW - 0.8, h: 1.9, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: C.slateLt, lineSpacing: 20,
  });

  strip(s, 5.62, 'A button that opens a new tab and a button that moves inside the course look identical. The label is the only thing that tells them apart.');
  s.addNotes(
    'Do this one in the tool, not on the slide. Change the colour, drag the corner radius to full round and ' +
    'back, hide the description text — three settings, thirty seconds, and everyone understands the panel.\n\n' +
    'The right-hand card exists because somebody in every room hunts for the text-size setting for ten minutes ' +
    'and feels stupid. Tell them it is not there. Naming a limit honestly buys you credibility for everything ' +
    'else you say about the tool.\n\n' +
    'The strip at the bottom is the design point of the whole section. Two buttons, identical on screen, ' +
    'completely different consequences. That gap is closed by writing, not by settings.'
  );
}

// ---- 8. Routes without restriction ---------------------------------------
{
  const s = contentSlide('Give a route, do not take one away',
    'The best use of a button stack is inside your own course — offering a way on without restricting navigation.');

  const colW = 5.3;
  card(s, M, 1.82, colW, 4.3);
  s.addText('AT THE END OF A LESSON', {
    x: M + 0.4, y: 2.06, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  const bx = M + 0.62, bw = colW - 1.24;
  s.addShape(pres.ShapeType.roundRect, {
    x: bx, y: 2.5, w: bw, h: 0.5, rectRadius: 0.08,
    fill: { color: 'EDF1F5' }, line: { color: C.line, width: 1 },
  });
  s.addText('…the learner finishes “Try it”', {
    x: bx, y: 2.5, w: bw, h: 0.5, margin: 0, align: 'center', valign: 'middle',
    fontFace: BODY, fontSize: 12.5, italic: true, color: C.slate,
  });
  [['Take the check', 'the quiz'], ['Read the template again', 'lesson 1'], ['Ask for a review', 'an email draft']]
    .forEach((b, i) => {
      const y = 3.16 + i * 0.82;
      s.addShape(pres.ShapeType.roundRect, {
        x: bx, y, w: bw, h: 0.5, rectRadius: 0.08,
        fill: { color: i === 2 ? C.paper : C.ember },
        line: { color: i === 2 ? C.ember : C.ember, width: 1 },
      });
      s.addText(b[0], {
        x: bx, y, w: bw, h: 0.5, margin: 0, align: 'center', valign: 'middle',
        fontFace: BODY, fontSize: 13, bold: true, color: i === 2 ? C.ember : 'FFFFFF',
      });
      s.addText('→ ' + b[1], {
        x: bx, y: y + 0.5, w: bw, h: 0.3, margin: 0, align: 'center',
        fontFace: BODY, fontSize: 10.5, color: C.slate,
      });
    });
  s.addText('Nobody is trapped. The sidebar still works. You have offered, not enforced.', {
    x: M + 0.4, y: 5.7, w: colW - 0.8, h: 0.34, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 11.5, italic: true, color: C.slate,
  });

  const rx = M + colW + 0.55, rw = CW - colW - 0.55;
  card(s, rx, 1.82, rw, 4.3);
  s.addText('LABEL IT WITH WHAT HAPPENS', {
    x: rx + 0.4, y: 2.06, w: rw - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  const pairs = [
    ['Click here', 'Open the refund policy'],
    ['Learn more', 'See the three escalation levels'],
    ['Next', 'Take the check'],
    ['Read on', 'Email the onboarding team'],
  ];
  pairs.forEach((p, i) => {
    const y = 2.56 + i * 0.78;
    s.addText(p[0], {
      x: rx + 0.36, y, w: 1.85, h: 0.4, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 13, color: C.bad, strike: true,
    });
    s.addText('→', {
      x: rx + 2.24, y, w: 0.34, h: 0.4, margin: 0, valign: 'middle', align: 'center',
      fontFace: BODY, fontSize: 13, color: C.slateLt,
    });
    s.addText(p[1], {
      x: rx + 2.66, y, w: rw - 3.02, h: 0.4, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 13, bold: true, color: C.ok,
    });
  });
  s.addText('The test: read the label with the screen covered. Can you say where you are about to land?', {
    x: rx + 0.36, y: 5.62, w: rw - 0.72, h: 0.42, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 12, italic: true, color: C.slate,
  });

  s.addNotes(
    'Two ideas on one slide because they are the same idea. A stack offers routes; the labels are what make ' +
    'the offer legible.\n\n' +
    'The left column matters because tomorrow you will teach restricted navigation, and somebody will want to ' +
    'lock the course down. Show them the cheaper move first: offer the route, keep the freedom. Restriction is ' +
    'for compliance, not for tidiness.\n\n' +
    'Run the right column as an exercise, not a reading. Cover the green side, ask the room to fix each label, ' +
    'then reveal. They will produce better ones than these, which is the point.\n\n' +
    'The covered-screen test at the bottom is the takeaway. It is the fastest label review there is.'
  );
}

// ---- 9. Section two -------------------------------------------------------
sectionSlide('SECTION TWO', 'Role-play with\nconsequences',
  'The Scenario block is the only place in Rise where a learner finds out what their choice costs. It is also the only block you should write before you open the editor.')
  .addNotes(
    'This is the hour that matters most this morning. Slow down for it.\n\n' +
    'Two honest framings before you start. First: scenario is a writing problem wearing a software costume — ' +
    'the block takes ten minutes to build and the branching takes an afternoon to write. Second: a bad ' +
    'scenario is worse than no scenario, because it feels like a trick and learners stop trusting the course.\n\n' +
    'Day 3 gave them one slide on this block. Expect to be re-teaching from close to zero.'
  );

// ---- 10. Anatomy ----------------------------------------------------------
{
  const s = contentSlide('What a scenario is made of',
    'Five things, nested. Learn the nesting and the editor stops being confusing.');

  const imgW = 3.5, imgH = imgW / (294 / 204);
  card(s, M, 1.86, imgW + 0.5, imgH + 1.4);
  s.addImage({ path: imgOf('Scenario'), x: M + 0.25, y: 2.1, w: imgW, h: imgH });
  s.addText('Interactive › Scenario', {
    x: M + 0.25, y: 2.16 + imgH + 0.1, w: imgW, h: 0.34, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 12, bold: true, color: C.ink,
  });
  s.addText('Photographic characters, a background,\nand a conversation that forks.', {
    x: M + 0.25, y: 2.5 + imgH + 0.1, w: imgW, h: 0.5, margin: 0, align: 'center',
    fontFace: BODY, fontSize: 11.5, color: C.slate, lineSpacing: 15,
  });

  const rx = M + imgW + 1.05, rw = CW - imgW - 1.05;
  const levels = [
    ['Scenario', 'the block itself', 0],
    ['Scene', 'one background, one character', 1],
    ['Content', 'Dialogue, or Text', 2],
    ['Response', 'up to three, each with a pose', 3],
    ['Feedback', 'what the character says back', 4],
    ['Go to', 'where that choice leads', 5],
  ];
  levels.forEach((l, i) => {
    const y = 1.94 + i * 0.72;
    const x = rx + l[2] * 0.42;
    const w = rw - l[2] * 0.42;
    const isLeaf = i >= 3;
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w, h: 0.58, rectRadius: 0.08,
      fill: { color: i === 0 ? C.ink : isLeaf ? C.emberPale : C.paper },
      line: { color: i === 0 ? C.ink : isLeaf ? C.emberDim : C.line, width: 1 },
    });
    s.addText(l[0], {
      x: x + 0.3, y, w: 2.2, h: 0.58, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 15, bold: true,
      color: i === 0 ? 'FFFFFF' : isLeaf ? C.emberInk : C.ink,
    });
    s.addText(l[1], {
      x: x + 2.5, y, w: w - 2.8, h: 0.58, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12.5,
      color: i === 0 ? C.slateLt : isLeaf ? C.emberInk : C.slate,
    });
  });

  s.addText('Everything a learner experiences happens in the bottom three rows. That is where your writing time goes.', {
    x: M, y: 6.36, w: CW, h: 0.4, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 12.5, italic: true, color: C.slate,
  });
  s.addNotes(
    'Draw this nesting on a whiteboard as you talk, even though it is on the slide. People who draw it ' +
    'remember it; people who read it do not.\n\n' +
    'The single most common source of confusion in the scenario editor is not knowing which level you are ' +
    'editing — a scene, a piece of content, or a response. Once the nesting is clear, the sidebar makes sense ' +
    'immediately.\n\n' +
    'The bottom line is the honest one. The block is quick. The writing is not.'
  );
}

// ---- 11. The pieces, exactly ----------------------------------------------
{
  const s = contentSlide('The pieces, exactly',
    'What Rise gives you, what it constrains, and the one field people skip.');
  const items = [
    ['Scene', 'Add Scene, at the top of the sidebar. Each new scene inherits the background and character of the one before, so you only change what changes. Drag to reorder, collapse to navigate.'],
    ['Character', 'Photographic characters from Content Library 360 only — filter by clothing, gender or age. One character per scene, though scenes can use different people. You can hide the character entirely.'],
    ['Background', 'Content Library, or upload your own. Use landscape images: only part of the picture shows. The background blur is a toggle if it fights the text.'],
    ['Scene description', 'Rise asks for one. Write it properly — it is what a screen-reader user gets instead of the picture, and it is the field everyone skips.'],
  ];
  const gap = 0.3, w = (CW - gap * 3) / 4;
  items.forEach((it, i) => {
    const x = M + i * (w + gap);
    card(s, x, 1.9, w, 3.9, i === 3 ? { fill: C.ink, line: C.ink } : {});
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.28, y: 2.16, w: 0.44, h: 0.44, rectRadius: 0.1,
      fill: { color: C.ember }, line: { type: 'none' },
    });
    s.addText(String(i + 1), {
      x: x + 0.28, y: 2.16, w: 0.44, h: 0.44, margin: 0, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 15, bold: true, color: 'FFFFFF',
    });
    s.addText(it[0], {
      x: x + 0.28, y: 2.74, w: w - 0.56, h: 0.76, margin: 0, valign: 'top',
      fontFace: HEAD, fontSize: 17, bold: true, color: i === 3 ? 'FFFFFF' : C.ink,
    });
    s.addText(it[1], {
      x: x + 0.28, y: 3.52, w: w - 0.56, h: 2.1, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12, color: i === 3 ? C.slateLt : C.slate, lineSpacing: 17,
    });
  });
  strip(s, 5.94, 'Accessibility is not a separate task you do later. The scene description is part of writing the scene.', { dot: false });
  s.addNotes(
    'Build a scene live while this slide is up — background, character, description — and narrate the ' +
    'constraints as you hit them.\n\n' +
    'The landscape-image point saves real pain. Portrait images get cropped brutally because only a band of ' +
    'the picture shows behind the dialogue.\n\n' +
    'Card four is dark because it is the one they will skip. Say plainly that a screen-reader user gets the ' +
    'scene description instead of the picture, and that an empty field means they get nothing at all.\n\n' +
    'The inheritance behaviour in card one is a genuine time-saver and almost nobody discovers it alone: add ' +
    'scene two and it already looks like scene one.'
  );
}

// ---- 12. Dialogue vs Text -------------------------------------------------
{
  const s = contentSlide('Two kinds of content, and a hard ceiling',
    'Every scene holds one or the other. Choosing right is most of what makes a scenario read well.');
  const colW = (CW - 0.9) / 2;

  card(s, M, 1.82, colW, 3.8, { fill: C.ink, line: C.ink });
  s.addText('DIALOGUE', {
    x: M + 0.42, y: 2.06, w: colW - 0.84, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText('A decision.', {
    x: M + 0.42, y: 2.44, w: colW - 0.84, h: 0.46, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 24, bold: true, color: 'FFFFFF',
  });
  s.addText([
    { text: 'The character speaks, the learner picks a reply', options: { bullet: true, breakLine: true } },
    { text: 'Up to three responses', options: { bullet: true, breakLine: true } },
    { text: 'Each response gets its own character pose', options: { bullet: true, breakLine: true } },
    { text: 'Each response gets feedback, spoken by the character', options: { bullet: true } },
  ], {
    x: M + 0.42, y: 3.06, w: colW - 0.84, h: 2.0, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: 'FFFFFF', lineSpacing: 20, paraSpaceAfter: 8,
  });
  s.addText('Use it for the moments that actually matter.', {
    x: M + 0.42, y: 5.14, w: colW - 0.84, h: 0.4, margin: 0,
    fontFace: BODY, fontSize: 12, italic: true, color: C.slateLt,
  });

  card(s, M + colW + 0.9, 1.82, colW, 3.8);
  s.addText('TEXT', {
    x: M + colW + 1.32, y: 2.06, w: colW - 0.84, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  s.addText('A single statement.', {
    x: M + colW + 1.32, y: 2.44, w: colW - 0.84, h: 0.46, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 24, bold: true, color: C.ink,
  });
  s.addText([
    { text: 'No responses, nothing for the learner to choose', options: { bullet: true, breakLine: true } },
    { text: 'One character pose, or no character at all', options: { bullet: true, breakLine: true } },
    { text: 'Sets the scene, resets after a wrong turn, or ends it', options: { bullet: true } },
  ], {
    x: M + colW + 1.32, y: 3.06, w: colW - 0.84, h: 2.2, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: C.ink, lineSpacing: 20, paraSpaceAfter: 8,
  });
  s.addText('Use it for the connective tissue between decisions.', {
    x: M + colW + 1.32, y: 5.14, w: colW - 0.84, h: 0.4, margin: 0,
    fontFace: BODY, fontSize: 12, italic: true, color: C.slate,
  });

  strip(s, 5.84, 'Three responses is a ceiling, not a target. If you need a fourth, you are looking at two decisions pretending to be one — split the scene.');
  s.addNotes(
    'The ceiling is the useful constraint here and it is worth defending rather than apologising for. Three ' +
    'plausible options is already hard to write; four is usually padding.\n\n' +
    'Two responses is completely legitimate. Say so — people assume they must fill all three slots and end up ' +
    'inventing a filler option, which is exactly what wrecks a scenario.\n\n' +
    'Text scenes are the under-used half. They are how you open ("here is the situation"), how you recover ' +
    '("that closed the conversation — try again") and how you land the point at the end. A scenario made ' +
    'entirely of dialogue feels relentless.'
  );
}

// ---- 13. Branching --------------------------------------------------------
{
  const s = contentSlide('Where each choice leads',
    'The Go to menu on every response. Four options, and the third one is what makes a scenario fair.');

  const gx = M, gw = 7.4;
  card(s, gx, 1.82, gw, 4.34);
  s.addText('THE SHAPE OF THIS AFTERNOON’S SCENARIO', {
    x: gx + 0.4, y: 2.04, w: gw - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 10.5, bold: true, color: C.slate, charSpacing: 1.6,
  });
  const nodes = [
    ['Scene 1', 'Dialogue — the opening line', 2.42, C.ember],
    ['Scene 2', 'Dialogue — “so what would help?”', 3.20, C.ember],
    ['Scene 4', 'Text — the ending', 3.98, C.ink],
    ['Scene 3', 'Text — the reset  →  Try again', 4.76, C.slate],
  ];
  nodes.forEach((n) => {
    s.addShape(pres.ShapeType.roundRect, {
      x: gx + 0.5, y: n[2], w: 2.1, h: 0.56, rectRadius: 0.08,
      fill: { color: n[3] }, line: { type: 'none' },
    });
    s.addText(n[0], {
      x: gx + 0.5, y: n[2], w: 2.1, h: 0.56, margin: 0, align: 'center', valign: 'middle',
      fontFace: BODY, fontSize: 13, bold: true, color: 'FFFFFF',
    });
    s.addText(n[1], {
      x: gx + 2.78, y: n[2], w: gw - 3.3, h: 0.56, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 12.5, color: C.slate,
    });
  });
  s.addText('A good answer moves down the chain. Anything else lands on Scene 3, which returns the learner to the decision — not to a telling-off.', {
    x: gx + 0.5, y: 5.42, w: gw - 1.0, h: 0.68, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 11.5, italic: true, color: C.slate, lineSpacing: 16,
  });

  const rx = gx + gw + 0.5, rw = CW - gw - 0.5;
  const opts = [
    ['Next content', 'The default. Straight on.'],
    ['A specific scene or content', 'The actual branching. Point anywhere.'],
    ['Try again', 'Back to the decision. Use it.'],
    ['End Scenario', 'Stop here, and hand back to the lesson.'],
  ];
  card(s, rx, 1.82, rw, 4.34);
  s.addText('THE GO TO MENU', {
    x: rx + 0.36, y: 2.04, w: rw - 0.72, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 10.5, bold: true, color: C.slate, charSpacing: 1.6,
  });
  opts.forEach((o, i) => {
    const y = 2.5 + i * 0.86;
    s.addText(o[0], {
      x: rx + 0.36, y, w: rw - 0.72, h: 0.34, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 15, bold: true, color: i === 2 ? C.ember : C.ink,
    });
    s.addText(o[1], {
      x: rx + 0.36, y: y + 0.32, w: rw - 0.72, h: 0.42, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12, color: C.slate, lineSpacing: 16,
    });
  });

  strip(s, 6.26, '“Try again” is the difference between a scenario and a trick question. It returns the learner to the decision instead of scolding them for it.', { h: 0.54, size: 12 });
  s.addNotes(
    'Walk the left diagram from the top. The point to land: the wrong turns do not dead-end and they do not ' +
    'punish. They loop.\n\n' +
    'Note what Scene 3 actually is — a Text scene whose only job is to explain the consequence and hand the ' +
    'decision back. That is the pattern worth copying, and it costs one scene.\n\n' +
    'On the right, the second option is where real branching lives, and it is the one people do not find. ' +
    'You can point any response at any content anywhere in the scenario, which is how you build a genuine ' +
    'fork rather than a straight line with commentary.\n\n' +
    'Warn them about the obvious failure: point everything at End Scenario and you have built a survey.'
  );
}

// ---- 14. Write it first ---------------------------------------------------
{
  const s = contentSlide('Write it before you open the editor',
    'Branching written live turns into a maze you cannot follow. Five steps, on paper, first.');
  const steps = [
    ['Name the ending', 'What does the learner do differently on Monday? Write that one sentence before anything else. Everything in the scenario either serves it or goes.'],
    ['Find the decision', 'One real moment where a competent person could plausibly go either way. If there is no such moment, you do not have a scenario — you have an explanation.'],
    ['Write three options a real person would pick', 'If one option is obviously stupid, it is not an option. It is filler, and it tells the learner the whole thing is a game.'],
    ['Write feedback as consequence', 'What happens next, in the character’s voice. Not a verdict on the learner. The character answers; the course does not grade.'],
    ['Give the wrong turn a way back', 'A Text scene that names the consequence, then Try again. A scenario without one is a trap, and traps teach people to stop trying.'],
  ];
  const gap = 0.24, w = (CW - gap * 4) / 5;
  steps.forEach((st, i) => {
    const x = M + i * (w + gap);
    card(s, x, 1.86, w, 4.0);
    s.addShape(pres.ShapeType.roundRect, {
      x: x + 0.24, y: 2.1, w: 0.42, h: 0.42, rectRadius: 0.1,
      fill: { color: C.ember }, line: { type: 'none' },
    });
    s.addText(String(i + 1), {
      x: x + 0.24, y: 2.1, w: 0.42, h: 0.42, margin: 0, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 14.5, bold: true, color: 'FFFFFF',
    });
    s.addText(st[0], {
      x: x + 0.24, y: 2.64, w: w - 0.48, h: 1.0, margin: 0, valign: 'top',
      fontFace: HEAD, fontSize: 15, bold: true, color: C.ink, lineSpacing: 19,
    });
    s.addText(st[1], {
      x: x + 0.24, y: 3.72, w: w - 0.48, h: 1.96, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 11.5, color: C.slate, lineSpacing: 16,
    });
  });
  strip(s, 5.98, 'Steps 1 to 3 happen on paper, away from Rise. Building a written scenario takes about ten minutes; writing one in the editor takes an afternoon and reads like it.', { dot: false });
  s.addNotes(
    'This is the slide to hand out. It is the method, and it is the thing that will still be useful to them ' +
    'in a year.\n\n' +
    'Step 1 is the one people skip and it is the one that saves the most time. Without a named ending you ' +
    'cannot tell which branches are worth writing.\n\n' +
    'Step 3 deserves a moment of honesty: writing three genuinely plausible options is hard, and it is where ' +
    'most published scenarios fail. The tell is an option nobody would ever choose.\n\n' +
    'If you have time, run this live: take a real situation from the room, get them to name the ending and ' +
    'the decision out loud, and write three options on the whiteboard together. Ten minutes, and it is the ' +
    'most useful ten minutes of the morning.'
  );
}

// ---- 15. Feedback that teaches --------------------------------------------
{
  const s = contentSlide('The character answers. The course does not grade.',
    'Scenario feedback is the hardest writing in the block, and the easiest to get wrong.');

  const colW = (CW - 0.6) / 2;
  const pairs = [
    ['Incorrect. Try again.', '“…Right. Okay. What do you want me to say?”'],
    ['That was the rude option.', '“Great — I’m heads-down on the launch, so I’ll crack on.”'],
    ['Correct! Well done.', '“Yeah. I got stuck on the pricing section and didn’t want to send something half-finished.”'],
  ];

  card(s, M, 1.82, colW, 4.24, { fill: 'FFF1EC', line: 'F3C9BC' });
  s.addText('A COURSE MARKING A LEARNER', {
    x: M + 0.4, y: 2.06, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.bad, charSpacing: 1.8,
  });
  card(s, M + colW + 0.6, 1.82, colW, 4.24, { fill: 'EEF7F2', line: 'C3E0D2' });
  s.addText('A PERSON REACTING', {
    x: M + colW + 1.0, y: 2.06, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ok, charSpacing: 1.8,
  });

  pairs.forEach((p, i) => {
    const y = 2.52 + i * 1.14;
    s.addText(p[0], {
      x: M + 0.4, y, w: colW - 0.8, h: 1.0, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 14, color: C.bad, lineSpacing: 20,
    });
    s.addText(p[1], {
      x: M + colW + 1.0, y, w: colW - 0.8, h: 1.0, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 14, color: '1F5E42', italic: true, lineSpacing: 20,
    });
  });

  strip(s, 6.2, 'If your feedback could be pasted into any scenario in any course, it is not feedback. It is a verdict with a costume on.', { h: 0.54, size: 12 });
  s.addNotes(
    'Read the left column aloud in a flat voice, then the right column in the character’s voice. The ' +
    'difference lands in the delivery in a way it does not on the page.\n\n' +
    'The third pair is the one that surprises people. “Correct! Well done.” feels harmless and it is still ' +
    'wrong here — it breaks the fiction and turns a conversation into a test. In the right-hand version the ' +
    'learner knows they did well because Priya opened up, which is exactly what happens in real life.\n\n' +
    'The bottom line is the test. Generic feedback is a verdict. Specific feedback is a consequence.'
  );
}

// ---- 16. Section three ----------------------------------------------------
sectionSlide('SECTION THREE', 'Watch me\nbuild one',
  'A blank course to a live link in forty-five minutes. Two lessons, a quiz, nothing skipped — including the parts that go wrong.')
  .addNotes(
    'Say the rule before you start: watch, do not build along. Building along means everyone falls behind at a ' +
    'different moment and the session becomes tech support. Their turn is the capstone.\n\n' +
    'Tell them what to watch for: the order of operations. Which decisions are cheap to change later and which ' +
    'are not. That is the thing a slide genuinely cannot teach.\n\n' +
    'Everything you need is in demo-script.md — the copy is written, so you never compose in front of the ' +
    'room. Make the three scripted mistakes; they are the best part.'
  );

// ---- 17. What you are about to watch --------------------------------------
{
  const s = contentSlide('What you are about to watch',
    'Nine blocks and four questions. Small enough to finish, real enough to matter.');

  const colW = (CW - 0.6) / 3;
  const lessons = [
    ['Lesson 1', 'Why feedback fails', ['Heading', 'Paragraph', 'Statement A', 'Numbered list', 'Button'], 'read'],
    ['Lesson 2', 'Try it', ['Paragraph with heading', 'Scenario — 3 scenes', 'Button stack'], 'do'],
    ['Quiz', 'Check yourself', ['Multiple Choice', 'Multiple Response', 'Fill in the Blank', 'Matching'], 'score'],
  ];
  lessons.forEach((l, i) => {
    const x = M + i * (colW + 0.3);
    card(s, x, 1.86, colW, 3.6, i === 2 ? { fill: C.ink, line: C.ink } : {});
    pill(s, x + 0.32, 2.12, 1.1, l[3], { fill: i === 2 ? C.ember : C.ink, size: 10 });
    s.addText(l[0], {
      x: x + 1.56, y: 2.08, w: colW - 1.9, h: 0.36, margin: 0, valign: 'middle',
      fontFace: BODY, fontSize: 11, bold: true, color: i === 2 ? C.slateLt : C.slate, charSpacing: 1.4,
    });
    s.addText(l[1], {
      x: x + 0.32, y: 2.6, w: colW - 0.64, h: 0.5, margin: 0, valign: 'top',
      fontFace: HEAD, fontSize: 19, bold: true, color: i === 2 ? 'FFFFFF' : C.ink,
    });
    s.addText(l[2].map((b, j) => ({
      text: b, options: { bullet: true, breakLine: j < l[2].length - 1 },
    })), {
      x: x + 0.32, y: 3.2, w: colW - 0.64, h: 2.0, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 13, color: i === 2 ? C.slateLt : C.slate,
      lineSpacing: 20, paraSpaceAfter: 6,
    });
  });

  card(s, M, 5.62, CW, 0.76);
  s.addText([
    { text: 'Watch the order, not the clicks.  ', options: { bold: true, color: C.ink } },
    { text: 'The scenario gets written before it gets built. The quiz questions get written while the content is still fresh. Publishing happens last and takes four minutes.', options: { color: C.slate } },
  ], {
    x: M + 0.38, y: 5.62, w: CW - 0.76, h: 0.76, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 13, lineSpacing: 18,
  });
  s.addNotes(
    'Put this up before you open Rise and leave it up for a minute. It is the map they will hold in their ' +
    'heads for the next forty-five minutes.\n\n' +
    'Say the numbers out loud — nine blocks, four questions. The most common capstone failure is scope, and ' +
    'this is the anchor that prevents it. Repeat the numbers when you hand out the brief.\n\n' +
    'The three-word tags matter more than they look: read, do, score. That is the shape of almost every good ' +
    'small course, and it is the shape their capstone should take.'
  );
}

// ---- 18. Four question types ----------------------------------------------
{
  const s = contentSlide('The four question types',
    'Everything a Rise quiz can ask. A fifth block, Draw from Question Bank, serves a random subset from a bank you build.');
  const qs = [
    ['Multiple Choice', 'One right answer. The default, and the one most often ruined by three obviously wrong distractors.'],
    ['Multiple Response', 'Several right answers. Say “select all that apply” in the stem or people pick one and stop.'],
    ['Fill in the Blank', 'A typed answer, matched as text. Add every spelling and synonym you will accept, or you will fail people who are right.'],
    ['Matching', 'Pairs. Excellent for term to definition. Useless if the pairs can be guessed from length or wording.'],
  ];
  const gap = 0.3, w = (CW - gap * 3) / 4;
  const imgW = w - 0.7, imgH = imgW / (294 / 204);
  qs.forEach((q, i) => {
    const x = M + i * (w + gap);
    card(s, x, 1.84, w, 4.1);
    s.addImage({ path: imgOf(q[0]), x: x + 0.35, y: 2.04, w: imgW, h: imgH });
    s.addText(q[0], {
      x: x + 0.2, y: 2.12 + imgH, w: w - 0.4, h: 0.5, margin: 0, align: 'center', valign: 'middle',
      fontFace: HEAD, fontSize: 15, bold: true, color: C.ink,
    });
    s.addText(q[1], {
      x: x + 0.3, y: 2.72 + imgH, w: w - 0.6, h: 1.5, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 11.5, color: C.slate, lineSpacing: 16,
    });
  });
  strip(s, 6.06, 'Write feedback on every question, right and wrong. An unexplained wrong answer teaches nothing and a bare “Correct” wastes the one moment you had their attention.');
  s.addNotes(
    'Four types, and the room will use Multiple Choice for everything unless you push. Push.\n\n' +
    'Matching is the underrated one — it is the fastest way to check that somebody can tell four similar ' +
    'things apart, and it is much harder to guess than a four-option MCQ.\n\n' +
    'Fill in the Blank is the one that generates support tickets. It matches text. If you accept “behaviour” ' +
    'but not “behavior”, you will fail half of a global audience for spelling. Show the accepted-answers field ' +
    'in the demo and add a second spelling in front of them.\n\n' +
    'Mention the question bank in one sentence and move on — it matters for real assessments and not for a ' +
    'four-question capstone.'
  );
}

// ---- 19. Quiz settings ----------------------------------------------------
{
  const s = contentSlide('Quiz settings, and what to actually do with them',
    'Defaults are choices somebody else made. Two of these are worth changing every single time.');
  table(s, [
    { label: 'Setting', w: 4.0, bold: true },
    { label: 'Default', w: 2.0, color: C.slate },
    { label: 'What to do', w: CW - 6.0, color: C.slate },
  ], [
    ['Passing Score', '80%', 'Do the arithmetic for your question count. 80% of four questions is four out of four.'],
    ['Shuffle Answer Choices', 'Off', 'Turn it on. It costs nothing and stops people remembering “it was the third one”.'],
    ['Reveal Answers', '—', 'Show incorrect answers with feedback, or nobody ever reads the feedback you wrote.'],
    ['Quiz Retries', 'Unlimited', 'Limit it only when the score has consequences. Fail every attempt and the learner is stuck.'],
    ['Randomize Question Order', 'Off', 'On for a real assessment. Off for four questions — it only makes support harder.'],
    ['Require Passing Score to Continue', 'Off', 'This is the gate. Know it exists before somebody asks you to lock the course.'],
    ['Skip Ahead If Passed', 'Off', 'Turns the quiz into a pre-test: pass it and the content you already know is skipped.'],
    ['Timer', 'Off, 10 min', 'Almost never. Time pressure tests typing speed, not knowledge.'],
  ], { rowH: 0.5, size: 12 });

  s.addText('Settings live behind the Settings button in the top right of the quiz editor — not in the block panel.', {
    x: M, y: 6.4, w: CW, h: 0.36, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 11.5, italic: true, color: C.slate,
  });
  s.addNotes(
    'Open the settings panel live and change the two that matter — passing score and shuffle — so they see ' +
    'where it lives. People genuinely do not find this panel on their own.\n\n' +
    'The passing-score arithmetic is the best single moment in this section. Ask the room what 80% of four ' +
    'questions is, wait for someone to work out that it means four out of four, then fix it. It lands far ' +
    'harder than being told.\n\n' +
    'Skip Ahead If Passed is recent and worth thirty seconds — for an audience who already know some of the ' +
    'material it is genuinely respectful, and it is off by default.\n\n' +
    'Retries deserves a warning: limit the attempts and a learner who fails them all cannot complete the ' +
    'course at all. That is a support call, not a design choice.'
  );
}

// ---- 20. Knowledge check vs quiz ------------------------------------------
{
  const s = contentSlide('Knowledge check or quiz?',
    'They look almost identical to a learner. Only one of them can ever report a score.');
  const colW = (CW - 0.9) / 2;

  card(s, M, 1.84, colW, 3.5);
  s.addText('KNOWLEDGE CHECK BLOCKS', {
    x: M + 0.4, y: 2.08, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.slate, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Question types placed as blocks, inside a normal lesson', options: { bullet: true, breakLine: true } },
    { text: 'Practice — checking understanding as you go', options: { bullet: true, breakLine: true } },
    { text: 'No score, no pass mark, no record', options: { bullet: true, breakLine: true } },
    { text: 'Use them freely; they cost the learner nothing', options: { bullet: true } },
  ], {
    x: M + 0.4, y: 2.5, w: colW - 0.8, h: 2.6, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: C.ink, lineSpacing: 20, paraSpaceAfter: 8,
  });

  card(s, M + colW + 0.9, 1.84, colW, 3.5, { fill: C.ink, line: C.ink });
  s.addText('A QUIZ LESSON', {
    x: M + colW + 1.3, y: 2.08, w: colW - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText([
    { text: 'A lesson type of its own, added from the outline', options: { bullet: true, breakLine: true } },
    { text: 'Scored, with a passing score and retries', options: { bullet: true, breakLine: true } },
    { text: 'The only thing an LMS can read as a result', options: { bullet: true, breakLine: true } },
    { text: 'Microlearning cannot hold one — you need a Course', options: { bullet: true } },
  ], {
    x: M + colW + 1.3, y: 2.5, w: colW - 0.8, h: 2.6, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13.5, color: 'FFFFFF', lineSpacing: 20, paraSpaceAfter: 8,
  });

  card(s, M, 5.56, CW, 0.86, { fill: C.emberPale, line: C.emberDim, flat: true });
  s.addText('Build a course entirely of knowledge checks and your LMS report will be empty.', {
    x: M + 0.4, y: 5.56, w: CW - 0.8, h: 0.86, margin: 0, valign: 'middle', align: 'center',
    fontFace: HEAD, fontSize: 19, bold: true, color: C.emberInk,
  });
  s.addNotes(
    'This slide prevents a specific, expensive mistake that happens in real organisations — a course full of ' +
    'knowledge checks, launched, and then a month of confusion about why the LMS shows nothing.\n\n' +
    'The distinction is not about difficulty or question type. It is exactly one thing: a quiz is a lesson and ' +
    'it reports; blocks do not.\n\n' +
    'Tie it back to Day 2. This is the second time Microlearning has bitten — it cannot hold a quiz, which is ' +
    'why the demo course was created as a Course.\n\n' +
    'Ask the room: "if your manager asks who completed the training, which of these two answers the question?" ' +
    'Let them say it.'
  );
}

// ---- 21. Section four -----------------------------------------------------
sectionSlide('SECTION FOUR', 'Four ways out',
  'Preview like a learner first. Then pick the door that matches who is taking this and what you need to know about them.')
  .addNotes(
    'Thirty minutes. The single decision that matters here is the door, and it is driven by one question: do ' +
    'you need to know who completed it?\n\n' +
    'Resist going deep on SCORM versions unless the room is publishing to a corporate LMS — in that case, ' +
    'spend the time and skip a scenario example this morning instead.'
  );

// ---- 22. Preview ----------------------------------------------------------
{
  const s = contentSlide('Preview like a learner, not like an author',
    'Rise reflows rather than laying out fixed pages, so you never design for a device — but you do have to look.');

  const views = [
    ['Desktop', 3.0, 1.9, 'the one you built in'],
    ['Tablet\nlandscape', 2.2, 1.5, ''],
    ['Tablet\nportrait', 1.5, 2.1, ''],
    ['Mobile\nlandscape', 1.7, 1.0, ''],
    ['Mobile\nportrait', 1.0, 1.85, 'where most of them are'],
  ];
  const totalW = views.reduce((a, v) => a + v[1], 0) + 0.5 * (views.length - 1);
  let x = M + (CW - totalW) / 2;
  const baseY = 4.3;
  views.forEach((v, i) => {
    const hot = i === 4;
    s.addShape(pres.ShapeType.roundRect, {
      x, y: baseY - v[2], w: v[1], h: v[2], rectRadius: 0.08,
      fill: { color: hot ? C.ember : C.paper },
      line: { color: hot ? C.ember : C.slateLt, width: hot ? 2 : 1.25 },
      shadow: shadow({ blur: 10, opacity: 0.2 }),
    });
    s.addText(v[0], {
      x: x - 0.2, y: baseY + 0.12, w: v[1] + 0.4, h: 0.6, margin: 0, align: 'center', valign: 'top',
      fontFace: BODY, fontSize: 12, bold: true, color: hot ? C.ember : C.ink, lineSpacing: 15,
    });
    if (v[3]) {
      s.addText(v[3], {
        x: x - 0.35, y: baseY + 0.74, w: v[1] + 0.7, h: 0.5, margin: 0, align: 'center', valign: 'top',
        fontFace: BODY, fontSize: 10.5, italic: true, color: C.slate, lineSpacing: 14,
      });
    }
    x += v[1] + 0.5;
  });

  strip(s, 5.62, 'Land on mobile portrait and stay there. Button labels wrap, scenario text crowds the character, and labelled-graphic markers collide — all of it invisible on your desktop.');
  s.addText('Preview is beside Publish, top right. Five views, one click each.', {
    x: M, y: 6.36, w: CW, h: 0.36, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 11.5, italic: true, color: C.slate,
  });
  s.addNotes(
    'Do this live on the demo course, all five views, and let them see you find something to fix. If nothing ' +
    'is broken, say so — but look properly, because normally something is.\n\n' +
    'The conceptual point is the contrast with PowerPoint and Storyline: Rise reflows like a website, so you ' +
    'are not designing per device. That is why previewing is checking, not designing.\n\n' +
    'Mobile portrait is where the defects are and it is the view everybody skips. Make the room promise to ' +
    'check it before Monday — it is on the capstone checklist for exactly this reason.'
  );
}

// ---- 23. Four ways out ----------------------------------------------------
{
  const s = contentSlide('Four ways out of Rise',
    'One question decides it: do you need to know who completed this?');
  table(s, [
    { label: 'Path', w: 2.6, bold: true, color: C.ember },
    { label: 'What you get', w: 3.5, color: C.ink },
    { label: 'What it can tell you', w: 3.2, color: C.slate },
    { label: 'Use it when', w: CW - 9.3, color: C.slate },
  ], [
    ['Quick Share', 'One public link, immediately', 'Views and guestbook entries', 'It needs to be in someone’s hands today'],
    ['Reach 360', 'Articulate’s own LMS', 'Full learner reporting, certificates', 'Your learners are outside the company'],
    ['LMS export', 'A package for your LMS', 'Whatever the standard supports', 'Someone needs a completion record'],
    ['Web / PDF', 'A zip you host, or a document', 'Nothing at all', 'You own the server, or you need a handout'],
  ], { rowH: 0.66 });

  card(s, M, 5.06, CW, 0.82);
  s.addText([
    { text: 'PDF is not a course.  ', options: { bold: true, color: C.ink } },
    { text: 'Every interaction becomes a static picture — your scenario turns into a screenshot. It is a useful reference handout and nothing more.', options: { color: C.slate } },
  ], {
    x: M + 0.38, y: 5.06, w: CW - 0.76, h: 0.82, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 13, lineSpacing: 18,
  });
  strip(s, 6.04, 'Review 360 is a fifth door, and it is not for learners. It is how you get in-context comments from stakeholders before any of this happens.', { h: 0.54, size: 12 });
  s.addNotes(
    'Run this as one question, asked out loud: do you need to know who completed it? No means Quick Share or ' +
    'Web. Yes means an LMS — yours or Reach.\n\n' +
    'Quick Share is the one that changes how teams work. A link in a channel in thirty seconds is a genuinely ' +
    'different thing from a two-week LMS upload cycle, and most of what your organisation publishes does not ' +
    'need tracking at all.\n\n' +
    'Reach 360 is worth real time only if your audience is external or deskless — contractors, partners, ' +
    'franchisees — people who will never have a corporate LMS account.\n\n' +
    'Mention Review 360 last and briefly. It is the step before publishing, and if your organisation reviews ' +
    'content by emailing Word documents around, it is the single biggest improvement available to them.'
  );
}

// ---- 24. LMS export -------------------------------------------------------
{
  const s = contentSlide('Exporting for an LMS',
    'Publish › LMS. Three decisions and a zip file.');
  const cols = [
    ['1 · Standard', [
      'xAPI (Tin Can API)', 'SCORM 2004', 'SCORM 1.2', 'AICC', 'cmi5',
    ], 'SCORM 1.2 when you are unsure. It is the one that works everywhere, and it is what most LMS teams will ask for.'],
    ['2 · Tracking', [
      'Completion percentage', 'Storyline block', 'Quiz result (courses only)',
    ], 'Quiz result is the one people want and the one that quietly requires an actual quiz lesson in the course.'],
    ['3 · Settings', [
      'Exit Course Link', 'Hide Cover Page', 'Allow Usage Data Collection', 'Reset Learner Progress', 'Hide LMS Interface', 'Only Load in LMS',
    ], 'The last three sit under “More settings”. Leave them alone unless your LMS administrator asks.'],
  ];
  const gap = 0.4, w = (CW - gap * 2) / 3;
  cols.forEach((c, i) => {
    const x = M + i * (w + gap);
    card(s, x, 1.84, w, 3.86);
    s.addText(c[0], {
      x: x + 0.36, y: 2.06, w: w - 0.72, h: 0.36, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 17, bold: true, color: C.ink,
    });
    s.addText(c[1].map((t, j) => ({
      text: t, options: { bullet: true, breakLine: j < c[1].length - 1 },
    })), {
      x: x + 0.36, y: 2.52, w: w - 0.72, h: 2.0, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12.5, color: C.ink, lineSpacing: 19, paraSpaceAfter: 4,
    });
    s.addText(c[2], {
      x: x + 0.36, y: 4.6, w: w - 0.72, h: 1.0, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 11.5, italic: true, color: C.slate, lineSpacing: 16,
    });
  });

  strip(s, 5.86, 'Only a quiz reports a score. xAPI and cmi5 launch from indexapi.html; a Web export launches from index.html. LMS and Web publishing need a standard Articulate 360 licence — Quick Share and PDF do not.', { h: 0.72 });
  s.addNotes(
    'Do this live and let the download finish on screen, so they see that the output is just a zip. Do not ' +
    'attempt an LMS upload in front of the room.\n\n' +
    'If nobody here publishes to a corporate LMS, spend two minutes and move on. If everybody does, this is ' +
    'the most valuable slide of the afternoon and you should take questions properly.\n\n' +
    'The launch-file detail is the thing LMS administrators ask for, and knowing it makes your learners look ' +
    'competent in a conversation they will definitely have.\n\n' +
    'The licence note heads off a support question. If LMS and Web are greyed out on somebody’s account, that ' +
    'is the reason.'
  );
}

// ---- 25. Quick Share ------------------------------------------------------
{
  const s = contentSlide('Quick Share, in detail',
    'One link, no LMS, no learner accounts. The fastest useful thing in the tool.');
  const items = [
    ['Password', 'Optional. Turn it on and nobody reaches the training without it.'],
    ['Guestbook', 'On by default. It asks every learner for a name and an email before they start — turn it off for a casual share.'],
    ['Live updates', 'Edit the course in Rise and the link shows the new version. No re-publishing, no new link.'],
    ['Engagement', 'How many individual learners, how many views, exportable as CSV. The view count includes repeat visits by the same person.'],
  ];
  const gap = 0.3, w = (CW - gap * 3) / 4;
  items.forEach((it, i) => {
    const x = M + i * (w + gap);
    card(s, x, 1.9, w, 3.5);
    s.addText(it[0], {
      x: x + 0.32, y: 2.16, w: w - 0.64, h: 0.44, margin: 0, valign: 'middle',
      fontFace: HEAD, fontSize: 17, bold: true, color: C.ink,
    });
    s.addShape(pres.ShapeType.rect, {
      x: x + 0.32, y: 2.68, w: 0.5, h: 0.035, fill: { color: C.ember }, line: { type: 'none' },
    });
    s.addText(it[1], {
      x: x + 0.32, y: 2.86, w: w - 0.64, h: 2.3, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12.5, color: C.slate, lineSpacing: 18,
    });
  });

  card(s, M, 5.58, CW, 0.8, { fill: C.ink, line: C.ink });
  s.addText([
    { text: 'What it is not:  ', options: { bold: true, color: C.ember } },
    { text: 'a tracking system. You get views and guestbook entries, never per-learner quiz scores. The link stays alive until you delete the training or the subscription lapses.', options: { color: 'FFFFFF' } },
  ], {
    x: M + 0.4, y: 5.58, w: CW - 0.8, h: 0.8, margin: 0, valign: 'middle',
    fontFace: BODY, fontSize: 13, lineSpacing: 18,
  });
  s.addNotes(
    'Publish the demo course for real here and paste the link into the room’s channel. Watching thirty people ' +
    'open your course on their phones is the most persuasive thing that happens all day.\n\n' +
    'Then do the live-updates trick: with the link open in a second tab, change a word in lesson 1 and ' +
    'refresh. That single move sells the feature better than any slide.\n\n' +
    'The guestbook default surprises people — it is on, and it asks every learner for a name and email. Fine ' +
    'for a stakeholder review, wrong for a casual share.\n\n' +
    'Be straight about the ceiling. People try to use Quick Share as compliance tracking and it is not that. ' +
    'If somebody needs to prove who completed training, they need an LMS.'
  );
}

// ---- 26. Section five -----------------------------------------------------
sectionSlide('SECTION FIVE', 'Your turn',
  'Same shape, your subject, published by Monday morning. Fifteen minutes on this — the detail is all in the brief.')
  .addNotes(
    'Fifteen minutes, and most of it should be questions. The brief is written down, so do not read it out.\n\n' +
    'Hand out capstone-brief.md — printed if you can. Walk the checklist on the next two slides, then take ' +
    'questions until they stop.\n\n' +
    'The most important thing you can say in this section: pick a topic you already know. The capstone tests ' +
    'the tool, not their subject expertise, and the people who choose something they have to research are the ' +
    'people who do not finish.'
  );

// ---- 27. The capstone -----------------------------------------------------
{
  const s = contentSlide('The capstone',
    'Two lessons, one quiz, nine to twelve blocks — the same shape you just watched, on a topic from your own work.');

  const colW = (CW - 0.6) / 3;
  const groups = [
    ['Structure', [
      'A Course, not Microlearning',
      'A title that reads as a promise',
      'Two lessons plus a quiz',
      '9–12 blocks in total',
    ]],
    ['Blocks', [
      'Four static families from Day 2',
      'One Day 3 interactive block',
      'One Scenario — 3+ scenes, a wrong branch, a Try again',
      'One Button or Button stack, labelled properly',
    ]],
    ['Finish', [
      'Four questions, three question types',
      'Feedback on every question',
      'Previewed on mobile portrait, and fixed',
      'A Quick Share link and a SCORM 1.2 export',
    ]],
  ];
  groups.forEach((g, i) => {
    const x = M + i * (colW + 0.3);
    card(s, x, 1.8, colW, 3.5);
    pill(s, x + 0.32, 2.06, 1.5, g[0], { fill: C.ink, size: 10.5 });
    s.addText(g[1].map((t, j) => ({
      text: t, options: { bullet: true, breakLine: j < g[1].length - 1 },
    })), {
      x: x + 0.32, y: 2.6, w: colW - 0.64, h: 2.5, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12.5, color: C.ink, lineSpacing: 19, paraSpaceAfter: 8,
    });
  });

  const dates = [
    ['Friday', 'Clinic — bring your outline, not your course'],
    ['Weekend', 'About three hours. Past four means your topic is too big'],
    ['Monday 9:00', 'Link, SCORM zip and three lines, in the channel'],
  ];
  const dw = (CW - 0.6) / 3;
  dates.forEach((d, i) => {
    const x = M + i * (dw + 0.3);
    card(s, x, 5.46, dw, 0.92, i === 2 ? { fill: C.ink, line: C.ink } : { fill: C.paper });
    s.addText(d[0], {
      x: x + 0.32, y: 5.58, w: dw - 0.64, h: 0.3, margin: 0,
      fontFace: BODY, fontSize: 11, bold: true, color: i === 2 ? C.ember : C.slate, charSpacing: 1.6,
    });
    s.addText(d[1], {
      x: x + 0.32, y: 5.88, w: dw - 0.64, h: 0.42, margin: 0, valign: 'top',
      fontFace: BODY, fontSize: 12, color: i === 2 ? 'FFFFFF' : C.ink, lineSpacing: 16,
    });
  });
  s.addNotes(
    'Do not read the columns. They have the brief in their hands; this slide exists so the room is looking at ' +
    'the same thing while you talk.\n\n' +
    'Say the three things that are not on the slide:\n' +
    '1. Pick a topic you can already explain in ten minutes. No research.\n' +
    '2. Write the scenario first, on paper. It is the hardest part and it changes the lessons.\n' +
    '3. Small and finished beats big and unfinished, every time.\n\n' +
    'The three-hour budget is a real number, not encouragement. Anyone past four hours has chosen too big a ' +
    'topic and should cut a lesson rather than lose a weekend.\n\n' +
    'Point at the clinic. Help is cheapest on Friday, when the outline is still on paper and changing it is free.'
  );
}

// ---- 28. Ready means this -------------------------------------------------
{
  const s = contentSlide('“Ready” means this',
    'Nothing here is about polish. All of it is about whether the course works for a learner who is not you.');
  table(s, [
    { label: '', w: 2.7, bold: true },
    { label: 'Ready', w: 4.5, color: C.ok },
    { label: 'Not ready yet', w: CW - 7.2, color: C.bad },
  ], [
    ['Title', 'Says what the learner will be able to do', 'Names a module, a system or a department'],
    ['Lesson 1', 'Under half text, and something to look at', 'A wall of paragraphs'],
    ['Scenario', 'Every option is one a real person might pick', 'One right answer and two silly ones'],
    ['Feedback', 'Explains the consequence, in the character’s voice', 'Tells the learner they were wrong'],
    ['Try again', 'Returns them to the decision', 'Missing, or dumps them at the end'],
    ['Buttons', 'The label says what happens', '“Click here”, “Learn more”, “Read on”'],
    ['Quiz', 'Feedback on every question, a reachable pass mark', 'Bare right and wrong, 80% on four questions'],
    ['The link', 'Opens for somebody else, first time', 'Never tested outside your own browser'],
  ], { top: 1.78, rowH: 0.47, size: 12 });

  strip(s, 6.12, 'Submit three lines with it: who it is for, what you would fix with another hour, and what fought you hardest in Rise. Monday’s session is built around those.', { h: 0.56, size: 12 });
  s.addNotes(
    'This is the marking scheme and there is nothing hidden in it. Say that plainly — people build better ' +
    'things when they know exactly what is being looked at.\n\n' +
    'The bottom row is the one that catches people. A Quick Share link that works in the browser you are ' +
    'signed into is not evidence of anything. Private window, or ask a colleague.\n\n' +
    'The three lines are not padding. The third one — what fought you hardest — is how you find out what this ' +
    'week failed to teach, and it shapes what you do on Monday.'
  );
}

// ---- 29. Close ------------------------------------------------------------
{
  const s = darkSlide();
  s.addText('BEFORE YOU GO', {
    x: M, y: 1.9, w: 6, h: 0.32, margin: 0,
    fontFace: BODY, fontSize: 12, bold: true, color: C.ember, charSpacing: 2.4,
  });
  s.addText('Write the scenario\nbefore you open\nthe editor.', {
    x: M, y: 2.34, w: 7.4, h: 2.4, margin: 0, valign: 'top',
    fontFace: HEAD, fontSize: 38, bold: true, color: 'FFFFFF', lineSpacing: 48,
  });
  s.addText('If you take one thing from today, take that. It is ten minutes on paper against an afternoon in a maze.', {
    x: M, y: 4.86, w: 6.8, h: 0.8, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 14.5, color: C.slateLt, lineSpacing: 22,
  });

  const rx = 8.5, rw = W - rx - M;
  card(s, rx, 1.9, rw, 3.9, { fill: C.inkSoft, line: '2A3242' });
  s.addText('TOMORROW — DAY 5', {
    x: rx + 0.4, y: 2.16, w: rw - 0.8, h: 0.3, margin: 0,
    fontFace: BODY, fontSize: 11, bold: true, color: C.ember, charSpacing: 1.8,
  });
  s.addText([
    { text: 'Theme — colours, fonts, your logo', options: { bullet: true, breakLine: true } },
    { text: 'Navigation — free against restricted, and when each is right', options: { bullet: true, breakLine: true } },
    { text: 'Accessibility, properly', options: { bullet: true, breakLine: true } },
    { text: 'Review 360 — comments from stakeholders in context', options: { bullet: true, breakLine: true } },
    { text: 'A clinic for your capstone. Bring the outline.', options: { bullet: true } },
  ], {
    x: rx + 0.4, y: 2.62, w: rw - 0.8, h: 3.0, margin: 0, valign: 'top',
    fontFace: BODY, fontSize: 13, color: 'FFFFFF', lineSpacing: 20, paraSpaceAfter: 9,
  });

  s.addText('Brief: capstone-brief.md   ·   Due Monday, 9:00 a.m.', {
    x: M, y: 6.1, w: CW, h: 0.36, margin: 0,
    fontFace: BODY, fontSize: 12.5, italic: true, color: C.slateLt,
  });
  footer(s, true);
  s.addNotes(
    'End on the sentence, not on the logistics. It is the one piece of craft from today that transfers to ' +
    'every authoring tool they will ever use.\n\n' +
    'Check three things before anybody leaves: they have the brief, they have picked a topic, and they know ' +
    'Friday’s clinic is for outlines rather than finished courses.\n\n' +
    'If the room struggled today, it was almost certainly with scenario feedback — writing consequence rather ' +
    'than verdict. Note who, and use the Friday clinic on them.'
  );
}

pres.writeFile({ fileName: OUT }).then(() => {
  console.log('wrote', OUT, '·', slideNo, 'slides');
});
