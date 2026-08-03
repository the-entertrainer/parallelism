const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, ImageRun,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, PageBreak,
  PageOrientation, Header, Footer, PageNumber, convertInchesToTwip
} = require('docx');

const MOD = process.argv[2];
const DATA = JSON.parse(fs.readFileSync(`findings_${MOD}.json`, 'utf8'));
const SLICES = JSON.parse(fs.readFileSync('data/slices.json', 'utf8'));
const SHOTS = Object.fromEntries(
  JSON.parse(fs.readFileSync(`data/${MOD}_shots.json`, 'utf8')).map(r => [r.id, r]));

const IMG_W = 660;               // px @96dpi -> ~6.9 in
const FONT = 'Arial';
const RED = 'C00000';
const NAVY = '003087';
const GREY = '595959';

const sz = (pt) => pt * 2;       // docx half-points

function p(text, opts = {}) {
  return new Paragraph({
    alignment: opts.align,
    spacing: { before: opts.before ?? 0, after: opts.after ?? 100 },
    indent: opts.indent,
    border: opts.border,
    shading: opts.shading,
    children: [new TextRun({
      text, font: FONT, size: sz(opts.size ?? 10.5),
      bold: opts.bold, italics: opts.italics, color: opts.color
    })]
  });
}

function runsPara(runs, opts = {}) {
  return new Paragraph({
    alignment: opts.align,
    spacing: { before: opts.before ?? 0, after: opts.after ?? 100 },
    indent: opts.indent,
    shading: opts.shading,
    border: opts.border,
    children: runs.map(r => new TextRun({
      text: r.t, font: FONT, size: sz(r.size ?? opts.size ?? 10.5),
      bold: r.bold, italics: r.italics, color: r.color
    }))
  });
}

function h(text, level, color) {
  return new Paragraph({
    heading: level,
    spacing: { before: level === HeadingLevel.HEADING_1 ? 320 : 240, after: 140 },
    children: [new TextRun({ text, font: FONT, bold: true, color: color || NAVY,
      size: sz(level === HeadingLevel.HEADING_1 ? 15 : level === HeadingLevel.HEADING_2 ? 12.5 : 11.5) })]
  });
}

function rule() {
  return new Paragraph({
    spacing: { before: 60, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D9D9D9' } },
    children: [new TextRun({ text: '', font: FONT, size: sz(2) })]
  });
}


// simple PNG dimension reader (IHDR)
function pngSize(file) {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20), buf: b };
}

function imageParas(screenKey) {
  const files = SLICES[screenKey];
  if (!files) return [p(`[screenshot missing: ${screenKey}]`, { italics: true, color: RED })];
  return files.map((f, i) => {
    const { w, h: ih, buf } = pngSize(f);
    const scale = IMG_W / w;
    return new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 60, after: files.length > 1 && i < files.length - 1 ? 40 : 140 },
      children: [new ImageRun({ data: buf, type: 'png',
        transformation: { width: Math.round(w * scale), height: Math.round(ih * scale) } })]
    });
  });
}

function infoTable(rows) {
  return new Table({
    columnWidths: [2200, 7400],
    width: { size: 9600, type: WidthType.DXA },
    rows: rows.map(([k, v]) => new TableRow({
      children: [
        new TableCell({
          width: { size: 2200, type: WidthType.DXA },
          shading: { type: ShadingType.CLEAR, fill: 'F2F2F2' },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [p(k, { bold: true, size: 10, after: 0 })]
        }),
        new TableCell({
          width: { size: 7400, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [p(v, { size: 10, after: 0 })]
        })
      ]
    }))
  });
}

// ---------------------------------------------------------------- build
const children = [];

// Cover
children.push(new Paragraph({ spacing: { before: 1400, after: 0 }, children: [
  new TextRun({ text: 'CONTENT QUALITY ASSURANCE AUDIT', font: FONT, bold: true, size: sz(12), color: GREY })
] }));
children.push(new Paragraph({ spacing: { before: 120, after: 60 }, children: [
  new TextRun({ text: DATA.title, font: FONT, bold: true, size: sz(28), color: NAVY })
] }));
children.push(new Paragraph({ spacing: { after: 500 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY } },
  children: [new TextRun({ text: 'Articulate Rise 360 module — editorial and design review', font: FONT, size: sz(12), color: GREY })]
}));

const totalItems = DATA.screens.reduce((n, s) => n + s.items.length, 0);
children.push(infoTable([
  ['Module', DATA.title],
  ['Review link', DATA.review_url],
  ['Review date', new Date().toISOString().slice(0, 10)],
  ['Screens with findings', String(DATA.screens.length)],
  ['Numbered findings', String(totalItems)],
  ['Additional observations', String(DATA.other.length)],
  ['Sources of truth', 'PayPal Style Guide (PayPal_Style_Guide_Final.pdf) and standard English usage. No storyboard or source content document was supplied, so content accuracy against source could not be checked; items needing source verification are called out explicitly.'],
  ['Method', 'Every lesson was opened in a desktop browser at 1400 px wide and read top to bottom. Every accordion, tab, flashcard and quiz question was opened and answered. Colours were sampled from the rendered page.']
]));

children.push(new Paragraph({ children: [new PageBreak()] }));

// How to read
children.push(h('How to read this report', HeadingLevel.HEADING_1));
children.push(p('Findings are grouped by the screen they appear on, in the order a learner meets them. Each screen shows a real screenshot of the module with every issue circled and numbered in red. Directly below the screenshot, each number has a comment and a reason.'));
children.push(p('The comment is written so it can be copied straight into a review tool as-is. The reason is written in plain English.'));
children.push(p('Nothing in this report has been posted to the Rise 360 review link.', { bold: true }));
children.push(rule());

// Summary by lesson
children.push(h('Findings by lesson', HeadingLevel.HEADING_1));
const byLesson = {};
DATA.screens.forEach(s => { byLesson[s.lesson] = (byLesson[s.lesson] || 0) + s.items.length; });
children.push(new Table({
  columnWidths: [7600, 2000],
  width: { size: 9600, type: WidthType.DXA },
  rows: [
    new TableRow({ children: [
      new TableCell({ width: { size: 7600, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: NAVY },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [p('Lesson', { bold: true, color: 'FFFFFF', size: 10, after: 0 })] }),
      new TableCell({ width: { size: 2000, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: NAVY },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [p('Findings', { bold: true, color: 'FFFFFF', size: 10, after: 0, align: AlignmentType.CENTER })] })
    ] }),
    ...Object.entries(byLesson).map(([l, n]) => new TableRow({ children: [
      new TableCell({ width: { size: 7600, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 },
        children: [p(l, { size: 10, after: 0 })] }),
      new TableCell({ width: { size: 2000, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 },
        children: [p(String(n), { size: 10, after: 0, align: AlignmentType.CENTER })] })
    ] }))
  ]
}));
children.push(new Paragraph({ spacing: { before: 200 }, children: [new PageBreak()] }));

// Screens
let currentLesson = null;
DATA.screens.forEach((sc, si) => {
  if (sc.lesson !== currentLesson) {
    if (currentLesson !== null) children.push(new Paragraph({ children: [new PageBreak()] }));
    currentLesson = sc.lesson;
    children.push(h(sc.lesson, HeadingLevel.HEADING_1));
  }
  children.push(h(`${sc.id}  ${sc.screen}`, HeadingLevel.HEADING_2));
  imageParas(sc.id.replace(/\./g, '_')).forEach(x => children.push(x));

  // Badges on the screenshot are numbered top-to-bottom; present the comments in the
  // same order so number 1 in the list is number 1 in the image.
  const ord = (SHOTS[sc.id] || {}).order;
  const ordered = (ord && ord.length === sc.items.length)
    ? ord.map(n => sc.items[n - 1]).filter(Boolean)
    : sc.items;
  ordered.forEach((it, i) => {
    children.push(runsPara(
      [{ t: `${i + 1}. `, bold: true, color: RED }, { t: it.comment }],
      { after: 60, indent: { left: 200 } }
    ));
    children.push(runsPara(
      [{ t: 'Reason: ', bold: true }, { t: it.reason }],
      { after: 200, indent: { left: 200 } }
    ));
  });
  if (si < DATA.screens.length - 1) children.push(rule());
});

// Other observations
children.push(new Paragraph({ children: [new PageBreak()] }));
children.push(h('Other observations', HeadingLevel.HEADING_1));
children.push(p('These items sit outside the numbered, screen-by-screen findings. They are module-wide patterns, instructional-design notes, or points that need someone else to confirm a fact.', { italics: true, color: GREY }));
DATA.other.forEach((o, i) => {
  children.push(runsPara([{ t: `${i + 1}. `, bold: true, color: RED }, { t: o.title, bold: true }], { after: 60 }));
  children.push(p(o.text, { after: 200, indent: { left: 200 } }));
});

const doc = new Document({
  creator: 'CQA review',
  title: `CQA audit — ${DATA.title}`,
  styles: { default: { document: { run: { font: FONT, size: sz(10.5) } } } },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840, orientation: PageOrientation.PORTRAIT },
        margin: { top: convertInchesToTwip(0.7), bottom: convertInchesToTwip(0.7),
                  left: convertInchesToTwip(0.75), right: convertInchesToTwip(0.75) }
      }
    },
    headers: { default: new Header({ children: [
      new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { after: 0 },
        children: [new TextRun({ text: `CQA audit — ${DATA.title}`, font: FONT, size: sz(8), color: GREY })] })
    ] }) },
    footers: { default: new Footer({ children: [
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0 },
        children: [new TextRun({ children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES],
          font: FONT, size: sz(8), color: GREY })] })
    ] }) },
    children
  }]
});

const outName = `CQA_Audit_${DATA.title.replace(/[^A-Za-z0-9]+/g, '_')}.docx`;
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(`out/${outName}`, buf);
  console.log('wrote out/' + outName, (buf.length / 1024 / 1024).toFixed(2) + ' MB',
    '| screens', DATA.screens.length, '| findings', totalItems, '| other', DATA.other.length);
});
