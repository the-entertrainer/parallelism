/**
 * "Parallelism" — 60-minute facilitator deck for instructional designers.
 *
 * 1920 x 1080 (13.333in x 7.5in @ 144dpi), dark liquid-glass visual system.
 * Visual grammar: rose = breaks, mint = holds, iris = structure/current.
 * Motif: the pill triad (three stacked pills — ragged in rose, matched in mint).
 *
 * Morph transitions are NOT written here (no JS library supports them);
 * repeating shapes carry stable `objectName`s so that inject_morph.py can
 * add <p159:morph> and PowerPoint can match them slide to slide.
 */
const path = require("path");
const PptxGenJS = require("pptxgenjs");
const { C, background, sheen, icon } = require("./assets");

const OUT = path.join(__dirname, "build", "parallelism-training.pptx");

// ---------------------------------------------------------------- geometry
const W = 13.333;
const H = 7.5;
const M = 0.62; // slide margin
const CW = W - M * 2; // content width 12.093
const FONT = "Calibri"; // one typeface throughout — the deck obeys its own rule

const NAV_Y = 6.88;
const TOTAL = 21;
const MAP_SLIDE = 10; // the five-places grid doubles as the deck's map

// ---------------------------------------------------------------- helpers
const NOUL = { style: "none" }; // hyperlinks stay clickable, lose the underline

const shadow = (o = {}) => ({
  type: "outer",
  color: "000000",
  blur: o.blur || 26,
  offset: o.offset === undefined ? 7 : o.offset,
  angle: o.angle || 90,
  opacity: o.opacity || 0.45,
});

/** The one motif: a liquid-glass panel. tone = neutral | breaks | holds | iris */
function glass(slide, x, y, w, h, opt = {}) {
  const tone = opt.tone || "neutral";
  const tint = { neutral: C.white, breaks: C.breaks, holds: C.holds, iris: C.iris }[tone];
  const fillT = tone === "neutral" ? (opt.strong ? 89 : 93) : 86;
  const lineT = tone === "neutral" ? 78 : 52;
  slide.addShape("roundRect", {
    x, y, w, h,
    rectRadius: opt.radius || 0.2,
    fill: { color: tint, transparency: opt.fillT === undefined ? fillT : opt.fillT },
    line: { color: tint, width: opt.lineW || 1, transparency: opt.lineT === undefined ? lineT : opt.lineT },
    shadow: shadow(opt.shadow || {}),
    ...(opt.objectName ? { objectName: opt.objectName } : {}),
  });
}

/** Small capsule label (status chips, timing chips, nav chips). */
function chip(slide, x, y, w, h, text, colour, opt = {}) {
  slide.addShape("roundRect", {
    x, y, w, h,
    rectRadius: h / 2,
    fill: { color: colour, transparency: opt.solid ? 12 : 80 },
    line: { color: colour, width: 0.75, transparency: opt.solid ? 20 : 45 },
    ...(opt.objectName ? { objectName: opt.objectName } : {}),
  });
  slide.addText(text, {
    x, y, w, h,
    align: "center", valign: "middle", margin: 0,
    fontFace: FONT, fontSize: opt.fontSize || 10.5, bold: true,
    charSpacing: opt.charSpacing === undefined ? 1.2 : opt.charSpacing,
    color: opt.solid ? C.ink : colour,
    ...(opt.link ? { hyperlink: { slide: opt.link }, underline: NOUL } : {}),
    ...(opt.textObjectName ? { objectName: opt.textObjectName } : {}),
  });
}

function kicker(slide, text, colour) {
  slide.addText(text, {
    x: M, y: 0.38, w: CW, h: 0.28, margin: 0, objectName: "chromeKicker",
    fontFace: FONT, fontSize: 11.5, bold: true, charSpacing: 3.4,
    color: colour || C.iris,
  });
}

function title(slide, text, opt = {}) {
  slide.addText(text, {
    x: M, y: opt.y === undefined ? 0.68 : opt.y, w: opt.w || CW, h: opt.h || 0.66, margin: 0,
    fontFace: FONT, fontSize: opt.fontSize || 34, bold: true, color: C.text,
    valign: "top",
    objectName: opt.objectName || "chromeTitle",
  });
}

function subtitle(slide, text, opt = {}) {
  slide.addText(text, {
    x: M, y: opt.y === undefined ? 1.34 : opt.y, w: opt.w || CW * 0.78, h: 0.4, margin: 0,
    fontFace: FONT, fontSize: opt.fontSize || 15.5, color: C.textDim, italic: !!opt.italic,
    objectName: "chromeSub",
  });
}

/** On-screen navigation: section + timing on the left, links on the right. */
function nav(slide, idx, section, minutes) {
  const left = minutes ? `${section}   ·   ${minutes} MIN` : section;
  slide.addText(left, {
    x: M, y: NAV_Y, w: 6, h: 0.3, margin: 0, valign: "middle",
    fontFace: FONT, fontSize: 10, bold: true, charSpacing: 2.4, color: C.textFaint,
    objectName: "navSection",
  });
  slide.addText(`${String(idx).padStart(2, "0")} / ${TOTAL}`, {
    x: W - M - 4.05, y: NAV_Y, w: 1.1, h: 0.3, margin: 0, valign: "middle", align: "right",
    fontFace: FONT, fontSize: 10, bold: true, charSpacing: 1.6, color: C.textFaint,
    objectName: "navCount",
  });
  const chips = [
    ["‹  BACK", Math.max(1, idx - 1)],
    ["MAP", MAP_SLIDE],
    ["NEXT  ›", Math.min(TOTAL, idx + 1)],
  ];
  let cx = W - M - 2.7;
  chips.forEach(([label, target], i) => {
    const w = i === 1 ? 0.66 : 0.92;
    chip(slide, cx, NAV_Y - 0.02, w, 0.34, label, C.textDim, {
      fontSize: 9.5, link: target,
      objectName: `navChip${i}`, textObjectName: `navChipText${i}`,
    });
    cx += w + 0.1;
  });
}

/** Bullet-free item list, laid out as ONE flowing text box so wrapped items
 *  push the next item down instead of overlapping it. */
function itemList(slide, x, y, w, items, colour, opt = {}) {
  const marker = (i) =>
    opt.numbered ? `${i + 1}.  ` : opt.marker === false ? "" : opt.marker === "dot" ? "\u2022  " : "\u2014  ";
  const runs = [];
  items.forEach((t, i) => {
    runs.push({ text: marker(i), options: { color: colour, bold: true } });
    runs.push({
      text: t,
      options: { color: opt.textColor || C.text, breakLine: i < items.length - 1 },
    });
  });
  slide.addText(runs, {
    x, y, w, h: opt.h || 3.0, margin: 0, valign: "top",
    fontFace: FONT, fontSize: opt.fontSize || 14,
    lineSpacing: opt.lineSpacing || 19,
    paraSpaceAfter: opt.gap === undefined ? 9 : opt.gap,
  });
}

/** The pill-triad motif. ragged=true → rose & mismatched; false → mint & matched. */
function pillTriad(slide, x, y, w, ragged, opt = {}) {
  const widths = ragged ? [1.0, 0.55, 0.82] : [1, 1, 1];
  const colour = ragged ? C.breaks : C.holds;
  const ph = opt.pillH || 0.3;
  const gap = opt.gap || 0.22;
  widths.forEach((f, i) => {
    slide.addShape("roundRect", {
      x, y: y + i * (ph + gap), w: w * f, h: ph,
      rectRadius: ph / 2,
      fill: { color: colour, transparency: 30 },
      line: { color: colour, width: 1, transparency: 10 },
      ...(opt.objectName ? { objectName: `${opt.objectName}${i + 1}` } : {}),
    });
  });
}

// ---------------------------------------------------------------- content
const PLACES = [
  { n: "01", name: "Learning\nobjectives", flat: "Learning objectives", ic: "FaBullseye" },
  { n: "02", name: "Bullet\nlists", flat: "Bullet lists", ic: "FaListUl" },
  { n: "03", name: "Nav & menu\nlabels", flat: "Nav & menu labels", ic: "FaCompass" },
  { n: "04", name: "Assessment\nitems", flat: "Assessment items", ic: "FaCircleQuestion" },
  { n: "05", name: "Module\narchitecture", flat: "Module architecture", ic: "FaSitemap" },
];

// grid geometry (slide 10 and slide 20) — must match so morph reads as one move
const GRID = { y: 2.28, h: 3.32, w: 2.24, gap: 0.22, x0: M };
const gridX = (i) => GRID.x0 + i * (GRID.w + GRID.gap);
// rail geometry (slides 11-15)
const RAIL = { x: M, w: 3.3, h: 0.82, gap: 0.16, y0: 1.72 };
const railY = (i) => RAIL.y0 + i * (RAIL.h + RAIL.gap);
const PANE = { x: 4.3, w: 4.05, gap: 0.31, y: 1.72, h: 4.74 }; // bottom aligns with the rail

async function main() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "HD", width: W, height: H });
  pres.layout = "HD";
  pres.author = "Parallelism training";
  pres.title = "Parallelism — a 60-minute working session";

  const BG = await background(false);
  const BG_HERO = await background(true);
  const SHEEN = await sheen(0.55);
  const SHEEN_SOFT = await sheen(0.3);

  const ic = {};
  for (const p of PLACES) ic[p.ic] = await icon(p.ic, "FFFFFF");
  for (const [k, col] of [
    ["FaShapes", C.iris], ["FaLayerGroup", C.iris], ["FaWaveSquare", C.iris],
    ["FaGaugeHigh", C.iris], ["FaSignsPost", C.iris], ["FaScaleBalanced", C.breaks],
    ["FaMagnifyingGlass", C.holds], ["FaTag", C.holds], ["FaWrench", C.holds],
    ["FaUser", C.iris], ["FaUserGroup", C.iris], ["FaClipboardCheck", C.iris],
  ]) ic[k] = await icon(k, col);

  const slide = (hero) => {
    const s = pres.addSlide();
    s.background = { data: hero ? BG_HERO : BG };
    return s;
  };

  // =============================================================== 1 TITLE
  {
    const s = slide(true);
    s.addImage({ data: SHEEN, x: -1.6, y: -1.9, w: 6.4, h: 6.4, transparency: 55 });
    glass(s, M, 1.32, 7.5, 4.5, { strong: true, radius: 0.28 });
    s.addText("A 60-MINUTE WORKING SESSION", {
      x: M + 0.62, y: 1.86, w: 6.2, h: 0.3, margin: 0,
      fontFace: FONT, fontSize: 11.5, bold: true, charSpacing: 3.6, color: C.iris,
    });
    s.addText("Parallelism", {
      x: M + 0.62, y: 2.24, w: 6.4, h: 1.3, margin: 0,
      fontFace: FONT, fontSize: 72, bold: true, color: C.text,
    });
    s.addText("Same job, same shape.", {
      x: M + 0.62, y: 3.56, w: 6.4, h: 0.5, margin: 0,
      fontFace: FONT, fontSize: 25, color: C.holds,
    });
    s.addText(
      "For instructional designers and content developers who already know the term — and want to catch it in their own drafts.",
      { x: M + 0.62, y: 4.18, w: 6.3, h: 0.9, margin: 0, fontFace: FONT, fontSize: 14.5, color: C.textDim, lineSpacing: 21 }
    );
    chip(s, M + 0.62, 5.18, 2.0, 0.36, "FACILITATOR DECK", C.iris, { fontSize: 9.5 });
    chip(s, M + 2.72, 5.18, 1.62, 0.36, "21 SLIDES", C.textDim, { fontSize: 9.5 });

    // motif, stated as the deck's own legend
    glass(s, 8.62, 1.32, 4.09, 4.5, { radius: 0.28 });
    s.addText("THE VISUAL GRAMMAR", {
      x: 9.02, y: 1.78, w: 3.3, h: 0.3, margin: 0,
      fontFace: FONT, fontSize: 10.5, bold: true, charSpacing: 2.6, color: C.textFaint,
    });
    pillTriad(s, 9.02, 2.2, 2.32, true, { pillH: 0.28, gap: 0.18 });
    s.addText("BREAKS", { x: 11.52, y: 2.54, w: 1.1, h: 0.3, margin: 0, fontFace: FONT, fontSize: 11, bold: true, charSpacing: 1.6, color: C.breaks });
    pillTriad(s, 9.02, 3.62, 2.32, false, { pillH: 0.28, gap: 0.18 });
    s.addText("HOLDS", { x: 11.52, y: 3.96, w: 1.1, h: 0.3, margin: 0, fontFace: FONT, fontSize: 11, bold: true, charSpacing: 1.6, color: C.holds });
    s.addText("Every slide in this deck uses these two colours the same way.", {
      x: 9.02, y: 4.92, w: 3.3, h: 0.8, margin: 0, fontFace: FONT, fontSize: 12, color: C.textDim, lineSpacing: 17,
    });
    nav(s, 1, "OPENING", 1);
    s.addNotes(
      "SAY: Welcome. Sixty minutes, and we're spending most of it on your own material, not mine. You all know what parallelism is. The goal today is narrower: catch it in your own drafts, before review does.\n\n" +
      "POINT AT the right-hand panel: pink means the structure broke, green means it holds. That's the only colour code in the deck — you'll see it about forty times.\n\n" +
      "EXPECT: low energy at minute one; that's fine, the icebreaker does the waking up.\n\n" +
      "IF THE ROOM STALLS: ask for a show of hands — who has had a reviewer say 'these objectives feel off' without saying why? Most hands go up. That's the session."
    );
  }

  // ======================================================= 2-5 ICEBREAKER
  const quotes = [
    {
      broken: "“I came, I was in the area seeing things, and then conquest happened.”",
      fixed: "“I came, I saw, I conquered.”",
      attr: "Julius Caesar, reporting from Pontus · 47 BC",
      diagBroke: ["came — simple past verb", "was seeing — past continuous + object", "conquest happened — noun as subject"],
      diagFix: ["came — simple past", "saw — simple past", "conquered — simple past"],
      tag: "THREE VERBS \u00b7 THREE SHAPES",
      tagFix: "THREE VERBS \u00b7 ONE SHAPE",
      size: 30,
    },
    {
      broken: "“Ask not what your country can do for you — the real question concerns your own contribution to it.”",
      fixed: "“Ask not what your country can do for you — ask what you can do for your country.”",
      attr: "John F. Kennedy, inaugural address · 20 January 1961",
      diagBroke: ["Clause 1 — imperative: ask not what…", "Clause 2 — declarative: the question concerns…", "The mirror is gone; only the meaning survives"],
      diagFix: ["Clause 1 — imperative: ask not what…", "Clause 2 — imperative: ask what…", "Same verb, same frame, terms swapped"],
      tag: "A MIRROR MISSING ONE SIDE",
      tagFix: "A MIRROR WITH BOTH SIDES",
      size: 25,
    },
  ];

  quotes.forEach((q, qi) => {
    const suffix = qi + 1; // objectName suffix keeps quote A and quote B apart
    [false, true].forEach((fixed) => {
      const s = slide(false);
      const tone = fixed ? "holds" : "breaks";
      const colour = fixed ? C.holds : C.breaks;
      kicker(s, fixed ? "ICEBREAKER  ·  THE RESTORED LINE" : "ICEBREAKER  ·  READ IT OUT LOUD", colour);
      title(s, fixed ? "Here is what it actually says." : "Something is wrong. Don’t say what yet.", {
        objectName: `quoteHead${suffix}`,
      });

      // quote card — identical objectName across the pair drives the morph
      glass(s, M, 1.72, CW, 2.62, { tone, radius: 0.26, objectName: `quoteCard${suffix}` });
      s.addImage({ data: SHEEN_SOFT, x: M - 0.5, y: 1.0, w: 3.2, h: 3.2, transparency: 68, objectName: `quoteSheen${suffix}` });
      s.addText(fixed ? q.fixed : q.broken, {
        x: M + 0.52, y: 1.94, w: CW - 1.04, h: 1.56, margin: 0, valign: "middle",
        fontFace: FONT, fontSize: q.size, bold: true, color: C.text, lineSpacing: q.size + 8,
        objectName: `quoteText${suffix}`,
      });
      s.addText(q.attr, {
        x: M + 0.52, y: 3.62, w: 7.4, h: 0.34, margin: 0,
        fontFace: FONT, fontSize: 13, italic: true, color: C.textDim,
        objectName: `quoteAttr${suffix}`,
      });
      chip(s, W - M - 4.0, 3.6, 3.48, 0.38, fixed ? q.tagFix : q.tag, colour, {
        fontSize: 9.5, objectName: `quoteTag${suffix}`, textObjectName: `quoteTagText${suffix}`,
      });

      // diagnosis panel
      glass(s, M, 4.56, 7.34, 2.0, { radius: 0.22, objectName: `quoteDiag${suffix}` });
      s.addText(fixed ? "WHAT CHANGED" : "WHAT YOUR EAR IS FLAGGING", {
        x: M + 0.42, y: 4.8, w: 6.6, h: 0.28, margin: 0,
        fontFace: FONT, fontSize: 10.5, bold: true, charSpacing: 2.6, color: colour,
        objectName: `quoteDiagHead${suffix}`,
      });
      itemList(s, M + 0.42, 5.18, 6.5, fixed ? q.diagFix : q.diagBroke, colour, { fontSize: 13.5, lineHeight: 0.38 });

      // facilitator prompt
      glass(s, 8.34, 4.56, 4.39, 2.0, { tone: fixed ? "holds" : "iris", radius: 0.22, objectName: `quotePrompt${suffix}` });
      s.addText(fixed ? "SAY THIS" : "ASK THE ROOM", {
        x: 8.72, y: 4.8, w: 3.7, h: 0.28, margin: 0,
        fontFace: FONT, fontSize: 10.5, bold: true, charSpacing: 2.6, color: fixed ? C.holds : C.iris,
        objectName: `quotePromptHead${suffix}`,
      });
      s.addText(
        fixed
          ? (qi === 0
            ? "“Nothing was added. Three words were made to match — and the line became quotable.”"
            : "“The idea never changed. The shape did. Sixty years later we still quote the shape.”")
          : "“On a count of three, one word for how that felt — not why.”",
        { x: 8.72, y: 5.16, w: 3.65, h: 1.2, margin: 0, fontFace: FONT, fontSize: 13.5, color: C.text, lineSpacing: 19,
          objectName: `quotePromptText${suffix}` }
      );

      const idx = 2 + qi * 2 + (fixed ? 1 : 0);
      nav(s, idx, "ICEBREAKER", fixed ? undefined : 4);
      s.addNotes(
        fixed
          ? `SAY: ${qi === 0 ? "Caesar" : "Kennedy"} wrote it like this. Read the real line, then the broken one again, back to back.\n\n` +
            "DO: advance slowly — the words re-assemble one at a time (Morph, by word). Let people watch it land; don't talk over the animation.\n\n" +
            "SAY: I didn't add information. I made the parts match. That's the whole move.\n\n" +
            "EXPECT: a laugh or an audible 'oh'. Someone usually says 'I read that first version three times and couldn't see it'. Perfect — that's the point of the next slide.\n\n" +
            "IF THE ROOM STALLS: count the verbs out loud — came, saw, conquered — one beat each. The rhythm does the arguing for you."
          : `DO: put this up and say nothing for a slow ten count. Let them read it twice.\n\n` +
            `SAY: This is a famous line. I've done something to it. Don't tell me what — tell me how it feels.\n\n` +
            "EXPECT: 'clunky', 'wordy', 'off', 'someone padded it'. Almost nobody names the grammar on the first pass, and that is exactly the finding you want on record.\n\n" +
            "DO: write two or three of their words on the whiteboard — you will point back at them in four minutes.\n\n" +
            "IF THE ROOM STALLS: ask 'would you post this on LinkedIn as a quote?' Nobody would. Ask why not — that unlocks it."
      );
    });
  });

  // =============================================================== 6 PIVOT
  {
    const s = slide(true);
    kicker(s, "THE PIVOT  ·  WHY THAT MATTERED", C.iris);
    title(s, "You felt it before you could name it.");
    subtitle(s, "Your learners do the same thing — and they blame the subject, not the sentence.", { w: CW });
    s.addImage({ data: SHEEN_SOFT, x: 8.4, y: -1.2, w: 6.2, h: 6.2, transparency: 60 });

    glass(s, M, 1.86, 5.7, 2.16, { strong: true, radius: 0.26 });
    s.addText(
      "“Clunky.”  “Wordy.”  “Off.”",
      { x: M + 0.46, y: 2.12, w: 4.9, h: 0.5, margin: 0, fontFace: FONT, fontSize: 22, bold: true, color: C.breaks }
    );
    s.addText(
      "Those are the words the room reached for. Not one of them is a grammatical diagnosis — and every one of them is accurate.",
      { x: M + 0.46, y: 2.74, w: 4.85, h: 1.1, margin: 0, fontFace: FONT, fontSize: 15, color: C.text, lineSpacing: 22 }
    );

    glass(s, M, 4.18, 5.7, 2.32, { tone: "iris", radius: 0.26 });
    s.addText("THE TRANSFER", {
      x: M + 0.46, y: 4.44, w: 4.9, h: 0.28, margin: 0,
      fontFace: FONT, fontSize: 10.5, bold: true, charSpacing: 2.6, color: C.iris,
    });
    s.addText(
      "A learner in your course cannot pause to diagnose your syntax. They register the friction, attribute it to the subject, and conclude the material is harder than it is.",
      { x: M + 0.46, y: 4.82, w: 4.85, h: 1.5, margin: 0, fontFace: FONT, fontSize: 14.5, color: C.text, lineSpacing: 21 }
    );

    const rows = [
      ["FaGaugeHigh", "They don’t report it", "No one files a ticket saying ‘clause forms are inconsistent’. They file ‘this module was confusing’ — or they file nothing and disengage."],
      ["FaSignsPost", "They read form as meaning", "Different shapes imply different kinds of thing. Your ragged list quietly tells them these five items aren’t the same kind of item."],
      ["FaScaleBalanced", "The cost lands on comprehension", "Effort spent reconciling shape is effort not spent on content. You built the load; the learner pays it."],
    ];
    let y = 1.86;
    rows.forEach(([iconName, head, body], i) => {
      glass(s, 6.72, y, 6.0, 1.46, { radius: 0.22, tone: i === 2 ? "breaks" : "neutral" });
      s.addShape("ellipse", {
        x: 7.02, y: y + 0.32, w: 0.62, h: 0.62,
        fill: { color: i === 2 ? C.breaks : C.iris, transparency: 72 },
        line: { color: i === 2 ? C.breaks : C.iris, width: 1, transparency: 40 },
      });
      s.addImage({ data: i === 2 ? ic.FaScaleBalanced : ic[iconName], x: 7.17, y: y + 0.47, w: 0.32, h: 0.32 });
      s.addText(head, {
        x: 7.82, y: y + 0.22, w: 4.6, h: 0.34, margin: 0,
        fontFace: FONT, fontSize: 17, bold: true, color: C.text,
      });
      s.addText(body, {
        x: 7.82, y: y + 0.6, w: 4.62, h: 0.8, margin: 0,
        fontFace: FONT, fontSize: 12.5, color: C.textDim, lineSpacing: 17,
      });
      y += 1.58;
    });
    nav(s, 6, "THE PIVOT", 2);
    s.addNotes(
      "SAY: Point at the whiteboard words. You produced a valid quality finding in four seconds, using no technical vocabulary. Your learners do the same thing — except they blame the subject instead of the sentence.\n\n" +
      "SAY: That's the reframe for today. Parallelism isn't an editorial nicety. It's load management.\n\n" +
      "EXPECT: nodding, and often a war story about a course that 'tested fine but felt bad'. Take one, keep it short, it will pay off in the audit exercise.\n\n" +
      "IF THE ROOM STALLS: name your own example — a module where the fix was purely structural and the satisfaction scores moved. If you don't have one, use the third card: the cost is real even when nobody can name it."
    );
  }

  // ========================================================== 7 DEFINITION
  {
    const s = slide(false);
    kicker(s, "DEFINITION  ·  THE THREE TESTS", C.iris);
    title(s, "Parallelism is three agreements, not one.");
    subtitle(s, "A series is parallel only when all three hold. Two out of three still reads as broken.");

    const tests = [
      ["FaShapes", "Same grammatical form", "Every item is the same part of speech in the same construction — all imperative verbs, all noun phrases, all full clauses.",
        "Draft · review · approving", "Draft · review · approve"],
      ["FaLayerGroup", "Same conceptual level", "Every item sits at the same altitude of abstraction — no strategy hiding among tasks, no whole among its own parts.",
        "Europe · Asia · Belgium", "Europe · Asia · Africa"],
      ["FaWaveSquare", "Same rhythm and length", "Items land within a comparable span. A twenty-word item beside a two-word item breaks the set.",
        "Save it · Send it · Archive it in the shared quarterly folder", "Save it · Send it · Archive it"],
    ];
    const cw = (CW - 0.44) / 3;
    tests.forEach(([iconName, head, body, bad, good], i) => {
      const x = M + i * (cw + 0.22);
      glass(s, x, 2.0, cw, 4.3, { radius: 0.24, strong: i === 1 });
      s.addShape("ellipse", {
        x: x + 0.42, y: 2.34, w: 0.72, h: 0.72,
        fill: { color: C.iris, transparency: 70 }, line: { color: C.iris, width: 1, transparency: 36 },
      });
      s.addImage({ data: ic[iconName], x: x + 0.6, y: 2.52, w: 0.36, h: 0.36 });
      s.addText(`0${i + 1}`, {
        x: x + cw - 1.0, y: 2.34, w: 0.6, h: 0.4, margin: 0, align: "right",
        fontFace: FONT, fontSize: 20, bold: true, color: C.textFaint,
      });
      s.addText(head, {
        x: x + 0.42, y: 3.24, w: cw - 0.84, h: 0.7, margin: 0,
        fontFace: FONT, fontSize: 19, bold: true, color: C.text, lineSpacing: 24,
      });
      s.addText(body, {
        x: x + 0.42, y: 3.96, w: cw - 0.84, h: 1.16, margin: 0,
        fontFace: FONT, fontSize: 13, color: C.textDim, lineSpacing: 18,
      });
      s.addText([{ text: "✕  ", options: { color: C.breaks, bold: true } }, { text: bad, options: { color: C.text } }], {
        x: x + 0.42, y: 5.16, w: cw - 0.84, h: 0.52, margin: 0,
        fontFace: FONT, fontSize: 12.5, italic: true, lineSpacing: 16,
      });
      s.addText([{ text: "✓  ", options: { color: C.holds, bold: true } }, { text: good, options: { color: C.text } }], {
        x: x + 0.42, y: 5.74, w: cw - 0.84, h: 0.5, margin: 0,
        fontFace: FONT, fontSize: 12.5, italic: true, lineSpacing: 16,
      });
    });
    nav(s, 7, "DEFINITION", 5);
    s.addNotes(
      "SAY: Most style guides stop at test one — match the words. That's why people still ship broken lists that pass a grammar check.\n\n" +
      "WALK the three cards. Spend longest on level: Europe, Asia, Belgium is the one everybody recognises and nobody catches in their own outline.\n\n" +
      "SAY: Rhythm feels like a preference. It isn't — it's the signal readers use to decide whether items are peers.\n\n" +
      "EXPECT: pushback on rhythm — 'sometimes one item just needs more words'. Agree, then answer: then it may not belong in the list, or the list needs a different frame.\n\n" +
      "IF THE ROOM STALLS: ask them to invent a fourth item for the middle example. Half the room will produce a country, which proves the test is already installed in their ear."
    );
  }

  // ========================================================== 8 WHY FOR ID
  {
    const s = slide(false);
    kicker(s, "STAKES  ·  WHY AN ID SHOULD CARE", C.iris);
    title(s, "Three reasons this is your job.");
    subtitle(s, "Not the copy editor’s — each of these is a design decision with a measurable cost.", { w: CW });

    const reasons = [
      ["Comparison cost", "Learners compare list items whether you intend it or not. Matched form makes the comparison free; mismatched form makes them re-parse before they can weigh.", "COGNITIVE LOAD", C.iris],
      ["Structure as signal", "Shape is instructional information. Consistent form says ‘these are peers’. Break it and you have silently taught a hierarchy you never designed.", "INSTRUCTIONAL DESIGN", C.iris],
      ["Assessment validity", "A non-parallel distractor is a giveaway: the odd option out can be picked on shape alone, so the item measures test-wiseness, not knowledge.", "THE SHARPEST ONE", C.breaks],
    ];
    let y = 1.86;
    reasons.forEach(([head, body, tag, colour], i) => {
      const isKey = i === 2;
      glass(s, M, y, CW, 1.4, { radius: 0.24, tone: isKey ? "breaks" : "neutral", strong: isKey });
      s.addText(`0${i + 1}`, {
        x: M + 0.44, y: y + 0.34, w: 0.9, h: 0.66, margin: 0,
        fontFace: FONT, fontSize: 32, bold: true, color: isKey ? C.breaks : C.textFaint,
      });
      s.addText(head, {
        x: M + 1.42, y: y + 0.22, w: 3.5, h: 0.42, margin: 0,
        fontFace: FONT, fontSize: 21, bold: true, color: C.text,
      });
      chip(s, M + 1.42, y + 0.7, 2.5, 0.34, tag, colour, { fontSize: 9 });
      s.addText(body, {
        x: M + 5.2, y: y + 0.24, w: 6.3, h: 0.92, margin: 0,
        fontFace: FONT, fontSize: 14, color: isKey ? C.text : C.textDim, lineSpacing: 20,
      });
      y += 1.54;
    });
    s.addText(
      "The first two cost comprehension. The third costs you your data — and you will defend that data in a stakeholder review.",
      { x: M, y: 6.46, w: CW, h: 0.34, margin: 0, fontFace: FONT, fontSize: 13.5, italic: true, color: C.textDim }
    );
    nav(s, 8, "STAKES", 5);
    s.addNotes(
      "SAY: If someone in your org treats this as an editing preference, these are the three arguments that end that conversation.\n\n" +
      "DO: move fast through one and two, then slow right down on three. Say plainly: a broken distractor is a scoring bug.\n\n" +
      "EXPECT: the assessment point is where SMEs and QA people sit up. Someone may say 'we've been doing that for years'. Don't let the room get defensive — next slide shows it, and the fix is thirty seconds of work.\n\n" +
      "IF THE ROOM STALLS: ask who has ever passed a quiz on a topic they didn't know. Most people have, and they know exactly how they did it."
    );
  }

  // ================================================ 9 ASSESSMENT VALIDITY
  {
    const s = slide(false);
    kicker(s, "STAKES  ·  WORKED EXAMPLE", C.breaks);
    title(s, "The distractor that answers itself.");
    subtitle(s, "Same stem, same content, same key. Only the shape of the options changed.");

    const bad = [
      "Reset it",
      "Escalate",
      "Notify the account owner in writing, document the incident in the CRM, and set a 24-hour follow-up reminder",
      "Wait",
    ];
    const good = [
      "Reset the credential",
      "Escalate to tier two",
      "Notify the account owner",
      "Close the ticket",
    ];

    const panes = [
      { x: M, tone: "breaks", label: "BREAKS", colour: C.breaks, head: "Gameable", opts: bad,
        ys: [3.2, 3.72, 4.24, 5.06], noteY: 5.5,
        note: "Option C is longer, hedged and complete. A learner who knows nothing about credentials still picks it — the shape says ‘this is the careful answer’." },
      { x: M + CW / 2 + 0.16, tone: "holds", label: "HOLDS", colour: C.holds, head: "Valid", opts: good,
        ys: [3.2, 3.72, 4.24, 4.76], noteY: 5.5,
        note: "Four verb phrases of comparable weight. Nothing distinguishes the key except knowing the procedure — which is what the item claims to measure." },
    ];
    const pw = CW / 2 - 0.16;
    panes.forEach((p) => {
      glass(s, p.x, 1.98, pw, 4.34, { tone: p.tone, radius: 0.24 });
      chip(s, p.x + 0.4, 2.24, 1.32, 0.34, p.label, p.colour, { fontSize: 9.5 });
      s.addText(p.head, {
        x: p.x + 1.86, y: 2.2, w: 2.4, h: 0.4, margin: 0,
        fontFace: FONT, fontSize: 19, bold: true, color: C.text,
      });
      s.addText("Stem · A user reports a locked account. What is your first action?", {
        x: p.x + 0.4, y: 2.72, w: pw - 0.8, h: 0.4, margin: 0,
        fontFace: FONT, fontSize: 13, italic: true, color: C.textDim, lineSpacing: 18,
      });
      p.opts.forEach((o, i) => {
        s.addText(
          [{ text: `${"ABCD"[i]}   `, options: { color: p.colour, bold: true } }, { text: o, options: { color: C.text } }],
          { x: p.x + 0.4, y: p.ys[i], w: pw - 0.8, h: 0.8, margin: 0, valign: "top",
            fontFace: FONT, fontSize: 13.5, lineSpacing: 18 }
        );
      });
      s.addText(p.note, {
        x: p.x + 0.4, y: p.noteY, w: pw - 0.8, h: 0.8, margin: 0,
        fontFace: FONT, fontSize: 12.5, color: C.textDim, lineSpacing: 17,
      });
    });
    nav(s, 9, "STAKES", 3);
    s.addNotes(
      "DO: read the left-hand stem and options aloud, then ask for a show of hands for C. It will be near-unanimous.\n\n" +
      "SAY: Nobody in this room troubleshoots accounts for a living, and we all just got it right. That item measured nothing.\n\n" +
      "SAY: The three tells are length, hedging, and completeness — the longest, most qualified, most thorough option is almost always the key, and test-wise learners know it.\n\n" +
      "EXPECT: someone says 'but the correct answer genuinely needs more detail'. Answer: then move the detail into the stem or the feedback, not the option.\n\n" +
      "IF THE ROOM STALLS: ask them to guess the key on the right-hand version. They can't — that's the deliverable."
    );
  }

  // ==================================================== 10 FIVE PLACES GRID
  {
    const s = slide(false);
    kicker(s, "THE FIELD GUIDE  ·  WHERE IT BREAKS", C.iris);
    title(s, "Five places parallelism breaks in a course.");
    subtitle(s, "Ranked by how often it ships. The next five slides take one each — real before, real after.", { w: CW });

    PLACES.forEach((p, i) => {
      const x = gridX(i);
      glass(s, x, GRID.y, GRID.w, 3.6, { radius: 0.24, objectName: `placeCard${i + 1}`, strong: false });
      s.addShape("ellipse", {
        x: x + 0.38, y: GRID.y + 0.42, w: 0.74, h: 0.74,
        fill: { color: C.iris, transparency: 66 }, line: { color: C.iris, width: 1, transparency: 34 },
        objectName: `placeDisc${i + 1}`,
      });
      s.addImage({ data: ic[p.ic], x: x + 0.57, y: GRID.y + 0.61, w: 0.36, h: 0.36, objectName: `placeIcon${i + 1}` });
      s.addText(p.n, {
        x: x + 0.38, y: GRID.y + 1.36, w: GRID.w - 0.76, h: 0.34, margin: 0,
        fontFace: FONT, fontSize: 12, bold: true, charSpacing: 2, color: C.textFaint,
        objectName: `placeNum${i + 1}`,
      });
      s.addText(p.name, {
        x: x + 0.38, y: GRID.y + 1.64, w: GRID.w - 0.5, h: 1.0, margin: 0,
        fontFace: FONT, fontSize: 16, bold: true, color: C.text, lineSpacing: 21,
        objectName: `placeName${i + 1}`,
      });
      chip(s, x + 0.38, GRID.y + 2.94, GRID.w - 0.76, 0.34, `SLIDE ${11 + i}`, C.textDim, {
        fontSize: 9, link: 11 + i, objectName: `placeJump${i + 1}`, textObjectName: `placeJumpText${i + 1}`,
      });
    });
    glass(s, M, 6.06, CW, 0.5, { radius: 0.18, tone: "iris" });
    s.addText(
      "Same treatment on each: one place, one broken example from the wild, one repaired version beside it.",
      { x: M + 0.4, y: 6.06, w: CW - 0.8, h: 0.5, margin: 0, valign: "middle",
        fontFace: FONT, fontSize: 13, color: C.text }
    );
    nav(s, 10, "FIELD GUIDE", 2);
    s.addNotes(
      "SAY: Five places. In my experience they break in roughly this order of frequency, and in exactly the reverse order of how expensive they are to fix later.\n\n" +
      "DO: read the five out loud and ask people to pick the one they're worried about in their current project. Note the tally — you'll use it to decide where to linger.\n\n" +
      "DO: these five cards are clickable, and they follow you as a rail on the next five slides. Point that out once; people navigate the deck themselves in the practice rounds.\n\n" +
      "EXPECT: objectives and bullets get most hands. Architecture gets almost none, which is why it's last and why it costs the most.\n\n" +
      "IF THE ROOM STALLS: pick number one yourself and move — the deep dives create the discussion."
    );
  }

  // ================================================= 11-15 THE FIVE PLACES
  const deep = [
    {
      before: ["Identify the three phases of onboarding", "Understanding of escalation paths", "The learner will be able to close a ticket", "Empathy"],
      after: ["Identify the three phases of onboarding", "Describe the two escalation paths", "Close a ticket in under three minutes", "Acknowledge a customer’s frustration in one sentence"],
      breakNote: "Four items, four constructions: a verb, a gerund phrase, a stem sentence, a bare noun. The set no longer reads as one set.",
      holdNote: "Every item is a measurable verb phrase completing the same stem. The list can now be scanned, sequenced and assessed.",
      test: "FAILS: FORM",
      say: "Stem once, verb every time.",
      notes:
        "SAY: This is the most-shipped break in our field, and it survives because each item looks fine on its own line.\n\n" +
        "DO: read the left column as one sentence, stem included: 'By the end of this module, learners will… understanding of escalation paths.' The grammar failure becomes audible.\n\n" +
        "SAY: Note what the fix did to 'Empathy'. Forcing the form forced a measurable behaviour. Parallelism is doing instructional design work here, not tidying.\n\n" +
        "EXPECT: recognition, sometimes discomfort — many will spot their own current draft. Say plainly that nobody's work is on the screen.\n\n" +
        "IF THE ROOM STALLS: put your own stem on the board and have them read each item into it out loud.",
    },
    {
      before: ["Faster reviews", "Cost", "You’ll reduce rework", "Compliance is improved", "Better for everyone honestly"],
      after: ["Faster reviews", "Lower cost", "Less rework", "Stronger compliance", "Fewer escalations"],
      breakNote: "A noun phrase, a bare noun, a second-person promise, a passive clause, and an opinion. Five shapes for five peers.",
      holdNote: "Adjective plus noun, five times. The pattern makes the items comparable at a glance — and exposes any weak one.",
      test: "FAILS: FORM + RHYTHM",
      say: "One shape, repeated, is a promise the reader can scan.",
      notes:
        "SAY: Bullets are where parallelism breaks fastest, because bullets get written last and edited by three people.\n\n" +
        "DO: cover the right column and ask which left-hand item is weakest. People pick 'Better for everyone honestly' — then point out they only found it because it broke shape. Broken form hides weak content; matched form exposes it.\n\n" +
        "SAY: When five items take one grammatical shape, a sixth weak item cannot hide.\n\n" +
        "EXPECT: a debate about whether 'Cost' means high or low. That ambiguity is the cost of the break — the fix removed it.\n\n" +
        "IF THE ROOM STALLS: have someone read the left list at speed. The stumble is the evidence.",
    },
    {
      before: ["Overview", "Start here", "How to submit a claim", "Why this matters", "Glossary"],
      after: ["Overview", "Setup", "Claim submission", "Policy rationale", "Glossary"],
      breakNote: "Nouns, an imperative, a how-to question, a why question. Learners cannot predict what a label will deliver, so they click to find out.",
      holdNote: "Five noun phrases naming a destination. The menu becomes predictable, and predictable menus get used instead of searched.",
      test: "FAILS: FORM + LEVEL",
      say: "A menu is a promise about what is behind the door.",
      notes:
        "SAY: Navigation is the most-read text in your course and the least-reviewed. Nobody proofreads a menu.\n\n" +
        "DO: ask what 'Start here' delivers versus 'Overview'. The room won't agree — that's the defect. Two labels, one destination class, no way to tell them apart.\n\n" +
        "SAY: Mixed labels don't just look untidy; they cost clicks. Every unpredictable label is a click spent finding out.\n\n" +
        "EXPECT: someone defends 'Start here' as friendly. Fair — then make them all imperative: Start, Set up, Submit, Understand, Look up. Parallel is the requirement; register is your choice.\n\n" +
        "IF THE ROOM STALLS: pull up any LMS menu in the room and read it aloud.",
    },
    {
      before: ["All of the following are required except", "Which is not a benefit", "The primary purpose of the intake form is to", "True or false: escalation is optional"],
      after: ["Which step is not required during intake?", "Which outcome is not a benefit of triage?", "What is the primary purpose of the intake form?", "Which action is required before escalation?"],
      breakNote: "Four stem shapes across one quiz: a sentence completion, a fragment, another completion, a true/false. Learners re-learn the format at every item.",
      holdNote: "Four direct questions in the same frame. The learner’s attention goes to the content, not to decoding what is being asked.",
      test: "FAILS: FORM · SEE SLIDE 09",
      say: "Decoding the question costs the item its measurement.",
      notes:
        "SAY: Slide nine was about options. This one is about stems, and it's the same disease.\n\n" +
        "DO: read all four left-hand stems in a row. The room feels the gear change at each one. That gear change is measurement noise.\n\n" +
        "SAY: Mixed stem formats inflate the difficulty of every item equally — which means your item statistics are measuring your formatting.\n\n" +
        "EXPECT: 'we're told to vary question types for engagement.' Distinguish item type from stem form: vary the type deliberately, keep the stem frame constant within a type.\n\n" +
        "IF THE ROOM STALLS: ask what a low score on the left-hand quiz would prove. Nothing you could act on.",
    },
    {
      before: ["Module 1 · Introduction", "Module 2 · Working with vendors", "Module 3 · Compliance", "Module 4 · What to do when things go wrong"],
      after: ["Module 1 · Vendor selection", "Module 2 · Vendor onboarding", "Module 3 · Vendor compliance", "Module 4 · Vendor offboarding"],
      breakNote: "A framing device, an activity, an abstract domain, and a scenario. Four different kinds of thing wearing the same numbered label — a level break, not a wording one.",
      holdNote: "One noun pattern across the lifecycle. The architecture now teaches the sequence before a single screen is built.",
      test: "FAILS: LEVEL",
      say: "If the titles aren’t a series, the curriculum isn’t one.",
      notes:
        "SAY: This is the expensive one. Wording breaks cost minutes; architecture breaks cost a rebuild.\n\n" +
        "DO: point out that every left-hand title is defensible on its own. The break only exists across the set — which is exactly why solo authors miss it and why you review titles as a column, never one at a time.\n\n" +
        "SAY: Read the right-hand column. You now know the sequence, the scope and what module five would be, without opening anything.\n\n" +
        "EXPECT: a genuine objection — 'introductions are standard'. Agree, and separate it: front matter isn't a module. Numbering it forces it into a series it doesn't belong to.\n\n" +
        "IF THE ROOM STALLS: ask what module five is, for each column. Only the right-hand one has an answer.",
    },
  ];

  PLACES.forEach((p, idx) => {
    const s = slide(false);
    const d = deep[idx];
    kicker(s, `THE FIELD GUIDE  ·  ${p.n} OF 05`, C.iris);
    title(s, p.flat);

    // persistent rail — identical objectNames as the grid, so morph slides them
    PLACES.forEach((q, j) => {
      const active = j === idx;
      const y = railY(j);
      glass(s, RAIL.x, y, RAIL.w, RAIL.h, {
        radius: 0.2, tone: active ? "iris" : "neutral",
        fillT: active ? 78 : 95, lineT: active ? 34 : 84,
        objectName: `placeCard${j + 1}`,
      });
      s.addShape("ellipse", {
        x: RAIL.x + 0.22, y: y + 0.19, w: 0.44, h: 0.44,
        fill: { color: C.iris, transparency: active ? 50 : 80 },
        line: { color: C.iris, width: 1, transparency: active ? 24 : 62 },
        objectName: `placeDisc${j + 1}`,
      });
      s.addImage({ data: ic[q.ic], x: RAIL.x + 0.32, y: y + 0.29, w: 0.24, h: 0.24,
        transparency: active ? 0 : 45, objectName: `placeIcon${j + 1}` });
      s.addText(q.n, {
        x: RAIL.x + 0.8, y: y + 0.13, w: 0.5, h: 0.26, margin: 0,
        fontFace: FONT, fontSize: 9.5, bold: true, charSpacing: 1.6,
        color: active ? C.iris : C.textFaint, objectName: `placeNum${j + 1}`,
      });
      s.addText(q.flat, {
        x: RAIL.x + 0.8, y: y + 0.37, w: RAIL.w - 1.0, h: 0.32, margin: 0,
        fontFace: FONT, fontSize: 14, bold: active, color: active ? C.text : C.textDim,
        objectName: `placeName${j + 1}`, hyperlink: { slide: 11 + j }, underline: NOUL,
      });
    });

    // before / after pair
    const pairs = [
      { x: PANE.x, tone: "breaks", colour: C.breaks, label: "BREAKS", items: d.before, note: d.breakNote },
      { x: PANE.x + PANE.w + PANE.gap, tone: "holds", colour: C.holds, label: "HOLDS", items: d.after, note: d.holdNote },
    ];
    pairs.forEach((pane) => {
      glass(s, pane.x, PANE.y, PANE.w, PANE.h, { tone: pane.tone, radius: 0.24, objectName: `pane${pane.label}` });
      chip(s, pane.x + 0.34, PANE.y + 0.26, 1.28, 0.34, pane.label, pane.colour, { fontSize: 9.5 });
      itemList(s, pane.x + 0.34, PANE.y + 0.86, PANE.w - 0.68, pane.items, pane.colour, {
        fontSize: 13.5, lineSpacing: 18, gap: 12, h: 2.6,
      });
      s.addText(pane.note, {
        x: pane.x + 0.34, y: PANE.y + 3.58, w: PANE.w - 0.68, h: 1.0, margin: 0,
        fontFace: FONT, fontSize: 12, color: C.textDim, lineSpacing: 17,
      });
    });

    // verdict sits in the title row, so nothing crowds the nav
    chip(s, PANE.x, 1.3, 2.9, 0.36, d.test, C.breaks, { fontSize: 9.5 });
    s.addText(d.say, {
      x: PANE.x + 3.06, y: 1.3, w: 5.36, h: 0.36, margin: 0, valign: "middle",
      fontFace: FONT, fontSize: 12.5, italic: true, color: C.holds,
    });

    nav(s, 11 + idx, "FIELD GUIDE", 2);
    s.addNotes(d.notes);
  });

  // ======================================================= 16 REPAIR METHOD
  {
    const s = slide(true);
    kicker(s, "THE METHOD  ·  TAKE THIS WITH YOU", C.holds);
    title(s, "Three moves. Under a minute per list.");
    subtitle(s, "Run it on any series before it ships. Mechanical on purpose — it should not require taste.", { w: CW });

    const steps = [
      ["FaMagnifyingGlass", "FIND THE SERIES", "Find the series",
        "Anything numbered, bulleted, comma-joined, tabbed or menued — including sets you never called a list, like four module titles in a column.",
        "“Is this two or more things doing the same job?”"],
      ["FaTag", "NAME THE FORM", "Name the form",
        "Say the shape of the strongest item out loud: ‘imperative verb plus object’, ‘noun phrase’, ‘full clause’. Naming beats eyeballing — and gives you a target.",
        "“What exactly is item one, grammatically?”"],
      ["FaWrench", "FORCE THE MATCH", "Force the match",
        "Rewrite every other item into that form, no exceptions. An item that refuses to fit is usually at the wrong level: promote it, demote it, or cut it.",
        "“If it won’t fit the pattern, it isn’t a peer.”"],
    ];
    const cw = (CW - 0.5) / 3;
    steps.forEach(([iconName, , head, body, quote], i) => {
      const x = M + i * (cw + 0.25);
      glass(s, x, 1.92, cw, 3.94, { tone: "holds", radius: 0.24 });
      s.addShape("ellipse", {
        x: x + 0.4, y: 2.26, w: 0.74, h: 0.74,
        fill: { color: C.holds, transparency: 70 }, line: { color: C.holds, width: 1, transparency: 34 },
      });
      s.addImage({ data: ic[iconName], x: x + 0.59, y: 2.45, w: 0.36, h: 0.36 });
      s.addText(`STEP ${i + 1}`, {
        x: x + cw - 1.4, y: 2.36, w: 1.0, h: 0.3, margin: 0, align: "right",
        fontFace: FONT, fontSize: 10, bold: true, charSpacing: 2.2, color: C.holds,
      });
      s.addText(head, {
        x: x + 0.4, y: 3.2, w: cw - 0.8, h: 0.44, margin: 0,
        fontFace: FONT, fontSize: 21, bold: true, color: C.text,
      });
      s.addText(body, {
        x: x + 0.4, y: 3.74, w: cw - 0.8, h: 1.4, margin: 0,
        fontFace: FONT, fontSize: 13, color: C.textDim, lineSpacing: 18,
      });
      s.addText(quote, {
        x: x + 0.4, y: 5.16, w: cw - 0.8, h: 0.62, margin: 0,
        fontFace: FONT, fontSize: 13, italic: true, color: C.holds, lineSpacing: 17,
      });
      if (i < 2) {
        s.addShape("chevron", {
          x: x + cw + 0.03, y: 3.8, w: 0.19, h: 0.34,
          fill: { color: C.holds, transparency: 45 }, line: { color: C.holds, width: 0.75, transparency: 30 },
        });
      }
    });
    glass(s, M, 5.98, CW, 0.6, { radius: 0.2, tone: "iris" });
    s.addText(
      "Find → Name → Force.  A list that survives all three is ready; a list that doesn’t was telling you something.",
      { x: M + 0.4, y: 5.98, w: CW - 0.8, h: 0.6, margin: 0, valign: "middle",
        fontFace: FONT, fontSize: 14, color: C.text }
    );
    nav(s, 16, "THE METHOD", 3);
    s.addNotes(
      "SAY: Three moves, and step two is the one everybody skips. If you don't name the form out loud, you end up matching vibes.\n\n" +
      "DO: demo it live on the objectives list from slide eleven, narrating each step in about forty seconds. Speed is the selling point.\n\n" +
      "SAY: Step three has a bonus. When an item won't fit, you've found a level problem — and that's a content finding, not a wording one.\n\n" +
      "EXPECT: someone asks for a checklist. This slide is it; tell them to screenshot it now, because it's the tool for all three practice rounds.\n\n" +
      "IF THE ROOM STALLS: ask for a volunteer list from a live project and run the method on it in real time. Two minutes, and the exercises land much harder afterwards."
    );
  }

  // ===================================================== 17-19 PRACTICE
  const practice = [
    {
      idx: 17, kick: "PRACTICE 01  ·  SOLO", mins: 5, iconName: "FaUser",
      title: "Spot the break.",
      sub: "Ninety seconds a set. Note which item breaks — and which of the three tests it fails.",
      sets: [
        { label: "SET A", items: ["Plan the sprint", "Run the standup", "Retrospective", "Ship the increment"] },
        { label: "SET B", items: ["Reduce handling time", "Improve first-contact resolution", "Customers are happier", "Cut repeat contacts"] },
        { label: "SET C", items: ["Word documents", "Spreadsheets", "Slide decks", "Microsoft Office"] },
      ],
      side: ["Answer with the test, not the fix.", "A: form — a bare noun among imperatives.", "B: form — a full clause among verb phrases.", "C: level — the container listed among its contents."],
      sideHead: "FACILITATOR KEY",
      notes:
        "DO: hands down, pens out, no discussion. Ninety seconds a set, and call time hard.\n\n" +
        "SAY: I don't want the rewrite yet. I want the diagnosis: which item, which test.\n\n" +
        "DO: take answers set by set, one person each. Set C is the teaching moment — the form is flawless, so anyone hunting only grammar declares it clean.\n\n" +
        "EXPECT: near-perfect scores on A and B, and a genuine split on C. That split is the argument for test two.\n\n" +
        "IF THE ROOM STALLS on C: ask 'is Microsoft Office the same kind of thing as a spreadsheet?' It resolves in one beat.",
    },
    {
      idx: 18, kick: "PRACTICE 02  ·  PAIRS", mins: 8, iconName: "FaUserGroup", numbered: true,
      title: "Repair a real objective set.",
      sub: "Eight minutes in pairs, applying find → name → force. One item will resist.",
      sets: [
        { label: "BY THE END OF THIS MODULE, LEARNERS WILL…", items: [
          "Identify the four intake channels",
          "Escalation triage",
          "Be able to demonstrate empathy on a live call",
          "Improve customer satisfaction across the support organisation",
          "Understand the CRM",
          "Log a call outcome within 60 seconds",
        ] },
      ],
      side: ["Items 2, 3 and 5 are form breaks — quick wins.", "Item 4 is the buried one: the grammar can be fixed, the level cannot. It is an organisational outcome, not a learner behaviour.", "Correct handling: cut item 4 from the objectives and move it to the business case."],
      sideHead: "FACILITATOR KEY",
      notes:
        "DO: pairs, one scribe, eight minutes. Ask each pair to name the target form before rewriting anything — that's step two and they will try to skip it.\n\n" +
        "EXPECT: every pair fixes 2, 3 and 5 within three minutes and then argues about 4. Let that argument run; it's the exercise.\n\n" +
        "SAY when you debrief: item four can be made grammatical — 'Improve customer satisfaction' is a fine verb phrase. It still doesn't belong, because it isn't at the same level: no learner does it on a call, and you can't assess it. Force the match and you expose the level break.\n\n" +
        "EXPECT: one pair keeps item four and rewrites it as 'Apply three satisfaction-driving behaviours'. That's a strong answer — they promoted the intent to the right altitude.\n\n" +
        "IF A PAIR STALLS: ask them to read each item into the stem and say who performs the action, and when."
    },
  ];

  practice.forEach((p) => {
    const s = slide(false);
    kicker(s, p.kick, C.iris);
    title(s, p.title, { w: CW - 1.9 });
    subtitle(s, p.sub, { w: CW - 1.9 });
    chip(s, W - M - 1.5, 0.72, 1.5, 0.42, `${p.mins} MIN`, C.iris, { fontSize: 11, solid: false });

    const mainW = 8.0;
    if (p.sets.length === 3) {
      const sw = (mainW - 0.44) / 3;
      p.sets.forEach((set, i) => {
        const x = M + i * (sw + 0.22);
        glass(s, x, 2.06, sw, 3.5, { radius: 0.22 });
        chip(s, x + 0.28, 2.3, 1.16, 0.32, set.label, C.iris, { fontSize: 9 });
        itemList(s, x + 0.28, 2.84, sw - 0.56, set.items, C.iris, {
          fontSize: 13, lineSpacing: 17, gap: 14, h: 2.5,
        });
      });
    } else {
      const set = p.sets[0];
      glass(s, M, 2.06, mainW, 3.5, { radius: 0.22, strong: true });
      s.addText(set.label, {
        x: M + 0.4, y: 2.32, w: mainW - 0.8, h: 0.3, margin: 0,
        fontFace: FONT, fontSize: 11, bold: true, charSpacing: 2.2, color: C.iris,
      });
      itemList(s, M + 0.4, 2.76, mainW - 0.8, set.items, C.iris, {
        fontSize: 14.5, lineSpacing: 20, gap: 10, h: 2.7, numbered: !!p.numbered,
      });
    }

    // facilitator side panel
    glass(s, M + mainW + 0.28, 2.06, CW - mainW - 0.28, 3.5, { tone: "holds", radius: 0.22 });
    const sx = M + mainW + 0.62;
    const sw2 = CW - mainW - 0.96;
    s.addText(p.sideHead, {
      x: sx, y: 2.32, w: sw2, h: 0.3, margin: 0,
      fontFace: FONT, fontSize: 10, bold: true, charSpacing: 2.2, color: C.holds,
    });
    itemList(s, sx, 2.76, sw2, p.side, C.holds, {
      marker: "dot", fontSize: 12.5, lineSpacing: 17, gap: 12, h: 2.7,
    });

    // running method reminder
    glass(s, M, 5.72, CW, 0.72, { radius: 0.2, tone: "iris" });
    s.addShape("ellipse", {
      x: M + 0.3, y: 5.87, w: 0.42, h: 0.42,
      fill: { color: C.iris, transparency: 62 }, line: { color: C.iris, width: 1, transparency: 34 },
    });
    s.addImage({ data: ic[p.iconName], x: M + 0.39, y: 5.96, w: 0.24, h: 0.24 });
    s.addText(
      "THE METHOD:   Find the series   →   Name the form   →   Force the match",
      { x: M + 0.9, y: 5.72, w: 9.5, h: 0.72, margin: 0, valign: "middle",
        fontFace: FONT, fontSize: 13, bold: true, charSpacing: 0.6, color: C.text }
    );
    chip(s, W - M - 1.86, 5.89, 1.56, 0.38, "SEE SLIDE 16", C.textDim, { fontSize: 9, link: 16 });

    nav(s, p.idx, "PRACTICE", p.mins);
    s.addNotes(p.notes);
  });

  // ============================== 19 PRACTICE 03 — the five-place audit rail
  // The five checkpoints ARE the five places, carrying the same objectNames, so
  // advancing to the summary folds this rail back into the opening grid (Morph).
  {
    const s = slide(false);
    kicker(s, "PRACTICE 03  ·  PAIRS  ·  LIVE WORK", C.iris);
    title(s, "Audit the module you’re building.", { w: CW - 1.9 });
    subtitle(s, "Ten minutes on a draft you can still change this week — the five places, in order.", { w: CW - 1.9 });
    chip(s, W - M - 1.5, 0.72, 1.5, 0.42, "10 MIN", C.iris, { fontSize: 11 });

    const checks = [
      "Read every item into your stem, out loud",
      "Take the longest one: find → name → force",
      "All noun phrases, or all imperatives?",
      "Stems first, then option length and hedging",
      "Read the titles as a column, not one at a time",
    ];
    const rowW = 8.0;
    PLACES.forEach((q, j) => {
      const y = 1.94 + j * 0.84;
      glass(s, M, y, rowW, 0.72, { radius: 0.2, objectName: `placeCard${j + 1}` });
      s.addShape("ellipse", {
        x: M + 0.26, y: y + 0.14, w: 0.44, h: 0.44,
        fill: { color: C.iris, transparency: 74 },
        line: { color: C.iris, width: 1, transparency: 46 },
        objectName: `placeDisc${j + 1}`,
      });
      s.addImage({ data: ic[q.ic], x: M + 0.36, y: y + 0.24, w: 0.24, h: 0.24, objectName: `placeIcon${j + 1}` });
      s.addText(q.n, {
        x: M + 0.84, y: y + 0.24, w: 0.42, h: 0.26, margin: 0,
        fontFace: FONT, fontSize: 9.5, bold: true, charSpacing: 1.6, color: C.textFaint,
        objectName: `placeNum${j + 1}`,
      });
      s.addText(q.flat, {
        x: M + 1.3, y: y + 0.21, w: 2.3, h: 0.3, margin: 0,
        fontFace: FONT, fontSize: 13.5, bold: true, color: C.text,
        objectName: `placeName${j + 1}`, hyperlink: { slide: 11 + j }, underline: NOUL,
      });
      s.addText(checks[j], {
        x: M + 3.72, y: y + 0.21, w: rowW - 3.96, h: 0.3, margin: 0,
        fontFace: FONT, fontSize: 12.5, color: C.textDim,
      });
    });

    // ground rules
    const px = M + rowW + 0.28;
    const pw = CW - rowW - 0.28;
    glass(s, px, 1.94, pw, 4.08, { tone: "holds", radius: 0.22 });
    s.addText("GROUND RULES", {
      x: px + 0.34, y: 2.2, w: pw - 0.68, h: 0.3, margin: 0,
      fontFace: FONT, fontSize: 10, bold: true, charSpacing: 2.2, color: C.holds,
    });
    itemList(s, px + 0.34, 2.64, pw - 0.68, [
      "Swap drafts: audit your partner’s module, not your own.",
      "Every finding names its test — form, level or rhythm.",
      "Leave with one fix you will make before Friday, written down.",
      "Stuck for material? Design module five from slide 15.",
    ], C.holds, { marker: "dot", fontSize: 12.5, lineSpacing: 17, gap: 12, h: 3.1 });

    chip(s, M, 6.28, 2.2, 0.38, "SEE SLIDE 16", C.textDim, { fontSize: 9, link: 16 });
    s.addText("THE METHOD:   Find the series   →   Name the form   →   Force the match", {
      x: M + 2.44, y: 6.28, w: 7.4, h: 0.38, margin: 0, valign: "middle",
      fontFace: FONT, fontSize: 12.5, bold: true, color: C.text,
    });

    nav(s, 19, "PRACTICE", 10);
    s.addNotes(
      "SAY: Something unfinished, please. Auditing a shipped course is entertainment; auditing a draft is work you get paid for.\n\n" +
      "DO: enforce the swap — people cannot see shape in their own prose, they read their own intent. Ten minutes, five checkpoints, in the order on screen.\n\n" +
      "SAY: These are the same five places from the field guide, in the same order. That's your standing audit sequence, not just today's exercise.\n\n" +
      "DO: circulate and ask 'which test did it fail?' A finding without a test name is an opinion, and opinions don't survive an SME review.\n\n" +
      "EXPECT: the objectives check finds something in nearly every draft, and at least one pair finds a level break in their module titles — the highest-value find of the session. Flag it to the room.\n\n" +
      "IF A PAIR HAS NOTHING IN PROGRESS: hand them slide fifteen's module titles and have them design module five, or give them your own live draft."
    );
  }

  // ============================================================= 20 SUMMARY
  {
    const s = slide(false);
    kicker(s, "SUMMARY  ·  THE WHOLE SESSION", C.iris);
    title(s, "Five places to look. Three moves to fix.");
    subtitle(s, "The rail folds back into the map. If you keep one row of this slide, keep the bottom one.", { w: CW });

    PLACES.forEach((p, i) => {
      const x = gridX(i);
      glass(s, x, GRID.y, GRID.w, 2.32, { radius: 0.24, objectName: `placeCard${i + 1}` });
      s.addShape("ellipse", {
        x: x + 0.38, y: GRID.y + 0.34, w: 0.66, h: 0.66,
        fill: { color: C.iris, transparency: 66 }, line: { color: C.iris, width: 1, transparency: 34 },
        objectName: `placeDisc${i + 1}`,
      });
      s.addImage({ data: ic[p.ic], x: x + 0.54, y: GRID.y + 0.5, w: 0.34, h: 0.34, objectName: `placeIcon${i + 1}` });
      s.addText(p.n, {
        x: x + 0.38, y: GRID.y + 1.12, w: GRID.w - 0.76, h: 0.28, margin: 0,
        fontFace: FONT, fontSize: 11, bold: true, charSpacing: 2, color: C.textFaint,
        objectName: `placeNum${i + 1}`,
      });
      s.addText(p.name, {
        x: x + 0.38, y: GRID.y + 1.42, w: GRID.w - 0.66, h: 0.8, margin: 0,
        fontFace: FONT, fontSize: 16, bold: true, color: C.text, lineSpacing: 20,
        objectName: `placeName${i + 1}`, hyperlink: { slide: 11 + i }, underline: NOUL,
      });
    });

    const moves = [
      ["FIND", "the series — anything doing the same job twice"],
      ["NAME", "the form — out loud, using grammar words"],
      ["FORCE", "the match — or admit the item isn’t a peer"],
    ];
    const mw = (CW - 0.44) / 3;
    moves.forEach(([verb, rest], i) => {
      const x = M + i * (mw + 0.22);
      glass(s, x, 4.96, mw, 1.42, { tone: "holds", radius: 0.22 });
      s.addText(`0${i + 1}`, {
        x: x + 0.34, y: 5.16, w: 0.5, h: 0.34, margin: 0,
        fontFace: FONT, fontSize: 13, bold: true, color: C.holds,
      });
      s.addText(verb, {
        x: x + 0.34, y: 5.46, w: mw - 0.68, h: 0.4, margin: 0,
        fontFace: FONT, fontSize: 22, bold: true, charSpacing: 1.4, color: C.text,
      });
      s.addText(rest, {
        x: x + 0.34, y: 5.82, w: mw - 0.68, h: 0.5, margin: 0,
        fontFace: FONT, fontSize: 12, color: C.textDim, lineSpacing: 16,
      });
    });
    nav(s, 20, "SUMMARY", 2);
    s.addNotes(
      "DO: advance and let the rail slide back into the grid before you speak — the bookend is worth three seconds of silence.\n\n" +
      "SAY: Top row is where to look. Bottom row is what to do. That's the session.\n\n" +
      "DO: ask each pair for one commitment out loud — which draft, which of the five, by when. Public commitments get done.\n\n" +
      "EXPECT: energy dips at the summary. Counter it by going straight to commitments instead of recapping content they just practised.\n\n" +
      "IF THE ROOM STALLS: go first — name the one place in your own current build you'll audit this week."
    );
  }

  // =============================================================== 21 CLOSE
  {
    const s = slide(true);
    s.addImage({ data: SHEEN, x: 7.0, y: 1.0, w: 7.2, h: 7.2, transparency: 58 });
    kicker(s, "CLOSE", C.holds);
    glass(s, M, 1.5, 8.5, 3.6, { strong: true, radius: 0.28 });
    s.addText("Same job, same shape.", {
      x: M + 0.6, y: 1.94, w: 7.3, h: 0.76, margin: 0,
      fontFace: FONT, fontSize: 38, bold: true, color: C.text,
    });
    s.addText("Same shape, same job.", {
      x: M + 0.6, y: 2.74, w: 7.3, h: 0.76, margin: 0,
      fontFace: FONT, fontSize: 38, bold: true, color: C.holds,
    });
    s.addText(
      "Said once, it’s a rule. Said twice in the same shape, it’s the demonstration.",
      { x: M + 0.6, y: 4.04, w: 7.2, h: 0.6, margin: 0, fontFace: FONT, fontSize: 15.5, italic: true, color: C.textDim }
    );

    glass(s, 9.5, 1.5, 3.21, 3.6, { radius: 0.28 });
    s.addText("BEFORE YOU LEAVE", {
      x: 9.86, y: 1.86, w: 2.6, h: 0.28, margin: 0,
      fontFace: FONT, fontSize: 10, bold: true, charSpacing: 2.4, color: C.textFaint,
    });
    itemList(s, 9.86, 2.3, 2.55, [
      "One draft, audited",
      "One fix, before Friday",
      "One list, read out loud",
    ], C.holds, { fontSize: 14, lineSpacing: 19, gap: 18, h: 1.6 });
    pillTriad(s, 9.86, 4.1, 2.3, false, { pillH: 0.2, gap: 0.12 });

    glass(s, M, 5.5, CW, 0.9, { tone: "iris", radius: 0.22 });
    s.addText(
      "This deck obeys its own rule: one typeface, one card shape, one meaning per colour — pink breaks, green holds — on all twenty-one slides.",
      { x: M + 0.44, y: 5.5, w: CW - 3.4, h: 0.9, margin: 0, valign: "middle",
        fontFace: FONT, fontSize: 14, color: C.text }
    );
    chip(s, W - M - 2.6, 5.76, 1.16, 0.38, "MAP", C.textDim, { fontSize: 9, link: MAP_SLIDE });
    chip(s, W - M - 1.34, 5.76, 1.34, 0.38, "METHOD", C.textDim, { fontSize: 9, link: 16 });
    nav(s, 21, "CLOSE", 1);
    s.addNotes(
      "SAY the two lines slowly, in order, and let the second one land. Same words, reversed, same shape — that's the rule performing itself.\n\n" +
      "SAY: You already had the ear for this; today you got the vocabulary and a repeatable procedure. Find the series, name the form, force the match.\n\n" +
      "DO: collect the commitments if you haven't, and point at the bottom strip — this deck was built under the constraint it teaches. Invite them to audit it and tell you where it slipped.\n\n" +
      "EXPECT: someone asks for the slides. Send the deck plus slide sixteen as a one-page job aid.\n\n" +
      "IF TIME IS SHORT: cut everything except the two lines and the three moves."
    );
  }

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
