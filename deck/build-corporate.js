/**
 * "Parallelism (in Grammar)" — corporate training deck.
 *
 * Content source: the presentation of the same name by Iffat Jahan Suchona
 * (Department of English, University of Dhaka). Every definition and every
 * example sentence in this deck is taken from that source; the sequencing,
 * layout, illustrations and speaker notes are the presentation layer.
 *
 * 1920 x 1080 (13.333in x 7.5in). Light corporate system: navy and white,
 * Cambria headings over Calibri body, red = NOT PARALLEL, green = PARALLEL.
 */
const path = require("path");
const PptxGenJS = require("pptxgenjs");
const { C, rails, balance, bridge, icon } = require("./assets-corporate");

const OUT = path.join(__dirname, "build", "parallelism-in-grammar.pptx");

// ---------------------------------------------------------------- geometry
const W = 13.333;
const H = 7.5;
const M = 0.7;
const CW = W - M * 2; // 11.933
const HEAD = "Cambria"; // headings — serif, ships with Office
const BODY = "Calibri"; // body and UI — sans, ships with Office
const FOOT_Y = 6.94;
const TOTAL = 14;
const CONTENTS = 2;

const softShadow = () => ({
  type: "outer", color: "9AA7BC", blur: 14, offset: 3, angle: 90, opacity: 0.28,
});

/** White card with a hairline edge — the deck's single repeating container. */
function card(slide, x, y, w, h, opt = {}) {
  const tone = opt.tone || "plain"; // plain | panel | bad | good | navy
  const fill = {
    plain: C.white, panel: C.panel, bad: C.redTint, good: C.greenTint, navy: C.navy,
  }[tone];
  const line = {
    plain: C.hairline, panel: C.hairline, bad: C.redLine, good: C.greenLine, navy: C.navySoft,
  }[tone];
  slide.addShape("roundRect", {
    x, y, w, h,
    rectRadius: opt.radius || 0.09,
    fill: { color: fill },
    line: { color: line, width: opt.lineW || 1 },
    ...(opt.shadow === false ? {} : { shadow: softShadow() }),
    ...(opt.objectName ? { objectName: opt.objectName } : {}),
  });
}

/** Small capsule label: NOT PARALLEL / PARALLEL / section tags. */
function tag(slide, x, y, w, h, text, colour, opt = {}) {
  slide.addShape("roundRect", {
    x, y, w, h,
    rectRadius: h / 2,
    fill: { color: colour },
    line: { color: colour, width: 1 },
    ...(opt.objectName ? { objectName: opt.objectName } : {}),
  });
  slide.addText(text, {
    x, y, w, h, align: "center", valign: "middle", margin: 0,
    fontFace: BODY, fontSize: opt.fontSize || 10.5, bold: true, charSpacing: 1.4,
    color: C.white,
    ...(opt.textObjectName ? { objectName: opt.textObjectName } : {}),
  });
}

function eyebrow(slide, text, colour) {
  slide.addText(text, {
    x: M, y: 0.5, w: CW, h: 0.26, margin: 0, objectName: "eyebrow",
    fontFace: BODY, fontSize: 11, bold: true, charSpacing: 3.2,
    color: colour || C.ochre,
  });
}

function heading(slide, text, opt = {}) {
  slide.addText(text, {
    x: M, y: opt.y === undefined ? 0.8 : opt.y, w: opt.w || CW, h: opt.h || 0.62, margin: 0,
    fontFace: HEAD, fontSize: opt.fontSize || 32, bold: true,
    color: opt.color || C.navy, valign: "top", objectName: opt.objectName || "heading",
  });
}

function standfirst(slide, text, opt = {}) {
  slide.addText(text, {
    x: M, y: opt.y === undefined ? 1.46 : opt.y, w: opt.w || CW, h: 0.36, margin: 0,
    fontFace: BODY, fontSize: opt.fontSize || 15, color: C.muted, objectName: "standfirst",
  });
}

/** Discreet footer: deck name, contents link, page number, prev/next. */
function footer(slide, idx) {
  slide.addText("Parallelism (in Grammar)", {
    x: M, y: FOOT_Y, w: 4.4, h: 0.28, margin: 0, valign: "middle",
    fontFace: BODY, fontSize: 9.5, color: C.muted, charSpacing: 0.6, objectName: "footDeck",
  });
  slide.addText(
    [
      { text: "CONTENTS", options: { color: C.muted, hyperlink: { slide: CONTENTS }, underline: { style: "none" } } },
      { text: "     ‹ ", options: { color: C.muted, hyperlink: { slide: Math.max(1, idx - 1) }, underline: { style: "none" } } },
      { text: " › ", options: { color: C.muted, hyperlink: { slide: Math.min(TOTAL, idx + 1) }, underline: { style: "none" } } },
      { text: `    ${String(idx).padStart(2, "0")} / ${TOTAL}`, options: { color: C.navy, bold: true } },
    ],
    {
      x: W - M - 4.4, y: FOOT_Y, w: 4.4, h: 0.28, margin: 0, valign: "middle", align: "right",
      fontFace: BODY, fontSize: 9.5, charSpacing: 0.8, objectName: "footNav",
    }
  );
}

/** Not-parallel / parallel example pair, stacked. Source wording verbatim. */
function examplePair(slide, x, y, w, bad, good, opt = {}) {
  const h = opt.h || 1.24;
  const gap = opt.gap || 0.26;
  [
    { label: "NOT PARALLEL", colour: C.red, tone: "bad", text: bad, y },
    { label: "PARALLEL", colour: C.green, tone: "good", text: good, y: y + h + gap },
  ].forEach((row) => {
    card(slide, x, row.y, w, h, { tone: row.tone, radius: 0.1 });
    tag(slide, x + 0.3, row.y + 0.24, 1.62, 0.32, row.label, row.colour, { fontSize: 9.5 });
    slide.addText(row.text, {
      x: x + 0.3, y: row.y + 0.62, w: w - 0.6, h: h - 0.78, margin: 0, valign: "top",
      fontFace: BODY, fontSize: opt.fontSize || 16.5, color: C.navy, lineSpacing: 23,
    });
  });
}

// ------------------------------------------------------- the four contexts
// (source: "Using Parallel Structure" — coordinating conjunctions, correlative
// conjunctions, phrases or clauses of comparison, lists)
const CONTEXTS = [
  {
    n: "01", name: "Coordinating\nconjunctions", flat: "Coordinating conjunctions", short: "Coordinating", ic: "FaPlus",
    cue: "for, and, but, or, yet",
    rule: "While adding two or more clauses or phrases with a coordinating conjunction (for, and, but, or, yet).",
    bad: "My friend took me dancing and to a show.",
    good: "My friend took me to a dance and a show.",
    note: "“Dancing” is a gerund, “to a show” a prepositional phrase. The repair makes both noun phrases after one preposition.",
    notes:
      "SAY: This is the most common context — two things joined by and, but or or.\n\n" +
      "DO: read the not-parallel version aloud, then ask what kind of word each half is. Dancing is a gerund; to a show is a prepositional phrase.\n\n" +
      "SAY: The fix isn't longer, it's matched: to a dance and a show.\n\n" +
      "IF THE ROOM STALLS: cover the second half and ask them to finish the sentence themselves — most people produce the parallel version naturally.",
  },
  {
    n: "02", name: "Correlative\nconjunctions", flat: "Correlative conjunctions", short: "Correlative", ic: "FaLink",
    cue: "not only…but also, either…or, neither…nor, if…then",
    rule: "While connecting two clauses or phrases with a correlative conjunction (not only…but also, either…or, neither…nor, if…then, etc.).",
    bad: "The dog not only likes to play fetch, but also chase cars.",
    good: "The dog likes not only to play fetch, but also to chase cars.",
    note: "Whatever follows the first half of the pair must match what follows the second: to play fetch / to chase cars.",
    notes:
      "SAY: Correlative conjunctions come in pairs, and the pair sets an expectation the reader hears.\n\n" +
      "DO: point at where the pair sits. In the broken version, not only lands before likes, so the two halves can never match. Moving likes in front of not only fixes the frame.\n\n" +
      "SAY: The test is simple — whatever follows the first half must be the same kind of thing that follows the second.\n\n" +
      "EXPECT: this is the one people get wrong most often in writing. Flag it as the highest-value slide of the four.",
  },
  {
    n: "03", name: "Phrases or clauses\nof comparison", flat: "Phrases or clauses of comparison", short: "Comparison", ic: "FaScaleBalanced",
    cue: "than, as",
    rule: "While joining two clauses or phrases with a word of comparison, such as than or as.",
    bad: "I would rather pay for my education than financial aid.",
    good: "I would rather pay for my education than receive financial aid.",
    note: "A comparison must compare like with like: pay … than receive, not pay … than a noun.",
    notes:
      "SAY: A comparison sets up two things that must be the same kind of thing.\n\n" +
      "DO: read the broken version literally — it compares paying with financial aid, which are not the same category. Adding the verb receive restores the match.\n\n" +
      "SAY: Watch for than and as. They are quiet, and they carry a grammatical obligation.\n\n" +
      "IF THE ROOM STALLS: ask what exactly is being compared with what. Saying it out loud exposes the mismatch.",
  },
  {
    n: "04", name: "Lists", flat: "Lists", short: "Lists", ic: "FaListUl",
    cue: "items in a series",
    rule: "While comparing elements in lists.",
    bad: "He made cake, pie and he baked dumplings.",
    good: "He made cake, pie and dumplings.",
    note: "Two nouns followed by a clause is not a list. Once the verb is stated, every item can hang from it.",
    notes:
      "SAY: Lists are where a reader most expects a pattern, so a break is most audible here.\n\n" +
      "DO: read the broken version and let the room hear the third item change shape. Two nouns, then a whole clause.\n\n" +
      "SAY: The verb is already there — he made. Every item can hang from it, so nothing needs repeating.\n\n" +
      "SAY: This applies to bulleted lists in documents and slides exactly as it does to a sentence.",
  },
];

// grid (slide 7) and rail (slides 8-11) geometry — shared objectNames drive Morph
const GRID = { y: 2.06, h: 2.5, w: 2.84, gap: 0.19, x0: M };
const gridX = (i) => GRID.x0 + i * (GRID.w + GRID.gap);
const RAIL = { x: M, w: 3.32, h: 0.96, gap: 0.16, y0: 2.0 };
const railY = (i) => RAIL.y0 + i * (RAIL.h + RAIL.gap);

async function main() {
  const pres = new PptxGenJS();
  pres.defineLayout({ name: "HD", width: W, height: H });
  pres.layout = "HD";
  pres.author = "Parallelism (in Grammar)";
  pres.title = "Parallelism (in Grammar)";
  pres.subject = "Adapted from the presentation by Iffat Jahan Suchona";

  const ART = {
    railsDark: await rails({ dark: true }),
    rails: await rails({ dark: false }),
    balance: await balance(),
    bridge: await bridge(),
  };
  const ic = {};
  for (const c of CONTEXTS) ic[c.ic] = await icon(c.ic, C.white);
  for (const [k, col] of [
    ["FaBookOpen", C.navy], ["FaMagnifyingGlass", C.navy], ["FaTableColumns", C.navy],
    ["FaLayerGroup", C.navy], ["FaCircleCheck", C.green], ["FaBook", C.navy],
  ]) ic[k] = await icon(k, col);

  const light = () => {
    const s = pres.addSlide();
    s.background = { color: C.white };
    return s;
  };
  const dark = () => {
    const s = pres.addSlide();
    s.background = { color: C.navy };
    return s;
  };

  // ============================================================== 1 TITLE
  {
    const s = dark();
    s.addText("TRAINING SESSION", {
      x: M, y: 1.5, w: 6.6, h: 0.3, margin: 0,
      fontFace: BODY, fontSize: 11.5, bold: true, charSpacing: 3.4, color: C.ochre,
    });
    s.addText("Parallelism", {
      x: M, y: 1.92, w: 7.0, h: 1.0, margin: 0,
      fontFace: HEAD, fontSize: 60, bold: true, color: C.white,
    });
    s.addText("(in Grammar)", {
      x: M, y: 2.92, w: 7.0, h: 0.7, margin: 0,
      fontFace: HEAD, fontSize: 34, italic: true, color: C.ochre,
    });
    s.addText(
      "The same grammatical structure within one or more sentences of similar phrases or clauses.",
      { x: M, y: 3.78, w: 6.5, h: 0.7, margin: 0, fontFace: BODY, fontSize: 15, color: "C7CFDD", lineSpacing: 22 }
    );

    card(s, M, 4.6, 6.5, 2.04, { tone: "navy", radius: 0.1, shadow: false });
    s.addText("PRESENTED BY", {
      x: M + 0.36, y: 4.84, w: 5.6, h: 0.26, margin: 0,
      fontFace: BODY, fontSize: 9.5, bold: true, charSpacing: 2.6, color: C.ochre,
    });
    s.addText("Iffat Jahan Suchona", {
      x: M + 0.36, y: 5.14, w: 5.8, h: 0.36, margin: 0,
      fontFace: HEAD, fontSize: 20, bold: true, color: C.white,
    });
    s.addText(
      "B.A (Honours) · M.A (Applied Linguistics and E.L.T)\nDepartment of English, University of Dhaka",
      { x: M + 0.36, y: 5.56, w: 5.8, h: 0.6, margin: 0, fontFace: BODY, fontSize: 12.5, color: "AEB9CC", lineSpacing: 18 }
    );
    s.addText("Contact: 01681444197  ·  iffatsuchona@gmail.com", {
      x: M + 0.36, y: 6.14, w: 5.8, h: 0.3, margin: 0,
      fontFace: BODY, fontSize: 12, color: C.ochre,
    });

    s.addImage({ data: ART.railsDark, x: 7.52, y: 2.46, w: 5.18, h: 2.24 });
    s.addText("A 14-slide working session on parallel structure", {
      x: 7.52, y: 4.94, w: 5.18, h: 0.3, margin: 0, align: "right",
      fontFace: BODY, fontSize: 11, color: "8794AB", charSpacing: 0.8,
    });
    s.addNotes(
      "SAY: Welcome. This session covers one idea — parallel structure — and four places you need it.\n\n" +
      "SAY: The material is adapted from Iffat Jahan Suchona's presentation on parallelism; every example you'll see comes from that source.\n\n" +
      "DO: set the expectation early — this is a working session, not a lecture. There are two before-and-after examples on almost every slide and I'll ask you to correct some of them out loud.\n\n" +
      "IF THE ROOM IS QUIET: ask who has had a sentence returned by a reviewer marked 'awkward' with no further explanation. Most hands go up; that is usually parallelism."
    );
  }

  // ============================================================ 2 CONTENTS
  {
    const s = light();
    eyebrow(s, "CONTENTS");
    heading(s, "What this session covers");
    standfirst(s, "One definition, one worked example, four contexts, and the references behind them.");

    const items = [
      ["FaBookOpen", "Parallelism is…", "The definition, and the balance it describes", "03 – 04"],
      ["FaMagnifyingGlass", "A worked example", "One sentence, taken apart and repaired", "05"],
      ["FaTableColumns", "Not parallel vs parallel", "Four short pairs, side by side", "06"],
      ["FaLayerGroup", "Using parallel structure", "The four contexts that demand it", "07 – 11"],
      ["FaCircleCheck", "Summary", "The rule, the four contexts, the test", "12"],
      ["FaBook", "References", "Where the material comes from", "13"],
    ];
    const cw = (CW - 0.4) / 2;
    items.forEach(([iconName, title, sub, pages], i) => {
      const x = M + (i % 2) * (cw + 0.4);
      const y = 2.06 + Math.floor(i / 2) * 1.5;
      card(s, x, y, cw, 1.28, { tone: i % 2 === 0 ? "plain" : "panel" });
      s.addShape("ellipse", {
        x: x + 0.34, y: y + 0.34, w: 0.6, h: 0.6,
        fill: { color: C.ochreTint }, line: { color: C.ochre, width: 1 },
      });
      s.addImage({ data: ic[iconName], x: x + 0.48, y: y + 0.48, w: 0.32, h: 0.32 });
      s.addText(title, {
        x: x + 1.14, y: y + 0.3, w: cw - 2.2, h: 0.36, margin: 0,
        fontFace: HEAD, fontSize: 18, bold: true, color: C.navy,
      });
      s.addText(sub, {
        x: x + 1.14, y: y + 0.7, w: cw - 2.0, h: 0.34, margin: 0,
        fontFace: BODY, fontSize: 12.5, color: C.muted,
      });
      s.addText(pages, {
        x: x + cw - 1.0, y: y + 0.3, w: 0.7, h: 0.34, margin: 0, align: "right",
        fontFace: BODY, fontSize: 12, bold: true, color: C.ochre,
      });
    });
    footer(s, 2);
    s.addNotes(
      "SAY: Here is the shape of the next forty minutes. Definition first, then one sentence taken apart, then four contexts.\n\n" +
      "DO: point out that the four contexts in the middle are the operational part — that's what people take back to their desks.\n\n" +
      "SAY: Every page number on this slide is clickable, so ask me to jump back at any point.\n\n" +
      "IF TIME IS SHORT: the four context slides and the summary are the core; the worked example can be cut."
    );
  }

  // =============================================================== 3 HOOK
  {
    const s = light();
    eyebrow(s, "STARTING POINT");
    heading(s, "“Parallelism… in grammar?”");
    standfirst(s, "The word comes from the railway: two rails, the same distance apart, all the way along.");

    s.addImage({ data: ART.rails, x: M, y: 2.16, w: 7.4, h: 3.21 });

    card(s, 8.4, 2.16, CW - 7.7, 3.21, { tone: "panel" });
    s.addText("What people expect", {
      x: 8.74, y: 2.5, w: 3.6, h: 0.36, margin: 0,
      fontFace: HEAD, fontSize: 19, bold: true, color: C.navy,
    });
    s.addText(
      "Something mathematical, or something about train tracks.",
      { x: 8.74, y: 2.92, w: 3.62, h: 0.62, margin: 0, fontFace: BODY, fontSize: 13.5, color: C.slate, lineSpacing: 19 }
    );
    s.addText("What it actually is", {
      x: 8.74, y: 3.74, w: 3.6, h: 0.36, margin: 0,
      fontFace: HEAD, fontSize: 19, bold: true, color: C.navy,
    });
    s.addText(
      "A balance inside the sentence. Items doing the same job are written in the same form — and the reader feels it immediately.",
      { x: 8.74, y: 4.16, w: 3.62, h: 1.0, margin: 0, fontFace: BODY, fontSize: 13.5, color: C.slate, lineSpacing: 19 }
    );
    card(s, M, 5.74, CW, 0.86, { tone: "panel" });
    s.addText(
      [
        { text: "The source deck opens on exactly this question:  ", options: { color: C.slate } },
        { text: "“umm… Parallelism in Grammar????”", options: { color: C.navy, bold: true, italic: true } },
      ],
      { x: M + 0.4, y: 5.74, w: CW - 0.8, h: 0.86, margin: 0, valign: "middle",
        fontFace: BODY, fontSize: 14 }
    );
    footer(s, 3);
    s.addNotes(
      "DO: put this up and ask the room what the word parallelism makes them think of. You will get geometry and railway lines.\n\n" +
      "SAY: Hold on to the railway picture — it is exactly right. Two rails, the same shape, the same distance apart, for as long as the track runs.\n\n" +
      "SAY: In a sentence, the rails are the items doing the same job. When one of them changes shape, the reader feels the bump even if they can't name it.\n\n" +
      "IF THE ROOM STALLS: ask whether they have ever read a sentence twice without knowing why. That bump is usually this."
    );
  }

  // ========================================================= 4 DEFINITION
  {
    const s = light();
    eyebrow(s, "DEFINITION");
    heading(s, "Parallelism is…");

    card(s, M, 1.86, 6.1, 3.5, { tone: "plain" });
    const defs = [
      "having the same grammatical structure within one or more sentences of similar phrases or clauses",
      "also known as parallel structure or parallel construction",
    ];
    defs.forEach((t, i) => {
      const y = 2.24 + i * 1.44;
      s.addShape("ellipse", {
        x: M + 0.4, y: y + 0.08, w: 0.34, h: 0.34,
        fill: { color: C.navy }, line: { color: C.navy, width: 1 },
      });
      s.addText(String(i + 1), {
        x: M + 0.4, y: y + 0.08, w: 0.34, h: 0.34, align: "center", valign: "middle", margin: 0,
        fontFace: BODY, fontSize: 11, bold: true, color: C.white,
      });
      s.addText(t, {
        x: M + 0.96, y, w: 4.86, h: 1.1, margin: 0,
        fontFace: BODY, fontSize: 16.5, color: C.navy, lineSpacing: 25,
      });
    });

    card(s, M, 5.58, 6.1, 0.92, { tone: "good" });
    s.addImage({ data: ic.FaCircleCheck, x: M + 0.32, y: 5.86, w: 0.34, h: 0.34 });
    s.addText("Similar phrases or clauses, written in the same pattern of words.", {
      x: M + 0.84, y: 5.58, w: 5.0, h: 0.92, margin: 0, valign: "middle",
      fontFace: BODY, fontSize: 14, bold: true, color: C.green,
    });

    s.addImage({ data: ART.balance, x: 7.3, y: 1.92, w: 5.2, h: 3.43 });
    s.addText(
      "Both sides of the sentence carry the same pattern of words — that equality is the whole idea.",
      { x: 7.3, y: 5.52, w: 5.2, h: 0.66, margin: 0, align: "center",
        fontFace: BODY, fontSize: 13, italic: true, color: C.muted, lineSpacing: 19 }
    );
    footer(s, 4);
    s.addNotes(
      "SAY: Read the definition slowly. The key phrase is 'the same grammatical structure' — not the same words, the same structure.\n\n" +
      "SAY: You will also meet it as parallel structure or parallel construction. Same thing, three names.\n\n" +
      "DO: point at the balance. Both pans carry the same pattern of words. If one side changes form, the sentence tips.\n\n" +
      "EXPECT: someone asks whether this is a rule or a preference. Answer: with conjunctions and comparisons it is a rule of grammar; in lists and headings it is a rule of clarity."
    );
  }

  // ==================================================== 5 WORKED EXAMPLE
  {
    const s = light();
    eyebrow(s, "WORKED EXAMPLE");
    heading(s, "One sentence, taken apart");
    standfirst(s, "Both halves answer “how did she speak?” — so both halves must take the same form.");

    // broken
    card(s, M, 2.06, CW, 1.94, { tone: "bad" });
    tag(s, M + 0.34, 2.3, 1.62, 0.32, "NOT PARALLEL", C.red, { fontSize: 9.5 });
    s.addText(
      [
        { text: "When talking to the group, she spoke ", options: { color: C.navy } },
        { text: "sincerely", options: { color: C.red, bold: true } },
        { text: " and ", options: { color: C.navy } },
        { text: "with passion", options: { color: C.red, bold: true } },
      ],
      { x: M + 0.34, y: 2.72, w: CW - 0.68, h: 0.5, margin: 0, fontFace: BODY, fontSize: 21 }
    );
    [
      { centre: 7.32, label: "adverb", w: 1.24 },
      { centre: 9.84, label: "prepositional phrase", w: 2.1 },
    ].forEach((a) => {
      s.addShape("line", {
        x: a.centre, y: 3.2, w: 0, h: 0.26, line: { color: C.red, width: 1.25, dashType: "dash" },
      });
      s.addShape("roundRect", {
        x: a.centre - a.w / 2, y: 3.46, w: a.w, h: 0.34, rectRadius: 0.17,
        fill: { color: C.white }, line: { color: C.redLine, width: 1 },
      });
      s.addText(a.label, {
        x: a.centre - a.w / 2, y: 3.46, w: a.w, h: 0.34, align: "center", valign: "middle", margin: 0,
        fontFace: BODY, fontSize: 10.5, bold: true, color: C.red,
      });
    });

    // repaired
    card(s, M, 4.28, CW, 1.94, { tone: "good" });
    tag(s, M + 0.34, 4.52, 1.34, 0.32, "PARALLEL", C.green, { fontSize: 9.5 });
    s.addText(
      [
        { text: "When talking to the group, she spoke ", options: { color: C.navy } },
        { text: "sincerely", options: { color: C.green, bold: true } },
        { text: " and ", options: { color: C.navy } },
        { text: "passionately", options: { color: C.green, bold: true } },
      ],
      { x: M + 0.34, y: 4.94, w: CW - 0.68, h: 0.5, margin: 0, fontFace: BODY, fontSize: 21 }
    );
    [7.33, 9.85].forEach((x) => {
      s.addShape("line", {
        x, y: 5.42, w: 0, h: 0.22, line: { color: C.green, width: 1.25, dashType: "dash" },
      });
    });
    s.addShape("line", {
      x: 7.33, y: 5.64, w: 2.52, h: 0, line: { color: C.green, width: 1.25, dashType: "dash" },
    });
    s.addShape("roundRect", {
      x: 8.59 - 0.95, y: 5.74, w: 1.9, h: 0.34, rectRadius: 0.17,
      fill: { color: C.white }, line: { color: C.greenLine, width: 1 },
    });
    s.addText("two adverbs", {
      x: 8.59 - 0.95, y: 5.74, w: 1.9, h: 0.34, align: "center", valign: "middle", margin: 0,
      fontFace: BODY, fontSize: 10.5, bold: true, color: C.green,
    });
    footer(s, 5);
    s.addNotes(
      "DO: read the first version aloud, then stop at 'and'. Ask what kind of word sincerely is, and what with passion is.\n\n" +
      "SAY: One is an adverb, the other a prepositional phrase. Both answer the same question — how did she speak? — so the reader expects them to look alike.\n\n" +
      "SAY: The repair is a single word: passionately. Nothing was added, nothing was cut, and the sentence now reads in one movement.\n\n" +
      "EXPECT: someone says the first version sounds fine. Agree that it is understandable, and point out that understandable and well-formed are different standards."
    );
  }

  // ================================================== 6 MORE EXAMPLES
  {
    const s = light();
    eyebrow(s, "MORE EXAMPLES");
    heading(s, "Not parallel, then parallel");
    standfirst(s, "Four sentences from the source, each broken in a different way.");

    const rows = [
      ["She likes jogging, and to read", "She likes jogging and reading."],
      ["The show is both enjoyable and it is educational", "The show is both enjoyable and educational"],
      ["He not only wants money but also fame", "He not only wants money but also wants fame"],
      ["Trying desperately to find his keys, he looked in his bedroom, the sofa and on the kitchen counter.",
        "Trying desperately to find his keys, he looked in his bedroom, under the sofa and on the kitchen counter."],
    ];
    const colW = (CW - 0.34) / 2;
    tag(s, M, 1.96, 1.72, 0.34, "NOT PARALLEL", C.red, { fontSize: 10 });
    tag(s, M + colW + 0.34, 1.96, 1.44, 0.34, "PARALLEL", C.green, { fontSize: 10 });

    let y = 2.44;
    rows.forEach(([bad, good], i) => {
      const h = i === 3 ? 1.1 : 0.86;
      card(s, M, y, colW, h, { tone: "bad", radius: 0.08, shadow: false });
      card(s, M + colW + 0.34, y, colW, h, { tone: "good", radius: 0.08, shadow: false });
      s.addText(bad, {
        x: M + 0.28, y: y + 0.14, w: colW - 0.56, h: h - 0.28, margin: 0, valign: "middle",
        fontFace: BODY, fontSize: 13.5, color: C.navy, lineSpacing: 19,
      });
      s.addText(good, {
        x: M + colW + 0.62, y: y + 0.14, w: colW - 0.56, h: h - 0.28, margin: 0, valign: "middle",
        fontFace: BODY, fontSize: 13.5, color: C.navy, lineSpacing: 19,
      });
      y += h + 0.13;
    });
    footer(s, 6);
    s.addNotes(
      "DO: work down the rows one at a time, covering the right-hand column until the room has offered a fix.\n\n" +
      "ROW 1: a gerund beside an infinitive — jogging and to read. ROW 2: an adjective beside a full clause. ROW 3: the correlative pair needs the verb repeated — wants money, wants fame. ROW 4: two prepositional phrases and a bare noun; under the sofa restores the pattern.\n\n" +
      "SAY: Notice that three of the four repairs are shorter or the same length. Parallelism is not padding.\n\n" +
      "IF THE ROOM IS CONFIDENT: skip straight to row four, which is the one people most often miss."
    );
  }

  // ============================================ 7 USING PARALLEL STRUCTURE
  {
    const s = light();
    eyebrow(s, "USING PARALLEL STRUCTURE");
    heading(s, "Four contexts that demand it");
    standfirst(s, "Whenever two or more elements are joined, the join sets an expectation about their form.");

    CONTEXTS.forEach((c, i) => {
      const x = gridX(i);
      card(s, x, GRID.y, GRID.w, GRID.h, { tone: "plain", objectName: `ctxCard${i + 1}` });
      s.addShape("ellipse", {
        x: x + 0.3, y: GRID.y + 0.32, w: 0.62, h: 0.62,
        fill: { color: C.navy }, line: { color: C.navy, width: 1 },
        objectName: `ctxDisc${i + 1}`,
      });
      s.addImage({ data: ic[c.ic], x: x + 0.45, y: GRID.y + 0.47, w: 0.32, h: 0.32, objectName: `ctxIcon${i + 1}` });
      s.addText(c.n, {
        x: x + GRID.w - 0.92, y: GRID.y + 0.34, w: 0.62, h: 0.32, margin: 0, align: "right",
        fontFace: BODY, fontSize: 13, bold: true, color: C.ochre, objectName: `ctxNum${i + 1}`,
      });
      s.addText(c.name, {
        x: x + 0.3, y: GRID.y + 1.06, w: GRID.w - 0.42, h: 0.8, margin: 0,
        fontFace: HEAD, fontSize: 15.5, bold: true, color: C.navy, lineSpacing: 21,
        objectName: `ctxName${i + 1}`, hyperlink: { slide: 8 + i }, underline: { style: "none" },
      });
      s.addText(c.cue, {
        x: x + 0.3, y: GRID.y + 1.92, w: GRID.w - 0.6, h: 0.44, margin: 0,
        fontFace: BODY, fontSize: 11.5, italic: true, color: C.muted, lineSpacing: 15,
        objectName: `ctxCue${i + 1}`,
      });
    });
    s.addImage({ data: ART.bridge, x: 3.17, y: 5.02, w: 7.0, h: 1.75 });
    footer(s, 7);
    s.addNotes(
      "SAY: Parallelism is not required everywhere — it is required wherever elements are joined. These are the four joins.\n\n" +
      "DO: read the four out loud with their signal words: and, not only…but also, than, and the comma series. Those signal words are what people should learn to notice.\n\n" +
      "SAY: The next four slides take one context each, with the same layout every time: the rule, then a not-parallel and a parallel version.\n\n" +
      "DO: the four cards are clickable if anyone wants to jump straight to one."
    );
  }

  // ========================================== 8-11 THE FOUR CONTEXTS
  CONTEXTS.forEach((c, idx) => {
    const s = light();
    eyebrow(s, `USING PARALLEL STRUCTURE  ·  ${c.n} OF 04`);
    heading(s, c.flat);

    // persistent rail — same objectNames as the grid, so PowerPoint morphs them
    CONTEXTS.forEach((q, j) => {
      const active = j === idx;
      const y = railY(j);
      card(s, RAIL.x, y, RAIL.w, RAIL.h, {
        tone: active ? "navy" : "panel", radius: 0.08, shadow: active,
        objectName: `ctxCard${j + 1}`,
      });
      s.addShape("ellipse", {
        x: RAIL.x + 0.26, y: y + 0.26, w: 0.44, h: 0.44,
        fill: { color: active ? C.ochre : C.hairline },
        line: { color: active ? C.ochre : C.hairline, width: 1 },
        objectName: `ctxDisc${j + 1}`,
      });
      s.addImage({
        data: ic[q.ic], x: RAIL.x + 0.36, y: y + 0.36, w: 0.24, h: 0.24,
        transparency: active ? 0 : 55, objectName: `ctxIcon${j + 1}`,
      });
      s.addText(q.n, {
        x: RAIL.x + RAIL.w - 0.72, y: y + 0.32, w: 0.44, h: 0.3, margin: 0, align: "right",
        fontFace: BODY, fontSize: 11, bold: true, color: active ? C.ochre : C.muted,
        objectName: `ctxNum${j + 1}`,
      });
      s.addText(q.short, {
        x: RAIL.x + 0.84, y: y + 0.3, w: RAIL.w - 1.5, h: 0.4, margin: 0, valign: "middle",
        fontFace: BODY, fontSize: 13, bold: active, color: active ? C.white : C.slate,
        objectName: `ctxName${j + 1}`, hyperlink: { slide: 8 + j }, underline: { style: "none" },
      });
    });

    const px = M + RAIL.w + 0.42;
    const pw = W - M - px;

    // the rule, quoted from the source
    card(s, px, 1.44, pw, 1.16, { tone: "panel" });
    s.addText("THE RULE", {
      x: px + 0.34, y: 1.62, w: 2.0, h: 0.26, margin: 0,
      fontFace: BODY, fontSize: 9.5, bold: true, charSpacing: 2.4, color: C.ochre,
    });
    s.addText(c.rule, {
      x: px + 0.34, y: 1.9, w: pw - 0.68, h: 0.6, margin: 0,
      fontFace: BODY, fontSize: 14.5, color: C.navy, lineSpacing: 21,
    });

    examplePair(s, px, 2.86, pw, c.bad, c.good, { h: 1.3, gap: 0.24 });

    card(s, px, 5.74, pw, 0.86, { tone: "plain" });
    s.addText("Why", {
      x: px + 0.34, y: 5.74, w: 0.6, h: 0.86, margin: 0, valign: "middle",
      fontFace: HEAD, fontSize: 14, bold: true, color: C.ochre,
    });
    s.addText(c.note, {
      x: px + 1.0, y: 5.74, w: pw - 1.34, h: 0.86, margin: 0, valign: "middle",
      fontFace: BODY, fontSize: 12.5, color: C.slate, lineSpacing: 17,
    });

    footer(s, 8 + idx);
    s.addNotes(c.notes);
  });

  // ============================================================ 12 SUMMARY
  {
    const s = light();
    eyebrow(s, "SUMMARY");
    heading(s, "One rule, four contexts, one test");
    standfirst(s, "Everything in this session, on a single page.");

    card(s, M, 2.0, CW, 1.06, { tone: "good" });
    s.addText("THE RULE", {
      x: M + 0.34, y: 2.2, w: 1.6, h: 0.26, margin: 0,
      fontFace: BODY, fontSize: 9.5, bold: true, charSpacing: 2.4, color: C.green,
    });
    s.addText(
      "Similar phrases or clauses take the same grammatical structure — the same pattern of words on both sides.",
      { x: M + 0.34, y: 2.48, w: CW - 0.68, h: 0.44, margin: 0,
        fontFace: BODY, fontSize: 16, bold: true, color: C.navy }
    );

    const cw = (CW - 0.57) / 4;
    CONTEXTS.forEach((c, i) => {
      const x = M + i * (cw + 0.19);
      // same objectNames as the grid and the rail: the rail folds back to here
      card(s, x, 3.32, cw, 1.94, { tone: "plain", objectName: `ctxCard${i + 1}` });
      s.addShape("ellipse", {
        x: x + 0.28, y: 3.5, w: 0.44, h: 0.44,
        fill: { color: C.navy }, line: { color: C.navy, width: 1 },
        objectName: `ctxDisc${i + 1}`,
      });
      s.addImage({ data: ic[c.ic], x: x + 0.38, y: 3.6, w: 0.24, h: 0.24, objectName: `ctxIcon${i + 1}` });
      s.addText(c.n, {
        x: x + 0.84, y: 3.58, w: 0.6, h: 0.3, margin: 0,
        fontFace: BODY, fontSize: 12, bold: true, color: C.ochre, objectName: `ctxNum${i + 1}`,
      });
      s.addText(c.name, {
        x: x + 0.28, y: 4.02, w: cw - 0.34, h: 0.72, margin: 0,
        fontFace: HEAD, fontSize: 15, bold: true, color: C.navy, lineSpacing: 19,
        objectName: `ctxName${i + 1}`, hyperlink: { slide: 8 + i }, underline: { style: "none" },
      });
      s.addText(c.cue, {
        x: x + 0.28, y: 4.82, w: cw - 0.44, h: 0.44, margin: 0,
        fontFace: BODY, fontSize: 10.5, italic: true, color: C.muted, lineSpacing: 14,
        objectName: `ctxCue${i + 1}`,
      });
    });

    card(s, M, 5.5, CW, 1.06, { tone: "panel" });
    s.addText("THE TEST", {
      x: M + 0.34, y: 5.7, w: 1.6, h: 0.26, margin: 0,
      fontFace: BODY, fontSize: 9.5, bold: true, charSpacing: 2.4, color: C.ochre,
    });
    s.addText(
      "Read the joined elements one after another. If they do not take the same form, rewrite the shorter one to match — never the reader’s expectation.",
      { x: M + 0.34, y: 5.98, w: CW - 0.68, h: 0.44, margin: 0,
        fontFace: BODY, fontSize: 14.5, color: C.navy }
    );
    footer(s, 12);
    s.addNotes(
      "SAY: If you keep one page from today, keep this one.\n\n" +
      "DO: read the rule, then name the four contexts, then the test. Thirty seconds, no examples — this is the recall version.\n\n" +
      "DO: ask each person for one place in their own writing where they will apply the test this week. Say it out loud, keep it short.\n\n" +
      "IF TIME ALLOWS: take a sentence from the room and run the test on it live."
    );
  }

  // ========================================================= 13 REFERENCES
  {
    const s = light();
    eyebrow(s, "REFERENCES");
    heading(s, "Where this material comes from");
    standfirst(s, "As cited in the source presentation.");

    const refs = [
      ["Parallelism (n.d.). In Grammarly Blog.", "https://www.grammarly.com/blog/parallelism/"],
      ["Macleod, S. (2002, November 15). Parallel Structure or Parallelism [Video file].", "https://www.youtube.com/watch?v=qvDNvS2M3QA"],
      ["Evergreen Writing Center.", "https://www.evergreen.edu/sites/default/files/writingcenter/handouts/grammar/parallel.pdf"],
    ];
    refs.forEach(([text, url], i) => {
      const y = 2.14 + i * 1.24;
      card(s, M, y, CW, 1.04, { tone: i % 2 === 0 ? "plain" : "panel" });
      s.addShape("ellipse", {
        x: M + 0.34, y: y + 0.32, w: 0.4, h: 0.4,
        fill: { color: C.navy }, line: { color: C.navy, width: 1 },
      });
      s.addText(String(i + 1), {
        x: M + 0.34, y: y + 0.32, w: 0.4, h: 0.4, align: "center", valign: "middle", margin: 0,
        fontFace: BODY, fontSize: 12, bold: true, color: C.white,
      });
      s.addText(text, {
        x: M + 0.98, y: y + 0.2, w: CW - 1.4, h: 0.34, margin: 0,
        fontFace: BODY, fontSize: 14.5, color: C.navy,
      });
      s.addText(url, {
        x: M + 0.98, y: y + 0.56, w: CW - 1.4, h: 0.32, margin: 0,
        fontFace: BODY, fontSize: 12, color: C.ochre,
        hyperlink: { url }, underline: { style: "none" },
      });
    });

    card(s, M, 5.86, CW, 0.72, { tone: "panel" });
    s.addText(
      "Source presentation: “Parallelism (in Grammar)” — Iffat Jahan Suchona, Department of English, University of Dhaka.",
      { x: M + 0.34, y: 5.86, w: CW - 0.68, h: 0.72, margin: 0, valign: "middle",
        fontFace: BODY, fontSize: 13, italic: true, color: C.slate }
    );
    footer(s, 13);
    s.addNotes(
      "SAY: Three sources sit behind this material, and all three are worth ten minutes each.\n\n" +
      "DO: point at the Evergreen handout in particular — it is a one-page reference people can keep on a desk.\n\n" +
      "SAY: The links are live in the file, so take the deck rather than photographing the slide.\n\n" +
      "DO: credit the source presentation explicitly when you re-use this deck internally."
    );
  }

  // ============================================================== 14 CLOSE
  {
    const s = dark();
    s.addText("THANK YOU", {
      x: M, y: 2.36, w: 8.0, h: 0.4, margin: 0,
      fontFace: BODY, fontSize: 12, bold: true, charSpacing: 4, color: C.ochre,
    });
    s.addText("Same job, same form.", {
      x: M, y: 2.9, w: 7.2, h: 0.86, margin: 0,
      fontFace: HEAD, fontSize: 40, bold: true, color: C.white,
    });
    s.addText(
      "Similar phrases or clauses, written in the same pattern of words — in a sentence, in a list, in every document you send.",
      { x: M, y: 3.92, w: 7.4, h: 0.9, margin: 0, fontFace: BODY, fontSize: 15.5, color: "C7CFDD", lineSpacing: 23 }
    );

    card(s, M, 5.02, 7.4, 1.06, { tone: "navy", radius: 0.1, shadow: false });
    s.addText(
      [
        { text: "Questions   ·   ", options: { color: "8794AB" } },
        { text: "iffatsuchona@gmail.com", options: { color: C.ochre, bold: true } },
        { text: "   ·   01681444197", options: { color: "8794AB" } },
      ],
      { x: M + 0.36, y: 5.02, w: 6.8, h: 1.06, margin: 0, valign: "middle",
        fontFace: BODY, fontSize: 14 }
    );

    s.addImage({ data: ART.railsDark, x: 8.62, y: 2.78, w: 4.08, h: 1.77 });
    s.addText("Iffat Jahan Suchona  ·  Department of English, University of Dhaka", {
      x: 8.62, y: 4.76, w: 4.08, h: 0.6, margin: 0, align: "right",
      fontFace: BODY, fontSize: 11.5, color: "8794AB", lineSpacing: 17,
    });
    s.addNotes(
      "SAY: One line to leave with — same job, same form. If two elements do the same work in a sentence, write them the same way.\n\n" +
      "DO: invite questions here rather than earlier; most people surface their own examples at this point, which is the best possible ending.\n\n" +
      "DO: share the deck and point at the summary slide and the Evergreen handout as the two things to keep.\n\n" +
      "IF NO ONE ASKS: offer one of your own — read out a broken sentence from a real document and repair it together."
    );
  }

  await pres.writeFile({ fileName: OUT });
  console.log("wrote", OUT);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
