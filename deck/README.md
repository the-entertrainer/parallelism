# Parallelism — PowerPoint decks

Two decks live here.

| File | What it is |
|---|---|
| `build/parallelism-in-grammar.pptx` | **Primary.** 14 slides, light corporate style, rebuilt strictly from the source presentation by Iffat Jahan Suchona (see below). Generator: `build-corporate.js` + `assets-corporate.js`. |
| `../parallelism_deck.html` | Interactive 16:9 HTML deck — 7 slides, dark Tailwind theme, keyboard-driven, same source content. See below. |
| `build/parallelism-training.pptx` | Earlier 21-slide facilitator deck on parallelism for instructional designers, dark liquid-glass style. Generator: `build.js` + `assets.js`. |

---

## 1. Parallelism (in Grammar) — corporate deck

Content source: **“Parallelism (in Grammar)”, Iffat Jahan Suchona, Department of
English, University of Dhaka** (11 slides, supplied as PDF). Every definition and
every example sentence is taken verbatim from that source, including the four
contexts, the four not-parallel/parallel pairs, the worked "sincerely / with
passion" example, and the three references. Sequencing, layout, illustrations and
speaker notes are the presentation layer.

```bash
node deck/build-corporate.js
python3 deck/inject_morph.py deck/build/parallelism-in-grammar.pptx
python3 deck/render.py    deck/build/parallelism-in-grammar.pptx -o deck/build/render-corp
python3 /root/.claude/skills/pptx/scripts/office/validate.py deck/build/parallelism-in-grammar.pptx
```

**Design system**

| Token | Value | Means |
|---|---|---|
| navy | `16233A` | dominant — headings, title and closing slides |
| ochre | `B4832B` | accent — eyebrows, numbers, emphasis |
| red / tint | `B3261E` / `FBEDEB` | **NOT PARALLEL** |
| green / tint | `1E7A4F` / `EAF4EF` | **PARALLEL** |
| panel / hairline | `F6F8FB` / `DFE5EE` | card surfaces and edges |

Typography is **Cambria** headings over **Calibri** body — both ship with Office
and both render true-to-width in QA. Structure is a light sandwich: navy title and
closing slides, white content slides. Illustrations are generated line art
(`assets-corporate.js`): parallel railway track and a balance beam, both echoing
the source deck's own images, plus a "join" diagram for the four contexts.

**Morph:** slides 8–12 use `byObject` — the four-context grid on slide 7 collapses
into the persistent rail on 8–11, and folds back into the summary row on 12.
Everything else fades.

---

## 1b. parallelism_deck.html — interactive web deck

Same source content, delivered as a single self-contained HTML file at the repo
root. Open it directly in a browser, or serve the folder.

- **7 slides:** interactive “Spot the brain glitch” hook · visual formula ·
  four before/after breakdowns (coordinating, correlative, comparison, lists) ·
  three-question lightning round.
- **Theme:** slate-950 ground, white type, emerald-400 = parallel,
  rose-500 = not parallel. Bento cards, split-screen comparisons and hero
  formula cards, with the layout mirrored on alternate breakdown slides.
- **Controls:** ← / → (also PageUp/PageDown, Space), Home / End, `Enter` to
  reveal the fix on a breakdown slide, `N` to toggle presenter notes. Progress
  bar, `03 / 07` counter and deep links (`#slide-4`).
- **Tailwind** loads from the CDN, so first load needs a connection.

```bash
python3 -m http.server 8123          # then open /parallelism_deck.html
curl -sL https://cdn.tailwindcss.com -o /tmp/tw.js   # cache for the test run
node deck/test-html-deck.js          # 45 assertions + screenshots
```

`deck/test-html-deck.js` drives the deck in Chromium: navigation, every
interaction, computed visibility, overflow inside the 16:9 frame, and stage
scaling at 1280×720. It writes screenshots to `deck/build/html-qa/`.

---

## 2. Parallelism — facilitator deck (earlier)

A 21-slide, ~60-minute working session on grammatical parallelism for
instructional designers and content developers. Companion to the interactive web
app in the repository root.

`build/parallelism-training.pptx` — 1920 × 1080 (13.333in × 7.5in), speaker notes
on all 21 slides, on-screen navigation, Morph transitions. Its content is original,
not drawn from the source presentation above.

## Build

```bash
npm install                                   # pptxgenjs, sharp, react-icons
node deck/build.js                            # -> deck/build/parallelism-training.pptx
python3 deck/inject_morph.py deck/build/parallelism-training.pptx
python3 deck/render.py    deck/build/parallelism-training.pptx   # PNG per slide, visual QA
python3 /root/.claude/skills/pptx/scripts/office/validate.py deck/build/parallelism-training.pptx
```

`inject_morph.py` must run after every `build.js`, and always before shipping:
the generator cannot write transitions, so a fresh build has none.

## Design system

| Token | Value | Means |
|---|---|---|
| rose | `FF4F6E` | **breaks** — mismatched form |
| mint | `2FE0A5` | **holds** — parallel form |
| iris | `8A7BFF` | structure, current item, facilitator instruction |
| ink | `05070F`–`0B1020` | dark mesh-gradient ground (generated PNG; pptxgenjs has no gradient fills) |

One typeface throughout (Calibri). One motif: the liquid-glass rounded panel,
plus the pill triad (three pills — ragged in rose, matched in mint) that states
the deck's own rule on the title and closing slides. The deck is built under the
constraint it teaches.

## Morph map

Transitions live on the slide being *entered*. Morph matches shapes by **name**
(`<p:cNvPr name>`, set through pptxgenjs `objectName`), never by position — so the
repeating cards share names across slides:

| Entering slide | Option | What morphs |
|---|---|---|
| 3 | `byWord` | broken Caesar line re-assembles word by word into the real one (`quoteText1`) |
| 5 | `byWord` | broken Kennedy line re-assembles into the real one (`quoteText2`) |
| 11 | `byObject` | five-card grid (slide 10) collapses into the persistent left rail (`placeCard1–5`, `placeIcon`, `placeName`, `placeNum`, `placeDisc`) |
| 12, 13, 14, 15 | `byObject` | the iris "current" highlight slides down the rail; the before/after panes swap (`paneBREAKS`, `paneHOLDS`) |
| 20 | `byObject` | the audit rail on slide 19 unfolds back into the opening grid — the bookend |
| all others | `fade` | quiet default so the deck feels of one piece |

Slide 19's five audit checkpoints deliberately *are* the five places, carrying the
same `objectName`s, which is what gives slide 20's fold-back something to morph.

`inject_morph.py` also namespaces pptxgenjs's auto-generated shape names
(`Shape 12`, `Text 7`, …) per slide. Those names collide across slides by
accident, and since Morph matches on name, a collision would tween two unrelated
objects. Deck chrome that *should* sit still — `chromeKicker`, `chromeTitle`,
`chromeSub`, `navSection`, `navCount`, `navChip*` — keeps stable names on purpose.

Each `<p159:morph>` is wrapped in `<mc:AlternateContent>` with a `<p:fade/>`
`<mc:Fallback>`, so Keynote, Google Slides and pre-2016 PowerPoint open the deck
and fall back to a fade instead of failing.

## Known limitation

**Morph cannot be verified in this environment.** LibreOffice renders static
frames only — it ignores `p159:morph` entirely — so `render.py` produces the
before/after keyframes of each morph, never the animation. Shape-name pairing,
transition XML placement and the fallback wrapper are verified programmatically;
the motion itself needs PowerPoint 2016+ (or Microsoft 365) on a real machine.

## On-screen navigation

Every slide carries `‹ BACK` / `MAP` / `NEXT ›` chips plus a slide counter and a
section-and-timing label. `MAP` jumps to slide 10; the five place cards there, and
the rail on slides 11–15, link to their deep-dive slides; practice slides link
back to the method on slide 16.
