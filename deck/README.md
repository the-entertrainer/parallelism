# Parallelism — facilitator deck (PowerPoint)

A 21-slide, ~60-minute working session on grammatical parallelism for
instructional designers and content developers. Companion to the interactive web
app in the repository root.

`build/parallelism-training.pptx` — 1920 × 1080 (13.333in × 7.5in), speaker notes
on all 21 slides, on-screen navigation, Morph transitions.

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
