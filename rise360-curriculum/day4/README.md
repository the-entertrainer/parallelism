# Day 4 — Decide, then ship

Facilitator deck for day 4 of the five-day Articulate Rise 360 beginner curriculum.
29 slides, 1920 × 1080, speaker notes on every slide.

`build/rise360-day4-decide-and-ship.pptx`

Two hand-outs sit beside it in `practice/`, and they are the working half of the day:

| File | What it is |
|---|---|
| `practice/demo-script.md` | The facilitator's run sheet for the live demo — every word of copy, every click path, timings, scripted mistakes and fallbacks |
| `practice/capstone-brief.md` | The participant hand-out — what to build, by when, and what "ready" means |

## Scope

| Part | Time | Covers |
|---|---|---|
| Buttons | 45 min | Button and Button stack: the five destinations, every setting, the one real limit, and routes without restriction |
| Scenario | 60 min | Anatomy, scenes, characters, backgrounds, Dialogue vs Text, the Go to menu, and a writing method |
| Watch me build one | 45 min | A blank course to a live link — two lessons, a quiz, published and exported |
| Ship it | 30 min | Multi-device preview, then Quick Share, Reach 360, LMS export and Web/PDF |
| Your turn | 15 min | The capstone hand-over |

The four question types and quiz settings are covered as part of the demo, which is where
Day 3 said they would land.

**Deliberately excluded** — theme, navigation and accessibility. Those are Day 5, along with
a clinic for the capstone.

## The demo course

*Feedback that lands* — two lessons and a quiz, nine blocks, four questions, one of each
question type. It uses the two blocks taught that morning, so the demo doubles as
consolidation. Everything the presenter needs to type is written out in `demo-script.md`,
including the full scenario with its branches and the appendix of paste-ready copy.

Three mistakes are scripted in on purpose — a misplaced block, an implausible scenario
option, and an 80% pass mark on a four-question quiz. Recovering from them in front of the
room teaches more than getting it right.

## The capstone

Same shape, participants' own subject, published by Monday morning: a Quick Share link, a
SCORM 1.2 export, and three lines of writing. About three hours of work. The brief carries a
build order, a rubric written as ready / not ready yet, and a printable checklist.

## Sourcing note

Product detail on this day is specific enough to be wrong if it goes stale, so it was
checked against Articulate's current documentation while the deck was built — the button
destination list and settings, the Scenario block's parts and limits, quiz setting defaults,
and every publish and export path including tracking options and launch files.

Rise ships changes every few weeks. Slide 6 tells the room in as many words that if the
interface disagrees with the deck, the interface wins — and both hand-outs repeat it.

## Design

Same system as Days 2 and 3: the block-card motif, the ember accent (`ED6A2A`) sampled from
the Rise thumbnails, dark title/section/close slides, light content slides, Cambria headings
with Calibri body.

## Build

```bash
npm install pptxgenjs
node tooling/build.js                     # -> build/rise360-day4-decide-and-ship.pptx
python3 /root/.claude/skills/pptx/scripts/office/validate.py build/rise360-day4-decide-and-ship.pptx
python3 /root/.claude/skills/pptx/scripts/office/soffice.py --headless --convert-to pdf build/rise360-day4-decide-and-ship.pptx
```

`tooling/build.js` reads the block screenshots from `../day2/assets/block-screenshots` and
the block metadata from `../day3/assets/rise_blocks_metadata.json`.
