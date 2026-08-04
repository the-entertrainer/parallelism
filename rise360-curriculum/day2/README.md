# Day 2 — Blocks that carry your content

Facilitator deck for day 2 of the five-day Articulate Rise 360 beginner curriculum.
29 slides, 1920 × 1080, speaker notes on every slide.

`build/rise360-day2-blocks.pptx`

## Scope

Day 2 is the **static (non-advanced) blocks** only — the nine content families and the
layout settings that go with them:

| Family | Members |
|---|---|
| Text | Paragraph · Paragraph with heading · Paragraph with subheading · Heading · Subheading · Two column · Table |
| Statement | Statement A–D · Note |
| Quote | Quote A–D · Quote on image · Quote carousel |
| List | Numbered · Checkbox · Bulleted |
| Image | Image centered · Image full width · Image & text · Text on image · Banner |
| Gallery | Carousel · Two column grid · Three column grid · Four column grid |
| Multimedia | Audio · Video · Embed · Attachment · Code snippet |
| Chart | Bar · Line · Pie |
| Divider | Continue · Divider · Numbered Divider · Spacer |

42 blocks, one screenshot each, all from the supplied asset pack.

Plus: the three ways to add a block, swapping a block's type without losing content,
padding, backgrounds and contrast, block entrance animations, the anti-clutter rules,
the five beginner mistakes, the lab brief and a done-checklist.

**Deliberately excluded** — interactive blocks (accordion, tabs, labeled graphic, process,
scenario, sorting, timeline, flashcards, buttons, Storyline) and the assessment blocks.
Those are Days 3 and 4, and the lab brief tells learners not to use them today.

## Structure

1. Title · where Day 2 sits · yesterday → today
2. **The stack** — a lesson is a vertical stack; three ways to add a block; the nine families
3. **The nine families** — one slide each: official purpose, every member, a per-member note, one thing to watch
4. **Layout & polish** — padding, backgrounds and contrast, entrance animations
5. **Keeping it readable** — four rules of thumb, five beginner mistakes
6. **Build it** — lab brief, done-checklist, tomorrow, close

## Design

The motif is the block card — a white rounded card with a soft shadow, the same object
a learner stacks in a Rise lesson, so the deck is assembled the way the tool is.

The accent (`ED6A2A`) is sampled from the Rise block screenshots themselves, so the
42 embedded thumbnails and the deck around them share one palette. Dark title, section
and closing slides; light content slides. Cambria headings, Calibri body — both render
true-to-width in the QA previewer.

## Sourcing note

The ≤50% text, 1–2 images, and 20-block figures are practitioner guidance from the
Articulate community, not Articulate's own documentation. The deck says so on the slide
rather than presenting them as vendor rules.

## Build

```bash
npm install pptxgenjs
node tooling/build.js                     # -> build/rise360-day2-blocks.pptx
python3 /root/.claude/skills/pptx/scripts/office/validate.py build/rise360-day2-blocks.pptx
python3 /root/.claude/skills/pptx/scripts/office/soffice.py --headless --convert-to pdf build/rise360-day2-blocks.pptx
```

`tooling/build.js` expects the block screenshots at `img/` relative to itself; they are
kept here in `assets/block-screenshots/` alongside `assets/rise_blocks_metadata.json`,
which supplies the official category names and purpose lines quoted on each family slide.
