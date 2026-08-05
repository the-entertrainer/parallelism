# Day 3 — Blocks that answer back

Facilitator deck for day 3 of the five-day Articulate Rise 360 beginner curriculum.
26 slides, 1920 × 1080, speaker notes on every slide.

`build/rise360-day3-interactive.pptx`

## Scope

Day 3 is the **interactive and advanced blocks** — twelve of them, grouped into six jobs:

| Job | Blocks | Example used in the deck |
|---|---|---|
| Reveal | Accordion · Tabs | Twelve refund FAQs; three pricing plans compared |
| Explore | Labeled Graphic | A card reader with five parts marked on the photo |
| Sequence | Process · Timeline | Pairing a reader; how the product line changed |
| Recall | Flashcard grid · Flashcard stack | Six product names; ten support FAQs as revision |
| Decide | Sorting Activity · Scenario | Phishing or safe; an angry refund call |
| Act | Button · Button stack | Download the job aid; pick your region |
| — | Storyline | The door out of Rise, with its costs stated |

Plus a decision table (which interaction for which job) and four rules that each prevent
a specific failure mode.

**Deliberately excluded** — the four question types and Draw from Question Bank. Those are
assessment, and they are Day 4. The lab brief says so explicitly.

## The bonus: a swipe deck Rise cannot build

The last third of the deck teaches the limits of the tool honestly, then works around them:

1. **What actually runs.** The Code snippet block *displays* code as monospaced text — nothing
   executes. The Embed block loads a page you host, and your HTML, CSS and JavaScript run
   inside it. Storyline runs a published Storyline interaction. Three blocks, three different
   answers to "can I use code in Rise?"
2. **The build.** `bonus/swipe-deck.html` — a Tinder-style swipe deck where each card is a
   scenario and the learner swipes right for "an interactive block earns its place here" or
   left for "a static block does the job". Self-referential on purpose: the bonus doubles as
   revision of the whole day.
3. **The complete source, in the deck.** Six slides carry all 166 lines of the file as code
   blocks. The build script reads the file from disk and chunks it, so the slides and the
   shipped file cannot drift apart — this is verified after every build.

The component is one self-contained HTML file. No dependencies, no build step, no account.
It opens straight from disk in any browser, and the six cards are a plain array at the top
for anyone who wants to swap in their own content.

What it does: pointer events (one path for mouse, touch and pen), rotation derived from
drag distance, a single 96px commit threshold, verdict stamps that fade in as the card
travels, a live-region announcement of the feedback, and — because a drag-only interaction
would exclude keyboard and screen-reader users — two buttons and the arrow keys running
through exactly the same commit path. A `prefers-reduced-motion` query drops the transition.

## Structure

1. Title · from reading to doing
2. **Twelve blocks, six jobs** — all twelve at a glance, then one slide per job
3. **Choosing well** — the decision table, the four rules
4. **Bonus** — what runs in Rise, the swipe deck, how it works, the full source
5. Lab brief · close

## Design

Same system as Day 2: the block-card motif, the ember accent (`ED6A2A`) sampled from the
Rise thumbnails, dark title/section/close slides, light content slides, Cambria headings
with Calibri body. Code blocks are Courier New on a near-black card.

## Build

```bash
npm install pptxgenjs
node tooling/build.js                     # -> build/rise360-day3-interactive.pptx
node tooling/shot.js                      # re-capture the three swipe screenshots
python3 /root/.claude/skills/pptx/scripts/office/validate.py build/rise360-day3-interactive.pptx
```

`tooling/build.js` reads `bonus/swipe-deck.html` at build time for the code slides, and the
block screenshots from `../day2/assets/block-screenshots`.
