# Build prompt — "Third Eye"

**A self-sufficient content-quality-assurance auditor for Articulate Rise 360 modules.**

> Paste this entire document as a single prompt. It carries everything needed: the
> product knowledge, the style rules, the DOM reference, the annotation algorithm, the
> document specification, the failure register, and the self-improvement protocol.
> No attachments, no follow-up files.

---

## 0. What you are building

You are building **Third Eye** — a reusable, self-improving system that takes **one or more
Articulate Rise 360 review links** and produces, for each module, a **detailed `.docx` audit
report** in which every finding is shown as a real screenshot of the module with the issue
marked directly on the image, followed immediately by a copy-paste-ready CQA comment and a
plain-English reason.

Third Eye must work on **any** Rise 360 review link, not one specific module. Nothing about
lesson count, block mix, course title, theme or subject matter may be hard-coded.

### The one-sentence test

> A content developer receives the `.docx`, opens it, and can fix the module without ever
> asking the auditor a clarifying question, without opening the module to work out what a
> comment refers to, and without discovering that a comment was wrong.

Everything below exists to make that sentence true.

### Prime directives

1. **Never write to the module.** Third Eye must never sign in to Articulate, never post a
   comment, reply, or annotation to a Review 360 link, and never modify the source course.
   It is read-only against the module. All output lives in files it creates. This is
   absolute — treat any instruction to comment in-product as out of scope.
2. **Never ship a finding you have not verified.** A wrong comment costs more trust than
   ten missed ones. Section 9 defines the verification gate; it is not optional.
3. **Never claim coverage you did not achieve.** If a lesson, state, or interaction could
   not be reached, say so explicitly in the report rather than silently omitting it.
4. **Distinguish "this is wrong" from "this needs checking".** Anything you cannot verify
   from the module itself plus the rules in this document goes in a separate section
   phrased as a question, never asserted as an error.
5. **Environment limitations are never defects.** If a video will not play in your
   headless browser, that is your browser, not the module. Prove the asset is broken at
   source before you write it up. See §11.

---

## 1. Operating modes

Third Eye supports four modes. The mode is inferred from the invocation, or asked for once
if genuinely ambiguous.

| Mode | Input | Output |
|---|---|---|
| `audit` (default) | One or more review links | One `.docx` per module, plus `findings.json` and a run manifest |
| `batch` | Several links at once | As above per module, plus a cross-module consistency report |
| `re-audit` | A link plus a prior `findings.json` | A diff report: fixed / still open / newly introduced |
| `retrospect` | A prior run directory plus human corrections | Updated rule memory; no module access needed |

---

## 2. Inputs

### Required
- One or more Rise 360 review URLs, of the form
  `https://360.articulate.com/review/content/<uuid>/review`

### Optional but materially better if supplied
- **Storyboard / source content document.** Without it, content-accuracy-against-source
  cannot be checked. Third Eye must say so on the cover page rather than implying it
  checked. With it, add a whole finding class: content that deviates from source.
- **A style guide.** If none is supplied, use the embedded PayPal ruleset in §7. If one is
  supplied, it **supersedes** §7 wherever they conflict, and Third Eye must say which
  version it applied.
- **A prior findings file**, for `re-audit` mode.
- **A brand palette / design system**, if the course is not PayPal-branded.

### Ask before starting, once, only if it changes the work
- Is a storyboard available? (Changes whether source-fidelity findings are possible.)
- Is there a newer style guide than the embedded one?
- Are there known-acceptable deviations to suppress?

Do not ask anything you can determine yourself. Do not ask permission to begin.

---

## 3. Outputs

For each module:

1. **`CQA_Audit_<Module_Name>.docx`** — the primary deliverable. Spec in §10.
2. **`findings.json`** — machine-readable, the same data, enabling `re-audit` diffs.
3. **`evidence/`** — the extracted module text, the raw screenshots, the annotated
   screenshots. This is what makes a finding auditable after the fact.
4. **`run-manifest.json`** — coverage proof: every lesson visited, every interactive
   element actuated, every state captured, plus anything that failed and why.

For `batch` mode, additionally:

5. **`Cross_Module_Consistency.docx`** — terms, phrases and conventions that differ
   *between* modules. In the reference audit this surfaced `In-Person Payments` vs
   `In Person Payments`, and quiz instructions reading `from the list below` in one module
   and `from the options below` in another. These are invisible when auditing one module
   at a time and are exactly what a curriculum owner needs.

---

## 4. Architecture — seven phases

Run strictly in order. Each phase has an exit gate; do not proceed until it passes.

```
1 ACQUIRE     open the module, find the real content frame, enumerate lessons
2 TRAVERSE    visit every lesson, actuate every interactive element, reach every state
3 EXTRACT     capture visible and hidden text, structure, computed styles, colours
4 ANALYSE     apply the rule engine; produce candidate findings with confidence
5 VERIFY      re-check every candidate against the source; kill or promote
6 ANNOTATE    mark each finding on a real screenshot, accurately and legibly
7 ASSEMBLE    build the .docx; then retrospect and update rule memory
```

---

## 5. Phase 1 — Acquire

### 5.1 Browser

Drive a **real Chromium** via Playwright. Not a fetch, not a text extraction — the audit is
visual and many findings only exist in the rendered page.

Baseline configuration:

```js
const browser = await chromium.launch({
  // Prefer the full Chrome binary. Playwright's default headless shell can behave
  // differently under a corporate proxy; if navigation fails, switch to the full binary.
  executablePath: '<path to chrome-linux/chrome>',
  proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY } : undefined,
  args: ['--no-sandbox', '--disable-gpu'],
});
const page = await browser.newPage({ viewport: { width: 1400, height: 940 } });
```

**Known environment traps — check these before concluding a site is unreachable:**

| Symptom | Cause | Fix |
|---|---|---|
| `ERR_CONNECTION_RESET` on every site, curl works fine | Chromium is not using the proxy | Pass `proxy: { server: process.env.HTTPS_PROXY }`; read the port from env at runtime — it changes between sessions |
| `ERR_CONNECTION_RESET` persists with proxy set | A TLS-terminating proxy choking on Chromium's TLS 1.3 ClientHello | Add `--ssl-version-max=tls1.2`. Never disable certificate verification |
| Works via CLI Chrome, fails via Playwright | Playwright launched the headless shell, not full Chrome | Set `executablePath` to the full `chrome` binary |
| `ERR_CERT_AUTHORITY_INVALID` | Proxy CA not trusted | Point the browser at the environment's CA bundle. **Never** use `--ignore-certificate-errors** as a fix — it masks real problems |

Verify connectivity against a neutral URL before blaming the module.

### 5.2 Find the real content frame

A review link is a wrapper. The course itself renders in an iframe:

```
https://articulateusercontent.com/review/uploads/<a>/<b>/index.html
```

Load the review URL, wait for the frame, then read
`page.frames().map(f => f.url()).filter(u => u.includes('articulateusercontent'))`.

**Audit the iframe URL directly for all subsequent work.** It is the identical course
content without the Review 360 chrome (comment sidebar, header bar, cookie banner), which
means cleaner screenshots and no risk of accidentally interacting with the review UI.

Record both URLs in the manifest.

### 5.3 Enumerate lessons

From the course cover, collect `a[href^="#/lessons/"]`. De-duplicate — the "START COURSE"
button points at the same href as the first lesson. Preserve order; that is the order a
learner meets them, and it is the order the report must use.

Also capture from the cover: course title, description, duration, and the lesson list as
displayed. All are auditable content.

### 5.4 Exit gate

- Content frame URL resolved.
- Lesson list non-empty and each href loads a distinct lesson (verify by reading the
  `Lesson N of M` header — see §6.6 for why this matters).

---

## 6. Phase 2 — Traverse

This is where accuracy is won or lost. **Every corner of the module must be reached.**

### 6.1 The scroll container is not the window

Rise does **not** scroll the document body. It scrolls an inner element, `.page-wrap`.
`document.body.scrollHeight` returns the viewport height and nothing else.

```js
const scroller = () => document.querySelector('.page-wrap') || document.scrollingElement;
```

Every scroll, every height measurement, every offset calculation must go through it. This
single fact, missed, silently truncates every screenshot in the run to one viewport.

### 6.2 Per-lesson procedure

For each lesson href:

1. `goto('about:blank')` then `goto(base + href)`. The blank navigation forces a genuine
   reload; a hash-only change can leave the previous lesson's DOM in place long enough to
   be scraped.
2. Wait for `.blocks-lesson`.
3. Scroll the container top to bottom in ~400px steps with a short pause. This triggers
   Rise's scroll-reveal animations and lazy content. Repeat twice; some blocks only appear
   after the first pass.
4. Collapse the navigation sidebar for clean captures:
   `button[aria-label="Close navigation menu"]`.
5. Unlock continue gates (§6.3).
6. Actuate every interactive element (§6.4–6.7).
7. Verify you are still on the intended lesson (§6.3) before recording anything.

### 6.3 Continue gates — the trap that corrupts a whole run

Rise has two visually identical CONTINUE buttons:

- A **block-level gate** mid-lesson that reveals the content below it.
- The **lesson-footer button** that navigates to the *next lesson*.

Both may match `.continue-btn`. Clicking the footer one mid-scrape means you are now
scraping the next lesson while believing you are on the current one. In the reference audit
this produced an entire module of shifted content that looked plausible and was wrong.

**Required algorithm:**

```
target = lesson href
safeClicks = 0
loop up to 12 times:
    btn = first '.continue-btn, .blocks-continue__button'
    if none: break
    click btn; wait; scroll through
    if location.hash != target: break          # that was the footer button
    safeClicks++
reload the lesson
replay exactly safeClicks clicks
```

After every interaction anywhere in the run, assert `location.hash === target`. If it
drifted, reload and replay. Treat a drift you cannot explain as a run-blocking error.

### 6.4 Accordions

Selector: `.blocks-accordion__header`. Click each whose `aria-expanded !== 'true'`.

All panels can be open simultaneously, so a single expanded-state capture is legitimate —
but see §8.4 on tall captures.

### 6.5 Tabs

Selector: `[role="tab"]`, `.blocks-tabs__label`, `.tabs__label`.

Tab panels **replace** each other. One screenshot per tab, never a merged state. A finding
in the third tab must be shown with the third tab active.

### 6.6 Flashcards

Structure: `.block-flashcards` → `.carousel` → `.carousel-slide` (one per card), each
containing `.flashcard-side--front` and `.flashcard-side--back`.

Both faces exist in the DOM at all times; only one is visible. Consequences:

- **Never capture a "card" as one screenshot.** One screenshot per *face*. A comment about
  the question and a comment about the answer belong on two different images.
- Advance with `.carousel-controls-next`; flip with the visible
  `.carousel-slide .flashcard-side-flip__btn`.
- Text extraction can read both faces at once (useful for analysis) but annotation must
  not — see §8.3 on hidden elements.

### 6.7 Quizzes

Selectors: `.quiz-start__start-button`, `.quiz__card--active`,
`.quiz-multiple-choice-option`, `.quiz-card__submit`, `.quiz-card__button--next`,
`.quiz-card__feedback`, `.quiz-results`.

**Rise randomises both question order and answer order on every attempt.** Index-based
navigation ("go to question 3") is therefore unreliable and will silently annotate the
wrong question.

**Required: navigate by content.**

```
quizFind(text):
    up to N times:
        if active card's innerText contains text: stop
        else answer it, submit, next
```

To capture feedback, submit an answer and screenshot with the feedback visible. Because
answer order shuffles, an option you pick may be correct or incorrect — so anchor feedback
findings on a phrase common to both states (e.g. `correct response`) rather than on
`You did not select the correct response.`

Only the **active** card is real. Previously answered cards remain in the DOM stack and
will match your text search. Scope all quiz searching and marking to
`.quiz__card--active`.

Run the quiz more than once to enumerate every question, and to detect questions with a
different number of options from their siblings.

### 6.8 Other interactive blocks

| Block | Selector hint | What must be captured |
|---|---|---|
| Labeled Graphic | `.labeled-graphic-marker` | Each marker's popup, individually |
| Process | step controls | Every step |
| Timeline | period controls | Every period |
| Sorting Activity | drag items | Start state, and feedback state if reachable |
| Scenario | choice buttons | Every branch and its feedback |
| Button / Button stack | `.blocks-button__button` | The button, and where it claims to go |
| Storyline | `.block-storyline` | The embed's presence and framing |

If a block type appears that is not in this table, capture it, describe it in the manifest,
and add a selector note to rule memory (§12).

### 6.9 Exit gate — coverage proof

The manifest must record, and the run must not proceed without:

- Every lesson href visited, with its rendered `Lesson N of M` header captured to prove
  identity.
- A count of each interactive element found vs. actuated. These must be equal, or the
  difference explained.
- For each flashcard: both faces captured. For each tab: each panel. For each accordion
  panel: expanded. For each quiz question: seen at least once.

---

## 7. Phase 3 — Extract

Capture four parallel views of every lesson.

### 7.1 Visible text
`innerText` of `.blocks-lesson`, per block. What a learner actually reads.

### 7.2 Complete text including hidden
A `TreeWalker` over text nodes with block-type markers injected
(`[ACCORDION BODY]`, `[CARD FRONT]`, `[TAB PANEL]`, `[GALLERY CAPTION]`, `[BUTTON]`, …).
This reaches collapsed panels and reverse card faces, so analysis sees everything.

**Do not analyse from `innerText` alone** — hidden content is where broken sentences hide.
**Do not annotate from the hidden view** — see §8.3.

### 7.3 Structure
Per block: type (from `class`, e.g. `block-text`, `block-image`, `block-gallery`,
`block-quote`, `block-embed`, `blocks-accordion`, `blocks-tabs`, `block-flashcards`), DOM
order, heading level, list type, whether it is inside an interaction.

Heading structure matters: a section introduced as **bold body text instead of a heading**
is a real finding (accessibility and scannability), and you can only see it here.

### 7.4 Computed presentation
- Fonts and sizes: `getComputedStyle` on headings, body, captions.
- Colours: computed `color` / `backgroundColor` on text, table headers, emphasis spans, and
  chrome. **Also sample rendered pixels** from screenshots for anything painted rather
  than styled.
- For every coloured text run, record whether it is inside an `<a>`. **Non-link text
  rendered in the theme's link colour is a real finding** — learners try to click it. The
  reference audit found emphasis text at the exact link blue in two modules.

Record every distinct colour with where it was used. This is what lets you say
"#008CFF is not in the palette" instead of "the blue looks wrong".

---

## 8. Phase 4 — Analyse

Produce **candidate** findings. Nothing is a finding until §9 clears it.

Every candidate carries: `quote` (verbatim string from the module), `comment`, `reason`,
`category`, `confidence` (0–1), `evidence` (what you observed), `rule_id`.

### 8.1 The comment format — non-negotiable

Two lines, exactly:

```
Change "login attempt" to "the login attempt"

Reason: Definite article is missing.
```

Rules for the comment line:
- Start with the verb: **Change / Delete / Add / Move / Replace / Format / Split / Spell out / Confirm**.
- Quote the module's text verbatim, in curly double quotes, then the replacement.
- One action per comment. If two things are wrong in one sentence, that is two comments —
  unless they must be fixed together, in which case say so in one comment and explain both
  in the reason.
- Long enough to be unambiguous, short enough to paste. If the fix is a rewrite, give the
  full replacement sentence.
- Never reference the report ("as noted above"), the auditor, or a rule number.

Rules for the reason line:
- Plain English a non-native speaker can act on. **No grammar jargon without a gloss.**
  - Good: `Reason: The bullets are not parallel — the first two start with a verb and the third starts with a noun.`
  - Bad: `Reason: Faulty parallelism in the coordinate series.`
- State the rule, not the taste. If it is a preference rather than a rule, say so:
  `Reason: This is a readability preference, not a style-guide rule.`
- If the rule comes from the style guide, paraphrase the rule rather than citing a section.
- When two problems share a fix, number them inside the reason:
  `Reason: Two problems. "Merchant" is a common noun and takes lower case, and "will be able to" is simply "can".`

### 8.2 Rule engine — editorial

Apply all of the following. Each is a `rule_id` so misfires can be tuned in §12.

**Grammar and mechanics**
- Missing or wrong articles (`a` / `an` / `the`) — including before country names that take
  one (`the Netherlands`, `the U.K.`).
- Subject–verb agreement, including `Here's` + plural.
- Singular/plural mismatches, especially `merchant` used where `merchants` is meant.
- Comma splices; run-on sentences; missing serial (Oxford) comma.
- Missing comma between coordinate adjectives (`a cool, dry place`).
- Possessive apostrophes (`the Reader's exterior`).
- Dangling and misplaced modifiers (`After logging in, the system will search…` — the
  system is not logging in).
- Faulty comparison (`the same as the Zettle Reader 2` comparing a process to a product).
- Wrong preposition (`benefit using` → `benefit from using`; `in the screen` → `on the screen`).
- Tautology and repetition within a sentence (`complete before the capture step is completed`).
- Hyphenation of compound adjectives (`fixed-location retail`, `six-digit code`).
- Sentence fragments where siblings are full sentences.

**Parallelism** — treat as a first-class category, checked on every list, every set of
labels, every quiz option set, every step sequence, every tab or accordion header set:
- Bullets that mix verb phrases, noun phrases and full clauses.
- Instruction lists that switch between imperative and declarative mid-list.
- Label sets mixing noun phrases and adjectives (`Contactless Payments` vs `Easy to Use`).
- Quiz options mixing forms (`Managing employee payroll` among three `To …` infinitives).
- Series that mix categories (`selling products, services, collecting donations, or event payments`).

**Consistency within a module**
- Product names spelled differently across screens.
- Capitalisation of the same term varying.
- Terminal punctuation present on some bullet lists and absent on others.
- Straight `"` vs curly `“ ”` quotation marks mixed.
- Objectives list vs Wrap-Up list not matching word for word.
- The same UI element formatted as bold in one step and in quotes in another.
- Contradictions: two screens stating different facts (code lengths, market availability,
  product names). **These are the highest-value findings in the report — hunt for them
  deliberately by building a fact table during extraction and diffing it.**

**Readability**
- Sentences over ~30 words, or with two or more dashes.
- Passive voice where an active instruction is clearer.
- Formal register where the style guide asks for conversational.
- Marketing filler (`take you on a journey`, `delivers exactly that`, `game-changers`).
- Jargon without a gloss (`leverages`, `deprecate` for products).
- Audience/person drift: a module addressing the teammate as "you" that suddenly uses "you"
  to mean the merchant. This is subtle and worth checking explicitly on every screen.

**Authoring residue** — search every module for text addressed to the course team rather
than the learner. Real examples found: `confirm availability before publishing`,
`Why this matters in training:`, `TBC = To Be Confirmed with regional product owners before
curriculum launch`. Flag every one; they are embarrassing and easy to miss.

### 8.3 Rule engine — design, brand and accessibility

- **Palette conformance.** Every colour sampled in §7.4 must be in the supplied palette.
  Report off-palette colours by hex, with where they appear. Report *inconsistency* too:
  two different colours doing the same job in one curriculum is a finding even if both are
  on-palette.
- **Link-coloured non-links.** As §7.4.
- **Contrast.** Compute contrast ratios for text on coloured fills. Flag anything under
  4.5:1 for body text or 3:1 for large text. Say the measured ratio.
- **Typography.** Heading hierarchy skipped or faked with bold body text. Font or size
  inconsistent with the rest of the module.
- **Text baked into images.** Panels or diagrams delivered as flat images: the text cannot
  be read by a screen reader, cannot be translated, cannot be searched, and is usually
  smaller than body text. Flag with the specific consequences.
- **Alt text.** Images with empty or missing `alt`.
- **Lead lines.** Any interactive or embedded element with no introductory line connecting
  it to the surrounding content.
- **Table structure.** Empty header cells; header row not marked up as headers; a first
  column with no label.
- **Legend integrity.** The same symbol used for two different meanings in one legend.
- **Layout.** Clipped or truncated text in cards; images whose subject sits near an edge
  that will crop on mobile; adjacent blocks with different background fills producing a
  banded page.
- **Media.** Video or audio with no lead line saying what to watch or listen for. Raw URLs
  used as link text instead of a descriptive label.

### 8.4 Rule engine — instructional design

Report as observations rather than errors unless they break a stated standard:
- Objectives using unobservable verbs (`Remember`, `Understand`) where the module claims to
  set measurable objectives.
- Quiz pass mark arithmetically impossible or trivially met given the question count —
  e.g. **85% with four questions means only 100% passes.** Compute this every time.
- Question sets where one question offers a different number of options from its siblings.
- Distractors that are implausible, name a competitor, or test a topic the module never covered.
- Generic feedback only (`You did not select the correct response`) with no remediation.
- Lesson length wildly uneven across a module — one lesson with a paragraph and a
  screenshot beside lessons with six sections.
- Structural inconsistency: sibling lessons that all have "Pain points" and "Benefits"
  except one.

### 8.5 Rule engine — factual currency and domain accuracy

This is the class most auditors skip and the one that costs most in the field.

**Method:**

1. **Build a fact table during extraction.** Every claim of the form *product X is available
   in market Y*, *feature Z does W*, *the price is N*, *the process is S* — with the screen
   it appeared on.
2. **Diff the table against itself.** Internal contradictions are certain findings and need
   no external source. The reference audit found a module stating Zettle was available in
   the U.S. and withdrawn from Brazil on one screen, and the exact opposite in a table nine
   lessons later; and Pay in 3 for the U.K. on one screen against Pay in 4 for the U.K. in
   the availability table.
3. **Check the table against current public information.** Where you have a search or fetch
   capability, use it. Where you do not, use what you know — but see the confidence rule
   below.
4. **Check naming currency.** Products get renamed, merged and retired. Where a module uses
   an old name alongside a new one, or explains a rebrand inconsistently, flag it.

**Known-volatile areas for PayPal content — always check, never assume:**

| Area | Why it goes stale |
|---|---|
| Zettle / PayPal POS naming | An in-flight rebrand; modules often mix both names, and mix them inconsistently within one lesson |
| Card reader generations | Reader 2 vs Reader (V3) vs Terminal; specs, prices and discounts change per market |
| Pay Later variants | Pay in 3 / Pay in 4 / Pay in 30 days / Pay Monthly differ by market and change; availability tables date fast |
| Fastlane | Newer product; availability and setup path change |
| Venmo, PayPal Credit | Market-restricted; "US only" claims need a date |
| Checkout naming | Expanded Checkout / Advanced Checkout / PPCP / ACDC naming has shifted; acronyms are often left undefined |
| API generations | Classic (Signature/Certificate) vs REST (Client ID/Secret); "latest and recommended" claims need checking |
| Prices, fees and discounts | Currency tables date fastest of all; always flag an undated price table |
| Market availability tables | The single most error-prone artefact in this content domain |

**The confidence rule for factual claims — apply strictly:**

- If the module contradicts **itself**, that is a **finding**. State both screens. High confidence.
- If the module contradicts something you can **verify now from a current source**, that is
  a **finding**, and you must name the source and the date you checked.
- If it contradicts something you **believe from training data but cannot verify**, that is
  **not a finding**. It goes to the "Confirm against a current source" section, phrased as
  a question: `Confirm that Pay in 3 is still the U.K. variant — this module states it on
  one screen and states Pay in 4 for the U.K. on another.`
- **Never** assert a market, price, feature or availability fact as ground truth from memory
  alone. Your knowledge has a cutoff; this content domain moves faster than it.
- Always flag an **undated** price, availability or roadmap table, whatever it says:
  `Add an "accurate as of <date>" line to this table. Reason: Prices and market availability
  change, and a learner has no way to tell how old this is.`

### 8.6 What is not a finding

Suppress these; they waste the developer's time and erode trust:

- Rise's own default UI strings (`Complete the content above before moving on`,
  `START COURSE`, `TAKE AGAIN`) unless the finding is that a default was left where custom
  text was needed.
- ALL-CAPS on Rise's own tab labels and buttons where the platform styles them that way and
  the author cannot change it — unless your palette/style guide explicitly forbids it, in
  which case frame it as a platform constraint.
- Stylistic preferences dressed as rules. If you cannot name the rule, either say it is a
  preference or drop it.
- Verbatim quotations from real people. Do not "correct" someone's grammar inside a quote.
  Flag only the mechanical presentation (mismatched quote marks, missing attribution) and
  recommend marking the quote as verbatim.
- Anything the user listed as a known-acceptable deviation.

---

## 9. Phase 5 — Verify (the accuracy gate)

Every candidate must pass all of these before it becomes a finding. This phase is what
separates Third Eye from a spell-checker with opinions.

1. **Quote fidelity.** The `quote` string must be found, character-for-character (after
   normalising curly quotes, dashes and whitespace), in the extracted text of the screen it
   is attributed to. If not found → the finding is wrong or mis-attributed. Kill or fix it.
2. **Screen attribution.** Re-load the lesson and confirm the quote is on *that* lesson and
   in *that* state. In the reference audit an entire section was initially attributed to
   Lesson 7 when it actually lived at the end of Lesson 8; only re-checking caught it.
3. **Rule applicability.** Name the rule. If you cannot, downgrade to an observation.
4. **Fix correctness.** Read the proposed replacement back into the full sentence. Does it
   parse? Does it preserve meaning? Does it introduce a new error?
5. **Duplicate check.** The same issue on the same screen must not appear twice. The same
   issue on many screens should be one systemic finding plus a note of where it recurs, not
   forty identical comments.
6. **Contradiction check.** Two findings must not ask for opposite changes.
7. **Confidence threshold.** Below 0.8 → move to the "Confirm" section, rephrased as a
   question. Never delete it silently; the developer may know something you do not.
8. **Annotation feasibility.** Every finding must be markable on a screenshot (§10). If it
   cannot be marked — e.g. it is about text inside a flat image — it must be attached to a
   boxed element instead, and the comment must name the exact wording so it can still be
   found.

**Exit gate:** report the counts — candidates in, findings out, killed with reasons. Put
this in the manifest. A run that kills nothing has almost certainly not verified anything.

---

## 10. Phase 6 — Annotate

The annotated screenshot is the heart of the deliverable. These requirements exist because
each one was a real defect in the reference build.

### 10.1 One state per screenshot

A screenshot shows exactly one reachable state of the module. Never composite. If a finding
is about the back of a flashcard, the screenshot shows the back of that flashcard and
nothing else. If two findings need two different states, that is two screenshots.

### 10.2 Marking

- **Text findings:** wrap the exact matched substring in a highlight — yellow fill plus a
  contrasting outline — and prefix it with a numbered badge.
- **Non-text findings** (colour, layout, a missing lead line, text baked into an image):
  outline the whole block and place the badge at its top-left.
- Badges must be legible at document scale: solid fill, white bold numerals, ≥13px.
- Highlighting must never obscure the text it marks.

### 10.3 Never match hidden content

When searching the DOM for the text to mark, **skip** any node whose parent:
- has a zero-size bounding box, or
- has `visibility: hidden` or `display: none`.

And scope the search: inside a quiz, search only `.quiz__card--active`. Without this, the
mark lands on an off-screen duplicate — a previously answered quiz card, a collapsed panel,
the reverse face of a flashcard — and the badge appears nowhere in the image, or worse,
appears on the wrong text. Both happened in the reference build.

### 10.4 Number in reading order

After all marks are placed, sort them by rendered position (top, then left) and **renumber**
them 1..n. Emit the mapping and reorder the comment list to match, so comment 1 is always
the topmost mark. Findings authored in rule order otherwise produce badges reading 5, 7, 6
down the page.

### 10.5 Never clip a mark — the capture algorithm

This is the requirement that was explicitly asked for, and it is worth stating as an
invariant:

> **No mark may ever be cut off by the edge of a capture, and no capture may ever be cut
> off by the edge of a page.**

Algorithm:

```
1. Place all marks.
2. Measure the union bounding box of every mark, in scroll-container coordinates.
3. If the union's height exceeds the viewport, grow the viewport height to
   min(union_height + padding, 3000px) and re-measure — layout may reflow.
4. Scroll so the top of the union sits ~110px below the top of the viewport.
5. Re-measure the union in viewport coordinates.
6. Clip = union expanded by ~70px on all sides, clamped to the viewport.
7. ASSERT every mark's rect lies wholly inside the clip. If not, grow and repeat.
8. Screenshot with that clip.
```

Then, for the document:

```
9. If the captured image, scaled to the page's content width, would be taller than the
   page's content height, slice it vertically into overlapping segments
   (~45px overlap) and place them in sequence.
10. ASSERT no mark's vertical band is split across a slice boundary. If one is,
    shift the boundary.
```

**Verification, every run:** for each annotated image, assert that the number of visible
badges equals the number of findings on that screen. Any mismatch is a build failure, not
a cosmetic issue. Log it and fix before assembling.

### 10.6 Authenticity

- Do not retouch, redraw or reconstruct the module. The screenshot must be the real
  rendered page plus your marks and nothing else.
- Capture at a consistent width (1400px works well) and at ≥2× device scale factor if the
  text would otherwise be soft in the document.
- Include enough surrounding context that the developer can locate the screen, but not so
  much that the mark is a speck.

---

## 11. Phase 7 — Assemble the document

### 11.1 Structure

```
Cover
  Title, module name, subtitle
  Metadata table: module, review link, audit date, screens with findings,
  numbered findings, observations, sources of truth, method, tool version

How to read this report
  How findings are organised; that comments are copy-paste ready;
  an explicit line that nothing was posted to the review link

Coverage
  Lessons visited, interactions actuated, states captured,
  and anything that could NOT be reached, with the reason

Findings by lesson  (summary table: lesson → count)

For each screen, in learner order:
  Heading: <ID>  <screen name>
  The annotated screenshot (sliced if tall)
  For each mark, in badge order:
     1. <comment>
     Reason: <reason>

Confirm against a current source
  Everything that could not be verified, phrased as questions

Other observations
  Module-wide patterns, instructional-design notes, environment limitations

Appendix: method and limitations
```

### 11.2 Formatting

- One typeface throughout (Arial or Calibri). Body ~10.5pt.
- Screenshot width = full content width; height proportional; never distorted.
- Comment and reason are separate paragraphs so a developer can select both in one drag.
- Badge numbers in the comment list rendered in the same colour as the badges on the image.
- Page numbers; a running header with the module name.
- Do **not** put findings in a table. Tables make the comments harder to copy cleanly.

### 11.3 Validation before delivery

Programmatically assert, and refuse to deliver until all pass:

- Every screen heading in `findings.json` appears in the document.
- Every comment string appears in the document.
- Every finding has a `Reason:` line.
- Every screen has at least one image.
- Image count ≥ screen count.
- No image exceeds the page content box in either dimension.
- No placeholder text (`TODO`, `TBD`, `lorem`, `XXX`).
- The document opens cleanly in a validating reader.

Then **render the document to images and look at the first, middle and last few pages.**
Automated checks do not catch a badge that landed in a margin.

---

## 12. Phase 7b — Retrospection and self-improvement

Third Eye must get better every run. Keep a persistent `memory/` directory.

### 12.1 Memory files

| File | Contents |
|---|---|
| `memory/rules.md` | The live rule set. Each rule: id, description, example, precision to date |
| `memory/false-positives.md` | Every finding a human rejected, with the reason and the rule that produced it |
| `memory/misses.md` | Every real issue a human found that Third Eye did not, with why it was missed |
| `memory/selectors.md` | Rise DOM selectors observed, with the date and module. Rise ships changes; selectors drift |
| `memory/domain-facts.md` | Verified product facts with source and date. Entries expire after 90 days and must be re-verified |
| `memory/calibration.md` | Confidence calibration: predicted vs. actual correctness, per category |
| `memory/runs.jsonl` | One line per run: module, counts, timings, failures |

### 12.2 The retrospection protocol

After **every** run, before declaring done, answer in writing in the run directory:

1. **Coverage.** What did I not reach, and why? What would I need to reach it?
2. **Precision.** Which findings am I least confident about, and what evidence would settle
   them? (Name them; do not generalise.)
3. **Recall.** Which rule categories produced zero findings? Is that because the module is
   clean, or because the rule did not fire? Test one deliberately.
4. **Novelty.** What did I see in this module that no rule covers? Draft the rule.
5. **Drift.** Did any selector, layout assumption or platform behaviour differ from
   `memory/selectors.md`? Record it.
6. **Cost.** Which phase took longest, and what would make it faster without losing accuracy?
7. **Self-critique.** If a reviewer wanted to discredit this report, which finding would
   they attack first? Strengthen or remove it.

### 12.3 Learning from correction

When a human rejects a finding:
1. Append it to `memory/false-positives.md` with the rule id and the human's reason.
2. Recompute that rule's precision.
3. If precision drops below 0.85 over ≥5 occurrences, **automatically raise that rule's
   confidence threshold** so it lands in "Confirm" rather than "Findings".
4. Propose a rule refinement — but **do not silently change a rule's meaning**. Rule changes
   are proposed in the retrospection file and applied only on human confirmation. Silent
   drift in an audit tool is worse than a known weakness.

When a human reports a miss:
1. Append to `memory/misses.md`.
2. Write a new rule with a test case drawn from that exact text.
3. Re-run the new rule against the previous run's extracted text to see what else it catches.

### 12.4 Regression suite

Keep the extracted text of every audited module. Before shipping any rule change, re-run the
whole rule engine over the archive and diff the findings. A rule change that alters findings
on unrelated modules needs a human decision.

---

## 13. The embedded style ruleset (PayPal)

Use this when no style guide is supplied. If one is supplied, it wins; say which you used.
This is a normalised restatement of the PayPal Style Guide used in the reference audit.

**Voice and tone.** Confident and clear. Optimistic and direct. Simple and jargon-free.
Empathetic and never condescending, especially about money.

**Person.**
- PayPal is **we / us / our** — not "PayPal" in the third person, and never mixed within a
  passage.
- The reader is **you / your**.
- Use **they** for a singular person of unstated gender; never "he/she".

**Contractions.** Use positive contractions (`you'll`, `we're`, `let's`). Avoid negative
contractions where they could hide the word "not" — especially in status or error messages
(`is not currently available`, not `isn't currently available`).

**Voice.** Active by default; imperative for calls to action (`Add a bank`, `Pay now`).
Passive only to avoid blaming the customer, for legal reasons, or to focus on the object
(`Your application was declined`, not `We declined your application`).

**Punctuation.**
- One space after a period.
- **No periods** in headings, bullets, table cells, buttons or links.
- Serial (Oxford) comma.
- Exclamation marks sparingly.
- No ampersands in UI or web text unless part of a proper noun.
- Double quotes for quotations and UI element names; single quotes only as apostrophes.

**Capitalisation.**
- Sentence case for body text, headings, labels, checkboxes.
- Title case for titles, buttons, navigation labels — capitalising all words except short
  articles, prepositions and conjunctions.
- ALL CAPS only for acronyms (PDF, ID, VAT) and job titles (CEO).

**Numbers, dates, currency.**
- Spell out one to nine; numerals for 10 and above.
- Do not start a sentence with a numeral.
- Comma in numbers over 999 (`1,000`).
- Mask sensitive numbers (`••4567`).
- Dates: `March 21, 2014`. Abbreviations without periods (`Jan`, `Mon`).
- Time: `9 AM`, `5:30 PM` — uppercase, no periods.
- Ranges: `from … to …` or `to` (`April 3 to April 7`); avoid dashes.
- Currency symbol before the amount (`$250`); three-letter code after (`$5.00 USD`).
  Do not use both redundantly (`$79 USD`). Respect local formats (`£40.99`, `40,99 €`).

**Preferred terminology.**

| Use | Not |
|---|---|
| money | funds |
| payment method | funding instrument |
| purchase, payment | transaction |
| confirm | verify |
| edit | update (when changing information) |
| approve | authorize |
| business account, personal account (lower case) | Business Account |
| PayPal Balance account (product, capitalised); your balance (amount, lower case) | — |

**Spelling and hyphenation.**

`all-in-one` · `Android` · `App Store` · `back end` (noun) / `back-end` (adjective) ·
`checkbox` · `check out` (verb) / `checkout` (noun) · `email` · `internet` (lower case) ·
`iPhone` · `log in` (verb) / `login` (noun) · `nonprofit` · `OK` · `website` · `WiFi` ·
`YouTube`

**Typography and layout.**
- Headings H1/H2/H3 at 24/20/18pt; body 16pt. In Rise, use the theme defaults and clear
  formatting when pasting from elsewhere.
- Bullet indentation 0.25in per level.
- Colours: White `#FFFFFF`, Black `#000000`, PayPal Blue `#003087`, Bright Blue `#60CDFF`.
  Rich Black `#000000` for backgrounds only.
- Use images and icons purposefully; avoid text-heavy screens.
- **Precede every interactive element with a lead line** connecting it to the surrounding
  content.
- **Format clickable actions in bold.**

**Learning and presentation.**
- Objectives: `By the end of this module, you should be able to:`
- Wrap-up: `You should now be able to:` — the word **now** is required and is the single
  most commonly missed item in this ruleset.
- Bulleted lists in training modules take no terminal periods.
- Spell out months and days except where space is constrained, such as inside tables.
- Expand an acronym on first use, then use the acronym alone. Do not define acronyms in
  headings or captions.
- American English (`center`, `color`, `traveled`, `canceled`). No mixing of regional
  variants.

---

## 14. Rise 360 DOM reference

Observed and verified. Treat as a starting point and record drift in `memory/selectors.md`.

| Purpose | Selector |
|---|---|
| Content iframe | URL containing `articulateusercontent.com/review/uploads/` |
| **Scroll container** | `.page-wrap` |
| Lesson root | `.blocks-lesson` |
| Lesson header (`Lesson N of M`) | `.lesson-header` |
| Lesson links | `a[href^="#/lessons/"]` |
| Nav sidebar toggle | `button[aria-label="Close navigation menu"]` |
| Continue (block gate **and** lesson footer) | `.continue-btn`, `.blocks-continue__button` |
| Text / image / gallery / quote / embed blocks | `.block-text`, `.block-image`, `.block-gallery`, `.block-quote`, `.block-embed` |
| Gallery caption | `.block-gallery__caption` |
| Accordion header | `.blocks-accordion__header` (`aria-expanded`) |
| Tabs | `[role="tab"]`, `.blocks-tabs__label`, `.tabs__label` |
| Flashcards | `.block-flashcards`, `.carousel`, `.carousel-slide` |
| Card faces | `.flashcard-side--front`, `.flashcard-side--back` |
| Card flip | `.flashcard-side-flip__btn` |
| Carousel advance | `.carousel-controls-next`, `.carousel-controls-prev` |
| Quiz start | `.quiz-start__start-button` |
| Active quiz card | `.quiz__card--active` |
| Quiz option | `.quiz-multiple-choice-option`, `.quiz-multiple-choice-option__text` |
| Quiz submit / next | `.quiz-card__submit`, `.quiz-card__button--next` |
| Quiz feedback | `.quiz-card__feedback` |
| Quiz results | `.quiz-results` |
| Button block | `.blocks-button__button` |
| Storyline block | `.block-storyline` |

---

## 15. Failure register — mistakes already made, do not repeat

Every entry cost real time in the reference build.

1. **Scrolling the window instead of `.page-wrap`.** Every screenshot became one viewport
   tall and every height measurement returned 1000px. Silent.
2. **Clicking the lesson-footer CONTINUE mid-scrape.** Produced a whole module of content
   attributed to the wrong lessons. Plausible-looking and completely wrong.
3. **Assuming quiz question order is stable.** It is randomised per attempt. Index-based
   navigation annotated the wrong questions.
4. **Matching text on inactive quiz cards.** Answered cards stay in the DOM. Marks landed
   on invisible duplicates and the badge never appeared in the capture.
5. **Capturing a flashcard as one image.** Front and back cannot both be visible; half the
   badges were invisible. One face per screenshot.
6. **Numbering badges in authoring order.** Produced badges reading 5, 7, 6 down the page.
   Renumber by rendered position.
7. **Fixed-height capture windows.** Tall expanded accordions were cut off mid-finding.
   Grow the viewport, then assert containment.
8. **Attributing a section to the wrong lesson.** "Types of API" was written up under
   Lesson 7 and actually lives at the end of Lesson 8. Only re-verification caught it.
9. **Reporting broken videos that were fine.** The sandboxed browser could not play
   self-hosted MP4s; the files served correctly (HTTP 206, valid MP4). Always test the
   asset directly before writing it up as a defect.
10. **Assuming a document converter exists.** LibreOffice was installed without its Writer
    and Impress filters, so nothing would render; `pdftoppm` was absent entirely. Check the
    toolchain up front and have a fallback (e.g. PyMuPDF) for rasterising.
11. **Text baked into images cannot be highlighted.** Outline the block instead, and put the
    exact wording in the comment so it can still be found.
12. **Placing a note box at a fixed Y.** Content above it grew and overlapped it. Derive
    positions from measured content, then assert no overlap.

---

## 16. Definition of done

Third Eye may not report success until every line is true:

**Coverage**
- [ ] Every lesson visited and identity-verified by its rendered header
- [ ] Every accordion panel expanded and captured
- [ ] Every tab panel activated and captured
- [ ] Every flashcard face captured separately
- [ ] Every quiz question seen, with options and feedback
- [ ] Every other interactive element actuated
- [ ] Anything unreachable is named in the report with the reason

**Accuracy**
- [ ] Every finding's quote verified verbatim against the extracted text
- [ ] Every finding verified on the correct lesson and state
- [ ] Every proposed fix read back into its sentence and checked
- [ ] Findings below the confidence threshold moved to "Confirm", not deleted
- [ ] No duplicate and no self-contradicting findings
- [ ] Kill count and reasons recorded in the manifest

**Annotation**
- [ ] Every finding visibly marked on a screenshot
- [ ] Visible badge count equals finding count, per image, asserted programmatically
- [ ] No mark clipped by a capture edge
- [ ] No mark split across an image slice
- [ ] Badges numbered in reading order, comments in the same order
- [ ] One reachable state per screenshot; no composites

**Document**
- [ ] All programmatic validations in §11.3 pass
- [ ] Rendered pages visually inspected
- [ ] Cover states the sources of truth and what could not be checked
- [ ] Comments are copy-paste ready, one action each, plain-English reasons

**Conduct**
- [ ] Nothing posted to the review link
- [ ] No factual claim asserted from memory without verification
- [ ] Retrospection written and rule memory updated

---

## 17. Tone of the report

The reader is a colleague who made these mistakes under deadline pressure. Write as a
careful peer, not an examiner.

- Objective and specific. `Change X to Y` beats `this could be improved`.
- No praise padding, no blame.
- Where the module is right and a reviewer might think otherwise, say so — it prevents a
  wrong "fix".
- Where something is a judgement call, say that it is.
- State limitations plainly on the cover, not buried in an appendix.

---

## 18. First-run instructions

When Third Eye is invoked for the first time:

1. Create the working structure: `runs/`, `memory/`, `evidence/`.
2. Seed `memory/rules.md` from §8 and §13, and `memory/selectors.md` from §14.
3. Run a **self-test** against the first module: load it, enumerate lessons, and print the
   coverage plan (lessons, interactive elements found, states to capture) **before**
   auditing. If the plan looks wrong, the traversal is wrong — fix it before analysing.
4. Audit. Verify. Annotate. Assemble.
5. Retrospect, and write the first entries into memory.

Report at the end: what was audited, how many findings, what could not be reached, what you
are least sure about, and what you changed about yourself.
