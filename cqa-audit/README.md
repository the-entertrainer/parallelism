# Rise 360 CQA audit — PayPal modules

Visual content-quality-assurance audit of two Articulate Rise 360 modules, covering
editorial (CQA) and design/brand compliance.

## Modules reviewed

| Module | Review link |
|---|---|
| In-Person Payments with PayPal | https://360.articulate.com/review/content/0aa5a2c1-aa69-4e63-8a7e-a9f233d5fc38/review |
| Products Deep Dive | https://360.articulate.com/review/content/737cabbf-b4ee-4c62-8ce6-0d9a16527cd6/review |

## Deliverables

`reports/` holds one `.docx` per module. Each report has, for every screen with findings:

1. A real screenshot of the module, with every issue outlined and numbered in red.
2. Directly beneath it, a numbered comment and reason for each mark. Numbers run
   top-to-bottom in the image, so number 1 in the list is number 1 on screen.

Comments are written to be copied straight into a review tool. Reasons are in plain English.

**No comments were posted to the Rise 360 review links.** All output lives in the documents.

## Sources of truth

- `PayPal_Style_Guide_Final.pdf` — voice and tone, typography, colour palette,
  punctuation, capitalisation, numbers/dates/currency, preferred terminology,
  the "phrases to avoid" table, and the learning/presentation conventions.
- Standard American English usage.

No storyboard or source content document was supplied, so content accuracy against
source could not be verified. Findings that need someone to confirm a fact against
source are called out explicitly rather than asserted.

## Method

Both modules were driven in a real desktop browser (Chromium via Playwright) at
1400 px wide. Every lesson was read top to bottom, and every accordion panel, tab,
flashcard face and quiz question was opened and answered. Theme colours were sampled
from the rendered page rather than assumed.

Quiz questions and answer options are shuffled on each attempt, so the tooling
navigates the quizzes by question text rather than by position.

## Known limitation

The self-hosted MP4 videos would not play in the audit browser, although the files
themselves are served correctly (HTTP 206, valid MP4). This is a limitation of the
audit environment, not a defect in the modules, and is recorded as such in both
reports. Embedded YouTube videos played normally. Video playback should be confirmed
manually.

## Layout

```
cqa-audit/
├── reports/     the two .docx deliverables
├── evidence/    extracted module text and the structured findings used to build them
└── tooling/     the scripts that drove the browser, marked up screenshots, built the docs
```
