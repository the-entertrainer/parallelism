# Third Eye

`THIRD_EYE_BUILD_PROMPT.md` — a single, self-sufficient prompt that builds a reusable CQA
auditor for Articulate Rise 360 modules.

Paste the whole file as one prompt. It carries its own product knowledge, style ruleset,
DOM reference, annotation algorithm, document specification, failure register and
self-improvement protocol, so it needs no attachments — no style-guide PDF, no sample
report, no selector notes.

## What the built system does

Takes any Rise 360 review link and produces a `.docx` in which every finding is a real
screenshot of the module with the issue marked on the image, followed immediately by a
copy-paste-ready CQA comment and a plain-English reason:

```
Change "login attempt" to "the login attempt"

Reason: Definite article is missing.
```

## Where it comes from

Distilled from the audit in `../cqa-audit/` — two PayPal modules, 76 annotated screens and
276 findings. Section 15 of the prompt is a failure register of the mistakes made during
that build (hidden-element mis-annotation, continue-gate corruption, quiz randomisation,
clipped captures, mis-attributed lesson placement), each with the rule that prevents it
recurring.

## Design decisions worth knowing

- **Read-only.** The system never signs in to Articulate, never comments on a review link,
  never touches the source course. All output lives in files it creates.
- **A confidence rule for domain facts.** Anything not verifiable from the module itself
  plus the embedded rules goes into a "Confirm against a current source" section phrased as
  a question — never asserted as an error. Recalled product facts are not evidence.
- **A clipping invariant.** No mark may be cut off by the edge of a capture, and no capture
  may be cut off by the edge of a page. The capture algorithm grows the viewport to fit and
  asserts containment before saving.
- **Environment limitations are never defects.** A video that will not play headless is the
  browser's fault until proven otherwise at source.
- **It improves itself.** After each run it writes to a memory set (rules, false positives,
  misses, selectors, domain facts, calibration) that the next run loads first.
