---
name: slide-content-qa
description: Verify that every piece of content was copied accurately and completely from an original (strategist-drafted) slide to its redesigned duplicate in a .pptx deck. Use whenever the user uploads a PowerPoint where originals are hidden and the redesigned duplicate follows, or asks to check, QA, audit, proof, diff, or compare slides before and after design, "make sure nothing got lost in the redesign", "did all the copy carry over", content transfer, or copy fidelity between two versions of a slide or two templates - even if they don't say "QA". Reports missing, changed, or added text, numbers, table and chart data, speaker notes, hyperlinks, images and alt text. Report only; never edits the deck.
---

# Slide content QA

## Why this exists

The design team's workflow: a strategist hands over a rough slide with **finalized copy**. The designer duplicates it, hides the original, and redesigns the duplicate directly after it. Copy is final, so any difference between the two is either a designer's mistake or something to confirm. A dropped footnote, a transposed digit, or lost speaker notes reaches the client unnoticed because nobody rereads copy that "was already approved".

The goal is **zero unexplained differences**. No single check is trusted alone: a deterministic script catches what eyes skim past, a visual pass catches what the script can't see, and a resolution step makes sure every flag is confirmed or explained.

This skill reports. It never edits the user's deck; the designer owns the fixes.

**Content only, never appearance.** The duplicate is supposed to look different; that is the designer's job. Do not report or comment on layout, position, overlap, spacing, fonts, colors, sizing, text-box structure, how bullets are grouped, or whether the slide "looks like" the original. The visual pass exists only to confirm that each piece of original content is present somewhere on the duplicate. Say nothing about design quality, even helpfully.

## Inputs

A `.pptx` where original slides are **hidden** and each is followed immediately by its visible duplicate. If the deck is a Google Slides export, "skipped" slides export as hidden and work the same way. If there are no hidden slides, or the pairing is different, ask which slides pair up and pass `--pairs 3:4,7:8` (original:duplicate, 1-based, hidden slides counted).

Other formats (PDF, Keynote, Figma Slides) lose hidden slides or structure. Ask for a `.pptx` export instead of guessing.

## Procedure

Scripts are in `scripts/` next to this file. Install dependencies if missing: `pip install python-pptx pillow --break-system-packages`. Rendering needs `soffice` and `pdftoppm`.

### 0. Confirm scope and pairing with the user (always, before checking anything)

A wrong pairing produces a wall of false failures or, worse, silently checks the wrong slides, so settle it first:

```bash
python scripts/check_pairs.py deck.pptx --list
```

This prints each hidden slide and the slide right after it, with their first line of text. Show the user that list and confirm three things, using the question tool when available:

1. **Which slides to check**: the whole deck, or a specific range or list (use `--only 5-12`). Don't assume "all".
2. **The originals are the hidden slides.** If the deck has no hidden slides, or hidden slides the user didn't expect (old backups, appendix), the convention is broken here and the pairing can't be trusted.
3. **The designed slide is the one immediately after each hidden slide.** Read the titles aloud to the user: "hidden slide 3 `Revenue grew 42%` -> slide 4 `Revenue grew 24%`". If any pair's titles look unrelated, or the script lists a hidden slide with no visible slide after it, ask before proceeding.

If anything is off, get the correct pairs from the user and pass them as `--pairs 3:4,7:8`. Don't proceed on a guess. Carry the confirmed `--pairs` / `--only` values through every later command, and state the confirmed scope at the top of the final report.

### 1. Prove the checker works (every session, takes seconds)

```bash
python scripts/selftest.py
```

It builds pairs with 12 known seeded errors (dropped footnote, changed digit, lost notes, lost link, swapped image, edited table cell, edited chart value, etc.) and must print `SELFTEST PASSED`. If it fails, stop and tell the user; a checker that fails its own test can't be trusted with their deck.

### 2. Run the automated check

```bash
python scripts/check_pairs.py deck.pptx --json report.json --md report.md --render renders/ [--only 5-12] [--pairs 3:4,7:8]
```

The script pairs each hidden slide with the next visible one, then compares the two with independent layers so one layer's blind spot is covered by another:

- **Block matching**: every original text block is matched to the duplicate as verbatim, cosmetic (case, punctuation, quote style), re-flowed (same words split, merged or moved between boxes), altered (with an inline `{-old-}{+new+}` diff), or missing. Leftover duplicate blocks are "added".
- **Word and number census**: all words and numbers on the slide are counted on both sides, regardless of which box they sit in. This catches a digit changed inside a sentence even if block matching was fooled by restructuring.
- **Non-text content**: hyperlinks (URL and anchor), images (compared by file bytes), alt text, embedded media, speaker notes. Tables, charts (series and values), SmartArt and grouped shapes are all included.

Verdicts per pair: `CLEAN` (nothing differs), `REVIEW` (differences that may be intentional, such as an image swapped or quotes straightened), `FAIL` (content missing or changed). Slide numbers count hidden slides, matching PowerPoint's thumbnail pane.

If `--render` fails, say so and do the visual pass another way (convert to PDF and look at pages); don't skip it.

### 3. Visual pass (the script can't see pixels)

Open each `renders/pair-OO-DD.png` (original left, duplicate right). For each pair, **first list everything visible on the original, then look for each item on the duplicate**. Reading the original first matters because glancing at both side by side makes the eye assume things match. This catches text baked into images or screenshots, charts that were pasted as pictures, logos that were redrawn (the script flags any image whose bytes changed), and anything the XML hid from the script.

Anything you find here that the script didn't flag is an extraction gap: report it as a finding and mention it so the checker can be improved.

### 4. Resolve every flag

For each `FAIL` and `REVIEW` item, re-read the actual strings in `report.json` (`inventory.original` vs `inventory.duplicate`) and confirm it, rather than paraphrasing from memory. Then classify:

- **Confirmed problem**: copy dropped, changed, or added. Quote both versions.
- **Needs the designer's answer**: cosmetic or likely intentional (an image replaced with a new asset, a typographic quote fix, a new label). Say what to confirm, don't dismiss it. The strategist's copy is final, so even a "harmless" edit is the designer's call to explain.
- **False alarm**: explain why (e.g., the checker paired two slides that aren't a real original/duplicate pair; check the pairing warning).

Never downgrade a `FAIL` yourself because it "looks fine". Numbers and links especially: a changed digit looks fine to the eye, which is why this skill exists.

### 5. Write the report

Lead with the answer, then a per-slide table, then detail. Use this shape:

```
Checked N pairs in <deck>: X clean, Y need review, Z have errors.

| Original -> Duplicate | Result | Issues |
| 3 -> 4 | FAIL | Footnote missing; "42%" became "24%" |

Slide 3 -> 4 (FAIL)
- Missing: "Source: Company filings, FY2023." (footnote)
- Changed: "Revenue grew {-42%-} {+24%+} in 2023"
```

Clean pairs get one line. Always include a short **not checked** section: unpaired slides, pairs that look wrongly matched, anything the visual pass couldn't confirm.

Word the verdict carefully. Say "no differences found by the automated check or the visual pass" only when both are true. Don't claim "100% accurate": the honest claim is "every item below was compared and nothing else differed", plus what was out of scope (fonts, layout, colors, animations, and anything the designer intentionally changed are not content and are not checked).

Save the report as `slide-content-qa-report.md` in the outputs folder and share it.

## Notes on judgment

- **Re-flowed text** (INFO in the report) means the exact words are all still present but in different boxes. This is normal in a redesign and is not an error; mention only the count unless a block was split in a way that could change meaning.
- **Designers' additions** such as new section labels or "Source:" lines are `REVIEW`, not errors, but list them: the copy is supposed to be final.
- **Slide numbers, dates and footers** that are auto-fields are ignored, since they legitimately differ.
- If the pairing warning fires (very little shared text), the two slides probably aren't a real pair; ask rather than reporting a wall of false failures.
- Speaker notes live on a separate part of the slide; a duplicate created by a tool that doesn't copy notes silently loses them, which is why the script checks them explicitly.
