# Morning Clips - Automated Daily Media Monitoring Pipeline

Prompt source, formatting macro and version history for an automated daily media
monitoring pipeline. The pipeline converts a single daily PDF of press clips into a
formatted Daily Media Monitoring Report.

This repository holds prompt text, macro code and documentation only. It contains no
client data, no source clip files, no tenant identifiers and no output reports.

---

## 1. Purpose

Replace a manual, analyst-produced daily clips report with a repeatable pipeline that
produces output matching the manual benchmark. The pipeline is assistive: a human reviews
every report before it is issued.

**Input:** one PDF per day containing that day's press clips.
**Output:** plain text, pasted into a Word document where a macro applies all formatting.

The report has two required sections:

1. **Summary** - condensed treatment of each article, grouped and ordered by section.
2. **Full Articles** - verbatim article bodies.

Both sections are required in a live report. Some archived test fixtures were reduced to
Summary-only on purpose, to validate that path in isolation.

---

## 2. Architecture

Single agent, classic orchestration, four inline prompt tools plus a document macro.

```
Daily PDF
   |
   v
[1] Extractor      Python / pypdf via code interpreter
   |               hybrid layout + character extraction
   |               emits [[BR]] paragraph delimiters
   v  RawTextPass
[2] Parser         structures raw text into ===ARTICLE=== blocks
   |               labels: HEADLINE / OUTLET / LINK / DATE / BYLINE / BODY
   v  ParsedBlocksPass
[3] Reporter       generates the Summary section
   |
   v  SummaryPass
[4] FullArticles   generates verbatim full-text bodies
   |
   v  FullArticlesPass
Plain text -> paste into Word -> formatting macro -> final report
```

### 2.1 Component specifications

| # | Tool | Model class | Temp | Output type | Role |
|---|------|-------------|------|-------------|------|
| 1 | Extractor | code interpreter (Python 3.12.x) | n/a | Text | PDF to raw text with paragraph delimiters |
| 2 | Parser | small general model | 0 | Text | Raw text to labelled article blocks |
| 3 | Reporter | large reasoning model | default | **Text** | Summary section |
| 4 | FullArticles | large general model | default | Text | Verbatim article bodies |

Notes:

- Reporter **must** remain on output type `Text`. Switching it to `Document` breaks
  downstream `.text` extraction.
- FullArticles was moved from the small model to the large model. The small model has an
  output-length ceiling that silently abridges large batches without erroring.
- Extractor uses layout mode for paragraph-boundary detection and character mode for the
  text itself. The `layout_mode_scale_weight` parameter has no effect once the
  space-collapsing step runs - do not tune it.
- The code interpreter is a real Python execution environment in the agent test chat, not a
  simulation of one.

### 2.2 Variable contract

After every prompt-tool invocation, extract the text explicitly:

```
SetVariable Global.ParsedBlocksPass = Text(Global.parsedBlocks.text)
```

Applies at each stage. Skipping the `Text()` extraction leaves an object where a string is
expected, and the next stage receives an empty or malformed input.

### 2.3 FullArticles design

FullArticles takes two inputs:

- **SummaryPass** - drives section membership and article order.
- **ParsedBlocksPass** - supplies the verbatim body text.

Format requirements: dash-bulleted headlines, no URL line, section headers printed only for
populated sections, document terminates with `-ENDS-`.

---

## 3. Formatting macro

`FormatMorningClips` - Word VBA, stored in `Normal.dotm`. Current version: v11.

All formatting, numbering, structural cleanup and keyword highlighting is applied
deterministically by the macro, not by any model.

### 3.1 Environment prerequisite

One-time, per machine: **AutoCorrect -> AutoFormat As You Type -> uncheck "Automatic
numbered lists".**

This cannot be set programmatically from the macOS build of VBA. It must be done by hand on
every machine that runs the macro. Without it, Word re-numbers pasted lists and overwrites
the macro's numbering.

### 3.2 Transfer

Export and import via `.bas` file only. Never transfer by copy-paste. Recipients on Windows
must unblock the downloaded file in file properties before importing.

### 3.3 Disabled code

`RepairLongParagraphs` is commented out, not deleted. It was confirmed to split a single
long paragraph six ways. Leave it disabled.

### 3.4 Pending macro work

Full Articles section detection is not yet implemented. It requires handling:

- un-numbered headlines
- byline lines
- long-form date lines
- the `-ENDS-` terminator

---

## 4. Keyword categories and highlighting

Three fixed tracked categories. The actual terms are held outside this repository in the
runtime configuration; only structure is documented here.

| Category | Highlight colour | Contents |
|----------|-----------------|----------|
| Cat1 | Green | one named individual |
| Cat2 | Yellow | the client organisation and its abbreviation |
| Cat3 | Turquoise | eleven peer and competitor firms |

**Highlighting rule:** highlight the exact tracked string only. Never extend the highlight
to surrounding words, to a title, or to a longer version of the name.

---

## 5. Summary-section editorial rules (Reporter)

Prompt structure:

```
A1 - A5    ALWAYS rules, hoisted to top level
SHAPE      ROUNDUP / COMMENT / SINGLE STORY
S1         funding gate
S2         case B
S3         head
S4         tail
QUOTES
BOILERPLATE
```

### 5.1 Structural rule about the prompt itself

**Unconditional keeps must be top-level minimal single lines. Never as clauses inside a
numbered step, and never with emphatic or qualifying language.**

A1 is a global rule, not scoped to one shape. Rewording it produces spillover across every
shape at once, in both directions - one revision improved two articles and regressed a
third, with four confirmed spillovers across three shapes in a single version. Treat any
edit to A1 as a full-batch regression risk.

### 5.2 Borderline paragraph test

Remove the tracked name from the paragraph. If it still reads as standalone general
context, drop the paragraph - unless an explicit override applies, or it is a
named-regulator or company-status paragraph.

Tolerance: over-keeping or under-keeping by one paragraph is generally acceptable against
the benchmark.

---

## 6. Version history

### Reporter

| Version | STEP3 chars | Full prompt chars | Batch score | Note |
|---------|-------------|-------------------|-------------|------|
| pre-rewrite | - | - | 5 / 14 | baseline |
| v2 | - | - | 12 / 19 | |
| v3 (draft) | - | - | not run | superseded, partially absorbed into v3b |
| **v3b** | ~9,220 | ~14,059 | **15 / 19** | current installed version |
| v3c | ~9,081 (-139) | - | planned | single-line A1 revert |

**v3b batch detail**

| Batch | Articles | Result |
|-------|----------|--------|
| Batch 1 | 6 | 5 / 5 acceptable |
| Batch 2 | 5 | 5 / 5 acceptable, four exact |
| Batch 3 | 6 | 4 / 6 - one partial, one regression |
| Batch 4 | 3 | 1 / 3 - two blocked upstream by the parser defect |

**v3 draft edits, superseded:** A1 dek carve-out, R1 two-condition dek test, R2 strict
source-order plus bridge clause, R5 roster-header locator.

**v3c planned change:** revert A1 to minimal wording -

> A1 ALWAYS keep body paragraph 1 (the lead), except a ROUNDUP dek under R1, where the next
> paragraph becomes the lead instead.

Expected to reverse all five confirmed A1 spillovers. Accepted trade-off: one article
regresses by one paragraph, which is within tolerance.

---

## 7. Open defects

| ID | Defect | Root cause | Fix target |
|----|--------|-----------|------------|
| A | One article regressed 5 -> 9 paragraphs on v3b | A1 rewrite spillover outside ROUNDUP; the emphatic clause weakened the S3 head stop | Reporter (v3c) |
| B | Two articles in batch 4 unusable | Parser corrupts `[[BR]]` to `[BR]]` from the second or third break onward; Reporter then receives one blob per article and re-segments by eye | **Parser** |
| C | Batch 3 structural fidelity: stripped rank labels, flattened stat blocks, dropped roster headers | same token corruption as B | **Parser** |
| D | Duplicate `1.` numbering on two articles | numbering step | Step 4 |
| E | Article order flipped within a section | ordering step | Step 4 |
| F | One article keeps an extra product-launch paragraph | within tolerance | not fixing |

### 7.1 The `[[BR]]` corruption is the root cause of run-to-run variance

Extractor output has clean `[[BR]]` tokens. Parser output degrades from the second or third
break onward. Every observed instance of non-deterministic output and structural fidelity
loss traces back to this.

**Do not attempt a Reporter-side tolerance patch.** This was tried and caused timeouts on
the largest batch. It was reverted. The Parser is the correct fix target.

---

## 8. Roadmap

In sequence:

1. **Parser `[[BR]]` token fix.** Blocks defects B and C, and all further Reporter tuning on
   the affected articles.
2. **Reporter v3c.** Single-line A1 revert. Re-run batch 4 to confirm the two blocked
   articles unblock; re-run the full batch to confirm defect A returns to benchmark and no
   new spillovers appear.
3. **Step 4 fixes.** Duplicate numbering; intra-section article ordering.
4. **FullArticles build.** Begin once the Reporter prompt is stable.
5. **Macro update** for Full Articles section detection.

---

## 9. Hard constraints

Architectural decisions already tested and closed. Do not reopen without new evidence.

- **No workflow-automation platform, no low-code app front end, no HTTP connectors.** Single
  agent only.
- **Classic orchestration with inline prompt tools is the only viable architecture.**
  - Connected agents strip data bindings at the boundary: *"variable data type not eligible
    to receive or return values"*.
  - Global variables fail across agent boundaries: *"Identifier not recognised"*.
  - Generative orchestration causes session drift.
- **All formatting stays in the macro.** No model is asked to produce formatted output.
- **Human review before issue.** Every report is checked by a person against the benchmark.
- Premium agent licensing is available; premium workflow-automation licensing is not.

---

## 10. Conventions

### Naming

- Prompt tools: PascalCase, named for their function, no positional numbers - `Extractor`,
  `Parser`, `Reporter`, `FullArticles`.
- Variables: `Pass` suffix - `RawTextPass`, `ParsedBlocksPass`, `SummaryPass`,
  `FullArticlesPass`.

### Prompt files

- `.txt`, never `.md`. One file per prompt.
- Version in the filename: `Reporter_v3b.txt`.
- No metadata inside the prompt file. Character counts, changelog and scorecards live in
  this README only.
- `.gitattributes` enforces `eol=lf`.

### Change discipline

- Surgical edits only. Minimal changes that do not perturb already-validated behaviour.
- Every prompt edit gets an explicit changelog entry and a character-count delta.
- Full section rewrites, not partial diffs.
- Each component is validated in isolation against a named benchmark batch before being
  chained.
- Regression traces are always cited against a named benchmark batch.

### Revert protocol

If a change regresses categorisation or reintroduces `[[BR]]` leakage, restore the saved
baseline verbatim. Do not patch forward.

### Git

- Always `git add <folder>`. Never `git add -A` - it cross-contaminates commit messages
  across components.
- Repository is private.

---

## 11. Testing

Four benchmark batches, produced manually and treated as the quality bar: 6, 5, 6 and 3
articles respectively. Fixture files are not committed.

Procedure per version:

1. Run all four batches.
2. Retain both the raw `.txt` output and the formatted `.docx` for combined content and
   formatting review.
3. Score per article: exact / acceptable / partial / fail.
4. Record the scorecard in section 6 of this README.

Formatting differences cannot be diagnosed from pasted plain text alone - plain-text pastes
strip Word auto-numbering and paragraph breaks. Always review the `.docx`.

---

## 12. Not in this repository

- Client data, source PDFs, generated reports
- Tracked keyword terms
- Tenant, environment and model instance identifiers
- Benchmark fixture files
