# Morning Clips - Automated Daily Media Monitoring Pipeline

**Status: complete and in daily production use.**

Prompt source, formatting macro and operating documentation for an automated daily media
monitoring pipeline. The pipeline converts a daily PDF of press clips into a formatted
Daily Media Monitoring Report in Word, replacing a fully manual process.

Validated against four manually produced benchmark reports and confirmed on live daily
runs. Every report passes human review before it is issued.

This repository holds prompt text, macro code and documentation only. No client data, no
source clips, no generated reports, no tenant or model-instance identifiers, no tracked
keyword terms, no client branding.

---

## 1. What it does

**Input:** one PDF per day containing that day's press clips. PDF is the only accepted
input format - the platform's file handling does not accept `.docx`.

**Output:** a single plain-text block returned in chat, pasted into a Word starter
template and formatted in one pass by a VBA macro.

The report contains two sections:

1. **Summary** - each kept article categorised into one of three fixed sections,
   numbered, trimmed to the paragraphs relevant to the tracked names, with tracked
   keywords highlighted by category colour.
2. **Full Articles** - verbatim full bodies of the kept articles, dash-bulleted
   headlines, byline, date, no URL line, section headers only for populated sections,
   terminating in `-ENDS-`.

Relevance to the tracked names is the only reason to keep text. The operator supplies the
report date; it is never inferred from article dates.

---

## 2. Architecture

Single agent, classic orchestration, four inline prompt tools chained through global
variables, plus a Word macro for all formatting.

```
Daily PDF
   |
   v
[1] Extractor      Python / pypdf via code interpreter
   |               hybrid extraction: plain-mode text, layout-derived paragraph cuts
   |               emits [[BR]] paragraph delimiters
   v  RawTextPass
[2] Parser         structures raw text into ===ARTICLE=== blocks
   |               labels: HEADLINE / OUTLET / LINK / DATE / BYLINE / BODY
   v  ParsedBlocksPass
[3] Reporter       categorises, trims, builds the Summary section
   |
   v  SummaryPass ------------+
[4] FullArticles              |  <- section membership and article order
   |  <- ParsedBlocksPass     |  <- verbatim body text
   v  FullArticlesPass
Single SendActivity (code-fenced)
   |
   v
Paste into starter template -> run FormatMorningClips once -> human review -> issue
```

### 2.1 Components

| # | Tool | Engine | Temp | Output type | Role |
|---|------|--------|------|-------------|------|
| 1 | Extractor | code interpreter, Python 3.12.x, pypdf | n/a | Text | PDF to raw text with paragraph delimiters |
| 2 | Parser | small general model | 0 | Text | Raw text to labelled article blocks, verbatim |
| 3 | Reporter | large reasoning model | default | **Text** | Categorisation, trimming, Summary build |
| 4 | FullArticles | large general model | 0 | Text | Verbatim article bodies |

Fixed decisions behind this table:

- Reporter output type must stay `Text`. `Document` output cannot carry real paragraph
  marks, bold or highlighting (platform limitation) and breaks the downstream `.text`
  extraction.
- The Parser stays on the small model at temperature 0. Reasoning models time out on
  large mechanical restructuring tasks.
- FullArticles runs on the full-size model, not the mini variant. The mini variant has an
  output-length ceiling that silently abridges the largest batches without erroring.
- The agent's own model setting has no effect on output - each prompt tool carries its
  own model.
- Each prompt field has a ceiling of roughly 8,000 characters. Every prompt in this
  repository was written against that budget; adding a rule means removing one.

### 2.2 Main flow

`OnRecognizedIntent` trigger (trigger phrases for "generate morning clips") ->
six `Blank()` variable clears -> Question (source file) -> Question (report date) ->
Extractor invoke -> Parser invoke -> Reporter invoke -> FullArticles invoke -> single
`SendActivity` -> end.

Two contracts that make the chain work:

1. **Text extraction at every hop.** After each prompt-tool invocation:
   `SetVariable Global.ParsedBlocksPass = Text(Global.parsedBlocks.text)`. Without the
   `Text()` wrapper an object is passed where a string is expected and the next stage
   receives nothing usable.
2. **`[[BR]]` is the universal break token.** Inter-tool handoffs never carry real
   newlines - they collapse in transit. `[[BR]]` marks every field and paragraph
   boundary; the Reporter and FullArticles convert to real breaks only at final output,
   emitting a blank line between paragraphs so the chat renderer's markdown handling
   cannot collapse them.

### 2.3 Output message wrapping

The `SendActivity` output is wrapped in a code fence, written in single-line quoted YAML.
Without the fence, the chat renderer parses the report as markdown: ordered lists are
renumbered (`12.` becomes `1.`) and rank labels, stat-block fields and entry labels are
damaged. The fence disables all of that.

### 2.4 Extractor normalisation

The Extractor takes character text from pypdf plain mode and paragraph boundaries from
layout mode. Its `norm()` pass applies: NFKD unicode normalisation; stem-anchored
ligature repair (sparing camelCase brand names); a space-before-period repair for a
character-mode extraction defect; URL line-wrap repair; line-wrap newlines to single
spaces with multi-space collapse; paragraph and page boundaries to `[[BR]]`; a
comma/semicolon join rule healing mid-sentence page splits; a width-based
paragraph-boundary rule (median line length) separating absorbed sub-headings; and
`-ENDS-` preserved as a standalone segment.

Do not switch the text source to layout mode - it inserts spaces inside URLs across the
whole batch. `layout_mode_scale_weight` has no effect after the space-collapsing step.

---

## 3. Categorisation and highlighting

### 3.1 Sections and precedence

Three fixed sections, strict precedence, one section per article:

| Section | Category | Highlight | Contents |
|---------|----------|-----------|----------|
| i | Cat1 | Green | one tracked individual |
| ii | Cat2 | Yellow | the client organisation, its abbreviation, and its former entity name |
| iii | Cat3 | Turquoise | a fixed list of peer and competitor firms |

Any Cat1 keyword -> section i; else any Cat2 -> section ii; else any Cat3 -> section iii;
no tracked keyword -> article excluded. Cat2 presence anywhere in an article forces
section ii regardless of how many Cat3 names appear.

The keyword terms themselves live in the runtime configuration, not in this repository.

### 3.2 Matching rules

These rules are the product of the hardest debugging in the project. Their exact wording
matters:

- **Section iii is the default; sections i and ii are exceptions.** Each exception
  requires a verifiable literal string match inside that article's own body.
  Prohibition-style wording ("never categorise by topic") failed; a literal string lookup
  holds.
- **Matches count in any body segment, not only prose sentences** - this covers roster
  lines of the form `Name | Firm | #rank`.
- **Whole-phrase matching only.** A keyword never matches inside an unrelated longer name
  or as a fragment of a common word.
- **Listed variants match; unlisted aliases do not.** Short forms of a tracked firm's
  name are matched via an explicit variant list, case-insensitively.
- **Per-article independence.** Every article is categorised on its own body alone.
  Without this instruction, articles in a batch bleed into each other's section
  assignments.

### 3.3 Highlighting

Every occurrence of every tracked keyword is highlighted in its own category colour,
including keywords from other categories appearing inside an article. Only the exact
keyword string is highlighted - never surrounding words, a title, or a longer form of the
name. Applied deterministically by the macro (`HiliteTerm`, `MatchCase = False`), never
by a model.

---

## 4. Summary editorial logic (Reporter)

Prompt structure: CONFIG -> STEP1 CATEGORISE -> STEP2 HIGHLIGHT -> STEP3 TRIM ->
STEP4 BUILD -> OUTPUT CONTRACT -> GUARDRAILS, with unconditional ALWAYS rules hoisted to
top level.

STEP3 decides an article's shape and trims accordingly:

- **SINGLE STORY** (default): keep the contiguous core block relevant to the tracked
  name, drop the tail. CASE A (tracked name is a party to the event) and CASE B (tracked
  name is only a background owner or parent) set how much survives; an OVERRIDE
  guarantees every tracked name surfaces at least once; a funding gate keeps short
  backer-funding stories whole.
- **ROUNDUP** (only when the article is built as parallel ranked or profiled entries and
  that list is the article's subject): keep the opening narrative and every entry
  containing a tracked keyword, verbatim, with source rank labels copied exactly - never
  renumbered. A list of names inside a prose story is not a roundup.

Rules about editing this prompt, learned the hard way:

- **Unconditional keeps are top-level minimal single lines** - never clauses inside a
  numbered step, never with emphatic or qualifying language. The global lead-keep rule
  was reworded once and produced five confirmed regressions across three shapes in a
  single version.
- **Stop editing when it works.** Every late-stage regression in this project came from a
  change chasing something that was not broken. The trimming logic, build rules and
  section structure that went untouched are the parts that never failed.

Accepted output bar: wrong categorisation is a failure; over- or under-keeping by about a
paragraph is within tolerance; omissions that are faithful to the source are not defects;
empty LINK fields on some source batches are source-faithful, not defects.

---

## 5. Formatting macro

`FormatMorningClips` - Word VBA, single module, stored in `Normal.dotm`, transferred as a
`.bas` file. All visual formatting, numbering, highlighting and structural cleanup is
deterministic and lives here. No model produces formatted output.

The macro is **one-shot and not idempotent** - run it exactly once on a fresh paste.

### 5.1 What it does, in order

Early text repair: strip the Reporter's own headline numbers (the macro applies its own);
`RepairArtifacts` (eaten spaces after full stops, broken hyphens, straight-to-curly
apostrophes - skips URL-containing paragraphs); junk-character cleanup (zero-width
spaces, BOM, non-breaking spaces); hyperlink field codes unlinked to plain URL text;
URL space repair; content-control and form-field removal.

Structural pass: title/byline/date splits; roster-line splits; doubled-entry-header
splits; horizontal-rule shape removal; roster rank-number regeneration (Word AutoFormat
strips them); attribution-roster pruning with a Full Articles boundary guard; colour-tag
stripping; social-channel label normalisation; the divider placed before the
`FULL ARTICLES` heading as a fixed-length em-dash line with paragraph borders cleared.

Numbering and layout: both sections numbered with a counter reset between them, as typed
text (not Word list objects); Summary indents and alignment; Full Articles flush left at
1.25 line spacing; Full Articles headlines converted to dash format last, after
numbering.

Highlighting: category-coloured keyword highlighting as in section 3.3.

### 5.2 Measured layout

All values measured from the manual benchmark document, applied in centimetres via
`CentimetersToPoints()`: sections 1.31/0.67, articles 1.94/1.31, social channels
1.27/0.64, "no relevant news" lines aligned under the section name rather than the
numeral. Margins 2.49 top, 2.54 bottom, 3.17 left and right. Sections, socials and
headlines use hanging indent plus tab stop so names land at a fixed position regardless
of numeral width.

### 5.3 Coding rules in this module

- No line-continuation characters anywhere - they corrupt on transfer.
- `ListFormat.RemoveNumbers` per paragraph, never document-wide (unreliable on Windows).
- `MatchWildcards = False` explicitly on every literal find; `wdFindStop`, never
  `wdFindContinue` (prevents infinite re-matching); every loop progress-guarded.
- `RepairLongParagraphs` is **commented out, not deleted** - it split one long paragraph
  six ways. It stays disabled.

---

## 6. Environment setup (per machine, one time)

1. **Word: AutoCorrect -> AutoFormat As You Type -> uncheck "Automatic numbered lists."**
   Cannot be set programmatically from the macOS build of VBA. Without it, Word renumbers
   the pasted report and overwrites the macro's numbering.
2. Import the macro into `Normal.dotm` via `.bas` Import File - never by pasting code.
   Windows users must unblock the downloaded `.bas` first (Properties -> Unblock).
3. Install the starter template: a Word document carrying the client letterhead as an
   inline header image. The template and logo are not in this repository.

---

## 7. Daily operating procedure

1. Trigger the agent with a trigger phrase.
2. Upload the day's clips PDF and give the report date when prompted.
3. Copy the single returned output block (inside the code fence).
4. Paste into a fresh copy of the starter template.
5. Run `FormatMorningClips` once.
6. **Human review before issue** - check categorisation and highlighting, and remove any
   photo captions. Caption removal is deliberately manual: it is pattern-matching on
   prose, and a human glance is safer than a rule that could catch a real sentence.

---

## 8. Validation

Benchmarked against four manually produced reports (6, 5, 6 and 3 articles) and confirmed
on live daily runs.

**Final state: all 19 benchmark articles coherent, on-topic and correctly categorised,
six exact matches to the benchmark, and live runs producing correct Summary and Full
Articles sections end to end.** Categorisation - the hard acceptance bar - is verified
across every batch: whole-phrase matching with no false positives, correct precedence,
per-article independence, correct colour-coded highlighting. Layout feedback from the
pilot review (indentation, spacing, letterhead) is closed.

Testing discipline used throughout, to be kept for any future change:

1. Run a changed batch **twice back-to-back** before attributing any difference to the
   change - the pipeline has inherent run-to-run variance.
2. One change at a time; full batch re-run before the next edit.
3. Retain both the raw `.txt` and the formatted `.docx`. Formatting is not diagnosable
   from pasted plain text - pastes strip Word numbering and paragraph breaks.
4. Diagnose from file data, never from screenshots.
5. Judge against the acceptance bar in section 4, not against byte-identity with the
   benchmark.

---

## 9. Known and accepted limitations

Documented so they are not mistaken for defects, and not "fixed" into regressions:

- **Delimiter degradation on the longest roundups.** `[[BR]]` tokens are clean out of the
  Extractor and can degrade in the Parser on the largest batches, which is the origin of
  the residual run-to-run variance. Every repair approach tried (a dedicated repair
  stage; Reporter-side tolerance) exceeded the runtime budget on publish and was
  reverted. Current behaviour is within the accepted bar. Leave it alone unless output
  actually fails review.
- **Original in-body paragraph boundaries are unrecoverable** once the source PDF does
  not carry them; reconstructed breaks are plausible, not identical. A small number of
  mid-sentence splits at page boundaries are known and accepted.
- **Embedded tables render as flattened tab-separated text**, matching the text form of
  the manual benchmark. No visual table reconstruction.
- **Numbering is typed text**, not Word `ListFormat` objects. Visually identical to the
  benchmark; a list-object implementation was assessed and not needed.
- **The tracked individual's section and the social-channels block** are format-verified
  but have not yet been exercised by a live day's input containing that coverage.
- **Photo captions** are removed by the human reviewer by design (section 7).
- Test-pane success does not guarantee published success - runtime limits differ.
  Anything that passes in test must be confirmed published.

---

## 10. Closed design decisions

Tested and settled. Do not reopen without new evidence.

- **Single agent, classic orchestration, inline prompt tools.** Connected agents strip
  data bindings at the boundary ("variable data type not eligible to receive or return
  values"); global variables do not cross agent boundaries ("Identifier not recognised");
  generative orchestration causes session drift and bleeds state between runs.
- **No workflow-automation platform, no low-code app front end, no HTTP connectors.**
  The available licensing covers the agent platform only; the macro also requires desktop
  Word regardless, so an automated document hand-off buys nothing.
- **PDF-only input.** A `.docx`-based Extractor is a dead end on this platform.
- **Labelled plain text with `===ARTICLE===` delimiters for handoff, not JSON.** JSON is
  fragile with article body content and unnecessary between language models.
- **All formatting in the macro, never in a prompt.** A tag-based highlighting scheme was
  assessed and rejected: it returns highlighting to the model and adds a stripping step
  for something the macro does deterministically.
- **Report date is operator-supplied**, set in a deterministic expression, never
  model-inferred.

### 10.1 Approaches tried and rejected

Recorded so they are not re-attempted:

- Deterministic delimiter-repair stage between Parser and Reporter - passed in test,
  timed out on publish.
- Reporter-side tolerance for malformed delimiters - timed out on the largest batch.
- Parser step forcing re-paragraphing of oversized segments - fired inconsistently
  between runs (model variance); prose re-paragraphing cannot be made deterministic in a
  prompt. Fully reverted; deterministic re-paragraphing belongs in the macro if anywhere.
- Semantic-rules Extractor rewrite - destroyed the paragraph-break signal; reverted to
  the mechanical rules.
- Layout-mode text source in the Extractor - broke every URL in the batch; reverted to
  plain-mode text with layout-derived cuts.
- A six-part rewrite of the Reporter trim step - regressed paragraph counts badly;
  reverted to the confirmed baseline, then improved by single surgical edits only.
- `.docx` Extractor rewrite - ruled out by the input constraint.

---

## 11. Troubleshooting

| Symptom | Cause |
|---------|-------|
| `12.` renders as `1.`; rank labels or stat-block fields damaged in chat | output not code-fenced; renderer is parsing markdown |
| Numbering overwritten after paste into Word | machine missing the AutoCorrect setting (section 6) |
| Spaces inside URLs (`w w w`) across a whole batch | Extractor text source switched to layout mode |
| Literal colour tags in the report text | Reporter prompt has regained a plain-text / highlight contradiction |
| A tracked firm routed to section iii instead of ii | precedence or whole-phrase matching wording altered |
| Section assignments bleed between articles | per-article independence instruction missing |
| One long paragraph split many ways | `RepairLongParagraphs` re-enabled |
| Whole articles missing from Full Articles on a large batch | FullArticles moved back to the mini model |
| Empty URL line, or an omission versus the benchmark | check the source first - both are usually source-faithful |
| Macro run twice on one document | not recoverable - re-paste and run once |

---

## 12. Repository conventions

- Prompt tools: PascalCase, named for function, no positional numbers - `Extractor`,
  `Parser`, `Reporter`, `FullArticles`.
- Global variables: `Pass` suffix - `RawTextPass`, `ParsedBlocksPass`, `SummaryPass`,
  `FullArticlesPass`.
- Prompts as `.txt`, one file per prompt, version in the filename (`Reporter_v3b.txt`),
  no metadata inside the file. Macro as `.bas`. `.gitattributes` enforces `eol=lf`.
- Change discipline, if anything is ever revisited: surgical edits only; full section
  rewrites, not partial diffs; a changelog entry and character-count delta per edit;
  regressions traced against a named benchmark batch; and if a change regresses
  categorisation or reintroduces delimiter leakage, restore the saved baseline verbatim -
  never patch forward.
- `git add <folder>` always; never `git add -A`.
- Repository is private.

---

## 13. Not in this repository

- Client data, source PDFs, generated reports
- Tracked keyword terms and variant lists
- Client letterhead and starter template
- Tenant, environment and model-instance identifiers
- Benchmark fixture files
