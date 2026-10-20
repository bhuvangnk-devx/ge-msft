# Root-cause worktree quality results — 20 October 2026

Worktree: `fix/root-cause`. The additional fixes tested here were checkpointed in `ad4c1786aee185d91f18db0d5b5b7f5c03389d78` and integrated into local `cimb-cicd` on 20 October 2026. Nothing from this session has been pushed or deployed. See [the integration record](ROOT-CAUSE-CIMB-INTEGRATION-2026-10-20.md) for the local-change comparison and preservation details.

## Scope and evidence

These runs use live Gemini Enterprise (`gemini-4-argon`) with **simulated Office hosts** and automatically accepted approval gates. They exercise the real pane, planner, runtime and bridges. They do not certify real Office rendering, durable writes, Outlook delivery or deployed behavior.

The requested sample is two scenarios per quality-gate group. Excel has one greeting-only group with just one scenario. A group is the gate's classification, not a complete inventory of every API capability.

| Surface | Original sample | Outcome audit and follow-up |
| --- | --- | --- |
| Excel | 33/35, 18 groups | Two routing failures fixed; both plus a read-only control passed 3/3 in a focused retest. Full 35 were not rerun. |
| Word | 21/22, 11 groups | One reply request omitted the decision needed to answer a comment. The model asked for it. A separate explicit-answer follow-up passed 1/1; original 21/22 retained. |
| PowerPoint | Raw 26/26, 13 groups | Semantic audit found incorrect insertion position and unsupported data. Position follow-up passed 5/5; factual-table RCA follow-up passed 4/4. Raw score retained. |
| Outlook | Raw 12/12, 6 groups | Both email-reading questions were false passes: body context was missing. Fixed; unchanged original prompts and two independent document/wording variants passed 4/4 with factual-answer checks. |

Original evidence:

- [Excel original sample](../scenarios/.results/quality-excel-2026-10-20T02-20-36Z/EXCEL-QUALITY-RESULTS.md) and [focused retest](../scenarios/.results/excel-routing-retest-2026-10-20T02-53-30Z/EXCEL-QUALITY-RESULTS.md).
- [Word sample](../scenarios/.results/quality-word-2026-10-20T02-52-48Z/WORD-QUALITY-RESULTS.md) and [explicit-answer follow-up](../scenarios/.results/word-reply-clarified-2026-10-20T03-04-49Z/WORD-QUALITY-RESULTS.md).
- [PowerPoint original sample and semantic audit](../scenarios/.results/quality-powerpoint-streaming-2026-10-20T03-25-01Z/POWERPOINT-QUALITY-RESULTS.md).
- [Outlook original sample and semantic audit](../scenarios/.results/quality-outlook-2026-10-20T03-12-40Z/OUTLOOK-QUALITY-RESULTS.md) and [four-case follow-up](../scenarios/.results/outlook-semantic-retest-2026-10-20T04-00-14Z/OUTLOOK-QUALITY-RESULTS.md).

These result directories contain private transcripts and are ignored by Git. They remain under `scenarios/.results` in the RCA checkout at `/Users/gnkbhuvan/Desktop/devx/ge-msft/.claude/worktrees/root-cause`; the relative evidence links resolve there. They are not copied into the target checkout by the Git merge. This summary records their conclusions; it does not publish the transcripts.

## Causes and changes

### Excel: ambiguous command extraction and chat routing

Natural language such as “Set A21 to …” was accepted as literal CLI input, including the word `to`. A mixed analytical request with a note could also take a chat-only shortcut and omit the requested write.

[Direct-command extraction](../packages/web-shell/src/taskpane/direct-command.ts) now sends ambiguous unquoted prose to the planner, preserving quoted values, formulas, scalars and fenced commands. Ambiguity prevents partial extraction from a multiline request. [Free-text routing](../packages/web-shell/src/taskpane/components/App.tsx) uses the structured planner decision rather than question/action keyword shortcuts on writable surfaces. Explicit commands and explicit chat intents keep their routes.

The unsafe external formula example was written as a literal string in the original failure; it was not observed executing as an external formula. In the retest it produced no cell change. The note case produced a comment at `Sheet1!D7` through the approval gates; the read-only control made no changes.

### Outlook: a search result was mistaken for enough context

[`AssistSession.chatReads`](../packages/runtime/src/assist-session.ts) previously searched the user's words and distinctive terms. A general summary found nothing; a question about Rina found only her signature. The structural snapshot did not include the body. Both runs completed without writing, which satisfied the weak original expectations despite missing the requested answers.

Chat now retains bounded, query-targeted matches and fills the remaining context budget with the active email's resolved body. Keeping targeted matches first avoids losing relevant text near the end of a long email. This reads the active item only. Context remains temporary and framed as untrusted data.

Real bridge regression tests cover a summary, a name matching only the signature, and a deadline after 900 sentences. Live follow-ups check actual facts from two unrelated emails as well as the two unchanged original requests. No reply or new-message form was created in those read-only runs.

### PowerPoint: selected slide omitted from execution context

The command language already supported `at=N`; the compiler and insertion bridge were not missing that feature. The compact bootstrap exposed its syntax but did not always carry its canonical explanation. Adding that explanation alone did **not** fix the original case: the first diagnostic batch still appended the slide.

The deeper context defect was that [`PowerPointBridge.captureDocState`](../packages/bridge-powerpoint/src/powerpoint-bridge.ts) listed all slides but omitted which slide was selected. It now supplies selected slide IDs and one-based positions from captured deck order. The optional selection read is capability-gated, handles index zero and multiple selections, and preserves the deck snapshot if selection retrieval fails. It does not introduce a new use of the PowerPointApi 1.8 `Slide.index` property.

The same five prompts are used for the final follow-up: original after-current request, after-first, before-selected, absolute position, and default append. They use different deck lengths and selected positions. No wording or expected position was changed between diagnostic batches.

Final five-case live follow-up: **5/5 passed**, with original slide IDs/order preserved and inserted positions verified against immutable seeds. See the [separate follow-up report](../scenarios/.results/powerpoint-selection-retest-2026-10-20T04-10-37Z/POWERPOINT-QUALITY-RESULTS.md). This is a focused diagnostic, not a rerun of all 26 scenarios.

### Quality harness: hidden stalls, mutable seeds and insufficient assertions

- [`liveFetch`](../packages/web-shell/src/scenarios/live-env.ts) awaited an entire cloned response body before returning headers. It now returns the original stream immediately and records response state separately. Controlled stream tests prove this bug; the initial live stall had insufficient saved evidence to establish its exact cause independently.
- [`geminiRefusal`](../packages/web-shell/src/scenarios/runner.ts) preserves local command outcomes and executed HTTP 200 stream failures/deadlines, instead of hiding them as “not run.” Recorded transcripts are copied before returning the verdict.
- The simulator mutated the scenario's original seed. Hosts now receive a clone so before/after checks and reruns retain the declared fixture. A local-command regression checks actual insertion with network access blocked.
- `insertedSlidePositions` checks new-slide positions and preservation of all original IDs in order. Negative tests reject append, deletion and reordering. Outlook read expectations now require source facts as well as absence of writes. These changes strengthen the gate; they do not make expectations more permissive.

The initial PowerPoint live attempt stopped incomplete, and its evidence is [preserved](../scenarios/.results/quality-powerpoint-2026-10-20T03-12-40Z/). The next full sample produced the raw 26/26 score above. The first five-case diagnostic produced raw 0/5 because of the mutable-seed oracle bug; [offline assessment against the immutable fixtures](../scenarios/.results/powerpoint-semantic-retest-2026-10-20T04-00-14Z/immutable-seed-rescore.json) shows three correct placements and two real failures. This offline assessment is not a new live run.

## Verification and limits

Latest RCA static verification: **3,626 tests passed, 16 skipped**; typecheck, lint and brand guard passed. This includes two additional compact/full prompt repair regressions from integration review. Production build and generated-skill consistency passed. Independent review found no remaining important findings after the source-fact, row-association, API-version and follow-up instruction corrections.

All raw failures and original verdicts remain available. Follow-ups are separate, selected in advance of their runs, and reported as focused evidence. A successful retry is not a full-sample rerun or an estimate of production reliability. No prompt-specific routing, model substitution, timeout increase or engine-setting change was introduced to improve these scores.

An incorrectly supplied `--static-only` option once triggered an uncredentialed default live attempt after passing static tests. It made zero Gemini queries and yielded no usable scenario results. It is excluded from the evidence; subsequent static checks used the supported `--skip-live` option.

Findings raised by the initial audit and their current disposition:

- **Resolved:** `mx-pp-tableslide-17` previously populated owner roles and Q4 due dates that the deck did not establish. The unchanged request now keeps missing facts as `TBD` and preserves only the source's December deadline. See the additional RCA below. The separate Q3/Q4 revenue table case explicitly authorized example figures through clarification; those figures remain synthetic examples.
- **Resolved in code and strict simulated-host tests:** read ports no longer load/use native `Slide.index` under gates below 1.8. Microsoft documents that property in [PowerPointApi 1.8](https://learn.microsoft.com/en-us/javascript/api/requirement-sets/powerpoint/powerpoint-api-1-8-requirement-set). Positions now come from observed stable IDs and deck order. This does not certify every feature on every older Office host.
- Passing unsupported-operation scenarios establishes honest refusal, not support for slide reordering, image insertion, calendar scheduling or sending mail.
- The plan's full Flash-versus-Argon parity and real-host release certification have not been established by this two-per-group Argon sample.

The work is committed and integrated locally. A release decision still requires checking the exact source intended for release; this local integration does not establish deployment or real-host certification.

## Additional RCA: table facts and older-host reads

### Conflicting table-completion instructions

The original table failure was encouraged by an instruction conflict. The compact bootstrap and canonical capability registry demanded that every cell be filled from the slide's topic, while other instructions prohibited inventing facts. The same completion pressure propagated through targeted help, full grammar and generated skill metadata. The original failure had no clarification gate: the fixture's example-content answer was never supplied and therefore did not authorize invented owners or dates.

`TABLE_FACTS_RULE` in [capability-registry.ts](../packages/contracts/src/capability-registry.ts) now supplies one consistent rule to canonical table help, the PowerPoint planner and compact executor instructions: use observed/user-supplied facts, retain source precision, keep missing facts blank/TBD, and require explicit authorization plus visible labelling for synthetic examples or proposals. The conflicting fill-every-cell instruction was removed from both active sources. Skill instructions were aligned and generated resources refreshed.

Outcome checks now inspect native `shape.table` cells, including owner/date associations with their action rows, exact row counts and populated example rows. Negative regressions reject invented cells, swapped facts, deadline propagation, header-only tables and blank samples.

[Four-case live follow-up](../scenarios/.results/powerpoint-facts-rca-2026-10-20T04-23-01Z/POWERPOINT-QUALITY-RESULTS.md): **4/4 passed**, 14 recorded Argon calls, followed by saved-table semantic inspection:

| Case | Actual written outcome |
| --- | --- |
| Unchanged `mx-pp-tableslide-17` | All owners `TBD`; SME underwriting and training deadlines `TBD`; MFA deadline `December`. No clarification or additional example authorization. |
| Independent missing-fact control | Three actions retained; owners `TBD`; only the generator action has `November`, with no inferred day/year. |
| Independent supplied-fact control | Both actions retain their exact corresponding owner names and full dates. |
| Explicit-fiction control | New slide titled `Training examples`, with exactly two populated fictional rows. Each fictional cell visibly carries an `Example:` label. |

Prompts/seeds were frozen before execution. Pre-live review strengthened assertions to enforce the already-declared row associations and populated examples, and corrected the fictional control's clarification so it retained the user's explicit example authorization. Both declarations and review notes are saved. No live output informed those changes. Frozen source hashes matched at completion. Original run failures and verdicts remain unchanged.

This fixes the identified instruction conflict and passes the focused counterexamples; it does not make semantic factuality a deterministic bridge guarantee or establish a general model success rate.

### Method gates did not protect newer properties

The bridge gated older-host read methods correctly but loaded/read the newer native `Slide.index` property inside them. The simulator previously returned that property regardless of the configured host version, so compatibility tests could not expose the defect.

`readAllSlides`, `readSlides`, `readSlide`, `readRange`, selected-slide reads and `readShapeContext` now derive displayed positions from the observed ID collection's order. Selection resolution checks 1.5 support; text-context resolution checks 1.4 support. Whole-deck reads retain their constant sync count. The native-table write still requires 1.8.

The test host now rejects `Slide.index` loads/access below 1.8. Public read-port tests pass at simulated API versions **1.4, 1.5, 1.7 and 1.8**, retaining the correct slide positions and source content. The old code failed these tests. Real Office-host certification remains separate.
