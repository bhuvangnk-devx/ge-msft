# Clear failure messages: say what happened, stop the loop

**Date:** 2026-10-14 · **Branch:** `cimb-cicd` · **Status (2026-10-14):** Excel slice built in worktree branch `fix/clear-failure-messages` (uncommitted); see "Built" below.
Supersedes the open questions in `BUG-HOST-REFUSAL-SHOWN-AS-ERROR.md` (same bug, wider scope).

## Requirement (user, 2026-10-14)

Every failed request ends with one plain sentence that says **which kind of failure** it was. A user
must never think "the plugin is broken" when the document blocked the action or the plugin does not
offer that action yet.

| Class | Example | User sees | Retry? | Report? |
|---|---|---|---|---|
| **Office refused** | locked cell / protected sheet, read-only file, Outlook item in read mode | "A1 on Sheet2 is locked (sheet is protected). Turn off protection: Review → Unprotect Sheet. Then ask again." | No | No |
| **Not supported yet** | pivot table, sort, filter, add sheet | "CN-GPT cannot create pivot tables yet." + what it can do instead | No | No |
| **Real error** | Gemini timeout, sign-in, our bug | current "That didn't finish" + Try again | Yes | Yes |

Office refused and Not supported: stop at the **first** occurrence, EN/Bahasa, no Try again.

## What happens today (evidence)

### A. Locked cell (Office refused)

1. Excel handlers do not catch Office errors; the shared dispatcher has no try/catch
   (`runtime/src/bridge-dispatch.ts:58-104`). All 7 Excel handlers throw.
2. Runtime swallows the error code: `runtime/src/recovery.ts:255-273` (write-cells) and
   `runtime/src/assist-session.ts:3531-3536` (all other kinds) → `outcome_unknown` ("uncertain").
3. The uncertain record blocks every later write to the sheet (`prepare_failed: An unresolved write overlaps`).
4. Actuation failures never count toward the repair cap (`failed` counts only dry-run errors,
   `assist-session.ts:1750, 1903`), so the model retries until turn 12 (+1 no-fence turn).
5. ~7 Activity steps per turn × 12 turns ≈ the 91 steps the user saw.
6. End: `incomplete` → `core.outcomes` guard throws → generic `ErrorCard` "That didn't finish" + Try again
   (`web-shell/.../MessageThread.tsx:1198-1225`). No error code maps to a specific message.
7. No pre-checks anywhere for protection / read-only (grep: nothing).

### B. Pivot table (Not supported yet)

1. Model's per-turn command list (`command-protocol.ts:1123`) correctly omits `/insert-pivot`.
2. But the planner prompt hard-codes "chart/pivot/sort/filter/worksheet" (`contracts/src/command-plan.ts:527`),
   and skill docs list `/insert-pivot` as usable (`excel-semantics.md:56`, `generated-capability-catalog.md:27`,
   `generated-command-catalog.md:53` — ~70 slash kinds, some not in the registry at all).
3. Plan steps are free text, never checked against capabilities → user approves a pivot plan.
4. `/insert-pivot` parses, fails type check "not supported on this surface"; repair instruction says
   "Fix each failing line" → model retries → `repair_exhausted` after 3 failing turns → same generic card.

### Same path, other cases (from code, not live-tested)

| App | Case | Today |
|---|---|---|
| Excel | protected sheet (all 7 writes), cell edit mode, missing sheet, merged/spill range, table overlap | throw → unknown → loop |
| Word | protected / read-only document, locked content control, bad OOXML | 11/12 handlers throw; add-comment errors mislabelled `anchor_drift` |
| Outlook | reply while in compose, size/recipient limits | throw or `outcome_unknown` |
| PowerPoint | error after mutation (non-text shape) | `outcome_unknown` (pre-mutation errors are already clean) |
| OneNote | read-only / locked section | throw |

Not shipped (listed to the model somewhere, no bridge handler):

| App | Kinds |
|---|---|
| Excel | insert-pivot, sort-range, filter-range, format-chart, manage-worksheet |
| PowerPoint | insert-image, apply-slide-layout |
| Outlook | add-categories, compose-appointment |
| OneNote | set-page-title, add-outline, append-rich-text, add-note-tag, create-section |
| Teams | post-card + estate-gated kinds (Teams out of scope for CIMB) |
| Release profile `internal-alpha-word-excel` | Excel table/chart/conditional disabled, but skill lists them as core verbs |

## Fix plan (one standard, all apps)

| # | Change | Where | Est. |
|---|---|---|---|
| 1 | Closed codes: `host_refused` + reason (`protected`, `read_only`, `edit_mode`, `conflict`, `not_found`, `unsupported_api`), and `not_supported` (capability). Helper `needsUserAction(code)`. | `contracts` | 1 h |
| 2 | Catch host throws once in the shared dispatcher; map Office `error.code` (`AccessDenied`, `InvalidOperationInCellEditMode`, `MergedRangeConflict`, `ItemNotFound`, `ApiNotFound`, …) to a reason. Not-applied only when certain (pre-check, or before-snapshot unchanged); else stays uncertain but carries the reason. A refusal creates no blocking recovery record. | `runtime/bridge-dispatch.ts`, `recovery.ts`, `assist-session.ts` | 3 h |
| 3 | Stop the loop on the first `host_refused` / `not_supported`: new loop end `needs-user`, no more model turns. Also stop on the same actuation failure twice. | `assist-session.ts` | 2 h |
| 4 | Pre-checks: Excel `worksheet.protection.protected` + `range.format.protection.locked` before write; Outlook compose mode for reply; Word document protection where the API allows. | bridges | 2 h Excel, +1 h each |
| 5 | Prevent up front: planner vocabulary from the live manifest (`command-plan.ts:527`); skill rule "if the request needs a command not in Available commands, say so with `done` and stop — do not approximate"; regenerate skill catalogs from the manifest + release profile. | `contracts`, `skill/` | 3 h |
| 6 | UI: a notice card (not the error card) with the reason text in EN/Bahasa, naming the sheet/range and the fix; no Try again. Reporter: no report for these classes. | `web-shell` (`ui-language.ts`, `MessageThread.tsx`, `reporter.ts`) | 3 h |
| 7 | Tests (offline, no live Gemini): scripted-client stack tests in `web-shell/src/test-harness` for locked cell and `/insert-pivot` → one turn, notice card, no Try again, no report; bridge tests with fake Office errors. | tests | 3 h |

Order: Excel (locked cell + pivot) end to end first (~2 days), then Word, Outlook, PowerPoint, OneNote (~½ day each).
Live check afterwards: protected sheet and "make a pivot table" in Excel on the web; only Excel live scenarios.

## Open decision

For **Not supported yet**, should the pane also offer an alternative the plugin can do (e.g. "I can build
a summary table with formulas instead — want that?")? Recommended: yes, as a suggestion only; it acts
only after the user asks.

## Built (2026-10-14, branch `fix/clear-failure-messages`)

- `contracts`: `host_refused` reasons + `TaskBlocker`.
- `runtime`: loop event `needs-user`, run status `needs-user`; stops at the first unsupported write or
  host refusal. Status is `needs-user` only when nothing landed; otherwise stays `incomplete`.
- Shared dispatcher maps only `InvalidOperationInCellEditMode` (whole request rejected) to a refusal.
  Security review: statement-level codes (AccessDenied, *Conflict, ItemNotFound, ApiNotFound) can follow
  a landed statement in a multi-write batch, so they stay `outcome_unknown`; bridges pre-check instead.
- Excel: protected-sheet pre-check before write-cells (locked cells) and format-cells.
- Pane: one plain sentence (EN/Bahasa) appended to the reply, no error card, no Try again, no report.
- Tests: dispatcher, runtime replay (3), Excel bridge (2), full-stack pane (4).

Not done (follow-ups): planner vocabulary / skill docs still mention pivot etc. (fix 5); pre-checks for
Word protection, Outlook read mode, other Excel writes (table/chart/conditional/comment); merged/spill.

## Phase 2 decision (user, 2026-10-15): honest stop, all apps

Approved fixes: truthful capabilities (planner, skill, help) generated from the live manifest; one rule
"if it is not possible, say so plainly, never approximate"; a progress guard (3 turns with no change and
no new information → plain stop); chart ranges accept quoted sheet names; bulk edits as one range write.

Constraints (hold for the whole project):
- **Easy to change later:** no hand-written lists of unsupported features. A capability that ships later
  must appear everywhere by being added to its bridge manifest / registry status only.
- **Word edits are always tracked changes** (one setting, not per-command checks).
- **No overfitting:** general rules only; nothing keyed to pivot / flag / animation etc.

## Client bugs not yet sorted (UAT/SIT sheet "list defect UAT CN-GPT (2).xlsx", re-checked 2026-10-15)

Strict re-check scenarios: `scenarios/uat-sit.json` (run in every quality gate). Feature requests
(Plug IN #9 feedback buttons) are out of scope here.

- [ ] **UAT Plug IN #7: new formula column not filled completely.** English run: all 20 rows correct but
      the error card shown (cell-by-cell writes hit the 8-writes cap); an earlier run filled only the
      first row and claimed the whole column. Fix: to-dos **#8** (fill a formula down a range) + **#7**
      (raise the cap; capped writes covered by a later block write are not "missing"). Scenarios:
      `uat-xl-7-formula-column`, `uat-xl-7-formula-column-id`, `mx-xl-formula-15`.
- [ ] **UAT Plug IN #2: complex PPT designs.** A 4-shape design now works; designs needing more than 8
      shape changes still stop at the per-turn cap. Fix: to-do **#7**. Scenario: `uat-pp-2-complex-design`
      (add a larger variant, e.g. 12 boxes).
- [ ] **UAT Plug IN #6: inconsistent output.** 17 queries still fail on the newest code (to-dos #1–#6):
      multi-sheet workbooks, invented commands ending on the error card, looping instead of asking,
      wrong action for the request, uncertain outcomes, answers missing facts.
- [ ] **SIT #5–#7:** rows have no description; ask CIMB (Meli) for steps and the expected result.

## Open to-dos (from the 2026-10-15 re-run: 45/62 former failures pass, 17 remain)

All general; none keyed to one feature. Not started.

- [ ] **1. Excel: list every worksheet** (name + used range) as context, not only the active sheet.
      Query: "Fill in the Summary sheet … from the Jakarta and Surabaya sheets" (mx-xl-block-19). ~1–2 h
- [ ] **2a. Accept plural/near kinds** in `list` (`comments` → `comment`). Word reply-19 lost a turn. ~30 min
- [ ] **2b. Correction budget spent, nothing written → plain "couldn't do this" message**, not the error
      card. Queries: PPT "Move the Outlook slide…", "Pindahkan slide ini ke paling akhir" (model used `mv`). ~1 h
- [ ] **3. Prompt rule: missing information → ask** (`done "<question>"`), never loop. Queries: Word
      "Jawab komentar yang ada lalu tandai selesai", "answer the revenue question". ~30 min
- [ ] **4. Read-only questions cannot write** (plan intent ask/summarize → block writes), keeping mixed
      requests ("which is lowest? leave a note") working. Queries: "Rank the top 3 regions", "can you
      chart this?", "which one is the lowest? leave a note…", Word "Reply to the comment: change it to
      30 days", PPT "Add one slide summarising…". ~1 h
- [ ] **5. Uncertain outcomes**: check in real Office, fix add-in or fake host. Queries: Excel "table pls",
      "Wrap the text … autofit", Word "Insert a highlighted box", PPT "add a Draft stamp". ~2–3 h
- [ ] **6. Answer questions from the whole document** (bounded), not only the snapshot. Queries: Word
      "whats the DSCR?", "who's doing wat after this meeting?". ~2 h
- [ ] **7. Raise the per-turn write cap for in-document changes** (8 → ~30; each change still on the
      approval card). UAT Plug IN #2 (complex PPT designs). ~1 h + security review
- [ ] **8. Fill a formula down a range** (`set H2:H31 "=…"` fills every row, shifting relative
      references, like Excel's fill). Today it writes only the first cell, and the model has then
      claimed the whole column was done (mx-xl-formula-15). With #7, also stop counting capped writes as
      missing when a later block write covered the same cells (uat-xl-7: all rows correct, error card
      shown). **UAT Plug IN #7.** ~2–3 h
- [ ] Follow-ups from security review: report `no_progress` stops to support; tracked-edit notice on
      Word < 1.6; Word undo must reject tracked changes when an inverse executor exists.
- [ ] Skill docs on the engine still list unshipped features (regenerate + `ge:skills` deploy, with OK).

## Testing within the Gemini quota (2026-10-15)

Pool ≈ 2,400 assistant queries/day (160 per Standard licence, shared with CIMB staging, resets
14:00 WIB); a scenario ≈ 3.2 queries, so all 707 ≈ a whole day. `bun run quality:check` now runs a
daily sample (QUALITY_SAMPLE=4 per capability → 165 scenarios ≈ 530 queries, rotating; all covered
in ~5 days), stops at the first quota 429 (the rest = not run, not failed) and prints queries used.
`--full` runs all 707. Keep total processes ≤ 4 (30k tokens/min limit).
