# Bug: Office refuses a change (protected sheet) → pane shows a generic error and loops

**Found:** 2026-10-14, staging `cngpt-web-00052-9f5`, Excel on the web, while testing error reports.
**Status:** root cause investigation started, **not confirmed**, nothing fixed.
**Owner:** next session. Use `superpowers:systematic-debugging` (finish Phase 1 before any fix).

---

## 1. What the user saw

1. Sheet2 was protected (Review → Protect Sheet → On).
2. Request: `Put the header "Region" in A1 and make it bold` → **Go ahead** → **Apply**.
3. The pane ran **12 model turns** (~75 s), then showed **"That didn't finish" + "Try again"**.
4. It never said **why**: the sheet is protected.

**Expected:** stop at the first refusal and say, in the user's language, something like
"Sheet2 is protected, so I could not change A1. Turn off protection (Review → Protect Sheet) and try again."
No retry loop, no generic error, no "Try again" for something that will fail again.

## 2. Evidence (real report `c2d9251611c1408ba5694408412840a6`)

From the stored report / the user's Copy diagnostics (same content):

| Time | Where | Code | Message |
|---|---|---|---|
| 06:57:44.677 | write `write-cells` | `outcome_unknown` | Host execution was interrupted. Inspect the cells before deciding what to do. |
| 06:57:44.679 | write `format-cells` | `prerequisite_failed` | skipped […] |
| 06:57:52.420 | write `format-cells` | `outcome_unknown` | Host execution was interrupted. Inspect the document before retrying. |
| 06:58:03.618 | write `write-cells` | `prepare_failed` | An unresolved write overlaps Sheet2 […] |
| 06:57:49 → 06:58:35 | command ×8 | – | outcome unknown […] |
| 06:58:46.553 | panel | – | The task did not complete all its operations… Some changes may already have applied. |

Run: `command`, status **`blocked`**, 12 turns, 24 tool calls, 4 failed effects
(`outcome_unknown`, `prerequisite_failed`, `outcome_unknown`, `prepare_failed`). Steps end with
`exhausted: Stopped after 12 turns` → `Task blocked. 0 changes applied; 4 not applied.`

**Note:** Office's own error (the "protected" refusal) appears **nowhere** in the diagnostics.

## 3. Code path found so far

The message comes from two places; both swallow the real host error:

- `packages/runtime/src/recovery.ts:255-262` — `catch {` around `actuate()`: any throw → record
  `state = 'uncertain'`, message "Host execution was interrupted. Inspect the cells…". The error
  object is discarded.
- `packages/runtime/src/assist-session.ts:3531-3536` — `catch (err)` after dispatch →
  `unknownActuationResult(…, 'Host execution was interrupted. Inspect the document before retrying.')`.
  `err` is discarded.
- `packages/contracts/src/actuation-outcome.ts` — `assessActuationResult` treats `outcome_unknown` as
  `uncertain`; `unknownActuationResult` builds it.

**Working hypothesis (unverified):** the Excel bridge's `actuate` **throws** Office's error
(`OfficeExtension.Error`, likely code `AccessDenied` / `InvalidOperation` for a protected sheet)
instead of returning `{ ok: false, error: { code } }`. The runtime cannot tell "Office refused
before writing" from "Office failed mid-write", so it says *uncertain*. The uncertain record then
blocks later writes to the same sheet (`prepare_failed: An unresolved write overlaps Sheet2`), and the
model keeps retrying until the 12-turn cap.

## 4. Phase 1 — confirm before fixing

1. **Get the real Office error.** In Excel on the web with a protected sheet, run in the pane console:
   ```js
   await Excel.run(async (c) => { c.workbook.worksheets.getItem('Sheet2').getRange('A1').values = [['x']]; await c.sync(); })
     .catch((e) => ({ code: e.code, message: e.message, debugInfo: e.debugInfo }))
   ```
   Record `code` and `debugInfo` (which statement, and whether anything was written).
2. **Find where the bridge throws.** `packages/bridge-excel/src/excel-bridge.ts` `actuate` for
   `write-cells` / `format-cells`: does it catch Office errors and return a structured result, or rethrow?
3. **Is it really a non-write?** Confirm the cells are unchanged after the refusal (it should be a
   clean refusal; Office validates protection before applying).
4. **Check the same path in Word, PowerPoint and Outlook bridges** (`bridge-*/src/*-bridge.ts`).
5. Write the failing test first (bridge test with a fake Office that throws `AccessDenied`).

## 5. The class of bug (not just protected sheets)

Any time **Office refuses** an action for a reason the user can fix, the pane must say the reason,
not a generic error. Check each of these:

| App | Refusal | What the user should hear |
|---|---|---|
| Excel | Protected sheet / locked cells | Sheet is protected; unprotect it |
| Excel | Protected workbook structure (add/rename/delete sheet) | Workbook structure is protected |
| Excel | Cell in edit mode (user typing) | Finish editing the cell, then try again |
| Excel | Write overlaps a table / merged cells / spill range / PivotTable | Name the obstacle |
| Excel | Filtered range, array formula part | Name the obstacle |
| Word | Restrict Editing / read-only document / content control locked | Document is protected / read-only |
| Word | Tracked changes required by policy | Explain |
| PowerPoint | Read-only presentation, locked shape (if API exposes) | Presentation is read-only |
| Outlook | Read mode instead of compose; item is sent / not editable | Open a draft or reply first |
| All | API requirement set not supported in this Office version | This Office version cannot do that; use X |
| All | Document opened read-only / co-authoring lock / file checked out | Explain |

## 6. Required behaviour (acceptance criteria)

1. **Classify host refusals as a known non-write**, distinct from `outcome_unknown`: a new error
   code family (e.g. `host_refused` with reason `protected` / `read_only` / `edit_mode` /
   `unsupported_api` / `overlap`), kept as closed codes so diagnostics stay content-free.
2. **Only true mid-write failures stay `outcome_unknown`.** A refusal must not create an uncertain
   recovery record and must not block later writes (`prepare_failed`).
3. **Stop the loop.** A refusal the model cannot fix (protection, read-only) ends the task at the first
   occurrence; the model gets one clear result, not 12 turns.
4. **Plain message in the user's language** (EN and Bahasa, `ui-language.ts`), naming the sheet/item and
   the fix. No "Try again" button when retrying cannot work; offer it only after the user can act.
5. **Prefer a pre-check where cheap:** e.g. `worksheet.protection.protected` before a write, so the
   pane can refuse before dispatch (certain non-write).
6. **Error reports:** these are user-fixable, not bugs. Decide whether to report them (suggest: no
   report, or `kind: error` with a closed `reason` so support can filter). Update `classifyFailure` in
   `packages/web-shell/src/reporter.ts` and the `kind` table in `docs/CNGPT-Reports-API.md`.
7. **Tests:** a bridge test per refusal type with a fake Office error; a runtime test that a refusal is
   `rejected`/`failed` (not `uncertain`) and does not block the next write; a controller/UI test for
   the message and no retry loop; replay the real case above.
8. **Live check:** protected sheet in Excel on the web → one turn, clear message, no "Try again".

## 7. Context for the next session

- Branch `cimb-cicd`. Staging = our Cloud Run `cngpt-web` (prod is CIMB-only). Only commit, push or
  deploy when the user says so.
- Uncommitted work in the tree: the error Reporter (`packages/web-shell/src/reporter.ts`, `host.ts`,
  `App.tsx`, `config.ts`, `main.tsx`, `auth-client.ts`, `controller.ts` `redactCredentials`,
  `gemini-client` `auth_exchange`) and an unrelated brand `promptName` change. Don't mix them into this fix's commit.
- Two confirmations stay (plan card, then approval card). Never show raw errors in chat.
- Run only the affected app's live scenarios, not the full quality gate.
- Error reports are on in staging: each failed test request is reported to the test API
  (`cngpt-reports-test`), which helps verify the fix (`kind`, codes, steps).
