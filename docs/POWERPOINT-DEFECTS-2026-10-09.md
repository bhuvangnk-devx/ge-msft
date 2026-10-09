# PowerPoint: shapes, false "done", and designed slides on long decks

**Status:** fixed on branch `fix/powerpoint-command-defects`, not committed. All three defects
reported from CIMB Niaga testing on 2026-10-09 were reproduced live in PowerPoint for the web (test
deck `Atlas-Q3-Review-TextOnly.pptx`, 12 slides), fixed, and re-run on the final build.

**Bottom line: the three reported failures came from six defects in our code, not from the model.**
Each fix is in the layer that caused it, and reuses what was already there: the existing router,
command compiler, plan helpers, context brief and snapshot builder. The bridge's list of shapes
moved into `contracts` so the compiler can use the same list.

**Files changed:**

| File | What changed |
|---|---|
| `packages/web-shell/src/taskpane/components/App.tsx` | `/shape …`, `/slide …` and other unknown `/words` go to the planner, not chat |
| `packages/runtime/src/command-protocol.ts` | `/add-shape shapeType=rectangle` draws a rectangle; an unknown shape type is refused before approval; the prompt says where `finish when=verified` can complete |
| `packages/contracts/src/capability.ts` | the list of drawable shapes, moved here from the bridge so the compiler and bridge share it |
| `packages/bridge-powerpoint/src/powerpoint-bridge.ts` | uses the shared list; writes `text=` inside a shape; reordered "shape not found" message |
| `packages/runtime/src/context-model.ts`, `assist-session.ts` | new tasks no longer see earlier tasks' write notes; a change plan cannot end with `done` before any write; actionable fence error |
| `packages/bridge-powerpoint/src/capture.ts`, `packages/content/src/doc-state-builder.ts` | the deck snapshot lists each slide once, numbered |
| `packages/contracts/src/capability-registry.ts` (+ regenerated `skill/…/m365-cli-1.0.json`, catalog) | what the model is told about `/add-shape` and `/add-table-slide` |

---

## In short

| # | Reported defect | Prompt | Before | After |
|---|---|---|---|---|
| 3 | Triangle / rectangle not drawn | `/shape Add a blue rectangle to the top right of slide 2.` | Fails: went to chat, which answered that it cannot draw shapes | Passes: blue rectangle, 1 approval |
| 3 | (same, without `/shape`) | Add a blue rectangle to the top right of slide 2. | Fails: 7 approvals, 10 failed writes; a "Draft" text box was added and the slide title recoloured | Passes via the same route |
| 1 | "Done" but nothing added | `/slide Add a summary slide with the three key findings: …` | Passed in our run; the reported failure is explained below | Passes |
| 2 | Complex design fails | CIMB Niaga CoE slide: dark title bar, 3 bullets, 2-column table (Indonesian) | Fails: no table after 12 turns | See [Live results](#live-results-after) |

---

## For everyone: what was going on

The add-in lets Gemini edit the deck through a small set of commands. A planner turns the request
into steps, the user confirms, and an executor writes the commands. These six things went wrong.

### 1. `/shape …` and `/slide …` went to chat, which cannot edit slides (defect 3)

PowerPoint's own commands are `/ask`, `/summarize`, `/explain`, `/draft` and `/review`. Any other
text that started with `/` was sent to plain chat. Chat cannot write, so Gemini replied that drawing
shapes is not supported. Such text is now routed like a normal request, without the slash, so
`/delete slide 3` reaches the planner as "delete slide 3".

### 2. "Rectangle" became an empty text box (defect 3)

The model writes `shapeType=rectangle` or `shapeType=triangle`. The add-in only knew `textBox`,
`geometric` and `line`, and quietly treated anything else as a text box with no text, which the
bridge then refused. The model was only ever shown a text-box example, so it never learned the
right form. In the live run it tried 7 times, each behind an approval card, then gave up and
recoloured the slide title instead.

Now `shapeType=` takes a shape name in any case (`rectangle`, `Triangle`, `ellipse`, …), the
command card's first example is a filled rectangle, and text given for a shape is written inside it.

### 3. New tasks read notes about earlier tasks' writes (defect 1)

Planner and executor requests are sent without chat history on purpose, so each task starts clean.
But a running list of past write outcomes ("Applied insert-slide …", "Attempted add-shape …
no_text") was attached to every request and never cleared, up to 40 notes. A new "add a slide"
task could read that a slide had already been added. This is the likely cause of the reported
`done "The SIT Slide has been successfully imported and verified."`: nothing in the request
mentioned a SIT slide. **This is inferred, not verified**: the reporter's session was not
recorded, and the same prompt passed in our runs. Settled write outcomes are now left out of these
requests. An uncertain outcome (a write that may have landed) is kept.

### 4. "Done" with nothing changed counted as success (defect 1)

If the model ended a confirmed "add a slide" plan with `done` before writing anything, the task was
reported complete. Now that first `done` is refused once, with a note to emit the write. A second
`done` stands, so a change that is already in place does not loop. A write the user rejected
counts as tried.

### 5. On a long deck the last slides were invisible (defect 2)

The deck summary the model sees listed every bullet as its own entry and stops at 60 entries. On a
15-slide deck it ended around slide 11. Asked to add a table to the slide it had just created
(slide 15), the model could not see it and spent all 12 turns searching. The `slides` command did
not help: on PowerPoint it lists only the selected slide. The summary now lists each slide once, by
title and number, for up to 60 slides.

### 6. The model kept ending with a `finish` that PowerPoint can never satisfy (defect 2)

`finish when=verified` completes a task only when every write reports a verified readback, and only
Excel cell writes do. The prompt offered it on every surface with "use where host readback is
supported", which the model cannot check. In two PowerPoint runs it ended its commands with
`finish when=verified` and no closing fence. The add-in refuses that by design, but the error did not
say how to fix it. One run sent the identical block three times until the task stopped. Another
recovered only on turn 11 of 12.

Now the prompt says only Excel cell writes return readback, and to end with `done` otherwise. The
fence error says to close the block or end it with `done`.

---

## For developers

### Live evidence (before)

Every run was in Office for the web with every approval card approved. Each model request and
reply was recorded, and results were read back from the deck through Office.js.

- **Defect 3, `/shape …`:** chat answer "There is no direct tool … that can … draw specific shapes".
- **Defect 3, no prefix:** the first block was
  `/add-shape slide=2 shapeType=rectangle left=576 …` plus `/format-shape slide=2 shape=rectangle fill=#0078D4`.
  Results were `no_text` and "Shape "rectangle" was not found". It repeated with `name=myRect`,
  `shape=myRectangle` and `text=" "`. It ended with `/add-shape … shapeType=textBox text="Draft"`
  and `/format-shape … shape=title fill=#0000FF`. Result: "Task completed. 2 changes applied; 10
  not applied".
- **Defect 2:** the snapshot carried `truncated=true` and stopped at slide 11. `slides` returned
  only slide 1 on all 3 calls. The run ended "Stopped after 12 turns" with no table.
- **Cross-task notes:** all 3 requests of a new task carried the previous task's
  "Attempted add-shape … no_text" notes.

### The changes

#### Routing (`App.tsx`)

- `withoutUnknownVerb(inv)`: when the composer found no palette verb but the line starts with `/`
  and has text after it, drop only the slash.
- It runs in `dispatch` after the pasted-CLI check, so `/add-shape slide=2 …` pasted as a command
  still runs directly.
- A bare `/word` is unchanged.
- Everything then flows through the existing `inferImplicitIntent` and `shouldUsePlannerForFreeText`.

#### Shape commands (`command-protocol.ts`, `capability.ts`, `powerpoint-bridge.ts`)

- `addShapeKind(props)` replaces `shapeTypeFromProp`:
  - `textBox`, `line`, or any name in `POWERPOINT_GEOMETRIES`, matched in any case;
  - `shapeType=geometric geometryType=…` still works;
  - an unknown name is a compile error that lists the valid names, so it goes back to the model
    before any approval card.
- `POWERPOINT_GEOMETRIES` and `powerpointGeometry()` moved from the bridge's private
  `GEOMETRIC_SHAPES` into `contracts`. The bridge still validates against the same whitelist, which
  stays the host parameter type.
- The bridge writes `text=` into a geometric shape's text frame. It was dropped before.
- The "shape not found" message lists the slide's shape ids first and says a new shape's id is in
  its add-shape result. Before, it led with "for the slide title use title", which is the hint the
  model followed when it recoloured the title.

#### Context brief (`context-model.ts`, `assist-session.ts`)

- `ContextModel.note` flags settled write outcomes.
- `pendingBrief({ withoutActuations: true })` leaves those notes out.
- `streamTurn` uses it only for isolated (session-less) planner and executor requests. A task's own
  outcomes still reach the model as its result blocks, and the deck itself arrives as the fresh
  snapshot.
- Chat turns are unchanged.

#### False "done" (`assist-session.ts`)

- `changeAwaited` is set when the task is a confirmed plan (`CONFIRMED_PLAN_OPEN`) whose text
  describes a change. This reuses the existing `commandIntentText` and `planDescribesChange` from
  `contracts`; no new parser.
- The `done` handler refuses the first `done` that arrives with no recorded effect and no share,
  as an advisory note, so it does not count toward the repair limit.

#### Snapshot (`capture.ts`, `doc-state-builder.ts`)

- `slideElementsToDocStateBlocks` emits one title heading per slide, without body lines (`read slide:N`
  returns those).
- The builder numbers slide headings itself and renders the number outside the quoted title:
  `- [slide] ref="slide:270#…" "Strategi AI …" (slide 15)`. A slide titled "Slide 2: …" therefore
  cannot pass for slide 2 (a security review finding).

#### `finish` guidance (`command-protocol.ts`, `assist-session.ts`)

- The bootstrap's existing two `finish` lines are reworded, not added to:
  - "Only Excel cell writes return readback";
  - "For any other write, inspect the outcome and emit done alone when complete".

  No logic changed. The runtime rule for completing a verified program is untouched.
- The parse error "finish when=verified requires a closed cmd fence" gains ": end the block with a
  line of three backticks, or end it with done instead".
- The parser in `contracts` and its Python twin in the skill bundle are unchanged, so they stay in
  lockstep.

### Security review

The `security-reviewer` agent found no critical issues and one high:

- **High: an early version let `/format-shape shape=last` resolve after approval.** It picked
  whatever shape was on top, so a failed `add-shape` before it would have let the format land on
  an existing shape. Fixed by removing `shape=last`.
- **Lows, all fixed:**
  - the "done" check now counts a completed share as tried;
  - slide numbers sit outside the quoted title;
  - only the slash of an unknown `/word` is dropped, not the word.

A re-check of the fixes found nothing above informational.

### Tests

- `bun run typecheck` clean; `bun run test` 2,860 passed, 16 skipped; `bun run lint` clean;
  `bun run skills:check` current.
- `packages/runtime/src/replay/powerpoint-2026-10-09.replay.test.ts` replays:
  - the refused `done`;
  - the cross-task notes;
  - the unclosed `finish`.

  Its 4 defect tests fail with the runtime fixes removed; its 2 control tests pass either way.
- Unit tests:
  - `command-protocol.test.ts`: shape names and refusals;
  - `powerpoint-bridge.test.ts`: text in a shape, message order, a 15-slide deck lists slide 15;
  - `doc-state-builder.test.ts`: slide numbering outside the title;
  - `context-model.test.ts`: brief filtering;
  - `capability.test.ts`: shape-name lookup;
  - `command-surface.integration.test.ts`: `/shape`, `/slide`, `/delete` routing.

### Live results (after)

Final build, every approval card approved, each result read back from the deck:

| Prompt | Result |
|---|---|
| `/shape Add a blue rectangle to the top right of slide 2.` | Pass, 1 approval: `GeometricShape` `#0F6CBD` at 560, 20, 140 × 60 |
| `/slide Add a summary slide with the three key findings: …` | Pass: title "Summary" and the 3 bullets |
| CIMB Niaga CoE slide (Indonesian) | Pass in 8 turns, 3 approvals: dark-blue (`#002060`) title bar with white text, 3 bullets, native table; no `finish` |

In that CIMB run, three first attempts failed and were corrected on later turns. These are the open
issues below:

- the title fill hit a slide that had not settled;
- `title=` was passed to a table on an existing slide;
- the table was placed below the slide.

Earlier runs on this branch:

- "Add a red triangle to the bottom left of slide 2" passed (1/1).
- The CIMB prompt passed 3 of 4 times. The failing run is the one that found defect 6.
- One run before the `finish` rewording passed only on turn 11 of 12.

All test shapes and slides were removed from the deck afterwards.

---

## Open issues

- **A slide added moments ago reads as empty on PowerPoint for the web.** Measured: no shapes for
  over 6.7 s after the add, by id and by position. A format on it in the same block fails, and the
  model redoes it on the next turn, which costs a second approval.
- **`/add-table-slide` on an existing slide refuses `title=`.** The model passed it in every CIMB
  run and corrected it on the next turn.
- **First placements can fall off the slide.** The model assumes a 540 pt tall slide; this deck is
  405 pt. The bridge refuses with the real size and the model corrects it, at the cost of an
  approval.
  - Putting the slide size into the snapshot fixed this in testing (a `slideSize` field).
  - It was left out to keep this change small; it is the first follow-up worth taking.
- **Native tables grow past their height.** PowerPoint sizes rows to fit, so a table placed at
  `top=200 height=160` ended about 30 pt below the slide.
- **Defect 2's exact prompt is unknown.** It was rebuilt from the screenshot; re-run it with the
  original wording when available.

## Considered and left out

These were tried during the investigation and removed to keep the change minimal. Each is a small,
separate follow-up if the open issues above matter in testing.

| Idea | Why left out |
|---|---|
| Slide size in the snapshot | New contract field; the bridge's off-slide error already self-corrects |
| Everyday shape names (`circle`, `square`, `rounded-rectangle`) | An unknown name is refused before approval with the valid list, and the model corrects it |
| Refusing an empty text box before approval | The bridge already refuses it |
| A "slide not settled yet" message for an empty new slide | The model recovered without it |
| A registry flag marking which writes return readback, to tailor the `finish` text per surface | A one-line rewording of the existing prompt text did the same job; see change 6 |
| `/format-shape shape=last` | Removed after the security review (see above) |
