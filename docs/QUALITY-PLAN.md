# CN-GPT Quality Plan

**Goal:** we find problems before users do. A user's first request in Word, Excel, PowerPoint or
Outlook works, says what it did, and never claims a change it did not make.

**Status (12 Oct 2026):** 269 of 280 real-user test requests pass (96%), up from 208 (74%) on
9 Oct. Median: 17 seconds and 2 Gemini calls per request.

---

## 1. The bugs we had, and why

Users and the client found most of these bugs before we did. They fall into six groups.

| Group | Examples | How we found out |
| --- | --- | --- |
| **1. The model had to guess** | Questions turned into edits ("Extract the action items" rewrote an email; "List the obligations" made a tracked change; an analysis request ended with no answer). Outlook mixed up reply and new email, and opened 4 drafts. PowerPoint could not reach a subtitle. Gemini made up figures when the sheet could not be read. | Users, then the test matrix |
| **2. Our own UI filters and wording** | Commands leaked into chat four times (cmd block, narrated line, wrapped bullets, `let` line). Blank bubble while streaming. Model Armor message overwritten. Copy answer failed in Office desktop. A guidance note cut off at 240 characters. | Users |
| **3. Office host behaviour** | PowerPoint for the web "GeneralException" and unconfirmed slides, then duplicate retries. Excel response size limit on whole-column selections. Clipboard blocked in Office desktop. | Users |
| **4. Hosting and network** | Untrusted certificate on `.local`. Chrome/Edge Local Network Access block. Chrome Secure DNS. VPN differences. | Client IT |
| **5. Missing features** | New chat needed a pane refresh. History did not show old messages. Diagnostics were lost when the pane closed. Chats from other apps were not labelled. | Users |
| **6. Test tooling** | Simulated Word, Excel and Outlook lacked APIs that real Office has, so whole capabilities went untested. | The test matrix |

### Root causes

1. **We tested our code, not the experience.** About 3,200 unit tests check our code against
   scripted model replies. Until 9 Oct nothing checked what real Gemini does with real user
   wording.
2. **There was no gate before deploy.** Changes went to the client's Cloud Run minutes after they
   were written, checked only by unit tests. Several bugs came from these quick fixes.
3. **The model decides with too little information.** Most behaviour bugs share one cause: the add-in
   did not tell Gemini something it needed (is this a question, is the email a draft, which shapes
   are on the slide, could the document be read). Each fix so far adds the missing fact.
4. **We fix symptoms with keyword rules.** "format / highlight / flag means a change" is a regular
   expression. Each patch fixes one wording and can break another. The security review caught this
   twice in one week.
5. **The simulators are not real Office.** Nothing runs in real Office before a release.
6. **We cannot see production.** We learn about failures from screenshots. We do not know the failure
   rate per capability.

---

## 2. What is in place now

### Quality gate (`bun run quality:check`)

Run it on a clean commit before every deploy. It takes about 40 minutes.

1. Static checks: type check, lint, unit tests (one retry for the known timing flake), brand guard.
2. Live checks: every scenario in `scenarios/*.json` (338 today) against real Gemini Enterprise,
   with simulated Office, one process per app.
3. Each capability is compared with `scenarios/quality-baseline.json`. The gate fails when a
   capability loses more than one pass against its baseline, when the overall rate is below 95%, or
   when a scenario produces no result.
4. The report goes to `scenarios/.results/quality-check.json`, with the commit, the verdict and a
   table of pass rate, median time and median Gemini calls per capability.

`bun run deploy:staging` refuses to deploy unless that report is a pass for the exact commit being
deployed and the working tree is clean.

When a change is meant to improve results, accept the new numbers with
`bun run quality:check --update-baseline` and commit the baseline file with the change.

### Guards against false claims

- If a reply says it changed the document but no write landed, the answer shows: "No change was made
  to your document in this request, whatever the reply above says." This applies to chat and command
  answers, in English and Bahasa Indonesia.
- A planner-confirmed request that changed nothing must answer. A bare `done` is refused once and
  asked for the answer or the reason nothing changed.
- If the document cannot be read, Gemini is told so and told not to guess its contents.

### The test matrix

`scenarios/matrix-*.json`: 10 real-user requests per capability, 28 capabilities, in English and
Bahasa Indonesia. Run one app with `GE_SCENARIO_SURFACE=excel bun run scenarios`.

---

## 3. What to do next

| # | Action | Effort | Prevents |
| --- | --- | --- | --- |
| 1 | **Internal environment.** The Cloud Run service the client uses is our only environment. Add a second service for our own testing, and deploy to the client's service only through `deploy:staging` after the gate passes. | ½ day | Users seeing untested builds |
| 2 | **Real-user test set.** Turn real user questions (from the Gemini sessions API, anonymised) into scenarios. Add multi-turn follow-ups, large documents and more Bahasa. Check facts in answers, not only what was written. | 1–2 days | Wording we never tried; made-up answers |
| 3 | **An explicit effect from the planner.** The planner states `effect: answer` or `effect: change`, replacing the keyword rules. | 1 day | The whole "question became an edit" group |
| 4 | **Simulator fidelity check.** A test fails when a bridge calls an Office API the simulator does not implement. | ½ day | Untested capabilities |
| 5 | **Real Office smoke test.** The top 20 scenarios in Office for the web on a test tenant before each client release. Manual checklist first. | 2 h per release | Host bugs like PowerPoint GeneralException |
| 6 | **Content-free outcome telemetry.** Route, status and error codes (the same fields as Copy diagnostics) to Cloud Logging, with a failure-rate dashboard. Needs a security review and client approval. | 2–3 days | Learning about failures from screenshots |

---

## 4. Gaps from the user's point of view

These are not bugs. They are places where the product is harder to use than it should be.

### Must fix before wider rollout

1. **The interface is in English only.** CIMB Niaga users work in Bahasa Indonesia. Answers follow
   the user's language, but buttons, approval cards, errors and Activity are in English. Only the
   disclaimer is localised.
2. **Error messages do not tell the user what to do.** Examples users can see today: "The task did
   not complete all its operations. Review the execution results before continuing", "outcome
   unknown", "(PowerPoint: GeneralException)". Each error needs a plain cause and a next step.
3. **There is no way to report a problem from the pane.** Users report through a spreadsheet or
   WhatsApp. Add "Report a problem" that sends the question, the diagnostics and an optional note,
   and a quick helpful / not helpful on each answer.
4. **Users do not know what they can ask.** There is no first-run guide per app. Add three to five
   example requests per app, in Bahasa and English, on the empty pane.

### Should fix

5. **Waiting time is not explained.** Planning takes 15–20 seconds before anything visible happens,
   and a long analysis can take 60–90 seconds. Show what is happening in plain words, and set
   expectations for long answers.
6. **A change needs two confirmations.** The user confirms the plan, then approves the changes. For
   low-risk changes (one cell, one format), one confirmation may be enough. This needs a security
   decision.
7. **Internal terms are visible.** The approval card and Activity show terms such as `write-cells`,
   `insert-slide`, `turn start` and shape ids. Users need "Change cell B2 to 100", not the command.
8. **Undo is unclear.** After a change there is no one-click undo except in the Excel analysis
   workbench. Users must know to use Office undo or reject tracked changes. Show how to undo after
   each change.
9. **Attached context is confusing.** "Attached context · 0 sources · + 3 available" does not say
   that the current selection is always used. Say "Using your selection" when that is the case.

### Later

10. Accessibility (keyboard-only use, screen readers) and Office mobile have not been tested.
11. History has no search, and chats from CN-GPT web show no source until that app labels its
    sessions.

---

## 5. Rules from now on

1. Deploy to the client's Cloud Run only with `bun run deploy:staging` after `bun run quality:check`
   passes.
2. Every reported bug gets a scenario that reproduces it before it is fixed.
3. A fix that tells the model something must fit the 240-character snapshot limit and be checked
   live, not only in unit tests.
4. Changes to routing, guardrails or what is stored on the device get a security review before
   commit.
