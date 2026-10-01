# CNGPT hardening plan (items 1–9)

Status: 1–5, 9 done · 6–8 planned. Owner: Bhuvan. Started 2026-10-01.

These items come from bugs found in manual testing on the CIMB test deployment (Cloud Run
`cngpt-web`), not from a full audit. Each item lists the problem, the evidence, the change, how it
is verified, and an effort estimate.

Every item ends with the same gate: `bun run typecheck`, `bun run test`, `bun run lint` clean, a
failing-first test for the bug, a Cloud Run redeploy, and a push to `cimb-cicd`. Items touching
guardrails or provenance (2, 7) also get a `security-reviewer` pass.

| # | Item | Effort | Order |
|---|---|---|---|
| 1 | Indonesian requests reach the command route | 0.5 day | now |
| 2 | Chat replies stop claiming changes they did not make | 2–3 h | 2nd |
| 3 | Read questions in the command route end with an answer | 0.5 day | 3rd |
| 4 | Skill bundle / web app version drift is visible | 2 h | 4th |
| 5 | "Copy diagnostics" for a failed run | 0.5 day | 5th |
| 9 | Request timeouts | 2–3 h | 6th |
| 8 | PowerPoint position gaps | 2 h | 7th |
| 7 | Hide expected provenance warnings | 1 h | 8th |
| 6 | Live smoke suite per app | 1–2 days | last (needs a test tenant) |

---

## 1. Indonesian requests reach the command route

**Problem.** The free-text router (`packages/web-shell/src/taskpane/components/App.tsx`:
`LEAD_IN`, `requestRe`, `OFFICE_ACTION_REQUEST_RE`, the per-surface `*_RE`) only knows English
verbs and lead-ins. CIMB staff writing "Tolong balas email ini" or "Buatkan grafik pendapatan"
land in chat, where nothing can be written.

**Change.**
- Add Indonesian lead-ins: `tolong`, `mohon`, `coba`, `bisa(kah)` / `dapatkah` / `bisakah` (+ `kamu`/`anda`),
  `saya mau` / `saya ingin` / `aku mau` / `saya perlu`, `ayo`, `silakan`.
- Add Indonesian action verbs with their common suffix forms (`-kan`, `-i`, `-lah`): `balas`, `tambah`,
  `buat`, `ubah`/`ganti`, `hapus`, `sisipkan`/`masukkan`, `tulis`, `perbaiki`, `format`, `urutkan`,
  `isi`, `tandai`, `sorot`, `pindahkan`, `terapkan`, `lampirkan`, `susun`, `edit`, `revisi`, `ringkas`
  is a READ (summarize) and stays out.
- Indonesian object nouns for the per-surface fast paths: `grafik`/`diagram`/`bagan` (chart),
  `slide`/`salindia` (slide), `balasan`/`email`/`surel`/`pesan` (Outlook draft), `komentar` (comment).
- Indonesian question words keep text in chat: `apa`, `bagaimana`, `kenapa`/`mengapa`, `siapa`,
  `kapan`, `di mana`, `berapa`, `jelaskan` (explain), `ringkas(kan)` (summarize).

**Verify.** Extend `routing-corpus.test.ts` with ~20 Indonesian questions (must stay in chat on every
surface) and ~20 Indonesian requests (must leave chat). Run it against the current router first to
record what fails today.

**Risk.** False positives move a question to the planner, whose "chat" verdict still answers it in
chat (one extra planner turn, no write). The corpus guards against this.

## 2. Chat replies stop claiming changes they did not make

**Problem.** The chat route has no write commands, but the model still answers "Here is a bar chart…"
or "I've changed the title". Seen in the `/visualize` scenario and the earlier "Change the title…" case.

**Change.**
- In `controller.ts` `send()`: when a finished chat answer claims a change ("I've created / added /
  inserted / updated / changed / replied", Indonesian equivalents) and no write ran, append a muted
  note: "No change was made to the document. To apply it, ask again as an action (e.g. /visualize)."
- Keep it a display note, never a rewrite of the model text.

**Verify.** Controller tests: claim + no write → note; plain answer → no note; command-route answers
unaffected.

## 3. Read questions in the command route end with an answer

**Problem.** "Read the last row for me" in Excel: the read succeeds, the model emits `done`, the
panel shows only "Done". Cause (diagnosed earlier): `previewOf` sends only the first row
(`excel-bridge.ts:1084`), `done` carries no message (`command-grammar.ts:224`), `SKILL.md` says
"emit `done` alone", and `controller.ts:1355` shows only a step.

**Change.**
- Grammar: `done "answer text"` (optional quoted answer), mirrored in `parse_commands.py`.
- Controller: show the `done` answer as the assistant message.
- Skill: "after a read-only request, finish with `done "<the answer>"`".
- Excel preview: include the last rows as well as the first (bounded).

**Verify.** Grammar + parity tests; a session test where read → `done "…"` shows the answer.

## 4. Skill bundle / web app version drift is visible

**Problem.** The web app and the Gemini Enterprise skill bundle ship separately (`bun run ge:skills`).
On 2026-10-01 the web app learned `slide … at=N` while the model's skill copy did not; nothing warned.

**Change.**
- Emit a grammar fingerprint (hash of `m365-cli-1.0.json`) into both the web build and the skill bundle.
- When the model's turn reports a different fingerprint (or the bundle metadata does), add a step:
  "Model instructions are older than this app — run `bun run ge:skills`".
- Add `ge:skills` to the release checklist in `RUN.md` and `scripts/cimb-release.sh`.

**Verify.** Unit test on the comparison; manual: deploy web without skills → warning shows.

## 5. "Copy diagnostics" for a failed run

**Problem.** No trace of a user's failed run reaches us; diagnosis is guesswork.

**Change.**
- A "Copy diagnostics" action on each run's Activity: route chosen (chat / planner / command),
  planner verdict, steps, command lines (no document text), error codes, app + host version, build id.
- Source: the execution ledger, which already excludes prompts and content.

**Verify.** Test that the export contains no document text (redaction test with a known secret string).

## 6. Live smoke suite per app

**Problem.** Every bug found on 2026-10-01 came from manual testing; unit tests use fake hosts.

**Change.** 10 scripted prompts per app (Word, Excel, PowerPoint, Outlook) run against a test tenant
before each release, building on the existing `test:streamassist:modes` live harness. Record pass/fail
in a release note.

**Needs.** A test tenant account and documents that can be reset. Blocked until those exist.

## 7. Hide expected provenance warnings

**Problem.** "⚠ provenance not recorded" appears on every Outlook/PowerPoint write. It is expected
(those apps cannot store provenance) but reads as an error.

**Change.** Show the warning only where the bridge advertises durable provenance (Word/Excel); keep it
in diagnostics (item 5) everywhere. Security review required (provenance visibility).

## 8. PowerPoint position gaps

**Problem.** `/add-table-slide slide=new` still appends; after `slide … at=N`, `slide=last` points at
the old last slide.

**Change.** `at=N` for `add-table-slide slide=new`; skill note already says use `slide=N` after `at=`.
Bridge reuses the `moveTo` path from `insert-slide`.

## 9. Request timeouts

**Problem.** No request timeout found in `packages/gemini-client` (only retry back-off). A stalled
planner or chat stream spins until the user cancels.

**Change.** An idle timeout on the stream (no event for ~60 s → abort with "The assistant did not
respond. Try again."), applied in `gemini-client` so chat, planner and command routes share it.
Must not abort a write already under way.

**Verify.** Fake-timer tests: idle stream aborts with the message; a slow-but-alive stream does not.
