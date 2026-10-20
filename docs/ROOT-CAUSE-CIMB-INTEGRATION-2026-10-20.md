# Root-cause integration into local cimb-cicd — 20 October 2026

Target confirmed by the user: local `cimb-cicd`, starting at `39fb0b3d385f60dd3bcf6117f4fcc31170cd4150`. `fix/root-cause` descends from that commit through seven existing RCA commits, followed by the additional fixes documented in [the quality report](PLUGIN-QUALITY-GATE-RESULTS-2026-10-20.md). Integration is local; pushing and deploying are separate actions.

## Local issue coverage

All substantive behavior and regression coverage in the target checkout's 13 modified code/test/skill/parser files is present in the root-cause working tree:

| Target's local change | Root-cause disposition |
| --- | --- |
| Task phrasing and permitted engine `suggested_questions` after planner/executor blocks | Present in planner, compact bootstrap and skills |
| Strip a closed trailing follow-up block without accepting unclosed/nested/extra programs | Same TypeScript helper, six grammar edge tests and Python mirror/self-tests |
| Strip follow-ups before execution and no-progress comparison | Present; root also strips before stale-line detection |
| Complete a valid program followed by engine questions | Same runtime acceptance test |
| Stop repeated programs despite changing questions | Same no-progress replay test |
| Refresh gcloud's potentially old cached access token every five minutes | Present; root additionally records streaming responses without buffering them before the client sees headers |
| Bootstrap headroom increased to 5,760 bytes | Superseded by 5,888 bytes, including the later border-format parameters; bounded disclosure tests retained |

Independent source review found no missing substantive local behavior or regression coverage. A three-way comparison found ten files fully covered without textual conflict and three overlapping hunks:

- `command-plan.ts`: preserve the target's task phrasing and the root-cause factual-table rule.
- `assist-session.ts`: preserve the target's follow-up stripper import and the root-cause command extractor import.
- `command-protocol.ts`: preserve the headroom rationale and the newer 5,888-byte bound.

These hunks use the root-cause version because it retains both intents. They are overlapping edits, not competing product behavior.

## Additional instruction gap found during review

Both versions still prohibited engine follow-ups in full disclosure and the no-fence repair prompt. Those paths are now aligned with compact disclosure. Full grammar uses task wording, permits one closed command block plus a trailing engine follow-up, and continues to reject executable Python/JSON/bash/unlabelled blocks. The repair text retains the same boundary and stays small enough to remain inline in execution state.

New public runtime regressions exercise compact and full disclosure through a repair, then an approved `A1` write followed by command-shaped text in `suggested_questions`. They verify exactly one approved write and no execution of the follow-up text. Both tests failed before the instruction correction. Existing hosted-code repair tests also caught an inline-budget spill; the repair wording was shortened without changing the refusal or raising the budget.

This establishes instruction consistency and local parsing/approval behavior. It does not prove that all Gemini-side stalls or Model Armor refusals have disappeared. The target's historical flakiness log and the root-cause quality reports describe different evidence: server-side stalls, prompt conflicts, and a separate client-side transcript-buffering defect must not be collapsed into one universal RCA.

## Preservation and validation

The target's unrelated local `.claude/settings.json`, `.gitignore`, `CLAUDE.md`, `docs/COMPETITIVE-COPILOT.md`, `docs/FLAKY-CAPABILITIES.md` and two untracked Copilot documents are outside the code integration. Preserve their contents and uncommitted status.

Before integration, save both checkouts' patches, changed/untracked files, hashes and three-way previews. Checkpoint the reviewed root-cause source, retain a named recoverable stash of only the 13 already-covered target edits, then fast-forward the target. Do not reapply duplicate patches over their merged implementations. Retain the RCA worktree and private test artifacts.

Local integration completed through checkpoint `ad4c1786aee185d91f18db0d5b5b7f5c03389d78`. All seven unrelated files retained their original byte hashes and uncommitted status. The 13 duplicate edits remain recoverable in stash `17595827cb010d421f1e87e97441e968d0c5916e`; they were not reapplied over the integrated implementations. The source worktree is retained.

Verify the actual target checkout after integration: typecheck, lint, full unit suite, brand guard, production build, generated-skill consistency and both Python parser self-tests. Refresh its local derived graph after code integration. Record the final branch/commit and preservation checks in the integration evidence directory. Real Office certification, a new full live cross-app run, remote push and deployment are not implied by those checks.
