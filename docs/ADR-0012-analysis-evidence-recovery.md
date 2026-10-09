# ADR-0012: Verified analysis, scoped evidence and recoverable cell effects

Date: 2026-09-07
Status: Implemented for Excel cell analysis; host validation remains a release gate.

ADR-0011 provides lifecycle hooks and task ownership. This change gives those hooks concrete
services: versioned artifacts and a constrained SQL engine, selected-scope evidence assembly,
contextual finding actions, readback verification, and an effect recovery journal.

The implementation is documented in [COMPUTE-RECOVERY.md](COMPUTE-RECOVERY.md). It implements the
Excel tabular slice of ADR-0010. PowerPoint normalized deck artifacts, the broader VFS tool registry,
native pivot generation and arbitrary inverse replay remain outside this implemented slice.

A separate `@ge/compute` package owns engine/artifact policy. `AnalysisWorkspace` and
`RecoveryCoordinator` own analysis and effect recovery. `AssistSession` orchestrates them through
existing ownership, lifecycle, capability and approval paths. React renders typed state and actions;
it never owns Office operations or executable model callbacks.

A write request carries source/destination preconditions. A write result separately reports host
acceptance (`ok`), readback (`verified`, `mismatch`, `unknown`) and checkpoint health
(`recoveryPending`). A landed write must never be reported as unapplied merely because its
readback or post-write receipt failed. Unverified writes block dependent effects and successful task
completion. Interrupted work is inspected, never automatically replayed.

ADR-0010's memory-only requirement remains true for compute artifacts. This ADR introduces a narrow,
explicit exception for the separate recovery journal: previous cell contents and reviewed requests
persist in the Office document so recovery and undo can survive a pane reload. The diagnostic ledger
remains metadata-only. Document permissions govern journal content; credentials and conversations
are excluded. Retention is bounded, unresolved receipts are retained, and approval authority is never
persisted.

Office provides neither atomic source-check/write nor cross-device journal compare-and-swap.
Same-origin locks and hashes reduce races; they do not establish transaction isolation or exactly-once
execution. New bridges can implement the snapshot/storage ports, but must define their own identity,
freshness, readback and inverse semantics before exposing recovery controls.

## Amendment (2026-10-09): saved diagnostics in browser storage

Approved by the product owner so users can send support diagnostics after the pane or Office was
closed. The pane keeps `{ question, diagnostic }` per request in `localStorage`
(`web-shell/src/saved-diagnostics.ts`), the one place request text persists on the client:

- **Scope:** only the text the user typed in the composer (or the quick action they chose), capped at
  1,000 characters, plus the content-free `diagnostics()` snapshot (routes, step kinds, redacted
  errors, run outcomes). Never document text, model output, inserted answers, history loaded from the
  server, or credentials.
- **Per user:** the key holds a SHA-256 of the signed-in account id, never the raw id. Without a
  signed-in account nothing is stored.
- **Retention, enforced at rest:** the newest 20 entries, none older than 7 days, about 512 K
  characters; expired entries are deleted, and the user can clear them.
- **Caveat:** people sharing one OS account and Office account on one device see the same list.

Conversation history itself is not stored on the client: the pane loads past chat turns from
Gemini Enterprise (`sessions.get`) when the user opens one.
