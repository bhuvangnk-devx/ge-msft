---
title: Document Evidence Workspace
kind: reference
skill: m365-surface-commander
topics: [evidence, snapshots, search, continuations, coverage, freshness, surgical-edits]
load_when: Reading a large source, following supplied evidence paths, or editing a selected target using broader document context.
---

# Document evidence workspace

Read the smallest sufficient evidence, preserve its source revision, and write only to the approved
native target. A paragraph edit can depend on other sections; a selected range can depend on other
tabs. Broader reading never widens the write target or removes exclusions.

## Choose live reads or captured evidence

- Use `read <selector>`, `search <text>`, `inspect <ref>` and `neighbors <ref>` for live host content.
  Use metadata discovery only when the target is unknown. `outline` supplies structure, not content.
- Use supplied `/evidence/...` paths for an immutable captured revision. `ls /evidence` discovers
  available snapshots; an empty listing means none was captured. It does not trigger capture.
- Use `read document` and each returned `next` command when complete live extraction is required.
  Reach `eof: true` and retain extraction limitations. Neither a last page nor a search result alone
  proves complete coverage. Do not flatten an unfinished paged read into a complete pipeline input.
- Reuse supplied evidence instead of pasting the entire source into a request. Read further when a
  semantic decision needs omitted content. Bounded results may span as many reads as host budgets
  permit; an exhausted budget is an incomplete task, not evidence that the source ended.

## Use executable CLI paths

| Need | CLI operation | Interpretation |
| --- | --- | --- |
| Discover captured revisions | `ls /evidence` | Returned names are snapshot IDs; never fabricate them. |
| Inspect capture scope | `cat /evidence/<snapshot>/manifest.json` | Coverage, source identity, limitations; check truncation. |
| Browse unit groups | `ls /evidence/<snapshot>/units` | Groups of up to 100 unit files. |
| Browse one group | `ls /evidence/<snapshot>/units/<group>` | Use returned file names. |
| Search captured text | `grep /evidence/<snapshot>/units "literal phrase"` | Case-sensitive literal match; bounded hit previews. |
| Read one exact unit | `cat <returned-unit-path>` | Unit ID, text, optional parent/related IDs, and text-window path. |
| Continue oversized text | `cat <textWindowPath>` then `cat <nextPath>` | Exact text windows; finish when `complete: true` and no next path remains. |

Replace angle-bracket placeholders with observed values before emitting commands. Absolute paths
use DocFs; `cat notes.md` and `grep notes.md "phrase"` still inspect named workspace artifacts.
`grep` over evidence returns a bounded first set of matching nodes, not all occurrences or a proof
of absence. Narrow to a known group or unit when necessary. Follow text windows to inspect every
occurrence within a unit. Do not assume regex, recursive AST queries, or `context=N` expands evidence
relationships; search semantics belong to each mount.

A unit file advertises `textWindowPath` before its text. If the raw file is truncated, use that path
instead of parsing incomplete JSON as complete content. Window responses carry `snapshotId`, `id`,
`offset`, `text`, `nextOffset`, `complete` and optional `nextPath`. Follow the returned path exactly;
offsets use JavaScript UTF-16 indices and must not be guessed as byte positions. Window files are
virtual; they need not appear in a directory listing. `head=N` limits lines, not JSON fields.

## Keep SDK operations at the host boundary

`EvidenceWorkspace.capture/query/readText/project`, `AssistSession.prepareEvidence`, snapshot reuse
and binding release are application SDK operations. They are not new CLI verbs. Do not emit
`capture`, `project`, `expand`, `prepareEvidence`, or a made-up `context` hint to invoke them.
If no snapshot is available, use advertised live reads or the existing `context` strategy request.
That request does not itself capture, upload, execute code, or mint a binding.

The host can query metadata with revision-bound cursors, expand declared parent/child/related links,
and assemble a bounded projection. Supplied `parentId`/`relatedIds` describe relationships inside a
snapshot; they are not necessarily native Office selectors. Resolve live edit targets through the
host's existing refs and observed text. Do not treat a unit ordinal as a stable paragraph, slide,
shape, sheet, or cell address. Relationship extraction is partial; absence of a link does not prove
that there is no semantic or formula dependency.

## Preserve coverage and required evidence

A projection's `sourceComplete: true` means capture reached EOF for its declared extraction scope.
`projectionComplete: false`, `omittedUnits` and `unresolvedRelations` still matter: the model sees a
subset, and lexical ranking does not establish semantic completeness. Search and expand until the
required definition, assumption or exception is observed. Never claim to have reasoned over every
unit merely because capture was complete. Required evidence that exceeds its budget must produce
an explicit narrower request or incomplete outcome; do not silently discard it.

Keep calculations in `analyze` artifacts when the host exposes local analysis. Capture actual native
ranges and use bounded previews for decisions. Evidence JSON is not automatically a DuckDB table or
an executable Office object. `context analytical` and `code-execution-preferred` are strategy hints;
claim computation only after a structured host result confirms execution. Do not emit generated
Office.js, Python or shell code as a substitute for native edit commands.

## Prepare a surgical effect

1. Identify the approved native target and exclusions from the task.
2. Observe its exact content; read governing definitions, assumptions, headers or exceptions from
   the broader permitted source. Preserve references to evidence used for the decision.
3. Produce the smallest native effect through live capabilities. Prefer one permitted bulk effect
   for one coherent range; do not widen the target to fit a convenient rectangle.
4. Let the host approve, revalidate evidence and native anchors, dispatch, and check readback.

An evidence-bound task validates the captured source at task start, after approval before a native
write, and before read-only completion. A change outside the target can invalidate its context.
The first dispatched effect consumes the binding, including uncertain effects. A second effect
requires fresh host-prepared context; another model turn does not refresh the binding. Do not batch
multiple effects under an evidence binding or invent a refresh command. Ordinary unbound programs
retain their advertised batching behavior. A preview token or a historical result is not verified
completion. After drift, let the host recapture/replan; after an uncertain write, inspect recovery
instead of replaying the effect.
