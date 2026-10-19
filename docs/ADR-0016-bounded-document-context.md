# ADR-0016 — Bounded Office content reads and complete-source summaries

**Status:** Implemented for review (October 2026); live Office and tenant acceptance pending.

## Context

An outline is useful discovery context, but cannot answer a request to summarize an entire source.
The former bare `read` path returned a compact outline. Natural-language “read the entire doc and
summarise it” could also be mistaken for an addressed CLI read. Eager document attachments followed
by resident-context compaction did not guarantee that the beginning and end reached the model.

## Decision

Keep extraction inside the add-in: StreamAssist cannot call the user's Office.js object model.
Add a shared `DocumentPage` contract and optional `DocBridge.readDocumentPage` port. Host bridges
own Office pagination; the runtime owns traversal, byte/call budgets, map/reduce requests and
coverage checks; the Gemini client owns the isolated provider route. No server is introduced.

Ordinary ambient snapshots contain metadata or a bounded paragraph prefix. A whole-document
summary uses every page of the supported extraction scope and bounded independent model requests.
It publishes only after EOF, successful reduction, and a matching second traversal. An extraction,
policy, budget or freshness failure is explicitly incomplete. Source limitations accompany the
answer. The second traversal is a drift check, not an atomic Office snapshot.

`read` returns real paged content with a continuation; `outline` remains metadata. Only recognizable
addressed reads or explicit fenced commands enter the direct CLI path. SDK clients can select
`scope: 'document'`; common English summary requests also select it. Conflicting extra grounding
fails visibly rather than being silently omitted.

## Consequences

Model requests stay bounded independently of file length, but complete reading still costs time
and total bytes. Reduction loses detail; exact spreadsheet computation remains in the Data workbench.
Hidden worksheets are included in workbook scope. Unsupported source objects are disclosed.
Independent summary requests do not advance the ordinary provider chat session; follow-up chat
cannot assume that summary is resident. Existing targeted search and attachment paths retain their
limitations and are not evidence of whole-source coverage.

See [DOCUMENT-CONTEXT.md](DOCUMENT-CONTEXT.md) for limits, APIs, tests and release gates.
