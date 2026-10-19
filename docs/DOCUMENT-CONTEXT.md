# Bounded Office document context

For a whole-source summary, ask **“read the entire doc and summarise it”**, **“summarize all 200
slides”**, or **“summarise the whole workbook”**. The runtime reads the open Office source to EOF,
summarizes bounded batches, reduces the partial summaries, then checks the source again before
publishing a final answer. A compact outline alone cannot satisfy this request.

## Office reads and model requests are separate budgets

| Surface | Office.js extraction | Scope and exclusions |
| --- | --- | --- |
| Word | `body.paragraphs.load({select: 'text,styleBuiltIn', skip, top: 41})`; consume 40 paragraphs and retain one lookahead | Main-body paragraphs, preserving empty paragraphs and order. Not rendered pages. Headers, footers, footnotes, comments, drawings and images are excluded. Table text is included only as exposed by body paragraphs; no table-structure guarantee. |
| PowerPoint | Slide IDs first; paged shape metadata, then up to 40 supported text frames in one `sync()` | Every slide, including empty slides; no 60-slide traversal ceiling. Speaker notes, grouped shapes, tables, charts, images and embedded objects are excluded and disclosed. Requires PowerPointApi 1.4. |
| Excel | All worksheet IDs/names/visibility; used-range dimensions; rectangular `getRangeByIndexes` reads of at most 1,024 cells, 32 columns and 64 rows | Values and formulas on every worksheet, including hidden sheets. Tiles advance across columns and then rows, respecting used-range offsets. Charts, images, comments and formatting are excluded. Requires ExcelApi 1.4. |

Word pages are layout-dependent: a 300-page document is traversed by paragraphs, not by an assumed
page count. Excel row count alone is also insufficient to predict cost: column count, text length
and formulas matter. One exceptionally large paragraph, shape or cell can still exceed a host
transport limit; the task reports failure rather than silently skipping it.

Ambient context is cheaper: Excel captures sheet/range metadata without reading cell values;
PowerPoint captures slide IDs without reading all shapes; Word reads only a bounded paragraph
prefix. A truncated inventory is labeled. Targeted questions and explicit attachments retain their
existing paths; this change does not turn every legacy search/attachment reader into a streaming
reader. In particular, Excel's legacy literal search/whole-sheet attachment can still load the active
used range, and PowerPoint's legacy search is limited to its existing slide prefix.

## Runtime and StreamAssist

1. `DocBridge.readDocumentPage({cursor, signal})` returns validated source units with stable-in-run
   IDs, a scope, extraction limitations, and an optional `nextCursor`. No cursor means EOF.
2. The runtime packs source fragments into at most 48 KiB of data per map request, including outer JSON escaping.
   Oversized units are split on Unicode code points, accounting for JSON escaping; text is not
   sliced away. The serialized query has a separate 96 KiB ceiling. The provider envelope adds
   fixed routing fields; these are byte budgets, not promises about provider token accounting.
3. Independent `isSessionLess: true` StreamAssist requests carry only the current source batch and
   task. The `document` route mounts no command/planner skill or client-selected search store.
   Existing conversation IDs, ambient selection text and old attached context are not replayed.
   Server-side engine configuration remains under tenant control.
4. Partial summaries are bounded to 12 KiB, with a requested length below 2,000 characters.
   Groups of up to four are reduced repeatedly. Source IDs and exceptions are requested at every
   stage. This is lossy semantic compression, not a proof that every fact survives the final summary.
5. The runtime re-reads every page and compares content hashes before publishing. A changed page,
   cursor cycle, duplicate source ID, budget overflow, failed host read, blocked model response or
   incomplete stream prevents a successful final summary. Cancellation stops subsequent calls.
   Office does not supply an atomic cross-host snapshot: two matching traversals detect drift,
   but cannot guarantee that no edit occurred after a page was checked.
6. The final answer includes deterministic coverage counts, scope and exclusions. Intermediate
   text and provider citation/provenance events are not published as final evidence. Source IDs
   requested from the model are textual references, not independently verified citations.

Default task limits: 256 model calls, 64 MiB of serialized source pages, 10,000 host pages and
1,000,000 distinct source-unit IDs. Hitting a limit reports an incomplete task; it never substitutes
an early-prefix summary. Repeated host traversal costs additional Office round trips. No full
source upload, local persisted index or background synchronization is introduced.

For exact spreadsheet totals, reconciliations or large structured comparisons, use the existing
Data workbench/`analyze` SQL path and send its bounded result to the model. A prose summary is not
a substitute for deterministic numerical computation.

## SDK and command usage

```ts
// Explicit scope avoids depending on the convenience English-language recognizer.
for await (const event of session.ask('Summarize the key decisions and exceptions', {
  scope: 'document',
  signal,
})) {
  // Render activity, the final token, done, or an explicit error.
}
```

`scope: 'attached'` uses the ordinary chat path, even when the wording mentions a whole document.
If a whole-source request also supplies `grounding.queryParts`, `dataStoreSpecs` or `fileIds`, it
fails with a scope-conflict error rather than silently ignoring those sources. The document route
reads the open file. Attached context already resident in the session is not part of that scope.

`AssistSessionOptions.documentSummary` exposes `maxCalls` and `maxSourceBytes`; both must be
positive safe integers. Increasing them increases total latency and provider usage, not request
size. The English recognizer is a convenience for common summary phrases, not a general intent
classifier. Unsupported surfaces return a clear error; they do not substitute an outline.

Bare `read` / `read document` now returns the first page of actual content. Follow the returned
`next` command (for example, `read document:40` in Word) until `eof: true`. `complete: true` means
that a single initial read reached EOF, not that the last page alone contains the whole source.
A pipeline cannot silently flatten an incomplete paged result into a complete document.
`outline` remains metadata. Natural-language requests beginning with “read” no longer execute as
arbitrary naked CLI selectors; explicit fenced CLI and addressed reads remain available.

Summary calls do not advance the existing server-side chat session. The answer is displayed in the
pane, but a later ordinary chat request cannot assume the provider remembers this isolated summary;
request the document again when a follow-up needs its contents.

## Validation and remaining release gates

Mocked Office tests traverse 200 slides (including a dense final slide), 6,000 Word paragraphs
(including blanks), and all cells in 30 worksheets of 1,500 rows with both 8- and 70-column shapes.
Those are extraction coverage tests, not measured live Office latency or a claim that every wide
workbook fits the default model-call budget. Runtime tests additionally cover 300 page-equivalents,
wire-byte limits, oversized Unicode text, EOF, cancellation, source drift, request isolation and
scope conflicts. Integration tests exercise the real StreamAssist request builder.

Live Word/Excel/PowerPoint requirement-set behavior, collection pagination and tenant StreamAssist
sessionless behavior still need acceptance in supported Office hosts. No live performance or
summary-quality result is claimed by the mocked tests.
