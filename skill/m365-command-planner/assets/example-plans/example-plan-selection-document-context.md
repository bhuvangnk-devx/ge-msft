---
title: Selected Edit With Document Context
kind: example
skill: m365-command-planner
surface: word
workflow: single-surface
topics: [selection, document-context, evidence, exclusions]
load_when: A user asks to edit only selected content while considering the entire document.
---

# Selected edit with document context

User: "Read the entire document, then rewrite only this selected paragraph to agree with its
terminology and exceptions. Leave all other paragraphs and figures unchanged."

```plan
intent rewrite
surface word
scope selection
context full-scope
step review the supported document content for terminology, definitions and exceptions
step rewrite only the selected paragraph as a tracked change using the relevant evidence
exclude preserve all other paragraphs and figures
confidence high
```

Keep `scope selection` as the edit target. `full-scope` describes the evidence needed for this
request; it does not authorize document-wide edits or require a full-file upload. The executor must
obtain coverage and current evidence through available host reads; the planner invents no snapshot
IDs, native anchors or document facts. If extraction cannot cover required content, report the gap.
