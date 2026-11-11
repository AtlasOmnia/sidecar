---
name: office-guidance
description: Use when drafting, editing or styling a menu, brief or document in native Sidecar Word.
version: 1.0.0
license: Apache-2.0
---
<!-- SPDX-License-Identifier: Apache-2.0 -->
# Sidecar Word guidance

Load this for ordinary requests, not only questions about the add-in. Guidance
belongs to this bundled plugin, not a profile-installed legacy operator guide.
Use only `sidecar_word_*` tools from this product's registered schemas; Workbench
web tools have a different namespace/host contract. Do not change setup, trust,
approval, pairing or runtime configuration to follow a document recipe.

1. Read [word workflow](references/word-workflow.md) before creating or editing.
2. Read [portable design playbook](references/design-playbook.md) for hierarchy,
   typography, spacing, chart/menu/deck readability and anti-generic QA.
3. See [provenance](references/provenance.md) for source/license boundaries.

Inspect first, preserve sources, make the smallest authorized change and read back.
A request to review alone does not authorize edits. Respect truncation and explicit
Unsupported errors; this skill adds no tools and claims no native-host acceptance.

## Explicit Word insertion scope

For model-authored OOXML require a COMPLETE Flat OPC package; never a bare
fragment or bare `w:document`. Prefer native text/style for a simple menu. Always
pass `location: "end"` by default; use start or selection only after scope is
verified and agreed. OOXML `selection` REPLACES the current selection. Workbench's
omitted OOXML location defaults to selection, whereas Sidecar defaults to end.
This guidance changes neither runtime default. Read the linked Word workflow first.
