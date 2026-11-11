<!-- SPDX-License-Identifier: Apache-2.0 -->
# Native Sidecar PowerPoint: editable, readable decks

Inspect `sidecar_powerpoint_get_presentation_overview` and
`sidecar_powerpoint_read_slide_content` for the target slide. Respect truncated
reads, real shapeId values and zero-based slideIndex. Preserve unrelated slides,
wording and order. Establish audience, purpose, facts, slide size and brand.
Use the linked design playbook for typography, grid, spacing and contrast.

Make each slide carry one message with one dominant title and supporting evidence.
Use consistent margins, readable labels/sources and a small palette. Start around
28–36 pt title and 18–24 pt body, then inspect at presentation size; reduce/split
content rather than shrinking everything. Keep editable text as text. A sparse
slide is not a reason for filler icons or generic cards.

## Supported tools, not assumed Office features

- Prefer `sidecar_powerpoint_create_text_slide` for title/body/bullets and a safe
  default layout. Await it, then read before adding another object.
- Use `sidecar_powerpoint_list_slide_layouts` for actual layoutId/slideMasterId;
  pass IDs to `sidecar_powerpoint_create_slide` or
  `sidecar_powerpoint_set_slide_layout` only when host support permits. Setting
  a layout/reordering requires PowerPointApi 1.8; respect Unsupported errors.
- `sidecar_powerpoint_add_text_box` and `sidecar_powerpoint_add_shape` accept
  slideIndex and left/top/width/height in points. Verify actual canvas size and
  keep objects within non-overlapping regions. shapeType must be supported by
  Office; do not invent values. No general move/resize tool is exposed.
- Read actual shapeId, then `sidecar_powerpoint_format_text` with slideIndex,
  shapeId and nested font {name, size, color, bold, italic}. Not fontSize.
  It requires PowerPointApi 1.4. Preserve omitted properties.
- Prefer explicit verified slide/shape identifiers with
  `sidecar_powerpoint_update_slide_text`; otherwise it acts on the selection.
  `sidecar_powerpoint_find_and_replace_text` can affect every matching shape
  when shapeId is absent: inspect matches/scope before writing.
- This Sidecar schema has no image-insertion tool, speaker-notes tool, animation,
  SmartArt or full theme editing. Do not borrow Workbench's image or notes tools.
  For images, ask the user to insert an approved asset in PowerPoint or agree
  a separate external workflow; do not rasterize the entire editable deck.
  Workbench web uses its own unprefixed tool namespace and image contract.

## Check the result

After each independent write, read slide inventory/content and check exact text,
slide count/order, shape IDs and unchanged neighbors. On timeout or uncertain
write, inspect before any retry; never duplicate a slide just to retry. Confirm
slide deletion or broad replacement; no undo/rollback promise without evidence.

Readback is not visual QA. Inspect at full presentation size and thumbnail scale:
clipping, overlap, line wraps, contrast, source readability, title/body balance,
image crop and charts/axes if present. Fix observed defects with approval and
inspect again. If host viewing is unavailable, state that visual acceptance is
pending. Static/schema/mock checks are not live model or native-host proof.

## Registered surface versus implementation inventory

This plugin has 15 registered PowerPoint operations. Image insertion, duplication
and speaker-notes capability stubs in source are not registered tool capabilities.
Workbench's separate shared image tool and implemented export/reinsert duplication
do not add Sidecar tools. Do not infer support or prohibition for another product
from the common subset. Inspect the active registered contract and exact host.
