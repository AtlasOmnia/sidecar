# Sidecar for PowerPoint — supported tools

The plugin exposes these fifteen implemented operations (not all individually live-verified) (each prefixed `sidecar_powerpoint_`):

- `sidecar_powerpoint_get_presentation_overview` — Get high-level details about the active PowerPoint presentation.
- `sidecar_powerpoint_read_selected_slides` — Read the currently selected slides from the active presentation.
- `sidecar_powerpoint_read_all_slides` — Read a summary of all slides in the active presentation.
- `sidecar_powerpoint_create_slide` — Create a new slide at the end of the presentation.
- `sidecar_powerpoint_create_text_slide` — Create one new slide with a title and optional body/bullets using safe non-overlapping default layout. Prefer this over multiple add_text_box calls when drafting deck content.
- `sidecar_powerpoint_add_text_box` — Add a text box to a specific slide.
- `sidecar_powerpoint_update_slide_text` — Replace the currently selected text in the presentation.
- `sidecar_powerpoint_add_shape` — Add a geometric shape to a specific slide.
- `sidecar_powerpoint_read_slide_content` — Read text content from shapes on a specific slide by index. Returns text from all text-bearing shapes (text boxes, rectangles with text, placeholders) on the slide.
- `sidecar_powerpoint_delete_slide` — Delete a slide by index. Has a built-in guard against deleting the last remaining slide — if the presentation has only one slide, returns an error instead.
- `sidecar_powerpoint_reorder_slide` — Move a slide from its current position to a new position within the presentation using Slide.moveTo (PowerPointApi 1.8).
- `sidecar_powerpoint_find_and_replace_text` — Find and replace text on a slide. Searches shape text content using JavaScript-side string matching and replaces all occurrences. If shapeId is provided, only that shape is affected; otherwise searches all text-bearing shapes on the slide.
- `sidecar_powerpoint_list_slide_layouts` — List all slide masters and their available layouts. Returns slide master names/id and the layout names/types for each master.
- `sidecar_powerpoint_set_slide_layout` — Change the layout of a specific slide by applying a layout ID. Uses Slide.applyLayout() (PowerPointApi 1.8).
- `sidecar_powerpoint_format_text` — Format text in a shape by setting font properties (bold, italic, color, size, name). Uses TextRange.font (PowerPointApi 1.4). Only provided properties are applied.

Speaker notes, slide duplication and image insertion are unsupported and are not plugin tools. Canonical pane handlers remain for backwards compatibility, not as advertised capabilities.

Tools require a trusted chat session paired with an approved sidebar. Office.js API support varies by host version. Live-tested in desktop PowerPoint for Mac at `315f1f0`, 2026-11-11 about 00:32:
Connect and Allow; one prompt changed slide 1's title and subtitle, verified by readback
from the deck, and produced exactly one request to Hermes. After a plugin reload with the
sidebar open, it showed “Lost the connection to Hermes on this computer.” with a Try again
button; Try again reconnected. Revoke dropped the sidebar to “needs your OK”.

In the earlier run at `3682e5f`, about 00:01 the same night (behaviour unchanged since),
one prompt changed slide 1's title and subtitle and added a second slide with three bullets,
verified by readback. The request stayed open for about 24 seconds with no timeout or repeat. After a plugin reload in that run the sidebar showed “Not connected” and did not reconnect within 10 seconds; reopening it connected.
Reversal on request restored the deck (no real undo). Approval was remembered after restarting
PowerPoint; the next Connect after revoke asked for approval again.

PowerPoint needs Try again (or reopen the sidebar) after a plugin reload; Excel and Word
reconnect by themselves. Both runs used a private test site via the `addin_base` setting,
with the CA trusted in the login keychain only.
Windows desktop PowerPoint is untested. Dark theme, keyboard-only use, PowerPoint for the web
and TLS-only trust were not tested. The plugin-served same-origin default in 1.2.0 has not yet been live-tested in Office. There is no real undo: Hermes reverses an edit by editing again. Save a copy before important changes.
