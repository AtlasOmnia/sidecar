<!-- SPDX-License-Identifier: Apache-2.0 -->
# Native Sidecar Word: safe editable menus, briefs and documents

Read `sidecar_word_get_document_overview`, `sidecar_word_read_document_text` and
`sidecar_word_get_document_structure` first. Respect maxLength and truncated
paragraph previews; they are not complete exact target identities. Preserve
surrounding content and establish insertion/replacement scope. Review/proofread
alone is read-only; give source text, correction and reason before applying edits.

## Prefer native text/style

For simple menus and briefs use native paragraphs and styles, not decorative XML.
`sidecar_word_insert_text` takes text OR paragraphs, never both. Text has start,
end or selection placement; selection REPLACES the selected range. Use explicit
`location: "end"` by default, never an implicit location.
`sidecar_word_replace_selection` replaces the current selection after an exact
`sidecar_word_read_selection` check. Do not assume a prior write selected its result.

Paragraph mode constructs 1–20 paragraphs, at start/end only, total 5,000 UTF-16
units; no embedded control characters or line/paragraph breaks inside each text.
Its styleBuiltIn enum is Normal and Heading1 through Heading6, not Title or
"Heading 1". Await completion, then inspect actual text and styleBuiltIn through
`sidecar_word_get_document_structure`. Native paragraph mode requires WordApi 1.3.

`sidecar_word_apply_style` supports `style` and optional `targetText`. For exact
targeting, read and pass one unique COMPLETE paragraph unchanged. Partial,
missing or ambiguous matches refuse. Literal search and encoded query length
are bounded to 255 UTF-16 units (a caret doubles when encoded); controls/breaks
are unsupported. Never truncate or split a target to fit. When targetText is
omitted it styles the current selection: read and confirm that scope first.
Heading1..Heading6 and Normal use built-in styles; other names depend on the
document's localized/custom style definitions. A style label does not guarantee
font, size, spacing or geometry. Keep long or ambiguous targets manual rather
than broadening replacement.

`sidecar_word_format_selection` supports bold, italic, color, highlightColor and
fontSize, not font family, margins or paragraph spacing. Ask for manual adjustment
or agree a simpler design where the schema cannot express the goal. Use
`sidecar_word_apply_edit_plan` for independent unique exact replacements; inspect
each per-edit result. Do not claim all corrections applied when any are skipped.
One write, independent readback, then the next mutation; dependent edits are not
an unordered batch. Workbench's Word contract differs: its style tool is
selection-only and it does NOT have this paragraph insertion or targetText mode.
Do not interchange tool prefixes or payloads.

Static schema example, not host-tested (fictional copy only):
```json
{"tool":"sidecar_word_insert_text","arguments":{"location":"end","paragraphs":[{"text":"Sample menu","styleBuiltIn":"Heading1"},{"text":"Coffee — USD 2.50 (fictional)","styleBuiltIn":"Normal"}]}}
```

## OOXML is a separate, risky contract

`sidecar_word_insert_ooxml` takes an XML string, not HTML, Markdown, a DOCX ZIP
or a file path. The executor bounds length to 20,000 UTF-16 units and passes the
raw string to Word; it does not establish XML/package validity. Do not borrow
Workbench's marker-regex preflight as a guarantee for this native path.
Native start/end use body insertion; selection uses range Replace. Confirm this
replacement scope explicitly. Do not escalate a simple menu just for decoration.

Word document XML is a document part rooted at w:document with w:body, w:p, w:r
and w:t. The w namespace is
http://schemas.openxmlformats.org/wordprocessingml/2006/main.
Full Flat OPC is a package rooted at pkg:package in
http://schemas.microsoft.com/office/2006/xmlPackage, containing named pkg:part
entries, content types and XML/binary data. A w:p fragment is neither a complete
document nor a complete package. The executor neither converts fragments nor
verifies their complete relationships. No universal accepted payload form is
established by a regex match.

Keep package relationships, document relationships, namespace bindings, styles.xml,
numbering, hyperlinks and image parts consistent. A w:pStyle identifier is not
a UI style display name. Do not use dangling IDs from a different document.
Use a serializer to escape &, <, > and attribute quotes, reject invalid XML
characters and preserve significant whitespace. Well-formed XML still needs
actual Word host compatibility validation. Never pass HTML as OOXML. The STATIC complete-package example below is not native OOXML repair evidence.

## FIRST failed OOXML insertion: stop

After the FIRST failed OOXML insertion stop all retries. Preserve the exact
failing payload and structured error safely, then read text/structure and compare
with the pre-write state. Redact private content in any report. GeneralException
alone leaves root cause unresolved: it does not prove nothing changed, and does
not prove a partial write either. Missing payload or failed/truncated readback
means uncertainty, not permission to try again.

Do not flatten, delete/reinsert or replace the whole body automatically. A simpler
native draft needs explicit user agreement, a confirmed safe target and one write
followed by readback. If that write fails, stop. Do not promise undo or rollback
without actual product/host evidence. Native text/style readback and visual
inspection are distinct from schema/mock tests; no new native-host or live model
acceptance is claimed. Inspect pagination, clipping, price alignment and contrast
in Word before declaring the designed document accepted.

## Model-authored OOXML: complete package only

Always author a COMPLETE Flat OPC package; never a bare fragment or bare
`w:document`, even when the runtime accepts such strings. This is authoring
guidance, not a new runtime preflight, conversion or host guarantee. Use native
text/style for a simple menu. Always pass `location: "end"` explicitly unless a
verified start/selection target is agreed. OOXML `selection` REPLACES the selected
range; confirm the destructive scope before calling.

Authoring checklist:
- One single root `pkg:package`; bind `xmlns:w` on `w:document`, and include
  `/_rels/.rels` with an officeDocument relationship to `/word/document.xml`.
- Treat source strings as data. Serialize text and attribute values with XML
  escaping; reject XML-invalid characters, DTD/entity declarations, Markdown
  fences and HTML. Namespace-aware parsing must succeed before using the example.
- Include a related `/word/styles.xml` part whenever using `w:pStyle`. Without
  it Word may silently ignore the style: this is not proof of an insertion error.
- Every internal target must exist; every document `r:id`/`r:embed` must resolve.
  Do not add images, numbering or hyperlinks without all parts/relationships.
  No arbitrary external links or external/linked content in model-authored XML.
  The product-generated brief's reported external-hyperlink history is a bounded
  exception, not a licence for arbitrary model-authored links.
- No `w:sectPr` in these model-authored examples. Keep schema child order:
  `w:pPr` spacing before jc; `w:rPr` b, i, color, sz, u when present.
- Budget the entire serialized wrapper AND content at no more than 20,000 UTF-16
  code units (not Python code points); supplementary characters consume two.
  Count after escaping. Static lint does not establish native compatibility.

### STATIC minimal package — not host-tested

Exact section-6 critic example, structured like Microsoft's minimal Flat OPC
sample: package relationships plus the main document part. Fictional copy and
prices. STATIC until these exact bytes pass native insertion, independent text/
structure readback and separately authorized undo/visual checks on the intended
host. This is not an incident fix. Example arguments are `ooxml` = the entire
block and `location: "end"`; never send a selected paragraph from the block alone.

```xml
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:spacing w:after="120"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="48"/></w:rPr><w:t>Sample Bakery</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="240" w:after="60"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="8A4B08"/><w:sz w:val="28"/></w:rPr><w:t>BAKERY</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Butter croissant — $3.50</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>
```

Primary contract references (documentation identifiers, not deployment endpoints):
Microsoft Learn, “Use Office Open XML in your Word add-in”, and Word.Body.insertOoxml;
OfficeDev/office-js issue #6940 records a Microsoft assignee requiring a complete
Flat OPC package and a fragment GeneralException at Body.insertOoxml. Relevant
source paths are `schemas/word.json`, `addin/word/executor.js` in Workbench and
`src/wordTools.ts` plus generated `tool_schemas.json` in Sidecar.
