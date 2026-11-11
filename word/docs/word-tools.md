# Hermes Word Tool Reference

The Office.js command engine is `src/wordTools.ts`. Every tool returns a JSON string with `ok: true` on success or `ok: false` and an error on failure.

## Current tools

| Tool | Purpose | Arguments |
|---|---|---|
| `get_document_overview` | Character, word, Word API paragraph, and selection summary. | none |
| `read_selection` | Read selected text. | none |
| `read_document_text` | Read bounded body text. | `maxLength?` |
| `apply_edit_plan` | Validate and apply up to 100 exact replacements in one Word batch; unique matches are changed, while missing or ambiguous matches are reported without mutation. | `edits[]`: `find`, `replace`, `matchCase?`, `matchWholeWord?`, `highlightColor?`, `reason?` |
| `get_document_structure` | Read whitespace-preserving bounded paragraph previews and actual styles. | `maxParagraphs?` (1-100, default 10) |
| `insert_text` | Insert plain text OR native styled paragraphs. | `text` OR `paragraphs[]`: `text`, `styleBuiltIn`; `location?`: `start`, `end` (plain text also supports `selection`) |
| `replace_selection` | Replace selection and optionally highlight it. | `text`, `highlightColor?` |
| `find_replace` | Replace body matches and optionally highlight them. | `find`, `replace`, `matchCase?`, `matchWholeWord?`, `highlightColor?` |
| `format_selection` | Apply direct font formatting. | `bold?`, `italic?`, `color?`, `highlightColor?`, `fontSize?` |
| `apply_style` | Apply a style to selected paragraphs or one unique exact complete paragraph; returns observed styles. | `style`, `targetText?` |
| `insert_ooxml` | Insert trusted WordprocessingML. | `ooxml`, `location?`: `start`, `end`, `selection` |
| `insert_comment` | Add a comment to the current selection (requires WordApi 1.4). | `text` |

## Aliases (backwards compatibility)

| Old name | Maps to | Reason |
|---|---|---|
| `get_document_text` | `read_document_text` | Legacy document-read name. |
| `read_document_range` | `read_document_text` | Range reads are not implemented; this currently reads bounded body text. |

Aliases are defined by `toolAliases` in `src/wordTools.ts`.

## Capability matrix

| Tool | I/O | Destructive | Needs selection |
|---|---|---:|---:|
| `get_document_overview` | read | no | no |
| `read_selection` | read | no | yes |
| `read_document_text` | read | no | no |
| `apply_edit_plan` | write, batch | yes | no |
| `get_document_structure` | read | no | no |
| `insert_text` | write | yes | only for `selection` |
| `replace_selection` | write | yes | yes |
| `find_replace` | write | yes, bulk | no |
| `format_selection` | write | visual | yes |
| `apply_style` | write | visual | only without `targetText` |
| `insert_ooxml` | write | yes | only for `selection` |
| `insert_comment` | write | annotation | yes |

## Safety and compatibility

- Historical pre-repair failure (superseded by accepted d4d0de9): tested live in desktop Word for Mac on 2026-11-11 at about 00:29, with the plugin and sidebar built from `7eaf1fc`. Connect and Allow worked. The prompt “Write the line Quarterly Summary as a Heading 1, then one short paragraph about fruit prices” wrote both paragraphs, which were read back from the document, but the first line ended as Normal, not Heading 1. OOXML merged the heading/body before an `apply_edit_plan` search failed with `SearchStringInvalidOrTooLong` on a retained 312-character ASCII find. Later recovery restored separate text paragraphs, but final style readback was Normal; the precise style-loss operation is unisolated. Revoke dropped the sidebar to “needs your OK”.

- Earlier the same night at `2387154` (historical run): heading text and a bulleted list written with readback; Heading 1 style applied when asked; reversal on request removed the text but left empty paragraphs (no real undo). Plugin reload with the sidebar open reconnected by itself; approval was remembered after restarting Word; the next Connect after revoke asked for approval again.

- The sidebar was served from a private test site via `addin_base`; the CA was trusted in the login keychain only. Windows desktop Word is not tested. Dark theme, keyboard-only use, Word for the web, the public GitHub Pages origin, and TLS-only trust are not tested.

- Retained evidence identifies the failed search as 312 ASCII characters, beyond Word’s 255-character search bound. OOXML had already merged the heading/body before that failure; final Heading 1 was Normal, but the exact operation losing style remains unisolated. The accepted d4d0de9 repair passed 429 JavaScript and 213 Python tests and bounded native Word 16.114.4/macOS 27.2 acceptance (exact Heading1/Normal paragraphs, targeted style repair, search boundaries, short-edit native Office Undo, bold/table/image preservation and revoke). Picture-only paragraph sharing, arbitrary rich formatting, fields/content controls, comprehensive native Unicode/encoded-boundary and 5000-character replacement remain unqualified. The repaired bounds are: source and encoded queries are capped at 255 UTF-16 units, replacement capacity stays 5000, actual styles are read back, and failed searches must stop destructive recovery. Guidance is not a hard execution fence; use native Office Undo for recovery after uncertain mutations.
- Approving a pane permits all known read and write tools on its bound document, including new content,
  rewrites, formatting and comments. The plugin rejects unknown tools, missing/invalid sessions and revoked
  approval. It does not classify prompts as review, annotation or correction to grant access.
- The gateway's agent runs the tool loop; the sidebar executes Office.js tickets. Legacy local-agent modes,
  snapshot budgets and repetition limits are not guarantees of this Hermes gateway path.
- There is no real undo tool. Hermes reverses its own edit by editing again; this is not transactional
  rollback. Save a backup/version before important edits.
- Read before writing and keep edits small and verifiable.
- `apply_edit_plan` validates edits before entering `Word.run()`, applies unique exact matches and returns
  an ordered result for each edit. Entries must target independent, non-overlapping text. Zero-match and
  multiple-match entries are not changed; this is not an all-or-nothing transaction.
- `insert_ooxml` accepts raw OOXML; use only content supplied or generated for the requested edit.
- The manifest requires WordApi 1.2. `insert_comment` checks WordApi 1.4 before `Word.run()` and refuses
  unsupported hosts. API availability is not evidence of Windows live acceptance.
- After updating or reloading the plugin, reopen the sidebar if it does not reconnect. Reversing an insert can leave empty paragraphs behind.

## Bounded search and style candidate (not native acceptance)

- Both direct-search tools reject sources over **255 UTF-16 code units** before entering Word, including any invalid member of a mixed plan. Literal carets encode as `^^`; the **encoded query also must fit 255**. No truncation, splitting or long-source resolver is implemented. Replacement strings still allow 5000 units.
- Searches explicitly disable wildcards, punctuation/space ignoring, prefix and suffix matching; declared case/whole-word settings remain honored. C0/C1 controls (including tabs, CR, LF and manual breaks), U+2028 and U+2029 are refused rather than mapped to guessed paragraph notation. NBSP and astral text within the UTF-16 bound are allowed, pending native qualification.
- Errors contain phase and mutation state, never raw Office error/debug text. Validation, search and comparison failures before writes report `mutationState: "none"` for **that call**, not earlier calls. A write or post-write readback failure reports `"unknown"`; Office batches are not ACID transactions and there is no automatic rollback/replay.
- Structure readback includes actual `style` and `styleBuiltIn`, whitespace-preserving text, full length and a `truncated` flag. Built-in style readback requires WordApi 1.3; unsupported hosts return `styleBuiltInSupported: false` and `styleBuiltIn: null`, not an inferred style.
- `apply_style` with `targetText` uses bounded case-sensitive literal search, requires one search match and one complete paragraph with exactly equal text, and then applies style to that paragraph object, not the selection. Ambiguous, missing, partial and changed-before-comparison targets refuse without writes. This conservative resolver can also refuse a unique complete paragraph if its text occurs elsewhere as a substring. It is not a lock against concurrent user edits after comparison. Use `Heading1` through `Heading6` or `Normal` for built-in application (WordApi 1.3 required); other style names explicitly use Word's localized/custom `style` path. Results include `observedStyles`, not merely the requested value.
- Typed `insert_text` paragraph mode accepts 1-20 objects, 5000 total text units, no embedded controls/breaks, `start`/`end` only. Each object requires `text` and `styleBuiltIn` (`Normal`, `Heading1`…`Heading6`). It uses native paragraph insertion; for a sole text-empty paragraph it inserts text at Start/End (never Replace), preserving opaque non-text content rather than assuming structural emptiness. It applies styles to the inserted/reused paragraph objects, awaits write sync, then loads actual paragraph text/styles. No incomplete OOXML package is generated. It does not reconstruct arbitrary rich formatting.
- Await each mutation, then read actual text, paragraph boundaries and styles before claiming success. Failed/refused search means **stop destructive recovery, inspect and report**. Do not delete/reinsert, flatten, split or blindly replay. For unknown mutations ask the user to inspect/use native Office Undo. Both local and server-owned prompts carry this policy; prompts are not a hard execution fence. No adapter write latch is implemented.
- The local loop additionally performs fresh bounded text (5000 units) and structure (100 paragraphs, 200-unit previews) reads after an awaited successful write batch and before its completion reply. These reads are attached to receipts/activity. A completion claim requires matching requested/observed style and a fresh fully visible exact target, or matching typed paragraphs at the requested boundary. Contradictory, truncated, unavailable or otherwise insufficient proof stops the local turn as unverified, without replay. General edits (including replacements) lack a target/preimage-aware verifier and therefore report an unverified write receipt, not completion. An executor failure with unknown mutation state also ends the local turn for inspection/native Undo. Custom style names are shown when the built-in sentinel is `Other`. Local search-schema rejections share content-free preflight diagnostics without executing invalid calls; Hermes's external `tool_call` schema rejection can still echo argument text before plugin invocation and is outside this plugin's control. This does not lock the document, prove out-of-window content, or enforce the server-owned agent's future tool choices.

Example creation in an empty disposable document:

```json
{"paragraphs":[{"text":"Quarterly Summary","styleBuiltIn":"Heading1"},{"text":"Fruit prices rose modestly this quarter.","styleBuiltIn":"Normal"}]}
```

Pass that to `sidecar_word_insert_text`, await it, then read `sidecar_word_get_document_structure` and `sidecar_word_read_document_text`. Style-only repair with the cursor elsewhere:

```json
{"style":"Heading1","targetText":"Quarterly Summary"}
```

Pass that to `sidecar_word_apply_style`, await it, then verify the heading is `Heading1` and the separate body remains `Normal`. Native tests must still verify boundary behavior, actual styles, unaffected rich content and Office Undo before shipping. Mock suites and static builds do not prove native behavior.

## Validation

```sh
npm test
npx tsc --noEmit
npm run build:static -- --base <installer-hosted-HTTPS-site>/
```

The installer must set `addin_base` to that same HTTPS site; no public sidebar endpoint is provided here.
