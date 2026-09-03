---
name: docx-editor
description: "Creates, opens, inspects, and edits Microsoft Word .docx files via a structured Document API (REST or MCP). Use when generating or editing Word/DOCX documents, when the user provides an existing .docx as base or template, or when filling {{placeholders}}, headings, tables, lists, images, captions, TOC, headers, and footers — without editing XML or using the Word UI."
license: MIT
compatibility: "Requires the AI DOCX Editor HTTP API (Python 3.12+, FastAPI) at AIDOC_BASE_URL or http://127.0.0.1:8000. Optional LibreOffice for PDF/PNG preview."
metadata:
  version: "1.1"
  protocol: "rest+mcp"
  open-standard: "https://agentskills.io"
---

# DOCX Editor (universal)

Edit Word documents through the Document API. Do **not** edit OOXML, unzip `.docx`, or drive Microsoft Word with mouse/keyboard.

Default transport: **HTTP REST**. Prefer the bundled CLI over ad-hoc curl.

```bash
python skills/docx-editor/scripts/docx.py health
```

If the API is down, start it from the project root:

```bash
uvicorn app.main:app --host 127.0.0.1 --port 8000
```

Override the base URL with `AIDOC_BASE_URL`. MCP stdio is optional (`python -m app.mcp_server`); it uses the same operations.

On Windows, pass command JSON with `--json-file path.json` (PowerShell breaks `--json "{...}"`).

## Hard rules

1. Operate on high-level blocks: document, heading, paragraph, table, image, list, header, footer, style, caption, TOC.
2. Address elements only by stable IDs (`heading_001`, `paragraph_002`, `table_001`).
3. If the user already has a `.docx`, **open or clone it**. Do not rebuild from a blank file.
4. Inspect with `inspect` (outline + placeholders + tables) then `structure` for IDs. Load a single `block` only when you need full detail.
5. Never overwrite the original file unless the user explicitly asks. Use `save --filename other.docx`.
6. Put files only in `documents/` or `templates/`. Filenames only — no `../`, no absolute paths.
7. Prefer Word styles (`Heading 1`…`Heading 4`, `Normal`, `Caption`) over manual font size.
8. Page numbers must be Word fields (`PAGE` / `NUMPAGES`). Sumário/listas must be TOC fields, not typed page digits.
9. New documents default to proofing language `pt-BR` (avoids red squiggles on Portuguese). Call `set_language` on cloned English files.
10. After edits, `save --filename …` then optionally `render`.

## Existing document as base / template

Use this path whenever the user attaches, names, or points to a `.docx` already created.

```text
- [ ] 1. health
- [ ] 2. Copy the file into documents/ (or templates/ if it is a reusable model)
- [ ] 3. Clone with create --template (preferred) OR open to edit in place
- [ ] 4. inspect — read outline, {{placeholders}}, tables, figures, captions, language
- [ ] 5. structure — keep the block ID list; do not invent IDs
- [ ] 6. Fill variables / edit by ID; preserve styles, header, footer, numbering
- [ ] 7. structure again
- [ ] 8. save --filename novo.docx  (never the original unless asked)
```

**Clone (new file, original untouched):**

```bash
python skills/docx-editor/scripts/docx.py create --filename proposta_cliente.docx --template proposta.docx --template-source templates
python skills/docx-editor/scripts/docx.py create --filename relatorio_v2.docx --template relatorio.docx --template-source documents
```

**Open (same file; save still requires `--filename` or `--overwrite-original`):**

```bash
python skills/docx-editor/scripts/docx.py open --filename contrato.docx
python skills/docx-editor/scripts/docx.py open --source templates --filename proposta.docx
```

**Read before editing:**

```bash
python skills/docx-editor/scripts/docx.py inspect --id doc_XXXXXXXX
python skills/docx-editor/scripts/docx.py structure --id doc_XXXXXXXX
python skills/docx-editor/scripts/docx.py block --id doc_XXXXXXXX --block heading_001
```

`inspect` returns: `outline` (títulos), `placeholders` (`{{VAR}}`), `tables` (id + header), `images`, `captions`, `styles_in_use`, `language`. Use that map. Do not guess the document structure from the filename.

Fill placeholders with `variables`. Change body text with `replace_text` / `update_block` on the IDs from `structure`. Keep existing Heading styles instead of inserting a parallel outline.

## Blank document workflow

Only when there is no base file:

```text
- [ ] 1. health
- [ ] 2. create
- [ ] 3. structure
- [ ] 4. edit by ID
- [ ] 5. structure again
- [ ] 6. save as a new .docx
- [ ] 7. render (optional)
```

```bash
python skills/docx-editor/scripts/docx.py create --filename relatorio.docx --title "Relatório" --language pt-BR
```

## Edit

Prefer one batched `cmd` so either all changes apply or none do. On Windows use `--json-file`:

```bash
python skills/docx-editor/scripts/docx.py cmd --id doc_XXXXXXXX --json-file documents/_cmd.json
```

Fill `{{VAR}}`:

```bash
python skills/docx-editor/scripts/docx.py variables --id doc_XXXXXXXX --json "{\"CLIENTE\":\"Empresa ABC\",\"DATA\":\"20/08/2026\"}"
```

## Sumário, listas e legendas

Insert **Word fields** (Word updates them on open; `updateFields` is enabled):

```json
{"action": "insert_toc", "after_id": "pagebreak_001", "title": "Sumário", "levels": "1-3"}
{"action": "insert_list_of_figures", "after_id": "heading_002", "title": "Lista de figuras"}
{"action": "insert_list_of_tables", "title": "Lista de tabelas"}
{"action": "insert_list_of_quadros", "title": "Lista de quadros"}
{"action": "insert_caption", "after_id": "image_001", "kind": "figura", "text": "Pipeline de edição"}
{"action": "insert_caption", "after_id": "table_001", "kind": "tabela", "text": "Tempo de produção"}
{"action": "insert_caption", "after_id": "table_002", "kind": "quadro", "text": "Comparativo de cláusulas"}
```

`kind`: `figura` | `tabela` | `quadro`. Image `caption=` also creates a SEQ Figura field. Do not type “Figura 1” by hand.

## Tables

Never set `width_cm` above **15.5** on A4 (16 cm overflows the page). Omit `width_cm` to fit 100% of the text area.

If an older file shows broken tables or red overflow marks:

```json
{"action": "normalize_tables", "width_cm": 15.5}
{"action": "set_language", "language": "pt-BR"}
```

Red underlines on correct Portuguese = proofing language `en-US`. Fix with `set_language`.

## Save and preview

```bash
python skills/docx-editor/scripts/docx.py save --id doc_XXXXXXXX --filename relatorio_final.docx
python skills/docx-editor/scripts/docx.py render --id doc_XXXXXXXX
```

The saved file is `documents/relatorio_final.docx` and must open in Microsoft Word.

`documents/` is disposable working storage (generated .docx, PNG, JSON). You may delete those files. Keep `.gitkeep`. Do not delete `documents/.workspace/` while a session is open. Reusable models live in `templates/`, not in `documents/`.

## Response shape

Success:

```json
{"success": true, "document_id": "doc_001", "created_block_id": "paragraph_032", "data": {}}
```

Error:

```json
{"success": false, "error": {"code": "BLOCK_NOT_FOUND", "message": "Block paragraph_999 was not found."}}
```

Stop on `success: false`. Do not retry blindly with the same payload. Read `error.code`.

## Commands (short list)

| action | required | notes |
| --- | --- | --- |
| `insert_heading` / `add_heading` | `text`, `level` (1–4) | optional `after_id` / `before_id` |
| `insert_paragraph` / `add_paragraph` | `text` | optional `style`, `runs`, `after_id` |
| `update_block` / `format` | `target_id` | `text`, `runs`, or `properties` |
| `delete_block` | `target_id` | |
| `move` | `target_id` + `after_id` or `before_id` | |
| `add_table` | `rows`, `columns` | optional `data`, `after_id`; `width_cm` ≤ 15.5 |
| `set_cell` | `table_id` or `target_id`, `row`, `column`, `text` | 0-based indexes |
| `normalize_tables` | | repairs overflow widths |
| `replace_text` | `search`, `replace` | works across split runs |
| `set_variables` | `variables` | `{{KEY}}` → value |
| `bullet_list` / `numbered_list` | `items` | optional `level` |
| `page_break` / `section_break` | | section: `break_type` = `next_page` \| `continuous` \| `even_page` \| `odd_page` |
| `set_header` / `set_footer` | | `text`, `include_page_number`, `include_document_name`, `include_date` |
| `set_language` | `language` | default `pt-BR` |
| `insert_toc` | | Sumário (`levels` default `1-3`) |
| `insert_list_of_figures` / `tables` / `quadros` | | TOC field by caption label |
| `insert_caption` | `text`, `kind` | SEQ field; `figura` \| `tabela` \| `quadro` |

Full action catalog: [references/commands.md](references/commands.md). REST map: [references/api.md](references/api.md).

## Runs (mixed formatting)

Do not flatten a paragraph into one string when bold/italic must be preserved:

```json
{
  "action": "insert_paragraph",
  "after_id": "heading_001",
  "content": {
    "runs": [
      {"text": "Este texto é "},
      {"text": "importante", "bold": true},
      {"text": "."}
    ]
  }
}
```

## Gotchas

- New documents often start with an empty `paragraph_001`. Insert after a heading rather than assuming index 0 is your content.
- IDs live in hidden Word bookmarks (`_Aidoc_…`) and survive save/reopen. `document_id` is an **in-memory session**; if the API process restarts, `open` again.
- `save` without `--filename` only writes the working copy. Always `save --filename …` to publish a user-visible `.docx`.
- Batch `commands` roll back as a whole if any item fails.
- `RENDERER_UNAVAILABLE` means LibreOffice is missing; the `.docx` is still valid — skip preview.
- Path traversal and absolute paths return `PATH_TRAVERSAL` / `INVALID_FILENAME`.
- Do not log or echo the full document body.
- A TOC preview is a cache; Word asks to update fields on open so page numbers fill in.

## MCP tools (same semantics)

`document_create` `document_open` `document_inspect` `document_get_structure` `document_get_block` `document_add_heading` `document_add_paragraph` `document_add_table` `document_add_image` `document_add_list` `document_update_block` `document_delete_block` `document_move_block` `document_find` `document_replace` `document_render` `document_save`

If MCP is connected, call these tools instead of the CLI. Do not duplicate logic.

## Install this skill on any agent

Copy the folder `skills/docx-editor/` (this directory) into the host's skills path:

| Host | Typical path |
| --- | --- |
| Cross-client (preferred) | `.agents/skills/docx-editor/` |
| Claude Code | `.claude/skills/docx-editor/` |
| Cursor | `.cursor/skills/docx-editor/` |
| GitHub Copilot | `.github/skills/docx-editor/` |
| Codex / Gemini / others | `.agents/skills/docx-editor/` |

The folder name **must** stay `docx-editor` (matches `name` in frontmatter). Spec: https://agentskills.io/specification
