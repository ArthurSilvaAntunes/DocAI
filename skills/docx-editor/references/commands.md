# Command catalog

`POST /documents/{id}/commands`

One object with `action`, or a batch:

```json
{"commands": [ {"action": "..."}, {"action": "..."} ]}
```

Batch is transactional: all succeed or the document is restored.

Extra keys are allowed. `content` is merged into the payload.

## Insert / update text

```json
{"action": "insert_heading", "after_id": "table_001", "text": "Riscos", "level": 1}
{"action": "insert_paragraph", "after_id": "heading_001", "text": "Novo conteúdo.", "style": "Normal"}
{"action": "insert_paragraph_before", "target_id": "paragraph_002", "text": "Antes."}
{"action": "update_block", "target_id": "paragraph_012", "text": "Texto novo"}
{"action": "format", "target_id": "paragraph_012", "properties": {"alignment": "justify", "space_after_pt": 8}}
```

Paragraph properties: `style`, `alignment` (`left|center|right|justify`), `indent_left_cm`, `space_before_pt`, `space_after_pt`, `line_spacing`, `keep_with_next`, `page_break_before`.

Run flags: `bold`, `italic`, `underline`, `strikethrough`, `font`, `font_size`, `highlight`, `color` (`RRGGBB`), `superscript`, `subscript`.

## Structure

```json
{"action": "delete_block", "target_id": "paragraph_009"}
{"action": "move", "target_id": "heading_003", "after_id": "table_001"}
{"action": "duplicate", "target_id": "paragraph_004"}
{"action": "page_break", "after_id": "paragraph_010"}
{"action": "section_break", "after_id": "heading_002", "break_type": "next_page"}
```

Aliases: `delete`, `delete_paragraph`, `delete_table`, `delete_image`, `insert_before`, `insert_after`.

## Tables

```json
{"action": "add_table", "after_id": "heading_002", "rows": 3, "columns": 2, "data": [["Item","Valor"],["A","10"]]}
{"action": "set_cell", "table_id": "table_001", "row": 1, "column": 1, "text": "R$ 1.500,00", "bold": true}
{"action": "add_row", "table_id": "table_001", "values": ["B", "20"]}
{"action": "delete_row", "table_id": "table_001", "index": 2}
{"action": "add_column", "table_id": "table_001"}
{"action": "delete_column", "table_id": "table_001", "index": 1}
{"action": "merge_cells", "table_id": "table_001", "start_row": 0, "start_col": 0, "end_row": 0, "end_col": 1}
```

## Lists, header, replace

```json
{"action": "bullet_list", "after_id": "heading_001", "items": ["Alpha", "Beta"]}
{"action": "numbered_list", "items": ["Um", "Dois"], "level": 0}
{"action": "replace_text", "search": "{{CLIENTE}}", "replace": "Empresa ABC Ltda."}
{"action": "set_variables", "variables": {"CLIENTE": "Empresa ABC", "DATA": "20/08/2026"}}
{"action": "set_header", "text": "Confidencial", "include_document_name": true}
{"action": "set_footer", "include_page_number": true}
{"action": "apply_style", "target_id": "paragraph_003", "style": "Normal"}
{"action": "create_style", "name": "CaptionBold", "type": "paragraph", "bold": true, "based_on": "Normal"}
```

Numbered list levels render as `1.` / `1.1.` / `1.1.1.`. Bullets: `•` / `○` / `▪`.

## Indexes, captions, language, tables

```json
{"action": "insert_toc", "after_id": "pagebreak_001", "title": "Sumário", "levels": "1-3"}
{"action": "insert_list_of_figures", "title": "Lista de figuras"}
{"action": "insert_list_of_tables", "title": "Lista de tabelas"}
{"action": "insert_list_of_quadros", "title": "Lista de quadros"}
{"action": "insert_caption", "after_id": "image_001", "kind": "figura", "text": "Pipeline de edição"}
{"action": "insert_caption", "after_id": "table_001", "kind": "tabela", "text": "Tempo médio de produção"}
{"action": "set_language", "language": "pt-BR"}
{"action": "normalize_tables", "width_cm": 15.5}
```

`kind` for captions: `figura` | `tabela` | `quadro`. Image uploads with `caption=` also emit a SEQ Figura field.

Table `width_cm` must be ≤ 15.5 on A4. Omit it to use 100% of the text width. `normalize_tables` repairs files whose cell widths were stored as EMU (overflow / red page marks).

