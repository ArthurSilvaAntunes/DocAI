# REST API reference

Base URL: `AIDOC_BASE_URL` or `http://127.0.0.1:8000`. Swagger: `/docs`.

All JSON responses use `{success, document_id, created_block_id, data, error}`.

## Documents

| Method | Path | Body |
| --- | --- | --- |
| POST | `/documents` | `{filename, title?, author?, template?, template_source?, language?, page?}` |
| POST | `/documents/open` | `{filename, source: "documents"\|"templates"}` |
| POST | `/documents/upload` | multipart file |
| GET | `/documents` | open sessions |
| GET | `/documents/{id}` | session info |
| GET | `/documents/{id}/inspect` | outline, placeholders, tables, figures, captions, language |
| GET | `/documents/{id}/structure` | compact block list |
| GET | `/documents/{id}/blocks/{block_id}` | full block |
| PATCH | `/documents/{id}/blocks/{block_id}` | `{text?, runs?, style?, ...}` |
| DELETE | `/documents/{id}/blocks/{block_id}` | |
| POST | `/documents/{id}/commands` | one command **or** `{commands:[...]}` |
| POST | `/documents/{id}/find` | `{text?, block_type?, style?, block_id?}` |
| POST | `/documents/{id}/replace` | `{search, replace, block_id?}` |
| POST | `/documents/{id}/variables` | `{variables:{KEY:value}}` |
| POST | `/documents/{id}/save` | `{filename?, overwrite_original?, keep_version?}` |
| POST | `/documents/{id}/save-as` | `{filename}` |
| GET | `/documents/{id}/download` | raw `.docx` |
| GET | `/documents/{id}/history` | operation log |
| PUT | `/documents/{id}/header` | `{text?, include_page_number?, include_document_name?, include_date?}` |
| PUT | `/documents/{id}/footer` | same |
| POST | `/documents/{id}/render` | DOCX→PDF→PNG |
| GET | `/documents/{id}/preview` | `{pages:[{page, image}]}` |

`page`: `{size:"A4"|"Letter", orientation:"portrait"|"landscape", margin_top_cm, margin_bottom_cm, margin_left_cm, margin_right_cm}`.

`template_source`: `templates` (default) or `documents`. `language` defaults to `pt-BR`.

## Content

| Method | Path | Notes |
| --- | --- | --- |
| POST | `/documents/{id}/paragraphs?after_id=&before_id=` | `{text, runs?, style?, alignment?, ...}` |
| POST | `/documents/{id}/headings?after_id=` | `{text, level}` |
| POST | `/documents/{id}/lists?after_id=` | `{items, list_type:"bullet"|"numbered", level?}` |
| POST | `/documents/{id}/page-breaks?after_id=` | |
| POST | `/documents/{id}/section-breaks?break_type=next_page&after_id=` | |
| POST | `/documents/{id}/tables?after_id=` | `{rows, columns, data?, style?}` |
| POST | `/documents/{id}/tables/{table_id}/cells` | `{row, column, text, bold?, alignment?, shading?}` |
| POST | `/documents/{id}/tables/{table_id}/rows` | `{values?}` |
| DELETE | `/documents/{id}/tables/{table_id}/rows/{index}` | |
| POST | `/documents/{id}/tables/{table_id}/columns` | `{values?}` |
| DELETE | `/documents/{id}/tables/{table_id}/columns/{index}` | |
| POST | `/documents/{id}/tables/{table_id}/merge` | `{start_row, start_col, end_row, end_col}` |
| POST | `/documents/{id}/images` | multipart `file` + query `width_cm`, `caption`, `after_id` |
| GET | `/documents/{id}/styles` | |
| GET | `/documents/{id}/styles/{name}` | |
| POST | `/documents/{id}/styles` | create/update |
| POST | `/documents/{id}/blocks/{block_id}/style` | `{style}` |

Row/column indexes are **0-based**.

## Error codes

`DOCUMENT_NOT_FOUND` `BLOCK_NOT_FOUND` `FILE_NOT_FOUND` `INVALID_FILENAME` `PATH_TRAVERSAL` `INVALID_DOCX` `STYLE_NOT_FOUND` `CELL_NOT_FOUND` `UNKNOWN_ACTION` `VALIDATION_ERROR` `RENDERER_UNAVAILABLE` `RENDER_FAILED` `API_UNAVAILABLE` `NO_ORIGINAL`
