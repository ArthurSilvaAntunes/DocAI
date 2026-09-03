# DocAI — AI DOCX Editor

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

Motor profissional de edição de documentos `.docx` controlável por IA.

A aplicação não substitui visualmente o Microsoft Word. Ela expõe uma **API estruturada** (REST e MCP) para agentes criarem, abrirem, lerem, editarem, formatarem, reorganizarem, renderizarem e salvarem documentos Word padrão.

O arquivo final continua sendo um `.docx` editável no Microsoft Word.

Repositório: https://github.com/ArthurSilvaAntunes/DocAI

## Licença

Este projeto é distribuído sob a [licença MIT](LICENSE). Você pode usar, copiar, modificar e distribuir o software, inclusive para uso comercial, desde que mantenha o aviso de copyright e da licença.

## Uso rápido

```bash
git clone https://github.com/ArthurSilvaAntunes/DocAI.git
cd DocAI
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
pip install -r requirements.txt
python -m scripts.make_template
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

Linux/macOS:

```bash
source .venv/bin/activate
pip install -r requirements.txt
python -m scripts.make_template
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

Abra o Swagger em http://127.0.0.1:8000/docs.

### Pedir um documento a uma IA (nova conversa)

1. Deixe a API ligada no endereço acima.
2. Abra um chat no Cursor (ou outro agente) **dentro desta pasta**.
3. Escreva, por exemplo:

```text
Use a skill docx-editor.
Gere uma proposta comercial para a Empresa ABC Ltda.,
data 20/08/2026, valor R$ 35.000,00, responsável Arthur Antunes.
Salve como proposta_abc.docx e não sobrescreva o original.
```

O arquivo sai em `documents/proposta_abc.docx` e abre no Microsoft Word.

Para preencher o template pronto:

```text
Use a skill docx-editor.
Abra o template proposta.docx, preencha {{CLIENTE}}, {{DATA}}, {{VALOR}}
e {{RESPONSAVEL}} e salve como proposta_abc.docx.
```

## Objetivo

Permitir que uma IA manipule conceitos de alto nível — documento, seção, parágrafo, título, tabela, imagem, lista, cabeçalho, rodapé e estilo — sem editar XML diretamente.

## Arquitetura

```text
AI / Agent
    |
    v
API REST / MCP
    |
    v
DocumentService
    |
    v
Document Model (Pydantic)
    |
    v
DOCX Engine
    |
    +--> python-docx
    |
    +--> WordprocessingML / lxml
    |
    v
DOCX
```

Camadas:

| Camada | Pacote |
| --- | --- |
| API | `app/api` |
| Negócio | `app/services` |
| Modelo | `app/models` |
| DOCX Adapter | `app/docx_engine` |
| Renderização | `app/services/rendering_service.py` |
| Storage | `app/storage` |

A lógica de edição vive no `DocumentService`. A API REST e o servidor MCP reutilizam o mesmo serviço.

## Instalação

Requisitos: Python 3.12+.

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

No Linux/macOS:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Crie o template de exemplo:

```bash
python -m scripts.make_template
```

## Como executar

API REST (Swagger em `/docs`):

```bash
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

Ou:

```bash
python -m app
```

Servidor MCP (stdio):

```bash
python -m app.mcp_server
```

## LibreOffice (preview PDF/imagens)

A renderização usa LibreOffice headless:

```bash
soffice --headless --convert-to pdf arquivo.docx
```

### Windows

1. Instale o [LibreOffice](https://www.libreoffice.org/download/).
2. O executável costuma ficar em `C:\Program Files\LibreOffice\program\soffice.exe`.
3. Se estiver em outro caminho, defina:

```text
set AIDOC_SOFFICE_PATH=C:\caminho\soffice.exe
```

### Linux

```bash
sudo apt install libreoffice
```

### macOS

```bash
brew install --cask libreoffice
```

As imagens de preview são gravadas em `previews/{document_id}/page_001.png`.

## Segurança de arquivos

A API não aceita caminhos arbitrários do sistema. Apenas arquivos dentro de:

```text
documents/
templates/
previews/
```

Nomes com `../` são rejeitados.

Ao abrir um arquivo existente, as alterações **não sobrescrevem o original** até um `save` explícito com `overwrite_original=true` ou um `save-as` para outro nome.

## Exemplos de API

### Criar documento

```http
POST /documents
```

```json
{
  "filename": "relatorio.docx",
  "title": "Relatório",
  "author": "Agente",
  "page": {
    "size": "A4",
    "orientation": "portrait",
    "margin_top_cm": 2.5,
    "margin_bottom_cm": 2.5,
    "margin_left_cm": 2.5,
    "margin_right_cm": 2.5
  }
}
```

Resposta:

```json
{
  "success": true,
  "document_id": "doc_a1b2c3d4"
}
```

### Abrir DOCX existente

Coloque o arquivo em `documents/` e chame:

```http
POST /documents/open
```

```json
{
  "filename": "relatorio.docx",
  "source": "documents"
}
```

Templates ficam em `templates/`:

```json
{
  "filename": "proposta.docx",
  "source": "templates"
}
```

### Inspecionar estrutura

```http
GET /documents/{document_id}/structure
```

```json
{
  "success": true,
  "document_id": "doc_001",
  "data": {
    "blocks": [
      {"id": "heading_001", "type": "heading", "level": 1, "text": "Introdução"},
      {"id": "paragraph_001", "type": "paragraph", "text": "Texto introdutório."},
      {"id": "table_001", "type": "table", "rows": 5, "columns": 4}
    ]
  }
}
```

### Detalhe de um bloco

```http
GET /documents/{document_id}/blocks/{block_id}
```

### Comando único (ideal para agentes)

```http
POST /documents/{document_id}/commands
```

```json
{
  "action": "insert_paragraph",
  "after_id": "heading_003",
  "content": {
    "text": "Novo conteúdo.",
    "style": "Normal"
  }
}
```

### Lote transacional

Se um comando falhar, nenhum é aplicado.

```json
{
  "commands": [
    {"action": "replace_text", "search": "{{CLIENTE}}", "replace": "Empresa ABC"},
    {"action": "replace_text", "search": "{{DATA}}", "replace": "20/08/2026"},
    {"action": "add_paragraph", "after_id": "heading_001", "text": "Documento gerado automaticamente."}
  ]
}
```

### Salvar como novo arquivo

```http
POST /documents/{document_id}/save
```

```json
{
  "filename": "relatorio_final.docx"
}
```

### Renderizar preview

```http
POST /documents/{document_id}/render
GET  /documents/{document_id}/preview
```

## Exemplo completo — criar documento

```bash
curl -s -X POST http://127.0.0.1:8000/documents ^
  -H "Content-Type: application/json" ^
  -d "{\"filename\":\"relatorio.docx\",\"title\":\"Relatório\"}"
```

Em seguida:

```bash
curl -s -X POST http://127.0.0.1:8000/documents/DOC_ID/headings ^
  -H "Content-Type: application/json" ^
  -d "{\"text\":\"Introdução\",\"level\":1}"
```

```bash
curl -s -X POST "http://127.0.0.1:8000/documents/DOC_ID/paragraphs?after_id=heading_001" ^
  -H "Content-Type: application/json" ^
  -d "{\"text\":\"Texto introdutório.\"}"
```

```bash
curl -s -X POST http://127.0.0.1:8000/documents/DOC_ID/save ^
  -H "Content-Type: application/json" ^
  -d "{\"filename\":\"relatorio.docx\"}"
```

## Exemplo — editar DOCX existente

1. Copie `contrato.docx` para `documents/`.
2. `POST /documents/open` com `{"filename":"contrato.docx"}`.
3. `GET /documents/{id}/structure` para obter IDs.
4. `PATCH /documents/{id}/blocks/{block_id}` para alterar um parágrafo.
5. `POST /documents/{id}/replace` para substituir `{{CLIENTE}}`.
6. `POST /documents/{id}/save` com `{"filename":"contrato_editado.docx"}`.
7. Abra `documents/contrato_editado.docx` no Microsoft Word.

## Skill universal (qualquer IA)

O pacote segue o padrão aberto [Agent Skills](https://agentskills.io/specification) (`SKILL.md` + `scripts/` + `references/`). Qualquer agente compatível (Claude Code, Codex, Copilot, Gemini, Cursor, Goose, etc.) pode usá-lo.

Pasta canônica (copie esta pasta para o host):

```text
skills/docx-editor/
```

Já está também nos caminhos de descoberta:

```text
.agents/skills/docx-editor/     # padrão cross-client
.cursor/skills/docx-editor/     # Cursor neste repositório
```

O agente deve ler `SKILL.md` e chamar a API via:

```bash
python skills/docx-editor/scripts/docx.py health
python skills/docx-editor/scripts/docx.py create --filename relatorio.docx --title "Relatório"
```

A API precisa estar no ar (`uvicorn app.main:app --host 127.0.0.1 --port 8000`).

## MCP

Ferramentas disponíveis:

```text
document_create
document_open
document_get_structure
document_get_block
document_add_heading
document_add_paragraph
document_add_table
document_add_image
document_add_list
document_update_block
document_delete_block
document_move_block
document_find
document_replace
document_render
document_save
```

## Testes

```bash
pytest -q
```

O teste de renderização é ignorado se o LibreOffice não estiver instalado.

## Diretórios controlados

```text
documents/   arquivos do usuário
templates/   modelos com placeholders {{VAR}}
previews/    PDF e PNG gerados
history/     versões document_v001.docx, ...
```

## Licença

MIT — veja [LICENSE](LICENSE). Copyright (c) 2026 Arthur Silva Antunes.
