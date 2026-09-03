#!/usr/bin/env python3
"""Stdlib CLI for the AI DOCX Editor REST API. Any agent can run this."""

from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.error
import urllib.request

DEFAULT_BASE = os.environ.get("AIDOC_BASE_URL", "http://127.0.0.1:8000").rstrip("/")


def api(method: str, path: str, payload: dict | None = None) -> dict:
    data = None
    headers = {"Accept": "application/json"}
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(DEFAULT_BASE + path, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
            body = resp.read().decode("utf-8")
            return json.loads(body) if body else {"success": True}
    except urllib.error.HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return {"success": False, "error": {"code": "HTTP_ERROR", "message": raw or str(exc)}}
    except urllib.error.URLError as exc:
        return {
            "success": False,
            "error": {
                "code": "API_UNAVAILABLE",
                "message": f"Cannot reach {DEFAULT_BASE} ({exc.reason}). Start: uvicorn app.main:app --host 127.0.0.1 --port 8000",
            },
        }


def emit(result: dict) -> int:
    json.dump(result, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")
    return 0 if result.get("success") else 1


def parse_json(raw: str) -> dict:
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        emit({"success": False, "error": {"code": "INVALID_JSON", "message": str(exc)}})
        raise SystemExit(1) from exc
    if not isinstance(value, dict):
        emit({"success": False, "error": {"code": "INVALID_JSON", "message": "JSON object required."}})
        raise SystemExit(1)
    return value


def main() -> int:
    parser = argparse.ArgumentParser(description="AI DOCX Editor client")
    sub = parser.add_subparsers(dest="cmd", required=True)

    sub.add_parser("health")

    p_create = sub.add_parser("create")
    p_create.add_argument("--filename", default="documento.docx")
    p_create.add_argument("--title")
    p_create.add_argument("--author")
    p_create.add_argument("--template")
    p_create.add_argument("--template-source", default="templates", choices=["documents", "templates"])
    p_create.add_argument("--language", default="pt-BR")
    p_create.add_argument("--size", default="A4", choices=["A4", "Letter"])
    p_create.add_argument("--orientation", default="portrait", choices=["portrait", "landscape"])

    p_open = sub.add_parser("open")
    p_open.add_argument("--filename", required=True)
    p_open.add_argument("--source", default="documents", choices=["documents", "templates"])

    p_id = argparse.ArgumentParser(add_help=False)
    p_id.add_argument("--id", required=True, help="document_id")

    sub.add_parser("structure", parents=[p_id])
    sub.add_parser("inspect", parents=[p_id])
    sub.add_parser("get", parents=[p_id])

    p_block = sub.add_parser("block", parents=[p_id])
    p_block.add_argument("--block", required=True)

    p_patch = sub.add_parser("patch", parents=[p_id])
    p_patch.add_argument("--block", required=True)
    p_patch.add_argument("--json", required=True)

    p_delete = sub.add_parser("delete", parents=[p_id])
    p_delete.add_argument("--block", required=True)

    p_cmd = sub.add_parser("cmd", parents=[p_id])
    p_cmd.add_argument("--json")
    p_cmd.add_argument("--json-file")

    p_find = sub.add_parser("find", parents=[p_id])
    p_find.add_argument("--text")
    p_find.add_argument("--type")
    p_find.add_argument("--style")

    p_replace = sub.add_parser("replace", parents=[p_id])
    p_replace.add_argument("--search", required=True)
    p_replace.add_argument("--replace", required=True)

    p_vars = sub.add_parser("variables", parents=[p_id])
    p_vars.add_argument("--json", required=True)

    p_save = sub.add_parser("save", parents=[p_id])
    p_save.add_argument("--filename")
    p_save.add_argument("--overwrite-original", action="store_true")

    sub.add_parser("render", parents=[p_id])
    sub.add_parser("preview", parents=[p_id])
    sub.add_parser("history", parents=[p_id])

    args = parser.parse_args()

    if args.cmd == "health":
        return emit(api("GET", "/health"))
    if args.cmd == "create":
        payload = {
            "filename": args.filename,
            "title": args.title,
            "author": args.author,
            "template": args.template,
            "template_source": args.template_source,
            "language": args.language,
            "page": {"size": args.size, "orientation": args.orientation},
        }
        payload = {k: v for k, v in payload.items() if v is not None}
        return emit(api("POST", "/documents", payload))
    if args.cmd == "open":
        return emit(api("POST", "/documents/open", {"filename": args.filename, "source": args.source}))
    if args.cmd == "structure":
        return emit(api("GET", f"/documents/{args.id}/structure"))
    if args.cmd == "inspect":
        return emit(api("GET", f"/documents/{args.id}/inspect"))
    if args.cmd == "get":
        return emit(api("GET", f"/documents/{args.id}"))
    if args.cmd == "block":
        return emit(api("GET", f"/documents/{args.id}/blocks/{args.block}"))
    if args.cmd == "patch":
        return emit(api("PATCH", f"/documents/{args.id}/blocks/{args.block}", parse_json(args.json)))
    if args.cmd == "delete":
        return emit(api("DELETE", f"/documents/{args.id}/blocks/{args.block}"))
    if args.cmd == "cmd":
        raw = args.json
        if args.json_file:
            from pathlib import Path

            raw = Path(args.json_file).read_text(encoding="utf-8")
        if not raw:
            return emit({"success": False, "error": {"code": "INVALID_JSON", "message": "Pass --json or --json-file."}})
        return emit(api("POST", f"/documents/{args.id}/commands", parse_json(raw)))
    if args.cmd == "find":
        return emit(api("POST", f"/documents/{args.id}/find", {"text": args.text, "block_type": args.type, "style": args.style}))
    if args.cmd == "replace":
        return emit(api("POST", f"/documents/{args.id}/replace", {"search": args.search, "replace": args.replace}))
    if args.cmd == "variables":
        return emit(api("POST", f"/documents/{args.id}/variables", {"variables": parse_json(args.json)}))
    if args.cmd == "save":
        payload = {"filename": args.filename, "overwrite_original": args.overwrite_original, "keep_version": True}
        payload = {k: v for k, v in payload.items() if v is not None}
        return emit(api("POST", f"/documents/{args.id}/save", payload))
    if args.cmd == "render":
        return emit(api("POST", f"/documents/{args.id}/render"))
    if args.cmd == "preview":
        return emit(api("GET", f"/documents/{args.id}/preview"))
    if args.cmd == "history":
        return emit(api("GET", f"/documents/{args.id}/history"))
    return 1


if __name__ == "__main__":
    sys.exit(main())
