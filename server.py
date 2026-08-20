#!/usr/bin/env python3
"""
IPO SEBI MCP Server — zero external dependencies (stdlib only).
Protocol: MCP JSON-RPC 2.0 over stdio (Content-Length framing).

Tools:
  search_sebi_filings   — list recent DRHP/RHP filings from SEBI
  download_ipo_pdf      — download a filing PDF to disk
  generate_ipo_lp       — full pipeline: PDF → extract text → Claude → .docx
  list_recent_pdfs      — list PDFs downloaded today
"""

import json
import os
import re
import sys
from datetime import datetime
from pathlib import Path

# Redirect stderr to log file — prevents any error output from corrupting stdout MCP framing
try:
    _log = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "server.log"), "a", encoding="utf-8")
    sys.stderr = _log
except Exception:
    pass  # if log can't be opened, stderr stays as-is (won't corrupt stdout on Windows with -u)

# Make imports work regardless of working directory
sys.path.insert(0, str(Path(__file__).parent))

import sebi
import pdf_extractor
import extractor
import docx_builder
from docx_builder import SAVE_DIR

# ── Tool definitions ──────────────────────────────────────────────────────────

TOOLS = [
    {
        "name": "search_sebi_filings",
        "description": (
            "Search SEBI's public-issues page for new DRHP and/or RHP filings. "
            "Returns company name, doc type, date, detail URL, and direct PDF URL."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "hours": {
                    "type": "integer",
                    "default": 24,
                    "description": "Look-back window in hours (default 24).",
                },
                "doc_type": {
                    "type": "string",
                    "enum": ["both", "drhp", "rhp"],
                    "default": "both",
                    "description": "Which filing type to search (default: both).",
                },
            },
            "required": [],
        },
    },
    {
        "name": "download_ipo_pdf",
        "description": (
            "Download an RHP or DRHP PDF from SEBI to the configured save directory. "
            "Use pdf_url from search_sebi_filings. Returns the local file path."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "pdf_url":  {"type": "string", "description": "Direct PDF URL from SEBI."},
                "company":  {"type": "string", "description": "Company name (used in filename)."},
                "doc_type": {"type": "string", "enum": ["RHP", "DRHP"], "description": "Filing type."},
            },
            "required": ["pdf_url", "company", "doc_type"],
        },
    },
    {
        "name": "generate_ipo_lp",
        "description": (
            "Full end-to-end pipeline for generating IPO Landing Page documents. "
            "batch=true (default): fetches ALL fresh filings on the latest date from SEBI, "
            "generates a Google Doc for each, and returns all URLs. "
            "batch=false: fetches only the single latest filing. "
            "doc_type controls DRHP or RHP. Can also accept a local pdf_path to skip the download."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "doc_type": {
                    "type": "string",
                    "enum": ["drhp", "rhp"],
                    "default": "drhp",
                    "description": "Filing type to fetch from SEBI (default: drhp).",
                },
                "batch": {
                    "type": "boolean",
                    "default": True,
                    "description": "True = generate LP for ALL filings on the latest date. False = single latest only.",
                },
                "date": {
                    "type": "string",
                    "description": "Specific date (YYYY-MM-DD). Omit to auto-detect the latest date.",
                },
                "pdf_path": {
                    "type": "string",
                    "description": "Path to an already-downloaded PDF (skips SEBI fetch).",
                },
                "company": {
                    "type": "string",
                    "description": "Company name override (optional, used with pdf_path).",
                },
                "no_cache": {
                    "type": "boolean",
                    "default": False,
                    "description": "Set true to ignore cached JSON and re-call Claude.",
                },
            },
            "required": [],
        },
    },
    {
        "name": "list_recent_pdfs",
        "description": "List PDF files downloaded today in the configured save directory.",
        "inputSchema": {"type": "object", "properties": {}, "required": []},
    },
]

# ── Tool handlers ─────────────────────────────────────────────────────────────

def tool_search(args):
    hours    = int(args.get("hours", 24))
    doc_type = args.get("doc_type", "both").lower()

    urls = []
    if doc_type in ("both", "drhp"): urls.append(sebi.SEBI_DRHP_URL)
    if doc_type in ("both", "rhp"):  urls.append(sebi.SEBI_RHP_URL)

    all_filings = []
    for url in urls:
        filings, error = sebi.fetch_filings(hours, listing_url=url)
        if error:
            return "Error: " + error
        all_filings.extend(filings)

    if not all_filings:
        return f"No new filings found on SEBI in the last {hours} hours."
    lines = [f"Found {len(all_filings)} filing(s) in last {hours} hours:\n"]
    for i, f in enumerate(all_filings, 1):
        lines.append(f"{i}. {f['company']}")
        lines.append(f"   Type   : {f['doc_type']}  |  Date: {f['date']}")
        lines.append(f"   Detail : {f['detail_url']}")
        lines.append(f"   PDF    : {f['pdf_url'] or 'Not found'}")
        lines.append("")
    return "\n".join(lines)


def tool_download(args):
    pdf_url  = args.get("pdf_url", "")
    company  = args.get("company", "Unknown")
    doc_type = args.get("doc_type", "RHP")

    safe     = re.sub(r"[^\w\s-]", "", company).strip().replace(" ", "_")
    date_str = datetime.now().strftime("%d%b%Y")
    filename = f"{safe}_{doc_type}_{date_str}.pdf"
    dest     = SAVE_DIR / filename

    data, err = sebi.http_download(pdf_url)
    if err:
        return "Download failed: " + err
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        with open(dest, "wb") as fh:
            fh.write(data)
    except OSError as exc:
        return f"Could not save file: {exc}"

    return (
        f"Downloaded.\n"
        f"  Company : {company}\n"
        f"  Type    : {doc_type}\n"
        f"  Saved to: {dest}\n\n"
        f"Call generate_ipo_lp with pdf_path=\"{dest}\" to build the LP document."
    )


def tool_generate(args):
    import generate_lp  # import here to keep MCP startup fast
    pdf_path = args.get("pdf_path", "")
    company  = args.get("company", "")
    no_cache = bool(args.get("no_cache", False))
    doc_type = args.get("doc_type", "drhp").lower()
    batch    = bool(args.get("batch", True))   # default: batch mode
    date_str = args.get("date", None)

    try:
        if pdf_path:
            # Single PDF mode
            if not os.path.exists(pdf_path):
                return f"PDF not found: {pdf_path}"
            if not company:
                company = os.path.splitext(os.path.basename(pdf_path))[0]
            doc_url, docx_path = generate_lp.run(
                pdf_path, company, doc_type=doc_type.upper(), cache=not no_cache
            )
            lines = ["LP document generated successfully."]
            if doc_url:   lines.append(f"  Google Doc : {doc_url}")
            if docx_path: lines.append(f"  Local .docx: {docx_path}")
            return "\n".join(lines)

        if batch:
            # Batch mode: all filings on the latest date (or specified date)
            results = generate_lp.run_batch(date_str, doc_type=doc_type, no_cache=no_cache)
            if not results:
                return f"No fresh {doc_type.upper()} filings found."
            lines = [f"Batch complete — {len(results)} LP(s) generated.\n"]
            for co, doc_url, docx_path in results:
                lines.append(f"  {co}")
                if doc_url:   lines.append(f"    Google Doc : {doc_url}")
                if docx_path: lines.append(f"    Local .docx: {docx_path}")
            return "\n".join(lines)
        else:
            # Single latest mode
            doc_url, docx_path = generate_lp.fetch_and_run(doc_type=doc_type, cache=not no_cache)
            lines = ["LP document generated successfully."]
            if doc_url:   lines.append(f"  Google Doc : {doc_url}")
            if docx_path: lines.append(f"  Local .docx: {docx_path}")
            return "\n".join(lines)

    except Exception as exc:
        return f"Generation failed: {exc}"


def tool_list_pdfs(_args):
    today = datetime.now().date()
    try:
        pdfs = [
            p for p in Path(SAVE_DIR).glob("*.pdf")
            if datetime.fromtimestamp(p.stat().st_mtime).date() == today
        ]
    except OSError as exc:
        return "Error: " + str(exc)
    if not pdfs:
        return "No PDFs downloaded today."
    lines = [f"Today's PDFs ({today.strftime('%d-%b-%Y')}):\n"]
    for p in sorted(pdfs, key=lambda x: x.stat().st_mtime, reverse=True):
        lines.append(f"  {p.name}  ({p.stat().st_size / 1_048_576:.1f} MB)")
    return "\n".join(lines)


HANDLERS = {
    "search_sebi_filings": tool_search,
    "download_ipo_pdf":    tool_download,
    "generate_ipo_lp":     tool_generate,
    "list_recent_pdfs":    tool_list_pdfs,
}

# ── MCP stdio transport ───────────────────────────────────────────────────────

def _read_message(stream):
    headers = {}
    while True:
        line = stream.readline()
        if not line:
            return None
        line = line.decode("utf-8").rstrip("\r\n")
        if not line:
            break
        if ":" in line:
            k, _, v = line.partition(":")
            headers[k.strip().lower()] = v.strip()
    length = int(headers.get("content-length", 0))
    if length == 0:
        return None
    return json.loads(stream.read(length).decode("utf-8"))


def _write_message(stream, obj):
    body = json.dumps(obj).encode("utf-8")
    stream.write(f"Content-Length: {len(body)}\r\n\r\n".encode("utf-8"))
    stream.write(body)
    stream.flush()


def _dispatch(msg):
    method = msg.get("method", "")
    msg_id = msg.get("id")
    params = msg.get("params") or {}

    if msg_id is None:
        return None  # notification — no response needed

    def ok(result):
        return {"jsonrpc": "2.0", "id": msg_id, "result": result}

    def err(code, message):
        return {"jsonrpc": "2.0", "id": msg_id, "error": {"code": code, "message": message}}

    if method == "initialize":
        return ok({
            "protocolVersion": "2024-11-05",
            "capabilities": {"tools": {}},
            "serverInfo": {"name": "ipo-sebi-server", "version": "3.0.0"},
        })
    if method == "tools/list":
        return ok({"tools": TOOLS})
    if method == "tools/call":
        name    = params.get("name", "")
        handler = HANDLERS.get(name)
        if not handler:
            return err(-32601, f"Unknown tool: {name}")
        try:
            text = handler(params.get("arguments") or {})
        except Exception as exc:
            text = f"Tool error: {exc}"
        return ok({"content": [{"type": "text", "text": text}]})
    if method == "ping":
        return ok({})
    return err(-32601, f"Method not found: {method}")


def main():
    stdin  = sys.stdin.buffer
    stdout = sys.stdout.buffer
    while True:
        try:
            msg = _read_message(stdin)
        except Exception:
            break
        if msg is None:
            break
        response = _dispatch(msg)
        if response is not None:
            _write_message(stdout, response)


if __name__ == "__main__":
    main()
