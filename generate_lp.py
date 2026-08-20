"""
IPO LP Generator — orchestrator.

Pipeline:
  1. Fetch DRHP(s) or RHP(s) from SEBI (or use --url / --pdf)
  2. Download PDF to a temp file (deleted after processing)
  3. Extract text (smart section-targeted sampling)
  4. Call Claude via AWS Bedrock → structured JSON (cached in project/cache/)
  5. Upload as Google Doc → return shareable URL
  6. Also save a local .docx to Desktop

Usage:
  python generate_lp.py                          # latest DRHP
  python generate_lp.py --type rhp               # latest RHP
  python generate_lp.py --date 2026-08-14        # ALL fresh DRHPs on that date
  python generate_lp.py --date 2026-08-12 --type rhp
  python generate_lp.py --url <detail_url>       # specific SEBI filing page
  python generate_lp.py --pdf <path>             # already-downloaded PDF
  python generate_lp.py --pdf <path> --no-cache
"""

import argparse
import json
import os
import re
import sys
import tempfile
import threading

sys.stdout.reconfigure(encoding="utf-8")

import sebi
import pdf_extractor
import extractor
import docx_builder
import gdocs

_LISTING_URL = {
    "drhp": sebi.SEBI_DRHP_URL,
    "rhp":  sebi.SEBI_RHP_URL,
}

_SKIP_KEYWORDS = ("corrigendum", "addendum", "abridged", "prospectus")


def _is_fresh(title):
    t = title.lower()
    return not any(k in t for k in _SKIP_KEYWORDS)


def run(pdf_path, company_name, doc_type="DRHP", cache=True):
    """
    Full pipeline given a local PDF path.
    Returns (google_doc_url, local_docx_path).
    No files are left behind — PDF and any temp data are caller's responsibility.
    """
    print(f"\n[{company_name}] Extracting text from PDF...")
    pdf_text = pdf_extractor.extract_text(pdf_path)
    print(f"[{company_name}] Extracted {len(pdf_text):,} characters")

    print(f"[{company_name}] Sending to Claude (~30s)...")
    data = extractor.extract(pdf_text)
    print(f"[{company_name}] Extracted: {data.get('company_full_name', company_name)}")

    doc_url = None
    try:
        print(f"[{company_name}] Uploading to Google Docs...")
        doc_url = gdocs.create(data)
        print(f"[{company_name}] Google Doc: {doc_url}")
    except Exception as ex:
        print(f"[{company_name}] Google Docs upload failed: {ex}")

    docx_path = None
    try:
        docx_path = docx_builder.build(data)
        print(f"[{company_name}] Local .docx: {docx_path}")
    except Exception as ex:
        print(f"[{company_name}] Local .docx save failed: {ex}")

    return doc_url, docx_path


def _download_tmp(pdf_url, company):
    """Download PDF to a temp file. Returns path (caller must delete when done)."""
    print(f"[{company}] Downloading PDF...")
    data, err = sebi.http_download(pdf_url)
    if err:
        raise RuntimeError(f"Download failed: {err}")
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.write(data)
    tmp.close()
    print(f"[{company}] Downloaded {len(data)/1048576:.1f} MB")
    return tmp.name


def fetch_and_run(doc_type="drhp", detail_url=None, cache=True):
    """
    Fetch the latest single filing from SEBI, process it, delete temp PDF.
    Returns (google_doc_url, local_docx_path).
    """
    listing_url = _LISTING_URL.get(doc_type.lower(), sebi.SEBI_DRHP_URL)
    if not detail_url:
        print(f"Fetching latest {doc_type.upper()} from SEBI...")
        company, detail_url = sebi.get_latest(listing_url=listing_url)
    else:
        company = detail_url.rstrip("/").split("/")[-1].replace("-", " ").title()

    print("Fetching PDF URL...")
    pdf_url  = sebi.get_pdf_url(detail_url)
    tmp_path = _download_tmp(pdf_url, company)
    try:
        return run(tmp_path, company, doc_type=doc_type.upper(), cache=cache)
    finally:
        os.unlink(tmp_path)


def run_batch(date_str, doc_type="drhp", no_cache=False):
    """
    Generate LPs for ALL fresh DRHPs/RHPs on a given date.
    If date_str is None, auto-detects the most recent date on the listing page.
    Downloads sequentially (SEBI dislikes parallel), processes in parallel threads.
    Returns list of (company, doc_url, docx_path).
    """
    from datetime import datetime

    listing_url = _LISTING_URL.get(doc_type.lower(), sebi.SEBI_DRHP_URL)
    html, err   = sebi.http_get(listing_url)
    if err:
        sys.exit(f"Could not fetch SEBI listing: {err}")

    all_rows = sebi._scrape_rows(html)

    if date_str is None:
        # Find the most recent date that has at least one fresh filing
        dates = [
            sebi._parse_date(r["date_text"]).date()
            for r in all_rows
            if sebi._parse_date(r["date_text"]) and _is_fresh(r["title"])
        ]
        if not dates:
            sys.exit(f"No fresh {doc_type.upper()} filings found on SEBI.")
        target = max(dates)
        print(f"Auto-detected latest {doc_type.upper()} date: {target.strftime('%d-%b-%Y')}")
    else:
        for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y", "%b %d, %Y"):
            try:
                target = datetime.strptime(date_str, fmt).date()
                break
            except ValueError:
                continue
        else:
            sys.exit(f"Unrecognised date format: {date_str}")

    fresh    = [
        r for r in all_rows
        if sebi._parse_date(r["date_text"]) and
           sebi._parse_date(r["date_text"]).date() == target and
           _is_fresh(r["title"])
    ]

    if not fresh:
        print(f"No fresh {doc_type.upper()} filings found on {target.strftime('%d-%b-%Y')}.")
        return []

    print(f"\nFound {len(fresh)} fresh {doc_type.upper()} filing(s) on {target.strftime('%d-%b-%Y')}:")
    for r in fresh:
        print(f"  • {r['title']}")

    # Download all PDFs sequentially into temp files
    downloaded = []  # list of (tmp_path, company, doc_label)
    for row in fresh:
        company   = sebi._clean_company(row["title"])
        doc_label = sebi._classify(row["title"])
        try:
            pdf_url  = sebi.get_pdf_url(row["detail_url"])
            tmp_path = _download_tmp(pdf_url, company)
            downloaded.append((tmp_path, company, doc_label))
        except Exception as ex:
            print(f"  SKIP {row['title']}: {ex}")

    if not downloaded:
        print("No PDFs downloaded successfully.")
        return []

    # Process in parallel threads
    results = [None] * len(downloaded)
    lock    = threading.Lock()

    def _process(i, tmp_path, company, doc_label):
        try:
            doc_url, docx_path = run(tmp_path, company,
                                     doc_type=doc_label, cache=not no_cache)
            with lock:
                results[i] = (company, doc_url, docx_path)
        except Exception as ex:
            print(f"[{company}] ERROR: {ex}")
            with lock:
                results[i] = (company, None, None)
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

    threads = [
        threading.Thread(target=_process, args=(i, p, c, d))
        for i, (p, c, d) in enumerate(downloaded)
    ]
    for t in threads: t.start()
    for t in threads: t.join()

    return [r for r in results if r]


def main():
    parser = argparse.ArgumentParser(description="Auto IPO LP Generator")
    parser.add_argument("--type",     default="drhp", choices=["drhp", "rhp"])
    parser.add_argument("--date",     help="Date to fetch (YYYY-MM-DD). Omit to auto-detect latest.")
    parser.add_argument("--batch",    action="store_true",
                        help="Generate LP for ALL filings on the latest date (or --date)")
    parser.add_argument("--url",      help="SEBI filing detail page URL")
    parser.add_argument("--pdf",      help="Already-downloaded PDF path (skips SEBI fetch)")
    parser.add_argument("--company",  help="Company name (used with --pdf)")
    parser.add_argument("--no-cache", action="store_true")
    args = parser.parse_args()

    if args.batch or args.date:
        results = run_batch(args.date, doc_type=args.type, no_cache=args.no_cache)
        print("\n" + "=" * 60)
        print(f"Batch complete — {len(results)} LP(s) generated")
        for company, doc_url, docx_path in results:
            print(f"\n  {company}")
            if doc_url:   print(f"    Google Doc : {doc_url}")
            if docx_path: print(f"    Local .docx: {docx_path}")
        print("=" * 60)
        return

    if args.pdf:
        pdf_path = args.pdf
        company  = args.company or os.path.splitext(os.path.basename(pdf_path))[0]
        doc_url, docx_path = run(pdf_path, company,
                                 doc_type=args.type.upper(), cache=not args.no_cache)
    else:
        doc_url, docx_path = fetch_and_run(doc_type=args.type,
                                           detail_url=args.url,
                                           cache=not args.no_cache)

    print("\n" + "=" * 60)
    if doc_url:   print(f"Google Doc : {doc_url}")
    if docx_path: print(f"Local .docx: {docx_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
