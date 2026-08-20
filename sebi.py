"""SEBI listing page scraper — shared by server.py and generate_lp.py."""
import re
from datetime import datetime, timedelta
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

SEBI_BASE     = "https://www.sebi.gov.in"
SEBI_DRHP_URL = f"{SEBI_BASE}/sebiweb/home/HomeAction.do?doListing=yes&sid=3&ssid=15&smid=10"
SEBI_RHP_URL  = f"{SEBI_BASE}/sebiweb/home/HomeAction.do?doListing=yes&sid=3&ssid=15&smid=11"

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Referer": "https://www.sebi.gov.in/",
}


def http_get(url, timeout=30):
    """Return (html_text, error_or_None)."""
    req = Request(url, headers=HTTP_HEADERS)
    try:
        with urlopen(req, timeout=timeout) as resp:
            return resp.read().decode("utf-8", errors="replace"), None
    except (URLError, HTTPError) as exc:
        return None, str(exc)


def http_download(url, timeout=180, retries=3):
    """Return (bytes, error_or_None). Reads in chunks and retries on incomplete reads."""
    for attempt in range(retries):
        req = Request(url, headers=HTTP_HEADERS)
        try:
            with urlopen(req, timeout=timeout) as resp:
                chunks = []
                while True:
                    chunk = resp.read(1024 * 1024)  # 1 MB at a time
                    if not chunk:
                        break
                    chunks.append(chunk)
                return b"".join(chunks), None
        except (URLError, HTTPError) as exc:
            if attempt == retries - 1:
                return None, str(exc)
        except Exception as exc:
            if attempt == retries - 1:
                return None, str(exc)
    return None, "Download failed after retries"


def _parse_date(text):
    text = text.strip()
    for fmt in ("%b %d, %Y", "%b  %d, %Y", "%d-%b-%Y", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _scrape_rows(html):
    """Parse listing HTML rows. Returns list of {date_text, title, detail_url}."""
    # Match any <tr> that contains a date <td> — SEBI omits role="row" on some rows
    rows = re.findall(r"<tr[^>]*>(.*?)</tr>", html, re.DOTALL | re.IGNORECASE)
    results = []
    for row in rows:
        date_m = re.search(r"<td>([^<]+)", row, re.IGNORECASE)
        if not date_m:
            continue
        date_text = date_m.group(1).strip()
        # Skip header/non-date rows
        if not re.search(r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b", date_text):
            continue

        # Extract href and title from the <a> tag
        # Use the title= attribute — it has the clean name before any embedded <br>/<a> HTML
        a_m = re.search(r'<a\s+href="([^"]+)"[^>]*title="([^"<]+)', row, re.IGNORECASE)
        if not a_m:
            # Fallback: anchor inner text
            a_m2 = re.search(r'<a\s+href="([^"]+)"[^>]*>\s*([^<\n]+)', row, re.IGNORECASE)
            if not a_m2:
                continue
            href, title = a_m2.group(1).strip(), a_m2.group(2).strip()
        else:
            href  = a_m.group(1).strip()
            title = a_m.group(2).strip()

        results.append({
            "date_text":  date_text,
            "title":      title,
            "detail_url": href if href.startswith("http") else SEBI_BASE + href,
        })
    return results


def _scrape_pdf_url(html):
    """Extract PDF URL from an iframe src on a filing detail page."""
    m = re.search(r"file=(https?://[^\s'\"]+\.pdf)", html, re.IGNORECASE)
    return m.group(1) if m else None


def _classify(title):
    t = title.upper()
    if "DRHP" in t or "DRAFT" in t:
        return "DRHP"
    if "RHP" in t or "RED HERRING" in t:
        return "RHP"
    return "PROSPECTUS"


def _clean_company(title):
    return re.sub(
        r"\s*[-–]\s*(DRHP|RHP|DRAFT RHP|PROSPECTUS|ABRIDGED PROSPECTUS).*$",
        "", title, flags=re.IGNORECASE,
    ).strip()


def fetch_filings(hours=24, listing_url=SEBI_DRHP_URL):
    """
    Fetch recent filings from a SEBI listing page.
    Returns (list_of_filings, error_or_None).
    Each filing: {company, doc_type, date, detail_url, pdf_url}.
    """
    cutoff = datetime.now() - timedelta(hours=hours)
    html, err = http_get(listing_url)
    if err:
        return [], "Could not fetch SEBI listing: " + err

    rows = _scrape_rows(html)
    if not rows:
        return [], "Listing page fetched but no rows parsed — SEBI HTML may have changed."

    results = []
    for row in rows:
        dt = _parse_date(row["date_text"])
        if dt is None or dt < cutoff:
            continue
        pdf_url = None
        detail_html, _ = http_get(row["detail_url"])
        if detail_html:
            pdf_url = _scrape_pdf_url(detail_html)
        results.append({
            "company":    _clean_company(row["title"]),
            "doc_type":   _classify(row["title"]),
            "date":       dt.strftime("%d-%b-%Y"),
            "detail_url": row["detail_url"],
            "pdf_url":    pdf_url,
        })
    return results, None


def get_latest(listing_url=SEBI_DRHP_URL, skip_corrigenda=True):
    """
    Return (company_name, detail_url) for the most recent non-corrigendum filing.
    Raises RuntimeError if nothing found.
    """
    html, err = http_get(listing_url)
    if err:
        raise RuntimeError("Could not fetch SEBI listing: " + err)
    rows = _scrape_rows(html)
    for row in rows:
        t = row["title"].upper()
        if skip_corrigenda and any(x in t for x in ("CORRIGENDUM", "ADDENDUM", "ABRIDGED")):
            continue
        return _clean_company(row["title"]), row["detail_url"]
    raise RuntimeError("No filing found on SEBI listing page.")


def get_pdf_url(detail_url):
    """Extract PDF URL from a SEBI filing detail page. Raises RuntimeError on failure."""
    html, err = http_get(detail_url)
    if err:
        raise RuntimeError(f"Could not fetch detail page: {err}")
    url = _scrape_pdf_url(html)
    if not url:
        raise RuntimeError(f"PDF URL not found in detail page: {detail_url}")
    return url
