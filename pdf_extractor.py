"""
PDF text extraction — stdlib only, no external dependencies.

Strategy: extract ALL text first, then build a targeted sample by locating
the actual financial sections. This avoids the problem of fixed front/tail
sampling missing tables buried in the middle of a long DRHP.

Target sample layout (fits well within Claude's context):
  - Front matter       : first 60k chars  (cover page, summary, objects, issue details)
  - Financial summary  : 80k around "Summary Financial Information" / KPIs
  - Restated financials: 80k around "Restated Financial" section header
  - Basis for price    : 60k around "Basis for Offer Price" (EPS, RoNW, peer table)
  - Tail               : last 30k chars   (catch anything near appendices)
Total: ~300k chars, all targeted
"""
import re
import zlib


# ── Low-level PDF stream parser ───────────────────────────────────────────────

def _parse_stream(data):
    """Parse a decompressed PDF content stream and return plain text."""
    try:
        s = data.decode("latin-1")
    except Exception:
        return ""
    result = []
    i = 0
    n = len(s)
    while i < n:
        if s[i] == "(":
            j = i + 1
            buf = []
            while j < n and s[j] != ")":
                if s[j] == "\\":
                    j += 1
                    if j < n:
                        e = s[j]
                        if e.isdigit():
                            o = e
                            if j + 1 < n and s[j + 1].isdigit(): j += 1; o += s[j]
                            if j + 1 < n and s[j + 1].isdigit(): j += 1; o += s[j]
                            try: buf.append(chr(int(o, 8)))
                            except ValueError: pass
                        elif 32 <= ord(e) <= 126:
                            buf.append(e)
                else:
                    buf.append(s[j])
                j += 1
            token = "".join(buf)
            k = j + 1
            while k < n and s[k] in " \t\r\n": k += 1
            if k < n and s[k:k+2] in ("Tj", "TJ"):
                clean = "".join(c for c in token if 32 <= ord(c) <= 126)
                if len(clean) > 1:
                    result.append(clean)
            i = j + 1

        elif s[i] == "[":
            j = i + 1
            arr = []
            while j < n and s[j] != "]":
                if s[j] == "(":
                    j += 1
                    while j < n and s[j] != ")":
                        if s[j] == "\\":
                            j += 1
                            if j < n:
                                e = s[j]
                                if e.isdigit():
                                    o = e
                                    if j + 1 < n and s[j + 1].isdigit(): j += 1; o += s[j]
                                    if j + 1 < n and s[j + 1].isdigit(): j += 1; o += s[j]
                                    try:
                                        c = chr(int(o, 8))
                                        if 32 <= ord(c) <= 126: arr.append(c)
                                    except ValueError: pass
                                elif 32 <= ord(e) <= 126:
                                    arr.append(e)
                        elif 32 <= ord(s[j]) <= 126:
                            arr.append(s[j])
                        j += 1
                j += 1
            k = j + 1
            while k < n and s[k] in " \t\r\n": k += 1
            if k < n and s[k:k+2] == "TJ":
                t = "".join(arr)
                if len(t) > 1: result.append(t)
            i = j + 1

        elif s[i:i+2] in ("Td", "TD", "T*"):
            result.append(" "); i += 2
        elif s[i:i+2] == "ET":
            result.append("\n"); i += 2
        else:
            i += 1

    return " ".join(result)


def _extract_full(path):
    """Extract ALL text from a PDF. Returns the full string (can be very large)."""
    with open(path, "rb") as f:
        data = f.read()
    pattern = re.compile(
        rb"<<([^>]*)>>\s*stream\r?\n(.*?)\r?\nendstream", re.DOTALL
    )
    texts = []
    for m in pattern.finditer(data):
        header, raw = m.group(1), m.group(2)
        if b"FlateDecode" not in header and b"Fl" not in header:
            continue
        try:
            dec = zlib.decompress(raw)
        except Exception:
            try:
                dec = zlib.decompress(raw, -15)
            except Exception:
                continue
        texts.append(_parse_stream(dec))

    full = "\n".join(texts)
    full = re.sub(r"[ \t]{2,}", " ", full)
    full = re.sub(r"\n{3,}", "\n\n", full)
    return full


# ── Section-finding helpers ───────────────────────────────────────────────────

# All patterns that signal financially relevant content
_FINANCIAL_ANCHORS = [
    r"summary\s+(?:of\s+)?(?:restated\s+)?financial\s+(?:information|statements|data)",
    r"key\s+(?:financial\s+)?(?:performance\s+)?(?:indicators|highlights|metrics)",
    r"restated\s+(?:summary\s+)?(?:financial|standalone|consolidated)",
    r"(?:selected\s+)?financial\s+information",
    r"statement\s+of\s+(?:profit|income)\s+and\s+loss",
    r"profit\s+(?:and\s+loss|&\s+loss)\s+(?:account|statement)",
    r"profit\s+before\s+(?:tax|taxation)",
    r"comparison\s+of\s+(?:our\s+)?kpis",
    r"ebitda\s+is\s+calculated",
    r"total\s+income",
    r"revenue\s+from\s+operations",
]

_VALUATION_ANCHORS = [
    r"basis\s+(?:of|for)\s+(?:the\s+)?offer\s+price",
    r"earnings\s+per\s+share",
    r"return\s+on\s+net\s+worth",
    r"ronw",
    r"net\s+asset\s+value",
    r"comparison\s+(?:of\s+)?(?:accounting\s+)?ratios",
    r"weighted\s+average\s+(?:eps|earnings|ronw|return)",
    r"face\s+value.*?(?:eps|earnings\s+per)",
]

# Combined single pattern for fast scanning across the full text
_ALL_ANCHORS = _FINANCIAL_ANCHORS + _VALUATION_ANCHORS


def _collect_ranges(text, patterns, window=40_000):
    """
    Find ALL matches for every pattern and return merged (start, end) ranges.
    Uses a smaller per-hit window (40k) since we collect all hits, not just first.
    """
    raw = []
    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            # Only include hits that have at least one number nearby (actual data, not glossary)
            vicinity = text[max(0, m.start()-100): m.start()+500]
            if re.search(r'\d+[\.,]\d+', vicinity):
                raw.append((max(0, m.start() - 1_000),
                             min(len(text), m.start() + window)))

    if not raw:
        return []
    raw.sort()
    merged = [list(raw[0])]
    for start, end in raw[1:]:
        if start <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], end)
        else:
            merged.append([start, end])
    return [(s, e) for s, e in merged]


def _budget_ranges(ranges, text_len, budget):
    """
    Given merged ranges sorted by position, pick ranges greedily up to `budget` chars.
    Prioritises earlier ranges (more likely to be summary sections).
    Returns list of (start, end).
    """
    selected = []
    used = 0
    for start, end in ranges:
        chunk = end - start
        if used + chunk <= budget:
            selected.append((start, end))
            used += chunk
        else:
            # Partial: take as much as fits
            remaining = budget - used
            if remaining > 5_000:
                selected.append((start, start + remaining))
            break
    return selected


def _slice(text, ranges):
    return [text[s:e] for s, e in ranges]


# ── Public API ────────────────────────────────────────────────────────────────

def extract_text(path):
    """
    Extract a targeted sample from a DRHP/RHP PDF covering ALL financial data.

    Strategy:
      1. Front matter (first 60k) — cover page, issue details, objects
      2. ALL hits for every financial/valuation anchor across the full text,
         each with a 40k window, merged and deduplicated — this catches tables
         wherever they appear in a 2M+ char document
      3. Tail (last 20k) — catch any appendix tables

    Total budget: 380k chars. Ranges are prioritised by document position so
    summary sections (earlier) are preferred over repetitive footnote sections.
    """
    full = _extract_full(path)
    n    = len(full)

    if n <= 380_000:
        return full  # small PDF — send everything

    BUDGET = 380_000

    # 1. Front matter — always include
    front    = full[:60_000]
    tail     = full[max(0, n - 20_000):]
    reserved = len(front) + len(tail)
    remaining_budget = BUDGET - reserved

    # 2. Find all financial/valuation anchor hits with data nearby
    all_ranges = _collect_ranges(full, _ALL_ANCHORS, window=40_000)

    # Remove ranges already covered by front/tail
    mid_ranges = [
        (s, e) for s, e in all_ranges
        if e > 60_000 and s < n - 20_000
    ]

    selected = _budget_ranges(mid_ranges, n, remaining_budget)
    parts    = [front] + _slice(full, selected) + [tail]

    combined = "\n\n[--- section break ---]\n\n".join(parts)
    return combined[:BUDGET]
