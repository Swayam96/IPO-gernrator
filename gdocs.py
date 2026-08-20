"""
Upload IPO LP as HTML → Google Doc (Drive API multipart upload).
Single API call — no batchUpdate, no index arithmetic.
Returns a shareable Google Docs URL.
"""
import json
import re
import ssl
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

_TOKEN_FILE = Path(__file__).parent / "token.json"

_SSL = ssl.create_default_context()
_SSL.check_hostname = False
_SSL.verify_mode    = ssl.CERT_NONE


# ── Google API helpers ────────────────────────────────────────────────────────

def _get_token():
    """Refresh and return a valid access token."""
    with open(_TOKEN_FILE) as f:
        t = json.load(f)
    data = urllib.parse.urlencode({
        "client_id":     t["client_id"],
        "client_secret": t["client_secret"],
        "refresh_token": t["refresh_token"],
        "grant_type":    "refresh_token",
    }).encode()
    req = urllib.request.Request(
        "https://oauth2.googleapis.com/token", data=data, method="POST"
    )
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urllib.request.urlopen(req, context=_SSL, timeout=30) as resp:
        new = json.loads(resp.read())
    t["access_token"] = new["access_token"]
    with open(_TOKEN_FILE, "w") as f:
        json.dump(t, f, indent=2)
    return new["access_token"]


def _api(method, url, body, token):
    req = urllib.request.Request(url, data=json.dumps(body).encode(), method=method)
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", "application/json")
    with urllib.request.urlopen(req, context=_SSL, timeout=30) as resp:
        return json.loads(resp.read())


def _upload_html(html, title, token):
    """Multipart upload: HTML → converted Google Doc."""
    boundary = "IPO_LP_UPLOAD_BOUNDARY"
    meta     = json.dumps({
        "name":     title,
        "mimeType": "application/vnd.google-apps.document",
    }).encode("utf-8")
    content = html.encode("utf-8")
    body = (
        f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n"
        .encode() + meta +
        f"\r\n--{boundary}\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n"
        .encode() + content +
        f"\r\n--{boundary}--".encode()
    )
    req = urllib.request.Request(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        data=body, method="POST",
    )
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type", f"multipart/related; boundary={boundary}")
    with urllib.request.urlopen(req, context=_SSL, timeout=60) as resp:
        return json.loads(resp.read())


# ── HTML builder (mirrors docx_builder sections) ─────────────────────────────

_STYLE = """
<style>
  body { font-family: Arial, sans-serif; font-size: 11pt; color: #1a1a1a; }
  h1   { color: #1F3864; font-size: 14pt; margin-top: 18pt; margin-bottom: 6pt; }
  p    { margin: 4pt 0; line-height: 1.4; }
  table{ border-collapse: collapse; width: 100%; margin: 8pt 0; }
  th   { background: #D6E4F0; font-weight: bold; padding: 4px 6px;
         border: 1px solid #BFBFBF; text-align: left; }
  td   { padding: 4px 6px; border: 1px solid #BFBFBF; }
  ul   { margin: 4pt 0 4pt 20pt; }
  li   { margin-bottom: 3pt; }
  .note{ font-style: italic; color: #555; }
  .draft{ font-style: italic; color: #C00; margin-bottom: 12pt; }
</style>
"""


def _e(t):
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _build_html(d):
    def g(key, default="TBA"):
        return str(d.get(key) or default) or default

    co    = g("company_full_name")
    fresh = g("fresh_issue_shares")
    ofs   = g("ofs_details")
    H     = [f"<!DOCTYPE html><html><head><meta charset='utf-8'>{_STYLE}</head><body>"]

    # helpers
    def h1(text):      H.append(f"<h1>{_e(text)}</h1>")
    def p(text):       H.append(f"<p>{_e(text)}</p>")
    def kv(l, v):      H.append(f"<p><b>{_e(l)}</b>{_e(v)}</p>")
    def note(text):    H.append(f'<p class="note">{_e(text)}</p>')
    def bullets(items):
        H.append("<ul>")
        for item in items: H.append(f"<li>{_e(item)}</li>")
        H.append("</ul>")

    def tbl(*rows):
        H.append('<table>')
        for row in rows:
            H.append("<tr>")
            for cell_text, is_hdr in row:
                tag = "th" if is_hdr else "td"
                H.append(f"<{tag}>{_e(str(cell_text))}</{tag}>")
            H.append("</tr>")
        H.append("</table>")

    HD = lambda t: (t, True)
    C  = lambda t: (t, False)

    # S1
    h1(f"{co} IPO")
    kv("Issue Size: ", f"{fresh} (Fresh Issue) + {ofs} (OFS)")
    kv("Price Band: ", "TBA")
    kv("Lot Size: ", "TBA")
    kv("Listing Date: ", "TBA")
    kv("Listing Exchange: ", g("listing_exchange"))

    # S2
    h1("IPO Timeline")
    kv("Opening Date: ", "TBA")
    kv("Closing Date: ", "TBA")
    kv("Allotment Date: ", "TBA")
    kv("Initiation of Refund: ", "TBA")

    # S3
    h1(f"About {co} IPO")
    p(f"The {co} IPO opens on TBA and closes on TBA. The allotment of shares will take place on TBA. "
      f"The credit of shares to the demat account will take place on TBA. "
      f"The initiation of refunds will take place on TBA. The listing of shares will take place on TBA.")
    p(f"The Equity Shares of the Company are proposed to be listed on {g('listing_exchange')}.")
    p(f"The offer consists of both a fresh issue and an offer for sale component. "
      f"The fresh issue will include {fresh}. The offer for sale portion includes {ofs}. "
      f"The total number of shares and aggregate amount are yet to be finalised.")
    p(f"{co} IPO's price band is set at TBA per share. The lot size for an application is TBA. "
      f"The minimum amount of investment required by a retail investor is TBA. {g('business_description')}")

    # S4
    h1(f"Objectives of {co} IPO")
    bullets([f"{i+1}. {obj}" for i, obj in enumerate(d.get("objects") or ["To be updated from RHP."])])
    p("Note: The Offer for Sale proceeds will accrue to the Promoter Selling Shareholders. "
      "The Company will not receive any proceeds from the Offer for Sale.")

    # S5
    h1(f"{co} IPO Valuation")
    tbl(
        [HD("Detail"), HD("Information")],
        [C("Upper Price Band"), C("TBA")],
        [C("Fresh Issue"), C(fresh)],
        [C("Offer for Sale"), C(ofs)],
        [C("EPS Diluted (in ₹) for FY 26"), C(g("eps_diluted_fy26"))],
    )

    # S6
    h1(f"{co} IPO Lot Size")
    tbl(
        [HD("Application"), HD("Lots"), HD("Shares"), HD("Amount")],
        [C("Individual investors (Retail) (Min)"), C("1"),  C("TBA"), C("TBA")],
        [C("Individual investors (Retail) (Max)"), C("13"), C("TBA"), C("TBA")],
        [C("S-HNI (Min)"), C("14"), C("TBA"), C("TBA")],
        [C("S-HNI (Max)"), C("67"), C("TBA"), C("TBA")],
        [C("B-HNI (Min)"), C("68"), C("TBA"), C("TBA")],
    )
    note("Lot size, shares per lot, and amounts to be updated once Price Band is announced with the Red Herring Prospectus.")

    # S7
    h1(f"{co} IPO Share Offer and Subscription Details")
    tbl(
        [HD("Investor Category"), HD("Shares Offered")],
        [C("QIBs"), C("Not more than 50% of the net offer")],
        [C("Non-institutional Investors (NIIs)"), C("Not less than 15% of the net offer")],
        [C("Retail-individual Investors (RIIs)"), C("Not less than 35% of the net offer")],
    )

    # S8
    h1("Industry Outlook")
    for key in ("industry_para_1", "industry_para_2", "industry_para_3"):
        if d.get(key): p(d[key])

    # S9
    h1(f"About {co}")
    for key in ("about_para_1", "about_para_2", "about_para_3"):
        if d.get(key): p(d[key])

    # S10
    h1(f"Strengths of {co}")
    bullets([f"{i+1}. {s}" for i, s in enumerate(d.get("strengths") or [])])

    # S11
    h1(f"Risks of {co}")
    bullets([f"{i+1}. {r}" for i, r in enumerate(d.get("risks") or [])])

    # S12
    h1(f"{co} Financials")
    tbl(
        [HD("Financial Year"), HD("Revenue from Operations (in ₹ crores)"),
         HD("Total Equity and Liabilities (in ₹ crores)"), HD("Return on Net Worth (in %)")],
        [C("FY 24"), C(g("revenue_fy24")), C(g("total_equity_fy24")), C(g("ronw_fy24"))],
        [C("FY 25"), C(g("revenue_fy25")), C(g("total_equity_fy25")), C(g("ronw_fy25"))],
        [C("FY 26"), C(g("revenue_fy26")), C(g("total_equity_fy26")), C(g("ronw_fy26"))],
    )

    # S13
    h1("Peer Comparison")
    peer_note = g("peer_comparison_note")
    p(peer_note)
    peers = d.get("peers") or []
    if peers and "no publicly listed" not in peer_note.lower():
        tbl(
            *([[HD("Company Name"), HD("Revenue (₹ crores)"), HD("P/E"), HD("EPS (Diluted) ₹"), HD("NAV per share ₹")]] +
              [[C(p.get("name","")), C(p.get("revenue","NA")), C(p.get("pe","NA")),
                C(p.get("eps","NA")), C(p.get("nav","NA"))] for p in peers])
        )

    # S14
    h1("Anchor Investor Bidding Date")
    p("Anchor Investor Bidding Date: TBA — one Working Day prior to the Bid/Offer Opening Date.")

    # S15
    h1("IPO Registrar and Book Running Lead Managers")
    kv("Registrar: ", g("registrar"))
    kv("Book Running Lead Manager: ", g("brlm"))
    kv("Contact Details: ", g("registered_office"))
    p(f"Email: {g('email')} | Phone: {g('phone')} | Website: {g('website')}")

    # S16
    h1(f"{co} Business Model")
    p(g("business_model"))

    # S17
    h1(f"{co} Growth Trajectory")
    p(f"{co}'s Total Income for FY 26 was ₹ {g('total_income_fy26')} crores, "
      f"whereas in FY 25 and FY 24 it was ₹ {g('total_income_fy25')} crores "
      f"and ₹ {g('total_income_fy24')} crores, respectively.")
    p(f"The Profit After Tax for FY 26 was ₹ {g('pat_fy26')} crores, "
      f"whereas in FY 25 and FY 24 it was ₹ {g('pat_fy25')} crores "
      f"and ₹ {g('pat_fy24')} crores, respectively.")
    p(f"Their EBITDA for FY 26 was ₹ {g('ebitda_fy26')} crores, "
      f"whereas in FY 25 and FY 24 it was ₹ {g('ebitda_fy25')} crores "
      f"and ₹ {g('ebitda_fy24')} crores, respectively.")

    # S18
    h1(f"{co} Market Position")
    p(g("about_para_3", f"{co} serves customers across India through its distribution network."))
    p(f"As of 31 March 2026, the company's Total Income, Profit After Tax, and EBITDA were "
      f"₹ {g('total_income_fy26')} crores, ₹ {g('pat_fy26')} crores, "
      f"and ₹ {g('ebitda_fy26')} crores, respectively.")

    # S19
    h1(f"{co} Profit and Loss Statement (in ₹ crores)")
    tbl(
        [HD("Parameter"), HD("FY 26"), HD("FY 25"), HD("FY 24")],
        [C("Total Income"),      C(g("total_income_fy26")), C(g("total_income_fy25")), C(g("total_income_fy24"))],
        [C("Profit Before Tax"), C(g("pbt_fy26")),          C(g("pbt_fy25")),          C(g("pbt_fy24"))],
        [C("Profit After Tax"),  C(g("pat_fy26")),          C(g("pat_fy25")),           C(g("pat_fy24"))],
        [C("EPS (Diluted) ₹"),   C(g("eps_diluted_fy26")),  C(g("eps_diluted_fy25")),  C(g("eps_diluted_fy24"))],
        [C("EBITDA"),            C(g("ebitda_fy26")),       C(g("ebitda_fy25")),       C(g("ebitda_fy24"))],
    )

    # S20
    h1(f"{co} Balance Sheet (in ₹ crores)")
    tbl(
        [HD("Parameter"), HD("FY 26"), HD("FY 25"), HD("FY 24")],
        [C("Profit Before Tax"),                 C(g("pbt_fy26")), C(g("pbt_fy25")), C(g("pbt_fy24"))],
        [C("Net Cash from Operating Activities"), C("TBA"), C("TBA"), C("TBA")],
        [C("Net Cash from Investing Activities"), C("TBA"), C("TBA"), C("TBA")],
        [C("Net Cash from Financing Activities"), C("TBA"), C("TBA"), C("TBA")],
        [C("Cash & Cash Equivalents"),            C("TBA"), C("TBA"), C("TBA")],
    )

    # S21
    h1(f"How to apply for {co} IPO?")
    kv("Step 1: ", "Log in to your Kotak Neo Demat account to access IPO investments. Next, select the current IPO section.")
    kv("Step 2: ", "Specify IPO details. Enter the number of lots and the price you wish to apply for.")
    kv("Step 3: ", "Enter UPI ID. After entering your UPI ID, click submit. This will place your bid with the exchange.")
    kv("Step 4: ", "Mandate Notification. Your UPI app will receive a mandate notification to block funds.")
    kv("Step 5: ", "Approve Request. Your funds will be blocked once you approve the mandate request on your UPI.")

    # S22
    h1("FAQs")
    chair = g("chairperson_or_md")
    if "Managing Director" in chair:   title_word = "Managing Director"
    elif any(x in chair for x in ("Chairperson", "Chairman")): title_word = "Chairperson"
    else:                              title_word = "Chief Executive Officer"
    name_only = re.split(r"\s*[\-,]|\s+(?:is|was)\s+", chair)[0].strip()
    kv(f"Who is the {title_word} of {co}? ", f"{name_only} is the {title_word} of {co}.")
    promoters = ", ".join(d.get("promoters") or ["TBA"])
    kv(f"Who are the Promoters of {co}? ", f"The Promoters of {co} are {promoters}.")

    # S23
    h1("Disclaimer")
    p("This article is for informational purposes only and does not constitute financial advice. "
      "It is not produced by the desk of the Kotak Securities Research Team, nor is it a report "
      "published by the Kotak Securities Research Team. The information presented is compiled from "
      "several secondary sources available on the internet and may change over time. Investors should "
      "conduct their own research and consult with financial professionals before making any investment "
      "decisions. Read the full disclaimer here.")
    p("Investments in securities market are subject to market risks, read all the related documents "
      "carefully before investing. Brokerage will not exceed SEBI prescribed limit. The securities are "
      "quoted as an example and not as a recommendation. SEBI Registration No-INZ000200137 Member Id "
      "NSE-08081; BSE-673; MSE-1024, MCX-56285, NCDEX-1262.")

    H.append("</body></html>")
    return "".join(H)


# ── Public API ────────────────────────────────────────────────────────────────

def create(d):
    """
    Build the 23-section LP as a Google Doc and return the shareable URL.
    Requires token.json in the same directory (run google_auth.py once to set up).
    """
    co    = str(d.get("company_full_name") or "Company")
    today = datetime.now().strftime("%d-%b-%Y")
    title = f"{co} IPO LP — DRHP Draft {today}"

    token  = _get_token()
    html   = _build_html(d)
    result = _upload_html(html, title, token)
    doc_id = result["id"]

    _api("POST",
         f"https://www.googleapis.com/drive/v3/files/{doc_id}/permissions",
         {"role": "reader", "type": "anyone"},
         token=token)

    return f"https://docs.google.com/document/d/{doc_id}/edit"
