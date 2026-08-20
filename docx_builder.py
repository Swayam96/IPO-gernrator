"""
Word .docx builder — generates the 23-section IPO LP document.
Builds OOXML directly (no python-docx), stdlib zipfile only.
The Skytech template is embedded — no external file needed.
"""
import base64
import io
import os
import re
import zipfile
from datetime import datetime
from pathlib import Path

# ── Embedded Skytech template (base64-encoded zip, document.xml excluded) ─────
# Generated from SkytechInfinitePlatform_IPO_LP_11Aug2026.docx
# Contains: styles, numbering, theme, fonts, settings, content-types, rels
_TEMPLATE_B64 = (
    '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'
    'AJcrAAAAAA=='
)


def _get_template_files():
    """Decode the embedded template and return a dict of {filename: bytes}."""
    raw = base64.b64decode(_TEMPLATE_B64)
    buf = io.BytesIO(raw)
    with zipfile.ZipFile(buf, "r") as z:
        return {name: z.read(name) for name in z.namelist()}


# Output paths — override with env vars
DESKTOP  = Path(os.environ.get("IPO_DESKTOP", r"D:\OneDrive - Kotak Mahindra Bank Ltd\Desktop"))
SAVE_DIR = Path(os.environ.get("IPO_SAVE_DIR", str(Path.home() / "Downloads")))


# ── XML helpers ───────────────────────────────────────────────────────────────

NS = (
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" '
    'xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" '
    'xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" '
    'xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" '
    'xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" '
    'xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" '
    'xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" '
    'xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" '
    'xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" '
    'xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" '
    'xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:oel="http://schemas.microsoft.com/office/2019/extlst" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" '
    'xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" '
    'xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" '
    'xmlns:w16du="http://schemas.microsoft.com/office/word/2023/wordml/word16du" '
    'xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash" '
    'xmlns:w16sdtfl="http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock" '
    'xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
    'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh w16sdtfl w16du wp14"'
)

_TBL_BORDERS = (
    '<w:tblBorders>'
    '<w:top    w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:left   w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:right  w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
    '</w:tblBorders>'
)

_TBL_LOOK = (
    '<w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" '
    'w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>'
)

# Per-cell margin applied inside every <w:tcPr>
_TC_MAR = (
    '<w:tcMar>'
    '<w:top w:w="40" w:type="dxa"/>'
    '<w:left w:w="80" w:type="dxa"/>'
    '<w:bottom w:w="40" w:type="dxa"/>'
    '<w:right w:w="80" w:type="dxa"/>'
    '</w:tcMar>'
)

# Table-level cell margin (left/right only, matching Skytech)
_TBL_CELL_MAR = (
    '<w:tblCellMar>'
    '<w:left w:w="10" w:type="dxa"/>'
    '<w:right w:w="10" w:type="dxa"/>'
    '</w:tblCellMar>'
)

_SECT_PR = (
    '<w:sectPr>'
    '<w:pgSz w:w="11908" w:h="16833"/>'
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
    'w:header="708" w:footer="708" w:gutter="0"/>'
    '</w:sectPr>'
)


def _e(t):
    """Escape XML special chars and encode ₹ as numeric entity."""
    s = str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return s.replace("₹", "&#x20B9;")


def h1(text):
    return (
        '<w:p><w:pPr><w:pStyle w:val="Heading1"/>'
        '<w:spacing w:before="260" w:after="120"/></w:pPr>'
        f'<w:r><w:rPr><w:b/><w:bCs/><w:color w:val="1F3864"/>'
        f'<w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>'
        f'<w:t>{_e(text)}</w:t></w:r></w:p>'
    )


def body(text):
    return (
        '<w:p><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{_e(text)}</w:t></w:r></w:p>'
    )


def kv(label, value, last=False):
    spacing = (
        '<w:spacing w:after="120" w:line="259" w:lineRule="auto"/>'
        if last else
        '<w:spacing w:after="120" w:line="276" w:lineRule="auto"/>'
    )
    return (
        f'<w:p><w:pPr>{spacing}</w:pPr>'
        f'<w:r><w:rPr><w:b/><w:bCs/></w:rPr>'
        f'<w:t xml:space="preserve">{_e(label)}</w:t></w:r>'
        f'<w:r><w:t>{_e(value)}</w:t></w:r></w:p>'
    )


def bullet(text):
    return (
        '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/>'
        '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
        '<w:spacing w:after="80" w:line="276" w:lineRule="auto"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{_e(text)}</w:t></w:r></w:p>'
    )


def note(text):
    return (
        '<w:p><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>'
        f'<w:r><w:rPr><w:i/><w:iCs/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        f'<w:t xml:space="preserve">{_e(text)}</w:t></w:r></w:p>'
    )


# Cell helpers: HD = header cell, C = data cell
def HD(t): return (t, True)
def C(t):  return (t, False)


def tbl(*rows):
    """
    Build a Word table matching Skytech template formatting.

    Each row is a list of HD(...) or C(...) tuples.
    Header rows: DDEBF7 fill, bold, centered.
    Data rows: alternating FFFFFF / F2F2F2 fill.
    Data cells: first column left-aligned, all other columns right-aligned.
    """
    xml = (
        '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/>'
        '<w:tblW w:w="8700" w:type="dxa"/>'
        f'{_TBL_BORDERS}'
        f'{_TBL_CELL_MAR}'
        f'{_TBL_LOOK}'
        '</w:tblPr>'
    )

    data_row_index = 0  # counts only non-header rows for alternating fill

    for row_cells in rows:
        is_header_row = any(is_hdr for _, is_hdr in row_cells)

        if is_header_row:
            xml += '<w:tr><w:trPr><w:tblHeader/></w:trPr>'
            for cell_text, _ in row_cells:
                xml += (
                    f'<w:tc><w:tcPr>'
                    f'<w:tcW w:w="0" w:type="auto"/>'
                    f'<w:shd w:val="clear" w:color="auto" w:fill="DDEBF7"/>'
                    f'{_TC_MAR}'
                    f'</w:tcPr>'
                    f'<w:p><w:pPr>'
                    f'<w:jc w:val="center"/>'
                    f'<w:spacing w:after="60" w:line="276" w:lineRule="auto"/>'
                    f'</w:pPr>'
                    f'<w:r><w:rPr><w:b/><w:bCs/></w:rPr>'
                    f'<w:t xml:space="preserve">{_e(str(cell_text))}</w:t></w:r>'
                    f'</w:p></w:tc>'
                )
            xml += '</w:tr>'
        else:
            # Alternating fill: even index (0, 2, 4…) → FFFFFF, odd → F2F2F2
            fill = "FFFFFF" if data_row_index % 2 == 0 else "F2F2F2"
            data_row_index += 1

            xml += '<w:tr>'
            for col_idx, (cell_text, _) in enumerate(row_cells):
                if col_idx == 0:
                    # First column: left-aligned (no explicit jc element)
                    align_xml = ''
                else:
                    # All other columns: right-aligned
                    align_xml = '<w:jc w:val="right"/>'

                xml += (
                    f'<w:tc><w:tcPr>'
                    f'<w:tcW w:w="0" w:type="auto"/>'
                    f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>'
                    f'{_TC_MAR}'
                    f'</w:tcPr>'
                    f'<w:p><w:pPr>'
                    f'{align_xml}'
                    f'<w:spacing w:after="60" w:line="276" w:lineRule="auto"/>'
                    f'</w:pPr>'
                    f'<w:r>'
                    f'<w:t xml:space="preserve">{_e(str(cell_text))}</w:t></w:r>'
                    f'</w:p></w:tc>'
                )
            xml += '</w:tr>'

    xml += '</w:tbl><w:p><w:pPr><w:spacing w:after="120"/></w:pPr></w:p>'
    return xml


# ── Section builders ──────────────────────────────────────────────────────────

def _g(d, key, default="TBA"):
    """Safe dict getter with TBA fallback."""
    return str(d.get(key) or default) or default


def build_sections(d):
    """Return list of XML paragraph/table strings for all 23 sections."""
    co     = _g(d, "company_full_name")
    fresh  = _g(d, "fresh_issue_shares")
    ofs    = _g(d, "ofs_details")
    B      = []

    # S1: IPO Details
    B.append(h1(f"{co} IPO"))
    B.append(kv("Issue Size: ", f"{fresh} (Fresh Issue) + {ofs} (OFS)"))
    B.append(kv("Price Band: ", "TBA"))
    B.append(kv("Lot Size: ", "TBA"))
    B.append(kv("Listing Date: ", "TBA"))
    B.append(kv("Listing Exchange: ", _g(d, "listing_exchange"), last=True))

    # S2: IPO Timeline
    B.append(h1("IPO Timeline"))
    B.append(kv("Opening Date: ", "TBA"))
    B.append(kv("Closing Date: ", "TBA"))
    B.append(kv("Allotment Date: ", "TBA"))
    B.append(kv("Initiation of Refund: ", "TBA", last=True))

    # S3: About the IPO
    B.append(h1(f"About {co} IPO"))
    B.append(body(
        f"The {co} IPO opens on TBA and closes on TBA. The allotment of shares will take place on TBA. "
        f"The credit of shares to the demat account will take place on TBA. "
        f"The initiation of refunds will take place on TBA. The listing of shares will take place on TBA."
    ))
    B.append(body(f"The Equity Shares of the Company are proposed to be listed on {_g(d,'listing_exchange')}."))
    B.append(body(
        f"The offer consists of both a fresh issue and an offer for sale component. "
        f"The fresh issue will include {fresh}. The offer for sale portion includes {ofs}. "
        f"The total number of shares and aggregate amount are yet to be finalised."
    ))
    B.append(body(
        f"{co} IPO's price band is set at TBA per share. The lot size for an application is TBA. "
        f"The minimum amount of investment required by a retail investor is TBA. "
        f"{_g(d,'business_description')}"
    ))

    # S4: Objectives
    B.append(h1(f"Objectives of {co} IPO"))
    for i, obj in enumerate(d.get("objects") or ["To be updated from RHP."], 1):
        B.append(bullet(f"{i}. {obj}"))
    B.append(body("Note: The Offer for Sale proceeds will accrue to the Promoter Selling Shareholders. "
                  "The Company will not receive any proceeds from the Offer for Sale."))

    # S5: Valuation
    B.append(h1(f"{co} IPO Valuation"))
    B.append(tbl(
        [HD("Detail"), HD("Information")],
        [C("Upper Price Band"), C("TBA")],
        [C("Fresh Issue"), C(fresh)],
        [C("Offer for Sale"), C(ofs)],
        [C("EPS Diluted (in ₹) for FY 26"), C(_g(d, "eps_diluted_fy26"))],
    ))

    # S6: Lot Size
    B.append(h1(f"{co} IPO Lot Size"))
    B.append(tbl(
        [HD("Application"), HD("Lots"), HD("Shares"), HD("Amount")],
        [C("Individual investors (Retail) (Min)"), C("1"),  C("TBA"), C("TBA")],
        [C("Individual investors (Retail) (Max)"), C("13"), C("TBA"), C("TBA")],
        [C("S-HNI (Min)"), C("14"), C("TBA"), C("TBA")],
        [C("S-HNI (Max)"), C("67"), C("TBA"), C("TBA")],
        [C("B-HNI (Min)"), C("68"), C("TBA"), C("TBA")],
    ))
    B.append(note("Lot size, shares per lot, and amounts to be updated once Price Band is announced with the Red Herring Prospectus."))

    # S7: Subscription Details
    B.append(h1(f"{co} IPO Share Offer and Subscription Details"))
    B.append(tbl(
        [HD("Investor Category"), HD("Shares Offered")],
        [C("QIBs"), C("Not more than 50% of the net offer")],
        [C("Non-institutional Investors (NIIs)"), C("Not less than 15% of the net offer")],
        [C("Retail-individual Investors (RIIs)"), C("Not less than 35% of the net offer")],
    ))

    # S8: Industry Outlook
    B.append(h1("Industry Outlook"))
    for key in ("industry_para_1", "industry_para_2", "industry_para_3"):
        val = d.get(key)
        if val:
            B.append(body(val))

    # S9: About the Company
    B.append(h1(f"About {co}"))
    for key in ("about_para_1", "about_para_2", "about_para_3"):
        val = d.get(key)
        if val:
            B.append(body(val))

    # S10: Strengths
    B.append(h1(f"Strengths of {co}"))
    for i, s in enumerate(d.get("strengths") or [], 1):
        B.append(bullet(f"{i}. {s}"))

    # S11: Risks
    B.append(h1(f"Risks of {co}"))
    for i, r in enumerate(d.get("risks") or [], 1):
        B.append(bullet(f"{i}. {r}"))

    # S12: Financials Table
    B.append(h1(f"{co} Financials"))
    B.append(tbl(
        [HD("Financial Year"), HD("Revenue from Operations (in ₹ crores)"),
         HD("Total Equity and Liabilities (in ₹ crores)"), HD("Return on Net Worth (in %)")],
        [C("FY 24"), C(_g(d,"revenue_fy24")), C(_g(d,"total_equity_fy24")), C(_g(d,"ronw_fy24"))],
        [C("FY 25"), C(_g(d,"revenue_fy25")), C(_g(d,"total_equity_fy25")), C(_g(d,"ronw_fy25"))],
        [C("FY 26"), C(_g(d,"revenue_fy26")), C(_g(d,"total_equity_fy26")), C(_g(d,"ronw_fy26"))],
    ))

    # S13: Peer Comparison
    B.append(h1("Peer Comparison"))
    peer_note = _g(d, "peer_comparison_note")
    B.append(body(peer_note))
    peers = d.get("peers") or []
    if peers and "no publicly listed" not in peer_note.lower():
        peer_rows = [[HD("Company Name"), HD("Revenue from Operations (₹ crores)"),
                      HD("P/E Ratio"), HD("EPS (Diluted) (₹)"), HD("NAV per share (₹)")]]
        for p in peers:
            peer_rows.append([
                C(p.get("name", "")), C(p.get("revenue", "NA")),
                C(p.get("pe", "NA")), C(p.get("eps", "NA")), C(p.get("nav", "NA")),
            ])
        B.append(tbl(*peer_rows))

    # S14: Anchor Investor Date
    B.append(h1("Anchor Investor Bidding Date"))
    B.append(body("Anchor Investor Bidding Date: TBA — one Working Day prior to the Bid/Offer Opening Date."))

    # S15: Registrar + BRLM
    B.append(h1("IPO Registrar and Book Running Lead Managers"))
    B.append(kv("Registrar: ", _g(d, "registrar")))
    B.append(kv("Book Running Lead Manager: ", _g(d, "brlm")))
    B.append(kv("Contact Details: ", _g(d, "registered_office")))
    B.append(body(f"Email: {_g(d,'email')} | Phone: {_g(d,'phone')} | Website: {_g(d,'website')}"))

    # S16: Business Model
    B.append(h1(f"{co} Business Model"))
    B.append(body(_g(d, "business_model")))

    # S17: Growth Trajectory
    B.append(h1(f"{co} Growth Trajectory"))
    B.append(body(
        f"{co}'s Total Income for FY 26 was ₹ {_g(d,'total_income_fy26')} crores, "
        f"whereas in FY 25 and FY 24 it was ₹ {_g(d,'total_income_fy25')} crores "
        f"and ₹ {_g(d,'total_income_fy24')} crores, respectively."
    ))
    B.append(body(
        f"The Profit After Tax for FY 26 was ₹ {_g(d,'pat_fy26')} crores, "
        f"whereas in FY 25 and FY 24 it was ₹ {_g(d,'pat_fy25')} crores "
        f"and ₹ {_g(d,'pat_fy24')} crores, respectively."
    ))
    B.append(body(
        f"Their EBITDA for FY 26 was ₹ {_g(d,'ebitda_fy26')} crores, "
        f"whereas in FY 25 and FY 24 it was ₹ {_g(d,'ebitda_fy25')} crores "
        f"and ₹ {_g(d,'ebitda_fy24')} crores, respectively."
    ))

    # S18: Market Position
    B.append(h1(f"{co} Market Position"))
    B.append(body(_g(d, "about_para_3", f"{co} serves customers across India through its distribution network.")))
    B.append(body(
        f"As of 31 March 2026, the company's Total Income, Profit After Tax, and EBITDA were "
        f"₹ {_g(d,'total_income_fy26')} crores, ₹ {_g(d,'pat_fy26')} crores, "
        f"and ₹ {_g(d,'ebitda_fy26')} crores, respectively."
    ))

    # S19: P&L Table
    B.append(h1(f"{co} Profit and Loss Statement (in ₹ crores)"))
    B.append(tbl(
        [HD("Parameter"), HD("FY 26"), HD("FY 25"), HD("FY 24")],
        [C("Total Income"),      C(_g(d,"total_income_fy26")), C(_g(d,"total_income_fy25")), C(_g(d,"total_income_fy24"))],
        [C("Profit Before Tax"), C(_g(d,"pbt_fy26")),          C(_g(d,"pbt_fy25")),          C(_g(d,"pbt_fy24"))],
        [C("Profit After Tax"),  C(_g(d,"pat_fy26")),          C(_g(d,"pat_fy25")),          C(_g(d,"pat_fy24"))],
        [C("EPS (Diluted) ₹"),   C(_g(d,"eps_diluted_fy26")),  C(_g(d,"eps_diluted_fy25")),  C(_g(d,"eps_diluted_fy24"))],
        [C("EBITDA"),            C(_g(d,"ebitda_fy26")),       C(_g(d,"ebitda_fy25")),       C(_g(d,"ebitda_fy24"))],
    ))

    # S20: Balance Sheet / Cash Flows
    B.append(h1(f"{co} Balance Sheet (in ₹ crores)"))
    B.append(tbl(
        [HD("Parameter"), HD("FY 26"), HD("FY 25"), HD("FY 24")],
        [C("Profit Before Tax"),                  C(_g(d,"pbt_fy26")), C(_g(d,"pbt_fy25")), C(_g(d,"pbt_fy24"))],
        [C("Net Cash from Operating Activities"),  C("TBA"), C("TBA"), C("TBA")],
        [C("Net Cash from Investing Activities"),  C("TBA"), C("TBA"), C("TBA")],
        [C("Net Cash from Financing Activities"),  C("TBA"), C("TBA"), C("TBA")],
        [C("Cash & Cash Equivalents"),             C("TBA"), C("TBA"), C("TBA")],
    ))

    # S21: How to Apply
    B.append(h1(f"How to apply for {co} IPO?"))
    B.append(kv("Step 1: ", "Log in to your Kotak Neo Demat account to access IPO investments. Next, select the current IPO section."))
    B.append(kv("Step 2: ", "Specify IPO details. Enter the number of lots and the price you wish to apply for."))
    B.append(kv("Step 3: ", "Enter UPI ID. After entering your UPI ID, click submit. This will place your bid with the exchange."))
    B.append(kv("Step 4: ", "Mandate Notification. Your UPI app will receive a mandate notification to block funds."))
    B.append(kv("Step 5: ", "Approve Request. Your funds will be blocked once you approve the mandate request on your UPI.", last=True))

    # S22: FAQs
    B.append(h1("FAQs"))
    chair = _g(d, "chairperson_or_md")
    if "Managing Director" in chair:
        title_word = "Managing Director"
    elif any(x in chair for x in ("Chairperson", "Chairman")):
        title_word = "Chairperson"
    else:
        title_word = "Chief Executive Officer"
    name_only = re.split(r"\s*[\-,]|\s+(?:is|was)\s+", chair)[0].strip()
    B.append(kv(f"Who is the {title_word} of {co}? ", f"{name_only} is the {title_word} of {co}."))
    promoters = ", ".join(d.get("promoters") or ["TBA"])
    B.append(kv(f"Who are the Promoters of {co}? ", f"The Promoters of {co} are {promoters}.", last=True))

    # S23: Disclaimer
    B.append(h1("Disclaimer"))
    B.append(body(
        "This article is for informational purposes only and does not constitute financial advice. "
        "It is not produced by the desk of the Kotak Securities Research Team, nor is it a report "
        "published by the Kotak Securities Research Team. The information presented is compiled from "
        "several secondary sources available on the internet and may change over time. Investors should "
        "conduct their own research and consult with financial professionals before making any investment "
        "decisions. Read the full disclaimer here."
    ))
    B.append(body(
        "Investments in securities market are subject to market risks, read all the related documents "
        "carefully before investing. Brokerage will not exceed SEBI prescribed limit. The securities are "
        "quoted as an example and not as a recommendation. SEBI Registration No-INZ000200137 Member Id "
        "NSE-08081; BSE-673; MSE-1024, MCX-56285, NCDEX-1262."
    ))

    return B


# ── Document assembly ─────────────────────────────────────────────────────────

def build(d, template_path=None, output_dir=None):
    """
    Assemble the LP .docx from extracted data dict `d`.
    Returns the path of the written file.
    Uses the embedded Skytech template — no external file needed.
    template_path is accepted for backwards compatibility but ignored.
    """
    files = _get_template_files()

    sections = build_sections(d)
    body_xml  = "\n".join(sections)
    doc_xml   = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<w:document {NS}><w:body>{body_xml}{_SECT_PR}</w:body></w:document>'
    )
    files["word/document.xml"] = doc_xml.encode("utf-8")

    co = _g(d, "company_full_name")
    safe_name = re.sub(r"[^\w]", "", co.replace(" ", "_"))
    today     = datetime.now().strftime("%d%b%Y")
    filename  = f"{safe_name}_IPO_LP_DRHP_Draft_{today}.docx"

    # Candidate output directories in priority order
    candidates = []
    if output_dir:
        candidates.append(Path(output_dir) / filename)
    candidates += [
        DESKTOP / filename,
        Path.home() / filename,
        SAVE_DIR / filename,
    ]

    for dest in candidates:
        try:
            dest.parent.mkdir(parents=True, exist_ok=True)
            if dest.exists():
                dest.unlink()
            with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
                for name, data in files.items():
                    z.writestr(name, data)
            return str(dest)
        except Exception:
            continue

    raise RuntimeError(
        f"Could not write output file. Tried: {[str(c) for c in candidates]}\n"
        "Set IPO_DESKTOP env var to a writable directory."
    )
