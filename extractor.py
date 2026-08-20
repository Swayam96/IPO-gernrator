"""
Claude extraction via AWS Bedrock (credential_process / env vars, no boto3).
Handles SigV4 signing + structured DRHP data extraction in one place.
"""
import configparser
import hashlib
import hmac
import json
import os
import re
import ssl
import subprocess
from datetime import datetime, timezone
from urllib.request import Request, urlopen

# ── Bedrock config (override with env vars) ───────────────────────────────────
_REGION    = os.environ.get("AWS_REGION", "us-east-1")
_MODEL     = os.environ.get("BEDROCK_MODEL", "us.anthropic.claude-sonnet-4-6")
_PROFILE   = os.environ.get("AWS_PROFILE", "claude-code")
_MAX_TOK   = 8000

# Smart PDF sampling: front covers narrative/objects/risks, tail covers financials
PDF_FRONT_CHARS = 300_000
PDF_TAIL_CHARS  = 100_000


# ── AWS credential resolution ─────────────────────────────────────────────────

def _get_credentials():
    ak = os.environ.get("AWS_ACCESS_KEY_ID", "")
    sk = os.environ.get("AWS_SECRET_ACCESS_KEY", "")
    st = os.environ.get("AWS_SESSION_TOKEN", "")
    if ak:
        return ak, sk, st
    cfg = configparser.ConfigParser()
    cfg.read(os.path.expanduser("~/.aws/config"))
    section = f"profile {_PROFILE}" if f"profile {_PROFILE}" in cfg else _PROFILE
    if section not in cfg or "credential_process" not in cfg[section]:
        raise RuntimeError(
            f"No AWS credentials. Set AWS_ACCESS_KEY_ID or configure "
            f"credential_process under [{section}] in ~/.aws/config."
        )
    result = subprocess.run(
        cfg[section]["credential_process"], shell=True,
        capture_output=True, text=True, timeout=30,
    )
    c = json.loads(result.stdout)
    return c.get("AccessKeyId", ""), c.get("SecretAccessKey", ""), c.get("SessionToken", "")


# ── SigV4 signing ─────────────────────────────────────────────────────────────

def _sign_and_call(body_bytes):
    ak, sk, st = _get_credentials()
    now    = datetime.now(timezone.utc)
    date_s = now.strftime("%Y%m%d")
    time_s = now.strftime("%Y%m%dT%H%M%SZ")
    host   = f"bedrock-runtime.{_REGION}.amazonaws.com"
    path   = f"/model/{_MODEL}/invoke"

    def _sha256(b): return hashlib.sha256(b).hexdigest()
    def _hmac(key, msg):
        k = key if isinstance(key, bytes) else key.encode()
        return hmac.new(k, msg.encode(), hashlib.sha256).digest()

    hdrs = {"content-type": "application/json", "host": host, "x-amz-date": time_s}
    if st:
        hdrs["x-amz-security-token"] = st

    signed  = ";".join(sorted(hdrs))
    canon_h = "".join(f"{k}:{v}\n" for k, v in sorted(hdrs.items()))
    scope   = f"{date_s}/{_REGION}/bedrock/aws4_request"
    canon   = "\n".join(["POST", path, "", canon_h, signed, _sha256(body_bytes)])
    s2s     = "\n".join(["AWS4-HMAC-SHA256", time_s, scope, _sha256(canon.encode())])
    sig_key = _hmac(_hmac(_hmac(_hmac(f"AWS4{sk}", date_s), _REGION), "bedrock"), "aws4_request")
    sig     = hmac.new(sig_key, s2s.encode(), hashlib.sha256).hexdigest()

    hdrs["authorization"] = (
        f"AWS4-HMAC-SHA256 Credential={ak}/{scope}, "
        f"SignedHeaders={signed}, Signature={sig}"
    )
    hdrs["content-length"] = str(len(body_bytes))

    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE  # corporate proxy — no public CA chain

    req = Request(f"https://{host}{path}", data=body_bytes, method="POST")
    for k, v in hdrs.items():
        req.add_header(k, v)
    with urlopen(req, context=ssl_ctx, timeout=120) as resp:
        return json.loads(resp.read())


def _call_claude(user_prompt, system_prompt):
    body = json.dumps({
        "anthropic_version": "bedrock-2023-05-31",
        "max_tokens": _MAX_TOK,
        "system": system_prompt,
        "messages": [{"role": "user", "content": user_prompt}],
    }).encode("utf-8")
    return _sign_and_call(body)["content"][0]["text"]


# ── Extraction prompt ─────────────────────────────────────────────────────────

_SYSTEM = (
    "You are an expert financial content writer for Kotak Securities specialising in "
    "IPO Landing Page (LP) documents. Extract structured data from DRHP text and return "
    "valid JSON only — no markdown fences, no explanation."
)

_PROMPT = """Extract all key data fields from the DRHP text below to write a 23-section IPO LP document.

RULES:
- Extract ONLY what is explicitly stated in the DRHP text.
- If a field is missing (e.g. price band, dates), set it to "TBA".
- All monetary figures must be in ₹ crores (convert: ₹ millions ÷ 10, ₹ lakhs ÷ 100).
- Verify the unit header of each financial table before extracting — state the unit in "financials_unit".
- Use ONLY post-issue diluted EPS from restated financials.
- Write all narrative fields in third-person (convert "our company / we / our" → "the Company / its").
- EBITDA = PBT + Finance Costs + Depreciation & Amortisation.

Return a JSON object with these exact keys:

{{
  "company_full_name": "exact legal name from cover page",
  "drhp_date": "e.g. July 28, 2026",
  "doc_type": "DRHP",
  "ipo_type": "SME or Mainboard",
  "listing_exchange": "BSE and NSE / BSE SME / NSE Emerge",
  "face_value": "e.g. ₹10 per share",
  "fresh_issue_shares": "e.g. shares aggregating up to ₹125 crore",
  "ofs_details": "e.g. up to X shares by [Seller 1] (X shares) and [Seller 2] (X shares)",
  "promoters": ["Name 1", "Name 2"],
  "brlm": "Full legal name(s)",
  "registrar": "Full legal name",
  "company_secretary": "Name",
  "chairperson_or_md": "Name and exact title (e.g. Managing Director)",
  "registered_office": "full address",
  "website": "url",
  "phone": "phone number",
  "email": "email",
  "cin": "CIN number",
  "incorporation_year": "year",

  "business_description": "2-3 sentences, third-person, specific facts. No AI filler phrases.",
  "business_model": "Opening sentence: The company earns its revenue through [description]. Then 2-3 sentences on revenue streams, channels, key customers.",

  "industry_para_1": "Sector definition, current market size with specific figures from DRHP.",
  "industry_para_2": "Key growth drivers with specific numbers from DRHP.",
  "industry_para_3": "Company's positioning to benefit, with figures.",

  "about_para_1": "Incorporation, CIN, HQ, what it does. Third-person.",
  "about_para_2": "Products/services, key certifications, manufacturing. Specific facts.",
  "about_para_3": "Distribution, customers, operational metrics. End with one financial figure.",

  "strengths": [
    "Strength 1 — ONE sentence, factual, specific number if available. No 'we believe'.",
    "Strength 2", "Strength 3", "Strength 4", "Strength 5"
  ],
  "risks": [
    "Risk 1 — ONE sentence: (a) what could go wrong and (b) why it matters.",
    "Risk 2", "Risk 3", "Risk 4", "Risk 5"
  ],

  "financials_unit": "e.g. in ₹ millions",
  "revenue_fy26": "value in ₹ crores",
  "revenue_fy25": "value in ₹ crores",
  "revenue_fy24": "value in ₹ crores",
  "total_income_fy26": "value in ₹ crores",
  "total_income_fy25": "value in ₹ crores",
  "total_income_fy24": "value in ₹ crores",
  "pbt_fy26": "value in ₹ crores",
  "pbt_fy25": "value in ₹ crores",
  "pbt_fy24": "value in ₹ crores",
  "pat_fy26": "value in ₹ crores",
  "pat_fy25": "value in ₹ crores",
  "pat_fy24": "value in ₹ crores",
  "ebitda_fy26": "calculated value in ₹ crores",
  "ebitda_fy25": "calculated value in ₹ crores",
  "ebitda_fy24": "calculated value in ₹ crores",
  "eps_diluted_fy26": "post-issue diluted EPS in ₹",
  "eps_diluted_fy25": "TBA if not in DRHP",
  "eps_diluted_fy24": "TBA if not in DRHP",
  "ronw_fy26": "e.g. 75.90%",
  "ronw_fy25": "e.g. 65.10%",
  "ronw_fy24": "e.g. 100.98%",
  "total_equity_fy26": "TBA or value in ₹ crores",
  "total_equity_fy25": "TBA or value in ₹ crores",
  "total_equity_fy24": "TBA or value in ₹ crores",

  "objects": ["Object 1 with ₹ amount if stated", "Object 2", "Object 3"],

  "peer_comparison_note": "Either 'The RHP does not provide a peer comparison as there are no publicly listed companies in India with an exclusively similar business model.' OR list peers as described in DRHP.",
  "peers": [
    {{"name": "Company Name", "revenue": "₹X crore", "pe": "X.Xx", "eps": "₹X.XX", "nav": "₹X.XX"}}
  ]
}}

DRHP TEXT:
{drhp_text}
"""


def extract(pdf_text):
    """Send DRHP text to Claude and return the structured data dict."""
    raw = _call_claude(_PROMPT.format(drhp_text=pdf_text), _SYSTEM)
    raw = re.sub(r"^```(?:json)?\s*", "", raw.strip())
    raw = re.sub(r"\s*```$", "", raw.strip())
    return json.loads(raw)
