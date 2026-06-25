import os
import re
import csv
import sys
import time
import json
import base64
import hashlib
import logging
from datetime import datetime, timedelta
from urllib.parse import urljoin, urlparse, urlsplit, urlunsplit, parse_qsl, urlencode

import requests
from bs4 import BeautifulSoup

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import pandas as pd
    import openpyxl
    _XLSX_AVAILABLE = True
except ImportError:
    _XLSX_AVAILABLE = False

try:
    import language_tool_python
    from sentence_transformers import SentenceTransformer, util as st_util
    _NLP_AVAILABLE = True
except ImportError:
    _NLP_AVAILABLE = False

# =============================================================================
#  CONFIG
# =============================================================================

BASE_URL = "https://jobs.newtimes.co.rw"

# Category to scrape. Override with env var NEWTIMES_CATEGORY=tenders|all
SCRAPE_CATEGORY = os.environ.get("NEWTIMES_CATEGORY", "announcement")

# PATCH A: archive uses ?q=&page=N format, category as &category= param
# The /jobs/search path is the SPA root — no sub-path segments for category.
def _archive_url() -> str:
    return f"{BASE_URL}/jobs/search"

JOBS_ARCHIVE_URL = _archive_url()

# JSON API — try both the standard endpoint and the category-specific one.
_NEWTIMES_API_URL = f"{BASE_URL}/api/jobs"
_NEWTIMES_API_ALT = f"{BASE_URL}/api/announcements"   # PATCH F: alternate endpoint

REQUIRE_PUBLIC_APPLY = os.environ.get("REQUIRE_PUBLIC_APPLY", "1") != "0"

REQUEST_DELAY   = float(os.environ.get("REQUEST_DELAY", "1.5"))
MAX_JOBS        = int(os.environ.get("MAX_JOBS", "0"))
MAX_PAGES       = int(os.environ.get("MAX_PAGES", "20"))
REQUEST_TIMEOUT = int(os.environ.get("REQUEST_TIMEOUT", "25"))

OUTPUT_FILE        = "newtimes_rwanda_jobs.xlsx"
PROCESSED_IDS_FILE = "newtimes_rwanda_processed.csv"
FLAGGED_FILE       = "newtimes_rwanda_flagged.csv"

_TRACKER_FIELDS = ["Job ID", "Job URL", "Job Title", "Company Name",
                   "Status", "Timestamp", "WP ID"]

_FLAGGED_FIELDS = ["Source", "Title", "Company", "Location", "Salary",
                   "Deadline", "Reason", "Apply Note", "Job URL", "Timestamp"]

# ── WordPress ─────────────────────────────────────────────────────────────────
WP_URL      = os.environ.get("WP_BASE_URL", "")
WP_USER     = os.environ.get("WP_USERNAME", "")
WP_PASSWORD = os.environ.get("WP_APP_PASSWORD", "")
WP_BASE      = WP_URL.rstrip("/")
WP_JOBS_URL  = f"{WP_BASE}/job-listings"
WP_MEDIA_URL = f"{WP_BASE}/media"

# ── Mistral ───────────────────────────────────────────────────────────────────
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY", "")
MISTRAL_MODEL   = "mistral-small-latest"
MISTRAL_URL     = "https://api.mistral.ai/v1/chat/completions"

ENABLE_PARAPHRASE = True

for _var, _val, _feature in [
    ("MISTRAL_API_KEY", MISTRAL_API_KEY, "paraphrasing"),
    ("WP_USERNAME",     WP_USER,         "WordPress posting"),
    ("WP_APP_PASSWORD", WP_PASSWORD,     "WordPress posting"),
]:
    if not _val:
        logging.getLogger(__name__).warning(
            f"Environment variable {_var} is not set — {_feature} will be disabled/skipped."
        )

JOB_TYPE_MAPPING = {
    "full-time": "full-time", "full time": "full-time",
    "part-time": "part-time", "part time": "part-time",
    "contract":  "contract",  "temporary": "temporary",
    "internship":"internship","freelance": "freelance",
    "volunteer": "volunteer", "permanent": "full-time",
    "invitation for bid": "contract",
    "tenders":   "contract",  "tender":   "contract",
}

# PATCH B: stronger browser fingerprint headers to avoid 403
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
    "Referer": BASE_URL + "/",
    "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "same-origin",
    "Sec-Fetch-User": "?1",
    "Upgrade-Insecure-Requests": "1",
    "Connection": "keep-alive",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# PATCH C: prime cookies by visiting the homepage before hitting any archive page
_session_primed = False

def _prime_session():
    """Visit the homepage once to acquire session cookies, avoiding 403 on archive."""
    global _session_primed
    if _session_primed:
        return
    try:
        r = SESSION.get(BASE_URL + "/", timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        _session_primed = True
        log_.info(f"Session primed — cookies: {dict(SESSION.cookies)}")
    except Exception as e:
        log_.warning(f"Session priming failed (continuing anyway): {e}")
        _session_primed = True   # don't retry on every call

RWANDA_LOCATIONS = [
    "Kigali", "Gasabo", "Kicukiro", "Nyarugenge",
    "Musanze", "Rubavu", "Ngororero", "Karongi", "Rutsiro",
    "Rusizi", "Nyamasheke", "Ruhango", "Muhanga", "Kamonyi",
    "Huye", "Nyanza", "Gisagara", "Nyaruguru", "Nyamagabe",
    "Rwamagana", "Kayonza", "Kirehe", "Ngoma", "Bugesera",
    "Nyagatare", "Gatsibo", "Rulindo", "Gakenke",
    "Burera", "Gicumbi",
    "Kigali Province", "Southern Province", "Northern Province",
    "Eastern Province", "Western Province",
]
DEFAULT_LOCATION = os.environ.get("NEWTIMES_DEFAULT_LOCATION", "Kigali, Rwanda")

_NON_APPLY_EMAIL_DOMAINS = ("newtimes.co.rw", "jobs.newtimes.co.rw")
_NON_APPLY_HOST_SUBSTR = (
    "newtimes.co.rw", "facebook.", "twitter.", "x.com", "linkedin.",
    "instagram.", "wa.me", "whatsapp", "t.me", "telegram",
    "plus.google", "pinterest.", "youtube.",
)
_NON_APPLY_PATH_SUBSTR = (
    "/member-", "action=login", "mode=register", "#share", "/share",
    "/wp-login", "/cart", "/checkout", "/POSTAJOB", "/employers",
    "/blog", "/jobs/search", "/jobs/map",
)

def _is_real_apply_email(email: str) -> bool:
    if not email or "@" not in email:
        return False
    dom = email.rsplit("@", 1)[-1].lower()
    return not any(dom == d or dom.endswith("." + d) for d in _NON_APPLY_EMAIL_DOMAINS)

# =============================================================================
#  LOGGING / COLOUR
# =============================================================================

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log_ = logging.getLogger(__name__)

_USE_COLOUR = sys.stdout.isatty()

def _c(code, text):
    return f"\033[{code}m{text}\033[0m" if _USE_COLOUR else text

C_HEADER  = lambda t: _c("1;36",  t)
C_LABEL   = lambda t: _c("1;33",  t)
C_VALUE   = lambda t: _c("97",    t)
C_DIM     = lambda t: _c("2",     t)
C_GREEN   = lambda t: _c("1;32",  t)
C_RED     = lambda t: _c("1;31",  t)
C_BLUE    = lambda t: _c("1;34",  t)
C_DIVIDER = lambda: _c("2", "─" * 80)

def log(msg):
    print(msg, flush=True)

EMAIL_PATTERN = re.compile(r"[A-Za-z0-9.+_-]+@[A-Za-z0-9-]+\.[A-Za-z0-9.-]+")
URL_PATTERN   = re.compile(r"https?://[^\s)>\"']+", re.I)

TRACKING_PARAM_PREFIXES = ("utm_",)
TRACKING_PARAM_EXACT = {
    "fbclid", "gclid", "msclkid", "mc_cid", "mc_eid", "ref", "referrer",
}

MONTHS = {
    "january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6,
    "july": 7, "august": 8, "september": 9, "october": 10, "november": 11,
    "december": 12, "jan": 1, "feb": 2, "mar": 3, "apr": 4, "jun": 6, "jul": 7,
    "aug": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dec": 12,
}

_META_DMY_RE = re.compile(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{4})\b")

TEXT_DATE_RE = re.compile(
    r"(\d{1,2})\s*(?:st|nd|rd|th)?\s+([A-Za-z]+)\s*[.,]?\s*(\d{4})", re.I
)

DEADLINE_LABELS = ("application deadline", "closing date", "deadline",
                   "expiry date", "expires", "submission deadline",
                   "deadline for applications", "closing date and time",
                   "date de clôture")

_APPLY_HEAD_PHRASES = re.compile(
    r"^(?:how\s*(?:and|&)\s*deadline\s*to\s*apply|how\s*to\s*apply(?:\s*(?:and|&)\s*deadline)?|"
    r"how\s*to\s*submit|to\s*apply|application\s*(?:and|&)\s*deadline|"
    r"mode\s*of\s*application|method\s*of\s*application|"
    r"application\s*(?:procedure|process|instructions?|method|guidelines?)|"
    r"submission\s*of\s*applications?|deadline\s*(?:and|&)?\s*(?:how\s*)?to\s*apply|"
    r"interested\s+candidates?|how\s+to\s+(?:bid|submit|tender)|"
    r"interested\s+companies|interested\s+parties|interested\s+suppliers?)\b",
    re.I,
)

_BODY_CUT_MARKERS = [
    "related jobs", "leave your thoughts", "you must be logged in",
    "email me jobs like these", "send to a friend", "company information",
    "leave a reply", "post a comment", "post a job", "log in",
    "create employer account", "employers ·", "about · contact",
]
_BODY_DROP_LINES = {
    "apply for this job", "save", "share", "share:", "bookmark job",
    "quick view", "send to friend", "send to a friend", "clear all",
    "filter", "view more", "post a job", "employers", "log in",
    "create employer account",
}

# =============================================================================
#  TEXT CLEANUP / SANITIZATION
# =============================================================================

_MOJIBAKE = [
    ("Â", ""), ("â€™", "'"), ("â€œ", '"'), ("â€\x9d", '"'), ("â€", '"'),
    ("â€¢", "•"), ("â„¢", "™"), ("\u00a0", " "), ("\u200b", ""), ("\ufeff", ""),
]

def _fix_mojibake(text: str) -> str:
    for pattern, replacement in _MOJIBAKE:
        text = text.replace(pattern, replacement)
    text = re.sub(r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]", "", text)
    return text

def sanitize_text(text, is_url=False) -> str:
    if not isinstance(text, str):
        text = str(text) if (text is not None and str(text) not in ("nan", "None", "NaN")) else ""
    text = text.strip()
    if text in ("nan", "None", "NaN", "", "N/A", "n/a", "NA", "na"):
        return ""
    text = _fix_mojibake(text)
    if is_url:
        return re.sub(r"[ \t\r\n\f\v]+", " ", text).strip()
    text = re.sub(r"#+\s*", "", text)
    text = re.sub(r"\*\*", "", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()

def clean_text(el):
    if el is None:
        return ""
    return re.sub(r"\s+", " ", el.get_text(" ", strip=True)).strip()

def extract_email(text):
    if not text:
        return ""
    m = EMAIL_PATTERN.search(text)
    return m.group(0) if m else ""

def strip_tracking_params(url):
    if not url:
        return url
    parts = urlsplit(url)
    if not parts.query:
        return url
    kept = []
    for key, value in parse_qsl(parts.query, keep_blank_values=True):
        key_lower = key.lower()
        if key_lower.startswith(TRACKING_PARAM_PREFIXES) or key_lower in TRACKING_PARAM_EXACT:
            continue
        kept.append((key, value))
    new_query = urlencode(kept)
    return urlunsplit((parts.scheme, parts.netloc, parts.path, new_query, parts.fragment))

# =============================================================================
#  BASIC HTTP / PARSING HELPERS
# =============================================================================

def get_soup(url: str) -> BeautifulSoup:
    _prime_session()
    resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    try:
        return BeautifulSoup(resp.text, "lxml")
    except Exception:
        return BeautifulSoup(resp.text, "html.parser")

def get_raw(url: str) -> str:
    """Return raw response text (for __NEXT_DATA__ extraction)."""
    _prime_session()
    resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    resp.encoding = resp.apparent_encoding or "utf-8"
    return resp.text

def get_json(url: str, params: dict = None) -> dict | None:
    """Attempt a JSON API call; return None on any failure."""
    _prime_session()
    try:
        hdrs = dict(SESSION.headers)
        hdrs["Accept"] = "application/json, text/javascript, */*"
        hdrs["Sec-Fetch-Dest"] = "empty"
        hdrs["Sec-Fetch-Mode"] = "cors"
        hdrs["X-Requested-With"] = "XMLHttpRequest"
        resp = SESSION.get(url, params=params, headers=hdrs, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        log_.debug(f"JSON API call failed ({url}): {e}")
        return None

def slugify(text, maxlen=80):
    s = re.sub(r"[^a-z0-9]+", "-", (text or "").lower()).strip("-")
    return s[:maxlen] or "job"

def html_block_to_text(el) -> str:
    if el is None:
        return ""
    for br in el.find_all("br"):
        br.replace_with("\n")
    for li in el.find_all("li"):
        txt = li.get_text(" ", strip=True)
        li.replace_with("\n• " + txt + "\n")
    for tag in el.find_all(["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "tr"]):
        tag.insert_before("\n")
        tag.insert_after("\n")
    text = el.get_text("\n")
    text = _fix_mojibake(text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

# =============================================================================
#  DATE / FIELD EXTRACTORS
# =============================================================================

def dmy_dates(text: str) -> list:
    out = []
    for d, m, y in _META_DMY_RE.findall(text or ""):
        try:
            out.append(datetime(int(y), int(m), int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def text_dates(text: str) -> list:
    out = []
    for d, mon, y in TEXT_DATE_RE.findall(text or ""):
        month = MONTHS.get(mon.lower())
        if not month:
            continue
        try:
            out.append(datetime(int(y), month, int(d)).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    return out

def parse_any_date(text: str) -> str:
    ds = dmy_dates(text)
    if ds:
        return ds[-1]
    ts = text_dates(text)
    return ts[-1] if ts else ""

def clean_title(raw: str) -> str:
    t = sanitize_text(raw)
    t = re.sub(r"\s*\|?\s*[Pp]ublished\s+on\s+.*", "", t)
    t = re.sub(r"\s*\|?\s*[Dd]eadline\s+.*", "", t)
    t = re.sub(r"\s*[–—-]\s*[Tt]he\s+[Nn]ew\s+[Tt]imes.*", "", t)
    t = re.sub(r"\s*\|\s*Rwanda jobs.*", "", t, flags=re.I)
    t = re.sub(r"\s+[Jj]ob\s+at\s+.+$", "", t)
    return t.strip()

def map_job_type(raw: str) -> str:
    key = (raw or "").lower().strip()
    return JOB_TYPE_MAPPING.get(key, "full-time")

def location_from_text(text: str) -> str:
    if text:
        for town in RWANDA_LOCATIONS:
            if re.search(rf"\b{re.escape(town)}\b", text, re.I):
                return town
    return DEFAULT_LOCATION

def extract_experience(text: str) -> str:
    if not text:
        return ""
    m = re.search(r"(?:at least|minimum(?: of)?)\s+\d+\s+years?[^.\n;]*", text, re.I)
    if m:
        return m.group(0).strip().rstrip(".")
    m = re.search(r"\b\d+\s+years?[^.\n;]*experience", text, re.I)
    if m:
        return m.group(0).strip().rstrip(".")
    return ""

def extract_salary(text: str) -> str:
    if not text:
        return ""
    m = re.search(
        r"(?:RWF|Rwf|FRW|RW\s*F)\s*([0-9]{1,3}(?:[,.\s][0-9]{3})*(?:\.[0-9]+)?)", text
    )
    if m:
        amt = re.sub(r"[\s,]", "", m.group(1))
        return f"RWF {amt}"
    m = re.search(r"\b(?:salary|remuneration|pay)\b[^.\n]{0,80}", text, re.I)
    if m and re.search(r"\d", m.group(0)):
        return m.group(0).strip().rstrip(".")
    return ""

# =============================================================================
#  CANONICAL NORMALISERS
# =============================================================================

def _kw_hit(text_low: str, keywords) -> bool:
    for k in keywords:
        kk = k.strip().lower()
        if not kk:
            continue
        esc = re.escape(kk)
        if len(kk) <= 3:
            pat = r"(?<![a-z0-9])" + esc + r"(?![a-z0-9])"
        else:
            pat = r"(?<![a-z0-9])" + esc + r"(?:es|s)?(?![a-z0-9])"
        if re.search(pat, text_low):
            return True
    return False

QUALIFICATION_TIERS = [
    ("PhD / Doctorate",          ["phd", "ph.d", "doctorate", "doctoral", "doctor of philosophy"]),
    ("Master's Degree",          ["master", "msc", "m.sc", "ma ", "m.a ", "mba", "m.b.a", "meng",
                                  "m.eng", "mphil", "postgraduate", "post-graduate", "post graduate"]),
    ("Bachelor's Degree",        ["bachelor", "bsc", "b.sc", "ba ", "b.a ", "beng", "b.eng", "bcom",
                                  "b.com", "bba", "llb", "degree in", "undergraduate degree",
                                  "honours degree", "hons"]),
    ("Higher National Diploma",  ["hnd", "hnc", "higher national diploma", "higher national certificate",
                                  "higher diploma", "advanced diploma"]),
    ("Diploma",                  ["diploma", "dip ", "dip.", "associate degree", "foundation degree"]),
    ("Professional Certification", ["acca", "cpa", "cfa", "cima", "pmp", "prince2", "cissp",
                                    "aws certified", "comptia", "cisco", "ccna", "ccnp", "shrm",
                                    "cipd", "chartered", "certified public", "certified financial",
                                    "certified project", "professional certification",
                                    "professional certificate"]),
    ("A-Levels / HSC",           ["a-level", "a level", "hsc", "higher school certificate", "ib diploma",
                                  "international baccalaureate", "gce advanced"]),
    ("O-Levels / School Certificate", ["o-level", "o level", "igcse", "gcse", "school certificate",
                                       "sc ", "cpe", "certificate of primary"]),
    ("No Formal Qualification Required", ["no qualification", "no degree", "no formal", "school leaver",
                                          "entry level", "no experience required", "training provided",
                                          "will train"]),
]

def extract_qualification(text: str) -> str:
    if not text:
        return ""
    if re.search(r"nursery|primary years|ib pyp|aged between|boys and girls", text, re.I):
        return ""
    lower = text.lower()
    for label, keywords in QUALIFICATION_TIERS:
        if _kw_hit(lower, keywords):
            return label
    return ""

NO_EXP_KW = ["no experience", "no prior experience", "fresh graduate", "freshers",
             "entry level", "entry-level", "0 years", "zero experience",
             "training provided", "will train", "no experience required"]
LESS1_KW  = ["less than 1 year", "under 1 year", "6 months", "less than a year",
             "some experience", "minimal experience"]

def years_to_band(n: int) -> str:
    if n <= 0:  return "No Experience Required"
    if n <= 2:  return "1 - 2 Years"
    if n <= 5:  return "3 - 5 Years"
    if n <= 10: return "6 - 10 Years"
    return "10+ Years"

_EXP_CAP = 20
_EXP_REQ_RE = re.compile(
    r"(?:minimum|min\.?|at\s+least|atleast|least|over|more\s+than|not\s+less\s+than|"
    r"minimum\s+of|a\s+minimum\s+of)\s+(?:of\s+)?(\d{1,2})\s*\+?\s*years?", re.I)
_EXP_YEARS_OF_RE = re.compile(r"(\d{1,2})\s*\+?\s*years?\s+of\b", re.I)
_EXP_ANY_YEARS_RE = re.compile(r"(\d{1,2})\s*\+?\s*years?", re.I)
_EXP_RANGE_RE = re.compile(r"(\d{1,2})\s*(?:-|–|to)\s*(\d{1,2})\s*years?", re.I)

def extract_experience_band(text: str) -> str:
    if not text:
        return ""
    low = text.lower()
    years = []
    for m in _EXP_REQ_RE.finditer(text):
        n = int(m.group(1))
        if 0 < n <= _EXP_CAP:
            years.append(n)
    for m in _EXP_YEARS_OF_RE.finditer(low):
        n = int(m.group(1))
        if 0 < n <= _EXP_CAP:
            years.append(n)
    for m in _EXP_ANY_YEARS_RE.finditer(low):
        n = int(m.group(1))
        if 0 < n <= _EXP_CAP and "experien" in low[m.end():m.end() + 60]:
            years.append(n)
    for m in _EXP_RANGE_RE.finditer(text):
        a = int(m.group(1))
        if 0 < a <= _EXP_CAP:
            years.append(a)
    if years:
        return years_to_band(min(years))
    if _kw_hit(low, NO_EXP_KW):
        return "No Experience Required"
    if _kw_hit(low, LESS1_KW):
        return "1 - 2 Years"
    return ""

FIELD_KEYWORD_MAP = [
    ("Information Technology",
     ["software engineer", "developer", "devops", "frontend", "backend", "full stack", "fullstack",
      "sysadmin", "cloud", "cybersecurity", "data engineer", "machine learning", "artificial intelligence",
      "ai/ml", "it support", "network engineer", "database", "kubernetes", "docker", "aws", "azure",
      "react", "node.js", "python developer", "java developer"],
     ["programming", "coding", "api", "agile", "scrum", "git", "linux", "server", "infrastructure", "software"]),
    ("Finance & Accounting",
     ["accountant", "auditor", "finance manager", "financial analyst", "cfo", "treasurer", "tax",
      "bookkeeper", "payroll", "budget analyst", "credit analyst", "investment", "portfolio manager",
      "risk analyst", "forex", "actuary", "acca", "cfa", "cpa"],
     ["financial", "accounting", "balance sheet", "p&l", "reconciliation", "ifrs", "gaap", "ledger", "invoicing"]),
    ("Sales & Business Development",
     ["sales executive", "sales manager", "business development", "account manager",
      "sales representative", "bd manager", "regional sales", "key account", "sales director",
      "commercial manager", "sales officer"],
     ["revenue", "pipeline", "crm", "leads", "prospects", "quota", "target", "upsell", "cross-sell", "b2b", "b2c"]),
    ("Marketing & Communications",
     ["marketing manager", "digital marketing", "seo", "sem", "content marketer", "social media manager",
      "brand manager", "marketing executive", "communications manager", "pr manager", "copywriter",
      "growth hacker", "email marketing", "campaign manager"],
     ["marketing", "branding", "advertising", "social media", "content", "campaign", "analytics"]),
    ("Human Resources",
     ["hr manager", "human resources", "recruiter", "talent acquisition", "hr business partner",
      "hrbp", "hr officer", "compensation", "benefits manager", "organisational development",
      "learning and development", "l&d", "hr generalist", "payroll manager"],
     ["recruitment", "onboarding", "performance management", "employee relations", "hr", "workforce"]),
    ("Engineering",
     ["mechanical engineer", "civil engineer", "electrical engineer", "structural engineer",
      "process engineer", "project engineer", "maintenance engineer", "production engineer",
      "quality engineer", "safety engineer", "site engineer", "design engineer"],
     ["engineering", "cad", "autocad", "solidworks", "manufacturing", "plant", "machinery", "commissioning"]),
    ("Healthcare & Medicine",
     ["doctor", "physician", "nurse", "pharmacist", "medical officer", "surgeon", "anaesthetist",
      "physiotherapist", "radiographer", "lab technician", "clinical", "healthcare manager",
      "occupational therapist", "dentist", "midwife"],
     ["hospital", "clinic", "patient", "medical", "health", "pharmaceutical", "diagnosis", "treatment"]),
    ("Education & Training",
     ["teacher", "lecturer", "professor", "trainer", "educator", "tutor", "school principal",
      "academic", "curriculum", "e-learning", "instructional designer", "teaching assistant"],
     ["school", "university", "college", "classroom", "students", "pedagogy", "curriculum", "education"]),
    ("Hospitality & Tourism",
     ["hotel manager", "front desk", "housekeeping", "chef", "sous chef", "food and beverage",
      "f&b manager", "restaurant manager", "bartender", "waiter", "concierge", "tour guide",
      "travel agent", "events coordinator", "catering"],
     ["hospitality", "hotel", "resort", "tourism", "guest", "accommodation", "restaurant", "kitchen"]),
    ("Logistics & Supply Chain",
     ["supply chain manager", "logistics coordinator", "warehouse manager", "fleet manager",
      "procurement manager", "purchasing manager", "import export", "freight", "shipping coordinator",
      "inventory manager", "demand planner"],
     ["logistics", "supply chain", "warehouse", "inventory", "freight", "procurement", "sourcing"]),
    ("Legal",
     ["lawyer", "attorney", "legal counsel", "paralegal", "compliance officer", "legal advisor",
      "solicitor", "barrister", "corporate counsel", "legal manager", "contract manager"],
     ["legal", "law", "contracts", "litigation", "regulatory", "compliance", "gdpr"]),
    ("Administration & Operations",
     ["office manager", "executive assistant", "administrative officer", "operations manager",
      "pa", "personal assistant", "receptionist", "data entry", "office administrator",
      "company secretary", "business analyst"],
     ["administration", "operations", "office", "coordination", "scheduling", "reporting", "clerical"]),
    ("Customer Service",
     ["customer service", "call centre", "customer success", "customer support", "help desk",
      "service advisor", "client relations", "customer experience", "contact centre"],
     ["customer", "support", "helpdesk", "tickets", "escalation", "satisfaction", "service level"]),
    ("Construction & Real Estate",
     ["quantity surveyor", "site supervisor", "project manager construction", "architect",
      "draughtsman", "property manager", "estate agent", "real estate", "building inspector",
      "land surveyor", "construction manager"],
     ["construction", "building", "property", "real estate", "site", "contractor", "tender"]),
    ("Manufacturing & Production",
     ["production manager", "quality control", "quality assurance", "qa", "qc", "factory manager",
      "plant manager", "production supervisor", "assembly", "cnc operator", "technician"],
     ["production", "manufacturing", "factory", "assembly", "quality", "lean", "six sigma"]),
    ("Design & Creative",
     ["graphic designer", "ui/ux", "product designer", "art director", "creative director",
      "animator", "illustrator", "photographer", "videographer", "motion designer", "web designer"],
     ["design", "creative", "adobe", "figma", "photoshop", "illustrator", "indesign", "sketch", "branding"]),
    ("Research & Science",
     ["research scientist", "data scientist", "lab researcher", "research analyst",
      "clinical researcher", "environmental scientist", "chemist", "biologist", "statistician"],
     ["research", "analysis", "data", "laboratory", "science", "experiment", "findings", "methodology"]),
    ("Security",
     ["security officer", "security guard", "security manager", "cctv", "loss prevention",
      "risk manager", "health and safety", "hse officer", "osh", "fire safety"],
     ["security", "safety", "risk", "surveillance", "patrol", "access control", "emergency"]),
    ("Media & Journalism",
     ["journalist", "editor", "reporter", "broadcast", "news anchor", "content creator",
      "media manager", "radio", "television", "producer", "scriptwriter"],
     ["media", "journalism", "broadcast", "news", "editorial", "publishing", "press"]),
    ("Non-Profit & Social Work",
     ["social worker", "ngo", "charity", "programme coordinator", "community development",
      "welfare officer", "case manager", "development officer", "fundraiser", "volunteer coordinator"],
     ["social", "ngo", "community", "welfare", "beneficiary", "donor", "impact", "charity"]),
]

_TENDER_TITLE_RE = re.compile(
    r"\b(?:rfq|rfp|reoi|eoi|itb|itt|spn|rfb|rfa|gpn|ifb|rfi)\b"
    r"|invitation\s+to\s+(?:bid|tender)|invitation\s+for\s+bids?"
    r"|request\s+for\s+(?:quotation|proposal|proposals|expression|expressions|bids?)"
    r"|expressions?\s+of\s+interest"
    r"|\btenders?\s+(?:notice|document|no\b)|procurement\s+notice"
    r"|specific\s+procurement|general\s+procurement"
    r"|call\s+for\s+(?:bid|bids|tender|tenders|proposal|proposals|expression|expressions|quotation)"
    r"|matching\s+grant|terms\s+of\s+reference|prior\s+notice\s+of\s+procurement"
    r"|itangazo\s+ry.isoko",
    re.I,
)
TENDER_FIELD = "Public Notices & Tenders"

def infer_field(title: str, description: str, fallback_categories: str = "") -> str:
    title_l = (title or "").lower()
    if _TENDER_TITLE_RE.search(title_l):
        return TENDER_FIELD
    text = f"{title}\n{description}".lower()
    for field, strong, _weak in FIELD_KEYWORD_MAP:
        if _kw_hit(text, strong):
            return field
    for field, _strong, weak in FIELD_KEYWORD_MAP:
        if _kw_hit(text, weak):
            return field
    if fallback_categories:
        cats = [c.strip() for c in fallback_categories.split(",") if c.strip()]
        for c in cats:
            if "tender" in c.lower() or "notice" in c.lower():
                return TENDER_FIELD
        if cats:
            return cats[0]
    return ""

# =============================================================================
#  NLP TOOLS (lazy init, optional)
# =============================================================================

_grammar_tool = None
_sim_model    = None

def _get_grammar_tool():
    global _grammar_tool
    if _grammar_tool is None and _NLP_AVAILABLE:
        try:
            _grammar_tool = language_tool_python.LanguageTool(
                "en-US", remote_server="https://api.languagetool.org")
        except Exception as e:
            log_.warning(f"LanguageTool init failed: {e}")
    return _grammar_tool

def _get_sim_model():
    global _sim_model
    if _sim_model is None and _NLP_AVAILABLE:
        try:
            _sim_model = SentenceTransformer("all-MiniLM-L6-v2", device="cpu")
        except Exception as e:
            log_.warning(f"SentenceTransformer init failed: {e}")
    return _sim_model

def grammar_correct(text: str) -> str:
    tool = _get_grammar_tool()
    if tool:
        try:
            return language_tool_python.utils.correct(text, tool.check(text))
        except Exception:
            pass
    return text

def similarity_score(a: str, b: str) -> float:
    model = _get_sim_model()
    if model:
        try:
            emb = model.encode([a, b], convert_to_tensor=True)
            return float(st_util.pytorch_cos_sim(emb[0], emb[1]))
        except Exception:
            pass
    def tokens(s):
        return set(re.sub(r"[^a-z0-9 ]", " ", s.lower()).split())
    ta, tb = tokens(a), tokens(b)
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(len(ta), len(tb))

def clean_output(text: str) -> str:
    text = _fix_mojibake(text)
    for pat in [r"\[/?INST\]", r"</?s>",
                r"(?i)(rewritten?|rephrased?|output|paraphrase[d]?)[:\s]+",
                r"\*\*", r"###", r"---"]:
        text = re.sub(pat, "", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return grammar_correct(text.strip())

# =============================================================================
#  MISTRAL API
# =============================================================================

def mistral_generate(prompt: str, max_tokens: int = 400, temperature: float = 0.7) -> str:
    if not MISTRAL_API_KEY:
        log_.warning("MISTRAL_API_KEY not set — skipping paraphrase")
        return ""
    try:
        response = requests.post(
            MISTRAL_URL,
            headers={
                "Authorization": f"Bearer {MISTRAL_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": MISTRAL_MODEL,
                "messages": [{"role": "user", "content": prompt}],
                "max_tokens": max_tokens,
                "temperature": temperature,
            },
            timeout=30,
        )
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        log_.error(f"Mistral API error: {e}")
        return ""

# =============================================================================
#  PARAPHRASE FUNCTIONS
# =============================================================================

def _print_wrapped(text: str, prefix: str = "   ", width: int = 100):
    words = text.split()
    line  = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= width:
            print(f"{prefix}{' '.join(line)}")
            line = []
    if line:
        print(f"{prefix}{' '.join(line)}")

def paraphrase_title(title: str) -> str:
    if not ENABLE_PARAPHRASE:
        return title
    clean = sanitize_text(title)
    if not clean:
        return title

    print(f"\n ┌─ TITLE PARAPHRASE {'─'*45}")
    print(f" │ Original : \"{clean}\"")
    print(f" │ {'─'*60}")

    best_result = None
    best_sim    = 0.0

    for attempt in range(4):
        temp = round(0.68 + attempt * 0.06, 2)
        print(f" │ Attempt {attempt+1} (temp={temp}):")

        prompt = (
            f"Rewrite this job title professionally using different words. "
            f"Output ONLY the rewritten title, nothing else. "
            f"Keep it between 4 and 12 words.\n\nJob title: {clean}"
        )

        raw    = mistral_generate(prompt, max_tokens=50, temperature=temp)
        result = clean_output(raw).split("\n")[0].strip().strip('"').strip("'")

        wc     = len(result.split()) if result else 0
        sim    = similarity_score(clean, result) if result else 0.0
        is_dup = result.lower().strip() == clean.lower().strip()

        print(f" │    Output  : \"{result}\"")
        print(f" │    Words   : {wc} | Similarity: {sim:.3f} | Duplicate: {'Yes' if is_dup else 'No'}")

        valid = bool(result) and 4 <= wc <= 14 and sim >= 0.55 and not is_dup

        if not valid:
            reasons = []
            if not result:  reasons.append("empty output")
            if wc < 4:      reasons.append(f"too short ({wc} words, min=4)")
            if wc > 14:     reasons.append(f"too long ({wc} words, max=14)")
            if sim < 0.55:  reasons.append(f"sim={sim:.3f} < 0.55")
            if is_dup:      reasons.append("identical to original")
            print(f" │    -> REJECTED — {', '.join(reasons)}")
        else:
            if sim > best_sim:
                best_sim    = sim
                best_result = result
                print(f" │    -> ACCEPTED — new best candidate (sim={sim:.3f})")
            else:
                print(f" │    -> VALID but not better than current best (best sim={best_sim:.3f})")

        print(f" │ {'─'*60}")
        time.sleep(1)

    if best_result:
        print(f" │ FINAL SELECTED : \"{best_result}\"")
        print(f" │    Similarity  : {best_sim:.3f}")
        print(f" └{'─'*65}")
        return best_result
    else:
        print(f" │ No valid paraphrase found -> Keeping original: \"{clean}\"")
        print(f" └{'─'*65}")
        return clean

def paraphrase_description(text: str) -> str:
    if not ENABLE_PARAPHRASE:
        return text
    clean = sanitize_text(text)
    if not clean:
        return text

    paragraphs  = [p.strip() for p in re.split(r"\n+", clean) if p.strip()]
    if not paragraphs:
        paragraphs = [clean]
    rewritten   = []
    success_count = 0

    print(f"\n ┌─ DESCRIPTION PARAPHRASE ({len(paragraphs)} paragraph(s)) {'─'*15}")

    for i, para in enumerate(paragraphs):
        orig_wc = len(para.split())

        print(f"\n │ ┌─ Paragraph {i+1}/{len(paragraphs)} {'─'*50}")
        print(f" │ │ ORIGINAL ({orig_wc} words):")
        _print_wrapped(para, prefix=" │ │    ")
        print(f" │ │ {'─'*60}")

        if orig_wc < 8:
            print(f" │ │ (kept — too short to paraphrase safely)")
            rewritten.append(para)
            print(f" │ └{'─'*62}")
            continue

        prompt = (
            f"Rewrite this job description paragraph professionally. "
            f"Keep ALL facts, requirements, and responsibilities. "
            f"Use different sentence structure and vocabulary. "
            f"Output ONLY the rewritten paragraph — no labels, no explanation.\n\n"
            f"Original:\n{para}"
        )

        best_result = None
        best_sim    = 0.0
        accepted_text = None

        for attempt in range(3):
            temp = round(0.65 + attempt * 0.08, 2)
            print(f" │ │ Attempt {attempt+1}/3 (temp={temp}):")

            raw    = mistral_generate(prompt, max_tokens=500, temperature=temp)
            result = clean_output(raw).strip()

            rw  = len(result.split()) if result else 0
            sim = similarity_score(para, result) if result and rw >= 5 else 0.0

            if result:
                print(f" │ │    Paraphrased ({rw} words, sim={sim:.3f}):")
                _print_wrapped(result, prefix=" │ │       ")
            else:
                print(f" │ │    Paraphrased : (no output from model)")

            valid = bool(result) and rw >= 8 and sim >= 0.48

            if not valid:
                reasons = []
                if not result: reasons.append("empty output")
                if rw < 8:     reasons.append(f"too short ({rw} words, min=8)")
                if sim < 0.48: reasons.append(f"sim={sim:.3f} < 0.48")
                print(f" │ │    -> REJECTED — {', '.join(reasons)}")
                if result and sim > best_sim:
                    best_sim    = sim
                    best_result = result
                    print(f" │ │       (stored as best fallback, sim={sim:.3f})")
            else:
                print(f" │ │    -> ACCEPTED on attempt {attempt+1}")
                rewritten.append(result)
                success_count += 1
                accepted_text = result
                break

            print(f" │ │ {'─'*60}")
            time.sleep(1)

        if accepted_text is None:
            print(f" │ │ {'─'*60}")
            if best_result and best_sim >= 0.40:
                print(f" │ │ FALLBACK — Using best attempt (sim={best_sim:.3f}):")
                _print_wrapped(best_result, prefix=" │ │    ")
                rewritten.append(best_result)
                success_count += 1
            else:
                print(f" │ │ KEPT ORIGINAL — no acceptable paraphrase (best sim={best_sim:.3f})")
                rewritten.append(para)

        print(f" │ └{'─'*62}")

    print(f"\n │ SUMMARY: {success_count}/{len(paragraphs)} paragraphs successfully paraphrased")
    print(f" └{'─'*80}\n")

    return "\n\n".join(rewritten)

# =============================================================================
#  DUPLICATE TRACKER
# =============================================================================

def _init_tracker():
    if not os.path.exists(PROCESSED_IDS_FILE):
        try:
            with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_TRACKER_FIELDS)
            log_.info(f"Tracker file created: {PROCESSED_IDS_FILE}")
        except Exception as e:
            log_.error(f"Could not create tracker file {PROCESSED_IDS_FILE}: {e}")

def load_processed_ids() -> tuple:
    _init_tracker()
    ids, urls = set(), set()
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("Job ID"):
                    ids.add(row["Job ID"].strip())
                if row.get("Job URL"):
                    urls.add(row["Job URL"].strip())
    except Exception as e:
        log_.error(f"Could not read tracker file: {e}")
    return ids, urls

def _upsert_row(job_id: str, updates: dict):
    _init_tracker()
    rows = []
    try:
        with open(PROCESSED_IDS_FILE, newline="", encoding="utf-8") as f:
            rows = list(csv.DictReader(f))
    except Exception as e:
        log_.error(f"Tracker read error: {e}")
        rows = []

    found = False
    for row in rows:
        if row.get("Job ID", "").strip() == str(job_id):
            row.update(updates)
            row["Timestamp"] = datetime.now().isoformat()
            found = True
            break

    if not found:
        new_row = {k: "" for k in _TRACKER_FIELDS}
        new_row["Job ID"]    = str(job_id)
        new_row["Timestamp"] = datetime.now().isoformat()
        new_row.update(updates)
        rows.append(new_row)

    try:
        with open(PROCESSED_IDS_FILE, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=_TRACKER_FIELDS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
    except Exception as e:
        log_.error(f"Tracker write error: {e}")

def make_job_id(job_url: str, title: str = "", company: str = "") -> str:
    # PATCH E: accept 5+ digit IDs (not just 6+)
    m = re.search(r"/jobs/(\d{5,})", job_url or "")
    if m:
        return m.group(1)
    if job_url:
        return hashlib.md5(job_url.encode()).hexdigest()[:16]
    return hashlib.md5(f"{title}{company}".encode()).hexdigest()[:16]

def mark_scraped(job_id, job_url, title, company):
    log_.info(f"Tracker -> scraped: {job_id} | {title}")
    _upsert_row(job_id, {
        "Job URL":      job_url,
        "Job Title":    title,
        "Company Name": company,
        "Status":       "scraped",
        "WP ID":        "",
    })

def mark_paraphrased(job_id):
    _upsert_row(job_id, {"Status": "paraphrased"})

def mark_posted(job_id, wp_id, wp_url):
    _upsert_row(job_id, {"Status": "posted", "WP ID": str(wp_id)})

def mark_failed(job_id, reason):
    _upsert_row(job_id, {"Status": f"failed|{reason}"})

# =============================================================================
#  FLAGGED CSV
# =============================================================================

def _init_flagged():
    if not os.path.exists(FLAGGED_FILE):
        try:
            with open(FLAGGED_FILE, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(_FLAGGED_FIELDS)
        except Exception as e:
            log_.error(f"Could not create flagged file {FLAGGED_FILE}: {e}")

def write_flagged(raw_job: dict, reason: str, apply_note: str):
    _init_flagged()
    try:
        with open(FLAGGED_FILE, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                "NewTimes Rwanda",
                raw_job.get("title", ""),
                raw_job.get("company_name", ""),
                raw_job.get("location", ""),
                raw_job.get("salary", ""),
                raw_job.get("deadline", ""),
                reason,
                apply_note,
                raw_job.get("job_url", ""),
                datetime.now().isoformat(),
            ])
    except Exception as e:
        log_.error(f"Flagged write error: {e}")

# =============================================================================
#  WORDPRESS POSTING
# =============================================================================

def _wp_auth_headers() -> dict:
    token = base64.b64encode(f"{WP_USER}:{WP_PASSWORD}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}

def get_or_create_term(taxonomy_url: str, name: str):
    if not name or not name.strip():
        return None
    slug = re.sub(r"[^a-z0-9-]", "-", name.lower().strip())
    h = _wp_auth_headers()
    try:
        r = requests.get(f"{taxonomy_url}?slug={slug}", headers=h, timeout=10, verify=False)
        terms = r.json()
        if isinstance(terms, list) and terms:
            return terms[0]["id"]
    except Exception:
        pass
    try:
        r = requests.post(taxonomy_url, json={"name": name, "slug": slug},
                          headers=h, auth=(WP_USER, WP_PASSWORD), timeout=10, verify=False)
        return r.json().get("id")
    except Exception as e:
        log_.error(f"Term create error '{name}': {e}")
        return None

def post_job_to_wordpress(job: dict) -> tuple:
    if not WP_USER or not WP_PASSWORD:
        log_.warning("WP_USERNAME / WP_APP_PASSWORD not set — skipping WordPress post")
        return None, None

    h = _wp_auth_headers()

    title       = sanitize_text(job.get("jobTitle", ""))
    description = sanitize_text(job.get("jobDescription", ""))
    if not title or not description:
        return None, None

    slug = re.sub(r"[^a-z0-9-]", "-", title.lower())[:80]
    try:
        r = requests.get(f"{WP_JOBS_URL}?slug={slug}", headers=h, timeout=10, verify=False)
        posts = r.json()
        if isinstance(posts, list) and posts:
            log_.info(f"Job already on WP: {title}")
            return posts[0]["id"], posts[0].get("link")
    except Exception:
        pass

    logo_url    = sanitize_text(job.get("companyLogo", ""), is_url=True)
    location    = sanitize_text(job.get("jobLocation", ""))
    raw_type    = sanitize_text(job.get("jobType", "")) or "Full-time"
    job_type_s  = JOB_TYPE_MAPPING.get(raw_type.lower().strip(), "full-time")
    company     = sanitize_text(job.get("companyName", ""))
    application = sanitize_text(job.get("application", ""), is_url=True)
    company_url = sanitize_text(job.get("companyUrl", ""), is_url=True)
    deadline    = sanitize_text(job.get("deadline", ""))
    co_website  = sanitize_text(job.get("companyWebsite", ""), is_url=True)
    qualif      = sanitize_text(job.get("jobQualifications", ""))
    experience  = sanitize_text(job.get("jobExperience", ""))
    co_address  = sanitize_text(job.get("companyAddress", ""))
    job_field   = sanitize_text(job.get("jobField", ""))
    salary      = sanitize_text(job.get("salaryRange", ""))
    about       = sanitize_text(job.get("companyDetails", ""))

    is_email = bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", application))
    is_url_v = bool(re.match(r"^https?://[^\s]+$", application))
    if not (is_email or is_url_v):
        application = ""

    attachment_id = None
    if logo_url:
        try:
            img_r = requests.get(logo_url, headers={"User-Agent": HEADERS["User-Agent"]}, timeout=15)
            if img_r.status_code == 200:
                ct  = img_r.headers.get("Content-Type", "image/jpeg")
                ext = "png" if "png" in ct else "jpg"
                fn  = re.sub(r"[^a-z0-9]", "-", company.lower()) + "-logo." + ext
                up_h = dict(_wp_auth_headers())
                up_h["Content-Disposition"] = f"attachment; filename={fn}"
                up_h["Content-Type"] = ct
                up_r = requests.post(WP_MEDIA_URL, headers=up_h, data=img_r.content,
                                     auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
                if up_r.status_code in (200, 201):
                    attachment_id = up_r.json().get("id")
        except Exception as e:
            log_.warning(f"Logo upload failed: {e}")

    region_term_id   = get_or_create_term(f"{WP_BASE}/job_listing_region", location)
    job_type_term_id = get_or_create_term(f"{WP_BASE}/job_listing_type",
                                           job_type_s.replace("-", " ").title())

    payload = {
        "title":          title,
        "content":        description,
        "status":         "publish",
        "featured_media": attachment_id or 0,
        "meta": {
            "_job_title":          title,
            "_job_location":       location,
            "_job_type":           job_type_s,
            "_job_description":    description,
            "_application":        application,
            "_company_url":        company_url,
            "_job_expires":        deadline,
            "_company_name":       company,
            "_company_website":    co_website,
            "_company_logo":       str(attachment_id) if attachment_id else "",
            "_company_address":    co_address,
            "_company_details":    about,
            "_job_qualifications": qualif,
            "_job_experiences":    experience,
            "_job_field":          job_field,
            "_job_salary":         salary,
        },
    }
    if region_term_id:   payload["job_listing_region"] = [region_term_id]
    if job_type_term_id: payload["job_listing_type"]   = [job_type_term_id]

    for attempt in range(3):
        try:
            r = requests.post(WP_JOBS_URL, json=payload, headers=h,
                              auth=(WP_USER, WP_PASSWORD), timeout=20, verify=False)
            r.raise_for_status()
            post = r.json()
            log_.info(f"Job posted: '{title}' -> WP ID {post.get('id')}")
            return post.get("id"), post.get("link")
        except Exception as e:
            log_.error(f"Job post attempt {attempt+1} failed: {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    return None, None

# =============================================================================
#  STEP 1 — COLLECT JOB DETAIL URLS
# =============================================================================

def _norm_detail_url(path_or_url: str) -> str:
    if not path_or_url:
        return ""
    if path_or_url.startswith("http"):
        url = path_or_url
    else:
        url = urljoin(BASE_URL + "/", path_or_url.lstrip("/"))
    p = urlsplit(url)
    path = p.path.rstrip("/")
    return urlunsplit(("https", "jobs.newtimes.co.rw", path, "", ""))

# PATCH E: lowered min digit count to 5
def _is_job_detail_path(path: str) -> bool:
    parts = [s for s in path.split("/") if s]
    if len(parts) != 2 or parts[0] != "jobs":
        return False
    return bool(re.match(r"^\d{5,}", parts[1]))

# PATCH D-1: extract job links from __NEXT_DATA__ / __NUXT__ JSON blobs
def _extract_links_from_page_data(html: str) -> list:
    """
    NewTimes uses a JS framework (Next.js or Nuxt). Job card data is embedded
    in a <script id="__NEXT_DATA__"> or window.__NUXT__ block in the HTML.
    Extract all /jobs/{id}-{slug} href values from these blobs.
    """
    links = []
    # Pattern 1: Next.js  <script id="__NEXT_DATA__" type="application/json">
    nd_match = re.search(r'<script[^>]+id=["\']__NEXT_DATA__["\'][^>]*>(.*?)</script>',
                         html, re.S | re.I)
    if nd_match:
        try:
            data = json.loads(nd_match.group(1))
            blob = json.dumps(data)
            for m in re.finditer(r'"/jobs/(\d{5,}-[^"]+)"', blob):
                links.append(f"/jobs/{m.group(1)}")
        except Exception:
            pass

    # Pattern 2: Nuxt  window.__NUXT__={...}
    nuxt_match = re.search(r'window\.__NUXT__\s*=\s*(\{.*?\})\s*;', html, re.S)
    if nuxt_match:
        try:
            blob = nuxt_match.group(1)
            for m in re.finditer(r'"/jobs/(\d{5,}-[^"]+)"', blob):
                links.append(f"/jobs/{m.group(1)}")
        except Exception:
            pass

    # Pattern 3: Any JSON-like string with jobs path (catch-all)
    for m in re.finditer(r'["\'](/jobs/\d{5,}-[^"\'\s>]+)["\']', html):
        links.append(m.group(1))

    return links


def _collect_via_api(category: str, max_pages: int) -> list:
    """Walk the JSON API. Returns canonical detail URLs, or [] if unavailable."""
    print(C_BLUE(f"\n  Trying JSON API: {_NEWTIMES_API_URL}"))
    seen, ordered = set(), []
    cat_map = {"announcement": "announcement", "tenders": "tenders", "all": None}
    cat_param = cat_map.get(category)

    # PATCH F: try both API endpoints
    api_endpoints = [_NEWTIMES_API_URL, _NEWTIMES_API_ALT]
    working_endpoint = None

    for endpoint in api_endpoints:
        params = {"page": 1, "per_page": 20}
        if cat_param:
            params["category"] = cat_param
        data = get_json(endpoint, params=params)
        if data is not None:
            working_endpoint = endpoint
            log(f"  API endpoint active: {endpoint}")
            break

    if working_endpoint is None:
        log(C_DIM("  JSON API unavailable — switching to HTML fallback."))
        return []

    # Walk pages on the working endpoint
    for page in range(1, max_pages + 1):
        params = {"page": page, "per_page": 20}
        if cat_param:
            params["category"] = cat_param
        data = get_json(working_endpoint, params=params)
        if data is None:
            break

        jobs_list = []
        if isinstance(data, list):
            jobs_list = data
        elif isinstance(data, dict):
            jobs_list = data.get("jobs") or data.get("data") or data.get("results") or []

        if not jobs_list:
            log(C_DIM(f"  API page {page}: empty — stopping."))
            break

        page_new = 0
        for item in jobs_list:
            jid   = str(item.get("id") or item.get("job_id") or "")
            slug  = item.get("slug") or slugify(item.get("title", ""))
            if not jid:
                continue
            detail_url = _norm_detail_url(f"/jobs/{jid}-{slug}")
            if detail_url not in seen:
                seen.add(detail_url)
                ordered.append(detail_url)
                page_new += 1

        log(f"    API page {page}: {page_new} new link(s) (total {len(ordered)})")

        total_pages = (data.get("total_pages") or data.get("last_page") or 0
                       if isinstance(data, dict) else 0)
        if total_pages and page >= total_pages:
            break
        if page_new == 0:
            break

        time.sleep(REQUEST_DELAY)

    return ordered


# PATCH A+D: corrected archive URL params and __NEXT_DATA__ extraction
def _collect_via_html(archive_url: str, max_pages: int) -> list:
    """
    Scrape archive HTML for job detail links.
    1. Try ?q=&page=N (correct query param format for this SPA)
    2. Check __NEXT_DATA__ / __NUXT__ blobs for embedded card data
    3. Fall back to raw href scanning
    """
    cat_param = ""
    if SCRAPE_CATEGORY == "announcement":
        cat_param = "announcement"
    elif SCRAPE_CATEGORY == "tenders":
        cat_param = "tenders"

    print(C_BLUE(f"\n  Collecting job links from HTML: {archive_url}"))
    seen, ordered = set(), []
    empty_streak = 0

    for page in range(1, max_pages + 1):
        # Build the correct query string: ?q=&page=N[&category=X]
        params = {"q": "", "page": page}
        if cat_param:
            params["category"] = cat_param

        param_str = "&".join(f"{k}={v}" for k, v in params.items())
        url = f"{archive_url}?{param_str}"

        try:
            _prime_session()
            resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            raw_html = resp.text
        except requests.HTTPError as e:
            status = getattr(e.response, 'status_code', '?')
            log(C_RED(f"  Page {page}: HTTP {status} — {url}"))
            if status == 403:
                log(C_RED("  403 Forbidden — site may require JS execution (Cloudflare/bot protection)."))
                log(C_DIM("  Tip: set NEWTIMES_COOKIE env var with a valid _session cookie from your browser."))
            break
        except Exception as e:
            log(C_DIM(f"  Page {page}: fetch error ({e}) — stopping."))
            break

        page_new = 0

        # Strategy 1: __NEXT_DATA__ / __NUXT__ embedded JSON
        embedded = _extract_links_from_page_data(raw_html)
        for path in embedded:
            norm = _norm_detail_url(path)
            if norm and norm not in seen and _is_job_detail_path(urlparse(norm).path):
                seen.add(norm)
                ordered.append(norm)
                page_new += 1

        # Strategy 2: raw href scanning (SSR or partial hydration)
        if page_new == 0:
            soup = BeautifulSoup(raw_html, "html.parser")
            for a in soup.find_all("a", href=True):
                href = a["href"].strip()
                if href.startswith("/"):
                    full = urljoin(BASE_URL, href)
                elif href.startswith("http"):
                    full = href
                else:
                    continue
                path = urlparse(full).path
                if not _is_job_detail_path(path):
                    continue
                norm = _norm_detail_url(full)
                if norm and norm not in seen:
                    seen.add(norm)
                    ordered.append(norm)
                    page_new += 1

        log(f"    HTML page {page}: {page_new} new link(s) (total {len(ordered)})")

        if page_new == 0:
            empty_streak += 1
            if empty_streak >= 2:
                break
        else:
            empty_streak = 0

        time.sleep(REQUEST_DELAY)

    return ordered


def collect_job_links(archive_url: str, category: str, max_pages: int) -> list:
    """Return ordered, de-duplicated detail URLs using API first, HTML fallback."""
    links = _collect_via_api(category, max_pages)
    if not links:
        links = _collect_via_html(archive_url, max_pages)
    return links

# =============================================================================
#  STEP 2 — PARSE ONE NEWTIMES DETAIL PAGE
# =============================================================================

_CONTENT_SELECTORS = [
    "div.job-description",
    "div.job-content",
    "div.single-job-content",
    "div.description",
    "section.job-description",
    "article.job-content",
    "div.content-area",
    "div.entry-content",
    "main .content",
    "main",
]

def _find_content_el(soup: BeautifulSoup):
    best, best_len = None, 0
    for sel in _CONTENT_SELECTORS:
        el = soup.select_one(sel)
        if el:
            txt = el.get_text(" ", strip=True)
            if len(txt) > best_len:
                best, best_len = el, len(txt)
        if best and best_len > 300:
            return best
    for tag in ("article", "section", "div"):
        for el in soup.find_all(tag):
            if el.find_parent(["nav", "header", "footer"]):
                continue
            txt = el.get_text(" ", strip=True)
            if len(txt) > best_len:
                best, best_len = el, len(txt)
    return best or soup.body or soup

def _is_real_apply_url(href: str) -> bool:
    if not href:
        return False
    low = href.lower()
    if low.startswith("mailto:") or low.startswith("#") or low.startswith("javascript:"):
        return False
    if not low.startswith("http"):
        return False
    if any(s in low for s in _NON_APPLY_HOST_SUBSTR):
        return False
    if any(s in low for s in _NON_APPLY_PATH_SUBSTR):
        return False
    return True

def _is_apply_heading_line(line: str) -> bool:
    s = line.strip().lstrip("•*-–—#:. ").strip()
    if not s or len(s.split()) > 9:
        return False
    return bool(_APPLY_HEAD_PHRASES.match(s))

def _split_description_and_apply(content_text: str):
    if not content_text:
        return "", ""

    lines = content_text.split("\n")
    kept = []
    for ln in lines:
        low = ln.strip().lower()
        if low in _BODY_DROP_LINES:
            continue
        if any(low.startswith(m) for m in _BODY_CUT_MARKERS):
            break
        kept.append(ln)

    apply_idx = None
    for i, ln in enumerate(kept):
        if _is_apply_heading_line(ln):
            apply_idx = i
            break

    if apply_idx is None:
        return "\n".join(kept).strip(), ""

    description = "\n".join(kept[:apply_idx]).strip()
    apply_text  = "\n".join(kept[apply_idx:]).strip()
    if not description:
        return "\n".join(kept).strip(), ""
    return description, apply_text


def _extract_meta_line(soup: BeautifulSoup) -> str:
    for sel in (".job-meta", ".job-header-meta", ".post-meta", ".listing-meta",
                "[class*='published']", "[class*='deadline']", "[class*='meta']"):
        el = soup.select_one(sel)
        if el:
            t = el.get_text(" ", strip=True)
            if "published" in t.lower() or "deadline" in t.lower():
                return t

    page_text = soup.get_text(" ")
    m = re.search(
        r"[Pp]ublished\s+on\s+\d{1,2}[/-]\d{1,2}[/-]\d{4}"
        r"(?:[^\n]{0,50}[Dd]eadline\s+\d{1,2}[/-]\d{1,2}[/-]\d{4})?",
        page_text
    )
    return m.group(0) if m else ""


def _extract_company(soup: BeautifulSoup, fallback_from_url_slug: str = "") -> tuple:
    for sel in (".company-name", ".employer-name", ".company-title",
                "[class*='company']", "[class*='employer']",
                "a[href*='/employers/']"):
        el = soup.select_one(sel)
        if el:
            name = el.get_text(" ", strip=True)
            href = el.get("href", "")
            if name and len(name) < 120:
                url = urljoin(BASE_URL, href) if href and not href.startswith("http") else href
                return name, url

    for prop in (("property", "article:author"), ("name", "author")):
        tag = soup.find("meta", attrs={prop[0]: prop[1]})
        if tag and tag.get("content"):
            return tag["content"].strip(), ""

    if fallback_from_url_slug:
        m = re.search(r"-at-([a-z0-9-]+?)(?:-published-on|-deadline|-$)", fallback_from_url_slug, re.I)
        if m:
            name = m.group(1).replace("-", " ").title()
            return name, ""

    return "NewTimes Rwanda Employer", ""


def scrape_job_detail(url: str) -> dict:
    soup = get_soup(url)

    og_title_tag = soup.find("meta", attrs={"property": "og:title"})
    og_title = (og_title_tag.get("content", "") if og_title_tag else "").strip()
    h1 = (soup.select_one("h1.job-title") or soup.select_one("h1.entry-title")
          or soup.select_one("h1") or soup.find("h1"))
    h1_text = h1.get_text(" ", strip=True) if h1 else ""
    raw_title = h1_text or og_title
    title = clean_title(raw_title)
    if not title and og_title:
        title = clean_title(og_title)

    logo = ""
    og_img = soup.find("meta", attrs={"property": "og:image"})
    if og_img and og_img.get("content"):
        logo = og_img["content"].strip()
    if not logo:
        emp_img = (soup.select_one("a[href*='/employers/'] img") or
                   soup.select_one(".company-logo img") or
                   soup.select_one(".employer-logo img"))
        if emp_img and emp_img.get("src"):
            logo = emp_img["src"].strip()

    url_slug = urlparse(url).path
    company_name, company_profile_url = _extract_company(soup, url_slug)

    company_website = ""
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        if (href.startswith("http") and
                not any(s in href.lower() for s in _NON_APPLY_HOST_SUBSTR) and
                "newtimes" not in href.lower() and
                a.get_text(strip=True).lower() in ("website", "visit website",
                                                   "company website", "official website")):
            company_website = href
            break

    meta_line = _extract_meta_line(soup)
    page_text  = soup.get_text("\n")

    date_posted = ""
    deadline    = ""
    pub_m = re.search(r"[Pp]ublished\s+on\s+(\d{1,2}[/-]\d{1,2}[/-]\d{4})", meta_line or page_text)
    if pub_m:
        date_posted = parse_any_date(pub_m.group(1))
    dead_m = re.search(r"[Dd]eadline\s+(\d{1,2}[/-]\d{1,2}[/-]\d{4})", meta_line or page_text)
    if dead_m:
        deadline = parse_any_date(dead_m.group(1))

    for lab in DEADLINE_LABELS:
        m = re.search(rf"{lab}\s*[:\-]?\s*([^\n<]{{3,60}})", page_text, re.I)
        if m:
            d = parse_any_date(m.group(1))
            if d:
                deadline = d
                break

    if not date_posted:
        date_posted = datetime.now().strftime("%Y-%m-%d")
    if not deadline:
        deadline = (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d")

    location = DEFAULT_LOCATION
    loc_m = re.search(
        r"([A-Za-z][A-Za-z\s]+),?\s*(Rwanda|Kigali Province|Southern Province|"
        r"Northern Province|Eastern Province|Western Province)",
        meta_line or page_text, re.I
    )
    if loc_m:
        raw_loc = f"{loc_m.group(1).strip()}, {loc_m.group(2).strip()}"
        location = raw_loc
    else:
        location = location_from_text(meta_line or page_text[:500])

    job_type = "full-time"
    type_m = re.search(
        r"\b(Full[\s-]?Time|Part[\s-]?Time|Contract|Internship|Freelance|Volunteer|"
        r"Temporary|Invitation\s+for\s+Bid|Tenders?)\b",
        meta_line or page_text[:600], re.I
    )
    if type_m:
        job_type = map_job_type(type_m.group(1).lower().strip())

    category_tag = ""
    for sel in (".job-category", ".category-tag", ".post-category",
                "a[href*='/jobs/search/']", "span.category"):
        el = soup.select_one(sel)
        if el:
            t = el.get_text(strip=True)
            if t and len(t) < 40:
                category_tag = t
                break

    content_el   = _find_content_el(soup)
    content_copy = BeautifulSoup(str(content_el), "lxml")
    content_text = html_block_to_text(content_copy)
    description, apply_text = _split_description_and_apply(content_text)
    if not description:
        description = content_text

    qual_block = ""
    qm = re.search(
        r"(?:^|\n)[ \t]*qualifications?(?:\s*(?:&|and)\s*experience)?(?:\s+\w+){0,3}\s*:?[ \t]*\n"
        r"(.*?)"
        r"(?:\n[ \t]*(?:how\s*(?:and|&)?\s*(?:deadline\s*)?to\s*apply|what\s+we\s+offer|"
        r"key\s+competenc|interested\s+candidates?|submission|mode\s+of\s+application)\b"
        r"|\n[ \t]*[A-Z][^\n]{0,60}:[ \t]*\n|\Z)",
        description, re.I | re.S
    )
    if qm:
        qual_block = qm.group(1).strip()[:1500]

    qualification = extract_qualification(qual_block or description)
    experience    = extract_experience_band(qual_block or description)

    job_field = infer_field(title, description, category_tag)

    apply_email = ""
    apply_url   = ""

    for a in content_el.find_all("a", href=True):
        href = a["href"].strip()
        if href.lower().startswith("mailto:"):
            cand = extract_email(href[7:])
            if cand and _is_real_apply_email(cand):
                apply_email = apply_email or cand
        elif _is_real_apply_url(href):
            apply_url = apply_url or strip_tracking_params(href)

    scan = apply_text or description
    if not apply_email:
        cand = extract_email(scan)
        if cand and _is_real_apply_email(cand):
            apply_email = cand
    if not apply_url:
        for u in URL_PATTERN.findall(scan):
            if _is_real_apply_url(u):
                apply_url = strip_tracking_params(u.rstrip(".,);"))
                break

    physical_submit = ""
    phys_m = re.search(
        r"(?:submit(?:ted)?|send|deliver|drop)\b[^.\n]{0,120}"
        r"(?:P\.?O\.?\s*Box|P/O\s*Box|[Aa]ddress|[Oo]ffice|[Hh]ead[Qq]uarters|[Pp]remises)[^.\n]{0,80}",
        apply_text or description, re.I
    )
    if phys_m:
        physical_submit = phys_m.group(0).strip()

    salary = extract_salary(description)

    return {
        "title":           title,
        "company_name":    company_name,
        "company_url":     company_profile_url,
        "company_website": company_website,
        "company_address": location,
        "company_logo":    logo,
        "job_type":        job_type,
        "location":        location,
        "job_field":       job_field,
        "job_categories":  category_tag,
        "date_posted":     date_posted,
        "deadline":        deadline,
        "description":     description,
        "qualification":   qualification,
        "experience":      experience,
        "salary":          salary,
        "apply_email":     apply_email,
        "apply_url":       apply_url,
        "apply_text":      apply_text,
        "physical_submit": physical_submit,
        "job_url":         _norm_detail_url(url),
    }

# =============================================================================
#  STEP 3 — DEDUPLICATE + PARAPHRASE + APPLY-RULE GATING
# =============================================================================

def process_job(raw_job: dict, processed_ids: set, processed_urls: set, seen_content: set):
    job_url  = raw_job.get("job_url", "")
    title    = raw_job.get("title", "")
    company  = raw_job.get("company_name", "")
    location = raw_job.get("location", "")

    if not title:
        return "duplicate", None

    job_id = make_job_id(job_url, title, company)

    if job_id in processed_ids or job_url in processed_urls:
        log(C_DIM(f"  Already processed (tracker) — skipped: {title}"))
        return "duplicate", None

    fingerprint = (title.lower().strip(), company.lower().strip(), location.lower().strip())
    if fingerprint in seen_content:
        log(C_DIM(f"  Duplicate content this run — skipped: {title}"))
        return "duplicate", None
    seen_content.add(fingerprint)

    apply_email = raw_job.get("apply_email", "")
    apply_url   = raw_job.get("apply_url", "")
    physical    = raw_job.get("physical_submit", "")
    qualifies   = bool(apply_email) or bool(apply_url) or bool(physical)

    if REQUIRE_PUBLIC_APPLY and not qualifies:
        write_flagged(raw_job,
                      "no public apply method (no email, external URL, or submission address found)",
                      raw_job.get("apply_text", "")[:300])
        log(C_RED(f"  FLAGGED (no public apply) — {title}"))
        return "flagged", None

    mark_scraped(job_id, job_url, title, company)
    processed_ids.add(job_id)
    processed_urls.add(job_url)

    description = raw_job.get("description", "")
    paraphrased_title = title
    paraphrased_desc  = description

    if ENABLE_PARAPHRASE and MISTRAL_API_KEY:
        print(C_BLUE(f"\n  Paraphrasing '{title}' ..."))
        paraphrased_title = paraphrase_title(title)
        paraphrased_desc  = paraphrase_description(description)
        mark_paraphrased(job_id)
    else:
        print(C_DIM("  Paraphrasing skipped (ENABLE_PARAPHRASE=False or MISTRAL_API_KEY not set)"))

    application = apply_email or apply_url or physical
    apply_method = ("email" if apply_email
                    else "external_url" if apply_url
                    else "physical_address" if physical
                    else "not_found")

    company_link = raw_job.get("company_website") or raw_job.get("company_url", "")

    return "ok", {
        "jobTitle":          paraphrased_title,
        "jobDescription":    paraphrased_desc,
        "companyDetails":    "",
        "originalTitle":     title,
        "originalDesc":      description,
        "jobType":           raw_job.get("job_type", "full-time"),
        "jobQualifications": raw_job.get("qualification", ""),
        "jobExperience":     raw_job.get("experience", ""),
        "jobLocation":       location,
        "jobField":          raw_job.get("job_field", ""),
        "datePosted":        raw_job.get("date_posted", datetime.now().strftime("%Y-%m-%d")),
        "deadline":          raw_job.get("deadline", ""),
        "application":       application,
        "companyUrl":        company_link,
        "companyName":       company,
        "companyLogo":       raw_job.get("company_logo", ""),
        "companyWebsite":    raw_job.get("company_website", ""),
        "companyAddress":    raw_job.get("company_address", "") or location,
        "jobUrl":            job_url,
        "salaryRange":       raw_job.get("salary", ""),
        "_jobId":            job_id,
        "_apply_method":     apply_method,
        "_apply_raw":        raw_job.get("apply_text", "")[:160],
    }

# =============================================================================
#  VERBOSE PRINTER
# =============================================================================

def print_job_verbose(index, job):
    desc = job.get("jobDescription", "")
    desc_preview = (desc[:400] + " [...]") if len(desc) > 400 else desc

    print()
    print(C_DIVIDER())
    print(C_HEADER(f"  JOB #{index}"))
    print(C_DIVIDER())
    print(f"  {C_LABEL('Title (original)')}    : {C_VALUE(job.get('originalTitle',''))}")
    print(f"  {C_LABEL('Title (paraphrased)')} : {C_GREEN(job.get('jobTitle',''))}")
    print(f"  {C_LABEL('Job Type')}             : {job.get('jobType','') or C_DIM('—')}")
    print(f"  {C_LABEL('Qualification')}        : {(job.get('jobQualifications','')[:120] or C_DIM('—'))}")
    print(f"  {C_LABEL('Experience')}           : {job.get('jobExperience','') or C_DIM('—')}")
    print(f"  {C_LABEL('Location')}             : {job.get('jobLocation','') or C_DIM('—')}")
    print(f"  {C_LABEL('Category/Field')}       : {job.get('jobField','') or C_DIM('—')}")
    print(f"  {C_LABEL('Salary')}               : {job.get('salaryRange','') or C_DIM('—')}")
    print(f"  {C_LABEL('Posted')}               : {job.get('datePosted','') or C_DIM('—')}")
    print(f"  {C_LABEL('Deadline')}             : {job.get('deadline','') or C_DIM('—')}")

    application = job.get("application", "")
    print(f"  {C_LABEL('Apply')}                : {C_GREEN(application) if application else C_DIM('— not found —')}")
    print(f"  {C_LABEL('Apply Method')}         : {C_DIM(job.get('_apply_method',''))}")

    print()
    print(f"  {C_BLUE('── EMPLOYER ─────────────────────────────────────────')}")
    print(f"  {C_LABEL('Name')}      : {C_VALUE(job.get('companyName','') or C_DIM('—'))}")
    print(f"  {C_LABEL('Website')}   : {job.get('companyWebsite','') or C_DIM('—')}")
    print(f"  {C_LABEL('Source')}    : {job.get('companyUrl','') or C_DIM('—')}")
    print(f"  {C_LABEL('Logo')}      : {job.get('companyLogo','') or C_DIM('— none —')}")

    print()
    print(f"  {C_BLUE('── DESCRIPTION PREVIEW ─────────────────────────────')}")
    print(desc_preview if desc_preview else C_DIM("   — no description —"))
    print(f"  {C_LABEL('Job URL')}   : {job.get('jobUrl','')}")
    print(C_DIVIDER())

# =============================================================================
#  EXCEL SAVE
# =============================================================================

EXCEL_HEADERS = [
    "Job Title", "Job Type", "Job Qualifications", "Job Experience",
    "Job Location", "Job Field", "Date Posted", "Deadline",
    "Job Description", "Application", "Company URL", "Company Name",
    "Company Logo", "Company Website", "Company Address",
    "Company Details", "Job URL", "Salary Range",
]

def _save_excel(jobs: list):
    if not _XLSX_AVAILABLE:
        log_.warning("pandas/openpyxl not installed — skipping Excel export")
        return
    if not jobs:
        return
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(EXCEL_HEADERS)
    for job in jobs:
        ws.append([
            job["jobTitle"], job["jobType"], job["jobQualifications"], job["jobExperience"],
            job["jobLocation"], job["jobField"], job["datePosted"], job["deadline"],
            job["jobDescription"], job["application"], job["companyUrl"], job["companyName"],
            job["companyLogo"], job["companyWebsite"], job["companyAddress"],
            job["companyDetails"], job["jobUrl"], job["salaryRange"],
        ])
    wb.save(OUTPUT_FILE)
    log_.info(f"Saved {len(jobs)} rows -> {OUTPUT_FILE}")

# =============================================================================
#  COOKIE INJECTION SUPPORT
# =============================================================================
# If the site is behind Cloudflare or needs a real browser session, export
# your browser cookies and set the NEWTIMES_COOKIE env var:
#   export NEWTIMES_COOKIE="_session=abc123; cf_clearance=xyz"
# The scraper will inject these into the session before any requests.

def _inject_cookies():
    raw = os.environ.get("NEWTIMES_COOKIE", "").strip()
    if not raw:
        return
    for pair in raw.split(";"):
        pair = pair.strip()
        if "=" in pair:
            k, v = pair.split("=", 1)
            SESSION.cookies.set(k.strip(), v.strip(), domain="jobs.newtimes.co.rw")
    log_.info(f"Injected {len(raw.split(';'))} cookie(s) from NEWTIMES_COOKIE env var")

# =============================================================================
#  MAIN
# =============================================================================

def main():
    start_time = datetime.now()

    # Inject any manually supplied cookies before any network calls
    _inject_cookies()

    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  NEWTIMES RWANDA SCRAPER + MISTRAL PARAPHRASE + WORDPRESS POSTING"))
    print(C_HEADER("=" * 80))
    print(f"  Source          : {BASE_URL}")
    print(f"  Archive URL     : {JOBS_ARCHIVE_URL}")
    print(f"  Category        : {SCRAPE_CATEGORY}")
    print(f"  Public-apply    : {'✅ enforced (flag others)' if REQUIRE_PUBLIC_APPLY else '❌ off (post all)'}")
    print(f"  Max new jobs    : {'unlimited' if not MAX_JOBS else MAX_JOBS}")
    print(f"  Max pages       : {MAX_PAGES}")
    print(f"  Paraphrase      : {'✅ enabled' if (ENABLE_PARAPHRASE and MISTRAL_API_KEY) else '❌ disabled'}")
    print(f"  WordPress post  : {'✅ enabled' if (WP_USER and WP_PASSWORD) else '❌ disabled'}")
    print(f"  Excel export    : {'✅ enabled' if _XLSX_AVAILABLE else '❌ disabled (pip install pandas openpyxl)'}")
    print(f"  NLP gating      : {'✅' if _NLP_AVAILABLE else '⚠️  no sentence-transformers / language-tool'}")
    print(f"  Cookie inject   : {'✅ ' + os.environ.get('NEWTIMES_COOKIE','')[:40] if os.environ.get('NEWTIMES_COOKIE') else '❌ none set'}")
    print(f"  Started         : {start_time.strftime('%Y-%m-%d %H:%M:%S')}")
    print(C_HEADER("=" * 80))

    _init_tracker()
    _init_flagged()
    processed_ids, processed_urls = load_processed_ids()
    print(f"  Tracker loaded: {len(processed_ids)} previously processed job IDs")

    try:
        job_links = collect_job_links(JOBS_ARCHIVE_URL, SCRAPE_CATEGORY, MAX_PAGES)
    except Exception as e:
        log(C_RED(f"  FATAL: could not collect job links: {e}"))
        return

    if not job_links:
        log(C_RED("  No job links found — nothing to do."))
        log(C_DIM("  If you got 403 errors, try:"))
        log(C_DIM("    1. export NEWTIMES_COOKIE='_session=<value>; cf_clearance=<value>'"))
        log(C_DIM("       (copy cookies from browser DevTools → Network → any request to jobs.newtimes.co.rw)"))
        log(C_DIM("    2. Set NEWTIMES_CATEGORY=all to broaden the search"))
        return
    print(C_GREEN(f"\n  Found {len(job_links)} job detail page(s) to process.\n"))

    jobs_out = []
    seen_content = set()
    posted_count = 0
    flagged_count = 0
    dup_count = 0
    errors = 0
    scraped = 0

    for link in job_links:
        if link in processed_urls:
            dup_count += 1
            log(C_DIM(f"  Already processed (tracker) — skipped: {link}"))
            continue

        try:
            raw_job = scrape_job_detail(link)
            scraped += 1
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR scraping {link} : {e}"))
            time.sleep(REQUEST_DELAY)
            continue

        try:
            status, job = process_job(raw_job, processed_ids, processed_urls, seen_content)
        except Exception as e:
            errors += 1
            log(C_RED(f"  ERROR processing '{raw_job.get('title','')}' : {e}"))
            continue

        if status == "duplicate":
            dup_count += 1
            time.sleep(REQUEST_DELAY)
            continue
        if status == "flagged":
            flagged_count += 1
            time.sleep(REQUEST_DELAY)
            continue

        jobs_out.append(job)
        print_job_verbose(len(jobs_out), job)

        print(C_BLUE("\n  Posting to WordPress …"))
        wp_id, wp_url = post_job_to_wordpress(job)
        if wp_id:
            mark_posted(job["_jobId"], wp_id, wp_url or "")
            posted_count += 1
            print(C_GREEN(f"  WP ID={wp_id}  {wp_url}"))
        else:
            mark_failed(job["_jobId"], "wp_post_failed_or_skipped")
            print(C_RED("  WordPress post failed / skipped"))

        if len(jobs_out) % 25 == 0:
            _save_excel(jobs_out)

        if MAX_JOBS and len(jobs_out) >= MAX_JOBS:
            log(f"\nMAX_JOBS limit ({MAX_JOBS}) reached, stopping.")
            break

        time.sleep(REQUEST_DELAY)

    _save_excel(jobs_out)

    end_time = datetime.now()
    duration = (end_time - start_time).total_seconds() / 60.0
    print()
    print(C_HEADER("=" * 80))
    print(C_HEADER("  SCRAPE COMPLETE"))
    print(C_HEADER("=" * 80))
    print(f"  {C_LABEL('Job links found')}           : {len(job_links)}")
    print(f"  {C_LABEL('Detail pages scraped')}      : {scraped}")
    print(f"  {C_LABEL('New jobs processed')}        : {C_GREEN(str(len(jobs_out)))}")
    print(f"  {C_LABEL('Posted to WordPress')}       : {C_GREEN(str(posted_count))}")
    print(f"  {C_LABEL('Flagged (no public apply)')} : {flagged_count}")
    print(f"  {C_LABEL('Duplicates skipped')}        : {dup_count}")
    print(f"  {C_LABEL('Errors')}                    : {C_RED(str(errors)) if errors else '0'}")
    print(f"  {C_LABEL('Duration')}                  : ~{duration:.1f} min")
    print(f"  {C_LABEL('Output file')}               : {OUTPUT_FILE}")
    print(f"  {C_LABEL('Tracker file')}              : {PROCESSED_IDS_FILE}")
    print(f"  {C_LABEL('Flagged file')}              : {FLAGGED_FILE}")

    if jobs_out:
        with_apply = sum(1 for j in jobs_out if j.get("application"))
        with_email = sum(1 for j in jobs_out if "@" in (j.get("application") or ""))
        with_url   = sum(1 for j in jobs_out
                         if j.get("application") and j.get("_apply_method") == "external_url")
        with_phys  = sum(1 for j in jobs_out
                         if j.get("_apply_method") == "physical_address")
        print(f"\n  {C_LABEL('Application methods:')}")
        print(f"    Email found       : {with_email}")
        print(f"    External URL      : {with_url}")
        print(f"    Physical address  : {with_phys}")
        print(f"    Total with apply  : {with_apply}")

        para_count = sum(1 for j in jobs_out if j.get("jobTitle") != j.get("originalTitle"))
        print(f"\n  {C_LABEL('Paraphrased titles')} : {para_count}/{len(jobs_out)}")

        with_deadline = sum(1 for j in jobs_out if j.get("deadline"))
        print(f"  {C_LABEL('Deadline captured')}  : {with_deadline}/{len(jobs_out)}")

    print(C_HEADER("=" * 80))


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
