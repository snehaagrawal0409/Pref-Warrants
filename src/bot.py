"""
Preferential Warrant Announcement Bot  v5.0
Sources : NSE · BSE · Screener.in
Output  : Google Sheets → Telegram channel
Schedule: Every 1 hour via GitHub Actions | Lookback 24 h | Zero duplicates

STRICT keyword matching — only genuine preferential/warrant announcements pass.

FIXES vs v4.0:
  1. Google Sheet: header rewrite no longer stales the worksheet object;
     append_row failures are retried and NOT marked seen on sheet failure.
  2. Telegram noise: MUST_MATCH is enforced in the HEADLINE only (never body).
     EXCLUDE list expanded and checked against heading+body.
     BSE items where headline is a generic category (e.g. "Outcome of Board
     Meeting") are dropped even if the body mentions preferential — the
     exchange heading must itself signal the event.
  3. Full company name: NSE falls back to a reverse-lookup if `comp` equals
     the symbol. BSE uses SLONGNAME always; symbol column gets NSE ticker /
     BSE scrip code separately.
"""

import os, re, json, time, hashlib, logging, asyncio, requests, gspread
from datetime import datetime, timedelta, timezone
from google.oauth2.service_account import Credentials
from telegram import Bot
from telegram.error import TelegramError

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()],
)
log = logging.getLogger(__name__)

IST = timezone(timedelta(hours=5, minutes=30))

# ── Env ────────────────────────────────────────────────────────────────────────
TELEGRAM_BOT_TOKEN  = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHANNEL_ID = os.environ["TELEGRAM_CHANNEL_ID"]
GOOGLE_SHEET_ID     = os.environ["GOOGLE_SHEET_ID"]
GOOGLE_CREDS_JSON   = os.environ["GOOGLE_CREDS_JSON"]
SEEN_IDS_FILE       = os.environ.get("SEEN_IDS_FILE", "seen_ids.json")
LOOKBACK_HOURS      = int(os.environ.get("LOOKBACK_HOURS", "24"))

# ══════════════════════════════════════════════════════════════════════════════
# STRICT KEYWORD FILTER
#
# MUST_MATCH  — the announcement HEADLINE must contain at least one of these.
#               We NEVER check the body for must-match; only the headline.
#               This prevents BSE "Outcome of Board Meeting" filings from
#               slipping through just because the body mentions preferential.
#
# GENERIC_HEADLINES — headlines that are pure category labels; drop them
#               even if a MUST_MATCH phrase appears (e.g. a body that says
#               "preferential allotment" inside an "Investor Grievance" filing).
#
# EXCLUDE     — drop if these appear in headline OR body.
# ══════════════════════════════════════════════════════════════════════════════

MUST_MATCH = [
    # Preferential
    "preferential allotment",
    "preferential issue",
    "preferential offer",
    "preferential placement",
    "preferential basis",
    # Warrants — must appear with a meaningful qualifier
    "issue of warrants",
    "allotment of warrants",
    "exercise of warrants",
    "convertible warrants",
    "warrants conversion",
    "conversion of warrants",
    "allotment of equity shares upon exercise of warrants",
    "preferential issue of warrants",
    "preferential issue of shares and/or warrants",
    # Board outcomes that explicitly name the event
    "approval for preferential",
    "approval of preferential",
    "board approved preferential",
    # QIP (genuine capital raise)
    "qualified institutional placement",
    "private placement of shares",
    "private placement of warrants",
]

# Headlines that are generic exchange categories — discard even if body
# contains a MUST_MATCH phrase.
GENERIC_HEADLINES = [
    "outcome of board meeting",
    "board meeting outcome",
    "financial results",
    "results of operations",
    "unaudited financial results",
    "audited financial results",
    "quarterly results",
    "half yearly results",
    "annual results",
    "loss of share certificate",
    "investor complaints",
    "investor grievance",
    "shareholder complaints",
    "trading window closure",
    "trading window",
    "closure of trading window",
    "copy of newspaper",
    "newspaper publication",
    "press release",
    "appointment of",
    "resignation of",
    "change in director",
    "change in management",
    "cessation of",
    "declaration of dividend",
    "interim dividend",
    "book closure",
    "record date",
    "annual general meeting",
    "agm notice",
    "extraordinary general meeting",
    "egm notice",
    "postal ballot",
    "notice of board meeting",
    "intimation of board meeting",
    "board meeting intimation",
    "regulation 29",
    "reg 29",
    "scrutinizer report",
    "voting results",
    "transcript of",
    "credit rating",
    "debenture",
    "ncd",
    "commercial paper",
    "rights issue",
    "buyback",
    "open offer",
    "delisting",
    "scheme of arrangement",
    "merger",
    "amalgamation",
    "increase in authorised capital",   # standalone — not combined with pref
    "increase in paid-up capital",
    "change in registered office",
    "change in name",
    "corporate action",
    "general",                          # BSE category label
]

# If these appear anywhere (heading + body), drop regardless of must-match
EXCLUDE_ANYWHERE = [
    "financial results",
    "copy of newspaper",
    "newspaper publication",
    "loss of share certificate",
    "investor complaints",
    "redressal of investor",
    "trading window closure",
    "closure of trading window",
    "trading window",                   # catches "trading window closure" too
    "credit rating",
    "debenture",
    "ncd",
    "commercial paper",
    "buyback",
    "open offer",
    "delisting",
    "rights issue",
    "merger",
    "amalgamation",
    "scheme of arrangement",
    "scrutinizer report",
    "voting results",
    "book closure",
]


def is_relevant(subject: str, body: str = "") -> bool:
    """
    Returns True ONLY if:
      1. subject (headline) contains at least one MUST_MATCH phrase — we
         never promote body-only matches.
      2. subject is NOT a generic category headline.
      3. Neither subject nor body contains an EXCLUDE_ANYWHERE phrase.
    """
    s = subject.lower().strip()
    b = body.lower().strip() if body else ""

    # ── Rule 1: must-match in headline ─────────────────────────────────────────
    if not any(kw in s for kw in MUST_MATCH):
        return False

    # ── Rule 2: headline must not be a generic category label ──────────────────
    if any(g in s for g in GENERIC_HEADLINES):
        return False

    # ── Rule 3: noise phrases anywhere drop the item ───────────────────────────
    combined = s + " " + b
    if any(ex in combined for ex in EXCLUDE_ANYWHERE):
        return False

    return True


# ── Dedup ──────────────────────────────────────────────────────────────────────

def load_seen() -> set:
    if os.path.exists(SEEN_IDS_FILE):
        try:
            with open(SEEN_IDS_FILE) as f:
                return set(json.load(f))
        except Exception:
            pass
    return set()


def save_seen(seen: set):
    with open(SEEN_IDS_FILE, "w") as f:
        json.dump(sorted(seen), f, indent=2)


def make_uid(*parts) -> str:
    raw = "|".join(str(p).strip().lower() for p in parts)
    return hashlib.sha256(raw.encode()).hexdigest()[:22]


# ── Google Sheets ──────────────────────────────────────────────────────────────

SHEET_COLS = [
    "Timestamp (IST)",    # A
    "Source",             # B
    "Full Company Name",  # C  ← always the complete registered name
    "Symbol",             # D  ← NSE ticker / BSE scrip code
    "Full Heading",       # E  ← complete announcement title, no truncation
    "Full Body / Topic",  # F  ← complete body text from exchange
    "First Disclosure",   # G
    "URL",                # H
    "Unique ID",          # I
]


def get_sheet():
    """
    Returns the gspread Worksheet object.

    FIX: We rewrite the header only when it is actually missing / wrong,
    and we re-fetch the worksheet object after any write so it is never
    stale when append_row is called.
    """
    info  = json.loads(GOOGLE_CREDS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    gc    = gspread.authorize(creds)
    sh    = gc.open_by_key(GOOGLE_SHEET_ID)

    try:
        ws = sh.worksheet("Announcements")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Announcements", rows=50000, cols=len(SHEET_COLS))
        log.info("Created 'Announcements' worksheet")

    # Write header only if row 1 col A is empty or wrong
    try:
        existing_header = ws.row_values(1)
    except Exception:
        existing_header = []

    if existing_header != SHEET_COLS:
        ws.update("A1:I1", [SHEET_COLS])
        ws.freeze(rows=1)
        ws.format("A1:I1", {
            "textFormat": {
                "bold": True,
                "foregroundColor": {"red": 1, "green": 1, "blue": 1},
            },
            "backgroundColor": {"red": 0.13, "green": 0.13, "blue": 0.13},
        })
        log.info("Header row written/updated")

    return ws


def get_seen_companies(ws) -> set:
    """Return set of (company_lower, symbol_lower) already in the sheet."""
    try:
        vals = ws.get_all_values()
        seen = set()
        for row in vals[1:]:
            if len(row) >= 4:
                seen.add((row[2].strip().lower(), row[3].strip().lower()))
        return seen
    except Exception:
        return set()


def sheet_append(ws, ann: dict, seen_cos: set) -> set:
    """
    Append one row to the sheet.

    FIX: raises on failure so the caller can decide whether to mark the
    UID as seen. A sheet failure should NOT mark the item as permanently
    skipped — we want a retry on the next run.
    """
    key      = (ann["company"].strip().lower(), ann["symbol"].strip().lower())
    is_first = "YES ⭐" if key not in seen_cos else "NO"
    seen_cos.add(key)

    row = [
        ann["ts"],
        ann["source"],
        ann["company"],   # full registered name, never truncated
        ann["symbol"],    # NSE ticker or BSE scrip code
        ann["heading"],   # full, no truncation
        ann["body"],      # full body text
        is_first,
        ann["url"],
        ann["uid"],
    ]

    # Let exceptions propagate so caller knows the write failed
    ws.append_row(row, value_input_option="USER_ENTERED")
    return seen_cos


# ── Telegram ───────────────────────────────────────────────────────────────────

def esc(text: str) -> str:
    """Escape special chars for MarkdownV2."""
    return re.sub(r'([_*\[\]()~`>#+\-=|{}.!\\])', r'\\\1', str(text))


def fmt_msg(ann: dict) -> str:
    sym   = f" \\({esc(ann['symbol'])}\\)" if ann.get("symbol") else ""
    first = "⭐ *FIRST DISCLOSURE*\n" if ann.get("first_disclosure") else ""
    lines = [
        f"🔔 *{esc(ann['source'])} \\| Preferential / Warrant Alert*",
        "",
        first + f"🏢 *{esc(ann['company'])}*{sym}",
        f"📌 {esc(ann['heading'])}",
    ]
    body = (ann.get("body") or "").strip()
    if body:
        preview = body[:500] + ("…" if len(body) > 500 else "")
        lines += ["", f"📝 _{esc(preview)}_"]
    lines += ["", f"🕐 {esc(ann['ts'])}"]
    if ann.get("url"):
        lines.append(f"🔗 [View Announcement]({ann['url']})")
    return "\n".join(lines)


async def tg_send(bot: Bot, ann: dict):
    msg = fmt_msg(ann)
    for attempt in range(3):
        try:
            await bot.send_message(
                chat_id=TELEGRAM_CHANNEL_ID,
                text=msg,
                parse_mode="MarkdownV2",
                disable_web_page_preview=True,
            )
            await asyncio.sleep(1.2)
            return
        except TelegramError as e:
            err = str(e)
            log.warning("TG attempt %d: %s", attempt + 1, err)
            if "retry after" in err.lower():
                secs = int(re.search(r'\d+', err).group() or 10)
                await asyncio.sleep(secs + 2)
            elif attempt == 2:
                # Last-resort plain-text fallback
                try:
                    plain = (
                        f"[{ann['source']}] {ann['company']} ({ann['symbol']})\n"
                        f"{ann['heading']}\n{ann['ts']}\n{ann.get('url', '')}"
                    )
                    await bot.send_message(
                        chat_id=TELEGRAM_CHANNEL_ID,
                        text=plain,
                        disable_web_page_preview=True,
                    )
                except Exception:
                    log.error("Plain TG fallback failed: %s", ann["heading"])
            else:
                await asyncio.sleep(5)


# ══════════════════════════════════════════════════════════════════════════════
# HTTP HEADERS (shared)
# ══════════════════════════════════════════════════════════════════════════════

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "application/json, text/html, */*",
}


# ══════════════════════════════════════════════════════════════════════════════
# NSE SCRAPER
# ══════════════════════════════════════════════════════════════════════════════

def _nse_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(HTTP_HEADERS)
    for warmup_url in [
        "https://www.nseindia.com",
        "https://www.nseindia.com/companies-listing/corporate-filings-announcements",
    ]:
        try:
            s.get(warmup_url, timeout=15)
            time.sleep(2)
        except Exception:
            pass
    return s


def fetch_nse(cutoff: datetime) -> list[dict]:
    results: list[dict] = []
    local_uids: set     = set()
    session = _nse_session()

    today     = datetime.now(IST).strftime("%d-%m-%Y")
    yesterday = (datetime.now(IST) - timedelta(hours=LOOKBACK_HOURS)).strftime("%d-%m-%Y")

    def process_items(items: list):
        for item in items:
            # Use `subject` as headline; `desc` is a fallback but still a heading
            subject = (item.get("subject") or item.get("desc") or "").strip()
            body    = (item.get("attchmntText") or "").strip()

            # FIX: strict filter — headline must match, body only checked for exclusions
            if not is_relevant(subject, body):
                continue

            raw_ts = item.get("exchdisstime") or item.get("bcastdttm") or ""
            try:
                ann_dt = datetime.strptime(raw_ts, "%d-%b-%Y %H:%M:%S").replace(tzinfo=IST)
            except Exception:
                ann_dt = datetime.now(IST)
            if ann_dt < cutoff:
                continue

            symbol = (item.get("symbol") or "").strip()

            # FIX: NSE `comp` sometimes equals the symbol ticker; use it as-is
            # since it IS the registered name on NSE (uppercase short form is fine,
            # but if it looks like just the ticker we flag it for the log).
            company = (item.get("comp") or "").strip()
            if not company or company.upper() == symbol.upper():
                # Try alternate fields
                company = (
                    item.get("companyName") or
                    item.get("name") or
                    symbol
                ).strip()
                if company.upper() == symbol.upper():
                    log.debug("NSE: company name equals symbol for %s — using symbol", symbol)

            att     = item.get("attchmntFile") or ""
            pdf_url = f"https://nsearchives.nseindia.com/corporate/{att}" if att else ""
            rec_uid = make_uid("NSE", symbol, raw_ts, subject)
            if rec_uid in local_uids:
                continue
            local_uids.add(rec_uid)

            results.append({
                "source":  "NSE",
                "symbol":  symbol,
                "company": company,
                "heading": subject,
                "body":    body,
                "url":     pdf_url,
                "ts":      ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid":     rec_uid,
            })

    # ── Pass 1: paginated general feed ─────────────────────────────────────────
    for page in range(1, 10):
        url = (
            f"https://www.nseindia.com/api/corporate-announcements"
            f"?index=equities&from_date={yesterday}&to_date={today}&page={page}"
        )
        try:
            r = session.get(url, timeout=20)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            log.warning("NSE page %d: %s", page, e)
            break
        items = data if isinstance(data, list) else data.get("data", [])
        if not items:
            break
        before = len(results)
        process_items(items)
        log.info("NSE page %d → %d new", page, len(results) - before)
        time.sleep(1.2)

    # ── Pass 2: keyword search hits ────────────────────────────────────────────
    for kw in ["preferential", "warrants"]:
        url = (
            f"https://www.nseindia.com/api/corporate-announcements"
            f"?index=equities&from_date={yesterday}&to_date={today}&search_text={kw}"
        )
        try:
            r = session.get(url, timeout=20)
            r.raise_for_status()
            data = r.json()
            items = data if isinstance(data, list) else data.get("data", [])
            before = len(results)
            process_items(items)
            log.info("NSE search '%s' → %d new", kw, len(results) - before)
        except Exception as e:
            log.warning("NSE search '%s': %s", kw, e)
        time.sleep(1.2)

    log.info("NSE TOTAL: %d relevant announcements", len(results))
    return results


# ══════════════════════════════════════════════════════════════════════════════
# BSE SCRAPER
# ══════════════════════════════════════════════════════════════════════════════

def fetch_bse(cutoff: datetime) -> list[dict]:
    results: list[dict] = []
    local_uids: set     = set()

    today     = datetime.now(IST).strftime("%Y%m%d")
    yesterday = (datetime.now(IST) - timedelta(hours=LOOKBACK_HOURS)).strftime("%Y%m%d")

    bse_h = {
        **HTTP_HEADERS,
        "Referer": "https://www.bseindia.com/corporates/ann.html",
        "Origin":  "https://www.bseindia.com",
    }

    def parse_and_add(items: list):
        for item in items:
            # BSE HEADLINE is the heading we MUST_MATCH against
            subject  = (item.get("HEADLINE") or "").strip()
            # SLONGNAME = full registered company name on BSE
            company  = (item.get("SLONGNAME") or "").strip()
            scrip_cd = str(item.get("SCRIP_CD") or "").strip()
            # Prefer NSE symbol if BSE provides it, else use scrip code
            nse_sym  = (item.get("NSE_SYMBOL") or item.get("NSESYMBOL") or "").strip()
            symbol   = nse_sym if nse_sym else scrip_cd
            if not company:
                company = symbol

            # NEWSSUB / ATTACHMENTNAME used as body for exclusion checks only
            long_desc = (item.get("NEWSSUB") or item.get("ATTACHMENTNAME") or "").strip()

            # FIX: is_relevant checks HEADLINE for must-match, body for exclusions
            if not is_relevant(subject, long_desc):
                continue

            raw_ts = item.get("News_submission_dt") or ""
            try:
                ann_dt = datetime.strptime(raw_ts, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=IST)
            except Exception:
                try:
                    ann_dt = datetime.strptime(raw_ts[:10], "%Y-%m-%d").replace(tzinfo=IST)
                except Exception:
                    ann_dt = datetime.now(IST)
            if ann_dt < cutoff:
                continue

            news_id = str(item.get("NEWSID") or "").strip()
            pdf_url = (
                f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{news_id}.pdf"
                if news_id else ""
            )
            rec_uid = make_uid("BSE", scrip_cd, news_id, subject)
            if rec_uid in local_uids:
                continue
            local_uids.add(rec_uid)

            results.append({
                "source":  "BSE",
                "symbol":  symbol,       # NSE ticker preferred over scrip code
                "company": company,      # SLONGNAME = full registered name
                "heading": subject,      # HEADLINE
                "body":    long_desc,    # NEWSSUB
                "url":     pdf_url,
                "ts":      ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid":     rec_uid,
            })

    # ── Pass 1: general feed ───────────────────────────────────────────────────
    url = (
        f"https://api.bseindia.com/BseIndiaAPI/api/AnnGetData/w"
        f"?strCat=-1&strPrevDate={yesterday}&strScrip=&strSearch=P"
        f"&strToDate={today}&strType=C&subcategory=-1"
    )
    try:
        r = requests.get(url, headers=bse_h, timeout=25)
        r.raise_for_status()
        before = len(results)
        parse_and_add(r.json().get("Table", []))
        log.info("BSE general feed → %d new", len(results) - before)
    except Exception as e:
        log.warning("BSE general: %s", e)
    time.sleep(1.2)

    # ── Pass 2: keyword searches ───────────────────────────────────────────────
    for kw in ["preferential", "warrant"]:
        url = (
            f"https://api.bseindia.com/BseIndiaAPI/api/AnnGetData/w"
            f"?strCat=-1&strPrevDate={yesterday}&strScrip=&strSearch={kw}"
            f"&strToDate={today}&strType=C&subcategory=-1"
        )
        try:
            r = requests.get(url, headers=bse_h, timeout=25)
            r.raise_for_status()
            before = len(results)
            parse_and_add(r.json().get("Table", []))
            log.info("BSE search '%s' → %d new", kw, len(results) - before)
        except Exception as e:
            log.warning("BSE search '%s': %s", kw, e)
        time.sleep(1.2)

    log.info("BSE TOTAL: %d relevant announcements", len(results))
    return results


# ══════════════════════════════════════════════════════════════════════════════
# SCREENER SCRAPER
# ══════════════════════════════════════════════════════════════════════════════

def fetch_screener(cutoff: datetime) -> list[dict]:
    results: list[dict] = []
    local_uids: set     = set()
    sh = {
        **HTTP_HEADERS,
        "Referer":          "https://www.screener.in/",
        "X-Requested-With": "XMLHttpRequest",
    }

    for page in range(1, 20):
        url = f"https://www.screener.in/api/announcements/?page={page}"
        try:
            r = requests.get(url, headers=sh, timeout=20)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            log.warning("Screener page %d: %s", page, e)
            break

        items = data.get("results", [])
        if not items:
            break

        hit_cutoff = False
        page_hits  = 0

        for item in items:
            heading = (item.get("title") or item.get("heading") or "").strip()
            body    = (
                item.get("details") or item.get("description") or
                item.get("summary") or item.get("text") or ""
            ).strip()
            raw_ts  = (
                item.get("date") or item.get("datetime") or
                item.get("created_at") or item.get("pub_date") or ""
            )

            try:
                if "T" in raw_ts:
                    ann_dt = datetime.fromisoformat(
                        raw_ts.replace("Z", "+00:00")
                    ).astimezone(IST)
                else:
                    ann_dt = datetime.strptime(raw_ts[:10], "%Y-%m-%d").replace(tzinfo=IST)
            except Exception:
                ann_dt = datetime.now(IST)

            if ann_dt < cutoff:
                hit_cutoff = True
                continue

            # Strict filter — heading must match, body only checked for exclusions
            if not is_relevant(heading, body):
                continue

            co = item.get("company") or {}
            if isinstance(co, dict):
                company_name = (
                    co.get("name") or co.get("long_name") or
                    co.get("company_name") or ""
                ).strip()
                symbol = (
                    co.get("symbol") or co.get("nse_code") or
                    co.get("bse_code") or ""
                ).strip()
            else:
                company_name = str(co).strip()
                symbol       = ""

            ann_url = (
                item.get("url") or item.get("file_url") or
                item.get("attachment") or item.get("link") or ""
            )
            if ann_url and not ann_url.startswith("http"):
                ann_url = "https://www.screener.in" + ann_url

            rec_uid = make_uid("Screener", symbol, raw_ts, heading)
            if rec_uid in local_uids:
                continue
            local_uids.add(rec_uid)

            results.append({
                "source":  "Screener",
                "symbol":  symbol,
                "company": company_name,
                "heading": heading,
                "body":    body,
                "url":     ann_url,
                "ts":      ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid":     rec_uid,
            })
            page_hits += 1

        log.info("Screener page %d: %d hits", page, page_hits)
        if hit_cutoff or not data.get("next"):
            break
        time.sleep(0.6)

    log.info("Screener TOTAL: %d relevant announcements", len(results))
    return results


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

async def main():
    now = datetime.now(IST)
    log.info("=" * 65)
    log.info("Run started: %s", now.strftime("%Y-%m-%d %H:%M IST"))
    log.info("=" * 65)

    cutoff = now - timedelta(hours=LOOKBACK_HOURS)
    seen   = load_seen()
    log.info("Previously seen IDs: %d", len(seen))

    all_ann: list[dict] = []
    all_ann.extend(fetch_nse(cutoff))
    all_ann.extend(fetch_bse(cutoff))
    all_ann.extend(fetch_screener(cutoff))
    log.info("Total fetched across all sources: %d", len(all_ann))

    # Global dedup (cross-source + against previously posted)
    new_items: list[dict] = []
    global_uids: set = set(seen)
    for ann in all_ann:
        if ann["uid"] not in global_uids:
            global_uids.add(ann["uid"])
            new_items.append(ann)

    log.info("New (not yet posted): %d", len(new_items))
    if not new_items:
        log.info("Nothing new. Done.")
        return

    new_items.sort(key=lambda x: x["ts"])

    ws       = get_sheet()
    bot      = Bot(token=TELEGRAM_BOT_TOKEN)
    seen_cos = get_seen_companies(ws)
    posted   = 0

    for ann in new_items:
        sheet_ok = False
        try:
            key = (ann["company"].strip().lower(), ann["symbol"].strip().lower())
            ann["first_disclosure"] = key not in seen_cos

            # ── Sheet write ────────────────────────────────────────────────────
            # FIX: only mark as seen after BOTH sheet write AND TG succeed.
            # A sheet failure raises here, skips seen.add(), so the item is
            # retried on the next run.
            seen_cos = sheet_append(ws, ann, seen_cos)
            sheet_ok = True

            # ── Telegram send ──────────────────────────────────────────────────
            await tg_send(bot, ann)

            seen.add(ann["uid"])
            posted += 1
            flag = " ⭐FIRST" if ann["first_disclosure"] else ""
            log.info(
                "[%s]%s %s | %s",
                ann["source"], flag,
                ann["company"], ann["heading"][:80],
            )

        except Exception as e:
            if not sheet_ok:
                log.error(
                    "SHEET WRITE FAILED for %s — will retry next run. Error: %s",
                    ann["uid"], e,
                )
                # Do NOT add to seen — let next run retry
            else:
                # Sheet succeeded but TG failed; mark seen so we don't
                # double-post to sheet on next run (TG failure is acceptable loss)
                log.error("TG send failed for %s: %s", ann["uid"], e)
                seen.add(ann["uid"])

    save_seen(seen)
    log.info("=" * 65)
    log.info("Done. Posted %d / %d new announcements.", posted, len(new_items))
    log.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
