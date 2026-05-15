"""
Preferential Warrant Announcement Bot  v3.0
Sources : NSE (3 endpoints) · BSE (category + search) · Screener.in
Output  : Google Sheets (full heading/body) → Telegram channel
Schedule: Every 1 hour via GitHub Actions | Lookback 24 h | Zero duplicates
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

# ── Keywords ───────────────────────────────────────────────────────────────────
# Any announcement whose subject OR body contains ANY of these (case-insensitive)
# will be captured.  Sorted longest-first so more-specific phrases win.
KEYWORDS = sorted([
    "allotment of equity shares upon exercise of warrants",
    "conversion of warrants into equity shares",
    "conversion of warrants into equity",
    "preferential issue of warrants",
    "preferential issue of shares",
    "qualified institutional placement",
    "initial subscription amount",
    "warrants conversion",
    "allotment of warrants",
    "exercise of warrants",
    "convertible warrants",
    "preferential allotment",
    "preferential placement",
    "preferential issue",
    "preferential basis",
    "private placement",
    "in-principle preferential",
    "in-principle approval",
    "promoter allotment",
    "non-promoter allotment",
    "issue of warrants",
    "balance consideration",
    "further allotment",
    "tranche allotment",
    "capital raising",
    "fund raising",
    "lock-in period",
    "listing stage",
    "issue price",
    "preferential",
    "warrants",
    "warrant",
], key=len, reverse=True)

HTTP_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Accept": "application/json, text/html, */*",
}


def is_relevant(text: str) -> bool:
    if not text:
        return False
    t = text.lower()
    return any(kw in t for kw in KEYWORDS)


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


def uid(*parts) -> str:
    raw = "|".join(str(p).strip().lower() for p in parts)
    return hashlib.sha256(raw.encode()).hexdigest()[:22]


# ── Google Sheets ──────────────────────────────────────────────────────────────

SHEET_COLS = [
    "Timestamp (IST)",   # A
    "Source",            # B
    "Company",           # C
    "Symbol",            # D
    "Full Heading",      # E  ← full announcement title (no truncation)
    "Full Body / Topic", # F  ← full text body (Screener summary / NSE body)
    "First Disclosure",  # G
    "URL",               # H
    "Unique ID",         # I
]


def get_sheet():
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
        # Always ensure headers are correct (handles old/missing columns)
        ws.update("A1:I1", [SHEET_COLS])
        ws.freeze(rows=1)
        ws.format("A1:I1", {"textFormat": {"bold": True},
                             "backgroundColor": {"red": 0.15, "green": 0.15, "blue": 0.15}})
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Announcements", rows=20000, cols=len(SHEET_COLS))
        ws.update("A1:I1", [SHEET_COLS])
        ws.freeze(rows=1)
        ws.format("A1:I1", {"textFormat": {"bold": True},
                             "backgroundColor": {"red": 0.15, "green": 0.15, "blue": 0.15}})
        log.info("Created 'Announcements' worksheet")
    return ws


def get_seen_companies(ws) -> set:
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
    key      = (ann["company"].strip().lower(), ann["symbol"].strip().lower())
    is_first = "YES ⭐" if key not in seen_cos else "NO"
    seen_cos.add(key)
    ws.append_row([
        ann["ts"], ann["source"], ann["company"], ann["symbol"],
        ann["heading"],   # FULL heading — no [:300] cap
        ann["body"],      # FULL body text
        is_first,
        ann["url"],
        ann["uid"],
    ], value_input_option="USER_ENTERED")
    return seen_cos


# ── Telegram ───────────────────────────────────────────────────────────────────

def esc(text: str) -> str:
    return re.sub(r'([_*\[\]()~`>#+\-=|{}.!\\])', r'\\\1', str(text))


def fmt_msg(ann: dict) -> str:
    sym   = f" \\({esc(ann['symbol'])}\\)" if ann.get("symbol") else ""
    first = "⭐ *FIRST DISCLOSURE*\n" if ann.get("first_disclosure") else ""
    lines = [
        f"🔔 *{esc(ann['source'])} \\| Preferential Warrant Alert*",
        "",
        first + f"🏢 *{esc(ann['company'])}*{sym}",
        f"📌 {esc(ann['heading'])}",
    ]
    body = ann.get("body", "")
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
                try:  # plain-text fallback
                    plain = (
                        f"[{ann['source']}] {ann['company']} ({ann['symbol']})\n"
                        f"{ann['heading']}\n{ann['ts']}\n{ann.get('url','')}"
                    )
                    await bot.send_message(chat_id=TELEGRAM_CHANNEL_ID,
                                           text=plain, disable_web_page_preview=True)
                except Exception:
                    log.error("Plain TG fallback failed: %s", ann["heading"])
            else:
                await asyncio.sleep(5)


# ══════════════════════════════════════════════════════════════════════════════
# NSE  — three endpoints to maximise coverage
# ══════════════════════════════════════════════════════════════════════════════

def _nse_session() -> requests.Session:
    """Return a warmed-up NSE session with cookies."""
    s = requests.Session()
    s.headers.update(HTTP_HEADERS)
    for url in [
        "https://www.nseindia.com",
        "https://www.nseindia.com/companies-listing/corporate-filings-announcements",
    ]:
        try:
            s.get(url, timeout=15)
            time.sleep(1.5)
        except Exception:
            pass
    return s


def fetch_nse(cutoff: datetime) -> list[dict]:
    results = []
    seen_ids_local: set = set()
    session = _nse_session()

    today     = datetime.now(IST).strftime("%d-%m-%Y")
    yesterday = (datetime.now(IST) - timedelta(hours=LOOKBACK_HOURS)).strftime("%d-%m-%Y")

    # ── Endpoint 1: general corporate announcements (all categories) ───────────
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
            log.warning("NSE ep1 page %d: %s", page, e)
            break

        items = data if isinstance(data, list) else data.get("data", [])
        if not items:
            break

        added = 0
        for item in items:
            subject = (item.get("subject") or item.get("desc") or "").strip()
            body    = (item.get("attchmntText") or "").strip()
            if not is_relevant(subject) and not is_relevant(body):
                continue

            raw_ts = item.get("exchdisstime") or item.get("bcastdttm") or ""
            try:
                ann_dt = datetime.strptime(raw_ts, "%d-%b-%Y %H:%M:%S").replace(tzinfo=IST)
            except Exception:
                ann_dt = datetime.now(IST)
            if ann_dt < cutoff:
                continue

            symbol  = (item.get("symbol") or "").strip()
            company = (item.get("comp") or symbol).strip()
            att     = item.get("attchmntFile") or ""
            pdf_url = f"https://nsearchives.nseindia.com/corporate/{att}" if att else ""
            rec_uid = uid("NSE", symbol, raw_ts, subject)
            if rec_uid in seen_ids_local:
                continue
            seen_ids_local.add(rec_uid)

            results.append({
                "source": "NSE", "symbol": symbol, "company": company,
                "heading": subject,   # FULL, no truncation
                "body":    body,      # FULL body text
                "url": pdf_url,
                "ts":  ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid": rec_uid,
            })
            added += 1

        log.info("NSE ep1 page %d: %d relevant", page, added)
        time.sleep(1)

    # ── Endpoint 2: search by keyword "preferential" ───────────────────────────
    for kw_search in ["preferential", "warrant"]:
        search_url = (
            f"https://www.nseindia.com/api/corporate-announcements"
            f"?index=equities&from_date={yesterday}&to_date={today}"
            f"&search_text={kw_search}"
        )
        try:
            r = session.get(search_url, timeout=20)
            r.raise_for_status()
            items = r.json()
            if not isinstance(items, list):
                items = items.get("data", [])
        except Exception as e:
            log.warning("NSE keyword search '%s': %s", kw_search, e)
            items = []

        for item in items:
            subject = (item.get("subject") or item.get("desc") or "").strip()
            body    = (item.get("attchmntText") or "").strip()
            raw_ts  = item.get("exchdisstime") or item.get("bcastdttm") or ""
            try:
                ann_dt = datetime.strptime(raw_ts, "%d-%b-%Y %H:%M:%S").replace(tzinfo=IST)
            except Exception:
                ann_dt = datetime.now(IST)
            if ann_dt < cutoff:
                continue

            symbol  = (item.get("symbol") or "").strip()
            company = (item.get("comp") or symbol).strip()
            att     = item.get("attchmntFile") or ""
            pdf_url = f"https://nsearchives.nseindia.com/corporate/{att}" if att else ""
            rec_uid = uid("NSE", symbol, raw_ts, subject)
            if rec_uid in seen_ids_local:
                continue
            seen_ids_local.add(rec_uid)

            results.append({
                "source": "NSE", "symbol": symbol, "company": company,
                "heading": subject, "body": body, "url": pdf_url,
                "ts":  ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid": rec_uid,
            })
        time.sleep(1)

    # ── Endpoint 3: NSE capital market / preferential category (cat=4) ─────────
    for cat in ["4", "8"]:   # 4 = Capital Market, 8 = Board Meeting (catches warrants)
        cat_url = (
            f"https://www.nseindia.com/api/corporate-announcements"
            f"?index=equities&from_date={yesterday}&to_date={today}&category={cat}"
        )
        try:
            r = session.get(cat_url, timeout=20)
            r.raise_for_status()
            items = r.json()
            if not isinstance(items, list):
                items = items.get("data", [])
        except Exception as e:
            log.warning("NSE cat %s: %s", cat, e)
            items = []

        for item in items:
            subject = (item.get("subject") or item.get("desc") or "").strip()
            body    = (item.get("attchmntText") or "").strip()
            if not is_relevant(subject) and not is_relevant(body):
                continue

            raw_ts  = item.get("exchdisstime") or item.get("bcastdttm") or ""
            try:
                ann_dt = datetime.strptime(raw_ts, "%d-%b-%Y %H:%M:%S").replace(tzinfo=IST)
            except Exception:
                ann_dt = datetime.now(IST)
            if ann_dt < cutoff:
                continue

            symbol  = (item.get("symbol") or "").strip()
            company = (item.get("comp") or symbol).strip()
            att     = item.get("attchmntFile") or ""
            pdf_url = f"https://nsearchives.nseindia.com/corporate/{att}" if att else ""
            rec_uid = uid("NSE", symbol, raw_ts, subject)
            if rec_uid in seen_ids_local:
                continue
            seen_ids_local.add(rec_uid)

            results.append({
                "source": "NSE", "symbol": symbol, "company": company,
                "heading": subject, "body": body, "url": pdf_url,
                "ts":  ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid": rec_uid,
            })
        time.sleep(1)

    log.info("NSE TOTAL: %d relevant", len(results))
    return results


# ══════════════════════════════════════════════════════════════════════════════
# BSE  — two endpoints + dedicated preferential category
# ══════════════════════════════════════════════════════════════════════════════

def fetch_bse(cutoff: datetime) -> list[dict]:
    results       = []
    seen_ids_local: set = set()
    today     = datetime.now(IST).strftime("%Y%m%d")
    yesterday = (datetime.now(IST) - timedelta(hours=LOOKBACK_HOURS)).strftime("%Y%m%d")
    bse_h = {
        **HTTP_HEADERS,
        "Referer": "https://www.bseindia.com/corporates/ann.html",
        "Origin":  "https://www.bseindia.com",
    }

    def parse_bse_item(item):
        subject  = (item.get("HEADLINE") or "").strip()
        scrip_cd = str(item.get("SCRIP_CD") or "").strip()
        company  = (item.get("SLONGNAME") or scrip_cd).strip()
        raw_ts   = item.get("News_submission_dt") or ""
        try:
            ann_dt = datetime.strptime(raw_ts, "%Y-%m-%dT%H:%M:%S").replace(tzinfo=IST)
        except Exception:
            try:
                ann_dt = datetime.strptime(raw_ts[:10], "%Y-%m-%d").replace(tzinfo=IST)
            except Exception:
                ann_dt = datetime.now(IST)
        news_id = str(item.get("NEWSID") or "").strip()
        pdf_url = (
            f"https://www.bseindia.com/xml-data/corpfiling/AttachHis/{news_id}.pdf"
            if news_id else ""
        )
        # BSE long description (not always present but capture if available)
        long_desc = (item.get("NEWSSUB") or item.get("ATTACHMENTNAME") or "").strip()
        return subject, scrip_cd, company, ann_dt, pdf_url, long_desc, news_id

    # ── Endpoint A: all categories ─────────────────────────────────────────────
    for strCat in ["-1", "4"]:  # -1 = all, 4 = Capital Issues
        api_url = (
            f"https://api.bseindia.com/BseIndiaAPI/api/AnnGetData/w"
            f"?strCat={strCat}&strPrevDate={yesterday}&strScrip=&strSearch=P"
            f"&strToDate={today}&strType=C&subcategory=-1"
        )
        try:
            r = requests.get(api_url, headers=bse_h, timeout=25)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            log.warning("BSE cat %s: %s", strCat, e)
            continue

        for item in data.get("Table", []):
            subject, scrip_cd, company, ann_dt, pdf_url, long_desc, news_id = parse_bse_item(item)
            if not is_relevant(subject) and not is_relevant(long_desc):
                continue
            if ann_dt < cutoff:
                continue
            rec_uid = uid("BSE", scrip_cd, news_id, subject)
            if rec_uid in seen_ids_local:
                continue
            seen_ids_local.add(rec_uid)
            results.append({
                "source": "BSE", "symbol": scrip_cd, "company": company,
                "heading": subject,
                "body":    long_desc,
                "url": pdf_url,
                "ts":  ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid": rec_uid,
            })
        time.sleep(1)

    # ── Endpoint B: BSE search by "preferential" / "warrant" ──────────────────
    for kw in ["preferential", "warrant"]:
        search_url = (
            f"https://api.bseindia.com/BseIndiaAPI/api/AnnGetData/w"
            f"?strCat=-1&strPrevDate={yesterday}&strScrip=&strSearch={kw}"
            f"&strToDate={today}&strType=C&subcategory=-1"
        )
        try:
            r = requests.get(search_url, headers=bse_h, timeout=25)
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            log.warning("BSE search '%s': %s", kw, e)
            continue

        for item in data.get("Table", []):
            subject, scrip_cd, company, ann_dt, pdf_url, long_desc, news_id = parse_bse_item(item)
            if ann_dt < cutoff:
                continue
            rec_uid = uid("BSE", scrip_cd, news_id, subject)
            if rec_uid in seen_ids_local:
                continue
            seen_ids_local.add(rec_uid)
            results.append({
                "source": "BSE", "symbol": scrip_cd, "company": company,
                "heading": subject,
                "body":    long_desc,
                "url": pdf_url,
                "ts":  ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid": rec_uid,
            })
        time.sleep(1)

    log.info("BSE TOTAL: %d relevant", len(results))
    return results


# ══════════════════════════════════════════════════════════════════════════════
# Screener.in  — paginate until cutoff, capture full details field
# ══════════════════════════════════════════════════════════════════════════════

def fetch_screener(cutoff: datetime) -> list[dict]:
    results = []
    seen_ids_local: set = set()
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
            # Screener provides a rich "details" field — capture in full
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

            if not is_relevant(heading) and not is_relevant(body):
                continue

            co = item.get("company") or {}
            if isinstance(co, dict):
                company_name = co.get("name") or ""
                symbol       = co.get("symbol") or co.get("bse_code") or ""
            else:
                company_name = str(co)
                symbol       = ""

            ann_url = (
                item.get("url") or item.get("file_url") or
                item.get("attachment") or item.get("link") or ""
            )
            if ann_url and not ann_url.startswith("http"):
                ann_url = "https://www.screener.in" + ann_url

            rec_uid = uid("Screener", symbol, raw_ts, heading)
            if rec_uid in seen_ids_local:
                continue
            seen_ids_local.add(rec_uid)

            results.append({
                "source": "Screener", "symbol": symbol, "company": company_name,
                "heading": heading,
                "body":    body,    # full details text — no truncation
                "url": ann_url,
                "ts":  ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid": rec_uid,
            })
            page_hits += 1

        log.info("Screener page %d: %d hits", page, page_hits)
        if hit_cutoff or not data.get("next"):
            break
        time.sleep(0.6)

    log.info("Screener TOTAL: %d relevant", len(results))
    return results


# ══════════════════════════════════════════════════════════════════════════════
# Main
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

    # Global dedup (across sources and against already-posted)
    new_items = []
    global_seen_uids: set = set(seen)
    for ann in all_ann:
        if ann["uid"] not in global_seen_uids:
            global_seen_uids.add(ann["uid"])
            new_items.append(ann)

    log.info("New (not yet posted): %d", len(new_items))
    if not new_items:
        log.info("Nothing new. Done.")
        return

    new_items.sort(key=lambda x: x["ts"])

    ws  = get_sheet()
    bot = Bot(token=TELEGRAM_BOT_TOKEN)
    seen_cos = get_seen_companies(ws)
    posted   = 0

    for ann in new_items:
        try:
            key = (ann["company"].strip().lower(), ann["symbol"].strip().lower())
            ann["first_disclosure"] = key not in seen_cos
            seen_cos = sheet_append(ws, ann, seen_cos)
            await tg_send(bot, ann)
            seen.add(ann["uid"])
            posted += 1
            flag = " ⭐FIRST" if ann["first_disclosure"] else ""
            log.info("[%s]%s %s | %s", ann["source"], flag, ann["company"], ann["heading"][:65])
        except Exception as e:
            log.error("Failed %s: %s", ann["uid"], e)
            seen.add(ann["uid"])

    save_seen(seen)
    log.info("=" * 65)
    log.info("Done. Posted %d new announcements.", posted)
    log.info("=" * 65)


if __name__ == "__main__":
    asyncio.run(main())
