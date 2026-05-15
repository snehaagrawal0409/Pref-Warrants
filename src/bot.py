"""
Preferential Warrant Announcement Bot
Sources : NSE · BSE · Screener.in
Output  : Google Sheets (first) → Telegram channel
Schedule: Every 1 hour via GitHub Actions
Lookback: Last 24 hours | Zero duplicates guaranteed
"""

import os
import re
import json
import time
import hashlib
import logging
import asyncio
import requests
import gspread
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

# ── Environment Config ─────────────────────────────────────────────────────────
TELEGRAM_BOT_TOKEN  = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHANNEL_ID = os.environ["TELEGRAM_CHANNEL_ID"]
GOOGLE_SHEET_ID     = os.environ["GOOGLE_SHEET_ID"]
GOOGLE_CREDS_JSON   = os.environ["GOOGLE_CREDS_JSON"]
SEEN_IDS_FILE       = os.environ.get("SEEN_IDS_FILE", "seen_ids.json")
LOOKBACK_HOURS      = int(os.environ.get("LOOKBACK_HOURS", "24"))

# ── Keywords ───────────────────────────────────────────────────────────────────
KEYWORDS = [
    # Core preferential
    "preferential issue",
    "preferential allotment",
    "preferential placement",
    "preferential basis",
    "preferential issue of warrants",
    "preferential issue of shares",
    # Warrants
    "issue of warrants",
    "convertible warrants",
    "allotment of warrants",
    "exercise of warrants",
    "warrants conversion",
    "conversion of warrants into equity",
    "allotment of equity shares upon exercise of warrants",
    # Capital markets
    "private placement",
    "qualified institutional placement",
    # Approvals & stages
    "in-principle approval",
    "in-principle preferential",
    # Allottee types
    "promoter allotment",
    "non-promoter allotment",
    # Financial terms
    "fund raising",
    "capital raising",
    "issue price",
    "lock-in period",
    "listing stage",
    "initial subscription amount",
    "balance consideration",
    "further allotment",
    "tranche allotment",
    # Broad catch-alls
    "warrant",
    "warrants",
    "preferential",
]

HEADERS = {
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


# ── Deduplication ──────────────────────────────────────────────────────────────

def load_seen_ids() -> set:
    if os.path.exists(SEEN_IDS_FILE):
        try:
            with open(SEEN_IDS_FILE) as f:
                return set(json.load(f))
        except Exception:
            pass
    return set()


def save_seen_ids(seen: set):
    with open(SEEN_IDS_FILE, "w") as f:
        json.dump(sorted(seen), f, indent=2)


def make_uid(*parts) -> str:
    raw = "|".join(str(p).strip().lower() for p in parts)
    return hashlib.sha256(raw.encode()).hexdigest()[:20]


# ── Google Sheets ──────────────────────────────────────────────────────────────

SHEET_HEADERS = [
    "Timestamp (IST)",   # A - when the announcement was made
    "Source",            # B - NSE / BSE / Screener
    "Company",           # C - full company name
    "Symbol",            # D - ticker / scrip code
    "Heading",           # E - announcement title
    "Summary",           # F - brief description (Screener provides this)
    "First Disclosure",  # G - YES if this company's first preferential announcement, else NO
    "URL",               # H - link to filing / PDF
    "Unique ID",         # I - dedup fingerprint
]


def get_sheet():
    creds_info = json.loads(GOOGLE_CREDS_JSON)
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(creds_info, scopes=scopes)
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        ws = sh.worksheet("Announcements")
        # If sheet exists but is missing the First Disclosure column, add it
        existing_headers = ws.row_values(1)
        if "First Disclosure" not in existing_headers and len(existing_headers) >= 6:
            # Insert "First Disclosure" before URL (was col G, now shifting)
            # Simplest: clear row 1 and rewrite all headers
            ws.update("A1:I1", [SHEET_HEADERS])
            ws.format("A1:I1", {"textFormat": {"bold": True}})
            log.info("Updated sheet headers to include 'First Disclosure'")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Announcements", rows=10000, cols=len(SHEET_HEADERS))
        ws.append_row(SHEET_HEADERS, value_input_option="RAW")
        ws.freeze(rows=1)
        ws.format("A1:I1", {"textFormat": {"bold": True}})
        # Set column widths for readability
        log.info("Created new 'Announcements' worksheet with all headers")
    return ws


def get_seen_companies(ws) -> set:
    """
    Read all company+symbol values already in the sheet.
    Used to determine if an announcement is the FIRST for that company.
    """
    try:
        all_values = ws.get_all_values()
        # Column C (index 2) = Company, Column D (index 3) = Symbol
        seen = set()
        for row in all_values[1:]:   # skip header
            if len(row) >= 4:
                key = (row[2].strip().lower(), row[3].strip().lower())
                seen.add(key)
        return seen
    except Exception:
        return set()


def sheet_append(ws, ann: dict, seen_companies: set) -> set:
    """
    Append one row. Determines First Disclosure by checking seen_companies.
    Updates seen_companies in-place and returns it.
    """
    key = (ann["company"].strip().lower(), ann["symbol"].strip().lower())
    is_first = "YES ⭐" if key not in seen_companies else "NO"
    seen_companies.add(key)

    ws.append_row([
        ann["ts"],
        ann["source"],
        ann["company"],
        ann["symbol"],
        ann["heading"],
        ann["summary"],
        is_first,
        ann["url"],
        ann["uid"],
    ], value_input_option="USER_ENTERED")

    return seen_companies


# ── Telegram ───────────────────────────────────────────────────────────────────

def escape_md(text: str) -> str:
    return re.sub(r'([_*\[\]()~`>#+\-=|{}.!\\])', r'\\\1', str(text))


def format_msg(ann: dict) -> str:
    sym_part = f" \\({escape_md(ann['symbol'])}\\)" if ann.get("symbol") else ""
    first_badge = "⭐ *FIRST DISCLOSURE*\n" if ann.get("first_disclosure") else ""
    lines = [
        f"🔔 *{escape_md(ann['source'])} \\| Preferential Warrant Alert*",
        "",
        first_badge + f"🏢 *{escape_md(ann['company'])}*{sym_part}",
        f"📌 {escape_md(ann['heading'])}",
    ]
    if ann.get("summary"):
        summary = ann["summary"][:450] + ("…" if len(ann["summary"]) > 450 else "")
        lines += ["", f"📝 _{escape_md(summary)}_"]
    lines += ["", f"🕐 {escape_md(ann['ts'])}"]
    if ann.get("url"):
        lines.append(f"🔗 [View Announcement]({ann['url']})")
    return "\n".join(lines)


async def send_telegram_async(bot: Bot, ann: dict):
    msg = format_msg(ann)
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
            log.warning("Telegram attempt %d: %s", attempt + 1, err)
            if "retry after" in err.lower():
                secs = int(re.search(r'\d+', err).group() or 10)
                await asyncio.sleep(secs + 2)
            elif attempt == 2:
                try:
                    plain = (
                        f"[{ann['source']}] {ann['company']} ({ann['symbol']})\n"
                        f"{ann['heading']}\n{ann['ts']}\n{ann.get('url','')}"
                    )
                    await bot.send_message(
                        chat_id=TELEGRAM_CHANNEL_ID,
                        text=plain,
                        disable_web_page_preview=True,
                    )
                except Exception:
                    log.error("Plain fallback also failed for: %s", ann["heading"])
            else:
                await asyncio.sleep(5)


# ── NSE Scraper ────────────────────────────────────────────────────────────────

def fetch_nse(cutoff: datetime) -> list[dict]:
    results = []
    session = requests.Session()
    try:
        session.get("https://www.nseindia.com", headers=HEADERS, timeout=15)
        time.sleep(2)
        session.get(
            "https://www.nseindia.com/companies-listing/corporate-filings-announcements",
            headers=HEADERS, timeout=15,
        )
        time.sleep(1)
    except Exception as e:
        log.warning("NSE warmup failed: %s", e)

    today     = datetime.now(IST).strftime("%d-%m-%Y")
    yesterday = (datetime.now(IST) - timedelta(hours=LOOKBACK_HOURS)).strftime("%d-%m-%Y")

    for page in range(1, 6):
        api_url = (
            f"https://www.nseindia.com/api/corporate-announcements"
            f"?index=equities&from_date={yesterday}&to_date={today}&page={page}"
        )
        try:
            resp = session.get(api_url, headers=HEADERS, timeout=20)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            log.warning("NSE page %d failed: %s", page, e)
            break

        items = data if isinstance(data, list) else data.get("data", [])
        if not items:
            break

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

            results.append({
                "source": "NSE", "symbol": symbol, "company": company,
                "heading": subject[:300], "summary": body[:500] if body else "",
                "url": pdf_url, "ts": ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid": make_uid("NSE", symbol, raw_ts, subject),
            })
        time.sleep(1)

    log.info("NSE: %d relevant", len(results))
    return results


# ── BSE Scraper ────────────────────────────────────────────────────────────────

def fetch_bse(cutoff: datetime) -> list[dict]:
    results = []
    today     = datetime.now(IST).strftime("%Y%m%d")
    yesterday = (datetime.now(IST) - timedelta(hours=LOOKBACK_HOURS)).strftime("%Y%m%d")
    bse_headers = {**HEADERS, "Referer": "https://www.bseindia.com/corporates/ann.html"}
    api_url = (
        f"https://api.bseindia.com/BseIndiaAPI/api/AnnGetData/w"
        f"?strCat=-1&strPrevDate={yesterday}&strScrip=&strSearch=P"
        f"&strToDate={today}&strType=C&subcategory=-1"
    )
    try:
        resp = requests.get(api_url, headers=bse_headers, timeout=25)
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        log.error("BSE fetch failed: %s", e)
        return results

    for item in data.get("Table", []):
        subject  = (item.get("HEADLINE") or "").strip()
        scrip_cd = str(item.get("SCRIP_CD") or "").strip()
        company  = (item.get("SLONGNAME") or scrip_cd).strip()
        if not is_relevant(subject):
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
        results.append({
            "source": "BSE", "symbol": scrip_cd, "company": company,
            "heading": subject[:300], "summary": "",
            "url": pdf_url, "ts": ann_dt.strftime("%Y-%m-%d %H:%M IST"),
            "uid": make_uid("BSE", scrip_cd, news_id, subject),
        })

    log.info("BSE: %d relevant", len(results))
    return results


# ── Screener.in Scraper ────────────────────────────────────────────────────────

def fetch_screener(cutoff: datetime) -> list[dict]:
    results = []
    screener_headers = {
        **HEADERS,
        "Referer": "https://www.screener.in/",
        "X-Requested-With": "XMLHttpRequest",
    }

    for page in range(1, 15):
        url = f"https://www.screener.in/api/announcements/?page={page}"
        try:
            resp = requests.get(url, headers=screener_headers, timeout=20)
            resp.raise_for_status()
            data = resp.json()
        except Exception as e:
            log.warning("Screener page %d failed: %s", page, e)
            break

        items = data.get("results", [])
        if not items:
            break

        hit_cutoff = False
        for item in items:
            heading = (item.get("title") or item.get("heading") or "").strip()
            summary = (
                item.get("details") or item.get("description") or item.get("summary") or ""
            ).strip()
            raw_ts = item.get("date") or item.get("datetime") or item.get("created_at") or ""

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

            if not is_relevant(heading) and not is_relevant(summary):
                continue

            company_raw  = item.get("company") or {}
            if isinstance(company_raw, dict):
                company_name = company_raw.get("name") or ""
                symbol       = company_raw.get("symbol") or company_raw.get("bse_code") or ""
            else:
                company_name = str(company_raw)
                symbol       = ""

            ann_url = item.get("url") or item.get("file_url") or item.get("attachment") or ""
            if ann_url and not ann_url.startswith("http"):
                ann_url = "https://www.screener.in" + ann_url

            results.append({
                "source": "Screener", "symbol": symbol, "company": company_name,
                "heading": heading[:300], "summary": summary[:500],
                "url": ann_url, "ts": ann_dt.strftime("%Y-%m-%d %H:%M IST"),
                "uid": make_uid("Screener", symbol, raw_ts, heading),
            })

        if hit_cutoff or not data.get("next"):
            break
        time.sleep(0.5)

    log.info("Screener: %d relevant", len(results))
    return results


# ── Main ───────────────────────────────────────────────────────────────────────

async def main():
    now_ist = datetime.now(IST)
    log.info("=" * 60)
    log.info("Bot run started: %s", now_ist.strftime("%Y-%m-%d %H:%M IST"))
    log.info("=" * 60)

    cutoff = now_ist - timedelta(hours=LOOKBACK_HOURS)
    seen   = load_seen_ids()
    log.info("Already seen: %d IDs", len(seen))

    all_ann: list[dict] = []
    all_ann.extend(fetch_nse(cutoff))
    all_ann.extend(fetch_bse(cutoff))
    all_ann.extend(fetch_screener(cutoff))
    log.info("Total fetched: %d", len(all_ann))

    new_items = [a for a in all_ann if a["uid"] not in seen]
    log.info("New (unseen): %d", len(new_items))

    if not new_items:
        log.info("Nothing new. Done.")
        return

    new_items.sort(key=lambda x: x["ts"])
    ws  = get_sheet()
    bot = Bot(token=TELEGRAM_BOT_TOKEN)
    posted = 0

    # Load companies already in the sheet (for First Disclosure detection)
    seen_companies = get_seen_companies(ws)
    log.info("Companies already in sheet: %d", len(seen_companies))

    for ann in new_items:
        try:
            # Determine first disclosure BEFORE appending (so this row counts)
            key = (ann["company"].strip().lower(), ann["symbol"].strip().lower())
            ann["first_disclosure"] = key not in seen_companies

            # 1. Google Sheet first
            seen_companies = sheet_append(ws, ann, seen_companies)

            # 2. Telegram second
            await send_telegram_async(bot, ann)

            seen.add(ann["uid"])
            posted += 1
            first_flag = " [FIRST]" if ann["first_disclosure"] else ""
            log.info("[%s]%s %s | %s", ann["source"], first_flag, ann["company"], ann["heading"][:60])
        except Exception as e:
            log.error("Failed for %s: %s", ann["uid"], e)
            seen.add(ann["uid"])

    save_seen_ids(seen)
    log.info("=" * 60)
    log.info("Done. Posted %d announcements.", posted)
    log.info("=" * 60)


if __name__ == "__main__":
    asyncio.run(main())
