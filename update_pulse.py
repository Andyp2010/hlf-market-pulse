import base64
import datetime
import io
import json
import os
import pathlib
import re
import time
import warnings
from html.parser import HTMLParser
import pandas as pd

# Silence cosmetic openpyxl stylesheet notices & Camoufox canvas notices in CI logs
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
warnings.filterwarnings("ignore", message=".*Canvas and WebGL.*")

# ---------------------------------------------------------------------------
# Official NZ Crown Data Sources (RBNZ, Stats NZ, MBIE Tenancy, Inland Revenue)
# ---------------------------------------------------------------------------
RBNZ_WARMUP_URL = "https://www.rbnz.govt.nz/statistics"

RBNZ_SOURCES = {
    "c31": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
            "https://www.rbnz.govt.nz/statistics/c31",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c31/hc31.xlsx",
        "match": "hc31",
    },
    "c32": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type",
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type2",
            "https://www.rbnz.govt.nz/statistics/c32",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c32/hc32.xlsx",
        "match": "hc32",
    },
    "c33": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose",
            "https://www.rbnz.govt.nz/statistics/c33",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c33/hc33.xlsx",
        "match": "hc33",
    },
    "c40": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use",
            "https://www.rbnz.govt.nz/statistics/c40",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c40/hc40.xlsx",
        "match": "hc40",
    },
    "c41": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi",
            "https://www.rbnz.govt.nz/statistics/c41",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c41/hc41.xlsx",
        "match": "hc41",
    },
    "b30": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-weighted-average-interest-rates",
            "https://www.rbnz.govt.nz/statistics/b30",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b30/hb30.xlsx",
        "match": "hb30",
    },
}

TENANCY_PAGE = "https://www.tenancy.govt.nz/about-tenancy-services/data-and-statistics/rental-bond-data/"
IRD_KS_NUM_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/withdrawals/number"
IRD_KS_AMT_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/withdrawals/amount"
IRD_KS_REG_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/withdrawals/kiwisaver-fund-withdrawals-by-region"
IRD_KS_TR_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/schemes/transfer"
IRD_KS_DATA_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/datasets"
IRD_SL_BAL_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/statistics-on-the-student-loan-scheme/statistics-on-student-loan-balances"

# Strict freshness gates (days since observation period end).
MAX_AGE_MONTHLY = 85
MAX_AGE_TENANCY = 120   # MBIE publishes private bond CSVs on a ~2-month processing lag
MAX_AGE_QUARTERLY = 160
MAX_AGE_ANNUAL = 420    # IRD annual scheme transfer & regional summaries publish each September for June FY


class SimpleTableExtractor(HTMLParser):
    """Zero-dependency HTML table parser for Inland Revenue regional KiwiSaver tables."""
    def __init__(self):
        super().__init__()
        self.tables = []
        self._cur_table = None
        self._cur_row = None
        self._cur_cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self._cur_table = []
        elif tag == "tr" and self._cur_table is not None:
            self._cur_row = []
        elif tag in ("td", "th") and self._cur_row is not None:
            self._cur_cell = []

    def handle_data(self, data):
        if self._cur_cell is not None:
            self._cur_cell.append(data)

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cur_cell is not None:
            self._cur_row.append("".join(self._cur_cell).strip())
            self._cur_cell = None
        elif tag == "tr" and self._cur_row is not None:
            if any(self._cur_row):
                self._cur_table.append(self._cur_row)
            self._cur_row = None
        elif tag == "table" and self._cur_table is not None:
            if self._cur_table:
                self.tables.append(self._cur_table)
            self._cur_table = None


def is_valid_xlsx(data: bytes) -> bool:
    return bool(data and len(data) > 6000 and data[:4] == b"PK\x03\x04")


def add_cb(url: str) -> str:
    sep = "&" if "?" in url else "?"
    return f"{url}{sep}_cb={int(time.time())}"


def month_end_iso(dt: pd.Timestamp) -> str:
    nxt = (dt.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
    last = nxt - datetime.timedelta(days=1)
    return last.strftime("%Y-%m-%d")


def is_fresh(as_of_iso: str, max_age_days: int) -> bool:
    try:
        obs_date = datetime.date.fromisoformat(as_of_iso[:10])
        age = (datetime.datetime.now(datetime.timezone.utc).date() - obs_date).days
        return -5 <= age <= max_age_days
    except Exception:
        return False


def fmt_nzd_millions(val_m: float) -> str:
    if abs(val_m) >= 1000.0:
        return f"${val_m / 1000.0:.2f}B"
    return f"${val_m:,.0f}m"


def plain_prior_year_pct(curr: float, prev: float) -> str:
    diff = curr - prev
    if abs(diff) < 0.25:
        return f"broadly steady alongside {prev:.1f}% a year earlier"
    direction = "up from" if diff > 0 else "down from"
    return f"{direction} {prev:.1f}% a year earlier"


def plain_yoy_growth(curr: float, prev: float) -> str:
    if not prev:
        return ""
    pct = ((curr - prev) / prev) * 100.0
    if abs(pct) < 0.25:
        return "virtually unchanged from a year earlier"
    direction = "up" if pct > 0 else "down"
    return f"{direction} {abs(pct):.1f}% on a year earlier"


# ---------------------------------------------------------------------------
# 3-Tier Anti-Cloudflare & HTTP Download Pipeline
# ---------------------------------------------------------------------------
def get_http_session():
    from curl_cffi import requests as cf_requests
    return cf_requests.Session(impersonate="chrome")


def open_camoufox():
    from camoufox.sync_api import Camoufox
    try:
        return Camoufox(headless=False, humanize=True, i_know_what_im_doing=True)
    except TypeError:
        return Camoufox(headless=False, humanize=True)


def fetch_url_bytes(session, url: str, referer: str = "", timeout: int = 45, use_cb: bool = True) -> bytes:
    headers = {
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-NZ,en;q=0.9",
        "Cache-Control": "no-cache",
    }
    if referer:
        headers["Referer"] = referer
    try:
        target = add_cb(url) if use_cb else url
        resp = session.get(target, headers=headers, timeout=timeout)
        return resp.content if resp.status_code == 200 else b""
    except Exception:
        return b""


def fetch_html_with_browser_fallback(session, urls: list[str], label: str) -> dict[str, str]:
    results = {}
    blocked = []

    for u in urls:
        raw = fetch_url_bytes(session, u, timeout=25, use_cb=False)
        html = raw.decode("utf-8", errors="ignore") if raw else ""
        low = html.lower()
        if len(html) > 2000 and "just a moment" not in low and "_incapsula_resource" not in low and "access denied" not in low:
            results[u] = html
        else:
            blocked.append(u)

    if not blocked:
        return results

    print(f"[{label}] {len(blocked)} URL(s) challenged on Tier 1; trying Camoufox browser...")
    try:
        with open_camoufox() as browser:
            page = browser.new_page()
            for u in blocked:
                try:
                    page.goto(u, wait_until="domcontentloaded", timeout=50000)
                    deadline = time.time() + 25
                    while time.time() < deadline:
                        title = (page.title() or "").lower()
                        body_html = page.content() or ""
                        if len(body_html) > 3000 and "just a moment" not in title and "_incapsula_resource" not in body_html.lower():
                            results[u] = body_html
                            break
                        page.wait_for_timeout(2000)
                except Exception as exc:
                    print(f"[{label}] Camoufox note on {u}: {exc}")
    except Exception as exc:
        print(f"[{label}] Camoufox launch note: {exc}")

    return results


def warmup_rbnz_session(page, tier_label: str) -> None:
    """Clear Cloudflare Turnstile & cookie banners once before clicking any .xlsx links."""
    try:
        page.goto(add_cb(RBNZ_WARMUP_URL), wait_until="domcontentloaded", timeout=50000)
        deadline = time.time() + 30
        while time.time() < deadline:
            title = (page.title() or "").lower()
            if title and "just a moment" not in title and "attention required" not in title:
                break
            page.wait_for_timeout(2000)
        page.wait_for_timeout(1500)
        for sel in ('button:has-text("Accept")', 'button:has-text("I agree")', 'button:has-text("Close")'):
            btn = page.locator(sel)
            if btn.count() > 0 and btn.first.is_visible():
                try:
                    btn.first.click(timeout=2000)
                except Exception:
                    pass
    except Exception as exc:
        print(f"[{tier_label}] Warmup note: {exc}")


def fetch_xlsx_in_browser_page(page, cfg: dict, dest: pathlib.Path, tier_label: str, allow_js_fetch: bool = False) -> bool:
    for p_url in cfg["pages"]:
        try:
            page.goto(add_cb(p_url), wait_until="domcontentloaded", timeout=60000)
            deadline = time.time() + 35
            while time.time() < deadline:
                if page.locator('a[href*=".xlsx"]').count() > 0 and "just a moment" not in (page.title() or "").lower():
                    break
                page.wait_for_timeout(2000)
            if page.locator('a[href*=".xlsx"]').count() > 0:
                break
        except Exception:
            continue

    page.wait_for_timeout(800)

    visible_locs, hidden_locs = [], []
    target = cfg["xlsx"]
    for link in page.locator('a[href*=".xlsx"]').all():
        href = link.get_attribute("href") or ""
        if cfg["match"] in href.lower():
            full = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
            if not visible_locs and not hidden_locs:
                target = full
            try:
                if link.is_visible():
                    visible_locs.append((link, full))
                else:
                    hidden_locs.append((link, full))
            except Exception:
                hidden_locs.append((link, full))

    ordered_candidates = visible_locs + hidden_locs
    if visible_locs:
        target = visible_locs[0][1]

    if allow_js_fetch:
        try:
            b64 = page.evaluate(
                """async (url) => {
                    const r = await fetch(url, { cache: 'no-store', credentials: 'include' });
                    if (!r.ok) return null;
                    const bytes = new Uint8Array(await r.arrayBuffer());
                    let bin = '';
                    for (let i = 0; i < bytes.length; i += 0x8000) {
                        bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
                    }
                    return btoa(bin);
                }""",
                add_cb(target),
            )
            if b64:
                raw = base64.b64decode(b64)
                if is_valid_xlsx(raw):
                    dest.write_bytes(raw)
                    print(f"[{tier_label}] Downloaded {cfg['match'].upper()} ({len(raw):,} bytes)")
                    return True
        except Exception as exc:
            print(f"[{tier_label}] In-browser fetch note ({cfg['match']}): {exc}")

    for attempt in range(2):
        for loc, _ in ordered_candidates[:3]:
            try:
                try:
                    loc.scroll_into_view_if_needed(timeout=3000)
                except Exception:
                    pass
                with page.expect_download(timeout=15000) as dl_info:
                    loc.click(force=True)
                dl_info.value.save_as(str(dest))
                if dest.exists() and is_valid_xlsx(dest.read_bytes()):
                    print(f"[{tier_label}] Click-downloaded {cfg['match'].upper()}")
                    return True
            except Exception:
                continue
        if attempt == 0:
            try:
                page.goto(add_cb(cfg["pages"][0]), wait_until="domcontentloaded", timeout=45000)
                page.wait_for_timeout(2000)
                ordered_candidates = [
                    (lnk, "") for lnk in page.locator('a[href*=".xlsx"]').all()
                    if cfg["match"] in (lnk.get_attribute("href") or "").lower()
                ]
            except Exception:
                break

    print(f"[{tier_label}] Could not download {cfg['match'].upper()} after retries.")
    return False


def download_rbnz_workbooks(dest_dir: pathlib.Path, session) -> dict:
    saved = {}
    for key, cfg in RBNZ_SOURCES.items():
        manual = dest_dir / f"{cfg['match']}.xlsx"
        if manual.exists() and is_valid_xlsx(manual.read_bytes()):
            saved[key] = manual

    # Tier 1: curl_cffi Chrome TLS impersonation
    for key in [k for k in RBNZ_SOURCES if k not in saved]:
        cfg = RBNZ_SOURCES[key]
        target = cfg["xlsx"]
        ref_used = cfg["pages"][0]
        for p_url in cfg["pages"]:
            page_bytes = fetch_url_bytes(session, p_url, timeout=25)
            if page_bytes:
                ref_used = p_url
                html = page_bytes.decode("utf-8", errors="ignore")
                for href in re.findall(r'href=["\']([^"\']*\.xlsx[^"\']*)["\']', html, flags=re.I):
                    if cfg["match"] in href.lower():
                        target = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
                        break
                break
        raw = fetch_url_bytes(session, target, referer=ref_used, timeout=45)
        if is_valid_xlsx(raw):
            out = dest_dir / f"tmp_{key}.xlsx"
            out.write_bytes(raw)
            saved[key] = out
            print(f"[RBNZ Tier 1 - curl_cffi] Downloaded {key.upper()} ({len(raw):,} bytes)")

    missing = [k for k in RBNZ_SOURCES if k not in saved]
    if not missing:
        return saved

    # Tier 2: Camoufox anti-detect Firefox under Xvfb
    try:
        with open_camoufox() as browser:
            page = browser.new_page()
            warmup_rbnz_session(page, "RBNZ Tier 2 - Camoufox")
            for key in list(missing):
                out = dest_dir / f"tmp_{key}.xlsx"
                if fetch_xlsx_in_browser_page(page, RBNZ_SOURCES[key], out, "RBNZ Tier 2 - Camoufox", allow_js_fetch=False):
                    saved[key] = out
    except Exception as exc:
        print(f"[RBNZ Tier 2 - Camoufox] Note: {exc}")

    missing = [k for k in RBNZ_SOURCES if k not in saved]
    if not missing:
        return saved

    # Tier 3: Headed Chromium via Playwright under Xvfb
    try:
        from playwright.sync_api import sync_playwright
        prof = pathlib.Path("/tmp/hlf_pulse_chrome")
        prof.mkdir(parents=True, exist_ok=True)
        with sync_playwright() as p:
            kw = dict(
                user_data_dir=str(prof),
                headless=False,
                accept_downloads=True,
                viewport={"width": 1366, "height": 900},
                args=["--disable-blink-features=AutomationControlled", "--no-sandbox"],
            )
            try:
                ctx = p.chromium.launch_persistent_context(channel="chrome", **kw)
            except Exception:
                ctx = p.chromium.launch_persistent_context(**kw)
            page = ctx.pages[0] if ctx.pages else ctx.new_page()
            warmup_rbnz_session(page, "RBNZ Tier 3 - Chrome")
            for key in list(missing):
                out = dest_dir / f"tmp_{key}.xlsx"
                if fetch_xlsx_in_browser_page(page, RBNZ_SOURCES[key], out, "RBNZ Tier 3 - Chrome", allow_js_fetch=True):
                    saved[key] = out
            ctx.close()
    except Exception as exc:
        print(f"[RBNZ Tier 3 - Chrome] Note: {exc}")

    return saved


# ---------------------------------------------------------------------------
# Multi-Sheet RBNZ Workbook Scanner (Case-Insensitive Series IDs & Hyphen-Safe)
# ---------------------------------------------------------------------------
def norm_label(txt: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[-\u2013\u2014]+", " ", txt.lower())).strip()


def load_rbnz_sheets(filepath: pathlib.Path) -> list:
    xls = pd.ExcelFile(filepath)
    parsed_tabs = []
    skip_words = ("description", "definition", "note", "metadata")
    for tab in xls.sheet_names:
        if any(w in tab.lower() for w in skip_words):
            continue
        raw = pd.read_excel(xls, sheet_name=tab, header=None)
        if raw.shape[0] < 10 or raw.shape[1] < 2:
            continue

        date_col = pd.to_datetime(raw.iloc[:, 0], errors="coerce", format="mixed")
        valid_dates = date_col.dropna()
        if len(valid_dates) < 6:
            continue
        first_data_row = int(valid_dates.index[0])
        hdr_rows = max(1, first_data_row)

        sids, labels = [], []
        for col_idx in range(raw.shape[1]):
            col_parts = [
                str(raw.iloc[r, col_idx]).strip()
                for r in range(hdr_rows)
                if pd.notna(raw.iloc[r, col_idx]) and str(raw.iloc[r, col_idx]).strip().lower() != "nan"
            ]
            sids.append((col_parts[-1] if col_parts else "").upper())
            joined = f"{tab} | " + " | ".join(col_parts)
            labels.append(norm_label(joined))

        df = raw.iloc[first_data_row:].copy()
        df["date"] = pd.to_datetime(df.iloc[:, 0], errors="coerce", format="mixed")
        df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
        parsed_tabs.append((tab, df, sids, labels))
    return parsed_tabs


def find_rbnz_series(tabs: list, exact_sids=(), include_all=(), exclude_any=(), val_bounds=None):
    inc_norm = [norm_label(k) for k in include_all]
    exc_norm = [norm_label(x) for x in exclude_any]

    for _, df, sids, _ in tabs:
        for sid in exact_sids:
            sid_u = sid.upper()
            if sid_u in sids:
                idx = sids.index(sid_u)
                ser = pd.to_numeric(df.iloc[:, idx], errors="coerce")
                valid = ser.dropna()
                if len(valid) and (val_bounds is None or (val_bounds[0] <= float(valid.iloc[-1]) <= val_bounds[1])):
                    return df["date"], ser, sid_u
    for _, df, sids, labels in tabs:
        for idx, lbl in enumerate(labels):
            if idx == 0:
                continue
            if all(k in lbl for k in inc_norm) and not any(x in lbl for x in exc_norm):
                ser = pd.to_numeric(df.iloc[:, idx], errors="coerce")
                valid = ser.dropna()
                if len(valid) and (val_bounds is None or (val_bounds[0] <= float(valid.iloc[-1]) <= val_bounds[1])):
                    return df["date"], ser, (sids[idx] or lbl[:45])
    raise KeyError(f"RBNZ series not found: sids={exact_sids}, include={include_all}")


# ---------------------------------------------------------------------------
# Conversational, Story-Driven, FMA-Safe Parsers
# ---------------------------------------------------------------------------
def build_rbnz_c31_tickers(filepath: pathlib.Path) -> tuple[dict, dict]:
    tabs = load_rbnz_sheets(filepath)
    dates, fhb_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AC1",), ("total lending", "first home"))
    _, oo_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AC2",), ("total lending", "other owner"))
    _, inv_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AN", "LVRN.MMB1.AC3"), ("total lending", "investor"))
    try:
        _, bus_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AP", "LVRN.MMB1.AC4"), ("total lending", "business"))
    except KeyError:
        bus_val = pd.Series([0.0] * len(dates))

    tot_val = fhb_val.fillna(0) + oo_val.fillna(0) + inv_val.fillna(0) + bus_val.fillna(0)
    _, fhb_hi_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.BC1",), ("80", "lending", "first home"), ("less", "equal"))
    _, fhb_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AC1",), ("borrowers", "first home"), ("80",))
    _, fhb_hi_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.BC1",), ("80", "borrowers", "first home"), ("less", "equal"))

    # Additional C31 series for FHB deep-dives & homeowner/investor equity positions
    _, oo_hi_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.BC2",), ("80", "lending", "other owner"), ("less", "equal"))
    _, oo_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AC2",), ("borrowers", "other owner"), ("80",))
    _, oo_hi_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.BC2",), ("80", "borrowers", "other owner"), ("less", "equal"))
    _, all_hi_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.BA",), ("80", "lending", "all borrower"), ("less", "equal"))
    _, all_hi_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.BA",), ("80", "borrowers", "all borrower"), ("less", "equal"))
    _, inv_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AN",), ("total", "borrowers", "investor"), ("70", "60", "80"))
    _, inv_gt70_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.DN",), ("70", "lending", "investor"), ("less", "equal"))
    _, inv_gt70_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.DN",), ("70", "borrowers", "investor"), ("less", "equal"))
    _, inv_gt60_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.FN",), ("60", "borrowers", "investor"), ("less", "equal"))
    _, inv_le60_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.TN",), ("60", "borrowers", "investor", "less"))

    i, iy = len(dates) - 1, max(0, len(dates) - 13)
    dt = dates.iloc[i]
    m_full = dt.strftime("%B %Y")
    as_of = month_end_iso(dt)

    low_cnt_pct = (float(fhb_hi_cnt.iloc[i]) / float(fhb_cnt.iloc[i])) * 100.0
    low_cnt_yoy = (float(fhb_hi_cnt.iloc[iy]) / float(fhb_cnt.iloc[iy])) * 100.0
    hi_dollars_m = float(fhb_hi_val.iloc[i])
    fhb_loans_int = int(round(float(fhb_cnt.iloc[i])))
    fhb_hi_loans_int = int(round(float(fhb_hi_cnt.iloc[i])))

    fhb_sh = (float(fhb_val.iloc[i]) / float(tot_val.iloc[i])) * 100.0
    fhb_sh_yoy = (float(fhb_val.iloc[iy]) / float(tot_val.iloc[iy])) * 100.0
    fhb_avg_k = int(round((float(fhb_val.iloc[i]) * 1000.0) / float(fhb_cnt.iloc[i])))

    inv_sh = (float(inv_val.iloc[i]) / float(tot_val.iloc[i])) * 100.0
    inv_sh_yoy = (float(inv_val.iloc[iy]) / float(tot_val.iloc[iy])) * 100.0
    inv_b = float(inv_val.iloc[i]) / 1000.0

    if low_cnt_pct >= 54.0:
        low_dep_html = (
            f"Buying a first home without a full 20% deposit is more common than many buyers expect. "
            f"In {m_full}, more than half of all first-home mortgages approved across New Zealand "
            f"(<strong>{low_cnt_pct:.1f}%</strong>—<strong>{fhb_hi_loans_int:,} of {fhb_loans_int:,} loans</strong>, or "
            f"<strong>${hi_dollars_m:,.0f}m</strong> in lending) went to buyers with <strong>under a 20% deposit</strong>, "
            f"{plain_prior_year_pct(low_cnt_pct, low_cnt_yoy)}."
        )
    elif low_cnt_pct >= 50.0:
        low_dep_html = (
            f"Buying a first home without a full 20% deposit is more common than many buyers expect. "
            f"In {m_full}, just over half of all first-home mortgages approved across New Zealand "
            f"(<strong>{low_cnt_pct:.1f}%</strong>—<strong>{fhb_hi_loans_int:,} of {fhb_loans_int:,} loans</strong>, or "
            f"<strong>${hi_dollars_m:,.0f}m</strong> in lending) went to buyers with <strong>under a 20% deposit</strong>, "
            f"{plain_prior_year_pct(low_cnt_pct, low_cnt_yoy)}."
        )
    elif low_cnt_pct >= 30.0:
        low_dep_html = (
            f"Getting onto the property ladder without a full 20% deposit remains a well-trodden path in New Zealand. "
            f"In {m_full}, <strong>{low_cnt_pct:.1f}%</strong> of approved first-home mortgages "
            f"(<strong>{fhb_hi_loans_int:,} of {fhb_loans_int:,} loans</strong>, totalling <strong>${hi_dollars_m:,.0f}m</strong>) "
            f"went to buyers with <strong>under a 20% deposit</strong>, {plain_prior_year_pct(low_cnt_pct, low_cnt_yoy)}."
        )
    else:
        low_dep_html = (
            f"Even when low-equity lending is tighter, banks still approve low-deposit home loans every month—in {m_full}, "
            f"<strong>{low_cnt_pct:.1f}%</strong> of approved first-home mortgages (<strong>${hi_dollars_m:,.0f}m</strong>) "
            f"went to buyers with <strong>under a 20% deposit</strong> ({plain_prior_year_pct(low_cnt_pct, low_cnt_yoy)})."
        )

    fhb_share_html = (
        f"First-home buyers continue to make up an active slice of the New Zealand market—in {m_full}, banks approved "
        f"<strong>{fhb_loans_int:,}</strong> first-home mortgages (<strong>{fhb_sh:.1f}%</strong> of all new mortgage lending, "
        f"{plain_prior_year_pct(fhb_sh, fhb_sh_yoy)}), with an average loan size of <strong>${fhb_avg_k:,}k</strong>."
    )

    gap = inv_sh - fhb_sh
    if gap > 0.75:
        inv_html = (
            f"Property investors are currently writing a larger share of new mortgage lending than first-home buyers—accounting for "
            f"<strong>{inv_sh:.1f}%</strong> of new commitments in {m_full} (<strong>${inv_b:.2f}B</strong>, "
            f"{plain_prior_year_pct(inv_sh, inv_sh_yoy)}), ahead of first-home buyers on <strong>{fhb_sh:.1f}%</strong>."
        )
    elif gap < -0.75:
        inv_html = (
            f"First-home buyers are currently borrowing more overall than property investors—in {m_full}, investors made up "
            f"<strong>{inv_sh:.1f}%</strong> of new mortgage lending (<strong>${inv_b:.2f}B</strong>, "
            f"{plain_prior_year_pct(inv_sh, inv_sh_yoy)}), compared with <strong>{fhb_sh:.1f}%</strong> for first-home buyers."
        )
    else:
        inv_html = (
            f"Property investors and first-home buyers are currently running neck and neck in the lending data—investors accounted for "
            f"<strong>{inv_sh:.1f}%</strong> of new mortgage lending in {m_full} (<strong>${inv_b:.2f}B</strong>), "
            f"right alongside first-home buyers at <strong>{fhb_sh:.1f}%</strong>."
        )

    # New FHB & Equity Position Tickers from C31
    fhb_lo_val_cur = float(fhb_val.iloc[i] - fhb_hi_val.iloc[i])
    fhb_lo_cnt_cur = float(fhb_cnt.iloc[i] - fhb_hi_cnt.iloc[i])
    avg_hi_loan = (float(fhb_hi_val.iloc[i]) * 1_000_000.0) / max(1.0, float(fhb_hi_cnt.iloc[i]))
    avg_lo_loan = (fhb_lo_val_cur * 1_000_000.0) / max(1.0, fhb_lo_cnt_cur)
    loan_gap = avg_hi_loan - avg_lo_loan

    if loan_gap > 10_000:
        fhb_gap_html = (
            f"First-home buyers with under a 20% deposit actually take out larger mortgages on average "
            f"(<strong>${avg_hi_loan:,.0f}</strong>) than buyers with a 20%+ deposit (<strong>${avg_lo_loan:,.0f}</strong>)—a "
            f"<strong>${abs(loan_gap):,.0f}</strong> difference in {m_full}, as higher-earning couples often use a "
            f"10%–15% deposit to buy sooner rather than renting for extra years to save 20%."
        )
    else:
        fhb_gap_html = (
            f"In {m_full}, New Zealand first-home buyers with under a 20% deposit borrowed an average of "
            f"<strong>${avg_hi_loan:,.0f}</strong>, compared with <strong>${avg_lo_loan:,.0f}</strong> for buyers with a 20%+ deposit."
        )

    fhb_hi_val_share = (float(fhb_hi_val.iloc[i]) / max(1.0, float(all_hi_val.iloc[i]))) * 100.0
    fhb_hi_cnt_share = (float(fhb_hi_cnt.iloc[i]) / max(1.0, float(all_hi_cnt.iloc[i]))) * 100.0
    fhb_priority_html = (
        f"When New Zealand banks approve low-deposit mortgages (&gt;80% LVR), most of their allowance goes to first-home buyers. "
        f"In {m_full}, <strong>{fhb_hi_val_share:.1f}% of all low-deposit mortgage dollars "
        f"({fmt_nzd_millions(float(fhb_hi_val.iloc[i]))} of {fmt_nzd_millions(float(all_hi_val.iloc[i]))})</strong> and "
        f"<strong>{fhb_hi_cnt_share:.1f}% of low-deposit loans ({fhb_hi_loans_int:,} of {int(round(float(all_hi_cnt.iloc[i]))):,})</strong> "
        f"went directly to first-home buyers."
    )

    std_dep_pct = (fhb_lo_cnt_cur / max(1.0, float(fhb_cnt.iloc[i]))) * 100.0
    fhb_20club_html = (
        f"Crossing the 20% deposit mark unlocks sharp advertised bank special rates and avoids low-equity margins. "
        f"In {m_full}, <strong>{int(round(fhb_lo_cnt_cur)):,} first-home buyers ({std_dep_pct:.1f}% of approvals, or "
        f"{fmt_nzd_millions(fhb_lo_val_cur)})</strong> bought with a <strong>20%+ deposit</strong>, borrowing an average of "
        f"<strong>${avg_lo_loan:,.0f}</strong>."
    )

    fhb_v = float(fhb_val.iloc[i])
    inv_v = float(inv_val.iloc[i])
    diff_fi = fhb_v - inv_v
    if diff_fi >= 0:
        fhb_vs_inv_html = (
            f"First-home buyers are out-borrowing property investors across New Zealand. In {m_full}, banks approved "
            f"<strong>{fmt_nzd_millions(fhb_v)} ({fhb_sh:.1f}% of lending)</strong> for first-home buyers compared with "
            f"<strong>{fmt_nzd_millions(inv_v)} ({inv_sh:.1f}%)</strong> for property investors—putting first-home buyers "
            f"<strong>{fmt_nzd_millions(abs(diff_fi))} ahead</strong>."
        )
    else:
        fhb_vs_inv_html = (
            f"In {m_full}, New Zealand first-home buyers accounted for <strong>{fmt_nzd_millions(fhb_v)} ({fhb_sh:.1f}% of lending)</strong> "
            f"alongside <strong>{fmt_nzd_millions(inv_v)} ({inv_sh:.1f}%)</strong> for property investors."
        )

    oo_tot_c = float(oo_cnt.iloc[i])
    oo_hi_c = float(oo_hi_cnt.iloc[i])
    oo_lo_c = oo_tot_c - oo_hi_c
    oo_lo_v = float(oo_val.iloc[i] - oo_hi_val.iloc[i])
    oo_20eq_pct = (oo_lo_c / max(1.0, oo_tot_c)) * 100.0
    oo_avg_20eq = (oo_lo_v * 1_000_000.0) / max(1.0, oo_lo_c)
    owner_eq_html = (
        f"Most existing Kiwi homeowners hold more usable equity than they realise. In {m_full}, "
        f"<strong>{oo_20eq_pct:.1f}%</strong> of existing owner-occupier mortgages approved across New Zealand "
        f"(<strong>{int(round(oo_lo_c)):,} of {int(round(oo_tot_c)):,} loans</strong>, or <strong>{fmt_nzd_millions(oo_lo_v)}</strong>) "
        f"had <strong>at least 20% equity</strong>, carrying an average loan size of <strong>${oo_avg_20eq:,.0f}</strong>."
    )

    inv_c_tot = float(inv_cnt.iloc[i])
    inv_40plus_pct = (float(inv_le60_cnt.iloc[i]) / max(1.0, inv_c_tot)) * 100.0
    inv_30to40_pct = (float(inv_gt60_cnt.iloc[i] - inv_gt70_cnt.iloc[i]) / max(1.0, inv_c_tot)) * 100.0
    inv_u30_cnt = int(round(float(inv_gt70_cnt.iloc[i])))
    inv_u30_pct = (inv_u30_cnt / max(1.0, inv_c_tot)) * 100.0
    inv_u30_val = float(inv_gt70_val.iloc[i])
    inv_eq_html = (
        f"You don't need 40% equity locked up to buy an investment property. In {m_full}, <strong>{inv_40plus_pct:.1f}%</strong> "
        f"of approved NZ property investors had 40%+ equity—while <strong>{inv_30to40_pct:.1f}%</strong> bought in the "
        f"<strong>30%–40% equity band</strong> and <strong>{inv_u30_pct:.1f}% ({inv_u30_cnt:,} loans, or {fmt_nzd_millions(inv_u30_val)})</strong> "
        f"were approved with <strong>under 30% equity</strong>."
    )

    oo_low_eq_pct = (oo_hi_c / max(1.0, oo_tot_c)) * 100.0
    oo_low_eq_prev = (float(oo_hi_cnt.iloc[iy]) / max(1.0, float(oo_cnt.iloc[iy]))) * 100.0
    mover_low_eq_html = (
        f"Low-deposit lending isn't restricted to first-home buyers. In {m_full}, New Zealand banks approved "
        f"<strong>{int(round(oo_hi_c)):,} mortgages ({fmt_nzd_millions(float(oo_hi_val.iloc[i]))})</strong> for existing owner-occupiers "
        f"moving or borrowing with <strong>under 20% equity (&gt;80% LVR)</strong>—accounting for <strong>{oo_low_eq_pct:.1f}%</strong> "
        f"of existing homeowner approvals, {plain_prior_year_pct(oo_low_eq_pct, oo_low_eq_prev)}."
    )

    src_c31 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type" '
        'target="_blank" rel="noopener">RBNZ Table C31</a>, &copy; Reserve Bank of New Zealand. Percentages &amp; averages derived by HLF. '
        'Historical market statistics for general information only, not personalised financial advice.'
    )

    def make_c31_item(html_str: str) -> dict:
        return {
            "chip": "RBNZ C31",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": html_str,
            "source_html": src_c31,
        }

    out = {
        "fhb-low-deposit": make_c31_item(low_dep_html),
        "fhb-share": make_c31_item(fhb_share_html),
        "investors": make_c31_item(inv_html),
        "fhb-loan-size-gap": make_c31_item(fhb_gap_html),
        "fhb-bank-lvr-priority": make_c31_item(fhb_priority_html),
        "fhb-20-deposit-club": make_c31_item(fhb_20club_html),
        "fhb-vs-investors": make_c31_item(fhb_vs_inv_html),
        "owner-equity-buffer": make_c31_item(owner_eq_html),
        "investor-equity-tiers": make_c31_item(inv_eq_html),
        "mover-low-equity": make_c31_item(mover_low_eq_html),
    }

    if i >= 23:
        r12_cnt = int(round(fhb_cnt.iloc[i - 11 : i + 1].sum()))
        r12_prev_cnt = int(round(fhb_cnt.iloc[i - 23 : i - 11].sum()))
        r12_val = float(fhb_val.iloc[i - 11 : i + 1].sum())
        r12_hi_cnt = int(round(fhb_hi_cnt.iloc[i - 11 : i + 1].sum()))
        r12_hi_pct = (r12_hi_cnt / max(1, r12_cnt)) * 100.0
        dir_r12 = "up from" if r12_cnt > r12_prev_cnt else "compared with"
        out["fhb-annual-milestone"] = make_c31_item(
            f"Over the 12 months to {m_full}, New Zealand banks approved <strong>{r12_cnt:,} first-home buyer mortgages</strong> "
            f"totalling <strong>{fmt_nzd_millions(r12_val)} in lending</strong> ({dir_r12} {r12_prev_cnt:,} loans a year earlier)—including "
            f"<strong>{r12_hi_cnt:,} buyers ({r12_hi_pct:.1f}%)</strong> who bought with <strong>under a 20% deposit</strong>."
        )

    by_ym = {}
    for idx_r in range(len(dates)):
        ym_key = dates.iloc[idx_r].strftime("%Y-%m")
        by_ym[ym_key] = {
            "fhb_cnt": float(fhb_cnt.iloc[idx_r] or 0.0),
            "inv_val_m": float(inv_val.iloc[idx_r] or 0.0),
            "oo_all_val_m": float((fhb_val.iloc[idx_r] or 0.0) + (oo_val.iloc[idx_r] or 0.0)),
        }
    return out, by_ym


def build_rbnz_c33_tickers(filepath: pathlib.Path) -> dict:
    tabs = load_rbnz_sheets(filepath)
    dates, top_v, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AZB",), ("lending", "top"))
    _, pur_v, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AZC",), ("lending", "purchase"))
    _, ref_v, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AZD",), ("lending", "provider"))
    _, oth_v, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AZE",), ("lending", "other"))
    _, top_c, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AZB",), ("borrowers", "top"))
    _, pur_c, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AZC",), ("borrowers", "purchase"))
    _, ref_c, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AZD",), ("borrowers", "provider"))
    _, oth_c, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AZE",), ("borrowers", "other"))

    j = len(dates) - 1
    dt = dates.iloc[j]
    m_full = dt.strftime("%B %Y")
    as_of = month_end_iso(dt)

    tot_v = float(top_v.iloc[j] + pur_v.iloc[j] + ref_v.iloc[j] + oth_v.iloc[j])
    tot_c = float(top_c.iloc[j] + pur_c.iloc[j] + ref_c.iloc[j] + oth_c.iloc[j])
    ref_sh = (float(ref_v.iloc[j]) / tot_v) * 100.0
    ref_avg_k = int(round((float(ref_v.iloc[j]) * 1000.0) / float(ref_c.iloc[j])))
    top_sh_c = (float(top_c.iloc[j]) / tot_c) * 100.0
    top_avg_k = int(round((float(top_v.iloc[j]) * 1000.0) / float(top_c.iloc[j])))
    top_val_m = float(top_v.iloc[j])
    top_cnt_int = int(round(float(top_c.iloc[j])))

    src_c33 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose" '
        'target="_blank" rel="noopener">RBNZ Table C33</a>, &copy; Reserve Bank of New Zealand. Derived by HLF. '
        'Historical market statistics for general information only, not personalised financial advice.'
    )
    return {
        "refinance": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"Moving banks is a regular part of the New Zealand mortgage market rather than a rare step—in {m_full} alone, "
                f"<strong>{int(round(float(ref_c.iloc[j]))):,}</strong> borrowers switched mortgage providers, moving "
                f"<strong>${float(ref_v.iloc[j])/1000.0:.2f}B</strong> in loans (<strong>{ref_sh:.1f}%</strong> of all new lending, "
                f"at an average loan size of <strong>${ref_avg_k:,}k</strong>)."
            ),
            "source_html": src_c33,
        },
        "top-ups": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"Many Kiwi homeowners tap into existing equity for renovations or upgrades rather than moving house—in {m_full}, "
                f"<strong>{top_cnt_int:,}</strong> borrowers topped up an existing mortgage, making up "
                f"<strong>{top_sh_c:.1f}%</strong> of all mortgage approvals by count (average top-up: <strong>${top_avg_k:,}k</strong>)."
            ),
            "source_html": src_c33,
        },
        "equity-topup-demand": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"Tapping into home equity without selling is one of the most common moves in New Zealand banking—in {m_full}, "
                f"homeowners unlocked <strong>{fmt_nzd_millions(top_val_m)} in mortgage top-ups</strong> across "
                f"<strong>{top_cnt_int:,} existing properties</strong> (averaging <strong>${top_avg_k:,}k</strong> per top-up) "
                f"for renovations, debt consolidation, and investment deposits."
            ),
            "source_html": src_c33,
        },
    }


def build_rbnz_c32_ticker(filepath: pathlib.Path, c31_by_ym: dict | None = None) -> dict:
    tabs = load_rbnz_sheets(filepath)
    dates, inv_io, _ = find_rbnz_series(
        tabs, (), ("investor", "interest only"), ("80", "less", "higher", "existing", "%", "share", "proportion")
    )
    _, oo_io, _ = find_rbnz_series(
        tabs, (), ("owner", "interest only"), ("80", "less", "higher", "existing", "%", "share", "proportion")
    )
    _, inv_tot, _ = find_rbnz_series(
        tabs,
        ("LVRN.MMB1.AN",),
        ("new lending by borrower type", "investors"),
        ("interest only", "principal", "80", "existing", "%"),
    )
    _, oo_tot, _ = find_rbnz_series(
        tabs,
        ("LVRN.MMB1.AC",),
        ("new lending by borrower type", "owner occupiers"),
        ("interest only", "principal", "80", "existing", "%"),
    )

    dt = dates.iloc[-1]
    inv_io_pct = (float(inv_io.iloc[-1]) / float(inv_tot.iloc[-1])) * 100.0
    oo_io_pct = (float(oo_io.iloc[-1]) / float(oo_tot.iloc[-1])) * 100.0

    m_full = dt.strftime("%B %Y")
    as_of = month_end_iso(dt)
    src_c32 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type" '
        'target="_blank" rel="noopener">RBNZ Table C32</a>, &copy; Reserve Bank of New Zealand. Derived by HLF. '
        'Historical market statistics for general information only, not personalised financial advice.'
    )
    io_item = {
        "chip": "RBNZ C32",
        "as_of": as_of,
        "max_age_days": MAX_AGE_MONTHLY,
        "html": (
            f"Loan structures look quite different between rental properties and family homes in the Reserve Bank figures—in {m_full}, "
            f"<strong>{inv_io_pct:.1f}%</strong> of new investor lending was set up on <strong>interest-only</strong> terms "
            f"(including revolving credit), compared with <strong>{oo_io_pct:.1f}%</strong> for owner-occupiers."
        ),
        "source_html": src_c32,
    }
    return {
        "interest-only": io_item,
        "investor-vs-owner-io": dict(io_item),
    }


def compute_c41_weighted_avg_income(tabs41: list, borrower_kw: str, region_mode: str = "nz") -> tuple[pd.Timestamp | None, float | None]:
    """
    Compute exact borrower-weighted average gross income (NZD) across TDTI bands in RBNZ C41.
    Supports both FHB bands (up to tdti > 6) and Investor bands (including 6 < tdti <= 7 and tdti > 7),
    as well as region_mode='nz' (nationwide) or 'akl' (Auckland).
    """
    band_keys = (
        "tdti <= 3",
        "3 < tdti <= 4",
        "4 < tdti <= 5",
        "5 < tdti <= 6",
        "6 < tdti <= 7",
        "tdti > 6",
        "tdti > 7",
    )

    for _, df41, _, labels41 in tabs41:
        inc_cols, avg_cols = {}, {}
        for idx, lbl in enumerate(labels41):
            if borrower_kw not in lbl or any(x in lbl for x in ("80", "lvr")):
                continue
            has_akl = "auckland" in lbl
            has_non = ("outside auckland" in lbl) or ("non auckland" in lbl) or ("excl auckland" in lbl)
            if region_mode == "nz" and (has_akl or has_non):
                continue
            if region_mode == "akl" and (not has_akl or has_non):
                continue

            for bk in band_keys:
                if bk in lbl:
                    if "average gross income" in lbl:
                        avg_cols[bk] = idx
                    elif "gross income" in lbl:
                        inc_cols[bk] = idx

        common = [bk for bk in band_keys if bk in inc_cols and bk in avg_cols]
        if len(common) >= 4:
            tot_inc_m = 0.0
            tot_borrowers_k = 0.0
            for bk in common:
                a_m = float(pd.to_numeric(df41.iloc[-1, inc_cols[bk]], errors="coerce") or 0.0)
                b_k = float(pd.to_numeric(df41.iloc[-1, avg_cols[bk]], errors="coerce") or 0.0)
                if b_k > 2000:
                    b_k /= 1000.0
                if a_m > 0 and b_k > 0:
                    tot_inc_m += a_m
                    tot_borrowers_k += (a_m / b_k)
            if tot_borrowers_k > 0:
                avg_nzd = (tot_inc_m / tot_borrowers_k) * 1000.0
                return df41["date"].iloc[-1], avg_nzd
    return None, None


def build_rbnz_dti_income_tickers(c40_path: pathlib.Path | None, c41_path: pathlib.Path | None) -> dict:
    out = {}
    if c40_path is not None:
        try:
            tabs40 = load_rbnz_sheets(c40_path)
            val_pct, dt_dti = None, None

            for _, df40, _, labels40 in tabs40:
                low_idxs, gran_hi_idxs, sub_gt6_idxs, unk_idxs = [], [], [], []
                for idx, lbl in enumerate(labels40):
                    if "new commitments to first home buyers" not in lbl:
                        continue
                    if any(x in lbl for x in ("auckland", "80", "lvr", "%", "share")):
                        continue
                    if any(lo in lbl for lo in ("tdti <= 3", "3 < tdti <= 4", "4 < tdti <= 5", "5 < tdti <= 6")):
                        low_idxs.append(idx)
                    elif "unknown" in lbl:
                        unk_idxs.append(idx)
                    elif "tdti > 6" in lbl:
                        sub_gt6_idxs.append((idx, "excl" in lbl))
                    elif any(hi in lbl for hi in ("6 < tdti", "7 < tdti", "8 < tdti", "9 < tdti", "tdti > 7", "tdti > 8", "tdti > 9", "tdti > 10")):
                        gran_hi_idxs.append(idx)

                if len(low_idxs) == 4 and (gran_hi_idxs or sub_gt6_idxs):
                    low_v = sum(float(pd.to_numeric(df40.iloc[-1, c_i], errors="coerce") or 0.0) for c_i in low_idxs)
                    unk_v = sum(float(pd.to_numeric(df40.iloc[-1, c_i], errors="coerce") or 0.0) for c_i in unk_idxs)
                    if gran_hi_idxs:
                        hi_v = sum(float(pd.to_numeric(df40.iloc[-1, c_i], errors="coerce") or 0.0) for c_i in gran_hi_idxs) + unk_v
                    else:
                        s_idx, is_excl_unk = sub_gt6_idxs[0]
                        hi_v = float(pd.to_numeric(df40.iloc[-1, s_idx], errors="coerce") or 0.0) + (unk_v if is_excl_unk else 0.0)
                    tot_v = low_v + hi_v
                    if tot_v > 0 and 0.1 <= (hi_v / tot_v) * 100.0 <= 45.0:
                        val_pct = (hi_v / tot_v) * 100.0
                        dt_dti = df40["date"].iloc[-1]
                        break

            if val_pct is not None and dt_dti is not None:
                m_full = dt_dti.strftime("%B %Y")
                diff_cap = 20.0 - val_pct
                if diff_cap >= 3.0:
                    dti_html = (
                        f"Most first-home lending currently sits well inside the Reserve Bank's Debt-to-Income limits—in {m_full}, "
                        f"<strong>{val_pct:.1f}%</strong> of first-home buyer lending had a <strong>DTI ratio above 6x</strong>, "
                        f"sitting <strong>{diff_cap:.1f} percentage points</strong> below RBNZ's 20% bank speed-limit allowance."
                    )
                elif diff_cap >= 0.5:
                    dti_html = (
                        f"In {m_full}, <strong>{val_pct:.1f}%</strong> of first-home buyer lending was approved with a "
                        f"<strong>Debt-to-Income (DTI) ratio above 6x</strong>—sitting <strong>{diff_cap:.1f} percentage points</strong> "
                        f"below RBNZ's 20% bank speed-limit allowance."
                    )
                else:
                    dti_html = (
                        f"In {m_full}, <strong>{val_pct:.1f}%</strong> of total first-home lending recorded a "
                        f"<strong>Debt-to-Income (DTI) ratio above 6x</strong>, alongside the Reserve Bank's 20% speed limit on non-exempt loans."
                    )
                out["dti-headroom"] = {
                    "chip": "RBNZ C40",
                    "as_of": month_end_iso(dt_dti),
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": dti_html,
                    "source_html": (
                        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use" '
                        'target="_blank" rel="noopener">RBNZ Table C40</a>, &copy; Reserve Bank of New Zealand. Derived by HLF. '
                        'Historical market statistics for general information only.'
                    ),
                }
        except Exception as exc:
            print(f"[C40 DTI] Parse note: {exc}")

    if c41_path is not None:
        try:
            tabs41 = load_rbnz_sheets(c41_path)
            dt_fhb, fhb_inc = compute_c41_weighted_avg_income(tabs41, "first home", "nz")
            _, fhb_akl_inc = compute_c41_weighted_avg_income(tabs41, "first home", "akl")
            _, inv_inc = compute_c41_weighted_avg_income(tabs41, "investor", "nz")

            src_c41 = (
                'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi" '
                'target="_blank" rel="noopener">RBNZ Table C41</a>, &copy; Reserve Bank of New Zealand. Derived by HLF for general information only.'
            )

            if dt_fhb is not None and fhb_inc and 40000 <= fhb_inc <= 600000:
                m_c41 = dt_fhb.strftime("%B %Y")
                as_of_c41 = month_end_iso(dt_fhb)

                if inv_inc and 40000 <= inv_inc <= 900000:
                    bgi_html = (
                        f"Because mortgage applications often combine two incomes, approved borrower earnings sit higher than single-earner "
                        f"averages—in {m_c41}, approved first-home buyers recorded an average gross borrower income of "
                        f"<strong>${fhb_inc/1000.0:,.1f}k</strong>, compared with <strong>${inv_inc/1000.0:,.1f}k</strong> for property investors."
                    )
                else:
                    bgi_html = (
                        f"Because mortgage applications often combine two incomes, approved first-home buyers in {m_c41} recorded an "
                        f"average gross borrower income of <strong>${fhb_inc/1000.0:,.1f}k</strong> nationwide."
                    )
                out["borrower-income"] = {
                    "chip": "RBNZ C41",
                    "as_of": as_of_c41,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": bgi_html,
                    "source_html": src_c41,
                }

                akl_clause = (
                    f" (and <strong>${fhb_akl_inc:,.0f}</strong> in Auckland)"
                    if fhb_akl_inc and 50000 <= fhb_akl_inc <= 600000
                    else ""
                )
                out["fhb-income-akl-vs-nz"] = {
                    "chip": "RBNZ C41",
                    "as_of": as_of_c41,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"Because most first-home applications combine two incomes, approved first-home buyers in {m_c41} averaged "
                        f"<strong>${fhb_inc:,.0f}</strong> in combined pre-tax income nationwide{akl_clause}—though single-income "
                        f"buyers regularly qualify using regional price points, boarder income, or Kāinga Ora support."
                    ),
                    "source_html": src_c41,
                }
        except Exception as exc:
            print(f"[C41 BGI] Parse note: {exc}")
    return out


def build_rbnz_b30_ticker(filepath: pathlib.Path) -> dict:
    tabs = load_rbnz_sheets(filepath)
    dates, r_1y, _ = find_rbnz_series(
        tabs, ("NCFA.MAM1PR.D1",), ("1 year",), ("value", "amount", "share", "investor"), val_bounds=(1.0, 15.0)
    )
    _, r_2y, _ = find_rbnz_series(
        tabs, ("NCFA.MAM1PR.D2",), ("2 year",), ("value", "amount", "share", "investor"), val_bounds=(1.0, 15.0)
    )
    dt = dates.iloc[-1]
    m_full = dt.strftime("%B %Y")
    v1, v2 = float(r_1y.iloc[-1]), float(r_2y.iloc[-1])

    return {
        "settled-rates": {
            "chip": "RBNZ B30",
            "as_of": month_end_iso(dt),
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"Advertised bank rates only tell half the story—across all new fixed mortgages actually settled in {m_full}, "
                f"New Zealand borrowers locked in a weighted-average rate of <strong>{v1:.2f}%</strong> for one year and "
                f"<strong>{v2:.2f}%</strong> for two years."
            ),
            "source_html": (
                'Source: <a href="https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-weighted-average-interest-rates" '
                'target="_blank" rel="noopener">RBNZ Table B30</a>, &copy; Reserve Bank of New Zealand. Historical market statistics for general information only.'
            ),
        }
    }


def build_stats_nz_tickers(session) -> dict:
    out = {}
    today = datetime.datetime.now(datetime.timezone.utc).date()

    bc_candidates = []
    for back in range(0, 4):
        y, m = today.year, today.month - back
        while m <= 0:
            m += 12
            y -= 1
        dt_cand = datetime.date(y, m, 1)
        m_slug = dt_cand.strftime("%B-%Y").lower()
        bc_candidates.append((dt_cand, f"https://www.stats.govt.nz/information-releases/building-consents-issued-{m_slug}/"))

    q_months = [(12, "december"), (9, "september"), (6, "june"), (3, "march")]
    cpi_candidates = [
        (yr, qm, qname, f"https://www.stats.govt.nz/information-releases/consumers-price-index-{qname}-{yr}-quarter/")
        for yr in (today.year, today.year - 1)
        for qm, qname in q_months
        if datetime.date(yr, qm, 1) <= today
    ]
    cpi_candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    cpi_candidates = cpi_candidates[:3]

    all_urls = [u for _, u in bc_candidates] + [u for _, _, _, u in cpi_candidates]
    fetched_html = fetch_html_with_browser_fallback(session, all_urls, "Stats NZ")

    for dt_cand, url in bc_candidates:
        html = fetched_html.get(url, "")
        if not html:
            continue
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))
        m_ann = re.search(
            r"year ended\s+[A-Za-z]+\s+\d{4},\s+the actual number of new dwellings consented was\s+([\d,]+),\s+(up|down|unchanged)?\s*([\d.]+)?",
            text,
            flags=re.I,
        )
        m_mth = re.search(
            r"there were\s+([\d,]+)\s+new dwellings consented,\s+comprising:\s*([\d,]+)\s+stand-alone houses[,;\s]+([\d,]+)\s+townhouses,\s+flats,\s+and units",
            text,
            flags=re.I,
        )
        if m_ann and m_mth:
            ann_tot = int(m_ann.group(1).replace(",", ""))
            dir_word = (m_ann.group(2) or "").lower()
            ann_pct = float(m_ann.group(3)) if m_ann.group(3) else 0.0
            mth_tot = int(m_mth.group(1).replace(",", ""))
            houses = int(m_mth.group(2).replace(",", ""))
            townhouses = int(m_mth.group(3).replace(",", ""))
            multi_tot = max(0, mth_tot - houses)
            multi_pct = (multi_tot / mth_tot) * 100.0
            m_full = dt_cand.strftime("%B %Y")
            as_of = month_end_iso(pd.Timestamp(dt_cand))

            if dir_word == "up" and ann_pct >= 2.0:
                nb_html = (
                    f"New-build supply across New Zealand has been gathering pace, with <strong>{ann_tot:,}</strong> new homes consented "
                    f"in the year to {m_full} (up <strong>{ann_pct:.1f}%</strong> on the previous year, including <strong>{mth_tot:,}</strong> "
                    f"dwellings consented in {m_full})."
                )
            elif dir_word == "down" and ann_pct >= 2.0:
                nb_html = (
                    f"New-build consenting across New Zealand has been running quieter over the past year, with <strong>{ann_tot:,}</strong> "
                    f"homes consented in the 12 months to {m_full} (down <strong>{ann_pct:.1f}%</strong> on a year earlier, including "
                    f"<strong>{mth_tot:,}</strong> dwellings in {m_full})."
                )
            else:
                nb_html = (
                    f"Across New Zealand, councils consented <strong>{ann_tot:,}</strong> new homes in the 12 months to {m_full}, "
                    f"including <strong>{mth_tot:,}</strong> dwellings in {m_full}."
                )

            if multi_pct >= 50.0:
                th_html = (
                    f"Townhouses and multi-unit developments now make up more than half of new residential consents—accounting for "
                    f"<strong>{multi_pct:.1f}%</strong> (<strong>{multi_tot:,}</strong> of <strong>{mth_tot:,}</strong> homes) consented in {m_full}, "
                    f"including <strong>{townhouses:,}</strong> townhouses, flats, and units."
                )
            else:
                th_html = (
                    f"Townhouses and multi-unit developments remain a major fixture of New Zealand's new housing stock—making up "
                    f"<strong>{multi_pct:.1f}%</strong> (<strong>{multi_tot:,}</strong> of <strong>{mth_tot:,}</strong> homes) consented in {m_full}, "
                    f"including <strong>{townhouses:,}</strong> townhouses, flats, and units."
                )

            src_bc = (
                f'Source: <a href="{url}" target="_blank" rel="noopener">Stats NZ — Building Consents Issued ({m_full})</a>, '
                'licensed under CC BY 4.0. Multi-unit share derived by HLF. Historical market statistics only.'
            )
            out["new-builds-annual"] = {
                "chip": "Stats NZ",
                "as_of": as_of,
                "max_age_days": MAX_AGE_MONTHLY,
                "html": nb_html,
                "source_html": src_bc,
            }
            out["townhouse-share"] = {
                "chip": "Stats NZ",
                "as_of": as_of,
                "max_age_days": MAX_AGE_MONTHLY,
                "html": th_html,
                "source_html": src_bc,
            }
            print(f"[Stats NZ Consents] Parsed {m_full}: {ann_tot:,} annual, {multi_pct:.1f}% multi-unit")
            break

    for yr, qm, qname, cpi_url in cpi_candidates:
        html = fetched_html.get(cpi_url, "")
        if not html:
            continue
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))
        m_cpi_all = re.search(r"CPI all groups\s+[-+]?[\d.]+\s+([-+]?[\d.]+)", text, flags=re.I)
        cpi_all_val = float(m_cpi_all.group(1)) if m_cpi_all else None

        m_build = re.search(
            r"(?:purchase of housing|purchase of new housing|cost of building a new home)[,:\s]+(up|down|fell|rose)?\s*([\d.]+)\s+percent",
            text,
            flags=re.I,
        )
        if m_build:
            direction = (m_build.group(1) or "up").lower()
            is_down = direction in ("down", "fell")
            val = float(m_build.group(2))
            signed_val = -val if is_down else val
            q_dt = datetime.date(yr, qm, 1)
            cpi_clause = f", alongside <strong>{cpi_all_val:.1f}%</strong> overall CPI inflation" if cpi_all_val is not None else ""

            if signed_val <= 4.5:
                bc_inf_html = (
                    f"Residential build-cost inflation has settled into single digits compared with the post-pandemic construction surge—"
                    f"Stats NZ recorded a <strong>{'-' if is_down else '+'}{val:.1f}%</strong> annual change in the cost of purchasing a newly built "
                    f"home in the {qname.capitalize()} {yr} quarter{cpi_clause}."
                )
            else:
                bc_inf_html = (
                    f"Construction pricing remains an active factor in new-build budgets, with Stats NZ recording a "
                    f"<strong>+{val:.1f}%</strong> annual increase in the cost of purchasing a newly built home in the "
                    f"{qname.capitalize()} {yr} quarter{cpi_clause}."
                )

            out["build-cost-inflation"] = {
                "chip": "Stats NZ CPI",
                "as_of": month_end_iso(pd.Timestamp(q_dt)),
                "max_age_days": MAX_AGE_QUARTERLY,
                "html": bc_inf_html,
                "source_html": (
                    f'Source: <a href="{cpi_url}" target="_blank" rel="noopener">Stats NZ — Consumers Price Index ({qname.capitalize()} {yr} quarter)</a>, '
                    'licensed under CC BY 4.0. Historical market statistics only.'
                ),
            }
            print(f"[Stats NZ CPI] Parsed {qname.capitalize()} {yr}: {'-' if is_down else '+'}{val:.1f}% YoY")
            break

    return out


def build_tenancy_tickers(session) -> dict:
    page_bytes = fetch_url_bytes(session, TENANCY_PAGE, timeout=30)
    if not page_bytes:
        raise RuntimeError("Could not reach Tenancy Services rental bond page.")
    html = page_bytes.decode("utf-8", errors="ignore")
    csv_url = ""
    for href in re.findall(r'href=["\']([^"\']*\.csv[^"\']*)["\']', html, flags=re.I):
        if "region" in href.lower():
            csv_url = "https://www.tenancy.govt.nz" + href if href.startswith("/") else href
            break
    if not csv_url:
        raise RuntimeError("Could not locate By-Region CSV link on Tenancy Services.")

    csv_bytes = fetch_url_bytes(session, csv_url, referer=TENANCY_PAGE, timeout=45)
    df = pd.read_csv(io.BytesIO(csv_bytes))
    df.columns = [c.strip() for c in df.columns]
    df["dt"] = pd.to_datetime(df["Time Frame"], errors="coerce", format="mixed")
    for col in ("Lodged Bonds", "Active Bonds", "Closed Bonds", "Median Rent"):
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "").str.strip(), errors="coerce")

    nat = df[df["Location"].astype(str).str.strip().str.upper() == "ALL"].dropna(subset=["dt", "Median Rent"]).sort_values("dt")
    last = nat.iloc[-1]
    dt = last["dt"]
    m_full = dt.strftime("%B %Y")
    as_of = month_end_iso(dt)
    nat_r = int(round(float(last["Median Rent"])))

    yoy_rows = nat[nat["dt"] == (dt - pd.DateOffset(years=1))]
    yoy_pct = ((float(last["Median Rent"]) - float(yoy_rows.iloc[-1]["Median Rent"])) / float(yoy_rows.iloc[-1]["Median Rent"]) * 100.0) if len(yoy_rows) else 0.0
    yoy_phrase = f" ({plain_yoy_growth(float(last['Median Rent']), float(yoy_rows.iloc[-1]['Median Rent']))})" if len(yoy_rows) else ""

    def reg_med(substr: str) -> int:
        sub = df[(df["dt"] == dt) & (df["Location"].astype(str).str.contains(substr, case=False, na=False))]
        return int(round(float(sub.iloc[-1]["Median Rent"]))) if len(sub) else 0

    akl_r = reg_med("Auckland")
    cant_r = reg_med("Canterbury")
    wgtn_r = reg_med("Wellington")

    if abs(yoy_pct) < 2.0:
        rent_opener = f"Private weekly rents across New Zealand have been relatively steady over the past year—the national median sat at <strong>${nat_r}/wk</strong> in {m_full}{yoy_phrase}"
    elif yoy_pct >= 2.0:
        rent_opener = f"Private weekly rents have continued to edge higher over the past year, reaching a national median of <strong>${nat_r}/wk</strong> in {m_full}{yoy_phrase}"
    else:
        rent_opener = f"Private weekly rents have eased over the past year to a national median of <strong>${nat_r}/wk</strong> in {m_full}{yoy_phrase}"

    src_mbie = (
        'Source: <a href="https://www.tenancy.govt.nz/about-tenancy-services/data-and-statistics/rental-bond-data/" '
        'target="_blank" rel="noopener">Ministry of Business, Innovation and Employment (Tenancy Services Rental Bond Data)</a>, '
        'licensed under CC BY 3.0 NZ. Derived by HLF. Historical market statistics only.'
    )

    return {
        "rent-yield-pulse": {
            "chip": "Tenancy Services",
            "as_of": as_of,
            "max_age_days": MAX_AGE_TENANCY,
            "html": (
                f"{rent_opener}, with regional medians at <strong>${akl_r}/wk</strong> in Auckland, "
                f"<strong>${wgtn_r}/wk</strong> in Wellington, and <strong>${cant_r}/wk</strong> in Canterbury."
            ),
            "source_html": src_mbie,
        },
        "rental-bonds-supply": {
            "chip": "Tenancy Services",
            "as_of": as_of,
            "max_age_days": MAX_AGE_TENANCY,
            "html": (
                f"Thousands of New Zealand households sign new leases every month—Tenancy Services logged "
                f"<strong>{int(round(float(last['Lodged Bonds']))):,}</strong> newly lodged private rental bonds in {m_full}, "
                f"across a nationwide pool of <strong>{int(round(float(last['Active Bonds']))):,}</strong> active private tenancies."
            ),
            "source_html": src_mbie,
        },
    }


def build_ird_tickers(session, c31_by_ym: dict | None = None) -> dict:
    out = {}

    sl_bytes = fetch_url_bytes(session, IRD_SL_BAL_PAGE, timeout=30)
    if sl_bytes:
        sl_text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", sl_bytes.decode("utf-8", errors="ignore")))
        m_med = re.search(
            r"Median student loan balance\s+(?:increased|decreased)\s+by\s+[\d.]+%\s+from\s+\$([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})\s+to\s+\$([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})",
            sl_text,
            flags=re.I,
        )
        if m_med:
            prev_bal = int(m_med.group(1).replace(",", ""))
            curr_bal = int(m_med.group(3).replace(",", ""))
            dt_sl = pd.to_datetime("1 " + m_med.group(4), errors="coerce")
            if pd.notna(dt_sl):
                m_full = dt_sl.strftime("%B %Y")
                out["student-loan-servicing"] = {
                    "chip": "Inland Revenue",
                    "as_of": month_end_iso(dt_sl),
                    "max_age_days": MAX_AGE_QUARTERLY,
                    "html": (
                        f"Student debt is a familiar line item on home loan applications—across Inland Revenue's {m_full} figures, "
                        f"the median NZ student loan balance sat at <strong>${curr_bal:,}</strong> ({plain_yoy_growth(curr_bal, prev_bal)}), "
                        f"where compulsory PAYE deductions factor into how banks test monthly servicing."
                    ),
                    "source_html": (
                        f'Source: <a href="{IRD_SL_BAL_PAGE}" target="_blank" rel="noopener">Inland Revenue — Student Loan Statistics ({dt_sl.strftime("%b %Y")})</a>, '
                        'licensed under CC BY 4.0. Servicing context by HLF for general information only.'
                    ),
                }

    num_bytes = fetch_url_bytes(session, IRD_KS_NUM_PAGE, timeout=30)
    amt_bytes = fetch_url_bytes(session, IRD_KS_AMT_PAGE, timeout=30)
    if num_bytes and amt_bytes:
        num_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", num_bytes.decode("utf-8", errors="ignore")))
        amt_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", amt_bytes.decode("utf-8", errors="ignore")))
        mn = re.search(
            r"In\s+([A-Za-z]+\s+\d{4}):.*?([\d,]+)\s+members have withdrawn their savings for the purpose of first home purchase",
            num_txt,
            flags=re.I,
        )
        ma = re.search(
            r"In\s+([A-Za-z]+\s+\d{4}):.*?\$([\d,.]+)\s+million was attributable to first home purchase(?:,\s+(?:up|down)\s+from\s+\$([\d,.]+)\s+million)?",
            amt_txt,
            flags=re.I,
        )
        if mn and ma:
            dt_ks = pd.to_datetime("1 " + ma.group(1), errors="coerce")
            if pd.notna(dt_ks):
                cnt = int(mn.group(2).replace(",", ""))
                amt_m = float(ma.group(2).replace(",", ""))
                avg_mem_k = int(round((amt_m * 1000.0) / cnt)) if cnt > 0 else 0
                m_full = dt_ks.strftime("%B %Y")
                ym_ks = dt_ks.strftime("%Y-%m")

                cross_note = f", or roughly <strong>${avg_mem_k:,}k</strong> per withdrawing member"
                has_rbnz_xref = False
                if c31_by_ym and ym_ks in c31_by_ym and c31_by_ym[ym_ks]["fhb_cnt"] > 0:
                    ks_per_loan_k = int(round((amt_m * 1000.0) / c31_by_ym[ym_ks]["fhb_cnt"]))
                    cross_note = (
                        f"—averaging <strong>~${avg_mem_k:,}k</strong> per withdrawing member, or roughly "
                        f"<strong>${ks_per_loan_k:,}k</strong> of KiwiSaver funds per approved first-home mortgage"
                    )
                    has_rbnz_xref = True

                out["kiwisaver-first-home"] = {
                    "chip": "IRD & RBNZ" if has_rbnz_xref else "Inland Revenue",
                    "as_of": month_end_iso(dt_ks),
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"For many buyers, KiwiSaver does much of the heavy lifting toward a first deposit—in {m_full}, "
                        f"<strong>{cnt:,}</strong> members withdrew a combined <strong>${amt_m:,.1f}m</strong> to put toward a "
                        f"first home{cross_note}."
                    ),
                    "source_html": (
                        f'Source: <a href="{IRD_KS_DATA_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Withdrawal Statistics</a>'
                        + (' cross-referenced with RBNZ Table C31' if has_rbnz_xref else '')
                        + ', licensed under CC BY 4.0. Derived by HLF for general information only.'
                    ),
                }

    # Regional KiwiSaver First-Home Withdrawals (16 Regions)
    reg_bytes = fetch_url_bytes(session, IRD_KS_REG_PAGE, timeout=30, use_cb=False)
    if reg_bytes:
        try:
            reg_html = reg_bytes.decode("utf-8", errors="ignore")
            m_yr = re.search(r"year ended 30 June (\d{4})", reg_html, flags=re.I)
            fy_yr_reg = int(m_yr.group(1)) if m_yr else datetime.datetime.now(datetime.timezone.utc).year
            ext = SimpleTableExtractor()
            ext.feed(reg_html)
            if len(ext.tables) >= 2:
                amt_map, cnt_map = {}, {}
                for row in ext.tables[0][1:]:
                    if len(row) >= 3:
                        r_name = row[0].strip()
                        v_str = re.sub(r"[^\d.]", "", row[-1])
                        if r_name and v_str and r_name.lower() not in ("other", "total", "region"):
                            amt_map[r_name.lower()] = (r_name, float(v_str))
                for row in ext.tables[1][1:]:
                    if len(row) >= 3:
                        r_name = row[0].strip()
                        c_str = re.sub(r"[^\d.]", "", row[-1])
                        if r_name and c_str and r_name.lower() in amt_map:
                            cnt_map[r_name.lower()] = int(round(float(c_str)))

                reg_list = []
                tot_nzd, tot_mem = 0.0, 0
                for k_low, (disp, nzd_val) in amt_map.items():
                    c_val = cnt_map.get(k_low, 0)
                    if c_val > 0:
                        tot_nzd += nzd_val
                        tot_mem += c_val
                        clean_disp = disp.replace("Bay Of Plenty", "Bay of Plenty").replace("Hawkes Bay", "Hawke's Bay")
                        reg_list.append((clean_disp, int(round(nzd_val / c_val))))

                if tot_mem > 0 and len(reg_list) >= 10:
                    reg_list.sort(key=lambda x: x[1], reverse=True)
                    nz_avg_p = int(round(tot_nzd / tot_mem))
                    nz_couple = int(round((nz_avg_p * 2) / 50.0) * 50)
                    wgtn_avg = dict(reg_list).get("Wellington", reg_list[0][1])
                    akl_avg = dict(reg_list).get("Auckland", reg_list[1][1])
                    out["fhb-regional-kiwisaver"] = {
                        "chip": "Inland Revenue",
                        "as_of": f"{fy_yr_reg}-06-30",
                        "max_age_days": MAX_AGE_ANNUAL,
                        "html": (
                            f"Pooling two KiwiSaver accounts is how most Kiwi couples build their deposit—over the year to June {fy_yr_reg}, "
                            f"<strong>{tot_mem:,} first-home buyers</strong> withdrew <strong>${tot_nzd/1_000_000_000.0:.2f}B</strong> from KiwiSaver, "
                            f"averaging <strong>${nz_avg_p:,} per person nationwide</strong> (<strong>~${nz_couple:,} per couple</strong>), "
                            f"led by <strong>Wellington (${wgtn_avg:,}/person)</strong> and <strong>Auckland (${akl_avg:,}/person)</strong>."
                        ),
                        "source_html": (
                            f'Source: <a href="{IRD_KS_REG_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Fund Withdrawals by Region (Year ended Jun {fy_yr_reg})</a>, '
                            'licensed under CC BY 4.0. Per-member averages derived by HLF for general information only.'
                        ),
                    }
        except Exception as exc:
            print(f"[IRD Regional KiwiSaver] Parse note: {exc}")

    tr_bytes = fetch_url_bytes(session, IRD_KS_TR_PAGE, timeout=30)
    if tr_bytes:
        tr_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", tr_bytes.decode("utf-8", errors="ignore")))
        mt = re.search(
            r"year ended\s+30\s+June\s+(\d{4}).*?scheme transfers\s+(?:increased|decreased),?\s+from\s+[\d,]+\s+in the\s+\d{4}\s+financial year to\s+([\d,]+)\s+in the\s+(\d{4})\s+financial year",
            tr_txt,
            flags=re.I,
        )
        if mt:
            fy_yr = int(mt.group(3))
            tr_tot = int(mt.group(2).replace(",", ""))
            mth_avg = int(round(tr_tot / 12.0))
            sw_item = {
                "chip": "Inland Revenue",
                "as_of": f"{fy_yr}-06-30",
                "max_age_days": MAX_AGE_ANNUAL,
                "html": (
                    f"Reviewing and changing KiwiSaver providers happens more regularly than many members realise—over the year to "
                    f"June {fy_yr}, <strong>{tr_tot:,}</strong> New Zealanders moved their balance to a new scheme, averaging "
                    f"<strong>{mth_avg:,}</strong> switches every month."
                ),
                "source_html": (
                    f'Source: <a href="{IRD_KS_TR_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Scheme Transfer Statistics</a> '
                    '(CC BY 4.0). Monthly average derived by HLF. General information only, not personalised KiwiSaver advice.'
                ),
            }
            out["kiwisaver-switches"] = sw_item
            out["kiwisaver-active-switchers"] = dict(sw_item)

    return out


# ---------------------------------------------------------------------------
# Main Orchestrator & Strict Per-Ticker Freshness Gate
# ---------------------------------------------------------------------------
def main() -> None:
    root = pathlib.Path(__file__).resolve().parent
    json_path = root / "nz-mortgage-pulse.json"
    html_path = root / "index.html"

    prev_tickers = {}
    if json_path.exists():
        try:
            prev_tickers = json.loads(json_path.read_text(encoding="utf-8")).get("tickers", {})
        except Exception:
            prev_tickers = {}

    session = get_http_session()
    fresh_candidates = dict(prev_tickers)
    live_parsed_count = 0

    rbnz_files = download_rbnz_workbooks(root, session)
    c31_by_ym = None

    if "c31" in rbnz_files:
        try:
            c31_out, c31_by_ym = build_rbnz_c31_tickers(rbnz_files["c31"])
            fresh_candidates.update(c31_out)
            live_parsed_count += len(c31_out)
        except Exception as exc:
            print(f"[RBNZ C31] Error: {exc}")

    if "c33" in rbnz_files:
        try:
            c33_out = build_rbnz_c33_tickers(rbnz_files["c33"])
            fresh_candidates.update(c33_out)
            live_parsed_count += len(c33_out)
        except Exception as exc:
            print(f"[RBNZ C33] Error: {exc}")

    if "c32" in rbnz_files:
        try:
            c32_out = build_rbnz_c32_ticker(rbnz_files["c32"], c31_by_ym)
            fresh_candidates.update(c32_out)
            live_parsed_count += len(c32_out)
        except Exception as exc:
            print(f"[RBNZ C32] Error: {exc}")

    if "c40" in rbnz_files or "c41" in rbnz_files:
        try:
            dti_out = build_rbnz_dti_income_tickers(rbnz_files.get("c40"), rbnz_files.get("c41"))
            fresh_candidates.update(dti_out)
            live_parsed_count += len(dti_out)
        except Exception as exc:
            print(f"[RBNZ C40/C41] Error: {exc}")

    if "b30" in rbnz_files:
        try:
            b30_out = build_rbnz_b30_ticker(rbnz_files["b30"])
            fresh_candidates.update(b30_out)
            live_parsed_count += len(b30_out)
        except Exception as exc:
            print(f"[RBNZ B30] Error: {exc}")

    for p in rbnz_files.values():
        if p.name.startswith("tmp_") and p.exists():
            os.remove(p)

    try:
        snz_out = build_stats_nz_tickers(session)
        fresh_candidates.update(snz_out)
        live_parsed_count += len(snz_out)
    except Exception as exc:
        print(f"[Stats NZ] Error: {exc}")

    try:
        ten_out = build_tenancy_tickers(session)
        fresh_candidates.update(ten_out)
        live_parsed_count += len(ten_out)
    except Exception as exc:
        print(f"[Tenancy Services] Error: {exc}")

    try:
        ird_out = build_ird_tickers(session, c31_by_ym)
        fresh_candidates.update(ird_out)
        live_parsed_count += len(ird_out)
    except Exception as exc:
        print(f"[IRD] Error: {exc}")

    verified_tickers = {}
    for k, item in fresh_candidates.items():
        if not isinstance(item, dict):
            continue
        as_of = item.get("as_of", "")
        max_age = int(item.get("max_age_days", MAX_AGE_MONTHLY))
        if item.get("html") and is_fresh(as_of, max_age):
            verified_tickers[k] = item
        else:
            print(f"[Failsafe Gate] Omitted stale/invalid ticker '{k}' (as_of={as_of}, max_age={max_age}d)")

    if live_parsed_count == 0 and not verified_tickers:
        raise RuntimeError("Zero live datasets parsed and zero valid fresh tickers remain; failing workflow honestly.")

    if verified_tickers == prev_tickers and html_path.exists():
        print(f"All {len(verified_tickers)} verified tickers unchanged; skipping redundant commit.")
        return

    now_utc = datetime.datetime.now(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    payload = {
        "publisher": "HLF — Home Loan Factory (homeloanfactory.co.nz)",
        "generated_utc": now_utc,
        "count": len(verified_tickers),
        "tickers": verified_tickers,
    }
    json_path.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")

    robots_tag = "index, follow" if verified_tickers else "noindex, nofollow"
    rows_html = "".join(
        f"<tr><td><code>{k}</code></td><td><strong>{v['chip']}</strong><br><span style='font-size:0.75rem;color:#6b7280;'>As at {v['as_of']}</span></td>"
        f"<td>{v['html']}<div style='margin-top:4px;font-size:0.76rem;color:#6b7280;'>{v['source_html']}</div></td></tr>"
        for k, v in sorted(verified_tickers.items())
    )
    html_path.write_text(
        f"""<!DOCTYPE html>
<html lang="en-NZ">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="robots" content="{robots_tag}">
  <title>NZ Mortgage, Housing &amp; KiwiSaver Market Pulse — Official Crown Data | HLF</title>
  <style>
    body {{ font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #111827; background: #f8fafc; margin: 0; padding: 24px 14px; line-height: 1.55; }}
    main {{ max-width: 1060px; margin: 0 auto; background: #ffffff; border-top: 5px solid #0b2559; border-radius: 16px; padding: 28px; box-shadow: 0 4px 20px rgba(11,37,89,0.05); }}
    h1 {{ margin: 0 0 8px; color: #0b2559; font-size: 1.4rem; letter-spacing: -0.02em; }}
    p {{ color: #374151; font-size: 0.9rem; margin: 0 0 14px; }}
    a {{ color: #186bd6; font-weight: 600; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; margin-top: 14px; font-variant-numeric: tabular-nums; }}
    th {{ background: #f8fafc; color: #0b2559; text-align: left; padding: 10px 12px; border-bottom: 2px solid #dbeafe; }}
    td {{ padding: 11px 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }}
    code {{ background: #eff6ff; padding: 2px 6px; border-radius: 4px; font-size: 0.78rem; color: #0b2559; }}
  </style>
</head>
<body>
  <main>
    <h1>NZ Mortgage, Housing &amp; KiwiSaver Pulse — Verified Reference Table</h1>
    <p>Machine-readable summary compiled automatically from official New Zealand Crown releases (RBNZ, Stats NZ, MBIE Tenancy Services, and Inland Revenue). Stale or unverified series are automatically excluded. Generated: <code>{now_utc}</code> by <a href="https://www.homeloanfactory.co.nz/">HLF (Home Loan Factory)</a> (<a href="https://www.homeloanfactory.co.nz/public-disclosures">Public Disclosures</a>).</p>
    <table>
      <thead><tr><th>Ticker Key</th><th>Source &amp; Period</th><th>Verified Statistic &amp; Attribution</th></tr></thead>
      <tbody>{rows_html}</tbody>
    </table>
  </main>
</body>
</html>""",
        encoding="utf-8",
    )
    print(f"Saved {len(verified_tickers)} verified fresh tickers to nz-mortgage-pulse.json & index.html.")


if __name__ == "__main__":
    main()
