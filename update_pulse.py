import base64
import datetime
import io
import json
import os
import pathlib
import re
import time
import pandas as pd

# ---------------------------------------------------------------------------
# Official NZ Crown Data Sources (RBNZ, Stats NZ, MBIE Tenancy, Inland Revenue)
# ---------------------------------------------------------------------------
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
IRD_KS_TR_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/schemes/transfer"
IRD_KS_DATA_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/datasets"
IRD_SL_BAL_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/statistics-on-the-student-loan-scheme/statistics-on-student-loan-balances"

# Strict freshness gates (days since observation period end).
# Any ticker older than max_age_days is stripped from JSON + HTML so stale data never renders.
MAX_AGE_MONTHLY = 85
MAX_AGE_TENANCY = 120   # MBIE publishes private bond CSVs on a ~2-month processing lag
MAX_AGE_QUARTERLY = 160
MAX_AGE_ANNUAL = 420    # IRD annual scheme transfer summaries publish each September for June FY


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


def pts_yoy(curr: float, prev: float) -> str:
    diff = curr - prev
    return f"{'+' if diff >= 0 else ''}{diff:.1f} pts YoY"


def pct_yoy(curr: float, prev: float) -> str:
    if not prev:
        return ""
    diff = ((curr - prev) / prev) * 100.0
    return f"{'+' if diff >= 0 else ''}{diff:.1f}% YoY"


# ---------------------------------------------------------------------------
# 3-Tier Anti-Cloudflare & HTTP Download Pipeline
# ---------------------------------------------------------------------------
def get_http_session():
    from curl_cffi import requests as cf_requests
    return cf_requests.Session(impersonate="chrome")


def fetch_url_bytes(session, url: str, referer: str = "", timeout: int = 45) -> bytes:
    headers = {"Cache-Control": "no-cache, no-store, must-revalidate", "Pragma": "no-cache"}
    if referer:
        headers["Referer"] = referer
    try:
        resp = session.get(add_cb(url), headers=headers, timeout=timeout)
        return resp.content if resp.status_code == 200 else b""
    except Exception:
        return b""


def fetch_xlsx_in_browser_page(page, cfg: dict, dest: pathlib.Path, tier_label: str) -> bool:
    for p_url in cfg["pages"]:
        try:
            page.goto(add_cb(p_url), wait_until="domcontentloaded", timeout=60000)
            deadline = time.time() + 35
            while time.time() < deadline:
                if page.locator('a[href*=".xlsx"]').count() > 0 and "just a moment" not in (page.title() or "").lower():
                    break
                page.wait_for_timeout(2500)
            if page.locator('a[href*=".xlsx"]').count() > 0:
                break
        except Exception:
            continue

    target = cfg["xlsx"]
    matched_loc = None
    for link in page.locator('a[href*=".xlsx"]').all():
        href = link.get_attribute("href") or ""
        if cfg["match"] in href.lower():
            target = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
            matched_loc = link
            break

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

    if matched_loc is not None:
        try:
            with page.expect_download(timeout=40000) as dl_info:
                matched_loc.click(force=True)
            dl_info.value.save_as(str(dest))
            if dest.exists() and is_valid_xlsx(dest.read_bytes()):
                print(f"[{tier_label}] Click-downloaded {cfg['match'].upper()}")
                return True
        except Exception as exc:
            print(f"[{tier_label}] Click-download note ({cfg['match']}): {exc}")

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
        from camoufox.sync_api import Camoufox
        with Camoufox(headless=False, humanize=True) as browser:
            page = browser.new_page()
            for key in list(missing):
                out = dest_dir / f"tmp_{key}.xlsx"
                if fetch_xlsx_in_browser_page(page, RBNZ_SOURCES[key], out, "RBNZ Tier 2 - Camoufox"):
                    saved[key] = out
    except Exception as exc:
        print(f"[RBNZ Tier 2 - Camoufox] Note: {exc}")

    missing = [k for k in RBNZ_SOURCES if k not in saved]
    if not missing:
        return saved

    # Tier 3: Headed Google Chrome / Chromium via Playwright under Xvfb
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
            for key in list(missing):
                out = dest_dir / f"tmp_{key}.xlsx"
                if fetch_xlsx_in_browser_page(page, RBNZ_SOURCES[key], out, "RBNZ Tier 3 - Chrome"):
                    saved[key] = out
            ctx.close()
    except Exception as exc:
        print(f"[RBNZ Tier 3 - Chrome] Note: {exc}")

    return saved


# ---------------------------------------------------------------------------
# Multi-Sheet RBNZ Workbook Scanner
# ---------------------------------------------------------------------------
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

        date_col = pd.to_datetime(raw.iloc[:, 0], errors="coerce")
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
            sids.append(col_parts[-1] if col_parts else "")
            labels.append(f"{tab.lower()} | " + " | ".join(p.lower() for p in col_parts))

        df = raw.iloc[first_data_row:].copy()
        df["date"] = pd.to_datetime(df.iloc[:, 0], errors="coerce")
        df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
        parsed_tabs.append((tab, df, sids, labels))
    return parsed_tabs


def find_rbnz_series(tabs: list, exact_sids=(), include_all=(), exclude_any=(), val_bounds=None):
    for _, df, sids, _ in tabs:
        for sid in exact_sids:
            if sid in sids:
                idx = sids.index(sid)
                ser = pd.to_numeric(df.iloc[:, idx], errors="coerce")
                valid = ser.dropna()
                if len(valid) and (val_bounds is None or (val_bounds[0] <= float(valid.iloc[-1]) <= val_bounds[1])):
                    return df["date"], ser, sid
    for _, df, sids, labels in tabs:
        for idx, lbl in enumerate(labels):
            if idx == 0:
                continue
            if all(k.lower() in lbl for k in include_all) and not any(x.lower() in lbl for x in exclude_any):
                ser = pd.to_numeric(df.iloc[:, idx], errors="coerce")
                valid = ser.dropna()
                if len(valid) and (val_bounds is None or (val_bounds[0] <= float(valid.iloc[-1]) <= val_bounds[1])):
                    return df["date"], ser, (sids[idx] or lbl[:45])
    raise KeyError(f"RBNZ series not found: sids={exact_sids}, include={include_all}")


# ---------------------------------------------------------------------------
# Fault-Isolated Parsers (100% Neutral, Value-Agnostic Templates)
# ---------------------------------------------------------------------------
def build_rbnz_c31_tickers(filepath: pathlib.Path) -> tuple[dict, dict]:
    tabs = load_rbnz_sheets(filepath)
    dates, fhb_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AC1",), ("total lending", "first home"))
    _, oo_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AC2",), ("total lending", "other owner"))
    _, inv_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AC3",), ("total lending", "investor"))
    try:
        _, bus_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AC4",), ("total lending", "business"))
    except KeyError:
        bus_val = pd.Series([0.0] * len(dates))

    tot_val = fhb_val.fillna(0) + oo_val.fillna(0) + inv_val.fillna(0) + bus_val.fillna(0)
    _, fhb_hi_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.BC1",), ("80", "lending", "first home"), ("less", "equal"))
    _, fhb_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AC1",), ("borrowers", "first home"), ("80",))
    _, fhb_hi_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.BC1",), ("80", "borrowers", "first home"), ("less", "equal"))

    i, iy = len(dates) - 1, max(0, len(dates) - 13)
    dt = dates.iloc[i]
    m_lbl = dt.strftime("%b %Y")
    as_of = month_end_iso(dt)

    low_cnt_pct = (float(fhb_hi_cnt.iloc[i]) / float(fhb_cnt.iloc[i])) * 100.0
    low_val_pct = (float(fhb_hi_val.iloc[i]) / float(fhb_val.iloc[i])) * 100.0
    low_val_yoy = (float(fhb_hi_val.iloc[iy]) / float(fhb_val.iloc[iy])) * 100.0

    fhb_sh = (float(fhb_val.iloc[i]) / float(tot_val.iloc[i])) * 100.0
    fhb_sh_yoy = (float(fhb_val.iloc[iy]) / float(tot_val.iloc[iy])) * 100.0
    fhb_avg_k = int(round((float(fhb_val.iloc[i]) * 1000.0) / float(fhb_cnt.iloc[i])))

    inv_sh = (float(inv_val.iloc[i]) / float(tot_val.iloc[i])) * 100.0
    inv_sh_yoy = (float(inv_val.iloc[iy]) / float(tot_val.iloc[iy])) * 100.0

    src_c31 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type" '
        'target="_blank" rel="noopener">RBNZ Table C31</a>, &copy; Reserve Bank of New Zealand. Percentages &amp; averages derived by HLF. '
        'Historical market statistics for general information only.'
    )

    out = {
        "fhb-low-deposit": {
            "chip": "RBNZ C31",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_lbl}, <strong>{low_cnt_pct:.1f}%</strong> of First Home Buyer mortgages approved across New Zealand "
                f"(<strong>{low_val_pct:.1f}%</strong> by dollar value, or <strong>${float(fhb_hi_val.iloc[i]):,.0f}m</strong>) "
                f"had <strong>under a 20% deposit</strong> ({pts_yoy(low_val_pct, low_val_yoy)})."
            ),
            "source_html": src_c31,
        },
        "fhb-share": {
            "chip": "RBNZ C31",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"First Home Buyers accounted for <strong>{fhb_sh:.1f}%</strong> of all new NZ mortgage lending in {m_lbl} "
                f"(<strong>{int(round(float(fhb_cnt.iloc[i]))):,}</strong> home loans approved, average loan <strong>${fhb_avg_k:,}k</strong>; {pts_yoy(fhb_sh, fhb_sh_yoy)})."
            ),
            "source_html": src_c31,
        },
        "investors": {
            "chip": "RBNZ C31",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"Property investors accounted for <strong>{inv_sh:.1f}%</strong> of new NZ mortgage lending in {m_lbl} "
                f"(<strong>${float(inv_val.iloc[i])/1000.0:.2f}B</strong>; {pts_yoy(inv_sh, inv_sh_yoy)}), compared with <strong>{fhb_sh:.1f}%</strong> for First Home Buyers."
            ),
            "source_html": src_c31,
        },
    }
    meta = {"ym": dt.strftime("%Y-%m"), "fhb_cnt": float(fhb_cnt.iloc[i]), "fhb_val_m": float(fhb_val.iloc[i])}
    return out, meta


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
    m_lbl = dt.strftime("%b %Y")
    as_of = month_end_iso(dt)

    tot_v = float(top_v.iloc[j] + pur_v.iloc[j] + ref_v.iloc[j] + oth_v.iloc[j])
    tot_c = float(top_c.iloc[j] + pur_c.iloc[j] + ref_c.iloc[j] + oth_c.iloc[j])
    ref_sh = (float(ref_v.iloc[j]) / tot_v) * 100.0
    ref_avg_k = int(round((float(ref_v.iloc[j]) * 1000.0) / float(ref_c.iloc[j])))
    top_sh_c = (float(top_c.iloc[j]) / tot_c) * 100.0
    top_avg_k = int(round((float(top_v.iloc[j]) * 1000.0) / float(top_c.iloc[j])))

    src_c33 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose" '
        'target="_blank" rel="noopener">RBNZ Table C33</a>, &copy; Reserve Bank of New Zealand. Derived by HLF. '
        'Historical market statistics for general information only.'
    )
    return {
        "refinance": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_lbl}, <strong>{int(round(float(ref_c.iloc[j]))):,}</strong> Kiwi borrowers switched mortgage providers—"
                f"totalling <strong>${float(ref_v.iloc[j])/1000.0:.2f}B</strong>, or <strong>{ref_sh:.1f}%</strong> of all new lending "
                f"(average loan switched: <strong>${ref_avg_k:,}k</strong>)."
            ),
            "source_html": src_c33,
        },
        "top-ups": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_lbl}, <strong>{int(round(float(top_c.iloc[j]))):,}</strong> existing NZ homeowners topped up their mortgage—"
                f"representing <strong>{top_sh_c:.1f}%</strong> of all new mortgage approvals by borrower count (average top-up: <strong>${top_avg_k:,}k</strong>)."
            ),
            "source_html": src_c33,
        },
    }


def build_rbnz_c32_ticker(filepath: pathlib.Path) -> dict:
    tabs = load_rbnz_sheets(filepath)
    dates, inv_io, _ = find_rbnz_series(tabs, (), ("investor", "interest only"), ("80", "less", "higher", "existing"))
    _, inv_pi, _ = find_rbnz_series(tabs, (), ("investor", "principal"), ("80", "less", "higher", "existing"))
    _, oo_io, _ = find_rbnz_series(tabs, (), ("owner", "interest only"), ("80", "less", "higher", "existing"))
    _, oo_pi, _ = find_rbnz_series(tabs, (), ("owner", "principal"), ("80", "less", "higher", "existing"))

    k = len(dates) - 1
    dt = dates.iloc[k]
    m_lbl = dt.strftime("%b %Y")
    inv_io_pct = (float(inv_io.iloc[k]) / float(inv_io.iloc[k] + inv_pi.iloc[k])) * 100.0
    oo_io_pct = (float(oo_io.iloc[k]) / float(oo_io.iloc[k] + oo_pi.iloc[k])) * 100.0

    return {
        "interest-only": {
            "chip": "RBNZ C32",
            "as_of": month_end_iso(dt),
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_lbl}, <strong>{inv_io_pct:.1f}%</strong> of new NZ property investor lending was structured on <strong>Interest-Only</strong> terms "
                f"(including revolving credit), compared with <strong>{oo_io_pct:.1f}%</strong> for owner-occupiers."
            ),
            "source_html": (
                'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type" '
                'target="_blank" rel="noopener">RBNZ Table C32</a>, &copy; Reserve Bank of New Zealand. Derived by HLF. '
                'Historical market statistics for general information only.'
            ),
        }
    }


def build_rbnz_dti_income_tickers(c40_path: pathlib.Path | None, c41_path: pathlib.Path | None) -> dict:
    out = {}
    if c40_path is not None:
        try:
            tabs40 = load_rbnz_sheets(c40_path)
            val_pct = None
            dt_dti = None
            try:
                dates, fhb_hi_dti, _ = find_rbnz_series(
                    tabs40, (), ("first home", "6", "%"), ("auckland", "less", "under"), val_bounds=(0.1, 45.0)
                )
                dt_dti = dates.iloc[-1]
                val_pct = float(fhb_hi_dti.iloc[-1])
            except KeyError:
                pass

            if val_pct is None:
                for _, df40, _, labels40 in tabs40:
                    tot_idx = None
                    over6_idxs = []
                    for idx, lbl in enumerate(labels40):
                        if "first home" not in lbl or "auckland" in lbl or "80" in lbl:
                            continue
                        if "total" in lbl and not any(b in lbl for b in ("5", "6", "7", "8", "9", "unknown")):
                            tot_idx = idx
                        elif any(h in lbl for h in ("> 6", ">6", "6 to", "7 to", "> 7", ">7", "8 to", "> 8", ">8", "9 to", "> 9", ">9", "over 6", "above 6", "greater than 6")):
                            over6_idxs.append(idx)
                    if tot_idx is not None and over6_idxs:
                        tot_v = float(pd.to_numeric(df40.iloc[-1, tot_idx], errors="coerce") or 0.0)
                        hi_v = sum(float(pd.to_numeric(df40.iloc[-1, c_i], errors="coerce") or 0.0) for c_i in over6_idxs)
                        if tot_v > 0 and 0.1 <= (hi_v / tot_v) * 100.0 <= 45.0:
                            val_pct = (hi_v / tot_v) * 100.0
                            dt_dti = df40["date"].iloc[-1]
                            break

            if val_pct is not None and dt_dti is not None:
                m_lbl = dt_dti.strftime("%b %Y")
                diff_cap = 20.0 - val_pct
                if diff_cap >= 0.5:
                    cap_clause = f", sitting <strong>{diff_cap:.1f} percentage points</strong> below RBNZ's 20% bank speed-limit allowance"
                elif diff_cap >= -0.5:
                    cap_clause = ", sitting right alongside RBNZ's 20% bank speed-limit allowance"
                else:
                    cap_clause = f", compared with RBNZ's 20% bank speed-limit allowance"

                out["dti-headroom"] = {
                    "chip": "RBNZ C40",
                    "as_of": month_end_iso(dt_dti),
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"In {m_lbl}, <strong>{val_pct:.1f}%</strong> of First Home Buyer lending was approved with a "
                        f"<strong>Debt-to-Income (DTI) ratio above 6x</strong>{cap_clause}."
                    ),
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
            dates41, fhb_bgi, _ = find_rbnz_series(tabs41, (), ("first home", "average"), ("auckland", "tdti", "dti"))
            _, inv_bgi, _ = find_rbnz_series(tabs41, (), ("investor", "average"), ("auckland", "tdti", "dti"))
            dt = dates41.iloc[-1]
            fhb_inc = float(fhb_bgi.iloc[-1])
            inv_inc = float(inv_bgi.iloc[-1])
            if fhb_inc < 2000:
                fhb_inc *= 1000.0
                inv_inc *= 1000.0
            if 40000 <= fhb_inc <= 600000:
                out["borrower-income"] = {
                    "chip": "RBNZ C41",
                    "as_of": month_end_iso(dt),
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"Approved NZ First Home Buyers recorded an average gross borrower income of <strong>${fhb_inc/1000.0:,.0f}k</strong> "
                        f"in {dt.strftime('%b %Y')} (across single and joint applications), compared with <strong>${inv_inc/1000.0:,.0f}k</strong> for property investors."
                    ),
                    "source_html": (
                        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi" '
                        'target="_blank" rel="noopener">RBNZ Table C41</a>, &copy; Reserve Bank of New Zealand. Historical market statistics for general information only.'
                    ),
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
    m_lbl = dt.strftime("%b %Y")
    v1, v2 = float(r_1y.iloc[-1]), float(r_2y.iloc[-1])

    return {
        "settled-rates": {
            "chip": "RBNZ B30",
            "as_of": month_end_iso(dt),
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"Across settled new NZ mortgages in {m_lbl}, borrowers achieved a weighted-average rate of "
                f"<strong>{v1:.2f}%</strong> on 1-year fixed terms and <strong>{v2:.2f}%</strong> on 2-year fixed terms."
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

    # 1. Monthly Building Consents Issued
    for back in range(0, 4):
        y, m = today.year, today.month - back
        while m <= 0:
            m += 12
            y -= 1
        dt_cand = datetime.date(y, m, 1)
        m_slug = dt_cand.strftime("%B-%Y").lower()
        url = f"https://www.stats.govt.nz/information-releases/building-consents-issued-{m_slug}/"
        raw = fetch_url_bytes(session, url, timeout=25)
        if not raw:
            continue
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", raw.decode("utf-8", errors="ignore")))

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
            yoy_lbl = "0.0% YoY" if dir_word == "unchanged" else f"{'-' if dir_word == 'down' else '+'}{ann_pct:.1f}% YoY"

            mth_tot = int(m_mth.group(1).replace(",", ""))
            houses = int(m_mth.group(2).replace(",", ""))
            townhouses = int(m_mth.group(3).replace(",", ""))
            multi_tot = max(0, mth_tot - houses)
            multi_pct = (multi_tot / mth_tot) * 100.0
            m_lbl = dt_cand.strftime("%b %Y")
            as_of = month_end_iso(pd.Timestamp(dt_cand))

            src_bc = (
                f'Source: <a href="{url}" target="_blank" rel="noopener">Stats NZ — Building Consents Issued ({m_lbl})</a>, '
                'licensed under CC BY 4.0. Multi-unit share derived by HLF. Historical market statistics only.'
            )
            out["new-builds-annual"] = {
                "chip": "Stats NZ",
                "as_of": as_of,
                "max_age_days": MAX_AGE_MONTHLY,
                "html": (
                    f"<strong>{ann_tot:,}</strong> new homes were consented across New Zealand in the year ended {m_lbl} "
                    f"(<strong>{yoy_lbl}</strong>), including <strong>{mth_tot:,}</strong> dwellings consented in {m_lbl}."
                ),
                "source_html": src_bc,
            }
            out["townhouse-share"] = {
                "chip": "Stats NZ",
                "as_of": as_of,
                "max_age_days": MAX_AGE_MONTHLY,
                "html": (
                    f"Multi-unit dwellings accounted for <strong>{multi_pct:.1f}%</strong> (<strong>{multi_tot:,}</strong> of <strong>{mth_tot:,}</strong>) "
                    f"of new NZ homes consented in {m_lbl}, including <strong>{townhouses:,}</strong> townhouses, flats, and units."
                ),
                "source_html": src_bc,
            }
            print(f"[Stats NZ Consents] Parsed {m_lbl}: {ann_tot:,} annual, {multi_pct:.1f}% multi-unit")
            break

    # 2. Quarterly CPI — New-Build Cost Change ("purchase of housing")
    q_months = [(12, "december"), (9, "september"), (6, "june"), (3, "march")]
    candidates = [
        (yr, qm, qname)
        for yr in (today.year, today.year - 1)
        for qm, qname in q_months
        if datetime.date(yr, qm, 1) <= today
    ]
    candidates.sort(reverse=True)

    for yr, qm, qname in candidates[:3]:
        cpi_url = f"https://www.stats.govt.nz/information-releases/consumers-price-index-{qname}-{yr}-quarter/"
        raw = fetch_url_bytes(session, cpi_url, timeout=25)
        if not raw:
            continue
        html_str = raw.decode("utf-8", errors="ignore")
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html_str))

        m_cpi_all = re.search(r"CPI all groups\s+[-+]?[\d.]+\s+([-+]?[\d.]+)", text, flags=re.I)
        cpi_all_val = float(m_cpi_all.group(1)) if m_cpi_all else None

        m_build = re.search(
            r"(?:purchase of housing|purchase of new housing|cost of building a new home)[,:\s]+(up|down|fell|rose)?\s*([\d.]+)\s+percent",
            text,
            flags=re.I,
        )
        if m_build:
            direction = (m_build.group(1) or "up").lower()
            sign = "-" if direction in ("down", "fell") else "+"
            val = float(m_build.group(2))
            q_dt = datetime.date(yr, qm, 1)
            cpi_clause = ""
            if cpi_all_val is not None:
                c_sign = "+" if cpi_all_val >= 0 else ""
                cpi_clause = f", compared with <strong>{c_sign}{cpi_all_val:.1f}% YoY</strong> headline CPI inflation"
            out["build-cost-inflation"] = {
                "chip": "Stats NZ CPI",
                "as_of": month_end_iso(pd.Timestamp(q_dt)),
                "max_age_days": MAX_AGE_QUARTERLY,
                "html": (
                    f"Annual NZ new-build price change (CPI <em>purchase of housing</em>) recorded "
                    f"<strong>{sign}{val:.1f}% YoY</strong> in the {qname.capitalize()} {yr} quarter{cpi_clause}."
                ),
                "source_html": (
                    f'Source: <a href="{cpi_url}" target="_blank" rel="noopener">Stats NZ — Consumers Price Index ({qname.capitalize()} {yr} quarter)</a>, '
                    'licensed under CC BY 4.0. Historical market statistics only.'
                ),
            }
            print(f"[Stats NZ CPI] Parsed {qname.capitalize()} {yr}: {sign}{val:.1f}% YoY")
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
    df["dt"] = pd.to_datetime(df["Time Frame"], errors="coerce")
    for col in ("Lodged Bonds", "Active Bonds", "Closed Bonds", "Median Rent"):
        df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "").str.strip(), errors="coerce")

    nat = df[df["Location"].astype(str).str.strip().str.upper() == "ALL"].dropna(subset=["dt", "Median Rent"]).sort_values("dt")
    last = nat.iloc[-1]
    dt = last["dt"]
    m_lbl = dt.strftime("%b %Y")
    as_of = month_end_iso(dt)

    yoy_rows = nat[nat["dt"] == (dt - pd.DateOffset(years=1))]
    yoy_clause = f" ({pct_yoy(float(last['Median Rent']), float(yoy_rows.iloc[-1]['Median Rent']))})" if len(yoy_rows) else ""

    def reg_med(substr: str) -> int:
        sub = df[(df["dt"] == dt) & (df["Location"].astype(str).str.contains(substr, case=False, na=False))]
        return int(round(float(sub.iloc[-1]["Median Rent"]))) if len(sub) else 0

    akl_r = reg_med("Auckland")
    cant_r = reg_med("Canterbury")
    wgtn_r = reg_med("Wellington")

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
                f"NZ median private weekly rent stood at <strong>${int(round(float(last['Median Rent'])))}/wk</strong> in {m_lbl}{yoy_clause}—"
                f"with regional medians at <strong>${akl_r}/wk</strong> in Auckland, "
                f"<strong>${wgtn_r}/wk</strong> in Wellington, and <strong>${cant_r}/wk</strong> in Canterbury."
            ),
            "source_html": src_mbie,
        },
        "rental-bonds-supply": {
            "chip": "Tenancy Services",
            "as_of": as_of,
            "max_age_days": MAX_AGE_TENANCY,
            "html": (
                f"Tenancy Services recorded <strong>{int(round(float(last['Lodged Bonds']))):,}</strong> newly lodged private rental bonds in {m_lbl}, "
                f"across <strong>{int(round(float(last['Active Bonds']))):,}</strong> active private tenancies nationwide."
            ),
            "source_html": src_mbie,
        },
    }


def build_ird_tickers(session, rbnz_c31_meta: dict | None = None) -> dict:
    out = {}

    # 1. IRD Student Loan Balances (Quarterly)
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
                out["student-loan-servicing"] = {
                    "chip": "Inland Revenue",
                    "as_of": month_end_iso(dt_sl),
                    "max_age_days": MAX_AGE_QUARTERLY,
                    "html": (
                        f"The median NZ student loan balance stood at <strong>${curr_bal:,}</strong> as at {dt_sl.strftime('%b %Y')} "
                        f"(<strong>{pct_yoy(curr_bal, prev_bal)}</strong>), across New Zealand and overseas-based borrowers."
                    ),
                    "source_html": (
                        f'Source: <a href="{IRD_SL_BAL_PAGE}" target="_blank" rel="noopener">Inland Revenue — Student Loan Statistics ({dt_sl.strftime("%b %Y")})</a>, '
                        'licensed under CC BY 4.0. Historical statistics for general information only.'
                    ),
                }

    # 2. IRD Monthly KiwiSaver First-Home Withdrawals (Number + Amount pages)
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
                prev_m = float(ma.group(3).replace(",", "")) if ma.group(3) else 0.0
                yoy_s = f" ({pct_yoy(amt_m, prev_m)})" if prev_m > 0 else ""
                m_lbl = dt_ks.strftime("%b %Y")
                cross_note = ""
                if rbnz_c31_meta and rbnz_c31_meta.get("ym") == dt_ks.strftime("%Y-%m") and rbnz_c31_meta.get("fhb_cnt", 0) > 0:
                    ks_per_loan_k = int(round((amt_m * 1000.0) / rbnz_c31_meta["fhb_cnt"]))
                    cross_note = f"—equating to <strong>~${ks_per_loan_k:,}k</strong> of KiwiSaver funds per approved first-home mortgage"

                out["kiwisaver-first-home"] = {
                    "chip": "IRD & RBNZ" if cross_note else "Inland Revenue",
                    "as_of": month_end_iso(dt_ks),
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"<strong>{cnt:,}</strong> KiwiSaver members withdrew a combined <strong>${amt_m:,.1f}m</strong> in {m_lbl}{yoy_s} "
                        f"for a first-home purchase{cross_note}."
                    ),
                    "source_html": (
                        f'Source: <a href="{IRD_KS_DATA_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Withdrawal Statistics</a>'
                        + (' cross-referenced with RBNZ Table C31' if cross_note else '')
                        + ', licensed under CC BY 4.0. Derived by HLF. General information only.'
                    ),
                }

    # 3. IRD KiwiSaver Scheme Transfers (Annual financial-year release)
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
            out["kiwisaver-switches"] = {
                "chip": "Inland Revenue",
                "as_of": f"{fy_yr}-06-30",
                "max_age_days": MAX_AGE_ANNUAL,
                "html": (
                    f"<strong>{tr_tot:,}</strong> KiwiSaver members transferred their balance to a new KiwiSaver scheme provider "
                    f"in the year ended June {fy_yr} (an average of <strong>{mth_avg:,}</strong> scheme switches per month)."
                ),
                "source_html": (
                    f'Source: <a href="{IRD_KS_TR_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Scheme Transfer Statistics</a> '
                    '(CC BY 4.0). Monthly average derived by HLF. General information only, not personalised KiwiSaver advice.'
                ),
            }

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
    c31_meta = None

    if "c31" in rbnz_files:
        try:
            c31_out, c31_meta = build_rbnz_c31_tickers(rbnz_files["c31"])
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
            c32_out = build_rbnz_c32_ticker(rbnz_files["c32"])
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
        ird_out = build_ird_tickers(session, c31_meta)
        fresh_candidates.update(ird_out)
        live_parsed_count += len(ird_out)
    except Exception as exc:
        print(f"[IRD] Error: {exc}")

    # STRICT FAILSAFE GATE: Drop any ticker that failed validation or exceeded max_age_days
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
