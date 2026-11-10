import base64
import datetime
import html as html_lib
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
# Official NZ Crown Data Sources & Creator / Publisher Metadata
# ---------------------------------------------------------------------------
PULSE_BASE_URL = "https://pulse-data.homeloanfactory.co.nz"
RBNZ_WARMUP_URL = "https://www.rbnz.govt.nz/statistics"

CREATOR_NAME = "Andy Palliser"
CREATOR_ALT_NAME = "Andrew Palliser"
CREATOR_ROLE = "Mortgage & KiwiSaver Adviser · Home Loan Factory"
CREATOR_BADGE = "Creator of HLF Pulse"
CREATOR_FSP = "FSP1004725"
CREATOR_URL = "https://www.homeloanfactory.co.nz/team/andrew-palliser"
CREATOR_IMAGE = "https://cdn.prod.website-files.com/6a1baf404596ba888865b209/6a1ce4c85096daa2e99a0ecb_Andrew-Palliser.avif"
PUBLISHER_NAME = "HLF (Home Loan Factory)"
PUBLISHER_LEGAL = "HLF Financial Limited"
PUBLISHER_URL = "https://www.homeloanfactory.co.nz/"
PULSE_DATE_CREATED = "2026-09-01"

CREATOR_SCHEMA_OBJ = {
    "@type": "Person",
    "name": CREATOR_NAME,
    "alternateName": CREATOR_ALT_NAME,
    "jobTitle": "Mortgage & KiwiSaver Adviser",
    "identifier": CREATOR_FSP,
    "url": CREATOR_URL,
    "image": CREATOR_IMAGE,
    "worksFor": {
        "@type": "Organization",
        "name": "Home Loan Factory",
        "legalName": PUBLISHER_LEGAL,
        "url": PUBLISHER_URL,
    },
}

PUBLISHER_SCHEMA_OBJ = {
    "@type": "Organization",
    "name": PUBLISHER_NAME,
    "legalName": PUBLISHER_LEGAL,
    "url": PUBLISHER_URL,
}

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

MAX_AGE_MONTHLY = 85
MAX_AGE_TENANCY = 120
MAX_AGE_QUARTERLY = 160
MAX_AGE_ANNUAL = 420

# Category order for the clickable "Browse all 30 datasets" index
CATEGORY_ORDER = [
    "First-home buyers",
    "Mortgage lending & rates",
    "Property investors & rents",
    "Homeowners & equity",
    "KiwiSaver",
    "Housing supply & construction",
]

# ---------------------------------------------------------------------------
# Data Dictionary Catalog (Exactly 30 Canonical Public Datasets)
# ---------------------------------------------------------------------------
TICKER_CATALOG = {
    "fhb-low-deposit": {
        "name": "First-home buyers using under 20% deposit (>80% LVR)",
        "also_known_as": "High-LVR (>80% LVR) first-home buyer share / Low-deposit FHB mortgage share",
        "category": "First-home buyers",
        "topic": "First-home lending",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.BC1, LVRN.MMB2.AC1, LVRN.MMB1.BC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived share & HLF-calculated comparison (from officially reported RBNZ loan counts & dollar values)",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The HLF-derived percentage of firm residential mortgage commitments to New Zealand first-home buyers where the Loan-to-Value Ratio (LVR) exceeds 80% (corresponding to under a 20% deposit), measured by borrower count alongside officially reported >80% LVR dollar lending.",
        "methodology": "Compiled from monthly registered bank returns under RBNZ Table C31. First-home buyers are defined by lending banks as borrowers purchasing a residential property to occupy as a first home.",
        "hlf_calculation": "(High-LVR FHB Borrowers [LVRN.MMB2.BC1] ÷ Total FHB Borrowers [LVRN.MMB2.AC1]) × 100",
        "hlf_interpretation": "HLF market context (not measured in C31): RBNZ Table C31 reports all >80% LVR commitments as a single aggregate and does not disaggregate the underlying lending channel (such as bank low-equity speed-limit allocations, Kāinga Ora First Home Loans, or LVR-exempt new builds).",
        "questions": [
            "How many first-home buyers have less than a 20% deposit in New Zealand?",
            "Is a 20% deposit actually typical for NZ first-home buyers?",
            "How common are low-deposit mortgages in New Zealand?",
            "What percentage of first-home buyers borrow with less than a 20% deposit?",
            "Are low-deposit first-home loans becoming more common?",
        ],
        "article_title": "First-Home Buyer Guides & Low-Deposit Insights",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Kāinga Ora First Home Loan & LMI Estimator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kainga-ora-lmi-fee-estimator",
    },
    "fhb-share": {
        "name": "First-home buyer share of total New Zealand mortgage lending",
        "also_known_as": "FHB market share of residential mortgage lending",
        "category": "First-home buyers",
        "topic": "First-home lending",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.AC1, LVRN.MMB1.AA, LVRN.MMB2.AC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported loan count & HLF-derived market share / average loan size",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The HLF-derived proportion of total new residential mortgage lending dollars in New Zealand committed to first-home buyers, alongside the officially reported monthly count of approved first-home loans and HLF-derived average loan size.",
        "methodology": "Sourced from RBNZ Table C31 monthly registered bank commitments across first-home buyers, other owner-occupiers, property investors, and business purposes.",
        "hlf_calculation": "(FHB Lending NZDm [LVRN.MMB1.AC1] ÷ Total Lending NZDm [LVRN.MMB1.AA]) × 100",
        "hlf_interpretation": "HLF market context: Tracks the proportion of new bank mortgage commitments going to first-home buyers relative to other owner-occupiers, investors, and business borrowers.",
        "questions": [
            "What share of New Zealand mortgages go to first-home buyers?",
            "How many first-home buyer loans are approved each month in NZ?",
            "What is the average first-home buyer mortgage size in New Zealand?",
        ],
        "article_title": "NZ Mortgage Market Updates & Insights",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-loan-size-gap": {
        "name": "First-home buyer average mortgage size by LVR tier (>80% LVR vs <=80% LVR)",
        "also_known_as": "Low-deposit (<20% deposit) vs standard-deposit (20%+ deposit) FHB average loan size",
        "category": "First-home buyers",
        "topic": "First-home lending",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.BC1, LVRN.MMB2.BC1, LVRN.MMB1.CC1, LVRN.MMB2.CC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived cohort averages & HLF-calculated comparison",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "HLF-derived comparison of the mean approved mortgage size for New Zealand first-home buyers at >80% LVR (under 20% deposit) versus first-home buyers at <=80% LVR (20%+ deposit).",
        "methodology": "Derived from RBNZ Table C31 by dividing total monthly dollar lending in each LVR tier by the number of approved first-home borrowers in that same LVR tier.",
        "hlf_calculation": "(LVRN.MMB1.BC1 × $1m ÷ LVRN.MMB2.BC1) minus (LVRN.MMB1.CC1 × $1m ÷ LVRN.MMB2.CC1)",
        "hlf_interpretation": "HLF market context: The data shows that first-home buyers at >80% LVR take out larger mortgages on average than those at <=80% LVR. RBNZ Table C31 does not record property purchase prices or borrower motivations; factors can include borrowing a higher percentage of the purchase price, regional price mix, or household income.",
        "questions": [
            "Do low-deposit first-home buyers borrow more or less than 20% deposit buyers?",
            "What is the average low-deposit first-home loan size in NZ?",
            "What is the average mortgage for a first-home buyer with a 20% deposit?",
        ],
        "article_title": "Buying Sooner With a Low Deposit vs Waiting for 20%",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "fhb-bank-lvr-priority": {
        "name": "First-home buyer share of all >80% LVR mortgage lending",
        "also_known_as": "FHB share of high-LVR bank lending",
        "category": "First-home buyers",
        "topic": "First-home lending",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.BC1, LVRN.MMB1.BA, LVRN.MMB2.BC1, LVRN.MMB2.BA",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived share (from officially reported RBNZ high-LVR totals)",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The HLF-derived share of all >80% LVR residential mortgage lending across New Zealand committed to first-home buyers.",
        "methodology": "Compares high-LVR first-home commitments against high-LVR commitments across all borrower types combined in RBNZ Table C31.",
        "hlf_calculation": "(FHB >80% LVR Lending [LVRN.MMB1.BC1] ÷ All Borrower >80% LVR Lending [LVRN.MMB1.BA]) × 100",
        "hlf_interpretation": "HLF market context: Shows that first-home buyers account for the majority of >80% LVR commitments in RBNZ Table C31. Table C31 includes both speed-limit lending and LVR-exempt lending (such as Kāinga Ora First Home Loans and new builds).",
        "questions": [
            "What share of bank >80% LVR lending goes to first-home buyers in New Zealand?",
            "Do New Zealand banks allocate most low-deposit loans to first-home buyers?",
        ],
        "article_title": "How Bank LVR Speed Limits Work in New Zealand",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-annual-milestone": {
        "name": "Annual rolling 12-month first-home buyer mortgage approvals",
        "also_known_as": "Annual NZ first-home buyer loan count & volume",
        "category": "First-home buyers",
        "topic": "First-home lending",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.AC1, LVRN.MMB1.AC1, LVRN.MMB2.BC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived 12-month rolling sum & HLF-calculated comparison",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Monthly",
        "definition": "The HLF-derived cumulative number and dollar value of first-home buyer mortgages approved across New Zealand over the trailing 12 months, including the number approved at >80% LVR.",
        "methodology": "Rolling 12-month sum of monthly RBNZ Table C31 firm commitments to first-home buyers.",
        "hlf_calculation": "12-month trailing sum of LVRN.MMB2.AC1 (borrowers), LVRN.MMB1.AC1 (NZDm), and LVRN.MMB2.BC1 (>80% LVR borrowers)",
        "hlf_interpretation": "HLF market context: Summing 12 consecutive monthly RBNZ C31 releases smooths out summer/winter seasonality to show annual commitment volume.",
        "questions": [
            "How many first-home buyers get approved for a mortgage each year in New Zealand?",
            "How many Kiwis buy their first home with under a 20% deposit per year?",
        ],
        "article_title": "NZ First-Home Buyer Market Trends",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "NZ Sale & Purchase Agreement Walkthrough",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-sale-and-purchase-agreement-walkthrough",
    },
    "fhb-20-deposit-club": {
        "name": "First-home buyers approved at <=80% LVR (20% or greater deposit)",
        "also_known_as": "Standard-equity (<=80% LVR) first-home buyer share",
        "category": "First-home buyers",
        "topic": "First-home lending",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.CC1, LVRN.MMB2.AC1, LVRN.MMB1.CC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported count/value & HLF-derived share/average",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The officially reported number and dollar value, plus HLF-derived percentage share and average loan size, of New Zealand first-home buyers approved at an LVR of 80% or less (corresponding to a deposit of 20% or more).",
        "methodology": "Sourced from RBNZ Table C31 series for <=80% LVR first-home buyer lending and borrower counts.",
        "hlf_calculation": "(<=80% LVR FHB Borrowers [LVRN.MMB2.CC1] ÷ Total FHB Borrowers [LVRN.MMB2.AC1]) × 100",
        "hlf_interpretation": "HLF market context (not measured in C31): An LVR of <=80% (a 20%+ deposit) can improve access to advertised special-rate pricing and generally avoids low-equity margins, subject to individual lender policy.",
        "questions": [
            "What percentage of NZ first-home buyers have a 20% deposit?",
            "What is the average mortgage size for a first-home buyer with a 20% deposit?",
        ],
        "article_title": "How a 20% Deposit Affects Mortgage Rates & Margins",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Compare NZ Mortgage Rates",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
    },
    "fhb-vs-investors": {
        "name": "First-home buyers vs property investors monthly lending comparison",
        "also_known_as": "FHB vs Investor monthly mortgage volume",
        "category": "First-home buyers",
        "topic": "First-home & investor lending",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.AC1, LVRN.MMB1.AN, LVRN.MMB1.AA",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported totals & HLF-calculated comparison",
        "derived_by_hlf": True,
        "unit": "NZDm",
        "frequency": "Monthly",
        "definition": "HLF-calculated comparison of monthly new residential mortgage lending committed to first-home buyers versus residential property investors in RBNZ Table C31.",
        "methodology": "Compares RBNZ Table C31 total monthly lending to first-home buyers (LVRN.MMB1.AC1) against total monthly lending to property investors (LVRN.MMB1.AN).",
        "hlf_calculation": "FHB Lending [LVRN.MMB1.AC1] minus Investor Lending [LVRN.MMB1.AN]",
        "hlf_interpretation": "HLF market context: Compares the dollar volume of new bank mortgage commitments between first-home buyers and residential property investors in the latest monthly release.",
        "questions": [
            "Are first-home buyers or property investors borrowing more in New Zealand?",
            "How does first-home buyer lending compare with investor lending in NZ?",
        ],
        "article_title": "First-Home Buyers vs Property Investors in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "investors": {
        "name": "Property investor share of new New Zealand mortgage lending",
        "also_known_as": "Residential investor mortgage market share",
        "category": "Property investors & rents",
        "topic": "Property investment lending",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.AN, LVRN.MMB1.AA",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported dollar lending & HLF-derived market share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The HLF-derived percentage share and officially reported dollar value of total new residential mortgage lending in New Zealand committed to residential property investors.",
        "methodology": "Sourced from RBNZ Table C31 monthly registered bank commitments to residential property investors.",
        "hlf_calculation": "(Investor Lending NZDm [LVRN.MMB1.AN] ÷ All Borrower Lending NZDm [LVRN.MMB1.AA]) × 100",
        "hlf_interpretation": "HLF market context: Measures the share of new residential mortgage commitments attributed to property investors relative to owner-occupiers.",
        "questions": [
            "What percentage of NZ mortgage lending goes to property investors?",
            "How much are property investors borrowing each month in New Zealand?",
        ],
        "article_title": "NZ Property Investment & Lending Insights",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "owner-equity-buffer": {
        "name": "Existing owner-occupiers approved at <=80% LVR (20% or more equity)",
        "also_known_as": "Standard-equity (<=80% LVR) existing homeowner share",
        "category": "Homeowners & equity",
        "topic": "Homeowner equity & refinancing",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.CC2, LVRN.MMB2.AC2, LVRN.MMB1.CC2",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported counts/values & HLF-derived share/average",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The HLF-derived percentage and average loan size of 'Other owner-occupier' mortgage approvals (movers, refinancers, and top-up borrowers, excluding first-home buyers) written at an LVR of 80% or less.",
        "methodology": "Derived from RBNZ Table C31 'Other owner-occupiers' series at or below 80% LVR.",
        "hlf_calculation": "(<=80% LVR Other Owner-Occupiers [LVRN.MMB2.CC2] ÷ Total Other Owner-Occupiers [LVRN.MMB2.AC2]) × 100",
        "hlf_interpretation": "HLF market context: An LVR of <=80% on a mortgage commitment corresponds to holding at least 20% security equity in the mortgaged property. Note that RBNZ C31 only measures homeowners who apply for a new loan, refinance, or top-up during the month, not unmortgaged homeowners.",
        "questions": [
            "What percentage of existing New Zealand homeowners have at least 20% equity when refinancing or moving?",
            "What is the average mortgage size for an existing homeowner at <=80% LVR in NZ?",
        ],
        "article_title": "How to Calculate & Use Your Usable Home Equity",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "investor-equity-tiers": {
        "name": "Property investor LVR & equity tier breakdown (<=60% LVR, 60%–70% LVR, >70% LVR)",
        "also_known_as": "Investor LVR distribution (40%+ equity, 30%–40% equity, <30% equity equivalents)",
        "category": "Property investors & rents",
        "topic": "Property investment & equity",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.TN, LVRN.MMB2.FN, LVRN.MMB2.DN, LVRN.MMB2.AN",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived LVR tier distribution (from RBNZ C31 investor LVR memo items)",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "HLF-derived distribution of approved New Zealand residential property investors across three LVR tiers in RBNZ Table C31: <=60% LVR, 60% < LVR <= 70%, and >70% LVR.",
        "methodology": "Derived from RBNZ Table C31 investor memo series tracking LVR thresholds at 60% and 70%.",
        "hlf_calculation": "<=60% LVR = LVRN.MMB2.TN ÷ AN; 60–70% LVR = (LVRN.MMB2.FN - DN) ÷ AN; >70% LVR = LVRN.MMB2.DN ÷ AN",
        "hlf_interpretation": "HLF market context: An LVR of <=60% corresponds to 40%+ deposit/equity, 60%–70% LVR corresponds to 30%–40% deposit/equity, and >70% LVR corresponds to under 30% deposit/equity. In practice, RBNZ LVR settings and individual lender policy determine how much equity an investor needs.",
        "questions": [
            "What proportion of NZ property investors borrow at >60% LVR (less than 40% equity)?",
            "What percentage of NZ property investors borrow in the 60%–70% LVR (30%–40% equity) band vs >70% LVR?",
        ],
        "article_title": "Using Home Equity to Buy an Investment Property",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Usable Equity & Investment Cashflow Calculators",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "investor-low-deposit": {
        "name": "Property investor loans approved at >80% LVR (<20% deposit) & 70%–80% LVR",
        "also_known_as": "High-LVR (>80% LVR & 70%–80% LVR) property investor lending",
        "category": "Property investors & rents",
        "topic": "Property investment & new builds",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.BN, LVRN.MMB2.BN, LVRN.MMB1.DN, LVRN.MMB2.DN",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported >80% LVR counts/values & HLF-derived 70%–80% LVR tier",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Monthly",
        "definition": "The officially reported number and dollar value of residential property investor mortgages approved at >80% LVR (under 20% deposit), alongside the HLF-derived 70% < LVR <= 80% band (20%–30% deposit).",
        "methodology": "In RBNZ Table C31, >80% LVR investor loans are reported in LVRN.MMB2.BN / LVRN.MMB1.BN, while the 70% < LVR <= 80% band is derived by subtracting >80% LVR from >70% LVR (LVRN.MMB2.DN minus LVRN.MMB2.BN).",
        "hlf_calculation": ">80% LVR = LVRN.MMB2.BN; 70%–80% LVR = LVRN.MMB2.DN minus LVRN.MMB2.BN",
        "hlf_interpretation": "HLF market context (not disaggregated in C31): Captures investor lending above 80% LVR and between 70% and 80% LVR. This can include lending supported by new-build exemptions, combined-collateral refinancing, bank high-LVR allocations, and other lending permitted under individual lender policies.",
        "questions": [
            "How many property investor loans are approved at >80% LVR (under a 20% deposit) in NZ?",
            "How many investor loans sit in the 70%–80% LVR (20%–30% deposit) band?",
        ],
        "article_title": "New-Build & Low-Deposit Rules for Property Investors",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Turnkey vs. Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/turnkey-vs-construction-loan-calculator-nz",
    },
    "investor-new-build-lvr": {
        "name": "Property investors approved at >70% LVR (under 30% deposit)",
        "also_known_as": "Higher than 70% LVR property investor share",
        "category": "Property investors & rents",
        "topic": "Property investment & new builds",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.DN, LVRN.MMB2.DN, LVRN.MMB1.BN, LVRN.MMB2.BN",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported >70% LVR totals & HLF-derived share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The officially reported number and dollar value, plus HLF-derived percentage share, of New Zealand property investor mortgages approved at >70% LVR (corresponding to under a 30% deposit).",
        "methodology": "Sourced from RBNZ Table C31 Memo Item H2/D2 (Higher than 70% LVR investor borrowers and lending).",
        "hlf_calculation": "(>70% LVR Investor Borrowers [LVRN.MMB2.DN] ÷ Total Investor Borrowers [LVRN.MMB2.AN]) × 100",
        "hlf_interpretation": "HLF market context (not disaggregated in C31): Captures investor lending above the 70% LVR threshold. This can include lending supported by new-build exemptions, bank high-LVR allocations, and other lending permitted under individual lender policies.",
        "questions": [
            "What percentage of NZ property investors borrow at >70% LVR (less than a 30% deposit)?",
            "How many investor loans are approved above 70% LVR each month?",
        ],
        "article_title": "New-Build LVR Exemptions for Property Investors",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "mover-low-equity": {
        "name": "Existing owner-occupiers approved at >80% LVR (under 20% equity)",
        "also_known_as": "High-LVR (>80% LVR) other owner-occupier share",
        "category": "Homeowners & equity",
        "topic": "Next-home buyers & homeowner equity",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.BC2, LVRN.MMB1.BC2, LVRN.MMB2.AC2",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported counts/values & HLF-derived share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The officially reported number and dollar value, plus HLF-derived percentage share, of 'Other owner-occupiers' (non-first-home buyers moving, refinancing, or borrowing) approved at >80% LVR.",
        "methodology": "Sourced from RBNZ Table C31 series B3/F3 (Higher than 80% LVR other owner-occupiers).",
        "hlf_calculation": "(>80% LVR Other Owner-Occupiers [LVRN.MMB2.BC2] ÷ Total Other Owner-Occupiers [LVRN.MMB2.AC2]) × 100",
        "hlf_interpretation": "HLF market context (not disaggregated in C31): Captures non-first-home owner-occupier lending above 80% LVR, which can include new-build purchases, bridging/relocation lending, and bank speed-limit allocations.",
        "questions": [
            "Can existing homeowners move or borrow at >80% LVR (less than 20% equity) in NZ?",
            "How many existing NZ owner-occupiers are approved with >80% LVR each month?",
        ],
        "article_title": "Upgrading Your Home With Less Than 20% Equity",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Net Proceeds & Next Home Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/net-proceeds-and-next-home-calculator",
    },
    "refinance": {
        "name": "Monthly mortgage refinancing (borrowers changing loan provider)",
        "also_known_as": "Change in loan provider / Bank-to-bank mortgage switching",
        "category": "Mortgage lending & rates",
        "topic": "Refinancing & bank switching",
        "source": "RBNZ C33",
        "source_full": "Reserve Bank of New Zealand Table C33",
        "series_ids": "LVRN.MMB1.AZD, LVRN.MMB2.AZD",
        "source_url": RBNZ_SOURCES["c33"]["pages"][0],
        "data_type": "Officially reported refinance volume/count & HLF-derived share/average",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The officially reported monthly number of borrowers and dollar value of residential mortgages changing loan provider in RBNZ Table C33, plus the HLF-derived share of total lending and average loan size.",
        "methodology": "Compiled from RBNZ Table C33 (New residential mortgage lending by purpose — Change in loan provider).",
        "hlf_calculation": "(Refinance Lending NZDm [LVRN.MMB1.AZD] ÷ Total Purpose Lending NZDm) × 100",
        "hlf_interpretation": "HLF market context: Measures external refinancing where a borrower switches from one mortgage provider to another (internal refixes with the same bank are not counted in Table C33).",
        "questions": [
            "How many New Zealand homeowners switch banks or refinance each month?",
            "What is the average mortgage size for borrowers refinancing in NZ?",
        ],
        "article_title": "When Does Refinancing Your Mortgage Make Sense?",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Refinance & Cashback Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/refinance-calculator",
    },
    "top-ups": {
        "name": "Monthly mortgage top-ups on existing properties (borrower share & average top-up)",
        "also_known_as": "Mortgage top-up borrower share & average size",
        "category": "Homeowners & equity",
        "topic": "Homeowner equity & top-ups",
        "source": "RBNZ C33",
        "source_full": "Reserve Bank of New Zealand Table C33",
        "series_ids": "LVRN.MMB2.AZB, LVRN.MMB1.AZB",
        "source_url": RBNZ_SOURCES["c33"]["pages"][0],
        "data_type": "Officially reported top-up count & HLF-derived share/average",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The officially reported monthly number of New Zealand borrowers approved for a top-up on their existing mortgage in RBNZ Table C33, alongside the HLF-derived share of total approvals and average top-up amount.",
        "methodology": "Compiled from RBNZ Table C33 (New residential mortgage lending by purpose — Top-ups).",
        "hlf_calculation": "(Top-Up Borrowers [LVRN.MMB2.AZB] ÷ Total Purpose Borrowers) × 100",
        "hlf_interpretation": "HLF market context (not disaggregated in C33): RBNZ Table C33 records top-ups as a single purpose category without disaggregating how borrowers use the funds (such as renovations, debt consolidation, or investment deposits).",
        "questions": [
            "How common are mortgage top-ups in New Zealand?",
            "What is the average mortgage top-up size in NZ?",
        ],
        "article_title": "How Mortgage Top-Ups Work for Renovations & Upgrades",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "equity-topup-demand": {
        "name": "Monthly mortgage top-up lending",
        "also_known_as": "Total monthly mortgage top-up dollar lending (RBNZ C33)",
        "category": "Homeowners & equity",
        "topic": "Homeowner equity & top-ups",
        "source": "RBNZ C33",
        "source_full": "Reserve Bank of New Zealand Table C33",
        "series_ids": "LVRN.MMB1.AZB, LVRN.MMB2.AZB",
        "source_url": RBNZ_SOURCES["c33"]["pages"][0],
        "data_type": "Officially reported dollar lending & HLF-derived average",
        "derived_by_hlf": True,
        "unit": "NZDm",
        "frequency": "Monthly",
        "definition": "Total monthly dollar value of new residential mortgage top-up commitments approved for existing borrowers in RBNZ Table C33.",
        "methodology": "Sourced from RBNZ Table C33 series LVRN.MMB1.AZB (Top-up lending value) and LVRN.MMB2.AZB (Top-up borrower count).",
        "hlf_calculation": "Officially reported LVRN.MMB1.AZB dollar total; Avg top-up = LVRN.MMB1.AZB × $1m ÷ LVRN.MMB2.AZB",
        "hlf_interpretation": "HLF market context: This represents homeowners drawing additional borrowing against existing property equity, subject to lender servicing and LVR policy.",
        "questions": [
            "How much mortgage top-up lending is approved in New Zealand each month?",
        ],
        "article_title": "Using a Mortgage Top-Up to Renovate or Invest",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "interest-only": {
        "name": "Interest-only share of new investor vs owner-occupier mortgages",
        "also_known_as": "Investor vs owner-occupier interest-only mortgage share / Payment-type split",
        "category": "Mortgage lending & rates",
        "topic": "Loan structure & interest-only",
        "source": "RBNZ C32",
        "source_full": "Reserve Bank of New Zealand Table C32",
        "series_ids": "LVRN.MMB1.AN, LVRN.MMB1.AC",
        "source_url": RBNZ_SOURCES["c32"]["pages"][0],
        "data_type": "HLF-derived share (from RBNZ C32 payment-type series)",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The HLF-derived percentage of new residential investor lending versus new owner-occupier lending set up on interest-only terms (including revolving credit facilities) in RBNZ Table C32.",
        "methodology": "Derived from RBNZ Table C32 (New and existing residential mortgage lending by payment type).",
        "hlf_calculation": "(New Investor Interest-Only NZDm ÷ Total New Investor NZDm) × 100",
        "hlf_interpretation": "HLF market context: Shows the difference in interest-only usage between residential investment loans and owner-occupied home loans, subject to lender interest-only term limits.",
        "questions": [
            "What percentage of NZ property investors use interest-only mortgages?",
            "How common are interest-only home loans for owner-occupiers vs investors?",
        ],
        "article_title": "Interest-Only vs Principal & Interest Mortgages in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "dti-headroom": {
        "name": "First-home buyer high-DTI (>6x) lending share & speed-limit comparison",
        "also_known_as": "RBNZ Debt-to-Income (>6x TDTI) first-home buyer share",
        "category": "First-home buyers",
        "topic": "Debt-to-Income (DTI) & borrowing power",
        "source": "RBNZ C40",
        "source_full": "Reserve Bank of New Zealand Table C40",
        "series_ids": "DTIN.MPV (First-home buyer TDTI bands)",
        "source_url": RBNZ_SOURCES["c40"]["pages"][0],
        "data_type": "HLF-derived share & HLF-calculated speed-limit comparison",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The HLF-derived percentage of total new first-home buyer mortgage lending committed at a Total Debt-to-Income (TDTI) ratio above 6x gross income in RBNZ Table C40, compared against the 20% RBNZ speed-limit threshold.",
        "methodology": "Derived from RBNZ Table C40 (Residential mortgage lending by Debt-to-Income purpose use).",
        "hlf_calculation": "(FHB Lending with TDTI > 6 ÷ Total FHB Lending) × 100; Benchmark difference = 20.0% minus High-DTI Share",
        "hlf_interpretation": "HLF market context: Note that RBNZ Table C40 totals include DTI-exempt new-build lending, whereas the Reserve Bank's 20% speed limit applies to non-exempt FHB lending per bank. In practice, borrowing capacity can be constrained by bank servicing test rates, LVR rules, or individual lender policy before the 6x DTI threshold is reached.",
        "questions": [
            "What percentage of NZ first-home buyer lending has a Debt-to-Income ratio above 6x?",
            "How does high-DTI first-home lending compare with RBNZ's 20% speed-limit threshold?",
        ],
        "article_title": "Being Able to Afford a House vs Feeling Safe Buying One",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/the-difference-between-being-able-to-afford-a-house-and-feeling-safe-buying-one",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "borrower-income": {
        "name": "Average gross borrower income on approved FHB mortgage applications",
        "also_known_as": "RBNZ C41 Borrower Gross Income (BGI) — First-home buyer weighted mean",
        "category": "First-home buyers",
        "topic": "Borrower income & servicing",
        "source": "RBNZ C41",
        "source_full": "Reserve Bank of New Zealand Table C41",
        "series_ids": "DTINI.MPA (Borrower Gross Income by TDTI band)",
        "source_url": RBNZ_SOURCES["c41"]["pages"][0],
        "data_type": "HLF-derived borrower-weighted harmonic average (from RBNZ C41 income bands)",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The HLF-derived borrower-weighted average annual pre-tax gross income across approved New Zealand first-home buyer mortgage applications (combining all single and joint applicants on the loan).",
        "methodology": "Computed by HLF from RBNZ Table C41 (Residential mortgage borrower gross income) across all Total Debt-to-Income bands.",
        "hlf_calculation": "Total Gross Income NZDm across all TDTI bands ÷ Implied Borrower Count (Band NZDm ÷ Band Avg Income NZD000s)",
        "hlf_interpretation": "HLF market context: RBNZ Table C41 publishes band-level figures rather than a single nationwide mean; the overall figure is calculated by HLF. Because C41 measures total gross income per mortgage application (often combining two earners plus eligible boarder income), it represents combined application income rather than an individual salary.",
        "questions": [
            "What is the average combined household income on an approved first-home mortgage in NZ?",
        ],
        "article_title": "How Banks Assess Your Income for a New Zealand Mortgage",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-income-akl-vs-nz": {
        "name": "Approved FHB gross borrower income by Debt-to-Income band — New Zealand",
        "also_known_as": "Approved first-home buyer gross application income by TDTI band (RBNZ C41)",
        "category": "First-home buyers",
        "topic": "First-home lending & borrower income",
        "source": "RBNZ C41",
        "source_full": "Reserve Bank of New Zealand Table C41",
        "series_ids": "DTINI.MPA (Nationwide First-Home Buyer TDTI income bands)",
        "source_url": RBNZ_SOURCES["c41"]["pages"][0],
        "data_type": "Officially reported TDTI band averages & HLF-derived borrower-weighted mean",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "Official RBNZ Table C41 average pre-tax gross borrower income across each of the five first-home buyer Total Debt-to-Income (TDTI) bands, alongside the HLF-derived borrower-weighted mean across all bands.",
        "methodology": "Sourced from RBNZ Table C41 (Residential mortgage borrower gross income) across the five first-home buyer TDTI bands.",
        "hlf_calculation": "Harmonic borrower-weighted mean across all 5 FHB TDTI bands (Sum of Band Gross Income NZDm ÷ Sum of Implied Band Borrowers)",
        "hlf_interpretation": "HLF market context: RBNZ Table C41 reports each TDTI band separately; the overall weighted mean is derived by HLF. Because mortgage applications frequently combine two incomes, application averages sit higher than single-earner salaries.",
        "questions": [
            "How does first-home buyer gross income vary across Debt-to-Income (TDTI) bands in RBNZ C41?",
            "What is the average gross borrower income on an approved first-home mortgage application in New Zealand?",
        ],
        "article_title": "Buying a First Home on a Single or Modest Income",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "settled-rates": {
        "name": "Weighted-average settled 1-year and 2-year fixed mortgage interest rates",
        "also_known_as": "Actual settled NZ fixed mortgage rates (RBNZ B30)",
        "category": "Mortgage lending & rates",
        "topic": "Mortgage interest rates",
        "source": "RBNZ B30",
        "source_full": "Reserve Bank of New Zealand Table B30",
        "series_ids": "NCFA.MAM1PR.D1, NCFA.MAM1PR.D2",
        "source_url": RBNZ_SOURCES["b30"]["pages"][0],
        "data_type": "Officially reported RBNZ weighted-average settled rates",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The dollar-weighted average fixed mortgage interest rates across all new residential mortgages settled by New Zealand registered banks for 1-year and 2-year fixed terms.",
        "methodology": "Directly reported by RBNZ in Table B30 (New residential mortgage weighted-average interest rates) for owner-occupiers/borrowers.",
        "hlf_calculation": "Officially reported RBNZ series NCFA.MAM1PR.D1 (1-year) and NCFA.MAM1PR.D2 (2-year)",
        "hlf_interpretation": "HLF market context: Table B30 measures the weighted-average rate on mortgages settled during the month. Differences from current advertised rates can reflect the timing lag between rate lock and settlement, borrower LVR mix, and lender pricing.",
        "questions": [
            "What is the actual average 1-year and 2-year mortgage rate settled in New Zealand?",
            "How do settled mortgage rates compare with advertised bank special rates?",
        ],
        "article_title": "Why Fixed Mortgage Rates Don't Simply Follow the OCR",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/why-fixed-mortgage-rates-dont-follow-the-ocr-nz",
        "tool_title": "Compare NZ Mortgage Rates & Historical Tracker",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
    },
    "new-builds-annual": {
        "name": "Annual new residential dwellings consented in New Zealand",
        "also_known_as": "Stats NZ annual residential building consents",
        "category": "Housing supply & construction",
        "topic": "New builds & housing supply",
        "source": "Stats NZ",
        "source_full": "Stats NZ — Building Consents Issued",
        "series_ids": "Stats NZ Building Consents Monthly Release",
        "source_url": "https://www.stats.govt.nz/",
        "data_type": "Officially reported Stats NZ figures",
        "derived_by_hlf": False,
        "unit": "dwellings",
        "frequency": "Monthly",
        "definition": "The total number of new residential dwellings consented across New Zealand over the trailing 12-month period and in the latest single month.",
        "methodology": "Compiled monthly by Stats NZ from building consent data supplied by New Zealand's 67 territorial authorities.",
        "hlf_calculation": "Officially reported Stats NZ annual and monthly dwelling consent totals",
        "hlf_interpretation": "HLF market context: Building consents measure local-authority approvals to build and do not measure completed dwellings or code-compliance certificates.",
        "questions": [
            "How many new homes are consented in New Zealand each year?",
            "Are New Zealand building consents increasing or decreasing?",
        ],
        "article_title": "Guide to New-Build & Turnkey Construction Loans",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "townhouse-share": {
        "name": "Multi-unit share of new residential building consents",
        "also_known_as": "Non-standalone dwelling consent share & townhouse/flat/unit share (Stats NZ)",
        "category": "Housing supply & construction",
        "topic": "New builds & housing supply",
        "source": "Stats NZ",
        "source_full": "Stats NZ — Building Consents Issued",
        "series_ids": "Stats NZ Dwelling Type Breakdown",
        "source_url": "https://www.stats.govt.nz/",
        "data_type": "Officially reported counts & HLF-derived multi-unit and townhouse/flat/unit shares",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The HLF-derived percentage of new residential building consents in New Zealand issued for multi-unit/non-standalone dwellings overall (townhouses, flats, units, apartments, and retirement village units), alongside the specific share for townhouses, flats, and units.",
        "methodology": "Derived from Stats NZ monthly building consent counts by dwelling type.",
        "hlf_calculation": "Multi-unit overall = (Total Consents − Stand-Alone Houses) ÷ Total Consents; Townhouses/flats/units specifically = Townhouses/Flats/Units ÷ Total Consents",
        "hlf_interpretation": "HLF market context: The share of multi-unit consents reflects the composition of consented new housing supply across stand-alone houses, townhouses/flats/units, apartments, and retirement village units.",
        "questions": [
            "What percentage of new homes consented in New Zealand are multi-unit dwellings?",
            "What percentage of new dwelling consents are specifically townhouses, flats, and units?",
        ],
        "article_title": "Turnkey vs Construction Loans for Townhouses & New Builds",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Turnkey vs. Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/turnkey-vs-construction-loan-calculator-nz",
    },
    "build-cost-inflation": {
        "name": "Annual residential construction cost inflation (CPI purchase of newly built home)",
        "also_known_as": "New-build housing cost inflation (Stats NZ CPI)",
        "category": "Housing supply & construction",
        "topic": "New builds & construction costs",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (CPI)",
        "series_ids": "Stats NZ CPI Housing & Household Utilities",
        "source_url": "https://www.stats.govt.nz/",
        "data_type": "Officially reported Stats NZ quarterly CPI inflation",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in the cost of purchasing a newly built home in New Zealand as measured by the Stats NZ Consumers Price Index (CPI).",
        "methodology": "Published quarterly by Stats NZ within the Consumers Price Index release.",
        "hlf_calculation": "Officially reported Stats NZ annual percentage change for purchase of new housing",
        "hlf_interpretation": "HLF market context: Tracks annual changes in residential construction pricing as measured within the CPI basket (excluding land prices).",
        "questions": [
            "How fast are house building costs rising in New Zealand?",
        ],
        "article_title": "Managing Fixed-Price vs Cost-Escalation Build Contracts",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "rent-yield-pulse": {
        "name": "New Zealand national and regional median weekly private rent",
        "also_known_as": "MBIE Tenancy Services median weekly rent",
        "category": "Property investors & rents",
        "topic": "Rents & property investment yields",
        "source": "MBIE Tenancy Services",
        "source_full": "MBIE Tenancy Services — Rental Bond Data",
        "series_ids": "Tenancy Services Monthly Regional Bond CSV (Median Rent)",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported median rents & HLF-calculated YoY comparison",
        "derived_by_hlf": True,
        "unit": "NZD/week",
        "frequency": "Monthly",
        "definition": "The median weekly rent recorded on newly lodged private residential tenancy bonds across New Zealand and in Auckland, Wellington, and Canterbury.",
        "methodology": "Compiled by the Ministry of Business, Innovation and Employment (MBIE) from private rental bonds lodged with Tenancy Services.",
        "hlf_calculation": "Officially reported national ('ALL') and regional median weekly rents, plus HLF-calculated YoY % change",
        "hlf_interpretation": "HLF market context: Measures median rents on newly lodged bonds during the month rather than existing tenancies that did not turn over.",
        "questions": [
            "What is the median weekly rent in New Zealand, Auckland, Wellington, and Canterbury?",
        ],
        "article_title": "Rental Yield & Cashflow Guide for NZ Investors",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "rental-bonds-supply": {
        "name": "Monthly newly lodged and active private rental bonds in New Zealand",
        "also_known_as": "Active NZ private rental tenancies & new bond lodgements",
        "category": "Property investors & rents",
        "topic": "Rental supply & tenancy activity",
        "source": "MBIE Tenancy Services",
        "source_full": "MBIE Tenancy Services — Rental Bond Data",
        "series_ids": "Tenancy Services Monthly Regional Bond CSV (Lodged & Active Bonds)",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported MBIE Tenancy Services bond counts",
        "derived_by_hlf": False,
        "unit": "bonds",
        "frequency": "Monthly",
        "definition": "The number of newly lodged private rental tenancy bonds in the latest month alongside the total pool of active private rental bonds held across New Zealand.",
        "methodology": "Published monthly by MBIE Tenancy Services from statutory residential tenancy bond lodgements.",
        "hlf_calculation": "Officially reported 'ALL' location Lodged Bonds and Active Bonds counts",
        "hlf_interpretation": "HLF market context: Tracks monthly tenancy bond lodgements and total active private bonds held by Tenancy Services.",
        "questions": [
            "How many private rental tenancies are active in New Zealand?",
        ],
        "article_title": "NZ Rental Market Supply & Investor Insights",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "student-loan-servicing": {
        "name": "New Zealand median student loan balance",
        "also_known_as": "IRD Student Loan Scheme median borrower balance",
        "category": "Mortgage lending & rates",
        "topic": "Borrowing power & student loans",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — Student Loan Scheme Statistics",
        "series_ids": "IRD Student Loan Balances Release",
        "source_url": IRD_SL_BAL_PAGE,
        "data_type": "Officially reported Inland Revenue median & HLF-calculated YoY change",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Quarterly",
        "definition": "The median student loan balance across New Zealand student loan borrowers reported by Inland Revenue (Te Tari Taake).",
        "methodology": "Published by Inland Revenue from administrative Student Loan Scheme records.",
        "hlf_calculation": "Officially reported IRD median balance and YoY % change",
        "hlf_interpretation": "HLF market context (not measured in IRD release): In bank mortgage servicing assessments, compulsory PAYE student loan deductions reduce net take-home income and therefore affect borrowing capacity.",
        "questions": [
            "What is the median student loan balance in New Zealand?",
            "How does a student loan affect mortgage borrowing power in NZ?",
        ],
        "article_title": "How Student Loans Impact Your Mortgage Borrowing Power",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "kiwisaver-first-home": {
        "name": "Monthly KiwiSaver first-home withdrawal count, value & average per member",
        "also_known_as": "Monthly IRD KiwiSaver first-home purchase withdrawals",
        "category": "KiwiSaver",
        "topic": "KiwiSaver & first-home deposits",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — KiwiSaver Withdrawal Statistics",
        "series_ids": "IRD Monthly KiwiSaver First-Home Withdrawals",
        "source_url": IRD_KS_DATA_PAGE,
        "data_type": "Officially reported IRD totals & HLF-derived per-member average",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The officially reported monthly number of KiwiSaver members withdrawing funds for a first-home purchase and total dollars withdrawn, plus the HLF-derived mean withdrawal per member.",
        "methodology": "Compiled monthly by Inland Revenue from KiwiSaver scheme provider reports.",
        "hlf_calculation": "Mean per Member = Total First-Home Withdrawal Dollars ÷ Withdrawing Members",
        "hlf_interpretation": "HLF market context: IRD reports individual member withdrawals rather than household applications; when two KiwiSaver members purchase together, both withdrawals are counted separately in the member total.",
        "questions": [
            "How much is the average KiwiSaver first-home withdrawal per member in New Zealand?",
        ],
        "article_title": "Using Your KiwiSaver for a First-Home Deposit",
        "article_url": "https://www.homeloanfactory.co.nz/kiwisaver",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "fhb-regional-kiwisaver": {
        "name": "Regional KiwiSaver first-home withdrawal averages per member across New Zealand",
        "also_known_as": "KiwiSaver first-home withdrawal average by NZ region (16 regions)",
        "category": "First-home buyers",
        "topic": "KiwiSaver & regional first-home deposits",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — KiwiSaver Fund Withdrawals by Region",
        "series_ids": "IRD Annual Regional KiwiSaver First-Home Purchase Withdrawals",
        "source_url": IRD_KS_REG_PAGE,
        "data_type": "Officially reported IRD regional totals & HLF-derived per-member averages",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual",
        "definition": "HLF-derived mean KiwiSaver first-home withdrawal amount per withdrawing member across New Zealand's 16 regions, calculated from Inland Revenue's annual regional withdrawal totals and member counts.",
        "methodology": "Published annually by Inland Revenue for the financial year ended 30 June based on member postcodes.",
        "hlf_calculation": "Regional Mean per Member = Regional First-Home Withdrawal Dollars ÷ Regional Withdrawing Members",
        "hlf_interpretation": "HLF market context: IRD reports withdrawals per individual member (not per couple). For illustrative household planning, two withdrawing members at the national mean of ~$45,005/member would total ~$90,010 combined.",
        "questions": [
            "How much KiwiSaver do first-home buyers withdraw per member in Wellington, Auckland, and Canterbury?",
        ],
        "article_title": "How Kiwis Piece Together a First-Home Deposit by Region",
        "article_url": "https://www.homeloanfactory.co.nz/kiwisaver",
        "tool_title": "HLF KiwiSaver & First-Home Advice",
        "tool_url": "https://www.homeloanfactory.co.nz/kiwisaver",
    },
    "kiwisaver-switches": {
        "name": "Annual and monthly KiwiSaver scheme transfers (members switching providers)",
        "also_known_as": "KiwiSaver provider switching / KiwiSaver scheme transfers / KiwiSaver member transfers",
        "category": "KiwiSaver",
        "topic": "KiwiSaver scheme transfers",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — KiwiSaver Scheme Transfer Statistics",
        "series_ids": "IRD Annual KiwiSaver Scheme Transfers",
        "source_url": IRD_KS_TR_PAGE,
        "data_type": "Officially reported IRD annual total & HLF-derived monthly average",
        "derived_by_hlf": True,
        "unit": "transfers",
        "frequency": "Annual",
        "definition": "The officially reported annual number of New Zealand KiwiSaver members who transferred their balance to a different KiwiSaver scheme provider over the June financial year, plus the HLF-derived monthly mean.",
        "methodology": "Compiled by Inland Revenue from KiwiSaver scheme transfers processed during the year ended 30 June.",
        "hlf_calculation": "Annual Scheme Transfers ÷ 12 months",
        "hlf_interpretation": "HLF market context: Measures inter-scheme transfers where members move their KiwiSaver balance to a new provider (internal fund-type changes within the same scheme are not counted in IRD's transfer total).",
        "questions": [
            "How many New Zealanders switch KiwiSaver providers each year?",
            "How many KiwiSaver scheme transfers happen per month in NZ?",
        ],
        "article_title": "How to Review Your KiwiSaver Fund & Provider",
        "article_url": "https://www.homeloanfactory.co.nz/kiwisaver",
        "tool_title": "Book a Free HLF KiwiSaver Review",
        "tool_url": "https://www.homeloanfactory.co.nz/kiwisaver",
    },
}

# Silent legacy alias map ONLY for nz-mortgage-pulse.json so existing Webflow bubbles never break
LEGACY_BUBBLE_ALIASES = {
    "kiwisaver-active-switchers": "kiwisaver-switches",
    "investor-vs-owner-io": "interest-only",
}


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


def make_history_list(dates: pd.Series, values: pd.Series, fmt_fn, n: int = 13) -> list[dict]:
    out = []
    start = max(0, len(dates) - n)
    for idx in range(len(dates) - 1, start - 1, -1):
        dt = dates.iloc[idx]
        v = values.iloc[idx]
        if pd.notna(dt) and pd.notna(v):
            out.append({
                "period": dt.strftime("%Y-%m"),
                "label": dt.strftime("%b %Y"),
                "value": round(float(v), 2),
                "formatted": fmt_fn(float(v)),
            })
    return out


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
# Multi-Sheet RBNZ Workbook Scanner + Published Date Extractor
# ---------------------------------------------------------------------------
def norm_label(txt: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[-\u2013\u2014]+", " ", txt.lower())).strip()


def extract_rbnz_published_date(filepath: pathlib.Path) -> str:
    """Extract official Published Date from the RBNZ 'Table Description' sheet if present."""
    try:
        xls = pd.ExcelFile(filepath)
        for tab in xls.sheet_names:
            if "description" in tab.lower():
                df_desc = pd.read_excel(xls, sheet_name=tab, header=None)
                for r in range(min(12, df_desc.shape[0])):
                    k_str = str(df_desc.iloc[r, 0] or "").strip().lower()
                    if "published date" in k_str and df_desc.shape[1] > 1:
                        dt_p = pd.to_datetime(df_desc.iloc[r, 1], errors="coerce")
                        if pd.notna(dt_p):
                            return dt_p.strftime("%Y-%m-%d")
    except Exception:
        pass
    return ""


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
# Reference-Grade Parsers (Strict Official Data / HLF Calculation / Interpretation)
# ---------------------------------------------------------------------------
def build_rbnz_c31_tickers(filepath: pathlib.Path) -> tuple[dict, dict]:
    pub_date = extract_rbnz_published_date(filepath)
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

    _, oo_hi_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.BC2",), ("80", "lending", "other owner"), ("less", "equal"))
    _, oo_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AC2",), ("borrowers", "other owner"), ("80",))
    _, oo_hi_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.BC2",), ("80", "borrowers", "other owner"), ("less", "equal"))
    _, all_hi_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.BA",), ("80", "lending", "all borrower"), ("less", "equal"))
    _, all_hi_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.BA",), ("80", "borrowers", "all borrower"), ("less", "equal"))
    _, inv_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AN",), ("total", "borrowers", "investor"), ("70", "60", "80"))
    _, inv_gt80_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.BN",), ("80", "lending", "investor"), ("less", "equal"))
    _, inv_gt80_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.BN",), ("80", "borrowers", "investor"), ("less", "equal"))
    _, inv_gt70_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.DN",), ("70", "lending", "investor"), ("less", "equal"))
    _, inv_gt70_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.DN",), ("70", "borrowers", "investor"), ("less", "equal"))
    _, inv_gt60_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.FN",), ("60", "borrowers", "investor"), ("less", "equal"))
    _, inv_le60_cnt, _ = find_rbnz_series(tabs, ("LVRN.MMB2.TN",), ("60", "borrowers", "investor", "less"))

    i, iy = len(dates) - 1, max(0, len(dates) - 13)
    dt = dates.iloc[i]
    m_full = dt.strftime("%B %Y")
    as_of = month_end_iso(dt)
    hist_start_lbl = f"Available from {dates.iloc[0].strftime('%B %Y')} ({len(dates)} monthly observations)"

    ser_low_dep_pct = (fhb_hi_cnt / fhb_cnt) * 100.0
    ser_fhb_sh_pct = (fhb_val / tot_val) * 100.0
    ser_inv_sh_pct = (inv_val / tot_val) * 100.0
    ser_fhb_hi_avg_k = (fhb_hi_val * 1000.0) / fhb_hi_cnt.replace(0, float("nan"))
    ser_fhb_lo_avg_k = ((fhb_val - fhb_hi_val) * 1000.0) / (fhb_cnt - fhb_hi_cnt).replace(0, float("nan"))
    ser_fhb_prio_pct = (fhb_hi_val / all_hi_val) * 100.0
    ser_fhb_20club_pct = ((fhb_cnt - fhb_hi_cnt) / fhb_cnt) * 100.0
    ser_fhb_vs_inv_m = fhb_val - inv_val
    ser_oo_20eq_pct = ((oo_cnt - oo_hi_cnt) / oo_cnt) * 100.0
    ser_inv_u30_pct = (inv_gt70_cnt / inv_cnt) * 100.0
    ser_oo_low_eq_pct = (oo_hi_cnt / oo_cnt) * 100.0

    low_cnt_pct = float(ser_low_dep_pct.iloc[i])
    low_cnt_yoy = float(ser_low_dep_pct.iloc[iy])
    hi_dollars_m = float(fhb_hi_val.iloc[i])
    fhb_loans_int = int(round(float(fhb_cnt.iloc[i])))
    fhb_hi_loans_int = int(round(float(fhb_hi_cnt.iloc[i])))
    tot_val_m = float(tot_val.iloc[i])

    fhb_sh = float(ser_fhb_sh_pct.iloc[i])
    fhb_sh_yoy = float(ser_fhb_sh_pct.iloc[iy])
    fhb_avg_k = int(round((float(fhb_val.iloc[i]) * 1000.0) / float(fhb_cnt.iloc[i])))

    inv_sh = float(ser_inv_sh_pct.iloc[i])
    inv_sh_yoy = float(ser_inv_sh_pct.iloc[iy])
    inv_b = float(inv_val.iloc[i]) / 1000.0

    low_dep_html = (
        f"In {m_full}, <strong>{low_cnt_pct:.1f}%</strong> of first-home mortgages approved across New Zealand "
        f"(<strong>{fhb_hi_loans_int:,} of {fhb_loans_int:,} loans</strong>, or <strong>${hi_dollars_m:,.0f}m</strong> in lending) "
        f"were written at an LVR above 80% (corresponding to <strong>under a 20% deposit</strong>, derived by HLF from RBNZ Table C31), "
        f"{plain_prior_year_pct(low_cnt_pct, low_cnt_yoy)}."
    )

    fhb_share_html = (
        f"In {m_full}, registered banks approved <strong>{fhb_loans_int:,}</strong> first-home buyer mortgages totalling "
        f"<strong>{fmt_nzd_millions(float(fhb_val.iloc[i]))}</strong>—representing an HLF-derived <strong>{fhb_sh:.1f}%</strong> share of all new "
        f"residential mortgage lending ({plain_prior_year_pct(fhb_sh, fhb_sh_yoy)}) and a mean loan size of <strong>${fhb_avg_k:,}k</strong>."
    )

    inv_html = (
        f"In {m_full}, residential property investors accounted for an HLF-derived <strong>{inv_sh:.1f}%</strong> share of new mortgage commitments "
        f"(<strong>${inv_b:.2f}B</strong> reported in RBNZ Table C31, {plain_prior_year_pct(inv_sh, inv_sh_yoy)}), compared with "
        f"<strong>{fhb_sh:.1f}%</strong> for first-home buyers."
    )

    fhb_lo_val_cur = float(fhb_val.iloc[i] - fhb_hi_val.iloc[i])
    fhb_lo_cnt_cur = float(fhb_cnt.iloc[i] - fhb_hi_cnt.iloc[i])
    fhb_lo_loans_int = int(round(fhb_lo_cnt_cur))
    avg_hi_loan = (float(fhb_hi_val.iloc[i]) * 1_000_000.0) / max(1.0, float(fhb_hi_cnt.iloc[i]))
    avg_lo_loan = (fhb_lo_val_cur * 1_000_000.0) / max(1.0, fhb_lo_cnt_cur)
    loan_gap = avg_hi_loan - avg_lo_loan

    fhb_gap_html = (
        f"In {m_full}, the HLF-derived mean mortgage size for first-home buyers at &gt;80% LVR (under 20% deposit) was "
        f"<strong>${avg_hi_loan:,.0f}</strong> ({fhb_hi_loans_int:,} loans / ${hi_dollars_m:,.0f}m), compared with "
        f"<strong>${avg_lo_loan:,.0f}</strong> for first-home buyers at &le;80% LVR ({fhb_lo_loans_int:,} loans / ${fhb_lo_val_cur:,.0f}m)—an "
        f"HLF-calculated difference of <strong>${abs(loan_gap):,.0f}</strong>."
    )

    fhb_hi_val_share = float(ser_fhb_prio_pct.iloc[i])
    fhb_hi_cnt_share = (float(fhb_hi_cnt.iloc[i]) / max(1.0, float(all_hi_cnt.iloc[i]))) * 100.0
    all_hi_loans_int = int(round(float(all_hi_cnt.iloc[i])))
    fhb_priority_html = (
        f"In {m_full}, first-home buyers accounted for an HLF-derived <strong>{fhb_hi_val_share:.1f}% of all &gt;80% LVR mortgage dollars "
        f"({fmt_nzd_millions(float(fhb_hi_val.iloc[i]))} of {fmt_nzd_millions(float(all_hi_val.iloc[i]))})</strong> and "
        f"<strong>{fhb_hi_cnt_share:.1f}% of &gt;80% LVR loans ({fhb_hi_loans_int:,} of {all_hi_loans_int:,})</strong> "
        f"reported in RBNZ Table C31."
    )

    std_dep_pct = float(ser_fhb_20club_pct.iloc[i])
    fhb_20club_html = (
        f"In {m_full}, RBNZ Table C31 recorded <strong>{fhb_lo_loans_int:,} first-home buyers ({fmt_nzd_millions(fhb_lo_val_cur)})</strong> "
        f"approved at <strong>&le;80% LVR</strong> (corresponding to a 20% or greater deposit)—representing an HLF-derived "
        f"<strong>{std_dep_pct:.1f}%</strong> share of first-home approvals and a mean loan size of <strong>${avg_lo_loan:,.0f}</strong>."
    )

    fhb_v = float(fhb_val.iloc[i])
    inv_v = float(inv_val.iloc[i])
    diff_fi = fhb_v - inv_v
    if diff_fi >= 0:
        fhb_vs_inv_html = (
            f"In {m_full}, RBNZ Table C31 recorded <strong>{fmt_nzd_millions(fhb_v)} ({fhb_sh:.1f}% of lending)</strong> in commitments to first-home buyers "
            f"compared with <strong>{fmt_nzd_millions(inv_v)} ({inv_sh:.1f}%)</strong> to residential property investors—an HLF-calculated difference of "
            f"<strong>+{fmt_nzd_millions(abs(diff_fi))}</strong> toward first-home buyers."
        )
    else:
        fhb_vs_inv_html = (
            f"In {m_full}, RBNZ Table C31 recorded <strong>{fmt_nzd_millions(fhb_v)} ({fhb_sh:.1f}% of lending)</strong> in commitments to first-home buyers "
            f"compared with <strong>{fmt_nzd_millions(inv_v)} ({inv_sh:.1f}%)</strong> to residential property investors."
        )

    oo_tot_c = float(oo_cnt.iloc[i])
    oo_hi_c = float(oo_hi_cnt.iloc[i])
    oo_lo_c = oo_tot_c - oo_hi_c
    oo_lo_v = float(oo_val.iloc[i] - oo_hi_val.iloc[i])
    oo_20eq_pct = float(ser_oo_20eq_pct.iloc[i])
    oo_avg_20eq = (oo_lo_v * 1_000_000.0) / max(1.0, oo_lo_c)
    owner_eq_html = (
        f"In {m_full}, RBNZ Table C31 recorded <strong>{int(round(oo_lo_c)):,} of {int(round(oo_tot_c)):,}</strong> existing ('other') "
        f"owner-occupier mortgage approvals (<strong>{fmt_nzd_millions(oo_lo_v)}</strong>) at <strong>&le;80% LVR</strong> (corresponding to 20%+ "
        f"security equity)—an HLF-derived share of <strong>{oo_20eq_pct:.1f}%</strong> and mean loan size of <strong>${oo_avg_20eq:,.0f}</strong>."
    )

    inv_c_tot = int(round(float(inv_cnt.iloc[i])))
    inv_le60_c = int(round(float(inv_le60_cnt.iloc[i])))
    inv_30to40_c = int(round(float(inv_gt60_cnt.iloc[i] - inv_gt70_cnt.iloc[i])))
    inv_40plus_pct = (inv_le60_c / max(1, inv_c_tot)) * 100.0
    inv_30to40_pct = (inv_30to40_c / max(1, inv_c_tot)) * 100.0
    inv_u30_cnt = int(round(float(inv_gt70_cnt.iloc[i])))
    inv_u30_pct = float(ser_inv_u30_pct.iloc[i])
    inv_u30_val = float(inv_gt70_val.iloc[i])
    inv_eq_html = (
        f"In {m_full}, RBNZ Table C31 investor commitments comprised an HLF-derived <strong>{inv_40plus_pct:.1f}%</strong> at "
        f"<strong>&le;60% LVR</strong> ({inv_le60_c:,} loans), <strong>{inv_30to40_pct:.1f}%</strong> between <strong>60% and 70% LVR</strong> "
        f"({inv_30to40_c:,} loans), and <strong>{inv_u30_pct:.1f}%</strong> at <strong>&gt;70% LVR</strong> ({inv_u30_cnt:,} loans, or "
        f"{fmt_nzd_millions(inv_u30_val)})."
    )

    inv_gt80_c = int(round(float(inv_gt80_cnt.iloc[i])))
    inv_gt80_v = float(inv_gt80_val.iloc[i])
    inv_gt80_c_prev = int(round(float(inv_gt80_cnt.iloc[iy])))
    inv_gt80_v_prev = float(inv_gt80_val.iloc[iy])
    inv_gt80_pct = (inv_gt80_c / max(1, inv_c_tot)) * 100.0
    inv_70to80_c = max(0, inv_u30_cnt - inv_gt80_c)
    inv_70to80_v = max(0.0, inv_u30_val - inv_gt80_v)
    inv_70to80_pct = (inv_70to80_c / max(1, inv_c_tot)) * 100.0

    inv_12m_clause = ""
    if i >= 11:
        inv_gt80_12m_c = int(round(float(inv_gt80_cnt.iloc[i - 11 : i + 1].sum())))
        inv_gt80_12m_v = float(inv_gt80_val.iloc[i - 11 : i + 1].sum())
        inv_12m_clause = f" (and an HLF-derived 12-month total of <strong>{inv_gt80_12m_c:,} loans / {fmt_nzd_millions(inv_gt80_12m_v)}</strong>)"

    inv_dir_prev = "up from" if inv_gt80_c >= inv_gt80_c_prev else "compared with"
    inv_low_dep_html = (
        f"In {m_full}, RBNZ Table C31 recorded <strong>{inv_gt80_c:,} property investor mortgages ({fmt_nzd_millions(inv_gt80_v)})</strong> "
        f"at <strong>&gt;80% LVR</strong> (an HLF-derived <strong>{inv_gt80_pct:.1f}%</strong> share of investors, {inv_dir_prev} "
        f"{inv_gt80_c_prev:,} loans / {fmt_nzd_millions(inv_gt80_v_prev)} a year earlier){inv_12m_clause}, alongside an HLF-derived "
        f"<strong>{inv_70to80_c:,} investor loans ({fmt_nzd_millions(inv_70to80_v)}, or {inv_70to80_pct:.1f}%)</strong> in the "
        f"<strong>70%–80% LVR</strong> band."
    )

    inv_nb_lvr_html = (
        f"In {m_full}, RBNZ Table C31 recorded <strong>{inv_u30_cnt:,} property investor loans ({fmt_nzd_millions(inv_u30_val)})</strong> "
        f"at <strong>&gt;70% LVR</strong> (an HLF-derived <strong>{inv_u30_pct:.1f}%</strong> share of investor approvals)—comprising "
        f"an HLF-derived <strong>{inv_70to80_c:,} loans ({fmt_nzd_millions(inv_70to80_v)})</strong> in the <strong>70%–80% LVR band</strong> and "
        f"<strong>{inv_gt80_c:,} reported loans ({fmt_nzd_millions(inv_gt80_v)})</strong> at <strong>&gt;80% LVR</strong>."
    )

    oo_low_eq_pct = float(ser_oo_low_eq_pct.iloc[i])
    oo_low_eq_prev = float(ser_oo_low_eq_pct.iloc[iy])
    mover_low_eq_html = (
        f"In {m_full}, RBNZ Table C31 recorded <strong>{int(round(oo_hi_c)):,} mortgages ({fmt_nzd_millions(float(oo_hi_val.iloc[i]))})</strong> "
        f"for existing ('other') owner-occupiers at <strong>&gt;80% LVR</strong>—an HLF-derived <strong>{oo_low_eq_pct:.1f}%</strong> "
        f"share of existing owner-occupier approvals, {plain_prior_year_pct(oo_low_eq_pct, oo_low_eq_prev)}."
    )

    src_c31 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type" '
        'target="_blank" rel="noopener">RBNZ Table C31</a>, &copy; Reserve Bank of New Zealand. Percentages &amp; averages derived by HLF. '
        'Historical market statistics for general information only, not personalised financial advice.'
    )

    def make_c31_item(html_str: str, official_raw: str, val_num: float, val_fmt: str, prev_fmt: str, chg_str: str, calc_live: str, hist_list: list) -> dict:
        return {
            "chip": "RBNZ C31",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": html_str,
            "official_data_summary": official_raw,
            "source_html": src_c31,
            "latest_period": m_full,
            "value": round(val_num, 2),
            "latest_value": val_fmt,
            "previous_period": f"{dates.iloc[iy].strftime('%B %Y')} ({prev_fmt})",
            "change": chg_str,
            "calculation_live": calc_live,
            "historical_data_from": hist_start_lbl,
            "history": hist_list,
        }

    out = {
        "fhb-low-deposit": make_c31_item(
            low_dep_html,
            f"RBNZ Table C31 ({m_full}): {fhb_hi_loans_int:,} high-LVR (>80% LVR) FHB borrowers (${hi_dollars_m:,.0f}m) out of {fhb_loans_int:,} total FHB borrowers (${fhb_v:,.0f}m).",
            low_cnt_pct, f"{low_cnt_pct:.1f}%", f"{low_cnt_yoy:.1f}%",
            f"{low_cnt_pct - low_cnt_yoy:+.1f} percentage points YoY",
            f"{fhb_hi_loans_int:,} ÷ {fhb_loans_int:,} = {low_cnt_pct:.1f}% (Lending: ${hi_dollars_m:,.0f}m ÷ ${fhb_v:,.0f}m = {(hi_dollars_m/max(1.0,fhb_v))*100:.1f}%)",
            make_history_list(dates, ser_low_dep_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-share": make_c31_item(
            fhb_share_html,
            f"RBNZ Table C31 ({m_full}): {fhb_loans_int:,} FHB borrowers and ${fhb_v:,.0f}m FHB lending out of ${tot_val_m:,.0f}m total new residential lending.",
            fhb_sh, f"{fhb_sh:.1f}% ({fhb_loans_int:,} loans)", f"{fhb_sh_yoy:.1f}%",
            f"{fhb_sh - fhb_sh_yoy:+.1f} percentage points YoY",
            f"${fhb_v:,.0f}m ÷ ${tot_val_m:,.0f}m = {fhb_sh:.1f}% (Avg loan: ${fhb_v:,.0f}m ÷ {fhb_loans_int:,} = ${fhb_avg_k:,}k)",
            make_history_list(dates, ser_fhb_sh_pct, lambda x: f"{x:.1f}%"),
        ),
        "investors": make_c31_item(
            inv_html,
            f"RBNZ Table C31 ({m_full}): {inv_c_tot:,} investor borrowers and ${inv_v:,.0f}m investor lending out of ${tot_val_m:,.0f}m total lending.",
            inv_sh, f"{inv_sh:.1f}% (${inv_b:.2f}B)", f"{inv_sh_yoy:.1f}%",
            f"{inv_sh - inv_sh_yoy:+.1f} percentage points YoY",
            f"${inv_v:,.0f}m ÷ ${tot_val_m:,.0f}m = {inv_sh:.1f}%",
            make_history_list(dates, ser_inv_sh_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-loan-size-gap": make_c31_item(
            fhb_gap_html,
            f"RBNZ Table C31 ({m_full}): >80% LVR FHB = ${hi_dollars_m:,.0f}m across {fhb_hi_loans_int:,} borrowers; <=80% LVR FHB = ${fhb_lo_val_cur:,.0f}m across {fhb_lo_loans_int:,} borrowers.",
            loan_gap / 1000.0, f"${avg_hi_loan:,.0f} (>80% LVR) vs ${avg_lo_loan:,.0f} (<=80% LVR)",
            f"${float(ser_fhb_hi_avg_k.iloc[iy])*1000:,.0f} vs ${float(ser_fhb_lo_avg_k.iloc[iy])*1000:,.0f}",
            f"${abs(loan_gap):,.0f} difference between LVR tiers",
            f"(${hi_dollars_m:,.0f}m ÷ {fhb_hi_loans_int:,} = ${avg_hi_loan:,.0f}) minus (${fhb_lo_val_cur:,.0f}m ÷ {fhb_lo_loans_int:,} = ${avg_lo_loan:,.0f}) = ${loan_gap:,.0f}",
            make_history_list(dates, ser_fhb_hi_avg_k, lambda x: f"${x:,.1f}k (>80% LVR avg)"),
        ),
        "fhb-bank-lvr-priority": make_c31_item(
            fhb_priority_html,
            f"RBNZ Table C31 ({m_full}): FHB >80% LVR = ${hi_dollars_m:,.0f}m ({fhb_hi_loans_int:,} loans); All Borrower >80% LVR = ${float(all_hi_val.iloc[i]):,.0f}m ({all_hi_loans_int:,} loans).",
            fhb_hi_val_share, f"{fhb_hi_val_share:.1f}% of >80% LVR dollars ({fhb_hi_cnt_share:.1f}% of loans)",
            f"{float(ser_fhb_prio_pct.iloc[iy]):.1f}%",
            f"{fhb_hi_val_share - float(ser_fhb_prio_pct.iloc[iy]):+.1f} percentage points YoY",
            f"Dollars: ${hi_dollars_m:,.0f}m ÷ ${float(all_hi_val.iloc[i]):,.0f}m = {fhb_hi_val_share:.1f}% | Loans: {fhb_hi_loans_int:,} ÷ {all_hi_loans_int:,} = {fhb_hi_cnt_share:.1f}%",
            make_history_list(dates, ser_fhb_prio_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-20-deposit-club": make_c31_item(
            fhb_20club_html,
            f"RBNZ Table C31 ({m_full}): {fhb_lo_loans_int:,} <=80% LVR FHB borrowers (${fhb_lo_val_cur:,.0f}m) out of {fhb_loans_int:,} total FHB borrowers.",
            std_dep_pct, f"{std_dep_pct:.1f}% ({fhb_lo_loans_int:,} loans, avg ${avg_lo_loan:,.0f})",
            f"{float(ser_fhb_20club_pct.iloc[iy]):.1f}%",
            f"{std_dep_pct - float(ser_fhb_20club_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{fhb_lo_loans_int:,} ÷ {fhb_loans_int:,} = {std_dep_pct:.1f}% (Avg loan: ${fhb_lo_val_cur:,.0f}m ÷ {fhb_lo_loans_int:,} = ${avg_lo_loan:,.0f})",
            make_history_list(dates, ser_fhb_20club_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-vs-investors": make_c31_item(
            fhb_vs_inv_html,
            f"RBNZ Table C31 ({m_full}): FHB lending = ${fhb_v:,.0f}m ({fhb_loans_int:,} loans); Investor lending = ${inv_v:,.0f}m ({inv_c_tot:,} loans).",
            diff_fi, f"FHB {fmt_nzd_millions(fhb_v)} vs Investors {fmt_nzd_millions(inv_v)}",
            f"{fmt_nzd_millions(float(ser_fhb_vs_inv_m.iloc[iy]))} FHB minus Investor gap",
            f"{fmt_nzd_millions(diff_fi)} FHB lead in {m_full}",
            f"${fhb_v:,.0f}m (FHB) minus ${inv_v:,.0f}m (Investors) = {diff_fi:+,.0f}m",
            make_history_list(dates, ser_fhb_vs_inv_m, lambda x: f"{x:+.0f}m (FHB vs Inv)"),
        ),
        "owner-equity-buffer": make_c31_item(
            owner_eq_html,
            f"RBNZ Table C31 ({m_full}): {int(round(oo_lo_c)):,} <=80% LVR other owner-occupier borrowers (${oo_lo_v:,.0f}m) out of {int(round(oo_tot_c)):,} total other owner-occupier borrowers.",
            oo_20eq_pct, f"{oo_20eq_pct:.1f}% (avg loan ${oo_avg_20eq:,.0f})",
            f"{float(ser_oo_20eq_pct.iloc[iy]):.1f}%",
            f"{oo_20eq_pct - float(ser_oo_20eq_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{int(round(oo_lo_c)):,} ÷ {int(round(oo_tot_c)):,} = {oo_20eq_pct:.1f}% (Avg loan: ${oo_lo_v:,.0f}m ÷ {int(round(oo_lo_c)):,} = ${oo_avg_20eq:,.0f})",
            make_history_list(dates, ser_oo_20eq_pct, lambda x: f"{x:.1f}%"),
        ),
        "investor-equity-tiers": make_c31_item(
            inv_eq_html,
            f"RBNZ Table C31 ({m_full}): Total investors = {inv_c_tot:,}; <=60% LVR = {inv_le60_c:,}; >60% LVR = {int(round(float(inv_gt60_cnt.iloc[i]))):,}; >70% LVR = {inv_u30_cnt:,} (${inv_u30_val:,.0f}m).",
            inv_40plus_pct, f"{inv_40plus_pct:.1f}% (<=60% LVR) | {inv_30to40_pct:.1f}% (60–70% LVR) | {inv_u30_pct:.1f}% (>70% LVR)",
            f"{100.0 - float(ser_inv_u30_pct.iloc[iy]):.1f}% at <=70% LVR",
            f"{inv_u30_pct - float(ser_inv_u30_pct.iloc[iy]):+.1f} pts shift in >70% LVR share",
            f"<=60% LVR: {inv_le60_c:,} ÷ {inv_c_tot:,} = {inv_40plus_pct:.1f}% | 60–70% LVR: {inv_30to40_c:,} ÷ {inv_c_tot:,} = {inv_30to40_pct:.1f}% | >70% LVR: {inv_u30_cnt:,} ÷ {inv_c_tot:,} = {inv_u30_pct:.1f}%",
            make_history_list(dates, 100.0 - ser_inv_u30_pct, lambda x: f"{x:.1f}% (<=70% LVR)"),
        ),
        "investor-low-deposit": make_c31_item(
            inv_low_dep_html,
            f"RBNZ Table C31 ({m_full}): >80% LVR investors = {inv_gt80_c:,} loans (${inv_gt80_v:,.0f}m); >70% LVR investors = {inv_u30_cnt:,} loans (${inv_u30_val:,.0f}m).",
            float(inv_gt80_c), f"{inv_gt80_c:,} loans (${inv_gt80_v:,.0f}m) >80% LVR + {inv_70to80_c:,} loans (${inv_70to80_v:,.0f}m) 70–80% LVR",
            f"{inv_gt80_c_prev:,} loans (${inv_gt80_v_prev:,.0f}m) >80% LVR",
            f"{inv_gt80_c - inv_gt80_c_prev:+d} loans ({inv_gt80_v - inv_gt80_v_prev:+.0f}m) YoY in >80% LVR tier",
            f">80% LVR: {inv_gt80_c:,} ÷ {inv_c_tot:,} = {inv_gt80_pct:.1f}% (${inv_gt80_v:,.0f}m) | 70–80% LVR: {inv_u30_cnt:,} - {inv_gt80_c:,} = {inv_70to80_c:,} loans (${inv_70to80_v:,.0f}m)",
            make_history_list(dates, inv_gt80_cnt, lambda x: f"{int(round(x)):,} loans (>80% LVR)"),
        ),
        "investor-new-build-lvr": make_c31_item(
            inv_nb_lvr_html,
            f"RBNZ Table C31 ({m_full}): >70% LVR investors = {inv_u30_cnt:,} loans (${inv_u30_val:,.0f}m) out of {inv_c_tot:,} total investor loans.",
            inv_u30_pct, f"{inv_u30_pct:.1f}% ({inv_u30_cnt:,} loans / ${inv_u30_val:,.0f}m at >70% LVR)",
            f"{float(ser_inv_u30_pct.iloc[iy]):.1f}%",
            f"{inv_u30_pct - float(ser_inv_u30_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{inv_u30_cnt:,} ÷ {inv_c_tot:,} = {inv_u30_pct:.1f}% ({inv_70to80_c:,} loans in 70–80% LVR + {inv_gt80_c:,} loans in >80% LVR)",
            make_history_list(dates, ser_inv_u30_pct, lambda x: f"{x:.1f}% (>70% LVR)"),
        ),
        "mover-low-equity": make_c31_item(
            mover_low_eq_html,
            f"RBNZ Table C31 ({m_full}): {int(round(oo_hi_c)):,} >80% LVR other owner-occupier loans (${float(oo_hi_val.iloc[i]):,.0f}m) out of {int(round(oo_tot_c)):,} total other owner-occupier loans.",
            oo_low_eq_pct, f"{oo_low_eq_pct:.1f}% ({int(round(oo_hi_c)):,} loans / ${float(oo_hi_val.iloc[i]):,.0f}m)",
            f"{oo_low_eq_prev:.1f}%",
            f"{oo_low_eq_pct - oo_low_eq_prev:+.1f} percentage points YoY",
            f"{int(round(oo_hi_c)):,} ÷ {int(round(oo_tot_c)):,} = {oo_low_eq_pct:.1f}%",
            make_history_list(dates, ser_oo_low_eq_pct, lambda x: f"{x:.1f}%"),
        ),
    }

    if i >= 23:
        r12_ser = fhb_cnt.rolling(12).sum()
        r12_cnt = int(round(fhb_cnt.iloc[i - 11 : i + 1].sum()))
        r12_prev_cnt = int(round(fhb_cnt.iloc[i - 23 : i - 11].sum()))
        r12_val = float(fhb_val.iloc[i - 11 : i + 1].sum())
        r12_hi_cnt = int(round(fhb_hi_cnt.iloc[i - 11 : i + 1].sum()))
        r12_hi_pct = (r12_hi_cnt / max(1, r12_cnt)) * 100.0
        dir_r12 = "up from" if r12_cnt > r12_prev_cnt else "compared with"
        out["fhb-annual-milestone"] = make_c31_item(
            f"Over the 12 months to {m_full}, RBNZ Table C31 recorded an HLF-derived rolling total of <strong>{r12_cnt:,} first-home buyer mortgages</strong> "
            f"totalling <strong>{fmt_nzd_millions(r12_val)} in lending</strong> ({dir_r12} {r12_prev_cnt:,} loans a year earlier)—including "
            f"<strong>{r12_hi_cnt:,} borrowers ({r12_hi_pct:.1f}%)</strong> approved at <strong>&gt;80% LVR (under a 20% deposit)</strong>.",
            f"RBNZ Table C31 (12 months to {m_full}): Sum of monthly FHB approvals = {r12_cnt:,} loans (${r12_val:,.0f}m); >80% LVR FHB approvals = {r12_hi_cnt:,} loans.",
            float(r12_cnt), f"{r12_cnt:,} loans ({fmt_nzd_millions(r12_val)})", f"{r12_prev_cnt:,} loans",
            f"{r12_cnt - r12_prev_cnt:+,d} annual FHB loans YoY",
            f"12-month sum = {r12_cnt:,} loans (${r12_val:,.0f}m); >80% LVR share = {r12_hi_cnt:,} ÷ {r12_cnt:,} = {r12_hi_pct:.1f}%",
            make_history_list(dates, r12_ser, lambda x: f"{int(round(x)):,} loans (12m rolling)"),
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
    pub_date = extract_rbnz_published_date(filepath)
    tabs = load_rbnz_sheets(filepath)
    dates, top_v, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AZB",), ("lending", "top"))
    _, pur_v, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AZC",), ("lending", "purchase"))
    _, ref_v, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AZD",), ("lending", "provider"))
    _, oth_v, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AZE",), ("lending", "other"))
    _, top_c, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AZB",), ("borrowers", "top"))
    _, pur_c, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AZC",), ("borrowers", "purchase"))
    _, ref_c, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AZD",), ("borrowers", "provider"))
    _, oth_c, _ = find_rbnz_series(tabs, ("LVRN.MMB2.AZE",), ("borrowers", "other"))

    j, jy = len(dates) - 1, max(0, len(dates) - 13)
    dt = dates.iloc[j]
    m_full = dt.strftime("%B %Y")
    as_of = month_end_iso(dt)
    hist_start_lbl = f"Available from {dates.iloc[0].strftime('%B %Y')} ({len(dates)} monthly observations)"

    tot_v_ser = top_v + pur_v + ref_v + oth_v
    tot_c_ser = top_c + pur_c + ref_c + oth_c
    ref_sh_ser = (ref_v / tot_v_ser) * 100.0
    top_sh_c_ser = (top_c / tot_c_ser) * 100.0

    tot_v = float(tot_v_ser.iloc[j])
    tot_c = int(round(float(tot_c_ser.iloc[j])))
    ref_val_m = float(ref_v.iloc[j])
    ref_cnt_int = int(round(float(ref_c.iloc[j])))
    ref_sh = float(ref_sh_ser.iloc[j])
    ref_avg_k = int(round((ref_val_m * 1000.0) / max(1, ref_cnt_int)))
    top_sh_c = float(top_sh_c_ser.iloc[j])
    top_val_m = float(top_v.iloc[j])
    top_cnt_int = int(round(float(top_c.iloc[j])))
    top_avg_k = int(round((top_val_m * 1000.0) / max(1, top_cnt_int)))

    src_c33 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose" '
        'target="_blank" rel="noopener">RBNZ Table C33</a>, &copy; Reserve Bank of New Zealand. Derived by HLF. '
        'Historical market statistics for general information only, not personalised financial advice.'
    )
    return {
        "refinance": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, RBNZ Table C33 recorded <strong>{ref_cnt_int:,}</strong> borrowers changing loan provider and "
                f"<strong>${ref_val_m/1000.0:.2f}B</strong> in refinanced loans—representing an HLF-derived <strong>{ref_sh:.1f}%</strong> "
                f"share of all new residential lending and a mean loan size of <strong>${ref_avg_k:,}k</strong>."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Change in loan provider = {ref_cnt_int:,} borrowers and ${ref_val_m:,.0f}m out of ${tot_v:,.0f}m total purpose lending.",
            "source_html": src_c33,
            "latest_period": m_full,
            "value": round(ref_sh, 2),
            "latest_value": f"{ref_sh:.1f}% (${ref_val_m/1000.0:.2f}B / {ref_cnt_int:,} loans)",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({float(ref_sh_ser.iloc[jy]):.1f}%)",
            "change": f"{ref_sh - float(ref_sh_ser.iloc[jy]):+.1f} percentage points YoY",
            "calculation_live": f"${ref_val_m:,.0f}m ÷ ${tot_v:,.0f}m = {ref_sh:.1f}% (Avg loan: ${ref_val_m:,.0f}m ÷ {ref_cnt_int:,} = ${ref_avg_k:,}k)",
            "historical_data_from": hist_start_lbl,
            "history": make_history_list(dates, ref_sh_ser, lambda x: f"{x:.1f}%"),
        },
        "top-ups": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, RBNZ Table C33 recorded <strong>{top_cnt_int:,}</strong> borrowers approved for a mortgage top-up on an existing property—"
                f"an HLF-derived <strong>{top_sh_c:.1f}%</strong> share of all mortgage approvals by count and a mean top-up of <strong>${top_avg_k:,}k</strong>."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Top-up approvals = {top_cnt_int:,} borrowers (${top_val_m:,.0f}m) out of {tot_c:,} total approvals.",
            "source_html": src_c33,
            "latest_period": m_full,
            "value": round(top_sh_c, 2),
            "latest_value": f"{top_sh_c:.1f}% ({top_cnt_int:,} top-ups, avg ${top_avg_k:,}k)",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({float(top_sh_c_ser.iloc[jy]):.1f}%)",
            "change": f"{top_sh_c - float(top_sh_c_ser.iloc[jy]):+.1f} percentage points YoY",
            "calculation_live": f"{top_cnt_int:,} ÷ {tot_c:,} = {top_sh_c:.1f}% (Avg top-up: ${top_val_m:,.0f}m ÷ {top_cnt_int:,} = ${top_avg_k:,}k)",
            "historical_data_from": hist_start_lbl,
            "history": make_history_list(dates, top_sh_c_ser, lambda x: f"{x:.1f}%"),
        },
        "equity-topup-demand": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, RBNZ Table C33 recorded <strong>{fmt_nzd_millions(top_val_m)} of mortgage top-up lending</strong> across "
                f"<strong>{top_cnt_int:,} borrowers</strong> (an HLF-derived mean of <strong>${top_avg_k:,}k</strong> per top-up)."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Top-up lending value = ${top_val_m:,.0f}m across {top_cnt_int:,} borrowers.",
            "source_html": src_c33,
            "latest_period": m_full,
            "value": round(top_val_m, 1),
            "latest_value": f"{fmt_nzd_millions(top_val_m)} of top-up lending across {top_cnt_int:,} borrowers",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({fmt_nzd_millions(float(top_v.iloc[jy]))})",
            "change": f"{top_val_m - float(top_v.iloc[jy]):+.0f}m YoY",
            "calculation_live": f"Directly reported ${top_val_m:,.0f}m of top-up lending ÷ {top_cnt_int:,} borrowers = ${top_avg_k:,}k avg",
            "historical_data_from": hist_start_lbl,
            "history": make_history_list(dates, top_v, lambda x: f"${x:,.0f}m"),
        },
    }


def build_rbnz_c32_ticker(filepath: pathlib.Path, c31_by_ym: dict | None = None) -> dict:
    pub_date = extract_rbnz_published_date(filepath)
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
    iy = max(0, len(dates) - 13)
    inv_io_ser = (inv_io / inv_tot) * 100.0
    oo_io_ser = (oo_io / oo_tot) * 100.0
    inv_io_pct = float(inv_io_ser.iloc[-1])
    oo_io_pct = float(oo_io_ser.iloc[-1])
    hist_start_lbl = f"Available from {dates.iloc[0].strftime('%B %Y')} ({len(dates)} monthly observations)"

    m_full = dt.strftime("%B %Y")
    as_of = month_end_iso(dt)
    src_c32 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type" '
        'target="_blank" rel="noopener">RBNZ Table C32</a>, &copy; Reserve Bank of New Zealand. Derived by HLF. '
        'Historical market statistics for general information only, not personalised financial advice.'
    )
    return {
        "interest-only": {
            "chip": "RBNZ C32",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, an HLF-derived <strong>{inv_io_pct:.1f}%</strong> of new residential investor mortgage lending was set up on "
                f"<strong>interest-only</strong> terms (including revolving credit), compared with <strong>{oo_io_pct:.1f}%</strong> "
                f"for owner-occupiers in RBNZ Table C32."
            ),
            "official_data_summary": f"RBNZ Table C32 ({m_full}): New investor interest-only = ${float(inv_io.iloc[-1]):,.0f}m of ${float(inv_tot.iloc[-1]):,.0f}m; New owner-occupier interest-only = ${float(oo_io.iloc[-1]):,.0f}m of ${float(oo_tot.iloc[-1]):,.0f}m.",
            "source_html": src_c32,
            "latest_period": m_full,
            "value": round(inv_io_pct, 2),
            "latest_value": f"{inv_io_pct:.1f}% investor IO vs {oo_io_pct:.1f}% owner-occupier IO",
            "previous_period": f"{dates.iloc[iy].strftime('%B %Y')} ({float(inv_io_ser.iloc[iy]):.1f}% investor IO)",
            "change": f"{inv_io_pct - float(inv_io_ser.iloc[iy]):+.1f} percentage points YoY (investor IO)",
            "calculation_live": f"Investors: ${float(inv_io.iloc[-1]):,.0f}m ÷ ${float(inv_tot.iloc[-1]):,.0f}m = {inv_io_pct:.1f}% | Owner-Occupiers: ${float(oo_io.iloc[-1]):,.0f}m ÷ ${float(oo_tot.iloc[-1]):,.0f}m = {oo_io_pct:.1f}%",
            "historical_data_from": hist_start_lbl,
            "history": make_history_list(dates, inv_io_ser, lambda x: f"{x:.1f}% (investor IO)"),
        }
    }


def compute_c41_weighted_avg_income(tabs41: list, borrower_kw: str, region_mode: str = "nz") -> tuple[pd.Timestamp | None, float | None, pd.Series | None, pd.Series | None, list[dict]]:
    band_defs = [
        ("tdti <= 3", "TDTI ≤ 3x income"),
        ("3 < tdti <= 4", "3x < TDTI ≤ 4x income"),
        ("4 < tdti <= 5", "4x < TDTI ≤ 5x income"),
        ("5 < tdti <= 6", "5x < TDTI ≤ 6x income"),
        ("6 < tdti <= 7", "6x < TDTI ≤ 7x income"),
        ("tdti > 6", "TDTI > 6x income"),
        ("tdti > 7", "TDTI > 7x income"),
    ]

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

            for bk, _ in band_defs:
                if bk in lbl:
                    if "average gross income" in lbl:
                        avg_cols[bk] = idx
                    elif "gross income" in lbl:
                        inc_cols[bk] = idx

        common = [(bk, b_label) for bk, b_label in band_defs if bk in inc_cols and bk in avg_cols]
        if len(common) >= 4:
            tot_inc_ser = pd.Series(0.0, index=df41.index)
            tot_cnt_ser = pd.Series(0.0, index=df41.index)
            band_rows = []
            for bk, b_label in common:
                a_m = pd.to_numeric(df41.iloc[:, inc_cols[bk]], errors="coerce")
                b_k = pd.to_numeric(df41.iloc[:, avg_cols[bk]], errors="coerce")
                if len(b_k.dropna()) and b_k.dropna().median() > 2000:
                    b_k = b_k / 1000.0
                valid = (a_m > 0) & (b_k > 0)
                tot_inc_ser = tot_inc_ser.add(a_m.where(valid, 0.0), fill_value=0.0)
                tot_cnt_ser = tot_cnt_ser.add((a_m / b_k).where(valid, 0.0), fill_value=0.0)

                last_a = float(a_m.iloc[-1]) if pd.notna(a_m.iloc[-1]) else 0.0
                last_b = float(b_k.iloc[-1]) if pd.notna(b_k.iloc[-1]) else 0.0
                if last_a > 0 and last_b > 0:
                    band_rows.append({
                        "band": b_label,
                        "avg_income_nzd": int(round(last_b * 1000.0)),
                        "total_income_m": round(last_a, 1),
                    })

            weighted_k_ser = tot_inc_ser / tot_cnt_ser.replace(0.0, float("nan"))
            valid_w = weighted_k_ser.dropna()
            if len(valid_w):
                return df41["date"].iloc[-1], float(valid_w.iloc[-1]) * 1000.0, df41["date"], weighted_k_ser, band_rows
    return None, None, None, None, []


def build_rbnz_dti_income_tickers(c40_path: pathlib.Path | None, c41_path: pathlib.Path | None) -> dict:
    out = {}
    if c40_path is not None:
        try:
            pub_c40 = extract_rbnz_published_date(c40_path)
            tabs40 = load_rbnz_sheets(c40_path)
            val_pct, dt_dti, dti_dates, dti_ser, last_hi_v, last_tot_v = None, None, None, None, 0.0, 0.0

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
                    low_s = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in low_idxs)
                    unk_s = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in unk_idxs)
                    if gran_hi_idxs:
                        hi_s = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in gran_hi_idxs) + unk_s
                    else:
                        s_idx, is_excl_unk = sub_gt6_idxs[0]
                        hi_s = pd.to_numeric(df40.iloc[:, s_idx], errors="coerce").fillna(0.0) + (unk_s if is_excl_unk else 0.0)
                    tot_s = (low_s + hi_s).replace(0.0, float("nan"))
                    pct_s = (hi_s / tot_s) * 100.0
                    last_v = float(pct_s.iloc[-1])
                    if 0.1 <= last_v <= 45.0:
                        val_pct = last_v
                        dt_dti = df40["date"].iloc[-1]
                        dti_dates = df40["date"]
                        dti_ser = pct_s
                        last_hi_v = float(hi_s.iloc[-1])
                        last_tot_v = float(tot_s.iloc[-1])
                        break

            if val_pct is not None and dt_dti is not None:
                m_full = dt_dti.strftime("%B %Y")
                diff_cap = 20.0 - val_pct
                iy = max(0, len(dti_dates) - 13)
                prev_v = float(dti_ser.iloc[iy]) if pd.notna(dti_ser.iloc[iy]) else val_pct
                hist_start_c40 = f"Available from {dti_dates.iloc[0].strftime('%B %Y')} ({len(dti_dates)} monthly observations)"
                dti_html = (
                    f"In {m_full}, an HLF-derived <strong>{val_pct:.1f}%</strong> of total first-home buyer mortgage commitments in RBNZ Table C40 "
                    f"had a <strong>Total Debt-to-Income (TDTI) ratio above 6x</strong> (${last_hi_v:,.1f}m of ${last_tot_v:,.1f}m)—sitting "
                    f"<strong>{diff_cap:.1f} percentage points</strong> below the 20% RBNZ speed-limit threshold on non-exempt lending."
                )
                out["dti-headroom"] = {
                    "chip": "RBNZ C40",
                    "as_of": month_end_iso(dt_dti),
                    "source_published": pub_c40,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": dti_html,
                    "official_data_summary": f"RBNZ Table C40 ({m_full}): FHB lending with TDTI > 6 = ${last_hi_v:,.1f}m out of ${last_tot_v:,.1f}m total FHB commitments.",
                    "source_html": (
                        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use" '
                        'target="_blank" rel="noopener">RBNZ Table C40</a>, &copy; Reserve Bank of New Zealand. Derived by HLF. '
                        'Historical market statistics for general information only.'
                    ),
                    "latest_period": m_full,
                    "value": round(val_pct, 2),
                    "latest_value": f"{val_pct:.1f}% (>6x DTI, {diff_cap:.1f} pts below 20% benchmark)",
                    "previous_period": f"{dti_dates.iloc[iy].strftime('%B %Y')} ({prev_v:.1f}%)",
                    "change": f"{val_pct - prev_v:+.1f} percentage points YoY",
                    "calculation_live": f"${last_hi_v:,.1f}m ÷ ${last_tot_v:,.1f}m = {val_pct:.1f}% (Benchmark diff: 20.0% - {val_pct:.1f}% = {diff_cap:.1f} pts)",
                    "historical_data_from": hist_start_c40,
                    "history": make_history_list(dti_dates, dti_ser, lambda x: f"{x:.1f}%"),
                }
        except Exception as exc:
            print(f"[C40 DTI] Parse note: {exc}")

    if c41_path is not None:
        try:
            pub_c41 = extract_rbnz_published_date(c41_path)
            tabs41 = load_rbnz_sheets(c41_path)
            dt_fhb, fhb_inc, d41_dates, fhb_inc_ser, fhb_bands = compute_c41_weighted_avg_income(tabs41, "first home", "nz")
            _, inv_inc, _, _, _ = compute_c41_weighted_avg_income(tabs41, "investor", "nz")

            src_c41 = (
                'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi" '
                'target="_blank" rel="noopener">RBNZ Table C41</a>, &copy; Reserve Bank of New Zealand. Derived by HLF for general information only.'
            )

            if dt_fhb is not None and fhb_inc and 40000 <= fhb_inc <= 600000:
                m_c41 = dt_fhb.strftime("%B %Y")
                as_of_c41 = month_end_iso(dt_fhb)
                iy41 = max(0, len(d41_dates) - 13)
                prev_inc_k = float(fhb_inc_ser.iloc[iy41]) if pd.notna(fhb_inc_ser.iloc[iy41]) else (fhb_inc / 1000.0)
                hist_41 = make_history_list(d41_dates, fhb_inc_ser, lambda x: f"${x*1000:,.0f}")
                hist_start_c41 = f"Available from {d41_dates.iloc[0].strftime('%B %Y')} ({len(d41_dates)} monthly observations)"

                band_summary_parts = [f"{b['band']}: ${b['avg_income_nzd']:,}" for b in fhb_bands]
                band_summary_str = " | ".join(band_summary_parts) if band_summary_parts else f"Weighted mean: ${fhb_inc:,.0f}"
                lo_band = min((b["avg_income_nzd"] for b in fhb_bands), default=int(round(fhb_inc)))
                hi_band = max((b["avg_income_nzd"] for b in fhb_bands), default=int(round(fhb_inc)))

                if inv_inc and 40000 <= inv_inc <= 900000:
                    bgi_html = (
                        f"Based on HLF's borrower-weighted calculation across RBNZ Table C41's Debt-to-Income bands, approved New Zealand "
                        f"first-home buyer mortgage applications averaged <strong>${fhb_inc:,.0f}</strong> in gross borrower income in {m_c41}, "
                        f"compared with <strong>${inv_inc:,.0f}</strong> across approved property investor applications."
                    )
                else:
                    bgi_html = (
                        f"Based on HLF's borrower-weighted calculation across RBNZ Table C41's five Debt-to-Income bands, approved New Zealand "
                        f"first-home buyer mortgage applications averaged <strong>${fhb_inc:,.0f}</strong> in gross borrower income in {m_c41} "
                        f"(with official RBNZ band averages ranging from <strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong>)."
                    )
                out["borrower-income"] = {
                    "chip": "RBNZ C41",
                    "as_of": as_of_c41,
                    "source_published": pub_c41,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": bgi_html,
                    "official_data_summary": f"RBNZ Table C41 ({m_c41}) Official FHB Band Averages: {band_summary_str}.",
                    "source_html": src_c41,
                    "latest_period": m_c41,
                    "value": int(round(fhb_inc)),
                    "latest_value": f"${fhb_inc:,.0f} FHB weighted mean" + (f" vs ${inv_inc:,.0f} Investor" if inv_inc else ""),
                    "previous_period": f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_k*1000:,.0f} FHB)",
                    "change": f"${fhb_inc - (prev_inc_k*1000):+,.0f} YoY",
                    "calculation_live": f"HLF borrower-weighted harmonic mean across C41 TDTI bands = ${fhb_inc:,.0f} (FHB)" + (f" | ${inv_inc:,.0f} (Investor)" if inv_inc else ""),
                    "historical_data_from": hist_start_c41,
                    "history": hist_41,
                }

                out["fhb-income-akl-vs-nz"] = {
                    "chip": "RBNZ C41",
                    "as_of": as_of_c41,
                    "source_published": pub_c41,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"In {m_c41}, official RBNZ Table C41 average gross borrower incomes across New Zealand first-home buyer Debt-to-Income "
                        f"bands ranged from <strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong>, with an HLF-derived borrower-weighted "
                        f"mean of <strong>${fhb_inc:,.0f}</strong> across all approved first-home applications."
                    ),
                    "official_data_summary": f"RBNZ Table C41 ({m_c41}) Official FHB Band Averages: {band_summary_str}.",
                    "source_html": src_c41,
                    "latest_period": m_c41,
                    "value": int(round(fhb_inc)),
                    "tdti_bands": fhb_bands,
                    "latest_value": f"${fhb_inc:,.0f} HLF weighted mean (Official C41 bands: ${lo_band:,}–${hi_band:,})",
                    "previous_period": f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_k*1000:,.0f} HLF weighted mean)",
                    "change": f"${fhb_inc - (prev_inc_k*1000):+,.0f} ({((fhb_inc - prev_inc_k*1000)/(prev_inc_k*1000))*100:+.1f}%) YoY",
                    "calculation_live": f"HLF harmonic mean across {len(fhb_bands)} FHB TDTI bands = ${fhb_inc:,.0f} ({band_summary_str})",
                    "historical_data_from": hist_start_c41,
                    "history": hist_41,
                }
        except Exception as exc:
            print(f"[C41 BGI] Parse note: {exc}")
    return out


def build_rbnz_b30_ticker(filepath: pathlib.Path) -> dict:
    pub_date = extract_rbnz_published_date(filepath)
    tabs = load_rbnz_sheets(filepath)
    dates, r_1y, _ = find_rbnz_series(
        tabs, ("NCFA.MAM1PR.D1",), ("1 year",), ("value", "amount", "share", "investor"), val_bounds=(1.0, 15.0)
    )
    _, r_2y, _ = find_rbnz_series(
        tabs, ("NCFA.MAM1PR.D2",), ("2 year",), ("value", "amount", "share", "investor"), val_bounds=(1.0, 15.0)
    )
    dt = dates.iloc[-1]
    iy = max(0, len(dates) - 13)
    m_full = dt.strftime("%B %Y")
    v1, v2 = float(r_1y.iloc[-1]), float(r_2y.iloc[-1])
    v1_prev = float(r_1y.iloc[iy])
    hist_start_b30 = f"Available from {dates.iloc[0].strftime('%B %Y')} ({len(dates)} monthly observations)"

    return {
        "settled-rates": {
            "chip": "RBNZ B30",
            "as_of": month_end_iso(dt),
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, RBNZ Table B30 reported a dollar-weighted average settled fixed mortgage interest rate of "
                f"<strong>{v1:.2f}%</strong> for 1-year terms and <strong>{v2:.2f}%</strong> for 2-year terms across New Zealand owner-occupiers."
            ),
            "official_data_summary": f"RBNZ Table B30 ({m_full}): Weighted-average settled owner-occupier rate = {v1:.2f}% (1-year fixed) and {v2:.2f}% (2-year fixed).",
            "source_html": (
                'Source: <a href="https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-weighted-average-interest-rates" '
                'target="_blank" rel="noopener">RBNZ Table B30</a>, &copy; Reserve Bank of New Zealand. Historical market statistics for general information only.'
            ),
            "latest_period": m_full,
            "value": round(v1, 2),
            "latest_value": f"{v1:.2f}% (1-year settled) | {v2:.2f}% (2-year settled)",
            "previous_period": f"{dates.iloc[iy].strftime('%B %Y')} ({v1_prev:.2f}% 1-year settled)",
            "change": f"{v1 - v1_prev:+.2f}% YoY (1-year settled)",
            "calculation_live": f"Directly reported RBNZ B30 series NCFA.MAM1PR.D1 = {v1:.2f}% (1Y) & NCFA.MAM1PR.D2 = {v2:.2f}% (2Y)",
            "historical_data_from": hist_start_b30,
            "history": make_history_list(dates, r_1y, lambda x: f"{x:.2f}% (1Y settled)"),
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
            tfu_pct = (townhouses / mth_tot) * 100.0
            other_multi = max(0, multi_tot - townhouses)
            other_multi_pct = (other_multi / mth_tot) * 100.0
            m_full = dt_cand.strftime("%B %Y")
            as_of = month_end_iso(pd.Timestamp(dt_cand))

            nb_html = (
                f"In the 12 months ended {m_full}, Stats NZ reported <strong>{ann_tot:,}</strong> new residential dwellings consented across New Zealand "
                f"({dir_word} <strong>{ann_pct:.1f}%</strong> year-on-year, with <strong>{mth_tot:,}</strong> dwellings consented in the month of {m_full})."
                if dir_word in ("up", "down") and ann_pct > 0
                else f"In the 12 months ended {m_full}, Stats NZ reported <strong>{ann_tot:,}</strong> new residential dwellings consented across New Zealand, including <strong>{mth_tot:,}</strong> dwellings in {m_full}."
            )

            th_html = (
                f"In {m_full}, an HLF-derived <strong>{multi_pct:.1f}%</strong> (<strong>{multi_tot:,} of {mth_tot:,}</strong>) of new residential dwelling "
                f"consents were for multi-unit/non-standalone dwellings (including townhouses, flats, units, apartments, and retirement village units)—with "
                f"townhouses, flats, and units specifically accounting for <strong>{tfu_pct:.1f}% ({townhouses:,} of {mth_tot:,})</strong> and "
                f"apartments plus retirement village units accounting for <strong>{other_multi_pct:.1f}% ({other_multi:,})</strong>."
            )

            src_bc = (
                f'Source: <a href="{url}" target="_blank" rel="noopener">Stats NZ — Building Consents Issued ({m_full})</a>, '
                'licensed under CC BY 4.0. Multi-unit and townhouse shares derived by HLF. Historical market statistics only.'
            )
            out["new-builds-annual"] = {
                "chip": "Stats NZ",
                "as_of": as_of,
                "max_age_days": MAX_AGE_MONTHLY,
                "html": nb_html,
                "official_data_summary": f"Stats NZ Building Consents ({m_full}): Annual consents = {ann_tot:,} ({dir_word} {ann_pct:.1f}% YoY); Monthly consents = {mth_tot:,}.",
                "source_html": src_bc,
                "latest_period": f"Year ended {m_full}",
                "value": ann_tot,
                "latest_value": f"{ann_tot:,} annual consents ({mth_tot:,} in {m_full})",
                "previous_period": "Prior 12-month period",
                "change": f"{dir_word.capitalize()} {ann_pct:.1f}% YoY" if dir_word else "Reported by Stats NZ",
                "calculation_live": f"Directly reported Stats NZ annual total = {ann_tot:,} ({mth_tot:,} in {m_full})",
                "historical_data_from": f"Stats NZ monthly release ({m_full})",
            }
            out["townhouse-share"] = {
                "chip": "Stats NZ",
                "as_of": as_of,
                "max_age_days": MAX_AGE_MONTHLY,
                "html": th_html,
                "official_data_summary": f"Stats NZ Building Consents ({m_full}): Total dwellings = {mth_tot:,}; Stand-alone houses = {houses:,}; Townhouses, flats & units = {townhouses:,}; Apartments & retirement units = {other_multi:,}.",
                "source_html": src_bc,
                "latest_period": m_full,
                "value": round(multi_pct, 1),
                "townhouse_flat_unit_share_pct": round(tfu_pct, 1),
                "latest_value": f"Multi-unit overall: {multi_pct:.1f}% ({multi_tot:,}/{mth_tot:,}) | Townhouses/flats/units: {tfu_pct:.1f}% ({townhouses:,}/{mth_tot:,})",
                "previous_period": f"Stand-alone houses: {100.0 - multi_pct:.1f}% ({houses:,}/{mth_tot:,})",
                "change": f"Apartments & retirement village units: {other_multi_pct:.1f}% ({other_multi:,}/{mth_tot:,})",
                "calculation_live": f"Multi-unit overall: ({mth_tot:,} − {houses:,}) ÷ {mth_tot:,} = {multi_tot:,} ÷ {mth_tot:,} = {multi_pct:.1f}% | Townhouses/flats/units specifically: {townhouses:,} ÷ {mth_tot:,} = {tfu_pct:.1f}%",
                "historical_data_from": f"Stats NZ monthly release ({m_full})",
            }
            print(f"[Stats NZ Consents] Parsed {m_full}: {ann_tot:,} annual, {multi_pct:.1f}% multi-unit ({tfu_pct:.1f}% townhouses/flats/units)")
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
            cpi_clause = f", alongside <strong>{cpi_all_val:.1f}%</strong> overall annual CPI inflation" if cpi_all_val is not None else ""

            bc_inf_html = (
                f"In the {qname.capitalize()} {yr} quarter, Stats NZ reported a <strong>{'-' if is_down else '+'}{val:.1f}%</strong> "
                f"annual change in the CPI index for purchasing a newly built home{cpi_clause}."
            )

            out["build-cost-inflation"] = {
                "chip": "Stats NZ CPI",
                "as_of": month_end_iso(pd.Timestamp(q_dt)),
                "max_age_days": MAX_AGE_QUARTERLY,
                "html": bc_inf_html,
                "official_data_summary": f"Stats NZ CPI ({qname.capitalize()} {yr} quarter): Purchase of newly built home = {'-' if is_down else '+'}{val:.1f}% YoY.",
                "source_html": (
                    f'Source: <a href="{cpi_url}" target="_blank" rel="noopener">Stats NZ — Consumers Price Index ({qname.capitalize()} {yr} quarter)</a>, '
                    'licensed under CC BY 4.0. Historical market statistics only.'
                ),
                "latest_period": f"{qname.capitalize()} {yr} quarter",
                "value": signed_val,
                "latest_value": f"{'-' if is_down else '+'}{val:.1f}% YoY",
                "previous_period": f"Headline CPI: {cpi_all_val:.1f}% YoY" if cpi_all_val is not None else "Prior year quarter",
                "change": f"{'-' if is_down else '+'}{val:.1f}% annual change",
                "calculation_live": f"Directly reported Stats NZ CPI purchase of new housing = {'-' if is_down else '+'}{val:.1f}% YoY",
                "historical_data_from": f"Stats NZ quarterly CPI release ({qname.capitalize()} {yr})",
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

    nat = df[df["Location"].astype(str).str.strip().str.upper() == "ALL"].dropna(subset=["dt", "Median Rent"]).sort_values("dt").reset_index(drop=True)
    last = nat.iloc[-1]
    dt = last["dt"]
    m_full = dt.strftime("%B %Y")
    as_of = month_end_iso(dt)
    nat_r = int(round(float(last["Median Rent"])))
    hist_start_mbie = f"Available from {nat['dt'].iloc[0].strftime('%B %Y')} ({len(nat)} monthly observations)"

    yoy_rows = nat[nat["dt"] == (dt - pd.DateOffset(years=1))]
    prev_r = int(round(float(yoy_rows.iloc[-1]["Median Rent"]))) if len(yoy_rows) else nat_r
    yoy_pct = ((float(last["Median Rent"]) - prev_r) / prev_r * 100.0) if len(yoy_rows) and prev_r else 0.0
    yoy_phrase = f" ({plain_yoy_growth(float(last['Median Rent']), float(yoy_rows.iloc[-1]['Median Rent']))}, derived by HLF)" if len(yoy_rows) else ""

    def reg_med(substr: str) -> int:
        sub = df[(df["dt"] == dt) & (df["Location"].astype(str).str.contains(substr, case=False, na=False))]
        return int(round(float(sub.iloc[-1]["Median Rent"]))) if len(sub) else 0

    akl_r = reg_med("Auckland")
    cant_r = reg_med("Canterbury")
    wgtn_r = reg_med("Wellington")

    src_mbie = (
        'Source: <a href="https://www.tenancy.govt.nz/about-tenancy-services/data-and-statistics/rental-bond-data/" '
        'target="_blank" rel="noopener">Ministry of Business, Innovation and Employment (Tenancy Services Rental Bond Data)</a>, '
        'licensed under CC BY 3.0 NZ. YoY change derived by HLF. Historical market statistics only.'
    )

    lodged_int = int(round(float(last["Lodged Bonds"])))
    active_int = int(round(float(last["Active Bonds"])))

    return {
        "rent-yield-pulse": {
            "chip": "Tenancy Services",
            "as_of": as_of,
            "max_age_days": MAX_AGE_TENANCY,
            "html": (
                f"In {m_full}, MBIE Tenancy Services reported a national median weekly rent across newly lodged private tenancy bonds of "
                f"<strong>${nat_r}/wk</strong>{yoy_phrase}, with regional medians of <strong>${akl_r}/wk</strong> in Auckland, "
                f"<strong>${wgtn_r}/wk</strong> in Wellington, and <strong>${cant_r}/wk</strong> in Canterbury."
            ),
            "official_data_summary": f"MBIE Tenancy Services ({m_full}): National ('ALL') median weekly rent = ${nat_r}/wk; Auckland = ${akl_r}/wk; Wellington = ${wgtn_r}/wk; Canterbury = ${cant_r}/wk.",
            "source_html": src_mbie,
            "latest_period": m_full,
            "value": nat_r,
            "latest_value": f"${nat_r}/wk NZ (${akl_r} AKL | ${wgtn_r} WGTN | ${cant_r} CANT)",
            "previous_period": f"12 months earlier (${prev_r}/wk NZ)",
            "change": f"{yoy_pct:+.1f}% YoY (${nat_r - prev_r:+d}/wk)",
            "calculation_live": f"Reported national median = ${nat_r}/wk; YoY = (${nat_r} - ${prev_r}) ÷ ${prev_r} = {yoy_pct:+.1f}%",
            "historical_data_from": hist_start_mbie,
            "history": make_history_list(nat["dt"], nat["Median Rent"], lambda x: f"${int(round(x))}/wk"),
        },
        "rental-bonds-supply": {
            "chip": "Tenancy Services",
            "as_of": as_of,
            "max_age_days": MAX_AGE_TENANCY,
            "html": (
                f"In {m_full}, MBIE Tenancy Services recorded <strong>{lodged_int:,}</strong> newly lodged private rental bonds "
                f"and <strong>{active_int:,}</strong> active private rental bonds nationwide."
            ),
            "official_data_summary": f"MBIE Tenancy Services ({m_full}): Nationwide ('ALL') Lodged Bonds = {lodged_int:,}; Active Bonds = {active_int:,}.",
            "source_html": src_mbie,
            "latest_period": m_full,
            "value": active_int,
            "latest_value": f"{lodged_int:,} lodged bonds | {active_int:,} active bonds",
            "previous_period": "Monthly MBIE bond registry",
            "change": f"{lodged_int:,} new private tenancies lodged in {m_full}",
            "calculation_live": f"Directly reported MBIE 'ALL' location: {lodged_int:,} lodged bonds & {active_int:,} active bonds",
            "historical_data_from": hist_start_mbie,
            "history": make_history_list(nat["dt"], nat["Lodged Bonds"], lambda x: f"{int(round(x)):,} lodged bonds"),
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
            prev_lbl = m_med.group(2)
            curr_bal = int(m_med.group(3).replace(",", ""))
            dt_sl = pd.to_datetime("1 " + m_med.group(4), errors="coerce")
            if pd.notna(dt_sl):
                m_full = dt_sl.strftime("%B %Y")
                pct_sl = ((curr_bal - prev_bal) / max(1, prev_bal)) * 100.0
                out["student-loan-servicing"] = {
                    "chip": "Inland Revenue",
                    "as_of": month_end_iso(dt_sl),
                    "max_age_days": MAX_AGE_QUARTERLY,
                    "html": (
                        f"In Inland Revenue's {m_full} Student Loan Scheme statistics, the median New Zealand student loan borrower balance "
                        f"was <strong>${curr_bal:,}</strong> ({plain_yoy_growth(curr_bal, prev_bal)} from ${prev_bal:,} in {prev_lbl})."
                    ),
                    "official_data_summary": f"Inland Revenue Student Loan Scheme ({m_full}): Median borrower balance = ${curr_bal:,} (compared with ${prev_bal:,} in {prev_lbl}).",
                    "source_html": (
                        f'Source: <a href="{IRD_SL_BAL_PAGE}" target="_blank" rel="noopener">Inland Revenue — Student Loan Statistics ({dt_sl.strftime("%b %Y")})</a>, '
                        'licensed under CC BY 4.0. Historical market statistics for general information only.'
                    ),
                    "latest_period": m_full,
                    "value": curr_bal,
                    "latest_value": f"${curr_bal:,} median balance",
                    "previous_period": f"{prev_lbl} (${prev_bal:,})",
                    "change": f"${curr_bal - prev_bal:+,d} ({pct_sl:+.1f}%)",
                    "calculation_live": f"Directly reported IRD median = ${curr_bal:,} (Change: (${curr_bal:,} - ${prev_bal:,}) ÷ ${prev_bal:,} = {pct_sl:+.1f}%)",
                    "historical_data_from": f"Inland Revenue release ({m_full})",
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
                prev_amt_m = float(ma.group(3).replace(",", "")) if ma.group(3) else None
                avg_mem_nzd = int(round((amt_m * 1_000_000.0) / max(1, cnt)))
                m_full = dt_ks.strftime("%B %Y")

                out["kiwisaver-first-home"] = {
                    "chip": "Inland Revenue",
                    "as_of": month_end_iso(dt_ks),
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"In {m_full}, Inland Revenue recorded <strong>{cnt:,}</strong> KiwiSaver members withdrawing a combined "
                        f"<strong>${amt_m:,.1f}m</strong> for a first-home purchase—an HLF-derived mean of "
                        f"<strong>${avg_mem_nzd:,}</strong> per withdrawing member."
                    ),
                    "official_data_summary": f"Inland Revenue KiwiSaver Withdrawals ({m_full}): First-home purchase withdrawals = {cnt:,} members and ${amt_m:,.1f}m.",
                    "source_html": (
                        f'Source: <a href="{IRD_KS_DATA_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Withdrawal Statistics</a>, '
                        'licensed under CC BY 4.0. Mean per member derived by HLF for general information only.'
                    ),
                    "latest_period": m_full,
                    "value": round(amt_m, 1),
                    "latest_value": f"${amt_m:,.1f}m across {cnt:,} members (${avg_mem_nzd:,}/member)",
                    "previous_period": f"${prev_amt_m:,.1f}m in prior period" if prev_amt_m else "Prior monthly release",
                    "change": f"{amt_m - prev_amt_m:+.1f}m vs prior period" if prev_amt_m else f"${avg_mem_nzd:,} mean per withdrawing member",
                    "calculation_live": f"${amt_m:,.1f}m ($195,700,000) ÷ {cnt:,} members = ${avg_mem_nzd:,}/member" if amt_m == 195.7 else f"${amt_m:,.1f}m ÷ {cnt:,} members = ${avg_mem_nzd:,}/member",
                    "historical_data_from": f"Inland Revenue monthly release ({m_full})",
                }

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
                        reg_list.append((clean_disp, int(round(nzd_val / c_val)), c_val, round(nzd_val / 1e6, 2)))

                if tot_mem > 0 and len(reg_list) >= 10:
                    reg_list.sort(key=lambda x: x[1], reverse=True)
                    nz_avg_p = int(round(tot_nzd / tot_mem))
                    high_reg_name, high_reg_avg = reg_list[0][0], reg_list[0][1]
                    low_reg_name, low_reg_avg = reg_list[-1][0], reg_list[-1][1]
                    akl_avg = dict((r[0], r[1]) for r in reg_list).get("Auckland", reg_list[1][1])
                    reg_hist = [
                        {"period": r[0], "label": r[0], "value": r[1], "formatted": f"${r[1]:,}/member ({r[2]:,} members, ${r[3]:.2f}m)"}
                        for r in reg_list
                    ]
                    out["fhb-regional-kiwisaver"] = {
                        "chip": "Inland Revenue",
                        "as_of": f"{fy_yr_reg}-06-30",
                        "max_age_days": MAX_AGE_ANNUAL,
                        "html": (
                            f"Over the year ended 30 June {fy_yr_reg}, Inland Revenue recorded <strong>{tot_mem:,} KiwiSaver members</strong> "
                            f"withdrawing <strong>${tot_nzd/1_000_000_000.0:.2f}B</strong> for a first home across New Zealand's 16 regions—an "
                            f"HLF-derived mean of <strong>${nz_avg_p:,} per withdrawing member nationwide</strong>, ranging from "
                            f"<strong>${high_reg_avg:,}/member in {high_reg_name}</strong> (and <strong>${akl_avg:,}/member in Auckland</strong>) "
                            f"to <strong>${low_reg_avg:,}/member in {low_reg_name}</strong>."
                        ),
                        "official_data_summary": f"Inland Revenue Regional KiwiSaver Withdrawals (Year ended 30 June {fy_yr_reg}): {tot_mem:,} members withdrew ${tot_nzd:,.0f} across 16 NZ regions.",
                        "source_html": (
                            f'Source: <a href="{IRD_KS_REG_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Fund Withdrawals by Region (Year ended Jun {fy_yr_reg})</a>, '
                            'licensed under CC BY 4.0. Per-member averages derived by HLF for general information only.'
                        ),
                        "latest_period": f"Year ended June {fy_yr_reg}",
                        "value": nz_avg_p,
                        "latest_value": f"${nz_avg_p:,}/member NZ ({tot_mem:,} members / ${tot_nzd/1e9:.2f}B total)",
                        "previous_period": f"Regional range: ${low_reg_avg:,}–${high_reg_avg:,} per member",
                        "change": f"Highest: {high_reg_name} ${high_reg_avg:,} | Lowest: {low_reg_name} ${low_reg_avg:,}",
                        "calculation_live": f"${tot_nzd:,.0f} ÷ {tot_mem:,} members = ${nz_avg_p:,}/member (Regional range: ${low_reg_avg:,}–${high_reg_avg:,})",
                        "historical_data_from": f"16 NZ Regions (Year ended 30 June {fy_yr_reg})",
                        "history": reg_hist,
                    }
        except Exception as exc:
            print(f"[IRD Regional KiwiSaver] Parse note: {exc}")

    tr_bytes = fetch_url_bytes(session, IRD_KS_TR_PAGE, timeout=30)
    if tr_bytes:
        tr_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", tr_bytes.decode("utf-8", errors="ignore")))
        mt = re.search(
            r"year ended\s+30\s+June\s+(\d{4}).*?scheme transfers\s+(?:increased|decreased),?\s+from\s+([\d,]+)\s+in the\s+(\d{4})\s+financial year to\s+([\d,]+)\s+in the\s+(\d{4})\s+financial year",
            tr_txt,
            flags=re.I,
        )
        if mt:
            prev_tr = int(mt.group(2).replace(",", ""))
            prev_fy = int(mt.group(3))
            tr_tot = int(mt.group(4).replace(",", ""))
            fy_yr = int(mt.group(5))
            mth_avg = int(round(tr_tot / 12.0))
            out["kiwisaver-switches"] = {
                "chip": "Inland Revenue",
                "as_of": f"{fy_yr}-06-30",
                "max_age_days": MAX_AGE_ANNUAL,
                "html": (
                    f"Over the year ended 30 June {fy_yr}, Inland Revenue recorded <strong>{tr_tot:,}</strong> KiwiSaver inter-scheme "
                    f"transfers (compared with {prev_tr:,} in FY{prev_fy})—an HLF-derived mean of <strong>{mth_avg:,}</strong> transfers per month."
                ),
                "official_data_summary": f"Inland Revenue KiwiSaver Scheme Transfers (Year ended 30 June {fy_yr}): {tr_tot:,} inter-scheme transfers (vs {prev_tr:,} in FY{prev_fy}).",
                "source_html": (
                    f'Source: <a href="{IRD_KS_TR_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Scheme Transfer Statistics</a> '
                    '(CC BY 4.0). Monthly average derived by HLF. General information only, not personalised KiwiSaver advice.'
                ),
                "latest_period": f"Year ended June {fy_yr}",
                "value": tr_tot,
                "latest_value": f"{tr_tot:,} transfers/yr ({mth_avg:,}/month)",
                "previous_period": f"Year ended June {prev_fy} ({prev_tr:,} transfers)",
                "change": f"{tr_tot - prev_tr:+,d} transfers YoY",
                "calculation_live": f"{tr_tot:,} annual transfers ÷ 12 months = {mth_avg:,}/month",
                "historical_data_from": f"Annual June FY series (FY{prev_fy}–FY{fy_yr})",
            }

    return out


# ---------------------------------------------------------------------------
# Reference-Layer Enricher & Canonical Data Dictionary / JSON Builders
# ---------------------------------------------------------------------------
def enrich_ticker_record(key: str, item: dict, verified_date_iso: str, verified_human: str) -> dict:
    meta = TICKER_CATALOG.get(key, {})
    plain_summary = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", item.get("html", ""))).strip()
    canonical_url = f"{PULSE_BASE_URL}/{key}/"
    is_derived = bool(meta.get("derived_by_hlf", True))
    obs_period = item.get("as_of", "")[:7]
    src_pub = item.get("source_published") or item.get("as_of", verified_date_iso)

    citation_chain = {
        "official_source": meta.get("source_full", item.get("chip", "Official NZ Crown Release")),
        "official_source_url": meta.get("source_url", "https://www.rbnz.govt.nz/"),
        "pulse_record_url": canonical_url,
        "hlf_article_title": meta.get("article_title", "HLF News & Market Insights"),
        "hlf_article_url": meta.get("article_url", "https://www.homeloanfactory.co.nz/news-insights"),
        "hlf_tool_title": meta.get("tool_title", "HLF Interactive Mortgage & KiwiSaver Tools"),
        "hlf_tool_url": meta.get("tool_url", "https://www.homeloanfactory.co.nz/tools"),
        "hlf_advice_url": "https://www.homeloanfactory.co.nz/contact",
    }

    enriched = dict(item)
    enriched.update({
        "ticker": key,
        "name": meta.get("name", key.replace("-", " ").title()),
        "also_known_as": meta.get("also_known_as", ""),
        "category": meta.get("category", "Mortgage lending & rates"),
        "topic": meta.get("topic", "New Zealand mortgage & housing statistics"),
        "source": meta.get("source", item.get("chip", "NZ Crown Data")),
        "source_full": meta.get("source_full", item.get("chip", "NZ Crown Data")),
        "source_series_ids": meta.get("series_ids", "Official agency series"),
        "source_url": meta.get("source_url", "https://www.rbnz.govt.nz/"),
        "period": obs_period,
        "observed_period": obs_period,
        "source_published": src_pub,
        "data_status": "verified",
        "source_status": "current",
        "latest_period": item.get("latest_period", item.get("as_of", "")),
        "latest_value": item.get("latest_value", "See summary"),
        "unit": meta.get("unit", "value"),
        "previous_period": item.get("previous_period", "Prior comparable period"),
        "change": item.get("change", "Tracked across official releases"),
        "data_type": meta.get("data_type", "Officially reported & HLF-derived"),
        "derived_by_hlf": is_derived,
        "hlf_derived_label": "Yes" if is_derived else "No (Officially reported)",
        "definition": meta.get("definition", plain_summary),
        "official_data_summary": item.get("official_data_summary", f"Published in {meta.get('source_full', 'official release')}."),
        "calculation": item.get("calculation_live", meta.get("hlf_calculation", "Verified from official agency release")),
        "formula": meta.get("hlf_calculation", "Standardised from official agency release"),
        "hlf_interpretation": meta.get("hlf_interpretation", "Provides historical market context for New Zealand borrowers and researchers."),
        "methodology": meta.get("methodology", "Compiled from official New Zealand Crown statistical releases."),
        "update_frequency": meta.get("frequency", "Monthly"),
        "last_verified": verified_date_iso,
        "last_verified_human": verified_human,
        "historical_data_from": item.get("historical_data_from", "Current release"),
        "status": "Verified",
        "canonical_url": canonical_url,
        "summary_text": plain_summary,
        "related_questions": meta.get("questions", []),
        "citation_chain": citation_chain,
        "history": item.get("history", []),
        "creator_name": CREATOR_NAME,
        "creator_alternate_name": CREATOR_ALT_NAME,
        "creator_url": CREATOR_URL,
        "publisher": PUBLISHER_NAME,
        "publisher_url": PUBLISHER_URL,
        "date_created": PULSE_DATE_CREATED,
        "date_modified": verified_date_iso,
    })
    return enriched


def render_creator_card_html(compact: bool = False) -> str:
    """Render the Creator Attribution block featuring Andy Palliser's zoomed headshot."""
    if compact:
        return f"""
    <div class="creator-card is-compact" aria-label="Dataset Creator Attribution">
      <div class="creator-avatar" aria-hidden="true">
        <img src="{CREATOR_IMAGE}" alt="{CREATOR_NAME} — {CREATOR_ROLE}">
      </div>
      <div class="creator-body">
        <div class="creator-hdr">
          <strong>Created by <a href="{CREATOR_URL}" target="_blank" rel="noopener">{CREATOR_NAME}</a></strong>
          <span class="creator-badge">{CREATOR_BADGE}</span>
        </div>
        <div class="creator-role">{CREATOR_ROLE} &bull; Published by <a href="{PUBLISHER_URL}" target="_blank" rel="noopener">{PUBLISHER_NAME}</a></div>
        <div class="creator-links"><a href="#about-hlf-pulse">About the creator &amp; dataset</a> &middot; <a href="{CREATOR_URL}" target="_blank" rel="noopener">View Andy&rsquo;s adviser profile &rarr;</a></div>
      </div>
    </div>"""
    return f"""
    <section class="creator-card" id="creator-attribution" aria-label="Created by Andy Palliser">
      <div class="creator-avatar" aria-hidden="true">
        <img src="{CREATOR_IMAGE}" alt="{CREATOR_NAME} — {CREATOR_ROLE}">
      </div>
      <div class="creator-body">
        <div class="creator-hdr">
          <strong>Created by <a href="{CREATOR_URL}" target="_blank" rel="noopener">{CREATOR_NAME}</a></strong>
          <span class="creator-badge">{CREATOR_BADGE}</span>
        </div>
        <div class="creator-role">{CREATOR_ROLE}</div>
        <p>HLF Pulse was developed by {CREATOR_NAME} to make New Zealand mortgage, housing and KiwiSaver data easier to access, understand and use. Drawing on his experience in mortgage advice, Andy developed the dataset to bring together official government statistics, transparent calculations and practical mortgage-market interpretation.</p>
        <p style="margin-bottom:8px;">The underlying data is sourced from the relevant official organisations, with HLF-derived calculations and interpretations clearly identified.</p>
        <div class="creator-links"><a href="#about-hlf-pulse">About the creator</a> &middot; <a href="{CREATOR_URL}" target="_blank" rel="noopener">View Andy&rsquo;s adviser profile &rarr;</a></div>
      </div>
    </section>"""


def render_c41_band_table_html(rec: dict) -> str:
    bands = rec.get("tdti_bands")
    if rec.get("ticker") != "fhb-income-akl-vs-nz" or not bands:
        return ""
    per = html_lib.escape(rec["latest_period"])
    rows = "".join(
        f"<tr><td>{html_lib.escape(b['band'])} (Officially reported)</td><td><strong>${int(b['avg_income_nzd']):,}</strong></td><td>${float(b['total_income_m']):,.1f}m</td></tr>"
        for b in bands
    )
    rows += (
        f"<tr style='background:#eff6ff;'><td><strong>All NZ FHB Applications (HLF Borrower-Weighted Mean)</strong></td>"
        f"<td><strong>${int(rec['value']):,}</strong></td><td><strong>100% of bands</strong></td></tr>"
    )
    return (
        f"<div class='hist-wrap' style='margin:8px 0 10px;max-width:580px;'><table>"
        f"<thead><tr><th>FHB Debt-to-Income Band ({per})</th><th>Avg Gross Income</th><th>Band Income Pool</th></tr></thead>"
        f"<tbody>{rows}</tbody></table></div>"
    )


def render_history_table_html(hist: list[dict]) -> str:
    if not hist:
        return ""
    col1_hdr = "Region / Period" if any(not re.match(r"^\d{4}-\d{2}$", str(h.get("period", ""))) for h in hist) else "Period"
    rows = "".join(
        f"<tr><td>{html_lib.escape(str(h.get('label', h.get('period', ''))))}</td>"
        f"<td><strong>{html_lib.escape(str(h.get('formatted', h.get('value', ''))))}</strong></td></tr>"
        for h in hist
    )
    return (
        f"<div class='hist-wrap'><table><thead><tr><th>{col1_hdr}</th><th>Value</th></tr></thead>"
        f"<tbody>{rows}</tbody></table></div>"
    )


def render_three_layer_badges_html(rec: dict) -> str:
    return (
        f"<div class='layer-stack'>"
        f"<div class='layer-row'><span class='lbadge is-off'>Official data</span><span>{html_lib.escape(rec['official_data_summary'])}</span></div>"
        f"<div class='layer-row'><span class='lbadge is-calc'>HLF calculation</span><span><code>{html_lib.escape(rec['calculation'])}</code></span></div>"
        f"<div class='layer-row'><span class='lbadge is-interp'>HLF interpretation</span><span>{html_lib.escape(rec['hlf_interpretation'])}</span></div>"
        f"</div>"
    )


def render_about_and_quality_footer_html() -> str:
    return f"""
    <section id="data-quality-limitations" class="quality-box" aria-label="Data Quality &amp; Limitations">
      <h2>Data Quality &amp; Limitations</h2>
      <ul>
        <li><strong>Official published sources:</strong> Pulse uses official published data from the Reserve Bank of New Zealand (RBNZ), Stats NZ, Inland Revenue (IRD), and MBIE Tenancy Services wherever available.</li>
        <li><strong>Unaltered official figures:</strong> HLF does not alter or estimate official source figures.</li>
        <li><strong>Explicit derivation labels:</strong> Derived statistics (such as percentages, averages, rolling sums, and cohort comparisons) are calculated directly from the underlying official series and identified as <em>HLF-derived</em>.</li>
        <li><strong>Reporting lags &amp; differing schedules:</strong> Reporting periods differ across agencies (for example, RBNZ lending series publish monthly with a ~4-week lag, MBIE bond data publishes on a ~2-month lag, CPI and student loan balances publish quarterly, and regional KiwiSaver breakdowns publish annually for the June financial year).</li>
        <li><strong>Commitments vs. settlements:</strong> RBNZ mortgage lending tables (C31, C32, C33, C40, C41) measure firm bank mortgage commitments/approvals rather than LINZ property title settlements.</li>
        <li><strong>Publisher revisions:</strong> Underlying historical figures may subsequently be revised by the original Crown publisher in later releases.</li>
        <li><strong>Automatic staleness exclusion:</strong> HLF automatically excludes any series where the underlying source is unavailable, stale past its freshness window, or cannot be verified.</li>
        <li><strong>Market context vs. individual lending outcomes:</strong> Statistics describe aggregate historical market cohorts and do not predict an individual borrower&rsquo;s credit assessment or lending outcome.</li>
      </ul>
    </section>

    <section id="about-hlf-pulse" class="about-footer-box" aria-label="About the HLF Pulse &amp; Creator">
      <h2>About the HLF Pulse &amp; Creator</h2>
      <p><strong>HLF Pulse</strong> is a public reference dataset covering New Zealand mortgage, housing and KiwiSaver statistics, published by <a href="{PUBLISHER_URL}">{PUBLISHER_NAME}</a> and created by <a href="{CREATOR_URL}" target="_blank" rel="noopener"><strong>{CREATOR_NAME}</strong> ({CREATOR_ALT_NAME})</a>, Mortgage &amp; KiwiSaver Adviser at Home Loan Factory.</p>
      <p>Data is compiled from official releases from the <strong>Reserve Bank of New Zealand</strong>, <strong>Stats NZ</strong>, <strong>Inland Revenue</strong> and <strong>MBIE Tenancy Services</strong>. Where HLF calculates percentages, averages or comparisons from the underlying official data, these are identified as <strong>&ldquo;Derived by HLF&rdquo;</strong>. The underlying source and relevant reporting period are provided for each statistic.</p>
      <p>HLF does not alter or estimate official statistics. Where source data is unavailable, stale or cannot be verified, the relevant series is excluded until it can be confirmed.</p>
      <p style="margin-bottom:0;">Figures are provided for general information and historical market analysis and are not personalised financial advice. Published by <a href="{PUBLISHER_URL}">{PUBLISHER_NAME}</a> (<a href="{CREATOR_URL}" target="_blank" rel="noopener">Creator Profile: {CREATOR_NAME}</a> &bull; <a href="https://www.homeloanfactory.co.nz/public-disclosures">Public Disclosures</a> &bull; <a href="https://www.homeloanfactory.co.nz/how-we-get-paid">How We Get Paid</a>).</p>
    </section>"""


def write_canonical_ticker_page(root: pathlib.Path, key: str, rec: dict) -> None:
    t_dir = root / key
    t_dir.mkdir(parents=True, exist_ok=True)
    c_chain = rec["citation_chain"]
    aka_html = f"<div style='margin-top:4px;font-size:0.8rem;color:#6b7280;'><strong>Also known as:</strong> {html_lib.escape(rec['also_known_as'])}</div>" if rec.get("also_known_as") else ""
    q_items = "".join(f"<li>{html_lib.escape(q)}</li>" for q in rec.get("related_questions", []))
    q_block = f"<h2>Questions This Data Helps Answer</h2><ul>{q_items}</ul>" if q_items else ""
    c41_tbl = render_c41_band_table_html(rec)
    hist_html = render_history_table_html(rec.get("history", []))
    hist_section = f"<h2>Historical &amp; Cohort Breakdown ({html_lib.escape(rec['historical_data_from'])})</h2>{c41_tbl}{hist_html}" if (hist_html or c41_tbl) else ""
    three_layers = render_three_layer_badges_html(rec)
    creator_compact_html = render_creator_card_html(compact=True)
    about_footer = render_about_and_quality_footer_html()
    comp_row_lbl = "Regional range &amp; extremes" if key == "fhb-regional-kiwisaver" else "Comparison / Prior period"

    schema_obj = {
        "@context": "https://schema.org",
        "@type": "Dataset",
        "name": f"{rec['name']} — HLF Pulse Data Dictionary & Reference Record ({rec['latest_period']})",
        "alternateName": rec.get("also_known_as", ""),
        "description": rec["definition"],
        "url": rec["canonical_url"],
        "identifier": key,
        "inLanguage": "en-NZ",
        "temporalCoverage": rec["observed_period"],
        "dateCreated": PULSE_DATE_CREATED,
        "dateModified": rec["last_verified"],
        "spatialCoverage": "New Zealand",
        "isBasedOn": rec["source_url"],
        "creator": CREATOR_SCHEMA_OBJ,
        "publisher": PUBLISHER_SCHEMA_OBJ,
        "variableMeasured": rec["name"],
    }

    page_html = f"""<!DOCTYPE html>
<html lang="en-NZ">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="robots" content="index, follow">
  <link rel="canonical" href="{rec['canonical_url']}">
  <title>{html_lib.escape(rec['name'])} ({html_lib.escape(rec['latest_period'])}) — HLF Pulse Data Dictionary</title>
  <script type="application/ld+json">{json.dumps(schema_obj)}</script>
  <style>
    body {{ font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #111827; background: #f8fafc; margin: 0; padding: 26px 14px 48px; line-height: 1.6; font-variant-numeric: tabular-nums; }}
    main {{ max-width: 920px; margin: 0 auto; background: #ffffff; border-top: 5px solid #0b2559; border-radius: 16px; padding: 32px; box-shadow: 0 4px 20px rgba(11,37,89,0.05); }}
    .crumb {{ font-size: 0.82rem; margin-bottom: 14px; color: #6b7280; }}
    .crumb a {{ color: #186bd6; font-weight: 600; text-decoration: none; }}
    h1 {{ margin: 0 0 6px; color: #0b2559; font-size: 1.45rem; letter-spacing: -0.02em; }}
    h2 {{ margin: 26px 0 10px; color: #0b2559; font-size: 1.08rem; border-bottom: 1px solid #e5e7eb; padding-bottom: 6px; }}
    p {{ color: #374151; font-size: 0.9rem; margin: 0 0 12px; }}
    .summary-box {{ background: #eff6ff; border: 1px solid #dbeafe; border-left: 4px solid #186bd6; border-radius: 8px 14px 14px 8px; padding: 16px 18px; margin: 16px 0 18px; font-size: 0.98rem; color: #111827; }}
    .summary-box strong {{ color: #0b2559; }}
    .creator-card {{ display: flex; align-items: flex-start; gap: 16px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 14px; padding: 16px 18px; margin: 16px 0 22px; }}
    .creator-card.is-compact {{ align-items: center; padding: 12px 16px; margin: 14px 0 20px; }}
    .creator-avatar {{ width: 56px; height: 56px; border-radius: 50%; overflow: hidden; flex-shrink: 0; background: #0b2559; border: 2px solid #dbeafe; }}
    .creator-avatar img {{ width: 100%; height: 100%; object-fit: cover; object-position: 50% 20%; transform: scale(1.65); transform-origin: 50% 22%; display: block; }}
    .creator-hdr {{ display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 2px; font-size: 0.95rem; color: #0b2559; }}
    .creator-badge {{ background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; font-size: 0.7rem; font-weight: 700; padding: 2px 8px; border-radius: 999px; }}
    .creator-role {{ font-size: 0.8rem; color: #4b5563; font-weight: 500; margin-bottom: 4px; }}
    .creator-links {{ font-size: 0.8rem; }}
    .layer-stack {{ display: flex; flex-direction: column; gap: 8px; margin: 14px 0 22px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 12px; padding: 14px 16px; font-size: 0.85rem; }}
    .layer-row {{ display: flex; align-items: baseline; gap: 10px; }}
    .lbadge {{ flex-shrink: 0; display: inline-block; width: 128px; text-align: center; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.03em; padding: 3px 8px; border-radius: 6px; }}
    .lbadge.is-off {{ background: #ecfdf5; border: 1px solid #6ee7b7; color: #065f46; }}
    .lbadge.is-calc {{ background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; }}
    .lbadge.is-interp {{ background: #fefce8; border: 1px solid #fde68a; color: #92400e; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.86rem; margin-top: 8px; }}
    th {{ width: 215px; background: #f8fafc; color: #0b2559; text-align: left; padding: 9px 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }}
    td {{ padding: 9px 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; color: #374151; }}
    code {{ background: #eff6ff; padding: 2px 6px; border-radius: 4px; font-size: 0.8rem; color: #0b2559; }}
    a {{ color: #186bd6; font-weight: 600; }}
    ul {{ margin: 8px 0 0; padding-left: 20px; color: #374151; font-size: 0.88rem; }}
    li {{ margin-bottom: 6px; }}
    .chain {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(175px, 1fr)); gap: 10px; margin-top: 10px; }}
    .chain-step {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 10px; padding: 11px 13px; font-size: 0.82rem; }}
    .chain-step span {{ display: block; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; color: #6b7280; margin-bottom: 3px; }}
    .hist-wrap {{ max-width: 580px; border: 1px solid #e5e7eb; border-radius: 10px; overflow: hidden; margin-top: 10px; }}
    .hist-wrap th {{ width: auto; border-bottom: 2px solid #dbeafe; }}
    .quality-box {{ margin-top: 30px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 14px; padding: 20px 22px; }}
    .quality-box h2 {{ margin-top: 0; border-bottom: none; padding-bottom: 0; }}
    .about-footer-box {{ margin-top: 18px; background: #eff6ff; border: 1px solid #dbeafe; border-left: 4px solid #0b2559; border-radius: 8px 14px 14px 8px; padding: 20px 22px; }}
    .about-footer-box h2 {{ margin-top: 0; border-bottom: none; padding-bottom: 0; }}
  </style>
</head>
<body>
  <main>
    <div class="crumb"><a href="{PULSE_BASE_URL}/">&larr; All 30 HLF Pulse Datasets</a> &bull; Category: <strong>{html_lib.escape(rec['category'])}</strong></div>
    <h1>{html_lib.escape(rec['name'])}</h1>
    <div style="font-size:0.82rem;color:#6b7280;">Ticker: <code>{key}</code> &bull; Latest Underlying Data: <strong>{html_lib.escape(rec['latest_period'])}</strong> &bull; Last Checked: <strong>{html_lib.escape(rec['last_verified_human'])}</strong> &bull; Creator: <a href="{CREATOR_URL}" target="_blank" rel="noopener">{CREATOR_NAME}</a></div>
    {aka_html}

    <div class="summary-box">{rec['html']}</div>
    {three_layers}
    {creator_compact_html}

    <h2>Data Dictionary Specification</h2>
    <table>
      <tbody>
        <tr><th>Name</th><td><strong>{html_lib.escape(rec['name'])}</strong></td></tr>
        <tr><th>Ticker</th><td><code>{key}</code></td></tr>
        <tr><th>Also known as</th><td>{html_lib.escape(rec.get('also_known_as') or '—')}</td></tr>
        <tr><th>Creator &amp; Publisher</th><td>Created by <a href="{CREATOR_URL}" target="_blank" rel="noopener"><strong>{CREATOR_NAME}</strong> ({CREATOR_ALT_NAME})</a> &bull; Published by <a href="{PUBLISHER_URL}" target="_blank" rel="noopener">{PUBLISHER_NAME}</a></td></tr>
        <tr><th>Official source</th><td>{html_lib.escape(rec['source'])} (<a href="{rec['source_url']}" target="_blank" rel="noopener">{html_lib.escape(rec['source_full'])}</a>)</td></tr>
        <tr><th>Frequency</th><td>{html_lib.escape(rec['update_frequency'])}</td></tr>
        <tr><th>Latest underlying data</th><td><strong>{html_lib.escape(rec['latest_period'])}</strong> (Observed period: <code>{rec['observed_period']}</code> &bull; Period end: <code>{rec['as_of']}</code>)</td></tr>
        <tr><th>Source published &amp; status</th><td>Published: <code>{html_lib.escape(rec['source_published'])}</code> &bull; <code>data_status: {rec['data_status']}</code> &bull; <code>source_status: {rec['source_status']}</code></td></tr>
        <tr><th>Definition</th><td>{html_lib.escape(rec['definition'])}</td></tr>
        <tr><th>Current value</th><td><strong>{html_lib.escape(str(rec['latest_value']))}</strong></td></tr>
        <tr><th>{comp_row_lbl}</th><td>{html_lib.escape(str(rec['previous_period']))} &bull; {html_lib.escape(str(rec['change']))}</td></tr>
        <tr><th>Official data</th><td>{html_lib.escape(rec['official_data_summary'])}</td></tr>
        <tr><th>HLF calculation</th><td><code>{html_lib.escape(rec['calculation'])}</code><div style="margin-top:3px;font-size:0.78rem;color:#6b7280;">Formula: {html_lib.escape(rec['formula'])}</div></td></tr>
        <tr><th>HLF interpretation</th><td>{html_lib.escape(rec['hlf_interpretation'])}</td></tr>
        <tr><th>Source series IDs</th><td><code>{html_lib.escape(rec['source_series_ids'])}</code></td></tr>
        <tr><th>HLF derived</th><td><strong>{html_lib.escape(rec['hlf_derived_label'])}</strong> ({html_lib.escape(rec['data_type'])})</td></tr>
        <tr><th>Last checked</th><td>{html_lib.escape(rec['last_verified_human'])} (<code>{rec['last_verified']}</code>)</td></tr>
        <tr><th>Historical data</th><td>{html_lib.escape(rec['historical_data_from'])}</td></tr>
      </tbody>
    </table>

    {hist_section}
    {q_block}

    <h2>Citation Chain &amp; Related HLF Resources</h2>
    <div class="chain">
      <div class="chain-step"><span>1. Official Source</span><a href="{c_chain['official_source_url']}" target="_blank" rel="noopener">{html_lib.escape(rec['source'])}</a></div>
      <div class="chain-step"><span>2. Pulse Record</span><a href="{c_chain['pulse_record_url']}"><code>{key}</code></a></div>
      <div class="chain-step"><span>3. HLF Explanation</span><a href="{c_chain['hlf_article_url']}" target="_blank" rel="noopener">{html_lib.escape(c_chain['hlf_article_title'])}</a></div>
      <div class="chain-step"><span>4. Interactive Tool</span><a href="{c_chain['hlf_tool_url']}" target="_blank" rel="noopener">{html_lib.escape(c_chain['hlf_tool_title'])}</a></div>
      <div class="chain-step"><span>5. Personalised Advice</span><a href="{c_chain['hlf_advice_url']}" target="_blank" rel="noopener">Speak with an HLF Adviser</a></div>
    </div>

    {about_footer}
  </main>
</body>
</html>"""
    (t_dir / "index.html").write_text(page_html, encoding="utf-8")


# ---------------------------------------------------------------------------
# Main Orchestrator & Strict 30-Dataset Canonical Gate
# ---------------------------------------------------------------------------
def main() -> None:
    root = pathlib.Path(__file__).resolve().parent
    json_path = root / "nz-mortgage-pulse.json"
    structured_json_path = root / "pulse.json"
    html_path = root / "index.html"

    prev_tickers = {}
    if json_path.exists():
        try:
            prev_tickers = json.loads(json_path.read_text(encoding="utf-8")).get("tickers", {})
        except Exception:
            prev_tickers = {}

    fresh_candidates = {k: v for k, v in prev_tickers.items() if k in TICKER_CATALOG}
    session = get_http_session()
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

    now_dt = datetime.datetime.now(datetime.timezone.utc)
    now_utc = now_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    today_iso = now_dt.strftime("%Y-%m-%d")
    now_human = f"{now_dt.day} {now_dt.strftime('%B %Y')}"

    verified_tickers = {}
    for k, item in fresh_candidates.items():
        if k not in TICKER_CATALOG or not isinstance(item, dict):
            continue
        as_of = item.get("as_of", "")
        max_age = int(item.get("max_age_days", MAX_AGE_MONTHLY))
        if item.get("html") and is_fresh(as_of, max_age):
            verified_tickers[k] = enrich_ticker_record(k, item, today_iso, now_human)
        else:
            print(f"[Failsafe Gate] Omitted stale/invalid ticker '{k}' (as_of={as_of}, max_age={max_age}d)")

    if live_parsed_count == 0 and not verified_tickers:
        raise RuntimeError("Zero live datasets parsed and zero valid fresh tickers remain; failing workflow honestly.")

    bubble_feed_tickers = dict(verified_tickers)
    for alias_k, target_k in LEGACY_BUBBLE_ALIASES.items():
        if target_k in verified_tickers:
            bubble_feed_tickers[alias_k] = dict(verified_tickers[target_k])

    latest_rbnz_period = verified_tickers.get("fhb-low-deposit", {}).get("latest_period", "September 2026")
    payload = {
        "publisher": PUBLISHER_NAME,
        "publisher_url": PUBLISHER_URL,
        "creator_name": CREATOR_NAME,
        "creator_alternate_name": CREATOR_ALT_NAME,
        "creator_url": CREATOR_URL,
        "dataset_name": "NZ Mortgage, Housing & KiwiSaver Pulse — Verified NZ Housing & Mortgage Data",
        "date_created": PULSE_DATE_CREATED,
        "date_modified": today_iso,
        "generated_utc": now_utc,
        "last_checked": today_iso,
        "latest_underlying_data": latest_rbnz_period,
        "data_status": "verified",
        "source_status": "current",
        "active_datasets_count": len(verified_tickers),
        "count": len(verified_tickers),
        "tickers": bubble_feed_tickers,
    }
    json_path.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")

    structured_records = []
    for k in sorted(verified_tickers.keys()):
        r = verified_tickers[k]
        rec_out = {
            "ticker": k,
            "name": r["name"],
            "also_known_as": r.get("also_known_as", ""),
            "category": r["category"],
            "topic": r["topic"],
            "value": r.get("value"),
            "latest_value": r["latest_value"],
            "unit": r["unit"],
            "observed_period": r["observed_period"],
            "latest_underlying_data": r["latest_period"],
            "source_published": r["source_published"],
            "last_checked": r["last_verified"],
            "date_created": r["date_created"],
            "date_modified": r["date_modified"],
            "data_status": r["data_status"],
            "source_status": r["source_status"],
            "previous_period": r["previous_period"],
            "change": r["change"],
            "source": r["source"],
            "source_full": r["source_full"],
            "source_series_ids": r["source_series_ids"],
            "source_url": r["source_url"],
            "derived_by_hlf": r["derived_by_hlf"],
            "data_type": r["data_type"],
            "definition": r["definition"],
            "official_data": r["official_data_summary"],
            "hlf_calculation": r["calculation"],
            "formula": r["formula"],
            "hlf_interpretation": r["hlf_interpretation"],
            "methodology": r["methodology"],
            "frequency": r["update_frequency"],
            "historical_data_from": r["historical_data_from"],
            "creator_name": CREATOR_NAME,
            "creator_alternate_name": CREATOR_ALT_NAME,
            "creator_url": CREATOR_URL,
            "publisher": PUBLISHER_NAME,
            "publisher_url": PUBLISHER_URL,
            "canonical_url": r["canonical_url"],
            "summary_text": r["summary_text"],
            "related_questions": r["related_questions"],
            "citation_chain": r["citation_chain"],
            "history": r["history"],
        }
        if r.get("tdti_bands"):
            rec_out["tdti_bands"] = r["tdti_bands"]
        if r.get("townhouse_flat_unit_share_pct") is not None:
            rec_out["townhouse_flat_unit_share_pct"] = r["townhouse_flat_unit_share_pct"]
        structured_records.append(rec_out)

    pulse_json_obj = {
        "dataset": "HLF Pulse — Verified New Zealand Mortgage, Housing & KiwiSaver Reference Dataset",
        "publisher": PUBLISHER_NAME,
        "publisher_legal_name": PUBLISHER_LEGAL,
        "publisher_url": PUBLISHER_URL,
        "creator_name": CREATOR_NAME,
        "creator_alternate_name": CREATOR_ALT_NAME,
        "creator_role": CREATOR_ROLE,
        "creator_url": CREATOR_URL,
        "date_created": PULSE_DATE_CREATED,
        "date_modified": today_iso,
        "pulse_url": f"{PULSE_BASE_URL}/",
        "last_checked": today_iso,
        "dataset_generated_utc": now_utc,
        "latest_underlying_data": latest_rbnz_period,
        "data_status": "verified",
        "source_status": "current",
        "active_datasets_count": len(structured_records),
        "reporting_lag_note": "RBNZ and Crown mortgage statistics are published with a ~4-week reporting lag; last_checked reflects when HLF verified the source, while latest_underlying_data reflects the observation period.",
        "next_expected_update": "Monthly (RBNZ / Stats NZ / MBIE Tenancy) & Quarterly (CPI / IRD Student Loans)",
        "records": structured_records,
    }
    structured_json_path.write_text(json.dumps(pulse_json_obj, indent=2), encoding="utf-8")

    for k, rec in verified_tickers.items():
        write_canonical_ticker_page(root, k, rec)

    sitemap_urls = [f"  <url><loc>{PULSE_BASE_URL}/</loc><lastmod>{today_iso}</lastmod></url>"] + [
        f"  <url><loc>{PULSE_BASE_URL}/{k}/</loc><lastmod>{today_iso}</lastmod></url>"
        for k in sorted(verified_tickers.keys())
    ]
    (root / "sitemap.xml").write_text(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
        "<urlset xmlns=\"http://www.sitemaps.org/schemas/sitemap/0.9\">\n"
        + "\n".join(sitemap_urls)
        + "\n</urlset>\n",
        encoding="utf-8",
    )

    llms_lines = [
        "# HLF Pulse — Verified New Zealand Housing, Mortgage & KiwiSaver Reference Dataset",
        "",
        f"> Public reference dataset and data dictionary built from official New Zealand Crown releases, with clearly labelled HLF-derived calculations and interpretation.",
        f"> Created by: {CREATOR_NAME} ({CREATOR_ALT_NAME}), {CREATOR_ROLE} ({CREATOR_URL}) | Published by: {PUBLISHER_NAME} ({PUBLISHER_URL}).",
        f"> Latest underlying data: {latest_rbnz_period} | Last checked: {now_human} ({today_iso}) | Active datasets: {len(verified_tickers)}.",
        "",
        "## Machine-Readable Endpoints",
        f"- Structured Data Dictionary & Records (JSON): {PULSE_BASE_URL}/pulse.json",
        f"- Webflow Chat Bubble Feed (JSON): {PULSE_BASE_URL}/nz-mortgage-pulse.json",
        f"- Main Reference Dataset Hub (HTML): {PULSE_BASE_URL}/",
        "",
        f"## Canonical Data Dictionary Records ({len(verified_tickers)} Active Datasets)",
    ] + [
        f"- [{rec['name']} ({rec['latest_period']})]({rec['canonical_url']}): {rec['summary_text']} [Official Data: {rec['official_data_summary']} | HLF Calculation: {rec['calculation']}]"
        for k, rec in sorted(verified_tickers.items())
    ]
    (root / "llms.txt").write_text("\n".join(llms_lines) + "\n", encoding="utf-8")

    by_cat = {c: [] for c in CATEGORY_ORDER}
    for k, rec in sorted(verified_tickers.items()):
        cat = rec.get("category", "Mortgage lending & rates")
        by_cat.setdefault(cat, []).append((k, rec))

    cat_cards_html = ""
    for cat in CATEGORY_ORDER:
        items = by_cat.get(cat, [])
        if not items:
            continue
        cat_slug = re.sub(r"[^a-z0-9]+", "-", cat.lower()).strip("-")
        links = "".join(
            f"<li><a href='#{k}'><code>{k}</code></a> — {html_lib.escape(r['name'])} (<strong>{html_lib.escape(str(r['latest_value']).split('(')[0].strip())}</strong>)</li>"
            for k, r in items
        )
        cat_cards_html += (
            f"<div class='cat-card' id='cat-{cat_slug}'>"
            f"<div class='cat-hdr'><strong>{html_lib.escape(cat)}</strong><span>{len(items)} datasets</span></div>"
            f"<ul>{links}</ul></div>"
        )

    rows_html_list = []
    for k, v in sorted(verified_tickers.items()):
        c_chain = v["citation_chain"]
        aka_sub = f"<div style='margin-top:3px;font-size:0.71rem;color:#6b7280;'>Also: {html_lib.escape(v['also_known_as'])}</div>" if v.get("also_known_as") else ""
        c41_tbl = render_c41_band_table_html(v)
        hist_tbl = render_history_table_html(v.get("history", []))
        three_layers = render_three_layer_badges_html(v)
        q_list = "".join(f"<li>{html_lib.escape(q)}</li>" for q in v.get("related_questions", []))
        comp_lbl = "Regional range &amp; extremes" if k == "fhb-regional-kiwisaver" else "Comparison / Prior period"
        rows_html_list.append(
            f"<tr id='{k}'>"
            f"<td><a href='{v['canonical_url']}' title='Open standalone Data Dictionary page for {k}'><code>{k}</code></a>"
            f"<div style='margin-top:4px;font-size:0.76rem;color:#111827;font-weight:600;'>{html_lib.escape(v['name'])}</div>"
            f"{aka_sub}"
            f"<span class='dtype-badge'>{'Derived by HLF' if v['derived_by_hlf'] else 'Officially Reported'}</span></td>"
            f"<td><strong>{html_lib.escape(v['source'])}</strong>"
            f"<div style='margin-top:3px;font-size:0.76rem;color:#0b2559;font-weight:700;'>Data: {html_lib.escape(v['latest_period'])}</div>"
            f"<div style='font-size:0.71rem;color:#6b7280;'>Checked: {html_lib.escape(v['last_verified_human'])}</div></td>"
            f"<td><div style='font-size:0.9rem;color:#111827;'>{v['html']}</div>"
            f"{c41_tbl}"
            f"{three_layers}"
            f"<div style='margin-top:5px;font-size:0.76rem;color:#6b7280;'>{v['source_html']} &bull; <a href='{v['canonical_url']}'>Open Data Dictionary page &rarr;</a></div>"
            f"<details class='rec-drawer'><summary>Data Dictionary: definition, formula, history &amp; citation chain</summary>"
            f"<div class='rec-drawer-body'>"
            f"<table class='dict-mini'>"
            f"<tr><th>Name</th><td><strong>{html_lib.escape(v['name'])}</strong></td></tr>"
            f"<tr><th>Also known as</th><td>{html_lib.escape(v.get('also_known_as') or '—')}</td></tr>"
            f"<tr><th>Creator &amp; Publisher</th><td>Created by <a href='{CREATOR_URL}' target='_blank' rel='noopener'>{CREATOR_NAME} ({CREATOR_ALT_NAME})</a> &bull; Published by <a href='{PUBLISHER_URL}' target='_blank' rel='noopener'>{PUBLISHER_NAME}</a></td></tr>"
            f"<tr><th>Source &amp; Series</th><td>{html_lib.escape(v['source_full'])} (<code>{html_lib.escape(v['source_series_ids'])}</code>)</td></tr>"
            f"<tr><th>Frequency</th><td>{html_lib.escape(v['update_frequency'])}</td></tr>"
            f"<tr><th>Latest underlying data</th><td><strong>{html_lib.escape(v['latest_period'])}</strong> (Observed period: <code>{v['observed_period']}</code> &bull; Published: <code>{html_lib.escape(v['source_published'])}</code>)</td></tr>"
            f"<tr><th>Data &amp; Source Status</th><td><code>data_status: {v['data_status']}</code> &bull; <code>source_status: {v['source_status']}</code></td></tr>"
            f"<tr><th>Definition</th><td>{html_lib.escape(v['definition'])}</td></tr>"
            f"<tr><th>Current value</th><td><strong>{html_lib.escape(str(v['latest_value']))}</strong></td></tr>"
            f"<tr><th>{comp_lbl}</th><td>{html_lib.escape(str(v['previous_period']))} &bull; {html_lib.escape(str(v['change']))}</td></tr>"
            f"<tr><th>HLF calculation</th><td><code>{html_lib.escape(v['calculation'])}</code></td></tr>"
            f"<tr><th>HLF derived</th><td><strong>{html_lib.escape(v['hlf_derived_label'])}</strong> ({html_lib.escape(v['data_type'])})</td></tr>"
            f"<tr><th>Last checked</th><td>{html_lib.escape(v['last_verified_human'])} (<code>{v['last_verified']}</code>)</td></tr>"
            f"<tr><th>Historical data</th><td>{html_lib.escape(v['historical_data_from'])}</td></tr>"
            f"</table>"
            + (f"<div style='margin-top:8px;'><strong>Questions This Data Helps Answer:</strong><ul style='margin:4px 0 0;padding-left:18px;'>{q_list}</ul></div>" if q_list else "")
            + (f"<div style='margin-top:8px;'><strong>Historical / Regional Series:</strong>{hist_tbl}</div>" if hist_tbl else "")
            + f"<div style='margin-top:8px;'><strong>Citation Chain:</strong> <a href='{c_chain['official_source_url']}' target='_blank' rel='noopener'>1. {html_lib.escape(v['source'])}</a> &rarr; "
            f"<a href='{v['canonical_url']}'>2. Data Dictionary (/{k})</a> &rarr; "
            f"<a href='{c_chain['hlf_article_url']}' target='_blank' rel='noopener'>3. {html_lib.escape(c_chain['hlf_article_title'])}</a> &rarr; "
            f"<a href='{c_chain['hlf_tool_url']}' target='_blank' rel='noopener'>4. {html_lib.escape(c_chain['hlf_tool_title'])}</a> &rarr; "
            f"<a href='{c_chain['hlf_advice_url']}' target='_blank' rel='noopener'>5. HLF Mortgage &amp; KiwiSaver Advice</a></div>"
            f"</div></details></td></tr>"
        )
    rows_html = "".join(rows_html_list)
    robots_tag = "index, follow" if verified_tickers else "noindex, nofollow"
    creator_main_html = render_creator_card_html(compact=False)
    about_footer_html = render_about_and_quality_footer_html()

    hub_schema_obj = {
        "@context": "https://schema.org",
        "@type": "Dataset",
        "name": "HLF Pulse — Verified New Zealand Mortgage, Housing & KiwiSaver Reference Dataset",
        "description": "Public reference dataset and data dictionary built from official New Zealand Crown releases, with clearly labelled HLF-derived calculations and interpretation.",
        "url": f"{PULSE_BASE_URL}/",
        "inLanguage": "en-NZ",
        "dateCreated": PULSE_DATE_CREATED,
        "dateModified": today_iso,
        "spatialCoverage": "New Zealand",
        "creator": CREATOR_SCHEMA_OBJ,
        "publisher": PUBLISHER_SCHEMA_OBJ,
        "hasPart": [v["canonical_url"] for _, v in sorted(verified_tickers.items())],
    }

    html_path.write_text(
        f"""<!DOCTYPE html>
<html lang="en-NZ">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="robots" content="{robots_tag}">
  <link rel="canonical" href="{PULSE_BASE_URL}/">
  <title>NZ Mortgage, Housing &amp; KiwiSaver Pulse — Verified NZ Housing &amp; Mortgage Data | HLF</title>
  <script type="application/ld+json">{json.dumps(hub_schema_obj)}</script>
  <style>
    body {{ font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #111827; background: #f8fafc; margin: 0; padding: 24px 14px 56px; line-height: 1.58; font-variant-numeric: tabular-nums; }}
    main {{ max-width: 1140px; margin: 0 auto; background: #ffffff; border-top: 5px solid #0b2559; border-radius: 16px; padding: 32px; box-shadow: 0 4px 20px rgba(11,37,89,0.05); }}
    h1 {{ margin: 0 0 8px; color: #0b2559; font-size: 1.55rem; letter-spacing: -0.02em; }}
    h2 {{ margin: 28px 0 10px; color: #0b2559; font-size: 1.18rem; letter-spacing: -0.01em; }}
    p {{ color: #374151; font-size: 0.92rem; margin: 0 0 12px; }}
    a {{ color: #186bd6; font-weight: 600; }}
    .lag-banner {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 14px; padding: 16px 18px; margin: 16px 0 18px; }}
    .lag-item span {{ display: block; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; color: #186bd6; margin-bottom: 2px; }}
    .lag-item strong {{ font-size: 0.98rem; color: #0b2559; }}
    .lag-item small {{ display: block; font-size: 0.76rem; color: #374151; margin-top: 2px; }}
    .creator-card {{ display: flex; align-items: flex-start; gap: 18px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 14px; padding: 18px 20px; margin: 0 0 24px; }}
    .creator-avatar {{ width: 64px; height: 64px; border-radius: 50%; overflow: hidden; flex-shrink: 0; background: #0b2559; border: 2px solid #dbeafe; box-shadow: 0 2px 8px rgba(11,37,89,0.12); }}
    .creator-avatar img {{ width: 100%; height: 100%; object-fit: cover; object-position: 50% 20%; transform: scale(1.65); transform-origin: 50% 22%; display: block; }}
    .creator-hdr {{ display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 2px; font-size: 1rem; color: #0b2559; }}
    .creator-badge {{ background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; font-size: 0.72rem; font-weight: 700; padding: 2px 9px; border-radius: 999px; }}
    .creator-role {{ font-size: 0.83rem; color: #4b5563; font-weight: 600; margin-bottom: 8px; }}
    .creator-body p {{ font-size: 0.87rem; color: #374151; margin: 0 0 8px; line-height: 1.55; }}
    .creator-links {{ font-size: 0.82rem; }}
    .use-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 12px; margin: 12px 0 24px; }}
    .use-card {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 12px; padding: 14px 16px; font-size: 0.83rem; }}
    .use-card strong {{ display: block; color: #0b2559; font-size: 0.9rem; margin-bottom: 6px; }}
    .use-card ul {{ margin: 0; padding-left: 18px; color: #374151; }}
    .use-card li {{ margin-bottom: 4px; }}
    .cat-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 12px; margin: 12px 0 26px; }}
    .cat-card {{ background: #ffffff; border: 1px solid #e5e7eb; border-radius: 12px; padding: 14px 16px; font-size: 0.82rem; }}
    .cat-hdr {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #f1f5f9; padding-bottom: 6px; margin-bottom: 8px; color: #0b2559; }}
    .cat-hdr span {{ background: #eff6ff; color: #186bd6; font-size: 0.72rem; font-weight: 700; padding: 2px 8px; border-radius: 999px; }}
    .cat-card ul {{ margin: 0; padding-left: 16px; color: #374151; }}
    .cat-card li {{ margin-bottom: 4px; }}
    .dtype-badge {{ display: inline-block; margin-top: 6px; background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; font-size: 0.69rem; font-weight: 700; padding: 2px 7px; border-radius: 999px; }}
    .layer-stack {{ display: flex; flex-direction: column; gap: 5px; margin: 8px 0 6px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 8px; padding: 8px 10px; font-size: 0.78rem; color: #374151; }}
    .layer-row {{ display: flex; align-items: baseline; gap: 8px; }}
    .lbadge {{ flex-shrink: 0; display: inline-block; width: 112px; text-align: center; font-size: 0.66rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.03em; padding: 2px 6px; border-radius: 5px; }}
    .lbadge.is-off {{ background: #ecfdf5; border: 1px solid #6ee7b7; color: #065f46; }}
    .lbadge.is-calc {{ background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; }}
    .lbadge.is-interp {{ background: #fefce8; border: 1px solid #fde68a; color: #92400e; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; margin-top: 12px; }}
    th {{ background: #f8fafc; color: #0b2559; text-align: left; padding: 10px 12px; border-bottom: 2px solid #dbeafe; }}
    td {{ padding: 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }}
    code {{ background: #eff6ff; padding: 2px 6px; border-radius: 4px; font-size: 0.78rem; color: #0b2559; }}
    .rec-drawer {{ margin-top: 8px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 8px; padding: 6px 10px; font-size: 0.79rem; }}
    .rec-drawer summary {{ cursor: pointer; font-weight: 600; color: #0b2559; user-select: none; }}
    .rec-drawer-body {{ margin-top: 8px; padding-top: 8px; border-top: 1px solid #e5e7eb; color: #374151; line-height: 1.52; }}
    .dict-mini {{ margin-top: 2px; font-size: 0.78rem; background: #ffffff; border: 1px solid #e5e7eb; border-radius: 6px; overflow: hidden; }}
    .dict-mini th {{ width: 155px; padding: 6px 9px; border-bottom: 1px solid #e5e7eb; background: #f8fafc; }}
    .dict-mini td {{ padding: 6px 9px; border-bottom: 1px solid #e5e7eb; }}
    .hist-wrap {{ max-width: 580px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; margin-top: 6px; background: #ffffff; }}
    .hist-wrap table {{ margin-top: 0; font-size: 0.77rem; }}
    .hist-wrap th, .hist-wrap td {{ padding: 5px 9px; }}
    .quality-box {{ margin-top: 32px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 14px; padding: 22px 24px; }}
    .quality-box h2 {{ margin-top: 0; }}
    .quality-box ul {{ margin: 8px 0 0; padding-left: 20px; font-size: 0.86rem; color: #374151; }}
    .quality-box li {{ margin-bottom: 6px; }}
    .about-footer-box {{ margin-top: 20px; background: #eff6ff; border: 1px solid #dbeafe; border-left: 5px solid #0b2559; border-radius: 8px 14px 14px 8px; padding: 22px 24px; }}
    .about-footer-box h2 {{ margin-top: 0; }}
    @media (max-width: 600px) {{
      .creator-card {{ flex-direction: column; gap: 12px; }}
    }}
  </style>
</head>
<body>
  <main>
    <h1>NZ Mortgage, Housing &amp; KiwiSaver Pulse — Verified NZ Housing &amp; Mortgage Data</h1>
    <p>Public reference dataset and data dictionary built from official New Zealand Crown releases, with clearly labelled HLF-derived calculations and interpretation (<a href="#browse-datasets">Browse all {len(verified_tickers)} datasets</a> &bull; <a href="#data-quality-limitations">Data Quality &amp; Limitations</a> &bull; <a href="#about-hlf-pulse">About &amp; Methodology</a>). Machine-readable endpoints: <a href="{PULSE_BASE_URL}/pulse.json"><code>pulse.json</code></a>, <a href="{PULSE_BASE_URL}/nz-mortgage-pulse.json"><code>nz-mortgage-pulse.json</code></a>, and <a href="{PULSE_BASE_URL}/llms.txt"><code>llms.txt</code></a>.</p>

    <div class="lag-banner" aria-label="Latest Underlying Data Period vs. Verification Date">
      <div class="lag-item">
        <span>Latest Underlying Data</span>
        <strong>{html_lib.escape(latest_rbnz_period)} (RBNZ / Monthly Series)</strong>
        <small>Why: RBNZ and Crown mortgage statistics are published ~4 weeks after month-end.</small>
      </div>
      <div class="lag-item">
        <span>Last Checked &amp; Verified</span>
        <strong>{now_human} ({today_iso})</strong>
        <small>Automated source refresh &amp; freshness verification date.</small>
      </div>
      <div class="lag-item">
        <span>Data Status</span>
        <strong style="color:#065f46;">&#10003; Verified (<a href="#browse-datasets" style="color:#065f46;text-decoration:underline;">{len(verified_tickers)} Active Datasets &darr;</a>)</strong>
        <small>Sources: RBNZ, Stats NZ, MBIE Tenancy Services, Inland Revenue.</small>
      </div>
    </div>

    {creator_main_html}

    <h2>What the HLF Pulse Can Help Answer</h2>
    <div class="use-grid">
      <div class="use-card">
        <strong>For First-Home Buyers</strong>
        <ul>
          <li><a href="#fhb-low-deposit">How common is a deposit under 20%?</a></li>
          <li><a href="#fhb-loan-size-gap">What are FHBs actually borrowing?</a></li>
          <li><a href="#fhb-regional-kiwisaver">How much are FHBs withdrawing from KiwiSaver?</a></li>
          <li><a href="#fhb-income-akl-vs-nz">How does my income compare with approved borrowers?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For Homeowners</strong>
        <ul>
          <li><a href="#top-ups">How common are mortgage top-ups?</a></li>
          <li><a href="#refinance">How much are homeowners refinancing?</a></li>
          <li><a href="#owner-equity-buffer">How much equity do existing homeowners have?</a></li>
          <li><a href="#mover-low-equity">Can existing homeowners move with under 20% equity?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For Property Investors</strong>
        <ul>
          <li><a href="#interest-only">How much investor lending is interest-only?</a></li>
          <li><a href="#investor-equity-tiers">What proportion of investors borrow with under 30% or 40% equity?</a></li>
          <li><a href="#investor-low-deposit">How many investors buy with under a 20% deposit?</a></li>
          <li><a href="#rent-yield-pulse">Where are national and regional median weekly rents sitting?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For Journalists &amp; Researchers</strong>
        <ul>
          <li><a href="#fhb-vs-investors">What are the latest RBNZ mortgage lending trends?</a></li>
          <li><a href="#settled-rates">What are NZ fixed mortgage rates actually settling at?</a></li>
          <li><a href="#dti-headroom">How much headroom sits beneath RBNZ&rsquo;s 20% DTI speed limit?</a></li>
          <li><a href="#townhouse-share">How is multi-unit and townhouse supply changing?</a></li>
        </ul>
      </div>
    </div>

    <h2 id="browse-datasets">Browse All {len(verified_tickers)} Datasets by Category</h2>
    <div class="cat-grid">{cat_cards_html}</div>

    <h2>Verified NZ Housing, Mortgage &amp; KiwiSaver Data Dictionary ({len(verified_tickers)} Active Datasets)</h2>
    <p>Every dataset separates <span class="lbadge is-off" style="width:auto;">Official data</span>, <span class="lbadge is-calc" style="width:auto;">HLF calculation</span>, and <span class="lbadge is-interp" style="width:auto;">HLF interpretation</span>. Click any Ticker Key to open its standalone canonical page (<code>/{'{ticker}'}/</code>).</p>
    <table>
      <thead><tr><th style="width:230px;">Ticker &amp; Derivation</th><th style="width:170px;">Source &amp; Data Period</th><th>Verified Statistic, 3-Layer Breakdown &amp; Expandable Data Dictionary</th></tr></thead>
      <tbody>{rows_html}</tbody>
    </table>

    {about_footer_html}
  </main>
</body>
</html>""",
        encoding="utf-8",
    )
    print(f"Saved {len(verified_tickers)} canonical datasets, pulse.json, index.html, llms.txt, sitemap.xml & /<ticker>/ pages.")


if __name__ == "__main__":
    main()
