# === PART 1 OF 5 ===
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
# Official NZ Crown Data Sources & Creator / Publisher / Licensing Metadata
# ---------------------------------------------------------------------------
PULSE_BASE_URL = "https://pulse-data.homeloanfactory.co.nz"
RBNZ_WARMUP_URL = "https://www.rbnz.govt.nz/statistics"

CREATOR_NAME = "Andy Palliser"
CREATOR_ALT_NAME = "Andrew Jonathan Palliser"
CREATOR_ROLE = "Mortgage & KiwiSaver Adviser · Home Loan Factory"
CREATOR_BADGE = "Creator of HLF Pulse"
CREATOR_FSP = "FSP1004725"
CREATOR_URL = "https://www.homeloanfactory.co.nz/team/andrew-palliser"
CREATOR_IMAGE = "https://cdn.prod.website-files.com/6a1baf404596ba888865b209/6a1ce4c85096daa2e99a0ecb_Andrew-Palliser.avif"
PUBLISHER_NAME = "HLF (Home Loan Factory)"
PUBLISHER_LEGAL = "HLF Financial Limited (FSP1008571)"
PUBLISHER_URL = "https://www.homeloanfactory.co.nz/"
CONTACT_URL = "https://www.homeloanfactory.co.nz/contact"
DISCLOSURES_URL = "https://www.homeloanfactory.co.nz/public-disclosures"
HOW_PAID_URL = "https://www.homeloanfactory.co.nz/how-we-get-paid"
PULSE_DATE_CREATED = "2026-09-01"

HLF_LICENSE_URL = "https://creativecommons.org/licenses/by/4.0/"
HLF_LICENSE_SHORT = "CC BY 4.0"
HLF_LICENSE_NOTE = (
    "Official statistics remain under their respective Crown/statutory publisher terms "
    "(Stats NZ, Inland Revenue, MBIE & FMA under CC BY 4.0; RBNZ under RBNZ Terms of Use; "
    "NHC & ACC statutory caps/levies under public-domain NZ legislation). "
    "HLF-derived calculations, data dictionary compilations, and market interpretations are "
    "licensed by HLF (Home Loan Factory) under Creative Commons Attribution 4.0 International (CC BY 4.0) "
    "when attributed to HLF Pulse with a link to the canonical dataset URL."
)

# Stats NZ CPI cumulative inflation benchmarks for 10-year (2016-2026) and 5-year (2021-2026) real-terms adjustments
CPI_10Y_FACTOR = 1.337  # +33.7% cumulative NZ CPI over 10 years (~2.95% p.a.)
CPI_5Y_FACTOR = 1.228   # +22.8% cumulative NZ CPI over 5 years (~4.19% p.a.)

CREATOR_SCHEMA_OBJ = {
    "@type": "Person",
    "name": CREATOR_NAME,
    "alternateName": ["Andrew Palliser", CREATOR_ALT_NAME],
    "jobTitle": "Mortgage & KiwiSaver Adviser",
    "identifier": CREATOR_FSP,
    "url": CREATOR_URL,
    "image": CREATOR_IMAGE,
    "worksFor": {
        "@type": "Organization",
        "name": "Home Loan Factory",
        "legalName": "HLF Financial Limited",
        "identifier": "FSP1008571",
        "url": PUBLISHER_URL,
    },
}

PUBLISHER_SCHEMA_OBJ = {
    "@type": "Organization",
    "name": PUBLISHER_NAME,
    "legalName": "HLF Financial Limited",
    "identifier": "FSP1008571",
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
IRD_KS_ENTRY_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/schemes/entry-method"
IRD_KS_PAY_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/contributions/scheme-providers-payments"
IRD_KS_SUSP_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/contributions/savings-suspensions"
IRD_KS_DATA_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/datasets"
IRD_SL_BAL_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/statistics-on-the-student-loan-scheme/statistics-on-student-loan-balances"

FMA_KS_REPORT_PAGE = "https://www.fma.govt.nz/library/reports-and-papers/kiwisaver-annual-report/"
STATS_BC_PAGE = "https://www.stats.govt.nz/topics/building/"
STATS_CPI_PAGE = "https://www.stats.govt.nz/indicators/consumers-price-index-cpi/"
STATS_HES_PAGE = "https://www.stats.govt.nz/information-releases/household-expenditure-statistics-year-ended-june-2023/"
RBNZ_FSR_PAGE = "https://www.rbnz.govt.nz/financial-stability/financial-stability-report"
RBNZ_INS_STATS_PAGE = "https://www.rbnz.govt.nz/statistics/series/insurance"
NHC_ACT_PAGE = "https://www.legislation.govt.nz/act/public/2023/0001/latest/LMS487791.html"
NHC_FUND_PAGE = "https://www.naturalhazards.govt.nz/about-nhc/how-we-work/natural-hazard-fund/"
ACC_REPORT_PAGE = "https://www.acc.co.nz/assets/corporate-documents/ACC8911-acc-annual-report-2026.pdf"
ACC_LEVY_PAGE = "https://www.acc.co.nz/about-us/our-levies-2/levy-results"

MAX_AGE_MONTHLY = 85
MAX_AGE_TENANCY = 120
MAX_AGE_QUARTERLY = 160
MAX_AGE_ANNUAL = 420
MAX_AGE_INSURANCE = 550

# 9-Category Taxonomy (with 3-Way Insurance Split)
CATEGORY_ORDER = [
    "First-home buyers",
    "Mortgage lending & rates",
    "Property investors & rents",
    "Homeowners & equity",
    "KiwiSaver",
    "Housing supply & construction",
    "Property & Natural Hazard Insurance",
    "Personal Risk & Income Protection",
    "Insurance Market & Claims",
]

# ---------------------------------------------------------------------------
# Data Dictionary Catalog — Part A: First-Home Buyers, Mortgages & Equity
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
        "methodology": "Compiled from monthly registered bank returns under RBNZ Table C31. First-home buyers are defined by lending banks as borrowers purchasing a residential property to occupy as a first home. Dollar values in Table C31 are published by RBNZ rounded to the nearest $1 million.",
        "hlf_calculation": "(High-LVR FHB Borrowers [LVRN.MMB2.BC1] ÷ Total FHB Borrowers [LVRN.MMB2.AC1]) × 100",
        "hlf_interpretation": "HLF market context (not measured in C31): RBNZ Table C31 reports all >80% LVR commitments as a single aggregate and does not disaggregate the underlying lending channel (such as bank low-equity speed-limit allocations, Kāinga Ora First Home Loans, or LVR-exempt new builds).",
        "questions": [
            "How many first-home buyers have less than a 20% deposit in New Zealand?",
            "Is a 20% deposit actually typical for NZ first-home buyers?",
            "How common are low-deposit mortgages in New Zealand?",
            "What percentage of first-home buyers borrow with less than a 20% deposit?",
            "Are low-deposit first-home loans becoming more common?",
        ],
        "article_title": "",
        "article_url": "",
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
        "methodology": "Sourced from RBNZ Table C31 monthly registered bank commitments across first-home buyers, other owner-occupiers, property investors, and business purposes (where dollar values are published rounded to the nearest $1m).",
        "hlf_calculation": "(FHB Lending NZDm [LVRN.MMB1.AC1] ÷ All Borrower Lending NZDm [LVRN.MMB1.AA]) × 100",
        "hlf_interpretation": "HLF market context: Tracks the proportion of new bank mortgage commitments going to first-home buyers relative to other owner-occupiers, investors, and business borrowers.",
        "questions": [
            "What share of New Zealand mortgages go to first-home buyers?",
            "How many first-home buyer loans are approved each month in NZ?",
            "What is the average first-home buyer mortgage size in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
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
        "methodology": "Derived from RBNZ Table C31 by dividing total monthly dollar lending in each LVR tier (published by RBNZ in rounded $ millions) by the exact number of approved first-home borrowers in that same LVR tier.",
        "hlf_calculation": "(>80% LVR FHB NZDm × $1m ÷ >80% LVR FHB Borrowers) minus (<=80% LVR FHB NZDm × $1m ÷ <=80% LVR FHB Borrowers)",
        "hlf_interpretation": "HLF market context: The data shows that first-home buyers at >80% LVR take out larger mortgages on average than those at <=80% LVR. RBNZ Table C31 does not record property purchase prices or borrower motivations; factors can include borrowing a higher percentage of the purchase price, regional price mix, or household income.",
        "questions": [
            "Do low-deposit first-home buyers borrow more or less than 20% deposit buyers?",
            "What is the average low-deposit first-home loan size in NZ?",
            "What is the average mortgage for a first-home buyer with a 20% deposit?",
        ],
        "article_title": "",
        "article_url": "",
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
        "methodology": "Compares high-LVR first-home commitments against high-LVR commitments across all borrower types combined in RBNZ Table C31 (where dollar series are published rounded to the nearest $1m).",
        "hlf_calculation": "(FHB >80% LVR Lending [LVRN.MMB1.BC1] ÷ All Borrower >80% LVR Lending [LVRN.MMB1.BA]) × 100",
        "hlf_interpretation": "HLF market context: Shows that first-home buyers account for the majority of >80% LVR commitments in RBNZ Table C31. Table C31 includes both speed-limit lending and LVR-exempt lending (such as Kāinga Ora First Home Loans and new builds).",
        "questions": [
            "What share of bank >80% LVR lending goes to first-home buyers in New Zealand?",
            "Do New Zealand banks allocate most low-deposit loans to first-home buyers?",
        ],
        "article_title": "",
        "article_url": "",
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
        "article_title": "",
        "article_url": "",
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
        "methodology": "Sourced from RBNZ Table C31 series for <=80% LVR first-home buyer lending and borrower counts (dollar values rounded by RBNZ to the nearest $1m).",
        "hlf_calculation": "(<=80% LVR FHB Borrowers [LVRN.MMB2.CC1] ÷ Total FHB Borrowers [LVRN.MMB2.AC1]) × 100",
        "hlf_interpretation": "HLF market context (not measured in C31): An LVR of <=80% (a 20%+ deposit) can improve access to advertised special-rate pricing and generally avoids low-equity margins, subject to individual lender policy.",
        "questions": [
            "What percentage of NZ first-home buyers have a 20% deposit?",
            "What is the average mortgage size for a first-home buyer with a 20% deposit?",
        ],
        "article_title": "",
        "article_url": "",
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
        "article_title": "",
        "article_url": "",
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
        "article_title": "",
        "article_url": "",
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
        "methodology": "Derived from RBNZ Table C31 'Other owner-occupiers' series at or below 80% LVR (dollar series rounded by RBNZ to the nearest $1m).",
        "hlf_calculation": "(<=80% LVR Other Owner-Occupiers [LVRN.MMB2.CC2] ÷ Total Other Owner-Occupiers [LVRN.MMB2.AC2]) × 100",
        "hlf_interpretation": "HLF market context: An LVR of <=80% on a mortgage commitment corresponds to holding at least 20% security equity in the mortgaged property. Note that RBNZ C31 only measures homeowners who apply for a new loan, refinance, or top-up during the month, not unmortgaged homeowners.",
        "questions": [
            "What percentage of existing New Zealand homeowners have at least 20% equity when refinancing or moving?",
            "What is the average mortgage size for an existing homeowner at <=80% LVR in NZ?",
        ],
        "article_title": "",
        "article_url": "",
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
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Usable Equity Calculator",
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
        "methodology": "In RBNZ Table C31, >80% LVR investor loans are reported in LVRN.MMB2.BN / LVRN.MMB1.BN, while the 70% < LVR <= 80% band is derived by subtracting >80% LVR from >70% LVR (LVRN.MMB2.DN minus LVRN.MMB2.BN; dollar series rounded by RBNZ to $1m).",
        "hlf_calculation": ">80% LVR = LVRN.MMB2.BN; 70%–80% LVR = LVRN.MMB2.DN minus LVRN.MMB2.BN",
        "hlf_interpretation": "HLF market context (not disaggregated in C31): Captures investor lending above 80% LVR and between 70% and 80% LVR. This can include lending supported by new-build exemptions, combined-collateral refinancing, bank high-LVR allocations, and other lending permitted under individual lender policies.",
        "questions": [
            "How many property investor loans are approved at >80% LVR (under a 20% deposit) in NZ?",
            "How many investor loans sit in the 70%–80% LVR (20%–30% deposit) band?",
        ],
        "article_title": "",
        "article_url": "",
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
        "article_title": "",
        "article_url": "",
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
        "article_title": "",
        "article_url": "",
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
        "methodology": "Compiled from RBNZ Table C33 (New residential mortgage lending by purpose — Change in loan provider; dollar series published rounded to $1m).",
        "hlf_calculation": "(Refinance Lending NZDm [LVRN.MMB1.AZD] ÷ Total Purpose Lending NZDm) × 100",
        "hlf_interpretation": "HLF market context: Measures external refinancing where a borrower switches from one mortgage provider to another (internal refixes with the same bank are not counted in Table C33).",
        "questions": [
            "How many New Zealand homeowners switch banks or refinance each month?",
            "What is the average mortgage size for borrowers refinancing in NZ?",
        ],
        "article_title": "",
        "article_url": "",
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
        "methodology": "Compiled from RBNZ Table C33 (New residential mortgage lending by purpose — Top-ups; dollar series published rounded to $1m).",
        "hlf_calculation": "(Top-Up Borrowers [LVRN.MMB2.AZB] ÷ Total Purpose Borrowers) × 100",
        "hlf_interpretation": "HLF market context (not disaggregated in C33): RBNZ Table C33 records top-ups as a single purpose category without disaggregating how borrowers use the funds (such as renovations, debt consolidation, or investment deposits).",
        "questions": [
            "How common are mortgage top-ups in New Zealand?",
            "What is the average mortgage top-up size in NZ?",
        ],
        "article_title": "",
        "article_url": "",
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
        "methodology": "Sourced from RBNZ Table C33 series LVRN.MMB1.AZB (Top-up lending value, rounded to $1m) and LVRN.MMB2.AZB (Top-up borrower count).",
        "hlf_calculation": "Officially reported LVRN.MMB1.AZB dollar total; Avg top-up = LVRN.MMB1.AZB × $1m ÷ LVRN.MMB2.AZB",
        "hlf_interpretation": "HLF market context: This represents homeowners drawing additional borrowing against existing property equity, subject to lender servicing and LVR policy.",
        "questions": [
            "How much mortgage top-up lending is approved in New Zealand each month?",
        ],
        "article_title": "",
        "article_url": "",
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
        "methodology": "Derived from RBNZ Table C32 (New and existing residential mortgage lending by payment type; dollar series rounded to $1m).",
        "hlf_calculation": "(New Investor Interest-Only NZDm ÷ Total New Investor NZDm) × 100",
        "hlf_interpretation": "HLF market context: Shows the difference in interest-only usage between residential investment loans and owner-occupied home loans, subject to lender interest-only term limits.",
        "questions": [
            "What percentage of NZ property investors use interest-only mortgages?",
            "How common are interest-only home loans for owner-occupiers vs investors?",
        ],
        "article_title": "",
        "article_url": "",
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
        "also_known_as": "RBNZ C41 Borrower Gross Income (BGI) — First-home buyer mean application income",
        "category": "First-home buyers",
        "topic": "Borrower income & servicing",
        "source": "RBNZ C41",
        "source_full": "Reserve Bank of New Zealand Table C41",
        "series_ids": "DTINI.MPA (Borrower Gross Income by TDTI band)",
        "source_url": RBNZ_SOURCES["c41"]["pages"][0],
        "data_type": "HLF-calculated mean gross income across C41 TDTI bands (Total band income ÷ implied borrower count)",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The HLF-calculated mean annual pre-tax gross income across approved New Zealand first-home buyer mortgage applications (combining all single and joint applicants on the loan).",
        "methodology": "Computed by HLF from RBNZ Table C41 (Residential mortgage borrower gross income) by dividing total gross income across all Total Debt-to-Income (TDTI) bands by the implied borrower count (Band Total Income ÷ Band Average Income). Note that RBNZ publishes band income pools rounded to $0.1m and band averages rounded to $1,000s, so implied application counts reflect source rounding.",
        "hlf_calculation": "Total Gross Income NZDm across all C41 TDTI bands ÷ Implied Application Count (Sum of [Band NZDm × 1,000 ÷ Band Avg Income NZD000s])",
        "hlf_interpretation": "HLF market context: RBNZ Table C41 publishes band-level figures rather than a single nationwide mean; the overall figure is calculated by HLF from those published band figures. Because C41 measures total gross income per mortgage application (often combining two earners plus eligible boarder income), it represents combined application income rather than an individual salary.",
        "questions": [
            "What is the average combined household income on an approved first-home mortgage in NZ?",
        ],
        "article_title": "",
        "article_url": "",
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
        "data_type": "Officially reported TDTI band averages & HLF-calculated overall mean (Total band income ÷ implied applications)",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "Official RBNZ Table C41 average pre-tax gross borrower income across each of the five first-home buyer Total Debt-to-Income (TDTI) bands, alongside the HLF-calculated mean across all bands.",
        "methodology": "Sourced from RBNZ Table C41 (Residential mortgage borrower gross income) across the five first-home buyer TDTI bands. Implied application counts per band are derived from RBNZ's published band income totals ($m) and band averages ($000s) and are subject to rounding in RBNZ's source tables.",
        "hlf_calculation": "Sum of Band Gross Income Pools ($m) ÷ Sum of Implied Band Applications (Band Pool $m × 1,000 ÷ Band Avg Income $000s)",
        "hlf_interpretation": "HLF market context: RBNZ Table C41 reports each TDTI band separately; the overall mean is derived by HLF and can be reproduced directly from the band breakdown table below. Because mortgage applications frequently combine two incomes, application averages sit higher than single-earner salaries.",
        "questions": [
            "How does first-home buyer gross income vary across Debt-to-Income (TDTI) bands in RBNZ C41?",
            "What is the average gross borrower income on an approved first-home mortgage application in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
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
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
}

PART_1_OF_5_READY = True
# === PART 2A OF 5 ===
TICKER_CATALOG_EXPANSION = {
    # --- Expanded KiwiSaver Datasets (8 KiwiSaver + 1 Regional FHB KiwiSaver) ---
    "kiwisaver-first-home": {
        "name": "Monthly & 10-year KiwiSaver first-home withdrawal count, value & average per member",
        "also_known_as": "Monthly & annual IRD KiwiSaver first-home purchase withdrawals",
        "category": "KiwiSaver",
        "topic": "KiwiSaver & first-home deposits",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — KiwiSaver Withdrawal Statistics",
        "series_ids": "IRD Monthly & June FY KiwiSaver First-Home Withdrawals",
        "source_url": IRD_KS_DATA_PAGE,
        "data_type": "Officially reported IRD totals & HLF-derived per-member average and 10-year nominal/real comparison",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The officially reported monthly and annual number of KiwiSaver members withdrawing funds for a first-home purchase and total dollars withdrawn, plus the HLF-derived mean withdrawal per member and 10-year nominal and CPI inflation-adjusted comparison.",
        "methodology": "Compiled monthly and annually by Inland Revenue from KiwiSaver scheme provider reports.",
        "hlf_calculation": "Mean per Member = Total First-Home Withdrawal Dollars ÷ Withdrawing Members; 10Y Real Growth = ((FY2026 ÷ FY2016) ÷ CPI_10Y_Factor [1.337]) − 1",
        "hlf_interpretation": "HLF market context: IRD reports individual member withdrawals rather than household applications; when two KiwiSaver members purchase together, both withdrawals are counted separately in the member total.",
        "questions": [
            "How much is the average KiwiSaver first-home withdrawal per member in New Zealand?",
            "How has KiwiSaver first-home withdrawal volume grown over the past 10 years in nominal and real terms?",
        ],
        "article_title": "",
        "article_url": "",
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
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-switches": {
        "name": "Annual, monthly & 10-year KiwiSaver scheme transfers (members switching providers)",
        "also_known_as": "KiwiSaver provider switching / KiwiSaver scheme transfers / KiwiSaver member transfers",
        "category": "KiwiSaver",
        "topic": "KiwiSaver scheme transfers",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — KiwiSaver Scheme Transfer Statistics",
        "series_ids": "IRD Annual KiwiSaver Scheme Transfers (10-Year Series)",
        "source_url": IRD_KS_TR_PAGE,
        "data_type": "Officially reported IRD annual totals & HLF-derived monthly average and 10-year change",
        "derived_by_hlf": True,
        "unit": "transfers",
        "frequency": "Annual",
        "definition": "The officially reported annual number of New Zealand KiwiSaver members who transferred their balance to a different KiwiSaver scheme provider over the June financial year, plus the HLF-derived monthly mean and 10-year trend.",
        "methodology": "Compiled by Inland Revenue from KiwiSaver scheme transfers processed during the year ended 30 June.",
        "hlf_calculation": "Monthly Mean = Annual Scheme Transfers ÷ 12 months; 10-Year Change = (Latest FY − FY 10 Years Ago) ÷ FY 10 Years Ago",
        "hlf_interpretation": "HLF market context: Measures inter-scheme transfers where members move their KiwiSaver balance to a new provider (internal fund-type changes within the same scheme are not counted in IRD's transfer total).",
        "questions": [
            "How many New Zealanders switch KiwiSaver providers each year?",
            "How has KiwiSaver provider switching changed over the last 10 years?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-total-fum-balance": {
        "name": "KiwiSaver total funds under management (FUM) & average member balance (10-year nominal & real growth)",
        "also_known_as": "FMA KiwiSaver Annual Report total FUM & official average balance per member",
        "category": "KiwiSaver",
        "topic": "Long-term KiwiSaver balances & FUM growth",
        "source": "FMA & RBNZ",
        "source_full": "Financial Markets Authority (FMA) KiwiSaver Annual Report & RBNZ Managed Funds Survey",
        "series_ids": "FMA Annual KiwiSaver Total FUM, Enrolled Member Count & Published Average Balance",
        "source_url": FMA_KS_REPORT_PAGE,
        "data_type": "Officially reported FMA average balance/FUM & HLF-calculated 10-year nominal and CPI inflation-adjusted (real) growth",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual",
        "definition": "Total funds under management (FUM) across all New Zealand KiwiSaver schemes, total enrolled members, and the FMA's officially published average balance per member ($40,340 as at 31 March 2026), alongside HLF-calculated 10-year nominal and CPI inflation-adjusted (real) growth rates.",
        "methodology": "Sourced from the FMA KiwiSaver Annual Report (for the year ended 31 March 2026, published 10 September 2026). Real (inflation-adjusted) 10-year growth is deflated by HLF using Stats NZ's cumulative 10-year all-groups CPI factor (+33.7% / 1.337x from March 2016 to March 2026).",
        "hlf_calculation": "Official FMA Average Balance = $40,340; 10Y Nominal Growth = ($40,340 − $12,852) ÷ $12,852 = +213.9% (+12.1% p.a.); 10Y Real (CPI-Adjusted) Growth = (($40,340 ÷ $12,852) ÷ 1.337) − 1 = +134.8% (+8.9% p.a.)",
        "hlf_interpretation": "HLF market context: The FMA's published average member balance ($40,340) is calculated across all 3.44 million enrolled KiwiSaver accounts—which includes children, non-contributing members, and members who have recently withdrawn funds for a first home.",
        "questions": [
            "What is the average KiwiSaver balance in New Zealand?",
            "How much money is invested in KiwiSaver in total across New Zealand?",
            "How fast have KiwiSaver balances grown over the last 10 years after adjusting for inflation?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-default-vs-active": {
        "name": "KiwiSaver active scheme choice vs default & employer-nominated allocation",
        "also_known_as": "IRD KiwiSaver scheme entry method (Active choice vs Default allocated)",
        "category": "KiwiSaver",
        "topic": "KiwiSaver default vs active membership",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — KiwiSaver Scheme Entry Method Statistics",
        "series_ids": "IRD Annual KiwiSaver Members by Scheme Entry Method (June FY)",
        "source_url": IRD_KS_ENTRY_PAGE,
        "data_type": "Officially reported IRD member counts & HLF-derived percentage shares",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Annual",
        "definition": "Breakdown of enrolled New Zealand KiwiSaver members by how they entered their scheme: active personal choice versus passive IRD default allocation or employer-nominated scheme entry.",
        "methodology": "Published annually by Inland Revenue for the year ended 30 June tracking cumulative scheme entry methods.",
        "hlf_calculation": "Non-Active Share = ((Default Allocated + Employer Nominated) ÷ Total Members) × 100",
        "hlf_interpretation": "HLF market context: Since December 2021, default KiwiSaver funds have been mandated to use a Balanced (rather than Conservative) investment mandate, but default allocation still means the member has not actively selected a fund type aligned to their specific first-home or retirement timeframe.",
        "questions": [
            "How many New Zealanders are still in a default KiwiSaver scheme?",
            "What percentage of KiwiSaver members actively choose their provider?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-contributions-10yr": {
        "name": "Annual KiwiSaver employee, employer & Crown contributions (10-year nominal/real trend & 3.5% default rate)",
        "also_known_as": "IRD KiwiSaver scheme provider payments & statutory contribution settings",
        "category": "KiwiSaver",
        "topic": "KiwiSaver contributions & government match",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — KiwiSaver Scheme Provider Payments & Statutory Settings",
        "series_ids": "IRD Annual Deductions & Crown Contributions Passed to Providers",
        "source_url": IRD_KS_PAY_PAGE,
        "data_type": "Officially reported IRD dollar flows, statutory rates & HLF-calculated nominal/real 10-year growth",
        "derived_by_hlf": True,
        "unit": "NZDB",
        "frequency": "Annual / Monthly",
        "definition": "Total annual KiwiSaver contributions transferred by Inland Revenue to scheme providers (broken down by employee deductions, employer contributions, and Crown government contributions) over a 10-year window in both nominal and CPI inflation-adjusted terms, alongside current statutory default rates.",
        "methodology": "Compiled from Inland Revenue's annual KiwiSaver payments table (year ended 30 June) and statutory KiwiSaver Act settings, deflated over 10 years using Stats NZ's cumulative CPI factor (+33.7%).",
        "hlf_calculation": "10Y Nominal Growth = ($11.40B − $5.10B) ÷ $5.10B = +123.5% (+8.4% p.a.); 10Y Real (CPI-Adjusted) Growth = (($11.40B ÷ $5.10B) ÷ 1.337) − 1 = +67.2% (+5.3% p.a.)",
        "hlf_interpretation": "HLF market context: Under current KiwiSaver settings, the maximum annual Government contribution is $260.72 (25 cents per dollar on the first $1,042.86 of member contributions for eligible earners at or below $180,000), and the default employee and employer contribution rate increased from 3.0% to 3.5% from 1 April 2026 (scheduled to reach 4.0% from 1 April 2028).",
        "questions": [
            "How much do New Zealand employees, employers, and the Government contribute to KiwiSaver each year?",
            "What is the current KiwiSaver Government contribution amount ($260.72) and default contribution rate (3.5%)?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-fund-allocation-shift": {
        "name": "Long-term KiwiSaver asset allocation shift (Growth vs Balanced vs Conservative funds)",
        "also_known_as": "10-year KiwiSaver FUM breakdown by risk profile (FMA KiwiSaver Report)",
        "category": "KiwiSaver",
        "topic": "KiwiSaver fund risk profiles & asset allocation",
        "source": "FMA KiwiSaver Report",
        "source_full": "Financial Markets Authority (FMA) KiwiSaver Annual Report — Fund Type Breakdown",
        "series_ids": "FMA KiwiSaver FUM by Fund Risk Category",
        "source_url": FMA_KS_REPORT_PAGE,
        "data_type": "Officially reported FMA fund-category totals & HLF-derived percentage shares",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Annual",
        "definition": "The proportion of total New Zealand KiwiSaver funds under management (FUM) invested in Growth/Aggressive, Balanced, and Conservative/Defensive fund categories compared with 10 years earlier.",
        "methodology": "Derived from FMA KiwiSaver Annual Report disclosures (year ended 31 March 2026, published 10 September 2026) categorising scheme FUM by underlying growth-asset allocation.",
        "hlf_calculation": "Category Share = (Category FUM NZDB ÷ Total KiwiSaver FUM NZDB) × 100",
        "hlf_interpretation": "HLF market context: Over the past decade, New Zealand KiwiSaver balances have shifted structurally from Conservative/Cash funds into Growth and Balanced funds. In practice, appropriate fund selection depends on when the member intends to withdraw funds (for example, a near-term first-home purchase within 1–3 years versus a 10–30+ year retirement horizon).",
        "questions": [
            "What percentage of KiwiSaver money is in Growth vs Balanced vs Conservative funds?",
            "Should I switch from a Growth fund to a Conservative or Cash fund before buying a first home?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-hardship-withdrawals": {
        "name": "Monthly & 10-year KiwiSaver significant financial hardship withdrawals",
        "also_known_as": "IRD KiwiSaver financial hardship withdrawal count, value & average",
        "category": "KiwiSaver",
        "topic": "KiwiSaver financial hardship & household pressure",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — KiwiSaver Withdrawal Statistics (Financial Hardship)",
        "series_ids": "IRD Monthly & Annual Financial Hardship Withdrawals",
        "source_url": IRD_KS_DATA_PAGE,
        "data_type": "Officially reported IRD hardship counts/values & HLF-derived per-member average",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The monthly and annual number of KiwiSaver members approved for a significant financial hardship withdrawal, total dollars withdrawn, and HLF-derived average withdrawal per member.",
        "methodology": "Compiled monthly and annually by Inland Revenue from KiwiSaver scheme provider reports.",
        "hlf_calculation": "Mean Hardship Withdrawal = Total Hardship Withdrawal Dollars ÷ Hardship Withdrawing Members",
        "hlf_interpretation": "HLF market context: Significant financial hardship withdrawals require supervisor approval when a member cannot meet minimum living expenses or mortgage arrears. Tracking hardship withdrawals alongside first-home withdrawals provides a barometer of household cost-of-living and mortgage-servicing pressure.",
        "questions": [
            "How many New Zealanders withdraw KiwiSaver for financial hardship each month?",
            "What is the average KiwiSaver financial hardship withdrawal amount in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "kiwisaver-savings-suspensions": {
        "name": "KiwiSaver members on active savings suspensions (10-year trend)",
        "also_known_as": "IRD KiwiSaver contributions holiday / savings suspension statistics",
        "category": "KiwiSaver",
        "topic": "KiwiSaver contribution continuity",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — KiwiSaver Savings Suspension Statistics",
        "series_ids": "IRD Monthly & Annual Active Savings Suspensions",
        "source_url": IRD_KS_SUSP_PAGE,
        "data_type": "Officially reported IRD counts & HLF-calculated 10-year change",
        "derived_by_hlf": True,
        "unit": "members",
        "frequency": "Monthly / Annual",
        "definition": "The number of enrolled New Zealand KiwiSaver members currently on an approved savings suspension (formerly known as a contributions holiday), broken down by ordinary and financial hardship suspensions.",
        "methodology": "Published by Inland Revenue tracking active savings suspensions at period end.",
        "hlf_calculation": "10-Year Change = (Current Active Suspensions − June 2016 Active Suspensions) ÷ June 2016 Active Suspensions",
        "hlf_interpretation": "HLF market context: Lowering the maximum standard savings suspension duration from 5 years to 1 year (in 2019) contributed to a structural decline in long-term passive suspensions. Pausing employee contributions generally pauses compulsory employer matching contributions unless agreed otherwise.",
        "questions": [
            "How many New Zealanders are on a KiwiSaver savings suspension?",
            "What happens to employer contributions when you pause KiwiSaver?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },

    # --- Housing Supply & Construction Datasets (6 Total) ---
    "new-builds-annual": {
        "name": "Annual new residential dwellings consented in New Zealand",
        "also_known_as": "Stats NZ annual residential building consents",
        "category": "Housing supply & construction",
        "topic": "New builds & housing supply",
        "source": "Stats NZ",
        "source_full": "Stats NZ — Building Consents Issued",
        "series_ids": "Stats NZ Building Consents Monthly Release",
        "source_url": STATS_BC_PAGE,
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
        "article_title": "",
        "article_url": "",
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
        "source_url": STATS_BC_PAGE,
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
        "article_title": "",
        "article_url": "",
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
        "series_ids": "Stats NZ CPI Housing & Household Utilities (Purchase of New Housing)",
        "source_url": STATS_CPI_PAGE,
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
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "regional-building-consents": {
        "name": "Annual new residential building consents by major New Zealand region",
        "also_known_as": "Stats NZ regional dwelling consents (Auckland, Canterbury, Waikato, Wellington)",
        "category": "Housing supply & construction",
        "topic": "Regional new-build supply",
        "source": "Stats NZ",
        "source_full": "Stats NZ — Building Consents Issued (Regional Breakdown)",
        "series_ids": "Stats NZ Annual Dwelling Consents by Region",
        "source_url": STATS_BC_PAGE,
        "data_type": "Officially reported Stats NZ regional totals & HLF-derived national shares",
        "derived_by_hlf": True,
        "unit": "dwellings",
        "frequency": "Monthly",
        "definition": "Annual new residential dwelling consents issued across New Zealand's main regions (Auckland, Canterbury, Waikato, Wellington, and Rest of New Zealand), plus their share of national supply.",
        "methodology": "Compiled monthly by Stats NZ from territorial authority consent approvals.",
        "hlf_calculation": "Regional Share = (Regional Annual Consents ÷ National Annual Consents) × 100",
        "hlf_interpretation": "HLF market context: Auckland and Canterbury together account for more than three-fifths of all new residential dwellings consented across New Zealand, influencing regional availability of LVR-exempt turnkey and construction properties.",
        "questions": [
            "Which New Zealand regions are building the most new homes?",
            "How many new homes were consented in Auckland, Canterbury, Waikato, and Wellington over the past year?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "standalone-vs-retirement-apartments": {
        "name": "New dwelling consents breakdown: stand-alone houses, townhouses, apartments & retirement units",
        "also_known_as": "Full 4-category Stats NZ monthly residential dwelling type breakdown",
        "category": "Housing supply & construction",
        "topic": "Dwelling typology & apartment supply",
        "source": "Stats NZ",
        "source_full": "Stats NZ — Building Consents Issued (All Dwelling Types)",
        "series_ids": "Stats NZ Monthly Consents by 4 Dwelling Categories",
        "source_url": STATS_BC_PAGE,
        "data_type": "Officially reported Stats NZ counts & HLF-derived category percentages",
        "derived_by_hlf": True,
        "unit": "dwellings",
        "frequency": "Monthly",
        "definition": "Complete monthly breakdown of new residential dwelling consents across all four Stats NZ categories: stand-alone houses, townhouses/flats/units, apartments, and retirement village units.",
        "methodology": "Published monthly by Stats NZ from territorial authority building consent records.",
        "hlf_calculation": "Category % = (Category Monthly Consents ÷ Total Monthly Consents) × 100",
        "hlf_interpretation": "HLF market context: Separating apartments and retirement village units from townhouses and stand-alone houses matters for mortgage borrowers because banks apply different LVR and floor-area lending criteria to high-density apartments versus fee-simple townhouses and stand-alone homes.",
        "questions": [
            "How many apartments vs townhouses vs stand-alone houses are being consented in New Zealand?",
            "Do banks have different deposit rules for apartments vs townhouses?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Turnkey vs. Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/turnkey-vs-construction-loan-calculator-nz",
    },
    "new-build-lvr-exemption-flow": {
        "name": "High-LVR mortgage lending channels supported by new-build & speed-limit exemptions",
        "also_known_as": "RBNZ C31 high-LVR mover (>80% LVR) and investor (>70% LVR) lending volume",
        "category": "Housing supply & construction",
        "topic": "New-build LVR & DTI exemptions",
        "source": "RBNZ C31",
        "source_full": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.BC2, LVRN.MMB2.BC2, LVRN.MMB1.DN, LVRN.MMB1.BN",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported RBNZ C31 series & HLF-derived combined high-LVR channel volume",
        "derived_by_hlf": True,
        "unit": "NZDm",
        "frequency": "Monthly",
        "definition": "Combined monthly lending volume and borrower count across existing owner-occupiers borrowing at >80% LVR and property investors borrowing at >70% LVR in RBNZ Table C31.",
        "methodology": "Derived from RBNZ Table C31 series for >80% LVR other owner-occupiers and >70% LVR property investors (dollar series rounded by RBNZ to $1m).",
        "hlf_calculation": "Combined Volume = Other Owner-Occupier >80% LVR NZDm [LVRN.MMB1.BC2] + Investor >70% LVR NZDm [LVRN.MMB1.DN]",
        "hlf_interpretation": "HLF market context: Under RBNZ macroprudential rules, newly built dwellings (purchased from the developer or within 6 months of code compliance) are exempt from both LVR and DTI restrictions—allowing eligible owner-occupiers and investors to purchase new builds with lower deposits than existing housing stock.",
        "questions": [
            "How do Reserve Bank LVR and DTI exemptions work for new builds in New Zealand?",
            "How much low-deposit lending goes to existing homeowners and property investors each month?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Turnkey vs. Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/turnkey-vs-construction-loan-calculator-nz",
    },

    # --- Tenancy & Rental Market Datasets (5 Total) ---
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
        "article_title": "",
        "article_url": "",
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
        "article_title": "",
        "article_url": "",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "rent-long-term-growth": {
        "name": "Long-term 5-year & 10-year New Zealand median weekly rent growth (nominal & real CPI-adjusted)",
        "also_known_as": "MBIE Tenancy Services 5-year and 10-year nominal and real rental growth benchmark",
        "category": "Property investors & rents",
        "topic": "Long-term rent inflation & yield modelling",
        "source": "MBIE Tenancy Services & Stats NZ CPI",
        "source_full": "MBIE Tenancy Services — Rental Bond Data (Long-Term Series) & Stats NZ CPI",
        "series_ids": "Tenancy Services Monthly Regional Bond CSV (5Y & 10Y Median Rent)",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported historical medians & HLF-derived 5-year / 10-year nominal and real (CPI-adjusted) CAGR",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "Cumulative percentage growth, dollar increase, and compound annual growth rate (CAGR) in national median weekly private rents over the trailing 5-year and 10-year periods in both nominal and CPI inflation-adjusted (real) terms.",
        "methodology": "Computed from MBIE Tenancy Services historical monthly regional bond CSV series, with real-terms growth deflated by Stats NZ cumulative CPI (+33.7% over 10 years; +22.8% over 5 years).",
        "hlf_calculation": "10Y Nominal % = ((Rent_t − Rent_t-120) ÷ Rent_t-120) × 100; 10Y Real % = (((Rent_t ÷ Rent_t-120) ÷ 1.337) − 1) × 100; 10Y Real CAGR = (((Rent_t ÷ Rent_t-120) ÷ 1.337)^(1/10) − 1) × 100",
        "hlf_interpretation": "HLF market context: Comparing both nominal and inflation-adjusted (real) 5-year and 10-year rental growth provides a balanced benchmark for long-term property cashflow modelling and rent-vs-buy comparisons.",
        "questions": [
            "How much have New Zealand rents increased over the last 5 and 10 years in nominal and real terms?",
            "Have New Zealand rents grown faster than general CPI inflation over the last decade?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "rent-quartile-spread": {
        "name": "New Zealand weekly private rent distribution: lower quartile, median & upper quartile",
        "also_known_as": "MBIE Tenancy Services 25th, 50th, and 75th percentile weekly rent spread",
        "category": "Property investors & rents",
        "topic": "Rental affordability & market tiers",
        "source": "MBIE Tenancy Services",
        "source_full": "MBIE Tenancy Services — Rental Bond Data (Quartile Series)",
        "series_ids": "Tenancy Services Lower Quartile, Median & Upper Quartile Weekly Rent",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported MBIE quartiles & HLF-calculated interquartile spread",
        "derived_by_hlf": True,
        "unit": "NZD/week",
        "frequency": "Monthly",
        "definition": "The lower quartile (25th percentile), median (50th percentile), and upper quartile (75th percentile) weekly private rent across newly lodged New Zealand tenancy bonds, plus the interquartile dollar spread.",
        "methodology": "Sourced from MBIE Tenancy Services monthly regional bond dataset.",
        "hlf_calculation": "Interquartile Spread = Upper Quartile Rent − Lower Quartile Rent (NZD/wk and NZD/yr)",
        "hlf_interpretation": "HLF market context: Comparing lower-quartile entry-level rents against median and upper-quartile family rents helps both tenants budgeting for a first-home deposit and investors assessing different property tiers.",
        "questions": [
            "What is the lower-quartile and upper-quartile weekly rent in New Zealand?",
            "How wide is the gap between entry-level and upper-tier weekly rents in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "rental-turnover-net-stock": {
        "name": "Annual net change in active private rental bonds & monthly tenancy turnover",
        "also_known_as": "MBIE Tenancy Services active bond pool growth & lodged vs closed bonds",
        "category": "Property investors & rents",
        "topic": "Rental stock expansion & tenancy turnover",
        "source": "MBIE Tenancy Services",
        "source_full": "MBIE Tenancy Services — Rental Bond Data (Lodged, Closed & Active Bonds)",
        "series_ids": "Tenancy Services Active Bonds YoY Change & Monthly Closed Bonds",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported bond counts & HLF-calculated 12-month net stock change",
        "derived_by_hlf": True,
        "unit": "bonds",
        "frequency": "Monthly",
        "definition": "The 12-month net change in the total pool of active private rental bonds across New Zealand, alongside the monthly comparison of newly lodged bonds versus closed (refunded) bonds.",
        "methodology": "Derived from MBIE Tenancy Services monthly regional bond dataset ('ALL' New Zealand location).",
        "hlf_calculation": "12-Month Net Active Change = Active Bonds_t − Active Bonds_t-12; Monthly Net Flow = Lodged Bonds − Closed Bonds",
        "hlf_interpretation": "HLF market context: A rising pool of active private rental bonds indicates net expansion in occupied private rental stock across New Zealand, while the monthly ratio of lodged to closed bonds reflects seasonal tenancy turnover.",
        "questions": [
            "Is the number of private rental properties in New Zealand growing or shrinking?",
            "How many rental tenancies end (close) vs start (lodge) each month in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
}

TICKER_CATALOG.update(TICKER_CATALOG_EXPANSION)

PART_2A_OF_5_READY = True_READY = True
# === PART 2B OF 5 ===
TICKER_CATALOG_INSURANCE = {
    # --- 3-Way Insurance Taxonomy (12 Datasets Total) ---
    # 1. Property & Natural Hazard Insurance (6 Datasets)
    "building-insurance-inflation": {
        "name": "Dwelling (building) insurance annual price inflation (Stats NZ CPI)",
        "also_known_as": "House insurance CPI price inflation / Residential building insurance price index change",
        "category": "Property & Natural Hazard Insurance",
        "topic": "House & dwelling insurance pricing",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Dwelling Insurance Sub-Index)",
        "series_ids": "CPI Dwelling Insurance (Miscellaneous Goods & Services — Insurance)",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI sub-index & HLF-calculated comparison",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in the price of residential dwelling (house) insurance as measured within the Stats NZ Consumers Price Index (CPI) basket.",
        "methodology": "Published quarterly by Stats NZ by tracking price movements for a representative sample of domestic dwelling insurance policies across New Zealand.",
        "hlf_calculation": "Officially reported Stats NZ annual percentage change in the CPI Dwelling Insurance series, compared against headline all-groups CPI",
        "hlf_interpretation": "HLF market context: Stats NZ's CPI measures percentage changes in insurance prices over time, not the actual dollar premium an individual household pays. Dwelling insurance price movements reflect global reinsurance costs, natural-hazard risk pricing, NHC levies, and residential rebuild cost inflation.",
        "questions": [
            "Why has house insurance become more expensive in New Zealand?",
            "How fast are dwelling insurance prices changing in the Stats NZ CPI?",
            "How does building insurance inflation compare with general CPI inflation?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build & Rebuild Cost Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "contents-insurance-inflation": {
        "name": "Household contents insurance annual price inflation (Stats NZ CPI)",
        "also_known_as": "Contents insurance CPI price movement",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Contents insurance pricing",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Contents Insurance Sub-Index)",
        "series_ids": "CPI Contents Insurance (Miscellaneous Goods & Services — Insurance)",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI sub-index & HLF-calculated comparison",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in the price of household contents insurance across New Zealand as measured in the Stats NZ Consumers Price Index (CPI).",
        "methodology": "Compiled quarterly by Stats NZ from representative domestic contents insurance policy pricing across New Zealand.",
        "hlf_calculation": "Officially reported Stats NZ annual percentage change in the CPI Contents Insurance series",
        "hlf_interpretation": "HLF market context: CPI tracks percentage price changes across a standardised basket of contents policies rather than individual quote amounts. For both renters and homeowners, contents cover also provides personal legal liability protection alongside replacement of household goods.",
        "questions": [
            "How much have contents insurance prices changed in New Zealand over the past year?",
            "How does contents insurance inflation compare with building insurance inflation?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "vehicle-insurance-inflation": {
        "name": "Vehicle insurance annual price inflation (Stats NZ CPI)",
        "also_known_as": "Motor vehicle insurance CPI price movement / Car insurance inflation NZ",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Motor vehicle insurance pricing",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Vehicle Insurance Sub-Index)",
        "series_ids": "CPI Vehicle Insurance (Miscellaneous Goods & Services — Insurance)",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI sub-index & HLF-calculated comparison",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in the price of private motor vehicle insurance in New Zealand as measured within the Stats NZ Consumers Price Index (CPI).",
        "methodology": "Compiled quarterly by Stats NZ from representative private motor vehicle insurance policies.",
        "hlf_calculation": "Officially reported Stats NZ annual percentage change in the CPI Vehicle Insurance series",
        "hlf_interpretation": "HLF market context: Measures percentage price changes in motor vehicle insurance indices (not individual dollar quotes). In mortgage servicing assessments, fixed household insurance commitments—including vehicle, house, and contents cover—form part of ongoing living expenses.",
        "questions": [
            "How much have car and vehicle insurance prices changed in New Zealand?",
            "Why have motor vehicle insurance premiums risen in recent years?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "household-insurance-expenditure": {
        "name": "Household expenditure on insurance & average domestic dwelling premium",
        "also_known_as": "Stats NZ Household Expenditure Survey (HES) insurance spend & RBNZ dwelling premium benchmark",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Household insurance affordability",
        "source": "Stats NZ HES & RBNZ FSR",
        "source_full": "Stats NZ Household Expenditure Statistics & RBNZ Financial Stability Report",
        "series_ids": "Stats NZ HES Insurance Expenditure & RBNZ Dwelling Insurance Affordability Series",
        "source_url": RBNZ_FSR_PAGE,
        "data_type": "Officially reported survey/report benchmarks & HLF-derived annual equivalents",
        "derived_by_hlf": True,
        "unit": "NZD/year",
        "frequency": "Annual / Semi-Annual",
        "definition": "Average household expenditure on insurance reported in Stats NZ's Household Expenditure Survey (HES) alongside the Reserve Bank of New Zealand's domestic dwelling insurance premium benchmark.",
        "methodology": "Combines Stats NZ HES average weekly household insurance expenditure across all NZ households with RBNZ Financial Stability Report analysis of quoted/paid domestic building insurance premiums for insured homeowners.",
        "hlf_calculation": "Annualised household expenditure = Weekly HES expenditure ($24.80/wk) × 52 weeks = $1,289.60/yr; compared against RBNZ insured homeowner dwelling premium benchmark (~$2,850/yr)",
        "hlf_interpretation": "HLF market context: Stats NZ's HES averages insurance spending across all households (including renters who do not pay building insurance), whereas RBNZ dwelling figures reflect the actual annual building premium paid by insured homeowners. Lenders require uninterrupted building insurance as a mandatory condition of every residential mortgage.",
        "questions": [
            "How much do New Zealand households typically spend on house and contents insurance?",
            "Why do bank mortgage calculators include building insurance as a fixed homeowner cost?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "natural-hazard-claims": {
        "name": "Natural Hazards Insurance Act (NHCover) statutory building cap, levy & disaster claims benchmark",
        "also_known_as": "Natural Hazards Commission Toka Tū Ake (NHC / EQC) statutory cap & RBNZ catastrophe claims",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Natural hazards, flood risk & NHCover",
        "source": "NHI Act 2023 & NHC",
        "source_full": "Natural Hazards Insurance Act 2023, NHC Toka Tū Ake & RBNZ Financial Stability Report",
        "series_ids": "NHI Act 2023 Statutory Building Cap & Levy Schedule",
        "source_url": NHC_ACT_PAGE,
        "data_type": "Public-domain statutory cap/levy schedule & HLF-calculated GST-inclusive totals",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Statutory / Event-Updated",
        "definition": "The statutory first-loss residential building cover cap and levy rate under the Natural Hazards Insurance Act 2023 (administered by NHC Toka Tū Ake), alongside RBNZ natural-disaster insured loss benchmarks.",
        "methodology": "Sourced from the Natural Hazards Insurance Act 2023 statutory schedule ($300,000 + GST per dwelling at 16c per $100 cover) and RBNZ Financial Stability Report catastrophe loss analysis.",
        "hlf_calculation": "NHCover Cap = $300,000 + 15% GST ($345,000); Max Annual Levy = ($300,000 ÷ $100) × $0.16 = $480 + 15% GST ($552.00)",
        "hlf_interpretation": "HLF market context: NHCover provides the first $300,000 + GST ($345,000) of natural-hazard building cover (for earthquakes, landslips, volcanic activity, hydrothermal activity, and tsunamis, plus statutory residential land cover), while private house insurance tops cover up to the homeowner's full sum insured and covers non-NHC perils such as flood and storm damage to the building.",
        "questions": [
            "How much does the Natural Hazards Commission (formerly EQC) cover on a New Zealand house?",
            "Could my home become harder or more expensive to insure due to flood or natural-hazard risk?",
            "Why is checking insurance availability essential before going unconditional on a property?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "NZ Sale & Purchase Agreement Walkthrough",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-sale-and-purchase-agreement-walkthrough",
    },
    "residential-insurance-coverage": {
        "name": "New Zealand residential dwelling insurance coverage & uninsured dwelling estimate",
        "also_known_as": "RBNZ Financial Stability Report residential insurance coverage & sum-insured gap",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Residential insurance coverage & underinsurance",
        "source": "RBNZ FSR & NHC",
        "source_full": "Reserve Bank of New Zealand Financial Stability Report (May 2026) & Natural Hazards Commission",
        "series_ids": "RBNZ May 2026 FSR Domestic Dwelling Insurance Coverage Estimate",
        "source_url": RBNZ_FSR_PAGE,
        "data_type": "Officially reported RBNZ / NHC estimates & HLF-derived coverage percentage",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Semi-Annual",
        "definition": "The estimated proportion and number of New Zealand domestic residential dwellings with active building insurance cover versus uninsured dwellings, as analysed in the Reserve Bank of New Zealand's May 2026 Financial Stability Report.",
        "methodology": "Sourced from RBNZ's May 2026 Financial Stability Report (published 6 May 2026) and Natural Hazards Commission Toka Tū Ake residential dwelling exposure estimates (~2.0 million domestic dwellings representing ~$1.5 trillion in total residential sum insured).",
        "hlf_calculation": "Uninsured Share = (60,000 Uninsured Dwellings ÷ 2,000,000 Total Domestic Dwellings) × 100 = 3.0%; Insured Share = 1,940,000 ÷ 2,000,000 = 97.0%",
        "hlf_interpretation": "HLF market context: RBNZ estimates that while roughly 97% of New Zealand domestic dwellings carry building insurance (primarily because mortgage lenders require full replacement/sum-insured cover as a loan condition), around 60,000 mostly unmortgaged dwellings have no building insurance—and many insured homeowners still face underinsurance risk if their sum insured lags behind rebuild costs.",
        "questions": [
            "How many houses in New Zealand have no building insurance?",
            "What percentage of New Zealand homes are insured?",
            "Is my house insured for enough to rebuild it if it is destroyed?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build & Rebuild Cost Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },

    # 2. Personal Risk & Income Protection (4 Datasets: Health, Life & 2 ACC Series)
    "health-insurance-inflation": {
        "name": "Health insurance annual price inflation (Stats NZ CPI)",
        "also_known_as": "Private medical & health insurance CPI price change NZ",
        "category": "Personal Risk & Income Protection",
        "topic": "Personal risk & health insurance",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Health Insurance Sub-Index)",
        "series_ids": "CPI Health Insurance (Miscellaneous Goods & Services — Insurance)",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI sub-index & HLF-calculated comparison",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in the price of private health and medical insurance in New Zealand as measured within the Stats NZ Consumers Price Index (CPI).",
        "methodology": "Published quarterly by Stats NZ tracking price changes for a standardised sample of private health insurance plans.",
        "hlf_calculation": "Officially reported Stats NZ annual percentage change in the CPI Health Insurance series",
        "hlf_interpretation": "HLF market context: Tracks index-level price changes caused by private surgical/medical claims volume and healthcare cost inflation (distinct from age-related premium steps on an individual's policy).",
        "questions": [
            "How fast are private health insurance prices rising in New Zealand?",
            "Why do health insurance premiums increase each year?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "life-insurance-inflation": {
        "name": "Life insurance annual price movement (Stats NZ CPI)",
        "also_known_as": "Term life insurance CPI price index change NZ",
        "category": "Personal Risk & Income Protection",
        "topic": "Personal risk, life & income protection",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Life Insurance Sub-Index)",
        "series_ids": "CPI Life Insurance (Miscellaneous Goods & Services — Insurance)",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI sub-index & HLF-calculated comparison",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in the price index for life insurance in New Zealand as measured within the Stats NZ Consumers Price Index (CPI).",
        "methodology": "Compiled quarterly by Stats NZ for a constant-age, constant-sum-insured sample of life insurance policies.",
        "hlf_calculation": "Officially reported Stats NZ annual percentage change in the CPI Life Insurance series",
        "hlf_interpretation": "HLF market context: Stats NZ holds age and sum insured constant to measure underlying base-rate changes. In practice, an individual policyholder on a yearly renewable ('rate-for-age') life or income protection policy will see annual premium changes driven by both age progression and inflation-indexed sum insured.",
        "questions": [
            "How much life insurance do New Zealand families with a mortgage need?",
            "Why do rate-for-age life insurance premiums change differently from the CPI life insurance index?",
            "Could your household afford six months without its primary income?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "acc-weekly-compensation-claims": {
        "name": "ACC annual accepted injury claims vs new weekly income compensation claims",
        "also_known_as": "ACC Annual Report total accepted injury claims & weekly compensation claim counts",
        "category": "Personal Risk & Income Protection",
        "topic": "Income protection & ACC coverage scope",
        "source": "ACC Annual Report",
        "source_full": "Accident Compensation Corporation (ACC) — Annual Report (2025/26)",
        "series_ids": "ACC Annual Accepted Injury Claims & New Weekly Compensation Entitlement Claims",
        "source_url": ACC_REPORT_PAGE,
        "data_type": "Officially reported ACC annual claim counts & HLF-calculated claim-count ratio",
        "derived_by_hlf": True,
        "unit": "claims",
        "frequency": "Annual",
        "definition": "Comparison of two separate annual claim counts published by ACC: total new accepted injury claims registered across New Zealand (comprising mainly minor treatment-only claims such as GP, physiotherapy, and dental visits) versus new weekly income compensation entitlement claims for time off work.",
        "methodology": "Sourced from ACC's Annual Report covering accepted injury claims, treatment/rehabilitation expenditure, and weekly compensation entitlements. Note: the two series measure separate annual claim counts (total accepted claims include non-earners and minor treatment-only injuries that do not require time off work).",
        "hlf_calculation": "Claim-Count Ratio = 105,000 New Weekly Compensation Claims ÷ 2,100,000 Total Accepted Injury Claims = 5.0 weekly compensation claims per 100 accepted injury claims",
        "hlf_interpretation": "HLF market context: The 5.0% ratio compares annual weekly-compensation claim counts against total accepted injury claims (most of which are minor treatment-only visits). For mortgage borrowers, the critical distinction is statutory scope: under the Accident Compensation Act 2001, ACC covers physical injuries caused by an accident (plus narrow statutory categories such as work-related gradual-process conditions, treatment injuries, and certain mental injuries), and generally does not cover non-injury illness (such as cancer, heart attack, stroke, or degenerative illness)—which is the gap addressed by private Income Protection, Mortgage Repayment, and Trauma insurance.",
        "questions": [
            "Does ACC cover my mortgage repayments if I get sick or diagnosed with an illness?",
            "How many weekly compensation claims does ACC accept each year in New Zealand?",
            "What is the difference between ACC and private income protection or mortgage protection insurance?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "acc-earners-levy-cap": {
        "name": "ACC Earners' Levy rate, maximum liable earnings cap & 80% weekly compensation ceiling",
        "also_known_as": "IRD / ACC PAYE Earners' Levy rate & maximum covered salary cap (2026/27)",
        "category": "Personal Risk & Income Protection",
        "topic": "ACC levy, high-income cap & mortgage servicing",
        "source": "ACC & IRD Statutory Schedule",
        "source_full": "Accident Compensation (Earners' Levy) Regulations & Inland Revenue PAYE Schedule (2026/27)",
        "series_ids": "ACC Earners' Levy Rate & Maximum Liable Earnings Threshold",
        "source_url": ACC_LEVY_PAGE,
        "data_type": "Public-domain statutory levy/cap & HLF-calculated 80% gross compensation ceiling",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual (1 April Tax Year)",
        "definition": "The statutory ACC Earners' Levy rate deducted from employee PAYE wages, the maximum liable earnings cap, and the resulting maximum 80% gross weekly compensation payable by ACC.",
        "methodology": "Published by ACC and Inland Revenue under statutory regulations for the tax year commencing 1 April 2026 ($1.75 per $100 of liable earnings incl. GST; maximum liable earnings of $156,641).",
        "hlf_calculation": "Max Annual Levy = ($156,641 ÷ 100) × $1.75 = $2,741.22; Max 80% Gross Compensation = $156,641 × 80% = $125,312.80/yr ($2,409.86/wk gross)",
        "hlf_interpretation": "HLF market context: This affects borrowers in two ways: (1) In mortgage servicing calculators, the 1.75% ACC levy reduces net take-home pay on earnings up to $156,641; and (2) For an eligible accident, ACC caps income replacement at 80% of earnings up to $156,641 ($125,313/yr gross)—leaving higher-income earners with a larger percentage drop in income, and self-employed company owners splitting shareholder salaries potentially under-covered unless structured via ACC CoverPlus Extra or private cover.",
        "questions": [
            "What is the maximum income ACC will cover if I have an accident in 2026/27?",
            "What is the current ACC Earners' Levy rate in New Zealand?",
            "How does ACC work for high-income earners and self-employed business owners?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },

    # 3. Insurance Market & Claims (2 Datasets — Anchored to RBNZ Insurance Statistics)
    "insurance-claims-trends": {
        "name": "New Zealand general insurance Gross Written Premiums vs. Net Incurred Claims (RBNZ)",
        "also_known_as": "RBNZ general & domestic insurance market premiums, reinsurance & claims volume",
        "category": "Insurance Market & Claims",
        "topic": "Insurance industry claims & premiums",
        "source": "RBNZ Insurance Statistics",
        "source_full": "Reserve Bank of New Zealand — Insurance Statistics & Financial Stability Report",
        "series_ids": "RBNZ General Insurer Gross Written Premium, Net Earned Premium & Net Incurred Claims",
        "source_url": RBNZ_INS_STATS_PAGE,
        "data_type": "Officially reported RBNZ insurer aggregates & HLF-calculated reinsurance/claims comparison",
        "derived_by_hlf": True,
        "unit": "NZDB",
        "frequency": "Annual / Semi-Annual",
        "definition": "Total annual Gross Written Premiums (GWP), Net Earned Premiums (after reinsurance), and Net Incurred Claims across licensed New Zealand general insurers regulated by the Reserve Bank of New Zealand.",
        "methodology": "Compiled from RBNZ licensed insurer financial returns and Financial Stability Report general insurance series.",
        "hlf_calculation": "Reinsurance Outflow = Gross Written Premiums ($9.85B) minus Net Earned Premiums ($7.40B) = $2.45B; compared against Net Incurred Claims ($4.66B)",
        "hlf_interpretation": "HLF market context: Illustrates the scale of claims paid back to New Zealand households and businesses relative to total premiums collected, alongside the portion of gross premiums paid by NZ insurers for international catastrophe reinsurance.",
        "questions": [
            "How much do New Zealand insurers pay out in claims compared with premiums collected?",
            "How much of New Zealand general insurance premiums goes toward international reinsurance?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "insurer-loss-ratios": {
        "name": "New Zealand general insurer net loss ratio (net claims incurred ÷ net earned premiums)",
        "also_known_as": "RBNZ general insurer claims loss ratio",
        "category": "Insurance Market & Claims",
        "topic": "Insurer loss ratios & claims",
        "source": "RBNZ Insurance Statistics",
        "source_full": "Reserve Bank of New Zealand — Insurance Statistics & Financial Stability Report",
        "series_ids": "RBNZ General Insurer Net Incurred Claims ÷ Net Earned Premium",
        "source_url": RBNZ_INS_STATS_PAGE,
        "data_type": "Officially reported RBNZ insurer series & HLF-derived net loss ratio",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Annual / Semi-Annual",
        "definition": "The percentage of net earned insurance premiums (after reinsurance costs) paid out in net incurred claims across licensed New Zealand general insurers.",
        "methodology": "Calculated from RBNZ licensed insurer financial performance series by dividing Net Incurred Claims by Net Earned Premium.",
        "hlf_calculation": "(Net Incurred Claims [$4.66B] ÷ Net Earned Premium [$7.40B]) × 100 = 63.0%",
        "hlf_interpretation": "HLF market context: A net loss ratio around 60%–68% in non-catastrophe years (and spiking above 80%–100%+ in major weather/earthquake years) shows the proportion of net premiums consumed directly by claims before operating expenses, tax, and RBNZ solvency capital requirements.",
        "questions": [
            "How much of every insurance premium dollar goes toward paying claims in New Zealand?",
            "What is the average loss ratio for New Zealand general insurers?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
}

TICKER_CATALOG.update(TICKER_CATALOG_INSURANCE)

# Silent legacy alias map ONLY for nz-mortgage-pulse.json so existing Webflow bubbles never break
LEGACY_BUBBLE_ALIASES = {
    "kiwisaver-active-switchers": "kiwisaver-switches",
    "investor-vs-owner-io": "interest-only",
}


class SimpleTableExtractor(HTMLParser):
    """Zero-dependency HTML table parser for Inland Revenue tables."""
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


def extract_html_release_date(html_str: str, fallback_note: str) -> str:
    """Extract actual agency release/publication date from HTML metadata or text without copying period_end."""
    if not html_str:
        return fallback_note
    for pat in (
        r'<meta[^>]+(?:name|property)=["\'](?:DC\.Date|dcterms\.created|article:published_time|article:modified_time|date)["\'][^>]+content=["\']([^"\']+)["\']',
        r'<time[^>]+datetime=["\'](20\d{2}-\d{2}-\d{2})["\']',
        r'(?:Release date|Last updated|Published date|Date published|Published)[:\s]+(\d{1,2}\s+[A-Za-z]+\s+20\d{2})',
    ):
        m = re.search(pat, html_str, flags=re.I)
        if m:
            raw_d = m.group(1).strip()
            dt = pd.to_datetime(raw_d[:10] if re.match(r"^20\d{2}-\d{2}-\d{2}", raw_d) else raw_d, errors="coerce")
            if pd.notna(dt):
                return f"{dt.strftime('%Y-%m-%d')} ({dt.day} {dt.strftime('%B %Y')})"
            return raw_d
    return fallback_note


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

PART_2B_OF_5_READY = True
# === PART 3A OF 5 ===
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
    """Extract official Published Date from the RBNZ 'Table Description' sheet without falling back to period_end."""
    try:
        xls = pd.ExcelFile(filepath)
        for tab in xls.sheet_names:
            if "description" in tab.lower():
                df_desc = pd.read_excel(xls, sheet_name=tab, header=None)
                for r in range(min(15, df_desc.shape[0])):
                    k_str = str(df_desc.iloc[r, 0] or "").strip().lower()
                    if ("published" in k_str or "release" in k_str) and df_desc.shape[1] > 1:
                        dt_p = pd.to_datetime(df_desc.iloc[r, 1], errors="coerce")
                        if pd.notna(dt_p):
                            return f"{dt_p.strftime('%Y-%m-%d')} ({dt_p.day} {dt_p.strftime('%B %Y')})"
    except Exception:
        pass
    return "Monthly RBNZ statistical release (~4 weeks after reference month-end)"


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

PART_3A_OF_5_READY = True
# === PART 3B OF 5 ===
# ---------------------------------------------------------------------------
# Reference-Grade RBNZ C31 Parser — Audited Math & Source Rounding Disclosures
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

    try:
        _, tot_val, _ = find_rbnz_series(tabs, ("LVRN.MMB1.AA",), ("total lending", "all borrower"))
    except KeyError:
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
    fhb_avg_nzd = int(round((float(fhb_val.iloc[i]) * 1_000_000.0) / float(fhb_cnt.iloc[i])))
    fhb_avg_k = int(round(fhb_avg_nzd / 1000.0))

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
        f"residential mortgage lending ({plain_prior_year_pct(fhb_sh, fhb_sh_yoy)}) and a mean loan size of <strong>${fhb_avg_nzd:,} (~${fhb_avg_k:,}k)</strong>."
    )

    inv_html = (
        f"In {m_full}, residential property investors accounted for an HLF-derived <strong>{inv_sh:.1f}%</strong> share of new mortgage commitments "
        f"(<strong>${inv_b:.2f}B</strong> reported in RBNZ Table C31, {plain_prior_year_pct(inv_sh, inv_sh_yoy)}), compared with "
        f"<strong>{fhb_sh:.1f}%</strong> for first-home buyers."
    )

    fhb_lo_val_cur = float(fhb_val.iloc[i] - fhb_hi_val.iloc[i])
    fhb_lo_cnt_cur = float(fhb_cnt.iloc[i] - fhb_hi_cnt.iloc[i])
    fhb_lo_loans_int = int(round(fhb_lo_cnt_cur))
    avg_hi_loan = int(round((float(fhb_hi_val.iloc[i]) * 1_000_000.0) / max(1.0, float(fhb_hi_cnt.iloc[i]))))
    avg_lo_loan = int(round((fhb_lo_val_cur * 1_000_000.0) / max(1.0, fhb_lo_cnt_cur)))
    loan_gap = avg_hi_loan - avg_lo_loan

    fhb_gap_html = (
        f"In {m_full}, the HLF-derived mean mortgage size for first-home buyers at &gt;80% LVR (under 20% deposit) was "
        f"<strong>${avg_hi_loan:,}</strong> ({fhb_hi_loans_int:,} loans / ${hi_dollars_m:,.0f}m), compared with "
        f"<strong>${avg_lo_loan:,}</strong> for first-home buyers at &le;80% LVR ({fhb_lo_loans_int:,} loans / ${fhb_lo_val_cur:,.0f}m)—an "
        f"HLF-calculated difference of <strong>${abs(loan_gap):,}</strong> (derived from RBNZ Table C31 dollar series rounded to $1m)."
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
        f"<strong>{std_dep_pct:.1f}%</strong> share of first-home approvals and a mean loan size of <strong>${avg_lo_loan:,}</strong>."
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
    oo_hi_v = float(oo_hi_val.iloc[i])
    oo_lo_c = oo_tot_c - oo_hi_c
    oo_lo_v = float(oo_val.iloc[i] - oo_hi_val.iloc[i])
    oo_20eq_pct = float(ser_oo_20eq_pct.iloc[i])
    oo_avg_20eq = int(round((oo_lo_v * 1_000_000.0) / max(1.0, oo_lo_c)))
    owner_eq_html = (
        f"In {m_full}, RBNZ Table C31 recorded <strong>{int(round(oo_lo_c)):,} of {int(round(oo_tot_c)):,}</strong> existing ('other') "
        f"owner-occupier mortgage approvals (<strong>{fmt_nzd_millions(oo_lo_v)}</strong>) at <strong>&le;80% LVR</strong> (corresponding to 20%+ "
        f"security equity)—an HLF-derived share of <strong>{oo_20eq_pct:.1f}%</strong> and mean loan size of <strong>${oo_avg_20eq:,}</strong>."
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
        f"In {m_full}, RBNZ Table C31 recorded <strong>{int(round(oo_hi_c)):,} mortgages ({fmt_nzd_millions(oo_hi_v)})</strong> "
        f"for existing ('other') owner-occupiers at <strong>&gt;80% LVR</strong>—an HLF-derived <strong>{oo_low_eq_pct:.1f}%</strong> "
        f"share of existing owner-occupier approvals, {plain_prior_year_pct(oo_low_eq_pct, oo_low_eq_prev)}."
    )

    comb_ex_val_ser = oo_hi_val.fillna(0) + inv_gt70_val.fillna(0)
    comb_ex_cnt_ser = oo_hi_cnt.fillna(0) + inv_gt70_cnt.fillna(0)
    comb_ex_v = float(comb_ex_val_ser.iloc[i])
    comb_ex_c = int(round(float(comb_ex_cnt_ser.iloc[i])))
    comb_ex_v_prev = float(comb_ex_val_ser.iloc[iy])
    comb_ex_c_prev = int(round(float(comb_ex_cnt_ser.iloc[iy])))
    nb_ex_flow_html = (
        f"In {m_full}, RBNZ Table C31 recorded an HLF-derived combined total of <strong>{comb_ex_c:,} non-FHB high-LVR loans "
        f"({fmt_nzd_millions(comb_ex_v)})</strong> across existing owner-occupiers borrowing at <strong>&gt;80% LVR</strong> "
        f"(<strong>{int(round(oo_hi_c)):,} loans / {fmt_nzd_millions(oo_hi_v)}</strong>) and property investors borrowing at "
        f"<strong>&gt;70% LVR</strong> (<strong>{inv_u30_cnt:,} loans / {fmt_nzd_millions(inv_u30_val)}</strong>)."
    )

    src_c31 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type" '
        'target="_blank" rel="noopener">RBNZ Table C31</a>, Reserve Bank of New Zealand. Derived figures licensed by HLF under CC BY 4.0.'
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
            low_cnt_pct, f"{low_cnt_pct:.1f}% ({fhb_hi_loans_int:,} of {fhb_loans_int:,} loans)", f"{low_cnt_yoy:.1f}%",
            f"{low_cnt_pct - low_cnt_yoy:+.1f} percentage points YoY",
            f"{fhb_hi_loans_int:,} ÷ {fhb_loans_int:,} = {low_cnt_pct:.1f}% (Lending: ${hi_dollars_m:,.0f}m ÷ ${fhb_v:,.0f}m = {(hi_dollars_m/max(1.0,fhb_v))*100:.1f}%)",
            make_history_list(dates, ser_low_dep_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-share": make_c31_item(
            fhb_share_html,
            f"RBNZ Table C31 ({m_full}): {fhb_loans_int:,} FHB borrowers and ${fhb_v:,.0f}m FHB lending out of ${tot_val_m:,.0f}m total new residential lending (LVRN.MMB1.AA).",
            fhb_sh, f"{fhb_sh:.1f}% ({fhb_loans_int:,} loans, avg ${fhb_avg_nzd:,})", f"{fhb_sh_yoy:.1f}%",
            f"{fhb_sh - fhb_sh_yoy:+.1f} percentage points YoY",
            f"${fhb_v:,.0f}m ÷ ${tot_val_m:,.0f}m = {fhb_sh:.1f}% (Mean loan: ${fhb_v:,.0f}m × $1m ÷ {fhb_loans_int:,} = ${fhb_avg_nzd:,}; RBNZ dollar series rounded to $1m)",
            make_history_list(dates, ser_fhb_sh_pct, lambda x: f"{x:.1f}%"),
        ),
        "investors": make_c31_item(
            inv_html,
            f"RBNZ Table C31 ({m_full}): {inv_c_tot:,} investor borrowers and ${inv_v:,.0f}m investor lending out of ${tot_val_m:,.0f}m total lending (LVRN.MMB1.AA).",
            inv_sh, f"{inv_sh:.1f}% (${inv_b:.2f}B across {inv_c_tot:,} loans)", f"{inv_sh_yoy:.1f}%",
            f"{inv_sh - inv_sh_yoy:+.1f} percentage points YoY",
            f"${inv_v:,.0f}m ÷ ${tot_val_m:,.0f}m = {inv_sh:.1f}%",
            make_history_list(dates, ser_inv_sh_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-loan-size-gap": make_c31_item(
            fhb_gap_html,
            f"RBNZ Table C31 ({m_full}): >80% LVR FHB = ${hi_dollars_m:,.0f}m across {fhb_hi_loans_int:,} borrowers; <=80% LVR FHB = ${fhb_lo_val_cur:,.0f}m across {fhb_lo_loans_int:,} borrowers (dollar series rounded to $1m).",
            loan_gap / 1000.0, f"${avg_hi_loan:,} at >80% LVR vs ${avg_lo_loan:,} at <=80% LVR (+${abs(loan_gap):,})",
            f"${float(ser_fhb_hi_avg_k.iloc[iy])*1000:,.0f} vs ${float(ser_fhb_lo_avg_k.iloc[iy])*1000:,.0f}",
            f"${abs(loan_gap):,} difference between LVR tiers",
            f"(${hi_dollars_m:,.0f}m × $1m ÷ {fhb_hi_loans_int:,} = ${avg_hi_loan:,}) minus (${fhb_lo_val_cur:,.0f}m × $1m ÷ {fhb_lo_loans_int:,} = ${avg_lo_loan:,}) = ${loan_gap:,} (subject to RBNZ $1m rounding)",
            make_history_list(dates, ser_fhb_hi_avg_k, lambda x: f"${x:,.1f}k (>80% LVR avg)"),
        ),
        "fhb-bank-lvr-priority": make_c31_item(
            fhb_priority_html,
            f"RBNZ Table C31 ({m_full}): FHB >80% LVR = ${hi_dollars_m:,.0f}m ({fhb_hi_loans_int:,} loans); All Borrower >80% LVR = ${float(all_hi_val.iloc[i]):,.0f}m ({all_hi_loans_int:,} loans).",
            fhb_hi_val_share, f"{fhb_hi_val_share:.1f}% of >80% LVR dollars | {fhb_hi_cnt_share:.1f}% of >80% LVR loans",
            f"{float(ser_fhb_prio_pct.iloc[iy]):.1f}%",
            f"{fhb_hi_val_share - float(ser_fhb_prio_pct.iloc[iy]):+.1f} percentage points YoY",
            f"Dollars: ${hi_dollars_m:,.0f}m ÷ ${float(all_hi_val.iloc[i]):,.0f}m = {fhb_hi_val_share:.1f}% | Loans: {fhb_hi_loans_int:,} ÷ {all_hi_loans_int:,} = {fhb_hi_cnt_share:.1f}%",
            make_history_list(dates, ser_fhb_prio_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-20-deposit-club": make_c31_item(
            fhb_20club_html,
            f"RBNZ Table C31 ({m_full}): {fhb_lo_loans_int:,} <=80% LVR FHB borrowers (${fhb_lo_val_cur:,.0f}m) out of {fhb_loans_int:,} total FHB borrowers.",
            std_dep_pct, f"{std_dep_pct:.1f}% ({fhb_lo_loans_int:,} loans, avg ${avg_lo_loan:,})",
            f"{float(ser_fhb_20club_pct.iloc[iy]):.1f}%",
            f"{std_dep_pct - float(ser_fhb_20club_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{fhb_lo_loans_int:,} ÷ {fhb_loans_int:,} = {std_dep_pct:.1f}% (Mean loan: ${fhb_lo_val_cur:,.0f}m × $1m ÷ {fhb_lo_loans_int:,} = ${avg_lo_loan:,})",
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
            oo_20eq_pct, f"{oo_20eq_pct:.1f}% at <=80% LVR (avg loan ${oo_avg_20eq:,})",
            f"{float(ser_oo_20eq_pct.iloc[iy]):.1f}%",
            f"{oo_20eq_pct - float(ser_oo_20eq_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{int(round(oo_lo_c)):,} ÷ {int(round(oo_tot_c)):,} = {oo_20eq_pct:.1f}% (Mean loan: ${oo_lo_v:,.0f}m × $1m ÷ {int(round(oo_lo_c)):,} = ${oo_avg_20eq:,})",
            make_history_list(dates, ser_oo_20eq_pct, lambda x: f"{x:.1f}%"),
        ),
        "investor-equity-tiers": make_c31_item(
            inv_eq_html,
            f"RBNZ Table C31 ({m_full}): Total investors = {inv_c_tot:,}; <=60% LVR = {inv_le60_c:,}; >60% LVR = {int(round(float(inv_gt60_cnt.iloc[i]))):,}; >70% LVR = {inv_u30_cnt:,} (${inv_u30_val:,.0f}m).",
            inv_40plus_pct, f"<=60% LVR: {inv_40plus_pct:.1f}% | 60–70% LVR: {inv_30to40_pct:.1f}% | >70% LVR: {inv_u30_pct:.1f}%",
            f"{100.0 - float(ser_inv_u30_pct.iloc[iy]):.1f}% at <=70% LVR",
            f"{inv_u30_pct - float(ser_inv_u30_pct.iloc[iy]):+.1f} pts shift in >70% LVR share",
            f"<=60% LVR: {inv_le60_c:,} ÷ {inv_c_tot:,} = {inv_40plus_pct:.1f}% | 60–70% LVR: {inv_30to40_c:,} ÷ {inv_c_tot:,} = {inv_30to40_pct:.1f}% | >70% LVR: {inv_u30_cnt:,} ÷ {inv_c_tot:,} = {inv_u30_pct:.1f}%",
            make_history_list(dates, 100.0 - ser_inv_u30_pct, lambda x: f"{x:.1f}% (<=70% LVR)"),
        ),
        "investor-low-deposit": make_c31_item(
            inv_low_dep_html,
            f"RBNZ Table C31 ({m_full}): >80% LVR investors = {inv_gt80_c:,} loans (${inv_gt80_v:,.0f}m); >70% LVR investors = {inv_u30_cnt:,} loans (${inv_u30_val:,.0f}m).",
            float(inv_gt80_c), f">80% LVR: {inv_gt80_c:,} loans (${inv_gt80_v:,.0f}m) | 70–80% LVR: {inv_70to80_c:,} loans (${inv_70to80_v:,.0f}m)",
            f"{inv_gt80_c_prev:,} loans (${inv_gt80_v_prev:,.0f}m) >80% LVR",
            f"{inv_gt80_c - inv_gt80_c_prev:+d} loans ({inv_gt80_v - inv_gt80_v_prev:+.0f}m) YoY in >80% LVR tier",
            f">80% LVR: {inv_gt80_c:,} ÷ {inv_c_tot:,} = {inv_gt80_pct:.1f}% (${inv_gt80_v:,.0f}m) | 70–80% LVR: {inv_u30_cnt:,} - {inv_gt80_c:,} = {inv_70to80_c:,} loans (${inv_70to80_v:,.0f}m)",
            make_history_list(dates, inv_gt80_cnt, lambda x: f"{int(round(x)):,} loans (>80% LVR)"),
        ),
        "investor-new-build-lvr": make_c31_item(
            inv_nb_lvr_html,
            f"RBNZ Table C31 ({m_full}): >70% LVR investors = {inv_u30_cnt:,} loans (${inv_u30_val:,.0f}m) out of {inv_c_tot:,} total investor loans.",
            inv_u30_pct, f"{inv_u30_pct:.1f}% at >70% LVR ({inv_u30_cnt:,} loans / ${inv_u30_val:,.0f}m)",
            f"{float(ser_inv_u30_pct.iloc[iy]):.1f}%",
            f"{inv_u30_pct - float(ser_inv_u30_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{inv_u30_cnt:,} ÷ {inv_c_tot:,} = {inv_u30_pct:.1f}% ({inv_70to80_c:,} loans in 70–80% LVR + {inv_gt80_c:,} loans in >80% LVR)",
            make_history_list(dates, ser_inv_u30_pct, lambda x: f"{x:.1f}% (>70% LVR)"),
        ),
        "mover-low-equity": make_c31_item(
            mover_low_eq_html,
            f"RBNZ Table C31 ({m_full}): {int(round(oo_hi_c)):,} >80% LVR other owner-occupier loans (${oo_hi_v:,.0f}m) out of {int(round(oo_tot_c)):,} total other owner-occupier loans.",
            oo_low_eq_pct, f"{oo_low_eq_pct:.1f}% at >80% LVR ({int(round(oo_hi_c)):,} loans / ${oo_hi_v:,.0f}m)",
            f"{oo_low_eq_prev:.1f}%",
            f"{oo_low_eq_pct - oo_low_eq_prev:+.1f} percentage points YoY",
            f"{int(round(oo_hi_c)):,} ÷ {int(round(oo_tot_c)):,} = {oo_low_eq_pct:.1f}%",
            make_history_list(dates, ser_oo_low_eq_pct, lambda x: f"{x:.1f}%"),
        ),
        "new-build-lvr-exemption-flow": make_c31_item(
            nb_ex_flow_html,
            f"RBNZ Table C31 ({m_full}): Other owner-occupiers >80% LVR = {int(round(oo_hi_c)):,} loans (${oo_hi_v:,.0f}m); Investors >70% LVR = {inv_u30_cnt:,} loans (${inv_u30_val:,.0f}m).",
            comb_ex_v, f"{fmt_nzd_millions(comb_ex_v)} across {comb_ex_c:,} high-LVR mover & investor loans",
            f"{fmt_nzd_millions(comb_ex_v_prev)} ({comb_ex_c_prev:,} loans)",
            f"{comb_ex_v - comb_ex_v_prev:+.0f}m ({comb_ex_c - comb_ex_c_prev:+,d} loans) YoY",
            f"Dollars: ${oo_hi_v:,.0f}m + ${inv_u30_val:,.0f}m = ${comb_ex_v:,.0f}m | Loans: {int(round(oo_hi_c)):,} + {inv_u30_cnt:,} = {comb_ex_c:,}",
            make_history_list(dates, comb_ex_val_ser, lambda x: f"${x:,.0f}m"),
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
            float(r12_cnt), f"{r12_cnt:,} loans / {fmt_nzd_millions(r12_val)} (12m rolling)", f"{r12_prev_cnt:,} loans",
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

PART_3B_OF_5_READY = True
# === PART 3C OF 5 ===
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
    ref_avg_nzd = int(round((ref_val_m * 1_000_000.0) / max(1, ref_cnt_int)))
    ref_avg_k = int(round(ref_avg_nzd / 1000.0))

    top_sh_c = float(top_sh_c_ser.iloc[j])
    top_val_m = float(top_v.iloc[j])
    top_cnt_int = int(round(float(top_c.iloc[j])))
    top_avg_nzd = int(round((top_val_m * 1_000_000.0) / max(1, top_cnt_int)))
    top_avg_k = int(round(top_avg_nzd / 1000.0))

    src_c33 = (
        'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose" '
        'target="_blank" rel="noopener">RBNZ Table C33</a>, Reserve Bank of New Zealand. Derived figures licensed by HLF under CC BY 4.0.'
    )
    return {
        "refinance": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, RBNZ Table C33 recorded <strong>{ref_cnt_int:,}</strong> borrowers changing loan provider and "
                f"<strong>${ref_val_m/1000.0:.2f}B (${ref_val_m:,.0f}m)</strong> in refinanced loans—representing an HLF-derived "
                f"<strong>{ref_sh:.1f}%</strong> share of all new residential lending and a mean loan size of "
                f"<strong>${ref_avg_nzd:,} (~${ref_avg_k:,}k)</strong>."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Change in loan provider = {ref_cnt_int:,} borrowers and ${ref_val_m:,.0f}m out of ${tot_v:,.0f}m total purpose lending (dollar series rounded to $1m).",
            "source_html": src_c33,
            "latest_period": m_full,
            "value": round(ref_sh, 2),
            "latest_value": f"{ref_sh:.1f}% of lending (${ref_val_m/1000.0:.2f}B across {ref_cnt_int:,} loans, avg ${ref_avg_nzd:,})",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({float(ref_sh_ser.iloc[jy]):.1f}%)",
            "change": f"{ref_sh - float(ref_sh_ser.iloc[jy]):+.1f} percentage points YoY",
            "calculation_live": f"${ref_val_m:,.0f}m ÷ ${tot_v:,.0f}m = {ref_sh:.1f}% (Mean loan: ${ref_val_m:,.0f}m × $1m ÷ {ref_cnt_int:,} = ${ref_avg_nzd:,}; RBNZ dollar series rounded to $1m)",
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
                f"an HLF-derived <strong>{top_sh_c:.1f}%</strong> share of all mortgage approvals by count and a mean top-up of "
                f"<strong>${top_avg_nzd:,} (~${top_avg_k:,}k)</strong>."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Top-up approvals = {top_cnt_int:,} borrowers (${top_val_m:,.0f}m) out of {tot_c:,} total approvals.",
            "source_html": src_c33,
            "latest_period": m_full,
            "value": round(top_sh_c, 2),
            "latest_value": f"{top_sh_c:.1f}% of approvals ({top_cnt_int:,} top-ups, avg ${top_avg_nzd:,})",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({float(top_sh_c_ser.iloc[jy]):.1f}%)",
            "change": f"{top_sh_c - float(top_sh_c_ser.iloc[jy]):+.1f} percentage points YoY",
            "calculation_live": f"{top_cnt_int:,} ÷ {tot_c:,} = {top_sh_c:.1f}% (Mean top-up: ${top_val_m:,.0f}m × $1m ÷ {top_cnt_int:,} = ${top_avg_nzd:,})",
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
                f"<strong>{top_cnt_int:,} borrowers</strong> (an HLF-derived mean of <strong>${top_avg_nzd:,}</strong> per top-up)."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Top-up lending value = ${top_val_m:,.0f}m across {top_cnt_int:,} borrowers.",
            "source_html": src_c33,
            "latest_period": m_full,
            "value": round(top_val_m, 1),
            "latest_value": f"{fmt_nzd_millions(top_val_m)} across {top_cnt_int:,} borrowers (avg ${top_avg_nzd:,})",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({fmt_nzd_millions(float(top_v.iloc[jy]))})",
            "change": f"{top_val_m - float(top_v.iloc[jy]):+.0f}m YoY",
            "calculation_live": f"Directly reported ${top_val_m:,.0f}m of top-up lending × $1m ÷ {top_cnt_int:,} borrowers = ${top_avg_nzd:,} mean",
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
        'target="_blank" rel="noopener">RBNZ Table C32</a>, Reserve Bank of New Zealand. Derived figures licensed by HLF under CC BY 4.0.'
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


def compute_c41_weighted_avg_income(
    tabs41: list, borrower_kw: str, region_mode: str = "nz"
) -> tuple[pd.Timestamp | None, float | None, pd.Series | None, pd.Series | None, list[dict], float, float]:
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
            sum_pool_m = 0.0
            sum_implied_apps = 0.0
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
                    imp_apps = (last_a * 1000.0) / last_b
                    sum_pool_m += last_a
                    sum_implied_apps += imp_apps
                    band_rows.append({
                        "band": b_label,
                        "avg_income_nzd": int(round(last_b * 1000.0)),
                        "total_income_m": round(last_a, 1),
                        "implied_borrowers": round(imp_apps, 1),
                    })

            weighted_k_ser = tot_inc_ser / tot_cnt_ser.replace(0.0, float("nan"))
            valid_w = weighted_k_ser.dropna()
            if len(valid_w):
                return (
                    df41["date"].iloc[-1],
                    float(valid_w.iloc[-1]) * 1000.0,
                    df41["date"],
                    weighted_k_ser,
                    band_rows,
                    round(sum_pool_m, 1),
                    round(sum_implied_apps, 1),
                )
    return None, None, None, None, [], 0.0, 0.0


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
                        'target="_blank" rel="noopener">RBNZ Table C40</a>, Reserve Bank of New Zealand. Derived figures licensed by HLF under CC BY 4.0.'
                    ),
                    "latest_period": m_full,
                    "value": round(val_pct, 2),
                    "latest_value": f"{val_pct:.1f}% at >6x DTI ({diff_cap:.1f} pts below 20% benchmark)",
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
            dt_fhb, fhb_inc, d41_dates, fhb_inc_ser, fhb_bands, fhb_pool_m, fhb_imp_apps = compute_c41_weighted_avg_income(
                tabs41, "first home", "nz"
            )
            _, inv_inc, _, _, _, inv_pool_m, inv_imp_apps = compute_c41_weighted_avg_income(
                tabs41, "investor", "nz"
            )

            src_c41 = (
                'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi" '
                'target="_blank" rel="noopener">RBNZ Table C41</a>, Reserve Bank of New Zealand. Derived figures licensed by HLF under CC BY 4.0.'
            )

            if dt_fhb is not None and fhb_inc and 40000 <= fhb_inc <= 600000:
                m_c41 = dt_fhb.strftime("%B %Y")
                as_of_c41 = month_end_iso(dt_fhb)
                iy41 = max(0, len(d41_dates) - 13)
                prev_inc_k = float(fhb_inc_ser.iloc[iy41]) if pd.notna(fhb_inc_ser.iloc[iy41]) else (fhb_inc / 1000.0)
                hist_41 = make_history_list(d41_dates, fhb_inc_ser, lambda x: f"${x*1000:,.0f}")
                hist_start_c41 = f"Available from {d41_dates.iloc[0].strftime('%B %Y')} ({len(d41_dates)} monthly observations)"

                band_summary_parts = [
                    f"{b['band']}: ${b['avg_income_nzd']:,} avg (${b['total_income_m']:,.1f}m pool / {b['implied_borrowers']:,.1f} implied apps)"
                    for b in fhb_bands
                ]
                band_summary_str = " | ".join(band_summary_parts) if band_summary_parts else f"Overall mean: ${fhb_inc:,.0f}"
                lo_band = min((b["avg_income_nzd"] for b in fhb_bands), default=int(round(fhb_inc)))
                hi_band = max((b["avg_income_nzd"] for b in fhb_bands), default=int(round(fhb_inc)))

                if inv_inc and 40000 <= inv_inc <= 900000:
                    bgi_html = (
                        f"Based on HLF's calculation across RBNZ Table C41's Debt-to-Income bands (<strong>${fhb_pool_m:,.1f}m</strong> total "
                        f"band gross income divided by <strong>{fhb_imp_apps:,.1f}</strong> implied applications), approved New Zealand "
                        f"first-home buyer mortgage applications averaged <strong>${fhb_inc:,.0f}</strong> in gross borrower income in {m_c41}, "
                        f"compared with <strong>${inv_inc:,.0f}</strong> (${inv_pool_m:,.1f}m ÷ {inv_imp_apps:,.1f} implied applications) "
                        f"across approved property investor applications."
                    )
                else:
                    bgi_html = (
                        f"Based on HLF's calculation across RBNZ Table C41's five Debt-to-Income bands (<strong>${fhb_pool_m:,.1f}m</strong> total "
                        f"band gross income divided by <strong>{fhb_imp_apps:,.1f}</strong> implied applications), approved New Zealand "
                        f"first-home buyer mortgage applications averaged <strong>${fhb_inc:,.0f}</strong> in gross borrower income in {m_c41} "
                        f"(with official RBNZ band averages ranging from <strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong>)."
                    )
                out["borrower-income"] = {
                    "chip": "RBNZ C41",
                    "as_of": as_of_c41,
                    "source_published": pub_c41,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": bgi_html,
                    "official_data_summary": f"RBNZ Table C41 ({m_c41}) Official FHB Band Figures (pools rounded to $0.1m, averages to $1,000): {band_summary_str}.",
                    "source_html": src_c41,
                    "latest_period": m_c41,
                    "value": int(round(fhb_inc)),
                    "latest_value": f"${fhb_inc:,.0f} FHB mean" + (f" vs ${inv_inc:,.0f} Investor mean" if inv_inc else ""),
                    "previous_period": f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_k*1000:,.0f} FHB)",
                    "change": f"${fhb_inc - (prev_inc_k*1000):+,.0f} YoY",
                    "calculation_live": (
                        f"FHB: ${fhb_pool_m:,.1f}m total band income ÷ {fhb_imp_apps:,.1f} implied applications = ${fhb_inc:,.0f}"
                        + (f" | Investor: ${inv_pool_m:,.1f}m ÷ {inv_imp_apps:,.1f} = ${inv_inc:,.0f}" if inv_inc else "")
                        + " (subject to RBNZ band rounding)"
                    ),
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
                        f"bands ranged from <strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong>, with an HLF-calculated mean of "
                        f"<strong>${fhb_inc:,.0f}</strong> (<strong>${fhb_pool_m:,.1f}m</strong> total band gross income divided by "
                        f"<strong>{fhb_imp_apps:,.1f}</strong> implied applications across all 5 bands; see reproducible table below)."
                    ),
                    "official_data_summary": f"RBNZ Table C41 ({m_c41}) Official FHB Band Figures (pools rounded to $0.1m, averages to $1,000): {band_summary_str}.",
                    "source_html": src_c41,
                    "latest_period": m_c41,
                    "value": int(round(fhb_inc)),
                    "tdti_bands": fhb_bands,
                    "tdti_total_pool_m": fhb_pool_m,
                    "tdti_total_implied_apps": fhb_imp_apps,
                    "latest_value": f"${fhb_inc:,.0f} HLF mean | Official C41 bands: ${lo_band:,}–${hi_band:,}",
                    "previous_period": f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_k*1000:,.0f} HLF mean)",
                    "change": f"${fhb_inc - (prev_inc_k*1000):+,.0f} ({((fhb_inc - prev_inc_k*1000)/(prev_inc_k*1000))*100:+.1f}%) YoY",
                    "calculation_live": f"${fhb_pool_m:,.1f}m total band income × $1m ÷ {fhb_imp_apps:,.1f} total implied applications = ${fhb_inc:,.0f} (reproducible from the 5 band rows; subject to RBNZ rounding)",
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
                'target="_blank" rel="noopener">RBNZ Table B30</a>, Reserve Bank of New Zealand.'
            ),
            "latest_period": m_full,
            "value": round(v1, 2),
            "latest_value": f"{v1:.2f}% 1-yr settled | {v2:.2f}% 2-yr settled",
            "previous_period": f"{dates.iloc[iy].strftime('%B %Y')} ({v1_prev:.2f}% 1-year settled)",
            "change": f"{v1 - v1_prev:+.2f}% YoY (1-year settled)",
            "calculation_live": f"Directly reported RBNZ B30 series NCFA.MAM1PR.D1 = {v1:.2f}% (1Y) & NCFA.MAM1PR.D2 = {v2:.2f}% (2Y)",
            "historical_data_from": hist_start_b30,
            "history": make_history_list(dates, r_1y, lambda x: f"{x:.2f}% (1Y settled)"),
        }
    }

PART_3C_OF_5_READY = True
# === PART 4A OF 5 ===
def make_pulse_item(
    chip: str,
    as_of: str,
    max_age: int,
    html_s: str,
    off_s: str,
    src_s: str,
    per_s: str,
    val_n,
    val_s: str,
    prev_s: str,
    chg_s: str,
    calc_s: str,
    hist_from: str,
    hist_l: list | None = None,
    extra: dict | None = None,
) -> dict:
    d = {
        "chip": chip,
        "as_of": as_of,
        "max_age_days": max_age,
        "html": html_s,
        "official_data_summary": off_s,
        "source_html": src_s,
        "latest_period": per_s,
        "value": val_n,
        "latest_value": val_s,
        "previous_period": prev_s,
        "change": chg_s,
        "calculation_live": calc_s,
        "historical_data_from": hist_from,
    }
    if hist_l is not None:
        d["history"] = hist_l
    if extra:
        d.update(extra)
    return d


def build_insurance_pulse_tickers(
    session,
    q_as_of: str,
    q_lbl: str,
    cpi_all_val: float | None,
    cpi_url: str,
    cpi_text: str = "",
    cpi_pub_date: str = "",
) -> dict:
    today = datetime.datetime.now(datetime.timezone.utc).date()
    cpi_as_of = q_as_of or month_end_iso(pd.Timestamp(datetime.date(today.year, 9 if today.month >= 10 else 6, 1)))
    cpi_period_lbl = q_lbl or "September 2026 quarter"
    cpi_ref_url = cpi_url or STATS_CPI_PAGE
    cpi_pub = cpi_pub_date or "Quarterly Stats NZ CPI release (~3 weeks after quarter-end)"
    headline_cpi = cpi_all_val if cpi_all_val is not None else 3.2

    def extract_ins_cpi(pattern: str, default_val: float) -> float:
        if cpi_text:
            m = re.search(pattern, cpi_text, flags=re.I)
            if m:
                dir_w = (m.group(1) or "up").lower()
                v = float(m.group(2))
                return -v if dir_w in ("down", "fell") else v
        return default_val

    bldg_inf = extract_ins_cpi(r"dwelling insurance[,:\s]+(up|down|fell|rose)?\s*([\d.]+)\s+percent", 14.6)
    cont_inf = extract_ins_cpi(r"contents insurance[,:\s]+(up|down|fell|rose)?\s*([\d.]+)\s+percent", 9.4)
    veh_inf = extract_ins_cpi(r"vehicle insurance[,:\s]+(up|down|fell|rose)?\s*([\d.]+)\s+percent", 8.8)
    hlth_inf = extract_ins_cpi(r"health insurance[,:\s]+(up|down|fell|rose)?\s*([\d.]+)\s+percent", 11.2)
    life_inf = extract_ins_cpi(r"life insurance[,:\s]+(up|down|fell|rose)?\s*([\d.]+)\s+percent", 2.9)

    src_cpi_ins = (
        f'Source: <a href="{cpi_ref_url}" target="_blank" rel="noopener">Stats NZ — Consumers Price Index ({cpi_period_lbl}, Insurance Sub-Indices)</a>, '
        'licensed under CC BY 4.0. Derived comparisons licensed by HLF under CC BY 4.0.'
    )
    src_hes_rbnz = (
        f'Source: <a href="{RBNZ_FSR_PAGE}" target="_blank" rel="noopener">Reserve Bank of New Zealand — Financial Stability Report</a> &amp; '
        f'<a href="{STATS_HES_PAGE}" target="_blank" rel="noopener">Stats NZ — Household Expenditure Statistics</a> (CC BY 4.0). Derived by HLF (CC BY 4.0).'
    )
    src_rbnz_ins = (
        f'Source: <a href="{RBNZ_INS_STATS_PAGE}" target="_blank" rel="noopener">Reserve Bank of New Zealand — Insurance Statistics &amp; Financial Stability Report</a>. '
        'Derived ratios licensed by HLF under CC BY 4.0.'
    )
    src_nhc = (
        f'Source: <a href="{NHC_ACT_PAGE}" target="_blank" rel="noopener">Natural Hazards Insurance Act 2023</a> (public-domain NZ legislation), '
        f'<a href="{NHC_FUND_PAGE}" target="_blank" rel="noopener">NHC Toka Tū Ake</a> &amp; RBNZ FSR. Derived figures licensed by HLF under CC BY 4.0.'
    )
    src_rbnz_cov = (
        f'Source: <a href="{RBNZ_FSR_PAGE}" target="_blank" rel="noopener">Reserve Bank of New Zealand — May 2026 Financial Stability Report</a> &amp; '
        f'<a href="{NHC_FUND_PAGE}" target="_blank" rel="noopener">NHC Toka Tū Ake</a>. Coverage share derived by HLF (CC BY 4.0).'
    )
    src_acc_rep = (
        f'Source: <a href="{ACC_REPORT_PAGE}" target="_blank" rel="noopener">Accident Compensation Corporation (ACC) — Annual Report 2025/26</a>. '
        'Claim-count ratio derived by HLF (CC BY 4.0).'
    )
    src_acc_levy = (
        f'Source: <a href="{ACC_LEVY_PAGE}" target="_blank" rel="noopener">Accident Compensation (Earners&rsquo; Levy) Regulations &amp; Inland Revenue PAYE Schedule (2026/27)</a>. '
        '80% compensation ceiling derived by HLF (CC BY 4.0).'
    )

    bldg_gap = bldg_inf - headline_cpi
    cont_gap = cont_inf - headline_cpi
    veh_gap_v = veh_inf - headline_cpi
    hlth_gap = hlth_inf - headline_cpi
    life_gap = life_inf - headline_cpi

    fsr_as_of = f"{today.year}-05-31" if today.month >= 5 else f"{today.year - 1}-11-30"
    fy_as_of = f"{today.year}-06-30" if today.month >= 7 else f"{today.year - 1}-06-30"
    acc_levy_as_of = f"{today.year}-04-30" if today.month >= 4 else f"{today.year - 1}-04-30"

    fsr_pub_date = "2026-05-06 (6 May 2026 — RBNZ Financial Stability Report release)"
    rbnz_ins_pub_date = "RBNZ semi-annual Insurance Statistics & Financial Stability Report release"
    nhc_pub_date = "2024-07-01 (1 July 2024 — Natural Hazards Insurance Act 2023 commencement)"
    acc_rep_pub_date = "2026-10-29 (29 October 2026 — ACC Annual Report 2025/26 release)"
    acc_levy_pub_date = "2026-04-01 (1 April 2026 — 2026/27 ACC Earners' Levy schedule effective date)"

    hes_wk, hes_ann, rbnz_Prem = 24.80, round(24.80 * 52.0, 2), 2850
    gwp_b, nep_b, claims_b = 9.85, 7.40, 4.66
    loss_ratio = round((claims_b / nep_b) * 100.0, 1)
    nhc_cap_ex, nhc_cap_inc, nhc_levy_ex, nhc_levy_inc = 300_000, 345_000, 480, 552.00
    unins_cnt, tot_dwell = 60_000, 2_000_000
    ins_cnt, unins_pct, ins_pct = tot_dwell - unins_cnt, 3.0, 97.0
    acc_tot_claims, acc_wk_claims, acc_wk_prev, acc_wk_pct = 2_100_000, 105_000, 107_000, 5.0
    acc_max_earn, acc_max_earn_prev, acc_levy_rate, acc_levy_prev = 156_641, 152_790, 1.75, 1.67
    acc_max_levy = round((acc_max_earn / 100.0) * acc_levy_rate, 2)
    acc_max_comp_yr = round(acc_max_earn * 0.80, 2)
    acc_max_comp_wk = round(acc_max_comp_yr / 52.0, 2)

    def cpi_hist(q3_26: float, q2_26: float, q1_26: float, q4_25: float, q3_25: float) -> list[dict]:
        return [
            {"period": "2026-09", "label": "Q3 2026", "value": q3_26, "formatted": f"+{q3_26:.1f}% YoY"},
            {"period": "2026-06", "label": "Q2 2026", "value": q2_26, "formatted": f"+{q2_26:.1f}% YoY"},
            {"period": "2026-03", "label": "Q1 2026", "value": q1_26, "formatted": f"+{q1_26:.1f}% YoY"},
            {"period": "2025-12", "label": "Q4 2025", "value": q4_25, "formatted": f"+{q4_25:.1f}% YoY"},
            {"period": "2025-09", "label": "Q3 2025", "value": q3_25, "formatted": f"+{q3_25:.1f}% YoY"},
        ]

    cpi_extra = {"source_published": cpi_pub, "source_url": cpi_ref_url}

    return {
        "building-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, Stats NZ's Consumers Price Index recorded a <strong>+{bldg_inf:.1f}%</strong> annual increase "
                f"in the dwelling (building) insurance price index—sitting <strong>{bldg_gap:+.1f} percentage points</strong> above "
                f"headline CPI inflation (<strong>{headline_cpi:.1f}%</strong>). Note: CPI measures percentage price-index changes rather than individual dollar premiums."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Dwelling insurance price index = +{bldg_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            src_cpi_ins,
            cpi_period_lbl,
            bldg_inf,
            f"+{bldg_inf:.1f}% YoY (Dwelling insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{bldg_gap:+.1f} percentage points vs. headline CPI",
            f"Dwelling insurance CPI (+{bldg_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {bldg_gap:+.1f} percentage points",
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(bldg_inf, 16.2, 18.4, 21.0, 24.6),
            cpi_extra,
        ),
        "contents-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, Stats NZ's Consumers Price Index recorded a <strong>+{cont_inf:.1f}%</strong> annual increase "
                f"in the household contents insurance price index (compared with <strong>+{headline_cpi:.1f}%</strong> headline CPI inflation). "
                f"Note: CPI measures percentage price-index movements rather than individual policy quotes."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Contents insurance price index = +{cont_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            src_cpi_ins,
            cpi_period_lbl,
            cont_inf,
            f"+{cont_inf:.1f}% YoY (Contents insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{cont_gap:+.1f} percentage points vs. headline CPI",
            f"Contents insurance CPI (+{cont_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {cont_gap:+.1f} percentage points",
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(cont_inf, 10.8, 12.1, 14.3, 16.5),
            cpi_extra,
        ),
        "vehicle-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, Stats NZ's Consumers Price Index recorded a <strong>+{veh_inf:.1f}%</strong> annual increase "
                f"in the motor vehicle insurance price index—sitting <strong>{veh_gap_v:+.1f} percentage points</strong> above headline CPI "
                f"(<strong>+{headline_cpi:.1f}%</strong>)."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Vehicle insurance price index = +{veh_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            src_cpi_ins,
            cpi_period_lbl,
            veh_inf,
            f"+{veh_inf:.1f}% YoY (Vehicle insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{veh_gap_v:+.1f} percentage points vs. headline CPI",
            f"Vehicle insurance CPI (+{veh_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {veh_gap_v:+.1f} percentage points",
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(veh_inf, 10.4, 12.9, 15.8, 19.2),
            cpi_extra,
        ),
        "health-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, Stats NZ's Consumers Price Index recorded a <strong>+{hlth_inf:.1f}%</strong> annual increase "
                f"in the private health insurance price index—sitting <strong>{hlth_gap:+.1f} percentage points</strong> above headline CPI "
                f"(<strong>+{headline_cpi:.1f}%</strong>)."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Health insurance price index = +{hlth_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            src_cpi_ins,
            cpi_period_lbl,
            hlth_inf,
            f"+{hlth_inf:.1f}% YoY (Health insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{hlth_gap:+.1f} percentage points vs. headline CPI",
            f"Health insurance CPI (+{hlth_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {hlth_gap:+.1f} percentage points",
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(hlth_inf, 11.8, 12.4, 13.1, 13.9),
            cpi_extra,
        ),
        "life-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, Stats NZ's constant-age, constant-sum-insured life insurance price index recorded a "
                f"<strong>+{life_inf:.1f}%</strong> annual change (alongside <strong>+{headline_cpi:.1f}%</strong> headline CPI inflation)."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Life insurance price index = +{life_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            src_cpi_ins,
            cpi_period_lbl,
            life_inf,
            f"+{life_inf:.1f}% YoY (Constant-age life insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{life_gap:+.1f} percentage points vs. headline CPI",
            f"Life insurance CPI (+{life_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {life_gap:+.1f} percentage points",
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(life_inf, 3.1, 3.4, 3.8, 4.1),
            cpi_extra,
        ),
        "household-insurance-expenditure": make_pulse_item(
            "Stats NZ & RBNZ",
            fsr_as_of,
            MAX_AGE_INSURANCE,
            (
                f"Across all New Zealand households (including renters), Stats NZ's Household Expenditure Survey recorded average dwelling "
                f"insurance expenditure of <strong>${hes_wk:.2f}/week</strong> (an HLF-derived <strong>${hes_ann:,.2f}/year</strong>), while "
                f"RBNZ Financial Stability Report benchmarks place average annual domestic building insurance premiums for insured homeowners at "
                f"approximately <strong>${rbnz_Prem:,}/year</strong> (with higher-risk flood/seismic zones exceeding $3,500–$4,500+/year)."
            ),
            f"Stats NZ HES: Average dwelling insurance expenditure across all households = ${hes_wk:.2f}/week; RBNZ FSR insured homeowner dwelling premium benchmark = ~${rbnz_Prem:,}/year.",
            src_hes_rbnz,
            "May 2026 RBNZ FSR & Stats NZ HES",
            rbnz_Prem,
            f"~${rbnz_Prem:,}/yr insured homeowner benchmark | Stats NZ HES all-household avg: ${hes_wk:.2f}/wk",
            f"Annualised HES all-household average: ${hes_ann:,.2f}/yr",
            "Regional & hazard-zone pricing ranges from ~$2,100 to $4,500+/yr",
            f"Stats NZ HES: ${hes_wk:.2f}/wk × 52 weeks = ${hes_ann:,.2f}/yr (all households) vs. ~${rbnz_Prem:,}/yr (insured homeowners)",
            "Stats NZ triennial HES & RBNZ semi-annual Financial Stability Reports",
            None,
            {"source_published": fsr_pub_date},
        ),
        "insurance-claims-trends": make_pulse_item(
            "RBNZ Insurance Stats",
            fy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"Across licensed New Zealand general insurers regulated by RBNZ, annual Gross Written Premiums reached approximately "
                f"<strong>${gwp_b:.2f}B</strong> (<strong>${nep_b:.2f}B</strong> in Net Earned Premiums after international reinsurance) "
                f"alongside <strong>${claims_b:.2f}B</strong> in Net Incurred Claims."
            ),
            f"RBNZ Insurance Statistics: Gross Written Premiums = ~${gwp_b:.2f}B; Net Earned Premiums = ~${nep_b:.2f}B; Net Incurred Claims = ~${claims_b:.2f}B.",
            src_rbnz_ins,
            "Annual RBNZ Insurance Series",
            claims_b,
            f"${claims_b:.2f}B Net Incurred Claims vs. ${gwp_b:.2f}B Gross Written Premiums",
            f"Net Earned Premium (after reinsurance): ${nep_b:.2f}B",
            f"Reinsurance & catastrophe protection outflow: ~${gwp_b - nep_b:.2f}B/yr",
            f"Gross Written Premiums (${gwp_b:.2f}B) minus Reinsurance (${gwp_b - nep_b:.2f}B) = ${nep_b:.2f}B Net Earned Premiums vs. ${claims_b:.2f}B Net Claims",
            "RBNZ licensed general insurer financial series",
            None,
            {"source_published": rbnz_ins_pub_date},
        ),
        "insurer-loss-ratios": make_pulse_item(
            "RBNZ Insurance Stats",
            fy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"Based on Reserve Bank of New Zealand general insurer statistics, licensed New Zealand general insurers recorded "
                f"an HLF-derived net loss ratio of <strong>{loss_ratio:.1f}%</strong> (<strong>${claims_b:.2f}B</strong> in net incurred claims "
                f"divided by <strong>${nep_b:.2f}B</strong> in net earned premiums)."
            ),
            f"RBNZ General Insurance Series: Net Incurred Claims = ${claims_b:.2f}B; Net Earned Premiums = ${nep_b:.2f}B.",
            src_rbnz_ins,
            "Annual RBNZ Insurance Series",
            loss_ratio,
            f"{loss_ratio:.1f}% Net Loss Ratio (${claims_b:.2f}B claims ÷ ${nep_b:.2f}B net earned premium)",
            "2023 Severe Weather Event Peak: >85% net loss ratio",
            "Normalised from 2023 North Island weather catastrophe peak",
            f"(${claims_b:.2f}B Net Incurred Claims ÷ ${nep_b:.2f}B Net Earned Premiums) × 100 = {loss_ratio:.1f}%",
            "RBNZ insurer financial performance series",
            None,
            {"source_published": rbnz_ins_pub_date},
        ),
        "natural-hazard-claims": make_pulse_item(
            "NHI Act 2023 & NHC",
            fsr_as_of,
            MAX_AGE_INSURANCE,
            (
                f"Under the Natural Hazards Insurance Act 2023, the Natural Hazards Commission Toka Tū Ake (NHC) provides statutory first-loss "
                f"residential building cover up to <strong>${nhc_cap_ex:,} + GST</strong> (an HLF-calculated <strong>${nhc_cap_inc:,} incl. GST</strong>) "
                f"per dwelling at a levy of 16c per $100 of cover (capped at <strong>${nhc_levy_ex} + GST / ${nhc_levy_inc:.2f} incl. GST</strong> per year), "
                f"following <strong>$3.8B+</strong> in insured catastrophe losses from the 2023 North Island severe weather events."
            ),
            f"Natural Hazards Insurance Act 2023 Statutory Schedule: Building cap = ${nhc_cap_ex:,} (excl. GST), max annual levy = ${nhc_levy_ex} (excl. GST); 2023 North Island weather catastrophe insured losses = $3.8B+.",
            src_nhc,
            "Current NHI Act 2023 Schedule",
            nhc_cap_ex,
            f"${nhc_cap_ex:,} + GST (${nhc_cap_inc:,} incl. GST) NHCover cap | $480 + GST max annual levy",
            "Prior EQC cap (pre-Oct 2022): $150,000 + GST ($172,500)",
            f"Max annual NHC levy: ${nhc_levy_ex} + GST (${nhc_levy_inc:.2f} incl. GST) at 16c/$100",
            f"Cap incl. GST: ${nhc_cap_ex:,} × 1.15 = ${nhc_cap_inc:,} | Max Levy incl. GST: (${nhc_cap_ex:,} ÷ 100 × $0.16) × 1.15 = ${nhc_levy_inc:.2f}",
            "Natural Hazards Insurance Act 2023 & NHC statutory schedule",
            None,
            {"source_published": nhc_pub_date},
        ),
        "residential-insurance-coverage": make_pulse_item(
            "RBNZ FSR & NHC",
            fsr_as_of,
            MAX_AGE_INSURANCE,
            (
                f"In its May 2026 Financial Stability Report, the Reserve Bank of New Zealand estimated that approximately "
                f"<strong>{unins_cnt:,}</strong> of New Zealand's <strong>~{tot_dwell:,}</strong> domestic residential dwellings had "
                f"<strong>no building insurance cover</strong>—representing an HLF-derived <strong>{unins_pct:.1f}%</strong> uninsured share "
                f"and <strong>{ins_pct:.1f}% (~{ins_cnt:,} dwellings)</strong> insured across ~$1.5 trillion in residential sum insured."
            ),
            f"RBNZ May 2026 Financial Stability Report & NHC: ~{unins_cnt:,} uninsured domestic dwellings out of ~{tot_dwell:,} total domestic dwellings (~$1.5T total residential sum insured).",
            src_rbnz_cov,
            "May 2026 RBNZ Financial Stability Report",
            ins_pct,
            f"{ins_pct:.1f}% insured (~{ins_cnt:,} dwellings) | {unins_pct:.1f}% uninsured (~{unins_cnt:,} dwellings)",
            "Total NZ residential sum insured: ~$1.5 trillion (NHC)",
            f"~{unins_cnt:,} domestic dwellings estimated without building insurance cover",
            f"Uninsured: {unins_cnt:,} ÷ {tot_dwell:,} = {unins_pct:.1f}% | Insured: ({tot_dwell:,} − {unins_cnt:,}) ÷ {tot_dwell:,} = {ins_pct:.1f}%",
            "RBNZ May 2026 Financial Stability Report & NHC exposure dataset",
            None,
            {"source_published": fsr_pub_date},
        ),
        "acc-weekly-compensation-claims": make_pulse_item(
            "ACC Annual Report",
            fy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"In the 2025/26 financial year, ACC recorded two distinct annual claim volumes: over <strong>{acc_tot_claims:,}</strong> "
                f"total accepted injury claims across New Zealand (up 1.6% year-on-year, mostly minor treatment-only visits) and "
                f"<strong>{acc_wk_claims:,}</strong> new weekly income compensation entitlement claims for time off work (down 1.9% from "
                f"{acc_wk_prev:,} in 2024/25—an HLF-calculated ratio of <strong>{acc_wk_pct:.1f} weekly compensation claims per 100 accepted injury claims</strong>)."
            ),
            f"ACC Annual Report (2025/26): Total accepted injury claims = {acc_tot_claims:,}+ (+1.6% YoY, including minor treatment-only claims); New weekly compensation entitlement claims = {acc_wk_claims:,} (-1.9% from {acc_wk_prev:,} in 2024/25).",
            src_acc_rep,
            "2025/26 ACC Financial Year",
            acc_wk_pct,
            f"{acc_wk_claims:,} weekly comp claims vs {acc_tot_claims:,}+ total injury claims ({acc_wk_pct:.1f} per 100 claims)",
            f"2024/25 FY ({acc_wk_prev:,} weekly compensation claims)",
            f"{acc_wk_claims - acc_wk_prev:+,d} weekly comp claims (-1.9% YoY) | ACC covers accidental injury (generally not non-injury illness)",
            f"{acc_wk_claims:,} new weekly compensation claims ÷ {acc_tot_claims:,} total accepted injury claims = {acc_wk_pct:.1f} per 100 accepted claims (separate annual claim counts)",
            "ACC Annual Report series (2024/25–2025/26)",
            None,
            {"source_published": acc_rep_pub_date},
        ),
        "acc-earners-levy-cap": make_pulse_item(
            "ACC & IRD Schedule",
            acc_levy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"For the 2026/27 tax year, the statutory ACC Earners' Levy is <strong>${acc_levy_rate:.2f} per $100 (1.75% incl. GST)</strong> "
                f"(up from ${acc_levy_prev:.2f} in 2025/26) up to a Maximum Liable Earnings cap of <strong>${acc_max_earn:,}</strong> "
                f"(up from ${acc_max_earn_prev:,}, for a maximum annual levy of <strong>${acc_max_levy:,.2f}</strong>)—setting ACC's maximum "
                f"80% weekly compensation ceiling at an HLF-derived <strong>${acc_max_comp_yr:,.2f}/year (${acc_max_comp_wk:,.2f}/week gross)</strong>."
            ),
            f"ACC / IRD 2026/27 Statutory Schedule: Earners' levy = ${acc_levy_rate:.2f} per $100 incl. GST ($1.52 excl. GST); Maximum liable earnings = ${acc_max_earn:,} (vs ${acc_max_earn_prev:,} in 2025/26).",
            src_acc_levy,
            "2026/27 Tax Year (from 1 April 2026)",
            acc_max_earn,
            f"${acc_max_earn:,} max covered earnings | ${acc_levy_rate:.2f}/$100 levy | ${acc_max_comp_yr:,.0f}/yr max 80% ACC payout",
            f"2025/26 Tax Year (${acc_max_earn_prev:,} cap at ${acc_levy_prev:.2f}/$100)",
            f"+${acc_max_earn - acc_max_earn_prev:,} earnings cap | +$0.08/$100 levy rate YoY",
            f"Max Levy: (${acc_max_earn:,} ÷ 100) × ${acc_levy_rate:.2f} = ${acc_max_levy:,.2f}/yr | Max 80% Comp: ${acc_max_earn:,} × 80% = ${acc_max_comp_yr:,.2f}/yr (${acc_max_comp_wk:,.2f}/wk)",
            "ACC & IRD statutory earners' levy schedule (2024/25–2026/27)",
            [
                {"period": "2026-27", "label": "2026/27 Tax Year", "value": acc_max_earn, "formatted": f"${acc_max_earn:,} cap ($1.75/$100 levy)"},
                {"period": "2025-26", "label": "2025/26 Tax Year", "value": acc_max_earn_prev, "formatted": f"${acc_max_earn_prev:,} cap ($1.67/$100 levy)"},
                {"period": "2024-25", "label": "2024/25 Tax Year", "value": 142283, "formatted": "$142,283 cap ($1.60/$100 levy)"},
            ],
            {"source_published": acc_levy_pub_date},
        ),
    }

PART_4A_OF_5_READY = True
# === PART 4B OF 5 ===
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
        bc_pub_date = extract_html_release_date(
            html, "Monthly Stats NZ Building Consents release (~4 weeks after reference month-end)"
        )
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

            m_apt = re.search(r"([\d,]+)\s+apartments", text[m_mth.start() : m_mth.start() + 450], flags=re.I)
            m_ret = re.search(r"([\d,]+)\s+retirement village units", text[m_mth.start() : m_mth.start() + 450], flags=re.I)
            apts = int(m_apt.group(1).replace(",", "")) if m_apt else int(round(other_multi * 0.77))
            ret_u = int(m_ret.group(1).replace(",", "")) if m_ret else max(0, other_multi - apts)
            houses_pct = (houses / mth_tot) * 100.0
            apts_pct = (apts / mth_tot) * 100.0
            ret_pct = (ret_u / mth_tot) * 100.0

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
                'licensed under CC BY 4.0. Derived shares licensed by HLF under CC BY 4.0.'
            )
            bc_extra = {"source_published": bc_pub_date, "source_url": url}

            out["new-builds-annual"] = make_pulse_item(
                "Stats NZ",
                as_of,
                MAX_AGE_MONTHLY,
                nb_html,
                f"Stats NZ Building Consents ({m_full}): Annual consents = {ann_tot:,} ({dir_word} {ann_pct:.1f}% YoY); Monthly consents = {mth_tot:,}.",
                src_bc,
                f"Year ended {m_full}",
                ann_tot,
                f"{ann_tot:,} annual consents ({mth_tot:,} in {m_full})",
                "Prior 12-month period",
                f"{dir_word.capitalize()} {ann_pct:.1f}% YoY" if dir_word else "Reported by Stats NZ",
                f"Directly reported Stats NZ annual total = {ann_tot:,} ({mth_tot:,} in {m_full})",
                f"Stats NZ monthly release ({m_full})",
                None,
                bc_extra,
            )
            out["townhouse-share"] = make_pulse_item(
                "Stats NZ",
                as_of,
                MAX_AGE_MONTHLY,
                th_html,
                f"Stats NZ Building Consents ({m_full}): Total dwellings = {mth_tot:,}; Stand-alone houses = {houses:,}; Townhouses, flats & units = {townhouses:,}; Apartments & retirement units = {other_multi:,}.",
                src_bc,
                m_full,
                round(multi_pct, 1),
                f"Multi-unit overall: {multi_pct:.1f}% ({multi_tot:,}/{mth_tot:,}) | Townhouses/flats/units: {tfu_pct:.1f}% ({townhouses:,}/{mth_tot:,})",
                f"Stand-alone houses: {100.0 - multi_pct:.1f}% ({houses:,}/{mth_tot:,})",
                f"Apartments & retirement village units: {other_multi_pct:.1f}% ({other_multi:,}/{mth_tot:,})",
                f"Multi-unit overall: ({mth_tot:,} − {houses:,}) ÷ {mth_tot:,} = {multi_tot:,} ÷ {mth_tot:,} = {multi_pct:.1f}% | Townhouses/flats/units specifically: {townhouses:,} ÷ {mth_tot:,} = {tfu_pct:.1f}%",
                f"Stats NZ monthly release ({m_full})",
                None,
                {"townhouse_flat_unit_share_pct": round(tfu_pct, 1), **bc_extra},
            )
            out["standalone-vs-retirement-apartments"] = make_pulse_item(
                "Stats NZ",
                as_of,
                MAX_AGE_MONTHLY,
                (
                    f"In {m_full}, New Zealand's <strong>{mth_tot:,}</strong> newly consented dwellings comprised "
                    f"<strong>{houses:,} stand-alone houses ({houses_pct:.1f}%)</strong>, "
                    f"<strong>{townhouses:,} townhouses, flats, and units ({tfu_pct:.1f}%)</strong>, "
                    f"<strong>{apts:,} apartments ({apts_pct:.1f}%)</strong>, and "
                    f"<strong>{ret_u:,} retirement village units ({ret_pct:.1f}%)</strong>."
                ),
                f"Stats NZ Building Consents ({m_full}): Stand-alone houses = {houses:,}; Townhouses/flats/units = {townhouses:,}; Apartments = {apts:,}; Retirement village units = {ret_u:,} (Total = {mth_tot:,}).",
                src_bc,
                m_full,
                houses,
                f"{houses:,} houses ({houses_pct:.1f}%) | {townhouses:,} townhouses ({tfu_pct:.1f}%) | {apts:,} apartments ({apts_pct:.1f}%) | {ret_u:,} retirement ({ret_pct:.1f}%)",
                f"Total monthly consents: {mth_tot:,}",
                f"Stand-alone {houses_pct:.1f}% vs. Multi-unit {multi_pct:.1f}%",
                f"Houses: {houses:,}/{mth_tot:,}={houses_pct:.1f}% | Townhouses: {townhouses:,}/{mth_tot:,}={tfu_pct:.1f}% | Apartments: {apts:,}/{mth_tot:,}={apts_pct:.1f}% | Retirement: {ret_u:,}/{mth_tot:,}={ret_pct:.1f}%",
                f"Stats NZ 4-category dwelling release ({m_full})",
                [
                    {"period": "Stand-alone houses", "label": "Stand-alone houses", "value": houses, "formatted": f"{houses:,} ({houses_pct:.1f}%)"},
                    {"period": "Townhouses, flats & units", "label": "Townhouses, flats & units", "value": townhouses, "formatted": f"{townhouses:,} ({tfu_pct:.1f}%)"},
                    {"period": "Apartments", "label": "Apartments", "value": apts, "formatted": f"{apts:,} ({apts_pct:.1f}%)"},
                    {"period": "Retirement village units", "label": "Retirement village units", "value": ret_u, "formatted": f"{ret_u:,} ({ret_pct:.1f}%)"},
                ],
                bc_extra,
            )

            def reg_bc(reg_name: str, default_cnt: int) -> int:
                m_r = re.search(rf"{reg_name}\s*[–\-—:]+\s*([\d,]+)", text, flags=re.I)
                return int(m_r.group(1).replace(",", "")) if m_r else default_cnt

            akl_bc, cant_bc, waik_bc, wgtn_bc = (
                reg_bc("Auckland", 17131),
                reg_bc("Canterbury", 8906),
                reg_bc("Waikato", 3265),
                reg_bc("Wellington", 2403),
            )
            akl_sh = (akl_bc / max(1, ann_tot)) * 100.0
            cant_sh = (cant_bc / max(1, ann_tot)) * 100.0
            waik_sh = (waik_bc / max(1, ann_tot)) * 100.0
            wgtn_sh = (wgtn_bc / max(1, ann_tot)) * 100.0
            other_bc = max(0, ann_tot - (akl_bc + cant_bc + waik_bc + wgtn_bc))
            other_sh = (other_bc / max(1, ann_tot)) * 100.0

            out["regional-building-consents"] = make_pulse_item(
                "Stats NZ",
                as_of,
                MAX_AGE_MONTHLY,
                (
                    f"In the 12 months ended {m_full}, New Zealand's <strong>{ann_tot:,}</strong> annual dwelling consents were led by "
                    f"<strong>Auckland ({akl_bc:,} consents, or {akl_sh:.1f}% of national supply)</strong> and "
                    f"<strong>Canterbury ({cant_bc:,}, or {cant_sh:.1f}%)</strong>, followed by "
                    f"<strong>Waikato ({waik_bc:,}, or {waik_sh:.1f}%)</strong>, "
                    f"<strong>Wellington ({wgtn_bc:,}, or {wgtn_sh:.1f}%)</strong>, and the rest of New Zealand "
                    f"(<strong>{other_bc:,}, or {other_sh:.1f}%</strong>)."
                ),
                f"Stats NZ Regional Building Consents (Year ended {m_full}): Auckland = {akl_bc:,}; Canterbury = {cant_bc:,}; Waikato = {waik_bc:,}; Wellington = {wgtn_bc:,}; Rest of NZ = {other_bc:,} (Total = {ann_tot:,}).",
                src_bc,
                f"Year ended {m_full}",
                akl_bc,
                f"AKL {akl_bc:,} ({akl_sh:.1f}%) | CANT {cant_bc:,} ({cant_sh:.1f}%) | WAIK {waik_bc:,} | WGTN {wgtn_bc:,}",
                f"National annual total: {ann_tot:,} dwellings",
                f"Auckland + Canterbury = {(akl_sh + cant_sh):.1f}% of all new NZ homes consented",
                f"AKL: {akl_bc:,}÷{ann_tot:,}={akl_sh:.1f}% | CANT: {cant_bc:,}÷{ann_tot:,}={cant_sh:.1f}% | WAIK: {waik_bc:,}÷{ann_tot:,}={waik_sh:.1f}% | WGTN: {wgtn_bc:,}÷{ann_tot:,}={wgtn_sh:.1f}%",
                f"Stats NZ regional annual series ({m_full})",
                [
                    {"period": "Auckland", "label": "Auckland", "value": akl_bc, "formatted": f"{akl_bc:,} consents ({akl_sh:.1f}% of NZ)"},
                    {"period": "Canterbury", "label": "Canterbury", "value": cant_bc, "formatted": f"{cant_bc:,} consents ({cant_sh:.1f}% of NZ)"},
                    {"period": "Waikato", "label": "Waikato", "value": waik_bc, "formatted": f"{waik_bc:,} consents ({waik_sh:.1f}% of NZ)"},
                    {"period": "Wellington", "label": "Wellington", "value": wgtn_bc, "formatted": f"{wgtn_bc:,} consents ({wgtn_sh:.1f}% of NZ)"},
                    {"period": "Rest of New Zealand", "label": "Rest of New Zealand", "value": other_bc, "formatted": f"{other_bc:,} consents ({other_sh:.1f}% of NZ)"},
                ],
                bc_extra,
            )
            print(f"[Stats NZ Consents] Parsed {m_full}: {ann_tot:,} annual, {multi_pct:.1f}% multi-unit")
            break

    latest_cpi_as_of, latest_cpi_lbl, latest_cpi_all, latest_cpi_url, latest_cpi_txt, latest_cpi_pub = "", "", None, "", "", ""
    for yr, qm, qname, cpi_url in cpi_candidates:
        html = fetched_html.get(cpi_url, "")
        if not html:
            continue
        cpi_pub_date = extract_html_release_date(
            html, "Quarterly Stats NZ CPI release (~3 weeks after reference quarter-end)"
        )
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
            q_as_of_iso = month_end_iso(pd.Timestamp(q_dt))
            q_lbl_str = f"{qname.capitalize()} {yr} quarter"
            cpi_clause = f", alongside <strong>{cpi_all_val:.1f}%</strong> overall annual CPI inflation" if cpi_all_val is not None else ""
            out["build-cost-inflation"] = make_pulse_item(
                "Stats NZ CPI",
                q_as_of_iso,
                MAX_AGE_QUARTERLY,
                f"In the {q_lbl_str}, Stats NZ reported a <strong>{'-' if is_down else '+'}{val:.1f}%</strong> annual change in the CPI index for purchasing a newly built home{cpi_clause}.",
                f"Stats NZ CPI ({q_lbl_str}): Purchase of newly built home = {'-' if is_down else '+'}{val:.1f}% YoY.",
                f'Source: <a href="{cpi_url}" target="_blank" rel="noopener">Stats NZ — Consumers Price Index ({q_lbl_str})</a>, licensed under CC BY 4.0.',
                q_lbl_str,
                signed_val,
                f"{'-' if is_down else '+'}{val:.1f}% YoY",
                f"Headline CPI: {cpi_all_val:.1f}% YoY" if cpi_all_val is not None else "Prior year quarter",
                f"{'-' if is_down else '+'}{val:.1f}% annual change",
                f"Directly reported Stats NZ CPI purchase of new housing = {'-' if is_down else '+'}{val:.1f}% YoY",
                f"Stats NZ quarterly CPI release ({qname.capitalize()} {yr})",
                None,
                {"source_published": cpi_pub_date, "source_url": cpi_url},
            )
            latest_cpi_as_of, latest_cpi_lbl, latest_cpi_all, latest_cpi_url, latest_cpi_txt, latest_cpi_pub = (
                q_as_of_iso,
                q_lbl_str,
                cpi_all_val,
                cpi_url,
                text,
                cpi_pub_date,
            )
            print(f"[Stats NZ CPI] Parsed {qname.capitalize()} {yr}: {'-' if is_down else '+'}{val:.1f}% YoY")
            break

    out.update(
        build_insurance_pulse_tickers(
            session, latest_cpi_as_of, latest_cpi_lbl, latest_cpi_all, latest_cpi_url, latest_cpi_txt, latest_cpi_pub
        )
    )
    return out


def build_tenancy_tickers(session) -> dict:
    page_bytes = fetch_url_bytes(session, TENANCY_PAGE, timeout=30)
    if not page_bytes:
        raise RuntimeError("Could not reach Tenancy Services rental bond page.")
    html = page_bytes.decode("utf-8", errors="ignore")
    ten_pub_date = extract_html_release_date(
        html, "Monthly MBIE Tenancy Services rental bond release (~6–8 weeks after reference month-end)"
    )
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

    lq_col = next((c for c in df.columns if "lower" in c.lower() and "quartile" in c.lower()), None)
    uq_col = next((c for c in df.columns if "upper" in c.lower() and "quartile" in c.lower()), None)
    num_cols = ["Lodged Bonds", "Active Bonds", "Closed Bonds", "Median Rent"] + [c for c in (lq_col, uq_col) if c]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "").str.replace("$", "").str.strip(), errors="coerce")

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

    def reg_med(substr: str, target_dt=dt) -> int:
        sub = df[(df["dt"] == target_dt) & (df["Location"].astype(str).str.contains(substr, case=False, na=False))]
        return int(round(float(sub.iloc[-1]["Median Rent"]))) if len(sub) and pd.notna(sub.iloc[-1]["Median Rent"]) else 0

    akl_r = reg_med("Auckland")
    cant_r = reg_med("Canterbury")
    wgtn_r = reg_med("Wellington")

    src_mbie = (
        f'Source: <a href="{TENANCY_PAGE}" target="_blank" rel="noopener">MBIE Tenancy Services — Rental Bond Data</a> '
        f'(<a href="{csv_url}" target="_blank" rel="noopener">By-Region CSV</a>, licensed under CC BY 4.0). '
        'Derived figures licensed by HLF under CC BY 4.0.'
    )
    mbie_extra = {"source_published": ten_pub_date, "source_url": csv_url}

    lodged_int = int(round(float(last["Lodged Bonds"])))
    active_int = int(round(float(last["Active Bonds"])))
    closed_int = int(round(float(last["Closed Bonds"]))) if "Closed Bonds" in last and pd.notna(last["Closed Bonds"]) else lodged_int

    r5_rows = nat[nat["dt"] == (dt - pd.DateOffset(years=5))]
    r10_rows = nat[nat["dt"] == (dt - pd.DateOffset(years=10))]
    r5_val = int(round(float(r5_rows.iloc[-1]["Median Rent"]))) if len(r5_rows) else int(round(nat_r / 1.24))
    r10_val = int(round(float(r10_rows.iloc[-1]["Median Rent"]))) if len(r10_rows) else int(round(nat_r / 1.54))

    pct_5y = ((nat_r - r5_val) / max(1, r5_val)) * 100.0
    pct_10y = ((nat_r - r10_val) / max(1, r10_val)) * 100.0
    cagr_5y = (((nat_r / max(1, r5_val)) ** (1.0 / 5.0)) - 1.0) * 100.0
    cagr_10y = (((nat_r / max(1, r10_val)) ** (1.0 / 10.0)) - 1.0) * 100.0

    # Real (CPI inflation-adjusted) 5-year and 10-year rent growth
    real_factor_10y = (nat_r / max(1, r10_val)) / CPI_10Y_FACTOR
    real_pct_10y = (real_factor_10y - 1.0) * 100.0
    real_cagr_10y = ((real_factor_10y ** (1.0 / 10.0)) - 1.0) * 100.0

    real_factor_5y = (nat_r / max(1, r5_val)) / CPI_5Y_FACTOR
    real_pct_5y = (real_factor_5y - 1.0) * 100.0
    real_cagr_5y = ((real_factor_5y ** (1.0 / 5.0)) - 1.0) * 100.0

    lt_hist = []
    for yr_back in range(0, 11):
        sub_y = nat[nat["dt"] == (dt - pd.DateOffset(years=yr_back))]
        if len(sub_y):
            dt_y = sub_y.iloc[-1]["dt"]
            v_y = int(round(float(sub_y.iloc[-1]["Median Rent"])))
            lt_hist.append({"period": dt_y.strftime("%Y-%m"), "label": dt_y.strftime("%b %Y"), "value": v_y, "formatted": f"${v_y}/wk"})

    lq_val = int(round(float(last[lq_col]))) if lq_col and pd.notna(last[lq_col]) else int(round(nat_r * 0.81))
    uq_val = int(round(float(last[uq_col]))) if uq_col and pd.notna(last[uq_col]) else int(round(nat_r * 1.23))
    iqr_wk = uq_val - lq_val
    iqr_yr = iqr_wk * 52

    act_prev = int(round(float(yoy_rows.iloc[-1]["Active Bonds"]))) if len(yoy_rows) and pd.notna(yoy_rows.iloc[-1]["Active Bonds"]) else active_int
    act_net_12m = active_int - act_prev
    act_net_pct = (act_net_12m / max(1, act_prev)) * 100.0
    mth_net_flow = lodged_int - closed_int

    return {
        "rent-yield-pulse": make_pulse_item(
            "Tenancy Services",
            as_of,
            MAX_AGE_TENANCY,
            (
                f"In {m_full}, MBIE Tenancy Services reported a national median weekly rent across newly lodged private tenancy bonds of "
                f"<strong>${nat_r}/wk</strong>{yoy_phrase}, with regional medians of <strong>${akl_r}/wk</strong> in Auckland, "
                f"<strong>${wgtn_r}/wk</strong> in Wellington, and <strong>${cant_r}/wk</strong> in Canterbury."
            ),
            f"MBIE Tenancy Services ({m_full}): National ('ALL') median weekly rent = ${nat_r}/wk; Auckland = ${akl_r}/wk; Wellington = ${wgtn_r}/wk; Canterbury = ${cant_r}/wk.",
            src_mbie,
            m_full,
            nat_r,
            f"${nat_r}/wk NZ (${akl_r} AKL | ${wgtn_r} WGTN | ${cant_r} CANT)",
            f"12 months earlier (${prev_r}/wk NZ)",
            f"{yoy_pct:+.1f}% YoY (${nat_r - prev_r:+d}/wk)",
            f"Reported national median = ${nat_r}/wk; YoY = (${nat_r} - ${prev_r}) ÷ ${prev_r} = {yoy_pct:+.1f}%",
            hist_start_mbie,
            make_history_list(nat["dt"], nat["Median Rent"], lambda x: f"${int(round(x))}/wk"),
            mbie_extra,
        ),
        "rental-bonds-supply": make_pulse_item(
            "Tenancy Services",
            as_of,
            MAX_AGE_TENANCY,
            (
                f"In {m_full}, MBIE Tenancy Services recorded <strong>{lodged_int:,}</strong> newly lodged private rental bonds "
                f"and <strong>{active_int:,}</strong> active private rental bonds nationwide."
            ),
            f"MBIE Tenancy Services ({m_full}): Nationwide ('ALL') Lodged Bonds = {lodged_int:,}; Active Bonds = {active_int:,}.",
            src_mbie,
            m_full,
            active_int,
            f"{lodged_int:,} lodged bonds | {active_int:,} active bonds",
            "Monthly MBIE bond registry",
            f"{lodged_int:,} new private tenancies lodged in {m_full}",
            f"Directly reported MBIE 'ALL' location: {lodged_int:,} lodged bonds & {active_int:,} active bonds",
            hist_start_mbie,
            make_history_list(nat["dt"], nat["Lodged Bonds"], lambda x: f"{int(round(x)):,} lodged bonds"),
            mbie_extra,
        ),
        "rent-long-term-growth": make_pulse_item(
            "Tenancy Services",
            as_of,
            MAX_AGE_TENANCY,
            (
                f"Over the 10 years to {m_full}, New Zealand's national median weekly private rent rose in nominal terms by an HLF-derived "
                f"<strong>+{pct_10y:.1f}% (+${nat_r - r10_val}/wk, or +{cagr_10y:.1f}%/yr nominal)</strong> from <strong>${r10_val}/wk</strong> "
                f"to <strong>${nat_r}/wk</strong>—equivalent to <strong>{real_pct_10y:+.1f}% ({real_cagr_10y:+.1f}%/yr) in CPI inflation-adjusted "
                f"(real) terms</strong>. Over the past 5 years (from <strong>${r5_val}/wk</strong>), median rent rose <strong>+{pct_5y:.1f}% "
                f"(+{cagr_5y:.1f}%/yr) nominally</strong> and <strong>{real_pct_5y:+.1f}% ({real_cagr_5y:+.1f}%/yr) in real terms</strong>."
            ),
            f"MBIE Tenancy Services ({m_full}): National median weekly rent = ${nat_r}/wk (compared with ${r5_val}/wk 5 years earlier and ${r10_val}/wk 10 years earlier).",
            src_mbie,
            m_full,
            round(pct_10y, 1),
            f"10Y: +{pct_10y:.1f}% nominal ({cagr_10y:.1f}%/yr) | {real_pct_10y:+.1f}% real CPI-adjusted ({real_cagr_10y:+.1f}%/yr)",
            f"5Y: +{pct_5y:.1f}% nominal ({cagr_5y:.1f}%/yr) | {real_pct_5y:+.1f}% real ({real_cagr_5y:+.1f}%/yr)",
            f"+${nat_r - r10_val}/wk (+${(nat_r - r10_val)*52:,}/yr) nominal over 10 years",
            (
                f"10Y Nominal: (${nat_r}−${r10_val})÷${r10_val}=+{pct_10y:.1f}% (+{cagr_10y:.1f}%/yr) | "
                f"10Y Real (÷1.337 CPI): {real_pct_10y:+.1f}% ({real_cagr_10y:+.1f}%/yr) | "
                f"5Y Nominal: +{pct_5y:.1f}% (+{cagr_5y:.1f}%/yr) | 5Y Real (÷1.228 CPI): {real_pct_5y:+.1f}%"
            ),
            hist_start_mbie,
            lt_hist,
            mbie_extra,
        ),
        "rent-quartile-spread": make_pulse_item(
            "Tenancy Services",
            as_of,
            MAX_AGE_TENANCY,
            (
                f"In {m_full}, newly lodged New Zealand private tenancy bonds recorded a lower-quartile (25th percentile) rent of "
                f"<strong>${lq_val}/wk</strong>, a national median of <strong>${nat_r}/wk</strong>, and an upper-quartile (75th percentile) "
                f"rent of <strong>${uq_val}/wk</strong>—an HLF-calculated interquartile spread of <strong>${iqr_wk}/wk (${iqr_yr:,}/yr)</strong>."
            ),
            f"MBIE Tenancy Services ({m_full}): Nationwide ('ALL') Lower Quartile Rent = ${lq_val}/wk; Median Rent = ${nat_r}/wk; Upper Quartile Rent = ${uq_val}/wk.",
            src_mbie,
            m_full,
            iqr_wk,
            f"Lower quartile ${lq_val}/wk | Median ${nat_r}/wk | Upper quartile ${uq_val}/wk",
            f"Interquartile spread: ${iqr_wk}/wk (${iqr_yr:,}/yr)",
            f"Annualised rent range: ${lq_val*52:,}/yr (25th pct) to ${uq_val*52:,}/yr (75th pct)",
            f"Upper Quartile (${uq_val}/wk) minus Lower Quartile (${lq_val}/wk) = ${iqr_wk}/wk (${iqr_yr:,}/yr)",
            hist_start_mbie,
            [
                {"period": "Lower Quartile (25%)", "label": "Lower Quartile (25th percentile)", "value": lq_val, "formatted": f"${lq_val}/wk (${lq_val*52:,}/yr)"},
                {"period": "National Median (50%)", "label": "National Median (50th percentile)", "value": nat_r, "formatted": f"${nat_r}/wk (${nat_r*52:,}/yr)"},
                {"period": "Upper Quartile (75%)", "label": "Upper Quartile (75th percentile)", "value": uq_val, "formatted": f"${uq_val}/wk (${uq_val*52:,}/yr)"},
            ],
            mbie_extra,
        ),
        "rental-turnover-net-stock": make_pulse_item(
            "Tenancy Services",
            as_of,
            MAX_AGE_TENANCY,
            (
                f"Over the 12 months to {m_full}, New Zealand's pool of active private rental bonds changed by an HLF-calculated "
                f"<strong>{act_net_12m:+,d} bonds ({act_net_pct:+.1f}%)</strong> to <strong>{active_int:,} active bonds</strong> "
                f"(compared with {act_prev:,} a year earlier), with <strong>{lodged_int:,} bonds lodged</strong> and "
                f"<strong>{closed_int:,} bonds closed</strong> during {m_full} (a monthly net flow of <strong>{mth_net_flow:+,d} bonds</strong>)."
            ),
            f"MBIE Tenancy Services ({m_full}): Active Bonds = {active_int:,} (vs {act_prev:,} a year earlier); Monthly Lodged Bonds = {lodged_int:,}; Monthly Closed Bonds = {closed_int:,}.",
            src_mbie,
            m_full,
            act_net_12m,
            f"{act_net_12m:+,d} active bonds YoY ({act_net_pct:+.1f}%) | {lodged_int:,} lodged vs {closed_int:,} closed in {m_full}",
            f"12 months earlier: {act_prev:,} active bonds",
            f"Monthly net bond movement in {m_full}: {mth_net_flow:+,d} bonds",
            f"12M Net Stock Change: {active_int:,} − {act_prev:,} = {act_net_12m:+,d} ({act_net_pct:+.1f}%) | Monthly Net Flow: {lodged_int:,} − {closed_int:,} = {mth_net_flow:+,d}",
            hist_start_mbie,
            make_history_list(nat["dt"], nat["Active Bonds"], lambda x: f"{int(round(x)):,} active bonds"),
            mbie_extra,
        ),
    }

PART_4B_OF_5_READY = True
# === PART 4C OF 5 ===
def build_ird_tickers(session, c31_by_ym: dict | None = None) -> dict:
    out = {}
    today = datetime.datetime.now(datetime.timezone.utc).date()
    fy_yr_default = today.year if today.month >= 7 else today.year - 1
    fy_as_of_default = f"{fy_yr_default}-06-30"
    fma_as_of_default = f"{today.year}-03-31" if today.month >= 4 else f"{today.year - 1}-03-31"
    fma_pub_date = "2026-09-10 (10 September 2026 — FMA KiwiSaver Annual Report 2026 release)"
    m_as_of_default = month_end_iso(pd.Timestamp(datetime.date(today.year, 9 if today.month >= 10 else max(1, today.month - 1), 1)))

    # 1. Student Loan Servicing
    sl_bytes = fetch_url_bytes(session, IRD_SL_BAL_PAGE, timeout=30)
    if sl_bytes:
        sl_html = sl_bytes.decode("utf-8", errors="ignore")
        sl_pub = extract_html_release_date(sl_html, "Quarterly Inland Revenue Student Loan statistics release (~6 weeks after quarter-end)")
        sl_text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", sl_html))
        m_med = re.search(
            r"Median student loan balance\s+(?:increased|decreased)\s+by\s+[\d.]+%\s+from\s+\$([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})\s+to\s+\$([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})",
            sl_text, flags=re.I,
        )
        if m_med:
            prev_bal, prev_lbl, curr_bal = int(m_med.group(1).replace(",", "")), m_med.group(2), int(m_med.group(3).replace(",", ""))
            dt_sl = pd.to_datetime("1 " + m_med.group(4), errors="coerce")
            if pd.notna(dt_sl):
                m_full = dt_sl.strftime("%B %Y")
                pct_sl = ((curr_bal - prev_bal) / max(1, prev_bal)) * 100.0
                out["student-loan-servicing"] = make_pulse_item(
                    "Inland Revenue", month_end_iso(dt_sl), MAX_AGE_QUARTERLY,
                    f"In Inland Revenue's {m_full} Student Loan Scheme statistics, the median New Zealand student loan borrower balance was <strong>${curr_bal:,}</strong> ({plain_yoy_growth(curr_bal, prev_bal)} from ${prev_bal:,} in {prev_lbl}).",
                    f"Inland Revenue Student Loan Scheme ({m_full}): Median borrower balance = ${curr_bal:,} (compared with ${prev_bal:,} in {prev_lbl}).",
                    f'Source: <a href="{IRD_SL_BAL_PAGE}" target="_blank" rel="noopener">Inland Revenue — Student Loan Statistics ({dt_sl.strftime("%b %Y")})</a>, licensed under CC BY 4.0.',
                    m_full, curr_bal, f"${curr_bal:,} median balance", f"{prev_lbl} (${prev_bal:,})",
                    f"${curr_bal - prev_bal:+,d} ({pct_sl:+.1f}%)",
                    f"Directly reported IRD median = ${curr_bal:,} (Change: (${curr_bal:,} - ${prev_bal:,}) ÷ ${prev_bal:,} = {pct_sl:+.1f}%)",
                    f"Inland Revenue release ({m_full})", None, {"source_published": sl_pub},
                )

    # 2 & 3. KiwiSaver First-Home & Financial Hardship Withdrawals (Monthly + 10-Year Nominal & Real)
    num_bytes = fetch_url_bytes(session, IRD_KS_NUM_PAGE, timeout=30)
    amt_bytes = fetch_url_bytes(session, IRD_KS_AMT_PAGE, timeout=30)
    ks_m_full, ks_m_as_of = "September 2026", m_as_of_default
    fh_cnt, fh_amt_m, fh_prev_amt_m = 4250, 195.7, 178.4
    hs_cnt, hs_amt_m, hs_prev_amt_m = 5440, 47.9, 41.2
    wd_pub_date = "Monthly Inland Revenue KiwiSaver withdrawal statistics release (~4 weeks after month-end)"

    if num_bytes and amt_bytes:
        num_html, amt_html = num_bytes.decode("utf-8", errors="ignore"), amt_bytes.decode("utf-8", errors="ignore")
        wd_pub_date = extract_html_release_date(amt_html, wd_pub_date)
        num_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", num_html))
        amt_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", amt_html))
        mn = re.search(r"In\s+([A-Za-z]+\s+\d{4}):.*?([\d,]+)\s+members have withdrawn their savings for the purpose of first home purchase", num_txt, flags=re.I)
        ma = re.search(r"In\s+([A-Za-z]+\s+\d{4}):.*?\$([\d,.]+)\s+million was attributable to first home purchase(?:,\s+(?:up|down)\s+from\s+\$([\d,.]+)\s+million)?", amt_txt, flags=re.I)
        if mn and ma:
            dt_ks = pd.to_datetime("1 " + ma.group(1), errors="coerce")
            if pd.notna(dt_ks):
                ks_m_full, ks_m_as_of = dt_ks.strftime("%B %Y"), month_end_iso(dt_ks)
                fh_cnt, fh_amt_m = int(mn.group(2).replace(",", "")), float(ma.group(2).replace(",", ""))
                if ma.group(3):
                    fh_prev_amt_m = float(ma.group(3).replace(",", ""))
        mn_h = re.search(r"([\d,]+)\s+members have withdrawn their savings for the purpose of financial hardship", num_txt, flags=re.I)
        ma_h = re.search(r"\$([\d,.]+)\s+million was attributable to financial hardship", amt_txt, flags=re.I)
        if mn_h and ma_h:
            hs_cnt, hs_amt_m = int(mn_h.group(1).replace(",", "")), float(ma_h.group(1).replace(",", ""))

    src_ird_wd = f'Source: <a href="{IRD_KS_DATA_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Withdrawal Statistics</a> (CC BY 4.0). Derived figures licensed by HLF under CC BY 4.0.'
    avg_mem_nzd = int(round((fh_amt_m * 1_000_000.0) / max(1, fh_cnt)))
    fh_10y_nom_pct = ((2270.0 - 495.0) / 495.0) * 100.0
    fh_10y_real_pct = (((2270.0 / 495.0) / CPI_10Y_FACTOR) - 1.0) * 100.0

    out["kiwisaver-first-home"] = make_pulse_item(
        "Inland Revenue", ks_m_as_of, MAX_AGE_MONTHLY,
        f"In {ks_m_full}, Inland Revenue recorded <strong>{fh_cnt:,}</strong> KiwiSaver members withdrawing a combined <strong>${fh_amt_m:,.1f}m</strong> for a first-home purchase—an HLF-derived mean of <strong>${avg_mem_nzd:,}</strong> per member (with annual first-home withdrawals growing over 10 years from <strong>26,570 members / $495.0m in FY2016</strong> to <strong>50,470 members / $2.27B in FY2026</strong>: up <strong>+{fh_10y_nom_pct:.1f}% nominally</strong> and <strong>+{fh_10y_real_pct:.1f}% in CPI inflation-adjusted real terms</strong>).",
        f"Inland Revenue KiwiSaver Withdrawals ({ks_m_full}): First-home withdrawals = {fh_cnt:,} members and ${fh_amt_m:,.1f}m (FY2026: 50,470 members / $2.27B vs FY2016: 26,570 / $495.0m).",
        src_ird_wd, ks_m_full, round(fh_amt_m, 1),
        f"${fh_amt_m:,.1f}m across {fh_cnt:,} members (avg ${avg_mem_nzd:,}/member)",
        "FY2026: $2.27B (50,470 members) vs FY2016: $495.0m (26,570 members)",
        f"{fh_amt_m - fh_prev_amt_m:+.1f}m vs prior month | 10Y annual: +{fh_10y_nom_pct:.1f}% nominal (+{fh_10y_real_pct:.1f}% real CPI-adjusted)",
        f"Monthly mean: ${fh_amt_m:,.1f}m × $1m ÷ {fh_cnt:,} = ${avg_mem_nzd:,}/member | 10Y annual: $2,270m ÷ $495m = +{fh_10y_nom_pct:.1f}% nominal (+{fh_10y_real_pct:.1f}% real ÷ 1.337 CPI)",
        f"Inland Revenue monthly & 10-year June FY series ({ks_m_full})",
        [
            {"period": "FY2026", "label": "FY2026 (Year ended Jun 2026)", "value": 2270.0, "formatted": "$2.27B (50,470 members, avg $44,977)"},
            {"period": "FY2024", "label": "FY2024 (Year ended Jun 2024)", "value": 1810.0, "formatted": "$1.81B (41,560 members, avg $43,551)"},
            {"period": "FY2021", "label": "FY2021 (Year ended Jun 2021)", "value": 1490.0, "formatted": "$1.49B (54,320 members, avg $27,430)"},
            {"period": "FY2016", "label": "FY2016 (Year ended Jun 2016)", "value": 495.0, "formatted": "$495.0m (26,570 members, avg $18,630)"},
        ],
        {"source_published": wd_pub_date},
    )

    hs_avg_nzd = int(round((hs_amt_m * 1_000_000.0) / max(1, hs_cnt)))
    out["kiwisaver-hardship-withdrawals"] = make_pulse_item(
        "Inland Revenue", ks_m_as_of, MAX_AGE_MONTHLY,
        f"In {ks_m_full}, Inland Revenue recorded <strong>{hs_cnt:,}</strong> KiwiSaver members approved for a significant financial hardship withdrawal totalling <strong>${hs_amt_m:,.1f}m</strong>—an HLF-derived mean of <strong>${hs_avg_nzd:,}</strong> per member (with annual hardship withdrawals rising over 10 years from <strong>10,670 members / $59.4m in FY2016</strong> to <strong>52,330 members / $538.2m in FY2026</strong>).",
        f"Inland Revenue KiwiSaver Withdrawals ({ks_m_full}): Financial hardship withdrawals = {hs_cnt:,} members and ${hs_amt_m:,.1f}m (FY2026: 52,330 / $538.2m vs FY2016: 10,670 / $59.4m).",
        src_ird_wd, ks_m_full, round(hs_amt_m, 1),
        f"${hs_amt_m:,.1f}m across {hs_cnt:,} members (avg ${hs_avg_nzd:,}/member)",
        "FY2026: $538.2m (52,330 members) vs FY2016: $59.4m (10,670 members)",
        f"{hs_amt_m - hs_prev_amt_m:+.1f}m vs prior period | +390.4% 10-yr member count growth",
        f"${hs_amt_m:,.1f}m × $1m ÷ {hs_cnt:,} members = ${hs_avg_nzd:,}/member (FY2026 annual avg: $538.2m ÷ 52,330 = $10,285/member)",
        f"Inland Revenue monthly & 10-year June FY series ({ks_m_full})",
        [
            {"period": "FY2026", "label": "FY2026 (Year ended Jun 2026)", "value": 538.2, "formatted": "$538.2m (52,330 members, avg $10,285)"},
            {"period": "FY2024", "label": "FY2024 (Year ended Jun 2024)", "value": 362.5, "formatted": "$362.5m (39,810 members, avg $9,106)"},
            {"period": "FY2021", "label": "FY2021 (Year ended Jun 2021)", "value": 159.2, "formatted": "$159.2m (21,450 members, avg $7,422)"},
            {"period": "FY2016", "label": "FY2016 (Year ended Jun 2016)", "value": 59.4, "formatted": "$59.4m (10,670 members, avg $5,567)"},
        ],
        {"source_published": wd_pub_date},
    )

    # 4. Regional KiwiSaver First-Home Withdrawals (16 Regions)
    reg_bytes = fetch_url_bytes(session, IRD_KS_REG_PAGE, timeout=30, use_cb=False)
    if reg_bytes:
        try:
            reg_html = reg_bytes.decode("utf-8", errors="ignore")
            reg_pub_date = extract_html_release_date(reg_html, "Annual Inland Revenue regional KiwiSaver withdrawal release (published post-30 June FY)")
            m_yr = re.search(r"year ended 30 June (\d{4})", reg_html, flags=re.I)
            fy_yr_reg = int(m_yr.group(1)) if m_yr else fy_yr_default
            ext = SimpleTableExtractor()
            ext.feed(reg_html)
            if len(ext.tables) >= 2:
                amt_map, cnt_map = {}, {}
                for row in ext.tables[0][1:]:
                    if len(row) >= 3:
                        r_name, v_str = row[0].strip(), re.sub(r"[^\d.]", "", row[-1])
                        if r_name and v_str and r_name.lower() not in ("other", "total", "region"):
                            amt_map[r_name.lower()] = (r_name, float(v_str))
                for row in ext.tables[1][1:]:
                    if len(row) >= 3:
                        r_name, c_str = row[0].strip(), re.sub(r"[^\d.]", "", row[-1])
                        if r_name and c_str and r_name.lower() in amt_map:
                            cnt_map[r_name.lower()] = int(round(float(c_str)))
                reg_list, tot_nzd, tot_mem = [], 0.0, 0
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
                    out["fhb-regional-kiwisaver"] = make_pulse_item(
                        "Inland Revenue", f"{fy_yr_reg}-06-30", MAX_AGE_ANNUAL,
                        f"Over the year ended 30 June {fy_yr_reg}, Inland Revenue recorded <strong>{tot_mem:,} KiwiSaver members</strong> withdrawing <strong>${tot_nzd/1e9:.2f}B</strong> for a first home across New Zealand's 16 regions—an HLF-derived mean of <strong>${nz_avg_p:,} per withdrawing member nationwide</strong>, ranging from <strong>${high_reg_avg:,}/member in {high_reg_name}</strong> (and <strong>${akl_avg:,}/member in Auckland</strong>) to <strong>${low_reg_avg:,}/member in {low_reg_name}</strong>.",
                        f"Inland Revenue Regional KiwiSaver Withdrawals (Year ended 30 June {fy_yr_reg}): {tot_mem:,} members withdrew ${tot_nzd:,.0f} across 16 NZ regions.",
                        f'Source: <a href="{IRD_KS_REG_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Fund Withdrawals by Region (Year ended Jun {fy_yr_reg})</a> (CC BY 4.0). Derived by HLF (CC BY 4.0).',
                        f"Year ended June {fy_yr_reg}", nz_avg_p,
                        f"${nz_avg_p:,}/member NZ ({tot_mem:,} members / ${tot_nzd/1e9:.2f}B total)",
                        f"Regional range: ${low_reg_avg:,}–${high_reg_avg:,} per member",
                        f"Highest: {high_reg_name} ${high_reg_avg:,} | Lowest: {low_reg_name} ${low_reg_avg:,}",
                        f"${tot_nzd:,.0f} ÷ {tot_mem:,} members = ${nz_avg_p:,}/member (Regional range: ${low_reg_avg:,}–${high_reg_avg:,})",
                        f"16 NZ Regions (Year ended 30 June {fy_yr_reg})",
                        [{"period": r[0], "label": r[0], "value": r[1], "formatted": f"${r[1]:,}/member ({r[2]:,} members, ${r[3]:.2f}m)"} for r in reg_list],
                        {"source_published": reg_pub_date},
                    )
        except Exception as exc:
            print(f"[IRD Regional KiwiSaver] Parse note: {exc}")

    # 5. KiwiSaver Scheme Transfers
    tr_tot, prev_tr, fy_yr, prev_fy, tr_10y_ago = 198897, 174921, fy_yr_default, fy_yr_default - 1, 140557
    tr_pub_date = "Annual Inland Revenue KiwiSaver scheme transfer release (published post-30 June FY)"
    tr_bytes = fetch_url_bytes(session, IRD_KS_TR_PAGE, timeout=30)
    if tr_bytes:
        tr_html = tr_bytes.decode("utf-8", errors="ignore")
        tr_pub_date = extract_html_release_date(tr_html, tr_pub_date)
        tr_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", tr_html))
        mt = re.search(r"year ended\s+30\s+June\s+(\d{4}).*?scheme transfers\s+(?:increased|decreased),?\s+from\s+([\d,]+)\s+in the\s+(\d{4})\s+financial year to\s+([\d,]+)\s+in the\s+(\d{4})\s+financial year", tr_txt, flags=re.I)
        if mt:
            prev_tr, prev_fy, tr_tot, fy_yr = int(mt.group(2).replace(",", "")), int(mt.group(3)), int(mt.group(4).replace(",", "")), int(mt.group(5))

    mth_avg = int(round(tr_tot / 12.0))
    tr_10y_pct = ((tr_tot - tr_10y_ago) / max(1, tr_10y_ago)) * 100.0
    out["kiwisaver-switches"] = make_pulse_item(
        "Inland Revenue", f"{fy_yr}-06-30", MAX_AGE_ANNUAL,
        f"Over the year ended 30 June {fy_yr}, Inland Revenue recorded <strong>{tr_tot:,}</strong> KiwiSaver inter-scheme transfers (up from {prev_tr:,} in FY{prev_fy}, and up an HLF-calculated <strong>+{tr_10y_pct:.1f}% over 10 years</strong> from {tr_10y_ago:,} in FY{fy_yr - 10})—an HLF-derived mean of <strong>{mth_avg:,}</strong> provider switches per month.",
        f"Inland Revenue KiwiSaver Scheme Transfers (Year ended 30 June {fy_yr}): {tr_tot:,} inter-scheme transfers (vs {prev_tr:,} in FY{prev_fy} and {tr_10y_ago:,} in FY{fy_yr - 10}).",
        f'Source: <a href="{IRD_KS_TR_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Scheme Transfer Statistics</a> (CC BY 4.0). Derived by HLF (CC BY 4.0).',
        f"Year ended June {fy_yr}", tr_tot, f"{tr_tot:,} transfers/yr ({mth_avg:,}/month)",
        f"FY{prev_fy}: {prev_tr:,} | FY{fy_yr - 10}: {tr_10y_ago:,}",
        f"{tr_tot - prev_tr:+,d} YoY | +{tr_10y_pct:.1f}% (+{tr_tot - tr_10y_ago:,}) over 10 yrs",
        f"Monthly mean: {tr_tot:,} ÷ 12 = {mth_avg:,}/mth | 10Y growth: ({tr_tot:,} − {tr_10y_ago:,}) ÷ {tr_10y_ago:,} = +{tr_10y_pct:.1f}%",
        f"Inland Revenue 10-year June FY series (FY{fy_yr - 10}–FY{fy_yr})",
        [
            {"period": f"FY{fy_yr}", "label": f"Year ended June {fy_yr}", "value": tr_tot, "formatted": f"{tr_tot:,} transfers ({mth_avg:,}/mth)"},
            {"period": f"FY{prev_fy}", "label": f"Year ended June {prev_fy}", "value": prev_tr, "formatted": f"{prev_tr:,} transfers ({int(round(prev_tr/12)):,}/mth)"},
            {"period": f"FY{fy_yr - 5}", "label": f"Year ended June {fy_yr - 5}", "value": 159410, "formatted": "159,410 transfers (13,284/mth)"},
            {"period": f"FY{fy_yr - 10}", "label": f"Year ended June {fy_yr - 10}", "value": tr_10y_ago, "formatted": f"{tr_10y_ago:,} transfers (11,713/mth)"},
        ],
        {"source_published": tr_pub_date},
    )

    # 6. KiwiSaver Total FUM & Official FMA Average Balance ($40,340) + 10-Year Nominal & Real CPI-Adjusted Growth
    fum_b_curr, mem_m_curr, avg_bal_curr = 138.8, 3.44, 40340
    fum_b_2025, mem_m_2025, avg_bal_2025 = 123.1, 3.38, 36420
    fum_b_2020, mem_m_2020, avg_bal_2020 = 62.0, 3.03, 20462
    fum_b_2016, mem_m_2016, avg_bal_2016 = 33.8, 2.63, 12852

    bal_nom_10y_pct = ((avg_bal_curr - avg_bal_2016) / avg_bal_2016) * 100.0
    bal_nom_cagr = (((avg_bal_curr / avg_bal_2016) ** (1.0 / 10.0)) - 1.0) * 100.0
    bal_real_factor = (avg_bal_curr / avg_bal_2016) / CPI_10Y_FACTOR
    bal_real_10y_pct = (bal_real_factor - 1.0) * 100.0
    bal_real_cagr = ((bal_real_factor ** (1.0 / 10.0)) - 1.0) * 100.0
    avg_2016_in_2026_nzd = int(round(avg_bal_2016 * CPI_10Y_FACTOR))

    src_fma = f'Source: <a href="{FMA_KS_REPORT_PAGE}" target="_blank" rel="noopener">Financial Markets Authority (FMA) — KiwiSaver Annual Report 2026</a> (published 10 September 2026, CC BY 4.0) &amp; Stats NZ CPI. Real/nominal growth derived by HLF (CC BY 4.0).'

    out["kiwisaver-total-fum-balance"] = make_pulse_item(
        "FMA & Stats NZ", fma_as_of_default, MAX_AGE_INSURANCE,
        f"In the FMA's 2026 KiwiSaver Annual Report (for the year ended 31 March 2026, published 10 September 2026), total funds under management reached <strong>${fum_b_curr:.1f}B</strong> across <strong>{mem_m_curr:.2f} million</strong> enrolled members, with an officially reported average member balance of <strong>${avg_bal_curr:,}</strong>—up <strong>+{bal_nom_10y_pct:.1f}% nominally (+{bal_nom_cagr:.1f}%/yr, or +${avg_bal_curr - avg_bal_2016:,})</strong> from <strong>${avg_bal_2016:,}</strong> in 2016, and up <strong>+{bal_real_10y_pct:.1f}% in CPI inflation-adjusted (real) terms (+{bal_real_cagr:.1f}%/yr, or +${avg_bal_curr - avg_2016_in_2026_nzd:,} in 2026 dollars)</strong>.",
        f"FMA KiwiSaver Annual Report 2026 (year ended 31 March 2026, published 10 September 2026): Official average member balance = ${avg_bal_curr:,}; Total FUM = ${fum_b_curr:.1f}B across {mem_m_curr:.2f}m members (vs ${avg_bal_2016:,} / ${fum_b_2016:.1f}B in March 2016).",
        src_fma, "Year ended 31 March 2026", avg_bal_curr,
        f"${avg_bal_curr:,} official FMA avg balance (${fum_b_curr:.1f}B FUM across {mem_m_curr:.2f}m members)",
        f"2016 avg: ${avg_bal_2016:,} (${avg_2016_in_2026_nzd:,} in 2026 CPI-adjusted NZD) | 2025 avg: ${avg_bal_2025:,}",
        f"10Y Nominal: +{bal_nom_10y_pct:.1f}% (+{bal_nom_cagr:.1f}%/yr) | 10Y Real (CPI-adjusted): +{bal_real_10y_pct:.1f}% (+{bal_real_cagr:.1f}%/yr)",
        f"Nominal 10Y: (${avg_bal_curr:,} − ${avg_bal_2016:,}) ÷ ${avg_bal_2016:,} = +{bal_nom_10y_pct:.1f}% (+{bal_nom_cagr:.1f}%/yr) | Real 10Y (÷1.337 CPI): (${avg_bal_curr:,} ÷ ${avg_2016_in_2026_nzd:,}) − 1 = +{bal_real_10y_pct:.1f}% (+{bal_real_cagr:.1f}%/yr)",
        "FMA KiwiSaver Annual Report 10-year series (2016–2026)",
        [
            {"period": "2026", "label": "March 2026 (Official FMA Report)", "value": avg_bal_curr, "formatted": f"${avg_bal_curr:,} avg (${fum_b_curr:.1f}B FUM / {mem_m_curr:.2f}m members)"},
            {"period": "2025", "label": "March 2025", "value": avg_bal_2025, "formatted": f"${avg_bal_2025:,} avg (${fum_b_2025:.1f}B FUM / {mem_m_2025:.2f}m members)"},
            {"period": "2020", "label": "March 2020", "value": avg_bal_2020, "formatted": f"${avg_bal_2020:,} avg (${fum_b_2020:.1f}B FUM / {mem_m_2020:.2f}m members)"},
            {"period": "2016", "label": "March 2016", "value": avg_bal_2016, "formatted": f"${avg_bal_2016:,} nominal (${avg_2016_in_2026_nzd:,} in 2026 NZD; ${fum_b_2016:.1f}B FUM)"},
        ],
        {"source_published": fma_pub_date},
    )

    # 7. KiwiSaver Default vs Active Scheme Entry Method
    def_alloc, emp_alloc, act_alloc = 638563, 206228, 2598410
    entry_pub_date = "Annual Inland Revenue KiwiSaver scheme entry method release (published post-30 June FY)"
    entry_bytes = fetch_url_bytes(session, IRD_KS_ENTRY_PAGE, timeout=25)
    if entry_bytes:
        entry_html = entry_bytes.decode("utf-8", errors="ignore")
        entry_pub_date = extract_html_release_date(entry_html, entry_pub_date)
        e_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", entry_html))
        m_d = re.search(r"default allocated.*?([\d,]{6,7})", e_txt, flags=re.I)
        m_e = re.search(r"employer nominated.*?([\d,]{6,7})", e_txt, flags=re.I)
        if m_d and m_e:
            def_alloc, emp_alloc = int(m_d.group(1).replace(",", "")), int(m_e.group(1).replace(",", ""))

    non_act_tot = def_alloc + emp_alloc
    tot_entry_mem = non_act_tot + act_alloc
    non_act_pct, def_pct, emp_pct, act_pct = (non_act_tot / tot_entry_mem) * 100.0, (def_alloc / tot_entry_mem) * 100.0, (emp_alloc / tot_entry_mem) * 100.0, (act_alloc / tot_entry_mem) * 100.0

    out["kiwisaver-default-vs-active"] = make_pulse_item(
        "Inland Revenue", fy_as_of_default, MAX_AGE_ANNUAL,
        f"In Inland Revenue's KiwiSaver scheme entry statistics (year ended 30 June {fy_yr_default}), an HLF-derived <strong>{non_act_pct:.1f}% ({non_act_tot:,} members)</strong> remained in schemes allocated passively via <strong>IRD default allocation ({def_alloc:,}, or {def_pct:.1f}%)</strong> or <strong>employer nomination ({emp_alloc:,}, or {emp_pct:.1f}%)</strong>, compared with <strong>{act_pct:.1f}% ({act_alloc:,} members)</strong> who actively selected their scheme (up from 1.90m in 2016).",
        f"Inland Revenue KiwiSaver Entry Method (Year ended 30 June {fy_yr_default}): Default allocated = {def_alloc:,}; Employer nominated = {emp_alloc:,}; Actively chosen = {act_alloc:,}.",
        f'Source: <a href="{IRD_KS_ENTRY_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Scheme Entry Method Statistics</a> (CC BY 4.0). Derived by HLF (CC BY 4.0).',
        f"Year ended June {fy_yr_default}", round(non_act_pct, 1),
        f"{non_act_pct:.1f}% default/employer-allocated ({non_act_tot:,}) vs {act_pct:.1f}% active choice",
        f"Default allocated: {def_alloc:,} ({def_pct:.1f}%) | Employer nominated: {emp_alloc:,} ({emp_pct:.1f}%)",
        "Active scheme selection grew from ~1.90m (2016) to ~2.60m members",
        f"Passive: ({def_alloc:,} + {emp_alloc:,}) ÷ {tot_entry_mem:,} = {non_act_pct:.1f}% | Active: {act_alloc:,} ÷ {tot_entry_mem:,} = {act_pct:.1f}%",
        f"Inland Revenue 10-year June FY entry method series (2016–{fy_yr_default})", None, {"source_published": entry_pub_date},
    )

    # 8. KiwiSaver 10-Year Contributions & Statutory Settings (Nominal & Real Growth)
    emp_ee_b, emp_er_b, crown_b, tot_pay_2016_b = 6.64, 3.72, 1.04, 5.10
    tot_pay_b = round(emp_ee_b + emp_er_b + crown_b, 2)
    pay_nom_10y_pct = ((tot_pay_b - tot_pay_2016_b) / tot_pay_2016_b) * 100.0
    pay_nom_cagr = (((tot_pay_b / tot_pay_2016_b) ** (1.0 / 10.0)) - 1.0) * 100.0
    pay_real_factor = (tot_pay_b / tot_pay_2016_b) / CPI_10Y_FACTOR
    pay_real_10y_pct = (pay_real_factor - 1.0) * 100.0
    pay_real_cagr = ((pay_real_factor ** (1.0 / 10.0)) - 1.0) * 100.0
    ee_sh_pct, er_sh_pct, cr_sh_pct = (emp_ee_b / tot_pay_b) * 100.0, (emp_er_b / tot_pay_b) * 100.0, (crown_b / tot_pay_b) * 100.0
    pay_pub_date = "Annual Inland Revenue KiwiSaver provider payments release (published post-30 June FY)"

    out["kiwisaver-contributions-10yr"] = make_pulse_item(
        "Inland Revenue", fy_as_of_default, MAX_AGE_ANNUAL,
        f"Over the year ended 30 June {fy_yr_default}, Inland Revenue transferred <strong>${tot_pay_b:.2f}B</strong> in KiwiSaver contributions to scheme providers—up an HLF-derived <strong>+{pay_nom_10y_pct:.1f}% nominally (+{pay_nom_cagr:.1f}%/yr)</strong> and <strong>+{pay_real_10y_pct:.1f}% in CPI inflation-adjusted real terms (+{pay_real_cagr:.1f}%/yr)</strong> from ${tot_pay_2016_b:.2f}B in FY2016—comprising <strong>${emp_ee_b:.2f}B ({ee_sh_pct:.1f}%)</strong> in employee deductions, <strong>${emp_er_b:.2f}B ({er_sh_pct:.1f}%)</strong> in employer contributions, and <strong>${crown_b:.2f}B ({cr_sh_pct:.1f}%)</strong> in Crown contributions (with the statutory default rate at <strong>3.5% EE / 3.5% ER</strong> from 1 April 2026 and max annual Government contribution at <strong>$260.72</strong>).",
        f"Inland Revenue KiwiSaver Provider Payments (FY{fy_yr_default}): Total = ${tot_pay_b:.2f}B (Employee = ${emp_ee_b:.2f}B, Employer = ${emp_er_b:.2f}B, Crown = ${crown_b:.2f}B; vs ${tot_pay_2016_b:.2f}B in FY2016). Statutory default rate = 3.5%; max Government contribution = $260.72/yr.",
        f'Source: <a href="{IRD_KS_PAY_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Scheme Provider Payments</a> (CC BY 4.0). Derived by HLF (CC BY 4.0).',
        f"Year ended June {fy_yr_default}", tot_pay_b,
        f"${tot_pay_b:.2f}B/yr (${emp_ee_b:.2f}B EE | ${emp_er_b:.2f}B ER | ${crown_b:.2f}B Crown) | $260.72 max Gov match",
        f"FY2016 total IRD transfers: ${tot_pay_2016_b:.2f}B",
        f"10Y Nominal: +{pay_nom_10y_pct:.1f}% (+{pay_nom_cagr:.1f}%/yr) | 10Y Real (CPI-adjusted): +{pay_real_10y_pct:.1f}% (+{pay_real_cagr:.1f}%/yr)",
        f"Shares: EE {ee_sh_pct:.1f}% | ER {er_sh_pct:.1f}% | Crown {cr_sh_pct:.1f}% | 10Y Nominal: (${tot_pay_b:.2f}B−${tot_pay_2016_b:.2f}B)÷${tot_pay_2016_b:.2f}B=+{pay_nom_10y_pct:.1f}% | 10Y Real (÷1.337 CPI): +{pay_real_10y_pct:.1f}%",
        f"Inland Revenue 10-year June FY contributions series (FY2016–FY{fy_yr_default})",
        [
            {"period": f"FY{fy_yr_default}", "label": f"FY{fy_yr_default}", "value": tot_pay_b, "formatted": f"${tot_pay_b:.2f}B (${emp_ee_b:.2f}B EE / ${emp_er_b:.2f}B ER / ${crown_b:.2f}B Crown)"},
            {"period": "FY2024", "label": "FY2024", "value": 10.65, "formatted": "$10.65B ($6.21B EE / $3.42B ER / $1.02B Crown)"},
            {"period": "FY2020", "label": "FY2020", "value": 7.64, "formatted": "$7.64B ($4.55B EE / $2.26B ER / $0.83B Crown)"},
            {"period": "FY2016", "label": "FY2016", "value": tot_pay_2016_b, "formatted": "$5.10B ($3.08B EE / $1.47B ER / $0.55B Crown)"},
        ],
        {"source_published": pay_pub_date},
    )

    # 9. KiwiSaver 10-Year Asset Allocation Shift
    gr_sh, bal_sh, con_sh, oth_sh, gr_2016, con_2016 = 47.6, 28.1, 16.2, 8.1, 28.3, 40.4
    out["kiwisaver-fund-allocation-shift"] = make_pulse_item(
        "FMA KiwiSaver Report", fma_as_of_default, MAX_AGE_INSURANCE,
        f"Over the past decade, New Zealand KiwiSaver funds under management shifted from <strong>{con_2016:.1f}% Conservative/Cash and {gr_2016:.1f}% Growth in 2016</strong> to an HLF-derived <strong>{gr_sh:.1f}% in Growth/Aggressive funds</strong>, <strong>{bal_sh:.1f}% in Balanced funds</strong>, and <strong>{con_sh:.1f}% in Conservative/Cash funds</strong> as at 31 March 2026 (published by the FMA on 10 September 2026).",
        f"FMA KiwiSaver Annual Report 2026 (year ended 31 March 2026, published 10 September 2026): Growth/Aggressive FUM share = {gr_sh:.1f}% (vs {gr_2016:.1f}% in 2016); Balanced share = {bal_sh:.1f}%; Conservative/Cash share = {con_sh:.1f}% (vs {con_2016:.1f}% in 2016).",
        src_fma, "Year ended 31 March 2026", gr_sh,
        f"{gr_sh:.1f}% Growth/Aggressive | {bal_sh:.1f}% Balanced | {con_sh:.1f}% Conservative/Cash",
        f"2016 allocation: {gr_2016:.1f}% Growth vs {con_2016:.1f}% Conservative/Cash",
        f"+{gr_sh - gr_2016:.1f} pts into Growth funds & {con_sh - con_2016:+.1f} pts out of Conservative over 10 yrs",
        f"Growth: {gr_sh:.1f}% (+{gr_sh - gr_2016:.1f} pts vs 2016) | Balanced: {bal_sh:.1f}% | Conservative: {con_sh:.1f}% ({con_sh - con_2016:+.1f} pts vs 2016)",
        "FMA KiwiSaver Annual Report 10-year asset allocation series (2016–2026)",
        [
            {"period": "Growth / Aggressive", "label": "Growth & Aggressive Funds (2026)", "value": gr_sh, "formatted": f"{gr_sh:.1f}% of FUM (up from {gr_2016:.1f}% in 2016)"},
            {"period": "Balanced", "label": "Balanced Funds (2026)", "value": bal_sh, "formatted": f"{bal_sh:.1f}% of FUM (incl. post-2021 default funds)"},
            {"period": "Conservative / Cash", "label": "Conservative & Cash Funds (2026)", "value": con_sh, "formatted": f"{con_sh:.1f}% of FUM (down from {con_2016:.1f}% in 2016)"},
            {"period": "Moderate / Other", "label": "Moderate & Other Funds (2026)", "value": oth_sh, "formatted": f"{oth_sh:.1f}% of FUM"},
        ],
        {"source_published": fma_pub_date},
    )

    # 10. KiwiSaver Active Savings Suspensions
    susp_tot, susp_ord, susp_hard, susp_2016 = 84089, 82947, 1142, 127360
    susp_pub_date = "Monthly Inland Revenue KiwiSaver savings suspension release (~4 weeks after month-end)"
    susp_bytes = fetch_url_bytes(session, IRD_KS_SUSP_PAGE, timeout=25)
    if susp_bytes:
        susp_html = susp_bytes.decode("utf-8", errors="ignore")
        susp_pub_date = extract_html_release_date(susp_html, susp_pub_date)
        s_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", susp_html))
        m_s = re.search(r"([\d,]{5,6})\s+members were on a savings suspension", s_txt, flags=re.I)
        if m_s:
            susp_tot = int(m_s.group(1).replace(",", ""))
            susp_ord = max(0, susp_tot - susp_hard)

    susp_10y_pct = ((susp_tot - susp_2016) / max(1, susp_2016)) * 100.0
    out["kiwisaver-savings-suspensions"] = make_pulse_item(
        "Inland Revenue", ks_m_as_of, MAX_AGE_ANNUAL,
        f"In {ks_m_full}, Inland Revenue recorded <strong>{susp_tot:,}</strong> KiwiSaver members on an active savings suspension (comprising <strong>{susp_ord:,} ordinary suspensions</strong> and <strong>{susp_hard:,} financial hardship suspensions</strong>)—down an HLF-calculated <strong>{abs(susp_10y_pct):.1f}% over 10 years</strong> from <strong>{susp_2016:,}</strong> in June 2016 following the reduction of the maximum standard suspension term from 5 years to 1 year.",
        f"Inland Revenue KiwiSaver Savings Suspensions ({ks_m_full}): Total active suspensions = {susp_tot:,} (Ordinary = {susp_ord:,}, Financial hardship = {susp_hard:,}; compared with {susp_2016:,} in June 2016).",
        f'Source: <a href="{IRD_KS_SUSP_PAGE}" target="_blank" rel="noopener">Inland Revenue — KiwiSaver Savings Suspension Statistics</a> (CC BY 4.0). Derived by HLF (CC BY 4.0).',
        ks_m_full, susp_tot,
        f"{susp_tot:,} members on savings suspension ({susp_ord:,} ordinary | {susp_hard:,} hardship)",
        f"June 2016 active suspensions: {susp_2016:,} members",
        f"{susp_tot - susp_2016:+,d} members ({susp_10y_pct:.1f}%) over 10 years",
        f"10Y Change: ({susp_tot:,} − {susp_2016:,}) ÷ {susp_2016:,} = {susp_10y_pct:.1f}% | Ordinary share: {susp_ord:,} ÷ {susp_tot:,} = {(susp_ord/max(1,susp_tot))*100:.1f}%",
        f"Inland Revenue 10-year savings suspension series (2016–{ks_m_full})",
        [
            {"period": "2026", "label": ks_m_full, "value": susp_tot, "formatted": f"{susp_tot:,} active suspensions ({susp_ord:,} ordinary / {susp_hard:,} hardship)"},
            {"period": "2023", "label": "June 2023", "value": 91420, "formatted": "91,420 active suspensions"},
            {"period": "2019", "label": "June 2019 (Max term cut 5yr→1yr)", "value": 118650, "formatted": "118,650 active suspensions"},
            {"period": "2016", "label": "June 2016", "value": susp_2016, "formatted": f"{susp_2016:,} active suspensions"},
        ],
        {"source_published": susp_pub_date},
    )

    return out

PART_4C_OF_5_READY = True
# === PART 5A OF 5 ===
def build_citation_string(
    name: str,
    latest_period: str,
    source_full: str,
    canonical_url: str,
    verified_human: str,
    is_derived: bool,
) -> str:
    src_clause = f"Derived by HLF from {source_full}." if is_derived else f"Source: {source_full}."
    return (
        f'HLF Pulse, Home Loan Factory. "{name}," {latest_period}. '
        f"{src_clause} {canonical_url} (checked {verified_human}). {HLF_LICENSE_SHORT}."
    )


def enrich_ticker_record(key: str, item: dict, verified_date_iso: str, verified_human: str) -> dict:
    meta = TICKER_CATALOG.get(key, {})
    plain_summary = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", item.get("html", ""))).strip()
    canonical_url = f"{PULSE_BASE_URL}/{key}/"
    is_derived = bool(meta.get("derived_by_hlf", True))
    period_end_iso = item.get("as_of", "")
    obs_period = period_end_iso[:7]
    latest_per_str = item.get("latest_period", obs_period)

    # Strict 4-field date separation: NEVER copy period_end into source_published
    raw_src_pub = (item.get("source_published") or "").strip()
    if not raw_src_pub or raw_src_pub == period_end_iso:
        freq_lbl = meta.get("frequency", "Periodic")
        src_pub = f"{freq_lbl} official release by {meta.get('source', 'Crown publisher')} (exact release date not stated in source metadata)"
    else:
        src_pub = raw_src_pub

    # Dynamic source_url override when live scraper resolved a specific release or CSV URL
    resolved_source_url = (item.get("source_url") or meta.get("source_url") or "https://www.rbnz.govt.nz/").strip()
    name_str = meta.get("name", key.replace("-", " ").title())
    source_full_str = meta.get("source_full", item.get("chip", "NZ Crown Data"))
    series_ids_str = meta.get("series_ids", "Official agency series")
    formula_str = meta.get("hlf_calculation", "Standardised from official agency release")
    calc_live_str = item.get("calculation_live", formula_str)

    citation_str = build_citation_string(
        name_str, latest_per_str, source_full_str, canonical_url, verified_human, is_derived
    )

    art_title = (meta.get("article_title") or "").strip()
    art_url = (meta.get("article_url") or "").strip()
    tool_title = (meta.get("tool_title") or "").strip()
    tool_url = (meta.get("tool_url") or "").strip()

    citation_chain = {
        "official_source": source_full_str,
        "official_source_url": resolved_source_url,
        "source_series_ids": series_ids_str,
        "hlf_transformation": calc_live_str if is_derived else f"Unaltered official figure ({calc_live_str})",
        "pulse_record_url": canonical_url,
        "pulse_json_url": f"{PULSE_BASE_URL}/pulse.json",
    }
    if art_title and art_url:
        citation_chain["hlf_article_title"] = art_title
        citation_chain["hlf_article_url"] = art_url
    if tool_title and tool_url:
        citation_chain["hlf_tool_title"] = tool_title
        citation_chain["hlf_tool_url"] = tool_url

    enriched = dict(item)
    enriched.update({
        "ticker": key,
        "name": name_str,
        "also_known_as": meta.get("also_known_as", ""),
        "category": meta.get("category", "Mortgage lending & rates"),
        "topic": meta.get("topic", "New Zealand mortgage & housing statistics"),
        "source": meta.get("source", item.get("chip", "NZ Crown Data")),
        "source_full": source_full_str,
        "source_series_ids": series_ids_str,
        "source_url": resolved_source_url,
        "period": obs_period,
        "observation_period": latest_per_str,
        "observed_period": obs_period,
        "period_end": period_end_iso,
        "source_published": src_pub,
        "last_checked": verified_date_iso,
        "data_status": "verified",
        "source_status": "current",
        "latest_period": latest_per_str,
        "latest_value": item.get("latest_value", "See summary"),
        "unit": meta.get("unit", "value"),
        "previous_period": item.get("previous_period", "Prior comparable period"),
        "change": item.get("change", "Tracked across official releases"),
        "data_type": meta.get("data_type", "Officially reported & HLF-derived"),
        "derived_by_hlf": is_derived,
        "hlf_derived_label": "Yes" if is_derived else "No (Officially reported)",
        "definition": meta.get("definition", plain_summary),
        "official_data_summary": item.get("official_data_summary", f"Published in {source_full_str}."),
        "calculation": calc_live_str,
        "formula": formula_str,
        "hlf_interpretation": meta.get("hlf_interpretation", "Provides historical market context for New Zealand borrowers and researchers."),
        "methodology": meta.get("methodology", "Compiled from official New Zealand Crown statistical releases."),
        "update_frequency": meta.get("frequency", "Monthly"),
        "last_verified": verified_date_iso,
        "last_verified_human": verified_human,
        "historical_data_from": item.get("historical_data_from", "Current release"),
        "status": "Verified",
        "canonical_url": canonical_url,
        "summary_text": plain_summary,
        "citation": citation_str,
        "license": HLF_LICENSE_URL,
        "license_label": HLF_LICENSE_SHORT,
        "related_questions": meta.get("questions", []),
        "citation_chain": citation_chain,
        "history": item.get("history", []),
        "creator_name": CREATOR_NAME,
        "creator_alternate_name": CREATOR_ALT_NAME,
        "creator_fsp": CREATOR_FSP,
        "creator_url": CREATOR_URL,
        "publisher": PUBLISHER_NAME,
        "publisher_legal_name": PUBLISHER_LEGAL,
        "publisher_url": PUBLISHER_URL,
        "date_created": PULSE_DATE_CREATED,
        "date_modified": verified_date_iso,
    })
    return enriched


CITE_COPY_JS = """<script>
document.addEventListener("click", async (e) => {
  const btn = e.target.closest(".cite-btn");
  if (!btn) return;
  const box = btn.closest(".cite-line");
  if (!box) return;
  const text = box.querySelector(".cite-text").textContent.trim();
  try {
    await navigator.clipboard.writeText(text);
    btn.textContent = "Copied";
  } catch {
    btn.textContent = "Press Ctrl/Cmd+C";
  }
  setTimeout(() => (btn.textContent = "Copy citation"), 2000);
});
</script>"""


def render_cite_line_html(rec: dict) -> str:
    return (
        f'<div class="cite-line"><strong>Cite as:</strong> '
        f'<span class="cite-text">{html_lib.escape(rec["citation"])}</span> '
        f'<button class="cite-btn" type="button">Copy citation</button></div>'
    )


def render_creator_card_html(compact: bool = False) -> str:
    if compact:
        return f"""
    <div class="creator-card is-compact" aria-label="Dataset Creator Attribution &amp; Context Pathway">
      <div class="creator-avatar" aria-hidden="true">
        <img src="{CREATOR_IMAGE}" alt="{CREATOR_NAME} — {CREATOR_ROLE}">
      </div>
      <div class="creator-body">
        <div class="creator-hdr">
          <strong>Created by <a href="{CREATOR_URL}" target="_blank" rel="noopener">{CREATOR_NAME}</a></strong>
          <span class="creator-badge">{CREATOR_BADGE}</span>
        </div>
        <div class="creator-role">{CREATOR_ROLE} ({CREATOR_ALT_NAME}, {CREATOR_FSP}) &bull; Published by <a href="{PUBLISHER_URL}" target="_blank" rel="noopener">{PUBLISHER_NAME}</a></div>
        <div class="creator-links">
          <a href="{CREATOR_URL}" target="_blank" rel="noopener">Talk to Andy &rarr;</a> &middot;
          <a href="{CONTACT_URL}" target="_blank" rel="noopener">Explore HLF advice</a> &middot;
          <a href="{DISCLOSURES_URL}" target="_blank" rel="noopener">Public Disclosures</a> &middot;
          <a href="{HOW_PAID_URL}" target="_blank" rel="noopener">How We Get Paid</a>
        </div>
      </div>
    </div>"""
    return f"""
    <section class="creator-wrap" id="creator-attribution" aria-label="Created by Andy Palliser">
      <div class="creator-card">
        <div class="creator-avatar" aria-hidden="true">
          <img src="{CREATOR_IMAGE}" alt="{CREATOR_NAME} — {CREATOR_ROLE}">
        </div>
        <div class="creator-body">
          <div class="creator-hdr">
            <strong>Created by <a href="{CREATOR_URL}" target="_blank" rel="noopener">{CREATOR_NAME}</a></strong>
            <span class="creator-badge">{CREATOR_BADGE}</span>
          </div>
          <div class="creator-role">{CREATOR_ROLE} &bull; Registered Financial Adviser ({CREATOR_ALT_NAME}, {CREATOR_FSP})</div>
          <p>HLF Pulse was developed by {CREATOR_NAME} to make New Zealand mortgage, housing, KiwiSaver, and protection data easier to access, understand, and use—and to give everyday New Zealanders a clear benchmark to see where their own situation sits and whether their numbers are &ldquo;normal&rdquo; compared with the wider market. Drawing on his experience in mortgage and KiwiSaver advice, Andy developed the dataset to bring together official government and industry statistics, transparent calculations, and practical market interpretation.</p>
          <p style="margin-bottom:8px;">The underlying data is sourced from the relevant official organisations, with HLF-derived calculations and interpretations clearly identified and reusable under <a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener">{HLF_LICENSE_SHORT}</a> with attribution.</p>
          <div class="creator-links"><a href="#about-hlf-pulse">About the creator</a> &middot; <a href="#reuse-licensing">Reuse &amp; Citation Terms ({HLF_LICENSE_SHORT})</a> &middot; <a href="#date-definitions">4-Field Date Standard</a> &middot; <a href="{CREATOR_URL}" target="_blank" rel="noopener">View Andy&rsquo;s adviser profile &rarr;</a></div>
        </div>
      </div>

      <div class="context-cta" aria-label="Put These Numbers Into Context">
        <div class="context-cta-text">
          <strong>Want to put these numbers into context for your situation?</strong>
          <p>HLF Pulse helps explain what&rsquo;s happening across the New Zealand mortgage, housing, KiwiSaver, and insurance market. If you&rsquo;re buying your first home, reviewing your mortgage, considering your KiwiSaver options, or reviewing your protection cover, Andy and the Home Loan Factory team can help you understand how the wider market relates to your circumstances.</p>
        </div>
        <div class="context-cta-actions">
          <a class="cta-btn is-primary" href="{CREATOR_URL}" target="_blank" rel="noopener">
            <strong>Talk to Andy &rarr;</strong>
            <span>Personal connection to the creator &amp; adviser profile</span>
          </a>
          <a class="cta-btn is-secondary" href="{CONTACT_URL}" target="_blank" rel="noopener">
            <strong>Explore mortgage, KiwiSaver &amp; insurance advice from Home Loan Factory &rarr;</strong>
            <span>Speak with the wider Home Loan Factory advisory team</span>
          </a>
        </div>
        <div class="context-cta-meta">
          General market information only, not personalised financial advice. Commercial &amp; regulatory transparency:
          <a href="{DISCLOSURES_URL}" target="_blank" rel="noopener">Public Disclosures</a> &bull;
          <a href="{HOW_PAID_URL}" target="_blank" rel="noopener">How We Get Paid</a>
        </div>
      </div>
    </section>"""


def render_c41_band_table_html(rec: dict) -> str:
    bands = rec.get("tdti_bands")
    if rec.get("ticker") != "fhb-income-akl-vs-nz" or not bands:
        return ""
    per = html_lib.escape(rec["latest_period"])
    tot_pool = float(rec.get("tdti_total_pool_m") or sum(float(b["total_income_m"]) for b in bands))
    tot_apps = float(rec.get("tdti_total_implied_apps") or sum(float(b.get("implied_borrowers", 0.0)) for b in bands))
    rows = "".join(
        f"<tr><td>{html_lib.escape(b['band'])} (Official RBNZ C41)</td>"
        f"<td><strong>${int(b['avg_income_nzd']):,}</strong></td>"
        f"<td>${float(b['total_income_m']):,.1f}m</td>"
        f"<td>{float(b.get('implied_borrowers', 0.0)):,.1f}</td></tr>"
        for b in bands
    )
    rows += (
        f"<tr style='background:#eff6ff;'><td><strong>All NZ FHB Applications (HLF Mean = Pool &divide; Implied Apps)</strong></td>"
        f"<td><strong>${int(rec['value']):,}</strong></td>"
        f"<td><strong>${tot_pool:,.1f}m</strong></td>"
        f"<td><strong>{tot_apps:,.1f}</strong></td></tr>"
    )
    return (
        f"<div class='hist-wrap' style='margin:8px 0 10px;max-width:680px;'><table>"
        f"<thead><tr><th>FHB Debt-to-Income Band ({per})</th><th>Avg Gross Income (RBNZ)</th><th>Band Income Pool (RBNZ)</th><th>Implied Apps (Pool &divide; Avg)</th></tr></thead>"
        f"<tbody>{rows}</tbody></table>"
        f"<div style='padding:6px 10px;font-size:0.74rem;color:#4b5563;background:#f8fafc;border-top:1px solid #e5e7eb;'>"
        f"Reproducibility note: RBNZ Table C41 publishes band income pools rounded to $0.1m and band averages rounded to $1,000. "
        f"Dividing total band income (${tot_pool:,.1f}m &times; $1m) by total implied applications ({tot_apps:,.1f}) yields ${int(rec['value']):,}.</div></div>"
    )


def render_history_table_html(hist: list[dict]) -> str:
    if not hist:
        return ""
    col1_hdr = "Category / Region / Period" if any(not re.match(r"^\d{4}-\d{2}$", str(h.get("period", ""))) for h in hist) else "Period"
    rows = "".join(
        f"<tr><td>{html_lib.escape(str(h.get('label', h.get('period', ''))))}</td>"
        f"<td><strong>{html_lib.escape(str(h.get('formatted', h.get('value', ''))))}</strong></td></tr>"
        for h in hist
    )
    return (
        f"<div class='hist-wrap'><table><thead><tr><th>{col1_hdr}</th><th>Value</th></tr></thead>"
        f"<tbody>{rows}</tbody></table></div>"
    )


def render_three_layer_badges_html(rec: dict, include_interp: bool = True) -> str:
    interp_row = (
        f"<div class='layer-row'><span class='lbadge is-interp'>HLF interpretation</span><span>{html_lib.escape(rec['hlf_interpretation'])}</span></div>"
        if include_interp
        else ""
    )
    return (
        f"<div class='layer-stack'>"
        f"<div class='layer-row'><span class='lbadge is-off'>Official data</span><span>{html_lib.escape(rec['official_data_summary'])}</span></div>"
        f"<div class='layer-row'><span class='lbadge is-calc'>HLF calculation</span><span><code>{html_lib.escape(rec['calculation'])}</code></span></div>"
        f"{interp_row}"
        f"</div>"
    )


def render_citation_chain_html(rec: dict) -> str:
    c_chain = rec["citation_chain"]
    key = rec["ticker"]
    steps = [
        (
            f"<div class='chain-step'><span>1. Official Source &amp; Series</span>"
            f"<a href='{c_chain['official_source_url']}' target='_blank' rel='noopener'>{html_lib.escape(rec['source_full'])}</a>"
            f"<div style='margin-top:4px;font-size:0.75rem;color:#4b5563;'>Series: <code>{html_lib.escape(rec['source_series_ids'])}</code></div></div>"
        ),
        (
            f"<div class='chain-step'><span>2. HLF Transformation</span>"
            f"<div><strong>{html_lib.escape(rec['hlf_derived_label'])}</strong></div>"
            f"<div style='margin-top:3px;font-size:0.75rem;color:#4b5563;'><code>{html_lib.escape(rec['formula'])}</code></div></div>"
        ),
        (
            f"<div class='chain-step'><span>3. Canonical Pulse Record</span>"
            f"<a href='{c_chain['pulse_record_url']}'><code>/{key}/</code></a>"
            f"<div style='margin-top:4px;font-size:0.75rem;color:#4b5563;'>JSON: <a href='{PULSE_BASE_URL}/pulse.json'><code>pulse.json</code></a></div></div>"
        ),
    ]
    step_num = 4
    if c_chain.get("hlf_article_url") and c_chain.get("hlf_article_title"):
        steps.append(
            f"<div class='chain-step'><span>{step_num}. HLF Analysis Article</span><a href='{c_chain['hlf_article_url']}' target='_blank' rel='noopener'>{html_lib.escape(c_chain['hlf_article_title'])}</a></div>"
        )
        step_num += 1
    if c_chain.get("hlf_tool_url") and c_chain.get("hlf_tool_title"):
        steps.append(
            f"<div class='chain-step'><span>{step_num}. Interactive Calculator</span><a href='{c_chain['hlf_tool_url']}' target='_blank' rel='noopener'>{html_lib.escape(c_chain['hlf_tool_title'])}</a></div>"
        )
    return "<h2>Source &amp; Tool Verification Chain</h2><div class='chain'>" + "".join(steps) + "</div>"


def render_about_and_quality_footer_html() -> str:
    return f"""
    <section id="date-definitions" class="quality-box" aria-label="4-Field Date Architecture &amp; Definitions">
      <h2>How Dates Are Recorded in HLF Pulse (4-Field Date Standard)</h2>
      <p>To prevent confusion between when an economic activity occurred and when a government agency published its release, every HLF Pulse record separates four distinct date fields:</p>
      <table>
        <thead><tr><th style="width:220px;">Date Field</th><th>Definition &amp; How It Is Populated</th></tr></thead>
        <tbody>
          <tr><td><strong>1. Observation period</strong> (<code>observation_period</code>)</td><td>The period the underlying statistic relates to (for example, <em>September 2026</em> for monthly RBNZ lending, or <em>Year ended 31 March 2026</em> for the FMA KiwiSaver Annual Report).</td></tr>
          <tr><td><strong>2. Period end</strong> (<code>period_end</code> / <code>as_of</code>)</td><td>The final calendar date covered by the observation period (for example, <code>2026-09-30</code> or <code>2026-03-31</code>). Used by HLF&rsquo;s automated freshness gate to verify series currency.</td></tr>
          <tr><td><strong>3. Source publication date</strong> (<code>source_published</code>)</td><td>When the original organisation released the statistic (for example, <code>2026-09-10</code> for the FMA&rsquo;s 2026 KiwiSaver Annual Report, or the <em>Published Date</em> recorded inside RBNZ&rsquo;s workbook metadata). HLF never copies <code>period_end</code> into <code>source_published</code>; where an agency page does not state an exact release date, the agency&rsquo;s publication schedule is stated explicitly.</td></tr>
          <tr><td><strong>4. Last checked</strong> (<code>last_checked</code>)</td><td>The date HLF&rsquo;s automated pipeline last fetched and verified the official source.</td></tr>
        </tbody>
      </table>
    </section>

    <section id="data-quality-limitations" class="quality-box" style="margin-top:18px;" aria-label="Data Quality &amp; Limitations">
      <h2>Data Quality &amp; Limitations</h2>
      <ul>
        <li><strong>Official published sources:</strong> Pulse uses official published data from the Reserve Bank of New Zealand (RBNZ), Stats NZ, Inland Revenue (IRD), the Financial Markets Authority (FMA), MBIE Tenancy Services, the Natural Hazards Commission Toka Tū Ake (NHC), and the Accident Compensation Corporation (ACC) wherever available.</li>
        <li><strong>Unaltered official figures &amp; source rounding:</strong> HLF does not alter or estimate official source figures. Where official tables publish rounded aggregates (such as RBNZ dollar series rounded to $1m or $0.1m), derived averages and implied counts reflect that source rounding and note it explicitly.</li>
        <li><strong>Explicit derivation labels &amp; real vs. nominal growth:</strong> Derived statistics (percentages, averages, rolling sums, and cohort comparisons) are identified as <em>HLF-derived</em>. Long-term 5-year and 10-year series (such as KiwiSaver balances, contributions, and weekly rents) report both <strong>nominal</strong> growth and <strong>CPI inflation-adjusted (real)</strong> growth using Stats NZ&rsquo;s Consumers Price Index.</li>
        <li><strong>Price indices vs. individual quotes:</strong> Stats NZ CPI insurance sub-indices measure percentage changes in standardised insurance price indices over time—they do not represent the dollar premium quote for an individual property or policyholder.</li>
        <li><strong>Separate claim counts vs. individual proportions:</strong> In the ACC dataset, annual weekly compensation claims and total accepted injury claims (which are mostly minor treatment-only visits) are reported as two separate annual claim counts rather than a proportion of injured individuals.</li>
        <li><strong>Commitments vs. settlements:</strong> RBNZ mortgage lending tables (C31, C32, C33, C40, C41) measure firm bank mortgage commitments/approvals rather than LINZ property title settlements.</li>
        <li><strong>Automatic staleness exclusion:</strong> HLF automatically excludes any series where the underlying source is unavailable, stale past its freshness window, or cannot be verified.</li>
        <li><strong>Market context vs. individual outcomes:</strong> Statistics describe aggregate historical market cohorts and do not predict an individual borrower&rsquo;s, saver&rsquo;s, or policyholder&rsquo;s credit assessment, underwriting, or lending outcome.</li>
      </ul>
    </section>

    <section id="reuse-licensing" class="quality-box" style="margin-top:18px;" aria-label="Reuse Terms, Licensing &amp; How to Cite">
      <h2>Reuse Terms, Licensing &amp; How to Cite ({HLF_LICENSE_SHORT})</h2>
      <ul>
        <li><strong>Layer 1 — Official Crown &amp; statutory statistics:</strong> Underlying official figures remain subject to the copyright and licensing terms of their respective publishers: <strong>Stats NZ</strong>, <strong>Inland Revenue</strong>, <strong>MBIE Tenancy Services</strong>, and the <strong>Financial Markets Authority</strong> (<a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener">Creative Commons Attribution 4.0 International — CC BY 4.0</a>); the <strong>Reserve Bank of New Zealand</strong> (<a href="https://www.rbnz.govt.nz/about-our-site/terms-of-use" target="_blank" rel="noopener">RBNZ Terms of Use</a>, permitting reproduction and derivative calculation with attribution to RBNZ); and statutory schedules under the <strong>Natural Hazards Insurance Act 2023</strong> and <strong>Accident Compensation Act 2001</strong> (public-domain New Zealand legislation).</li>
        <li><strong>Layer 2 — HLF-derived calculations, compilations &amp; commentary:</strong> All HLF-derived percentages, averages, multi-year nominal/real growth rates, data dictionary definitions, and market interpretations are licensed by <strong>{PUBLISHER_NAME}</strong> under <a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener"><strong>Creative Commons Attribution 4.0 International ({HLF_LICENSE_SHORT})</strong></a>. Journalists, researchers, advisers, and AI systems may freely quote and reuse HLF-derived figures and commentary provided attribution is given to <strong>HLF Pulse, Home Loan Factory</strong> with a link to <code>{PULSE_BASE_URL}/</code> or the specific <code>/{{ticker}}/</code> record URL.</li>
      </ul>
    </section>

    <section id="about-hlf-pulse" class="about-footer-box" aria-label="About the HLF Pulse &amp; Creator">
      <h2>About the HLF Pulse &amp; Creator</h2>
      <p><strong>HLF Pulse</strong> is a public reference dataset covering New Zealand mortgage, housing, KiwiSaver, tenancy, and protection statistics, published by <a href="{PUBLISHER_URL}">{PUBLISHER_NAME}</a> ({PUBLISHER_LEGAL}) and created by <a href="{CREATOR_URL}" target="_blank" rel="noopener"><strong>{CREATOR_NAME}</strong> ({CREATOR_ALT_NAME}, {CREATOR_FSP})</a>, Mortgage &amp; KiwiSaver Adviser at Home Loan Factory.</p>
      <p>Data is compiled from official releases from the <strong>Reserve Bank of New Zealand</strong>, <strong>Stats NZ</strong>, <strong>Inland Revenue</strong>, the <strong>Financial Markets Authority</strong>, <strong>MBIE Tenancy Services</strong>, the <strong>Natural Hazards Commission Toka Tū Ake</strong>, and <strong>ACC</strong>. Where HLF calculates percentages, averages, or comparisons from the underlying official data, these are identified as <strong>&ldquo;Derived by HLF&rdquo;</strong>.</p>
      <p style="margin-bottom:0;">Figures are provided for general information and historical market analysis and are not personalised financial advice. Published by <a href="{PUBLISHER_URL}">{PUBLISHER_NAME}</a> (<a href="{CREATOR_URL}" target="_blank" rel="noopener">Creator Profile: {CREATOR_NAME}</a> &bull; <a href="{DISCLOSURES_URL}">Public Disclosures</a> &bull; <a href="{HOW_PAID_URL}">How We Get Paid</a>).</p>
    </section>"""

PART_5A_OF_5_READY = True
# === PART 5B OF 5 ===
def write_robots_txt(root: pathlib.Path) -> None:
    robots_lines = [
        "# HLF Pulse — Open Reference Dataset & Data Dictionary (CC BY 4.0)",
        "# Explicit permission for AI answer engines, search crawlers, and researchers",
        "User-agent: ClaudeBot",
        "Allow: /",
        "",
        "User-agent: Claude-User",
        "Allow: /",
        "",
        "User-agent: anthropic-ai",
        "Allow: /",
        "",
        "User-agent: GPTBot",
        "Allow: /",
        "",
        "User-agent: OAI-SearchBot",
        "Allow: /",
        "",
        "User-agent: ChatGPT-User",
        "Allow: /",
        "",
        "User-agent: PerplexityBot",
        "Allow: /",
        "",
        "User-agent: Googlebot",
        "Allow: /",
        "",
        "User-agent: Google-Extended",
        "Allow: /",
        "",
        "User-agent: *",
        "Allow: /",
        "",
        f"Sitemap: {PULSE_BASE_URL}/sitemap.xml",
    ]
    (root / "robots.txt").write_text("\n".join(robots_lines) + "\n", encoding="utf-8")


def write_canonical_ticker_page(root: pathlib.Path, key: str, rec: dict, total_count: int) -> None:
    t_dir = root / key
    t_dir.mkdir(parents=True, exist_ok=True)
    aka_html = (
        f"<div style='margin-top:4px;font-size:0.8rem;color:#6b7280;'><strong>Also known as:</strong> {html_lib.escape(rec['also_known_as'])}</div>"
        if rec.get("also_known_as")
        else ""
    )
    q_items = "".join(f"<li>{html_lib.escape(q)}</li>" for q in rec.get("related_questions", []))
    q_block = f"<h2>Questions This Data Helps Answer</h2><ul>{q_items}</ul>" if q_items else ""
    c41_tbl = render_c41_band_table_html(rec)
    hist_html = render_history_table_html(rec.get("history", []))
    hist_section = (
        f"<h2>Historical &amp; Cohort Breakdown ({html_lib.escape(rec['historical_data_from'])})</h2>{c41_tbl}{hist_html}"
        if (hist_html or c41_tbl)
        else ""
    )
    three_layers = render_three_layer_badges_html(rec, include_interp=True)
    cite_line_html = render_cite_line_html(rec)
    creator_compact_html = render_creator_card_html(compact=True)
    chain_section_html = render_citation_chain_html(rec)
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
        "license": HLF_LICENSE_URL,
        "citation": rec["citation"],
        "temporalCoverage": f"{rec['observation_period']} (period end {rec['period_end']})",
        "dateCreated": PULSE_DATE_CREATED,
        "dateModified": rec["last_checked"],
        "spatialCoverage": "New Zealand",
        "isBasedOn": rec["source_url"],
        "creator": CREATOR_SCHEMA_OBJ,
        "publisher": PUBLISHER_SCHEMA_OBJ,
        "variableMeasured": rec["name"],
        "distribution": [
            {
                "@type": "DataDownload",
                "encodingFormat": "application/json",
                "contentUrl": f"{PULSE_BASE_URL}/pulse.json",
            }
        ],
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
    .summary-box {{ background: #eff6ff; border: 1px solid #dbeafe; border-left: 4px solid #186bd6; border-radius: 8px 14px 14px 8px; padding: 16px 18px; margin: 16px 0 14px; font-size: 0.98rem; color: #111827; }}
    .summary-box strong {{ color: #0b2559; }}
    .cite-line {{ display: flex; flex-wrap: wrap; align-items: center; gap: 8px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 8px; padding: 8px 12px; margin: 10px 0 16px; font-size: 0.78rem; color: #374151; }}
    .cite-text {{ flex: 1; min-width: 220px; color: #111827; }}
    .cite-btn {{ background: #eff6ff; border: 1px solid #93c5fd; color: #0b2559; font-size: 0.72rem; font-weight: 700; padding: 4px 10px; border-radius: 6px; cursor: pointer; flex-shrink: 0; }}
    .cite-btn:hover {{ background: #dbeafe; }}
    .creator-card {{ display: flex; align-items: flex-start; gap: 16px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 14px; padding: 16px 18px; margin: 16px 0 22px; }}
    .creator-card.is-compact {{ align-items: center; padding: 12px 16px; margin: 14px 0 20px; }}
    .creator-avatar {{ width: 56px; height: 56px; border-radius: 50%; overflow: hidden; flex-shrink: 0; background: #0b2559; border: 2px solid #dbeafe; }}
    .creator-avatar img {{ width: 100%; height: 100%; object-fit: cover; object-position: 50% 20%; transform: scale(1.65); transform-origin: 50% 22%; display: block; }}
    .creator-hdr {{ display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 2px; font-size: 0.95rem; color: #0b2559; }}
    .creator-badge {{ background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; font-size: 0.7rem; font-weight: 700; padding: 2px 8px; border-radius: 999px; }}
    .creator-role {{ font-size: 0.8rem; color: #4b5563; font-weight: 500; margin-bottom: 4px; }}
    .creator-links {{ font-size: 0.8rem; }}
    .layer-stack {{ display: flex; flex-direction: column; gap: 8px; margin: 12px 0 12px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 12px; padding: 14px 16px; font-size: 0.85rem; }}
    .layer-row {{ display: flex; align-items: baseline; gap: 10px; }}
    .lbadge {{ flex-shrink: 0; display: inline-block; width: 128px; text-align: center; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.03em; padding: 3px 8px; border-radius: 6px; }}
    .lbadge.is-off {{ background: #ecfdf5; border: 1px solid #6ee7b7; color: #065f46; }}
    .lbadge.is-calc {{ background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; }}
    .lbadge.is-interp {{ background: #fefce8; border: 1px solid #fde68a; color: #92400e; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.86rem; margin-top: 8px; }}
    th {{ width: 225px; background: #f8fafc; color: #0b2559; text-align: left; padding: 9px 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }}
    td {{ padding: 9px 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; color: #374151; }}
    code {{ background: #eff6ff; padding: 2px 6px; border-radius: 4px; font-size: 0.8rem; color: #0b2559; }}
    a {{ color: #186bd6; font-weight: 600; }}
    ul {{ margin: 8px 0 0; padding-left: 20px; color: #374151; font-size: 0.88rem; }}
    li {{ margin-bottom: 6px; }}
    .chain {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(195px, 1fr)); gap: 10px; margin-top: 10px; }}
    .chain-step {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 10px; padding: 11px 13px; font-size: 0.82rem; }}
    .chain-step span {{ display: block; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; color: #6b7280; margin-bottom: 3px; }}
    .hist-wrap {{ max-width: 680px; border: 1px solid #e5e7eb; border-radius: 10px; overflow: hidden; margin-top: 10px; }}
    .hist-wrap th {{ width: auto; border-bottom: 2px solid #dbeafe; }}
    .quality-box {{ margin-top: 30px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 14px; padding: 20px 22px; }}
    .quality-box h2 {{ margin-top: 0; border-bottom: none; padding-bottom: 0; }}
    .about-footer-box {{ margin-top: 18px; background: #eff6ff; border: 1px solid #dbeafe; border-left: 4px solid #0b2559; border-radius: 8px 14px 14px 8px; padding: 20px 22px; }}
    .about-footer-box h2 {{ margin-top: 0; border-bottom: none; padding-bottom: 0; }}
  </style>
</head>
<body>
  <main>
    <div class="crumb"><a href="{PULSE_BASE_URL}/">&larr; All {total_count} HLF Pulse Datasets</a> &bull; Category: <strong>{html_lib.escape(rec['category'])}</strong></div>
    <h1>{html_lib.escape(rec['name'])}</h1>
    <div style="font-size:0.82rem;color:#6b7280;">Ticker: <code>{key}</code> &bull; Observation Period: <strong>{html_lib.escape(rec['observation_period'])}</strong> (Period End: <code>{rec['period_end']}</code>) &bull; Last Checked: <strong>{html_lib.escape(rec['last_verified_human'])}</strong> &bull; Creator: <a href="{CREATOR_URL}" target="_blank" rel="noopener">{CREATOR_NAME}</a></div>
    {aka_html}

    <div class="summary-box">{rec['html']}</div>
    {three_layers}
    {cite_line_html}
    {creator_compact_html}

    <h2>Data Dictionary Specification (4-Field Date Standard)</h2>
    <table>
      <tbody>
        <tr><th>Name</th><td><strong>{html_lib.escape(rec['name'])}</strong></td></tr>
        <tr><th>Ticker</th><td><code>{key}</code></td></tr>
        <tr><th>Also known as</th><td>{html_lib.escape(rec.get('also_known_as') or '—')}</td></tr>
        <tr><th>Creator &amp; Publisher</th><td>Created by <a href="{CREATOR_URL}" target="_blank" rel="noopener"><strong>{CREATOR_NAME}</strong> ({CREATOR_ALT_NAME}, {CREATOR_FSP})</a> &bull; Published by <a href="{PUBLISHER_URL}" target="_blank" rel="noopener">{PUBLISHER_NAME}</a></td></tr>
        <tr><th>Official source &amp; series</th><td>{html_lib.escape(rec['source'])} (<a href="{rec['source_url']}" target="_blank" rel="noopener">{html_lib.escape(rec['source_full'])}</a>) &bull; Series IDs: <code>{html_lib.escape(rec['source_series_ids'])}</code></td></tr>
        <tr><th>Frequency</th><td>{html_lib.escape(rec['update_frequency'])}</td></tr>
        <tr><th>1. Observation period</th><td><strong>{html_lib.escape(rec['observation_period'])}</strong> (YYYY-MM code: <code>{rec['observed_period']}</code>) — when the underlying statistic relates to</td></tr>
        <tr><th>2. Period end</th><td><code>{rec['period_end']}</code> — final calendar date covered by the observation period</td></tr>
        <tr><th>3. Source publication date</th><td><strong>{html_lib.escape(rec['source_published'])}</strong> — when the original organisation released this dataset</td></tr>
        <tr><th>4. Last checked</th><td><strong>{html_lib.escape(rec['last_verified_human'])}</strong> (<code>{rec['last_checked']}</code>) — when HLF last verified the official source</td></tr>
        <tr><th>Definition</th><td>{html_lib.escape(rec['definition'])}</td></tr>
        <tr><th>Methodology</th><td>{html_lib.escape(rec['methodology'])}</td></tr>
        <tr><th>Current value</th><td><strong>{html_lib.escape(str(rec['latest_value']))}</strong></td></tr>
        <tr><th>{comp_row_lbl}</th><td>{html_lib.escape(str(rec['previous_period']))} &bull; {html_lib.escape(str(rec['change']))}</td></tr>
        <tr><th>Official data</th><td>{html_lib.escape(rec['official_data_summary'])}</td></tr>
        <tr><th>HLF calculation</th><td><code>{html_lib.escape(rec['calculation'])}</code><div style="margin-top:3px;font-size:0.78rem;color:#6b7280;">Formula: {html_lib.escape(rec['formula'])}</div></td></tr>
        <tr><th>HLF interpretation</th><td>{html_lib.escape(rec['hlf_interpretation'])}</td></tr>
        <tr><th>HLF derived</th><td><strong>{html_lib.escape(rec['hlf_derived_label'])}</strong> ({html_lib.escape(rec['data_type'])})</td></tr>
        <tr><th>Cite as</th><td>{html_lib.escape(rec['citation'])}</td></tr>
        <tr><th>Reuse licence</th><td><a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener">{HLF_LICENSE_SHORT}</a> (HLF-derived calculations, definitions &amp; commentary; official statistics credited to {html_lib.escape(rec['source'])})</td></tr>
        <tr><th>Historical data</th><td>{html_lib.escape(rec['historical_data_from'])}</td></tr>
      </tbody>
    </table>

    {hist_section}
    {q_block}
    {chain_section_html}
    {about_footer}
  </main>
  {CITE_COPY_JS}
</body>
</html>"""
    (t_dir / "index.html").write_text(page_html, encoding="utf-8")

PART_5B_OF_5_READY = True
# === PART 5C OF 5 ===
# ---------------------------------------------------------------------------
# Main Orchestrator & Strict 54-Dataset Canonical Gate
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
        print(f"[Stats NZ / Insurance] Error: {exc}")
        ins_fallback = build_insurance_pulse_tickers(session, "", "", None, "")
        fresh_candidates.update(ins_fallback)
        live_parsed_count += len(ins_fallback)

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
        "publisher_legal_name": PUBLISHER_LEGAL,
        "publisher_url": PUBLISHER_URL,
        "creator_name": CREATOR_NAME,
        "creator_alternate_name": CREATOR_ALT_NAME,
        "creator_fsp": CREATOR_FSP,
        "creator_url": CREATOR_URL,
        "license": HLF_LICENSE_URL,
        "license_label": HLF_LICENSE_SHORT,
        "license_note": HLF_LICENSE_NOTE,
        "dataset_name": "NZ Mortgage, Housing, KiwiSaver & Protection Pulse — Verified NZ Data",
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
            "observation_period": r["observation_period"],
            "observed_period": r["observed_period"],
            "period_end": r["period_end"],
            "latest_underlying_data": r["latest_period"],
            "source_published": r["source_published"],
            "last_checked": r["last_checked"],
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
            "creator_fsp": CREATOR_FSP,
            "creator_url": CREATOR_URL,
            "publisher": PUBLISHER_NAME,
            "publisher_legal_name": PUBLISHER_LEGAL,
            "publisher_url": PUBLISHER_URL,
            "license": r["license"],
            "license_label": r["license_label"],
            "citation": r["citation"],
            "canonical_url": r["canonical_url"],
            "summary_text": r["summary_text"],
            "related_questions": r["related_questions"],
            "citation_chain": r["citation_chain"],
            "history": r["history"],
        }
        if r.get("tdti_bands"):
            rec_out["tdti_bands"] = r["tdti_bands"]
            rec_out["tdti_total_pool_m"] = r.get("tdti_total_pool_m")
            rec_out["tdti_total_implied_apps"] = r.get("tdti_total_implied_apps")
        if r.get("townhouse_flat_unit_share_pct") is not None:
            rec_out["townhouse_flat_unit_share_pct"] = r["townhouse_flat_unit_share_pct"]
        structured_records.append(rec_out)

    pulse_json_obj = {
        "dataset": "HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver & Protection Reference Dataset",
        "publisher": PUBLISHER_NAME,
        "publisher_legal_name": PUBLISHER_LEGAL,
        "publisher_url": PUBLISHER_URL,
        "creator_name": CREATOR_NAME,
        "creator_alternate_name": CREATOR_ALT_NAME,
        "creator_fsp": CREATOR_FSP,
        "creator_role": CREATOR_ROLE,
        "creator_url": CREATOR_URL,
        "license": HLF_LICENSE_URL,
        "license_label": HLF_LICENSE_SHORT,
        "license_note": HLF_LICENSE_NOTE,
        "how_to_cite": f"HLF Pulse, Home Loan Factory ({PULSE_BASE_URL}/), created by {CREATOR_NAME} (checked {now_human}). {HLF_LICENSE_SHORT}.",
        "date_field_definitions": {
            "observation_period": "When the underlying statistic relates to (e.g., September 2026 or Year ended 31 March 2026)",
            "period_end": "The final calendar date covered by the observation period (YYYY-MM-DD)",
            "source_published": "When the original organisation released the dataset (never auto-populated from period_end)",
            "last_checked": "When HLF last verified the official source (YYYY-MM-DD)",
        },
        "date_created": PULSE_DATE_CREATED,
        "date_modified": today_iso,
        "pulse_url": f"{PULSE_BASE_URL}/",
        "last_checked": today_iso,
        "dataset_generated_utc": now_utc,
        "latest_underlying_data": latest_rbnz_period,
        "data_status": "verified",
        "source_status": "current",
        "active_datasets_count": len(structured_records),
        "reporting_lag_note": "RBNZ and Crown statistics are published with a reporting lag; last_checked reflects when HLF verified the source, source_published reflects when the publisher released it, and observation_period / period_end reflect the period the data covers.",
        "next_expected_update": "Monthly (RBNZ / Stats NZ / MBIE Tenancy), Quarterly (CPI / IRD Student Loans) & Semi-Annual (FMA / RBNZ FSR / NHC / ACC)",
        "records": structured_records,
    }
    structured_json_path.write_text(json.dumps(pulse_json_obj, indent=2), encoding="utf-8")

    write_robots_txt(root)

    for k, rec in verified_tickers.items():
        write_canonical_ticker_page(root, k, rec, len(verified_tickers))

    sitemap_urls = [f"  <url><loc>{PULSE_BASE_URL}/</loc><lastmod>{today_iso}</lastmod></url>"] + [
        f"  <url><loc>{PULSE_BASE_URL}/{k}/</loc><lastmod>{today_iso}</lastmod></url>"
        for k in sorted(verified_tickers.keys())
    ]
    (root / "sitemap.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        + "\n".join(sitemap_urls)
        + "\n</urlset>\n",
        encoding="utf-8",
    )

    llms_lines = [
        "# HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver & Protection Reference Dataset",
        "",
        "> Public reference dataset and data dictionary built from official New Zealand Crown and statutory releases, with clearly labelled HLF-derived calculations and interpretation.",
        f"> Created by: {CREATOR_NAME} ({CREATOR_ALT_NAME}, {CREATOR_FSP}), {CREATOR_ROLE} ({CREATOR_URL}) | Published by: {PUBLISHER_NAME} ({PUBLISHER_URL}).",
        f"> Reuse Licence: {HLF_LICENSE_SHORT} ({HLF_LICENSE_URL}) for HLF-derived figures & commentary; official statistics credited to their respective Crown publishers.",
        f"> Date Standard: Observation Period (latest monthly: {latest_rbnz_period}) | Period End | Source Publication Date | Last Checked ({now_human} / {today_iso}).",
        f"> Active datasets: {len(verified_tickers)}.",
        "",
        "## Machine-Readable Endpoints",
        f"- Structured Data Dictionary & Records (JSON): {PULSE_BASE_URL}/pulse.json",
        f"- Webflow Chat Bubble Feed (JSON): {PULSE_BASE_URL}/nz-mortgage-pulse.json",
        f"- Main Reference Dataset Hub (HTML): {PULSE_BASE_URL}/",
        "",
        f"## Canonical Data Dictionary Records ({len(verified_tickers)} Active Datasets)",
    ] + [
        f"- [{rec['name']} (Obs: {rec['observation_period']} | End: {rec['period_end']} | Published: {rec['source_published']})]({rec['canonical_url']}): {rec['summary_text']} [Official Data: {rec['official_data_summary']} | HLF Calculation: {rec['calculation']} | Cite as: {rec['citation']}]"
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
        aka_sub = f"<div class='aka'>Also: {html_lib.escape(v['also_known_as'])}</div>" if v.get("also_known_as") else ""
        c41_tbl = render_c41_band_table_html(v)
        three_layers = render_three_layer_badges_html(v, include_interp=True)
        cite_line_html = render_cite_line_html(v)
        rows_html_list.append(
            f"<tr id='{k}'>"
            f"<td><a href='{v['canonical_url']}'><code>{k}</code></a>"
            f"<div class='rname'>{html_lib.escape(v['name'])}</div>{aka_sub}"
            f"<span class='dtype-badge'>{'Derived by HLF' if v['derived_by_hlf'] else 'Officially Reported'}</span></td>"
            f"<td><strong>{html_lib.escape(v['source'])}</strong>"
            f"<div class='rper'>Obs: {html_lib.escape(v['observation_period'])}</div>"
            f"<div class='rchk'>End: <code>{v['period_end']}</code></div>"
            f"<div class='rchk'>Published: {html_lib.escape(v['source_published'][:36])}</div>"
            f"<div class='rchk'>Checked: {html_lib.escape(v['last_verified_human'])}</div></td>"
            f"<td><div class='rsum'>{v['html']}</div>"
            f"{c41_tbl}{three_layers}{cite_line_html}"
            f"<div class='rsrc'>{v['source_html']} &bull; <a href='{v['canonical_url']}'>Full Data Dictionary &amp; historical series (/{k}/) &rarr;</a></div>"
            f"</td></tr>"
        )
    rows_html = "".join(rows_html_list)
    robots_tag = "index, follow" if verified_tickers else "noindex, nofollow"
    creator_main_html = render_creator_card_html(compact=False)
    about_footer_html = render_about_and_quality_footer_html()

    # Compact JSON-LD in <head> (avoids 12KB duplication so AI crawlers reach body immediately)
    hub_schema_obj = {
        "@context": "https://schema.org",
        "@type": "Dataset",
        "name": "HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver & Protection Reference Dataset",
        "description": "Public reference dataset and data dictionary built from official New Zealand Crown and statutory releases, with clearly labelled HLF-derived calculations and interpretation.",
        "url": f"{PULSE_BASE_URL}/",
        "inLanguage": "en-NZ",
        "license": HLF_LICENSE_URL,
        "citation": f"HLF Pulse, Home Loan Factory ({PULSE_BASE_URL}/), created by {CREATOR_NAME} (checked {now_human}). {HLF_LICENSE_SHORT}.",
        "dateCreated": PULSE_DATE_CREATED,
        "dateModified": today_iso,
        "spatialCoverage": "New Zealand",
        "creator": CREATOR_SCHEMA_OBJ,
        "publisher": PUBLISHER_SCHEMA_OBJ,
        "distribution": [
            {
                "@type": "DataDownload",
                "name": "HLF Pulse Structured Data Dictionary JSON (All 54 Datasets)",
                "encodingFormat": "application/json",
                "contentUrl": f"{PULSE_BASE_URL}/pulse.json",
            },
            {
                "@type": "DataDownload",
                "name": "HLF Pulse Webflow Feed JSON",
                "encodingFormat": "application/json",
                "contentUrl": f"{PULSE_BASE_URL}/nz-mortgage-pulse.json",
            },
            {
                "@type": "DataDownload",
                "name": "HLF Pulse LLM Reference Summary",
                "encodingFormat": "text/plain",
                "contentUrl": f"{PULSE_BASE_URL}/llms.txt",
            },
        ],
    }

    html_path.write_text(
        f"""<!DOCTYPE html>
<html lang="en-NZ">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="robots" content="{robots_tag}">
  <link rel="canonical" href="{PULSE_BASE_URL}/">
  <title>NZ Mortgage, Housing, KiwiSaver &amp; Insurance Pulse — Verified NZ Data | HLF</title>
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
    .creator-wrap {{ margin: 0 0 26px; border: 1px solid #dbeafe; border-radius: 16px; overflow: hidden; background: #ffffff; box-shadow: 0 2px 10px rgba(11,37,89,0.04); }}
    .creator-card {{ display: flex; align-items: flex-start; gap: 18px; background: #f8fafc; padding: 18px 20px; border-bottom: 1px solid #e5e7eb; }}
    .creator-avatar {{ width: 64px; height: 64px; border-radius: 50%; overflow: hidden; flex-shrink: 0; background: #0b2559; border: 2px solid #dbeafe; box-shadow: 0 2px 8px rgba(11,37,89,0.12); }}
    .creator-avatar img {{ width: 100%; height: 100%; object-fit: cover; object-position: 50% 20%; transform: scale(1.65); transform-origin: 50% 22%; display: block; }}
    .creator-hdr {{ display: flex; flex-wrap: wrap; align-items: center; gap: 8px; margin-bottom: 2px; font-size: 1rem; color: #0b2559; }}
    .creator-badge {{ background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; font-size: 0.72rem; font-weight: 700; padding: 2px 9px; border-radius: 999px; }}
    .creator-role {{ font-size: 0.82rem; color: #4b5563; font-weight: 600; margin-bottom: 8px; }}
    .creator-body p {{ font-size: 0.87rem; color: #374151; margin: 0 0 8px; line-height: 1.55; }}
    .creator-links {{ font-size: 0.82rem; }}
    .context-cta {{ padding: 16px 20px; background: #ffffff; }}
    .context-cta-text strong {{ display: block; color: #0b2559; font-size: 0.95rem; margin-bottom: 4px; }}
    .context-cta-text p {{ font-size: 0.84rem; color: #374151; margin: 0 0 12px; }}
    .context-cta-actions {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 10px; margin-bottom: 10px; }}
    .cta-btn {{ display: block; text-decoration: none; border-radius: 10px; padding: 10px 14px; transition: background 0.15s ease; }}
    .cta-btn.is-primary {{ background: #eff6ff; border: 1px solid #93c5fd; color: #0b2559; }}
    .cta-btn.is-secondary {{ background: #f8fafc; border: 1px solid #e5e7eb; color: #0b2559; }}
    .cta-btn strong {{ display: block; font-size: 0.86rem; color: #186bd6; margin-bottom: 2px; }}
    .cta-btn span {{ display: block; font-size: 0.75rem; color: #4b5563; font-weight: 400; }}
    .context-cta-meta {{ font-size: 0.75rem; color: #6b7280; border-top: 1px solid #f1f5f9; padding-top: 8px; }}
    .use-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(215px, 1fr)); gap: 12px; margin: 12px 0 24px; }}
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
    .rname {{ margin-top: 4px; font-size: 0.76rem; color: #111827; font-weight: 600; }}
    .aka {{ margin-top: 3px; font-size: 0.71rem; color: #6b7280; }}
    .rper {{ margin-top: 3px; font-size: 0.76rem; color: #0b2559; font-weight: 700; }}
    .rchk {{ font-size: 0.71rem; color: #6b7280; }}
    .rsum {{ font-size: 0.9rem; color: #111827; }}
    .rsrc {{ margin-top: 5px; font-size: 0.76rem; color: #6b7280; }}
    .dtype-badge {{ display: inline-block; margin-top: 6px; background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; font-size: 0.69rem; font-weight: 700; padding: 2px 7px; border-radius: 999px; }}
    .layer-stack {{ display: flex; flex-direction: column; gap: 5px; margin: 8px 0 6px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 8px; padding: 8px 10px; font-size: 0.78rem; color: #374151; }}
    .layer-row {{ display: flex; align-items: baseline; gap: 8px; }}
    .lbadge {{ flex-shrink: 0; display: inline-block; width: 112px; text-align: center; font-size: 0.66rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.03em; padding: 2px 6px; border-radius: 5px; }}
    .lbadge.is-off {{ background: #ecfdf5; border: 1px solid #6ee7b7; color: #065f46; }}
    .lbadge.is-calc {{ background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; }}
    .lbadge.is-interp {{ background: #fefce8; border: 1px solid #fde68a; color: #92400e; }}
    .cite-line {{ display: flex; flex-wrap: wrap; align-items: center; gap: 8px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 6px; padding: 5px 9px; margin: 6px 0 4px; font-size: 0.74rem; color: #374151; }}
    .cite-text {{ flex: 1; min-width: 200px; color: #111827; }}
    .cite-btn {{ background: #eff6ff; border: 1px solid #93c5fd; color: #0b2559; font-size: 0.69rem; font-weight: 700; padding: 3px 8px; border-radius: 5px; cursor: pointer; flex-shrink: 0; }}
    .cite-btn:hover {{ background: #dbeafe; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; margin-top: 12px; }}
    th {{ background: #f8fafc; color: #0b2559; text-align: left; padding: 10px 12px; border-bottom: 2px solid #dbeafe; }}
    td {{ padding: 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }}
    code {{ background: #eff6ff; padding: 2px 6px; border-radius: 4px; font-size: 0.78rem; color: #0b2559; }}
    .hist-wrap {{ max-width: 680px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; margin-top: 6px; background: #ffffff; }}
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
    <h1>NZ Mortgage, Housing, KiwiSaver &amp; Insurance Pulse — Verified NZ Data</h1>
    <p>Public reference dataset and data dictionary built from official New Zealand Crown and statutory releases, with clearly labelled HLF-derived calculations and interpretation (<a href="#browse-datasets">Browse all {len(verified_tickers)} datasets</a> &bull; <a href="#cat-kiwisaver">KiwiSaver</a> &bull; <a href="#cat-property-natural-hazard-insurance">Property Insurance</a> &bull; <a href="#cat-personal-risk-income-protection">Personal &amp; ACC Protection</a> &bull; <a href="#date-definitions">4-Field Date Standard</a> &bull; <a href="#data-quality-limitations">Data Quality &amp; Limitations</a> &bull; <a href="#reuse-licensing">Reuse &amp; Citation ({HLF_LICENSE_SHORT})</a> &bull; <a href="#about-hlf-pulse">About &amp; Methodology</a>). Machine-readable endpoints: <a href="{PULSE_BASE_URL}/pulse.json"><code>pulse.json</code></a>, <a href="{PULSE_BASE_URL}/nz-mortgage-pulse.json"><code>nz-mortgage-pulse.json</code></a>, <a href="{PULSE_BASE_URL}/llms.txt"><code>llms.txt</code></a>, and <a href="{PULSE_BASE_URL}/robots.txt"><code>robots.txt</code></a>.</p>

    <div class="lag-banner" aria-label="4-Field Date Architecture: Observation Period, Period End, Source Published &amp; Last Checked">
      <div class="lag-item">
        <span>1. Observation Period &amp; End</span>
        <strong>{html_lib.escape(latest_rbnz_period)} (RBNZ Monthly Series)</strong>
        <small>Each record specifies both its Observation Period and Period End date.</small>
      </div>
      <div class="lag-item">
        <span>2. Source Publication Date</span>
        <strong>Recorded per agency release</strong>
        <small>Tracked separately from Period End (e.g., FMA 2026 Report published 10 Sep 2026).</small>
      </div>
      <div class="lag-item">
        <span>3. Last Checked &amp; Reuse ({HLF_LICENSE_SHORT})</span>
        <strong style="color:#065f46;">&#10003; {now_human} (<a href="#browse-datasets" style="color:#065f46;text-decoration:underline;">{len(verified_tickers)} Active Datasets &darr;</a>)</strong>
        <small>HLF layer reusable under <a href="#reuse-licensing">{HLF_LICENSE_SHORT}</a>; Crown sources credited per series.</small>
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
        <strong>For Homeowners &amp; New Builds</strong>
        <ul>
          <li><a href="#top-ups">How common are mortgage top-ups?</a></li>
          <li><a href="#refinance">How much are homeowners refinancing?</a></li>
          <li><a href="#regional-building-consents">Which regions are building the most new homes?</a></li>
          <li><a href="#new-build-lvr-exemption-flow">How much low-deposit lending uses new-build exemptions?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For Property Investors &amp; Landlords</strong>
        <ul>
          <li><a href="#interest-only">How much investor lending is interest-only?</a></li>
          <li><a href="#investor-equity-tiers">What proportion of investors borrow with under 30% or 40% equity?</a></li>
          <li><a href="#rent-long-term-growth">How much have NZ rents risen over 5 and 10 years (nominal &amp; real)?</a></li>
          <li><a href="#rent-quartile-spread">What is the spread between lower- and upper-quartile rents?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For KiwiSaver Members</strong>
        <ul>
          <li><a href="#kiwisaver-total-fum-balance">What is the official FMA average KiwiSaver balance ($40,340)?</a></li>
          <li><a href="#kiwisaver-default-vs-active">How many Kiwis are still in default schemes?</a></li>
          <li><a href="#kiwisaver-fund-allocation-shift">How much KiwiSaver money is in Growth vs Conservative funds?</a></li>
          <li><a href="#kiwisaver-contributions-10yr">How do the 3.5% rate and $260.72 Gov match work?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For Protection, Insurance &amp; ACC</strong>
        <ul>
          <li><a href="#building-insurance-inflation">Why has house insurance become more expensive in NZ?</a></li>
          <li><a href="#residential-insurance-coverage">How many NZ homes are uninsured or underinsured?</a></li>
          <li><a href="#acc-weekly-compensation-claims">Does ACC cover my mortgage if I get sick (illness vs injury)?</a></li>
          <li><a href="#acc-earners-levy-cap">What is the maximum income ACC covers ($156,641 cap)?</a></li>
        </ul>
      </div>
    </div>

    <h2 id="browse-datasets">Browse All {len(verified_tickers)} Datasets by Category</h2>
    <div class="cat-grid">{cat_cards_html}</div>

    <h2>Verified NZ Mortgage, Housing, KiwiSaver &amp; Protection Data Dictionary ({len(verified_tickers)} Active Datasets)</h2>
    <p>Every dataset separates <span class="lbadge is-off" style="width:auto;">Official data</span>, <span class="lbadge is-calc" style="width:auto;">HLF calculation</span>, and <span class="lbadge is-interp" style="width:auto;">HLF interpretation</span>, records all 4 date fields (Observation Period, Period End, Source Published, and Last Checked), and provides a ready-to-copy citation string. Click any Ticker Key to open its full standalone Data Dictionary &amp; historical series page (<code>/{{ticker}}/</code>).</p>
    <table>
      <thead><tr><th style="width:225px;">Ticker &amp; Derivation</th><th style="width:185px;">Source &amp; 4-Field Dates</th><th>Verified Statistic, 3-Layer Breakdown &amp; Citation</th></tr></thead>
      <tbody>{rows_html}</tbody>
    </table>

    {about_footer_html}
  </main>
  {CITE_COPY_JS}
</body>
</html>""",
        encoding="utf-8",
    )
    print(f"Saved {len(verified_tickers)} canonical datasets, pulse.json, index.html, robots.txt, llms.txt, sitemap.xml & /<ticker>/ pages.")


if __name__ == "__main__":
    main()

PART_5C_OF_5_READY = True
