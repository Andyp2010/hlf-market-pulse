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
# Official NZ Crown Data Sources (RBNZ, Stats NZ, MBIE Tenancy, Inland Revenue)
# ---------------------------------------------------------------------------
PULSE_BASE_URL = "https://pulse-data.homeloanfactory.co.nz"
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

MAX_AGE_MONTHLY = 85
MAX_AGE_TENANCY = 120
MAX_AGE_QUARTERLY = 160
MAX_AGE_ANNUAL = 420

# ---------------------------------------------------------------------------
# Reference-Layer Catalog: Definitions, Methodology, Questions & Citation Chain
# ---------------------------------------------------------------------------
TICKER_CATALOG = {
    "fhb-low-deposit": {
        "name": "First-home buyers with less than 20% deposit",
        "topic": "First-home lending",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.BC1, LVRN.MMB2.AC1, LVRN.MMB1.BC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived share & HLF-calculated comparison (from officially reported RBNZ loan counts & dollar values)",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of firm residential mortgage commitments to New Zealand first-home buyers where the Loan-to-Value Ratio (LVR) exceeds 80% (meaning the buyer has less than a 20% deposit), measured by borrower/loan count alongside officially reported high-LVR dollar lending.",
        "methodology": "Compiled from monthly registered bank returns under RBNZ Table C31. First-home buyers are defined by lending banks as borrowers purchasing a residential property to occupy as a first home.",
        "hlf_calculation": "(High-LVR FHB Borrowers [LVRN.MMB2.BC1] / Total FHB Borrowers [LVRN.MMB2.AC1]) * 100, compared against the same month 12 months earlier.",
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
        "topic": "First-home lending",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.AC1, LVRN.MMB1.AA, LVRN.MMB2.AC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported loan count & HLF-derived market share / average loan size",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The proportion of total new residential mortgage lending dollars in New Zealand committed to first-home buyers, alongside the officially reported monthly count of approved first-home loans and HLF-derived average loan size.",
        "methodology": "Sourced from RBNZ Table C31 monthly registered bank commitments across first-home buyers, other owner-occupiers, property investors, and business purposes.",
        "hlf_calculation": "Share = (FHB Lending NZDm [LVRN.MMB1.AC1] / Total Lending NZDm [LVRN.MMB1.AA]) * 100. Avg Loan = (LVRN.MMB1.AC1 * 1,000,000) / LVRN.MMB2.AC1.",
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
        "name": "First-home buyer average mortgage size by deposit tier (<20% vs 20%+ deposit)",
        "topic": "First-home lending",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.BC1, LVRN.MMB2.BC1, LVRN.MMB1.CC1, LVRN.MMB2.CC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived cohort averages & HLF-calculated comparison",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "Comparison of the average approved mortgage size for New Zealand first-home buyers with under a 20% deposit (>80% LVR) versus first-home buyers with a 20%+ deposit (<=80% LVR).",
        "methodology": "Derived from RBNZ Table C31 by dividing total monthly dollar lending in each LVR tier by the number of approved first-home borrowers in that same LVR tier.",
        "hlf_calculation": "Low-Deposit Avg = (LVRN.MMB1.BC1 * 1,000,000) / LVRN.MMB2.BC1. 20%+ Deposit Avg = (LVRN.MMB1.CC1 * 1,000,000) / LVRN.MMB2.CC1.",
        "questions": [
            "Do low-deposit first-home buyers borrow more or less than 20% deposit buyers?",
            "What is the average low-deposit first-home loan size in NZ?",
            "Should I wait to save a 20% deposit or buy sooner with 10%?",
        ],
        "article_title": "Buying Sooner With a Low Deposit vs Waiting for 20%",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "fhb-bank-lvr-priority": {
        "name": "First-home buyer share of all bank low-deposit (>80% LVR) lending",
        "topic": "First-home lending",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.BC1, LVRN.MMB1.BA, LVRN.MMB2.BC1, LVRN.MMB2.BA",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived share (from officially reported RBNZ high-LVR totals)",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The share of all >80% LVR (under 20% deposit/equity) residential mortgage lending across New Zealand that banks allocate to first-home buyers.",
        "methodology": "Compares high-LVR first-home commitments against high-LVR commitments across all borrower types combined in RBNZ Table C31.",
        "hlf_calculation": "Dollar Share = (LVRN.MMB1.BC1 / LVRN.MMB1.BA) * 100. Borrower Share = (LVRN.MMB2.BC1 / LVRN.MMB2.BA) * 100.",
        "questions": [
            "Do New Zealand banks prioritise first-home buyers for low-deposit loans?",
            "What share of bank low-equity (>80% LVR) quotas goes to first-home buyers?",
        ],
        "article_title": "How Bank LVR Speed Limits Work in New Zealand",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-annual-milestone": {
        "name": "Annual rolling 12-month first-home buyer mortgage approvals",
        "topic": "First-home lending",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.AC1, LVRN.MMB1.AC1, LVRN.MMB2.BC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived 12-month rolling sum & HLF-calculated comparison",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Monthly",
        "definition": "The cumulative number and dollar value of first-home buyer mortgages approved across New Zealand over the trailing 12 months, including the number approved with under a 20% deposit.",
        "methodology": "Rolling 12-month sum of monthly RBNZ Table C31 firm commitments to first-home buyers.",
        "hlf_calculation": "Sum of trailing 12 months of LVRN.MMB2.AC1 (borrowers), LVRN.MMB1.AC1 (NZDm), and LVRN.MMB2.BC1 (>80% LVR borrowers).",
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
        "name": "First-home buyers approved with a 20% or greater deposit (<=80% LVR)",
        "topic": "First-home lending",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.CC1, LVRN.MMB2.AC1, LVRN.MMB1.CC1",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported count/value & HLF-derived share/average",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The number, percentage share, and average loan size of New Zealand first-home buyers approved with a deposit of 20% or more (LVR <= 80%).",
        "methodology": "Sourced from RBNZ Table C31 series for <=80% LVR first-home buyer lending and borrower counts.",
        "hlf_calculation": "Share = (LVRN.MMB2.CC1 / LVRN.MMB2.AC1) * 100. Avg Loan = (LVRN.MMB1.CC1 * 1,000,000) / LVRN.MMB2.CC1.",
        "questions": [
            "What percentage of NZ first-home buyers have a 20% deposit?",
            "What is the average mortgage size for a first-home buyer with a 20% deposit?",
        ],
        "article_title": "How a 20% Deposit Unlocks Special Mortgage Rates",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Compare NZ Mortgage Rates",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
    },
    "fhb-vs-investors": {
        "name": "First-home buyers vs property investors monthly lending comparison",
        "topic": "First-home & investor lending",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.AC1, LVRN.MMB1.AN, LVRN.MMB1.AA",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported totals & HLF-calculated comparison",
        "derived_by_hlf": True,
        "unit": "NZDm",
        "frequency": "Monthly",
        "definition": "Head-to-head comparison of monthly new residential mortgage lending committed to first-home buyers versus residential property investors across New Zealand.",
        "methodology": "Compares RBNZ Table C31 total monthly lending to first-home buyers (LVRN.MMB1.AC1) against total monthly lending to property investors (LVRN.MMB1.AN).",
        "hlf_calculation": "Difference = LVRN.MMB1.AC1 - LVRN.MMB1.AN; Market shares = each series divided by LVRN.MMB1.AA.",
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
        "topic": "Property investment lending",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.AN, LVRN.MMB1.AA",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported dollar lending & HLF-derived market share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage share and dollar value of total new residential mortgage lending in New Zealand committed to residential property investors.",
        "methodology": "Sourced from RBNZ Table C31 monthly registered bank commitments to residential property investors.",
        "hlf_calculation": "Investor Share = (Investor Lending NZDm [LVRN.MMB1.AN] / All Borrower Lending NZDm [LVRN.MMB1.AA]) * 100.",
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
        "name": "Existing owner-occupiers approved with 20% or more equity (<=80% LVR)",
        "topic": "Homeowner equity & refinancing",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.CC2, LVRN.MMB2.AC2, LVRN.MMB1.CC2",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported counts/values & HLF-derived share/average",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of existing owner-occupier mortgage approvals (movers, refinancers, and top-up borrowers, excluding first-home buyers) with at least 20% equity (LVR <= 80%), plus their average mortgage size.",
        "methodology": "Derived from RBNZ Table C31 'Other owner-occupiers' series at or below 80% LVR.",
        "hlf_calculation": "Share = (LVRN.MMB2.CC2 / LVRN.MMB2.AC2) * 100. Avg Loan = (LVRN.MMB1.CC2 * 1,000,000) / LVRN.MMB2.CC2.",
        "questions": [
            "How much equity do existing New Zealand homeowners typically have?",
            "What is the average mortgage size for an existing homeowner in NZ?",
            "What percentage of refinancing or moving homeowners have at least 20% equity?",
        ],
        "article_title": "How to Calculate & Use Your Usable Home Equity",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "investor-equity-tiers": {
        "name": "Property investor equity tier breakdown (40%+, 30%–40%, and <30% equity)",
        "topic": "Property investment & equity",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.TN, LVRN.MMB2.FN, LVRN.MMB2.DN, LVRN.MMB2.AN",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "HLF-derived equity tier distribution (from RBNZ C31 investor LVR memo items)",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "Breakdown of approved New Zealand residential property investors across three equity tiers: 40%+ equity (<=60% LVR), 30%–40% equity (60% < LVR <= 70%), and under 30% equity (>70% LVR).",
        "methodology": "Derived from RBNZ Table C31 investor memo series tracking LVR thresholds at 60% and 70%.",
        "hlf_calculation": "40%+ equity = LVRN.MMB2.TN / LVRN.MMB2.AN; 30%–40% equity = (LVRN.MMB2.FN - LVRN.MMB2.DN) / LVRN.MMB2.AN; <30% equity = LVRN.MMB2.DN / LVRN.MMB2.AN.",
        "questions": [
            "How much deposit or equity do you need to buy an investment property in NZ?",
            "What percentage of NZ property investors buy with a 30% deposit vs 40%?",
        ],
        "article_title": "Using Home Equity to Buy an Investment Property",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Usable Equity & Investment Cashflow Calculators",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "investor-low-deposit": {
        "name": "Property investor loans approved with under 20% deposit (>80% LVR) & 20%–30% deposit",
        "topic": "Property investment & new builds",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.BN, LVRN.MMB2.BN, LVRN.MMB1.DN, LVRN.MMB2.DN",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported >80% LVR counts/values & HLF-derived 70%–80% LVR new-build tier",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Monthly",
        "definition": "The number and dollar value of residential property investor mortgages approved with strictly under a 20% deposit (>80% LVR), alongside the 20%–30% deposit tier (70% < LVR <= 80%) where standard 20%-deposit new-build investor loans sit.",
        "methodology": "In RBNZ Table C31, strictly under-20% deposit investor loans are recorded in >80% LVR (LVRN.MMB2.BN / LVRN.MMB1.BN), while exact 20.0% deposit (80.0% LVR) new-build investor loans sit inside 70% < LVR <= 80% (LVRN.MMB2.DN minus LVRN.MMB2.BN).",
        "hlf_calculation": ">80% LVR = LVRN.MMB2.BN & LVRN.MMB1.BN. 70%–80% LVR band = (LVRN.MMB2.DN - LVRN.MMB2.BN) & (LVRN.MMB1.DN - LVRN.MMB1.BN).",
        "questions": [
            "Can property investors buy with a 20% deposit or less in New Zealand?",
            "How many investor loans are approved with under a 20% or 30% deposit in NZ?",
            "How do new-build LVR exemptions work for property investors?",
        ],
        "article_title": "Buying a New-Build Investment Property With a 20% Deposit",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Turnkey vs. Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/turnkey-vs-construction-loan-calculator-nz",
    },
    "investor-new-build-lvr": {
        "name": "Property investors approved with under a 30% deposit (>70% LVR new-build & speed-limit tier)",
        "topic": "Property investment & new builds",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB1.DN, LVRN.MMB2.DN, LVRN.MMB1.BN, LVRN.MMB2.BN",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported >70% LVR totals & HLF-derived share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The total number, dollar value, and percentage share of New Zealand property investor mortgages approved with under a 30% deposit (>70% LVR), capturing 20%-deposit new-build exemptions and bank speed-limit allocations.",
        "methodology": "Sourced from RBNZ Table C31 Memo Item H2/D2 (Higher than 70% LVR investor borrowers and lending).",
        "hlf_calculation": "Share = (LVRN.MMB2.DN / LVRN.MMB2.AN) * 100; split into 70%–80% LVR and >80% LVR tiers.",
        "questions": [
            "What percentage of NZ property investors buy with less than a 30% deposit?",
            "How many investors use the 20% deposit new-build LVR exemption?",
        ],
        "article_title": "New-Build LVR Exemptions for Property Investors",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "mover-low-equity": {
        "name": "Existing owner-occupiers approved with under 20% equity (>80% LVR)",
        "topic": "Next-home buyers & homeowner equity",
        "source": "Reserve Bank of New Zealand Table C31",
        "series_ids": "LVRN.MMB2.BC2, LVRN.MMB1.BC2, LVRN.MMB2.AC2",
        "source_url": RBNZ_SOURCES["c31"]["pages"][0],
        "data_type": "Officially reported counts/values & HLF-derived share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The number, dollar value, and percentage share of existing owner-occupiers (non-first-home buyers moving or borrowing) approved with under 20% equity (>80% LVR).",
        "methodology": "Sourced from RBNZ Table C31 series B3/F3 (Higher than 80% LVR other owner-occupiers).",
        "hlf_calculation": "Share = (LVRN.MMB2.BC2 / LVRN.MMB2.AC2) * 100.",
        "questions": [
            "Can existing homeowners move or buy a second home with less than 20% equity?",
            "How many existing NZ homeowners get approved with under a 20% deposit?",
        ],
        "article_title": "Upgrading Your Home With Less Than 20% Equity",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Net Proceeds & Next Home Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/net-proceeds-and-next-home-calculator",
    },
    "refinance": {
        "name": "Monthly mortgage refinancing ( borrowers switching bank providers)",
        "topic": "Refinancing & bank switching",
        "source": "Reserve Bank of New Zealand Table C33",
        "series_ids": "LVRN.MMB1.AZD, LVRN.MMB2.AZD",
        "source_url": RBNZ_SOURCES["c33"]["pages"][0],
        "data_type": "Officially reported refinance volume/count & HLF-derived share/average",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The monthly number of borrowers and total dollar value of residential mortgages switching from one mortgage provider to another in New Zealand.",
        "methodology": "Compiled from RBNZ Table C33 (New residential mortgage lending by purpose — Change in loan provider).",
        "hlf_calculation": "Share = (Refinance NZDm [LVRN.MMB1.AZD] / Total Purpose Lending NZDm) * 100. Avg Refinance Loan = (LVRN.MMB1.AZD * 1,000,000) / LVRN.MMB2.AZD.",
        "questions": [
            "How many New Zealand homeowners switch banks or refinance each month?",
            "What is the average mortgage size for borrowers refinancing in NZ?",
            "Is switching mortgage providers common in New Zealand?",
        ],
        "article_title": "When Does Refinancing Your Mortgage Make Sense?",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Refinance & Cashback Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/refinance-calculator",
    },
    "top-ups": {
        "name": "Monthly mortgage top-ups on existing properties (borrower share & average top-up)",
        "topic": "Homeowner equity & top-ups",
        "source": "Reserve Bank of New Zealand Table C33",
        "series_ids": "LVRN.MMB2.AZB, LVRN.MMB1.AZB",
        "source_url": RBNZ_SOURCES["c33"]["pages"][0],
        "data_type": "Officially reported top-up count & HLF-derived share/average",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The monthly number of New Zealand homeowners approved for an equity top-up on their existing mortgage, their share of total approvals, and average top-up amount.",
        "methodology": "Compiled from RBNZ Table C33 (New residential mortgage lending by purpose — Top-ups).",
        "hlf_calculation": "Borrower Share = (Top-up Borrowers [LVRN.MMB2.AZB] / Total Purpose Borrowers) * 100. Avg Top-up = (LVRN.MMB1.AZB * 1,000,000) / LVRN.MMB2.AZB.",
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
        "name": "Total monthly dollar equity unlocked via mortgage top-ups",
        "topic": "Homeowner equity & top-ups",
        "source": "Reserve Bank of New Zealand Table C33",
        "series_ids": "LVRN.MMB1.AZB, LVRN.MMB2.AZB",
        "source_url": RBNZ_SOURCES["c33"]["pages"][0],
        "data_type": "Officially reported dollar volume & HLF-derived average",
        "derived_by_hlf": True,
        "unit": "NZDm",
        "frequency": "Monthly",
        "definition": "Total monthly dollar lending unlocked by existing New Zealand homeowners through mortgage top-ups without selling their property.",
        "methodology": "Sourced from RBNZ Table C33 series LVRN.MMB1.AZB (Top-up lending value) and LVRN.MMB2.AZB (Top-up borrower count).",
        "hlf_calculation": "Directly reported LVRN.MMB1.AZB dollar total and LVRN.MMB2.AZB count, with average top-up = (LVRN.MMB1.AZB * 1,000,000) / LVRN.MMB2.AZB.",
        "questions": [
            "How much equity do New Zealand homeowners unlock via mortgage top-ups each month?",
            "Can I top up my mortgage to fund renovations or an investment deposit?",
        ],
        "article_title": "Using a Mortgage Top-Up to Renovate or Invest",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "interest-only": {
        "name": "Interest-only share of new investor vs owner-occupier mortgages",
        "topic": "Loan structure & interest-only",
        "source": "Reserve Bank of New Zealand Table C32",
        "series_ids": "LVRN.MMB1.AN, LVRN.MMB1.AC",
        "source_url": RBNZ_SOURCES["c32"]["pages"][0],
        "data_type": "HLF-derived share (from RBNZ C32 payment-type series)",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of new residential investor lending versus new owner-occupier lending set up on interest-only terms (including revolving credit facilities).",
        "methodology": "Derived from RBNZ Table C32 (New and existing residential mortgage lending by payment type).",
        "hlf_calculation": "Investor IO % = (New Investor Interest-Only NZDm / Total New Investor NZDm) * 100; Owner-Occupier IO % = (New Owner-Occupier Interest-Only NZDm / Total New Owner-Occupier NZDm) * 100.",
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
        "name": "First-home buyer high-DTI (>6x) lending share & bank speed-limit headroom",
        "topic": "Debt-to-Income (DTI) & borrowing power",
        "source": "Reserve Bank of New Zealand Table C40",
        "series_ids": "DTIN.MPV (First-home buyer TDTI bands)",
        "source_url": RBNZ_SOURCES["c40"]["pages"][0],
        "data_type": "HLF-derived share & HLF-calculated speed-limit headroom",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of new first-home buyer mortgage lending committed at a Total Debt-to-Income (TDTI) ratio above 6x gross income, compared against the Reserve Bank's 20% speed-limit allowance.",
        "methodology": "Derived from RBNZ Table C40 (Residential mortgage lending by Debt-to-Income purpose use).",
        "hlf_calculation": "High-DTI Share = (FHB Lending with TDTI > 6 / Total FHB Lending) * 100. Unused Headroom = 20.0% - High-DTI Share.",
        "questions": [
            "Are Reserve Bank DTI rules limiting first-home buyers in New Zealand?",
            "What percentage of NZ first-home buyers borrow more than 6 times their income?",
            "How much headroom do banks have under the 20% DTI speed limit?",
        ],
        "article_title": "Being Able to Afford a House vs Feeling Safe Buying One",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/the-difference-between-being-able-to-afford-a-house-and-feeling-safe-buying-one",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "borrower-income": {
        "name": "Average gross borrower income for approved first-home buyers vs investors",
        "topic": "Borrower income & servicing",
        "source": "Reserve Bank of New Zealand Table C41",
        "series_ids": "DTINI.MPA (Borrower Gross Income by TDTI band)",
        "source_url": RBNZ_SOURCES["c41"]["pages"][0],
        "data_type": "HLF-derived borrower-weighted harmonic average (from RBNZ C41 income bands)",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The borrower-weighted average annual pre-tax gross income reported on approved New Zealand first-home buyer and property investor mortgage applications (combining single and joint applicants).",
        "methodology": "Computed from RBNZ Table C41 (Residential mortgage borrower gross income) across all Total Debt-to-Income bands.",
        "hlf_calculation": "Total gross income dollars across all TDTI bands divided by implied borrower count (Band Gross Income NZDm / Band Average Gross Income NZD000s).",
        "questions": [
            "What is the average income of a first-home buyer in New Zealand?",
            "How much do approved property investors earn on average in NZ?",
        ],
        "article_title": "How Banks Assess Your Income for a New Zealand Mortgage",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-income-akl-vs-nz": {
        "name": "Approved first-home buyer gross income — New Zealand vs Auckland",
        "topic": "First-home lending & income",
        "source": "Reserve Bank of New Zealand Table C41",
        "series_ids": "DTINI.MPA (Nationwide & Auckland FHB series)",
        "source_url": RBNZ_SOURCES["c41"]["pages"][0],
        "data_type": "HLF-derived borrower-weighted average & HLF interpretation",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "Comparison of borrower-weighted average pre-tax household income on approved first-home mortgage applications nationwide versus Auckland.",
        "methodology": "Derived from RBNZ Table C41 nationwide and Auckland first-home buyer gross income series.",
        "hlf_calculation": "Harmonic borrower-weighted mean across TDTI bands for all NZ first-home buyers and Auckland first-home buyers.",
        "questions": [
            "What is the average first-home buyer household income in Auckland vs New Zealand?",
            "Can I buy a first home in NZ if my income is below the national average?",
        ],
        "article_title": "Buying a First Home on a Single or Modest Income",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "settled-rates": {
        "name": "Weighted-average settled 1-year and 2-year fixed mortgage interest rates",
        "topic": "Mortgage interest rates",
        "source": "Reserve Bank of New Zealand Table B30",
        "series_ids": "NCFA.MAM1PR.D1, NCFA.MAM1PR.D2",
        "source_url": RBNZ_SOURCES["b30"]["pages"][0],
        "data_type": "Officially reported RBNZ weighted-average settled rates",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The actual dollar-weighted average fixed mortgage interest rates locked in across all new residential mortgages settled by New Zealand registered banks for 1-year and 2-year fixed terms (including unadvertised bank rate discounts).",
        "methodology": "Directly reported by RBNZ in Table B30 (New residential mortgage weighted-average interest rates) for owner-occupiers/borrowers.",
        "hlf_calculation": "Officially reported RBNZ series NCFA.MAM1PR.D1 (1-year fixed) and NCFA.MAM1PR.D2 (2-year fixed).",
        "questions": [
            "What is the actual average 1-year and 2-year mortgage rate Kiwis are getting?",
            "How do settled mortgage rates compare with advertised bank special rates?",
            "Is my bank's rollover mortgage rate competitive?",
        ],
        "article_title": "Why Fixed Mortgage Rates Don't Simply Follow the OCR",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/why-fixed-mortgage-rates-dont-follow-the-ocr-nz",
        "tool_title": "Compare NZ Mortgage Rates & Historical Tracker",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
    },
    "new-builds-annual": {
        "name": "Annual new residential dwellings consented in New Zealand",
        "topic": "New builds & housing supply",
        "source": "Stats NZ — Building Consents Issued",
        "series_ids": "Stats NZ Building Consents Monthly Release",
        "source_url": "https://www.stats.govt.nz/",
        "data_type": "Officially reported Stats NZ figures",
        "derived_by_hlf": False,
        "unit": "dwellings",
        "frequency": "Monthly",
        "definition": "The total number of new residential dwellings consented across New Zealand over the trailing 12-month period and in the latest single month.",
        "methodology": "Compiled monthly by Stats NZ from building consent data supplied by New Zealand's 67 territorial authorities.",
        "hlf_calculation": "Directly reported Stats NZ annual and monthly dwelling consent totals and year-on-year percentage change.",
        "questions": [
            "How many new homes are being built or consented in New Zealand each year?",
            "Are New Zealand building consents increasing or decreasing?",
        ],
        "article_title": "Guide to New-Build & Turnkey Construction Loans",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "townhouse-share": {
        "name": "Multi-unit and townhouse share of new residential building consents",
        "topic": "New builds & housing supply",
        "source": "Stats NZ — Building Consents Issued",
        "series_ids": "Stats NZ Dwelling Type Breakdown",
        "source_url": "https://www.stats.govt.nz/",
        "data_type": "HLF-derived share (from officially reported Stats NZ dwelling counts)",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of new residential building consents in New Zealand issued for multi-unit homes (townhouses, flats, units, apartments, and retirement village units) versus stand-alone houses.",
        "methodology": "Derived from Stats NZ monthly building consent counts by dwelling type.",
        "hlf_calculation": "Multi-Unit Share = ((Total Consents - Stand-Alone Houses) / Total Consents) * 100.",
        "questions": [
            "What percentage of new homes built in New Zealand are townhouses or multi-units?",
            "Are more townhouses than stand-alone houses being consented in NZ?",
        ],
        "article_title": "Turnkey vs Construction Loans for Townhouses & New Builds",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Turnkey vs. Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/turnkey-vs-construction-loan-calculator-nz",
    },
    "build-cost-inflation": {
        "name": "Annual residential construction cost inflation (CPI purchase of newly built home)",
        "topic": "New builds & construction costs",
        "source": "Stats NZ — Consumers Price Index (CPI)",
        "series_ids": "Stats NZ CPI Housing & Household Utilities",
        "source_url": "https://www.stats.govt.nz/",
        "data_type": "Officially reported Stats NZ quarterly CPI inflation",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in the cost of purchasing a newly built home in New Zealand as measured by the Stats NZ Consumers Price Index (CPI).",
        "methodology": "Published quarterly by Stats NZ within the Consumers Price Index release.",
        "hlf_calculation": "Officially reported Stats NZ annual percentage change for purchase of new housing alongside headline CPI.",
        "questions": [
            "How fast are house building costs rising in New Zealand?",
            "Has residential construction cost inflation slowed down in NZ?",
        ],
        "article_title": "Managing Fixed-Price vs Cost-Escalation Build Contracts",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "rent-yield-pulse": {
        "name": "New Zealand national and regional median weekly private rent",
        "topic": "Rents & property investment yields",
        "source": "MBIE Tenancy Services — Rental Bond Data",
        "series_ids": "Tenancy Services Monthly Regional Bond CSV (Median Rent)",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported median rents & HLF-calculated YoY comparison",
        "derived_by_hlf": True,
        "unit": "NZD/week",
        "frequency": "Monthly",
        "definition": "The median weekly rent recorded on newly lodged private residential tenancy bonds across New Zealand and in Auckland, Wellington, and Canterbury.",
        "methodology": "Compiled by the Ministry of Business, Innovation and Employment (MBIE) from private rental bonds lodged with Tenancy Services.",
        "hlf_calculation": "Directly reported national ('ALL') and regional median weekly rents, plus HLF-calculated year-on-year percentage change.",
        "questions": [
            "What is the median weekly rent in New Zealand right now?",
            "How much is median weekly rent in Auckland, Wellington, and Christchurch/Canterbury?",
            "Are New Zealand weekly rents rising or flat?",
        ],
        "article_title": "Rental Yield & Cashflow Guide for NZ Investors",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "rental-bonds-supply": {
        "name": "Monthly newly lodged and active private rental bonds in New Zealand",
        "topic": "Rental supply & tenancy activity",
        "source": "MBIE Tenancy Services — Rental Bond Data",
        "series_ids": "Tenancy Services Monthly Regional Bond CSV (Lodged & Active Bonds)",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported MBIE Tenancy Services bond counts",
        "derived_by_hlf": False,
        "unit": "bonds",
        "frequency": "Monthly",
        "definition": "The number of newly lodged private rental tenancy bonds in the latest month alongside the total pool of active private rental bonds held across New Zealand.",
        "methodology": "Published monthly by MBIE Tenancy Services from statutory residential tenancy bond lodgements.",
        "hlf_calculation": "Officially reported 'ALL' location Lodged Bonds and Active Bonds counts.",
        "questions": [
            "How many private rental properties and tenancies are active in New Zealand?",
            "How many new rental bonds are lodged each month in NZ?",
        ],
        "article_title": "NZ Rental Market Supply & Investor Insights",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "student-loan-servicing": {
        "name": "New Zealand median student loan balance",
        "topic": "Borrowing power & student loans",
        "source": "Inland Revenue — Student Loan Scheme Statistics",
        "series_ids": "IRD Student Loan Balances Release",
        "source_url": IRD_SL_BAL_PAGE,
        "data_type": "Officially reported Inland Revenue median & HLF-calculated YoY change",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Quarterly",
        "definition": "The median student loan balance across New Zealand student loan borrowers reported by Inland Revenue (Te Tari Taake).",
        "methodology": "Published by Inland Revenue from administrative Student Loan Scheme records; compulsory 12% PAYE repayments above the repayment threshold reduce net monthly mortgage servicing capacity.",
        "hlf_calculation": "Directly reported IRD median balance and year-on-year percentage change.",
        "questions": [
            "What is the average or median student loan balance in New Zealand?",
            "How does a student loan affect my mortgage borrowing power in NZ?",
        ],
        "article_title": "How Student Loans Impact Your Mortgage Borrowing Power",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "kiwisaver-first-home": {
        "name": "Monthly KiwiSaver first-home withdrawal count, value & average per buyer",
        "topic": "KiwiSaver & first-home deposits",
        "source": "Inland Revenue — KiwiSaver Withdrawal Statistics (cross-referenced with RBNZ C31)",
        "series_ids": "IRD Monthly KiwiSaver First-Home Withdrawals & RBNZ LVRN.MMB2.AC1",
        "source_url": IRD_KS_DATA_PAGE,
        "data_type": "Officially reported IRD totals & HLF-derived per-member / per-loan averages",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The monthly number of KiwiSaver members withdrawing funds to purchase a first home, total dollars withdrawn, and HLF-derived average withdrawal per member and per approved first-home mortgage.",
        "methodology": "Compiled monthly by Inland Revenue from KiwiSaver scheme provider reports and cross-referenced by HLF against RBNZ Table C31 monthly first-home mortgage approvals.",
        "hlf_calculation": "Avg per Member = Total First-Home Withdrawal Dollars / Withdrawing Members. Avg per Approved FHB Loan = Total Withdrawal Dollars / RBNZ C31 FHB Borrowers.",
        "questions": [
            "How much is the average KiwiSaver first-home withdrawal in New Zealand?",
            "How many Kiwis use KiwiSaver to buy their first home each month?",
        ],
        "article_title": "Using Your KiwiSaver for a First-Home Deposit",
        "article_url": "https://www.homeloanfactory.co.nz/kiwisaver",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "fhb-regional-kiwisaver": {
        "name": "Regional KiwiSaver first-home withdrawal averages across New Zealand",
        "topic": "KiwiSaver & regional first-home deposits",
        "source": "Inland Revenue — KiwiSaver Fund Withdrawals by Region",
        "series_ids": "IRD Annual Regional KiwiSaver First-Home Purchase Withdrawals",
        "source_url": IRD_KS_REG_PAGE,
        "data_type": "Officially reported IRD regional totals & HLF-derived per-person / couple benchmarks",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual",
        "definition": "Average KiwiSaver first-home withdrawal amount per withdrawing member (and two-member couple benchmark) across New Zealand's 16 regions.",
        "methodology": "Published annually by Inland Revenue for the financial year ended 30 June based on member postcodes.",
        "hlf_calculation": "Regional Avg per Person = Regional First-Home Withdrawal Dollars / Regional Withdrawing Members. Couple Benchmark = 2 * Regional Avg per Person.",
        "questions": [
            "How much KiwiSaver do first-home buyers withdraw in Wellington, Auckland, and Canterbury?",
            "Which New Zealand region has the highest average KiwiSaver first-home withdrawal?",
        ],
        "article_title": "How Kiwis Piece Together a First-Home Deposit by Region",
        "article_url": "https://www.homeloanfactory.co.nz/kiwisaver",
        "tool_title": "HLF KiwiSaver & First-Home Advice",
        "tool_url": "https://www.homeloanfactory.co.nz/kiwisaver",
    },
    "kiwisaver-switches": {
        "name": "Annual and monthly KiwiSaver scheme transfers (members switching providers)",
        "topic": "KiwiSaver scheme transfers",
        "source": "Inland Revenue — KiwiSaver Scheme Transfer Statistics",
        "series_ids": "IRD Annual KiwiSaver Scheme Transfers",
        "source_url": IRD_KS_TR_PAGE,
        "data_type": "Officially reported IRD annual total & HLF-derived monthly average",
        "derived_by_hlf": True,
        "unit": "transfers",
        "frequency": "Annual",
        "definition": "The total number of New Zealand KiwiSaver members who transferred their balance to a different KiwiSaver scheme provider over the June financial year, plus the monthly average.",
        "methodology": "Compiled by Inland Revenue from KiwiSaver scheme transfers processed during the year ended 30 June.",
        "hlf_calculation": "Directly reported annual scheme transfers divided by 12 to derive average monthly provider switches.",
        "questions": [
            "How many New Zealanders switch KiwiSaver providers each year?",
            "Is it common to change KiwiSaver funds or providers in NZ?",
        ],
        "article_title": "How to Review Your KiwiSaver Fund & Provider",
        "article_url": "https://www.homeloanfactory.co.nz/kiwisaver",
        "tool_title": "Book a Free HLF KiwiSaver Review",
        "tool_url": "https://www.homeloanfactory.co.nz/kiwisaver",
    },
}
# Alias keys share the same catalog metadata as their primary key
TICKER_CATALOG["investor-vs-owner-io"] = dict(TICKER_CATALOG["interest-only"], name="Investor vs owner-occupier interest-only mortgage share")
TICKER_CATALOG["kiwisaver-active-switchers"] = dict(TICKER_CATALOG["kiwisaver-switches"], name="Active KiwiSaver provider switching volume")


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
    """Extract the most recent n historical observations (latest first) for reference tables."""
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
# Multi-Sheet RBNZ Workbook Scanner
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
# Conversational, Story-Driven, FMA-Safe Parsers + Historical Series
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

    # Full time-series vectors for 13-month history tables
    ser_low_dep_pct = (fhb_hi_cnt / fhb_cnt) * 100.0
    ser_fhb_sh_pct = (fhb_val / tot_val) * 100.0
    ser_inv_sh_pct = (inv_val / tot_val) * 100.0
    ser_fhb_hi_avg_k = (fhb_hi_val * 1000.0) / fhb_hi_cnt.replace(0, float("nan"))
    ser_fhb_lo_avg_k = ((fhb_val - fhb_hi_val) * 1000.0) / (fhb_cnt - fhb_hi_cnt).replace(0, float("nan"))
    ser_fhb_gap_k = ser_fhb_hi_avg_k - ser_fhb_lo_avg_k
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

    fhb_sh = float(ser_fhb_sh_pct.iloc[i])
    fhb_sh_yoy = float(ser_fhb_sh_pct.iloc[iy])
    fhb_avg_k = int(round((float(fhb_val.iloc[i]) * 1000.0) / float(fhb_cnt.iloc[i])))

    inv_sh = float(ser_inv_sh_pct.iloc[i])
    inv_sh_yoy = float(ser_inv_sh_pct.iloc[iy])
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

    fhb_hi_val_share = float(ser_fhb_prio_pct.iloc[i])
    fhb_hi_cnt_share = (float(fhb_hi_cnt.iloc[i]) / max(1.0, float(all_hi_cnt.iloc[i]))) * 100.0
    fhb_priority_html = (
        f"When New Zealand banks approve low-deposit mortgages (&gt;80% LVR), most of their allowance goes to first-home buyers. "
        f"In {m_full}, <strong>{fhb_hi_val_share:.1f}% of all low-deposit mortgage dollars "
        f"({fmt_nzd_millions(float(fhb_hi_val.iloc[i]))} of {fmt_nzd_millions(float(all_hi_val.iloc[i]))})</strong> and "
        f"<strong>{fhb_hi_cnt_share:.1f}% of low-deposit loans ({fhb_hi_loans_int:,} of {int(round(float(all_hi_cnt.iloc[i]))):,})</strong> "
        f"went directly to first-home buyers."
    )

    std_dep_pct = float(ser_fhb_20club_pct.iloc[i])
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
    oo_20eq_pct = float(ser_oo_20eq_pct.iloc[i])
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
    inv_u30_pct = float(ser_inv_u30_pct.iloc[i])
    inv_u30_val = float(inv_gt70_val.iloc[i])
    inv_eq_html = (
        f"You don't need 40% equity locked up to buy an investment property. In {m_full}, <strong>{inv_40plus_pct:.1f}%</strong> "
        f"of approved NZ property investors had 40%+ equity—while <strong>{inv_30to40_pct:.1f}%</strong> bought in the "
        f"<strong>30%–40% equity band</strong> and <strong>{inv_u30_pct:.1f}% ({inv_u30_cnt:,} loans, or {fmt_nzd_millions(inv_u30_val)})</strong> "
        f"were approved with <strong>under 30% equity</strong>."
    )

    inv_gt80_c = int(round(float(inv_gt80_cnt.iloc[i])))
    inv_gt80_v = float(inv_gt80_val.iloc[i])
    inv_gt80_c_prev = int(round(float(inv_gt80_cnt.iloc[iy])))
    inv_gt80_v_prev = float(inv_gt80_val.iloc[iy])
    inv_gt80_pct = (inv_gt80_c / max(1.0, inv_c_tot)) * 100.0
    inv_70to80_c = max(0, inv_u30_cnt - inv_gt80_c)
    inv_70to80_v = max(0.0, inv_u30_val - inv_gt80_v)
    inv_70to80_pct = (inv_70to80_c / max(1.0, inv_c_tot)) * 100.0

    inv_12m_clause = ""
    if i >= 11:
        inv_gt80_12m_c = int(round(float(inv_gt80_cnt.iloc[i - 11 : i + 1].sum())))
        inv_gt80_12m_v = float(inv_gt80_val.iloc[i - 11 : i + 1].sum())
        inv_12m_clause = f" (and <strong>{inv_gt80_12m_c:,} loans / {fmt_nzd_millions(inv_gt80_12m_v)}</strong> over the past 12 months)"

    inv_dir_prev = "up from" if inv_gt80_c >= inv_gt80_c_prev else "compared with"
    inv_low_dep_html = (
        f"Even for property investors, low-deposit approvals happen every month through new-build exemptions and bank speed limits. "
        f"In {m_full}, New Zealand banks approved <strong>{inv_gt80_c:,} investor mortgages ({fmt_nzd_millions(inv_gt80_v)}, or "
        f"{inv_gt80_pct:.1f}% of investors)</strong> with <strong>strictly under a 20% deposit (&gt;80% LVR)</strong>—{inv_dir_prev} "
        f"{inv_gt80_c_prev:,} loans ({fmt_nzd_millions(inv_gt80_v_prev)}) a year earlier{inv_12m_clause}—plus another "
        f"<strong>{inv_70to80_c:,} investor loans ({fmt_nzd_millions(inv_70to80_v)}, or {inv_70to80_pct:.1f}%)</strong> in the "
        f"<strong>20%–30% deposit tier (70%–80% LVR)</strong> where standard 20%-deposit new-build investment loans sit."
    )

    inv_nb_lvr_html = (
        f"Because brand-new builds are exempt from Reserve Bank LVR restrictions, nearly one in five property investors buys with less "
        f"than the standard 30% existing-property deposit. In {m_full}, <strong>{inv_u30_cnt:,} investor loans "
        f"({fmt_nzd_millions(inv_u30_val)}, or {inv_u30_pct:.1f}% of investor approvals)</strong> were approved with "
        f"<strong>under a 30% deposit (&gt;70% LVR)</strong>—comprising <strong>{inv_70to80_c:,} loans ({fmt_nzd_millions(inv_70to80_v)})</strong> "
        f"in the <strong>20%–30% deposit band</strong> (which includes 20%-deposit new builds) and <strong>{inv_gt80_c:,} loans "
        f"({fmt_nzd_millions(inv_gt80_v)})</strong> with <strong>under a 20% deposit (&gt;80% LVR)</strong>."
    )

    oo_low_eq_pct = float(ser_oo_low_eq_pct.iloc[i])
    oo_low_eq_prev = float(ser_oo_low_eq_pct.iloc[iy])
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

    def make_c31_item(html_str: str, val_num: float, val_fmt: str, prev_fmt: str, chg_str: str, hist_list: list) -> dict:
        return {
            "chip": "RBNZ C31",
            "as_of": as_of,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": html_str,
            "source_html": src_c31,
            "latest_period": m_full,
            "value": round(val_num, 2),
            "latest_value": val_fmt,
            "previous_period": f"{dates.iloc[iy].strftime('%B %Y')} ({prev_fmt})",
            "change": chg_str,
            "history": hist_list,
        }

    out = {
        "fhb-low-deposit": make_c31_item(
            low_dep_html, low_cnt_pct, f"{low_cnt_pct:.1f}%", f"{low_cnt_yoy:.1f}%",
            f"{low_cnt_pct - low_cnt_yoy:+.1f} percentage points YoY",
            make_history_list(dates, ser_low_dep_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-share": make_c31_item(
            fhb_share_html, fhb_sh, f"{fhb_sh:.1f}% ({fhb_loans_int:,} loans)", f"{fhb_sh_yoy:.1f}%",
            f"{fhb_sh - fhb_sh_yoy:+.1f} percentage points YoY",
            make_history_list(dates, ser_fhb_sh_pct, lambda x: f"{x:.1f}%"),
        ),
        "investors": make_c31_item(
            inv_html, inv_sh, f"{inv_sh:.1f}% (${inv_b:.2f}B)", f"{inv_sh_yoy:.1f}%",
            f"{inv_sh - inv_sh_yoy:+.1f} percentage points YoY",
            make_history_list(dates, ser_inv_sh_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-loan-size-gap": make_c31_item(
            fhb_gap_html, loan_gap / 1000.0, f"${avg_hi_loan:,.0f} (<20% dep) vs ${avg_lo_loan:,.0f} (20%+ dep)",
            f"${float(ser_fhb_hi_avg_k.iloc[iy])*1000:,.0f} vs ${float(ser_fhb_lo_avg_k.iloc[iy])*1000:,.0f}",
            f"${abs(loan_gap):,.0f} gap between deposit tiers",
            make_history_list(dates, ser_fhb_hi_avg_k, lambda x: f"${x:,.1f}k (<20% dep avg)"),
        ),
        "fhb-bank-lvr-priority": make_c31_item(
            fhb_priority_html, fhb_hi_val_share, f"{fhb_hi_val_share:.1f}% of >80% LVR dollars ({fhb_hi_cnt_share:.1f}% of loans)",
            f"{float(ser_fhb_prio_pct.iloc[iy]):.1f}%",
            f"{fhb_hi_val_share - float(ser_fhb_prio_pct.iloc[iy]):+.1f} percentage points YoY",
            make_history_list(dates, ser_fhb_prio_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-20-deposit-club": make_c31_item(
            fhb_20club_html, std_dep_pct, f"{std_dep_pct:.1f}% ({int(round(fhb_lo_cnt_cur)):,} loans, avg ${avg_lo_loan:,.0f})",
            f"{float(ser_fhb_20club_pct.iloc[iy]):.1f}%",
            f"{std_dep_pct - float(ser_fhb_20club_pct.iloc[iy]):+.1f} percentage points YoY",
            make_history_list(dates, ser_fhb_20club_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-vs-investors": make_c31_item(
            fhb_vs_inv_html, diff_fi, f"FHB {fmt_nzd_millions(fhb_v)} vs Investors {fmt_nzd_millions(inv_v)}",
            f"{fmt_nzd_millions(float(ser_fhb_vs_inv_m.iloc[iy]))} FHB minus Investor gap",
            f"{fmt_nzd_millions(diff_fi)} FHB lead in {m_full}",
            make_history_list(dates, ser_fhb_vs_inv_m, lambda x: f"{x:+.0f}m (FHB vs Inv)"),
        ),
        "owner-equity-buffer": make_c31_item(
            owner_eq_html, oo_20eq_pct, f"{oo_20eq_pct:.1f}% (avg loan ${oo_avg_20eq:,.0f})",
            f"{float(ser_oo_20eq_pct.iloc[iy]):.1f}%",
            f"{oo_20eq_pct - float(ser_oo_20eq_pct.iloc[iy]):+.1f} percentage points YoY",
            make_history_list(dates, ser_oo_20eq_pct, lambda x: f"{x:.1f}%"),
        ),
        "investor-equity-tiers": make_c31_item(
            inv_eq_html, inv_40plus_pct, f"{inv_40plus_pct:.1f}% (40%+ eq) | {inv_30to40_pct:.1f}% (30–40% eq) | {inv_u30_pct:.1f}% (<30% eq)",
            f"{100.0 - float(ser_inv_u30_pct.iloc[iy]):.1f}% with 30%+ equity",
            f"{inv_u30_pct - float(ser_inv_u30_pct.iloc[iy]):+.1f} pts shift in <30% equity share",
            make_history_list(dates, 100.0 - ser_inv_u30_pct, lambda x: f"{x:.1f}% (30%+ equity)"),
        ),
        "investor-low-deposit": make_c31_item(
            inv_low_dep_html, float(inv_gt80_c), f"{inv_gt80_c:,} loans (${inv_gt80_v:,.0f}m) <20% dep + {inv_70to80_c:,} loans (${inv_70to80_v:,.0f}m) 20–30% dep",
            f"{inv_gt80_c_prev:,} loans (${inv_gt80_v_prev:,.0f}m) <20% dep",
            f"{inv_gt80_c - inv_gt80_c_prev:+d} loans ({inv_gt80_v - inv_gt80_v_prev:+.0f}m) YoY in >80% LVR tier",
            make_history_list(dates, inv_gt80_cnt, lambda x: f"{int(round(x)):,} loans (>80% LVR)"),
        ),
        "investor-new-build-lvr": make_c31_item(
            inv_nb_lvr_html, inv_u30_pct, f"{inv_u30_pct:.1f}% ({inv_u30_cnt:,} loans / ${inv_u30_val:,.0f}m <30% deposit)",
            f"{float(ser_inv_u30_pct.iloc[iy]):.1f}%",
            f"{inv_u30_pct - float(ser_inv_u30_pct.iloc[iy]):+.1f} percentage points YoY",
            make_history_list(dates, ser_inv_u30_pct, lambda x: f"{x:.1f}% (<30% dep)"),
        ),
        "mover-low-equity": make_c31_item(
            mover_low_eq_html, oo_low_eq_pct, f"{oo_low_eq_pct:.1f}% ({int(round(oo_hi_c)):,} loans / ${float(oo_hi_val.iloc[i]):,.0f}m)",
            f"{oo_low_eq_prev:.1f}%",
            f"{oo_low_eq_pct - oo_low_eq_prev:+.1f} percentage points YoY",
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
            f"Over the 12 months to {m_full}, New Zealand banks approved <strong>{r12_cnt:,} first-home buyer mortgages</strong> "
            f"totalling <strong>{fmt_nzd_millions(r12_val)} in lending</strong> ({dir_r12} {r12_prev_cnt:,} loans a year earlier)—including "
            f"<strong>{r12_hi_cnt:,} buyers ({r12_hi_pct:.1f}%)</strong> who bought with <strong>under a 20% deposit</strong>.",
            float(r12_cnt), f"{r12_cnt:,} loans ({fmt_nzd_millions(r12_val)})", f"{r12_prev_cnt:,} loans",
            f"{r12_cnt - r12_prev_cnt:+,d} annual FHB loans YoY",
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

    tot_v_ser = top_v + pur_v + ref_v + oth_v
    tot_c_ser = top_c + pur_c + ref_c + oth_c
    ref_sh_ser = (ref_v / tot_v_ser) * 100.0
    top_sh_c_ser = (top_c / tot_c_ser) * 100.0

    ref_sh = float(ref_sh_ser.iloc[j])
    ref_avg_k = int(round((float(ref_v.iloc[j]) * 1000.0) / float(ref_c.iloc[j])))
    top_sh_c = float(top_sh_c_ser.iloc[j])
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
            "latest_period": m_full,
            "value": round(ref_sh, 2),
            "latest_value": f"{ref_sh:.1f}% (${float(ref_v.iloc[j])/1000.0:.2f}B / {int(round(float(ref_c.iloc[j]))):,} loans)",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({float(ref_sh_ser.iloc[jy]):.1f}%)",
            "change": f"{ref_sh - float(ref_sh_ser.iloc[jy]):+.1f} percentage points YoY",
            "history": make_history_list(dates, ref_sh_ser, lambda x: f"{x:.1f}%"),
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
            "latest_period": m_full,
            "value": round(top_sh_c, 2),
            "latest_value": f"{top_sh_c:.1f}% ({top_cnt_int:,} top-ups, avg ${top_avg_k:,}k)",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({float(top_sh_c_ser.iloc[jy]):.1f}%)",
            "change": f"{top_sh_c - float(top_sh_c_ser.iloc[jy]):+.1f} percentage points YoY",
            "history": make_history_list(dates, top_sh_c_ser, lambda x: f"{x:.1f}%"),
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
            "latest_period": m_full,
            "value": round(top_val_m, 1),
            "latest_value": f"{fmt_nzd_millions(top_val_m)} ({top_cnt_int:,} top-ups)",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({fmt_nzd_millions(float(top_v.iloc[jy]))})",
            "change": f"{top_val_m - float(top_v.iloc[jy]):+.0f}m YoY",
            "history": make_history_list(dates, top_v, lambda x: f"${x:,.0f}m"),
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
    iy = max(0, len(dates) - 13)
    inv_io_ser = (inv_io / inv_tot) * 100.0
    oo_io_ser = (oo_io / oo_tot) * 100.0
    inv_io_pct = float(inv_io_ser.iloc[-1])
    oo_io_pct = float(oo_io_ser.iloc[-1])

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
        "latest_period": m_full,
        "value": round(inv_io_pct, 2),
        "latest_value": f"{inv_io_pct:.1f}% investor IO vs {oo_io_pct:.1f}% owner-occupier IO",
        "previous_period": f"{dates.iloc[iy].strftime('%B %Y')} ({float(inv_io_ser.iloc[iy]):.1f}% investor IO)",
        "change": f"{inv_io_pct - float(inv_io_ser.iloc[iy]):+.1f} percentage points YoY (investor IO)",
        "history": make_history_list(dates, inv_io_ser, lambda x: f"{x:.1f}% (investor IO)"),
    }
    return {
        "interest-only": io_item,
        "investor-vs-owner-io": dict(io_item),
    }


def compute_c41_weighted_avg_income(tabs41: list, borrower_kw: str, region_mode: str = "nz") -> tuple[pd.Timestamp | None, float | None, pd.Series | None, pd.Series | None]:
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
            tot_inc_ser = pd.Series(0.0, index=df41.index)
            tot_cnt_ser = pd.Series(0.0, index=df41.index)
            for bk in common:
                a_m = pd.to_numeric(df41.iloc[:, inc_cols[bk]], errors="coerce")
                b_k = pd.to_numeric(df41.iloc[:, avg_cols[bk]], errors="coerce")
                if len(b_k.dropna()) and b_k.dropna().median() > 2000:
                    b_k = b_k / 1000.0
                valid = (a_m > 0) & (b_k > 0)
                tot_inc_ser = tot_inc_ser.add(a_m.where(valid, 0.0), fill_value=0.0)
                tot_cnt_ser = tot_cnt_ser.add((a_m / b_k).where(valid, 0.0), fill_value=0.0)

            weighted_k_ser = tot_inc_ser / tot_cnt_ser.replace(0.0, float("nan"))
            valid_w = weighted_k_ser.dropna()
            if len(valid_w):
                return df41["date"].iloc[-1], float(valid_w.iloc[-1]) * 1000.0, df41["date"], weighted_k_ser
    return None, None, None, None


def build_rbnz_dti_income_tickers(c40_path: pathlib.Path | None, c41_path: pathlib.Path | None) -> dict:
    out = {}
    if c40_path is not None:
        try:
            tabs40 = load_rbnz_sheets(c40_path)
            val_pct, dt_dti, dti_dates, dti_ser = None, None, None, None

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
                        break

            if val_pct is not None and dt_dti is not None:
                m_full = dt_dti.strftime("%B %Y")
                diff_cap = 20.0 - val_pct
                iy = max(0, len(dti_dates) - 13)
                prev_v = float(dti_ser.iloc[iy]) if pd.notna(dti_ser.iloc[iy]) else val_pct
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
                    "latest_period": m_full,
                    "value": round(val_pct, 2),
                    "latest_value": f"{val_pct:.1f}% (>6x DTI, {diff_cap:.1f} pts below 20% cap)",
                    "previous_period": f"{dti_dates.iloc[iy].strftime('%B %Y')} ({prev_v:.1f}%)",
                    "change": f"{val_pct - prev_v:+.1f} percentage points YoY",
                    "history": make_history_list(dti_dates, dti_ser, lambda x: f"{x:.1f}%"),
                }
        except Exception as exc:
            print(f"[C40 DTI] Parse note: {exc}")

    if c41_path is not None:
        try:
            tabs41 = load_rbnz_sheets(c41_path)
            dt_fhb, fhb_inc, d41_dates, fhb_inc_ser = compute_c41_weighted_avg_income(tabs41, "first home", "nz")
            _, fhb_akl_inc, _, _ = compute_c41_weighted_avg_income(tabs41, "first home", "akl")
            _, inv_inc, _, _ = compute_c41_weighted_avg_income(tabs41, "investor", "nz")

            src_c41 = (
                'Source: <a href="https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi" '
                'target="_blank" rel="noopener">RBNZ Table C41</a>, &copy; Reserve Bank of New Zealand. Derived by HLF for general information only.'
            )

            if dt_fhb is not None and fhb_inc and 40000 <= fhb_inc <= 600000:
                m_c41 = dt_fhb.strftime("%B %Y")
                as_of_c41 = month_end_iso(dt_fhb)
                iy41 = max(0, len(d41_dates) - 13)
                prev_inc_k = float(fhb_inc_ser.iloc[iy41]) if pd.notna(fhb_inc_ser.iloc[iy41]) else (fhb_inc / 1000.0)
                hist_41 = make_history_list(d41_dates, fhb_inc_ser, lambda x: f"${x:,.1f}k")

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
                    "latest_period": m_c41,
                    "value": round(fhb_inc, 0),
                    "latest_value": f"${fhb_inc/1000.0:,.1f}k (FHB)" + (f" vs ${inv_inc/1000.0:,.1f}k (Investor)" if inv_inc else ""),
                    "previous_period": f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_k:,.1f}k FHB)",
                    "change": f"{(fhb_inc/1000.0) - prev_inc_k:+.1f}k YoY",
                    "history": hist_41,
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
                    "latest_period": m_c41,
                    "value": round(fhb_inc, 0),
                    "latest_value": f"${fhb_inc:,.0f} NZ" + (f" | ${fhb_akl_inc:,.0f} Auckland" if fhb_akl_inc else ""),
                    "previous_period": f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_k*1000:,.0f} NZ)",
                    "change": f"{(fhb_inc/1000.0) - prev_inc_k:+.1f}k YoY",
                    "history": hist_41,
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
    iy = max(0, len(dates) - 13)
    m_full = dt.strftime("%B %Y")
    v1, v2 = float(r_1y.iloc[-1]), float(r_2y.iloc[-1])
    v1_prev = float(r_1y.iloc[iy])

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
            "latest_period": m_full,
            "value": round(v1, 2),
            "latest_value": f"{v1:.2f}% (1-year settled) | {v2:.2f}% (2-year settled)",
            "previous_period": f"{dates.iloc[iy].strftime('%B %Y')} ({v1_prev:.2f}% 1-year settled)",
            "change": f"{v1 - v1_prev:+.2f}% YoY (1-year settled)",
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
                "latest_period": f"Year ended {m_full}",
                "value": ann_tot,
                "latest_value": f"{ann_tot:,} annual consents ({mth_tot:,} in {m_full})",
                "previous_period": "Prior 12-month period",
                "change": f"{dir_word.capitalize()} {ann_pct:.1f}% YoY" if dir_word else "Reported by Stats NZ",
            }
            out["townhouse-share"] = {
                "chip": "Stats NZ",
                "as_of": as_of,
                "max_age_days": MAX_AGE_MONTHLY,
                "html": th_html,
                "source_html": src_bc,
                "latest_period": m_full,
                "value": round(multi_pct, 1),
                "latest_value": f"{multi_pct:.1f}% ({multi_tot:,} of {mth_tot:,} consents)",
                "previous_period": f"Stand-alone houses: {houses:,}",
                "change": f"{townhouses:,} townhouses, flats & units in {m_full}",
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
                "latest_period": f"{qname.capitalize()} {yr} quarter",
                "value": signed_val,
                "latest_value": f"{'-' if is_down else '+'}{val:.1f}% YoY",
                "previous_period": f"Headline CPI: {cpi_all_val:.1f}% YoY" if cpi_all_val is not None else "Prior year quarter",
                "change": f"{'-' if is_down else '+'}{val:.1f}% annual change",
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

    yoy_rows = nat[nat["dt"] == (dt - pd.DateOffset(years=1))]
    prev_r = int(round(float(yoy_rows.iloc[-1]["Median Rent"]))) if len(yoy_rows) else nat_r
    yoy_pct = ((float(last["Median Rent"]) - prev_r) / prev_r * 100.0) if len(yoy_rows) and prev_r else 0.0
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

    lodged_int = int(round(float(last["Lodged Bonds"])))
    active_int = int(round(float(last["Active Bonds"])))

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
            "latest_period": m_full,
            "value": nat_r,
            "latest_value": f"${nat_r}/wk NZ (${akl_r} AKL | ${wgtn_r} WGTN | ${cant_r} CANT)",
            "previous_period": f"12 months earlier (${prev_r}/wk NZ)",
            "change": f"{yoy_pct:+.1f}% YoY (${nat_r - prev_r:+d}/wk)",
            "history": make_history_list(nat["dt"], nat["Median Rent"], lambda x: f"${int(round(x))}/wk"),
        },
        "rental-bonds-supply": {
            "chip": "Tenancy Services",
            "as_of": as_of,
            "max_age_days": MAX_AGE_TENANCY,
            "html": (
                f"Thousands of New Zealand households sign new leases every month—Tenancy Services logged "
                f"<strong>{lodged_int:,}</strong> newly lodged private rental bonds in {m_full}, "
                f"across a nationwide pool of <strong>{active_int:,}</strong> active private tenancies."
            ),
            "source_html": src_mbie,
            "latest_period": m_full,
            "value": active_int,
            "latest_value": f"{lodged_int:,} lodged bonds | {active_int:,} active bonds",
            "previous_period": "Monthly MBIE bond registry",
            "change": f"{lodged_int:,} new private tenancies lodged in {m_full}",
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
                    "latest_period": m_full,
                    "value": curr_bal,
                    "latest_value": f"${curr_bal:,} median balance",
                    "previous_period": f"{prev_lbl} (${prev_bal:,})",
                    "change": f"${curr_bal - prev_bal:+,d} ({((curr_bal - prev_bal)/max(1, prev_bal))*100:+.1f}%)",
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
                    "latest_period": m_full,
                    "value": round(amt_m, 1),
                    "latest_value": f"${amt_m:,.1f}m across {cnt:,} members (~${avg_mem_k:,}k/member)",
                    "previous_period": f"${prev_amt_m:,.1f}m in prior period" if prev_amt_m else "Prior monthly release",
                    "change": f"{amt_m - prev_amt_m:+.1f}m vs prior period" if prev_amt_m else f"~${avg_mem_k:,}k per withdrawing member",
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
                    nz_couple = int(round((nz_avg_p * 2) / 50.0) * 50)
                    wgtn_avg = dict((r[0], r[1]) for r in reg_list).get("Wellington", reg_list[0][1])
                    akl_avg = dict((r[0], r[1]) for r in reg_list).get("Auckland", reg_list[1][1])
                    reg_hist = [
                        {"period": r[0], "label": r[0], "value": r[1], "formatted": f"${r[1]:,}/person ({r[2]:,} members, ${r[3]:.2f}m)"}
                        for r in reg_list
                    ]
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
                        "latest_period": f"Year ended June {fy_yr_reg}",
                        "value": nz_avg_p,
                        "latest_value": f"${nz_avg_p:,}/person NZ (~${nz_couple:,}/couple across {tot_mem:,} members)",
                        "previous_period": f"Wellington: ${wgtn_avg:,}/person | Auckland: ${akl_avg:,}/person",
                        "change": f"${tot_nzd/1e9:.2f}B withdrawn across 16 NZ regions",
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
                "latest_period": f"Year ended June {fy_yr}",
                "value": tr_tot,
                "latest_value": f"{tr_tot:,} transfers/yr ({mth_avg:,}/month)",
                "previous_period": f"Year ended June {prev_fy} ({prev_tr:,} transfers)",
                "change": f"{tr_tot - prev_tr:+,d} transfers YoY",
            }
            out["kiwisaver-switches"] = sw_item
            out["kiwisaver-active-switchers"] = dict(sw_item)

    return out


# ---------------------------------------------------------------------------
# Reference-Layer Enricher & Canonical Page / JSON Builders
# ---------------------------------------------------------------------------
def enrich_ticker_record(key: str, item: dict, verified_date_iso: str) -> dict:
    meta = TICKER_CATALOG.get(key, {})
    plain_summary = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", item.get("html", ""))).strip()
    canonical_url = f"{PULSE_BASE_URL}/{key}/"

    citation_chain = {
        "official_source": meta.get("source", item.get("chip", "Official NZ Crown Release")),
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
        "topic": meta.get("topic", "New Zealand mortgage & housing statistics"),
        "source": meta.get("source", item.get("chip", "NZ Crown Data")),
        "source_series_ids": meta.get("series_ids", "Official agency series"),
        "source_url": meta.get("source_url", "https://www.rbnz.govt.nz/"),
        "period": item.get("as_of", "")[:7],
        "latest_period": item.get("latest_period", item.get("as_of", "")),
        "latest_value": item.get("latest_value", "See summary"),
        "unit": meta.get("unit", "value"),
        "previous_period": item.get("previous_period", "Prior comparable period"),
        "change": item.get("change", "Tracked across official releases"),
        "data_type": meta.get("data_type", "Officially reported & HLF-derived"),
        "derived_by_hlf": bool(meta.get("derived_by_hlf", True)),
        "definition": meta.get("definition", plain_summary),
        "methodology": meta.get("methodology", "Compiled from official New Zealand Crown statistical releases."),
        "hlf_calculation": meta.get("hlf_calculation", "Standardised and verified by HLF from official agency tables."),
        "update_frequency": meta.get("frequency", "Monthly"),
        "last_verified": verified_date_iso,
        "status": "Verified",
        "canonical_url": canonical_url,
        "summary_text": plain_summary,
        "related_questions": meta.get("questions", []),
        "citation_chain": citation_chain,
        "history": item.get("history", []),
    })
    return enriched


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


def write_canonical_ticker_page(root: pathlib.Path, key: str, rec: dict, now_human: str) -> None:
    t_dir = root / key
    t_dir.mkdir(parents=True, exist_ok=True)
    c_chain = rec["citation_chain"]
    q_items = "".join(f"<li>{html_lib.escape(q)}</li>" for q in rec.get("related_questions", []))
    q_block = f"<h2>Questions This Data Helps Answer</h2><ul>{q_items}</ul>" if q_items else ""
    hist_html = render_history_table_html(rec.get("history", []))
    hist_section = f"<h2>Historical &amp; Regional Values</h2>{hist_html}" if hist_html else ""

    schema_obj = {
        "@context": "https://schema.org",
        "@type": "Dataset",
        "name": f"{rec['name']} — HLF Pulse Reference Record ({rec['latest_period']})",
        "description": rec["definition"],
        "url": rec["canonical_url"],
        "identifier": key,
        "inLanguage": "en-NZ",
        "temporalCoverage": rec["as_of"],
        "spatialCoverage": "New Zealand",
        "isBasedOn": rec["source_url"],
        "creator": {"@type": "Organization", "name": "HLF (Home Loan Factory)", "url": "https://www.homeloanfactory.co.nz/"},
        "variableMeasured": rec["name"],
    }

    page_html = f"""<!DOCTYPE html>
<html lang="en-NZ">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="robots" content="index, follow">
  <link rel="canonical" href="{rec['canonical_url']}">
  <title>{html_lib.escape(rec['name'])} ({html_lib.escape(rec['latest_period'])}) — HLF Pulse Reference Data</title>
  <script type="application/ld+json">{json.dumps(schema_obj)}</script>
  <style>
    body {{ font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #111827; background: #f8fafc; margin: 0; padding: 26px 14px 48px; line-height: 1.6; font-variant-numeric: tabular-nums; }}
    main {{ max-width: 920px; margin: 0 auto; background: #ffffff; border-top: 5px solid #0b2559; border-radius: 16px; padding: 32px; box-shadow: 0 4px 20px rgba(11,37,89,0.05); }}
    .crumb {{ font-size: 0.82rem; margin-bottom: 14px; color: #6b7280; }}
    .crumb a {{ color: #186bd6; font-weight: 600; text-decoration: none; }}
    h1 {{ margin: 0 0 8px; color: #0b2559; font-size: 1.45rem; letter-spacing: -0.02em; }}
    h2 {{ margin: 26px 0 10px; color: #0b2559; font-size: 1.08rem; border-bottom: 1px solid #e5e7eb; padding-bottom: 6px; }}
    .summary-box {{ background: #eff6ff; border: 1px solid #dbeafe; border-left: 4px solid #186bd6; border-radius: 8px 14px 14px 8px; padding: 16px 18px; margin: 16px 0 22px; font-size: 0.98rem; color: #111827; }}
    .summary-box strong {{ color: #0b2559; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.86rem; margin-top: 8px; }}
    th {{ width: 220px; background: #f8fafc; color: #0b2559; text-align: left; padding: 9px 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }}
    td {{ padding: 9px 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; color: #374151; }}
    code {{ background: #eff6ff; padding: 2px 6px; border-radius: 4px; font-size: 0.8rem; color: #0b2559; }}
    a {{ color: #186bd6; font-weight: 600; }}
    ul {{ margin: 8px 0 0; padding-left: 20px; color: #374151; font-size: 0.9rem; }}
    li {{ margin-bottom: 5px; }}
    .chain {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px; margin-top: 10px; }}
    .chain-step {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 10px; padding: 11px 13px; font-size: 0.82rem; }}
    .chain-step span {{ display: block; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; color: #6b7280; margin-bottom: 3px; }}
    .hist-wrap {{ max-width: 520px; border: 1px solid #e5e7eb; border-radius: 10px; overflow: hidden; margin-top: 10px; }}
    .hist-wrap th {{ width: auto; border-bottom: 2px solid #dbeafe; }}
  </style>
</head>
<body>
  <main>
    <div class="crumb"><a href="{PULSE_BASE_URL}/">&larr; All HLF Pulse Reference Series</a> &bull; Canonical Statistic Record</div>
    <h1>{html_lib.escape(rec['name'])}</h1>
    <div style="font-size:0.82rem;color:#6b7280;">Ticker: <code>{key}</code> &bull; Topic: <strong>{html_lib.escape(rec['topic'])}</strong> &bull; Status: <strong>{rec['status']}</strong></div>

    <div class="summary-box">{rec['html']}</div>

    <h2>Structured Data Record</h2>
    <table>
      <tbody>
        <tr><th>Ticker ID</th><td><code>{key}</code></td></tr>
        <tr><th>Statistic Name</th><td><strong>{html_lib.escape(rec['name'])}</strong></td></tr>
        <tr><th>Topic</th><td>{html_lib.escape(rec['topic'])}</td></tr>
        <tr><th>Official Source</th><td><a href="{rec['source_url']}" target="_blank" rel="noopener">{html_lib.escape(rec['source'])}</a></td></tr>
        <tr><th>Official Series IDs</th><td><code>{html_lib.escape(rec['source_series_ids'])}</code></td></tr>
        <tr><th>Latest Data Period</th><td><strong>{html_lib.escape(rec['latest_period'])}</strong> (Observation end: <code>{rec['as_of']}</code>)</td></tr>
        <tr><th>Latest Value</th><td><strong>{html_lib.escape(str(rec['latest_value']))}</strong></td></tr>
        <tr><th>Previous Comparable Period</th><td>{html_lib.escape(str(rec['previous_period']))}</td></tr>
        <tr><th>Change</th><td>{html_lib.escape(str(rec['change']))}</td></tr>
        <tr><th>Data Type</th><td><strong>{html_lib.escape(rec['data_type'])}</strong></td></tr>
        <tr><th>Definition</th><td>{html_lib.escape(rec['definition'])}</td></tr>
        <tr><th>Official Methodology</th><td>{html_lib.escape(rec['methodology'])}</td></tr>
        <tr><th>HLF Calculation</th><td>{html_lib.escape(rec['hlf_calculation'])}</td></tr>
        <tr><th>Update Frequency</th><td>{html_lib.escape(rec['update_frequency'])}</td></tr>
        <tr><th>Last Source Refresh</th><td>{now_human} (<code>{rec['last_verified']}</code>)</td></tr>
        <tr><th>Canonical URL</th><td><a href="{rec['canonical_url']}">{rec['canonical_url']}</a></td></tr>
      </tbody>
    </table>

    {hist_section}
    {q_block}

    <h2>Citation Chain &amp; Related HLF Resources</h2>
    <div class="chain">
      <div class="chain-step"><span>1. Official Source</span><a href="{c_chain['official_source_url']}" target="_blank" rel="noopener">{html_lib.escape(c_chain['official_source'])}</a></div>
      <div class="chain-step"><span>2. Pulse Record</span><a href="{c_chain['pulse_record_url']}"><code>{key}</code></a></div>
      <div class="chain-step"><span>3. HLF Explanation</span><a href="{c_chain['hlf_article_url']}" target="_blank" rel="noopener">{html_lib.escape(c_chain['hlf_article_title'])}</a></div>
      <div class="chain-step"><span>4. Interactive Tool</span><a href="{c_chain['hlf_tool_url']}" target="_blank" rel="noopener">{html_lib.escape(c_chain['hlf_tool_title'])}</a></div>
      <div class="chain-step"><span>5. Personalised Advice</span><a href="{c_chain['hlf_advice_url']}" target="_blank" rel="noopener">Speak with an HLF Adviser</a></div>
    </div>

    <p style="margin-top:24px;font-size:0.78rem;color:#6b7280;">{rec['source_html']} Published by <a href="https://www.homeloanfactory.co.nz/">HLF (Home Loan Factory)</a> (<a href="https://www.homeloanfactory.co.nz/public-disclosures">Public Disclosures</a>).</p>
  </main>
</body>
</html>"""
    (t_dir / "index.html").write_text(page_html, encoding="utf-8")


# ---------------------------------------------------------------------------
# Main Orchestrator & Strict Per-Ticker Freshness Gate
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

    now_dt = datetime.datetime.now(datetime.timezone.utc)
    now_utc = now_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    today_iso = now_dt.strftime("%Y-%m-%d")
    now_human = now_dt.strftime("%-d %B %Y")

    verified_tickers = {}
    for k, item in fresh_candidates.items():
        if not isinstance(item, dict):
            continue
        as_of = item.get("as_of", "")
        max_age = int(item.get("max_age_days", MAX_AGE_MONTHLY))
        if item.get("html") and is_fresh(as_of, max_age):
            verified_tickers[k] = enrich_ticker_record(k, item, today_iso)
        else:
            print(f"[Failsafe Gate] Omitted stale/invalid ticker '{k}' (as_of={as_of}, max_age={max_age}d)")

    if live_parsed_count == 0 and not verified_tickers:
        raise RuntimeError("Zero live datasets parsed and zero valid fresh tickers remain; failing workflow honestly.")

    # Compare core html/as_of fields AND check if reference-layer files exist before skipping commit
    prev_core = {k: (v.get("as_of"), v.get("html")) for k, v in prev_tickers.items() if isinstance(v, dict)}
    curr_core = {k: (v.get("as_of"), v.get("html")) for k, v in verified_tickers.items()}
    canonical_ready = all((root / k / "index.html").exists() for k in verified_tickers)

    if curr_core == prev_core and html_path.exists() and structured_json_path.exists() and canonical_ready:
        print(f"All {len(verified_tickers)} verified tickers and canonical pages unchanged; skipping redundant commit.")
        return

    # 1. Write backward-compatible nz-mortgage-pulse.json (for Webflow Chat Bubbles)
    payload = {
        "publisher": "HLF — Home Loan Factory (homeloanfactory.co.nz)",
        "dataset_name": "NZ Mortgage, Housing & KiwiSaver Pulse — Verified NZ Housing & Mortgage Data",
        "generated_utc": now_utc,
        "last_source_refresh": today_iso,
        "status": "Verified",
        "count": len(verified_tickers),
        "tickers": verified_tickers,
    }
    json_path.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")

    # 2. Write structured machine-readable pulse.json (for AI retrieval, researchers & developers)
    pulse_records_Payload = {
        "dataset": "HLF Pulse — Verified New Zealand Mortgage, Housing & KiwiSaver Reference Dataset",
        "publisher": "HLF (Home Loan Factory)",
        "homepage": "https://www.homeloanfactory.co.nz/",
        "pulse_url": f"{PULSE_BASE_URL}/",
        "dataset_generated_utc": now_utc,
        "last_source_refresh": today_iso,
        "next_expected_update": "Monthly (RBNZ / Stats NZ / MBIE Tenancy) & Quarterly (CPI / IRD Student Loans)",
        "data_status": "Verified",
        "active_series_count": len(verified_tickers),
        "sources": [
            "Reserve Bank of New Zealand (Tables C31, C32, C33, C40, C41, B30)",
            "Stats NZ (Building Consents Issued & Consumers Price Index)",
            "Ministry of Business, Innovation and Employment (Tenancy Services Rental Bond Data)",
            "Inland Revenue (KiwiSaver Withdrawals, Scheme Transfers & Student Loan Balances)",
        ],
        "records": [verified_tickers[k] for k in sorted(verified_tickers.keys())],
    }
    structured_json_path.write_text(json.dumps(pulse_records_Payload, indent=2), encoding="utf-8")

    # 3. Write individual canonical reference pages (/<ticker>/index.html) + sitemap.xml + llms.txt
    for k, rec in verified_tickers.items():
        write_canonical_ticker_page(root, k, rec, now_human)

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
        f"> Public reference dataset maintained by HLF (Home Loan Factory, https://www.homeloanfactory.co.nz/). Last verified: {now_human} ({now_utc}).",
        "",
        "## Machine-Readable Endpoints",
        f"- Structured Reference Records (JSON): {PULSE_BASE_URL}/pulse.json",
        f"- Webflow Chat Bubble Feed (JSON): {PULSE_BASE_URL}/nz-mortgage-pulse.json",
        f"- Main Reference Dataset Hub (HTML): {PULSE_BASE_URL}/",
        "",
        "## Canonical Statistic Reference Pages",
    ] + [
        f"- [{rec['name']} ({rec['latest_period']})]({rec['canonical_url']}): {rec['summary_text']}"
        for k, rec in sorted(verified_tickers.items())
    ]
    (root / "llms.txt").write_text("\n".join(llms_lines) + "\n", encoding="utf-8")

    # 4. Write upgraded index.html Reference Dataset Hub
    robots_tag = "index, follow" if verified_tickers else "noindex, nofollow"
    latest_rbnz_period = verified_tickers.get("fhb-low-deposit", {}).get("latest_period", "Latest Monthly Release")

    rows_html_list = []
    for k, v in sorted(verified_tickers.items()):
        c_chain = v["citation_chain"]
        hist_tbl = render_history_table_html(v.get("history", []))
        q_list = "".join(f"<li>{html_lib.escape(q)}</li>" for q in v.get("related_questions", []))
        rows_html_list.append(
            f"<tr id='{k}'>"
            f"<td><a href='{v['canonical_url']}'><code>{k}</code></a><div style='margin-top:4px;font-size:0.75rem;color:#374151;font-weight:600;'>{html_lib.escape(v['name'])}</div>"
            f"<span class='dtype-badge'>{html_lib.escape(v['data_type'].split('(')[0].strip())}</span></td>"
            f"<td><strong>{html_lib.escape(v['chip'])}</strong><br><span style='font-size:0.76rem;color:#0b2559;font-weight:600;'>Period: {html_lib.escape(v['latest_period'])}</span>"
            f"<br><span style='font-size:0.72rem;color:#6b7280;'>Obs end: {v['as_of']}</span></td>"
            f"<td><div style='font-size:0.9rem;color:#111827;'>{v['html']}</div>"
            f"<div style='margin-top:5px;font-size:0.76rem;color:#6b7280;'>{v['source_html']} &bull; <a href='{v['canonical_url']}'>Canonical record &rarr;</a></div>"
            f"<details class='rec-drawer'><summary>View full data record, methodology, history &amp; citation chain</summary>"
            f"<div class='rec-drawer-body'>"
            f"<div><strong>Latest Value:</strong> {html_lib.escape(str(v['latest_value']))} &bull; <strong>Previous:</strong> {html_lib.escape(str(v['previous_period']))} &bull; <strong>Change:</strong> {html_lib.escape(str(v['change']))}</div>"
            f"<div style='margin-top:6px;'><strong>Data Type:</strong> {html_lib.escape(v['data_type'])} &bull; <strong>Series IDs:</strong> <code>{html_lib.escape(v['source_series_ids'])}</code></div>"
            f"<div style='margin-top:6px;'><strong>Definition:</strong> {html_lib.escape(v['definition'])}</div>"
            f"<div style='margin-top:6px;'><strong>Methodology &amp; HLF Calculation:</strong> {html_lib.escape(v['methodology'])} Formula: <code>{html_lib.escape(v['hlf_calculation'])}</code></div>"
            + (f"<div style='margin-top:8px;'><strong>Questions This Data Helps Answer:</strong><ul style='margin:4px 0 0;padding-left:18px;'>{q_list}</ul></div>" if q_list else "")
            + (f"<div style='margin-top:8px;'><strong>Historical / Regional Values:</strong>{hist_tbl}</div>" if hist_tbl else "")
            + f"<div style='margin-top:8px;'><strong>Citation Chain:</strong> <a href='{c_chain['official_source_url']}' target='_blank' rel='noopener'>1. Official Source ({html_lib.escape(v['chip'])})</a> &rarr; "
            f"<a href='{v['canonical_url']}'>2. Pulse Record (/{k})</a> &rarr; "
            f"<a href='{c_chain['hlf_article_url']}' target='_blank' rel='noopener'>3. {html_lib.escape(c_chain['hlf_article_title'])}</a> &rarr; "
            f"<a href='{c_chain['hlf_tool_url']}' target='_blank' rel='noopener'>4. {html_lib.escape(c_chain['hlf_tool_title'])}</a> &rarr; "
            f"<a href='{c_chain['hlf_advice_url']}' target='_blank' rel='noopener'>5. HLF Mortgage &amp; KiwiSaver Advice</a></div>"
            f"</div></details></td></tr>"
        )
    rows_html = "".join(rows_html_list)

    html_path.write_text(
        f"""<!DOCTYPE html>
<html lang="en-NZ">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <meta name="robots" content="{robots_tag}">
  <link rel="canonical" href="{PULSE_BASE_URL}/">
  <title>NZ Mortgage, Housing &amp; KiwiSaver Pulse — Verified NZ Housing &amp; Mortgage Data | HLF</title>
  <style>
    body {{ font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; color: #111827; background: #f8fafc; margin: 0; padding: 24px 14px 56px; line-height: 1.58; font-variant-numeric: tabular-nums; }}
    main {{ max-width: 1120px; margin: 0 auto; background: #ffffff; border-top: 5px solid #0b2559; border-radius: 16px; padding: 32px; box-shadow: 0 4px 20px rgba(11,37,89,0.05); }}
    h1 {{ margin: 0 0 8px; color: #0b2559; font-size: 1.55rem; letter-spacing: -0.02em; }}
    h2 {{ margin: 28px 0 10px; color: #0b2559; font-size: 1.18rem; letter-spacing: -0.01em; }}
    p {{ color: #374151; font-size: 0.92rem; margin: 0 0 12px; }}
    a {{ color: #186bd6; font-weight: 600; }}
    .status-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 10px; margin: 18px 0 24px; }}
    .status-tile {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 10px; padding: 10px 12px; }}
    .status-tile span {{ display: block; font-size: 0.7rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.04em; color: #6b7280; margin-bottom: 3px; }}
    .status-tile strong {{ font-size: 0.86rem; color: #0b2559; }}
    .about-box {{ background: #eff6ff; border: 1px solid #dbeafe; border-radius: 14px; padding: 20px 22px; margin-bottom: 24px; }}
    .about-box h2 {{ margin-top: 0; }}
    .dtype-list {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 10px; margin-top: 12px; }}
    .dtype-card {{ background: #ffffff; border: 1px solid #dbeafe; border-radius: 10px; padding: 10px 12px; font-size: 0.82rem; color: #374151; }}
    .dtype-card strong {{ display: block; color: #0b2559; margin-bottom: 2px; }}
    .dtype-badge {{ display: inline-block; margin-top: 6px; background: #eff6ff; border: 1px solid #bfdbfe; color: #186bd6; font-size: 0.69rem; font-weight: 700; padding: 2px 7px; border-radius: 999px; }}
    table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem; margin-top: 12px; }}
    th {{ background: #f8fafc; color: #0b2559; text-align: left; padding: 10px 12px; border-bottom: 2px solid #dbeafe; }}
    td {{ padding: 12px; border-bottom: 1px solid #e5e7eb; vertical-align: top; }}
    code {{ background: #eff6ff; padding: 2px 6px; border-radius: 4px; font-size: 0.78rem; color: #0b2559; }}
    .rec-drawer {{ margin-top: 8px; background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 8px; padding: 6px 10px; font-size: 0.79rem; }}
    .rec-drawer summary {{ cursor: pointer; font-weight: 600; color: #0b2559; user-select: none; }}
    .rec-drawer-body {{ margin-top: 8px; padding-top: 8px; border-top: 1px solid #e5e7eb; color: #374151; line-height: 1.52; }}
    .hist-wrap {{ max-width: 460px; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; margin-top: 6px; background: #ffffff; }}
    .hist-wrap table {{ margin-top: 0; font-size: 0.77rem; }}
    .hist-wrap th, .hist-wrap td {{ padding: 5px 9px; }}
    .defs-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 12px; margin-top: 12px; }}
    .def-card {{ background: #f8fafc; border: 1px solid #e5e7eb; border-radius: 12px; padding: 14px 16px; font-size: 0.84rem; color: #374151; }}
    .def-card strong {{ display: block; color: #0b2559; font-size: 0.9rem; margin-bottom: 4px; }}
  </style>
</head>
<body>
  <main>
    <h1>NZ Mortgage, Housing &amp; KiwiSaver Pulse — Verified NZ Housing &amp; Mortgage Data</h1>
    <p>Public reference dataset maintained by <a href="https://www.homeloanfactory.co.nz/">HLF (Home Loan Factory)</a>. Machine-readable endpoints: <a href="{PULSE_BASE_URL}/pulse.json"><code>pulse.json</code></a> (structured records), <a href="{PULSE_BASE_URL}/nz-mortgage-pulse.json"><code>nz-mortgage-pulse.json</code></a> (live ticker feed), and <a href="{PULSE_BASE_URL}/llms.txt"><code>llms.txt</code></a>.</p>

    <div class="status-grid" aria-label="Dataset Verification &amp; Update Status">
      <div class="status-tile"><span>Dataset Generated</span><strong>{now_human}</strong></div>
      <div class="status-tile"><span>Last Source Refresh</span><strong>{now_human} ({today_iso})</strong></div>
      <div class="status-tile"><span>Latest Data Period</span><strong>{html_lib.escape(latest_rbnz_period)} (Monthly)</strong></div>
      <div class="status-tile"><span>Next Expected Update</span><strong>Monthly / Quarterly</strong></div>
      <div class="status-tile"><span>Data Status</span><strong style="color:#065f46;">&#10003; Verified ({len(verified_tickers)} Active Series)</strong></div>
      <div class="status-tile"><span>Official Sources</span><strong>RBNZ, Stats NZ, MBIE, IRD</strong></div>
    </div>

    <section class="about-box" aria-label="About This Dataset">
      <h2>What is the HLF Pulse?</h2>
      <p>The <strong>HLF Pulse</strong> is a public reference dataset covering New Zealand mortgage, housing, and KiwiSaver statistics. It brings together selected statistics from official New Zealand Crown sources including the <strong>Reserve Bank of New Zealand (RBNZ)</strong>, <strong>Stats NZ</strong>, <strong>MBIE Tenancy Services</strong>, and <strong>Inland Revenue (IRD)</strong>.</p>
      <p>HLF does not create the underlying official statistics. We collect, standardise, and, where relevant, derive selected measures from those published datasets so they can be used consistently across HLF articles, interactive calculators, and educational resources.</p>
      <p style="margin-bottom:0;">Each statistic identifies its underlying official source, reporting period, methodology, historical values, and whether the figure is directly reported or derived by HLF. The dataset is designed to make New Zealand mortgage and housing data easier for people, researchers, publishers, and information systems to find, understand, and verify.</p>

      <div class="dtype-list">
        <div class="dtype-card"><strong>1. Officially reported</strong>Figures published directly by the Crown agency (e.g., loan counts, dollar lending totals, settled rates in RBNZ B30, building consents, median rents).</div>
        <div class="dtype-card"><strong>2. HLF-derived</strong>Proportions, shares, and per-borrower averages calculated by HLF from two official series (e.g., <code>LVRN.MMB2.BC1 &divide; LVRN.MMB2.AC1</code>).</div>
        <div class="dtype-card"><strong>3. HLF-calculated comparison</strong>Year-on-year percentage point changes, 12-month rolling totals, cohort gaps, and headroom against RBNZ speed limits.</div>
        <div class="dtype-card"><strong>4. HLF interpretation</strong>Plain-English educational context explaining how lending rules, bank test rates, or exemptions apply in practice.</div>
      </div>
    </section>

    <h2>Verified NZ Housing, Mortgage &amp; KiwiSaver Data Records</h2>
    <p>Click any Ticker Key to open its standalone canonical reference page (<code>/{'{ticker}'}/</code>), or expand the drawer beneath any statistic to inspect its structured record, formula, 13-period history, and citation chain.</p>
    <table>
      <thead><tr><th style="width:230px;">Ticker &amp; Data Type</th><th style="width:165px;">Source &amp; Period</th><th>Verified Statistic, Attribution &amp; Expandable Data Record</th></tr></thead>
      <tbody>{rows_html}</tbody>
    </table>

    <h2>Explicit Definitions &amp; Official Methodology Reference</h2>
    <div class="defs-grid">
      <div class="def-card"><strong>First-Home Buyer (FHB) Lending (RBNZ C31)</strong>Firm mortgage commitments issued by registered NZ banks to borrowers purchasing a residential property to occupy as their first home. Measures bank loan commitments rather than LINZ title settlements.</div>
      <div class="def-card"><strong>Low-Deposit / High-LVR Lending (&gt;80% LVR)</strong>Mortgages where the loan amount exceeds 80.0% of the property's value (meaning the buyer's deposit or equity is strictly under 20.0%). Loans with an exact 20.0% deposit (80.0% LVR) are classified by RBNZ inside <code>&le;80% LVR</code>.</div>
      <div class="def-card"><strong>Property Investor Lending &amp; LVR Tiers (RBNZ C31)</strong>Lending secured against residential investment property. RBNZ tracks investor LVR tiers at <code>&le;60% LVR</code> (40%+ equity), <code>60%&ndash;70% LVR</code> (30%&ndash;40% equity), <code>70%&ndash;80% LVR</code> (20%&ndash;30% equity, including 20%-deposit new builds), and <code>&gt;80% LVR</code> (&lt;20% equity).</div>
      <div class="def-card"><strong>Debt-to-Income (TDTI) Ratio (RBNZ C40)</strong>Total borrower debt (including existing mortgages, student loans, and credit limits) divided by annual gross before-tax borrower income. RBNZ's DTI speed limit allows banks up to 20% of new FHB lending above 6x DTI (new builds are exempt).</div>
      <div class="def-card"><strong>Borrower Gross Income — BGI (RBNZ C41)</strong>Total pre-tax annual gross income reported on approved mortgage applications (combining single and joint earners, plus eligible boarder/rental income). HLF computes the exact borrower-weighted harmonic mean across all TDTI bands.</div>
      <div class="def-card"><strong>Settled Mortgage Rates vs. Advertised Rates (RBNZ B30)</strong>RBNZ Table B30 records the dollar-weighted average interest rate across mortgages actually settled during the month, capturing negotiated and unadvertised bank rate discounts rather than carded rack rates.</div>
    </div>
  </main>
</body>
</html>""",
        encoding="utf-8",
    )
    print(f"Saved {len(verified_tickers)} verified tickers, pulse.json, index.html, llms.txt, sitemap.xml & canonical /<ticker>/ pages.")


if __name__ == "__main__":
    main()
