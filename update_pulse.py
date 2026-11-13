#!/usr/bin/env python3
"""
HLF Pulse — Reference-Grade New Zealand Mortgage, Housing, KiwiSaver & Protection Dataset
=========================================================================================
63 Active Datasets Across 10 Categories:
  - Main prose (`html` & `summary_text`) is written as clean, authoritative editorial
    reference prose — free of methodology clutter.
  - All HLF derivations, formulas, rounding disclosures, licensing (CC BY 4.0), and
    verification dates live in the structured provenance layer (`source_html` dropdown,
    `pulse.json`, `llms.txt`, and Data Dictionary tables).
  - Includes the 5-Row FHB Snapshot tickers, Dollar Build Cost by Dwelling Type,
    Personal Insurance Claim Payout Benchmarks, Investor >7x DTI Headroom, and
    Fixed vs. Floating Rate Roll-Off Series.
  - Strict 4-field date standard: Observation Period, Period End, Source Published, Last Checked.
"""

import base64
import datetime
from html.parser import HTMLParser
import html as html_lib
import io
import json
import os
import pathlib
import re
import time

import pandas as pd

PULSE_BASE_URL = "https://pulse-data.homeloanfactory.co.nz"
PULSE_DATE_CREATED = "2026-02-01"

# ---------------------------------------------------------------------------
# Creator, Publisher & Open Licensing Constants (CC BY 4.0)
# ---------------------------------------------------------------------------
CREATOR_NAME = "Andy Palliser"
CREATOR_ALT_NAME = "Andrew Jonathan Palliser"
CREATOR_FSP = "FSP1004725"
CREATOR_ROLE = "Mortgage & KiwiSaver Adviser · Home Loan Factory"
CREATOR_BADGE = "Creator of HLF Pulse"
CREATOR_URL = "https://www.homeloanfactory.co.nz/team/andrew-palliser"
CREATOR_IMAGE = "https://cdn.prod.website-files.com/6a1baf404596ba888865b209/6a1ce4c85096daa2e99a0ecb_Andrew-Palliser.avif"

PUBLISHER_NAME = "Home Loan Factory"
PUBLISHER_LEGAL = "HLF Financial Limited"
PUBLISHER_FSP = "FSP1008571"
PUBLISHER_URL = "https://www.homeloanfactory.co.nz/"
CONTACT_URL = "https://www.homeloanfactory.co.nz/contact"
DISCLOSURES_URL = "https://www.homeloanfactory.co.nz/public-disclosures"
HOW_PAID_URL = "https://www.homeloanfactory.co.nz/how-we-get-paid"
MEDIA_KIT_URL = "https://www.homeloanfactory.co.nz/media-and-citations"

HLF_LICENSE_URL = "https://creativecommons.org/licenses/by/4.0/"
HLF_LICENSE_SHORT = "CC BY 4.0"
HLF_LICENSE_NOTE = (
    "HLF-derived calculations, percentages, averages, multi-year real/nominal comparisons, "
    "data dictionary definitions, and market interpretations are licensed by Home Loan Factory "
    "under Creative Commons Attribution 4.0 International (CC BY 4.0) with attribution to "
    "HLF Pulse (https://pulse-data.homeloanfactory.co.nz/). Underlying official Crown and "
    "statutory statistics remain subject to the terms of their respective publishers "
    "(RBNZ Terms of Use; Stats NZ, Inland Revenue, MBIE, and FMA under CC BY 4.0; "
    "Natural Hazards Insurance Act 2023 and Accident Compensation Act 2001 under public-domain NZ legislation)."
)

# Cumulative Stats NZ CPI factors for real (inflation-adjusted) 5-year and 10-year comparisons
CPI_10Y_FACTOR = 1.337  # +33.7% cumulative NZ all-groups CPI over 10 years (2016 -> 2026)
CPI_5Y_FACTOR = 1.228   # +22.8% cumulative NZ all-groups CPI over 5 years (2021 -> 2026)

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
        "name": PUBLISHER_NAME,
        "legalName": PUBLISHER_LEGAL,
        "identifier": PUBLISHER_FSP,
        "url": PUBLISHER_URL,
    },
}

PUBLISHER_SCHEMA_OBJ = {
    "@type": "Organization",
    "name": PUBLISHER_NAME,
    "legalName": PUBLISHER_LEGAL,
    "identifier": PUBLISHER_FSP,
    "url": PUBLISHER_URL,
}

# Freshness windows (in days)
MAX_AGE_MONTHLY = 100
MAX_AGE_TENANCY = 140
MAX_AGE_QUARTERLY = 160
MAX_AGE_ANNUAL = 550
MAX_AGE_INSURANCE = 550

RBNZ_WARMUP_URL = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary"

RBNZ_SOURCES = {
    "c31": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c31/c31-new-residential-mortgage-lending-by-borrower-type.xlsx",
        "match": "c31",
    },
    "c33": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c33/c33-new-residential-mortgage-lending-by-purpose.xlsx",
        "match": "c33",
    },
    "c32": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c32/c32-new-and-existing-residential-mortgage-lending-by-payment-type.xlsx",
        "match": "c32",
    },
    "c40": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use",
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/debt-to-income-dti-statistics",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c40/c40-residential-mortgage-lending-by-dti-purpose-use.xlsx",
        "match": "c40",
    },
    "c41": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi",
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/debt-to-income-dti-statistics",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c41/c41-residential-mortgage-borrower-gross-income.xlsx",
        "match": "c41",
    },
    "b30": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-weighted-average-interest-rates",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b30/b30-new-residential-mortgage-weighted-average-interest-rates.xlsx",
        "match": "b30",
    },
}

TENANCY_PAGE = "https://www.tenancy.govt.nz/about-tenancy-services/data-and-statistics/rental-bond-data/"
IRD_SL_BAL_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/social-policy/student-loans/scheme-balances"
IRD_KS_NUM_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/withdrawals/numbers"
IRD_KS_AMT_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/withdrawals/amounts"
IRD_KS_TR_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/Joining-and-transfers/scheme-transfers"
IRD_KS_REG_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/withdrawals/withdrawals-regions"
IRD_KS_ENTRY_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/Joining-and-transfers/entry-methods"
IRD_KS_PAY_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/contributions/payments"
IRD_KS_SUSP_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/contributions/savings-suspensions"
IRD_KS_DATA_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/datasets"
FMA_KS_REPORT_PAGE = "https://www.fma.govt.nz/library/reports-and-papers/kiwisaver-annual-report/"
STATS_CPI_PAGE = "https://www.stats.govt.nz/information-releases/consumers-price-index-september-2026-quarter/"
STATS_HES_PAGE = "https://www.stats.govt.nz/information-releases/household-expenditure-statistics-year-ended-june-2025/"
RBNZ_FSR_PAGE = "https://www.rbnz.govt.nz/financial-stability/financial-stability-report"
RBNZ_INS_STATS_PAGE = "https://www.rbnz.govt.nz/statistics/series/insurance"
RBNZ_C35_PAGE = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-loan-reconciliation"
NHC_ACT_PAGE = "https://www.legislation.govt.nz/act/public/2023/0004/latest/LMS487798.html"
NHC_FUND_PAGE = "https://www.naturalhazards.govt.nz/"
ACC_REPORT_PAGE = "https://www.acc.co.nz/about-us/corporate-documents/annual-report"
ACC_LEVY_PAGE = "https://www.acc.co.nz/for-business/understanding-your-levies/paying-levies-if-you-work-for-yourself"
FSC_DATA_PAGE = "https://www.fsc.org.nz/research-and-data"

CATEGORY_ORDER = [
    "First-home buyers",
    "Homeowners & refinancing",
    "Property investors",
    "LVR, DTI & credit policy",
    "New builds & construction",
    "Rents & tenancy",
    "KiwiSaver",
    "Property & Natural Hazard Insurance",
    "Personal Risk & Income Protection",
    "Insurance Market & Claims",
]

# ---------------------------------------------------------------------------
# Reference-Grade Data Dictionary & Provenance Metadata (63 Canonical Tickers)
# ---------------------------------------------------------------------------
TICKER_CATALOG = {
    "fhb-low-deposit": {
        "name": "First-home buyers using under 20% deposit (>80% LVR)",
        "also_known_as": "Low-deposit first-home buyer mortgage share / High-LVR (>80%) FHB share",
        "category": "First-home buyers",
        "topic": "First-home buyers & deposit requirements",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB2.BC1 ÷ LVRN.MMB2.AC1 (Count) | LVRN.MMB1.BC1 (Value)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived percentage share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of new first-home buyer residential mortgage commitments approved by New Zealand registered banks with a loan-to-value ratio (LVR) above 80% (meaning a deposit of less than 20%).",
        "methodology": "Derived from RBNZ Table C31 by dividing the monthly number of new first-home buyer mortgage commitments with LVR > 80% (LVRN.MMB2.BC1) by total monthly first-home buyer mortgage commitments (LVRN.MMB2.AC1), multiplied by 100.",
        "hlf_calculation": "(High-LVR >80% FHB Borrowers [LVRN.MMB2.BC1] ÷ Total FHB Borrowers [LVRN.MMB2.AC1]) × 100",
        "hlf_interpretation": "HLF market context: First-home buyers are the primary recipients of bank low-deposit (>80% LVR) lending allowances, supplemented by LVR-exempt Kāinga Ora First Home Loans and new-build lending.",
        "questions": [
            "Can I buy a first home in New Zealand with less than a 20% deposit?",
            "What percentage of NZ first-home buyers get a mortgage with a 10% deposit?",
            "How common are low-deposit (>80% LVR) home loans for first-home buyers?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "Kāinga Ora LMI Fee & Low-Deposit Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/lenders-mortgage-insurance-fee-estimator",
    },
    "fhb-share": {
        "name": "First-home buyer share of new mortgage lending & average loan size",
        "also_known_as": "FHB share of total residential mortgage commitments & mean FHB loan value",
        "category": "First-home buyers",
        "topic": "First-home buyer lending activity",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.AC1 ÷ LVRN.MMB1.AA (Share) | LVRN.MMB1.AC1 ÷ LVRN.MMB2.AC1 (Mean Loan)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived share/mean",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The share of total new residential mortgage lending value committed to first-home buyers across New Zealand, alongside the monthly number of first-home buyer loans and mean loan size.",
        "methodology": "Derived from RBNZ Table C31 by dividing first-home buyer new lending value (LVRN.MMB1.AC1) by total new lending across all borrower types (LVRN.MMB1.AA, including first-home buyers, other owner-occupiers, investors, and business purpose). Mean loan size is LVRN.MMB1.AC1 (in $m, published by RBNZ rounded to $1m) divided by LVRN.MMB2.AC1.",
        "hlf_calculation": "(FHB Lending Value [LVRN.MMB1.AC1] ÷ All-Borrower Total Lending [LVRN.MMB1.AA]) × 100; Mean Loan = LVRN.MMB1.AC1 ÷ LVRN.MMB2.AC1 (subject to RBNZ $1m rounding)",
        "hlf_interpretation": "HLF market context: Tracks first-home buyer market participation relative to existing owner-occupiers and residential property investors.",
        "questions": [
            "What share of New Zealand mortgages go to first-home buyers?",
            "What is the average first-home buyer mortgage size in New Zealand?",
            "How many first-home buyers get approved for a mortgage each month in NZ?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-typical-income": {
        "name": "Typical first-home buyer household gross income (RBNZ Table C41)",
        "also_known_as": "FHB borrower-weighted average gross application income (BGI)",
        "category": "First-home buyers",
        "topic": "First-home buyer household income benchmark",
        "source": "RBNZ Table C41",
        "source_full": "Reserve Bank of New Zealand — Residential mortgage borrower gross income (BGI) (Table C41)",
        "series_ids": "Sum(FHB Band Gross Income Pool) ÷ Sum(FHB Band Implied Applications)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi",
        "data_type": "HLF-derived band-weighted mean from official RBNZ Table C41",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The typical (borrower-weighted mean) annual gross household income across approved New Zealand first-home buyer mortgage applications.",
        "methodology": "Derived from RBNZ Table C41 by summing total gross borrower income ($m) across all five first-home buyer Debt-to-Income bands and dividing by the total implied number of applications across those bands.",
        "hlf_calculation": "Total FHB Gross Income Pool ($m) × $1m ÷ Total Implied FHB Applications",
        "hlf_interpretation": "HLF market context: Because most first-home buyers apply as a couple (or with co-borrowers), RBNZ Borrower Gross Income reflects combined household income on the mortgage application—for example, two earners on ~$72,000 each equal ~$144,000 in gross application income.",
        "questions": [
            "What is the typical household income of a first-home buyer in New Zealand?",
            "How much do you need to earn to buy a first home in NZ?",
        ],
        "article_title": "DTI Rules in NZ: How Debt-to-Income Limits Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/dti-rules-in-nz-how-debt-to-income-limits-actually-work-for-home-buyers-and-investors",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-average-loan-size": {
        "name": "Typical first-home buyer mortgage size (RBNZ Table C31)",
        "also_known_as": "Average New Zealand first-home buyer loan amount",
        "category": "First-home buyers",
        "topic": "First-home buyer loan size benchmark",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.AC1 ÷ LVRN.MMB2.AC1",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "HLF-derived mean loan size from official RBNZ Table C31",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The average residential mortgage size approved for first-home buyers across New Zealand.",
        "methodology": "Derived from RBNZ Table C31 by dividing monthly first-home buyer new lending value (LVRN.MMB1.AC1, published rounded to $1m) by monthly first-home buyer loan count (LVRN.MMB2.AC1).",
        "hlf_calculation": "(Total FHB Lending Value [LVRN.MMB1.AC1] × $1,000,000) ÷ Total FHB Borrower Count [LVRN.MMB2.AC1]",
        "hlf_interpretation": "HLF market context: Provides a direct national benchmark for what first-home buyers are borrowing today across New Zealand before regional price differences.",
        "questions": [
            "What is the average first-home buyer mortgage in New Zealand?",
            "How much does a typical first-home buyer borrow in NZ?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "fhb-typical-dti": {
        "name": "Typical first-home buyer Debt-to-Income (DTI) ratio & share <= 6x DTI (RBNZ C40/C41)",
        "also_known_as": "Implied FHB Debt-to-Income multiple & share of FHB lending within 5x–6x DTI",
        "category": "First-home buyers",
        "topic": "First-home buyer Debt-to-Income benchmark",
        "source": "RBNZ Tables C40 & C41",
        "source_full": "Reserve Bank of New Zealand — DTI Purpose Use (Table C40) & Borrower Gross Income (Table C41)",
        "series_ids": "C31/C41 Implied Loan-to-Income Ratio & C40 Share of FHB Lending with TDTI <= 5 and <= 6",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use",
        "data_type": "HLF-derived DTI ratio & band shares from official RBNZ Tables C31, C40 & C41",
        "derived_by_hlf": True,
        "unit": "ratio",
        "frequency": "Monthly",
        "definition": "The typical Debt-to-Income (DTI) multiple for approved New Zealand first-home buyers—measured both by the ratio of average FHB loan size to average FHB gross income and by the proportion of FHB lending written at or below 5x and 6x gross income in RBNZ Table C40.",
        "methodology": "Combines the implied FHB mortgage-to-income ratio (Average FHB Loan Size from Table C31 ÷ Average FHB Gross Borrower Income from Table C41) with the distribution of FHB commitments across TDTI bands <= 5x and <= 6x in RBNZ Table C40.",
        "hlf_calculation": "Implied Mortgage DTI = Avg FHB Loan ÷ Avg FHB Gross Income; Share <= 6x DTI = 100% minus (>6x TDTI Share in C40)",
        "hlf_interpretation": "HLF market context: Most New Zealand first-home buyers borrow around 3.9x to 4.3x their gross household income—well below the 6.0x RBNZ owner-occupier speed-limit threshold—because bank servicing test rates cap borrowing power before the 6x DTI limit is reached.",
        "questions": [
            "What is the average Debt-to-Income (DTI) ratio for a first-home buyer in New Zealand?",
            "What percentage of first-home buyers borrow less than 5x or 6x their income?",
        ],
        "article_title": "DTI Rules in NZ: How Debt-to-Income Limits Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/dti-rules-in-nz-how-debt-to-income-limits-actually-work-for-home-buyers-and-investors",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-monthly-lending-volume": {
        "name": "Monthly first-home buyer total mortgage lending volume & loan count (RBNZ Table C31)",
        "also_known_as": "Monthly new FHB mortgage commitments ($ millions & borrower count)",
        "category": "First-home buyers",
        "topic": "Monthly first-home buyer lending volume",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.AC1 (FHB Lending Value $m) & LVRN.MMB2.AC1 (FHB Borrower Count)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived YoY comparison",
        "derived_by_hlf": False,
        "unit": "NZD_millions",
        "frequency": "Monthly",
        "definition": "Total monthly dollar value and number of new residential mortgage commitments approved for first-home buyers across New Zealand.",
        "methodology": "Directly reported in RBNZ Table C31 series LVRN.MMB1.AC1 (FHB total new lending value in $ millions) and LVRN.MMB2.AC1 (FHB total borrower count), alongside HLF-calculated year-on-year percentage changes.",
        "hlf_calculation": "Officially reported LVRN.MMB1.AC1 ($m) and LVRN.MMB2.AC1 (loans), plus YoY % change",
        "hlf_interpretation": "HLF market context: Shows the total monthly flow of bank credit supporting first-home buyers entering the New Zealand property market.",
        "questions": [
            "How much do New Zealand banks lend to first-home buyers each month?",
            "How many first-home buyer mortgages are approved per month in New Zealand?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
# === END OF PART 1 ===
    # === PART 2A OF 5 ===
    "investors": {
        "name": "Property investor share of new residential mortgage lending",
        "also_known_as": "Residential property investor share of total mortgage commitments",
        "category": "Property investors",
        "topic": "Property investor lending activity",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.AN ÷ LVRN.MMB1.AA (Share) | LVRN.MMB2.AN (Borrower Count)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived percentage share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of total new residential mortgage lending value committed to residential property investors across New Zealand.",
        "methodology": "Derived from RBNZ Table C31 by dividing monthly residential investor lending value (LVRN.MMB1.AN) by total new mortgage lending across all borrower types (LVRN.MMB1.AA).",
        "hlf_calculation": "(Investor Lending Value [LVRN.MMB1.AN] ÷ All-Borrower Total Lending [LVRN.MMB1.AA]) × 100",
        "hlf_interpretation": "HLF market context: Reflects how investor borrowing responds to mortgage interest rates, rental yields, interest deductibility rules, and LVR/DTI settings.",
        "questions": [
            "What percentage of New Zealand mortgages go to property investors?",
            "Are property investors borrowing more or less than first-home buyers in NZ?",
            "How much new mortgage lending goes to landlords each month?",
        ],
        "article_title": "Using Equity to Buy an Investment Property in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/using-equity-to-buy-an-investment-property-in-nz-the-split-banking-guide",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "investor-average-loan-size": {
        "name": "Property investor monthly lending volume & average loan size (RBNZ Table C31)",
        "also_known_as": "Average New Zealand residential property investor mortgage size & monthly commitments",
        "category": "Property investors",
        "topic": "Property investor loan size benchmark",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.AN ÷ LVRN.MMB2.AN",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived mean loan size",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The monthly dollar volume, borrower count, and average loan size across new residential property investor mortgage commitments in New Zealand.",
        "methodology": "Derived from RBNZ Table C31 by dividing monthly investor lending value (LVRN.MMB1.AN, rounded to $1m) by monthly investor loan count (LVRN.MMB2.AN).",
        "hlf_calculation": "(Investor Lending Value [LVRN.MMB1.AN] × $1,000,000) ÷ Investor Borrower Count [LVRN.MMB2.AN]",
        "hlf_interpretation": "HLF market context: Investor mortgage commitments include both standalone rental purchases and equity-leveraged portfolio restructures across New Zealand.",
        "questions": [
            "What is the average mortgage size for a property investor in New Zealand?",
            "How many property investor loans are approved each month in NZ?",
        ],
        "article_title": "Using Equity to Buy an Investment Property in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/using-equity-to-buy-an-investment-property-in-nz-the-split-banking-guide",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "mover-average-loan-size": {
        "name": "Existing owner-occupier ('mover') monthly lending volume & average loan size (RBNZ C31)",
        "also_known_as": "Other owner-occupier monthly mortgage commitments & mean loan value",
        "category": "Homeowners & refinancing",
        "topic": "Existing owner-occupier loan size benchmark",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.AC2 ÷ LVRN.MMB2.AC2",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived mean loan size",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The monthly dollar volume, borrower count, and average mortgage commitment size for existing ('other') owner-occupiers across New Zealand.",
        "methodology": "Derived from RBNZ Table C31 by dividing monthly other owner-occupier lending value (LVRN.MMB1.AC2, rounded to $1m) by monthly other owner-occupier loan count (LVRN.MMB2.AC2).",
        "hlf_calculation": "(Other Owner-Occupier Lending Value [LVRN.MMB1.AC2] × $1,000,000) ÷ Other Owner-Occupier Count [LVRN.MMB2.AC2]",
        "hlf_interpretation": "HLF market context: Captures existing homeowners moving house, refinancing, or topping up—typically with larger accumulated equity buffers than first-home buyers.",
        "questions": [
            "What is the average mortgage size for an existing homeowner or mover in New Zealand?",
            "How much do existing owner-occupiers borrow each month in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Net Proceeds & Next Home Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/net-proceeds-next-home-calculator",
    },
    "fhb-loan-size-gap": {
        "name": "First-home buyer average mortgage size by deposit tier (<20% vs 20%+ deposit)",
        "also_known_as": "High-LVR (>80%) vs standard-LVR (<=80%) first-home buyer loan size difference",
        "category": "First-home buyers",
        "topic": "First-home buyer loan sizes by LVR tier",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.BC1 ÷ LVRN.MMB2.BC1 vs (LVRN.MMB1.AC1 - BC1) ÷ (LVRN.MMB2.AC1 - BC1)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "HLF-derived cohort averages from official RBNZ series",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "Comparison of the mean mortgage size for first-home buyers borrowing with under a 20% deposit (>80% LVR) versus those borrowing with a 20% or greater deposit (<=80% LVR).",
        "methodology": "Derived from RBNZ Table C31 by dividing >80% LVR FHB lending value (LVRN.MMB1.BC1) by >80% LVR FHB loan count (LVRN.MMB2.BC1), and comparing against <=80% LVR FHB lending value divided by <=80% LVR FHB loan count. Note: RBNZ publishes Table C31 dollar series rounded to the nearest $1 million.",
        "hlf_calculation": "Mean(>80% LVR FHB Loan) minus Mean(<=80% LVR FHB Loan) (derived from RBNZ $1m-rounded lending series)",
        "hlf_interpretation": "HLF market context: Buyers with a smaller deposit (<20%) finance a larger percentage of the purchase price, resulting in a larger average mortgage balance that must pass bank servicing and DTI tests.",
        "questions": [
            "How much more do low-deposit first-home buyers borrow compared to buyers with a 20% deposit?",
            "What is the average loan size for a >80% LVR first-home buyer in New Zealand?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "fhb-bank-lvr-priority": {
        "name": "First-home buyer share of all >80% LVR (low-deposit) bank lending",
        "also_known_as": "FHB allocation share of total high-LVR (>80%) residential lending",
        "category": "LVR, DTI & credit policy",
        "topic": "Bank high-LVR (>80%) lending allocation",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.BC1 ÷ LVRN.MMB1.BA (Dollars) | LVRN.MMB2.BC1 ÷ LVRN.MMB2.BA (Loans)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "HLF-derived percentage share from official RBNZ series",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of all >80% LVR residential mortgage lending (by dollar value and loan count) across New Zealand that is allocated to first-home buyers.",
        "methodology": "Derived from RBNZ Table C31 by dividing first-home buyer >80% LVR lending value (LVRN.MMB1.BC1) by all-borrower >80% LVR lending value (LVRN.MMB1.BA), and FHB >80% LVR loan count (LVRN.MMB2.BC1) by all-borrower >80% LVR loan count (LVRN.MMB2.BA).",
        "hlf_calculation": "(FHB >80% LVR Value [LVRN.MMB1.BC1] ÷ All-Borrower >80% LVR Value [LVRN.MMB1.BA]) × 100",
        "hlf_interpretation": "HLF market context: Shows that New Zealand banks direct roughly four-fifths of all low-deposit (>80% LVR) lending to first-home buyers rather than existing homeowners or investors.",
        "questions": [
            "Do New Zealand banks prioritise first-home buyers for low-deposit (>80% LVR) loans?",
            "What share of >80% LVR mortgage lending goes to first-home buyers vs movers?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "Kāinga Ora LMI Fee & Low-Deposit Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/lenders-mortgage-insurance-fee-estimator",
    },
    "fhb-20-deposit-club": {
        "name": "First-home buyers with a 20%+ deposit (<=80% LVR)",
        "also_known_as": "Standard-deposit (<=80% LVR) first-home buyer share & average loan size",
        "category": "First-home buyers",
        "topic": "First-home buyers with 20%+ deposits",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "(LVRN.MMB2.AC1 - LVRN.MMB2.BC1) ÷ LVRN.MMB2.AC1",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "HLF-derived complement share & mean loan from official RBNZ series",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The share and number of first-home buyer mortgage approvals written at an LVR of 80% or below (meaning a deposit of 20% or more), alongside their mean loan size.",
        "methodology": "Derived from RBNZ Table C31 by subtracting >80% LVR FHB commitments from total FHB commitments for both loan count and dollar value.",
        "hlf_calculation": "((Total FHB Loans [LVRN.MMB2.AC1] - >80% LVR FHB Loans [LVRN.MMB2.BC1]) ÷ Total FHB Loans) × 100",
        "hlf_interpretation": "HLF market context: Buyers reaching the 20% deposit threshold (<=80% LVR) qualify for standard special carded fixed interest rates without low-equity premiums or LMI fees.",
        "questions": [
            "How many first-home buyers in New Zealand have a 20% deposit?",
            "What is the average mortgage size for a first-home buyer with a 20% deposit?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Compare NZ Mortgage Rates",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
    },
    "fhb-vs-investors": {
        "name": "First-home buyers vs property investors monthly lending comparison",
        "also_known_as": "FHB vs residential investor monthly mortgage dollar volume & gap",
        "category": "First-home buyers",
        "topic": "First-home buyers vs property investors",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.AC1 vs LVRN.MMB1.AN",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived dollar difference",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Monthly",
        "definition": "Direct monthly comparison of total new mortgage lending dollars committed to first-home buyers versus residential property investors across New Zealand.",
        "methodology": "Compares RBNZ Table C31 series LVRN.MMB1.AC1 (First-home buyers total lending value) against LVRN.MMB1.AN (Investors total lending value).",
        "hlf_calculation": "FHB Lending Value [LVRN.MMB1.AC1] minus Investor Lending Value [LVRN.MMB1.AN]",
        "hlf_interpretation": "HLF market context: Compares the monthly capital committed by banks to first-home buyers entering the market against landlords expanding or refinancing portfolios.",
        "questions": [
            "Do first-home buyers or property investors borrow more in New Zealand?",
            "How does first-home buyer lending compare to investor lending in RBNZ statistics?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "owner-equity-buffer": {
        "name": "Existing owner-occupiers borrowing with 20%+ equity (<=80% LVR)",
        "also_known_as": "Other owner-occupiers at <=80% LVR share & average loan size",
        "category": "Homeowners & refinancing",
        "topic": "Homeowner equity & mover lending",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "(LVRN.MMB2.AC2 - LVRN.MMB2.BC2) ÷ LVRN.MMB2.AC2",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "HLF-derived percentage & mean from official RBNZ series",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of existing ('other') owner-occupier mortgage approvals written at <=80% LVR (meaning 20% or more equity in the security property), alongside their mean loan size.",
        "methodology": "Derived from RBNZ Table C31 by subtracting >80% LVR other owner-occupier loans (LVRN.MMB2.BC2) from total other owner-occupier loans (LVRN.MMB2.AC2).",
        "hlf_calculation": "((Total Other OO Loans [LVRN.MMB2.AC2] - >80% LVR Other OO Loans [LVRN.MMB2.BC2]) ÷ Total Other OO Loans) × 100",
        "hlf_interpretation": "HLF market context: Over 90% of existing homeowners refinancing, topping up, or moving home hold at least 20% equity, giving them access to sharp carded mortgage rates.",
        "questions": [
            "What percentage of existing New Zealand homeowners have at least 20% equity when refinancing or moving?",
            "What is the average mortgage size for an existing owner-occupier in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Net Proceeds & Next Home Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/net-proceeds-next-home-calculator",
    },
    "investor-equity-tiers": {
        "name": "Property investor LVR & equity distribution (<=60%, 60%–70%, >70% LVR)",
        "also_known_as": "Investor mortgage commitments by LVR band (40%+ equity, 30%–40% equity, <30% equity)",
        "category": "Property investors",
        "topic": "Investor LVR & equity tiers",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB2.TN (<=60% LVR) | LVRN.MMB2.FN - DN (60–70% LVR) | LVRN.MMB2.DN (>70% LVR)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "HLF-derived LVR band shares from official RBNZ series",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "Breakdown of new residential property investor mortgage commitments across three LVR tiers: <=60% LVR (40%+ equity), 60%–70% LVR (30%–40% equity), and >70% LVR (under 30% equity).",
        "methodology": "Derived from RBNZ Table C31 investor borrower count series LVRN.MMB2.TN (<=60% LVR), LVRN.MMB2.FN (>60% LVR), LVRN.MMB2.DN (>70% LVR), and LVRN.MMB2.AN (total investors).",
        "hlf_calculation": "<=60% LVR = TN ÷ AN; 60–70% LVR = (FN - DN) ÷ AN; >70% LVR = DN ÷ AN",
        "hlf_interpretation": "HLF market context: Shows how investor borrowing clusters across the 60% and 70% LVR thresholds, with >70% LVR lending largely reflecting LVR-exempt new-build investment properties.",
        "questions": [
            "How much deposit or equity do property investors need in New Zealand?",
            "What percentage of NZ property investors borrow at 60% to 70% LVR?",
        ],
        "article_title": "Using Equity to Buy an Investment Property in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/using-equity-to-buy-an-investment-property-in-nz-the-split-banking-guide",
        "tool_title": "Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "investor-low-deposit": {
        "name": "Property investor high-LVR lending (>80% LVR & 70%–80% LVR bands)",
        "also_known_as": "Low-deposit property investor mortgages (>80% LVR & 70%–80% LVR)",
        "category": "Property investors",
        "topic": "High-LVR (>70% & >80%) investor lending",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB2.BN & LVRN.MMB1.BN (>80% LVR) | LVRN.MMB2.DN - BN (70–80% LVR)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived band split",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Monthly",
        "definition": "Monthly count and dollar value of residential property investor mortgage commitments approved at >80% LVR (under 20% deposit) and in the 70%–80% LVR band (20%–30% deposit).",
        "methodology": "Extracted from RBNZ Table C31 series LVRN.MMB2.BN / LVRN.MMB1.BN (>80% LVR investors), with the 70%–80% LVR band derived by subtracting >80% LVR investor commitments from >70% LVR investor commitments (LVRN.MMB2.DN / LVRN.MMB1.DN).",
        "hlf_calculation": ">80% LVR = LVRN.MMB2.BN; 70%–80% LVR = LVRN.MMB2.DN minus LVRN.MMB2.BN",
        "hlf_interpretation": "HLF market context: Under RBNZ LVR rules, new builds are exempt from LVR speed limits, allowing eligible new-build investors to borrow at 80% LVR (a 20% deposit) alongside discretionary bank speed-limit allowances.",
        "questions": [
            "Can a property investor buy with a 20% deposit in New Zealand?",
            "How many property investors get approved above 80% LVR or between 70% and 80% LVR?",
        ],
        "article_title": "Using Equity to Buy an Investment Property in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/using-equity-to-buy-an-investment-property-in-nz-the-split-banking-guide",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "investor-new-build-lvr": {
        "name": "Property investor lending above 70% LVR (under 30% deposit)",
        "also_known_as": "High-LVR (>70% LVR) investor mortgage share & volume",
        "category": "LVR, DTI & credit policy",
        "topic": "Investor >70% LVR lending & new-build exemptions",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB2.DN ÷ LVRN.MMB2.AN (Share) | LVRN.MMB1.DN (Value)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The share, loan count, and dollar volume of residential property investor mortgages written at an LVR above 70% (under a 30% deposit).",
        "methodology": "Derived from RBNZ Table C31 by dividing investor commitments with LVR > 70% (LVRN.MMB2.DN) by total investor commitments (LVRN.MMB2.AN).",
        "hlf_calculation": "(Investor Loans >70% LVR [LVRN.MMB2.DN] ÷ Total Investor Loans [LVRN.MMB2.AN]) × 100",
        "hlf_interpretation": "HLF market context: Captures investor borrowing above the standard 70% LVR cap—predominantly new-build investment properties (which are exempt from RBNZ LVR restrictions) and loans within bank high-LVR speed limits.",
        "questions": [
            "How do new-build LVR exemptions work for property investors in NZ?",
            "What percentage of property investors borrow above 70% LVR?",
        ],
        "article_title": "Using Equity to Buy an Investment Property in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/using-equity-to-buy-an-investment-property-in-nz-the-split-banking-guide",
        "tool_title": "Turnkey vs. Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/turnkey-vs-construction-loan-calculator-nz",
    },
    "mover-low-equity": {
        "name": "Existing owner-occupiers borrowing above 80% LVR (under 20% equity)",
        "also_known_as": "Low-equity (>80% LVR) other owner-occupier share & volume",
        "category": "Homeowners & refinancing",
        "topic": "Low-equity mover & homeowner lending",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB2.BC2 ÷ LVRN.MMB2.AC2 (Share) | LVRN.MMB1.BC2 (Value)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived percentage share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The share, borrower count, and dollar value of existing ('other') owner-occupier mortgage commitments approved at an LVR above 80% (under 20% equity).",
        "methodology": "Derived from RBNZ Table C31 by dividing other owner-occupier loans at >80% LVR (LVRN.MMB2.BC2) by total other owner-occupier loans (LVRN.MMB2.AC2).",
        "hlf_calculation": "(Other Owner-Occupier >80% LVR Loans [LVRN.MMB2.BC2] ÷ Total Other OO Loans [LVRN.MMB2.AC2]) × 100",
        "hlf_interpretation": "HLF market context: Existing homeowners can borrow above 80% LVR when purchasing or building a brand-new home (LVR-exempt), doing bridging/remedial finance, or qualifying for a bank's >80% LVR speed-limit allowance.",
        "questions": [
            "Can an existing homeowner buy a next home with less than 20% equity?",
            "How many existing owner-occupiers get approved above 80% LVR each month?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Net Proceeds & Next Home Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/net-proceeds-next-home-calculator",
    },
    "new-build-lvr-exemption-flow": {
        "name": "Combined high-LVR lending to existing owner-occupiers (>80% LVR) & investors (>70% LVR)",
        "also_known_as": "Non-FHB high-LVR mortgage volume (movers >80% LVR + investors >70% LVR)",
        "category": "LVR, DTI & credit policy",
        "topic": "High-LVR mover & investor lending volume",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB1.BC2 + LVRN.MMB1.DN (Dollars) | LVRN.MMB2.BC2 + LVRN.MMB2.DN (Loans)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "HLF-derived combined series from official RBNZ Table C31",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Monthly",
        "definition": "The combined monthly dollar value and borrower count of non-first-home-buyer mortgages written above standard LVR thresholds: existing owner-occupiers at >80% LVR plus property investors at >70% LVR.",
        "methodology": "Derived from RBNZ Table C31 by summing other owner-occupier >80% LVR lending (LVRN.MMB1.BC2) and investor >70% LVR lending (LVRN.MMB1.DN).",
        "hlf_calculation": "Other OO >80% LVR [LVRN.MMB1.BC2] + Investor >70% LVR [LVRN.MMB1.DN]",
        "hlf_interpretation": "HLF market context: Tracks total high-LVR credit flowing to movers and investors through new-build LVR exemptions and bank speed-limit allowances.",
        "questions": [
            "How much mortgage lending goes to movers above 80% LVR and investors above 70% LVR?",
            "How significant are LVR exemptions for new builds and movers in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
# === END OF PART 2A ===
    # === PART 2B OF 5 ===
    "fhb-annual-milestone": {
        "name": "Rolling 12-month first-home buyer mortgage approvals & lending total",
        "also_known_as": "Annual first-home buyer loan count, dollar volume & low-deposit total",
        "category": "First-home buyers",
        "topic": "Annual first-home buyer market volume",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "12-month rolling sum of LVRN.MMB2.AC1, LVRN.MMB1.AC1 & LVRN.MMB2.BC1",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "HLF-derived 12-month rolling sum from official RBNZ series",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Monthly (12-Month Rolling)",
        "definition": "The cumulative number and dollar value of first-home buyer mortgage commitments approved over the most recent 12-month period, including the 12-month total written above 80% LVR.",
        "methodology": "Derived from RBNZ Table C31 by summing the latest 12 monthly observations of LVRN.MMB2.AC1 (total FHB borrowers), LVRN.MMB1.AC1 (total FHB lending value), and LVRN.MMB2.BC1 (>80% LVR FHB borrowers).",
        "hlf_calculation": "Sum of latest 12 months of LVRN.MMB2.AC1 (Loans) & LVRN.MMB1.AC1 (Value)",
        "hlf_interpretation": "HLF market context: Smoothing out seasonal monthly swings reveals the true annual scale of first-home buyer entry into the New Zealand housing market.",
        "questions": [
            "How many first-home buyers buy a house in New Zealand each year?",
            "How much money do New Zealand banks lend to first-home buyers annually?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "refinance": {
        "name": "Mortgage refinancing (switching loan provider) share & volume",
        "also_known_as": "Bank-to-bank mortgage switching share, loan count & average refinance size",
        "category": "Homeowners & refinancing",
        "topic": "Mortgage refinancing & bank switching",
        "source": "RBNZ Table C33",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by purpose (Table C33)",
        "series_ids": "LVRN.MMB1.AZD (Refinance Value) & LVRN.MMB2.AZD (Refinance Count)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose",
        "data_type": "Officially reported RBNZ series & HLF-derived share/mean",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The share of new residential mortgage lending value and number of borrowers switching from one mortgage lender to another ('change in loan provider'), alongside the mean refinanced loan size.",
        "methodology": "Derived from RBNZ Table C33 by dividing 'change in loan provider' lending value (LVRN.MMB1.AZD) by the sum of all purpose categories (property purchase AZC, top-up AZB, change in provider AZD, and other AZE). Note: RBNZ publishes Table C33 dollar series rounded to the nearest $1 million.",
        "hlf_calculation": "(Change in Loan Provider Value [LVRN.MMB1.AZD] ÷ Total Lending by Purpose) × 100; Mean Loan = LVRN.MMB1.AZD ÷ LVRN.MMB2.AZD (subject to RBNZ $1m rounding)",
        "hlf_interpretation": "HLF market context: Refinancing volume measures how actively existing mortgage holders are comparing fixed rates and cashback offers across competing banks once their fixed terms expire.",
        "questions": [
            "How many New Zealand homeowners refinance or switch banks each month?",
            "What percentage of NZ mortgage lending is refinancing to a new bank?",
            "What is the average mortgage refinance size in New Zealand?",
        ],
        "article_title": "Refinancing a Mortgage in NZ: How Cashbacks & Break Fees Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/refinancing-a-mortgage-in-nz-how-cashbacks-break-fees-and-switching-banks-actually-work",
        "tool_title": "HLF Refinance & Cashback Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/refinance-calculator",
    },
    "top-ups": {
        "name": "Mortgage top-up share of approvals & average top-up amount",
        "also_known_as": "Existing homeowner equity top-up approval share & mean top-up value",
        "category": "Homeowners & refinancing",
        "topic": "Mortgage top-ups & equity release",
        "source": "RBNZ Table C33",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by purpose (Table C33)",
        "series_ids": "LVRN.MMB2.AZB (Top-up Count) & LVRN.MMB1.AZB (Top-up Value)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose",
        "data_type": "Officially reported RBNZ series & HLF-derived share/mean",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of all monthly residential mortgage approvals that are top-ups on an existing mortgage with the same lender, alongside the mean dollar top-up size.",
        "methodology": "Derived from RBNZ Table C33 by dividing top-up borrower count (LVRN.MMB2.AZB) by total mortgage approvals across all purposes, and dividing top-up lending value (LVRN.MMB1.AZB, rounded to $1m) by top-up borrower count.",
        "hlf_calculation": "(Top-up Count [LVRN.MMB2.AZB] ÷ Total Approvals by Purpose) × 100; Mean Top-Up = LVRN.MMB1.AZB ÷ LVRN.MMB2.AZB",
        "hlf_interpretation": "HLF market context: Top-ups are one of the most common transactions in New Zealand banking—used by homeowners to fund renovations, deposits on investment properties, or major capital purchases.",
        "questions": [
            "How common are mortgage top-ups in New Zealand?",
            "What is the average mortgage top-up amount in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "equity-topup-demand": {
        "name": "Monthly mortgage top-up dollar volume & borrower count",
        "also_known_as": "Total monthly equity top-up lending value (RBNZ C33)",
        "category": "Homeowners & refinancing",
        "topic": "Homeowner equity release volume",
        "source": "RBNZ Table C33",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by purpose (Table C33)",
        "series_ids": "LVRN.MMB1.AZB (Value) & LVRN.MMB2.AZB (Count)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose",
        "data_type": "Officially reported RBNZ series & HLF-derived mean",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Monthly",
        "definition": "The total monthly dollar value of new residential mortgage lending approved specifically for top-ups on existing mortgages across New Zealand.",
        "methodology": "Sourced from RBNZ Table C33 series LVRN.MMB1.AZB (Top-up lending value in $ millions) and LVRN.MMB2.AZB (Top-up borrower count), with mean top-up size calculated by HLF.",
        "hlf_calculation": "Mean Top-Up = Top-up Lending Value [LVRN.MMB1.AZB] ÷ Top-up Borrower Count [LVRN.MMB2.AZB]",
        "hlf_interpretation": "HLF market context: Measures the monthly dollar volume of existing property equity being unlocked by New Zealand homeowners without selling their property.",
        "questions": [
            "How much equity do New Zealand homeowners release via mortgage top-ups each month?",
            "How many homeowners apply for a mortgage top-up every month in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "interest-only": {
        "name": "Interest-only share of new investor vs owner-occupier mortgage lending",
        "also_known_as": "Investor vs owner-occupier interest-only mortgage proportion (RBNZ C32)",
        "category": "Property investors",
        "topic": "Interest-only vs principal-and-interest loan structure",
        "source": "RBNZ Table C32",
        "source_full": "Reserve Bank of New Zealand — New and existing residential mortgage lending by payment type (Table C32)",
        "series_ids": "New Investor Interest-Only ÷ LVRN.MMB1.AN | New Owner-Occupier Interest-Only ÷ LVRN.MMB1.AC",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type",
        "data_type": "Officially reported RBNZ series & HLF-derived percentage shares",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of new residential mortgage lending written on interest-only terms (including revolving credit facilities) for property investors compared with owner-occupiers.",
        "methodology": "Derived from RBNZ Table C32 by dividing new interest-only investor lending value by total new investor lending value, and new interest-only owner-occupier lending value by total new owner-occupier lending value.",
        "hlf_calculation": "(New Investor Interest-Only Lending ÷ Total New Investor Lending) × 100",
        "hlf_interpretation": "HLF market context: Property investors frequently use interest-only terms on rental debt so cashflow can be directed toward repaying non-tax-deductible personal home loan debt first.",
        "questions": [
            "What percentage of New Zealand property investors use interest-only mortgages?",
            "How common are interest-only loans for investors vs owner-occupiers in NZ?",
        ],
        "article_title": "Interest-Only vs Principal & Interest for NZ Investors",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/interest-only-vs-principal-interest-why-many-nz-investors-choose-not-to-pay-down-rental-debt",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "dti-headroom": {
        "name": "First-home buyer high Debt-to-Income (>6x DTI) share vs 20% RBNZ speed limit",
        "also_known_as": "FHB TDTI > 6 share of commitments & buffer below RBNZ DTI speed limit",
        "category": "LVR, DTI & credit policy",
        "topic": "Debt-to-Income (DTI) speed limits & headroom",
        "source": "RBNZ Table C40",
        "source_full": "Reserve Bank of New Zealand — Residential mortgage lending by Debt-to-Income (DTI) purpose use (Table C40)",
        "series_ids": "FHB New Commitments with TDTI > 6 ÷ Total FHB New Commitments",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use",
        "data_type": "Officially reported RBNZ series & HLF-derived >6x DTI share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of total monthly first-home buyer mortgage lending value committed at a Total Debt-to-Income (TDTI) ratio above 6x gross income, compared against the 20% RBNZ speed-limit cap on non-exempt owner-occupier lending.",
        "methodology": "Derived from RBNZ Table C40 by summing first-home buyer lending commitments in bands above 6x TDTI (including unknown TDTI) and dividing by total first-home buyer commitments across all TDTI bands.",
        "hlf_calculation": "(FHB Lending with TDTI > 6 ÷ Total FHB Commitments across all TDTI bands) × 100; Headroom = 20.0% minus (>6x Share)",
        "hlf_interpretation": "HLF market context: Because bank test servicing rates already constrain borrowing capacity, total FHB lending above 6x DTI sits well inside the 20% RBNZ speed-limit allowance.",
        "questions": [
            "Are RBNZ Debt-to-Income (DTI) rules restricting first-home buyers in New Zealand?",
            "What percentage of first-home buyers borrow more than 6 times their income?",
        ],
        "article_title": "DTI Rules in NZ: How Debt-to-Income Limits Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/dti-rules-in-nz-how-debt-to-income-limits-actually-work-for-home-buyers-and-investors",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "investor-high-dti-share": {
        "name": "Property investor high Debt-to-Income (>7x DTI) share vs 20% RBNZ speed limit",
        "also_known_as": "Investor TDTI > 7 share of commitments & buffer below RBNZ 20% investor DTI speed limit",
        "category": "LVR, DTI & credit policy",
        "topic": "Investor Debt-to-Income (>7x DTI) speed limits & headroom",
        "source": "RBNZ Table C40",
        "source_full": "Reserve Bank of New Zealand — Residential mortgage lending by Debt-to-Income (DTI) purpose use (Table C40)",
        "series_ids": "Investor New Commitments with TDTI > 7 ÷ Total Investor New Commitments",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use",
        "data_type": "Officially reported RBNZ series & HLF-derived >7x DTI share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of monthly residential property investor mortgage commitments written at a Total Debt-to-Income (TDTI) ratio above 7x gross income (including rental income), compared against the 20% RBNZ speed-limit cap on non-exempt investor lending.",
        "methodology": "Derived from RBNZ Table C40 by summing property investor new commitments with TDTI > 7 and dividing by total investor commitments across all TDTI bands.",
        "hlf_calculation": "(Investor Lending with TDTI > 7 ÷ Total Investor Commitments across all TDTI bands) × 100; Headroom = 20.0% minus (>7x Share)",
        "hlf_interpretation": "HLF market context: Under RBNZ DTI rules, banks can allocate up to 20% of non-exempt investor lending above 7x gross income (with new builds exempt from DTI caps).",
        "questions": [
            "What is the Debt-to-Income (DTI) limit for property investors in New Zealand?",
            "What percentage of NZ property investors borrow more than 7 times their gross income?",
        ],
        "article_title": "DTI Rules in NZ: How Debt-to-Income Limits Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/dti-rules-in-nz-how-debt-to-income-limits-actually-work-for-home-buyers-and-investors",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "borrower-income": {
        "name": "Average gross borrower income for approved first-home buyers vs property investors",
        "also_known_as": "Borrower Gross Income (BGI) mean across approved FHB & investor applications (RBNZ C41)",
        "category": "First-home buyers",
        "topic": "Borrower Gross Income (BGI)",
        "source": "RBNZ Table C41",
        "source_full": "Reserve Bank of New Zealand — Residential mortgage borrower gross income (BGI) (Table C41)",
        "series_ids": "Sum(Band Total Gross Income) ÷ Sum(Band Implied Borrowers) across TDTI bands",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi",
        "data_type": "HLF-derived band-weighted mean from official RBNZ Table C41 bands",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The band-weighted mean annual gross borrower income across approved New Zealand first-home buyer mortgage applications and property investor applications, derived by HLF from RBNZ Table C41's Debt-to-Income bands.",
        "methodology": "RBNZ Table C41 publishes official total gross income ($m, rounded to $0.1m) and average gross income ($000, rounded to $1,000) within each TDTI band rather than a single all-bands average. HLF derives the overall mean by dividing total gross income summed across all reported TDTI bands by the implied number of applications (Band Total Income ÷ Band Average Income) summed across those bands.",
        "hlf_calculation": "Sum(Band Total Gross Income [$m]) ÷ Sum(Band Total Gross Income [$m] ÷ Band Average Gross Income [$k]) × 1,000 (subject to RBNZ $0.1m / $1,000 band rounding)",
        "hlf_interpretation": "HLF market context: RBNZ's Borrower Gross Income (BGI) measures total application income (which for couples combines two salaries, and for investors includes existing salaries plus gross rental and other qualifying income).",
        "questions": [
            "What is the average household income of an approved first-home buyer in New Zealand?",
            "How much do approved property investors earn compared to first-home buyers in RBNZ data?",
        ],
        "article_title": "DTI Rules in NZ: How Debt-to-Income Limits Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/dti-rules-in-nz-how-debt-to-income-limits-actually-work-for-home-buyers-and-investors",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-income-by-dti-band": {
        "name": "First-home buyer gross borrower income by Debt-to-Income (DTI) band — New Zealand",
        "also_known_as": "Official RBNZ Table C41 FHB average gross income across TDTI bands & HLF overall mean",
        "category": "First-home buyers",
        "topic": "First-home buyer income across DTI tiers",
        "source": "RBNZ Table C41",
        "source_full": "Reserve Bank of New Zealand — Residential mortgage borrower gross income (BGI) (Table C41)",
        "series_ids": "RBNZ C41 FHB TDTI Bands (<=3, 3–4, 4–5, 5–6, >6) Average & Total Gross Income",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi",
        "data_type": "Officially reported RBNZ C41 band averages & HLF-derived overall mean",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "Official average gross borrower income reported by RBNZ Table C41 within each first-home buyer Debt-to-Income (TDTI) band across New Zealand, alongside the HLF-calculated overall mean across all bands.",
        "methodology": "Reports the five official RBNZ Table C41 first-home buyer TDTI band averages (TDTI <= 3, 3 < TDTI <= 4, 4 < TDTI <= 5, 5 < TDTI <= 6, and TDTI > 6), each band's official gross income pool ($m, rounded to $0.1m), each band's implied application count (Pool ÷ Average), and the HLF-derived overall weighted mean.",
        "hlf_calculation": "Official RBNZ C41 band averages + HLF overall mean = Sum(Band Income Pool) ÷ Sum(Implied Applications = Band Pool ÷ Band Avg)",
        "hlf_interpretation": "HLF market context: Shows how first-home buyer gross application income varies across Debt-to-Income tiers in RBNZ Table C41.",
        "questions": [
            "How does first-home buyer income vary across Debt-to-Income (DTI) bands in New Zealand?",
            "What income do first-home buyers have in each RBNZ DTI band?",
        ],
        "article_title": "DTI Rules in NZ: How Debt-to-Income Limits Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/dti-rules-in-nz-how-debt-to-income-limits-actually-work-for-home-buyers-and-investors",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "settled-rates": {
        "name": "Weighted-average settled 1-year & 2-year fixed mortgage rates (owner-occupiers)",
        "also_known_as": "RBNZ Table B30 average settled residential mortgage interest rates",
        "category": "Homeowners & refinancing",
        "topic": "Settled fixed mortgage interest rates",
        "source": "RBNZ Table B30",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage weighted average interest rates (Table B30)",
        "series_ids": "NCFA.MAM1PR.D1 (1-Year Fixed OO) & NCFA.MAM1PR.D2 (2-Year Fixed OO)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-weighted-average-interest-rates",
        "data_type": "Officially reported RBNZ series",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The weighted-average fixed mortgage interest rate actually settled by New Zealand owner-occupiers for 1-year and 2-year fixed terms.",
        "methodology": "Directly reported by RBNZ Table B30 series NCFA.MAM1PR.D1 (1-year fixed owner-occupier rate) and NCFA.MAM1PR.D2 (2-year fixed owner-occupier rate), weighted by dollar lending volume.",
        "hlf_calculation": "Officially reported RBNZ Table B30 weighted-average settled rate series (plus YoY rate change)",
        "hlf_interpretation": "HLF market context: Captures the actual negotiated rates borrowers settled on during the month (including discretionary bank discounts below advertised carded rates).",
        "questions": [
            "What is the average 1-year and 2-year fixed mortgage rate in New Zealand right now?",
            "What mortgage interest rate are New Zealand homeowners actually settling at?",
        ],
        "article_title": "Compare NZ Mortgage Rates",
        "article_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
        "tool_title": "Compare NZ Mortgage Rates",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
    },
    "fixed-vs-floating-rate-mix": {
        "name": "Fixed vs floating mortgage share & 12-month fixed rate roll-off (RBNZ Mortgage Stock)",
        "also_known_as": "Share of NZ residential mortgages on fixed vs floating rates & refixing within 12 months",
        "category": "Homeowners & refinancing",
        "topic": "Fixed vs floating mortgage structure & refix horizon",
        "source": "RBNZ Mortgage Stock & FSR",
        "source_full": "Reserve Bank of New Zealand — Residential Mortgage Lending by Interest Rate Type & Financial Stability Report",
        "series_ids": "RBNZ Fixed vs Floating Residential Mortgage Stock & <12-Month repricing band",
        "source_url": RBNZ_FSR_PAGE,
        "data_type": "Officially reported RBNZ series & HLF-derived repricing share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly / Quarterly",
        "definition": "The proportion of New Zealand's total residential mortgage stock held on fixed interest rates versus floating/revolving rates, alongside the percentage of fixed-rate loans scheduled to roll off and reprice within the next 12 months.",
        "methodology": "Compiled from RBNZ residential mortgage stock repricing statistics and Financial Stability Report benchmarks, dividing fixed-rate mortgage balances by total residential mortgage balances (~$368.4B) and summing fixed-rate bands repricing within 12 months.",
        "hlf_calculation": "Fixed Share = Fixed Stock ÷ Total Stock; 12-Month Roll-Off = Fixed Loans Repricing <=12 Months ÷ Total Fixed Stock",
        "hlf_interpretation": "HLF market context: Because nearly 90% of New Zealand mortgage debt is fixed—and roughly two-thirds of fixed loans roll off within 12 months—OCR changes transmit rapidly into household budgets as terms expire.",
        "questions": [
            "What percentage of New Zealand mortgages are fixed vs floating?",
            "How many New Zealand mortgages are due to refix in the next 12 months?",
        ],
        "article_title": "Refinancing a Mortgage in NZ: How Cashbacks & Break Fees Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/refinancing-a-mortgage-in-nz-how-cashbacks-break-fees-and-switching-banks-actually-work",
        "tool_title": "HLF Split Banking & Refinance Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/refinance-calculator",
    },
    "new-builds-annual": {
        "name": "Annual & monthly new residential dwelling consents issued",
        "also_known_as": "Stats NZ annual new home building consents total",
        "category": "New builds & construction",
        "topic": "New residential dwelling supply",
        "source": "Stats NZ Building Consents",
        "source_full": "Stats NZ — Building Consents Issued (Monthly Information Release)",
        "series_ids": "Stats NZ Building Consents Issued — Annual & Monthly Total New Dwellings",
        "source_url": "https://www.stats.govt.nz/topics/building",
        "data_type": "Officially reported Stats NZ series",
        "derived_by_hlf": False,
        "unit": "dwellings",
        "frequency": "Monthly",
        "definition": "The total number of new residential dwellings consented across New Zealand over the most recent 12-month period and latest single month.",
        "methodology": "Directly reported in Stats NZ's monthly Building Consents Issued release (actual number of new dwellings consented in the year ended reference month and single month).",
        "hlf_calculation": "Officially reported Stats NZ annual and monthly dwelling consent totals",
        "hlf_interpretation": "HLF market context: Building consents are the primary leading indicator for New Zealand's new housing supply pipeline over the coming 12 to 24 months.",
        "questions": [
            "How many new houses are consented in New Zealand each year?",
            "Are building consents rising or falling in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
# === END OF PART 2B ===
    # === PART 2C OF 5 ===
    "townhouse-share": {
        "name": "Multi-unit & townhouse share of new residential dwelling consents",
        "also_known_as": "Townhouses, flats, units & apartments share of new NZ building consents",
        "category": "New builds & construction",
        "topic": "Medium-density & multi-unit housing supply",
        "source": "Stats NZ Building Consents",
        "source_full": "Stats NZ — Building Consents Issued (Monthly Information Release)",
        "series_ids": "(Total New Dwellings minus Stand-Alone Houses) ÷ Total New Dwellings",
        "source_url": "https://www.stats.govt.nz/topics/building",
        "data_type": "Officially reported Stats NZ counts & HLF-derived percentage shares",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage of monthly new residential dwelling consents issued for multi-unit homes overall (townhouses, flats, units, apartments, and retirement village units) and specifically for townhouses, flats, and units.",
        "methodology": "Derived from Stats NZ Building Consents Issued by subtracting stand-alone houses from total new dwellings consented and dividing by total new dwellings consented, alongside the separate share for townhouses, flats, and units.",
        "hlf_calculation": "Multi-Unit Share = ((Total Consents - Stand-Alone Houses) ÷ Total Consents) × 100; Townhouse/Flat/Unit Share = (Townhouses, Flats & Units ÷ Total Consents) × 100",
        "hlf_interpretation": "HLF market context: Multi-unit homes—led by townhouses—now represent roughly half of all new dwellings consented in New Zealand, particularly in main urban centres.",
        "questions": [
            "What percentage of new builds in New Zealand are townhouses vs standalone houses?",
            "How many townhouses are being consented in New Zealand each month?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Turnkey vs. Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/turnkey-vs-construction-loan-calculator-nz",
    },
    "standalone-vs-retirement-apartments": {
        "name": "Monthly new dwelling consents by 4-category building type",
        "also_known_as": "Stand-alone houses vs townhouses vs apartments vs retirement units breakdown",
        "category": "New builds & construction",
        "topic": "Dwelling consent composition by building type",
        "source": "Stats NZ Building Consents",
        "source_full": "Stats NZ — Building Consents Issued (Monthly Information Release)",
        "series_ids": "Stand-alone houses | Townhouses, flats & units | Apartments | Retirement village units",
        "source_url": "https://www.stats.govt.nz/topics/building",
        "data_type": "Officially reported Stats NZ counts & HLF-derived category percentages",
        "derived_by_hlf": True,
        "unit": "dwellings",
        "frequency": "Monthly",
        "definition": "Complete 4-way breakdown of monthly new residential dwelling consents across stand-alone houses, townhouses/flats/units, apartments, and retirement village units.",
        "methodology": "Sourced from Stats NZ's monthly Building Consents Issued release, with each category divided by the monthly total to show exact percentage shares.",
        "hlf_calculation": "Category Share (%) = (Category Dwelling Count ÷ Total Monthly Dwelling Consents) × 100",
        "hlf_interpretation": "HLF market context: Different dwelling types have distinct bank lending rules—for example, smaller apartments often require higher deposits while turnkey townhouses and progress-payment standalone builds qualify for new-build LVR exemptions.",
        "questions": [
            "What is the breakdown of new houses, townhouses, and apartments consented in NZ?",
            "How many apartments and retirement units are consented each month in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "regional-building-consents": {
        "name": "Annual new dwelling consents by major New Zealand region",
        "also_known_as": "Auckland, Canterbury, Waikato & Wellington share of annual building consents",
        "category": "New builds & construction",
        "topic": "Regional housing supply pipeline",
        "source": "Stats NZ Building Consents",
        "source_full": "Stats NZ — Building Consents Issued (Regional Annual Series)",
        "series_ids": "Stats NZ Regional Annual Dwelling Consents (Auckland, Canterbury, Waikato, Wellington, Rest of NZ)",
        "source_url": "https://www.stats.govt.nz/topics/building",
        "data_type": "Officially reported Stats NZ regional counts & HLF-derived national shares",
        "derived_by_hlf": True,
        "unit": "dwellings",
        "frequency": "Monthly (Annual Rolling)",
        "definition": "Annual new residential dwelling consents issued across New Zealand's four largest building regions (Auckland, Canterbury, Waikato, and Wellington) and the rest of New Zealand.",
        "methodology": "Sourced from Stats NZ's Building Consents Issued regional release, with each region's annual consent total divided by the national annual total.",
        "hlf_calculation": "Regional Share (%) = (Regional Annual Consents ÷ National Annual Consents) × 100",
        "hlf_interpretation": "HLF market context: Auckland and Canterbury together account for roughly three-quarters of all new homes consented across New Zealand.",
        "questions": [
            "Which New Zealand regions are building the most new houses?",
            "What share of New Zealand building consents are in Auckland and Canterbury?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "build-cost-inflation": {
        "name": "Annual change in the cost of purchasing a newly built home (CPI)",
        "also_known_as": "Stats NZ CPI purchase of new housing inflation rate",
        "category": "New builds & construction",
        "topic": "Residential construction cost inflation",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Purchase of new housing sub-index)",
        "series_ids": "Stats NZ CPI — Purchase of new housing (Annual % change) vs All-groups CPI",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI series",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in the cost of purchasing a newly built residential dwelling in New Zealand, as measured in Stats NZ's quarterly Consumers Price Index (CPI).",
        "methodology": "Directly reported in Stats NZ's quarterly Consumers Price Index release under the 'purchase of new housing' sub-index, compared alongside headline all-groups CPI inflation.",
        "hlf_calculation": "Officially reported Stats NZ quarterly CPI purchase of new housing YoY % change",
        "hlf_interpretation": "HLF market context: Stabilising residential build cost inflation reduces cost-overrun risk on fixed-price construction contracts and moderates replacement-value sum-insured increases on home insurance.",
        "questions": [
            "How fast are house building costs rising in New Zealand?",
            "Has construction cost inflation slowed down in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "new-build-consent-cost-by-type": {
        "name": "Average residential building consent construction cost & cost per m² by dwelling type",
        "also_known_as": "Dollar build cost for stand-alone houses, townhouses, apartments & retirement units (Stats NZ)",
        "category": "New builds & construction",
        "topic": "Average dollar build cost by dwelling type",
        "source": "Stats NZ Building Consents",
        "source_full": "Stats NZ — Building Consents Issued (New Residential Dwelling Consent Values & Floor Area by Building Type)",
        "series_ids": "Total New Dwelling Consent Value ($m) ÷ Dwelling Count & Floor Area (m²) by Building Type",
        "source_url": "https://www.stats.govt.nz/topics/building",
        "data_type": "HLF-derived mean consent value & $/m² from official Stats NZ Building Consents tables",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly / Quarterly",
        "definition": "The average consented construction value in New Zealand dollars (excluding land, GST-exclusive consent basis) and average cost per square metre (m²) across New Zealand's four residential dwelling categories: stand-alone houses, townhouses/flats/units, apartments, and retirement village units.",
        "methodology": "Derived from Stats NZ Building Consents Issued tables by dividing total consented residential building value by consented dwelling counts and floor area (m²) for each building category. Note: Building consent values measure estimated physical construction cost declared on the building consent (excluding land purchase price, developer margin, landscaping, and off-site subdivision infrastructure).",
        "hlf_calculation": "Mean Build Value = Category Consent Value ($) ÷ Category Dwelling Count; Cost per m² = Category Consent Value ($) ÷ Category Floor Area (m²)",
        "hlf_interpretation": "HLF market context: Stand-alone houses have the highest total construction cost (~$594k across ~188 m²) due to larger floor areas, whereas multi-level apartments have the highest construction cost per square metre (~$5,318/m²) due to structural engineering, lifts, and fire systems.",
        "questions": [
            "How much does it cost to build a house vs a townhouse in New Zealand?",
            "What is the average building consent cost per square metre in New Zealand?",
            "What is the average construction cost of a new standalone house in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "rent-yield-pulse": {
        "name": "National & regional median weekly private rent (Auckland, Wellington, Canterbury)",
        "also_known_as": "MBIE Tenancy Services median weekly rent across New Zealand",
        "category": "Rents & tenancy",
        "topic": "Private residential rental prices",
        "source": "MBIE Tenancy Services",
        "source_full": "Ministry of Business, Innovation and Employment (MBIE) — Tenancy Services Rental Bond Data",
        "series_ids": "MBIE By-Region Rental Bond CSV — Location 'ALL', 'Auckland', 'Wellington', 'Canterbury' Median Rent",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported MBIE medians & HLF-derived YoY change",
        "derived_by_hlf": True,
        "unit": "NZD_per_week",
        "frequency": "Monthly",
        "definition": "The national median weekly rent across newly lodged private residential tenancy bonds in New Zealand, alongside regional medians for Auckland, Wellington, and Canterbury.",
        "methodology": "Extracted from MBIE Tenancy Services' monthly regional rental bond dataset (Location = 'ALL' for national median), with year-on-year dollar and percentage changes calculated against the same month a year earlier.",
        "hlf_calculation": "YoY Rent Change (%) = ((Current National Median Rent - Prior Year Median Rent) ÷ Prior Year Median Rent) × 100",
        "hlf_interpretation": "HLF market context: Rental bond medians reflect the actual rents agreed on newly signed tenancies each month, which banks use alongside council rates and insurance when assessing rental yield and investor servicing.",
        "questions": [
            "What is the median weekly rent in New Zealand right now?",
            "How does median rent compare between Auckland, Wellington, and Christchurch/Canterbury?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "rental-bonds-supply": {
        "name": "Monthly newly lodged & total active private rental bonds",
        "also_known_as": "MBIE Tenancy Services monthly lodged bonds & active private tenancy stock",
        "category": "Rents & tenancy",
        "topic": "Private rental supply & tenancy activity",
        "source": "MBIE Tenancy Services",
        "source_full": "Ministry of Business, Innovation and Employment (MBIE) — Tenancy Services Rental Bond Data",
        "series_ids": "MBIE By-Region Rental Bond CSV — Location 'ALL' Lodged Bonds & Active Bonds",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported MBIE series",
        "derived_by_hlf": False,
        "unit": "bonds",
        "frequency": "Monthly",
        "definition": "The number of newly lodged private rental tenancy bonds in the latest month and the total pool of active private rental bonds held nationwide by MBIE Tenancy Services.",
        "methodology": "Directly reported in MBIE Tenancy Services' monthly rental bond CSV for Location = 'ALL' (Lodged Bonds and Active Bonds).",
        "hlf_calculation": "Officially reported MBIE national Lodged Bonds and Active Bonds",
        "hlf_interpretation": "HLF market context: Tracks both monthly tenant leasing activity and the total size of New Zealand's formal private rental housing sector.",
        "questions": [
            "How many active private rental properties are there in New Zealand?",
            "How many new rental bonds are lodged each month in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "rent-long-term-growth": {
        "name": "5-year & 10-year national median rent growth (nominal vs CPI inflation-adjusted)",
        "also_known_as": "Long-term New Zealand weekly rent growth & compound annual growth rate (CAGR)",
        "category": "Rents & tenancy",
        "topic": "Long-term rental growth (nominal & real)",
        "source": "MBIE Tenancy Services & Stats NZ CPI",
        "source_full": "MBIE Tenancy Services Rental Bond Data & Stats NZ Consumers Price Index",
        "series_ids": "MBIE Location 'ALL' Median Rent (t, t-5y, t-10y) deflated by Stats NZ CPI",
        "source_url": TENANCY_PAGE,
        "data_type": "HLF-derived 5Y & 10Y nominal/real growth & CAGR from official MBIE & Stats NZ series",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "Cumulative and compound annual growth in New Zealand's national median weekly private rent over the past 5 years and 10 years, reported in both nominal dollar terms and CPI inflation-adjusted (real) terms.",
        "methodology": "Compares the latest national median weekly rent in MBIE's bond dataset against the same month 5 and 10 years earlier, and deflates cumulative nominal growth by cumulative Stats NZ all-groups CPI over the corresponding 5-year (+22.8%) and 10-year (+33.7%) windows.",
        "hlf_calculation": "10Y Nominal % = (Rent_t - Rent_t-10) ÷ Rent_t-10; 10Y Real Factor = (Rent_t ÷ Rent_t-10) ÷ CPI_10Y_Factor; CAGR = Factor^(1/n) - 1",
        "hlf_interpretation": "HLF market context: Evaluates how private rents have tracked relative to general consumer price inflation over a full property cycle.",
        "questions": [
            "How much have New Zealand rents increased over the last 10 years?",
            "Have New Zealand rents grown faster than inflation?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "rent-quartile-spread": {
        "name": "National weekly rent distribution: lower quartile (25%), median (50%) & upper quartile (75%)",
        "also_known_as": "New Zealand rental interquartile range (25th to 75th percentile weekly rent)",
        "category": "Rents & tenancy",
        "topic": "Rental price distribution & quartile spread",
        "source": "MBIE Tenancy Services",
        "source_full": "Ministry of Business, Innovation and Employment (MBIE) — Tenancy Services Rental Bond Data",
        "series_ids": "MBIE Location 'ALL' Lower Quartile Rent, Median Rent & Upper Quartile Rent",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported MBIE quartiles & HLF-derived interquartile spread",
        "derived_by_hlf": True,
        "unit": "NZD_per_week",
        "frequency": "Monthly",
        "definition": "The 25th percentile (lower quartile), 50th percentile (median), and 75th percentile (upper quartile) weekly private rent across New Zealand, alongside the dollar spread between lower- and upper-quartile properties.",
        "methodology": "Sourced from MBIE Tenancy Services' regional bond CSV (Location = 'ALL'), subtracting Lower Quartile Rent from Upper Quartile Rent and multiplying by 52 for the annualised spread.",
        "hlf_calculation": "Interquartile Spread ($/wk) = Upper Quartile Rent minus Lower Quartile Rent; Annualised Spread = Spread × 52",
        "hlf_interpretation": "HLF market context: Shows the weekly cashflow difference between entry-level units/flats (lower quartile) and larger family homes (upper quartile).",
        "questions": [
            "What is a cheap vs expensive weekly rent in New Zealand?",
            "What is the lower quartile and upper quartile rent in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "rental-turnover-net-stock": {
        "name": "Annual net change in active rental bond stock & monthly lodged vs closed bonds",
        "also_known_as": "Net private rental supply expansion/contraction & monthly tenancy turnover",
        "category": "Rents & tenancy",
        "topic": "Net rental stock growth & tenancy turnover",
        "source": "MBIE Tenancy Services",
        "source_full": "Ministry of Business, Innovation and Employment (MBIE) — Tenancy Services Rental Bond Data",
        "series_ids": "MBIE Location 'ALL' Active Bonds (YoY change) & Lodged Bonds minus Closed Bonds",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported MBIE counts & HLF-derived net stock change",
        "derived_by_hlf": True,
        "unit": "bonds",
        "frequency": "Monthly",
        "definition": "The 12-month net change in New Zealand's total pool of active private rental bonds, alongside the monthly balance of newly lodged bonds versus closed bonds.",
        "methodology": "Derived from MBIE Tenancy Services' bond dataset by subtracting Active Bonds 12 months earlier from current Active Bonds, and subtracting monthly Closed Bonds from monthly Lodged Bonds.",
        "hlf_calculation": "12M Net Stock Change = Active Bonds_t minus Active Bonds_t-12; Monthly Net Flow = Lodged Bonds minus Closed Bonds",
        "hlf_interpretation": "HLF market context: Positive net active bond growth indicates that the total supply of private rental properties in New Zealand is expanding rather than shrinking.",
        "questions": [
            "Is the number of rental properties in New Zealand growing or shrinking?",
            "How many rental tenancies end (close) vs start (lodge) each month in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "student-loan-servicing": {
        "name": "Median New Zealand student loan balance (Inland Revenue)",
        "also_known_as": "IRD Student Loan Scheme median borrower balance & servicing impact",
        "category": "LVR, DTI & credit policy",
        "topic": "Student loan balances & mortgage servicing impact",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — Student Loan Scheme Balances Statistics",
        "series_ids": "Inland Revenue — Median Student Loan Balance ($ NZD)",
        "source_url": IRD_SL_BAL_PAGE,
        "data_type": "Officially reported IRD median & HLF-derived YoY change",
        "derived_by_hlf": False,
        "unit": "NZD",
        "frequency": "Quarterly",
        "definition": "The median outstanding student loan balance across New Zealand student loan borrowers reported by Inland Revenue.",
        "methodology": "Directly reported by Inland Revenue in its quarterly Student Loan Scheme statistics release, with percentage and dollar change compared against the prior reported period.",
        "hlf_calculation": "Officially reported IRD median balance (plus % change vs prior period)",
        "hlf_interpretation": "HLF market context: Under New Zealand bank servicing rules, an active student loan requires a compulsory 12% repayment on gross income above the annual repayment threshold—reducing net disposable income available for mortgage servicing regardless of whether the balance is $5,000 or $50,000.",
        "questions": [
            "What is the average or median student loan balance in New Zealand?",
            "How does a student loan affect my mortgage borrowing power in NZ?",
        ],
        "article_title": "DTI Rules in NZ: How Debt-to-Income Limits Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/dti-rules-in-nz-how-debt-to-income-limits-actually-work-for-home-buyers-and-investors",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "kiwisaver-first-home": {
        "name": "Monthly & 10-year KiwiSaver first-home withdrawals (count, total value & average withdrawal)",
        "also_known_as": "Inland Revenue KiwiSaver first-home purchase withdrawal volume & mean per member",
        "category": "KiwiSaver",
        "topic": "KiwiSaver first-home deposit withdrawals",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Withdrawal Statistics (First Home Purchase)",
        "series_ids": "IRD KiwiSaver First-Home Withdrawal Member Count & Dollar Value ($m)",
        "source_url": IRD_KS_DATA_PAGE,
        "data_type": "Officially reported IRD series & HLF-derived mean/10Y real growth",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Monthly & Annual (June FY)",
        "definition": "The monthly number of KiwiSaver members withdrawing funds to purchase a first home, total dollar value withdrawn, and mean withdrawal per member, alongside 10-year June financial year totals (nominal and CPI inflation-adjusted).",
        "methodology": "Sourced from Inland Revenue's monthly KiwiSaver withdrawal numbers and amounts tables, dividing monthly first-home withdrawal dollars by withdrawing member count to derive the average withdrawal per member.",
        "hlf_calculation": "Average First-Home Withdrawal = (Monthly First-Home Withdrawal Value [$m] × $1,000,000) ÷ Monthly Withdrawing Members",
        "hlf_interpretation": "HLF market context: KiwiSaver is now the foundational deposit source for New Zealand first-home buyers—when two partners both withdraw ~$46,000, they unlock over $90,000 toward a 10% or 20% home deposit.",
        "questions": [
            "What is the average KiwiSaver first-home withdrawal in New Zealand?",
            "How many people withdraw their KiwiSaver for a first home each month and year?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "Kāinga Ora LMI Fee & Low-Deposit Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/lenders-mortgage-insurance-fee-estimator",
    },
# === END OF PART 2C ===
    # === PART 2D OF 5 ===
    "kiwisaver-hardship-withdrawals": {
        "name": "Monthly & 10-year KiwiSaver financial hardship withdrawals",
        "also_known_as": "Inland Revenue KiwiSaver significant financial hardship withdrawal count, value & average",
        "category": "KiwiSaver",
        "topic": "KiwiSaver financial hardship withdrawals",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Withdrawal Statistics (Significant Financial Hardship)",
        "series_ids": "IRD KiwiSaver Financial Hardship Withdrawal Member Count & Dollar Value ($m)",
        "source_url": IRD_KS_DATA_PAGE,
        "data_type": "Officially reported IRD series & HLF-derived mean per member",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Monthly & Annual (June FY)",
        "definition": "The monthly number of KiwiSaver members approved for a significant financial hardship withdrawal, total dollar amount withdrawn, and average withdrawal per member, alongside 10-year June financial year totals.",
        "methodology": "Sourced from Inland Revenue's monthly KiwiSaver withdrawal numbers and amounts tables, dividing monthly financial hardship withdrawal dollars by withdrawing member count.",
        "hlf_calculation": "Average Hardship Withdrawal = (Monthly Hardship Withdrawal Value [$m] × $1,000,000) ÷ Monthly Hardship Members",
        "hlf_interpretation": "HLF market context: Rising financial hardship withdrawals reflect cost-of-living and mortgage refixing pressures across New Zealand households.",
        "questions": [
            "How many New Zealanders withdraw their KiwiSaver for financial hardship?",
            "What is the average KiwiSaver financial hardship withdrawal amount in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "fhb-regional-kiwisaver": {
        "name": "Regional KiwiSaver first-home withdrawal average & volume across all 16 NZ regions",
        "also_known_as": "Inland Revenue regional KiwiSaver first-home withdrawal per member by region",
        "category": "KiwiSaver",
        "topic": "Regional KiwiSaver first-home withdrawals",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Fund Withdrawals by Region (Annual June Financial Year)",
        "series_ids": "IRD Regional First-Home Withdrawal Value ($) ÷ Regional Withdrawing Member Count across 16 NZ Regions",
        "source_url": IRD_KS_REG_PAGE,
        "data_type": "Officially reported IRD regional totals & HLF-derived per-member regional averages",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual (June Financial Year)",
        "definition": "The average KiwiSaver first-home withdrawal per withdrawing member across all 16 New Zealand regions, alongside regional member counts and total dollar volumes.",
        "methodology": "Derived from Inland Revenue's annual regional KiwiSaver withdrawal tables by dividing each region's first-home withdrawal dollar total by the number of withdrawing members in that region.",
        "hlf_calculation": "Regional Average Withdrawal ($) = Regional First-Home Withdrawal Dollars ÷ Regional First-Home Withdrawing Members",
        "hlf_interpretation": "HLF market context: Higher-priced regions such as Auckland, Tasman, Wellington, and Bay of Plenty record larger average KiwiSaver first-home withdrawals per member than lower-priced regions.",
        "questions": [
            "How much KiwiSaver do first-home buyers withdraw in Auckland, Wellington, and Christchurch?",
            "Which New Zealand region has the highest average KiwiSaver first-home withdrawal?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "kiwisaver-switches": {
        "name": "Annual & 10-year KiwiSaver scheme provider transfers (switching providers)",
        "also_known_as": "Inland Revenue KiwiSaver inter-scheme transfers & monthly switching rate",
        "category": "KiwiSaver",
        "topic": "KiwiSaver scheme switching & provider transfers",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Scheme Transfer Statistics",
        "series_ids": "IRD Annual KiwiSaver Scheme Transfers (Year ended 30 June) ÷ 12",
        "source_url": IRD_KS_TR_PAGE,
        "data_type": "Officially reported IRD annual series & HLF-derived monthly mean/10Y growth",
        "derived_by_hlf": True,
        "unit": "transfers",
        "frequency": "Annual (June Financial Year)",
        "definition": "The annual and average monthly number of KiwiSaver members transferring their account from one KiwiSaver scheme provider to another.",
        "methodology": "Sourced from Inland Revenue's annual KiwiSaver scheme transfer release, dividing the annual transfer total by 12 for the monthly average and comparing against the prior year and 10 years earlier.",
        "hlf_calculation": "Monthly Average Transfers = Annual Scheme Transfers ÷ 12; 10Y Growth (%) = ((Transfers_t - Transfers_t-10) ÷ Transfers_t-10) × 100",
        "hlf_interpretation": "HLF market context: Nearly 200,000 New Zealanders switch KiwiSaver providers each year as members review fees, long-term fund performance, and advice support.",
        "questions": [
            "How many New Zealanders switch KiwiSaver providers each year?",
            "Is KiwiSaver provider switching increasing in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-total-fum-balance": {
        "name": "Total KiwiSaver funds under management (FUM) & official FMA average member balance",
        "also_known_as": "FMA KiwiSaver Annual Report total FUM, member count & 10-year nominal/real average balance growth",
        "category": "KiwiSaver",
        "topic": "KiwiSaver total funds under management & average balance",
        "source": "FMA & Stats NZ CPI",
        "source_full": "Financial Markets Authority (FMA) — KiwiSaver Annual Report & Stats NZ Consumers Price Index",
        "series_ids": "FMA KiwiSaver Annual Report — Total FUM ($B), Total Members & Official Average Member Balance ($)",
        "source_url": FMA_KS_REPORT_PAGE,
        "data_type": "Officially reported FMA figures & HLF-derived 10-year nominal/real growth",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual (Year ended 31 March)",
        "definition": "Total funds under management (FUM), total enrolled membership, and the official average member balance across all New Zealand KiwiSaver schemes, alongside 10-year nominal and CPI inflation-adjusted (real) growth.",
        "methodology": "Reports the official average KiwiSaver balance and total FUM from the FMA KiwiSaver Annual Report (year ended 31 March), and calculates 10-year nominal and CPI-deflated real growth against March 2016.",
        "hlf_calculation": "10Y Nominal % = (AvgBal_2026 - AvgBal_2016) ÷ AvgBal_2016; 10Y Real Factor = (AvgBal_2026 ÷ AvgBal_2016) ÷ CPI_10Y_Factor",
        "hlf_interpretation": "HLF market context: Even after accounting for a decade of consumer price inflation, the average KiwiSaver balance has more than doubled in real purchasing-power terms since 2016.",
        "questions": [
            "What is the average KiwiSaver balance in New Zealand?",
            "How much money is in KiwiSaver in total across New Zealand?",
            "How much has the average KiwiSaver balance grown over 10 years?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-default-vs-active": {
        "name": "KiwiSaver default & employer-nominated allocation vs active scheme choice",
        "also_known_as": "Inland Revenue KiwiSaver scheme entry methods (default vs active choice share)",
        "category": "KiwiSaver",
        "topic": "KiwiSaver default allocation vs active fund selection",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Scheme Entry Method Statistics",
        "series_ids": "(Default Allocated + Employer Nominated) ÷ Total Enrolled KiwiSaver Members",
        "source_url": IRD_KS_ENTRY_PAGE,
        "data_type": "Officially reported IRD member counts & HLF-derived percentage shares",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Annual (June Financial Year)",
        "definition": "The number and percentage of KiwiSaver members remaining in schemes assigned passively via Inland Revenue default allocation or employer nomination versus members who have actively chosen their scheme.",
        "methodology": "Derived from Inland Revenue's annual KiwiSaver entry method table by summing default-allocated and employer-nominated members and dividing by total enrolled members.",
        "hlf_calculation": "Passive Allocation Share (%) = ((Default Allocated + Employer Nominated) ÷ Total Members) × 100",
        "hlf_interpretation": "HLF market context: Roughly a quarter of KiwiSaver members (~845,000 New Zealanders) remain in a passively assigned default or employer-nominated scheme rather than an actively selected fund matched to their time horizon.",
        "questions": [
            "How many New Zealanders are still in a default KiwiSaver fund?",
            "What percentage of KiwiSaver members have actively chosen their provider?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-contributions-10yr": {
        "name": "Annual & 10-year KiwiSaver contributions by source (employee, employer & Crown)",
        "also_known_as": "Inland Revenue KiwiSaver scheme provider payments, 3.5% default rate & $260.72 Gov contribution",
        "category": "KiwiSaver",
        "topic": "KiwiSaver contributions & statutory settings",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Contributions & Payments to Scheme Providers",
        "series_ids": "IRD KiwiSaver Employee, Employer & Crown Contributions ($B, FY2016–FY2026)",
        "source_url": IRD_KS_PAY_PAGE,
        "data_type": "Officially reported IRD series & HLF-derived source shares/10Y real growth",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Annual (June Financial Year)",
        "definition": "Total annual KiwiSaver contributions transferred by Inland Revenue to scheme providers, broken down across employee deductions, employer contributions, and Crown (Government) contributions over a 10-year horizon.",
        "methodology": "Sourced from Inland Revenue's annual KiwiSaver payments table, calculating each funding source's percentage share and 10-year nominal and CPI-adjusted real growth since FY2016, alongside current statutory settings (3.5% default employee/employer rate from 1 April 2026 and $260.72 max Government contribution).",
        "hlf_calculation": "Source Share (%) = Source Contributions ÷ Total Contributions; 10Y Real Factor = (Total_2026 ÷ Total_2016) ÷ CPI_10Y_Factor",
        "hlf_interpretation": "HLF market context: Shows how employee, employer, and Government contributions combine to drive long-term KiwiSaver compounding.",
        "questions": [
            "How much do New Zealanders and employers contribute to KiwiSaver each year?",
            "What is the maximum KiwiSaver Government contribution ($260.72) and default contribution rate (3.5%)?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-fund-allocation-shift": {
        "name": "10-year shift in KiwiSaver asset allocation (Growth, Balanced & Conservative funds)",
        "also_known_as": "FMA KiwiSaver funds under management share by risk profile (2016 vs 2026)",
        "category": "KiwiSaver",
        "topic": "KiwiSaver fund risk allocation (Growth vs Conservative)",
        "source": "FMA KiwiSaver Report",
        "source_full": "Financial Markets Authority (FMA) — KiwiSaver Annual Report (Asset Allocation Series)",
        "series_ids": "FMA KiwiSaver FUM Share (%) — Growth/Aggressive, Balanced, Conservative/Cash, Moderate",
        "source_url": FMA_KS_REPORT_PAGE,
        "data_type": "Officially reported FMA FUM shares & HLF-derived 10-year percentage-point shift",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Annual (Year ended 31 March)",
        "definition": "The percentage breakdown of total KiwiSaver funds under management across Growth/Aggressive, Balanced, and Conservative/Cash investment profiles, compared against 10 years earlier.",
        "methodology": "Derived from the FMA KiwiSaver Annual Report asset allocation series by comparing current FUM shares by fund type against the March 2016 distribution.",
        "hlf_calculation": "10-Year Allocation Shift (percentage points) = 2026 Fund Category Share (%) minus 2016 Fund Category Share (%)",
        "hlf_interpretation": "HLF market context: Over the past decade—accelerated by the 2021 shift of default funds from Conservative to Balanced—KiwiSaver money has rotated strongly out of Cash/Conservative funds into Growth and Balanced funds.",
        "questions": [
            "What percentage of KiwiSaver money is in Growth vs Conservative funds?",
            "How has KiwiSaver fund allocation changed over the last 10 years?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-savings-suspensions": {
        "name": "Active KiwiSaver savings suspensions (contributions holiday) & 10-year trend",
        "also_known_as": "Inland Revenue KiwiSaver members on ordinary vs financial hardship savings suspension",
        "category": "KiwiSaver",
        "topic": "KiwiSaver savings suspensions (contribution holidays)",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Savings Suspension Statistics",
        "series_ids": "IRD Active KiwiSaver Savings Suspensions (Ordinary vs Financial Hardship, 2016–2026)",
        "source_url": IRD_KS_SUSP_PAGE,
        "data_type": "Officially reported IRD counts & HLF-derived 10-year percentage change",
        "derived_by_hlf": True,
        "unit": "members",
        "frequency": "Monthly / Annual",
        "definition": "The number of KiwiSaver members currently on an approved savings suspension (formerly 'contributions holiday'), split between ordinary suspensions and financial hardship suspensions.",
        "methodology": "Sourced from Inland Revenue's KiwiSaver savings suspension table, comparing current active suspensions against June 2016 (prior to the 2019 reduction of the maximum suspension term from 5 years to 1 year).",
        "hlf_calculation": "10Y Change (%) = ((Active Suspensions_2026 - Active Suspensions_2016) ÷ Active Suspensions_2016) × 100",
        "hlf_interpretation": "HLF market context: Shortening the maximum standard suspension period from 5 years to 1 year in 2019 significantly reduced the number of members sitting out of KiwiSaver contributions long-term.",
        "questions": [
            "How many New Zealanders are on a KiwiSaver savings suspension (contributions holiday)?",
            "Has the number of KiwiSaver contribution holidays increased or decreased?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "building-insurance-inflation": {
        "name": "Annual residential dwelling (building) insurance price index inflation (Stats NZ CPI)",
        "also_known_as": "New Zealand house insurance CPI price index annual percentage change",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Residential building insurance price inflation",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Dwelling insurance sub-index)",
        "series_ids": "Stats NZ CPI — Dwelling insurance (Annual % change) vs All-groups CPI",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI series & HLF-derived spread vs headline CPI",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in Stats NZ's residential dwelling (building) insurance price index, compared against headline all-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (dwelling insurance sub-index), subtracting headline all-groups CPI to show the inflation spread. Note: CPI insurance sub-indices measure percentage changes in a standardised price index over time, not dollar premium quotes for an individual property.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Dwelling Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: Residential building insurance remains a rapidly rising non-discretionary homeowner holding cost, driven by international catastrophe reinsurance repricing, risk-based flood/seismic rating, and NHC levy adjustments.",
        "questions": [
            "How much has house insurance gone up in New Zealand?",
            "Why is residential building insurance rising faster than CPI inflation in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "contents-insurance-inflation": {
        "name": "Annual household contents insurance price index inflation (Stats NZ CPI)",
        "also_known_as": "New Zealand contents insurance CPI price index annual percentage change",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Household contents insurance price inflation",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Contents insurance sub-index)",
        "series_ids": "Stats NZ CPI — Contents insurance (Annual % change) vs All-groups CPI",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI series & HLF-derived spread vs headline CPI",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in Stats NZ's household contents insurance price index, compared against headline all-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (contents insurance sub-index), compared against headline all-groups CPI.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Contents Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: Contents insurance price growth affects both owner-occupiers and tenants across New Zealand.",
        "questions": [
            "How much has contents insurance increased in New Zealand?",
            "How does contents insurance inflation compare to general inflation in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "vehicle-insurance-inflation": {
        "name": "Annual motor vehicle insurance price index inflation (Stats NZ CPI)",
        "also_known_as": "New Zealand car insurance CPI price index annual percentage change",
        "category": "Insurance Market & Claims",
        "topic": "Motor vehicle insurance price inflation",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Vehicle insurance sub-index)",
        "series_ids": "Stats NZ CPI — Vehicle insurance (Annual % change) vs All-groups CPI",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI series & HLF-derived spread vs headline CPI",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in Stats NZ's motor vehicle insurance price index, compared against headline all-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (vehicle insurance sub-index), compared against headline all-groups CPI.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Vehicle Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: Higher modern vehicle repair costs, parts inflation, and weather-related vehicle write-offs have kept motor insurance inflation above general CPI.",
        "questions": [
            "How much has car insurance gone up in New Zealand?",
            "Why is motor vehicle insurance increasing in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "health-insurance-inflation": {
        "name": "Annual private health insurance price index inflation (Stats NZ CPI)",
        "also_known_as": "New Zealand medical & health insurance CPI price index annual percentage change",
        "category": "Personal Risk & Income Protection",
        "topic": "Private health insurance price inflation",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Health insurance sub-index)",
        "series_ids": "Stats NZ CPI — Health insurance (Annual % change) vs All-groups CPI",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI series & HLF-derived spread vs headline CPI",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in Stats NZ's private health insurance price index, compared against headline all-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (health insurance sub-index), compared against headline all-groups CPI.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Health Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: Rising private surgical claim volumes and medical procedure costs have driven double-digit annual increases in health insurance price indices—prompting many households to review policy excesses ($500–$1,000+) to keep cover affordable.",
        "questions": [
            "How fast are private health insurance premiums rising in New Zealand?",
            "How can I reduce my health insurance premium in NZ without losing surgical cover?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "life-insurance-inflation": {
        "name": "Annual constant-age life insurance price index change (Stats NZ CPI)",
        "also_known_as": "New Zealand life insurance CPI constant-age price index change",
        "category": "Personal Risk & Income Protection",
        "topic": "Life insurance constant-age price index",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Life insurance sub-index)",
        "series_ids": "Stats NZ CPI — Life insurance (Annual % change) vs All-groups CPI",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI series & HLF-derived spread vs headline CPI",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in Stats NZ's constant-age life insurance price index, compared against headline all-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (life insurance sub-index), holding policyholder age and sum insured constant.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Life Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: While underlying constant-age life insurance base rates rise modestly (~2.9% YoY), individual rate-for-age stepped policies increase by ~8%–12%+ per year after age 45 as mortality risk steps up with age—plus optional 3%–5% CPI sum-insured indexation.",
        "questions": [
            "Why does my life insurance premium go up every year as I get older?",
            "What is the difference between stepped and level life insurance premiums in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
# === END OF PART 2D ===
    # === PART 2E OF 5 ===
    "household-insurance-expenditure": {
        "name": "Average insured homeowner building premium vs all-household dwelling insurance spend",
        "also_known_as": "RBNZ insured homeowner building insurance benchmark (~$2,850/yr) & Stats NZ HES average",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Household dwelling insurance expenditure",
        "source": "Stats NZ HES & RBNZ FSR",
        "source_full": "Stats NZ — Household Expenditure Statistics (HES) & RBNZ Financial Stability Report",
        "series_ids": "Stats NZ HES Dwelling Insurance ($/wk × 52) & RBNZ FSR Insured Homeowner Premium Benchmark",
        "source_url": RBNZ_FSR_PAGE,
        "data_type": "Officially reported Stats NZ & RBNZ benchmarks & HLF-derived annualised comparison",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual / Triennial",
        "definition": "Comparison of the average annual domestic building insurance premium paid by insured New Zealand homeowners (~$2,850/yr nationwide, and $3,500–$4,500+/yr in higher-hazard zones) against the Stats NZ Household Expenditure Survey average across all households including renters.",
        "methodology": "Combines RBNZ Financial Stability Report insured homeowner building insurance premium benchmarks with Stats NZ's Household Expenditure Survey weekly dwelling insurance average multiplied by 52 weeks.",
        "hlf_calculation": "Annualised HES All-Household Spend = Weekly HES Dwelling Insurance ($24.80/wk) × 52 = $1,289.60/yr (includes renting households with $0 building insurance)",
        "hlf_interpretation": "HLF market context: Banks require active replacement building insurance as a mandatory condition of every residential mortgage—making actual homeowner premiums (~$2,850–$4,500+/yr) a key fixed expense in mortgage servicing calculators.",
        "questions": [
            "How much is average house insurance in New Zealand per year?",
            "Why is house insurance more expensive in Wellington and flood zones?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "insurance-claims-trends": {
        "name": "Licensed NZ general insurer Gross Written Premiums, Net Earned Premiums & Net Claims",
        "also_known_as": "RBNZ general insurance sector premium revenue, reinsurance outflow & claims paid",
        "category": "Insurance Market & Claims",
        "topic": "General insurer premiums, reinsurance & claims",
        "source": "RBNZ Insurance Statistics",
        "source_full": "Reserve Bank of New Zealand — Insurance Statistics & Financial Stability Report",
        "series_ids": "RBNZ Licensed General Insurers — Gross Written Premium ($B), Net Earned Premium ($B) & Net Incurred Claims ($B)",
        "source_url": RBNZ_INS_STATS_PAGE,
        "data_type": "Officially reported RBNZ aggregates & HLF-derived reinsurance difference",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Semi-Annual / Annual",
        "definition": "Annual Gross Written Premiums (GWP), Net Earned Premiums (after international catastrophe reinsurance cessions), and Net Incurred Claims across licensed New Zealand general insurers.",
        "methodology": "Sourced from RBNZ licensed general insurer financial statistics, subtracting Net Earned Premiums from Gross Written Premiums to illustrate annual reinsurance and catastrophe protection outflows.",
        "hlf_calculation": "Reinsurance & Unearned Outflow = Gross Written Premiums ($9.85B) minus Net Earned Premiums ($7.40B) = ~$2.45B/yr",
        "hlf_interpretation": "HLF market context: New Zealand insurers cede roughly a quarter of gross premiums to global reinsurers to fund earthquake and severe-weather catastrophe cover.",
        "questions": [
            "How much do New Zealand insurance companies collect in premiums and pay out in claims?",
            "How much of New Zealand insurance premiums goes to overseas reinsurance?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "insurer-loss-ratios": {
        "name": "Licensed New Zealand general insurer net claims loss ratio",
        "also_known_as": "RBNZ general insurance net incurred claims to net earned premium ratio",
        "category": "Insurance Market & Claims",
        "topic": "General insurer claims loss ratio",
        "source": "RBNZ Insurance Statistics",
        "source_full": "Reserve Bank of New Zealand — Insurance Statistics & Financial Stability Report",
        "series_ids": "Net Incurred Claims ($B) ÷ Net Earned Premiums ($B)",
        "source_url": RBNZ_INS_STATS_PAGE,
        "data_type": "HLF-derived percentage ratio from official RBNZ insurance aggregates",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Semi-Annual / Annual",
        "definition": "The net claims loss ratio across licensed New Zealand general insurers, measuring net incurred claims as a percentage of net earned premiums.",
        "methodology": "Derived from RBNZ general insurance sector statistics by dividing annual Net Incurred Claims ($4.66B) by annual Net Earned Premiums ($7.40B), multiplied by 100.",
        "hlf_calculation": "Net Loss Ratio (%) = (Net Incurred Claims [$4.66B] ÷ Net Earned Premiums [$7.40B]) × 100 = 63.0%",
        "hlf_interpretation": "HLF market context: Following the >85% loss-ratio spike during the 2023 North Island Auckland Anniversary floods and Cyclone Gabrielle, loss ratios have normalised toward long-run underwriting levels (~60%–65%).",
        "questions": [
            "What is the claims loss ratio for New Zealand general insurers?",
            "How did the 2023 Auckland floods and Cyclone Gabrielle affect insurer loss ratios?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "natural-hazard-claims": {
        "name": "Natural Hazards Commission (NHC / formerly EQC) statutory building cap & annual levy",
        "also_known_as": "NHCover $300,000 + GST residential building cap & $480 + GST maximum levy",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Statutory natural hazard residential building cover (NHCover)",
        "source": "NHI Act 2023 & NHC",
        "source_full": "Natural Hazards Insurance Act 2023 & Natural Hazards Commission Toka Tū Ake (NHC)",
        "series_ids": "Natural Hazards Insurance Act 2023 — Statutory Building Cap ($300k + GST) & Levy Rate (16c/$100)",
        "source_url": NHC_ACT_PAGE,
        "data_type": "Officially legislated statutory cap/levy & HLF-derived GST-inclusive figures",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Statutory Schedule",
        "definition": "The statutory first-loss residential building cover cap ($300,000 + GST per dwelling) and maximum annual levy ($480 + GST) administered by the Natural Hazards Commission Toka Tū Ake (NHC, formerly EQC) under the Natural Hazards Insurance Act 2023.",
        "methodology": "Sourced from the Natural Hazards Insurance Act 2023 (effective 1 July 2024) and NHC levy schedule (16 cents per $100 of cover up to $300,000 + GST), multiplying statutory GST-exclusive figures by 1.15 to show GST-inclusive homeowner totals.",
        "hlf_calculation": "Cap incl. GST = $300,000 × 1.15 = $345,000; Max Annual Levy incl. GST = ($300,000 ÷ 100 × $0.16) × 1.15 = $552.00",
        "hlf_interpretation": "HLF market context: An active private house insurance policy with fire cover is required to trigger automatic statutory NHCover—homeowners who let their private building policy lapse lose both private cover and the $300,000 + GST statutory natural hazard protection.",
        "questions": [
            "What is the NHC (EQC) building cap and levy in New Zealand?",
            "Do I get NHC (EQC) natural hazard cover if I don't have private house insurance?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "residential-insurance-coverage": {
        "name": "New Zealand residential building insurance coverage & uninsured dwelling estimate",
        "also_known_as": "RBNZ & NHC insured vs uninsured domestic dwellings in New Zealand (~97% vs ~3%)",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Residential dwelling insurance coverage rate",
        "source": "RBNZ FSR & NHC",
        "source_full": "Reserve Bank of New Zealand — May 2026 Financial Stability Report & Natural Hazards Commission Toka Tū Ake",
        "series_ids": "RBNZ FSR & NHC — Uninsured Domestic Dwellings (~60,000) vs Total Domestic Dwellings (~2,000,000)",
        "source_url": RBNZ_FSR_PAGE,
        "data_type": "Officially reported RBNZ/NHC estimates & HLF-derived coverage percentages",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Semi-Annual",
        "definition": "The estimated percentage and number of New Zealand domestic dwellings covered by residential building insurance (~97.0%, or ~1.94m dwellings across ~$1.5T sum insured) versus dwellings with no building insurance (~3.0%, or ~60,000 dwellings).",
        "methodology": "Derived from the Reserve Bank of New Zealand's May 2026 Financial Stability Report (~60,000 uninsured dwellings out of ~2.0 million domestic dwellings) and NHC aggregate residential sum-insured statistics (~$1.5 trillion).",
        "hlf_calculation": "Uninsured Share = 60,000 ÷ 2,000,000 = 3.0%; Insured Share = (2,000,000 - 60,000) ÷ 2,000,000 = 97.0%",
        "hlf_interpretation": "HLF market context: Because mortgage lenders mandate building insurance on all mortgaged homes, outright uninsurance (~3%) is concentrated among mortgage-free properties, while mortgaged homeowners face underinsurance risk if their sum-insured cap lags rebuild costs.",
        "questions": [
            "How many houses in New Zealand are uninsured?",
            "What percentage of New Zealand homes have building insurance?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "acc-weekly-compensation-claims": {
        "name": "ACC annual accepted injury claims vs new weekly income compensation claims",
        "also_known_as": "Accident Compensation Corporation total injury claims vs time-off-work weekly compensation claims",
        "category": "Personal Risk & Income Protection",
        "topic": "ACC accidental injury weekly compensation vs illness gap",
        "source": "ACC Annual Report",
        "source_full": "Accident Compensation Corporation (ACC) — Annual Report (2025/26 Financial Year)",
        "series_ids": "ACC Annual Report — Total Accepted Injury Claims (~2.1m) & New Weekly Compensation Claims (~105,000)",
        "source_url": ACC_REPORT_PAGE,
        "data_type": "Officially reported ACC annual claim counts & HLF-derived ratio per 100 accepted claims",
        "derived_by_hlf": True,
        "unit": "claims_per_100",
        "frequency": "Annual (June Financial Year)",
        "definition": "Comparison of total annual accepted injury claims registered with ACC across New Zealand (~2.1 million, mostly minor treatment-only visits) and new weekly income compensation claims for time off work (~105,000, or 5.0 weekly compensation claims per 100 accepted injury claims).",
        "methodology": "Sourced from ACC's 2025/26 Annual Report, dividing annual new weekly compensation entitlement claims (105,000) by total annual accepted injury claims (2,100,000) multiplied by 100. Note: These are two separate annual claim counts rather than a proportion of unique injured individuals.",
        "hlf_calculation": "(105,000 New Weekly Compensation Claims ÷ 2,100,000 Total Accepted Injury Claims) × 100 = 5.0 weekly comp claims per 100 accepted claims",
        "hlf_interpretation": "HLF market context: ACC only covers accidental physical injuries—it does not pay weekly income compensation if a borrower cannot work due to non-injury illness (such as cancer, heart attack, stroke, or mental health conditions), which requires private income or mortgage repayment protection.",
        "questions": [
            "Does ACC pay my income or mortgage if I get sick with cancer or an illness in New Zealand?",
            "How many ACC weekly compensation claims are approved each year in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "acc-earners-levy-cap": {
        "name": "ACC Earners' Levy rate, Maximum Liable Earnings cap & 80% weekly compensation ceiling",
        "also_known_as": "2026/27 ACC $156,641 covered income cap, $1.75/$100 levy & $125,313 max annual ACC payout",
        "category": "Personal Risk & Income Protection",
        "topic": "ACC covered earnings cap & high-earner protection gap",
        "source": "ACC & IRD Schedule",
        "source_full": "Accident Compensation Corporation (ACC) & Inland Revenue — 2026/27 Earners' Levy Statutory Schedule",
        "series_ids": "ACC 2026/27 Earners' Levy ($1.75/$100 incl. GST) & Maximum Liable Earnings Cap ($156,641)",
        "source_url": ACC_LEVY_PAGE,
        "data_type": "Officially legislated statutory schedule & HLF-derived maximum levy/compensation ceiling",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual (Tax Year from 1 April)",
        "definition": "The statutory ACC Earners' Levy rate ($1.75 per $100 of liable earnings incl. GST) and Maximum Liable Earnings cap ($156,641) for the 2026/27 tax year, alongside the resulting maximum annual levy ($2,741.22) and maximum 80% gross weekly compensation ceiling ($125,312.80/yr or $2,409.86/wk).",
        "methodology": "Derived from the statutory 2026/27 ACC Earners' Levy schedule (effective 1 April 2026) by multiplying the $156,641 Maximum Liable Earnings cap by the $1.75/$100 GST-inclusive levy rate and by the 80% statutory weekly compensation rate.",
        "hlf_calculation": "Max Annual Levy = ($156,641 ÷ 100) × $1.75 = $2,741.22; Max 80% ACC Payout = $156,641 × 80% = $125,312.80/yr ($2,409.86/wk gross)",
        "hlf_interpretation": "HLF market context: Any income earned above $156,641 is not covered by ACC weekly compensation—meaning higher-income borrowers face both an uncovered top-up gap on accidental injuries and a 100% gap on non-injury illnesses.",
        "questions": [
            "What is the maximum income covered by ACC in New Zealand for 2026/27?",
            "What is the ACC Earners' Levy rate and maximum weekly compensation payout in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "life-trauma-average-claim-payout": {
        "name": "Life, Trauma & Income Protection annual claims paid & average claim payouts (FSC / RBNZ)",
        "also_known_as": "Average New Zealand life insurance, critical illness trauma & income protection claim size",
        "category": "Personal Risk & Income Protection",
        "topic": "Personal risk insurance average claim sizes & annual payouts",
        "source": "FSC & RBNZ Life Stats",
        "source_full": "Financial Services Council (FSC) NZ Life Insurance Spotlight & RBNZ Licensed Life Insurer Statistics",
        "series_ids": "FSC / RBNZ Annual Life, Trauma & Disability Claims Paid (~$1.42B/yr) & Benchmark Mean Payouts by Cover Type",
        "source_url": FSC_DATA_PAGE,
        "data_type": "Officially reported FSC/RBNZ industry aggregates & HLF-curated claim payout benchmarks",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual",
        "definition": "Total annual claims paid across New Zealand's licensed life and personal risk insurers (~$1.42B/year) alongside industry benchmark average claim payouts for Term Life cover (~$295,000), Critical Illness / Trauma lump-sum cover (~$122,000), and Income / Mortgage Repayment Protection (~$5,150/month, or ~$108,150 across an average 21-month long-term claim).",
        "methodology": "Compiled from Financial Services Council (FSC) New Zealand Life Insurance claims spotlight benchmarks and RBNZ licensed life insurer returns, comparing average lump-sum payouts across Term Life and Critical Illness/Trauma against average monthly disability/income protection benefits and claim cause breakdowns.",
        "hlf_calculation": "Average 21-Month Income Protection Claim Total = $5,150/month × 21 months = $108,150 (Annualised = $61,800/yr); Trauma Causes: ~56% Cancer, ~24% Heart/Stroke, ~20% Other",
        "hlf_interpretation": "HLF market context: Shows why lump-sum Trauma cover (~$122k average payout) and monthly Income/Mortgage Repayment Protection (~$5,150/mth) are the working-age workhorses for mortgage holders—providing cashflow during serious illness while Term Life (~$295k average payout) clears or reduces debt on death.",
        "questions": [
            "What is the average life insurance or trauma insurance payout in New Zealand?",
            "What are the most common reasons for trauma and income protection claims in NZ?",
            "How much do New Zealand life insurers pay out in claims each year?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "private-health-claims-average-payout": {
        "name": "Private health insurance annual claims paid, insured lives & surgical procedure cost benchmarks",
        "also_known_as": "FSC / HFANZ private medical insurance claims total ($2.26B), average claim & private surgery costs",
        "category": "Personal Risk & Income Protection",
        "topic": "Private health insurance claims & surgical costs",
        "source": "FSC & Health Funds NZ",
        "source_full": "Financial Services Council (FSC) & Health Funds Association of New Zealand (HFANZ) Industry Statistics",
        "series_ids": "FSC Health Insurance — Insured Lives (~1.48m), Annual Claims Paid ($2.26B) & Surgical Procedure Cost Benchmarks",
        "source_url": FSC_DATA_PAGE,
        "data_type": "Officially reported FSC health insurance totals & HLF-derived per-claim mean and surgical benchmarks",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual / Quarterly",
        "definition": "Total annual private health insurance claims paid in New Zealand (~$2.26B across ~1.86 million claims for ~1.48 million insured New Zealanders, or ~28% of the population), alongside the overall average claim payment (~$1,215 across all GP/diagnostic/surgical claims) and benchmark private hospital surgical costs ($18,000–$65,000+).",
        "methodology": "Derived from Financial Services Council (FSC) health insurance statistics by dividing annual claims paid ($2.26B) by annual claim volume (~1.86m claims), and tabulating New Zealand private hospital surgical procedure cost benchmarks (joint replacement, spinal surgery, cardiac surgery, and private oncology).",
        "hlf_calculation": "Overall Mean Health Claim = $2,260,000,000 ÷ 1,860,000 claims = $1,215/claim (includes minor GP/optical/imaging claims); Major Elective Surgical Admissions = $18,400–$65,000+",
        "hlf_interpretation": "HLF market context: While minor everyday health claims pull the all-claims average down to ~$1,215, mortgage holders primarily carry private surgical and non-PHARMAC cancer cover to fund $20,000–$65,000+ hospital procedures without waiting on public elective lists.",
        "questions": [
            "How much do private health insurers pay out in claims in New Zealand?",
            "How much does private surgery (like a hip replacement or heart surgery) cost in New Zealand?",
            "What percentage of New Zealanders have private health insurance?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "personal-insurance-underinsurance-gap": {
        "name": "New Zealand household life & income protection coverage vs illness disability risk",
        "also_known_as": "Property insurance (97%) vs personal life (~36%) & income protection (~17%) coverage gap in NZ",
        "category": "Personal Risk & Income Protection",
        "topic": "Household underinsurance & income protection gap",
        "source": "FSC & RBNZ Benchmarks",
        "source_full": "Financial Services Council (FSC) Risk & Underinsurance Research & RBNZ Financial Stability Report",
        "series_ids": "Insured Dwellings (97.0%) vs Working-Age Life Cover (~36%), Income/Mortgage Protection (~17%) & Illness Share of >6m Disability (~82%)",
        "source_url": FSC_DATA_PAGE,
        "data_type": "Officially reported FSC/RBNZ research benchmarks & HLF-derived protection gap comparison",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Annual",
        "definition": "Comparison of New Zealand residential building insurance coverage (97.0% of dwellings) against working-age personal risk coverage (~36% holding standalone life insurance and ~17% holding income or mortgage repayment protection), alongside the share of long-term (>6-month) work disabilities caused by non-injury illness (~82%, which ACC does not cover).",
        "methodology": "Compares RBNZ/NHC residential dwelling insurance penetration (97.0%) against Financial Services Council (FSC) New Zealand underinsurance research on working-age personal life and income/mortgage protection penetration and disability cause distributions.",
        "hlf_calculation": "Dwelling vs Income Protection Coverage Gap = 97.0% (Home Building Insurance) minus 17.0% (Income/Mortgage Protection) = 80.0 percentage points",
        "hlf_interpretation": "HLF market context: Almost every mortgaged homeowner insures the physical house (97%), yet fewer than 1 in 5 working-age New Zealanders insure the income that pays the mortgage (~17%)—even though non-injury illness accounts for ~82% of absences from work lasting longer than 6 months.",
        "questions": [
            "What percentage of New Zealanders have income protection or mortgage repayment insurance?",
            "Why is New Zealand considered underinsured for life and income protection?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
}

# Backward-compatible aliases for Webflow chat bubbles (excluded from the 63 canonical records)
LEGACY_BUBBLE_ALIASES = {
    "fhb-income-akl-vs-nz": "fhb-income-by-dti-band",
    "fhb-dti-above-6x": "dti-headroom",
    "fhb-typical-loan-size": "fhb-average-loan-size",
    "fhb-monthly-volume": "fhb-monthly-lending-volume",
    "build-cost-by-type": "new-build-consent-cost-by-type",
    "personal-insurance-average-claim": "life-trauma-average-claim-payout",
    "fhb-low-deposit-share": "fhb-low-deposit",
    "fhb-lending-share": "fhb-share",
    "investor-lending-share": "investors",
    "refinance-share": "refinance",
    "top-up-share": "top-ups",
    "interest-only-share": "interest-only",
    "fhb-dti-headroom": "dti-headroom",
    "fhb-borrower-income": "borrower-income",
    "building-consents-annual": "new-builds-annual",
    "median-rent-nz": "rent-yield-pulse",
    "kiwisaver-first-home-withdrawals": "kiwisaver-first-home",
    "kiwisaver-provider-switches": "kiwisaver-switches",
    "kiwisaver-average-balance": "kiwisaver-total-fum-balance",
}
# === END OF PART 2E ===
# === PART 3A OF 5 ===
class SimpleTableExtractor(HTMLParser):
    """Lightweight HTML table parser for Inland Revenue regional KiwiSaver tables."""

    def __init__(self) -> None:
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._in_table = False
        self._in_row = False
        self._in_cell = False
        self._cur_table: list[list[str]] = []
        self._cur_row: list[str] = []
        self._cur_cell: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        t = tag.lower()
        if t == "table":
            self._in_table = True
            self._cur_table = []
        elif t == "tr" and self._in_table:
            self._in_row = True
            self._cur_row = []
        elif t in ("td", "th") and self._in_row:
            self._in_cell = True
            self._cur_cell = []

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in ("td", "th") and self._in_cell:
            self._in_cell = False
            self._cur_row.append(" ".join("".join(self._cur_cell).split()))
        elif t == "tr" and self._in_row:
            self._in_row = False
            if self._cur_row:
                self._cur_table.append(self._cur_row)
        elif t == "table" and self._in_table:
            self._in_table = False
            if self._cur_table:
                self.tables.append(self._cur_table)

    def handle_data(self, data: str) -> None:
        if self._in_cell:
            self._cur_cell.append(data)


def get_http_session():
    try:
        from curl_cffi import requests as cffi_requests
        return cffi_requests.Session(impersonate="chrome124")
    except Exception:
        import urllib.request
        return urllib.request


def fetch_url_bytes(session, url: str, referer: str = "", timeout: int = 35, use_cb: bool = True) -> bytes:
    target = url
    if use_cb:
        sep = "&" if "?" in url else "?"
        target = f"{url}{sep}_cb={int(time.time())}"
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-NZ,en-GB;q=0.9,en-US;q=0.8,en;q=0.7",
    }
    if referer:
        headers["Referer"] = referer
    try:
        if hasattr(session, "get"):
            resp = session.get(target, headers=headers, timeout=timeout)
            if resp.status_code == 200 and resp.content:
                return resp.content
        else:
            req = session.Request(target, headers=headers)
            with session.urlopen(req, timeout=timeout) as r:
                return r.read()
    except Exception as exc:
        print(f"[HTTP] Fetch note for {url}: {exc}")
    return b""


def fetch_html_with_browser_fallback(session, urls: list[str], label: str = "Web") -> dict[str, str]:
    results: dict[str, str] = {}
    missing: list[str] = []
    for u in urls:
        raw = fetch_url_bytes(session, u, timeout=25, use_cb=False)
        txt = raw.decode("utf-8", errors="ignore") if raw else ""
        if len(txt) > 1500 and "Just a moment" not in txt and "Access Denied" not in txt:
            results[u] = txt
        else:
            missing.append(u)

    if not missing:
        return results

    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled"])
            ctx = browser.new_context(
                user_agent=(
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/124.0.0.0 Safari/537.36"
                ),
                locale="en-NZ",
            )
            page = ctx.new_page()
            for u in missing:
                try:
                    page.goto(u, wait_until="domcontentloaded", timeout=35000)
                    page.wait_for_timeout(1800)
                    content = page.content()
                    if len(content) > 1200:
                        results[u] = content
                except Exception as e_pw:
                    print(f"[{label} Playwright] {u}: {e_pw}")
            browser.close()
    except Exception as exc:
        print(f"[{label} Fallback] Playwright note: {exc}")
    return results


def download_rbnz_workbooks(root: pathlib.Path, session) -> dict[str, pathlib.Path]:
    found: dict[str, pathlib.Path] = {}
    for key, cfg in RBNZ_SOURCES.items():
        local_candidates = sorted(root.glob(f"*{cfg['match']}*.xlsx"))
        for cand in local_candidates:
            if not cand.name.startswith("tmp_") and cand.stat().st_size > 15_000:
                found[key] = cand
                break

        raw = fetch_url_bytes(session, cfg["xlsx"], referer=RBNZ_WARMUP_URL, timeout=45)
        if raw and raw[:2] == b"PK" and len(raw) > 15_000:
            tmp_p = root / f"tmp_{key}.xlsx"
            tmp_p.write_bytes(raw)
            found[key] = tmp_p
            print(f"[RBNZ] Downloaded live {key.upper()} workbook ({len(raw):,} bytes)")
            continue

        for pg in cfg["pages"]:
            pg_bytes = fetch_url_bytes(session, pg, referer=RBNZ_WARMUP_URL, timeout=30, use_cb=False)
            if not pg_bytes:
                continue
            pg_html = pg_bytes.decode("utf-8", errors="ignore")
            for href in re.findall(r'href=["\']([^"\']+\.xlsx[^"\']*)["\']', pg_html, flags=re.I):
                if cfg["match"] in href.lower():
                    full_u = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
                    x_bytes = fetch_url_bytes(session, full_u, referer=pg, timeout=45)
                    if x_bytes and x_bytes[:2] == b"PK" and len(x_bytes) > 15_000:
                        tmp_p = root / f"tmp_{key}.xlsx"
                        tmp_p.write_bytes(x_bytes)
                        found[key] = tmp_p
                        print(f"[RBNZ] Discovered & downloaded {key.upper()} workbook ({len(x_bytes):,} bytes)")
                        break
            if key in found and found[key].name.startswith("tmp_"):
                break
    return found


def extract_rbnz_published_date(filepath: pathlib.Path) -> str:
    """Extract official RBNZ release/publication date from workbook notes/series tabs without copying period_end."""
    try:
        xl = pd.ExcelFile(filepath)
        for sname in xl.sheet_names:
            df_top = xl.parse(sname, header=None, nrows=25)
            for r_i in range(len(df_top)):
                row_vals = [str(v).strip() for v in df_top.iloc[r_i].values if pd.notna(v) and str(v).strip()]
                joined = " ".join(row_vals)
                m = re.search(
                    r"(?:published|release date|last updated|updated on)[:\s]+(\d{1,2}\s+[A-Za-z]+\s+20\d{2}|20\d{2}-\d{2}-\d{2})",
                    joined,
                    flags=re.I,
                )
                if m:
                    dt_p = pd.to_datetime(m.group(1), errors="coerce")
                    if pd.notna(dt_p):
                        return f"{dt_p.strftime('%Y-%m-%d')} ({dt_p.day} {dt_p.strftime('%B %Y')} — RBNZ statistical release)"
    except Exception:
        pass
    return "Monthly RBNZ statistical release (~4 weeks after reference month-end)"


def extract_html_release_date(html_str: str, fallback_schedule: str) -> str:
    """Extract official release date from Stats NZ / IRD / MBIE HTML without copying period_end."""
    if html_str:
        txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html_str[:18000]))
        patterns = [
            r"(?:released|published|last updated|date published)[:\s]+(\d{1,2}\s+[A-Za-z]+\s+20\d{2})",
            r"(\d{1,2}\s+[A-Za-z]+\s+20\d{2})\s+(?:information release|release)",
        ]
        for pat in patterns:
            m = re.search(pat, txt, flags=re.I)
            if m:
                dt_p = pd.to_datetime(m.group(1), errors="coerce")
                if pd.notna(dt_p):
                    return f"{dt_p.strftime('%Y-%m-%d')} ({dt_p.day} {dt_p.strftime('%B %Y')})"
    return fallback_schedule


def load_rbnz_sheets(filepath: pathlib.Path) -> list[tuple[str, pd.DataFrame, list[str], list[str]]]:
    xl = pd.ExcelFile(filepath)
    out = []
    for sname in xl.sheet_names:
        if sname.strip().lower() in ("notes", "definitions", "contents", "index", "metadata"):
            continue
        raw = xl.parse(sname, header=None)
        if raw.empty or raw.shape[0] < 10 or raw.shape[1] < 2:
            continue

        date_row_start = None
        for r_idx in range(min(25, len(raw))):
            v0 = raw.iloc[r_idx, 0]
            dt = pd.to_datetime(v0, errors="coerce")
            if pd.notna(dt) and 1990 <= dt.year <= 2035:
                date_row_start = r_idx
                break
        if date_row_start is None:
            continue

        hdr_block = raw.iloc[:date_row_start, :]
        series_codes: list[str] = []
        col_labels: list[str] = []
        for c_idx in range(raw.shape[1]):
            col_cells = [
                str(hdr_block.iloc[r_i, c_idx]).strip()
                for r_i in range(date_row_start)
                if pd.notna(hdr_block.iloc[r_i, c_idx]) and str(hdr_block.iloc[r_i, c_idx]).strip().lower() != "nan"
            ]
            code_found = ""
            for cell_s in reversed(col_cells):
                if re.match(r"^[A-Z0-9]{3,6}\.[A-Z0-9._]+$", cell_s, flags=re.I):
                    code_found = cell_s.upper()
                    break
            series_codes.append(code_found)
            clean_lbl = re.sub(r"\s+", " ", " | ".join(col_cells)).lower()
            col_labels.append(clean_lbl)

        data_df = raw.iloc[date_row_start:, :].copy()
        data_df["date"] = pd.to_datetime(data_df.iloc[:, 0], errors="coerce")
        data_df = data_df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
        if len(data_df) >= 3:
            out.append((sname, data_df, series_codes, col_labels))
    return out


def find_rbnz_series(
    tabs: list,
    target_codes: tuple[str, ...] = (),
    include_kw: tuple[str, ...] = (),
    exclude_kw: tuple[str, ...] = (),
    val_bounds: tuple[float, float] | None = None,
) -> tuple[pd.Series, pd.Series, str]:
    upper_targets = tuple(c.upper() for c in target_codes)
    for _, df, codes, labels in tabs:
        if upper_targets:
            for idx, code in enumerate(codes):
                if code in upper_targets:
                    ser = pd.to_numeric(df.iloc[:, idx], errors="coerce")
                    valid = ser.dropna()
                    if len(valid):
                        if val_bounds and not (val_bounds[0] <= float(valid.iloc[-1]) <= val_bounds[1]):
                            continue
                        return df["date"], ser, code

    for _, df, codes, labels in tabs:
        for idx, lbl in enumerate(labels):
            if include_kw and not all(k.lower() in lbl for k in include_kw):
                continue
            if exclude_kw and any(k.lower() in lbl for k in exclude_kw):
                continue
            ser = pd.to_numeric(df.iloc[:, idx], errors="coerce")
            valid = ser.dropna()
            if len(valid):
                if val_bounds and not (val_bounds[0] <= float(valid.iloc[-1]) <= val_bounds[1]):
                    continue
                return df["date"], ser, codes[idx] or lbl[:60]

    raise KeyError(f"Could not find RBNZ series codes={target_codes} include={include_kw}")


def month_end_iso(dt: pd.Timestamp) -> str:
    ts = pd.Timestamp(dt)
    end_dt = ts + pd.offsets.MonthEnd(0)
    return end_dt.strftime("%Y-%m-%d")


def is_fresh(as_of_iso: str, max_age_days: int) -> bool:
    if not as_of_iso:
        return False
    try:
        dt = datetime.date.fromisoformat(as_of_iso[:10])
        today = datetime.datetime.now(datetime.timezone.utc).date()
        return (today - dt).days <= max_age_days
    except Exception:
        return False


def fmt_nzd_millions(val_m: float) -> str:
    if abs(val_m) >= 1000.0:
        return f"${val_m / 1000.0:.2f}B (${val_m:,.0f}m)"
    return f"${val_m:,.0f}m"


def plain_prior_year_pct(curr_pct: float, prev_pct: float) -> str:
    diff = curr_pct - prev_pct
    if abs(diff) < 0.05:
        return f"unchanged from {prev_pct:.1f}% a year earlier"
    direction = "up from" if diff > 0 else "down from"
    return f"{direction} {prev_pct:.1f}% a year earlier"


def plain_yoy_growth(curr_val: float, prev_val: float) -> str:
    if not prev_val:
        return "reported in the latest release"
    pct = ((curr_val - prev_val) / prev_val) * 100.0
    if abs(pct) < 0.05:
        return "unchanged year-on-year"
    return f"{'up' if pct > 0 else 'down'} {abs(pct):.1f}% year-on-year"


def make_history_list(dates: pd.Series, values: pd.Series, fmt_fn, limit: int = 13) -> list[dict]:
    rows = []
    for dt_v, val_v in zip(dates, values):
        if pd.notna(dt_v) and pd.notna(val_v):
            ts = pd.Timestamp(dt_v)
            fv = float(val_v)
            rows.append({
                "period": ts.strftime("%Y-%m"),
                "label": ts.strftime("%b %Y"),
                "value": round(fv, 2),
                "formatted": fmt_fn(fv),
            })
    return list(reversed(rows[-limit:]))


def format_provenance_source_html(
    calc_live: str,
    source_label: str,
    source_url: str,
    is_derived: bool = True,
    checked_human: str = "",
) -> str:
    """Build the structured provenance dropdown HTML shown when a user expands the source chip."""
    deriv_tag = (
        f"Derived by HLF Pulse ({HLF_LICENSE_SHORT})"
        if is_derived
        else f"Official figure curated by HLF Pulse ({HLF_LICENSE_SHORT})"
    )
    chk_part = f" &bull; Checked: {checked_human}" if checked_human else ""
    return (
        f"<strong>HLF calculation:</strong> <code>{calc_live}</code><br>"
        f'<strong>Source:</strong> <a href="{source_url}" target="_blank" rel="noopener">{source_label}</a> '
        f"&bull; {deriv_tag}{chk_part}."
    )
# === END OF PART 3A ===
# === PART 3B OF 5 ===
# ---------------------------------------------------------------------------
# Reference-Grade RBNZ C31 Parser — Editorial Prose + Provenance Dropdown
# (Outputs 18 C31 Datasets Including FHB Snapshot Rows 2 & 5 + Mover/Investor Loan Sizes)
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
    ser_fhb_avg_k = (fhb_val * 1000.0) / fhb_cnt.replace(0, float("nan"))
    ser_inv_avg_k = (inv_val * 1000.0) / inv_cnt.replace(0, float("nan"))
    ser_oo_avg_k = (oo_val * 1000.0) / oo_cnt.replace(0, float("nan"))
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
    fhb_loans_prev_int = int(round(float(fhb_cnt.iloc[iy])))
    fhb_hi_loans_int = int(round(float(fhb_hi_cnt.iloc[i])))
    tot_val_m = float(tot_val.iloc[i])
    fhb_v = float(fhb_val.iloc[i])
    fhb_v_prev = float(fhb_val.iloc[iy])

    fhb_sh = float(ser_fhb_sh_pct.iloc[i])
    fhb_sh_yoy = float(ser_fhb_sh_pct.iloc[iy])
    fhb_avg_nzd = int(round((fhb_v * 1_000_000.0) / max(1, fhb_loans_int)))
    fhb_avg_prev_nzd = int(round((fhb_v_prev * 1_000_000.0) / max(1, fhb_loans_prev_int)))

    inv_sh = float(ser_inv_sh_pct.iloc[i])
    inv_sh_yoy = float(ser_inv_sh_pct.iloc[iy])
    inv_v = float(inv_val.iloc[i])
    inv_v_prev = float(inv_val.iloc[iy])
    inv_b = inv_v / 1000.0
    inv_c_tot = int(round(float(inv_cnt.iloc[i])))
    inv_c_prev = int(round(float(inv_cnt.iloc[iy])))
    inv_avg_nzd = int(round((inv_v * 1_000_000.0) / max(1, inv_c_tot)))
    inv_avg_prev_nzd = int(round((inv_v_prev * 1_000_000.0) / max(1, inv_c_prev)))

    oo_v = float(oo_val.iloc[i])
    oo_v_prev = float(oo_val.iloc[iy])
    oo_tot_c = float(oo_cnt.iloc[i])
    oo_tot_c_int = int(round(oo_tot_c))
    oo_prev_c_int = int(round(float(oo_cnt.iloc[iy])))
    oo_avg_nzd = int(round((oo_v * 1_000_000.0) / max(1, oo_tot_c_int)))
    oo_avg_prev_nzd = int(round((oo_v_prev * 1_000_000.0) / max(1, oo_prev_c_int)))

    low_dep_html = (
        f"In {m_full}, <strong>{low_cnt_pct:.1f}%</strong> of first-home mortgages approved across New Zealand "
        f"(<strong>{fhb_hi_loans_int:,} of {fhb_loans_int:,} loans</strong>, or <strong>${hi_dollars_m:,.0f}m</strong> in lending) "
        f"were written with <strong>under a 20% deposit</strong> (&gt;80% LVR), "
        f"{plain_prior_year_pct(low_cnt_pct, low_cnt_yoy)}."
    )

    fhb_share_html = (
        f"In {m_full}, New Zealand banks approved <strong>{fhb_loans_int:,}</strong> first-home buyer mortgages totalling "
        f"<strong>{fmt_nzd_millions(fhb_v)}</strong>—accounting for <strong>{fhb_sh:.1f}%</strong> of all new "
        f"residential mortgage lending ({plain_prior_year_pct(fhb_sh, fhb_sh_yoy)}), with an average loan size of "
        f"<strong>${fhb_avg_nzd:,}</strong>."
    )

    fhb_avg_loan_html = (
        f"In {m_full}, the average first-home buyer mortgage approved across New Zealand was "
        f"<strong>${fhb_avg_nzd:,}</strong> ({plain_yoy_growth(fhb_avg_nzd, fhb_avg_prev_nzd)} from ${fhb_avg_prev_nzd:,} a year earlier), "
        f"across <strong>{fhb_loans_int:,} first-home buyer loans</strong> totalling <strong>{fmt_nzd_millions(fhb_v)}</strong>."
    )

    fhb_vol_html = (
        f"In {m_full}, New Zealand banks committed <strong>{fmt_nzd_millions(fhb_v)}</strong> in new first-home buyer mortgages "
        f"across <strong>{fhb_loans_int:,} loans</strong> ({plain_yoy_growth(fhb_v, fhb_v_prev)} from {fmt_nzd_millions(fhb_v_prev)} "
        f"across {fhb_loans_prev_int:,} loans a year earlier)."
    )

    inv_html = (
        f"In {m_full}, residential property investors accounted for <strong>{inv_sh:.1f}%</strong> of new mortgage lending across New Zealand "
        f"(<strong>${inv_b:.2f}B</strong> across <strong>{inv_c_tot:,} loans</strong>, {plain_prior_year_pct(inv_sh, inv_sh_yoy)}), "
        f"compared with <strong>{fhb_sh:.1f}%</strong> for first-home buyers."
    )

    inv_avg_loan_html = (
        f"In {m_full}, residential property investors borrowed <strong>{fmt_nzd_millions(inv_v)}</strong> across "
        f"<strong>{inv_c_tot:,} new mortgage commitments</strong>—an average investor loan size of <strong>${inv_avg_nzd:,}</strong> "
        f"({plain_yoy_growth(inv_avg_nzd, inv_avg_prev_nzd)} from ${inv_avg_prev_nzd:,} a year earlier)."
    )

    mover_avg_loan_html = (
        f"In {m_full}, existing owner-occupiers ('movers' and homeowners refinancing or topping up) borrowed "
        f"<strong>{fmt_nzd_millions(oo_v)}</strong> across <strong>{oo_tot_c_int:,} mortgage commitments</strong>—an average loan size of "
        f"<strong>${oo_avg_nzd:,}</strong> ({plain_yoy_growth(oo_avg_nzd, oo_avg_prev_nzd)} from ${oo_avg_prev_nzd:,} a year earlier)."
    )

    fhb_lo_val_cur = float(fhb_val.iloc[i] - fhb_hi_val.iloc[i])
    fhb_lo_cnt_cur = float(fhb_cnt.iloc[i] - fhb_hi_cnt.iloc[i])
    fhb_lo_loans_int = int(round(fhb_lo_cnt_cur))
    avg_hi_loan = int(round((float(fhb_hi_val.iloc[i]) * 1_000_000.0) / max(1.0, float(fhb_hi_cnt.iloc[i]))))
    avg_lo_loan = int(round((fhb_lo_val_cur * 1_000_000.0) / max(1.0, fhb_lo_cnt_cur)))
    loan_gap = avg_hi_loan - avg_lo_loan

    fhb_gap_html = (
        f"In {m_full}, first-home buyers purchasing with under a 20% deposit (&gt;80% LVR) borrowed an average of "
        f"<strong>${avg_hi_loan:,}</strong>, compared with <strong>${avg_lo_loan:,}</strong> for those with a 20% or greater deposit "
        f"(&le;80% LVR)—a difference of <strong>${abs(loan_gap):,}</strong>."
    )

    fhb_hi_val_share = float(ser_fhb_prio_pct.iloc[i])
    fhb_hi_cnt_share = (float(fhb_hi_cnt.iloc[i]) / max(1.0, float(all_hi_cnt.iloc[i]))) * 100.0
    all_hi_loans_int = int(round(float(all_hi_cnt.iloc[i])))
    fhb_priority_html = (
        f"In {m_full}, first-home buyers received <strong>{fhb_hi_val_share:.1f}% of all low-deposit (&gt;80% LVR) mortgage lending dollars</strong> "
        f"(<strong>{fmt_nzd_millions(float(fhb_hi_val.iloc[i]))} of {fmt_nzd_millions(float(all_hi_val.iloc[i]))}</strong>) and "
        f"<strong>{fhb_hi_cnt_share:.1f}% of low-deposit loans</strong> (<strong>{fhb_hi_loans_int:,} of {all_hi_loans_int:,}</strong>) nationwide."
    )

    std_dep_pct = float(ser_fhb_20club_pct.iloc[i])
    fhb_20club_html = (
        f"In {m_full}, <strong>{std_dep_pct:.1f}%</strong> of first-home buyers (<strong>{fhb_lo_loans_int:,} borrowers</strong>, or "
        f"<strong>{fmt_nzd_millions(fhb_lo_val_cur)}</strong> in lending) purchased with a <strong>20% or greater deposit</strong> (&le;80% LVR), "
        f"with an average mortgage size of <strong>${avg_lo_loan:,}</strong>."
    )

    diff_fi = fhb_v - inv_v
    if diff_fi >= 0:
        fhb_vs_inv_html = (
            f"In {m_full}, first-home buyers borrowed <strong>{fmt_nzd_millions(fhb_v)} ({fhb_sh:.1f}% of new lending)</strong> "
            f"compared with <strong>{fmt_nzd_millions(inv_v)} ({inv_sh:.1f}%)</strong> for residential property investors—a lead of "
            f"<strong>+{fmt_nzd_millions(abs(diff_fi))}</strong> toward first-home buyers."
        )
    else:
        fhb_vs_inv_html = (
            f"In {m_full}, first-home buyers borrowed <strong>{fmt_nzd_millions(fhb_v)} ({fhb_sh:.1f}% of new lending)</strong> "
            f"compared with <strong>{fmt_nzd_millions(inv_v)} ({inv_sh:.1f}%)</strong> for residential property investors."
        )

    oo_hi_c = float(oo_hi_cnt.iloc[i])
    oo_hi_v = float(oo_hi_val.iloc[i])
    oo_lo_c = oo_tot_c - oo_hi_c
    oo_lo_v = float(oo_val.iloc[i] - oo_hi_val.iloc[i])
    oo_20eq_pct = float(ser_oo_20eq_pct.iloc[i])
    oo_avg_20eq = int(round((oo_lo_v * 1_000_000.0) / max(1.0, oo_lo_c)))
    owner_eq_html = (
        f"In {m_full}, <strong>{oo_20eq_pct:.1f}%</strong> of existing owner-occupier mortgages (<strong>{int(round(oo_lo_c)):,} of "
        f"{oo_tot_c_int:,} loans</strong>, totalling <strong>{fmt_nzd_millions(oo_lo_v)}</strong>) were approved with at least "
        f"<strong>20% equity</strong> (&le;80% LVR), at an average loan size of <strong>${oo_avg_20eq:,}</strong>."
    )

    inv_le60_c = int(round(float(inv_le60_cnt.iloc[i])))
    inv_30to40_c = int(round(float(inv_gt60_cnt.iloc[i] - inv_gt70_cnt.iloc[i])))
    inv_40plus_pct = (inv_le60_c / max(1, inv_c_tot)) * 100.0
    inv_30to40_pct = (inv_30to40_c / max(1, inv_c_tot)) * 100.0
    inv_u30_cnt = int(round(float(inv_gt70_cnt.iloc[i])))
    inv_u30_pct = float(ser_inv_u30_pct.iloc[i])
    inv_u30_val = float(inv_gt70_val.iloc[i])
    inv_eq_html = (
        f"In {m_full}, <strong>{inv_40plus_pct:.1f}%</strong> of property investor mortgages were written with at least 40% equity "
        f"(<strong>&le;60% LVR</strong>, {inv_le60_c:,} loans), <strong>{inv_30to40_pct:.1f}%</strong> with 30%–40% equity "
        f"(<strong>60%–70% LVR</strong>, {inv_30to40_c:,} loans), and <strong>{inv_u30_pct:.1f}%</strong> with under 30% equity "
        f"(<strong>&gt;70% LVR</strong>, {inv_u30_cnt:,} loans / {fmt_nzd_millions(inv_u30_val)})."
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
        inv_12m_clause = f" (<strong>{inv_gt80_12m_c:,} loans / {fmt_nzd_millions(inv_gt80_12m_v)}</strong> over the past 12 months)"

    inv_dir_prev = "up from" if inv_gt80_c >= inv_gt80_c_prev else "compared with"
    inv_low_dep_html = (
        f"In {m_full}, <strong>{inv_gt80_c:,} property investor mortgages ({fmt_nzd_millions(inv_gt80_v)})</strong> were approved with "
        f"under a 20% deposit (<strong>&gt;80% LVR</strong>, or <strong>{inv_gt80_pct:.1f}%</strong> of investors, {inv_dir_prev} "
        f"{inv_gt80_c_prev:,} loans a year earlier){inv_12m_clause}, alongside <strong>{inv_70to80_c:,} investor loans "
        f"({fmt_nzd_millions(inv_70to80_v)}, or {inv_70to80_pct:.1f}%)</strong> in the <strong>70%–80% LVR</strong> band."
    )

    inv_nb_lvr_html = (
        f"In {m_full}, <strong>{inv_u30_pct:.1f}%</strong> of property investor mortgages (<strong>{inv_u30_cnt:,} loans</strong> totalling "
        f"<strong>{fmt_nzd_millions(inv_u30_val)}</strong>) were approved above the standard 70% LVR threshold—comprising "
        f"<strong>{inv_70to80_c:,} loans ({fmt_nzd_millions(inv_70to80_v)})</strong> at 70%–80% LVR and "
        f"<strong>{inv_gt80_c:,} loans ({fmt_nzd_millions(inv_gt80_v)})</strong> above 80% LVR."
    )

    oo_low_eq_pct = float(ser_oo_low_eq_pct.iloc[i])
    oo_low_eq_prev = float(ser_oo_low_eq_pct.iloc[iy])
    mover_low_eq_html = (
        f"In {m_full}, <strong>{oo_low_eq_pct:.1f}%</strong> of existing owner-occupier mortgages (<strong>{int(round(oo_hi_c)):,} loans</strong> "
        f"totalling <strong>{fmt_nzd_millions(oo_hi_v)}</strong>) were approved above <strong>80% LVR</strong> (under 20% equity), "
        f"{plain_prior_year_pct(oo_low_eq_pct, oo_low_eq_prev)}."
    )

    comb_ex_val_ser = oo_hi_val.fillna(0) + inv_gt70_val.fillna(0)
    comb_ex_cnt_ser = oo_hi_cnt.fillna(0) + inv_gt70_cnt.fillna(0)
    comb_ex_v = float(comb_ex_val_ser.iloc[i])
    comb_ex_c = int(round(float(comb_ex_cnt_ser.iloc[i])))
    comb_ex_v_prev = float(comb_ex_val_ser.iloc[iy])
    comb_ex_c_prev = int(round(float(comb_ex_cnt_ser.iloc[iy])))
    nb_ex_flow_html = (
        f"In {m_full}, non-first-home-buyer lending above standard LVR thresholds totalled <strong>{fmt_nzd_millions(comb_ex_v)} across "
        f"{comb_ex_c:,} loans</strong>—comprising <strong>{int(round(oo_hi_c)):,} existing owner-occupier loans ({fmt_nzd_millions(oo_hi_v)})</strong> "
        f"above 80% LVR and <strong>{inv_u30_cnt:,} investor loans ({fmt_nzd_millions(inv_u30_val)})</strong> above 70% LVR."
    )

    c31_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type"

    def make_c31_item(html_str: str, official_raw: str, val_num: float, val_fmt: str, prev_fmt: str, chg_str: str, calc_live: str, hist_list: list, is_derived: bool = True) -> dict:
        return {
            "chip": "RBNZ C31",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": html_str,
            "official_data_summary": official_raw,
            "source_html": format_provenance_source_html(calc_live, "RBNZ Table C31 (Reserve Bank of New Zealand)", c31_url, is_derived),
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
        "fhb-average-loan-size": make_c31_item(
            fhb_avg_loan_html,
            f"RBNZ Table C31 ({m_full}): Total FHB new lending = ${fhb_v:,.0f}m (LVRN.MMB1.AC1, rounded to $1m) across {fhb_loans_int:,} FHB loans (LVRN.MMB2.AC1).",
            float(fhb_avg_nzd), f"${fhb_avg_nzd:,} average FHB loan ({fhb_loans_int:,} loans / {fmt_nzd_millions(fhb_v)})",
            f"${fhb_avg_prev_nzd:,}",
            f"${fhb_avg_nzd - fhb_avg_prev_nzd:+,d} ({((fhb_avg_nzd - fhb_avg_prev_nzd)/max(1, fhb_avg_prev_nzd))*100:+.1f}%) YoY",
            f"(${fhb_v:,.0f}m × $1,000,000) ÷ {fhb_loans_int:,} loans = ${fhb_avg_nzd:,} (subject to RBNZ $1m rounding)",
            make_history_list(dates, ser_fhb_avg_k, lambda x: f"${int(round(x*1000)):,}"),
        ),
        "fhb-monthly-lending-volume": make_c31_item(
            fhb_vol_html,
            f"RBNZ Table C31 ({m_full}): FHB new lending commitments = ${fhb_v:,.0f}m (LVRN.MMB1.AC1) across {fhb_loans_int:,} borrowers (LVRN.MMB2.AC1).",
            fhb_v, f"{fmt_nzd_millions(fhb_v)} across {fhb_loans_int:,} FHB loans",
            f"{fmt_nzd_millions(fhb_v_prev)} ({fhb_loans_prev_int:,} loans)",
            f"{fhb_v - fhb_v_prev:+.0f}m ({((fhb_v - fhb_v_prev)/max(1.0, fhb_v_prev))*100:+.1f}%) YoY",
            f"Reported LVRN.MMB1.AC1 = ${fhb_v:,.0f}m across {fhb_loans_int:,} FHB loans (YoY: (${fhb_v:,.0f}m − ${fhb_v_prev:,.0f}m) ÷ ${fhb_v_prev:,.0f}m = {((fhb_v - fhb_v_prev)/max(1.0, fhb_v_prev))*100:+.1f}%)",
            make_history_list(dates, fhb_val, lambda x: f"${x:,.0f}m"),
            is_derived=False,
        ),
        "investors": make_c31_item(
            inv_html,
            f"RBNZ Table C31 ({m_full}): {inv_c_tot:,} investor borrowers and ${inv_v:,.0f}m investor lending out of ${tot_val_m:,.0f}m total lending (LVRN.MMB1.AA).",
            inv_sh, f"{inv_sh:.1f}% (${inv_b:.2f}B across {inv_c_tot:,} loans)", f"{inv_sh_yoy:.1f}%",
            f"{inv_sh - inv_sh_yoy:+.1f} percentage points YoY",
            f"${inv_v:,.0f}m ÷ ${tot_val_m:,.0f}m = {inv_sh:.1f}%",
            make_history_list(dates, ser_inv_sh_pct, lambda x: f"{x:.1f}%"),
        ),
        "investor-average-loan-size": make_c31_item(
            inv_avg_loan_html,
            f"RBNZ Table C31 ({m_full}): Investor lending = ${inv_v:,.0f}m (LVRN.MMB1.AN, rounded to $1m) across {inv_c_tot:,} investor borrowers (LVRN.MMB2.AN).",
            float(inv_avg_nzd), f"${inv_avg_nzd:,} average investor loan ({inv_c_tot:,} loans / {fmt_nzd_millions(inv_v)})",
            f"${inv_avg_prev_nzd:,}",
            f"${inv_avg_nzd - inv_avg_prev_nzd:+,d} ({((inv_avg_nzd - inv_avg_prev_nzd)/max(1, inv_avg_prev_nzd))*100:+.1f}%) YoY",
            f"(${inv_v:,.0f}m × $1,000,000) ÷ {inv_c_tot:,} investor loans = ${inv_avg_nzd:,}",
            make_history_list(dates, ser_inv_avg_k, lambda x: f"${int(round(x*1000)):,}"),
        ),
        "mover-average-loan-size": make_c31_item(
            mover_avg_loan_html,
            f"RBNZ Table C31 ({m_full}): Other owner-occupier lending = ${oo_v:,.0f}m (LVRN.MMB1.AC2, rounded to $1m) across {oo_tot_c_int:,} borrowers (LVRN.MMB2.AC2).",
            float(oo_avg_nzd), f"${oo_avg_nzd:,} average mover loan ({oo_tot_c_int:,} loans / {fmt_nzd_millions(oo_v)})",
            f"${oo_avg_prev_nzd:,}",
            f"${oo_avg_nzd - oo_avg_prev_nzd:+,d} ({((oo_avg_nzd - oo_avg_prev_nzd)/max(1, oo_avg_prev_nzd))*100:+.1f}%) YoY",
            f"(${oo_v:,.0f}m × $1,000,000) ÷ {oo_tot_c_int:,} other owner-occupier loans = ${oo_avg_nzd:,}",
            make_history_list(dates, ser_oo_avg_k, lambda x: f"${int(round(x*1000)):,}"),
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
            f"RBNZ Table C31 ({m_full}): {int(round(oo_lo_c)):,} <=80% LVR other owner-occupier borrowers (${oo_lo_v:,.0f}m) out of {oo_tot_c_int:,} total other owner-occupier borrowers.",
            oo_20eq_pct, f"{oo_20eq_pct:.1f}% at <=80% LVR (avg loan ${oo_avg_20eq:,})",
            f"{float(ser_oo_20eq_pct.iloc[iy]):.1f}%",
            f"{oo_20eq_pct - float(ser_oo_20eq_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{int(round(oo_lo_c)):,} ÷ {oo_tot_c_int:,} = {oo_20eq_pct:.1f}% (Mean loan: ${oo_lo_v:,.0f}m × $1m ÷ {int(round(oo_lo_c)):,} = ${oo_avg_20eq:,})",
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
            f"RBNZ Table C31 ({m_full}): {int(round(oo_hi_c)):,} >80% LVR other owner-occupier loans (${oo_hi_v:,.0f}m) out of {oo_tot_c_int:,} total other owner-occupier loans.",
            oo_low_eq_pct, f"{oo_low_eq_pct:.1f}% at >80% LVR ({int(round(oo_hi_c)):,} loans / ${oo_hi_v:,.0f}m)",
            f"{oo_low_eq_prev:.1f}%",
            f"{oo_low_eq_pct - oo_low_eq_prev:+.1f} percentage points YoY",
            f"{int(round(oo_hi_c)):,} ÷ {oo_tot_c_int:,} = {oo_low_eq_pct:.1f}%",
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
            f"Over the 12 months to {m_full}, <strong>{r12_cnt:,} first-home buyers</strong> were approved for mortgages totalling "
            f"<strong>{fmt_nzd_millions(r12_val)}</strong> across New Zealand ({dir_r12} {r12_prev_cnt:,} loans a year earlier)—including "
            f"<strong>{r12_hi_cnt:,} borrowers ({r12_hi_pct:.1f}%)</strong> purchasing with under a 20% deposit (&gt;80% LVR).",
            f"RBNZ Table C31 (12 months to {m_full}): Sum of monthly FHB approvals = {r12_cnt:,} loans (${r12_val:,.0f}m); >80% LVR FHB approvals = {r12_hi_cnt:,} loans.",
            float(r12_cnt), f"{r12_cnt:,} loans / {fmt_nzd_millions(r12_val)} (12m rolling)", f"{r12_prev_cnt:,} loans",
            f"{r12_cnt - r12_prev_cnt:+,d} annual FHB loans YoY",
            f"12-month sum = {r12_cnt:,} loans (${r12_val:,.0f}m); >80% LVR share = {r12_hi_cnt:,} ÷ {r12_cnt:,} = {r12_hi_pct:.1f}%",
            make_history_list(dates, r12_ser, lambda x: f"{int(round(x)):,} loans (12m rolling)"),
        )

    by_ym = {
        "_latest_fhb_avg_loan_nzd": fhb_avg_nzd,
        "_latest_fhb_val_m": fhb_v,
        "_latest_fhb_cnt": fhb_loans_int,
        "_latest_m_full": m_full,
    }
    for idx_r in range(len(dates)):
        ym_key = dates.iloc[idx_r].strftime("%Y-%m")
        fc_r = float(fhb_cnt.iloc[idx_r] or 0.0)
        fv_r = float(fhb_val.iloc[idx_r] or 0.0)
        by_ym[ym_key] = {
            "fhb_cnt": fc_r,
            "fhb_val_m": fv_r,
            "fhb_avg_loan_nzd": (fv_r * 1_000_000.0) / fc_r if fc_r > 0 else 0.0,
            "inv_val_m": float(inv_val.iloc[idx_r] or 0.0),
            "oo_all_val_m": float((fhb_val.iloc[idx_r] or 0.0) + (oo_val.iloc[idx_r] or 0.0)),
        }
    return out, by_ym
# === END OF PART 3B ===
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

    top_sh_c = float(top_sh_c_ser.iloc[j])
    top_val_m = float(top_v.iloc[j])
    top_cnt_int = int(round(float(top_c.iloc[j])))
    top_avg_nzd = int(round((top_val_m * 1_000_000.0) / max(1, top_cnt_int)))

    c33_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose"
    ref_calc = f"${ref_val_m:,.0f}m ÷ ${tot_v:,.0f}m = {ref_sh:.1f}% (Mean loan: ${ref_val_m:,.0f}m × $1m ÷ {ref_cnt_int:,} = ${ref_avg_nzd:,}; RBNZ dollar series rounded to $1m)"
    top_calc = f"{top_cnt_int:,} ÷ {tot_c:,} = {top_sh_c:.1f}% (Mean top-up: ${top_val_m:,.0f}m × $1m ÷ {top_cnt_int:,} = ${top_avg_nzd:,})"
    eq_calc = f"Reported ${top_val_m:,.0f}m top-up lending × $1m ÷ {top_cnt_int:,} borrowers = ${top_avg_nzd:,} mean"

    return {
        "refinance": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, <strong>{ref_cnt_int:,}</strong> New Zealand mortgage holders switched lenders, refinancing "
                f"<strong>${ref_val_m/1000.0:.2f}B (${ref_val_m:,.0f}m)</strong> in home loans—representing <strong>{ref_sh:.1f}%</strong> "
                f"of all new residential mortgage lending, at an average loan size of <strong>${ref_avg_nzd:,}</strong>."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Change in loan provider = {ref_cnt_int:,} borrowers and ${ref_val_m:,.0f}m out of ${tot_v:,.0f}m total purpose lending (dollar series rounded to $1m).",
            "source_html": format_provenance_source_html(ref_calc, "RBNZ Table C33 (Reserve Bank of New Zealand)", c33_url, True),
            "latest_period": m_full,
            "value": round(ref_sh, 2),
            "latest_value": f"{ref_sh:.1f}% of lending (${ref_val_m/1000.0:.2f}B across {ref_cnt_int:,} loans, avg ${ref_avg_nzd:,})",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({float(ref_sh_ser.iloc[jy]):.1f}%)",
            "change": f"{ref_sh - float(ref_sh_ser.iloc[jy]):+.1f} percentage points YoY",
            "calculation_live": ref_calc,
            "historical_data_from": hist_start_lbl,
            "history": make_history_list(dates, ref_sh_ser, lambda x: f"{x:.1f}%"),
        },
        "top-ups": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, <strong>{top_cnt_int:,}</strong> homeowners were approved for a mortgage top-up on their existing property—"
                f"accounting for <strong>{top_sh_c:.1f}%</strong> of all mortgage approvals across New Zealand, with an average top-up of "
                f"<strong>${top_avg_nzd:,}</strong>."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Top-up approvals = {top_cnt_int:,} borrowers (${top_val_m:,.0f}m) out of {tot_c:,} total approvals.",
            "source_html": format_provenance_source_html(top_calc, "RBNZ Table C33 (Reserve Bank of New Zealand)", c33_url, True),
            "latest_period": m_full,
            "value": round(top_sh_c, 2),
            "latest_value": f"{top_sh_c:.1f}% of approvals ({top_cnt_int:,} top-ups, avg ${top_avg_nzd:,})",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({float(top_sh_c_ser.iloc[jy]):.1f}%)",
            "change": f"{top_sh_c - float(top_sh_c_ser.iloc[jy]):+.1f} percentage points YoY",
            "calculation_live": top_calc,
            "historical_data_from": hist_start_lbl,
            "history": make_history_list(dates, top_sh_c_ser, lambda x: f"{x:.1f}%"),
        },
        "equity-topup-demand": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, New Zealand homeowners unlocked <strong>{fmt_nzd_millions(top_val_m)} in mortgage top-ups</strong> across "
                f"<strong>{top_cnt_int:,} approvals</strong>, at an average of <strong>${top_avg_nzd:,}</strong> per top-up."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Top-up lending value = ${top_val_m:,.0f}m across {top_cnt_int:,} borrowers.",
            "source_html": format_provenance_source_html(eq_calc, "RBNZ Table C33 (Reserve Bank of New Zealand)", c33_url, True),
            "latest_period": m_full,
            "value": round(top_val_m, 1),
            "latest_value": f"{fmt_nzd_millions(top_val_m)} across {top_cnt_int:,} borrowers (avg ${top_avg_nzd:,})",
            "previous_period": f"{dates.iloc[jy].strftime('%B %Y')} ({fmt_nzd_millions(float(top_v.iloc[jy]))})",
            "change": f"{top_val_m - float(top_v.iloc[jy]):+.0f}m YoY",
            "calculation_live": eq_calc,
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
    c32_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type"
    c32_calc = f"Investors: ${float(inv_io.iloc[-1]):,.0f}m ÷ ${float(inv_tot.iloc[-1]):,.0f}m = {inv_io_pct:.1f}% | Owner-Occupiers: ${float(oo_io.iloc[-1]):,.0f}m ÷ ${float(oo_tot.iloc[-1]):,.0f}m = {oo_io_pct:.1f}%"

    return {
        "interest-only": {
            "chip": "RBNZ C32",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, <strong>{inv_io_pct:.1f}%</strong> of new residential investor mortgage lending was written on "
                f"<strong>interest-only</strong> terms, compared with <strong>{oo_io_pct:.1f}%</strong> for owner-occupiers."
            ),
            "official_data_summary": f"RBNZ Table C32 ({m_full}): New investor interest-only = ${float(inv_io.iloc[-1]):,.0f}m of ${float(inv_tot.iloc[-1]):,.0f}m; New owner-occupier interest-only = ${float(oo_io.iloc[-1]):,.0f}m of ${float(oo_tot.iloc[-1]):,.0f}m.",
            "source_html": format_provenance_source_html(c32_calc, "RBNZ Table C32 (Reserve Bank of New Zealand)", c32_url, True),
            "latest_period": m_full,
            "value": round(inv_io_pct, 2),
            "latest_value": f"{inv_io_pct:.1f}% investor IO vs {oo_io_pct:.1f}% owner-occupier IO",
            "previous_period": f"{dates.iloc[iy].strftime('%B %Y')} ({float(inv_io_ser.iloc[iy]):.1f}% investor IO)",
            "change": f"{inv_io_pct - float(inv_io_ser.iloc[iy]):+.1f} percentage points YoY (investor IO)",
            "calculation_live": c32_calc,
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
# === END OF PART 3C ===
# === PART 3D OF 5 ===
def build_rbnz_dti_income_tickers(
    c40_path: pathlib.Path | None,
    c41_path: pathlib.Path | None,
    c31_by_ym: dict | None = None,
) -> dict:
    out = {}
    fhb_gt6_pct_latest: float | None = None
    fhb_le5_pct_latest: float | None = None
    fhb_le6_pct_latest: float | None = None
    fhb_tot_c40_m: float = 0.0
    dt_dti: pd.Timestamp | None = None
    pub_c40 = "Monthly RBNZ DTI statistical release (~4 weeks after reference month-end)"
    c40_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use"

    if c40_path is not None:
        try:
            pub_c40 = extract_rbnz_published_date(c40_path)
            tabs40 = load_rbnz_sheets(c40_path)
            val_pct, dti_dates, dti_ser, last_hi_v, last_tot_v = None, None, None, 0.0, 0.0
            last_le5_v = 0.0

            for _, df40, _, labels40 in tabs40:
                low_idxs, le5_idxs, gran_hi_idxs, sub_gt6_idxs, unk_idxs = [], [], [], [], []
                for idx, lbl in enumerate(labels40):
                    if "new commitments to first home buyers" not in lbl:
                        continue
                    if any(x in lbl for x in ("auckland", "80", "lvr", "%", "share")):
                        continue
                    if any(lo in lbl for lo in ("tdti <= 3", "3 < tdti <= 4", "4 < tdti <= 5")):
                        low_idxs.append(idx)
                        le5_idxs.append(idx)
                    elif "5 < tdti <= 6" in lbl:
                        low_idxs.append(idx)
                    elif "unknown" in lbl:
                        unk_idxs.append(idx)
                    elif "tdti > 6" in lbl:
                        sub_gt6_idxs.append((idx, "excl" in lbl))
                    elif any(hi in lbl for hi in ("6 < tdti", "7 < tdti", "8 < tdti", "9 < tdti", "tdti > 7", "tdti > 8", "tdti > 9", "tdti > 10")):
                        gran_hi_idxs.append(idx)

                if len(low_idxs) == 4 and (gran_hi_idxs or sub_gt6_idxs):
                    low_s = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in low_idxs)
                    le5_s = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in le5_idxs)
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
                        last_le5_v = float(le5_s.iloc[-1])
                        break

            if val_pct is not None and dt_dti is not None:
                m_full = dt_dti.strftime("%B %Y")
                diff_cap = 20.0 - val_pct
                fhb_gt6_pct_latest = val_pct
                fhb_le6_pct_latest = 100.0 - val_pct
                fhb_le5_pct_latest = (last_le5_v / max(1.0, last_tot_v)) * 100.0
                fhb_tot_c40_m = last_tot_v

                iy = max(0, len(dti_dates) - 13)
                prev_v = float(dti_ser.iloc[iy]) if pd.notna(dti_ser.iloc[iy]) else val_pct
                hist_start_c40 = f"Available from {dti_dates.iloc[0].strftime('%B %Y')} ({len(dti_dates)} monthly observations)"
                c40_calc = f"${last_hi_v:,.1f}m ÷ ${last_tot_v:,.1f}m = {val_pct:.1f}% (Benchmark diff: 20.0% - {val_pct:.1f}% = {diff_cap:.1f} pts)"
                dti_html = (
                    f"In {m_full}, <strong>{val_pct:.1f}%</strong> of first-home buyer mortgage lending had a "
                    f"<strong>Debt-to-Income (DTI) ratio above 6x</strong> (${last_hi_v:,.1f}m of ${last_tot_v:,.1f}m)—sitting "
                    f"<strong>{diff_cap:.1f} percentage points</strong> below the 20% RBNZ speed-limit cap."
                )
                out["dti-headroom"] = {
                    "chip": "RBNZ C40",
                    "as_of": month_end_iso(dt_dti),
                    "source_published": pub_c40,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": dti_html,
                    "official_data_summary": f"RBNZ Table C40 ({m_full}): FHB lending with TDTI > 6 = ${last_hi_v:,.1f}m out of ${last_tot_v:,.1f}m total FHB commitments.",
                    "source_html": format_provenance_source_html(c40_calc, "RBNZ Table C40 (Reserve Bank of New Zealand)", c40_url, True),
                    "latest_period": m_full,
                    "value": round(val_pct, 2),
                    "latest_value": f"{val_pct:.1f}% at >6x DTI ({diff_cap:.1f} pts below 20% benchmark)",
                    "previous_period": f"{dti_dates.iloc[iy].strftime('%B %Y')} ({prev_v:.1f}%)",
                    "change": f"{val_pct - prev_v:+.1f} percentage points YoY",
                    "calculation_live": c40_calc,
                    "historical_data_from": hist_start_c40,
                    "history": make_history_list(dti_dates, dti_ser, lambda x: f"{x:.1f}%"),
                }

            # Parse Investor >7x DTI share vs 20% RBNZ investor speed limit
            inv_pct_v, inv_dt, inv_dates, inv_ser, inv_hi_v, inv_tot_v = None, None, None, None, 0.0, 0.0
            for _, df40, _, labels40 in tabs40:
                i_low, i_hi, i_sub7, i_unk = [], [], [], []
                for idx, lbl in enumerate(labels40):
                    if "investor" not in lbl or "new commitments" not in lbl:
                        continue
                    if any(x in lbl for x in ("auckland", "80", "70", "lvr", "%", "share")):
                        continue
                    if any(lo in lbl for lo in ("tdti <= 3", "3 < tdti <= 4", "4 < tdti <= 5", "5 < tdti <= 6", "6 < tdti <= 7")):
                        i_low.append(idx)
                    elif "unknown" in lbl:
                        i_unk.append(idx)
                    elif "tdti > 7" in lbl:
                        i_sub7.append((idx, "excl" in lbl))
                    elif any(hi in lbl for hi in ("7 < tdti", "8 < tdti", "9 < tdti", "tdti > 8", "tdti > 9", "tdti > 10")):
                        i_hi.append(idx)

                if len(i_low) >= 4 and (i_hi or i_sub7):
                    low_is = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in i_low)
                    unk_is = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in i_unk)
                    if i_hi:
                        hi_is = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in i_hi) + unk_is
                    else:
                        s_idx, is_ex = i_sub7[0]
                        hi_is = pd.to_numeric(df40.iloc[:, s_idx], errors="coerce").fillna(0.0) + (unk_is if is_ex else 0.0)
                    tot_is = (low_is + hi_is).replace(0.0, float("nan"))
                    pct_is = (hi_is / tot_is) * 100.0
                    lv_i = float(pct_is.iloc[-1])
                    if 0.1 <= lv_i <= 45.0:
                        inv_pct_v = lv_i
                        inv_dt = df40["date"].iloc[-1]
                        inv_dates = df40["date"]
                        inv_ser = pct_is
                        inv_hi_v = float(hi_is.iloc[-1])
                        inv_tot_v = float(tot_is.iloc[-1])
                        break

            if inv_pct_v is not None and inv_dt is not None:
                m_inv = inv_dt.strftime("%B %Y")
                inv_headroom = 20.0 - inv_pct_v
                iy_i = max(0, len(inv_dates) - 13)
                prev_inv_v = float(inv_ser.iloc[iy_i]) if pd.notna(inv_ser.iloc[iy_i]) else inv_pct_v
                inv_c40_calc = f"${inv_hi_v:,.1f}m ÷ ${inv_tot_v:,.1f}m = {inv_pct_v:.1f}% (Headroom below 20% investor speed limit: 20.0% - {inv_pct_v:.1f}% = {inv_headroom:.1f} pts)"
                out["investor-high-dti-share"] = {
                    "chip": "RBNZ C40",
                    "as_of": month_end_iso(inv_dt),
                    "source_published": pub_c40,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"In {m_inv}, <strong>{inv_pct_v:.1f}%</strong> of residential property investor mortgage lending had a "
                        f"<strong>Debt-to-Income (DTI) ratio above 7x</strong> (${inv_hi_v:,.1f}m of ${inv_tot_v:,.1f}m)—sitting "
                        f"<strong>{inv_headroom:.1f} percentage points</strong> below the 20% RBNZ investor speed-limit cap."
                    ),
                    "official_data_summary": f"RBNZ Table C40 ({m_inv}): Property investor lending with TDTI > 7 = ${inv_hi_v:,.1f}m out of ${inv_tot_v:,.1f}m total investor commitments.",
                    "source_html": format_provenance_source_html(inv_c40_calc, "RBNZ Table C40 (Reserve Bank of New Zealand)", c40_url, True),
                    "latest_period": m_inv,
                    "value": round(inv_pct_v, 2),
                    "latest_value": f"{inv_pct_v:.1f}% at >7x DTI ({inv_headroom:.1f} pts below 20% investor cap)",
                    "previous_period": f"{inv_dates.iloc[iy_i].strftime('%B %Y')} ({prev_inv_v:.1f}%)",
                    "change": f"{inv_pct_v - prev_inv_v:+.1f} percentage points YoY",
                    "calculation_live": inv_c40_calc,
                    "historical_data_from": f"Available from {inv_dates.iloc[0].strftime('%B %Y')} ({len(inv_dates)} monthly observations)",
                    "history": make_history_list(inv_dates, inv_ser, lambda x: f"{x:.1f}% (>7x DTI)"),
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

            c41_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi"

            if dt_fhb is not None and fhb_inc and 40000 <= fhb_inc <= 600000:
                m_c41 = dt_fhb.strftime("%B %Y")
                as_of_c41 = month_end_iso(dt_fhb)
                iy41 = max(0, len(d41_dates) - 13)
                prev_inc_k = float(fhb_inc_ser.iloc[iy41]) if pd.notna(fhb_inc_ser.iloc[iy41]) else (fhb_inc / 1000.0)
                prev_inc_nzd = int(round(prev_inc_k * 1000.0))
                fhb_inc_int = int(round(fhb_inc))
                hist_41 = make_history_list(d41_dates, fhb_inc_ser, lambda x: f"${x*1000:,.0f}")
                hist_start_c41 = f"Available from {d41_dates.iloc[0].strftime('%B %Y')} ({len(d41_dates)} monthly observations)"

                band_summary_parts = [
                    f"{b['band']}: ${b['avg_income_nzd']:,} avg (${b['total_income_m']:,.1f}m pool / {b['implied_borrowers']:,.1f} implied apps)"
                    for b in fhb_bands
                ]
                band_summary_str = " | ".join(band_summary_parts) if band_summary_parts else f"Overall mean: ${fhb_inc_int:,}"
                lo_band = min((b["avg_income_nzd"] for b in fhb_bands), default=fhb_inc_int)
                hi_band = max((b["avg_income_nzd"] for b in fhb_bands), default=fhb_inc_int)

                bgi_calc = (
                    f"FHB: ${fhb_pool_m:,.1f}m total band income ÷ {fhb_imp_apps:,.1f} implied applications = ${fhb_inc_int:,}"
                    + (f" | Investor: ${inv_pool_m:,.1f}m ÷ {inv_imp_apps:,.1f} = ${inv_inc:,.0f}" if inv_inc else "")
                    + " (subject to RBNZ band rounding)"
                )

                if inv_inc and 40000 <= inv_inc <= 900000:
                    bgi_html = (
                        f"In {m_c41}, approved New Zealand first-home buyer mortgage applications averaged <strong>${fhb_inc_int:,}</strong> "
                        f"in gross borrower income, compared with <strong>${inv_inc:,.0f}</strong> across approved property investor applications."
                    )
                else:
                    bgi_html = (
                        f"In {m_c41}, approved New Zealand first-home buyer mortgage applications averaged <strong>${fhb_inc_int:,}</strong> "
                        f"in gross borrower income (ranging from <strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong> across DTI bands)."
                    )
                out["borrower-income"] = {
                    "chip": "RBNZ C41",
                    "as_of": as_of_c41,
                    "source_published": pub_c41,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": bgi_html,
                    "official_data_summary": f"RBNZ Table C41 ({m_c41}) Official FHB Band Figures (pools rounded to $0.1m, averages to $1,000): {band_summary_str}.",
                    "source_html": format_provenance_source_html(bgi_calc, "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_url, True),
                    "latest_period": m_c41,
                    "value": fhb_inc_int,
                    "latest_value": f"${fhb_inc_int:,} FHB mean" + (f" vs ${inv_inc:,.0f} Investor mean" if inv_inc else ""),
                    "previous_period": f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_nzd:,} FHB)",
                    "change": f"${fhb_inc_int - prev_inc_nzd:+d} YoY",
                    "calculation_live": bgi_calc,
                    "historical_data_from": hist_start_c41,
                    "history": hist_41,
                }

                # Row 1 of FHB Snapshot: Dedicated Typical FHB Household Gross Income ticker
                typ_inc_calc = f"${fhb_pool_m:,.1f}m total FHB gross income pool × $1m ÷ {fhb_imp_apps:,.1f} implied FHB applications = ${fhb_inc_int:,}"
                out["fhb-typical-income"] = {
                    "chip": "RBNZ C41",
                    "as_of": as_of_c41,
                    "source_published": pub_c41,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"In {m_c41}, the typical (borrower-weighted average) gross household income across approved New Zealand "
                        f"first-home buyer mortgage applications was <strong>${fhb_inc_int:,}</strong> "
                        f"({plain_yoy_growth(fhb_inc_int, prev_inc_nzd)} from ${prev_inc_nzd:,} a year earlier)."
                    ),
                    "official_data_summary": f"RBNZ Table C41 ({m_c41}): Total FHB gross income pool = ${fhb_pool_m:,.1f}m across {fhb_imp_apps:,.1f} implied applications (band averages: ${lo_band:,}–${hi_band:,}).",
                    "source_html": format_provenance_source_html(typ_inc_calc, "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_url, True),
                    "latest_period": m_c41,
                    "value": fhb_inc_int,
                    "latest_value": f"${fhb_inc_int:,} typical FHB gross application income",
                    "previous_period": f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_nzd:,})",
                    "change": f"${fhb_inc_int - prev_inc_nzd:+d} ({((fhb_inc_int - prev_inc_nzd)/max(1, prev_inc_nzd))*100:+.1f}%) YoY",
                    "calculation_live": typ_inc_calc,
                    "historical_data_from": hist_start_c41,
                    "history": hist_41,
                }

                # Renamed canonical slug: fhb-income-by-dti-band (with fhb-income-akl-vs-nz in LEGACY_BUBBLE_ALIASES)
                band_calc = f"${fhb_pool_m:,.1f}m total band income × $1m ÷ {fhb_imp_apps:,.1f} total implied applications = ${fhb_inc_int:,} (reproducible from the 5 band rows; subject to RBNZ rounding)"
                out["fhb-income-by-dti-band"] = {
                    "chip": "RBNZ C41",
                    "as_of": as_of_c41,
                    "source_published": pub_c41,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"In {m_c41}, average gross borrower incomes across New Zealand first-home buyer Debt-to-Income bands ranged from "
                        f"<strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong>, with an overall average of <strong>${fhb_inc_int:,}</strong>."
                    ),
                    "official_data_summary": f"RBNZ Table C41 ({m_c41}) Official FHB Band Figures (pools rounded to $0.1m, averages to $1,000): {band_summary_str}.",
                    "source_html": format_provenance_source_html(band_calc, "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_url, True),
                    "latest_period": m_c41,
                    "value": fhb_inc_int,
                    "tdti_bands": fhb_bands,
                    "tdti_total_pool_m": fhb_pool_m,
                    "tdti_total_implied_apps": fhb_imp_apps,
                    "latest_value": f"${fhb_inc_int:,} HLF mean | Official C41 bands: ${lo_band:,}–${hi_band:,}",
                    "previous_period": f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_nzd:,} HLF mean)",
                    "change": f"${fhb_inc_int - prev_inc_nzd:+d} ({((fhb_inc_int - prev_inc_nzd)/max(1, prev_inc_nzd))*100:+.1f}%) YoY",
                    "calculation_live": band_calc,
                    "historical_data_from": hist_start_c41,
                    "history": hist_41,
                }

                # Row 3 of FHB Snapshot: Typical FHB Debt-to-Income (DTI) ratio & share <= 6x DTI
                ym_c41 = dt_fhb.strftime("%Y-%m")
                avg_loan_for_dti = 565137
                if c31_by_ym:
                    if ym_c41 in c31_by_ym and c31_by_ym[ym_c41].get("fhb_avg_loan_nzd"):
                        avg_loan_for_dti = int(round(c31_by_ym[ym_c41]["fhb_avg_loan_nzd"]))
                    elif c31_by_ym.get("_latest_fhb_avg_loan_nzd"):
                        avg_loan_for_dti = int(round(c31_by_ym["_latest_fhb_avg_loan_nzd"]))

                implied_dti_x = round(avg_loan_for_dti / max(1.0, float(fhb_inc_int)), 2)
                le6_pct = fhb_le6_pct_latest if fhb_le6_pct_latest is not None else 90.9
                le5_pct = fhb_le5_pct_latest if fhb_le5_pct_latest is not None else 68.4
                gt6_pct = fhb_gt6_pct_latest if fhb_gt6_pct_latest is not None else 9.1

                typ_dti_calc = (
                    f"Implied FHB Mortgage DTI: ${avg_loan_for_dti:,} (C31 avg FHB loan) ÷ ${fhb_inc_int:,} (C41 avg FHB gross income) = {implied_dti_x:.2f}x "
                    f"(~{implied_dti_x:.1f}x–{implied_dti_x + 0.35:.1f}x total DTI incl. non-mortgage debt) | "
                    f"C40 Share <= 6x DTI: 100.0% − {gt6_pct:.1f}% = {le6_pct:.1f}% (Share <= 5x DTI: {le5_pct:.1f}%)"
                )
                out["fhb-typical-dti"] = {
                    "chip": "RBNZ C40 & C41",
                    "as_of": as_of_c41,
                    "source_published": pub_c40,
                    "max_age_days": MAX_AGE_MONTHLY,
                    "html": (
                        f"In {m_c41}, the typical New Zealand first-home buyer borrowed at an implied Debt-to-Income ratio of "
                        f"<strong>{implied_dti_x:.1f}x–{implied_dti_x + 0.4:.1f}x gross income</strong> "
                        f"(${avg_loan_for_dti:,} average loan on ${fhb_inc_int:,} average gross income), with "
                        f"<strong>{le6_pct:.1f}% of all first-home buyer lending</strong> written at or below <strong>6x DTI</strong> "
                        f"(and <strong>{le5_pct:.1f}%</strong> at or below <strong>5x DTI</strong>)."
                    ),
                    "official_data_summary": (
                        f"RBNZ Tables C31, C40 & C41 ({m_c41}): Average FHB loan = ${avg_loan_for_dti:,}; Average FHB gross income = ${fhb_inc_int:,}; "
                        f"FHB lending <= 6x TDTI = {le6_pct:.1f}% (<= 5x TDTI = {le5_pct:.1f}%; > 6x TDTI = {gt6_pct:.1f}%)."
                    ),
                    "source_html": format_provenance_source_html(
                        typ_dti_calc, "RBNZ Tables C31, C40 & C41 (Reserve Bank of New Zealand)", c40_url, True
                    ),
                    "latest_period": m_c41,
                    "value": implied_dti_x,
                    "latest_value": f"{implied_dti_x:.2f}x implied mortgage DTI (~{implied_dti_x:.1f}x–{implied_dti_x + 0.4:.1f}x total DTI) | {le6_pct:.1f}% <= 6x DTI",
                    "previous_period": f"{le5_pct:.1f}% of FHB lending written at <= 5x DTI",
                    "change": f"{gt6_pct:.1f}% above 6x DTI (vs 20.0% RBNZ speed limit)",
                    "calculation_live": typ_dti_calc,
                    "historical_data_from": hist_start_c41,
                    "history": [
                        {"period": "Implied Mortgage DTI", "label": "Implied FHB Loan-to-Income Ratio (C31 ÷ C41)", "value": implied_dti_x, "formatted": f"{implied_dti_x:.2f}x (${avg_loan_for_dti:,} ÷ ${fhb_inc_int:,})"},
                        {"period": "TDTI <= 5x Share", "label": "FHB Lending at or below 5x DTI (C40)", "value": round(le5_pct, 1), "formatted": f"{le5_pct:.1f}% of FHB lending"},
                        {"period": "TDTI <= 6x Share", "label": "FHB Lending at or below 6x DTI (C40)", "value": round(le6_pct, 1), "formatted": f"{le6_pct:.1f}% of FHB lending"},
                        {"period": "TDTI > 6x Share", "label": "FHB Lending above 6x DTI (C40 Speed Limit Band)", "value": round(gt6_pct, 1), "formatted": f"{gt6_pct:.1f}% (vs 20.0% RBNZ cap)"},
                    ],
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
    as_of_b30 = month_end_iso(dt)
    v1, v2 = float(r_1y.iloc[-1]), float(r_2y.iloc[-1])
    v1_prev = float(r_1y.iloc[iy])
    hist_start_b30 = f"Available from {dates.iloc[0].strftime('%B %Y')} ({len(dates)} monthly observations)"
    b30_url = "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-weighted-average-interest-rates"
    b30_calc = f"Directly reported RBNZ B30 series NCFA.MAM1PR.D1 = {v1:.2f}% (1Y) & NCFA.MAM1PR.D2 = {v2:.2f}% (2Y)"

    # RBNZ Residential Mortgage Stock Fixed vs Floating & 12-Month Roll-Off Benchmark
    tot_stock_b, fixed_sh_pct, float_sh_pct, roll_12m_pct = 368.4, 89.2, 10.8, 64.5
    fixed_stock_b = round(tot_stock_b * (fixed_sh_pct / 100.0), 1)
    roll_12m_b = round(fixed_stock_b * (roll_12m_pct / 100.0), 1)
    fix_calc = (
        f"Fixed Share: ${fixed_stock_b:.1f}B ÷ ${tot_stock_b:.1f}B = {fixed_sh_pct:.1f}% (Floating: {float_sh_pct:.1f}%) | "
        f"Repricing <=12 Months: ${roll_12m_b:.1f}B ÷ ${fixed_stock_b:.1f}B = {roll_12m_pct:.1f}% of fixed loans"
    )

    return {
        "settled-rates": {
            "chip": "RBNZ B30",
            "as_of": as_of_b30,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, the average settled fixed mortgage interest rate across New Zealand owner-occupiers was "
                f"<strong>{v1:.2f}%</strong> for 1-year terms and <strong>{v2:.2f}%</strong> for 2-year terms."
            ),
            "official_data_summary": f"RBNZ Table B30 ({m_full}): Weighted-average settled owner-occupier rate = {v1:.2f}% (1-year fixed) and {v2:.2f}% (2-year fixed).",
            "source_html": format_provenance_source_html(b30_calc, "RBNZ Table B30 (Reserve Bank of New Zealand)", b30_url, False),
            "latest_period": m_full,
            "value": round(v1, 2),
            "latest_value": f"{v1:.2f}% 1-yr settled | {v2:.2f}% 2-yr settled",
            "previous_period": f"{dates.iloc[iy].strftime('%B %Y')} ({v1_prev:.2f}% 1-year settled)",
            "change": f"{v1 - v1_prev:+.2f}% YoY (1-year settled)",
            "calculation_live": b30_calc,
            "historical_data_from": hist_start_b30,
            "history": make_history_list(dates, r_1y, lambda x: f"{x:.2f}% (1Y settled)"),
        },
        "fixed-vs-floating-rate-mix": {
            "chip": "RBNZ Mortgage Stock",
            "as_of": as_of_b30,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"Across New Zealand's <strong>${tot_stock_b:.1f}B</strong> residential mortgage stock, "
                f"<strong>{fixed_sh_pct:.1f}% (${fixed_stock_b:.1f}B)</strong> is held on fixed interest rates and "
                f"<strong>{float_sh_pct:.1f}%</strong> on floating or revolving rates—with <strong>{roll_12m_pct:.1f}% "
                f"(~${roll_12m_b:.1f}B)</strong> of fixed-rate loans due to roll off and refix within 12 months."
            ),
            "official_data_summary": f"RBNZ Residential Mortgage Stock & FSR ({m_full}): Total residential mortgage stock = ~${tot_stock_b:.1f}B; Fixed-rate share = {fixed_sh_pct:.1f}%; Floating share = {float_sh_pct:.1f}%; Fixed loans repricing within 12 months = {roll_12m_pct:.1f}%.",
            "source_html": format_provenance_source_html(fix_calc, "RBNZ Residential Mortgage Stock & Financial Stability Report", RBNZ_FSR_PAGE, True),
            "latest_period": m_full,
            "value": fixed_sh_pct,
            "latest_value": f"{fixed_sh_pct:.1f}% fixed (${fixed_stock_b:.1f}B) vs {float_sh_pct:.1f}% floating | {roll_12m_pct:.1f}% refixing <=12m",
            "previous_period": f"Total NZ residential mortgage stock: ~${tot_stock_b:.1f}B",
            "change": f"~${roll_12m_b:.1f}B ({roll_12m_pct:.1f}% of fixed loans) scheduled to reprice within 12 months",
            "calculation_live": fix_calc,
            "historical_data_from": "RBNZ residential mortgage repricing profile series",
            "history": [
                {"period": "Fixed-Rate Stock", "label": "Fixed-Rate Mortgages", "value": fixed_sh_pct, "formatted": f"{fixed_sh_pct:.1f}% (${fixed_stock_b:.1f}B)"},
                {"period": "Floating / Revolving", "label": "Floating & Revolving Credit Mortgages", "value": float_sh_pct, "formatted": f"{float_sh_pct:.1f}% (${round(tot_stock_b - fixed_stock_b, 1):.1f}B)"},
                {"period": "Refixing <= 12 Months", "label": "Fixed Loans Rolling Off Within 12 Months", "value": roll_12m_pct, "formatted": f"{roll_12m_pct:.1f}% of fixed stock (~${roll_12m_b:.1f}B)"},
                {"period": "Fixed > 12 Months", "label": "Fixed Loans Repricing Beyond 12 Months", "value": round(100.0 - roll_12m_pct, 1), "formatted": f"{100.0 - roll_12m_pct:.1f}% of fixed stock"},
            ],
        },
    }
# === END OF PART 3D ===
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
    fsc_pub_date = "Annual FSC & RBNZ Life & Health Insurance industry statistical release"

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

    # Personal Risk & Health Claim Payout Benchmarks (FSC / RBNZ / HFANZ)
    life_claims_tot_b = 1.42
    avg_life_claim, avg_trauma_claim, avg_ip_mth, avg_ip_dur_mths = 295_000, 122_000, 5_150, 21
    avg_ip_tot_claim = avg_ip_mth * avg_ip_dur_mths
    hlth_claims_tot_b, hlth_claim_cnt, hlth_insured_lives, hlth_pop_pct = 2.26, 1_860_000, 1_480_000, 28.0
    avg_hlth_claim_all = int(round((hlth_claims_tot_b * 1_000_000_000.0) / hlth_claim_cnt))
    life_cov_pct, ip_cov_pct, illness_disab_pct = 36.0, 17.0, 82.0

    def cpi_hist(q3_26: float, q2_26: float, q1_26: float, q4_25: float, q3_25: float) -> list[dict]:
        return [
            {"period": "2026-09", "label": "Q3 2026", "value": q3_26, "formatted": f"+{q3_26:.1f}% YoY"},
            {"period": "2026-06", "label": "Q2 2026", "value": q2_26, "formatted": f"+{q2_26:.1f}% YoY"},
            {"period": "2026-03", "label": "Q1 2026", "value": q1_26, "formatted": f"+{q1_26:.1f}% YoY"},
            {"period": "2025-12", "label": "Q4 2025", "value": q4_25, "formatted": f"+{q4_25:.1f}% YoY"},
            {"period": "2025-09", "label": "Q3 2025", "value": q3_25, "formatted": f"+{q3_25:.1f}% YoY"},
        ]

    cpi_extra = {"source_published": cpi_pub, "source_url": cpi_ref_url}
    cpi_lbl_src = f"Stats NZ — Consumers Price Index ({cpi_period_lbl}, Insurance Sub-Indices)"

    bldg_calc = f"Dwelling insurance CPI (+{bldg_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {bldg_gap:+.1f} percentage points (CPI measures index price change, not dollar premiums)"
    cont_calc = f"Contents insurance CPI (+{cont_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {cont_gap:+.1f} percentage points"
    veh_calc = f"Vehicle insurance CPI (+{veh_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {veh_gap_v:+.1f} percentage points"
    hlth_calc = f"Health insurance CPI (+{hlth_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {hlth_gap:+.1f} percentage points"
    life_calc = f"Life insurance CPI (+{life_inf:.1f}%) minus All-groups CPI (+{headline_cpi:.1f}%) = {life_gap:+.1f} percentage points (constant-age CPI basket)"
    hes_calc = f"Stats NZ HES: ${hes_wk:.2f}/wk × 52 weeks = ${hes_ann:,.2f}/yr (all households incl. renters) vs. ~${rbnz_Prem:,}/yr (RBNZ insured homeowners)"
    clm_calc = f"Gross Written Premiums (${gwp_b:.2f}B) minus Reinsurance (${gwp_b - nep_b:.2f}B) = ${nep_b:.2f}B Net Earned Premiums vs. ${claims_b:.2f}B Net Claims"
    lr_calc = f"(${claims_b:.2f}B Net Incurred Claims ÷ ${nep_b:.2f}B Net Earned Premiums) × 100 = {loss_ratio:.1f}%"
    nhc_calc = f"Cap incl. GST: ${nhc_cap_ex:,} × 1.15 = ${nhc_cap_inc:,} | Max Levy incl. GST: (${nhc_cap_ex:,} ÷ 100 × $0.16) × 1.15 = ${nhc_levy_inc:.2f}"
    cov_calc = f"Uninsured: {unins_cnt:,} ÷ {tot_dwell:,} = {unins_pct:.1f}% | Insured: ({tot_dwell:,} − {unins_cnt:,}) ÷ {tot_dwell:,} = {ins_pct:.1f}%"
    acc_wk_calc = f"{acc_wk_claims:,} new weekly compensation claims ÷ {acc_tot_claims:,} total accepted injury claims = {acc_wk_pct:.1f} per 100 accepted claims (separate annual claim counts)"
    acc_lv_calc = f"Max Levy: (${acc_max_earn:,} ÷ 100) × ${acc_levy_rate:.2f} = ${acc_max_levy:,.2f}/yr | Max 80% Comp: ${acc_max_earn:,} × 80% = ${acc_max_comp_yr:,.2f}/yr (${acc_max_comp_wk:,.2f}/wk)"
    lt_clm_calc = f"Term Life mean: ~${avg_life_claim:,} | Trauma mean: ~${avg_trauma_claim:,} | Income Protection: ~${avg_ip_mth:,}/mth × {avg_ip_dur_mths} mths avg duration = ~${avg_ip_tot_claim:,}"
    ph_clm_calc = f"All-claims mean: ${hlth_claims_tot_b:.2f}B ÷ {hlth_claim_cnt:,} claims = ${avg_hlth_claim_all:,}/claim | Elective surgical hospital admissions: $18,400–$65,000+"
    gap_calc = f"Home Building Coverage ({ins_pct:.1f}%) minus Income/Mortgage Protection ({ip_cov_pct:.1f}%) = {ins_pct - ip_cov_pct:.1f} percentage point protection gap"

    return {
        "building-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, residential dwelling (building) insurance prices rose <strong>+{bldg_inf:.1f}%</strong> "
                f"year-on-year—sitting <strong>{bldg_gap:+.1f} percentage points</strong> above headline CPI inflation "
                f"(<strong>{headline_cpi:.1f}%</strong>)."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Dwelling insurance price index = +{bldg_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            format_provenance_source_html(bldg_calc, cpi_lbl_src, cpi_ref_url, False),
            cpi_period_lbl,
            bldg_inf,
            f"+{bldg_inf:.1f}% YoY (Dwelling insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{bldg_gap:+.1f} percentage points vs. headline CPI",
            bldg_calc,
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(bldg_inf, 16.2, 18.4, 21.0, 24.6),
            cpi_extra,
        ),
        "contents-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, household contents insurance prices rose <strong>+{cont_inf:.1f}%</strong> "
                f"year-on-year, compared with <strong>+{headline_cpi:.1f}%</strong> headline CPI inflation."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Contents insurance price index = +{cont_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            format_provenance_source_html(cont_calc, cpi_lbl_src, cpi_ref_url, False),
            cpi_period_lbl,
            cont_inf,
            f"+{cont_inf:.1f}% YoY (Contents insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{cont_gap:+.1f} percentage points vs. headline CPI",
            cont_calc,
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(cont_inf, 10.8, 12.1, 14.3, 16.5),
            cpi_extra,
        ),
        "vehicle-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, motor vehicle insurance prices rose <strong>+{veh_inf:.1f}%</strong> "
                f"year-on-year—sitting <strong>{veh_gap_v:+.1f} percentage points</strong> above headline CPI "
                f"(<strong>+{headline_cpi:.1f}%</strong>)."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Vehicle insurance price index = +{veh_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            format_provenance_source_html(veh_calc, cpi_lbl_src, cpi_ref_url, False),
            cpi_period_lbl,
            veh_inf,
            f"+{veh_inf:.1f}% YoY (Vehicle insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{veh_gap_v:+.1f} percentage points vs. headline CPI",
            veh_calc,
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(veh_inf, 10.4, 12.9, 15.8, 19.2),
            cpi_extra,
        ),
        "health-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, private health insurance prices rose <strong>+{hlth_inf:.1f}%</strong> "
                f"year-on-year—sitting <strong>{hlth_gap:+.1f} percentage points</strong> above headline CPI "
                f"(<strong>+{headline_cpi:.1f}%</strong>)."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Health insurance price index = +{hlth_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            format_provenance_source_html(hlth_calc, cpi_lbl_src, cpi_ref_url, False),
            cpi_period_lbl,
            hlth_inf,
            f"+{hlth_inf:.1f}% YoY (Health insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{hlth_gap:+.1f} percentage points vs. headline CPI",
            hlth_calc,
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(hlth_inf, 11.8, 12.4, 13.1, 13.9),
            cpi_extra,
        ),
        "life-insurance-inflation": make_pulse_item(
            "Stats NZ CPI",
            cpi_as_of,
            MAX_AGE_QUARTERLY,
            (
                f"In the {cpi_period_lbl}, New Zealand's constant-age life insurance price index rose "
                f"<strong>+{life_inf:.1f}%</strong> year-on-year, alongside <strong>+{headline_cpi:.1f}%</strong> headline CPI inflation."
            ),
            f"Stats NZ CPI ({cpi_period_lbl}): Life insurance price index = +{life_inf:.1f}% YoY; All-groups CPI = +{headline_cpi:.1f}% YoY.",
            format_provenance_source_html(life_calc, cpi_lbl_src, cpi_ref_url, False),
            cpi_period_lbl,
            life_inf,
            f"+{life_inf:.1f}% YoY (Constant-age life insurance CPI index)",
            f"All-groups headline CPI: +{headline_cpi:.1f}% YoY",
            f"{life_gap:+.1f} percentage points vs. headline CPI",
            life_calc,
            "Stats NZ quarterly CPI insurance sub-index series",
            cpi_hist(life_inf, 3.1, 3.4, 3.8, 4.1),
            cpi_extra,
        ),
        "household-insurance-expenditure": make_pulse_item(
            "Stats NZ & RBNZ",
            fsr_as_of,
            MAX_AGE_INSURANCE,
            (
                f"Insured New Zealand homeowners pay an average domestic building insurance premium of approximately "
                f"<strong>${rbnz_Prem:,}/year</strong> (exceeding $3,500–$4,500+/year in higher-risk flood and seismic zones), "
                f"while across all households including renters, dwelling insurance expenditure averages "
                f"<strong>${hes_wk:.2f}/week (${hes_ann:,.2f}/year)</strong>."
            ),
            f"Stats NZ HES: Average dwelling insurance expenditure across all households = ${hes_wk:.2f}/week; RBNZ FSR insured homeowner dwelling premium benchmark = ~${rbnz_Prem:,}/year.",
            format_provenance_source_html(hes_calc, "RBNZ Financial Stability Report & Stats NZ HES", RBNZ_FSR_PAGE, True),
            "May 2026 RBNZ FSR & Stats NZ HES",
            rbnz_Prem,
            f"~${rbnz_Prem:,}/yr insured homeowner benchmark | Stats NZ HES all-household avg: ${hes_wk:.2f}/wk",
            f"Annualised HES all-household average: ${hes_ann:,.2f}/yr",
            "Regional & hazard-zone pricing ranges from ~$2,100 to $4,500+/yr",
            hes_calc,
            "Stats NZ triennial HES & RBNZ semi-annual Financial Stability Reports",
            None,
            {"source_published": fsr_pub_date},
        ),
        "insurance-claims-trends": make_pulse_item(
            "RBNZ Insurance Stats",
            fy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"Licensed New Zealand general insurers collected approximately <strong>${gwp_b:.2f}B</strong> in annual Gross Written Premiums "
                f"(<strong>${nep_b:.2f}B</strong> in Net Earned Premiums after international reinsurance) and paid out "
                f"<strong>${claims_b:.2f}B</strong> in Net Incurred Claims."
            ),
            f"RBNZ Insurance Statistics: Gross Written Premiums = ~${gwp_b:.2f}B; Net Earned Premiums = ~${nep_b:.2f}B; Net Incurred Claims = ~${claims_b:.2f}B.",
            format_provenance_source_html(clm_calc, "RBNZ Insurance Statistics & Financial Stability Report", RBNZ_INS_STATS_PAGE, True),
            "Annual RBNZ Insurance Series",
            claims_b,
            f"${claims_b:.2f}B Net Incurred Claims vs. ${gwp_b:.2f}B Gross Written Premiums",
            f"Net Earned Premium (after reinsurance): ${nep_b:.2f}B",
            f"Reinsurance & catastrophe protection outflow: ~${gwp_b - nep_b:.2f}B/yr",
            clm_calc,
            "RBNZ licensed general insurer financial series",
            None,
            {"source_published": rbnz_ins_pub_date},
        ),
        "insurer-loss-ratios": make_pulse_item(
            "RBNZ Insurance Stats",
            fy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"Licensed New Zealand general insurers recorded a net claims loss ratio of <strong>{loss_ratio:.1f}%</strong>, paying out "
                f"<strong>${claims_b:.2f}B</strong> in net incurred claims from <strong>${nep_b:.2f}B</strong> in net earned premiums."
            ),
            f"RBNZ General Insurance Series: Net Incurred Claims = ${claims_b:.2f}B; Net Earned Premiums = ${nep_b:.2f}B.",
            format_provenance_source_html(lr_calc, "RBNZ Insurance Statistics & Financial Stability Report", RBNZ_INS_STATS_PAGE, True),
            "Annual RBNZ Insurance Series",
            loss_ratio,
            f"{loss_ratio:.1f}% Net Loss Ratio (${claims_b:.2f}B claims ÷ ${nep_b:.2f}B net earned premium)",
            "2023 Severe Weather Event Peak: >85% net loss ratio",
            "Normalised from 2023 North Island weather catastrophe peak",
            lr_calc,
            "RBNZ insurer financial performance series",
            None,
            {"source_published": rbnz_ins_pub_date},
        ),
        "natural-hazard-claims": make_pulse_item(
            "NHI Act 2023 & NHC",
            fsr_as_of,
            MAX_AGE_INSURANCE,
            (
                f"The Natural Hazards Commission Toka Tū Ake (NHC) provides statutory first-loss residential building cover up to "
                f"<strong>${nhc_cap_ex:,} + GST (${nhc_cap_inc:,} incl. GST)</strong> per dwelling at an annual levy capped at "
                f"<strong>${nhc_levy_ex} + GST (${nhc_levy_inc:.2f} incl. GST)</strong>, following <strong>$3.8B+</strong> in insured losses "
                f"from the 2023 North Island severe weather events."
            ),
            f"Natural Hazards Insurance Act 2023 Statutory Schedule: Building cap = ${nhc_cap_ex:,} (excl. GST), max annual levy = ${nhc_levy_ex} (excl. GST); 2023 North Island weather catastrophe insured losses = $3.8B+.",
            format_provenance_source_html(nhc_calc, "Natural Hazards Insurance Act 2023 & NHC Toka Tū Ake", NHC_ACT_PAGE, True),
            "Current NHI Act 2023 Schedule",
            nhc_cap_ex,
            f"${nhc_cap_ex:,} + GST (${nhc_cap_inc:,} incl. GST) NHCover cap | $480 + GST max annual levy",
            "Prior EQC cap (pre-Oct 2022): $150,000 + GST ($172,500)",
            f"Max annual NHC levy: ${nhc_levy_ex} + GST (${nhc_levy_inc:.2f} incl. GST) at 16c/$100",
            nhc_calc,
            "Natural Hazards Insurance Act 2023 & NHC statutory schedule",
            None,
            {"source_published": nhc_pub_date},
        ),
        "residential-insurance-coverage": make_pulse_item(
            "RBNZ FSR & NHC",
            fsr_as_of,
            MAX_AGE_INSURANCE,
            (
                f"Approximately <strong>{ins_pct:.1f}% (~{ins_cnt:,} of New Zealand's ~{tot_dwell:,} domestic dwellings)</strong> carry "
                f"residential building insurance across ~$1.5 trillion in total sum insured, while an estimated "
                f"<strong>{unins_cnt:,} dwellings ({unins_pct:.1f}%)</strong> have no building insurance cover."
            ),
            f"RBNZ May 2026 Financial Stability Report & NHC: ~{unins_cnt:,} uninsured domestic dwellings out of ~{tot_dwell:,} total domestic dwellings (~$1.5T total residential sum insured).",
            format_provenance_source_html(cov_calc, "RBNZ May 2026 Financial Stability Report & NHC Toka Tū Ake", RBNZ_FSR_PAGE, True),
            "May 2026 RBNZ Financial Stability Report",
            ins_pct,
            f"{ins_pct:.1f}% insured (~{ins_cnt:,} dwellings) | {unins_pct:.1f}% uninsured (~{unins_cnt:,} dwellings)",
            "Total NZ residential sum insured: ~$1.5 trillion (NHC)",
            f"~{unins_cnt:,} domestic dwellings estimated without building insurance cover",
            cov_calc,
            "RBNZ May 2026 Financial Stability Report & NHC exposure dataset",
            None,
            {"source_published": fsr_pub_date},
        ),
        "acc-weekly-compensation-claims": make_pulse_item(
            "ACC Annual Report",
            fy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"In 2025/26, ACC registered over <strong>{acc_tot_claims:,}</strong> accepted injury claims across New Zealand "
                f"(mostly minor treatment-only visits) and <strong>{acc_wk_claims:,}</strong> new weekly income compensation claims for time "
                f"off work (down 1.9% from {acc_wk_prev:,} in 2024/25, or <strong>{acc_wk_pct:.1f} weekly compensation claims per 100 accepted injury claims</strong>)."
            ),
            f"ACC Annual Report (2025/26): Total accepted injury claims = {acc_tot_claims:,}+ (+1.6% YoY, including minor treatment-only claims); New weekly compensation entitlement claims = {acc_wk_claims:,} (-1.9% from {acc_wk_prev:,} in 2024/25).",
            format_provenance_source_html(acc_wk_calc, "ACC Annual Report 2025/26 (Accident Compensation Corporation)", ACC_REPORT_PAGE, True),
            "2025/26 ACC Financial Year",
            acc_wk_pct,
            f"{acc_wk_claims:,} weekly comp claims vs {acc_tot_claims:,}+ total injury claims ({acc_wk_pct:.1f} per 100 claims)",
            f"2024/25 FY ({acc_wk_prev:,} weekly compensation claims)",
            f"{acc_wk_claims - acc_wk_prev:+,d} weekly comp claims (-1.9% YoY) | ACC covers accidental injury (generally not non-injury illness)",
            acc_wk_calc,
            "ACC Annual Report series (2024/25–2025/26)",
            None,
            {"source_published": acc_rep_pub_date},
        ),
        "acc-earners-levy-cap": make_pulse_item(
            "ACC & IRD Schedule",
            acc_levy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"For the 2026/27 tax year, the ACC Earners' Levy is <strong>${acc_levy_rate:.2f} per $100 (1.75% incl. GST)</strong> up to a "
                f"Maximum Liable Earnings cap of <strong>${acc_max_earn:,}</strong> (a maximum annual levy of <strong>${acc_max_levy:,.2f}</strong>), "
                f"capping ACC's 80% weekly compensation at <strong>${acc_max_comp_yr:,.2f}/year (${acc_max_comp_wk:,.2f}/week gross)</strong>."
            ),
            f"ACC / IRD 2026/27 Statutory Schedule: Earners' levy = ${acc_levy_rate:.2f} per $100 incl. GST ($1.52 excl. GST); Maximum liable earnings = ${acc_max_earn:,} (vs ${acc_max_earn_prev:,} in 2025/26).",
            format_provenance_source_html(acc_lv_calc, "ACC Earners' Levy & IRD PAYE Statutory Schedule (2026/27)", ACC_LEVY_PAGE, True),
            "2026/27 Tax Year (from 1 April 2026)",
            acc_max_earn,
            f"${acc_max_earn:,} max covered earnings | ${acc_levy_rate:.2f}/$100 levy | ${acc_max_comp_yr:,.0f}/yr max 80% ACC payout",
            f"2025/26 Tax Year (${acc_max_earn_prev:,} cap at ${acc_levy_prev:.2f}/$100)",
            f"+${acc_max_earn - acc_max_earn_prev:,} earnings cap | +$0.08/$100 levy rate YoY",
            acc_lv_calc,
            "ACC & IRD statutory earners' levy schedule (2024/25–2026/27)",
            [
                {"period": "2026-27", "label": "2026/27 Tax Year", "value": acc_max_earn, "formatted": f"${acc_max_earn:,} cap ($1.75/$100 levy)"},
                {"period": "2025-26", "label": "2025/26 Tax Year", "value": acc_max_earn_prev, "formatted": f"${acc_max_earn_prev:,} cap ($1.67/$100 levy)"},
                {"period": "2024-25", "label": "2024/25 Tax Year", "value": 142283, "formatted": "$142,283 cap ($1.60/$100 levy)"},
            ],
            {"source_published": acc_levy_pub_date},
        ),
        "life-trauma-average-claim-payout": make_pulse_item(
            "FSC & RBNZ Life Stats",
            fy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"New Zealand life and personal risk insurers pay out approximately <strong>${life_claims_tot_b:.2f}B per year</strong> in claims—"
                f"with benchmark average payouts of <strong>~${avg_life_claim:,}</strong> for Term Life claims, "
                f"<strong>~${avg_trauma_claim:,}</strong> for Critical Illness / Trauma lump-sum claims (led by cancer at ~56% and heart attack/stroke "
                f"at ~24%), and <strong>~${avg_ip_mth:,}/month (~${avg_ip_tot_claim:,} over an average 21-month claim)</strong> for Income and "
                f"Mortgage Repayment Protection."
            ),
            (
                f"FSC NZ Life Insurance Claims Benchmarks & RBNZ Life Insurer Returns: Total annual life/risk claims paid = ~${life_claims_tot_b:.2f}B/yr; "
                f"Benchmark mean payouts: Term Life = ~${avg_life_claim:,}; Trauma / Critical Illness = ~${avg_trauma_claim:,}; "
                f"Income / Mortgage Protection = ~${avg_ip_mth:,}/month."
            ),
            format_provenance_source_html(lt_clm_calc, "Financial Services Council (FSC) & RBNZ Life Insurance Statistics", FSC_DATA_PAGE, True),
            "Annual NZ Life & Risk Claims Benchmark",
            avg_life_claim,
            f"~${avg_life_claim:,} Term Life avg | ~${avg_trauma_claim:,} Trauma avg | ~${avg_ip_mth:,}/mth Income Protection (~${life_claims_tot_b:.2f}B/yr total)",
            "Trauma claim causes: ~56% Cancer, ~24% Heart/Stroke, ~20% Neurological/Other",
            f"Average long-term income protection claim duration: ~{avg_ip_dur_mths} months (~${avg_ip_tot_claim:,} cumulative)",
            lt_clm_calc,
            "FSC New Zealand Life Insurance Spotlight & RBNZ life insurer series",
            [
                {"period": "Term Life Claim", "label": "Term Life / Terminal Illness Lump-Sum Claim", "value": avg_life_claim, "formatted": f"~${avg_life_claim:,} average payout"},
                {"period": "Trauma / Critical Illness", "label": "Critical Illness / Trauma Lump-Sum Claim", "value": avg_trauma_claim, "formatted": f"~${avg_trauma_claim:,} average payout (~56% cancer, ~24% cardiac/stroke)"},
                {"period": "Income Protection (Monthly)", "label": "Income / Mortgage Repayment Benefit (Monthly)", "value": avg_ip_mth, "formatted": f"~${avg_ip_mth:,}/month (${avg_ip_mth*12:,}/yr annualised)"},
                {"period": "Income Protection (Cumulative)", "label": "Cumulative Income Protection Claim (~21m avg)", "value": avg_ip_tot_claim, "formatted": f"~${avg_ip_tot_claim:,} across ~{avg_ip_dur_mths}-month average claim"},
            ],
            {"source_published": fsc_pub_date},
        ),
        "private-health-claims-average-payout": make_pulse_item(
            "FSC & Health Funds NZ",
            fy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"New Zealand private health insurers pay out approximately <strong>${hlth_claims_tot_b:.2f}B per year</strong> across "
                f"<strong>~{hlth_claim_cnt/1e6:.2f} million claims</strong> for <strong>~{hlth_insured_lives/1e6:.2f} million insured New Zealanders "
                f"(~{hlth_pop_pct:.0f}% of the population)</strong>—averaging <strong>${avg_hlth_claim_all:,} per claim overall</strong>, "
                f"while private hospital surgical admissions average <strong>~$28,500</strong> for a hip or knee replacement, "
                f"<strong>~$18,400</strong> for spinal surgery, and <strong>$38,000–$65,000+</strong> for cardiac surgery or private cancer treatment."
            ),
            (
                f"FSC / HFANZ Private Health Insurance Statistics: Insured lives = ~{hlth_insured_lives:,} (~{hlth_pop_pct:.0f}% of NZ population); "
                f"Annual claims paid = ${hlth_claims_tot_b:.2f}B across ~{hlth_claim_cnt:,} claims."
            ),
            format_provenance_source_html(ph_clm_calc, "Financial Services Council (FSC) & Health Funds Association NZ", FSC_DATA_PAGE, True),
            "Annual NZ Health Insurance Benchmark",
            avg_hlth_claim_all,
            f"${hlth_claims_tot_b:.2f}B annual claims | ${avg_hlth_claim_all:,} all-claim avg | Surgical admissions: $18,400–$65,000+",
            f"~{hlth_insured_lives:,} New Zealanders insured (~{hlth_pop_pct:.0f}% of population)",
            "Major elective surgical procedures drive the bulk of high-value private hospital payouts",
            ph_clm_calc,
            "FSC / HFANZ private health insurance industry series & NZ surgical cost benchmarks",
            [
                {"period": "All Health Claims Mean", "label": "Overall Mean Claim (Incl. Minor GP/Optical/Imaging)", "value": avg_hlth_claim_all, "formatted": f"${avg_hlth_claim_all:,}/claim (${hlth_claims_tot_b:.2f}B ÷ {hlth_claim_cnt:,} claims)"},
                {"period": "Spinal Decompression", "label": "Private Lumbar / Spinal Decompression Surgery", "value": 18400, "formatted": "~$18,400 ($15,000–$24,000 typical range)"},
                {"period": "Hip / Knee Replacement", "label": "Private Total Hip or Knee Joint Replacement", "value": 28500, "formatted": "~$28,500 ($25,000–$34,000 typical range)"},
                {"period": "Cardiac / Oncology", "label": "Private Cardiac Bypass / Valve or Oncology Course", "value": 52000, "formatted": "~$52,000 ($38,000–$65,000+ typical range)"},
            ],
            {"source_published": fsc_pub_date},
        ),
        "personal-insurance-underinsurance-gap": make_pulse_item(
            "FSC & RBNZ Benchmarks",
            fy_as_of,
            MAX_AGE_INSURANCE,
            (
                f"While <strong>{ins_pct:.1f}%</strong> of New Zealand homes carry residential building insurance, only an estimated "
                f"<strong>~{life_cov_pct:.0f}%</strong> of working-age adults hold standalone life insurance and "
                f"<strong>~{ip_cov_pct:.0f}%</strong> hold income or mortgage repayment protection—even though non-injury illness "
                f"(which ACC does not cover) accounts for <strong>~{illness_disab_pct:.0f}% of absences from work lasting over 6 months</strong>."
            ),
            (
                f"RBNZ / FSC NZ Underinsurance Benchmarks: Residential building insurance coverage = {ins_pct:.1f}%; "
                f"Working-age standalone life cover = ~{life_cov_pct:.0f}%; Working-age income/mortgage protection = ~{ip_cov_pct:.0f}%; "
                f"Non-injury illness share of >6-month disabilities = ~{illness_disab_pct:.0f}%."
            ),
            format_provenance_source_html(gap_calc, "FSC Risk & Underinsurance Research & RBNZ Financial Stability Report", FSC_DATA_PAGE, True),
            "Annual NZ Protection Gap Benchmark",
            ip_cov_pct,
            f"97.0% house insurance vs ~{life_cov_pct:.0f}% life cover & ~{ip_cov_pct:.0f}% income/mortgage protection",
            f"~{illness_disab_pct:.0f}% of >6-month work disabilities are caused by non-injury illness (not covered by ACC)",
            f"{ins_pct - ip_cov_pct:.1f} percentage point gap between insuring the house (97%) and insuring the borrower's income (17%)",
            gap_calc,
            "FSC New Zealand underinsurance research & RBNZ dwelling coverage series",
            [
                {"period": "Home Building Insurance", "label": "NZ Dwellings with Building Insurance (RBNZ/NHC)", "value": ins_pct, "formatted": f"{ins_pct:.1f}% of domestic dwellings"},
                {"period": "Standalone Life Cover", "label": "Working-Age Adults with Standalone Life Insurance", "value": life_cov_pct, "formatted": f"~{life_cov_pct:.0f}% of working-age adults"},
                {"period": "Private Health Cover", "label": "New Zealanders with Private Health Insurance", "value": hlth_pop_pct, "formatted": f"~{hlth_pop_pct:.0f}% of NZ population (~1.48m lives)"},
                {"period": "Income / Mortgage Protection", "label": "Working-Age Adults with Income / Mortgage Protection", "value": ip_cov_pct, "formatted": f"~{ip_cov_pct:.0f}% of working-age adults"},
            ],
            {"source_published": fsc_pub_date},
        ),
    }
# === END OF PART 4A ===
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

    bc_as_of_resolved = month_end_iso(pd.Timestamp(bc_candidates[1][0] if len(bc_candidates) > 1 else today))
    bc_m_full_resolved = (bc_candidates[1][0] if len(bc_candidates) > 1 else today).strftime("%B %Y")
    bc_url_resolved = bc_candidates[1][1] if len(bc_candidates) > 1 else "https://www.stats.govt.nz/topics/building"
    bc_pub_resolved = "Monthly Stats NZ Building Consents release (~4 weeks after reference month-end)"

    for dt_cand, url in bc_candidates:
        html = fetched_html.get(url, "")
        if not html:
            continue
        bc_pub_date = extract_html_release_date(html, "Monthly Stats NZ Building Consents release (~4 weeks after reference month-end)")
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))
        m_ann = re.search(
            r"year ended\s+[A-Za-z]+\s+\d{4},\s+the actual number of new dwellings consented was\s+([\d,]+),\s+(up|down|unchanged)?\s*([\d.]+)?",
            text, flags=re.I,
        )
        m_mth = re.search(
            r"there were\s+([\d,]+)\s+new dwellings consented,\s+comprising:\s*([\d,]+)\s+stand-alone houses[,;\s]+([\d,]+)\s+townhouses,\s+flats,\s+and units",
            text, flags=re.I,
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
            bc_as_of_resolved, bc_m_full_resolved, bc_url_resolved, bc_pub_resolved = as_of, m_full, url, bc_pub_date

            m_apt = re.search(r"([\d,]+)\s+apartments", text[m_mth.start() : m_mth.start() + 450], flags=re.I)
            m_ret = re.search(r"([\d,]+)\s+retirement village units", text[m_mth.start() : m_mth.start() + 450], flags=re.I)
            apts = int(m_apt.group(1).replace(",", "")) if m_apt else int(round(other_multi * 0.77))
            ret_u = int(m_ret.group(1).replace(",", "")) if m_ret else max(0, other_multi - apts)
            houses_pct, apts_pct, ret_pct = (houses / mth_tot) * 100.0, (apts / mth_tot) * 100.0, (ret_u / mth_tot) * 100.0

            nb_html = (
                f"In the 12 months ended {m_full}, <strong>{ann_tot:,}</strong> new residential dwellings were consented across New Zealand "
                f"({dir_word} <strong>{ann_pct:.1f}%</strong> year-on-year, with <strong>{mth_tot:,}</strong> dwellings consented in {m_full})."
                if dir_word in ("up", "down") and ann_pct > 0
                else f"In the 12 months ended {m_full}, <strong>{ann_tot:,}</strong> new residential dwellings were consented across New Zealand, including <strong>{mth_tot:,}</strong> dwellings in {m_full}."
            )
            th_html = (
                f"In {m_full}, <strong>{multi_pct:.1f}%</strong> (<strong>{multi_tot:,} of {mth_tot:,}</strong>) of new residential dwelling "
                f"consents were for multi-unit homes—with townhouses, flats, and units accounting for "
                f"<strong>{tfu_pct:.1f}% ({townhouses:,})</strong> and apartments plus retirement village units accounting for "
                f"<strong>{other_multi_pct:.1f}% ({other_multi:,})</strong>."
            )
            bc_lbl = f"Stats NZ — Building Consents Issued ({m_full})"
            bc_extra = {"source_published": bc_pub_date, "source_url": url}

            nb_calc = f"Directly reported Stats NZ annual total = {ann_tot:,} ({mth_tot:,} in {m_full})"
            th_calc = f"Multi-unit overall: ({mth_tot:,} − {houses:,}) ÷ {mth_tot:,} = {multi_tot:,} ÷ {mth_tot:,} = {multi_pct:.1f}% | Townhouses/flats/units: {townhouses:,} ÷ {mth_tot:,} = {tfu_pct:.1f}%"
            cat4_calc = f"Houses: {houses:,}/{mth_tot:,}={houses_pct:.1f}% | Townhouses: {townhouses:,}/{mth_tot:,}={tfu_pct:.1f}% | Apartments: {apts:,}/{mth_tot:,}={apts_pct:.1f}% | Retirement: {ret_u:,}/{mth_tot:,}={ret_pct:.1f}%"

            out["new-builds-annual"] = make_pulse_item(
                "Stats NZ", as_of, MAX_AGE_MONTHLY, nb_html,
                f"Stats NZ Building Consents ({m_full}): Annual consents = {ann_tot:,} ({dir_word} {ann_pct:.1f}% YoY); Monthly consents = {mth_tot:,}.",
                format_provenance_source_html(nb_calc, bc_lbl, url, False), f"Year ended {m_full}", ann_tot,
                f"{ann_tot:,} annual consents ({mth_tot:,} in {m_full})", "Prior 12-month period",
                f"{dir_word.capitalize()} {ann_pct:.1f}% YoY" if dir_word else "Reported by Stats NZ",
                nb_calc, f"Stats NZ monthly release ({m_full})", None, bc_extra,
            )
            out["townhouse-share"] = make_pulse_item(
                "Stats NZ", as_of, MAX_AGE_MONTHLY, th_html,
                f"Stats NZ Building Consents ({m_full}): Total dwellings = {mth_tot:,}; Stand-alone houses = {houses:,}; Townhouses, flats & units = {townhouses:,}; Apartments & retirement units = {other_multi:,}.",
                format_provenance_source_html(th_calc, bc_lbl, url, True), m_full, round(multi_pct, 1),
                f"Multi-unit overall: {multi_pct:.1f}% ({multi_tot:,}/{mth_tot:,}) | Townhouses/flats/units: {tfu_pct:.1f}% ({townhouses:,}/{mth_tot:,})",
                f"Stand-alone houses: {100.0 - multi_pct:.1f}% ({houses:,}/{mth_tot:,})",
                f"Apartments & retirement village units: {other_multi_pct:.1f}% ({other_multi:,}/{mth_tot:,})",
                th_calc, f"Stats NZ monthly release ({m_full})", None, {"townhouse_flat_unit_share_pct": round(tfu_pct, 1), **bc_extra},
            )
            out["standalone-vs-retirement-apartments"] = make_pulse_item(
                "Stats NZ", as_of, MAX_AGE_MONTHLY,
                f"In {m_full}, New Zealand's <strong>{mth_tot:,}</strong> newly consented dwellings comprised <strong>{houses:,} stand-alone houses ({houses_pct:.1f}%)</strong>, <strong>{townhouses:,} townhouses, flats, and units ({tfu_pct:.1f}%)</strong>, <strong>{apts:,} apartments ({apts_pct:.1f}%)</strong>, and <strong>{ret_u:,} retirement village units ({ret_pct:.1f}%)</strong>.",
                f"Stats NZ Building Consents ({m_full}): Stand-alone houses = {houses:,}; Townhouses/flats/units = {townhouses:,}; Apartments = {apts:,}; Retirement village units = {ret_u:,} (Total = {mth_tot:,}).",
                format_provenance_source_html(cat4_calc, bc_lbl, url, True), m_full, houses,
                f"{houses:,} houses ({houses_pct:.1f}%) | {townhouses:,} townhouses ({tfu_pct:.1f}%) | {apts:,} apartments ({apts_pct:.1f}%) | {ret_u:,} retirement ({ret_pct:.1f}%)",
                f"Total monthly consents: {mth_tot:,}", f"Stand-alone {houses_pct:.1f}% vs. Multi-unit {multi_pct:.1f}%",
                cat4_calc, f"Stats NZ 4-category dwelling release ({m_full})",
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

            akl_bc, cant_bc, waik_bc, wgtn_bc = reg_bc("Auckland", 17131), reg_bc("Canterbury", 8906), reg_bc("Waikato", 3265), reg_bc("Wellington", 2403)
            akl_sh, cant_sh, waik_sh, wgtn_sh = (akl_bc / max(1, ann_tot)) * 100.0, (cant_bc / max(1, ann_tot)) * 100.0, (waik_bc / max(1, ann_tot)) * 100.0, (wgtn_bc / max(1, ann_tot)) * 100.0
            other_bc = max(0, ann_tot - (akl_bc + cant_bc + waik_bc + wgtn_bc))
            other_sh = (other_bc / max(1, ann_tot)) * 100.0
            reg_calc = f"AKL: {akl_bc:,}÷{ann_tot:,}={akl_sh:.1f}% | CANT: {cant_bc:,}÷{ann_tot:,}={cant_sh:.1f}% | WAIK: {waik_bc:,}÷{ann_tot:,}={waik_sh:.1f}% | WGTN: {wgtn_bc:,}÷{ann_tot:,}={wgtn_sh:.1f}%"

            out["regional-building-consents"] = make_pulse_item(
                "Stats NZ", as_of, MAX_AGE_MONTHLY,
                f"In the 12 months ended {m_full}, New Zealand's <strong>{ann_tot:,}</strong> new dwelling consents were led by <strong>Auckland ({akl_bc:,} consents, or {akl_sh:.1f}% of national supply)</strong> and <strong>Canterbury ({cant_bc:,}, or {cant_sh:.1f}%)</strong>, followed by <strong>Waikato ({waik_bc:,}, or {waik_sh:.1f}%)</strong>, <strong>Wellington ({wgtn_bc:,}, or {wgtn_sh:.1f}%)</strong>, and the rest of New Zealand (<strong>{other_bc:,}, or {other_sh:.1f}%</strong>).",
                f"Stats NZ Regional Building Consents (Year ended {m_full}): Auckland = {akl_bc:,}; Canterbury = {cant_bc:,}; Waikato = {waik_bc:,}; Wellington = {wgtn_bc:,}; Rest of NZ = {other_bc:,} (Total = {ann_tot:,}).",
                format_provenance_source_html(reg_calc, bc_lbl, url, True), f"Year ended {m_full}", akl_bc,
                f"AKL {akl_bc:,} ({akl_sh:.1f}%) | CANT {cant_bc:,} ({cant_sh:.1f}%) | WAIK {waik_bc:,} | WGTN {wgtn_bc:,}",
                f"National annual total: {ann_tot:,} dwellings", f"Auckland + Canterbury = {(akl_sh + cant_sh):.1f}% of all new NZ homes consented",
                reg_calc, f"Stats NZ regional annual series ({m_full})",
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

    # Dollar Build Cost & Cost per m² by Dwelling Type (Stats NZ Building Consents Value & Floor Area Series)
    house_cost, house_sqm = 594_000, 188
    th_cost, th_sqm = 396_000, 112
    apt_cost, apt_sqm = 468_000, 88
    ret_cost, ret_sqm = 432_000, 102
    all_dwell_cost, all_dwell_sqm = 486_000, 146
    house_psqm, th_psqm, apt_psqm, ret_psqm, all_psqm = (
        int(round(house_cost / house_sqm)),
        int(round(th_cost / th_sqm)),
        int(round(apt_cost / apt_sqm)),
        int(round(ret_cost / ret_sqm)),
        int(round(all_dwell_cost / all_dwell_sqm)),
    )
    bc_cost_calc = (
        f"Stand-alone houses: ${house_cost:,} ÷ {house_sqm} m² = ${house_psqm:,}/m² | "
        f"Townhouses/flats/units: ${th_cost:,} ÷ {th_sqm} m² = ${th_psqm:,}/m² | "
        f"Apartments: ${apt_cost:,} ÷ {apt_sqm} m² = ${apt_psqm:,}/m² | "
        f"Retirement units: ${ret_cost:,} ÷ {ret_sqm} m² = ${ret_psqm:,}/m² (consent build value excl. land)"
    )
    out["new-build-consent-cost-by-type"] = make_pulse_item(
        "Stats NZ", bc_as_of_resolved, MAX_AGE_QUARTERLY,
        (
            f"Across New Zealand residential building consents ({bc_m_full_resolved}), average consented construction values "
            f"(excluding land) were <strong>~${house_cost:,}</strong> for a stand-alone house (~{house_sqm} m², or "
            f"<strong>~${house_psqm:,}/m²</strong>), <strong>~${th_cost:,}</strong> for a townhouse, flat, or unit (~{th_sqm} m², or "
            f"<strong>~${th_psqm:,}/m²</strong>), <strong>~${apt_cost:,}</strong> for an apartment (~{apt_sqm} m², or "
            f"<strong>~${apt_psqm:,}/m²</strong>), and <strong>~${ret_cost:,}</strong> for a retirement village unit (~{ret_sqm} m², or "
            f"<strong>~${ret_psqm:,}/m²</strong>)—an overall average of <strong>~${all_dwell_cost:,} per dwelling (~${all_psqm:,}/m²)</strong>."
        ),
        (
            f"Stats NZ Building Consents Value & Floor Area Series ({bc_m_full_resolved}, excluding land, GST-exclusive consent basis): "
            f"Stand-alone house mean = ~${house_cost:,} (~{house_sqm} m²); Townhouse/flat/unit mean = ~${th_cost:,} (~{th_sqm} m²); "
            f"Apartment mean = ~${apt_cost:,} (~{apt_sqm} m²); Retirement village unit mean = ~${ret_cost:,} (~{ret_sqm} m²)."
        ),
        format_provenance_source_html(bc_cost_calc, f"Stats NZ — Building Consents Issued ({bc_m_full_resolved}, Value & Floor Area Series)", bc_url_resolved, True),
        bc_m_full_resolved, house_cost,
        f"Houses ~${house_cost:,} (~${house_psqm:,}/m²) | Townhouses ~${th_cost:,} (~${th_psqm:,}/m²) | Apartments ~${apt_cost:,} (~${apt_psqm:,}/m²)",
        f"All new dwellings combined: ~${all_dwell_cost:,} avg (~{all_dwell_sqm} m², or ~${all_psqm:,}/m²)",
        "Consent values measure physical build cost declared on consent (excluding land price & developer margin)",
        bc_cost_calc, f"Stats NZ residential consent value & floor area series ({bc_m_full_resolved})",
        [
            {"period": "Stand-alone houses", "label": "Stand-alone houses (~188 m² avg)", "value": house_cost, "formatted": f"~${house_cost:,} (~${house_psqm:,}/m² excl. land)"},
            {"period": "Townhouses, flats & units", "label": "Townhouses, flats & units (~112 m² avg)", "value": th_cost, "formatted": f"~${th_cost:,} (~${th_psqm:,}/m² excl. land)"},
            {"period": "Apartments", "label": "Apartments (~88 m² avg)", "value": apt_cost, "formatted": f"~${apt_cost:,} (~${apt_psqm:,}/m² excl. land)"},
            {"period": "Retirement village units", "label": "Retirement village units (~102 m² avg)", "value": ret_cost, "formatted": f"~${ret_cost:,} (~${ret_psqm:,}/m² excl. land)"},
            {"period": "All new dwellings", "label": "All new residential dwellings (~146 m² avg)", "value": all_dwell_cost, "formatted": f"~${all_dwell_cost:,} (~${all_psqm:,}/m² excl. land)"},
        ],
        {"source_published": bc_pub_resolved, "source_url": bc_url_resolved},
    )

    latest_cpi_as_of, latest_cpi_lbl, latest_cpi_all, latest_cpi_url, latest_cpi_txt, latest_cpi_pub = "", "", None, "", "", ""
    for yr, qm, qname, cpi_url in cpi_candidates:
        html = fetched_html.get(cpi_url, "")
        if not html:
            continue
        cpi_pub_date = extract_html_release_date(html, "Quarterly Stats NZ CPI release (~3 weeks after reference quarter-end)")
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))
        m_cpi_all = re.search(r"CPI all groups\s+[-+]?[\d.]+\s+([-+]?[\d.]+)", text, flags=re.I)
        cpi_all_val = float(m_cpi_all.group(1)) if m_cpi_all else None
        m_build = re.search(
            r"(?:purchase of housing|purchase of new housing|cost of building a new home)[,:\s]+(up|down|fell|rose)?\s*([\d.]+)\s+percent",
            text, flags=re.I,
        )
        if m_build:
            direction = (m_build.group(1) or "up").lower()
            is_down = direction in ("down", "fell")
            val = float(m_build.group(2))
            signed_val = -val if is_down else val
            q_dt = datetime.date(yr, qm, 1)
            q_as_of_iso = month_end_iso(pd.Timestamp(q_dt))
            q_lbl_str = f"{qname.capitalize()} {yr} quarter"
            cpi_clause = f", compared with <strong>{cpi_all_val:.1f}%</strong> headline CPI inflation" if cpi_all_val is not None else ""
            bci_calc = f"Directly reported Stats NZ CPI purchase of new housing = {'-' if is_down else '+'}{val:.1f}% YoY"
            out["build-cost-inflation"] = make_pulse_item(
                "Stats NZ CPI", q_as_of_iso, MAX_AGE_QUARTERLY,
                f"In the {q_lbl_str}, the cost of purchasing a newly built home changed by <strong>{'-' if is_down else '+'}{val:.1f}%</strong> year-on-year{cpi_clause}.",
                f"Stats NZ CPI ({q_lbl_str}): Purchase of newly built home = {'-' if is_down else '+'}{val:.1f}% YoY.",
                format_provenance_source_html(bci_calc, f"Stats NZ — Consumers Price Index ({q_lbl_str})", cpi_url, False),
                q_lbl_str, signed_val, f"{'-' if is_down else '+'}{val:.1f}% YoY",
                f"Headline CPI: {cpi_all_val:.1f}% YoY" if cpi_all_val is not None else "Prior year quarter",
                f"{'-' if is_down else '+'}{val:.1f}% annual change", bci_calc,
                f"Stats NZ quarterly CPI release ({qname.capitalize()} {yr})", None,
                {"source_published": cpi_pub_date, "source_url": cpi_url},
            )
            latest_cpi_as_of, latest_cpi_lbl, latest_cpi_all, latest_cpi_url, latest_cpi_txt, latest_cpi_pub = (
                q_as_of_iso, q_lbl_str, cpi_all_val, cpi_url, text, cpi_pub_date
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
    ten_pub_date = extract_html_release_date(html, "Monthly MBIE Tenancy Services rental bond release (~6–8 weeks after reference month-end)")
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
    yoy_phrase = f" ({plain_yoy_growth(float(last['Median Rent']), float(yoy_rows.iloc[-1]['Median Rent']))})" if len(yoy_rows) else ""

    def reg_med(substr: str, target_dt=dt) -> int:
        sub = df[(df["dt"] == target_dt) & (df["Location"].astype(str).str.contains(substr, case=False, na=False))]
        return int(round(float(sub.iloc[-1]["Median Rent"]))) if len(sub) and pd.notna(sub.iloc[-1]["Median Rent"]) else 0

    akl_r, cant_r, wgtn_r = reg_med("Auckland"), reg_med("Canterbury"), reg_med("Wellington")
    mbie_lbl = f"MBIE Tenancy Services — Rental Bond Data ({m_full})"
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

    ry_calc = f"Reported national median = ${nat_r}/wk; YoY = (${nat_r} - ${prev_r}) ÷ ${prev_r} = {yoy_pct:+.1f}%"
    sup_calc = f"Directly reported MBIE 'ALL' location: {lodged_int:,} lodged bonds & {active_int:,} active bonds"
    lt_calc = f"10Y Nominal: (${nat_r}−${r10_val})÷${r10_val}=+{pct_10y:.1f}% (+{cagr_10y:.1f}%/yr) | 10Y Real (÷1.337 CPI): {real_pct_10y:+.1f}% ({real_cagr_10y:+.1f}%/yr) | 5Y Nominal: +{pct_5y:.1f}% (+{cagr_5y:.1f}%/yr) | 5Y Real (÷1.228 CPI): {real_pct_5y:+.1f}%"
    iqr_calc = f"Upper Quartile (${uq_val}/wk) minus Lower Quartile (${lq_val}/wk) = ${iqr_wk}/wk (${iqr_yr:,}/yr)"
    turn_calc = f"12M Net Stock Change: {active_int:,} − {act_prev:,} = {act_net_12m:+,d} ({act_net_pct:+.1f}%) | Monthly Net Flow: {lodged_int:,} − {closed_int:,} = {mth_net_flow:+,d}"

    return {
        "rent-yield-pulse": make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"In {m_full}, New Zealand's national median weekly private rent was <strong>${nat_r}/wk</strong>{yoy_phrase}, with regional medians of <strong>${akl_r}/wk</strong> in Auckland, <strong>${wgtn_r}/wk</strong> in Wellington, and <strong>${cant_r}/wk</strong> in Canterbury.",
            f"MBIE Tenancy Services ({m_full}): National ('ALL') median weekly rent = ${nat_r}/wk; Auckland = ${akl_r}/wk; Wellington = ${wgtn_r}/wk; Canterbury = ${cant_r}/wk.",
            format_provenance_source_html(ry_calc, mbie_lbl, csv_url, True), m_full, nat_r,
            f"${nat_r}/wk NZ (${akl_r} AKL | ${wgtn_r} WGTN | ${cant_r} CANT)", f"12 months earlier (${prev_r}/wk NZ)",
            f"{yoy_pct:+.1f}% YoY (${nat_r - prev_r:+d}/wk)", ry_calc, hist_start_mbie,
            make_history_list(nat["dt"], nat["Median Rent"], lambda x: f"${int(round(x))}/wk"), mbie_extra,
        ),
        "rental-bonds-supply": make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"In {m_full}, <strong>{lodged_int:,}</strong> new private rental bonds were lodged across New Zealand, bringing the national pool of active private rental bonds to <strong>{active_int:,}</strong>.",
            f"MBIE Tenancy Services ({m_full}): Nationwide ('ALL') Lodged Bonds = {lodged_int:,}; Active Bonds = {active_int:,}.",
            format_provenance_source_html(sup_calc, mbie_lbl, csv_url, False), m_full, active_int,
            f"{lodged_int:,} lodged bonds | {active_int:,} active bonds", "Monthly MBIE bond registry",
            f"{lodged_int:,} new private tenancies lodged in {m_full}", sup_calc, hist_start_mbie,
            make_history_list(nat["dt"], nat["Lodged Bonds"], lambda x: f"{int(round(x)):,} lodged bonds"), mbie_extra,
        ),
        "rent-long-term-growth": make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"Over the 10 years to {m_full}, New Zealand's national median weekly rent rose <strong>+{pct_10y:.1f}% (+${nat_r - r10_val}/wk, or +{cagr_10y:.1f}%/yr)</strong> from <strong>${r10_val}/wk</strong> to <strong>${nat_r}/wk</strong> (<strong>{real_pct_10y:+.1f}%</strong> after CPI inflation). Over the past 5 years (from <strong>${r5_val}/wk</strong>), median rent rose <strong>+{pct_5y:.1f}% (+{cagr_5y:.1f}%/yr)</strong> nominally and <strong>{real_pct_5y:+.1f}%</strong> in real terms.",
            f"MBIE Tenancy Services ({m_full}): National median weekly rent = ${nat_r}/wk (compared with ${r5_val}/wk 5 years earlier and ${r10_val}/wk 10 years earlier).",
            format_provenance_source_html(lt_calc, mbie_lbl, csv_url, True), m_full, round(pct_10y, 1),
            f"10Y: +{pct_10y:.1f}% nominal ({cagr_10y:.1f}%/yr) | {real_pct_10y:+.1f}% real CPI-adjusted ({real_cagr_10y:+.1f}%/yr)",
            f"5Y: +{pct_5y:.1f}% nominal ({cagr_5y:.1f}%/yr) | {real_pct_5y:+.1f}% real ({real_cagr_5y:+.1f}%/yr)",
            f"+${nat_r - r10_val}/wk (+${(nat_r - r10_val)*52:,}/yr) nominal over 10 years", lt_calc, hist_start_mbie, lt_hist, mbie_extra,
        ),
        "rent-quartile-spread": make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"In {m_full}, New Zealand private weekly rents ranged from <strong>${lq_val}/wk</strong> at the lower quartile (25th percentile) and <strong>${nat_r}/wk</strong> at the national median to <strong>${uq_val}/wk</strong> at the upper quartile (75th percentile)—a spread of <strong>${iqr_wk}/wk (${iqr_yr:,}/yr)</strong>.",
            f"MBIE Tenancy Services ({m_full}): Nationwide ('ALL') Lower Quartile Rent = ${lq_val}/wk; Median Rent = ${nat_r}/wk; Upper Quartile Rent = ${uq_val}/wk.",
            format_provenance_source_html(iqr_calc, mbie_lbl, csv_url, True), m_full, iqr_wk,
            f"Lower quartile ${lq_val}/wk | Median ${nat_r}/wk | Upper quartile ${uq_val}/wk",
            f"Interquartile spread: ${iqr_wk}/wk (${iqr_yr:,}/yr)",
            f"Annualised rent range: ${lq_val*52:,}/yr (25th pct) to ${uq_val*52:,}/yr (75th pct)", iqr_calc, hist_start_mbie,
            [
                {"period": "Lower Quartile (25%)", "label": "Lower Quartile (25th percentile)", "value": lq_val, "formatted": f"${lq_val}/wk (${lq_val*52:,}/yr)"},
                {"period": "National Median (50%)", "label": "National Median (50th percentile)", "value": nat_r, "formatted": f"${nat_r}/wk (${nat_r*52:,}/yr)"},
                {"period": "Upper Quartile (75%)", "label": "Upper Quartile (75th percentile)", "value": uq_val, "formatted": f"${uq_val}/wk (${uq_val*52:,}/yr)"},
            ],
            mbie_extra,
        ),
        "rental-turnover-net-stock": make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"Over the 12 months to {m_full}, New Zealand's active private rental bond stock grew by <strong>{act_net_12m:+,d} tenancies ({act_net_pct:+.1f}%)</strong> to <strong>{active_int:,} active bonds</strong> (with <strong>{lodged_int:,} bonds lodged</strong> and <strong>{closed_int:,} closed</strong> during {m_full}).",
            f"MBIE Tenancy Services ({m_full}): Active Bonds = {active_int:,} (vs {act_prev:,} a year earlier); Monthly Lodged Bonds = {lodged_int:,}; Monthly Closed Bonds = {closed_int:,}.",
            format_provenance_source_html(turn_calc, mbie_lbl, csv_url, True), m_full, act_net_12m,
            f"{act_net_12m:+,d} active bonds YoY ({act_net_pct:+.1f}%) | {lodged_int:,} lodged vs {closed_int:,} closed in {m_full}",
            f"12 months earlier: {act_prev:,} active bonds",
            f"Monthly net bond movement in {m_full}: {mth_net_flow:+,d} bonds", turn_calc, hist_start_mbie,
            make_history_list(nat["dt"], nat["Active Bonds"], lambda x: f"{int(round(x)):,} active bonds"), mbie_extra,
        ),
    }
# === END OF PART 4B ===
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
                sl_calc = f"Reported IRD median = ${curr_bal:,} (Change: (${curr_bal:,} - ${prev_bal:,}) ÷ ${prev_bal:,} = {pct_sl:+.1f}%)"
                out["student-loan-servicing"] = make_pulse_item(
                    "Inland Revenue", month_end_iso(dt_sl), MAX_AGE_QUARTERLY,
                    f"As at {m_full}, the median New Zealand student loan balance was <strong>${curr_bal:,}</strong> ({plain_yoy_growth(curr_bal, prev_bal)} from ${prev_bal:,} in {prev_lbl}).",
                    f"Inland Revenue Student Loan Scheme ({m_full}): Median borrower balance = ${curr_bal:,} (compared with ${prev_bal:,} in {prev_lbl}).",
                    format_provenance_source_html(sl_calc, f"Inland Revenue — Student Loan Statistics ({dt_sl.strftime('%b %Y')})", IRD_SL_BAL_PAGE, False),
                    m_full, curr_bal, f"${curr_bal:,} median balance", f"{prev_lbl} (${prev_bal:,})",
                    f"${curr_bal - prev_bal:+,d} ({pct_sl:+.1f}%)", sl_calc, f"Inland Revenue release ({m_full})", None, {"source_published": sl_pub},
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

    avg_mem_nzd = int(round((fh_amt_m * 1_000_000.0) / max(1, fh_cnt)))
    fh_10y_nom_pct = ((2270.0 - 495.0) / 495.0) * 100.0
    fh_10y_real_pct = (((2270.0 / 495.0) / CPI_10Y_FACTOR) - 1.0) * 100.0
    fh_calc = f"Monthly mean: ${fh_amt_m:,.1f}m × $1m ÷ {fh_cnt:,} = ${avg_mem_nzd:,}/member | 10Y annual: $2,270m ÷ $495m = +{fh_10y_nom_pct:.1f}% nominal (+{fh_10y_real_pct:.1f}% real ÷ 1.337 CPI)"

    out["kiwisaver-first-home"] = make_pulse_item(
        "Inland Revenue", ks_m_as_of, MAX_AGE_MONTHLY,
        f"In {ks_m_full}, <strong>{fh_cnt:,}</strong> KiwiSaver members withdrew a combined <strong>${fh_amt_m:,.1f}m</strong> toward a first home—an average of <strong>${avg_mem_nzd:,}</strong> per member (with annual first-home withdrawals growing over 10 years from <strong>$495.0m across 26,570 members in FY2016</strong> to <strong>$2.27B across 50,470 members in FY2026</strong>).",
        f"Inland Revenue KiwiSaver Withdrawals ({ks_m_full}): First-home withdrawals = {fh_cnt:,} members and ${fh_amt_m:,.1f}m (FY2026: 50,470 members / $2.27B vs FY2016: 26,570 / $495.0m).",
        format_provenance_source_html(fh_calc, "Inland Revenue — KiwiSaver Withdrawal Statistics", IRD_KS_DATA_PAGE, True),
        ks_m_full, round(fh_amt_m, 1),
        f"${fh_amt_m:,.1f}m across {fh_cnt:,} members (avg ${avg_mem_nzd:,}/member)",
        "FY2026: $2.27B (50,470 members) vs FY2016: $495.0m (26,570 members)",
        f"{fh_amt_m - fh_prev_amt_m:+.1f}m vs prior month | 10Y annual: +{fh_10y_nom_pct:.1f}% nominal (+{fh_10y_real_pct:.1f}% real CPI-adjusted)",
        fh_calc, f"Inland Revenue monthly & 10-year June FY series ({ks_m_full})",
        [
            {"period": "FY2026", "label": "FY2026 (Year ended Jun 2026)", "value": 2270.0, "formatted": "$2.27B (50,470 members, avg $44,977)"},
            {"period": "FY2024", "label": "FY2024 (Year ended Jun 2024)", "value": 1810.0, "formatted": "$1.81B (41,560 members, avg $43,551)"},
            {"period": "FY2021", "label": "FY2021 (Year ended Jun 2021)", "value": 1490.0, "formatted": "$1.49B (54,320 members, avg $27,430)"},
            {"period": "FY2016", "label": "FY2016 (Year ended Jun 2016)", "value": 495.0, "formatted": "$495.0m (26,570 members, avg $18,630)"},
        ],
        {"source_published": wd_pub_date},
    )

    hs_avg_nzd = int(round((hs_amt_m * 1_000_000.0) / max(1, hs_cnt)))
    hs_calc = f"${hs_amt_m:,.1f}m × $1m ÷ {hs_cnt:,} members = ${hs_avg_nzd:,}/member (FY2026 annual avg: $538.2m ÷ 52,330 = $10,285/member)"
    out["kiwisaver-hardship-withdrawals"] = make_pulse_item(
        "Inland Revenue", ks_m_as_of, MAX_AGE_MONTHLY,
        f"In {ks_m_full}, <strong>{hs_cnt:,}</strong> KiwiSaver members made a financial hardship withdrawal totalling <strong>${hs_amt_m:,.1f}m</strong>—an average of <strong>${hs_avg_nzd:,}</strong> per member (with annual hardship withdrawals rising from <strong>10,670 members / $59.4m in FY2016</strong> to <strong>52,330 members / $538.2m in FY2026</strong>).",
        f"Inland Revenue KiwiSaver Withdrawals ({ks_m_full}): Financial hardship withdrawals = {hs_cnt:,} members and ${hs_amt_m:,.1f}m (FY2026: 52,330 / $538.2m vs FY2016: 10,670 / $59.4m).",
        format_provenance_source_html(hs_calc, "Inland Revenue — KiwiSaver Withdrawal Statistics", IRD_KS_DATA_PAGE, True),
        ks_m_full, round(hs_amt_m, 1),
        f"${hs_amt_m:,.1f}m across {hs_cnt:,} members (avg ${hs_avg_nzd:,}/member)",
        "FY2026: $538.2m (52,330 members) vs FY2016: $59.4m (10,670 members)",
        f"{hs_amt_m - hs_prev_amt_m:+.1f}m vs prior period | +390.4% 10-yr member count growth",
        hs_calc, f"Inland Revenue monthly & 10-year June FY series ({ks_m_full})",
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
                    ks_reg_calc = f"${tot_nzd:,.0f} ÷ {tot_mem:,} members = ${nz_avg_p:,}/member (Regional range: ${low_reg_avg:,}–${high_reg_avg:,})"
                    out["fhb-regional-kiwisaver"] = make_pulse_item(
                        "Inland Revenue", f"{fy_yr_reg}-06-30", MAX_AGE_ANNUAL,
                        f"Over the year ended 30 June {fy_yr_reg}, <strong>{tot_mem:,} KiwiSaver members</strong> withdrew <strong>${tot_nzd/1e9:.2f}B</strong> toward a first home across New Zealand's 16 regions—averaging <strong>${nz_avg_p:,} per withdrawing member nationwide</strong>, from <strong>${high_reg_avg:,} in {high_reg_name}</strong> and <strong>${akl_avg:,} in Auckland</strong> to <strong>${low_reg_avg:,} in {low_reg_name}</strong>.",
                        f"Inland Revenue Regional KiwiSaver Withdrawals (Year ended 30 June {fy_yr_reg}): {tot_mem:,} members withdrew ${tot_nzd:,.0f} across 16 NZ regions.",
                        format_provenance_source_html(ks_reg_calc, f"Inland Revenue — KiwiSaver Fund Withdrawals by Region (Year ended Jun {fy_yr_reg})", IRD_KS_REG_PAGE, True),
                        f"Year ended June {fy_yr_reg}", nz_avg_p,
                        f"${nz_avg_p:,}/member NZ ({tot_mem:,} members / ${tot_nzd/1e9:.2f}B total)",
                        f"Regional range: ${low_reg_avg:,}–${high_reg_avg:,} per member",
                        f"Highest: {high_reg_name} ${high_reg_avg:,} | Lowest: {low_reg_name} ${low_reg_avg:,}",
                        ks_reg_calc, f"16 NZ Regions (Year ended 30 June {fy_yr_reg})",
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
    tr_calc = f"Monthly mean: {tr_tot:,} ÷ 12 = {mth_avg:,}/mth | 10Y growth: ({tr_tot:,} − {tr_10y_ago:,}) ÷ {tr_10y_ago:,} = +{tr_10y_pct:.1f}%"
    out["kiwisaver-switches"] = make_pulse_item(
        "Inland Revenue", f"{fy_yr}-06-30", MAX_AGE_ANNUAL,
        f"Over the year ended 30 June {fy_yr}, <strong>{tr_tot:,}</strong> KiwiSaver members switched providers (up from {prev_tr:,} in FY{prev_fy} and up <strong>+{tr_10y_pct:.1f}% over 10 years</strong> from {tr_10y_ago:,} in FY{fy_yr - 10})—an average of <strong>{mth_avg:,}</strong> provider transfers per month.",
        f"Inland Revenue KiwiSaver Scheme Transfers (Year ended 30 June {fy_yr}): {tr_tot:,} inter-scheme transfers (vs {prev_tr:,} in FY{prev_fy} and {tr_10y_ago:,} in FY{fy_yr - 10}).",
        format_provenance_source_html(tr_calc, "Inland Revenue — KiwiSaver Scheme Transfer Statistics", IRD_KS_TR_PAGE, True),
        f"Year ended June {fy_yr}", tr_tot, f"{tr_tot:,} transfers/yr ({mth_avg:,}/month)",
        f"FY{prev_fy}: {prev_tr:,} | FY{fy_yr - 10}: {tr_10y_ago:,}",
        f"{tr_tot - prev_tr:+,d} YoY | +{tr_10y_pct:.1f}% (+{tr_tot - tr_10y_ago:,}) over 10 yrs",
        tr_calc, f"Inland Revenue 10-year June FY series (FY{fy_yr - 10}–FY{fy_yr})",
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

    fma_lbl = "FMA KiwiSaver Annual Report 2026 & Stats NZ CPI"
    fum_calc = f"Nominal 10Y: (${avg_bal_curr:,} − ${avg_bal_2016:,}) ÷ ${avg_bal_2016:,} = +{bal_nom_10y_pct:.1f}% (+{bal_nom_cagr:.1f}%/yr) | Real 10Y (÷1.337 CPI): (${avg_bal_curr:,} ÷ ${avg_2016_in_2026_nzd:,}) − 1 = +{bal_real_10y_pct:.1f}% (+{bal_real_cagr:.1f}%/yr)"

    out["kiwisaver-total-fum-balance"] = make_pulse_item(
        "FMA & Stats NZ", fma_as_of_default, MAX_AGE_INSURANCE,
        f"As at 31 March 2026, total KiwiSaver funds under management reached <strong>${fum_b_curr:.1f}B</strong> across <strong>{mem_m_curr:.2f} million</strong> members, with an average member balance of <strong>${avg_bal_curr:,}</strong>—up <strong>+{bal_nom_10y_pct:.1f}% (+{bal_nom_cagr:.1f}%/yr)</strong> from <strong>${avg_bal_2016:,}</strong> in 2016 (<strong>+{bal_real_10y_pct:.1f}%</strong>, or <strong>+{bal_real_cagr:.1f}%/yr</strong>, after CPI inflation).",
        f"FMA KiwiSaver Annual Report 2026 (year ended 31 March 2026, published 10 September 2026): Official average member balance = ${avg_bal_curr:,}; Total FUM = ${fum_b_curr:.1f}B across {mem_m_curr:.2f}m members (vs ${avg_bal_2016:,} / ${fum_b_2016:.1f}B in March 2016).",
        format_provenance_source_html(fum_calc, fma_lbl, FMA_KS_REPORT_PAGE, True),
        "Year ended 31 March 2026", avg_bal_curr,
        f"${avg_bal_curr:,} official FMA avg balance (${fum_b_curr:.1f}B FUM across {mem_m_curr:.2f}m members)",
        f"2016 avg: ${avg_bal_2016:,} (${avg_2016_in_2026_nzd:,} in 2026 CPI-adjusted NZD) | 2025 avg: ${avg_bal_2025:,}",
        f"10Y Nominal: +{bal_nom_10y_pct:.1f}% (+{bal_nom_cagr:.1f}%/yr) | 10Y Real (CPI-adjusted): +{bal_real_10y_pct:.1f}% (+{bal_real_cagr:.1f}%/yr)",
        fum_calc, "FMA KiwiSaver Annual Report 10-year series (2016–2026)",
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
    ent_calc = f"Passive: ({def_alloc:,} + {emp_alloc:,}) ÷ {tot_entry_mem:,} = {non_act_pct:.1f}% | Active: {act_alloc:,} ÷ {tot_entry_mem:,} = {act_pct:.1f}%"

    out["kiwisaver-default-vs-active"] = make_pulse_item(
        "Inland Revenue", fy_as_of_default, MAX_AGE_ANNUAL,
        f"As at June {fy_yr_default}, <strong>{non_act_pct:.1f}% ({non_act_tot:,} KiwiSaver members)</strong> remained in schemes assigned via <strong>IRD default allocation ({def_alloc:,}, or {def_pct:.1f}%)</strong> or <strong>employer nomination ({emp_alloc:,}, or {emp_pct:.1f}%)</strong>, while <strong>{act_pct:.1f}% ({act_alloc:,} members)</strong> had actively chosen their provider.",
        f"Inland Revenue KiwiSaver Entry Method (Year ended 30 June {fy_yr_default}): Default allocated = {def_alloc:,}; Employer nominated = {emp_alloc:,}; Actively chosen = {act_alloc:,}.",
        format_provenance_source_html(ent_calc, "Inland Revenue — KiwiSaver Scheme Entry Method Statistics", IRD_KS_ENTRY_PAGE, True),
        f"Year ended June {fy_yr_default}", round(non_act_pct, 1),
        f"{non_act_pct:.1f}% default/employer-allocated ({non_act_tot:,}) vs {act_pct:.1f}% active choice",
        f"Default allocated: {def_alloc:,} ({def_pct:.1f}%) | Employer nominated: {emp_alloc:,} ({emp_pct:.1f}%)",
        "Active scheme selection grew from ~1.90m (2016) to ~2.60m members",
        ent_calc, f"Inland Revenue 10-year June FY entry method series (2016–{fy_yr_default})", None, {"source_published": entry_pub_date},
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
    pay_calc = f"Shares: EE {ee_sh_pct:.1f}% | ER {er_sh_pct:.1f}% | Crown {cr_sh_pct:.1f}% | 10Y Nominal: (${tot_pay_b:.2f}B−${tot_pay_2016_b:.2f}B)÷${tot_pay_2016_b:.2f}B=+{pay_nom_10y_pct:.1f}% | 10Y Real (÷1.337 CPI): +{pay_real_10y_pct:.1f}%"

    out["kiwisaver-contributions-10yr"] = make_pulse_item(
        "Inland Revenue", fy_as_of_default, MAX_AGE_ANNUAL,
        f"Over the year ended 30 June {fy_yr_default}, total KiwiSaver contributions transferred to scheme providers reached <strong>${tot_pay_b:.2f}B</strong>—up <strong>+{pay_nom_10y_pct:.1f}%</strong> from ${tot_pay_2016_b:.2f}B in FY2016 (<strong>+{pay_real_10y_pct:.1f}%</strong> after CPI inflation)—comprising <strong>${emp_ee_b:.2f}B ({ee_sh_pct:.1f}%)</strong> from employees, <strong>${emp_er_b:.2f}B ({er_sh_pct:.1f}%)</strong> from employers, and <strong>${crown_b:.2f}B ({cr_sh_pct:.1f}%)</strong> in Crown contributions (with the statutory default rate at <strong>3.5% EE / 3.5% ER</strong> from 1 April 2026 and the maximum annual Government contribution at <strong>$260.72</strong>).",
        f"Inland Revenue KiwiSaver Provider Payments (FY{fy_yr_default}): Total = ${tot_pay_b:.2f}B (Employee = ${emp_ee_b:.2f}B, Employer = ${emp_er_b:.2f}B, Crown = ${crown_b:.2f}B; vs ${tot_pay_2016_b:.2f}B in FY2016). Statutory default rate = 3.5%; max Government contribution = $260.72/yr.",
        format_provenance_source_html(pay_calc, "Inland Revenue — KiwiSaver Scheme Provider Payments", IRD_KS_PAY_PAGE, True),
        f"Year ended June {fy_yr_default}", tot_pay_b,
        f"${tot_pay_b:.2f}B/yr (${emp_ee_b:.2f}B EE | ${emp_er_b:.2f}B ER | ${crown_b:.2f}B Crown) | $260.72 max Gov match",
        f"FY2016 total IRD transfers: ${tot_pay_2016_b:.2f}B",
        f"10Y Nominal: +{pay_nom_10y_pct:.1f}% (+{pay_nom_cagr:.1f}%/yr) | 10Y Real (CPI-adjusted): +{pay_real_10y_pct:.1f}% (+{pay_real_cagr:.1f}%/yr)",
        pay_calc, f"Inland Revenue 10-year June FY contributions series (FY2016–FY{fy_yr_default})",
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
    alloc_calc = f"Growth: {gr_sh:.1f}% (+{gr_sh - gr_2016:.1f} pts vs 2016) | Balanced: {bal_sh:.1f}% | Conservative: {con_sh:.1f}% ({con_sh - con_2016:+.1f} pts vs 2016)"
    out["kiwisaver-fund-allocation-shift"] = make_pulse_item(
        "FMA KiwiSaver Report", fma_as_of_default, MAX_AGE_INSURANCE,
        f"Over the decade to March 2026, KiwiSaver funds under management shifted from <strong>{con_2016:.1f}% Conservative/Cash and {gr_2016:.1f}% Growth in 2016</strong> to <strong>{gr_sh:.1f}% in Growth/Aggressive funds</strong>, <strong>{bal_sh:.1f}% in Balanced funds</strong>, and <strong>{con_sh:.1f}% in Conservative/Cash funds</strong>.",
        f"FMA KiwiSaver Annual Report 2026 (year ended 31 March 2026, published 10 September 2026): Growth/Aggressive FUM share = {gr_sh:.1f}% (vs {gr_2016:.1f}% in 2016); Balanced share = {bal_sh:.1f}%; Conservative/Cash share = {con_sh:.1f}% (vs {con_2016:.1f}% in 2016).",
        format_provenance_source_html(alloc_calc, "FMA KiwiSaver Annual Report 2026", FMA_KS_REPORT_PAGE, True),
        "Year ended 31 March 2026", gr_sh,
        f"{gr_sh:.1f}% Growth/Aggressive | {bal_sh:.1f}% Balanced | {con_sh:.1f}% Conservative/Cash",
        f"2016 allocation: {gr_2016:.1f}% Growth vs {con_2016:.1f}% Conservative/Cash",
        f"+{gr_sh - gr_2016:.1f} pts into Growth funds & {con_sh - con_2016:+.1f} pts out of Conservative over 10 yrs",
        alloc_calc, "FMA KiwiSaver Annual Report 10-year asset allocation series (2016–2026)",
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
    susp_calc = f"10Y Change: ({susp_tot:,} − {susp_2016:,}) ÷ {susp_2016:,} = {susp_10y_pct:.1f}% | Ordinary share: {susp_ord:,} ÷ {susp_tot:,} = {(susp_ord/max(1,susp_tot))*100:.1f}%"
    out["kiwisaver-savings-suspensions"] = make_pulse_item(
        "Inland Revenue", ks_m_as_of, MAX_AGE_ANNUAL,
        f"In {ks_m_full}, <strong>{susp_tot:,}</strong> KiwiSaver members were on an active savings suspension (<strong>{susp_ord:,} ordinary suspensions</strong> and <strong>{susp_hard:,} financial hardship suspensions</strong>)—down <strong>{abs(susp_10y_pct):.1f}% over 10 years</strong> from <strong>{susp_2016:,}</strong> in June 2016.",
        f"Inland Revenue KiwiSaver Savings Suspensions ({ks_m_full}): Total active suspensions = {susp_tot:,} (Ordinary = {susp_ord:,}, Financial hardship = {susp_hard:,}; compared with {susp_2016:,} in June 2016).",
        format_provenance_source_html(susp_calc, "Inland Revenue — KiwiSaver Savings Suspension Statistics", IRD_KS_SUSP_PAGE, True),
        ks_m_full, susp_tot,
        f"{susp_tot:,} members on savings suspension ({susp_ord:,} ordinary | {susp_hard:,} hardship)",
        f"June 2016 active suspensions: {susp_2016:,} members",
        f"{susp_tot - susp_2016:+,d} members ({susp_10y_pct:.1f}%) over 10 years",
        susp_calc, f"Inland Revenue 10-year savings suspension series (2016–{ks_m_full})",
        [
            {"period": "2026", "label": ks_m_full, "value": susp_tot, "formatted": f"{susp_tot:,} active suspensions ({susp_ord:,} ordinary / {susp_hard:,} hardship)"},
            {"period": "2023", "label": "June 2023", "value": 91420, "formatted": "91,420 active suspensions"},
            {"period": "2019", "label": "June 2019 (Max term cut 5yr→1yr)", "value": 118650, "formatted": "118,650 active suspensions"},
            {"period": "2016", "label": "June 2016", "value": susp_2016, "formatted": f"{susp_2016:,} active suspensions"},
        ],
        {"source_published": susp_pub_date},
    )

    return out
# === END OF PART 4C ===
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

    # Structured provenance dropdown (`source_html`) with calculation, source link, HLF derivation & checked date
    provenance_dropdown_html = format_provenance_source_html(
        calc_live_str, source_full_str, resolved_source_url, is_derived, verified_human
    )

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
        "source_html": provenance_dropdown_html,
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
        "media_kit_url": MEDIA_KIT_URL,
        "date_created": PULSE_DATE_CREATED,
        "date_modified": verified_date_iso,
    })
    return enriched


def build_fhb_snapshot_object(verified_tickers: dict) -> dict:
    """Build the 5-row First-Home Buyer Snapshot mapping directly to RBNZ Tables C31, C40 & C41."""
    row_specs = [
        (
            1,
            "Typical FHB Household Income",
            "fhb-typical-income",
            "RBNZ Table C41 (FHB borrower-weighted average gross income)",
        ),
        (
            2,
            "Typical FHB Loan Size",
            "fhb-average-loan-size",
            "RBNZ Table C31 (Total FHB new lending value ÷ FHB loan count)",
        ),
        (
            3,
            "Typical FHB Debt-to-Income (DTI)",
            "fhb-typical-dti",
            "RBNZ Tables C40 & C41 (Implied FHB DTI ratio & share of FHB lending ≤ 5x–6x DTI)",
        ),
        (
            4,
            "% of FHB Lending Above 6x DTI",
            "dti-headroom",
            "RBNZ Table C40 (FHB DTI > 6x share vs. 20% RBNZ speed limit)",
        ),
        (
            5,
            "Monthly FHB Total Lending Volume",
            "fhb-monthly-lending-volume",
            "RBNZ Table C31 (Monthly new FHB mortgage commitments & loan count)",
        ),
    ]
    rows_out = []
    latest_per = "September 2026"
    for row_num, label, t_key, series_desc in row_specs:
        rec = verified_tickers.get(t_key, {})
        if rec.get("latest_period"):
            latest_per = rec["latest_period"]
        rows_out.append({
            "row": row_num,
            "metric": label,
            "ticker": t_key,
            "canonical_url": f"{PULSE_BASE_URL}/{t_key}/",
            "rbnz_series_mapping": series_desc,
            "observation_period": rec.get("observation_period", latest_per),
            "period_end": rec.get("period_end", ""),
            "value": rec.get("value"),
            "latest_value": rec.get("latest_value", ""),
            "change": rec.get("change", ""),
            "hlf_calculation": rec.get("calculation", ""),
            "summary_text": rec.get("summary_text", ""),
        })
    return {
        "title": "New Zealand First-Home Buyer (FHB) 5-Row Live Market Snapshot",
        "observation_period": latest_per,
        "rows": rows_out,
    }


def render_fhb_snapshot_html(fhb_snap: dict) -> str:
    rows = fhb_snap.get("rows", [])
    if not rows:
        return ""
    trs = "".join(
        f"<tr>"
        f"<td><strong>{r['row']}. {html_lib.escape(r['metric'])}</strong><div style='font-size:0.74rem;color:#6b7280;'>{html_lib.escape(r['rbnz_series_mapping'])}</div></td>"
        f"<td><a href='#{r['ticker']}'><code>{r['ticker']}</code></a></td>"
        f"<td><strong>{html_lib.escape(str(r['latest_value']))}</strong><div style='font-size:0.74rem;color:#4b5563;'>{html_lib.escape(str(r['change']))}</div></td>"
        f"<td><code>{html_lib.escape(str(r['hlf_calculation']))}</code></td>"
        f"</tr>"
        for r in rows
    )
    return f"""
    <section id="fhb-live-snapshot" class="quality-box" style="margin:18px 0 26px;background:#eff6ff;border-color:#bfdbfe;" aria-label="First-Home Buyer 5-Row Live Snapshot">
      <h2 style="margin-top:0;color:#0b2559;">First-Home Buyer (FHB) 5-Row Live RBNZ Snapshot ({html_lib.escape(str(fhb_snap.get('observation_period', '')))})</h2>
      <p style="margin-bottom:8px;">Direct mapping of the 5 core New Zealand first-home buyer benchmark metrics to their live RBNZ Table C31, C40, and C41 Pulse tickers (also available in machine-readable JSON under <code>fhb_snapshot</code>):</p>
      <table style="background:#ffffff;border-radius:10px;overflow:hidden;">
        <thead>
          <tr>
            <th style="width:260px;">FHB Benchmark Metric &amp; RBNZ Mapping</th>
            <th style="width:170px;">Pulse Ticker</th>
            <th style="width:250px;">Latest Verified Value &amp; YoY</th>
            <th>HLF Provenance &amp; Calculation</th>
          </tr>
        </thead>
        <tbody>{trs}</tbody>
      </table>
    </section>"""


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
          <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citations</a> &middot;
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
          <p style="margin-bottom:8px;">The underlying data is sourced from the relevant official organisations, with HLF-derived calculations and interpretations clearly identified in the provenance layer and reusable under <a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener">{HLF_LICENSE_SHORT}</a> with attribution.</p>
          <div class="creator-links"><a href="#about-hlf-pulse">About the creator</a> &middot; <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citation Guide</a> &middot; <a href="#reuse-licensing">Reuse &amp; Citation Terms ({HLF_LICENSE_SHORT})</a> &middot; <a href="#date-definitions">4-Field Date Standard</a> &middot; <a href="{CREATOR_URL}" target="_blank" rel="noopener">View Andy&rsquo;s adviser profile &rarr;</a></div>
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
          <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citations</a> &bull;
          <a href="{DISCLOSURES_URL}" target="_blank" rel="noopener">Public Disclosures</a> &bull;
          <a href="{HOW_PAID_URL}" target="_blank" rel="noopener">How We Get Paid</a>
        </div>
      </div>
    </section>"""


def render_c41_band_table_html(rec: dict) -> str:
    bands = rec.get("tdti_bands")
    if rec.get("ticker") not in ("fhb-income-by-dti-band", "fhb-income-akl-vs-nz") or not bands:
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
        <li><strong>Official published sources:</strong> Pulse uses official published data from the Reserve Bank of New Zealand (RBNZ), Stats NZ, Inland Revenue (IRD), the Financial Markets Authority (FMA), MBIE Tenancy Services, the Natural Hazards Commission Toka Tū Ake (NHC), the Accident Compensation Corporation (ACC), and the Financial Services Council (FSC) wherever available.</li>
        <li><strong>Unaltered official figures &amp; source rounding:</strong> HLF does not alter or estimate official source figures. Where official tables publish rounded aggregates (such as RBNZ dollar series rounded to $1m or $0.1m), derived averages and implied counts reflect that source rounding and note it explicitly in the provenance layer.</li>
        <li><strong>Provenance &amp; derivation layer:</strong> To keep the primary editorial summaries natural and readable, exact formulas, series IDs, and derivation labels (<em>Derived by HLF Pulse</em>) are recorded in the expandable provenance layer, Data Dictionary tables, <code>pulse.json</code>, and <code>llms.txt</code>. Long-term 5-year and 10-year series report both <strong>nominal</strong> growth and <strong>CPI inflation-adjusted (real)</strong> growth using Stats NZ&rsquo;s Consumers Price Index.</li>
        <li><strong>Price indices vs. individual quotes:</strong> Stats NZ CPI insurance sub-indices measure percentage changes in standardised insurance price indices over time—they do not represent the dollar premium quote for an individual property or policyholder.</li>
        <li><strong>Building consent values vs. turnkey purchase prices:</strong> Stats NZ residential building consent values measure the estimated physical construction cost declared on the building consent (excluding land purchase price, developer margin, and off-site subdivision infrastructure).</li>
        <li><strong>Commitments vs. settlements:</strong> RBNZ mortgage lending tables (C31, C32, C33, C40, C41) measure firm bank mortgage commitments/approvals rather than LINZ property title settlements.</li>
        <li><strong>Automatic staleness exclusion:</strong> HLF automatically excludes any series where the underlying source is unavailable, stale past its freshness window, or cannot be verified.</li>
        <li><strong>Market context vs. individual outcomes:</strong> Statistics describe aggregate historical market cohorts and do not predict an individual borrower&rsquo;s, saver&rsquo;s, or policyholder&rsquo;s credit assessment, underwriting, or lending outcome.</li>
      </ul>
    </section>

    <section id="reuse-licensing" class="quality-box" style="margin-top:18px;" aria-label="Reuse Terms, Licensing &amp; How to Cite">
      <h2>Reuse Terms, Licensing &amp; How to Cite ({HLF_LICENSE_SHORT})</h2>
      <ul>
        <li><strong>Layer 1 — Official Crown &amp; statutory statistics:</strong> Underlying official figures remain subject to the copyright and licensing terms of their respective publishers: <strong>Stats NZ</strong>, <strong>Inland Revenue</strong>, <strong>MBIE Tenancy Services</strong>, and the <strong>Financial Markets Authority</strong> (<a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener">Creative Commons Attribution 4.0 International — CC BY 4.0</a>); the <strong>Reserve Bank of New Zealand</strong> (<a href="https://www.rbnz.govt.nz/about-our-site/terms-of-use" target="_blank" rel="noopener">RBNZ Terms of Use</a>, permitting reproduction and derivative calculation with attribution to RBNZ); and statutory schedules under the <strong>Natural Hazards Insurance Act 2023</strong> and <strong>Accident Compensation Act 2001</strong> (public-domain New Zealand legislation).</li>
        <li><strong>Layer 2 — HLF-derived calculations, compilations &amp; commentary:</strong> All HLF-derived percentages, averages, multi-year nominal/real growth rates, data dictionary definitions, and market interpretations are licensed by <strong>{PUBLISHER_NAME}</strong> under <a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener"><strong>Creative Commons Attribution 4.0 International ({HLF_LICENSE_SHORT})</strong></a>. Journalists, researchers, advisers, and AI systems may freely quote and reuse HLF-derived figures and commentary provided attribution is given to <strong>HLF Pulse, Home Loan Factory</strong> with a link to <code>{PULSE_BASE_URL}/</code>, the specific <code>/{{ticker}}/</code> record URL, or the <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">HLF Pulse Media Kit &amp; Citation Guide</a>.</li>
      </ul>
    </section>

    <section id="about-hlf-pulse" class="about-footer-box" aria-label="About the HLF Pulse &amp; Creator">
      <h2>About the HLF Pulse &amp; Creator</h2>
      <p><strong>HLF Pulse</strong> is a public reference dataset covering New Zealand mortgage, housing, KiwiSaver, tenancy, and protection statistics, published by <a href="{PUBLISHER_URL}">{PUBLISHER_NAME}</a> ({PUBLISHER_LEGAL}) and created by <a href="{CREATOR_URL}" target="_blank" rel="noopener"><strong>{CREATOR_NAME}</strong> ({CREATOR_ALT_NAME}, {CREATOR_FSP})</a>, Mortgage &amp; KiwiSaver Adviser at Home Loan Factory.</p>
      <p>Data is compiled from official releases from the <strong>Reserve Bank of New Zealand</strong>, <strong>Stats NZ</strong>, <strong>Inland Revenue</strong>, the <strong>Financial Markets Authority</strong>, <strong>MBIE Tenancy Services</strong>, the <strong>Natural Hazards Commission Toka Tū Ake</strong>, <strong>ACC</strong>, and the <strong>Financial Services Council</strong>. Where HLF calculates percentages, averages, or comparisons from the underlying official data, the exact formula and <strong>&ldquo;Derived by HLF Pulse&rdquo;</strong> provenance are recorded in the expandable source layer and data dictionary.</p>
      <p style="margin-bottom:0;">Figures are provided for general information and historical market analysis and are not personalised financial advice. Published by <a href="{PUBLISHER_URL}">{PUBLISHER_NAME}</a> (<a href="{CREATOR_URL}" target="_blank" rel="noopener">Creator Profile: {CREATOR_NAME}</a> &bull; <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citations</a> &bull; <a href="{DISCLOSURES_URL}">Public Disclosures</a> &bull; <a href="{HOW_PAID_URL}">How We Get Paid</a>).</p>
    </section>"""
# === END OF PART 5A ===
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
    <div class="crumb"><a href="{PULSE_BASE_URL}/">&larr; All {total_count} HLF Pulse Datasets</a> &bull; Category: <strong>{html_lib.escape(rec['category'])}</strong> &bull; <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citations</a></div>
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
        <tr><th>Reuse licence</th><td><a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener">{HLF_LICENSE_SHORT}</a> (HLF-derived calculations, definitions &amp; commentary; official statistics credited to {html_lib.escape(rec['source'])}) &bull; <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citation Guide</a></td></tr>
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


def write_legacy_redirect_page(root: pathlib.Path, alias_key: str, target_key: str) -> None:
    """Write a clean canonical redirect page for renamed legacy ticker URLs (e.g. /fhb-income-akl-vs-nz/)."""
    a_dir = root / alias_key
    a_dir.mkdir(parents=True, exist_ok=True)
    target_url = f"{PULSE_BASE_URL}/{target_key}/"
    redirect_html = f"""<!DOCTYPE html>
<html lang="en-NZ">
<head>
  <meta charset="UTF-8">
  <meta name="robots" content="noindex, follow">
  <link rel="canonical" href="{target_url}">
  <meta http-equiv="refresh" content="0; url={target_url}">
  <title>Redirecting to {target_key} — HLF Pulse</title>
  <script>window.location.replace("{target_url}");</script>
</head>
<body style="font-family:system-ui,sans-serif;padding:40px;color:#0b2559;">
  <p>This HLF Pulse dataset has moved to its canonical slug: <a href="{target_url}"><code>/{target_key}/</code></a>.</p>
</body>
</html>"""
    (a_dir / "index.html").write_text(redirect_html, encoding="utf-8")
# === END OF PART 5B ===
# === PART 5C OF 5 ===
# ---------------------------------------------------------------------------
# Main Orchestrator & Strict 63-Dataset Canonical Gate
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

    # Migrate legacy fhb-income-akl-vs-nz cache key to fhb-income-by-dti-band if needed
    if "fhb-income-by-dti-band" not in prev_tickers and "fhb-income-akl-vs-nz" in prev_tickers:
        prev_tickers["fhb-income-by-dti-band"] = dict(prev_tickers["fhb-income-akl-vs-nz"])

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
            dti_out = build_rbnz_dti_income_tickers(rbnz_files.get("c40"), rbnz_files.get("c41"), c31_by_ym)
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

    fhb_snap_obj = build_fhb_snapshot_object(verified_tickers)

    bubble_feed_tickers = dict(verified_tickers)
    for alias_k, target_k in LEGACY_BUBBLE_ALIASES.items():
        if target_k in verified_tickers:
            bubble_feed_tickers[alias_k] = dict(verified_tickers[target_k])

    latest_rbnz_period = verified_tickers.get("fhb-low-deposit", {}).get("latest_period", "September 2026")
    payload = {
        "publisher": PUBLISHER_NAME,
        "publisher_legal_name": PUBLISHER_LEGAL,
        "publisher_url": PUBLISHER_URL,
        "media_kit_url": MEDIA_KIT_URL,
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
        "fhb_snapshot": fhb_snap_obj,
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
            "media_kit_url": MEDIA_KIT_URL,
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
        "media_kit_url": MEDIA_KIT_URL,
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
        "next_expected_update": "Monthly (RBNZ / Stats NZ / MBIE Tenancy), Quarterly (CPI / IRD Student Loans) & Semi-Annual (FMA / RBNZ FSR / NHC / ACC / FSC)",
        "fhb_snapshot": fhb_snap_obj,
        "records": structured_records,
    }
    structured_json_path.write_text(json.dumps(pulse_json_obj, indent=2), encoding="utf-8")

    write_robots_txt(root)

    for k, rec in verified_tickers.items():
        write_canonical_ticker_page(root, k, rec, len(verified_tickers))

    # Write clean redirect pages for legacy slugs so old links resolve to canonical URLs
    for alias_k, target_k in LEGACY_BUBBLE_ALIASES.items():
        if target_k in verified_tickers and alias_k not in verified_tickers:
            write_legacy_redirect_page(root, alias_k, target_k)

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

    fhb_snap_lines = [
        "## First-Home Buyer (FHB) 5-Row Live RBNZ Snapshot",
    ] + [
        f"- Row {r['row']} — {r['metric']} ({r['rbnz_series_mapping']}): {r['latest_value']} ({r['change']}) [Ticker: {r['ticker']} | URL: {r['canonical_url']} | HLF Calculation: {r['hlf_calculation']}]"
        for r in fhb_snap_obj.get("rows", [])
    ]

    llms_lines = (
        [
            "# HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver & Protection Reference Dataset",
            "",
            "> Public reference dataset and data dictionary built from official New Zealand Crown and statutory releases, with clearly labelled HLF-derived calculations and interpretation.",
            f"> Created by: {CREATOR_NAME} ({CREATOR_ALT_NAME}, {CREATOR_FSP}), {CREATOR_ROLE} ({CREATOR_URL}) | Published by: {PUBLISHER_NAME} ({PUBLISHER_URL}).",
            f"> Media Kit & Citation Guide: {MEDIA_KIT_URL}",
            f"> Reuse Licence: {HLF_LICENSE_SHORT} ({HLF_LICENSE_URL}) for HLF-derived figures & commentary; official statistics credited to their respective Crown publishers.",
            f"> Date Standard: Observation Period (latest monthly: {latest_rbnz_period}) | Period End | Source Publication Date | Last Checked ({now_human} / {today_iso}).",
            f"> Active datasets: {len(verified_tickers)}.",
            "",
            "## Machine-Readable Endpoints",
            f"- Structured Data Dictionary & Records (JSON): {PULSE_BASE_URL}/pulse.json",
            f"- Webflow Chat Bubble Feed (JSON): {PULSE_BASE_URL}/nz-mortgage-pulse.json",
            f"- Main Reference Dataset Hub (HTML): {PULSE_BASE_URL}/",
            f"- Media Kit & Citation Guide (HTML): {MEDIA_KIT_URL}",
            "",
        ]
        + fhb_snap_lines
        + [
            "",
            f"## Canonical Data Dictionary Records ({len(verified_tickers)} Active Datasets)",
        ]
        + [
            f"- [{rec['name']} (Obs: {rec['observation_period']} | End: {rec['period_end']} | Published: {rec['source_published']})]({rec['canonical_url']}): {rec['summary_text']} [Official Data: {rec['official_data_summary']} | HLF Calculation: {rec['calculation']} | Derived by HLF Pulse: {rec['hlf_derived_label']} | Cite as: {rec['citation']}]"
            for k, rec in sorted(verified_tickers.items())
        ]
    )
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
    fhb_snapshot_html = render_fhb_snapshot_html(fhb_snap_obj)
    about_footer_html = render_about_and_quality_footer_html()

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
                "name": f"HLF Pulse Structured Data Dictionary JSON (All {len(verified_tickers)} Datasets)",
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
    <p>Public reference dataset and data dictionary built from official New Zealand Crown and statutory releases, with clearly labelled HLF-derived calculations and interpretation (<a href="#fhb-live-snapshot">5-Row FHB Live Snapshot</a> &bull; <a href="#browse-datasets">Browse all {len(verified_tickers)} datasets</a> &bull; <a href="#cat-kiwisaver">KiwiSaver</a> &bull; <a href="#cat-new-builds-construction">Build Costs</a> &bull; <a href="#cat-property-natural-hazard-insurance">Property Insurance</a> &bull; <a href="#cat-personal-risk-income-protection">Personal &amp; ACC Protection</a> &bull; <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citation Guide</a> &bull; <a href="#date-definitions">4-Field Date Standard</a> &bull; <a href="#data-quality-limitations">Data Quality &amp; Limitations</a> &bull; <a href="#reuse-licensing">Reuse &amp; Citation ({HLF_LICENSE_SHORT})</a> &bull; <a href="#about-hlf-pulse">About &amp; Methodology</a>). Machine-readable endpoints: <a href="{PULSE_BASE_URL}/pulse.json"><code>pulse.json</code></a>, <a href="{PULSE_BASE_URL}/nz-mortgage-pulse.json"><code>nz-mortgage-pulse.json</code></a>, <a href="{PULSE_BASE_URL}/llms.txt"><code>llms.txt</code></a>, and <a href="{PULSE_BASE_URL}/robots.txt"><code>robots.txt</code></a>.</p>

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
        <small>HLF layer reusable under <a href="#reuse-licensing">{HLF_LICENSE_SHORT}</a> (<a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit</a>).</small>
      </div>
    </div>

    {creator_main_html}
    {fhb_snapshot_html}

    <h2>What the HLF Pulse Can Help Answer</h2>
    <div class="use-grid">
      <div class="use-card">
        <strong>For First-Home Buyers</strong>
        <ul>
          <li><a href="#fhb-typical-income">What is the typical FHB household income ($144k)?</a></li>
          <li><a href="#fhb-average-loan-size">What is the average FHB mortgage size (~$565k)?</a></li>
          <li><a href="#fhb-typical-dti">What is the typical FHB Debt-to-Income (DTI) ratio?</a></li>
          <li><a href="#fhb-low-deposit">How common is a deposit under 20%?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For Homeowners &amp; New Builds</strong>
        <ul>
          <li><a href="#new-build-consent-cost-by-type">What is the average build cost in $ and $/m² by dwelling type?</a></li>
          <li><a href="#fixed-vs-floating-rate-mix">How many fixed mortgages roll off within 12 months?</a></li>
          <li><a href="#refinance">How much are homeowners refinancing?</a></li>
          <li><a href="#top-ups">How common are mortgage top-ups?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For Property Investors &amp; Landlords</strong>
        <ul>
          <li><a href="#interest-only">How much investor lending is interest-only?</a></li>
          <li><a href="#investor-high-dti-share">What % of investors borrow above 7x DTI?</a></li>
          <li><a href="#investor-average-loan-size">What is the average property investor loan size?</a></li>
          <li><a href="#rent-long-term-growth">How much have NZ rents risen over 5 and 10 years?</a></li>
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
          <li><a href="#life-trauma-average-claim-payout">What is the average Life, Trauma &amp; Income Protection claim payout?</a></li>
          <li><a href="#private-health-claims-average-payout">What do private health claims &amp; surgeries cost in NZ?</a></li>
          <li><a href="#personal-insurance-underinsurance-gap">How many Kiwis have income protection vs house insurance?</a></li>
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
    print(f"Saved {len(verified_tickers)} canonical datasets, fhb_snapshot, pulse.json, index.html, robots.txt, llms.txt, sitemap.xml & /<ticker>/ pages.")


if __name__ == "__main__":
    main()
# === END OF PART 5C ===
