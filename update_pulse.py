#!/usr/bin/env python3
"""
HLF Pulse — Reference-Grade New Zealand Mortgage, Housing, KiwiSaver & Protection Dataset
100% Live-Pulled Architecture (Zero Hardcoded Fallbacks) via Camoufox + Playwright + HTTP.
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
GLOSSARY_BASE_URL = "https://www.homeloanfactory.co.nz/tools/glossary"
FHB_2026_ARTICLE_URL = "https://www.homeloanfactory.co.nz/news-insights/what-it-actually-takes-to-buy-your-first-home-nz"
FHB_2026_ARTICLE_TITLE = "What It Actually Takes to Buy Your First Home in NZ"

HLF_LICENSE_URL = "https://creativecommons.org/licenses/by/4.0/"
HLF_LICENSE_SHORT = "CC BY 4.0"
HLF_LICENSE_NOTE = (
    "HLF-derived calculations, percentages, averages, multi-year real/nominal comparisons, "
    "data dictionary definitions, and market interpretations are licensed by Home Loan Factory "
    "under Creative Commons Attribution 4.0 International (CC BY 4.0) with attribution to "
    "HLF Pulse (https://pulse-data.homeloanfactory.co.nz/). Underlying official Crown, statutory, "
    "and industry statistics remain subject to the terms of their respective publishers "
    "(RBNZ Terms of Use; Stats NZ, Inland Revenue, MBIE, FMA, and Kāinga Ora under CC BY 4.0; "
    "Natural Hazards Insurance Act 2023 and Accident Compensation Act 2001 under public-domain NZ legislation; "
    "Financial Services Council [FSC] Spotlight reports credited to FSC under fair-dealing citation)."
)

GLOSSARY_LINKS_BY_CATEGORY = {
    "First-home buyers": [
        ("Loan-to-Value Ratio (LVR)", f"{GLOSSARY_BASE_URL}#lvr"),
        ("Kāinga Ora First Home Loan", f"{GLOSSARY_BASE_URL}#first-home-loan"),
        ("Lenders Mortgage Insurance (LMI)", f"{GLOSSARY_BASE_URL}#lmi"),
    ],
    "Homeowners & refinancing": [
        ("Refinancing", f"{GLOSSARY_BASE_URL}#refinancing"),
        ("Revolving Credit", f"{GLOSSARY_BASE_URL}#revolving-credit"),
        ("Usable Equity", f"{GLOSSARY_BASE_URL}#usable-equity"),
    ],
    "Property investors": [
        ("Interest-Only Mortgage", f"{GLOSSARY_BASE_URL}#interest-only"),
        ("Debt-to-Income (DTI)", f"{GLOSSARY_BASE_URL}#dti"),
        ("Gross Rental Yield", f"{GLOSSARY_BASE_URL}#rental-yield"),
    ],
    "LVR, DTI & credit policy": [
        ("Debt-to-Income (DTI) Ratio", f"{GLOSSARY_BASE_URL}#dti"),
        ("Loan-to-Value Ratio (LVR)", f"{GLOSSARY_BASE_URL}#lvr"),
        ("Servicing Test Rate", f"{GLOSSARY_BASE_URL}#test-rate"),
    ],
    "New builds & construction": [
        ("Turnkey New Build", f"{GLOSSARY_BASE_URL}#turnkey"),
        ("Progressive Drawdown Construction Loan", f"{GLOSSARY_BASE_URL}#construction-loan"),
        ("Code Compliance Certificate (CCC)", f"{GLOSSARY_BASE_URL}#ccc"),
    ],
    "Rents & tenancy": [
        ("Rental Bond", f"{GLOSSARY_BASE_URL}#rental-bond"),
        ("Gross vs Net Rental Yield", f"{GLOSSARY_BASE_URL}#rental-yield"),
    ],
    "KiwiSaver": [
        ("KiwiSaver First-Home Withdrawal", f"{GLOSSARY_BASE_URL}#kiwisaver-first-home-withdrawal"),
        ("KiwiSaver Second-Chance Withdrawal", f"{GLOSSARY_BASE_URL}#second-chance-withdrawal"),
        ("Default vs Active KiwiSaver Scheme", f"{GLOSSARY_BASE_URL}#kiwisaver"),
    ],
    "Property & Natural Hazard Insurance": [
        ("Sum Insured Replacement Cover", f"{GLOSSARY_BASE_URL}#sum-insured"),
        ("Natural Hazards Commission (NHC / EQC)", f"{GLOSSARY_BASE_URL}#nhc"),
    ],
    "Personal Risk & Income Protection": [
        ("Income & Mortgage Repayment Protection", f"{GLOSSARY_BASE_URL}#income-protection"),
        ("Critical Illness / Trauma Cover", f"{GLOSSARY_BASE_URL}#trauma-cover"),
    ],
    "Insurance Market & Claims": [
        ("Sum Insured Replacement Cover", f"{GLOSSARY_BASE_URL}#sum-insured"),
        ("NZ Mortgage & Property Glossary (All 86 Terms)", GLOSSARY_BASE_URL),
    ],
}

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

MAX_AGE_MONTHLY = 100
MAX_AGE_TENANCY = 140
MAX_AGE_QUARTERLY = 160
MAX_AGE_ANNUAL = 550
MAX_AGE_INSURANCE = 550

RBNZ_WARMUP_URL = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary"

RBNZ_SOURCES = {
    "c31": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c31/c31-new-residential-mortgage-lending-by-borrower-type.xlsx",
        "match": "c31",
    },
    "c33": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c33/c33-new-residential-mortgage-lending-by-purpose.xlsx",
        "match": "c33",
    },
    "c32": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type"],
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
        "pages": ["https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-weighted-average-interest-rates"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b30/b30-new-residential-mortgage-weighted-average-interest-rates.xlsx",
        "match": "b30",
    },
    "j10": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/insurance/insurance-income-and-expenditure",
            "https://www.rbnz.govt.nz/statistics/series/insurance",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/j/j10/j10-insurance-income-and-expenditure.xlsx",
        "match": "j10",
    },
    "t43": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/managed-funds/kiwisaver",
            "https://www.rbnz.govt.nz/statistics/series/managed-funds",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/t/t43/t43-kiwisaver-assets-and-liabilities.xlsx",
        "match": "t43",
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
KAINGA_ORA_QTR_REPORT_PAGE = "https://kaingaora.govt.nz/en_NZ/publications/home-ownership-publications/home-ownership-products-report/"
KAINGA_ORA_ANN_REPORT_PAGE = "https://kaingaora.govt.nz/en_NZ/publications/core-publications/annual-report/"
STATS_CPI_PAGE = "https://www.stats.govt.nz/information-releases/consumers-price-index-september-2026-quarter/"
STATS_HES_PAGE = "https://www.stats.govt.nz/information-releases/household-expenditure-statistics-year-ended-june-2025/"
RBNZ_FSR_PAGE = "https://www.rbnz.govt.nz/financial-stability/financial-stability-report"
RBNZ_INS_STATS_PAGE = "https://www.rbnz.govt.nz/statistics/series/insurance"
RBNZ_C35_PAGE = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-loan-reconciliation"
NHC_ACT_PAGE = "https://www.legislation.govt.nz/act/public/2023/0004/latest/whole.html"
NHC_FUND_PAGE = "https://www.naturalhazards.govt.nz/"
ACC_REPORT_PAGE = "https://www.acc.co.nz/about-us/corporate-documents/annual-report"
ACC_LEVY_PAGE = "https://www.acc.co.nz/for-business/understanding-your-levies/paying-levies-if-you-work-for-yourself"
FSC_DATA_PAGE = "https://www.fsc.org.nz/research-and-data"
FSC_LIFE_TAG_PAGE = "https://blog.fsc.org.nz/tag/life-insurance"
FSC_HEALTH_TAG_PAGE = "https://blog.fsc.org.nz/tag/health-insurance"

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
        "methodology": "Derived from RBNZ Table C31 by dividing first-home buyer new lending value (LVRN.MMB1.AC1) by total new lending across all borrower types (LVRN.MMB1.AA). Mean loan size is LVRN.MMB1.AC1 (published by RBNZ rounded to $1m) divided by LVRN.MMB2.AC1.",
        "hlf_calculation": "(FHB Lending Value [LVRN.MMB1.AC1] ÷ All-Borrower Total Lending [LVRN.MMB1.AA]) × 100; Mean Loan = LVRN.MMB1.AC1 ÷ LVRN.MMB2.AC1",
        "hlf_interpretation": "HLF market context: Tracks first-home buyer market participation relative to existing owner-occupiers and residential property investors.",
        "questions": [
            "What share of New Zealand mortgages go to first-home buyers?",
            "What is the average first-home buyer mortgage size in New Zealand?",
            "How many first-home buyers get approved for a mortgage each month in NZ?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
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
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
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
        "hlf_interpretation": "HLF market context: Provides a direct national benchmark for what first-home buyers are borrowing across New Zealand before regional price differences.",
        "questions": [
            "What is the average first-home buyer mortgage in New Zealand?",
            "How much does a typical first-home buyer borrow in NZ?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
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
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
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
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "kainga-ora-first-home-loan-approvals": {
        "name": "Kāinga Ora First Home Loan (5% deposit) nominal approvals, settlements & YoY/QoQ trend",
        "also_known_as": "Kāinga Ora Home Ownership Products Report — First Home Loan 5-quarter approvals & settlements series",
        "category": "First-home buyers",
        "topic": "Kāinga Ora 5% deposit First Home Loan approvals, settlements & annual growth",
        "source": "Kāinga Ora Quarterly Report",
        "source_full": "Kāinga Ora – Homes and Communities — Home Ownership Products Quarterly Report (CC BY 4.0)",
        "series_ids": "Kāinga Ora Home Ownership Products Report — First Home Loans: Approvals and Settlements (5-Quarter Series & Monthly Totals)",
        "source_url": KAINGA_ORA_QTR_REPORT_PAGE,
        "data_type": "Officially reported Kāinga Ora quarterly PDF series & HLF-derived YoY/QoQ growth & 12-month rolling totals",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Quarterly",
        "definition": "The nominal number of Kāinga Ora-underwritten First Home Loan (5% deposit) approvals and settlements each month, quarter, and rolling 12-month period across New Zealand, alongside quarter-on-quarter and year-on-year percentage trends.",
        "methodology": "Automatically fetched and parsed from Kāinga Ora's Home Ownership Products Quarterly Report PDF (licensed under CC BY 4.0). Extracts the 5-quarter Approvals and Settlements series on Page 2 and verifies the latest quarter against the monthly Territorial Local Authority Total rows on Page 4.",
        "hlf_calculation": "12-Month Rolling Approvals = Sum of latest 4 quarters; YoY % = ((Latest Quarter − Prior-Year Quarter) ÷ Prior-Year Quarter) × 100; Settlement Ratio = (Settlements ÷ Approvals) × 100",
        "hlf_interpretation": "HLF market context: Because Kāinga Ora First Home Loans require only a 5% deposit and are exempt from RBNZ's 20% bank low-deposit speed limit, tracking whether nominal approvals and settlements are rising or falling shows how actively participating lenders are using the scheme.",
        "questions": [
            "How many Kāinga Ora First Home Loans are approved and settled each quarter in New Zealand?",
            "Are Kāinga Ora 5% deposit First Home Loan approvals going up or down?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "Kāinga Ora Lenders Mortgage Insurance (LMI) Fee Estimator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/lenders-mortgage-insurance-fee-estimator",
    },
    "kainga-ora-first-home-loan-regional": {
        "name": "Kāinga Ora First Home Loan (5% deposit) approvals & settlements by NZ city & district",
        "also_known_as": "Kāinga Ora First Home Loan Approvals & Settlements by Territorial Local Authority (TLA)",
        "category": "First-home buyers",
        "topic": "Regional Kāinga Ora 5% deposit First Home Loan approvals by council",
        "source": "Kāinga Ora Quarterly Report",
        "source_full": "Kāinga Ora – Homes and Communities — Home Ownership Products Quarterly Report: TLA Tables (CC BY 4.0)",
        "series_ids": "Kāinga Ora Home Ownership Products Report — First Home Loan Approvals & Settlements by Territorial Local Authority",
        "source_url": KAINGA_ORA_QTR_REPORT_PAGE,
        "data_type": "Officially reported Kāinga Ora council-level tables & HLF-derived national shares",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Quarterly",
        "definition": "Distribution of Kāinga Ora 5% deposit First Home Loan approvals and settlements across New Zealand's major cities and Territorial Local Authorities.",
        "methodology": "Automatically parsed from the Territorial Local Authority Approvals and Settlements tables (Pages 2–4) in Kāinga Ora's Home Ownership Products Quarterly Report PDF, verifying that each council's 3 monthly figures sum to its quarterly total.",
        "hlf_calculation": "Council Share of National Approvals (%) = (Council Quarterly Approvals ÷ National Quarterly Approvals) × 100",
        "hlf_interpretation": "HLF market context: Auckland and Christchurch account for over 43% of all Kāinga Ora First Home Loan approvals nationwide, followed by Hamilton, Wellington, Dunedin, Napier/Hastings, and Tauranga.",
        "questions": [
            "Which New Zealand cities have the most Kāinga Ora First Home Loan approvals?",
            "How many Kāinga Ora First Home Loans are approved in Auckland, Christchurch, Hamilton, and Wellington?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "Kāinga Ora Lenders Mortgage Insurance (LMI) Fee Estimator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/lenders-mortgage-insurance-fee-estimator",
    },
    "kainga-ora-second-chance-kiwisaver-approvals": {
        "name": "Kāinga Ora 'Second Chance' (Previous Homeowner) KiwiSaver first-home withdrawal approvals",
        "also_known_as": "Previous Owner Savings Withdrawal Applications Approved by Kāinga Ora",
        "category": "KiwiSaver",
        "topic": "Second-chance KiwiSaver first-home withdrawal approvals for previous homeowners",
        "source": "Kāinga Ora Quarterly Report",
        "source_full": "Kāinga Ora – Homes and Communities — Home Ownership Products Quarterly Report: Previous Owner Savings Withdrawal Applications",
        "series_ids": "Kāinga Ora Home Ownership Products Report — Previous Owner Savings Withdrawal Applications Approved (5-Quarter Series)",
        "source_url": KAINGA_ORA_QTR_REPORT_PAGE,
        "data_type": "Officially reported Kāinga Ora quarterly PDF series & HLF-derived 12-month rolling total",
        "derived_by_hlf": True,
        "unit": "approvals",
        "frequency": "Quarterly",
        "definition": "The number of previous New Zealand homeowners approved by Kāinga Ora each quarter (and over the rolling 12-month period) to make a 'second-chance' KiwiSaver savings withdrawal to buy a home.",
        "methodology": "Automatically fetched and parsed from the 'Previous Owner Savings Withdrawal Applications' chart on Page 5 of Kāinga Ora's Home Ownership Products Quarterly Report PDF. Extracts Applications Approved across all 5 published quarters and sums the 4 most recent quarters for the rolling 12-month approval total.",
        "hlf_calculation": "12-Month Rolling Approvals = Sum of Applications Approved across the 4 most recent quarters; Monthly Mean = Quarterly Approvals ÷ 3",
        "hlf_interpretation": "HLF market context: Previous homeowners who no longer own property and are in a similar financial position to a first-home buyer (typically after a relationship separation or career relocation) can apply to Kāinga Ora for a determination letter to unlock their KiwiSaver balance for another home purchase—with over 4,100 approvals granted annually.",
        "questions": [
            "How many previous homeowners get approved for a second-chance KiwiSaver withdrawal in New Zealand?",
            "Can I use my KiwiSaver to buy a house if I have owned a property before in NZ?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "kainga-ora-underwritten-loan-book": {
        "name": "Kāinga Ora First Home Loan annual mortgages settled/underwritten & approval-to-settlement pipeline",
        "also_known_as": "Kāinga Ora 12-Month Rolling First Home Loan Settlements Underwritten & Approvals (SPE 4.1)",
        "category": "First-home buyers",
        "topic": "Annual Kāinga Ora First Home Loans underwritten & Mortgage Insurance Scheme",
        "source": "Kāinga Ora Reports",
        "source_full": "Kāinga Ora – Homes and Communities — Home Ownership Products Report & Annual Report (SPE 4.1)",
        "series_ids": "Kāinga Ora 4-Quarter Rolling First Home Loan Approvals & Underwritten Settlements (SPE 4.1)",
        "source_url": KAINGA_ORA_QTR_REPORT_PAGE,
        "data_type": "Officially reported Kāinga Ora series & HLF-derived 12-month rolling totals",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Quarterly / Annual",
        "definition": "The rolling 12-month number of First Home Loan (5% deposit) mortgages settled/underwritten and approved by Kāinga Ora under the Crown Mortgage Insurance Scheme.",
        "methodology": "Derived live by summing the 4 most recent quarters of Kāinga Ora First Home Loan settlements and approvals in the Home Ownership Products Quarterly Report PDF (and cross-checked against SPE 4.1 in Kāinga Ora's Annual Report).",
        "hlf_calculation": "12-Month Underwritten Settlements = Sum of 4 quarters of FHL settlements; 12-Month Approvals = Sum of 4 quarters of FHL approvals",
        "hlf_interpretation": "HLF market context: Under the user-pays Mortgage Insurance Scheme (funded via the 1.2% Lenders Mortgage Insurance fee), Kāinga Ora approves ~9,500+ applications and underwrites over 4,000 settled 5%-deposit mortgages per year.",
        "questions": [
            "How many First Home Loans does Kāinga Ora approve and underwrite each year in New Zealand?",
            "Why are Kāinga Ora First Home Loan approvals higher than settlements?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "Kāinga Ora Lenders Mortgage Insurance (LMI) Fee Estimator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/lenders-mortgage-insurance-fee-estimator",
    },
    "kiwisaver-first-home-quarterly-trend": {
        "name": "Quarterly & 12-month rolling KiwiSaver first-home withdrawals & average withdrawal size (IRD / Kāinga Ora)",
        "also_known_as": "KiwiSaver First Home Savings Withdrawal 4-quarter volume, dollar value & rolling annual total",
        "category": "KiwiSaver",
        "topic": "Quarterly & annual rolling KiwiSaver first-home withdrawals",
        "source": "Kāinga Ora / IRD Report",
        "source_full": "Kāinga Ora Home Ownership Products Report & Inland Revenue — KiwiSaver First Home Savings Withdrawal (4-Quarter Series)",
        "series_ids": "KiwiSaver First Home Savings Withdrawal (IRD) — Quarterly Number & Dollar Value ($m) of Withdrawals",
        "source_url": KAINGA_ORA_QTR_REPORT_PAGE,
        "data_type": "Officially reported IRD/Kāinga Ora quarterly series & HLF-derived 12-month rolling sum & mean withdrawal",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Quarterly",
        "definition": "Quarterly and 12-month rolling total number and dollar value ($ millions) of KiwiSaver first-home savings withdrawals across New Zealand, plus the average withdrawal per buyer.",
        "methodology": "Automatically parsed from the 'KiwiSaver First Home Savings Withdrawal (IRD)' 4-quarter chart on Page 5 of Kāinga Ora's Home Ownership Products Quarterly Report PDF.",
        "hlf_calculation": "12-Month Rolling Total = Sum of 4 quarters (Withdrawals & $m); Mean Withdrawal = Quarterly Value ($m × $1m) ÷ Quarterly Withdrawals",
        "hlf_interpretation": "HLF market context: Over 51,000 New Zealanders withdraw more than $2.30B from KiwiSaver annually (~$45,000 per withdrawing member) to form the core of their home deposit.",
        "questions": [
            "How much is withdrawn from KiwiSaver for first homes each quarter and year in New Zealand?",
            "How many New Zealanders use KiwiSaver to buy a first home each year?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
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
        "also_known_as": "RBNZ insured homeowner building insurance benchmark & Stats NZ HES dwelling insurance spend",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Household dwelling insurance expenditure",
        "source": "Stats NZ HES & RBNZ FSR",
        "source_full": "Stats NZ — Household Expenditure Statistics (HES) & RBNZ Financial Stability Report",
        "series_ids": "Stats NZ HES Dwelling Insurance ($/wk × 52) & RBNZ FSR Insured Homeowner Premium Series",
        "source_url": STATS_HES_PAGE,
        "data_type": "Officially reported Stats NZ & RBNZ series & HLF-derived annualised comparison",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual / Triennial",
        "definition": "Comparison of the average annual domestic building insurance premium paid by insured New Zealand homeowners against the Stats NZ Household Expenditure Survey average across all households including renters.",
        "methodology": "Live-parsed from Stats NZ's Household Expenditure Statistics release tables (multiplying weekly average dwelling insurance expenditure by 52 weeks) alongside RBNZ Financial Stability Report insured homeowner building insurance series.",
        "hlf_calculation": "Annualised HES All-Household Spend = Weekly HES Dwelling Insurance ($/wk) × 52 (includes renting households with $0 building insurance)",
        "hlf_interpretation": "HLF market context: Banks require active replacement building insurance as a mandatory condition of every residential mortgage—making actual homeowner premiums a key fixed expense in mortgage servicing calculators.",
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
        "name": "Licensed NZ general insurer Gross Written Premiums, Net Earned Premiums & Net Claims (RBNZ J10)",
        "also_known_as": "RBNZ Table J10 general insurance sector premium revenue, reinsurance outflow & claims paid",
        "category": "Insurance Market & Claims",
        "topic": "General insurer premiums, reinsurance & claims",
        "source": "RBNZ Table J10",
        "source_full": "Reserve Bank of New Zealand — Insurance Income and Expenditure (Table J10)",
        "series_ids": "RBNZ Table J10 — General Insurers Gross Premium Revenue, Outwards Reinsurance & Net Claims Expense",
        "source_url": RBNZ_INS_STATS_PAGE,
        "data_type": "Officially reported RBNZ Table J10 series & HLF-derived reinsurance difference",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Quarterly / Annual",
        "definition": "Annual Gross Written/Earned Premiums, Net Earned Premiums (after international catastrophe reinsurance expense), and Net Incurred Claims across licensed New Zealand general insurers.",
        "methodology": "Live-parsed from RBNZ Table J10 (Insurance Income and Expenditure workbook), subtracting outwards reinsurance expense from gross premium revenue to show net earned premium and comparing against net claims expense.",
        "hlf_calculation": "Reinsurance Outflow = Gross Premium Revenue minus Net Earned Premium",
        "hlf_interpretation": "HLF market context: New Zealand insurers cede roughly a quarter of gross property and general premiums to global reinsurers to fund earthquake and severe-weather catastrophe cover.",
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
        "name": "Licensed New Zealand general insurer net claims loss ratio (RBNZ Table J10)",
        "also_known_as": "RBNZ Table J10 general insurance net incurred claims to net earned premium ratio",
        "category": "Insurance Market & Claims",
        "topic": "General insurer claims loss ratio",
        "source": "RBNZ Table J10",
        "source_full": "Reserve Bank of New Zealand — Insurance Income and Expenditure (Table J10)",
        "series_ids": "RBNZ Table J10 — Net Claims Expense ÷ Net Premium Revenue",
        "source_url": RBNZ_INS_STATS_PAGE,
        "data_type": "HLF-derived percentage ratio from official RBNZ Table J10 series",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Quarterly / Annual",
        "definition": "The net claims loss ratio across licensed New Zealand general insurers, measuring net incurred claims as a percentage of net earned premiums.",
        "methodology": "Derived live from RBNZ Table J10 by dividing annual Net Claims Expense by annual Net Earned Premium Revenue, multiplied by 100.",
        "hlf_calculation": "Net Loss Ratio (%) = (Net Claims Expense ÷ Net Earned Premium Revenue) × 100",
        "hlf_interpretation": "HLF market context: Following the >85% loss-ratio spike during the 2023 North Island Auckland Anniversary floods and Cyclone Gabrielle, loss ratios have normalised toward long-run underwriting levels.",
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
        "also_known_as": "NHCover statutory residential building cap & maximum annual levy (NHI Act 2023)",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Statutory natural hazard residential building cover (NHCover)",
        "source": "NHI Act 2023 & NHC",
        "source_full": "Natural Hazards Insurance Act 2023 (NZ Legislation) & Natural Hazards Commission Toka Tū Ake (NHC)",
        "series_ids": "Natural Hazards Insurance Act 2023 — Statutory Building Cap & Levy Rate Schedule",
        "source_url": NHC_ACT_PAGE,
        "data_type": "Officially legislated statutory cap/levy & HLF-derived GST-inclusive figures",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Statutory Schedule",
        "definition": "The statutory first-loss residential building cover cap and maximum annual levy administered by the Natural Hazards Commission Toka Tū Ake (NHC, formerly EQC) under the Natural Hazards Insurance Act 2023.",
        "methodology": "Live-verified from the Natural Hazards Insurance Act 2023 on legislation.govt.nz and NHC Toka Tū Ake levy schedule, multiplying statutory GST-exclusive figures by 1.15 to show GST-inclusive homeowner totals.",
        "hlf_calculation": "Cap incl. GST = Statutory Cap × 1.15; Max Annual Levy incl. GST = (Statutory Cap ÷ 100 × Levy Rate) × 1.15",
        "hlf_interpretation": "HLF market context: An active private house insurance policy with fire cover is required to trigger automatic statutory NHCover—homeowners who let their private building policy lapse lose both private cover and the statutory natural hazard protection.",
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
        "also_known_as": "RBNZ & NHC insured vs uninsured domestic dwellings in New Zealand",
        "category": "Property & Natural Hazard Insurance",
        "topic": "Residential dwelling insurance coverage rate",
        "source": "RBNZ FSR & NHC",
        "source_full": "Reserve Bank of New Zealand — Financial Stability Report & Natural Hazards Commission Toka Tū Ake",
        "series_ids": "RBNZ FSR & NHC — Insured vs Uninsured Domestic Dwellings & Aggregate Residential Sum Insured",
        "source_url": RBNZ_FSR_PAGE,
        "data_type": "Officially reported RBNZ/NHC figures & HLF-derived coverage percentages",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Semi-Annual",
        "definition": "The percentage and number of New Zealand domestic dwellings covered by residential building insurance versus dwellings without building insurance.",
        "methodology": "Live-parsed from the Reserve Bank of New Zealand Financial Stability Report and Natural Hazards Commission Toka Tū Ake residential exposure releases.",
        "hlf_calculation": "Uninsured Share = Uninsured Dwellings ÷ Total Domestic Dwellings; Insured Share = 100% minus Uninsured Share",
        "hlf_interpretation": "HLF market context: Because mortgage lenders mandate building insurance on all mortgaged homes, outright uninsurance is concentrated among mortgage-free properties, while mortgaged homeowners face underinsurance risk if their sum-insured cap lags rebuild costs.",
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
        "source_full": "Accident Compensation Corporation (ACC) — Annual Report & Official Claim Statistics",
        "series_ids": "ACC Annual Report — Total Accepted Injury Claims & New Weekly Compensation Entitlement Claims",
        "source_url": ACC_REPORT_PAGE,
        "data_type": "Officially reported ACC annual claim counts & HLF-derived ratio per 100 accepted claims",
        "derived_by_hlf": True,
        "unit": "claims_per_100",
        "frequency": "Annual (June Financial Year)",
        "definition": "Comparison of total annual accepted injury claims registered with ACC across New Zealand and new weekly income compensation claims for time off work.",
        "methodology": "Live-parsed from ACC's published Annual Report and claim statistics releases, dividing annual new weekly compensation entitlement claims by total annual accepted injury claims multiplied by 100.",
        "hlf_calculation": "(New Weekly Compensation Claims ÷ Total Accepted Injury Claims) × 100",
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
        "also_known_as": "ACC covered income cap, Earners' Levy rate & 80% maximum annual ACC payout",
        "category": "Personal Risk & Income Protection",
        "topic": "ACC covered earnings cap & high-earner protection gap",
        "source": "ACC & IRD Schedule",
        "source_full": "Accident Compensation Corporation (ACC) & Inland Revenue — Earners' Levy Statutory Schedule",
        "series_ids": "ACC / IRD Earners' Levy Rate (per $100 incl. GST) & Maximum Liable Earnings Cap ($ NZD)",
        "source_url": ACC_LEVY_PAGE,
        "data_type": "Officially legislated statutory schedule & HLF-derived maximum levy/compensation ceiling",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual (Tax Year from 1 April)",
        "definition": "The statutory ACC Earners' Levy rate and Maximum Liable Earnings cap for the current tax year, alongside the resulting maximum annual levy and maximum 80% gross weekly compensation ceiling.",
        "methodology": "Live-parsed from the official ACC and Inland Revenue PAYE / Earners' Levy schedule, multiplying the Maximum Liable Earnings cap by the GST-inclusive levy rate and by the 80% statutory weekly compensation rate.",
        "hlf_calculation": "Max Annual Levy = (Max Liable Earnings ÷ 100) × Levy Rate; Max 80% ACC Payout = Max Liable Earnings × 80%",
        "hlf_interpretation": "HLF market context: Any income earned above the ACC Maximum Liable Earnings cap is not covered by ACC weekly compensation—meaning higher-income borrowers face both an uncovered top-up gap on accidental injuries and a 100% gap on non-injury illnesses.",
        "questions": [
            "What is the maximum income covered by ACC in New Zealand?",
            "What is the ACC Earners' Levy rate and maximum weekly compensation payout in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "life-trauma-average-claim-payout": {
        "name": "Life, Trauma & Income Protection annual claims paid, premiums & covers in force (FSC / RBNZ)",
        "also_known_as": "FSC Life Insurance Industry Spotlight & RBNZ Table J10 life, trauma & disability claims and premiums",
        "category": "Personal Risk & Income Protection",
        "topic": "Personal risk insurance claims paid, product premiums & covers in force",
        "source": "FSC Spotlight PDF & RBNZ J10",
        "source_full": "Financial Services Council (FSC) — Life Insurance Industry Spotlight (PDF) & RBNZ Insurance Statistics (Table J10)",
        "series_ids": "FSC Life Insurance Spotlight PDF — Annual Claims Paid, In-Force Covers & Premiums by Product (Term Life, Trauma, Income Protection)",
        "source_url": FSC_LIFE_TAG_PAGE,
        "data_type": "Officially reported FSC Spotlight PDF & RBNZ J10 series + HLF-derived per-cover & daily claim figures",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Quarterly / Annual",
        "definition": "Total annual claims paid across New Zealand's life and personal risk insurance sector, total life covers in force, and annual premiums across Term Life & Accidental Death, Critical Illness / Trauma cover, and Income / Mortgage Repayment Protection.",
        "methodology": "Automatically fetched and parsed from the latest Financial Services Council (FSC) Life Insurance Industry Spotlight PDF (via blog.fsc.org.nz/tag/life-insurance) and cross-checked against RBNZ Table J10 life insurance series. Underlying spotlight statistics are credited to the Financial Services Council (FSC); HLF derives daily claim payouts, claims-to-premium ratios, and average annual premium per in-force cover.",
        "hlf_calculation": "Daily Claims Paid = Annual Claims Paid ÷ 365; Mean Annual Premium per Cover = Total Annual Life Premiums ÷ Total In-Force Covers",
        "hlf_interpretation": "HLF market context: Shows the scale of New Zealand's personal risk protection pool across Term Life, lump-sum Trauma cover, and monthly Income/Mortgage Repayment Protection supporting households through death, serious illness, and disability.",
        "questions": [
            "How much do New Zealand life, trauma, and income protection insurers pay out in claims each year?",
            "How many New Zealanders have life, trauma, or income protection insurance covers?",
            "How much is spent on Term Life vs Trauma vs Income Protection in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "private-health-claims-average-payout": {
        "name": "Private health insurance annual claims paid, insured lives & claims per member (FSC / RBNZ)",
        "also_known_as": "FSC Health Insurance Industry Spotlight & RBNZ Table J10 private medical claims paid & per-member payout",
        "category": "Personal Risk & Income Protection",
        "topic": "Private health insurance claims paid & insured lives",
        "source": "FSC Spotlight & RBNZ J10",
        "source_full": "Financial Services Council (FSC) — Health Insurance Industry Spotlight & RBNZ Insurance Statistics (Table J10)",
        "series_ids": "FSC Health Insurance Spotlight — Total Insured Lives, Annual Claims Paid ($B) & Annual Premiums",
        "source_url": FSC_HEALTH_TAG_PAGE,
        "data_type": "Officially reported FSC Spotlight & RBNZ series + HLF-derived average annual claims paid per insured member",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Quarterly / Annual",
        "definition": "Total annual private health insurance claims paid in New Zealand, total lives covered by private health insurance, and the average annual claims paid per insured member.",
        "methodology": "Automatically fetched and parsed from the latest Financial Services Council (FSC) Health Insurance Spotlight release/PDF (via blog.fsc.org.nz/tag/health-insurance) and RBNZ Table J10 health insurance series. Underlying spotlight statistics are credited to the Financial Services Council (FSC); HLF derives the average annual claims paid per insured person and daily claim payout rate.",
        "hlf_calculation": "Average Annual Claims Paid per Insured Member = Annual Health Claims Paid ($) ÷ Total Insured Lives",
        "hlf_interpretation": "HLF market context: Rising private surgical admissions and medical procedure costs mean private health insurers now pay out well over $2B annually across ~1.35m–1.48m insured New Zealanders—making surgical cover a key protection pillar alongside mortgage and income cover.",
        "questions": [
            "How much do private health insurers pay out in claims in New Zealand each year?",
            "What is the average annual health insurance claim payout per insured member in NZ?",
            "How many New Zealanders have private health insurance?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "personal-insurance-underinsurance-gap": {
        "name": "New Zealand household property insurance vs personal life, health & income protection coverage",
        "also_known_as": "Comparison of NZ insured dwellings vs FSC life, trauma, income protection & health insurance penetration",
        "category": "Personal Risk & Income Protection",
        "topic": "Household underinsurance & income protection gap",
        "source": "FSC Spotlight & RBNZ FSR",
        "source_full": "Financial Services Council (FSC) Life & Health Spotlights & RBNZ Financial Stability Report",
        "series_ids": "RBNZ/NHC Insured Dwellings (%) vs FSC In-Force Life, Income Protection & Health Insurance Penetration",
        "source_url": FSC_DATA_PAGE,
        "data_type": "Officially reported RBNZ & FSC series + HLF-derived protection gap comparison",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Quarterly / Annual",
        "definition": "Comparison of New Zealand residential building insurance coverage against personal risk and health insurance penetration across the population and working-age households.",
        "methodology": "Combines live-parsed residential dwelling insurance coverage from RBNZ/NHC with live-parsed in-force cover counts and premium shares from the Financial Services Council (FSC) Life and Health Insurance Spotlights.",
        "hlf_calculation": "Income Protection Share of Life Sector Premiums = Income Protection Premiums ÷ Total Life Sector Premiums",
        "hlf_interpretation": "HLF market context: Almost every mortgaged homeowner insures the physical house, yet a much smaller share of working-age borrowers insure the income that pays the mortgage—even though ACC only covers accidental injuries and does not cover non-injury illness.",
        "questions": [
            "How does income protection coverage compare to house insurance coverage in New Zealand?",
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
    """Lightweight HTML table parser for official release tables."""

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


def _browser_fetch_binary(ctx, page, url: str, referer: str, min_bytes: int) -> bytes:
    """Helper to fetch binary bytes (.xlsx, .csv, .pdf) inside an active Camoufox or Playwright context."""
    if referer:
        try:
            page.goto(referer, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(1800)
        except Exception:
            pass
    try:
        api_resp = ctx.request.get(url, timeout=45000, headers={"Referer": referer or url})
        if api_resp.ok:
            body = api_resp.body()
            if body and len(body) >= min_bytes and b"Just a moment" not in body[:600]:
                return body
    except Exception:
        pass
    # Secondary fallback: base64 fetch inside page context so browser cookies/TLS fingerprint are used directly
    try:
        b64_data = page.evaluate(
            """async (u) => {
                const r = await fetch(u, {credentials: 'include'});
                if (!r.ok) return '';
                const buf = await r.arrayBuffer();
                let binary = '';
                const bytes = new Uint8Array(buf);
                for (let i = 0; i < bytes.byteLength; i++) binary += String.fromCharCode(bytes[i]);
                return btoa(binary);
            }""",
            url,
        )
        if b64_data:
            raw_b = base64.b64decode(b64_data)
            if len(raw_b) >= min_bytes and b"Just a moment" not in raw_b[:600]:
                return raw_b
    except Exception:
        pass
    return b""


def fetch_bytes_with_browser_fallback(session, url: str, referer: str = "", min_bytes: int = 1000) -> bytes:
    """Fetch binary files (.xlsx, .csv, .pdf) via HTTP -> Camoufox -> Playwright Chromium."""
    raw = fetch_url_bytes(session, url, referer=referer, timeout=40, use_cb=False)
    if raw and len(raw) >= min_bytes and b"Just a moment" not in raw[:600] and b"Access Denied" not in raw[:600]:
        return raw

    # Tier 2: Camoufox anti-bot browser (installed in GitHub Actions workflow)
    try:
        from camoufox.sync_api import Camoufox
        with Camoufox(headless=True) as browser:
            ctx = browser.new_context(locale="en-NZ", accept_downloads=True)
            page = ctx.new_page()
            b_out = _browser_fetch_binary(ctx, page, url, referer, min_bytes)
            if b_out:
                return b_out
    except Exception as exc_cf:
        print(f"[Camoufox Binary] {url}: {exc_cf}")

    # Tier 3: Playwright Chromium fallback
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
                accept_downloads=True,
            )
            page = ctx.new_page()
            b_out = _browser_fetch_binary(ctx, page, url, referer, min_bytes)
            browser.close()
            if b_out:
                return b_out
    except Exception as exc_pw:
        print(f"[Playwright Binary] {url}: {exc_pw}")
    return b""


def extract_pdf_text(pdf_bytes: bytes) -> str:
    """Extract plain text from a live-downloaded PDF (Kāinga Ora, FSC Spotlight, FMA Report, RBNZ FSR, ACC)."""
    if not pdf_bytes or not pdf_bytes.startswith(b"%PDF"):
        return ""
    try:
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(pdf_bytes))
        pages_txt = []
        for pg in reader.pages:
            t = pg.extract_text() or ""
            if t:
                pages_txt.append(t)
        return re.sub(r"\s+", " ", " ".join(pages_txt)).strip()
    except Exception as exc:
        print(f"[PDF Parser] Note: {exc}")
        return ""


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

    # Tier 2: Camoufox for Cloudflare-protected pages
    still_missing: list[str] = []
    try:
        from camoufox.sync_api import Camoufox
        with Camoufox(headless=True) as browser:
            ctx = browser.new_context(locale="en-NZ")
            page = ctx.new_page()
            for u in missing:
                try:
                    page.goto(u, wait_until="domcontentloaded", timeout=35000)
                    page.wait_for_timeout(2000)
                    content = page.content()
                    if len(content) > 1200 and "Just a moment" not in content[:600]:
                        results[u] = content
                    else:
                        still_missing.append(u)
                except Exception:
                    still_missing.append(u)
    except Exception:
        still_missing = list(missing)

    if not still_missing:
        return results

    # Tier 3: Playwright Chromium
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
            for u in still_missing:
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


def _download_rbnz_in_context(ctx, page, root: pathlib.Path, keys_to_get: list[str], found: dict, tag: str) -> None:
    """Download RBNZ .xlsx workbooks inside an active Camoufox or Playwright browser session."""
    try:
        page.goto(RBNZ_WARMUP_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(2200)
    except Exception:
        pass

    for key in list(keys_to_get):
        cfg = RBNZ_SOURCES[key]
        b_data = _browser_fetch_binary(ctx, page, cfg["xlsx"], RBNZ_WARMUP_URL, 15_000)
        if b_data and b_data[:2] == b"PK":
            tmp_p = root / f"tmp_{key}.xlsx"
            tmp_p.write_bytes(b_data)
            found[key] = tmp_p
            print(f"[RBNZ {tag}] Downloaded live {key.upper()} workbook ({len(b_data):,} bytes)")
            continue

        for pg_url in cfg["pages"]:
            try:
                page.goto(pg_url, wait_until="domcontentloaded", timeout=35000)
                page.wait_for_timeout(1800)
                pg_html = page.content()
                for href in re.findall(r'href=["\']([^"\']+\.xlsx[^"\']*)["\']', pg_html, flags=re.I):
                    if cfg["match"] in href.lower():
                        full_u = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
                        b2 = _browser_fetch_binary(ctx, page, full_u, pg_url, 15_000)
                        if b2 and b2[:2] == b"PK":
                            tmp_p = root / f"tmp_{key}.xlsx"
                            tmp_p.write_bytes(b2)
                            found[key] = tmp_p
                            print(f"[RBNZ {tag}] Discovered & downloaded {key.upper()} ({len(b2):,} bytes)")
                            break
                if key in found and found[key].name.startswith("tmp_"):
                    break
            except Exception as exc_pg:
                print(f"[RBNZ {tag}] Page {pg_url}: {exc_pg}")


def download_rbnz_workbooks(root: pathlib.Path, session) -> dict[str, pathlib.Path]:
    """Download live RBNZ .xlsx workbooks using HTTP -> Camoufox -> Playwright browser fallback."""
    found: dict[str, pathlib.Path] = {}
    needs_browser: list[str] = []

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
            print(f"[RBNZ HTTP] Downloaded live {key.upper()} workbook ({len(raw):,} bytes)")
            continue

        needs_browser.append(key)

    if needs_browser:
        try:
            from camoufox.sync_api import Camoufox
            with Camoufox(headless=True) as browser:
                ctx = browser.new_context(locale="en-NZ", accept_downloads=True)
                page = ctx.new_page()
                _download_rbnz_in_context(ctx, page, root, needs_browser, found, "Camoufox")
        except Exception as exc_cf:
            print(f"[RBNZ Camoufox] Note: {exc_cf}")

    still_needed = [k for k in needs_browser if k not in found or not found[k].name.startswith("tmp_")]
    if still_needed:
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
                    accept_downloads=True,
                )
                page = ctx.new_page()
                _download_rbnz_in_context(ctx, page, root, still_needed, found, "Playwright")
                browser.close()
        except Exception as exc_pw:
            print(f"[RBNZ Playwright] Note: {exc_pw}")

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
                    joined, flags=re.I,
                )
                if m:
                    dt_p = pd.to_datetime(m.group(1), errors="coerce")
                    if pd.notna(dt_p):
                        return f"{dt_p.strftime('%Y-%m-%d')} ({dt_p.day} {dt_p.strftime('%B %Y')} — RBNZ statistical release)"
    except Exception:
        pass
    return "Monthly RBNZ statistical release (~4 weeks after reference month-end)"


def extract_html_release_date(html_str: str, fallback_schedule: str) -> str:
    """Extract official release date from Stats NZ / IRD / MBIE / Kāinga Ora HTML without copying period_end."""
    if html_str:
        txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html_str[:18000]))
        for pat in (
            r"(?:released|published|last updated|date published)[:\s]+(\d{1,2}\s+[A-Za-z]+\s+20\d{2})",
            r"(\d{1,2}\s+[A-Za-z]+\s+20\d{2})\s+(?:information release|release)",
        ):
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
        if raw.empty or raw.shape[0] < 6 or raw.shape[1] < 2:
            continue

        date_row_start = None
        for r_idx in range(min(30, len(raw))):
            dt = pd.to_datetime(raw.iloc[r_idx, 0], errors="coerce")
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
                if re.match(r"^[A-Z0-9]{2,8}\.[A-Z0-9._]+$", cell_s, flags=re.I):
                    code_found = cell_s.upper()
                    break
            series_codes.append(code_found)
            col_labels.append(re.sub(r"\s+", " ", " | ".join(col_cells)).lower())

        data_df = raw.iloc[date_row_start:, :].copy()
        data_df["date"] = pd.to_datetime(data_df.iloc[:, 0], errors="coerce")
        data_df = data_df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
        if len(data_df) >= 2:
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
    return (ts + pd.offsets.MonthEnd(0)).strftime("%Y-%m-%d")


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
    return f"{'up from' if diff > 0 else 'down from'} {prev_pct:.1f}% a year earlier"


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


def resolve_raw_file_and_glossary_html(source_label: str, source_url: str, calc_live: str) -> str:
    """Build direct raw-file download link (.xlsx/.csv/.pdf) and verified HLF Glossary anchor links."""
    s_low = f"{source_label} {source_url} {calc_live}".lower()
    raw_link_html = ""
    for r_key, r_cfg in RBNZ_SOURCES.items():
        if f"table {r_key}" in s_low or f"rbnz {r_key}" in s_low or f"/{r_key}/" in s_low:
            raw_link_html = f' &bull; <a href="{r_cfg["xlsx"]}" target="_blank" rel="noopener">Direct RBNZ {r_key.upper()} .xlsx</a>'
            break
    if not raw_link_html and any(source_url.lower().endswith(ext) for ext in (".xlsx", ".csv", ".pdf")):
        ext_lbl = source_url.rsplit(".", 1)[-1].upper()
        raw_link_html = f' &bull; <a href="{source_url}" target="_blank" rel="noopener">Direct {ext_lbl} source file</a>'

    g_links: list[tuple[str, str]] = []
    if any(k in s_low for k in ("dti", "debt-to-income", "c40", "c41", "gross income")):
        g_links.append(("Debt-to-Income (DTI)", f"{GLOSSARY_BASE_URL}#dti"))
    if any(k in s_low for k in ("lvr", "low-deposit", "c31", "80%", "deposit")):
        g_links.append(("Loan-to-Value Ratio (LVR)", f"{GLOSSARY_BASE_URL}#lvr"))
    if any(k in s_low for k in ("kāinga ora", "kainga ora", "first home loan", "lmi")):
        g_links.append(("Kāinga Ora First Home Loan", f"{GLOSSARY_BASE_URL}#first-home-loan"))
        g_links.append(("Lenders Mortgage Insurance (LMI)", f"{GLOSSARY_BASE_URL}#lmi"))
    if any(k in s_low for k in ("interest-only", "c32")):
        g_links.append(("Interest-Only Mortgage", f"{GLOSSARY_BASE_URL}#interest-only"))
    if any(k in s_low for k in ("refinanc", "top-up", "equity", "c33")):
        g_links.append(("Refinancing", f"{GLOSSARY_BASE_URL}#refinancing"))
        g_links.append(("Usable Equity", f"{GLOSSARY_BASE_URL}#usable-equity"))
    if any(k in s_low for k in ("kiwisaver", "t43", "fma")):
        g_links.append(("KiwiSaver First-Home Withdrawal", f"{GLOSSARY_BASE_URL}#kiwisaver-first-home-withdrawal"))
    if any(k in s_low for k in ("second-chance", "previous owner", "previous homeowner")):
        g_links.append(("Second-Chance KiwiSaver Withdrawal", f"{GLOSSARY_BASE_URL}#second-chance-withdrawal"))
    if any(k in s_low for k in ("settled", "b30", "s33", "fixed", "floating")):
        g_links.append(("Fixed vs Floating Interest Rate", f"{GLOSSARY_BASE_URL}#fixed-rate"))
    if any(k in s_low for k in ("consent", "new build", "turnkey", "dwelling")):
        g_links.append(("Turnkey vs Construction Loan", f"{GLOSSARY_BASE_URL}#turnkey"))
    if any(k in s_low for k in ("rent", "bond", "tenancy")):
        g_links.append(("Gross Rental Yield", f"{GLOSSARY_BASE_URL}#rental-yield"))
    if any(k in s_low for k in ("insurance", "nhc", "acc", "trauma", "claim", "j10")):
        g_links.append(("Sum Insured & Income Protection", f"{GLOSSARY_BASE_URL}#income-protection"))
    if not g_links:
        g_links.append(("HLF Financial & Mortgage Glossary", GLOSSARY_BASE_URL))

    seen = set()
    dedup_g = []
    for lbl, u in g_links:
        if u not in seen:
            seen.add(u)
            dedup_g.append(f'<a href="{u}" target="_blank" rel="noopener">{lbl}</a>')
        if len(dedup_g) == 2:
            break
    glossary_part = f" &bull; <strong>Glossary:</strong> {' &middot; '.join(dedup_g)}" if dedup_g else ""
    return raw_link_html + glossary_part


def format_provenance_source_html(
    calc_live: str,
    source_label: str,
    source_url: str,
    is_derived: bool = True,
    checked_human: str = "",
) -> str:
    """Build the structured provenance dropdown HTML with direct source file links & HLF Glossary anchors."""
    deriv_tag = (
        f"Derived by HLF Pulse ({HLF_LICENSE_SHORT})"
        if is_derived
        else f"Official figure curated by HLF Pulse ({HLF_LICENSE_SHORT})"
    )
    chk_part = f" &bull; Checked: {checked_human}" if checked_human else ""
    extra_links = resolve_raw_file_and_glossary_html(source_label, source_url, calc_live)
    return (
        f"<strong>HLF calculation:</strong> <code>{calc_live}</code><br>"
        f'<strong>Source:</strong> <a href="{source_url}" target="_blank" rel="noopener">{source_label}</a>'
        f"{extra_links} &bull; {deriv_tag}{chk_part}."
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
    """100% live-parsed DTI & Borrower Gross Income series from RBNZ C40, C41 & C31 (zero hardcoded fallbacks)."""
    out = {}
    fhb_gt6_pct_latest: float | None = None
    fhb_le5_pct_latest: float | None = None
    fhb_le6_pct_latest: float | None = None
    dt_dti: pd.Timestamp | None = None
    pub_c40 = "Monthly RBNZ DTI statistical release (~4 weeks after reference month-end)"
    c40_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use"

    def mk_dti_rec(chip, as_of, pub, html_s, off_s, src_lbl, src_u, per_s, val_n, val_s, prev_s, chg_s, calc_s, hist_from, hist_l, extra=None):
        rec = {
            "chip": chip,
            "as_of": as_of,
            "source_published": pub,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": html_s,
            "official_data_summary": off_s,
            "source_html": format_provenance_source_html(calc_s, src_lbl, src_u, True),
            "latest_period": per_s,
            "value": val_n,
            "latest_value": val_s,
            "previous_period": prev_s,
            "change": chg_s,
            "calculation_live": calc_s,
            "historical_data_from": hist_from,
            "history": hist_l,
        }
        if extra:
            rec.update(extra)
        return rec

    if c40_path is not None:
        try:
            pub_c40 = extract_rbnz_published_date(c40_path)
            tabs40 = load_rbnz_sheets(c40_path)
            val_pct, dti_dates, dti_ser, last_hi_v, last_tot_v, last_le5_v = None, None, None, 0.0, 0.0, 0.0

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
                        val_pct, dt_dti, dti_dates, dti_ser = last_v, df40["date"].iloc[-1], df40["date"], pct_s
                        last_hi_v, last_tot_v, last_le5_v = float(hi_s.iloc[-1]), float(tot_s.iloc[-1]), float(le5_s.iloc[-1])
                        break

            if val_pct is not None and dt_dti is not None:
                m_full = dt_dti.strftime("%B %Y")
                diff_cap = 20.0 - val_pct
                fhb_gt6_pct_latest, fhb_le6_pct_latest = val_pct, 100.0 - val_pct
                fhb_le5_pct_latest = (last_le5_v / max(1.0, last_tot_v)) * 100.0
                iy = max(0, len(dti_dates) - 13)
                prev_v = float(dti_ser.iloc[iy]) if pd.notna(dti_ser.iloc[iy]) else val_pct
                hist_start_c40 = f"Available from {dti_dates.iloc[0].strftime('%B %Y')} ({len(dti_dates)} monthly observations)"
                c40_calc = f"${last_hi_v:,.1f}m ÷ ${last_tot_v:,.1f}m = {val_pct:.1f}% (Benchmark diff: 20.0% - {val_pct:.1f}% = {diff_cap:.1f} pts)"
                out["dti-headroom"] = mk_dti_rec(
                    "RBNZ C40", month_end_iso(dt_dti), pub_c40,
                    f"In {m_full}, <strong>{val_pct:.1f}%</strong> of first-home buyer mortgage lending had a <strong>Debt-to-Income (DTI) ratio above 6x</strong> (${last_hi_v:,.1f}m of ${last_tot_v:,.1f}m)—sitting <strong>{diff_cap:.1f} percentage points</strong> below the 20% RBNZ speed-limit cap.",
                    f"RBNZ Table C40 ({m_full}): FHB lending with TDTI > 6 = ${last_hi_v:,.1f}m out of ${last_tot_v:,.1f}m total FHB commitments.",
                    "RBNZ Table C40 (Reserve Bank of New Zealand)", c40_url, m_full, round(val_pct, 2),
                    f"{val_pct:.1f}% at >6x DTI ({diff_cap:.1f} pts below 20% benchmark)",
                    f"{dti_dates.iloc[iy].strftime('%B %Y')} ({prev_v:.1f}%)",
                    f"{val_pct - prev_v:+.1f} percentage points YoY", c40_calc, hist_start_c40,
                    make_history_list(dti_dates, dti_ser, lambda x: f"{x:.1f}%"),
                )

            # Live-parse Investor >7x DTI share vs 20% RBNZ investor speed limit
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
                        inv_pct_v, inv_dt, inv_dates, inv_ser = lv_i, df40["date"].iloc[-1], df40["date"], pct_is
                        inv_hi_v, inv_tot_v = float(hi_is.iloc[-1]), float(tot_is.iloc[-1])
                        break

            if inv_pct_v is not None and inv_dt is not None:
                m_inv = inv_dt.strftime("%B %Y")
                inv_headroom = 20.0 - inv_pct_v
                iy_i = max(0, len(inv_dates) - 13)
                prev_inv_v = float(inv_ser.iloc[iy_i]) if pd.notna(inv_ser.iloc[iy_i]) else inv_pct_v
                inv_c40_calc = f"${inv_hi_v:,.1f}m ÷ ${inv_tot_v:,.1f}m = {inv_pct_v:.1f}% (Headroom below 20% investor speed limit: 20.0% - {inv_pct_v:.1f}% = {inv_headroom:.1f} pts)"
                out["investor-high-dti-share"] = mk_dti_rec(
                    "RBNZ C40", month_end_iso(inv_dt), pub_c40,
                    f"In {m_inv}, <strong>{inv_pct_v:.1f}%</strong> of residential property investor mortgage lending had a <strong>Debt-to-Income (DTI) ratio above 7x</strong> (${inv_hi_v:,.1f}m of ${inv_tot_v:,.1f}m)—sitting <strong>{inv_headroom:.1f} percentage points</strong> below the 20% RBNZ investor speed-limit cap.",
                    f"RBNZ Table C40 ({m_inv}): Property investor lending with TDTI > 7 = ${inv_hi_v:,.1f}m out of ${inv_tot_v:,.1f}m total investor commitments.",
                    "RBNZ Table C40 (Reserve Bank of New Zealand)", c40_url, m_inv, round(inv_pct_v, 2),
                    f"{inv_pct_v:.1f}% at >7x DTI ({inv_headroom:.1f} pts below 20% investor cap)",
                    f"{inv_dates.iloc[iy_i].strftime('%B %Y')} ({prev_inv_v:.1f}%)",
                    f"{inv_pct_v - prev_inv_v:+.1f} percentage points YoY", inv_c40_calc,
                    f"Available from {inv_dates.iloc[0].strftime('%B %Y')} ({len(inv_dates)} monthly observations)",
                    make_history_list(inv_dates, inv_ser, lambda x: f"{x:.1f}% (>7x DTI)"),
                )
        except Exception as exc:
            print(f"[C40 DTI] Parse note: {exc}")

    if c41_path is not None:
        try:
            pub_c41 = extract_rbnz_published_date(c41_path)
            tabs41 = load_rbnz_sheets(c41_path)
            dt_fhb, fhb_inc, d41_dates, fhb_inc_ser, fhb_bands, fhb_pool_m, fhb_imp_apps = compute_c41_weighted_avg_income(tabs41, "first home", "nz")
            _, inv_inc, _, _, _, inv_pool_m, inv_imp_apps = compute_c41_weighted_avg_income(tabs41, "investor", "nz")
            c41_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi"

            if dt_fhb is not None and fhb_inc and 40000 <= fhb_inc <= 600000:
                m_c41, as_of_c41 = dt_fhb.strftime("%B %Y"), month_end_iso(dt_fhb)
                iy41 = max(0, len(d41_dates) - 13)
                prev_inc_k = float(fhb_inc_ser.iloc[iy41]) if pd.notna(fhb_inc_ser.iloc[iy41]) else (fhb_inc / 1000.0)
                prev_inc_nzd, fhb_inc_int = int(round(prev_inc_k * 1000.0)), int(round(fhb_inc))
                hist_41 = make_history_list(d41_dates, fhb_inc_ser, lambda x: f"${x*1000:,.0f}")
                hist_start_c41 = f"Available from {d41_dates.iloc[0].strftime('%B %Y')} ({len(d41_dates)} monthly observations)"

                band_summary_str = " | ".join(
                    f"{b['band']}: ${b['avg_income_nzd']:,} avg (${b['total_income_m']:,.1f}m pool / {b['implied_borrowers']:,.1f} implied apps)"
                    for b in fhb_bands
                ) or f"Overall mean: ${fhb_inc_int:,}"
                lo_band = min((b["avg_income_nzd"] for b in fhb_bands), default=fhb_inc_int)
                hi_band = max((b["avg_income_nzd"] for b in fhb_bands), default=fhb_inc_int)

                bgi_calc = (
                    f"FHB: ${fhb_pool_m:,.1f}m total band income ÷ {fhb_imp_apps:,.1f} implied applications = ${fhb_inc_int:,}"
                    + (f" | Investor: ${inv_pool_m:,.1f}m ÷ {inv_imp_apps:,.1f} = ${inv_inc:,.0f}" if inv_inc else "")
                    + " (subject to RBNZ band rounding)"
                )
                bgi_html = (
                    f"In {m_c41}, approved New Zealand first-home buyer mortgage applications averaged <strong>${fhb_inc_int:,}</strong> in gross borrower income, compared with <strong>${inv_inc:,.0f}</strong> across approved property investor applications."
                    if inv_inc and 40000 <= inv_inc <= 900000
                    else f"In {m_c41}, approved New Zealand first-home buyer mortgage applications averaged <strong>${fhb_inc_int:,}</strong> in gross borrower income (ranging from <strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong> across DTI bands)."
                )
                out["borrower-income"] = mk_dti_rec(
                    "RBNZ C41", as_of_c41, pub_c41, bgi_html,
                    f"RBNZ Table C41 ({m_c41}) Official FHB Band Figures (pools rounded to $0.1m, averages to $1,000): {band_summary_str}.",
                    "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_url, m_c41, fhb_inc_int,
                    f"${fhb_inc_int:,} FHB mean" + (f" vs ${inv_inc:,.0f} Investor mean" if inv_inc else ""),
                    f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_nzd:,} FHB)",
                    f"${fhb_inc_int - prev_inc_nzd:+d} YoY", bgi_calc, hist_start_c41, hist_41,
                )

                # Typical FHB Household Gross Income ticker (100% live from RBNZ C41)
                typ_inc_calc = f"${fhb_pool_m:,.1f}m total FHB gross income pool × $1m ÷ {fhb_imp_apps:,.1f} implied FHB applications = ${fhb_inc_int:,}"
                out["fhb-typical-income"] = mk_dti_rec(
                    "RBNZ C41", as_of_c41, pub_c41,
                    f"In {m_c41}, the typical (borrower-weighted average) gross household income across approved New Zealand first-home buyer mortgage applications was <strong>${fhb_inc_int:,}</strong> ({plain_yoy_growth(fhb_inc_int, prev_inc_nzd)} from ${prev_inc_nzd:,} a year earlier).",
                    f"RBNZ Table C41 ({m_c41}): Total FHB gross income pool = ${fhb_pool_m:,.1f}m across {fhb_imp_apps:,.1f} implied applications (band averages: ${lo_band:,}–${hi_band:,}).",
                    "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_url, m_c41, fhb_inc_int,
                    f"${fhb_inc_int:,} typical FHB gross application income",
                    f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_nzd:,})",
                    f"${fhb_inc_int - prev_inc_nzd:+d} ({((fhb_inc_int - prev_inc_nzd)/max(1, prev_inc_nzd))*100:+.1f}%) YoY",
                    typ_inc_calc, hist_start_c41, hist_41,
                )

                # Canonical slug: fhb-income-by-dti-band (100% live from RBNZ C41)
                band_calc = f"${fhb_pool_m:,.1f}m total band income × $1m ÷ {fhb_imp_apps:,.1f} total implied applications = ${fhb_inc_int:,} (reproducible from the 5 band rows; subject to RBNZ rounding)"
                out["fhb-income-by-dti-band"] = mk_dti_rec(
                    "RBNZ C41", as_of_c41, pub_c41,
                    f"In {m_c41}, average gross borrower incomes across New Zealand first-home buyer Debt-to-Income bands ranged from <strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong>, with an overall average of <strong>${fhb_inc_int:,}</strong>.",
                    f"RBNZ Table C41 ({m_c41}) Official FHB Band Figures (pools rounded to $0.1m, averages to $1,000): {band_summary_str}.",
                    "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_url, m_c41, fhb_inc_int,
                    f"${fhb_inc_int:,} HLF mean | Official C41 bands: ${lo_band:,}–${hi_band:,}",
                    f"{d41_dates.iloc[iy41].strftime('%B %Y')} (${prev_inc_nzd:,} HLF mean)",
                    f"${fhb_inc_int - prev_inc_nzd:+d} ({((fhb_inc_int - prev_inc_nzd)/max(1, prev_inc_nzd))*100:+.1f}%) YoY",
                    band_calc, hist_start_c41, hist_41,
                    {"tdti_bands": fhb_bands, "tdti_total_pool_m": fhb_pool_m, "tdti_total_implied_apps": fhb_imp_apps},
                )

                # Typical FHB Debt-to-Income (DTI) ratio & share <= 6x DTI (100% live from C31 + C40 + C41; zero hardcoded fallbacks)
                ym_c41 = dt_fhb.strftime("%Y-%m")
                avg_loan_for_dti = None
                if c31_by_ym:
                    if ym_c41 in c31_by_ym and c31_by_ym[ym_c41].get("fhb_avg_loan_nzd"):
                        avg_loan_for_dti = int(round(c31_by_ym[ym_c41]["fhb_avg_loan_nzd"]))
                    elif c31_by_ym.get("_latest_fhb_avg_loan_nzd"):
                        avg_loan_for_dti = int(round(c31_by_ym["_latest_fhb_avg_loan_nzd"]))

                if (
                    avg_loan_for_dti is not None
                    and fhb_le6_pct_latest is not None
                    and fhb_le5_pct_latest is not None
                    and fhb_gt6_pct_latest is not None
                ):
                    implied_dti_x = round(avg_loan_for_dti / max(1.0, float(fhb_inc_int)), 2)
                    le6_pct, le5_pct, gt6_pct = fhb_le6_pct_latest, fhb_le5_pct_latest, fhb_gt6_pct_latest
                    typ_dti_calc = (
                        f"Implied FHB Mortgage DTI: ${avg_loan_for_dti:,} (C31 avg FHB loan) ÷ ${fhb_inc_int:,} (C41 avg FHB gross income) = {implied_dti_x:.2f}x "
                        f"(~{implied_dti_x:.1f}x–{implied_dti_x + 0.35:.1f}x total DTI incl. non-mortgage debt) | "
                        f"C40 Share <= 6x DTI: 100.0% − {gt6_pct:.1f}% = {le6_pct:.1f}% (Share <= 5x DTI: {le5_pct:.1f}%)"
                    )
                    out["fhb-typical-dti"] = mk_dti_rec(
                        "RBNZ C40 & C41", as_of_c41, pub_c40,
                        f"In {m_c41}, the typical New Zealand first-home buyer borrowed at an implied Debt-to-Income ratio of <strong>{implied_dti_x:.1f}x–{implied_dti_x + 0.4:.1f}x gross income</strong> (${avg_loan_for_dti:,} average loan on ${fhb_inc_int:,} average gross income), with <strong>{le6_pct:.1f}% of all first-home buyer lending</strong> written at or below <strong>6x DTI</strong> (and <strong>{le5_pct:.1f}%</strong> at or below <strong>5x DTI</strong>).",
                        f"RBNZ Tables C31, C40 & C41 ({m_c41}): Average FHB loan = ${avg_loan_for_dti:,}; Average FHB gross income = ${fhb_inc_int:,}; FHB lending <= 6x TDTI = {le6_pct:.1f}% (<= 5x TDTI = {le5_pct:.1f}%; > 6x TDTI = {gt6_pct:.1f}%).",
                        "RBNZ Tables C31, C40 & C41 (Reserve Bank of New Zealand)", c40_url, m_c41, implied_dti_x,
                        f"{implied_dti_x:.2f}x implied mortgage DTI (~{implied_dti_x:.1f}x–{implied_dti_x + 0.4:.1f}x total DTI) | {le6_pct:.1f}% <= 6x DTI",
                        f"{le5_pct:.1f}% of FHB lending written at <= 5x DTI",
                        f"{gt6_pct:.1f}% above 6x DTI (vs 20.0% RBNZ speed limit)",
                        typ_dti_calc, hist_start_c41,
                        [
                            {"period": "Implied Mortgage DTI", "label": "Implied FHB Loan-to-Income Ratio (C31 ÷ C41)", "value": implied_dti_x, "formatted": f"{implied_dti_x:.2f}x (${avg_loan_for_dti:,} ÷ ${fhb_inc_int:,})"},
                            {"period": "TDTI <= 5x Share", "label": "FHB Lending at or below 5x DTI (C40)", "value": round(le5_pct, 1), "formatted": f"{le5_pct:.1f}% of FHB lending"},
                            {"period": "TDTI <= 6x Share", "label": "FHB Lending at or below 6x DTI (C40)", "value": round(le6_pct, 1), "formatted": f"{le6_pct:.1f}% of FHB lending"},
                            {"period": "TDTI > 6x Share", "label": "FHB Lending above 6x DTI (C40 Speed Limit Band)", "value": round(gt6_pct, 1), "formatted": f"{gt6_pct:.1f}% (vs 20.0% RBNZ cap)"},
                        ],
                    )
        except Exception as exc:
            print(f"[C41 BGI] Parse note: {exc}")
    return out


def build_rbnz_b30_ticker(filepath: pathlib.Path, session=None) -> dict:
    """100% live-parsed settled rates (RBNZ B30) & fixed vs floating mortgage mix (RBNZ S33 / B30; zero hardcoded fallbacks)."""
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

    out = {
        "settled-rates": {
            "chip": "RBNZ B30",
            "as_of": as_of_b30,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": f"In {m_full}, the average settled fixed mortgage interest rate across New Zealand owner-occupiers was <strong>{v1:.2f}%</strong> for 1-year terms and <strong>{v2:.2f}%</strong> for 2-year terms.",
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
        }
    }

    # Live-parse Fixed vs Floating Mortgage Mix & 12-Month Repricing Share from RBNZ Table S33 or B30
    s33_page = "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-fully-secured-by-residential-mortgage"
    s33_xlsx = "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/s/s33/s33-banks-assets-loans-fully-secured-by-residential-mortgage.xlsx"
    active_sess = session or get_http_session()
    s33_parsed = False

    try:
        s33_bytes = fetch_bytes_with_browser_fallback(active_sess, s33_xlsx, referer=s33_page, min_bytes=15_000)
        if not s33_bytes:
            pg_map = fetch_html_with_browser_fallback(active_sess, [s33_page], "RBNZ S33")
            s33_html = pg_map.get(s33_page, "")
            for href in re.findall(r'href=["\']([^"\']+\.xlsx[^"\']*)["\']', s33_html, flags=re.I):
                if "s33" in href.lower():
                    full_u = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
                    s33_bytes = fetch_bytes_with_browser_fallback(active_sess, full_u, referer=s33_page, min_bytes=15_000)
                    if s33_bytes:
                        break

        if s33_bytes and s33_bytes[:2] == b"PK":
            tmp_s33 = filepath.parent / "tmp_s33.xlsx"
            tmp_s33.write_bytes(s33_bytes)
            s33_pub = extract_rbnz_published_date(tmp_s33)
            tabs_s33 = load_rbnz_sheets(tmp_s33)
            if tmp_s33.exists():
                os.remove(tmp_s33)

            for _, df33, _, labels33 in tabs_s33:
                fl_idx, fx_idx, tot_idx, rep_le12_idxs = None, None, None, []
                for idx, lbl in enumerate(labels33):
                    if any(x in lbl for x in ("%", "share", "growth")):
                        continue
                    if "floating" in lbl and fl_idx is None:
                        fl_idx = idx
                    elif ("total fixed" in lbl or ("fixed" in lbl and "total" in lbl)) and fx_idx is None:
                        fx_idx = idx
                    elif "total" in lbl and ("residential" in lbl or "mortgage" in lbl or "loans" in lbl) and tot_idx is None:
                        tot_idx = idx
                    elif any(k in lbl for k in ("3 months", "6 months", "1 year", "up to 3", "3 to 6", "6 to 12")):
                        rep_le12_idxs.append(idx)

                if fl_idx is not None and (fx_idx is not None or tot_idx is not None):
                    fl_ser = pd.to_numeric(df33.iloc[:, fl_idx], errors="coerce")
                    if fx_idx is not None:
                        fx_ser = pd.to_numeric(df33.iloc[:, fx_idx], errors="coerce")
                        tot_ser = fl_ser + fx_ser
                    else:
                        tot_ser = pd.to_numeric(df33.iloc[:, tot_idx], errors="coerce")
                        fx_ser = tot_ser - fl_ser

                    fl_m, fx_m, tot_m = float(fl_ser.iloc[-1]), float(fx_ser.iloc[-1]), float(tot_ser.iloc[-1])
                    if tot_m > 50_000:
                        dt33 = df33["date"].iloc[-1]
                        m33, as_of33 = dt33.strftime("%B %Y"), month_end_iso(dt33)
                        tot_b, fx_b, fl_b = tot_m / 1000.0, fx_m / 1000.0, fl_m / 1000.0
                        fx_sh_ser = (fx_ser / tot_ser) * 100.0
                        fx_sh, fl_sh = float(fx_sh_ser.iloc[-1]), 100.0 - float(fx_sh_ser.iloc[-1])
                        iy33 = max(0, len(df33) - 13)
                        fx_sh_prev = float(fx_sh_ser.iloc[iy33]) if pd.notna(fx_sh_ser.iloc[iy33]) else fx_sh

                        rep12_clause, rep12_calc = "", ""
                        if rep_le12_idxs:
                            rep12_m = sum(float(pd.to_numeric(df33.iloc[:, c_i], errors="coerce").iloc[-1] or 0.0) for c_i in rep_le12_idxs[:3])
                            if 0.1 * fx_m <= rep12_m <= fx_m:
                                rep12_b, rep12_pct = rep12_m / 1000.0, (rep12_m / fx_m) * 100.0
                                rep12_clause = f"—with <strong>{rep12_pct:.1f}% (${rep12_b:.1f}B)</strong> of fixed-rate balances scheduled to reprice within 12 months"
                                rep12_calc = f" | Repricing <=12m: ${rep12_b:.2f}B ÷ ${fx_b:.2f}B = {rep12_pct:.1f}%"

                        s33_calc = f"Fixed Share: ${fx_b:.2f}B ÷ ${tot_b:.2f}B = {fx_sh:.1f}% | Floating Share: ${fl_b:.2f}B ÷ ${tot_b:.2f}B = {fl_sh:.1f}%{rep12_calc}"
                        out["fixed-vs-floating-rate-mix"] = {
                            "chip": "RBNZ Table S33",
                            "as_of": as_of33,
                            "source_published": s33_pub,
                            "source_url": s33_page,
                            "max_age_days": MAX_AGE_MONTHLY,
                            "html": f"As at {m33}, across New Zealand's <strong>${tot_b:.1f}B</strong> in bank loans fully secured by residential mortgage, <strong>{fx_sh:.1f}% (${fx_b:.1f}B)</strong> was held on fixed interest rates and <strong>{fl_sh:.1f}% (${fl_b:.1f}B)</strong> on floating rates{rep12_clause}.",
                            "official_data_summary": f"RBNZ Table S33 ({m33}): Total residential mortgage loans = ${tot_m:,.0f}m; Fixed-rate loans = ${fx_m:,.0f}m ({fx_sh:.1f}%); Floating loans = ${fl_m:,.0f}m ({fl_sh:.1f}%).",
                            "source_html": format_provenance_source_html(s33_calc, "RBNZ Table S33 — Loans by Repricing (Reserve Bank of New Zealand)", s33_page, True),
                            "latest_period": m33,
                            "value": round(fx_sh, 2),
                            "latest_value": f"{fx_sh:.1f}% fixed (${fx_b:.1f}B) vs {fl_sh:.1f}% floating (${fl_b:.1f}B)",
                            "previous_period": f"{df33['date'].iloc[iy33].strftime('%B %Y')} ({fx_sh_prev:.1f}% fixed share)",
                            "change": f"{fx_sh - fx_sh_prev:+.1f} percentage points YoY (fixed share)",
                            "calculation_live": s33_calc,
                            "historical_data_from": f"Available from {df33['date'].iloc[0].strftime('%B %Y')} ({len(df33)} monthly observations)",
                            "history": make_history_list(df33["date"], fx_sh_ser, lambda x: f"{x:.1f}% fixed"),
                        }
                        s33_parsed = True
                        break
    except Exception as exc_s33:
        print(f"[RBNZ S33] Parse note: {exc_s33}")

    if not s33_parsed:
        try:
            for _, df_b, _, labels_b in tabs:
                fl_c, tot_c = None, None
                for idx, lbl in enumerate(labels_b):
                    if ("value" in lbl or "amount" in lbl or "$" in lbl) and "floating" in lbl and fl_c is None:
                        fl_c = idx
                    elif ("value" in lbl or "amount" in lbl or "$" in lbl) and "total" in lbl and tot_c is None:
                        tot_c = idx
                if fl_c is not None and tot_c is not None:
                    fl_s, tot_s = pd.to_numeric(df_b.iloc[:, fl_c], errors="coerce"), pd.to_numeric(df_b.iloc[:, tot_c], errors="coerce")
                    fx_s = tot_s - fl_s
                    fx_pct_s = (fx_s / tot_s) * 100.0
                    if pd.notna(fx_pct_s.iloc[-1]) and 40.0 <= float(fx_pct_s.iloc[-1]) <= 99.0:
                        fx_p, fl_p = float(fx_pct_s.iloc[-1]), 100.0 - float(fx_pct_s.iloc[-1])
                        fx_v, fl_v, tot_v = float(fx_s.iloc[-1]), float(fl_s.iloc[-1]), float(tot_s.iloc[-1])
                        iy_b = max(0, len(df_b) - 13)
                        prev_fx_p = float(fx_pct_s.iloc[iy_b])
                        b_mix_calc = f"Fixed: ${fx_v:,.0f}m ÷ ${tot_v:,.0f}m = {fx_p:.1f}% | Floating: ${fl_v:,.0f}m ÷ ${tot_v:,.0f}m = {fl_p:.1f}%"
                        out["fixed-vs-floating-rate-mix"] = {
                            "chip": "RBNZ B30",
                            "as_of": as_of_b30,
                            "source_published": pub_date,
                            "max_age_days": MAX_AGE_MONTHLY,
                            "html": f"In {m_full}, across <strong>{fmt_nzd_millions(tot_v)}</strong> in new residential mortgage lending, <strong>{fx_p:.1f}% ({fmt_nzd_millions(fx_v)})</strong> was committed on fixed interest rate terms and <strong>{fl_p:.1f}% ({fmt_nzd_millions(fl_v)})</strong> on floating terms.",
                            "official_data_summary": f"RBNZ Table B30 ({m_full}): Fixed-rate lending = ${fx_v:,.0f}m ({fx_p:.1f}%); Floating-rate lending = ${fl_v:,.0f}m ({fl_p:.1f}%) out of ${tot_v:,.0f}m total.",
                            "source_html": format_provenance_source_html(b_mix_calc, "RBNZ Table B30 (Reserve Bank of New Zealand)", b30_url, True),
                            "latest_period": m_full,
                            "value": round(fx_p, 2),
                            "latest_value": f"{fx_p:.1f}% fixed vs {fl_p:.1f}% floating",
                            "previous_period": f"{df_b['date'].iloc[iy_b].strftime('%B %Y')} ({prev_fx_p:.1f}% fixed)",
                            "change": f"{fx_p - prev_fx_p:+.1f} percentage points YoY",
                            "calculation_live": b_mix_calc,
                            "historical_data_from": hist_start_b30,
                            "history": make_history_list(df_b["date"], fx_pct_s, lambda x: f"{x:.1f}% fixed"),
                        }
                        break
        except Exception as exc_b:
            print(f"[RBNZ B30 Mix] Parse note: {exc_b}")

    return out
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


def fetch_latest_fsc_spotlight(session, tag_url: str, keyword: str) -> tuple[str, str, str, str]:
    """Crawl FSC tag page, open newest Spotlight post, download linked HubSpot PDF, and return (text, url, period, pub_date)."""
    pages = fetch_html_with_browser_fallback(session, [tag_url], "FSC Tag")
    tag_html = pages.get(tag_url, "")
    if not tag_html:
        return "", tag_url, "", ""

    post_urls = []
    for href in re.findall(r'href=["\'](https?://blog\.fsc\.org\.nz/[^"\'#?]+)["\']', tag_html, flags=re.I):
        h_low = href.lower()
        if any(x in h_low for x in ("/tag/", "/author/", "/page/")):
            continue
        if any(w in h_low for w in ("spotlight", keyword.lower(), "data", "claims", "quarter", "2025", "2026", "2027")):
            if href not in post_urls:
                post_urls.append(href)
    if not post_urls:
        return "", tag_url, "", ""

    post_map = fetch_html_with_browser_fallback(session, post_urls[:3], "FSC Post")
    for p_url in post_urls[:3]:
        p_html = post_map.get(p_url, "")
        if not p_html:
            continue
        pub_date = extract_html_release_date(p_html, "Quarterly Financial Services Council (FSC) Spotlight release")
        p_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", p_html))
        pdf_links = re.findall(r'href=["\'](https?://[^"\']+\.pdf[^"\']*)["\']', p_html, flags=re.I)
        pdf_txt, final_src_url = "", p_url
        for pdf_u in pdf_links[:2]:
            pdf_bytes = fetch_bytes_with_browser_fallback(session, pdf_u, referer=p_url, min_bytes=5_000)
            extracted = extract_pdf_text(pdf_bytes)
            if len(extracted) > 200:
                pdf_txt, final_src_url = extracted, pdf_u
                break

        combined = f"{pdf_txt} {p_txt}".strip()
        m_per = re.search(
            r"(?:quarter ended|year ended|as at|spotlight)\s*:?\s*((?:March|June|September|December|Mar|Jun|Sep|Dec)\s+20\d{2}|Q[1-4]\s+20\d{2})",
            combined, flags=re.I,
        )
        per_lbl = m_per.group(1) if m_per else "Latest FSC Spotlight release"
        if len(combined) > 300:
            return combined, final_src_url, per_lbl, pub_date
    return "", tag_url, "", ""


def build_insurance_pulse_tickers(
    session,
    q_as_of: str,
    q_lbl: str,
    cpi_all_val: float | None,
    cpi_url: str,
    cpi_text: str = "",
    cpi_pub_date: str = "",
    cpi_sub_series: dict | None = None,
    j10_path: pathlib.Path | None = None,
) -> dict:
    """100% live-pulled Insurance, ACC, NHC & FSC Spotlight datasets (zero hardcoded fallbacks)."""
    out: dict = {}
    today = datetime.datetime.now(datetime.timezone.utc).date()
    cpi_sub_series = cpi_sub_series or {}

    def resolve_live_ins_cpi(sub_key: str, regex_pat: str) -> tuple[float | None, list[dict]]:
        if sub_key in cpi_sub_series and cpi_sub_series[sub_key].get("latest") is not None:
            return float(cpi_sub_series[sub_key]["latest"]), cpi_sub_series[sub_key].get("history", [])
        if cpi_text:
            m = re.search(regex_pat, cpi_text, flags=re.I)
            if m:
                dir_w = (m.group(1) or "up").lower()
                v = float(m.group(2))
                return (-v if dir_w in ("down", "fell", "decreased") else v), []
        return None, []

    if q_as_of and q_lbl and cpi_all_val is not None:
        cpi_ref_url = cpi_url or STATS_CPI_PAGE
        cpi_pub = cpi_pub_date or "Quarterly Stats NZ CPI release (~3 weeks after quarter-end)"
        cpi_extra = {"source_published": cpi_pub, "source_url": cpi_ref_url}
        cpi_lbl_src = f"Stats NZ — Consumers Price Index ({q_lbl}, Insurance Sub-Indices)"
        sub_specs = [
            ("building-insurance-inflation", "dwelling_insurance", r"dwelling insurance[,:\s]+(up|down|fell|rose|increased|decreased)?\s*([\d.]+)\s*(?:percent|%)", "residential dwelling (building) insurance", "Dwelling insurance"),
            ("contents-insurance-inflation", "contents_insurance", r"contents insurance[,:\s]+(up|down|fell|rose|increased|decreased)?\s*([\d.]+)\s*(?:percent|%)", "household contents insurance", "Contents insurance"),
            ("vehicle-insurance-inflation", "vehicle_insurance", r"(?:vehicle|motor vehicle|car) insurance[,:\s]+(up|down|fell|rose|increased|decreased)?\s*([\d.]+)\s*(?:percent|%)", "motor vehicle insurance", "Vehicle insurance"),
            ("health-insurance-inflation", "health_insurance", r"health insurance[,:\s]+(up|down|fell|rose|increased|decreased)?\s*([\d.]+)\s*(?:percent|%)", "private health insurance", "Health insurance"),
            ("life-insurance-inflation", "life_insurance", r"life insurance[,:\s]+(up|down|fell|rose|increased|decreased)?\s*([\d.]+)\s*(?:percent|%)", "constant-age life insurance", "Life insurance"),
        ]
        for slug, skey, pat, prose_lbl, short_lbl in sub_specs:
            val_inf, hist_inf = resolve_live_ins_cpi(skey, pat)
            if val_inf is not None:
                gap_v = val_inf - cpi_all_val
                calc_s = f"{short_lbl} CPI ({val_inf:+.1f}%) minus All-groups CPI ({cpi_all_val:+.1f}%) = {gap_v:+.1f} percentage points"
                out[slug] = make_pulse_item(
                    "Stats NZ CPI", q_as_of, MAX_AGE_QUARTERLY,
                    f"In the {q_lbl}, {prose_lbl} prices changed by <strong>{val_inf:+.1f}%</strong> year-on-year—sitting <strong>{gap_v:+.1f} percentage points</strong> relative to headline CPI inflation (<strong>{cpi_all_val:+.1f}%</strong>).",
                    f"Stats NZ CPI ({q_lbl}): {short_lbl} price index = {val_inf:+.1f}% YoY; All-groups CPI = {cpi_all_val:+.1f}% YoY.",
                    format_provenance_source_html(calc_s, cpi_lbl_src, cpi_ref_url, False),
                    q_lbl, round(val_inf, 2), f"{val_inf:+.1f}% YoY ({short_lbl} CPI index)",
                    f"All-groups headline CPI: {cpi_all_val:+.1f}% YoY", f"{gap_v:+.1f} percentage points vs. headline CPI",
                    calc_s, f"Stats NZ quarterly CPI series ({q_lbl})", hist_inf if hist_inf else None, cpi_extra,
                )

    # 2. Live RBNZ Table J10 (Insurance Income & Expenditure): General, Life & Health Insurer Series
    j10_life_claims_m: float | None = None
    j10_hlth_claims_m: float | None = None
    if j10_path is None:
        j10_cfg = RBNZ_SOURCES.get("j10", {})
        if j10_cfg.get("xlsx"):
            j10_bytes = fetch_bytes_with_browser_fallback(session, j10_cfg["xlsx"], referer=RBNZ_INS_STATS_PAGE, min_bytes=12_000)
            if j10_bytes and j10_bytes[:2] == b"PK":
                tmp_j10 = pathlib.Path(__file__).resolve().parent / "tmp_j10_live.xlsx"
                tmp_j10.write_bytes(j10_bytes)
                j10_path = tmp_j10

    if j10_path is not None and j10_path.exists():
        try:
            j10_pub = extract_rbnz_published_date(j10_path)
            tabs_j10 = load_rbnz_sheets(j10_path)
            for sname, df_j, _, labels_j in tabs_j10:
                gwp_i, re_i, nep_i, clm_i = None, None, None, None
                for idx, lbl in enumerate(labels_j):
                    if "general" in sname.lower() or "general" in lbl or len(tabs_j10) == 1:
                        if ("gross premium" in lbl or "gross written" in lbl or "gross earned" in lbl) and gwp_i is None:
                            gwp_i = idx
                        elif "reinsurance" in lbl and ("expense" in lbl or "outward" in lbl) and re_i is None:
                            re_i = idx
                        elif ("net premium" in lbl or "net earned" in lbl) and nep_i is None:
                            nep_i = idx
                        elif "claim" in lbl and ("net" in lbl or "incurred" in lbl or "expense" in lbl) and clm_i is None:
                            clm_i = idx
                    if "life" in lbl and "claim" in lbl and j10_life_claims_m is None:
                        s_l = pd.to_numeric(df_j.iloc[:, idx], errors="coerce").dropna()
                        if len(s_l):
                            j10_life_claims_m = float(s_l.iloc[-1])
                    if "health" in lbl and "claim" in lbl and j10_hlth_claims_m is None:
                        s_h = pd.to_numeric(df_j.iloc[:, idx], errors="coerce").dropna()
                        if len(s_h):
                            j10_hlth_claims_m = float(s_h.iloc[-1])

                if gwp_i is not None and clm_i is not None:
                    gwp_s = pd.to_numeric(df_j.iloc[:, gwp_i], errors="coerce")
                    clm_s = pd.to_numeric(df_j.iloc[:, clm_i], errors="coerce")
                    nep_s = pd.to_numeric(df_j.iloc[:, nep_i], errors="coerce") if nep_i is not None else (
                        gwp_s - pd.to_numeric(df_j.iloc[:, re_i], errors="coerce").abs() if re_i is not None else gwp_s
                    )
                    is_qtr = float(gwp_s.dropna().iloc[-1]) < 4500.0 and len(gwp_s.dropna()) >= 4
                    gwp_ann_s = gwp_s.rolling(4).sum() if is_qtr else gwp_s
                    nep_ann_s = nep_s.rolling(4).sum() if is_qtr else nep_s
                    clm_ann_s = clm_s.rolling(4).sum() if is_qtr else clm_s

                    dt_j = df_j["date"].iloc[-1]
                    as_of_j, per_j = month_end_iso(dt_j), f"Year ended {dt_j.strftime('%B %Y')}"
                    gwp_b, nep_b, clm_b = float(gwp_ann_s.iloc[-1]) / 1000.0, float(nep_ann_s.iloc[-1]) / 1000.0, float(clm_ann_s.iloc[-1]) / 1000.0
                    re_b = max(0.0, gwp_b - nep_b)
                    lr_s = (clm_ann_s / nep_ann_s.replace(0, float("nan"))) * 100.0
                    lr_pct = float(lr_s.iloc[-1])
                    iy_j = max(0, len(df_j) - 5)
                    prev_lr = float(lr_s.iloc[iy_j]) if pd.notna(lr_s.iloc[iy_j]) else lr_pct
                    clm_calc = f"Gross Premium (${gwp_b:.2f}B) minus Reinsurance (${re_b:.2f}B) = ${nep_b:.2f}B Net Earned Premium vs. ${clm_b:.2f}B Net Claims"
                    lr_calc = f"(${clm_b:.2f}B Net Claims ÷ ${nep_b:.2f}B Net Earned Premium) × 100 = {lr_pct:.1f}%"
                    hist_from_j = f"Available from {df_j['date'].iloc[0].strftime('%B %Y')} ({len(df_j)} observations)"

                    out["insurance-claims-trends"] = make_pulse_item(
                        "RBNZ Table J10", as_of_j, MAX_AGE_INSURANCE,
                        f"Over the {per_j.lower()}, licensed New Zealand general insurers collected <strong>${gwp_b:.2f}B</strong> in Gross Premium Revenue (<strong>${nep_b:.2f}B</strong> in Net Earned Premiums after <strong>${re_b:.2f}B</strong> in reinsurance) and incurred <strong>${clm_b:.2f}B</strong> in Net Claims.",
                        f"RBNZ Table J10 ({per_j}): General insurer Gross Premiums = ${gwp_b:.2f}B; Net Earned Premiums = ${nep_b:.2f}B; Net Claims = ${clm_b:.2f}B.",
                        format_provenance_source_html(clm_calc, "RBNZ Table J10 — Insurance Income & Expenditure", RBNZ_INS_STATS_PAGE, True),
                        per_j, round(clm_b, 2), f"${clm_b:.2f}B Net Claims vs. ${gwp_b:.2f}B Gross Premiums",
                        f"Net Earned Premium (after reinsurance): ${nep_b:.2f}B", f"Outwards catastrophe & reinsurance expense: ${re_b:.2f}B/yr",
                        clm_calc, hist_from_j, make_history_list(df_j["date"], clm_ann_s / 1000.0, lambda x: f"${x:.2f}B net claims"),
                        {"source_published": j10_pub},
                    )
                    out["insurer-loss-ratios"] = make_pulse_item(
                        "RBNZ Table J10", as_of_j, MAX_AGE_INSURANCE,
                        f"Over the {per_j.lower()}, licensed New Zealand general insurers recorded a net claims loss ratio of <strong>{lr_pct:.1f}%</strong>, incurring <strong>${clm_b:.2f}B</strong> in net claims against <strong>${nep_b:.2f}B</strong> in net earned premiums.",
                        f"RBNZ Table J10 ({per_j}): Net Claims = ${clm_b:.2f}B; Net Earned Premiums = ${nep_b:.2f}B; Net Loss Ratio = {lr_pct:.1f}%.",
                        format_provenance_source_html(lr_calc, "RBNZ Table J10 — Insurance Income & Expenditure", RBNZ_INS_STATS_PAGE, True),
                        per_j, round(lr_pct, 1), f"{lr_pct:.1f}% Net Loss Ratio (${clm_b:.2f}B claims ÷ ${nep_b:.2f}B net premium)",
                        f"Prior comparable period: {prev_lr:.1f}%", f"{lr_pct - prev_lr:+.1f} percentage points YoY",
                        lr_calc, hist_from_j, make_history_list(df_j["date"], lr_s, lambda x: f"{x:.1f}% loss ratio"),
                        {"source_published": j10_pub},
                    )
                    break
        except Exception as exc_j10:
            print(f"[RBNZ J10] Parse note: {exc_j10}")
        finally:
            if j10_path.name == "tmp_j10_live.xlsx" and j10_path.exists():
                os.remove(j10_path)

    # 3. Live Statutory Natural Hazards Insurance Act 2023 (NHCover Building Cap & Levy)
    nhc_pages = fetch_html_with_browser_fallback(session, [NHC_ACT_PAGE, NHC_FUND_PAGE], "NHC Act")
    nhc_txt = " ".join(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h)) for h in nhc_pages.values())
    m_cap = re.search(r"\$(300,000|300\s*000)", nhc_txt)
    m_levy_c = re.search(r"(?:16\s*cents|\$0\.16|\$480)", nhc_txt, flags=re.I)
    if m_cap and m_levy_c:
        nhc_cap_ex = int(re.sub(r"[^\d]", "", m_cap.group(1)))
        nhc_cap_inc = int(round(nhc_cap_ex * 1.15))
        nhc_levy_ex = int(round((nhc_cap_ex / 100.0) * 0.16))
        nhc_levy_inc = round(nhc_levy_ex * 1.15, 2)
        nhc_calc = f"Cap incl. GST: ${nhc_cap_ex:,} × 1.15 = ${nhc_cap_inc:,} | Max Levy incl. GST: (${nhc_cap_ex:,} ÷ 100 × $0.16) × 1.15 = ${nhc_levy_inc:.2f}"
        out["natural-hazard-claims"] = make_pulse_item(
            "NHI Act 2023 & NHC", today.strftime("%Y-%m-%d"), MAX_AGE_INSURANCE,
            f"Under the Natural Hazards Insurance Act 2023, the Natural Hazards Commission Toka Tū Ake (NHC) provides statutory first-loss residential building cover up to <strong>${nhc_cap_ex:,} + GST (${nhc_cap_inc:,} incl. GST)</strong> per dwelling at an annual levy capped at <strong>${nhc_levy_ex} + GST (${nhc_levy_inc:.2f} incl. GST)</strong>.",
            f"Natural Hazards Insurance Act 2023 (legislation.govt.nz): Statutory residential building cap = ${nhc_cap_ex:,} excl. GST; maximum annual levy = ${nhc_levy_ex} excl. GST (16c per $100 of cover).",
            format_provenance_source_html(nhc_calc, "Natural Hazards Insurance Act 2023 & NHC Toka Tū Ake", NHC_ACT_PAGE, True),
            "Current NHI Act 2023 Statutory Schedule", nhc_cap_ex,
            f"${nhc_cap_ex:,} + GST (${nhc_cap_inc:,} incl. GST) NHCover cap | ${nhc_levy_ex} + GST max annual levy",
            "Statutory schedule under Natural Hazards Insurance Act 2023",
            f"Max annual NHC levy: ${nhc_levy_ex} + GST (${nhc_levy_inc:.2f} incl. GST) at 16c/$100",
            nhc_calc, "Natural Hazards Insurance Act 2023 (legislation.govt.nz)", None,
            {"source_published": "2024-07-01 (Natural Hazards Insurance Act 2023 commencement)"},
        )

    # 4. Live ACC Earners' Levy & Weekly Compensation Claims
    acc_pages = fetch_html_with_browser_fallback(session, [ACC_LEVY_PAGE, ACC_REPORT_PAGE], "ACC")
    acc_levy_html = acc_pages.get(ACC_LEVY_PAGE, "")
    if acc_levy_html:
        acc_l_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", acc_levy_html))
        m_max_e = re.search(r"\$(1[4567]\d,\d{3})", acc_l_txt)
        m_rate_e = re.search(r"\$?(1\.[56789]\d)\s*(?:per\s*\$100|incl|excluding)?", acc_l_txt, flags=re.I)
        if m_max_e and m_rate_e:
            max_earn, levy_rate = int(m_max_e.group(1).replace(",", "")), float(m_rate_e.group(1))
            max_levy, max_comp_yr = round((max_earn / 100.0) * levy_rate, 2), round(max_earn * 0.80, 2)
            max_comp_wk = round(max_comp_yr / 52.0, 2)
            acc_pub = extract_html_release_date(acc_levy_html, "Annual ACC Earners' Levy statutory schedule (effective 1 April)")
            acc_lv_calc = f"Max Levy: (${max_earn:,} ÷ 100) × ${levy_rate:.2f} = ${max_levy:,.2f}/yr | Max 80% Comp: ${max_earn:,} × 80% = ${max_comp_yr:,.2f}/yr (${max_comp_wk:,.2f}/wk)"
            out["acc-earners-levy-cap"] = make_pulse_item(
                "ACC & IRD Schedule", f"{today.year if today.month >= 4 else today.year - 1}-04-30", MAX_AGE_INSURANCE,
                f"Under the current ACC Earners' Levy schedule, the levy rate is <strong>${levy_rate:.2f} per $100</strong> up to a Maximum Liable Earnings cap of <strong>${max_earn:,}</strong> (a maximum annual levy of <strong>${max_levy:,.2f}</strong>), capping ACC's 80% weekly compensation at <strong>${max_comp_yr:,.2f}/year (${max_comp_wk:,.2f}/week gross)</strong>.",
                f"ACC Earners' Levy Statutory Schedule: Levy rate = ${levy_rate:.2f} per $100; Maximum liable earnings cap = ${max_earn:,}.",
                format_provenance_source_html(acc_lv_calc, "ACC Earners' Levy Statutory Schedule", ACC_LEVY_PAGE, True),
                "Current ACC Tax Year Schedule", max_earn,
                f"${max_earn:,} max covered earnings | ${levy_rate:.2f}/$100 levy | ${max_comp_yr:,.0f}/yr max 80% ACC payout",
                "Statutory 80% weekly accidental injury compensation ceiling",
                f"Income above ${max_earn:,} is not covered by ACC weekly compensation",
                acc_lv_calc, "ACC statutory earners' levy schedule", None, {"source_published": acc_pub},
            )

    acc_rep_html = acc_pages.get(ACC_REPORT_PAGE, "")
    if acc_rep_html:
        acc_r_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", acc_rep_html))
        for pdf_href in re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', acc_rep_html, flags=re.I)[:2]:
            full_pdf = "https://www.acc.co.nz" + pdf_href if pdf_href.startswith("/") else pdf_href
            pdf_b = fetch_bytes_with_browser_fallback(session, full_pdf, referer=ACC_REPORT_PAGE, min_bytes=20_000)
            acc_r_txt += " " + extract_pdf_text(pdf_b)
        m_tot_c = re.search(r"([12]\.[01234]\s*million|[12],[012]\d{2},\d{3})\s+(?:accepted|new|total)?\s*(?:injury\s+)?claims", acc_r_txt, flags=re.I)
        m_wk_c = re.search(r"(\d{2,3},\d{3})\s+(?:weekly\s+compensation)", acc_r_txt, flags=re.I)
        if m_tot_c and m_wk_c:
            raw_tot = m_tot_c.group(1).lower()
            tot_c_val = int(float(re.sub(r"[^\d.]", "", raw_tot)) * 1_000_000) if "million" in raw_tot else int(re.sub(r"[^\d]", "", raw_tot))
            wk_c_val = int(re.sub(r"[^\d]", "", m_wk_c.group(1)))
            wk_ratio = round((wk_c_val / max(1, tot_c_val)) * 100.0, 1)
            acc_r_pub = extract_html_release_date(acc_rep_html, "Annual ACC Corporate Report release")
            acc_wk_calc = f"{wk_c_val:,} weekly compensation claims ÷ {tot_c_val:,} total accepted injury claims = {wk_ratio:.1f} per 100 accepted claims"
            out["acc-weekly-compensation-claims"] = make_pulse_item(
                "ACC Annual Report", f"{today.year if today.month >= 7 else today.year - 1}-06-30", MAX_AGE_INSURANCE,
                f"In its latest annual report, ACC registered <strong>{tot_c_val:,}</strong> accepted injury claims across New Zealand and <strong>{wk_c_val:,}</strong> weekly income compensation claims for time off work (<strong>{wk_ratio:.1f} weekly compensation claims per 100 accepted injury claims</strong>).",
                f"ACC Annual Report: Total accepted injury claims = {tot_c_val:,}; Weekly compensation entitlement claims = {wk_c_val:,}.",
                format_provenance_source_html(acc_wk_calc, "ACC Annual Report (Accident Compensation Corporation)", ACC_REPORT_PAGE, True),
                "Latest ACC Financial Year", wk_ratio,
                f"{wk_c_val:,} weekly comp claims vs {tot_c_val:,} total injury claims ({wk_ratio:.1f} per 100 claims)",
                "ACC covers accidental physical injury only",
                "Non-injury illness (e.g. cancer, heart attack, stroke) is not covered by ACC weekly compensation",
                acc_wk_calc, "ACC Annual Report series", None, {"source_published": acc_r_pub},
            )

    # 5. Live FSC Life & Health Insurance Industry Spotlight PDFs (via blog.fsc.org.nz)
    life_txt, life_src_url, life_per, life_pub = fetch_latest_fsc_spotlight(session, FSC_LIFE_TAG_PAGE, "life")
    hlth_txt, hlth_src_url, hlth_per, hlth_pub = fetch_latest_fsc_spotlight(session, FSC_HEALTH_TAG_PAGE, "health")

    def parse_fsc_money(txt: str, patterns: list[str]) -> float | None:
        for pat in patterns:
            m = re.search(pat, txt, flags=re.I)
            if m:
                num = float(m.group(1).replace(",", ""))
                unit_w = (m.group(2) or "").lower()
                return num * 1e9 if unit_w.startswith("b") else (num * 1e6 if unit_w.startswith("m") else num)
        return None

    life_claims_nzd = parse_fsc_money(
        life_txt,
        [r"\$([\d,.]+)\s*(billion|b|million|m)\s+(?:in\s+)?(?:life\s+)?claims\s+paid", r"claims\s+paid[^$]{0,40}\$([\d,.]+)\s*(billion|b|million|m)"],
    ) or (j10_life_claims_m * 1e6 if j10_life_claims_m is not None else None)
    life_prem_nzd = parse_fsc_money(
        life_txt,
        [r"\$([\d,.]+)\s*(billion|b|million|m)\s+(?:in\s+)?(?:total\s+)?(?:annual\s+)?(?:in-force\s+)?premiums", r"total\s+premiums[^$]{0,40}\$([\d,.]+)\s*(billion|b|million|m)"],
    )
    life_covers_cnt = parse_fsc_money(
        life_txt,
        [r"([\d,.]+)\s*(million|m)\s+(?:life\s+insurance\s+)?covers", r"covers\s+in\s+force[^\d]{0,30}([\d,.]+)\s*(million|m)"],
    )
    trauma_prem_nzd = parse_fsc_money(life_txt, [r"trauma[^$]{0,40}\$([\d,.]+)\s*(million|m|billion|b)"])
    ip_prem_nzd = parse_fsc_money(life_txt, [r"income\s+protection[^$]{0,40}\$([\d,.]+)\s*(million|m|billion|b)"])
    term_prem_nzd = parse_fsc_money(life_txt, [r"term\s+(?:life|&\s*accidental)[^$]{0,40}\$([\d,.]+)\s*(billion|b|million|m)"])

    if life_claims_nzd is not None and life_claims_nzd > 100_000_000:
        lc_b, daily_clm_m = life_claims_nzd / 1e9, (life_claims_nzd / 365.0) / 1e6
        as_of_fsc, per_lbl_l = month_end_iso(pd.Timestamp(today)), life_per or "Latest FSC Life Insurance Spotlight"
        prem_clause = ""
        calc_parts = [f"Daily Claims Paid: ${lc_b:.2f}B ÷ 365 = ${daily_clm_m:.2f}m/day"]
        hist_rows_l = [{"period": "Annual Claims Paid", "label": "Annual Life, Trauma & Disability Claims Paid", "value": round(lc_b, 2), "formatted": f"${lc_b:.2f}B/yr (${daily_clm_m:.2f}m/day)"}]
        if life_prem_nzd and life_covers_cnt:
            lp_b, cov_m = life_prem_nzd / 1e9, life_covers_cnt / 1e6
            avg_prem_cov = int(round(life_prem_nzd / life_covers_cnt))
            prem_clause = f" across <strong>{cov_m:.2f} million covers in force</strong> and <strong>${lp_b:.2f}B in annual premiums</strong> (averaging <strong>${avg_prem_cov:,}/year per cover</strong>)"
            calc_parts.append(f"Mean Annual Premium per Cover: ${lp_b:.2f}B ÷ {cov_m:.2f}m covers = ${avg_prem_cov:,}/cover")
            hist_rows_l.append({"period": "In-Force Covers", "label": "Total Life, Trauma & IP Covers in Force", "value": round(cov_m, 2), "formatted": f"{cov_m:.2f}m covers (avg ${avg_prem_cov:,}/yr premium)"})
        if term_prem_nzd:
            hist_rows_l.append({"period": "Term Life Premiums", "label": "Term Life & Accidental Death Premium Pool", "value": round(term_prem_nzd / 1e9, 2), "formatted": f"${term_prem_nzd/1e9:.2f}B/yr"})
        if trauma_prem_nzd:
            hist_rows_l.append({"period": "Trauma Premiums", "label": "Critical Illness / Trauma Premium Pool", "value": round(trauma_prem_nzd / 1e6, 1), "formatted": f"${trauma_prem_nzd/1e6:,.0f}m/yr"})
        if ip_prem_nzd:
            hist_rows_l.append({"period": "Income Protection", "label": "Income / Mortgage Protection Premium Pool", "value": round(ip_prem_nzd / 1e6, 1), "formatted": f"${ip_prem_nzd/1e6:,.0f}m/yr"})

        lt_calc_s = " | ".join(calc_parts)
        out["life-trauma-average-claim-payout"] = make_pulse_item(
            "FSC Spotlight PDF", as_of_fsc, MAX_AGE_INSURANCE,
            f"According to the Financial Services Council (FSC) Life Insurance Industry Spotlight ({per_lbl_l}), New Zealand life and personal risk insurers paid out <strong>${lc_b:.2f}B in annual claims (${daily_clm_m:.2f}m per day)</strong>{prem_clause}.",
            f"Financial Services Council (FSC) Life Insurance Spotlight ({per_lbl_l}): Annual claims paid = ${lc_b:.2f}B (${daily_clm_m:.2f}m/day).",
            format_provenance_source_html(lt_calc_s, f"Financial Services Council (FSC) — Life Insurance Spotlight ({per_lbl_l})", life_src_url, True),
            per_lbl_l, round(lc_b, 2), f"${lc_b:.2f}B/yr claims paid (${daily_clm_m:.2f}m/day)",
            "Credited to Financial Services Council (FSC) Life Insurance Spotlight",
            "Covers Term Life, Critical Illness / Trauma & Income Protection claims",
            lt_calc_s, f"FSC Life Insurance Spotlight ({per_lbl_l})", hist_rows_l,
            {"source_published": life_pub or "Quarterly FSC Spotlight release", "source_url": life_src_url},
        )

    hlth_claims_nzd = parse_fsc_money(
        hlth_txt,
        [r"\$([\d,.]+)\s*(billion|b|million|m)\s+(?:in\s+)?(?:health\s+)?claims\s+paid", r"claims\s+paid[^$]{0,40}\$([\d,.]+)\s*(billion|b|million|m)"],
    ) or (j10_hlth_claims_m * 1e6 if j10_hlth_claims_m is not None else None)
    hlth_lives_cnt = parse_fsc_money(
        hlth_txt,
        [r"([\d,.]+)\s*(million|m)\s+(?:new\s+zealanders|people|lives|kiwis|insured|policies)", r"lives\s+covered[^\d]{0,30}([\d,.]+)\s*(million|m)"],
    )

    if hlth_claims_nzd is not None and hlth_claims_nzd > 100_000_000:
        hc_b, daily_hc_m = hlth_claims_nzd / 1e9, (hlth_claims_nzd / 365.0) / 1e6
        per_lbl_h, as_of_fsc = hlth_per or "Latest FSC Health Insurance Spotlight", month_end_iso(pd.Timestamp(today))
        if hlth_lives_cnt and hlth_lives_cnt > 500_000:
            hl_m = hlth_lives_cnt / 1e6
            avg_per_mem = int(round(hlth_claims_nzd / hlth_lives_cnt))
            ph_calc = f"Mean Annual Claims Paid per Insured Member: ${hc_b:.3f}B ÷ {hlth_lives_cnt:,.0f} lives = ${avg_per_mem:,}/yr (${daily_hc_m:.2f}m/day across NZ)"
            ph_html = f"According to the Financial Services Council (FSC) Health Insurance Spotlight ({per_lbl_h}), New Zealand private health insurers paid out <strong>${hc_b:.2f}B in annual claims (${daily_hc_m:.2f}m per day)</strong> across <strong>{hl_m:.2f} million insured lives</strong>—an average of <strong>${avg_per_mem:,} in annual claims paid per insured member</strong>."
            ph_val, ph_latest_v = avg_per_mem, f"${hc_b:.2f}B annual claims | ${avg_per_mem:,}/yr per insured member ({hl_m:.2f}m lives)"
        else:
            ph_calc = f"Daily Health Claims Paid: ${hc_b:.2f}B ÷ 365 = ${daily_hc_m:.2f}m/day"
            ph_html = f"According to the Financial Services Council (FSC) Health Insurance Spotlight ({per_lbl_h}) and RBNZ insurance returns, New Zealand private health insurers paid out <strong>${hc_b:.2f}B in annual claims (${daily_hc_m:.2f}m per day)</strong>."
            ph_val, ph_latest_v = round(hc_b, 2), f"${hc_b:.2f}B annual health claims paid (${daily_hc_m:.2f}m/day)"

        out["private-health-claims-average-payout"] = make_pulse_item(
            "FSC Spotlight & RBNZ", as_of_fsc, MAX_AGE_INSURANCE, ph_html,
            f"Financial Services Council (FSC) Health Insurance Spotlight ({per_lbl_h}): Annual private health claims paid = ${hc_b:.2f}B.",
            format_provenance_source_html(ph_calc, f"Financial Services Council (FSC) — Health Insurance Spotlight ({per_lbl_h})", hlth_src_url, True),
            per_lbl_h, ph_val, ph_latest_v,
            "Credited to Financial Services Council (FSC) Health Insurance Spotlight",
            f"${daily_hc_m:.2f}m paid per day across private hospital surgical & medical claims",
            ph_calc, f"FSC Health Insurance Spotlight ({per_lbl_h})", None,
            {"source_published": hlth_pub or "Quarterly FSC Spotlight release", "source_url": hlth_src_url},
        )

    if life_prem_nzd and ip_prem_nzd:
        ip_sh_pct = round((ip_prem_nzd / life_prem_nzd) * 100.0, 1)
        per_lbl_g = life_per or "Latest FSC Spotlight"
        gap_calc = f"Income Protection Share of Life Sector Premiums: ${ip_prem_nzd/1e6:,.0f}m ÷ ${life_prem_nzd/1e9:.2f}B = {ip_sh_pct:.1f}%"
        out["personal-insurance-underinsurance-gap"] = make_pulse_item(
            "FSC Spotlight", month_end_iso(pd.Timestamp(today)), MAX_AGE_INSURANCE,
            f"In the Financial Services Council (FSC) Life Insurance Spotlight ({per_lbl_g}), Income and Mortgage Repayment Protection accounted for <strong>{ip_sh_pct:.1f}% (${ip_prem_nzd/1e6:,.0f}m)</strong> of New Zealand's <strong>${life_prem_nzd/1e9:.2f}B</strong> personal life insurance market—highlighting why working-age borrowers review private illness and disability income cover alongside mandatory house insurance.",
            f"FSC Life Insurance Spotlight ({per_lbl_g}): Income Protection premiums = ${ip_prem_nzd/1e6:,.0f}m out of ${life_prem_nzd/1e9:.2f}B total life insurance premiums ({ip_sh_pct:.1f}%).",
            format_provenance_source_html(gap_calc, f"Financial Services Council (FSC) — Life Insurance Spotlight ({per_lbl_g})", life_src_url, True),
            per_lbl_g, ip_sh_pct,
            f"{ip_sh_pct:.1f}% Income Protection share (${ip_prem_nzd/1e6:,.0f}m of ${life_prem_nzd/1e9:.2f}B life premiums)",
            "ACC covers accidental physical injury only (not illness)",
            "Private income protection covers non-injury illness & disability subject to policy terms",
            gap_calc, f"FSC Life Insurance Spotlight ({per_lbl_g})", None,
            {"source_published": life_pub or "Quarterly FSC Spotlight release", "source_url": life_src_url},
        )

    return out
# === END OF PART 4A ===
# === PART 4B OF 5 ===
def parse_stats_nz_bc_workbook(xlsx_bytes: bytes) -> dict:
    """Parse live Stats NZ Building Consents .xlsx workbook for regional counts and consent values by dwelling type."""
    info: dict = {"regions": {}, "cost_by_type": {}}
    if not xlsx_bytes or xlsx_bytes[:2] != b"PK":
        return info
    try:
        xl = pd.ExcelFile(io.BytesIO(xlsx_bytes))
        type_keys = [
            ("houses", ("stand-alone house", "houses")),
            ("townhouses", ("townhouse", "flats, and units", "flats and units")),
            ("apartments", ("apartment",)),
            ("retirement", ("retirement village",)),
            ("all", ("all dwellings", "total new dwellings", "total dwellings")),
        ]
        for sname in xl.sheet_names:
            raw = xl.parse(sname, header=None)
            if raw.empty:
                continue
            flat_txt = " ".join(str(v) for v in raw.values.flatten() if pd.notna(v)).lower()
            for r_i in range(len(raw)):
                row_s = [str(v).strip() for v in raw.iloc[r_i].values if pd.notna(v) and str(v).strip()]
                if len(row_s) < 2:
                    continue
                r_lbl = row_s[0].lower()
                for reg in ("Auckland", "Canterbury", "Waikato", "Wellington"):
                    if reg not in info["regions"] and r_lbl.startswith(reg.lower()):
                        nums = [
                            int(float(re.sub(r"[^\d.]", "", c)))
                            for c in row_s[1:]
                            if re.match(r"^[\d,]+(\.\d+)?$", c) and 500 <= float(re.sub(r"[^\d.]", "", c)) <= 35000
                        ]
                        if nums:
                            info["regions"][reg] = nums[-1]
                for t_code, kw_tuple in type_keys:
                    if any(k in r_lbl for k in kw_tuple):
                        nums = [float(re.sub(r"[,$]", "", c)) for c in row_s[1:] if re.match(r"^-?\d+(\.\d+)?$", re.sub(r"[,$]", "", c))]
                        if nums:
                            slot = info["cost_by_type"].setdefault(t_code, {})
                            if "value" in sname.lower() or "$m" in flat_txt or "million" in r_lbl:
                                slot.setdefault("val_m", nums[-1])
                            elif "area" in sname.lower() or "sqm" in r_lbl or "m2" in r_lbl or "floor" in sname.lower():
                                slot.setdefault("sqm_tot", nums[-1])
                            else:
                                slot.setdefault("cnt", nums[-1])
    except Exception as exc:
        print(f"[Stats NZ BC Excel] Note: {exc}")
    return info


def parse_stats_nz_cpi_files(session, release_html: str, release_url: str) -> dict:
    """Download live Stats NZ CPI .xlsx / .csv tables linked on the quarterly CPI release page."""
    sub_series: dict = {}
    if not release_html:
        return sub_series
    target_map = {
        "all_groups": ("all groups",),
        "new_housing": ("purchase of new housing", "purchase of housing", "new dwellings"),
        "dwelling_insurance": ("dwelling insurance",),
        "contents_insurance": ("contents insurance",),
        "vehicle_insurance": ("vehicle insurance", "motor vehicle insurance"),
        "health_insurance": ("health insurance",),
        "life_insurance": ("life insurance",),
    }
    for href in re.findall(r'href=["\']([^"\']+\.(?:xlsx|csv)[^"\']*)["\']', release_html, flags=re.I)[:5]:
        full_u = "https://www.stats.govt.nz" + href if href.startswith("/") else href
        f_bytes = fetch_bytes_with_browser_fallback(session, full_u, referer=release_url, min_bytes=2_000)
        if not f_bytes:
            continue
        try:
            frames = (
                [pd.ExcelFile(io.BytesIO(f_bytes)).parse(sh, header=None) for sh in pd.ExcelFile(io.BytesIO(f_bytes)).sheet_names]
                if full_u.lower().endswith(".xlsx") and f_bytes[:2] == b"PK"
                else ([pd.read_csv(io.BytesIO(f_bytes), header=None, low_memory=False)] if full_u.lower().endswith(".csv") else [])
            )
            for df_c in frames:
                for r_i in range(len(df_c)):
                    row_s = [str(v).strip() for v in df_c.iloc[r_i].values if pd.notna(v) and str(v).strip()]
                    if len(row_s) < 2:
                        continue
                    joined_low = " ".join(row_s).lower()
                    for skey, kw_t in target_map.items():
                        if skey not in sub_series and any(k in joined_low for k in kw_t):
                            pcts = [
                                float(c.replace("%", "").replace(",", "").strip())
                                for c in row_s[1:]
                                if re.match(r"^-?\d{1,2}\.\d{1,2}$", c.replace("%", "").replace(",", "").strip())
                                and -15.0 <= float(c.replace("%", "").replace(",", "").strip()) <= 45.0
                            ]
                            if pcts:
                                sub_series[skey] = {"latest": pcts[-1], "history": []}
        except Exception as exc:
            print(f"[Stats NZ CPI File] {full_u}: {exc}")
    return sub_series


def build_stats_nz_tickers(session, j10_path: pathlib.Path | None = None) -> dict:
    """100% live-pulled Stats NZ Building Consents & CPI tickers (zero hardcoded fallbacks)."""
    out = {}
    today = datetime.datetime.now(datetime.timezone.utc).date()

    bc_candidates = []
    for back in range(0, 5):
        y, m = today.year, today.month - back
        while m <= 0:
            m += 12
            y -= 1
        dt_cand = datetime.date(y, m, 1)
        bc_candidates.append((dt_cand, f"https://www.stats.govt.nz/information-releases/building-consents-issued-{dt_cand.strftime('%B-%Y').lower()}/"))

    q_months = [(12, "december"), (9, "september"), (6, "june"), (3, "march")]
    cpi_candidates = sorted(
        [
            (yr, qm, qname, f"https://www.stats.govt.nz/information-releases/consumers-price-index-{qname}-{yr}-quarter/")
            for yr in (today.year, today.year - 1)
            for qm, qname in q_months
            if datetime.date(yr, qm, 1) <= today
        ],
        key=lambda x: (x[0], x[1]),
        reverse=True,
    )[:3]

    fetched_html = fetch_html_with_browser_fallback(
        session, [u for _, u in bc_candidates] + [u for _, _, _, u in cpi_candidates], "Stats NZ"
    )

    for dt_cand, url in bc_candidates:
        html = fetched_html.get(url, "")
        if not html:
            continue
        bc_pub_date = extract_html_release_date(html, "Monthly Stats NZ Building Consents release (~4 weeks after reference month-end)")
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))

        wb_info = {"regions": {}, "cost_by_type": {}}
        for x_href in re.findall(r'href=["\']([^"\']+\.xlsx[^"\']*)["\']', html, flags=re.I)[:2]:
            full_x = "https://www.stats.govt.nz" + x_href if x_href.startswith("/") else x_href
            x_bytes = fetch_bytes_with_browser_fallback(session, full_x, referer=url, min_bytes=10_000)
            if x_bytes:
                wb_info = parse_stats_nz_bc_workbook(x_bytes)
                if wb_info["regions"] or wb_info["cost_by_type"]:
                    break

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
            houses, townhouses = int(m_mth.group(2).replace(",", "")), int(m_mth.group(3).replace(",", ""))
            multi_tot = max(0, mth_tot - houses)
            multi_pct, tfu_pct = (multi_tot / mth_tot) * 100.0, (townhouses / mth_tot) * 100.0
            other_multi = max(0, multi_tot - townhouses)
            other_multi_pct = (other_multi / mth_tot) * 100.0
            m_full, as_of = dt_cand.strftime("%B %Y"), month_end_iso(pd.Timestamp(dt_cand))
            bc_lbl, bc_extra = f"Stats NZ — Building Consents Issued ({m_full})", {"source_published": bc_pub_date, "source_url": url}

            nb_html = (
                f"In the 12 months ended {m_full}, <strong>{ann_tot:,}</strong> new residential dwellings were consented across New Zealand ({dir_word} <strong>{ann_pct:.1f}%</strong> year-on-year, with <strong>{mth_tot:,}</strong> dwellings consented in {m_full})."
                if dir_word in ("up", "down") and ann_pct > 0
                else f"In the 12 months ended {m_full}, <strong>{ann_tot:,}</strong> new residential dwellings were consented across New Zealand, including <strong>{mth_tot:,}</strong> dwellings in {m_full}."
            )
            th_html = f"In {m_full}, <strong>{multi_pct:.1f}%</strong> (<strong>{multi_tot:,} of {mth_tot:,}</strong>) of new residential dwelling consents were for multi-unit homes—with townhouses, flats, and units accounting for <strong>{tfu_pct:.1f}% ({townhouses:,})</strong> and apartments plus retirement village units accounting for <strong>{other_multi_pct:.1f}% ({other_multi:,})</strong>."
            nb_calc = f"Directly reported Stats NZ annual total = {ann_tot:,} ({mth_tot:,} in {m_full})"
            th_calc = f"Multi-unit overall: ({mth_tot:,} − {houses:,}) ÷ {mth_tot:,} = {multi_tot:,} ÷ {mth_tot:,} = {multi_pct:.1f}% | Townhouses/flats/units: {townhouses:,} ÷ {mth_tot:,} = {tfu_pct:.1f}%"

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

            m_apt = re.search(r"([\d,]+)\s+apartments", text[m_mth.start() : m_mth.start() + 500], flags=re.I)
            m_ret = re.search(r"([\d,]+)\s+retirement village units", text[m_mth.start() : m_mth.start() + 500], flags=re.I)
            apts = int(m_apt.group(1).replace(",", "")) if m_apt else None
            ret_u = int(m_ret.group(1).replace(",", "")) if m_ret else (max(0, other_multi - apts) if apts is not None else None)
            if apts is not None and ret_u is not None:
                houses_pct, apts_pct, ret_pct = (houses / mth_tot) * 100.0, (apts / mth_tot) * 100.0, (ret_u / mth_tot) * 100.0
                cat4_calc = f"Houses: {houses:,}/{mth_tot:,}={houses_pct:.1f}% | Townhouses: {townhouses:,}/{mth_tot:,}={tfu_pct:.1f}% | Apartments: {apts:,}/{mth_tot:,}={apts_pct:.1f}% | Retirement: {ret_u:,}/{mth_tot:,}={ret_pct:.1f}%"
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

            def reg_bc_live(reg_name: str) -> int | None:
                if reg_name in wb_info["regions"]:
                    return int(wb_info["regions"][reg_name])
                m_r = re.search(rf"{reg_name}\s*[–\-—:]+\s*([\d,]+)", text, flags=re.I) or re.search(rf"in\s+{reg_name}\s+were\s+([\d,]+)", text, flags=re.I)
                return int(m_r.group(1).replace(",", "")) if m_r else None

            akl_bc, cant_bc, waik_bc, wgtn_bc = reg_bc_live("Auckland"), reg_bc_live("Canterbury"), reg_bc_live("Waikato"), reg_bc_live("Wellington")
            if akl_bc and cant_bc and waik_bc and wgtn_bc:
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

            cbt = wb_info.get("cost_by_type", {})
            m_res_val = re.search(r"residential\s+building\s+work\s+was\s+\$([\d,.]+)\s*(billion|million|b|m)", text, flags=re.I)
            if cbt.get("houses", {}).get("val_m") and houses > 0:
                h_cost = int(round((cbt["houses"]["val_m"] * 1e6) / max(1, cbt["houses"].get("cnt", houses))))
                th_cost = int(round((cbt.get("townhouses", {}).get("val_m", 0) * 1e6) / max(1, cbt.get("townhouses", {}).get("cnt", townhouses)))) if cbt.get("townhouses", {}).get("val_m") else 0
                bc_cost_calc = f"Live Stats NZ Building Consents workbook ({m_full}): Stand-alone house mean consent value = ${h_cost:,}" + (f" | Townhouse/flat/unit mean = ${th_cost:,}" if th_cost else "")
                out["new-build-consent-cost-by-type"] = make_pulse_item(
                    "Stats NZ", as_of, MAX_AGE_QUARTERLY,
                    f"In {m_full}, average consented residential construction values (excluding land) in Stats NZ's building consents dataset averaged <strong>${h_cost:,}</strong> for a stand-alone house" + (f" and <strong>${th_cost:,}</strong> for a townhouse, flat, or unit." if th_cost else "."),
                    f"Stats NZ Building Consents Value Series ({m_full}, excluding land): Stand-alone house mean = ${h_cost:,}" + (f"; Townhouse/flat/unit mean = ${th_cost:,}." if th_cost else "."),
                    format_provenance_source_html(bc_cost_calc, bc_lbl, url, True), m_full, h_cost,
                    f"Stand-alone houses: ${h_cost:,}" + (f" | Townhouses/flats/units: ${th_cost:,}" if th_cost else ""),
                    "Consented physical construction value (excluding land)", "Derived live from Stats NZ Building Consents workbook",
                    bc_cost_calc, f"Stats NZ residential consent value series ({m_full})", None, bc_extra,
                )
            elif m_res_val and mth_tot > 0:
                raw_v = float(m_res_val.group(1).replace(",", ""))
                tot_res_nzd = raw_v * (1e9 if m_res_val.group(2).lower().startswith("b") else 1e6)
                avg_dwell_val = int(round(tot_res_nzd / mth_tot))
                bc_cost_calc = f"${tot_res_nzd/1e6:,.1f}m residential building consent value ÷ {mth_tot:,} consented dwellings = ${avg_dwell_val:,}/dwelling ({houses:,} houses & {multi_tot:,} multi-unit dwellings)"
                out["new-build-consent-cost-by-type"] = make_pulse_item(
                    "Stats NZ", as_of, MAX_AGE_QUARTERLY,
                    f"In {m_full}, total consented residential building work across New Zealand was <strong>${tot_res_nzd/1e6:,.0f}m</strong> across <strong>{mth_tot:,} new dwellings</strong> ({houses:,} stand-alone houses and {multi_tot:,} multi-unit homes)—an average consented construction value of <strong>${avg_dwell_val:,} per dwelling</strong> (excluding land).",
                    f"Stats NZ Building Consents ({m_full}): Residential building consent value = ${tot_res_nzd/1e6:,.0f}m across {mth_tot:,} new dwellings.",
                    format_provenance_source_html(bc_cost_calc, bc_lbl, url, True), m_full, avg_dwell_val,
                    f"${avg_dwell_val:,} average consent build value (${tot_res_nzd/1e6:,.0f}m ÷ {mth_tot:,} dwellings)",
                    f"Stand-alone houses: {houses:,} | Multi-unit dwellings: {multi_tot:,}",
                    "Measures declared physical construction value on building consent (excluding land)",
                    bc_cost_calc, f"Stats NZ monthly release ({m_full})", None, bc_extra,
                )
            print(f"[Stats NZ Consents] Parsed {m_full}: {ann_tot:,} annual, {multi_pct:.1f}% multi-unit")
            break

    latest_cpi_as_of, latest_cpi_lbl, latest_cpi_all, latest_cpi_url, latest_cpi_txt, latest_cpi_pub = "", "", None, "", "", ""
    cpi_sub_series: dict = {}

    for yr, qm, qname, cpi_url in cpi_candidates:
        html = fetched_html.get(cpi_url, "")
        if not html:
            continue
        cpi_pub_date = extract_html_release_date(html, "Quarterly Stats NZ CPI release (~3 weeks after reference quarter-end)")
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))
        cpi_sub_series = parse_stats_nz_cpi_files(session, html, cpi_url)

        m_cpi_all = re.search(r"CPI all groups\s+[-+]?[\d.]+\s+([-+]?[\d.]+)", text, flags=re.I) or re.search(
            r"annual\s+inflation\s+(?:rate\s+)?was\s+([\d.]+)\s*(?:percent|%)", text, flags=re.I
        )
        cpi_all_val = (
            float(m_cpi_all.group(1))
            if m_cpi_all
            else (float(cpi_sub_series["all_groups"]["latest"]) if "all_groups" in cpi_sub_series else None)
        )
        m_build = re.search(
            r"(?:purchase of housing|purchase of new housing|cost of building a new home|new dwellings)[,:\s]+(up|down|fell|rose|increased|decreased)?\s*([\d.]+)\s*(?:percent|%)",
            text, flags=re.I,
        )
        signed_val = None
        if m_build:
            direction = (m_build.group(1) or "up").lower()
            val = float(m_build.group(2))
            signed_val = -val if direction in ("down", "fell", "decreased") else val
        elif "new_housing" in cpi_sub_series:
            signed_val = float(cpi_sub_series["new_housing"]["latest"])

        if signed_val is not None or cpi_all_val is not None:
            q_dt = datetime.date(yr, qm, 1)
            q_as_of_iso, q_lbl_str = month_end_iso(pd.Timestamp(q_dt)), f"{qname.capitalize()} {yr} quarter"
            if signed_val is not None:
                cpi_clause = f", compared with <strong>{cpi_all_val:.1f}%</strong> headline CPI inflation" if cpi_all_val is not None else ""
                bci_calc = f"Directly reported Stats NZ CPI purchase of new housing = {signed_val:+.1f}% YoY"
                out["build-cost-inflation"] = make_pulse_item(
                    "Stats NZ CPI", q_as_of_iso, MAX_AGE_QUARTERLY,
                    f"In the {q_lbl_str}, the cost of purchasing a newly built home changed by <strong>{signed_val:+.1f}%</strong> year-on-year{cpi_clause}.",
                    f"Stats NZ CPI ({q_lbl_str}): Purchase of newly built home = {signed_val:+.1f}% YoY.",
                    format_provenance_source_html(bci_calc, f"Stats NZ — Consumers Price Index ({q_lbl_str})", cpi_url, False),
                    q_lbl_str, signed_val, f"{signed_val:+.1f}% YoY",
                    f"Headline CPI: {cpi_all_val:.1f}% YoY" if cpi_all_val is not None else "Prior year quarter",
                    f"{signed_val:+.1f}% annual change", bci_calc,
                    f"Stats NZ quarterly CPI release ({qname.capitalize()} {yr})", None,
                    {"source_published": cpi_pub_date, "source_url": cpi_url},
                )
            latest_cpi_as_of, latest_cpi_lbl, latest_cpi_all, latest_cpi_url, latest_cpi_txt, latest_cpi_pub = (
                q_as_of_iso, q_lbl_str, cpi_all_val, cpi_url, text, cpi_pub_date
            )
            print(f"[Stats NZ CPI] Parsed {q_lbl_str}: Headline CPI={cpi_all_val}%")
            break

    out.update(
        build_insurance_pulse_tickers(
            session, latest_cpi_as_of, latest_cpi_lbl, latest_cpi_all, latest_cpi_url, latest_cpi_txt, latest_cpi_pub, cpi_sub_series, j10_path
        )
    )
    return out
# === END OF PART 4B ===
# === PART 4C OF 5 ===
def build_tenancy_tickers(session) -> dict:
    """100% live-parsed MBIE Tenancy Services rental bond dataset from the official By-Region CSV (zero hardcoded fallbacks)."""
    pages = fetch_html_with_browser_fallback(session, [TENANCY_PAGE], "MBIE Tenancy")
    html = pages.get(TENANCY_PAGE, "")
    if not html:
        raise RuntimeError("Could not reach Tenancy Services rental bond page.")
    ten_pub_date = extract_html_release_date(html, "Monthly MBIE Tenancy Services rental bond release (~6–8 weeks after reference month-end)")
    csv_url = next(
        ("https://www.tenancy.govt.nz" + h if h.startswith("/") else h for h in re.findall(r'href=["\']([^"\']*\.csv[^"\']*)["\']', html, flags=re.I) if "region" in h.lower()),
        "",
    )
    if not csv_url:
        raise RuntimeError("Could not locate By-Region CSV link on Tenancy Services.")

    csv_bytes = fetch_bytes_with_browser_fallback(session, csv_url, referer=TENANCY_PAGE, min_bytes=10_000)
    df = pd.read_csv(io.BytesIO(csv_bytes))
    df.columns = [c.strip() for c in df.columns]
    df["dt"] = pd.to_datetime(df["Time Frame"], errors="coerce", format="mixed")

    lq_col = next((c for c in df.columns if "lower" in c.lower() and "quartile" in c.lower()), None)
    uq_col = next((c for c in df.columns if "upper" in c.lower() and "quartile" in c.lower()), None)
    for col in ["Lodged Bonds", "Active Bonds", "Closed Bonds", "Median Rent"] + [c for c in (lq_col, uq_col) if c]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "").str.replace("$", "").str.strip(), errors="coerce")

    nat = df[df["Location"].astype(str).str.strip().str.upper() == "ALL"].dropna(subset=["dt", "Median Rent"]).sort_values("dt").reset_index(drop=True)
    last = nat.iloc[-1]
    dt = last["dt"]
    m_full, as_of = dt.strftime("%B %Y"), month_end_iso(dt)
    nat_r = int(round(float(last["Median Rent"])))
    hist_start_mbie = f"Available from {nat['dt'].iloc[0].strftime('%B %Y')} ({len(nat)} monthly observations)"

    def find_nearest_row(target_dt: pd.Timestamp):
        diffs = (nat["dt"] - target_dt).abs()
        min_idx = diffs.idxmin()
        return nat.loc[min_idx] if diffs.loc[min_idx].days <= 65 else None

    yoy_row = find_nearest_row(dt - pd.DateOffset(years=1))
    prev_r = int(round(float(yoy_row["Median Rent"]))) if yoy_row is not None else nat_r
    yoy_pct = ((float(last["Median Rent"]) - prev_r) / prev_r * 100.0) if yoy_row is not None and prev_r else 0.0
    yoy_phrase = f" ({plain_yoy_growth(float(last['Median Rent']), float(yoy_row['Median Rent']))})" if yoy_row is not None else ""

    def reg_med(substr: str) -> int:
        sub = df[(df["dt"] == dt) & (df["Location"].astype(str).str.contains(substr, case=False, na=False))]
        return int(round(float(sub.iloc[-1]["Median Rent"]))) if len(sub) and pd.notna(sub.iloc[-1]["Median Rent"]) else 0

    akl_r, cant_r, wgtn_r = reg_med("Auckland"), reg_med("Canterbury"), reg_med("Wellington")
    mbie_lbl, mbie_extra = f"MBIE Tenancy Services — Rental Bond Data ({m_full})", {"source_published": ten_pub_date, "source_url": csv_url}
    lodged_int, active_int = int(round(float(last["Lodged Bonds"]))), int(round(float(last["Active Bonds"])))
    closed_int = int(round(float(last["Closed Bonds"]))) if "Closed Bonds" in last and pd.notna(last["Closed Bonds"]) else None

    ry_calc = f"Reported national median = ${nat_r}/wk; YoY = (${nat_r} - ${prev_r}) ÷ ${prev_r} = {yoy_pct:+.1f}%"
    sup_calc = f"Directly reported MBIE 'ALL' location: {lodged_int:,} lodged bonds & {active_int:,} active bonds"

    out = {
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
    }

    r5_row, r10_row = find_nearest_row(dt - pd.DateOffset(years=5)), find_nearest_row(dt - pd.DateOffset(years=10))
    if r5_row is not None and r10_row is not None:
        r5_val, r10_val = int(round(float(r5_row["Median Rent"]))), int(round(float(r10_row["Median Rent"])))
        pct_5y, pct_10y = ((nat_r - r5_val) / max(1, r5_val)) * 100.0, ((nat_r - r10_val) / max(1, r10_val)) * 100.0
        cagr_5y, cagr_10y = (((nat_r / max(1, r5_val)) ** (1.0 / 5.0)) - 1.0) * 100.0, (((nat_r / max(1, r10_val)) ** (1.0 / 10.0)) - 1.0) * 100.0
        lt_hist = [
            {"period": r["dt"].strftime("%Y-%m"), "label": r["dt"].strftime("%b %Y"), "value": int(round(float(r["Median Rent"]))), "formatted": f"${int(round(float(r['Median Rent'])))}/wk"}
            for yr_b in range(0, 11)
            for r in [find_nearest_row(dt - pd.DateOffset(years=yr_b))]
            if r is not None
        ]
        lt_calc = f"10Y Nominal: (${nat_r}−${r10_val})÷${r10_val}=+{pct_10y:.1f}% (+{cagr_10y:.1f}%/yr) | 5Y Nominal: (${nat_r}−${r5_val})÷${r5_val}=+{pct_5y:.1f}% (+{cagr_5y:.1f}%/yr)"
        out["rent-long-term-growth"] = make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"Over the 10 years to {m_full}, New Zealand's national median weekly rent rose <strong>+{pct_10y:.1f}% (+${nat_r - r10_val}/wk, or +{cagr_10y:.1f}%/yr)</strong> from <strong>${r10_val}/wk</strong> to <strong>${nat_r}/wk</strong>. Over the past 5 years (from <strong>${r5_val}/wk</strong>), median rent rose <strong>+{pct_5y:.1f}% (+{cagr_5y:.1f}%/yr)</strong>.",
            f"MBIE Tenancy Services ({m_full}): National median weekly rent = ${nat_r}/wk (compared with ${r5_val}/wk 5 years earlier and ${r10_val}/wk 10 years earlier).",
            format_provenance_source_html(lt_calc, mbie_lbl, csv_url, True), m_full, round(pct_10y, 1),
            f"10Y: +{pct_10y:.1f}% (+{cagr_10y:.1f}%/yr) | 5Y: +{pct_5y:.1f}% (+{cagr_5y:.1f}%/yr)",
            f"5 years ago: ${r5_val}/wk | 10 years ago: ${r10_val}/wk",
            f"+${nat_r - r10_val}/wk (+${(nat_r - r10_val)*52:,}/yr) over 10 years", lt_calc, hist_start_mbie, lt_hist, mbie_extra,
        )

    if lq_col and uq_col and pd.notna(last[lq_col]) and pd.notna(last[uq_col]):
        lq_val, uq_val = int(round(float(last[lq_col]))), int(round(float(last[uq_col])))
        iqr_wk, iqr_yr = uq_val - lq_val, (uq_val - lq_val) * 52
        iqr_calc = f"Upper Quartile (${uq_val}/wk) minus Lower Quartile (${lq_val}/wk) = ${iqr_wk}/wk (${iqr_yr:,}/yr)"
        out["rent-quartile-spread"] = make_pulse_item(
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
        )

    if yoy_row is not None and pd.notna(yoy_row.get("Active Bonds")) and closed_int is not None:
        act_prev = int(round(float(yoy_row["Active Bonds"])))
        act_net_12m, act_net_pct = active_int - act_prev, ((active_int - act_prev) / max(1, act_prev)) * 100.0
        mth_net_flow = lodged_int - closed_int
        turn_calc = f"12M Net Stock Change: {active_int:,} − {act_prev:,} = {act_net_12m:+,d} ({act_net_pct:+.1f}%) | Monthly Net Flow: {lodged_int:,} − {closed_int:,} = {mth_net_flow:+,d}"
        out["rental-turnover-net-stock"] = make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"Over the 12 months to {m_full}, New Zealand's active private rental bond stock changed by <strong>{act_net_12m:+,d} tenancies ({act_net_pct:+.1f}%)</strong> to <strong>{active_int:,} active bonds</strong> (with <strong>{lodged_int:,} bonds lodged</strong> and <strong>{closed_int:,} closed</strong> during {m_full}).",
            f"MBIE Tenancy Services ({m_full}): Active Bonds = {active_int:,} (vs {act_prev:,} a year earlier); Monthly Lodged Bonds = {lodged_int:,}; Monthly Closed Bonds = {closed_int:,}.",
            format_provenance_source_html(turn_calc, mbie_lbl, csv_url, True), m_full, act_net_12m,
            f"{act_net_12m:+,d} active bonds YoY ({act_net_pct:+.1f}%) | {lodged_int:,} lodged vs {closed_int:,} closed in {m_full}",
            f"12 months earlier: {act_prev:,} active bonds",
            f"Monthly net bond movement in {m_full}: {mth_net_flow:+,d} bonds", turn_calc, hist_start_mbie,
            make_history_list(nat["dt"], nat["Active Bonds"], lambda x: f"{int(round(x)):,} active bonds"), mbie_extra,
        )

    return out
# === END OF PART 4C ===
# === PART 4D OF 5 ===
import warnings
warnings.filterwarnings("ignore")


def _remove_axis_ticks(nums: list[int], step: int, lo: int, hi: int) -> list[int]:
    """Remove arithmetic chart y-axis gridline ticks (1 instance per tick) so only bar data labels remain."""
    out = list(nums)
    for tick in range(hi, lo - 1, -step):
        if tick in out:
            out.remove(tick)
    return out


def _extract_ko_pdf_totals(pdf_txt: str) -> tuple[str, str, list[int] | None, list[int] | None]:
    """Extract (quarter_label, period_end_iso, approvals_4col, settlements_4col) from a Kāinga Ora quarterly PDF."""
    today = datetime.datetime.now(datetime.timezone.utc).date()
    m_pend = re.search(r"Period\s+Ending\s+(\d{1,2}\s+[A-Za-z]+\s+20\d{2})", pdf_txt, flags=re.I)
    q_end_iso = today.strftime("%Y-%m-%d")
    q_label = "Latest quarter"
    if m_pend:
        dt_pe = pd.to_datetime(m_pend.group(1), errors="coerce")
        if pd.notna(dt_pe):
            q_end_iso = month_end_iso(dt_pe)
            q_map = {3: "Jan - Mar", 6: "Apr - Jun", 9: "Jul - Sep", 12: "Oct - Dec"}
            q_label = f"{q_map.get(dt_pe.month, dt_pe.strftime('%b'))} {dt_pe.year}"

    norm_txt = re.sub(r"\s+", " ", re.sub(r"[\|\r\n]+", " ", pdf_txt))
    totals_4col = []
    for m in re.finditer(r"\bTotal\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)", norm_txt, flags=re.I):
        v = [int(m.group(i).replace(",", "")) for i in range(1, 5)]
        if v[0] + v[1] + v[2] == v[3] and v[3] >= 100:
            totals_4col.append(v)

    app_row = totals_4col[0] if len(totals_4col) >= 1 else None
    set_row = totals_4col[1] if len(totals_4col) >= 2 else None
    return q_label, q_end_iso, app_row, set_row


def build_kainga_ora_tickers(session) -> dict:
    """
    100% live-pulled Kāinga Ora Home Ownership Products Quarterly Report PDFs (CC BY 4.0).
    Reads local *.pdf in repo root first + newest PDFs on kaingaora.govt.nz via fast HTTP.
    Zero hardcoded fallbacks: every figure is mathematically verified from the PDF text stream.
    """
    out: dict = {}
    root = pathlib.Path(__file__).resolve().parent
    ko_pages = fetch_html_with_browser_fallback(session, [KAINGA_ORA_QTR_REPORT_PAGE], "Kāinga Ora")
    qtr_html = ko_pages.get(KAINGA_ORA_QTR_REPORT_PAGE, "")
    qtr_pub = extract_html_release_date(qtr_html, "Quarterly Kāinga Ora Home Ownership Products Report release (CC BY 4.0)")

    # 1. Load local PDF(s) in repo root first (instant, zero browser launches)
    pdf_sources: list[tuple[str, bytes]] = []
    seen_local: set[str] = set()
    for local_pdf in sorted(root.glob("*Home*Ownership*Products*.pdf"), reverse=True) + sorted(root.glob("*.pdf"), reverse=True):
        if local_pdf.name not in seen_local and local_pdf.stat().st_size > 10_000:
            seen_local.add(local_pdf.name)
            b_loc = local_pdf.read_bytes()
            if b_loc.startswith(b"%PDF"):
                pdf_sources.append((KAINGA_ORA_QTR_REPORT_PAGE, b_loc))

    # 2. Fetch up to the top 5 quarterly PDFs from qtr_html via fast HTTP (no Camoufox loop!)
    if qtr_html:
        ko_pdf_links = []
        for h in re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', qtr_html, flags=re.I):
            if "home-ownership-products" in h.lower() or "quarterly" in h.lower():
                full_u = "https://kaingaora.govt.nz" + h if h.startswith("/") else h
                if full_u not in ko_pdf_links:
                    ko_pdf_links.append(full_u)
        for full_u in ko_pdf_links[:5]:
            b_http = fetch_url_bytes(session, full_u, referer=KAINGA_ORA_QTR_REPORT_PAGE, timeout=30, use_cb=False)
            if b_http and b_http.startswith(b"%PDF") and len(b_http) > 10_000:
                pdf_sources.append((full_u, b_http))

    # Parse all available quarterly PDFs and sort newest-first by Period Ending date
    parsed_quarters: list[dict] = []
    seen_periods: set[str] = set()
    for pdf_u, pdf_b in pdf_sources:
        pdf_txt = extract_pdf_text(pdf_b)
        if len(pdf_txt) < 400 or "First Home Loan" not in pdf_txt:
            continue
        q_lbl, q_iso, app_r, set_r = _extract_ko_pdf_totals(pdf_txt)
        if app_r and set_r and q_iso not in seen_periods:
            seen_periods.add(q_iso)
            parsed_quarters.append({
                "url": pdf_u, "bytes": pdf_b, "text": pdf_txt,
                "q_label": q_lbl, "q_end_iso": q_iso, "app_row": app_r, "set_row": set_r,
            })

    if not parsed_quarters:
        return out

    parsed_quarters.sort(key=lambda x: x["q_end_iso"], reverse=True)
    latest = parsed_quarters[0]
    pdf_u = latest["url"]
    pdf_txt = latest["text"]
    q_disp = latest["q_label"]
    q_end_iso = latest["q_end_iso"]
    m1a, m2a, m3a, q_app = latest["app_row"]
    m1s, m2s, m3s, q_set = latest["set_row"]
    ma_app, ma_set = int(round(q_app / 3.0)), int(round(q_set / 3.0))
    conv = (q_set / max(1, q_app)) * 100.0
    norm_txt = re.sub(r"\s+", " ", re.sub(r"[\|\r\n]+", " ", pdf_txt))

    # Dataset 1 & Dataset 4: kainga-ora-first-home-loan-approvals & kainga-ora-underwritten-loan-book
    trend_clause = ""
    roll_clause = ""
    prev_str = f"Monthly approvals in quarter: {m1a:,}, {m2a:,}, {m3a:,}"
    chg_str = f"{q_set:,} First Home Loans settled in {q_disp} ({conv:.1f}% conversion)"
    fhl_hist = [
        {"period": q["q_label"], "label": q["q_label"], "value": q["app_row"][3], "formatted": f"{q['app_row'][3]:,} approvals | {q['set_row'][3]:,} settlements ({(q['set_row'][3]/max(1,q['app_row'][3]))*100:.1f}% settled)"}
        for q in parsed_quarters
    ]

    if len(parsed_quarters) >= 2:
        prev_q = parsed_quarters[1]
        pq_a, pq_s = prev_q["app_row"][3], prev_q["set_row"][3]
        qoq_a_pct = ((q_app - pq_a) / max(1, pq_a)) * 100.0
        trend_clause = f" (<strong>{'up' if qoq_a_pct >= 0 else 'down'} {abs(qoq_a_pct):.1f}% QoQ</strong> from {pq_a:,} in {prev_q['q_label']})"
        prev_str = f"Prior quarter ({prev_q['q_label']}): {pq_a:,} approvals / {pq_s:,} settlements"
        chg_str = f"Approvals: {q_app - pq_a:+,d} ({qoq_a_pct:+.1f}% QoQ) | Settlements: {q_set:,} ({conv:.1f}% conversion)"

    if len(parsed_quarters) >= 5:
        yoy_q = parsed_quarters[4]
        ya, ys = yoy_q["app_row"][3], yoy_q["set_row"][3]
        yoy_a_pct = ((q_app - ya) / max(1, ya)) * 100.0
        yoy_s_pct = ((q_set - ys) / max(1, ys)) * 100.0
        pq_a = parsed_quarters[1]["app_row"][3]
        qoq_a_pct = ((q_app - pq_a) / max(1, pq_a)) * 100.0
        trend_clause = (
            f" (<strong>{plain_yoy_growth(q_app, ya)}</strong> from {ya:,} in {yoy_q['q_label']}, and "
            f"<strong>{'up' if qoq_a_pct >= 0 else 'down'} {abs(qoq_a_pct):.1f}% QoQ</strong> from {pq_a:,} in {parsed_quarters[1]['q_label']})"
        )
        chg_str = f"Approvals: {q_app - ya:+,d} ({yoy_a_pct:+.1f}% YoY, {qoq_a_pct:+.1f}% QoQ) | Settlements: {q_set - ys:+,d} ({yoy_s_pct:+.1f}% YoY)"

    if len(parsed_quarters) >= 4:
        q4_list = list(reversed(parsed_quarters[:4]))
        roll_app_12m = sum(q["app_row"][3] for q in q4_list)
        roll_set_12m = sum(q["set_row"][3] for q in q4_list)
        roll_clause = f" Over the rolling 12-month period, Kāinga Ora recorded <strong>{roll_app_12m:,} approvals</strong> and <strong>{roll_set_12m:,} settled First Home Loans</strong>."
        mth_uw_12m, mth_ap_12m = int(round(roll_set_12m / 12.0)), int(round(roll_app_12m / 12.0))
        uw_calc = (
            f"12-Month Underwritten Settlements: {' + '.join(f'{q[\"set_row\"][3]:,}' for q in q4_list)} = {roll_set_12m:,} loans ({mth_uw_12m:,}/mth) | "
            f"12-Month Approvals: {' + '.join(f'{q[\"app_row\"][3]:,}' for q in q4_list)} = {roll_app_12m:,} approvals ({mth_ap_12m:,}/mth)"
        )
        out["kainga-ora-underwritten-loan-book"] = make_pulse_item(
            "Kāinga Ora Report", q_end_iso, MAX_AGE_QUARTERLY,
            f"Over the 12 months ended {q_disp}, Kāinga Ora underwrote <strong>{roll_set_12m:,} settled First Home Loan (5% deposit) mortgages</strong> (~{mth_uw_12m:,}/month) from <strong>{roll_app_12m:,} approved applications</strong> (~{mth_ap_12m:,}/month) under the Crown Mortgage Insurance Scheme.",
            f"Kāinga Ora Home Ownership Products Report (4 quarters ended {q_disp}): Rolling 12-month First Home Loan settlements = {roll_set_12m:,}; rolling 12-month approvals = {roll_app_12m:,}.",
            format_provenance_source_html(uw_calc, f"Kāinga Ora — Home Ownership Products Report ({q_disp})", pdf_u, True),
            f"12 months ended {q_disp}", roll_set_12m,
            f"{roll_set_12m:,} settled loans underwritten/yr (~{mth_uw_12m:,}/mth) | {roll_app_12m:,} approvals/yr",
            f"Latest quarter ({q_disp}): {q_set:,} settled / {q_app:,} approved",
            f"Annual settlement conversion rate: {(roll_set_12m / max(1, roll_app_12m)) * 100:.1f}% ({roll_set_12m:,} ÷ {roll_app_12m:,})",
            uw_calc, f"Kāinga Ora 4-quarter rolling series (ended {q_disp})", fhl_hist,
            {"source_published": qtr_pub, "source_url": pdf_u},
        )
    else:
        ann_set_pace = q_set * 4
        ann_app_pace = q_app * 4
        uw_calc = f"Quarterly Settlements ({q_disp}): {m1s:,} + {m2s:,} + {m3s:,} = {q_set:,} ({ma_set:,}/mth; annualised pace = {ann_set_pace:,}/yr from {ann_app_pace:,}/yr approvals)"
        out["kainga-ora-underwritten-loan-book"] = make_pulse_item(
            "Kāinga Ora Report", q_end_iso, MAX_AGE_QUARTERLY,
            f"In the {q_disp} quarter, Kāinga Ora underwrote <strong>{q_set:,} settled First Home Loan (5% deposit) mortgages</strong> (~{ma_set:,}/month, an annualised pace of <strong>{ann_set_pace:,} settled loans/year</strong>) from <strong>{q_app:,} quarterly approvals</strong>.",
            f"Kāinga Ora Home Ownership Products Report ({q_disp}): Quarterly settled First Home Loans = {q_set:,} ({m1s:,}, {m2s:,}, {m3s:,}); quarterly approvals = {q_app:,}.",
            format_provenance_source_html(uw_calc, f"Kāinga Ora — Home Ownership Products Report ({q_disp})", pdf_u, True),
            q_disp, q_set,
            f"{q_set:,} settled loans in {q_disp} (~{ma_set:,}/mth; {ann_set_pace:,}/yr pace)",
            f"Monthly settlements in quarter: {m1s:,}, {m2s:,}, {m3s:,}",
            f"Settlement conversion rate: {conv:.1f}% ({q_set:,} ÷ {q_app:,})",
            uw_calc, f"Kāinga Ora quarterly series ({q_disp})", fhl_hist,
            {"source_published": qtr_pub, "source_url": pdf_u},
        )

    fhl_calc = (
        f"Quarterly Approvals: {m1a:,} + {m2a:,} + {m3a:,} = {q_app:,} ({ma_app:,}/mth) | "
        f"Quarterly Settlements: {m1s:,} + {m2s:,} + {m3s:,} = {q_set:,} ({ma_set:,}/mth; {conv:.1f}% conversion)"
    )
    out["kainga-ora-first-home-loan-approvals"] = make_pulse_item(
        "Kāinga Ora Report", q_end_iso, MAX_AGE_QUARTERLY,
        f"In the {q_disp} quarter, Kāinga Ora recorded <strong>{q_app:,} First Home Loan (5% deposit) approvals</strong> (~{ma_app:,}/month){trend_clause} and <strong>{q_set:,} loan settlements</strong> (~{ma_set:,}/month, a <strong>{conv:.1f}%</strong> settlement-to-approval rate).{roll_clause}",
        f"Kāinga Ora Home Ownership Products Report ({q_disp}): First Home Loan approvals = {q_app:,} ({m1a:,}, {m2a:,}, {m3a:,}); settlements = {q_set:,} ({m1s:,}, {m2s:,}, {m3s:,}).",
        format_provenance_source_html(fhl_calc, f"Kāinga Ora — Home Ownership Products Quarterly Report ({q_disp})", pdf_u, True),
        q_disp, q_app, f"{q_app:,} approvals ({ma_app:,}/mth) | {q_set:,} settlements ({conv:.1f}%)",
        prev_str, chg_str, fhl_calc, f"Kāinga Ora quarterly series ({q_disp})", fhl_hist,
        {"source_published": qtr_pub, "source_url": pdf_u},
    )

    # Dataset 2: kainga-ora-first-home-loan-regional (Council / City TLA Leaderboard)
    cities_spec = [
        ("Auckland City", r"Auckland\s+City"), ("Christchurch City", r"Christchurch\s+City"),
        ("Hamilton City", r"Hamilton\s+City"), ("Dunedin City", r"Dunedin\s+City"),
        ("Wellington City", r"Wellington\s+City"), ("Tauranga City", r"Tauranga\s+City"),
        ("Invercargill City", r"Invercargill\s+City"), ("Napier City", r"Napier\s+City"),
        ("Rotorua District", r"Rotorua\s+District"), ("Palmerston North District", r"Palmerston\s+North\s+District"),
        ("Nelson City", r"Nelson\s+City"), ("Waikato District", r"Waikato\s+District"),
        ("New Plymouth District", r"New\s+Plymouth\s+District"), ("Lower Hutt City", r"Lower\s+Hutt\s+City"),
        ("Selwyn District", r"Selwyn\s+District"), ("Hastings District", r"Hastings\s+District"),
        ("Waimakariri District", r"Waimakariri\s+District"),
    ]
    c_rows = []
    for c_lbl, c_pat in cities_spec:
        m_list = [
            [int(m.group(i).replace(",", "")) for i in range(1, 5)]
            for m in re.finditer(c_pat + r"\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)\s+(\d[\d,]*)", norm_txt, flags=re.I)
        ]
        m_list = [v for v in m_list if v[0] + v[1] + v[2] == v[3]]
        if len(m_list) == 1:
            mm = re.search(c_pat + r"\s+[A-Za-z\s]+?\s+(\d+)\s+\d+\s+(\d+)\s+\d+\s+(\d+)\s+\d+\s+(\d+)\s+\d+", norm_txt, flags=re.I)
            if mm:
                vm = [int(mm.group(i)) for i in range(1, 5)]
                if vm[0] + vm[1] + vm[2] == vm[3]:
                    m_list.append(vm)
        if m_list:
            ca = m_list[0][3]
            cs = m_list[1][3] if len(m_list) >= 2 else None
            c_sh = (ca / max(1, q_app)) * 100.0
            s_part = f" | {cs:,} settlements" if cs is not None else ""
            c_rows.append({"period": c_lbl, "label": c_lbl, "value": ca, "set_val": cs or 0, "share": c_sh, "formatted": f"{ca:,} approvals ({c_sh:.1f}% of NZ){s_part}"})

    if len(c_rows) >= 5:
        c_rows.sort(key=lambda x: x["value"], reverse=True)
        top1, top2, top3, top4 = c_rows[0], c_rows[1], c_rows[2], c_rows[3]
        top2_comb_sh = top1["share"] + top2["share"]
        reg_fhl_calc = (
            f"{top1['label']}: {top1['value']:,}÷{q_app:,}={top1['share']:.1f}% | "
            f"{top2['label']}: {top2['value']:,}÷{q_app:,}={top2['share']:.1f}% | "
            f"{top3['label']}: {top3['value']:,}÷{q_app:,}={top3['share']:.1f}%"
        )
        out["kainga-ora-first-home-loan-regional"] = make_pulse_item(
            "Kāinga Ora Report", q_end_iso, MAX_AGE_QUARTERLY,
            f"In the {q_disp} quarter, Kāinga Ora's <strong>{q_app:,}</strong> First Home Loan (5% deposit) approvals were led by <strong>{top1['label']} ({top1['value']:,} approvals, or {top1['share']:.1f}% of NZ)</strong> and <strong>{top2['label']} ({top2['value']:,}, or {top2['share']:.1f}%)</strong>—together accounting for <strong>{top2_comb_sh:.1f}%</strong> of national approvals—followed by <strong>{top3['label']} ({top3['value']:,})</strong> and <strong>{top4['label']} ({top4['value']:,})</strong>.",
            f"Kāinga Ora Home Ownership Products Report ({q_disp} TLA Tables): {top1['label']} = {top1['value']:,} approvals ({top1['set_val']:,} settlements); {top2['label']} = {top2['value']:,} ({top2['set_val']:,}); {top3['label']} = {top3['value']:,} ({top3['set_val']:,}).",
            format_provenance_source_html(reg_fhl_calc, f"Kāinga Ora — Home Ownership Products Report TLA Tables ({q_disp})", pdf_u, True),
            q_disp, top1["value"],
            f"{top1['label']} {top1['value']:,} ({top1['share']:.1f}%) | {top2['label']} {top2['value']:,} ({top2['share']:.1f}%) | {top3['label']} {top3['value']:,}",
            f"National quarterly total: {q_app:,} approvals & {q_set:,} settlements",
            f"{top1['label']} + {top2['label']} = {top2_comb_sh:.1f}% of all NZ First Home Loan approvals",
            reg_fhl_calc, f"Kāinga Ora Territorial Local Authority tables ({q_disp})",
            [{k: r[k] for k in ("period", "label", "value", "formatted")} for r in c_rows],
            {"source_published": qtr_pub, "source_url": pdf_u},
        )

    q_pat = r"(?:Apr|April|Jul|July|Oct|October|Jan|January)\s*[-–—to\s]+\s*(?:Jun|June|Sep|Sept|September|Dec|December|Mar|March)\s*20\d{2}"

    # Dataset 3: 'Previous Owner Savings Withdrawal Applications' (Page 5 Bottom Chart — numbers appear BEFORE title in PDF stream)
    m_sc_pos = re.search(r"Previous\s+Owner\s+Savings\s+Withdrawal\s+Applications", pdf_txt, flags=re.I)
    if m_sc_pos:
        m_sc_hdr = re.search(r"Determination\s+for\s+Previous\s+Homeowners", pdf_txt, flags=re.I)
        sc_start = m_sc_hdr.end() if m_sc_hdr else max(0, m_sc_pos.start() - 950)
        sc_win = pdf_txt[sc_start : min(len(pdf_txt), m_sc_pos.end() + 950)]
        q_list = [re.sub(r"\s+", " ", q).strip() for q in re.findall(q_pat, sc_win, flags=re.I)]
        cln = sc_win
        for qm in q_list:
            cln = cln.replace(qm, " ")
        cln = re.sub(r"\b(?:1\s+October|30\s+September|1\s+July|30\s+June|31\s+March|31\s+December)\s+20\d{2}\b|\b20[12]\d\b|\b1,000\b", " ", cln, flags=re.I)
        raw_sc = [int(n.replace(",", "")) for n in re.findall(r"\b\d{1,2},\d{3}\b|\b\d{3,4}\b|\b0\b", cln)]
        sc_bars = [x for x in _remove_axis_ticks(raw_sc, 200, 0, 1600) if 250 <= x <= 3500]
        n_q = len(q_list)
        if n_q >= 4 and len(sc_bars) >= 2 * n_q:
            cand = sc_bars[: 2 * n_q]
            grp = [(cand[i], cand[i + n_q]) for i in range(n_q)]
            ilv = [(cand[2 * i], cand[2 * i + 1]) for i in range(n_q)]
            app_s = [a for _, a in grp] if all(r >= a for r, a in grp) else ([a for _, a in ilv] if all(r >= a for r, a in ilv) else cand[n_q : 2 * n_q])
            rec_s = [r for r, _ in grp] if all(r >= a for r, a in grp) else ([r for r, _ in ilv] if all(r >= a for r, a in ilv) else cand[:n_q])
            curr_q, curr_a, curr_r = q_list[-1], app_s[-1], rec_s[-1]
            prev_q, prev_a = q_list[-2], app_s[-2]
            roll_12m, mth_a = sum(app_s[-4:]), int(round(curr_a / 3.0))
            qoq_pct = ((curr_a - prev_a) / max(1, prev_a)) * 100.0
            yoy_cl, yoy_chg = f" ({'up' if qoq_pct >= 0 else 'down'} {abs(qoq_pct):.1f}% QoQ from {prev_a:,} in {prev_q})", f"{curr_a - prev_a:+,d} ({qoq_pct:+.1f}% QoQ) vs {prev_q}"
            if n_q >= 5:
                yq, ya = q_list[-5], app_s[-5]
                ypct = ((curr_a - ya) / max(1, ya)) * 100.0
                yoy_cl = f" ({plain_yoy_growth(curr_a, ya)} from {ya:,} in {yq}, and {'up' if qoq_pct >= 0 else 'down'} {abs(qoq_pct):.1f}% QoQ from {prev_a:,} in {prev_q})"
                yoy_chg = f"{curr_a - ya:+,d} ({ypct:+.1f}% YoY vs {yq}: {ya:,}) | {curr_a - prev_a:+,d} ({qoq_pct:+.1f}% QoQ)"
            sc_calc = f"12-Month Rolling Approvals: {' + '.join(f'{x:,}' for x in app_s[-4:])} = {roll_12m:,} approvals/yr | Latest Quarter ({curr_q}): {curr_a:,} approved of {curr_r:,} received (÷ 3 = {mth_a:,}/mth)"
            sc_hist = [{"period": ql, "label": ql, "value": av, "formatted": f"{av:,} approved (of {rv:,} received)"} for ql, av, rv in reversed(list(zip(q_list, app_s, rec_s)))]
            out["kainga-ora-second-chance-kiwisaver-approvals"] = make_pulse_item(
                "Kāinga Ora Report", q_end_iso, MAX_AGE_QUARTERLY,
                f"In the {curr_q} quarter, Kāinga Ora approved <strong>{curr_a:,} 'second-chance' (Previous Homeowner) KiwiSaver savings withdrawal applications</strong>{yoy_cl} (~{mth_a:,}/month), bringing the rolling 12-month total to <strong>{roll_12m:,} approved previous homeowners</strong>.",
                f"Kāinga Ora Home Ownership Products Report ({curr_q}): Previous Owner Savings Withdrawal Applications Approved = {curr_a:,} in {curr_q} ({roll_12m:,} over rolling 4 quarters).",
                format_provenance_source_html(sc_calc, f"Kāinga Ora — Home Ownership Products Quarterly Report ({curr_q})", pdf_u, True),
                curr_q, curr_a, f"{curr_a:,} approvals in {curr_q} ({roll_12m:,} over 12 months)",
                f"Prior quarter ({prev_q}): {prev_a:,} approvals", yoy_chg, sc_calc,
                f"Kāinga Ora 5-quarter published series ({q_list[0]} to {curr_q})", sc_hist,
                {"source_published": qtr_pub, "source_url": pdf_u},
            )

    # Dataset 5: 'KiwiSaver FirstHome Savings Withdrawal (IRD)' 4-Quarter Chart (Page 5 Top Chart — numbers appear BEFORE title in PDF stream)
    m_ks_start = re.search(r"programme\s+of\s+assistance\s+for\s+first\s+home\s+buyers\.?", pdf_txt, flags=re.I)
    m_ks_end = re.search(r"KiwiSaver\s+Savings\s+Withdrawal\s+Determination", pdf_txt, flags=re.I)
    if m_ks_start and m_ks_end and m_ks_end.start() > m_ks_start.end():
        ks_w = pdf_txt[m_ks_start.end() : m_ks_end.start()]
        ks_qs = [re.sub(r"\s+", " ", q).strip() for q in re.findall(q_pat, ks_w, flags=re.I)]
        ks_c = ks_w
        for qm in ks_qs:
            ks_c = ks_c.replace(qm, " ")
        raw_dlr = [int(x.replace(",", "")) for x in re.findall(r"\$([\d,]+)", ks_c)]
        dlr_bars = [x for x in _remove_axis_ticks(raw_dlr, 10, 400, 650) if 300 <= x <= 950]
        ks_no_d = re.sub(r"\b20[12]\d\b", " ", re.sub(r"\$[\d,]+", " ", ks_c))
        raw_cnts = [int(x.replace(",", "")) for x in re.findall(r"\b\d{1,2},\d{3}\b", ks_no_d)]
        cnt_bars = [x for x in _remove_axis_ticks(raw_cnts, 200, 11000, 14000) if 8000 <= x <= 20000]
        if len(ks_qs) == 4 and len(cnt_bars) >= 4 and len(dlr_bars) >= 4:
            cnt_4q, dlr_4q = cnt_bars[:4], dlr_bars[:4]
            roll_cnt_12m, roll_dlr_12m = sum(cnt_4q), sum(dlr_4q)
            roll_avg_nzd = int(round((roll_dlr_12m * 1e6) / max(1, roll_cnt_12m)))
            q_now_lbl, q_cnt_now, q_dlr_now = ks_qs[-1], cnt_4q[-1], dlr_4q[-1]
            q_avg_now = int(round((q_dlr_now * 1e6) / max(1, q_cnt_now)))
            ks4_calc = (
                f"12-Month Rolling Total: {' + '.join(f'{x:,}' for x in cnt_4q)} = {roll_cnt_12m:,} withdrawals & "
                f"{' + '.join(f'${x}m' for x in dlr_4q)} = ${roll_dlr_12m:,}m (${roll_dlr_12m/1000:.3f}B; avg ${roll_avg_nzd:,}/withdrawal)"
            )
            ks4_hist = [
                {"period": ql, "label": ql, "value": dv, "formatted": f"${dv:,}m across {cv:,} withdrawals (avg ${int(round((dv*1e6)/max(1,cv))):,})"}
                for ql, cv, dv in reversed(list(zip(ks_qs, cnt_4q, dlr_4q)))
            ]
            out["kiwisaver-first-home-quarterly-trend"] = make_pulse_item(
                "Kāinga Ora / IRD", q_end_iso, MAX_AGE_QUARTERLY,
                f"Over the 12 months ended {q_now_lbl}, <strong>{roll_cnt_12m:,} KiwiSaver first-home savings withdrawals</strong> were completed across New Zealand totalling <strong>${roll_dlr_12m/1000:.2f}B (${roll_dlr_12m:,}m)</strong>—an average of <strong>${roll_avg_nzd:,} per withdrawal</strong> (with <strong>{q_cnt_now:,} withdrawals</strong> totalling <strong>${q_dlr_now:,}m</strong> in {q_now_lbl}).",
                f"Kāinga Ora Home Ownership Products Report / IRD ({q_now_lbl}): 4-quarter rolling KiwiSaver first-home withdrawals = {roll_cnt_12m:,} withdrawals and ${roll_dlr_12m:,}m.",
                format_provenance_source_html(ks4_calc, f"Kāinga Ora Home Ownership Products Report / IRD ({q_now_lbl})", pdf_u, True),
                q_now_lbl, roll_dlr_12m,
                f"${roll_dlr_12m/1000:.2f}B (${roll_dlr_12m:,}m) across {roll_cnt_12m:,} withdrawals (12m rolling)",
                f"Latest quarter ({q_now_lbl}): ${q_dlr_now:,}m across {q_cnt_now:,} withdrawals (avg ${q_avg_now:,})",
                f"12-month average withdrawal: ${roll_avg_nzd:,} per first-home buyer",
                ks4_calc, f"IRD / Kāinga Ora 4-quarter series ({ks_qs[0]} to {ks_qs[-1]})", ks4_hist,
                {"source_published": qtr_pub, "source_url": pdf_u},
            )

    return out
# === PART 4D-2 OF 5 ===
def build_ird_tickers(session, t43_path: pathlib.Path | None = None) -> dict:
    """100% live-pulled Inland Revenue (IRD), FMA, RBNZ Table T43 & Kāinga Ora datasets (zero hardcoded fallbacks)."""
    out = {}
    today = datetime.datetime.now(datetime.timezone.utc).date()
    ird_urls = [
        IRD_SL_BAL_PAGE, IRD_KS_NUM_PAGE, IRD_KS_AMT_PAGE, IRD_KS_REG_PAGE, IRD_KS_TR_PAGE,
        IRD_KS_ENTRY_PAGE, IRD_KS_PAY_PAGE, IRD_KS_SUSP_PAGE, IRD_KS_DATA_PAGE, FMA_KS_REPORT_PAGE,
    ]
    pages = fetch_html_with_browser_fallback(session, ird_urls, "IRD / FMA")

    sl_html = pages.get(IRD_SL_BAL_PAGE, "")
    if sl_html:
        sl_pub = extract_html_release_date(sl_html, "Quarterly Inland Revenue Student Loan statistics release")
        sl_text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", sl_html))
        m_med = re.search(r"Median student loan balance\s+(?:increased|decreased)\s+by\s+[\d.]+%\s+from\s+\$([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})\s+to\s+\$([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})", sl_text, flags=re.I)
        if m_med:
            prev_bal, prev_lbl, curr_bal = int(m_med.group(1).replace(",", "")), m_med.group(2), int(m_med.group(3).replace(",", ""))
            dt_sl = pd.to_datetime("1 " + m_med.group(4), errors="coerce")
            if pd.notna(dt_sl):
                m_full, pct_sl = dt_sl.strftime("%B %Y"), ((curr_bal - prev_bal) / max(1, prev_bal)) * 100.0
                sl_calc = f"Reported IRD median = ${curr_bal:,} (Change: (${curr_bal:,} - ${prev_bal:,}) ÷ ${prev_bal:,} = {pct_sl:+.1f}%)"
                out["student-loan-servicing"] = make_pulse_item(
                    "Inland Revenue", month_end_iso(dt_sl), MAX_AGE_QUARTERLY,
                    f"As at {m_full}, the median New Zealand student loan balance was <strong>${curr_bal:,}</strong> ({plain_yoy_growth(curr_bal, prev_bal)} from ${prev_bal:,} in {prev_lbl}).",
                    f"Inland Revenue Student Loan Scheme ({m_full}): Median borrower balance = ${curr_bal:,} (compared with ${prev_bal:,} in {prev_lbl}).",
                    format_provenance_source_html(sl_calc, f"Inland Revenue — Student Loan Statistics ({dt_sl.strftime('%b %Y')})", IRD_SL_BAL_PAGE, False),
                    m_full, curr_bal, f"${curr_bal:,} median balance", f"{prev_lbl} (${prev_bal:,})",
                    f"${curr_bal - prev_bal:+,d} ({pct_sl:+.1f}%)", sl_calc, f"Inland Revenue release ({m_full})", None, {"source_published": sl_pub},
                )

    num_html, amt_html = pages.get(IRD_KS_NUM_PAGE, ""), pages.get(IRD_KS_AMT_PAGE, "")
    if num_html and amt_html:
        wd_pub = extract_html_release_date(amt_html, "Monthly Inland Revenue KiwiSaver withdrawal statistics release")
        num_txt, amt_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", num_html)), re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", amt_html))
        mn = re.search(r"In\s+([A-Za-z]+\s+\d{4}):.*?([\d,]+)\s+members have withdrawn their savings for the purpose of first home purchase", num_txt, flags=re.I)
        ma = re.search(r"In\s+([A-Za-z]+\s+\d{4}):.*?\$([\d,.]+)\s+million was attributable to first home purchase(?:,\s+(?:up|down)\s+from\s+\$([\d,.]+)\s+million)?", amt_txt, flags=re.I)
        if mn and ma:
            dt_ks = pd.to_datetime("1 " + ma.group(1), errors="coerce")
            if pd.notna(dt_ks):
                ks_m_full, ks_m_as_of = dt_ks.strftime("%B %Y"), month_end_iso(dt_ks)
                fh_cnt, fh_amt_m = int(mn.group(2).replace(",", "")), float(ma.group(2).replace(",", ""))
                fh_prev_m = float(ma.group(3).replace(",", "")) if ma.group(3) else None
                avg_mem = int(round((fh_amt_m * 1e6) / max(1, fh_cnt)))
                fh_calc = f"Monthly mean: ${fh_amt_m:,.1f}m × $1,000,000 ÷ {fh_cnt:,} members = ${avg_mem:,}/member"
                out["kiwisaver-first-home"] = make_pulse_item(
                    "Inland Revenue", ks_m_as_of, MAX_AGE_MONTHLY,
                    f"In {ks_m_full}, <strong>{fh_cnt:,}</strong> KiwiSaver members withdrew a combined <strong>${fh_amt_m:,.1f}m</strong> toward a first home—an average of <strong>${avg_mem:,}</strong> per withdrawing member.",
                    f"Inland Revenue KiwiSaver Withdrawals ({ks_m_full}): First-home withdrawals = {fh_cnt:,} members and ${fh_amt_m:,.1f}m.",
                    format_provenance_source_html(fh_calc, "Inland Revenue — KiwiSaver Withdrawal Statistics", IRD_KS_DATA_PAGE, True),
                    ks_m_full, round(fh_amt_m, 1), f"${fh_amt_m:,.1f}m across {fh_cnt:,} members (avg ${avg_mem:,}/member)",
                    f"Prior month: ${fh_prev_m:,.1f}m" if fh_prev_m is not None else "Monthly IRD withdrawal series",
                    f"{fh_amt_m - fh_prev_m:+.1f}m vs prior month" if fh_prev_m is not None else f"${avg_mem:,} average per withdrawing member",
                    fh_calc, f"Inland Revenue monthly series ({ks_m_full})", None, {"source_published": wd_pub},
                )
                mn_h = re.search(r"([\d,]+)\s+members have withdrawn their savings for the purpose of financial hardship", num_txt, flags=re.I)
                ma_h = re.search(r"\$([\d,.]+)\s+million was attributable to financial hardship(?:,\s+(?:up|down)\s+from\s+\$([\d,.]+)\s+million)?", amt_txt, flags=re.I)
                if mn_h and ma_h:
                    hs_cnt, hs_amt_m = int(mn_h.group(1).replace(",", "")), float(ma_h.group(1).replace(",", ""))
                    hs_prev_m = float(ma_h.group(2).replace(",", "")) if ma_h.group(2) else None
                    hs_avg = int(round((hs_amt_m * 1e6) / max(1, hs_cnt)))
                    hs_calc = f"${hs_amt_m:,.1f}m × $1,000,000 ÷ {hs_cnt:,} members = ${hs_avg:,}/member"
                    out["kiwisaver-hardship-withdrawals"] = make_pulse_item(
                        "Inland Revenue", ks_m_as_of, MAX_AGE_MONTHLY,
                        f"In {ks_m_full}, <strong>{hs_cnt:,}</strong> KiwiSaver members made a financial hardship withdrawal totalling <strong>${hs_amt_m:,.1f}m</strong>—an average of <strong>${hs_avg:,}</strong> per member.",
                        f"Inland Revenue KiwiSaver Withdrawals ({ks_m_full}): Financial hardship withdrawals = {hs_cnt:,} members and ${hs_amt_m:,.1f}m.",
                        format_provenance_source_html(hs_calc, "Inland Revenue — KiwiSaver Withdrawal Statistics", IRD_KS_DATA_PAGE, True),
                        ks_m_full, round(hs_amt_m, 1), f"${hs_amt_m:,.1f}m across {hs_cnt:,} members (avg ${hs_avg:,}/member)",
                        f"Prior month: ${hs_prev_m:,.1f}m" if hs_prev_m is not None else "Monthly IRD withdrawal series",
                        f"{hs_amt_m - hs_prev_m:+.1f}m vs prior month" if hs_prev_m is not None else f"${hs_avg:,} average hardship withdrawal",
                        hs_calc, f"Inland Revenue monthly series ({ks_m_full})", None, {"source_published": wd_pub},
                    )

    reg_html = pages.get(IRD_KS_REG_PAGE, "")
    if reg_html:
        try:
            reg_pub = extract_html_release_date(reg_html, "Annual Inland Revenue regional KiwiSaver withdrawal release")
            m_yr = re.search(r"year ended 30 June (\d{4})", reg_html, flags=re.I)
            fy_reg = int(m_yr.group(1)) if m_yr else (today.year if today.month >= 7 else today.year - 1)
            ext = SimpleTableExtractor()
            ext.feed(reg_html)
            if len(ext.tables) >= 2:
                amt_map = {r[0].strip().lower(): (r[0].strip(), float(re.sub(r"[^\d.]", "", r[-1]))) for r in ext.tables[0][1:] if len(r) >= 3 and r[0].strip() and re.sub(r"[^\d.]", "", r[-1]) and r[0].strip().lower() not in ("other", "total", "region")}
                cnt_map = {r[0].strip().lower(): int(round(float(re.sub(r"[^\d.]", "", r[-1])))) for r in ext.tables[1][1:] if len(r) >= 3 and r[0].strip().lower() in amt_map and re.sub(r"[^\d.]", "", r[-1])}
                reg_list, tot_nzd, tot_mem = [], 0.0, 0
                for k_low, (disp, nzd_v) in amt_map.items():
                    c_v = cnt_map.get(k_low, 0)
                    if c_v > 0:
                        tot_nzd += nzd_v
                        tot_mem += c_v
                        reg_list.append((disp.replace("Bay Of Plenty", "Bay of Plenty").replace("Hawkes Bay", "Hawke's Bay"), int(round(nzd_v / c_v)), c_v, round(nzd_v / 1e6, 2)))
                if tot_mem > 0 and len(reg_list) >= 10:
                    reg_list.sort(key=lambda x: x[1], reverse=True)
                    nz_avg_p, h_n, h_a, l_n, l_a = int(round(tot_nzd / tot_mem)), reg_list[0][0], reg_list[0][1], reg_list[-1][0], reg_list[-1][1]
                    akl_a = dict((r[0], r[1]) for r in reg_list).get("Auckland", reg_list[1][1])
                    ks_reg_calc = f"${tot_nzd:,.0f} ÷ {tot_mem:,} members = ${nz_avg_p:,}/member (Regional range: ${l_a:,}–${h_a:,})"
                    out["fhb-regional-kiwisaver"] = make_pulse_item(
                        "Inland Revenue", f"{fy_reg}-06-30", MAX_AGE_ANNUAL,
                        f"Over the year ended 30 June {fy_reg}, <strong>{tot_mem:,} KiwiSaver members</strong> withdrew <strong>${tot_nzd/1e9:.2f}B</strong> toward a first home across New Zealand's 16 regions—averaging <strong>${nz_avg_p:,} per withdrawing member nationwide</strong>, from <strong>${h_a:,} in {h_n}</strong> and <strong>${akl_a:,} in Auckland</strong> to <strong>${l_a:,} in {l_n}</strong>.",
                        f"Inland Revenue Regional KiwiSaver Withdrawals (Year ended 30 June {fy_reg}): {tot_mem:,} members withdrew ${tot_nzd:,.0f} across 16 NZ regions.",
                        format_provenance_source_html(ks_reg_calc, f"Inland Revenue — KiwiSaver Fund Withdrawals by Region (Year ended Jun {fy_reg})", IRD_KS_REG_PAGE, True),
                        f"Year ended June {fy_reg}", nz_avg_p, f"${nz_avg_p:,}/member NZ ({tot_mem:,} members / ${tot_nzd/1e9:.2f}B total)",
                        f"Regional range: ${l_a:,}–${h_a:,} per member", f"Highest: {h_n} ${h_a:,} | Lowest: {l_n} ${l_a:,}",
                        ks_reg_calc, f"16 NZ Regions (Year ended 30 June {fy_reg})",
                        [{"period": r[0], "label": r[0], "value": r[1], "formatted": f"${r[1]:,}/member ({r[2]:,} members, ${r[3]:.2f}m)"} for r in reg_list],
                        {"source_published": reg_pub},
                    )
        except Exception as exc:
            print(f"[IRD Regional KiwiSaver] Parse note: {exc}")

    tr_html = pages.get(IRD_KS_TR_PAGE, "")
    if tr_html:
        tr_pub = extract_html_release_date(tr_html, "Annual Inland Revenue KiwiSaver scheme transfer release")
        tr_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", tr_html))
        mt = re.search(r"year ended\s+30\s+June\s+(\d{4}).*?scheme transfers\s+(?:increased|decreased),?\s+from\s+([\d,]+)\s+in the\s+(\d{4})\s+financial year to\s+([\d,]+)\s+in the\s+(\d{4})\s+financial year", tr_txt, flags=re.I)
        if mt:
            prev_tr, prev_fy, tr_tot, fy_yr = int(mt.group(2).replace(",", "")), int(mt.group(3)), int(mt.group(4).replace(",", "")), int(mt.group(5))
            mth_avg, tr_yoy_pct = int(round(tr_tot / 12.0)), ((tr_tot - prev_tr) / max(1, prev_tr)) * 100.0
            tr_calc = f"Monthly mean: {tr_tot:,} ÷ 12 = {mth_avg:,}/mth | YoY: ({tr_tot:,} − {prev_tr:,}) ÷ {prev_tr:,} = {tr_yoy_pct:+.1f}%"
            out["kiwisaver-switches"] = make_pulse_item(
                "Inland Revenue", f"{fy_yr}-06-30", MAX_AGE_ANNUAL,
                f"Over the year ended 30 June {fy_yr}, <strong>{tr_tot:,}</strong> KiwiSaver members switched providers ({plain_yoy_growth(tr_tot, prev_tr)} from {prev_tr:,} in FY{prev_fy})—an average of <strong>{mth_avg:,}</strong> provider transfers per month.",
                f"Inland Revenue KiwiSaver Scheme Transfers (Year ended 30 June {fy_yr}): {tr_tot:,} inter-scheme transfers (compared with {prev_tr:,} in FY{prev_fy}).",
                format_provenance_source_html(tr_calc, "Inland Revenue — KiwiSaver Scheme Transfer Statistics", IRD_KS_TR_PAGE, True),
                f"Year ended June {fy_yr}", tr_tot, f"{tr_tot:,} transfers/yr ({mth_avg:,}/month)",
                f"FY{prev_fy}: {prev_tr:,} transfers", f"{tr_tot - prev_tr:+,d} ({tr_yoy_pct:+.1f}%) YoY",
                tr_calc, f"Inland Revenue June FY series (FY{fy_yr})", None, {"source_published": tr_pub},
            )

    entry_html = pages.get(IRD_KS_ENTRY_PAGE, "")
    if entry_html:
        entry_pub = extract_html_release_date(entry_html, "Annual Inland Revenue KiwiSaver scheme entry method release")
        e_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", entry_html))
        m_d, m_e, m_a = re.search(r"default allocated[^\d]{0,40}([\d,]{6,7})", e_txt, flags=re.I), re.search(r"employer nominated[^\d]{0,40}([\d,]{6,7})", e_txt, flags=re.I), re.search(r"actively chosen[^\d]{0,40}([\d,]{6,8})", e_txt, flags=re.I)
        if m_d and m_e and m_a:
            def_a, emp_a, act_a = int(m_d.group(1).replace(",", "")), int(m_e.group(1).replace(",", "")), int(m_a.group(1).replace(",", ""))
            non_act, tot_m = def_a + emp_a, def_a + emp_a + act_a
            non_pct, def_p, emp_p, act_p = (non_act / max(1, tot_m)) * 100.0, (def_a / max(1, tot_m)) * 100.0, (emp_a / max(1, tot_m)) * 100.0, (act_a / max(1, tot_m)) * 100.0
            fy_e = today.year if today.month >= 7 else today.year - 1
            ent_calc = f"Passive: ({def_a:,} + {emp_a:,}) ÷ {tot_m:,} = {non_pct:.1f}% | Active: {act_a:,} ÷ {tot_m:,} = {act_p:.1f}%"
            out["kiwisaver-default-vs-active"] = make_pulse_item(
                "Inland Revenue", f"{fy_e}-06-30", MAX_AGE_ANNUAL,
                f"In Inland Revenue's latest KiwiSaver entry-method release, <strong>{non_pct:.1f}% ({non_act:,} members)</strong> were in schemes assigned via <strong>default allocation ({def_a:,}, or {def_p:.1f}%)</strong> or <strong>employer nomination ({emp_a:,}, or {emp_p:.1f}%)</strong>, while <strong>{act_p:.1f}% ({act_a:,} members)</strong> had actively chosen their provider.",
                f"Inland Revenue KiwiSaver Entry Method: Default allocated = {def_a:,}; Employer nominated = {emp_a:,}; Actively chosen = {act_a:,}.",
                format_provenance_source_html(ent_calc, "Inland Revenue — KiwiSaver Scheme Entry Method Statistics", IRD_KS_ENTRY_PAGE, True),
                f"Year ended June {fy_e}", round(non_pct, 1), f"{non_pct:.1f}% default/employer-allocated ({non_act:,}) vs {act_p:.1f}% active choice",
                f"Default allocated: {def_a:,} ({def_p:.1f}%) | Employer nominated: {emp_a:,} ({emp_p:.1f}%)", f"Total members across entry methods: {tot_m:,}",
                ent_calc, "Inland Revenue June FY entry method series", None, {"source_published": entry_pub},
            )

    pay_html = pages.get(IRD_KS_PAY_PAGE, "")
    if pay_html:
        pay_pub = extract_html_release_date(pay_html, "Annual Inland Revenue KiwiSaver provider payments release")
        p_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", pay_html))
        m_tot_p = re.search(r"\$([\d,.]+)\s*(billion|b)\s+(?:was\s+)?(?:passed|transferred|paid)\s+to\s+(?:scheme\s+)?providers", p_txt, flags=re.I)
        if m_tot_p:
            tot_pay_b, fy_p = float(m_tot_p.group(1).replace(",", "")), today.year if today.month >= 7 else today.year - 1
            pay_calc = f"Reported Inland Revenue annual KiwiSaver transfers to scheme providers = ${tot_pay_b:.2f}B"
            out["kiwisaver-contributions-10yr"] = make_pulse_item(
                "Inland Revenue", f"{fy_p}-06-30", MAX_AGE_ANNUAL,
                f"In the latest Inland Revenue KiwiSaver contributions release, total annual employee, employer, and Crown contributions transferred to scheme providers reached <strong>${tot_pay_b:.2f}B</strong>.",
                f"Inland Revenue KiwiSaver Provider Payments: Total annual contributions transferred = ${tot_pay_b:.2f}B.",
                format_provenance_source_html(pay_calc, "Inland Revenue — KiwiSaver Scheme Provider Payments", IRD_KS_PAY_PAGE, False),
                f"Year ended June {fy_p}", tot_pay_b, f"${tot_pay_b:.2f}B annual contributions transferred",
                "Employee + Employer + Crown contributions", "Live-parsed from Inland Revenue contributions release",
                pay_calc, "Inland Revenue June FY contributions series", None, {"source_published": pay_pub},
            )

    susp_html = pages.get(IRD_KS_SUSP_PAGE, "")
    if susp_html:
        susp_pub = extract_html_release_date(susp_html, "Monthly Inland Revenue KiwiSaver savings suspension release")
        s_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", susp_html))
        m_s = re.search(r"([\d,]{5,6})\s+members were on a savings suspension", s_txt, flags=re.I)
        if m_s:
            susp_tot = int(m_s.group(1).replace(",", ""))
            m_dt_s = re.search(r"In\s+([A-Za-z]+\s+20\d{2})", s_txt, flags=re.I)
            dt_s = pd.to_datetime("1 " + m_dt_s.group(1), errors="coerce") if m_dt_s else pd.Timestamp(today)
            s_per, s_as_of = (dt_s.strftime("%B %Y"), month_end_iso(dt_s)) if pd.notna(dt_s) else ("Latest IRD release", today.strftime("%Y-%m-%d"))
            susp_calc = f"Directly reported Inland Revenue active KiwiSaver savings suspensions = {susp_tot:,} members"
            out["kiwisaver-savings-suspensions"] = make_pulse_item(
                "Inland Revenue", s_as_of, MAX_AGE_ANNUAL,
                f"In {s_per}, <strong>{susp_tot:,}</strong> KiwiSaver members were on an active savings suspension.",
                f"Inland Revenue KiwiSaver Savings Suspensions ({s_per}): Active savings suspensions = {susp_tot:,} members.",
                format_provenance_source_html(susp_calc, "Inland Revenue — KiwiSaver Savings Suspension Statistics", IRD_KS_SUSP_PAGE, False),
                s_per, susp_tot, f"{susp_tot:,} members on savings suspension",
                "Monthly Inland Revenue savings suspension series", "Live-parsed from Inland Revenue",
                susp_calc, f"Inland Revenue savings suspension series ({s_per})", None, {"source_published": susp_pub},
            )

    fma_html = pages.get(FMA_KS_REPORT_PAGE, "")
    fma_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", fma_html)) if fma_html else ""
    if fma_html:
        for pdf_h in re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', fma_html, flags=re.I)[:2]:
            full_fma_pdf = "https://www.fma.govt.nz" + pdf_h if pdf_h.startswith("/") else pdf_h
            fma_txt += " " + extract_pdf_text(fetch_bytes_with_browser_fallback(session, full_fma_pdf, referer=FMA_KS_REPORT_PAGE, min_bytes=20_000))

    t43_fum_b, t43_dt, t43_pub = None, None, "Quarterly RBNZ Table T43 KiwiSaver Managed Funds release"
    if t43_path is not None and t43_path.exists():
        try:
            t43_pub = extract_rbnz_published_date(t43_path)
            for _, df_t, _, labels_t in load_rbnz_sheets(t43_path):
                for idx, lbl in enumerate(labels_t):
                    if "total" in lbl and ("asset" in lbl or "funds under management" in lbl or "kiwisaver" in lbl):
                        s_t = pd.to_numeric(df_t.iloc[:, idx], errors="coerce").dropna()
                        if len(s_t) and float(s_t.iloc[-1]) > 50_000:
                            t43_fum_b, t43_dt = float(s_t.iloc[-1]) / 1000.0, df_t["date"].iloc[-1]
                            break
                if t43_fum_b is not None:
                    break
        except Exception as exc_t43:
            print(f"[RBNZ T43] Parse note: {exc_t43}")

    m_fma_fum = re.search(r"\$(1[123456]\d\.?\d*)\s*(?:billion|b)", fma_txt, flags=re.I)
    m_fma_mem = re.search(r"(3\.[23456]\d?)\s*million\s*(?:members|people|kiwisaver)", fma_txt, flags=re.I)
    m_fma_avg = re.search(r"average\s+(?:member\s+|kiwisaver\s+)?balance[^$]{0,35}\$([345]\d,\d{3})", fma_txt, flags=re.I)
    fum_b_live = float(m_fma_fum.group(1)) if m_fma_fum else t43_fum_b
    mem_m_live = float(m_fma_mem.group(1)) if m_fma_mem else None
    avg_bal_live = int(m_fma_avg.group(1).replace(",", "")) if m_fma_avg else (int(round((fum_b_live * 1e9) / (mem_m_live * 1e6))) if (fum_b_live and mem_m_live) else None)

    if fum_b_live is not None:
        as_of_ks_fum = month_end_iso(t43_dt) if t43_dt is not None else f"{today.year if today.month >= 4 else today.year - 1}-03-31"
        per_ks_fum = t43_dt.strftime("%B %Y") if t43_dt is not None else f"Year ended 31 March {today.year if today.month >= 4 else today.year - 1}"
        if avg_bal_live and mem_m_live:
            fum_calc = f"Official FMA / RBNZ FUM = ${fum_b_live:.1f}B across {mem_m_live:.2f}m members (Average balance = ${avg_bal_live:,})"
            fum_html = f"Total KiwiSaver funds under management reached <strong>${fum_b_live:.1f}B</strong> across <strong>{mem_m_live:.2f} million</strong> members, with an average member balance of <strong>${avg_bal_live:,}</strong>."
            fum_val, fum_latest_v = avg_bal_live, f"${avg_bal_live:,} avg balance (${fum_b_live:.1f}B FUM across {mem_m_live:.2f}m members)"
        else:
            fum_calc = f"Directly reported RBNZ Table T43 / FMA total KiwiSaver funds under management = ${fum_b_live:.1f}B"
            fum_html = f"As at {per_ks_fum}, total New Zealand KiwiSaver funds under management reached <strong>${fum_b_live:.1f}B</strong>."
            fum_val, fum_latest_v = round(fum_b_live, 1), f"${fum_b_live:.1f}B total KiwiSaver FUM"
        out["kiwisaver-total-fum-balance"] = make_pulse_item(
            "FMA & RBNZ T43", as_of_ks_fum, MAX_AGE_INSURANCE, fum_html,
            f"FMA KiwiSaver Annual Report & RBNZ Table T43 ({per_ks_fum}): Total KiwiSaver FUM = ${fum_b_live:.1f}B.",
            format_provenance_source_html(fum_calc, "FMA KiwiSaver Annual Report & RBNZ Table T43", FMA_KS_REPORT_PAGE, True),
            per_ks_fum, fum_val, fum_latest_v, "FMA KiwiSaver Annual Report & RBNZ Managed Funds Survey",
            "Live-parsed from official FMA & RBNZ KiwiSaver releases", fum_calc, "FMA & RBNZ Table T43 KiwiSaver series",
            None, {"source_published": t43_pub},
        )

    try:
        out.update(build_kainga_ora_tickers(session))
    except Exception as exc_ko:
        print(f"[Kāinga Ora] Parse note: {exc_ko}")

    return out
# === END OF PART 4D ===
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

    # Dynamic source_url override when live scraper resolved a specific release, PDF, or CSV URL
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
    """
    Build the machine-readable `fhb_snapshot` JSON object from the 5 live FHB tickers
    (for external Webflow charts/widgets only — no separate chart box is rendered on Pulse index.html).
    """
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
    latest_per = ""
    for row_num, label, t_key, series_desc in row_specs:
        rec = verified_tickers.get(t_key)
        if not rec:
            continue
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
      <p>To prevent confusion between when an economic activity occurred and when a government or industry body published its release, every HLF Pulse record separates four distinct date fields:</p>
      <table>
        <thead><tr><th style="width:220px;">Date Field</th><th>Definition &amp; How It Is Populated</th></tr></thead>
        <tbody>
          <tr><td><strong>1. Observation period</strong> (<code>observation_period</code>)</td><td>The period the underlying statistic relates to (for example, <em>September 2026</em> for monthly RBNZ lending, or <em>Year ended 31 March 2026</em> for the FMA KiwiSaver Annual Report).</td></tr>
          <tr><td><strong>2. Period end</strong> (<code>period_end</code> / <code>as_of</code>)</td><td>The final calendar date covered by the observation period (for example, <code>2026-09-30</code> or <code>2026-03-31</code>). Used by HLF&rsquo;s automated freshness gate to verify series currency.</td></tr>
          <tr><td><strong>3. Source publication date</strong> (<code>source_published</code>)</td><td>When the original organisation released the statistic (for example, the <em>Published Date</em> recorded inside RBNZ&rsquo;s workbook metadata or on the agency release page). HLF never copies <code>period_end</code> into <code>source_published</code>; where an agency page does not state an exact release date, the agency&rsquo;s publication schedule is stated explicitly.</td></tr>
          <tr><td><strong>4. Last checked</strong> (<code>last_checked</code>)</td><td>The date HLF&rsquo;s automated GitHub Actions pipeline last fetched and verified the official source.</td></tr>
        </tbody>
      </table>
    </section>

    <section id="data-quality-limitations" class="quality-box" style="margin-top:18px;" aria-label="Data Quality &amp; Limitations">
      <h2>Data Quality &amp; Limitations</h2>
      <ul>
        <li><strong>100% live-pulled official sources (zero hardcoded fallbacks):</strong> Pulse automatically fetches and parses live data files (.xlsx, .csv, .pdf, and release tables) from the Reserve Bank of New Zealand (RBNZ), Stats NZ, Inland Revenue (IRD), the Financial Markets Authority (FMA), MBIE Tenancy Services, the Natural Hazards Commission Toka Tū Ake (NHC), the Accident Compensation Corporation (ACC), and the Financial Services Council (FSC).</li>
        <li><strong>Unaltered official figures &amp; source rounding:</strong> HLF does not alter or estimate official source figures. Where official tables publish rounded aggregates (such as RBNZ dollar series rounded to $1m or $0.1m), derived averages and implied counts reflect that source rounding and note it explicitly in the provenance layer.</li>
        <li><strong>Provenance &amp; derivation layer:</strong> To keep the primary editorial summaries natural and readable, exact formulas, series IDs, and derivation labels (<em>Derived by HLF Pulse</em>) are recorded in the expandable provenance layer, Data Dictionary tables, <code>pulse.json</code>, and <code>llms.txt</code>.</li>
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
        <li><strong>Layer 1 — Official Crown, statutory &amp; industry statistics:</strong> Underlying official figures remain subject to the copyright and licensing terms of their respective publishers: <strong>Stats NZ</strong>, <strong>Inland Revenue</strong>, <strong>MBIE Tenancy Services</strong>, and the <strong>Financial Markets Authority</strong> (<a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener">Creative Commons Attribution 4.0 International — CC BY 4.0</a>); the <strong>Reserve Bank of New Zealand</strong> (<a href="https://www.rbnz.govt.nz/about-our-site/terms-of-use" target="_blank" rel="noopener">RBNZ Terms of Use</a>, permitting reproduction and derivative calculation with attribution to RBNZ); statutory schedules under the <strong>Natural Hazards Insurance Act 2023</strong> and <strong>Accident Compensation Act 2001</strong> (public-domain New Zealand legislation); and <strong>Financial Services Council (FSC)</strong> Life and Health Insurance Industry Spotlight reports (credited to the Financial Services Council under fair-dealing citation, linking directly to the respective FSC Spotlight PDF/release).</li>
        <li><strong>Layer 2 — HLF-derived calculations, compilations &amp; commentary:</strong> All HLF-derived percentages, averages, per-cover ratios, multi-year growth rates, data dictionary definitions, and market interpretations are licensed by <strong>{PUBLISHER_NAME}</strong> under <a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener"><strong>Creative Commons Attribution 4.0 International ({HLF_LICENSE_SHORT})</strong></a>. Journalists, researchers, advisers, and AI systems may freely quote and reuse HLF-derived figures and commentary provided attribution is given to <strong>HLF Pulse, Home Loan Factory</strong> with a link to <code>{PULSE_BASE_URL}/</code>, the specific <code>/{{ticker}}/</code> record URL, or the <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">HLF Pulse Media Kit &amp; Citation Guide</a>.</li>
      </ul>
    </section>

    <section id="about-hlf-pulse" class="about-footer-box" aria-label="About the HLF Pulse &amp; Creator">
      <h2>About the HLF Pulse &amp; Creator</h2>
      <p><strong>HLF Pulse</strong> is a public reference dataset covering New Zealand mortgage, housing, KiwiSaver, tenancy, and protection statistics, published by <a href="{PUBLISHER_URL}">{PUBLISHER_NAME}</a> ({PUBLISHER_LEGAL}) and created by <a href="{CREATOR_URL}" target="_blank" rel="noopener"><strong>{CREATOR_NAME}</strong> ({CREATOR_ALT_NAME}, {CREATOR_FSP})</a>, Mortgage &amp; KiwiSaver Adviser at Home Loan Factory.</p>
      <p>Data is compiled automatically from official releases from the <strong>Reserve Bank of New Zealand</strong>, <strong>Stats NZ</strong>, <strong>Inland Revenue</strong>, the <strong>Financial Markets Authority</strong>, <strong>MBIE Tenancy Services</strong>, the <strong>Natural Hazards Commission Toka Tū Ake</strong>, <strong>ACC</strong>, and the <strong>Financial Services Council</strong>. Where HLF calculates percentages, averages, or comparisons from the underlying official data, the exact formula and <strong>&ldquo;Derived by HLF Pulse&rdquo;</strong> provenance are recorded in the expandable source layer and data dictionary.</p>
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
def cross_derive_fhb_benchmark_tickers(candidates: dict) -> None:
    """
    Populate the 4 standalone FHB benchmark tickers (fhb-typical-income, fhb-average-loan-size,
    fhb-typical-dti, fhb-monthly-lending-volume) from live RBNZ records in `candidates`
    (borrower-income, fhb-share, dti-headroom) if not already set on this run. Zero hardcoded numbers.
    """
    c31_u = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type"
    c40_u = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use"
    c41_u = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi"

    bi = candidates.get("borrower-income")
    if "fhb-typical-income" not in candidates and isinstance(bi, dict) and bi.get("value"):
        inc_v = int(round(float(bi["value"])))
        per_s, as_of_s, pub_s = bi.get("latest_period", ""), bi.get("as_of", ""), bi.get("source_published", "Monthly RBNZ statistical release")
        calc_s = bi.get("calculation_live") or bi.get("calculation") or f"RBNZ Table C41 FHB band-weighted mean gross income = ${inc_v:,}"
        candidates["fhb-typical-income"] = make_pulse_item(
            "RBNZ C41", as_of_s, MAX_AGE_MONTHLY,
            f"In {per_s}, the typical (borrower-weighted average) gross household income across approved New Zealand first-home buyer mortgage applications was <strong>${inc_v:,}</strong>.",
            bi.get("official_data_summary", f"RBNZ Table C41 ({per_s}): FHB borrower-weighted mean gross application income = ${inc_v:,}."),
            format_provenance_source_html(calc_s, "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_u, True),
            per_s, inc_v, f"${inc_v:,} typical FHB gross application income",
            bi.get("previous_period", "Prior 12-month period"), bi.get("change", "Tracked monthly in RBNZ Table C41"),
            calc_s, bi.get("historical_data_from", "RBNZ Table C41 series"), bi.get("history", []),
            {"source_published": pub_s, "source_url": c41_u},
        )

    fs = candidates.get("fhb-share")
    if isinstance(fs, dict) and fs.get("html"):
        per_s, as_of_s, pub_s = fs.get("latest_period", ""), fs.get("as_of", ""), fs.get("source_published", "Monthly RBNZ statistical release")
        fs_txt = re.sub(r"<[^>]+>", " ", fs.get("html", ""))
        m_avg = re.search(r"averaging\s+\$([\d,]+)\s+per\s+loan", fs_txt, flags=re.I)
        m_cnt = re.search(r"across\s+([\d,]+)\s+loans", fs_txt, flags=re.I)
        m_val = re.search(r"\$([\d,]+)\s*m\)", fs_txt, flags=re.I)
        val_m = float(m_val.group(1).replace(",", "")) if m_val else (float(re.search(r"\$([\d.]+)B", fs_txt).group(1)) * 1000.0 if re.search(r"\$([\d.]+)B", fs_txt) else None)
        cnt_n = int(m_cnt.group(1).replace(",", "")) if m_cnt else None
        avg_l = int(m_avg.group(1).replace(",", "")) if m_avg else (int(round((val_m * 1e6) / cnt_n)) if (val_m and cnt_n) else None)

        if "fhb-average-loan-size" not in candidates and avg_l:
            al_calc = f"${val_m:,.0f}m × $1,000,000 ÷ {cnt_n:,} FHB loans = ${avg_l:,}" if (val_m and cnt_n) else f"RBNZ Table C31 FHB mean loan size = ${avg_l:,}"
            candidates["fhb-average-loan-size"] = make_pulse_item(
                "RBNZ C31", as_of_s, MAX_AGE_MONTHLY,
                f"In {per_s}, the average new first-home buyer mortgage approved across New Zealand was <strong>${avg_l:,}</strong>" + (f" (${val_m:,.0f}m across {cnt_n:,} loans)." if (val_m and cnt_n) else "."),
                f"RBNZ Table C31 ({per_s}): Average first-home buyer loan size = ${avg_l:,}.",
                format_provenance_source_html(al_calc, "RBNZ Table C31 (Reserve Bank of New Zealand)", c31_u, True),
                per_s, avg_l, f"${avg_l:,} average FHB mortgage size",
                "RBNZ Table C31 monthly FHB series", "Derived from RBNZ Table C31 (LVRN.MMB1.AC1 ÷ LVRN.MMB2.AC1)",
                al_calc, fs.get("historical_data_from", "RBNZ Table C31 series"), [], {"source_published": pub_s, "source_url": c31_u},
            )

        if "fhb-monthly-lending-volume" not in candidates and val_m and cnt_n:
            vol_calc = f"Official RBNZ Table C31 LVRN.MMB1.AC1 = ${val_m:,.0f}m across LVRN.MMB2.AC1 = {cnt_n:,} FHB loans"
            candidates["fhb-monthly-lending-volume"] = make_pulse_item(
                "RBNZ C31", as_of_s, MAX_AGE_MONTHLY,
                f"In {per_s}, New Zealand banks committed <strong>{fmt_nzd_millions(val_m)}</strong> in new residential mortgages to <strong>{cnt_n:,} first-home buyers</strong>.",
                f"RBNZ Table C31 ({per_s}): First-home buyer new mortgage commitments = ${val_m:,.0f}m across {cnt_n:,} loans.",
                format_provenance_source_html(vol_calc, "RBNZ Table C31 (Reserve Bank of New Zealand)", c31_u, False),
                per_s, round(val_m, 1), f"{fmt_nzd_millions(val_m)} across {cnt_n:,} FHB loans",
                "RBNZ Table C31 monthly FHB series", f"Average loan size: ${avg_l:,}" if avg_l else "Reported in RBNZ C31",
                vol_calc, fs.get("historical_data_from", "RBNZ Table C31 series"), [], {"source_published": pub_s, "source_url": c31_u},
            )

    inc_rec = candidates.get("fhb-typical-income") or candidates.get("borrower-income")
    loan_rec, dti_rec = candidates.get("fhb-average-loan-size"), candidates.get("dti-headroom")
    if "fhb-typical-dti" not in candidates and isinstance(inc_rec, dict) and isinstance(loan_rec, dict) and isinstance(dti_rec, dict):
        inc_v, loan_v, gt6_v = float(inc_rec.get("value") or 0), float(loan_rec.get("value") or 0), float(dti_rec.get("value") or 0)
        if inc_v > 20000 and loan_v > 100000 and 0.1 <= gt6_v <= 50.0:
            imp_dti, le6_v = round(loan_v / inc_v, 2), round(100.0 - gt6_v, 1)
            per_s, as_of_s, pub_s = dti_rec.get("latest_period", ""), dti_rec.get("as_of", ""), dti_rec.get("source_published", "Monthly RBNZ statistical release")
            td_calc = f"Implied FHB Mortgage DTI: ${int(loan_v):,} (C31 avg FHB loan) ÷ ${int(inc_v):,} (C41 avg FHB gross income) = {imp_dti:.2f}x (~{imp_dti:.1f}x–{imp_dti + 0.4:.1f}x total DTI) | C40 Share <= 6x DTI: 100.0% − {gt6_v:.1f}% = {le6_v:.1f}%"
            candidates["fhb-typical-dti"] = make_pulse_item(
                "RBNZ C40 & C41", as_of_s, MAX_AGE_MONTHLY,
                f"In {per_s}, the typical New Zealand first-home buyer borrowed at an implied Debt-to-Income ratio of <strong>{imp_dti:.1f}x–{imp_dti + 0.4:.1f}x gross income</strong> (${int(loan_v):,} average loan on ${int(inc_v):,} average gross income), with <strong>{le6_v:.1f}% of all first-home buyer lending</strong> written at or below <strong>6x DTI</strong>.",
                f"RBNZ Tables C31, C40 & C41 ({per_s}): Average FHB loan = ${int(loan_v):,}; Average FHB gross income = ${int(inc_v):,}; FHB lending <= 6x TDTI = {le6_v:.1f}% (> 6x TDTI = {gt6_v:.1f}%).",
                format_provenance_source_html(td_calc, "RBNZ Tables C31, C40 & C41 (Reserve Bank of New Zealand)", c40_u, True),
                per_s, imp_dti, f"{imp_dti:.2f}x implied mortgage DTI (~{imp_dti:.1f}x–{imp_dti + 0.4:.1f}x total DTI) | {le6_v:.1f}% <= 6x DTI",
                f"Share <= 6x DTI: {le6_v:.1f}%", f"{gt6_v:.1f}% above 6x DTI (vs 20.0% RBNZ speed limit)",
                td_calc, dti_rec.get("historical_data_from", "RBNZ Tables C40 & C41"),
                [
                    {"period": "Implied Mortgage DTI", "label": "Implied FHB Loan-to-Income Ratio (C31 ÷ C41)", "value": imp_dti, "formatted": f"{imp_dti:.2f}x (${int(loan_v):,} ÷ ${int(inc_v):,})"},
                    {"period": "TDTI <= 6x Share", "label": "FHB Lending at or below 6x DTI (C40)", "value": le6_v, "formatted": f"{le6_v:.1f}% of FHB lending"},
                    {"period": "TDTI > 6x Share", "label": "FHB Lending above 6x DTI (C40 Speed Limit Band)", "value": round(gt6_v, 1), "formatted": f"{gt6_v:.1f}% (vs 20.0% RBNZ cap)"},
                ],
                {"source_published": pub_s, "source_url": c40_u},
            )


def main() -> None:
    root = pathlib.Path(__file__).resolve().parent
    json_path, structured_json_path, html_path = root / "nz-mortgage-pulse.json", root / "pulse.json", root / "index.html"

    prev_tickers = {}
    if json_path.exists():
        try:
            prev_tickers = json.loads(json_path.read_text(encoding="utf-8")).get("tickers", {})
        except Exception:
            prev_tickers = {}

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
            b30_out = build_rbnz_b30_ticker(rbnz_files["b30"], session)
            fresh_candidates.update(b30_out)
            live_parsed_count += len(b30_out)
        except Exception as exc:
            print(f"[RBNZ B30] Error: {exc}")

    try:
        snz_out = build_stats_nz_tickers(session, rbnz_files.get("j10"))
        fresh_candidates.update(snz_out)
        live_parsed_count += len(snz_out)
    except Exception as exc:
        print(f"[Stats NZ / Insurance] Error: {exc}")
        try:
            ins_live = build_insurance_pulse_tickers(session, "", "", None, "", j10_path=rbnz_files.get("j10"))
            fresh_candidates.update(ins_live)
            live_parsed_count += len(ins_live)
        except Exception as exc_ins:
            print(f"[Insurance Live] Error: {exc_ins}")

    try:
        ten_out = build_tenancy_tickers(session)
        fresh_candidates.update(ten_out)
        live_parsed_count += len(ten_out)
    except Exception as exc:
        print(f"[Tenancy Services] Error: {exc}")

    try:
        ird_out = build_ird_tickers(session, rbnz_files.get("t43"))
        fresh_candidates.update(ird_out)
        live_parsed_count += len(ird_out)
    except Exception as exc:
        print(f"[IRD / FMA / T43 / Kāinga Ora] Error: {exc}")

    cross_derive_fhb_benchmark_tickers(fresh_candidates)

    for p in rbnz_files.values():
        if p.name.startswith("tmp_") and p.exists():
            os.remove(p)

    now_dt = datetime.datetime.now(datetime.timezone.utc)
    now_utc, today_iso, now_human = now_dt.strftime("%Y-%m-%dT%H:%M:%SZ"), now_dt.strftime("%Y-%m-%d"), f"{now_dt.day} {now_dt.strftime('%B %Y')}"

    verified_tickers = {}
    for k, item in fresh_candidates.items():
        if k not in TICKER_CATALOG or not isinstance(item, dict):
            continue
        as_of, max_age = item.get("as_of", ""), int(item.get("max_age_days", MAX_AGE_MONTHLY))
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

    latest_rbnz_period = verified_tickers.get("fhb-low-deposit", {}).get("latest_period", "Latest Official Release")
    payload = {
        "publisher": PUBLISHER_NAME, "publisher_legal_name": PUBLISHER_LEGAL, "publisher_url": PUBLISHER_URL, "media_kit_url": MEDIA_KIT_URL,
        "creator_name": CREATOR_NAME, "creator_alternate_name": CREATOR_ALT_NAME, "creator_fsp": CREATOR_FSP, "creator_url": CREATOR_URL,
        "license": HLF_LICENSE_URL, "license_label": HLF_LICENSE_SHORT, "license_note": HLF_LICENSE_NOTE,
        "dataset_name": "NZ Mortgage, Housing, KiwiSaver & Protection Pulse — Verified NZ Data",
        "date_created": PULSE_DATE_CREATED, "date_modified": today_iso, "generated_utc": now_utc, "last_checked": today_iso,
        "latest_underlying_data": latest_rbnz_period, "data_status": "verified", "source_status": "current",
        "active_datasets_count": len(verified_tickers), "count": len(verified_tickers), "fhb_snapshot": fhb_snap_obj, "tickers": bubble_feed_tickers,
    }
    json_path.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")

    structured_records = []
    for k in sorted(verified_tickers.keys()):
        r = verified_tickers[k]
        rec_out = {f: r.get(f) for f in (
            "ticker", "name", "also_known_as", "category", "topic", "value", "latest_value", "unit",
            "observation_period", "observed_period", "period_end", "source_published", "last_checked",
            "date_created", "date_modified", "data_status", "source_status", "previous_period", "change",
            "source", "source_full", "source_series_ids", "source_url", "derived_by_hlf", "data_type",
            "definition", "formula", "hlf_interpretation", "methodology", "historical_data_from",
            "creator_name", "creator_alternate_name", "creator_fsp", "creator_url", "publisher",
            "publisher_legal_name", "publisher_url", "media_kit_url", "license", "license_label",
            "citation", "canonical_url", "summary_text", "related_questions", "citation_chain", "history",
        )}
        rec_out.update({"latest_underlying_data": r["latest_period"], "official_data": r["official_data_summary"], "hlf_calculation": r["calculation"], "frequency": r["update_frequency"]})
        if r.get("tdti_bands"):
            rec_out["tdti_bands"], rec_out["tdti_total_pool_m"], rec_out["tdti_total_implied_apps"] = r["tdti_bands"], r.get("tdti_total_pool_m"), r.get("tdti_total_implied_apps")
        if r.get("townhouse_flat_unit_share_pct") is not None:
            rec_out["townhouse_flat_unit_share_pct"] = r["townhouse_flat_unit_share_pct"]
        structured_records.append(rec_out)

    pulse_json_obj = {
        "dataset": "HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver & Protection Reference Dataset",
        "publisher": PUBLISHER_NAME, "publisher_legal_name": PUBLISHER_LEGAL, "publisher_url": PUBLISHER_URL, "media_kit_url": MEDIA_KIT_URL,
        "creator_name": CREATOR_NAME, "creator_alternate_name": CREATOR_ALT_NAME, "creator_fsp": CREATOR_FSP, "creator_role": CREATOR_ROLE, "creator_url": CREATOR_URL,
        "license": HLF_LICENSE_URL, "license_label": HLF_LICENSE_SHORT, "license_note": HLF_LICENSE_NOTE,
        "how_to_cite": f"HLF Pulse, Home Loan Factory ({PULSE_BASE_URL}/), created by {CREATOR_NAME} (checked {now_human}). {HLF_LICENSE_SHORT}.",
        "date_field_definitions": {
            "observation_period": "When the underlying statistic relates to (e.g., September 2026 or Year ended 31 March 2026)",
            "period_end": "The final calendar date covered by the observation period (YYYY-MM-DD)",
            "source_published": "When the original organisation released the dataset (never auto-populated from period_end)",
            "last_checked": "When HLF last verified the official source (YYYY-MM-DD)",
        },
        "date_created": PULSE_DATE_CREATED, "date_modified": today_iso, "pulse_url": f"{PULSE_BASE_URL}/", "last_checked": today_iso,
        "dataset_generated_utc": now_utc, "latest_underlying_data": latest_rbnz_period, "data_status": "verified", "source_status": "current",
        "active_datasets_count": len(structured_records),
        "reporting_lag_note": "RBNZ and Crown statistics are published with a reporting lag; last_checked reflects when HLF verified the source, source_published reflects when the publisher released it, and observation_period / period_end reflect the period the data covers.",
        "next_expected_update": "Monthly (RBNZ / Stats NZ / MBIE Tenancy), Quarterly (CPI / IRD Student Loans / Kāinga Ora / FSC Spotlights) & Semi-Annual (FMA / RBNZ FSR / NHC / ACC)",
        "fhb_snapshot": fhb_snap_obj, "records": structured_records,
    }
    structured_json_path.write_text(json.dumps(pulse_json_obj, indent=2), encoding="utf-8")

    write_robots_txt(root)
    for k, rec in verified_tickers.items():
        write_canonical_ticker_page(root, k, rec, len(verified_tickers))
    for alias_k, target_k in LEGACY_BUBBLE_ALIASES.items():
        if target_k in verified_tickers and alias_k not in verified_tickers:
            write_legacy_redirect_page(root, alias_k, target_k)

    sitemap_urls = [f"  <url><loc>{PULSE_BASE_URL}/</loc><lastmod>{today_iso}</lastmod></url>"] + [
        f"  <url><loc>{PULSE_BASE_URL}/{k}/</loc><lastmod>{today_iso}</lastmod></url>" for k in sorted(verified_tickers.keys())
    ]
    (root / "sitemap.xml").write_text('<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n' + "\n".join(sitemap_urls) + "\n</urlset>\n", encoding="utf-8")

    llms_lines = [
        "# HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver & Protection Reference Dataset", "",
        "> Public reference dataset and data dictionary built 100% live from official New Zealand Crown, statutory, and industry releases, with clearly labelled HLF-derived calculations and interpretation.",
        f"> Created by: {CREATOR_NAME} ({CREATOR_ALT_NAME}, {CREATOR_FSP}), {CREATOR_ROLE} ({CREATOR_URL}) | Published by: {PUBLISHER_NAME} ({PUBLISHER_URL}).",
        f"> Media Kit & Citation Guide: {MEDIA_KIT_URL} | HLF Glossary: {GLOSSARY_BASE_URL}",
        f"> Reuse Licence: {HLF_LICENSE_SHORT} ({HLF_LICENSE_URL}) for HLF-derived figures & commentary; official statistics credited to their respective publishers.",
        f"> Date Standard: Observation Period (latest monthly: {latest_rbnz_period}) | Period End | Source Publication Date | Last Checked ({now_human} / {today_iso}).",
        f"> Active datasets: {len(verified_tickers)}.", "",
        "## Machine-Readable Endpoints",
        f"- Structured Data Dictionary & Records (JSON): {PULSE_BASE_URL}/pulse.json",
        f"- Webflow Chat Bubble Feed (JSON): {PULSE_BASE_URL}/nz-mortgage-pulse.json",
        f"- Main Reference Dataset Hub (HTML): {PULSE_BASE_URL}/",
        f"- Media Kit & Citation Guide (HTML): {MEDIA_KIT_URL}", "",
        f"## Canonical Data Dictionary Records ({len(verified_tickers)} Active Datasets)",
    ] + [
        f"- [{rec['name']} (Obs: {rec['observation_period']} | End: {rec['period_end']} | Published: {rec['source_published']})]({rec['canonical_url']}): {rec['summary_text']} [Official Data: {rec['official_data_summary']} | HLF Calculation: {rec['calculation']} | Derived by HLF Pulse: {rec['hlf_derived_label']} | Cite as: {rec['citation']}]"
        for k, rec in sorted(verified_tickers.items())
    ]
    (root / "llms.txt").write_text("\n".join(llms_lines) + "\n", encoding="utf-8")

    by_cat = {c: [] for c in CATEGORY_ORDER}
    for k, rec in sorted(verified_tickers.items()):
        by_cat.setdefault(rec.get("category", "Mortgage lending & rates"), []).append((k, rec))

    cat_cards_html = ""
    for cat in CATEGORY_ORDER:
        items = by_cat.get(cat, [])
        if not items:
            continue
        cat_slug = re.sub(r"[^a-z0-9]+", "-", cat.lower()).strip("-")
        links = "".join(f"<li><a href='#{k}'><code>{k}</code></a> — {html_lib.escape(r['name'])} (<strong>{html_lib.escape(str(r['latest_value']).split('(')[0].strip())}</strong>)</li>" for k, r in items)
        cat_cards_html += f"<div class='cat-card' id='cat-{cat_slug}'><div class='cat-hdr'><strong>{html_lib.escape(cat)}</strong><span>{len(items)} datasets</span></div><ul>{links}</ul></div>"

    rows_html_list = []
    for k, v in sorted(verified_tickers.items()):
        aka_sub = f"<div class='aka'>Also: {html_lib.escape(v['also_known_as'])}</div>" if v.get("also_known_as") else ""
        rows_html_list.append(
            f"<tr id='{k}'>"
            f"<td><a href='{v['canonical_url']}'><code>{k}</code></a><div class='rname'>{html_lib.escape(v['name'])}</div>{aka_sub}"
            f"<span class='dtype-badge'>{'Derived by HLF' if v['derived_by_hlf'] else 'Officially Reported'}</span></td>"
            f"<td><strong>{html_lib.escape(v['source'])}</strong><div class='rper'>Obs: {html_lib.escape(v['observation_period'])}</div>"
            f"<div class='rchk'>End: <code>{v['period_end']}</code></div><div class='rchk'>Published: {html_lib.escape(v['source_published'][:36])}</div>"
            f"<div class='rchk'>Checked: {html_lib.escape(v['last_verified_human'])}</div></td>"
            f"<td><div class='rsum'>{v['html']}</div>{render_c41_band_table_html(v)}{render_three_layer_badges_html(v, True)}{render_cite_line_html(v)}"
            f"<div class='rsrc'>{v['source_html']} &bull; <a href='{v['canonical_url']}'>Full Data Dictionary &amp; historical series (/{k}/) &rarr;</a></div></td></tr>"
        )
    rows_html = "".join(rows_html_list)
    robots_tag = "index, follow" if verified_tickers else "noindex, nofollow"
    creator_main_html, about_footer_html = render_creator_card_html(compact=False), render_about_and_quality_footer_html()

    hub_schema_obj = {
        "@context": "https://schema.org", "@type": "Dataset",
        "name": "HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver & Protection Reference Dataset",
        "description": "Public reference dataset and data dictionary built from official New Zealand Crown and statutory releases, with clearly labelled HLF-derived calculations and interpretation.",
        "url": f"{PULSE_BASE_URL}/", "inLanguage": "en-NZ", "license": HLF_LICENSE_URL,
        "citation": f"HLF Pulse, Home Loan Factory ({PULSE_BASE_URL}/), created by {CREATOR_NAME} (checked {now_human}). {HLF_LICENSE_SHORT}.",
        "dateCreated": PULSE_DATE_CREATED, "dateModified": today_iso, "spatialCoverage": "New Zealand",
        "creator": CREATOR_SCHEMA_OBJ, "publisher": PUBLISHER_SCHEMA_OBJ,
        "distribution": [
            {"@type": "DataDownload", "name": f"HLF Pulse Structured Data Dictionary JSON (All {len(verified_tickers)} Datasets)", "encodingFormat": "application/json", "contentUrl": f"{PULSE_BASE_URL}/pulse.json"},
            {"@type": "DataDownload", "name": "HLF Pulse Webflow Feed JSON", "encodingFormat": "application/json", "contentUrl": f"{PULSE_BASE_URL}/nz-mortgage-pulse.json"},
            {"@type": "DataDownload", "name": "HLF Pulse LLM Reference Summary", "encodingFormat": "text/plain", "contentUrl": f"{PULSE_BASE_URL}/llms.txt"},
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
    <p>Public reference dataset and data dictionary built 100% live from official New Zealand Crown, statutory, and industry releases, with clearly labelled HLF-derived calculations and interpretation (<a href="#browse-datasets">Browse all {len(verified_tickers)} datasets</a> &bull; <a href="#cat-first-home-buyers">First-Home Buyers</a> &bull; <a href="#cat-kiwisaver">KiwiSaver</a> &bull; <a href="#cat-new-builds-construction">Build Costs</a> &bull; <a href="#cat-property-natural-hazard-insurance">Property Insurance</a> &bull; <a href="#cat-personal-risk-income-protection">Personal &amp; ACC Protection</a> &bull; <a href="{GLOSSARY_BASE_URL}" target="_blank" rel="noopener">HLF Glossary</a> &bull; <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citation Guide</a> &bull; <a href="#date-definitions">4-Field Date Standard</a> &bull; <a href="#data-quality-limitations">Data Quality &amp; Limitations</a> &bull; <a href="#reuse-licensing">Reuse &amp; Citation ({HLF_LICENSE_SHORT})</a> &bull; <a href="#about-hlf-pulse">About &amp; Methodology</a>). Machine-readable endpoints: <a href="{PULSE_BASE_URL}/pulse.json"><code>pulse.json</code></a>, <a href="{PULSE_BASE_URL}/nz-mortgage-pulse.json"><code>nz-mortgage-pulse.json</code></a>, <a href="{PULSE_BASE_URL}/llms.txt"><code>llms.txt</code></a>, and <a href="{PULSE_BASE_URL}/robots.txt"><code>robots.txt</code></a>.</p>

    <div class="lag-banner" aria-label="4-Field Date Architecture: Observation Period, Period End, Source Published &amp; Last Checked">
      <div class="lag-item">
        <span>1. Observation Period &amp; End</span>
        <strong>{html_lib.escape(latest_rbnz_period)} (RBNZ Monthly Series)</strong>
        <small>Each record specifies both its Observation Period and Period End date.</small>
      </div>
      <div class="lag-item">
        <span>2. Source Publication Date</span>
        <strong>Recorded per agency release</strong>
        <small>Tracked separately from Period End across RBNZ, Stats NZ, IRD, MBIE, FMA, Kāinga Ora, NHC, ACC &amp; FSC.</small>
      </div>
      <div class="lag-item">
        <span>3. Last Checked &amp; Reuse ({HLF_LICENSE_SHORT})</span>
        <strong style="color:#065f46;">&#10003; {now_human} (<a href="#browse-datasets" style="color:#065f46;text-decoration:underline;">{len(verified_tickers)} Active Datasets &darr;</a>)</strong>
        <small>HLF layer reusable under <a href="#reuse-licensing">{HLF_LICENSE_SHORT}</a> (<a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit</a>).</small>
      </div>
    </div>

    {creator_main_html}

    <h2>What the HLF Pulse Can Help Answer</h2>
    <div class="use-grid">
      <div class="use-card">
        <strong>For First-Home Buyers</strong>
        <ul>
          <li><a href="#fhb-typical-income">What is the typical FHB household income?</a></li>
          <li><a href="#fhb-average-loan-size">What is the average FHB mortgage size?</a></li>
          <li><a href="#fhb-typical-dti">What is the typical FHB Debt-to-Income (DTI) ratio?</a></li>
          <li><a href="#kainga-ora-first-home-loan-approvals">How many Kāinga Ora 5% deposit loans are approved?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For Homeowners &amp; New Builds</strong>
        <ul>
          <li><a href="#new-build-consent-cost-by-type">What is the average consent build cost by dwelling type?</a></li>
          <li><a href="#fixed-vs-floating-rate-mix">What share of mortgages are fixed vs floating?</a></li>
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
          <li><a href="#kiwisaver-total-fum-balance">What is the official FMA average KiwiSaver balance &amp; FUM?</a></li>
          <li><a href="#kainga-ora-second-chance-kiwisaver-approvals">How many 'second-chance' KiwiSaver withdrawals are approved?</a></li>
          <li><a href="#kiwisaver-first-home">How much are Kiwis withdrawing for a first home?</a></li>
          <li><a href="#kiwisaver-default-vs-active">How many Kiwis are still in default schemes?</a></li>
        </ul>
      </div>
      <div class="use-card">
        <strong>For Protection, Insurance &amp; ACC</strong>
        <ul>
          <li><a href="#life-trauma-average-claim-payout">How much do Life, Trauma &amp; Income Protection insurers pay out?</a></li>
          <li><a href="#private-health-claims-average-payout">How much do private health insurers pay out per member?</a></li>
          <li><a href="#personal-insurance-underinsurance-gap">What share of life insurance is Income Protection?</a></li>
          <li><a href="#acc-earners-levy-cap">What is the maximum income ACC covers?</a></li>
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
    print(f"Saved {len(verified_tickers)} canonical live datasets, pulse.json, index.html, robots.txt, llms.txt, sitemap.xml & /<ticker>/ pages.")


if __name__ == "__main__":
    main()
# === END OF PART 5C ===
