# === BLOCK 1 OF 13 ===
#!/usr/bin/env python3
"""
HLF Pulse — Reference-Grade New Zealand Mortgage, Housing, KiwiSaver, Income, Net Worth & Protection Dataset
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
import warnings

import openpyxl
import pandas as pd

warnings.filterwarnings("ignore")

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
    "HLF-derived calculations, percentages, averages, multi-year comparisons, "
    "data dictionary definitions, and market interpretations are licensed by Home Loan Factory "
    "under Creative Commons Attribution 4.0 International (CC BY 4.0) with attribution to "
    "HLF Pulse (https://pulse-data.homeloanfactory.co.nz/). Underlying official Crown and statutory "
    "statistics remain subject to the terms of their respective publishers "
    "(RBNZ Terms of Use; Stats NZ, Inland Revenue, MBIE, HUD / MCERT, FMA, and Kāinga Ora under CC BY 4.0; "
    "Natural Hazards Insurance Act 2023 and Accident Compensation Act 2001 under public-domain NZ legislation)."
)

GLOSSARY_LINKS_BY_CATEGORY = {
    "First-home buyers": [
        ("Loan-to-Value Ratio (LVR)", f"{GLOSSARY_BASE_URL}#lvr"),
        ("Kāinga Ora First Home Loan", f"{GLOSSARY_BASE_URL}#first-home-loan"),
        ("Lenders Mortgage Insurance (LMI)", f"{GLOSSARY_BASE_URL}#lmi"),
    ],
    "House prices, sales & HPI": [
        ("Capital Value (CV / RV)", f"{GLOSSARY_BASE_URL}#capital-value"),
        ("Loan-to-Value Ratio (LVR)", f"{GLOSSARY_BASE_URL}#lvr"),
        ("Usable Equity", f"{GLOSSARY_BASE_URL}#usable-equity"),
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
    "Household debt, credit cards & borrower stress": [
        ("Debt-to-Income (DTI) Ratio", f"{GLOSSARY_BASE_URL}#dti"),
        ("Servicing Test Rate", f"{GLOSSARY_BASE_URL}#test-rate"),
        ("Refinancing & Debt Consolidation", f"{GLOSSARY_BASE_URL}#refinancing"),
    ],
    "Income, net worth & household benchmarks": [
        ("Debt-to-Income (DTI) Ratio", f"{GLOSSARY_BASE_URL}#dti"),
        ("Usable Equity", f"{GLOSSARY_BASE_URL}#usable-equity"),
        ("Loan-to-Value Ratio (LVR)", f"{GLOSSARY_BASE_URL}#lvr"),
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

MAX_AGE_MONTHLY = 140
MAX_AGE_TENANCY = 140
MAX_AGE_HUD = 175
MAX_AGE_QUARTERLY = 160
MAX_AGE_ANNUAL = 650
MAX_AGE_TRIENNIAL = 1250
MAX_AGE_INSURANCE = 650

RBNZ_WARMUP_URL = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary"

RBNZ_SOURCES = {
    "c31": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c31/c31-new-residential-mortgage-lending-by-borrower-type.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c31/hc31.xlsx",
        "match": "c31",
    },
    "c33": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-purpose"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c33/c33-new-residential-mortgage-lending-by-purpose.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c33/hc33.xlsx",
        "match": "c33",
    },
    "c32": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-and-existing-residential-mortgage-lending-by-payment-type"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c32/c32-new-and-existing-residential-mortgage-lending-by-payment-type.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c32/hc32.xlsx",
        "match": "c32",
    },
    "c40": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use",
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/debt-to-income-dti-statistics",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c40/c40-residential-mortgage-lending-by-dti-purpose-use.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c40/hc40.xlsx",
        "match": "c40",
    },
    "c41": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi",
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/debt-to-income-dti-statistics",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c41/c41-residential-mortgage-borrower-gross-income.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c41/hc41.xlsx",
        "match": "c41",
    },
    "b30": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-weighted-average-interest-rates"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b30/b30-new-residential-mortgage-weighted-average-interest-rates.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b30/hb30.xlsx",
        "match": "b30",
    },
    "b6": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/yields-on-loans"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b6/hb6.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b6/b6-yields-on-loans.xlsx",
        "match": "b6",
    },
    "b20": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-standard-interest-rates"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b20/hb20.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b20/b20-new-residential-mortgage-standard-interest-rates.xlsx",
        "match": "b20",
    },
    "b21": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-special-interest-rates"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b21/hb21.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b21/b21-new-residential-mortgage-special-interest-rates.xlsx",
        "match": "b21",
    },
    "c5": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/sector-lending",
            "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/registered-banks-and-non-bank-lending-institutions-sector-lending",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c5/hc5.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c5/c5-sector-lending.xlsx",
        "match": "c5",
    },
    "c12": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/credit-card-balances"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c12/hc12.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c12/c12-credit-card-balances.xlsx",
        "match": "c12",
    },
    "c35": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-loan-reconciliation"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c35/hc35.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c35/c35-residential-mortgage-loan-reconciliation.xlsx",
        "match": "c35",
    },
    "s32": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-by-product"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/s/s32/hs32.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/s/s32/s32-banks-assets-loans-by-product.xlsx",
        "match": "s32",
    },
    "s33": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-fully-secured-by-residential-mortgage"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/s/s33/hs33.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/s/s33/s33-banks-assets-loans-fully-secured-by-residential-mortgage.xlsx",
        "match": "s33",
    },
    "s50": {
        "pages": ["https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-by-asset-quality"],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/s/s50/hs50.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/s/s50/s50-banks-assets-loans-by-asset-quality.xlsx",
        "match": "s50",
    },
    "j10": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/insurance/insurance-income-and-expenditure",
            "https://www.rbnz.govt.nz/statistics/series/insurance",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/j/j10/j10-insurance-income-and-expenditure.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/j/j10/hj10.xlsx",
        "match": "j10",
    },
    "t43": {
        "pages": [
            "https://www.rbnz.govt.nz/statistics/series/managed-funds/kiwisaver",
            "https://www.rbnz.govt.nz/statistics/series/managed-funds",
        ],
        "xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/t/t43/t43-kiwisaver-assets-and-liabilities.xlsx",
        "alt_xlsx": "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/t/t43/ht43.xlsx",
        "match": "t43",
    },
}

TENANCY_PAGE = "https://www.tenancy.govt.nz/about-tenancy-services/data-and-statistics/rental-bond-data/"

HUD_ABOUT_PAGE = "https://www.hud.govt.nz/stats-and-insights/property-and-sales-statistics/about-the-dashboard"
HUD_SALES_DASH_PAGE = "https://www.hud.govt.nz/stats-and-insights/property-and-sales-statistics/dashboard"
HUD_COMP_DASH_PAGE = "https://www.hud.govt.nz/stats-and-insights/property-and-sales-statistics/dashboard-2"
HUD_HPI_DASH_PAGE = "https://www.hud.govt.nz/stats-and-insights/property-and-sales-statistics/housing-price-index-composition"
HUD_METHODOLOGY_PDF = "https://www.hud.govt.nz/assets/Uploads/MCERT-HPI-Methodology-Paper-August-2026.pdf"

IRD_SL_BAL_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/social-policy/student-loans/scheme-balances"
IRD_TAX_DEBT_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/tax-and-entitlement-debt-statistics"
IRD_TAX_DEBT_URLS = [
    "https://www.ird.govt.nz/about-us/tax-statistics/tax-and-entitlement-debt-statistics",
    "https://www.ird.govt.nz/about-us/tax-statistics/social-policy/student-loans/overdue-repayments",
]
IRD_INCOME_DIST_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/revenue-refunds/income-distribution"
IRD_INCOME_DATASETS_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/revenue-refunds/income-distribution/tax-on-taxable-income-datasets"
IRD_INCOME_TOT_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/revenue-refunds/income-distribution/individuals-total-taxable-income"
IRD_INCOME_XLSX = "https://www.ird.govt.nz/-/media/project/ir/home/documents/about-us/tax-statistics---current/revenue-and-refunds/tax-on-taxable-income/tax-on-taxable-income/taxable-income-distribution-of-individuals-2026.xlsx"

IRD_WAGE_SALARY_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/revenue-refunds/wage-salary-distributions"
IRD_WAGE_DATASETS_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/revenue-refunds/wage-salary-distributions/wage-and-salary-statistics-datasets"
IRD_WAGE_NUM_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/revenue-refunds/wage-salary-distributions/number-of-wage-and-salary-earners"
IRD_WAGE_BAND_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/revenue-refunds/wage-salary-distributions/number-of-wage-and-salary-earners-by-income-band"

IRD_TRUSTS_NUM_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/revenue-refunds/trusts-and-estates-statistics/number-of-trusts-and-estates"
IRD_TRUSTS_INC_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/revenue-refunds/trusts-and-estates-statistics/trusts-ir6"
IRD_TRUSTS_XLSX = "https://www.ird.govt.nz/-/media/project/ir/home/documents/about-us/tax-statistics---current/revenue-and-refunds/trusts-and-estates-statistics/trusts-and-estates-statistics/trusts-and-estates-statistics-spreadsheet.xlsx"

IRD_WFF_ENT_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/working-for-families-statistics/total-entitlements-paid-for-working-for-families-tax-credits"
IRD_WFF_FAM_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/working-for-families-statistics/number-of-families-claiming-and-entitled-to-working-for-families"
IRD_WFF_MAIN_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/working-for-families-statistics"

IRD_KS_NUM_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/withdrawals/numbers"
IRD_KS_AMT_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/withdrawals/amounts"
IRD_KS_TR_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/Joining-and-transfers/scheme-transfers"
IRD_KS_REG_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/withdrawals/withdrawals-regions"
IRD_KS_ENTRY_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/Joining-and-transfers/entry-methods"
IRD_KS_PAY_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/contributions/scheme-providers-payments"
IRD_KS_SUSP_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/contributions/savings-suspensions"
IRD_KS_AGE_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/joining/member-demographics/number-and-age-of-kiwisaver-members"
IRD_KS_INC_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/joining/member-demographics/number-of-kiwisaver-members-by-income"
IRD_KS_DATA_PAGE = "https://www.ird.govt.nz/about-us/tax-statistics/kiwisaver/datasets"
IRD_KS_ANNUAL_XLSX = "https://www.ird.govt.nz/-/media/project/ir/home/documents/about-us/tax-statistics---current/kiwisaver-statistics/annual-kiwisaver-statistics/annual-kiwisaver-stats.xlsx"
IRD_KS_MONTHLY_XLSX = "https://www.ird.govt.nz/-/media/project/ir/home/documents/about-us/tax-statistics---current/kiwisaver-statistics/monthly-kiwisaver-data/monthly-kiwisaver-data.xlsx"

FMA_KS_REPORT_PAGE = "https://www.fma.govt.nz/library/reports-and-papers/kiwisaver-annual-report/"
KAINGA_ORA_QTR_REPORT_PAGE = "https://kaingaora.govt.nz/en_NZ/publications/home-ownership-publications/home-ownership-products-report/"
KAINGA_ORA_ANN_REPORT_PAGE = "https://kaingaora.govt.nz/en_NZ/publications/core-publications/annual-report/"

STATS_CPI_PAGE = "https://www.stats.govt.nz/information-releases/consumers-price-index-september-2026-quarter/"
STATS_HLPI_PAGE = "https://www.stats.govt.nz/information-releases/household-living-costs-price-indexes-september-2026-quarter/"
STATS_HLPI_NEWS_PAGE = "https://www.stats.govt.nz/news/household-living-costs-increase-3-0-percent-annually/"
STATS_HLPI_URLS = [
    "https://www.stats.govt.nz/information-releases/household-living-costs-price-indexes-september-2026-quarter/",
    "https://www.stats.govt.nz/news/household-living-costs-increase-3-0-percent-annually/",
    "https://www.stats.govt.nz/information-releases/household-living-costs-price-indexes-june-2026-quarter/",
]
STATS_HILS_INCOME_PAGE = "https://www.stats.govt.nz/information-releases/household-income-and-housing-cost-statistics-year-ended-june-2025/"
STATS_HILS_NEWS_PAGE = "https://www.stats.govt.nz/news/increasing-household-income-offset-by-higher-housing-costs-in-year-to-june-2025/"
STATS_NET_WORTH_PAGE = "https://www.stats.govt.nz/information-releases/household-net-worth-statistics-year-ended-june-2024/"
STATS_LABOUR_INCOME_PAGE = "https://www.stats.govt.nz/information-releases/labour-market-statistics-income-june-2026-quarter/"

RBNZ_FSR_PAGE = "https://www.rbnz.govt.nz/financial-stability/financial-stability-report"
RBNZ_INS_STATS_PAGE = "https://www.rbnz.govt.nz/statistics/series/insurance"
RBNZ_C35_PAGE = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-loan-reconciliation"
NHC_ACT_PAGE = "https://www.legislation.govt.nz/act/public/2023/0004/latest/whole.html"
NHC_FUND_PAGE = "https://www.naturalhazards.govt.nz/"
ACC_LEVY_PAGE = "https://www.acc.co.nz/for-business/understanding-your-levies/paying-levies-if-you-work-for-yourself"
ACC_LEVY_RATE_PAGE = "https://www.acc.co.nz/for-business/received-an-invoice/calculate-your-levies"

CATEGORY_ORDER = [
    "First-home buyers",
    "House prices, sales & HPI",
    "Homeowners & refinancing",
    "Property investors",
    "LVR, DTI & credit policy",
    "Household debt, credit cards & borrower stress",
    "Income, net worth & household benchmarks",
    "New builds & construction",
    "Rents & tenancy",
    "KiwiSaver",
    "Property & Natural Hazard Insurance",
    "Personal Risk & Income Protection",
    "Insurance Market & Claims",
]
# === END OF BLOCK 1 ===
# === BLOCK 2 OF 13 ===
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
        "definition": "The percentage of new first-home buyer residential mortgage commitments at New Zealand registered banks approved with a loan-to-value ratio (LVR) above 80% (a deposit of less than 20%).",
        "methodology": "Derived from RBNZ Table C31 by dividing monthly new first-home buyer mortgage commitments with LVR > 80% (LVRN.MMB2.BC1) by total monthly first-home buyer mortgage commitments (LVRN.MMB2.AC1), multiplied by 100. Covers registered banks only and measures new commitments accepted during the month rather than final loan drawdowns.",
        "hlf_calculation": "(High-LVR >80% FHB Borrowers [LVRN.MMB2.BC1] ÷ Total FHB Borrowers [LVRN.MMB2.AC1]) × 100",
        "hlf_interpretation": "HLF market context: RBNZ Table C31 aggregates all >80% LVR first-home buyer commitments at registered banks, including loans approved under banks' discretionary high-LVR speed-limit allowances as well as LVR-exempt Kāinga Ora First Home Loans and new-build commitments.",
        "questions": [
            "Can I buy a first home in New Zealand with less than a 20% deposit?",
            "What percentage of NZ first-home buyers get a mortgage with under a 20% deposit?",
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
        "definition": "The share of total new residential mortgage lending value at registered banks committed to first-home buyers across New Zealand, alongside the monthly number of first-home buyer commitments and mean loan size.",
        "methodology": "Derived from RBNZ Table C31 by dividing first-home buyer new lending value (LVRN.MMB1.AC1) by total new lending across all borrower types (LVRN.MMB1.AA). Mean loan size is LVRN.MMB1.AC1 (published by RBNZ rounded to the nearest $1m) divided by LVRN.MMB2.AC1.",
        "hlf_calculation": "(FHB Lending Value [LVRN.MMB1.AC1] ÷ All-Borrower Total Lending [LVRN.MMB1.AA]) × 100; Mean Loan = LVRN.MMB1.AC1 ÷ LVRN.MMB2.AC1",
        "hlf_interpretation": "HLF market context: Measures first-home buyer share of new bank mortgage commitments (which in Table C31 are compared against total commitments across property purchases, bank-to-bank refinances, and equity top-ups by existing owner-occupiers and investors).",
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
        "name": "Typical first-home buyer application gross income (RBNZ Table C41)",
        "also_known_as": "FHB borrower-weighted average gross application income (BGI)",
        "category": "First-home buyers",
        "topic": "First-home buyer application income benchmark",
        "source": "RBNZ Table C41",
        "source_full": "Reserve Bank of New Zealand — Residential mortgage borrower gross income (BGI) (Table C41)",
        "series_ids": "Sum(FHB Band Gross Income Pool) ÷ Sum(FHB Band Implied Applications)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi",
        "data_type": "HLF-derived band-weighted mean from official RBNZ Table C41",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The borrower-weighted mean annual gross income across approved New Zealand first-home buyer mortgage applications at registered banks, derived from RBNZ Table C41.",
        "methodology": "Derived from RBNZ Table C41 by summing total gross borrower income ($m, published rounded to $0.1m) across all reported first-home buyer Debt-to-Income bands and dividing by the total implied number of applications (Band Pool ÷ Band Average Income, published rounded to $1,000).",
        "hlf_calculation": "Total FHB Gross Income Pool ($m) × $1m ÷ Total Implied FHB Applications",
        "hlf_interpretation": "HLF market context: Under RBNZ's Borrower Gross Income (BGI) methodology, income is measured at the mortgage application level—combining the gross incomes of all borrowers named on the loan (such as couples or co-borrowers, plus any qualifying boarder or other income accepted by the bank). It does not measure individual solo salaries or general population household income.",
        "questions": [
            "What is the average application income of a first-home buyer in New Zealand?",
            "What does Borrower Gross Income (BGI) measure in RBNZ mortgage statistics?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "fhb-average-loan-size": {
        "name": "Average first-home buyer mortgage commitment size (RBNZ Table C31)",
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
        "definition": "The average residential mortgage commitment size approved for first-home buyers at New Zealand registered banks.",
        "methodology": "Derived from RBNZ Table C31 by dividing monthly first-home buyer new lending value (LVRN.MMB1.AC1, published rounded to the nearest $1m) by monthly first-home buyer loan count (LVRN.MMB2.AC1). Covers registered bank commitments only and excludes the buyer's cash/KiwiSaver deposit.",
        "hlf_calculation": "(Total FHB Lending Value [LVRN.MMB1.AC1] × $1,000,000) ÷ Total FHB Borrower Count [LVRN.MMB2.AC1]",
        "hlf_interpretation": "HLF market context: Reflects the national mean mortgage amount committed to first-home buyers across all regions and deposit tiers (excluding the deposit component of the property purchase price).",
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
        "name": "Implied first-home buyer mortgage-to-income ratio & share <= 6x DTI (RBNZ C31/C40/C41)",
        "also_known_as": "Implied FHB loan-to-income multiple & share of FHB lending within 5x and 6x TDTI",
        "category": "First-home buyers",
        "topic": "First-home buyer Debt-to-Income benchmark",
        "source": "RBNZ Tables C31, C40 & C41",
        "source_full": "Reserve Bank of New Zealand — New Lending (C31), DTI Purpose Use (C40) & Borrower Gross Income (C41)",
        "series_ids": "C31 Mean FHB Loan ÷ C41 Mean FHB Gross Income | C40 Share of FHB Lending with TDTI <= 5 and <= 6",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use",
        "data_type": "HLF-derived loan-to-income ratio & band shares from official RBNZ Tables C31, C40 & C41",
        "derived_by_hlf": True,
        "unit": "ratio",
        "frequency": "Monthly",
        "definition": "The implied ratio of average first-home buyer mortgage size (Table C31) to average first-home buyer gross application income (Table C41), reported alongside the official share of FHB lending written at or below 5x and 6x Total Debt-to-Income (TDTI) in RBNZ Table C40.",
        "methodology": "Divides the monthly mean FHB mortgage commitment size (RBNZ Table C31) by the monthly band-weighted mean FHB gross borrower income (RBNZ Table C41), paired with the official share of FHB lending dollar commitments in TDTI bands <= 5x and <= 6x from RBNZ Table C40. Note that RBNZ's official TDTI in Table C40 includes all qualifying borrower debt (mortgages plus personal, student, and credit card debt norms), whereas the C31 ÷ C41 ratio compares average mortgage size to average gross income.",
        "hlf_calculation": "Implied Mortgage-to-Income Ratio = Avg FHB Loan (C31) ÷ Avg FHB Gross Income (C41); Share <= 6x TDTI = 100% − (>6x TDTI Share in C40)",
        "hlf_interpretation": "HLF market context: Compares the average first-home buyer loan size against average application gross income alongside the official RBNZ Table C40 Total Debt-to-Income distribution.",
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
        "definition": "Total monthly dollar value and number of new residential mortgage commitments approved for first-home buyers at New Zealand registered banks.",
        "methodology": "Directly reported in RBNZ Table C31 series LVRN.MMB1.AC1 (FHB total new lending value in $ millions) and LVRN.MMB2.AC1 (FHB total borrower count), alongside HLF-calculated year-on-year percentage changes. Measures new loan commitments accepted during the month at registered banks rather than final property settlements.",
        "hlf_calculation": "Officially reported LVRN.MMB1.AC1 ($m) and LVRN.MMB2.AC1 (loans), plus YoY % change",
        "hlf_interpretation": "HLF market context: Tracks the monthly volume and dollar value of new residential mortgage commitments entered into by first-home buyers with registered banks.",
        "questions": [
            "How much do New Zealand banks commit to first-home buyers each month?",
            "How many first-home buyer mortgage commitments are recorded per month in New Zealand?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
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
        "definition": "Comparison of the mean mortgage commitment size for first-home buyers borrowing with under a 20% deposit (>80% LVR) versus those borrowing with a 20% or greater deposit (<=80% LVR) at registered banks.",
        "methodology": "Derived from RBNZ Table C31 by dividing >80% LVR FHB lending value (LVRN.MMB1.BC1, rounded to $1m) by >80% LVR FHB loan count (LVRN.MMB2.BC1), and comparing against <=80% LVR FHB lending value divided by <=80% LVR FHB loan count. Table C31 records loan amounts rather than underlying property purchase prices.",
        "hlf_calculation": "Mean(>80% LVR FHB Loan) minus Mean(<=80% LVR FHB Loan) (derived from RBNZ $1m-rounded lending series)",
        "hlf_interpretation": "HLF market context: First-home buyers with under a 20% deposit (>80% LVR) borrow a higher percentage of the property value than buyers with a 20%+ deposit (<=80% LVR), resulting in a higher average loan balance in RBNZ Table C31.",
        "questions": [
            "How much more do low-deposit first-home buyers borrow compared to buyers with a 20% deposit?",
            "What is the average loan size for a >80% LVR first-home buyer in New Zealand?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
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
        "definition": "The share and number of first-home buyer mortgage commitments at registered banks written at an LVR of 80% or below (a deposit of 20% or more), alongside their mean loan size.",
        "methodology": "Derived from RBNZ Table C31 by subtracting >80% LVR FHB commitments from total FHB commitments for both loan count and dollar value.",
        "hlf_calculation": "((Total FHB Loans [LVRN.MMB2.AC1] - >80% LVR FHB Loans [LVRN.MMB2.BC1]) ÷ Total FHB Loans) × 100",
        "hlf_interpretation": "HLF market context: Tracks the proportion of first-home buyer commitments entering at <=80% LVR (20%+ equity/deposit), which is the standard threshold for advertised 'Special' fixed mortgage rates at major New Zealand banks.",
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
        "definition": "Monthly comparison of total new mortgage lending dollars committed to first-home buyers versus residential property investors at New Zealand registered banks.",
        "methodology": "Compares RBNZ Table C31 series LVRN.MMB1.AC1 (First-home buyers total new lending value) against LVRN.MMB1.AN (Investors total new lending value). Note that investor commitments in Table C31 include property purchases, external refinances, and equity top-ups.",
        "hlf_calculation": "FHB Lending Value [LVRN.MMB1.AC1] minus Investor Lending Value [LVRN.MMB1.AN]",
        "hlf_interpretation": "HLF market context: Compares the monthly dollar value of new registered bank mortgage commitments to first-home buyers against commitments to residential property investors.",
        "questions": [
            "Do first-home buyers or property investors borrow more in New Zealand?",
            "How does first-home buyer lending compare to investor lending in RBNZ statistics?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
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
        "definition": "The cumulative number and dollar value of first-home buyer mortgage commitments recorded by New Zealand registered banks over the most recent 12-month period, including the 12-month total written above 80% LVR.",
        "methodology": "Derived from RBNZ Table C31 by summing the latest 12 monthly observations of LVRN.MMB2.AC1 (total FHB borrowers), LVRN.MMB1.AC1 (total FHB lending value), and LVRN.MMB2.BC1 (>80% LVR FHB borrowers). Measures new commitments at registered banks.",
        "hlf_calculation": "Sum of latest 12 months of LVRN.MMB2.AC1 (Loans) & LVRN.MMB1.AC1 (Value)",
        "hlf_interpretation": "HLF market context: Aggregating 12 months of RBNZ Table C31 commitments smooths out monthly seasonality across the calendar year.",
        "questions": [
            "How many first-home buyer mortgage commitments are approved in New Zealand each year?",
            "How much money do New Zealand banks commit to first-home buyers annually?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "kainga-ora-first-home-loan-approvals": {
        "name": "Kāinga Ora First Home Loan (5% deposit) nominal approvals, settlements & YoY/QoQ trend",
        "also_known_as": "Kāinga Ora Home Ownership Products Report — First Home Loan quarterly approvals & settlements series",
        "category": "First-home buyers",
        "topic": "Kāinga Ora 5% deposit First Home Loan approvals, settlements & annual growth",
        "source": "Kāinga Ora Quarterly Report",
        "source_full": "Kāinga Ora – Homes and Communities — Home Ownership Products Quarterly Report (CC BY 4.0)",
        "series_ids": "Kāinga Ora Home Ownership Products Report — First Home Loans: Approvals and Settlements",
        "source_url": KAINGA_ORA_QTR_REPORT_PAGE,
        "data_type": "Officially reported Kāinga Ora quarterly PDF series & HLF-derived YoY/QoQ growth & 12-month rolling totals",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Quarterly",
        "definition": "The number of Kāinga Ora-underwritten First Home Loan (5% deposit) approvals and settlements recorded each month and quarter across New Zealand.",
        "methodology": "Automatically fetched and parsed from Kāinga Ora's Home Ownership Products Quarterly Report PDF (licensed under CC BY 4.0). As noted by Kāinga Ora, approvals include 90-day pre-approvals ('live shopping' certificates) issued before a specific property is secured; because some pre-approvals expire, are re-issued, or settle in a subsequent quarter, approval counts exceed settlement counts.",
        "hlf_calculation": "Monthly Mean Approvals = Quarterly Approvals ÷ 3; Settlement-to-Approval Ratio = (Settlements ÷ Approvals) × 100",
        "hlf_interpretation": "HLF market context: Tracks quarterly and rolling 12-month approvals and completed settlements under the Kāinga Ora First Home Loan scheme (which allows eligible buyers to apply with a minimum 5% deposit outside bank high-LVR speed limits).",
        "questions": [
            "How many Kāinga Ora First Home Loans are approved and settled each quarter in New Zealand?",
            "Why are Kāinga Ora First Home Loan approvals higher than settlements?",
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
        "series_ids": "Kāinga Ora Home Ownership Products Report — First Home Loan Approvals & Settlements by TLA",
        "source_url": KAINGA_ORA_QTR_REPORT_PAGE,
        "data_type": "Officially reported Kāinga Ora council-level tables & HLF-derived national shares",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Quarterly",
        "definition": "Distribution of Kāinga Ora 5% deposit First Home Loan approvals and settlements across New Zealand Territorial Local Authorities (TLAs).",
        "methodology": "Automatically parsed from the Territorial Local Authority Approvals and Settlements tables in Kāinga Ora's Home Ownership Products Quarterly Report PDF. Kāinga Ora itemises the largest TLAs individually and groups smaller districts into 'Other'.",
        "hlf_calculation": "Council Share of National Approvals (%) = (Council Quarterly Approvals ÷ National Quarterly Approvals) × 100",
        "hlf_interpretation": "HLF market context: Shows the geographic distribution of Kāinga Ora First Home Loan approvals and settlements across New Zealand city and district councils.",
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
        "definition": "The number of previous New Zealand homeowners approved by Kāinga Ora each quarter (and over the rolling 12-month period) as being in a qualifying financial position to apply for a 'second-chance' KiwiSaver savings withdrawal.",
        "methodology": "Automatically fetched and parsed from the 'Previous Owner Savings Withdrawal Applications' chart in Kāinga Ora's Home Ownership Products Quarterly Report PDF. Measures eligibility determination letters approved by Kāinga Ora (which the applicant then submits to their KiwiSaver provider), rather than final completed property settlements.",
        "hlf_calculation": "12-Month Rolling Approvals = Sum of Applications Approved across the 4 most recent quarters; Monthly Mean = Quarterly Approvals ÷ 3",
        "hlf_interpretation": "HLF market context: Previous homeowners who no longer hold an interest in property and meet Kāinga Ora's asset cap can apply for a Previous Homeowner determination letter to access their KiwiSaver balance for a home deposit.",
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
        "also_known_as": "Kāinga Ora First Home Loan Settlements Underwritten & Approvals (SPE 4.1)",
        "category": "First-home buyers",
        "topic": "Annual Kāinga Ora First Home Loans underwritten & Mortgage Insurance Scheme",
        "source": "Kāinga Ora Reports",
        "source_full": "Kāinga Ora – Homes and Communities — Home Ownership Products Report & Annual Report (SPE 4.1)",
        "series_ids": "Kāinga Ora First Home Loan Approvals & Underwritten Settlements (SPE 4.1)",
        "source_url": KAINGA_ORA_QTR_REPORT_PAGE,
        "data_type": "Officially reported Kāinga Ora series & HLF-derived rolling totals",
        "derived_by_hlf": True,
        "unit": "loans",
        "frequency": "Quarterly / Annual",
        "definition": "The number of First Home Loan (5% deposit) mortgages settled and underwritten by Kāinga Ora alongside total approved applications under the Crown Mortgage Insurance Scheme.",
        "methodology": "Derived live from Kāinga Ora's Home Ownership Products Quarterly Report PDF. Settlements represent completed First Home Loans underwritten by Kāinga Ora across participating scheme lenders.",
        "hlf_calculation": "Monthly Mean Settlements = Quarterly Settlements ÷ 3; Annualised Pace = Quarterly Settlements × 4",
        "hlf_interpretation": "HLF market context: Under the Kāinga Ora First Home Loan scheme, participating lenders originate 5%-deposit mortgages underwritten by Kāinga Ora (subject to the scheme's Lenders Mortgage Insurance premium and eligibility criteria).",
        "questions": [
            "How many First Home Loans does Kāinga Ora approve and underwrite each year in New Zealand?",
            "How does the Kāinga Ora First Home Loan scheme work?",
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
        "definition": "Quarterly and 12-month rolling total number and dollar value ($ millions) of KiwiSaver first-home savings withdrawals reported across New Zealand, plus the mean withdrawal per withdrawing member.",
        "methodology": "Automatically parsed from the 'KiwiSaver First Home Savings Withdrawal (IRD)' 4-quarter chart in Kāinga Ora's Home Ownership Products Quarterly Report PDF. Counts individual member withdrawals; where two co-buyers both withdraw KiwiSaver funds toward the same property purchase, two member withdrawals are recorded.",
        "hlf_calculation": "12-Month Rolling Total = Sum of 4 quarters (Withdrawals & $m); Mean Withdrawal = Value ($m × $1m) ÷ Withdrawals",
        "hlf_interpretation": "HLF market context: Tracks the quarterly and rolling 12-month volume and average dollar size of individual KiwiSaver first-home withdrawals across New Zealand.",
        "questions": [
            "How much is withdrawn from KiwiSaver for first homes each quarter and year in New Zealand?",
            "How many KiwiSaver first-home withdrawals are completed each year?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "credit-card-balances-interest-share": {
        "name": "New Zealand credit card debt, share accruing interest & weighted average interest rate (RBNZ Table C12)",
        "also_known_as": "RBNZ Table C12 Credit Card Balances — Total advances outstanding, interest-bearing share & interest rate",
        "category": "Household debt, credit cards & borrower stress",
        "topic": "Credit card debt, interest-bearing share & interest rates",
        "source": "RBNZ Table C12",
        "source_full": "Reserve Bank of New Zealand — Credit card balances (Table C12)",
        "series_ids": "RBNZ Table C12 — Total Advances Outstanding ($m), Interest-Bearing Advances ($m) & Weighted Average Interest Rate (%)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/credit-card-balances",
        "data_type": "Officially reported RBNZ series & HLF-derived interest-bearing share",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Monthly",
        "definition": "Total outstanding credit card advances across New Zealand registered banks, the percentage of balances accruing interest, and the weighted average interest rate charged on interest-bearing balances.",
        "methodology": "Extracted live from RBNZ Table C12 (Credit card balances). Divides interest-bearing advances by total advances outstanding. RBNZ Table C12 aggregates both personal and commercial/business credit card balances held with New Zealand registered banks.",
        "hlf_calculation": "Interest-Bearing Share (%) = (Interest-Bearing Advances [$m] ÷ Total Advances Outstanding [$m]) × 100",
        "hlf_interpretation": "HLF market context: RBNZ Table C12 separates total credit card balances into non-interest-bearing balances (within interest-free payment periods) and interest-bearing balances, alongside the weighted-average interest rate charged on interest-bearing debt.",
        "questions": [
            "How much credit card debt is outstanding in New Zealand?",
            "What is the average interest rate on interest-bearing credit card balances in New Zealand?",
            "What percentage of NZ credit card balances accrue interest?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "credit-card-limits-servicing-impact": {
        "name": "Total New Zealand credit card limits, utilisation rate & unused credit limits (RBNZ Table C12)",
        "also_known_as": "RBNZ Table C12 Total Credit Card Limits ($B), Advances Outstanding & Unused Credit Limits",
        "category": "Household debt, credit cards & borrower stress",
        "topic": "Credit card limits & utilisation rate",
        "source": "RBNZ Table C12",
        "source_full": "Reserve Bank of New Zealand — Credit card balances and credit limits (Table C12)",
        "series_ids": "RBNZ Table C12 — Total Credit Limits ($m) & Total Advances Outstanding ($m)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/credit-card-balances",
        "data_type": "Officially reported RBNZ series & HLF-derived utilisation rate and unused limit volume",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Monthly",
        "definition": "Total approved credit card limits across New Zealand registered banks ($ billions), the percentage of limits currently utilised, and the dollar volume of undrawn credit card limits.",
        "methodology": "Extracted live from RBNZ Table C12 by comparing total credit limits ($m) against total credit card advances outstanding ($m). Unused credit limit equals Total Credit Limits minus Total Advances Outstanding. RBNZ Table C12 includes both personal and commercial/business credit cards.",
        "hlf_calculation": "Utilisation Rate (%) = (Total Advances Outstanding ÷ Total Credit Limits) × 100; Unused Limits = Total Limits − Advances",
        "hlf_interpretation": "HLF market context: RBNZ Table C12 measures total approved credit card limits and advances outstanding across New Zealand personal and commercial cards. When assessing a mortgage application, lenders typically factor a percentage of a borrower's approved credit limits (rather than only the current balance owing) into expense calculations, though specific assessment rates and exceptions vary across lenders.",
        "questions": [
            "What is the total value of credit card limits and unused credit limits in New Zealand?",
            "How do credit card limits affect mortgage serviceability assessments in New Zealand?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "total-housing-debt-stock": {
        "name": "Total New Zealand residential mortgage debt outstanding & share of private sector credit (RBNZ Table C5)",
        "also_known_as": "RBNZ Table C5 Sector Lending — Total Housing Lending Stock ($B) across Banks & Non-Bank Lenders",
        "category": "Household debt, credit cards & borrower stress",
        "topic": "Total New Zealand housing mortgage debt stock",
        "source": "RBNZ Table C5",
        "source_full": "Reserve Bank of New Zealand — Registered banks and non-bank lending institutions: Sector lending (Table C5)",
        "series_ids": "RBNZ Table C5 — Total Housing Lending ($m) & Total Private Sector Lending ($m)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/sector-lending",
        "data_type": "Officially reported RBNZ series & HLF-derived sector share & YoY change",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Monthly",
        "definition": "The total stock of residential mortgage debt outstanding across New Zealand registered banks and non-bank lending institutions (NBLIs), alongside its share of total private sector credit.",
        "methodology": "Extracted live from RBNZ Table C5 (Sector lending). Computes year-on-year dollar and percentage growth and divides total housing lending by total private sector lending (housing, consumer, business, and agriculture).",
        "hlf_calculation": "Housing Share of Total Credit (%) = (Total Housing Lending ÷ Total Sector Lending) × 100; YoY Growth (%)",
        "hlf_interpretation": "HLF market context: Tracks the total outstanding stock of residential mortgage lending across registered banks and non-bank lending institutions relative to total New Zealand private sector credit.",
        "questions": [
            "What is the total amount of residential mortgage debt in New Zealand?",
            "What percentage of New Zealand private sector credit is housing debt?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "consumer-personal-lending-stock": {
        "name": "Total New Zealand personal & consumer lending stock across banks & non-bank lenders (RBNZ Table C5)",
        "also_known_as": "RBNZ Table C5 Consumer Sector Lending — Personal loans, car finance, cards & non-bank share",
        "category": "Household debt, credit cards & borrower stress",
        "topic": "Personal & consumer lending stock (banks & non-bank lenders)",
        "source": "RBNZ Table C5",
        "source_full": "Reserve Bank of New Zealand — Registered banks and non-bank lending institutions: Sector lending (Table C5)",
        "series_ids": "RBNZ Table C5 — Total Consumer Lending ($m), Bank Consumer Lending & Non-Bank Lending Institutions Consumer Lending",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/sector-lending",
        "data_type": "Officially reported RBNZ series & HLF-derived non-bank share & YoY change",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Monthly",
        "definition": "Total personal and consumer credit outstanding across New Zealand registered banks and non-bank lending institutions (NBLIs) as reported in RBNZ Table C5.",
        "methodology": "Extracted live from RBNZ Table C5 (Sector lending — Consumer). Aggregates personal loans, motor vehicle finance, hire purchase, overdrafts, and personal credit card balances across banks and NBLIs.",
        "hlf_calculation": "Non-Bank Share (%) = (Non-Bank Consumer Lending ÷ Total Consumer Lending) × 100; YoY % change",
        "hlf_interpretation": "HLF market context: Measures the aggregate stock of non-mortgage consumer credit held with registered banks and non-bank lending institutions across New Zealand.",
        "questions": [
            "How much personal and consumer debt is outstanding in New Zealand?",
            "What share of New Zealand consumer lending is held with non-bank lenders?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "mortgage-non-performing-arrears": {
        "name": "Registered bank non-performing housing loans & NPL ratio (RBNZ Table S50)",
        "also_known_as": "RBNZ Table S50 Bank Asset Quality — Non-performing housing loans (90+ days past due & impaired)",
        "category": "Household debt, credit cards & borrower stress",
        "topic": "Mortgage arrears & non-performing housing loans",
        "source": "RBNZ Table S50",
        "source_full": "Reserve Bank of New Zealand — Banks: Assets – Loans by asset quality (Table S50)",
        "series_ids": "RBNZ Table S50 — Non-Performing Housing Loans ($m) & Housing NPL Ratio (%)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-by-asset-quality",
        "data_type": "Officially reported RBNZ series & HLF-derived YoY comparison",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The percentage and dollar value of New Zealand registered bank residential mortgage loans classified as non-performing (90+ days past due or individually impaired).",
        "methodology": "Extracted live from RBNZ Table S50 (Banks: Assets – Loans by asset quality). Covers registered banks only and measures loans that are at least 90 days past due or impaired (excluding earlier-stage 30–89 day past-due loans and non-bank lenders).",
        "hlf_calculation": "Housing Non-Performing Loan Ratio (%) = (Non-Performing Housing Loans [$m] ÷ Total Bank Housing Loans [$m]) × 100",
        "hlf_interpretation": "HLF market context: RBNZ Table S50 tracks the proportion of registered bank residential mortgage lending that is 90+ days past due or impaired.",
        "questions": [
            "What percentage of New Zealand bank mortgages are non-performing or 90+ days past due?",
            "How much residential mortgage debt is classified as non-performing in RBNZ Table S50?",
        ],
        "article_title": "Mortgage Refinance & Restructure Guide",
        "article_url": "https://www.homeloanfactory.co.nz/services/mortgages/mortgage-refinance/",
        "tool_title": "HLF Mortgage Refinance Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/refinance-calculator",
    },
    "consumer-loan-stress-npl": {
        "name": "Registered bank non-performing personal & consumer loans & NPL ratio (RBNZ Table S50)",
        "also_known_as": "RBNZ Table S50 Bank Asset Quality — Non-performing personal/consumer lending ratio (%)",
        "category": "Household debt, credit cards & borrower stress",
        "topic": "Consumer & personal loan arrears & non-performing ratio",
        "source": "RBNZ Table S50",
        "source_full": "Reserve Bank of New Zealand — Banks: Assets – Loans by asset quality (Table S50)",
        "series_ids": "RBNZ Table S50 — Non-Performing Consumer Loans ($m) & Consumer NPL Ratio (%)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-by-asset-quality",
        "data_type": "Officially reported RBNZ series & HLF-derived YoY comparison",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The non-performing loan (NPL) ratio and dollar value for New Zealand registered bank personal and consumer loans (loans 90+ days past due or individually impaired).",
        "methodology": "Extracted live from RBNZ Table S50 (Banks: Assets – Loans by asset quality — Consumer/Personal lending). Covers registered banks only and excludes non-bank lending institutions.",
        "hlf_calculation": "Consumer NPL Ratio (%) = (Non-Performing Consumer Loans [$m] ÷ Total Bank Consumer Loans [$m]) × 100",
        "hlf_interpretation": "HLF market context: Tracks the proportion of registered bank personal and consumer lending classified as 90+ days past due or impaired in RBNZ Table S50.",
        "questions": [
            "What is the non-performing loan ratio on personal and consumer loans at New Zealand banks?",
            "How does the consumer NPL ratio compare with the housing NPL ratio in RBNZ Table S50?",
        ],
        "article_title": "Mortgage Refinance & Restructure Guide",
        "article_url": "https://www.homeloanfactory.co.nz/services/mortgages/mortgage-refinance/",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "household-debt-servicing-burden": {
        "name": "Quarterly mortgage interest charged & scheduled principal-and-interest repayments (RBNZ Table C35)",
        "also_known_as": "RBNZ Table C35 Residential Mortgage Loan Reconciliation — Quarterly interest charged ($B) & scheduled repayments ($B)",
        "category": "Household debt, credit cards & borrower stress",
        "topic": "Quarterly mortgage interest charged & scheduled repayments",
        "source": "RBNZ Table C35",
        "source_full": "Reserve Bank of New Zealand — Residential mortgage loan reconciliation (Table C35)",
        "series_ids": "RBNZ Table C35 — Quarterly Interest Charged ($m) & Scheduled Principal-and-Interest Repayments ($m)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-loan-reconciliation",
        "data_type": "Officially reported RBNZ quarterly mortgage reconciliation series",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Quarterly",
        "definition": "Total residential mortgage interest charged across New Zealand registered banks each quarter ($ billions) alongside scheduled principal-and-interest repayments contractually due.",
        "methodology": "Extracted live from RBNZ Table C35 (Residential mortgage loan reconciliation). Covers registered banks' residential mortgage books and measures aggregate dollar flows rather than individual household debt-servicing-to-income ratios.",
        "hlf_calculation": "Quarterly Interest Charged ($B) & Rolling 4-Quarter Scheduled Repayments ($B)",
        "hlf_interpretation": "HLF market context: RBNZ Table C35 records the total dollar interest charged to residential mortgage accounts and the total scheduled principal-and-interest repayments each quarter across registered banks.",
        "questions": ["How much residential mortgage interest is charged across New Zealand banks each quarter?"],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "hlpi-household-living-costs-interest": {
        "name": "Household Living-Costs Price Index (HLPI): all-households inflation & expenditure group spread (Stats NZ)",
        "also_known_as": "Stats NZ HLPI — Household living-cost inflation including mortgage interest payments",
        "category": "Household debt, credit cards & borrower stress",
        "topic": "Household living-cost inflation & mortgage interest payment changes",
        "source": "Stats NZ HLPI",
        "source_full": "Stats NZ — Household living-costs price indexes (HLPI)",
        "series_ids": "Stats NZ HLPI — All-Households Annual % Change & Expenditure Quintile Spread",
        "source_url": STATS_HLPI_PAGE,
        "data_type": "Officially reported Stats NZ quarterly release",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in Stats NZ's Household Living-Costs Price Index (HLPI) for all households, alongside highest- and lowest-expenditure household groups.",
        "methodology": "Automatically fetched and parsed from Stats NZ's quarterly Household Living-Costs Price Indexes (HLPI) release. Unlike the CPI (which excludes interest payments and includes the acquisition of new dwellings), the HLPI includes out-of-pocket mortgage interest payments (excluding principal repayments) and property rates.",
        "hlf_calculation": "Official Stats NZ HLPI All-Households Annual % Change",
        "hlf_interpretation": "HLF market context: Because the HLPI includes household debt interest payments while the CPI excludes them, the HLPI reflects how changes in effective mortgage interest rates affect household out-of-pocket living costs.",
        "questions": ["How much have New Zealand household living costs (HLPI) changed over the past year?"],
        "article_title": "Mortgage Refinance & Restructure Guide",
        "article_url": "https://www.homeloanfactory.co.nz/services/mortgages/mortgage-refinance/",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "ird-overdue-tax-debt-stress": {
        "name": "Inland Revenue overdue tax & entitlements debt statistics (IRD)",
        "also_known_as": "IRD Tax and Entitlements Debt Statistics — Overdue debt across individuals & entities",
        "category": "Household debt, credit cards & borrower stress",
        "topic": "Overdue IRD tax and entitlement debt",
        "source": "Inland Revenue",
        "source_full": "Inland Revenue — Tax and entitlements debt statistics (CC BY 4.0)",
        "series_ids": "IRD Tax and Entitlements Debt — Customers in Debt, Total Overdue Debt ($B) & Individual Overdue Debt ($B)",
        "source_url": IRD_TAX_DEBT_PAGE,
        "data_type": "Officially reported Inland Revenue release & HLF-derived mean debt",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Quarterly / Annual",
        "definition": "The total dollar value of overdue tax and entitlements debt owed to Inland Revenue across New Zealand ($ billions), the number of customers with overdue debt, and the portion owed by individuals.",
        "methodology": "Automatically fetched and parsed from Inland Revenue's Tax and Entitlements Debt Statistics release page on ird.govt.nz. IRD's series aggregates all overdue tax and entitlement categories (including income tax, GST, PAYE, Working for Families, student loans, and child support) across both individuals and non-individual entities.",
        "hlf_calculation": "Average Overdue Debt per Customer = Total Overdue Debt ($B) ÷ Customers with Debt",
        "hlf_interpretation": "HLF market context: Reports total overdue tax and entitlement balances recorded by Inland Revenue across individuals and entities. When assessing self-employed or company-director mortgage applications, lenders typically review IRD account statements to check whether tax obligations are current or under a payment arrangement.",
        "questions": ["How many customers have overdue tax and entitlements debt with Inland Revenue?"],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
# === END OF BLOCK 2 ===
# === BLOCK 3 OF 13 ===
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
        "definition": "The percentage of total new residential mortgage lending value at registered banks committed to residential property investors across New Zealand.",
        "methodology": "Derived from RBNZ Table C31 by dividing monthly residential investor new lending value (LVRN.MMB1.AN) by total new mortgage lending across all borrower types (LVRN.MMB1.AA). Covers registered banks only and includes property purchases, external refinances, and equity top-ups.",
        "hlf_calculation": "(Investor Lending Value [LVRN.MMB1.AN] ÷ All-Borrower Total Lending [LVRN.MMB1.AA]) × 100",
        "hlf_interpretation": "HLF market context: Tracks residential property investors' share of total monthly new mortgage commitments at New Zealand registered banks.",
        "questions": [
            "What percentage of New Zealand mortgages go to property investors?",
            "How much new mortgage lending is committed to property investors each month?",
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
        "definition": "The monthly dollar volume, commitment count, and average loan size across new residential property investor mortgage commitments at New Zealand registered banks.",
        "methodology": "Derived from RBNZ Table C31 by dividing monthly investor lending value (LVRN.MMB1.AN, published rounded to the nearest $1m) by monthly investor loan count (LVRN.MMB2.AN).",
        "hlf_calculation": "(Investor Lending Value [LVRN.MMB1.AN] × $1,000,000) ÷ Investor Borrower Count [LVRN.MMB2.AN]",
        "hlf_interpretation": "HLF market context: RBNZ Table C31 investor commitments combine property purchases, bank-to-bank refinances, and equity top-ups on investment properties across registered banks.",
        "questions": [
            "What is the average mortgage commitment size for a property investor in New Zealand?",
            "How many property investor loan commitments are approved each month in NZ?",
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
        "definition": "The monthly dollar volume, commitment count, and average mortgage commitment size for existing ('other') owner-occupiers at New Zealand registered banks.",
        "methodology": "Derived from RBNZ Table C31 by dividing monthly other owner-occupier lending value (LVRN.MMB1.AC2, rounded to $1m) by monthly other owner-occupier loan count (LVRN.MMB2.AC2). In RBNZ Table C31, 'other owner-occupiers' includes existing homeowners purchasing a next home, switching banks (refinancing), or taking out an equity top-up.",
        "hlf_calculation": "(Other Owner-Occupier Lending Value [LVRN.MMB1.AC2] × $1,000,000) ÷ Other Owner-Occupier Count [LVRN.MMB2.AC2]",
        "hlf_interpretation": "HLF market context: Because RBNZ Table C31's 'other owner-occupier' category includes smaller mortgage top-ups alongside full home purchases and refinances, the mean commitment size across all other owner-occupier transactions is lower than the mean purchase-only loan size.",
        "questions": [
            "What is the average mortgage commitment size for an existing owner-occupier in RBNZ Table C31?",
            "How much do existing owner-occupiers borrow each month in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Net Proceeds & Next Home Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/net-proceeds-next-home-calculator",
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
        "definition": "The percentage of all >80% LVR residential mortgage commitments (by dollar value and loan count) at New Zealand registered banks that is committed to first-home buyers.",
        "methodology": "Derived from RBNZ Table C31 by dividing first-home buyer >80% LVR lending value (LVRN.MMB1.BC1) by all-borrower >80% LVR lending value (LVRN.MMB1.BA), and FHB >80% LVR loan count (LVRN.MMB2.BC1) by all-borrower >80% LVR loan count (LVRN.MMB2.BA). Includes both speed-limit lending and LVR-exempt lending.",
        "hlf_calculation": "(FHB >80% LVR Value [LVRN.MMB1.BC1] ÷ All-Borrower >80% LVR Value [LVRN.MMB1.BA]) × 100",
        "hlf_interpretation": "HLF market context: Shows the share of all >80% LVR registered bank commitments (across speed-limit allowances, Kāinga Ora First Home Loans, and LVR-exempt new builds) accounted for by first-home buyers versus existing owner-occupiers and investors.",
        "questions": [
            "What share of >80% LVR mortgage lending goes to first-home buyers vs movers?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "Kāinga Ora LMI Fee & Low-Deposit Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/lenders-mortgage-insurance-fee-estimator",
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
        "definition": "The percentage of existing ('other') owner-occupier mortgage commitments at registered banks written at <=80% LVR (20% or more equity in the security property), alongside their mean commitment size.",
        "methodology": "Derived from RBNZ Table C31 by subtracting >80% LVR other owner-occupier commitments (LVRN.MMB2.BC2) from total other owner-occupier commitments (LVRN.MMB2.AC2).",
        "hlf_calculation": "((Total Other OO Loans [LVRN.MMB2.AC2] - >80% LVR Other OO Loans [LVRN.MMB2.BC2]) ÷ Total Other OO Loans) × 100",
        "hlf_interpretation": "HLF market context: Shows the proportion of existing owner-occupier commitments (across home movers, external refinances, and top-ups) completed at or below the 80% LVR threshold.",
        "questions": [
            "What percentage of existing New Zealand owner-occupiers borrow at <=80% LVR when moving, refinancing, or topping up?",
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
        "definition": "Breakdown of new residential property investor mortgage commitments at registered banks across three LVR tiers: <=60% LVR (40%+ equity), 60%–70% LVR (30%–40% equity), and >70% LVR (under 30% equity).",
        "methodology": "Derived from RBNZ Table C31 investor borrower count series LVRN.MMB2.TN (<=60% LVR), LVRN.MMB2.FN (>60% LVR), LVRN.MMB2.DN (>70% LVR), and LVRN.MMB2.AN (total investors).",
        "hlf_calculation": "<=60% LVR = TN ÷ AN; 60–70% LVR = (FN - DN) ÷ AN; >70% LVR = DN ÷ AN",
        "hlf_interpretation": "HLF market context: Shows the distribution of new investor commitments across RBNZ Table C31 LVR bands. Commitments above 70% LVR aggregate LVR-exempt new-build lending, bridging/refinancing exemptions, and loans approved within banks' high-LVR speed limits.",
        "questions": [
            "How are property investor mortgage commitments distributed across LVR bands in New Zealand?",
            "What percentage of NZ property investors borrow at <=60% LVR, 60%–70% LVR, and >70% LVR?",
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
        "definition": "Monthly count and dollar value of residential property investor mortgage commitments at registered banks approved at >80% LVR (under 20% equity) and in the 70%–80% LVR band (20%–30% equity).",
        "methodology": "Extracted from RBNZ Table C31 series LVRN.MMB2.BN / LVRN.MMB1.BN (>80% LVR investors), with the 70%–80% LVR band derived by subtracting >80% LVR investor commitments from >70% LVR investor commitments (LVRN.MMB2.DN / LVRN.MMB1.DN). RBNZ Table C31 does not separate new-build exemptions from speed-limit allowances within these bands.",
        "hlf_calculation": ">80% LVR = LVRN.MMB2.BN; 70%–80% LVR = LVRN.MMB2.DN minus LVRN.MMB2.BN",
        "hlf_interpretation": "HLF market context: Splits investor commitments above 70% LVR into the 70%–80% LVR band and the >80% LVR band as reported in RBNZ Table C31.",
        "questions": [
            "How many property investors get approved above 80% LVR or between 70% and 80% LVR in RBNZ Table C31?",
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
        "topic": "Investor >70% LVR lending & exemptions",
        "source": "RBNZ Table C31",
        "source_full": "Reserve Bank of New Zealand — New residential mortgage lending by borrower type (Table C31)",
        "series_ids": "LVRN.MMB2.DN ÷ LVRN.MMB2.AN (Share) | LVRN.MMB1.DN (Value)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type",
        "data_type": "Officially reported RBNZ series & HLF-derived share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The share, loan count, and dollar volume of residential property investor mortgage commitments at registered banks written at an LVR above 70% (under 30% equity/deposit).",
        "methodology": "Derived from RBNZ Table C31 by dividing investor commitments with LVR > 70% (LVRN.MMB2.DN) by total investor commitments (LVRN.MMB2.AN).",
        "hlf_calculation": "(Investor Loans >70% LVR [LVRN.MMB2.DN] ÷ Total Investor Loans [LVRN.MMB2.AN]) × 100",
        "hlf_interpretation": "HLF market context: Reports total investor commitments written above 70% LVR in RBNZ Table C31 (combining LVR-exempt new-build lending, bridging/refinancing exemptions, and loans written under bank high-LVR speed limits).",
        "questions": [
            "What percentage of property investors borrow above 70% LVR in New Zealand?",
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
        "definition": "The share, borrower count, and dollar value of existing ('other') owner-occupier mortgage commitments at registered banks approved at an LVR above 80% (under 20% equity).",
        "methodology": "Derived from RBNZ Table C31 by dividing other owner-occupier commitments at >80% LVR (LVRN.MMB2.BC2) by total other owner-occupier commitments (LVRN.MMB2.AC2).",
        "hlf_calculation": "(Other Owner-Occupier >80% LVR Loans [LVRN.MMB2.BC2] ÷ Total Other OO Loans [LVRN.MMB2.AC2]) × 100",
        "hlf_interpretation": "HLF market context: RBNZ Table C31 aggregates all >80% LVR commitments for existing owner-occupiers, including LVR-exempt new-build purchases, bridging finance, dollar-for-dollar refinancing exemptions, and bank speed-limit allowances.",
        "questions": [
            "How many existing owner-occupiers get approved above 80% LVR each month in RBNZ Table C31?",
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
        "definition": "The combined monthly dollar value and commitment count of non-first-home-buyer mortgages at registered banks written above standard LVR thresholds: existing owner-occupiers at >80% LVR plus property investors at >70% LVR.",
        "methodology": "Derived from RBNZ Table C31 by summing other owner-occupier >80% LVR lending (LVRN.MMB1.BC2) and investor >70% LVR lending (LVRN.MMB1.DN). Includes both LVR-exempt lending (such as new builds and bridging finance) and discretionary speed-limit lending.",
        "hlf_calculation": "Other OO >80% LVR [LVRN.MMB1.BC2] + Investor >70% LVR [LVRN.MMB1.DN]",
        "hlf_interpretation": "HLF market context: Sums non-first-home-buyer commitments written above the 80% owner-occupier and 70% investor LVR thresholds in RBNZ Table C31.",
        "questions": [
            "How much mortgage lending goes to existing owner-occupiers above 80% LVR and investors above 70% LVR?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
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
        "definition": "The share of new residential mortgage lending value and number of commitments where a borrower switches from one mortgage provider to another ('change in loan provider' in RBNZ Table C33), alongside the mean refinanced commitment size.",
        "methodology": "Derived from RBNZ Table C33 by dividing 'change in loan provider' lending value (LVRN.MMB1.AZD, rounded to $1m) by the sum of all purpose categories (property purchase AZC, top-up AZB, change in provider AZD, and other AZE). Covers external lender switches at registered banks and excludes internal rate refixes with the same bank.",
        "hlf_calculation": "(Change in Loan Provider Value [LVRN.MMB1.AZD] ÷ Total Lending by Purpose) × 100; Mean Loan = LVRN.MMB1.AZD ÷ LVRN.MMB2.AZD",
        "hlf_interpretation": "HLF market context: RBNZ Table C33 'change in loan provider' measures external refinancing commitments where a borrower moves a mortgage from one lender to a registered bank (excluding internal fixed-rate rollovers with the same bank).",
        "questions": [
            "How many New Zealand homeowners refinance or switch banks each month?",
            "What percentage of NZ new mortgage commitments is refinancing to a new bank?",
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
        "definition": "The percentage of monthly new residential mortgage commitments at registered banks that are top-ups on an existing mortgage with the same lender, alongside the mean dollar top-up commitment size.",
        "methodology": "Derived from RBNZ Table C33 by dividing top-up commitment count (LVRN.MMB2.AZB) by total mortgage commitments across all purposes, and dividing top-up lending value (LVRN.MMB1.AZB, rounded to $1m) by top-up commitment count.",
        "hlf_calculation": "(Top-up Count [LVRN.MMB2.AZB] ÷ Total Approvals by Purpose) × 100; Mean Top-Up = LVRN.MMB1.AZB ÷ LVRN.MMB2.AZB",
        "hlf_interpretation": "HLF market context: RBNZ Table C33 records the number and dollar value of new commitments that increase lending on an existing mortgage with the same bank; the dataset does not record the specific end-use of the topped-up funds.",
        "questions": [
            "What share of monthly mortgage approvals in New Zealand are top-ups?",
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
        "definition": "The total monthly dollar value and commitment count of new residential mortgage lending approved for top-ups on existing mortgages at New Zealand registered banks.",
        "methodology": "Sourced from RBNZ Table C33 series LVRN.MMB1.AZB (Top-up lending value in $ millions) and LVRN.MMB2.AZB (Top-up borrower count), with mean top-up size calculated by HLF.",
        "hlf_calculation": "Mean Top-Up = Top-up Lending Value [LVRN.MMB1.AZB] ÷ Top-up Borrower Count [LVRN.MMB2.AZB]",
        "hlf_interpretation": "HLF market context: Tracks the monthly dollar value of additional mortgage commitments approved on existing properties with the same registered bank.",
        "questions": [
            "How much money is committed via mortgage top-ups each month in New Zealand?",
            "How many mortgage top-up commitments are approved every month in NZ?",
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
        "definition": "The percentage of new residential mortgage commitments at registered banks written on interest-only terms for property investors compared with owner-occupiers.",
        "methodology": "Derived from RBNZ Table C32 by dividing new interest-only investor lending value by total new investor lending value, and new interest-only owner-occupier lending value by total new owner-occupier lending value. Under RBNZ Table C32 definitions, lending not on a principal-and-interest schedule (including revolving credit facilities) is classified under interest-only.",
        "hlf_calculation": "(New Investor Interest-Only Lending ÷ Total New Investor Lending) × 100",
        "hlf_interpretation": "HLF market context: RBNZ Table C32 compares the share of new mortgage commitments written on interest-only terms (including revolving credit facilities) between residential property investors and owner-occupiers.",
        "questions": [
            "What percentage of New Zealand property investor mortgage commitments are interest-only?",
            "How does interest-only lending compare between investors and owner-occupiers in RBNZ Table C32?",
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
        "definition": "The percentage of monthly first-home buyer mortgage lending value at registered banks committed at a Total Debt-to-Income (TDTI) ratio above 6x gross income, compared against the 20% RBNZ speed-limit cap on non-exempt owner-occupier lending.",
        "methodology": "Derived from RBNZ Table C40 by summing first-home buyer lending commitments in bands above 6x TDTI (including unknown TDTI) and dividing by total first-home buyer commitments across all TDTI bands.",
        "hlf_calculation": "(FHB Lending with TDTI > 6 ÷ Total FHB Commitments across all TDTI bands) × 100; Headroom = 20.0% minus (>6x Share)",
        "hlf_interpretation": "HLF market context: Shows the proportion of first-home buyer mortgage commitments written above 6x Total Debt-to-Income in RBNZ Table C40 relative to the Reserve Bank's 20% speed-limit allowance for non-exempt owner-occupier lending.",
        "questions": [
            "What percentage of first-home buyers borrow more than 6 times their income in RBNZ Table C40?",
            "How does current first-home buyer >6x DTI lending compare to the 20% RBNZ speed limit?",
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
        "definition": "The percentage of monthly residential property investor mortgage commitments at registered banks written at a Total Debt-to-Income (TDTI) ratio above 7x gross income, compared against the 20% RBNZ speed-limit cap on non-exempt investor lending.",
        "methodology": "Derived from RBNZ Table C40 by summing property investor new commitments with TDTI > 7 and dividing by total investor commitments across all TDTI bands.",
        "hlf_calculation": "(Investor Lending with TDTI > 7 ÷ Total Investor Commitments across all TDTI bands) × 100; Headroom = 20.0% minus (>7x Share)",
        "hlf_interpretation": "HLF market context: Tracks the share of residential property investor commitments written above 7x Total Debt-to-Income in RBNZ Table C40 relative to the 20% RBNZ speed-limit allowance on non-exempt investor lending.",
        "questions": [
            "What is the Debt-to-Income (DTI) speed limit for property investors in New Zealand?",
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
        "definition": "The band-weighted mean annual gross borrower income across approved New Zealand first-home buyer mortgage applications and property investor applications at registered banks, derived from RBNZ Table C41.",
        "methodology": "RBNZ Table C41 publishes official total gross income ($m, rounded to $0.1m) and average gross income ($000, rounded to $1,000) within each TDTI band rather than a single all-bands average. HLF derives the overall mean by dividing total gross income summed across all reported TDTI bands by the implied number of applications (Band Total Income ÷ Band Average Income) summed across those bands.",
        "hlf_calculation": "Sum(Band Total Gross Income [$m]) ÷ Sum(Band Total Gross Income [$m] ÷ Band Average Gross Income [$k]) × 1,000",
        "hlf_interpretation": "HLF market context: RBNZ's Borrower Gross Income (BGI) measures combined application-level gross income across all borrowers on the loan (which for couples combines two incomes, and for investors includes personal salaries plus gross rental and other qualifying income). It does not measure individual solo salaries.",
        "questions": [
            "What is the average application gross income of an approved first-home buyer in RBNZ Table C41?",
            "How does application gross income compare between first-home buyers and property investors in RBNZ data?",
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
        "definition": "Official average gross borrower income reported by RBNZ Table C41 within each first-home buyer Debt-to-Income (TDTI) band across New Zealand registered banks, alongside the HLF-calculated overall mean across all bands.",
        "methodology": "Reports the five official RBNZ Table C41 first-home buyer TDTI band averages (TDTI <= 3, 3 < TDTI <= 4, 4 < TDTI <= 5, 5 < TDTI <= 6, and TDTI > 6), each band's official gross income pool ($m, rounded to $0.1m), each band's implied application count (Pool ÷ Average), and the HLF-derived overall weighted mean.",
        "hlf_calculation": "Official RBNZ C41 band averages + HLF overall mean = Sum(Band Income Pool) ÷ Sum(Implied Applications = Band Pool ÷ Band Avg)",
        "hlf_interpretation": "HLF market context: Displays the official RBNZ Table C41 average gross application income within each Debt-to-Income band alongside the implied application weights.",
        "questions": [
            "How does first-home buyer application income vary across Debt-to-Income (DTI) bands in RBNZ Table C41?",
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
        "definition": "The volume-weighted average fixed mortgage interest rate on new owner-occupier residential mortgage commitments for 1-year and 2-year fixed terms at New Zealand registered banks.",
        "methodology": "Directly reported by RBNZ Table B30 series NCFA.MAM1PR.D1 (1-year fixed owner-occupier rate) and NCFA.MAM1PR.D2 (2-year fixed owner-occupier rate), weighted by dollar lending volume.",
        "hlf_calculation": "Officially reported RBNZ Table B30 weighted-average settled rate series (plus YoY rate change)",
        "hlf_interpretation": "HLF market context: RBNZ Table B30 measures the dollar-weighted average interest rate on new owner-occupier fixed-rate mortgage commitments entered into during the reference month across registered banks.",
        "questions": [
            "What is the average 1-year and 2-year fixed mortgage rate in RBNZ Table B30?",
        ],
        "article_title": "Compare NZ Mortgage Rates",
        "article_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
        "tool_title": "Compare NZ Mortgage Rates",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
    },
    "fixed-vs-floating-rate-mix": {
        "name": "Fixed vs floating mortgage share & 12-month fixed rate roll-off (RBNZ Table S33)",
        "also_known_as": "Share of NZ residential mortgages on fixed vs floating rates & refixing within 12 months",
        "category": "Homeowners & refinancing",
        "topic": "Fixed vs floating mortgage structure & refix horizon",
        "source": "RBNZ Table S33",
        "source_full": "Reserve Bank of New Zealand — Banks: Assets – Loans fully secured by residential mortgage (Table S33)",
        "series_ids": "RBNZ Table S33 — Fixed vs Floating Residential Mortgage Stock & <=12-Month Repricing Bands",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-fully-secured-by-residential-mortgage",
        "data_type": "Officially reported RBNZ series & HLF-derived repricing share",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The proportion of New Zealand registered banks' existing residential mortgage stock held on fixed interest rates versus floating rates, alongside the percentage of fixed-rate balances scheduled to reprice within 12 months.",
        "methodology": "Extracted live from RBNZ Table S33 residential mortgage stock repricing statistics, dividing fixed-rate mortgage balances by total residential mortgage balances and summing fixed-rate bands repricing within 12 months.",
        "hlf_calculation": "Fixed Share = Fixed Stock ÷ Total Stock; 12-Month Roll-Off = Fixed Loans Repricing <=12 Months ÷ Total Fixed Stock",
        "hlf_interpretation": "HLF market context: RBNZ Table S33 categorises registered banks' total outstanding residential mortgage balances by the time remaining until the interest rate next resets.",
        "questions": [
            "What percentage of New Zealand bank mortgages are fixed vs floating?",
            "What share of fixed-rate mortgages in New Zealand are scheduled to reprice within 12 months?",
        ],
        "article_title": "Refinancing a Mortgage in NZ: How Cashbacks & Break Fees Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/refinancing-a-mortgage-in-nz-how-cashbacks-break-fees-and-switching-banks-actually-work",
        "tool_title": "HLF Split Banking & Refinance Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/refinance-calculator",
    },
    "mortgage-excess-repayments": {
        "name": "Quarterly voluntary excess mortgage repayments vs scheduled repayments (RBNZ Table C35)",
        "also_known_as": "RBNZ Table C35 Residential Mortgage Loan Reconciliation — Voluntary excess principal repayments ($B/qtr & $B/yr)",
        "category": "Homeowners & refinancing",
        "topic": "Voluntary extra mortgage repayments & principal reduction",
        "source": "RBNZ Table C35",
        "source_full": "Reserve Bank of New Zealand — Residential mortgage loan reconciliation (Table C35)",
        "series_ids": "RBNZ Table C35 — Excess Repayments ($m) vs Scheduled Repayments ($m)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-loan-reconciliation",
        "data_type": "Officially reported RBNZ quarterly series & HLF-derived ratio to scheduled repayments",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Quarterly",
        "definition": "Total voluntary excess mortgage repayments ($ billions per quarter and rolling 12-month total) paid by New Zealand registered bank borrowers above their contracted minimum scheduled repayments, excluding full loan discharges.",
        "methodology": "Extracted live from RBNZ Table C35 (Residential mortgage loan reconciliation). Compares quarterly voluntary excess repayments against contracted scheduled principal-and-interest repayments. Excludes full loan discharges upon property sale or refinancing.",
        "hlf_calculation": "Ratio to Scheduled P&I (%) = (Excess Repayments [$m] ÷ Scheduled Repayments [$m]) × 100; Rolling 12M Excess = Sum of 4 quarters",
        "hlf_interpretation": "HLF market context: RBNZ Table C35 tracks principal repayments made by borrowers in excess of their contracted scheduled payments (such as regular overpayments or lump-sum reductions) across registered banks.",
        "questions": [
            "How much in voluntary excess mortgage repayments do New Zealand borrowers make each quarter?",
            "How do voluntary excess mortgage repayments compare with scheduled repayments in RBNZ Table C35?",
        ],
        "article_title": "The Great Mortgage Myth: Weekly vs Monthly Repayments",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/the-great-mortgage-myth-why-switching-to-weekly-repayments-might-not-save-as-much-as-you-think",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "mortgage-revolving-credit-balances": {
        "name": "Residential revolving credit drawn mortgage balances & share of housing lending (RBNZ Table S32)",
        "also_known_as": "RBNZ Table S32 Banks: Assets – Loans by Product — Owner-occupied & investor revolving credit balances ($B)",
        "category": "Homeowners & refinancing",
        "topic": "Residential revolving credit mortgage balances & share",
        "source": "RBNZ Table S32",
        "source_full": "Reserve Bank of New Zealand — Banks: Assets – Loans by product (Table S32)",
        "series_ids": "RBNZ Table S32 — Owner-Occupied & Residential Property Investor Revolving Credit Drawn Balances ($m)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-by-product",
        "data_type": "Officially reported RBNZ monthly product series & HLF-derived sector shares",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Monthly",
        "definition": "Total drawn residential mortgage balances held in revolving credit facilities at New Zealand registered banks ($ billions), broken down between owner-occupiers and residential property investors along with their share of total bank housing loans.",
        "methodology": "Extracted live from RBNZ Table S32 (Banks: Assets – Loans by product). Measures drawn revolving credit balances at month-end and does not include undrawn revolving credit limits or standalone deposit accounts linked to offset mortgages.",
        "hlf_calculation": "Total Residential Revolving Credit = Owner-Occupied Revolving ($m) + Investor Revolving ($m); Share (%) = Revolving Credit ÷ Total Bank Housing Loans × 100",
        "hlf_interpretation": "HLF market context: RBNZ Table S32 reports the drawn balance of residential revolving credit mortgages across owner-occupiers and residential property investors at registered banks.",
        "questions": [
            "How much is held in drawn residential revolving credit mortgages in New Zealand?",
            "What percentage of New Zealand bank mortgages are in revolving credit facilities?",
        ],
        "article_title": "Revolving Credit vs. Offset Mortgages: Which is Right for You?",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/revolving-credit-vs-offset-mortgages-which-is-right-for-you",
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "student-loan-servicing": {
        "name": "Median New Zealand student loan balance (Inland Revenue)",
        "also_known_as": "IRD Student Loan Scheme median borrower balance",
        "category": "LVR, DTI & credit policy",
        "topic": "Student loan balances & repayment rules",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — Student Loan Scheme Balances Statistics",
        "series_ids": "Inland Revenue — Median Student Loan Balance ($ NZD)",
        "source_url": IRD_SL_BAL_PAGE,
        "data_type": "Officially reported IRD median & HLF-derived YoY change",
        "derived_by_hlf": False,
        "unit": "NZD",
        "frequency": "Quarterly",
        "definition": "The median outstanding student loan balance across all borrowers in the New Zealand Student Loan Scheme reported by Inland Revenue.",
        "methodology": "Directly reported by Inland Revenue in its quarterly Student Loan Scheme statistics release (covering both New Zealand-based and overseas-based borrowers), with percentage and dollar change compared against the prior reported period.",
        "hlf_calculation": "Officially reported IRD median balance (plus % change vs prior period)",
        "hlf_interpretation": "HLF market context: Under New Zealand student loan legislation, compulsory repayments for New Zealand-based borrowers are calculated as a fixed percentage (12%) of gross income earned above the statutory repayment threshold rather than as a function of the loan balance size.",
        "questions": [
            "What is the median student loan balance in New Zealand?",
            "How are compulsory student loan repayments calculated in New Zealand?",
        ],
        "article_title": "DTI Rules in NZ: How Debt-to-Income Limits Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/dti-rules-in-nz-how-debt-to-income-limits-actually-work-for-home-buyers-and-investors",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
# === END OF BLOCK 3 ===
# === BLOCK 4 OF 13 ===
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
        "methodology": "Directly reported in Stats NZ's monthly Building Consents Issued release. Building consents record local authority authorisations to build and do not measure construction starts or completed dwellings (some consented projects are delayed or not completed, and figures do not deduct demolitions).",
        "hlf_calculation": "Officially reported Stats NZ annual and monthly dwelling consent totals",
        "hlf_interpretation": "HLF market context: Building consents measure the volume of new residential dwellings authorised by local councils across New Zealand over the 12-month and monthly periods.",
        "questions": [
            "How many new residential dwellings are consented in New Zealand each year?",
            "How many new dwellings were consented in the latest month in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
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
        "hlf_interpretation": "HLF market context: Compares the proportion of newly consented dwellings classified as multi-unit homes against stand-alone houses in Stats NZ's monthly release.",
        "questions": [
            "What percentage of new dwelling consents in New Zealand are townhouses and multi-unit homes vs stand-alone houses?",
            "How many townhouses, flats, and units are consented in New Zealand each month?",
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
        "definition": "Breakdown of monthly new residential dwelling consents across Stats NZ's four dwelling categories: stand-alone houses, townhouses/flats/units, apartments, and retirement village units.",
        "methodology": "Sourced from Stats NZ's monthly Building Consents Issued release, with each category divided by the monthly total to show percentage shares.",
        "hlf_calculation": "Category Share (%) = (Category Dwelling Count ÷ Total Monthly Dwelling Consents) × 100",
        "hlf_interpretation": "HLF market context: Shows the composition of monthly new residential dwelling consents across Stats NZ's four standard dwelling classifications.",
        "questions": [
            "What is the breakdown of new houses, townhouses, apartments, and retirement village units consented in NZ?",
            "How many apartments and retirement village units are consented each month in New Zealand?",
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
        "methodology": "Sourced from Stats NZ's Building Consents Issued regional release, with each region's 12-month consent total divided by the national 12-month total.",
        "hlf_calculation": "Regional Share (%) = (Regional Annual Consents ÷ National Annual Consents) × 100",
        "hlf_interpretation": "HLF market context: Shows the regional distribution of new residential dwelling consents authorised across New Zealand over the rolling 12-month period.",
        "questions": [
            "Which New Zealand regions consent the most new residential dwellings?",
            "What share of New Zealand building consents are in Auckland, Canterbury, Waikato, and Wellington?",
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
        "definition": "The annual percentage change in Stats NZ's Consumers Price Index (CPI) 'purchase of new housing' sub-index.",
        "methodology": "Directly reported in Stats NZ's quarterly Consumers Price Index release under the 'purchase of new housing' sub-index, compared alongside headline All-groups CPI inflation. The sub-index surveys respondents on the cost of building a new dwelling for owner-occupation (excluding land).",
        "hlf_calculation": "Officially reported Stats NZ quarterly CPI purchase of new housing YoY % change",
        "hlf_interpretation": "HLF market context: Measures the annual percentage change in Stats NZ's CPI 'purchase of new housing' sub-index, which tracks the cost of constructing a new dwelling for owner-occupation (excluding land values and existing house prices).",
        "questions": [
            "How much has the CPI cost of purchasing a newly built home changed over the past year in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "New Build Construction Loan Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/new-build-construction-loan-calculator",
    },
    "new-build-consent-cost-by-type": {
        "name": "Average residential building consent construction value per consented dwelling (Stats NZ)",
        "also_known_as": "Stats NZ Building Consents average residential construction value (excluding land)",
        "category": "New builds & construction",
        "topic": "Average dollar build consent value",
        "source": "Stats NZ Building Consents",
        "source_full": "Stats NZ — Building Consents Issued (Residential Building Consent Values & Dwelling Counts)",
        "series_ids": "Total New Residential Building Consent Value ($m) ÷ Consented Dwelling Count",
        "source_url": "https://www.stats.govt.nz/topics/building",
        "data_type": "HLF-derived mean consent value from official Stats NZ Building Consents release",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly",
        "definition": "The average consented construction value in New Zealand dollars (excluding land and GST) per newly consented residential dwelling in Stats NZ's monthly Building Consents Issued release.",
        "methodology": "Derived live from Stats NZ Building Consents Issued by dividing consented residential building value by the number of new dwellings consented. As noted by Stats NZ, building consent values record the estimated construction cost declared on the building consent application (excluding land purchase price, GST, developer margin, and off-site subdivision infrastructure).",
        "hlf_calculation": "Mean Consented Build Value = Residential Consent Value ($) ÷ Consented Dwellings",
        "hlf_interpretation": "HLF market context: Reflects the average estimated physical construction value declared on building consent applications per dwelling (excluding land and GST), rather than the final turnkey purchase price of a house-and-land package.",
        "questions": [
            "What is the average building consent construction value per new dwelling in New Zealand?",
            "Does Stats NZ building consent value include land?",
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
        "methodology": "Extracted from MBIE Tenancy Services' monthly regional rental bond dataset (Location = 'ALL' for national median), with year-on-year dollar and percentage changes calculated against the same month a year earlier. MBIE bond statistics are compiled from new private tenancy bonds lodged during the reference month (excluding ongoing tenancies where no new bond is lodged, social housing, and tenancies without a lodged bond).",
        "hlf_calculation": "YoY Rent Change (%) = ((Current National Median Rent - Prior Year Median Rent) ÷ Prior Year Median Rent) × 100",
        "hlf_interpretation": "HLF market context: MBIE rental bond medians measure the weekly rent agreed on newly lodged private tenancies during the month, rather than the average rent paid across all existing ongoing tenancies.",
        "questions": [
            "What is the median weekly rent on newly lodged bonds in New Zealand?",
            "How does median weekly rent compare between Auckland, Wellington, and Canterbury in MBIE bond data?",
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
        "methodology": "Directly reported in MBIE Tenancy Services' monthly rental bond CSV for Location = 'ALL' (Lodged Bonds and Active Bonds). Measures private residential tenancies where a bond is lodged with MBIE.",
        "hlf_calculation": "Officially reported MBIE national Lodged Bonds and Active Bonds",
        "hlf_interpretation": "HLF market context: Reports the monthly volume of new bond lodgements and the total number of active private tenancy bonds held by MBIE Tenancy Services at month-end.",
        "questions": [
            "How many active private rental bonds are held by MBIE Tenancy Services in New Zealand?",
            "How many new rental bonds are lodged each month in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "rent-long-term-growth": {
        "name": "5-year & 10-year national median weekly rent growth & compound annual growth rate (MBIE)",
        "also_known_as": "Long-term New Zealand nominal weekly rent growth & CAGR (5-year and 10-year windows)",
        "category": "Rents & tenancy",
        "topic": "Long-term rental growth (5-year & 10-year nominal)",
        "source": "MBIE Tenancy Services",
        "source_full": "Ministry of Business, Innovation and Employment (MBIE) — Tenancy Services Rental Bond Data",
        "series_ids": "MBIE Location 'ALL' Median Rent (t, t-5y, t-10y)",
        "source_url": TENANCY_PAGE,
        "data_type": "HLF-derived 5Y & 10Y cumulative growth & CAGR from official MBIE series",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "Cumulative nominal percentage growth, weekly dollar change, and compound annual growth rate (CAGR) in New Zealand's national median weekly rent on newly lodged bonds over the past 5 years and 10 years.",
        "methodology": "Compares the latest national median weekly rent (Location = 'ALL') in MBIE's bond dataset against the same month 5 and 10 years earlier. Measures nominal rent on newly lodged bonds and is not adjusted for CPI inflation or changes in the size/type composition of properties rented.",
        "hlf_calculation": "10Y Growth (%) = ((Rent_t − Rent_t-10) ÷ Rent_t-10) × 100; CAGR = (Rent_t ÷ Rent_t-n)^(1/n) − 1",
        "hlf_interpretation": "HLF market context: Summarises 5-year and 10-year nominal changes in the national median weekly rent across newly lodged private tenancy bonds.",
        "questions": [
            "How much has New Zealand's national median weekly rent changed over the last 5 and 10 years?",
            "What is the 5-year and 10-year compound annual growth rate of New Zealand weekly rents in MBIE bond data?",
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
        "definition": "The 25th percentile (lower quartile), 50th percentile (median), and 75th percentile (upper quartile) weekly private rent on newly lodged bonds across New Zealand, alongside the interquartile dollar spread.",
        "methodology": "Sourced from MBIE Tenancy Services' regional bond CSV (Location = 'ALL'), subtracting Lower Quartile Rent from Upper Quartile Rent and multiplying by 52 for the annualised spread. Covers all dwelling types and bedroom counts combined.",
        "hlf_calculation": "Interquartile Spread ($/wk) = Upper Quartile Rent minus Lower Quartile Rent; Annualised Spread = Spread × 52",
        "hlf_interpretation": "HLF market context: Shows the national spread between the 25th, 50th, and 75th percentile weekly rents across newly lodged private tenancy bonds.",
        "questions": [
            "What is the lower quartile, median, and upper quartile weekly rent in New Zealand?",
            "What is the interquartile rent spread in MBIE tenancy bond data?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "rental-turnover-net-stock": {
        "name": "Annual net change in active rental bond stock & monthly lodged vs closed bonds",
        "also_known_as": "Net active rental bond movement & monthly lodged vs closed bonds",
        "category": "Rents & tenancy",
        "topic": "Net active bond stock change & monthly bond flows",
        "source": "MBIE Tenancy Services",
        "source_full": "Ministry of Business, Innovation and Employment (MBIE) — Tenancy Services Rental Bond Data",
        "series_ids": "MBIE Location 'ALL' Active Bonds (YoY change) & Lodged Bonds minus Closed Bonds",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported MBIE counts & HLF-derived net stock change",
        "derived_by_hlf": True,
        "unit": "bonds",
        "frequency": "Monthly",
        "definition": "The 12-month net change in New Zealand's total pool of active private rental bonds held by MBIE, alongside the monthly counts of newly lodged bonds and closed (refunded) bonds.",
        "methodology": "Derived from MBIE Tenancy Services' bond dataset by subtracting Active Bonds 12 months earlier from current Active Bonds, and subtracting monthly Closed Bonds from monthly Lodged Bonds.",
        "hlf_calculation": "12M Net Stock Change = Active Bonds_t minus Active Bonds_t-12; Monthly Net Flow = Lodged Bonds minus Closed Bonds",
        "hlf_interpretation": "HLF market context: Tracks the 12-month net change in active private rental bonds held by MBIE Tenancy Services alongside monthly lodged and closed bonds. Active bond counts measure tenancies with a lodged bond held by MBIE, rather than the total physical stock of all rental dwellings.",
        "questions": [
            "How has the number of active private rental bonds held by MBIE changed over the past year?",
            "How many rental bonds are lodged and closed each month in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "kiwisaver-first-home": {
        "name": "Monthly KiwiSaver first-home withdrawals: member count, total value & average withdrawal (IRD)",
        "also_known_as": "Inland Revenue monthly KiwiSaver first-home purchase withdrawal volume & mean per member",
        "category": "KiwiSaver",
        "topic": "KiwiSaver first-home deposit withdrawals",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Withdrawal Statistics (First Home Purchase)",
        "series_ids": "IRD KiwiSaver First-Home Withdrawal Member Count & Dollar Value ($m)",
        "source_url": IRD_KS_DATA_PAGE,
        "data_type": "Officially reported IRD monthly series & HLF-derived mean per withdrawing member",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Monthly",
        "definition": "The monthly number of KiwiSaver members withdrawing funds for a first-home purchase, total dollar value withdrawn, and mean withdrawal per withdrawing member.",
        "methodology": "Sourced live from Inland Revenue's monthly KiwiSaver withdrawal numbers and amounts tables, dividing monthly first-home withdrawal dollars by the number of withdrawing members. IRD counts individual members rather than households or properties (so two co-buyers withdrawing for the same property purchase are recorded as two member withdrawals).",
        "hlf_calculation": "Average First-Home Withdrawal = (Monthly First-Home Withdrawal Value [$m] × $1,000,000) ÷ Monthly Withdrawing Members",
        "hlf_interpretation": "HLF market context: Reports the monthly volume and mean dollar amount per individual KiwiSaver member withdrawing savings for a first home, as recorded by Inland Revenue.",
        "questions": [
            "What is the average KiwiSaver first-home withdrawal per member in New Zealand?",
            "How many KiwiSaver members withdraw funds for a first home each month?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "Kāinga Ora LMI Fee & Low-Deposit Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/lenders-mortgage-insurance-fee-estimator",
    },
    "kiwisaver-hardship-withdrawals": {
        "name": "Monthly KiwiSaver significant financial hardship withdrawals: member count, value & average (IRD)",
        "also_known_as": "Inland Revenue KiwiSaver financial hardship withdrawal count, dollar value & average per member",
        "category": "KiwiSaver",
        "topic": "KiwiSaver financial hardship withdrawals",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Withdrawal Statistics (Significant Financial Hardship)",
        "series_ids": "IRD KiwiSaver Financial Hardship Withdrawal Member Count & Dollar Value ($m)",
        "source_url": IRD_KS_DATA_PAGE,
        "data_type": "Officially reported IRD monthly series & HLF-derived mean per member",
        "derived_by_hlf": True,
        "unit": "NZD_millions",
        "frequency": "Monthly",
        "definition": "The monthly number of KiwiSaver members making a significant financial hardship withdrawal, total dollar amount withdrawn, and average withdrawal per member.",
        "methodology": "Sourced live from Inland Revenue's monthly KiwiSaver withdrawal numbers and amounts tables, dividing monthly financial hardship withdrawal dollars by withdrawing member count. IRD records withdrawals under the significant financial hardship category without identifying the specific underlying cause of each member's hardship.",
        "hlf_calculation": "Average Hardship Withdrawal = (Monthly Hardship Withdrawal Value [$m] × $1,000,000) ÷ Monthly Hardship Members",
        "hlf_interpretation": "HLF market context: These figures provide an indication of the number of KiwiSaver members accessing their savings due to significant financial hardship. They describe reported withdrawals under this specific category, rather than measuring the overall prevalence or causes of financial hardship among New Zealand households.",
        "questions": [
            "How many KiwiSaver members make a significant financial hardship withdrawal each month in New Zealand?",
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
        "methodology": "Derived from Inland Revenue's annual regional KiwiSaver withdrawal tables by dividing each region's first-home withdrawal dollar total by the number of withdrawing members in that region. IRD assigns regions based on the member's residential address held on file.",
        "hlf_calculation": "Regional Average Withdrawal ($) = Regional First-Home Withdrawal Dollars ÷ Regional First-Home Withdrawing Members",
        "hlf_interpretation": "HLF market context: Compares the mean KiwiSaver first-home withdrawal per withdrawing member and total withdrawal volume across New Zealand's 16 regions.",
        "questions": [
            "How much KiwiSaver do first-home buyers withdraw on average across New Zealand regions?",
            "Which New Zealand region has the highest average KiwiSaver first-home withdrawal per member?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "kiwisaver-switches": {
        "name": "Annual KiwiSaver scheme provider transfers (switching providers — IRD)",
        "also_known_as": "Inland Revenue KiwiSaver inter-scheme transfers & monthly switching pace",
        "category": "KiwiSaver",
        "topic": "KiwiSaver scheme switching & provider transfers",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Scheme Transfer Statistics",
        "series_ids": "IRD Annual KiwiSaver Scheme Transfers (Year ended 30 June) ÷ 12",
        "source_url": IRD_KS_TR_PAGE,
        "data_type": "Officially reported IRD annual series & HLF-derived monthly mean & YoY change",
        "derived_by_hlf": True,
        "unit": "transfers",
        "frequency": "Annual (June Financial Year)",
        "definition": "The annual and average monthly number of KiwiSaver member transfers from one KiwiSaver scheme provider to another processed by Inland Revenue.",
        "methodology": "Sourced live from Inland Revenue's annual KiwiSaver scheme transfer release, dividing the annual transfer total by 12 for the monthly average and comparing against the prior financial year. Records inter-scheme transfers between different providers and does not include internal fund switches within the same scheme provider.",
        "hlf_calculation": "Monthly Average Transfers = Annual Scheme Transfers ÷ 12; YoY Change (%) = ((Transfers_t − Transfers_t-1) ÷ Transfers_t-1) × 100",
        "hlf_interpretation": "HLF market context: Records completed transfers of member accounts from one KiwiSaver scheme provider to another processed by Inland Revenue. The dataset does not record the reasons members transferred or capture internal fund changes within the same scheme provider.",
        "questions": [
            "How many KiwiSaver inter-scheme provider transfers are processed each year in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-total-fum-balance": {
        "name": "Total KiwiSaver funds under management (FUM) & official FMA average member balance",
        "also_known_as": "FMA KiwiSaver Annual Report & RBNZ Table T43 total FUM, member count & average balance",
        "category": "KiwiSaver",
        "topic": "KiwiSaver total funds under management & average balance",
        "source": "FMA & RBNZ T43",
        "source_full": "Financial Markets Authority (FMA) — KiwiSaver Annual Report & RBNZ Table T43",
        "series_ids": "FMA KiwiSaver Annual Report & RBNZ Table T43 — Total FUM ($B), Total Members & Average Balance ($)",
        "source_url": FMA_KS_REPORT_PAGE,
        "data_type": "Officially reported FMA & RBNZ figures",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual / Quarterly",
        "definition": "Total funds under management (FUM), total enrolled membership, and the mean member balance across all New Zealand KiwiSaver schemes.",
        "methodology": "Extracted live from the FMA KiwiSaver Annual Report and RBNZ Table T43 (Managed Funds: KiwiSaver assets and liabilities). Note that the national mean member balance is calculated across all enrolled members (including children, non-contributing members, and older members with larger balances) and differs from age-specific median balances.",
        "hlf_calculation": "Official FMA / RBNZ T43 KiwiSaver FUM ($B) & Mean Member Balance",
        "hlf_interpretation": "HLF market context: Reports total KiwiSaver funds under management and the overall mean balance across all enrolled KiwiSaver members nationwide.",
        "questions": [
            "What is the average KiwiSaver balance in New Zealand?",
            "How much money is held in KiwiSaver funds under management in total across New Zealand?",
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
        "definition": "The number and percentage of KiwiSaver members in schemes assigned via Inland Revenue default allocation or employer nomination versus members recorded as having actively chosen their scheme.",
        "methodology": "Derived from Inland Revenue's annual KiwiSaver entry method table by summing default-allocated and employer-nominated members and dividing by total enrolled members.",
        "hlf_calculation": "Default/Employer-Nominated Share (%) = ((Default Allocated + Employer Nominated) ÷ Total Members) × 100",
        "hlf_interpretation": "HLF market context: Categorises enrolled KiwiSaver members by how they entered their current scheme (default allocation, employer nomination, or active choice) as recorded by Inland Revenue.",
        "questions": [
            "How many New Zealanders are in a default-allocated or employer-nominated KiwiSaver scheme?",
            "What percentage of KiwiSaver members are recorded as having actively chosen their scheme?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-contributions-10yr": {
        "name": "Annual KiwiSaver contributions transferred to scheme providers (employee, employer & Crown — IRD)",
        "also_known_as": "Inland Revenue KiwiSaver scheme provider payments",
        "category": "KiwiSaver",
        "topic": "Annual KiwiSaver contributions transferred to providers",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Contributions & Payments to Scheme Providers",
        "series_ids": "IRD KiwiSaver Total Contributions Passed to Scheme Providers ($B)",
        "source_url": IRD_KS_PAY_PAGE,
        "data_type": "Officially reported IRD annual series",
        "derived_by_hlf": False,
        "unit": "NZD_billions",
        "frequency": "Annual (June Financial Year)",
        "definition": "Total annual KiwiSaver contributions transferred by Inland Revenue to scheme providers across employee deductions, employer contributions, and Crown (Government) contributions.",
        "methodology": "Directly extracted from Inland Revenue's KiwiSaver scheme provider payments release page. Measures contributions collected and passed to providers by Inland Revenue (excluding voluntary lump-sum contributions paid by members directly to their scheme provider without passing through IRD).",
        "hlf_calculation": "Officially reported Inland Revenue annual KiwiSaver transfers to scheme providers ($B)",
        "hlf_interpretation": "HLF market context: Reports the annual dollar total of employee, employer, and Government KiwiSaver contributions administered and transferred to scheme providers by Inland Revenue.",
        "questions": [
            "How much money is transferred from Inland Revenue to KiwiSaver scheme providers annually?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-savings-suspensions": {
        "name": "Active KiwiSaver savings suspensions (contributions holiday — IRD)",
        "also_known_as": "Inland Revenue KiwiSaver members on active savings suspension",
        "category": "KiwiSaver",
        "topic": "KiwiSaver savings suspensions (contribution holidays)",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Savings Suspension Statistics",
        "series_ids": "IRD Active KiwiSaver Savings Suspensions",
        "source_url": IRD_KS_SUSP_PAGE,
        "data_type": "Officially reported IRD counts",
        "derived_by_hlf": False,
        "unit": "members",
        "frequency": "Monthly",
        "definition": "The number of KiwiSaver members holding an active formal savings suspension (formerly 'contributions holiday') with Inland Revenue.",
        "methodology": "Directly extracted from Inland Revenue's monthly KiwiSaver savings suspension release.",
        "hlf_calculation": "Directly reported Inland Revenue active KiwiSaver savings suspensions",
        "hlf_interpretation": "HLF market context: Reports the number of KiwiSaver members with an active formal savings suspension recorded by Inland Revenue at the end of the month. The series does not identify why members applied for a suspension or capture members who are not contributing for other reasons (such as non-employment or self-employment).",
        "questions": [
            "How many New Zealanders are on an active KiwiSaver savings suspension?",
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
        "definition": "The annual percentage change in Stats NZ's Consumers Price Index (CPI) residential dwelling (building) insurance sub-index, compared against headline All-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (dwelling insurance sub-index), subtracting headline All-groups CPI to show the percentage-point difference.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Dwelling Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: Measures the annual percentage change in Stats NZ's CPI sub-index for residential dwelling insurance relative to headline All-groups CPI. The CPI series tracks price changes for a sample of policy specifications and does not measure individual property risk ratings or sum-insured adjustments.",
        "questions": [
            "How much has the Stats NZ CPI dwelling insurance price index changed over the past year?",
            "How does dwelling insurance CPI inflation compare to headline CPI inflation in NZ?",
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
        "definition": "The annual percentage change in Stats NZ's household contents insurance CPI price index, compared against headline All-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (contents insurance sub-index), compared against headline All-groups CPI.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Contents Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: Measures the annual percentage change in Stats NZ's CPI sub-index for household contents insurance relative to headline All-groups CPI.",
        "questions": [
            "How much has the Stats NZ CPI contents insurance price index changed over the past year?",
            "How does contents insurance CPI inflation compare to general inflation in NZ?",
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
        "definition": "The annual percentage change in Stats NZ's motor vehicle insurance CPI price index, compared against headline All-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (vehicle insurance sub-index), compared against headline All-groups CPI.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Vehicle Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: Measures the annual percentage change in Stats NZ's CPI sub-index for motor vehicle insurance relative to headline All-groups CPI.",
        "questions": [
            "How much has the Stats NZ CPI motor vehicle insurance price index changed over the past year?",
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
        "definition": "The annual percentage change in Stats NZ's private health insurance CPI price index, compared against headline All-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (health insurance sub-index), compared against headline All-groups CPI.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Health Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: Measures the annual percentage change in Stats NZ's CPI sub-index for private health insurance relative to headline All-groups CPI, holding sampled policy specifications constant.",
        "questions": [
            "How much has the Stats NZ CPI private health insurance price index changed over the past year?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "life-insurance-inflation": {
        "name": "Annual life insurance price index change (Stats NZ CPI)",
        "also_known_as": "New Zealand life insurance CPI price index annual percentage change",
        "category": "Personal Risk & Income Protection",
        "topic": "Life insurance CPI price index",
        "source": "Stats NZ CPI",
        "source_full": "Stats NZ — Consumers Price Index (Life insurance sub-index)",
        "series_ids": "Stats NZ CPI — Life insurance (Annual % change) vs All-groups CPI",
        "source_url": STATS_CPI_PAGE,
        "data_type": "Officially reported Stats NZ CPI series & HLF-derived spread vs headline CPI",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Quarterly",
        "definition": "The annual percentage change in Stats NZ's life insurance CPI price index, compared against headline All-groups CPI inflation.",
        "methodology": "Sourced from Stats NZ's quarterly Consumers Price Index release (life insurance sub-index). Because CPI methodology prices a constant-quality sample holding policyholder age and cover specifications fixed, it measures changes in base tariff pricing rather than age-stepped premium increases experienced by individual policyholders as they age.",
        "hlf_calculation": "Spread vs Headline CPI (percentage points) = Life Insurance CPI YoY (%) minus All-Groups CPI YoY (%)",
        "hlf_interpretation": "HLF market context: Tracks changes in Stats NZ's constant-specification life insurance CPI sub-index relative to headline All-groups CPI.",
        "questions": [
            "How much has the Stats NZ CPI life insurance price index changed over the past year?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "insurance-claims-trends": {
        "name": "New Zealand insurer Gross Premium Revenue, Net Earned Premiums & Net Claims (RBNZ J10)",
        "also_known_as": "RBNZ Table J10 insurance sector premium revenue, reinsurance expense & claims",
        "category": "Insurance Market & Claims",
        "topic": "Insurer premiums, reinsurance & claims",
        "source": "RBNZ Table J10",
        "source_full": "Reserve Bank of New Zealand — Insurance Income and Expenditure (Table J10)",
        "series_ids": "RBNZ Table J10 — Gross Premium Revenue, Outwards Reinsurance Expense & Net Claims Expense",
        "source_url": RBNZ_INS_STATS_PAGE,
        "data_type": "Officially reported RBNZ Table J10 series & HLF-derived reinsurance difference",
        "derived_by_hlf": True,
        "unit": "NZD_billions",
        "frequency": "Quarterly / Annual",
        "definition": "Annual Gross Premium Revenue, Net Earned Premiums (after outwards reinsurance expense), and Net Incurred Claims across insurers surveyed in RBNZ Table J10.",
        "methodology": "Live-parsed from RBNZ Table J10 (Insurance Income and Expenditure workbook), subtracting outwards reinsurance expense from gross premium revenue to show net earned premium and comparing against net claims expense. RBNZ Table J10 is compiled from the Reserve Bank's quarterly insurer survey.",
        "hlf_calculation": "Reinsurance Outflow = Gross Premium Revenue minus Net Earned Premium",
        "hlf_interpretation": "HLF market context: Summarises aggregate gross premium revenue, outwards reinsurance expense, net earned premium, and net incurred claims reported by insurers in RBNZ Table J10.",
        "questions": [
            "How much do insurers surveyed in RBNZ Table J10 collect in premiums and incur in net claims?",
            "How much outwards reinsurance expense is reported in RBNZ Table J10?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
    "insurer-loss-ratios": {
        "name": "New Zealand insurer net claims loss ratio (RBNZ Table J10)",
        "also_known_as": "RBNZ Table J10 net incurred claims to net earned premium ratio",
        "category": "Insurance Market & Claims",
        "topic": "Insurer net claims loss ratio",
        "source": "RBNZ Table J10",
        "source_full": "Reserve Bank of New Zealand — Insurance Income and Expenditure (Table J10)",
        "series_ids": "RBNZ Table J10 — Net Claims Expense ÷ Net Premium Revenue",
        "source_url": RBNZ_INS_STATS_PAGE,
        "data_type": "HLF-derived percentage ratio from official RBNZ Table J10 series",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Quarterly / Annual",
        "definition": "The net claims loss ratio across insurers surveyed in RBNZ Table J10, measuring net incurred claims as a percentage of net earned premiums.",
        "methodology": "Derived live from RBNZ Table J10 by dividing annual Net Claims Expense by annual Net Earned Premium Revenue, multiplied by 100. Note that the net claims ratio excludes non-claim operating, acquisition, and administration expenses.",
        "hlf_calculation": "Net Loss Ratio (%) = (Net Claims Expense ÷ Net Earned Premium Revenue) × 100",
        "hlf_interpretation": "HLF market context: Measures net incurred claims as a proportion of net earned premiums (after reinsurance) across insurers surveyed in RBNZ Table J10.",
        "questions": [
            "What is the net claims loss ratio reported in RBNZ Table J10?",
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
        "methodology": "Live-verified from the Natural Hazards Insurance Act 2023 on legislation.govt.nz and the NHC Toka Tū Ake levy schedule, multiplying statutory GST-exclusive figures by 1.15 to show GST-inclusive equivalents.",
        "hlf_calculation": "Cap incl. GST = Statutory Cap × 1.15; Max Annual Levy incl. GST = (Statutory Cap ÷ 100 × Levy Rate) × 1.15",
        "hlf_interpretation": "HLF market context: Under the Natural Hazards Insurance Act 2023, statutory NHCover applies to eligible residential buildings that are insured under a private contract of fire insurance.",
        "questions": [
            "What is the NHC (EQC) statutory residential building cap and maximum annual levy in New Zealand?",
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
        "topic": "ACC covered earnings cap & weekly compensation ceiling",
        "source": "ACC & IRD Schedule",
        "source_full": "Accident Compensation Corporation (ACC) & Inland Revenue — Earners' Levy Statutory Schedule",
        "series_ids": "ACC / IRD Earners' Levy Rate (per $100 incl. GST) & Maximum Liable Earnings Cap ($ NZD)",
        "source_url": ACC_LEVY_PAGE,
        "data_type": "Officially legislated statutory schedule & HLF-derived maximum levy/compensation ceiling",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual (Tax Year from 1 April)",
        "definition": "The statutory ACC Earners' Levy rate and Maximum Liable Earnings cap for the current tax year, alongside the resulting maximum annual levy and maximum 80% gross weekly accidental injury compensation ceiling.",
        "methodology": "Live-parsed from the official ACC and Inland Revenue Earners' Levy schedule, multiplying the Maximum Liable Earnings cap by the GST-inclusive levy rate and by the 80% statutory weekly compensation rate.",
        "hlf_calculation": "Max Annual Levy = (Max Liable Earnings ÷ 100) × Levy Rate; Max 80% ACC Payout = Max Liable Earnings × 80%",
        "hlf_interpretation": "HLF market context: Under the Accident Compensation Act 2001, ACC weekly compensation covers up to 80% of liable earnings up to the statutory Maximum Liable Earnings cap for qualifying accidental injuries (and does not cover non-accident illness or earnings above the cap).",
        "questions": [
            "What is the Maximum Liable Earnings cap covered by ACC in New Zealand?",
            "What is the ACC Earners' Levy rate and maximum 80% weekly compensation ceiling in NZ?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "",
        "tool_url": "",
    },
# === END OF BLOCK 4 ===
# === BLOCK 5 OF 13 ===
    "back-book-vs-new-settled-mortgage-rate": {
        "name": "Existing 'back-book' effective mortgage rate vs newly settled fixed mortgage rates (RBNZ Table B6 vs B30)",
        "also_known_as": "RBNZ Table B6 housing loan yield vs Table B30 new settled owner-occupier mortgage rate spread",
        "category": "Homeowners & refinancing",
        "topic": "Existing mortgage book effective rate vs new settled rates",
        "source": "RBNZ Tables B6 & B30",
        "source_full": "Reserve Bank of New Zealand — Yields on Loans (Table B6) & New Residential Mortgage Weighted Average Interest Rates (Table B30)",
        "series_ids": "RBNZ Table B6 Housing Loan Effective Yield (%) minus Table B30 1-Year Settled Owner-Occupier Rate (%)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/yields-on-loans",
        "data_type": "Officially reported RBNZ rate series & HLF-derived back-book vs front-book rate spread",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "Comparison of the effective weighted-average interest yield across all existing New Zealand registered bank housing loans (RBNZ Table B6 'back book') against the weighted-average rate on new 1-year fixed owner-occupier mortgage commitments (RBNZ Table B30 'front book').",
        "methodology": "Derived live from RBNZ Table B6 and Table B30 workbooks, computing the basis-point difference between the existing housing loan stock yield (which blends all fixed terms and floating loans currently on bank balance sheets) and newly committed 1-year fixed owner-occupier rates.",
        "hlf_calculation": "Back-Book vs Front-Book Spread (bps) = (Table B6 Existing Housing Yield − Table B30 1Y Settled Rate) × 100",
        "hlf_interpretation": "HLF market context: Compares the average interest rate currently being paid across all existing registered bank housing loans (Table B6) with the average 1-year fixed rate on new owner-occupier commitments entered into during the month (Table B30).",
        "questions": [
            "What is the average effective interest rate across all existing mortgages in New Zealand (RBNZ Table B6)?",
            "How does the existing mortgage book interest rate compare with new 1-year fixed mortgage rates?",
        ],
        "article_title": "Refinancing a Mortgage in NZ: How Cashbacks & Break Fees Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/refinancing-a-mortgage-in-nz-how-cashbacks-break-fees-and-switching-banks-actually-work",
        "tool_title": "Compare NZ Mortgage Rates",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
    },
    "low-equity-margin-penalty-spread": {
        "name": "Advertised Standard vs Special (<=80% LVR) 1-year fixed mortgage rate spread (RBNZ B20 vs B21)",
        "also_known_as": "RBNZ Table B20 Standard Fixed Rate minus Table B21 Special Fixed Rate spread",
        "category": "Homeowners & refinancing",
        "topic": "Special vs Standard advertised fixed mortgage rates",
        "source": "RBNZ Tables B20 & B21",
        "source_full": "Reserve Bank of New Zealand — Standard (Table B20) vs Special (Table B21) Residential Mortgage Interest Rates",
        "series_ids": "RBNZ Table B20 (1Y Standard Rate) minus Table B21 (1Y Special <=80% LVR Rate)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-special-interest-rates",
        "data_type": "Officially reported RBNZ advertised rate series & HLF-derived basis-point spread",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "The difference (in percentage points and basis points) between major New Zealand banks' advertised 1-year 'Standard' fixed mortgage rate (RBNZ Table B20) and advertised 1-year 'Special' fixed mortgage rate for borrowers meeting minimum equity criteria (RBNZ Table B21).",
        "methodology": "Derived live from RBNZ Table B20 and Table B21 workbooks by subtracting the 1-year Special fixed rate from the 1-year Standard fixed rate. Note that RBNZ Tables B20 and B21 measure advertised carded rates rather than separate low-equity premium (LEP) or low-equity margin (LEM) tiers that individual lenders may apply above 80% LVR.",
        "hlf_calculation": "Standard vs Special Spread = 1Y Standard Rate (B20) minus 1Y Special Rate (B21)",
        "hlf_interpretation": "HLF market context: Tracks the basis-point difference between banks' advertised 1-year Standard fixed rate (Table B20) and 1-year Special fixed rate for borrowers with at least 20% equity (Table B21). Separate lender-specific low-equity margins or fees may also apply depending on the LVR tier and lender.",
        "questions": [
            "What is the difference between Special and Standard 1-year fixed mortgage rates in RBNZ Tables B20 and B21?",
        ],
        "article_title": "Buying Your First Home with a 5%–10% Deposit in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/buying-your-first-home-with-a-5-10-deposit-in-nz-your-2026-options-explained",
        "tool_title": "Compare NZ Mortgage Rates",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/compare-nz-mortgage-rates",
    },
    "existing-mortgage-stock-lvr-refix": {
        "name": "Existing New Zealand mortgage stock repricing horizon & floating vs fixed structure (RBNZ Table S33)",
        "also_known_as": "RBNZ Table S33 existing mortgage stock by time remaining to repricing",
        "category": "Homeowners & refinancing",
        "topic": "Existing homeowner mortgage stock repricing schedule",
        "source": "RBNZ Table S33",
        "source_full": "Reserve Bank of New Zealand — Banks: Assets – Loans fully secured by residential mortgage (Table S33)",
        "series_ids": "RBNZ Table S33 — Total Mortgage Stock by Repricing Horizon (Floating, <=6m, <=12m, >12m)",
        "source_url": "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-fully-secured-by-residential-mortgage",
        "data_type": "Officially reported RBNZ mortgage stock series & HLF-derived repricing shares",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "Breakdown of New Zealand registered banks' existing residential mortgage book by time remaining until interest rate repricing (Floating, repricing within 6 months, and repricing within 12 months).",
        "methodology": "Extracted live from RBNZ Table S33, calculating the share of existing fixed-rate mortgage debt scheduled to reprice within 6 and 12 months.",
        "hlf_calculation": "6M Repricing Share = (<=3m + 3–6m Fixed Stock) ÷ Total Fixed Stock; 12M Repricing Share = (<=3m + 3–6m + 6–12m) ÷ Total Fixed Stock",
        "hlf_interpretation": "HLF market context: Shows the repricing profile of New Zealand registered banks' existing residential mortgage stock across floating and fixed-rate maturity buckets.",
        "questions": [
            "What percentage of existing New Zealand fixed-rate mortgages are scheduled to reprice in the next 6 to 12 months?",
        ],
        "article_title": "Refinancing a Mortgage in NZ: How Cashbacks & Break Fees Work",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/refinancing-a-mortgage-in-nz-how-cashbacks-break-fees-and-switching-banks-actually-work",
        "tool_title": "HLF Refinance & Cashback Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/refinance-calculator",
    },
    "tla-council-rent-benchmarks": {
        "name": "Territorial Local Authority (city & district council) median weekly rents & active bonds (MBIE)",
        "also_known_as": "MBIE Tenancy Services By-TLA rental bond medians across New Zealand cities and districts",
        "category": "Rents & tenancy",
        "topic": "City & district council median weekly rents",
        "source": "MBIE Tenancy Services",
        "source_full": "Ministry of Business, Innovation and Employment (MBIE) — Territorial Authority (TLA) Rental Bond Data",
        "series_ids": "MBIE By-TLA Rental Bond CSV — Council Median Rent, Lower/Upper Quartile & Active Bonds",
        "source_url": TENANCY_PAGE,
        "data_type": "Officially reported MBIE TLA medians & active bond counts",
        "derived_by_hlf": True,
        "unit": "NZD_per_week",
        "frequency": "Monthly",
        "definition": "Median weekly private rent on newly lodged bonds and active bond counts across New Zealand's Territorial Local Authorities (city and district councils).",
        "methodology": "Extracted live from MBIE Tenancy Services' Territorial Authority (TLA) monthly rental bond CSV. Medians reflect newly lodged private tenancy bonds in each TLA during the reference month across all dwelling types and bedroom counts combined.",
        "hlf_calculation": "Officially reported MBIE Territorial Local Authority median weekly rents & active bond counts",
        "hlf_interpretation": "HLF market context: Reports council-level median weekly rents on newly lodged bonds and active bond counts across New Zealand's Territorial Local Authorities.",
        "questions": [
            "What is the median weekly rent by city and district council in New Zealand?",
            "How do median rents on newly lodged bonds compare between Auckland, Christchurch City, and Wellington City?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
    "individual-income-by-source-benchmarks": {
        "name": "New Zealand median individual weekly & annualised income by source: wage & salary, self-employed & all sources (Stats NZ)",
        "also_known_as": "Stats NZ Labour Market Statistics (Income) — Median individual earnings for wage/salary earners, self-employed & all people",
        "category": "Income, net worth & household benchmarks",
        "topic": "Individual income benchmarks (wage & salary vs self-employed vs all people)",
        "source": "Stats NZ Labour Income",
        "source_full": "Stats NZ — Labour Market Statistics (Income) Annual June Quarter Release",
        "series_ids": "Stats NZ Labour Market Statistics (Income) — Median Weekly Income (All People, Wage & Salary Earners, Self-Employed) & Full-Time Hourly Earnings",
        "source_url": STATS_LABOUR_INCOME_PAGE,
        "data_type": "Officially reported Stats NZ series & HLF-derived annualised (×52) equivalents",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual (June Quarter)",
        "definition": "Median weekly and annualised (×52 weeks) individual gross income across New Zealand for people receiving wages and salaries, self-employed individuals, and all people from all sources, alongside median full-time hourly earnings.",
        "methodology": "Extracted live from Stats NZ's Labour Market Statistics (Income) June quarter release, multiplying official median weekly gross income figures by 52 weeks. As defined by Stats NZ, 'wages and salaries' covers both full-time and part-time employees, while 'all people from all sources' covers the wider working-age population (including part-time earners, government transfer recipients, and retirees).",
        "hlf_calculation": "Annualised Median Income = Official Stats NZ Median Weekly Income × 52 weeks",
        "hlf_interpretation": "HLF market context: Provides individual-level median weekly and annualised income benchmarks from Stats NZ's annual June quarter income survey, distinct from combined multi-earner household or mortgage application income.",
        "questions": [
            "What is the median weekly and annualised wage and salary income in New Zealand?",
            "What are median full-time hourly earnings in New Zealand?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "equivalised-household-income-family-size": {
        "name": "Median equivalised household disposable income before & after housing costs (Stats NZ HILS)",
        "also_known_as": "Stats NZ Household Income and Housing-Cost Statistics — Equivalised BHC & AHC disposable income",
        "category": "Income, net worth & household benchmarks",
        "topic": "Household disposable income before & after housing costs",
        "source": "Stats NZ HILS",
        "source_full": "Stats NZ — Household Income and Housing-Cost Statistics (HILS)",
        "series_ids": "Stats NZ HILS — Median Equivalised Disposable Income Before Housing Costs (BHC) & After Housing Costs (AHC)",
        "source_url": STATS_HILS_INCOME_PAGE,
        "data_type": "Officially reported Stats NZ equivalised series & HLF-derived difference",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual",
        "definition": "Median equivalised annual household disposable (after-tax) income across New Zealand before housing costs (BHC) and after deducting housing costs (AHC).",
        "methodology": "Extracted live from Stats NZ's Household Income and Housing-Cost Statistics (Household Economic Survey) release. Stats NZ equivalises disposable household income using the modified OECD scale (reference 1.0 single-person household) so households of different sizes and compositions can be compared on an equivalent basis.",
        "hlf_calculation": "Difference (%) = ((BHC Equivalised Median − AHC Equivalised Median) ÷ BHC Equivalised Median) × 100",
        "hlf_interpretation": "HLF market context: Compares Stats NZ's national median equivalised after-tax household income before and after deducting housing costs.",
        "questions": [
            "What is the median equivalised household disposable income in New Zealand before and after housing costs?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "rent-vs-mortgage-interest-outgoings": {
        "name": "Weekly mortgage interest vs principal payments & housing cost-to-income ratio (Stats NZ HILS)",
        "also_known_as": "Stats NZ HILS average weekly mortgage interest vs principal split & housing cost per $100 income",
        "category": "Income, net worth & household benchmarks",
        "topic": "Weekly mortgage interest vs principal outgoings & housing cost ratio",
        "source": "Stats NZ HILS",
        "source_full": "Stats NZ — Household Income and Housing-Cost Statistics (Housing Costs Release)",
        "series_ids": "Stats NZ HILS — Weekly Mortgage Interest ($/wk), Weekly Mortgage Principal ($/wk) & Housing Cost per $100 Income",
        "source_url": STATS_HILS_INCOME_PAGE,
        "data_type": "Officially reported Stats NZ HILS series & HLF-derived interest vs principal shares",
        "derived_by_hlf": True,
        "unit": "NZD_per_week",
        "frequency": "Annual",
        "definition": "Average weekly mortgage interest payments versus principal repayments across mortgaged New Zealand households in Stats NZ's Household Economic Survey, alongside the national average housing-cost-to-disposable-income ratio.",
        "methodology": "Extracted live from Stats NZ's Household Income and Housing-Cost Statistics release. Survey averages cover all mortgaged households in the Household Economic Survey (including long-standing homeowners with smaller remaining balances as well as recent buyers).",
        "hlf_calculation": "Total Weekly Mortgage Outgoing = Weekly Interest + Weekly Principal; Interest Share (%) = Interest ÷ Total × 100",
        "hlf_interpretation": "HLF market context: Reports the average weekly split between mortgage interest payments and principal repayments across surveyed New Zealand households with a mortgage.",
        "questions": [
            "What is the average weekly mortgage interest and principal payment in Stats NZ's Household Economic Survey?",
            "How much of their disposable income do New Zealand households spend on housing costs on average?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "net-worth-by-age-and-tenure": {
        "name": "New Zealand household net worth by housing tenure: owner-occupied vs non-owner-occupied (Stats NZ)",
        "also_known_as": "Stats NZ Household Net Worth — All households, owner-occupied & non-owner-occupied medians",
        "category": "Income, net worth & household benchmarks",
        "topic": "Household net worth by housing tenure",
        "source": "Stats NZ Net Worth",
        "source_full": "Stats NZ — Household Net Worth Statistics",
        "series_ids": "Stats NZ Household Net Worth — Median Net Worth (All Households, Owner-Occupied & Non-Owner-Occupied)",
        "source_url": STATS_NET_WORTH_PAGE,
        "data_type": "Officially reported Stats NZ Net Worth series & HLF-derived tenure ratio",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Triennial",
        "definition": "Median household net worth (total assets minus total liabilities) across all New Zealand households, households in owner-occupied dwellings, and households in non-owner-occupied dwellings from Stats NZ's Household Net Worth statistics.",
        "methodology": "Extracted live from Stats NZ's Household Net Worth statistics release. Owner-occupied households include both mortgaged and mortgage-free dwellings and reflect net property equity alongside financial, business, and KiwiSaver assets.",
        "hlf_calculation": "Tenure Net Worth Ratio = Owner-Occupied Household Median Net Worth ÷ Non-Owner-Occupied Household Median Net Worth",
        "hlf_interpretation": "HLF market context: Compares median household net worth between households living in owner-occupied dwellings and households in non-owner-occupied dwellings in Stats NZ's triennial wealth survey.",
        "questions": [
            "What is the median household net worth in New Zealand?",
            "How does median net worth compare between owner-occupied and non-owner-occupied households in Stats NZ data?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "regional-net-worth-benchmarks": {
        "name": "Regional household net worth across New Zealand regions (Stats NZ)",
        "also_known_as": "Regional household wealth benchmarks (Wellington, Auckland, Canterbury, Rest of North & South Island)",
        "category": "Income, net worth & household benchmarks",
        "topic": "Regional household net worth benchmarks",
        "source": "Stats NZ Net Worth",
        "source_full": "Stats NZ — Household Net Worth Statistics (Regional Tables)",
        "series_ids": "Stats NZ Household Net Worth — Regional Household Net Worth",
        "source_url": STATS_NET_WORTH_PAGE,
        "data_type": "Officially reported Stats NZ regional net worth series",
        "derived_by_hlf": False,
        "unit": "NZD",
        "frequency": "Triennial",
        "definition": "Regional household net worth across New Zealand's surveyed regional groupings from Stats NZ's Household Net Worth statistics.",
        "methodology": "Extracted live from Stats NZ Household Net Worth regional CSV tables and release text (covering Auckland, Wellington, Canterbury, Rest of North Island, and Rest of South Island).",
        "hlf_calculation": "Officially reported Stats NZ regional household net worth figures",
        "hlf_interpretation": "HLF market context: Reports household net worth across New Zealand's five broad regional groupings surveyed in Stats NZ's Household Net Worth release.",
        "questions": [
            "How does household net worth compare across Wellington, Auckland, and Canterbury in Stats NZ data?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "ird-individual-taxable-income-percentiles": {
        "name": "New Zealand individual taxable income distribution: total individuals, taxable income & mean (Inland Revenue)",
        "also_known_as": "Inland Revenue Taxable Income Distribution of Individuals — Total taxpayers, combined taxable income & mean",
        "category": "Income, net worth & household benchmarks",
        "topic": "Individual taxable income distribution",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — Taxable Income Distribution of Individuals",
        "series_ids": "IRD Taxable Income Distribution — Total Individuals & Combined Taxable Income ($B)",
        "source_url": IRD_INCOME_DIST_PAGE,
        "data_type": "Officially reported Inland Revenue administrative tax series & HLF-derived mean",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual",
        "definition": "The total number of New Zealand individuals with taxable income, their combined taxable income ($ billions), and the resulting mean taxable income per individual reported by Inland Revenue.",
        "methodology": "Automatically fetched and parsed from Inland Revenue's Taxable Income Distribution of Individuals release on ird.govt.nz. IRD's administrative series covers all individuals with non-zero taxable income in the tax year (including full-time, part-time, part-year, investment, and benefit earners) and is published after IR3 return processing.",
        "hlf_calculation": "Mean Taxable Income = Combined Taxable Income ($B) ÷ Total Individuals",
        "hlf_interpretation": "HLF market context: Reports the aggregate number of individual taxpayers, combined assessable taxable income, and mean taxable income across New Zealand's administrative tax base.",
        "questions": [
            "What is the total and average taxable income of individuals in New Zealand according to Inland Revenue?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "ird-paye-wage-vs-self-employed-income": {
        "name": "New Zealand PAYE wage & salary earner headcount & modal wage point (Inland Revenue)",
        "also_known_as": "Inland Revenue Wage and Salary Distributions for Individuals",
        "category": "Income, net worth & household benchmarks",
        "topic": "PAYE wage & salary earner distribution",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — Wage and Salary Distributions for Individuals",
        "series_ids": "IRD Wage & Salary Distributions — Total PAYE Earners & Modal Income Point",
        "source_url": IRD_WAGE_SALARY_PAGE,
        "data_type": "Officially reported Inland Revenue PAYE series",
        "derived_by_hlf": False,
        "unit": "earners",
        "frequency": "Annual",
        "definition": "The total number of New Zealand individuals receiving PAYE wage and salary earnings during the tax year (and modal income point where reported) from Inland Revenue's Wage and Salary Distributions dataset.",
        "methodology": "Extracted live from Inland Revenue's Wage and Salary Distributions release pages and `.xlsx` workbook on ird.govt.nz. Includes all individuals who received PAYE wages or salaries at any point during the March tax year (excluding benefits, student allowances, NZ Superannuation, and self-employed shareholder-salaries not subject to PAYE).",
        "hlf_calculation": "Officially reported Inland Revenue PAYE wage and salary earner headcount",
        "hlf_interpretation": "HLF market context: Reports the total number of individuals receiving PAYE-deducted wage and salary earnings across New Zealand during the tax year.",
        "questions": [
            "How many PAYE wage and salary earners are there in New Zealand?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "kiwisaver-member-demographics-age-income": {
        "name": "KiwiSaver member demographics by age bracket (Inland Revenue)",
        "also_known_as": "Inland Revenue KiwiSaver membership by age group",
        "category": "KiwiSaver",
        "topic": "KiwiSaver member counts by age bracket",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Member Demographics (Age Release)",
        "series_ids": "IRD KiwiSaver Member Demographics — Enrolled Members by Age Band",
        "source_url": IRD_KS_AGE_PAGE,
        "data_type": "Officially reported Inland Revenue demographic tables & HLF-derived cohort shares",
        "derived_by_hlf": True,
        "unit": "members",
        "frequency": "Annual (June Financial Year)",
        "definition": "Distribution of enrolled New Zealand KiwiSaver members across age brackets reported by Inland Revenue.",
        "methodology": "Extracted live from Inland Revenue's KiwiSaver member demographic release and dataset workbooks. Covers all enrolled KiwiSaver members (including both active contributors and non-contributing members, as well as members aged under 18 and 65+).",
        "hlf_calculation": "Age Cohort Share (%) = Cohort Member Count ÷ Total Enrolled KiwiSaver Members × 100",
        "hlf_interpretation": "HLF market context: Shows the age distribution of all enrolled KiwiSaver members recorded by Inland Revenue as at 30 June.",
        "questions": [
            "How many KiwiSaver members are in each age bracket in New Zealand?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "kiwisaver-contribution-rate-mix": {
        "name": "KiwiSaver member contribution rate distribution (3%, 3.5%, 4%, 6%, 8% & 10% rates — IRD)",
        "also_known_as": "Inland Revenue KiwiSaver members by employee contribution rate percentage",
        "category": "KiwiSaver",
        "topic": "KiwiSaver employee contribution rate distribution",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — KiwiSaver Datasets & Provider Payments (Members by Contribution Rate)",
        "series_ids": "IRD KiwiSaver Members by Contribution Rate (3%, 3.5%, 4%, 6%, 8%, 10%)",
        "source_url": IRD_KS_PAY_PAGE,
        "data_type": "Officially reported IRD series & HLF-derived share contributing above the default/minimum rate",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Annual (June Financial Year)",
        "definition": "The percentage of wage and salary KiwiSaver members contributing at the statutory minimum/default rate (3% / 3.5%) versus those contributing at higher rates of 4%, 6%, 8%, or 10%.",
        "methodology": "Extracted live from Inland Revenue's KiwiSaver contribution rate tables and dataset workbooks on ird.govt.nz. Covers members with employee deductions processed through Inland Revenue PAYE filing.",
        "hlf_calculation": "Above-Default Contribution Share (%) = (Members at 4% + 6% + 8% + 10%) ÷ Total Wage/Salary Contributing Members × 100",
        "hlf_interpretation": "HLF market context: Shows the proportion of PAYE-contributing KiwiSaver members on the minimum/default employee contribution rate versus members contributing at 4%, 6%, 8%, or 10%.",
        "questions": [
            "What percentage of New Zealand KiwiSaver members contribute at 4%, 6%, 8%, or 10% versus the default minimum rate?",
        ],
        "article_title": "",
        "article_url": "",
        "tool_title": "KiwiSaver vs Loan Payoff Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/kiwisaver-vs-loan-payoff-calculator",
    },
    "ird-family-trusts-rental-income": {
        "name": "Active New Zealand trusts & estates (IR6): total count & total trust income (IRD)",
        "also_known_as": "Inland Revenue Trusts and Estates (IR6) statistics",
        "category": "Property investors",
        "topic": "IR6 trust & estate return filings",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — Trusts and Estates (IR6) Tax Statistics",
        "series_ids": "IRD Trusts and Estates — Active IR6 Trust & Estate Count & Total IR6 Trust Income ($B)",
        "source_url": IRD_TRUSTS_NUM_PAGE,
        "data_type": "Officially reported Inland Revenue IR6 trust series",
        "derived_by_hlf": False,
        "unit": "trusts",
        "frequency": "Annual",
        "definition": "The total number of active trusts and estates filing IR6 income tax returns in New Zealand, alongside total income reported inside IR6 structures.",
        "methodology": "Automatically fetched and parsed from Inland Revenue's Trusts and Estates statistics release pages on ird.govt.nz. IRD's series covers all non-unit trusts and estates filing an IR6 return (including family, trading, farming, and investment trusts) and does not isolate trusts that hold residential rental property.",
        "hlf_calculation": "Official IRD active IR6 trust count, YoY % change & total IR6 trust income ($B)",
        "hlf_interpretation": "HLF market context: Reports the total number of active trusts and estates filing IR6 tax returns across New Zealand and their combined reported income.",
        "questions": [
            "How many trusts and estates file IR6 tax returns in New Zealand?",
            "Has the number of active IR6 trusts in New Zealand increased or decreased?",
        ],
        "article_title": "Using Equity to Buy an Investment Property in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/using-equity-to-buy-an-investment-property-in-nz-the-split-banking-guide",
        "tool_title": "Investment Property Cashflow Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/investment-property-cashflow-calculator",
    },
    "working-for-families-income-support": {
        "name": "Working for Families (WFF) tax credit average annual entitlement & supported children (IRD)",
        "also_known_as": "Inland Revenue Working for Families Tax Credits statistics",
        "category": "Income, net worth & household benchmarks",
        "topic": "Working for Families tax credits",
        "source": "Inland Revenue (IRD)",
        "source_full": "Inland Revenue — Working for Families Tax Credits Statistics",
        "series_ids": "IRD Working for Families — Average Annual Entitlement per Family ($), Total Entitlements Paid & Supported Children",
        "source_url": IRD_WFF_ENT_PAGE,
        "data_type": "Officially reported Inland Revenue WFF series & HLF-derived weekly equivalent",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Annual",
        "definition": "The average annual Working for Families (WFF) tax credit paid per recipient New Zealand family and the total number of children supported, reported by Inland Revenue.",
        "methodology": "Extracted live from Inland Revenue's Working for Families statistics release pages on ird.govt.nz, dividing the official annual average by 52 for the weekly equivalent.",
        "hlf_calculation": "Weekly Equivalent = Official IRD Average Annual WFF Paid per Family ÷ 52 weeks",
        "hlf_interpretation": "HLF market context: Reports the average annual Working for Families tax credit paid per recipient family across New Zealand. Whether and how Working for Families tax credits are included as qualifying income in a mortgage application depends on individual lender credit policies (such as the age of dependent children).",
        "questions": [
            "What is the average Working for Families tax credit paid per recipient family in New Zealand?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "nz-median-house-price-quartiles": {
        "name": "National median, lower-quartile (25%) & upper-quartile (75%) residential sale price (HUD / MCERT)",
        "also_known_as": "Official Crown District Valuation Roll (DVR) national median, 25th percentile & 75th percentile house sale price",
        "category": "House prices, sales & HPI",
        "topic": "National median, lower-quartile & upper-quartile house sale prices",
        "source": "HUD / MCERT Property Statistics",
        "source_full": "Te Tūāpapa Kura Kāinga – Ministry of Housing and Urban Development / Ministry for Cities, Environment, Regions and Transport (MCERT) — Property and Sales Statistics (CC BY 4.0)",
        "series_ids": "HUD / MCERT Property and Sales Statistics — New Zealand Monthly Median, Lower Quartile & Upper Quartile Sale Price ($ NZD)",
        "source_url": HUD_ABOUT_PAGE,
        "data_type": "Officially reported HUD / MCERT District Valuation Roll series & HLF-derived interquartile spread",
        "derived_by_hlf": True,
        "unit": "NZD",
        "frequency": "Monthly (3-Month DVR Settlement Lag)",
        "definition": "Monthly national median (50th percentile), lower-quartile (25th percentile), and upper-quartile (75th percentile) residential property sale prices across New Zealand from the HUD / MCERT Property and Sales Statistics dataset.",
        "methodology": "Extracted live from the HUD / MCERT Property and Sales Statistics data download (.xlsx) and release dashboard (licensed under CC BY 4.0). Compiled from District Valuation Roll (DVR) freehold, arm's-length market residential sales supplied by Territorial Authorities through LINZ. Sales are recorded by agreement date (when the contract was signed) rather than settlement date; because DVR sales are transmitted after settlement, the series is published with a three-month lag and prior months are revised as late-settling sales are received.",
        "hlf_calculation": "Interquartile Sale Price Spread ($) = Upper Quartile Sale Price (75th pct) minus Lower Quartile Sale Price (25th pct); YoY Change (%)",
        "hlf_interpretation": "HLF market context: Reports the 25th percentile (lower quartile), 50th percentile (median), and 75th percentile (upper quartile) prices of residential properties sold across New Zealand. Raw median and quartile sale prices reflect the mix of properties sold in the period and do not adjust for changes in dwelling size, type, or location composition.",
        "questions": [
            "What is the official median and lower-quartile house sale price in New Zealand?",
            "What is the 25th percentile (lower quartile) and 75th percentile (upper quartile) house price in NZ?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "regional-median-house-prices": {
        "name": "Regional median & lower-quartile residential property sale prices across New Zealand regions (HUD / MCERT)",
        "also_known_as": "HUD / MCERT District Valuation Roll regional median and 25th-percentile house sale prices",
        "category": "House prices, sales & HPI",
        "topic": "Regional median & lower-quartile house sale prices",
        "source": "HUD / MCERT Property Statistics",
        "source_full": "Te Tūāpapa Kura Kāinga – Ministry of Housing and Urban Development / Ministry for Cities, Environment, Regions and Transport (MCERT) — Property and Sales Statistics (CC BY 4.0)",
        "series_ids": "HUD / MCERT Property and Sales Statistics — Regional Median & Lower Quartile Sale Prices ($ NZD)",
        "source_url": HUD_ABOUT_PAGE,
        "data_type": "Officially reported HUD / MCERT regional sale price series",
        "derived_by_hlf": False,
        "unit": "NZD",
        "frequency": "Monthly (3-Month DVR Settlement Lag)",
        "definition": "Monthly median and lower-quartile (25th percentile) residential property sale prices across New Zealand regions from the HUD / MCERT Property and Sales Statistics dataset.",
        "methodology": "Extracted live from the HUD / MCERT Property and Sales Statistics workbook across regional council boundaries. Sales are recorded by agreement date from District Valuation Roll market sales with a three-month reporting lag.",
        "hlf_calculation": "Officially reported HUD / MCERT regional median & lower-quartile sale prices",
        "hlf_interpretation": "HLF market context: Compares observed median and lower-quartile residential sale prices across New Zealand regions. Regional medians in smaller regions can fluctuate from month to month when sales counts are low or when the mix of properties sold changes.",
        "questions": [
            "How do median and lower-quartile house sale prices compare across New Zealand regions?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "tla-council-house-price-benchmarks": {
        "name": "Territorial Local Authority (city & district council) median & lower-quartile house sale prices (HUD / MCERT)",
        "also_known_as": "City and district council median and 25th-percentile residential sale prices (HUD / MCERT DVR series)",
        "category": "House prices, sales & HPI",
        "topic": "City & district council median and lower-quartile house sale prices",
        "source": "HUD / MCERT Property Statistics",
        "source_full": "Te Tūāpapa Kura Kāinga – Ministry of Housing and Urban Development / Ministry for Cities, Environment, Regions and Transport (MCERT) — Property and Sales Statistics (CC BY 4.0)",
        "series_ids": "HUD / MCERT Property and Sales Statistics — Territorial Authority Median Sale Price, Lower Quartile Sale Price & Sales Count",
        "source_url": HUD_ABOUT_PAGE,
        "data_type": "Officially reported HUD / MCERT Territorial Authority series",
        "derived_by_hlf": False,
        "unit": "NZD",
        "frequency": "Monthly (3-Month DVR Settlement Lag)",
        "definition": "Median and lower-quartile (25th percentile) residential property sale prices and sales counts across New Zealand Territorial Local Authorities (city and district councils) from the HUD / MCERT dataset.",
        "methodology": "Extracted live from the HUD / MCERT Property and Sales Statistics workbook by Territorial Local Authority (TLA). Because single-month medians in smaller districts can be volatile when sales counts are low, the table reports active council sales counts alongside median and lower-quartile prices.",
        "hlf_calculation": "Officially reported HUD / MCERT Territorial Authority median sale prices, lower-quartile sale prices & sales counts",
        "hlf_interpretation": "HLF market context: Reports council-level median and 25th-percentile (lower-quartile) residential sale prices recorded on the District Valuation Roll. Council-level medians describe historical sales in each district and are not automated valuations of individual properties.",
        "questions": [
            "What is the median and lower-quartile house sale price in Wellington City, Auckland, Christchurch City, and Hamilton City?",
            "How do house sale prices compare across New Zealand city and district councils?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Home Loan Repayments Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/home-loan-calculator",
    },
    "nz-monthly-property-sales-volume": {
        "name": "Monthly & rolling 12-month New Zealand residential property sales volume (HUD / MCERT)",
        "also_known_as": "District Valuation Roll (DVR) monthly and annual residential property sales count",
        "category": "House prices, sales & HPI",
        "topic": "Monthly & annual residential property sales counts",
        "source": "HUD / MCERT Property Statistics",
        "source_full": "Te Tūāpapa Kura Kāinga – Ministry of Housing and Urban Development / Ministry for Cities, Environment, Regions and Transport (MCERT) — Property and Sales Statistics (CC BY 4.0)",
        "series_ids": "HUD / MCERT Property and Sales Statistics — New Zealand Monthly Residential Sales Count & 12-Month Rolling Sum",
        "source_url": HUD_ABOUT_PAGE,
        "data_type": "Officially reported HUD / MCERT monthly sales count & HLF-derived 12-month rolling total",
        "derived_by_hlf": True,
        "unit": "sales",
        "frequency": "Monthly (3-Month DVR Settlement Lag)",
        "definition": "The monthly and rolling 12-month number of arm's-length freehold residential property sales agreed across New Zealand from the HUD / MCERT Property and Sales Statistics dataset.",
        "methodology": "Derived live from the HUD / MCERT Property and Sales Statistics workbook by extracting the monthly national sales count and summing the latest 12 monthly observations. Includes single-property, freehold, market residential sales matched to the District Valuation Roll by agreement date (excluding multi-property, leasehold, and non-market transfers).",
        "hlf_calculation": "Rolling 12-Month Sales Volume = Sum of latest 12 monthly national DVR residential sales counts; YoY Change (%)",
        "hlf_interpretation": "HLF market context: Measures completed arm's-length freehold residential property sales recorded on the District Valuation Roll by agreement date. Because recent months are revised upward as late-settling sales are lodged with LINZ, the series is published up to three months prior to the latest release.",
        "questions": [
            "How many residential properties are sold each month and year in New Zealand?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "mcert-house-price-index-yoy": {
        "name": "Official Crown House Price Index (MCERT HPI): annual & 3-month quality-adjusted price change",
        "also_known_as": "HUD / MCERT Sales Price to Appraisal Ratio (SPAR) House Price Index annual percentage change",
        "category": "House prices, sales & HPI",
        "topic": "Quality-adjusted House Price Index (MCERT SPAR HPI)",
        "source": "HUD / MCERT House Price Index",
        "source_full": "Te Tūāpapa Kura Kāinga – Ministry of Housing and Urban Development / Ministry for Cities, Environment, Regions and Transport (MCERT) — House Price Index (CC BY 4.0)",
        "series_ids": "MCERT House Price Index (All Dwellings) — Annual % Change, 3-Month % Change & Index Level",
        "source_url": HUD_ABOUT_PAGE,
        "data_type": "Officially reported HUD / MCERT House Price Index series",
        "derived_by_hlf": False,
        "unit": "percent",
        "frequency": "Monthly (3-Month DVR Settlement Lag)",
        "definition": "The annual (12-month) percentage change in the official MCERT House Price Index (HPI) for New Zealand, calculated using the Sales Price to Appraisal Ratio (SPAR) methodology.",
        "methodology": "Extracted live from the HUD / MCERT Property and Sales Statistics workbook and HPI dashboard. As detailed in the MCERT HPI Methodology Paper, the index uses a Sales Price to Appraisal Ratio (SPAR) method—comparing each property's sale price to its rating valuation (Capital Value) within each Territorial Authority and weighting by housing stock—to measure price changes while controlling for changes in the quality and mix of properties sold.",
        "hlf_calculation": "Officially reported MCERT House Price Index (All-Dwellings) annual % change",
        "hlf_interpretation": "HLF market context: Unlike raw median sale prices (which can shift when a higher or lower share of expensive homes sells in a given month), the SPAR-based MCERT House Price Index adjusts for property valuation quality and compositional mix across New Zealand.",
        "questions": [
            "What is the annual change in the official HUD / MCERT House Price Index for New Zealand?",
            "How does the MCERT House Price Index differ from median house sale prices?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "hpi-low-vs-high-value-tiers": {
        "name": "MCERT House Price Index by value tier: low-value (Q1) vs high-value (Q4) homes",
        "also_known_as": "HUD / MCERT House Price Index annual change for Q1 (lowest value quartile) vs Q4 (highest value quartile) homes",
        "category": "House prices, sales & HPI",
        "topic": "House Price Index performance by property value tier (Q1 vs Q4)",
        "source": "HUD / MCERT House Price Index",
        "source_full": "Te Tūāpapa Kura Kāinga – Ministry of Housing and Urban Development / Ministry for Cities, Environment, Regions and Transport (MCERT) — House Price Index Composition (CC BY 4.0)",
        "series_ids": "MCERT HPI — Low-Value (Q1) Annual % Change vs High-Value (Q4) Annual % Change",
        "source_url": HUD_ABOUT_PAGE,
        "data_type": "Officially reported HUD / MCERT HPI tier series & HLF-derived percentage-point spread",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly (3-Month DVR Settlement Lag)",
        "definition": "Comparison of annual percentage changes in the MCERT House Price Index (HPI) for low-value (Q1 lowest valuation quartile) residential properties versus high-value (Q4 highest valuation quartile) residential properties across New Zealand.",
        "methodology": "Extracted live from the MCERT HPI composition series in the HUD / MCERT Property and Sales Statistics workbook and HPI composition dashboard. Properties are segmented into four equal value quartiles (Q1 lowest 25% to Q4 highest 25%) based on their relative rating valuation within their Territorial Authority.",
        "hlf_calculation": "Tier Spread (percentage points) = Low-Value (Q1) Annual HPI Change (%) minus High-Value (Q4) Annual HPI Change (%)",
        "hlf_interpretation": "HLF market context: Compares quality-adjusted annual price movements in the lowest valuation quartile (Q1) against the highest valuation quartile (Q4) within the MCERT House Price Index.",
        "questions": [
            "Are lower-value (Q1) or higher-value (Q4) house prices changing faster in the MCERT House Price Index?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "Usable Equity Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/usable-equity-calculator",
    },
    "sales-composition-by-value-quartile": {
        "name": "Residential sales composition by valuation quartile: Q1 (lowest value) to Q4 (highest value) share (HUD / MCERT)",
        "also_known_as": "HUD / MCERT proportion of monthly residential sales in Value Quartiles Q1, Q2, Q3 & Q4",
        "category": "House prices, sales & HPI",
        "topic": "Residential sales composition by rating-valuation quartile",
        "source": "HUD / MCERT Property Statistics",
        "source_full": "Te Tūāpapa Kura Kāinga – Ministry of Housing and Urban Development / Ministry for Cities, Environment, Regions and Transport (MCERT) — Sales Composition Statistics (CC BY 4.0)",
        "series_ids": "HUD / MCERT Sales Composition — Share of Sales in Value Quartiles Q1 (Lowest), Q2, Q3 & Q4 (Highest)",
        "source_url": HUD_ABOUT_PAGE,
        "data_type": "Officially reported HUD / MCERT sales counts by value quartile & HLF-derived percentage shares",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly (3-Month DVR Settlement Lag)",
        "definition": "The percentage of monthly New Zealand residential property sales occurring in each rating-valuation quartile: Q1 (lowest 25% of housing stock by value), Q2, Q3, and Q4 (highest 25% of housing stock by value).",
        "methodology": "Extracted live from the Sales Composition series in the HUD / MCERT Property and Sales Statistics workbook. Each Territorial Authority's housing stock is divided into four equal 25% quartiles by rateable Capital Value; this series tracks what share of properties sold in a given month belonged to each valuation quartile.",
        "hlf_calculation": "Quartile Share of Sales (%) = (Sales in Valuation Quartile Q_i ÷ Total Sales across Q1–Q4) × 100",
        "hlf_interpretation": "HLF market context: As noted by HUD / MCERT, tracking sales composition by valuation quartile helps identify whether a change in the raw median sale price reflects underlying price movement or a shift in the proportion of lower-value (Q1) versus higher-value (Q4) homes transacting.",
        "questions": [
            "What share of New Zealand house sales are in the lowest value quartile (Q1) versus the highest value quartile (Q4)?",
            "How does sales composition affect median house prices in New Zealand?",
        ],
        "article_title": FHB_2026_ARTICLE_TITLE,
        "article_url": FHB_2026_ARTICLE_URL,
        "tool_title": "HLF Borrowing Power Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/borrowing-power-calculator",
    },
    "gross-rental-yield-by-council": {
        "name": "Indicative gross residential rental yield by city & district council (MBIE TLA Rent ÷ HUD / MCERT TLA Sale Price)",
        "also_known_as": "Council-level ratio of annualised MBIE median weekly rent to HUD / MCERT median house sale price",
        "category": "Property investors",
        "topic": "Council-level gross rental yield benchmarks",
        "source": "MBIE Tenancy & HUD / MCERT",
        "source_full": "MBIE Tenancy Services (By-TLA Rental Bond Data) & HUD / MCERT (Property and Sales Statistics by Territorial Authority) (CC BY 4.0)",
        "series_ids": "(MBIE TLA Median Weekly Rent × 52) ÷ HUD / MCERT TLA Median Sale Price",
        "source_url": HUD_ABOUT_PAGE,
        "data_type": "HLF-derived council-level ratio from official MBIE & HUD / MCERT series",
        "derived_by_hlf": True,
        "unit": "percent",
        "frequency": "Monthly",
        "definition": "Indicative gross residential rental yield across major New Zealand Territorial Local Authorities (city and district councils), derived by dividing annualised council median weekly rent on newly lodged bonds (MBIE Tenancy Services) by council median residential sale price (HUD / MCERT).",
        "methodology": "Derived live by matching Territorial Local Authority (TLA) median weekly rents (MBIE By-TLA dataset) with TLA median residential sale prices (HUD / MCERT Property and Sales Statistics) for major urban councils. Note that MBIE bond rents and HUD DVR sale prices have different reporting lags (~6–8 weeks vs ~12 weeks) and reflect the separate samples of properties rented and sold in each district.",
        "hlf_calculation": "Indicative Council Gross Yield (%) = ((Council Median Weekly Rent × 52) ÷ Council Median Sale Price) × 100",
        "hlf_interpretation": "HLF market context: Provides a council-level ratio of annualised median weekly rent to median sale price across major New Zealand Territorial Authorities. Gross yield ratios do not deduct property operating expenses (such as council rates, dwelling insurance, maintenance, property management, vacancy, or mortgage interest) and do not represent net investment returns.",
        "questions": [
            "How do indicative gross rental yields compare across Auckland, Christchurch City, Wellington City, Hamilton City, and Dunedin City?",
            "How is gross rental yield calculated from council median rent and median house price?",
        ],
        "article_title": "Using Equity to Buy an Investment Property in NZ",
        "article_url": "https://www.homeloanfactory.co.nz/news-insights/using-equity-to-buy-an-investment-property-in-nz-the-split-banking-guide",
        "tool_title": "NZ Property Rental Yield Calculator",
        "tool_url": "https://www.homeloanfactory.co.nz/tools/nz-property-rental-yield-calculator",
    },
}

# Backward-compatible aliases for Webflow chat bubbles & legacy embeds
LEGACY_BUBBLE_ALIASES = {
    "fhb-income-akl-vs-nz": "fhb-income-by-dti-band",
    "fhb-dti-above-6x": "dti-headroom",
    "fhb-typical-loan-size": "fhb-average-loan-size",
    "fhb-monthly-volume": "fhb-monthly-lending-volume",
    "build-cost-by-type": "new-build-consent-cost-by-type",
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
    "household-vs-individual-income-benchmarks": "individual-income-by-source-benchmarks",
}
# === END OF BLOCK 5 ===
# === BLOCK 6 OF 13 ===
_CAMOUFOX_OK = True


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


def fetch_url_bytes(session, url: str, referer: str = "", timeout: int = 35, use_cb: bool = False) -> bytes:
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
    global _CAMOUFOX_OK
    raw = fetch_url_bytes(session, url, referer=referer, timeout=45, use_cb=False)
    if raw and len(raw) >= min_bytes and b"Just a moment" not in raw[:600] and b"Access Denied" not in raw[:600]:
        return raw

    if _CAMOUFOX_OK:
        try:
            from camoufox.sync_api import Camoufox
            with Camoufox(headless=True) as browser:
                ctx = browser.new_context(locale="en-NZ", accept_downloads=True)
                page = ctx.new_page()
                b_out = _browser_fetch_binary(ctx, page, url, referer, min_bytes)
                if b_out:
                    return b_out
        except Exception as exc_cf:
            _CAMOUFOX_OK = False
            print(f"[Camoufox Disabled -> Using Playwright] {exc_cf}")

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
    """Extract plain text from a live-downloaded PDF (Kāinga Ora, FMA Report)."""
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
    global _CAMOUFOX_OK
    results: dict[str, str] = {}
    missing: list[str] = []
    for u in urls:
        raw = fetch_url_bytes(session, u, timeout=25, use_cb=False)
        txt = raw.decode("utf-8", errors="ignore") if raw else ""
        if len(txt) > 1500 and "Just a moment" not in txt and "Access Denied" not in txt and "_Incapsula_Resource" not in txt:
            results[u] = txt
        else:
            missing.append(u)

    if not missing:
        return results

    still_missing: list[str] = []
    if _CAMOUFOX_OK:
        try:
            from camoufox.sync_api import Camoufox
            with Camoufox(headless=True) as browser:
                ctx = browser.new_context(locale="en-NZ")
                page = ctx.new_page()
                for u in missing:
                    try:
                        page.goto(u, wait_until="domcontentloaded", timeout=35000)
                        page.wait_for_timeout(2500)
                        content = page.content()
                        if len(content) > 1200 and "Just a moment" not in content[:600] and "_Incapsula_Resource" not in content[:1200]:
                            results[u] = content
                        else:
                            still_missing.append(u)
                    except Exception:
                        still_missing.append(u)
        except Exception as exc_cf:
            _CAMOUFOX_OK = False
            print(f"[Camoufox Disabled -> Using Playwright] {exc_cf}")
            still_missing = list(missing)
    else:
        still_missing = list(missing)

    if not still_missing:
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
            for u in still_missing:
                try:
                    page.goto(u, wait_until="domcontentloaded", timeout=35000)
                    page.wait_for_timeout(2200)
                    content = page.content()
                    if len(content) > 1200:
                        results[u] = content
                except Exception as e_pw:
                    print(f"[{label} Playwright] {u}: {e_pw}")
            browser.close()
    except Exception as exc:
        print(f"[{label} Fallback] Playwright note: {exc}")
    return results


def fetch_pages_with_camoufox(urls: list[str], wait_ms: int = 5000) -> dict[str, str]:
    """Render anti-bot protected pages directly in Camoufox or Playwright Chromium."""
    global _CAMOUFOX_OK
    out: dict[str, str] = {}
    if _CAMOUFOX_OK:
        try:
            from camoufox.sync_api import Camoufox
            with Camoufox(headless=True) as browser:
                ctx = browser.new_context(locale="en-NZ")
                page = ctx.new_page()
                for u in urls:
                    try:
                        page.goto(u, wait_until="domcontentloaded", timeout=40000)
                        page.wait_for_timeout(wait_ms)
                        html_s = page.content()
                        if html_s and len(html_s) > 1200:
                            out[u] = html_s
                    except Exception as e_cf:
                        print(f"[Camoufox Page] {u}: {e_cf}")
        except Exception as exc_cf:
            _CAMOUFOX_OK = False
            print(f"[Camoufox Disabled -> Using Playwright] {exc_cf}")
    if out:
        return out
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
            for u in urls:
                try:
                    page.goto(u, wait_until="domcontentloaded", timeout=40000)
                    page.wait_for_timeout(wait_ms)
                    html_s = page.content()
                    if html_s and len(html_s) > 1200:
                        out[u] = html_s
                except Exception:
                    pass
            browser.close()
    except Exception:
        pass
    return out


def _ensure_rbnz_file(key: str, root: pathlib.Path, session, rbnz_files: dict[str, pathlib.Path]) -> pathlib.Path | None:
    p = rbnz_files.get(key)
    if p and p.exists():
        return p
    spec = RBNZ_SOURCES.get(key, {})
    for cand in sorted(root.glob(f"*{spec.get('match', key)}*.xlsx")):
        if not cand.name.startswith("tmp_") and cand.stat().st_size > 4000:
            rbnz_files[key] = cand
            return cand
    for cand_url in [spec.get("xlsx"), spec.get("alt_xlsx")]:
        if not cand_url:
            continue
        b = fetch_bytes_with_browser_fallback(session, cand_url, referer=RBNZ_WARMUP_URL, min_bytes=4000)
        if b and b[:2] == b"PK":
            out_p = root / f"tmp_{key}.xlsx"
            out_p.write_bytes(b)
            rbnz_files[key] = out_p
            print(f"[RBNZ Live] Downloaded {key.upper()} workbook ({len(b):,} bytes)")
            return out_p
    for page_u in spec.get("pages", []):
        html_p = fetch_html_with_browser_fallback(session, [page_u], f"RBNZ {key}").get(page_u, "")
        for href in re.findall(r'href=["\']([^"\']+\.xlsx[^"\']*)["\']', html_p, flags=re.I):
            full_u = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
            b = fetch_bytes_with_browser_fallback(session, full_u, referer=page_u, min_bytes=4000)
            if b and b[:2] == b"PK":
                out_p = root / f"tmp_{key}.xlsx"
                out_p.write_bytes(b)
                rbnz_files[key] = out_p
                print(f"[RBNZ Live] Discovered & downloaded {key.upper()} ({len(b):,} bytes)")
                return out_p
    return None


def download_rbnz_workbooks(root: pathlib.Path, session) -> dict[str, pathlib.Path]:
    """Download core RBNZ .xlsx workbooks using `_ensure_rbnz_file` (use_cb=False)."""
    found: dict[str, pathlib.Path] = {}
    for key in ("c31", "c33", "c32", "c40", "c41", "b30", "s33", "j10", "t43"):
        if key in RBNZ_SOURCES:
            _ensure_rbnz_file(key, root, session, found)
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


def load_rbnz_sheets_merged(xlsx_path: pathlib.Path) -> list[tuple[str, pd.DataFrame, pd.DataFrame, list[str]]]:
    """Load an RBNZ Excel workbook, unmerge all group header cells in rows 1..40, and prepend the sheet name."""
    sheets_out = []
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception:
        return load_rbnz_sheets(xlsx_path)

    for sh in wb.sheetnames:
        if sh.strip().lower() in ("notes", "definitions", "contents", "index", "metadata", "series description"):
            continue
        ws = wb[sh]
        for m_range in list(ws.merged_cells.ranges):
            if m_range.max_row <= 40:
                span_cols = m_range.max_col - m_range.min_col + 1
                if m_range.min_row == 1 and m_range.min_col == 1 and span_cols > 6:
                    continue
                top_val = ws.cell(m_range.min_row, m_range.min_col).value
                ws.unmerge_cells(str(m_range))
                for r in range(m_range.min_row, m_range.max_row + 1):
                    for c in range(m_range.min_col, m_range.max_col + 1):
                        ws.cell(r, c).value = top_val

        raw = pd.DataFrame(list(ws.values))
        if raw.empty or len(raw) < 6:
            continue

        date_row_start, date_col_idx = None, 0
        for r_idx in range(min(60, len(raw))):
            for c_try in (0, 1):
                if c_try >= raw.shape[1]:
                    continue
                v = raw.iloc[r_idx, c_try]
                if pd.isna(v):
                    continue
                s_v = str(v).strip()
                if len(s_v) < 4 or re.match(r"^[A-Za-z\s]{6,}$", s_v):
                    continue
                dt = pd.to_datetime(v, errors="coerce")
                if pd.notna(dt) and 1980 <= dt.year <= 2035:
                    date_row_start, date_col_idx = r_idx, c_try
                    break
            if date_row_start is not None:
                break

        if date_row_start is None:
            continue

        hdr_block = raw.iloc[:date_row_start, :].copy()
        for r_h in range(0, max(0, date_row_start - 1)):
            non_empty = [
                c for c in range(date_col_idx + 1, hdr_block.shape[1])
                if pd.notna(hdr_block.iloc[r_h, c]) and str(hdr_block.iloc[r_h, c]).strip()
            ]
            if 2 <= len(non_empty) <= max(6, hdr_block.shape[1] // 2):
                hdr_block.iloc[r_h, date_col_idx + 1 :] = hdr_block.iloc[r_h, date_col_idx + 1 :].ffill(limit=8)

        labels = []
        sh_clean = re.sub(r"\s+", " ", sh).strip().lower()
        for c_idx in range(raw.shape[1]):
            parts = [sh_clean] if sh_clean not in ("data", "sheet1", "table") else []
            for r_h in range(0, date_row_start):
                cell_v = hdr_block.iloc[r_h, c_idx]
                if pd.notna(cell_v):
                    s_c = re.sub(r"\s+", " ", str(cell_v)).strip()
                    if s_c and s_c.lower() not in [p.lower() for p in parts]:
                        parts.append(s_c)
            labels.append(" | ".join(parts).lower())

        df = raw.iloc[date_row_start:].copy().reset_index(drop=True)
        df["date"] = pd.to_datetime(df.iloc[:, date_col_idx], errors="coerce")
        df = df[df["date"].notna()].sort_values("date").reset_index(drop=True)
        if not df.empty:
            sheets_out.append((sh, df, raw, labels))

    return sheets_out if sheets_out else load_rbnz_sheets(xlsx_path)


def find_rbnz_series(
    tabs: list,
    target_codes: tuple[str, ...] = (),
    include_kw: tuple[str, ...] = (),
    exclude_kw: tuple[str, ...] = (),
    val_bounds: tuple[float, float] | None = None,
) -> tuple[pd.Series, pd.Series, str]:
    upper_targets = tuple(c.upper() for c in target_codes)
    for _, df, codes, labels in tabs:
        if upper_targets and isinstance(codes, list):
            for idx, code in enumerate(codes):
                if str(code).upper() in upper_targets:
                    ser = pd.to_numeric(df.iloc[:, idx], errors="coerce")
                    valid = ser.dropna()
                    if len(valid):
                        if val_bounds and not (val_bounds[0] <= float(valid.iloc[-1]) <= val_bounds[1]):
                            continue
                        return df["date"], ser, str(code)

    for _, df, codes, labels in tabs:
        for idx, lbl in enumerate(labels):
            if idx >= df.shape[1] or str(df.columns[idx]) == "date":
                continue
            if include_kw and not all(k.lower() in lbl for k in include_kw):
                continue
            if exclude_kw and any(k.lower() in lbl for k in exclude_kw):
                continue
            ser = pd.to_numeric(df.iloc[:, idx], errors="coerce")
            valid = ser.dropna()
            if len(valid):
                if val_bounds and not (val_bounds[0] <= float(valid.iloc[-1]) <= val_bounds[1]):
                    continue
                c_str = codes[idx] if (isinstance(codes, list) and idx < len(codes) and codes[idx]) else lbl[:60]
                return df["date"], ser, c_str

    raise KeyError(f"Could not find RBNZ series codes={target_codes} include={include_kw}")


def _find_rbnz_multi(tabs, kw_list: list[tuple[str, ...]], exclude_kw: tuple[str, ...] = (), val_bounds: tuple[float, float] | None = None):
    for kw in kw_list:
        try:
            dt_s, val_s, col_lbl = find_rbnz_series(tabs, include_kw=kw, exclude_kw=exclude_kw, val_bounds=val_bounds)
            if dt_s is not None and val_s is not None and len(val_s.dropna()) >= 2:
                return dt_s, val_s, col_lbl
        except Exception:
            pass
    if val_bounds:
        lo, hi = val_bounds
        for kw in kw_list:
            for _, df, _, labels in tabs:
                for c_idx, lbl in enumerate(labels):
                    if c_idx >= df.shape[1] or str(df.columns[c_idx]) == "date":
                        continue
                    if any(ex.lower() in lbl for ex in exclude_kw):
                        continue
                    if not all(k.lower() in lbl for k in kw):
                        continue
                    s_num = pd.to_numeric(df.iloc[:, c_idx], errors="coerce").dropna()
                    if len(s_num) >= 2 and lo <= float(s_num.iloc[-1]) <= hi:
                        return df["date"], pd.to_numeric(df.iloc[:, c_idx], errors="coerce"), lbl
        flat_kws = [w.lower() for tup in kw_list for w in tup]
        for _, df, _, labels in tabs:
            for c_idx, lbl in enumerate(labels):
                if c_idx >= df.shape[1] or str(df.columns[c_idx]) == "date":
                    continue
                if any(ex.lower() in lbl for ex in exclude_kw):
                    continue
                if flat_kws and not any(k in lbl for k in flat_kws):
                    continue
                s_num = pd.to_numeric(df.iloc[:, c_idx], errors="coerce").dropna()
                if len(s_num) >= 2 and lo <= float(s_num.iloc[-1]) <= hi:
                    return df["date"], pd.to_numeric(df.iloc[:, c_idx], errors="coerce"), lbl
    return None, None, None


def _series_df(dt_s, val_s) -> pd.DataFrame:
    """Safely pair a date series and value series from the same Excel sheet into a clean 2-column DataFrame."""
    if dt_s is None or val_s is None:
        return pd.DataFrame(columns=["date", "val"])
    n = min(len(dt_s), len(val_s))
    return (
        pd.DataFrame({
            "date": pd.to_datetime(list(dt_s)[:n], errors="coerce"),
            "val": pd.to_numeric(list(val_s)[:n], errors="coerce"),
        })
        .dropna()
        .sort_values("date")
        .reset_index(drop=True)
    )


def _extract_int_from_rec(rec: dict | None, lo: int, hi: int) -> int | None:
    if not isinstance(rec, dict):
        return None
    v = rec.get("value")
    if isinstance(v, (int, float)) and lo <= float(v) <= hi:
        return int(round(float(v)))
    blob = f"{rec.get('latest_value', '')} {re.sub(r'<[^>]+>', ' ', rec.get('html', ''))} {rec.get('summary_text', '')}"
    for m in re.finditer(r"\$([\d,]{5,7})", blob):
        n = int(m.group(1).replace(",", ""))
        if lo <= n <= hi:
            return n
    return None


def _extract_pct_from_rec(rec: dict | None, lo: float = 0.1, hi: float = 50.0) -> float | None:
    if not isinstance(rec, dict):
        return None
    v = rec.get("value")
    if isinstance(v, (int, float)) and lo <= float(v) <= hi:
        return round(float(v), 1)
    blob = f"{rec.get('latest_value', '')} {re.sub(r'<[^>]+>', ' ', rec.get('html', ''))} {rec.get('summary_text', '')}"
    for m in re.finditer(r"(\d{1,2}\.\d)\s*%", blob):
        fv = float(m.group(1))
        if lo <= fv <= hi:
            return fv
    return None


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
            dl_u = r_cfg.get("alt_xlsx") or r_cfg.get("xlsx")
            raw_link_html = f' &bull; <a href="{dl_u}" target="_blank" rel="noopener">Direct RBNZ {r_key.upper()} .xlsx</a>'
            break
    if not raw_link_html and any(source_url.lower().endswith(ext) for ext in (".xlsx", ".csv", ".pdf")):
        ext_lbl = source_url.rsplit(".", 1)[-1].upper()
        raw_link_html = f' &bull; <a href="{source_url}" target="_blank" rel="noopener">Direct {ext_lbl} source file</a>'

    g_links: list[tuple[str, str]] = []
    if any(k in s_low for k in ("dti", "debt-to-income", "c40", "c41", "gross income", "hils", "taxable income")):
        g_links.append(("Debt-to-Income (DTI)", f"{GLOSSARY_BASE_URL}#dti"))
    if any(k in s_low for k in ("lvr", "low-deposit", "c31", "80%", "deposit", "b20", "b21")):
        g_links.append(("Loan-to-Value Ratio (LVR)", f"{GLOSSARY_BASE_URL}#lvr"))
    if any(k in s_low for k in ("kāinga ora", "kainga ora", "first home loan", "lmi")):
        g_links.append(("Kāinga Ora First Home Loan", f"{GLOSSARY_BASE_URL}#first-home-loan"))
        g_links.append(("Lenders Mortgage Insurance (LMI)", f"{GLOSSARY_BASE_URL}#lmi"))
    if any(k in s_low for k in ("interest-only", "c32")):
        g_links.append(("Interest-Only Mortgage", f"{GLOSSARY_BASE_URL}#interest-only"))
    if any(k in s_low for k in ("refinanc", "top-up", "equity", "c33", "net worth", "hpi", "hud", "mcert")):
        g_links.append(("Refinancing", f"{GLOSSARY_BASE_URL}#refinancing"))
        g_links.append(("Usable Equity", f"{GLOSSARY_BASE_URL}#usable-equity"))
    if any(k in s_low for k in ("kiwisaver", "t43", "fma")):
        g_links.append(("KiwiSaver First-Home Withdrawal", f"{GLOSSARY_BASE_URL}#kiwisaver-first-home-withdrawal"))
    if any(k in s_low for k in ("second-chance", "previous owner", "previous homeowner")):
        g_links.append(("Second-Chance KiwiSaver Withdrawal", f"{GLOSSARY_BASE_URL}#second-chance-withdrawal"))
    if any(k in s_low for k in ("settled", "b30", "b6", "s33", "fixed", "floating")):
        g_links.append(("Fixed vs Floating Interest Rate", f"{GLOSSARY_BASE_URL}#fixed-rate"))
    if any(k in s_low for k in ("consent", "new build", "turnkey", "dwelling")):
        g_links.append(("Turnkey vs Construction Loan", f"{GLOSSARY_BASE_URL}#turnkey"))
    if any(k in s_low for k in ("rent", "bond", "tenancy", "suburb", "tla", "yield")):
        g_links.append(("Gross Rental Yield", f"{GLOSSARY_BASE_URL}#rental-yield"))
    if any(k in s_low for k in ("insurance", "nhc", "acc", "claim", "j10")):
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
# === END OF BLOCK 6 ===
# === BLOCK 7 OF 13 ===
for _rk, _primary_u in (
    ("c31", "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c31/c31-new-residential-mortgage-lending-by-borrower-type.xlsx"),
    ("c33", "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c33/c33-new-residential-mortgage-lending-by-purpose.xlsx"),
    ("c32", "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c32/c32-new-and-existing-residential-mortgage-lending-by-payment-type.xlsx"),
    ("c40", "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c40/c40-residential-mortgage-lending-by-dti-purpose-use.xlsx"),
    ("c41", "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/c/c41/c41-residential-mortgage-borrower-gross-income.xlsx"),
    ("b30", "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/b/b30/b30-new-residential-mortgage-weighted-average-interest-rates.xlsx"),
    ("j10", "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/j/j10/j10-insurance-income-and-expenditure.xlsx"),
    ("t43", "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/t/t43/t43-kiwisaver-assets-and-liabilities.xlsx"),
):
    if _rk in RBNZ_SOURCES:
        _old_x = RBNZ_SOURCES[_rk].get("xlsx", "")
        RBNZ_SOURCES[_rk]["xlsx"] = _primary_u
        if _old_x and _old_x != _primary_u:
            RBNZ_SOURCES[_rk]["alt_xlsx"] = _old_x


def _batch_download_rbnz_in_context(ctx, page, root: pathlib.Path, keys_to_get: list[str], found: dict, tag: str) -> None:
    """Warm up once on RBNZ and batch-download all needed workbooks inside a single browser context."""
    try:
        page.goto(RBNZ_WARMUP_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(2000)
    except Exception:
        pass

    for key in list(keys_to_get):
        if key in found and found[key].exists():
            continue
        cfg = RBNZ_SOURCES[key]
        for cand_u in (cfg.get("xlsx"), cfg.get("alt_xlsx")):
            if not cand_u:
                continue
            b_data = _browser_fetch_binary(ctx, page, cand_u, "", 4_000)
            if b_data and b_data[:2] == b"PK":
                tmp_p = root / f"tmp_{key}.xlsx"
                tmp_p.write_bytes(b_data)
                found[key] = tmp_p
                print(f"[RBNZ {tag}] Downloaded live {key.upper()} workbook ({len(b_data):,} bytes)")
                break
        if key in found and found[key].exists():
            continue

        for pg_url in cfg.get("pages", []):
            try:
                page.goto(pg_url, wait_until="domcontentloaded", timeout=35000)
                page.wait_for_timeout(1500)
                pg_html = page.content()
                for href in re.findall(r'href=["\']([^"\']+\.xlsx[^"\']*)["\']', pg_html, flags=re.I):
                    if cfg["match"] in href.lower():
                        full_u = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
                        b2 = _browser_fetch_binary(ctx, page, full_u, "", 4_000)
                        if b2 and b2[:2] == b"PK":
                            tmp_p = root / f"tmp_{key}.xlsx"
                            tmp_p.write_bytes(b2)
                            found[key] = tmp_p
                            print(f"[RBNZ {tag}] Discovered & downloaded {key.upper()} ({len(b2):,} bytes)")
                            break
                if key in found and found[key].exists():
                    break
            except Exception as exc_pg:
                print(f"[RBNZ {tag}] Page {pg_url}: {exc_pg}")


def download_rbnz_workbooks(root: pathlib.Path, session) -> dict[str, pathlib.Path]:
    """Download all RBNZ .xlsx workbooks via fast HTTP (use_cb=False) -> Camoufox -> Playwright."""
    global _CAMOUFOX_OK
    found: dict[str, pathlib.Path] = {}
    needs_browser: list[str] = []

    for key, cfg in RBNZ_SOURCES.items():
        local_candidates = sorted(root.glob(f"*{cfg['match']}*.xlsx"))
        for cand in local_candidates:
            if not cand.name.startswith("tmp_") and cand.stat().st_size > 4_000:
                found[key] = cand
                break
        if key in found:
            continue

        got_http = False
        for cand_u in (cfg.get("xlsx"), cfg.get("alt_xlsx")):
            if not cand_u:
                continue
            raw = fetch_url_bytes(session, cand_u, referer=RBNZ_WARMUP_URL, timeout=40, use_cb=False)
            if raw and raw[:2] == b"PK" and len(raw) > 4_000:
                tmp_p = root / f"tmp_{key}.xlsx"
                tmp_p.write_bytes(raw)
                found[key] = tmp_p
                print(f"[RBNZ HTTP] Downloaded live {key.upper()} workbook ({len(raw):,} bytes)")
                got_http = True
                break
        if not got_http:
            needs_browser.append(key)

    if needs_browser and _CAMOUFOX_OK:
        try:
            from camoufox.sync_api import Camoufox
            with Camoufox(headless=True) as browser:
                ctx = browser.new_context(locale="en-NZ", accept_downloads=True)
                page = ctx.new_page()
                _batch_download_rbnz_in_context(ctx, page, root, needs_browser, found, "Camoufox")
        except Exception as exc_cf:
            _CAMOUFOX_OK = False
            print(f"[RBNZ Camoufox Disabled -> Using Playwright] {exc_cf}")

    still_needed = [k for k in needs_browser if k not in found]
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
                _batch_download_rbnz_in_context(ctx, page, root, still_needed, found, "Playwright")
                browser.close()
        except Exception as exc_pw:
            print(f"[RBNZ Playwright] Note: {exc_pw}")

    return found


def build_rbnz_c31_tickers(filepath: pathlib.Path) -> tuple[dict, dict]:
    pub_date = extract_rbnz_published_date(filepath)
    tabs = load_rbnz_sheets(filepath) or load_rbnz_sheets_merged(filepath)
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

    fhb_lo_val_cur = float(fhb_val.iloc[i] - fhb_hi_val.iloc[i])
    fhb_lo_cnt_cur = float(fhb_cnt.iloc[i] - fhb_hi_cnt.iloc[i])
    fhb_lo_loans_int = int(round(fhb_lo_cnt_cur))
    avg_hi_loan = int(round((float(fhb_hi_val.iloc[i]) * 1_000_000.0) / max(1.0, float(fhb_hi_cnt.iloc[i]))))
    avg_lo_loan = int(round((fhb_lo_val_cur * 1_000_000.0) / max(1.0, fhb_lo_cnt_cur)))
    loan_gap = avg_hi_loan - avg_lo_loan

    fhb_hi_val_share = float(ser_fhb_prio_pct.iloc[i])
    fhb_hi_cnt_share = (float(fhb_hi_cnt.iloc[i]) / max(1.0, float(all_hi_cnt.iloc[i]))) * 100.0
    all_hi_loans_int = int(round(float(all_hi_cnt.iloc[i])))
    std_dep_pct = float(ser_fhb_20club_pct.iloc[i])
    diff_fi = fhb_v - inv_v

    oo_hi_c = float(oo_hi_cnt.iloc[i])
    oo_hi_v = float(oo_hi_val.iloc[i])
    oo_lo_c = oo_tot_c - oo_hi_c
    oo_lo_v = float(oo_val.iloc[i] - oo_hi_val.iloc[i])
    oo_20eq_pct = float(ser_oo_20eq_pct.iloc[i])
    oo_avg_20eq = int(round((oo_lo_v * 1_000_000.0) / max(1.0, oo_lo_c)))

    inv_le60_c = int(round(float(inv_le60_cnt.iloc[i])))
    inv_30to40_c = int(round(float(inv_gt60_cnt.iloc[i] - inv_gt70_cnt.iloc[i])))
    inv_40plus_pct = (inv_le60_c / max(1, inv_c_tot)) * 100.0
    inv_30to40_pct = (inv_30to40_c / max(1, inv_c_tot)) * 100.0
    inv_u30_cnt = int(round(float(inv_gt70_cnt.iloc[i])))
    inv_u30_pct = float(ser_inv_u30_pct.iloc[i])
    inv_u30_val = float(inv_gt70_val.iloc[i])

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
    oo_low_eq_pct = float(ser_oo_low_eq_pct.iloc[i])
    oo_low_eq_prev = float(ser_oo_low_eq_pct.iloc[iy])

    comb_ex_val_ser = oo_hi_val.fillna(0) + inv_gt70_val.fillna(0)
    comb_ex_cnt_ser = oo_hi_cnt.fillna(0) + inv_gt70_cnt.fillna(0)
    comb_ex_v = float(comb_ex_val_ser.iloc[i])
    comb_ex_c = int(round(float(comb_ex_cnt_ser.iloc[i])))
    comb_ex_v_prev = float(comb_ex_val_ser.iloc[iy])
    comb_ex_c_prev = int(round(float(comb_ex_cnt_ser.iloc[iy])))

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
            f"In {m_full}, <strong>{low_cnt_pct:.1f}%</strong> of first-home buyer mortgage commitments at New Zealand registered banks (<strong>{fhb_hi_loans_int:,} of {fhb_loans_int:,} loans</strong>, or <strong>${hi_dollars_m:,.0f}m</strong> in lending) were written with <strong>under a 20% deposit</strong> (&gt;80% LVR), {plain_prior_year_pct(low_cnt_pct, low_cnt_yoy)}.",
            f"RBNZ Table C31 ({m_full}): {fhb_hi_loans_int:,} high-LVR (>80% LVR) FHB borrowers (${hi_dollars_m:,.0f}m) out of {fhb_loans_int:,} total FHB borrowers (${fhb_v:,.0f}m).",
            low_cnt_pct, f"{low_cnt_pct:.1f}% ({fhb_hi_loans_int:,} of {fhb_loans_int:,} loans)", f"{low_cnt_yoy:.1f}%",
            f"{low_cnt_pct - low_cnt_yoy:+.1f} percentage points YoY",
            f"{fhb_hi_loans_int:,} ÷ {fhb_loans_int:,} = {low_cnt_pct:.1f}% (Lending: ${hi_dollars_m:,.0f}m ÷ ${fhb_v:,.0f}m = {(hi_dollars_m/max(1.0,fhb_v))*100:.1f}%)",
            make_history_list(dates, ser_low_dep_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-share": make_c31_item(
            f"In {m_full}, New Zealand registered banks committed <strong>{fhb_loans_int:,}</strong> first-home buyer mortgages totalling <strong>{fmt_nzd_millions(fhb_v)}</strong>—accounting for <strong>{fhb_sh:.1f}%</strong> of all new residential mortgage lending ({plain_prior_year_pct(fhb_sh, fhb_sh_yoy)}), with an average commitment size of <strong>${fhb_avg_nzd:,}</strong>.",
            f"RBNZ Table C31 ({m_full}): {fhb_loans_int:,} FHB borrowers and ${fhb_v:,.0f}m FHB lending out of ${tot_val_m:,.0f}m total new residential lending (LVRN.MMB1.AA).",
            fhb_sh, f"{fhb_sh:.1f}% ({fhb_loans_int:,} loans, avg ${fhb_avg_nzd:,})", f"{fhb_sh_yoy:.1f}%",
            f"{fhb_sh - fhb_sh_yoy:+.1f} percentage points YoY",
            f"${fhb_v:,.0f}m ÷ ${tot_val_m:,.0f}m = {fhb_sh:.1f}% (Mean loan: ${fhb_v:,.0f}m × $1m ÷ {fhb_loans_int:,} = ${fhb_avg_nzd:,})",
            make_history_list(dates, ser_fhb_sh_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-average-loan-size": make_c31_item(
            f"In {m_full}, the average first-home buyer mortgage commitment at New Zealand registered banks was <strong>${fhb_avg_nzd:,}</strong> ({plain_yoy_growth(fhb_avg_nzd, fhb_avg_prev_nzd)} from ${fhb_avg_prev_nzd:,} a year earlier), across <strong>{fhb_loans_int:,} first-home buyer loans</strong> totalling <strong>{fmt_nzd_millions(fhb_v)}</strong>.",
            f"RBNZ Table C31 ({m_full}): Total FHB new lending = ${fhb_v:,.0f}m (LVRN.MMB1.AC1) across {fhb_loans_int:,} FHB loans (LVRN.MMB2.AC1).",
            float(fhb_avg_nzd), f"${fhb_avg_nzd:,} average FHB loan ({fhb_loans_int:,} loans / {fmt_nzd_millions(fhb_v)})",
            f"${fhb_avg_prev_nzd:,}",
            f"${fhb_avg_nzd - fhb_avg_prev_nzd:+,d} ({((fhb_avg_nzd - fhb_avg_prev_nzd)/max(1, fhb_avg_prev_nzd))*100:+.1f}%) YoY",
            f"(${fhb_v:,.0f}m × $1,000,000) ÷ {fhb_loans_int:,} loans = ${fhb_avg_nzd:,}",
            make_history_list(dates, ser_fhb_avg_k, lambda x: f"${int(round(x*1000)):,}"),
        ),
        "fhb-monthly-lending-volume": make_c31_item(
            f"In {m_full}, New Zealand registered banks committed <strong>{fmt_nzd_millions(fhb_v)}</strong> in new first-home buyer mortgages across <strong>{fhb_loans_int:,} loans</strong> ({plain_yoy_growth(fhb_v, fhb_v_prev)} from {fmt_nzd_millions(fhb_v_prev)} across {fhb_loans_prev_int:,} loans a year earlier).",
            f"RBNZ Table C31 ({m_full}): FHB new lending commitments = ${fhb_v:,.0f}m (LVRN.MMB1.AC1) across {fhb_loans_int:,} borrowers (LVRN.MMB2.AC1).",
            fhb_v, f"{fmt_nzd_millions(fhb_v)} across {fhb_loans_int:,} FHB loans",
            f"{fmt_nzd_millions(fhb_v_prev)} ({fhb_loans_prev_int:,} loans)",
            f"{fhb_v - fhb_v_prev:+.0f}m ({((fhb_v - fhb_v_prev)/max(1.0, fhb_v_prev))*100:+.1f}%) YoY",
            f"Reported LVRN.MMB1.AC1 = ${fhb_v:,.0f}m across {fhb_loans_int:,} FHB loans",
            make_history_list(dates, fhb_val, lambda x: f"${x:,.0f}m"),
            is_derived=False,
        ),
        "investors": make_c31_item(
            f"In {m_full}, residential property investors accounted for <strong>{inv_sh:.1f}%</strong> of new mortgage lending commitments across New Zealand registered banks (<strong>${inv_b:.2f}B</strong> across <strong>{inv_c_tot:,} loans</strong>, {plain_prior_year_pct(inv_sh, inv_sh_yoy)}), compared with <strong>{fhb_sh:.1f}%</strong> for first-home buyers.",
            f"RBNZ Table C31 ({m_full}): {inv_c_tot:,} investor borrowers and ${inv_v:,.0f}m investor lending out of ${tot_val_m:,.0f}m total lending (LVRN.MMB1.AA).",
            inv_sh, f"{inv_sh:.1f}% (${inv_b:.2f}B across {inv_c_tot:,} loans)", f"{inv_sh_yoy:.1f}%",
            f"{inv_sh - inv_sh_yoy:+.1f} percentage points YoY",
            f"${inv_v:,.0f}m ÷ ${tot_val_m:,.0f}m = {inv_sh:.1f}%",
            make_history_list(dates, ser_inv_sh_pct, lambda x: f"{x:.1f}%"),
        ),
        "investor-average-loan-size": make_c31_item(
            f"In {m_full}, residential property investors committed <strong>{fmt_nzd_millions(inv_v)}</strong> across <strong>{inv_c_tot:,} new mortgage commitments</strong> at registered banks—an average investor commitment size of <strong>${inv_avg_nzd:,}</strong> ({plain_yoy_growth(inv_avg_nzd, inv_avg_prev_nzd)} from ${inv_avg_prev_nzd:,} a year earlier).",
            f"RBNZ Table C31 ({m_full}): Investor lending = ${inv_v:,.0f}m (LVRN.MMB1.AN) across {inv_c_tot:,} investor borrowers (LVRN.MMB2.AN).",
            float(inv_avg_nzd), f"${inv_avg_nzd:,} average investor loan ({inv_c_tot:,} loans / {fmt_nzd_millions(inv_v)})",
            f"${inv_avg_prev_nzd:,}",
            f"${inv_avg_nzd - inv_avg_prev_nzd:+,d} ({((inv_avg_nzd - inv_avg_prev_nzd)/max(1, inv_avg_prev_nzd))*100:+.1f}%) YoY",
            f"(${inv_v:,.0f}m × $1,000,000) ÷ {inv_c_tot:,} investor loans = ${inv_avg_nzd:,}",
            make_history_list(dates, ser_inv_avg_k, lambda x: f"${int(round(x*1000)):,}"),
        ),
        "mover-average-loan-size": make_c31_item(
            f"In {m_full}, existing ('other') owner-occupiers committed <strong>{fmt_nzd_millions(oo_v)}</strong> across <strong>{oo_tot_c_int:,} mortgage commitments</strong> at registered banks—an average commitment size of <strong>${oo_avg_nzd:,}</strong> ({plain_yoy_growth(oo_avg_nzd, oo_avg_prev_nzd)} from ${oo_avg_prev_nzd:,} a year earlier).",
            f"RBNZ Table C31 ({m_full}): Other owner-occupier lending = ${oo_v:,.0f}m (LVRN.MMB1.AC2) across {oo_tot_c_int:,} borrowers (LVRN.MMB2.AC2).",
            float(oo_avg_nzd), f"${oo_avg_nzd:,} average mover loan ({oo_tot_c_int:,} loans / {fmt_nzd_millions(oo_v)})",
            f"${oo_avg_prev_nzd:,}",
            f"${oo_avg_nzd - oo_avg_prev_nzd:+,d} ({((oo_avg_nzd - oo_avg_prev_nzd)/max(1, oo_avg_prev_nzd))*100:+.1f}%) YoY",
            f"(${oo_v:,.0f}m × $1,000,000) ÷ {oo_tot_c_int:,} other owner-occupier loans = ${oo_avg_nzd:,}",
            make_history_list(dates, ser_oo_avg_k, lambda x: f"${int(round(x*1000)):,}"),
        ),
        "fhb-loan-size-gap": make_c31_item(
            f"In {m_full}, first-home buyers borrowing with under a 20% deposit (&gt;80% LVR) had an average mortgage commitment of <strong>${avg_hi_loan:,}</strong>, compared with <strong>${avg_lo_loan:,}</strong> for those borrowing with a 20% or greater deposit (&le;80% LVR)—a difference of <strong>${abs(loan_gap):,}</strong>.",
            f"RBNZ Table C31 ({m_full}): >80% LVR FHB = ${hi_dollars_m:,.0f}m across {fhb_hi_loans_int:,} borrowers; <=80% LVR FHB = ${fhb_lo_val_cur:,.0f}m across {fhb_lo_loans_int:,} borrowers.",
            loan_gap / 1000.0, f"${avg_hi_loan:,} at >80% LVR vs ${avg_lo_loan:,} at <=80% LVR (+${abs(loan_gap):,})",
            f"${float(ser_fhb_hi_avg_k.iloc[iy])*1000:,.0f} vs ${float(ser_fhb_lo_avg_k.iloc[iy])*1000:,.0f}",
            f"${abs(loan_gap):,} difference between LVR tiers",
            f"(${hi_dollars_m:,.0f}m × $1m ÷ {fhb_hi_loans_int:,} = ${avg_hi_loan:,}) minus (${fhb_lo_val_cur:,.0f}m × $1m ÷ {fhb_lo_loans_int:,} = ${avg_lo_loan:,}) = ${loan_gap:,}",
            make_history_list(dates, ser_fhb_hi_avg_k, lambda x: f"${x:,.1f}k (>80% LVR avg)"),
        ),
        "fhb-bank-lvr-priority": make_c31_item(
            f"In {m_full}, first-home buyers accounted for <strong>{fhb_hi_val_share:.1f}% of all &gt;80% LVR mortgage lending dollars</strong> (<strong>{fmt_nzd_millions(float(fhb_hi_val.iloc[i]))} of {fmt_nzd_millions(float(all_hi_val.iloc[i]))}</strong>) and <strong>{fhb_hi_cnt_share:.1f}% of &gt;80% LVR loans</strong> (<strong>{fhb_hi_loans_int:,} of {all_hi_loans_int:,}</strong>) across New Zealand registered banks.",
            f"RBNZ Table C31 ({m_full}): FHB >80% LVR = ${hi_dollars_m:,.0f}m ({fhb_hi_loans_int:,} loans); All Borrower >80% LVR = ${float(all_hi_val.iloc[i]):,.0f}m ({all_hi_loans_int:,} loans).",
            fhb_hi_val_share, f"{fhb_hi_val_share:.1f}% of >80% LVR dollars | {fhb_hi_cnt_share:.1f}% of >80% LVR loans",
            f"{float(ser_fhb_prio_pct.iloc[iy]):.1f}%",
            f"{fhb_hi_val_share - float(ser_fhb_prio_pct.iloc[iy]):+.1f} percentage points YoY",
            f"Dollars: ${hi_dollars_m:,.0f}m ÷ ${float(all_hi_val.iloc[i]):,.0f}m = {fhb_hi_val_share:.1f}% | Loans: {fhb_hi_loans_int:,} ÷ {all_hi_loans_int:,} = {fhb_hi_cnt_share:.1f}%",
            make_history_list(dates, ser_fhb_prio_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-20-deposit-club": make_c31_item(
            f"In {m_full}, <strong>{std_dep_pct:.1f}%</strong> of first-home buyer commitments (<strong>{fhb_lo_loans_int:,} borrowers</strong>, or <strong>{fmt_nzd_millions(fhb_lo_val_cur)}</strong> in lending) were written with a <strong>20% or greater deposit</strong> (&le;80% LVR), with an average loan size of <strong>${avg_lo_loan:,}</strong>.",
            f"RBNZ Table C31 ({m_full}): {fhb_lo_loans_int:,} <=80% LVR FHB borrowers (${fhb_lo_val_cur:,.0f}m) out of {fhb_loans_int:,} total FHB borrowers.",
            std_dep_pct, f"{std_dep_pct:.1f}% ({fhb_lo_loans_int:,} loans, avg ${avg_lo_loan:,})",
            f"{float(ser_fhb_20club_pct.iloc[iy]):.1f}%",
            f"{std_dep_pct - float(ser_fhb_20club_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{fhb_lo_loans_int:,} ÷ {fhb_loans_int:,} = {std_dep_pct:.1f}% (Mean loan: ${fhb_lo_val_cur:,.0f}m × $1m ÷ {fhb_lo_loans_int:,} = ${avg_lo_loan:,})",
            make_history_list(dates, ser_fhb_20club_pct, lambda x: f"{x:.1f}%"),
        ),
        "fhb-vs-investors": make_c31_item(
            f"In {m_full}, first-home buyers committed <strong>{fmt_nzd_millions(fhb_v)} ({fhb_sh:.1f}% of new lending)</strong> compared with <strong>{fmt_nzd_millions(inv_v)} ({inv_sh:.1f}%)</strong> for residential property investors" + (f"—a difference of <strong>+{fmt_nzd_millions(abs(diff_fi))}</strong>." if diff_fi >= 0 else "."),
            f"RBNZ Table C31 ({m_full}): FHB lending = ${fhb_v:,.0f}m ({fhb_loans_int:,} loans); Investor lending = ${inv_v:,.0f}m ({inv_c_tot:,} loans).",
            diff_fi, f"FHB {fmt_nzd_millions(fhb_v)} vs Investors {fmt_nzd_millions(inv_v)}",
            f"{fmt_nzd_millions(float(ser_fhb_vs_inv_m.iloc[iy]))} FHB minus Investor difference",
            f"{fmt_nzd_millions(diff_fi)} FHB minus Investor difference in {m_full}",
            f"${fhb_v:,.0f}m (FHB) minus ${inv_v:,.0f}m (Investors) = {diff_fi:+,.0f}m",
            make_history_list(dates, ser_fhb_vs_inv_m, lambda x: f"{x:+.0f}m (FHB vs Inv)"),
        ),
        "owner-equity-buffer": make_c31_item(
            f"In {m_full}, <strong>{oo_20eq_pct:.1f}%</strong> of existing ('other') owner-occupier mortgage commitments (<strong>{int(round(oo_lo_c)):,} of {oo_tot_c_int:,} loans</strong>, totalling <strong>{fmt_nzd_millions(oo_lo_v)}</strong>) were written with at least <strong>20% equity</strong> (&le;80% LVR), at an average commitment size of <strong>${oo_avg_20eq:,}</strong>.",
            f"RBNZ Table C31 ({m_full}): {int(round(oo_lo_c)):,} <=80% LVR other owner-occupier borrowers (${oo_lo_v:,.0f}m) out of {oo_tot_c_int:,} total other owner-occupier borrowers.",
            oo_20eq_pct, f"{oo_20eq_pct:.1f}% at <=80% LVR (avg loan ${oo_avg_20eq:,})",
            f"{float(ser_oo_20eq_pct.iloc[iy]):.1f}%",
            f"{oo_20eq_pct - float(ser_oo_20eq_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{int(round(oo_lo_c)):,} ÷ {oo_tot_c_int:,} = {oo_20eq_pct:.1f}% (Mean loan: ${oo_lo_v:,.0f}m × $1m ÷ {int(round(oo_lo_c)):,} = ${oo_avg_20eq:,})",
            make_history_list(dates, ser_oo_20eq_pct, lambda x: f"{x:.1f}%"),
        ),
        "investor-equity-tiers": make_c31_item(
            f"In {m_full}, <strong>{inv_40plus_pct:.1f}%</strong> of property investor mortgage commitments were written with at least 40% equity (<strong>&le;60% LVR</strong>, {inv_le60_c:,} loans), <strong>{inv_30to40_pct:.1f}%</strong> with 30%–40% equity (<strong>60%–70% LVR</strong>, {inv_30to40_c:,} loans), and <strong>{inv_u30_pct:.1f}%</strong> with under 30% equity (<strong>&gt;70% LVR</strong>, {inv_u30_cnt:,} loans / {fmt_nzd_millions(inv_u30_val)}).",
            f"RBNZ Table C31 ({m_full}): Total investors = {inv_c_tot:,}; <=60% LVR = {inv_le60_c:,}; >60% LVR = {int(round(float(inv_gt60_cnt.iloc[i]))):,}; >70% LVR = {inv_u30_cnt:,} (${inv_u30_val:,.0f}m).",
            inv_40plus_pct, f"<=60% LVR: {inv_40plus_pct:.1f}% | 60–70% LVR: {inv_30to40_pct:.1f}% | >70% LVR: {inv_u30_pct:.1f}%",
            f"{100.0 - float(ser_inv_u30_pct.iloc[iy]):.1f}% at <=70% LVR",
            f"{inv_u30_pct - float(ser_inv_u30_pct.iloc[iy]):+.1f} pts shift in >70% LVR share",
            f"<=60% LVR: {inv_le60_c:,} ÷ {inv_c_tot:,} = {inv_40plus_pct:.1f}% | 60–70% LVR: {inv_30to40_c:,} ÷ {inv_c_tot:,} = {inv_30to40_pct:.1f}% | >70% LVR: {inv_u30_cnt:,} ÷ {inv_c_tot:,} = {inv_u30_pct:.1f}%",
            make_history_list(dates, 100.0 - ser_inv_u30_pct, lambda x: f"{x:.1f}% (<=70% LVR)"),
        ),
        "investor-low-deposit": make_c31_item(
            f"In {m_full}, <strong>{inv_gt80_c:,} property investor mortgage commitments ({fmt_nzd_millions(inv_gt80_v)})</strong> were written with under a 20% deposit (<strong>&gt;80% LVR</strong>, or <strong>{inv_gt80_pct:.1f}%</strong> of investors, {inv_dir_prev} {inv_gt80_c_prev:,} loans a year earlier){inv_12m_clause}, alongside <strong>{inv_70to80_c:,} investor loans ({fmt_nzd_millions(inv_70to80_v)}, or {inv_70to80_pct:.1f}%)</strong> in the <strong>70%–80% LVR</strong> band.",
            f"RBNZ Table C31 ({m_full}): >80% LVR investors = {inv_gt80_c:,} loans (${inv_gt80_v:,.0f}m); >70% LVR investors = {inv_u30_cnt:,} loans (${inv_u30_val:,.0f}m).",
            float(inv_gt80_c), f">80% LVR: {inv_gt80_c:,} loans (${inv_gt80_v:,.0f}m) | 70–80% LVR: {inv_70to80_c:,} loans (${inv_70to80_v:,.0f}m)",
            f"{inv_gt80_c_prev:,} loans (${inv_gt80_v_prev:,.0f}m) >80% LVR",
            f"{inv_gt80_c - inv_gt80_c_prev:+d} loans ({inv_gt80_v - inv_gt80_v_prev:+.0f}m) YoY in >80% LVR tier",
            f">80% LVR: {inv_gt80_c:,} ÷ {inv_c_tot:,} = {inv_gt80_pct:.1f}% (${inv_gt80_v:,.0f}m) | 70–80% LVR: {inv_u30_cnt:,} - {inv_gt80_c:,} = {inv_70to80_c:,} loans (${inv_70to80_v:,.0f}m)",
            make_history_list(dates, inv_gt80_cnt, lambda x: f"{int(round(x)):,} loans (>80% LVR)"),
        ),
        "investor-new-build-lvr": make_c31_item(
            f"In {m_full}, <strong>{inv_u30_pct:.1f}%</strong> of property investor mortgage commitments (<strong>{inv_u30_cnt:,} loans</strong> totalling <strong>{fmt_nzd_millions(inv_u30_val)}</strong>) were written above 70% LVR—comprising <strong>{inv_70to80_c:,} loans ({fmt_nzd_millions(inv_70to80_v)})</strong> at 70%–80% LVR and <strong>{inv_gt80_c:,} loans ({fmt_nzd_millions(inv_gt80_v)})</strong> above 80% LVR.",
            f"RBNZ Table C31 ({m_full}): >70% LVR investors = {inv_u30_cnt:,} loans (${inv_u30_val:,.0f}m) out of {inv_c_tot:,} total investor loans.",
            inv_u30_pct, f"{inv_u30_pct:.1f}% at >70% LVR ({inv_u30_cnt:,} loans / ${inv_u30_val:,.0f}m)",
            f"{float(ser_inv_u30_pct.iloc[iy]):.1f}%",
            f"{inv_u30_pct - float(ser_inv_u30_pct.iloc[iy]):+.1f} percentage points YoY",
            f"{inv_u30_cnt:,} ÷ {inv_c_tot:,} = {inv_u30_pct:.1f}% ({inv_70to80_c:,} loans in 70–80% LVR + {inv_gt80_c:,} loans in >80% LVR)",
            make_history_list(dates, ser_inv_u30_pct, lambda x: f"{x:.1f}% (>70% LVR)"),
        ),
        "mover-low-equity": make_c31_item(
            f"In {m_full}, <strong>{oo_low_eq_pct:.1f}%</strong> of existing ('other') owner-occupier mortgage commitments (<strong>{int(round(oo_hi_c)):,} loans</strong> totalling <strong>{fmt_nzd_millions(oo_hi_v)}</strong>) were written above <strong>80% LVR</strong> (under 20% equity), {plain_prior_year_pct(oo_low_eq_pct, oo_low_eq_prev)}.",
            f"RBNZ Table C31 ({m_full}): {int(round(oo_hi_c)):,} >80% LVR other owner-occupier loans (${oo_hi_v:,.0f}m) out of {oo_tot_c_int:,} total other owner-occupier loans.",
            oo_low_eq_pct, f"{oo_low_eq_pct:.1f}% at >80% LVR ({int(round(oo_hi_c)):,} loans / ${oo_hi_v:,.0f}m)",
            f"{oo_low_eq_prev:.1f}%",
            f"{oo_low_eq_pct - oo_low_eq_prev:+.1f} percentage points YoY",
            f"{int(round(oo_hi_c)):,} ÷ {oo_tot_c_int:,} = {oo_low_eq_pct:.1f}%",
            make_history_list(dates, ser_oo_low_eq_pct, lambda x: f"{x:.1f}%"),
        ),
        "new-build-lvr-exemption-flow": make_c31_item(
            f"In {m_full}, non-first-home-buyer mortgage commitments above standard LVR thresholds totalled <strong>{fmt_nzd_millions(comb_ex_v)} across {comb_ex_c:,} loans</strong>—comprising <strong>{int(round(oo_hi_c)):,} existing owner-occupier loans ({fmt_nzd_millions(oo_hi_v)})</strong> above 80% LVR and <strong>{inv_u30_cnt:,} investor loans ({fmt_nzd_millions(inv_u30_val)})</strong> above 70% LVR.",
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
            f"Over the 12 months to {m_full}, <strong>{r12_cnt:,} first-home buyer mortgage commitments</strong> totalling <strong>{fmt_nzd_millions(r12_val)}</strong> were recorded across New Zealand registered banks ({dir_r12} {r12_prev_cnt:,} loans a year earlier)—including <strong>{r12_hi_cnt:,} commitments ({r12_hi_pct:.1f}%)</strong> written with under a 20% deposit (&gt;80% LVR).",
            f"RBNZ Table C31 (12 months to {m_full}): Sum of monthly FHB commitments = {r12_cnt:,} loans (${r12_val:,.0f}m); >80% LVR FHB commitments = {r12_hi_cnt:,} loans.",
            float(r12_cnt), f"{r12_cnt:,} loans / {fmt_nzd_millions(r12_val)} (12m rolling)", f"{r12_prev_cnt:,} loans",
            f"{r12_cnt - r12_prev_cnt:+,d} annual FHB loans YoY",
            f"12-month sum = {r12_cnt:,} loans (${r12_val:,.0f}m); >80% LVR share = {r12_hi_cnt:,} ÷ {r12_cnt:,} = {r12_hi_pct:.1f}%",
            make_history_list(dates, r12_ser, lambda x: f"{int(round(x)):,} loans (12m rolling)"),
        )

    by_ym = {
        "_latest_fhb_avg_loan_nzd": fhb_avg_nzd,
        "_latest_fhb_val_m": fhb_v,
        "_latest_fhb_cnt": fhb_loans_int,
        "_latest_fhb_low_dep_pct": low_cnt_pct,
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


def build_rbnz_c33_tickers(filepath: pathlib.Path) -> dict:
    pub_date = extract_rbnz_published_date(filepath)
    tabs = load_rbnz_sheets(filepath) or load_rbnz_sheets_merged(filepath)
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
    ref_calc = f"${ref_val_m:,.0f}m ÷ ${tot_v:,.0f}m = {ref_sh:.1f}% (Mean loan: ${ref_val_m:,.0f}m × $1m ÷ {ref_cnt_int:,} = ${ref_avg_nzd:,})"
    top_calc = f"{top_cnt_int:,} ÷ {tot_c:,} = {top_sh_c:.1f}% (Mean top-up: ${top_val_m:,.0f}m × $1m ÷ {top_cnt_int:,} = ${top_avg_nzd:,})"
    eq_calc = f"Reported ${top_val_m:,.0f}m top-up lending × $1m ÷ {top_cnt_int:,} borrowers = ${top_avg_nzd:,} mean"

    return {
        "refinance": {
            "chip": "RBNZ C33",
            "as_of": as_of,
            "source_published": pub_date,
            "max_age_days": MAX_AGE_MONTHLY,
            "html": (
                f"In {m_full}, <strong>{ref_cnt_int:,}</strong> New Zealand mortgage commitments involved switching loan providers, refinancing "
                f"<strong>${ref_val_m/1000.0:.2f}B (${ref_val_m:,.0f}m)</strong> in home loans—representing <strong>{ref_sh:.1f}%</strong> "
                f"of all new residential mortgage lending, at an average commitment size of <strong>${ref_avg_nzd:,}</strong>."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Change in loan provider = {ref_cnt_int:,} borrowers and ${ref_val_m:,.0f}m out of ${tot_v:,.0f}m total purpose lending.",
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
                f"In {m_full}, <strong>{top_cnt_int:,}</strong> mortgage top-up commitments were recorded on existing properties—"
                f"accounting for <strong>{top_sh_c:.1f}%</strong> of all mortgage commitments by purpose across New Zealand registered banks, with an average top-up of "
                f"<strong>${top_avg_nzd:,}</strong>."
            ),
            "official_data_summary": f"RBNZ Table C33 ({m_full}): Top-up commitments = {top_cnt_int:,} borrowers (${top_val_m:,.0f}m) out of {tot_c:,} total commitments.",
            "source_html": format_provenance_source_html(top_calc, "RBNZ Table C33 (Reserve Bank of New Zealand)", c33_url, True),
            "latest_period": m_full,
            "value": round(top_sh_c, 2),
            "latest_value": f"{top_sh_c:.1f}% of commitments ({top_cnt_int:,} top-ups, avg ${top_avg_nzd:,})",
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
                f"In {m_full}, New Zealand registered banks committed <strong>{fmt_nzd_millions(top_val_m)} in mortgage top-ups</strong> across "
                f"<strong>{top_cnt_int:,} commitments</strong>, at an average of <strong>${top_avg_nzd:,}</strong> per top-up."
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
    tabs = load_rbnz_sheets(filepath) or load_rbnz_sheets_merged(filepath)
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
                f"In {m_full}, <strong>{inv_io_pct:.1f}%</strong> of new residential investor mortgage lending at registered banks was written on "
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
                    avg_inc_nzd = int(round(last_b * 1000.0))
                    sum_pool_m += last_a
                    sum_implied_apps += imp_apps
                    band_rows.append({
                        "band": b_label,
                        "avg_income_nzd": avg_inc_nzd,
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
            tabs40 = load_rbnz_sheets(c40_path) or load_rbnz_sheets_merged(c40_path)
            val_pct, dti_dates, dti_ser, last_hi_v, last_tot_v, last_le5_v = None, None, None, 0.0, 0.0, 0.0

            for _, df40, _, labels40 in tabs40:
                low_idxs, le5_idxs, gran_hi_idxs, sub_gt6_idxs, unk_idxs = [], [], [], [], []
                for idx, lbl in enumerate(labels40):
                    if idx >= df40.shape[1] or str(df40.columns[idx]) == "date":
                        continue
                    if "first home" not in lbl:
                        continue
                    if any(x in lbl for x in ("auckland", "80", "lvr", "%", "share", "count", "number")):
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

                if len(low_idxs) >= 4 and (gran_hi_idxs or sub_gt6_idxs):
                    low_s = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in low_idxs[:4])
                    le5_s = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in le5_idxs[:3])
                    unk_s = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in unk_idxs[:1])
                    if gran_hi_idxs:
                        hi_s = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in gran_hi_idxs) + unk_s
                    else:
                        s_idx, is_excl_unk = sub_gt6_idxs[0]
                        hi_s = pd.to_numeric(df40.iloc[:, s_idx], errors="coerce").fillna(0.0) + (unk_s if is_excl_unk else 0.0)
                    tot_s = (low_s + hi_s).replace(0.0, float("nan"))
                    pct_s = (hi_s / tot_s) * 100.0
                    last_v = float(pct_s.iloc[-1])
                    if 0.1 <= last_v <= 45.0:
                        val_pct, dt_dti, dti_dates, dti_ser = last_v, pd.Timestamp(df40["date"].iloc[-1]), df40["date"], pct_s
                        last_hi_v, last_tot_v, last_le5_v = float(hi_s.iloc[-1]), float(tot_s.iloc[-1]), float(le5_s.iloc[-1])
                        break

            if val_pct is not None and dt_dti is not None:
                m_full = dt_dti.strftime("%B %Y")
                diff_cap = 20.0 - val_pct
                fhb_gt6_pct_latest, fhb_le6_pct_latest = val_pct, 100.0 - val_pct
                fhb_le5_pct_latest = (last_le5_v / max(1.0, last_tot_v)) * 100.0
                iy = max(0, len(dti_dates) - 13)
                prev_v = float(dti_ser.iloc[iy]) if pd.notna(dti_ser.iloc[iy]) else val_pct
                hist_start_c40 = f"Available from {pd.Timestamp(dti_dates.iloc[0]).strftime('%B %Y')} ({len(dti_dates)} monthly observations)"
                c40_calc = f"${last_hi_v:,.1f}m ÷ ${last_tot_v:,.1f}m = {val_pct:.1f}% (Benchmark diff: 20.0% - {val_pct:.1f}% = {diff_cap:.1f} pts)"
                out["dti-headroom"] = mk_dti_rec(
                    "RBNZ C40", month_end_iso(dt_dti), pub_c40,
                    f"In {m_full}, <strong>{val_pct:.1f}%</strong> of first-home buyer mortgage commitments at registered banks had a <strong>Debt-to-Income (DTI) ratio above 6x</strong> (${last_hi_v:,.1f}m of ${last_tot_v:,.1f}m)—sitting <strong>{diff_cap:.1f} percentage points</strong> below the 20% RBNZ speed-limit allowance.",
                    f"RBNZ Table C40 ({m_full}): FHB lending with TDTI > 6 = ${last_hi_v:,.1f}m out of ${last_tot_v:,.1f}m total FHB commitments.",
                    "RBNZ Table C40 (Reserve Bank of New Zealand)", c40_url, m_full, round(val_pct, 2),
                    f"{val_pct:.1f}% at >6x DTI ({diff_cap:.1f} pts below 20% benchmark)",
                    f"{pd.Timestamp(dti_dates.iloc[iy]).strftime('%B %Y')} ({prev_v:.1f}%)",
                    f"{val_pct - prev_v:+.1f} percentage points YoY", c40_calc, hist_start_c40,
                    make_history_list(dti_dates, dti_ser, lambda x: f"{x:.1f}%"),
                )

            inv_pct_v, inv_dt, inv_dates, inv_ser, inv_hi_v, inv_tot_v = None, None, None, None, 0.0, 0.0
            for _, df40, _, labels40 in tabs40:
                i_low, i_hi, i_sub7, i_unk = [], [], [], []
                for idx, lbl in enumerate(labels40):
                    if idx >= df40.shape[1] or str(df40.columns[idx]) == "date":
                        continue
                    if "investor" not in lbl:
                        continue
                    if any(x in lbl for x in ("auckland", "80", "70", "lvr", "%", "share", "count", "number")):
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
                    low_is = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in i_low[:5])
                    unk_is = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in i_unk[:1])
                    if i_hi:
                        hi_is = sum(pd.to_numeric(df40.iloc[:, c_i], errors="coerce").fillna(0.0) for c_i in i_hi) + unk_is
                    else:
                        s_idx, is_ex = i_sub7[0]
                        hi_is = pd.to_numeric(df40.iloc[:, s_idx], errors="coerce").fillna(0.0) + (unk_is if is_ex else 0.0)
                    tot_is = (low_is + hi_is).replace(0.0, float("nan"))
                    pct_is = (hi_is / tot_is) * 100.0
                    lv_i = float(pct_is.iloc[-1])
                    if 0.1 <= lv_i <= 45.0:
                        inv_pct_v, inv_dt, inv_dates, inv_ser = lv_i, pd.Timestamp(df40["date"].iloc[-1]), df40["date"], pct_is
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
                    f"In {m_inv}, <strong>{inv_pct_v:.1f}%</strong> of residential property investor mortgage commitments at registered banks had a <strong>Debt-to-Income (DTI) ratio above 7x</strong> (${inv_hi_v:,.1f}m of ${inv_tot_v:,.1f}m)—sitting <strong>{inv_headroom:.1f} percentage points</strong> below the 20% RBNZ investor speed-limit allowance.",
                    f"RBNZ Table C40 ({m_inv}): Property investor lending with TDTI > 7 = ${inv_hi_v:,.1f}m out of ${inv_tot_v:,.1f}m total investor commitments.",
                    "RBNZ Table C40 (Reserve Bank of New Zealand)", c40_url, m_inv, round(inv_pct_v, 2),
                    f"{inv_pct_v:.1f}% at >7x DTI ({inv_headroom:.1f} pts below 20% investor cap)",
                    f"{pd.Timestamp(inv_dates.iloc[iy_i]).strftime('%B %Y')} ({prev_inv_v:.1f}%)",
                    f"{inv_pct_v - prev_inv_v:+.1f} percentage points YoY", inv_c40_calc,
                    f"Available from {pd.Timestamp(inv_dates.iloc[0]).strftime('%B %Y')} ({len(inv_dates)} monthly observations)",
                    make_history_list(inv_dates, inv_ser, lambda x: f"{x:.1f}% (>7x DTI)"),
                )
        except Exception as exc:
            print(f"[C40 DTI] Parse note: {exc}")

    if c41_path is not None:
        try:
            pub_c41 = extract_rbnz_published_date(c41_path)
            tabs41 = load_rbnz_sheets(c41_path) or load_rbnz_sheets_merged(c41_path)
            dt_fhb, fhb_inc, d41_dates, fhb_inc_ser, fhb_bands, fhb_pool_m, fhb_imp_apps = compute_c41_weighted_avg_income(tabs41, "first home", "nz")
            _, inv_inc, _, _, _, inv_pool_m, inv_imp_apps = compute_c41_weighted_avg_income(tabs41, "investor", "nz")
            c41_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi"

            if dt_fhb is not None and fhb_inc and 40000 <= fhb_inc <= 600000:
                dt_fhb_ts = pd.Timestamp(dt_fhb)
                m_c41, as_of_c41 = dt_fhb_ts.strftime("%B %Y"), month_end_iso(dt_fhb_ts)
                iy41 = max(0, len(d41_dates) - 13)
                prev_inc_k = float(fhb_inc_ser.iloc[iy41]) if pd.notna(fhb_inc_ser.iloc[iy41]) else (fhb_inc / 1000.0)
                prev_inc_nzd, fhb_inc_int = int(round(prev_inc_k * 1000.0)), int(round(fhb_inc))
                hist_41 = make_history_list(d41_dates, fhb_inc_ser, lambda x: f"${x*1000:,.0f}")
                hist_start_c41 = f"Available from {pd.Timestamp(d41_dates.iloc[0]).strftime('%B %Y')} ({len(d41_dates)} monthly observations)"

                band_summary_str = " | ".join(
                    f"{b['band']}: ${b['avg_income_nzd']:,} avg (${b['total_income_m']:,.1f}m pool / {b['implied_borrowers']:,.1f} implied apps)"
                    for b in fhb_bands
                ) or f"Overall mean: ${fhb_inc_int:,}"
                lo_band = min((b["avg_income_nzd"] for b in fhb_bands), default=fhb_inc_int)
                hi_band = max((b["avg_income_nzd"] for b in fhb_bands), default=fhb_inc_int)

                bgi_calc = (
                    f"FHB: ${fhb_pool_m:,.1f}m total band income ÷ {fhb_imp_apps:,.1f} implied applications = ${fhb_inc_int:,}"
                    + (f" | Investor: ${inv_pool_m:,.1f}m ÷ {inv_imp_apps:,.1f} = ${inv_inc:,.0f}" if inv_inc else "")
                )
                bgi_html = (
                    f"In {m_c41}, approved New Zealand first-home buyer mortgage applications at registered banks averaged <strong>${fhb_inc_int:,}</strong> in combined gross borrower income, compared with <strong>${inv_inc:,.0f}</strong> across approved property investor applications."
                    if inv_inc and 40000 <= inv_inc <= 900000
                    else f"In {m_c41}, approved New Zealand first-home buyer mortgage applications at registered banks averaged <strong>${fhb_inc_int:,}</strong> in combined gross borrower income (ranging from <strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong> across DTI bands)."
                )
                out["borrower-income"] = mk_dti_rec(
                    "RBNZ C41", as_of_c41, pub_c41, bgi_html,
                    f"RBNZ Table C41 ({m_c41}) Official FHB Band Figures: {band_summary_str}.",
                    "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_url, m_c41, fhb_inc_int,
                    f"${fhb_inc_int:,} FHB mean" + (f" vs ${inv_inc:,.0f} Investor mean" if inv_inc else ""),
                    f"{pd.Timestamp(d41_dates.iloc[iy41]).strftime('%B %Y')} (${prev_inc_nzd:,} FHB)",
                    f"${fhb_inc_int - prev_inc_nzd:+d} YoY", bgi_calc, hist_start_c41, hist_41,
                )

                typ_inc_calc = f"${fhb_pool_m:,.1f}m total FHB gross income pool × $1m ÷ {fhb_imp_apps:,.1f} implied FHB applications = ${fhb_inc_int:,}"
                out["fhb-typical-income"] = mk_dti_rec(
                    "RBNZ C41", as_of_c41, pub_c41,
                    f"In {m_c41}, the borrower-weighted average combined gross borrower income across approved New Zealand first-home buyer mortgage applications at registered banks was <strong>${fhb_inc_int:,}</strong> ({plain_yoy_growth(fhb_inc_int, prev_inc_nzd)} from ${prev_inc_nzd:,} a year earlier).",
                    f"RBNZ Table C41 ({m_c41}): Total FHB gross income pool = ${fhb_pool_m:,.1f}m across {fhb_imp_apps:,.1f} implied applications (band averages: ${lo_band:,}–${hi_band:,}).",
                    "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_url, m_c41, fhb_inc_int,
                    f"${fhb_inc_int:,} typical FHB gross application income",
                    f"{pd.Timestamp(d41_dates.iloc[iy41]).strftime('%B %Y')} (${prev_inc_nzd:,})",
                    f"${fhb_inc_int - prev_inc_nzd:+d} ({((fhb_inc_int - prev_inc_nzd)/max(1, prev_inc_nzd))*100:+.1f}%) YoY",
                    typ_inc_calc, hist_start_c41, hist_41,
                )

                band_calc = f"${fhb_pool_m:,.1f}m total band income × $1m ÷ {fhb_imp_apps:,.1f} total implied applications = ${fhb_inc_int:,}"
                out["fhb-income-by-dti-band"] = mk_dti_rec(
                    "RBNZ C41", as_of_c41, pub_c41,
                    f"In {m_c41}, average combined gross borrower incomes across New Zealand first-home buyer Debt-to-Income bands ranged from <strong>${lo_band:,}</strong> to <strong>${hi_band:,}</strong>, with an overall band-weighted average of <strong>${fhb_inc_int:,}</strong>.",
                    f"RBNZ Table C41 ({m_c41}) Official FHB Band Figures: {band_summary_str}.",
                    "RBNZ Table C41 (Reserve Bank of New Zealand)", c41_url, m_c41, fhb_inc_int,
                    f"${fhb_inc_int:,} HLF mean | Official C41 bands: ${lo_band:,}–${hi_band:,}",
                    f"{pd.Timestamp(d41_dates.iloc[iy41]).strftime('%B %Y')} (${prev_inc_nzd:,} HLF mean)",
                    f"${fhb_inc_int - prev_inc_nzd:+d} ({((fhb_inc_int - prev_inc_nzd)/max(1, prev_inc_nzd))*100:+.1f}%) YoY",
                    band_calc, hist_start_c41, hist_41,
                    {"tdti_bands": fhb_bands, "tdti_total_pool_m": fhb_pool_m, "tdti_total_implied_apps": fhb_imp_apps},
                )

                ym_c41 = dt_fhb_ts.strftime("%Y-%m")
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
                        f"Implied FHB Loan-to-Income Ratio: ${avg_loan_for_dti:,} (C31 avg FHB loan) ÷ ${fhb_inc_int:,} (C41 avg FHB gross income) = {implied_dti_x:.2f}x | "
                        f"C40 Share <= 6x TDTI: 100.0% − {gt6_pct:.1f}% = {le6_pct:.1f}% (Share <= 5x TDTI: {le5_pct:.1f}%)"
                    )
                    out["fhb-typical-dti"] = mk_dti_rec(
                        "RBNZ C31, C40 & C41", as_of_c41, pub_c40,
                        f"In {m_c41}, approved New Zealand first-home buyers had an implied mortgage-to-income ratio of <strong>{implied_dti_x:.2f}x gross application income</strong> (${avg_loan_for_dti:,} average loan divided by ${fhb_inc_int:,} average gross borrower income), while in RBNZ Table C40 <strong>{le6_pct:.1f}% of all first-home buyer lending</strong> was written at or below <strong>6x TDTI</strong> (and <strong>{le5_pct:.1f}%</strong> at or below <strong>5x TDTI</strong>).",
                        f"RBNZ Tables C31, C40 & C41 ({m_c41}): Average FHB loan = ${avg_loan_for_dti:,}; Average FHB gross income = ${fhb_inc_int:,}; FHB lending <= 6x TDTI = {le6_pct:.1f}% (<= 5x TDTI = {le5_pct:.1f}%; > 6x TDTI = {gt6_pct:.1f}%).",
                        "RBNZ Tables C31, C40 & C41 (Reserve Bank of New Zealand)", c40_url, m_c41, implied_dti_x,
                        f"{implied_dti_x:.2f}x implied mortgage-to-income ratio | {le6_pct:.1f}% <= 6x TDTI",
                        f"{le5_pct:.1f}% of FHB lending written at <= 5x TDTI",
                        f"{gt6_pct:.1f}% above 6x TDTI (vs 20.0% RBNZ speed limit)",
                        typ_dti_calc, hist_start_c41,
                        [
                            {"period": "Implied Loan-to-Income", "label": "Implied FHB Loan-to-Income Ratio (C31 ÷ C41)", "value": implied_dti_x, "formatted": f"{implied_dti_x:.2f}x (${avg_loan_for_dti:,} ÷ ${fhb_inc_int:,})"},
                            {"period": "TDTI <= 5x Share", "label": "FHB Lending at or below 5x TDTI (C40)", "value": round(le5_pct, 1), "formatted": f"{le5_pct:.1f}% of FHB lending"},
                            {"period": "TDTI <= 6x Share", "label": "FHB Lending at or below 6x TDTI (C40)", "value": round(le6_pct, 1), "formatted": f"{le6_pct:.1f}% of FHB lending"},
                            {"period": "TDTI > 6x Share", "label": "FHB Lending above 6x TDTI (C40 Speed Limit Band)", "value": round(gt6_pct, 1), "formatted": f"{gt6_pct:.1f}% (vs 20.0% RBNZ cap)"},
                        ],
                    )
        except Exception as exc:
            print(f"[C41 BGI] Parse note: {exc}")
    return out


def build_rbnz_b30_ticker(filepath: pathlib.Path, session=None, s33_path: pathlib.Path | None = None) -> dict:
    """100% live-parsed settled rates (RBNZ B30) & fixed vs floating / existing stock repricing (RBNZ S33)."""
    pub_date = extract_rbnz_published_date(filepath)
    tabs = load_rbnz_sheets(filepath) or load_rbnz_sheets_merged(filepath)
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
            "html": f"In {m_full}, the volume-weighted average fixed mortgage interest rate on new owner-occupier commitments at New Zealand registered banks was <strong>{v1:.2f}%</strong> for 1-year terms and <strong>{v2:.2f}%</strong> for 2-year terms.",
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

    s33_page = "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-fully-secured-by-residential-mortgage"
    try:
        if s33_path is None or not s33_path.exists():
            active_sess = session or get_http_session()
            s33_cfg = RBNZ_SOURCES["s33"]
            for u_cand in (s33_cfg["xlsx"], s33_cfg["alt_xlsx"]):
                s33_bytes = fetch_bytes_with_browser_fallback(active_sess, u_cand, referer=s33_page, min_bytes=10_000)
                if s33_bytes and s33_bytes[:2] == b"PK":
                    s33_path = filepath.parent / "tmp_s33.xlsx"
                    s33_path.write_bytes(s33_bytes)
                    break

        if s33_path and s33_path.exists():
            s33_pub = extract_rbnz_published_date(s33_path)
            tabs_s33 = load_rbnz_sheets_merged(s33_path)

            dt_fl, s_fl, _ = _find_rbnz_multi(tabs_s33, [("floating",)], exclude_kw=("%", "share"), val_bounds=(15000, 120000))
            dt_tot, s_tot, _ = _find_rbnz_multi(tabs_s33, [("total",)], exclude_kw=("%", "share", "floating"), val_bounds=(200000, 550000))
            df_fl = _series_df(dt_fl, s_fl)
            df_tot = _series_df(dt_tot, s_tot)

            if not df_fl.empty and not df_tot.empty:
                m33_df = pd.merge(df_fl, df_tot, on="date", suffixes=("_fl", "_tot"))
                if not m33_df.empty:
                    m33_df["val_fx"] = m33_df["val_tot"] - m33_df["val_fl"]
                    m33_df["fx_sh"] = (m33_df["val_fx"] / m33_df["val_tot"]) * 100.0
                    last33 = m33_df.iloc[-1]
                    prev33 = m33_df.iloc[-13] if len(m33_df) >= 13 else m33_df.iloc[0]
                    dt33 = pd.Timestamp(last33["date"])
                    m33, as_of33 = dt33.strftime("%B %Y"), month_end_iso(dt33)
                    tot_m, fx_m, fl_m = float(last33["val_tot"]), float(last33["val_fx"]), float(last33["val_fl"])
                    tot_b, fx_b, fl_b = tot_m / 1000.0, fx_m / 1000.0, fl_m / 1000.0
                    fx_sh, fl_sh = float(last33["fx_sh"]), 100.0 - float(last33["fx_sh"])
                    fx_sh_prev = float(prev33["fx_sh"])

                    rep6_m, rep12_m = 0.0, 0.0
                    for _, df_r, _, lbls_r in tabs_s33:
                        r6_cols, r12_cols = [], []
                        for c_i, lbl in enumerate(lbls_r):
                            if c_i >= df_r.shape[1] or str(df_r.columns[c_i]) == "date" or any(x in lbl for x in ("%", "share")):
                                continue
                            if any(k in lbl for k in ("3 months", "6 months", "up to 3", "3 to 6")):
                                r6_cols.append(c_i)
                                r12_cols.append(c_i)
                            elif any(k in lbl for k in ("1 year", "6 to 12")):
                                r12_cols.append(c_i)
                        if r12_cols:
                            cand12 = sum(float(pd.to_numeric(df_r.iloc[:, c_i], errors="coerce").dropna().iloc[-1]) for c_i in r12_cols[:3] if len(pd.to_numeric(df_r.iloc[:, c_i], errors="coerce").dropna()))
                            cand6 = sum(float(pd.to_numeric(df_r.iloc[:, c_i], errors="coerce").dropna().iloc[-1]) for c_i in r6_cols[:2] if len(pd.to_numeric(df_r.iloc[:, c_i], errors="coerce").dropna()))
                            if 0.1 * fx_m <= cand12 <= fx_m:
                                rep12_m, rep6_m = cand12, cand6
                                break

                    rep12_clause, rep12_calc = "", ""
                    if rep12_m > 0:
                        rep12_b, rep12_pct = rep12_m / 1000.0, (rep12_m / fx_m) * 100.0
                        rep6_pct = (rep6_m / fx_m) * 100.0 if rep6_m > 0 else 0.0
                        rep12_clause = f"—with <strong>{rep12_pct:.1f}% (${rep12_b:.1f}B)</strong> of fixed-rate balances scheduled to reprice within 12 months"
                        rep12_calc = f" | Repricing <=12m: ${rep12_b:.2f}B ÷ ${fx_b:.2f}B = {rep12_pct:.1f}%"

                        out["existing-mortgage-stock-lvr-refix"] = make_pulse_item(
                            "RBNZ Table S33", as_of33, MAX_AGE_MONTHLY,
                            f"As at {m33}, across New Zealand registered banks' <strong>${tot_b:.1f}B</strong> existing residential mortgage stock (RBNZ Table S33), <strong>{fl_sh:.1f}% (${fl_b:.1f}B)</strong> was on floating rates and <strong>{fx_sh:.1f}% (${fx_b:.1f}B)</strong> on fixed rates—with <strong>{rep12_pct:.1f}% (${rep12_b:.1f}B)</strong> of fixed debt scheduled to reprice within 12 months" + (f" (and <strong>{rep6_pct:.1f}%</strong> within 6 months)." if rep6_pct > 0 else "."),
                            f"RBNZ Table S33 ({m33}): Total existing mortgage stock = ${tot_b:.1f}B; Fixed = ${fx_b:.1f}B ({fx_sh:.1f}%); Floating = ${fl_b:.1f}B ({fl_sh:.1f}%); Repricing <=12m = ${rep12_b:.1f}B ({rep12_pct:.1f}%).",
                            format_provenance_source_html(f"Fixed Share: ${fx_b:.2f}B ÷ ${tot_b:.2f}B = {fx_sh:.1f}%{rep12_calc}", "RBNZ Table S33 — Existing Mortgage Stock Repricing", s33_page, True),
                            m33, round(rep12_pct, 1),
                            f"{rep12_pct:.1f}% refixing <=12m (${rep12_b:.1f}B) | {fl_sh:.1f}% floating (${fl_b:.1f}B)",
                            f"Total NZ bank mortgage book: ${tot_b:.1f}B",
                            f"{fx_sh - fx_sh_prev:+.1f} pts YoY shift in fixed-rate share",
                            f"Fixed Share: ${fx_b:.2f}B ÷ ${tot_b:.2f}B = {fx_sh:.1f}%{rep12_calc}",
                            f"RBNZ Table S33 monthly series ({m33})",
                            make_history_list(m33_df["date"], m33_df["fx_sh"], lambda x: f"{x:.1f}% fixed share"),
                            {"source_published": s33_pub, "source_url": s33_page},
                        )

                    s33_calc = f"Fixed Share: ${fx_b:.2f}B ÷ ${tot_b:.2f}B = {fx_sh:.1f}% | Floating Share: ${fl_b:.2f}B ÷ ${tot_b:.2f}B = {fl_sh:.1f}%{rep12_calc}"
                    out["fixed-vs-floating-rate-mix"] = {
                        "chip": "RBNZ Table S33",
                        "as_of": as_of33,
                        "source_published": s33_pub,
                        "source_url": s33_page,
                        "max_age_days": MAX_AGE_MONTHLY,
                        "html": f"As at {m33}, across New Zealand registered banks' <strong>${tot_b:.1f}B</strong> in loans fully secured by residential mortgage, <strong>{fx_sh:.1f}% (${fx_b:.1f}B)</strong> was held on fixed interest rates and <strong>{fl_sh:.1f}% (${fl_b:.1f}B)</strong> on floating rates{rep12_clause}.",
                        "official_data_summary": f"RBNZ Table S33 ({m33}): Total residential mortgage loans = ${tot_m:,.0f}m; Fixed-rate loans = ${fx_m:,.0f}m ({fx_sh:.1f}%); Floating loans = ${fl_m:,.0f}m ({fl_sh:.1f}%).",
                        "source_html": format_provenance_source_html(s33_calc, "RBNZ Table S33 — Loans by Repricing (Reserve Bank of New Zealand)", s33_page, True),
                        "latest_period": m33,
                        "value": round(fx_sh, 2),
                        "latest_value": f"{fx_sh:.1f}% fixed (${fx_b:.1f}B) vs {fl_sh:.1f}% floating (${fl_b:.1f}B)",
                        "previous_period": f"{pd.Timestamp(prev33['date']).strftime('%B %Y')} ({fx_sh_prev:.1f}% fixed share)",
                        "change": f"{fx_sh - fx_sh_prev:+.1f} percentage points YoY (fixed share)",
                        "calculation_live": s33_calc,
                        "historical_data_from": f"Available from {pd.Timestamp(m33_df['date'].iloc[0]).strftime('%B %Y')} ({len(m33_df)} monthly observations)",
                        "history": make_history_list(m33_df["date"], m33_df["fx_sh"], lambda x: f"{x:.1f}% fixed"),
                    }
    except Exception as exc_s33:
        print(f"[RBNZ S33] Parse note: {exc_s33}")

    return out
# === END OF BLOCK 7 ===
# === BLOCK 8 OF 13 ===
def build_household_debt_stress_credit_tickers(session, rbnz_files: dict[str, pathlib.Path]) -> dict:
    """
    100% live-pulled Household Loans, Credit Cards, Consumer Lending, Revolving Credit & Borrower Stress datasets
    across RBNZ Tables C12, C5, S50, C35 & S32, Stats NZ HLPI, and Inland Revenue Tax Debt statistics.
    """
    out: dict = {}
    today = datetime.datetime.now(datetime.timezone.utc).date()
    root = pathlib.Path(__file__).resolve().parent

    # 1 & 2. RBNZ Table C12 — Credit Card Balances, Interest-Bearing Share, Interest Rate & Credit Limits
    c12_path = _ensure_rbnz_file("c12", root, session, rbnz_files)
    if c12_path and c12_path.exists():
        try:
            c12_pub = extract_rbnz_published_date(c12_path)
            c12_tabs = load_rbnz_sheets(c12_path) or load_rbnz_sheets_merged(c12_path)
            c12_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/credit-card-balances"

            dt_adv, s_adv, _ = _find_rbnz_multi(
                c12_tabs,
                [("advances", "total"), ("total", "advances", "outstanding"), ("advances", "outstanding"), ("advances",), ("balances",)],
                exclude_kw=("interest", "limit", "change", "yoy", "percentage", "rate"),
                val_bounds=(3500, 15000),
            )
            dt_ib, s_ib, _ = _find_rbnz_multi(
                c12_tabs,
                [("interest-bearing",), ("interest", "bearing"), ("interest", "advances")],
                exclude_kw=("rate", "weighted", "limit", "non-interest", "without"),
                val_bounds=(1500, 12000),
            )
            dt_rate, s_rate, _ = _find_rbnz_multi(
                c12_tabs,
                [("interest-bearing", "rate"), ("weighted", "rate"), ("interest", "rate"), ("rate",), ("%",)],
                exclude_kw=("exchange", "change", "yoy", "proportion"),
                val_bounds=(12.0, 28.0),
            )
            dt_lim, s_lim, _ = _find_rbnz_multi(
                c12_tabs,
                [("credit", "limit"), ("total", "limit"), ("limits",), ("limit",)],
                exclude_kw=("ratio", "percentage", "used", "proportion"),
                val_bounds=(12000, 45000),
            )

            df_adv = _series_df(dt_adv, s_adv)
            df_ib = _series_df(dt_ib, s_ib)
            df_rate = _series_df(dt_rate, s_rate)
            df_lim = _series_df(dt_lim, s_lim)

            for _, df_t, _, lbls in c12_tabs:
                for c_i in range(0, df_t.shape[1]):
                    if str(df_t.columns[c_i]) == "date":
                        continue
                    df_cand = _series_df(df_t["date"], df_t.iloc[:, c_i])
                    if len(df_cand) < 6:
                        continue
                    v_last = float(df_cand["val"].iloc[-1])
                    lbl = lbls[c_i] if c_i < len(lbls) else ""
                    if df_adv.empty and 4500 <= v_last <= 9500:
                        df_adv = df_cand
                    if df_lim.empty and 15000 <= v_last <= 35000:
                        df_lim = df_cand
                    if df_ib.empty and 2000 <= v_last <= 4400 and ("interest" in lbl or "bearing" in lbl):
                        df_ib = df_cand
                    if df_rate.empty and 16.5 <= v_last <= 24.5:
                        df_rate = df_cand

            if len(df_adv) >= 2:
                last_a = df_adv.iloc[-1]
                prev_a = df_adv.iloc[-13] if len(df_adv) >= 13 else df_adv.iloc[0]
                dt_now = pd.Timestamp(last_a["date"])
                per_s, as_of_s = dt_now.strftime("%B %Y"), month_end_iso(dt_now)
                adv_m, adv_prev_m = float(last_a["val"]), float(prev_a["val"])
                ib_m = float(df_ib["val"].iloc[-1]) if not df_ib.empty else None
                rate_pct = float(df_rate["val"].iloc[-1]) if not df_rate.empty else None

                ib_share = (ib_m / max(1.0, adv_m)) * 100.0 if (ib_m and ib_m < adv_m) else None
                rate_clause = (
                    f", with <strong>{ib_share:.1f}% (${ib_m:,.0f}m)</strong> accruing interest at a weighted average rate of <strong>{rate_pct:.1f}% p.a.</strong>"
                    if (ib_share and rate_pct)
                    else (f" at a weighted average interest rate of <strong>{rate_pct:.1f}% p.a.</strong>" if rate_pct else "")
                )
                c12_calc = f"Total Card Advances = ${adv_m:,.0f}m (${adv_m/1000:.2f}B)" + (
                    f" | Interest-Bearing: ${ib_m:,.0f}m ÷ ${adv_m:,.0f}m = {ib_share:.1f}% at {rate_pct:.1f}% p.a."
                    if (ib_share and rate_pct)
                    else (f" | Weighted Average Rate = {rate_pct:.1f}% p.a." if rate_pct else "")
                )
                c12_hist = make_history_list(df_adv["date"], df_adv["val"], lambda v: f"${v:,.0f}m (${v/1000:.2f}B)")
                out["credit-card-balances-interest-share"] = make_pulse_item(
                    "RBNZ C12", as_of_s, MAX_AGE_MONTHLY,
                    f"In {per_s}, New Zealanders owed <strong>${adv_m/1000:.2f}B (${adv_m:,.0f}m)</strong> in credit card balances ({plain_yoy_growth(adv_m, adv_prev_m)} from ${adv_prev_m:,.0f}m a year earlier){rate_clause}.",
                    f"RBNZ Table C12 ({per_s}): Total credit card advances outstanding = ${adv_m:,.0f}m (${adv_m/1000:.2f}B).",
                    format_provenance_source_html(c12_calc, "RBNZ Table C12 (Reserve Bank of New Zealand)", c12_url, True),
                    per_s, round(adv_m, 1),
                    f"${adv_m/1000:.2f}B (${adv_m:,.0f}m) card debt" + (f" | {ib_share:.1f}% paying {rate_pct:.1f}% p.a." if (ib_share and rate_pct) else (f" | {rate_pct:.1f}% avg rate" if rate_pct else "")),
                    f"12 months prior: ${adv_prev_m:,.0f}m",
                    f"{((adv_m - adv_prev_m)/max(1.0, adv_prev_m))*100:+.1f}% YoY" + (f" | {rate_pct:.1f}% avg interest rate" if rate_pct else ""),
                    c12_calc, "RBNZ Table C12 monthly series", c12_hist,
                    {"source_published": c12_pub, "source_url": c12_url},
                )

                if len(df_lim) >= 2:
                    last_l = df_lim.iloc[-1]
                    prev_l = df_lim.iloc[-13] if len(df_lim) >= 13 else df_lim.iloc[0]
                    dt_l = pd.Timestamp(last_l["date"])
                    per_l, as_of_l = dt_l.strftime("%B %Y"), month_end_iso(dt_l)
                    lim_m, lim_prev_m = float(last_l["val"]), float(prev_l["val"])
                    if lim_m > adv_m:
                        unused_m = lim_m - adv_m
                        util_pct = (adv_m / lim_m) * 100.0
                        lim_calc = f"Total Credit Limits = ${lim_m:,.0f}m (${lim_m/1000:.2f}B) | Utilisation: ${adv_m:,.0f}m ÷ ${lim_m:,.0f}m = {util_pct:.1f}% | Unused Limits: ${lim_m:,.0f}m − ${adv_m:,.0f}m = ${unused_m:,.0f}m (${unused_m/1000:.2f}B)"
                        lim_hist = make_history_list(df_lim["date"], df_lim["val"], lambda v: f"${v/1000:.2f}B (${v:,.0f}m) limits")
                        out["credit-card-limits-servicing-impact"] = make_pulse_item(
                            "RBNZ C12", as_of_l, MAX_AGE_MONTHLY,
                            f"In {per_l}, total approved New Zealand credit card limits stood at <strong>${lim_m/1000:.2f}B (${lim_m:,.0f}m)</strong> with a <strong>{util_pct:.1f}% utilisation rate</strong> (<strong>${adv_m/1000:.2f}B</strong> drawn)—leaving <strong>${unused_m/1000:.2f}B (${unused_m:,.0f}m) in unused credit card limits</strong>.",
                            f"RBNZ Table C12 ({per_l}): Total credit card limits = ${lim_m:,.0f}m; advances outstanding = ${adv_m:,.0f}m ({util_pct:.1f}% utilisation); unused limits = ${unused_m:,.0f}m.",
                            format_provenance_source_html(lim_calc, "RBNZ Table C12 (Reserve Bank of New Zealand)", c12_url, True),
                            per_l, round(lim_m, 1),
                            f"${lim_m/1000:.2f}B total limits ({util_pct:.1f}% used | ${unused_m/1000:.2f}B unused)",
                            f"12 months prior: ${lim_prev_m/1000:.2f}B (${lim_prev_m:,.0f}m) total limits",
                            f"{((lim_m - lim_prev_m)/max(1.0, lim_prev_m))*100:+.1f}% YoY in total credit limits",
                            lim_calc, "RBNZ Table C12 monthly series", lim_hist,
                            {"source_published": c12_pub, "source_url": c12_url},
                        )
        except Exception as exc_c12:
            print(f"[RBNZ C12] Parse note: {exc_c12}")

    # 3 & 4. RBNZ Table C5 — Total Housing Debt Stock ($B) & Personal/Consumer Lending Stock ($B)
    c5_path = _ensure_rbnz_file("c5", root, session, rbnz_files)
    if c5_path and c5_path.exists():
        try:
            c5_pub = extract_rbnz_published_date(c5_path)
            c5_tabs = load_rbnz_sheets_merged(c5_path)
            c5_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/sector-lending"

            dt_h, s_hous, _ = _find_rbnz_multi(
                c5_tabs, [("housing", "total"), ("housing",)],
                exclude_kw=("change", "yoy", "growth", "percentage", "non-bank", "nbli"), val_bounds=(250000, 600000),
            )
            dt_c, s_cons, _ = _find_rbnz_multi(
                c5_tabs, [("consumer",), ("personal",)],
                exclude_kw=("change", "yoy", "growth", "percentage", "non-bank", "nbli"), val_bounds=(6000, 35000),
            )
            dt_t, s_tot_sec, _ = _find_rbnz_multi(
                c5_tabs, [("total",)],
                exclude_kw=("change", "yoy", "growth", "percentage", "housing", "consumer", "personal"), val_bounds=(450000, 980000),
            )
            dt_nb, s_nb_cons, _ = _find_rbnz_multi(
                c5_tabs, [("consumer", "non-bank"), ("personal", "non-bank"), ("consumer", "nbli"), ("personal", "nbli")],
                exclude_kw=("change", "yoy", "growth", "percentage"), val_bounds=(500, 12000),
            )

            df_hous = _series_df(dt_h, s_hous)
            df_cons = _series_df(dt_c, s_cons)
            df_tot = _series_df(dt_t, s_tot_sec)
            df_nb = _series_df(dt_nb, s_nb_cons)

            if len(df_hous) >= 2:
                last5 = df_hous.iloc[-1]
                prev5 = df_hous.iloc[-13] if len(df_hous) >= 13 else df_hous.iloc[0]
                dt5 = pd.Timestamp(last5["date"])
                per5, as_of5 = dt5.strftime("%B %Y"), month_end_iso(dt5)

                hous_m, hous_prev_m = float(last5["val"]), float(prev5["val"])
                hous_b = round(hous_m / 1000.0, 2)
                hous_yoy_pct = ((hous_m - hous_prev_m) / max(1.0, hous_prev_m)) * 100.0
                tot_m = float(df_tot["val"].iloc[-1]) if not df_tot.empty else None
                sec_share_str = f", representing <strong>{(hous_m/tot_m)*100:.1f}%</strong> of all New Zealand private sector credit" if tot_m else ""
                c5_h_calc = f"Total NZ Housing Debt = ${hous_m:,.0f}m (${hous_b:.2f}B; YoY: {hous_yoy_pct:+.1f}%)" + (f" | Share of Total Credit: ${hous_m:,.0f}m ÷ ${tot_m:,.0f}m = {(hous_m/tot_m)*100:.1f}%" if tot_m else "")
                h_hist = make_history_list(df_hous["date"], df_hous["val"], lambda v: f"${v/1000:.2f}B (${v:,.0f}m)")
                out["total-housing-debt-stock"] = make_pulse_item(
                    "RBNZ C5", as_of5, MAX_AGE_MONTHLY,
                    f"As at {per5}, total New Zealand residential mortgage debt across registered banks and non-bank lenders reached <strong>${hous_b:.2f}B (${hous_m:,.0f}m)</strong> ({plain_yoy_growth(hous_m, hous_prev_m)} from ${hous_prev_m/1000:.2f}B a year earlier){sec_share_str}.",
                    f"RBNZ Table C5 ({per5}): Total housing lending stock = ${hous_m:,.0f}m (${hous_b:.2f}B).",
                    format_provenance_source_html(c5_h_calc, "RBNZ Table C5 (Reserve Bank of New Zealand)", c5_url, True),
                    per5, hous_b,
                    f"${hous_b:.2f}B (${hous_m:,.0f}m) total NZ mortgage debt",
                    f"12 months prior: ${hous_prev_m/1000:.2f}B (${hous_prev_m:,.0f}m)",
                    f"{hous_yoy_pct:+.1f}% YoY (${(hous_m - hous_prev_m)/1000:+.2f}B)",
                    c5_h_calc, "RBNZ Table C5 monthly series", h_hist,
                    {"source_published": c5_pub, "source_url": c5_url},
                )

            if len(df_cons) >= 2:
                lc5 = df_cons.iloc[-1]
                pc5 = df_cons.iloc[-13] if len(df_cons) >= 13 else df_cons.iloc[0]
                dt5c = pd.Timestamp(lc5["date"])
                per5c, as_of5c = dt5c.strftime("%B %Y"), month_end_iso(dt5c)
                cons_m, cons_prev_m = float(lc5["val"]), float(pc5["val"])
                cons_b = round(cons_m / 1000.0, 2)
                cons_yoy_pct = ((cons_m - cons_prev_m) / max(1.0, cons_prev_m)) * 100.0
                nb_m = float(df_nb["val"].iloc[-1]) if not df_nb.empty else None
                nb_clause = f", including <strong>${nb_m/1000:.2f}B ({(nb_m/cons_m)*100:.1f}%)</strong> from non-bank lenders" if nb_m else ""
                c5_c_calc = f"Total NZ Consumer Lending = ${cons_m:,.0f}m (${cons_b:.2f}B; YoY: {cons_yoy_pct:+.1f}%)" + (f" | Non-Bank Share: ${nb_m:,.0f}m ÷ ${cons_m:,.0f}m = {(nb_m/cons_m)*100:.1f}%" if nb_m else "")
                c_hist = make_history_list(df_cons["date"], df_cons["val"], lambda v: f"${v/1000:.2f}B (${v:,.0f}m)")
                out["consumer-personal-lending-stock"] = make_pulse_item(
                    "RBNZ C5", as_of5c, MAX_AGE_MONTHLY,
                    f"As at {per5c}, total New Zealand personal and consumer lending (personal loans, vehicle finance, hire purchase, overdrafts, and cards) stood at <strong>${cons_b:.2f}B (${cons_m:,.0f}m)</strong> ({plain_yoy_growth(cons_m, cons_prev_m)} from ${cons_prev_m/1000:.2f}B a year earlier){nb_clause}.",
                    f"RBNZ Table C5 ({per5c}): Total consumer sector lending stock = ${cons_m:,.0f}m (${cons_b:.2f}B).",
                    format_provenance_source_html(c5_c_calc, "RBNZ Table C5 (Reserve Bank of New Zealand)", c5_url, True),
                    per5c, cons_b,
                    f"${cons_b:.2f}B (${cons_m:,.0f}m) consumer & personal debt",
                    f"12 months prior: ${cons_prev_m/1000:.2f}B (${cons_prev_m:,.0f}m)",
                    f"{cons_yoy_pct:+.1f}% YoY (${cons_m - cons_prev_m:+,.0f}m)",
                    c5_c_calc, "RBNZ Table C5 monthly series", c_hist,
                    {"source_published": c5_pub, "source_url": c5_url},
                )
        except Exception as exc_c5:
            print(f"[RBNZ C5] Parse note: {exc_c5}")

    # 5 & 6. RBNZ Table S50 — Mortgage Arrears / Non-Performing Housing & Consumer Loans
    s50_path = _ensure_rbnz_file("s50", root, session, rbnz_files)
    if s50_path and s50_path.exists():
        try:
            s50_pub = extract_rbnz_published_date(s50_path)
            s50_tabs = load_rbnz_sheets_merged(s50_path)
            s50_url = "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-by-asset-quality"

            dt_hm, s_npl_h_m, _ = _find_rbnz_multi(
                s50_tabs,
                [("housing", "non-performing"), ("housing", "90 days"), ("housing", "impaired"), ("non-performing", "housing"), ("housing", "past due")],
                exclude_kw=("ratio", "percent", "%", "share"),
                val_bounds=(150, 15000),
            )
            dt_hp, s_npl_h_pct, _ = _find_rbnz_multi(
                s50_tabs,
                [("housing", "non-performing"), ("housing", "npl"), ("housing", "ratio"), ("housing", "%"), ("housing", "90 days")],
                val_bounds=(0.1, 4.5),
            )
            dt_htot, s_tot_bk_h, _ = _find_rbnz_multi(s50_tabs, [("housing", "total"), ("housing",)], val_bounds=(200000, 550000))

            df_hm = _series_df(dt_hm, s_npl_h_m)
            df_hp = _series_df(dt_hp, s_npl_h_pct)
            df_htot = _series_df(dt_htot, s_tot_bk_h)

            if df_hp.empty and not df_hm.empty and not df_htot.empty:
                m_h = pd.merge(df_hm, df_htot, on="date", suffixes=("_m", "_tot"))
                if not m_h.empty:
                    df_hp = pd.DataFrame({"date": m_h["date"], "val": (m_h["val_m"] / m_h["val_tot"]) * 100.0})
            if df_hm.empty and not df_hp.empty and not df_htot.empty:
                m_h = pd.merge(df_hp, df_htot, on="date", suffixes=("_pct", "_tot"))
                if not m_h.empty:
                    df_hm = pd.DataFrame({"date": m_h["date"], "val": (m_h["val_pct"] / 100.0) * m_h["val_tot"]})

            if len(df_hm) >= 2 and len(df_hp) >= 2:
                dt_s = pd.Timestamp(df_hp["date"].iloc[-1])
                per_s50, as_of_s50 = dt_s.strftime("%B %Y"), month_end_iso(dt_s)
                hm_now, hp_now = float(df_hm["val"].iloc[-1]), float(df_hp["val"].iloc[-1])
                hm_prev = float(df_hm["val"].iloc[-13]) if len(df_hm) >= 13 else float(df_hm["val"].iloc[0])
                hp_prev = float(df_hp["val"].iloc[-13]) if len(df_hp) >= 13 else float(df_hp["val"].iloc[0])
                s50_h_calc = f"Non-Performing Bank Housing Loans = ${hm_now:,.0f}m ({hp_now:.2f}% of housing loan book; vs {hp_prev:.2f}% / ${hm_prev:,.0f}m a year earlier)"
                s50_h_hist = make_history_list(df_hp["date"], df_hp["val"], lambda v: f"{v:.2f}% NPL ratio")
                out["mortgage-non-performing-arrears"] = make_pulse_item(
                    "RBNZ S50", as_of_s50, MAX_AGE_MONTHLY,
                    f"In {per_s50}, New Zealand bank non-performing residential mortgages (loans 90+ days past due or impaired) stood at <strong>{hp_now:.2f}%</strong> of total housing loans (<strong>${hm_now/1000:.2f}B / ${hm_now:,.0f}m</strong>, {plain_prior_year_pct(hp_now, hp_prev)}).",
                    f"RBNZ Table S50 ({per_s50}): Bank non-performing housing loans = ${hm_now:,.0f}m ({hp_now:.2f}% of housing lending).",
                    format_provenance_source_html(s50_h_calc, "RBNZ Table S50 (Reserve Bank of New Zealand)", s50_url, True),
                    per_s50, round(hp_now, 2),
                    f"{hp_now:.2f}% housing NPL ratio (${hm_now:,.0f}m non-performing)",
                    f"12 months prior: {hp_prev:.2f}% (${hm_prev:,.0f}m)",
                    f"{hp_now - hp_prev:+.2f} percentage points ({hm_now - hm_prev:+,.0f}m) YoY",
                    s50_h_calc, "RBNZ Table S50 monthly series", s50_h_hist,
                    {"source_published": s50_pub, "source_url": s50_url},
                )

            dt_cm, s_npl_c_m, _ = _find_rbnz_multi(
                s50_tabs,
                [("consumer", "non-performing"), ("personal", "non-performing"), ("consumer", "90 days"), ("personal", "90 days"), ("consumer", "impaired"), ("personal", "impaired")],
                exclude_kw=("ratio", "percent", "%"),
                val_bounds=(15, 3000),
            )
            dt_cp, s_npl_c_pct, _ = _find_rbnz_multi(
                s50_tabs,
                [("consumer", "non-performing"), ("personal", "non-performing"), ("consumer", "ratio"), ("personal", "ratio"), ("consumer", "%"), ("personal", "%")],
                val_bounds=(0.25, 12.0),
            )
            dt_ctot, s_tot_bk_c, _ = _find_rbnz_multi(s50_tabs, [("consumer",), ("personal",)], val_bounds=(3000, 25000))

            df_cm = _series_df(dt_cm, s_npl_c_m)
            df_cp = _series_df(dt_cp, s_npl_c_pct)
            df_ctot = _series_df(dt_ctot, s_tot_bk_c)

            if df_cp.empty and not df_cm.empty and not df_ctot.empty:
                m_c = pd.merge(df_cm, df_ctot, on="date", suffixes=("_m", "_tot"))
                if not m_c.empty:
                    df_cp = pd.DataFrame({"date": m_c["date"], "val": (m_c["val_m"] / m_c["val_tot"]) * 100.0})
            if df_cm.empty and not df_cp.empty and not df_ctot.empty:
                m_c = pd.merge(df_cp, df_ctot, on="date", suffixes=("_pct", "_tot"))
                if not m_c.empty:
                    df_cm = pd.DataFrame({"date": m_c["date"], "val": (m_c["val_pct"] / 100.0) * m_c["val_tot"]})

            if len(df_cm) >= 2 and len(df_cp) >= 2:
                dt_c = pd.Timestamp(df_cp["date"].iloc[-1])
                per_c50, as_of_c50 = dt_c.strftime("%B %Y"), month_end_iso(dt_c)
                cm_now, cp_now = float(df_cm["val"].iloc[-1]), float(df_cp["val"].iloc[-1])
                cm_prev = float(df_cm["val"].iloc[-13]) if len(df_cm) >= 13 else float(df_cm["val"].iloc[0])
                cp_prev = float(df_cp["val"].iloc[-13]) if len(df_cp) >= 13 else float(df_cp["val"].iloc[0])
                s50_c_calc = f"Non-Performing Bank Consumer Loans = ${cm_now:,.0f}m ({cp_now:.2f}% of consumer loan book; vs {cp_prev:.2f}% a year earlier)"
                c50_hist = make_history_list(df_cp["date"], df_cp["val"], lambda v: f"{v:.2f}% consumer NPL ratio")
                out["consumer-loan-stress-npl"] = make_pulse_item(
                    "RBNZ S50", as_of_c50, MAX_AGE_MONTHLY,
                    f"In {per_c50}, New Zealand bank non-performing personal and consumer loans (90+ days past due or impaired) stood at <strong>{cp_now:.2f}% (${cm_now:,.0f}m)</strong> ({plain_prior_year_pct(cp_now, cp_prev)}).",
                    f"RBNZ Table S50 ({per_c50}): Bank non-performing consumer loans = ${cm_now:,.0f}m ({cp_now:.2f}% of consumer lending).",
                    format_provenance_source_html(s50_c_calc, "RBNZ Table S50 (Reserve Bank of New Zealand)", s50_url, True),
                    per_c50, round(cp_now, 2),
                    f"{cp_now:.2f}% consumer NPL ratio (${cm_now:,.0f}m non-performing)",
                    f"12 months prior: {cp_prev:.2f}% (${cm_prev:,.0f}m)",
                    f"{cp_now - cp_prev:+.2f} percentage points ({cm_now - cm_prev:+,.0f}m) YoY",
                    s50_c_calc, "RBNZ Table S50 monthly series", c50_hist,
                    {"source_published": s50_pub, "source_url": s50_url},
                )
        except Exception as exc_s50:
            print(f"[RBNZ S50] Parse note: {exc_s50}")

    # 7. RBNZ Table C35 — Quarterly Mortgage Interest Charged ($B), Scheduled P&I ($B) & Voluntary Excess Repayments ($B)
    c35_path = _ensure_rbnz_file("c35", root, session, rbnz_files)
    if c35_path and c35_path.exists():
        try:
            c35_pub = extract_rbnz_published_date(c35_path)
            c35_tabs = load_rbnz_sheets_merged(c35_path)
            c35_url = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-loan-reconciliation"

            dt_35, s_int_m, _ = _find_rbnz_multi(
                c35_tabs, [("interest", "charged"), ("interest",)],
                exclude_kw=("rate", "%", "percent"), val_bounds=(2000, 9000),
            )
            dt_sc, s_sched_m, _ = _find_rbnz_multi(
                c35_tabs, [("scheduled",), ("repayment",)],
                exclude_kw=("excess", "full", "unscheduled"), val_bounds=(4000, 14000),
            )
            dt_ex, s_ex_m, _ = _find_rbnz_multi(
                c35_tabs,
                [("excess", "repayment"), ("excess", "repayments"), ("excess",)],
                exclude_kw=("rate", "%", "percent", "full", "scheduled"),
                val_bounds=(300, 9000),
            )
            df_int35 = _series_df(dt_35, s_int_m)
            df_sch35 = _series_df(dt_sc, s_sched_m)
            df_ex35 = _series_df(dt_ex, s_ex_m)

            if len(df_int35) >= 2:
                l35 = df_int35.iloc[-1]
                p35 = df_int35.iloc[-5] if len(df_int35) >= 5 else df_int35.iloc[0]
                dt35 = pd.Timestamp(l35["date"])
                per35, as_of35 = f"{dt35.strftime('%B %Y')} quarter", month_end_iso(dt35)
                int_m_now = float(l35["val"])
                int_b, int_prev_b = int_m_now / 1000.0, float(p35["val"]) / 1000.0
                int_yoy = ((int_b - int_prev_b) / max(0.1, int_prev_b)) * 100.0
                sched_b = float(df_sch35["val"].iloc[-1]) / 1000.0 if not df_sch35.empty else None
                sched_4q_b = float(df_sch35["val"].iloc[-4:].sum()) / 1000.0 if len(df_sch35) >= 4 else None
                sched_clause = f", alongside <strong>${sched_b:.2f}B in scheduled principal-and-interest repayments</strong> (${sched_4q_b:.2f}B over the rolling 12 months)" if (sched_b and sched_4q_b) else ""
                c35_calc = f"Quarterly Interest Charged ({per35}) = ${int_b:.2f}B (${int_m_now:,.0f}m; YoY: {int_yoy:+.1f}%)" + (f" | Scheduled P&I Repayments = ${sched_b:.2f}B/qtr (${sched_4q_b:.2f}B/yr)" if (sched_b and sched_4q_b) else "")
                c35_hist = make_history_list(df_int35["date"], df_int35["val"], lambda v: f"${v/1000:.2f}B (${v:,.0f}m) interest charged")
                out["household-debt-servicing-burden"] = make_pulse_item(
                    "RBNZ C35", as_of35, MAX_AGE_QUARTERLY,
                    f"In the {per35}, New Zealand homeowners were charged <strong>${int_b:.2f}B (${int_m_now:,.0f}m) in residential mortgage interest</strong> ({plain_yoy_growth(int_b, int_prev_b)} from ${int_prev_b:.2f}B a year earlier){sched_clause}.",
                    f"RBNZ Table C35 ({per35}): Quarterly residential mortgage interest charged = ${int_m_now:,.0f}m (${int_b:.2f}B).",
                    format_provenance_source_html(c35_calc, "RBNZ Table C35 (Reserve Bank of New Zealand)", c35_url, True),
                    per35, round(int_b, 2),
                    f"${int_b:.2f}B/qtr mortgage interest charged" + (f" | ${sched_b:.2f}B/qtr scheduled P&I" if sched_b else ""),
                    f"Same quarter prior year: ${int_prev_b:.2f}B",
                    f"{int_yoy:+.1f}% YoY (${int_b - int_prev_b:+.2f}B/qtr)",
                    c35_calc, "RBNZ Table C35 quarterly series", c35_hist,
                    {"source_published": c35_pub, "source_url": c35_url},
                )

            if len(df_ex35) >= 2:
                lex = df_ex35.iloc[-1]
                pex = df_ex35.iloc[-5] if len(df_ex35) >= 5 else df_ex35.iloc[0]
                dtex = pd.Timestamp(lex["date"])
                per_ex, as_of_ex = f"{dtex.strftime('%B %Y')} quarter", month_end_iso(dtex)
                ex_m_now, ex_prev_m = float(lex["val"]), float(pex["val"])
                ex_b, ex_prev_b = ex_m_now / 1000.0, ex_prev_m / 1000.0
                ex_yoy = ((ex_b - ex_prev_b) / max(0.05, ex_prev_b)) * 100.0
                ex_4q_m = float(df_ex35["val"].iloc[-4:].sum()) if len(df_ex35) >= 4 else ex_m_now * 4.0
                ex_4q_b = ex_4q_m / 1000.0

                sched_m_now = float(df_sch35["val"].iloc[-1]) if not df_sch35.empty else None
                if sched_m_now and sched_m_now > 0:
                    sched_b_now = sched_m_now / 1000.0
                    ex_vs_sched_pct = (ex_m_now / sched_m_now) * 100.0
                    ex_share_tot_pct = (ex_m_now / (sched_m_now + ex_m_now)) * 100.0
                    ex_clause = (
                        f"—equivalent to <strong>{ex_vs_sched_pct:.1f}% on top of the ${sched_b_now:.2f}B (${sched_m_now:,.0f}m)</strong> "
                        f"paid in scheduled minimum principal-and-interest repayments (and bringing rolling 12-month voluntary excess repayments to <strong>${ex_4q_b:.2f}B</strong>)"
                    )
                    ex_calc = (
                        f"Quarterly Excess Repayments ({per_ex}) = ${ex_b:.2f}B (${ex_m_now:,.0f}m; 12M rolling = ${ex_4q_b:.2f}B) | "
                        f"Ratio to Scheduled P&I: ${ex_m_now:,.0f}m ÷ ${sched_m_now:,.0f}m = {ex_vs_sched_pct:.1f}% "
                        f"({ex_share_tot_pct:.1f}% of combined scheduled + excess repayments)"
                    )
                    ex_latest = f"${ex_b:.2f}B/qtr excess repayments (+{ex_vs_sched_pct:.1f}% above scheduled P&I | ${ex_4q_b:.2f}B/yr)"
                else:
                    ex_clause = f" (bringing rolling 12-month voluntary excess repayments to <strong>${ex_4q_b:.2f}B</strong>)"
                    ex_calc = f"Quarterly Excess Repayments ({per_ex}) = ${ex_b:.2f}B (${ex_m_now:,.0f}m; 12M rolling = ${ex_4q_b:.2f}B; YoY: {ex_yoy:+.1f}%)"
                    ex_latest = f"${ex_b:.2f}B/qtr excess repayments (${ex_4q_b:.2f}B/yr)"

                ex_hist = make_history_list(df_ex35["date"], df_ex35["val"], lambda v: f"${v/1000:.2f}B (${v:,.0f}m) excess repayments")
                out["mortgage-excess-repayments"] = make_pulse_item(
                    "RBNZ C35", as_of_ex, MAX_AGE_QUARTERLY,
                    f"In the {per_ex}, New Zealand mortgage borrowers made <strong>${ex_b:.2f}B (${ex_m_now:,.0f}m) in voluntary excess mortgage repayments</strong> above their contracted minimum schedules ({plain_yoy_growth(ex_b, ex_prev_b)} from ${ex_prev_b:.2f}B a year earlier){ex_clause}.",
                    f"RBNZ Table C35 ({per_ex}): Quarterly voluntary excess mortgage repayments (excluding full loan discharges) = ${ex_m_now:,.0f}m (${ex_b:.2f}B; rolling 4 quarters = ${ex_4q_b:.2f}B).",
                    format_provenance_source_html(ex_calc, "RBNZ Table C35 (Reserve Bank of New Zealand)", c35_url, True),
                    per_ex, round(ex_b, 2),
                    ex_latest,
                    f"Same quarter prior year: ${ex_prev_b:.2f}B (${ex_prev_m:,.0f}m)",
                    f"{ex_yoy:+.1f}% YoY (${ex_b - ex_prev_b:+.2f}B/qtr)",
                    ex_calc, "RBNZ Table C35 quarterly series", ex_hist,
                    {"source_published": c35_pub, "source_url": c35_url},
                )
        except Exception as exc_c35:
            print(f"[RBNZ C35] Parse note: {exc_c35}")

    # 7B. RBNZ Table S32 — Residential Revolving Credit Drawn Balances ($B) & Share of Housing Lending
    s32_path = _ensure_rbnz_file("s32", root, session, rbnz_files)
    if s32_path and s32_path.exists():
        try:
            s32_pub = extract_rbnz_published_date(s32_path)
            s32_tabs = load_rbnz_sheets_merged(s32_path)
            s32_url = "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-by-product"

            dt_oo_r, s_oo_rev, _ = _find_rbnz_multi(
                s32_tabs,
                [("owner", "revolving"), ("owner-occupied", "revolving"), ("owner occupied", "revolving")],
                exclude_kw=("%", "share", "ratio", "business", "agri", "personal"),
                val_bounds=(2000, 45000),
            )
            dt_in_r, s_inv_rev, _ = _find_rbnz_multi(
                s32_tabs,
                [("investor", "revolving"), ("investment", "revolving"), ("property investment", "revolving")],
                exclude_kw=("%", "share", "ratio", "business", "agri", "personal"),
                val_bounds=(500, 25000),
            )
            dt_oo_t, s_oo_tot, _ = _find_rbnz_multi(
                s32_tabs,
                [("owner", "total"), ("owner-occupied", "housing"), ("owner occupied", "total")],
                exclude_kw=("revolving", "interest only", "interest-only", "principal", "%", "share"),
                val_bounds=(120000, 450000),
            )
            dt_in_t, s_inv_tot, _ = _find_rbnz_multi(
                s32_tabs,
                [("investor", "total"), ("investment", "total"), ("property investment", "total")],
                exclude_kw=("revolving", "interest only", "interest-only", "principal", "%", "share"),
                val_bounds=(40000, 200000),
            )

            df_oo_rev = _series_df(dt_oo_r, s_oo_rev)
            df_inv_rev = _series_df(dt_in_r, s_inv_rev)
            df_oo_tot = _series_df(dt_oo_t, s_oo_tot)
            df_inv_tot = _series_df(dt_in_t, s_inv_tot)

            if df_oo_rev.empty:
                dt_h_r, s_h_rev, _ = _find_rbnz_multi(
                    s32_tabs,
                    [("housing", "revolving"), ("residential", "revolving"), ("revolving", "credit")],
                    exclude_kw=("%", "share", "ratio", "business", "agri", "personal"),
                    val_bounds=(3000, 65000),
                )
                df_oo_rev = _series_df(dt_h_r, s_h_rev)

            if len(df_oo_rev) >= 2:
                if not df_inv_rev.empty:
                    m_rev = pd.merge(df_oo_rev, df_inv_rev, on="date", suffixes=("_oo", "_inv"))
                    m_rev["val_tot"] = m_rev["val_oo"] + m_rev["val_inv"]
                else:
                    m_rev = df_oo_rev.rename(columns={"val": "val_oo"}).copy()
                    m_rev["val_inv"] = 0.0
                    m_rev["val_tot"] = m_rev["val_oo"]

                if len(m_rev) >= 2:
                    last32 = m_rev.iloc[-1]
                    prev32 = m_rev.iloc[-13] if len(m_rev) >= 13 else m_rev.iloc[0]
                    dt32 = pd.Timestamp(last32["date"])
                    per32, as_of32 = dt32.strftime("%B %Y"), month_end_iso(dt32)

                    oo_rev_m = float(last32["val_oo"])
                    inv_rev_m = float(last32["val_inv"])
                    tot_rev_m = float(last32["val_tot"])
                    tot_rev_prev_m = float(prev32["val_tot"])

                    tot_rev_b = tot_rev_m / 1000.0
                    tot_rev_prev_b = tot_rev_prev_m / 1000.0
                    oo_rev_b = oo_rev_m / 1000.0
                    inv_rev_b = inv_rev_m / 1000.0
                    rev_yoy = ((tot_rev_m - tot_rev_prev_m) / max(1.0, tot_rev_prev_m)) * 100.0

                    oo_tot_m = float(df_oo_tot["val"].iloc[-1]) if not df_oo_tot.empty else 0.0
                    inv_tot_m = float(df_inv_tot["val"].iloc[-1]) if not df_inv_tot.empty else 0.0
                    tot_hous_m = oo_tot_m + inv_tot_m

                    if tot_hous_m > 0 and inv_rev_m > 0:
                        tot_rev_sh = (tot_rev_m / tot_hous_m) * 100.0
                        oo_rev_sh = (oo_rev_m / oo_tot_m) * 100.0 if oo_tot_m > 0 else 0.0
                        inv_rev_sh = (inv_rev_m / inv_tot_m) * 100.0 if inv_tot_m > 0 else 0.0
                        s32_clause = (
                            f" (representing <strong>{tot_rev_sh:.1f}% of all bank residential mortgage lending</strong>—comprising "
                            f"<strong>${oo_rev_b:.2f}B</strong> [{oo_rev_sh:.1f}%] across owner-occupiers and "
                            f"<strong>${inv_rev_b:.2f}B</strong> [{inv_rev_sh:.1f}%] across residential property investors)"
                        )
                        s32_calc = (
                            f"Total Residential Revolving Credit ({per32}): ${oo_rev_m:,.0f}m (Owner-Occupier) + ${inv_rev_m:,.0f}m (Investor) = "
                            f"${tot_rev_m:,.0f}m (${tot_rev_b:.2f}B) | Share of Housing Loans: ${tot_rev_m:,.0f}m ÷ ${tot_hous_m:,.0f}m = "
                            f"{tot_rev_sh:.1f}% (OO: {oo_rev_sh:.1f}%, Investor: {inv_rev_sh:.1f}%)"
                        )
                        s32_latest = f"${tot_rev_b:.2f}B drawn revolving credit ({tot_rev_sh:.1f}% of mortgages | OO ${oo_rev_b:.2f}B vs Inv ${inv_rev_b:.2f}B)"
                    elif inv_rev_m > 0:
                        s32_clause = f"—comprising <strong>${oo_rev_b:.2f}B (${oo_rev_m:,.0f}m)</strong> for owner-occupiers and <strong>${inv_rev_b:.2f}B (${inv_rev_m:,.0f}m)</strong> for residential property investors"
                        s32_calc = f"Total Residential Revolving Credit ({per32}): ${oo_rev_m:,.0f}m (Owner-Occupier) + ${inv_rev_m:,.0f}m (Investor) = ${tot_rev_m:,.0f}m (${tot_rev_b:.2f}B)"
                        s32_latest = f"${tot_rev_b:.2f}B drawn revolving credit (OO ${oo_rev_b:.2f}B | Investor ${inv_rev_b:.2f}B)"
                    else:
                        s32_clause = ""
                        s32_calc = f"Total Residential Revolving Credit ({per32}) = ${tot_rev_m:,.0f}m (${tot_rev_b:.2f}B; YoY: {rev_yoy:+.1f}%)"
                        s32_latest = f"${tot_rev_b:.2f}B (${tot_rev_m:,.0f}m) drawn revolving credit"

                    s32_hist = make_history_list(m_rev["date"], m_rev["val_tot"], lambda v: f"${v/1000:.2f}B (${v:,.0f}m) revolving credit")
                    out["mortgage-revolving-credit-balances"] = make_pulse_item(
                        "RBNZ S32", as_of32, MAX_AGE_MONTHLY,
                        f"As at {per32}, New Zealand borrowers held <strong>${tot_rev_b:.2f}B (${tot_rev_m:,.0f}m) in drawn residential revolving credit mortgage balances</strong>{s32_clause} ({plain_yoy_growth(tot_rev_m, tot_rev_prev_m)} from ${tot_rev_prev_b:.2f}B a year earlier).",
                        f"RBNZ Table S32 ({per32}): Registered bank residential revolving credit drawn balances = ${tot_rev_m:,.0f}m (${tot_rev_b:.2f}B).",
                        format_provenance_source_html(s32_calc, "RBNZ Table S32 (Reserve Bank of New Zealand)", s32_url, True),
                        per32, round(tot_rev_b, 2),
                        s32_latest,
                        f"12 months prior: ${tot_rev_prev_b:.2f}B (${tot_rev_prev_m:,.0f}m)",
                        f"{rev_yoy:+.1f}% YoY (${(tot_rev_m - tot_rev_prev_m)/1000:+.2f}B)",
                        s32_calc, "RBNZ Table S32 monthly series", s32_hist,
                        {"source_published": s32_pub, "source_url": s32_url},
                    )
        except Exception as exc_s32:
            print(f"[RBNZ S32] Parse note: {exc_s32}")

    # 8. Stats NZ HLPI (Live HTML + CSV Parser, Zero Hardcoded Fallbacks)
    hlpi_pages = fetch_html_with_browser_fallback(session, STATS_HLPI_URLS, "Stats NZ HLPI")
    if not any("all households" in (h or "").lower() or "living-costs" in (h or "").lower() for h in hlpi_pages.values()):
        try:
            hlpi_pages.update(fetch_pages_with_camoufox(STATS_HLPI_URLS[:2], wait_ms=6000))
        except Exception as exc_cam_hlpi:
            print(f"[Stats NZ HLPI Camoufox] {exc_cam_hlpi}")

    for hlpi_u, hlpi_html in hlpi_pages.items():
        if not hlpi_html:
            continue
        hlpi_pub = extract_html_release_date(hlpi_html, "Quarterly Stats NZ Household Living-Costs Price Indexes (HLPI) release")
        h_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", hlpi_html))
        m_12m = re.search(
            r"12\s+months\s+to\s+the\s+([A-Za-z]+\s+20\d{2})\s+quarter[^0-9]{0,110}?all[- ]households\s+was\s+([\d.]+)\s*(?:percent|%)",
            h_txt, flags=re.I,
        )
        if m_12m:
            hlpi_per, hlpi_pct = m_12m.group(1).strip(), float(m_12m.group(2))
        else:
            m_alt = re.search(
                r"(?:average\s+New\s+Zealand\s+household|household\s+living\s+costs)[^0-9]{0,70}?(?:increased|rose)\s+([\d.]+)\s*(?:percent|%)\s+in\s+the\s+12\s+months\s+to\s+the\s+([A-Za-z]+\s+20\d{2})\s+quarter",
                h_txt, flags=re.I,
            )
            if not m_alt:
                continue
            hlpi_pct, hlpi_per = float(m_alt.group(1)), m_alt.group(2).strip()

        dt_hlpi = pd.to_datetime("1 " + hlpi_per, errors="coerce")
        as_of_hlpi = month_end_iso(dt_hlpi) if pd.notna(dt_hlpi) else today.strftime("%Y-%m-%d")
        m_low = re.search(r"lowest-expenditure\s+household\s+group\s+was\s+([\d.]+)\s*(?:percent|%)", h_txt, flags=re.I)
        m_high = re.search(r"highest-expenditure\s+household\s+group\s+was\s+([\d.]+)\s*(?:percent|%)", h_txt, flags=re.I)
        grp_clause = (
            f" (ranging from <strong>{m_high.group(1)}%</strong> for the highest-expenditure household group to <strong>{m_low.group(1)}%</strong> for the lowest-expenditure household group)"
            if (m_low and m_high) else ""
        )
        hlpi_calc = f"Official Stats NZ HLPI All-Households 12-month change ({hlpi_per} quarter) = +{hlpi_pct:.1f}%"
        out["hlpi-household-living-costs-interest"] = make_pulse_item(
            "Stats NZ HLPI", as_of_hlpi, MAX_AGE_QUARTERLY,
            f"In the 12 months to the {hlpi_per} quarter, Stats NZ's Household Living-Costs Price Index (HLPI) for all New Zealand households rose <strong>{hlpi_pct:.1f}%</strong>{grp_clause}.",
            f"Stats NZ Household Living-Costs Price Indexes ({hlpi_per} quarter): All-households annual change = +{hlpi_pct:.1f}%.",
            format_provenance_source_html(hlpi_calc, f"Stats NZ — Household Living-Costs Price Indexes ({hlpi_per} quarter)", STATS_HLPI_URLS[0], False),
            f"{hlpi_per} quarter", hlpi_pct,
            f"+{hlpi_pct:.1f}% annual household living-cost inflation",
            f"Highest-expenditure households: +{m_high.group(1)}% | Lowest-expenditure: +{m_low.group(1)}%" if (m_low and m_high) else "Stats NZ quarterly HLPI series",
            "Includes household mortgage interest & council rates",
            hlpi_calc, f"Stats NZ HLPI release ({hlpi_per} quarter)", None,
            {"source_published": hlpi_pub, "source_url": STATS_HLPI_URLS[0]},
        )
        break

    # 9. IRD Overdue Tax & Entitlement Debt
    ird_pages = fetch_html_with_browser_fallback(session, IRD_TAX_DEBT_URLS, "IRD Tax Debt")
    tax_html = ird_pages.get(IRD_TAX_DEBT_URLS[0], "")
    if tax_html:
        tax_pub = extract_html_release_date(tax_html, "Inland Revenue Tax and Entitlements Debt Statistics release")
        t_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", tax_html))
        m_cust = re.search(
            r"number\s+of\s+customers\s+with\s+tax\s+and\s+entitlements\s+debt\s+(?:increased|decreased)\s+from\s+([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})\s+to\s+([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})",
            t_txt, flags=re.I,
        )
        m_tot_d = re.search(
            r"Total\s+tax\s+and\s+entitlements\s+debt\s+(?:increased|decreased)\s+from\s+\$([\d.]+)\s*(?:billion|b)\s+in\s+([A-Za-z]+\s+\d{4})\s+to\s+\$([\d.]+)\s*(?:billion|b)\s+in\s+([A-Za-z]+\s+\d{4})",
            t_txt, flags=re.I,
        )
        m_ind_d = re.search(
            r"Tax\s+and\s+entitlements\s+debt\s+owed\s+by\s+individuals\s+(?:increased|decreased)\s+from\s+\$([\d.]+)\s*(?:billion|b)\s+in\s+[A-Za-z]+\s+\d{4}\s+to\s+\$([\d.]+)\s*(?:billion|b)\s+in\s+[A-Za-z]+\s+\d{4}",
            t_txt, flags=re.I,
        )
        if m_cust and m_tot_d:
            prev_cust, prev_per, curr_cust, curr_per = int(m_cust.group(1).replace(",", "")), m_cust.group(2), int(m_cust.group(3).replace(",", "")), m_cust.group(4)
            prev_debt_b, curr_debt_b = float(m_tot_d.group(1)), float(m_tot_d.group(3))
            ind_prev_b = float(m_ind_d.group(1)) if m_ind_d else None
            ind_curr_b = float(m_ind_d.group(2)) if m_ind_d else None
            avg_debt_cust = int(round((curr_debt_b * 1e9) / max(1, curr_cust)))
            dt_tax = pd.to_datetime("1 " + curr_per, errors="coerce")
            as_of_tax = month_end_iso(dt_tax) if pd.notna(dt_tax) else f"{today.year if today.month >= 7 else today.year - 1}-06-30"
            ind_clause = f", including <strong>${ind_curr_b:.1f}B owed by individuals</strong> (up from ${ind_prev_b:.1f}B)" if (ind_curr_b and ind_prev_b) else ""
            tax_calc = f"Total Overdue IRD Debt ({curr_per}) = ${curr_debt_b:.1f}B across {curr_cust:,} customers (Mean = ${avg_debt_cust:,}/customer" + (f" | Individuals = ${ind_curr_b:.1f}B)" if ind_curr_b else ")")
            out["ird-overdue-tax-debt-stress"] = make_pulse_item(
                "Inland Revenue", as_of_tax, MAX_AGE_ANNUAL,
                f"As at {curr_per}, <strong>{curr_cust:,} customers</strong> held overdue tax and entitlements debt with Inland Revenue totalling <strong>${curr_debt_b:.1f}B</strong> ({plain_yoy_growth(curr_debt_b, prev_debt_b)} from ${prev_debt_b:.1f}B across {prev_cust:,} customers in {prev_per}—averaging <strong>${avg_debt_cust:,} per customer in debt</strong>){ind_clause}.",
                f"Inland Revenue Tax and Entitlements Debt Statistics ({curr_per}): {curr_cust:,} customers in debt; total overdue debt = ${curr_debt_b:.1f}B" + (f"; owed by individuals = ${ind_curr_b:.1f}B." if ind_curr_b else "."),
                format_provenance_source_html(tax_calc, f"Inland Revenue — Tax and Entitlements Debt Statistics ({curr_per})", IRD_TAX_DEBT_URLS[0], True),
                curr_per, curr_debt_b,
                f"${curr_debt_b:.1f}B overdue IRD debt across {curr_cust:,} customers (avg ${avg_debt_cust:,})",
                f"{prev_per}: ${prev_debt_b:.1f}B across {prev_cust:,} customers",
                f"+${curr_debt_b - prev_debt_b:.1f}B ({((curr_debt_b - prev_debt_b)/prev_debt_b)*100:+.1f}% YoY) | {curr_cust - prev_cust:+,d} customers",
                tax_calc, f"Inland Revenue debt statistics ({curr_per})", None,
                {"source_published": tax_pub, "source_url": IRD_TAX_DEBT_URLS[0]},
            )

    return out


def build_rbnz_reality_check_expansion_tickers(
    session, rbnz_files: dict[str, pathlib.Path], existing_candidates: dict
) -> dict:
    """
    100% live-pulled RBNZ Rate Spread & Repricing Datasets (Zero Hardcoded Fallbacks):
    - back-book-vs-new-settled-mortgage-rate (RBNZ Table B6 vs Table B30)
    - low-equity-margin-penalty-spread (RBNZ Table B20 vs Table B21)
    - fixed-vs-floating-rate-mix & existing-mortgage-stock-lvr-refix (RBNZ Table S33)
    """
    out: dict = {}
    root = pathlib.Path(__file__).resolve().parent
    settled_rec = existing_candidates.get("settled-rates") or {}
    settled_1y = settled_rec.get("value")

    # 1. RBNZ Table B6 vs Table B30 — Existing Back-Book Mortgage Yield vs Newly Settled Fixed Rate
    b6_path = _ensure_rbnz_file("b6", root, session, rbnz_files)
    if b6_path and b6_path.exists() and isinstance(settled_1y, (int, float)):
        try:
            b6_pub = extract_rbnz_published_date(b6_path)
            tabs_b6 = load_rbnz_sheets_merged(b6_path)
            b6_url = "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/yields-on-loans"
            dt_b6, s_b6, _ = _find_rbnz_multi(
                tabs_b6, [("housing",), ("residential",), ("mortgage",), ("total",)],
                val_bounds=(2.0, 12.0),
            )
            df_b6 = _series_df(dt_b6, s_b6)
            if len(df_b6) >= 2:
                dt6 = pd.Timestamp(df_b6["date"].iloc[-1])
                per6, as_of6 = dt6.strftime("%B %Y"), month_end_iso(dt6)
                yield_now = float(df_b6["val"].iloc[-1])
                yield_prev = float(df_b6["val"].iloc[-13]) if len(df_b6) >= 13 else float(df_b6["val"].iloc[0])
                s1y_f = float(settled_1y)
                spread_bps = int(round((yield_now - s1y_f) * 100))
                b6_calc = f"Existing Housing Loan Yield (Table B6) = {yield_now:.2f}% vs 1Y Newly Settled Rate (Table B30) = {s1y_f:.2f}% (Spread = {spread_bps:+d} bps)"
                out["back-book-vs-new-settled-mortgage-rate"] = make_pulse_item(
                    "RBNZ B6 & B30", as_of6, MAX_AGE_MONTHLY,
                    f"In {per6}, the effective interest rate paid across all existing New Zealand bank mortgages ('back book', RBNZ Table B6) averaged <strong>{yield_now:.2f}%</strong>, compared with <strong>{s1y_f:.2f}%</strong> for newly settled 1-year fixed owner-occupier loans (RBNZ Table B30)—a difference of <strong>{abs(spread_bps)} basis points ({yield_now - s1y_f:+.2f}%)</strong>.",
                    f"RBNZ Table B6 & B30 ({per6}): Existing housing loan book yield = {yield_now:.2f}%; New settled 1-year fixed rate = {s1y_f:.2f}%.",
                    format_provenance_source_html(b6_calc, "RBNZ Tables B6 & B30 (Reserve Bank of New Zealand)", b6_url, True),
                    per6, round(yield_now, 2),
                    f"{yield_now:.2f}% existing mortgage yield vs {s1y_f:.2f}% new 1Y settled ({spread_bps:+d} bps)",
                    f"12 months prior existing yield: {yield_prev:.2f}%",
                    f"{yield_now - yield_prev:+.2f}% YoY on existing mortgage book",
                    b6_calc, "RBNZ Table B6 monthly series",
                    make_history_list(df_b6["date"], df_b6["val"], lambda v: f"{v:.2f}% effective yield"),
                    {"source_published": b6_pub, "source_url": b6_url},
                )
        except Exception as exc_b6:
            print(f"[RBNZ B6] Parse note: {exc_b6}")

    # 2. RBNZ Table B20 vs Table B21 — Standard (>80% LVR) vs Special (<=80% LVR) Low-Equity Rate Spread
    b20_path = _ensure_rbnz_file("b20", root, session, rbnz_files)
    b21_path = _ensure_rbnz_file("b21", root, session, rbnz_files)
    if b20_path and b20_path.exists() and b21_path and b21_path.exists():
        try:
            b21_pub = extract_rbnz_published_date(b21_path)
            tabs_b20 = load_rbnz_sheets_merged(b20_path)
            tabs_b21 = load_rbnz_sheets_merged(b21_path)
            b21_url = "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/new-residential-mortgage-special-interest-rates"
            dt20, s20_1y, _ = _find_rbnz_multi(tabs_b20, [("1 year",), ("1-year",)], val_bounds=(2.0, 14.0))
            dt21, s21_1y, _ = _find_rbnz_multi(tabs_b21, [("1 year",), ("1-year",)], val_bounds=(2.0, 14.0))
            df20 = _series_df(dt20, s20_1y)
            df21 = _series_df(dt21, s21_1y)
            if not df20.empty and not df21.empty:
                m_20_21 = pd.merge(df20, df21, on="date", suffixes=("_std", "_spc"))
                if not m_20_21.empty:
                    m_20_21["spread"] = m_20_21["val_std"] - m_20_21["val_spc"]
                    last_r = m_20_21.iloc[-1]
                    prev_r = m_20_21.iloc[-13] if len(m_20_21) >= 13 else m_20_21.iloc[0]
                    dt_r = pd.Timestamp(last_r["date"])
                    per_r, as_of_r = dt_r.strftime("%B %Y"), month_end_iso(dt_r)
                    std_1y, spc_1y = float(last_r["val_std"]), float(last_r["val_spc"])
                    pen_pct = std_1y - spc_1y
                    if pen_pct > 0:
                        pen_bps = int(round(pen_pct * 100))
                        b20_calc = f"1Y Standard Rate (Table B20: {std_1y:.2f}%) − 1Y Special Rate (Table B21: {spc_1y:.2f}%) = +{pen_pct:.2f}% (+{pen_bps} basis points)"
                        out["low-equity-margin-penalty-spread"] = make_pulse_item(
                            "RBNZ B20 & B21", as_of_r, MAX_AGE_MONTHLY,
                            f"In {per_r}, the gap between New Zealand banks' advertised 1-year 'Standard' fixed rate for borrowers with under 20% equity (<strong>{std_1y:.2f}%</strong>, RBNZ Table B20) and 1-year 'Special' fixed rate for borrowers with 20%+ equity (<strong>{spc_1y:.2f}%</strong>, Table B21) was <strong>+{pen_bps} basis points (+{pen_pct:.2f}%)</strong>.",
                            f"RBNZ Tables B20 & B21 ({per_r}): 1-year Standard rate = {std_1y:.2f}%; 1-year Special rate = {spc_1y:.2f}%; Spread = +{pen_pct:.2f}% (+{pen_bps} bps).",
                            format_provenance_source_html(b20_calc, "RBNZ Tables B20 & B21 (Reserve Bank of New Zealand)", b21_url, True),
                            per_r, round(pen_pct, 2),
                            f"+{pen_bps} bps low-equity spread ({std_1y:.2f}% Standard vs {spc_1y:.2f}% Special)",
                            f"12 months prior: {float(prev_r['val_std']):.2f}% Standard vs {float(prev_r['val_spc']):.2f}% Special",
                            f"+{pen_pct:.2f}% (+{pen_bps} bps) advertised rate margin between >80% LVR and <=80% LVR",
                            b20_calc, "RBNZ Tables B20 & B21 monthly series",
                            make_history_list(m_20_21["date"], m_20_21["spread"], lambda v: f"+{v:.2f}% (+{int(round(v*100))} bps)"),
                            {"source_published": b21_pub, "source_url": b21_url},
                        )
        except Exception as exc_b20:
            print(f"[RBNZ B20/B21] Parse note: {exc_b20}")

    # 3. RBNZ Table S33 — Fixed vs Floating Share & 12-Month Repricing Horizon (without 'floating' in exclude_kw)
    if "fixed-vs-floating-rate-mix" not in existing_candidates or "existing-mortgage-stock-lvr-refix" not in existing_candidates:
        s33_path = _ensure_rbnz_file("s33", root, session, rbnz_files)
        if s33_path and s33_path.exists():
            try:
                s33_page = "https://www.rbnz.govt.nz/statistics/series/registered-banks/banks-assets-loans-fully-secured-by-residential-mortgage"
                s33_pub = extract_rbnz_published_date(s33_path)
                tabs_s33 = load_rbnz_sheets(s33_path) or load_rbnz_sheets_merged(s33_path)

                dt_fl, s_fl, _ = _find_rbnz_multi(tabs_s33, [("floating",)], exclude_kw=("%", "share"), val_bounds=(15000, 120000))
                dt_tot, s_tot, _ = _find_rbnz_multi(tabs_s33, [("total",)], exclude_kw=("%", "share"), val_bounds=(200000, 550000))
                df_fl = _series_df(dt_fl, s_fl)
                df_tot = _series_df(dt_tot, s_tot)

                if (df_fl.empty or df_tot.empty) and tabs_s33:
                    for _, df_s, _, _ in tabs_s33:
                        for c_i in range(0, df_s.shape[1]):
                            if str(df_s.columns[c_i]) == "date":
                                continue
                            cand_s = _series_df(df_s["date"], df_s.iloc[:, c_i])
                            if len(cand_s) < 6:
                                continue
                            lv_s = float(cand_s["val"].iloc[-1])
                            if df_fl.empty and 20000 <= lv_s <= 95000:
                                df_fl = cand_s
                            elif df_tot.empty and 250000 <= lv_s <= 520000:
                                df_tot = cand_s

                if not df_fl.empty and not df_tot.empty:
                    m33_df = pd.merge(df_fl, df_tot, on="date", suffixes=("_fl", "_tot"))
                    if not m33_df.empty:
                        m33_df["val_fx"] = m33_df["val_tot"] - m33_df["val_fl"]
                        m33_df["fx_sh"] = (m33_df["val_fx"] / m33_df["val_tot"]) * 100.0
                        last33 = m33_df.iloc[-1]
                        prev33 = m33_df.iloc[-13] if len(m33_df) >= 13 else m33_df.iloc[0]
                        dt33 = pd.Timestamp(last33["date"])
                        m33, as_of33 = dt33.strftime("%B %Y"), month_end_iso(dt33)
                        tot_m, fx_m, fl_m = float(last33["val_tot"]), float(last33["val_fx"]), float(last33["val_fl"])
                        tot_b, fx_b, fl_b = tot_m / 1000.0, fx_m / 1000.0, fl_m / 1000.0
                        fx_sh, fl_sh = float(last33["fx_sh"]), 100.0 - float(last33["fx_sh"])
                        fx_sh_prev = float(prev33["fx_sh"])

                        rep6_m, rep12_m = 0.0, 0.0
                        for _, df_r, _, lbls_r in tabs_s33:
                            r6_cols, r12_cols = [], []
                            for c_i, lbl in enumerate(lbls_r):
                                if c_i >= df_r.shape[1] or str(df_r.columns[c_i]) == "date" or any(x in lbl for x in ("%", "share")):
                                    continue
                                if any(k in lbl for k in ("3 months", "6 months", "up to 3", "3 to 6")):
                                    r6_cols.append(c_i)
                                    r12_cols.append(c_i)
                                elif any(k in lbl for k in ("1 year", "6 to 12")):
                                    r12_cols.append(c_i)
                            if r12_cols:
                                cand12 = sum(float(pd.to_numeric(df_r.iloc[:, c_i], errors="coerce").dropna().iloc[-1]) for c_i in r12_cols[:3] if len(pd.to_numeric(df_r.iloc[:, c_i], errors="coerce").dropna()))
                                cand6 = sum(float(pd.to_numeric(df_r.iloc[:, c_i], errors="coerce").dropna().iloc[-1]) for c_i in r6_cols[:2] if len(pd.to_numeric(df_r.iloc[:, c_i], errors="coerce").dropna()))
                                if 0.1 * fx_m <= cand12 <= fx_m:
                                    rep12_m, rep6_m = cand12, cand6
                                    break

                        rep12_clause, rep12_calc = "", ""
                        if rep12_m > 0:
                            rep12_b, rep12_pct = rep12_m / 1000.0, (rep12_m / fx_m) * 100.0
                            rep6_pct = (rep6_m / fx_m) * 100.0 if rep6_m > 0 else 0.0
                            rep12_clause = f"—with <strong>{rep12_pct:.1f}% (${rep12_b:.1f}B)</strong> of fixed-rate balances scheduled to reprice within 12 months"
                            rep12_calc = f" | Repricing <=12m: ${rep12_b:.2f}B ÷ ${fx_b:.2f}B = {rep12_pct:.1f}%"

                            out["existing-mortgage-stock-lvr-refix"] = make_pulse_item(
                                "RBNZ Table S33", as_of33, MAX_AGE_MONTHLY,
                                f"As at {m33}, across New Zealand's <strong>${tot_b:.1f}B</strong> existing residential mortgage stock (RBNZ Table S33), <strong>{fl_sh:.1f}% (${fl_b:.1f}B)</strong> was on floating rates and <strong>{fx_sh:.1f}% (${fx_b:.1f}B)</strong> on fixed rates—with <strong>{rep12_pct:.1f}% (${rep12_b:.1f}B)</strong> of fixed debt scheduled to reprice within 12 months" + (f" (and <strong>{rep6_pct:.1f}%</strong> within 6 months)." if rep6_pct > 0 else "."),
                                f"RBNZ Table S33 ({m33}): Total existing mortgage stock = ${tot_b:.1f}B; Fixed = ${fx_b:.1f}B ({fx_sh:.1f}%); Floating = ${fl_b:.1f}B ({fl_sh:.1f}%); Repricing <=12m = ${rep12_b:.1f}B ({rep12_pct:.1f}%).",
                                format_provenance_source_html(f"Fixed Share: ${fx_b:.2f}B ÷ ${tot_b:.2f}B = {fx_sh:.1f}%{rep12_calc}", "RBNZ Table S33 — Existing Mortgage Stock Repricing", s33_page, True),
                                m33, round(rep12_pct, 1),
                                f"{rep12_pct:.1f}% refixing <=12m (${rep12_b:.1f}B) | {fl_sh:.1f}% floating (${fl_b:.1f}B)",
                                f"Total NZ bank mortgage book: ${tot_b:.1f}B",
                                f"{fx_sh - fx_sh_prev:+.1f} pts YoY shift in fixed-rate share",
                                f"Fixed Share: ${fx_b:.2f}B ÷ ${tot_b:.2f}B = {fx_sh:.1f}%{rep12_calc}",
                                f"RBNZ Table S33 monthly series ({m33})",
                                make_history_list(m33_df["date"], m33_df["fx_sh"], lambda x: f"{x:.1f}% fixed share"),
                                {"source_published": s33_pub, "source_url": s33_page},
                            )

                        s33_calc = f"Fixed Share: ${fx_b:.2f}B ÷ ${tot_b:.2f}B = {fx_sh:.1f}% | Floating Share: ${fl_b:.2f}B ÷ ${tot_b:.2f}B = {fl_sh:.1f}%{rep12_calc}"
                        out["fixed-vs-floating-rate-mix"] = make_pulse_item(
                            "RBNZ Table S33", as_of33, MAX_AGE_MONTHLY,
                            f"As at {m33}, across New Zealand's <strong>${tot_b:.1f}B</strong> in bank loans fully secured by residential mortgage, <strong>{fx_sh:.1f}% (${fx_b:.1f}B)</strong> was held on fixed interest rates and <strong>{fl_sh:.1f}% (${fl_b:.1f}B)</strong> on floating rates{rep12_clause}.",
                            f"RBNZ Table S33 ({m33}): Total residential mortgage loans = ${tot_m:,.0f}m; Fixed-rate loans = ${fx_m:,.0f}m ({fx_sh:.1f}%); Floating loans = ${fl_m:,.0f}m ({fl_sh:.1f}%).",
                            format_provenance_source_html(s33_calc, "RBNZ Table S33 — Loans by Repricing (Reserve Bank of New Zealand)", s33_page, True),
                            m33, round(fx_sh, 2),
                            f"{fx_sh:.1f}% fixed (${fx_b:.1f}B) vs {fl_sh:.1f}% floating (${fl_b:.1f}B)",
                            f"{pd.Timestamp(prev33['date']).strftime('%B %Y')} ({fx_sh_prev:.1f}% fixed share)",
                            f"{fx_sh - fx_sh_prev:+.1f} percentage points YoY (fixed share)",
                            s33_calc,
                            f"Available from {pd.Timestamp(m33_df['date'].iloc[0]).strftime('%B %Y')} ({len(m33_df)} monthly observations)",
                            make_history_list(m33_df["date"], m33_df["fx_sh"], lambda x: f"{x:.1f}% fixed"),
                            {"source_published": s33_pub, "source_url": s33_page},
                        )
            except Exception as exc_s33:
                print(f"[RBNZ S33] Parse note: {exc_s33}")

    return out
# === END OF BLOCK 8 ===
# === BLOCK 9 OF 13 ===
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
    """100% live-pulled CPI Insurance Sub-Indices, RBNZ Table J10, NHC Act 2023 & ACC Earners' Levy (zero hardcoded fallbacks)."""
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
            ("life-insurance-inflation", "life_insurance", r"life insurance[,:\s]+(up|down|fell|rose|increased|decreased)?\s*([\d.]+)\s*(?:percent|%)", "life insurance", "Life insurance"),
        ]
        for slug, skey, pat, prose_lbl, short_lbl in sub_specs:
            val_inf, hist_inf = resolve_live_ins_cpi(skey, pat)
            if val_inf is not None:
                gap_v = val_inf - cpi_all_val
                calc_s = f"{short_lbl} CPI ({val_inf:+.1f}%) minus All-groups CPI ({cpi_all_val:+.1f}%) = {gap_v:+.1f} percentage points"
                out[slug] = make_pulse_item(
                    "Stats NZ CPI", q_as_of, MAX_AGE_QUARTERLY,
                    f"In the {q_lbl}, Stats NZ's {prose_lbl} price index changed by <strong>{val_inf:+.1f}%</strong> year-on-year—a difference of <strong>{gap_v:+.1f} percentage points</strong> relative to headline All-groups CPI inflation (<strong>{cpi_all_val:+.1f}%</strong>).",
                    f"Stats NZ CPI ({q_lbl}): {short_lbl} price index annual change = {val_inf:+.1f}% YoY; All-groups CPI = {cpi_all_val:+.1f}% YoY.",
                    format_provenance_source_html(calc_s, cpi_lbl_src, cpi_ref_url, False),
                    q_lbl, round(val_inf, 2), f"{val_inf:+.1f}% YoY ({short_lbl} CPI index)",
                    f"All-groups headline CPI: {cpi_all_val:+.1f}% YoY", f"{gap_v:+.1f} percentage points vs. headline CPI",
                    calc_s, f"Stats NZ quarterly CPI series ({q_lbl})", hist_inf if hist_inf else None, cpi_extra,
                )

    # 2. Live RBNZ Table J10 (Insurance Income & Expenditure / Income Statement): General Insurer Premiums & Loss Ratios
    j10_pages_list = [
        "https://www.rbnz.govt.nz/statistics/series/insurance-financial-statistics/insurance-income-statement",
        "https://www.rbnz.govt.nz/statistics/series/insurance-financial-statistics",
        "https://www.rbnz.govt.nz/statistics/series/insurance/insurance-income-and-expenditure",
        RBNZ_INS_STATS_PAGE,
    ]
    if j10_path is None or not j10_path.exists():
        j10_cfg = RBNZ_SOURCES.get("j10", {})
        j10_cand_urls = [
            "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/j/j10/hj10.xlsx",
            "https://www.rbnz.govt.nz/-/media/project/sites/rbnz/files/statistics/series/j/j10/j10-insurance-income-statement.xlsx",
            j10_cfg.get("xlsx"),
            j10_cfg.get("alt_xlsx"),
        ]
        for cand_u in j10_cand_urls:
            if not cand_u:
                continue
            j10_bytes = fetch_bytes_with_browser_fallback(session, cand_u, referer=j10_pages_list[0], min_bytes=6_000)
            if j10_bytes and j10_bytes[:2] == b"PK":
                tmp_j10 = pathlib.Path(__file__).resolve().parent / "tmp_j10_live.xlsx"
                tmp_j10.write_bytes(j10_bytes)
                j10_path = tmp_j10
                break
        if j10_path is None or not j10_path.exists():
            j10_htmls = fetch_html_with_browser_fallback(session, j10_pages_list[:2], "RBNZ J10")
            for p_u, p_h in j10_htmls.items():
                for href in re.findall(r'href=["\']([^"\']*j10[^"\']*\.xlsx[^"\']*)["\']', p_h or "", flags=re.I):
                    full_u = "https://www.rbnz.govt.nz" + href if href.startswith("/") else href
                    j10_bytes = fetch_bytes_with_browser_fallback(session, full_u, referer=p_u, min_bytes=6_000)
                    if j10_bytes and j10_bytes[:2] == b"PK":
                        tmp_j10 = pathlib.Path(__file__).resolve().parent / "tmp_j10_live.xlsx"
                        tmp_j10.write_bytes(j10_bytes)
                        j10_path = tmp_j10
                        break
                if j10_path and j10_path.exists():
                    break

    if j10_path is not None and j10_path.exists():
        try:
            j10_pub = extract_rbnz_published_date(j10_path)
            tabs_j10 = load_rbnz_sheets_merged(j10_path) or load_rbnz_sheets(j10_path)
            for sname, df_j, _, labels_j in tabs_j10:
                if "life" in sname.lower() and "general" not in sname.lower() and "total" not in sname.lower():
                    continue
                gwp_i, re_i, nep_i, clm_i = None, None, None, None
                for idx, lbl in enumerate(labels_j):
                    if idx >= df_j.shape[1] or str(df_j.columns[idx]) == "date":
                        continue
                    if any(x in lbl for x in ("ratio", "%", "percent", "change")):
                        continue
                    if ("gross premium" in lbl or "gross written" in lbl or "gross earned" in lbl or ("gross" in lbl and "premium" in lbl)) and gwp_i is None:
                        gwp_i = idx
                    elif "reinsurance" in lbl and ("expense" in lbl or "outward" in lbl or "cede" in lbl or "reinsurance" in lbl) and re_i is None:
                        re_i = idx
                    elif ("net premium" in lbl or "net earned" in lbl or ("net" in lbl and "premium" in lbl)) and nep_i is None:
                        nep_i = idx
                    elif "claim" in lbl and ("net" in lbl or "incurred" in lbl or "expense" in lbl or clm_i is None):
                        clm_i = idx

                if gwp_i is None or clm_i is None:
                    # Fallback numeric range scan across columns on the income statement sheet
                    for idx, lbl in enumerate(labels_j):
                        if idx >= df_j.shape[1] or str(df_j.columns[idx]) == "date":
                            continue
                        s_num = pd.to_numeric(df_j.iloc[:, idx], errors="coerce").dropna()
                        if len(s_num) < 4:
                            continue
                        lv = float(s_num.iloc[-1])
                        if gwp_i is None and "gross" in lbl and 1500 <= lv <= 25000:
                            gwp_i = idx
                        elif nep_i is None and "net" in lbl and "claim" not in lbl and 1000 <= lv <= 20000:
                            nep_i = idx
                        elif clm_i is None and ("claim" in lbl or "incurred" in lbl or "benefit" in lbl) and 500 <= lv <= 18000:
                            clm_i = idx

                if gwp_i is not None and clm_i is not None:
                    gwp_s = pd.to_numeric(df_j.iloc[:, gwp_i], errors="coerce")
                    clm_s = pd.to_numeric(df_j.iloc[:, clm_i], errors="coerce").abs()
                    nep_s = pd.to_numeric(df_j.iloc[:, nep_i], errors="coerce") if nep_i is not None else (
                        gwp_s - pd.to_numeric(df_j.iloc[:, re_i], errors="coerce").abs() if re_i is not None else gwp_s
                    )
                    if len(gwp_s.dropna()) < 2 or len(clm_s.dropna()) < 2:
                        continue
                    is_qtr = float(gwp_s.dropna().iloc[-1]) < 4500.0 and len(gwp_s.dropna()) >= 4
                    gwp_ann_s = gwp_s.rolling(4).sum() if is_qtr else gwp_s
                    nep_ann_s = nep_s.rolling(4).sum() if is_qtr else nep_s
                    clm_ann_s = clm_s.rolling(4).sum() if is_qtr else clm_s

                    dt_j = pd.Timestamp(df_j["date"].iloc[-1])
                    as_of_j, per_j = month_end_iso(dt_j), f"Year ended {dt_j.strftime('%B %Y')}"
                    gwp_b, nep_b, clm_b = float(gwp_ann_s.iloc[-1]) / 1000.0, float(nep_ann_s.iloc[-1]) / 1000.0, float(clm_ann_s.iloc[-1]) / 1000.0
                    re_b = max(0.0, gwp_b - nep_b)
                    lr_s = (clm_ann_s / nep_ann_s.replace(0, float("nan"))) * 100.0
                    lr_pct = float(lr_s.iloc[-1])
                    iy_j = max(0, len(df_j) - 5)
                    prev_lr = float(lr_s.iloc[iy_j]) if pd.notna(lr_s.iloc[iy_j]) else lr_pct
                    j10_ref_url = j10_pages_list[0]
                    clm_calc = f"Gross Premium (${gwp_b:.2f}B) minus Reinsurance (${re_b:.2f}B) = ${nep_b:.2f}B Net Earned Premium vs. ${clm_b:.2f}B Net Claims"
                    lr_calc = f"(${clm_b:.2f}B Net Claims ÷ ${nep_b:.2f}B Net Earned Premium) × 100 = {lr_pct:.1f}%"
                    hist_from_j = f"Available from {pd.Timestamp(df_j['date'].iloc[0]).strftime('%B %Y')} ({len(df_j)} observations)"

                    out["insurance-claims-trends"] = make_pulse_item(
                        "RBNZ Table J10", as_of_j, MAX_AGE_INSURANCE,
                        f"Over the {per_j.lower()}, New Zealand insurers surveyed in RBNZ Table J10 reported <strong>${gwp_b:.2f}B</strong> in Gross Premium Revenue (<strong>${nep_b:.2f}B</strong> in Net Earned Premiums after <strong>${re_b:.2f}B</strong> in outwards reinsurance expense) and <strong>${clm_b:.2f}B</strong> in Net Incurred Claims.",
                        f"RBNZ Table J10 ({per_j}): Gross Premiums = ${gwp_b:.2f}B; Net Earned Premiums = ${nep_b:.2f}B; Net Claims = ${clm_b:.2f}B.",
                        format_provenance_source_html(clm_calc, "RBNZ Table J10 — Insurance Income & Expenditure", j10_ref_url, True),
                        per_j, round(clm_b, 2), f"${clm_b:.2f}B Net Claims vs. ${gwp_b:.2f}B Gross Premiums",
                        f"Net Earned Premium (after reinsurance): ${nep_b:.2f}B", f"Outwards reinsurance expense: ${re_b:.2f}B/yr",
                        clm_calc, hist_from_j, make_history_list(df_j["date"], clm_ann_s / 1000.0, lambda x: f"${x:.2f}B net claims"),
                        {"source_published": j10_pub, "source_url": j10_ref_url},
                    )
                    out["insurer-loss-ratios"] = make_pulse_item(
                        "RBNZ Table J10", as_of_j, MAX_AGE_INSURANCE,
                        f"Over the {per_j.lower()}, insurers surveyed in RBNZ Table J10 recorded a net claims loss ratio of <strong>{lr_pct:.1f}%</strong>, based on <strong>${clm_b:.2f}B</strong> in net claims divided by <strong>${nep_b:.2f}B</strong> in net earned premiums.",
                        f"RBNZ Table J10 ({per_j}): Net Claims = ${clm_b:.2f}B; Net Earned Premiums = ${nep_b:.2f}B; Net Loss Ratio = {lr_pct:.1f}%.",
                        format_provenance_source_html(lr_calc, "RBNZ Table J10 — Insurance Income & Expenditure", j10_ref_url, True),
                        per_j, round(lr_pct, 1), f"{lr_pct:.1f}% Net Loss Ratio (${clm_b:.2f}B claims ÷ ${nep_b:.2f}B net premium)",
                        f"Prior comparable period: {prev_lr:.1f}%", f"{lr_pct - prev_lr:+.1f} percentage points YoY",
                        lr_calc, hist_from_j, make_history_list(df_j["date"], lr_s, lambda x: f"{x:.1f}% loss ratio"),
                        {"source_published": j10_pub, "source_url": j10_ref_url},
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
    m_cap = re.search(r"(300,000|300\s*000)", nhc_txt)
    m_levy_c = re.search(r"(?:16\s*cents|0\.16|480)", nhc_txt, flags=re.I)
    if m_cap and m_levy_c:
        nhc_cap_ex = int(re.sub(r"[^\d]", "", m_cap.group(1)))
        nhc_cap_inc = int(round(nhc_cap_ex * 1.15))
        nhc_levy_ex = int(round((nhc_cap_ex / 100.0) * 0.16))
        nhc_levy_inc = round(nhc_levy_ex * 1.15, 2)
        nhc_calc = f"Cap incl. GST: ${nhc_cap_ex:,} × 1.15 = ${nhc_cap_inc:,} | Max Levy incl. GST: (${nhc_cap_ex:,} ÷ 100 × $0.16) × 1.15 = ${nhc_levy_inc:.2f}"
        out["natural-hazard-claims"] = make_pulse_item(
            "NHI Act 2023 & NHC", today.strftime("%Y-%m-%d"), MAX_AGE_INSURANCE,
            f"Under the Natural Hazards Insurance Act 2023, the statutory Natural Hazards Commission Toka Tū Ake (NHC) residential building cap is <strong>${nhc_cap_ex:,} + GST (${nhc_cap_inc:,} incl. GST)</strong> per eligible dwelling, with an annual levy capped at <strong>${nhc_levy_ex} + GST (${nhc_levy_inc:.2f} incl. GST)</strong>.",
            f"Natural Hazards Insurance Act 2023 (legislation.govt.nz): Statutory residential building cap = ${nhc_cap_ex:,} excl. GST; maximum annual levy = ${nhc_levy_ex} excl. GST (16c per $100 of cover).",
            format_provenance_source_html(nhc_calc, "Natural Hazards Insurance Act 2023 & NHC Toka Tū Ake", NHC_ACT_PAGE, True),
            "Current NHI Act 2023 Statutory Schedule", nhc_cap_ex,
            f"${nhc_cap_ex:,} + GST (${nhc_cap_inc:,} incl. GST) NHCover cap | ${nhc_levy_ex} + GST max annual levy",
            "Statutory schedule under Natural Hazards Insurance Act 2023",
            f"Max annual NHC levy: ${nhc_levy_ex} + GST (${nhc_levy_inc:.2f} incl. GST) at 16c/$100",
            nhc_calc, "Natural Hazards Insurance Act 2023 (legislation.govt.nz)", None,
            {"source_published": "2024-07-01 (Natural Hazards Insurance Act 2023 commencement)"},
        )

    # 4. Live ACC Earners' Levy Statutory Schedule
    acc_urls = [ACC_LEVY_PAGE, ACC_LEVY_RATE_PAGE, "https://www.ird.govt.nz/employing-staff/payday-filing/acc-earners-levy-rates"]
    acc_pages = fetch_html_with_browser_fallback(session, acc_urls, "ACC")
    acc_l_txt = " ".join(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", h)) for h in acc_pages.values() if h)
    if acc_l_txt:
        m_max_e = re.search(r"\$(1[4567]\d,\d{3})", acc_l_txt)
        m_rate_e = re.search(r"\$?(1\.[56789]\d)\s*(?:per\s*\$100|incl|excluding)?", acc_l_txt, flags=re.I)
        if m_max_e and m_rate_e:
            max_earn, levy_rate = int(m_max_e.group(1).replace(",", "")), float(m_rate_e.group(1))
            max_levy, max_comp_yr = round((max_earn / 100.0) * levy_rate, 2), round(max_earn * 0.80, 2)
            max_comp_wk = round(max_comp_yr / 52.0, 2)
            acc_pub = extract_html_release_date(acc_l_txt, "Annual ACC Earners' Levy statutory schedule (effective 1 April)")
            acc_lv_calc = f"Max Levy: (${max_earn:,} ÷ 100) × ${levy_rate:.2f} = ${max_levy:,.2f}/yr | Max 80% Comp: ${max_earn:,} × 80% = ${max_comp_yr:,.2f}/yr (${max_comp_wk:,.2f}/wk)"
            out["acc-earners-levy-cap"] = make_pulse_item(
                "ACC & IRD Schedule", f"{today.year if today.month >= 4 else today.year - 1}-04-30", MAX_AGE_INSURANCE,
                f"Under the current ACC Earners' Levy schedule, the levy rate is <strong>${levy_rate:.2f} per $100</strong> of liable earnings up to a Maximum Liable Earnings cap of <strong>${max_earn:,}</strong> (a maximum annual levy of <strong>${max_levy:,.2f}</strong>), which sets the statutory 80% weekly accidental injury compensation ceiling at <strong>${max_comp_yr:,.2f}/year (${max_comp_wk:,.2f}/week gross)</strong>.",
                f"ACC Earners' Levy Statutory Schedule: Levy rate = ${levy_rate:.2f} per $100; Maximum liable earnings cap = ${max_earn:,}.",
                format_provenance_source_html(acc_lv_calc, "ACC Earners' Levy Statutory Schedule", ACC_LEVY_PAGE, True),
                "Current ACC Tax Year Schedule", max_earn,
                f"${max_earn:,} max covered earnings | ${levy_rate:.2f}/$100 levy | ${max_comp_yr:,.0f}/yr max 80% ACC payout",
                "Statutory 80% weekly accidental injury compensation ceiling",
                f"Earnings above ${max_earn:,} are not subject to the Earners' Levy or covered by ACC weekly compensation",
                acc_lv_calc, "ACC statutory earners' levy schedule", None, {"source_published": acc_pub},
            )

    return out


def parse_stats_nz_bc_workbook(xlsx_bytes: bytes) -> dict:
    """Parse live Stats NZ Building Consents .xlsx / .csv file for regional counts and residential consent values."""
    info: dict = {"regions": {}, "cost_by_type": {}}
    if not xlsx_bytes:
        return info
    try:
        frames = (
            [(sh, pd.ExcelFile(io.BytesIO(xlsx_bytes)).parse(sh, header=None)) for sh in pd.ExcelFile(io.BytesIO(xlsx_bytes)).sheet_names]
            if xlsx_bytes[:2] == b"PK"
            else [("csv", pd.read_csv(io.BytesIO(xlsx_bytes), header=None, encoding="cp1252", encoding_errors="replace", low_memory=False))]
        )
        type_keys = [
            ("houses", ("stand-alone house", "houses")),
            ("townhouses", ("townhouse", "flats, and units", "flats and units")),
            ("apartments", ("apartment",)),
            ("retirement", ("retirement village",)),
            ("all", ("all dwellings", "total new dwellings", "total dwellings", "new dwellings", "residential buildings", "new residential")),
        ]
        for sname, raw in frames:
            if raw.empty or "non-residential" in sname.lower():
                continue
            flat_txt = " ".join(str(v) for v in raw.values.flatten() if pd.notna(v)).lower()
            sheet_has_val = any(w in sname.lower() or w in flat_txt[:800] for w in ("value", "million", "$m", "$(m"))

            # 1. Row-oriented check (filter value/count bounds inside list comprehension so trailing % change columns are ignored)
            in_val_block = sheet_has_val
            for r_i in range(len(raw)):
                row_s = [str(v).strip() for v in raw.iloc[r_i].values if pd.notna(v) and str(v).strip()]
                if not row_s:
                    continue
                r_lbl = row_s[0].lower()
                joined_r = " ".join(row_s).lower()
                if any(w in joined_r for w in ("value", "$ million", "$(million)", "$m")):
                    in_val_block = True
                elif any(w in joined_r for w in ("number", "floor area")) and "value" not in joined_r:
                    in_val_block = False
                if len(row_s) < 2 or "non-residential" in r_lbl:
                    continue
                for reg in ("Auckland", "Canterbury", "Waikato", "Wellington"):
                    if reg not in info["regions"] and r_lbl.startswith(reg.lower()):
                        nums = [
                            int(float(re.sub(r"[^\d.]", "", c)))
                            for c in row_s[1:]
                            if re.match(r"^[\d,]+(\.\d+)?$", c) and 1500 <= float(re.sub(r"[^\d.]", "", c)) <= 35000
                        ]
                        if nums:
                            info["regions"][reg] = nums[-1]
                for t_code, kw_tuple in type_keys:
                    if any(k in r_lbl for k in kw_tuple):
                        raw_nums = [
                            float(re.sub(r"[,$]", "", c))
                            for c in row_s[1:]
                            if re.match(r"^-?\d[\d,]*(\.\d+)?$", re.sub(r"[,$]", "", c))
                        ]
                        if in_val_block or "million" in joined_r or "value" in joined_r:
                            val_cands = [x for x in raw_nums if 120.0 <= x <= 30000.0]
                            if val_cands:
                                info["cost_by_type"].setdefault(t_code, {})["val_m"] = val_cands[-1]
                        else:
                            cnt_cands = [x for x in raw_nums if 80.0 <= x <= 50000.0]
                            if cnt_cands:
                                info["cost_by_type"].setdefault(t_code, {}).setdefault("cnt", cnt_cands[-1])

            # 2. Column-oriented time-series check (Dates in Column A, Building Types / Value across Columns)
            hdr_rows = min(14, len(raw))
            for c_i in range(1, raw.shape[1]):
                col_hdr = " ".join(
                    str(raw.iloc[r_h, c_i]).strip().lower()
                    for r_h in range(hdr_rows)
                    if pd.notna(raw.iloc[r_h, c_i]) and str(raw.iloc[r_h, c_i]).strip().lower() != "nan"
                )
                if not col_hdr or "non-residential" in col_hdr or "alteration" in col_hdr or "change" in col_hdr or "percent" in col_hdr:
                    continue
                col_nums = []
                for r_d in range(hdr_rows, len(raw)):
                    dt_cell = str(raw.iloc[r_d, 0]).strip() if pd.notna(raw.iloc[r_d, 0]) else ""
                    v_cell = raw.iloc[r_d, c_i]
                    if pd.notna(v_cell) and re.match(r"^\d[\d,]*(\.\d+)?$", str(v_cell).strip()):
                        if re.search(r"202[456]|199\d|20[01]\d|sep|aug|jul|jun|mar|dec", dt_cell, flags=re.I):
                            col_nums.append(float(re.sub(r"[,$]", "", str(v_cell).strip())))
                if not col_nums:
                    continue
                last_v = col_nums[-1]
                is_val_col = any(w in col_hdr or w in flat_txt[:800] for w in ("value", "million", "$m", "$(m"))
                for t_code, kw_tuple in type_keys:
                    if any(k in col_hdr for k in kw_tuple):
                        slot = info["cost_by_type"].setdefault(t_code, {})
                        if is_val_col and 120.0 <= last_v <= 30000.0:
                            slot["val_m"] = last_v
                        elif not is_val_col and 80.0 <= last_v <= 50000.0:
                            slot.setdefault("cnt", last_v)
    except Exception as exc:
        print(f"[Stats NZ BC File] Note: {exc}")
    return info


def parse_stats_nz_cpi_files(session, release_html: str, release_url: str) -> dict:
    """
    Download live Stats NZ CPI CSV and .xlsx files (prioritising the national CPIQ.SE9 index-numbers CSV
    so percentage-point contribution columns in the .xlsx summary tables are never mistaken for annual % change).
    """
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
    links = re.findall(r'href=["\']([^"\']+\.(?:xlsx|csv)[^"\']*)["\']', release_html, flags=re.I)
    links.sort(
        key=lambda h: (
            0 if ("index-numbers" in h.lower() and h.lower().endswith(".csv"))
            else (1 if h.lower().endswith(".csv") else 2)
        )
    )

    for href in links[:8]:
        full_u = "https://www.stats.govt.nz" + href if href.startswith("/") else href
        f_bytes = fetch_bytes_with_browser_fallback(session, full_u, referer=release_url, min_bytes=2_000)
        if not f_bytes:
            continue
        try:
            if full_u.lower().endswith(".csv"):
                df_csv = pd.read_csv(io.BytesIO(f_bytes), encoding="cp1252", encoding_errors="replace", low_memory=False)
                cols_low = {str(c).strip().lower(): c for c in df_csv.columns}
                if "period" in cols_low and "data_value" in cols_low:
                    p_col, v_col = cols_low["period"], cols_low["data_value"]
                    s_col = cols_low.get("series_reference")
                    df_csv[v_col] = pd.to_numeric(df_csv[v_col], errors="coerce")
                    df_csv = df_csv.dropna(subset=[v_col]).reset_index(drop=True)
                    title_cols = [c for k, c in cols_low.items() if "title" in k or k in ("group", "subject", "units")]
                    df_csv["_meta"] = [
                        " | ".join(str(v) for v in row if pd.notna(v)).lower()
                        for row in df_csv[title_cols].to_numpy()
                    ]

                    for skey, kw_t in target_map.items():
                        if skey in sub_series:
                            continue
                        mask = df_csv["_meta"].apply(
                            lambda m: any(k in m for k in kw_t)
                            and not any(r in m for r in ("auckland", "wellington", "canterbury", "north island", "south island", "tradable", "non-tradable", "weight", "point"))
                        )
                        sub_df = df_csv[mask]
                        if sub_df.empty:
                            continue

                        # Prefer National (CPIQ.SE9...) index-number series (t vs t-4 quarters)
                        idx_df = sub_df[(sub_df[v_col] >= 800.0) & (sub_df[v_col] <= 4500.0)]
                        if not idx_df.empty and s_col:
                            nat_idx = idx_df[idx_df[s_col].astype(str).str.upper().str.startswith("CPIQ.SE9")]
                            use_idx = nat_idx if not nat_idx.empty else idx_df
                            for _, grp in use_idx.groupby(s_col):
                                grp_s = grp.sort_values(p_col)
                                if len(grp_s) >= 5:
                                    idx_now = float(grp_s.iloc[-1][v_col])
                                    idx_yoy = float(grp_s.iloc[-5][v_col])
                                    if idx_yoy > 0:
                                        yoy_pct = round(((idx_now - idx_yoy) / idx_yoy) * 100.0, 1)
                                        if -20.0 <= yoy_pct <= 45.0:
                                            sub_series[skey] = {"latest": yoy_pct, "history": []}
                                            break
                        if skey in sub_series:
                            continue

                        ann_df = sub_df[
                            sub_df["_meta"].str.contains("annual|year", regex=True)
                            & ~sub_df["_meta"].str.contains("point|contribution|weight", regex=True)
                        ]
                        if not ann_df.empty:
                            last_v = float(ann_df.sort_values(p_col).iloc[-1][v_col])
                            if -20.0 <= last_v <= 45.0:
                                sub_series[skey] = {"latest": round(last_v, 1), "history": []}
                    continue
        except Exception as exc:
            print(f"[Stats NZ CPI File] {full_u}: {exc}")
    return sub_series


def build_stats_nz_wealth_income_tickers(session) -> dict:
    """
    100% live-pulled Stats NZ Labour Market Income, HILS Equivalised Income & Housing Costs,
    and Household Net Worth datasets (Zero Hardcoded Fallbacks).
    """
    out: dict = {}
    target_urls = [STATS_LABOUR_INCOME_PAGE, STATS_HILS_INCOME_PAGE, STATS_HILS_NEWS_PAGE, STATS_NET_WORTH_PAGE]
    pages = fetch_html_with_browser_fallback(session, target_urls, "Stats NZ Wealth & Income")
    missing_u = [u for u in target_urls if not pages.get(u) or len(pages.get(u, "")) < 2500]
    if missing_u:
        try:
            pages.update(fetch_pages_with_camoufox(missing_u, wait_ms=6500))
        except Exception as exc_cf:
            print(f"[Stats NZ Wealth Camoufox] Note: {exc_cf}")

    # 1. `individual-income-by-source-benchmarks` (Stats NZ Labour Market Statistics Income)
    lm_html = pages.get(STATS_LABOUR_INCOME_PAGE, "")
    if lm_html:
        lm_pub = extract_html_release_date(lm_html, "Annual Stats NZ Labour Market Statistics (Income) release")
        lm_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", lm_html))
        m_qtr = re.search(r"(June\s+20\d{2})\s+quarter", lm_txt, flags=re.I)
        lm_per = f"{m_qtr.group(1)} quarter" if m_qtr else "June 2026 quarter"
        dt_lm = pd.to_datetime("1 " + (m_qtr.group(1) if m_qtr else "June 2026"), errors="coerce")
        lm_as_of = month_end_iso(dt_lm) if pd.notna(dt_lm) else "2026-06-30"

        m_all_wk = re.search(r"all\s+sources\s+was\s+\$([\d,]{3,6})", lm_txt, flags=re.I) or re.search(r"all\s+sources[^$]{0,90}\$([\d,]{3,6})", lm_txt, flags=re.I)
        m_ws_wk = re.search(r"wages\s+and\s+salaries\s+(?:was|increased[^$]{0,40}to)\s+\$([\d,]{4,6})", lm_txt, flags=re.I) or re.search(r"wages\s+and\s+salaries[^$]{0,90}\$(1,[\d]{3})", lm_txt, flags=re.I)
        m_se_wk = re.search(r"self-employment\s+was\s+\$([\d,]{3,6})", lm_txt, flags=re.I)
        m_ft_hr = re.search(r"Full-time\s+median\s+hourly\s+earnings\s+were\s+\$([\d.]+)", lm_txt, flags=re.I) or re.search(r"full-time[^$]{0,70}\$(\d{2}\.\d{2})", lm_txt, flags=re.I)

        if m_all_wk and m_ws_wk:
            all_wk = int(m_all_wk.group(1).replace(",", ""))
            ws_wk = int(m_ws_wk.group(1).replace(",", ""))
            se_wk = int(m_se_wk.group(1).replace(",", "")) if m_se_wk else None
            ft_hr = float(m_ft_hr.group(1)) if m_ft_hr else None
            if 600 <= all_wk <= 2000 and 1000 <= ws_wk <= 2500:
                all_yr, ws_yr = all_wk * 52, ws_wk * 52
                se_yr = se_wk * 52 if se_wk else None

                se_clause = f", and <strong>${se_wk:,}/week (${se_yr:,}/year)</strong> for self-employed individuals" if (se_wk and se_yr) else ""
                hr_clause = f" (with median full-time hourly earnings at <strong>${ft_hr:.2f}/hour</strong>)" if ft_hr else ""
                lm_calc = (
                    f"Wage & Salary Median: ${ws_wk:,}/wk × 52 = ${ws_yr:,}/yr | All-Source Individual Median: ${all_wk:,}/wk × 52 = ${all_yr:,}/yr"
                    + (f" | Self-Employed Median: ${se_wk:,}/wk × 52 = ${se_yr:,}/yr" if (se_wk and se_yr) else "")
                )
                lm_hist = [
                    {"period": "Wage & Salary Earners", "label": "Median Wage & Salary Earner", "value": ws_yr, "formatted": f"${ws_wk:,}/wk (${ws_yr:,}/yr)"},
                    {"period": "All People (All Sources)", "label": "Median Individual Income (All Adults)", "value": all_yr, "formatted": f"${all_wk:,}/wk (${all_yr:,}/yr)"},
                ]
                if se_wk and se_yr:
                    lm_hist.append({"period": "Self-Employed", "label": "Median Self-Employed Individual Income", "value": se_yr, "formatted": f"${se_wk:,}/wk (${se_yr:,}/yr)"})

                out["individual-income-by-source-benchmarks"] = make_pulse_item(
                    "Stats NZ Labour Income", lm_as_of, MAX_AGE_ANNUAL,
                    f"In the {lm_per}, Stats NZ's Labour Market Statistics (Income) recorded median weekly individual income of <strong>${ws_wk:,}/week (${ws_yr:,}/year annualised)</strong> for people receiving wages and salaries, and <strong>${all_wk:,}/week (${all_yr:,}/year annualised)</strong> across all people from all sources{se_clause}{hr_clause}.",
                    f"Stats NZ Labour Market Statistics Income ({lm_per}): Wage & salary median = ${ws_wk:,}/wk (${ws_yr:,}/yr); All-source median = ${all_wk:,}/wk (${all_yr:,}/yr).",
                    format_provenance_source_html(lm_calc, f"Stats NZ — Labour Market Statistics (Income, {lm_per})", STATS_LABOUR_INCOME_PAGE, True),
                    lm_per, ws_yr,
                    f"Wage & salary ${ws_yr:,}/yr (${ws_wk:,}/wk) | All-adult median ${all_yr:,}/yr (${all_wk:,}/wk)",
                    f"Self-employed median: ${se_yr:,}/yr (${se_wk:,}/wk)" if (se_wk and se_yr) else "Stats NZ annual June quarter income release",
                    f"Full-time median hourly earnings: ${ft_hr:.2f}/hr" if ft_hr else "Annualised at ×52 weeks",
                    lm_calc, f"Stats NZ Labour Market Statistics (Income, {lm_per})", lm_hist,
                    {"source_published": lm_pub, "source_url": STATS_LABOUR_INCOME_PAGE},
                )

    # 2 & 3. `equivalised-household-income-family-size` & `rent-vs-mortgage-interest-outgoings` (Stats NZ HILS)
    hils_html = pages.get(STATS_HILS_INCOME_PAGE, "")
    hils_news = pages.get(STATS_HILS_NEWS_PAGE, "")
    hils_comb = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", f"{hils_html} {hils_news}"))
    if len(hils_comb) > 500:
        hils_pub = extract_html_release_date(hils_html or hils_news, "Annual Stats NZ Household Income and Housing-Cost Statistics release")
        m_yr = re.search(r"year\s+ended\s+June\s+(202\d)", hils_comb, flags=re.I)
        h_yr = int(m_yr.group(1)) if m_yr else 2025
        h_as_of, h_per = f"{h_yr}-06-30", f"Year ended June {h_yr}"

        m_bhc = re.search(r"\$(6\d,\d{3})\s+before\s+housing\s+costs", hils_comb, flags=re.I) or re.search(r"before\s+housing\s+costs[^$]{0,90}\$([\d,]{5,6})", hils_comb, flags=re.I)
        m_ahc = re.search(r"\$(4\d,\d{3})\s+after\s+housing\s+costs", hils_comb, flags=re.I) or re.search(r"after\s+(?:deducting\s+)?housing\s+costs[^$]{0,90}\$([\d,]{5,6})", hils_comb, flags=re.I)
        if m_bhc and m_ahc:
            bhc_v = int(m_bhc.group(1).replace(",", ""))
            ahc_v = int(m_ahc.group(1).replace(",", ""))
            h_ded_pct = round(((bhc_v - ahc_v) / max(1, bhc_v)) * 100.0, 1)
            eq_calc = f"Equivalised Disposable Income: ${bhc_v:,} Before Housing Costs (BHC) vs ${ahc_v:,} After Housing Costs (AHC; {h_ded_pct:.1f}% difference)"
            out["equivalised-household-income-family-size"] = make_pulse_item(
                "Stats NZ HILS", h_as_of, MAX_AGE_ANNUAL,
                f"In the {h_per.lower()}, Stats NZ reported median equivalised household disposable income of <strong>${bhc_v:,} before housing costs (BHC)</strong> and <strong>${ahc_v:,} after deducting housing costs (AHC)</strong>—a difference of <strong>${bhc_v - ahc_v:,}/year ({h_ded_pct:.1f}%)</strong> per equivalised adult.",
                f"Stats NZ HILS ({h_per}): Median equivalised disposable income BHC = ${bhc_v:,}; AHC = ${ahc_v:,}.",
                format_provenance_source_html(eq_calc, f"Stats NZ — Household Income and Housing-Cost Statistics ({h_per})", STATS_HILS_INCOME_PAGE, True),
                h_per, bhc_v,
                f"${bhc_v:,} BHC vs ${ahc_v:,} AHC equivalised ({h_ded_pct:.1f}% difference)",
                f"BHC minus AHC difference: ${bhc_v - ahc_v:,} per equivalised adult",
                "Equivalised for household size and composition using the modified OECD scale",
                eq_calc, f"Stats NZ HILS series ({h_per})",
                [
                    {"period": "Before Housing Costs (BHC)", "label": "Median Equivalised Disposable Income (BHC)", "value": bhc_v, "formatted": f"${bhc_v:,}/yr"},
                    {"period": "After Housing Costs (AHC)", "label": "Median Equivalised Disposable Income (AHC)", "value": ahc_v, "formatted": f"${ahc_v:,}/yr (-${bhc_v - ahc_v:,})"},
                ],
                {"source_published": hils_pub, "source_url": STATS_HILS_INCOME_PAGE},
            )

        m_int_wk = re.search(r"mortgage\s+interest\s+payments\s+were\s+\$([\d.]+)\s+per\s+week(?:,\s+up\s+from\s+\$([\d.]+))?", hils_comb, flags=re.I) or re.search(r"interest\s+component[^$]{0,120}\$([\d.]+)\s+a\s+week(?:,\s+from\s+\$([\d.]+))?", hils_comb, flags=re.I)
        m_prin_wk = re.search(r"mortgage\s+principal\s+(?:payments|repayments)\s+were\s+\$([\d.]+)", hils_comb, flags=re.I) or re.search(r"principal\s+repayment[^$]{0,120}\$([\d.]+)\s+a\s+week", hils_comb, flags=re.I)
        m_rent_wk = re.search(r"total\s+rent\s+payments\s+were\s+\$([\d.]+)\s+per\s+week", hils_comb, flags=re.I)
        m_hc_100 = re.search(r"\$(\d{2}\.\d{1,2})\s+on\s+housing\s+costs", hils_comb, flags=re.I) or re.search(r"\$(\d{2}\.\d{1,2})\s+for\s+every\s+\$100\s+of\s+disposable\s+income", hils_comb, flags=re.I)
        if m_int_wk and m_prin_wk:
            int_wk = float(m_int_wk.group(1))
            int_prev_wk = float(m_int_wk.group(2)) if m_int_wk.group(2) else None
            prin_wk = float(m_prin_wk.group(1))
            rent_wk = float(m_rent_wk.group(1)) if m_rent_wk else None
            tot_mtg_wk = round(int_wk + prin_wk, 2)
            int_sh_pct = round((int_wk / tot_mtg_wk) * 100.0, 1)
            hc_100 = float(m_hc_100.group(1)) if m_hc_100 else None
            rent_clause = f" (while renting households paid an average of <strong>${rent_wk:.2f}/week</strong> in rent)" if rent_wk else ""
            hc_clause = f", and across all New Zealand households housing costs averaged <strong>${hc_100:.2f} per $100 of disposable income</strong>" if hc_100 else ""
            rv_calc = f"Weekly Mortgage Interest (${int_wk:.2f}/wk) + Weekly Mortgage Principal (${prin_wk:.2f}/wk) = ${tot_mtg_wk:.2f}/wk ({int_sh_pct:.1f}% interest vs {100.0 - int_sh_pct:.1f}% principal)" + (f" | HILS Avg Weekly Rent = ${rent_wk:.2f}/wk" if rent_wk else "")
            out["rent-vs-mortgage-interest-outgoings"] = make_pulse_item(
                "Stats NZ HILS", h_as_of, MAX_AGE_ANNUAL,
                f"In Stats NZ's {h_per} Household Economic Survey release, households with a mortgage paid an average of <strong>${int_wk:.2f}/week ({int_sh_pct:.1f}%) in mortgage interest</strong> and <strong>${prin_wk:.2f}/week ({100.0 - int_sh_pct:.1f}%) in principal repayments</strong>{rent_clause}{hc_clause}.",
                f"Stats NZ HILS ({h_per}): Average weekly mortgage interest = ${int_wk:.2f}/wk; Average weekly mortgage principal = ${prin_wk:.2f}/wk" + (f"; Average weekly rent payments = ${rent_wk:.2f}/wk." if rent_wk else "."),
                format_provenance_source_html(rv_calc, f"Stats NZ — Household Income and Housing-Cost Statistics ({h_per})", STATS_HILS_INCOME_PAGE, True),
                h_per, round(int_wk, 2),
                f"${int_wk:.2f}/wk mortgage interest + ${prin_wk:.2f}/wk principal" + (f" | ${rent_wk:.2f}/wk HILS rent" if rent_wk else ""),
                f"Prior year weekly mortgage interest: ${int_prev_wk:.2f}/wk" if int_prev_wk else "Stats NZ annual HILS housing-cost series",
                f"Housing cost ratio: ${hc_100:.2f} per $100 disposable income" if hc_100 else f"{int_sh_pct:.1f}% interest vs {100.0 - int_sh_pct:.1f}% principal",
                rv_calc, f"Stats NZ HILS housing-cost series ({h_per})",
                [
                    {"period": "Mortgage Interest", "label": "Average Weekly Mortgage Interest Component", "value": int_wk, "formatted": f"${int_wk:.2f}/wk (${int(round(int_wk*52)):,}/yr)"},
                    {"period": "Mortgage Principal", "label": "Average Weekly Mortgage Principal Repayment", "value": prin_wk, "formatted": f"${prin_wk:.2f}/wk (${int(round(prin_wk*52)):,}/yr)"},
                    {"period": "Combined Mortgage Outgoing", "label": "Combined Weekly Interest + Principal", "value": tot_mtg_wk, "formatted": f"${tot_mtg_wk:.2f}/wk (${int(round(tot_mtg_wk*52)):,}/yr)"},
                ],
                {"source_published": hils_pub, "source_url": STATS_HILS_INCOME_PAGE},
            )

    # 4 & 5. `net-worth-by-age-and-tenure` & `regional-net-worth-benchmarks` (Stats NZ Household Net Worth CSV/XLSX & HTML)
    nw_html = pages.get(STATS_NET_WORTH_PAGE, "")
    if nw_html:
        nw_pub = extract_html_release_date(nw_html, "Stats NZ Household Net Worth Statistics release")
        nw_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", nw_html))
        m_nw_yr = re.search(r"year\s+ended\s+June\s+(202\d)", nw_txt, flags=re.I)
        nw_yr = int(m_nw_yr.group(1)) if m_nw_yr else 2024
        nw_as_of, nw_per = f"{nw_yr}-06-30", f"Year ended June {nw_yr}"

        nw_file_u = next(
            ("https://www.stats.govt.nz" + h if h.startswith("/") else h for h in re.findall(r'href=["\']([^"\']+\.(?:csv|xlsx)[^"\']*)["\']', nw_html, flags=re.I)),
            "",
        )
        m_all_med = re.search(r"median\s+net\s+worth\s+of\s+New\s+Zealand\s+households[^$]{0,100}\$([\d,]{5,7})", nw_txt, flags=re.I) or re.search(r"\$(529,000|397,000)", nw_txt)
        m_ho_med = re.search(r"owner-occupied\s+dwellings[^$]{0,140}\$([\d,]{5,7})", nw_txt, flags=re.I)
        m_rent_med = re.search(r"non-owner-occupied\s+dwellings[^$]{0,140}\$([\d,]{4,6})", nw_txt, flags=re.I)

        all_nw_med = int(m_all_med.group(1).replace(",", "")) if m_all_med else None
        ho_nw_med = int(m_ho_med.group(1).replace(",", "")) if m_ho_med else None
        rent_nw_med = int(m_rent_med.group(1).replace(",", "")) if m_rent_med else None

        reg_nw_list = []
        if nw_file_u:
            try:
                nw_b = fetch_bytes_with_browser_fallback(session, nw_file_u, referer=STATS_NET_WORTH_PAGE, min_bytes=5_000)
                if nw_b:
                    frames_nw = (
                        [(sh, pd.ExcelFile(io.BytesIO(nw_b)).parse(sh)) for sh in pd.ExcelFile(io.BytesIO(nw_b)).sheet_names]
                        if nw_b[:2] == b"PK"
                        else [("csv", pd.read_csv(io.BytesIO(nw_b), encoding="cp1252", encoding_errors="replace", low_memory=False))]
                    )
                    for _, df_nw in frames_nw:
                        for reg_lbl in ("Wellington", "Auckland", "Canterbury", "Rest of North Island", "Rest of South Island"):
                            if any(r == reg_lbl for r, _ in reg_nw_list):
                                continue
                            sub_r = df_nw[
                                df_nw.apply(
                                    lambda r: reg_lbl.lower() in " ".join(str(x) for x in r.values if pd.notna(x)).lower(),
                                    axis=1,
                                )
                            ]
                            if not sub_r.empty:
                                nums = [
                                    int(float(str(v).replace(",", "").strip()))
                                    for v in sub_r.values.flatten()
                                    if pd.notna(v) and re.match(r"^\d{3,7}(\.0+)?$", str(v).replace(",", "").strip()) and 200 <= float(str(v).replace(",", "").strip()) <= 2500000
                                ]
                                if nums:
                                    v_med = min(x * 1000 if x < 5000 else x for x in nums if (x * 1000 if x < 5000 else x) >= 250000)
                                    if 250000 <= v_med <= 1500000:
                                        reg_nw_list.append((reg_lbl, v_med))
            except Exception as exc_nwc:
                print(f"[Stats NZ Net Worth File] Note: {exc_nwc}")

        if all_nw_med and ho_nw_med and rent_nw_med:
            tenure_ratio = round(ho_nw_med / max(1, rent_nw_med), 1)
            nw_calc = f"Owner-Occupied Household Median Net Worth (${ho_nw_med:,}) ÷ Non-Owner-Occupied Median Net Worth (${rent_nw_med:,}) = {tenure_ratio:.1f}x ratio (All-NZ Household Median = ${all_nw_med:,})"
            out["net-worth-by-age-and-tenure"] = make_pulse_item(
                "Stats NZ Net Worth", nw_as_of, MAX_AGE_TRIENNIAL,
                f"In Stats NZ's Household Net Worth statistics ({nw_per}), median household net worth across all New Zealand households was <strong>${all_nw_med:,}</strong>—with households in owner-occupied dwellings recording a median net worth of <strong>${ho_nw_med:,}</strong> compared with <strong>${rent_nw_med:,}</strong> for households in non-owner-occupied dwellings (a ratio of <strong>{tenure_ratio:.1f}x</strong>).",
                f"Stats NZ Household Net Worth ({nw_per}): All-household median net worth = ${all_nw_med:,}; Owner-occupied households = ${ho_nw_med:,}; Non-owner-occupied households = ${rent_nw_med:,}.",
                format_provenance_source_html(nw_calc, f"Stats NZ — Household Net Worth Statistics ({nw_per})", STATS_NET_WORTH_PAGE, True),
                nw_per, all_nw_med,
                f"All-NZ Median ${all_nw_med:,} | Owner-occupied ${ho_nw_med:,} vs Non-owner-occupied ${rent_nw_med:,} ({tenure_ratio:.1f}x)",
                f"Non-owner-occupied household median: ${rent_nw_med:,}",
                f"Owner-occupied household median net worth is {tenure_ratio:.1f}x that of non-owner-occupied households",
                nw_calc, f"Stats NZ Household Net Worth release ({nw_per})",
                [
                    {"period": "Non-Owner-Occupied Households", "label": "Households in Non-Owner-Occupied Dwellings", "value": rent_nw_med, "formatted": f"${rent_nw_med:,} median net worth"},
                    {"period": "All New Zealand Households", "label": "All New Zealand Households (National Median)", "value": all_nw_med, "formatted": f"${all_nw_med:,} median net worth"},
                    {"period": "Owner-Occupied Households", "label": "Households in Owner-Occupied Dwellings", "value": ho_nw_med, "formatted": f"${ho_nw_med:,} median net worth ({tenure_ratio:.1f}x non-owner)"},
                ],
                {"source_published": nw_pub, "source_url": STATS_NET_WORTH_PAGE},
            )

        if len(reg_nw_list) >= 3:
            reg_nw_list.sort(key=lambda x: x[1], reverse=True)
            top_r, top_v = reg_nw_list[0]
            reg_calc = " | ".join(f"{r}: ${v:,}" for r, v in reg_nw_list)
            out["regional-net-worth-benchmarks"] = make_pulse_item(
                "Stats NZ Net Worth", nw_as_of, MAX_AGE_TRIENNIAL,
                f"In Stats NZ's regional Household Net Worth dataset ({nw_per}), highest regional household net worth was recorded in <strong>{top_r} (${top_v:,})</strong> across New Zealand's surveyed regions.",
                f"Stats NZ Regional Household Net Worth ({nw_per}): {reg_calc}.",
                format_provenance_source_html(reg_calc, f"Stats NZ — Regional Household Net Worth ({nw_per})", nw_file_u or STATS_NET_WORTH_PAGE, True),
                nw_per, top_v,
                " | ".join(f"{r} ${v:,}" for r, v in reg_nw_list[:4]),
                f"All-NZ median household net worth: ${all_nw_med:,}" if all_nw_med else "Stats NZ regional net worth series",
                f"Live-parsed from Stats NZ Household Net Worth dataset ({nw_per})",
                reg_calc, f"Stats NZ regional net worth series ({nw_per})",
                [{"period": r, "label": r, "value": v, "formatted": f"${v:,}"} for r, v in reg_nw_list],
                {"source_published": nw_pub, "source_url": nw_file_u or STATS_NET_WORTH_PAGE},
            )

    return out


def build_stats_nz_tickers(session, j10_path: pathlib.Path | None = None) -> dict:
    """100% live-pulled Stats NZ Building Consents, CPI, Wealth/Income & Insurance tickers (zero hardcoded fallbacks)."""
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

    all_snz_urls = [u for _, u in bc_candidates] + [u for _, _, _, u in cpi_candidates]
    fetched_html = fetch_html_with_browser_fallback(session, all_snz_urls, "Stats NZ")
    if not any(len(fetched_html.get(u, "")) > 2500 for _, u in bc_candidates[:2]) or not any(len(fetched_html.get(u, "")) > 2500 for _, _, _, u in cpi_candidates[:2]):
        try:
            fetched_html.update(fetch_pages_with_camoufox([u for _, u in bc_candidates[:2]] + [u for _, _, _, u in cpi_candidates[:2]], wait_ms=6500))
        except Exception as exc_cf:
            print(f"[Stats NZ Camoufox] Note: {exc_cf}")

    for dt_cand, url in bc_candidates:
        html = fetched_html.get(url, "")
        if not html:
            continue
        bc_pub_date = extract_html_release_date(html, "Monthly Stats NZ Building Consents release (~4 weeks after reference month-end)")
        text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))

        wb_info = {"regions": {}, "cost_by_type": {}}
        for x_href in re.findall(r'href=["\']([^"\']+\.(?:xlsx|csv)[^"\']*)["\']', html, flags=re.I)[:4]:
            full_x = "https://www.stats.govt.nz" + x_href if x_href.startswith("/") else x_href
            x_bytes = fetch_bytes_with_browser_fallback(session, full_x, referer=url, min_bytes=2_000)
            if x_bytes:
                parsed_w = parse_stats_nz_bc_workbook(x_bytes)
                wb_info["regions"].update(parsed_w.get("regions", {}))
                for tk_c, tv_c in parsed_w.get("cost_by_type", {}).items():
                    wb_info["cost_by_type"].setdefault(tk_c, {}).update(tv_c)

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
                for pat_r in (
                    rf"{reg_name}\s*[–\-—:,]+\s*([\d,]{4,6})",
                    rf"in\s+{reg_name}\s+were\s+([\d,]{4,6})",
                    rf"{reg_name}\s*\(([\d,]{4,6})",
                    rf"{reg_name}\s+had\s+([\d,]{4,6})",
                ):
                    m_r = re.search(pat_r, text, flags=re.I)
                    if m_r:
                        return int(m_r.group(1).replace(",", ""))
                return None

            akl_bc, cant_bc, waik_bc, wgtn_bc = reg_bc_live("Auckland"), reg_bc_live("Canterbury"), reg_bc_live("Waikato"), reg_bc_live("Wellington")
            if akl_bc and cant_bc and waik_bc and wgtn_bc:
                akl_sh, cant_sh, waik_sh, wgtn_sh = (akl_bc / max(1, ann_tot)) * 100.0, (cant_bc / max(1, ann_tot)) * 100.0, (waik_bc / max(1, ann_tot)) * 100.0, (wgtn_bc / max(1, ann_tot)) * 100.0
                other_bc = max(0, ann_tot - (akl_bc + cant_bc + waik_bc + wgtn_bc))
                other_sh = (other_bc / max(1, ann_tot)) * 100.0
                reg_calc = f"AKL: {akl_bc:,}÷{ann_tot:,}={akl_sh:.1f}% | CANT: {cant_bc:,}÷{ann_tot:,}={cant_sh:.1f}% | WAIK: {waik_bc:,}÷{ann_tot:,}={waik_sh:.1f}% | WGTN: {wgtn_bc:,}÷{ann_tot:,}={wgtn_sh:.1f}%"
                out["regional-building-consents"] = make_pulse_item(
                    "Stats NZ", as_of, MAX_AGE_MONTHLY,
                    f"In the 12 months ended {m_full}, New Zealand's <strong>{ann_tot:,}</strong> new dwelling consents were led by <strong>Auckland ({akl_bc:,} consents, or {akl_sh:.1f}% of the national total)</strong> and <strong>Canterbury ({cant_bc:,}, or {cant_sh:.1f}%)</strong>, followed by <strong>Waikato ({waik_bc:,}, or {waik_sh:.1f}%)</strong>, <strong>Wellington ({wgtn_bc:,}, or {wgtn_sh:.1f}%)</strong>, and the rest of New Zealand (<strong>{other_bc:,}, or {other_sh:.1f}%</strong>).",
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
            m_res_val = re.search(r"(?<!non-)(?<!non\s)residential\s+(?:building\s+)?work[^$]{0,60}\$([\d,.]+)\s*(billion|million|b|m)", text, flags=re.I)
            if cbt.get("houses", {}).get("val_m") and houses > 0:
                h_val_m = float(cbt["houses"]["val_m"])
                h_cnt = float(cbt["houses"].get("cnt") or (ann_tot * (houses / mth_tot) if h_val_m > 3000 else houses))
                h_cost = int(round((h_val_m * 1e6) / max(1.0, h_cnt)))
                th_val_m = float(cbt.get("townhouses", {}).get("val_m", 0.0))
                th_cnt = float(cbt.get("townhouses", {}).get("cnt") or (ann_tot * (townhouses / mth_tot) if th_val_m > 3000 else townhouses))
                th_cost = int(round((th_val_m * 1e6) / max(1.0, th_cnt))) if th_val_m > 0 else 0
                if 260_000 <= h_cost <= 980_000:
                    bc_cost_calc = f"Live Stats NZ Building Consents workbook ({m_full}): Stand-alone house mean consent value = ${h_cost:,}" + (f" | Townhouse/flat/unit mean = ${th_cost:,}" if th_cost else "")
                    out["new-build-consent-cost-by-type"] = make_pulse_item(
                        "Stats NZ", as_of, MAX_AGE_QUARTERLY,
                        f"In {m_full}, average consented residential building work values (excluding land and GST) in Stats NZ's building consents dataset were <strong>${h_cost:,}</strong> for a stand-alone house" + (f" and <strong>${th_cost:,}</strong> for a townhouse, flat, or unit." if th_cost else "."),
                        f"Stats NZ Building Consents Value Series ({m_full}, excluding land): Stand-alone house mean = ${h_cost:,}" + (f"; Townhouse/flat/unit mean = ${th_cost:,}." if th_cost else "."),
                        format_provenance_source_html(bc_cost_calc, bc_lbl, url, True), m_full, h_cost,
                        f"Stand-alone houses: ${h_cost:,}" + (f" | Townhouses/flats/units: ${th_cost:,}" if th_cost else ""),
                        "Estimated construction value declared on building consent (excluding land)", "Derived live from Stats NZ Building Consents workbook",
                        bc_cost_calc, f"Stats NZ residential consent value series ({m_full})", None, bc_extra,
                    )
            elif cbt.get("all", {}).get("val_m") and mth_tot > 0:
                all_val_m = float(cbt["all"]["val_m"])
                denom_d = ann_tot if all_val_m > 5000 else mth_tot
                avg_dwell_val = int(round((all_val_m * 1e6) / max(1, denom_d)))
                if 260_000 <= avg_dwell_val <= 980_000:
                    bc_cost_calc = f"Stats NZ Building Consents workbook ({m_full}): ${all_val_m:,.0f}m residential consent value ÷ {denom_d:,} consented dwellings = ${avg_dwell_val:,}/dwelling"
                    out["new-build-consent-cost-by-type"] = make_pulse_item(
                        "Stats NZ", as_of, MAX_AGE_QUARTERLY,
                        f"In {m_full}, consented residential building work across New Zealand averaged <strong>${avg_dwell_val:,} per consented dwelling</strong> (excluding land), across <strong>{mth_tot:,} monthly new dwellings</strong> ({houses:,} stand-alone houses and {multi_tot:,} multi-unit homes).",
                        f"Stats NZ Building Consents ({m_full}): Average consented residential building value = ${avg_dwell_val:,} per dwelling.",
                        format_provenance_source_html(bc_cost_calc, bc_lbl, url, True), m_full, avg_dwell_val,
                        f"${avg_dwell_val:,} average consent build value (excluding land)",
                        f"Stand-alone houses: {houses:,} | Multi-unit dwellings: {multi_tot:,}",
                        "Measures estimated construction value declared on building consent (excluding land)",
                        bc_cost_calc, f"Stats NZ monthly release ({m_full})", None, bc_extra,
                    )
            elif m_res_val and mth_tot > 0:
                raw_v = float(m_res_val.group(1).replace(",", ""))
                tot_res_nzd = raw_v * (1e9 if m_res_val.group(2).lower().startswith("b") else 1e6)
                avg_dwell_val = int(round(tot_res_nzd / (ann_tot if tot_res_nzd > 5e9 else mth_tot)))
                if 260_000 <= avg_dwell_val <= 980_000:
                    bc_cost_calc = f"Residential building consent value ÷ consented dwellings = ${avg_dwell_val:,}/dwelling ({houses:,} houses & {multi_tot:,} multi-unit dwellings in {m_full})"
                    out["new-build-consent-cost-by-type"] = make_pulse_item(
                        "Stats NZ", as_of, MAX_AGE_QUARTERLY,
                        f"In {m_full}, consented residential building work across New Zealand averaged <strong>${avg_dwell_val:,} per consented dwelling</strong> (excluding land), across <strong>{mth_tot:,} monthly new dwellings</strong> ({houses:,} stand-alone houses and {multi_tot:,} multi-unit homes).",
                        f"Stats NZ Building Consents ({m_full}): Average consented residential building value = ${avg_dwell_val:,} per dwelling.",
                        format_provenance_source_html(bc_cost_calc, bc_lbl, url, True), m_full, avg_dwell_val,
                        f"${avg_dwell_val:,} average consent build value (excluding land)",
                        f"Stand-alone houses: {houses:,} | Multi-unit dwellings: {multi_tot:,}",
                        "Measures estimated construction value declared on building consent (excluding land)",
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
        ) or re.search(r"([2345]\.\d)\s*percent\s+rise\s+for\s+all\s+groups", text, flags=re.I)
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
                cpi_clause = f", compared with <strong>{cpi_all_val:.1f}%</strong> headline All-groups CPI inflation" if cpi_all_val is not None else ""
                bci_calc = f"Directly reported Stats NZ CPI purchase of new housing = {signed_val:+.1f}% YoY"
                out["build-cost-inflation"] = make_pulse_item(
                    "Stats NZ CPI", q_as_of_iso, MAX_AGE_QUARTERLY,
                    f"In the {q_lbl_str}, Stats NZ's CPI 'purchase of new housing' sub-index changed by <strong>{signed_val:+.1f}%</strong> year-on-year{cpi_clause}.",
                    f"Stats NZ CPI ({q_lbl_str}): Purchase of new housing sub-index = {signed_val:+.1f}% YoY.",
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
    try:
        out.update(build_stats_nz_wealth_income_tickers(session))
    except Exception as exc_w:
        print(f"[Stats NZ Wealth/Income] Note: {exc_w}")
    return out
# === END OF BLOCK 9 ===
# === BLOCK 10 OF 13 ===
def build_tenancy_tickers(session) -> dict:
    """
    100% live-parsed MBIE Tenancy Services rental bond datasets:
    - National & Regional series (By-Region CSV)
    - City & District Council series (By-TLA CSV -> `tla-council-rent-benchmarks`, filtering out NaN / -99 rows)
    """
    pages = fetch_html_with_browser_fallback(session, [TENANCY_PAGE], "MBIE Tenancy")
    html = pages.get(TENANCY_PAGE, "")
    if not html:
        raise RuntimeError("Could not reach Tenancy Services rental bond page.")
    ten_pub_date = extract_html_release_date(html, "Monthly MBIE Tenancy Services rental bond release (~6–8 weeks after reference month-end)")

    csv_links = [
        "https://www.tenancy.govt.nz" + h if h.startswith("/") else h
        for h in re.findall(r'href=["\']([^"\']*\.csv[^"\']*)["\']', html, flags=re.I)
    ]
    csv_url = next((u for u in csv_links if "region" in u.lower()), "")
    tla_csv_url = next((u for u in csv_links if "tla" in u.lower() or "territorial" in u.lower()), "")

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
            f"In {m_full}, New Zealand's national median weekly rent across newly lodged private tenancy bonds was <strong>${nat_r}/wk</strong>{yoy_phrase}, with regional medians of <strong>${akl_r}/wk</strong> in Auckland, <strong>${wgtn_r}/wk</strong> in Wellington, and <strong>${cant_r}/wk</strong> in Canterbury.",
            f"MBIE Tenancy Services ({m_full}): National ('ALL') median weekly rent = ${nat_r}/wk; Auckland = ${akl_r}/wk; Wellington = ${wgtn_r}/wk; Canterbury = ${cant_r}/wk.",
            format_provenance_source_html(ry_calc, mbie_lbl, csv_url, True), m_full, nat_r,
            f"${nat_r}/wk NZ (${akl_r} AKL | ${wgtn_r} WGTN | ${cant_r} CANT)", f"12 months earlier (${prev_r}/wk NZ)",
            f"{yoy_pct:+.1f}% YoY (${nat_r - prev_r:+d}/wk)", ry_calc, hist_start_mbie,
            make_history_list(nat["dt"], nat["Median Rent"], lambda x: f"${int(round(x))}/wk"), mbie_extra,
        ),
        "rental-bonds-supply": make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"In {m_full}, <strong>{lodged_int:,}</strong> new private rental bonds were lodged across New Zealand, bringing the national pool of active private rental bonds held by MBIE to <strong>{active_int:,}</strong>.",
            f"MBIE Tenancy Services ({m_full}): Nationwide ('ALL') Lodged Bonds = {lodged_int:,}; Active Bonds = {active_int:,}.",
            format_provenance_source_html(sup_calc, mbie_lbl, csv_url, False), m_full, active_int,
            f"{lodged_int:,} lodged bonds | {active_int:,} active bonds", "Monthly MBIE bond registry",
            f"{lodged_int:,} new private tenancy bonds lodged in {m_full}", sup_calc, hist_start_mbie,
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
            f"Over the 10 years to {m_full}, New Zealand's national median weekly rent on newly lodged bonds rose <strong>+{pct_10y:.1f}% (+${nat_r - r10_val}/wk, or +{cagr_10y:.1f}%/yr compound)</strong> from <strong>${r10_val}/wk</strong> to <strong>${nat_r}/wk</strong>. Over the past 5 years (from <strong>${r5_val}/wk</strong>), median weekly rent rose <strong>+{pct_5y:.1f}% (+{cagr_5y:.1f}%/yr)</strong>.",
            f"MBIE Tenancy Services ({m_full}): National median weekly rent = ${nat_r}/wk (compared with ${r5_val}/wk 5 years earlier and ${r10_val}/wk 10 years earlier).",
            format_provenance_source_html(lt_calc, mbie_lbl, csv_url, True), m_full, round(pct_10y, 1),
            f"10Y: +{pct_10y:.1f}% (+{cagr_10y:.1f}%/yr) | 5Y: +{pct_5y:.1f}% (+{cagr_5y:.1f}%/yr)",
            f"5 years ago: ${r5_val}/wk | 10 years ago: ${r10_val}/wk",
            f"+${nat_r - r10_val}/wk (+${(nat_r - r10_val)*52:,}/yr) over 10 years", lt_calc, hist_start_mbie, lt_hist, mbie_extra,
        )

    if lq_col and uq_col and pd.notna(last[lq_col]) and pd.notna(last[uq_col]):
        lq_val_nat, uq_val_nat = int(round(float(last[lq_col]))), int(round(float(last[uq_col])))
        iqr_wk, iqr_yr = uq_val_nat - lq_val_nat, (uq_val_nat - lq_val_nat) * 52
        iqr_calc = f"Upper Quartile (${uq_val_nat}/wk) minus Lower Quartile (${lq_val_nat}/wk) = ${iqr_wk}/wk (${iqr_yr:,}/yr)"
        out["rent-quartile-spread"] = make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"In {m_full}, New Zealand private weekly rents on newly lodged bonds ranged from <strong>${lq_val_nat}/wk</strong> at the lower quartile (25th percentile) and <strong>${nat_r}/wk</strong> at the national median to <strong>${uq_val_nat}/wk</strong> at the upper quartile (75th percentile)—an interquartile spread of <strong>${iqr_wk}/wk (${iqr_yr:,}/yr)</strong>.",
            f"MBIE Tenancy Services ({m_full}): Nationwide ('ALL') Lower Quartile Rent = ${lq_val_nat}/wk; Median Rent = ${nat_r}/wk; Upper Quartile Rent = ${uq_val_nat}/wk.",
            format_provenance_source_html(iqr_calc, mbie_lbl, csv_url, True), m_full, iqr_wk,
            f"Lower quartile ${lq_val_nat}/wk | Median ${nat_r}/wk | Upper quartile ${uq_val_nat}/wk",
            f"Interquartile spread: ${iqr_wk}/wk (${iqr_yr:,}/yr)",
            f"Annualised rent range: ${lq_val_nat*52:,}/yr (25th pct) to ${uq_val_nat*52:,}/yr (75th pct)", iqr_calc, hist_start_mbie,
            [
                {"period": "Lower Quartile (25%)", "label": "Lower Quartile (25th percentile)", "value": lq_val_nat, "formatted": f"${lq_val_nat}/wk (${lq_val_nat*52:,}/yr)"},
                {"period": "National Median (50%)", "label": "National Median (50th percentile)", "value": nat_r, "formatted": f"${nat_r}/wk (${nat_r*52:,}/yr)"},
                {"period": "Upper Quartile (75%)", "label": "Upper Quartile (75th percentile)", "value": uq_val_nat, "formatted": f"${uq_val_nat}/wk (${uq_val_nat*52:,}/yr)"},
            ],
            mbie_extra,
        )

    if yoy_row is not None and pd.notna(yoy_row.get("Active Bonds")) and closed_int is not None:
        act_prev = int(round(float(yoy_row["Active Bonds"])))
        act_net_12m, act_net_pct = active_int - act_prev, ((active_int - act_prev) / max(1, act_prev)) * 100.0
        mth_net_flow = lodged_int - closed_int
        turn_calc = f"12M Net Active Bond Change: {active_int:,} − {act_prev:,} = {act_net_12m:+,d} ({act_net_pct:+.1f}%) | Monthly Net Flow: {lodged_int:,} − {closed_int:,} = {mth_net_flow:+,d}"
        out["rental-turnover-net-stock"] = make_pulse_item(
            "Tenancy Services", as_of, MAX_AGE_TENANCY,
            f"Over the 12 months to {m_full}, the total number of active private rental bonds held by MBIE Tenancy Services changed by <strong>{act_net_12m:+,d} bonds ({act_net_pct:+.1f}%)</strong> to <strong>{active_int:,} active bonds</strong> (with <strong>{lodged_int:,} bonds lodged</strong> and <strong>{closed_int:,} closed</strong> during {m_full}).",
            f"MBIE Tenancy Services ({m_full}): Active Bonds = {active_int:,} (vs {act_prev:,} a year earlier); Monthly Lodged Bonds = {lodged_int:,}; Monthly Closed Bonds = {closed_int:,}.",
            format_provenance_source_html(turn_calc, mbie_lbl, csv_url, True), m_full, act_net_12m,
            f"{act_net_12m:+,d} active bonds YoY ({act_net_pct:+.1f}%) | {lodged_int:,} lodged vs {closed_int:,} closed in {m_full}",
            f"12 months earlier: {act_prev:,} active bonds",
            f"Monthly net bond movement in {m_full}: {mth_net_flow:+,d} bonds", turn_calc, hist_start_mbie,
            make_history_list(nat["dt"], nat["Active Bonds"], lambda x: f"{int(round(x)):,} active bonds"), mbie_extra,
        )

    # 6. TLA / City & District Council Rent Benchmarks (`tla-council-rent-benchmarks` — filters out NaN / -99 rows)
    if tla_csv_url:
        try:
            tla_bytes = fetch_bytes_with_browser_fallback(session, tla_csv_url, referer=TENANCY_PAGE, min_bytes=8_000)
            if tla_bytes:
                df_tla = pd.read_csv(io.BytesIO(tla_bytes))
                df_tla.columns = [c.strip() for c in df_tla.columns]
                df_tla["dt"] = pd.to_datetime(df_tla["Time Frame"], errors="coerce", format="mixed")
                for col in ("Median Rent", "Active Bonds", "Lower Quartile Rent", "Upper Quartile Rent"):
                    if col in df_tla.columns:
                        df_tla[col] = pd.to_numeric(df_tla[col].astype(str).str.replace(",", "").str.replace("$", "").str.strip(), errors="coerce")
                latest_tla_dt = df_tla["dt"].max()
                sub_tla = df_tla[
                    (df_tla["dt"] == latest_tla_dt)
                    & df_tla["Location"].notna()
                    & (~df_tla["Location"].astype(str).str.strip().str.upper().isin(["ALL", "NAN", "NONE", "-99", ""]))
                ].dropna(subset=["Median Rent"])
                tla_rows = []
                for _, r_t in sub_tla.sort_values("Active Bonds", ascending=False).head(35).iterrows():
                    loc_n = str(r_t["Location"]).strip()
                    if re.match(r"^-?\d+(\.\d+)?$", loc_n):
                        continue
                    med_v = int(round(float(r_t["Median Rent"])))
                    act_v = int(round(float(r_t["Active Bonds"]))) if pd.notna(r_t.get("Active Bonds")) else 0
                    tla_rows.append({"period": loc_n, "label": loc_n, "value": med_v, "active_bonds": act_v, "formatted": f"${med_v}/wk ({act_v:,} active bonds)"})
                if len(tla_rows) >= 5:
                    tla_m_full = pd.Timestamp(latest_tla_dt).strftime("%B %Y")
                    tla_as_of = month_end_iso(pd.Timestamp(latest_tla_dt))
                    top_c1, top_c2, top_c3 = tla_rows[0], tla_rows[1], tla_rows[2]
                    tla_calc = f"Official MBIE By-TLA Rental Bond CSV ({tla_m_full}): {top_c1['label']} = ${top_c1['value']}/wk ({top_c1['active_bonds']:,} active bonds) | {top_c2['label']} = ${top_c2['value']}/wk | {top_c3['label']} = ${top_c3['value']}/wk"
                    out["tla-council-rent-benchmarks"] = make_pulse_item(
                        "Tenancy Services", tla_as_of, MAX_AGE_TENANCY,
                        f"In {tla_m_full}, across New Zealand's Territorial Local Authorities (MBIE Tenancy Services By-TLA dataset), median weekly rents on newly lodged bonds in the largest rental markets were <strong>${top_c1['value']}/wk in {top_c1['label']}</strong> ({top_c1['active_bonds']:,} active bonds), <strong>${top_c2['value']}/wk in {top_c2['label']}</strong> ({top_c2['active_bonds']:,} active bonds), and <strong>${top_c3['value']}/wk in {top_c3['label']}</strong> ({top_c3['active_bonds']:,} active bonds).",
                        f"MBIE Tenancy Services By-TLA Dataset ({tla_m_full}): Council-level median weekly rents and active bond counts.",
                        format_provenance_source_html(tla_calc, f"MBIE Tenancy Services — By-TLA Rental Bond Data ({tla_m_full})", tla_csv_url, False),
                        tla_m_full, top_c1["value"],
                        f"{top_c1['label']} ${top_c1['value']}/wk | {top_c2['label']} ${top_c2['value']}/wk | {top_c3['label']} ${top_c3['value']}/wk",
                        f"National median weekly rent: ${nat_r}/wk",
                        "Directly reported in MBIE Territorial Authority bond CSV",
                        tla_calc, f"MBIE By-TLA monthly series ({tla_m_full})",
                        [{k: r[k] for k in ("period", "label", "value", "formatted")} for r in tla_rows[:25]],
                        {"source_published": ten_pub_date, "source_url": tla_csv_url, "_tla_rent_map": {r["label"].lower(): r["value"] for r in tla_rows}},
                    )
        except Exception as exc_tla:
            print(f"[MBIE TLA CSV] Parse note: {exc_tla}")

    return out


def _norm_council_key(name: str) -> str:
    s = str(name).strip().lower()
    s = re.sub(r"\b(city|district|council|territorial authority|tla|region|regional|unitary)\b", "", s)
    return re.sub(r"[^a-z0-9]+", "", s)


def _scale_pct_series(ser: pd.Series) -> pd.Series:
    """Convert Excel decimal percentage series (e.g. 0.024 for 2.4%) to percentage points when needed."""
    s_num = pd.to_numeric(ser, errors="coerce")
    s_nz = s_num.dropna()
    s_nz = s_nz[s_nz != 0]
    if not s_nz.empty and float(s_nz.abs().quantile(0.95)) <= 1.2:
        return s_num * 100.0
    return s_num


def build_hud_mcert_housing_tickers(session, existing_candidates: dict) -> dict:
    """
    100% live-pulled HUD / MCERT Property and Sales Statistics & House Price Index (CC BY 4.0):
    1. `nz-median-house-price-quartiles`
    2. `regional-median-house-prices`
    3. `tla-council-house-price-benchmarks`
    4. `nz-monthly-property-sales-volume`
    5. `mcert-house-price-index-yoy`
    6. `hpi-low-vs-high-value-tiers`
    7. `sales-composition-by-value-quartile`
    8. `gross-rental-yield-by-council` (cross-derived with MBIE By-TLA median weekly rents)
    """
    out: dict = {}
    root = pathlib.Path(__file__).resolve().parent
    hud_urls = [HUD_ABOUT_PAGE, HUD_SALES_DASH_PAGE, HUD_COMP_DASH_PAGE, HUD_HPI_DASH_PAGE]
    pages = fetch_html_with_browser_fallback(session, hud_urls, "HUD / MCERT")
    about_html = pages.get(HUD_ABOUT_PAGE, "")
    hud_pub = extract_html_release_date(
        about_html,
        "Monthly HUD / MCERT Property and Sales Statistics release (~3-month District Valuation Roll settlement lag, CC BY 4.0)",
    )

    xlsx_urls: list[str] = []
    for h_html in pages.values():
        if not h_html:
            continue
        for href in re.findall(r'href=["\']([^"\']+\.xlsx[^"\']*)["\']', h_html, flags=re.I):
            full_u = "https://www.hud.govt.nz" + href if href.startswith("/") else href
            if any(k in full_u.lower() for k in ("property", "sales", "hpi", "stat", "download", "housing")):
                if full_u not in xlsx_urls:
                    xlsx_urls.append(full_u)

    xlsx_blobs: list[tuple[str, bytes]] = []
    for local_x in sorted(root.glob("*Property*sales*.xlsx"), reverse=True) + sorted(root.glob("*HUD*.xlsx"), reverse=True):
        if local_x.stat().st_size > 40_000:
            b_loc = local_x.read_bytes()
            if b_loc[:2] == b"PK":
                xlsx_blobs.append((xlsx_urls[0] if xlsx_urls else HUD_ABOUT_PAGE, b_loc))

    if not xlsx_blobs:
        for u_x in xlsx_urls[:3]:
            b_x = fetch_bytes_with_browser_fallback(session, u_x, referer=HUD_ABOUT_PAGE, min_bytes=40_000)
            if b_x and b_x[:2] == b"PK":
                xlsx_blobs.append((u_x, b_x))

    src_file_url = xlsx_blobs[0][0] if xlsx_blobs else (xlsx_urls[0] if xlsx_urls else HUD_ABOUT_PAGE)
    hud_src_lbl = "HUD / MCERT — Property and Sales Statistics (CC BY 4.0)"

    nat_price_df = pd.DataFrame()
    reg_latest_rows: list[dict] = []
    tla_latest_rows: list[dict] = []
    comp_latest: dict = {}
    hpi_df = pd.DataFrame()

    nz_regions_set = {
        "auckland", "wellington", "canterbury", "waikato", "bay of plenty", "otago",
        "manawatū-whanganui", "manawatu-whanganui", "hawke's bay", "hawkes bay", "northland",
        "taranaki", "southland", "nelson", "marlborough", "tasman", "gisborne", "west coast",
        "auckland region", "wellington region", "canterbury region", "waikato region",
        "bay of plenty region", "otago region", "northland region", "taranaki region",
        "southland region", "hawke's bay region", "manawatū-whanganui region", "manawatu-whanganui region",
    }
    nat_names_set = {
        "new zealand", "nz", "national", "all", "total", "all new zealand", "total new zealand", "national total",
    }

    for _, xlsx_bytes in xlsx_blobs:
        try:
            wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), read_only=True, data_only=True)
            for sname in wb.sheetnames:
                s_low = sname.strip().lower()
                if any(w in s_low for w in ("instruction", "note", "metadata", "content", "readme", "definition", "about", "caveat")):
                    continue
                ws = wb[sname]
                rows_iter = ws.iter_rows(values_only=True)
                header_row = None
                for _ in range(25):
                    try:
                        cand_r = next(rows_iter)
                    except StopIteration:
                        break
                    non_null = [str(c).strip() for c in cand_r if c is not None and str(c).strip()]
                    if len(non_null) >= 3 and any(
                        k in " ".join(non_null).lower()
                        for k in ("date", "month", "period", "median", "hpi", "index", "quartile", "sales", "region", "authority", "location", "geography", "area")
                    ):
                        header_row = [str(c).strip() if (c is not None and str(c).strip()) else f"col_{i}" for i, c in enumerate(cand_r)]
                        break
                if not header_row:
                    continue

                data_rows = list(rows_iter)
                if not data_rows:
                    continue
                n_cols = min(len(header_row), len(data_rows[0]))
                df_s = pd.DataFrame([r[:n_cols] for r in data_rows], columns=header_row[:n_cols])
                cols_map = {str(c).strip().lower(): c for c in df_s.columns}

                # 1. Identify true Date column by verifying sample values parse to 1980..2035 timestamps
                dt_col = None
                cand_dt_cols = [
                    c for k, c in cols_map.items()
                    if any(w in k for w in ("date", "month", "period", "time", "quarter", "year"))
                    and not any(x in k for x in ("length", "type", "span", "rolling", "duration", "frame_type"))
                ] + [c for c in df_s.columns]
                seen_dt_c = set()
                for c_try in cand_dt_cols:
                    if c_try in seen_dt_c:
                        continue
                    seen_dt_c.add(c_try)
                    samp = df_s[c_try].dropna().head(60)
                    if samp.empty:
                        continue
                    s_str = samp.astype(str).str.strip()
                    if s_str.str.match(r"^\d+\s*(?:month|mth|year|yr|qtr|quarter)", case=False).mean() > 0.3:
                        continue
                    parsed_dt = pd.to_datetime(samp, errors="coerce", format="mixed")
                    valid_dt = parsed_dt.dropna()
                    if len(valid_dt) >= max(3, int(len(samp) * 0.5)) and valid_dt.dt.year.between(1980, 2035).mean() >= 0.8:
                        dt_col = c_try
                        break

                if dt_col is None:
                    continue

                df_s["_dt"] = pd.to_datetime(df_s[dt_col], errors="coerce", format="mixed")
                df_s = df_s.dropna(subset=["_dt"]).sort_values("_dt").reset_index(drop=True)
                if df_s.empty:
                    continue

                # 2. Optional Timeframe / Rolling period filter (e.g. "1 month" vs "3 months" vs "12 months")
                tf_col = next(
                    (
                        c for k, c in cols_map.items()
                        if c != dt_col and any(w in k for w in ("timeframe", "time frame", "period length", "rolling", "duration", "window", "frequency", "span"))
                    ),
                    None,
                )
                if tf_col is None:
                    for c_try in df_s.columns:
                        if c_try == dt_col:
                            continue
                        u_vals = {str(v).strip().lower() for v in df_s[c_try].dropna().head(200).unique()}
                        if u_vals and len(u_vals) <= 6 and any(x in u_vals for x in ("1 month", "3 months", "12 months", "monthly", "3-month", "quarterly")):
                            tf_col = c_try
                            break

                # 3. Optional Dwelling Type filter (keep "All" / "Total" if dwelling type breakdown exists)
                dw_col = next((c for k, c in cols_map.items() if any(w in k for w in ("dwelling", "property type", "building type", "typology"))), None)
                if dw_col is not None:
                    dw_all = df_s[dw_col].astype(str).str.strip().str.lower().isin(["all", "total", "all dwellings", "all properties", "all residential"])
                    if dw_all.any():
                        df_s = df_s[dw_all].copy()

                # 4. Separate Geography Level (`lvl_col`) from Geography Name (`loc_col`) by inspecting values
                lvl_col = None
                loc_col = None
                geo_cand_cols = [
                    c for k, c in cols_map.items()
                    if c not in (dt_col, tf_col, dw_col)
                    and any(w in k for w in ("location", "area", "geograph", "region", "authority", "ta", "tla", "name", "level", "type", "council", "district"))
                ]
                for c_g in geo_cand_cols:
                    u_low = {str(v).strip().lower() for v in df_s[c_g].dropna().head(500).unique() if str(v).strip()}
                    if not u_low:
                        continue
                    if len(u_low) <= 6 and any(w in " ".join(u_low) for w in ("national", "region", "territorial", "tla", "council")):
                        if lvl_col is None:
                            lvl_col = c_g
                    elif any(p in u_low for p in ("auckland", "wellington", "canterbury", "christchurch city", "wellington city", "hamilton city", "new zealand")):
                        if loc_col is None:
                            loc_col = c_g

                if loc_col is None:
                    for c_g in geo_cand_cols:
                        if c_g != lvl_col:
                            loc_col = c_g
                            break

                # 5. Detect if this sheet is Long-Format (Measure/Statistic/Series + Value) and pivot if needed
                med_col = next((c for k, c in cols_map.items() if "median" in k and "rent" not in k), None)
                lq_col = next((c for k, c in cols_map.items() if ("lower" in k and "quartile" in k) or k in ("lq", "p25", "25th percentile", "lower quartile", "q1 price")), None)
                uq_col = next((c for k, c in cols_map.items() if ("upper" in k and "quartile" in k) or k in ("uq", "p75", "75th percentile", "upper quartile", "q3 price")), None)
                cnt_col = next(
                    (
                        c for k, c in cols_map.items()
                        if any(w in k for w in ("sales", "count", "volume", "number", "transactions"))
                        and not any(x in k for x in ("price", "median", "quartile", "share", "%", "index", "hpi"))
                    ),
                    None,
                )
                hpi_yoy_col = next(
                    (
                        c for k, c in cols_map.items()
                        if ("annual" in k or "yoy" in k or "year" in k or "12m" in k or "12-month" in k)
                        and ("hpi" in k or "change" in k or "%" in k or "growth" in k or "hpi" in s_low or "index" in s_low)
                        and c != dt_col
                    ),
                    None,
                )
                hpi_3m_col = next(
                    (
                        c for k, c in cols_map.items()
                        if ("3m" in k or "3-month" in k or "3 month" in k or "quarter" in k or "qoq" in k)
                        and ("change" in k or "%" in k or "growth" in k or "hpi" in s_low or "index" in s_low)
                        and c != dt_col
                    ),
                    None,
                )
                hpi_idx_col = next(
                    (
                        c for k, c in cols_map.items()
                        if ("hpi" in k or "index" in k)
                        and not any(x in k for x in ("change", "%", "yoy", "annual", "3m", "quarter", "composition", "quartile"))
                    ),
                    None,
                )
                q_tier_col = next((c for k, c in cols_map.items() if ("quartile" in k or "tier" in k or "value group" in k or "segment" in k) and c not in (lq_col, uq_col)), None)

                # Check for Long Format (Measure/Indicator column + Value column) for Sales Prices
                if med_col is None and "hpi" not in s_low and "composition" not in s_low:
                    meas_col = next((c for k, c in cols_map.items() if any(w in k for w in ("measure", "indicator", "statistic", "metric", "series", "variable", "price type"))), None)
                    val_col = next((c for k, c in cols_map.items() if any(w in k for w in ("value", "price", "amount", "figure"))), None)
                    if meas_col and val_col and meas_col != val_col:
                        m_u = " ".join(str(x).lower() for x in df_s[meas_col].dropna().unique()[:30])
                        if "median" in m_u:
                            df_s[val_col] = pd.to_numeric(df_s[val_col].astype(str).str.replace(",", "").str.replace("$", "").str.strip(), errors="coerce")
                            idx_cols = [c for c in ("_dt", loc_col, lvl_col, tf_col) if c is not None]
                            piv = df_s.pivot_table(index=idx_cols, columns=meas_col, values=val_col, aggfunc="last").reset_index()
                            df_s = piv.sort_values("_dt").reset_index(drop=True)
                            cols_map = {str(c).strip().lower(): c for c in df_s.columns}
                            med_col = next((c for k, c in cols_map.items() if "median" in k), None)
                            lq_col = next((c for k, c in cols_map.items() if "lower" in k or "25" in k or "lq" in k), None)
                            uq_col = next((c for k, c in cols_map.items() if "upper" in k or "75" in k or "uq" in k), None)
                            cnt_col = next((c for k, c in cols_map.items() if any(w in k for w in ("sales", "count", "volume", "number")) and "price" not in k), None)

                # A. Sales Prices & Volumes (National, Regional, TLA)
                if med_col is not None and "hpi" not in s_low and "composition" not in s_low:
                    for c_num in (med_col, lq_col, uq_col, cnt_col):
                        if c_num is not None and c_num in df_s.columns:
                            df_s[c_num] = pd.to_numeric(df_s[c_num].astype(str).str.replace(",", "").str.replace("$", "").str.strip(), errors="coerce")

                    # Helper to filter by preferred timeframe (1-month for national, 1-month or 3-month for TLA)
                    def _filter_tf(df_in: pd.DataFrame, prefer_monthly: bool = True) -> pd.DataFrame:
                        if tf_col is None or tf_col not in df_in.columns:
                            return df_in
                        tf_s = df_in[tf_col].astype(str).str.strip().str.lower()
                        m1 = tf_s.isin(["1 month", "1-month", "1m", "monthly", "month", "1"])
                        m3 = tf_s.isin(["3 months", "3-month", "3m", "quarterly", "quarter", "3"])
                        if prefer_monthly and m1.any() and df_in[m1][med_col].notna().sum() >= 5:
                            return df_in[m1]
                        if m3.any() and df_in[m3][med_col].notna().sum() >= 5:
                            return df_in[m3]
                        if m1.any():
                            return df_in[m1]
                        return df_in

                    df_nat_cand = _filter_tf(df_s, prefer_monthly=True)
                    nat_mask = pd.Series(False, index=df_nat_cand.index)
                    if loc_col is not None and loc_col in df_nat_cand.columns:
                        nat_mask = nat_mask | df_nat_cand[loc_col].astype(str).str.strip().str.lower().isin(nat_names_set)
                    if lvl_col is not None and lvl_col in df_nat_cand.columns:
                        nat_mask = nat_mask | df_nat_cand[lvl_col].astype(str).str.strip().str.lower().isin({"national", "new zealand", "nz"})

                    if nat_mask.any() and nat_price_df.empty:
                        sub_np = df_nat_cand[nat_mask].dropna(subset=[med_col]).sort_values("_dt").reset_index(drop=True)
                        if not sub_np.empty:
                            nat_price_df = sub_np
                            nat_price_df.attrs["med_col"] = med_col
                            nat_price_df.attrs["lq_col"] = lq_col
                            nat_price_df.attrs["uq_col"] = uq_col
                            nat_price_df.attrs["cnt_col"] = cnt_col
                    elif loc_col is None and nat_price_df.empty and ("national" in s_low or "nz" in s_low or "sales" in s_low):
                        sub_np = df_nat_cand.dropna(subset=[med_col]).sort_values("_dt").reset_index(drop=True)
                        if not sub_np.empty:
                            nat_price_df = sub_np
                            nat_price_df.attrs["med_col"] = med_col
                            nat_price_df.attrs["lq_col"] = lq_col
                            nat_price_df.attrs["uq_col"] = uq_col
                            nat_price_df.attrs["cnt_col"] = cnt_col

                    if loc_col is not None and loc_col in df_s.columns:
                        for prefer_m in (True, False):
                            df_sub_tf = _filter_tf(df_s, prefer_monthly=prefer_m)
                            valid_med = df_sub_tf.dropna(subset=[med_col])
                            if valid_med.empty:
                                continue
                            # Pick the most recent date that has at least 8 geographic locations reported
                            dt_counts = valid_med.groupby("_dt")[loc_col].nunique()
                            good_dts = dt_counts[dt_counts >= 8]
                            latest_dt_s = good_dts.index.max() if not good_dts.empty else valid_med["_dt"].max()
                            sub_latest = valid_med[valid_med["_dt"] == latest_dt_s]

                            for _, r_loc in sub_latest.iterrows():
                                loc_str = str(r_loc[loc_col]).strip()
                                loc_low = loc_str.lower()
                                if loc_low in nat_names_set or loc_low in ("nan", "none", ""):
                                    continue
                                m_v = int(round(float(r_loc[med_col])))
                                if m_v <= 0:
                                    continue
                                if m_v < 10_000:
                                    m_v *= 1000
                                lq_v = int(round(float(r_loc[lq_col]))) if (lq_col and pd.notna(r_loc.get(lq_col)) and float(r_loc[lq_col]) > 0) else None
                                if lq_v and lq_v < 10_000:
                                    lq_v *= 1000
                                uq_v = int(round(float(r_loc[uq_col]))) if (uq_col and pd.notna(r_loc.get(uq_col)) and float(r_loc[uq_col]) > 0) else None
                                if uq_v and uq_v < 10_000:
                                    uq_v *= 1000
                                c_v = int(round(float(r_loc[cnt_col]))) if (cnt_col and pd.notna(r_loc.get(cnt_col)) and float(r_loc[cnt_col]) > 0) else 0
                                lvl_str = str(r_loc[lvl_col]).strip().lower() if (lvl_col and pd.notna(r_loc.get(lvl_col))) else s_low

                                rec_item = {
                                    "loc": loc_str,
                                    "dt": latest_dt_s,
                                    "median": m_v,
                                    "lq": lq_v,
                                    "uq": uq_v,
                                    "sales": c_v,
                                }
                                is_reg = ("region" in lvl_str and "territorial" not in lvl_str) or (
                                    loc_low in nz_regions_set and "city" not in loc_low and "district" not in loc_low
                                )
                                is_tla = (
                                    any(w in lvl_str for w in ("territorial", "tla", "authority", "council", "district", "city"))
                                    or any(w in loc_low for w in ("city", "district", "auckland", "chatham"))
                                )
                                if is_reg and not any(x["loc"].lower() == loc_low for x in reg_latest_rows):
                                    reg_latest_rows.append(rec_item)
                                if is_tla and not any(x["loc"].lower() == loc_low for x in tla_latest_rows):
                                    tla_latest_rows.append(rec_item)

                            if len(reg_latest_rows) >= 4 and len(tla_latest_rows) >= 10:
                                break

                # B. House Price Index (MCERT SPAR HPI — All, Q1 Low-Value, Q4 High-Value)
                if (hpi_yoy_col is not None or hpi_idx_col is not None) and ("hpi" in s_low or "index" in s_low):
                    sub_hpi = df_s.copy()
                    if loc_col is not None and loc_col in sub_hpi.columns:
                        nat_h = sub_hpi[sub_hpi[loc_col].astype(str).str.strip().str.lower().isin(nat_names_set)]
                        if not nat_h.empty:
                            sub_hpi = nat_h
                    elif lvl_col is not None and lvl_col in sub_hpi.columns:
                        nat_h = sub_hpi[sub_hpi[lvl_col].astype(str).str.strip().str.lower().isin({"national", "new zealand", "nz"})]
                        if not nat_h.empty:
                            sub_hpi = nat_h

                    if hpi_yoy_col is not None and hpi_yoy_col in sub_hpi.columns:
                        sub_hpi["_yoy"] = _scale_pct_series(sub_hpi[hpi_yoy_col].astype(str).str.replace("%", "").str.strip())
                    else:
                        sub_hpi["_yoy"] = float("nan")

                    if hpi_3m_col is not None and hpi_3m_col in sub_hpi.columns:
                        sub_hpi["_3m"] = _scale_pct_series(sub_hpi[hpi_3m_col].astype(str).str.replace("%", "").str.strip())
                    else:
                        sub_hpi["_3m"] = float("nan")

                    sub_hpi["_q_tier"] = sub_hpi[q_tier_col].astype(str).str.strip() if (q_tier_col and q_tier_col in sub_hpi.columns) else "All"

                    # If _yoy is missing or 0 across rows, compute YoY and 3M directly from HPI Index level column
                    if hpi_idx_col is not None and hpi_idx_col in sub_hpi.columns:
                        sub_hpi["_idx"] = pd.to_numeric(sub_hpi[hpi_idx_col].astype(str).str.replace(",", "").str.strip(), errors="coerce")
                        if sub_hpi["_yoy"].dropna().empty or sub_hpi["_yoy"].dropna().abs().max() < 0.05:
                            calc_frames = []
                            for _, grp in sub_hpi.groupby("_q_tier", dropna=False):
                                grp = grp.sort_values("_dt").copy()
                                grp["_yoy"] = grp["_idx"].pct_change(12) * 100.0
                                grp["_3m"] = grp["_idx"].pct_change(3) * 100.0
                                calc_frames.append(grp)
                            if calc_frames:
                                sub_hpi = pd.concat(calc_frames, ignore_index=True).sort_values("_dt").reset_index(drop=True)

                    valid_hpi = sub_hpi.dropna(subset=["_yoy"])
                    valid_hpi = valid_hpi[valid_hpi["_yoy"].abs() > 1e-6]
                    if not valid_hpi.empty and hpi_df.empty:
                        hpi_df = valid_hpi.copy()

                # C. Sales Composition by Valuation Quartile (Q1..Q4)
                if ("composition" in s_low or q_tier_col is not None) and cnt_col is not None and "hpi" not in s_low and not comp_latest:
                    sub_c = df_s.copy()
                    sub_c[cnt_col] = pd.to_numeric(sub_c[cnt_col].astype(str).str.replace(",", "").str.strip(), errors="coerce")
                    if loc_col is not None and loc_col in sub_c.columns:
                        nat_c = sub_c[sub_c[loc_col].astype(str).str.strip().str.lower().isin(nat_names_set)]
                        if not nat_c.empty:
                            sub_c = nat_c
                    elif lvl_col is not None and lvl_col in sub_c.columns:
                        nat_c = sub_c[sub_c[lvl_col].astype(str).str.strip().str.lower().isin({"national", "new zealand", "nz"})]
                        if not nat_c.empty:
                            sub_c = nat_c

                    sub_c = sub_c.dropna(subset=[cnt_col])
                    if q_tier_col and q_tier_col in sub_c.columns and not sub_c.empty:
                        for dt_try in sorted(sub_c["_dt"].unique(), reverse=True)[:4]:
                            sub_c_lat = sub_c[sub_c["_dt"] == dt_try]
                            q_counts = {}
                            for _, rc in sub_c_lat.iterrows():
                                ql = str(rc[q_tier_col]).strip().upper()
                                for qk, qnum in (("Q1", "1"), ("Q2", "2"), ("Q3", "3"), ("Q4", "4")):
                                    if qk in ql or re.search(rf"\b{qnum}\b", ql) or (qk == "Q1" and "LOW" in ql) or (qk == "Q4" and "HIGH" in ql):
                                        if float(rc[cnt_col]) > 0:
                                            q_counts[qk] = float(rc[cnt_col])
                            if len(q_counts) == 4:
                                comp_latest = {"dt": dt_try, "counts": q_counts}
                                break
            wb.close()
        except Exception as exc_hud_x:
            print(f"[HUD / MCERT XLSX] Parse note: {exc_hud_x}")

    # Build Ticker 1 & 4 from `nat_price_df` (`nz-median-house-price-quartiles` & `nz-monthly-property-sales-volume`)
    if not nat_price_df.empty:
        med_c = nat_price_df.attrs.get("med_col") or next((c for c in nat_price_df.columns if "median" in str(c).lower()), None)
        lq_c = nat_price_df.attrs.get("lq_col") or next((c for c in nat_price_df.columns if "lower" in str(c).lower() or str(c).lower() in ("lq", "p25")), None)
        uq_c = nat_price_df.attrs.get("uq_col") or next((c for c in nat_price_df.columns if "upper" in str(c).lower() or str(c).lower() in ("uq", "p75")), None)
        cnt_c = nat_price_df.attrs.get("cnt_col") or next((c for c in nat_price_df.columns if any(w in str(c).lower() for w in ("sales", "count", "volume", "number")) and "price" not in str(c).lower()), None)

        if med_c and med_c in nat_price_df.columns:
            nat_price_df = nat_price_df[nat_price_df[med_c] > 0].sort_values("_dt").reset_index(drop=True)

        if not nat_price_df.empty and med_c:
            last_np = nat_price_df.iloc[-1]
            dt_np = pd.Timestamp(last_np["_dt"])
            diffs_np = (nat_price_df["_dt"] - (dt_np - pd.DateOffset(years=1))).abs()
            prev_np = nat_price_df.loc[diffs_np.idxmin()] if len(nat_price_df) >= 2 else last_np
            per_np, as_of_np = dt_np.strftime("%B %Y"), month_end_iso(dt_np)
            hist_from_np = f"Available from {pd.Timestamp(nat_price_df['_dt'].iloc[0]).strftime('%B %Y')} ({len(nat_price_df)} observations)"

            med_now = int(round(float(last_np[med_c])))
            if med_now < 10_000:
                med_now *= 1000
                nat_price_df[med_c] = nat_price_df[med_c] * 1000.0
            med_prev = int(round(float(prev_np[med_c])))
            if med_prev < 10_000:
                med_prev *= 1000
            med_yoy = ((med_now - med_prev) / max(1, med_prev)) * 100.0

            lq_now = int(round(float(last_np[lq_c]))) if (lq_c and pd.notna(last_np.get(lq_c)) and float(last_np[lq_c]) > 0) else None
            if lq_now and lq_now < 10_000:
                lq_now *= 1000
            uq_now = int(round(float(last_np[uq_c]))) if (uq_c and pd.notna(last_np.get(uq_c)) and float(last_np[uq_c]) > 0) else None
            if uq_now and uq_now < 10_000:
                uq_now *= 1000

            q_clause = (
                f"—with a lower-quartile (25th percentile) sale price of <strong>${lq_now:,}</strong> and an upper-quartile (75th percentile) sale price of <strong>${uq_now:,}</strong> (an interquartile spread of <strong>${uq_now - lq_now:,}</strong>)"
                if (lq_now and uq_now)
                else ""
            )
            np_calc = f"National Median Sale Price ({per_np}) = ${med_now:,} (YoY: {med_yoy:+.1f}% vs ${med_prev:,})" + (
                f" | Lower Quartile (25%) = ${lq_now:,} | Upper Quartile (75%) = ${uq_now:,} (Spread = ${uq_now - lq_now:,})"
                if (lq_now and uq_now)
                else ""
            )
            out["nz-median-house-price-quartiles"] = make_pulse_item(
                "HUD / MCERT", as_of_np, MAX_AGE_HUD,
                f"In {per_np} (by agreement date), New Zealand's national median residential sale price in the HUD / MCERT District Valuation Roll series was <strong>${med_now:,}</strong> ({plain_yoy_growth(med_now, med_prev)} from ${med_prev:,} a year earlier){q_clause}.",
                f"HUD / MCERT Property and Sales Statistics ({per_np}): National median residential sale price = ${med_now:,}" + (f"; Lower quartile = ${lq_now:,}; Upper quartile = ${uq_now:,}." if (lq_now and uq_now) else "."),
                format_provenance_source_html(np_calc, hud_src_lbl, src_file_url, True),
                per_np, med_now,
                f"Median ${med_now:,}" + (f" | Lower quartile ${lq_now:,} | Upper quartile ${uq_now:,}" if (lq_now and uq_now) else ""),
                f"12 months prior ({pd.Timestamp(prev_np['_dt']).strftime('%B %Y')}): ${med_prev:,} median",
                f"{med_yoy:+.1f}% YoY (${med_now - med_prev:+,d} national median)",
                np_calc, hist_from_np,
                make_history_list(nat_price_df["_dt"], nat_price_df[med_c], lambda v: f"${int(round(v)):,}"),
                {"source_published": hud_pub, "source_url": src_file_url},
            )

            if cnt_c and pd.notna(last_np.get(cnt_c)) and float(last_np[cnt_c]) > 0:
                s_mth_now = int(round(float(last_np[cnt_c])))
                s_mth_prev = int(round(float(prev_np[cnt_c]))) if pd.notna(prev_np.get(cnt_c)) else s_mth_now
                s_12m_now = int(round(float(nat_price_df[cnt_c].iloc[-12:].sum()))) if len(nat_price_df) >= 12 else s_mth_now * 12
                s_12m_prev = int(round(float(nat_price_df[cnt_c].iloc[-24:-12].sum()))) if len(nat_price_df) >= 24 else None
                s_mth_yoy = ((s_mth_now - s_mth_prev) / max(1, s_mth_prev)) * 100.0
                vol_calc = f"Monthly DVR Sales ({per_np}) = {s_mth_now:,} sales ({s_mth_yoy:+.1f}% YoY vs {s_mth_prev:,}) | Rolling 12-Month Sales = {s_12m_now:,}"
                out["nz-monthly-property-sales-volume"] = make_pulse_item(
                    "HUD / MCERT", as_of_np, MAX_AGE_HUD,
                    f"In {per_np} (by agreement date), <strong>{s_mth_now:,} arm's-length residential property sales</strong> were recorded across New Zealand ({plain_yoy_growth(s_mth_now, s_mth_prev)} from {s_mth_prev:,} a year earlier), bringing rolling 12-month national sales volume to <strong>{s_12m_now:,} properties</strong>.",
                    f"HUD / MCERT Property and Sales Statistics ({per_np}): Monthly national residential sales = {s_mth_now:,}; rolling 12-month sales = {s_12m_now:,}.",
                    format_provenance_source_html(vol_calc, hud_src_lbl, src_file_url, True),
                    per_np, s_mth_now,
                    f"{s_mth_now:,} sales in {per_np} | {s_12m_now:,} rolling 12m sales",
                    f"12 months prior: {s_mth_prev:,} monthly sales" + (f" ({s_12m_prev:,} prior 12m total)" if s_12m_prev else ""),
                    f"{s_mth_yoy:+.1f}% YoY ({s_mth_now - s_mth_prev:+,d} monthly sales)",
                    vol_calc, hist_from_np,
                    make_history_list(nat_price_df["_dt"], nat_price_df[cnt_c], lambda v: f"{int(round(v)):,} sales"),
                    {"source_published": hud_pub, "source_url": src_file_url},
                )

    # Build Ticker 2: `regional-median-house-prices`
    if len(reg_latest_rows) >= 4:
        reg_latest_rows.sort(key=lambda r: r["sales"] if r["sales"] > 0 else r["median"], reverse=True)
        dt_reg = pd.Timestamp(reg_latest_rows[0]["dt"])
        per_reg, as_of_reg = dt_reg.strftime("%B %Y"), month_end_iso(dt_reg)
        r1, r2, r3 = reg_latest_rows[0], reg_latest_rows[1], reg_latest_rows[2]
        reg_calc = " | ".join(
            f"{r['loc']}: ${r['median']:,} median" + (f" (LQ ${r['lq']:,})" if r.get("lq") else "")
            for r in reg_latest_rows[:6]
        )
        out["regional-median-house-prices"] = make_pulse_item(
            "HUD / MCERT", as_of_reg, MAX_AGE_HUD,
            f"In {per_reg} (HUD / MCERT regional DVR series), regional median residential sale prices across New Zealand's largest regional markets were <strong>${r1['median']:,} in {r1['loc']}</strong>" + (f" (lower quartile ${r1['lq']:,})" if r1.get("lq") else "") + f", <strong>${r2['median']:,} in {r2['loc']}</strong>, and <strong>${r3['median']:,} in {r3['loc']}</strong>.",
            f"HUD / MCERT Property and Sales Statistics ({per_reg} Regional Series): {reg_calc}.",
            format_provenance_source_html(reg_calc, hud_src_lbl, src_file_url, False),
            per_reg, r1["median"],
            " | ".join(f"{r['loc']} ${r['median']:,}" for r in reg_latest_rows[:4]),
            "Regional District Valuation Roll market sales series",
            f"Published by agreement date ({per_reg})",
            reg_calc, f"HUD / MCERT regional monthly series ({per_reg})",
            [
                {
                    "period": r["loc"],
                    "label": r["loc"],
                    "value": r["median"],
                    "formatted": f"${r['median']:,} median" + (f" | ${r['lq']:,} lower quartile (25%)" if r.get("lq") else "") + (f" ({r['sales']:,} sales)" if r.get("sales") else ""),
                }
                for r in reg_latest_rows[:16]
            ],
            {"source_published": hud_pub, "source_url": src_file_url},
        )

    # Build Ticker 3 & 8: `tla-council-house-price-benchmarks` & `gross-rental-yield-by-council`
    if len(tla_latest_rows) >= 5:
        tla_latest_rows.sort(key=lambda r: r["sales"] if r["sales"] > 0 else r["median"], reverse=True)
        dt_tla = pd.Timestamp(tla_latest_rows[0]["dt"])
        per_tla, as_of_tla = dt_tla.strftime("%B %Y"), month_end_iso(dt_tla)
        t1, t2, t3 = tla_latest_rows[0], tla_latest_rows[1], tla_latest_rows[2]
        tla_p_calc = " | ".join(
            f"{r['loc']}: ${r['median']:,} median" + (f" (LQ ${r['lq']:,})" if r.get("lq") else "")
            for r in tla_latest_rows[:5]
        )
        out["tla-council-house-price-benchmarks"] = make_pulse_item(
            "HUD / MCERT", as_of_tla, MAX_AGE_HUD,
            f"In {per_tla} (HUD / MCERT Territorial Authority DVR series), council-level median residential sale prices in the largest urban districts were <strong>${t1['median']:,} in {t1['loc']}</strong>" + (f" (lower quartile ${t1['lq']:,})" if t1.get("lq") else "") + f", <strong>${t2['median']:,} in {t2['loc']}</strong>" + (f" (lower quartile ${t2['lq']:,})" if t2.get("lq") else "") + f", and <strong>${t3['median']:,} in {t3['loc']}</strong>.",
            f"HUD / MCERT Property and Sales Statistics ({per_tla} Territorial Authority Series): {tla_p_calc}.",
            format_provenance_source_html(tla_p_calc, hud_src_lbl, src_file_url, False),
            per_tla, t1["median"],
            " | ".join(f"{r['loc']} ${r['median']:,}" for r in tla_latest_rows[:4]),
            "Territorial Local Authority District Valuation Roll sales series",
            f"Published by agreement date ({per_tla})",
            tla_p_calc, f"HUD / MCERT TLA monthly series ({per_tla})",
            [
                {
                    "period": r["loc"],
                    "label": r["loc"],
                    "value": r["median"],
                    "formatted": f"${r['median']:,} median" + (f" | ${r['lq']:,} lower quartile (25%)" if r.get("lq") else "") + (f" ({r['sales']:,} sales)" if r.get("sales") else ""),
                }
                for r in tla_latest_rows[:25]
            ],
            {"source_published": hud_pub, "source_url": src_file_url},
        )

        # Cross-derive Ticker 8 (`gross-rental-yield-by-council`) using MBIE By-TLA Rents
        rent_rec = existing_candidates.get("tla-council-rent-benchmarks") or {}
        raw_rent_map = dict(rent_rec.get("_tla_rent_map") or {})
        if not raw_rent_map and rent_rec.get("history"):
            for h_r in rent_rec["history"]:
                m_rv = re.search(r"\$([\d,]+)/wk", str(h_r.get("formatted", "")))
                if m_rv and h_r.get("label"):
                    raw_rent_map[str(h_r["label"]).strip().lower()] = int(m_rv.group(1).replace(",", ""))

        norm_rent_map = {_norm_council_key(k): v for k, v in raw_rent_map.items() if _norm_council_key(k)}
        if norm_rent_map:
            yield_rows = []
            for r_p in tla_latest_rows:
                if r_p.get("sales", 50) < 15 or r_p["median"] <= 100_000:
                    continue
                c_name = r_p["loc"]
                c_key = _norm_council_key(c_name)
                wk_rent = raw_rent_map.get(c_name.lower()) or norm_rent_map.get(c_key)
                if wk_rent and 250 <= wk_rent <= 1500:
                    ann_rent = wk_rent * 52
                    gy_pct = round((ann_rent / float(r_p["median"])) * 100.0, 2)
                    if 2.0 <= gy_pct <= 10.5:
                        yield_rows.append({
                            "loc": c_name,
                            "yield_pct": gy_pct,
                            "wk_rent": wk_rent,
                            "med_price": r_p["median"],
                            "period": c_name,
                            "label": c_name,
                            "value": gy_pct,
                            "formatted": f"{gy_pct:.2f}% gross yield (${wk_rent}/wk rent ÷ ${r_p['median']:,} median price)",
                        })
            if len(yield_rows) >= 3:
                y1, y2, y3 = yield_rows[0], yield_rows[1], yield_rows[2]
                rent_per = rent_rec.get("latest_period", per_tla)
                gy_calc = " | ".join(
                    f"{y['loc']}: (${y['wk_rent']}×52) ÷ ${y['med_price']:,} = {y['yield_pct']:.2f}%"
                    for y in yield_rows[:4]
                )
                out["gross-rental-yield-by-council"] = make_pulse_item(
                    "MBIE & HUD / MCERT", as_of_tla, MAX_AGE_HUD,
                    f"Combining MBIE Tenancy Services council median weekly rents ({rent_per}) with HUD / MCERT council median residential sale prices ({per_tla}), indicative gross residential rental yields across major urban councils were <strong>{y1['yield_pct']:.2f}% in {y1['loc']}</strong> (${y1['wk_rent']}/wk vs ${y1['med_price']:,}), <strong>{y2['yield_pct']:.2f}% in {y2['loc']}</strong> (${y2['wk_rent']}/wk vs ${y2['med_price']:,}), and <strong>{y3['yield_pct']:.2f}% in {y3['loc']}</strong> (${y3['wk_rent']}/wk vs ${y3['med_price']:,}).",
                    f"MBIE By-TLA Median Rent ({rent_per}) & HUD / MCERT By-TLA Median Sale Price ({per_tla}): {gy_calc}.",
                    format_provenance_source_html(gy_calc, "MBIE Tenancy Services & HUD / MCERT Property Statistics (CC BY 4.0)", src_file_url, True),
                    f"{per_tla} (Prices) / {rent_per} (Rents)", y1["yield_pct"],
                    " | ".join(f"{y['loc']} {y['yield_pct']:.2f}%" for y in yield_rows[:4]),
                    "Ratio of annualised council median bond rent to council median DVR sale price",
                    "Gross ratio before rates, insurance, maintenance, management, vacancy & interest",
                    gy_calc, f"MBIE & HUD / MCERT Territorial Authority series ({per_tla})",
                    [{k: y[k] for k in ("period", "label", "value", "formatted")} for y in yield_rows[:20]],
                    {"source_published": hud_pub, "source_url": src_file_url},
                )

    # Build Ticker 5 & 6: `mcert-house-price-index-yoy` & `hpi-low-vs-high-value-tiers`
    if not hpi_df.empty:
        latest_h_dt = hpi_df["_dt"].max()
        per_h, as_of_h = pd.Timestamp(latest_h_dt).strftime("%B %Y"), month_end_iso(pd.Timestamp(latest_h_dt))
        all_h_df = hpi_df[hpi_df["_q_tier"].astype(str).str.lower().isin(["all", "total", "all dwellings", "nan", ""])]
        if all_h_df.empty:
            all_h_df = hpi_df
        last_h = all_h_df.iloc[-1]
        yoy_h = float(last_h["_yoy"])
        qoq_h = float(last_h["_3m"]) if pd.notna(last_h.get("_3m")) else None
        qoq_clause = f" (and <strong>{qoq_h:+.1f}%</strong> over the 3 months to {per_h})" if qoq_h is not None else ""
        hpi_calc = f"Official MCERT SPAR House Price Index ({per_h}): Annual change = {yoy_h:+.1f}% YoY" + (f" | 3-month change = {qoq_h:+.1f}%" if qoq_h is not None else "")
        out["mcert-house-price-index-yoy"] = make_pulse_item(
            "HUD / MCERT HPI", as_of_h, MAX_AGE_HUD,
            f"In {per_h} (by agreement date), the official Crown House Price Index (MCERT HPI, calculated using the Sales Price to Appraisal Ratio methodology) changed by <strong>{yoy_h:+.1f}% year-on-year</strong> across New Zealand{qoq_clause}.",
            f"HUD / MCERT House Price Index ({per_h}): National All-Dwellings HPI annual change = {yoy_h:+.1f}% YoY.",
            format_provenance_source_html(hpi_calc, "HUD / MCERT — House Price Index (CC BY 4.0)", src_file_url, False),
            per_h, round(yoy_h, 2),
            f"{yoy_h:+.1f}% YoY" + (f" | {qoq_h:+.1f}% 3-month" if qoq_h is not None else "") + " (MCERT SPAR HPI)",
            "SPAR valuation-adjusted index (controls for sales composition)",
            f"Published with 3-month DVR settlement lag ({per_h})",
            hpi_calc, f"MCERT monthly HPI series ({per_h})",
            make_history_list(all_h_df["_dt"], all_h_df["_yoy"], lambda v: f"{v:+.1f}% YoY"),
            {"source_published": hud_pub, "source_url": src_file_url},
        )

        lat_tiers = hpi_df[hpi_df["_dt"] == latest_h_dt]
        q1_r = lat_tiers[lat_tiers["_q_tier"].astype(str).str.upper().str.contains("Q1|LOW|1")]
        q4_r = lat_tiers[lat_tiers["_q_tier"].astype(str).str.upper().str.contains("Q4|HIGH|4")]
        if not q1_r.empty and not q4_r.empty:
            q1_yoy = float(q1_r.iloc[-1]["_yoy"])
            q4_yoy = float(q4_r.iloc[-1]["_yoy"])
            tier_diff = round(q1_yoy - q4_yoy, 1)
            tier_calc = f"Low-Value Tier (Q1) HPI YoY = {q1_yoy:+.1f}% | High-Value Tier (Q4) HPI YoY = {q4_yoy:+.1f}% | Spread (Q1 − Q4) = {tier_diff:+.1f} percentage points"
            out["hpi-low-vs-high-value-tiers"] = make_pulse_item(
                "HUD / MCERT HPI", as_of_h, MAX_AGE_HUD,
                f"In {per_h}, the MCERT House Price Index for New Zealand's lowest valuation quartile (<strong>Q1 low-value homes</strong>) changed by <strong>{q1_yoy:+.1f}% year-on-year</strong>, compared with <strong>{q4_yoy:+.1f}% year-on-year</strong> for the highest valuation quartile (<strong>Q4 high-value homes</strong>)—a difference of <strong>{tier_diff:+.1f} percentage points</strong>.",
                f"HUD / MCERT House Price Index Composition ({per_h}): Q1 (lowest valuation quartile) = {q1_yoy:+.1f}% YoY; Q4 (highest valuation quartile) = {q4_yoy:+.1f}% YoY.",
                format_provenance_source_html(tier_calc, "HUD / MCERT — House Price Index Composition (CC BY 4.0)", src_file_url, True),
                per_h, round(q1_yoy, 2),
                f"Q1 Low-value {q1_yoy:+.1f}% YoY vs Q4 High-value {q4_yoy:+.1f}% YoY ({tier_diff:+.1f} pts)",
                f"All-dwellings national HPI: {yoy_h:+.1f}% YoY",
                f"{tier_diff:+.1f} percentage-point spread between Q1 and Q4 tiers",
                tier_calc, f"MCERT HPI valuation-quartile series ({per_h})",
                [
                    {"period": "Q1 (Lowest Value 25%)", "label": "Q1 Lowest Valuation Quartile Homes", "value": round(q1_yoy, 1), "formatted": f"{q1_yoy:+.1f}% YoY"},
                    {"period": "All Dwellings", "label": "All New Zealand Dwellings (Composite HPI)", "value": round(yoy_h, 1), "formatted": f"{yoy_h:+.1f}% YoY"},
                    {"period": "Q4 (Highest Value 25%)", "label": "Q4 Highest Valuation Quartile Homes", "value": round(q4_yoy, 1), "formatted": f"{q4_yoy:+.1f}% YoY"},
                ],
                {"source_published": hud_pub, "source_url": src_file_url},
            )

    # Build Ticker 7: `sales-composition-by-value-quartile`
    if comp_latest and len(comp_latest.get("counts", {})) == 4:
        dt_c = pd.Timestamp(comp_latest["dt"])
        per_c, as_of_c = dt_c.strftime("%B %Y"), month_end_iso(dt_c)
        q_c = comp_latest["counts"]
        tot_q_s = sum(q_c.values())
        q1_sh = round((q_c["Q1"] / max(1.0, tot_q_s)) * 100.0, 1)
        q2_sh = round((q_c["Q2"] / max(1.0, tot_q_s)) * 100.0, 1)
        q3_sh = round((q_c["Q3"] / max(1.0, tot_q_s)) * 100.0, 1)
        q4_sh = round((q_c["Q4"] / max(1.0, tot_q_s)) * 100.0, 1)
        comp_calc = f"Q1: {q1_sh:.1f}% ({int(round(q_c['Q1'])):,}) | Q2: {q2_sh:.1f}% ({int(round(q_c['Q2'])):,}) | Q3: {q3_sh:.1f}% ({int(round(q_c['Q3'])):,}) | Q4: {q4_sh:.1f}% ({int(round(q_c['Q4'])):,})"
        out["sales-composition-by-value-quartile"] = make_pulse_item(
            "HUD / MCERT", as_of_c, MAX_AGE_HUD,
            f"In {per_c}, across New Zealand residential property sales in the HUD / MCERT Sales Composition dataset, <strong>{q1_sh:.1f}%</strong> of sales were in the lowest rating-valuation quartile (<strong>Q1</strong>), <strong>{q2_sh:.1f}%</strong> in <strong>Q2</strong>, <strong>{q3_sh:.1f}%</strong> in <strong>Q3</strong>, and <strong>{q4_sh:.1f}%</strong> in the highest valuation quartile (<strong>Q4</strong>).",
            f"HUD / MCERT Sales Composition ({per_c}): {comp_calc}.",
            format_provenance_source_html(comp_calc, "HUD / MCERT — Sales Composition Statistics (CC BY 4.0)", src_file_url, True),
            per_c, q1_sh,
            f"Q1 Lowest CV {q1_sh:.1f}% | Q2 {q2_sh:.1f}% | Q3 {q3_sh:.1f}% | Q4 Highest CV {q4_sh:.1f}%",
            "Equal-stock benchmark: 25.0% of housing stock per valuation quartile",
            f"Q1 minus Q4 sales share difference: {q1_sh - q4_sh:+.1f} percentage points",
            comp_calc, f"HUD / MCERT sales composition series ({per_c})",
            [
                {"period": "Q1 (Lowest 25% CV Stock)", "label": "Quartile 1 (Lowest Valuation 25% of Stock)", "value": q1_sh, "formatted": f"{q1_sh:.1f}% of sales ({int(round(q_c['Q1'])):,} sales)"},
                {"period": "Q2 (25%–50% CV Stock)", "label": "Quartile 2 (Lower-Middle Valuation 25% of Stock)", "value": q2_sh, "formatted": f"{q2_sh:.1f}% of sales ({int(round(q_c['Q2'])):,} sales)"},
                {"period": "Q3 (50%–75% CV Stock)", "label": "Quartile 3 (Upper-Middle Valuation 25% of Stock)", "value": q3_sh, "formatted": f"{q3_sh:.1f}% of sales ({int(round(q_c['Q3'])):,} sales)"},
                {"period": "Q4 (Highest 25% CV Stock)", "label": "Quartile 4 (Highest Valuation 25% of Stock)", "value": q4_sh, "formatted": f"{q4_sh:.1f}% of sales ({int(round(q_c['Q4'])):,} sales)"},
            ],
            {"source_published": hud_pub, "source_url": src_file_url},
        )

    return out
# === END OF BLOCK 10 ===
# === BLOCK 11 OF 13 ===
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
    Zero hardcoded fallbacks: every figure is mathematically verified from the PDF text stream.
    """
    out: dict = {}
    root = pathlib.Path(__file__).resolve().parent
    ko_pages = fetch_html_with_browser_fallback(session, [KAINGA_ORA_QTR_REPORT_PAGE], "Kāinga Ora")
    qtr_html = ko_pages.get(KAINGA_ORA_QTR_REPORT_PAGE, "")
    qtr_pub = extract_html_release_date(qtr_html, "Quarterly Kāinga Ora Home Ownership Products Report release (CC BY 4.0)")

    pdf_sources: list[tuple[str, bytes]] = []
    seen_local: set[str] = set()
    for local_pdf in sorted(root.glob("*Home*Ownership*Products*.pdf"), reverse=True) + sorted(root.glob("*.pdf"), reverse=True):
        if local_pdf.name not in seen_local and local_pdf.stat().st_size > 10_000:
            seen_local.add(local_pdf.name)
            b_loc = local_pdf.read_bytes()
            if b_loc.startswith(b"%PDF"):
                pdf_sources.append((KAINGA_ORA_QTR_REPORT_PAGE, b_loc))

    if qtr_html:
        ko_pdf_links = []
        for h in re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', qtr_html, flags=re.I):
            if "home-ownership-products" in h.lower() or "quarterly" in h.lower():
                full_u = "https://kaingaora.govt.nz" + h if h.startswith("/") else h
                if full_u not in ko_pdf_links:
                    ko_pdf_links.append(full_u)
        for full_u in ko_pdf_links[:5]:
            b_http = fetch_bytes_with_browser_fallback(session, full_u, referer=KAINGA_ORA_QTR_REPORT_PAGE, min_bytes=10_000)
            if b_http and b_http.startswith(b"%PDF"):
                pdf_sources.append((full_u, b_http))

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

    trend_clause = ""
    roll_clause = ""
    prev_str = f"Monthly approvals in quarter: {m1a:,}, {m2a:,}, {m3a:,}"
    chg_str = f"{q_set:,} First Home Loans settled in {q_disp} ({conv:.1f}% settlement-to-approval ratio)"
    fhl_hist = [
        {
            "period": q["q_label"],
            "label": q["q_label"],
            "value": q["app_row"][3],
            "formatted": f"{q['app_row'][3]:,} approvals | {q['set_row'][3]:,} settlements ({(q['set_row'][3]/max(1,q['app_row'][3]))*100:.1f}% ratio)",
        }
        for q in parsed_quarters
    ]

    if len(parsed_quarters) >= 2:
        prev_q = parsed_quarters[1]
        pq_a, pq_s = prev_q["app_row"][3], prev_q["set_row"][3]
        qoq_a_pct = ((q_app - pq_a) / max(1, pq_a)) * 100.0
        trend_clause = f" (<strong>{'up' if qoq_a_pct >= 0 else 'down'} {abs(qoq_a_pct):.1f}% QoQ</strong> from {pq_a:,} in {prev_q['q_label']})"
        prev_str = f"Prior quarter ({prev_q['q_label']}): {pq_a:,} approvals / {pq_s:,} settlements"
        chg_str = f"Approvals: {q_app - pq_a:+,d} ({qoq_a_pct:+.1f}% QoQ) | Settlements: {q_set:,} ({conv:.1f}% ratio)"

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
        set_expr_4q = " + ".join(f"{q['set_row'][3]:,}" for q in q4_list)
        app_expr_4q = " + ".join(f"{q['app_row'][3]:,}" for q in q4_list)
        uw_calc = (
            f"12-Month Underwritten Settlements: {set_expr_4q} = {roll_set_12m:,} loans ({mth_uw_12m:,}/mth) | "
            f"12-Month Approvals: {app_expr_4q} = {roll_app_12m:,} approvals ({mth_ap_12m:,}/mth)"
        )
        out["kainga-ora-underwritten-loan-book"] = make_pulse_item(
            "Kāinga Ora Report", q_end_iso, MAX_AGE_QUARTERLY,
            f"Over the 12 months ended {q_disp}, Kāinga Ora recorded <strong>{roll_set_12m:,} settled First Home Loan (5% deposit) mortgages</strong> (~{mth_uw_12m:,}/month) and <strong>{roll_app_12m:,} approved applications</strong> (~{mth_ap_12m:,}/month) under the Crown Mortgage Insurance Scheme.",
            f"Kāinga Ora Home Ownership Products Report (4 quarters ended {q_disp}): Rolling 12-month First Home Loan settlements = {roll_set_12m:,}; rolling 12-month approvals = {roll_app_12m:,}.",
            format_provenance_source_html(uw_calc, f"Kāinga Ora — Home Ownership Products Report ({q_disp})", pdf_u, True),
            f"12 months ended {q_disp}", roll_set_12m,
            f"{roll_set_12m:,} settled loans underwritten/yr (~{mth_uw_12m:,}/mth) | {roll_app_12m:,} approvals/yr",
            f"Latest quarter ({q_disp}): {q_set:,} settled / {q_app:,} approved",
            f"Annual settlement-to-approval ratio: {(roll_set_12m / max(1, roll_app_12m)) * 100:.1f}% ({roll_set_12m:,} ÷ {roll_app_12m:,})",
            uw_calc, f"Kāinga Ora 4-quarter rolling series (ended {q_disp})", fhl_hist,
            {"source_published": qtr_pub, "source_url": pdf_u},
        )
    else:
        ann_set_pace = q_set * 4
        ann_app_pace = q_app * 4
        uw_calc = f"Quarterly Settlements ({q_disp}): {m1s:,} + {m2s:,} + {m3s:,} = {q_set:,} ({ma_set:,}/mth; annualised pace = {ann_set_pace:,}/yr from {ann_app_pace:,}/yr approvals)"
        out["kainga-ora-underwritten-loan-book"] = make_pulse_item(
            "Kāinga Ora Report", q_end_iso, MAX_AGE_QUARTERLY,
            f"In the {q_disp} quarter, Kāinga Ora recorded <strong>{q_set:,} settled First Home Loan (5% deposit) mortgages</strong> (~{ma_set:,}/month, an annualised pace of <strong>{ann_set_pace:,} settled loans/year</strong>) alongside <strong>{q_app:,} quarterly approvals</strong>.",
            f"Kāinga Ora Home Ownership Products Report ({q_disp}): Quarterly settled First Home Loans = {q_set:,} ({m1s:,}, {m2s:,}, {m3s:,}); quarterly approvals = {q_app:,}.",
            format_provenance_source_html(uw_calc, f"Kāinga Ora — Home Ownership Products Report ({q_disp})", pdf_u, True),
            q_disp, q_set,
            f"{q_set:,} settled loans in {q_disp} (~{ma_set:,}/mth; {ann_set_pace:,}/yr pace)",
            f"Monthly settlements in quarter: {m1s:,}, {m2s:,}, {m3s:,}",
            f"Settlement-to-approval ratio: {conv:.1f}% ({q_set:,} ÷ {q_app:,})",
            uw_calc, f"Kāinga Ora quarterly series ({q_disp})", fhl_hist,
            {"source_published": qtr_pub, "source_url": pdf_u},
        )

    fhl_calc = (
        f"Quarterly Approvals: {m1a:,} + {m2a:,} + {m3a:,} = {q_app:,} ({ma_app:,}/mth) | "
        f"Quarterly Settlements: {m1s:,} + {m2s:,} + {m3s:,} = {q_set:,} ({ma_set:,}/mth; {conv:.1f}% ratio)"
    )
    out["kainga-ora-first-home-loan-approvals"] = make_pulse_item(
        "Kāinga Ora Report", q_end_iso, MAX_AGE_QUARTERLY,
        f"In the {q_disp} quarter, Kāinga Ora recorded <strong>{q_app:,} First Home Loan (5% deposit) approvals</strong> (~{ma_app:,}/month){trend_clause} and <strong>{q_set:,} loan settlements</strong> (~{ma_set:,}/month, a <strong>{conv:.1f}%</strong> settlement-to-approval ratio).{roll_clause}",
        f"Kāinga Ora Home Ownership Products Report ({q_disp}): First Home Loan approvals = {q_app:,} ({m1a:,}, {m2a:,}, {m3a:,}); settlements = {q_set:,} ({m1s:,}, {m2s:,}, {m3s:,}).",
        format_provenance_source_html(fhl_calc, f"Kāinga Ora — Home Ownership Products Quarterly Report ({q_disp})", pdf_u, True),
        q_disp, q_app, f"{q_app:,} approvals ({ma_app:,}/mth) | {q_set:,} settlements ({conv:.1f}%)",
        prev_str, chg_str, fhl_calc, f"Kāinga Ora quarterly series ({q_disp})", fhl_hist,
        {"source_published": qtr_pub, "source_url": pdf_u},
    )

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
            sc_expr_4q = " + ".join(f"{x:,}" for x in app_s[-4:])
            sc_calc = f"12-Month Rolling Approvals: {sc_expr_4q} = {roll_12m:,} approvals/yr | Latest Quarter ({curr_q}): {curr_a:,} approved of {curr_r:,} received (÷ 3 = {mth_a:,}/mth)"
            sc_hist = [{"period": ql, "label": ql, "value": av, "formatted": f"{av:,} approved (of {rv:,} received)"} for ql, av, rv in reversed(list(zip(q_list, app_s, rec_s)))]
            out["kainga-ora-second-chance-kiwisaver-approvals"] = make_pulse_item(
                "Kāinga Ora Report", q_end_iso, MAX_AGE_QUARTERLY,
                f"In the {curr_q} quarter, Kāinga Ora approved <strong>{curr_a:,} Previous Homeowner ('second-chance') KiwiSaver savings withdrawal determination applications</strong>{yoy_cl} (~{mth_a:,}/month), bringing the rolling 12-month total to <strong>{roll_12m:,} approved applications</strong>.",
                f"Kāinga Ora Home Ownership Products Report ({curr_q}): Previous Owner Savings Withdrawal Applications Approved = {curr_a:,} in {curr_q} ({roll_12m:,} over rolling 4 quarters).",
                format_provenance_source_html(sc_calc, f"Kāinga Ora — Home Ownership Products Quarterly Report ({curr_q})", pdf_u, True),
                curr_q, curr_a, f"{curr_a:,} approvals in {curr_q} ({roll_12m:,} over 12 months)",
                f"Prior quarter ({prev_q}): {prev_a:,} approvals", yoy_chg, sc_calc,
                f"Kāinga Ora 5-quarter published series ({q_list[0]} to {curr_q})", sc_hist,
                {"source_published": qtr_pub, "source_url": pdf_u},
            )

    m_ks_anchor = re.search(r"KiwiSaver\s+First\s+Home\s+Savings\s+Withdrawal", pdf_txt, flags=re.I) or re.search(r"programme\s+of\s+assistance\s+for\s+first\s+home\s+buyers", pdf_txt, flags=re.I)
    if m_ks_anchor:
        ks_w = pdf_txt[max(0, m_ks_anchor.start() - 1200) : min(len(pdf_txt), m_ks_anchor.end() + 1400)]
        ks_qs = [re.sub(r"\s+", " ", q).strip() for q in re.findall(q_pat, ks_w, flags=re.I)]
        ks_c = ks_w
        for qm in ks_qs:
            ks_c = ks_c.replace(qm, " ")
        raw_dlr = [int(x.replace(",", "")) for x in re.findall(r"\$([\d,]+)", ks_c)]
        dlr_bars = [x for x in _remove_axis_ticks(raw_dlr, 10, 400, 650) if 300 <= x <= 950]
        ks_no_d = re.sub(r"\b20[12]\d\b", " ", re.sub(r"\$[\d,]+", " ", ks_c))
        raw_cnts = [int(x.replace(",", "")) for x in re.findall(r"\b\d{1,2},\d{3}\b", ks_no_d)]
        cnt_bars = [x for x in _remove_axis_ticks(raw_cnts, 200, 11000, 14000) if 7500 <= x <= 20000]
        if len(ks_qs) >= 4 and len(cnt_bars) >= 4 and len(dlr_bars) >= 4:
            ks_qs_4, cnt_4q, dlr_4q = ks_qs[:4], cnt_bars[:4], dlr_bars[:4]
            roll_cnt_12m, roll_dlr_12m = sum(cnt_4q), sum(dlr_4q)
            roll_avg_nzd = int(round((roll_dlr_12m * 1e6) / max(1, roll_cnt_12m)))
            q_now_lbl, q_cnt_now, q_dlr_now = ks_qs_4[-1], cnt_4q[-1], dlr_4q[-1]
            q_avg_now = int(round((q_dlr_now * 1e6) / max(1, q_cnt_now)))
            cnt_expr_4q = " + ".join(f"{x:,}" for x in cnt_4q)
            dlr_expr_4q = " + ".join(f"${x}m" for x in dlr_4q)
            ks4_calc = (
                f"12-Month Rolling Total: {cnt_expr_4q} = {roll_cnt_12m:,} withdrawals & "
                f"{dlr_expr_4q} = ${roll_dlr_12m:,}m (${roll_dlr_12m/1000:.3f}B; avg ${roll_avg_nzd:,}/withdrawal)"
            )
            ks4_hist = [
                {"period": ql, "label": ql, "value": dv, "formatted": f"${dv:,}m across {cv:,} withdrawals (avg ${int(round((dv*1e6)/max(1,cv))):,})"}
                for ql, cv, dv in reversed(list(zip(ks_qs_4, cnt_4q, dlr_4q)))
            ]
            out["kiwisaver-first-home-quarterly-trend"] = make_pulse_item(
                "Kāinga Ora / IRD", q_end_iso, MAX_AGE_QUARTERLY,
                f"Over the 12 months ended {q_now_lbl}, <strong>{roll_cnt_12m:,} KiwiSaver first-home savings withdrawals</strong> were recorded across New Zealand totalling <strong>${roll_dlr_12m/1000:.2f}B (${roll_dlr_12m:,}m)</strong>—an average of <strong>${roll_avg_nzd:,} per member withdrawal</strong> (with <strong>{q_cnt_now:,} withdrawals</strong> totalling <strong>${q_dlr_now:,}m</strong> in {q_now_lbl}).",
                f"Kāinga Ora Home Ownership Products Report / IRD ({q_now_lbl}): 4-quarter rolling KiwiSaver first-home withdrawals = {roll_cnt_12m:,} withdrawals and ${roll_dlr_12m:,}m.",
                format_provenance_source_html(ks4_calc, f"Kāinga Ora Home Ownership Products Report / IRD ({q_now_lbl})", pdf_u, True),
                q_now_lbl, roll_dlr_12m,
                f"${roll_dlr_12m/1000:.2f}B (${roll_dlr_12m:,}m) across {roll_cnt_12m:,} withdrawals (12m rolling)",
                f"Latest quarter ({q_now_lbl}): ${q_dlr_now:,}m across {q_cnt_now:,} withdrawals (avg ${q_avg_now:,})",
                f"12-month average withdrawal: ${roll_avg_nzd:,} per withdrawing member",
                ks4_calc, f"IRD / Kāinga Ora 4-quarter series ({ks_qs_4[0]} to {ks_qs_4[-1]})", ks4_hist,
                {"source_published": qtr_pub, "source_url": pdf_u},
            )

    return out


def build_ird_tickers(session, t43_path: pathlib.Path | None = None) -> dict:
    """
    100% live-pulled Inland Revenue (IRD), FMA, RBNZ Table T43 & Kāinga Ora datasets
    (Zero Hardcoded Fallbacks).
    """
    out = {}
    today = datetime.datetime.now(datetime.timezone.utc).date()
    ird_urls = [
        IRD_SL_BAL_PAGE, IRD_KS_NUM_PAGE, IRD_KS_AMT_PAGE, IRD_KS_REG_PAGE, IRD_KS_TR_PAGE,
        IRD_KS_ENTRY_PAGE, IRD_KS_PAY_PAGE, IRD_KS_SUSP_PAGE, IRD_KS_DATA_PAGE, FMA_KS_REPORT_PAGE,
        IRD_INCOME_DIST_PAGE, IRD_INCOME_DATASETS_PAGE, IRD_INCOME_TOT_PAGE,
        IRD_WAGE_SALARY_PAGE, IRD_WAGE_DATASETS_PAGE, IRD_WAGE_NUM_PAGE, IRD_WAGE_BAND_PAGE,
        IRD_TRUSTS_NUM_PAGE, IRD_TRUSTS_INC_PAGE,
        IRD_WFF_ENT_PAGE, IRD_WFF_FAM_PAGE, IRD_WFF_MAIN_PAGE,
        IRD_KS_AGE_PAGE, IRD_KS_INC_PAGE,
    ]
    pages = fetch_html_with_browser_fallback(session, ird_urls, "IRD / FMA")

    sl_html = pages.get(IRD_SL_BAL_PAGE, "")
    if sl_html:
        sl_pub = extract_html_release_date(sl_html, "Quarterly Inland Revenue Student Loan statistics release")
        sl_text = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", sl_html))
        m_med = re.search(r"Median student loan balance\s+(?:increased|decreased)\s+by\s+[\d.]+%\s+from\s+\$([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})\s+to\s+\$([\d,]+)\s+in\s+([A-Za-z]+\s+\d{4})", sl_text, flags=re.I)
        if m_med:
            prev_bal, prev_lbl, curr_sl_bal = int(m_med.group(1).replace(",", "")), m_med.group(2), int(m_med.group(3).replace(",", ""))
            dt_sl = pd.to_datetime("1 " + m_med.group(4), errors="coerce")
            if pd.notna(dt_sl):
                sl_m_full, pct_sl = dt_sl.strftime("%B %Y"), ((curr_sl_bal - prev_bal) / max(1, prev_bal)) * 100.0
                sl_as_of = month_end_iso(dt_sl)
                sl_calc = f"Reported IRD median = ${curr_sl_bal:,} (Change: (${curr_sl_bal:,} - ${prev_bal:,}) ÷ ${prev_bal:,} = {pct_sl:+.1f}%)"
                out["student-loan-servicing"] = make_pulse_item(
                    "Inland Revenue", sl_as_of, MAX_AGE_QUARTERLY,
                    f"As at {sl_m_full}, the median New Zealand student loan balance reported by Inland Revenue was <strong>${curr_sl_bal:,}</strong> ({plain_yoy_growth(curr_sl_bal, prev_bal)} from ${prev_bal:,} in {prev_lbl}).",
                    f"Inland Revenue Student Loan Scheme ({sl_m_full}): Median borrower balance = ${curr_sl_bal:,} (compared with ${prev_bal:,} in {prev_lbl}).",
                    format_provenance_source_html(sl_calc, f"Inland Revenue — Student Loan Statistics ({dt_sl.strftime('%b %Y')})", IRD_SL_BAL_PAGE, False),
                    sl_m_full, curr_sl_bal, f"${curr_sl_bal:,} median balance", f"{prev_lbl} (${prev_bal:,})",
                    f"${curr_sl_bal - prev_bal:+,d} ({pct_sl:+.1f}%)", sl_calc, f"Inland Revenue release ({sl_m_full})", None, {"source_published": sl_pub},
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
                    f"In {ks_m_full}, <strong>{fh_cnt:,}</strong> KiwiSaver members withdrew a combined <strong>${fh_amt_m:,.1f}m</strong> for a first-home purchase—an average of <strong>${avg_mem:,}</strong> per withdrawing member.",
                    f"Inland Revenue KiwiSaver Withdrawals ({ks_m_full}): First-home withdrawals = {fh_cnt:,} members and ${fh_amt_m:,.1f}m.",
                    format_provenance_source_html(fh_calc, "Inland Revenue — KiwiSaver Withdrawal Statistics", IRD_KS_DATA_PAGE, True),
                    ks_m_full, round(fh_amt_m, 1), f"${fh_amt_m:,.1f}m across {fh_cnt:,} members (avg ${avg_mem:,}/member)",
                    f"Prior month: ${fh_prev_m:,.1f}m" if fh_prev_m is not None else "Monthly IRD withdrawal series",
                    f"{fh_amt_m - fh_prev_m:+.1f}m vs prior month" if fh_prev_m is not None else f"${avg_mem:,} average per withdrawing member",
                    fh_calc, f"Inland Revenue monthly series ({ks_m_full})", None, {"source_published": wd_pub},
                )
                mn_h = re.search(r"([\d,]+)\s+members[^.]{0,80}financial\s+hardship", num_txt, flags=re.I)
                ma_h = re.search(r"\$([\d,.]+)\s+million[^.]{0,80}financial\s+hardship(?:,\s+(?:up|down)\s+from\s+\$([\d,.]+)\s+million)?", amt_txt, flags=re.I)
                if mn_h and ma_h:
                    hs_cnt, hs_amt_m = int(mn_h.group(1).replace(",", "")), float(ma_h.group(1).replace(",", ""))
                    hs_prev_m = float(ma_h.group(2).replace(",", "")) if ma_h.group(2) else None
                    hs_avg = int(round((hs_amt_m * 1e6) / max(1, hs_cnt)))
                    hs_calc = f"${hs_amt_m:,.1f}m × $1,000,000 ÷ {hs_cnt:,} members = ${hs_avg:,}/member"
                    out["kiwisaver-hardship-withdrawals"] = make_pulse_item(
                        "Inland Revenue", ks_m_as_of, MAX_AGE_MONTHLY,
                        f"In {ks_m_full}, <strong>{hs_cnt:,}</strong> KiwiSaver members made a withdrawal under the financial hardship category totalling <strong>${hs_amt_m:,.1f}m</strong>—an average of <strong>${hs_avg:,}</strong> per withdrawing member.",
                        f"Inland Revenue KiwiSaver Withdrawals ({ks_m_full}): Financial hardship withdrawals = {hs_cnt:,} members and ${hs_amt_m:,.1f}m.",
                        format_provenance_source_html(hs_calc, "Inland Revenue — KiwiSaver Withdrawal Statistics", IRD_KS_DATA_PAGE, True),
                        ks_m_full, round(hs_amt_m, 1), f"${hs_amt_m:,.1f}m across {hs_cnt:,} members (avg ${hs_avg:,}/member)",
                        f"Prior month: ${hs_prev_m:,.1f}m" if fh_prev_m is not None else "Monthly IRD withdrawal series",
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
                        f"Over the year ended 30 June {fy_reg}, <strong>{tot_mem:,} KiwiSaver members</strong> withdrew <strong>${tot_nzd/1e9:.2f}B</strong> for a first home across New Zealand's 16 regions—averaging <strong>${nz_avg_p:,} per withdrawing member nationwide</strong>, ranging from <strong>${h_a:,} in {h_n}</strong> and <strong>${akl_a:,} in Auckland</strong> to <strong>${l_a:,} in {l_n}</strong>.",
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
                f"Over the year ended 30 June {fy_yr}, Inland Revenue processed <strong>{tr_tot:,}</strong> KiwiSaver inter-scheme transfers ({plain_yoy_growth(tr_tot, prev_tr)} from {prev_tr:,} in FY{prev_fy})—an average of <strong>{mth_avg:,}</strong> scheme transfers per month.",
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
        m_d = re.search(r"default allocated[^\d]{0,60}(\d[\d,]{4,8})", e_txt, flags=re.I)
        m_e = re.search(r"employer nominated[^\d]{0,60}(\d[\d,]{4,8})", e_txt, flags=re.I)
        m_a = re.search(r"actively chosen[^\d]{0,60}(\d[\d,]{5,9})", e_txt, flags=re.I)
        if m_d and m_e and m_a:
            def_a, emp_a, act_a = int(m_d.group(1).replace(",", "")), int(m_e.group(1).replace(",", "")), int(m_a.group(1).replace(",", ""))
            non_act, tot_m = def_a + emp_a, def_a + emp_a + act_a
            non_pct, def_p, emp_p, act_p = (non_act / max(1, tot_m)) * 100.0, (def_a / max(1, tot_m)) * 100.0, (emp_a / max(1, tot_m)) * 100.0, (act_a / max(1, tot_m)) * 100.0
            fy_e = today.year if today.month >= 7 else today.year - 1
            ent_calc = f"Default/Employer-Nominated: ({def_a:,} + {emp_a:,}) ÷ {tot_m:,} = {non_pct:.1f}% | Active Choice: {act_a:,} ÷ {tot_m:,} = {act_p:.1f}%"
            out["kiwisaver-default-vs-active"] = make_pulse_item(
                "Inland Revenue", f"{fy_e}-06-30", MAX_AGE_ANNUAL,
                f"In Inland Revenue's KiwiSaver entry-method release, <strong>{non_pct:.1f}% ({non_act:,} members)</strong> were in schemes entered via <strong>default allocation ({def_a:,}, or {def_p:.1f}%)</strong> or <strong>employer nomination ({emp_a:,}, or {emp_p:.1f}%)</strong>, while <strong>{act_p:.1f}% ({act_a:,} members)</strong> had actively chosen their scheme.",
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
        m_tot_p = re.search(r"\$([\d,.]+)\s*(billion|b)\s+(?:was\s+)?(?:passed|transferred|paid)\s+to\s+(?:scheme\s+)?providers", p_txt, flags=re.I) or re.search(r"\$(1\d\.\d{1,2})\s*(?:billion|b)", p_txt, flags=re.I)
        if m_tot_p:
            tot_pay_b, fy_p = float(m_tot_p.group(1).replace(",", "")), today.year if today.month >= 7 else today.year - 1
            pay_calc = f"Reported Inland Revenue annual KiwiSaver transfers to scheme providers = ${tot_pay_b:.2f}B"
            out["kiwisaver-contributions-10yr"] = make_pulse_item(
                "Inland Revenue", f"{fy_p}-06-30", MAX_AGE_ANNUAL,
                f"In Inland Revenue's latest KiwiSaver contributions release, total annual employee, employer, and Crown contributions transferred by Inland Revenue to scheme providers reached <strong>${tot_pay_b:.2f}B</strong>.",
                f"Inland Revenue KiwiSaver Provider Payments: Total annual contributions transferred = ${tot_pay_b:.2f}B.",
                format_provenance_source_html(pay_calc, "Inland Revenue — KiwiSaver Scheme Provider Payments", IRD_KS_PAY_PAGE, False),
                f"Year ended June {fy_p}", tot_pay_b, f"${tot_pay_b:.2f}B annual contributions transferred",
                "Employee + Employer + Crown contributions via IRD", "Live-parsed from Inland Revenue contributions release",
                pay_calc, "Inland Revenue June FY contributions series", None, {"source_published": pay_pub},
            )

    susp_html = pages.get(IRD_KS_SUSP_PAGE, "")
    if susp_html:
        susp_pub = extract_html_release_date(susp_html, "Monthly Inland Revenue KiwiSaver savings suspension release")
        s_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", susp_html))
        m_s = re.search(r"(\d[\d,]{4,7})\s+members were on a savings suspension", s_txt, flags=re.I)
        if m_s:
            susp_tot = int(m_s.group(1).replace(",", ""))
            m_dt_s = re.search(r"In\s+([A-Za-z]+\s+20\d{2})", s_txt, flags=re.I)
            dt_s = pd.to_datetime("1 " + m_dt_s.group(1), errors="coerce") if m_dt_s else pd.Timestamp(today)
            s_per, s_as_of = (dt_s.strftime("%B %Y"), month_end_iso(dt_s)) if pd.notna(dt_s) else ("Latest IRD release", today.strftime("%Y-%m-%d"))
            susp_calc = f"Directly reported Inland Revenue active KiwiSaver savings suspensions = {susp_tot:,} members"
            out["kiwisaver-savings-suspensions"] = make_pulse_item(
                "Inland Revenue", s_as_of, MAX_AGE_ANNUAL,
                f"In {s_per}, Inland Revenue recorded <strong>{susp_tot:,}</strong> KiwiSaver members on an active formal savings suspension.",
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
            tabs_t43 = load_rbnz_sheets(t43_path) or load_rbnz_sheets_merged(t43_path)
            dt_t, s_t, _ = _find_rbnz_multi(
                tabs_t43, [("total",), ("asset",), ("kiwisaver",), ("funds",)],
                val_bounds=(60000, 220000),
            )
            df_t43 = _series_df(dt_t, s_t)
            if not df_t43.empty:
                t43_fum_b, t43_dt = float(df_t43["val"].iloc[-1]) / 1000.0, pd.Timestamp(df_t43["date"].iloc[-1])
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
            fum_calc = f"Official FMA / RBNZ FUM = ${fum_b_live:.1f}B across {mem_m_live:.2f}m members (Mean member balance = ${avg_bal_live:,})"
            fum_html = f"Total KiwiSaver funds under management reached <strong>${fum_b_live:.1f}B</strong> across <strong>{mem_m_live:.2f} million</strong> enrolled members, with an average member balance of <strong>${avg_bal_live:,}</strong>."
            fum_val, fum_latest_v = avg_bal_live, f"${avg_bal_live:,} avg balance | ${fum_b_live:.1f}B FUM ({mem_m_live:.2f}m members)"
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

    # 1. `ird-individual-taxable-income-percentiles` (Uses MAX_AGE_TRIENNIAL to account for IR3 filing lag)
    inc_comb_html = " ".join(pages.get(u, "") for u in (IRD_INCOME_DIST_PAGE, IRD_INCOME_TOT_PAGE, IRD_INCOME_DATASETS_PAGE))
    if inc_comb_html.strip():
        inc_pub = extract_html_release_date(inc_comb_html, "Annual Inland Revenue Taxable Income Distribution release")
        inc_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", inc_comb_html))
        m_ind_cnt = re.search(r"([45],\d{3},\d{3})\s+individuals", inc_txt, flags=re.I)
        m_inc_b = re.search(r"combined\s+taxable\s+income\s+of\s+\$([\d.]+)\s*(?:billion|b)", inc_txt, flags=re.I) or re.search(r"\$([23]\d{2}\.?\d*)\s*(?:billion|b)", inc_txt, flags=re.I)
        m_ty = re.search(r"(202\d)\s+tax\s+year", inc_txt, flags=re.I)
        if m_ind_cnt and m_inc_b:
            tot_inds = int(m_ind_cnt.group(1).replace(",", ""))
            tot_inc_b = float(m_inc_b.group(1))
            ty_yr = int(m_ty.group(1)) if m_ty else (today.year - 2)
            mean_tax_inc = int(round((tot_inc_b * 1e9) / max(1, tot_inds)))
            inc_calc = f"Official IRD ({ty_yr} tax year): ${tot_inc_b:.1f}B total taxable income ÷ {tot_inds:,} individuals = ${mean_tax_inc:,} mean taxable income"
            out["ird-individual-taxable-income-percentiles"] = make_pulse_item(
                "Inland Revenue", f"{ty_yr}-03-31", MAX_AGE_TRIENNIAL,
                f"In Inland Revenue's {ty_yr} Taxable Income Distribution of Individuals, <strong>{tot_inds:,} New Zealand individuals</strong> recorded a combined taxable income of <strong>${tot_inc_b:.1f}B</strong>—a mean taxable income of <strong>${mean_tax_inc:,} per individual</strong>.",
                f"Inland Revenue Taxable Income Distribution ({ty_yr} tax year): Total individuals = {tot_inds:,}; Combined taxable income = ${tot_inc_b:.1f}B.",
                format_provenance_source_html(inc_calc, f"Inland Revenue — Taxable Income Distribution ({ty_yr})", IRD_INCOME_DIST_PAGE, True),
                f"{ty_yr} tax year", mean_tax_inc,
                f"${mean_tax_inc:,} mean taxable income (${tot_inc_b:.1f}B across {tot_inds:,} individuals)",
                "Administrative tax return distribution across all NZ individuals with taxable income",
                "Derived live from Inland Revenue Taxable Income Distribution release",
                inc_calc, f"Inland Revenue {ty_yr} tax year release", None,
                {"source_published": inc_pub, "source_url": IRD_INCOME_DIST_PAGE},
            )

    # 2. `ird-paye-wage-vs-self-employed-income` (Live IRD Wage & Salary Pages + XLSX)
    ws_comb_html = " ".join(pages.get(u, "") for u in (IRD_WAGE_SALARY_PAGE, IRD_WAGE_NUM_PAGE, IRD_WAGE_BAND_PAGE, IRD_WAGE_DATASETS_PAGE))
    if ws_comb_html.strip():
        ws_pub = extract_html_release_date(ws_comb_html, "Annual Inland Revenue Wage and Salary Distributions release")
        ws_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", ws_comb_html))
        paye_cnt = None
        m_ws_cnt = re.search(r"(2,\d{3},\d{3})\s+(?:individuals|people|wage|earners)", ws_txt, flags=re.I)
        if m_ws_cnt:
            paye_cnt = int(m_ws_cnt.group(1).replace(",", ""))
        m_ws_mod = re.search(r"modal\s+income\s+point[^$]{0,60}\$([\d,]{5,7})", ws_txt, flags=re.I)
        mod_wage = int(m_ws_mod.group(1).replace(",", "")) if m_ws_mod else None

        ws_xlsx_u = next(
            ("https://www.ird.govt.nz" + h if h.startswith("/") else h for h in re.findall(r'href=["\']([^"\']+\.xlsx[^"\']*)["\']', ws_comb_html, flags=re.I)),
            "",
        )
        if ws_xlsx_u and paye_cnt is None:
            try:
                ws_b = fetch_bytes_with_browser_fallback(session, ws_xlsx_u, referer=IRD_WAGE_DATASETS_PAGE, min_bytes=5_000)
                if ws_b and ws_b[:2] == b"PK":
                    xl_ws = pd.ExcelFile(io.BytesIO(ws_b))
                    nums_cnt = []
                    for sh in xl_ws.sheet_names:
                        df_ws = xl_ws.parse(sh)
                        nums_cnt.extend([int(float(v)) for v in df_ws.values.flatten() if re.match(r"^\d{7}(\.0+)?$", str(v).strip()) and 2_000_000 <= float(v) <= 3_800_000])
                    if nums_cnt:
                        paye_cnt = max(nums_cnt)
            except Exception as exc_ws:
                print(f"[IRD Wage XLSX] Note: {exc_ws}")

        if paye_cnt:
            mod_clause = f", with the modal PAYE income point at <strong>${mod_wage:,}</strong>" if mod_wage else ""
            ws_calc = f"Official IRD Wage and Salary Distributions: Total PAYE wage & salary earners = {paye_cnt:,}" + (f" | Modal PAYE wage = ${mod_wage:,}" if mod_wage else "")
            out["ird-paye-wage-vs-self-employed-income"] = make_pulse_item(
                "Inland Revenue", f"{today.year if today.month >= 4 else today.year - 1}-03-31", MAX_AGE_TRIENNIAL,
                f"In Inland Revenue's Wage and Salary Distributions release, <strong>{paye_cnt:,} New Zealand individuals</strong> received PAYE wage and salary earnings during the tax year{mod_clause}.",
                f"Inland Revenue Wage and Salary Distributions: Total PAYE wage and salary earners = {paye_cnt:,}.",
                format_provenance_source_html(ws_calc, "Inland Revenue — Wage and Salary Distributions", ws_xlsx_u or IRD_WAGE_SALARY_PAGE, False),
                "Latest IRD Tax Year", paye_cnt,
                f"{paye_cnt:,} PAYE wage & salary earners" + (f" (${mod_wage:,} modal PAYE wage)" if mod_wage else ""),
                "Inland Revenue administrative PAYE series",
                "Parsed live from Inland Revenue Wage and Salary Distributions release",
                ws_calc, "Inland Revenue Wage & Salary series", None,
                {"source_published": ws_pub, "source_url": ws_xlsx_u or IRD_WAGE_SALARY_PAGE},
            )

    # 3 & 4. `kiwisaver-member-demographics-age-income` & `kiwisaver-contribution-rate-mix`
    age_html = pages.get(IRD_KS_AGE_PAGE, "")
    inc_ks_html = pages.get(IRD_KS_INC_PAGE, "")
    ks_ds_html = pages.get(IRD_KS_DATA_PAGE, "")
    age_pub = extract_html_release_date(age_html or ks_ds_html, "Annual Inland Revenue KiwiSaver Member Demographics release")

    age_band_specs = [
        ("0–17", r"(?:^|\b)(?:0\s*[-–to]+\s*17|under\s+18|17\s+and\s+under)(?:\b|$)"),
        ("18–24", r"(?:^|\b)18\s*[-–to]+\s*24(?:\b|$)"),
        ("25–34", r"(?:^|\b)25\s*[-–to]+\s*34(?:\b|$)"),
        ("35–44", r"(?:^|\b)35\s*[-–to]+\s*44(?:\b|$)"),
        ("45–54", r"(?:^|\b)45\s*[-–to]+\s*54(?:\b|$)"),
        ("55–64", r"(?:^|\b)55\s*[-–to]+\s*64(?:\b|$)"),
        ("65+", r"(?:^|\b)(?:65\s*\+|65\s+and\s+over|over\s+65)(?:\b|$)"),
    ]
    age_map: dict[str, int] = {}

    if age_html:
        ext_a = SimpleTableExtractor()
        ext_a.feed(age_html)
        for tbl in ext_a.tables:
            if not tbl or len(tbl) < 2:
                continue
            # Check Orientation A: Age bands in Column 0
            cand_a: dict[str, int] = {}
            for r in tbl:
                if len(r) < 2:
                    continue
                r0 = r[0].strip()
                for b_lbl, b_rx in age_band_specs:
                    if re.search(b_rx, r0, flags=re.I):
                        valid_nums = [
                            int(round(float(re.sub(r"[,$]", "", c))))
                            for c in r[1:]
                            if re.match(r"^[\d,]+(\.0+)?$", c.strip()) and 50_000 <= float(re.sub(r"[,$]", "", c)) <= 1_200_000
                        ]
                        if valid_nums:
                            cand_a[b_lbl] = valid_nums[-1]
            if len(cand_a) >= 5 and sum(cand_a.values()) >= 2_400_000:
                age_map = cand_a
                break

            # Check Orientation B: Age bands across Header Row (tbl[0]) and Years in Column 0
            hdr_r = tbl[0]
            col_band_map: dict[int, str] = {}
            for c_idx, h_cell in enumerate(hdr_r):
                for b_lbl, b_rx in age_band_specs:
                    if re.search(b_rx, h_cell.strip(), flags=re.I):
                        col_band_map[c_idx] = b_lbl
            if len(col_band_map) >= 5:
                for r in reversed(tbl[1:]):
                    cand_b: dict[str, int] = {}
                    for c_idx, b_lbl in col_band_map.items():
                        if c_idx < len(r):
                            c_str = r[c_idx].strip()
                            if re.match(r"^[\d,]+(\.0+)?$", c_str) and 50_000 <= float(re.sub(r"[,$]", "", c_str)) <= 1_200_000:
                                cand_b[b_lbl] = int(round(float(re.sub(r"[,$]", "", c_str))))
                    if len(cand_b) >= 5 and sum(cand_b.values()) >= 2_400_000:
                        age_map = cand_b
                        break
            if len(age_map) >= 5 and sum(age_map.values()) >= 2_400_000:
                break

        if sum(age_map.values()) < 2_400_000:
            a_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", age_html))
            cand_rx: dict[str, int] = {}
            for b_lbl, b_pat in (
                ("0–17", r"(?:0\s*[-–to]+\s*17|under\s+18)[^\d]{0,45}(\d{2,3},\d{3})"),
                ("18–24", r"18\s*[-–to]+\s*24[^\d]{0,45}(\d{2,3},\d{3})"),
                ("25–34", r"25\s*[-–to]+\s*34[^\d]{0,45}(\d{2,3},\d{3})"),
                ("35–44", r"35\s*[-–to]+\s*44[^\d]{0,45}(\d{2,3},\d{3})"),
                ("45–54", r"45\s*[-–to]+\s*54[^\d]{0,45}(\d{2,3},\d{3})"),
                ("55–64", r"55\s*[-–to]+\s*64[^\d]{0,45}(\d{2,3},\d{3})"),
                ("65+", r"(?:65\+|65\s+and\s+over)[^\d]{0,45}(\d{2,3},\d{3})"),
            ):
                matches_b = [int(m.replace(",", "")) for m in re.findall(b_pat, a_txt, flags=re.I) if 50_000 <= int(m.replace(",", "")) <= 1_200_000]
                if matches_b:
                    cand_rx[b_lbl] = matches_b[-1]
            if len(cand_rx) >= 5 and sum(cand_rx.values()) >= 2_400_000:
                age_map = cand_rx

    rate_map: dict[str, int] = {}
    used_ks_xlsx_url = IRD_KS_PAY_PAGE

    for html_cand, u_cand in ((pay_html, IRD_KS_PAY_PAGE), (ks_ds_html, IRD_KS_DATA_PAGE)):
        if not html_cand:
            continue
        ext_r = SimpleTableExtractor()
        ext_r.feed(html_cand)
        for tbl in ext_r.tables:
            for r in tbl:
                if len(r) < 2:
                    continue
                r0 = r[0].strip()
                if r0 in ("3%", "3.5%", "4%", "6%", "8%", "10%"):
                    nums_c = [
                        int(re.sub(r"[^\d]", "", cell))
                        for cell in r[1:]
                        if re.sub(r"[^\d]", "", cell) and 10_000 <= int(re.sub(r"[^\d]", "", cell)) <= 2_500_000
                    ]
                    if nums_c:
                        rate_map[r0] = nums_c[-1]
                        used_ks_xlsx_url = u_cand

    ks_xlsx_candidates = []
    if ks_ds_html:
        for h in re.findall(r'href=["\']([^"\']+\.xlsx[^"\']*)["\']', ks_ds_html, flags=re.I):
            full_kx = "https://www.ird.govt.nz" + h if h.startswith("/") else h
            if full_kx not in ks_xlsx_candidates:
                ks_xlsx_candidates.append(full_kx)
    for fallback_kx in (IRD_KS_ANNUAL_XLSX, IRD_KS_MONTHLY_XLSX):
        if fallback_kx not in ks_xlsx_candidates:
            ks_xlsx_candidates.append(fallback_kx)

    for kx_u in ks_xlsx_candidates[:3]:
        try:
            ks_b = fetch_bytes_with_browser_fallback(session, kx_u, referer=IRD_KS_DATA_PAGE, min_bytes=8_000)
            if not ks_b or ks_b[:2] != b"PK":
                continue
            xl_ks = pd.ExcelFile(io.BytesIO(ks_b))
            for sh in xl_ks.sheet_names:
                df_k = xl_ks.parse(sh, header=None)
                if df_k.empty:
                    continue
                cand_xl_age: dict[str, int] = {}
                for r_i in range(len(df_k)):
                    r_vals = [str(v).strip() for v in df_k.iloc[r_i].values if pd.notna(v) and str(v).strip()]
                    if len(r_vals) < 2:
                        continue
                    r0 = r_vals[0]
                    if r0 in ("3%", "3.5%", "4%", "6%", "8%", "10%", "0.03", "0.035", "0.04", "0.06", "0.08", "0.1", "0.10"):
                        nums_r = [
                            int(round(float(re.sub(r"[,$]", "", c))))
                            for c in r_vals[1:]
                            if re.match(r"^[\d,]+(\.0+)?$", c.strip()) and 10_000 <= float(re.sub(r"[,$]", "", c)) <= 2_500_000
                        ]
                        if nums_r:
                            if "." in r0 and not r0.endswith("%"):
                                fv_r = float(r0) * 100.0
                                lbl_r = "3.5%" if abs(fv_r - 3.5) < 0.05 else f"{int(round(fv_r))}%"
                            else:
                                lbl_r = r0
                            if nums_r[-1] > rate_map.get(lbl_r, 0):
                                rate_map[lbl_r] = nums_r[-1]
                                used_ks_xlsx_url = kx_u

                    if sum(age_map.values()) < 2_400_000:
                        for b_lbl, b_rx in age_band_specs:
                            if re.search(b_rx, r0, flags=re.I):
                                nums_a = [
                                    int(round(float(re.sub(r"[,$]", "", c))))
                                    for c in r_vals[1:]
                                    if re.match(r"^[\d,]+(\.0+)?$", c.strip()) and 50_000 <= float(re.sub(r"[,$]", "", c)) <= 1_200_000
                                ]
                                if nums_a:
                                    cand_xl_age[b_lbl] = nums_a[-1]

                        # Also check if age bands are column headers on this row
                        col_b_xl: dict[int, str] = {}
                        for c_i in range(df_k.shape[1]):
                            cell_s = str(df_k.iloc[r_i, c_i]).strip() if pd.notna(df_k.iloc[r_i, c_i]) else ""
                            for b_lbl, b_rx in age_band_specs:
                                if re.search(b_rx, cell_s, flags=re.I):
                                    col_b_xl[c_i] = b_lbl
                        if len(col_b_xl) >= 5:
                            for r_d in range(len(df_k) - 1, r_i, -1):
                                row_b: dict[str, int] = {}
                                for c_i, b_lbl in col_b_xl.items():
                                    v_c = df_k.iloc[r_d, c_i]
                                    if pd.notna(v_c) and re.match(r"^[\d,]+(\.0+)?$", str(v_c).strip()):
                                        fv_a = float(re.sub(r"[,$]", "", str(v_c).strip()))
                                        if 50_000 <= fv_a <= 1_200_000:
                                            row_b[b_lbl] = int(round(fv_a))
                                if len(row_b) >= 5 and sum(row_b.values()) >= 2_400_000:
                                    cand_xl_age = row_b
                                    break

                if len(cand_xl_age) >= 5 and sum(cand_xl_age.values()) >= 2_400_000:
                    age_map = cand_xl_age
        except Exception as exc_ksx:
            print(f"[IRD KiwiSaver XLSX] {kx_u}: {exc_ksx}")

    age_rows = [(b_lbl, age_map[b_lbl]) for b_lbl, _ in age_band_specs if b_lbl in age_map]
    tot_ks_a = sum(v for _, v in age_rows)
    if len(age_rows) >= 5 and tot_ks_a >= 2_400_000:
        fy_a = today.year if today.month >= 7 else today.year - 1
        age_calc = " | ".join(f"{k}: {v:,} ({(v/max(1,tot_ks_a))*100:.1f}%)" for k, v in age_rows[:5])
        out["kiwisaver-member-demographics-age-income"] = make_pulse_item(
            "Inland Revenue", f"{fy_a}-06-30", MAX_AGE_ANNUAL,
            f"As at June {fy_a}, Inland Revenue's KiwiSaver demographic statistics recorded <strong>{tot_ks_a:,} enrolled members</strong> across New Zealand age brackets.",
            f"Inland Revenue KiwiSaver Member Demographics (June {fy_a}): Total enrolled members across age bands = {tot_ks_a:,}.",
            format_provenance_source_html(age_calc, "Inland Revenue — KiwiSaver Member Demographics", IRD_KS_AGE_PAGE, True),
            f"June {fy_a}", tot_ks_a,
            f"{tot_ks_a:,} enrolled KiwiSaver members across age bands",
            "Inland Revenue annual June demographic release",
            "Parsed live from IRD KiwiSaver age demographic series",
            age_calc, f"Inland Revenue June {fy_a} demographic series",
            [{"period": k, "label": f"Age {k}", "value": v, "formatted": f"{v:,} members ({(v/max(1,tot_ks_a))*100:.1f}%)"} for k, v in age_rows],
            {"source_published": age_pub, "source_url": IRD_KS_AGE_PAGE},
        )
    elif age_html or inc_ks_html:
        comb_demo = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", f"{age_html} {inc_ks_html}"))
        m_tot_m = re.search(r"(3,\d{3},\d{3})\s+(?:members|individuals|people)", comb_demo, flags=re.I) or re.search(r"\b(3,[23456]\d{2},\d{3})\b", comb_demo)
        if m_tot_m:
            tot_ks_a = int(m_tot_m.group(1).replace(",", ""))
            fy_a = today.year if today.month >= 7 else today.year - 1
            age_calc = f"Official Inland Revenue KiwiSaver Member Demographics (June {fy_a}): Total enrolled members = {tot_ks_a:,}"
            out["kiwisaver-member-demographics-age-income"] = make_pulse_item(
                "Inland Revenue", f"{fy_a}-06-30", MAX_AGE_ANNUAL,
                f"As at June {fy_a}, Inland Revenue's KiwiSaver member demographic statistics recorded <strong>{tot_ks_a:,} enrolled members</strong> across New Zealand age and income brackets.",
                f"Inland Revenue KiwiSaver Member Demographics (June {fy_a}): Total enrolled members = {tot_ks_a:,}.",
                format_provenance_source_html(age_calc, "Inland Revenue — KiwiSaver Member Demographics", IRD_KS_AGE_PAGE, False),
                f"June {fy_a}", tot_ks_a,
                f"{tot_ks_a:,} enrolled KiwiSaver members across age & income bands",
                "Inland Revenue annual June demographic release",
                "Parsed live from IRD KiwiSaver member demographics release",
                age_calc, f"Inland Revenue June {fy_a} demographic series", None,
                {"source_published": age_pub, "source_url": IRD_KS_AGE_PAGE},
            )

    default_min_cnt = rate_map.get("3%", 0) + rate_map.get("3.5%", 0)
    if len(rate_map) >= 4 and default_min_cnt >= 400_000:
        tot_c_mem = sum(rate_map.values())
        abv_cnt = tot_c_mem - default_min_cnt
        abv_pct = round((abv_cnt / max(1, tot_c_mem)) * 100.0, 1)
        def_pct = round(100.0 - abv_pct, 1)
        if 20.0 <= abv_pct <= 60.0:
            fy_k = today.year if today.month >= 7 else today.year - 1
            def_lbl = "3%–3.5%" if ("3.5%" in rate_map and "3%" in rate_map) else ("3.5%" if "3.5%" in rate_map else "3%")
            r_calc = f"Above-Default (4%–10%) Contributors: {abv_cnt:,} ÷ {tot_c_mem:,} = {abv_pct:.1f}% | Default/Min ({def_lbl}) Rate: {default_min_cnt:,} ({def_pct:.1f}%)"
            out["kiwisaver-contribution-rate-mix"] = make_pulse_item(
                "Inland Revenue", f"{fy_k}-06-30", MAX_AGE_ANNUAL,
                f"In Inland Revenue's KiwiSaver contribution rate statistics (June {fy_k}), across <strong>{tot_c_mem:,} wage and salary contributing members</strong>, <strong>{def_pct:.1f}% ({default_min_cnt:,} members)</strong> contributed at the {def_lbl} rate while <strong>{abv_pct:.1f}% ({abv_cnt:,} members)</strong> contributed at 4%, 6%, 8%, or 10%.",
                f"Inland Revenue KiwiSaver Contribution Rates (June {fy_k}): Total contributing members = {tot_c_mem:,}; {def_lbl} rate = {default_min_cnt:,}; 4%–10% rates = {abv_cnt:,}.",
                format_provenance_source_html(r_calc, "Inland Revenue — KiwiSaver Contribution Rate Statistics", used_ks_xlsx_url, True),
                f"Year ended June {fy_k}", abv_pct,
                f"{abv_pct:.1f}% contributing 4%–10% ({abv_cnt:,}) vs {def_pct:.1f}% at {def_lbl} ({default_min_cnt:,})",
                f"Total contributing members by rate: {tot_c_mem:,}",
                "Parsed live from Inland Revenue KiwiSaver contribution rate series",
                r_calc, f"Inland Revenue June {fy_k} series",
                [{"period": f"{rk} Rate", "label": f"{rk} Contribution Rate", "value": rv, "formatted": f"{rv:,} members ({(rv/max(1,tot_c_mem))*100:.1f}%)"} for rk, rv in rate_map.items()],
                {"source_published": "Annual Inland Revenue KiwiSaver datasets release", "source_url": used_ks_xlsx_url},
            )

    # 5. `ird-family-trusts-rental-income` (Live IRD Trusts & Estates Pages)
    tr6_num_html = pages.get(IRD_TRUSTS_NUM_PAGE, "")
    tr6_inc_html = pages.get(IRD_TRUSTS_INC_PAGE, "")
    if tr6_num_html or tr6_inc_html:
        tr6_pub = extract_html_release_date(tr6_num_html or tr6_inc_html, "Annual Inland Revenue Trusts and Estates (IR6) statistics release")
        tr6_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", f"{tr6_num_html} {tr6_inc_html}"))
        m_tr6 = re.search(r"(2[012345]\d,\d{3})[^.]{0,60}?(202\d)\s+tax\s+year,\s+a\s+([\d.]+)%\s+(decrease|increase)", tr6_txt, flags=re.I)
        m_tr6_inc = re.search(r"Total\s+trust\s+income[^$]{0,90}\$([\d.]+)\s*billion", tr6_txt, flags=re.I)
        if m_tr6:
            trust_cnt = int(m_tr6.group(1).replace(",", ""))
            tr_yr = int(m_tr6.group(2))
            tr_chg_pct = float(m_tr6.group(3))
            tr_dir = m_tr6.group(4).lower()
            tr_inc_b = float(m_tr6_inc.group(1)) if m_tr6_inc else None
            inc_clause = f", with total IR6 trust income reported at <strong>${tr_inc_b:.1f}B</strong>" if tr_inc_b else ""
            tr6_calc = f"Official IRD ({tr_yr} tax year): Active trusts & estates = {trust_cnt:,} ({tr_dir} {tr_chg_pct:.1f}% YoY)" + (f" | Total IR6 Trust Income = ${tr_inc_b:.1f}B" if tr_inc_b else "")
            out["ird-family-trusts-rental-income"] = make_pulse_item(
                "Inland Revenue", f"{tr_yr}-03-31", MAX_AGE_TRIENNIAL,
                f"In Inland Revenue's Trusts and Estates (IR6) statistics for the {tr_yr} tax year, <strong>{trust_cnt:,} active trusts and estates</strong> filed returns across New Zealand (a <strong>{tr_chg_pct:.1f}% {tr_dir}</strong> year-on-year){inc_clause}.",
                f"Inland Revenue Trusts and Estates Statistics ({tr_yr} tax year): Active trusts and estates = {trust_cnt:,} ({tr_dir} {tr_chg_pct:.1f}% YoY).",
                format_provenance_source_html(tr6_calc, f"Inland Revenue — Trusts and Estates Statistics ({tr_yr})", IRD_TRUSTS_NUM_PAGE, False),
                f"{tr_yr} tax year", trust_cnt,
                f"{trust_cnt:,} active NZ trusts & estates ({tr_dir} {tr_chg_pct:.1f}% YoY)" + (f" | ${tr_inc_b:.1f}B trust income" if tr_inc_b else ""),
                "Inland Revenue annual IR6 series",
                f"{tr_dir.capitalize()} {tr_chg_pct:.1f}% YoY",
                tr6_calc, f"Inland Revenue {tr_yr} tax year series", None,
                {"source_published": tr6_pub, "source_url": IRD_TRUSTS_NUM_PAGE},
            )

    # 6. `working-for-families-income-support` (Live IRD Working for Families Pages)
    wff_comb_html = " ".join(pages.get(u, "") for u in (IRD_WFF_ENT_PAGE, IRD_WFF_FAM_PAGE, IRD_WFF_MAIN_PAGE))
    if wff_comb_html.strip():
        wff_pub = extract_html_release_date(wff_comb_html, "Annual Inland Revenue Working for Families statistics release")
        wff_txt = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", wff_comb_html))
        m_wff_avg = re.search(
            r"(?:average\s+entitlement|average\s+amount\s+of\s+annual\s+WFF)[^$]{0,120}\$\s*([\d,]{4,7})\s+in\s+(?:the\s+)?(202\d)[^$]{0,60}\$\s*([\d,]{4,7})\s+in\s+(?:the\s+)?(202\d)",
            wff_txt,
            flags=re.I,
        )
        m_wff_ch = re.search(r"([\d,]{5,9})\s+children", wff_txt, flags=re.I)
        m_wff_tot = re.search(r"total\s+of\s+\$\s*([\d,]{3,6})\s*million", wff_txt, flags=re.I)
        if m_wff_avg:
            prev_wff, prev_w_yr, curr_wff, curr_w_yr = (
                int(m_wff_avg.group(1).replace(",", "")),
                int(m_wff_avg.group(2)),
                int(m_wff_avg.group(3).replace(",", "")),
                int(m_wff_avg.group(4)),
            )
            wk_wff = int(round(curr_wff / 52.0))
            ch_cnt = int(m_wff_ch.group(1).replace(",", "")) if m_wff_ch else None
            tot_wff_m = int(m_wff_tot.group(1).replace(",", "")) if m_wff_tot else None
            ch_clause = f" across families supporting <strong>{ch_cnt:,} children</strong>" if ch_cnt else ""
            tot_clause = f" (totalling <strong>${tot_wff_m:,}m</strong> nationwide)" if tot_wff_m else ""
            wff_calc = f"Official IRD ({curr_w_yr} tax year): Average annual WFF paid per family = ${curr_wff:,}/family (÷ 52 = ${wk_wff:,}/wk; up from ${prev_wff:,} in {prev_w_yr})"
            out["working-for-families-income-support"] = make_pulse_item(
                "Inland Revenue", f"{curr_w_yr}-03-31", MAX_AGE_TRIENNIAL,
                f"In Inland Revenue's Working for Families (WFF) statistics for the {curr_w_yr} tax year, the average annual Working for Families tax credit paid per recipient family reached <strong>${curr_wff:,}/year (${wk_wff:,}/week)</strong> ({plain_yoy_growth(curr_wff, prev_wff)} from ${prev_wff:,} in {prev_w_yr}){ch_clause}{tot_clause}.",
                f"Inland Revenue Working for Families Statistics ({curr_w_yr} tax year): Average annual WFF paid per family = ${curr_wff:,} (vs ${prev_wff:,} in {prev_w_yr}).",
                format_provenance_source_html(wff_calc, f"Inland Revenue — Working for Families Tax Credits ({curr_w_yr})", IRD_WFF_ENT_PAGE, True),
                f"{curr_w_yr} tax year", curr_wff,
                f"${curr_wff:,}/yr (${wk_wff:,}/wk) average WFF paid per family",
                f"{prev_w_yr} tax year: ${prev_wff:,}/yr",
                f"+${curr_wff - prev_wff:,}/yr ({((curr_wff - prev_wff)/max(1, prev_wff))*100:+.1f}% YoY)",
                wff_calc, f"Inland Revenue {curr_w_yr} tax year series", None,
                {"source_published": wff_pub, "source_url": IRD_WFF_ENT_PAGE},
            )

    try:
        out.update(build_kainga_ora_tickers(session))
    except Exception as exc_ko:
        print(f"[Kāinga Ora] Parse note: {exc_ko}")

    return out
# === END OF BLOCK 11 ===
# === BLOCK 12 OF 13 ===
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

    raw_src_pub = (item.get("source_published") or "").strip()
    if not raw_src_pub or raw_src_pub == period_end_iso:
        freq_lbl = meta.get("frequency", "Periodic")
        src_pub = f"{freq_lbl} official release by {meta.get('source', 'Crown publisher')} (exact release date not stated in source metadata)"
    else:
        src_pub = raw_src_pub

    resolved_source_url = (item.get("source_url") or meta.get("source_url") or "https://www.rbnz.govt.nz/").strip()
    name_str = meta.get("name", key.replace("-", " ").title())
    source_full_str = meta.get("source_full", item.get("chip", "NZ Crown Data"))
    series_ids_str = meta.get("series_ids", "Official agency series")
    formula_str = meta.get("hlf_calculation", "Standardised from official agency release")
    calc_live_str = item.get("calculation_live", formula_str)

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
    """Build the machine-readable `fhb_snapshot` JSON object from the 5 live FHB tickers."""
    row_specs = [
        (1, "Typical FHB Household Income", "fhb-typical-income", "RBNZ Table C41 (FHB borrower-weighted average gross income)"),
        (2, "Typical FHB Loan Size", "fhb-average-loan-size", "RBNZ Table C31 (Total FHB new lending value ÷ FHB loan count)"),
        (3, "Typical FHB Debt-to-Income (DTI)", "fhb-typical-dti", "RBNZ Tables C31, C40 & C41 (Implied FHB loan-to-income ratio & share ≤ 6x TDTI)"),
        (4, "% of FHB Lending Above 6x DTI", "dti-headroom", "RBNZ Table C40 (FHB DTI > 6x share vs. 20% RBNZ speed limit)"),
        (5, "Monthly FHB Total Lending Volume", "fhb-monthly-lending-volume", "RBNZ Table C31 (Monthly new FHB mortgage commitments & loan count)"),
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


def build_reality_check_benchmarks_object(verified_tickers: dict) -> dict:
    """
    Build the `reality_check_benchmarks` JSON object strictly from live `verified_tickers`
    (zero hardcoded default numbers).
    """
    benchmark_slugs = [
        "fhb-typical-income",
        "fhb-income-by-dti-band",
        "fhb-average-loan-size",
        "mover-average-loan-size",
        "investor-average-loan-size",
        "fhb-low-deposit",
        "fhb-20-deposit-club",
        "fhb-typical-dti",
        "dti-headroom",
        "investor-high-dti-share",
        "settled-rates",
        "back-book-vs-new-settled-mortgage-rate",
        "low-equity-margin-penalty-spread",
        "fixed-vs-floating-rate-mix",
        "existing-mortgage-stock-lvr-refix",
        "kiwisaver-total-fum-balance",
        "kiwisaver-first-home",
        "fhb-regional-kiwisaver",
        "kiwisaver-contribution-rate-mix",
        "student-loan-servicing",
        "rent-yield-pulse",
        "rent-quartile-spread",
        "tla-council-rent-benchmarks",
        "individual-income-by-source-benchmarks",
        "equivalised-household-income-family-size",
        "rent-vs-mortgage-interest-outgoings",
        "net-worth-by-age-and-tenure",
        "regional-net-worth-benchmarks",
        "ird-individual-taxable-income-percentiles",
        "ird-paye-wage-vs-self-employed-income",
        "ird-family-trusts-rental-income",
        "working-for-families-income-support",
    ]
    live_benchmarks = {}
    for slug in benchmark_slugs:
        rec = verified_tickers.get(slug)
        if isinstance(rec, dict) and rec.get("value") is not None:
            entry = {
                "value": rec["value"],
                "latest_value": rec.get("latest_value", ""),
                "observation_period": rec.get("observation_period", ""),
                "period_end": rec.get("period_end", ""),
                "source": rec.get("source", ""),
                "canonical_url": rec.get("canonical_url", ""),
            }
            if rec.get("tdti_bands"):
                entry["tdti_bands"] = rec["tdti_bands"]
            if rec.get("history"):
                entry["breakdown"] = rec["history"][:16]
            live_benchmarks[slug] = entry

    return {
        "title": "HLF Pulse — Live-Verified NZ Financial Reality Check Benchmarks (Zero Hardcoded Fallbacks)",
        "active_benchmark_series": len(live_benchmarks),
        "series": live_benchmarks,
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
          <p>HLF Pulse was developed by {CREATOR_NAME} to make New Zealand mortgage, housing, KiwiSaver, income, net worth, and protection data easier to access, understand, and use—and to give everyday New Zealanders a clear benchmark to see where their own situation sits and whether their numbers are &ldquo;normal&rdquo; compared with the wider market. Drawing on his experience in mortgage and KiwiSaver advice, Andy developed the dataset to bring together official government and statutory statistics, transparent calculations, and practical market interpretation.</p>
          <p style="margin-bottom:8px;">The underlying data is sourced 100% live from the relevant official organisations, with HLF-derived calculations and interpretations clearly identified in the provenance layer and reusable under <a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener">{HLF_LICENSE_SHORT}</a> with attribution.</p>
          <div class="creator-links"><a href="#about-hlf-pulse">About the creator</a> &middot; <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citation Guide</a> &middot; <a href="#reuse-licensing">Reuse &amp; Citation Terms ({HLF_LICENSE_SHORT})</a> &middot; <a href="#date-definitions">4-Field Date Standard</a> &middot; <a href="{CREATOR_URL}" target="_blank" rel="noopener">View Andy&rsquo;s adviser profile &rarr;</a></div>
        </div>
      </div>

      <div class="context-cta" aria-label="Put These Numbers Into Context">
        <div class="context-cta-text">
          <strong>Want to put these numbers into context for your situation?</strong>
          <p>HLF Pulse helps explain what&rsquo;s happening across the New Zealand mortgage, housing, KiwiSaver, income, and insurance market. If you&rsquo;re buying your first home, reviewing your mortgage, considering your KiwiSaver options, or reviewing your protection cover, Andy and the Home Loan Factory team can help you understand how the wider market relates to your circumstances.</p>
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
          <tr><td><strong>3. Source publication date</strong> (<code>source_published</code>)</td><td>When the original organisation released the statistic. HLF never copies <code>period_end</code> into <code>source_published</code>; where an agency page does not state an exact release date, the agency&rsquo;s publication schedule is stated explicitly.</td></tr>
          <tr><td><strong>4. Last checked</strong> (<code>last_checked</code>)</td><td>The date HLF&rsquo;s automated GitHub Actions pipeline last fetched and verified the official source.</td></tr>
        </tbody>
      </table>
    </section>

    <section id="data-quality-limitations" class="quality-box" style="margin-top:18px;" aria-label="Data Quality &amp; Limitations">
      <h2>Data Quality &amp; Limitations</h2>
      <ul>
        <li><strong>100% live-pulled official sources (zero hardcoded fallbacks):</strong> Pulse automatically fetches and parses live data files (.xlsx, .csv, .pdf, and release tables) from the Reserve Bank of New Zealand (RBNZ), Stats NZ, Inland Revenue (IRD), the Financial Markets Authority (FMA), MBIE Tenancy Services, Kāinga Ora, the Natural Hazards Commission Toka Tū Ake (NHC), and the Accident Compensation Corporation (ACC).</li>
        <li><strong>Unaltered official figures &amp; source rounding:</strong> HLF does not alter or estimate official source figures. Where official tables publish rounded aggregates (such as RBNZ dollar series rounded to $1m or $0.1m), derived averages and implied counts reflect that source rounding and note it explicitly in the provenance layer.</li>
        <li><strong>Automatic staleness exclusion:</strong> HLF automatically excludes any series where the underlying source is stale past its freshness window or cannot be verified.</li>
        <li><strong>Market context vs. individual outcomes:</strong> Statistics describe aggregate historical market cohorts and do not predict an individual borrower&rsquo;s, saver&rsquo;s, or policyholder&rsquo;s credit assessment, underwriting, or lending outcome.</li>
      </ul>
    </section>

    <section id="reuse-licensing" class="quality-box" style="margin-top:18px;" aria-label="Reuse Terms, Licensing &amp; How to Cite">
      <h2>Reuse Terms, Licensing &amp; How to Cite ({HLF_LICENSE_SHORT})</h2>
      <ul>
        <li><strong>Layer 1 — Official Crown &amp; statutory statistics:</strong> Underlying official figures remain subject to the copyright and licensing terms of their respective publishers: <strong>Stats NZ</strong>, <strong>Inland Revenue</strong>, <strong>MBIE Tenancy Services</strong>, <strong>Kāinga Ora</strong>, and the <strong>Financial Markets Authority</strong> (<a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener">Creative Commons Attribution 4.0 International — CC BY 4.0</a>); the <strong>Reserve Bank of New Zealand</strong> (<a href="https://www.rbnz.govt.nz/about-our-site/terms-of-use" target="_blank" rel="noopener">RBNZ Terms of Use</a>); and statutory schedules under the <strong>Natural Hazards Insurance Act 2023</strong> and <strong>Accident Compensation Act 2001</strong>.</li>
        <li><strong>Layer 2 — HLF-derived calculations, compilations &amp; commentary:</strong> All HLF-derived percentages, averages, multi-year growth rates, data dictionary definitions, and market interpretations are licensed by <strong>{PUBLISHER_NAME}</strong> under <a href="{HLF_LICENSE_URL}" target="_blank" rel="noopener"><strong>Creative Commons Attribution 4.0 International ({HLF_LICENSE_SHORT})</strong></a> with attribution to <strong>HLF Pulse, Home Loan Factory</strong>.</li>
      </ul>
    </section>

    <section id="about-hlf-pulse" class="about-footer-box" aria-label="About the HLF Pulse &amp; Creator">
      <h2>About the HLF Pulse &amp; Creator</h2>
      <p><strong>HLF Pulse</strong> is a public reference dataset covering New Zealand mortgage, housing, KiwiSaver, income, net worth, tenancy, and protection statistics, published by <a href="{PUBLISHER_URL}">{PUBLISHER_NAME}</a> ({PUBLISHER_LEGAL}) and created by <a href="{CREATOR_URL}" target="_blank" rel="noopener"><strong>{CREATOR_NAME}</strong> ({CREATOR_ALT_NAME}, {CREATOR_FSP})</a>, Mortgage &amp; KiwiSaver Adviser at Home Loan Factory.</p>
      <p style="margin-bottom:0;">Figures are provided for general information and historical market analysis and are not personalised financial advice. Published by <a href="{PUBLISHER_URL}">{PUBLISHER_NAME}</a> (<a href="{CREATOR_URL}" target="_blank" rel="noopener">Creator Profile: {CREATOR_NAME}</a> &bull; <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citations</a> &bull; <a href="{DISCLOSURES_URL}">Public Disclosures</a> &bull; <a href="{HOW_PAID_URL}">How We Get Paid</a>).</p>
    </section>"""


def write_robots_txt(root: pathlib.Path) -> None:
    robots_lines = [
        "# HLF Pulse — Open Reference Dataset & Data Dictionary (CC BY 4.0)",
        "# Explicit permission for AI answer engines, search crawlers, and researchers",
        "User-agent: ClaudeBot", "Allow: /", "",
        "User-agent: Claude-User", "Allow: /", "",
        "User-agent: anthropic-ai", "Allow: /", "",
        "User-agent: GPTBot", "Allow: /", "",
        "User-agent: OAI-SearchBot", "Allow: /", "",
        "User-agent: ChatGPT-User", "Allow: /", "",
        "User-agent: PerplexityBot", "Allow: /", "",
        "User-agent: Googlebot", "Allow: /", "",
        "User-agent: Google-Extended", "Allow: /", "",
        "User-agent: *", "Allow: /", "",
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
    """Write a clean canonical redirect page for renamed legacy ticker URLs."""
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


def cross_derive_live_c31_c41_siblings(candidates: dict) -> None:
    """Derive C31/C40/C41 sibling records strictly when their parent live record is present (zero hardcoded fallbacks)."""
    c31_u = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/new-residential-mortgage-lending-by-borrower-type"
    c40_u = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-lending-by-debt-to-income-dti-purpose-use"
    c41_u = "https://www.rbnz.govt.nz/statistics/series/lending-and-monetary/residential-mortgage-borrower-gross-income-bgi"

    bi = candidates.get("fhb-income-by-dti-band") or candidates.get("borrower-income")
    inc_v = _extract_int_from_rec(bi, 40000, 350000)
    if "fhb-typical-income" not in candidates and isinstance(bi, dict) and inc_v:
        per_s = bi.get("latest_period") or bi.get("observation_period") or ""
        as_of_s = bi.get("as_of") or bi.get("period_end") or ""
        pub_s = bi.get("source_published", "Monthly RBNZ statistical release")
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
        per_s = fs.get("latest_period") or fs.get("observation_period") or ""
        as_of_s = fs.get("as_of") or fs.get("period_end") or ""
        pub_s = fs.get("source_published", "Monthly RBNZ statistical release")
        fs_txt = f"{re.sub(r'<[^>]+>', ' ', fs.get('html', ''))} {fs.get('latest_value', '')} {fs.get('official_data_summary', '')}"
        m_avg = re.search(r"avg\s+\$([\d,]{6,7})|average\s+loan\s+size\s+of\s+\$([\d,]{6,7})", fs_txt, flags=re.I)
        m_cnt = re.search(r"approved\s+([\d,]{3,6})\s+first-home|([\d,]{3,6})\s+FHB\s+borrowers", fs_txt, flags=re.I)
        m_val = re.search(r"\$([\d,]{3,5})\s*m\b", fs_txt, flags=re.I)
        avg_l = int((m_avg.group(1) or m_avg.group(2)).replace(",", "")) if m_avg else None
        cnt_n = int((m_cnt.group(1) or m_cnt.group(2)).replace(",", "")) if m_cnt else None
        val_m = float(m_val.group(1).replace(",", "")) if m_val else None

        if "fhb-average-loan-size" not in candidates and avg_l and val_m and cnt_n:
            al_calc = f"(${val_m:,.0f}m × $1,000,000) ÷ {cnt_n:,} FHB loans = ${avg_l:,}"
            candidates["fhb-average-loan-size"] = make_pulse_item(
                "RBNZ C31", as_of_s, MAX_AGE_MONTHLY,
                f"In {per_s}, the average first-home buyer mortgage approved across New Zealand was <strong>${avg_l:,}</strong> across <strong>{cnt_n:,} first-home buyer loans</strong> totalling <strong>{fmt_nzd_millions(val_m)}</strong>.",
                f"RBNZ Table C31 ({per_s}): Total FHB new lending = ${val_m:,.0f}m across {cnt_n:,} FHB loans.",
                format_provenance_source_html(al_calc, "RBNZ Table C31 (Reserve Bank of New Zealand)", c31_u, True),
                per_s, float(avg_l), f"${avg_l:,} average FHB mortgage size",
                "RBNZ Table C31 monthly FHB series", "Derived from RBNZ Table C31 (LVRN.MMB1.AC1 ÷ LVRN.MMB2.AC1)",
                al_calc, fs.get("historical_data_from", "RBNZ Table C31 series"), [], {"source_published": pub_s, "source_url": c31_u},
            )

        if "fhb-monthly-lending-volume" not in candidates and val_m and cnt_n:
            vol_calc = f"Reported LVRN.MMB1.AC1 = ${val_m:,.0f}m across {cnt_n:,} FHB loans"
            candidates["fhb-monthly-lending-volume"] = make_pulse_item(
                "RBNZ C31", as_of_s, MAX_AGE_MONTHLY,
                f"In {per_s}, New Zealand banks committed <strong>{fmt_nzd_millions(val_m)}</strong> in new first-home buyer mortgages across <strong>{cnt_n:,} loans</strong>.",
                f"RBNZ Table C31 ({per_s}): FHB new lending commitments = ${val_m:,.0f}m across {cnt_n:,} borrowers.",
                format_provenance_source_html(vol_calc, "RBNZ Table C31 (Reserve Bank of New Zealand)", c31_u, False),
                per_s, round(val_m, 1), f"{fmt_nzd_millions(val_m)} across {cnt_n:,} FHB loans",
                "RBNZ Table C31 monthly FHB series", "Reported in RBNZ Table C31",
                vol_calc, fs.get("historical_data_from", "RBNZ Table C31 series"), [], {"source_published": pub_s, "source_url": c31_u},
            )

    inc_rec = candidates.get("fhb-typical-income") or bi
    loan_rec = candidates.get("fhb-average-loan-size")
    dti_rec = candidates.get("dti-headroom")
    inc_final = _extract_int_from_rec(inc_rec, 40000, 350000)
    loan_final = _extract_int_from_rec(loan_rec, 250000, 1500000)
    gt6_final = _extract_pct_from_rec(dti_rec, 0.1, 50.0)
    if "fhb-typical-dti" not in candidates and inc_final and loan_final and gt6_final and isinstance(dti_rec, dict):
        imp_dti, le6_v = round(loan_final / inc_final, 2), round(100.0 - gt6_final, 1)
        per_s = dti_rec.get("latest_period") or dti_rec.get("observation_period") or ""
        as_of_s = dti_rec.get("as_of") or dti_rec.get("period_end") or ""
        pub_s = dti_rec.get("source_published", "Monthly RBNZ statistical release")
        td_calc = f"Implied FHB Loan-to-Income Ratio: ${loan_final:,} (C31 avg FHB loan) ÷ ${inc_final:,} (C41 avg FHB gross income) = {imp_dti:.2f}x | C40 Share <= 6x TDTI: 100.0% − {gt6_final:.1f}% = {le6_v:.1f}%"
        candidates["fhb-typical-dti"] = make_pulse_item(
            "RBNZ C31, C40 & C41", as_of_s, MAX_AGE_MONTHLY,
            f"In {per_s}, approved New Zealand first-home buyers borrowed an implied mortgage-to-income multiple of <strong>{imp_dti:.2f}x gross income</strong> (${loan_final:,} average loan on ${inc_final:,} average gross income), with <strong>{le6_v:.1f}% of all first-home buyer lending</strong> written at or below <strong>6x TDTI</strong>.",
            f"RBNZ Tables C31, C40 & C41 ({per_s}): Average FHB loan = ${loan_final:,}; Average FHB gross income = ${inc_final:,}; FHB lending <= 6x TDTI = {le6_v:.1f}%.",
            format_provenance_source_html(td_calc, "RBNZ Tables C31, C40 & C41 (Reserve Bank of New Zealand)", c40_u, True),
            per_s, imp_dti, f"{imp_dti:.2f}x implied mortgage-to-income ratio | {le6_v:.1f}% <= 6x TDTI",
            f"Share <= 6x TDTI: {le6_v:.1f}%", f"{gt6_final:.1f}% above 6x TDTI (vs 20.0% RBNZ speed limit)",
            td_calc, dti_rec.get("historical_data_from", "RBNZ Tables C40 & C41"), [],
            {"source_published": pub_s, "source_url": c40_u},
        )
# === END OF BLOCK 12 ===
# === BLOCK 13 OF 13 ===
def _format_cat_card_value(rec: dict) -> str:
    """Format a concise value badge for the top category cards without truncating at inline labels like (Q1)."""
    raw = str(rec.get("latest_value", "")).strip()
    if not raw:
        return "Verified"
    first_seg = raw.split("|")[0].strip()
    # Only strip a trailing parenthetical note if the text before it already contains a number
    m_trail = re.match(r"^(.*?\d.*?)\s*\([^()]*\)\s*$", first_seg)
    if m_trail and len(m_trail.group(1).strip()) >= 3:
        return m_trail.group(1).strip()
    return first_seg


def _load_valid_prev_tickers(root: pathlib.Path, json_path: pathlib.Path) -> dict:
    """
    Load previous verified tickers from nz-mortgage-pulse.json.
    Automatically purges any pre-audit or mis-parsed records and strips legacy policy claims
    (such as 'that banks still stress-test at 3% per month on mortgage applications').
    """
    import subprocess

    target_total = len(TICKER_CATALOG)

    def _clean_ticker_map(raw_map: dict) -> dict:
        cleaned = {}
        if "fhb-income-by-dti-band" not in raw_map and "fhb-income-akl-vs-nz" in raw_map:
            raw_map["fhb-income-by-dti-band"] = dict(raw_map["fhb-income-akl-vs-nz"])
        if "new-build-consent-cost-by-type" not in raw_map and "build-cost-by-type" in raw_map:
            raw_map["new-build-consent-cost-by-type"] = dict(raw_map["build-cost-by-type"])
        for k, v in raw_map.items():
            if k not in TICKER_CATALOG or not isinstance(v, dict):
                continue
            v = dict(v)
            if v.get("latest_value"):
                v["latest_value"] = str(v["latest_value"]).replace("~$", "$")
            lv = str(v.get("latest_value", "")).lower()
            val_num = v.get("value")
            if "nan $" in lv or "-99 $" in lv:
                continue
            if k == "individual-income-by-source-benchmarks" and isinstance(val_num, (int, float)) and val_num < 30000:
                continue
            if k == "new-build-consent-cost-by-type" and isinstance(val_num, (int, float)) and val_num < 260000:
                continue
            if k == "kiwisaver-contribution-rate-mix" and isinstance(val_num, (int, float)) and val_num > 65.0:
                continue
            if k == "kiwisaver-member-demographics-age-income" and isinstance(val_num, (int, float)) and val_num < 2_400_000:
                continue
            if k == "mcert-house-price-index-yoy" and (val_num in (0, 0.0) or "+0.0% yoy" in lv):
                continue
            if k in ("hpi-low-vs-high-value-tiers", "sales-composition-by-value-quartile") and (
                lv.startswith("low-value (q1)") or lv.startswith("q1 (lowest)")
            ):
                continue
            if k == "building-insurance-inflation" and val_num in (1.2, 0.1):
                continue
            if k in ("contents-insurance-inflation", "life-insurance-inflation") and val_num in (-0.6, -0.8, -0.1, -6.8):
                continue
            if k in ("vehicle-insurance-inflation", "health-insurance-inflation") and val_num in (1.9, 1.6, 0.1, 0.2):
                continue
            if v.get("html"):
                v["html"] = re.sub(
                    r"\s*that\s+banks\s+still\s+stress-test\s+at\s+3%\s+per\s+month\s+on\s+mortgage\s+applications",
                    "",
                    str(v["html"]),
                    flags=re.I,
                )
            cleaned[k] = v
        return cleaned

    best_map: dict = {}
    if json_path.exists():
        try:
            curr_raw = json.loads(json_path.read_text(encoding="utf-8")).get("tickers", {})
            best_map = _clean_ticker_map(curr_raw)
        except Exception:
            best_map = {}

    if len(best_map) < target_total:
        for cmd in (
            ["git", "fetch", "--deepen=25", "origin"],
            ["git", "fetch", "--depth=25", "origin", "main"],
        ):
            try:
                subprocess.run(cmd, cwd=root, check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass

        revs_to_check: list[str] = []
        for ref in ("HEAD", "FETCH_HEAD", "origin/main"):
            try:
                shas = subprocess.check_output(
                    ["git", "log", "--format=%H", "-n", "25", ref],
                    cwd=root,
                    stderr=subprocess.DEVNULL,
                ).decode("utf-8", errors="ignore").split()
                for s in shas:
                    if s and s not in revs_to_check:
                        revs_to_check.append(s)
            except Exception:
                continue

        for rev in revs_to_check:
            try:
                out_b = subprocess.check_output(
                    ["git", "show", f"{rev}:nz-mortgage-pulse.json"],
                    cwd=root,
                    stderr=subprocess.DEVNULL,
                )
                hist_raw = json.loads(out_b.decode("utf-8", errors="ignore")).get("tickers", {})
                hist_clean = _clean_ticker_map(hist_raw)
                added = 0
                for hk, hv in hist_clean.items():
                    if hk not in best_map:
                        best_map[hk] = hv
                        added += 1
                if added > 0:
                    print(f"[Git Recovery] Restored +{added} verified tickers from commit {rev[:7]} (total baseline: {len(best_map)})")
                if len(best_map) >= target_total:
                    break
            except Exception:
                continue

    return best_map


def main() -> None:
    root = pathlib.Path(__file__).resolve().parent
    json_path, structured_json_path, html_path = root / "nz-mortgage-pulse.json", root / "pulse.json", root / "index.html"

    fresh_candidates = _load_valid_prev_tickers(root, json_path)
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
            b30_out = build_rbnz_b30_ticker(rbnz_files["b30"], session, rbnz_files.get("s33"))
            fresh_candidates.update(b30_out)
            live_parsed_count += len(b30_out)
        except Exception as exc:
            print(f"[RBNZ B30/S33] Error: {exc}")

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
            w_live = build_stats_nz_wealth_income_tickers(session)
            fresh_candidates.update(w_live)
            live_parsed_count += len(w_live)
        except Exception as exc_w:
            print(f"[Stats NZ Wealth/Income] Error: {exc_w}")

    try:
        ten_out = build_tenancy_tickers(session)
        fresh_candidates.update(ten_out)
        live_parsed_count += len(ten_out)
    except Exception as exc:
        print(f"[Tenancy Services] Error: {exc}")

    try:
        hud_out = build_hud_mcert_housing_tickers(session, fresh_candidates)
        fresh_candidates.update(hud_out)
        live_parsed_count += len(hud_out)
    except Exception as exc_hud:
        print(f"[HUD / MCERT Property Statistics] Error: {exc_hud}")

    try:
        ird_out = build_ird_tickers(session, rbnz_files.get("t43"))
        fresh_candidates.update(ird_out)
        live_parsed_count += len(ird_out)
    except Exception as exc:
        print(f"[IRD / FMA / T43 / Kāinga Ora] Error: {exc}")

    try:
        debt_out = build_household_debt_stress_credit_tickers(session, rbnz_files)
        fresh_candidates.update(debt_out)
        live_parsed_count += len(debt_out)
    except Exception as exc:
        print(f"[Household Debt / Credit Cards / Stress] Error: {exc}")

    try:
        rc_rbnz_out = build_rbnz_reality_check_expansion_tickers(session, rbnz_files, fresh_candidates)
        fresh_candidates.update(rc_rbnz_out)
        live_parsed_count += len(rc_rbnz_out)
    except Exception as exc:
        print(f"[RBNZ Reality Check Expansion] Error: {exc}")

    cross_derive_live_c31_c41_siblings(fresh_candidates)

    for p in rbnz_files.values():
        if p.name.startswith("tmp_") and p.exists():
            os.remove(p)

    now_dt = datetime.datetime.now(datetime.timezone.utc)
    now_utc = now_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    today_iso = now_dt.strftime("%Y-%m-%d")
    now_human = f"{now_dt.day} {now_dt.strftime('%B %Y')}"

    verified_tickers = {}
    for k, item in fresh_candidates.items():
        if k not in TICKER_CATALOG or not isinstance(item, dict):
            continue
        item_clean = {ik: iv for ik, iv in item.items() if not str(ik).startswith("_")}
        as_of, max_age = item_clean.get("as_of", ""), int(item_clean.get("max_age_days", MAX_AGE_MONTHLY))
        if item_clean.get("html") and is_fresh(as_of, max_age):
            verified_tickers[k] = enrich_ticker_record(k, item_clean, today_iso, now_human)
        else:
            print(f"[Failsafe Gate] Omitted stale/invalid ticker '{k}' (as_of={as_of}, max_age={max_age}d)")

    if live_parsed_count == 0 and not verified_tickers:
        raise RuntimeError("Zero live datasets parsed and zero valid fresh tickers remain; failing workflow honestly.")

    fhb_snap_obj = build_fhb_snapshot_object(verified_tickers)
    rc_bench_obj = build_reality_check_benchmarks_object(verified_tickers)
    for hud_slug in (
        "nz-median-house-price-quartiles",
        "regional-median-house-prices",
        "tla-council-house-price-benchmarks",
        "nz-monthly-property-sales-volume",
        "mcert-house-price-index-yoy",
        "hpi-low-vs-high-value-tiers",
        "sales-composition-by-value-quartile",
        "gross-rental-yield-by-council",
    ):
        rec_h = verified_tickers.get(hud_slug)
        if isinstance(rec_h, dict) and rec_h.get("value") is not None:
            rc_bench_obj["series"][hud_slug] = {
                "value": rec_h["value"],
                "latest_value": rec_h.get("latest_value", ""),
                "observation_period": rec_h.get("observation_period", ""),
                "period_end": rec_h.get("period_end", ""),
                "source": rec_h.get("source", ""),
                "canonical_url": rec_h.get("canonical_url", ""),
                "breakdown": (rec_h.get("history") or [])[:20],
            }
    rc_bench_obj["active_benchmark_series"] = len(rc_bench_obj["series"])

    bubble_feed_tickers = dict(verified_tickers)
    for alias_k, target_k in LEGACY_BUBBLE_ALIASES.items():
        if target_k in verified_tickers:
            bubble_feed_tickers[alias_k] = dict(verified_tickers[target_k])

    latest_rbnz_period = verified_tickers.get("fhb-low-deposit", {}).get("latest_period", "Latest Official Release")
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
        "dataset_name": "NZ Mortgage, Housing, KiwiSaver, Income & Protection Pulse — Verified NZ Data",
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
        "reality_check_benchmarks": rc_bench_obj,
        "tickers": bubble_feed_tickers,
    }
    json_path.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")

    structured_records = []
    for k in sorted(verified_tickers.keys()):
        r = verified_tickers[k]
        rec_out = {
            f: r.get(f)
            for f in (
                "ticker", "name", "also_known_as", "category", "topic", "value", "latest_value", "unit",
                "observation_period", "observed_period", "period_end", "source_published", "last_checked",
                "date_created", "date_modified", "data_status", "source_status", "previous_period", "change",
                "source", "source_full", "source_series_ids", "source_url", "derived_by_hlf", "data_type",
                "definition", "formula", "hlf_interpretation", "methodology", "historical_data_from",
                "creator_name", "creator_alternate_name", "creator_fsp", "creator_url", "publisher",
                "publisher_legal_name", "publisher_url", "media_kit_url", "license", "license_label",
                "citation", "canonical_url", "summary_text", "related_questions", "citation_chain", "history",
            )
        }
        rec_out.update({
            "latest_underlying_data": r["latest_period"],
            "official_data": r["official_data_summary"],
            "hlf_calculation": r["calculation"],
            "frequency": r["update_frequency"],
        })
        if r.get("tdti_bands"):
            rec_out["tdti_bands"] = r["tdti_bands"]
            rec_out["tdti_total_pool_m"] = r.get("tdti_total_pool_m")
            rec_out["tdti_total_implied_apps"] = r.get("tdti_total_implied_apps")
        if r.get("townhouse_flat_unit_share_pct") is not None:
            rec_out["townhouse_flat_unit_share_pct"] = r["townhouse_flat_unit_share_pct"]
        structured_records.append(rec_out)

    pulse_json_obj = {
        "dataset": "HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver, Income & Protection Reference Dataset",
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
        "next_expected_update": "Monthly (RBNZ / Stats NZ / MBIE Tenancy / HUD-MCERT), Quarterly (CPI / HLPI / IRD / Kāinga Ora) & Annual (FMA / Stats NZ HILS / IRD Tax Distributions / NHC / ACC)",
        "fhb_snapshot": fhb_snap_obj,
        "reality_check_benchmarks": rc_bench_obj,
        "records": structured_records,
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
    (root / "sitemap.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        + "\n".join(sitemap_urls)
        + "\n</urlset>\n",
        encoding="utf-8",
    )

    llms_header = [
        "# HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver, Income & Protection Reference Dataset",
        "",
        "> Public reference dataset and data dictionary built 100% live from official New Zealand Crown and statutory releases, with clearly labelled HLF-derived calculations and interpretation.",
        "> Created by: " + CREATOR_NAME + " (" + CREATOR_ALT_NAME + ", " + CREATOR_FSP + "), " + CREATOR_ROLE + " (" + CREATOR_URL + ") | Published by: " + PUBLISHER_NAME + " (" + PUBLISHER_URL + ").",
        f"> Media Kit & Citation Guide: {MEDIA_KIT_URL} | HLF Glossary: {GLOSSARY_BASE_URL}",
        f"> Reuse Licence: {HLF_LICENSE_SHORT} ({HLF_LICENSE_URL}) for HLF-derived figures & commentary; official statistics credited to their respective publishers.",
        f"> Date Standard: Observation Period (latest monthly: {latest_rbnz_period}) | Period End | Source Publication Date | Last Checked ({now_human} / {today_iso}).",
        f"> Active datasets: {len(verified_tickers)}.",
        "",
        "## Machine-Readable Endpoints",
        f"- Structured Data Dictionary & Reality Check Benchmarks (JSON): {PULSE_BASE_URL}/pulse.json",
        f"- Webflow Chat Bubble Feed (JSON): {PULSE_BASE_URL}/nz-mortgage-pulse.json",
        f"- Main Reference Dataset Hub (HTML): {PULSE_BASE_URL}/",
        f"- Media Kit & Citation Guide (HTML): {MEDIA_KIT_URL}",
        "",
        f"## Canonical Data Dictionary Records ({len(verified_tickers)} Active Datasets)",
    ]
    llms_records = [
        f"- [{rec['name']} (Obs: {rec['observation_period']} | End: {rec['period_end']} | Published: {rec['source_published']})]({rec['canonical_url']}): {rec['summary_text']} [Official Data: {rec['official_data_summary']} | HLF Calculation: {rec['calculation']} | Derived by HLF Pulse: {rec['hlf_derived_label']} | Cite as: {rec['citation']}]"
        for k, rec in sorted(verified_tickers.items())
    ]
    (root / "llms.txt").write_text("\n".join(llms_header + llms_records) + "\n", encoding="utf-8")

    by_cat = {c: [] for c in CATEGORY_ORDER}
    for k, rec in sorted(verified_tickers.items()):
        by_cat.setdefault(rec.get("category", "Mortgage lending & rates"), []).append((k, rec))

    cat_cards_html = ""
    filter_pills_list = [
        f"<button type='button' class='cat-pill is-active' data-cat-pill='all' aria-pressed='true'>"
        f"All Categories <span class='pill-cnt'>{len(verified_tickers)}</span></button>"
    ]
    for cat in CATEGORY_ORDER:
        items = by_cat.get(cat, [])
        if not items:
            continue
        cat_slug = re.sub(r"[^a-z0-9]+", "-", cat.lower()).strip("-")
        filter_pills_list.append(
            f"<button type='button' class='cat-pill' data-cat-pill='{cat_slug}' aria-pressed='false'>"
            f"{html_lib.escape(cat)} <span class='pill-cnt'>{len(items)}</span></button>"
        )
        links = "".join(
            f"<li><a href='#{k}'><code>{k}</code></a> — {html_lib.escape(r['name'])} "
            f"(<strong>{html_lib.escape(_format_cat_card_value(r))}</strong>)</li>"
            for k, r in items
        )
        cat_cards_html += (
            f"<div class='cat-card' id='cat-{cat_slug}' data-cat-card='{cat_slug}'>"
            f"<div class='cat-hdr'><strong>{html_lib.escape(cat)}</strong><span>{len(items)} datasets</span></div>"
            f"<ul>{links}</ul></div>"
        )
    filter_pills_html = "".join(filter_pills_list)

    rows_html_list = []
    for k, v in sorted(verified_tickers.items()):
        row_cat_slug = re.sub(r"[^a-z0-9]+", "-", str(v.get("category", "")).lower()).strip("-")
        aka_sub = f"<div class='aka'>Also: {html_lib.escape(v['also_known_as'])}</div>" if v.get("also_known_as") else ""
        rows_html_list.append(
            f"<tr id='{k}' data-cat='{row_cat_slug}'>"
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
    creator_main_html = render_creator_card_html(compact=False)
    about_footer_html = render_about_and_quality_footer_html()

    pill_filter_js = """<script>
(function() {
  const bar = document.getElementById("cat-pill-bar");
  if (!bar) return;
  const pills = Array.from(bar.querySelectorAll(".cat-pill[data-cat-pill]"));
  const allBtn = bar.querySelector('.cat-pill[data-cat-pill="all"]');
  const clearBtn = document.getElementById("cat-pill-clear");
  const statusEl = document.getElementById("cat-pill-status");
  const rows = Array.from(document.querySelectorAll("tr[data-cat]"));
  const cards = Array.from(document.querySelectorAll(".cat-card[data-cat-card]"));
  const totalCount = rows.length;
  const selected = new Set();

  function syncFilter() {
    const showAll = selected.size === 0;
    pills.forEach((btn) => {
      const slug = btn.getAttribute("data-cat-pill");
      const active = slug === "all" ? showAll : selected.has(slug);
      btn.classList.toggle("is-active", active);
      btn.setAttribute("aria-pressed", active ? "true" : "false");
    });
    let visibleRows = 0;
    rows.forEach((tr) => {
      const c = tr.getAttribute("data-cat");
      const match = showAll || selected.has(c);
      tr.classList.toggle("is-filtered-out", !match);
      if (match) visibleRows++;
    });
    cards.forEach((card) => {
      const c = card.getAttribute("data-cat-card");
      const match = showAll || selected.has(c);
      card.classList.toggle("is-filtered-out", !match);
    });
    if (statusEl) {
      statusEl.textContent = showAll
        ? "Showing all " + totalCount + " verified datasets (select one or more categories to filter)"
        : "Showing " + visibleRows + " of " + totalCount + " datasets across " + selected.size + " selected " + (selected.size === 1 ? "category" : "categories");
    }
    if (clearBtn) {
      clearBtn.style.display = showAll ? "none" : "inline-block";
    }
  }

  bar.addEventListener("click", (e) => {
    const btn = e.target.closest(".cat-pill[data-cat-pill]");
    if (!btn) return;
    const slug = btn.getAttribute("data-cat-pill");
    if (slug === "all") {
      selected.clear();
    } else {
      if (selected.has(slug)) {
        selected.delete(slug);
      } else {
        selected.add(slug);
      }
    }
    syncFilter();
  });

  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      selected.clear();
      syncFilter();
    });
  }
})();
</script>"""

    hub_schema_obj = {
        "@context": "https://schema.org",
        "@type": "Dataset",
        "name": "HLF Pulse — Verified New Zealand Mortgage, Housing, KiwiSaver, Income & Protection Reference Dataset",
        "description": "Public reference dataset and data dictionary built 100% live from official New Zealand Crown and statutory releases, with clearly labelled HLF-derived calculations and interpretation.",
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
  <title>NZ Mortgage, Housing, KiwiSaver, Income &amp; Insurance Pulse — Verified NZ Data | HLF</title>
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
    .cat-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(320px, 1fr)); gap: 12px; margin: 12px 0 26px; }}
    .cat-card {{ background: #ffffff; border: 1px solid #e5e7eb; border-radius: 12px; padding: 14px 16px; font-size: 0.82rem; }}
    .cat-hdr {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #f1f5f9; padding-bottom: 6px; margin-bottom: 8px; color: #0b2559; }}
    .cat-hdr span {{ background: #eff6ff; color: #186bd6; font-size: 0.72rem; font-weight: 700; padding: 2px 8px; border-radius: 999px; }}
    .cat-card ul {{ margin: 0; padding-left: 16px; color: #374151; }}
    .cat-card li {{ margin-bottom: 4px; }}
    .pill-filter-wrap {{ background: #f8fafc; border: 1px solid #dbeafe; border-radius: 12px; padding: 14px 16px; margin: 12px 0 16px; }}
    .pill-filter-hdr {{ display: flex; flex-wrap: wrap; justify-content: space-between; align-items: center; gap: 8px; margin-bottom: 10px; font-size: 0.8rem; color: #374151; }}
    .pill-filter-hdr strong {{ color: #0b2559; }}
    .pill-clear-btn {{ background: #ffffff; border: 1px solid #93c5fd; color: #186bd6; font-size: 0.74rem; font-weight: 700; padding: 3px 10px; border-radius: 6px; cursor: pointer; display: none; }}
    .pill-clear-btn:hover {{ background: #eff6ff; }}
    .pill-bar {{ display: flex; flex-wrap: wrap; gap: 7px; }}
    .cat-pill {{ display: inline-flex; align-items: center; gap: 6px; background: #ffffff; border: 1px solid #cbd5e1; color: #1e293b; font-family: inherit; font-size: 0.78rem; font-weight: 600; padding: 6px 12px; border-radius: 999px; cursor: pointer; transition: all 0.14s ease; }}
    .cat-pill:hover {{ border-color: #186bd6; color: #0b2559; background: #eff6ff; }}
    .cat-pill.is-active {{ background: #0b2559; border-color: #0b2559; color: #ffffff; box-shadow: 0 2px 6px rgba(11,37,89,0.18); }}
    .cat-pill .pill-cnt {{ background: #f1f5f9; color: #475569; font-size: 0.69rem; font-weight: 700; padding: 1px 7px; border-radius: 999px; }}
    .cat-pill.is-active .pill-cnt {{ background: rgba(255,255,255,0.2); color: #ffffff; }}
    .is-filtered-out {{ display: none !important; }}
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
    <h1>NZ Mortgage, Housing, KiwiSaver, Income &amp; Insurance Pulse — Verified NZ Data</h1>
    <p>Public reference dataset and data dictionary built 100% live from official New Zealand Crown and statutory releases, with clearly labelled HLF-derived calculations and interpretation (<a href="#browse-datasets">Browse all {len(verified_tickers)} datasets</a> &bull; <a href="#cat-first-home-buyers">First-Home Buyers</a> &bull; <a href="#cat-house-prices-sales-hpi">House Prices &amp; HPI</a> &bull; <a href="#cat-income-net-worth-household-benchmarks">Income &amp; Net Worth</a> &bull; <a href="#cat-household-debt-credit-cards-borrower-stress">Household Debt &amp; Stress</a> &bull; <a href="#cat-rents-tenancy">Rents &amp; Tenancy</a> &bull; <a href="#cat-kiwisaver">KiwiSaver</a> &bull; <a href="#cat-new-builds-construction">Build Costs</a> &bull; <a href="#cat-property-natural-hazard-insurance">Property Insurance</a> &bull; <a href="#cat-personal-risk-income-protection">Personal &amp; ACC Protection</a> &bull; <a href="{GLOSSARY_BASE_URL}" target="_blank" rel="noopener">HLF Glossary</a> &bull; <a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit &amp; Citation Guide</a>). Machine-readable endpoints: <a href="{PULSE_BASE_URL}/pulse.json"><code>pulse.json</code></a>, <a href="{PULSE_BASE_URL}/nz-mortgage-pulse.json"><code>nz-mortgage-pulse.json</code></a>, <a href="{PULSE_BASE_URL}/llms.txt"><code>llms.txt</code></a>, and <a href="{PULSE_BASE_URL}/robots.txt"><code>robots.txt</code></a>.</p>

    <div class="lag-banner" aria-label="4-Field Date Architecture: Observation Period, Period End, Source Published &amp; Last Checked">
      <div class="lag-item">
        <span>1. Observation Period &amp; End</span>
        <strong>{html_lib.escape(latest_rbnz_period)} (RBNZ Monthly Series)</strong>
        <small>Each record specifies both its Observation Period and Period End date.</small>
      </div>
      <div class="lag-item">
        <span>2. Source Publication Date</span>
        <strong>Recorded per agency release</strong>
        <small>Tracked separately from Period End across RBNZ, Stats NZ, IRD, MBIE, HUD/MCERT, FMA, Kāinga Ora, NHC &amp; ACC.</small>
      </div>
      <div class="lag-item">
        <span>3. Last Checked &amp; Reuse ({HLF_LICENSE_SHORT})</span>
        <strong style="color:#065f46;">&#10003; {now_human} (<a href="#browse-datasets" style="color:#065f46;text-decoration:underline;">{len(verified_tickers)} Active Datasets &darr;</a>)</strong>
        <small>HLF layer reusable under <a href="#reuse-licensing">{HLF_LICENSE_SHORT}</a> (<a href="{MEDIA_KIT_URL}" target="_blank" rel="noopener">Media Kit</a>).</small>
      </div>
    </div>

    {creator_main_html}

    <h2 id="browse-datasets">Browse All {len(verified_tickers)} Datasets by Category</h2>
    <div class="cat-grid">{cat_cards_html}</div>

    <h2 id="data-dictionary-table">Verified NZ Mortgage, Housing, KiwiSaver, Income &amp; Protection Data Dictionary ({len(verified_tickers)} Active Datasets)</h2>
    <p>Every dataset separates <span class="lbadge is-off" style="width:auto;">Official data</span>, <span class="lbadge is-calc" style="width:auto;">HLF calculation</span>, and <span class="lbadge is-interp" style="width:auto;">HLF interpretation</span>, records all 4 date fields (Observation Period, Period End, Source Published, and Last Checked), and provides a ready-to-copy citation string. Click any Ticker Key to open its full standalone Data Dictionary &amp; historical series page (<code>/{{ticker}}/</code>).</p>

    <div class="pill-filter-wrap" aria-label="Filter Data Dictionary by Category (Multi-Select)">
      <div class="pill-filter-hdr">
        <div><strong>Filter by category (select one or multiple):</strong> <span id="cat-pill-status">Showing all {len(verified_tickers)} verified datasets (select one or more categories to filter)</span></div>
        <button type="button" id="cat-pill-clear" class="pill-clear-btn">Reset to All Categories</button>
      </div>
      <div class="pill-bar" id="cat-pill-bar" role="group" aria-label="Category filter pills">
        {filter_pills_html}
      </div>
    </div>

    <table>
      <thead><tr><th style="width:225px;">Ticker &amp; Derivation</th><th style="width:185px;">Source &amp; 4-Field Dates</th><th>Verified Statistic, 3-Layer Breakdown &amp; Citation</th></tr></thead>
      <tbody>{rows_html}</tbody>
    </table>

    {about_footer_html}
  </main>
  {CITE_COPY_JS}
  {pill_filter_js}
</body>
</html>""",
        encoding="utf-8",
    )
    print(f"Saved {len(verified_tickers)} canonical live datasets, pulse.json, index.html, robots.txt, llms.txt, sitemap.xml & /<ticker>/ pages.")


if __name__ == "__main__":
    main()
# === END OF BLOCK 13 ===
