"""
formats.py
Centralized configuration module for Upload Format schemas (Quantity vs Amount).
Defines program technical value classifications, header/sub-header schemas, and utility functions.
"""

import pandas as pd

PROGRAM_FORMAT_TYPE = {
    "M7_TP_VOL01": "Amount",
    "M7_TP_VOL05": "Amount",
    "M7_TP_VOL03": "Amount",
    "M7_TP_VOL07": "Amount",
    "M7_TP_VOL02": "Quantity",
    "M7_TP_VOL06": "Quantity",
    "M7_TP_VOL04": "Quantity",
    "M7_TP_VOL08": "Quantity",
}

# Sub-headings per section tag for Quantity format
QUANTITY_HEADERS = {
    "HEADER": ["Request Description"],
    "AGRMT": [
        "Master Agreement External ID", "Master Agreement Description", "Company",
        "Business Unit", "Legal Valid From", "Legal Valid To"
    ],
    "SUBAGRMT": [
        "Processing Sequence", "Sub Agreement External ID", "Sub Agreement Description",
        "Sub Agreement Program", "Period Profile", "Accrual Frequency", "Settlement Frequency",
        "GST Indicator", "Key for Prov G/L", "Legal Valid From", "Legal Valid To"
    ],
    "ELCUST": [
        "Sub-Agreement External ID", "Customer External ID", "Flexible Group",
        "Valid From", "Valid To"
    ],
    "FGHD": [
        "Description", "Flexible Group Type", "Flexible Group", "Flexible Group Request"
    ],
    "FGIT": [
        "Flexible Group Category", "Include/Exclude", "Material External ID", "Material Group",
        "Capacity", "Series", "Star Rating", "Star Rating Year", "Feature 1", "Feature 2",
        "Feature 3", "Feature 4", "Feature 5", "Feature 6", "Feature 7", "Feature 8",
        "Feature 9", "Feature 10", "Customer Group", "Region", "State", "Territory/Sales Area",
        "Plant", "Cluster of Customer", "Liquidation", "Subset", "Set Number", "Flexible Group Request"
    ],
    "ELGBBASE": [
        "Sub-Agreement External ID", "Base Value", "Valid From", "Valid To"
    ],
    "ELGBASEVAL": [
        "Sub-Agreement External ID", "Base Value", "Source", "Source Field",
        "Sequence", "Contribution", "Eligibility %", "Valid From", "Valid To"
    ],
    "ELGBPRDGRP": [
        "Sub-Agreement External ID", "Group Name", "Group Type", "Source", "Flexible Group",
        "Valid From", "Valid To"
    ],
    "BENEFITS": [
        "Sub-Agreement External ID", "Payout Group", "Attain Group", "Rate", "Currency(%)",
        # "Target Type", "Target Value", "Target Unit", "Valid From", "Valid To", "Product External ID"
        "Target Type", "Target Value", "Target Unit", "Valid From", "Valid To"
    ],
    "BSA": [
        "Bracket", "Scale Type", "Unit", "Unit Type"
    ],
    "BS": [
        "Dimension Value1", "Rate", "Calculation Derivation Record"
    ],
    "ACCRTE": [
        "Sub-Agreement External ID", "Payout Group", "Rate", "Currency(%)",
        # "Valid From", "Valid To", "Product External ID"
        "Valid From", "Valid To"
    ],
    "END": [
        "Sub-Agreement External ID", "Payout Group", "Rate", "Currency(%)",
        "Valid From", "Valid To", "Product External ID"
    ]
}

# Sub-headings per section tag for Amount format
AMOUNT_HEADERS = {
    "HEADER": ["Request Description"],
    "AGRMT": [
        "Master Agreement External ID", "Master Agreement Description", "Company",
        "Business Unit", "Legal Valid From", "Legal Valid To"
    ],
    "SUBAGRMT": [
        "Processing Sequence", "Sub Agreement External ID", "Sub Agreement Description",
        "Sub Agreement Program", "Period Profile", "Accrual Frequency", "Settlement Frequency",
        "GST Indicator", "Key for Prov G/L", "Legal Valid From", "Legal Valid To"
    ],
    "ELCUST": [
        "Sub-Agreement External ID", "Customer External ID", "Flexible Group",
        "Valid From", "Valid To"
    ],
    "FGHD": [
        "Description", "Flexible Group Type", "Flexible Group", "Flexible Group Request"
    ],
    "FGIT": [
        "Flexible Group Category", "Include/Exclude", "Material External ID", "Material Group",
        "Capacity", "Series", "Star Rating", "Star Rating Year", "Feature 1", "Feature 2",
        "Feature 3", "Feature 4", "Feature 5", "Feature 6", "Feature 7", "Feature 8",
        "Feature 9", "Feature 10", "Customer Group", "Region", "State", "Territory/Sales Area",
        "Plant", "Cluster of Customer", "Liquidation", "Subset", "Set Number", "Flexible Group Request"
    ],
    "ELGBBASE": [
        "Sub-Agreement External ID", "Base Value", "Valid From", "Valid To"
    ],
    "ELGBASEVAL": [
        "Sub-Agreement External ID", "Base Value", "Source", "Source Field",
        "Sequence", "Contribution", "Eligibility %", "Valid From", "Valid To"
    ],
    "ELGBPRDGRP": [
        "Sub-Agreement External ID", "Group Name", "Group Type", "Source", "Flexible Group",
        "Valid From", "Valid To"
    ],
    "BENEFITS": [
        "Sub-Agreement External ID", "Payout Group", "Attain Group", "Rate",
        "Target Type", "Target Value", "Target Unit", "Valid From", "Valid To", "Product External ID"
    ],
    "BSA": [
        "Bracket", "Scale Type", "Unit", "Unit Type"
    ],
    "BS": [
        "Dimension Value1", "Rate", "Calculation Derivation Record"
    ],
    "ACCRTE": [
        "Sub-Agreement External ID", "Payout Group", "Rate",
        "Valid From", "Valid To", "Product External ID"
    ],
    "END": [
        "Sub-Agreement External ID", "Payout Group", "Rate",
        "Valid From", "Valid To", "Product External ID"
    ]
}


def get_program_format_type(program_code, default_format="Amount"):
    """
    Returns the format type ('Amount' or 'Quantity') for a given Sub Agreement Program technical code.
    Defaults to default_format ('Amount') if unrecognized or missing.
    """
    if not program_code or pd.isna(program_code):
        return default_format
    code = str(program_code).strip().upper()
    return PROGRAM_FORMAT_TYPE.get(code, default_format)


def get_section_headers(format_type, tag_name):
    """
    Returns the sub-headings list for a given format type ('Quantity' or 'Amount') and section tag name.
    """
    fmt = str(format_type).strip().capitalize()
    headers_dict = QUANTITY_HEADERS if fmt == "Quantity" else AMOUNT_HEADERS
    return headers_dict.get(tag_name.upper(), [])
