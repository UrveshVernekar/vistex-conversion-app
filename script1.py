import pandas as pd
import numpy as np
import openpyxl
import os
import re
import datetime

from formats import get_program_format_type, get_section_headers

# Embedded 3-level Header Template (Row 0: Section, Row 1: Tag, Row 2: Column Name)
# Eliminates external dependency on Output.xlsx
TEMPLATE_SECTION_HEADER = [
    "Request Header", "Master Agreement Header", None, None, None, None, None,
    "Sub Agreement Header", None, None, None, None, None, None, None,
    "Eligibility of Customer Validity", None, None, None, None,
    "Eligibility of Customer Group", None, None, None, None, None, None,
    "Eligibility of Customer Group Category\n", None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
    "Source & base Units ", None, None, None, None, None, None, None, None,
    "Eligible Product Grouping ", None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
    "Benefits and respective Scales ", None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
    "Accrual Rule ", None, None, None, None
]


TEMPLATE_TAG_HEADER = [
    "HEADER", "AGRMT", None, None, None, None, None,
    "SUBAGRMT", None, None, None, None, None, None, None,
    "ELCUST", None, None, None, None,
    "FGHD", None, None, None, None, None, None,
    "FGIT", None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
    "ELGBBASE", None, "ELGBASEVAL", None, None, None, None, None, None,
    "ELGBPRDGRP", None, None, None, None, None, None,
    "FGHD", None, None, None, None, None, None,
    "FGIT", None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
    "FGHD", None, None, None, None, None, None,
    "FGIT", None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None, None,
    "BENEFITS", None, None, None, None, None, None,
    "BSA", None, None, None,
    "BS", None,
    "BENEFITS", None, None, None, None, None, None,
    "BSA", None, None, None,
    "BS", None,
    "ACCRTE", None, None, None, None
]


TEMPLATE_COLUMN_HEADER = [
    "MR AG Upload 21", "Master Agreement External ID", "Master Agreement Description", "Company", "Business Unit", "Valid From", "Valid To",
    "Processing Sequence", "Sub Agreement External ID", "Sub Agreement Description", "Sub Agreement Program", "Valid From", "Valid To", "Period Profile", "Settlement Frequency",
    "Sub-Agreement External ID", "Customer External ID", "Flexible Group", "Valid From", "Valid To",
    "Description", "Object Type", "Local", "Flexible Group Type", "Flexible Group Request", "Valid From", "Valid To",
    "Flexible Group Category", "Include/Exclude", "Material External ID", "Product Hierarchy", "Material Group", "Capacity", "Series", "Star Rating", "Star Rating Year", "Feature 1", "Feature 2", "Feature 3", "Feature 4", "Feature 5", "Feature 6", "Feature 7", "Feature 8", "Feature 9", "Feature 10", "Customer Group", "Region", "State", "Territory/Sales Area", "Plant", "Cluster of Customer", "Subset", "Set Number", "Flexible Group Request", "Request Action",
    "Sub-Agreement External ID", "Base Value",
    "Sub-Agreement External ID", "Base Value", "Source", "Source Field", "Sequence", "Contribution", "Eligibility %",
    "Sub-Agreement External ID", "Group Name", "Group Type", "Source", "Flexible Group", "Valid From", "Valid To",
    "Description", "Object Type", "Local", "Flexible Group Type", "Flexible Group Request", "Valid From", "Valid To",
    "Flexible Group Category", "Include/Exclude", "Material External ID", "Product Hierarchy", "Material Group", "Capacity", "Series", "Star Rating", "Star Rating Year", "Feature 1", "Feature 2", "Feature 3", "Feature 4", "Feature 5", "Feature 6", "Feature 7", "Feature 8", "Feature 9", "Feature 10", "Customer Group", "Region", "State", "Territory/Sales Area", "Plant", "Cluster of Customer", "Subset", "Set Number", "Flexible Group Request", "Request Action",
    "Description", "Object Type", "Local", "Flexible Group Type", "Flexible Group Request", "Valid From", "Valid To",
    "Flexible Group Category", "Include/Exclude", "Material External ID", "Product Hierarchy", "Material Group", "Capacity", "Series", "Star Rating", "Star Rating Year", "Feature 1", "Feature 2", "Feature 3", "Feature 4", "Feature 5", "Feature 6", "Feature 7", "Feature 8", "Feature 9", "Feature 10", "Customer Group", "Region", "State", "Territory/Sales Area", "Plant", "Cluster of Customer", "Subset", "Set Number", "Flexible Group Request", "Request Action",
    "Sub-Agreement External ID", "Payout Group", "Rate Type", "Rate", "Payout Unit", "Payout Uni", "Dummy",
    "Bracket", "Scale Type", "Unit", "Unit Type",
    "Dimension Value1", "Rate",
    "Sub-Agreement External ID", "Payout Group", "Rate Type", "Rate", "Payout Unit", "Payout Uni", "Dummy",
    "Bracket", "Scale Type", "Unit", "Unit Type",
    "Dimension Value1", "Rate",
    "Sub-Agreement External ID", "Payout Group", "Rate", "Valid From", "Valid To"
]


def clean_cell_value(val):
    """Formats date values into dd-mm-yyyy string format and converts numeric values to native numbers."""
    if pd.isna(val) or val is None or val == '':
        return val
    if isinstance(val, bool):
        return val
    if isinstance(val, (pd.Timestamp, np.datetime64, datetime.datetime, datetime.date)):
        return pd.to_datetime(val).strftime('%d-%m-%Y')
    
    val_str = str(val).strip()
    match_ymd = re.match(r'^(\d{4})[-/](\d{2})[-/](\d{2})', val_str)
    if match_ymd:
        y, m, d = match_ymd.groups()
        return f"{d}-{m}-{y}"
    match_dmy = re.match(r'^(\d{2})[-/](\d{2})[-/](\d{4})', val_str)
    if match_dmy:
        d, m, y = match_dmy.groups()
        return f"{d}-{m}-{y}"
        
    if re.match(r'^-?\d+$', val_str):
        if len(val_str) > 1 and val_str.startswith('0'):
            return val_str
        try:
            return int(val_str)
        except ValueError:
            pass
            
    if re.match(r'^-?\d+\.\d+$', val_str):
        try:
            return float(val_str)
        except ValueError:
            pass

    return val


def normalize_str(s):
    """
    Normalizes string by removing whitespace, special characters, and converting to lowercase
    for robust sub-heading matching across variable input Excel formats.
    """
    if not s:
        return ''
    s = str(s).lower().strip()
    return re.sub(r'[^a-z0-9]', '', s)


def parse_upload_format(input_file):
    """
    Dynamically extracts sub-headings and data blocks from input Excel or CSV file.
    Reads sub-headings from header rows preceding each section tag dynamically so that
    changing or custom sub-headings are automatically recognized.
    """
    sheets_data = {}
    
    if input_file.lower().endswith('.csv'):
        sheets_dict = {'Sheet1': pd.read_csv(input_file, header=None)}
    else:
        excel_file = pd.ExcelFile(input_file)
        sheets_dict = {sn: pd.read_excel(excel_file, sheet_name=sn, header=None) for sn in excel_file.sheet_names}
    
    for sheet_name, df in sheets_dict.items():
        if df.empty:
            continue
            
        master_agreements = []
        current_master = None
        current_sub = None
        
        current_headers_map = {}
        current_fghd_context = None # 'ELCUST' or 'ELGBPRDGRP'
        
        for row_idx in range(len(df)):
            row = df.iloc[row_idx]
            row_vals = [row[c] for c in range(len(row))]
            
            # Find non-null items with column index
            items = [(c, str(val).strip()) for c, val in enumerate(row_vals) if pd.notna(val) and str(val).strip() != '']
            if not items:
                continue
                
            row_str = " | ".join([v for _, v in items])
            
            # Determine if this row is a Tag row
            tag = None
            for c, val in items:
                if val in ['HEADER', 'AGRMT', 'SUBAGRMT', 'ELCUST', 'ELGBBASE', 'ELGBASEVAL', 'ELGBPRDGRP', 'BENEFITS', 'BSA', 'BS', 'FGHD', 'FGIT', 'ACCRTE', 'END']:
                    tag = val
                    break
            
            # Sub-heading / Header line detection if tag is None
            if tag is None:
                current_headers_map = {}
                for c, val in items:
                    current_headers_map[c] = val
                
                # Update FGHD context based on section header
                if 'Eligibility of Customer Group' in row_str:
                    current_fghd_context = 'ELCUST'
                elif 'Eligible Product Grouping- Flexi Group' in row_str or 'Eligible Product Grouping' in row_str:
                    current_fghd_context = 'ELGBPRDGRP'
                continue
                
            # Process Tag row using dynamically mapped sub-headings
            if tag == 'HEADER':
                req_desc = None
                for c, val in items:
                    if c >= 3:
                        req_desc = val
                        break
                if not current_master:
                    current_master = {
                        'header': {'Request Description': req_desc},
                        'agrmt': {},
                        'sub_agreements': []
                    }
                    master_agreements.append(current_master)
                else:
                    current_master['header']['Request Description'] = req_desc
                    
            elif tag == 'AGRMT':
                if not current_master:
                    current_master = {
                        'header': {},
                        'agrmt': {},
                        'sub_agreements': []
                    }
                    master_agreements.append(current_master)
                
                agrmt_data = {}
                for c, val in items:
                    h_name = current_headers_map.get(c)
                    if h_name and h_name not in ['Sub Heading', 'AGRMT', 'Value']:
                        agrmt_data[h_name] = val
                current_master['agrmt'] = agrmt_data
                
            elif tag == 'SUBAGRMT':
                if not current_master:
                    current_master = {
                        'header': {},
                        'agrmt': {},
                        'sub_agreements': []
                    }
                    master_agreements.append(current_master)
                    
                sub_data = {}
                for c, val in items:
                    h_name = current_headers_map.get(c)
                    if h_name and h_name not in ['Sub Heading', 'SUBAGRMT', 'Value', 'Can repeat']:
                        sub_data[h_name] = val
                        
                current_sub = {
                    'subagrmt': sub_data,
                    'elcust': [],
                    'elgbbase': [],
                    'elgbaseval': [],
                    'elgbprdgrp': [],
                    'benefits': [],
                    'bsa': [],
                    'bs': [],
                    'fghd_elcust': [],
                    'fgit_elcust': [],
                    'fghd_prdgrp': [],
                    'fgit_prdgrp': [],
                    'accrte': []
                }
                current_master['sub_agreements'].append(current_sub)
                
            elif tag == 'END':
                # 'END' delimiter marks the end of the current sub agreement
                current_sub = None
                current_fghd_context = None
                
            elif current_sub is not None:
                data_dict = {}
                for c, val in items:
                    h_name = current_headers_map.get(c)
                    if h_name and h_name not in ['Sub Heading', 'Value', 'Can repeat', tag]:
                        data_dict[h_name] = val
                        
                if tag == 'ELCUST':
                    current_sub['elcust'].append(data_dict)
                elif tag == 'ELGBBASE':
                    current_sub['elgbbase'].append(data_dict)
                elif tag == 'ELGBASEVAL':
                    current_sub['elgbaseval'].append(data_dict)
                elif tag == 'ELGBPRDGRP':
                    current_sub['elgbprdgrp'].append(data_dict)
                    current_fghd_context = 'ELGBPRDGRP'
                elif tag == 'BENEFITS':
                    current_sub['benefits'].append(data_dict)
                elif tag == 'BSA':
                    current_sub['bsa'].append(data_dict)
                elif tag == 'BS':
                    current_sub['bs'].append(data_dict)
                elif tag == 'FGHD':
                    # Determine context based on content or previous section
                    is_cust = (current_fghd_context == 'ELCUST')
                    for k_name, val in data_dict.items():
                        if normalize_str(val) in ['m7cu', 'custgroup', 'clgroup', 'customergroup']:
                            is_cust = True
                            break
                    if is_cust:
                        current_sub['fghd_elcust'].append(data_dict)
                    else:
                        current_sub['fghd_prdgrp'].append(data_dict)
                elif tag == 'FGIT':
                    is_cust = (current_fghd_context == 'ELCUST')
                    for k_name in data_dict.keys():
                        if normalize_str(k_name) in ['customergroup', 'clusterofcustomer', 'customerid']:
                            is_cust = True
                            break
                    if is_cust:
                        current_sub['fgit_elcust'].append(data_dict)
                    else:
                        current_sub['fgit_prdgrp'].append(data_dict)
                elif tag == 'ACCRTE':
                    current_sub['accrte'].append(data_dict)
                    
        sheets_data[sheet_name] = master_agreements
        
    return sheets_data


def convert_upload_to_output(input_upload_file='Upload Format.xlsx',
                            output_file='output1.xlsx',
                            template_output_file=None):
    """
    Converts input upload Excel file into output Excel file formatted according to the template schema.
    Dynamically matches sub-headings from the input file to the target columns.
    """
    if not os.path.exists(input_upload_file):
        raise FileNotFoundError(f"Input file '{input_upload_file}' not found.")

    # 1. Parse Upload Format dynamically reading sub-headings
    parsed_sheets = parse_upload_format(input_upload_file)
    
    # 2. Header Resolution: Load from file if provided, otherwise use embedded template
    if template_output_file and os.path.exists(template_output_file):
        df_out_tmpl = pd.read_excel(template_output_file, sheet_name=0, header=None)
        sec_row = list(df_out_tmpl.iloc[0])
        tag_row = list(df_out_tmpl.iloc[1])
        col_row = list(df_out_tmpl.iloc[2])
    else:
        sec_row = TEMPLATE_SECTION_HEADER
        tag_row = TEMPLATE_TAG_HEADER
        col_row = TEMPLATE_COLUMN_HEADER
    
    num_cols = len(col_row)
    
    # Pre-normalize template column headers for fast flexible matching
    norm_col_row = [normalize_str(c) for c in col_row]
    
    def get_col_index(section_key, col_name):
        col_name_clean = str(col_name).strip()
        norm_in_col = normalize_str(col_name_clean)
        
        ranges = {
            'HEADER': (0, 0),
            'AGRMT': (1, 6),
            'SUBAGRMT': (7, 14),
            'ELCUST': (15, 19),
            'FGHD_CUST': (20, 26),
            'FGIT_CUST': (27, 55),
            'ELGBBASE': (56, 57),
            'ELGBASEVAL': (58, 64),
            'ELGBPRDGRP': (65, 71),
            'FGHD_PRD': (72, 78),
            'FGIT_PRD': (79, 107),
            'BENEFITS': (144, 150),
            'BSA': (151, 154),
            'BS': (155, 156),
            'BENEFITS2': (157, 163),
            'BSA2': (164, 167),
            'BS2': (168, 169),
            'ACCRTE': (170, 174),
        }
        
        if section_key not in ranges:
            return None
            
        start_c, end_c = ranges[section_key]
        
        # Broad alias mapping for sub-heading variations
        aliases = {
            'Master Agreement External ID': ['masteragreementexternalid', 'masteragreementid', 'externalid', 'masteragreement'],
            'Master Agreement Description': ['masteragreementdescription', 'masterdescription', 'description'],
            'Company': ['company', 'companycode'],
            'Business Unit': ['businessunit', 'bu'],
            'Legal Valid From': ['validfrom', 'legalvalidfrom', 'validfromdate', 'fromdate'],
            'Legal Valid To': ['validto', 'legalvalidto', 'validtodate', 'todate'],
            'Processing Sequence': ['processingsequence', 'sequence', 'procseq', 'seq'],
            'Sub Agreement External ID': ['subagreementexternalid', 'subagreementid', 'externalid', 'subagreement'],
            'Sub Agreement Description': ['subagreementdescription', 'description', 'subagreementname'],
            'Sub Agreement Program': ['subagreementprogram', 'program', 'programid'],
            'Period Profile': ['periodprofile', 'profile'],
            'Settlement Frequency': ['settlementfrequency', 'settlementfreq', 'frequency'],
            'Customer External ID': ['customerexternalid', 'customerid', 'customer'],
            'Flexible Group': ['flexiblegroup', 'flexigroup', 'fg'],
            'Valid From': ['validfrom', 'legalvalidfrom', 'fromdate'],
            'Valid To': ['validto', 'legalvalidto', 'todate'],
            'Group Name': ['groupname', 'productgroupname', 'name'],
            'Group Type': ['grouptype', 'type'],
            'Source': ['source'],
            'Payout Group': ['payoutgroup', 'payout'],
            'Target Type': ['targettype', 'ratetype', 'type'],
            'Target Value': ['targetvalue', 'rate', 'value', 'amount'],
            'Currency(%)': ['currency', 'payoutunit', 'unit', 'currencypercent'],
            'Target Unit': ['targetunit', 'payoutuni', 'unit'],
            'Bracket': ['bracket', 'scalebracket'],
            'Scale Type': ['scaletype'],
            'Unit': ['unit', 'payoutunit', 'currency'],
            'Unit Type': ['unittype'],
            'Dimension Value1': ['dimensionvalue1', 'dimensionvalue', 'scalevalue', 'dimension1', 'val1'],
            'Rate': ['rate', 'value', 'amount', 'targetvalue'],
            'Description': ['description', 'groupdescription'],
            'Flexible Group Type': ['flexiblegrouptype', 'flexigrouptype'],
            'Flexible Group Request': ['flexiblegrouprequest', 'flexigrouprequest'],
            'Flexible Group Category': ['flexiblegroupcategory', 'flexigroupcategory', 'category'],
            'Include/Exclude': ['includeexclude', 'incexc', 'include', 'exclude'],
            'Material External ID': ['materialexternalid', 'materialid', 'material'],
            'Material Group': ['materialgroup', 'matgroup'],
            'Customer Group': ['customergroup', 'custgroup'],
            'Set Number': ['setnumber', 'setno', 'set']
        }
        
        # 1. Direct or normalized match in template slice
        for c in range(start_c, end_c + 1):
            if norm_in_col == norm_col_row[c]:
                return c
                
        # 2. Match via alias dictionary
        possible_norm_aliases = []
        for k, val_list in aliases.items():
            norm_k = normalize_str(k)
            norm_vals = [normalize_str(v) for v in val_list]
            if norm_in_col == norm_k or norm_in_col in norm_vals:
                possible_norm_aliases.append(norm_k)
                possible_norm_aliases.extend(norm_vals)
                
        possible_norm_aliases = list(dict.fromkeys(possible_norm_aliases))
        
        for c in range(start_c, end_c + 1):
            if norm_col_row[c] in possible_norm_aliases:
                return c
                
        # 3. Substring match fallback
        for c in range(start_c, end_c + 1):
            if norm_in_col and (norm_in_col in norm_col_row[c] or norm_col_row[c] in norm_in_col):
                return c
                
        return None

    try:
        writer = pd.ExcelWriter(output_file, engine='openpyxl')
    except PermissionError:
        alt_file = output_file.replace('.xlsx', '_out.xlsx')
        print(f"Warning: File '{output_file}' is locked/open in Excel. Generating '{alt_file}' instead.")
        output_file = alt_file
        writer = pd.ExcelWriter(output_file, engine='openpyxl')
    
    for sheet_name, master_agreements in parsed_sheets.items():
        all_rows = []
        # Header rows 0, 1, 2
        all_rows.append(sec_row)
        all_rows.append(tag_row)
        all_rows.append(col_row)
        
        for master in master_agreements:
            header_info = master['header']
            agrmt_info = master['agrmt']
            sub_agreements = master['sub_agreements']
            
            for sub in sub_agreements:
                max_sub_rows = max(
                    len(sub['bs']),
                    len(sub['elcust']),
                    len(sub['elgbbase']),
                    len(sub['elgbaseval']),
                    len(sub['elgbprdgrp']),
                    len(sub['benefits']),
                    len(sub['bsa']),
                    len(sub['fghd_elcust']),
                    len(sub['fgit_elcust']),
                    len(sub['fghd_prdgrp']),
                    len(sub['fgit_prdgrp']),
                    len(sub['accrte']),
                    1
                )
                
                for r in range(max_sub_rows):
                    row_data = [np.nan] * num_cols
                    
                    # 1. AGRMT (Master Agreement Header)
                    for k, v in agrmt_info.items():
                        c_idx = get_col_index('AGRMT', k)
                        if c_idx is not None:
                            row_data[c_idx] = v
                            
                    # 2. HEADER (Request Description)
                    if r == 0:
                        for k, v in header_info.items():
                            c_idx = get_col_index('HEADER', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 3. SUBAGRMT (Sub Agreement Header)
                    if r == 0:
                        sub_hdr = sub['subagrmt']
                        for k, v in sub_hdr.items():
                            c_idx = get_col_index('SUBAGRMT', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 4. ELCUST (Eligibility of Customer Validity)
                    if r < len(sub['elcust']):
                        elcust_item = sub['elcust'][r]
                        for k, v in elcust_item.items():
                            c_idx = get_col_index('ELCUST', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                    c_sub_id = get_col_index('ELCUST', 'Sub-Agreement External ID')
                    if c_sub_id is not None and pd.isna(row_data[c_sub_id]) and r < len(sub['elcust']):
                        row_data[c_sub_id] = sub['subagrmt'].get('Sub Agreement External ID')
                        
                    # 4b. ELGBBASE (Eligible Base)
                    if r < len(sub['elgbbase']):
                        item = sub['elgbbase'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('ELGBBASE', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 4c. ELGBASEVAL (Eligible Base Values)
                    if r < len(sub['elgbaseval']):
                        item = sub['elgbaseval'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('ELGBASEVAL', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                        
                    # 5. FGHD Customer Group
                    if r < len(sub['fghd_elcust']):
                        item = sub['fghd_elcust'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('FGHD_CUST', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 6. FGIT Customer Group
                    if r < len(sub['fgit_elcust']):
                        item = sub['fgit_elcust'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('FGIT_CUST', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 7. ELGBPRDGRP (Eligible Product Grouping)
                    if r < len(sub['elgbprdgrp']):
                        item = sub['elgbprdgrp'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('ELGBPRDGRP', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 8. FGHD Product Grouping
                    if r < len(sub['fghd_prdgrp']):
                        item = sub['fghd_prdgrp'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('FGHD_PRD', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 9. FGIT Product Grouping
                    if r < len(sub['fgit_prdgrp']):
                        item = sub['fgit_prdgrp'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('FGIT_PRD', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 10. BENEFITS
                    if r < len(sub['benefits']):
                        item = sub['benefits'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('BENEFITS', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 11. BSA (Scale Attribute)
                    if r < len(sub['bsa']):
                        item = sub['bsa'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('BSA', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 12. BS (Scales)
                    if r < len(sub['bs']):
                        item = sub['bs'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('BS', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    # 13. ACCRTE (Accrual Rate)
                    if r < len(sub['accrte']):
                        item = sub['accrte'][r]
                        for k, v in item.items():
                            c_idx = get_col_index('ACCRTE', k)
                            if c_idx is not None:
                                row_data[c_idx] = v
                                
                    row_data = [clean_cell_value(v) for v in row_data]
                    all_rows.append(row_data)
                    
        df_out_sheet = pd.DataFrame(all_rows)
        df_out_sheet.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        # Generate CSV output alongside .xlsx
        if len(parsed_sheets) == 1:
            csv_file = os.path.splitext(output_file)[0] + '.csv'
        else:
            csv_file = f"{os.path.splitext(output_file)[0]}_{sheet_name}.csv"
        df_out_sheet.to_csv(csv_file, index=False, header=False)
        
        # Format numeric cells explicitly in openpyxl so Excel defaults them to Number category
        ws = writer.sheets[sheet_name]
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, (int, float, np.integer, np.floating)) and not isinstance(cell.value, bool):
                    if isinstance(cell.value, (int, np.integer)) or (isinstance(cell.value, float) and cell.value.is_integer()):
                        cell.number_format = '0'
                    else:
                        cell.number_format = '0.##'
        
    writer.close()
    base_name = os.path.splitext(output_file)[0]
    csv_info = f"'{base_name}.csv'" if len(parsed_sheets) == 1 else f"'{base_name}_*.csv'"
    print(f"Successfully generated '{output_file}' and {csv_info}.")


if __name__ == '__main__':
    # convert_upload_to_output('Amount % - Multiple SA.xlsx', 'output1.xlsx')
    convert_upload_to_output('Quantity - Multiple SA.xlsx', 'output2.xlsx')