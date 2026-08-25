import pandas as pd
import numpy as np
import openpyxl
import os
import datetime
import re

from formats import get_program_format_type, get_section_headers


def clean_cell_value(val):
    """Formats date values into dd/mm/yyyy string format and converts numeric values to native numbers."""
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
        return f"{d}/{m}/{y}"
    match_dmy = re.match(r'^(\d{2})[-/](\d{2})[-/](\d{4})', val_str)
    if match_dmy:
        d, m, y = match_dmy.groups()
        return f"{d}/{m}/{y}"
        
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


def convert_output_to_upload(input_output_file='output1.xlsx', output_reverse_file='reverse-output1.xlsx', remove_first_two_cols=True):
    """
    Performs the reverse transformation of script1.py:
    Takes 'output1.xlsx' (175-column flat table) and converts it back into 
    the vertical block section format of 'Upload Format.xlsx' (saved as 'reverse-output1.xlsx').
    Removes column A and column B if remove_first_two_cols is True.
    Adapts headers dynamically for Amount vs Quantity formats based on Sub Agreement Program.
    """
    if not os.path.exists(input_output_file):
        raise FileNotFoundError(f"Input file '{input_output_file}' not found.")

    try:
        writer = pd.ExcelWriter(output_reverse_file, engine='openpyxl')
    except PermissionError:
        alt_file = output_reverse_file.replace('.xlsx', '_out.xlsx')
        print(f"Warning: File '{output_reverse_file}' is locked/open in Excel. Generating '{alt_file}' instead.")
        output_reverse_file = alt_file
        writer = pd.ExcelWriter(output_reverse_file, engine='openpyxl')

    def add_if_unique(item_list, new_item, key_field):
        val = new_item.get(key_field)
        if not val or pd.isna(val):
            return
        for existing in item_list:
            if existing.get(key_field) == val:
                return
        item_list.append(new_item)

    if input_output_file.lower().endswith('.csv'):
        sheets_dict = {'Sheet1': pd.read_csv(input_output_file, header=None)}
    else:
        excel_in = pd.ExcelFile(input_output_file)
        sheets_dict = {sn: pd.read_excel(excel_in, sheet_name=sn, header=None) for sn in excel_in.sheet_names}

    total_sheets = len(sheets_dict)

    for sheet_name, df_out in sheets_dict.items():
        if df_out.empty or len(df_out) < 4:
            # Empty sheet or header only
            df_empty = pd.DataFrame()
            df_empty.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            if total_sheets == 1:
                csv_file = os.path.splitext(output_reverse_file)[0] + '.csv'
            else:
                csv_file = f"{os.path.splitext(output_reverse_file)[0]}_{sheet_name}.csv"
            df_empty.to_csv(csv_file, index=False, header=False)
            continue

        sec_row = list(df_out.iloc[0])
        tag_row = list(df_out.iloc[1])
        col_row = list(df_out.iloc[2])

        data_rows = df_out.iloc[3:].copy()

        # Build dynamic section ranges for current sheet
        def build_dynamic_ranges(tag_row_list, total_cols):
            blocks = []
            for idx, tag in enumerate(tag_row_list):
                if pd.notna(tag) and str(tag).strip():
                    blocks.append((str(tag).strip().upper(), idx))
            
            if not blocks:
                return {}

            raw_ranges = []
            for i in range(len(blocks)):
                tag_name, start_idx = blocks[i]
                end_idx = blocks[i+1][1] - 1 if i + 1 < len(blocks) else total_cols - 1
                raw_ranges.append((tag_name, start_idx, end_idx))

            dynamic_ranges = {}
            fghd_count = 0
            fgit_count = 0
            benefits_count = 0
            bsa_count = 0
            bs_count = 0

            for tag_name, start_c, end_c in raw_ranges:
                if tag_name == 'HEADER':
                    dynamic_ranges['HEADER'] = (start_c, end_c)
                elif tag_name == 'AGRMT':
                    dynamic_ranges['AGRMT'] = (start_c, end_c)
                elif tag_name == 'SUBAGRMT':
                    dynamic_ranges['SUBAGRMT'] = (start_c, end_c)
                elif tag_name == 'ELCUST':
                    dynamic_ranges['ELCUST'] = (start_c, end_c)
                elif tag_name == 'FGHD':
                    fghd_count += 1
                    key = 'FGHD_CUST' if fghd_count == 1 else ('FGHD_PRD' if fghd_count == 2 else f'FGHD_{fghd_count}')
                    dynamic_ranges[key] = (start_c, end_c)
                elif tag_name == 'FGIT':
                    fgit_count += 1
                    key = 'FGIT_CUST' if fgit_count == 1 else ('FGIT_PRD' if fgit_count == 2 else f'FGIT_{fgit_count}')
                    dynamic_ranges[key] = (start_c, end_c)
                elif tag_name == 'ELGBBASE':
                    dynamic_ranges['ELGBBASE'] = (start_c, end_c)
                elif tag_name == 'ELGBASEVAL':
                    dynamic_ranges['ELGBASEVAL'] = (start_c, end_c)
                elif tag_name == 'ELGBPRDGRP':
                    dynamic_ranges['ELGBPRDGRP'] = (start_c, end_c)
                elif tag_name == 'BENEFITS':
                    benefits_count += 1
                    key = 'BENEFITS' if benefits_count == 1 else f'BENEFITS{benefits_count}'
                    dynamic_ranges[key] = (start_c, end_c)
                elif tag_name == 'BSA':
                    bsa_count += 1
                    key = 'BSA' if bsa_count == 1 else f'BSA{bsa_count}'
                    dynamic_ranges[key] = (start_c, end_c)
                elif tag_name == 'BS':
                    bs_count += 1
                    key = 'BS' if bs_count == 1 else f'BS{bs_count}'
                    dynamic_ranges[key] = (start_c, end_c)
                elif tag_name == 'ACCRTE':
                    dynamic_ranges['ACCRTE'] = (start_c, end_c)
                else:
                    dynamic_ranges[tag_name] = (start_c, end_c)

            return dynamic_ranges

        ranges = build_dynamic_ranges(tag_row, len(col_row))

        def get_col_idx(section_key, col_name, fallback=None):
            if section_key in ranges:
                start_c, end_c = ranges[section_key]
                norm_col_name = str(col_name).strip().lower().replace(" ", "").replace("-", "")
                for i in range(start_c, min(end_c + 1, len(col_row))):
                    c_val = str(col_row[i]).strip().lower().replace(" ", "").replace("-", "") if pd.notna(col_row[i]) else ""
                    if c_val == norm_col_name:
                        return i
            return fallback

        def get_val(row_list, section_key, col_name, fallback_index=None):
            idx = get_col_idx(section_key, col_name, fallback_index)
            if idx is not None and idx < len(row_list):
                v = row_list[idx]
                if pd.notna(v):
                    return v
            return None

        # Parse output data rows into Master Agreement & Sub Agreement objects
        master_agreements = []
        curr_master = None
        curr_sub = None

        for idx, row in data_rows.iterrows():
            row_list = list(row)
            
            # Check Master Agreement ID in Col 1
            ma_id = get_val(row_list, 'AGRMT', 'Master Agreement External ID', 1)
            ma_id_str = str(ma_id).strip() if ma_id else None

            # Check Sub Agreement ID or Processing Sequence
            sa_id = get_val(row_list, 'SUBAGRMT', 'Sub Agreement External ID', 8)
            sa_id_str = str(sa_id).strip() if sa_id else None
            proc_seq = get_val(row_list, 'SUBAGRMT', 'Processing Sequence', 7)

            if not curr_master or (ma_id_str and curr_master['ma_id'] != ma_id_str):
                curr_master = {
                    'ma_id': ma_id_str,
                    'header': {'Request Description': get_val(row_list, 'HEADER', 'MR AG Upload 21', 0) or ma_id_str},
                    'agrmt': {
                        'Master Agreement External ID': get_val(row_list, 'AGRMT', 'Master Agreement External ID', 1) or '',
                        'Master Agreement Description': get_val(row_list, 'AGRMT', 'Master Agreement Description', 2) or '',
                        'Company': get_val(row_list, 'AGRMT', 'Company', 3) or '',
                        'Business Unit': get_val(row_list, 'AGRMT', 'Business Unit', 4) or '',
                        'Legal Valid From': get_val(row_list, 'AGRMT', 'Valid From', 5) or '',
                        'Legal Valid To': get_val(row_list, 'AGRMT', 'Valid To', 6) or ''
                    },
                    'sub_agreements': []
                }
                master_agreements.append(curr_master)

            if sa_id_str or proc_seq or not curr_master['sub_agreements']:
                curr_sub = {
                    'subagrmt': {
                        'Processing Sequence': proc_seq if proc_seq is not None else len(curr_master['sub_agreements']) + 1,
                        'Sub Agreement External ID': sa_id_str if sa_id_str else '',
                        'Sub Agreement Description': get_val(row_list, 'SUBAGRMT', 'Sub Agreement Description', 9) or sa_id_str or '',
                        'Sub Agreement Program': get_val(row_list, 'SUBAGRMT', 'Sub Agreement Program', 10) or '',
                        'Period Profile': get_val(row_list, 'SUBAGRMT', 'Period Profile', 11) or '',
                        'Accrual Frequency': get_val(row_list, 'SUBAGRMT', 'Accrual Frequency', 12) or 'MD',
                        'Settlement Frequency': get_val(row_list, 'SUBAGRMT', 'Settlement Frequency', 13) or '',
                        'GST Indicator': get_val(row_list, 'SUBAGRMT', 'GST Indicator', 14) or 'N',
                        'Key for Prov G/L': get_val(row_list, 'SUBAGRMT', 'Key for Prov G/L', 15) or 'T',
                        'Legal Valid From': get_val(row_list, 'SUBAGRMT', 'Valid From', 16) or '',
                        'Legal Valid To': get_val(row_list, 'SUBAGRMT', 'Valid To', 17) or ''
                    },
                    'elcust': [],
                    'elgbbase': [],
                    'elgbaseval': [],
                    'elgbprdgrp': [],
                    'benefits': [],
                    'bsa': [],
                    'bs': [],
                    'fghd_cust': [],
                    'fgit_cust': [],
                    'fghd_prd': [],
                    'fgit_prd': [],
                    'accrte': []
                }
                curr_master['sub_agreements'].append(curr_sub)

            if curr_sub:
                sub_ext_id = curr_sub['subagrmt'].get('Sub Agreement External ID', '')
                
                # ELCUST
                cust_ext_id = get_val(row_list, 'ELCUST', 'Customer External ID')
                if cust_ext_id:
                    add_if_unique(curr_sub['elcust'], {
                        'Sub-Agreement External ID': get_val(row_list, 'ELCUST', 'Sub-Agreement External ID') or sub_ext_id,
                        'Customer External ID': cust_ext_id,
                        'Flexible Group': get_val(row_list, 'ELCUST', 'Flexible Group') or '',
                        'Valid From': get_val(row_list, 'ELCUST', 'Valid From') or '',
                        'Valid To': get_val(row_list, 'ELCUST', 'Valid To') or ''
                    }, 'Customer External ID')

                # ELGBBASE
                base_val = get_val(row_list, 'ELGBBASE', 'Base Value')
                if base_val:
                    add_if_unique(curr_sub['elgbbase'], {
                        'Sub-Agreement External ID': get_val(row_list, 'ELGBBASE', 'Sub-Agreement External ID') or sub_ext_id,
                        'Base Value': base_val,
                        'Valid From': get_val(row_list, 'ELGBBASE', 'Valid From') or '',
                        'Valid To': get_val(row_list, 'ELGBBASE', 'Valid To') or ''
                    }, 'Base Value')

                # ELGBASEVAL
                base_val_detail = get_val(row_list, 'ELGBASEVAL', 'Base Value')
                if base_val_detail:
                    add_if_unique(curr_sub['elgbaseval'], {
                        'Sub-Agreement External ID': get_val(row_list, 'ELGBASEVAL', 'Sub-Agreement External ID') or sub_ext_id,
                        'Base Value': base_val_detail,
                        'Source': get_val(row_list, 'ELGBASEVAL', 'Source') or '',
                        'Source Field': get_val(row_list, 'ELGBASEVAL', 'Source Field') or '',
                        'Sequence': get_val(row_list, 'ELGBASEVAL', 'Sequence') or 1,
                        'Contribution': get_val(row_list, 'ELGBASEVAL', 'Contribution') or '+',
                        'Eligibility %': get_val(row_list, 'ELGBASEVAL', 'Eligibility %') or 100,
                        'Valid From': get_val(row_list, 'ELGBASEVAL', 'Valid From') or '',
                        'Valid To': get_val(row_list, 'ELGBASEVAL', 'Valid To') or ''
                    }, 'Base Value')

                # FGHD Cust
                fghd_c_desc = get_val(row_list, 'FGHD_CUST', 'Description')
                if fghd_c_desc:
                    add_if_unique(curr_sub['fghd_cust'], {
                        'Description': fghd_c_desc,
                        'Object Type': get_val(row_list, 'FGHD_CUST', 'Object Type') or '',
                        'Local': get_val(row_list, 'FGHD_CUST', 'Local') or '',
                        'Flexible Group Type': get_val(row_list, 'FGHD_CUST', 'Flexible Group Type') or '',
                        'Flexible Group Request': get_val(row_list, 'FGHD_CUST', 'Flexible Group Request') or '',
                        'Valid From': get_val(row_list, 'FGHD_CUST', 'Valid From') or '',
                        'Valid To': get_val(row_list, 'FGHD_CUST', 'Valid To') or ''
                    }, 'Description')

                # FGIT Cust
                fgit_c_cat = get_val(row_list, 'FGIT_CUST', 'Flexible Group Category')
                if fgit_c_cat:
                    fgit_c_item = {
                        'Flexible Group Category': fgit_c_cat,
                        'Include/Exclude': get_val(row_list, 'FGIT_CUST', 'Include/Exclude') or 'X',
                        'Material External ID': get_val(row_list, 'FGIT_CUST', 'Material External ID') or '',
                        'Material Group': get_val(row_list, 'FGIT_CUST', 'Material Group') or '',
                        'Capacity': get_val(row_list, 'FGIT_CUST', 'Capacity') or '',
                        'Customer Group': get_val(row_list, 'FGIT_CUST', 'Customer Group') or '',
                        'Region': get_val(row_list, 'FGIT_CUST', 'Region') or '',
                        'State': get_val(row_list, 'FGIT_CUST', 'State') or '',
                        'Territory/Sales Area': get_val(row_list, 'FGIT_CUST', 'Territory/Sales Area') or '',
                        'Plant': get_val(row_list, 'FGIT_CUST', 'Plant') or '',
                        'Cluster of Customer': get_val(row_list, 'FGIT_CUST', 'Cluster of Customer') or '',
                        'Liquidation': get_val(row_list, 'FGIT_CUST', 'Liquidation') or '',
                        'Subset': get_val(row_list, 'FGIT_CUST', 'Subset') or '',
                        'Set Number': get_val(row_list, 'FGIT_CUST', 'Set Number') or 1,
                        'Flexible Group Request': get_val(row_list, 'FGIT_CUST', 'Flexible Group Request') or ''
                    }
                    if fgit_c_item not in curr_sub['fgit_cust']:
                        curr_sub['fgit_cust'].append(fgit_c_item)

                # ELGBPRDGRP
                grp_name = get_val(row_list, 'ELGBPRDGRP', 'Group Name')
                if grp_name:
                    add_if_unique(curr_sub['elgbprdgrp'], {
                        'Sub-Agreement External ID': get_val(row_list, 'ELGBPRDGRP', 'Sub-Agreement External ID') or sub_ext_id,
                        'Group Name': grp_name,
                        'Group Type': get_val(row_list, 'ELGBPRDGRP', 'Group Type') or '',
                        'Source': get_val(row_list, 'ELGBPRDGRP', 'Source') or '',
                        'Flexible Group': get_val(row_list, 'ELGBPRDGRP', 'Flexible Group') or '',
                        'Valid From': get_val(row_list, 'ELGBPRDGRP', 'Valid From') or '',
                        'Valid To': get_val(row_list, 'ELGBPRDGRP', 'Valid To') or ''
                    }, 'Group Name')

                # FGHD Prd
                fghd_p_desc = get_val(row_list, 'FGHD_PRD', 'Description')
                if fghd_p_desc:
                    add_if_unique(curr_sub['fghd_prd'], {
                        'Description': fghd_p_desc,
                        'Object Type': get_val(row_list, 'FGHD_PRD', 'Object Type') or '',
                        'Local': get_val(row_list, 'FGHD_PRD', 'Local') or '',
                        'Flexible Group Type': get_val(row_list, 'FGHD_PRD', 'Flexible Group Type') or '',
                        'Flexible Group Request': get_val(row_list, 'FGHD_PRD', 'Flexible Group Request') or '',
                        'Valid From': get_val(row_list, 'FGHD_PRD', 'Valid From') or '',
                        'Valid To': get_val(row_list, 'FGHD_PRD', 'Valid To') or ''
                    }, 'Description')

                # FGIT Prd
                fgit_p_cat = get_val(row_list, 'FGIT_PRD', 'Flexible Group Category')
                if fgit_p_cat:
                    fgit_p_item = {
                        'Flexible Group Category': fgit_p_cat,
                        'Include/Exclude': get_val(row_list, 'FGIT_PRD', 'Include/Exclude') or 'X',
                        'Material External ID': get_val(row_list, 'FGIT_PRD', 'Material External ID') or '',
                        'Product Hierarchy': get_val(row_list, 'FGIT_PRD', 'Product Hierarchy') or '',
                        'Material Group': get_val(row_list, 'FGIT_PRD', 'Material Group') or '',
                        'Capacity': get_val(row_list, 'FGIT_PRD', 'Capacity') or '',
                        'Series': get_val(row_list, 'FGIT_PRD', 'Series') or '',
                        'Star Rating': get_val(row_list, 'FGIT_PRD', 'Star Rating') or '',
                        'Star Rating Year': get_val(row_list, 'FGIT_PRD', 'Star Rating Year') or '',
                        'Feature 1': get_val(row_list, 'FGIT_PRD', 'Feature 1') or '',
                        'Feature 2': get_val(row_list, 'FGIT_PRD', 'Feature 2') or '',
                        'Feature 3': get_val(row_list, 'FGIT_PRD', 'Feature 3') or '',
                        'Feature 4': get_val(row_list, 'FGIT_PRD', 'Feature 4') or '',
                        'Feature 5': get_val(row_list, 'FGIT_PRD', 'Feature 5') or '',
                        'Feature 6': get_val(row_list, 'FGIT_PRD', 'Feature 6') or '',
                        'Feature 7': get_val(row_list, 'FGIT_PRD', 'Feature 7') or '',
                        'Feature 8': get_val(row_list, 'FGIT_PRD', 'Feature 8') or '',
                        'Feature 9': get_val(row_list, 'FGIT_PRD', 'Feature 9') or '',
                        'Feature 10': get_val(row_list, 'FGIT_PRD', 'Feature 10') or '',
                        'Customer Group': get_val(row_list, 'FGIT_PRD', 'Customer Group') or '',
                        'Region': get_val(row_list, 'FGIT_PRD', 'Region') or '',
                        'State': get_val(row_list, 'FGIT_PRD', 'State') or '',
                        'Territory/Sales Area': get_val(row_list, 'FGIT_PRD', 'Territory/Sales Area') or '',
                        'Plant': get_val(row_list, 'FGIT_PRD', 'Plant') or '',
                        'Cluster of Customer': get_val(row_list, 'FGIT_PRD', 'Cluster of Customer') or '',
                        'Liquidation': get_val(row_list, 'FGIT_PRD', 'Liquidation') or '',
                        'Subset': get_val(row_list, 'FGIT_PRD', 'Subset') or '',
                        'Set Number': get_val(row_list, 'FGIT_PRD', 'Set Number') or 1,
                        'Flexible Group Request': get_val(row_list, 'FGIT_PRD', 'Flexible Group Request') or ''
                    }
                    if fgit_p_item not in curr_sub['fgit_prd']:
                        curr_sub['fgit_prd'].append(fgit_p_item)

                # BENEFITS
                payout_grp = get_val(row_list, 'BENEFITS', 'Payout Group')
                if payout_grp:
                    sub_vf = curr_sub['subagrmt'].get('Legal Valid From', '')
                    sub_vt = curr_sub['subagrmt'].get('Legal Valid To', '')
                    add_if_unique(curr_sub['benefits'], {
                        'Sub-Agreement External ID': get_val(row_list, 'BENEFITS', 'Sub-Agreement External ID') or sub_ext_id,
                        'Payout Group': payout_grp,
                        'Attain Group': get_val(row_list, 'BENEFITS', 'Attain Group') or '',
                        'Rate': get_val(row_list, 'BENEFITS', 'Rate') or '',
                        'Target Type': get_val(row_list, 'BENEFITS', 'Target Type') or '',
                        'Target Value': get_val(row_list, 'BENEFITS', 'Target Value') or '',
                        'Currency(%)': get_val(row_list, 'BENEFITS', 'Target Unit') or get_val(row_list, 'BENEFITS', 'Currency(%)') or '',
                        'Target Unit': get_val(row_list, 'BENEFITS', 'Target Unit') or '',
                        'Valid From': get_val(row_list, 'BENEFITS', 'Valid From') or sub_vf,
                        'Valid To': get_val(row_list, 'BENEFITS', 'Valid To') or sub_vt
                    }, 'Payout Group')

                # BSA
                bsa_brk = get_val(row_list, 'BSA', 'Bracket')
                if bsa_brk:
                    add_if_unique(curr_sub['bsa'], {
                        'Bracket': bsa_brk,
                        'Scale Type': get_val(row_list, 'BSA', 'Scale Type') or '',
                        'Unit': get_val(row_list, 'BSA', 'Unit') or '',
                        'Unit Type': get_val(row_list, 'BSA', 'Unit Type') or ''
                    }, 'Bracket')

                # ACCRTE
                acc_pg = get_val(row_list, 'ACCRTE', 'Payout Group')
                if acc_pg:
                    add_if_unique(curr_sub['accrte'], {
                        'Sub-Agreement External ID': get_val(row_list, 'ACCRTE', 'Sub-Agreement External ID') or sub_ext_id,
                        'Payout Group': acc_pg,
                        'Rate': get_val(row_list, 'ACCRTE', 'Rate') or '',
                        'Currency(%)': 'INR',
                        'Valid From': get_val(row_list, 'ACCRTE', 'Valid From') or '',
                        'Valid To': get_val(row_list, 'ACCRTE', 'Valid To') or ''
                    }, 'Payout Group')

                # BS (Scales)
                bs_dim = get_val(row_list, 'BS', 'Dimension Value1')
                if bs_dim:
                    curr_sub['bs'].append({
                        'Dimension Value1': bs_dim,
                        'Rate': get_val(row_list, 'BS', 'Rate') or '',
                        'Calculation Derivation Record': get_val(row_list, 'BS', 'Calculation Derivation Record') or ''
                    })

        # Reconstruct vertical block section rows matching Upload Format.xlsx layout
        output_rows = []

        for master in master_agreements:
            h_info = master['header']
            a_info = master['agrmt']

            # Request Header
            output_rows.append(['Request Header', None, None, 'Request Description'])
            output_rows.append([None, 'Value', 'HEADER', h_info.get('Request Description', master['ma_id'])])

            # Master Agreement Header
            output_rows.append(['Master Agreement Header', 'Sub Heading ', None, 'Master Agreement External ID', 'Master Agreement Description', 'Company', 'Business Unit', 'Legal Valid From', 'Legal Valid To'])
            output_rows.append([None, 'Value', 'AGRMT', a_info.get('Master Agreement External ID', master['ma_id']), a_info.get('Master Agreement Description', master['ma_id']), a_info.get('Company', ''), a_info.get('Business Unit', ''), a_info.get('Legal Valid From', ''), a_info.get('Legal Valid To', '')])

            for sub_idx, sub in enumerate(master['sub_agreements']):
                s_info = sub['subagrmt']
                seq_num = s_info.get('Processing Sequence', sub_idx + 1)
                sa_id = s_info.get('Sub Agreement External ID', '')
                prog_code = s_info.get('Sub Agreement Program', '')
                
                # Determine format type (Amount vs Quantity) based on program code
                fmt_type = get_program_format_type(prog_code)
                is_quantity = (fmt_type == "Quantity")

                # Sub Agreement Header
                output_rows.append([f'Sub Agreement-{seq_num} Header', 'Sub Heading ', None, 'Processing Sequence', 'Sub Agreement External ID', 'Sub Agreement Description', 'Sub Agreement Program', 'Period Profile', 'Accrual Frequency', 'Settlement Frequency', 'GST Indicator', 'Key for Prov G/L', 'Legal Valid From', 'Legal Valid To'])
                output_rows.append([None, 'Value', 'SUBAGRMT', seq_num, sa_id, s_info.get('Sub Agreement Description', sa_id), prog_code, s_info.get('Period Profile', ''), 'MD', s_info.get('Settlement Frequency', ''), 'N', 'T', s_info.get('Legal Valid From', ''), s_info.get('Legal Valid To', '')])

                # ELCUST
                if sub['elcust']:
                    output_rows.append(['Eligibility of Customer Validity', 'Sub Heading ', None, 'Sub-Agreement External ID', 'Customer External ID', 'Flexible Group', 'Valid From', 'Valid To'])
                    for el in sub['elcust']:
                        output_rows.append(['Can repeat', 'Value', 'ELCUST', ' ', el.get('Customer External ID', ''), None, el.get('Valid From', ''), el.get('Valid To', '')])

                # ELGBBASE
                if sub['elgbbase']:
                    eb_hdrs = get_section_headers(fmt_type, "ELGBBASE")
                    output_rows.append(['Source & base Units ', 'Sub Heading ', None] + eb_hdrs)
                    for eb in sub['elgbbase']:
                        output_rows.append(['Can repeat', 'Value', 'ELGBBASE', sa_id, eb.get('Base Value', ''), eb.get('Valid From', ''), eb.get('Valid To', '')])

                # ELGBASEVAL
                if sub['elgbaseval']:
                    ev_hdrs = get_section_headers(fmt_type, "ELGBASEVAL")
                    output_rows.append(['Source & base Units ', 'Sub Heading ', None] + ev_hdrs)
                    for ev in sub['elgbaseval']:
                        output_rows.append(['Can repeat', 'Value', 'ELGBASEVAL', sa_id, ev.get('Base Value', ''), ev.get('Source', 'MXDS'), ev.get('Source Field', ''), ev.get('Sequence', 1), ev.get('Contribution', '+'), ev.get('Eligibility %', 100), ev.get('Valid From', ''), ev.get('Valid To', '')])

                # ELGBPRDGRP
                if sub['elgbprdgrp']:
                    output_rows.append(['Eligible Product Grouping ', 'Sub Heading ', None, 'Sub-Agreement External ID', 'Group Name', 'Group Type', 'Source', 'Flexible Group', 'Valid From', 'Valid To'])
                    for prd in sub['elgbprdgrp']:
                        output_rows.append(['Can repeat', 'Value', 'ELGBPRDGRP', sa_id, prd.get('Group Name', ''), prd.get('Group Type', 'P'), prd.get('Source', 'MXDS'), prd.get('Flexible Group', ''), prd.get('Valid From', ''), prd.get('Valid To', '')])

                # BENEFITS, BSA, BS
                ben_hdrs = get_section_headers(fmt_type, "BENEFITS")
                
                if len(sub['benefits']) > 1:
                    # Multiple BENEFITS sets (e.g. FLU and MW)
                    num_b = len(sub['benefits'])
                    total_bs = len(sub['bs'])
                    bs_per_b = total_bs // num_b if num_b > 0 else total_bs
                    
                    for b_idx, b in enumerate(sub['benefits']):
                        output_rows.append(['Benefits and respective Scales ', 'Sub Heading ', None] + ben_hdrs)
                        if is_quantity:
                            b_row = [None, 'Value', 'BENEFITS', sa_id, b.get('Payout Group', ''), b.get('Attain Group', ''), b.get('Rate', ''), b.get('Currency(%)', 'INR'), b.get('Target Type', 'XQTY'), b.get('Target Value', ''), b.get('Target Unit', 'PCS'), b.get('Valid From', ''), b.get('Valid To', '')]
                        else:
                            b_row = [None, 'Value', 'BENEFITS', sa_id, b.get('Payout Group', ''), b.get('Attain Group', ''), b.get('Rate', ''), b.get('Target Type', 'XQTY'), b.get('Target Value', ''), b.get('Target Unit', 'PCS'), b.get('Valid From', ''), b.get('Valid To', '')]
                        output_rows.append(b_row)
                        
                        bsa_item = sub['bsa'][b_idx] if b_idx < len(sub['bsa']) else (sub['bsa'][0] if sub['bsa'] else {})
                        output_rows.append(['Scale Attrubute', 'Sub Heading ', None, 'Bracket', 'Scale Type', 'Unit', 'Unit Type'])
                        output_rows.append([None, 'Value', 'BSA', bsa_item.get('Bracket', 'MS'), bsa_item.get('Scale Type', 'A'), bsa_item.get('Unit', '%'), bsa_item.get('Unit Type', 'MXAU')])
                        
                        # Corresponding BS scale rows
                        start_bs = b_idx * bs_per_b
                        end_bs = (b_idx + 1) * bs_per_b if b_idx < num_b - 1 else total_bs
                        sub_bs_chunk = sub['bs'][start_bs:end_bs]
                        
                        if sub_bs_chunk:
                            output_rows.append(['Scales', 'Sub Heading ', None, 'Dimension Value1', 'Rate', 'Calculation Derivation Record'])
                            for bs in sub_bs_chunk:
                                output_rows.append(['Can repeat', 'Value', 'BS', bs.get('Dimension Value1', ''), bs.get('Rate', '')])
                else:
                    if sub['benefits']:
                        output_rows.append(['Benefits and respective Scales ', 'Sub Heading ', None] + ben_hdrs)
                        for b in sub['benefits']:
                            if is_quantity:
                                b_row = [None, 'Value', 'BENEFITS', sa_id, b.get('Payout Group', ''), b.get('Attain Group', ''), b.get('Rate', ''), b.get('Currency(%)', 'INR'), b.get('Target Type', 'XQTY'), b.get('Target Value', ''), b.get('Target Unit', 'PCS'), b.get('Valid From', ''), b.get('Valid To', '')]
                            else:
                                b_row = [None, 'Value', 'BENEFITS', sa_id, b.get('Payout Group', ''), b.get('Attain Group', ''), b.get('Rate', ''), b.get('Target Type', 'XQTY'), b.get('Target Value', ''), b.get('Target Unit', 'PCS'), b.get('Valid From', ''), b.get('Valid To', '')]
                            output_rows.append(b_row)

                    if sub['bsa']:
                        output_rows.append(['Scale Attrubute', 'Sub Heading ', None, 'Bracket', 'Scale Type', 'Unit', 'Unit Type'])
                        for bsa in sub['bsa']:
                            output_rows.append([None, 'Value', 'BSA', bsa.get('Bracket', 'MS'), bsa.get('Scale Type', 'A'), bsa.get('Unit', '%'), bsa.get('Unit Type', 'MXAU')])

                    if sub['bs']:
                        output_rows.append(['Scales', 'Sub Heading ', None, 'Dimension Value1', 'Rate', 'Calculation Derivation Record'])
                        for bs in sub['bs']:
                            output_rows.append(['Can repeat', 'Value', 'BS', bs.get('Dimension Value1', ''), bs.get('Rate', '')])

                # FGHD / FGIT Customer Group
                if sub['fghd_cust']:
                    for fghd_idx, fghd in enumerate(sub['fghd_cust']):
                        output_rows.append(['Eligibility of Customer Group- Flexi Group', 'Sub Heading ', None, 'Description', 'Flexible Group Type', 'Flexible Group', 'Flexible Group Request'])
                        output_rows.append(['Can repeat', 'Value', 'FGHD', fghd.get('Description', ''), ' ', ' ', fghd.get('Flexible Group Request', '')])
                        
                        fg_req = fghd.get('Flexible Group Request')
                        fg_desc = fghd.get('Description')
                        matching_fgits = [
                            f for f in sub['fgit_cust']
                            if (fg_req and f.get('Flexible Group Request') == fg_req) or
                               (fg_desc and f.get('Customer Group') == fg_desc)
                        ]
                        if not matching_fgits and len(sub['fgit_cust']) > fghd_idx:
                            matching_fgits = [sub['fgit_cust'][fghd_idx]]
                            
                        if matching_fgits:
                            fgit_hdr = [None, None, None, 'Flexible Group Category', 'Include/Exclude', 'Material External ID', 'Material Group', 'Capacity', 'Series', 'Star Rating', 'Star Rating Year', 'Feature 1', 'Feature 2', 'Feature 3', 'Feature 4', 'Feature 5', 'Feature 6', 'Feature 7', 'Feature 8', 'Feature 9', 'Feature 10', 'Customer Group', 'Region', 'State', 'Territory/Sales Area', 'Plant', 'Cluster of Customer', 'Liquidation', 'Subset', 'Set Number', 'Flexible Group Request']
                            output_rows.append(fgit_hdr)
                            for fgit in matching_fgits:
                                fgit_row = [None] * 31
                                fgit_row[0] = 'Can repeat'
                                fgit_row[2] = 'FGIT'
                                fgit_row[3] = fgit.get('Flexible Group Category', '')
                                fgit_row[4] = fgit.get('Include/Exclude', 'X')
                                fgit_row[5] = fgit.get('Material External ID', '')
                                fgit_row[6] = fgit.get('Material Group', '')
                                fgit_row[7] = fgit.get('Capacity', '')
                                fgit_row[21] = fgit.get('Customer Group', '')
                                fgit_row[22] = fgit.get('Region', '')
                                fgit_row[23] = fgit.get('State', '')
                                fgit_row[24] = fgit.get('Territory/Sales Area', '')
                                fgit_row[25] = fgit.get('Plant', '')
                                fgit_row[26] = fgit.get('Cluster of Customer', '')
                                fgit_row[28] = fgit.get('Subset', '')
                                fgit_row[29] = fgit.get('Set Number', 1)
                                fgit_row[30] = fgit.get('Flexible Group Request', fg_req if fg_req else '')
                                output_rows.append(fgit_row)

                # FGHD / FGIT Product Group
                if sub['fghd_prd']:
                    for fghd_idx, fghd in enumerate(sub['fghd_prd']):
                        output_rows.append(['Eligible Product Grouping- Flexi Group', 'Sub Heading ', None, 'Description', 'Flexible Group Type', 'Flexible Group', 'Flexible Group Request'])
                        output_rows.append(['Can repeat', 'Value', 'FGHD', fghd.get('Description', ''), fghd.get('Flexible Group Type', 'M7PR'), ' ', fghd.get('Flexible Group Request', '')])
                        
                        fg_req = fghd.get('Flexible Group Request')
                        fg_desc = fghd.get('Description')
                        matching_fgits = [
                            f for f in sub['fgit_prd']
                            if (fg_req and f.get('Flexible Group Request') == fg_req) or
                               (fg_desc and f.get('Material Group') == fg_desc)
                        ]
                        if not matching_fgits and len(sub['fgit_prd']) > fghd_idx:
                            matching_fgits = [sub['fgit_prd'][fghd_idx]]
                            
                        if matching_fgits:
                            fgit_hdr = [None, None, None, 'Flexible Group Category', 'Include/Exclude', 'Material External ID', 'Material Group', 'Capacity', 'Series', 'Star Rating', 'Star Rating Year', 'Feature 1', 'Feature 2', 'Feature 3', 'Feature 4', 'Feature 5', 'Feature 6', 'Feature 7', 'Feature 8', 'Feature 9', 'Feature 10', 'Customer Group', 'Region', 'State', 'Territory/Sales Area', 'Plant', 'Cluster of Customer', 'Liquidation', 'Subset', 'Set Number', 'Flexible Group Request']
                            output_rows.append(fgit_hdr)
                            for fgit in matching_fgits:
                                fgit_row = [None] * 31
                                fgit_row[0] = 'Can repeat'
                                fgit_row[2] = 'FGIT'
                                fgit_row[3] = fgit.get('Flexible Group Category', '')
                                fgit_row[4] = fgit.get('Include/Exclude', 'X')
                                fgit_row[5] = fgit.get('Material External ID', '')
                                fgit_row[6] = fgit.get('Material Group', '')
                                fgit_row[7] = fgit.get('Capacity', '')
                                fgit_row[8] = fgit.get('Series', '')
                                fgit_row[9] = fgit.get('Star Rating', '')
                                fgit_row[10] = fgit.get('Star Rating Year', '')
                                fgit_row[11] = fgit.get('Feature 1', '')
                                fgit_row[12] = fgit.get('Feature 2', '')
                                fgit_row[13] = fgit.get('Feature 3', '')
                                fgit_row[14] = fgit.get('Feature 4', '')
                                fgit_row[15] = fgit.get('Feature 5', '')
                                fgit_row[16] = fgit.get('Feature 6', '')
                                fgit_row[17] = fgit.get('Feature 7', '')
                                fgit_row[18] = fgit.get('Feature 8', '')
                                fgit_row[19] = fgit.get('Feature 9', '')
                                fgit_row[20] = fgit.get('Feature 10', '')
                                fgit_row[21] = fgit.get('Customer Group', '')
                                fgit_row[22] = fgit.get('Region', '')
                                fgit_row[23] = fgit.get('State', '')
                                fgit_row[24] = fgit.get('Territory/Sales Area', '')
                                fgit_row[25] = fgit.get('Plant', '')
                                fgit_row[26] = fgit.get('Cluster of Customer', '')
                                fgit_row[28] = fgit.get('Subset', '')
                                fgit_row[29] = fgit.get('Set Number', 1)
                                fgit_row[30] = fgit.get('Flexible Group Request', fg_req if fg_req else '')
                                output_rows.append(fgit_row)

                # ACCRTE
                if sub['accrte']:
                    acc_hdrs = get_section_headers(fmt_type, "ACCRTE")
                    output_rows.append(['Accrual Rate', 'Sub Heading ', None] + acc_hdrs)
                    for acc in sub['accrte']:
                        if is_quantity:
                            acc_row = ['Accrual Rate', 'Sub Heading ', 'ACCRTE', sa_id, acc.get('Payout Group', ''), acc.get('Rate', ''), acc.get('Currency(%)', 'INR'), acc.get('Valid From', ''), acc.get('Valid To', '')]
                        else:
                            acc_row = ['Accrual Rate', 'Sub Heading ', 'ACCRTE', sa_id, acc.get('Payout Group', ''), acc.get('Rate', ''), acc.get('Valid From', ''), acc.get('Valid To', '')]
                        output_rows.append(acc_row)

                # Sub agreement END delimiter
                output_rows.append([None, 'Value', 'END'])
                # Empty row after sub agreement
                # output_rows.append([])

        cleaned_output_rows = [[clean_cell_value(v) for v in row] for row in output_rows]
        df_reverse = pd.DataFrame(cleaned_output_rows)
        if remove_first_two_cols and df_reverse.shape[1] >= 2:
            df_reverse = df_reverse.iloc[:, 2:]
        df_reverse.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        # Generate CSV output alongside .xlsx
        if total_sheets == 1:
            csv_file = os.path.splitext(output_reverse_file)[0] + '.csv'
        else:
            csv_file = f"{os.path.splitext(output_reverse_file)[0]}_{sheet_name}.csv"
        df_reverse.to_csv(csv_file, index=False, header=False)

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
    base_name = os.path.splitext(output_reverse_file)[0]
    csv_info = f"'{base_name}.csv'" if total_sheets == 1 else f"'{base_name}_*.csv'"
    print(f"Successfully generated '{output_reverse_file}' and {csv_info}.")


if __name__ == '__main__':
    convert_output_to_upload('output1.xlsx', 'reverse-output1.xlsx')
    # convert_output_to_upload('output2.xlsx', 'reverse-output2.xlsx')