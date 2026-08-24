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

        # Parse output data rows into Master Agreement & Sub Agreement objects
        master_agreements = []
        curr_master = None
        curr_sub = None

        for idx, row in data_rows.iterrows():
            row_list = list(row)
            
            # Check Master Agreement ID in Col 1
            ma_id = str(row_list[1]).strip() if pd.notna(row_list[1]) else None
            # Check Sub Agreement ID in Col 8 or Processing Sequence in Col 7
            sa_id = str(row_list[8]).strip() if pd.notna(row_list[8]) else None
            proc_seq = row_list[7] if pd.notna(row_list[7]) else None

            if not curr_master or (ma_id and curr_master['ma_id'] != ma_id):
                curr_master = {
                    'ma_id': ma_id,
                    'header': {'Request Description': row_list[0] if pd.notna(row_list[0]) else ma_id},
                    'agrmt': {
                        'Master Agreement External ID': row_list[1] if pd.notna(row_list[1]) else '',
                        'Master Agreement Description': row_list[2] if pd.notna(row_list[2]) else '',
                        'Company': row_list[3] if pd.notna(row_list[3]) else '',
                        'Business Unit': row_list[4] if pd.notna(row_list[4]) else '',
                        'Legal Valid From': row_list[5] if pd.notna(row_list[5]) else '',
                        'Legal Valid To': row_list[6] if pd.notna(row_list[6]) else ''
                    },
                    'sub_agreements': []
                }
                master_agreements.append(curr_master)

            if sa_id or proc_seq or not curr_master['sub_agreements']:
                curr_sub = {
                    'subagrmt': {
                        'Processing Sequence': proc_seq if pd.notna(proc_seq) else len(curr_master['sub_agreements']) + 1,
                        'Sub Agreement External ID': sa_id if sa_id else '',
                        'Sub Agreement Description': row_list[9] if pd.notna(row_list[9]) else sa_id,
                        'Sub Agreement Program': row_list[10] if pd.notna(row_list[10]) else '',
                        'Period Profile': row_list[13] if pd.notna(row_list[13]) else '',
                        'Settlement Frequency': row_list[14] if pd.notna(row_list[14]) else '',
                        'Legal Valid From': row_list[11] if pd.notna(row_list[11]) else '',
                        'Legal Valid To': row_list[12] if pd.notna(row_list[12]) else ''
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
                if pd.notna(row_list[16]):
                    add_if_unique(curr_sub['elcust'], {
                        'Sub-Agreement External ID': row_list[15] if pd.notna(row_list[15]) else sub_ext_id,
                        'Customer External ID': row_list[16],
                        'Flexible Group': row_list[17] if pd.notna(row_list[17]) else '',
                        'Valid From': row_list[18] if pd.notna(row_list[18]) else '',
                        'Valid To': row_list[19] if pd.notna(row_list[19]) else ''
                    }, 'Customer External ID')

                # ELGBBASE
                if pd.notna(row_list[57]):
                    add_if_unique(curr_sub['elgbbase'], {
                        'Sub-Agreement External ID': row_list[56] if pd.notna(row_list[56]) else sub_ext_id,
                        'Base Value': row_list[57],
                        'Valid From': row_list[11] if pd.notna(row_list[11]) else '',
                        'Valid To': row_list[12] if pd.notna(row_list[12]) else ''
                    }, 'Base Value')

                # ELGBASEVAL
                if pd.notna(row_list[59]):
                    add_if_unique(curr_sub['elgbaseval'], {
                        'Sub-Agreement External ID': row_list[58] if pd.notna(row_list[58]) else sub_ext_id,
                        'Base Value': row_list[59],
                        'Source': row_list[60] if pd.notna(row_list[60]) else '',
                        'Source Field': row_list[61] if pd.notna(row_list[61]) else '',
                        'Sequence': row_list[62] if pd.notna(row_list[62]) else 1,
                        'Contribution': row_list[63] if pd.notna(row_list[63]) else '+',
                        'Eligibility %': row_list[64] if pd.notna(row_list[64]) else 100,
                        'Valid From': row_list[11] if pd.notna(row_list[11]) else '',
                        'Valid To': row_list[12] if pd.notna(row_list[12]) else ''
                    }, 'Base Value')

                # FGHD Cust
                if pd.notna(row_list[20]):
                    add_if_unique(curr_sub['fghd_cust'], {
                        'Description': row_list[20],
                        'Object Type': row_list[21] if pd.notna(row_list[21]) else '',
                        'Local': row_list[22] if pd.notna(row_list[22]) else '',
                        'Flexible Group Type': row_list[23] if pd.notna(row_list[23]) else '',
                        'Flexible Group Request': row_list[24] if pd.notna(row_list[24]) else '',
                        'Valid From': row_list[25] if pd.notna(row_list[25]) else '',
                        'Valid To': row_list[26] if pd.notna(row_list[26]) else ''
                    }, 'Description')

                # FGIT Cust
                if any(pd.notna(row_list[c]) for c in range(27, 55)):
                    fgit_c_item = {
                        'Flexible Group Category': row_list[27] if pd.notna(row_list[27]) else '',
                        'Include/Exclude': row_list[28] if pd.notna(row_list[28]) else 'X',
                        'Material External ID': row_list[29] if pd.notna(row_list[29]) else '',
                        'Material Group': row_list[31] if pd.notna(row_list[31]) else '',
                        'Capacity': row_list[32] if pd.notna(row_list[32]) else '',
                        'Customer Group': row_list[46] if pd.notna(row_list[46]) else '',
                        'Region': row_list[47] if pd.notna(row_list[47]) else '',
                        'State': row_list[48] if pd.notna(row_list[48]) else '',
                        'Territory/Sales Area': row_list[49] if pd.notna(row_list[49]) else '',
                        'Plant': row_list[50] if pd.notna(row_list[50]) else '',
                        'Cluster of Customer': row_list[51] if pd.notna(row_list[51]) else '',
                        'Subset': row_list[52] if pd.notna(row_list[52]) else '',
                        'Set Number': row_list[53] if pd.notna(row_list[53]) else 1,
                        'Flexible Group Request': row_list[54] if pd.notna(row_list[54]) else ''
                    }
                    if fgit_c_item not in curr_sub['fgit_cust']:
                        curr_sub['fgit_cust'].append(fgit_c_item)

                # ELGBPRDGRP
                if pd.notna(row_list[66]):
                    add_if_unique(curr_sub['elgbprdgrp'], {
                        'Sub-Agreement External ID': row_list[65] if pd.notna(row_list[65]) else sub_ext_id,
                        'Group Name': row_list[66],
                        'Group Type': row_list[67] if pd.notna(row_list[67]) else '',
                        'Source': row_list[68] if pd.notna(row_list[68]) else '',
                        'Flexible Group': row_list[69] if pd.notna(row_list[69]) else '',
                        'Valid From': row_list[70] if pd.notna(row_list[70]) else '',
                        'Valid To': row_list[71] if pd.notna(row_list[71]) else ''
                    }, 'Group Name')

                # FGHD Prd
                if pd.notna(row_list[72]):
                    add_if_unique(curr_sub['fghd_prd'], {
                        'Description': row_list[72],
                        'Object Type': row_list[73] if pd.notna(row_list[73]) else '',
                        'Local': row_list[74] if pd.notna(row_list[74]) else '',
                        'Flexible Group Type': row_list[75] if pd.notna(row_list[75]) else '',
                        'Flexible Group Request': row_list[76] if pd.notna(row_list[76]) else '',
                        'Valid From': row_list[77] if pd.notna(row_list[77]) else '',
                        'Valid To': row_list[78] if pd.notna(row_list[78]) else ''
                    }, 'Description')

                # FGIT Prd
                if any(pd.notna(row_list[c]) for c in range(79, 107)):
                    fgit_p_item = {
                        'Flexible Group Category': row_list[79] if pd.notna(row_list[79]) else '',
                        'Include/Exclude': row_list[80] if pd.notna(row_list[80]) else 'X',
                        'Material External ID': row_list[81] if pd.notna(row_list[81]) else '',
                        'Product Hierarchy': row_list[82] if pd.notna(row_list[82]) else '',
                        'Material Group': row_list[83] if pd.notna(row_list[83]) else '',
                        'Capacity': row_list[84] if pd.notna(row_list[84]) else '',
                        'Series': row_list[85] if pd.notna(row_list[85]) else '',
                        'Star Rating': row_list[86] if pd.notna(row_list[86]) else '',
                        'Star Rating Year': row_list[87] if pd.notna(row_list[87]) else '',
                        'Feature 1': row_list[88] if pd.notna(row_list[88]) else '',
                        'Feature 2': row_list[89] if pd.notna(row_list[89]) else '',
                        'Feature 3': row_list[90] if pd.notna(row_list[90]) else '',
                        'Feature 4': row_list[91] if pd.notna(row_list[91]) else '',
                        'Feature 5': row_list[92] if pd.notna(row_list[92]) else '',
                        'Feature 6': row_list[93] if pd.notna(row_list[93]) else '',
                        'Feature 7': row_list[94] if pd.notna(row_list[94]) else '',
                        'Feature 8': row_list[95] if pd.notna(row_list[95]) else '',
                        'Feature 9': row_list[96] if pd.notna(row_list[96]) else '',
                        'Feature 10': row_list[97] if pd.notna(row_list[97]) else '',
                        'Customer Group': row_list[98] if pd.notna(row_list[98]) else '',
                        'Region': row_list[99] if pd.notna(row_list[99]) else '',
                        'State': row_list[100] if pd.notna(row_list[100]) else '',
                        'Territory/Sales Area': row_list[101] if pd.notna(row_list[101]) else '',
                        'Plant': row_list[102] if pd.notna(row_list[102]) else '',
                        'Cluster of Customer': row_list[103] if pd.notna(row_list[103]) else '',
                        'Subset': row_list[104] if pd.notna(row_list[104]) else '',
                        'Set Number': row_list[105] if pd.notna(row_list[105]) else 1,
                        'Flexible Group Request': row_list[106] if pd.notna(row_list[106]) else ''
                    }
                    if fgit_p_item not in curr_sub['fgit_prd']:
                        curr_sub['fgit_prd'].append(fgit_p_item)

                # BENEFITS
                if pd.notna(row_list[145]):
                    sub_vf = curr_sub['subagrmt'].get('Legal Valid From', '')
                    sub_vt = curr_sub['subagrmt'].get('Legal Valid To', '')
                    add_if_unique(curr_sub['benefits'], {
                        'Sub-Agreement External ID': row_list[144] if pd.notna(row_list[144]) else sub_ext_id,
                        'Payout Group': row_list[145],
                        'Attain Group': '',
                        'Rate': '',
                        'Target Type': row_list[146] if pd.notna(row_list[146]) else '',
                        'Target Value': row_list[147] if pd.notna(row_list[147]) else '',
                        'Currency(%)': row_list[148] if pd.notna(row_list[148]) else '',
                        'Target Unit': row_list[149] if pd.notna(row_list[149]) else '',
                        'Valid From': sub_vf,
                        'Valid To': sub_vt
                    }, 'Payout Group')

                # BSA
                if pd.notna(row_list[151]):
                    add_if_unique(curr_sub['bsa'], {
                        'Bracket': row_list[151],
                        'Scale Type': row_list[152] if pd.notna(row_list[152]) else '',
                        'Unit': row_list[153] if pd.notna(row_list[153]) else '',
                        'Unit Type': row_list[154] if pd.notna(row_list[154]) else ''
                    }, 'Bracket')

                # ACCRTE
                if pd.notna(row_list[171]):
                    add_if_unique(curr_sub['accrte'], {
                        'Sub-Agreement External ID': row_list[170] if pd.notna(row_list[170]) else sub_ext_id,
                        'Payout Group': row_list[171],
                        'Rate': row_list[172] if pd.notna(row_list[172]) else '',
                        'Currency(%)': row_list[148] if pd.notna(row_list[148]) else 'INR',
                        'Valid From': row_list[173] if pd.notna(row_list[173]) else '',
                        'Valid To': row_list[174] if pd.notna(row_list[174]) else ''
                    }, 'Payout Group')

                # BS (Scales)
                if pd.notna(row_list[155]):
                    curr_sub['bs'].append({
                        'Dimension Value1': row_list[155],
                        'Rate': row_list[156] if pd.notna(row_list[156]) else ''
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