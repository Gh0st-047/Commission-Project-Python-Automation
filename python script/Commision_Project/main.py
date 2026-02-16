import pandas as pd
import os
import glob
import datetime
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name
import re

# ==============================================================================
# CONFIGURATION
# ==============================================================================
INPUT_FOLDER = 'Input_Raw'
OUTPUT_FOLDER = 'Output'

os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Commission Logic
PLAN_MAP = {
    1600: {'Weekly': 369.23, 'BiWeekly': 738.46, 'SemiMonthly': 800, 'Monthly': 1600},
    1400: {'Weekly': 323.08, 'BiWeekly': 646.15, 'SemiMonthly': 700, 'Monthly': 1400},
    1200: {'Weekly': 276.92, 'BiWeekly': 553.85, 'SemiMonthly': 600, 'Monthly': 1200},
    1000: {'Weekly': 230.77, 'BiWeekly': 461.54, 'SemiMonthly': 500, 'Monthly': 1000}
}

# ==============================================================================
# 1. LOGIC ENGINE
# ==============================================================================

def extract_date_from_filename(filename):
    """Extract date from filename using various patterns"""
    patterns = [
        r'(\d{1,2})[_-](\d{1,2})[_-](\d{4})',  # MM_DD_YYYY or MM-DD-YYYY
        r'(\d{1,2})\.(\d{1,2})\.(\d{4})',       # MM.DD.YYYY
        r'(\d{2})(\d{2})(\d{4})',               # MMDDYYYY
    ]
    
    for pattern in patterns:
        match = re.search(pattern, filename)
        if match:
            month, day, year = match.groups()
            try:
                return datetime.datetime(int(year), int(month), int(day))
            except:
                continue
    return None

def get_frequency_from_deduction(deduction):
    """Determine payment frequency from deduction amount"""
    deduction = abs(deduction)
    if deduction == 0: 
        return None, None
    
    # Check each frequency with tolerance
    tolerance = 1.0
    
    if any(abs(deduction - v['Weekly']) < tolerance for v in PLAN_MAP.values()):
        return 52, "Weekly"
    if any(abs(deduction - v['BiWeekly']) < tolerance for v in PLAN_MAP.values()):
        return 26, "BiWeekly"
    if any(abs(deduction - v['SemiMonthly']) < tolerance for v in PLAN_MAP.values()):
        return 24, "SemiMonthly"
    if any(abs(deduction - v['Monthly']) < tolerance for v in PLAN_MAP.values()):
        return 12, "Monthly"
        
    return None, "Unknown"

def process_raw_files():
    """Process all CSV/Excel files from Input_Raw folder"""
    files = glob.glob(os.path.join(INPUT_FOLDER, '*.*'))
    valid_files = [f for f in files if f.endswith(('.csv', '.xlsx', '.xls'))]
    
    if not valid_files:
        print("⚠️ No files in Input_Raw!")
        return []

    processed = []
    
    for filepath in valid_files:
        try:
            # Read file
            if filepath.endswith('.csv'):
                df = pd.read_csv(filepath, dtype=str)
            else:
                df = pd.read_excel(filepath, dtype=str)
        except Exception as e:
            print(f"❌ Skipping {filepath}: {e}")
            continue
            
        df.columns = df.columns.str.strip()
        
        # Find required columns
        ded_col = next((c for c in df.columns if 'ppc' in c.lower() and '125' in c.lower()), None)
        date_col = next((c for c in df.columns if 'date' in c.lower()), None)
        id_col = 'SSN' if 'SSN' in df.columns else df.columns[0]
        
        if not ded_col:
            print(f"⚠️ Skipping {filepath}: No PPC125 column found")
            continue

        df = df.dropna(subset=[id_col])
        
        # Clean deduction column - handle both positive and negative values
        df[ded_col] = df[ded_col].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False)
        df[ded_col] = pd.to_numeric(df[ded_col], errors='coerce').fillna(0)
        
        # Determine frequency from sample data
        freq = 52
        freq_name = "Weekly"
        sample = df[df[ded_col] != 0].head(20)
        
        for val in sample[ded_col]:
            f, n = get_frequency_from_deduction(val)
            if f:
                freq, freq_name = f, n
                break
        
        # NEW: Check if the file contains multiple dates
        # Parse dates in the date column if it exists
        if date_col:
            df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
            unique_dates = df[date_col].dropna().unique()
            
            # If we have more than one unique date, split the file by date
            if len(unique_dates) > 1:
                print(f"📦 Multi-week file detected: {os.path.basename(filepath)} ({len(unique_dates)} weeks)")
                
                # Group by date and create separate packets
                for check_date in sorted(unique_dates):
                    week_df = df[df[date_col] == check_date].copy()
                    
                    if not week_df.empty:
                        processed.append({
                            'df': week_df,
                            'date': pd.Timestamp(check_date).to_pydatetime(),
                            'freq': freq,
                            'freq_name': freq_name,
                            'ded_col': ded_col,
                            'id_col': id_col,
                            'date_col': date_col,
                            'filename': f"{os.path.basename(filepath)} ({pd.Timestamp(check_date).strftime('%m/%d/%Y')})"
                        })
                continue  # Skip the single-file processing below
        
        # OLD: Single date file (original behavior)
        check_date = None
        if date_col:
            check_date = pd.to_datetime(df[date_col], errors='coerce').max()
        
        if pd.isna(check_date) or check_date is None:
            check_date = extract_date_from_filename(os.path.basename(filepath))
        
        if check_date is None:
            check_date = datetime.datetime.now()
            print(f"⚠️ No date found for {filepath}, using current date")
        
        processed.append({
            'df': df,
            'date': check_date,
            'freq': freq,
            'freq_name': freq_name,
            'ded_col': ded_col,
            'id_col': id_col,
            'date_col': date_col,
            'filename': os.path.basename(filepath)
        })
        
    # Sort by date (oldest first)
    processed.sort(key=lambda x: x['date'])
    return processed

# ==============================================================================
# 2. EXCEL BUILDER - CLIENT FORMAT
# ==============================================================================

def build_full_report(packets):
    """Build Excel report matching client requirements"""
    if not packets: 
        print("❌ No valid data found.")
        return

    report_date = packets[0]['date']
    freq_name = packets[0]['freq_name']
    freq_val = packets[0]['freq'] if packets[0]['freq'] else 52
    
    filename = f"Commission_Report_Output.xlsx"
    out_path = os.path.join(OUTPUT_FOLDER, filename)
    
    workbook = xlsxwriter.Workbook(out_path, {'nan_inf_to_errors': True})
    
    # Formats
    fmt_header = workbook.add_format({
        'bold': True, 
        'bg_color': '#D9E1F2', 
        'border': 1, 
        'align': 'center',
        'valign': 'vcenter'
    })
    fmt_currency = workbook.add_format({'num_format': '$#,##0.00'})
    fmt_text = workbook.add_format({'num_format': '@'})
    fmt_date_header = workbook.add_format({
        'bold': True,
        'bg_color': '#4472C4',
        'font_color': '#FFFFFF',
        'border': 1,
        'align': 'center'
    })
    
    # Color formats for commission columns
    fmt_charles = workbook.add_format({'num_format': '$#,##0.00', 'bg_color': '#D9E1F2'})
    fmt_harry = workbook.add_format({'num_format': '$#,##0.00', 'bg_color': '#E2EFDA'})
    fmt_lighthouse = workbook.add_format({'num_format': '$#,##0.00', 'bg_color': '#FCE4D6'})
    
    # Grand totals
    fmt_total_header = workbook.add_format({
        'bold': True,
        'bg_color': '#000000',
        'font_color': '#FFFFFF',
        'align': 'center',
        'border': 1
    })
    fmt_total_value = workbook.add_format({
        'num_format': '$#,##0.00',
        'bold': True,
        'bg_color': '#FFFF00',
        'border': 1,
        'align': 'center',
        'font_size': 12
    })

    master_ssn = set()
    unpaid_data = []
    
    # Track payment status for each employee across all weeks
    employee_payments = {}  # {ssn: [week1_amount, week2_amount, ...]}
    
    # Step 1: Create date-named tabs (e.g., "12.7", "12.14", "12.25")
    for i, p in enumerate(packets):
        tab_date = f"{p['date'].month}.{p['date'].day}"  # Format: "12.7" (Windows compatible)
        # Excel tab names can't exceed 31 chars
        tab_name = tab_date[:31]
        
        ws = workbook.add_worksheet(tab_name)
        
        df = p['df']
        ded_col = p['ded_col']
        id_col = p['id_col']
        date_col = p['date_col']
        
        paid = df[df[ded_col] != 0].copy()
        unpaid = df[df[ded_col] == 0].copy()
        
        # Collect all unique SSNs from this file
        all_ids = df[id_col].dropna().astype(str).str.strip().unique()
        master_ssn.update(all_ids)
        
        # Track who paid and who didn't in this week
        paid_ssns = set(paid[id_col].dropna().astype(str).str.strip())
        
        for ssn in all_ids:
            if ssn not in employee_payments:
                employee_payments[ssn] = []
            
            # Check if this employee paid in this week
            if ssn in paid_ssns:
                amount = paid[paid[id_col].astype(str).str.strip() == ssn][ded_col].iloc[0]
                employee_payments[ssn].append(amount)
            else:
                employee_payments[ssn].append(0)  # Missing payment
        
        # Write headers: SSN | PPC125 | MM/DD/YYYY
        ws.write(0, 0, "SSN", fmt_header)
        ws.write(0, 1, "PPC125", fmt_header)
        ws.write(0, 2, p['date'].strftime('%m/%d/%Y'), fmt_header)
        
        ws.set_column(0, 0, 15)
        ws.set_column(1, 2, 12)
        
        # Write only employees who appear in THIS specific file
        # (not all employees from master_ssn)
        all_ssns_in_file = df[id_col].dropna().astype(str).str.strip().unique()
        
        row_idx = 1
        for ssn in sorted(all_ssns_in_file):
            ws.write_string(row_idx, 0, ssn, fmt_text)
            
            # Check if this employee paid in this week
            employee_row = paid[paid[id_col].astype(str).str.strip() == ssn]
            
            if not employee_row.empty:
                # Employee paid - write the deduction
                val = -abs(employee_row[ded_col].iloc[0])
                ws.write_number(row_idx, 1, val, fmt_currency)
            else:
                # Employee didn't pay - leave blank
                ws.write_string(row_idx, 1, "", fmt_text)
            
            row_idx += 1
        
        # Add total row at the bottom of each weekly tab
        ws.write_string(row_idx, 0, "", fmt_text)  # Empty SSN cell
        ws.write_formula(row_idx, 1, f'=SUM(B2:B{row_idx})', fmt_currency)
    
    # Identify perfect employees (paid in ALL weeks) vs imperfect (missing ANY week)
    perfect_employees = []
    imperfect_employees = []
    
    num_weeks = len(packets)
    
    # Also track plan level for sorting (based on first payment)
    employee_plan_levels = {}  # {ssn: plan_number}
    
    for ssn, payments in employee_payments.items():
        # Check if employee paid in all weeks
        if len(payments) == num_weeks and all(amount != 0 for amount in payments):
            perfect_employees.append(ssn)
            
            # Determine plan level from first non-zero payment for sorting
            first_payment = abs(payments[0])
            if freq_name == "Weekly":
                if first_payment >= 360: employee_plan_levels[ssn] = 1600
                elif first_payment >= 315: employee_plan_levels[ssn] = 1400
                elif first_payment >= 270: employee_plan_levels[ssn] = 1200
                elif first_payment >= 220: employee_plan_levels[ssn] = 1000
                else: employee_plan_levels[ssn] = 0
            elif freq_name == "BiWeekly":
                if first_payment >= 720: employee_plan_levels[ssn] = 1600
                elif first_payment >= 630: employee_plan_levels[ssn] = 1400
                elif first_payment >= 540: employee_plan_levels[ssn] = 1200
                elif first_payment >= 450: employee_plan_levels[ssn] = 1000
                else: employee_plan_levels[ssn] = 0
            elif freq_name == "SemiMonthly":
                if first_payment >= 780: employee_plan_levels[ssn] = 1600
                elif first_payment >= 680: employee_plan_levels[ssn] = 1400
                elif first_payment >= 580: employee_plan_levels[ssn] = 1200
                elif first_payment >= 480: employee_plan_levels[ssn] = 1000
                else: employee_plan_levels[ssn] = 0
            else:  # Monthly
                if first_payment >= 1550: employee_plan_levels[ssn] = 1600
                elif first_payment >= 1350: employee_plan_levels[ssn] = 1400
                elif first_payment >= 1150: employee_plan_levels[ssn] = 1200
                elif first_payment >= 950: employee_plan_levels[ssn] = 1000
                else: employee_plan_levels[ssn] = 0
        else:
            # Find which weeks they missed
            missed_weeks = []
            for i, amount in enumerate(payments):
                if amount == 0:
                    missed_weeks.append(packets[i]['date'].strftime('%m/%d/%Y'))
            
            # Add to unpaid with reason
            reason = f"Missing payment in week(s): {', '.join(missed_weeks)}" if missed_weeks else "Incomplete data"
            imperfect_employees.append([ssn, 0, reason])
    
    # Step 2: Create Unpaid tab with imperfect employees (same structure as Commissions)
    ws_unpaid = workbook.add_worksheet("Unpaid")
    
    # Same header structure as Commissions tab
    ws_unpaid.write(0, 0, "SSN", fmt_header)
    ws_unpaid.set_column(0, 0, 15)
    
    # Build same column structure as Commissions
    current_col = 1
    unpaid_ppc_cols = []
    unpaid_plan_cols = []
    unpaid_charles_cols = []
    unpaid_harry_cols = []
    unpaid_lighthouse_cols = []
    
    for i, p in enumerate(packets):
        date_display = p['date'].strftime('%m/%d/%Y')
        
        # Date header spanning 5 columns
        ws_unpaid.merge_range(0, current_col, 0, current_col + 4, date_display, fmt_date_header)
        
        # Sub-headers
        ws_unpaid.write(1, current_col, "PPC125", fmt_header)
        ws_unpaid.write(1, current_col + 1, "Plan", fmt_header)
        ws_unpaid.write(1, current_col + 2, "Charles", fmt_header)
        ws_unpaid.write(1, current_col + 3, "Harry", fmt_header)
        ws_unpaid.write(1, current_col + 4, "LightHouse", fmt_header)
        
        ws_unpaid.set_column(current_col, current_col, 12)
        ws_unpaid.set_column(current_col + 1, current_col + 1, 12)
        ws_unpaid.set_column(current_col + 2, current_col + 4, 11)
        
        # Track column indices
        unpaid_ppc_cols.append(current_col)
        unpaid_plan_cols.append(current_col + 1)
        unpaid_charles_cols.append(current_col + 2)
        unpaid_harry_cols.append(current_col + 3)
        unpaid_lighthouse_cols.append(current_col + 4)
        
        current_col += 5
    
    # Write imperfect employees data
    sorted_imperfect = sorted([emp[0] for emp in imperfect_employees])
    
    for row_num, ssn in enumerate(sorted_imperfect):
        excel_row = row_num + 2
        ws_unpaid.write_string(row_num + 2, 0, ssn, fmt_text)
        
        for i, p in enumerate(packets):
            tab_date = f"{p['date'].month}.{p['date'].day}"
            
            # PPC125 VLOOKUP - will show 0 if not found
            vlookup = f'=IFERROR(VLOOKUP($A{excel_row+1},\'{tab_date}\'!A:B,2,FALSE),0)'
            ws_unpaid.write_formula(row_num + 2, unpaid_ppc_cols[i], vlookup, fmt_currency)
            
            ppc_cell = xl_rowcol_to_cell(row_num + 2, unpaid_ppc_cols[i])
            
            # Plan determination
            if freq_name == "Weekly":
                plan_formula = f'=IF(ABS({ppc_cell})>=360,"Plan 1600",IF(ABS({ppc_cell})>=315,"Plan 1400",IF(ABS({ppc_cell})>=270,"Plan 1200",IF(ABS({ppc_cell})>=220,"Plan 1000",""))))'
            elif freq_name == "BiWeekly":
                plan_formula = f'=IF(ABS({ppc_cell})>=720,"Plan 1600",IF(ABS({ppc_cell})>=630,"Plan 1400",IF(ABS({ppc_cell})>=540,"Plan 1200",IF(ABS({ppc_cell})>=450,"Plan 1000",""))))'
            elif freq_name == "SemiMonthly":
                plan_formula = f'=IF(ABS({ppc_cell})>=780,"Plan 1600",IF(ABS({ppc_cell})>=680,"Plan 1400",IF(ABS({ppc_cell})>=580,"Plan 1200",IF(ABS({ppc_cell})>=480,"Plan 1000",""))))'
            else:  # Monthly
                plan_formula = f'=IF(ABS({ppc_cell})>=1550,"Plan 1600",IF(ABS({ppc_cell})>=1350,"Plan 1400",IF(ABS({ppc_cell})>=1150,"Plan 1200",IF(ABS({ppc_cell})>=950,"Plan 1000",""))))'
            
            ws_unpaid.write_formula(row_num + 2, unpaid_plan_cols[i], plan_formula)
            
            plan_cell = xl_rowcol_to_cell(row_num + 2, unpaid_plan_cols[i])
            
            # Commission formulas (will be 0 when plan is empty)
            charles_formula = f'=IF({plan_cell}="Plan 1600",15*12/{freq_val},IF({plan_cell}="Plan 1400",10*12/{freq_val},IF({plan_cell}="Plan 1200",5*12/{freq_val},IF({plan_cell}="Plan 1000",1.5*12/{freq_val},0))))'
            harry_formula = f'=IF({plan_cell}="Plan 1600",97*12/{freq_val},IF({plan_cell}="Plan 1400",78*12/{freq_val},IF({plan_cell}="Plan 1200",60*12/{freq_val},IF({plan_cell}="Plan 1000",25*12/{freq_val},0))))'
            lighthouse_formula = f'=IF({plan_cell}="Plan 1600",25*12/{freq_val},IF({plan_cell}="Plan 1400",20*12/{freq_val},IF({plan_cell}="Plan 1200",15*12/{freq_val},IF({plan_cell}="Plan 1000",2*12/{freq_val},0))))'
            
            ws_unpaid.write_formula(row_num + 2, unpaid_charles_cols[i], charles_formula, fmt_charles)
            ws_unpaid.write_formula(row_num + 2, unpaid_harry_cols[i], harry_formula, fmt_harry)
            ws_unpaid.write_formula(row_num + 2, unpaid_lighthouse_cols[i], lighthouse_formula, fmt_lighthouse)
    
    # Step 3: Create Commissions Dashboard (MUST BE FIRST TAB)
    ws_comm = workbook.add_worksheet("Commissions")
    # Move Commissions to first position, Unpaid to second
    workbook.worksheets_objs.insert(0, workbook.worksheets_objs.pop())  # Move Commissions to position 0
    workbook.worksheets_objs.insert(1, workbook.worksheets_objs.pop(-1))  # Move Unpaid to position 1
    
    ws_comm.freeze_panes(1, 1)
    ws_comm.write(0, 0, "SSN", fmt_header)
    ws_comm.set_column(0, 0, 15)
    
    # Build column structure
    current_col = 1
    ppc_cols = []
    plan_cols = []
    charles_cols = []
    harry_cols = []
    lighthouse_cols = []
    
    for i, p in enumerate(packets):
        tab_date = f"{p['date'].month}.{p['date'].day}"  # Windows compatible
        date_display = p['date'].strftime('%m/%d/%Y')
        
        # Date header spanning 5 columns
        ws_comm.merge_range(0, current_col, 0, current_col + 4, date_display, fmt_date_header)
        
        # Sub-headers
        ws_comm.write(1, current_col, "PPC125", fmt_header)
        ws_comm.write(1, current_col + 1, "Plan", fmt_header)
        ws_comm.write(1, current_col + 2, "Charles", fmt_header)
        ws_comm.write(1, current_col + 3, "Harry", fmt_header)
        ws_comm.write(1, current_col + 4, "LightHouse", fmt_header)
        
        ws_comm.set_column(current_col, current_col, 12)
        ws_comm.set_column(current_col + 1, current_col + 1, 12)
        ws_comm.set_column(current_col + 2, current_col + 4, 11)
        
        # Track column indices
        ppc_cols.append(current_col)
        plan_cols.append(current_col + 1)
        charles_cols.append(current_col + 2)
        harry_cols.append(current_col + 3)
        lighthouse_cols.append(current_col + 4)
        
        current_col += 5
    
    # Write SSN data - ONLY PERFECT EMPLOYEES, sorted by plan level
    # Sort by plan (1600→1400→1200→1000), then alphabetically within each plan
    sorted_ssns = sorted(perfect_employees, key=lambda ssn: (-employee_plan_levels.get(ssn, 0), ssn))
    
    for row_num, ssn in enumerate(sorted_ssns):
        excel_row = row_num + 2  # Row 2 onwards (headers are rows 0-1)
        ws_comm.write_string(row_num + 2, 0, ssn, fmt_text)
        
        for i, p in enumerate(packets):
            tab_date = f"{p['date'].month}.{p['date'].day}"  # Windows compatible
            
            # PPC125 VLOOKUP
            vlookup = f'=IFERROR(VLOOKUP($A{excel_row+1},\'{tab_date}\'!A:B,2,FALSE),0)'
            ws_comm.write_formula(row_num + 2, ppc_cols[i], vlookup, fmt_currency)
            
            ppc_cell = xl_rowcol_to_cell(row_num + 2, ppc_cols[i])
            
            # Plan determination - check actual PPC125 value against expected amounts
            # More accurate than rounding annualized values
            if freq_name == "Weekly":
                plan_formula = f'=IF(ABS({ppc_cell})>=360,"Plan 1600",IF(ABS({ppc_cell})>=315,"Plan 1400",IF(ABS({ppc_cell})>=270,"Plan 1200",IF(ABS({ppc_cell})>=220,"Plan 1000",""))))'
            elif freq_name == "BiWeekly":
                plan_formula = f'=IF(ABS({ppc_cell})>=720,"Plan 1600",IF(ABS({ppc_cell})>=630,"Plan 1400",IF(ABS({ppc_cell})>=540,"Plan 1200",IF(ABS({ppc_cell})>=450,"Plan 1000",""))))'
            elif freq_name == "SemiMonthly":
                plan_formula = f'=IF(ABS({ppc_cell})>=780,"Plan 1600",IF(ABS({ppc_cell})>=680,"Plan 1400",IF(ABS({ppc_cell})>=580,"Plan 1200",IF(ABS({ppc_cell})>=480,"Plan 1000",""))))'
            else:  # Monthly
                plan_formula = f'=IF(ABS({ppc_cell})>=1550,"Plan 1600",IF(ABS({ppc_cell})>=1350,"Plan 1400",IF(ABS({ppc_cell})>=1150,"Plan 1200",IF(ABS({ppc_cell})>=950,"Plan 1000",""))))'
            
            ws_comm.write_formula(row_num + 2, plan_cols[i], plan_formula)
            
            plan_cell = xl_rowcol_to_cell(row_num + 2, plan_cols[i])
            
            # Commission formulas using nested IF (compatible with all Excel versions)
            charles_formula = f'=IF({plan_cell}="Plan 1600",15*12/{freq_val},IF({plan_cell}="Plan 1400",10*12/{freq_val},IF({plan_cell}="Plan 1200",5*12/{freq_val},IF({plan_cell}="Plan 1000",1.5*12/{freq_val},0))))'
            harry_formula = f'=IF({plan_cell}="Plan 1600",97*12/{freq_val},IF({plan_cell}="Plan 1400",78*12/{freq_val},IF({plan_cell}="Plan 1200",60*12/{freq_val},IF({plan_cell}="Plan 1000",25*12/{freq_val},0))))'
            lighthouse_formula = f'=IF({plan_cell}="Plan 1600",25*12/{freq_val},IF({plan_cell}="Plan 1400",20*12/{freq_val},IF({plan_cell}="Plan 1200",15*12/{freq_val},IF({plan_cell}="Plan 1000",2*12/{freq_val},0))))'
            
            ws_comm.write_formula(row_num + 2, charles_cols[i], charles_formula, fmt_charles)
            ws_comm.write_formula(row_num + 2, harry_cols[i], harry_formula, fmt_harry)
            ws_comm.write_formula(row_num + 2, lighthouse_cols[i], lighthouse_formula, fmt_lighthouse)
    
    last_data_row = len(sorted_ssns) + 2
    
    # Add weekly subtotals row
    subtotal_row = last_data_row + 2  # Leave one blank row
    
    ws_comm.write(subtotal_row, 0, "Weekly Totals", fmt_total_header)
    
    # Format for weekly subtotals
    fmt_weekly_total = workbook.add_format({
        'num_format': '$#,##0.00',
        'bold': True,
        'bg_color': '#FFE699',
        'border': 1,
        'top': 2
    })
    
    for i in range(len(packets)):
        # Skip PPC125 and Plan columns, only total the commission columns
        charles_col = charles_cols[i]
        harry_col = harry_cols[i]
        lighthouse_col = lighthouse_cols[i]
        
        # Write sum formulas for each agent in this week
        ws_comm.write_formula(subtotal_row, charles_col, 
            f'=SUM({xl_col_to_name(charles_col)}3:{xl_col_to_name(charles_col)}{last_data_row})',
            fmt_weekly_total)
        ws_comm.write_formula(subtotal_row, harry_col,
            f'=SUM({xl_col_to_name(harry_col)}3:{xl_col_to_name(harry_col)}{last_data_row})',
            fmt_weekly_total)
        ws_comm.write_formula(subtotal_row, lighthouse_col,
            f'=SUM({xl_col_to_name(lighthouse_col)}3:{xl_col_to_name(lighthouse_col)}{last_data_row})',
            fmt_weekly_total)
    
    # Step 4: Grand Totals (right side)
    totals_col = current_col + 1
    
    ws_comm.write(0, totals_col, "GRAND TOTALS", fmt_total_header)
    ws_comm.write(1, totals_col, "Charles", fmt_total_header)
    ws_comm.write(1, totals_col + 1, "Harry", fmt_total_header)
    ws_comm.write(1, totals_col + 2, "LightHouse", fmt_total_header)
    
    def build_sum_formula(cols):
        ranges = []
        for c in cols:
            col_letter = xl_col_to_name(c)
            ranges.append(f"{col_letter}3:{col_letter}{last_data_row}")  # Sum only employee rows, not weekly totals
        return f"=SUM({','.join(ranges)})"
    
    ws_comm.write_formula(2, totals_col, build_sum_formula(charles_cols), fmt_total_value)
    ws_comm.write_formula(2, totals_col + 1, build_sum_formula(harry_cols), fmt_total_value)
    ws_comm.write_formula(2, totals_col + 2, build_sum_formula(lighthouse_cols), fmt_total_value)
    
    ws_comm.set_column(totals_col, totals_col + 2, 18)
    
    workbook.close()
    print(f"\n✅ REPORT GENERATED: {filename}")
    print(f"📊 Frequency Detected: {freq_name} (÷{freq_val})")
    print(f"📅 Date Range: {packets[0]['date'].strftime('%m/%d/%Y')} - {packets[-1]['date'].strftime('%m/%d/%Y')}")
    print(f"👥 Total Employees: {len(master_ssn)}")
    print(f"✅ Perfect Employees (paid all weeks): {len(perfect_employees)}")
    print(f"❌ Imperfect Employees (moved to Unpaid): {len(imperfect_employees)}")
    print(f"📄 Files Processed: {len(packets)}")

if __name__ == "__main__":
    print("=" * 60)
    print("COMMISSION REPORT GENERATOR")
    print("=" * 60)
    packets = process_raw_files()
    build_full_report(packets)