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

# Group Types
GROUP_TYPE_HARRY = "Harry's Group"
GROUP_TYPE_ADAM = "Adam's Group"
GROUP_TYPE_OTHER = "Other Groups"
GROUP_TYPE_DYNAMIC = "Dynamic Group"

# Dynamic Groups Storage
DYNAMIC_GROUPS = {}

# Commission Logic - Base PLAN_MAP for frequency detection
PLAN_MAP = {
    1600: {'Weekly': 369.23, 'BiWeekly': 738.46, 'SemiMonthly': 800, 'Monthly': 1600},
    1400: {'Weekly': 323.08, 'BiWeekly': 646.15, 'SemiMonthly': 700, 'Monthly': 1400},
    1200: {'Weekly': 276.92, 'BiWeekly': 553.85, 'SemiMonthly': 600, 'Monthly': 1200},
    1000: {'Weekly': 230.77, 'BiWeekly': 461.54, 'SemiMonthly': 500, 'Monthly': 1000}
}

# ==============================================================================
# SYSTEM 1: HARRY'S GROUP - CLIENT-BASED RATES
# ==============================================================================

# Harry's Downline - Client-based rates
HARRY_DOWNLINE_RATES = {
    'AMERISTAR': {
        'Agent1': {'1600/1400/1200': 35, '1000': 15},
        'Agent2': {'1600/1400/1200': 35, '1000': 15}
    },
    'JANUS': {
        'Agent1': {'1600/1400/1200': 35, '1000': 15},
        'Agent2': {'1600/1400/1200': 35, '1000': 15}
    },
    'CONFIDENCE': {
        'Agent1': {'1600/1400/1200': 15, '1000': 5},
        'Agent2': {'1600/1400/1200': 15, '1000': 5}
    },
    'CRESCENT': {
        'Agent1': {'1600/1400/1200': 15, '1000': 10},
        'Agent2': {'1600/1400/1200': 15, '1000': 10}
    },
    'MEDALLION HC/SPANISH LAKES': {
        'Agent1': {'1600/1400/1200': 20, '1000': 10},
        'Agent2': {'1600/1400/1200': 20, '1000': 10}
    },
    'METROPOLITAN': {
        'Agent1': {'1600/1400/1200': 35, '1000': 15},
        'Agent2': {'1600/1400/1200': 35, '1000': 15}
    }
}

# Confidence special multipliers based on number of weeks
CONFIDENCE_MULTIPLIERS = {
    2: {'1000': 5, 'other': 15},
    3: {'1000': 2.31, 'other': 5},
    4: {'1000': 1.15, 'other': 3.75},
    5: {'1000': 1.15, 'other': 3}
}

# ==============================================================================
# ADAM'S GROUP - BROKER-BASED RATES
# ==============================================================================

# Adam's Group Brokers - Individual plan-based rates
ADAMS_GROUP_AGENTS = {
    'OBouley Light House': {'1600': 15, '1400': 15, '1200': 10, '1000': 5},
    'CBsupport': {'1600': 20, '1400': 13, '1200': 10, '1000': 5.25},
    'ALFRED LEOPOLD': {'1600': 20, '1400': 17, '1200': 15, '1000': 5.25},
    'Adam Charon': {'1600': 82, '1400': 63, '1200': 45, '1000': 13}
}

# ==============================================================================
# SYSTEM 2: OTHER GROUPS - TIER-BASED RATES
# ==============================================================================

# Tier Rates - Extracted from tier.xlsx
TIER_RATES = {
    '70': {
        'PPC1600': 107,
        'PPC1400': 88,
        'PPC1200': 70,
        'PPC1000': 25
    },
    '60': {
        'PPC1600': 97,
        'PPC1400': 78,
        'PPC1200': 60,
        'PPC1000': 25
    },
    '50': {
        'PPC1600': 87,
        'PPC1400': 68,
        'PPC1200': 50,
        'PPC1000': 15
    },
    '45': {
        'PPC1600': 82,
        'PPC1400': 63,
        'PPC1200': 45,
        'PPC1000': 13
    },
    '40': {
        'PPC1600': 77,
        'PPC1400': 58,
        'PPC1200': 40,
        'PPC1000': 12.5
    },
    '35': {
        'PPC1600': 72,
        'PPC1400': 53,
        'PPC1200': 35,
        'PPC1000': 10
    },
    '30': {
        'PPC1600': 52,
        'PPC1400': 37,
        'PPC1200': 30,
        'PPC1000': 8
    },
    '25': {
        'PPC1600': 44,
        'PPC1400': 32,
        'PPC1200': 25,
        'PPC1000': 7.5
    },
    '20': {
        'PPC1600': 37,
        'PPC1400': 27,
        'PPC1200': 20,
        'PPC1000': 6
    },
    '15': {
        'PPC1600': 30,
        'PPC1400': 22,
        'PPC1200': 15,
        'PPC1000': 5
    }
}

# ==============================================================================
# HELPER FUNCTIONS
# ==============================================================================

def get_rate_for_plan(rates, plan):
    """Extract rate for a specific plan from rates dictionary.
    Handles both old format ('1600/1400/1200': value) and new format ('1600': value, etc.)
    """
    if not rates:
        return 0
    
    # New format: individual plan keys like '1600', '1400', etc.
    if plan in rates:
        return rates[plan]
    
    # Old format: grouped key like '1600/1400/1200'
    if '1600/1400/1200' in rates and plan in ['1600', '1400', '1200']:
        return rates['1600/1400/1200']
    
    # Fallback
    return rates.get('1000', 0)

# ==============================================================================
# 1. LOGIC ENGINE - FILE PROCESSING
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
        return None, "Unknown"
    
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

def detect_plan_from_amount(amount, freq_name):
    """Detect plan level from deduction amount"""
    if amount == 0:
        return None
    
    amount = abs(amount)
    tolerance = 1.0
    
    for plan_num, freq_rates in PLAN_MAP.items():
        if abs(amount - freq_rates[freq_name]) < tolerance:
            return f'PPC{plan_num}'
    
    return 'PPC1000'  # Default fallback

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
        
        # Extract date
        check_date = None
        if date_col:
            check_date = pd.to_datetime(df[date_col], errors='coerce').max()
        
        if pd.isna(check_date) or check_date is None:
            check_date = extract_date_from_filename(os.path.basename(filepath))
        
        if check_date is None:
            check_date = datetime.datetime.now()
            print(f"⚠️ No date found for {filepath}, using current date")
        
        # Clean deduction column
        df[ded_col] = df[ded_col].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False)
        df[ded_col] = pd.to_numeric(df[ded_col], errors='coerce').fillna(0)
        
        # Determine frequency
        freq = 52
        freq_name = "Weekly"
        sample = df[df[ded_col] != 0].head(20)
        
        for val in sample[ded_col]:
            f, n = get_frequency_from_deduction(val)
            if f:
                freq, freq_name = f, n
                break
        
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
# 2. TIER-BASED COMMISSION CALCULATIONS (System 2)
# ==============================================================================

def get_employee_plan_counts(packets, perfect_employees, employee_plan_levels):
    """
    Count how many perfect employees are in each plan level
    
    Returns: {'PPC1600': 25, 'PPC1400': 18, 'PPC1200': 12, 'PPC1000': 8}
    """
    plan_counts = {'PPC1600': 0, 'PPC1400': 0, 'PPC1200': 0, 'PPC1000': 0}
    
    for ssn in perfect_employees:
        plan = employee_plan_levels.get(ssn)
        if plan and plan in plan_counts:
            plan_counts[plan] += 1
    
    return plan_counts

def calculate_tier_commission(plan_counts, tier):
    """
    Calculate commission for an agent based on their tier
    
    Args:
        plan_counts: {'PPC1600': 25, 'PPC1400': 18, 'PPC1200': 12, 'PPC1000': 8}
        tier: '70', '60', '50', etc.
    
    Returns:
        Total commission amount
    """
    if tier not in TIER_RATES:
        return 0
    
    rates = TIER_RATES[tier]
    total = 0
    
    for plan, count in plan_counts.items():
        rate = rates.get(plan, 0)
        total += count * rate
    
    return total

def calculate_override_commission(plan_counts, client_tier, agent_tier):
    """
    Calculate override commission (difference between client and agent tier)
    
    Args:
        plan_counts: Employee count by plan
        client_tier: Client's tier (e.g., '70')
        agent_tier: Agent's tier (e.g., '50')
    
    Returns:
        Override commission amount
    """
    if client_tier not in TIER_RATES or agent_tier not in TIER_RATES:
        return 0
    
    client_rates = TIER_RATES[client_tier]
    agent_rates = TIER_RATES[agent_tier]
    
    override = 0
    for plan, count in plan_counts.items():
        client_rate = client_rates.get(plan, 0)
        agent_rate = agent_rates.get(plan, 0)
        override += count * (client_rate - agent_rate)
    
    return override

# ==============================================================================
# 3. HARRY'S GROUP REPORT BUILDER (System 1)
# ==============================================================================

def build_harry_group_report(packets, selected_client=None, group_type=GROUP_TYPE_HARRY):
    """Build Excel report for Harry's Group or Adam's Group with client-based rates, plan counting, and downline commissions"""
    if not packets: 
        print("❌ No valid data found.")
        return

    report_date = packets[0]['date']
    freq_name = packets[0]['freq_name']
    freq_val = packets[0]['freq'] if packets[0]['freq'] else 52
    
    # Include client name in filename if specified
    client_suffix = f"_{selected_client}" if selected_client else ""
    filename = f"Commission_Report_Harry{client_suffix}_{report_date.strftime('%B_%Y')}.xlsx"
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
    employee_payments = {}
    
    # Step 1: Create date-named tabs
    for i, p in enumerate(packets):
        tab_date = f"{p['date'].month}.{p['date'].day}"
        tab_name = tab_date[:31]
        
        ws = workbook.add_worksheet(tab_name)
        
        df = p['df']
        ded_col = p['ded_col']
        id_col = p['id_col']
        
        paid = df[df[ded_col] != 0].copy()
        
        all_ids = df[id_col].dropna().astype(str).str.strip().unique()
        master_ssn.update(all_ids)
        
        paid_ssns = set(paid[id_col].dropna().astype(str).str.strip())
        
        for ssn in all_ids:
            if ssn not in employee_payments:
                employee_payments[ssn] = []
            
            if ssn in paid_ssns:
                amount = paid[paid[id_col].astype(str).str.strip() == ssn][ded_col].iloc[0]
                employee_payments[ssn].append(amount)
            else:
                employee_payments[ssn].append(0)
        
        # Write headers
        ws.write(0, 0, "SSN", fmt_header)
        ws.write(0, 1, "PPC125", fmt_header)
        ws.write(0, 2, p['date'].strftime('%m/%d/%Y'), fmt_header)
        
        ws.set_column(0, 0, 15)
        ws.set_column(1, 2, 12)
        
        all_ssns_in_file = df[id_col].dropna().astype(str).str.strip().unique()
        
        row_idx = 1
        for ssn in sorted(all_ssns_in_file):
            ws.write_string(row_idx, 0, ssn, fmt_text)
            
            employee_row = paid[paid[id_col].astype(str).str.strip() == ssn]
            
            if not employee_row.empty:
                val = -abs(employee_row[ded_col].iloc[0])
                ws.write_number(row_idx, 1, val, fmt_currency)
            else:
                ws.write_string(row_idx, 1, "", fmt_text)
            
            row_idx += 1
        
        # Add total row
        ws.write_string(row_idx, 0, "", fmt_text)
        ws.write_formula(row_idx, 1, f'=SUM(B2:B{row_idx})', fmt_currency)
    
    # Identify perfect vs imperfect employees
    perfect_employees = []
    imperfect_employees = []
    num_weeks = len(packets)
    employee_plan_levels = {}
    
    for ssn, payments in employee_payments.items():
        if len(payments) == num_weeks and all(amount != 0 for amount in payments):
            perfect_employees.append(ssn)
            
            # Determine plan level from first non-zero payment
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
            missed_weeks = []
            for i, amount in enumerate(payments):
                if amount == 0:
                    missed_weeks.append(packets[i]['date'].strftime('%m/%d/%Y'))
            
            reason = f"Missing payment in week(s): {', '.join(missed_weeks)}" if missed_weeks else "Incomplete data"
            imperfect_employees.append([ssn, 0, reason])
    
    # Step 2: Create Unpaid tab
    ws_unpaid = workbook.add_worksheet("Unpaid")
    ws_unpaid.write(0, 0, "SSN", fmt_header)
    ws_unpaid.set_column(0, 0, 15)
    
    current_col = 1
    unpaid_ppc_cols = []
    unpaid_plan_cols = []
    unpaid_charles_cols = []
    unpaid_harry_cols = []
    unpaid_lighthouse_cols = []
    
    for i, p in enumerate(packets):
        date_display = p['date'].strftime('%m/%d/%Y')
        
        ws_unpaid.merge_range(0, current_col, 0, current_col + 4, date_display, fmt_date_header)
        
        ws_unpaid.write(1, current_col, "PPC125", fmt_header)
        ws_unpaid.write(1, current_col + 1, "Plan", fmt_header)
        ws_unpaid.write(1, current_col + 2, "Charles", fmt_header)
        ws_unpaid.write(1, current_col + 3, "Harry", fmt_header)
        ws_unpaid.write(1, current_col + 4, "LightHouse", fmt_header)
        
        ws_unpaid.set_column(current_col, current_col, 12)
        ws_unpaid.set_column(current_col + 1, current_col + 1, 12)
        ws_unpaid.set_column(current_col + 2, current_col + 4, 11)
        
        unpaid_ppc_cols.append(current_col)
        unpaid_plan_cols.append(current_col + 1)
        unpaid_charles_cols.append(current_col + 2)
        unpaid_harry_cols.append(current_col + 3)
        unpaid_lighthouse_cols.append(current_col + 4)
        
        current_col += 5
    
    sorted_imperfect = sorted([emp[0] for emp in imperfect_employees])
    
    for row_num, ssn in enumerate(sorted_imperfect):
        excel_row = row_num + 2
        ws_unpaid.write_string(row_num + 2, 0, ssn, fmt_text)
        
        for i, p in enumerate(packets):
            tab_date = f"{p['date'].month}.{p['date'].day}"
            
            vlookup = f'=IFERROR(VLOOKUP($A{excel_row+1},\'{tab_date}\'!A:B,2,FALSE),0)'
            ws_unpaid.write_formula(row_num + 2, unpaid_ppc_cols[i], vlookup, fmt_currency)
            
            ppc_cell = xl_rowcol_to_cell(row_num + 2, unpaid_ppc_cols[i])
            
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
            
            charles_formula = f'=IF({plan_cell}="Plan 1600",15*12/{freq_val},IF({plan_cell}="Plan 1400",10*12/{freq_val},IF({plan_cell}="Plan 1200",5*12/{freq_val},IF({plan_cell}="Plan 1000",1.5*12/{freq_val},0))))'
            harry_formula = f'=IF({plan_cell}="Plan 1600",97*12/{freq_val},IF({plan_cell}="Plan 1400",78*12/{freq_val},IF({plan_cell}="Plan 1200",60*12/{freq_val},IF({plan_cell}="Plan 1000",25*12/{freq_val},0))))'
            lighthouse_formula = f'=IF({plan_cell}="Plan 1600",25*12/{freq_val},IF({plan_cell}="Plan 1400",20*12/{freq_val},IF({plan_cell}="Plan 1200",15*12/{freq_val},IF({plan_cell}="Plan 1000",2*12/{freq_val},0))))'
            
            ws_unpaid.write_formula(row_num + 2, unpaid_charles_cols[i], charles_formula, fmt_charles)
            ws_unpaid.write_formula(row_num + 2, unpaid_harry_cols[i], harry_formula, fmt_harry)
            ws_unpaid.write_formula(row_num + 2, unpaid_lighthouse_cols[i], lighthouse_formula, fmt_lighthouse)
    
    # Step 3: Create Commissions Dashboard
    ws_comm = workbook.add_worksheet("Commissions")
    workbook.worksheets_objs.insert(0, workbook.worksheets_objs.pop())
    workbook.worksheets_objs.insert(1, workbook.worksheets_objs.pop(-1))
    
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
        tab_date = f"{p['date'].month}.{p['date'].day}"
        date_display = p['date'].strftime('%m/%d/%Y')
        
        ws_comm.merge_range(0, current_col, 0, current_col + 4, date_display, fmt_date_header)
        
        ws_comm.write(1, current_col, "PPC125", fmt_header)
        ws_comm.write(1, current_col + 1, "Plan", fmt_header)
        ws_comm.write(1, current_col + 2, "Charles", fmt_header)
        ws_comm.write(1, current_col + 3, "Harry", fmt_header)
        ws_comm.write(1, current_col + 4, "LightHouse", fmt_header)
        
        ws_comm.set_column(current_col, current_col, 12)
        ws_comm.set_column(current_col + 1, current_col + 1, 12)
        ws_comm.set_column(current_col + 2, current_col + 4, 11)
        
        ppc_cols.append(current_col)
        plan_cols.append(current_col + 1)
        charles_cols.append(current_col + 2)
        harry_cols.append(current_col + 3)
        lighthouse_cols.append(current_col + 4)
        
        current_col += 5
    
    # Write perfect employees
    sorted_ssns = sorted(perfect_employees, key=lambda ssn: (-employee_plan_levels.get(ssn, 0), ssn))
    
    for row_num, ssn in enumerate(sorted_ssns):
        excel_row = row_num + 2
        ws_comm.write_string(row_num + 2, 0, ssn, fmt_text)
        
        for i, p in enumerate(packets):
            tab_date = f"{p['date'].month}.{p['date'].day}"
            
            vlookup = f'=IFERROR(VLOOKUP($A{excel_row+1},\'{tab_date}\'!A:B,2,FALSE),0)'
            ws_comm.write_formula(row_num + 2, ppc_cols[i], vlookup, fmt_currency)
            
            ppc_cell = xl_rowcol_to_cell(row_num + 2, ppc_cols[i])
            
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
            
            charles_formula = f'=IF({plan_cell}="Plan 1600",15*12/{freq_val},IF({plan_cell}="Plan 1400",10*12/{freq_val},IF({plan_cell}="Plan 1200",5*12/{freq_val},IF({plan_cell}="Plan 1000",1.5*12/{freq_val},0))))'
            harry_formula = f'=IF({plan_cell}="Plan 1600",97*12/{freq_val},IF({plan_cell}="Plan 1400",78*12/{freq_val},IF({plan_cell}="Plan 1200",60*12/{freq_val},IF({plan_cell}="Plan 1000",25*12/{freq_val},0))))'
            lighthouse_formula = f'=IF({plan_cell}="Plan 1600",25*12/{freq_val},IF({plan_cell}="Plan 1400",20*12/{freq_val},IF({plan_cell}="Plan 1200",15*12/{freq_val},IF({plan_cell}="Plan 1000",2*12/{freq_val},0))))'
            
            ws_comm.write_formula(row_num + 2, charles_cols[i], charles_formula, fmt_charles)
            ws_comm.write_formula(row_num + 2, harry_cols[i], harry_formula, fmt_harry)
            ws_comm.write_formula(row_num + 2, lighthouse_cols[i], lighthouse_formula, fmt_lighthouse)
    
    last_data_row = len(sorted_ssns) + 2
    subtotal_row = last_data_row + 2
    
    ws_comm.write(subtotal_row, 0, "Weekly Totals", fmt_total_header)
    
    fmt_weekly_total = workbook.add_format({
        'num_format': '$#,##0.00',
        'bold': True,
        'bg_color': '#FFE699',
        'border': 1,
        'top': 2
    })
    
    for i in range(len(packets)):
        charles_col = charles_cols[i]
        harry_col = harry_cols[i]
        lighthouse_col = lighthouse_cols[i]
        
        ws_comm.write_formula(subtotal_row, charles_col, 
            f'=SUM({xl_col_to_name(charles_col)}3:{xl_col_to_name(charles_col)}{last_data_row})',
            fmt_weekly_total)
        ws_comm.write_formula(subtotal_row, harry_col,
            f'=SUM({xl_col_to_name(harry_col)}3:{xl_col_to_name(harry_col)}{last_data_row})',
            fmt_weekly_total)
        ws_comm.write_formula(subtotal_row, lighthouse_col,
            f'=SUM({xl_col_to_name(lighthouse_col)}3:{xl_col_to_name(lighthouse_col)}{last_data_row})',
            fmt_weekly_total)
    
    # Step 4: Grand Totals
    totals_col = current_col + 1
    
    ws_comm.write(0, totals_col, "GRAND TOTALS", fmt_total_header)
    ws_comm.write(1, totals_col, "Charles", fmt_total_header)
    ws_comm.write(1, totals_col + 1, "Harry", fmt_total_header)
    ws_comm.write(1, totals_col + 2, "LightHouse", fmt_total_header)
    
    def build_sum_formula(cols):
        ranges = []
        for c in cols:
            col_letter = xl_col_to_name(c)
            ranges.append(f"{col_letter}3:{col_letter}{last_data_row}")
        return f"=SUM({','.join(ranges)})"
    
    ws_comm.write_formula(2, totals_col, build_sum_formula(charles_cols), fmt_total_value)
    ws_comm.write_formula(2, totals_col + 1, build_sum_formula(harry_cols), fmt_total_value)
    ws_comm.write_formula(2, totals_col + 2, build_sum_formula(lighthouse_cols), fmt_total_value)
    
    ws_comm.set_column(totals_col, totals_col + 2, 18)
    
    # ==============================================================================
    # Step 5: PLAN COUNTING SECTION
    # ==============================================================================
    
    plan_count_start_row = 5
    plan_count_col = totals_col
    
    fmt_plan_count_header = workbook.add_format({
        'bold': True,
        'bg_color': '#FFF2CC',
        'border': 1,
        'align': 'center',
        'font_size': 11
    })
    fmt_plan_count_value = workbook.add_format({
        'bold': True,
        'bg_color': '#FFF2CC',
        'border': 1,
        'align': 'center',
        'font_size': 11
    })
    
    ws_comm.merge_range(plan_count_start_row, plan_count_col, plan_count_start_row, plan_count_col + 2, 
                        "PLAN COUNTING", fmt_plan_count_header)
    
    # Build plan counting formulas based on number of weeks
    if num_weeks == 2:
        ws_comm.merge_range(plan_count_start_row + 1, plan_count_col, plan_count_start_row + 1, plan_count_col + 2,
                           "BiWeekly - 2 Payroll Weeks", fmt_header)
        ws_comm.write(plan_count_start_row + 2, plan_count_col, "Plan 1000 Count:", fmt_plan_count_header)
        ws_comm.write(plan_count_start_row + 3, plan_count_col, "Other Plans Count:", fmt_plan_count_header)
        
        if len(plan_cols) >= 2:
            col_C = xl_col_to_name(plan_cols[0])
            col_I = xl_col_to_name(plan_cols[1])
            
            plan_1000_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_I}3:{col_I}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 2, plan_count_col + 1, plan_1000_formula, fmt_plan_count_value)
            
            other_plans_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_I}3:{col_I}{last_data_row})))=0),--((ISNUMBER(SEARCH("Plan 1200",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_C}3:{col_C}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_I}3:{col_I}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 3, plan_count_col + 1, other_plans_formula, fmt_plan_count_value)
    
    elif num_weeks == 3:
        ws_comm.merge_range(plan_count_start_row + 1, plan_count_col, plan_count_start_row + 1, plan_count_col + 2,
                           "BiWeekly - 3 Payroll Weeks", fmt_header)
        ws_comm.write(plan_count_start_row + 2, plan_count_col, "Plan 1000 Count:", fmt_plan_count_header)
        ws_comm.write(plan_count_start_row + 3, plan_count_col, "Other Plans Count:", fmt_plan_count_header)
        
        if len(plan_cols) >= 3:
            col_C = xl_col_to_name(plan_cols[0])
            col_I = xl_col_to_name(plan_cols[1])
            col_O = xl_col_to_name(plan_cols[2])
            
            plan_1000_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_O}3:{col_O}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 2, plan_count_col + 1, plan_1000_formula, fmt_plan_count_value)
            
            other_plans_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_O}3:{col_O}{last_data_row})))=0),--((ISNUMBER(SEARCH("Plan 1200",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_C}3:{col_C}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_I}3:{col_I}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_O}3:{col_O}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_O}3:{col_O}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_O}3:{col_O}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 3, plan_count_col + 1, other_plans_formula, fmt_plan_count_value)
    
    elif num_weeks == 4:
        ws_comm.merge_range(plan_count_start_row + 1, plan_count_col, plan_count_start_row + 1, plan_count_col + 2,
                           "Weekly - 4 Payroll Weeks", fmt_header)
        ws_comm.write(plan_count_start_row + 2, plan_count_col, "Plan 1000 Count:", fmt_plan_count_header)
        ws_comm.write(plan_count_start_row + 3, plan_count_col, "Other Plans Count:", fmt_plan_count_header)
        
        if len(plan_cols) >= 4:
            col_C = xl_col_to_name(plan_cols[0])
            col_I = xl_col_to_name(plan_cols[1])
            col_O = xl_col_to_name(plan_cols[2])
            col_U = xl_col_to_name(plan_cols[3])
            
            plan_1000_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_O}3:{col_O}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_U}3:{col_U}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 2, plan_count_col + 1, plan_1000_formula, fmt_plan_count_value)
            
            other_plans_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_O}3:{col_O}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_U}3:{col_U}{last_data_row})))=0),--((ISNUMBER(SEARCH("Plan 1200",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_C}3:{col_C}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_C}3:{col_C}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_I}3:{col_I}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_I}3:{col_I}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_O}3:{col_O}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_O}3:{col_O}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_O}3:{col_O}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_U}3:{col_U}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_U}3:{col_U}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_U}3:{col_U}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 3, plan_count_col + 1, other_plans_formula, fmt_plan_count_value)
    
    # ==============================================================================
    # Step 6: HARRY'S DOWNLINE SECTION
    # ==============================================================================
    
    downline_start_row = plan_count_start_row + 6
    downline_col = totals_col
    
    fmt_downline_header = workbook.add_format({
        'bold': True,
        'bg_color': '#C6E0B4',
        'border': 1,
        'align': 'center',
        'font_size': 12
    })
    fmt_downline_client = workbook.add_format({
        'bold': True,
        'bg_color': '#E2EFDA',
        'border': 1,
        'align': 'left'
    })
    fmt_downline_agent = workbook.add_format({
        'bg_color': '#F4F7F0',
        'border': 1,
        'align': 'left',
        'indent': 1
    })
    fmt_downline_commission = workbook.add_format({
        'num_format': '$#,##0.00',
        'bg_color': '#E2EFDA',
        'border': 1
    })
    
    # Determine which rates to use based on group type and write header
    if group_type == GROUP_TYPE_ADAM:
        group_rates = ADAMS_GROUP_AGENTS
        group_name = "Adam's Group Brokers"
        ws_comm.merge_range(downline_start_row, downline_col, downline_start_row, downline_col + 3, 
                            "ADAM'S GROUP COMMISSIONS", fmt_downline_header)
    else:
        group_rates = HARRY_DOWNLINE_RATES
        group_name = "Harry's Downline"
        ws_comm.merge_range(downline_start_row, downline_col, downline_start_row, downline_col + 3, 
                            "HARRY'S DOWNLINE COMMISSIONS", fmt_downline_header)
    
    current_downline_row = downline_start_row + 2
    
    ws_comm.write(current_downline_row, downline_col, "Client/Agent", fmt_header)
    ws_comm.write(current_downline_row, downline_col + 1, "Plan 1000 Count", fmt_header)
    ws_comm.write(current_downline_row, downline_col + 2, "Other Plans Count", fmt_header)
    ws_comm.write(current_downline_row, downline_col + 3, "Commission", fmt_header)
    
    ws_comm.set_column(downline_col, downline_col, 25)
    ws_comm.set_column(downline_col + 1, downline_col + 2, 18)
    ws_comm.set_column(downline_col + 3, downline_col + 3, 15)
    
    current_downline_row += 1
    
    # Process each client in Harry's downline OR agent in Adam's group
    if group_type == GROUP_TYPE_ADAM:
        # For Adam's Group: No client layer, just agents
        for agent_name, rates in group_rates.items():
            ws_comm.write(current_downline_row, downline_col, agent_name, fmt_downline_agent)
            
            # Reference to plan count cells
            plan_1000_count_cell = xl_rowcol_to_cell(plan_count_start_row + 2, plan_count_col + 1)
            other_plans_count_cell = xl_rowcol_to_cell(plan_count_start_row + 3, plan_count_col + 1)
            
            # Show counts
            ws_comm.write_formula(current_downline_row, downline_col + 1, f'={plan_1000_count_cell}', fmt_plan_count_value)
            ws_comm.write_formula(current_downline_row, downline_col + 2, f'={other_plans_count_cell}', fmt_plan_count_value)
            
            # Calculate commission using individual plan rates
            rate_1000 = get_rate_for_plan(rates, '1000')
            rate_1600 = get_rate_for_plan(rates, '1600')
            
            commission_formula = f'=({plan_1000_count_cell}*{rate_1000})+({other_plans_count_cell}*{rate_1600})'
            ws_comm.write_formula(current_downline_row, downline_col + 3, commission_formula, fmt_downline_commission)
            
            current_downline_row += 1
    else:
        # For Harry's Group: Iterate clients and their agents
        clients_to_process = group_rates.items()
        if selected_client and selected_client in group_rates:
            clients_to_process = [(selected_client, group_rates[selected_client])]
        
        for client_name, agents in clients_to_process:
            ws_comm.write(current_downline_row, downline_col, client_name, fmt_downline_client)
            current_downline_row += 1
            
            for agent_name, rates in agents.items():
                ws_comm.write(current_downline_row, downline_col, f"  {agent_name}", fmt_downline_agent)
                
                # Reference to plan count cells
                plan_1000_count_cell = xl_rowcol_to_cell(plan_count_start_row + 2, plan_count_col + 1)
                other_plans_count_cell = xl_rowcol_to_cell(plan_count_start_row + 3, plan_count_col + 1)
                
                # Show counts
                ws_comm.write_formula(current_downline_row, downline_col + 1, f'={plan_1000_count_cell}', fmt_plan_count_value)
                ws_comm.write_formula(current_downline_row, downline_col + 2, f'={other_plans_count_cell}', fmt_plan_count_value)
                
                # Calculate commission with CONFIDENCE multipliers
                if client_name == 'CONFIDENCE' and num_weeks in CONFIDENCE_MULTIPLIERS:
                    rate_1000 = CONFIDENCE_MULTIPLIERS[num_weeks]['1000']
                    rate_other = CONFIDENCE_MULTIPLIERS[num_weeks]['other']
                else:
                    rate_1000 = get_rate_for_plan(rates, '1000')
                    rate_other = get_rate_for_plan(rates, '1600')
                
                commission_formula = f'=({plan_1000_count_cell}*{rate_1000})+({other_plans_count_cell}*{rate_other})'
                ws_comm.write_formula(current_downline_row, downline_col + 3, commission_formula, fmt_downline_commission)
                
                current_downline_row += 1
    
    workbook.close()
    
    # Generate appropriate success message based on group type
    if group_type == GROUP_TYPE_ADAM:
        print(f"\n✅ ADAM'S GROUP REPORT GENERATED: {filename}")
        print(f"📊 Brokers: All 4 (Light House, CBsupport, ALFRED LEOPOLD, Adam Charon)")
        num_agents = len(ADAMS_GROUP_AGENTS)
    else:
        client_info = f" - {selected_client}" if selected_client else " (All Clients)"
        print(f"\n✅ HARRY'S GROUP REPORT GENERATED: {filename}")
        print(f"📊 Client{client_info}")
        num_agents = 1 if selected_client else len(HARRY_DOWNLINE_RATES)
    
    print(f"📊 Frequency: {freq_name} (÷{freq_val})")
    print(f"📅 Date Range: {packets[0]['date'].strftime('%m/%d/%Y')} - {packets[-1]['date'].strftime('%m/%d/%Y')}")
    print(f"👥 Total Employees: {len(master_ssn)}")
    print(f"✅ Perfect Employees: {len(perfect_employees)}")
    print(f"❌ Imperfect Employees: {len(imperfect_employees)}")
    print(f"📋 Features:")
    print(f"   ✓ Plan Counting ({num_weeks} weeks)")
    
    if group_type == GROUP_TYPE_ADAM:
        print(f"   ✓ Adam's Brokers Commissions ({num_agents} brokers)")
    else:
        print(f"   ✓ Harry's Downline Commissions ({num_agents} client{'s' if num_agents > 1 else ''})")
    
    if selected_client == 'CONFIDENCE' and num_weeks in CONFIDENCE_MULTIPLIERS:
        print(f"   ✓ CONFIDENCE multipliers applied for {num_weeks} weeks")

# ==============================================================================
# 4. DYNAMIC GROUP REPORT BUILDER
# ==============================================================================

def build_dynamic_group_report(packets, group_config):
    """Build Excel report for dynamic groups - EXACTLY like Harry's Group with custom agents"""
    if not packets: 
        print("❌ No valid data found.")
        return
    
    # Extract configuration
    group_name = group_config.get('group_name', 'Dynamic Group')
    main_agents = group_config.get('main_agents', {})  # {name: percentage}
    sub_agents = group_config.get('sub_agents', {})    # {name: {1600: rate, ...}}
    
    report_date = packets[0]['date']
    freq_name = packets[0]['freq_name']
    freq_val = packets[0]['freq'] if packets[0]['freq'] else 52
    
    filename = f"Commission_Report_{group_name}_{report_date.strftime('%B_%Y')}.xlsx"
    out_path = os.path.join(OUTPUT_FOLDER, filename)
    
    workbook = xlsxwriter.Workbook(out_path, {'nan_inf_to_errors': True})
    
    # FORMATS (SAME as Harry's)
    fmt_header = workbook.add_format({
        'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    fmt_currency = workbook.add_format({'num_format': '$#,##0.00'})
    fmt_text = workbook.add_format({'num_format': '@'})
    fmt_date_header = workbook.add_format({
        'bold': True, 'bg_color': '#4472C4', 'font_color': '#FFFFFF', 'border': 1, 'align': 'center'
    })
    
    # Color formats for each main agent (cycle through colors)
    colors = ['#D9E1F2', '#E2EFDA', '#FCE4D6', '#F4B084', '#C5E0B4', '#FFE699']
    agent_formats_map = {}
    for i, agent_name in enumerate(main_agents.keys()):
        agent_formats_map[agent_name] = workbook.add_format({
            'num_format': '$#,##0.00',
            'bg_color': colors[i % len(colors)]
        })
    
    fmt_total_header = workbook.add_format({
        'bold': True, 'bg_color': '#000000', 'font_color': '#FFFFFF', 'align': 'center', 'border': 1
    })
    fmt_total_value = workbook.add_format({
        'num_format': '$#,##0.00', 'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center', 'font_size': 12
    })
    
    master_ssn = set()
    employee_payments = {}
    
    # STEP 1: Create date-named tabs (SAME as Harry's Group)
    for i, p in enumerate(packets):
        tab_date = f"{p['date'].month}.{p['date'].day}"
        tab_name = tab_date[:31]
        ws = workbook.add_worksheet(tab_name)
        
        df = p['df']
        ded_col = p['ded_col']
        id_col = p['id_col']
        
        paid = df[df[ded_col] != 0].copy()
        all_ids = df[id_col].dropna().astype(str).str.strip().unique()
        master_ssn.update(all_ids)
        
        paid_ssns = set(paid[id_col].dropna().astype(str).str.strip())
        
        for ssn in all_ids:
            if ssn not in employee_payments:
                employee_payments[ssn] = []
            if ssn in paid_ssns:
                amount = paid[paid[id_col].astype(str).str.strip() == ssn][ded_col].iloc[0]
                employee_payments[ssn].append(amount)
            else:
                employee_payments[ssn].append(0)
        
        # Write tab data
        ws.write(0, 0, "SSN", fmt_header)
        ws.write(0, 1, "PPC125", fmt_header)
        ws.write(0, 2, p['date'].strftime('%m/%d/%Y'), fmt_header)
        ws.set_column(0, 0, 15)
        ws.set_column(1, 2, 12)
        
        all_ssns_in_file = df[id_col].dropna().astype(str).str.strip().unique()
        
        row_idx = 1
        for ssn in sorted(all_ssns_in_file):
            ws.write_string(row_idx, 0, ssn, fmt_text)
            employee_row = paid[paid[id_col].astype(str).str.strip() == ssn]
            if not employee_row.empty:
                val = -abs(employee_row[ded_col].iloc[0])
                ws.write_number(row_idx, 1, val, fmt_currency)
            else:
                ws.write_string(row_idx, 1, "", fmt_text)
            row_idx += 1
        
        ws.write_string(row_idx, 0, "")
        ws.write_formula(row_idx, 1, f'=SUM(B2:B{row_idx})', fmt_currency)
    
    # Identify perfect vs imperfect employees
    perfect_employees = []
    imperfect_employees = []
    num_weeks = len(packets)
    employee_plan_levels = {}
    
    for ssn, payments in employee_payments.items():
        if len(payments) == num_weeks and all(amount != 0 for amount in payments):
            perfect_employees.append(ssn)
            first_payment = abs(payments[0])
            
            # Determine plan from payment amount
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
            imperfect_employees.append(ssn)
    
    # STEP 2: Create Unpaid tab (SAME as Harry's Group with agent columns)
    ws_unpaid = workbook.add_worksheet("Unpaid")
    ws_unpaid.write(0, 0, "SSN", fmt_header)
    ws_unpaid.set_column(0, 0, 15)
    
    current_col = 1
    unpaid_ppc_cols = []
    unpaid_plan_cols = []
    unpaid_agent_cols = {}  # {agent_name: [col_week1, col_week2, ...]}
    
    for agent_name in main_agents.keys():
        unpaid_agent_cols[agent_name] = []
    
    for i, p in enumerate(packets):
        date_display = p['date'].strftime('%m/%d/%Y')
        num_agents = len(main_agents)
        
        # Merge header for this week + all agents
        ws_unpaid.merge_range(0, current_col, 0, current_col + num_agents + 1, date_display, fmt_date_header)
        
        ws_unpaid.write(1, current_col, "PPC125", fmt_header)
        ws_unpaid.write(1, current_col + 1, "Plan", fmt_header)
        
        ws_unpaid.set_column(current_col, current_col, 12)
        ws_unpaid.set_column(current_col + 1, current_col + 1, 12)
        
        unpaid_ppc_cols.append(current_col)
        unpaid_plan_cols.append(current_col + 1)
        
        current_col += 2
        
        # Agent commission headers
        for agent_name in main_agents.keys():
            ws_unpaid.write(1, current_col, agent_name, fmt_header)
            ws_unpaid.set_column(current_col, current_col, 11)
            unpaid_agent_cols[agent_name].append(current_col)
            current_col += 1
    
    sorted_imperfect = sorted(imperfect_employees)
    
    for row_num, ssn in enumerate(sorted_imperfect):
        excel_row = row_num + 2
        ws_unpaid.write_string(row_num + 2, 0, ssn, fmt_text)
        
        for i, p in enumerate(packets):
            tab_date = f"{p['date'].month}.{p['date'].day}"
            
            # VLOOKUP PPC from date tab
            vlookup = f'=IFERROR(VLOOKUP($A{excel_row+1},\'{tab_date}\'!A:B,2,FALSE),0)'
            ws_unpaid.write_formula(row_num + 2, unpaid_ppc_cols[i], vlookup, fmt_currency)
            
            ppc_cell = xl_rowcol_to_cell(row_num + 2, unpaid_ppc_cols[i])
            
            # Plan detection formula
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
            
            # FOR EACH AGENT: Calculate commission based on Plan and Agent's percentage
            for agent_name, agent_pct in main_agents.items():
                commission_formula = f'=IF({plan_cell}="Plan 1600",(1600*{agent_pct}/100*12/{freq_val}),IF({plan_cell}="Plan 1400",(1400*{agent_pct}/100*12/{freq_val}),IF({plan_cell}="Plan 1200",(1200*{agent_pct}/100*12/{freq_val}),IF({plan_cell}="Plan 1000",(1000*{agent_pct}/100*12/{freq_val}),0))))'
                
                agent_col = unpaid_agent_cols[agent_name][i]
                ws_unpaid.write_formula(row_num + 2, agent_col, commission_formula, agent_formats_map[agent_name])
    
    # STEP 3: Create Commissions tab (THE MAIN DIFFERENCE!)
    # THIS part MUST be like Harry's: PPC + Plan + Agent1 + Agent2 + Agent3... per WEEK
    ws_comm = workbook.add_worksheet("Commissions")
    workbook.worksheets_objs.insert(0, workbook.worksheets_objs.pop())
    workbook.worksheets_objs.insert(1, workbook.worksheets_objs.pop(-1))
    
    ws_comm.freeze_panes(1, 1)
    ws_comm.write(0, 0, "SSN", fmt_header)
    ws_comm.set_column(0, 0, 15)
    
    # Build column structure: FOR EACH WEEK:
    #   - PPC column
    #   - Plan column  
    #   - Agent1 commission column
    #   - Agent2 commission column
    #   - etc.
    
    current_col = 1
    ppc_cols = []
    plan_cols = []
    agent_cols = {}  # {agent_name: [col_week1, col_week2, ...]}
    
    for agent_name in main_agents.keys():
        agent_cols[agent_name] = []
    
    for i, p in enumerate(packets):
        tab_date = f"{p['date'].month}.{p['date'].day}"
        date_display = p['date'].strftime('%m/%d/%Y')
        
        # Merge header for this week + all agents
        num_agents = len(main_agents)
        ws_comm.merge_range(0, current_col, 0, current_col + num_agents + 1, date_display, fmt_date_header)
        
        # PPC and Plan header
        ws_comm.write(1, current_col, "PPC125", fmt_header)
        ws_comm.write(1, current_col + 1, "Plan", fmt_header) 
        ppc_cols.append(current_col)
        plan_cols.append(current_col + 1)
        
        current_col += 2
        
        # Agent commission headers
        for agent_name in main_agents.keys():
            ws_comm.write(1, current_col, agent_name, fmt_header)
            agent_cols[agent_name].append(current_col)
            current_col += 1
        
        ws_comm.set_column(ppc_cols[-1], ppc_cols[-1], 12)
        ws_comm.set_column(plan_cols[-1], plan_cols[-1], 12)
    
    # Write perfect employees
    sorted_ssns = sorted(perfect_employees, key=lambda ssn: (-employee_plan_levels.get(ssn, 0), ssn))
    
    for row_num, ssn in enumerate(sorted_ssns):
        excel_row = row_num + 2
        ws_comm.write_string(row_num + 2, 0, ssn, fmt_text)
        
        for i, p in enumerate(packets):
            tab_date = f"{p['date'].month}.{p['date'].day}"
            
            # VLOOKUP PPC from date tab
            vlookup = f'=IFERROR(VLOOKUP($A{excel_row+1},\'{tab_date}\'!A:B,2,FALSE),0)'
            ws_comm.write_formula(row_num + 2, ppc_cols[i], vlookup, fmt_currency)
            
            ppc_cell = xl_rowcol_to_cell(row_num + 2, ppc_cols[i])
            
            # Plan detection formula
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
            
            # FOR EACH AGENT: Calculate commission based on Plan and Agent's percentage
            # Commission formula: IF Plan=1600 then (monthly_1600 * pct / 100 * 12 / freq), etc.
            for agent_name, agent_pct in main_agents.items():
                # Use PLAN_MAP to get monthly amounts
                # Commission = (monthly_for_plan * percentage / 100) * 12 / freq_val
                commission_formula = f'=IF({plan_cell}="Plan 1600",(1600*{agent_pct}/100*12/{freq_val}),IF({plan_cell}="Plan 1400",(1400*{agent_pct}/100*12/{freq_val}),IF({plan_cell}="Plan 1200",(1200*{agent_pct}/100*12/{freq_val}),IF({plan_cell}="Plan 1000",(1000*{agent_pct}/100*12/{freq_val}),0))))'
                
                agent_col = agent_cols[agent_name][i]
                ws_comm.write_formula(row_num + 2, agent_col, commission_formula, agent_formats_map[agent_name])
    
    # STEP 4: Weekly Totals row
    last_data_row = len(sorted_ssns) + 2
    subtotal_row = last_data_row + 2
    
    ws_comm.write(subtotal_row, 0, "Weekly Totals", fmt_total_header)
    
    for agent_name in main_agents.keys():
        for i, col in enumerate(agent_cols[agent_name]):
            ws_comm.write_formula(subtotal_row, col,
                f'=SUM({xl_col_to_name(col)}3:{xl_col_to_name(col)}{last_data_row})',
                workbook.add_format({'bold': True, 'bg_color': '#FFE699', 'border': 1}))
    
    # STEP 5: Grand Totals
    totals_col = current_col + 1
    
    ws_comm.write(0, totals_col, "GRAND TOTALS", fmt_total_header)
    
    # Write agent names ACROSS COLUMNS (horizontal layout)
    col_offset = 0
    for agent_name in main_agents.keys():
        ws_comm.write(1, totals_col + col_offset, agent_name, fmt_total_header)
        col_offset += 1
    
    # Write grand total formulas ACROSS COLUMNS (horizontal layout)
    col_offset = 0
    for agent_name in main_agents.keys():
        cols_for_agent = agent_cols[agent_name]
        ranges = [f"{xl_col_to_name(c)}3:{xl_col_to_name(c)}{last_data_row}" for c in cols_for_agent]
        grand_total_formula = f"=SUM({','.join(ranges)})"
        ws_comm.write_formula(2, totals_col + col_offset, grand_total_formula, fmt_total_value)
        col_offset += 1
    
    # Set column widths for all agent columns
    num_agents = len(main_agents)
    ws_comm.set_column(totals_col, totals_col + num_agents - 1, 18)
    
    # ==============================================================================
    # STEP 6: PLAN COUNTING SECTION (LIKE HARRY'S GROUP)
    # ==============================================================================
    
    plan_count_start_row = 5
    plan_count_col = totals_col
    
    fmt_plan_count_header = workbook.add_format({
        'bold': True,
        'bg_color': '#FFF2CC',
        'border': 1,
        'align': 'center',
        'font_size': 11
    })
    fmt_plan_count_value = workbook.add_format({
        'bold': True,
        'bg_color': '#FFF2CC',
        'border': 1,
        'align': 'center',
        'font_size': 11
    })
    
    ws_comm.merge_range(plan_count_start_row, plan_count_col, plan_count_start_row, plan_count_col + 2, 
                        "PLAN COUNTING", fmt_plan_count_header)
    
    # Build plan counting formulas based on number of weeks
    if num_weeks == 2:
        ws_comm.merge_range(plan_count_start_row + 1, plan_count_col, plan_count_start_row + 1, plan_count_col + 2,
                           "BiWeekly - 2 Payroll Weeks", fmt_header)
        ws_comm.write(plan_count_start_row + 2, plan_count_col, "Plan 1000 Count:", fmt_plan_count_header)
        ws_comm.write(plan_count_start_row + 3, plan_count_col, "Other Plans Count:", fmt_plan_count_header)
        
        if len(plan_cols) >= 2:
            col_1 = xl_col_to_name(plan_cols[0])
            col_2 = xl_col_to_name(plan_cols[1])
            
            plan_1000_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_2}3:{col_2}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 2, plan_count_col + 1, plan_1000_formula, fmt_plan_count_value)
            
            other_plans_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_2}3:{col_2}{last_data_row})))=0),--((ISNUMBER(SEARCH("Plan 1200",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_1}3:{col_1}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_2}3:{col_2}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 3, plan_count_col + 1, other_plans_formula, fmt_plan_count_value)
    
    elif num_weeks == 3:
        ws_comm.merge_range(plan_count_start_row + 1, plan_count_col, plan_count_start_row + 1, plan_count_col + 2,
                           "BiWeekly - 3 Payroll Weeks", fmt_header)
        ws_comm.write(plan_count_start_row + 2, plan_count_col, "Plan 1000 Count:", fmt_plan_count_header)
        ws_comm.write(plan_count_start_row + 3, plan_count_col, "Other Plans Count:", fmt_plan_count_header)
        
        if len(plan_cols) >= 3:
            col_1 = xl_col_to_name(plan_cols[0])
            col_2 = xl_col_to_name(plan_cols[1])
            col_3 = xl_col_to_name(plan_cols[2])
            
            plan_1000_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_3}3:{col_3}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 2, plan_count_col + 1, plan_1000_formula, fmt_plan_count_value)
            
            other_plans_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_3}3:{col_3}{last_data_row})))=0),--((ISNUMBER(SEARCH("Plan 1200",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_1}3:{col_1}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_2}3:{col_2}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_3}3:{col_3}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_3}3:{col_3}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_3}3:{col_3}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 3, plan_count_col + 1, other_plans_formula, fmt_plan_count_value)
    
    elif num_weeks == 4:
        ws_comm.merge_range(plan_count_start_row + 1, plan_count_col, plan_count_start_row + 1, plan_count_col + 2,
                           "Weekly - 4 Payroll Weeks", fmt_header)
        ws_comm.write(plan_count_start_row + 2, plan_count_col, "Plan 1000 Count:", fmt_plan_count_header)
        ws_comm.write(plan_count_start_row + 3, plan_count_col, "Other Plans Count:", fmt_plan_count_header)
        
        if len(plan_cols) >= 4:
            col_1 = xl_col_to_name(plan_cols[0])
            col_2 = xl_col_to_name(plan_cols[1])
            col_3 = xl_col_to_name(plan_cols[2])
            col_4 = xl_col_to_name(plan_cols[3])
            
            plan_1000_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_3}3:{col_3}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_4}3:{col_4}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 2, plan_count_col + 1, plan_1000_formula, fmt_plan_count_value)
            
            other_plans_formula = f'=SUMPRODUCT(--((ISNUMBER(SEARCH("Plan 1000",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_3}3:{col_3}{last_data_row}))+ISNUMBER(SEARCH("Plan 1000",{col_4}3:{col_4}{last_data_row})))=0),--((ISNUMBER(SEARCH("Plan 1200",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_1}3:{col_1}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_1}3:{col_1}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_2}3:{col_2}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_2}3:{col_2}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_3}3:{col_3}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_3}3:{col_3}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_3}3:{col_3}{last_data_row})))>0),--((ISNUMBER(SEARCH("Plan 1200",{col_4}3:{col_4}{last_data_row}))+ISNUMBER(SEARCH("Plan 1400",{col_4}3:{col_4}{last_data_row}))+ISNUMBER(SEARCH("Plan 1600",{col_4}3:{col_4}{last_data_row})))>0))'
            ws_comm.write_formula(plan_count_start_row + 3, plan_count_col + 1, other_plans_formula, fmt_plan_count_value)
    
    # ==============================================================================
    # STEP 7: SUB-AGENTS DOWNLINE SECTION (IF sub_agents EXIST)
    # ==============================================================================
    
    if sub_agents:
        downline_start_row = plan_count_start_row + 6
        downline_col = totals_col
        
        fmt_downline_header = workbook.add_format({
            'bold': True,
            'bg_color': '#C6E0B4',
            'border': 1,
            'align': 'center',
            'font_size': 12
        })
        fmt_downline_agent = workbook.add_format({
            'bg_color': '#F4F7F0',
            'border': 1,
            'align': 'left',
            'indent': 1
        })
        fmt_downline_commission = workbook.add_format({
            'num_format': '$#,##0.00',
            'bg_color': '#E2EFDA',
            'border': 1
        })
        
        ws_comm.merge_range(downline_start_row, downline_col, downline_start_row, downline_col + 3, 
                            f"{group_name.upper()} - SUB-AGENTS COMMISSIONS", fmt_downline_header)
        
        current_downline_row = downline_start_row + 2
        
        ws_comm.write(current_downline_row, downline_col, "Agent", fmt_header)
        ws_comm.write(current_downline_row, downline_col + 1, "Plan 1000 Count", fmt_header)
        ws_comm.write(current_downline_row, downline_col + 2, "Other Plans Count", fmt_header)
        ws_comm.write(current_downline_row, downline_col + 3, "Commission", fmt_header)
        
        ws_comm.set_column(downline_col, downline_col, 25)
        ws_comm.set_column(downline_col + 1, downline_col + 2, 18)
        ws_comm.set_column(downline_col + 3, downline_col + 3, 15)
        
        current_downline_row += 1
        
        # Process each sub-agent
        for agent_name, rates in sub_agents.items():
            ws_comm.write(current_downline_row, downline_col, agent_name, fmt_downline_agent)
            
            # Reference to plan count cells
            plan_1000_count_cell = xl_rowcol_to_cell(plan_count_start_row + 2, plan_count_col + 1)
            other_plans_count_cell = xl_rowcol_to_cell(plan_count_start_row + 3, plan_count_col + 1)
            
            # Show counts
            ws_comm.write_formula(current_downline_row, downline_col + 1, f'={plan_1000_count_cell}', fmt_plan_count_value)
            ws_comm.write_formula(current_downline_row, downline_col + 2, f'={other_plans_count_cell}', fmt_plan_count_value)
            
            # Calculate commission using individual plan rates
            rate_1000 = get_rate_for_plan(rates, '1000')
            rate_other = get_rate_for_plan(rates, '1600')
            
            commission_formula = f'=({plan_1000_count_cell}*{rate_1000})+({other_plans_count_cell}*{rate_other})'
            ws_comm.write_formula(current_downline_row, downline_col + 3, commission_formula, fmt_downline_commission)
            
            current_downline_row += 1
    
    workbook.close()
    
    print(f"\n✅ DYNAMIC GROUP REPORT GENERATED: {out_path}")
    print(f"📁 Group: {group_name}")
    print(f"📊 Main Agents: {', '.join(main_agents.keys())}")
    if sub_agents:
        print(f"📊 Sub-Agents: {', '.join(sub_agents.keys())}")
    print(f"📊 Frequency: {freq_name} (÷{freq_val})")
    print(f"📅 Date Range: {packets[0]['date'].strftime('%m/%d/%Y')} - {packets[-1]['date'].strftime('%m/%d/%Y')}")
    print(f"👥 Total Employees: {len(master_ssn)}")
    print(f"✅ Perfect Employees: {len(perfect_employees)}")
    print(f"❌ Imperfect Employees: {len(imperfect_employees)}")
    print(f"📋 Features:")
    print(f"   ✓ Plan Counting ({num_weeks} weeks)")
    if sub_agents:
        print(f"   ✓ Sub-Agents Downline Commissions ({len(sub_agents)} agents)")

# ==============================================================================
# 4. TIER-BASED GROUP REPORT BUILDER (System 2)
# ==============================================================================

def build_tier_group_report(packets, group_config):
    """
    Build Excel report for tier-based groups with hierarchical structure
    
    group_config = {
        'group_name': '100 Academy',
        'main_agent': {'name': 'Agent 1', 'tier': '35'},
        'sub_agents': [
            {'name': 'Agent 2', 'tier': '30'},
            {'name': 'Agent 3', 'tier': '25'}
        ]
    }
    """
    if not packets: 
        print("❌ No valid data found.")
        return

    group_name = group_config.get('group_name', 'Group')
    main_agent = group_config.get('main_agent', {'name': 'Main Agent', 'tier': '35'})
    sub_agents = group_config.get('sub_agents', [])
    
    report_date = packets[0]['date']
    freq_name = packets[0]['freq_name']
    freq_val = packets[0]['freq'] if packets[0]['freq'] else 52
    
    filename = f"Commission_Report_{group_name}_{report_date.strftime('%B_%Y')}.xlsx"
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
    
    fmt_agent_header = workbook.add_format({
        'bold': True,
        'bg_color': '#C6E0B4',
        'border': 1,
        'align': 'center',
        'font_size': 11
    })
    
    fmt_client_header = workbook.add_format({
        'bold': True,
        'bg_color': '#FFD966',
        'border': 1,
        'align': 'center',
        'font_size': 11
    })
    
    fmt_downline_agent = workbook.add_format({
        'indent': 1,
        'border': 1
    })

    master_ssn = set()
    employee_payments = {}
    
    # Step 1: Create date-named tabs
    for i, p in enumerate(packets):
        tab_date = f"{p['date'].month}.{p['date'].day}"
        tab_name = tab_date[:31]
        
        ws = workbook.add_worksheet(tab_name)
        
        df = p['df']
        ded_col = p['ded_col']
        id_col = p['id_col']
        
        all_ids = df[id_col].dropna().astype(str).str.strip().unique()
        master_ssn.update(all_ids)
        
        paid = df[df[ded_col] != 0].copy()
        paid_ssns = set(paid[id_col].dropna().astype(str).str.strip())
        
        for ssn in all_ids:
            if ssn not in employee_payments:
                employee_payments[ssn] = []
            
            if ssn in paid_ssns:
                amount = paid[paid[id_col].astype(str).str.strip() == ssn][ded_col].iloc[0]
                employee_payments[ssn].append(amount)
            else:
                employee_payments[ssn].append(0)
        
        ws.write(0, 0, "SSN", fmt_header)
        ws.write(0, 1, "PPC125", fmt_header)
        ws.write(0, 2, p['date'].strftime('%m/%d/%Y'), fmt_header)
        ws.set_column(0, 0, 15)
        ws.set_column(1, 2, 12)
        
        row_idx = 1
        for ssn in sorted(all_ids):
            ws.write(row_idx, 0, ssn, fmt_text)
            
            if ssn in paid_ssns:
                amount = paid[paid[id_col].astype(str).str.strip() == ssn][ded_col].iloc[0]
                ws.write(row_idx, 1, amount, fmt_currency)
                ws.write(row_idx, 2, p['date'].strftime('%m/%d/%Y'))
            else:
                ws.write(row_idx, 1, 0, fmt_currency)
                ws.write(row_idx, 2, "UNPAID")
            
            row_idx += 1
    
    # Identify perfect vs imperfect
    perfect_employees = []
    imperfect_employees = []
    num_weeks = len(packets)
    employee_plan_levels = {}
    
    for ssn, payments in employee_payments.items():
        if len(payments) != num_weeks:
            continue
        
        all_paid = all(p != 0 for p in payments)
        
        first_payment = None
        for p in payments:
            if p != 0:
                first_payment = p
                break
        
        if first_payment and first_payment != 0:
            plan = detect_plan_from_amount(first_payment, packets[0]['freq_name'])
            employee_plan_levels[ssn] = plan
            
            if all_paid:
                perfect_employees.append(ssn)
            else:
                imperfect_employees.append((ssn, payments))
    
    # Get plan counts
    plan_counts = get_employee_plan_counts(packets, perfect_employees, employee_plan_levels)
    
    # Step 2: Create Unpaid tab
    ws_unpaid = workbook.add_worksheet("Unpaid")
    ws_unpaid.write(0, 0, "SSN", fmt_header)
    ws_unpaid.set_column(0, 0, 15)
    
    current_col = 1
    for i, p in enumerate(packets):
        ws_unpaid.write(0, current_col, f"Week {i+1}", fmt_date_header)
        ws_unpaid.write(1, current_col, "PPC", fmt_header)
        ws_unpaid.write(1, current_col + 1, "Plan", fmt_header)
        ws_unpaid.set_column(current_col, current_col + 1, 10)
        current_col += 2
    
    sorted_imperfect = sorted([emp[0] for emp in imperfect_employees])
    
    for row_num, ssn in enumerate(sorted_imperfect):
        ws_unpaid.write(row_num + 2, 0, ssn, fmt_text)
        
        payments = employee_payments.get(ssn, [])
        col = 1
        for payment in payments:
            ws_unpaid.write(row_num + 2, col, payment, fmt_currency)
            plan = detect_plan_from_amount(payment, packets[0]['freq_name']) if payment != 0 else ""
            ws_unpaid.write(row_num + 2, col + 1, plan)
            col += 2
    
    # Step 3: Create Commissions Dashboard
    ws_comm = workbook.add_worksheet("Commissions")
    workbook.worksheets_objs.insert(0, workbook.worksheets_objs.pop())
    workbook.worksheets_objs.insert(1, workbook.worksheets_objs.pop(-1))
    
    ws_comm.freeze_panes(1, 1)
    ws_comm.write(0, 0, "SSN", fmt_header)
    ws_comm.set_column(0, 0, 15)
    
    # Build column structure
    current_col = 1
    ppc_cols = []
    plan_cols = []
    
    for i, p in enumerate(packets):
        ws_comm.write(0, current_col, f"Week {i+1}", fmt_date_header)
        ws_comm.write(1, current_col, "PPC", fmt_header)
        ws_comm.write(1, current_col + 1, "Plan", fmt_header)
        
        ppc_cols.append(current_col)
        plan_cols.append(current_col + 1)
        
        ws_comm.set_column(current_col, current_col, 10)
        ws_comm.set_column(current_col + 1, current_col + 1, 8)
        
        current_col += 2
    
    # Write perfect employees
    sorted_ssns = sorted(perfect_employees, key=lambda ssn: (-int(employee_plan_levels.get(ssn, 'PPC1000').replace('PPC', '')), ssn))
    
    for row_num, ssn in enumerate(sorted_ssns):
        ws_comm.write(row_num + 2, 0, ssn, fmt_text)
        
        payments = employee_payments.get(ssn, [])
        for i, payment in enumerate(payments):
            col = ppc_cols[i]
            ws_comm.write(row_num + 2, col, payment, fmt_currency)
            
            plan = detect_plan_from_amount(payment, packets[0]['freq_name'])
            ws_comm.write(row_num + 2, col + 1, plan)
    
    last_data_row = len(sorted_ssns) + 2
    
    # Step 4: Agent Commission Summary Section
    summary_start_row = last_data_row + 4
    
    ws_comm.merge_range(summary_start_row, 0, summary_start_row, 6,
                        f"COMMISSION SUMMARY - {group_name.upper()}", fmt_total_header)
    
    # Headers
    header_row = summary_start_row + 2
    ws_comm.write(header_row, 0, "Agent Name", fmt_header)
    ws_comm.write(header_row, 1, "Tier", fmt_header)
    ws_comm.write(header_row, 2, "PPC1600", fmt_header)
    ws_comm.write(header_row, 3, "PPC1400", fmt_header)
    ws_comm.write(header_row, 4, "PPC1200", fmt_header)
    ws_comm.write(header_row, 5, "PPC1000", fmt_header)
    ws_comm.write(header_row, 6, "Commission", fmt_header)
    
    ws_comm.set_column(0, 0, 25)
    ws_comm.set_column(1, 6, 12)
    
    # Calculate and display sub-agent commissions
    data_row = header_row + 1
    sub_agent_total = 0
    
    for agent in sub_agents:
        agent_name = agent.get('name', 'Sub Agent')
        agent_tier = agent.get('tier', '25')
        
        agent_commission = calculate_tier_commission(plan_counts, agent_tier)
        sub_agent_total += agent_commission
        
        ws_comm.write(data_row, 0, agent_name)
        ws_comm.write(data_row, 1, f"Tier {agent_tier}")
        ws_comm.write(data_row, 2, plan_counts['PPC1600'])
        ws_comm.write(data_row, 3, plan_counts['PPC1400'])
        ws_comm.write(data_row, 4, plan_counts['PPC1200'])
        ws_comm.write(data_row, 5, plan_counts['PPC1000'])
        ws_comm.write(data_row, 6, agent_commission, fmt_currency)
        
        data_row += 1
    
    # Main agent commission (their tier + override from sub-agents)
    data_row += 1
    main_agent_name = main_agent.get('name', 'Main Agent')
    main_agent_tier = main_agent.get('tier', '35')
    
    ws_comm.write(data_row, 0, f"{main_agent_name} (Main Agent)", fmt_client_header)
    ws_comm.write(data_row, 1, f"Tier {main_agent_tier}", fmt_client_header)
    
    # Main agent gets their own tier commission
    main_commission = calculate_tier_commission(plan_counts, main_agent_tier)
    
    # Plus override from all sub-agents
    main_override = 0
    for agent in sub_agents:
        agent_tier = agent.get('tier', '25')
        override = calculate_override_commission(plan_counts, main_agent_tier, agent_tier)
        main_override += override
    
    total_main_commission = main_commission + main_override
    
    ws_comm.write(data_row, 2, plan_counts['PPC1600'])
    ws_comm.write(data_row, 3, plan_counts['PPC1400'])
    ws_comm.write(data_row, 4, plan_counts['PPC1200'])
    ws_comm.write(data_row, 5, plan_counts['PPC1000'])
    ws_comm.write(data_row, 6, total_main_commission, fmt_total_value)
    
    # Show breakdown
    data_row += 1
    ws_comm.write(data_row, 0, f"  • Own Tier {main_agent_tier}", fmt_downline_agent)
    ws_comm.write(data_row, 6, main_commission, fmt_currency)
    data_row += 1
    ws_comm.write(data_row, 0, f"  • Override from Sub-Agents", fmt_downline_agent)
    ws_comm.write(data_row, 6, main_override, fmt_currency)
    
    # Grand total
    data_row += 2
    ws_comm.write(data_row, 5, "GRAND TOTAL:", fmt_total_header)
    grand_total = sub_agent_total + total_main_commission
    ws_comm.write(data_row, 6, grand_total, fmt_total_value)
    
    workbook.close()
    
    print(f"\n✅ TIER GROUP REPORT GENERATED: {filename}")
    print(f"\n📁 Group: {group_name}")
    print(f"\n👤 Main Agent: {main_agent_name} (Tier {main_agent_tier})")
    print(f"   • Own Commission: ${main_commission:,.2f}")
    print(f"   • Override from Sub-Agents: ${main_override:,.2f}")
    print(f"   • Total: ${total_main_commission:,.2f}")
    
    if sub_agents:
        print(f"\n👥 Sub-Agents: {len(sub_agents)}")
        for agent in sub_agents:
            agent_name = agent.get('name')
            agent_tier = agent.get('tier')
            comm = calculate_tier_commission(plan_counts, agent_tier)
            print(f"   - {agent_name} (Tier {agent_tier}): ${comm:,.2f}")
        print(f"\n💵 Sub-Agents Total: ${sub_agent_total:,.2f}")
    else:
        print(f"\n👥 Sub-Agents: None")
    
    print(f"\n💰 GRAND TOTAL: ${grand_total:,.2f}")
    print(f"� Frequency: {freq_name} (÷{freq_val})")
    print(f"📅 Date Range: {packets[0]['date'].strftime('%m/%d/%Y')} - {packets[-1]['date'].strftime('%m/%d/%Y')}")
    print(f"✅ Perfect Employees: {len(perfect_employees)}")
    print(f"   - PPC1600: {plan_counts['PPC1600']}")
    print(f"   - PPC1400: {plan_counts['PPC1400']}")
    print(f"   - PPC1200: {plan_counts['PPC1200']}")
    print(f"   - PPC1000: {plan_counts['PPC1000']}")

# ==============================================================================
# 5. MAIN REPORT BUILDER (Router)
# ==============================================================================

def build_full_report(packets, group_type=GROUP_TYPE_HARRY, config=None):
    """
    Main report builder - routes to appropriate sub-builder
    
    Args:
        packets: Processed employee data
        group_type: "Harry's Group", "Adam's Group", "Tier-based", or "Dynamic Group"
        config: Configuration dict containing group-specific settings
    """
    if group_type == GROUP_TYPE_HARRY:
        selected_client = config.get('selected_client') if config else None
        build_harry_group_report(packets, selected_client, group_type=GROUP_TYPE_HARRY)
    elif group_type == GROUP_TYPE_ADAM:
        build_harry_group_report(packets, selected_client=None, group_type=GROUP_TYPE_ADAM)
    elif group_type == GROUP_TYPE_DYNAMIC:
        if not config:
            print("❌ Group configuration required for Dynamic Group mode!")
            return
        build_dynamic_group_report(packets, config)
    else:
        # Other Groups (Tier-based)
        if not config:
            print("❌ Group configuration required for Tier-based Groups mode!")
            return
        build_tier_group_report(packets, config)

# ==============================================================================
# 6. INTERACTIVE CLI
# ==============================================================================

def validate_tier(tier_str):
    """Validate tier number"""
    valid_tiers = ['70', '60', '50', '45', '40', '35', '30', '25', '20', '15']
    return tier_str in valid_tiers

def get_dynamic_group_config():
    """Get configuration for a dynamic group with custom agents and rates"""
    print("\n" + "-" * 60)
    print("DYNAMIC GROUP SETUP")
    print("-" * 60)
    print("This will create a report similar to Harry's Group but with")
    print("your custom agent names and commission rates.")
    
    group_name = input("\n📁 Enter Group Name: ").strip()
    if not group_name:
        group_name = "Custom Group"
    
    # Get main agents (like Charles, Harry, Lighthouse)
    print("\n" + "=" * 60)
    print("MAIN AGENTS SETUP")
    print("=" * 60)
    print("These are the top-level agents (like Charles, Harry, Lighthouse)")
    print("They will appear in the Grand Total section.\n")
    
    main_agents = {}
    main_agent_num = 1
    
    while True:
        print(f"\n👤 MAIN AGENT #{main_agent_num}")
        agent_name = input(f"   Agent Name (or press Enter to finish): ").strip()
        if not agent_name:
            if main_agent_num == 1:
                print("   ❌ At least one main agent is required!")
                continue
            break
        
        # Get commission percentage for this main agent
        while True:
            try:
                commission_pct = float(input(f"   Commission % for {agent_name} (e.g., 10 for 10%): ").strip())
                if 0 <= commission_pct <= 100:
                    main_agents[agent_name] = commission_pct
                    print(f"   ✅ Main agent added: {agent_name} ({commission_pct}%)")
                    break
                else:
                    print("   ❌ Please enter a percentage between 0 and 100")
            except ValueError:
                print(f"   ❌ Please enter a valid number")
        
        main_agent_num += 1
        add_more = input("\n   ➕ Add another main agent? (y/n): ").strip().lower()
        if add_more != 'y':
            break
    
    # Get sub-agents (downline agents like in Harry's downline)
    print("\n" + "=" * 60)
    print("SUB-AGENTS / DOWNLINE AGENTS SETUP")
    print("=" * 60)
    print("These are downline agents (like Agent1, Agent2 in Harry's clients)")
    print("They will appear in the Downline Commissions section.\n")
    
    sub_agents = {}
    sub_agent_num = 1
    
    while True:
        print(f"\n👥 SUB-AGENT #{sub_agent_num}")
        agent_name = input(f"   Agent Name (or press Enter to finish): ").strip()
        if not agent_name:
            if sub_agent_num == 1:
                print("   ℹ️  No sub-agents added (Main agents only)")
            break
        
        # Get rates for each plan level
        agent_rates = {}
        for plan in ['1600', '1400', '1200', '1000']:
            while True:
                try:
                    rate = float(input(f"   Rate for PPC{plan}: $").strip())
                    agent_rates[plan] = rate
                    break
                except ValueError:
                    print(f"   ❌ Please enter a valid number")
        
        sub_agents[agent_name] = agent_rates
        print(f"   ✅ Sub-agent added: {agent_name}")
        
        sub_agent_num += 1
        add_more = input("\n   ➕ Add another sub-agent? (y/n): ").strip().lower()
        if add_more != 'y':
            break
    
    # Display configuration summary
    print("\n" + "=" * 60)
    print("CONFIGURATION SUMMARY")
    print("=" * 60)
    print(f"\n📁 Group: {group_name}")
    print(f"\n👤 Main Agents: {len(main_agents)}")
    for name, pct in main_agents.items():
        print(f"   • {name}: {pct}%")
    
    if sub_agents:
        print(f"\n👥 Sub-Agents: {len(sub_agents)}")
        for name in sub_agents.keys():
            print(f"   • {name}")
    else:
        print(f"\n👥 Sub-Agents: None")
    
    confirm = input("\n✅ Proceed with this configuration? (y/n): ").strip().lower()
    if confirm != 'y':
        print("\n❌ Configuration cancelled.")
        return None
    
    return {
        'group_name': group_name,
        'main_agents': main_agents,
        'sub_agents': sub_agents
    }

def get_user_input():
    """Interactive command-line interface to get configuration"""
    print("\n" + "=" * 60)
    print("COMMISSION REPORT GENERATOR - FINAL VERSION")
    print("=" * 60)
    print("\n📋 SELECT GROUP TYPE:\n")
    print("1. Harry's Group (Charles/Harry/Lighthouse + Client Downlines)")
    print("2. Adam's Group (Brokers: Light House, CBsupport, ALFRED LEOPOLD, Adam Charon)")
    print("3. Other Groups (Tier-based with Client + Agents)")
    print("4. Dynamic Group (Create custom group with custom agents & rates)")
    
    while True:
        choice = input("\nEnter your choice (1, 2, 3, or 4): ").strip()
        if choice == '1':
            print("\n✅ Selected: Harry's Group")
            
            # Ask which client to process
            print("\n" + "-" * 60)
            print("SELECT CLIENT")
            print("-" * 60)
            print("\nAvailable Clients:")
            client_list = list(HARRY_DOWNLINE_RATES.keys())
            for i, client in enumerate(client_list, 1):
                print(f"{i}. {client}")
            
            while True:
                client_choice = input(f"\nEnter client number (1-{len(client_list)}): ").strip()
                try:
                    client_idx = int(client_choice) - 1
                    if 0 <= client_idx < len(client_list):
                        selected_client = client_list[client_idx]
                        print(f"\n✅ Selected Client: {selected_client}")
                        return GROUP_TYPE_HARRY, {'selected_client': selected_client, 'group_type': GROUP_TYPE_HARRY}
                    else:
                        print(f"❌ Please enter a number between 1 and {len(client_list)}")
                except ValueError:
                    print("❌ Please enter a valid number")
                    
        elif choice == '2':
            print("\n✅ Selected: Adam's Group")
            
            # Show available agents in Adam's Group
            print("\n" + "-" * 60)
            print("SELECT AGENTS TO PROCESS")
            print("-" * 60)
            print("\nAvailable Brokers:")
            agent_list = list(ADAMS_GROUP_AGENTS.keys())
            for i, agent in enumerate(agent_list, 1):
                print(f"{i}. {agent}")
            
            # For now, process all agents (user can select specific ones later)
            print("\n📌 Processing all brokers in Adam's Group")
            return GROUP_TYPE_ADAM, {'group_type': GROUP_TYPE_ADAM}
            
        elif choice == '3':
            print("\n✅ Selected: Other Groups (Tier-based)")
            break
        elif choice == '4':
            print("\n✅ Selected: Dynamic Group")
            config = get_dynamic_group_config()
            if config is None:
                return None, None
            return GROUP_TYPE_DYNAMIC, config
        else:
            print("❌ Invalid choice. Please enter 1, 2, 3, or 4.")
    
    # Get group information for Tier-based groups
    print("\n" + "-" * 60)
    print("GROUP CONFIGURATION")
    print("-" * 60)
    
    group_name = input("\n📁 Enter Group Name (e.g., '100 Academy'): ").strip()
    if not group_name:
        group_name = "Group"
    
    # Get TOP/MAIN agent
    print("\n" + "-" * 60)
    print("TOP/MAIN AGENT")
    print("-" * 60)
    
    main_agent_name = input("\n👤 Enter Main Agent Name: ").strip()
    if not main_agent_name:
        main_agent_name = "Main Agent"
    
    # Get main agent tier
    while True:
        main_agent_tier = input(f"🎯 Enter Tier for {main_agent_name} (70, 60, 50, 45, 40, 35, 30, 25, 20, 15): ").strip()
        if validate_tier(main_agent_tier):
            print(f"   ✅ Main Agent: {main_agent_name} (Tier {main_agent_tier})")
            break
        else:
            print("   ❌ Invalid tier! Please choose from: 70, 60, 50, 45, 40, 35, 30, 25, 20, 15")
    
    # Get sub-agents
    print("\n" + "-" * 60)
    print("SUB-AGENTS (Below Main Agent)")
    print("-" * 60)
    print("Note: Sub-agents should have LOWER tiers than the main agent")
    
    sub_agents = []
    sub_agent_num = 1
    
    while True:
        print(f"\n👥 SUB-AGENT #{sub_agent_num}")
        
        sub_agent_name = input(f"   Name (or press Enter to finish): ").strip()
        if not sub_agent_name:
            if sub_agent_num == 1:
                print("   ℹ️  No sub-agents added (Main agent only)")
            break
        
        # Get sub-agent tier
        while True:
            sub_agent_tier = input(f"   Tier for {sub_agent_name} (70, 60, 50, 45, 40, 35, 30, 25, 20, 15): ").strip()
            if validate_tier(sub_agent_tier):
                if int(sub_agent_tier) >= int(main_agent_tier):
                    print(f"   ⚠️ Warning: Sub-agent tier ({sub_agent_tier}) should be LOWER than Main Agent tier ({main_agent_tier})")
                    confirm = input(f"   Continue anyway? (y/n): ").strip().lower()
                    if confirm != 'y':
                        continue
                print(f"   ✅ Sub-agent added: {sub_agent_name} (Tier {sub_agent_tier})")
                sub_agents.append({'name': sub_agent_name, 'tier': sub_agent_tier})
                break
            else:
                print("   ❌ Invalid tier! Please choose from: 70, 60, 50, 45, 40, 35, 30, 25, 20, 15")
        
        sub_agent_num += 1
        
        add_more = input("\n   ➕ Add another sub-agent? (y/n): ").strip().lower()
        if add_more != 'y':
            break
    
    # Build config
    group_config = {
        'group_name': group_name,
        'main_agent': {
            'name': main_agent_name,
            'tier': main_agent_tier
        },
        'sub_agents': sub_agents
    }
    
    # Display summary
    print("\n" + "=" * 60)
    print("CONFIGURATION SUMMARY")
    print("=" * 60)
    print(f"\n📁 Group: {group_name}")
    print(f"\n👤 Main Agent: {main_agent_name} (Tier {main_agent_tier})")
    print(f"   - Earns: Tier {main_agent_tier} rates + Override from sub-agents")
    if sub_agents:
        print(f"\n👥 Sub-Agents: {len(sub_agents)}")
        for i, agent in enumerate(sub_agents, 1):
            print(f"   {i}. {agent['name']} (Tier {agent['tier']}) - Earns: Tier {agent['tier']} rates only")
    else:
        print(f"\n👥 Sub-Agents: None (Main agent only)")
    
    confirm = input("\n✅ Proceed with this configuration? (y/n): ").strip().lower()
    if confirm != 'y':
        print("\n❌ Configuration cancelled. Exiting...")
        return None, None
    
    return GROUP_TYPE_OTHER, group_config

# ==============================================================================
# 7. MAIN EXECUTION
# ==============================================================================

if __name__ == "__main__":
    # Get user configuration
    group_type, client_config = get_user_input()
    
    if group_type is None:
        exit(0)
    
    # Process files
    print("\n" + "=" * 60)
    print("PROCESSING FILES")
    print("=" * 60)
    
    packets = process_raw_files()
    
    if packets:
        build_full_report(packets, group_type, client_config)
        print("\n" + "=" * 60)
        print("✅ REPORT GENERATION COMPLETE!")
        print("=" * 60)
        print(f"\n📁 Check the Output folder for your report.")
    else:
        print("\n❌ No files to process. Please add files to Input_Raw folder.")
