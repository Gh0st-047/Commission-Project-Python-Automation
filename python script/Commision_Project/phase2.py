import pandas as pd
import glob
import os
import re
from datetime import datetime
import sys
# Force UTF-8 output for Windows consoles
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell, xl_col_to_name

# ==============================================================================
# CONFIGURATION
# ==============================================================================
INPUT_FOLDER = 'Input_Raw'
OUTPUT_FOLDER = 'Output'

os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Group Types
GROUP_TYPE_HARRY = "Harry's Group"
GROUP_TYPE_OTHER = "Other Groups"

# Commission Logic - Original PLAN_MAP for Harry's Group
PLAN_MAP = {
    1600: {'Weekly': 369.23, 'BiWeekly': 738.46, 'SemiMonthly': 800, 'Monthly': 1600},
    1400: {'Weekly': 323.08, 'BiWeekly': 646.15, 'SemiMonthly': 700, 'Monthly': 1400},
    1200: {'Weekly': 276.92, 'BiWeekly': 553.85, 'SemiMonthly': 600, 'Monthly': 1200},
    1000: {'Weekly': 230.77, 'BiWeekly': 461.54, 'SemiMonthly': 500, 'Monthly': 1000}
}

# Tier Rates - Extracted from tier.xlsx
# Each tier shows the commission rate per employee per plan level
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
            return datetime(int(year), int(month), int(day))
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
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)
        except Exception as e:
            print(f"⚠️ Could not read {filepath}: {e}")
            continue
            
        df.columns = df.columns.str.strip()
        
        # Find required columns
        ded_col = next((c for c in df.columns if 'ppc' in c.lower() and '125' in c.lower()), None)
        date_col = next((c for c in df.columns if 'date' in c.lower()), None)
        id_col = 'SSN' if 'SSN' in df.columns else df.columns[0]
        
        if not ded_col:
            print(f"⚠️ No PPC125 column in {filepath}")
            continue

        df = df.dropna(subset=[id_col])
        
        # Clean deduction column
        df[ded_col] = df[ded_col].astype(str).str.replace('$', '', regex=False).str.replace(',', '', regex=False)
        df[ded_col] = pd.to_numeric(df[ded_col], errors='coerce').fillna(0)
        
        # Determine frequency
        freq = 52
        freq_name = "Weekly"
        sample = df[df[ded_col] != 0].head(20)
        
        for val in sample[ded_col]:
            f, fn = get_frequency_from_deduction(val)
            if f:
                freq = f
                freq_name = fn
                break
        
        # Extract date
        check_date = None
        if date_col:
            check_date = pd.to_datetime(df[date_col].dropna().iloc[0], errors='coerce')
        
        if pd.isna(check_date) or check_date is None:
            check_date = extract_date_from_filename(os.path.basename(filepath))
        
        if check_date is None:
            print(f"⚠️ Could not extract date from {filepath}")
            continue
        
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
        
    # Sort by date
    processed.sort(key=lambda x: x['date'])
    return processed

# ==============================================================================
# 2. TIER COMMISSION CALCULATIONS
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
# 3. HARRY'S GROUP REPORT BUILDER (From main.py)
# ==============================================================================

def build_harry_group_report(packets):
    """Build Excel report for Harry's Group (Milestone 1 format)"""
    if not packets: 
        print("❌ No valid data found.")
        return

    report_date = packets[0]['date']
    freq_name = packets[0]['freq_name']
    freq_val = packets[0]['freq'] if packets[0]['freq'] else 52
    
    filename = f"Commission_Report_Harry_Output.xlsx"
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
            ws.write(row_idx, 0, ssn, fmt_text)
            
            if ssn in paid_ssns:
                amount = paid[paid[id_col].astype(str).str.strip() == ssn][ded_col].iloc[0]
                ws.write(row_idx, 1, amount, fmt_currency)
                ws.write(row_idx, 2, p['date'].strftime('%m/%d/%Y'))
            else:
                ws.write(row_idx, 1, 0, fmt_currency)
                ws.write(row_idx, 2, "UNPAID")
            
            row_idx += 1
    
    # Identify perfect vs imperfect employees
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
            plan_number = int(plan.replace('PPC', '')) if plan else 1000
            employee_plan_levels[ssn] = plan
            
            if all_paid:
                perfect_employees.append(ssn)
            else:
                imperfect_employees.append((ssn, payments))
    
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
    charles_cols = []
    harry_cols = []
    lighthouse_cols = []
    
    for i, p in enumerate(packets):
        ws_comm.write(0, current_col, f"Week {i+1}", fmt_date_header)
        ws_comm.write(1, current_col, "PPC", fmt_header)
        ws_comm.write(1, current_col + 1, "Plan", fmt_header)
        ws_comm.write(1, current_col + 2, "Charles", fmt_header)
        ws_comm.write(1, current_col + 3, "Harry", fmt_header)
        ws_comm.write(1, current_col + 4, "LightHouse", fmt_header)
        
        ppc_cols.append(current_col)
        charles_cols.append(current_col + 2)
        harry_cols.append(current_col + 3)
        lighthouse_cols.append(current_col + 4)
        
        ws_comm.set_column(current_col, current_col, 10)
        ws_comm.set_column(current_col + 1, current_col + 1, 8)
        ws_comm.set_column(current_col + 2, current_col + 4, 12)
        
        current_col += 5
    
    # Write perfect employees
    sorted_ssns = sorted(perfect_employees, key=lambda ssn: (-int(employee_plan_levels.get(ssn, '0').replace('PPC', '')) if employee_plan_levels.get(ssn) else 0, ssn))
    
    for row_num, ssn in enumerate(sorted_ssns):
        ws_comm.write(row_num + 2, 0, ssn, fmt_text)
        
        payments = employee_payments.get(ssn, [])
        for i, payment in enumerate(payments):
            col = ppc_cols[i]
            ws_comm.write(row_num + 2, col, payment, fmt_currency)
            
            plan = detect_plan_from_amount(payment, packets[0]['freq_name'])
            ws_comm.write(row_num + 2, col + 1, plan)
            
            # Calculate commissions
            freq = packets[i]['freq']
            charles_comm = payment / freq if freq else 0
            harry_comm = charles_comm * 2
            lighthouse_comm = charles_comm * 3
            
            ws_comm.write(row_num + 2, charles_cols[i], charles_comm, fmt_charles)
            ws_comm.write(row_num + 2, harry_cols[i], harry_comm, fmt_harry)
            ws_comm.write(row_num + 2, lighthouse_cols[i], lighthouse_comm, fmt_lighthouse)
    
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
        start_row = 2
        end_row = last_data_row
        
        charles_formula = f"=SUM({xl_col_to_name(charles_cols[i])}{start_row+1}:{xl_col_to_name(charles_cols[i])}{end_row})"
        harry_formula = f"=SUM({xl_col_to_name(harry_cols[i])}{start_row+1}:{xl_col_to_name(harry_cols[i])}{end_row})"
        lighthouse_formula = f"=SUM({xl_col_to_name(lighthouse_cols[i])}{start_row+1}:{xl_col_to_name(lighthouse_cols[i])}{end_row})"
        
        ws_comm.write_formula(subtotal_row, charles_cols[i], charles_formula, fmt_weekly_total)
        ws_comm.write_formula(subtotal_row, harry_cols[i], harry_formula, fmt_weekly_total)
        ws_comm.write_formula(subtotal_row, lighthouse_cols[i], lighthouse_formula, fmt_weekly_total)
    
    # Grand Totals
    totals_col = current_col + 1
    
    ws_comm.write(0, totals_col, "GRAND TOTALS", fmt_total_header)
    ws_comm.write(1, totals_col, "Charles", fmt_total_header)
    ws_comm.write(1, totals_col + 1, "Harry", fmt_total_header)
    ws_comm.write(1, totals_col + 2, "LightHouse", fmt_total_header)
    
    def build_sum_formula(cols):
        refs = [f"{xl_col_to_name(c)}{subtotal_row+1}" for c in cols]
        return f"={'+'.join(refs)}"
    
    ws_comm.write_formula(2, totals_col, build_sum_formula(charles_cols), fmt_total_value)
    ws_comm.write_formula(2, totals_col + 1, build_sum_formula(harry_cols), fmt_total_value)
    ws_comm.write_formula(2, totals_col + 2, build_sum_formula(lighthouse_cols), fmt_total_value)
    
    ws_comm.set_column(totals_col, totals_col + 2, 18)
    
    workbook.close()
    print(f"\n✅ HARRY'S GROUP REPORT GENERATED: {filename}")
    print(f"📊 Frequency: {freq_name} (÷{freq_val})")
    print(f"📅 Date Range: {packets[0]['date'].strftime('%m/%d/%Y')} - {packets[-1]['date'].strftime('%m/%d/%Y')}")
    print(f"👥 Total Employees: {len(master_ssn)}")
    print(f"✅ Perfect Employees: {len(perfect_employees)}")
    print(f"❌ Imperfect Employees: {len(imperfect_employees)}")

# ==============================================================================
# 4. TIER-BASED GROUP REPORT BUILDER (Option A - Hierarchical)
# ==============================================================================

def build_tier_group_report(packets, client_config):
    """
    Build Excel report for tier-based groups with hierarchical structure
    
    client_config = {
        'client_name': 'John',
        'client_tier': '70',
        'agents': [
            {'name': 'Agent 1', 'tier': '50'},
            {'name': 'Agent 2', 'tier': '40'}
        ]
    }
    """
    if not packets: 
        print("❌ No valid data found.")
        return

    client_name = client_config.get('client_name', 'Client')
    client_tier = client_config.get('client_tier', '70')
    agents = client_config.get('agents', [])
    
    report_date = packets[0]['date']
    freq_name = packets[0]['freq_name']
    freq_val = packets[0]['freq'] if packets[0]['freq'] else 52
    
    filename = f"Commission_Report_{client_name}_{report_date.strftime('%B_%Y')}.xlsx"
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

    master_ssn = set()
    employee_payments = {}
    
    # Step 1: Create date-named tabs (same as Harry's)
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
                        f"COMMISSION SUMMARY - {client_name.upper()}'S GROUP", fmt_total_header)
    
    # Headers
    header_row = summary_start_row + 2
    ws_comm.write(header_row, 0, "Agent Name", fmt_header)
    ws_comm.write(header_row, 1, "Tier", fmt_header)
    ws_comm.write(header_row, 2, "PPC1600", fmt_header)
    ws_comm.write(header_row, 3, "PPC1400", fmt_header)
    ws_comm.write(header_row, 4, "PPC1200", fmt_header)
    ws_comm.write(header_row, 5, "PPC1000", fmt_header)
    ws_comm.write(header_row, 6, "Total Commission", fmt_header)
    
    ws_comm.set_column(0, 0, 20)
    ws_comm.set_column(1, 6, 12)
    
    # Calculate and display agent commissions
    data_row = header_row + 1
    agent_totals = []
    
    for agent in agents:
        agent_name = agent.get('name', 'Agent')
        agent_tier = agent.get('tier', '50')
        
        agent_commission = calculate_tier_commission(plan_counts, agent_tier)
        agent_totals.append(agent_commission)
        
        ws_comm.write(data_row, 0, agent_name)
        ws_comm.write(data_row, 1, f"Tier {agent_tier}")
        ws_comm.write(data_row, 2, plan_counts['PPC1600'])
        ws_comm.write(data_row, 3, plan_counts['PPC1400'])
        ws_comm.write(data_row, 4, plan_counts['PPC1200'])
        ws_comm.write(data_row, 5, plan_counts['PPC1000'])
        ws_comm.write(data_row, 6, agent_commission, fmt_currency)
        
        data_row += 1
    
    # Client override commission
    data_row += 1
    ws_comm.write(data_row, 0, f"{client_name} (Client Override)", fmt_client_header)
    ws_comm.write(data_row, 1, f"Tier {client_tier}", fmt_client_header)
    
    client_override = 0
    for agent in agents:
        agent_tier = agent.get('tier', '50')
        override = calculate_override_commission(plan_counts, client_tier, agent_tier)
        client_override += override
    
    ws_comm.write(data_row, 2, plan_counts['PPC1600'])
    ws_comm.write(data_row, 3, plan_counts['PPC1400'])
    ws_comm.write(data_row, 4, plan_counts['PPC1200'])
    ws_comm.write(data_row, 5, plan_counts['PPC1000'])
    ws_comm.write(data_row, 6, client_override, fmt_total_value)
    
    # Grand total
    data_row += 2
    ws_comm.write(data_row, 5, "GRAND TOTAL:", fmt_total_header)
    grand_total = sum(agent_totals) + client_override
    ws_comm.write(data_row, 6, grand_total, fmt_total_value)
    
    workbook.close()
    
    print(f"\n✅ TIER GROUP REPORT GENERATED: {filename}")
    print(f"👤 Client: {client_name} (Tier {client_tier})")
    print(f"👥 Agents: {len(agents)}")
    for agent in agents:
        agent_tier = agent.get('tier')
        agent_name = agent.get('name')
        comm = calculate_tier_commission(plan_counts, agent_tier)
        print(f"   - {agent_name} (Tier {agent_tier}): ${comm:,.2f}")
    print(f"💰 Client Override: ${client_override:,.2f}")
    print(f"💵 Grand Total: ${grand_total:,.2f}")
    print(f"📊 Frequency: {freq_name} (÷{freq_val})")
    print(f"📅 Date Range: {packets[0]['date'].strftime('%m/%d/%Y')} - {packets[-1]['date'].strftime('%m/%d/%Y')}")
    print(f"✅ Perfect Employees: {len(perfect_employees)}")
    print(f"   - PPC1600: {plan_counts['PPC1600']}")
    print(f"   - PPC1400: {plan_counts['PPC1400']}")
    print(f"   - PPC1200: {plan_counts['PPC1200']}")
    print(f"   - PPC1000: {plan_counts['PPC1000']}")

# ==============================================================================
# 5. MAIN REPORT BUILDER (Router)
# ==============================================================================

def build_full_report(packets, group_type=GROUP_TYPE_HARRY, client_config=None):
    """
    Main report builder - routes to appropriate sub-builder
    
    Args:
        packets: Processed employee data
        group_type: "Harry's Group" or "Other Groups"
        client_config: For Other Groups, contains client and agents config
    """
    if group_type == GROUP_TYPE_HARRY:
        build_harry_group_report(packets)
    else:
        if not client_config:
            print("❌ Client configuration required for Other Groups mode!")
            return
        build_tier_group_report(packets, client_config)

# ==============================================================================
# 6. INTERACTIVE CLI
# ==============================================================================

def validate_tier(tier_str):
    """Validate tier number"""
    valid_tiers = ['70', '60', '50', '45', '40', '35', '30', '25', '20', '15']
    return tier_str in valid_tiers

def get_user_input():
    """Interactive command-line interface to get configuration"""
    print("\n" + "=" * 60)
    print("COMMISSION REPORT GENERATOR - PHASE 2")
    print("=" * 60)
    print("\n📋 SELECT GROUP TYPE:\n")
    print("1. Harry's Group (Charles/Harry/Lighthouse)")
    print("2. Other Groups (Tier-based with Client + Agents)")
    
    while True:
        choice = input("\nEnter your choice (1 or 2): ").strip()
        if choice == '1':
            print("\n✅ Selected: Harry's Group")
            return GROUP_TYPE_HARRY, None
        elif choice == '2':
            print("\n✅ Selected: Other Groups (Tier-based)")
            break
        else:
            print("❌ Invalid choice. Please enter 1 or 2.")
    
    # Get client information
    print("\n" + "-" * 60)
    print("CLIENT CONFIGURATION")
    print("-" * 60)
    
    client_name = input("\n👤 Enter Client Name: ").strip()
    if not client_name:
        client_name = "Client"
    
    # Get client tier
    while True:
        client_tier = input(f"🎯 Enter Client Tier for {client_name} (70, 60, 50, 45, 40, 35, 30, 25, 20, 15): ").strip()
        if validate_tier(client_tier):
            print(f"   ✅ Valid tier: {client_tier}")
            break
        else:
            print("   ❌ Invalid tier! Please choose from: 70, 60, 50, 45, 40, 35, 30, 25, 20, 15")
    
    # Get agents
    print("\n" + "-" * 60)
    print("AGENT CONFIGURATION")
    print("-" * 60)
    
    agents = []
    agent_num = 1
    
    while True:
        print(f"\n👥 AGENT #{agent_num}")
        
        agent_name = input(f"   Name (or press Enter to skip): ").strip()
        if not agent_name:
            if agent_num == 1:
                print("   ⚠️ You must add at least one agent!")
                continue
            else:
                break
        
        # Get agent tier
        while True:
            agent_tier = input(f"   Tier for {agent_name} (70, 60, 50, 45, 40, 35, 30, 25, 20, 15): ").strip()
            if validate_tier(agent_tier):
                # Check if agent tier is lower than client tier
                if int(agent_tier) >= int(client_tier):
                    print(f"   ⚠️ Warning: Agent tier ({agent_tier}) should typically be lower than Client tier ({client_tier})")
                    confirm = input(f"   Continue anyway? (y/n): ").strip().lower()
                    if confirm != 'y':
                        continue
                print(f"   ✅ Agent added: {agent_name} (Tier {agent_tier})")
                agents.append({'name': agent_name, 'tier': agent_tier})
                break
            else:
                print("   ❌ Invalid tier! Please choose from: 70, 60, 50, 45, 40, 35, 30, 25, 20, 15")
        
        agent_num += 1
        
        # Ask if more agents
        add_more = input("\n   ➕ Add another agent? (y/n): ").strip().lower()
        if add_more != 'y':
            break
    
    # Build config
    client_config = {
        'client_name': client_name,
        'client_tier': client_tier,
        'agents': agents
    }
    
    # Display summary
    print("\n" + "=" * 60)
    print("CONFIGURATION SUMMARY")
    print("=" * 60)
    print(f"\n👤 Client: {client_name} (Tier {client_tier})")
    print(f"👥 Agents: {len(agents)}")
    for i, agent in enumerate(agents, 1):
        print(f"   {i}. {agent['name']} (Tier {agent['tier']})")
    
    confirm = input("\n✅ Proceed with this configuration? (y/n): ").strip().lower()
    if confirm != 'y':
        print("\n❌ Configuration cancelled. Exiting...")
        return None, None
    
    return GROUP_TYPE_OTHER, client_config

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
