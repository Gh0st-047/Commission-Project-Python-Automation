# This file shows the CORRECTED build_dynamic_group_report function structure
# Copy this logic into final.py replacing the old build_dynamic_group_report

def build_dynamic_group_report_CORRECT(packets, group_config):
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
        ws.write(0, 1, "PPC Deduction", fmt_header)
        ws.write(0, 2, "Plan", fmt_header)
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
    
    # STEP 2: Create Unpaid tab (SAME as Harry's Group)
    ws_unpaid = workbook.add_worksheet("Unpaid")
    ws_unpaid.write(0, 0, "SSN", fmt_header)
    
    col = 1
    for i, p in enumerate(packets):
        ws_unpaid.merge_range(0, col, 0, col + 1, f"Week {i+1}", fmt_header)
        ws_unpaid.write(1, col, "PPC", fmt_header)
        ws_unpaid.write(1, col + 1, "Plan", fmt_header)
        col += 2
    
    ws_unpaid.set_column(0, 0, 15)
    
    sorted_imperfect = sorted(imperfect_employees)
    for row_num, ssn in enumerate(sorted_imperfect):
        ws_unpaid.write(row_num + 2, 0, ssn, fmt_text)
        payments = employee_payments.get(ssn, [])
        col = 1
        for payment in payments:
            ws_unpaid.write(row_num + 2, col, payment, fmt_currency)
            plan = detect_plan_from_amount(payment, freq_name) if payment != 0 else ""
            ws_unpaid.write(row_num + 2, col + 1, plan)
            col += 2
    
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
    
    row = 1
    for agent_name in main_agents.keys():
        ws_comm.write(row, totals_col, agent_name, fmt_total_header)
        row += 1
    
    row = 2
    for agent_name in main_agents.keys():
        cols_for_agent = agent_cols[agent_name]
        ranges = [f"{xl_col_to_name(c)}3:{xl_col_to_name(c)}{last_data_row}" for c in cols_for_agent]
        grand_total_formula = f"=SUM({','.join(ranges)})"
        ws_comm.write_formula(row, totals_col, grand_total_formula, fmt_total_value)
        row += 1
    
    ws_comm.set_column(totals_col, totals_col, 18)
    
    workbook.close()
    
    print(f"\n✅ DYNAMIC GROUP REPORT GENERATED: {out_path}")
    print(f"📁 Group: {group_name}")
    print(f"📊 Agents: {', '.join(main_agents.keys())}")
    print(f"📊 Frequency: {freq_name} (÷{freq_val})")
    print(f"📅 Date Range: {packets[0]['date'].strftime('%m/%d/%Y')} - {packets[-1]['date'].strftime('%m/%d/%Y')}")
    print(f"👥 Total Employees: {len(master_ssn)}")
    print(f"✅ Perfect Employees: {len(perfect_employees)}")
    print(f"❌ Imperfect Employees: {len(imperfect_employees)}")
