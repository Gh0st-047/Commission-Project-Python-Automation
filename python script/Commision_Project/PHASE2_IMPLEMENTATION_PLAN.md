# PHASE 2 IMPLEMENTATION PLAN

## 📋 Project Overview
Create **phase2.py** that combines:
1. ✅ All Milestone 1 features from main.py
2. ✅ Tier-based commission system for non-Harry groups
3. ✅ Embedded tier rate data from tier.xlsx

---

## 🎯 Core Features to Implement

### **PART A: Carry Over from main.py (Milestone 1)**
- [x] Process CSV/Excel files from Input_Raw folder
- [x] Extract dates from filenames
- [x] Detect payment frequency (Weekly/BiWeekly/SemiMonthly/Monthly)
- [x] Identify perfect employees (paid all weeks) vs imperfect (missing payments)
- [x] Create date-named tabs (e.g., "12.7", "12.14")
- [x] Create Commissions tab (perfect employees only)
- [x] Create Unpaid tab (imperfect employees)
- [x] Plan detection (PPC1600, PPC1400, PPC1200, PPC1000)
- [x] Commission calculation using PLAN_MAP
- [x] Weekly totals row
- [x] Grand totals (Charles, Harry, Lighthouse) - **for Harry's group only**

### **PART B: NEW - Tier-Based Commission System**
- [ ] **TIER_RATES dictionary** - Store all tier commission rates
- [ ] **Group selection system** - Allow user to choose group type (Harry's vs Other Groups)
- [ ] **Dynamic agent configuration** - Support multiple agents with different tiers
- [ ] **Tier-based commission calculation** - Use tier rates instead of PLAN_MAP
- [ ] **Flexible report generation** - Different output format for non-Harry groups
- [ ] **Agent commission breakdown** - Show per-agent earnings by tier

---

## 📊 Data Structure: TIER_RATES

```python
# Extracted from tier.xlsx - Complete tier rate table
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
```

---

## 🏗️ Architecture Design

### **1. Configuration Section**
```python
# Group types
GROUP_TYPE_HARRY = "Harry's Group"
GROUP_TYPE_OTHER = "Other Groups"

# Agent configuration structure
AGENT_CONFIG = {
    'group_name': 'Sales Team Alpha',  # Custom group name
    'agents': [
        {'name': 'Agent 1', 'tier': '70'},
        {'name': 'Agent 2', 'tier': '60'},
        {'name': 'Agent 3', 'tier': '50'},
    ]
}
```

### **2. Main Functions**

#### `process_raw_files()` - ✅ Keep from main.py
- No changes needed
- Returns processed packets with employee data

#### `get_employee_plan_counts(packets, perfect_employees)` - **NEW**
```python
def get_employee_plan_counts(packets, perfect_employees):
    """
    Count how many perfect employees are in each plan level
    Returns: {'PPC1600': 25, 'PPC1400': 18, 'PPC1200': 12, 'PPC1000': 8}
    """
    plan_counts = {'PPC1600': 0, 'PPC1400': 0, 'PPC1200': 0, 'PPC1000': 0}
    
    for ssn in perfect_employees:
        # Determine plan from first payment amount
        first_payment = get_first_payment(ssn, packets)
        plan = detect_plan_from_amount(first_payment)
        if plan:
            plan_counts[plan] += 1
    
    return plan_counts
```

#### `calculate_tier_commission(plan_counts, tier)` - **NEW**
```python
def calculate_tier_commission(plan_counts, tier):
    """
    Calculate commission for an agent based on their tier
    
    Args:
        plan_counts: {'PPC1600': 25, 'PPC1400': 18, ...}
        tier: '70', '60', '50', etc.
    
    Returns:
        Total commission amount
    """
    rates = TIER_RATES[tier]
    total = 0
    
    for plan, count in plan_counts.items():
        rate = rates.get(plan, 0)
        total += count * rate
    
    return total
```

#### `build_harry_group_report(packets, workbook)` - **REFACTORED**
- Extract Harry's group logic from main.py
- Keep Charles/Harry/Lighthouse grand totals
- Keep colored columns
- Keep existing format

#### `build_tier_group_report(packets, workbook, agent_config)` - **NEW**
```python
def build_tier_group_report(packets, workbook, agent_config):
    """
    Generate report for tier-based groups (non-Harry)
    
    Structure:
    - Date-named tabs (same as Harry's)
    - Commissions tab:
        * SSN column
        * Week columns (PPC, Plan)
        * Agent commission columns (one per agent)
        * Total row per agent
    - Unpaid tab (same structure)
    - Agent Summary section at bottom
    """
```

#### `build_full_report(packets, group_type, agent_config=None)` - **MODIFIED**
```python
def build_full_report(packets, group_type=GROUP_TYPE_HARRY, agent_config=None):
    """
    Main report builder - routes to appropriate sub-builder
    
    Args:
        packets: Processed employee data
        group_type: "Harry's Group" or "Other Groups"
        agent_config: For Other Groups, contains agents and tiers
    """
    if group_type == GROUP_TYPE_HARRY:
        build_harry_group_report(packets, workbook)
    else:
        build_tier_group_report(packets, workbook, agent_config)
```

---

## 📝 Output Format Differences

### **Harry's Group (Existing)**
```
Commissions Tab:
SSN | Week1 PPC | Week1 Plan | Week1 Charles | Week1 Harry | Week1 Lighthouse | ... | GRAND TOTALS
----|-----------|------------|---------------|-------------|------------------|-----|-------------
123 | $738.46   | 1600       | $24.62        | $49.23      | $73.85           | ... | Charles: $X
456 | $646.15   | 1400       | ...           | ...         | ...              | ... | Harry: $Y
... | ...       | ...        | ...           | ...         | ...              | ... | Lighthouse: $Z

Weekly Totals Row
```

### **Other Groups (NEW)**
```
Commissions Tab:
SSN | Week1 PPC | Week1 Plan | Week2 PPC | Week2 Plan | ... | Agent1 (Tier70) | Agent2 (Tier60) | Agent3 (Tier50)
----|-----------|------------|-----------|------------|-----|-----------------|-----------------|----------------
123 | $738.46   | 1600       | $738.46   | 1600       | ... | -               | -               | -
456 | $646.15   | 1400       | $646.15   | 1400       | ... | -               | -               | -

[Blank rows]

AGENT COMMISSION SUMMARY
Agent Name      | Tier | PPC1600 Count | PPC1400 Count | PPC1200 Count | PPC1000 Count | Total Commission
----------------|------|---------------|---------------|---------------|---------------|------------------
Agent 1         | 70   | 25            | 18            | 12            | 8             | $7,234.00
Agent 2         | 60   | 25            | 18            | 12            | 8             | $6,429.00
Agent 3         | 50   | 25            | 18            | 12            | 8             | $5,519.00
----------------|------|---------------|---------------|---------------|---------------|------------------
TOTAL           |      |               |               |               |               | $19,182.00
```

---

## 🎨 User Interface Flow (For Future Dashboard)

### **Step 1: Select Group Type**
```
[ ] Harry's Group (pre-configured Charles/Harry/Lighthouse)
[ ] Other Groups (tier-based custom agents)
```

### **Step 2: Configure Agents (if Other Groups selected)**
```
Group Name: [________________]

Agents:
1. Name: [________] Tier: [Dropdown: 70, 60, 50, 45, 40, 35, 30, 25, 20, 15]
2. Name: [________] Tier: [Dropdown: 70, 60, 50, 45, 40, 35, 30, 25, 20, 15]
3. Name: [________] Tier: [Dropdown: 70, 60, 50, 45, 40, 35, 30, 25, 20, 15]

[+ Add Agent]
```

### **Step 3: Upload Files & Generate Report**
```
Upload payroll files → Process → Generate Excel report
```

---

## 🔧 Implementation Steps

### **Phase 2A: Setup & Data Extraction**
1. Create `phase2.py` file
2. Copy all imports and basic structure from `main.py`
3. Add `TIER_RATES` dictionary with all tier data
4. Add `GROUP_TYPE_HARRY` and `GROUP_TYPE_OTHER` constants

### **Phase 2B: Refactor Existing Code**
5. Keep `process_raw_files()` unchanged
6. Keep `extract_date_from_filename()` unchanged
7. Keep `get_frequency_from_deduction()` unchanged

### **Phase 2C: Build New Tier Functions**
8. Implement `get_employee_plan_counts(packets, perfect_employees)`
9. Implement `detect_plan_from_amount(amount, freq_name)`
10. Implement `calculate_tier_commission(plan_counts, tier)`

### **Phase 2D: Split Report Builders**
11. Extract Harry's logic into `build_harry_group_report()`
12. Create new `build_tier_group_report()` for Other Groups
13. Modify `build_full_report()` to route based on group_type

### **Phase 2E: Testing Configuration**
14. Add command-line configuration at bottom:
```python
if __name__ == "__main__":
    # Configuration
    GROUP_TYPE = GROUP_TYPE_OTHER  # or GROUP_TYPE_HARRY
    
    AGENT_CONFIG = {
        'group_name': 'Sales Team Alpha',
        'agents': [
            {'name': 'John Doe', 'tier': '70'},
            {'name': 'Jane Smith', 'tier': '60'},
            {'name': 'Bob Johnson', 'tier': '50'},
        ]
    }
    
    packets = process_raw_files()
    build_full_report(packets, GROUP_TYPE, AGENT_CONFIG)
```

---

## ✅ Testing Checklist

- [ ] Test Harry's Group mode (should work exactly like main.py)
- [ ] Test Other Groups mode with 1 agent
- [ ] Test Other Groups mode with 3 agents
- [ ] Test with different tier combinations (70/60/50, 45/35/25, etc.)
- [ ] Verify plan counts are correct
- [ ] Verify tier commission calculations match manual calculation
- [ ] Test with 2-week, 3-week, 4-week payroll data
- [ ] Verify perfect vs imperfect employee sorting
- [ ] Check Excel formatting (colors, borders, formulas)

---

## 🚀 Future Enhancements (Phase 3 - Dashboard)

- [ ] Web interface with Flask/Streamlit
- [ ] Dropdown for group selection
- [ ] Dynamic agent addition with (+) button
- [ ] Real-time commission preview
- [ ] Export to Excel from web interface
- [ ] Save agent configurations for reuse
- [ ] Historical report comparison

---

## 📦 File Structure

```
Commision_Project/
├── Input_Raw/           # Input CSV/Excel files
├── Output/              # Generated reports
├── main.py              # Milestone 1 (Harry's Group only)
├── phase2.py            # ⭐ NEW - Supports both Harry's & Other Groups
├── tier.xlsx            # Reference (not needed at runtime)
├── PHASE2_FEATURES.md   # Harry's downline documentation
├── PHASE2_TIER_EXPLANATION.md  # Tier system explanation
└── PHASE2_IMPLEMENTATION_PLAN.md  # This file
```

---

## 🎯 Success Criteria

**phase2.py is complete when:**

1. ✅ It can run in "Harry's Group" mode and produce identical output to main.py
2. ✅ It can run in "Other Groups" mode with configurable agents/tiers
3. ✅ Tier commissions calculate correctly for all tiers (70, 60, 50, 45, 40, 35, 30, 25, 20, 15)
4. ✅ Perfect employee filtering works across both modes
5. ✅ Excel output is clean, formatted, and ready for client review
6. ✅ Console output shows processing summary

---

## ⚡ Ready to Implement?

**Estimated Lines of Code:** ~800-1000 lines
**Estimated Implementation Time:** 2-3 hours
**Complexity:** Medium-High

**Next Step:** Await your green flag to proceed with implementation! 🚦

---

**Created:** January 30, 2026
**Status:** 🟡 Awaiting Approval
**Developer:** Fawaad S.
