# Commission Report Generator

## Quick Start Guide

### 1. Install Python
- Download and install Python 3.8 or higher from [python.org](https://www.python.org/downloads/)
- During installation, check "Add Python to PATH"

### 2. Create Virtual Environment
```bash
python -m venv env
```

### 3. Activate Environment
**Windows:**
```bash
env\Scripts\activate
```

**Mac/Linux:**
```bash
source env/bin/activate
```

### 4. Install Dependencies
```bash
pip install -r requirements.txt
```

### 5. Prepare Input Files
- Create a folder named `Input_Raw` in the project directory (if not already created)
- Place all your Excel/CSV payroll files inside `Input_Raw` folder

### 6. Run the Script
```bash
python main.py
```

### 7. Get Results
- Find the generated report in the `Output` folder
- Report name: `Commission_Report_[Month]_[Year].xlsx`

---

## Output Structure
The Excel report contains:
- **Date tabs** (e.g., "12.7", "12.14") - Weekly payroll data
- **Commissions tab** - Employees who paid all weeks
- **Unpaid tab** - Employees with missing payments
- Commission calculations for Charles, Harry, and LightHouse

---

## Troubleshooting
- If no files found: Check that Excel/CSV files are in `Input_Raw` folder
- If errors occur: Ensure files have PPC125 and SSN columns
- Python not recognized: Reinstall Python and check "Add to PATH"

---

## Support
For issues or questions, contact your developer.
