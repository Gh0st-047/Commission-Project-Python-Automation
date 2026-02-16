from fpdf import FPDF

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'Project Bible: Automated Commission Dashboard', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

    def chapter_title(self, title):
        self.set_font('Arial', 'B', 12)
        self.set_fill_color(230, 230, 230)
        self.cell(0, 10, title, 0, 1, 'L', 1)
        self.ln(4)

    def body_text(self, text):
        self.set_font('Arial', '', 11)
        self.multi_cell(0, 6, text)
        self.ln()

pdf = PDF()
pdf.add_page()

# 1. THE PROBLEM
pdf.chapter_title("1. The Client's Goal (The 'Why')")
pdf.body_text(
    "The client manages an insurance program where employees pay premiums via payroll deductions. "
    "Currently, the client performs a painful manual process every pay period:\n"
    "1. Downloads raw payroll reports (CSV/Excel).\n"
    "2. Looks at the deduction column (PPC125).\n"
    "3. Manually calculates 'Is this Weekly or Monthly?'\n"
    "4. Manually figures out which 'Plan' the agent sold (Plan 1000, 1600, etc.).\n"
    "5. Pastes this data into a Master Excel Template to calculate commissions for agents (Charles, Harry, Lighthouse)."
)

# 2. THE SCRIBE EXPLAINED
pdf.chapter_title("2. What was the Scribe Link?")
pdf.body_text(
    "The Scribe link was the 'Instruction Manual' for the logic. It taught us how to reverse-engineer the 'Plan' "
    "using only the deduction amount. It provided the 'Golden Formula':\n\n"
    "Formula: (12 * Deduction) / Frequency = Plan Value\n\n"
    "Example from Scribe:\n"
    "- If Deduction is $600 and Frequency is Semi-Monthly (24 payments/yr):\n"
    "- Math: (12 * 600) / 24 = 300 (Wait, math check: (600*24)/12 = 1200).\n"
    "- Result: The employee is on 'Plan 1200'.\n\n"
    "The Scribe told us: 'Look at the money, do the math, find the Plan.'"
)

# 3. THE DATA FLOW
pdf.chapter_title("3. The Data Flow (Inputs & Outputs)")
pdf.body_text(
    "INPUT (The Raw Data):\n"
    "- Files: 'Patriot Payroll Confirmation Reports'.\n"
    "- Key Column: 'D-ppc 125' (The money deducted).\n"
    "- Weirdness: The numbers are often NEGATIVE (e.g., -369.23).\n"
    "- Action: Your script must convert them to positive to do the math, but keep them negative for the final report.\n\n"
    "OUTPUT (The Templates):\n"
    "- Files: 'Commissions Template (Weekly).xlsx', '(BiWeekly).xlsx', etc.\n"
    "- Structure: These files have empty tabs named W1, W2, W3.\n"
    "- Action: Your script fills these empty tabs. The main 'Commissions' tab has pre-built Excel formulas that read W1/W2 and show the final money."
)

# 4. THE LOGIC ENGINE
pdf.chapter_title("4. The Logic Engine (Your Python Code)")
pdf.body_text(
    "Your script acts as a 'Detective'. It opens a raw file and grabs a deduction (e.g., $369.23).\n\n"
    "It tests the number against the rules:\n"
    "- Is (369.23 * 52) / 12 a valid Plan? -> YES (Plan 1600). -> Verdict: WEEKLY File.\n"
    "- Is (369.23 * 26) / 12 a valid Plan? -> NO. -> Verdict: Not Bi-Weekly.\n\n"
    "Once it knows it is 'Weekly', it selects the 'Weekly Template' and pastes the data."
)

# 5. SCOPE & PHASE 2
pdf.chapter_title("5. Project Scope (Phase 1 vs Phase 2)")
pdf.body_text(
    "PHASE 1 (Now - Due Friday):\n"
    "- Build a Web Dashboard on Google Cloud.\n"
    "- User uploads raw files (1-4 files).\n"
    "- System detects frequency, cleans data, fills the W1/W2/W3 tabs.\n"
    "- User downloads the finished Excel.\n\n"
    "PHASE 2 (Later):\n"
    "- Advanced breakdowns for other downline agents.\n"
    "- Handling complex split-commission rules."
)

pdf.output("Project_Bible_Summary.pdf")
print("PDF Generated: Project_Bible_Summary.pdf")