# Financial-Aid-Packaging-Tool
Generates a plain-speak financial aid overview page for students
import os
import shutil
import openpyxl
import math  # Imported to handle ceiling rounding

# Direct, explicit import optimized for Python environment
import pypdf
from pypdf import PdfWriter as PdfMerger

# Import ReportLab layout components
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# =====================================================================
# FORCE ABSOLUTE WORKING DIRECTORY
# =====================================================================
desktop_dir = r"C:\Users\cbaggarly\OneDrive - Goodwill Industries of MiGA and the CSRA\Desktop"
os.chdir(desktop_dir)

excel_file = "Macon Campus Package Database.xlsm"
temp_excel_snapshot = "temp_snapshot_package.xlsm"

# =====================================================================
# 1. SILENT EXCEL SCANNER (TARGETED TO RESTRUCTURED COORD MATRIX)
# =====================================================================
selected_flyers = []
student_name = "Unknown_Student"
program_name = "Unknown_Program"
program_cost = 0.0
pell_award = 0.0
sub_gross = 0.0          
unsub_gross = 0.0        

# Restructured Q4 Explicit Variable Buckets
q4_anticipated_pell = 0.0
q4_anticipated_sub = 0.0
q4_anticipated_unsub = 0.0

loan_total = 0.0
award_total = 0.0
va_type = ""             

q_costs = []

def safe_float(val):
    if val is None:
        return 0.0
    val_str = str(val).strip()
    if val_str.startswith("#") or not val_str:
        return 0.0
    try:
        return float(val_str.replace(",", "").replace("$", ""))
    except ValueError:
        return 0.0

try:
    shutil.copy2(excel_file, temp_excel_snapshot)
    wb = openpyxl.load_workbook(temp_excel_snapshot, data_only=True, read_only=True)
    
    if "Package_Data" in wb.sheetnames:
        ws = wb["Package_Data"]
        
        student_name = str(ws["B16"].value).strip() if ws["B16"].value is not None else "Unknown_Student"
        program_name = str(ws["B17"].value).strip() if ws["B17"].value is not None else "Unknown_Program"
        
        program_cost = safe_float(ws["B18"].value)
        pell_award = safe_float(ws["B19"].value)
        sub_gross = safe_float(ws["B20"].value)    
        unsub_gross = safe_float(ws["B21"].value)  
        
        # Hunting inside your new dedicated Q4 block fields!
        q4_anticipated_pell = safe_float(ws["B22"].value)
        q4_anticipated_sub = safe_float(ws["B23"].value)
        q4_anticipated_unsub = safe_float(ws["B24"].value)
        
        loan_total = safe_float(ws["B26"].value)   
        award_total = safe_float(ws["B27"].value)  
        
        if ws["F4"].value is not None:
            va_type = str(ws["F4"].value).strip()
        
        for row in range(2, 8):
            flyer_name_val = ws.cell(row=row, column=4).value   
            checked_val = ws.cell(row=row, column=5).value      
            
            if flyer_name_val is not None:
                flyer_name = str(flyer_name_val).strip()
                is_checked = False
                if checked_val is not None:
                    is_checked = str(checked_val).strip().upper() in ["X", "YES", "TRUE", "1", "1.0", "CHECK"]
                if is_checked:
                    selected_flyers.append(flyer_name)

    if "Program_Master" in wb.sheetnames:
        pm_ws = wb["Program_Master"]
        for row in range(2, 20):
            row_prog = str(pm_ws.cell(row=row, column=1).value).strip().upper()
            if row_prog == program_name.upper():
                q1_val = safe_float(pm_ws.cell(row=row, column=3).value)
                q2_val = safe_float(pm_ws.cell(row=row, column=4).value)
                q3_val = safe_float(pm_ws.cell(row=row, column=5).value)
                q4_val = safe_float(pm_ws.cell(row=row, column=6).value)
                q_costs = [q1_val, q2_val, q3_val, q4_val]
                break

    wb.close()
finally:
    if os.path.exists(temp_excel_snapshot):
        os.remove(temp_excel_snapshot)

if student_name == "nan" or not student_name or student_name.startswith("#"):
    student_name = "Unknown_Student"

FLYER_MAP = {
    "Dependent Student": os.path.join(desktop_dir, r"Packet_Flyers\Dependent_Plus_Flyer.pdf"),
    "Student Loans": os.path.join(desktop_dir, r"Packet_Flyers\Student_Loans_Flyer.pdf"),
    "VA Student": os.path.join(desktop_dir, r"Packet_Flyers\VA_Flyer.pdf"),
    "Scholarship Flyer": os.path.join(desktop_dir, r"Packet_Flyers\Scholarships_Flyer.pdf"),
    "Helms GFS App": os.path.join(desktop_dir, r"Packet_Flyers\Helms College General Fund Scholarship Application.pdf"),
    "ION Tuition": os.path.join(desktop_dir, r"Packet_Flyers\ION_Tuition_Flyer.pdf")
}

is_va_student = "VA Student" in selected_flyers

if is_va_student:
    if "33" in va_type:
        va_label = "Chapter 33 Post-9/11"
    elif "35" in va_type:
        va_label = "Chapter 35 DEA"
    else:
        va_label = "VA Educational Benefits"
else:
    va_label = ""

# =====================================================================
# 2. NET LEDGER MATH ENGINE (REVISED FOR FLEXIBLE FUNDING)
# =====================================================================
program_upper = program_name.upper()
is_culinary_diploma = "CULINARY DIPLOMA" in program_upper

if is_culinary_diploma:
    total_program_quarters = 4  
    ay1_quarters = 3            
elif "CULINARY AAS" in program_upper:
    total_program_quarters = 3
    ay1_quarters = 3
else:
    total_program_quarters = 2 if "CERTIFICATE" in program_upper else 3
    ay1_quarters = total_program_quarters

# Program Master Safe Scan Guard
if not q_costs:
    q_costs = [program_cost / float(total_program_quarters)] * total_program_quarters

# Precision Ay1 Calculator Matrix
pell_per_q = round(pell_award / ay1_quarters, 2) if ay1_quarters > 0 else 0.0
pell_disb_list = [pell_per_q] * ay1_quarters

origination_fee_rate = 0.01057

# AY1 Subsidized Loan Stripper
sub_net_total = sub_gross * (1.0 - origination_fee_rate)
sub_base = int(sub_net_total / ay1_quarters) if ay1_quarters > 0 else 0
sub_disb_list = [sub_base] * ay1_quarters
sub_remainder = int(round(sub_net_total - (sub_base * ay1_quarters), 2)) if ay1_quarters > 0 else 0
for i in range(sub_remainder):
    sub_disb_list[i] += 1

# AY1 Unsubsidized Loan Stripper
unsub_net_total = unsub_gross * (1.0 - origination_fee_rate)
unsub_base = int(round(unsub_net_total / ay1_quarters, 2)) if ay1_quarters > 0 else 0
unsub_disb_list = [unsub_base] * ay1_quarters
unsub_remainder = int(round(unsub_net_total - (unsub_base * ay1_quarters), 2)) if ay1_quarters > 0 else 0
for i in range(unsub_remainder):
    unsub_disb_list[i] += 1

# CALCULATE QUARTERLY AID (Includes loans regardless of is_va_student)
quarterly_net_aid = []
for i in range(ay1_quarters):
    q_aid = pell_disb_list[i] + sub_disb_list[i] + unsub_disb_list[i]
    quarterly_net_aid.append(q_aid)

running_balances = []
current_balance = 0.0

for i in range(ay1_quarters):
    q_charge = q_costs[i] if i < len(q_costs) else 0.0
    q_aid = quarterly_net_aid[i]
    current_balance = current_balance + q_charge - q_aid
    running_balances.append(round(current_balance, 2))

q3_ending_balance = running_balances[-1] if len(running_balances) > 0 else 0.0

# Q4 FUNDING MATRIX
if is_culinary_diploma:
    q4_charge = q_costs[3] if len(q_costs) > 3 else 0.0
    q4_net_sub = q4_anticipated_sub * (1.0 - origination_fee_rate)
    q4_net_unsub = q4_anticipated_unsub * (1.0 - origination_fee_rate)
    q4_true_combined_net_aid = q4_anticipated_pell + q4_net_sub + q4_net_unsub
    
    q4_ending_balance = q3_ending_balance + q4_charge - q4_true_combined_net_aid
    balance_refund = q3_ending_balance  
    total_net_funding = sum(quarterly_net_aid) + q4_true_combined_net_aid
else:
    balance_refund = current_balance
    total_net_funding = sum(quarterly_net_aid)
    q4_ending_balance = 0.0
    q4_true_combined_net_aid = 0.0

# Apply ceiling rules to stipend refund paths to enforce whole numbers (e.g. $167.59 becomes $168)
rounded_balance_refund = math.ceil(abs(balance_refund)) if balance_refund < 0 else balance_refund
rounded_q4_ending_balance = math.ceil(abs(q4_ending_balance)) if q4_ending_balance < 0 else q4_ending_balance

# =====================================================================
# 3. GENERATE TAILORED COVER SHEET LAYOUT (FSA COMPLIANT VERBIAGE)
# =====================================================================
cover_pdf = "temp_cover.pdf"
doc = SimpleDocTemplate(
    cover_pdf, 
    pagesize=letter, 
    leftMargin=36, 
    rightMargin=36, 
    topMargin=30, 
    bottomMargin=30
)
styles = getSampleStyleSheet()
story = []

NAVY = colors.HexColor("#002D62")
GOLD = colors.HexColor("#F4A261")
DARK_GRAY = colors.HexColor("#333333")
LIGHT_BG = colors.HexColor("#F4F6F9")

title_style = ParagraphStyle('TStyle', fontName='Helvetica-Bold', fontSize=21, leading=25, textColor=colors.white, alignment=1, spaceAfter=4)
subtitle_style = ParagraphStyle('SubStyle', fontName='Helvetica-Bold', fontSize=13, leading=16, textColor=GOLD, alignment=1, spaceAfter=2)
program_style = ParagraphStyle('PStyle', fontName='Helvetica', fontSize=10.5, leading=14, textColor=colors.white, alignment=1)
section_style = ParagraphStyle('SecStyle', fontName='Helvetica-Bold', fontSize=13, leading=16, textColor=NAVY, spaceBefore=12, spaceAfter=4)
th_style = ParagraphStyle('THStyle', fontName='Helvetica-Bold', fontSize=9.5, leading=12, textColor=colors.white, alignment=1)
td_style = ParagraphStyle('TDStyle', fontName='Helvetica', fontSize=9.5, leading=12, textColor=DARK_GRAY, alignment=1)
td_left = ParagraphStyle('TDLeft', fontName='Helvetica', fontSize=9.5, leading=12, textColor=DARK_GRAY, alignment=0)
outcome_style = ParagraphStyle('OutcomeStyle', fontName='Helvetica', fontSize=10.5, leading=14, textColor=DARK_GRAY)
disc_style = ParagraphStyle('DiscStyle', fontName='Helvetica', fontSize=9, leading=13, textColor=DARK_GRAY, leftIndent=15, firstLineIndent=-10)
loan_box_style = ParagraphStyle('LoanBoxStyle', fontName='Helvetica', fontSize=10, textColor=NAVY, leading=13)

# COMPLIANCE UPDATE: Clarified that "Aid" includes both grants and loans to be repaid
if is_culinary_diploma:
    funding_label = f"Total Estimated Financial Aid Offered: ${total_net_funding:,.2f}"
else:
    funding_label = f"Total Estimated Financial Aid Package: ${total_net_funding:,.2f}"

header_data = [
    [Paragraph("FINANCIAL AID PACKAGING OVERVIEW", title_style)],
    [Paragraph(f"Prepared For: {student_name} {f'({va_label})' if is_va_student else ''}", subtitle_style)],
    [Paragraph(f"Program: {program_name} | Total Program Cost: ${program_cost:,.2f} | {funding_label}", program_style)]
]
header_table = Table(header_data, colWidths=[540])
header_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), NAVY),
    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('TOPPADDING', (0,0), (-1,-1), 12),
    ('BOTTOMPADDING', (0,0), (-1,-1), 12),
]))
story.append(header_table)
story.append(Spacer(1, 8))

# Dynamic Differentiator Text Mapping (Updated for FSA Compliance)
if is_va_student and "33" in va_type:
    ch33_refund_total = math.ceil(sum(pell_disb_list) + q4_anticipated_pell)
    # COMPLIANCE UPDATE: Changed "direct stipend refund" to "estimated credit balance refund"
    outcome_text = f"<b>CHAPTER 33 POST-9/11 BENEFIT SUMMARY:</b> Your program instructional tuition and fee costs are certified and paid directly to the school by the VA. Your Federal Pell Grant is applied to your ledger, creating an estimated credit balance refund of <b>${ch33_refund_total:,.0f}</b> distributed across your scheduled checkpoints."
    box_bg = colors.HexColor("#E2F0D9") 
    box_line = colors.HexColor("#385723")
    
elif is_va_student and "35" in va_type:
    total_pell_all_q = sum(pell_disb_list) + q4_anticipated_pell
    ch35_balance_due = program_cost - total_pell_all_q
    outcome_text = f"<b>CHAPTER 35 DEA BENEFIT SUMMARY:</b> Under Chapter 35, the VA distributes monthly allowance stipends <b>directly to the student</b>, not the school. On the institutional ledger, your Federal Pell Grant reduces your direct program costs, leaving an estimated remaining balance of <b>${ch35_balance_due:,.2f}</b>. <b>An Institutional Payment Arrangement is required</b> to resolve this remaining balance."
    box_bg = colors.HexColor("#FCE4D6") 
    box_line = colors.HexColor("#C65911")
    
elif is_culinary_diploma:
    if q4_ending_balance <= 0:
        # COMPLIANCE UPDATE: Replaced "surplus" with "estimated credit balance"
        outcome_text = f"<b>ESTIMATED PROGRAM CREDIT BALANCE:</b> Your combined financial aid allocations across all 4 quarters completely cover your total program costs, leaving a projected final program credit balance of <b>${rounded_q4_ending_balance:,.0f}</b>."
        box_bg = colors.HexColor("#E2F0D9") 
        box_line = colors.HexColor("#385723")
    else:
        # COMPLIANCE UPDATE: Removed "liability" to soften and generalize options transparently
        outcome_text = f"<b>ESTIMATED OUT-OF-POCKET BALANCE DUE (4-QUARTER PATH):</b> After applying your current financial aid disbursements and projecting next year's anticipated aid against your full program costs, an estimated open balance of <b>${q4_ending_balance:,.2f}</b> remains across the duration of your program. <b>Payment Options Available:</b> A customized Institutional Payment Plan will be provided to break this balance down into monthly installments. Additionally, you may be evaluated for internal <b>Institutional Aid</b> options (such as scholarships or grants) which can further reduce your out-of-pocket balance."
        box_bg = colors.HexColor("#FCE4D6") 
        box_line = colors.HexColor("#C65911")
        
else:
    if balance_refund <= 0:
        # COMPLIANCE UPDATE: Avoided treating credit refunds as guaranteed stipends
        outcome_text = f"<b>ESTIMATED CREDIT BALANCE REFUND:</b> Your estimated financial aid awards completely clear your current academic year costs, leaving an estimated credit balance of <b>${rounded_balance_refund:,.0f}</b>. Credit balance refunds are processed and distributed by the school following federal guidelines after your final scheduled funding post."
        box_bg = colors.HexColor("#E2F0D9") 
        box_line = colors.HexColor("#385723")
    else:
        outcome_text = f"<b>ESTIMATED OUT-OF-POCKET BALANCE DUE:</b> After applying your net financial aid disbursements against direct institutional costs, there remains an estimated net open balance of <b>${balance_refund:,.2f}</b> to be cleared for this academic cycle. <b>Payment Options Available:</b> To assist with this balance, a tailored Institutional Payment Plan will be provided to establish monthly installment options. You may also be eligible for internal <b>Institutional Aid</b> programs that can be applied directly to your account to reduce your overall balance due."
        box_bg = colors.HexColor("#FCE4D6") 
        box_line = colors.HexColor("#C65911")

outcome_table = Table([[Paragraph(outcome_text, outcome_style)]], colWidths=[540])
outcome_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), box_bg),
    ('BOX', (0,0), (-1,-1), 1.2, box_line),
    ('TOPPADDING', (0,0), (-1,-1), 8),
    ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ('LEFTPADDING', (0,0), (-1,-1), 10),
    ('RIGHTPADDING', (0,0), (-1,-1), 10),
]))
story.append(outcome_table)

story.append(Paragraph("Funding Breakdown & Ledger Timeline", section_style))

timeline_data = [
    [
        Paragraph("Term Milestone", th_style), 
        Paragraph("Tuition & Fees", th_style), 
        Paragraph("Pell Grant", th_style),
        Paragraph("Sub. Loan (Net)", th_style),
        Paragraph("Unsub. Loan (Net)", th_style),
        Paragraph("Estimated Balance", th_style) # COMPLIANCE: Avoid calling it a definitive "Running" or final balance
    ]
]

for i in range(ay1_quarters):
    milestone_label = f"Quarter {i+1} Start"
    q_charge_val = q_costs[i] if i < len(q_costs) else 0.0
    q_pell_val = pell_disb_list[i] if i < len(pell_disb_list) else 0.0
    q_sub_val = float(sub_disb_list[i]) if i < len(sub_disb_list) else 0.0
    q_unsub_val = float(unsub_disb_list[i]) if i < len(unsub_disb_list) else 0.0
    q_bal_val = running_balances[i] if i < len(running_balances) else 0.0
    
    if is_va_student and "33" in va_type:
        bal_display = Paragraph("<font color='#385723'><b>VA Certified</b></font>", td_style)
    else:
        if q_bal_val < 0:
            rounded_q_bal = math.ceil(abs(q_bal_val))
            bal_display = Paragraph(f"-${rounded_q_bal:,.0f} (Estimated Credit)", td_style)
        else:
            bal_display = Paragraph(f"${q_bal_val:,.2f}", td_style)
    
    timeline_data.append([
        Paragraph(milestone_label, td_left),
        Paragraph(f"${q_charge_val:,.2f}", td_style),
        Paragraph(f"${q_pell_val:,.2f}" if q_pell_val > 0 else "$0.00", td_style),
        Paragraph(f"${q_sub_val:,.2f}" if q_sub_val > 0 else "$0.00", td_style),
        Paragraph(f"${q_unsub_val:,.2f}" if q_unsub_val > 0 else "$0.00", td_style),
        bal_display
    ])

if is_culinary_diploma:
    q4_charge_val = q_costs[3] if len(q_costs) > 3 else 0.0
    if is_va_student and "33" in va_type:
        q4_bal_display = Paragraph("<font color='#385723'><b>VA Certified</b></font>", td_style)
    else:
        if q4_ending_balance < 0:
            q4_bal_display = Paragraph(f"-${rounded_q4_ending_balance:,.0f} (Estimated Credit)", td_style)
        else:
            q4_bal_display = Paragraph(f"${q4_ending_balance:,.2f}", td_style)
    
    timeline_data.append([
        Paragraph("Quarter 4 Start<br/><font size='7' color='#555555'>(Next Aid Year)</font>", td_left),
        Paragraph(f"${q4_charge_val:,.2f}", td_style),
        Paragraph(f"${q4_anticipated_pell:,.2f}" if q4_anticipated_pell > 0 else "$0.00", td_style),
        Paragraph(f"${q4_net_sub:,.2f}" if q4_net_sub > 0 else "$0.00", td_style),
        Paragraph(f"${q4_net_unsub:,.2f}" if q4_net_unsub > 0 else "$0.00", td_style),
        q4_bal_display
    ])

timeline_table = Table(timeline_data, colWidths=[105, 87, 87, 87, 87, 87])
timeline_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_BG]),
]))
story.append(timeline_table)
story.append(Spacer(1, 8))

# Final Loan Box Logic (Updated for compliance)
if is_va_student and "33" in va_type:
    loan_box_text = f"""
    <b>MILITARY EDUCATION BENEFITS POLICY: ESTIMATED $0.00 LOAN DEBT OFFERED</b><br/>
    <font size='9' color='#333333'>Based on current packaging, your institutional tuition and fees are covered fully via Chapter 33 veterans' benefits and non-repayable Federal Pell Grants; you currently have no federal student loan debt obligations scheduled for this cycle.</font><br/><br/>
    <b>VA Post-9/11 Certification Notice:</b><br/>
    <font size='8.5' color='#555555'>
    Tuition payments are sent by the VA directly to the institution. Any alterations in your scheduled enrollment or attendance must be reported immediately to prevent retroactive benefit recalculations and personal overpayment debts to the government.
    </font>
    """
elif is_va_student and "35" in va_type:
    loan_box_text = f"""
    <b>CHAPTER 35 REPAYMENT &amp; LEDGER DISCLOSURE: PAYMENT ARRANGEMENT PLAN</b><br/>
    <font size='9' color='#333333'>The VA does not pay the school directly under Chapter 35. Federal Pell Grants apply directly to your institutional bill; however, a payment arrangement plan is required to address any remaining balance.</font><br/><br/>
    <b>Student Payment Responsibility:</b><br/>
    <font size='8.5' color='#555555'>
    You remain responsible for resolving the estimated <b>${ch35_balance_due:,.2f}</b> account balance. You must coordinate an official payment plan with the financial aid office to manage this institutional obligation.
    </font>
    """
else:
    if ay1_quarters > 0 and loan_total > 0:
        avg_gross_q = loan_total / ay1_quarters
        avg_net_q = (sum(sub_disb_list) + sum(unsub_disb_list)) / ay1_quarters
        avg_fee_q = avg_gross_q - avg_net_q
    else:
        avg_gross_q = 0.00
        avg_net_q = 0.00
        avg_fee_q = 0.00
        
    loan_box_text = f"""
    <b>TOTAL ESTIMATED STUDENT LOAN DEBT TO BE REPAID: ${loan_total:,.2f}</b><br/>
    <font size='9' color='#333333'>This gross amount represents the total principal balance you are legally responsible for repaying, plus any accrued interest, once your grace period ends.</font><br/><br/>
    <b>IMPORTANT DISCLOSURE: Gross Award vs. Net Applied Amount</b><br/>
    <font size='8.5' color='#555555'>
    The U.S. Department of Education charges a mandatory <b>1.057% loan origination fee</b> on federal student loans. This fee is automatically deducted from your gross award before funds are applied to your student account ledger. Consequently, there is a minor variance between the gross award shown on your initial award letter and the actual net capital received by the school.<br/><br/>
    <b>Per-Disbursement Period Breakdown (Quarterly Average):</b><br/>
    &bull; Gross Loan Value Scheduled: <b>${avg_gross_q:,.2f}</b> | Mandatory Federal Fee Deducted (1.057%): <b>-${avg_fee_q:,.2f}</b> | Net Cash Applied to School Ledger: <b>${avg_net_q:,.2f}</b>
    </font>
    """

loan_box_table = Table([[Paragraph(loan_box_text, loan_box_style)]], colWidths=[540])
loan_box_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), LIGHT_BG),
    ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#002D62")),
    ('TOPPADDING', (0,0), (-1,-1), 8),
    ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ('LEFTPADDING', (0,0), (-1,-1), 10),
    ('RIGHTPADDING', (0,0), (-1,-1), 10),
]))
story.append(loan_box_table)

story.append(Paragraph("Critical Disclosures", section_style))

if is_va_student and "33" in va_type:
    d1 = "<b>VA Funding Timelines:</b> Post-9/11 tuition disbursements dispatch on a distinct federal timeline and post directly to the institution independently from Pell Grant schedules."
    d2 = "<b>Pell Grant Autonomy:</b> Your Pell Grant eligibility remains independent from GI Bill restrictions and applies directly to your institutional charges."
    d3 = "<b>Credit Balance Processing:</b> Eligible credit balance refunds are processed following federal regulatory guidelines once institutional hour milestones are confirmed."
elif is_va_student and "35" in va_type:
    d1 = "<b>Direct Student Payments:</b> The VA will distribute your Chapter 35 monthly allowance payments directly to your personal bank account, not to the school."
    d2 = "<b>Ledger Application:</b> Pell Grants apply to your institutional account first, reducing your immediate balance due to the business office."
    d3 = "<b>Payment Arrangements:</b> Please connect with the financial aid office to establish an approved payment plan to resolve your remaining institutional balance."
else:
    d1 = "<b>True Net Values:</b> Figures inside the timeline schedule above reflect the automatic removal of the mandatory 1.057% federal loan origination fee."
    d2 = "<b>30-Day Delay Rule:</b> First-time, first-year student borrowers are subject to a mandatory federal 30-day delayed disbursement hold starting from the first day of classes."
    d3 = "<b>Receipt Notifications:</b> Automated notifications indicate that federal funds have been received to clear your institutional ledger, rather than implying a direct cash distribution."

story.append(Paragraph(f"&bull; {d1}", disc_style))
story.append(Spacer(1, 3))
story.append(Paragraph(f"&bull; {d2}", disc_style))
story.append(Spacer(1, 3))
story.append(Paragraph(f"&bull; {d3}", disc_style))

doc.build(story)
# =====================================================================
# 4. COMPILE AND STITCH PORTFOLIO TO CUSTOM DESTINATION FOLDER
# =====================================================================
parent_folder_name = "FA_Info_Packets"
master_parent_dir = os.path.join(desktop_dir, parent_folder_name)
os.makedirs(master_parent_dir, exist_ok=True)

name_parts = student_name.split()
if len(name_parts) >= 2:
    last_name = name_parts[-1]
    first_name = " ".join(name_parts[:-1])
    formatted_name = f"{last_name}, {first_name}"
else:
    formatted_name = student_name

student_folder_name = f"{formatted_name} FA Info"
final_output_dir = os.path.join(master_parent_dir, student_folder_name)
os.makedirs(final_output_dir, exist_ok=True)

merger = PdfMerger()
merger.append(cover_pdf)

for flyer_name in selected_flyers:
    file_path = FLYER_MAP.get(flyer_name)
    if file_path and os.path.exists(file_path):
        merger.append(file_path)

output_filename = f"{formatted_name} FA Info.pdf"
final_pdf_path = os.path.join(final_output_dir, output_filename)

with open(final_pdf_path, "wb") as f_out:
    merger.write(f_out)
merger.close()

if os.path.exists(cover_pdf): 
    os.remove(cover_pdf)

print(f"Success! Finalized packet generated successfully at: {final_pdf_path}")

import os
import shutil
import openpyxl
import math

# Direct, explicit import optimized for Python environment
import pypdf
from pypdf import PdfWriter as PdfMerger

# Import ReportLab layout components
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# =====================================================================
# FORCE ABSOLUTE WORKING DIRECTORY
# =====================================================================
desktop_dir = r"C:\Users\cbaggarly\OneDrive - Goodwill Industries of MiGA and the CSRA\Desktop"
os.chdir(desktop_dir)

excel_file = "Macon Campus Package Database.xlsm"
temp_excel_snapshot = "temp_snapshot_package.xlsm"

# =====================================================================
# 1. SILENT EXCEL SCANNER (DYNAMIC PROGRAM TERMINOLOGY DETECTION)
# =====================================================================
selected_flyers = []
student_name = "Unknown_Student"
program_name = "Unknown_Program"
program_cost = 0.0
pell_award = 0.0
loan_total = 0.0
award_total = 0.0
va_type = "" 

q_costs = [] # Array to hold dynamic quarter charges pulled from Program_Master

def safe_float(val):
    if val is None:
        return 0.0
    val_str = str(val).strip()
    if val_str.startswith("#") or not val_str:
        return 0.0
    try:
        for char in ["$", ",", " ", " "]:  # Strip out currency symbols and hidden spaces
            val_str = val_str.replace(char, "")
        return float(val_str)
    except ValueError:
        return 0.0

try:
    shutil.copy2(excel_file, temp_excel_snapshot)
    wb = openpyxl.load_workbook(temp_excel_snapshot, data_only=True, read_only=True)
    
    if "Package_Data" in wb.sheetnames:
        ws = wb["Package_Data"]
        
        student_name = str(ws["B1"].value).strip() if ws["B1"].value is not None else "Unknown_Student"
        program_name = str(ws["B2"].value).strip() if ws["B2"].value is not None else "Unknown_Program"
        
        program_cost = safe_float(ws["B3"].value)
        pell_award = safe_float(ws["B4"].value)
        loan_total = safe_float(ws["B8"].value)   
        award_total = safe_float(ws["B9"].value)  
        
        if ws["F4"].value is not None:
            va_type = str(ws["F4"].value).strip()
        
        for row in range(2, 8):
            flyer_name_val = ws.cell(row=row, column=4).value   
            checked_val = ws.cell(row=row, column=5).value      
            
            if flyer_name_val is not None:
                flyer_name = str(flyer_name_val).strip()
                is_checked = False
                if checked_val is not None:
                    is_checked = str(checked_val).strip().upper() in ["X", "YES", "TRUE", "1", "1.0", "CHECK"]
                if is_checked:
                    selected_flyers.append(flyer_name)

    # Dynamic Charge Tracker mapping directly from master matrix
    if "Program_Master" in wb.sheetnames:
        pm_ws = wb["Program_Master"]
        for row in range(2, 25):
            row_prog = str(pm_ws.cell(row=row, column=1).value).strip().upper()
            if row_prog == program_name.upper().strip():
                # Pull across all possible structural terms
                q1_val = safe_float(pm_ws.cell(row=row, column=3).value)
                q2_val = safe_float(pm_ws.cell(row=row, column=4).value)
                q3_val = safe_float(pm_ws.cell(row=row, column=5).value)
                q4_val = safe_float(pm_ws.cell(row=row, column=6).value)
                
                # Dynamic append keeping zero-cost trailing fields clean
                raw_costs = [q1_val, q2_val, q3_val, q4_val]
                q_costs = [c for c in raw_costs if c > 0]
                break

    wb.close()
finally:
    if os.path.exists(temp_excel_snapshot):
        os.remove(temp_excel_snapshot)

if student_name == "nan" or not student_name or student_name.startswith("#"):
    student_name = "Unknown_Student"

FLYER_MAP = {
    "Dependent Student": os.path.join(desktop_dir, r"Packet_Flyers\Dependent_Plus_Flyer.pdf"),
    "Student Loans": os.path.join(desktop_dir, r"Packet_Flyers\Student_Loans_Flyer.pdf"),
    "VA Student": os.path.join(desktop_dir, r"Packet_Flyers\VA_Flyer.pdf"),
    "Scholarship Flyer": os.path.join(desktop_dir, r"Packet_Flyers\Scholarships_Flyer.pdf"),
    "Helms GFS App": os.path.join(desktop_dir, r"Packet_Flyers\Helms College General Fund Scholarship Application.pdf"),
    "ION Tuition": os.path.join(desktop_dir, r"Packet_Flyers\ION_Tuition_Flyer.pdf")
}

is_va_student = "VA Student" in selected_flyers
va_label = f"Chapter {33 if '33' in va_type else 35} {'Post-9/11' if '33' in va_type else 'DEA'}" if is_va_student else ""

# =====================================================================
# 2. NET LEDGER MATH ENGINE (RECONCILED CASCADING ARCHITECTURE)
# =====================================================================
origination_fee_rate = 0.01057
program_upper = program_name.upper()
is_clock_hour = "HVAC" in program_upper or "MSMA" in program_upper or "ELECTRICIAN" in program_upper

if not q_costs:
    q_costs = [program_cost / 2.0, program_cost / 2.0]

total_terms = len(q_costs)

# Step A: Split allocations into clean upfront variables
half_pell = round(pell_award / 2.0, 2)

if is_va_student and "33" in va_type:
    loan_total = 0.0
    total_net_loan = 0.0
    net_loan_disb = 0.0
    fee_per_disb = 0.0
    gross_per_disb = 0.0
else:
    # Anchor the per-disbursement values to exact rounded halves
    gross_per_disb = round(loan_total / 2.0, 2)
    fee_per_disb = round(gross_per_disb * origination_fee_rate, 2)
    net_loan_disb = round(gross_per_disb - fee_per_disb, 2)
    # Define the exact net pool to reconcile loose pennies against
    total_net_loan = round(loan_total - (fee_per_disb * 2), 2)

# Step B: Map chronological allocations with strict cent-tracking
pell_posts = [0.0] * total_terms
loan_posts = [0.0] * total_terms

if total_terms >= 1:
    pell_posts[0] = half_pell
    loan_posts[0] = net_loan_disb

if total_terms >= 3:
    pell_posts[2] = round(pell_award - half_pell, 2)
    loan_posts[2] = round(total_net_loan - net_loan_disb, 2)  # Ties out perfectly to total net loan pool
elif total_terms == 2:
    pell_posts[1] = round(pell_award - half_pell, 2)
    loan_posts[1] = round(total_net_loan - net_loan_disb, 2)

# Step C: Reconciled Timeline Ledger Loop with Whole-Dollar Rounding
timeline_ledger = []
running_balance = 0.0

for i in range(total_terms):
    term_charge = round(q_costs[i], 2)
    term_pell = round(pell_posts[i], 2)
    term_loan = round(loan_posts[i], 2)
    
    if is_va_student and "33" in va_type:
        running_balance = round(running_balance + term_charge - (term_charge + term_pell), 2)
    elif is_va_student and "35" in va_type:
        running_balance = round(running_balance + term_charge - term_pell, 2)
    else:
        running_balance = round(running_balance + term_charge - (term_pell + term_loan), 2)
        
    if abs(running_balance) < 0.01:
        display_balance = 0.0
    else:
        # Intercept and round up to the nearest whole integer per term milestone
        if running_balance < 0:
            display_balance = -float(math.ceil(abs(running_balance)))
        else:
            display_balance = float(math.ceil(running_balance))
        
    timeline_ledger.append({
        "label": f"Term / Quarter {i+1} Start",
        "charge": term_charge,
        "pell": term_pell,
        "loan": term_loan,
        "balance": display_balance  # Injects the clean rounded value straight into the table dataset
    })

# Step D: Extract Final Reconciled Balance Refund
balance_refund = timeline_ledger[-1]["balance"]

total_net_funding = pell_award if is_va_student else round(pell_award + total_net_loan, 2)
ch33_refund_value = pell_award

# =====================================================================
# 3. GENERATE COVER SHEET LAYOUT
# =====================================================================
cover_pdf = "temp_cover.pdf"
doc = SimpleDocTemplate(cover_pdf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=30, bottomMargin=30)
styles = getSampleStyleSheet()
story = []

NAVY = colors.HexColor("#002D62")
GOLD = colors.HexColor("#F4A261")
DARK_GRAY = colors.HexColor("#333333")
LIGHT_BG = colors.HexColor("#F4F6F9")

title_style = ParagraphStyle('TStyle', fontName='Helvetica-Bold', fontSize=20, leading=24, textColor=colors.white, alignment=1, spaceAfter=4)
subtitle_style = ParagraphStyle('SubStyle', fontName='Helvetica-Bold', fontSize=13, leading=16, textColor=GOLD, alignment=1, spaceAfter=2)
program_style = ParagraphStyle('PStyle', fontName='Helvetica', fontSize=10.5, leading=14, textColor=colors.white, alignment=1)
section_style = ParagraphStyle('SecStyle', fontName='Helvetica-Bold', fontSize=13, leading=16, textColor=NAVY, spaceBefore=12, spaceAfter=4)
th_style = ParagraphStyle('THStyle', fontName='Helvetica-Bold', fontSize=9.5, leading=12, textColor=colors.white, alignment=1)
td_style = ParagraphStyle('TDStyle', fontName='Helvetica', fontSize=9.5, leading=12, textColor=DARK_GRAY, alignment=1)
td_left = ParagraphStyle('TDLeft', fontName='Helvetica', fontSize=9.5, leading=12, textColor=DARK_GRAY, alignment=0)
outcome_style = ParagraphStyle('OutcomeStyle', fontName='Helvetica', fontSize=10.5, textColor=DARK_GRAY, leading=14)

# COMPLIANCE UPDATE: Clearly specifies that estimated packaging includes all award types.
funding_label = f"Total Federal Grants: ${total_net_funding:,.2f} ({va_label})" if is_va_student else f"Total Estimated Financial Aid Package: ${total_net_funding:,.2f}"

header_data = [
    [Paragraph("FINANCIAL AID PACKAGING OVERVIEW", title_style)],
    [Paragraph(f"Prepared For: {student_name} {f'({va_label})' if is_va_student else ''}", subtitle_style)],
    [Paragraph(f"Program: {program_name} | Program Cost: ${program_cost:,.2f} | {funding_label}", program_style)]
]
header_table = Table(header_data, colWidths=[540])
header_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), NAVY),
    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('TOPPADDING', (0,0), (-1,-1), 12),
    ('BOTTOMPADDING', (0,0), (-1,-1), 12),
]))
story.append(header_table)
story.append(Spacer(1, 8))

if is_va_student and "33" in va_type:
    # COMPLIANCE UPDATE: Swapped out "stipend refund" and "surplus" for "estimated credit balance refund"
    outcome_text = (
        f"<b>CHAPTER 33 POST-9/11 SUMMARY:</b> Your program instructional tuition and fee costs are certified "
        f"and paid directly to the school by the VA. Your Federal Pell Grant applies to your ledger, "
        f"creating an estimated credit balance refund of <b>${ch33_refund_value:,.2f}</b>. "
        f"<b>CREDIT BALANCE ISSUANCE POLICY:</b> Eligible credit balance refunds are processed "
        f"and distributed following federal guidelines <u>AFTER</u> your scheduled funding posts to your account ledger "
        f"and the student account reflects a clear $0.00 institutional balance."
    )
    box_bg = colors.HexColor("#E2F0D9") 
    box_line = colors.HexColor("#385723")
elif is_va_student and "35" in va_type:
    ch35_real_balance = program_cost - pell_award
    # COMPLIANCE UPDATE: Softened "payment contract required" to "payment arrangement plan required"
    outcome_text = (
        f"<b>CHAPTER 35 DEA SUMMARY:</b> The VA distributes monthly allowance stipends <b>directly to the student</b>, "
        f"not the school. Your Federal Pell Grant reduces your direct program costs on the institutional ledger, "
        f"leaving an estimated remaining balance of <b>${ch35_real_balance:,.2f}</b>. An institutional payment arrangement "
        f"plan is required for all recipients of Chapter 35 VA benefits in order to resolve the remaining balance."
    )
    box_bg = colors.HexColor("#FCE4D6") 
    box_line = colors.HexColor("#C65911")
else:
    if balance_refund <= 0:
        # COMPLIANCE UPDATE: Disclosures framing credit balances as estimated refunds under federal guidelines
        outcome_text = f"<b>ESTIMATED CREDIT BALANCE :</b> Your net financial aid awards completely clear your program costs, leaving an estimated credit balance of <b>${abs(balance_refund):,.2f}</b>. Credit balances are processed and issued following federal guidelines <u>AFTER</u> your final scheduled financial aid disbursement has posted to your ledger and your active student account shows a $0.00 institutional balance."
        box_bg = colors.HexColor("#E2F0D9") 
        box_line = colors.HexColor("#385723")
    else:
        # COMPLIANCE UPDATE: Standardized payment plan terminology
        outcome_text = f"<b>ESTIMATED OUT-OF-POCKET BALANCE DUE:</b> After applying your net financial aid disbursements against program costs, there remains an estimated net open balance of <b>${balance_refund:,.2f}</b> to be cleared for this academic cycle. <br/><br/><b>Payment Options Available:</b> A customized Institutional Payment Plan will be provided to break this balance down into manageable monthly installments."
        box_bg = colors.HexColor("#FCE4D6") 
        box_line = colors.HexColor("#C65911")

outcome_table = Table([[Paragraph(outcome_text, outcome_style)]], colWidths=[540])
outcome_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), box_bg),
    ('BOX', (0,0), (-1,-1), 1.2, box_line),
    ('TOPPADDING', (0,0), (-1,-1), 10),
    ('BOTTOMPADDING', (0,0), (-1,-1), 10),
    ('LEFTPADDING', (0,0), (-1,-1), 10),
    ('RIGHTPADDING', (0,0), (-1,-1), 10),
]))
story.append(outcome_table)

story.append(Paragraph("Account Ledger Timeline", section_style))

timeline_data = [
    [
        Paragraph("Term Milestone", th_style), 
        Paragraph("Tuition & Fees", th_style), 
        Paragraph("Pell Grant", th_style),
        Paragraph("Net Student Loans", th_style),
        Paragraph("Estimated Balance", th_style) # COMPLIANCE UPDATE: Changed "Rolling Balance" to "Estimated Balance"
    ]
]

for row in timeline_ledger:
    p_charge = row["charge"]
    p_pell = row["pell"]
    p_loan = row["loan"]
    p_bal = row["balance"]
    
    if is_va_student and "33" in va_type:
        bal_str = Paragraph("<font color='#385723'><b>VA Certified</b></font>", td_style)
    else:
        bal_str = Paragraph(f"-${abs(p_bal):,.2f} (Estimated Credit)" if p_bal < 0 else f"${p_bal:,.2f}", td_style)
        
    timeline_data.append([
        Paragraph(row["label"], td_left),
        Paragraph(f"${p_charge:,.2f}", td_style),
        Paragraph(f"${p_pell:,.2f}" if p_pell > 0 else "$0.00", td_style),
        Paragraph(f"${p_loan:,.2f}" if p_loan > 0 else "$0.00", td_style),
        bal_str
    ])

timeline_table = Table(timeline_data, colWidths=[130, 105, 100, 100, 105])
timeline_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,0), NAVY),
    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#D9D9D9")),
    ('TOPPADDING', (0,0), (-1,-1), 6),
    ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_BG]),
]))
story.append(timeline_table)

if is_clock_hour:
    hour_milestone = "440 clock hours" if "HVAC" in program_upper else "445.5 clock hours"
    gate_warning_text = f"⚠️ <b>CLOCK-HOUR ENROLLMENT REGULATORY DISCLOSURE:</b> This program operates on a federal clock-hour framework (Total Program Hours: {'880' if 'HVAC' in program_upper else '891'}). While institutional program costs are assessed on a term structure, your financial aid allocations are divided into two equal disbursements. The first half posts on Day 1. The second half <b>WILL NOT disburse until you have successfully completed {hour_milestone}</b>. Attendance directly dictates disbursement dates."
    
    gate_table = Table([[Paragraph(gate_warning_text, ParagraphStyle('GateStyle', fontName='Helvetica-Oblique', fontSize=8.5, leading=11, textColor=colors.HexColor("#7F6000")))]], colWidths=[540])
    gate_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#FFF2CC")),
        ('BOX', (0,0), (-1,-1), 0.8, colors.HexColor("#D6B656")),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 8),
        ('RIGHTPADDING', (0,0), (-1,-1), 8),
    ]))
    story.append(Spacer(1, 4))
    story.append(gate_table)

story.append(Spacer(1, 6))

if is_va_student and "33" in va_type:
    loan_box_text = f"""
    <b>MILITARY EDUCATION BENEFITS POLICY: ESTIMATED $0.00 LOAN DEBT OFFERED</b><br/>
    <font size='9' color='#333333'>Based on current packaging, your institutional tuition and fees are covered fully via Chapter 33 veterans' benefits and non-repayable Federal Pell Grants; you currently have no federal student loan debt obligations scheduled for this cycle.</font><br/><br/>
    <b>VA Post-9/11 Certification Notice:</b><br/>
    <font size='8.5' color='#555555'>
    Tuition payments are sent by the VA directly to the institution. Any alterations in your scheduled enrollment or attendance must be reported immediately to prevent retroactive benefit recalculations and personal overpayment debts to the government.
    </font>
    """
elif is_va_student and "35" in va_type:
    loan_box_text = f"""
    <b>CHAPTER 35 REPAYMENT & LEDGER DISCLOSURE: PAYMENT ARRANGEMENT PLAN</b><br/>
    <font size='9' color='#333333'>The VA does not pay the school directly under Chapter 35. Federal Pell Grants apply directly to your institutional bill; however, an institutional payment arrangement plan is required to address any remaining balance.</font><br/><br/>
    <b>Student Payment Responsibility:</b><br/>
    <font size='8.5' color='#555555'>
    You remain responsible for resolving the estimated <b>${balance_refund:,.2f}</b> account balance. You must coordinate an approved payment plan with the financial aid office to manage this institutional obligation over the course of your program.
    </font>
    """
else:
    loan_box_text = f"""
    <b>TOTAL ESTIMATED STUDENT LOAN DEBT TO BE REPAID: ${loan_total:,.2f}</b><br/>
    <font size='9' color='#333333'>This gross amount represents the total principal balance you are legally responsible for repaying, less accrued interest, once your grace period ends.</font><br/><br/>
    <b>IMPORTANT DISCLOSURE: Gross Award vs. Net Applied Amount</b><br/>
    <font size='8.5' color='#555555'>
    The U.S. Department of Education charges a <b>1.057% loan origination fee</b> on federal student loans. This fee is automatically deducted from your gross award before funds are applied to your student account ledger. Consequently, there is a minor variance between the gross award shown on your initial award letter and the actual net funding received by Helms College.<br/>
    <b>Per-Disbursement Period Breakdown (Two-Installment Split):</b><br/>
    &bull; Gross Loan Value Scheduled: <b>${gross_per_disb:,.2f}</b> | Origination Fee (1.057%): <b>-${fee_per_disb:,.2f}</b> | Net Cash Applied to School Ledger: <b>${net_loan_disb:,.2f}</b>
    </font>
    """

loan_box_style = ParagraphStyle('LoanBoxStyle', fontName='Helvetica', fontSize=10, textColor=NAVY, leading=13)
loan_box_table = Table([[Paragraph(loan_box_text, loan_box_style)]], colWidths=[540])
loan_box_table.setStyle(TableStyle([
    ('BACKGROUND', (0,0), (-1,-1), LIGHT_BG),
    ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#002D62")),
    ('TOPPADDING', (0,0), (-1,-1), 8),
    ('BOTTOMPADDING', (0,0), (-1,-1), 8),
    ('LEFTPADDING', (0,0), (-1,-1), 10),
    ('RIGHTPADDING', (0,0), (-1,-1), 10),
]))
story.append(loan_box_table)

story.append(Paragraph("Critical Disclosures", section_style))
disc_style = ParagraphStyle('DiscStyle', fontName='Helvetica', fontSize=9, textColor=DARK_GRAY, leading=13, leftIndent=15, firstLineIndent=-10)

if is_va_student and "33" in va_type:
    d1 = "<b>VA Funding Timelines:</b> Post-9/11 tuition disbursements dispatch on a distinct federal timeline and post directly to the institution independently from Pell Grant schedules."
    d2 = "<b>Pell Grant Autonomy:</b> Your Pell Grant eligibility remains independent from GI Bill restrictions and applies directly to your institutional charges."
    d3 = "<b>Credit Balance Processing:</b> Eligible credit balance refunds are processed following federal regulatory guidelines once institutional hour milestones are confirmed."
elif is_va_student and "35" in va_type:
    d1 = "<b>Direct Student Payments:</b> The VA will distribute your Chapter 35 monthly allowance payments directly to your personal bank account, not to the school."
    d2 = "<b>Ledger Application:</b> Pell Grants apply to your institutional account first, reducing your immediate balance due to the business office."
    d3 = "<b>Payment Arrangements:</b> Please connect with the financial aid office to establish an approved payment plan to resolve your remaining institutional balance."
else:
    d1 = "<b>True Net Values:</b> Figures inside the timeline schedule above reflect the automatic removal of the mandatory 1.057% federal loan origination fee."
    d2 = "<b>30-Day Delay Rule:</b> First-time, first-year student borrowers are subject to a mandatory federal 30-day delayed disbursement hold starting from the first day of classes."
    d3 = "<b>Receipt Notifications:</b> Automated notifications indicate that federal funds have been received to clear your institutional ledger, rather than implying a direct cash distribution."

story.append(Paragraph(f"&bull; {d1}", disc_style))
story.append(Spacer(1, 3))
story.append(Paragraph(f"&bull; {d2}", disc_style))
story.append(Spacer(1, 3))
story.append(Paragraph(f"&bull; {d3}", disc_style))

doc.build(story)

# =====================================================================
# 4. COMPILE AND STITCH PORTFOLIO TO CUSTOM DESTINATION FOLDER
# =====================================================================
parent_folder_name = "FA_Info_Packets"
master_parent_dir = os.path.join(desktop_dir, parent_folder_name)
os.makedirs(master_parent_dir, exist_ok=True)

name_parts = student_name.split()
if len(name_parts) >= 2:
    last_name = name_parts[-1]
    first_name = " ".join(name_parts[:-1])
    formatted_name = f"{last_name}, {first_name}"
else:
    formatted_name = student_name

student_folder_name = f"{formatted_name} FA Info"
final_output_dir = os.path.join(master_parent_dir, student_folder_name)
os.makedirs(final_output_dir, exist_ok=True)

merger = PdfMerger()
merger.append(cover_pdf)

for flyer_name in selected_flyers:
    file_path = FLYER_MAP.get(flyer_name)
    if file_path and os.path.exists(file_path):
        merger.append(file_path)

output_filename = f"{formatted_name} FA Info.pdf"
final_pdf_path = os.path.join(final_output_dir, output_filename)

with open(final_pdf_path, "wb") as f_out:
    merger.write(f_out)
merger.close()

if os.path.exists(cover_pdf): 
    os.remove(cover_pdf)

print(f"Success! Script executed flawlessly. Output saved to {final_pdf_path}")
