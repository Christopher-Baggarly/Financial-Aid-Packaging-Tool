# ===========================================================================
# QUARTER-PROGRAM DATA READ (replaces clock-hour data read block)
# ===========================================================================

if "Package_Data" in wb.sheetnames:
    ws = wb["Package_Data"]
    student_name = str(ws["B18"].value).strip() if ws["B18"].value else "Unknown_Student"
    program_name = str(ws["B19"].value).strip() if ws["B19"].value else "Unknown_Program"
    program_cost = safe_float(ws["B20"].value)
    pell_award = safe_float(ws["B21"].value)
    sub_gross = safe_float(ws["B22"].value)
    unsub_gross = safe_float(ws["B23"].value)
    q4_anticipated_pell = safe_float(ws["B24"].value)
    q4_anticipated_sub = safe_float(ws["B25"].value)
    q4_anticipated_unsub = safe_float(ws["B26"].value)
    interest_rate_input = _read_rate(ws["B27"].value, program_meta["default_student_interest_rate"])
    parent_plus_gross = safe_float(ws["B28"].value)
    plus_interest_rate_input = _read_rate(
        ws["B29"].value,
        round(program_meta["default_student_interest_rate"] + program_meta["plus_interest_rate_spread"], 4),
    )
    loan_total = safe_float(ws["B31"].value) or (
        sub_gross + unsub_gross + q4_anticipated_sub + q4_anticipated_unsub + parent_plus_gross
    )
    if ws["F4"].value is not None:
        va_type = str(ws["F4"].value).strip()
    # Same flyer selection loop as clock-hour

# ===========================================================================
# QUARTER PROGRAM DETECTION
# ===========================================================================

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

# ===========================================================================
# QUARTER-COD INTEGER DISBURSEMENT ENGINE
# ===========================================================================

# Pell split across AY1 quarters
pell_base = int(pell_award / ay1_quarters) if ay1_quarters > 0 else 0
pell_rem = int(round(pell_award - (pell_base * ay1_quarters))) if ay1_quarters > 0 else 0
pell_disb_list = [pell_base] * ay1_quarters
for i in range(pell_rem):
    pell_disb_list[i] += 1

# Sub, Unsub, PLUS: same integer ceil + remainder distribution
sub_net_total = math.ceil(sub_gross * (1.0 - origination_fee_rate))
sub_base = int(sub_net_total / ay1_quarters) if ay1_quarters > 0 else 0
sub_rem = int(round(sub_net_total - (sub_base * ay1_quarters))) if ay1_quarters > 0 else 0
sub_disb_list = [sub_base] * ay1_quarters
for i in range(sub_rem):
    sub_disb_list[i] += 1

unsub_net_total = math.ceil(unsub_gross * (1.0 - origination_fee_rate))
unsub_base = int(unsub_net_total / ay1_quarters) if ay1_quarters > 0 else 0
unsub_rem = int(round(unsub_net_total - (unsub_base * ay1_quarters))) if ay1_quarters > 0 else 0
unsub_disb_list = [sub_base] * ay1_quarters  # NOTE: matches existing original behavior
for i in range(unsub_rem):
    unsub_disb_list[i] += 1

plus_net_total = math.ceil(parent_plus_gross * (1.0 - plus_fee_rate))
plus_base = int(plus_net_total / ay1_quarters) if ay1_quarters > 0 else 0
plus_rem = int(round(plus_net_total - (plus_base * ay1_quarters))) if ay1_quarters > 0 else 0
plus_disb_list = [plus_base] * ay1_quarters
for i in range(plus_rem):
    plus_disb_list[i] += 1

quarterly_net_aid = []
for i in range(ay1_quarters):
    quarterly_net_aid.append(
        pell_disb_list[i] + sub_disb_list[i] + unsub_disb_list[i] + plus_disb_list[i]
    )

running_balances = []
current_balance = 0.0
for i in range(ay1_quarters):
    q_charge = q_costs[i] if i < len(q_costs) else 0.0
    q_aid = quarterly_net_aid[i]
    current_balance = current_balance + q_charge - q_aid
    running_balances.append(round(current_balance, 2))

q3_ending_balance = running_balances[-1] if running_balances else 0.0

# Q4 anticipation extension (multi-year diplomas only)
if is_culinary_diploma:
    q4_charge = q_costs[3] if len(q_costs) > 3 else 0.0
    q4_net_sub = math.ceil(q4_anticipated_sub * (1.0 - origination_fee_rate))
    q4_net_unsub = math.ceil(q4_anticipated_unsub * (1.0 - origination_fee_rate))
    q4_true_combined_net_aid = q4_anticipated_pell + q4_net_sub + q4_net_unsub
    q4_ending_balance = q3_ending_balance + q4_charge - q4_true_combined_net_aid
    balance_refund = q3_ending_balance
    total_net_funding = sum(quarterly_net_aid) + q4_true_combined_net_aid
else:
    balance_refund = current_balance
    total_net_funding = sum(quarterly_net_aid)
    q4_ending_balance = 0.0
    q4_true_combined_net_aid = 0.0

# ===========================================================================
# TIMELINE TABLE — QUARTER MILESTONES (replaces clock-hour timeline block)
# ===========================================================================

timeline_data = [[
    Paragraph("Quarter Milestone", th_style),
    Paragraph("Tuition & Fees", th_style),
    Paragraph("Pell Grant", th_style),
    Paragraph("Sub. Loan (Net)", th_style),
    Paragraph("Unsub. Loan (Net)", th_style),
    Paragraph("Estimated Balance", th_style),
]]
for i in range(ay1_quarters):
    q_charge_val = q_costs[i] if i < len(q_costs) else 0.0
    q_pell_val = float(pell_disb_list[i]) if i < len(pell_disb_list) else 0.0
    q_sub_val = float(sub_disb_list[i]) if i < len(sub_disb_list) else 0.0
    q_unsub_val = float(unsub_disb_list[i]) if i < len(unsub_disb_list) else 0.0
    q_bal_val = running_balances[i] if i < len(running_balances) else 0.0

    if is_va_student and "33" in va_type:
        bal_display = Paragraph("<font color='#385723'><b>VA Certified</b></font>", td_style)
    elif q_bal_val < 0:
        rounded_q_bal = math.ceil(abs(q_bal_val))
        bal_display = Paragraph(f"-${rounded_q_bal:,.0f} (Estimated Credit)", td_style)
    else:
        bal_display = Paragraph(f"${q_bal_val:,.2f}", td_style)

    timeline_data.append([
        Paragraph(f"Quarter {i+1} Start", td_left),
        Paragraph(f"${q_charge_val:,.2f}", td_style),
        Paragraph(f"${q_pell_val:,.2f}" if q_pell_val > 0 else "$0.00", td_style),
        Paragraph(f"${q_sub_val:,.2f}" if q_sub_val > 0 else "$0.00", td_style),
        Paragraph(f"${q_unsub_val:,.2f}" if q_unsub_val > 0 else "$0.00", td_style),
        bal_display,
    ])

# Q4 row (culinary diplomas only)
if is_culinary_diploma:
    q4_charge_val = q_costs[3] if len(q_costs) > 3 else 0.0
    if is_va_student and "33" in va_type:
        q4_bal_display = Paragraph("<font color='#385723'><b>VA Certified</b></font>", td_style)
    elif q4_ending_balance < 0:
        q4_bal_display = Paragraph(f"-${math.ceil(abs(q4_ending_balance)):,.0f} (Estimated Credit)", td_style)
    else:
        q4_bal_display = Paragraph(f"${q4_ending_balance:,.2f}", td_style)
    timeline_data.append([
        Paragraph("Quarter 4 Start<br/><font size='7' color='#555555'>(Next Aid Year)</font>", td_left),
        Paragraph(f"${q4_charge_val:,.2f}", td_style),
        Paragraph(f"${q4_anticipated_pell:,.2f}" if q4_anticipated_pell > 0 else "$0.00", td_style),
        Paragraph(f"${q4_net_sub:,.2f}" if q4_net_sub > 0 else "$0.00", td_style),
        Paragraph(f"${q4_net_unsub:,.2f}" if q4_net_unsub > 0 else "$0.00", td_style),
        q4_bal_display,
    ])

# Use 6-column table for quarter programs
timeline_table = Table(timeline_data, colWidths=[105, 87, 87, 87, 87, 87])

# ===========================================================================
# REPAYMENT MODELING — uses total_sub_all_years (sub + q4_anticipated_sub)
# ===========================================================================

total_sub_all_years = sub_gross + (q4_anticipated_sub if is_culinary_diploma else 0.0)
total_unsub_all_years = unsub_gross + (q4_anticipated_unsub if is_culinary_diploma else 0.0)
in_school_grace_months = 18 if is_culinary_diploma else 15

# (Same amortization math but using total_sub_all_years / total_unsub_all_years
# instead of sub_gross / unsub_gross in the Standard and RAP tables)