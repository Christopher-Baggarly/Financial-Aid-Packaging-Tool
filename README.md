**Federal Financial Aid Packaging & Packet Generation Engine**

**The Business Problem**

Career-college Financial Aid Offices are required to deliver standardized, regulatorily-compliant aid disclosure packets to every admitted student before class start. Manual packet assembly pulls staff off high-value work (cohort-default prevention, default counseling, regulatory updates) and produces inconsistent layouts across packets. Worse, because federal disbursement rules differ between clock-hour programs (HVAC, MSMA) and quarter-hour programs (Culinary Diploma, AAS), most offices either build one hacky spreadsheet that fails for the other program type or duplicate the work by hand. Neither is auditable, neither scales, and neither protects the institution when a federal program review asks how a Pell disbursement was calculated for a specific student.

**The Architecture & Logic**

This engine is two parallel packagers driven by a shared institutional config, producing publication-quality PDF packets that satisfy the federal disclosure requirements every Title IV institution owes its students.

_Two parallel scripts, two federal frameworks_. generate_package.py handles clock-hour programs (two equal COD payment periods, gated by hours-and-weeks thresholds per 34 CFR § 668.4(b)). generate_package_quarter.py handles quarter-hour programs (variable quarter disbursements plus Q4 anticipation for multi-year culinary diplomas). The split is deliberate: the federal math is different, and trying to force both into one script is how real FA offices end up with miscalculated Pell disbursements.

_Config-driven institutional identity_. School name, branding colors, accounts-manager contact, flyer map, origination fee rates, credit-balance refund window, and per-program clock-hour thresholds all live in config/institutional.json. The same engine that runs for one school can be pointed at another school by editing one config file. No code changes.

_COD integer disbursement math_. Federal Title IV disbursement requires integer dollar amounts that match across the school's ledger, the COD system, and the financial aid award letter. The packagers compute net disbursements per period using integer ceiling logic so the two-period Pell split lands on the exact award total without penny drift, then build the running ledger timeline that students see on page one of their packet.

_VA bifurcation_. Chapter 33 (Post-9/11) zeros out loan components, marks the ledger VA Certified, and produces a Pell-driven credit balance. Chapter 35 (DEA) keeps Pell applied to institutional charges, calculates the remaining balance the student owes the school directly, and flags a payment arrangement. Non-VA packets follow the standard Title IV packaging path. The same engine produces three structurally different packets from one decision tree.

_Interest capitalization modeling_. Page two of every packet that contains unsubsidized or PLUS debt projects the in-school interest that will capitalize at repayment entry, then shows both Standard 10-Year amortization and a RAP income-driven payment range (minimum $10.00/mo per current federal guidelines). Parent PLUS is shown in a separate visual block with maroon header because it is the parent borrower's legal obligation, not the student's.

_Clock-hour regulatory callout_. Every clock-hour packet carries a visible disclosure that the second disbursement will not post until the student has completed the required clock-hour and weeks threshold (34 CFR § 668.4(b)). This is the disclosure that protects the institution during a program review if a student later claims they were promised earlier access to funds.

**Federal Compliance Coverage**

34 CFR § 668.4(b) — clock-hour payment period definition and second-disbursement gating
34 CFR § 668.164(h) — 14-day Title IV credit balance refund timing
34 CFR § 685.303(b)(5) — first-time, first-year 30-day disbursement delay
Higher Education Act § 455 — Pell Grant maximum award and COD reporting
§ 668.202 — Direct Loan annual loan limits and grade-level progression
§ 674.33 — FSEOG selection criteria (policy-driven, not auto-funded by this engine)
RAP (Repayment Assistance Plan) — 2024-25 federal income-driven repayment framework

**Sample Output**
Three test packets are included in samples/:

_Portfolio, Test FA Info.pdf_ — MSMA clock-hour program, two payment periods, full Pell + Sub + Unsub loan schedule, $520 estimated credit balance.

_Portfolio, Test B FA Info.pdf_ — Culinary Diploma quarter program, four quarters with Q4 anticipation, Pell + Sub + Unsub + Parent PLUS, $301 estimated program credit balance after full 4-quarter projection.

_Portfolio, Test C FA Info.pdf_ — HVAC clock-hour program with Chapter 33 VA benefits. Tuition certified to the VA, Pell applied to ledger, zero loan debt offered.

Each packet contains a one-page regulatory cover sheet, an optional two-page repayment modeling page when loan components exist, and the selected institutional flyers (Dependent Student, VA Student, Scholarships, Loan Counseling, IonTuition) merged into one PDF ready for the student folder.

**Tech Stack**

Python 3.10+ — runtime

openpyxl — Excel input reading (the institutional packaging workbook)

reportlab — PDF cover sheet generation with regulatory callouts, ledger tables, and amortization modeling

pypdf — PDF merging (cover sheet + selected institutional flyers)

config-driven — no hardcoded paths, school names, or branding colors in source code

**Repository Layout**
.
├── README.md
├── requirements.txt
├── config/
│   └── institutional.json          # School identity, branding, fee rates, thresholds
├── generate_package.py             # Clock-hour program packager
├── generate_package_quarter.py     # Quarter-hour program packager
├── data/
│   ├── sample_student.xlsx         # Synthetic input workbook (FERPA-safe)
│   └── sample_student_quarter.xlsx # Synthetic input workbook (FERPA-safe)
├── flyers/                         # Selected informational flyers (PDFs)
├── samples/                        # Pre-generated sample output PDFs
└── examples/
    └── run_demo.py                 # End-to-end walkthrough script


**Quick Start**
# 1. Install dependencies
pip install -r requirements.txt

# 2. Verify config and sample data are in place
ls config/institutional.json
ls data/sample_student.xlsx

# 3. Generate a clock-hour packet
python generate_package.py

# 4. Generate a quarter-program packet
python generate_package_quarter.py

# 5. Output lands in ./output/FA_Info_Packets/<LastName>, <FirstName> FA Info/<LastName>, <FirstName> 

To swap to a different institution, edit config/institutional.json only. No code changes required.


**What This Engine Does Not Do**

This engine automates packaging math and disclosure delivery. It does not:

-Originate or transmit COD records to the Department of Education
-Replace PowerFAIDS, Banner, or other SIS systems of record
-Generate ISIRs or perform federal needs analysis
-Auto-award FSEOG (institutional policy decisions remain with the FA Director)
-The architectural decision was to automate the high-volume, error-prone calculation and disclosure work, and leave the policy judgment calls to humans.

**Author**

Built by Christopher Baggarly — Financial Aid administration with hands-on responsibility for Title IV packaging, COD reconciliation, and student-facing disclosure delivery. Engineered to solve a specific pain point (manual packet assembly eating staff hours) while preserving the federal compliance floor that an auditor would check.



