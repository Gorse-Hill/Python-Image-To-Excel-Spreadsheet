from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from datetime import datetime

# Create workbook and sheets
wb = Workbook()
discussion_ws = wb.active
discussion_ws.title = "Discussion and Meeting Tracker"
budget_ws = wb.create_sheet(title="Budget Management")

# =============================================================================
# DISCUSSION & MEETING TRACKER SHEET
# =============================================================================

# Set headers and formatting
discussion_headers = [
    "Discussion ID", "Date", "Topic", "Description", "Type",
    "Partner Organization/University", "Key Participants", "Location",
    "Outcomes/Actions", "Responsible Person", "Deadline", "Status", "Notes"
]

discussion_ws.append(discussion_headers)

# Set column widths
col_widths = [16, 12, 30, 40, 18, 30, 30, 20, 40, 20, 12, 12, 40]
for i, width in enumerate(col_widths, 1):
    discussion_ws.column_dimensions[get_column_letter(i)].width = width

# Add data validation
type_dv = DataValidation(
    type="list", 
    formula1='"Internal Discussion,University Meeting,Scholar Meeting,Public Forum"',
    showDropDown=True
)
discussion_ws.add_data_validation(type_dv)
type_dv.add(f"E2:E1000")

status_dv = DataValidation(
    type="list", 
    formula1='"Pending,In Progress,Completed"',
    showDropDown=True
)
discussion_ws.add_data_validation(status_dv)
status_dv.add(f"L2:L1000")

# Add auto-ID formula
discussion_ws["A2"] = '="SLIF-DISC-"&TEXT(ROW()-1,"000")'
discussion_ws["A2"].number_format = '@'  # Text format

# Add conditional formatting
red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

# Overdue actions
discussion_ws.conditional_formatting.add(
    f"K2:K1000",
    FormulaRule(
        formula=[f'AND(K2<TODAY(), L2<>"Completed")'],
        stopIfTrue=False,
        fill=red_fill
    )
)

# Upcoming deadlines (within 7 days)
discussion_ws.conditional_formatting.add(
    f"K2:K1000",
    FormulaRule(
        formula=[f'AND(K2>=TODAY(), K2<=TODAY()+7, L2="Pending")'],
        stopIfTrue=False,
        fill=yellow_fill
    )
)

# Freeze header row
discussion_ws.freeze_panes = "A2"

# =============================================================================
# BUDGET MANAGEMENT SHEET
# =============================================================================

# Set headers
budget_headers = [
    "Budget ID", "Discussion ID", "Date", "Description", "Category", "Type",
    "Amount (USD)", "Payment Method", "Receipt/Reference", "Status"
]

budget_ws.append(budget_headers)

# Set column widths
budget_widths = [14, 14, 12, 40, 16, 12, 14, 16, 20, 12]
for i, width in enumerate(budget_widths, 1):
    budget_ws.column_dimensions[get_column_letter(i)].width = width

# Add data validation
category_dv = DataValidation(
    type="list", 
    formula1='"Travel,Accommodation,Venue,Supplies,Honorarium,Marketing,Other"',
    showDropDown=True
)
budget_ws.add_data_validation(category_dv)
category_dv.add(f"E2:E1000")

type_dv_budget = DataValidation(
    type="list", 
    formula1='"Expense,Income,Allocation"',
    showDropDown=True
)
budget_ws.add_data_validation(type_dv_budget)
type_dv_budget.add(f"F2:F1000")

method_dv = DataValidation(
    type="list", 
    formula1='"Cash,Bank Transfer,Mobile Money"',
    showDropDown=True
)
budget_ws.add_data_validation(method_dv)
method_dv.add(f"H2:H1000")

status_dv_budget = DataValidation(
    type="list", 
    formula1='"Planned,Pending,Completed"',
    showDropDown=True
)
budget_ws.add_data_validation(status_dv_budget)
status_dv_budget.add(f"J2:J1000")

# Add Discussion ID dropdown (references first sheet)
budget_ws["B2"] = '=IFERROR(INDEX(\'Discussion and Meeting Tracker\'!A:A, MATCH(1, (\'Discussion and Meeting Tracker\'!A:A<>"")*1, 0), "")'
budget_ws["B2"].number_format = '@'

# Add auto-ID formula
budget_ws["A2"] = '="SLIF-BUDG-"&TEXT(ROW()-1,"000")'
budget_ws["A2"].number_format = '@'  # Text format

# Add summary formulas
budget_ws["M1"] = "BUDGET SUMMARY"
budget_ws["M2"] = "Total Allocated:"
budget_ws["N2"] = '=SUMIF(F:F,"Income",G:G)+SUMIF(F:F,"Allocation",G:G)'
budget_ws["M3"] = "Total Expenses:"
budget_ws["N3"] = '=SUMIF(F:F,"Expense",G:G)'
budget_ws["M4"] = "Remaining Balance:"
budget_ws["N4"] = "=N2+N3"  # Since expenses are negative

# Format summary
summary_fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
for cell in ["M1", "M2", "M3", "M4", "N2", "N3", "N4"]:
    budget_ws[cell].fill = summary_fill
budget_ws["N2"].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
budget_ws["N3"].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'
budget_ws["N4"].number_format = '"$"#,##0.00_);[Red]("$"#,##0.00)'

# Add conditional formatting
low_balance_rule = FormulaRule(
    formula=["$N$4<500"],
    stopIfTrue=False,
    fill=PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
)
budget_ws.conditional_formatting.add("N4", low_balance_rule)

overdue_payment_rule = FormulaRule(
    formula=['AND($J2="Pending", $C2<TODAY())'],
    stopIfTrue=False,
    font=Font(color="FF0000")
)
budget_ws.conditional_formatting.add("C2:C1000", overdue_payment_rule)

# Freeze header row
budget_ws.freeze_panes = "A2"

# =============================================================================
# FINALIZE AND SAVE
# =============================================================================

# Format headers
header_fill = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
header_alignment = Alignment(horizontal="center", wrap_text=True)

for ws in [discussion_ws, budget_ws]:
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment

# Set wrap text for long columns
for col in ["D", "I", "M"]:  # Description, Outcomes/Actions, Notes
    for cell in discussion_ws[col]:
        cell.alignment = Alignment(wrap_text=True)

# Save the workbook
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
filename = f"SLIF_Activities_Tracker_{timestamp}.xlsx"
wb.save(filename)

print(f"Template created successfully: {filename}")
