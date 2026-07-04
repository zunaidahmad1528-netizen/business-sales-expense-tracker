from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

wb = Workbook()

# Color palette
NAVY = "1F3864"
LIGHT_BLUE = "D9E2F3"
GREEN = "C6EFCE"
RED_FILL = "FFC7CE"
WHITE_FONT = Font(color="FFFFFF", bold=True, name="Arial", size=11)
HEADER_FILL = PatternFill("solid", start_color=NAVY)
TITLE_FONT = Font(bold=True, size=16, color=NAVY, name="Arial")
SUBTITLE_FONT = Font(italic=True, size=10, color="595959", name="Arial")
LABEL_FONT = Font(bold=True, name="Arial", size=11)
NORMAL_FONT = Font(name="Arial", size=11)
BLUE_INPUT = Font(color="0000FF", name="Arial", size=11)
BLACK_FORMULA = Font(color="000000", name="Arial", size=11)
thin = Side(style="thin", color="B7B7B7")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# ============ SHEET 1: INSTRUCTIONS ============
ws0 = wb.active
ws0.title = "Instructions"
ws0.sheet_view.showGridLines = False
ws0.column_dimensions['A'].width = 4
ws0.column_dimensions['B'].width = 90

ws0['B2'] = "Business Sales & Expense Tracker"
ws0['B2'].font = TITLE_FONT
ws0['B3'] = "Auto-calculating dashboard — just enter your data, everything else updates itself"
ws0['B3'].font = SUBTITLE_FONT

steps = [
    ("How to use this sheet", LABEL_FONT),
    ("1. Go to the 'Sales Entry' tab and add one row per sale (date, product, quantity, price).", NORMAL_FONT),
    ("2. Go to the 'Expense Entry' tab and add one row per expense (date, category, amount).", NORMAL_FONT),
    ("3. Open the 'Dashboard' tab — totals, profit, and monthly summary update automatically.", NORMAL_FONT),
    ("4. Do not edit cells with black formula text (columns marked 'Auto') — they calculate on their own.", NORMAL_FONT),
    ("5. Only type in cells shown in blue — those are your input cells.", NORMAL_FONT),
    ("", NORMAL_FONT),
    ("Tips", LABEL_FONT),
    ("- Use the dropdown in the Category column to keep expense names consistent.", NORMAL_FONT),
    ("- The Dashboard automatically pulls the latest numbers — no manual updating needed.", NORMAL_FONT),
    ("- You can add as many rows as you like; formulas are already extended to row 500.", NORMAL_FONT),
]
r = 5
for text, font in steps:
    ws0[f'B{r}'] = text
    ws0[f'B{r}'].font = font
    r += 1

# ============ SHEET 2: SALES ENTRY ============
ws1 = wb.create_sheet("Sales Entry")
ws1.sheet_view.showGridLines = False
headers1 = ["Date", "Product Name", "Quantity Sold", "Price per Unit (₹)", "Total Sale (₹) [Auto]"]
widths1 = [14, 28, 16, 20, 20]
for i, (h, w) in enumerate(zip(headers1, widths1), start=1):
    col = get_column_letter(i)
    ws1.column_dimensions[col].width = w
    cell = ws1[f'{col}1']
    cell.value = h
    cell.font = WHITE_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
ws1.row_dimensions[1].height = 30

for row in range(2, 501):
    ws1[f'A{row}'].number_format = 'DD-MMM-YYYY'
    ws1[f'D{row}'].number_format = '₹#,##0'
    ws1[f'E{row}'] = f'=IF(AND(C{row}<>"",D{row}<>""),C{row}*D{row},"")'
    ws1[f'E{row}'].font = BLACK_FORMULA
    ws1[f'E{row}'].number_format = '₹#,##0'
    for col in ['A', 'B', 'C', 'D']:
        ws1[f'{col}{row}'].font = BLUE_INPUT
    for col in ['A', 'B', 'C', 'D', 'E']:
        ws1[f'{col}{row}'].border = BORDER

# ============ SHEET 3: EXPENSE ENTRY ============
ws2 = wb.create_sheet("Expense Entry")
ws2.sheet_view.showGridLines = False
headers2 = ["Date", "Category", "Description", "Amount (₹)"]
widths2 = [14, 20, 30, 18]
for i, (h, w) in enumerate(zip(headers2, widths2), start=1):
    col = get_column_letter(i)
    ws2.column_dimensions[col].width = w
    cell = ws2[f'{col}1']
    cell.value = h
    cell.font = WHITE_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER
ws2.row_dimensions[1].height = 30

dv = DataValidation(type="list", formula1='"Rent,Salaries,Inventory/Stock,Utilities,Marketing,Transport,Maintenance,Other"', allow_blank=True)
ws2.add_data_validation(dv)

for row in range(2, 501):
    ws2[f'A{row}'].number_format = 'DD-MMM-YYYY'
    ws2[f'D{row}'].number_format = '₹#,##0'
    for col in ['A', 'B', 'C', 'D']:
        ws2[f'{col}{row}'].font = BLUE_INPUT
        ws2[f'{col}{row}'].border = BORDER
    dv.add(ws2[f'B{row}'])

# ============ SHEET 4: DASHBOARD ============
ws3 = wb.create_sheet("Dashboard")
ws3.sheet_view.showGridLines = False
ws3.column_dimensions['A'].width = 4
for col in ['B', 'C', 'D', 'E']:
    ws3.column_dimensions[col].width = 22

ws3['B2'] = "Business Dashboard"
ws3['B2'].font = TITLE_FONT
ws3['B3'] = "Auto-updates from Sales Entry and Expense Entry tabs"
ws3['B3'].font = SUBTITLE_FONT

# Summary cards
summary_labels = ["Total Sales (₹)", "Total Expenses (₹)", "Net Profit (₹)", "Profit Margin (%)"]
summary_formulas = [
    "=SUM('Sales Entry'!E2:E500)",
    "=SUM('Expense Entry'!D2:D500)",
    "=C5-C6",
    '=IF(C5=0,0,C7/C5)'
]
row_start = 5
for i, (label, formula) in enumerate(zip(summary_labels, summary_formulas)):
    r = row_start + i
    ws3[f'B{r}'] = label
    ws3[f'B{r}'].font = LABEL_FONT
    ws3[f'C{r}'] = formula
    ws3[f'C{r}'].font = Font(bold=True, size=13, color=NAVY, name="Arial")
    ws3[f'C{r}'].border = BORDER
    ws3[f'B{r}'].border = BORDER
    if "Margin" in label:
        ws3[f'C{r}'].number_format = '0.0%'
    else:
        ws3[f'C{r}'].number_format = '₹#,##0'
    ws3[f'B{r}'].fill = PatternFill("solid", start_color=LIGHT_BLUE)

# Conditional-style profit highlight (manual fill based on sign not needed; keep simple)

# Expense by Category summary
ws3['B11'] = "Expense Breakdown by Category"
ws3['B11'].font = LABEL_FONT
cats = ["Rent", "Salaries", "Inventory/Stock", "Utilities", "Marketing", "Transport", "Maintenance", "Other"]
ws3['B12'] = "Category"
ws3['C12'] = "Total (₹)"
for c in ['B12', 'C12']:
    ws3[c].font = WHITE_FONT
    ws3[c].fill = HEADER_FILL
    ws3[c].border = BORDER
    ws3[c].alignment = Alignment(horizontal="center")

for i, cat in enumerate(cats):
    r = 13 + i
    ws3[f'B{r}'] = cat
    ws3[f'B{r}'].font = NORMAL_FONT
    ws3[f'B{r}'].border = BORDER
    ws3[f'C{r}'] = f"=SUMIF('Expense Entry'!$B$2:$B$500,B{r},'Expense Entry'!$D$2:$D$500)"
    ws3[f'C{r}'].font = BLACK_FORMULA
    ws3[f'C{r}'].number_format = '₹#,##0'
    ws3[f'C{r}'].border = BORDER

# Monthly Sales Summary
ws3['E11'] = "Monthly Sales Summary"
ws3['E11'].font = LABEL_FONT
ws3['E12'] = "Month"
ws3['F12'] = "Sales (₹)"
ws3.column_dimensions['F'].width = 18
for c in ['E12', 'F12']:
    ws3[c].font = WHITE_FONT
    ws3[c].fill = HEADER_FILL
    ws3[c].border = BORDER
    ws3[c].alignment = Alignment(horizontal="center")

months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
for i, m in enumerate(months):
    r = 13 + i
    month_num = i + 1
    ws3[f'E{r}'] = m
    ws3[f'E{r}'].font = NORMAL_FONT
    ws3[f'E{r}'].border = BORDER
    ws3[f'F{r}'] = f'=SUMPRODUCT((\'Sales Entry\'!$A$2:$A$500<>"")*(MONTH(\'Sales Entry\'!$A$2:$A$500)={month_num})*\'Sales Entry\'!$C$2:$C$500*\'Sales Entry\'!$D$2:$D$500)'
    ws3[f'F{r}'].font = BLACK_FORMULA
    ws3[f'F{r}'].number_format = '₹#,##0'
    ws3[f'F{r}'].border = BORDER

wb.save('/home/claude/template/Business_Sales_Expense_Tracker.xlsx')
print("saved")
