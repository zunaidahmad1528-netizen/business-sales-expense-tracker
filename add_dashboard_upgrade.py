from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList

wb = load_workbook('Business_Sales_Expense_Tracker_v3.xlsx')

NAVY = "1F3864"
LIGHT_BLUE = "D9E2F3"
GREEN_FILL = "E2EFDA"
GOLD_FILL = "FFF2CC"
WHITE_FONT = Font(color="FFFFFF", bold=True, name="Arial", size=11)
HEADER_FILL = PatternFill("solid", start_color=NAVY)
LABEL_FONT = Font(bold=True, name="Arial", size=11)
NORMAL_FONT = Font(name="Arial", size=11)
BLUE_INPUT = Font(color="0000FF", name="Arial", size=11)
INSIGHT_FONT = Font(bold=True, name="Arial", size=11, color=NAVY)
thin = Side(style="thin", color="B7B7B7")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

ws3 = wb['Dashboard']

# clear stray old values first (before merging cells) to avoid MergedCell write errors
for row in range(11, 25):
    for col in ['B', 'C', 'E', 'F']:
        ws3[f'{col}{row}'] = None

# ============ INSIGHTS BOX ============
ws3['B10'] = "Key Insights (Auto-Updates)"
ws3['B10'].font = LABEL_FONT

ws3.merge_cells('B11:E11')
ws3['B11'] = '="Top Seller: "&IFERROR(INDEX(H42:H71,MATCH(MAX(I42:I71),I42:I71,0)),"—")&"  ("&MAX(I42:I71)&" units sold)"'
ws3['B11'].font = INSIGHT_FONT
ws3['B11'].fill = PatternFill("solid", start_color=GREEN_FILL)
ws3['B11'].alignment = Alignment(vertical="center", indent=1)
ws3['B11'].border = BORDER

ws3.merge_cells('B12:E12')
ws3['B12'] = '="Highest Expense: "&IFERROR(INDEX(B14:B21,MATCH(MAX(C14:C21),C14:C21,0)),"—")&"  (₹"&TEXT(MAX(C14:C21),"#,##0")&")"'
ws3['B12'].font = INSIGHT_FONT
ws3['B12'].fill = PatternFill("solid", start_color=GOLD_FILL)
ws3['B12'].alignment = Alignment(vertical="center", indent=1)
ws3['B12'].border = BORDER

ws3.merge_cells('B13:E13')
ws3['B13'] = '="Status: "&IF(C7>0,"Profitable","Loss-making — review your expenses")'
ws3['B13'].font = INSIGHT_FONT
ws3['B13'].fill = PatternFill("solid", start_color=LIGHT_BLUE)
ws3['B13'].alignment = Alignment(vertical="center", indent=1)
ws3['B13'].border = BORDER

# shift existing Category Breakdown & Monthly Summary down by 3 rows (was row 11-24, now 14-27)
# Category Breakdown header was B11/row12 header/13-20 data -> now B14/row15 header/16-23 data
ws3['B14'] = "Expense Breakdown by Category"
ws3['B14'].font = LABEL_FONT
ws3['B15'] = "Category"
ws3['C15'] = "Total (₹)"
for c in ['B15', 'C15']:
    ws3[c].font = WHITE_FONT
    ws3[c].fill = HEADER_FILL
    ws3[c].border = BORDER
    ws3[c].alignment = Alignment(horizontal="center")

cats = ["Rent", "Salaries", "Inventory/Stock", "Utilities", "Marketing", "Transport", "Maintenance", "Other"]
for i, cat in enumerate(cats):
    old_r = 13 + i
    new_r = 16 + i
    ws3[f'B{new_r}'] = cat
    ws3[f'B{new_r}'].font = NORMAL_FONT
    ws3[f'B{new_r}'].border = BORDER
    ws3[f'C{new_r}'] = f"=SUMIF('Expense Entry'!$B$2:$B$500,B{new_r},'Expense Entry'!$D$2:$D$500)"
    ws3[f'C{new_r}'].font = Font(color="000000", name="Arial", size=11)
    ws3[f'C{new_r}'].number_format = '₹#,##0'
    ws3[f'C{new_r}'].border = BORDER

# Monthly summary shift: old E11 title/E12 header/E13:F24 -> new E14 title/E15 header/E16:F27
ws3['E14'] = "Monthly Sales Summary"
ws3['E14'].font = LABEL_FONT
ws3['E15'] = "Month"
ws3['F15'] = "Sales (₹)"
for c in ['E15', 'F15']:
    ws3[c].font = WHITE_FONT
    ws3[c].fill = HEADER_FILL
    ws3[c].border = BORDER
    ws3[c].alignment = Alignment(horizontal="center")

months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
for i, m in enumerate(months):
    old_r = 13 + i
    new_r = 16 + i
    month_num = i + 1
    ws3[f'E{new_r}'] = m
    ws3[f'E{new_r}'].font = NORMAL_FONT
    ws3[f'E{new_r}'].border = BORDER
    ws3[f'F{new_r}'] = f'=SUMPRODUCT((\'Sales Entry\'!$A$2:$A$500<>"")*(MONTH(\'Sales Entry\'!$A$2:$A$500)={month_num})*\'Sales Entry\'!$C$2:$C$500*\'Sales Entry\'!$D$2:$D$500)'
    ws3[f'F{new_r}'].font = Font(color="000000", name="Arial", size=11)
    ws3[f'F{new_r}'].number_format = '₹#,##0'
    ws3[f'F{new_r}'].border = BORDER

# ============ CHARTS ============
bar = BarChart()
bar.type = "col"
bar.title = "Monthly Sales Trend"
bar.y_axis.title = "Sales (₹)"
bar.x_axis.title = "Month"
bar.style = 10
data = Reference(ws3, min_col=6, min_row=15, max_row=27)
cats_ref = Reference(ws3, min_col=5, min_row=16, max_row=27)
bar.add_data(data, titles_from_data=True)
bar.set_categories(cats_ref)
bar.height = 8
bar.width = 18
bar.legend = None
ws3.add_chart(bar, "B30")

pie = PieChart()
pie.title = "Expense Breakdown by Category"
data2 = Reference(ws3, min_col=3, min_row=15, max_row=23)
cats2 = Reference(ws3, min_col=2, min_row=16, max_row=23)
pie.add_data(data2, titles_from_data=True)
pie.set_categories(cats2)
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
pie.height = 8
pie.width = 12
ws3.add_chart(pie, "B50")

# ============ CUSTOM FIELD COLUMNS ============
ws1 = wb['Sales Entry']
ws1.column_dimensions['F'].width = 22
ws1['F1'] = "Custom Field (optional)"
ws1['F1'].font = WHITE_FONT
ws1['F1'].fill = HEADER_FILL
ws1['F1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws1['F1'].border = BORDER
for row in range(2, 501):
    ws1[f'F{row}'].font = BLUE_INPUT
    ws1[f'F{row}'].border = BORDER

ws2 = wb['Expense Entry']
ws2.column_dimensions['E'].width = 22
ws2['E1'] = "Custom Field (optional)"
ws2['E1'].font = WHITE_FONT
ws2['E1'].fill = HEADER_FILL
ws2['E1'].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws2['E1'].border = BORDER
for row in range(2, 501):
    ws2[f'E{row}'].font = BLUE_INPUT
    ws2[f'E{row}'].border = BORDER

# ============ INSTRUCTIONS UPDATE ============
ws0 = wb['Instructions']
ws0['B21'] = "9. Want to track something extra (customer name, payment mode, notes)? Use the 'Custom Field'"
ws0['B21'].font = NORMAL_FONT
ws0['B22'] = "   column already added in Sales Entry / Expense Entry — just rename the header in row 1 and fill it in."
ws0['B22'].font = NORMAL_FONT
ws0['B23'] = "   This column is safe to edit and won't break any formulas."
ws0['B23'].font = NORMAL_FONT
ws0['B24'] = "10. Dashboard now shows auto-insights (top seller, highest expense, profit status) and two charts —"
ws0['B24'].font = NORMAL_FONT
ws0['B25'] = "    a monthly sales bar chart and an expense category pie chart. Both update automatically."
ws0['B25'].font = NORMAL_FONT

wb.save('Business_Sales_Expense_Tracker_v3.xlsx')
print("done")
