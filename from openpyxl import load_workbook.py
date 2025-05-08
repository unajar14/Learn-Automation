from openpyxl import load_workbook
from openpyxl.styles import Font

wb = load_workbook('report.xlsx')

sheet = wb['Report']

sheet['A1'] = 'Sales Report'
sheet['A2'] = 'January'
sheet['A1'].font = Font('Arial', bold=True, size=20)
sheet['A2'].font = Font('Calibri', bold=True, size=14)

wb.save('report_jan.xlsx')
