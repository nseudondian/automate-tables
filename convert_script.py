from openpyxl import load_workbook
from openpyxl.styles import Font

month = 'february'
wb = load_workbook('report.xlsx')
sheet = wb['Report']

sheet['A1'] = 'Sales Report'
sheet["A2"] = month

sheet['A1'].font = Font('Calibri', bold=True, size=20 )
sheet['A2'].font = Font('Calibri', bold=True, size=10 )

wb.save(f'report_{month}.xlsx')