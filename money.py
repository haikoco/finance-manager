from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime

dt = datetime.now()

budget = []

wb = Workbook()
ws = wb.active
ws.title = 'many'

filename = 'pas1s.xlsx'
wb.save(filename)