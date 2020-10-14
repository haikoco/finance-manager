from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime

dt = datetime.now()

passw = []


wb = load_workbook(filename='pas1s.xlsx')
sheet = wb.active
sheet.title = 'data'
q = len(sheet['A'])
print(q)
row = q
sheet['A'+str(row)] = '#'
sheet['B'+str(row)] = 'site'
sheet['C'+str(row)] = 'login'
sheet['D'+str(row)] = 'pass'

i=q-1
while True:
	i += 1
	site = input('Название: ')
	login = input('Логин: ')
	passwo = input('Пароль: ')
	if site==str(1):
		break
	else:
		passw.append([i, site, login, passwo])
		print(passw)
for item in passw:
			 row +=1
			 sheet['A'+str(row)] = item[0]
			 sheet['B'+str(row)] = item[1]
			 sheet['C'+str(row)] = item[2]
			 sheet['D'+str(row)] = item[3]		

    
filename = 'pas1s.xlsx'
wb.save(filename)