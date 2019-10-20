import openpyxl
from formula import *
from randdata import *

customer_name = input("Enter your name (without any special character or space): ")
file_path = f"customer/{customer_name}.xlsx"

wb = openpyxl.load_workbook('prototype.xlsx', data_only=True)

# ws for September sheet
ws = wb.active
wb.active = 2

# wo for October sheet
wo = wb.active

# Salary
salary = int(input('Enter your salary: '))
ws['C5'] = salary
wo['C5'] = salary

# gender
gender = ''
while (gender not in ['m', 'f']):
    gender = input('Enter your gender (M)ale/(F)emale: ').lower()

# for i in range(5, 35):
#     food(ws, i)
#     travel(ws, i)
#     personal(ws, i, gender)
#     enterainment(ws, i)
#     other(ws, i)

# monthly(ws, wo)

def write(sheet, month='sep'):
    if month == 'sep':
        stop = 35
    else:
        stop = 36
    
    for i in range(5, stop):
        food(sheet, i)

write(ws)
# Save file
wb.save(file_path)