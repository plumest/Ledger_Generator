import math
import openpyxl
from random import randint, sample

def food(sep, oct):
    expend = {'Food@KU': [80, 100, 120, 130, 150, 200],
             'Food@Major': [260, 275, 279, 289, 299, 319, 320, 322, 349, 367, 389, 399, 450],
             'Food@Central': [279, 289, 299, 319, 320, 322, 349, 367, 389, 399, 450],
             'Food@ThaMall': [279, 289, 299, 319, 320, 322, 349, 367, 389, 399, 450],
             'Food@7-11': [20, 30, 39, 45, 50, 69, 79, 100, 120, 143, 159, 179, 249, 300, 319, 320],
             'Food@Amonphan': [80, 100, 120, 130, 150, 200],
             'Buffe': [199, 209, 269, 279, 289, 299, 349, 399, 499, 599, 699, 799, 899, 999, 1099, 1199, 1299, 1399, 1599],
             'SamSteak': [39, 79, 89, 109, 119, 129, 139, 239, 245]
        }

def travel(sep, oct):
    expend = {'Bus': [8, 12, 13, 15, 18, 21, 25],
             'BTS': [16, 23, 33, 37, 40, 44],
             'MRT': [21, 25, 32, 42],
             'Taxi': [30, 40, 50, 70, 107, 200, 210, 250, 300]
        }

def education(sep, oct):
    expend = {'Books': [89, 90, 100, 115, 120, 150, 189, 190, 213, 225, 240, 245, 255, 265, 279, 319, 345, 350, 450, 550, 777],
             'Courses': [300, 360, 390, 450, 500, 550, 790, 1290, 1590, 3990],
             'Stationary': [5, 7, 10, 39, 55, 69, 79, 89, 99, 109, 120, 129, 139, 249, 349, 769],
        }

def personal(sep, oct):
    gender = ''
    while gender.lower() not in ['m', 'f']:
        gender = input('Enter your gender (M)ale/(F)emale: ')

    if gender == 'f':
        expend = {'Haircut': [300, 500, 1000, 1500, 2000, 3000, 4000],
                'Cosmetics': [390, 550, 569, 750, 790, 1200, 1235, 1790, 2000, 2190, 2490, 2550, 2590, 3190, 3550, 4790, 5390],
                'Manicure': [200, 250, 300, 450, 500],
        }

    else:
        expend = {'Haircut': [150, 200, 300, 500, 1000, 1500, 2000, 3000, 4000],
                'Cosmetics': [390, 550, 569, 750, 790, 1200, 1235, 1790, 2000, 2190, 2490, 2550, 2590, 3190, 3550, 4790, 5390],
                'Manicure': [200, 250, 300, 450, 500],
        }

def enterainment(sep, oct):
    expend = {'Movie': [100, 110, 120, 240, 400],
             'Concert': [390, 490, 590, 790, 1000, 1200, 1500, 2000, 2900, 3000, 3500, 3900, 4500, 5500, 6500],
             'Games': [129, 150, 199, 209, 300, 329, 399, 419, 450, 500, 790, 1000, 1200, 1290, 2000, 2590, 2990]
        }

def other(sheet, index):
    expend = {'Boots': [390, 790, 890, 1290, 1390, 1550, 1990, 2990, 3990, 4500, 4590, 4900, 5900, 6000, 6100, 7900, 8900, 9750],
             'Clothes': [69, 100, 120, 139, 159, 259, 300, 350, 400, 500, 790, 890, 1000, 1200, 1300, 1450, 1590, 1990, 2900],
             'Accessories': [39, 59, 79, 99, 390, 590, 790, 890, 1290, 1390, 2900, 3900, 4500, 4900, 5500, 6900]
        }
    remaining = int(sheet[f'Q{index}'].internal_value)
    price = math.inf
    count = 0

    while price > remaining:
        pick_expend = sample(expend.items(), 1)[0]
        item = pick_expend[0]
        price = sample(pick_expend[1], 1)[0]
        count += 1
        if count == 4:
            item = ''
            price = 0

    sheet[f'N{index}'] = item
    sheet[f'O{index}'] = price


def monthly(sep, oct):
    print('~~ Fill an interger!!!')
    rent = int(input('Enter your Rent cost: '))
    sep['P39'] = rent
    oct['P39'] = rent
    net = int(input('Enter your internet cost: '))
    sep['P40'] = net
    oct['P40'] = net
    utility = int(input('Enter your utility cost (water/electricity bills): '))
    sep['P41'] = randint(utility-100, utility+100)
    oct['P41'] = randint(utility-100, utility+100)

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

for i in range(5, 35):
    other(ws, i)

# monthly(ws, wo)

# Save file
wb.save(file_path)