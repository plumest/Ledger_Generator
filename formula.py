def expense(sheet):
    for i in range(5, 36):
        sheet[f'P{i}'] = f'=SUM(E{i},G{i},I{i},K{i},M{i},O{i})'

    sheet['P45'] = '=SUM(P36:P44)'
    sheet['P46'] = '=C36-P45'

def remaining(sheet):
    sheet['Q5'] = '=C5-P5'
    for i in range(6, 36):
        sheet[f'Q{i}'] = f'=Q{i-1}+C{i}-P{i}'
    sheet['Q36'] = '=C36-P36'

def total(sheet):
    sheet['C36'] = '=SUM(C5:C35)'
    sheet['E36'] = '=SUM(E5:E35)'
    sheet['G36'] = '=SUM(G5:G35)'
    sheet['I36'] = '=SUM(I5:I35)'
    sheet['K36'] = '=SUM(K5:K35)'
    sheet['M36'] = '=SUM(M5:M35)'
    sheet['O36'] = '=SUM(O5:O35)'
    sheet['P36'] = '=SUM(P5:P35)'

def oct_summary(sheet):
    sheet[f'C51'] = '=C36'
    sheet[f'C52'] = '=SUM(C54:C62)'
    sheet[f'C53'] = '=C51-C52'
    sheet[f'C54'] = '=P39'
    sheet[f'C55'] = '=P40'
    sheet[f'C56'] = '=P41'
    sheet[f'C57'] = '=E36'
    sheet[f'C58'] = '=G36'
    sheet[f'C59'] = '=I36'
    sheet[f'C60'] = '=K36'
    sheet[f'C61'] = '=M36'
    sheet[f'C62'] = '=O36'

def sep_summary(sheet):
    oct_summary(sheet)
    for i in range(51, 63):
        sheet[f'D{i}'] = f'=October!C{i}'