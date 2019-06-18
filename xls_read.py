# -*- coding: utf-8 -*-
"""
Created on Wed Jun 12 12:31:24 2019

@author: Andrej
"""

import openpyxl

print('Skripta za samodejni vnos službe v Google koledar\n')

wb = openpyxl.load_workbook('urnik_RIV_2019.xlsx')

sheet = wb.active

def user(): #Najde ime v stoplcih iz vrne vrstico
    name = input('Vnesi svoje ime:  ')
    for index, row in enumerate(sheet.iter_rows()):
        for cell in row:
            if cell.value == name.upper():
                return (index +1)
                print(index+1)
    pass
row = user()
start_cell = 'E{}'.format(row)
end_cell = 'AH{}'.format(row)

def read_shift():   #Preberi izmene v seznam
    cell = sheet[start_cell:end_cell]
    my_list = []
    for c1 in cell:
        for c_val in c1:
            my_list.append(c_val.value)
    return my_list
    pass
shift_list = read_shift()

def dopust(): #Sivo obarvane celice so prosti dnevi
    cell_dod = sheet['E{}'.format(row+1):'AH{}'.format(row+1)]
    fraj_list = []
    for c2 in cell_dod:
        for fraj in c2:
            fraj_list.append(fraj.value)
    return fraj_list
    pass
fraj_list = dopust()


def working_days(): #Vnese dopust v seznam s šihti
    oznake_dopust = ['LD','KU','PO','IZPIT','PP','B']
    for day,shift in enumerate(fraj_list):
        if shift in oznake_dopust:
            shift_list.pop(day)
            shift_list.insert(day,'None')
    pass
working_days()
def read_days(): #Preberi vrstni red dni v seznam
    days_list = []
    cell = sheet['E2':'AH2']
    for c1 in cell:
        for c_val in c1:
            days_list.append(c_val.value)
    return days_list
    pass
days_list = read_days()

def month(): #Preberi mesec
    cell = sheet['A1'].value
    months = {'januar':'01','februar':'02', 'marec':'03', 'april':'04', 'maj':'05', 'junij':'06', 'julij':'07', 'avgust':'08', 'september':'09', 'oktober':'10', 'november':'11', 'december':'12'}
    for month,date in months.items():
        if month == cell.lower():
            return date
    pass

def vodja_odkrivanje(): #Preberi barve fonta za odkrivanje in vodje v seznam
    vodja_list = []
    cell = sheet[start_cell:end_cell]
    for c1 in cell:
        for c_font_color in c1:
            if c_font_color.font.color is not None:
                vodja_list.append(c_font_color.font.color.rgb)
            else:
                vodja_list.append('Fraj')
    return vodja_list
    pass

def duration(day): #Kombinacija treh seznamov za določanje začetka in konca službe (h1,m1,h2,m2)
    if shift_list[day]=='P':
        if days_list[day] == 'PET' or days_list[day] == 'SOB':
            return(15,00,21,00)
        elif vodja_odkrivanje[day] == 'FFFF0000': #Rdeča - vodja
            return(15,00,23,00)
        else:
            return(15,00,20,00)
    elif shift_list[day] =='D':
        if vodja_odkrivanje[day] == 'FF00B0F0': #Modra - odkrivanje
            return(8,00,15,00)
        elif vodja_odkrivanje[day] == 'FFFF0000': #Rdeča - vodja
            return(8,30,15,00)
        else:
            return(9,00,15,00)
    elif shift_list[day] =='8-13':
        return(8,00,13,00)
    elif shift_list[day] =='8-17' or shift_list[day]=='9-17':
        return(8,00,17,00)
    elif shift_list[day] =='13-21':
        return(13,00,21,00)
    elif shift_list[day] =='V': #Rdeča - vodja
        return(13,00,21,00)
    elif shift_list[day] =='C':
        return(9,00,20,00)
    pass

def koncni_seznam():
    koncni_seznam = ['Nicti_dan']
    for i in range(0, len(days_list)):
        koncni_seznam.append(duration(i))
    return koncni_seznam
    pass

shift_list = read_shift()
working_days()
vodja_odkrivanje = vodja_odkrivanje()
days_list = read_days()
fraj_list = dopust()
month = month()
seznam = koncni_seznam()