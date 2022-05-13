from calendar import weekday

import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt
import datetime

#Establish a connection to a workbook
hours_report = xw.Book(r"C:\Users\15402\Desktop\python_hours_script\hours_report.xlsx")
schedule = xw.Book(r"C:\Users\15402\Desktop\python_hours_script\schedule.xlsx")


#Instantiate a sheet object
june_schedule = schedule.sheets['test_june']
extra_horas = hours_report.sheets['HORA EXTRA']

practice = hours_report.sheets['Sheet2']
lista = hours_report.sheets['LISTA']
"""
#Reading/Writing values to/from ranges is easy
practice.range('A1').value = 'Hello World'
practice.range('A1').value

practice.range('A1').value = [['Foo 1', 'Foo 2', 'Foo 3'], [10.0, 20.0, 30.0]]
practice.range('A1').expand().value

#Using pandas dataframe
df = pd.DataFrame([[1,2],[3,4]], columns=['a','b'])
practice.range('A1').value = df
practice.range('A1').options(pd.DataFrame, expand = 'table').value

#Matplotlib figures can be shown as pictures in Excel
#fig = plt.figure()
#plt.plot([1,2,3,4,5])
#practice.pictures.add(fig,name='Myplot',update = True)

#practice.cells.clear_contents()
"""
"""
for cell in practice.range('B3:B32'):
    if cell != 'N':
        for val in practice.range('C3:C32'):
            val.value=cell.value
"""
#list = [3:32]

#Variables
diurno=lista.range('C1').value
nocturno=lista.range('C2').value
domD=lista.range('C3').value
domN=lista.range('C4').value

columnA = 'A' #Date
columnB = 'B' #holds document number and name as well as shift
columnC = 'C' #document
columnD = 'D' # name
columnE = 'E' # date
columnF = 'F'

row = 3 #Starting row

for i in range(3,33):
    cellA = columnA + str(i)
    cellB = columnB + str(i)
    cellC = columnC + str(i)
    cellD = columnD + str(i)
    cellE = columnE +str(i)
    cellF = columnF + str(i)

    place = columnB+str(i)
    if practice.range(place).value != 'O':
        practice.range(cellC).value = practice.range('B1').value # this value needs to incrument to C1,D1,E1 etc.
        practice.range(cellD).value = practice.range('B2').value # this value needs to increment to C2, D2, E2 etc.
        practice.range(cellE).value = practice.range(cellA).value
        if practice.range(cellA).value.weekday() == 6 and practice.range(place).value == 'N':
            practice.range(cellF).value = domN
        elif practice.range(cellA).value.weekday() == 6 and practice.range(place).value == 'D':
            practice.range(cellF).value = domD
        elif practice.range(cellA).value.weekday() == 6 and practice.range(place).value == 'D':
            practice.range(cellF).value = diurno
        else:
            practice.range(cellF).value = nocturno
    else:
        practice.range(cellC).value = ''
    row = row + 1
