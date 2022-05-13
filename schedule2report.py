from calendar import weekday

import pandas as pd
import xlwings as xw
import matplotlib.pyplot as plt
import datetime
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string

#Establish a connection to a workbook
hours_report = xw.Book(r"C:\Users\15402\Desktop\python_hours_script\hours_report.xlsx")
schedule = xw.Book(r"C:\Users\15402\Desktop\python_hours_script\miner_schedule.xlsx")
wb = openpyxl.load_workbook(r"C:\Users\15402\Desktop\python_hours_script\miner_schedule.xlsx")

#Instantiate a sheet object
june_schedule = schedule.sheets['MUZO EMPLOYEES JUNE']
extra_horas = schedule.sheets['HORA EXTRA']
sheet = wb['MUZO EMPLOYEES JUNE']
lista = schedule.sheets['LISTA']

#Variables for novedad o tipo de hora 'G'
diurno=lista.range('C1').value
nocturno=lista.range('C2').value
domD=lista.range('C3').value
domN=lista.range('C4').value

#variables for turno normal del trabajador el el dia 'H'
day_turno_hrs=lista.range('B1').value
night_turno_hrs=lista.range('B2').value

#variables for turno en que se efectua la novedad 'I'
A_turno=lista.range('D1').value
C_turno=lista.range('D3').value
A_turno_domingo=lista.range('D4').value
C_turno_domingo=lista.range('D6').value

#variables for hora en que se efectua la novedad 'J'
day_start_hr=lista.range('B9').value
night_start_hr=lista.range('B10').value

#variables for actividad realiza and numero de horas
actividad = 'RAMPA JD'
num_hours = 2
#

columnD = 'D' #names column
columnC = 'C' #cedula column

count = 0
bigdict = {}
for i in range(8,sheet.max_row+1): #looping down names
    name = columnD + str(i)
    cedula = columnC + str(i)

    name = june_schedule.range(name).value
    cedula = june_schedule.range(cedula).value


    for x in range(column_index_from_string('E'), column_index_from_string('AH')+1, 1): #looping through shifts
        #print(sheet.cell(row=i, column=x).value)
        shift = sheet.cell(row=i, column=x).value
        date = sheet.cell(row=7, column=x).value
        if shift != 'O' and name != 'NEW PERSON': #the None clause gets rid of new miner
            dicts={'name':name,'cedula':cedula,'date':date,'shift':shift}
            count+=1

            #keys = ['name','cedula','date','shift']
            #values = [name,cedula,date,shift]
            #dicts[keys['name']] = values[name]
            #print(dicts,count)
            bigdict[count]=dicts
            #print(count)

            #print(name,cedula,date,shift)
#print(bigdict)
#print("LENGTH OF BIGDICT",str(len(bigdict)))

columnC = 'C' #documento,----- cedula
columnD = 'D' #nombre apellido,----- name
columnG = 'G' # novedad/tipo de hora----- diurno,nocturno,domD,domN
columnH = 'H' #turno normal del trabajador el el dia ------ day_turno_hrs,night_turno_hrs
columnI = 'I' #turno en que se efectua la novedad------ A_turno, C_turno, A_turno_domingo, C_turno_domingo
columnJ = 'J' #hora en que se efectua la novedad ----------- day_start_hr, night_start_hr
columnK = 'K' #actividad realizada------actividad
columnL = 'L'
columnM = 'M' #numero de horas ---------  num_horas

row = 5 #starting row



for entry in range(len(bigdict)):
    cellC = columnC + str(entry+row)
    cellD = columnD + str(entry+row)
    cellG = columnG + str(entry+row)
    cellH = columnH + str(entry+row)
    cellI = columnI + str(entry+row)
    cellJ = columnJ + str(entry+row)
    cellK = columnK + str(entry+row)
    cellL = columnL + str(entry + row)
    cellM = columnM + str(entry+row)

    extra_horas.range(cellC).value = bigdict[entry+1]['cedula']
    extra_horas.range(cellD).value = bigdict[entry+1]['name']
    extra_horas.range(cellL).value = bigdict[entry + 1]['date']

    if bigdict[entry+1]['date'].weekday() == 6 and bigdict[entry+1]['shift'] == 'N':
        extra_horas.range(cellG).value = domN
        extra_horas.range(cellH).value = night_turno_hrs
        extra_horas.range(cellI).value = C_turno_domingo
        extra_horas.range(cellJ).value = night_start_hr
        extra_horas.range(cellK).value = actividad
        extra_horas.range(cellM).value = num_hours
    elif bigdict[entry+1]['date'].weekday() == 6 and bigdict[entry+1]['shift'] == 'D':
        extra_horas.range(cellG).value = domD
        extra_horas.range(cellH).value = day_turno_hrs
        extra_horas.range(cellI).value = A_turno_domingo
        extra_horas.range(cellJ).value = day_start_hr
        extra_horas.range(cellK).value = actividad
        extra_horas.range(cellM).value = num_hours
    elif bigdict[entry+1]['date'].weekday() != 6 and bigdict[entry+1]['shift'] == 'N':
        extra_horas.range(cellG).value = nocturno
        extra_horas.range(cellH).value = night_turno_hrs
        extra_horas.range(cellI).value = C_turno
        extra_horas.range(cellJ).value = night_start_hr
        extra_horas.range(cellK).value = actividad
        extra_horas.range(cellM).value = num_hours
    else:
        extra_horas.range(cellG).value = diurno
        extra_horas.range(cellH).value = day_turno_hrs
        extra_horas.range(cellI).value = A_turno
        extra_horas.range(cellJ).value = day_start_hr
        extra_horas.range(cellK).value = actividad
        extra_horas.range(cellM).value = num_hours

print("complete")









