#-*- coding: utf-8 -*-
'''
@author: birdlin
'''
import openpyxl
mAh = 0
Wh = 0

wb = openpyxl.load_workbook('rest+discharge.xlsx', read_only=False)
print (wb.get_sheet_names())

ws = wb.get_sheet_by_name ('sheet1')
print (ws['B2'].value)
print (type(ws['B2'].value))
print (ws.max_row)
print (ws.max_column)
row_index = 1

for row in ws.rows:
    if row_index == 1 :
        ws.cell(row=row_index, column=len(row)+1).value='mV'
        ws.cell(row=row_index, column=len(row)+2).value='mA'
        ws.cell(row=row_index, column=len(row)+3).value='mAh'
        ws.cell(row=row_index, column=len(row)+4).value='Wh'
        row_index += 1
        continue
    Total_Vol = row[0].value * 1000
    Current_mA = row[1].value * 1000 
    mAh += Current_mA * 10 /3600
    Wh += Current_mA/1000 * Total_Vol/1000 * 10 /3600
    ws.cell(row=row_index, column=len(row)+1).value = Total_Vol
    ws.cell(row=row_index, column=len(row)+2).value = Current_mA
    ws.cell(row=row_index, column=len(row)+3).value = mAh
    ws.cell(row=row_index, column=len(row)+4).value = Wh    
    row_index += 1

wb.save(filename = 'new_rest+discharge.xlsx')



