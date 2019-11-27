#=============================
# Name       :workbook_sheet_cell.py
# Argument   :
# Email      :lukyandy3162@gmail.com
# Author     :srhuang
# History    :
#    20191126:Initial
#=============================

#===============
#import section
#===============
import os
import sys
import openpyxl
#import pandas as pd


#================
#variable section
#================
workbook='example.xlsx'
target_sheet='sheet1'

#=================
#argument section
#=================

#=================
#function section
#=================

#===============
#progress start
#===============

#open work book
wb=openpyxl.load_workbook(workbook)

#get sheets
all_sheet=wb.get_sheet_names()
for sheet in all_sheet:
    if sheet==target_sheet:
        sheet_name=sheet
    print sheet

sheet=wb.get_sheet_by_name(sheet_name)
print sheet['A1'].value

#get cell
print sheet.cell(row=1, column=2).value

#1~11, add 2 for each step
for i in range(2,13,2):
    print(i, sheet.cell(row=i, column=2).value)

#get max column and max row
print "max row :"+str(sheet.max_row)
print "max column :"+str(sheet.max_column)

#get area
for rowOfCellObject in sheet['A2':'C9']:
    #print each row
    for cellObject in rowOfCellObject:
        print cellObject.coordinate, cellObject.value 
    print "-----END OF ROW-----"

#get specific row
for cellObject in sheet['B']:
    print cellObject.value

#new and save excel file
all_sheet=wb.get_sheet_names()
sheet2=wb.get_sheet_by_name(all_sheet[1])
sheet2.title='new'
wb.save(workbook)

#new and delete sheet
wb.create_sheet(index=3, title='new')
wb.save(workbook)

all_sheet=wb.get_sheet_names()
wb.remove_sheet(wb.get_sheet_by_name(all_sheet[2]))
print wb.get_sheet_names()
wb.save(workbook)

