#=============================
# Name       :dimension.py
# Argument   :
# Email      :lukyandy3162@gmail.com
# Author     :srhuang
# History    :
#    20191127:Initial
#=============================

#===============
#import section
#===============
import os
import sys
import openpyxl


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
sheet=wb.get_sheet_by_name(target_sheet)

#formula
sheet['C'+str(sheet.max_row+1)]='=SUM(C2:C10)'

#save
wb.save(workbook)
