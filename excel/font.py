#=============================
# Name       :font.py
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
from openpyxl.styles import Font


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

#get Font object
Font1 = Font(name='Calibri', size=24, italic=False)
sheet['A1'].font=Font1

#save
wb.save(workbook)




