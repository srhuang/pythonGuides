#=============================
# Name       :graph.py
# Argument   :
# Email      :lukyandy3162@gmail.com
# Author     :srhuang
# History    :
#    20191128:Initial
#=============================

#===============
#import section
#===============
import os
import sys
import openpyxl
from openpyxl.chart import BarChart,Reference

#================
#variable section
#================
workbook='example'
input='input/'+workbook+'.xlsx'
output='output/'+workbook+'.xlsx'
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
wb=openpyxl.load_workbook(input)
sheet=wb.get_sheet_by_name(target_sheet)

#graph
refObject=Reference(sheet, min_col=3, min_row=2, max_col=3, max_row=12)
chart=BarChart()
chart.add_data(refObject)

chart.title="BAR-CHART"
chart.x_axis.title="X-axis"
chart.y_axis.title="Y-axis"

sheet.add_chart(chart, "E3")

#save
wb.save(output)




