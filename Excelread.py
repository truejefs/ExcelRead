# -*- coding: utf-8 -*-
"""
Created on Fri Feb  8 20:11:31 2019

@author: truejefs
Copyright (c) 2019 truejefs
"""

from xlrd import open_workbook, xldate_as_tuple
from os.path import splitext

START_ROW=1
Employee_column=0


file_path="C:/Python_Scripts/ExcellRead/ExcellRead.xlsx"

with open_workbook(file_path) as rb: #Using a with clause to ensure the workbook is closed in case the code crashes.
    r_sheet = rb.sheet_by_index(0) # read only copy to introspect the file
    
    employee="Jack"
    
    for row_index in range(START_ROW, r_sheet.nrows): #searching for the employee
            if r_sheet.cell_value(row_index,Employee_column) == employee:
                for col_index in range(Employee_column+1, r_sheet.ncols): #printing data
                    day=xldate_as_tuple(r_sheet.cell_value(0,col_index), rb.datemode)
                    day=str(day[2]) + '/' + str(day[1]) + '/' + str(day[0])
                    print(day + ": " + r_sheet.cell_value(row_index,col_index))
                continue
            elif row_index == r_sheet.nrows:
                print('Employee not found')
