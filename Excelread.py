# -*- coding: utf-8 -*-
"""
Created on Fri Feb  8 20:11:31 2019

@author: truejefs
Copyright (c) 2019 truejefs
"""

from xlrd import open_workbook, xldate_as_tuple
from os.path import splitext

START_ROW=1 # Assume row 0 (1 in excel) contains the header, so data starts ar row 1 (2 in excel)
EmployeeID_column=0 # This is the column in which the employee IDs are listed in the excel spreadsheet.
employeeID="Jack" # ID of the employee you wan to search for
file_path="C:/Python_Scripts/ExcellRead/ExcellRead.xlsx" # Path to the excel file

with open_workbook(file_path) as workbook_obj: # Set workbook_obj as the workbook object by opening the workbook (the with clause ensures the workbook will be closed when the script terminates)
    worksheet_obj = workbook_obj.sheet_by_index(0) # Set r-sheet as the first sheet object within the workbook_obj workbook.
    
    for row_index in range(START_ROW, worksheet_obj.nrows): # Scanning through all filled rows of the sheet looking for the employeeID
            if worksheet_obj.cell_value(row_index,EmployeeID_column) == employeeID: # If the value of the cell at (row_index, EmployeeID_column) matches the employeeID being searched 
                for col_index in range(EmployeeID_column+1, worksheet_obj.ncols): # the row matching the EmployeeID has been found: now scanning through columns in this row to print the rota data
                    day=xldate_as_tuple(worksheet_obj.cell_value(0,col_index), workbook_obj.datemode) # read each date (top row) and record as a tuple in the format: (y,m,d,h,m,s). This is because Excel and PYTHON dates are not compatible
                    day=str(day[2]) + '/' + str(day[1]) + '/' + str(day[0]) # Rearranging the three first elements in the day tuple as a string with the format: d/m/y
                    print(day + ": " + worksheet_obj.cell_value(row_index,col_index)) # Printing the data to the screen
                break # This line makes sure the script stops searching for the employeeID once it has already been found and printed.
            elif row_index == worksheet_obj.nrows: # after scanning all rows in the excel sheet and the employeeID match is not found, print 'Employee not found'
                print('Employee not found')
