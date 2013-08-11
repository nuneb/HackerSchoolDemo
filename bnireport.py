#!/usr/bin/env python

import xlrd
import xlwt

from xlrd import cellname, XL_CELL_TEXT

workbook = xlrd.open_workbook( '/cygdrive/c//Users/nboyadjian/Desktop/Google Drive/Fabulous Statistics.xlsx')
sheet = workbook.sheet_by_name('Raw Data')
#print sheet.name#print sheet.nrows#print sheet.ncols

#create unique names
#create dictionary with unique names as keys

#give me indecies of interesting columns
memberName_index = 0
gi_index = 0
go_index = 0
v_index = 0
meet_index = 0
a_index = 0
for col_index in range(sheet.ncols):
    if sheet.cell(0, col_index).value == 'Full Name':
        memberName_index = col_index
    if sheet.cell(0, col_index).value == 'GI':
        gi_index = col_index
    if sheet.cell(0, col_index).value == 'GO':
        go_index = col_index
    if sheet.cell(0, col_index).value == 'Visitors':
        v_index = col_index
    if sheet.cell(0, col_index).value == '1-to-1':
        meet_index = col_index
    if sheet.cell(0, col_index).value == 'Attendance':
        a_index = col_index

# defining a new function so that we can check if a cell is a number
def getValidatedValue(i):
    if isinstance(i, float): #if the cell is a float 
        return int(i) #give me the integer
    else: # if it's not a float, we might have empty or blank or strings or spaces
        return 0 #give me a 0

report = {} #create dictionary for storing member data

# loop over entire table and save relevant data
for row_index in range(1, sheet.nrows):
    name = sheet.cell(row_index, memberName_index).value
    
    print name
    a = sheet.cell(row_index, a_index).value
    gi = sheet.cell(row_index, gi_index).value
    gi = getValidatedValue(gi)
    
    go = sheet.cell(row_index, go_index).value
    go = getValidatedValue(go)
    
    v = sheet.cell(row_index, v_index).value
    v = getValidatedValue(v)
    
    meet = sheet.cell(row_index, meet_index).value
    meet = getValidatedValue(meet)
    
    # decide if a member name is already in dictionary
    if name not in report: 
        #if a member's name is not in the dictionary, add it as a key to the dictionary
        report[name] = {'weeks': 0, 'ihaves': 0}
    # if attedance cell is empty meeting has not occured
    # doing this so we don't have to do date math
    if len(a) > 0:
        report[name]['weeks'] += 1
        report[name]['ihaves'] += gi + go + v + meet

#for name in sorted(report.keys()):
#    score = (report[name]['ihaves'] - report[name]['weeks'])
#    print name, report[name]['ihaves'], report[name]['weeks'], score

#data = [sheet.cell_value(0, col) for col in range(sheet.ncols)]

#workbook = xlwt.Workbook()
#sheet = workbook.add_sheet('Baseline')

#for index, value in enumerate(data):
#    sheet.write(0, index, value)

#workbook.save('/cygdrive/c//Users/nboyadjian/Desktop/ExcelPython/BaselinePerformance.xls')
