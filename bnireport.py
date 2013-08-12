#!/usr/bin/env python

########################################################################################
#
# bnireport.py
#
# (Explanation is also in README)
# This program takes data collected on a weekly basis for performance measurement
# by person name, by week they were a member, and gives us a scorecard for that person.
# A negative score means they are behind and a positive score means they are ahead
# by that amount.
#
########################################################################################

#install modules xlrd and xlwt from http://www.python-excel.org/
#import modules for use
import xlrd
import xlwt
from xlrd import cellname, XL_CELL_TEXT
from xlwt import *

##### FUNCTION DEFINITIONS

# Defining a new function so that we can check if a cell is a number
def getValidatedValue(i):
    #if the cell is a float 
    if isinstance(i, float):
        #give me the integer
        return int(i)
    # if it's not a float, we might have empty or blank or strings or spaces
    else:
        #give me a 0
        return 0 



##### MAIN PROGRAM CODE

#access workbook that contains the data in the existing directory
workbook = xlrd.open_workbook( './bnidatabefore.xlsx')
#go to sheet name 'Raw Data'
sheet = workbook.sheet_by_name('Raw Data')

# Loop over entire table and give me indicies of interesting columns
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


# Create dictionary with unique member names as keys
report = {}

# Loop over entire table and save relevant data
for row_index in range(1, sheet.nrows):
    # Give me the value of what's in that cell
    name = sheet.cell(row_index, memberName_index).value
    # If the name cell is empty skip it because it's possible for there to be
    # empty rows in the table
    if name == '':
        # Otherwise continue with the loop
        continue
    a = sheet.cell(row_index, a_index).value
    gi = sheet.cell(row_index, gi_index).value
    # With gi, go, v, meet we convert the value of what's in the cell to an
    # integer because we want to avoid errors that we might get if the cell
    # contains strings, empty, blank, etc. Since Excel believes that the data
    # is numeric, it is stored as a float and so goes to python as a float.
    # We want integers, so we convert it to integers using the function.
    gi = getValidatedValue(gi)
    go = sheet.cell(row_index, go_index).value
    go = getValidatedValue(go)
    
    v = sheet.cell(row_index, v_index).value
    v = getValidatedValue(v)
    
    meet = sheet.cell(row_index, meet_index).value
    meet = getValidatedValue(meet)
    
    # Decide if a member name is already in dictionary
    if name not in report: 
        # If a member's name is not in the dictionary, add it as a key to the
        # dictionary and create it's corresponding values, which are also keys
        # that will hold values.
        report[name] = {'weeks': 0, 'ihaves': 0}
    # If attedance cell is empty, the meeting has not occured
    # We are doing this so we don't have to do date math
    if len(a) > 0:
        # If there is something in the cell, add 1, which means the person has
        # been a member for n weeks
        report[name]['weeks'] += 1
        # Add up the performance data, which has already been converted to
        # integers and save it in 'ihaves'
        report[name]['ihaves'] += gi + go + v + meet

# Did this to get output in python interpreter to see what output looks like,
# but it's not necessary for this program to run.
for name in sorted(report.keys()):
    ihaves = report[name]['ihaves']
    weeks = report[name]['weeks']
    score = ihaves - weeks
    print name, ihaves, weeks, score

# Create new workbook using xlwt
workbook = xlwt.Workbook()
# Create new worksheet in the new workbook name the sheet 'CumulativeScore'
sheet = workbook.add_sheet('CumulativeScore')
# Call on the easyxf function so that we can format the headers
bold = xlwt.easyxf("font: bold on")
# Make the headers and format them to be bold
sheet.write(0,0, 'Full Name', bold)
sheet.write(0,1, 'IHaves', bold)
sheet.write(0,2, '#Weeks', bold)
sheet.write(0,3, 'Score', bold)

# Loop on every unique key in the dictionary, which is called to be sorted in
# alphabetical order and write in the new spreadsheet each name, ihaves, weeks,
# and add the new score variable that we created as we are going along the loop iteration
# Make a row counter
row = 1 
for name in sorted(report.keys()):
    sheet.write(row, 0, name)
    sheet.write(row, 1, report[name]['ihaves'])
    sheet.write(row, 2, report[name]['weeks'])
    score = report[name]['ihaves'] - report[name]['weeks']
    sheet.write(row, 3, score)
    # Add 1 everytime we are done with each row to advance to the next row
    row = row + 1 

# BOOM! you're done, save the new file in the same directory and call it 'Performance'
workbook.save('./Performance.xls')