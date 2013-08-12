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

# defining a new function so that we can check if a cell is a number
def getValidatedValue(i):
    if isinstance(i, float): #if the cell is a float 
        return int(i) #give me the integer
    else: # if it's not a float, we might have empty or blank or strings or spaces
        return 0 #give me a 0




##### MAIN PROGRAM CODE

#access workbook that contains the data in the existing directory
workbook = xlrd.open_workbook( './bnidatabefore.xlsx')
#go to sheet name 'Raw Data'
sheet = workbook.sheet_by_name('Raw Data')

#give me indicies of interesting columns
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


#create dictionary with unique member names as keys
report = {}

#loop over entire table and save relevant data
for row_index in range(1, sheet.nrows):
    name = sheet.cell(row_index, memberName_index).value #give me the value of what's in that cell
    # if the name cell is empty skip it because it's possible for there empty rows in the table
    if name == '':
        continue  #otherwise continue with the loop
    a = sheet.cell(row_index, a_index).value
    gi = sheet.cell(row_index, gi_index).value
    gi = getValidatedValue(gi) # with gi, go, v, meet we convert the value of what's in the cell to an integer because
                            #we want to avoid errors that we might get if the cell contains strings, empty, blank, etc and since
                            #python reads numbers as floats from excel, we want integers, we convert it to integers.
    go = sheet.cell(row_index, go_index).value
    go = getValidatedValue(go)
    
    v = sheet.cell(row_index, v_index).value
    v = getValidatedValue(v)
    
    meet = sheet.cell(row_index, meet_index).value
    meet = getValidatedValue(meet)
    
    # decide if a member name is already in dictionary
    if name not in report: 
        #if a member's name is not in the dictionary, add it as a key to the dictionary
        #and create it's corresponding values, which are also keys that will hold values
        report[name] = {'weeks': 0, 'ihaves': 0}
    # if attedance cell is empty, the meeting has not occured
    # we are doing this so we don't have to do date math
    if len(a) > 0:
        # if there is something in the cell, add 1, which means the person has been a member for n weeks
        report[name]['weeks'] += 1
        # add up the performance data, which has already been converted to integers and save it in 'ihaves'
        report[name]['ihaves'] += gi + go + v + meet

# did this to get output in python interpreter to see what output looks like,
# but it's not necessary for this program to run.
for name in sorted(report.keys()):
    ihaves = report[name]['ihaves']
    weeks = report[name]['weeks']
    score = ihaves - weeks
    print name, ihaves, weeks, score

# create new workbook using xlwt
workbook = xlwt.Workbook()
# create new worksheet in the new workbook name the sheet 'CumulativeScore'
sheet = workbook.add_sheet('CumulativeScore')
# call on the easyxf function so that we can format the headers for the data
bold = xlwt.easyxf("font: bold on")
# make the headers and format them to be bold
sheet.write(0,0, 'Full Name', bold)
sheet.write(0,1, 'IHaves', bold)
sheet.write(0,2, '#Weeks', bold)
sheet.write(0,3, 'Score', bold)

# loop on every unique key in the dictionary, which is called to be sorted in alphabetical order
# and write in the new spreadsheet each name, ihaves, weeks, and add the new score variable that
# we created as we are going along the loop iteration 
row = 1 #make a row counter
for name in sorted(report.keys()):
    sheet.write(row, 0, name)
    sheet.write(row, 1, report[name]['ihaves'])
    sheet.write(row, 2, report[name]['weeks'])
    score = report[name]['ihaves'] - report[name]['weeks']
    sheet.write(row, 3, score)
    row = row + 1 # add 1 everytime we are done with each row to advance to the next row

# BOOM! you're done, save the new spreashseet and call it 'Performance'
workbook.save('./Performance.xls')
