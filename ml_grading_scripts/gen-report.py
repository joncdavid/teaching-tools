#!/usr/bin/env python3
# file: gen-report.py
# date: Nov 16, 2023
# author: Jon David
# description:
#   Given an Excel file (*.xlsx), generate a report in the form of an ORG (*.org) text file.
#

import sys
# import openpyxl
import os

from reportGenerator import *

STR_HELP = 'Requires a command line arg.\npython gen-report.py <name of project to use>.\nName of project can be ' \
           '\"p1\", \"p2\", \"p3\", \"p4\".'

fname = "cs529-Sp24-Details.xlsx"
name_project1_sheet = "Project1"
name_project2_sheet = "Project2"
name_project3_sheet = "Project3"
name_project4_sheet = "Project4"
selected_sheet_name = None  ## to be selected by command line args

## found the bug: column IDs are not sequential after Yaz separated the sheets...
## thanks Yaz for making it confusing <_< ...
rowID_project1_list = [*range(3, 25, 1)]
selected_rowID_list = None  ## to be selected by command line args


wb = openpyxl.load_workbook( fname )
names_all_sheets = wb.sheetnames
# print( names_all_sheets )

if len( sys.argv ) < 2:
    print( STR_HELP )
    sys.exit(1)
    
class_selection = sys.argv[1]
if class_selection == "p1":
    selected_sheet_name = name_project1_sheet
    selected_rowID_list = rowID_project1_list
elif class_selection == "p2":
    selected_sheet_name = name_project2_sheet
    selected_rowID_list = None
elif class_selection == "p3":
    selected_sheet_name = name_project3_sheet
    selected_rowID_list = None
elif class_selection == "p4":
    selected_sheet_name = name_project4_sheet
    selected_rowID_list = None
else:
    print( STR_HELP )
    sys.exit(1)

# sheet = None
# for name in names_all_sheets:
#     if name == selected_sheet_name:
#         sheet = wb[name]
#         break
sheet = wb[selected_sheet_name]

numRows = sheet.max_row
numCols = sheet.max_column

# print( "[debug] numRows: {}".format( numRows ))
# print( "[debug] numCols: {}".format( numCols ))

factory = OrgReportFactory( sheet )
i = 0
for rowID in selected_rowID_list:
    r = factory.genReport( rowID )  ## returns OrgReport
    #r.print_me()
    # print("[debug] rowID: {}, studentName: {}".format( rowID, r.student1Name ))
    rec = r.student1Name.split(',')
    rec2 = []
    rec3 = []
    rec4 = []
    if r.student2Name is not None:
        rec2 = r.student2Name.split(',')
    if r.student3Name is not None:
        rec3 = r.student3Name.split(',')
    if r.student4Name is not None:
        rec4 = r.student4Name.split(',')
    filename = rec[0]
    if rec2:
        filename += '_' + rec2[0]
    if rec3:
        filename += '_' + rec3[0]
    if rec4:
        filename += '_' + rec4[0]
    outf = "output/{}.txt".format( filename ) ## lastname only## "test.org"
    if not os.path.exists('output'):
        os.mkdir('output')
    r.write_to_file( outf )
    print( "[status] ({},{}) wrote file: {}".format( rowID, r.student1Name, outf ))


