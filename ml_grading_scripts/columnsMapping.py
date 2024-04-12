#!/usr/bin/env python3
# file: columnsMapping.py
# date: Nov 16, 2023
# author: Jon David
# description:
#   define mapping between items and column index
#--------------------------------------------------------------------
# notes:
#   Excel indices start with 1
#   our input source has no column A, columns begin at B (index 2)
# 

## columnStartIndex = 5  ## column E

colMap = {}

colMap[ "student1Name" ] = 2
colMap[ "student2Name" ] = 3
colMap[ "student3Name" ] = 4
colMap[ "student4Name" ] = 5

## Overall Project Grade
colMap[ "p_grade" ] = 8 ## column H

## Extra Points
colMap["extra"] = 9

## Deliverables
colMap[ "pdf" ] = 10 ## column J
colMap[ "pdf_comments" ] = 11
colMap[ "pkg" ] = 12 ## column L
colMap[ "pkg_comments" ] = 13
colMap[ "manifest" ] = 14  ## column N
colMap[ "manifest_comments" ] = 15
colMap[ "team_contributions" ] = 16  ## column P
colMap[ "team_contributions_comments" ] = 17
colMap[ "kaggle" ] = 18  ## column R
colMap[ "kaggle_comments" ] = 19

## Code
colMap[ "clean" ] = 21  ## column U
colMap[ "clean_comments" ] = 22
colMap[ "ig_impl" ] = 23  ## column W
colMap[ "ig_impl_comments" ] = 24
colMap[ "tree_impl" ] = 25  ## column Y
colMap[ "tree_impl_comments" ] = 26
colMap[ "forest_impl" ] = 27  ## column AA
colMap[ "forest_impl_comments" ] = 28

#Analysis
colMap[ "p1" ] = 30  ## column AD
colMap[ "p1_comment" ] = 31
colMap[ "p2" ] = 32  ## column AF
colMap[ "p2_comment" ] = 33
colMap[ "p3" ] = 34  ## column AH
colMap[ "p3_comment" ] = 35
colMap[ "p4" ] = 36  ## column AJ
colMap[ "p4_comment" ] = 37
colMap[ "p5" ] = 38  ## column AL
colMap[ "p5_comment" ] = 39
colMap[ "p6" ] = 40  ## column AN
colMap[ "p6_comment" ] = 41
colMap[ "p7" ] = 42  ## column AP
colMap[ "p7_comment" ] = 43
