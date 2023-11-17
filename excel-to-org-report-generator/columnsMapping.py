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

colMap[ "studentName" ] = 2

## Quesition 1
colMap[ "q1grade" ] = 5 ## column E
colMap[ "q1comment" ] = 6  

## Question 2
colMap[ "q2agrade" ] = 7  ## column G
colMap[ "q2acomment" ] = 8
colMap[ "q2bgrade" ] = 9  ## column I
colMap[ "q2bcomment" ] = 10
colMap[ "q2cgrade" ] = 11  ## column K
colMap[ "q2ccomment" ] = 12
colMap[ "q2dgrade" ] = 13  ## column M
colMap[ "q2dcomment" ] = 14
colMap[ "q2egrade" ] = 15  ## column O
colMap[ "q2ecomment" ] = 16
colMap[ "q2fgrade" ] = 17  ## column Q
colMap[ "q2fcomment" ] = 18
colMap[ "q2fbonusgrade" ] = 19  ## column S

colMap[ "q3agrade" ] = 20  ## column T
colMap[ "q3acomment" ] = 21
colMap[ "q3bgrade" ] = 22  ## column V
colMap[ "q3bcomment" ] = 23

colMap[ "q4agrade" ] = 24  ## column X
colMap[ "q4acomment" ] = 25
colMap[ "q4bgrade" ] = 26  ## column Z
colMap[ "q4bcomment" ] = 27

colMap[ "q5DFSgrade" ] = 28  ## column AB
colMap[ "q5DFScomment" ] = 29
colMap[ "q5BFSgrade" ] = 30  ## column AD
colMap[ "q5BFScomment" ] = 31
colMap[ "q5HILLgrade" ] = 32  ## column AF
colMap[ "q5HILLcomment" ] = 33
colMap[ "q5BESTgrade" ] = 34  ## column AG
colMap[ "q5BESTcomment" ] = 35
colMap[ "q5ASTARgrade" ] = 36  ## column AJ
colMap[ "q5ASTARcomment" ] = 37
colMap[ "q5fgrade" ] = 38  ## column AL
colMap[ "q5fcomment" ] = 39

colMap[ "q6agrade" ] = 40  ## column AN
colMap[ "q6acomment" ] = 41
colMap[ "q6bgrade" ] = 42  ## column AP
colMap[ "q6bcomment" ] = 43
colMap[ "q6cgrade" ] = 44  ## column AR
colMap[ "q6ccomment" ] = 45
colMap[ "q6dgrade" ] = 46  ## column AT
colMap[ "q6dcomment" ] = 47 
colMap[ "q6egrade" ] = 48  ## column AV
colMap[ "q6ecomment" ] = 49
colMap[ "q6fgrade" ] = 50  ## column AX
colMap[ "q6fcomment" ] = 51

colMap[ "q7bonusgrade" ] = 52  ## column AAB
colMap[ "q7bonuscomment" ] = 53
