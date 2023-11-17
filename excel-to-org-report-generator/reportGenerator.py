#!/usr/bin/env python3
# filename: report-generator.py
# date: Nov 16, 2023
# author: Jon David
# description:
#   class definition that writes the report
#

import openpyxl

from columnsMapping import *


class OrgReport(object):
    def __init__(self):
        self.studentName = None
        
        self.q1grade = None
        self.q1comment = None
        
        self.q2agrade = None
        self.q2acomment = None
        self.q2bgrade = None
        self.q2bcomment = None
        self.q2cgrade = None
        self.q2ccomment = None
        self.q2dgrade = None
        self.q2dcomment = None
        self.q2egrade = None
        self.q2ecomment = None
        self.q2fgrade = None
        self.q2fcomment = None
        self.q2fbonusgrade = None
        self.q2fbonuscomment = None

        self.q3agrade = None
        self.q3acomment = None
        self.q3bgrade = None
        self.q3bcomment = None

        self.q4agrade = None
        self.q4acomment = None
        self.q4bgrade = None
        self.q4bcomment = None

        self.q5BFSgrade = None
        self.q5BFScomment = None
        self.q5DFSgrade = None
        self.q5DFScomment = None
        self.q5HILLgrade = None
        self.q5HILLcomment = None
        self.q5BESTgrade = None
        self.q5BESTcomment = None
        self.q5ASTARgrade = None
        self.q5ASTARcomment = None
        self.q5fgrade = None
        self.q5fcomment = None

        self.q6agrade = None
        self.q6acomment = None
        self.q6bgrade = None
        self.q6bcomment = None
        self.q6cgrade = None
        self.q6ccomment = None
        self.q6dgrade = None
        self.q6dcomment = None
        self.q6egrade = None
        self.q6ecomment = None
        self.q6fgrade = None
        self.q6fcomment = None

        self.q7bonusgrade = None
        self.q7bonuscomment = None

        self.score_q1 = 0.0
        self.score_q2 = 0.0
        self.score_q3 = 0.0
        self.score_q4 = 0.0
        self.score_q5 = 0.0
        self.score_q6 = 0.0
        self.score_q7 = 0.0

    def print_me(self):
        print( "Printing OrgReport..." )
        print( "\t studentName: {}".format( self.studentName ))
        
        print( "\t q1grade: {}".format( self.q1grade ))
        print( "\t q1comment: {}".format( self.q1comment ))
        
        print( "\t q2agrade: {}".format( self.q2agrade ))
        print( "\t q2comment: {}".format( self.q2acomment ))
        print( "\t q2bgrade: {}".format( self.q2bgrade ))
        print( "\t q2bcomment: {}".format( self.q2bcomment ))
        print( "\t q2cgrade: {}".format( self.q2cgrade ))
        print( "\t q2ccomment: {}".format( self.q2ccomment ))
        print( "\t q2dgrade: {}".format( self.q2dgrade ))
        print( "\t q2dcomment: {}".format( self.q2dcomment ))
        print( "\t q2egrade: {}".format( self.q2egrade ))
        print( "\t q2ecomment: {}".format( self.q2ecomment ))
        print( "\t q2fgrade: {}".format( self.q2fgrade ))
        print( "\t q2fcomment: {}".format( self.q2fcomment ))
        print( "\t q2fbonusgrade: {}".format( self.q2fbonusgrade ))
        print( "\t q2fbonuscomment: {}".format( self.q2fbonuscomment ))

        print( "\t q3agrade: {}".format( self.q3agrade ))
        print( "\t q3acomment: {}".format( self.q3acomment ))
        print( "\t q3bgrade: {}".format( self.q3bgrade ))
        print( "\t q3bcomment: {}".format( self.q3bcomment ))

        print( "\t q4agrade: {}".format( self.q4agrade ))
        print( "\t q4acomment: {}".format( self.q4acomment ))
        print( "\t q4bgrade: {}".format( self.q4bgrade ))
        print( "\t q4bcomment: {}".format( self.q4bcomment ))


        print( "\t q5DFSgrade: {}".format( self.q5DFSgrade ))
        print( "\t q5DFScomment: {}".format( self.q5DFScomment ))
        print( "\t q5BFSgrade: {}".format( self.q5BFSgrade ))
        print( "\t q5BFScomment: {}".format( self.q5BFScomment ))
        print( "\t q5HILLgrade: {}".format( self.q5HILLgrade ))
        print( "\t q5HILLcomment: {}".format( self.q5HILLcomment ))
        print( "\t q5BESTgrade: {}".format( self.q5BESTgrade ))
        print( "\t q5BESTcomment: {}".format( self.q5BESTcomment ))
        print( "\t q5ASTARgrade: {}".format( self.q5ASTARgrade ))
        print( "\t q5ASTARcomment: {}".format( self.q5ASTARcomment ))
        print( "\t q5fgrade: {}".format( self.q5fgrade ))
        print( "\t q5fcomment: {}".format( self.q5fcomment ))

        print( "\t q6agrade: {}".format( self.q6agrade ))
        print( "\t q6acomment: {}".format( self.q6acomment ))
        print( "\t q6bgrade: {}".format( self.q6bgrade ))
        print( "\t q6bcomment: {}".format( self.q6bcomment ))
        print( "\t q6cgrade: {}".format( self.q6cgrade ))
        print( "\t q6ccomment: {}".format( self.q6ccomment ))
        print( "\t q6dgrade: {}".format( self.q6dgrade ))
        print( "\t q6dcomment: {}".format( self.q6dcomment ))
        print( "\t q6egrade: {}".format( self.q6egrade ))
        print( "\t q6ecomment: {}".format( self.q6ecomment ))
        print( "\t q6fgrade: {}".format( self.q6fgrade ))
        print( "\t q6fcomment: {}".format( self.q6fcomment ))

        print( "\t q7bonusgrade: {}".format( self.q7bonusgrade ))
        print( "\t q7bonuscomment: {}".format( self.q7bonuscomment ))

    def _none2good( self, x ):
        if not x:
            return "good."
        return x
    
    def write_to_file(self, fname):
        f = open( fname, 'w' )
        f.write( "#+TITLE: Midterm Report for student: {}\n".format( self.studentName ) )
        f.write( "#+SUBTITLE: Score: {}\n\n".format( self.score_midterm ))
        
        f.write( "* Question 1 [{}/8 points]\n".format( self.score_q1 ))
        f.write( "+ q1: {}/8, notes: {}\n\n".format( self.q1grade, self._none2good(self.q1comment) ))

        f.write( "* Question 2 [{}/30 points]\n".format( self.score_q2 ) )
        f.write( "+ q2a: {}/5, notes: {}\n".format( self.q2agrade, self._none2good(self.q2acomment) ))
        f.write( "+ q2b: {}/5, notes: {}\n".format( self.q2bgrade, self._none2good(self.q2bcomment) ))
        f.write( "+ q2c: {}/5, notes: {}\n".format( self.q2cgrade, self._none2good(self.q2ccomment) ))
        f.write( "+ q2d: {}/5, notes: {}\n".format( self.q2dgrade, self._none2good(self.q2dcomment) ))
        f.write( "+ q2e: {}/5, notes: {}\n".format( self.q2egrade, self._none2good(self.q2ecomment) ))
        f.write( "+ q2f: {}/5, notes: {}\n".format( self.q2fgrade, self._none2good(self.q2fcomment) ))
        f.write( "+ q2fbonus/5: {}, notes: {}\n\n".format( self.q2fbonusgrade,
                                                           self.q2fbonuscomment ))


        f.write( "* Question 3 [{}/10 points] \n".format(self.score_q3) )
        f.write( "+ q3a: {}/5, notes: {}\n".format( self.q3agrade, self._none2good(self.q3acomment) ))
        f.write( "+ q3b: {}/5, notes: {}\n\n".format( self.q3bgrade, self._none2good(self.q3bcomment) ))

        f.write( "* Question 4 [{}/10 points]\n".format(self.score_q4) )
        f.write( "+ q4a: {}/5, notes: {}\n".format( self.q4agrade, self._none2good(self.q4acomment) ))
        f.write( "+ q4b: {}/5, notes: {}\n\n".format( self.q4bgrade, self._none2good(self.q4bcomment) ))

        f.write( "* Question 5 [{}/30 points]\n".format(self.score_q5) )
        f.write( "+ q5 (DFS): {}/5, notes: {}\n".format( self.q5DFSgrade, self._none2good(self.q5DFScomment) ))
        f.write( "+ q5 (BFS): {}/5, notes: {}\n".format( self.q5BFSgrade, self._none2good(self.q5BFScomment) ))
        f.write( "+ q5 (Hill Climbing): {}/5, notes: {}\n".format( self.q5HILLgrade,
                                                                   self._none2good(self.q5HILLcomment) ))
        f.write( "+ q5 (Best First): {}/5, notes: {}\n".format( self.q5BESTgrade,
                                                                self._none2good(self.q5BESTcomment) ))
        f.write( "+ q5 (A*): {}/5, notes: {}\n".format( self.q5ASTARgrade,
                                                        self._none2good(self.q5ASTARcomment) ))
        f.write( "+ q5f: {}/5, notes: {}\n\n".format( self.q5fgrade, self._none2good(self.q5fcomment) ))

        f.write( "* Question 6 [{}/12 points]\n".format( self.score_q6 ) )
        f.write( "+ q6a: {}/2, notes: {}\n".format( self.q6agrade, self._none2good(self.q6acomment) ))
        f.write( "+ q6b: {}/2, notes: {}\n".format( self.q6bgrade, self._none2good(self.q6bcomment) ))
        f.write( "+ q6c: {}/2, notes: {}\n".format( self.q6cgrade, self._none2good(self.q6ccomment) ))
        f.write( "+ q6d: {}/2, notes: {}\n".format( self.q6dgrade, self._none2good(self.q6dcomment) ))
        f.write( "+ q6e: {}/2, notes: {}\n".format( self.q6egrade, self._none2good(self.q6ecomment) ))
        f.write( "+ q6f: {}/2, notes: {}\n\n".format( self.q6fgrade, self._none2good(self.q6fcomment) ))

        f.write( "* Question 7 (bonus) [{} points]\n".format( self.score_q7 ) )
        f.write( "+ q7bonus: {}, notes: {}\n\n".format( self.q7bonusgrade, self.q7bonuscomment ))
        f.close()

class OrgReportFactory(object):
    def __init__(self, sheet):
        self.sheet = sheet

    def _getVal( self, targetLabel, sheet, rowID ):
        #print( "[debug] in _getVal() ..." )
        #print( "\t[debug] targetLabel: {}".format( targetLabel ))
        #print( "\t[debug] rowID: {}".format( rowID ))
        colID = colMap[ targetLabel ]
        #print( "\t[debug] colID: {}".format( colID ))
        cell_obj = sheet.cell( row=rowID, column=colID )
        val = cell_obj.value
        #print( "\t[debug] val: {}".format( val ))
        return val

    def genReport(self, rowID ):
        r = OrgReport()
        r.studentName = self._getVal( "studentName", self.sheet, rowID )
        r = self._pop_q1( rowID, r )
        r = self._pop_q2( rowID, r )
        r = self._pop_q3( rowID, r )
        r = self._pop_q4( rowID, r )
        r = self._pop_q5( rowID, r )
        r = self._pop_q6( rowID, r )
        r = self._pop_q7( rowID, r )
        r.score_midterm = r.score_q1 + r.score_q2 + r.score_q3 + r.score_q4 + \
            r.score_q5 + r.score_q6 + r.score_q7
        return r

    def _pop_q1( self, rowID, r ):
        r.q1grade = self._getVal( "q1grade", self.sheet, rowID )
        r.q1comment = self._getVal ( "q1comment", self.sheet, rowID )
        #print( "[debug] r.q1grade: {}".format( r.q1grade ))
        r.score_q1 = float( r.q1grade )
        return r
    
    def _pop_q2( self, rowID, r ):
        r.q2agrade = self._getVal( "q2agrade", self.sheet, rowID )
        r.q2acomment = self._getVal( "q2acomment", self.sheet, rowID )
        r.q2bgrade = self._getVal( "q2bgrade", self.sheet, rowID )
        r.q2bcomment = self._getVal( "q2bcomment", self.sheet, rowID )
        r.q2cgrade = self._getVal( "q2cgrade", self.sheet, rowID )
        r.q2ccomment = self._getVal( "q2ccomment", self.sheet, rowID )
        r.q2dgrade = self._getVal( "q2dgrade", self.sheet, rowID )
        r.q2dcomment = self._getVal( "q2dcomment", self.sheet, rowID )
        r.q2egrade = self._getVal( "q2egrade", self.sheet, rowID )
        r.q2ecomment = self._getVal( "q2ecomment", self.sheet, rowID )
        r.q2fgrade = self._getVal( "q2fgrade", self.sheet, rowID )
        r.q2fcomment = self._getVal( "q2fcomment", self.sheet, rowID )
        r.q2fbonusgrade = self._getVal( "q2fbonusgrade", self.sheet, rowID )
        r.score_q2 = float(r.q2agrade) + float(r.q2bgrade) + float(r.q2cgrade) + float(r.q2dgrade) + \
            float(r.q2egrade) + float(r.q2fgrade) + float(r.q2fbonusgrade)
        return r

    def _pop_q3( self, rowID, r ):
        r.q3agrade = self._getVal( "q3agrade", self.sheet, rowID )
        r.q3acomment = self._getVal( "q3acomment", self.sheet, rowID )
        r.q3bgrade = self._getVal( "q3bgrade", self.sheet, rowID )
        r.q3bcomment = self._getVal( "q3bcomment", self.sheet, rowID )
        r.score_q3 = float(r.q3agrade) + float(r.q3bgrade)
        return r

    def _pop_q4( self, rowID, r ):
        r.q4agrade = self._getVal( "q4agrade", self.sheet, rowID )
        r.q4acomment = self._getVal( "q4acomment", self.sheet, rowID )
        r.q4bgrade = self._getVal( "q4bgrade", self.sheet, rowID )
        r.q4bcomment = self._getVal( "q4bcomment", self.sheet, rowID )
        r.score_q4 = float(r.q4agrade) + float(r.q4bgrade)
        return r

    def _pop_q5( self, rowID, r ):
        r.q5DFSgrade = self._getVal( "q5DFSgrade", self.sheet, rowID )
        r.q5DFScomment = self._getVal( "q5DFScomment", self.sheet, rowID )
        r.q5BFSgrade = self._getVal( "q5BFSgrade", self.sheet, rowID )
        r.q5BFScomment = self._getVal( "q5BFScomment", self.sheet, rowID )
        r.q5HILLgrade = self._getVal( "q5HILLgrade", self.sheet, rowID )
        r.q5HILLcomment = self._getVal( "q5HILLcomment", self.sheet, rowID )
        r.q5BESTgrade = self._getVal( "q5BESTgrade", self.sheet, rowID )
        r.q5BESTcomment = self._getVal( "q5BESTcomment", self.sheet, rowID )
        r.q5ASTARgrade = self._getVal( "q5ASTARgrade", self.sheet, rowID )
        r.q5ASTARcomment = self._getVal( "q5ASTARcomment", self.sheet, rowID )
        r.q5fgrade = self._getVal( "q5fgrade", self.sheet, rowID )
        r.q5fcomment = self._getVal( "q5fcomment", self.sheet, rowID )
        r.score_q5 = float(r.q5DFSgrade) + float(r.q5BFSgrade) + float(r.q5HILLgrade) + \
            float(r.q5BESTgrade) + float(r.q5ASTARgrade) + float(r.q5fgrade)
        return r

    def _pop_q6( self, rowID, r ):
        r.q6agrade = self._getVal( "q6agrade", self.sheet, rowID )
        r.q6acomment = self._getVal( "q6acomment", self.sheet, rowID )
        r.q6bgrade = self._getVal( "q6bgrade", self.sheet, rowID )
        r.q6bcomment = self._getVal( "q6bcomment", self.sheet, rowID )
        r.q6cgrade = self._getVal( "q6cgrade", self.sheet, rowID )
        r.q6ccomment = self._getVal( "q6ccomment", self.sheet, rowID )
        r.q6dgrade = self._getVal( "q6dgrade", self.sheet, rowID )
        r.q6dcomment = self._getVal( "q6dcomment", self.sheet, rowID )
        r.q6egrade = self._getVal( "q6egrade", self.sheet, rowID )
        r.q6ecomment = self._getVal( "q6ecomment", self.sheet, rowID )
        r.q6fgrade = self._getVal( "q6fgrade", self.sheet, rowID )
        r.q6fcomment = self._getVal( "q6fcomment", self.sheet, rowID )
        r.score_q6 = float(r.q6agrade) + float(r.q6bgrade) + float(r.q6cgrade) + float(r.q6dgrade) + \
            float(r.q6egrade) + float(r.q6fgrade)
        return r

    def _pop_q7( self, rowID, r ):
        r.q7bonusgrade = self._getVal( "q7bonusgrade", self.sheet, rowID )
        r.q7bonuscomment = self._getVal( "q7bonuscomment", self.sheet, rowID )
        r.score_q7 = float(r.q7bonusgrade)
        return r
