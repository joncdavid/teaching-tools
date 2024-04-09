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
        self.student1Name = None
        self.student2Name = None
        self.student3Name = None
        self.student4Name = None
        
        self.p_grade = 0.0
        self.project_score = 0.0

        self.pdf = 0.0
        self.pdf_comments = None
        self.pkg = 0.0
        self.pkg_comments = None
        self.manifest = 0.0
        self.manifest_comments = None
        self.team_contributions = 0.0
        self.team_contributions_comments = None
        self.kaggle = 0.0
        self.kaggle_comments = None
        self.deliverables_score = 0.0

        self.clean = 0.0
        self.clean_comments = None
        self.ig_impl = 0.0
        self.ig_impl_comments = None
        self.tree_impl = 0.0
        self.tree_impl_comments = None
        self.forest_impl = 0.0
        self.forest_impl_comments = None
        self.code_score = 0.0

        self.p1 = 0.0
        self.p1_comment = None
        self.p2 = 0.0
        self.p2_comment = None
        self.p3 = 0.0
        self.p3_comment = None
        self.p4 = 0.0
        self.p4_comment = None
        self.p5 = 0.0
        self.p5_comment = None
        self.p6 = 0.0
        self.p6_comment = None
        self.p7 = 0.0
        self.p7_comment = None
        self.analysis_score = 0.0

    def print_me(self):
        print( "Printing OrgReport..." )
        if self.student2Name is None:
            print( "\t studentName: {}".format( self.student1Name ))
        elif self.student3Name is None:
            print("\t studentNames: {}, {}".format(self.student1Name, self.student2Name))
        elif self.student4Name is None:
            print("\t studentNames: {}, {}, {}".format(self.student1Name, self.student2Name, self.student3Name))
        else:
            print("\t studentNames: {}, {}, {}, {}".format(self.student1Name, self.student2Name, self.student3Name, self.student4Name))

        print( "\t overall grade: {}".format( self.p_grade ))

        print( "\t pdf: {}".format( self.pdf ))
        print( "\t pdf_comments: {}".format( self.pdf_comments ))
        print( "\t pkg: {}".format( self.pkg ))
        print( "\t pkg_comments: {}".format( self.pkg_comments ))
        print( "\t manifest: {}".format( self.manifest ))
        print( "\t manifest_comments: {}".format( self.manifest_comments ))
        print( "\t team_contributions: {}".format( self.team_contributions ))
        print( "\t team_contributions_comments: {}".format( self.team_contributions_comments ))
        print( "\t kaggle: {}".format( self.kaggle ))
        print( "\t kaggle_comments: {}".format( self.kaggle_comments ))

        print( "\t clean: {}".format( self.clean ))
        print( "\t clean_comments: {}".format( self.clean_comments ))
        print( "\t ig_impl: {}".format( self.ig_impl ))
        print( "\t ig_impl_comments: {}".format( self.ig_impl_comments ))
        print( "\t tree_impl: {}".format( self.tree_impl ))
        print( "\t tree_impl_comments: {}".format( self.tree_impl_comments ))
        print( "\t forest_impl: {}".format( self.forest_impl ))
        print( "\t forest_impl_comments: {}".format( self.forest_impl_comments ))

        print( "\t p1: {}".format( self.p1 ))
        print( "\t p1_comment: {}".format( self.p1_comment ))
        print( "\t p2: {}".format( self.p2 ))
        print( "\t p2_comment: {}".format( self.p2_comment ))
        print( "\t p3: {}".format( self.p3 ))
        print( "\t p3_comment: {}".format( self.p3_comment ))
        print( "\t p4: {}".format( self.p4 ))
        print( "\t p4_comment: {}".format( self.p4_comment ))
        print( "\t p5: {}".format( self.p5 ))
        print( "\t p5_comment: {}".format( self.p5_comment ))
        print( "\t p6: {}".format( self.p6 ))
        print( "\t p6_comment: {}".format( self.p6_comment ))
        print( "\t p7 (bonus): {}".format( self.p7 ))
        print( "\t p7_comment: {}".format( self.p7_comment ))

    def _none2good( self, x ):
        if not x:
            return "good."
        return x
    
    def write_to_file(self, fname):
        f = open( fname, 'w' )
        if self.student2Name is None:
            student_names = "{}".format(self.student1Name)
        elif self.student3Name is None:
            student_names = "{}, {}".format(self.student1Name, self.student2Name)
        elif self.student4Name is None:
            student_names = "{}, {}, {}".format(self.student1Name, self.student2Name, self.student3Name)
        else:
            student_names = "{}, {}, {}, {}".format(self.student1Name, self.student2Name, self.student3Name,
                                                           self.student4Name)

        f.write( "Project 1 Report for: {}\n".format( student_names ) )
        f.write( "Score: {}/90\n\n".format( self.project_score ))

        f.write( "Deliverables [{}/10 points]\n".format(self.deliverables_score))
        f.write( "+ PDF: {}/2, notes: {}\n".format( self.pdf, self._none2good(self.pdf_comments) ))
        f.write( "+ Package: {}/2, notes: {}\n".format( self.pkg, self._none2good(self.pkg_comments) ))
        f.write( "+ File Manifest: {}/2, notes: {}\n".format( self.manifest, self._none2good(self.manifest_comments) ))
        f.write( "+ Team Contributions: {}/2, notes: {}\n".format( self.team_contributions, self._none2good(self.team_contributions_comments) ))
        f.write( "+ Kaggle Score: {}/2, notes: {}\n\n".format( self.kaggle, self._none2good(self.kaggle_comments) ))

        f.write( "Code [{}/35 points] \n".format(self.code_score) )
        f.write( "+ Clean: {}/5, notes: {}\n".format( self.clean, self._none2good(self.clean_comments) ))
        f.write( "+ IG Methods Implementation: {}/10, notes: {}\n".format( self.ig_impl, self._none2good(self.ig_impl_comments) ))
        f.write( "+ Decision Tree Implementation: {}/10, notes: {}\n".format( self.tree_impl, self._none2good(self.tree_impl_comments) ))
        f.write( "+ Random Forest Implementation: {}/10, notes: {}\n\n".format( self.forest_impl, self._none2good(self.forest_impl_comments) ))

        f.write( "Analysis [{}/45 points]\n".format(self.analysis_score) )
        f.write( "+ Compare/contrast of IG methods: {}/10, notes: {}\n".format( self.p1, self._none2good(self.p1_comment) ))
        f.write( "+ Compare/contrast of alphas for Chi Square: {}/10, notes: {}\n".format( self.p2, self._none2good(self.p2_comment) ))
        f.write( "+ Selected IG Criteria Method: {}/5, notes: {}\n".format( self.p3, self._none2good(self.p3_comment) ))
        f.write( "+ Missing Data: {}/5, notes: {}\n".format( self.p4, self._none2good(self.p4_comment) ))
        f.write( "+ Numerical Features: {}/5, notes: {}\n".format( self.p5, self._none2good(self.p5_comment) ))
        f.write( "+ Data Imbalance: {}/10, notes: {}\n".format( self.p6, self._none2good(self.p6_comment) ))
        f.write( "+ Optimizations (bonus): {}/10, notes: {}\n".format(self.p7, self.p7_comment))

        f.close()

class OrgReportFactory(object):
    def __init__(self, sheet):
        self.sheet = sheet

    def _getVal( self, targetLabel, sheet, rowID ):
        colID = colMap[ targetLabel ]
        cell_obj = sheet.cell( row=rowID, column=colID )
        val = cell_obj.value
        return val

    def genReport(self, rowID ):
        r = OrgReport()
        r.student1Name = self._getVal( "student1Name", self.sheet, rowID )
        r.student2Name = self._getVal("student2Name", self.sheet, rowID)
        r.student3Name = self._getVal("student3Name", self.sheet, rowID)
        r.student4Name = self._getVal("student4Name", self.sheet, rowID)
        r = self._pop_overall_project_grade(rowID, r)
        r = self._pop_deliverables(rowID, r)
        r = self._pop_code(rowID, r)
        r = self._pop_analysis(rowID, r)
        r.project_score = r.deliverables_score + r.analysis_score + r.code_score
        return r

    def _pop_overall_project_grade(self, rowID, r):
        r.p_grade =self._getVal('p_grade', self.sheet, rowID)
        # r.p_score = float(r.p_grade)
        return r

    def _pop_deliverables(self, rowID, r):
        r.pdf = self._getVal( "pdf", self.sheet, rowID )
        r.pdf_comments = self._getVal ( "pdf_comments", self.sheet, rowID )
        r.pkg = self._getVal("pkg", self.sheet, rowID)
        r.pkg_comments = self._getVal("pkg_comments", self.sheet, rowID)
        r.manifest = self._getVal("manifest", self.sheet, rowID)
        r.manifest_comments = self._getVal("manifest_comments", self.sheet, rowID)
        r.team_contributions = self._getVal("team_contributions", self.sheet, rowID)
        r.team_contributions_comments = self._getVal("team_contributions_comments", self.sheet, rowID)
        r.kaggle = self._getVal("kaggle", self.sheet, rowID)
        r.kaggle_comments = self._getVal("kaggle_comments", self.sheet, rowID)
        r.deliverables_score = float(r.pdf) + float(r.pkg) + float(r.manifest) + float(r.team_contributions) + \
                               float(r.kaggle)
        return r
    
    def _pop_code(self, rowID, r):
        r.clean = self._getVal( "clean", self.sheet, rowID )
        r.clean = max(0, r.clean)
        r.clean_comments = self._getVal( "clean_comments", self.sheet, rowID )
        r.ig_impl = self._getVal( "ig_impl", self.sheet, rowID )
        r.ig_impl_comments = self._getVal( "ig_impl_comments", self.sheet, rowID )
        r.tree_impl = self._getVal( "tree_impl", self.sheet, rowID )
        r.tree_impl = max(0, r.tree_impl)
        r.tree_impl_comments = self._getVal( "tree_impl_comments", self.sheet, rowID )
        r.forest_impl = self._getVal( "forest_impl", self.sheet, rowID )
        r.forest_impl_comments = self._getVal( "forest_impl_comments", self.sheet, rowID )
        r.code_score = float(r.clean) + float(r.ig_impl) + float(r.tree_impl) + float(r.forest_impl)
        return r

    def _pop_analysis(self, rowID, r):
        r.p1 = self._getVal( "p1", self.sheet, rowID )
        r.p1_comment = self._getVal( "p1_comment", self.sheet, rowID )
        r.p2 = self._getVal( "p2", self.sheet, rowID )
        r.p2_comment = self._getVal( "p2_comment", self.sheet, rowID )
        r.p3 = self._getVal( "p3", self.sheet, rowID )
        r.p3_comment = self._getVal( "p3_comment", self.sheet, rowID )
        r.p4 = self._getVal( "p4", self.sheet, rowID )
        r.p4_comment = self._getVal( "p4_comment", self.sheet, rowID )
        r.p5 = self._getVal( "p5", self.sheet, rowID )
        r.p5_comment = self._getVal( "p5_comment", self.sheet, rowID )
        r.p6 = self._getVal( "p6", self.sheet, rowID )
        r.p6_comment = self._getVal( "p6_comment", self.sheet, rowID )
        r.p7 = self._getVal("p7", self.sheet, rowID)
        r.p7_comment = self._getVal("p7_comment", self.sheet, rowID)
        r.analysis_score = float(r.p1) + float(r.p2) + float(r.p3) + \
            float(r.p4) + float(r.p5) + float(r.p6) + float(r.p7)
        return r

