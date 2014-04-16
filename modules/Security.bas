Attribute VB_Name = "Security"
''  Program: Unit Tracking System (UTS)
''  Created by Simon Peter Campbell
''
''  The purpose of UTS is to provide an automated tracking system for individual BTEC units.
''  The aim is to create a fully functional and bug free system that is scalable in terms
''  of different amounts of criteria, students & sorting preferences.
''
''  This module covers everything to do with spreadsheet security. The locking of cells and
''  protecting of sheets is all done here.
''
''  Copyright 2012 Simon Peter Campbell
''  This file is part of Unit Tracking System (UTS).
''
''  Unit Tracking System is free software: you can redistribute it and/or modify it under
''  the terms of the GNU General Public License as published by the Free Software Foundation,
''  either version 3 of the License, or any later version. Unit Tracking System is distributed
''  in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
''  warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General
''  Public License for more details. You should have received a copy of the GNU General Public
''  License along with Unit Tracking System. If not, see <http://www.gnu.org/licenses/>.
''
''
''                    DD/MM/YYYY
''  Version 1.0 BETA (24/07/2012)
''
''  TODO:
''  1) Fully test the code - In Progress

Sub doSecurity()
    Dim Pass As Integer:          Pass = frmSettings.numPass.Value
    Dim Merit As Integer:         Merit = frmSettings.numMerit.Value
    Dim Distinction As Integer:   Distinction = frmSettings.numDistinction.Value
    Dim numCriteria As Integer:   numCriteria = Pass + Merit + Distinction
    Dim numStudents As Integer:   numStudents = frmSettings.numStudents.Value
    
    Dim topRow As Range:        Set topRow = Range("B6", Cells(6, 4 + numCriteria))
    Dim topTitles As Range:     Set topTitles = Range("B6", "D8")
    Dim criteriaRow As Range:   Set criteriaRow = Range("E8", Cells(8, 4 + numCriteria))
    Dim support As Range:       Set support = Range("D9", Cells(8 + numStudents, 4))
    Dim gradeCol As Range:      Set gradeCol = Range(Cells(6, 5 + numCriteria), Cells(8 + numStudents, 5 + numCriteria))
    Dim notesCell As Range:     Set notesCell = Range(Cells(6, 5 + numCriteria + 1), Cells(8, 5 + numCriteria + 1))
    Dim pointCol As Range:      Set pointCol = Range(Cells(6, 5 + numCriteria + 2), Cells(8 + numStudents, 5 + numCriteria + 2))
    Dim courseTitle As Range:   Set courseTitle = Cells(2, 5 + numCriteria)
    Dim unitGroup As Range:     Set unitGroup = Range(Cells(4, 5 + numCriteria), Cells(4, 5 + numCriteria + 1))
    
    Unit1.Cells.Locked = False
    variables.Cells.Locked = False
    topRow.Locked = True
    topTitles.Locked = True
    criteriaRow.Locked = True
    support.Locked = True
    gradeCol.Locked = True
    notesCell.Locked = True
    pointCol.Locked = True
    courseTitle.MergeArea.Locked = True
    unitGroup.Locked = True
        
    Unit1.Activate: ProtectSheet '' This is a recorded macro
    variables.Activate: ProtectSheet
    variables.Visible = xlSheetHidden
    Unit1.Activate
End Sub

Sub unlockSheets()
    Unit1.Unprotect Password:="21102560"
    variables.Unprotect Password:="21102560"
End Sub

