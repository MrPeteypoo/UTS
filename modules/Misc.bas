Attribute VB_Name = "Misc"
''  Program: Unit Tracking System (UTS)
''  Created by Simon Peter Campbell
''
''  The purpose of UTS is to provide an automated tracking system for individual BTEC units.
''  The aim is to create a fully functional and bug free system that is scalable in terms
''  of different amounts of criteria, students & sorting preferences.
''
''  All miscellaneous functions and procedures appear here. This module is for any algorithms
''  that don't fit into the other modules and are usually more general purpose.
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
''                       DD/MM/YYYY
''  Version 1.0.1 FINAL (27/09/2012)

Public Pass As Integer
Public Merit As Integer
Public Distinction As Integer
Public Students As Integer

Private Sub TrackerSettings_Click()
    Load frmSettings
    frmSettings.Show
End Sub

Function obtainColLetter(ByVal address As String) As String
'' A simple function that returns the column letter for any cell via its address function.
'' (columnabsolute:=false) must be specified or this won't work.
    obtainColLetter = Left(address & "", Application.WorksheetFunction.Find(":", address) - 1)
End Function

Sub writeVariables()
    variables.Select
    Dim Students As Range:      Set Students = Cells(6, 2)
    Dim Passes As Range:        Set Passes = Cells(7, 2)
    Dim Merits As Range:        Set Merits = Cells(8, 2)
    Dim Distinctions As Range:  Set Distinctions = Cells(9, 2)
    '' 10 is skipped because it is no longer needed
    Dim Assignments As Range:   Set Assignments = Cells(11, 2)
    Dim criteria As Range:      Set criteria = Cells(12, 2)
    Dim GradeColumn As Range:   Set GradeColumn = Cells(13, 2)
    Dim Bottom As Range:        Set Bottom = Cells(14, 2)
    Dim sortType As Range:      Set sortType = Cells(15, 2)
    Dim courseTitle As Range:   Set courseTitle = Cells(16, 2)
    Dim unitTitle As Range:     Set unitTitle = Cells(17, 2)
    Dim groupTitle As Range:    Set groupTitle = Cells(18, 2)
    
    Students.Value = frmSettings.numStudents.Value
    Passes.Value = frmSettings.numPass.Value
    Merits.Value = frmSettings.numMerit.Value
    Distinctions.Value = frmSettings.numDistinction.Value
    
    Dim numCriteria As Integer: numCriteria = Passes.Value + Merits.Value + Distinctions.Value
    Dim endCriteria As String: endCriteria = obtainColLetter(Cells(1, 4 + numCriteria).EntireColumn.address(columnabsolute:=False))
    
    Assignments.Value = "E7:" & endCriteria & 8
    criteria.Value = "E9:" & endCriteria & 8 + Students.Value
    GradeColumn.Value = obtainColLetter(Cells(1, 5 + numCriteria).EntireColumn.address(columnabsolute:=False))
    Bottom.Value = "B" & 9 + frmSettings.numStudents.Value
    If frmSettings.radAlphabet.Value = True Then sortType.Value = 1
    If frmSettings.radGrade.Value = True Then Cells("15", "B").Value = 2
    If frmSettings.radLeader.Value = True Then Cells("15", "B").Value = 3
    courseTitle.Value = frmSettings.txtCourse.Value
    unitTitle.Value = frmSettings.txtUnit.Value
    groupTitle.Value = frmSettings.txtGroup.Value
    Unit1.Select
End Sub

Sub determineCriteria(ByRef result As Integer, ByVal target As Range)
'' This procedure finds out what column the modified cell is in. This happens to make
'' sure that the cells are coloured correctly.
    Dim critLoop As Boolean: critLoop = True
    target.Select
    Do
        If Selection.Columns.count > 1 Then
            Select Case Selection.Cells(1, 1).Value
            Case "PASS"
                result = 1
            Case "MERIT"
                result = 2
            Case "DISTINCTION"
                result = 3
            End Select
        End If
        If Not ActiveCell.row = 1 Then
            Selection.Offset(-1, 0).Select
        Else
            critLoop = False
        End If
    Loop Until critLoop = False
End Sub

Sub StudentColour(ByVal row As Integer, ByVal target As Range)
'' A simple procedure that colours in an entire students row if they have achieved
'' a grade.
    Dim grade As Range: Set grade = Range(variables.Cells(13, 2).Value & row)
    Range("B" & row, "C" & row).Select
    Select Case grade.Value
    Case "Distinction"
        ColourDistinction
        Range(grade, grade.Offset(0, 2)).Select
        ColourDistinction '' colour the heading cells too.
    Case "Merit"
        ColourMerit
        Range(grade, grade.Offset(0, 2)).Select
        ColourMerit
    Case "Pass"
        ColourPass
        Range(grade, grade.Offset(0, 2)).Select
        ColourPass
    Case "Pass Referral"
        ColourPassR
        Range(grade, grade.Offset(0, 2)).Select
        ColourPassR
    Case "Unsafe"
        ColourRisk
        Range(grade, grade.Offset(0, 2)).Select
        ColourRisk
    Case Else
        If target.address <> ActiveCell.EntireRow.address Then
            ColourRemove
            Range(grade, grade.Offset(0, 2)).Select
            ColourRemove '' z is used for sorting reasons if no grade has been achieved
                         '' so the grade column is given white text to make it invisible.
            grade.Font.ThemeColor = xlThemeColorDark1
            grade.Font.TintAndShade = 0
        End If
    End Select
End Sub
