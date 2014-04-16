Attribute VB_Name = "Criteria"
''  Program: Unit Tracking System (UTS)
''  Created by Simon Peter Campbell
''
''  The purpose of UTS is to provide an automated tracking system for individual BTEC units.
''  The aim is to create a fully functional and bug free system that is scalable in terms
''  of different amounts of criteria, students & sorting preferences.
''
''  This module contains the code that adds and removes each students criteria.
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
''  2) Fully test rewritten code.

''  Modular  ''

Sub DoCriteria()
    Dim Pass As Integer:        Pass = frmSettings.numPass.Value
    Dim Merit As Integer:       Merit = frmSettings.numMerit.Value
    Dim Distinction As Integer: Distinction = frmSettings.numDistinction.Value
    Dim criteria As Integer:    criteria = Pass + Merit + Distinction
    Dim Students As Integer:    Students = frmSettings.numStudents.Value
    '' These integers store the column number that addCriteria should start looping from.
    Dim PassHome As Integer:        PassHome = 5
    Dim MeritHome As Integer:       MeritHome = 5 + Pass
    Dim DistinctionHome As Integer: DistinctionHome = 5 + Pass + Merit
    Dim HomeAdjustment As Integer
    
    '' Unmerge all merged titles to prevent errors when adding/deleting columns.
    Range("E6", Cells(6, 6 + criteria)).UnMerge
    addCriteria "P", "M", 11, Pass, PassHome
    addCriteria "M", "D", 6, Merit, MeritHome
    addCriteria "D", "P", 4, Distinction, DistinctionHome
    addDeadlines criteria, Students
    
    '' Bold and centre all headings!
    With Range("E6", Cells(8, 4 + criteria))
        .Font.Bold = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .ColumnWidth = 3.57
    End With
    '' Format the criteria cells.
    With Range("E9", Cells(8 + Students, 4 + criteria))
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.name = "Wingdings 2"
        .Font.Size = 11
        .Font.Strikethrough = False
        .Font.Superscript = False
        .Font.Subscript = False
        .Font.OutlineFont = False
        .Font.Shadow = False
        .Font.Underline = xlUnderlineStyleNone
        .Font.TintAndShade = 0
        .Font.ThemeFont = xlThemeFontNone
    End With
    '' Next look for any missing assignment numbers and set them to 1.
    For i = 1 To criteria
        Cells(7, 4 + i).Select
        Select Case Selection.Value
        Case ""
            Selection.Value = 1
        End Select
    Next i
    '' Format the criteria headings (Pass, Merit and Distinction).
    criteriaHeadings "PASS", Range("E6", Cells(6, 4 + Pass))
    criteriaHeadings "MERIT", Range(Cells(6, 5 + Pass), Cells(6, 4 + Pass + Merit))
    criteriaHeadings "DISTINCTION", Range(Cells(6, 5 + Pass + Merit), Cells(6, 4 + criteria))
    addBorders Range(Cells(7, 5), Cells(9 + Students, 4 + criteria))
End Sub

Private Sub addCriteria(PMD As String, altPMD As String, maxCriteria As Integer, newCriteria As Integer, home As Integer)
'' The aim of this procedure is to determine how many criteria of a certain type (P, M or D)
'' exist on the tracker already, when the number has been determined the procedure will
'' add or remove any necessary columns.
    '' Discover how many Criteria are currently selected.
    Dim count As Integer:       count = 0
    Dim curCriteria As Integer: curCriteria = 0
    Dim loopRange As Range
    For i = 1 To maxCriteria
        Set loopRange = Cells(8, home + i - 1)
        If loopRange.Cells(1, 1).Value = PMD & i Then curCriteria = curCriteria + 1
    Next i '' Cells(1, 1) is used to make sure the tracker doesn't crash if it encounters a merged cell.
    
    '' Compare the current Criteria to the desired Criteria to find out if any
    '' criteria need to be added or removed.
    Cells(8, home + curCriteria).Select
    Select Case newCriteria
    Case Is > curCriteria
        For i = 1 To (newCriteria - curCriteria)
            Selection.EntireColumn.Insert shift:=xlRight
            Selection.EntireColumn.ClearFormats
            count = count + 1
        Next i
        For i = 1 To newCriteria
            Cells(8, home + i - 1).Value = (PMD & i)
        Next i
    Case Is < curCriteria
        Do
            Set loopRange = Cells(8, home + newCriteria)
            loopRange.EntireColumn.Delete
            count = count - 1
        Loop Until count = newCriteria - curCriteria
    End Select
End Sub

Private Sub criteriaHeadings(name As String, ByVal target As Range)
    With target
        .Merge
        .Value = name
        .Font.Bold = True
        .Font.Size = 12
    End With
    addBorders target
    
    Select Case name
    Case "PASS"
        If target.Cells.count < 2 Then target.Font.Size = 8
        ColourPass
    Case "MERIT"
        If target.Cells.count < 2 Then target.Font.Size = 8
        ColourMerit
    Case "DISTINCTION"
        If target.Cells.count <= 2 Then target.Font.Size = 8
        ColourDistinction
    End Select
End Sub

Private Sub addDeadlines(ByVal criteria As Integer, ByVal Students As Integer)
    Range(Cells(9 + Students, 5), Cells(9 + Students, 4 + criteria)).UnMerge
    Dim workingRange As Range: Set workingRange = Range(Cells(9 + Students, 5), Cells(9 + Students, 4 + criteria))
    Dim curCell As Range
    Dim tDate As Date: tDate = Date
    tDate = Format(Date, "dd-mm-yyyy")
    
    For Each curCell In workingRange
        With Range(curCell, curCell.Offset(3, 0))
            .Merge
            .Font.Bold = True
            .Font.Size = 12
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlRight
            .Orientation = -90
            .Select
        End With
        If Selection.Text = "" Then Selection.Value = tDate
    Next curCell
End Sub

