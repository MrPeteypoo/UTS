Attribute VB_Name = "Extras"
''  Program: Unit Tracking System (UTS)
''  Created by Simon Peter Campbell
''
''  The purpose of UTS is to provide an automated tracking system for individual BTEC units.
''  The aim is to create a fully functional and bug free system that is scalable in terms
''  of different amounts of criteria, students & sorting preferences.
''
''  This module contains the code that formats all the additional headings such as "overall
''  grade". It also fills in relevant formulas and formats the points column.
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
''
''  TODO:
''

Sub DoExtras()
    Dim Pass As Integer:          Pass = frmSettings.numPass.Value
    Dim Merit As Integer:         Merit = frmSettings.numMerit.Value
    Dim Distinction As Integer:   Distinction = frmSettings.numDistinction.Value
    Dim numCriteria As Integer:   numCriteria = Pass + Merit + Distinction
    Dim gradeCell As Range:   Set gradeCell = Cells(6, (5 + numCriteria))
    Dim notesCell As Range:   Set notesCell = gradeCell.Offset(0, 1)
    Dim pointCell As Range:   Set pointCell = notesCell.Offset(0, 1)
    Dim unitsCell As Range:   Set unitsCell = gradeCell.Offset(-4, 0)
    Dim keyHome As Range:     Set keyHome = unitsCell.Offset(0, 3)
    
    formatGrade gradeCell
    formatNotes notesCell
    formatPoint pointCell
    formatUnitTitles unitsCell
    formatExtras frmSettings.numStudents.Value
    formatKey keyHome
End Sub

Private Sub formatGrade(ByVal heading As Range)
'' Format the grades heading and assign the correct formula
    Dim Pass As Integer:          Pass = frmSettings.numPass.Value
    Dim Merit As Integer:         Merit = frmSettings.numMerit.Value
    Dim Distinction As Integer:   Distinction = frmSettings.numDistinction.Value
    Dim numCriteria As Integer:   numCriteria = Pass + Merit + Distinction
    Dim numStudents As Integer:   numStudents = frmSettings.numStudents.Value
    Dim PassEnd As String:        PassEnd = obtainColLetter(Cells(1, 4 + Pass).EntireColumn.address(columnabsolute:=False))
    Dim MeritEnd As String:       MeritEnd = obtainColLetter(Cells(1, 4 + Pass + Merit).EntireColumn.address(columnabsolute:=False))
    Dim DistinctionEnd As String: DistinctionEnd = obtainColLetter(Cells(1, 4 + numCriteria).EntireColumn.address(columnabsolute:=False))
    
    heading.UnMerge
    Range(heading.address, heading.Offset(2, 0)).Select
    ColourGrey
    With Selection
        .Merge
        .Value = "Overall Grade"
        .Font.Bold = True
        .Font.Size = 16
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .EntireColumn.ColumnWidth = 25
    End With
    Range(heading.Offset(1, 0), heading.Offset(numStudents, 0)).Select
    With Selection
        .ClearFormats
        .Font.Bold = True
        .Font.ThemeColor = xlThemeColorDark1
        .Font.TintAndShade = 0
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    
    Dim r As Integer: r = 9 '' The start of the student rows.
    Dim curCell As Range
    For Each curCell In Selection
'       Now the incredibly complex formula for calculating grades! I doubt it is understandable
'       so I broke it down in comment form with a mixture of Excel and pseudocode. It shows the
'       steps taken in determining the grade.
'       curCell.Formula = "=
'       IF(COUNTIF(criteria range = 'R')= number-criteria, "Distinction" grade,
'       IF(COUNTIF(pass + merit range = 'R')= pass+merit critera, "Merit" grade,
'       IF(COUNTIF(pass range = 'R')= pass criteria, "Pass" grade,
'       IF(SUM(COUNTIF(pass range = 'R'), COUNTIF(pass range = '8'))=pass criteria, "Pass Referral", _
'       IF(SUM(COUNTIF(pass range = 'R'), COUNTIF(pass range = 'T'), COUNTIF(pass range = '8'))=pass criteria, "Unsafe", "z" for sorting purposes))))))"
        curCell.Formula = "=IF(COUNTIF($E" & r & ":$" & DistinctionEnd & r & ", ""R"")=" & numCriteria & ", ""Distinction"",IF(COUNTIF($E" & r & ":$" & MeritEnd & r & ", ""R"")=" & Pass + Merit & ", ""Merit"",IF(COUNTIF($E" & r & ":$" & PassEnd & r & ", ""R"")=" & Pass & ", ""Pass"", IF(SUM(COUNTIF($E" & r & ":$" & PassEnd & r & ", ""R""), COUNTIF($E" & r & ":$" & PassEnd & r & ", ""8""))=" & Pass & ", ""Pass Referral"", IF(SUM(COUNTIF($E" & r & ":$" & PassEnd & r & ", ""R""), COUNTIF($E" & r & ":$" & PassEnd & r & ", ""T""), COUNTIF($E" & r & ":$" & PassEnd & r & ", ""8""))=" & Pass & ", ""Unsafe"", ""z"")))))"
        StudentColour r, Cells(curCell.row, curCell.Column)
        r = r + 1
    Next curCell
    addBorders Range(heading.address, heading.Offset(numStudents, 0))
End Sub

Private Sub formatNotes(ByVal heading As Range)
'' This procedure formats the notes column.
    Dim numStudents As Integer: numStudents = frmSettings.numStudents.Value
    heading.UnMerge
    Range(heading, heading.Offset(2, 0)).Select
    ColourGrey
    With Selection
        .Merge
        .Value = "Notes"
        .Font.Bold = True
        .Font.Size = 16
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .EntireColumn.ColumnWidth = 25
    End With
    With Range(heading.Offset(1, 0), heading.Offset(numStudents, 0))
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    addBorders Range(heading, heading.Offset(numStudents, 0))
End Sub

Private Sub formatPoint(ByVal heading As Range)
'' Simple procedure to give the correct formatting and formula to every cell in
'' the points heading.
    Dim Pass As Integer:          Pass = frmSettings.numPass.Value
    Dim Merit As Integer:         Merit = frmSettings.numMerit.Value
    Dim Distinction As Integer:   Distinction = frmSettings.numDistinction.Value
    Dim numCriteria As Integer:   numCriteria = Pass + Merit + Distinction
    Dim numStudents As Integer:   numStudents = frmSettings.numStudents.Value
    Dim PassEnd As String:        PassEnd = obtainColLetter(Cells(1, 4 + Pass).EntireColumn.address(columnabsolute:=False))
    Dim MeritEnd As String:       MeritEnd = obtainColLetter(Cells(1, 4 + Pass + Merit).EntireColumn.address(columnabsolute:=False))
    Dim DistinctionEnd As String: DistinctionEnd = obtainColLetter(Cells(1, 4 + numCriteria).EntireColumn.address(columnabsolute:=False))
    Dim sortType As Integer:      sortType = variables.Cells(15, 2).Value
    Dim gradeLetter As String:    gradeLetter = obtainColLetter(heading.Offset(0, -2).EntireColumn.address(columnabsolute:=False))
    Dim r As Integer: r = 9 '' The first student row
    '' Format all of the cells and add borders to them.
    heading.UnMerge
    Range(heading, heading.Offset(2, 0)).Select
    ColourGrey
    With Selection
        .Merge
        .Value = "Points"
        .Font.Bold = True
        .Font.Size = 12
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .Orientation = -90
    End With
    Range(heading.Offset(1, 0), heading.Offset(numStudents, 0)).Select
    With Selection
        .Font.Bold = True
        .Font.Size = 11
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .EntireColumn.ColumnWidth = 2.29
    End With
    '' Next add the correct formula depending on the users choice.
    Dim curCell As Range
    For Each curCell In Selection
        Select Case sortType
        Case 1 '' 1 is alphabetic
            curCell.Formula = 1
        Case 2 '' 2 is grade hence the complicated formula
            curCell.Formula = "=IF($" & gradeLetter & r & " = ""Distinction"", 4, IF($" & gradeLetter & r & " = ""Merit"", 3, IF($" & gradeLetter & r & " = ""Pass"",2,IF($" & gradeLetter & r & " = ""Pass Referral"",1,IF($" & gradeLetter & r & " = ""z"",0,IF($" & gradeLetter & r & " = ""Unsafe"",-1,0))))))"
        Case 3 '' 3 is points, 3 points for a pass, 1 for referral and -1 for a missed deadline
            curCell.Formula = "=SUM(COUNTIF($E" & r & ":$" & DistinctionEnd & r & ", ""R"") * 3, COUNTIF($E" & r & ":$" & DistinctionEnd & r & ", ""8"") * 1, COUNTIF($E" & r & ":$" & DistinctionEnd & r & ", ""T"") * -1)"
        End Select
        r = r + 1
    Next curCell
    '' Since the points column isn't necessary for grade of alphabetic, hide the column.
    If sortType = 1 Or sortType = 2 Then
        ActiveCell.EntireColumn.Hidden = True
    ElseIf sortType = 3 Then
        ActiveCell.EntireColumn.Hidden = False
    End If
    addBorders Range(heading.Offset(0, -1), heading.Offset(numStudents, 0))
End Sub

Private Sub formatUnitTitles(ByVal heading As Range)
'' The extra titles that display the unit name, group code and course title are formatted here.
    Dim workingRange As Range: Set workingRange = Range(heading, heading.Offset(3, 0))
    workingRange.UnMerge
    With Range(workingRange.Cells(1, 1), workingRange.Cells(1, 2))
        .Merge
        .Value = "Course Title"
        .Font.Bold = True
        .Font.Size = 12
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Select
    End With
    ColourGrey
    With Range(workingRange.Cells(2, 1), workingRange.Cells(2, 2))
        .Merge
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Value = frmSettings.txtCourse.Value
    End With
    With Range(workingRange.Cells(3, 1), workingRange.Cells(3, 2))
        .Font.Bold = True
        .Font.Size = 12
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Select
    End With
    ColourGrey
    workingRange.Cells(3, 1).Value = "Unit"
    workingRange.Cells(3, 2).Value = "Group"
    With Range(workingRange.Cells(4, 1), workingRange.Cells(4, 2))
        .Font.Size = 12
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .ShrinkToFit = True
    End With
    workingRange.Cells(4, 1).Value = frmSettings.txtUnit.Value
    workingRange.Cells(4, 2).Value = frmSettings.txtGroup.Value
    addBorders workingRange
End Sub

Private Sub formatExtras(ByVal Students As Integer)
    With Range(Cells(4, 2), Cells(5, 3))
        .Merge
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Value = "Unit Tracking System"
    End With
    
    Dim workingRange As Range: Set workingRange = Range(Cells(10 + Students, 2), Cells(10 + Students + 2, 3))
    With workingRange
        .Merge
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Select
    End With
    If workingRange.Cells(1, 1).Value = "" Then workingRange.Value = "ENTER MESSAGE HERE!"
    AddThickBorders
End Sub

Private Sub formatKey(ByVal home As Range)
    With Range(home, home.Offset(2, 2))
        .Clear
        .ColumnWidth = 2.71
    End With
    With Range(home, home.Offset(2, 0))
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
    End With
    
    Dim sortType As Integer: sortType = variables.Cells(15, 2).Value
    If sortType > 3 Then Set home = home.Offset(0, 1)
    
    Dim workingRange As Range: Set workingRange = Range(home, home.Offset(2, 1))
    Range(home, home.Offset(2, 0)).Font.name = "Wingdings 2"
    workingRange.Cells(1, 1).Value = "R"
    workingRange.Cells(2, 1).Value = "8"
    workingRange.Cells(3, 1).Value = "T"
    workingRange.Cells(1, 2).Value = "Pass (R)"
    workingRange.Cells(2, 2).Value = "Referral (8)"
    workingRange.Cells(3, 2).Value = "Deadline Missed (T)"
    workingRange.Cells.Locked = True
End Sub
