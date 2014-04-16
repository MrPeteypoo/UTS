VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Unit1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''  Program: Unit Tracking System (UTS)
''  Created by Simon Peter Campbell
''
''  The purpose of UTS is to provide an automated tracking system for individual BTEC units.
''  The aim is to create a fully functional and bug free system that is scalable in terms
''  of different amounts of criteria, students & sorting preferences.
''
''  This particular part of the program deals with the automation of the colouring of cells,
''  the detection of grade changes and finally the sorting of the students.
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
''  1) Fully test the rewritten code - In Progress

Private Sub Worksheet_Change(ByVal target As Range)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    variables.Visible = xlSheetVisible
    
    Dim prevCell As Range:      Set prevCell = Selection
    Dim Assignments As Range:   Set Assignments = Range(variables.Cells(11, 2).Value)
    Dim criteria As Range:      Set criteria = Range(variables.Cells(12, 2).Value)
    
    If Not Intersect(target, Assignments) Is Nothing Then
        AssignmentNumbers Assignments
    '' If the assignment numbers have been modified, change the colours respectively.
    ElseIf Not Intersect(target, criteria) Is Nothing Then
        CriteriaChanged target
    '' If any criteria has been modified then change the colours respectively.
    End If
    
    SortAlgorithm
    prevCell.Select
    variables.Visible = xlSheetHidden
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub AssignmentNumbers(ByVal Assignments As Range)
'' This section of code changes the colours of every assignment number and criteria that resides
'' below it, it does this by looping through the entire assignment number row so one edit refreshes
'' the entire row.
    Dim Students As Integer: Students = variables.Cells(6, 2).Value
    Dim curCell As Range
    For Each curCell In Assignments '' A loop will run through the entire assignment number range.
        If curCell.row = 7 Then
            Range(curCell, curCell.Offset(1, 0)).Select '' Select both the assignment number cell and the adjacent criteria cell.
            Select Case curCell.Value
            Case "1"
                ColourAss1
                curCell.Offset(Students + 2, 0).Select
                ColourAss1
            Case "2"
                ColourAss2
                curCell.Offset(Students + 2, 0).Select
                ColourAss2
            Case "3"
                ColourAss3
                curCell.Offset(Students + 2, 0).Select
                ColourAss3
            Case "4"
                ColourAss4
                curCell.Offset(Students + 2, 0).Select
                ColourAss4
            Case "5"
                ColourAss5
                curCell.Offset(Students + 2, 0).Select
                ColourAss5
            Case Else
                ColourRemove
                curCell.Offset(Students + 2, 0).Select
                ColourRemove
            End Select
        End If
    Next curCell
End Sub

Private Sub CriteriaChanged(ByVal target As Range)
'' This procedure is dedicated to changing the colours of cells in the criteria
'' section to their respective colour. First the procedure tests if an entire row
'' has been selected for deletion, this is done to avoid lengthy wait times because
'' every cell being deleted gets tested for different values.
''
'' When ready the system will loop through every modified cell and test their values
'' and colour in the cell accordingly. Once all cells have been coloured the students
'' row is changed in the event of a grade being met.
    '' pre-condition: A row hasn't been selected for deletion.
    If target.Cells.count < 2000 Then
        Dim curCell As Range
        Dim result As Integer
        For Each curCell In target
            result = 0
            Select Case curCell.Value
            Case "R"
                determineCriteria result, curCell
                Select Case result
                Case 1
                    curCell.Select
                    ColourPass
                Case 2
                    curCell.Select
                    ColourMerit
                Case 3
                    curCell.Select
                    ColourDistinction
                End Select
            Case "8"
                determineCriteria result, curCell
                Select Case result
                Case 1
                    curCell.Select
                    ColourPassR
                Case 2
                    curCell.Select
                    ColourMeritR
                Case 3
                    curCell.Select
                    ColourDistinctionR
                End Select
            Case "T"
                curCell.Select
                ColourRisk
            Case Else
                curCell.Select
                ColourRemove
            End Select
            StudentColour curCell.row, target
        Next curCell
    End If
End Sub

Private Sub SortAlgorithm()
'' Simple sorting algorithm, the number of students and criteria is obtained then the
'' sort type is determined. Once this is done the table is sorted accordingly.
    Dim Students As Integer: Students = variables.Cells(6, 2).Value
    Dim criteria As Integer: criteria = variables.Cells(7, 2).Value + variables.Cells(8, 2).Value + variables.Cells(9, 2).Value
    Dim sortType As Integer: sortType = variables.Cells(15, 2).Value
    Dim SortRange As Range: Set SortRange = Range("B9", Cells(8 + Students, 4 + criteria + 3))
    
    Select Case sortType
    Case 1
        SortAlphabet SortRange
    Case 2
        SortGrade SortRange
    Case 3 '' Same as case 2 because the way the points are accumlated is what makes the sort types different.
        SortGrade SortRange
    End Select
End Sub

