Attribute VB_Name = "Students"
''  Program: Unit Tracking System (UTS)
''  Created by Simon Peter Campbell
''
''  The purpose of UTS is to provide an automated tracking system for individual BTEC units.
''  The aim is to create a fully functional and bug free system that is scalable in terms
''  of different amounts of criteria, students & sorting preferences.
''
''  This module contains the code involved in the addition and removal of students from
''  the tracker. First the current amount of students is determined, then if changes are
''  required to meet the users requirements they are performed, otherwise nothing happens.
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
''  1) Fully test the module

Sub DoStudents()
    '' This is the stage where the current students on the tracker are determined.
    Dim changesRequired As Boolean
    Dim curStudents As Integer: curStudents = 0
    Dim newStudents As Integer: newStudents = frmSettings.numStudents.Value
    
    '' A loop is run to find out exactly how many students are still present.
    Range("B" & (9 + curStudents)).Select
    If Selection.Cells(1, 1).Value <> "end" Then
        Do Until Selection.Value = "end"
            curStudents = curStudents + 1
            Range("B" & (9 + curStudents)).Select
        Loop
    '' else curStudents will remain 0.
    End If
    addStudents curStudents, newStudents
    frmSettings.numStudents.Value = newStudents
    Columns("A:A").ColumnWidth = 2.29
    Columns("B:B").ColumnWidth = 26.71
    Columns("C:C").ColumnWidth = 17.57
    Columns("D:D").ColumnWidth = 2.29
    Rows("1:" & 8 + newStudents).RowHeight = 15.75
End Sub

Private Sub addStudents(curStudents As Integer, newStudents As Integer)
    '' Now the system finds out if students need to be added or subtracted.
    Dim count As Integer: count = 0
    Range("B" & (9 + curStudents)).Select
    Select Case newStudents
    Case Is > curStudents
        For i = 1 To (newStudents - curStudents)
            Selection.EntireRow.Insert shift:=xlUp
            Selection.EntireRow.ClearFormats
        Next i
    Case Is < curStudents
        Do
            Range("B" & 9 + newStudents).Select
            Selection.EntireRow.Delete
            count = count - 1
        Loop Until count = newStudents - curStudents 'Selection.Value = "end"
    End Select
    '' Polish the tracker off
    addBorders Range("B9", "D" & (8 + newStudents))
    Range("D9", "D" & (8 + newStudents)).Select
    ColourGrey
End Sub

