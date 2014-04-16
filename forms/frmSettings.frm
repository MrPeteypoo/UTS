VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Tracker Settings"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''  Program: Unit Tracking System (UTS)
''  Created by Simon Peter Campbell
''
''  The purpose of UTS is to provide an automated tracking system for individual BTEC units.
''  The aim is to create a fully functional and bug free system that is scalable in terms
''  of different amounts of criteria, students & sorting preferences.
''
''  This forms objective is to allow the user to customise the tracker to suit their needs.
''  The options given on the form effect the way the tracker works, the code contained within
''  handles these different events.
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

Private Sub UserForm_Initialize()
    '' Add items to the controls.
    '' The maximum values for each option are set here.
    For i = 1 To 30
        numStudents.AddItem (i)
    Next i
    For i = 1 To 11
        numPass.AddItem (i)
    Next i
    For i = 1 To 6
        numMerit.AddItem (i)
    Next i
    For i = 1 To 4
        numDistinction.AddItem (i)
    Next i
    '' Set the initial values of each control.
    defaultOptions
End Sub

Private Sub UserForm_Activate()
    '' Set the values of each control
    defaultOptions
End Sub

Private Sub cmdApply_Click()
    '' Confirm whether they wish to continue.
    Dim confirmation As Integer: confirmation = MsgBox("Are you sure you wish to continue? The tracker contents will be changed, if you have removed any students or criteria then it will be lost forever.", vbYesNo, "Confirmation")
    If confirmation = vbYes Then
    '' Start with a little housekeeping
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        unlockSheets
        Unit1.Cells.UnMerge
        variables.Visible = xlSheetVisible
        help.Visible = xlSheetVisible
            
        '' Next the sort type must be set before anything else.
        If frmSettings.radAlphabet.Value = True Then
            variables.Cells("15", "B").Value = 1
        ElseIf frmSettings.radGrade.Value = True Then
            variables.Cells("15", "B").Value = 2
        ElseIf frmSettings.radLeader.Value = True Then
            variables.Cells("15", "B").Value = 3
        End If
        
        '' After that check to see if students, criteria and formatting need adding.
        DoStudents
        DoCriteria
        ''Store variables in second sheet.
        writeVariables
        DoExtras
        doSecurity
        
        Dim temp As Integer: temp = Cells(7, 5).Value
        Cells(7, 5).ClearContents
        Application.EnableEvents = True
        Cells(7, 5).Value = temp
        Cells(2, 2).Select
    Else ''Show that the operation has been cancelled.
        MsgBox "Operation cancelled. Nothing was changed.", vbOKOnly, "Cancelled"
    End If
End Sub

Private Sub cmdWithdraw_Click()
    '' This button tells the tracker to delete all students that have withdrawn.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    variables.Visible = xlSheetVisible
    unlockSheets
    Dim confirmation As Integer: confirmation = MsgBox("Are you sure you want to delete withdrawn students from the tracker? This cannot be reversed.", vbYesNo, "Caution")
    If confirmation = vbNo Then
        MsgBox "Operation cancelled. Nothing was changed.", vbOKOnly, "Cancelled"
    ElseIf confirmation = vbYes Then
        '' Select D8 as the starting point and look for the grade column.
        Cells(8, 4).Select
        For i = 1 To 100 '' 100 being a safe maximum.
            Selection.Offset(0, 1).Select
            If Selection.Cells(1, 1).Value = "Overall Grade" Then Exit For
        Next i
        Selection.Offset(0, 1).Select
        
        '' Now go through each student and test if they should be deleted.
        Dim Students As Integer: Students = variables.Cells(6, 2).Value
        For i = 1 To Students
            Selection.Offset(1, 0).Select
            If Selection.Value = "Withdrawn" Or Selection.Value = "withdrawn" Then
                Students = Students - 1
                Selection.EntireRow.Delete
                Selection.Offset(-1, 0).Select
            End If
        Next i
        If Students = variables.Cells(6, 2).Value Then _
            MsgBox "No students were removed. Make sure they have ""Withdrawn"" in their notes cell.", vbOKOnly, "Error"
        variables.Cells(6, 2).Value = Students
        defaultOptions
    End If
    writeVariables
    doSecurity
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Private Sub cmdCancel_Click()
    Me.Hide '' Hide the window when the user clicks cancel.
End Sub

Private Sub defaultOptions()
    '' The currently selected values of the tracker are stored on a separate sheet.
    '' This procedure obtains those and mirrors the controls to reflect previous choices.
    variables.Activate
    numStudents.Value = Cells(6, 2).Value
    numPass.Value = Cells(7, 2).Value
    numMerit.Value = Cells(8, 2).Value
    numDistinction.Value = Cells(9, 2).Value
    If Cells(15, 2).Value = 1 Then
        radAlphabet.Value = True
    ElseIf Cells(15, 2).Value = 2 Then
        radGrade.Value = True
    ElseIf Cells(15, 2).Value = 3 Then
        radLeader.Value = True
    End If
    txtCourse.Value = Cells(16, 2).Value
    txtUnit.Value = Cells(17, 2).Value
    txtGroup.Value = Cells(18, 2).Value
    Unit1.Activate
End Sub


