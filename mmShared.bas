Attribute VB_Name = "mmShared"
Option Explicit

' A Library of funtions and subroutines for use with SharePoint
' Author: Marc Mitsialis
' Version: 1.0.1
' Date: 2017-03-01
' Changelog:
'   [2017-03-01]
'       Fix: Added "Option Explicit"
'       Fix: Cleaned undefined variables.
'       Add: added more comments.
'       Update: added more descriptive information in message boxes.




Sub DoCheckIn(docCheckIn As String, Optional strCommitMessage As String, Optional bMakePublic As Boolean = False, Optional bMessage As Boolean)
    ' docCheckIn is the pathname of the workbook to check in
    ' strCommitMessage is the text to be inserted in the check in comments
    ' bMakePublic is a flag to choose the "Keep document checked out" choice in the CheckIn dialog
    ' bMessage is a flag to provide a message to the user
    
    ' Determine if workbook can be checked in.
    If Workbooks(docCheckIn).CanCheckIn = True Then
        Select Case bMessage
            Case True
                Workbooks(docCheckIn).CheckIn SaveChanges:=True, Comments:=strCommitMessage, MakePublic:=True
                If strCommitMessage = "" Then
                    MsgBox docCheckIn & " has been checked in. "
                Else
                    MsgBox docCheckIn & " has been checked in. Commit message was: " & strCommitMessage
                End If
            Case False
                Workbooks(docCheckIn).CheckIn SaveChanges:=True, Comments:=strCommitMessage, MakePublic:=True
        End Select
    Else
        MsgBox "This file cannot be checked in at this time. Please try again later." & vbCrLf & _
               "Check in the SharePoint site an ensure you have it checked out. " & vbCrLf & _
               "If checked out to another user, ask them to cancel their checkout. Then  " & vbCrLf & _
               "check it out to yourself."
    End If
End Sub

Sub DoCheckOut(docCheckOut As String, Optional bMessage As Boolean = False)
    ' docCheckIn is the pathname of the workbook to check in
    ' bMessage is a flag to provide a message to the user
    
    ' Determine if workbook can be checked out.
    If Workbooks.CanCheckOut(docCheckOut) = True Then
        Select Case bMessage
            Case True
                Workbooks.CheckOut docCheckOut
                MsgBox docCheckOut & " has been checked out."
        
            Case False
                Workbooks.CheckOut docCheckOut
        End Select
    Else
        MsgBox "This file could not be checked out." & vbCrLf & vbCrLf & _
               "Review the SharePoint site an ensure the document ( " & _
               docCheckOut & " ) is not checked out to another user." & vbCrLf & vbCrLf & _
               "If checked out to another user, ask them to cancel their checkout." & vbCrLf & _
               "You may manually check out the document in SharePoint. "
    End If
 
End Sub

Sub ConvertAllToValues()
    '
    'Originally Adapted from OZgrid.com
    ' http://www.ozgrid.com/forum/showthread.php?t=38064
    '
    Dim OldSelection As Range
    Dim HiddenSheets() As Boolean
    Dim Goahead As Integer, n As Integer, i As Integer
    Goahead = MsgBox("This will irreversibly convert all formulas in the workbook to values. Continue?", vbOKCancel, "Confirm conversion to values only")
    
    ' ThisWorkbook.VBProject.Name
        
    If Goahead = vbOK Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
         
        n = Sheets.Count
        ReDim HiddenSheets(1 To n) As Boolean
         
        For i = 1 To n
            If Sheets(i).Visible = False Then HiddenSheets(i) = True
            Sheets(i).Visible = True
        Next
         
        Set OldSelection = Selection.Cells
        Worksheets.Select
        Cells.Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues
         
        Cells(OldSelection.Row, OldSelection.Column).Select
        Sheets(OldSelection.Worksheet.Name).Select
         
        Application.CutCopyMode = False
         
        For i = 1 To n
            Sheets(i).Visible = Not HiddenSheets(i)
        Next
         
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub


