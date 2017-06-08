Attribute VB_Name = "DashboardAdmin"
Option Explicit

' A Library of funtions and subroutines for use administering the Client Server Incident Dashboard
' Author: Marc Mitsialis
' Version: 0.0.3
' Date: 2017-03-06
' Changelog:
'   [2017-03-01]
'       New:
'            Increment_All()
'            Decrement_All()
'            PrepareForCapture()
'            RecordIncident()
'            Increment_Range()
'            Decrement_Range()
'   [2017-03-06]
'       New: Convert Increment_Range into IncrementDecrement_Range to handle arbitrary adjustments.
'       Update: Increment_All and Decrement_All to pass the adjustement value (handles weekends and public holidays.)
'       Delete: DecrementRange and IncrementRange
'       Update: Updated all subroutine for use with the RibbonBar
'   [2017-03-07]
'       Initial Release 0.0.3



Const strRngIncidentFreeDays As String = "$C$7:$D$36"
Dim docThisWorkbook As Workbook
Dim docThisWorkbookname As String
Dim shThisWorkSheet As Worksheet


'Callback for customButton304 onAction
Sub Increment_All(control As IRibbonControl, Optional intIncrementAmount As Integer = 1)

    ' Increment the 'Incident Free Days' for all services by one
    ' Use the range defined in the constant strRngIncidentFreeDays

    ' TODO: Change increment_Range to add an arbitrary number. handled gaps between evaluation dates

    Dim rngTargetToChange As Range
    Dim IncrementResult As Boolean

    Set docThisWorkbook = ActiveWorkbook
    Set rngTargetToChange = Range(strRngIncidentFreeDays)




    ' Increment IncidentFreeDays by "1"
    intIncrementAmount = 1 * intIncrementAmount
    IncrementResult = IncrementDecrement_Range(intIncrementAmount, rngTargetToChange)
    docThisWorkbook.Activate

    ActiveSheet.Calculate


End Sub

'Callback for customButton303 onAction
Sub Decrement_All(control As IRibbonControl, Optional intIncrementAmount As Integer = 1)

    ' Decrement the 'Incident Free Days' for all services by one
    ' Use the range defined in the constant strRngIncidentFreeDays
    ' Only to be used to 'fix' errors, by running from the Run Macro dialog.

    ' TODO: Change increment_Range to add an arbitrary number. handled gaps between evaluation dates

    Dim rngTargetToChange As Range
    Dim IncrementResult As Boolean

    Set docThisWorkbook = ActiveWorkbook
    Set rngTargetToChange = Range(strRngIncidentFreeDays)

    ' Decrement IncidentFreeDays by "1"
    intIncrementAmount = -1 * intIncrementAmount
    IncrementResult = IncrementDecrement_Range(intIncrementAmount, rngTargetToChange)
    ActiveSheet.Calculate

End Sub

'Callback for customButton300 onAction
Sub PrepareForCapture(control As IRibbonControl)
    ' Prepares the worksheet for data capture

    Dim dtToday As Date
    Dim dtStartOfYear As Date
    Dim dtDayToEvaluate As Date                  ' The day for which we are capturing data, 24hours ending 05:59
    Dim strEvaluationDate As String              ' Evaluation date as entered by user
    Dim arrEvaluationDate() As String            ' Array of yyyy, mm, dd as provided by user
    Dim dtRevisedEvaluationDate As Date          ' user provided evaluatation date
    Dim intDaysIntoYear As Integer               ' Calculated for the dtDayToEvaluate
    Dim intIncrementAmount As Integer            ' Days to adjust evaluation days by
    Dim iLetsGo As Integer
    Dim bdocCheckoutDone                         ' Flag if the document checkout is successful
    Dim rngTargetToChange As Range               ' Defines the range to update
    Dim IncrementResult As Boolean               ' Result of IncrementDecrement_Range()
    Dim intCurrentDaysIntoYear As Integer
    Dim dtCurrentDayToEvaluate As Date
    Dim bCheckDays As Boolean
    Dim bCheckDayToEvaluate As Boolean

    Set docThisWorkbook = ActiveWorkbook
    docThisWorkbookname = docThisWorkbook.FullName

    Worksheets("Integr8 Incident Dashboard").Activate
    Range("$A$1").Activate
    Range("$A$1").Select


    ' Set up the dates for the evaluation
    dtToday = Date
    dtStartOfYear = 1 * DateSerial(Year(dtToday), 1, 1)
    dtDayToEvaluate = Date
    intDaysIntoYear = DateDiff("y", dtStartOfYear, dtDayToEvaluate)
    intIncrementAmount = dtDayToEvaluate - CDate(Mid(ActiveWorkbook.Names("EvaluationDate"), 2))
    Set rngTargetToChange = Range(strRngIncidentFreeDays)


    ' Confirm information details and provide an opportunity to amend if necessary.
    intCurrentDaysIntoYear = CInt(Mid(ActiveWorkbook.Names("DaysIntoYear").RefersTo, 2))
    dtCurrentDayToEvaluate = CDate(Mid(ActiveWorkbook.Names("EvaluationDate"), 2))

    bCheckDays = intCurrentDaysIntoYear = intDaysIntoYear
    bCheckDayToEvaluate = dtCurrentDayToEvaluate = dtDayToEvaluate

    If bCheckDays And bCheckDayToEvaluate Then
        MsgBox "Dashboard has already evaluated and updated for today."
        Exit Sub
    End If


    iLetsGo = MsgBox( _
              "Evaluating the period ending """ & _
              FormatDateTime(dtDayToEvaluate, vbLongDate) & " 05:59""" & _
              vbCrLf & vbCrLf & _
              "Evaluated as """ & intDaysIntoYear & """ days into this year." & _
              vbCrLf & vbCrLf & _
              "The Incident Free days will be adjusted by """ & intIncrementAmount & """ days. " & _
              vbCrLf & vbCrLf & _
              "The document will be checked out for modifications. " & _
              vbCrLf & vbCrLf & _
              "When you are ready press ""Yes""?. " & vbCrLf & _
              "If you wish to adjust the evaluation date press ""No"".  " & vbCrLf & _
              "If you wish to cancel press ""Cancel"". ", _
              vbYesNoCancel, "Ready to Capture Incident Data")

    Select Case iLetsGo
        Case vbYes
            ' Go ahead and increment the incident free days.
            ' Increment the 'Incident Free Days' for all services
            ' TODO: Convert DoCheckOut to a function to allow error handling
            ' TODO: Introduce automatic checkout when figured out why current function is not working.
            ' DoCheckOut docCheckOut:=docThisWorkbookname, bMessage:=True

            ' Determine the difference between dtDayToEvaluate and ActiveWorkbook.Names("EvaluationDate")
            ' then increment by that amount.
            Set rngTargetToChange = Range(strRngIncidentFreeDays)
            IncrementResult = IncrementDecrement_Range(intIncrementAmount, rngTargetToChange)

            ' Update "Days into Year" count and "EvaluationDate" for the worksheet formulas and dashboard
            ActiveWorkbook.Names("DaysIntoYear").RefersTo = intDaysIntoYear
            ActiveWorkbook.Names("EvaluationDate").RefersTo = dtDayToEvaluate
            Range("$B$2").Value = intDaysIntoYear
            Range("$B$3").Value = FormatDateTime(dtDayToEvaluate, vbLongDate) & " 05:59"

        Case vbNo

            ' TODO: Add logic to loop until valid date value is provided
            ' Inputbox to present and collect the date of evaluation
            strEvaluationDate = InputBox( _
                                "Evaluating the period ending """ & _
                                FormatDateTime(dtDayToEvaluate, vbLongDate) & " 05:59""" & _
                                vbCrLf & vbCrLf & _
                                "Evaluated as """ & intDaysIntoYear & """ days into this year." & _
                                vbCrLf & vbCrLf & _
                                "Update the Evaluation Date if necessary. Use the format (yyyy-mm-dd)?", _
                                "Evaluation Date", _
                                CStr(Year(dtDayToEvaluate)) & "-" & CStr(Month(dtDayToEvaluate)) & "-" & CStr(Day(dtDayToEvaluate)))
            ' FormatDateTime(dtDayToEvaluate, vbShortDate))

            ' Set the evaluation date to calculated or user provided value.
            If strEvaluationDate = "" Then
                ' User pressed cancel or blanked the input field. Exit Sub
                MsgBox "You pressed cancel or emptied the input field. Please try again!"
            Else
                ' Confirm information details and provide an opportunity to amend if necessary.
                intCurrentDaysIntoYear = CInt(Mid(ActiveWorkbook.Names("DaysIntoYear").RefersTo, 2))
                dtCurrentDayToEvaluate = CDate(Mid(ActiveWorkbook.Names("EvaluationDate"), 2))

                bCheckDays = intCurrentDaysIntoYear = intDaysIntoYear
                bCheckDayToEvaluate = dtCurrentDayToEvaluate = dtDayToEvaluate

                If bCheckDays And bCheckDayToEvaluate Then
                    MsgBox "Dashboard has already evaluated and updated for today."
                    Exit Sub
                End If

                ' TODO: Convert DoCheckOut to a function to allow error handling
                ' TODO: Introduce automatic checkout when figured out why current function is not working.
                ' DoCheckOut docCheckOut:=docThisWorkbookname, bMessage:=True

                ' Determine the difference between dtDayToEvaluate and ActiveWorkbook.Names("EvaluationDate")
                ' then increment by that amount.
                ' Change the adjustments based on new input date.
                arrEvaluationDate() = Split(strEvaluationDate, "-", 3, vbTextCompare)
                dtRevisedEvaluationDate = DateSerial(CInt(arrEvaluationDate(0)), CInt(arrEvaluationDate(1)), CInt(arrEvaluationDate(2)))
                dtDayToEvaluate = dtRevisedEvaluationDate
                intDaysIntoYear = DateDiff("y", dtStartOfYear, dtDayToEvaluate)
                intIncrementAmount = dtDayToEvaluate - CDate(Mid(ActiveWorkbook.Names("EvaluationDate"), 2))

                ' Increment the 'Incident Free Days' for all services
                IncrementResult = IncrementDecrement_Range(intIncrementAmount, rngTargetToChange)
                ' Update "Days into Year" count for the worksheet formulas
                ActiveWorkbook.Names("DaysIntoYear").RefersTo = intDaysIntoYear
                ' Update the "Days into Year" in the dashboard
                Range("$B$2").Value = intDaysIntoYear
                ' Update the "Evaluation up to" in the dashboard
                Range("$B$3").Value = FormatDateTime(dtDayToEvaluate, vbLongDate) & " 05:59"
                ActiveSheet.Calculate

            End If

        Case vbCancel

        Case Else

    End Select

End Sub

'Callback for customButton301 onAction
Sub AllDone(control As IRibbonControl)
    ' All done, check in the document and close

    Set docThisWorkbook = ActiveWorkbook
    docThisWorkbook.CheckIn SaveChanges:=True, Comments:="Incident free days updated at " & Now(), MakePublic:=True

End Sub

'Callback for customButton302 onAction
Sub RecordIncident(control As IRibbonControl)

    Dim iSelectedServiceColumn                   ' Column number of the active cell to find "Service" and update incident count.
    Dim iSelectedServiceRow                      ' Row number of the active cell to find "Service" and update incident count.
    Dim strService                               ' The determined "Service"
    Dim strIncidentSeverity As String            'string indicating incident severity for messages
    Dim iRecordIncident As Integer               ' Result of question to record incident

    If Intersect(ActiveCell, Range(strRngIncidentFreeDays)) Is Nothing Then
        MsgBox "Please select a cell in the range " & _
               strRngIncidentFreeDays & " to record an incident, then try again.  "
        Range(strRngIncidentFreeDays).Select
        Exit Sub
    Else
        iSelectedServiceColumn = ActiveCell.Column
        iSelectedServiceRow = ActiveCell.Row

        Select Case iSelectedServiceColumn
            Case 3
                strIncidentSeverity = "SEV1"
                strService = ActiveCell.Offset(0, -1).Value
            Case 4
                strIncidentSeverity = "SEV1"
                strService = ActiveCell.Offset(0, -2).Value
        End Select

        ' Confirm we will record the incident or cancel
        iRecordIncident = MsgBox( _
                          "You are about to record a """ & strIncidentSeverity & """ for the service """ & _
                          strService & """ in the 24 hour period ending """ & _
                          FormatDateTime(CDate(Mid(ActiveWorkbook.Names("EvaluationDate"), 2)), vbLongDate) & _
                          " 05:59""" & _
                          vbCrLf & vbCrLf & _
                          "If the fault occured in the " & strService & " service, select ""Yes"". " & _
                          vbCrLf & vbCrLf & _
                          "If the " & strService & " service, was affected by an external fault, select ""No"". " & _
                          vbCrLf & vbCrLf & _
                          " Please select Cancel if you do not want to record an incident ", _
                          vbYesNoCancel, "Record an Incident")

        Select Case iRecordIncident
            Case vbYes                           ' Record an incident for which we are responsible
                Select Case iSelectedServiceColumn
                    Case 3
                        ' Record a SEV1 incident
                        ActiveCell.Offset(0, 5).Value = ActiveCell.Offset(0, 5).Value + 1
                        ActiveCell.Value = 0
                        ActiveSheet.Calculate

                    Case 4
                        ' Record a SEV2 incident
                        ActiveCell.Offset(0, 4).Value = ActiveCell.Offset(0, 4).Value + 1
                        ActiveCell.Value = 0
                        ActiveSheet.Calculate

                End Select

            Case vbNo                            ' Record an incident for which we are NOT responsible
                Select Case iSelectedServiceColumn
                    Case 3
                        ' Record a SEV1 incident
                        ActiveCell.Offset(0, 6).Value = ActiveCell.Offset(0, 6).Value + 1
                        ActiveCell.Value = 0
                        ActiveSheet.Calculate

                    Case 4
                        ' Record a SEV2 incident
                        ActiveCell.Offset(0, 5).Value = ActiveCell.Offset(0, 5).Value + 1
                        ActiveCell.Value = 0
                        ActiveSheet.Calculate

                End Select

            Case vbCancel                        ' Abandon recording an incident.
                Exit Sub

        End Select

    End If


End Sub

Function IncrementDecrement_Range(intIncrementAmount As Integer, Optional ByVal rngTargetToChange As Range)
    ' Increment the value of each cell in a range by one.
    ' Range may be "1 x 1" or "p x q" in dimension

    Dim rngTargetCell As Range
    ' Set RangeTochange = Range("C6:D35")
    For Each rngTargetCell In rngTargetToChange
        rngTargetCell.Value = rngTargetCell.Value + intIncrementAmount
    Next rngTargetCell



End Function


