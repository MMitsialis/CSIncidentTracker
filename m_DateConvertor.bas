Attribute VB_Name = "m_DateConvertor"
Option Explicit

Private mStrLastPattern As String
Private mStrSourceDate As String
Private mDatResult As Date

Public Function fctDateFromString(strDate As String) As Date
    mStrSourceDate = strDate
    mDatResult = 0
    If TryConvert("(^\d{2})\.(\d{2})\.(\d{4})$", "$2/$1/$3") Then 'DD.MM.YYYY
    ElseIf TryConvert("(^\d{2})\.(\d{2})\.(\d{2})$", "$2/$1/20$3") Then 'DD.MM.YY
    ElseIf TryConvert("(^\d{4})(\d{2})\.(\d{2})$", "$2/$3/$1") Then 'YYYYMMDD
    ElseIf TryConvert("(^\d{2})/(\d{2})/(\d{4})$", "$1/$2/$3") Then 'MM/DD/YYYY
    ElseIf TryConvert("(^\d{2})/(\d{2})/(\d{2})$", "$1/$2/20$3") Then 'MM/DD/YY
    ElseIf TryConvert("(^\d{1})/(\d{1})/(\d{4})$", "0$1/0$2/$3") Then 'M/D/YYYY
    ElseIf TryConvert("(^\d{1})/(\d{1})/(\d{2})$", "0$1/0$2/20$3") Then 'M/D/YY
    End If
    If mDatResult = 0 Then Debug.Print "Cannot find matching format for " & strDate
    fctDateFromString = mDatResult
End Function

Private Function TryConvert(strFrom As String, strTo As String) As Boolean
    If RegExMatch(strFrom) Then
        mDatResult = RegExConvert("$1/$2/$3")
        TryConvert = (mDatResult <> 0)
    End If
End Function

Private Function RegExMatch(strPattern As String) As Boolean
    mStrLastPattern = strPattern
    With CreateObject("VBScript.RegExp")
        .Pattern = strPattern
        .IgnoreCase = True
        .MultiLine = False
        RegExMatch = .Test(mStrSourceDate)
    End With
End Function

Private Function RegExConvert(strReplacePattern As String) As Date
    On Error Resume Next
    With CreateObject("VBScript.RegExp")
        .Pattern = mStrLastPattern
        .IgnoreCase = True
        .MultiLine = False
        RegExConvert = CDate(.Replace(mStrSourceDate, strReplacePattern))
        If Err.Number Then
            Err.Clear
            RegExConvert = 0
        End If
    End With

End Function

