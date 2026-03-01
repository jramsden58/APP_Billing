Attribute VB_Name = "modHelpers"
'==============================================================================
' modHelpers - Utility Helpers Module
' APP Billing System
'
' Contains shared helper routines used by forms and other modules.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' UnloadSuperUser - Deferred unload for frmSuperUser when auth fails
' Called via Application.OnTime from UserForm_Initialize
'------------------------------------------------------------------------------
Public Sub UnloadSuperUser()
    On Error Resume Next
    Unload frmSuperUser
End Sub

'------------------------------------------------------------------------------
' Show_DailyExport - Opens the Daily Export form
' Can be assigned to a button on the Home sheet
'------------------------------------------------------------------------------
Public Sub Show_DailyExport()
    frmDailyExport.Show
End Sub

'------------------------------------------------------------------------------
' ParseDateDMY - Parses a DD/MM/YYYY date string regardless of system locale
'
' Always interprets as DD/MM/YYYY. Never uses IsDate/CDate which depend on
' the Windows locale setting.
'
' Returns: Parsed Date value
' Raises Error 13 if format is invalid
'------------------------------------------------------------------------------
Public Function ParseDateDMY(ByVal sDate As String) As Date
    Dim parts() As String
    sDate = Trim(sDate)

    If Len(sDate) = 0 Then
        Err.Raise 13, "ParseDateDMY", "Date string is empty."
    End If

    parts = Split(sDate, "/")
    If UBound(parts) <> 2 Then
        Err.Raise 13, "ParseDateDMY", "Invalid date format. Use DD/MM/YYYY."
    End If

    Dim iDay As Long, iMonth As Long, iYear As Long
    On Error GoTo InvalidParts
    iDay = CInt(parts(0))
    iMonth = CInt(parts(1))
    iYear = CInt(parts(2))
    On Error GoTo 0

    ' Basic range validation
    If iMonth < 1 Or iMonth > 12 Then
        Err.Raise 13, "ParseDateDMY", "Month must be between 1 and 12."
    End If
    If iDay < 1 Or iDay > 31 Then
        Err.Raise 13, "ParseDateDMY", "Day must be between 1 and 31."
    End If
    If iYear < 1900 Or iYear > 2100 Then
        Err.Raise 13, "ParseDateDMY", "Year must be between 1900 and 2100."
    End If

    ' Use DateSerial which handles month-end overflow (e.g., Feb 30 -> Mar 2)
    Dim dtResult As Date
    dtResult = DateSerial(iYear, iMonth, iDay)

    ' Verify the date didn't overflow (e.g., Feb 30 becoming Mar 2)
    If Day(dtResult) <> iDay Or Month(dtResult) <> iMonth Then
        Err.Raise 13, "ParseDateDMY", "Invalid date: " & sDate
    End If

    ParseDateDMY = dtResult
    Exit Function

InvalidParts:
    Err.Raise 13, "ParseDateDMY", "Date contains non-numeric parts: " & sDate
End Function

'------------------------------------------------------------------------------
' TryParseDateDMY - Attempts to parse DD/MM/YYYY, returns success/failure
'
' Parameters:
'   sDate    - The date string to parse
'   dtResult - [out] The parsed date if successful
'
' Returns: True on success, False on failure
'------------------------------------------------------------------------------
Public Function TryParseDateDMY(ByVal sDate As String, ByRef dtResult As Date) As Boolean
    On Error GoTo ParseFailed
    dtResult = ParseDateDMY(sDate)
    TryParseDateDMY = True
    Exit Function

ParseFailed:
    TryParseDateDMY = False
End Function

'------------------------------------------------------------------------------
' IsValidDateDMY - Checks if a string is a valid DD/MM/YYYY date
'------------------------------------------------------------------------------
Public Function IsValidDateDMY(ByVal sDate As String) As Boolean
    Dim dtDummy As Date
    IsValidDateDMY = TryParseDateDMY(sDate, dtDummy)
End Function

'------------------------------------------------------------------------------
' IsValidTime24 - Checks if a string is a valid HH:MM time (24-hour)
'------------------------------------------------------------------------------
Public Function IsValidTime24(ByVal sTime As String) As Boolean
    On Error GoTo InvalidTime

    sTime = Trim(sTime)
    If Len(sTime) = 0 Then
        IsValidTime24 = False
        Exit Function
    End If

    Dim parts() As String
    parts = Split(sTime, ":")
    If UBound(parts) <> 1 Then
        IsValidTime24 = False
        Exit Function
    End If

    Dim iHour As Long, iMin As Long
    iHour = CInt(parts(0))
    iMin = CInt(parts(1))

    IsValidTime24 = (iHour >= 0 And iHour <= 23 And iMin >= 0 And iMin <= 59)
    Exit Function

InvalidTime:
    IsValidTime24 = False
End Function

'------------------------------------------------------------------------------
' FormatTimestamp - Returns a properly formatted DD/MM/YYYY HH:nn:SS timestamp
' VBA's Format "MM" always returns month; use "nn" for minutes.
'------------------------------------------------------------------------------
Public Function FormatTimestamp(ByVal dtValue As Date) As String
    FormatTimestamp = Format(dtValue, "DD/MM/YYYY HH:nn:SS")
End Function

'------------------------------------------------------------------------------
' EnsureSheetExists - Creates a worksheet if it doesn't exist
'
' Returns: The worksheet (existing or newly created)
'------------------------------------------------------------------------------
Public Function EnsureSheetExists(ByVal sName As String) As Worksheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sName
    End If

    Set EnsureSheetExists = ws
End Function
