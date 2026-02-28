Attribute VB_Name = "modNetworkIO"
'==============================================================================
' modNetworkIO - Network File I/O Module
' APP Billing System
'
' Handles reading and writing billing data to per-user daily Excel files
' on the network share. Uses per-user files to avoid lock conflicts with
' 15+ concurrent users.
'==============================================================================
Option Explicit

' Column headers for the daily data files (must match DailyDatabase)
Private Const NUM_COLUMNS As Long = 28 ' A through AB (added SyncStatus)

' DailyDatabase column indices
Public Const COL_SERIAL As Long = 1
Public Const COL_ANESTH As Long = 2
Public Const COL_SITE As Long = 3
Public Const COL_DATE As Long = 4
Public Const COL_SHIFT As Long = 5
Public Const COL_ONCALL As Long = 6
Public Const COL_SHIFTTYPE As Long = 7
Public Const COL_PROCCODE As Long = 8
Public Const COL_STARTTIME As Long = 9
Public Const COL_FINTIME As Long = 10
Public Const COL_MAXIC As Long = 11
Public Const COL_CONSULT As Long = 12
Public Const COL_MOD1 As Long = 13
Public Const COL_MOD2 As Long = 14
Public Const COL_MOD3 As Long = 15
Public Const COL_RESUS As Long = 16
Public Const COL_OBS As Long = 17
Public Const COL_ACUTEPAIN As Long = 18
Public Const COL_CHRONPAIN As Long = 19
Public Const COL_MISC As Long = 20
Public Const COL_WCBNUM As Long = 21
Public Const COL_WCBSIDE As Long = 22
Public Const COL_WCBDIAG As Long = 23
Public Const COL_WCBINJ As Long = 24
Public Const COL_WCBDATE As Long = 25
Public Const COL_SUBMBY As Long = 26
Public Const COL_SUBMON As Long = 27
Public Const COL_SYNCSTATUS As Long = 28

'------------------------------------------------------------------------------
' GetHeaders - Returns an array of column headers
'------------------------------------------------------------------------------
Private Function GetHeaders() As Variant
    GetHeaders = Array("S #", "Anesthesiologist", "Site", "Date of Service", _
                       "Shift Name", "On Call", "Shift Type", "Surgical Procedure Code", _
                       "Procedure Start Time", "Procedure Finish Time", "Maximum IC Level", _
                       "Consults", "Fee Modifier 1", "Fee Modifier 2", "Fee Modifier 3", _
                       "Resuscitation", "Obstetrics", "Acute Pain", _
                       "Diagnostic and Chronic Pain", "Miscellaneous Fee Items", _
                       "WCB Number", "Side", "Diagnostic Code", "Injury Code", _
                       "Date of Injury", "Submitted By", "Submitted On", "Sync Status")
End Function

'------------------------------------------------------------------------------
' SaveToNetwork - Saves a single row of data to the user's daily network file
'
' Parameters:
'   wsSource - The DailyDatabase worksheet
'   lRow     - The row number in DailyDatabase to save
'
' Returns True on success, False on failure
'------------------------------------------------------------------------------
Public Function SaveToNetwork(ByVal wsSource As Worksheet, ByVal lRow As Long) As Boolean
    On Error GoTo ErrHandler

    ' Get the anesthesiologist name and date for file naming
    Dim sAnesth As String
    sAnesth = CStr(wsSource.Cells(lRow, COL_ANESTH).Value)

    Dim dtService As Date
    Dim sDateVal As String
    sDateVal = CStr(wsSource.Cells(lRow, COL_DATE).Value)

    ' Try to parse the date
    If IsDate(sDateVal) Then
        dtService = CDate(sDateVal)
    Else
        ' Try DD/MM/YYYY format
        Dim parts() As String
        parts = Split(sDateVal, "/")
        If UBound(parts) = 2 Then
            dtService = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
        Else
            dtService = Date ' Fallback to today
        End If
    End If

    ' Build file path
    Dim sFilePath As String
    sFilePath = GetUserDailyFilePath(sAnesth, dtService)

    If Len(sFilePath) = 0 Then
        SaveToNetwork = False
        Exit Function
    End If

    ' Retry logic - 3 attempts with delay
    Dim iAttempt As Long
    For iAttempt = 1 To 3
        If TrySaveToFile(wsSource, lRow, sFilePath) Then
            ' Mark as synced in the DailyDatabase
            wsSource.Cells(lRow, COL_SYNCSTATUS).Value = "Synced"
            SaveToNetwork = True
            Exit Function
        End If

        ' Wait before retry
        If iAttempt < 3 Then
            Application.Wait Now + TimeSerial(0, 0, 2)
        End If
    Next iAttempt

    ' All retries failed
    wsSource.Cells(lRow, COL_SYNCSTATUS).Value = "Pending"
    SaveToNetwork = False
    Exit Function

ErrHandler:
    wsSource.Cells(lRow, COL_SYNCSTATUS).Value = "Error: " & Err.Description
    SaveToNetwork = False
End Function

'------------------------------------------------------------------------------
' TrySaveToFile - Attempts a single save operation to the network file
'------------------------------------------------------------------------------
Private Function TrySaveToFile(ByVal wsSource As Worksheet, ByVal lRow As Long, _
                                ByVal sFilePath As String) As Boolean
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lNextRow As Long
    Dim bNewFile As Boolean

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Check if file exists
    If Dir(sFilePath) <> "" Then
        ' Open existing file
        Set wb = Workbooks.Open(sFilePath, UpdateLinks:=0, ReadOnly:=False)
        Set ws = wb.Sheets(1)
        lNextRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row + 1
        bNewFile = False
    Else
        ' Create new file
        Set wb = Workbooks.Add(xlWBATWorksheet)
        Set ws = wb.Sheets(1)
        ws.Name = "DailyData"

        ' Write headers
        Dim headers As Variant
        headers = GetHeaders()
        Dim j As Long
        For j = 0 To UBound(headers)
            ws.Cells(1, j + 1).Value = headers(j)
        Next j

        ' Format headers
        With ws.Range(ws.Cells(1, 1), ws.Cells(1, NUM_COLUMNS))
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With

        lNextRow = 2
        bNewFile = True
    End If

    ' Copy row data (columns A through AA from source, skip AB=SyncStatus)
    Dim i As Long
    For i = 1 To COL_SUBMON ' Columns 1 through 27
        ws.Cells(lNextRow, i).Value = wsSource.Cells(lRow, i).Value
    Next i

    ' Set the serial number for the network file
    ws.Cells(lNextRow, COL_SERIAL).Value = lNextRow - 1

    ' Set sync status in network file
    ws.Cells(lNextRow, COL_SYNCSTATUS).Value = "Synced"

    ' Auto-fit columns on new files
    If bNewFile Then
        ws.Columns.AutoFit
    End If

    ' Save and close quickly
    If bNewFile Then
        wb.SaveAs sFilePath, FileFormat:=xlOpenXMLWorkbook
    Else
        wb.Save
    End If
    wb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    TrySaveToFile = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    TrySaveToFile = False
End Function

'------------------------------------------------------------------------------
' ReadUserDailyData - Reads a user's daily data file into a 2D array
'
' Returns: Variant array with data, or Empty if file doesn't exist/error
'------------------------------------------------------------------------------
Public Function ReadUserDailyData(ByVal sUserName As String, _
                                   ByVal dtDate As Date) As Variant
    On Error GoTo ErrHandler

    Dim sFilePath As String
    sFilePath = GetUserDailyFilePath(sUserName, dtDate)

    If Len(sFilePath) = 0 Or Dir(sFilePath) = "" Then
        ReadUserDailyData = Empty
        Exit Function
    End If

    Application.ScreenUpdating = False
    Dim wb As Workbook
    Set wb = Workbooks.Open(sFilePath, ReadOnly:=True, UpdateLinks:=0)

    Dim ws As Worksheet
    Set ws = wb.Sheets(1)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row

    If lastRow < 2 Then
        wb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        ReadUserDailyData = Empty
        Exit Function
    End If

    ' Read data (skip header)
    ReadUserDailyData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, NUM_COLUMNS)).Value

    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    ReadUserDailyData = Empty
End Function

'------------------------------------------------------------------------------
' ReadAllUsersDailyData - Reads all user files for a given date
'
' Returns: Collection of arrays, each element is a user's data
'------------------------------------------------------------------------------
Public Function ReadAllUsersDailyData(ByVal dtDate As Date) As Collection
    On Error GoTo ErrHandler

    Dim col As New Collection
    Dim sMonthFolder As String
    sMonthFolder = GetNetworkPath() & FOLDER_DATA & "\" & Format(dtDate, "YYYY-MM") & "\"

    If Dir(sMonthFolder, vbDirectory) = "" Then
        Set ReadAllUsersDailyData = col
        Exit Function
    End If

    ' Find all files matching *_YYYYMMDD.xlsx
    Dim sDateSuffix As String
    sDateSuffix = "_" & Format(dtDate, "YYYYMMDD") & ".xlsx"

    Dim sFile As String
    sFile = Dir(sMonthFolder & "*" & sDateSuffix)

    Application.ScreenUpdating = False

    Do While Len(sFile) > 0
        Dim wb As Workbook
        Set wb = Workbooks.Open(sMonthFolder & sFile, ReadOnly:=True, UpdateLinks:=0)

        Dim ws As Worksheet
        Set ws = wb.Sheets(1)

        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row

        If lastRow >= 2 Then
            Dim vData As Variant
            vData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, NUM_COLUMNS)).Value
            col.Add vData
        End If

        wb.Close SaveChanges:=False
        sFile = Dir()
    Loop

    Application.ScreenUpdating = True
    Set ReadAllUsersDailyData = col
    Exit Function

ErrHandler:
    On Error Resume Next
    Application.ScreenUpdating = True
    Set ReadAllUsersDailyData = New Collection
End Function

'------------------------------------------------------------------------------
' GetUserFilesForDate - Returns a collection of file paths for a given date
'------------------------------------------------------------------------------
Public Function GetUserFilesForDate(ByVal dtDate As Date) As Collection
    On Error GoTo ErrHandler

    Dim col As New Collection
    Dim sMonthFolder As String
    sMonthFolder = GetNetworkPath() & FOLDER_DATA & "\" & Format(dtDate, "YYYY-MM") & "\"

    If Dir(sMonthFolder, vbDirectory) = "" Then
        Set GetUserFilesForDate = col
        Exit Function
    End If

    Dim sDateSuffix As String
    sDateSuffix = "_" & Format(dtDate, "YYYYMMDD") & ".xlsx"

    Dim sFile As String
    sFile = Dir(sMonthFolder & "*" & sDateSuffix)

    Do While Len(sFile) > 0
        col.Add sMonthFolder & sFile
        sFile = Dir()
    Loop

    Set GetUserFilesForDate = col
    Exit Function

ErrHandler:
    Set GetUserFilesForDate = New Collection
End Function

'------------------------------------------------------------------------------
' SyncPendingRecords - Re-sends any locally-saved records that failed to sync
'------------------------------------------------------------------------------
Public Function SyncPendingRecords() As Long
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row

    Dim lSynced As Long
    lSynced = 0

    Dim i As Long
    For i = 2 To lastRow
        Dim sStatus As String
        sStatus = CStr(ws.Cells(i, COL_SYNCSTATUS).Value)

        ' Try to sync records that are Pending or have Error status
        If sStatus = "Pending" Or Left(sStatus, 5) = "Error" Or Len(sStatus) = 0 Then
            If SaveToNetwork(ws, i) Then
                lSynced = lSynced + 1
            End If
        End If
    Next i

    SyncPendingRecords = lSynced
    Exit Function

ErrHandler:
    SyncPendingRecords = lSynced
End Function

'------------------------------------------------------------------------------
' GetSyncStats - Returns sync status counts as a string
'------------------------------------------------------------------------------
Public Function GetSyncStats() As String
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row

    If lastRow < 2 Then
        GetSyncStats = "No records"
        Exit Function
    End If

    Dim lTotal As Long, lSynced As Long, lPending As Long, lError As Long
    Dim i As Long

    For i = 2 To lastRow
        lTotal = lTotal + 1
        Dim sStatus As String
        sStatus = CStr(ws.Cells(i, COL_SYNCSTATUS).Value)

        If sStatus = "Synced" Then
            lSynced = lSynced + 1
        ElseIf sStatus = "Pending" Then
            lPending = lPending + 1
        ElseIf Left(sStatus, 5) = "Error" Then
            lError = lError + 1
        Else
            lPending = lPending + 1 ' Treat empty as pending
        End If
    Next i

    GetSyncStats = "Total: " & lTotal & " | Synced: " & lSynced & _
                   " | Pending: " & lPending & " | Errors: " & lError
    Exit Function

ErrHandler:
    GetSyncStats = "Unable to read sync status"
End Function
