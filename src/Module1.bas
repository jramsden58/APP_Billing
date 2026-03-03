Attribute VB_Name = "Module1"
'==============================================================================
' Module1 - Main Application Logic
' APP Billing System
'
' Contains form display routines and core Submit/Reset logic.
' Updated to include network save and sync status tracking.
'==============================================================================
Option Explicit

' Public variable to pass last saved row number from Submit to callers
Public g_lLastSavedRow As Long

'------------------------------------------------------------------------------
' Reset - Resets the frmSaveData form to default values
'------------------------------------------------------------------------------
Public Sub Reset()
    On Error GoTo ErrHandler

    With frmSaveData
        ' Site defaults
        .optRCH.Value = True
        .optERH.Value = False

        ' Date placeholder
        .txtDteOfSer.Value = "DD/MM/YYYY"

        ' On Call
        .chxOnCall.Value = False

        ' Shift type
        .optOR.Value = True
        .optOutOfOR.Value = False

        ' Procedure fields
        .txtSurgProcCode.Value = ""
        .txtProcStrtTime.Value = "HHMMhr"
        .txtProcFinTime.Value = "HHMMhr"
        .txtMaxIC.Value = ""

        ' Repopulate all list boxes with full item lists (clears search text too)
        .RepopulateAllLists

        ' WCB fields
        .txtWCBNum.Value = ""
        .txtWCBInjSide.Value = ""
        .txtWCBDiagCode.Value = ""
        .txtWCBInjCode.Value = ""
        .txtWCBDteofInj.Value = "DD/MM/YYYY"

        ' Reset background colors
        .txtDteOfSer.BackColor = &HFFFFFF
        .txtSurgProcCode.BackColor = &HFFFFFF
        .txtProcStrtTime.BackColor = &HFFFFFF
        .txtProcFinTime.BackColor = &HFFFFFF
        .txtMaxIC.BackColor = &HFFFFFF
        .txtWCBNum.BackColor = &HFFFFFF
        .txtWCBInjSide.BackColor = &HFFFFFF
        .txtWCBDiagCode.BackColor = &HFFFFFF
        .txtWCBInjCode.BackColor = &HFFFFFF
        .txtWCBDteofInj.BackColor = &HFFFFFF
    End With

    Exit Sub

ErrHandler:
    ' Silently handle missing controls during reset (form may not be fully loaded)
    Resume Next
End Sub

'------------------------------------------------------------------------------
' Submit - Saves form data to DailyDatabase and network share
' Returns True on success, False on failure
' Sets public g_lLastSavedRow so the caller can report the row number
'------------------------------------------------------------------------------
Public Function Submit() As Boolean
    On Error GoTo ErrHandler

    Dim sStep As String
    sStep = "Opening DailyDatabase sheet"

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    ' Find next empty row using Find (searches from bottom for last cell with data)
    ' This avoids End(xlUp) which can be fooled by stray content at the sheet bottom
    sStep = "Finding next empty row"
    Dim lRow As Long
    Dim rngLast As Range
    Set rngLast = ws.Columns(COL_ANESTH).Find(What:="*", LookIn:=xlValues, _
                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If rngLast Is Nothing Then
        ' Column B is completely empty - start at row 2 (row 1 = header)
        lRow = 2
    ElseIf rngLast.Row < 2 Then
        ' Only the header row has content
        lRow = 2
    Else
        lRow = rngLast.Row + 1
    End If

    sStep = "Writing form data to row " & lRow

    With frmSaveData
        ' Column A: Serial number (simple value, not formula)
        ws.Cells(lRow, COL_SERIAL).Value = lRow - 1

        ' Column B: Anesthesiologist (two-column list: column 0 = name)
        If .lstAnesth.ListIndex >= 0 Then
            ws.Cells(lRow, COL_ANESTH).Value = .lstAnesth.List(.lstAnesth.ListIndex, 0)
        End If

        ' Column C: Site
        If .optRCH.Value Then
            ws.Cells(lRow, COL_SITE).Value = "RCH"
        ElseIf .optERH.Value Then
            ws.Cells(lRow, COL_SITE).Value = "ERH"
        End If

        ' Column D: Date of Service
        Dim sDate As String
        sDate = .txtDteOfSer.Value
        If sDate <> "DD/MM/YYYY" And Len(sDate) > 0 Then
            ws.Cells(lRow, COL_DATE).Value = sDate
        End If

        ' Column E: Shift Name
        If .lstShftName.ListIndex >= 0 Then
            ws.Cells(lRow, COL_SHIFT).Value = .lstShftName.List(.lstShftName.ListIndex)
        End If

        ' Column F: On Call (store as "Yes"/"No" string for consistency)
        ws.Cells(lRow, COL_ONCALL).Value = IIf(.chxOnCall.Value, "Yes", "No")

        ' Column G: Shift Type
        If .optOR.Value Then
            ws.Cells(lRow, COL_SHIFTTYPE).Value = "OR"
        ElseIf .optOutOfOR.Value Then
            ws.Cells(lRow, COL_SHIFTTYPE).Value = "Out of OR"
        End If

        ' Column H: Surgical Procedure Code
        ws.Cells(lRow, COL_PROCCODE).Value = .txtSurgProcCode.Value

        ' Column I: Procedure Start Time
        Dim sStart As String
        sStart = .txtProcStrtTime.Value
        If sStart <> "HHMMhr" And Len(sStart) > 0 Then
            ws.Cells(lRow, COL_STARTTIME).Value = sStart
        End If

        ' Column J: Procedure Finish Time
        Dim sFinish As String
        sFinish = .txtProcFinTime.Value
        If sFinish <> "HHMMhr" And Len(sFinish) > 0 Then
            ws.Cells(lRow, COL_FINTIME).Value = sFinish
        End If

        ' Column K: Maximum IC Level
        ws.Cells(lRow, COL_MAXIC).Value = .txtMaxIC.Value

        ' Column L: Consults (two-column list: column 0 = code)
        If .lstEval.ListIndex >= 0 Then
            ws.Cells(lRow, COL_CONSULT).Value = .lstEval.List(.lstEval.ListIndex, 0)
        End If

        ' Columns M-O: Fee Modifiers (two-column list: column 0 = code)
        If .lstMod1.ListIndex >= 0 Then
            ws.Cells(lRow, COL_MOD1).Value = .lstMod1.List(.lstMod1.ListIndex, 0)
        End If
        If .lstMod2.ListIndex >= 0 Then
            ws.Cells(lRow, COL_MOD2).Value = .lstMod2.List(.lstMod2.ListIndex, 0)
        End If
        If .lstMod3.ListIndex >= 0 Then
            ws.Cells(lRow, COL_MOD3).Value = .lstMod3.List(.lstMod3.ListIndex, 0)
        End If

        ' Column P: Resuscitation (two-column list: column 0 = code)
        If .lstResus.ListIndex >= 0 Then
            ws.Cells(lRow, COL_RESUS).Value = .lstResus.List(.lstResus.ListIndex, 0)
        End If

        ' Column Q: Obstetrics (two-column list: column 0 = code)
        If .lstObs.ListIndex >= 0 Then
            ws.Cells(lRow, COL_OBS).Value = .lstObs.List(.lstObs.ListIndex, 0)
        End If

        ' Column R: Acute Pain (two-column list: column 0 = code)
        If .lstAcPain.ListIndex >= 0 Then
            ws.Cells(lRow, COL_ACUTEPAIN).Value = .lstAcPain.List(.lstAcPain.ListIndex, 0)
        End If

        ' Column S: Diagnostic and Chronic Pain (two-column list: column 0 = code)
        If .lstChPain.ListIndex >= 0 Then
            ws.Cells(lRow, COL_CHRONPAIN).Value = .lstChPain.List(.lstChPain.ListIndex, 0)
        End If

        ' Column T: Miscellaneous (two-column list: column 0 = code)
        If .lstMisc.ListIndex >= 0 Then
            ws.Cells(lRow, COL_MISC).Value = .lstMisc.List(.lstMisc.ListIndex, 0)
        End If

        ' Columns U-Y: WCB fields
        ws.Cells(lRow, COL_WCBNUM).Value = .txtWCBNum.Value

        ws.Cells(lRow, COL_WCBSIDE).Value = .txtWCBInjSide.Value

        ws.Cells(lRow, COL_WCBDIAG).Value = .txtWCBDiagCode.Value

        ws.Cells(lRow, COL_WCBINJ).Value = .txtWCBInjCode.Value

        Dim sWCBDate As String
        sWCBDate = .txtWCBDteofInj.Value
        If sWCBDate <> "DD/MM/YYYY" And Len(sWCBDate) > 0 Then
            ws.Cells(lRow, COL_WCBDATE).Value = sWCBDate
        End If

        ' Column Z: Submitted By (Windows username, consistent with file naming)
        sStep = "Writing Submitted By"
        ws.Cells(lRow, COL_SUBMBY).Value = GetCurrentUser()

        ' Column AA: Submitted On (timestamp - use nn for minutes, not MM)
        ws.Cells(lRow, COL_SUBMON).Value = FormatTimestamp(Now)

        ' Column AB: Sync Status (initially empty, set by SaveToNetwork)
        ws.Cells(lRow, COL_SYNCSTATUS).Value = ""
    End With

    ' VERIFY data was actually written before saving
    sStep = "Verifying data write"
    Dim sVerify As String
    sVerify = CStr(ws.Cells(lRow, COL_ANESTH).Value)
    If Len(sVerify) = 0 Then
        MsgBox "WARNING: Data write verification failed!" & vbCrLf & _
               "Row " & lRow & ", Column B (Anesthesiologist) is empty after writing." & vbCrLf & _
               "The workbook file is: " & ThisWorkbook.FullName & vbCrLf & _
               "The sheet name is: " & ws.Name, _
               vbCritical, "Write Verification Failed"
        Submit = False
        Exit Function
    End If

    ' Persist workbook to disk so data survives form close / Excel exit
    sStep = "Saving workbook"
    ThisWorkbook.Save

    ' Verify data survived the save (catches Workbook_BeforeSave interference)
    sStep = "Post-save verification"
    Dim sPostSave As String
    sPostSave = CStr(ws.Cells(lRow, COL_ANESTH).Value)
    If Len(sPostSave) = 0 Then
        MsgBox "WARNING: Data disappeared after ThisWorkbook.Save!" & vbCrLf & _
               "Something (possibly a Workbook_BeforeSave event) is clearing data." & vbCrLf & _
               "Check the ThisWorkbook module in VBA Editor (Alt+F11) for event code.", _
               vbCritical, "Post-Save Verification Failed"
        Submit = False
        Exit Function
    End If

    ' Store the row number for the success message
    g_lLastSavedRow = lRow

    ' Save to network share
    sStep = "Syncing to network"
    If IsNetworkAvailable() Then
        Dim bSynced As Boolean
        bSynced = SaveToNetwork(ws, lRow)

        If Not bSynced Then
            MsgBox "Data saved locally but could not be synced to the network share." & vbCrLf & _
                   "The record has been marked as 'Pending' and will be synced later." & vbCrLf & vbCrLf & _
                   "Use the 'Sync' button on the Home page to retry.", _
                   vbExclamation, "Network Sync Warning"
        End If
    Else
        ws.Cells(lRow, COL_SYNCSTATUS).Value = "Pending"
        MsgBox "Network share is not available. Data has been saved locally only." & vbCrLf & _
               "The record will be synced when the network connection is restored.", _
               vbExclamation, "Offline Mode"
    End If

    Submit = True
    Exit Function

ErrHandler:
    Submit = False
    MsgBox "Error saving data at step: " & sStep & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical, "Save Error"
End Function

'------------------------------------------------------------------------------
' Show_Form1 - Opens the Data Entry form
'------------------------------------------------------------------------------
Public Sub Show_Form1()
    frmSaveData.Show
End Sub

'------------------------------------------------------------------------------
' Show_Form2 - Opens the Print Data / PDF form
'------------------------------------------------------------------------------
Public Sub Show_Form2()
    frmPrntData.Show
End Sub

'------------------------------------------------------------------------------
' Show_Form3 - Opens the Superuser Access form
'------------------------------------------------------------------------------
Public Sub Show_Form3()
    On Error GoTo ErrHandler
    frmSuperUser.Show
    Exit Sub
ErrHandler:
    MsgBox "Error opening SuperUser form:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "Source: " & Err.Source, vbCritical, "SuperUser Form Error"
End Sub

'------------------------------------------------------------------------------
' Show_Form4 - Opens the Daily Export form
'------------------------------------------------------------------------------
Public Sub Show_Form4()
    frmDailyExport.Show
End Sub

'------------------------------------------------------------------------------
' Show_DailyData / Show_DailyExport - Opens the Daily Export form
' Aliases called directly by the "Export Daily Data" button on the Home sheet
'------------------------------------------------------------------------------
Public Sub Show_DailyData()
    frmDailyExport.Show
End Sub

Public Sub Show_DailyExport()
    frmDailyExport.Show
End Sub

'------------------------------------------------------------------------------
' SyncNow - Syncs all pending records to the network
'------------------------------------------------------------------------------
Public Sub SyncNow()
    On Error GoTo ErrHandler

    If Not IsNetworkAvailable() Then
        MsgBox "Network share is not available. Please check your connection.", _
               vbExclamation, "Network Unavailable"
        Exit Sub
    End If

    Dim lSynced As Long
    lSynced = SyncPendingRecords()

    If lSynced > 0 Then
        MsgBox lSynced & " record(s) successfully synced to the network share.", _
               vbInformation, "Sync Complete"
    Else
        MsgBox "No pending records to sync.", vbInformation, "Sync Status"
    End If

    ' Update status on Home sheet
    UpdateHomeStatus
    Exit Sub

ErrHandler:
    MsgBox "Error during sync: " & Err.Description, vbCritical, "Sync Error"
End Sub

'------------------------------------------------------------------------------
' UpdateHomeStatus - Updates the sync status display on the Home sheet
'------------------------------------------------------------------------------
Public Sub UpdateHomeStatus()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Home")

    ' Display sync stats in a designated cell
    ws.Range("A20").Value = "Sync Status: " & GetSyncStats()
    ws.Range("A20").Font.Size = 9
    ws.Range("A20").Font.Color = RGB(100, 100, 100)

    ' Display network connection status
    If IsNetworkAvailable() Then
        ws.Range("A21").Value = "Network: Connected"
        ws.Range("A21").Font.Color = RGB(0, 128, 0)
    Else
        ws.Range("A21").Value = "Network: Disconnected"
        ws.Range("A21").Font.Color = RGB(200, 0, 0)
    End If
    ws.Range("A21").Font.Size = 9

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' InitialSetup - Run once to set up the application
'------------------------------------------------------------------------------
Public Sub InitialSetup()
    ' Create Settings sheet
    InitializeSettingsSheet

    ' Prompt for network path
    ShowConfigDialog

    ' Always create folder structure first — this also creates the base path
    ' if it is a new local folder (e.g. C:\...\test that doesn't exist yet)
    EnsureNetworkFolders

    ' Create SuperUsers file if the path is now accessible
    If IsNetworkAvailable() Then
        Dim sPath As String
        sPath = GetNetworkPath() & FOLDER_CONFIG & "\SuperUsers.xlsx"
        If Dir(sPath) = "" Then
            CreateSuperUsersFile
        End If
    End If

    ' Ensure SearchData sheet exists for search functionality
    EnsureSheetExists "SearchData"

    ' Ensure ExportLog sheet exists for tracking exported dates
    Dim wsExport As Worksheet
    Set wsExport = EnsureSheetExists("ExportLog")
    If Len(wsExport.Cells(1, 1).Value) = 0 Then
        wsExport.Cells(1, 1).Value = "Date"
        wsExport.Cells(1, 2).Value = "ExportedBy"
        wsExport.Cells(1, 3).Value = "ExportedOn"
    End If
    wsExport.Visible = xlSheetVeryHidden

    ' Add Sync Status header to DailyDatabase if missing
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")
    If ws.Cells(1, COL_SYNCSTATUS).Value = "" Then
        ws.Cells(1, COL_SYNCSTATUS).Value = "Sync Status"
    End If

    ' Clean up any stray data beyond the real data range
    CleanDailyDatabase

    ' Update Home sheet status
    UpdateHomeStatus

    MsgBox "Initial setup complete!" & vbCrLf & vbCrLf & _
           "Network Path: " & GetNetworkPath() & vbCrLf & _
           "User: " & GetCurrentUser(), _
           vbInformation, "Setup Complete"
End Sub

'------------------------------------------------------------------------------
' CleanDailyDatabase - Removes stray content beyond the real data range
' Run this to fix issues where End(xlUp) finds content at the bottom of the sheet
'------------------------------------------------------------------------------
Public Sub CleanDailyDatabase()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    ' Find the true last row of data (searching from the top)
    Dim lTrueLastRow As Long
    Dim rngLast As Range
    Set rngLast = ws.UsedRange.Find(What:="*", LookIn:=xlValues, _
                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

    If rngLast Is Nothing Then
        lTrueLastRow = 1 ' Only header or empty
    Else
        lTrueLastRow = rngLast.Row
    End If

    ' Check if UsedRange extends far beyond the true data
    Dim lUsedLastRow As Long
    lUsedLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    If lUsedLastRow > lTrueLastRow + 10 Then
        ' Clear everything below the real data (plus a small buffer)
        Dim lClearFrom As Long
        lClearFrom = lTrueLastRow + 1
        ws.Rows(lClearFrom & ":" & lUsedLastRow).Delete
        MsgBox "Cleaned up " & (lUsedLastRow - lClearFrom + 1) & _
               " stray rows from DailyDatabase." & vbCrLf & _
               "Data now ends at row " & lTrueLastRow & ".", _
               vbInformation, "Cleanup Complete"
    End If

    Exit Sub
ErrHandler:
    MsgBox "Cleanup error: " & Err.Description, vbExclamation, "Cleanup"
End Sub
