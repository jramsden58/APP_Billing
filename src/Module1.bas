Attribute VB_Name = "Module1"
'==============================================================================
' Module1 - Main Application Logic
' APP Billing System
'
' Contains form display routines and core Submit/Reset logic.
' Updated to include network save and sync status tracking.
'==============================================================================
Option Explicit

' Form positioning variables
Dim iWidth As Integer
Dim iHeight As Integer
Dim iLeft As Integer
Dim iTop As Integer
Dim bState As Boolean

'------------------------------------------------------------------------------
' Reset - Resets the frmSaveData form to default values
'------------------------------------------------------------------------------
Public Sub Reset()
    On Error GoTo ErrHandler

    With frmSaveData
        ' Set default anesthesiologist
        .lstAnesth.ListIndex = 0

        ' Site defaults
        .optRCH.Value = True
        .optERH.Value = False

        ' Date placeholder
        .txtDteOfSer.Value = "DD/MM/YYYY"

        ' Shift defaults
        .lstShftName.ListIndex = -1

        ' On Call
        .chxOnCall.Value = False

        ' Shift type
        .optOR.Value = True
        .optOutOfOR.Value = False

        ' Procedure fields
        .txtSurgProcCode.Value = ""
        .txtProcStrtTime.Value = "HH:MM"
        .txtProcFinTime.Value = "HH:MM"
        .txtMaxIC.Value = ""

        ' Fee item lists - clear selections
        .lstEval.ListIndex = -1
        .lstMod1.ListIndex = -1
        .lstMod2.ListIndex = -1
        .lstMod3.ListIndex = -1
        .lstResus.ListIndex = -1
        .lstObs.ListIndex = -1
        .lstAcPain.ListIndex = -1
        .lstChPain.ListIndex = -1
        .lstMisc.ListIndex = -1

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
'------------------------------------------------------------------------------
Public Sub Submit()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    ' Find next empty row (use End(xlUp) to handle gaps from deleted rows)
    Dim lRow As Long
    lRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row + 1

    With frmSaveData
        ' Column A: Serial number formula
        ws.Cells(lRow, COL_SERIAL).Formula = "=Row()-1"

        ' Column B: Anesthesiologist
        If .lstAnesth.ListIndex >= 0 Then
            ws.Cells(lRow, COL_ANESTH).Value = .lstAnesth.Value
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
            ws.Cells(lRow, COL_SHIFT).Value = .lstShftName.Value
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
        If sStart <> "HH:MM" Then
            ws.Cells(lRow, COL_STARTTIME).Value = sStart
        End If

        ' Column J: Procedure Finish Time
        Dim sFinish As String
        sFinish = .txtProcFinTime.Value
        If sFinish <> "HH:MM" Then
            ws.Cells(lRow, COL_FINTIME).Value = sFinish
        End If

        ' Column K: Maximum IC Level
        ws.Cells(lRow, COL_MAXIC).Value = .txtMaxIC.Value

        ' Column L: Consults
        If .lstEval.ListIndex >= 0 Then
            ws.Cells(lRow, COL_CONSULT).Value = .lstEval.Value
        End If

        ' Columns M-O: Fee Modifiers
        If .lstMod1.ListIndex >= 0 Then
            ws.Cells(lRow, COL_MOD1).Value = .lstMod1.Value
        End If
        If .lstMod2.ListIndex >= 0 Then
            ws.Cells(lRow, COL_MOD2).Value = .lstMod2.Value
        End If
        If .lstMod3.ListIndex >= 0 Then
            ws.Cells(lRow, COL_MOD3).Value = .lstMod3.Value
        End If

        ' Column P: Resuscitation
        If .lstResus.ListIndex >= 0 Then
            ws.Cells(lRow, COL_RESUS).Value = .lstResus.Value
        End If

        ' Column Q: Obstetrics
        If .lstObs.ListIndex >= 0 Then
            ws.Cells(lRow, COL_OBS).Value = .lstObs.Value
        End If

        ' Column R: Acute Pain
        If .lstAcPain.ListIndex >= 0 Then
            ws.Cells(lRow, COL_ACUTEPAIN).Value = .lstAcPain.Value
        End If

        ' Column S: Diagnostic and Chronic Pain
        If .lstChPain.ListIndex >= 0 Then
            ws.Cells(lRow, COL_CHRONPAIN).Value = .lstChPain.Value
        End If

        ' Column T: Miscellaneous
        If .lstMisc.ListIndex >= 0 Then
            ws.Cells(lRow, COL_MISC).Value = .lstMisc.Value
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
        ws.Cells(lRow, COL_SUBMBY).Value = GetCurrentUser()

        ' Column AA: Submitted On (timestamp - use nn for minutes, not MM)
        ws.Cells(lRow, COL_SUBMON).Value = FormatTimestamp(Now)

        ' Column AB: Sync Status (initially empty, set by SaveToNetwork)
        ws.Cells(lRow, COL_SYNCSTATUS).Value = ""
    End With

    ' Save to network share
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

    Exit Sub

ErrHandler:
    MsgBox "Error saving data: " & Err.Description, vbCritical, "Save Error"
End Sub

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
    frmSuperUser.Show
End Sub

'------------------------------------------------------------------------------
' Show_Form4 - Opens the Daily Export form
'------------------------------------------------------------------------------
Public Sub Show_Form4()
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

    ' Create folder structure
    If IsNetworkAvailable() Then
        EnsureNetworkFolders

        ' Create SuperUsers file if it doesn't exist
        Dim sPath As String
        sPath = GetNetworkPath() & FOLDER_CONFIG & "\SuperUsers.xlsx"
        If Dir(sPath) = "" Then
            CreateSuperUsersFile
        End If
    End If

    ' Ensure SearchData sheet exists for search functionality
    EnsureSheetExists "SearchData"

    ' Add Sync Status header to DailyDatabase if missing
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")
    If ws.Cells(1, COL_SYNCSTATUS).Value = "" Then
        ws.Cells(1, COL_SYNCSTATUS).Value = "Sync Status"
    End If

    ' Update Home sheet status
    UpdateHomeStatus

    MsgBox "Initial setup complete!" & vbCrLf & vbCrLf & _
           "Network Path: " & GetNetworkPath() & vbCrLf & _
           "User: " & GetCurrentUser(), _
           vbInformation, "Setup Complete"
End Sub
