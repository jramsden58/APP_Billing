VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSaveData
   Caption         =   "APP Billing - Data Entry"
   ClientHeight    =   10000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   OleObjectBlob   =   "frmSaveData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSaveData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmSaveData - Data Entry Form
' APP Billing System
'
' Handles user input for patient procedures and billing data.
' Features: Search, edit, delete functionality
'==============================================================================
Option Explicit

' Module-level variable to track the row being edited (0 = not editing)
Private m_lEditRow As Long

'------------------------------------------------------------------------------
' Form Initialize - Sets up the form with default values
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    m_lEditRow = 0
    Call Reset
End Sub

'------------------------------------------------------------------------------
' Save Button - Saves data to local database and network share
'------------------------------------------------------------------------------
Private Sub cmdSave_Click()
    On Error GoTo ErrHandler

    ' Validate required fields
    If Not ValidateForm() Then Exit Sub

    If MsgBox("Save this record?", vbYesNo + vbQuestion, "Confirm Save") = vbYes Then
        ' If editing, delete the old record now (just before saving the replacement)
        If m_lEditRow > 0 Then
            Dim wsEdit As Worksheet
            Set wsEdit = ThisWorkbook.Sheets("DailyDatabase")
            ' Verify the row still exists and is the same record
            If m_lEditRow <= wsEdit.Cells(wsEdit.Rows.Count, COL_ANESTH).End(xlUp).Row Then
                wsEdit.Rows(m_lEditRow).Delete
            End If
            m_lEditRow = 0
        End If

        Call Submit
        Call Reset
        MsgBox "Record saved successfully.", vbInformation, "Saved"
    End If
    Exit Sub

ErrHandler:
    m_lEditRow = 0
    MsgBox "Error saving record: " & Err.Description, vbCritical, "Save Error"
End Sub

'------------------------------------------------------------------------------
' Exit Button - Closes the form
'------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    ' Warn about unsaved edit
    Dim sWarning As String
    If m_lEditRow > 0 Then
        sWarning = "You are currently editing a record. " & _
                   "If you exit now, the original record will be preserved." & vbCrLf & vbCrLf & _
                   "Are you sure you want to exit?"
    Else
        sWarning = "Are you sure you want to exit? Any unsaved data will be lost."
    End If

    If MsgBox(sWarning, vbYesNo + vbQuestion, "Confirm Exit") = vbYes Then
        m_lEditRow = 0
        Call Reset
        Unload Me
    End If
End Sub

'------------------------------------------------------------------------------
' Search Button - Searches records in DailyDatabase
'------------------------------------------------------------------------------
Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler

    Dim sSearchTerm As String
    sSearchTerm = InputBox("Enter search term:" & vbCrLf & vbCrLf & _
                          "Searches across Anesthesiologist, Date, and Procedure Code fields.", _
                          "Search Records")

    If Len(sSearchTerm) = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    ' Ensure SearchData sheet exists
    Dim wsSearch As Worksheet
    Set wsSearch = EnsureSheetExists("SearchData")
    wsSearch.Cells.ClearContents

    ' Copy headers
    ws.Rows(1).Copy wsSearch.Rows(1)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row

    Dim lOutRow As Long
    lOutRow = 2

    Dim i As Long
    For i = 2 To lastRow
        ' Search in Anesthesiologist, Date, and Procedure Code columns
        If InStr(1, CStr(ws.Cells(i, COL_ANESTH).Value), sSearchTerm, vbTextCompare) > 0 Or _
           InStr(1, CStr(ws.Cells(i, COL_DATE).Value), sSearchTerm, vbTextCompare) > 0 Or _
           InStr(1, CStr(ws.Cells(i, COL_PROCCODE).Value), sSearchTerm, vbTextCompare) > 0 Then

            ws.Rows(i).Copy wsSearch.Rows(lOutRow)
            lOutRow = lOutRow + 1
        End If
    Next i

    If lOutRow = 2 Then
        MsgBox "No records found matching '" & sSearchTerm & "'.", _
               vbInformation, "Search Results"
    Else
        MsgBox (lOutRow - 2) & " record(s) found. Results are on the SearchData sheet.", _
               vbInformation, "Search Results"
        ThisWorkbook.Sheets("SearchData").Activate
    End If

    Exit Sub
ErrHandler:
    MsgBox "Search error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Delete Button - Deletes the last entered record
'------------------------------------------------------------------------------
Private Sub cmdDelete_Click()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No records to delete.", vbInformation, "Delete"
        Exit Sub
    End If

    ' Show last record info for confirmation
    Dim sInfo As String
    sInfo = "Delete the last record?" & vbCrLf & vbCrLf & _
            "Anesthesiologist: " & ws.Cells(lastRow, COL_ANESTH).Value & vbCrLf & _
            "Date: " & ws.Cells(lastRow, COL_DATE).Value & vbCrLf & _
            "Procedure: " & ws.Cells(lastRow, COL_PROCCODE).Value & vbCrLf & _
            "Submitted: " & ws.Cells(lastRow, COL_SUBMON).Value & vbCrLf & vbCrLf & _
            "Note: This only deletes the local copy. Network copy is not affected."

    If MsgBox(sInfo, vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
        ws.Rows(lastRow).Delete
        MsgBox "Record deleted locally.", vbInformation, "Deleted"
    End If

    Exit Sub
ErrHandler:
    MsgBox "Delete error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Edit Button - Loads the last record into the form for editing
' The original record is NOT deleted until the user clicks Save.
'------------------------------------------------------------------------------
Private Sub cmdEdit_Click()
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No records to edit.", vbInformation, "Edit"
        Exit Sub
    End If

    ' Load data into form
    With Me
        ' Find anesthesiologist in list
        Dim sAnesth As String
        sAnesth = CStr(ws.Cells(lastRow, COL_ANESTH).Value)
        Dim k As Long
        For k = 0 To .lstAnesth.ListCount - 1
            If .lstAnesth.List(k) = sAnesth Then
                .lstAnesth.ListIndex = k
                Exit For
            End If
        Next k

        ' Site
        If CStr(ws.Cells(lastRow, COL_SITE).Value) = "RCH" Then
            .optRCH.Value = True
        Else
            .optERH.Value = True
        End If

        ' Date
        .txtDteOfSer.Value = CStr(ws.Cells(lastRow, COL_DATE).Value)

        ' Shift Name - find in list
        Dim sShift As String
        sShift = CStr(ws.Cells(lastRow, COL_SHIFT).Value)
        If Len(sShift) > 0 Then
            For k = 0 To .lstShftName.ListCount - 1
                If .lstShftName.List(k) = sShift Then
                    .lstShftName.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Shift type
        If CStr(ws.Cells(lastRow, COL_SHIFTTYPE).Value) = "OR" Then
            .optOR.Value = True
        Else
            .optOutOfOR.Value = True
        End If

        ' On Call (handle both Boolean and "Yes"/"No" string)
        Dim vOnCall As Variant
        vOnCall = ws.Cells(lastRow, COL_ONCALL).Value
        .chxOnCall.Value = (vOnCall = True Or LCase(CStr(vOnCall & "")) = "yes")

        ' Procedure fields
        .txtSurgProcCode.Value = CStr(ws.Cells(lastRow, COL_PROCCODE).Value)
        .txtProcStrtTime.Value = CStr(ws.Cells(lastRow, COL_STARTTIME).Value)
        .txtProcFinTime.Value = CStr(ws.Cells(lastRow, COL_FINTIME).Value)
        .txtMaxIC.Value = CStr(ws.Cells(lastRow, COL_MAXIC).Value)

        ' Consults - find in list
        Dim sVal As String
        sVal = CStr(ws.Cells(lastRow, COL_CONSULT).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstEval.ListCount - 1
                If .lstEval.List(k) = sVal Then
                    .lstEval.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Fee Modifier 1
        sVal = CStr(ws.Cells(lastRow, COL_MOD1).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstMod1.ListCount - 1
                If .lstMod1.List(k) = sVal Then
                    .lstMod1.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Fee Modifier 2
        sVal = CStr(ws.Cells(lastRow, COL_MOD2).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstMod2.ListCount - 1
                If .lstMod2.List(k) = sVal Then
                    .lstMod2.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Fee Modifier 3
        sVal = CStr(ws.Cells(lastRow, COL_MOD3).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstMod3.ListCount - 1
                If .lstMod3.List(k) = sVal Then
                    .lstMod3.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Resuscitation
        sVal = CStr(ws.Cells(lastRow, COL_RESUS).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstResus.ListCount - 1
                If .lstResus.List(k) = sVal Then
                    .lstResus.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Obstetrics
        sVal = CStr(ws.Cells(lastRow, COL_OBS).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstObs.ListCount - 1
                If .lstObs.List(k) = sVal Then
                    .lstObs.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Acute Pain
        sVal = CStr(ws.Cells(lastRow, COL_ACUTEPAIN).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstAcPain.ListCount - 1
                If .lstAcPain.List(k) = sVal Then
                    .lstAcPain.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Chronic Pain
        sVal = CStr(ws.Cells(lastRow, COL_CHRONPAIN).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstChPain.ListCount - 1
                If .lstChPain.List(k) = sVal Then
                    .lstChPain.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Miscellaneous
        sVal = CStr(ws.Cells(lastRow, COL_MISC).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstMisc.ListCount - 1
                If .lstMisc.List(k) = sVal Then
                    .lstMisc.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' WCB fields
        .txtWCBNum.Value = CStr(ws.Cells(lastRow, COL_WCBNUM).Value)
        .txtWCBInjSide.Value = CStr(ws.Cells(lastRow, COL_WCBSIDE).Value)
        .txtWCBDiagCode.Value = CStr(ws.Cells(lastRow, COL_WCBDIAG).Value)
        .txtWCBInjCode.Value = CStr(ws.Cells(lastRow, COL_WCBINJ).Value)

        Dim sWCBDate As String
        sWCBDate = CStr(ws.Cells(lastRow, COL_WCBDATE).Value)
        If Len(sWCBDate) > 0 Then
            .txtWCBDteofInj.Value = sWCBDate
        End If
    End With

    ' Store the row being edited — do NOT delete until Save is clicked
    m_lEditRow = lastRow

    MsgBox "Record loaded for editing. Make your changes and click Save." & vbCrLf & _
           "The original record will be replaced when you save.", _
           vbInformation, "Edit Mode"

    Exit Sub
ErrHandler:
    m_lEditRow = 0
    MsgBox "Edit error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Placeholder text handlers - Clear placeholder on mouse click
'------------------------------------------------------------------------------
Private Sub txtDteOfSer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                   ByVal X As Single, ByVal Y As Single)
    If txtDteOfSer.Value = "DD/MM/YYYY" Then
        txtDteOfSer.Value = ""
    End If
End Sub

Private Sub txtProcStrtTime_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                       ByVal X As Single, ByVal Y As Single)
    If txtProcStrtTime.Value = "HH:MM" Then
        txtProcStrtTime.Value = ""
    End If
End Sub

Private Sub txtProcFinTime_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                      ByVal X As Single, ByVal Y As Single)
    If txtProcFinTime.Value = "HH:MM" Then
        txtProcFinTime.Value = ""
    End If
End Sub

Private Sub txtWCBDteofInj_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                      ByVal X As Single, ByVal Y As Single)
    If txtWCBDteofInj.Value = "DD/MM/YYYY" Then
        txtWCBDteofInj.Value = ""
    End If
End Sub

'------------------------------------------------------------------------------
' ValidateForm - Validates required fields before saving
'------------------------------------------------------------------------------
Private Function ValidateForm() As Boolean
    Dim bValid As Boolean
    bValid = True

    ' Reset all backgrounds
    txtDteOfSer.BackColor = &HFFFFFF
    txtSurgProcCode.BackColor = &HFFFFFF
    txtProcStrtTime.BackColor = &HFFFFFF
    txtProcFinTime.BackColor = &HFFFFFF

    ' Check anesthesiologist selected
    If lstAnesth.ListIndex < 0 Then
        MsgBox "Please select an anesthesiologist.", vbExclamation, "Validation"
        bValid = False
    End If

    ' Check date is valid DD/MM/YYYY
    Dim sDate As String
    sDate = txtDteOfSer.Value
    If sDate = "DD/MM/YYYY" Or Len(sDate) = 0 Then
        txtDteOfSer.BackColor = &HC0C0FF ' Light red
        bValid = False
    ElseIf Not IsValidDateDMY(sDate) Then
        txtDteOfSer.BackColor = &HC0C0FF
        MsgBox "Invalid date format. Please use DD/MM/YYYY.", vbExclamation, "Validation"
        bValid = False
    End If

    ' Check procedure code
    If Len(txtSurgProcCode.Value) = 0 Then
        txtSurgProcCode.BackColor = &HC0C0FF
        bValid = False
    End If

    ' Check start time is valid HH:MM
    Dim sStart As String
    sStart = txtProcStrtTime.Value
    If sStart = "HH:MM" Or Len(sStart) = 0 Then
        txtProcStrtTime.BackColor = &HC0C0FF
        bValid = False
    ElseIf Not IsValidTime24(sStart) Then
        txtProcStrtTime.BackColor = &HC0C0FF
        MsgBox "Invalid start time. Please use HH:MM (24-hour format).", vbExclamation, "Validation"
        bValid = False
    End If

    ' Check finish time is valid HH:MM
    Dim sFinish As String
    sFinish = txtProcFinTime.Value
    If sFinish = "HH:MM" Or Len(sFinish) = 0 Then
        txtProcFinTime.BackColor = &HC0C0FF
        bValid = False
    ElseIf Not IsValidTime24(sFinish) Then
        txtProcFinTime.BackColor = &HC0C0FF
        MsgBox "Invalid finish time. Please use HH:MM (24-hour format).", vbExclamation, "Validation"
        bValid = False
    End If

    ' Check WCB date if entered
    Dim sWCBDate As String
    sWCBDate = txtWCBDteofInj.Value
    If sWCBDate <> "DD/MM/YYYY" And Len(sWCBDate) > 0 Then
        If Not IsValidDateDMY(sWCBDate) Then
            txtWCBDteofInj.BackColor = &HC0C0FF
            MsgBox "Invalid WCB date of injury. Please use DD/MM/YYYY.", vbExclamation, "Validation"
            bValid = False
        End If
    End If

    If Not bValid Then
        MsgBox "Please fill in all required fields (highlighted in red).", _
               vbExclamation, "Validation Error"
    End If

    ValidateForm = bValid
End Function
