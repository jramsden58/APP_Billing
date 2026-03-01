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
' Date format: DD/MM/YYYY (auto-formatted with slash insertion)
' Time format: HHMMhr (e.g., 0800hr) 24-hour clock
'==============================================================================
Option Explicit

' Module-level variable to track the row being edited (0 = not editing)
Private m_lEditRow As Long

' Flag to prevent recursive formatting in Change events
Private m_bFormatting As Boolean

'------------------------------------------------------------------------------
' Form Initialize - Sets up the form with default values
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    m_lEditRow = 0
    m_bFormatting = False

    ' Enable first-letter match on case detail list boxes
    ' so typing a letter/number jumps to the first matching item
    On Error Resume Next
    lstEval.MatchEntry = fmMatchEntryFirstLetter
    lstMod1.MatchEntry = fmMatchEntryFirstLetter
    lstMod2.MatchEntry = fmMatchEntryFirstLetter
    lstMod3.MatchEntry = fmMatchEntryFirstLetter
    lstResus.MatchEntry = fmMatchEntryFirstLetter
    lstObs.MatchEntry = fmMatchEntryFirstLetter
    lstAcPain.MatchEntry = fmMatchEntryFirstLetter
    lstChPain.MatchEntry = fmMatchEntryFirstLetter
    lstMisc.MatchEntry = fmMatchEntryFirstLetter
    lstShftName.MatchEntry = fmMatchEntryFirstLetter
    On Error GoTo 0

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
        ' Save the new/updated record first
        If Not Submit() Then
            ' Submit failed (error already shown by Submit)
            Exit Sub
        End If

        ' If editing, delete the old record AFTER successful save
        If m_lEditRow > 0 Then
            Dim wsEdit As Worksheet
            Set wsEdit = ThisWorkbook.Sheets("DailyDatabase")
            ' Verify the row still exists
            If m_lEditRow <= wsEdit.Cells(wsEdit.Rows.Count, COL_ANESTH).End(xlUp).Row Then
                wsEdit.Rows(m_lEditRow).Delete
            End If
            m_lEditRow = 0
        End If

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

    ' Temporarily disable formatting to avoid auto-format interference during load
    m_bFormatting = True

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

        ' Start Time - convert legacy HH:MM to HHMMhr if needed
        Dim sTime As String
        sTime = CStr(ws.Cells(lastRow, COL_STARTTIME).Value)
        If InStr(sTime, ":") > 0 Then
            sTime = Replace(sTime, ":", "") & "hr"
        End If
        .txtProcStrtTime.Value = sTime

        ' Finish Time - convert legacy HH:MM to HHMMhr if needed
        sTime = CStr(ws.Cells(lastRow, COL_FINTIME).Value)
        If InStr(sTime, ":") > 0 Then
            sTime = Replace(sTime, ":", "") & "hr"
        End If
        .txtProcFinTime.Value = sTime

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

    ' Re-enable formatting
    m_bFormatting = False

    ' Store the row being edited - do NOT delete until Save is clicked
    m_lEditRow = lastRow

    MsgBox "Record loaded for editing. Make your changes and click Save." & vbCrLf & _
           "The original record will be replaced when you save.", _
           vbInformation, "Edit Mode"

    Exit Sub
ErrHandler:
    m_bFormatting = False
    m_lEditRow = 0
    MsgBox "Edit error: " & Err.Description, vbCritical, "Error"
End Sub

'==============================================================================
' DATE FIELD AUTO-FORMATTING (DD/MM/YYYY)
' Auto-inserts "/" separators as the user types digits
'==============================================================================

'--- Date of Service ---
Private Sub txtDteOfSer_Enter()
    If txtDteOfSer.Value = "DD/MM/YYYY" Then
        txtDteOfSer.Value = ""
    End If
End Sub

Private Sub txtDteOfSer_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtDteOfSer.Value)) = 0 Then
        txtDteOfSer.Value = "DD/MM/YYYY"
    End If
End Sub

Private Sub txtDteOfSer_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits - slashes are auto-inserted
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDteOfSer_Change()
    If m_bFormatting Then Exit Sub
    If txtDteOfSer.Value = "DD/MM/YYYY" Or txtDteOfSer.Value = "" Then Exit Sub
    FormatDateField txtDteOfSer
End Sub

'--- WCB Date of Injury ---
Private Sub txtWCBDteofInj_Enter()
    If txtWCBDteofInj.Value = "DD/MM/YYYY" Then
        txtWCBDteofInj.Value = ""
    End If
End Sub

Private Sub txtWCBDteofInj_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtWCBDteofInj.Value)) = 0 Then
        txtWCBDteofInj.Value = "DD/MM/YYYY"
    End If
End Sub

Private Sub txtWCBDteofInj_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits - slashes are auto-inserted
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWCBDteofInj_Change()
    If m_bFormatting Then Exit Sub
    If txtWCBDteofInj.Value = "DD/MM/YYYY" Or txtWCBDteofInj.Value = "" Then Exit Sub
    FormatDateField txtWCBDteofInj
End Sub

'==============================================================================
' TIME FIELD AUTO-FORMATTING (HHMMhr)
' User types 4 digits, "hr" suffix is auto-appended
'==============================================================================

'--- Procedure Start Time ---
Private Sub txtProcStrtTime_Enter()
    If txtProcStrtTime.Value = "HHMMhr" Then
        txtProcStrtTime.Value = ""
    End If
End Sub

Private Sub txtProcStrtTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtProcStrtTime.Value)) = 0 Then
        txtProcStrtTime.Value = "HHMMhr"
    End If
End Sub

Private Sub txtProcStrtTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits - "hr" suffix is auto-appended
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtProcStrtTime_Change()
    If m_bFormatting Then Exit Sub
    If txtProcStrtTime.Value = "HHMMhr" Or txtProcStrtTime.Value = "" Then Exit Sub
    FormatTimeField txtProcStrtTime
End Sub

'--- Procedure Finish Time ---
Private Sub txtProcFinTime_Enter()
    If txtProcFinTime.Value = "HHMMhr" Then
        txtProcFinTime.Value = ""
    End If
End Sub

Private Sub txtProcFinTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtProcFinTime.Value)) = 0 Then
        txtProcFinTime.Value = "HHMMhr"
    End If
End Sub

Private Sub txtProcFinTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits - "hr" suffix is auto-appended
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtProcFinTime_Change()
    If m_bFormatting Then Exit Sub
    If txtProcFinTime.Value = "HHMMhr" Or txtProcFinTime.Value = "" Then Exit Sub
    FormatTimeField txtProcFinTime
End Sub

'==============================================================================
' FORMAT HELPER FUNCTIONS
'==============================================================================

'------------------------------------------------------------------------------
' ExtractDigits - Returns only digit characters from a string
'------------------------------------------------------------------------------
Private Function ExtractDigits(ByVal s As String) As String
    Dim i As Long
    Dim sResult As String
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            sResult = sResult & Mid(s, i, 1)
        End If
    Next i
    ExtractDigits = sResult
End Function

'------------------------------------------------------------------------------
' FormatDateField - Auto-inserts "/" separators for DD/MM/YYYY format
'------------------------------------------------------------------------------
Private Sub FormatDateField(ByRef ctl As MSForms.TextBox)
    m_bFormatting = True

    Dim sDigits As String
    sDigits = ExtractDigits(ctl.Value)

    ' Limit to 8 digits (DDMMYYYY)
    If Len(sDigits) > 8 Then sDigits = Left(sDigits, 8)

    ' Build formatted string with "/" separators
    Dim sFormatted As String
    If Len(sDigits) <= 2 Then
        sFormatted = sDigits
    ElseIf Len(sDigits) <= 4 Then
        sFormatted = Left(sDigits, 2) & "/" & Mid(sDigits, 3)
    Else
        sFormatted = Left(sDigits, 2) & "/" & Mid(sDigits, 3, 2) & "/" & Mid(sDigits, 5)
    End If

    If sFormatted <> ctl.Value Then
        ctl.Value = sFormatted
        ctl.SelStart = Len(sFormatted)
    End If

    m_bFormatting = False
End Sub

'------------------------------------------------------------------------------
' FormatTimeField - Auto-appends "hr" suffix when 4 digits are entered
'------------------------------------------------------------------------------
Private Sub FormatTimeField(ByRef ctl As MSForms.TextBox)
    m_bFormatting = True

    Dim sDigits As String
    sDigits = ExtractDigits(ctl.Value)

    ' Limit to 4 digits (HHMM)
    If Len(sDigits) > 4 Then sDigits = Left(sDigits, 4)

    ' Build formatted string - append "hr" when 4 digits entered
    Dim sFormatted As String
    If Len(sDigits) = 4 Then
        sFormatted = sDigits & "hr"
    Else
        sFormatted = sDigits
    End If

    If sFormatted <> ctl.Value Then
        ctl.Value = sFormatted
        ' Position cursor before "hr" suffix
        If Len(sDigits) = 4 Then
            ctl.SelStart = 4
        Else
            ctl.SelStart = Len(sFormatted)
        End If
    End If

    m_bFormatting = False
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

    ' Check start time is valid HHMMhr (24-hour)
    Dim sStart As String
    sStart = txtProcStrtTime.Value
    If sStart = "HHMMhr" Or Len(sStart) = 0 Then
        txtProcStrtTime.BackColor = &HC0C0FF
        bValid = False
    ElseIf Not IsValidTime24(sStart) Then
        txtProcStrtTime.BackColor = &HC0C0FF
        MsgBox "Invalid start time. Please enter 4 digits in 24-hour format (e.g., 0800hr).", _
               vbExclamation, "Validation"
        bValid = False
    End If

    ' Check finish time is valid HHMMhr (24-hour)
    Dim sFinish As String
    sFinish = txtProcFinTime.Value
    If sFinish = "HHMMhr" Or Len(sFinish) = 0 Then
        txtProcFinTime.BackColor = &HC0C0FF
        bValid = False
    ElseIf Not IsValidTime24(sFinish) Then
        txtProcFinTime.BackColor = &HC0C0FF
        MsgBox "Invalid finish time. Please enter 4 digits in 24-hour format (e.g., 1630hr).", _
               vbExclamation, "Validation"
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
