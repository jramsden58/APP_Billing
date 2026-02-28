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
' Fixed: WCB date of injury field typo
' Added: Search, edit, delete functionality (previously commented out)
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Form Initialize - Sets up the form with default values
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Call Reset
End Sub

'------------------------------------------------------------------------------
' Save Button - Saves data to local database and network share
'------------------------------------------------------------------------------
Private Sub cmdSave_Click()
    ' Validate required fields
    If Not ValidateForm() Then Exit Sub

    If MsgBox("Save this record?", vbYesNo + vbQuestion, "Confirm Save") = vbYes Then
        Call Submit
        Call Reset
        MsgBox "Record saved successfully.", vbInformation, "Saved"
    End If
End Sub

'------------------------------------------------------------------------------
' Exit Button - Closes the form
'------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    If MsgBox("Are you sure you want to exit? Any unsaved data will be lost.", _
              vbYesNo + vbQuestion, "Confirm Exit") = vbYes Then
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

    Dim wsSearch As Worksheet
    Set wsSearch = ThisWorkbook.Sheets("SearchData")
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
' Delete Button - Deletes the currently selected/last entered record
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
            "Submitted: " & ws.Cells(lastRow, COL_SUBMON).Value

    If MsgBox(sInfo, vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
        ws.Rows(lastRow).Delete
        MsgBox "Record deleted.", vbInformation, "Deleted"
    End If

    Exit Sub
ErrHandler:
    MsgBox "Delete error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Edit Button - Loads the last record into the form for editing
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

        ' Shift type
        If CStr(ws.Cells(lastRow, COL_SHIFTTYPE).Value) = "OR" Then
            .optOR.Value = True
        Else
            .optOutOfOR.Value = True
        End If

        ' On Call
        .chxOnCall.Value = ws.Cells(lastRow, COL_ONCALL).Value

        ' Procedure fields
        .txtSurgProcCode.Value = CStr(ws.Cells(lastRow, COL_PROCCODE).Value)
        .txtProcStrtTime.Value = CStr(ws.Cells(lastRow, COL_STARTTIME).Value)
        .txtProcFinTime.Value = CStr(ws.Cells(lastRow, COL_FINTIME).Value)
        .txtMaxIC.Value = CStr(ws.Cells(lastRow, COL_MAXIC).Value)

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

    ' Delete the old record so the edited version replaces it
    ws.Rows(lastRow).Delete

    MsgBox "Record loaded for editing. Make your changes and click Save.", _
           vbInformation, "Edit Mode"

    Exit Sub
ErrHandler:
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

    ' Check date
    If txtDteOfSer.Value = "DD/MM/YYYY" Or Len(txtDteOfSer.Value) = 0 Then
        txtDteOfSer.BackColor = &HC0C0FF ' Light red
        bValid = False
    End If

    ' Check procedure code
    If Len(txtSurgProcCode.Value) = 0 Then
        txtSurgProcCode.BackColor = &HC0C0FF
        bValid = False
    End If

    ' Check start time
    If txtProcStrtTime.Value = "HH:MM" Or Len(txtProcStrtTime.Value) = 0 Then
        txtProcStrtTime.BackColor = &HC0C0FF
        bValid = False
    End If

    ' Check finish time
    If txtProcFinTime.Value = "HH:MM" Or Len(txtProcFinTime.Value) = 0 Then
        txtProcFinTime.BackColor = &HC0C0FF
        bValid = False
    End If

    If Not bValid Then
        MsgBox "Please fill in all required fields (highlighted in red).", _
               vbExclamation, "Validation Error"
    End If

    ValidateForm = bValid
End Function
