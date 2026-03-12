VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrntData
   Caption         =   "APP Billing - Print Daily Report"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6000
   OleObjectBlob   =   "frmPrntData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrntData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmPrntData - Print Daily Report / PDF Generation Form
' APP Billing System
'
' FORM CONTROLS NEEDED (create in VBA Editor form designer):
'   lblUser        - Label showing current user
'   lblDate        - Label "Date of Service:"
'   txtReportDate  - TextBox for date entry (DD/MM/YYYY, auto-formatted)
'   lblAnesth      - Label "Anesthesiologist:"
'   lstAnesth      - ListBox for anesthesiologist (filter-as-you-type)
'   txtShftSrtTime - TextBox for shift start time (HHMMhr, auto-formatted)
'   txtShftFinTime - TextBox for shift finish time (HHMMhr, auto-formatted)
'   cmdSearch      - CommandButton "Search"
'   lstDataBse     - ListBox showing search results (4 columns)
'   cmdPreview     - CommandButton "Preview"
'   cmdGeneratePDF - CommandButton "Generate PDF"
'   cmdExit        - CommandButton "Exit"
'   lblStatus      - Label for status messages
'
' NOTE: txtShftSrtTime, txtShftFinTime, cmdSearch, and lstDataBse must be
'       added in the VBA Editor form designer.  All other controls existed
'       in the original form.  Until the new controls are added the form
'       still opens and all original functionality continues to work.
'==============================================================================
Option Explicit

' Anesthesiologist list data (two columns: name, MSP#) — mirrors frmSaveData
Private m_aAnesth As Variant
Private m_sSearchAnesth As String

' Flag to prevent recursive formatting in Change events
Private m_bFormatting As Boolean

'------------------------------------------------------------------------------
' Form Initialize
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Dim sStep As String
    On Error GoTo ErrHandler

    m_bFormatting = False
    m_sSearchAnesth = ""

    ' Non-critical display controls
    sStep = "Setting user label"
    On Error Resume Next
    lblUser.Caption = "Logged in as: " & Application.UserName
    On Error GoTo ErrHandler

    ' Date field - show placeholder, not today's date
    sStep = "Setting date placeholder"
    txtReportDate.Value = "DD/MM/YYYY"

    ' Time fields (new controls - use Controls() to avoid compile error if
    ' the user hasn't yet added them to the form designer)
    sStep = "Setting time placeholders"
    On Error Resume Next
    Me.Controls("txtShftSrtTime").Value = "HHMMhr"
    Me.Controls("txtShftFinTime").Value = "HHMMhr"
    On Error GoTo ErrHandler

    ' Configure lstAnesth as two-column (name + MSP#) — same as frmSaveData
    sStep = "Configuring anesthesiologist list"
    SetupTwoColumnListBox lstAnesth, "80;80"

    ' Verify LookupLists sheet exists
    sStep = "Opening LookupLists sheet"
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("LookupLists")
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        MsgBox "The LookupLists sheet is missing from this workbook." & vbCrLf & _
               "Please ensure the workbook was set up correctly (run InitialSetup).", _
               vbCritical, "Setup Error"
        Application.OnTime Now, "UnloadPrntData"
        Exit Sub
    End If

    ' Load anesthesiologist names + MSP numbers into two-column array
    sStep = "Loading anesthesiologist names"
    m_aAnesth = LoadTwoColumnsToArray(ws, 1, 2)     ' Columns A (name), B (MSP#)

    ' Populate list with all names on open
    sStep = "Populating anesthesiologist list"
    PopulateFullList2Col lstAnesth, m_aAnesth

    ' Clear result list if it exists
    On Error Resume Next
    Me.Controls("lstDataBse").Clear
    lblStatus.Caption = ""
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Error in frmPrntData.Initialize at step '" & sStep & "':" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Form Error"
    Application.OnTime Now, "UnloadPrntData"
End Sub

'==============================================================================
' ANESTHESIOLOGIST LIST - filter-as-you-type
'==============================================================================

Private Sub lstAnesth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress2Col lstAnesth, m_aAnesth, m_sSearchAnesth, KeyAscii
End Sub

'==============================================================================
' DATE FIELD - txtReportDate (DD/MM/YYYY, auto-formatted)
' txtReportDate exists in the original form so referenced directly by name.
'==============================================================================

Private Sub txtReportDate_Enter()
    If txtReportDate.Value = "DD/MM/YYYY" Then txtReportDate.Value = ""
End Sub

Private Sub txtReportDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtReportDate.Value)) = 0 Then txtReportDate.Value = "DD/MM/YYYY"
End Sub

Private Sub txtReportDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtReportDate_Change()
    If m_bFormatting Then Exit Sub
    If txtReportDate.Value = "DD/MM/YYYY" Or Len(txtReportDate.Value) = 0 Then Exit Sub
    FormatDateField txtReportDate
End Sub

'==============================================================================
' TIME FIELDS - txtShftSrtTime / txtShftFinTime (HHMMhr, auto-formatted)
' These are NEW controls to be added in the form designer.
' Their sub names compile without error; control references inside each sub
' use Me.Controls("name") so no "Variable not defined" even before the
' controls are placed on the form.
'==============================================================================

Private Sub txtShftSrtTime_Enter()
    On Error Resume Next
    Dim ctl As MSForms.TextBox
    Set ctl = Me.Controls("txtShftSrtTime")
    If Not ctl Is Nothing Then
        If ctl.Value = "HHMMhr" Then ctl.Value = ""
    End If
    On Error GoTo 0
End Sub

Private Sub txtShftSrtTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim ctl As MSForms.TextBox
    Set ctl = Me.Controls("txtShftSrtTime")
    If Not ctl Is Nothing Then
        If Len(Trim(ctl.Value)) = 0 Then ctl.Value = "HHMMhr"
    End If
    On Error GoTo 0
End Sub

Private Sub txtShftSrtTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtShftSrtTime_Change()
    If m_bFormatting Then Exit Sub
    On Error Resume Next
    Dim ctl As MSForms.TextBox
    Set ctl = Me.Controls("txtShftSrtTime")
    On Error GoTo 0
    If ctl Is Nothing Then Exit Sub
    If ctl.Value = "HHMMhr" Or Len(ctl.Value) = 0 Then Exit Sub
    FormatTimeField ctl
End Sub

Private Sub txtShftFinTime_Enter()
    On Error Resume Next
    Dim ctl As MSForms.TextBox
    Set ctl = Me.Controls("txtShftFinTime")
    If Not ctl Is Nothing Then
        If ctl.Value = "HHMMhr" Then ctl.Value = ""
    End If
    On Error GoTo 0
End Sub

Private Sub txtShftFinTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Dim ctl As MSForms.TextBox
    Set ctl = Me.Controls("txtShftFinTime")
    If Not ctl Is Nothing Then
        If Len(Trim(ctl.Value)) = 0 Then ctl.Value = "HHMMhr"
    End If
    On Error GoTo 0
End Sub

Private Sub txtShftFinTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtShftFinTime_Change()
    If m_bFormatting Then Exit Sub
    On Error Resume Next
    Dim ctl As MSForms.TextBox
    Set ctl = Me.Controls("txtShftFinTime")
    On Error GoTo 0
    If ctl Is Nothing Then Exit Sub
    If ctl.Value = "HHMMhr" Or Len(ctl.Value) = 0 Then Exit Sub
    FormatTimeField ctl
End Sub

'==============================================================================
' LIST BOX HELPERS — identical to frmSaveData implementations
'==============================================================================

Private Sub SetupTwoColumnListBox(ByRef lst As MSForms.ListBox, ByVal sWidths As String)
    On Error Resume Next
    lst.ColumnCount = 2
    lst.ColumnWidths = sWidths
    lst.MatchEntry = fmMatchEntryNone
    lst.RowSource = ""
    On Error GoTo 0
End Sub

Private Function LoadTwoColumnsToArray(ByVal ws As Worksheet, ByVal lCol1 As Long, _
                                       ByVal lCol2 As Long) As Variant
    Dim lr1 As Long, lr2 As Long
    lr1 = ws.Cells(ws.Rows.Count, lCol1).End(xlUp).Row
    lr2 = ws.Cells(ws.Rows.Count, lCol2).End(xlUp).Row
    Dim lastRow As Long
    lastRow = IIf(lr1 > lr2, lr1, lr2)

    Dim lCount As Long
    Dim i As Long
    For i = 2 To lastRow
        If Len(Trim(CStr(ws.Cells(i, lCol1).Value))) > 0 Or _
           Len(Trim(CStr(ws.Cells(i, lCol2).Value))) > 0 Then
            lCount = lCount + 1
        End If
    Next i

    If lCount = 0 Then
        LoadTwoColumnsToArray = Empty
        Exit Function
    End If

    Dim result() As String
    ReDim result(1 To lCount, 1 To 2)
    Dim idx As Long
    For i = 2 To lastRow
        Dim s1 As String, s2 As String
        s1 = Trim(CStr(ws.Cells(i, lCol1).Value))
        s2 = Trim(CStr(ws.Cells(i, lCol2).Value))
        If Len(s1) > 0 Or Len(s2) > 0 Then
            idx = idx + 1
            result(idx, 1) = s1
            result(idx, 2) = s2
        End If
    Next i

    LoadTwoColumnsToArray = result
End Function

Private Sub PopulateFullList2Col(ByRef lst As MSForms.ListBox, ByRef vItems As Variant)
    lst.Clear
    If IsEmpty(vItems) Then Exit Sub
    Dim i As Long
    For i = LBound(vItems, 1) To UBound(vItems, 1)
        lst.AddItem vItems(i, 1)
        lst.List(lst.ListCount - 1, 1) = vItems(i, 2)
    Next i
End Sub

Private Sub FilterListBox2Col(ByRef lst As MSForms.ListBox, ByRef vItems As Variant, _
                              ByVal sSearch As String)
    lst.Clear
    If IsEmpty(vItems) Then Exit Sub
    If Len(sSearch) = 0 Then
        Dim j As Long
        For j = LBound(vItems, 1) To UBound(vItems, 1)
            lst.AddItem vItems(j, 1)
            lst.List(lst.ListCount - 1, 1) = vItems(j, 2)
        Next j
        Exit Sub
    End If
    Dim i As Long
    Dim sLower As String
    sLower = LCase(sSearch)
    For i = LBound(vItems, 1) To UBound(vItems, 1)
        If LCase(Left(vItems(i, 1), Len(sSearch))) = sLower Or _
           LCase(Left(vItems(i, 2), Len(sSearch))) = sLower Then
            lst.AddItem vItems(i, 1)
            lst.List(lst.ListCount - 1, 1) = vItems(i, 2)
        End If
    Next i
    If lst.ListCount = 1 Then lst.ListIndex = 0
End Sub

Private Sub HandleListKeyPress2Col(ByRef lst As MSForms.ListBox, ByRef vItems As Variant, _
                                   ByRef sSearch As String, ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8  ' Backspace
            If Len(sSearch) > 0 Then sSearch = Left(sSearch, Len(sSearch) - 1)
        Case 27 ' Escape
            sSearch = ""
        Case Else
            sSearch = sSearch & Chr(KeyAscii)
    End Select

    FilterListBox2Col lst, vItems, sSearch

    If lst.ListCount = 0 And Len(sSearch) > 0 Then
        MsgBox "No matching item found for '" & sSearch & "'." & vbCrLf & _
               "The list has been reset. Please try again.", _
               vbExclamation, "Item Not Found"
        sSearch = ""
        FilterListBox2Col lst, vItems, sSearch
    End If

    KeyAscii = 0
End Sub

'==============================================================================
' FORMAT HELPER FUNCTIONS
'==============================================================================

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

Private Sub FormatDateField(ByRef ctl As MSForms.TextBox)
    m_bFormatting = True
    Dim sDigits As String
    sDigits = ExtractDigits(ctl.Value)
    If Len(sDigits) > 8 Then sDigits = Left(sDigits, 8)
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

Private Sub FormatTimeField(ByRef ctl As MSForms.TextBox)
    m_bFormatting = True
    Dim sDigits As String
    sDigits = ExtractDigits(ctl.Value)
    If Len(sDigits) > 4 Then sDigits = Left(sDigits, 4)
    Dim sFormatted As String
    If Len(sDigits) = 4 Then
        sFormatted = sDigits & "hr"
    Else
        sFormatted = sDigits
    End If
    If sFormatted <> ctl.Value Then
        ctl.Value = sFormatted
        If Len(sDigits) = 4 Then
            ctl.SelStart = 4
        Else
            ctl.SelStart = Len(sFormatted)
        End If
    End If
    m_bFormatting = False
End Sub

'==============================================================================
' SEARCH - cmdSearch
' Searches DailyDatabase by selected anesthesiologist + date.
' Results shown in lstDataBse (4 columns: Proc Code / Start / Finish / IC).
' lstDataBse is a NEW control — accessed via Me.Controls() to avoid compile
' errors before it has been added to the form designer.
'==============================================================================

Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler

    If lstAnesth.ListIndex < 0 Then
        MsgBox "Please select an anesthesiologist to search.", vbExclamation, "Validation"
        Exit Sub
    End If

    If Len(txtReportDate.Value) = 0 Or txtReportDate.Value = "DD/MM/YYYY" Then
        MsgBox "Please enter a date to search.", vbExclamation, "Validation"
        Exit Sub
    End If

    Dim dtDate As Date
    On Error Resume Next
    dtDate = ParseDateInput(txtReportDate.Value)
    If Err.Number <> 0 Then
        On Error GoTo ErrHandler
        MsgBox "Invalid date format. Please use DD/MM/YYYY.", vbExclamation, "Validation"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    ' Retrieve result list box via Controls collection (new control)
    Dim lst As MSForms.ListBox
    On Error Resume Next
    Set lst = Me.Controls("lstDataBse")
    On Error GoTo ErrHandler
    If lst Is Nothing Then
        MsgBox "Result list box (lstDataBse) has not been added to the form yet." & vbCrLf & _
               "Please add it in the VBA Editor form designer.", _
               vbExclamation, "Control Missing"
        Exit Sub
    End If

    Dim sAnesth As String
    sAnesth = lstAnesth.Value

    On Error Resume Next
    lblStatus.Caption = "Searching..."
    DoEvents

    lst.Clear
    lst.ColumnCount = 4
    lst.ColumnWidths = "70;40;40;30"
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row

    Dim lFound As Long
    lFound = 0
    Dim i As Long

    For i = 2 To lastRow
        Dim sRowAnesth As String
        sRowAnesth = Trim(CStr(ws.Cells(i, COL_ANESTH).Value))

        If InStr(1, sRowAnesth, sAnesth, vbTextCompare) > 0 Then
            ' Date comparison: handle both text "DD/MM/YYYY" and numeric date serial
            ' (older records may have been stored as a numeric serial by Excel)
            Dim bDateMatch As Boolean
            Dim dtRowDate As Date
            Dim vDateCell As Variant
            vDateCell = ws.Cells(i, COL_DATE).Value
            If IsNumeric(vDateCell) And Not IsEmpty(vDateCell) Then
                dtRowDate = CDate(vDateCell)
                bDateMatch = (DateValue(dtRowDate) = DateValue(dtDate))
            Else
                bDateMatch = TryParseDateDMY(Trim(CStr(vDateCell)), dtRowDate) And _
                             (DateValue(dtRowDate) = DateValue(dtDate))
            End If

            If bDateMatch Then
                On Error Resume Next
                lst.AddItem CStr(ws.Cells(i, COL_PROCCODE).Value)
                lst.List(lst.ListCount - 1, 1) = CStr(ws.Cells(i, COL_STARTTIME).Value)
                lst.List(lst.ListCount - 1, 2) = CStr(ws.Cells(i, COL_FINTIME).Value)
                lst.List(lst.ListCount - 1, 3) = CStr(ws.Cells(i, COL_MAXIC).Value)
                On Error GoTo ErrHandler
                lFound = lFound + 1
            End If
        End If
    Next i

    On Error Resume Next
    If lFound = 0 Then
        lblStatus.Caption = "No records found."
        MsgBox "No records found for " & sAnesth & " on " & Format(dtDate, "DD/MM/YYYY") & ".", _
               vbInformation, "Search Results"
    Else
        lblStatus.Caption = lFound & " record(s) found."
    End If
    On Error GoTo 0
    Exit Sub

ErrHandler:
    On Error Resume Next
    lblStatus.Caption = "Search failed."
    MsgBox "Search error: " & Err.Description, vbCritical, "Search Error"
End Sub

'------------------------------------------------------------------------------
' Preview Button
' Populates ORReportingForm, hides this form, and adds a "Return to Form"
' button on the sheet so the user can switch back easily.
'------------------------------------------------------------------------------
Private Sub cmdPreview_Click()
    On Error GoTo ErrHandler
    If Not ValidateInputs() Then Exit Sub

    Dim sAnesth As String
    sAnesth = lstAnesth.Value

    Dim dtDate As Date
    dtDate = ParseDateInput(txtReportDate.Value)

    lblStatus.Caption = "Generating preview..."
    DoEvents

    Dim sShiftStart As String, sShiftFin As String
    On Error Resume Next
    sShiftStart = Me.Controls("txtShftSrtTime").Value
    If sShiftStart = "HHMMhr" Then sShiftStart = ""
    sShiftFin = Me.Controls("txtShftFinTime").Value
    If sShiftFin = "HHMMhr" Then sShiftFin = ""
    On Error GoTo ErrHandler

    GenerateDailyPDF sAnesth, dtDate, bPreview:=True, _
                     sShiftStart:=sShiftStart, sShiftFin:=sShiftFin
    ' GenerateDailyPDF already activates the sheet; add button then hide form
    AddReturnButton
    Me.Hide
    Exit Sub

ErrHandler:
    lblStatus.Caption = "Preview failed."
    MsgBox "Error generating preview: " & Err.Description, vbCritical, "Preview Error"
End Sub

'------------------------------------------------------------------------------
' Generate PDF Button
' Generates PDF, clears ORReportingForm, and opens the PDF automatically.
' The form stays visible so the user can generate another or exit normally.
'------------------------------------------------------------------------------
Private Sub cmdGeneratePDF_Click()
    On Error GoTo ErrHandler
    If Not ValidateInputs() Then Exit Sub

    Dim sAnesth As String
    sAnesth = lstAnesth.Value

    Dim dtDate As Date
    dtDate = ParseDateInput(txtReportDate.Value)

    lblStatus.Caption = "Generating PDF..."
    DoEvents

    Dim sShiftStart As String, sShiftFin As String
    On Error Resume Next
    sShiftStart = Me.Controls("txtShftSrtTime").Value
    If sShiftStart = "HHMMhr" Then sShiftStart = ""
    sShiftFin = Me.Controls("txtShftFinTime").Value
    If sShiftFin = "HHMMhr" Then sShiftFin = ""
    On Error GoTo ErrHandler

    Dim sResult As String
    sResult = GenerateDailyPDF(sAnesth, dtDate, bPreview:=False, _
                               sShiftStart:=sShiftStart, sShiftFin:=sShiftFin)

    If Len(sResult) > 0 Then
        lblStatus.Caption = "PDF saved: " & sResult
        ' Remove any stale Return button, then open the PDF automatically
        RemoveReturnButton
        On Error Resume Next
        Shell "explorer.exe """ & sResult & """", vbNormalFocus
        On Error GoTo ErrHandler
    Else
        lblStatus.Caption = "PDF generation failed or no data found."
    End If
    Exit Sub

ErrHandler:
    lblStatus.Caption = "PDF generation failed."
    MsgBox "Error generating PDF: " & Err.Description, vbCritical, "PDF Error"
End Sub

'------------------------------------------------------------------------------
' Exit Button
'------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'------------------------------------------------------------------------------
' ValidateInputs
'------------------------------------------------------------------------------
Private Function ValidateInputs() As Boolean
    ValidateInputs = True

    If lstAnesth.ListIndex < 0 Then
        MsgBox "Please select an anesthesiologist.", vbExclamation, "Validation"
        ValidateInputs = False
        Exit Function
    End If

    If Len(txtReportDate.Value) = 0 Or txtReportDate.Value = "DD/MM/YYYY" Then
        MsgBox "Please enter a date.", vbExclamation, "Validation"
        ValidateInputs = False
        Exit Function
    End If

    Dim dtTest As Date
    On Error Resume Next
    dtTest = ParseDateInput(txtReportDate.Value)
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Invalid date format. Please use DD/MM/YYYY.", vbExclamation, "Validation"
        ValidateInputs = False
        Exit Function
    End If
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------
' ParseDateInput
'------------------------------------------------------------------------------
Private Function ParseDateInput(ByVal sDate As String) As Date
    ParseDateInput = ParseDateDMY(sDate)
End Function
