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
'   lstAnesth      - ListBox for anesthesiologist (filter-as-you-type, 3-letter)
'   txtShftSrtTime - TextBox for shift start time (HHMMhr, auto-formatted)
'   txtShftFinTime - TextBox for shift finish time (HHMMhr, auto-formatted)
'   cmdSearch      - CommandButton "Search"
'   lstDataBse     - ListBox showing search results (4 columns)
'   cmdPreview     - CommandButton "Preview"
'   cmdGeneratePDF - CommandButton "Generate PDF"
'   cmdExit        - CommandButton "Exit"
'   lblStatus      - Label for status messages
'==============================================================================
Option Explicit

' Master anesthesiologist list for filter-as-you-type
Private m_aAnesthNames() As String
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

    ' Date field
    sStep = "Setting date field"
    txtReportDate.Value = Format(Date, "DD/MM/YYYY")

    ' Time fields
    sStep = "Setting time fields"
    On Error Resume Next
    txtShftSrtTime.Value = "HHMMhr"
    txtShftFinTime.Value = "HHMMhr"
    On Error GoTo ErrHandler

    ' Configure lstAnesth for manual filter-as-you-type
    sStep = "Configuring anesthesiologist list"
    On Error Resume Next
    lstAnesth.MatchEntry = fmMatchEntryNone
    lstAnesth.RowSource = ""
    On Error GoTo ErrHandler

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

    ' Load anesthesiologist names into master array
    sStep = "Loading anesthesiologist names"
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Count non-empty rows first
    Dim lCount As Long
    Dim i As Long
    Dim sCellVal As String
    lCount = 0
    For i = 2 To lastRow
        sCellVal = ""
        On Error Resume Next
        sCellVal = Trim(CStr(ws.Cells(i, 1).Value))
        On Error GoTo ErrHandler
        If Len(sCellVal) > 0 Then lCount = lCount + 1
    Next i

    If lCount > 0 Then
        ReDim m_aAnesthNames(1 To lCount)
        Dim idx As Long
        idx = 0
        For i = 2 To lastRow
            sCellVal = ""
            On Error Resume Next
            sCellVal = Trim(CStr(ws.Cells(i, 1).Value))
            On Error GoTo ErrHandler
            If Len(sCellVal) > 0 Then
                idx = idx + 1
                m_aAnesthNames(idx) = sCellVal
            End If
        Next i
    Else
        ReDim m_aAnesthNames(0 To 0)
        m_aAnesthNames(0) = ""
    End If

    ' Populate list with all names
    sStep = "Populating anesthesiologist list"
    PopulateAnesthList ""

    ' Try to pre-select the current user's name
    sStep = "Pre-selecting user in list"
    Dim sUser As String
    sUser = Application.UserName
    For i = 0 To lstAnesth.ListCount - 1
        If InStr(1, lstAnesth.List(i), sUser, vbTextCompare) > 0 Then
            lstAnesth.ListIndex = i
            Exit For
        End If
    Next i

    On Error Resume Next
    lblStatus.Caption = ""
    lstDataBse.Clear
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Error in frmPrntData.Initialize at step '" & sStep & "':" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Form Error"
    Application.OnTime Now, "UnloadPrntData"
End Sub

'==============================================================================
' ANESTHESIOLOGIST LIST - filter-as-you-type (3-letter match)
'==============================================================================

Private Sub lstAnesth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 ' Backspace
            If Len(m_sSearchAnesth) > 0 Then
                m_sSearchAnesth = Left(m_sSearchAnesth, Len(m_sSearchAnesth) - 1)
            End If
        Case 27 ' Escape
            m_sSearchAnesth = ""
        Case Else
            m_sSearchAnesth = m_sSearchAnesth & Chr(KeyAscii)
    End Select

    PopulateAnesthList m_sSearchAnesth

    ' If no matches, notify and reset to full list
    If lstAnesth.ListCount = 0 And Len(m_sSearchAnesth) > 0 Then
        MsgBox "No anesthesiologist found matching '" & m_sSearchAnesth & "'." & vbCrLf & _
               "The list has been reset. Please try again.", _
               vbExclamation, "Not Found"
        m_sSearchAnesth = ""
        PopulateAnesthList ""
    End If

    KeyAscii = 0
End Sub

'------------------------------------------------------------------------------
' PopulateAnesthList - Filters lstAnesth by prefix match
'------------------------------------------------------------------------------
Private Sub PopulateAnesthList(ByVal sFilter As String)
    On Error Resume Next
    lstAnesth.Clear

    If Not IsAnesthArrayReady() Then
        On Error GoTo 0
        Exit Sub
    End If

    Dim i As Long
    Dim sLower As String
    sLower = LCase(sFilter)

    For i = LBound(m_aAnesthNames) To UBound(m_aAnesthNames)
        If Len(m_aAnesthNames(i)) > 0 Then
            If Len(sFilter) = 0 Or _
               LCase(Left(m_aAnesthNames(i), Len(sFilter))) = sLower Then
                lstAnesth.AddItem m_aAnesthNames(i)
            End If
        End If
    Next i

    ' Auto-select if exactly one match remains
    If lstAnesth.ListCount = 1 Then lstAnesth.ListIndex = 0
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' IsAnesthArrayReady - Returns True if m_aAnesthNames is dimensioned and usable
'------------------------------------------------------------------------------
Private Function IsAnesthArrayReady() As Boolean
    On Error GoTo NotReady
    Dim n As Long
    n = UBound(m_aAnesthNames)
    IsAnesthArrayReady = (n >= LBound(m_aAnesthNames) And Len(m_aAnesthNames(LBound(m_aAnesthNames))) > 0)
    Exit Function
NotReady:
    IsAnesthArrayReady = False
End Function

'==============================================================================
' DATE FIELD - txtReportDate (DD/MM/YYYY, auto-formatted)
'==============================================================================

Private Sub txtReportDate_Enter()
    If txtReportDate.Value = "DD/MM/YYYY" Then txtReportDate.Value = ""
End Sub

Private Sub txtReportDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtReportDate.Value)) = 0 Then txtReportDate.Value = "DD/MM/YYYY"
End Sub

Private Sub txtReportDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits; slashes are auto-inserted by FormatDateField
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtReportDate_Change()
    If m_bFormatting Then Exit Sub
    If txtReportDate.Value = "DD/MM/YYYY" Or Len(txtReportDate.Value) = 0 Then Exit Sub
    FormatDateField txtReportDate
End Sub

'==============================================================================
' TIME FIELDS - txtShftSrtTime / txtShftFinTime (HHMMhr, auto-formatted)
'==============================================================================

Private Sub txtShftSrtTime_Enter()
    If txtShftSrtTime.Value = "HHMMhr" Then txtShftSrtTime.Value = ""
End Sub

Private Sub txtShftSrtTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtShftSrtTime.Value)) = 0 Then txtShftSrtTime.Value = "HHMMhr"
End Sub

Private Sub txtShftSrtTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits; "hr" suffix is auto-appended
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtShftSrtTime_Change()
    If m_bFormatting Then Exit Sub
    If txtShftSrtTime.Value = "HHMMhr" Or Len(txtShftSrtTime.Value) = 0 Then Exit Sub
    FormatTimeField txtShftSrtTime
End Sub

Private Sub txtShftFinTime_Enter()
    If txtShftFinTime.Value = "HHMMhr" Then txtShftFinTime.Value = ""
End Sub

Private Sub txtShftFinTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtShftFinTime.Value)) = 0 Then txtShftFinTime.Value = "HHMMhr"
End Sub

Private Sub txtShftFinTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtShftFinTime_Change()
    If m_bFormatting Then Exit Sub
    If txtShftFinTime.Value = "HHMMhr" Or Len(txtShftFinTime.Value) = 0 Then Exit Sub
    FormatTimeField txtShftFinTime
End Sub

'==============================================================================
' FORMAT HELPER FUNCTIONS
'==============================================================================

'------------------------------------------------------------------------------
' ExtractDigits - Returns only the digit characters from a string
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
' FormatDateField - Auto-inserts "/" separators for DD/MM/YYYY as user types
'------------------------------------------------------------------------------
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

'------------------------------------------------------------------------------
' FormatTimeField - Auto-appends "hr" suffix when 4 digits are entered
'------------------------------------------------------------------------------
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
' Searches DailyDatabase by selected anesthesiologist + date
' Results shown in lstDataBse (columns: Proc Code, Start, Finish, IC)
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

    Dim sAnesth As String
    sAnesth = lstAnesth.Value

    Dim sDateStr As String
    sDateStr = Format(dtDate, "DD/MM/YYYY")

    On Error Resume Next
    lblStatus.Caption = "Searching..."
    DoEvents

    ' Configure result list: 4 columns (Proc Code | Start | Finish | IC)
    lstDataBse.Clear
    lstDataBse.ColumnCount = 4
    lstDataBse.ColumnWidths = "70;40;40;30"
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
        Dim sRowDate As String
        sRowAnesth = Trim(CStr(ws.Cells(i, COL_ANESTH).Value))
        sRowDate   = Trim(CStr(ws.Cells(i, COL_DATE).Value))

        If InStr(1, sRowAnesth, sAnesth, vbTextCompare) > 0 And sRowDate = sDateStr Then
            On Error Resume Next
            lstDataBse.AddItem CStr(ws.Cells(i, COL_PROCCODE).Value)
            lstDataBse.List(lstDataBse.ListCount - 1, 1) = CStr(ws.Cells(i, COL_STARTTIME).Value)
            lstDataBse.List(lstDataBse.ListCount - 1, 2) = CStr(ws.Cells(i, COL_FINTIME).Value)
            lstDataBse.List(lstDataBse.ListCount - 1, 3) = CStr(ws.Cells(i, COL_MAXIC).Value)
            On Error GoTo ErrHandler
            lFound = lFound + 1
        End If
    Next i

    On Error Resume Next
    If lFound = 0 Then
        lblStatus.Caption = "No records found."
        MsgBox "No records found for " & sAnesth & " on " & sDateStr & ".", _
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
' Preview Button - Shows the populated ORReportingForm
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

    Dim sResult As String
    sResult = GenerateDailyPDF(sAnesth, dtDate, bPreview:=True)

    lblStatus.Caption = "Preview ready. Check the ORReportingForm sheet."
    Exit Sub

ErrHandler:
    lblStatus.Caption = "Preview failed."
    MsgBox "Error generating preview: " & Err.Description, vbCritical, "Preview Error"
End Sub

'------------------------------------------------------------------------------
' Generate PDF Button - Creates and saves the PDF report
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

    Dim sResult As String
    sResult = GenerateDailyPDF(sAnesth, dtDate, bPreview:=False)

    If Len(sResult) > 0 Then
        lblStatus.Caption = "PDF saved: " & sResult

        If MsgBox("PDF generated successfully. Open the file?", _
                  vbYesNo + vbQuestion, "PDF Ready") = vbYes Then
            On Error Resume Next
            Shell "explorer.exe """ & sResult & """", vbNormalFocus
            On Error GoTo 0
        End If
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
' ValidateInputs - Checks that required fields are filled
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

    ' Validate date format
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
' ParseDateInput - Parses a DD/MM/YYYY date string (locale-safe)
'------------------------------------------------------------------------------
Private Function ParseDateInput(ByVal sDate As String) As Date
    ParseDateInput = ParseDateDMY(sDate)
End Function
