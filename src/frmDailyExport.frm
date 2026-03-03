VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDailyExport
   Caption         =   "APP Billing - Daily Data Export"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5500
   OleObjectBlob   =   "frmDailyExport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDailyExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmDailyExport - Daily Data Export Form
' APP Billing System
'
' Allows users to export all users' daily data into a single Excel file.
' After export, the date is locked - no further edits by the user for that date.
' Available to all users (not superuser-restricted).
'
' FORM CONTROLS NEEDED (create in VBA Editor form designer):
'   lblTitle     - Label "Export All Users' Daily Data"
'   lblDate      - Label "Date:"
'   txtExportDate - TextBox for date (DD/MM/YYYY)
'   cmdExport    - CommandButton "Export"
'   cmdExit      - CommandButton "Exit"
'   lblStatus    - Label for status messages
'   chkOpenFile  - CheckBox "Open file after export"
'   chkPDF       - CheckBox "Also generate PDF reports (via ORReportingForm)"
'==============================================================================
Option Explicit

' Flag to prevent recursive formatting in Change events
Private m_bFormatting As Boolean

'------------------------------------------------------------------------------
' Form Initialize
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    txtExportDate.Value = "DD/MM/YYYY"   ' placeholder — user types the actual date
    lblStatus.Caption = ""
    chkOpenFile.Value = True
    On Error Resume Next
    chkPDF.Value = False
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Export Button
'------------------------------------------------------------------------------
Private Sub cmdExport_Click()
    On Error GoTo ErrHandler

    ' Validate date
    If Len(txtExportDate.Value) = 0 Or txtExportDate.Value = "DD/MM/YYYY" Then
        MsgBox "Please enter a date.", vbExclamation, "Validation"
        Exit Sub
    End If

    Dim dtDate As Date
    dtDate = ParseDate(txtExportDate.Value)

    ' Check if already exported
    If IsDateExported(dtDate) Then
        If MsgBox("This date has already been exported." & vbCrLf & _
                  "Do you want to export again?", _
                  vbYesNo + vbQuestion, "Already Exported") = vbNo Then
            Exit Sub
        End If
    End If

    lblStatus.Caption = "Exporting data for " & Format(dtDate, "DD/MM/YYYY") & "..."
    DoEvents

    Dim sResult As String
    sResult = ConsolidateDailyData(dtDate)

    If Len(sResult) > 0 Then
        ' Mark the date as exported (locks further edits for this user/date)
        MarkDateExported dtDate

        lblStatus.Caption = "Export complete: " & sResult

        ' Generate per-anesthesiologist PDFs via ORReportingForm if requested
        Dim bPDF As Boolean
        On Error Resume Next
        bPDF = chkPDF.Value
        On Error GoTo ErrHandler

        If bPDF Then
            lblStatus.Caption = "Generating PDFs via ORReportingForm..."
            DoEvents
            Dim sPDFFolder As String
            sPDFFolder = GenerateConsolidatedPDF(dtDate)
            If Len(sPDFFolder) > 0 Then
                lblStatus.Caption = "Export and PDF reports complete."
            Else
                lblStatus.Caption = "Excel export complete. No data for PDF generation."
            End If
        End If

        If chkOpenFile.Value Then
            Workbooks.Open sResult
        Else
            MsgBox "Export complete." & vbCrLf & vbCrLf & _
                   "Excel file saved to:" & vbCrLf & sResult & _
                   IIf(bPDF And Len(sPDFFolder) > 0, vbCrLf & vbCrLf & _
                   "PDF reports saved to:" & vbCrLf & sPDFFolder, ""), _
                   vbInformation, "Export Complete"
        End If
    Else
        lblStatus.Caption = "No data found for export."
    End If

    Exit Sub
ErrHandler:
    lblStatus.Caption = "Export failed."
    MsgBox "Export error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Exit Button
'------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'==============================================================================
' DATE FIELD - txtExportDate (DD/MM/YYYY, auto-formatted)
'==============================================================================

Private Sub txtExportDate_Enter()
    If txtExportDate.Value = "DD/MM/YYYY" Then txtExportDate.Value = ""
End Sub

Private Sub txtExportDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtExportDate.Value)) = 0 Then txtExportDate.Value = "DD/MM/YYYY"
End Sub

Private Sub txtExportDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits; slashes are auto-inserted by FormatDateField
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtExportDate_Change()
    If m_bFormatting Then Exit Sub
    If txtExportDate.Value = "DD/MM/YYYY" Or Len(txtExportDate.Value) = 0 Then Exit Sub
    FormatDateField txtExportDate
End Sub

'==============================================================================
' FORMAT HELPERS
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

'------------------------------------------------------------------------------
' ParseDate - Helper to parse DD/MM/YYYY date strings (locale-safe)
'------------------------------------------------------------------------------
Private Function ParseDate(ByVal sDate As String) As Date
    ParseDate = ParseDateDMY(sDate)
End Function
