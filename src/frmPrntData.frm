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
' Allows users to select a date and generate a PDF daily report
' using the ORReportingForm template.
'
' FORM CONTROLS NEEDED (create in VBA Editor form designer):
'   lblUser        - Label showing current user
'   lblDate        - Label "Date of Service:"
'   txtReportDate  - TextBox for date entry (DD/MM/YYYY)
'   lstAnesth      - ListBox for anesthesiologist selection
'   lblAnesth      - Label "Anesthesiologist:"
'   cmdPreview     - CommandButton "Preview"
'   cmdGeneratePDF - CommandButton "Generate PDF"
'   cmdExit        - CommandButton "Exit"
'   lblStatus      - Label for status messages
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Form Initialize
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Set current user display
    lblUser.Caption = "Logged in as: " & Application.UserName

    ' Set default date to today
    txtReportDate.Value = Format(Date, "DD/MM/YYYY")

    ' Populate anesthesiologist list from LookupLists
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LookupLists")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    lstAnesth.Clear
    Dim i As Long
    For i = 2 To lastRow
        If Len(ws.Cells(i, 1).Value) > 0 Then
            lstAnesth.AddItem ws.Cells(i, 1).Value
        End If
    Next i

    ' Try to select the current user's name in the list
    Dim sUser As String
    sUser = Application.UserName
    For i = 0 To lstAnesth.ListCount - 1
        If InStr(1, lstAnesth.List(i), sUser, vbTextCompare) > 0 Then
            lstAnesth.ListIndex = i
            Exit For
        End If
    Next i

    lblStatus.Caption = ""
End Sub

'------------------------------------------------------------------------------
' Preview Button - Shows the populated ORReportingForm
'------------------------------------------------------------------------------
Private Sub cmdPreview_Click()
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
End Sub

'------------------------------------------------------------------------------
' Generate PDF Button - Creates and saves the PDF report
'------------------------------------------------------------------------------
Private Sub cmdGeneratePDF_Click()
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

        ' Offer to open the PDF
        If MsgBox("PDF generated successfully. Open the file?", _
                  vbYesNo + vbQuestion, "PDF Ready") = vbYes Then
            On Error Resume Next
            Shell "explorer.exe """ & sResult & """", vbNormalFocus
            On Error GoTo 0
        End If
    Else
        lblStatus.Caption = "PDF generation failed or no data found."
    End If
End Sub

'------------------------------------------------------------------------------
' Exit Button
'------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'------------------------------------------------------------------------------
' Date field placeholder
'------------------------------------------------------------------------------
Private Sub txtReportDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                     ByVal X As Single, ByVal Y As Single)
    If txtReportDate.Value = "DD/MM/YYYY" Then
        txtReportDate.Value = ""
    End If
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
' ParseDateInput - Parses a DD/MM/YYYY date string
'------------------------------------------------------------------------------
Private Function ParseDateInput(ByVal sDate As String) As Date
    If IsDate(sDate) Then
        ParseDateInput = CDate(sDate)
        Exit Function
    End If

    Dim parts() As String
    parts = Split(sDate, "/")
    If UBound(parts) = 2 Then
        ParseDateInput = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
    Else
        Err.Raise 13, , "Invalid date format"
    End If
End Function
