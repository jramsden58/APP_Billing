VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDailyExport
   Caption         =   "APP Billing - Daily Data Export"
   ClientHeight    =   3500
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
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' Form Initialize
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    txtExportDate.Value = Format(Date, "DD/MM/YYYY")
    lblStatus.Caption = ""
    chkOpenFile.Value = True
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

    lblStatus.Caption = "Exporting data for " & Format(dtDate, "DD/MM/YYYY") & "..."
    DoEvents

    Dim sResult As String
    sResult = ConsolidateDailyData(dtDate)

    If Len(sResult) > 0 Then
        lblStatus.Caption = "Export complete: " & sResult

        If chkOpenFile.Value Then
            Workbooks.Open sResult
        Else
            MsgBox "Export complete." & vbCrLf & vbCrLf & _
                   "File saved to:" & vbCrLf & sResult, _
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

'------------------------------------------------------------------------------
' Date placeholder handler
'------------------------------------------------------------------------------
Private Sub txtExportDate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                     ByVal X As Single, ByVal Y As Single)
    If txtExportDate.Value = "DD/MM/YYYY" Then
        txtExportDate.Value = ""
    End If
End Sub

'------------------------------------------------------------------------------
' ParseDate - Helper to parse DD/MM/YYYY date strings
'------------------------------------------------------------------------------
Private Function ParseDate(ByVal sDate As String) As Date
    If IsDate(sDate) Then
        ParseDate = CDate(sDate)
        Exit Function
    End If

    Dim parts() As String
    parts = Split(sDate, "/")
    If UBound(parts) = 2 Then
        ParseDate = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
    Else
        Err.Raise 13, , "Invalid date format. Use DD/MM/YYYY."
    End If
End Function
