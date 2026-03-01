VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSuperUser
   Caption         =   "APP Billing - Superuser Access"
   ClientHeight    =   6000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8000
   OleObjectBlob   =   "frmSuperUser.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSuperUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmSuperUser - Superuser Access Form
' APP Billing System
'
' Provides authenticated superuser access to:
'   - Browse and open data files on the network share
'   - Run daily consolidation
'   - Run date range consolidation
'   - Manage SuperUsers list (Admin only)
'   - Configure system settings (Admin only)
'   - Change superuser password (Admin only)
'
' FORM CONTROLS NEEDED (create in VBA Editor form designer):
'   lblTitle       - Label "Superuser Access Panel"
'   lblUser        - Label showing current user and access level
'   lblStatus      - Label for status messages
'
'   ' Data Browsing Frame
'   fraDataBrowse  - Frame "Browse Data Files"
'   txtBrowseDate  - TextBox for date (DD/MM/YYYY)
'   cmdListFiles   - CommandButton "List Files"
'   lstFiles       - ListBox showing available files
'   cmdOpenFile    - CommandButton "Open Selected"
'
'   ' Consolidation Frame
'   fraConsolidate - Frame "Data Consolidation"
'   txtConsDate    - TextBox for single date consolidation
'   cmdConsolidate - CommandButton "Consolidate Day"
'   txtStartDate   - TextBox for range start
'   txtEndDate     - TextBox for range end
'   cmdConsRange   - CommandButton "Consolidate Range"
'
'   ' Admin Frame (visible to Admin only)
'   fraAdmin       - Frame "Administration"
'   cmdManageUsers - CommandButton "Manage SuperUsers"
'   cmdSettings    - CommandButton "System Settings"
'   cmdChangePwd   - CommandButton "Change Password"
'   cmdInitSetup   - CommandButton "Initial Setup"
'
'   cmdLogout      - CommandButton "Logout"
'   cmdExit        - CommandButton "Exit"
'==============================================================================
Option Explicit

Private m_colFiles As Collection

'------------------------------------------------------------------------------
' Form Initialize - Authenticate and set up UI
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Disable all controls before authentication to prevent bypass
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        On Error Resume Next
        ctrl.Enabled = False
        On Error GoTo 0
    Next ctrl

    ' Attempt authentication
    If Not AuthenticateSuperUser() Then
        MsgBox "Authentication failed. Superuser access denied.", _
               vbCritical, "Access Denied"
        ' Defer unload to after Initialize completes
        Application.OnTime Now, "UnloadSuperUser"
        Exit Sub
    End If

    ' Re-enable all controls after successful auth
    For Each ctrl In Me.Controls
        On Error Resume Next
        ctrl.Enabled = True
        On Error GoTo 0
    Next ctrl

    ' Set up user display
    lblUser.Caption = "User: " & Application.UserName & " | Access: " & GetAccessLevel()

    ' Set default dates
    txtBrowseDate.Value = Format(Date, "DD/MM/YYYY")
    txtConsDate.Value = Format(Date, "DD/MM/YYYY")
    txtStartDate.Value = Format(Date - 7, "DD/MM/YYYY")
    txtEndDate.Value = Format(Date, "DD/MM/YYYY")

    ' Show/hide admin features based on access level
    fraAdmin.Visible = IsAdmin()

    lblStatus.Caption = "Authenticated successfully."
    Set m_colFiles = New Collection
End Sub

'------------------------------------------------------------------------------
' List Files Button - Shows data files for selected date
'------------------------------------------------------------------------------
Private Sub cmdListFiles_Click()
    On Error GoTo ErrHandler

    Dim dtDate As Date
    dtDate = ParseDate(txtBrowseDate.Value)

    Set m_colFiles = GetUserFilesForDate(dtDate)

    lstFiles.Clear

    If m_colFiles.Count = 0 Then
        lstFiles.AddItem "(No files found for " & Format(dtDate, "DD/MM/YYYY") & ")"
        lblStatus.Caption = "No files found."
    Else
        Dim sFile As Variant
        For Each sFile In m_colFiles
            ' Show just the filename, not full path
            Dim sName As String
            sName = Mid(CStr(sFile), InStrRev(CStr(sFile), "\") + 1)
            lstFiles.AddItem sName
        Next sFile
        lblStatus.Caption = m_colFiles.Count & " file(s) found."
    End If

    Exit Sub
ErrHandler:
    MsgBox "Error listing files: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Open File Button - Opens the selected file
'------------------------------------------------------------------------------
Private Sub cmdOpenFile_Click()
    If lstFiles.ListIndex < 0 Then
        MsgBox "Please select a file to open.", vbExclamation, "No Selection"
        Exit Sub
    End If

    If m_colFiles.Count = 0 Then Exit Sub

    On Error GoTo ErrHandler
    Dim sPath As String
    sPath = m_colFiles(lstFiles.ListIndex + 1)

    Workbooks.Open sPath, ReadOnly:=True
    lblStatus.Caption = "Opened: " & lstFiles.Value
    Exit Sub

ErrHandler:
    MsgBox "Error opening file: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Consolidate Day Button
'------------------------------------------------------------------------------
Private Sub cmdConsolidate_Click()
    On Error GoTo ErrHandler

    Dim dtDate As Date
    dtDate = ParseDate(txtConsDate.Value)

    lblStatus.Caption = "Consolidating data for " & Format(dtDate, "DD/MM/YYYY") & "..."
    DoEvents

    Dim sResult As String
    sResult = ConsolidateDailyData(dtDate)

    If Len(sResult) > 0 Then
        lblStatus.Caption = "Consolidated file saved."
        If MsgBox("Consolidation complete. Open the file?" & vbCrLf & vbCrLf & _
                  sResult, vbYesNo + vbQuestion, "Consolidation Complete") = vbYes Then
            Workbooks.Open sResult
        End If
    Else
        lblStatus.Caption = "Consolidation produced no output."
    End If

    Exit Sub
ErrHandler:
    MsgBox "Consolidation error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Consolidate Range Button
'------------------------------------------------------------------------------
Private Sub cmdConsRange_Click()
    On Error GoTo ErrHandler

    Dim dtStart As Date
    dtStart = ParseDate(txtStartDate.Value)

    Dim dtEnd As Date
    dtEnd = ParseDate(txtEndDate.Value)

    lblStatus.Caption = "Consolidating data from " & Format(dtStart, "DD/MM/YYYY") & _
                        " to " & Format(dtEnd, "DD/MM/YYYY") & "..."
    DoEvents

    Dim sResult As String
    sResult = ConsolidateDateRange(dtStart, dtEnd)

    If Len(sResult) > 0 Then
        lblStatus.Caption = "Range consolidation complete."
        If MsgBox("Consolidation complete. Open the file?" & vbCrLf & vbCrLf & _
                  sResult, vbYesNo + vbQuestion, "Consolidation Complete") = vbYes Then
            Workbooks.Open sResult
        End If
    Else
        lblStatus.Caption = "Range consolidation produced no output."
    End If

    Exit Sub
ErrHandler:
    MsgBox "Consolidation error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Manage SuperUsers Button (Admin only)
'------------------------------------------------------------------------------
Private Sub cmdManageUsers_Click()
    On Error GoTo ErrHandler

    Dim sPath As String
    sPath = GetNetworkPath() & FOLDER_CONFIG & "\SuperUsers.xlsx"

    If Dir(sPath) = "" Then
        If MsgBox("SuperUsers.xlsx does not exist. Create it?", _
                  vbYesNo + vbQuestion, "Create File") = vbYes Then
            CreateSuperUsersFile
        End If
    Else
        Workbooks.Open sPath
        MsgBox "SuperUsers.xlsx is now open for editing." & vbCrLf & _
               "Remember to save when done.", vbInformation, "Manage SuperUsers"
    End If

    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' System Settings Button (Admin only)
'------------------------------------------------------------------------------
Private Sub cmdSettings_Click()
    ShowConfigDialog
End Sub

'------------------------------------------------------------------------------
' Change Password Button (Admin only)
'------------------------------------------------------------------------------
Private Sub cmdChangePwd_Click()
    ChangeSuperUserPassword
End Sub

'------------------------------------------------------------------------------
' Initial Setup Button (Admin only)
'------------------------------------------------------------------------------
Private Sub cmdInitSetup_Click()
    If MsgBox("Run initial system setup?" & vbCrLf & vbCrLf & _
              "This will:" & vbCrLf & _
              "- Create/verify the Settings sheet" & vbCrLf & _
              "- Prompt for network path configuration" & vbCrLf & _
              "- Create network folder structure" & vbCrLf & _
              "- Create SuperUsers.xlsx if needed" & vbCrLf & _
              "- Add Sync Status column to DailyDatabase", _
              vbYesNo + vbQuestion, "Initial Setup") = vbYes Then
        InitialSetup
        lblStatus.Caption = "Initial setup complete."
    End If
End Sub

'------------------------------------------------------------------------------
' Logout Button
'------------------------------------------------------------------------------
Private Sub cmdLogout_Click()
    LogOut
    lblStatus.Caption = "Logged out."
    MsgBox "You have been logged out.", vbInformation, "Logged Out"
    Unload Me
End Sub

'------------------------------------------------------------------------------
' Exit Button
'------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    Unload Me
End Sub

'------------------------------------------------------------------------------
' ParseDate - Helper to parse DD/MM/YYYY date strings (locale-safe)
'------------------------------------------------------------------------------
Private Function ParseDate(ByVal sDate As String) As Date
    ParseDate = ParseDateDMY(sDate)
End Function
