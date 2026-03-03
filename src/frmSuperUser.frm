VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSuperUser
   Caption         =   "APP Billing - Superuser Access"
   ClientHeight    =   8000
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
'   - Search across all fields with date range filtering
'   - Browse and open data files on the network share
'   - Run daily consolidation
'   - Run bi-weekly / date range consolidation
'   - Manage SuperUsers list (Admin only)
'   - Configure system settings (Admin only)
'   - Change superuser password (Admin only)
'
' FORM CONTROLS NEEDED (create in VBA Editor form designer):
'   lblTitle       - Label "Superuser Access Panel"
'   lblUser        - Label showing current user and access level
'   lblStatus      - Label for status messages
'
'   ' Search Frame
'   fraSearch      - Frame "Search Database"
'   txtSearchTerm  - TextBox for search text
'   txtSearchDateFrom - TextBox for date range start (DD/MM/YYYY)
'   txtSearchDateTo   - TextBox for date range end (DD/MM/YYYY)
'   cboSearchField - ComboBox for field selection (or "All Fields")
'   cmdSearch      - CommandButton "Search"
'   cmdClearSearch - CommandButton "Clear"
'
'   ' Data Browsing Frame
'   fraDataBrowse  - Frame "Browse Data Files"
'   txtBrowseDate  - TextBox for date (DD/MM/YYYY)
'   cmdListFiles   - CommandButton "List Files"
'   lstFiles       - ListBox showing available files
'   cmdOpenFile    - CommandButton "Open Selected"
'
'   ' Consolidation Frame
'   fraConsolidate - Frame "Bi-Weekly / Date Range Consolidation"
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
    On Error GoTo ErrHandler

    ' Disable all controls before authentication to prevent bypass
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        On Error Resume Next
        ctrl.Enabled = False
        On Error GoTo 0
    Next ctrl

    ' Attempt authentication
    On Error GoTo ErrHandler
    If Not AuthenticateSuperUser() Then
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

    ' Set up controls - all wrapped in On Error Resume Next
    ' so missing controls won't crash the form
    On Error Resume Next

    ' Set up user display
    lblUser.Caption = "User: " & Application.UserName & " | Access: " & GetAccessLevel()

    ' Set default dates
    txtBrowseDate.Value = Format(Date, "DD/MM/YYYY")
    txtConsDate.Value = Format(Date, "DD/MM/YYYY")
    txtStartDate.Value = Format(Date - 7, "DD/MM/YYYY")
    txtEndDate.Value = Format(Date, "DD/MM/YYYY")

    ' Set up search controls
    txtSearchTerm.Value = ""
    txtSearchDateFrom.Value = Format(Date - 30, "DD/MM/YYYY")
    txtSearchDateTo.Value = Format(Date, "DD/MM/YYYY")

    ' Populate search field dropdown
    cboSearchField.Clear
    cboSearchField.AddItem "All Fields"
    cboSearchField.AddItem "Anesthesiologist"
    cboSearchField.AddItem "Site"
    cboSearchField.AddItem "Date of Service"
    cboSearchField.AddItem "Shift Name"
    cboSearchField.AddItem "On Call"
    cboSearchField.AddItem "Shift Type"
    cboSearchField.AddItem "Procedure Code"
    cboSearchField.AddItem "Start Time"
    cboSearchField.AddItem "Finish Time"
    cboSearchField.AddItem "IC Level"
    cboSearchField.AddItem "Consults"
    cboSearchField.AddItem "Fee Modifier 1"
    cboSearchField.AddItem "Fee Modifier 2"
    cboSearchField.AddItem "Fee Modifier 3"
    cboSearchField.AddItem "Resuscitation"
    cboSearchField.AddItem "Obstetrics"
    cboSearchField.AddItem "Acute Pain"
    cboSearchField.AddItem "Chronic Pain"
    cboSearchField.AddItem "Miscellaneous"
    cboSearchField.AddItem "WCB Number"
    cboSearchField.AddItem "Side"
    cboSearchField.AddItem "Diagnostic Code"
    cboSearchField.AddItem "Injury Code"
    cboSearchField.AddItem "Date of Injury"
    cboSearchField.AddItem "Submitted By"
    cboSearchField.AddItem "Submitted On"
    cboSearchField.ListIndex = 0 ' Default to "All Fields"

    ' Show/hide admin features based on access level
    fraAdmin.Visible = IsAdmin()

    lblStatus.Caption = "Authenticated successfully."
    Set m_colFiles = New Collection

    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Error initializing SuperUser form:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Ensure all form controls exist in the VBA Editor form designer." & vbCrLf & _
           "See the control list at the top of frmSuperUser.frm for required controls.", _
           vbCritical, "SuperUser Form Error"
    Application.OnTime Now, "UnloadSuperUser"
End Sub

'------------------------------------------------------------------------------
' Search Button - Searches across all fields in network data files
'------------------------------------------------------------------------------
Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler

    Dim sSearchTerm As String
    sSearchTerm = Trim(txtSearchTerm.Value)

    If Len(sSearchTerm) = 0 Then
        MsgBox "Please enter a search term.", vbExclamation, "Validation"
        Exit Sub
    End If

    ' Parse date range
    Dim dtFrom As Date, dtTo As Date
    dtFrom = ParseDate(txtSearchDateFrom.Value)
    dtTo = ParseDate(txtSearchDateTo.Value)

    If dtFrom > dtTo Then
        MsgBox "Start date must be before end date.", vbExclamation, "Validation"
        Exit Sub
    End If

    ' Determine which column to search (0 = all)
    Dim lSearchCol As Long
    lSearchCol = GetSearchColumnIndex(cboSearchField.ListIndex)

    lblStatus.Caption = "Searching..."
    DoEvents

    ' Ensure SearchData sheet exists and clear it
    Dim wsSearch As Worksheet
    Set wsSearch = EnsureSheetExists("SearchData")
    wsSearch.Cells.ClearContents

    ' Write headers
    Dim headers As Variant
    headers = Array("S #", "Anesthesiologist", "Site", "Date of Service", _
                   "Shift Name", "On Call", "Shift Type", "Surgical Procedure Code", _
                   "Procedure Start Time", "Procedure Finish Time", "Maximum IC Level", _
                   "Consults", "Fee Modifier 1", "Fee Modifier 2", "Fee Modifier 3", _
                   "Resuscitation", "Obstetrics", "Acute Pain", _
                   "Diagnostic and Chronic Pain", "Miscellaneous Fee Items", _
                   "WCB Number", "Side", "Diagnostic Code", "Injury Code", _
                   "Date of Injury", "Submitted By", "Submitted On", "Source File")
    Dim h As Long
    For h = 0 To UBound(headers)
        wsSearch.Cells(1, h + 1).Value = headers(h)
    Next h

    ' Format headers
    With wsSearch.Range(wsSearch.Cells(1, 1), wsSearch.Cells(1, 28))
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    Dim lOutRow As Long
    lOutRow = 2

    ' Search across all dates in range
    Dim dtCurrent As Date
    For dtCurrent = dtFrom To dtTo
        ' Get all user files for this date
        Dim colData As Collection
        Set colData = ReadAllUsersDailyData(dtCurrent)

        ' Search through each user's data
        Dim vDataSet As Variant
        For Each vDataSet In colData
            If IsArray(vDataSet) Then
                Dim lRows As Long
                ' Handle both single-row and multi-row arrays
                On Error Resume Next
                lRows = UBound(vDataSet, 1)
                On Error GoTo ErrHandler

                Dim r As Long
                For r = 1 To lRows
                    Dim bMatch As Boolean
                    bMatch = False

                    If lSearchCol = 0 Then
                        ' Search all columns
                        Dim c As Long
                        For c = 1 To NUM_COLUMNS
                            If InStr(1, CStr(vDataSet(r, c)), sSearchTerm, vbTextCompare) > 0 Then
                                bMatch = True
                                Exit For
                            End If
                        Next c
                    Else
                        ' Search specific column
                        If InStr(1, CStr(vDataSet(r, lSearchCol)), sSearchTerm, vbTextCompare) > 0 Then
                            bMatch = True
                        End If
                    End If

                    If bMatch Then
                        ' Write matching row to SearchData
                        For c = 1 To NUM_COLUMNS - 1 ' Skip SyncStatus (col 28)
                            wsSearch.Cells(lOutRow, c).Value = vDataSet(r, c)
                        Next c
                        ' Column 28: Source date
                        wsSearch.Cells(lOutRow, 28).Value = Format(dtCurrent, "DD/MM/YYYY")
                        lOutRow = lOutRow + 1
                    End If
                Next r
            End If
        Next vDataSet
    Next dtCurrent

    ' Also search local DailyDatabase
    Dim wsLocal As Worksheet
    Set wsLocal = ThisWorkbook.Sheets("DailyDatabase")
    Dim lastRow As Long
    lastRow = wsLocal.Cells(wsLocal.Rows.Count, COL_ANESTH).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        ' Check if record date is in range
        Dim dtRecDate As Date
        Dim sRecDate As String
        sRecDate = CStr(wsLocal.Cells(i, COL_DATE).Value)
        If TryParseDateDMY(sRecDate, dtRecDate) Then
            If dtRecDate >= dtFrom And dtRecDate <= dtTo Then
                Dim bLocalMatch As Boolean
                bLocalMatch = False

                If lSearchCol = 0 Then
                    Dim lc As Long
                    For lc = 1 To NUM_COLUMNS
                        If InStr(1, CStr(wsLocal.Cells(i, lc).Value), sSearchTerm, vbTextCompare) > 0 Then
                            bLocalMatch = True
                            Exit For
                        End If
                    Next lc
                Else
                    If InStr(1, CStr(wsLocal.Cells(i, lSearchCol).Value), sSearchTerm, vbTextCompare) > 0 Then
                        bLocalMatch = True
                    End If
                End If

                If bLocalMatch Then
                    For lc = 1 To NUM_COLUMNS - 1
                        wsSearch.Cells(lOutRow, lc).Value = wsLocal.Cells(i, lc).Value
                    Next lc
                    wsSearch.Cells(lOutRow, 28).Value = "Local"
                    lOutRow = lOutRow + 1
                End If
            End If
        End If
    Next i

    ' Apply AutoFilter
    If lOutRow > 2 Then
        wsSearch.Range("A1").AutoFilter
        wsSearch.Columns.AutoFit
    End If

    If lOutRow = 2 Then
        lblStatus.Caption = "No records found."
        MsgBox "No records found matching '" & sSearchTerm & "'.", _
               vbInformation, "Search Results"
    Else
        lblStatus.Caption = (lOutRow - 2) & " record(s) found."
        MsgBox (lOutRow - 2) & " record(s) found. Results are on the SearchData sheet.", _
               vbInformation, "Search Results"
        ThisWorkbook.Sheets("SearchData").Activate
    End If

    Exit Sub
ErrHandler:
    lblStatus.Caption = "Search failed."
    MsgBox "Search error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Clear Search Button
'------------------------------------------------------------------------------
Private Sub cmdClearSearch_Click()
    On Error Resume Next
    txtSearchTerm.Value = ""
    cboSearchField.ListIndex = 0
    txtSearchDateFrom.Value = Format(Date - 30, "DD/MM/YYYY")
    txtSearchDateTo.Value = Format(Date, "DD/MM/YYYY")
    lblStatus.Caption = "Search cleared."
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' GetSearchColumnIndex - Maps ComboBox selection to column index
' Returns 0 for "All Fields"
'------------------------------------------------------------------------------
Private Function GetSearchColumnIndex(ByVal lCboIndex As Long) As Long
    Select Case lCboIndex
        Case 0: GetSearchColumnIndex = 0   ' All Fields
        Case 1: GetSearchColumnIndex = COL_ANESTH
        Case 2: GetSearchColumnIndex = COL_SITE
        Case 3: GetSearchColumnIndex = COL_DATE
        Case 4: GetSearchColumnIndex = COL_SHIFT
        Case 5: GetSearchColumnIndex = COL_ONCALL
        Case 6: GetSearchColumnIndex = COL_SHIFTTYPE
        Case 7: GetSearchColumnIndex = COL_PROCCODE
        Case 8: GetSearchColumnIndex = COL_STARTTIME
        Case 9: GetSearchColumnIndex = COL_FINTIME
        Case 10: GetSearchColumnIndex = COL_MAXIC
        Case 11: GetSearchColumnIndex = COL_CONSULT
        Case 12: GetSearchColumnIndex = COL_MOD1
        Case 13: GetSearchColumnIndex = COL_MOD2
        Case 14: GetSearchColumnIndex = COL_MOD3
        Case 15: GetSearchColumnIndex = COL_RESUS
        Case 16: GetSearchColumnIndex = COL_OBS
        Case 17: GetSearchColumnIndex = COL_ACUTEPAIN
        Case 18: GetSearchColumnIndex = COL_CHRONPAIN
        Case 19: GetSearchColumnIndex = COL_MISC
        Case 20: GetSearchColumnIndex = COL_WCBNUM
        Case 21: GetSearchColumnIndex = COL_WCBSIDE
        Case 22: GetSearchColumnIndex = COL_WCBDIAG
        Case 23: GetSearchColumnIndex = COL_WCBINJ
        Case 24: GetSearchColumnIndex = COL_WCBDATE
        Case 25: GetSearchColumnIndex = COL_SUBMBY
        Case 26: GetSearchColumnIndex = COL_SUBMON
        Case Else: GetSearchColumnIndex = 0
    End Select
End Function

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
