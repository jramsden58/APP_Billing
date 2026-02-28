Attribute VB_Name = "modSecurity"
'==============================================================================
' modSecurity - Superuser Access Control Module
' APP Billing System
'
' Provides authentication combining Windows username verification against
' a SuperUsers.xlsx file on the network share AND password verification.
'==============================================================================
Option Explicit

' Column layout in SuperUsers.xlsx
Private Const SU_COL_USERNAME As Long = 1    ' Column A: Windows Username
Private Const SU_COL_DISPLAYNAME As Long = 2 ' Column B: Display Name
Private Const SU_COL_ACCESS As Long = 3      ' Column C: Access Level (Admin/ReadOnly)

' Cache for authentication state during session
Private m_bAuthenticated As Boolean
Private m_sAccessLevel As String

'------------------------------------------------------------------------------
' IsSuperUser - Checks if current Windows user is in the SuperUsers list
'------------------------------------------------------------------------------
Public Function IsSuperUser() As Boolean
    On Error GoTo ErrHandler

    Dim sCurrentUser As String
    sCurrentUser = LCase(Environ("USERNAME"))

    Dim sSuperUsersPath As String
    sSuperUsersPath = GetNetworkPath() & FOLDER_CONFIG & "\SuperUsers.xlsx"

    ' Check if file exists
    If Dir(sSuperUsersPath) = "" Then
        IsSuperUser = False
        Exit Function
    End If

    ' Open the SuperUsers file silently
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(sSuperUsersPath, ReadOnly:=True, UpdateLinks:=0)

    Dim ws As Worksheet
    Set ws = wb.Sheets(1)

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, SU_COL_USERNAME).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow ' Skip header row
        If LCase(Trim(ws.Cells(i, SU_COL_USERNAME).Value)) = sCurrentUser Then
            m_sAccessLevel = Trim(ws.Cells(i, SU_COL_ACCESS).Value)
            IsSuperUser = True
            wb.Close SaveChanges:=False
            Application.ScreenUpdating = True
            Exit Function
        End If
    Next i

    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    IsSuperUser = False
    Exit Function

ErrHandler:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    IsSuperUser = False
End Function

'------------------------------------------------------------------------------
' AuthenticateSuperUser - Full authentication: username check + password
' Returns True if authentication succeeds
'------------------------------------------------------------------------------
Public Function AuthenticateSuperUser() As Boolean
    ' Already authenticated this session?
    If m_bAuthenticated Then
        AuthenticateSuperUser = True
        Exit Function
    End If

    ' Step 1: Check Windows username against SuperUsers list
    If Not IsSuperUser() Then
        MsgBox "Access denied." & vbCrLf & vbCrLf & _
               "Your Windows account (" & Environ("USERNAME") & ") is not authorized " & _
               "for superuser access." & vbCrLf & vbCrLf & _
               "Contact an administrator to be added to the SuperUsers list.", _
               vbCritical, "Authentication Failed"
        AuthenticateSuperUser = False
        Exit Function
    End If

    ' Step 2: Prompt for password
    Dim sStoredHash As String
    sStoredHash = GetSuperUserPassword()

    ' If no password has been set yet, prompt to create one
    If Len(sStoredHash) = 0 Then
        Dim sNewPwd As String
        sNewPwd = InputBox("No superuser password has been set." & vbCrLf & _
                          "Please create a superuser password:", _
                          "Set Superuser Password")
        If Len(sNewPwd) = 0 Then
            AuthenticateSuperUser = False
            Exit Function
        End If

        Dim sConfirm As String
        sConfirm = InputBox("Please confirm the password:", "Confirm Password")
        If sNewPwd <> sConfirm Then
            MsgBox "Passwords do not match. Please try again.", vbExclamation, "Password Mismatch"
            AuthenticateSuperUser = False
            Exit Function
        End If

        SetSuperUserPassword sNewPwd
        m_bAuthenticated = True
        AuthenticateSuperUser = True
        MsgBox "Superuser password has been set successfully.", vbInformation, "Password Set"
        Exit Function
    End If

    ' Verify password (3 attempts)
    Dim iAttempts As Long
    For iAttempts = 1 To 3
        Dim sInput As String
        sInput = InputBox("Enter superuser password:" & vbCrLf & _
                         "Attempt " & iAttempts & " of 3", _
                         "Superuser Authentication")

        If Len(sInput) = 0 Then
            AuthenticateSuperUser = False
            Exit Function
        End If

        If SimpleHash(sInput) = sStoredHash Then
            m_bAuthenticated = True
            AuthenticateSuperUser = True
            Exit Function
        Else
            If iAttempts < 3 Then
                MsgBox "Incorrect password. Please try again.", vbExclamation, "Authentication Failed"
            End If
        End If
    Next iAttempts

    MsgBox "Maximum authentication attempts exceeded.", vbCritical, "Access Denied"
    AuthenticateSuperUser = False
End Function

'------------------------------------------------------------------------------
' GetAccessLevel - Returns the access level of the current superuser
'------------------------------------------------------------------------------
Public Function GetAccessLevel() As String
    If Len(m_sAccessLevel) > 0 Then
        GetAccessLevel = m_sAccessLevel
    Else
        GetAccessLevel = "None"
    End If
End Function

'------------------------------------------------------------------------------
' IsAdmin - Returns True if current user has Admin access level
'------------------------------------------------------------------------------
Public Function IsAdmin() As Boolean
    If Not m_bAuthenticated Then
        IsAdmin = False
        Exit Function
    End If
    IsAdmin = (LCase(m_sAccessLevel) = "admin")
End Function

'------------------------------------------------------------------------------
' LogOut - Clears authentication state
'------------------------------------------------------------------------------
Public Sub LogOut()
    m_bAuthenticated = False
    m_sAccessLevel = ""
End Sub

'------------------------------------------------------------------------------
' IsAuthenticated - Returns current authentication state
'------------------------------------------------------------------------------
Public Function IsAuthenticated() As Boolean
    IsAuthenticated = m_bAuthenticated
End Function

'------------------------------------------------------------------------------
' CreateSuperUsersFile - Creates the initial SuperUsers.xlsx on the network
' Call once during initial system setup
'------------------------------------------------------------------------------
Public Sub CreateSuperUsersFile()
    On Error GoTo ErrHandler

    Dim sPath As String
    sPath = GetNetworkPath() & FOLDER_CONFIG & "\SuperUsers.xlsx"

    ' Check if already exists
    If Dir(sPath) <> "" Then
        If MsgBox("SuperUsers.xlsx already exists. Overwrite?", _
                  vbYesNo + vbQuestion, "Confirm") = vbNo Then
            Exit Sub
        End If
    End If

    ' Create new workbook
    Dim wb As Workbook
    Set wb = Workbooks.Add(xlWBATWorksheet)

    Dim ws As Worksheet
    Set ws = wb.Sheets(1)
    ws.Name = "SuperUsers"

    ' Headers
    ws.Cells(1, SU_COL_USERNAME).Value = "Windows Username"
    ws.Cells(1, SU_COL_DISPLAYNAME).Value = "Display Name"
    ws.Cells(1, SU_COL_ACCESS).Value = "Access Level"

    ' Format headers
    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Add the current user as the first admin
    ws.Cells(2, SU_COL_USERNAME).Value = Environ("USERNAME")
    ws.Cells(2, SU_COL_DISPLAYNAME).Value = Application.UserName
    ws.Cells(2, SU_COL_ACCESS).Value = "Admin"

    ' Auto-fit columns
    ws.Columns("A:C").AutoFit

    ' Add data validation for Access Level
    With ws.Range("C2:C100").Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Admin,ReadOnly"
    End With

    ' Ensure config folder exists
    CreateFolderIfNotExists GetNetworkPath() & FOLDER_CONFIG

    ' Save
    Application.DisplayAlerts = False
    wb.SaveAs sPath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    wb.Close SaveChanges:=False

    MsgBox "SuperUsers.xlsx created successfully at:" & vbCrLf & sPath & vbCrLf & vbCrLf & _
           "Your account (" & Environ("USERNAME") & ") has been added as Admin.", _
           vbInformation, "Setup Complete"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "Error creating SuperUsers file: " & Err.Description, vbCritical, "Setup Error"
End Sub

'------------------------------------------------------------------------------
' ChangeSuperUserPassword - Allows an authenticated admin to change the password
'------------------------------------------------------------------------------
Public Sub ChangeSuperUserPassword()
    If Not AuthenticateSuperUser() Then Exit Sub
    If Not IsAdmin() Then
        MsgBox "Only administrators can change the superuser password.", _
               vbExclamation, "Access Denied"
        Exit Sub
    End If

    Dim sNewPwd As String
    sNewPwd = InputBox("Enter new superuser password:", "Change Password")
    If Len(sNewPwd) = 0 Then Exit Sub

    Dim sConfirm As String
    sConfirm = InputBox("Confirm new password:", "Change Password")
    If sNewPwd <> sConfirm Then
        MsgBox "Passwords do not match.", vbExclamation, "Password Mismatch"
        Exit Sub
    End If

    SetSuperUserPassword sNewPwd
    MsgBox "Superuser password changed successfully.", vbInformation, "Password Changed"
End Sub
