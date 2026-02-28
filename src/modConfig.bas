Attribute VB_Name = "modConfig"
'==============================================================================
' modConfig - Configuration & Settings Module
' APP Billing System
'
' Manages application settings stored in the VeryHidden "Settings" sheet.
' Provides helpers for network path, user identity, and folder structure.
'==============================================================================
Option Explicit

' Settings sheet cell locations
Private Const SETTINGS_SHEET As String = "Settings"
Private Const ROW_NETWORK_PATH As Long = 1
Private Const ROW_SUPERUSER_PWD As Long = 2
Private Const ROW_DEFAULT_SITE As Long = 3
Private Const COL_KEY As Long = 1   ' Column A
Private Const COL_VALUE As Long = 2 ' Column B

' Network subfolder names
Public Const FOLDER_DATA As String = "Data"
Public Const FOLDER_DAILY_EXPORTS As String = "DailyExports"
Public Const FOLDER_PDF_REPORTS As String = "PDFReports"
Public Const FOLDER_CONFIG As String = "Config"

'------------------------------------------------------------------------------
' GetNetworkPath - Returns the configured network share path
'------------------------------------------------------------------------------
Public Function GetNetworkPath() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    GetNetworkPath = Trim(CStr(ws.Cells(ROW_NETWORK_PATH, COL_VALUE).Value))

    ' Ensure trailing backslash
    If Len(GetNetworkPath) > 0 And Right(GetNetworkPath, 1) <> "\" Then
        GetNetworkPath = GetNetworkPath & "\"
    End If
    Exit Function
ErrHandler:
    GetNetworkPath = ""
End Function

'------------------------------------------------------------------------------
' SetNetworkPath - Updates the network share path in Settings
'------------------------------------------------------------------------------
Public Sub SetNetworkPath(ByVal sPath As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)

    ' Ensure trailing backslash
    If Len(sPath) > 0 And Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

    ws.Cells(ROW_NETWORK_PATH, COL_VALUE).Value = sPath
    Exit Sub
ErrHandler:
    MsgBox "Error saving network path: " & Err.Description, vbCritical, "Configuration Error"
End Sub

'------------------------------------------------------------------------------
' GetSuperUserPassword - Returns the stored superuser password hash
'------------------------------------------------------------------------------
Public Function GetSuperUserPassword() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    GetSuperUserPassword = CStr(ws.Cells(ROW_SUPERUSER_PWD, COL_VALUE).Value)
    Exit Function
ErrHandler:
    GetSuperUserPassword = ""
End Function

'------------------------------------------------------------------------------
' SetSuperUserPassword - Stores the superuser password (hashed)
'------------------------------------------------------------------------------
Public Sub SetSuperUserPassword(ByVal sPassword As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    ws.Cells(ROW_SUPERUSER_PWD, COL_VALUE).Value = SimpleHash(sPassword)
    Exit Sub
ErrHandler:
    MsgBox "Error saving password: " & Err.Description, vbCritical, "Configuration Error"
End Sub

'------------------------------------------------------------------------------
' SimpleHash - Basic hash function for password storage
' Note: This is not cryptographically secure but provides obfuscation.
' For production use, consider a more robust hashing approach.
'------------------------------------------------------------------------------
Public Function SimpleHash(ByVal sText As String) As String
    Dim i As Long
    Dim lHash As Long
    lHash = 5381
    For i = 1 To Len(sText)
        lHash = ((lHash * 33) Xor Asc(Mid(sText, i, 1))) And &H7FFFFFFF
    Next i
    SimpleHash = CStr(lHash)
End Function

'------------------------------------------------------------------------------
' GetDefaultSite - Returns the default hospital site
'------------------------------------------------------------------------------
Public Function GetDefaultSite() As String
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    GetDefaultSite = CStr(ws.Cells(ROW_DEFAULT_SITE, COL_VALUE).Value)
    If Len(GetDefaultSite) = 0 Then GetDefaultSite = "RCH"
    Exit Function
ErrHandler:
    GetDefaultSite = "RCH"
End Function

'------------------------------------------------------------------------------
' GetCurrentUser - Returns the Windows login username (sanitized for filenames)
'------------------------------------------------------------------------------
Public Function GetCurrentUser() As String
    Dim sUser As String
    sUser = Environ("USERNAME")
    If Len(sUser) = 0 Then sUser = Application.UserName

    ' Sanitize for use in filenames - remove invalid characters
    Dim i As Long
    Dim sClean As String
    Dim c As String
    For i = 1 To Len(sUser)
        c = Mid(sUser, i, 1)
        Select Case c
            Case "A" To "Z", "a" To "z", "0" To "9", "_", "-", ".", " "
                sClean = sClean & c
            Case Else
                sClean = sClean & "_"
        End Select
    Next i

    GetCurrentUser = sClean
End Function

'------------------------------------------------------------------------------
' GetCurrentUserDisplayName - Returns Application.UserName for display
'------------------------------------------------------------------------------
Public Function GetCurrentUserDisplayName() As String
    GetCurrentUserDisplayName = Application.UserName
End Function

'------------------------------------------------------------------------------
' IsNetworkAvailable - Checks if the network share is accessible
'------------------------------------------------------------------------------
Public Function IsNetworkAvailable() As Boolean
    On Error GoTo ErrHandler
    Dim sPath As String
    sPath = GetNetworkPath()

    If Len(sPath) = 0 Then
        IsNetworkAvailable = False
        Exit Function
    End If

    IsNetworkAvailable = (Dir(sPath, vbDirectory) <> "")
    Exit Function
ErrHandler:
    IsNetworkAvailable = False
End Function

'------------------------------------------------------------------------------
' EnsureNetworkFolders - Creates the required folder structure on the share
'------------------------------------------------------------------------------
Public Function EnsureNetworkFolders() As Boolean
    On Error GoTo ErrHandler
    Dim sBase As String
    sBase = GetNetworkPath()

    If Len(sBase) = 0 Then
        EnsureNetworkFolders = False
        Exit Function
    End If

    ' Create main subfolders
    CreateFolderIfNotExists sBase & FOLDER_DATA
    CreateFolderIfNotExists sBase & FOLDER_DAILY_EXPORTS
    CreateFolderIfNotExists sBase & FOLDER_PDF_REPORTS
    CreateFolderIfNotExists sBase & FOLDER_CONFIG

    EnsureNetworkFolders = True
    Exit Function
ErrHandler:
    EnsureNetworkFolders = False
End Function

'------------------------------------------------------------------------------
' EnsureMonthFolder - Creates the monthly subfolder under Data (YYYY-MM)
'------------------------------------------------------------------------------
Public Function EnsureMonthFolder(ByVal dtDate As Date) As String
    Dim sBase As String
    sBase = GetNetworkPath()
    If Len(sBase) = 0 Then
        EnsureMonthFolder = ""
        Exit Function
    End If

    Dim sMonth As String
    sMonth = Format(dtDate, "YYYY-MM")

    Dim sPath As String
    sPath = sBase & FOLDER_DATA & "\" & sMonth
    CreateFolderIfNotExists sPath

    EnsureMonthFolder = sPath & "\"
End Function

'------------------------------------------------------------------------------
' GetUserDailyFileName - Returns the filename for a user's daily data file
'------------------------------------------------------------------------------
Public Function GetUserDailyFileName(ByVal sUserName As String, ByVal dtDate As Date) As String
    ' Sanitize username for filename
    Dim sClean As String
    sClean = Replace(sUserName, " ", "_")
    sClean = Replace(sClean, ",", "")

    GetUserDailyFileName = sClean & "_" & Format(dtDate, "YYYYMMDD") & ".xlsx"
End Function

'------------------------------------------------------------------------------
' GetUserDailyFilePath - Returns the full path for a user's daily data file
'------------------------------------------------------------------------------
Public Function GetUserDailyFilePath(ByVal sUserName As String, ByVal dtDate As Date) As String
    Dim sMonthFolder As String
    sMonthFolder = EnsureMonthFolder(dtDate)

    If Len(sMonthFolder) = 0 Then
        GetUserDailyFilePath = ""
        Exit Function
    End If

    GetUserDailyFilePath = sMonthFolder & GetUserDailyFileName(sUserName, dtDate)
End Function

'------------------------------------------------------------------------------
' CreateFolderIfNotExists - Creates a folder (and parent folders) if missing
'------------------------------------------------------------------------------
Public Sub CreateFolderIfNotExists(ByVal sPath As String)
    On Error Resume Next
    If Dir(sPath, vbDirectory) = "" Then
        ' Use MkDir for each level - build path incrementally
        Dim parts() As String
        Dim buildPath As String
        Dim i As Long

        ' Handle UNC paths
        If Left(sPath, 2) = "\\" Then
            parts = Split(Mid(sPath, 3), "\")
            buildPath = "\\" & parts(0) & "\" & parts(1)
            ' Start from index 2 (skip server and share)
            For i = 2 To UBound(parts)
                If Len(parts(i)) > 0 Then
                    buildPath = buildPath & "\" & parts(i)
                    If Dir(buildPath, vbDirectory) = "" Then
                        MkDir buildPath
                    End If
                End If
            Next i
        Else
            parts = Split(sPath, "\")
            buildPath = parts(0)
            For i = 1 To UBound(parts)
                If Len(parts(i)) > 0 Then
                    buildPath = buildPath & "\" & parts(i)
                    If Dir(buildPath, vbDirectory) = "" Then
                        MkDir buildPath
                    End If
                End If
            Next i
        End If
    End If
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' InitializeSettingsSheet - Creates the Settings sheet if it doesn't exist
' Call this once during initial setup
'------------------------------------------------------------------------------
Public Sub InitializeSettingsSheet()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = SETTINGS_SHEET
    End If

    ' Set up keys if empty
    If Len(ws.Cells(ROW_NETWORK_PATH, COL_KEY).Value) = 0 Then
        ws.Cells(ROW_NETWORK_PATH, COL_KEY).Value = "NetworkSharePath"
        ws.Cells(ROW_NETWORK_PATH, COL_VALUE).Value = "\\server\APP_Billing\"
    End If

    If Len(ws.Cells(ROW_SUPERUSER_PWD, COL_KEY).Value) = 0 Then
        ws.Cells(ROW_SUPERUSER_PWD, COL_KEY).Value = "SuperUserPassword"
        ws.Cells(ROW_SUPERUSER_PWD, COL_VALUE).Value = ""
    End If

    If Len(ws.Cells(ROW_DEFAULT_SITE, COL_KEY).Value) = 0 Then
        ws.Cells(ROW_DEFAULT_SITE, COL_KEY).Value = "DefaultSite"
        ws.Cells(ROW_DEFAULT_SITE, COL_VALUE).Value = "RCH"
    End If

    ' Make the sheet VeryHidden (only accessible via VBA)
    ws.Visible = xlSheetVeryHidden
End Sub

'------------------------------------------------------------------------------
' ShowConfigDialog - Prompts user to configure the network path
'------------------------------------------------------------------------------
Public Sub ShowConfigDialog()
    Dim sPath As String
    sPath = GetNetworkPath()

    sPath = InputBox("Enter the network share path for APP Billing data storage:" & vbCrLf & vbCrLf & _
                     "Example: \\server\APP_Billing\" & vbCrLf & vbCrLf & _
                     "Current path: " & IIf(Len(sPath) > 0, sPath, "(not set)"), _
                     "Network Path Configuration", sPath)

    If Len(sPath) > 0 Then
        SetNetworkPath sPath

        If IsNetworkAvailable() Then
            EnsureNetworkFolders
            MsgBox "Network path configured successfully." & vbCrLf & _
                   "Folder structure verified.", vbInformation, "Configuration"
        Else
            MsgBox "Warning: The specified path is not currently accessible." & vbCrLf & _
                   "The path has been saved but please verify the network connection.", _
                   vbExclamation, "Configuration Warning"
        End If
    End If
End Sub
