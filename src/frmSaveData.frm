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
' Features: Filter-as-you-type list boxes, search, edit, delete
' Date format: DD/MM/YYYY (auto-formatted with slash insertion)
' Time format: HHMMhr (e.g., 0800hr) 24-hour clock
'
' List boxes start empty. As the user types, matching items appear
' (case-insensitive prefix match). Backspace removes last character,
' Escape clears the filter.
'==============================================================================
Option Explicit

' Module-level variable to track the row being edited (0 = not editing)
Private m_lEditRow As Long

' Flag to prevent recursive formatting in Change events
Private m_bFormatting As Boolean

' Master item arrays for filter-as-you-type list boxes
' Loaded once from LookupLists during Initialize, used to filter on each keypress
Private m_aAnesth() As String
Private m_aShftName() As String
Private m_aEval() As String
Private m_aMod1() As String
Private m_aMod2() As String
Private m_aMod3() As String
Private m_aResus() As String
Private m_aObs() As String
Private m_aAcPain() As String
Private m_aChPain() As String
Private m_aMisc() As String

' Per-listbox search text for filter-as-you-type
Private m_sSearchAnesth As String
Private m_sSearchShftName As String
Private m_sSearchEval As String
Private m_sSearchMod1 As String
Private m_sSearchMod2 As String
Private m_sSearchMod3 As String
Private m_sSearchResus As String
Private m_sSearchObs As String
Private m_sSearchAcPain As String
Private m_sSearchChPain As String
Private m_sSearchMisc As String

'------------------------------------------------------------------------------
' Form Initialize - Sets up the form with default values
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    m_lEditRow = 0
    m_bFormatting = False

    ' Disable built-in match behavior - we handle filtering ourselves
    On Error Resume Next
    lstAnesth.MatchEntry = fmMatchEntryNone
    lstEval.MatchEntry = fmMatchEntryNone
    lstMod1.MatchEntry = fmMatchEntryNone
    lstMod2.MatchEntry = fmMatchEntryNone
    lstMod3.MatchEntry = fmMatchEntryNone
    lstResus.MatchEntry = fmMatchEntryNone
    lstObs.MatchEntry = fmMatchEntryNone
    lstAcPain.MatchEntry = fmMatchEntryNone
    lstChPain.MatchEntry = fmMatchEntryNone
    lstMisc.MatchEntry = fmMatchEntryNone
    lstShftName.MatchEntry = fmMatchEntryNone
    On Error GoTo 0

    ' Load master item arrays from LookupLists sheet
    LoadMasterLists

    ' Clear all search strings
    ClearAllSearchText

    ' Clear all list boxes so they start empty
    ClearAllListBoxes

    Call Reset
End Sub

'------------------------------------------------------------------------------
' LoadMasterLists - Loads all list items from LookupLists into arrays
'------------------------------------------------------------------------------
Private Sub LoadMasterLists()
    On Error Resume Next
    Dim wsLookup As Worksheet
    Set wsLookup = ThisWorkbook.Sheets("LookupLists")
    If wsLookup Is Nothing Then Exit Sub
    On Error GoTo 0

    m_aAnesth = LoadColumnToArray(wsLookup, 1)      ' Column A: Anesthesiologists
    m_aShftName = LoadColumnToArray(wsLookup, 3)     ' Column C: Shift Names
    m_aEval = LoadColumnToArray(wsLookup, 4)         ' Column D: Consults/Eval
    m_aMod1 = LoadColumnToArray(wsLookup, 5)         ' Column E: Fee Modifier 1
    m_aMod2 = LoadColumnToArray(wsLookup, 6)         ' Column F: Fee Modifier 2
    m_aMod3 = LoadColumnToArray(wsLookup, 7)         ' Column G: Fee Modifier 3
    m_aResus = LoadColumnToArray(wsLookup, 8)        ' Column H: Resuscitation
    m_aObs = LoadColumnToArray(wsLookup, 9)          ' Column I: Obstetrics
    m_aAcPain = LoadColumnToArray(wsLookup, 10)      ' Column J: Acute Pain
    m_aChPain = LoadColumnToArray(wsLookup, 11)      ' Column K: Chronic Pain
    m_aMisc = LoadColumnToArray(wsLookup, 12)        ' Column L: Miscellaneous
End Sub

'------------------------------------------------------------------------------
' LoadColumnToArray - Reads non-empty cells from a column into a string array
'------------------------------------------------------------------------------
Private Function LoadColumnToArray(ByVal ws As Worksheet, ByVal lCol As Long) As String()
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, lCol).End(xlUp).Row

    Dim lCount As Long
    lCount = 0

    ' Count non-empty cells (skip header row 1)
    Dim i As Long
    For i = 2 To lastRow
        If Len(Trim(CStr(ws.Cells(i, lCol).Value))) > 0 Then
            lCount = lCount + 1
        End If
    Next i

    Dim result() As String
    If lCount = 0 Then
        ReDim result(0 To 0)
        result(0) = ""
        LoadColumnToArray = result
        Exit Function
    End If

    ReDim result(1 To lCount)
    Dim idx As Long
    idx = 0
    For i = 2 To lastRow
        Dim sVal As String
        sVal = Trim(CStr(ws.Cells(i, lCol).Value))
        If Len(sVal) > 0 Then
            idx = idx + 1
            result(idx) = sVal
        End If
    Next i

    LoadColumnToArray = result
End Function

'------------------------------------------------------------------------------
' ClearAllSearchText - Resets all search strings
'------------------------------------------------------------------------------
Private Sub ClearAllSearchText()
    m_sSearchAnesth = ""
    m_sSearchShftName = ""
    m_sSearchEval = ""
    m_sSearchMod1 = ""
    m_sSearchMod2 = ""
    m_sSearchMod3 = ""
    m_sSearchResus = ""
    m_sSearchObs = ""
    m_sSearchAcPain = ""
    m_sSearchChPain = ""
    m_sSearchMisc = ""
End Sub

'------------------------------------------------------------------------------
' ClearAllListBoxes - Clears all list boxes so they appear empty
'------------------------------------------------------------------------------
Private Sub ClearAllListBoxes()
    On Error Resume Next
    lstAnesth.Clear
    lstShftName.Clear
    lstEval.Clear
    lstMod1.Clear
    lstMod2.Clear
    lstMod3.Clear
    lstResus.Clear
    lstObs.Clear
    lstAcPain.Clear
    lstChPain.Clear
    lstMisc.Clear
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' FilterListBox - Filters a list box based on search text
' Shows items that match the search text (case-insensitive prefix match)
'------------------------------------------------------------------------------
Private Sub FilterListBox(ByRef lst As MSForms.ListBox, ByRef allItems() As String, _
                          ByVal sSearch As String)
    lst.Clear

    ' If search is empty, list stays empty (blank until user types)
    If Len(sSearch) = 0 Then Exit Sub

    ' Show items whose prefix matches the search text
    Dim i As Long
    For i = LBound(allItems) To UBound(allItems)
        If Len(allItems(i)) > 0 Then
            If LCase(Left(allItems(i), Len(sSearch))) = LCase(sSearch) Then
                lst.AddItem allItems(i)
            End If
        End If
    Next i

    ' Auto-select if only one match
    If lst.ListCount = 1 Then
        lst.ListIndex = 0
    End If
End Sub

'------------------------------------------------------------------------------
' PopulateFullList - Shows all items in a list box (used when loading for edit)
'------------------------------------------------------------------------------
Private Sub PopulateFullList(ByRef lst As MSForms.ListBox, ByRef allItems() As String)
    lst.Clear
    Dim i As Long
    For i = LBound(allItems) To UBound(allItems)
        If Len(allItems(i)) > 0 Then
            lst.AddItem allItems(i)
        End If
    Next i
End Sub

'==============================================================================
' LIST BOX KEYPRESS HANDLERS - Filter-as-you-type
' Each handler: appends typed char, handles Backspace/Escape, re-filters
'==============================================================================

Private Sub lstAnesth_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstAnesth, m_aAnesth, m_sSearchAnesth, KeyAscii
End Sub

Private Sub lstShftName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstShftName, m_aShftName, m_sSearchShftName, KeyAscii
End Sub

Private Sub lstEval_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstEval, m_aEval, m_sSearchEval, KeyAscii
End Sub

Private Sub lstMod1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstMod1, m_aMod1, m_sSearchMod1, KeyAscii
End Sub

Private Sub lstMod2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstMod2, m_aMod2, m_sSearchMod2, KeyAscii
End Sub

Private Sub lstMod3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstMod3, m_aMod3, m_sSearchMod3, KeyAscii
End Sub

Private Sub lstResus_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstResus, m_aResus, m_sSearchResus, KeyAscii
End Sub

Private Sub lstObs_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstObs, m_aObs, m_sSearchObs, KeyAscii
End Sub

Private Sub lstAcPain_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstAcPain, m_aAcPain, m_sSearchAcPain, KeyAscii
End Sub

Private Sub lstChPain_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstChPain, m_aChPain, m_sSearchChPain, KeyAscii
End Sub

Private Sub lstMisc_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleListKeyPress lstMisc, m_aMisc, m_sSearchMisc, KeyAscii
End Sub

'------------------------------------------------------------------------------
' HandleListKeyPress - Common handler for list box key presses
'------------------------------------------------------------------------------
Private Sub HandleListKeyPress(ByRef lst As MSForms.ListBox, ByRef allItems() As String, _
                               ByRef sSearch As String, ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 ' Backspace
            If Len(sSearch) > 0 Then
                sSearch = Left(sSearch, Len(sSearch) - 1)
            End If
        Case 27 ' Escape
            sSearch = ""
        Case Else
            sSearch = sSearch & Chr(KeyAscii)
    End Select

    FilterListBox lst, allItems, sSearch

    ' Consume the key so VBA doesn't try its own matching
    KeyAscii = 0
End Sub

'------------------------------------------------------------------------------
' Save Button - Saves data to local database and network share
'------------------------------------------------------------------------------
Private Sub cmdSave_Click()
    On Error GoTo ErrHandler

    ' Validate required fields
    If Not ValidateForm() Then Exit Sub

    ' Check if the date of service has been exported (locked)
    Dim sDateOfService As String
    sDateOfService = txtDteOfSer.Value
    If sDateOfService <> "DD/MM/YYYY" And Len(sDateOfService) > 0 Then
        Dim dtService As Date
        If TryParseDateDMY(sDateOfService, dtService) Then
            If IsDateExported(dtService) Then
                ' Block unless superuser
                If Not IsAuthenticated() Then
                    MsgBox "This date (" & sDateOfService & ") has been exported and locked." & vbCrLf & _
                           "No further changes are allowed for this date." & vbCrLf & vbCrLf & _
                           "Contact a superuser if changes are required.", _
                           vbExclamation, "Date Locked"
                    Exit Sub
                End If
            End If
        End If
    End If

    If MsgBox("Save this record?", vbYesNo + vbQuestion, "Confirm Save") = vbYes Then
        ' Save the new/updated record first
        If Not Submit() Then
            ' Submit failed (error already shown by Submit)
            Exit Sub
        End If

        ' If editing, delete the old record AFTER successful save
        If m_lEditRow > 0 Then
            Dim wsEdit As Worksheet
            Set wsEdit = ThisWorkbook.Sheets("DailyDatabase")
            ' Verify the row still exists
            If m_lEditRow <= wsEdit.Cells(wsEdit.Rows.Count, COL_ANESTH).End(xlUp).Row Then
                wsEdit.Rows(m_lEditRow).Delete
            End If
            m_lEditRow = 0
        End If

        Call Reset
        MsgBox "Record saved successfully.", vbInformation, "Saved"
    End If
    Exit Sub

ErrHandler:
    m_lEditRow = 0
    MsgBox "Error saving record: " & Err.Description, vbCritical, "Save Error"
End Sub

'------------------------------------------------------------------------------
' Exit Button - Closes the form
'------------------------------------------------------------------------------
Private Sub cmdExit_Click()
    ' Warn about unsaved edit
    Dim sWarning As String
    If m_lEditRow > 0 Then
        sWarning = "You are currently editing a record. " & _
                   "If you exit now, the original record will be preserved." & vbCrLf & vbCrLf & _
                   "Are you sure you want to exit?"
    Else
        sWarning = "Are you sure you want to exit? Any unsaved data will be lost."
    End If

    If MsgBox(sWarning, vbYesNo + vbQuestion, "Confirm Exit") = vbYes Then
        m_lEditRow = 0
        Unload Me
    End If
End Sub

'------------------------------------------------------------------------------
' Search Button - Searches records in DailyDatabase (superuser only)
'------------------------------------------------------------------------------
Private Sub cmdSearch_Click()
    On Error GoTo ErrHandler

    ' Search is restricted to superusers
    If Not AuthenticateSuperUser() Then
        MsgBox "Search functionality is restricted to superusers.", _
               vbExclamation, "Access Denied"
        Exit Sub
    End If

    Dim sSearchTerm As String
    sSearchTerm = InputBox("Enter search term:" & vbCrLf & vbCrLf & _
                          "Searches across Anesthesiologist, Date, and Procedure Code fields.", _
                          "Search Records")

    If Len(sSearchTerm) = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    ' Ensure SearchData sheet exists
    Dim wsSearch As Worksheet
    Set wsSearch = EnsureSheetExists("SearchData")
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
' Delete Button - Deletes the last entered record
' Regular users can only delete records submitted today.
' Past records require superuser authentication.
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

    ' Check if the record's date of service has been exported (locked)
    Dim sRecordDate As String
    sRecordDate = CStr(ws.Cells(lastRow, COL_DATE).Value)
    Dim dtRecordDate As Date
    If TryParseDateDMY(sRecordDate, dtRecordDate) Then
        If IsDateExported(dtRecordDate) Then
            If Not IsAuthenticated() Then
                MsgBox "This record's date (" & sRecordDate & ") has been exported and locked." & vbCrLf & _
                       "Contact a superuser to make changes.", _
                       vbExclamation, "Date Locked"
                Exit Sub
            End If
        End If
    End If

    ' Check if the record was submitted today (day-only editing restriction)
    If Not IsRecordFromToday(ws, lastRow) Then
        ' Past record - require superuser authentication
        If Not AuthenticateSuperUser() Then
            MsgBox "You can only delete records submitted today." & vbCrLf & _
                   "Superuser access is required to delete past records.", _
                   vbExclamation, "Access Denied"
            Exit Sub
        End If
    End If

    ' Show last record info for confirmation
    Dim sInfo As String
    sInfo = "Delete the last record?" & vbCrLf & vbCrLf & _
            "Anesthesiologist: " & ws.Cells(lastRow, COL_ANESTH).Value & vbCrLf & _
            "Date: " & ws.Cells(lastRow, COL_DATE).Value & vbCrLf & _
            "Procedure: " & ws.Cells(lastRow, COL_PROCCODE).Value & vbCrLf & _
            "Submitted: " & ws.Cells(lastRow, COL_SUBMON).Value & vbCrLf & vbCrLf & _
            "Note: This only deletes the local copy. Network copy is not affected."

    If MsgBox(sInfo, vbYesNo + vbExclamation, "Confirm Delete") = vbYes Then
        ws.Rows(lastRow).Delete
        MsgBox "Record deleted locally.", vbInformation, "Deleted"
    End If

    Exit Sub
ErrHandler:
    MsgBox "Delete error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' Edit Button - Loads the last record into the form for editing
' Regular users can only edit records submitted today.
' Past records require superuser authentication.
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

    ' Check if the record's date of service has been exported (locked)
    Dim sRecordDate As String
    sRecordDate = CStr(ws.Cells(lastRow, COL_DATE).Value)
    Dim dtRecordDate As Date
    If TryParseDateDMY(sRecordDate, dtRecordDate) Then
        If IsDateExported(dtRecordDate) Then
            If Not IsAuthenticated() Then
                MsgBox "This record's date (" & sRecordDate & ") has been exported and locked." & vbCrLf & _
                       "Contact a superuser to make changes.", _
                       vbExclamation, "Date Locked"
                Exit Sub
            End If
        End If
    End If

    ' Check if the record was submitted today (day-only editing restriction)
    If Not IsRecordFromToday(ws, lastRow) Then
        ' Past record - require superuser authentication
        If Not AuthenticateSuperUser() Then
            MsgBox "You can only edit records submitted today." & vbCrLf & _
                   "Superuser access is required to edit past records.", _
                   vbExclamation, "Access Denied"
            Exit Sub
        End If
    End If

    ' Temporarily disable formatting to avoid auto-format interference during load
    m_bFormatting = True

    ' Populate full lists so we can find items during edit load
    PopulateFullList lstAnesth, m_aAnesth
    PopulateFullList lstShftName, m_aShftName
    PopulateFullList lstEval, m_aEval
    PopulateFullList lstMod1, m_aMod1
    PopulateFullList lstMod2, m_aMod2
    PopulateFullList lstMod3, m_aMod3
    PopulateFullList lstResus, m_aResus
    PopulateFullList lstObs, m_aObs
    PopulateFullList lstAcPain, m_aAcPain
    PopulateFullList lstChPain, m_aChPain
    PopulateFullList lstMisc, m_aMisc

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

        ' Shift Name - find in list
        Dim sShift As String
        sShift = CStr(ws.Cells(lastRow, COL_SHIFT).Value)
        If Len(sShift) > 0 Then
            For k = 0 To .lstShftName.ListCount - 1
                If .lstShftName.List(k) = sShift Then
                    .lstShftName.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Shift type
        If CStr(ws.Cells(lastRow, COL_SHIFTTYPE).Value) = "OR" Then
            .optOR.Value = True
        Else
            .optOutOfOR.Value = True
        End If

        ' On Call (handle both Boolean and "Yes"/"No" string)
        Dim vOnCall As Variant
        vOnCall = ws.Cells(lastRow, COL_ONCALL).Value
        .chxOnCall.Value = (vOnCall = True Or LCase(CStr(vOnCall & "")) = "yes")

        ' Procedure fields
        .txtSurgProcCode.Value = CStr(ws.Cells(lastRow, COL_PROCCODE).Value)

        ' Start Time - convert legacy HH:MM to HHMMhr if needed
        Dim sTime As String
        sTime = CStr(ws.Cells(lastRow, COL_STARTTIME).Value)
        If InStr(sTime, ":") > 0 Then
            sTime = Replace(sTime, ":", "") & "hr"
        End If
        .txtProcStrtTime.Value = sTime

        ' Finish Time - convert legacy HH:MM to HHMMhr if needed
        sTime = CStr(ws.Cells(lastRow, COL_FINTIME).Value)
        If InStr(sTime, ":") > 0 Then
            sTime = Replace(sTime, ":", "") & "hr"
        End If
        .txtProcFinTime.Value = sTime

        .txtMaxIC.Value = CStr(ws.Cells(lastRow, COL_MAXIC).Value)

        ' Consults - find in list
        Dim sVal As String
        sVal = CStr(ws.Cells(lastRow, COL_CONSULT).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstEval.ListCount - 1
                If .lstEval.List(k) = sVal Then
                    .lstEval.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Fee Modifier 1
        sVal = CStr(ws.Cells(lastRow, COL_MOD1).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstMod1.ListCount - 1
                If .lstMod1.List(k) = sVal Then
                    .lstMod1.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Fee Modifier 2
        sVal = CStr(ws.Cells(lastRow, COL_MOD2).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstMod2.ListCount - 1
                If .lstMod2.List(k) = sVal Then
                    .lstMod2.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Fee Modifier 3
        sVal = CStr(ws.Cells(lastRow, COL_MOD3).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstMod3.ListCount - 1
                If .lstMod3.List(k) = sVal Then
                    .lstMod3.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Resuscitation
        sVal = CStr(ws.Cells(lastRow, COL_RESUS).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstResus.ListCount - 1
                If .lstResus.List(k) = sVal Then
                    .lstResus.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Obstetrics
        sVal = CStr(ws.Cells(lastRow, COL_OBS).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstObs.ListCount - 1
                If .lstObs.List(k) = sVal Then
                    .lstObs.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Acute Pain
        sVal = CStr(ws.Cells(lastRow, COL_ACUTEPAIN).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstAcPain.ListCount - 1
                If .lstAcPain.List(k) = sVal Then
                    .lstAcPain.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Chronic Pain
        sVal = CStr(ws.Cells(lastRow, COL_CHRONPAIN).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstChPain.ListCount - 1
                If .lstChPain.List(k) = sVal Then
                    .lstChPain.ListIndex = k
                    Exit For
                End If
            Next k
        End If

        ' Miscellaneous
        sVal = CStr(ws.Cells(lastRow, COL_MISC).Value)
        If Len(sVal) > 0 Then
            For k = 0 To .lstMisc.ListCount - 1
                If .lstMisc.List(k) = sVal Then
                    .lstMisc.ListIndex = k
                    Exit For
                End If
            Next k
        End If

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

    ' Re-enable formatting
    m_bFormatting = False

    ' Store the row being edited - do NOT delete until Save is clicked
    m_lEditRow = lastRow

    MsgBox "Record loaded for editing. Make your changes and click Save." & vbCrLf & _
           "The original record will be replaced when you save.", _
           vbInformation, "Edit Mode"

    Exit Sub
ErrHandler:
    m_bFormatting = False
    m_lEditRow = 0
    MsgBox "Edit error: " & Err.Description, vbCritical, "Error"
End Sub

'------------------------------------------------------------------------------
' IsRecordFromToday - Checks if a record's "Submitted On" date matches today
'------------------------------------------------------------------------------
Private Function IsRecordFromToday(ByVal ws As Worksheet, ByVal lRow As Long) As Boolean
    On Error GoTo NotToday

    Dim sSubmittedOn As String
    sSubmittedOn = CStr(ws.Cells(lRow, COL_SUBMON).Value)

    If Len(sSubmittedOn) = 0 Then
        IsRecordFromToday = False
        Exit Function
    End If

    ' The "Submitted On" field is in "DD/MM/YYYY HH:nn:SS" format
    ' Extract the date part (first 10 characters)
    Dim sDatePart As String
    If Len(sSubmittedOn) >= 10 Then
        sDatePart = Left(sSubmittedOn, 10)
    Else
        sDatePart = sSubmittedOn
    End If

    Dim dtSubmitted As Date
    If TryParseDateDMY(sDatePart, dtSubmitted) Then
        IsRecordFromToday = (dtSubmitted = Date)
    Else
        IsRecordFromToday = False
    End If
    Exit Function

NotToday:
    IsRecordFromToday = False
End Function

'==============================================================================
' DATE FIELD AUTO-FORMATTING (DD/MM/YYYY)
' Auto-inserts "/" separators as the user types digits
'==============================================================================

'--- Date of Service ---
Private Sub txtDteOfSer_Enter()
    If txtDteOfSer.Value = "DD/MM/YYYY" Then
        txtDteOfSer.Value = ""
    End If
End Sub

Private Sub txtDteOfSer_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtDteOfSer.Value)) = 0 Then
        txtDteOfSer.Value = "DD/MM/YYYY"
    End If
End Sub

Private Sub txtDteOfSer_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits - slashes are auto-inserted
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDteOfSer_Change()
    If m_bFormatting Then Exit Sub
    If txtDteOfSer.Value = "DD/MM/YYYY" Or txtDteOfSer.Value = "" Then Exit Sub
    FormatDateField txtDteOfSer
End Sub

'--- WCB Date of Injury ---
Private Sub txtWCBDteofInj_Enter()
    If txtWCBDteofInj.Value = "DD/MM/YYYY" Then
        txtWCBDteofInj.Value = ""
    End If
End Sub

Private Sub txtWCBDteofInj_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtWCBDteofInj.Value)) = 0 Then
        txtWCBDteofInj.Value = "DD/MM/YYYY"
    End If
End Sub

Private Sub txtWCBDteofInj_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits - slashes are auto-inserted
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtWCBDteofInj_Change()
    If m_bFormatting Then Exit Sub
    If txtWCBDteofInj.Value = "DD/MM/YYYY" Or txtWCBDteofInj.Value = "" Then Exit Sub
    FormatDateField txtWCBDteofInj
End Sub

'==============================================================================
' TIME FIELD AUTO-FORMATTING (HHMMhr)
' User types 4 digits, "hr" suffix is auto-appended
'==============================================================================

'--- Procedure Start Time ---
Private Sub txtProcStrtTime_Enter()
    If txtProcStrtTime.Value = "HHMMhr" Then
        txtProcStrtTime.Value = ""
    End If
End Sub

Private Sub txtProcStrtTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtProcStrtTime.Value)) = 0 Then
        txtProcStrtTime.Value = "HHMMhr"
    End If
End Sub

Private Sub txtProcStrtTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits - "hr" suffix is auto-appended
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtProcStrtTime_Change()
    If m_bFormatting Then Exit Sub
    If txtProcStrtTime.Value = "HHMMhr" Or txtProcStrtTime.Value = "" Then Exit Sub
    FormatTimeField txtProcStrtTime
End Sub

'--- Procedure Finish Time ---
Private Sub txtProcFinTime_Enter()
    If txtProcFinTime.Value = "HHMMhr" Then
        txtProcFinTime.Value = ""
    End If
End Sub

Private Sub txtProcFinTime_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Len(Trim(txtProcFinTime.Value)) = 0 Then
        txtProcFinTime.Value = "HHMMhr"
    End If
End Sub

Private Sub txtProcFinTime_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Only allow digits - "hr" suffix is auto-appended
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtProcFinTime_Change()
    If m_bFormatting Then Exit Sub
    If txtProcFinTime.Value = "HHMMhr" Or txtProcFinTime.Value = "" Then Exit Sub
    FormatTimeField txtProcFinTime
End Sub

'==============================================================================
' FORMAT HELPER FUNCTIONS
'==============================================================================

'------------------------------------------------------------------------------
' ExtractDigits - Returns only digit characters from a string
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
' FormatDateField - Auto-inserts "/" separators for DD/MM/YYYY format
'------------------------------------------------------------------------------
Private Sub FormatDateField(ByRef ctl As MSForms.TextBox)
    m_bFormatting = True

    Dim sDigits As String
    sDigits = ExtractDigits(ctl.Value)

    ' Limit to 8 digits (DDMMYYYY)
    If Len(sDigits) > 8 Then sDigits = Left(sDigits, 8)

    ' Build formatted string with "/" separators
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

    ' Limit to 4 digits (HHMM)
    If Len(sDigits) > 4 Then sDigits = Left(sDigits, 4)

    ' Build formatted string - append "hr" when 4 digits entered
    Dim sFormatted As String
    If Len(sDigits) = 4 Then
        sFormatted = sDigits & "hr"
    Else
        sFormatted = sDigits
    End If

    If sFormatted <> ctl.Value Then
        ctl.Value = sFormatted
        ' Position cursor before "hr" suffix
        If Len(sDigits) = 4 Then
            ctl.SelStart = 4
        Else
            ctl.SelStart = Len(sFormatted)
        End If
    End If

    m_bFormatting = False
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

    ' Check date is valid DD/MM/YYYY
    Dim sDate As String
    sDate = txtDteOfSer.Value
    If sDate = "DD/MM/YYYY" Or Len(sDate) = 0 Then
        txtDteOfSer.BackColor = &HC0C0FF ' Light red
        bValid = False
    ElseIf Not IsValidDateDMY(sDate) Then
        txtDteOfSer.BackColor = &HC0C0FF
        MsgBox "Invalid date format. Please use DD/MM/YYYY.", vbExclamation, "Validation"
        bValid = False
    End If

    ' Check procedure code
    If Len(txtSurgProcCode.Value) = 0 Then
        txtSurgProcCode.BackColor = &HC0C0FF
        bValid = False
    End If

    ' Check start time is valid HHMMhr (24-hour)
    Dim sStart As String
    sStart = txtProcStrtTime.Value
    If sStart = "HHMMhr" Or Len(sStart) = 0 Then
        txtProcStrtTime.BackColor = &HC0C0FF
        bValid = False
    ElseIf Not IsValidTime24(sStart) Then
        txtProcStrtTime.BackColor = &HC0C0FF
        MsgBox "Invalid start time. Please enter 4 digits in 24-hour format (e.g., 0800hr).", _
               vbExclamation, "Validation"
        bValid = False
    End If

    ' Check finish time is valid HHMMhr (24-hour)
    Dim sFinish As String
    sFinish = txtProcFinTime.Value
    If sFinish = "HHMMhr" Or Len(sFinish) = 0 Then
        txtProcFinTime.BackColor = &HC0C0FF
        bValid = False
    ElseIf Not IsValidTime24(sFinish) Then
        txtProcFinTime.BackColor = &HC0C0FF
        MsgBox "Invalid finish time. Please enter 4 digits in 24-hour format (e.g., 1630hr).", _
               vbExclamation, "Validation"
        bValid = False
    End If

    ' Check WCB date if entered
    Dim sWCBDate As String
    sWCBDate = txtWCBDteofInj.Value
    If sWCBDate <> "DD/MM/YYYY" And Len(sWCBDate) > 0 Then
        If Not IsValidDateDMY(sWCBDate) Then
            txtWCBDteofInj.BackColor = &HC0C0FF
            MsgBox "Invalid WCB date of injury. Please use DD/MM/YYYY.", vbExclamation, "Validation"
            bValid = False
        End If
    End If

    If Not bValid Then
        MsgBox "Please fill in all required fields (highlighted in red).", _
               vbExclamation, "Validation Error"
    End If

    ValidateForm = bValid
End Function
