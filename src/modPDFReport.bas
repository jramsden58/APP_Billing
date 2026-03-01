Attribute VB_Name = "modPDFReport"
'==============================================================================
' modPDFReport - PDF Report Generation Module
' APP Billing System
'
' Populates the ORReportingForm sheet with daily data and exports to PDF.
' Uses the existing form layout which supports 6 procedure entries per page.
'==============================================================================
Option Explicit

' ORReportingForm layout constants (based on existing sheet structure)
' Header fields
Private Const FORM_NAME_CELL As String = "C3"       ' Anesthesiologist name
Private Const FORM_MSP_CELL As String = "L3"        ' MSP Billing #
Private Const FORM_SITE_CELL As String = "C5"        ' Site
Private Const FORM_SHIFT_CELL As String = "C6"       ' Shift Name
Private Const FORM_SHIFTTYPE_CELL As String = "H6"   ' Shift Type
Private Const FORM_ONCALL_CELL As String = "L6"      ' On Call
Private Const FORM_DATE_CELL As String = "C8"        ' Date of Service

' Procedure block starting rows (6 blocks, each ~7 rows)
Private Const PROC_START_ROWS As String = "10,17,24,31,38,45"
' Within each block, relative offsets:
'   Row +0: Consult, Procedure Code, IC Level, fee codes
'   Row +1: Procedure Start/Finish times, WCB fields
'   (Exact layout follows existing ORReportingForm structure)

'------------------------------------------------------------------------------
' GenerateDailyPDF - Main entry point for PDF report generation
'
' Parameters:
'   sUserName - Anesthesiologist name (as stored in data)
'   dtDate    - Date of service to report on
'   bPreview  - If True, shows the form instead of saving PDF
'
' Returns: Path to saved PDF, or empty string on failure/preview
'------------------------------------------------------------------------------
Public Function GenerateDailyPDF(ByVal sUserName As String, _
                                  ByVal dtDate As Date, _
                                  Optional ByVal bPreview As Boolean = False) As String
    On Error GoTo ErrHandler

    ' Get data for this user and date
    Dim vData As Variant
    vData = GetUserData(sUserName, dtDate)

    If IsEmpty(vData) Then
        MsgBox "No data found for " & sUserName & " on " & Format(dtDate, "DD/MM/YYYY") & ".", _
               vbInformation, "No Data"
        GenerateDailyPDF = ""
        Exit Function
    End If

    ' Get MSP number from LookupLists
    Dim sMSP As String
    sMSP = GetMSPNumber(sUserName)

    ' Populate the form
    PopulateORForm vData, sUserName, sMSP, dtDate

    If bPreview Then
        ' Activate the sheet for preview
        ThisWorkbook.Sheets("ORReportingForm").Activate
        GenerateDailyPDF = ""
        Exit Function
    End If

    ' Determine the number of pages needed
    Dim lRecords As Long
    If IsArray(vData) Then
        lRecords = UBound(vData, 1)
    Else
        lRecords = 1
    End If

    Dim lPages As Long
    lPages = Int((lRecords - 1) / 6) + 1

    ' Export to PDF
    Dim sPDFPath As String
    sPDFPath = GetPDFPath(sUserName, dtDate)

    If Len(sPDFPath) = 0 Then
        ' Network not available, save locally
        sPDFPath = ThisWorkbook.Path & "\" & _
                   Replace(Replace(sUserName, " ", "_"), ",", "") & _
                   "_" & Format(dtDate, "YYYYMMDD") & ".pdf"
    End If

    ExportToPDF sPDFPath

    ' If more than 6 records, generate additional pages
    If lRecords > 6 Then
        Dim lPage As Long
        For lPage = 2 To lPages
            Dim lStartIdx As Long
            lStartIdx = ((lPage - 1) * 6) + 1

            ' Populate next batch of 6 records
            PopulateORFormPage vData, sUserName, sMSP, dtDate, lStartIdx

            ' Append to PDF (save as separate file then note for user)
            Dim sPagePath As String
            sPagePath = Replace(sPDFPath, ".pdf", "_Page" & lPage & ".pdf")
            ExportToPDF sPagePath
        Next lPage

        MsgBox "Report generated with " & lPages & " pages." & vbCrLf & _
               "Note: Due to Excel PDF limitations, each page is a separate file." & vbCrLf & vbCrLf & _
               "Files saved to: " & Left(sPDFPath, InStrRev(sPDFPath, "\")), _
               vbInformation, "PDF Report"
    Else
        MsgBox "PDF report saved to:" & vbCrLf & sPDFPath, vbInformation, "PDF Report"
    End If

    GenerateDailyPDF = sPDFPath
    Exit Function

ErrHandler:
    MsgBox "Error generating PDF: " & Err.Description, vbCritical, "PDF Error"
    GenerateDailyPDF = ""
End Function

'------------------------------------------------------------------------------
' GetUserData - Gets data for a specific user and date
' Tries network share first, falls back to local DailyDatabase
'------------------------------------------------------------------------------
Private Function GetUserData(ByVal sUserName As String, ByVal dtDate As Date) As Variant
    ' Try network share first
    If IsNetworkAvailable() Then
        Dim vNetData As Variant
        vNetData = ReadUserDailyData(sUserName, dtDate)
        If Not IsEmpty(vNetData) Then
            GetUserData = vNetData
            Exit Function
        End If
    End If

    ' Fall back to local DailyDatabase
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DailyDatabase")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_ANESTH).End(xlUp).Row

    If lastRow < 2 Then
        GetUserData = Empty
        Exit Function
    End If

    ' Count matching records first
    Dim lCount As Long
    lCount = 0
    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, COL_ANESTH).Value) = sUserName Then
            Dim dtRow As Date
            If TryParseDate(CStr(ws.Cells(i, COL_DATE).Value), dtRow) Then
                If dtRow = dtDate Then
                    lCount = lCount + 1
                End If
            End If
        End If
    Next i

    If lCount = 0 Then
        GetUserData = Empty
        Exit Function
    End If

    ' Build array of matching records
    Dim vResult() As Variant
    ReDim vResult(1 To lCount, 1 To NUM_COLUMNS)

    Dim lIdx As Long
    lIdx = 0
    For i = 2 To lastRow
        If CStr(ws.Cells(i, COL_ANESTH).Value) = sUserName Then
            If TryParseDate(CStr(ws.Cells(i, COL_DATE).Value), dtRow) Then
                If dtRow = dtDate Then
                    lIdx = lIdx + 1
                    Dim j As Long
                    For j = 1 To NUM_COLUMNS
                        vResult(lIdx, j) = ws.Cells(i, j).Value
                    Next j
                End If
            End If
        End If
    Next i

    GetUserData = vResult
End Function

'------------------------------------------------------------------------------
' PopulateORForm - Fills in the ORReportingForm sheet with data
'------------------------------------------------------------------------------
Private Sub PopulateORForm(ByVal vData As Variant, ByVal sUserName As String, _
                           ByVal sMSP As String, ByVal dtDate As Date)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ORReportingForm")

    ' Clear previous data
    ClearORForm ws

    ' Header fields
    ws.Range(FORM_NAME_CELL).Value = sUserName
    ws.Range(FORM_MSP_CELL).Value = sMSP
    ws.Range(FORM_DATE_CELL).Value = Format(dtDate, "DD/MM/YYYY")

    ' Get common fields from first record
    If IsArray(vData) And UBound(vData, 1) >= 1 Then
        ws.Range(FORM_SITE_CELL).Value = vData(1, COL_SITE)
        ws.Range(FORM_SHIFT_CELL).Value = vData(1, COL_SHIFT)
        ws.Range(FORM_SHIFTTYPE_CELL).Value = vData(1, COL_SHIFTTYPE)
        ws.Range(FORM_ONCALL_CELL).Value = IIf(vData(1, COL_ONCALL) = True Or _
            LCase(CStr(vData(1, COL_ONCALL) & "")) = "yes", "Yes", "No")
    End If

    ' Populate procedure blocks (up to 6)
    Dim procRows() As String
    procRows = Split(PROC_START_ROWS, ",")

    Dim lMax As Long
    If IsArray(vData) Then
        lMax = Application.Min(UBound(vData, 1), 6)
    Else
        lMax = 0
    End If

    Dim idx As Long
    For idx = 1 To lMax
        Dim lProcRow As Long
        lProcRow = CLng(procRows(idx - 1))

        ' Procedure Code
        ws.Cells(lProcRow, 3).Value = vData(idx, COL_PROCCODE)        ' Col C

        ' Consult
        ws.Cells(lProcRow, 1).Value = vData(idx, COL_CONSULT)         ' Col A

        ' IC Level
        ws.Cells(lProcRow, 5).Value = vData(idx, COL_MAXIC)           ' Col E

        ' Modifiers
        ws.Cells(lProcRow, 7).Value = vData(idx, COL_MOD1)            ' Col G
        ws.Cells(lProcRow, 8).Value = vData(idx, COL_MOD2)            ' Col H
        ws.Cells(lProcRow, 9).Value = vData(idx, COL_MOD3)            ' Col I

        ' Resuscitation
        ws.Cells(lProcRow, 10).Value = vData(idx, COL_RESUS)          ' Col J

        ' Obstetrics
        ws.Cells(lProcRow, 11).Value = vData(idx, COL_OBS)            ' Col K

        ' Acute Pain
        ws.Cells(lProcRow + 1, 7).Value = vData(idx, COL_ACUTEPAIN)   ' Next row, Col G

        ' Chronic Pain
        ws.Cells(lProcRow + 1, 8).Value = vData(idx, COL_CHRONPAIN)   ' Next row, Col H

        ' Miscellaneous
        ws.Cells(lProcRow + 1, 9).Value = vData(idx, COL_MISC)        ' Next row, Col I

        ' Procedure times
        ws.Cells(lProcRow + 2, 3).Value = vData(idx, COL_STARTTIME)   ' Start time
        ws.Cells(lProcRow + 2, 5).Value = vData(idx, COL_FINTIME)     ' Finish time

        ' WCB fields
        ws.Cells(lProcRow + 2, 7).Value = vData(idx, COL_WCBNUM)      ' WCB #
        ws.Cells(lProcRow + 2, 9).Value = vData(idx, COL_WCBDATE)     ' Date of Injury
        ws.Cells(lProcRow + 2, 10).Value = vData(idx, COL_WCBSIDE)    ' Injury Side
        ws.Cells(lProcRow + 2, 11).Value = vData(idx, COL_WCBINJ)     ' Injury Type
    Next idx
End Sub

'------------------------------------------------------------------------------
' PopulateORFormPage - Populates form for a specific page (records 7-12, etc.)
'------------------------------------------------------------------------------
Private Sub PopulateORFormPage(ByVal vData As Variant, ByVal sUserName As String, _
                                ByVal sMSP As String, ByVal dtDate As Date, _
                                ByVal lStartIdx As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ORReportingForm")

    ' Clear and re-populate header
    ClearORForm ws
    ws.Range(FORM_NAME_CELL).Value = sUserName
    ws.Range(FORM_MSP_CELL).Value = sMSP
    ws.Range(FORM_DATE_CELL).Value = Format(dtDate, "DD/MM/YYYY") & " (cont.)"
    ws.Range(FORM_SITE_CELL).Value = vData(1, COL_SITE)
    ws.Range(FORM_SHIFT_CELL).Value = vData(1, COL_SHIFT)
    ws.Range(FORM_SHIFTTYPE_CELL).Value = vData(1, COL_SHIFTTYPE)
    ws.Range(FORM_ONCALL_CELL).Value = IIf(vData(1, COL_ONCALL) = True Or _
        LCase(CStr(vData(1, COL_ONCALL) & "")) = "yes", "Yes", "No")

    ' Populate procedure blocks starting from lStartIdx
    Dim procRows() As String
    procRows = Split(PROC_START_ROWS, ",")

    Dim lMax As Long
    lMax = Application.Min(UBound(vData, 1), lStartIdx + 5)

    Dim idx As Long
    Dim blockIdx As Long
    blockIdx = 0
    For idx = lStartIdx To lMax
        Dim lProcRow As Long
        lProcRow = CLng(procRows(blockIdx))

        ws.Cells(lProcRow, 3).Value = vData(idx, COL_PROCCODE)
        ws.Cells(lProcRow, 1).Value = vData(idx, COL_CONSULT)
        ws.Cells(lProcRow, 5).Value = vData(idx, COL_MAXIC)
        ws.Cells(lProcRow, 7).Value = vData(idx, COL_MOD1)
        ws.Cells(lProcRow, 8).Value = vData(idx, COL_MOD2)
        ws.Cells(lProcRow, 9).Value = vData(idx, COL_MOD3)
        ws.Cells(lProcRow, 10).Value = vData(idx, COL_RESUS)
        ws.Cells(lProcRow, 11).Value = vData(idx, COL_OBS)
        ws.Cells(lProcRow + 1, 7).Value = vData(idx, COL_ACUTEPAIN)
        ws.Cells(lProcRow + 1, 8).Value = vData(idx, COL_CHRONPAIN)
        ws.Cells(lProcRow + 1, 9).Value = vData(idx, COL_MISC)
        ws.Cells(lProcRow + 2, 3).Value = vData(idx, COL_STARTTIME)
        ws.Cells(lProcRow + 2, 5).Value = vData(idx, COL_FINTIME)
        ws.Cells(lProcRow + 2, 7).Value = vData(idx, COL_WCBNUM)
        ws.Cells(lProcRow + 2, 9).Value = vData(idx, COL_WCBDATE)
        ws.Cells(lProcRow + 2, 10).Value = vData(idx, COL_WCBSIDE)
        ws.Cells(lProcRow + 2, 11).Value = vData(idx, COL_WCBINJ)

        blockIdx = blockIdx + 1
    Next idx
End Sub

'------------------------------------------------------------------------------
' ClearORForm - Clears all data cells in the ORReportingForm
'------------------------------------------------------------------------------
Private Sub ClearORForm(ByVal ws As Worksheet)
    On Error Resume Next

    ' Clear header fields
    ws.Range(FORM_NAME_CELL).Value = ""
    ws.Range(FORM_MSP_CELL).Value = ""
    ws.Range(FORM_SITE_CELL).Value = ""
    ws.Range(FORM_SHIFT_CELL).Value = ""
    ws.Range(FORM_SHIFTTYPE_CELL).Value = ""
    ws.Range(FORM_ONCALL_CELL).Value = ""
    ws.Range(FORM_DATE_CELL).Value = ""

    ' Clear all procedure blocks
    Dim procRows() As String
    procRows = Split(PROC_START_ROWS, ",")

    Dim i As Long
    For i = 0 To UBound(procRows)
        Dim lRow As Long
        lRow = CLng(procRows(i))

        ' Clear 3 rows per block, columns A through L
        ws.Range(ws.Cells(lRow, 1), ws.Cells(lRow + 2, 12)).ClearContents
    Next i

    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' ExportToPDF - Exports the ORReportingForm to PDF
'------------------------------------------------------------------------------
Private Sub ExportToPDF(ByVal sTargetPath As String)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ORReportingForm")

    ' Ensure the directory exists
    Dim sDir As String
    sDir = Left(sTargetPath, InStrRev(sTargetPath, "\"))
    CreateFolderIfNotExists sDir

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=sTargetPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    Exit Sub

ErrHandler:
    Err.Raise Err.Number, "ExportToPDF", "Error exporting PDF: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' GetPDFPath - Returns the network path for a user's daily PDF
'------------------------------------------------------------------------------
Private Function GetPDFPath(ByVal sUserName As String, ByVal dtDate As Date) As String
    If Not IsNetworkAvailable() Then
        GetPDFPath = ""
        Exit Function
    End If

    Dim sBase As String
    sBase = GetNetworkPath() & FOLDER_PDF_REPORTS & "\"

    ' Ensure folder exists
    CreateFolderIfNotExists sBase

    Dim sClean As String
    sClean = Replace(Replace(sUserName, " ", "_"), ",", "")

    GetPDFPath = sBase & sClean & "_" & Format(dtDate, "YYYYMMDD") & ".pdf"
End Function

'------------------------------------------------------------------------------
' GetMSPNumber - Looks up MSP billing number from the AnesthList
'------------------------------------------------------------------------------
Private Function GetMSPNumber(ByVal sUserName As String) As String
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LookupLists")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If CStr(ws.Cells(i, 1).Value) = sUserName Then
            GetMSPNumber = CStr(ws.Cells(i, 2).Value)
            Exit Function
        End If
    Next i

    GetMSPNumber = ""
    Exit Function

ErrHandler:
    GetMSPNumber = ""
End Function

'------------------------------------------------------------------------------
' TryParseDate - Attempts to parse a date string (DD/MM/YYYY format, locale-safe)
'------------------------------------------------------------------------------
Private Function TryParseDate(ByVal sDate As String, ByRef dtResult As Date) As Boolean
    TryParseDate = TryParseDateDMY(sDate, dtResult)
End Function
