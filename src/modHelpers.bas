Attribute VB_Name = "modHelpers"
'==============================================================================
' modHelpers - Utility Helpers Module
' APP Billing System
'
' Contains helper routines used by forms and other modules.
'==============================================================================
Option Explicit

'------------------------------------------------------------------------------
' UnloadSuperUser - Deferred unload for frmSuperUser when auth fails
' Called via Application.OnTime from UserForm_Initialize
'------------------------------------------------------------------------------
Public Sub UnloadSuperUser()
    On Error Resume Next
    Unload frmSuperUser
End Sub

'------------------------------------------------------------------------------
' Show_DailyExport - Opens the Daily Export form
' Can be assigned to a button on the Home sheet
'------------------------------------------------------------------------------
Public Sub Show_DailyExport()
    frmDailyExport.Show
End Sub
