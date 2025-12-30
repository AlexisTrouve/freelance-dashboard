Attribute VB_Name = "mod_Refresh"
Option Explicit

' =============================================================================
' Module: mod_Refresh
' Description: Main refresh and rebuild macros for the Dashboard
' =============================================================================

Sub RefreshDashboard()
    ' Main refresh macro - recalculates formulas and refreshes pivot tables
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrorHandler

    ' 1. Recalculate all formulas
    Application.CalculateFull

    ' 2. Refresh all Pivot Tables
    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    ' 3. Refresh external connections (if any)
    On Error Resume Next
    Dim conn As WorkbookConnection
    For Each conn In ThisWorkbook.Connections
        conn.Refresh
    Next conn
    On Error GoTo ErrorHandler

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Dashboard refreshed!" & vbCrLf & _
           "- Formulas recalculated" & vbCrLf & _
           "- Pivot tables updated", vbInformation, "Refresh Complete"
    Exit Sub

ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error during refresh: " & Err.Description, vbCritical, "Error"
End Sub

Sub QuickRefresh()
    ' Silent refresh without message box
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.CalculateFull

    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
End Sub

Sub RebuildAll()
    ' Rebuilds entire dashboard from scratch
    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False

    ' Call other macros
    CreatePivotTables
    CreateCharts
    CreateSlicers
    ApplyDesign

    Application.ScreenUpdating = True
    MsgBox "Dashboard rebuilt entirely!", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error during rebuild: " & Err.Description, vbCritical, "Error"
End Sub
