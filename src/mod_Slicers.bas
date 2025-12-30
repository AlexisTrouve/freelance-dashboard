Attribute VB_Name = "mod_Slicers"
Option Explicit

' =============================================================================
' Module: mod_Slicers
' Description: Creates Slicers for interactive filtering
' =============================================================================

Sub CreateSlicers()
    Dim wsDash As Worksheet
    Dim wsTCD As Worksheet
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim pt As PivotTable

    On Error GoTo ErrorHandler

    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsTCD = ThisWorkbook.Sheets("TCD_Data")

    ' Delete existing slicers
    On Error Resume Next
    Dim existingSC As SlicerCache
    For Each existingSC In ThisWorkbook.SlicerCaches
        existingSC.Delete
    Next existingSC
    On Error GoTo ErrorHandler

    ' Slicer 1: ClientID on TCD_CA_Client
    Set pt = wsTCD.PivotTables("TCD_CA_Client")

    Set sc = ThisWorkbook.SlicerCaches.Add2( _
        pt, _
        "ClientID", _
        "Slicer_ClientID")

    Set sl = sc.Slicers.Add( _
        wsDash, _
        , _
        "ClientID", _
        "Client", _
        wsDash.Range("K1").Left, _
        wsDash.Range("K1").Top, _
        150, _
        180)

    sl.Style = "SlicerStyleLight1"

    ' Connect slicer to other pivot tables
    On Error Resume Next
    sc.PivotTables.AddPivotTable wsTCD.PivotTables("TCD_CA_Mois")
    On Error GoTo ErrorHandler

    ' Slicer 2: Date on TCD_CA_Mois
    Set pt = wsTCD.PivotTables("TCD_CA_Mois")

    On Error Resume Next
    Set sc = ThisWorkbook.SlicerCaches.Add2( _
        pt, _
        "Date", _
        "Slicer_Date")

    If Not sc Is Nothing Then
        Set sl = sc.Slicers.Add( _
            wsDash, _
            , _
            "Date", _
            "Periode", _
            wsDash.Range("K10").Left, _
            wsDash.Range("K10").Top, _
            150, _
            150)

        sl.Style = "SlicerStyleLight2"
    End If
    On Error GoTo ErrorHandler

    MsgBox "Slicers created successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error creating slicers: " & Err.Description, vbCritical, "Error"
End Sub

Sub ClearSlicerFilters()
    ' Clears all slicer filters
    Dim sc As SlicerCache

    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
    On Error GoTo 0
End Sub
