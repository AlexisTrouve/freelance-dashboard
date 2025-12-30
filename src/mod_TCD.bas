Attribute VB_Name = "mod_TCD"
Option Explicit

' =============================================================================
' Module: mod_TCD
' Description: Creates Pivot Tables for the Freelance Dashboard
' =============================================================================

Sub CreatePivotTables()
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim pc As PivotCache

    On Error GoTo ErrorHandler

    ' Create TCD sheet if it doesn't exist
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("TCD_Data")
    On Error GoTo ErrorHandler

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "TCD_Data"
    End If

    ' Clear the sheet
    ws.Cells.Clear

    ' TCD 1: Revenue by Client
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="tbl_Revenus")

    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("A1"), _
        TableName:="TCD_CA_Client")

    With pt
        .PivotFields("ClientID").Orientation = xlRowField
        .PivotFields("ClientID").Position = 1
        With .PivotFields("Montant")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "# ##0 $"
            .Name = "CA Total"
        End With
    End With

    ' TCD 2: Revenue by Month
    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("E1"), _
        TableName:="TCD_CA_Mois")

    With pt
        .PivotFields("Date").Orientation = xlRowField
        .PivotFields("Date").Position = 1
        With .PivotFields("Montant")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "# ##0 $"
            .Name = "CA Mensuel"
        End With
    End With

    ' Group dates by month and year
    On Error Resume Next
    pt.PivotFields("Date").DataRange.Cells(1).Group _
        Start:=True, End:=True, _
        Periods:=Array(False, False, False, False, True, False, True)
    On Error GoTo ErrorHandler

    ' TCD 3: Hours by Project
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:="tbl_Temps")

    Set pt = pc.CreatePivotTable( _
        TableDestination:=ws.Range("I1"), _
        TableName:="TCD_Heures_Projet")

    With pt
        .PivotFields("Projet").Orientation = xlRowField
        .PivotFields("Projet").Position = 1
        With .PivotFields("Heures")
            .Orientation = xlDataField
            .Function = xlSum
            .NumberFormat = "0.0"
            .Name = "Total Heures"
        End With
    End With

    ws.Activate

    MsgBox "3 Pivot Tables created successfully!", vbInformation, "TCD Created"
    Exit Sub

ErrorHandler:
    MsgBox "Error creating Pivot Tables: " & Err.Description, vbCritical, "Error"
End Sub
