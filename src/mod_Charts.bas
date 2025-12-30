Attribute VB_Name = "mod_Charts"
Option Explicit

' =============================================================================
' Module: mod_Charts
' Description: Creates Charts for the Freelance Dashboard
' =============================================================================

Sub CreateCharts()
    Dim wsDash As Worksheet
    Dim wsTCD As Worksheet
    Dim cht As ChartObject

    On Error GoTo ErrorHandler

    Set wsDash = ThisWorkbook.Sheets("Dashboard")
    Set wsTCD = ThisWorkbook.Sheets("TCD_Data")

    ' Delete existing charts
    On Error Resume Next
    For Each cht In wsDash.ChartObjects
        cht.Delete
    Next cht
    On Error GoTo ErrorHandler

    ' Chart 1: Pie Chart - Revenue by Client
    Set cht = wsDash.ChartObjects.Add( _
        Left:=wsDash.Range("D3").Left, _
        Top:=wsDash.Range("D3").Top, _
        Width:=250, _
        Height:=200)

    With cht.Chart
        .SetSourceData Source:=wsTCD.PivotTables("TCD_CA_Client").TableRange1
        .ChartType = xlPie
        .HasTitle = True
        .ChartTitle.Text = "CA par Client"
        .HasLegend = True
        .Legend.Position = xlLegendPositionRight
    End With

    ' Chart 2: Column Chart - Revenue by Month
    Set cht = wsDash.ChartObjects.Add( _
        Left:=wsDash.Range("D12").Left, _
        Top:=wsDash.Range("D12").Top, _
        Width:=250, _
        Height:=200)

    With cht.Chart
        .SetSourceData Source:=wsTCD.PivotTables("TCD_CA_Mois").TableRange1
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "CA par Mois"
        .HasLegend = False
    End With

    ' Chart 3: Bar Chart - Hours by Project
    Set cht = wsDash.ChartObjects.Add( _
        Left:=wsDash.Range("H3").Left, _
        Top:=wsDash.Range("H3").Top, _
        Width:=250, _
        Height:=200)

    With cht.Chart
        .SetSourceData Source:=wsTCD.PivotTables("TCD_Heures_Projet").TableRange1
        .ChartType = xlBarClustered
        .HasTitle = True
        .ChartTitle.Text = "Heures par Projet"
        .HasLegend = False
    End With

    Exit Sub

ErrorHandler:
    MsgBox "Error creating charts: " & Err.Description, vbCritical, "Error"
End Sub
