Attribute VB_Name = "mod_Design"
Option Explicit

' =============================================================================
' Module: mod_Design
' Description: Applies professional design to the Dashboard
' Color Palette:
'   - Primary (Blue): #2C3E50 / RGB(44, 62, 80)
'   - Accent (Green): #27AE60 / RGB(39, 174, 96)
'   - Neutral (Gray): #ECF0F1 / RGB(236, 240, 241)
'   - Alert (Red): #E74C3C / RGB(231, 76, 60)
' =============================================================================

Sub ApplyDesign()
    Dim wsDash As Worksheet

    On Error GoTo ErrorHandler

    Set wsDash = ThisWorkbook.Sheets("Dashboard")

    ' Colors
    Dim bleuFonce As Long, vert As Long, grisClair As Long, rouge As Long
    bleuFonce = RGB(44, 62, 80)
    vert = RGB(39, 174, 96)
    grisClair = RGB(236, 240, 241)
    rouge = RGB(231, 76, 60)

    ' Euro symbol
    Dim euroSymbol As String
    euroSymbol = Chr(128)

    With wsDash
        ' Hide gridlines
        ActiveWindow.DisplayGridlines = False

        ' Light gray background
        .Cells.Interior.Color = grisClair

        ' Dashboard Title (A1:C1)
        .Range("A1:C1").Merge
        With .Range("A1")
            .Font.Name = "Calibri"
            .Font.Size = 24
            .Font.Bold = True
            .Font.Color = bleuFonce
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
            .RowHeight = 40
        End With

        ' Section header "KPIs Principaux" (A3)
        With .Range("A3")
            .Font.Name = "Calibri"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = bleuFonce
        End With

        ' KPI Labels (A4:A9)
        With .Range("A4:A9")
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Color = bleuFonce
        End With

        ' KPI Values (B4:B9)
        With .Range("B4:B9")
            .Font.Name = "Calibri"
            .Font.Size = 16
            .Font.Bold = True
            .Font.Color = bleuFonce
            .HorizontalAlignment = xlRight
        End With

        ' Number formats
        .Range("B4").NumberFormat = "# ##0 [$" & euroSymbol & "-40C]"
        .Range("B5").NumberFormat = "# ##0 [$" & euroSymbol & "-40C]"
        .Range("B6").NumberFormat = "0.0 ""h"""
        .Range("B7").NumberFormat = "0.00 [$" & euroSymbol & "-40C]""/h"""
        .Range("B8").NumberFormat = "0"
        .Range("B9").NumberFormat = "0.0 ""h"""

        ' Top Client (A11:B11)
        With .Range("A11")
            .Font.Name = "Calibri"
            .Font.Size = 11
            .Font.Color = bleuFonce
        End With
        With .Range("B11")
            .Font.Name = "Calibri"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = vert
        End With

        ' Statistics Section (A13:B16)
        With .Range("A13")
            .Font.Name = "Calibri"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = bleuFonce
        End With
        .Range("A14:A16").Font.Color = bleuFonce
        .Range("B14:B16").Font.Bold = True
        .Range("B14:B15").NumberFormat = "DD/MM/YYYY"

        ' Column widths
        .Columns("A").ColumnWidth = 25
        .Columns("B").ColumnWidth = 18
        .Columns("C").ColumnWidth = 3
        .Columns("D:G").ColumnWidth = 12
        .Columns("H:J").ColumnWidth = 12
        .Columns("K:L").ColumnWidth = 15

        ' Light borders around KPIs
        With .Range("A4:B9").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(189, 195, 199)
        End With

        With .Range("A11:B11").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(189, 195, 199)
        End With

        With .Range("A14:B16").Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(189, 195, 199)
        End With

        ' Row heights
        .Rows("1").RowHeight = 40
        .Rows("2").RowHeight = 10
        .Rows("3:16").RowHeight = 22
    End With

    MsgBox "Design applied successfully!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Error applying design: " & Err.Description, vbCritical, "Error"
End Sub
