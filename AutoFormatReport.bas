Attribute VB_Name = "AutoFormatReport"
'==========================================================================
' AutoFormatReport
'
' Purpose: Given a raw data dump on a sheet named "RawData" (Date | Ticker |
' Close | Volume), this macro builds a formatted summary table with
' conditional formatting, on a new "Summary" sheet, entirely inside Excel.
'
' This is the VBA-only counterpart to weekly_report_generator.py: useful
' when the data has already been pasted into Excel (e.g. from a Bloomberg
' or internal terminal export) and you want one-click formatting instead of
' re-running a Python script.
'
' How to use:
'   1. Open Excel, press Alt+F11 to open the VBA editor.
'   2. File > Import File... and select this .bas file.
'      (Or create a new Module and paste the code below.)
'   3. Put your raw data on a sheet named "RawData" with headers in row 1:
'      Date | Ticker | Close | Volume
'   4. Press F5, or run the "BuildSummary" macro from Developer > Macros.
'==========================================================================

Sub BuildSummary()

    Dim wsRaw As Worksheet
    Dim wsSummary As Worksheet
    Dim lastRow As Long
    Dim tickers As Collection
    Dim t As Variant
    Dim i As Long, outRow As Long
    Dim firstClose As Double, lastClose As Double, pctChange As Double
    Dim avgVolume As Double, volSum As Double, volCount As Long

    Set wsRaw = ThisWorkbook.Sheets("RawData")

    ' Remove old Summary sheet if it exists, then recreate it
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsSummary = ThisWorkbook.Sheets.Add(Before:=wsRaw)
    wsSummary.Name = "Summary"

    ' --- Title ---
    With wsSummary.Range("A1")
        .Value = "Weekly Market Report - " & Format(Date, "mmmm dd, yyyy")
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(31, 41, 55)
    End With

    ' --- Headers ---
    Dim headers As Variant
    headers = Array("Ticker", "Last Close", "Weekly % Change", "Avg Volume")
    For i = 0 To UBound(headers)
        With wsSummary.Cells(3, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 41, 55)
            .HorizontalAlignment = xlCenter
        End With
    Next i

    ' --- Collect distinct tickers from RawData ---
    Set tickers = New Collection
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, "B").End(xlUp).Row

    On Error Resume Next
    For i = 2 To lastRow
        tickers.Add wsRaw.Cells(i, 2).Value, CStr(wsRaw.Cells(i, 2).Value)
    Next i
    On Error GoTo 0

    ' --- Compute per-ticker stats and write to Summary ---
    outRow = 4
    For Each t In tickers
        firstClose = 0: lastClose = 0: volSum = 0: volCount = 0

        For i = 2 To lastRow
            If wsRaw.Cells(i, 2).Value = t Then
                If firstClose = 0 Then firstClose = wsRaw.Cells(i, 3).Value
                lastClose = wsRaw.Cells(i, 3).Value
                volSum = volSum + wsRaw.Cells(i, 4).Value
                volCount = volCount + 1
            End If
        Next i

        If firstClose > 0 And volCount > 0 Then
            pctChange = (lastClose - firstClose) / firstClose
            avgVolume = volSum / volCount

            wsSummary.Cells(outRow, 1).Value = t
            wsSummary.Cells(outRow, 2).Value = lastClose
            wsSummary.Cells(outRow, 2).NumberFormat = "$#,##0.00"
            wsSummary.Cells(outRow, 3).Value = pctChange
            wsSummary.Cells(outRow, 3).NumberFormat = "0.00%"
            wsSummary.Cells(outRow, 4).Value = avgVolume
            wsSummary.Cells(outRow, 4).NumberFormat = "#,##0"

            outRow = outRow + 1
        End If
    Next t

    ' --- Conditional formatting on the % change column ---
    Dim pctRange As Range
    Set pctRange = wsSummary.Range("C4:C" & (outRow - 1))
    pctRange.FormatConditions.Delete

    With pctRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = RGB(198, 239, 206)
        .Font.Color = RGB(0, 97, 0)
    End With

    With pctRange.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = RGB(255, 199, 206)
        .Font.Color = RGB(156, 0, 6)
    End With

    ' --- Column widths ---
    wsSummary.Columns("A:D").AutoFit

    wsSummary.Activate
    MsgBox "Summary built for " & (outRow - 4) & " tickers.", vbInformation

End Sub
