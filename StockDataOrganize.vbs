Attribute VB_Name = "Module1"
Sub StockDataOrganize()
'This macro is designed to work on a datatable of stock price organized as
'consecutive daily data starting with the first January data point and ending with final December price
'in columns of <ticker> <date> <open> <high> <low> <close> <vol> in columns A through G.
'1 create a summary table for each ticker in columns I through Q
'2 define and initialize variables
'3 cycle through column A. When ticker symbol changes, capture change in yearly price, percentage change and total stock volume
'4 cycle through summary table. calculate the greatest increase, decrease and stock volume in greatest hits table
'5 format summary and greatest hits tables, including conditional formatting
'6 repeat for all worksheets in workbook

For Each ws In Worksheets 'set up for cycling through worksheets

'1 create a summary table for each ticker in columns I through Q
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    ws.Range("p1").Value = "Ticker"
    ws.Range("q1").Value = "Value"
    ws.Range("o2").Value = "Greatest % Increase"
    ws.Range("o3").Value = "Greatest % Decrease"
    ws.Range("o4").Value = "Greatest Total Volume"

'2 define variables
    Dim i As Long
    Dim j As Long
    Dim y As Long
    Dim z As Long
    Dim OpenPrice As Double
    Dim ClosingPrice As Double
    Dim RowCount As Long
    Dim TotalVolume As Double
    Dim TableCount As Long

'2 initialize variables
    y = ws.Cells(Rows.Count, 1).End(xlUp).Row ' finds the last row with data
    OpenPrice = ws.Cells(2, 3).Value 'Sets first open price
    RowCount = 2
    TableCount = 2

'3 cycle through column A. When ticker symbol changes, capture change in yearly price, percentage change and total stock volume
    For i = 2 To y
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            'Record needed information
                ClosingPrice = ws.Cells(i, 6).Value
                TotalVolume = Application.Sum(ws.Range(ws.Cells(RowCount, 7), ws.Cells(i, 7)))
            'Input data in summary table
                ws.Cells(TableCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(TableCount, 10).Value = ClosingPrice - OpenPrice
                If OpenPrice <> 0 Then ws.Cells(TableCount, 11).Value = (ClosingPrice - OpenPrice) / OpenPrice 'if openprice = 0 then % calc blows up
                ws.Cells(TableCount, 12).Value = TotalVolume
            'Reset or increment Values
                OpenPrice = ws.Cells(i + 1, 3).Value
                ClosingPrice = 0
                RowCount = ws.Cells(i + 1, 6).Row
                TableCount = TableCount + 1
                TotalVolume = 0
        End If
 
    Next i

'Check Sums. Total volume in column g should be the same as total volume in summary table
If Application.Sum(ws.Columns("g")) <> Application.Sum(ws.Columns("L")) Then MsgBox ("Yeah your total volumes don't tie out mate")

'4 cycle through summary table. calculate the greatest increase, decrease and stock volume in greatest hits table
z = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row ' finds the last row in summary table with data
    For j = 2 To z 'cycle through summary table and populate greatest hits table
        If ws.Cells(j, 11).Value = Application.Max(ws.Columns("K")) Then
            ws.Range("P2").Value = ws.Cells(j, 9).Value
            ws.Range("q2").Value = ws.Cells(j, 11).Value
        ElseIf ws.Cells(j, 11).Value = Application.Min(ws.Columns("K")) Then
            ws.Range("P3").Value = ws.Cells(j, 9).Value
            ws.Range("q3").Value = ws.Cells(j, 11).Value
        ElseIf ws.Cells(j, 12).Value = Application.Max(ws.Columns("l")) Then
            ws.Range("P4").Value = ws.Cells(j, 9).Value
            ws.Range("q4").Value = ws.Cells(j, 12).Value
        End If
    Next j

'5 format summary and greatest hits tables, including conditional formatting
    ws.Columns("i:q").EntireColumn.AutoFit
    ws.Columns("j:j").NumberFormat = "#,##0.00_);(#,##0.00)"
    ws.Columns("l:l").NumberFormat = "#,##0_);[Red](#,##0)"
    ws.Range("q4").NumberFormat = "#,##0_);[Red](#,##0)"
    ws.Columns("k").NumberFormat = "0.00%"
    ws.Range("q2:q3").NumberFormat = "0.00%"

'5 conditional formatting, using z as last row in summary table
    With ws.Range("j" & 2 & ":j" & z).FormatConditions.Add(xlCellValue, xlGreater, "=0")
        .Interior.Color = vbGreen
    End With

    With ws.Range("j" & 2 & ":j" & z).FormatConditions.Add(xlCellValue, xlLess, "=0")
        .Interior.Color = vbRed
    End With

'6 repeat for all worksheets in workbook


Next ws

End Sub

