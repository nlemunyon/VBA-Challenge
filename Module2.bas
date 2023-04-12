Attribute VB_Name = "Module1"
Sub MultiYear()


Dim ws As Worksheet
Dim i As Long
Dim lastrow As Long
Dim ticker_name As String
Dim stock_total As Double


Application.ScreenUpdating = False


For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


stock_total = 0
row_summary = 2

For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


ticker_name = ws.Cells(i, 1).Value
stock_total = stock_total + ws.Cells(i, 7).Value
ws.Cells(row_summary, 9).Value = ticker_name
year_open = ws.Cells(i - 250, 3).Value
year_close = ws.Cells(i, 6).Value
year_change = year_close - year_open
ws.Cells(row_summary, 10).Value = year_change

If ws.Cells(row_summary, 10).Value < 0 Then
ws.Cells(row_summary, 10).Interior.Color = vbRed
Else
ws.Cells(row_summary, 10).Interior.Color = vbGreen
End If

percentchange = (year_close - year_open) / year_open
ws.Cells(row_summary, 11).Value = percentchange
ws.Cells(row_summary, 11).NumberFormat = "0.00%"
ws.Cells(row_summary, 12).Value = stock_total

row_summary = row_summary + 1
stock_total = 0

Else

stock_total = stock_total + ws.Cells(i, 7).Value

End If

Next i

increase = ws.Application.WorksheetFunction.Max(ws.Range("K:K"))
ws.Cells(2, 17).Value = increase
ws.Cells(2, 17).NumberFormat = "0.00%"

decrease = ws.Application.WorksheetFunction.Min(ws.Range("K:K"))
ws.Cells(3, 17).Value = decrease
ws.Cells(3, 17).NumberFormat = "0.00%"

Volume = ws.Application.WorksheetFunction.Max(ws.Range("L:L"))
ws.Cells(4, 17).Value = Volume
ws.Cells(4, 17).NumberFormat = "0"


ticker_increase = ws.Application.WorksheetFunction.XLookup(ws.Range("Q2"), ws.Range("K:K"), ws.Range("I:I"))
ticker_decrease = ws.Application.WorksheetFunction.XLookup(ws.Range("Q3"), ws.Range("K:K"), ws.Range("I:I"))
ticker_volume = ws.Application.WorksheetFunction.XLookup(ws.Range("Q4"), ws.Range("L:L"), ws.Range("I:I"))

ws.Cells(2, 16).Value = ticker_increase
ws.Cells(3, 16).Value = ticker_decrease
ws.Cells(4, 16).Value = ticker_volume


ws.Range("I:L").EntireColumn.AutoFit
ws.Range("O:Q").EntireColumn.AutoFit

Next ws

Application.ScreenUpdating = True







End Sub

