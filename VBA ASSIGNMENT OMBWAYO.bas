Attribute VB_Name = "Module1"
Sub Multiple_year_stock()

For Each ws In Worksheets

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

Dim SumTicker As Integer
Dim i As Long
Dim Ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim StockVolTotal As Double
Dim MaxPercent As Double
Dim MinPercent As Double
Dim MaxVolume As Long

openprice = ws.Cells(2, 3).Value
Ticker = ""
SumTicker = 1
YearlyChange = 0
PercentChange = 0
StockVolTotal = 0

For i = 2 To LastRow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
SumTicker = SumTicker + 1
Ticker = ws.Cells(i, 1).Value
ws.Cells(SumTicker, "I").Value = Ticker
closeprice = ws.Cells(i, 6).Value
YearlyChange = closeprice - openprice
ws.Cells(SumTicker, "J").Value = YearlyChange

StockVolTotal = StockVolTotal + ws.Cells(i, 7).Value
ws.Cells(SumTicker, "L").Value = StockVolTotal
PercentChange = (YearlyChange / openprice)
ws.Cells(SumTicker, "K").Value = PercentChange
ws.Cells(SumTicker, "K").NumberFormat = "0.00%"
openprice = ws.Cells(i + 1, 3).Value
SumTicker = SumTicker + 1

StockVolTotal = 0

Else
StockVolTotal = StockVolTotal + ws.Cells(i, 7).Value

End If

Next i

YearlyChangelastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To YearlyChangelastrow
If ws.Cells(i, 10) > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 10
Else
ws.Cells(i, 10).Interior.ColorIndex = 3


End If
MaxPercent = 0
MinPercent = 0
MaxVolume = 0

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"

MaxPercent = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & YearlyChangelastrow)), ws.Range("K2:K" & YearlyChangelastrow), 0)
ws.Range("O2") = ws.Cells(MaxPercent + 1, 9)
MinPercent = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & YearlyChangelastrow)), ws.Range("K2:K" & YearlyChangelastrow), 0)
MaxVolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & YearlyChangelastrow)), ws.Range("L2:L" & YearlyChangelastrow), 0)
ws.Range("O3") = ws.Cells(MinPercent + 1, 9)
ws.Range("O4") = ws.Cells(MaxVolume + 1, 9)
ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & YearlyChangelastrow))
ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & YearlyChangelastrow)) * 100
ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & YearlyChangelastrow)) * 100

Next i




Next ws

End Sub
