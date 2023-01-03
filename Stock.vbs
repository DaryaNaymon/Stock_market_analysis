Sub Stock()

For Each ws In Worksheets
'(ws. on each Range and each Cells)

'Headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

Dim Ticker As String
Dim OpenPrice, ClosedPrice As Double
Dim Total As Double
Dim Start As Integer
Dim greatestIncNum, greatestDecNum, greatestVolNum As Double
Dim greatestIncTicker, greatestDecTicker, greatestVolTicker As String

Total = 0
Start = 2
OpenPrice = ws.Range("C2").Value

For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
    Total = Total + ws.Cells(i, "G").Value
    Ticker = ws.Cells(i, "A").Value
        If ws.Cells(i + 1, "A").Value <> ws.Cells(i, "A").Value Then
            ws.Cells(Start, "I").Value = Ticker
            ws.Cells(Start, "L").Value = Total
            ClosedPrice = ws.Cells(i, "F").Value
            ws.Cells(Start, "J").Value = ClosedPrice - OpenPrice
            If ws.Cells(Start, "J").Value > 0 Then
            ws.Cells(Start, "J").Interior.ColorIndex = 4
            Else
            ws.Cells(Start, "J").Interior.ColorIndex = 3
            End If
            If OpenPrice <> 0 Then
            ws.Cells(Start, "K").Value = FormatPercent((ClosedPrice - OpenPrice) / OpenPrice, 2)
            Else
            ws.Cells(Start, "K").Value = Null
            End If
            If ws.Cells(Start, "K").Value > greatestIncNum Then
                greatestIncTicker = ws.Cells(Start, "I").Value
                greatestIncNum = ws.Cells(Start, "K").Value
            End If
            If ws.Cells(Start, "K").Value < greatestDecNum Then
                greatestDecTicker = ws.Cells(Start, "I").Value
                greatestDecNum = ws.Cells(Start, "K").Value
            End If
            If ws.Cells(Start, "L").Value > greatestVolNum Then
                greatestVolTicker = ws.Cells(Start, "I").Value
                greatestVolNum = ws.Cells(Start, "L").Value
            End If
            
            OpenPrice = ws.Cells(i + 1, "C").Value
            Total = 0
            Start = Start + 1
            
        End If
Next i
ws.Cells(1, "P").Value = "Ticker"
ws.Cells(1, "Q").Value = "Value"
ws.Cells(2, "O").Value = "Greatest % Increase"
ws.Cells(3, "O").Value = "Greatest % Decrease"
ws.Cells(4, "O").Value = "Greatest Total Volume"
ws.Cells(2, "P").Value = greatestIncTicker
ws.Cells(2, "Q").Value = FormatPercent(greatestIncNum, 2)
ws.Cells(3, "P").Value = greatestDecTicker
ws.Cells(3, "Q").Value = FormatPercent(greatestDecNum, 2)
ws.Cells(4, "P").Value = greatestVolTicker
ws.Cells(4, "Q").Value = greatestVolNum
Next ws
End Sub

