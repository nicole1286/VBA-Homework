Sub Testing_HW()

'make variable for holding ticker name
Dim Ticker_Name As String

'make variable for holding stock volume
Dim Stock_Volume As Double

Stock_Volume = 0

'track each ticker name location in total volume
Dim Total_Volume_Row As Integer
Total_Volume_Row = 2

'Loop through all tickers
For i = 2 To 71226

'Check if still same ticker name
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'label ticker names
Ticker_Name = Cells(i, 1).Value

'add to total volume col
Stock_Volume = Stock_Volume + Cells(i, 7).Value

'Copy ticker name to summary
Range("I" & Total_Volume_Row).Value = Ticker_Name

'Copy stock vol to total col
Range("J" & Total_Volume_Row).Value = Stock_Volume

'Add 1 to total summary
Total_Volume_Row = Total_Volume_Row + 1

'Reset total
Stock_Volume = 0

Else
Stock_Volume = Stock_Volume + Cells(i, 7).Value

End If

Next i

End Sub
