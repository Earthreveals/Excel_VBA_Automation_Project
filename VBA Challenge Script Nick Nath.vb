Sub Stocktickeranalysis()

For Each ws In Worksheets

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percentage Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Variant
Dim Total_Stock_Volume As Variant
Dim Summary_Table As Integer
Dim MaxIncrease As Variant
Dim MaxDecrease As Variant
Dim MaxVolume As Variant
Summary_Table = 2
lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Ticker_open = ws.Cells(2, 3).Value
Total_Stock_Volume = 0

' Loop through each row with data
   
For i = 2 To lastRow
'Calculate Total Stock Volume

Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
'Calculate Yearly Change

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

        Ticker = ws.Cells(i, 1).Value
        Ticker_close = ws.Cells(i, 6).Value
        Yearly_Change = Ticker_close - Ticker_open
'Calculate percent yearly change

If Ticker_open <> 0 Then
Percent_Change = (Yearly_Change / Ticker_open)
Else
Percent_Change = 0

End If

'Output the metrics for the last row

ws.Cells(Summary_Table, 9).Value = Ticker
ws.Cells(Summary_Table, 10).Value = Yearly_Change
ws.Cells(Summary_Table, 11).Value = Percent_Change * 100
ws.Cells(Summary_Table, 12).Value = Total_Stock_Volume
ws.Cells(Summary_Table, 11).Value = FormatPercent(Percent_Change)

'Conditional formatting to yearly change and percent change column

If Yearly_Change > 0 Then
ws.Cells(Summary_Table, 10).Interior.Color = RGB(0, 255, 0)

Else
ws.Cells(Summary_Table, 10).Interior.Color = RGB(255, 0, 0)

End If

Summary_Table = Summary_Table + 1
Ticker_open = ws.Cells(i + 1, 3).Value
Total_Stock_Volume = 0

End If

Next i

'Find the greatest % increase, % decrease, and total volume
MaxIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table))
MaxDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table))
MaxVolume = Application.WorksheetFunction.Max(ws.Range("L2:K" & Summary_Table))

'Find the tickers associated with the max values
maxIncreaseTicker = ws.Cells(Application.WorksheetFunction.Match(MaxIncrease, ws.Range("K2:K" & Summary_Table), 0) + 1, 9).Value
maxDecreaseTicker = ws.Cells(Application.WorksheetFunction.Match(MaxDecrease, ws.Range("K2:K" & Summary_Table), 0) + 1, 9).Value
maxVolumeTicker = ws.Cells(Application.WorksheetFunction.Match(MaxVolume, ws.Range("L2:L" & Summary_Table), 0) + 1, 9).Value


'Output the greatest % increase, decrease and total volume

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(2, 16).Value = maxIncreaseTicker
ws.Cells(3, 16).Value = maxDecreaseTicker
ws.Cells(4, 16).Value = maxVolumeTicker
ws.Cells(2, 17).Value = FormatPercent(MaxIncrease)
ws.Cells(3, 17).Value = FormatPercent(MaxDecrease)
ws.Cells(4, 17).Value = MaxVolume
ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"
Next ws

End Sub

