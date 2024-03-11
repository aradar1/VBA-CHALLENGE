
Sub Ticker()

'Dim ws As Worksheets


For Each ws In Worksheets

    Dim WorksheetName As String
    Dim Ticker, TickerRow, Volume As Double
    Dim OpenPrice, ClosePrice, PriceChangePercent, YearlyChange As Double
    Dim GreatestInc, GreatestDec, GreatestVol As Double
    Ticker = 0
    TickerRow = 0
    Volume = 0
    PriceChangePercent = 0
    GreatestInc = 0
    GreatestDec = 0
    GreatestVol = 0
    TickerRow = 1
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    WorksheetName = ws.Name
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    
    OpenPrice = ws.Cells(2, 3).Value
    
    For i = 2 To LastRow
    
        Volume = Volume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            TickerRow = TickerRow + 1
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(TickerRow, 9).Value = Ticker
            ws.Cells(TickerRow, 12).Value = Volume
            ClosePrice = ws.Cells(i, 6).Value 'Set Close Price
            YearlyChange = ClosePrice - OpenPrice 'Set YearlyChange
            ws.Cells(TickerRow, 10).Value = YearlyChange 'Output Yearly Change
            PriceChangePercent = (YearlyChange / OpenPrice)
            ws.Cells(TickerRow, 11).Value = PriceChangePercent
            If PriceChangePercent > GreatestInc Then
                GreatestInc = PriceChangePercent
                ws.Cells(2, 16) = Ticker
                ws.Cells(2, 17) = GreatestInc
            End If
            If (PriceChangePercent < GreatestDec) Then
                GreatestDec = PriceChangePercent
                ws.Cells(3, 16) = Ticker
                ws.Cells(3, 17) = GreatestDec
            End If
            If Volume > GreatestVol Then
                GreatestVol = Volume
                ws.Cells(4, 16) = Ticker
                ws.Cells(4, 17) = GreatestVol
            End If
            
            OpenPrice = ws.Cells(i + 1, 3).Value 'New Open Price
            Volume = 0 'Reset variable for next ticker
       ' ElseIf Left(ws.Cells(i + 1, 2), 4) <> Left(ws.Cells(i, 2), 4) Then
            'ClosePrice = ws.Cells(i, 6).Value
            'OpenPrice = ws.Cells(i + 1, 3).Value 'New Open Price
            
       ' ElseIf OpenPrice <> 0 Then
       ' PriceChangePercent = (PriceChangePercent / OpenPrice) * 100
        End If
    Next i
    
    For i = 2 To LastRow
        ws.Cells(i, 10).Style = "Currency" 'Changes to percent with Decimal Places
        'ws.Cells(i, 11).Style = "Percent"
        ws.Cells(i, 11).Value = Format(ws.Cells(i, 11).Value, "0.00%")
        If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        ws.Cells(i, 11).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        ws.Cells(i, 11).Interior.ColorIndex = 3
        ElseIf ws.Cells(i, 10).Value = 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 15
        ws.Cells(i, 11).Interior.ColorIndex = 15
        End If
    Next i
        ws.Range("Q2").Interior.ColorIndex = 4
        ws.Range("Q2").Value = Format(ws.Range("Q2").Value, "0.00%")
        ws.Range("Q3").Interior.ColorIndex = 3
        ws.Range("Q3").Value = Format(ws.Range("Q3").Value, "0.00%")
        ws.Range("Q4").Interior.ColorIndex = 15
Next ws
            
    
    
    
End Sub






