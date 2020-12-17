Sub Ticker()
   For Each ws In Worksheets
        Dim TickerName As String
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        Dim TickerYearlyChange As Double
        Dim TickerPercentChange As Double
        Dim TotalStockVolume As LongLong
        TotalStockVolume = 0
        Dim OpeningValue As Double
        Dim ClosingValue As Double
        Dim TickerHeader As String
        Dim YearlyChangeHeader As String
        Dim PercentChangeHeader As String
        Dim TotalStockVolumeHeader As String
        Dim GreatestIncrease As Double
        Dim GreatestDecrease As Double
        Dim GreatestTotalVolume As LongLong
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestTotalVolume = 0
        TickerHeader = "Ticker"
        YearlyChangeHeader = "Yearly Change"
        PercentChangeHeader = "Percent Change"
        TotalStockVolumeHeader = "Total Stock Volume"

        ws.Cells(1, 9).Value = TickerHeader
        ws.Cells(1, 10).Value = YearlyChangeHeader
        ws.Cells(1, 11).Value = PercentChangeHeader
        ws.Cells(1, 12).Value = TotalStockVolumeHeader
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = TickerHeader
        ws.Range("Q1").Value = "Value"
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To LastRow
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    OpeningValue = ws.Cells(i, 3).Value

            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerName = ws.Cells(i, 1).Value
                ClosingValue = ws.Cells(i, 6)
                TickerYearlyChange = ClosingValue - OpeningValue

                If OpeningValue <> 0 Then
                    TickerPercentChange = ((ClosingValue - OpeningValue) / OpeningValue)
                Else
                    TickerPercentChange = 0
                End If
                    
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                ws.Range("I" & Summary_Table_Row).Value = TickerName
                ws.Range("J" & Summary_Table_Row).Value = TickerYearlyChange

                    If TickerYearlyChange < 0 Then
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                    End If
                    
                ws.Range("K" & Summary_Table_Row).Value = TickerPercentChange
                ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume

                If TickerPercentChange > GreatestIncrease Then
                    GreatestIncrease = TickerPercentChange
                    ws.Range("P2").Value = TickerName
                    ws.Range("Q2").Value = GreatestIncrease
                End If

                If TickerPercentChange < GreatestDecrease Then
                    GreatestDecrease = TickerPercentChange
                    ws.Range("P3").Value = TickerName
                    ws.Range("Q3").Value = GreatestDecrease
                End If

                If TotalStockVolume > GreatestTotalVolume Then
                    GreatestTotalVolume = TotalStockVolume
                    ws.Range("P4").Value = TickerName
                    ws.Range("Q4").Value = GreatestTotalVolume
                End If

                Summary_Table_Row = Summary_Table_Row + 1

                TickerYearlyChange = 0
                TickerPercentChange = 0
                TotalStockVolume = 0
                OpeningValue = 0
                ClosingValue = 0

            Else
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
End Sub