Sub Ticker()
'Create variable for ticker name as string
Dim TickerName as String
' Create summary table
Dim Summary_Table_Row as Long
Summary_Table_Row = 2
Dim TickerOpeningTotal as Double
Dim TickerClosingTotal as Double
Dim TickerYearlyChange as Double
Dim TickerPercentChange as Double
Dim TotalStockVolume as LongLong
TotalStockVolume = 0

'Create loop for work sheets
    For each ws in Worksheets
        Dim WorkSheetName as String
        WorkSheetName = ws.Name
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Finding the ticker and the next ticker
            'create loop for rows iterating to LastRow
            For i = 2 to LastRow
                 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ' Set Ticker name
                    TickerName = ws.Cells(i, 1).Value
                    ' Add this last ticker to the total opening and closing values
                    TickerOpeningTotal = TickerOpeningTotal + ws.Cells(i,3).Value
                    TickerClosingTotal = TickerClosingTotal + ws.Cells(i, 6).Value
                    'Get yearly change 
                    TickerYearlyChange = TickerClosingTotal - TickerOpeningTotal
                    ' Get percent change 
                    TickerPercentChange = (TickerYearlyChange / TickerOpeningTotal) * 100
                    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value 
                    
                    ws.Range("I" & Summary_Table_Row).Value = TickerName
                    ws.Range("J" & Summary_Table_Row).Value = TickerYearlyChange
                        If TickerYearlyChange < 0 Then 
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        Else
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        End If
                    'Print out ticker percent yearly change
                    ws.Range("K" & Summary_Table_Row).Value = TickerPercentChange
                    'Print out total stock volume
                    ws.Range("L" & Summary_Table_Row).Value = TotalStockVolume

                    Summary_Table_Row = Summary_Table_Row + 1  

                    'Reset Ticker values
                    TickerOpeningTotal = 0
                    TickerClosingTotal = 0
                    TickerYearlyChange = 0
                    TickerPercentChange = 0
                    TotalStockVolume = 0
                 Else 
                    TickerOpeningTotal = TickerOpeningTotal + ws.Cells(i,3).Value
                    TickerClosingTotal = TickerClosingTotal + ws.Cells(i, 6).Value
                    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value 

                End If
            Next i
    Next ws

End Sub
