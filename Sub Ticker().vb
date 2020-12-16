Sub Ticker()
'Create variable for ticker name as string
Dim TickerName As String
' Create summary table
Dim Summary_Table_Row As Long
Summary_Table_Row = 2
Dim TickerYearlyChange As Double
Dim TickerPercentChange As Double
Dim TotalStockVolume As LongLong
Dim OpeningValue As Double
Dim ClosingValue As Double
TotalStockVolume = 0





'
   ' For Each ws In Worksheets
       ' Dim WorkSheetName As String
       ' WorkSheetName = ws.Name
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To LastRow
                If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                    OpeningValue = Cells(i, 3).Value

                 ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    
                    TickerName = Cells(i, 1).Value
                
                    ClosingValue = Cells(i, 6)
                    
                    TickerYearlyChange = ClosingValue - OpeningValue
                    If OpeningValue <> 0 Then
                        TickerPercentChange = ((ClosingValue - OpeningValue) / OpeningValue)
                    Else
                        TickerPercentChange = 0
                    End If
                    
                    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
                    
                    Range("I" & Summary_Table_Row).Value = TickerName
                    Range("J" & Summary_Table_Row).Value = TickerYearlyChange
                        If TickerYearlyChange < 0 Then
                            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                        Else
                            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        End If
                    
                    Range("K" & Summary_Table_Row).Value = TickerPercentChange
                    Range("L" & Summary_Table_Row).Value = TotalStockVolume

                    Summary_Table_Row = Summary_Table_Row + 1

                    TickerYearlyChange = 0
                    TickerPercentChange = 0
                    TotalStockVolume = 0
                    OpeningValue = 0
                    ClosingValue = 0
                 Else
                    TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

                End If
            Next i
   ' Next ws

End Sub
