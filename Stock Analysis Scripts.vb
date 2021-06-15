Sub StockAnalysis()
    Dim i As Long
    Dim LastRow As Long
    Dim TickerName As String
    Dim SummaryTable As Integer
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim Opening As Double
    Dim Closing As Double
    Dim ws as Worksheet

    For Each ws In Worksheets
        SummaryTable = 2
        TotalStockVolume = 0
        Opening = ws.Cells(2, 3).Value
    
        'Add Variables to sheet
        ws.Range("i1") = "Ticker"
        ws.Range("j1") = "Yearly Change"
        ws.Range("k1") = "Percent Change"
        ws.Range("l1") = "Total Stock Volume"

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Id and sorting ticker

        For i = 2 To LastRow
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                TickerName = ws.Cells(i, 1).Value
                
                ws.Range("I" & SummaryTable).Value = TickerName
                ws.Range("l" & SummaryTable).Value = TotalStockVolume
                Closing = ws.Cells(i, 6)
                YearlyChange = Closing - Opening
                If Opening = 0  Then
                    Opening = 1
                    Closing = Closing + 1
                End If
                PercentChange = round(((Closing - Opening) / Opening),2)
                ws.Range("J" & SummaryTable).Value = YearlyChange
                ws.Range("k" & SummaryTable).Value = PercentChange
                ws.Range("K" & SummaryTable).Style = "Percent"
                SummaryTable = SummaryTable + 1
                TotalStockVolume = 0
                Opening = ws.Cells(i + 1, 3).Value
            End If
                
            If YearlyChange < 0 Then
                ws.Range("J" & SummaryTable - 1).Interior.ColorIndex = 3

            ElseIf YearlyChange > 0 Then
                ws.Range("J" & SummaryTable - 1).Interior.ColorIndex = 4
            Else
                ws.Range("J" & SummaryTable - 1).Interior.ColorIndex = 0

            End If
            ws.Range("J1").Interior.ColorIndex = 0
            
        Next i
            
    Next ws
End Sub