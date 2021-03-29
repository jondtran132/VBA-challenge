Sub stockAnalysis()
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    currentRow = 2
    openPrice = Cells(2, 3).Value
    stockVol = 0
    
    gIncTicker = Cells(2, 1).Value
    gIncVal = 0
    gDecTicker = Cells(2, 1).Value
    gDecVal = 0
    gVolTicker = Cells(2, 1).Value
    gVolVal = stockVol
    
    
    For i = 2 To lastRow
        stockVol = stockVol + Cells(i, 7).Value
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            Cells(currentRow, 9).Value = Cells(i, 1).Value
            
            closePrice = Cells(i, 6).Value
            
            If openPrice <> 0 Then
                Cells(currentRow, 11).Value = (closePrice - openPrice) / openPrice
            Else
                Cells(currentRow, 11).Value = (closePrice - openPrice)
            End If
            Cells(currentRow, 11).NumberFormat = "0.00%"
            
            Cells(currentRow, 10).Value = Round(closePrice - openPrice, 2)
            If Cells(currentRow, 10).Value < 0 Then
                Cells(currentRow, 10).Interior.ColorIndex = 3
                If Cells(currentRow, 11).Value < gDecVal Then
                    gDecVal = Cells(currentRow, 11).Value
                    gDecTicker = Cells(i, 1).Value
                End If
            Else
                Cells(currentRow, 10).Interior.ColorIndex = 4
                If Cells(currentRow, 11).Value > gIncVal Then
                    gIncVal = Cells(currentRow, 11).Value
                    gIncTicker = Cells(i, 1).Value
                End If
            End If
            
            Cells(currentRow, 12).Value = stockVol
            
            If stockVol > gVolVal Then
                gVolVal = stockVol
                gVolTicker = Cells(i, 1).Value
            End If
            
            openPrice = Cells(i + 1, 6).Value
            stockVol = 0
            currentRow = currentRow + 1
        End If
    Next i
    
    Range("P2").Value = gIncTicker
    Range("Q2").Value = gIncVal
    Range("Q2").NumberFormat = "0.00%"
    Range("P3").Value = gDecTicker
    Range("Q3").Value = gDecVal
    Range("Q3").NumberFormat = "0.00%"
    Range("P4").Value = gVolTicker
    Range("Q4").Value = gVolVal

End Sub

