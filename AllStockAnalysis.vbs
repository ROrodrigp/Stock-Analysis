Sub AllStockAnalysis()

    Worksheets("All Stock Analysis").Activate
    range("A1").Value = "All Stocks (2018)"
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    
    Dim tickers(11) As String
    Dim rowEnd As Integer
    Dim rangeFound As range
    Dim sh As Worksheet
    Dim startingPrice As Double
    Dim endingprice As Double
    
    Set sh = ThisWorkbook.Sheets("2018")
    Set rangeFound = sh.range("A1")
    rowStart = 2
    
    'Find the number of rows to loop over
    rowEnd = sh.range(rangeFound, rangeFound.End(xlDown)).Rows.Count
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    
    
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        For j = rowStart To rowEnd
        
            Worksheets("2018").Activate
            
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                endingprice = Cells(j, 6).Value
            End If
            
        Next j
        
        Worksheets("All Stock Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingprice / startingPrice - 1
    Next i

End Sub
