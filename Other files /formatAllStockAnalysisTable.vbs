Sub formatAllStockAnalysisTable()

    Worksheets("All Stock Analysis").Activate
    range("A3:C3").Font.Bold = True
    range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    range("B4:B15").NumberFormat = "$#,##0.00"
    range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3).Value > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3).Value < 0 Then
            Cells(i, 3).Interior.Color = vbRed
        Else
            Cells(i, 3).Interior.Color = xlNone
        End If
        
    Next i
    
End Sub
