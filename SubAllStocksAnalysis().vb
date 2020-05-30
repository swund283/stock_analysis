Sub AllStocksAnalysis()
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Dim tickers(12) As String
    Dim totalVolume(12) As Double
    Dim StartingPriceArray(12) As Single
    Dim endingPrice(12) As Single

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


    Worksheets(yearValue).Activate

    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row


    For i = 0 To 11
        
        
        ticker = tickers(i)
        totalVolume(i) = 0
        

        Worksheets(yearValue).Activate

        'loop over all the rows
        For j = 2 To RowCount

            If Cells(j, 1).Value = ticker Then

                'increase totalVolume by the value in the current row
                totalVolume(i) = totalVolume(i) + Cells(j, 8).Value
                
                

            End If

            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                StartingPriceArray(i) = Cells(j, 6).Value

            End If

            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice(i) = Cells(j, 6).Value

            End If

        Next j

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume(i)
        Cells(4 + i, 3).Value = endingPrice(i) / StartingPriceArray(i) - 1

    Next i

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            Cells(i, 3).Interior.Color = vbGreen

        Else

            Cells(i, 3).Interior.Color = vbRed

        End If

    Next i

End Sub
