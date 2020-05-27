
Sub AllStocksAnalysis()
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Dim tickers(12) As String
    Dim volume(12)
    Dim startPrice(12)
    Dim endPrice(12)
    
    
'ticker array
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
    
'volume array
    volume(0) = "AY"
    volume(1) = "CSIQ"
    volume(2) = "DQ"
    volume(3) = "ENPH"
    volume(4) = "FSLR"
    volume(5) = "HASI"
    volume(6) = "JKS"
    volume(7) = "RUN"
    volume(8) = "SEDG"
    volume(9) = "SPWR"
    volume(10) = "TERP"
    volume(11) = "VSLR"
    
   'starting price array
    startPrice(0) = "AY"
    startPrice(1) = "CSIQ"
    startPrice(2) = "DQ"
    startPrice(3) = "ENPH"
    startPrice(4) = "FSLR"
    startPrice(5) = "HASI"
    startPrice(6) = "JKS"
    startPrice(7) = "RUN"
    startPrice(8) = "SEDG"
    startPrice(9) = "SPWR"
    startPrice(10) = "TERP"
    startPrice(11) = "VSLR"
    
'end price array
    endPrice(0) = "AY"
    endPrice(1) = "CSIQ"
    endPrice(2) = "DQ"
    endPrice(3) = "ENPH"
    endPrice(4) = "FSLR"
    endPrice(5) = "HASI"
    endPrice(6) = "JKS"
    endPrice(7) = "RUN"
    endPrice(8) = "SEDG"
    endPrice(9) = "SPWR"
    endPrice(10) = "TERP"
    endPrice(11) = "VSLR"


    Worksheets(yearValue).Activate

    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    Dim startingPrice As Single
    Dim endingPrice As Single

        For i = 0 To 11

        ticker = tickers(i)
        totalVolume = 0

        Worksheets(yearValue).Activate

        'loop over all the rows
        
        For j = 2 To RowCount

            If Cells(j, 1).Value = ticker Then

                'increase totalVolume by the value in the current row
                totalVolume = totalVolume + Cells(j, 8).Value

            End If

            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                startingPrice = Cells(j, 6).Value

            End If

            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value

            End If

        Next j

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

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
