Sub NewAllStocksAnalysis()
    Worksheets("All Stocks Analysis").Activate
    'prompt box
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'adds a year header to the analysis page based on the initial input
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    'array for tickers
    Dim tickers(12) As String
    
    '3 STORAGE arrarys for output value.  made totalVolumeArray very large to prevent overflow errors i was getting...
    Dim totalVolumeArray(12) As Double
    Dim startingPriceArray(12) As Single
    Dim endingPriceArray(12) As Single
    
    'defining the names of the tickers as string
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

    'opens the correct tab form the use input in prompt box
    Worksheets(yearValue).Activate

    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
   
   'setting the volume for ALL 12 tickers to zero for the totalVolumeArray initial values.  This needs to happen outside the loop
   For v = 0 To 11
        totalVolumeArray(tickerIndex) = 0
   Next v

    
    'define the starting place in the array
    tickerIndex = 0
    
    
    'go row by row
    For j = 2 To RowCount
        
            'increase totalVolumeArray by the value in the current row dont need if statement anymore -  im storing it in the array as we go row by row
            totalVolumeArray(tickerIndex) = totalVolumeArray(tickerIndex) + Cells(j, 8).Value
                        
                        
                     'logs starting price if value before is different
                    If Cells(j - 1, 1).Value <> Cells(j, 1).Value Then
        
                        startingPriceArray(tickerIndex) = Cells(j, 6).Value
        
                    End If
        
                    '2 things: logs ending price if value after different 2) raises the ticker index # by one when the name changes in row 1, only works beacuse its sorted
                    If Cells(j + 1, 1).Value <> Cells(j, 1).Value Then
        
                        endingPriceArray(tickerIndex) = Cells(j, 6).Value
                        
                        tickerIndex = tickerIndex + 1
                        
                    End If
                    
                   
        
        Next j


        'output of values from the arrays - need a cleaner formula via loop
    For o = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + o, 1).Value = tickers(o)
        Cells(4 + o, 2).Value = totalVolumeArray(o)
        Cells(4 + o, 3).Value = endingPriceArray(o) / startingPriceArray(o) - 1
    Next o

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