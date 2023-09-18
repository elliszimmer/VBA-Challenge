Sub StockAnalysisLoop()

    For Each ws In Worksheets
        ws.Activate

        'Variable for holding Ticker
        Dim Ticker As String
        
        'Variable for holding totals
        Dim TotalVolume As Double
        
        'Create Summary Table
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        'Track location for each in summary table
        Dim TableRow As Integer
        TableRow = 2
            OpenPricePointer = 2
            TotalVolume = 0
        
        'Loop through all tickers for yearly change
         
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To lastrow
        'See if we are still in the same ticker, if not:
            If Cells(i, "A").Value <> Cells(i + 1, "A").Value Then
        'Set the ticker
                Ticker = Cells(i, "A").Value
        'Add to the Total
                OpenPrice = Cells(OpenPricePointer, "C").Value
                ClosePrice = Cells(i, "F").Value
                TotalVolume = TotalVolume + Cells(i, "G").Value
        'Print the Ticker in the summary table
                Cells(TableRow, "I").Value = Ticker
        'Print the yearly change total
                Cells(TableRow, "J").Value = ClosePrice - OpenPrice
        'Print the percent change
                Cells(TableRow, "K").Value = "%" & Round((ClosePrice - OpenPrice) / OpenPrice * 100, 2)
        'Print the Total Stock Volume
                Cells(TableRow, "L").Value = TotalVolume
        'Add one to thhe summary table row
                TableRow = TableRow + 1
        'Reset the totals
                TotalVolume = 0
                OpenPricePointer = i + 1
        'If the cell immediately following a row is the same ticker:
            Else
                TotalVolume = TotalVolume + Cells(i, "G").Value
            End If
            
        Next i
        
        
        'greatest%increase greatest%decrease and greatest total volume fields
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
        lastrow = Cells(Rows.Count, "I").End(xlUp).Row
        
        GreatestIncrease = 0
        GreatestIncreaseTicker = ""
        GreatestDecrease = 0
        GreatestDecreaseTicker = ""
        GreatestVolume = 0
        GreatestVolumeTicker = ""
        
        For i = 2 To lastrow
            If Cells(i, "K").Value > GreatestIncrease Then
                GreatestIncrease = Cells(i, "K").Value
                GreatestIncreaseTicker = Cells(i, "I").Value
            
            End If
            
            If Cells(i, "K").Value < GreatestDecrease Then
                GreatestDecrease = Cells(i, "K").Value
                GreatestDecreaseTicker = Cells(i, "I").Value
            
            End If
            
            If Cells(i, "L").Value > GreatestVolume Then
                GreatestVolume = Cells(i, "L").Value
                GreatestVolumeTicker = Cells(i, "I").Value
                
            End If
             
        Next i
        
        Range("P2").Value = GreatestIncrease
        Range("O2").Value = GreatestIncreaseTicker
        Range("P3").Value = GreatestDecrease
        Range("O3").Value = GreatestDecreaseTicker
        Range("P4").Value = GreatestVolume
        Range("O4").Value = GreatestVolumeTicker
                            
        'use conditional formatting to highlight Yearly Changes (positive and negative changes)
        For i = 2 To lastrow
            If Cells(i, "J").Value > 0 Then
                Cells(i, "J").Interior.ColorIndex = 4
            ElseIf Cells(i, "J").Value < 0 Then
                Cells(i, "J").Interior.ColorIndex = 3
            End If
        Next i
        
        For i = 2 To lastrow
            If Cells(i, "K").Value > 0 Then
               Cells(i, "K").Interior.ColorIndex = 4
            ElseIf Cells(i, "K").Value < 0 Then
                Cells(i, "K").Interior.ColorIndex = 3
            End If
        Next i
                        
        Columns("A:P").AutoFit
    
    'run on every worksheet at once
    Next ws

    MsgBox ("Complete!")
    
End Sub