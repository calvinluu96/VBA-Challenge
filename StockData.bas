Attribute VB_Name = "StockData"
Sub Ticker()
    On Error Resume Next
    
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Add headers to columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Volume"
    
    ' Create counter so it knows how many ticker symbols to list
    Dim NumTickers As Integer
    NumTickers = 1
    
    ' Year open => (first row i, column 3)
    ' Year close => (last row i, column 6)
    Dim open_price As Double
    Dim close_price As Double
    
    ' Create variable to store total volume for a given ticker
    Dim total_volume As LongLong
    ' initialize total volume to 0
    total_volume = 0
    
    ' store initial opening price
    open_price = Cells(2, 3).Value
    
    For I = 2 To LastRow
        ' Update volume using by getting sum from column G (7th column)
        total_volume = total_volume + Cells(I, 7).Value
        
        If open_price = 0 Then
            open_price = Cells(I, 3).Value
        End If
        
        ' Check if the ticker symbol is different from the next row symbol
        If (Cells(I, 1).Value <> Cells(I + 1, 1)) Then
            ' Store closing price
            close_price = Cells(I, 6).Value
            
            ' Move down to next row
            NumTickers = NumTickers + 1
            
            ' Print new ticker symbols into column I
            Cells(NumTickers, 9).Value = Cells(I, 1).Value
            
            ' Calculate Yearly Change and print into column J
            Cells(NumTickers, 10).Value = close_price - open_price
            
            ' For Yearly Change, set fill color to green (4) if positive, red (3) if negative
            If Cells(NumTickers, 10).Value < 0 Then
                Cells(NumTickers, 10).Interior.ColorIndex = 3
            Else
                Cells(NumTickers, 10).Interior.ColorIndex = 4
            End If
            
            ' Calculate Percent Change = Yearly Change / Open Price and print into column K
            Cells(NumTickers, 11).Value = Cells(NumTickers, 10).Value / open_price
            
            ' Format Percent Change into 0.00%
            Cells(NumTickers, 11).NumberFormat = "0.00%"
            
            ' Print Total Volume into column L
            Cells(NumTickers, 12).Value = total_volume
            
            ' reset open_price to value in next row
             open_price = Cells(I + 1, 3).Value
            
            ' reset Total Volume to 0 for next ticker
            total_volume = 0
        End If
    Next I
    
    ' Add headers and row titles
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    
    
    ' Create and initialize variables for Greatest Changes and Total Volumes
    Dim max_increase As Double
    Dim max_decrease As Double
    Dim max_volume As LongLong
    
    ' Determine which Tickers are the maximums for each category
    For j = 2 To NumTickers
        ' For checking max increase
        If Cells(j, 11).Value > max_increase Then
            max_increase = Cells(j, 11).Value
            Cells(2, 15).Value = Cells(j, 9).Value
            Cells(2, 16).Value = Cells(j, 11).Value
            Cells(2, 16).NumberFormat = "0.00%"
        End If
        ' For checking max decrease
        If Cells(j, 11).Value < max_decrease Then
            max_decrease = Cells(j, 11).Value
            Cells(3, 15).Value = Cells(j, 9).Value
            Cells(3, 16).Value = Cells(j, 11).Value
            Cells(3, 16).NumberFormat = "0.00%"
        End If
        ' For checking max increase
        If Cells(j, 12).Value > max_volume Then
            max_volume = Cells(j, 12).Value
            Cells(4, 15).Value = Cells(j, 9).Value
            Cells(4, 16).Value = Cells(j, 12).Value
        End If
    Next j
                      
End Sub
Sub WorksheetLoop()

    Dim WS_Count As Integer
    Dim I As Integer
    
    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
    
    ' Begin the loop.
    For I = 1 To WS_Count
        Worksheets(I).Select
        Call Ticker
    Next I

End Sub
