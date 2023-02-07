# VBA-challenge
Jessica Hartman - VBA Module 2 Challenge

Sub multiYearStockData()

   ' Declare and set worksheet
    Dim ws As Worksheet
    
    ' Loop through all stocks for one year
    For Each ws In Worksheets
    
    ' Create the column headings
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ws.Range("A1:O4").Columns.AutoFit
     
     ' Define Ticker variable
    Dim Ticker As String
    Ticker = " "
    
    ' Dim tickerVolume As Double
     ' tickerVolume = 0    ' Is this different than stock volume?
    
    ' Create variable to hold stock volume
    Dim stockVolume As Double
    stockVolume = 0
    
     ' variable to hold the new rows for ticker and volume
    Dim nRows As Integer
    nRows = 2 ' the first row to populate in new columns will be row 2
    
    ' Set initial and last row for worksheet
    Dim lastRow As Long    ' why would we dim this as long?
    
    ' Define lastRow of worksheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set new variables for prices and percent changes
    Dim openPrice As Double
    openPrice = 0
    Dim closePrice As Double
    closePrice = 0
    Dim priceChange As Double
    priceChange = 0
    Dim priceChangePercent As Double
    priceChangePercent = 0
    
    ' Trying to populate ticker output
    ' Dim tickerRow As Long
    ' tickerRow = 1
    
    ' Do loop of current worksheet to last row
    For i = 2 To lastRow
    
    ' Ticker symbol output
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ' set the ticker name
    Ticker = ws.Cells(i, 1).Value
    
    ' add the stock volume total
    stockVolume = stockVolume + ws.Cells(i, 7).Value
    
    ' display the ticker name in new column
    ws.Cells(nRows, 9).Value = Ticker
    
    ' display the stock volume total in new column
    ws.Cells(nRows, 12).Value = stockVolume
    
    ' add 1 to the stock volume to go to next row
    ' newRows = newRows + 1
    
    ' reset the total volume for the next ticker
    stockVolume = 0
    Else
    
    ' if there is no change in the ticker, keep adding the stock volume
    stockVolume = stockVolume + ws.Cells(i, 7).Value
      
      ' Calculate change in price
      closePrice = closePrice + ws.Cells(i, 6).Value
      openPrice = openPrice + ws.Cells(i, 3).Value
     
     ' Calculate the difference in closing price from opening price
     priceChange = closePrice - openPrice
     
     ' display the price change in new column J
     ws.Cells(nRows, 10).Value = priceChange
     
     '  Calculate percent change from opening to closing price
     priceChangePercent = (priceChange / openPrice) * 100      ' why? come back to this
     
     ' display the percent change in new column K
     ws.Cells(nRows, 11).Value = priceChangePercent
     
     ' add 1 to the change to the price change
     ' newRows = newRows + 1
      
    End If
    
    Next i
    
    Next ws
    
End Sub
