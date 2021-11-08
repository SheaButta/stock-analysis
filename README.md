# Stock Analysis using Excel VBA

Dataset: [VBA Challenge - Stock Analysis](https://github.com/SheaButta/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Overview of Project

### Purpose
The purpose of this project is to analyze existing stock data and refactor legacy VBA code to increase 
processing performance. Although legacy VBA code processed the data in just over one (1) second; this effort 
will visualize a performance gain. The datasets are separated by two (2) worksheets in the "VBA Challenge - Stock Analysis" file.  
The two worksheets are;
- 2017
- 2018


## Results

### Analysis of 2017 and 2018 Refactoring
Using the 2017 dataset, the refactoring of the VBA code visualized all stocks (except one (1)) with returns of 5% or 
greater. The execution time for the original code was 1 (1.063) second while the refactored code completed in under 1 second.

![2017 Refactored Performance Gain](https://github.com/SheaButta/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

Using the 2018 dataset, the refactoring of the VBA code visualized two (2) stocks with returns over 80%.  The other stocks had no return and
may suggest to sell.  The execution time for the orignal VBA code was 1 second, while the refactored code improved to under 1 second.

![2018 Refactored Performance Gain](https://github.com/SheaButta/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)

The below refactored code included four (4) code changes which is documented below. 

    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
          If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
          
               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

           End If
           
                
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        
           If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

               tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

           End If
            

            '3d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
                       
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

## Summary

### Refactoring Stock Analysis
In summary, refactoring the VBA code proved to be an advantage because of the visual performance gain processing the 
2017 and 2018 stock data.  Documenting the edits, using "arrays" and "for loops" in the VBA code were critical additions 
to this performance gain.  The color coding of positve and negative returns also makes this effort worthwhile as it would 
be very appealing to senior management.  Although the code may look more complex since the updates, having clear documenation 
will only help with future development.  







