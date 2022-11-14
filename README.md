# Module_2_stock-analysis
## Overview of Project
Utilizing VBA Steve has asked for help to analyze stock data to convince his parents to diversify their funds.  After writing code for Steve, He was able to analyze the data from 2017 and 2018 using a button created in the data sheet.  Steve was very happy with the work that was done and wanted to use the sheet to analyze larger data sets.  

Having concern about the code taking a long time to execute, we refactored the code to compare difference in time needed to run the code.  We checked the results of the output data to ensure the changes in code still produced accurate data.      
## Results
After refactoring the code to find and compute the data as it is looping through in one single pass we were able to cut the run time down significantly.  In the original code (shown first) there were two variables i and j.  For each i the code would set it to 0 then run though each j and output the data before returning to the next i and then repeat those steps until all i's where computed.  In the refactored code (shown second) we eliminated the j variable which allowed the code to run through just one time the rows and pull the needed data, thus reducing the run time. The screenshots below show the difference in run time.  

#### Original Code

```ruby

Sub AllStocksAnalysis()

'1) Format the output sheet on the "All Stocks Analysis" worksheet.

Worksheets("All Stocks Analysis").Activate

Cells(1, 1).Value = "All Stocks (2018)"
Cells(3, 1).Value = "Ticker"
Cells(3, 2).Value = "Total Daily Volume"
Cells(3, 3).Value = "Return"

'2) Initialize an array of all tickers.
Dim tickers(11) As String
    
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


'3a) Initialize variables for the starting price and ending price.
Dim startingPrice As Double
Dim endingPrice As Double

'3b) Activate the data worksheet.
Worksheets("2018").Activate

'3c) Find the number of rows to loop over.

RowCount = Cells(Rows.Count, "A").End(xlUp).Row

'4) Loop through the tickers.
For i = 0 To 11
    
    ticker = tickers(i)
    
    totalVolume = 0
'5) Loop through rows in the data
Worksheets("2018").Activate

For j = 2 To RowCount

'5a) Find the total volume for the current ticker.
If Cells(j, 1).Value = ticker Then

    totalVolume = totalVolume + Cells(j, 8).Value
    
End If

'5b) Find the starting price for the current ticker.
If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

    startingPrice = Cells(j, 6).Value    
    
End If

'5c) Find the ending price for the current ticker.

If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

    endingPrice = Cells(j, 6).Value
    
End If

Next j

'6) Output the data for the current ticker.


Worksheets("All Stocks Analysis").Activate

Cells(4 + i, 1).Value = ticker
Cells(4 + i, 2).Value = totalVolume
Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1

Next i    

End Sub
```

#### Refactored Code

```ruby
    
    '1a) Create a ticker Index
    
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
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If

            '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
          
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
    
```
    
### Elapsed times for 2018 Original and Refactored codes
![VBA_Challenge_2018](https://user-images.githubusercontent.com/116597421/201556210-3d2d99d0-cfde-40bd-812c-18a2bb4b6011.PNG)
![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/116597421/201556221-c0b9f1d8-2cf5-4db2-acd5-3352ab8b99e5.PNG)


## Summary
There are both advantages and disadvantages of refactoring code.  Some of the advantages are that the code is usually condensed and can be easier to find mistakes.  Another is that refactoring usually decreases the run time, therefore allowing larger data sets to be run quickly and not tying up the computer.  Refactoring also improves readability.  Some of the disadvantages of refactoring are that it can be expensive to refactor legacy code.  Also, refactoring can be very time consuming.  

Within the VBA script there are some advantages and disadvantages also.  In the original script it was easier to follow and less implied logic.  Each variable was run one during each loop.  But, this was what took the original more time to run.  In the refactored script the code called for only one loop through the data.  In doing so any errors later in the script would affect functioning earlier on.  This caused time to find and fix the errors as they were not easily pointed out when debugging.  
