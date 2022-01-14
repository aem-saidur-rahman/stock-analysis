# stock-analysis

**Project Overview**

A friend, Steve is doing research on Green Energy stocks, a company that his parents are thinking of buying stocks. Purpose of this analysis is to help him take the right decision by providing some tools using Visual Basic Applications (VBA) from Microsoft Excel.

Some of the tasks included in this project were:

-Create VBA macro that triggers pop-ups and inputs which transform cell values

-Use loops (also nested loops) and conditionals to direct VBA script flow

-Apply coding skills such as syntax recollection, pattern recognition, problem decomposition, and debugging.

-Finally, refactoring was used to improve runtime.

**Results**

-----------Comparing Stock Performance---------

In the yearly comparison of stocks, it showed that in the early return (percentage increase or decrease in price from the beginning of the year to the end of the year), 11 out of 12 stocks had a positive outcome. DQ (the company in which Steve's parents are interested) had highest growth of 199.45%! But in 2018 only 2 out of 12 stocks had positive percentage. DQ's had the highest **drop** among 12 shares with 62.5%.
Which clearly suggests that now is not the right time to invest in DQ.

-----------Comparing Code Performance---------

In the initial code below tasks were performed

-Formatting the output sheet on All Stocks
 Analysis worksheet

-Activating the worksheet

-Creating title and header row

-Initializing array of all tickers

-Initializing variables for starting price and ending price

-loop through tickers to find total volume and yearly return

 -Timer and formatting

 Then some refactoring was done to make the code more efficient. Here 4 different arrays were created; tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickers array was used to establish the ticker symbol of a stock. I matched the other three arrays with the tickers array by using a variable called the tickerIndex.

 This gave a improvement in the runtime. Both the codes and run time comparison are given below;

 *Original code*
 
    '2) Initialize array of all tickers
    
    Dim tickers(12) As String

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

    '3a) Initialize variables for starting price and ending price

    Dim startingPrice As Double
    Dim endingPrice As Double

    '3b) Activate data worksheet
    
    Worksheets(yearValue).Activate

    '3c) Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row

    '4) Loop through tickers
    
    For i = 0 To 11
        ticker = tickers(i)
        TotalVolume = 0
        Worksheets(yearValue).Activate
    
       '5) loop through rows in the data
       
    For j = 2 To RowCount
    
           '5a) Get total volume for current ticker

     If Cells(j, 2).Value = ticker Then

            'increase totalVolume by the value in the current row
            TotalVolume = TotalVolume + Cells(j, 9).Value
    
    End If
    
           '5b) get starting price for current ticker

        If Cells(j - 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
            'set starting price
            startingPrice = Cells(j, 7).Value

        End If

           '5c) get ending price for current ticker
           
           If Cells(j + 1, 2).Value <> ticker And Cells(j, 2).Value = ticker Then
            'set ending price
            endingPrice = Cells(j, 7).Value

        End If

       Next j
       '6) Output data for current ticker

    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = TotalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   *Refactored*

    '1a) Create a ticker Index
         tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long

    Dim tickerStartingPrices(12) As Single

    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        
        For t = 0 To 11
        tickerVolumes(t) = 0
        Next t
        
    ''2b) Loop over all the rows in the spreadsheet.
          For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
             tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
             If (Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex)) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
          If (Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex)) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            

            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = (tickerEndingPrices(i)) / (tickerStartingPrices(i)) - 1
        
    Next i
    

**Summary**

Refactoring helps the code to become more readable with reduced complexity, however its a time consuming process; specially when the code is very long

Original VBA scprirt was detailed step by step, which is helpful for someone new to understand. However it wasn't efficient and taking more time. Refactored script obviously saved time by making it more structered. Although it saved the macro runtime but it increased the overall writing time of the script.
