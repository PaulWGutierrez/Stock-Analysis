# Steve's 2017 and 2018 Stock-Analysis
## 1.) Overview of Project

#### The overview of the project was to help Steve with his parents investing into DAQO New Energy Corp stock. He was not sure if that was the right direction to put all of their investing into so he wanted to research other green energy stocks including DAQO. Within this research the yearly return between 2017 and 2018 will be calculated and Steve will verify if the return increased or decreased between those two years. During thisproject the Visual Basic for Applications "VBA" tool will be used to help Steve get solutions.

## 2.) Results

#### There are major differences between both years. Overall the majority of the stocks made positive returns in 2017 except for ticker "TERP". While in 2018 the majority of the stocks made a negative return except for tickers "ENPH" and "RUN", but "ENPH" still had a decrease in its return being 129.5% to 81.9%. While "RUN" made a major increase in its return being 5.5% to 84.0%.
<img width="560" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/86431959/125217774-7a994d80-e28f-11eb-994b-6d4934b51702.png">
<img width="559" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/86431959/125217777-7d943e00-e28f-11eb-8f08-ae39e8ddf604.png">

#### When it comes to the running time of the code the refactored script ran much faster for both 2017 and 2018 analysis.
##### 2017 Orginal Script
<img width="266" alt="2017 1st run time" src="https://user-images.githubusercontent.com/86431959/125218074-1fb42600-e290-11eb-954c-8bc38e68a5ab.png">

##### 2017 Refactored Script
<img width="270" alt="2017 2nd run time" src="https://user-images.githubusercontent.com/86431959/125218092-2c387e80-e290-11eb-88f4-09ff2c888b2f.png">

##### 2018 Orginal Script
<img width="274" alt="2018 1st run time" src="https://user-images.githubusercontent.com/86431959/125218135-4a9e7a00-e290-11eb-9de4-184bda7d6c57.png">

##### 2018 Refactored Script
<img width="276" alt="2018 2nd run time" src="https://user-images.githubusercontent.com/86431959/125218145-4d996a80-e290-11eb-8e8c-9521ecbe0192.png">

#### Sample of "All Stock Analysis" Code:

Sub AllStocksAnalysis()
    Dim startTime As Single
    Dim endTime As Single
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
        startTime = Timer
        
    '1) Format the output sheet on the "All Stocks Analysis" worksheet.
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks"
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    '2) Initialize an array of all tickers.
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
    '3a) Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingPrice As Single
    '3b) Activate the data worksheet.
    Worksheets("2018").Activate
    '3c) Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '4) Loop through the tickers.
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
        '5) Loop through rows in the data.
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
            '5a) Get the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then

                 totalVolume = totalVolume + Cells(j, 8).Value
        
            End If
            '5b) get the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                 startingPrice = Cells(j, 6).Value
        
            End If
        
            '5c) get the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value
        
            End If
        Next j
        '6) Output the data for the current ticker.
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
        
        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
End Sub
