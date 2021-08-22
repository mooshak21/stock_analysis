# Stock Analysis 

## Overview of Project
<p>Steve has recently graduated with a degree in finance and really wants learn about more stocks in the market. Since his parents are invested in Daqo, he wanted to create an even more in-depth analyzation of the Daily Volume of different stocks in the market, but with more of a focus on optimization. We will also only be focusing on data from 207-2018.</p>

### Purpose
Steve's main focus for our task is to modify what we have already created for him throughout the module, but to refactor the code to make it more efficient, as we will be looking at more data. Using different techniques we will see if we can make the runtime of the program faster by emplying different coding concepts. 

## Results

**Refactored VBA Code**

    Sub AllStocksAnalysisRefactored()
        Dim startTime As Single
        Dim endTime  As Single

        yearValue = InputBox("What year would you like to run the analysis on?")

        startTime = Timer
    
        'Format the output sheet on All Stocks Analysis worksheet
        Worksheets("All Stocks Analysis").Activate
    
        Range("A1").Value = "All Stocks (" + yearValue + ")"
    
        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

        'Initialize array of all tickers
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
    
        'Activate data worksheet
        Worksheets(yearValue).Activate
    
        'Get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
        '1a) Create a ticker Index
        Dim tickerIndex As Single
        tickerIndex = 0

        '1b) Create three output arrays
        ReDim tickerVolumes(12) As Long
        ReDim tickerStartingPrices(12) As Single
        ReDim tickerEndingPrices(12) As Single
    
        ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
        Next i
        
        ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
            '3c) check if the current row is the last row with the selected ticker
            'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
                '3d Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            End If
        Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    Worksheets("All Stocks Analysis").Activate
    For i = 0 To 11
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
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
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

![outcomesVlaunch](https://github.com/mooshak21/kickstarter-analysis/blob/main/Resources/OutcomesLaunchPivot.png)
![outcomesVlaunch](https://github.com/mooshak21/kickstarter-analysis/blob/main/Resources/Theater_Outcomes_vs_Launch.png "Theater Outcomes vs. Launch Date")

### Analysis of Outcomes Based on Goals
<p>We can conclude that having a fundraising goal of $35K-40K is around the max amount that would work. We can conclude this because the chart plateaus in that range and then sharply decreases in regards to percentage successful. Based on my data, the highest success rates occur in the <$1000 and $1000-$4999 ranges with 74% and 70% success rates, respectively. The price range from $40000-$44999 provides a 63% success rate as well, but there are only 8 entries within that range, so that might not be the best place to look.</p>
  


### Challenges and Difficulties Encountered
<p>No Real challenges on my end. I feel like the instructions were very clear and made the process much easier. All I had to look up was the function:<br>
COUNTIF - https://support.microsoft.com/en-us/office/countifs-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842</p>

## Results

