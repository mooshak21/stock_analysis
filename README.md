# Stock Analysis 

## Overview of Project
<p>Steve has recently graduated with a degree in finance and really wants learn about more stocks in the market. Since his parents are invested in Daqo, he wanted to create an even more in-depth analyzation of the Daily Volume of different stocks in the market, but with more of a focus on optimization. We will also only be focusing on data from 2017-2018.</p>

### Purpose

Steve's main focus for our task is to modify what we have already created for him throughout the module, but to refactor the code to make it more efficient, as we will be looking at more data. Using different techniques we will see if we can make the runtime of the program faster by emplying different coding concepts. 

---

## Results

### **Refactored VBA Code**

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

--
### **Comparing the runtime of the refactored and original code**
--

#### **Refactored Runtime (2017):**
![outcomesVlaunch](https://github.com/mooshak21/stock_analysis/blob/main/Resources/NVBA_Challenge_2017.png)

#### **Original Runtime(2017):**
![outcomesVlaunch](https://github.com/mooshak21/stock_analysis/blob/main/Resources/VBA_Challenge_2017.png)

#### **Refactored Runtime (2018):**
![outcomesVlaunch](https://github.com/mooshak21/stock_analysis/blob/main/Resources/NVBA_Challenge_2018.png)

#### **Original Runtime(2018):**
![outcomesVlaunch](https://github.com/mooshak21/stock_analysis/blob/main/Resources/VBA_Challenge_2018.png)

**Significance of these results:**
>We can clearly see a difference between the original code runtime and the refactored code runtime based on these images. For 2017, the approximate runtime for the refactored and original code were **0.059** sec and **0.477** respectively. This means the refactored code was approximately **8.1 times faster**. For 2018, the approximate runtime for the refactored and original code were **0.059** sec and **0.480** respectively. This means the refactored code was approximately **8.2 times faster**. This gives us great insight on the advantages of the new version of the vba code because if we were to translate this to a large scale of analyzing company in the market, it would be much more useful to use this!

--
### **Chart comparison:**
--

![outcomesVlaunch](https://github.com/mooshak21/stock_analysis/blob/main/Resources/Stock_Value_2017.png)
![outcomesVlaunch](https://github.com/mooshak21/stock_analysis/blob/main/Resources/Stock_Value_2018.png)



