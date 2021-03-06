# Stock Analysis 

## Overview of Project
<p>Steve has recently graduated with a degree in finance and really wants learn about more stocks in the market. Since his parents are invested in Daqo, he wanted to create an even more in-depth analyzation of the Daily Volume of different stocks in the market, but with more of a focus on optimization. We will also only be focusing on data from 2017-2018.</p>

### Purpose

Steve's main focus for our task is to modify what we have already created for him throughout the module, but to refactor the code to make it more efficient, as we will be looking at more data. Using different techniques we will see if we can make the runtime of the program faster by employing different coding concepts. 

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
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
        
            '3c) check if the current row is the last row with the selected ticker
            'If the next row????????s ticker doesn????????t match, increase the tickerIndex.
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

---
### **Comparing the runtime of the refactored and original code**
---

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

---
### **Performace chart comparison:**
---

#### **Performace chart (2017):**
![outcomesVlaunch](https://github.com/mooshak21/stock_analysis/blob/main/Resources/Stock_Value_2017.png)
#### **Performace chart (2018):**
![outcomesVlaunch](https://github.com/mooshak21/stock_analysis/blob/main/Resources/Stock_Value_2018.png)

**Analysis of the charts:**
>I will be looking closely at the DQ ticker as that is the stock that we want to focus on for Steve. In 2017, DQ had an amazing year with 199.4% return which is outstanding and is the highest of the group of stocks we were looking at. However, in 2018 that was not the case. DQ had a worse year with -62.6% return. One thing to keep in mind is that the daily total volume increase by about 3 times its value in 2017. This is something Steve can look at to see how popular the stock was and how drastically it grew in only 1 year.

---

## Summary
### Advantages/Disadvantages of Refactoring Code
#### Advantages
1. Refactoring code in any language can help decrease its complexity and in turn increase its efficiency. This allows users to create programs that can run faster on larger data sets. For example, when I was in college we could complete a function in O(n^2) time but to improve it you could complete an action within the same loop instead of using another loop making the function run in O(n) time. 
2. It can make the code more readable. Like my example above, there are many ways to change code for the better, by using less lines and allowing the reader to easily digest it. If someone were to write something in 100 lines but it could be written in 50, that would make it much easier to read.

#### Disadvantages
1. Simplicity of code can be a disadvantage. When someone is first starting to code, it is most likely that they will be writing code that runs slower with more lines; however, they are usually writing the code in an easier way rather than using methods they haven't learned yet or that are more difficult. Refactoring code is this sense can be more difficult to understand for more people. 
2. Refactoring code can be more time consuming. Because you usually want to refactor code to make it run faster, there might be more time involved in the planning process on how you want execute it. Researching better functions/methods/macros to use can take a longer amount of time than using an easier approach. 

### Advantages/Disadvantages of our refactored & original VBA script
#### Advantages
1. We were able to get faster runtime. As we concluded in the results and analysis, the refactored version was around 8 times faster than the original. This makes the process less intensive on our computer and easier to run. 
2. We had to write less code in the long run and the code was written in a more efficient way. Using loops rather than hard-coding things is always the better option. 

#### Disadvantages
1. Personally, it just took more time understanding the new procedure we were performing in the new code. It takes time and hands-on-experience to get better at using different funtions/methods/macros. 
2. Overall, since the dataset was small, making the code more efficient wasn't really necessary because the processes were too intensive anyways.
