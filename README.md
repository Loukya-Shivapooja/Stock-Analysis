# Stock-Analysis
Stock analysis using VBA
# Overview of Project
## Background of the Project
Steve parents are passionate about green energy. They believed that when fossile fuels used up there will be more reliable on alternate energy production. They are interested in green energy including Hydro Energy, Geothermal Energy, Wind Energy & Bio Energy. Steve Parents haven't done much research and instead wanted to invest in DAQO new energy corporation company which makes silicon wafers for solar panels, its ticker is DQ.
Steve wanted to analyse handful of green energy stocks in addition to DQ stocks. He created an excel file with stock data which he wanted to Analyse.
## Purpose of the Project
In this project, An Excel file with Green Energy Stock Data for the year 2017 & 2018 are given. we need to perform an analysis to see the Stock performance for the year 2017 & 2018 which helps to understand trends in the market.
However, we will be using an extension to excel built to automate tasks.
**Visual Basic for Applications** in short, **VBA** is a programming language which interacts with excel. It can read and write cells in worksheets make calculations and uses complex project to perform analysis. Using code through automated analysis allows to reuse it to any stock analysis and finally to create an interface to allow the users to perform functions with a click of a button.
we can learn more about VBA from the begineers website. 
This site helps us to understand VBA in detail [excel VBA](https://www.homeandlearn.org/index.html).
# Results
An Analysis was performed to find each ticker **Total Daily Volume** and **Returns**. If we sum up all of the daily volume of each ticker we will have a rough idea of how often it gets traded. To find the returns, the stock's price at the end of the year is divided by the price at the beginning of the year, and converted to show percentage growth or loss. This indicates how much return in investment on a given ticker, with positive (green) values indicating increased value and negative (red) indicating losses.

<img width="252" alt="stocks 2017" src="https://user-images.githubusercontent.com/107584361/175803515-74e10648-c600-4e16-b6ae-749edbeca77d.png">.  <img width="248" alt="Stocks 2018" src="https://user-images.githubusercontent.com/107584361/175803391-b3b5e1c7-d1c8-4858-a2c5-fe78c80069d8.png">.

The Green Energy stocks for the year 2017 has a high ratio of positive return except for one ticker. Analysis for the year 2018 shows a complete different picture. The majority of the stocks have negative returns. The drop was significant. The DQ stock had almost **200%** yearly return in 2017, but in 2018 the stock dropped and finished the year with **negative 63%**.
These results indicate a risky investment. The stock trend is not stable and might not be worth investing all the money in DQ stocks.
### Code Comparision
Two Macros **“AllStockAnalysis”** and **“AllStockAnalysisRefactored”** have the same output. Codes has performed calculations for the given stocks and returned data on a new worksheet **All Stock Analysis**. The AllStockAnalysis uses nested loop where as AllStockAnalysisRefactored uses indexing. The two codes are performed to understand the **importance of refactoring**.
### Refactored Code 
#### Ticker (column A):
Array `Dim tickers(12) As String` holds 12 tickers. Variable `tickerIndex` access array indexes and returning values in the table.
#### Total Daily Volume (colum B):
To calculate The Total Volume for one particular ticker : `tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value` To make this code work we need to create an array `Dim tickerVolumes(12) As Long` that holds 12 elements, and use our new variable `tickerIndex` to access ticker index in order to store the right value for the right ticker. Before and after the ticker changes the equation sums up total daily volume.
#### Yearly Return (colum C):
The following code calculates Yearly Return
```
If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
    
If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value 
    End If
```
By using `if then statements` and applying conditions the variable `tickerIndex` helps us find the starting and ending point of an old/new ticker in the data. Arrays `Dim tickerStartingPrices(12) As Single` and `Dim tickerEndingPrices(12) As Single` store captured values. 

#### For loops
For loops are responsible for executing the code in a repetitive manner until the condition is met. Incrementing a variable by 1 `tickerIndex = tickerIndex + 1` is responsible to move to the next ticker. By initializing arrays `tickerVolumes(tickerIndex) = 0` we reset the total volume to zero, before entering the loop again.
#### Formatting
In order to make the final table organized and visually pleasing, the code also contain formatting syntax.
```
Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
```
All formating that is possible in Excel, we can execute in VBA as well. By selecting a cell Cells(1, 1) or a range Range("A3:C3") we define where we want to apply formatting. There are plenty of useful sites online where we can find clear formatting instructions. This link provides some of the formatting options [Formatting](https://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/).

**conditional formatting**.
```
 dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd  
        If Cells(i, 3) > 0 Then 
            Cells(i, 3).Interior.Color = vbGreen  
        Else
            Cells(i, 3).Interior.Color = vbRed 
        End If   
    Next i
```
#### Run Time
The time taken for execution of **AllStockAnalysis** and **AllStockAnalysisRefactored** are showen in the below pictures.

<img width="207" alt="ALLstocks 2017" src="https://user-images.githubusercontent.com/107584361/175827737-1edaed50-b87a-48b6-bb24-19d65dac947f.png">.    <img width="199" alt="Time 2017" src="https://user-images.githubusercontent.com/107584361/175827781-7cb746d6-8397-4a38-ad5a-de9e8c415041.png">

**AllStockAnalysis** ran for `0.88 sec` whereas **AllStockAnalysisRefactored** ran for `0.17 sec`. The **AllStocksAnalysisRefactored** has ran 5 times faster than **AllstockAnalysis**.
# Summary
Refactoring is a key part of the coding process. When refactoring code, we aren’t adding new functionality; we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. 
The goal with refactoring is typically to improve readability, reduce complexity (and thus increase speed), and streamline the code, making it easier to maintain or extend.
### Advantages and Disadvantages of refactoring Code
Every process has its own advantages and disadvantages 
#### Advantages
The advantages of refactoring code are to improve code:

1. **Efficiency** - code is taking fewer steps, therefore taking up less computer memory and taking-up less time to execute the code,
2. **Readability** - code is easier to understand.
3. **Functionality** - fixing any bugs that might have been overlooked in the original code.
#### Disadvantages
The disadvantages of refactoring code can be:

1. **Frustrating and time-consuming** - we might not be aware of the purpose of the code and its functionality. Especially when the code is not well commented and we could spend a lot of time figuring out what specific lines or blocks of code are supposed to do. That's why the good documentation and commenting the code is very important.
2. **Less efficient** - by refactoring the code, we could end up with a less efficient script.
### Pros and Cons of refactoring the original VBA Script
The Orginal Code was a simple and easy, step by step process of nested looping, an iterative process within which multiple additional iterative processes are contained. In the refactored code, code stays in same loop gathers all data and stored in an array. Both has its own pros and. cons, by refactoring the code has run more faster i.e., 5 times faster making more efficient and faster, less time taking. 
