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
An Analysis was performed to find each ticker **Total Daily Volume** and **Returns**. If we sum up all of the daily volume of each ticker we will have a rough idea of how often it gets traded. to find the returns, the stock's price at the end of the year is divided by the price at the beginning of the year, and converted to show percentage growth or loss. This indicates how much return in investment on a given ticker, with positive (green) values indicating increased value and negative (red) indicating losses.

<img width="252" alt="stocks 2017" src="https://user-images.githubusercontent.com/107584361/175803515-74e10648-c600-4e16-b6ae-749edbeca77d.png">.  <img width="248" alt="Stocks 2018" src="https://user-images.githubusercontent.com/107584361/175803391-b3b5e1c7-d1c8-4858-a2c5-fe78c80069d8.png">.

The Green Energy stocks for the year 2017 has a high ratio of positive return except for one ticker. Analysis for the year 2018 shows a complete different picture. The majority of the stocks have negative returns. The drop was significant. The DQ stock had almost **200%** yearly return in 2017, but in 2018 the stock dropped and finished the year with **negative 63%**.
These results indicate a risky investment. The stock trend is not stable and might not be worth investing all the money in DQ stocks.
### Creating a Macro
Both scripts **“AllStockAnalysis”** and **“AllStockAnalysisRefactored”** have the same output. Codes run calculations for the given stocks 
[stocks Data](VBA Challenge.xlsm) and return data on a new worksheet **All_Stock_Analysis**. The idea of presenting two codes with the same output is to highlight the **importance of refactoring**.
#### Ticker (column A):
Array `Dim tickers(12) As String` holds 12 tickers. Variable `tickerIndex` access array indexes and returning values in the table.
#### Total Daily Volume (colum B):
To calculate The Total Volume for a particular ticker : `tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 9).Value` To make this code work we need to create an array `Dim tickerVolumes(12) As Long` that holds 12 elements, and use our new variable `tickerIndex` to access ticker index in order to store the right value for the right ticker. Before and after the ticker changes the equation sums up total daily volume.
#### Yearly Return (colum C):
The following code calculates Yearly Return
```
If Cells(i - 1, 2).Value <> tickers(tickerIndex) And Cells(i, 2).Value = tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 7).Value`
    End If
    
If Cells(i + 1, 2).Value <> tickers(tickerIndex) And Cells(i, 2).Value = tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 7).Value 
    End If
```
To make this code work we need to use conditionals or `if statements`. The variable `tickerIndex` helps us find the starting and ending point of an old/new ticker in the dataset. Arrays `Dim tickerStartingPrices(12) As Single` and `Dim tickerEndingPrices(12) As Single` store captured values. 

#### For loops
For loops are responsible for executing the code in a repetitive manner until the condition is met. Incrementing a variable by 1 `tickerIndex = tickerIndex + 1` is responsible to move to the next ticker. By initializing arrays `tickerVolumes(tickerIndex) = 0` we reset the total volume to zero, before entering the loop again.
#### Formatting
In order to make the final table organized and visually pleasing, the code also contain formatting syntax.
```
Range("A1").Font.Italic = True
Cells(1, 1).Font.Size = 14
Range("B4").NumberFormat = "#,##0"
```
All formating that is possible in Excel, we can execute in VBA as well. By selecting a cell Cells(1, 1) or a range Range("A3:C3") we define where we want to apply formatting. There are plenty of useful sites online where we can find clear formatting instructions. This link provides some of the formatting options [Formatting](https://www.excelhowto.com/macros/formatting-a-range-of-cells-in-excel-vba/).

**conditional formatting**.
```
 If Cells(i, 3) < 0 Then 'set a condition
    Cells(i, 3).Interior.Color = vbRed 'color cell red.
End If
```



