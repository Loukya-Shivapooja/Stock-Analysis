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

![Stocks_2017]("https://user-images.githubusercontent.com/107584361/175802193-a293cf21-4e57-4e5b-b342-1083bd7f2335.png").

The Green Energy stocks for the year 2017 has a high ratio of positive return except for one ticker. Analysis for the year 2018 shows a complete different picture. The majority of the stocks have negative returns. The drop was significant. The DQ stock had almost **200%** yearly return in 2017, but in 2018 the stock dropped and finished the year with **negative 63%**.
These results indicate a risky investment. The stock trend is not stable and might not be worth investing all the money in DQ stocks.
### Creating a Macro
Both scripts **“AllStockAnalysis”** and **“AllStockAnalysisRefactored”** have the same output. Codes run calculations for the given stocks and return data on a new worksheet **All_Stock_Analysis**. The idea of presenting two codes with the same output is to highlight the **importance of refactoring**.
