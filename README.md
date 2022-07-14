# VBA_Challenge
## Overview of Project
Steve, a young graduate who's parents want to invest all their funds in DAQO new energy co-operation with ticker DQ. However Steve wantsto diversify his parent's funds by analysing other green energy stocks in addition to DAQO stock.

#####
Steve has his stock data in an excel workbook and needs help to know how the various stocks performed in 2017 and 2018.
In order to help Steve do this analysis and also make it easy for him to analyse any stock for any year and while reducing accidents and errors. We created VBA macro to automate the analysis for all the stocks steve wants learn about with the click of a button

## Results
Our original code which was adding up the volumes of the tickers along the rows while looping through the various stock tickers. And also comparing the tickers making sure we are storing the starting price and ending price of stocks with the same tickers tickers. 

#### Refactored Code
We refactored our code making the tickers and resulting starting price and ending price as well as volume sum an array. The code we used to achieve for the volume sum for instance looked like this:

For i = 2 To RowCount

If Cells(i, 1).Value = tickers(tickerIndex) Then

tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

End If

Where tickerVolumes was adding down the volumes of all the various groups of tickers. And tickerIndex was the indicator of the index of the particular ticker being worked on.

#### Performance
The time run of our original code was for the year 2017 was 0.2851562 secs and that for 2018 was 0.2890625. The runtime for both years were very similar. 

However after the refactoring the runtime reduced to 0.078125 for both 2017 and 2018. Screenshots are shown in Figure 1 for 2017 and Figure 2 for 2018 displaying the runtimes: 

![Screenshot1](https://github.com/Elfreda2019/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017.png)

Figure 1, Screenshot for 2017 code runtime.

![Screenshot2](https://github.com/Elfreda2019/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.png)

Figure 2, Screenshot for 2018 code runtime
