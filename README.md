# Stock-Analysis


# Kickstarting with Excel

## Overview of Project



### Purpose



## Results

### Stock Performance Between 2017 And 2018
The below table compares some stocks in terms of their total daily volumes and their yearly return in 2018. Total daily volumes column shows the total number of traded shares for each stock. The code that I wrote to calculate the total daily volumes is as follows:

    Dim tickerVolumes(12) As Long
    For i = 0 To 11
       tickerVolumes(i) = 0
     
     Next i
      
     tickerindex = 0
     For i = 2 To RowCount
     tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value

The return column shows the percentage increase or decrease in price from the begining of the year to the end of the year. In other words, how much your investment grow or shrunk by the end of the year. I wrote the following code to calculate the yearly return percentage:

    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then

               tickerStartingPrices(tickerindex) = Cells(i, 6).Value

          End If
    If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then

              tickerEndingPrices(tickerindex) = Cells(i, 6).Value

          End If     
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    
Based on the below table,just two stocks (ENPH and RUN) have the positive yearly return. It is clear that ENPH has the most total number of traded shares. Also, the return percentage of this stock is the highest one(81.9%). It means that if someone invest on ENPH, his investment will increase by 81.9% by the end of 2018. The second stock in terms of the total daily volumes is SPWR, while the return percentage of SPWR is negetive.


This table also indicates the total daily volume and the yearly return percentage of stocks in 2017. In terms of traded shares, SPWR has the highest number, but the percentage of the yearly return for this stock is about 23%. Although, DQ has the lowest number of traded shares, it has the highest percentage of increasement in price from the beginning of the 2017 to the end of the 2017.All the TERP.If someone wants to select one of the stocks for investmen, he/she should select one's that has the highest total number of traded shares and the highest yearly return. 


### Execution Times of the Original and Refactored Script

## Summary

1-What are the advantages or disadvantages of refactoring code?

2-How do these pros and cons apply to refactoring the original VBA script?
