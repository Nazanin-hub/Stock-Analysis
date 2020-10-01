# Stock-Analysis

## Overview of Project

This project is about Helping Steve and his family to analyze stockdata. Steve's family want to invest all their money into Daqo New Energy Corporation. Steve is concerned about diversifying their fund. He wants to analyze a handful of green energy stocks in addition to Daqo's stock. He created an Excel file containing the stockdata. So, I helped him to analyze data by refactoring the provided VBA code. The VBA code automates the analysis and refactoring makes the VBA script run faster.

### Purpose

Refactoring VBA Excel code to analyze stockdata faster

## Results

### Stock Performance Between 2017 And 2018

The below tables compares some stocks in terms of their total daily volumes and their yearly return in 2017 and 2018. The total daily volumes column shows the total number of traded shares for each stock. The VBA code that I wrote to calculate the total daily volumes is as follows:

     Dim tickerindex As Single
       tickerindex = 0
       
    Dim tickerVolumes(12) As Long
    For i = 0 To 11
       tickerVolumes(i) = 0
     Next i
      
     tickerindex = 0
     For i = 2 To RowCount
       tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
       
     For i = 0 To 11
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
 
The return columns in the below tabels show the percentage increase or decrease in price from the begining of the year to the end of the year. In other words, how much your investment grow or shrunk by the end of the year. I wrote the following VBA code to calculate the yearly return percentage:

     Dim tickerindex As Single
       tickerindex = 0
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    For i = 2 To RowCount
       If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then

               tickerStartingPrices(tickerindex) = Cells(i, 6).Value

       End If
      If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then

              tickerEndingPrices(tickerindex) = Cells(i, 6).Value

      End If  
      
    For i = 0 To 11
     Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
     
    
 ![](https://github.com/Nazanin-hub/Stock-Analysis/blob/master/All%20stocks%20table%20-2017.png)
 ![](https://github.com/Nazanin-hub/Stock-Analysis/blob/master/All%20stocks%20table%20-2018.png)
    
#### Stock Analysis in 2018:   
    
As the above right table shows, just two stocks (ENPH and RUN) have a positive yearly return. It is clear that ENPH has the most total number of traded shares. Also, the return percentage of this stock is the highest one (81.9%). It means that if someone invests in ENPH, his/her investment will increase by 81.9% by the end of 2018. The second stock in terms of the total daily volumes is SPWR, while the return percentage of SPWR is negative. The best option for investment is ENPH or RUN. Because both are among the highest stocks in terms of the total number of traded shares and the percentage of yearly return.  

#### Stock Analysis in 2017:

The above left table indicates the total daily volume and the yearly return percentage of stocks in 2017. In terms of traded shares, SPWR has the highest number, but the percentage of the yearly return for this stock is about 23%. Although, DQ has the lowest number of traded shares, it has the highest percentage of increase in price from the beginning of the 2017 to the end of the 2017. All the stocks have the positive yearly return except TERP stock. If someone wants to select one of the stocks for investment, he/she should select the one that has the highest total number of traded shares and the yearly return. So, FSLR and SEDQ could be the best options. 

#### Comparison of Stocks in 2017 and 2018:

Some noticeable differences between these two tables:

   - Most of the stocks in 2017 have a positive yearly return compared to 2018. 
   - The highest percentage of the yearly return in 2018 is 84%, while, in 2017 is about 199.4%. 
   - The highest total number of traded shares in 2017 is about 782,187,000, while, the highest one in 2018 belongs to ENPH stock (607,473,500). 
  
Based on the data of these two tables, ENPH stock, on average, could be one of the most profitable stocks.

### Execution Times of the Original and Refactored Script:

Based on the below images: 
 - The execution times of the refactored script is less than the original script. 
 - The elapsed run time for the original code of table 2017 was 1.13 seconds.
 - The elapsed run time for the original code of table 2018 was 1.13 seconds. 
 - The elapsed run time for the refactored code of table 2017 is just 0.21 seconds. 
 - The elapsed run time for the refactored code of table 2018 is just 0.18 seconds. 

![](https://github.com/Nazanin-hub/Stock-Analysis/blob/master/VBA_Challenge_2017%20.png)
![](https://github.com/Nazanin-hub/Stock-Analysis/blob/master/VBA_Challenge_2018.png)

## Summary

1-What are the advantages or disadvantages of refactoring code?

- Advantages:

    1. Refactoring helps programming faster and improves performance.
    2. Refactoring helps finding bugs.
    3. Refactoring makes software easier to understand.
    4. Refactoring makes the code clean and organized.
    5. Refactoring removes redundant, unused code and comments.
   
- Disdvantages:

    1. It's risky when the application is big.
    2. It's risky when the existing code doesn't have proper test cases. 
    3. It's risky when developers do not understand what's all about.
    4. It may introduce bugs.
    
2-How do these pros and cons apply to refactoring the original VBA script?

   - Advantages:

       1. Refactored code performs the task much faster than the original VBA.
       2. Refactored code helps me to understand the original code better.
       3. Refactored code helps me to split out long functions into more manageable bite.So, it makes my code clean and organized.
      
   - Disadvantages:

       1. At first the understanding the new code was difficult, but after working with that, it was more easier than original code.
       2. refactored code introduce some new bugs, but I could debug it in a short time. 
    
