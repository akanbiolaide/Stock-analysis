# Stock-analysis
# VBA Challenge

## Overview of Project
This project will allow us to analyze real stock market data in order to be able to pull data information such as return percentage, greatest total volume, and also percent change and total stock volume 

### Purpose

The purpose of this project is to be able to use VBA in analyzing real stock market data


## Results
 

 


 

 


Ticker Index set to Zero;
 

Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices

Dim tickers(12) As String
Dim startingPrice As Single
   Dim endingPrice As Single
   
The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays
 

The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices
'5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If

Comparing the all-stocks analysis for year 2017, and 2018, in 2017 the stocks  have better return percentage than 2018. The stocks had better performance.

### Summary
1.	What are the advantages or disadvantages of refactoring code?
One advantage if refactoring is that it helps in maintaining the code , and keeps the code clean and organize. Disadvantage is the uncertainty of the developer.
2.	How do these pros and cons apply to refactoring the original VBA script?
Clean and organize code is easier to follow  and read, which make its less complicated to following up.

![image](https://user-images.githubusercontent.com/79291308/110276138-83f07180-7fa0-11eb-9930-3b4825956dc9.png)
