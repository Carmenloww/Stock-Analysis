# Stock-Analysis

## Overview of Project
### Purpose
The purpose of this challenge is to edit or refactor the VBA code to loop through all the data one time in order to collect the same information. Then we will determine whether refactoring the code was successfully made the VBA script run faster. A written analysis that explains the finding will be presented. Refactoring is a fundamental part of the coding process. You're making the code more efficient with fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Results

1. Created a tickerIndex variable to set it equal to zero before looping over the rows.
  ![Create tickerIndex.png](https://github.com/Carmenloww/Stock-analysis/blob/master/Resources/Create%20tickerIndex.png)
2. Created three output arrays for tickers(),tickerVolumes(), tickerStartingPrices() , and tickerEndingPrices().
![Create Arrays Output.png](https://github.com/Carmenloww/Stock-analysis/blob/master/Resources/Create%20Arrays%20Output.png)
3. Access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays using the tickerIndex
![Create Loop.png](https://github.com/Carmenloww/Stock-analysis/blob/master/Resources/Create%20Loop.png)
4. Script the loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
![The script loops.png](https://github.com/Carmenloww/Stock-analysis/blob/master/Resources/The%20script%20loops.png)
5. Code for formatting the cells in the spreadsheet is working.

6. There are comments to explain the purpose of the code

7. The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook will match the outputs from the AllStockAnalysis in the module.

8. The pop-up messages will show the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png
![VBA_Challenge_2017.png](https://github.com/Carmenloww/Stock-analysis/blob/master/Resources/Screen%20Shot%202020-10-03%20at%2012.01.42%20PM.png)
![VBA_Challenge_2018.png](https://github.com/Carmenloww/Stock-analysis/blob/master/Resources/Screen%20Shot%202020-10-03%20at%2012.01.26%20PM.png)

![VBA_Challenge_2017.png](https://github.com/Carmenloww/Stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

In 2017, all of the stocks had positive Returns except for TERP (-7.2%). "DQ" made the best yearly return with 199.4% but it has the lowest total Daily Volume (35,796,200) in 2017.

![VBA_Challenge_2018.png](https://github.com/Carmenloww/Stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)

In 2018, all stocks had a negative Return percentage except for  ENPH (81.9%) and RUN (84%). They both had positive yearly Returns with high Total Daily Volumes and outperformed than other green stocks.

## Summary

What are the advantages or disadvantages of refactoring code?

When you are code refactoring, you are optimizing the existing code without adding any functionality. The advantages of code refactoring are to improve the performance of the code more efficiently and to be easier to understand or read. The disadvantages of code refactoring are that it is time-consuming, and you are likely to spend much more time solving the problem. It is possible that the coding may go wrong due to the complexity of the code. 

How do these pros and cons apply to refactoring the original VBA script?


