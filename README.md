# An Analysis of Stock Data (2017 - 2018)
Performing analysis on stock data for investment research

# Working with VBA 

## Overview of Project

### Purpose
The purpose of this analysis is to work with VBA in order to analyze financial data. In this project, we are working with stock data. The purpose of building our own scripts using VBA is to understand the core concepts involved in developing code, which include for loops, nested loops, conditionals, and arrays. This challenge provided an opportunity to learn how to refactor a piece of code that was already developed and edit it so it will be able to serve a similar purpose to what it was created for, without having to re-write the whole script. This challenge also allowed the exploration of measuring the execution performance of a code, which will display as a message box after the script runs.

## Results

## Stock Comparison
In this project, we are comparing the stocks with the tickers: AY, CSIQ, DQ, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, VSLR in the years 2017 and 2018. Using the macros created, we can quickly make this analysis with a click of a button. Using the Analysis button created, we can see the Ticker, Total Daily Volume, and Return displayed in three columns of the worksheet.

### Performance Results
![2017 image](https://github.com/julianneitliong/stock-analysis/blob/43547980d89f1ff58c353014cb3498b4c142d178/all%20stocks%202017.jpg)
![2018 image](https://github.com/julianneitliong/stock-analysis/blob/43547980d89f1ff58c353014cb3498b4c142d178/all%20stocks%202018.jpg)

When comparing the stock performance of 2017 and 2018, we can see that 2017 resulted in favorable returns. We can easily make this analysis using the conditional formatting of the colors in the Return column. The stocks in 2018 contained a lot of negative returns, thus resulting in a majority of red-colored cells. Though it is very difficult to draw conclusions about the list of stocks' overall performance when we are only looking at a 2-year history, we can say that 2017 was a better year for the overall stock market (or this portfolio) because a majority of the stocks did well. We can also say that the stock ENPH performed well in both 2017 and 2018 and the stock TERP did not do well in both years. 

### Execution Results
Execution results for the refractored code 2017 and 2018

![new 2017](https://github.com/julianneitliong/stock-analysis/blob/43547980d89f1ff58c353014cb3498b4c142d178/performance%202017.jpg)
![new 2018](https://github.com/julianneitliong/stock-analysis/blob/43547980d89f1ff58c353014cb3498b4c142d178/performance%202018.jpg)

Execution results for the original code 2017 and 2018

![original 2017](https://github.com/julianneitliong/stock-analysis/blob/43547980d89f1ff58c353014cb3498b4c142d178/original_performance%202017.jpg)
![original 2018](https://github.com/julianneitliong/stock-analysis/blob/43547980d89f1ff58c353014cb3498b4c142d178/original_performance%202018.jpg)

The refactored code ran a little longer than the original. This is due to the refactored script containing formatting code and having to execute those lines of code as well. In addition to iterating over the Ticker symbols, their values, and run calculations, the script also needed to check the results and apply the condition of red or green, change the cells to percentages, and reformat the Total Daily Volume. 

## Summary Statement
### 1) What are the advantages or disadvantages of refactoring code?
I found that the advantages of refactoring code is that it saves time by not having to re-write the entire code again. Instead, you can take the original code and replace or add lines and values where needed. One disadvantage I've encountered in refactoring this code is that it is harder to trace back where your error is if the code does not run properly. 

### 2) How do these pros and cons apply to refactoring the original VBA script?
When I refactored the original VBA script for this analysis, I found it helpful to know that the original script runs fine and that all I needed to do was change a few lines for this challenge. If I made a mistake, I can revert back to the original code and start again. My biggest challenge was that I kept coming across errors in the new script and it was very difficult for me to trace back exactly where my errors start in the code. This meant I did not fully grasp the concepts of the original script because I couldn't pin point exactly where the errors were. Overall, this was a good learning experience and very applicable to real-world situations.
