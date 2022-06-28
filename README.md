# stock-analysis.
Module 2 data analytics
## Overview of Project
### Purpose
The objective of this project is helping Steve to analyze some historic shares information order to advice his parents about their investment on green technologies companies. They wanted to invest all your savings on DAQO a company dedicated to make solar energy supplies but they don’t know so much about that business and they don’t have enough information to make good decisions. 
So, Steve needed to analyze data of 2017 and 2018 years in order to find the yearly return for each stock. The first goal was analyzing the stock market data for those years on 12 different stocks and determine whether or not Steve's parents should invest in the stock. The yearly return is the percentage difference in price from the beginning and the end of the year. 
To reach this goal, Steve decided to use VBA since it gave him an excellent tool to analyze the data an to build code that could be run on different scenarios like each year.
Besides that, this project aimed to refactor the Microsoft Excel VBA code to  improve the time efficiency of the original code defined in the Excel spreadsheet “Green_Stocks.xlsm” subroutine called ”yearValueAnalysis()”. The main improvement was loop through all the data one time instead of several times as the original code did
## Results
### The stock performance between 2017 and 2018:
It seems that investment in DQ is not advisable and convenient decision since, although its return on 2017 was very good (199.4%), it decreased dramatically on 2018 (-62.6%). Instead of that, “ENPH” and “RUN” are the best Tickers that are worth investing in, since in the two consecutive years 2017 and 2018 they were the only ones that did not have a downward trend in the final price, therefore the percentage of annual return for both years was positive.
In addition to the above, these tickers had an increase of more than 150% in the volume of shares so we can be deduced that the final price of these tickers tends to rise for the following year 2019.
### Refactoring the Code: 
 The runtime of the original code and Refactor for each year will be shown below.  
#### Original Code Run-Times 
 
 
#### Refactored Code Run-Times 

 
 

We can see from the screenshots taken at each year's run that the runtime was substantially reduced with the refactored code because this code included going through all the data once instead of using a loop for each ticker. For this runtime enhancement three output arrays were declared tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
The analysis with original code for each year took approximately two seconds to run, whereas the analysis with refactoring code took approximately 0.37 seconds to run. It means that refactored code was more efficient that the original one in more than 80%  
#### Original Code:
 
#### Refactoring code:
 
## Advantages or disadvantages of refactoring code
There are some reasons why a company might want to refactor existing code, the most common is to increase the performance of the process, so one of the affected metrics is execution time. Making the code more efficient implies that we will obtain a very important benefit and it is the reduction of the execution times of the process, use less memory and ease the reading and understanding of the programming code by other programmers.

In the other hand, the most important disadvantage of refactoring is the risk of error and waste of time if applications are very large and there is no adequate documentation so it could take a lot of time and eventually the result is not good enough. Refactoring needs expert programmers, at least as good as whom wrote the original code, so the final result is as good as expected.

## How do these pros and cons apply to refactoring the original VBA script?

It is evident that the refactoring of the original code made the code more efficient since it substantially improved the execution time and consequently it made that the analysis could be faster and more efficient as well.

There is nothing bad related with this code refactoring. May be an inexperienced programmer or a beginner would find this code a little difficult to understand so that programmer decides to build the code in an easier way.
