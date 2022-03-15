# Stock analysis for all stocks 2017 and 2018

## Overview of the project
Steve uses this worksheet to find out which stocks have performed better over the years, so he can help investors like his parents make informed decision and lower their risk of losing money in the stock market.
His original worksheet worked well with limited number of stocks, but for larger set of data, it was too slow and taking a lot of his computer’s processing resources, so he needed us to keep the worksheets functionalities intact but do something, so the file runs smoother and faster.

## Results

### Stock performance
First let’s look at the stocks’ performance in 2017 and 2018.
At the first glance, we can see that the returns on the stocks, was generally much better in 2017 than 2018. Using If Then Statements in the code, we used formatting to get this information at a glance.

![2017-2018_comparison](/resources/2017-2018_comparison.png "Comparison of 2017 and 2018 stock performance")

### VBA Code performance
Stock performance aside, we wanted to refactor the code to improve the VBA code performance.	
When refactoring, we will be looking at the code to see how we can improve the script, so it runs faster and more optimal.

#### Original code
By looking at the script, we see that the original code uses nested loops, which loops 12 times (number of tickers) by 3013 times (number of rows in the dataset)

![Original-VBA-NestedLoop](/resources/Original-VBA-NestedLoop.png "Original VBA Code with Nested Loop")

 and you can see the elapsed time for script to run for each year, using the original code.

![2017_Runtime-Original_Code](/resources/2017_Runtime-Original_Code.png "2017 Runtime - Original Code") ![2018_Runtime-Original_Code](/resources/2018_Runtime-Original_Code.png "2018 Runtime - Original Code")

#### Refactored code

We would like to change the code, so that by going to through each row of data set, we can collect all the data we need and not need to go through the rows again for each ticker. 

To do this, we have created three more arrays in the code, to keep the values of Starting Price of each stock, Ending Price of each Stock and Total Volume of each stock during the year, in those arrays by going through the rows of the dataset only once, and then use the stored values to calculate Total Volume and Returns of each stock.

![Refactored-VBA-Array & Loop](/resources/Refactored-VBA-Array&Loop.png "Refactored VBA -Array & Loop")

And again, we calculated the elapsed time to run the script for each year

![2017_Runtime-Refactored_Code](/resources/2017_Runtime-Refactored_Code.png "2017 Runtime - Refactored Code") ![2018_Runtime-Refactored_Code](/resources/2018_Runtime-Refactored_Code.png "2018 Runtime - Refactored Code")

We see that by refactoring the code the way we did, we greatly improved the runtimes of the script.


## Summary
*Advantages and disadvantages of Refactoring code:*

Refactoring the code is done to improve the code so that it runs faster, taking fewer steps and using less system resources.
And in most cases it is much faster and cheaper to refactor a code than to start re-writing a code from scratch.
But if the script is complicated, done by someone else with not enough comments and not very well readable, then it might lead to some inconsistencies that are harder to catch and might defeat the purpose of improving the code with less budget and time than re-writing.

*How do pros and cons apply to refactoring the original VBA?*

In our example, refactoring the code significantly reduced the runtime, in comparison to the original script. By using arrays, it eliminated the need for nested loops which meant taking much fewer steps and leading to shorter runtime. The original code was not too big, and was written by us, so we knew all the ins and outs of the code, it was readable, had enough comments, so, the refactoring turned out successful with minimal issues.


## Dataset & VBA Scripts:
To see the Excel data file, run macros and see the VBA scripts you can download the file by clicking [here](/VBA_Challenge.xlsm "download the XLSM file").
