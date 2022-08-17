# stock-analysis
Stock-analysis project for module 2
## Overview of Project
Analyzing the stock market over the last few years in order to find the yearly return
### Purpose
Analyzing 2017 and 2018 returns by finding the change over the year from the starting to the ending price.
## Code and explanation
The code below does the following:
1. Ask the user for the year he/she would like to run
2. Defines the tickers index
3. Finds the total daily volume of each ticker
4. Find the starting and ending price for each ticker by chacking if the row before or after is matching the current row.
5. Populate the result in a new worksheet and mark the return that increased in green and the ones that decreased in red.
![Code example.png](https://github.com/Yahel-Epel/stock-analysis/blob/main/Resources/Code_example.png)

## Results
- 2017 performance VS 2018: in the chart below we can see that 2017 performance was better from 2018 for most of the tickers. While ENPH and RUN were the only tickers that increase in 2018 the rest decreased. 
We recomend checking if there were other market changes that affect the stock market in 2018. 

![2017 All Stocks.png](https://github.com/Yahel-Epel/stock-analysis/blob/main/Resources/2017_All_Stocks.png)
![2018 All Stocks.png](https://github.com/Yahel-Epel/stock-analysis/blob/main/Resources/2018_All_Stocks.png)
- The change of the execution times of the original script and the refactored script- we can see that the execution time of the original script (2017 1.0469 seconds and 2018 1.0625 seconds) was longer than the refactored script below. In the refactored script the code was more efficient. 

![VBA_2017.png](https://github.com/Yahel-Epel/stock-analysis/blob/main/Resources/VBA_2017.png)
![VBA_2018.png](https://github.com/Yahel-Epel/stock-analysis/blob/main/Resources/VBA_2018.png)
## Summary
1. The advantages of refactoring code are: 
    - having fewer steps 
    - Using less memory
    - Improving the logic of the code to make it easier for future users to read. 
- The disadvantage is:
     - Sometimes it takes longer to refactor the script than start form the beginning.
     - If we need an answer for only one question it can be faster to change only the necessary lines. 
2. These pros and cons apply to refactoring the original VBA- 
    - The refactoring code runs faster as we can see in the execution times. 
    - We can have a perspective at each stock in one place and it can help us make a decision.  
    - There is no loop nest that takes longer to rung and create. 
