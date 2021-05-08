# Stock-analysis

## Overview of Project

This workbook contains the tables of stock value returns in the past years and the refactoring code, which can loop through any similar set of large stock dataset one time to successfully collect the necessary information faster with a click of a button.
### The purpose 
This project aims to help Steve's parents evaluate and compare the stock's performance between the 2017 and 2018 datasets at a glance, and they can also use it for future needs.

## Results
### Compare the stock performance between 2017 and 2018
The following is the function output produced by the program VBA. These tables provide the results of annual stock's volume and returns for 2017 and 2018.

### Returns in 2017
![ VBA_Challenge_2017.png ](https://github.com/summerginger/stock-analysis/blob/main/VBA_Challenge_2017.png)
The table above illustrates the yearly return in 2017, we can see that the highest yield is the DQ stock, valued at 199.4%, and the lowest yield is Run stock at 5.5%. And there was only one negative return throughout the year, which was Terp at -7.3%. Overall, the vast majority of stock returns in 2017 were good.

### Returns in 2018
Compare the same pile of stocks in 2018. In just over a year, the stock market has undergone tremendous changes. For the highest stocks in 2017, DQ stocks fell to the lowest -62.6% in 2018. The same surprising story happened to Run Stock, but its value soared to the highest point of return, reaching 84% on the positive side. Overall, the stock market performance in 2018 was very worrying.

![ VBA_Challenge_2018.png ](https://github.com/summerginger/stock-analysis/blob/main/VBA_Challenge_2018.png)

### Conclustion and Recommandation
To conclude, the stock market can take stairs going up and windows going down. These tables only provide the returns and volumes; it is difficult to tell where the institutions may be putting their money. Therefore,  we recommend that Steve also needs to consider other ways to protect and predict his parents' further investments.

### Compare the execution times of the original script and the refactored script
In the future, Steve's parents may want to perform their analysis on larger datasets, and this timer code can help them know how fast the VBA can compile the results. The refactored code is faster than the original script. 

![2017 original]( https://github.com/summerginger/stock-analysis/blob/main/2017%20original.png) 

![2018 original]( https://github.com/summerginger/stock-analysis/blob/main/2018%20original.png) 

## Summary of Advantages and Disadvantages of Refactoring Code in General
The advantages of refactoring code are that it helps find bugs, makes the program run faster, more concise, cleaner, more readable, and more understandable. It is not changing functionality. However, the disadvantages are that it can be time-consuming, testing slow down the development, and are prone to more defects.

## Summary of How do these Pros and Cons Apply to Refactoring the Original VBA Script?
The best thing about reapply the same code is that there is no repetition. It is a fast, free tool to automate any similar task and debug function help to find errors and easy to maintain. Moreover, the breakpoints for the loop is very helpful, and it helps visualize the actual value for that indicator. However, it has a lot of broad-reaching effects. It is difficult to tracking and change all places of variables, especially in a large dataset. And it is case sensitive and space sensitive. It also involves planning, trying, researching, and testing. It can end up spending a good chunk of time as well.
  
### *The buttons' instruction
We assigned macros code to the button to make it user-friendly. So here is a simple instruction for Steve's parents to follow, to run a particular year of stock analysis, click the button "Run all stock analysis," and then enter the specific year. The tables will display that year's information. To make it run faster or to clear the worksheet, click the "Reset" button.  

![buttons.png]( https://github.com/summerginger/stock-analysis/blob/main/buttons.png)




  




	
