# Green Stocks Analysis

## Analysis of the Trading Volume and Yearly Returns of Green Technology Stocks for 2017 and 2018

### The goal of this analysis was to gather and process data on the yearly returns and trading volume of green technology stocks for the years 2017 and 2018. This is being done in order to provide a data driven recommendation on what green technology stocks to invest in. Additionally, the original code for gathering this data was refactored and the speed of the macro was compared to determine which method is faster.

## Method, Results, and Analysis

### Method

A macro was built using VBA to automate the processing of the stock data. The macro is designed to pull out the starting and ending prices of each stock, as well as sum the total trading volume of each stock. The macro then outputs the values into an automatically formated table in the sheet "All Stocks Analysis" with 3 columns reporting Stock Ticker, Return, and Total Volume, where return is simply the ending price divided by the starting price expressed as a percentage.

#### Code

The first section of the code declares the variables that will track how long it takes the code to run. It also creates and fills the header cells, and prompts the user to input what year they would like to analyze.

![Code1](/Resources/VBA_Challenge_Code1.png)

The second section of the code declares four arrays that will hold the data. Each array has 12 elements with the ticker array having 12 strings for the stock tickers. The code also creates a variable that counts the number of rows in the data, which will be used to loop the macro over the correct number of rows and maintain the macro's fidelity if any rows are added or removed in future analyses. Also, the tickerIndex variable is declared and set to 0. This tickerIndex will be used in the following loops to track which ticker the macro is gathering data from.

![Code2](/Resources/VBA_Challenge_Code2.png)

The third section of code contains the meat of the macro. First, a for loop is run 12 times which will reset the value of each ticker's tickerVolume to 0 preparing for the calculation that will sum up all the volume and store it in accordance with the correct ticker. Next another for loop is created that loops over all the rows (except the header row). This loop first adds the volume of the current row to the tickerVolumes variable, then checks to see if the row above is different from the current ticker. If it is in fact different then the value of the starting price is stored. Then it checks if the row below is different from the current ticker and if it is then the value of the ending price is stored. Also, if this condition is true it advances the tickerindex by 1, which will start storing the volume, starting price, and ending price data the loop is gathering in the next elements of their respective arrays.

![Code3](/Resources/VBA_Challenge_Code3.png)

Finally, the  fourth section of code activates the "All Stocks Analysis" sheet and populates the cells of the table with the gathered data.

![Code4](/Resources/VBA_Challenge_Code4.png)

Some additional code at the end further formats the table with borders, colors, and some conditional formatting where green returns are positive and red returns are negative allowing the user to quickly see which stocks are providing a net return on the year.

![Code5](/Resources/VBA_Challenge_Code5.png)

### Results

![2017 Data](/Resources/VBA_Challenge_2017_Data.png)

![2018 Data](/Resources/VBA_Challenge_2018_Data.png)

### Analysis

As can be seen in the above tables, most of the green technology stocks performed much better in 2017 than in 2018. However, considering the goal of this analysis was to make recommendations for future investment, only the stocks ENPH and RUN are recommended.

In addition to the stock analysis, the code was refactored with the introduction of the tickerIndex variable. I believe the objective of using a tickerIndex is that the code is using an integer data type to track and store each ticker's data rather than the string data type the old code was using. This difference (I think) would use less memory and thus speed up the macro's run time, which would be relevant when using it to process very large data sets.

The original code:

![Original Code](/Resources/Green_Stocks_Code1,png)

The refactored code using a **tickerIndex** integer variable rather than a **ticker** string variable to advance the loop:

![Code3](/Resources/VBA_Challenge_Code3.png)

The run times for the original code were:

![Original Code 2017](/Resources/Green_Stocks_2017.png)
![Original Code 2018](/Resources/Green_Stocks_2018.png)

And the run times for the refactored code were:

![Refactored Code 2017](/Resources/VBA-Challenge_2017.png)
![Refactored Code 2018](/Resources/VBA-Challenge_2018.png)

As can be seen, the speed of the macro was improved by approximately one-hundreth of a second for both years. While this may not seem like much, this analysis only ran over 12 stocks. This time difference would certainly add up if the analysis were run on hundreds or even thousdands of stocks!

## Summary

### Advantages or Disadvantages of Refactoring Code

Refactoring code can have many advantages. One of the main advantages was demonstrated in this analysis and is the potential to streamline the code and improve run-time. Another advantage could be an improvement in the readability. The more comprehensible a code is the easier it will be to change or maintain the code in the future, or even to adapt the code to new applications. 

A potential disadvantage of refactoring code is the introduction of new bugs. However, this disadvantage should be completely avoided by properly saving the original code or even simply properly creating branches in GitHub. 

Again, as was demonstrated in this analysis, the refactoring of the original green stocks VBA code not only demonstrated the improvement in run-time, but also an improvement in readability of the code. If a future analysis is needed involving a larger data set of stocks I believe it would be easier to return to the refactored code and begin adapting it.


