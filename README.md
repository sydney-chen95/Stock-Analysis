# VBA Stock Market Challenge 

## Overview of Project

### Purpose
The purpose of this project is to help our friend Steve to analyze the entire stock market over the last few years. Steve's parents were interested in a stock called DQ so this is the stock we will analyze first. Then we will analyze the remaining stocks in our data to see which one performed the best and should be considerable for Steve's parents. We wrote a code on Excel's VBA editor to gather information for our initial stock DQ and then slowly increased the workload to a dozen more stocks. Furthermore, we want to know if there is a better way to use the code that we already wrote to make the code run better and more efficiently. We will refactor the code to make it loop through each stock faster with less memory and fewer steps.

## Results

### Analysis
After writing the code and running it to get all the information needed, I was able to create an output worksheet with each stock's ticker name, it's total daily volume and return rate with the respective year.


In 2017, the majority of the stocks in our analysis had a positive rate of return except for one which is TERP. TERP ended up with a -7.2% rate of return. Frankly, the stock that was in favor of Steve's parents, DQ, had the highest rate of return amongst the stocks. DQ had a 199.4% return rate that year. Other stocks that did well are SEDG with a return of 184.5%, ENPH with a return of 129.5% and FSLR with a return of 101.3%. The code had a timer setup to track how long it took to run through all the data. The results came in 0.75XX seconds for the year 2017. [INSERT PNG FOR INITIAL 2017 RUN TIME]


In 2018, the majority of the stocks in our analysis had a negative rate of return. The initially favored stock, DQ, is no longer doing as well as we had hoped. It took a dip down to -62.6% rate of return. The only stocks with a positive rate of return are ENPH and RUN with a rate of return of 81.9% and 84.0% respectively. The code ran in 0.765625 seconds for the year 2018. [INSERT PNG FOR INITIAL 2018 RUN TIME]


I was able to gather information on all 12 stocks in this dataset but to expand the worksheet to cover the entire stock market over the last few years, I will have to refactor the code to loop through all the data to collect the same information. In this challenge I dim a variable named tickerIndex as an integer type of data to create nested loops for looping through rows of data and collecting the appropriate values needed. Additionally, set up arrays for each ticker name, its volume, and starting/ending prices to calculate the rate of return. The nested loop will loop through each of the items mentioned for every ticker name. With this refactored code, the run time reduced to 0.1503906 seconds for both 2017 and 2018 as shown below.
[INSERT BOTH REFACTORED TIME PNG]


## Summary

- What are the advantages or disadvantages of refactoring code?
The advantages of refactoring a code is most definitely its maintainability and quality. The point of refactoring a code is to customize it to better fit our needs. It will be a lot easier to understand where the code came from. It will also run a lot faster to improve the quality of the code. Some disadvantages can come from more complex codes. It is time consuming to refactor a code that is harder to understand which can lead to more testing to avoid any bugs.   

- How do these pros and cons apply to refactoring the original VBA script?
Some of these pros and cons apply to refactoring the original VBA script. Since the original script wasn't as complex, refactoring it won't be too difficult as long as we identify new variables and organize the code as we progress. Inserting comments helped with the organization process of refactoring the code. The refactored code has 4 arrays so using different letters to represent the iterators helped with any confusion while we created the nested loops. In the end, everything was worth the time and effort since our code ran a lot faster and more efficiently. 
