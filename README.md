# Stock Analysis VBA Challenge

## Overview of Project

### Purpose
The purpose of this project was to create a macro that would calculate the total daily volume and percent return for each stock ticker. We can run the analysis for either 2017 or 2018. Any tickers with a negative return are indicated with red cells, while tickers with positive returns are indicated in green.

## Results

### Stock Performance

Generally speaking, our analysis showed that 2017 was a better year for stock performance than 2018. In 2017, only one stock had a negative return: TERP. On the other hand, in 2018 all of the stocks had negative returns, with the exception of ENPH and RUN.
Additionally, in 2017, four stocks (DQ, SEDG, ENPH, and FSLR) had returns greater than 100% and were performing very highly, but in 2018 these returns decreased. DQ, SEDG, and FSLR show negative returns in 2018, while returns for ENPH decreased from 129.5% to 81.9%.

### Script Execution Times
For both the 2017 and 2018 analyses, the script run time decreased with the refactored code because we created arrays to hold the output data, rather than having the output data reported through each run of the for loop.

#### 2017 Analysis
My original script took 1.105469 seconds to complete the 2017 analysis. The refactored script ran in 0.203125 seconds. 

![Time to execute refactored script for 2017](https://github.com/secicciari/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)


#### 2018 Analysis
My original script took 1.117188 seconds to complete the 2018 analysis. The refactored script ran in 0.1953125 seconds.

![Time to execute refactored script for 2018](https://github.com/secicciari/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

## Summary
One advantage of refactoring code is that it can make the code more efficient so that it runs more quickly and doesn't require as much memory. Refactoring can also result in cleaner code that is easier to maintain and understand. By refactoring the original VBA script, I created code that ran more quickly and that is easy to update if there are ever additional tickers that need to be included in the analysis. Establishing arrays and variables early in the script, it is easy to follow what is happening and make adjustments to adapt the analysis to include additional data.

However, refactoring code requires spending time and resources that may not end up being worthwhile. For example, since the VBA script is not analyzing a huge amount of data and won't be used on a regular basis, the additional time that it took to refactor the code may not have made a meaningful difference in the grand scheme of things. That time could have been spent writing scripts for additional analyses or refactoring more complex scripts.
