# Stock Analysis

## Overview of the Project
This project is for Steve, who is analyzing 12 green energy stocks' performance to help his parents diversify their investment portfolio. Stock data is from 2017 and 2018 and includes tickers, dates, prices of stocks throughout the day, and volume. The code was originally written for the analysis, and a second code was refractored to see if it could run more efficiently. Code in the green_stocks.xlsm file is the original code, and code in the VBA_Challenge.xlsm code is refractored. 

## Results
### 2017 vs 2018 Stock performance 
Stocks in 2017 performed better than stocks in 2018. 
- In 2017, all but one stock (Ticker TERP) had a positive return.
- In 2018, only two stocks (Tickers ENPH and RUN) had positive returns. 
- Total daily volume was higher in 2018 ($3,306,038,200.00) than in 2017 ($3,166,639,100.00). 

### 2017 vs 2018 and original vs refractored execution times
- Codes for 2017 ran faster than 2018:
  - 0.273 vs. 0.285 in original
  - 0.289 vs. 0.293 in refractored
- Original code ran faster than the refractored code:
  - 0.273 vs. 0.289 in 2017
  - 0.285 vs. 0.293 in 2018
 
![2017 original](https://github.com/emariecovey/stock-analysis/blob/main/Stock_analysis_2017.png)
![2018 original](https://github.com/emariecovey/stock-analysis/blob/main/Stock_analysis_2018.png)
![2017 refractored](https://github.com/emariecovey/stock-analysis/blob/main/VBA_Challenge_2017.png)
![2018 refractored](https://github.com/emariecovey/stock-analysis/blob/main/VBA_Challenge_2018.png)

## Summary
### Advantages and Disadvantages of Refractoring Code
- Advantages:
  - Code could be made to run quicker/more efficiently
  - Code could become more "future proof" so that if you added/changed data, the code would still work
  - In this example, the refractored VBA code had the advantage of making the volume, start price, and end price into arrays. This helps the code appear  more clearly to me personally, because all of the statements creating the arrays are grouped together. It makes sense to me too when I'm reading the code, to remember that the array is going through each of the "i" elements in the for loop when the code is written arrayname(i). Also, the code is not overwriting the variables again and again, and each of the elements in the array are constant once they've been written. 
- Disadvantages:
  - Code could become "broken" if too many changes are made and errors not corrected
  - Code may not actually run faster or be more efficient, depending on the original code and how it is refractored
  - In this example, the refractored VBA code did not actually run more quickly. Also, If we add to the number of stocks in the array, we would need to add to the code where we declared the array 
