# 02 - Stock Analysis by Alec Ngai

## Overview of Project

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Our client Steve and his parents wish to have a workbook in which they can analyze the entire stock market over the last few years with a click of a button. To achieve this we must refactor our code, this is key to analyze a large dataset with compilation time. We are given a dataset in which we named ***"VBA_Challenge.xlsm"*** this dataset contains of two sheets categorized by the year **2017** and **2018**. Within both sheets there consists of 8 columns: 

**Ticker**:  This ticker assigns this row of data to stock, with a letter combination to identify the stock.  \
**Date**: When the row of data was collected \
**Open**: Price at which the stock openned at in the market \
**High**: Highest price the stock sold for during that day the market was open \
**Low**: Lowest price the stock sold for during that day the market was open \
**Close**: Closing price of the stock when the stock market closed \
**Adj Close**: The **adjusted closing** price amends a stock's **closing** price to reflect that stock's value after accounting for any corporate actions. \
**Volume**: The amount of stocks exchanged during the market open 

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; We will create an VBA macro to solve our clients needs, intially we will have two different macros, one is the original, and the other is the refactored one. In addition, we will also include a clear function for the client if he wishes to remove the analysis form the sheet. We will include 3 buttons linked to each macro for the user to have an easier experience using out macro. This macro will allow the user to select the year in which the data exists, then will automatically preform and output the analysis formatted to the sheet **All Stocks Analysis**. 

## Results

![2017_Analysis](https://github.com/alecngai/02-Stock_Analysis/blob/main/Resources/VBA_Challenge_2017.png)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Here we can see the analysis of the **2017** stocks, we can see the best preforming stocks were **DQ and SEDG**, with the biggest return difference of **199.4% and 184.5%**, and the worst preforming stocks were **RUN and, TERP**, with the worst return of **5.5% and -7.2%**. If the client were to buy DQ and SEDG in the beginning of 2017 it would be the most optimal investment,  allowing our investor to more than double their intial investment.  However, this data is not very useful to the client in predicting wether or not he should buy the stock in the present, it is only useful to tell the client if he should have bought in the past.  The most traded stock was **FSLR (101.3%) for 684,181,400 and SPWR (184.5%) for 782,187,000**, and the least traded was **DQ(199.4%) for 35,796,200 and HASI(25.8%) for 80,949,300** we can see here by this data, that the  least traded stock infact has the highest profit return, altough the most traded has a low return it is still positive, and the second most traded stock **FSLR is positive 101.3% return**. With this infomation we can conclude, that total daily volume has no direct corelation to yearly return.  The sum of **Total Daily Volume is 3,166,639,100** and the **average return** of the market was **67.3%**, this meant **2017** was a very good year to invest in, as you had a great chance to make a return. 

![2018_Analysis](https://github.com/alecngai/02-Stock_Analysis/blob/main/Resources/VBA_Challenge_2018.png)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; As you can see **2018** was not a good year for the stock market, all the stocks are negative return except for two stocks **RUN (84.0%) for 502,757,100 and ENPH (81.9%) for 607,473,500**. Previously in 2017 **RUN** was one of the worst preforming stocks and **ENPH** was a strong preformer having a yearly return of **129.5%**.  The sum of **Total Daily Volume is 3,306,038,200** and the **average return** of the market was **-8.5%**, this meant **2017** was a very risky year to invest in, I would not recommend investing in 2018 rather to wait out the stock market to be more in a bull market. However, through the two years **ENPH** is doing very well and I believe would be a good investment in the long run, if **ENPH** is able to maintain such a growth even during a horrible year for the stock market, there is a probability that in 2019, it will also preform well. The whole year most stock are not good investment, they are all in a bear market, this would most likely due to Trumps trade war with China. During the period of 2018 becuase of the trade many goods and services were halted between the two nations, causing many production delays and complete work arounds for some companies which caused the stock market to have such a negative return of **-8.5%** for the year on average. 

## Summary

### Advantages and Disadvantages of refactoring code in general

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; In general, refactoring code allows the user to have a much more cleaner code, less variables and more efficient loops. This means the code will utilize less memory for storage and processing, and the compilation time will be fast. Due to the code being refactored it means it also more simple and easier to be upgraded by another user. The logic of the user is more straight forward and readable, unlike having multiple nested loops, it does not tell the reader clearly what the code is doing until the reader has read through it all, but with refactored code, you can more easily see the logic. 

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; The disadvantage is the time needed to refactor the code, if your client wants a solution, there might not be a need to refactor the code. In this case, it is not working on live data, and there are no time constraints therefore runtime is irrelevant to the user, it is purely just for the smoothness of the experience. It serves no analytical advantage other than reducing compilation time. This is useful for online services or lives code that is always being called by an API, the more efficient the last hang time you will have your users experience when using your live service.  Another disadvantage is the chance of making a mistake and altering your output, in this case, we have clear set data where we can compare the changes, however, this may not always be the case, therefore it is up to the programmer whether or not refactoring is worth the time investment. 

### Advantages and Disadvantages of the original and refactored VBA script

![Runtimes](https://github.com/alecngai/02-Stock_Analysis/blob/main/Resources/Runtimes.png)


| Year | Original (s) | Refactored (s) |	 Difference |
|--|--|--|--|
| 2017 |  0.8930664 |0.1479492 | 6.04 |
| 2018 | 0.8869629 | 0.144043 | 6.16 |

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; In the above table is the run times comparing the original and refactored code. Please note, my code was not following the original format provided, I made my own edits to fit my coding style which I felt was more efficient. However, the main difference between using a nested loop vs refactoring is present.  As you can see the refactored code is roughly 6 times faster than the original. This is an amazing advantage in terms of compilation time, as this code can be scaled to work with larger datasets, rather than the original which uses a nested loop, with bigger datasets compilation time can be an issue and impede workflow. Another advantage is the code is much more readable, less complex, and easier to maintain and add to future builds. 

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; In this case there is not much disadvantage to refactoring the code, both code depend on the data to be presorted as their conditionals are created for this scenario. The only disadvantage is the time needed to refactor the code, the original code ran within 1 second, and was still efficient for this data set, so technically there was really no need to refactor the code, it only makes sense if the data set was bigger. 

### Code Changes from original to refactored 

![Ticker](https://github.com/alecngai/02-Stock_Analysis/blob/main/Resources/Code_Ticker.png)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; My code is different from the original in multiple ways as I found it to be very inefficient, first thing is I created ticker as a Variant and added all the values as an array, this allowed me to assign it all at once, instead of using an assignment call for all 12 tickers. This already increased efference a tad by reducing the number of assignment calls from 12 to just one. 

![Error](https://github.com/alecngai/02-Stock_Analysis/blob/main/Resources/Code_Error.png)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  I added an error checker to make sure that the user does indeed enter the year 2017 or 2018 as that is what we had data for. This was just a simple error checker and would exit the macro if it was triggered. 

![Core](https://github.com/alecngai/02-Stock_Analysis/blob/main/Resources/Code_Core.png)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; In the original code, it called for three separate if statements, this is redundant and would run a check 3 times per loop, I converted it into simple if and else if loop which still had the same functionality but would short circuit if an earlier condition was met, as there was no need to check all three conditions each time.  Added an index to get rid of an entire nested loop, which is the main time saver for this code. 

![Output](https://github.com/alecngai/02-Stock_Analysis/blob/main/Resources/Code_Output.png)

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; For the output section, I combined it with 2a, as initially the first thing the original code asks us to do is to initialize the tickerVolume, tickerStartingPrice, and tickerEndingPrices to zero, however, when you first call them they are already zero, so you do not need to reinitialize at the very top, you can do this at the bottom, this saves the code from repeating another for loop which iterates from 0 to 11. I combined this with the output, and after the output for that ticker has been done, it will set it to zero. 