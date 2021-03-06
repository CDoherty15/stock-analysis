# stock-analysis
Module 2: VBA on Wallstreet
# Module 2 Challenge: Refactoring Steve’s stock data analysis
Refactoring the code will help the code run more efficiently. To ensure this, we will analyze and compare the original code and the refactored code.
## Overview
	Purpose of this data is to help Steve find the total daily volume and yearly return for these stocks. VBAs in excel help with creating macros and are very helpful when dealing with stock data in excel. Daily volume is the total number of shares traded during the day, measuring how actively a stock is traded. Yearly return is the percentage difference in price from the beginning of the year to the end.
	Steve has already figured out a way to compile and format the data, but his boss figures he can refactor the code to be somehow be cleaner and quicker. We have helped Steve with the first part and now it is time to refactor his code. Below will discuss the actual analysis of the results and then analyze the refactored code.
## Results
	This section is split into two sections: the first will analysis the two years of stocks and the second section will compare the difference of the original and refactored code.
### Table Analysis
#### All Stocks: 2017 & 2018 tables
##### 2017 Table
 
##### 2018 Table
 
#### Analysis
	In 2017, we see a lot of green in the return column. To a stock broker, all of these companies ended the year in green, meaning between the beginning and end of the year, the stocks had a positive increase in daily volume. The only stock in the red was TERP. In 2018, we see the opposite, almost all of the stocks were in the red. Showing that 2018 was not a good year in the stock market for the observed stocks. 
	We can see that tickers ENPH and RUN both stayed in the green through 2017 and 2018. ENPH’s return was lower in 2018 than 2017, but stayed significantly high in the green. RUN’s return did the opposite and rose significantly in 2018. When Steve is presenting these results, the actual return volume should only be slightly considered but the stocks being in the green should be extra emphasized. Eleven out of the twelve stocks in 2017 were green and only two of them stayed green in 2018. Steve should be telling his boss to watch and probably invest in these stocks since they stayed green not only for 2 years in a row, but stayed green in 2018 compared to its counterparts.
	Another observation that would be good to note is that TERP might have been in the red in 2017 and in 2018, but the return percentage slightly increased (-7.2% to -5.0%). Steve should mention to look out for the TERP ticker in-case it reaches the green at any point.
	A final observation to make would be to mention the DQ ticker. In 2017, DQ had the highest return at 199.4%, but in 2018, had the lowest return at -62.6%. However, we can also see that in 2017, DQ had the lowest value for total daily volume and the second lowest in 2018. This is probably due to the stock being valued very low in 2017 and this caused a high amount of trading activity that year. In the following year, brokers probably didn’t think this stock was not worth what it increased too, causing for a tremendous decrease in return percentage and still having a low daily volume compared to its counterparts. 

### Original vs Refactored Code
#### 2017 Times
##### Original 2017
 
##### Refactored 2017
 
#### 2018 Times
##### Original 2018
 
##### Refactored 2018
 

## Summary of Results
### Advantages or Disadvantages of refactoring code
	The point of refactoring our code was mentioned at the beginning of this document, refactoring code helps run code more efficiently, which is why refactoring can be an advantage. As we can see by the time screenshots, that is true. With the original code, it took over 1.5 seconds for both years, and almost 2 seconds for 2018. After refactoring, both years run under 0.5 seconds, and in fact, both ran just under 0.4 seconds. Without any calculations, we can see the refactored code runs at least 1.1 seconds faster than the original code. 2 seconds for running code is terribly slow, but longer a code takes can sometimes run into a time error and even if the code is right, no results will output. Here we only were looking at 12 tickers, which still included over 3,000 data entries in the two year data sheets, but 2 seconds isn’t anything to worry about. Imagine if the list of observed tickers was 15 or 20, those extra tickers will cause for more rows of data, causing the code to work through more lines of data causing it to work harder. 
	A disadvantage of refactoring code will be the time consumption on typing the code. Looking at the actual excel spreadsheet, going to the “Stocks (Original vs Refactor)” worksheet, there are two buttons: “All Stock Analysis” and “All Stock Analysis (Refactored)”. The first button prints the results with the original code, while the other button prints the refactored code. Running both codes, we get the same results for both codes, showing that refactoring might not be necessary since we just get the same results. 
### How do they apply to refactoring the original VBA script 
	Coding involves attention to detail, refactoring in this case involved creating an index variable in order for the code to classify and calculate the values for the ticker at a quicker rate. Mistyping the index in any of the for loops can cause the code to not work because of an error. This problem is not that of an issue because VBA will point out where the error is and stop running the code. But sometimes a mistype can lead to the code running, but the output is wrong because the code could have skipped a row, combine rows 1 & 2, or really any combination to where the wrong result is outputted. This is a bigger issue cause we don’t know where the error is because the code was written properly according to excel, and therefore we will have to inspect every single line to see where the mistake was. On top of this, refactoring causes the code in the VBA to be longer and in some cases can be clunkier and harder to sift through. Even with the helpful comments, comments just add more lines of code, causing the code to be longer and clunkier. So we could refactor the code with no comments, which could cause for more issues when trying to edit; or take the time commitment to type out the code.

## Conclusion
	Refactoring the code helped the code run faster but causes the VBA script to be longer and clunkier. The advantage here will be that when we add more tickers and years, refactoring will be worth it so excel or the computer doesn’t crash. However, for analyzing only 12 stocks one year at a time, refactoring may not be necessary. Both years with the original code took over 1.5 seconds but both took under 2 seconds. These are not tremendously long times and for this case, are perfectly fine. I would recommend to Steve if he was just to analyze one year at a time, the original code will be fine and if he wants to challenge himself with refactoring, then go for it. If he were to add more tickers and potentially look at multiple years at once, then refactoring would be recommended and worth the tedious editing multiple lines of code. 
![image](https://user-images.githubusercontent.com/79118630/110221353-d85cf980-7e99-11eb-9c73-2bc109ccd2ac.png)
