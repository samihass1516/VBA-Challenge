# VBA-Challenge

### Introduction:

In this analysis, we are helping our friend Steve decide if he should invest in the stock market or not.  So the purpose of this project is to refactor a Microsoft Excel VBA code to get specific information from the years 2017 and 2018.  The data collected from these two years are included in two charts with stock information on 12 different stocks.  Each stock we are presented with contains the following:
  1.	Ticker value
  2.	Date the stock was issued
  3.	Opening and Closing Price
  4.	Highest and Lowest Price
  And our goal is the return on each of these stocks.

### Analysis


I had to Copy and paste the code from a different application onto Excel and then refactor the data given to me.  As seen in the example provided below.  So, I created a program flow that loops through all tickers.  Using the Nested Loops, I ran a FOR loop inside another FOR loop to find:
  1.	Total volume of current tickers
  2.	Total start price of current tickers
  3.	Total end price of current tickers
  4.	Output the data for the current tickers  


![VBA_Challenge_1 (2)](https://user-images.githubusercontent.com/114379268/196966823-8fb66043-4bd1-4098-b756-01f2758f3795.png)

### Results




![VBA_Challenge_2017](https://user-images.githubusercontent.com/114379268/196967390-a5148382-bd08-4f6f-bd4b-bb41d50f196b.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/114379268/196967406-8da400af-da0a-4997-a9ad-28b1b7d01d18.png)

As we can see from the charts provided, we can determine that in 2017, the stocks were better than in 2018.  In the Return column, the color (green) indicates the increased value, and the (red) shows the losses.  

### Summary

Refactoring was a great way to find the stock market results faster and more efficiently. In the 2017 worksheet, the code ran in .65 seconds, and .13 seconds in the 2018 worksheets.  However, putting the data together took more time and effort to get the results.  If the code is not correctly annotated, it could mess up the whole report and result in refactoring and going through the information.   
