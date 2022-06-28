# Refactoring VBA Code
## Project Overview

### Purpose
Refactoring VBA script, prepared to analyze all stocks data over last few years, to achieve more effecient code implementaton 
using improved logic that takes less CPU time and RAM space. 

## Result

- **2017 VS 2018**

In 2017, most of the stocks have performed well and gave positive returns except "TERP". Also, based on "Total Daily Volumes",number of stocks traded more in 2017. 
In 2017, "DQ" gave highest return of 199.4%.

![StockAnalysis_2017](https://user-images.githubusercontent.com/107717882/176259178-26978724-d468-48e2-80db-e02f8e5a7950.png)

In 2018, most of the stocks couldn't performed well and gave negative returns except "ENPH" and "RUN". They both gave more than 80% return and "RUN" made to top by 84%. Also, based on "Total Daily Volumes", number of stocks traded in  2018 are less compared to 2017. 

![StockAnalysis_2018](https://user-images.githubusercontent.com/107717882/176259255-93a77ac3-5ec2-4bbd-ac59-ea60e88fd3c7.png)

- **Original Script VS Refactored Script**

From bleow images, it is clear that refactored code takes less execution time than original code.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/107717882/176259867-9e1fe757-71d5-4248-b551-0004218f57e9.png)

In original code we are excuting nested loops, for each ticker in array we are lopping over all rows; for 12 tickers; appx 3000 rows. We are excuting same code 36000 times. 

![VBA_Challenge_2018](https://user-images.githubusercontent.com/107717882/176259987-ba2d8aef-55d8-40c7-b5c3-30c92baf85c6.png)

Looking at the refactored code, it is obvious that loop is executing only once for all rows in dataset. Here, **tickers** array and **Ticker** column
in years(2017/2018) dataset are sorted so refactred code works most efficiently. The ticker index incremented when we moved to next ticker in row with improved 
logic of code. Total volumes, starting prices and ending prices are stored in corsponding arrays using ticker index. 

## Summary
  - Advantages and Disadvanteges of Refactoring Code
  
The advantages of refactoring codes are easier readability, less complexity and easier maintaiblity. A new user can easily read the refactored code as there is no complex flow of code. Easier readable code is easier to fix bugs and errors. Adding new features, unpdate old features and maintain code is much more easy and reliable.

On the other side, refactoring is time consuming, specifically for very large codes. Refactoring code increases speed of excusion but sometimes it takes more 
memory. Sometimes refactoring works only on some specific dataset , but fails on larger dataset and this can be lead to new bugs or errors. 
  
  - Advantagesand Disadvantagesof Originaland Refactored VBA Scripts
  
Look at refactored All Stocks Analysis, excution time is less as there's only one loop, whereas in originalcode there are nested lopps which take more 
ececution time.

Considering refactored VBA script, we have more easy and specific names to variables that we dont have in original code. 

In original VBA code, all calculatiions performed during looping and stored in temporary variables and ouput directy to data sheet. If we want to get total volume of "RUN" stock for future use, we won't have it, it is calculated, oputputed but not stored.

In refactored code we ovrcome this, we defined arrays of toalvolumes, starting prices, ending prices for each tickers. This makes us to use values for future calculatoins.

Memeory space using by arrays are more than temporiry variables but it is negotiable as we are getting clean, efficient, easily readable code with less execution
time.

Our dataset is sorted on **"Ticker"** as well as out **"tickers"**array in code; so CPU calculating very fast returns. But either of these or both are not sorted or someone added one more ticker at the end of array/dataset, then it will be errorsum calculation or execution. We must sort these two before execution.
 
This is really small dataset compared to real word data, so refacroing is easy. For larger dataset, it will be time consuming. 



