# Stock Market Trend Analysis

## Background 

The purpose of this project is to determine the annual changes for all the stocks listed in the US stock exchange during 2014-2016.  To accomplish this, a VBA script will be written from the given CSV datasets.

## Requirements

#### 1. Create a VBA script that will display the following information for each stock:

* Ticker Symbol
* Annual Price Change 
* Annual Percent Price Change 
* Annual Total Stock Volume

#### 2. Use conditional formatting that will highlight positive change in green and negative change in red.

#### 3. Upon retrieving these results, determine the ticker symbol with the following:

* Greatest Increase in Yearly Percent Price Change
* Greatest Decrease in Yearly Percent Price Change
* Greatest Yearly Total Stock Volume

## Datasets:
* [Multiple Year Stock Data](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Final%20Results/Multiple_year_stock_data.xlsx)

* [Single Year Stock Data (Test)](https://github.com/cecileung1208/VBA-challenge/blob/main/Test/Unit%202%20-%20VBA_Homework_Instructions_Resources_alphabetical_testing%20-%20Verifying.xlsm) 

## Method 
* Start with the Single Year Stock Data and do the following for all stocks:
  * Sort the data in chronological order for all stocks.
  * Determine the opening price for the first day of the year and the closing price for the last day of the year to calculate price change and % price change.
* Write a VBA script that displays the following: 
  * Stock Ticker Symbol
  * Yearly Price Change
  * Yearly % Price Change
  * Yearly total stock volume.
  * Apply Conditional Formatting to the Yearly Price Change
* Upon successfully displaying the above information, write another VBA script based on the results to determine the following:
  * Yearly annual increase in percent price change 
  * Yearly annual decrease in ppercent price change 
  * Yearly annual total volume.
* Repeat process for the Multiple Year Stock Data.


## Results

2014

Based on the results in the below image, the 2014 results show that:

 * Ticker DM has the Maximum % Price Increase at 5,581.60%
 * Ticker CBO has the Minimum % Price Decrease at 95.73%
 * Ticker BAC has the Highest Stock Volume at 21,595,474,700 shares

![Image](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Images/2014%20Results.png)

2015

Based on the results in the below image, the 2015 results show that:

 * Ticker ARR has the Maximum % Price Increase at 491.30%
 * Ticker KRI.W has the Minimum % Price Decrease at 98.59%
 * Ticker BAC has the Highest Stock Volume at 21,277,761,900 shares
 
![Image](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Images/2015%20Results.png)

2016

Based on the results in the below image, the 2016 results show that:

 * Ticker SD has the Maximum % Price Increase at 11,675.00%
 * Ticker DYN.W has the Minimum % Price Decrease at 91.49%
 * Ticker BAC has the Highest Stock Volume at 27,428,529,600 shares

![Image](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Images/2016%20Results.png)




