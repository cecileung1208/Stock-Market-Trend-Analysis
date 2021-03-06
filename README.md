# Stock Market Trend Analysis

![Image](https://www.gannett-cdn.com/-mm-/b2b05a4ab25f4fca0316459e1c7404c537a89702/c=0-0-1365-768/local/-/media/2018/12/20/USATODAY/usatsports/stock-market-price-ticker-display.jpg?width=660&height=372&fit=crop&format=pjpg&auto=webp)

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

* [Single Year Stock Data - Test](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Test/Single%20Year%20Stock%20Data.xlsm) 

* [Multiple Year Stock Data](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Final%20Results/Multiple_year_stock_data.xlsm)


## Method 
* Start with the Single Year Stock Data and do the following for all stocks:
  * Sort the data in chronological order for all stocks.
  * Determine the opening price for the first day of the year and the closing price for the last day of the year to calculate price change and % price change.
* Write a VBA script for Individual Stock Analysis that displays the following:
  * Stock Ticker Symbol
  * Yearly Price Change
  * Yearly % Price Change
  * Yearly total stock volume
  * Apply Conditional Formatting to the Yearly Price Change
* Upon successfully obtaining the results from the Individual Stock Analysis , write another VBA script for Multiple Stock Analysis that displays the following:
  * Stock Ticker with Maximum Increase in % Price Change and the % Amount
  * Stock Ticker with Maximum Decrease in % Price Change and the % Amount
  * Stock Ticker with Maximum Stock Volume and the number of shares
* Repeat process for the Multiple Year Stock Data.


## VBA Scripts

The following scripts are written for the given CSV Datasets:

* [Single Year Stock Data - Test](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Test/Single%20Year%20Stock%20Data.xlsm)
  * [Individual Stock Summary - Test](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Test/Individual%20Stock%20Summary%20-%20Test.bas)
  * [Multiple Stock Analysis - Test](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Test/Multiple%20Stock%20Summary%20-%20Test.bas)
 
* [Multiple Year Stock Data](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Final%20Results/Multiple_year_stock_data.xlsm)
  * [Individual Stock Summary - Final](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Final%20Results/Individual%20Stock%20Summary%20-%20Final.bas)
  * [Multiple Stock Analysis - Final](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Final%20Results/Multiple%20Stock%20Summary%20-%20Final.bas)

## Results
2014
 * Ticker DM has the Maximum % Price Increase at 5,581.60%
 * Ticker CBO has the Minimum % Price Decrease at 95.73%
 * Ticker BAC has the Highest Stock Volume at 21,595,474,700 shares

![Image](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Images/2014%20Results.png)

2015
 * Ticker ARR has the Maximum % Price Increase at 491.30%
 * Ticker KMI.W has the Minimum % Price Decrease at 98.59%
 * Ticker BAC has the Highest Stock Volume at 21,277,761,900 shares
 
![Image](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Images/2015%20Results.png)

2016
 * Ticker SD has the Maximum % Price Increase at 11,675.00%
 * Ticker DYN.W has the Minimum % Price Decrease at 91.49%
 * Ticker BAC has the Highest Stock Volume at 27,428,529,600 shares

![Image](https://github.com/cecileung1208/Stock-Market-Trend-Analysis/blob/main/Images/2016%20Results.png)




