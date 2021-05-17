# VBA-challenge

## Background

This assignment utilizes VBA scripting to analyze real stock market data.

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Used while developing scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run scripts on this data to generate the final report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)


* This script loops through all the S&P stock data for one year and outputs the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Conditional formatting will highlight positive change in green and negative change in red.

* The result should look as follows.

![moderate_solution](Images/moderate_solution.png)

The solution also returns the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:

![hard_solution](Images/hard_solution.png)
