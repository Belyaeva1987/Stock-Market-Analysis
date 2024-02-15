# Stock-Market-Analysis #
## General info ##
This data has information about different stocks. The information includes different parameters, such as: Ticker, Date, Opening Price, Closing Price, Volume, etc. 

## Goals ##
The goal is to analyze the data and extract the following for each year: 
*	The ticker symbol.
*	Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
*	The percentage changes from the opening price at the beginning of a given year to the closing price at the end of that year.
*	The total stock volume of the stock. 

## Solution steps ##
To have the required information extracted, VBA script was used to process all stocks by identifying each of the unique tickers listed. After the unique ticker was identified, the total volume of stocks, yearly change of the price, and the percentage change for each unique ticker was calculated and put in the summary table. 
As a final step of the analysis, the information which Stocks had the greatest % increase, the greatest % decrease and the greatest total volume was identified and placed in a separate table. 
