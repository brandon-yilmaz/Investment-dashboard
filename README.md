# Investment-dashboard

## QUICK SUMMARY

Hi! This investment dashboard app retrieves stock market prices from Yahoo Finance API. The main feature of the app is that it calculates an internal share price 
(like an Investmend fund). This share price reflects your overall performance taking into regard the investment outflows and inflows.
Powerful metrics calculated as output plots:

1) Cumulative return based on the internal share price.
2) Share price of your fund.
3) Portfolio volatility (10d rolling window).
4) Last picture of your performance (main metrics of the last day).
5) Table viewer in order to navigate within the output of the data wrangling process.

Additionally, SP500 is used as benchmarck for the Porfolio volatily and Cumulative return metrics.

## Data input (xlsx template file)
The input data is collected from two Excel spreadsheets (automatically created within the app). Once you initialize the app, you will be asked to select a XLSX file.
This XLSX file template is automatically created by pressing the 'Data template' button and stored in your desktop. The XLSX template file will contain both spreadsheets.
The Excel book contains a kind of 'transactions book' and the SYMBOLS related to each asset necessary for the API to work. Symbols can be found 
in Yahoo finance within each asset's site. You will need to fill out both (transactions book spreadsheet and symbols) spreadsheets.



## Libraries to install
Pandas
Matplotlib
Yfinance
Numpy
Pandastable
tkinter
xlsxwriter
