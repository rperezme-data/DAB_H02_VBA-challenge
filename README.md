# VBA-challenge

## VBA Script for Stock Analysis
<br>
This Visual Basic script analyzes stock market data to obtain the annual price change, annual percent change, and accumulated trading volume for every stock in the source data file (Excel Workbook).
<br><br>

The structure of the source data file must be as follows:
1. The information for each year has to be presented in a separate Excel Sheet.
2. The data must be sorted by stock, then by date: oldest to newest (as is common practice in financial time series).
3. Trading data should use the following columns as headers:
    | ticker | date | open | high | low | close | vol
    | --- | --- | --- | --- | --- | --- | ---
<br>

## Results of the Analysis
<br>
The results of the analysis will be summarized in a table with the following format:

| Ticker | Yearly Change | Percent Change | Total Stock Volume 
| --- | --- | --- | --- 

<br>
Where:

    Yearly price change = [ Opening price at the beginning of a given year ] / [ Closing price at the end of that year ]

    Percent change = [ Yearly price change ] / [ Opening price at the beginning of a given year ]

## Additional Report
<br>
An additional report will display the Ticker and Value for:

- The greatest percent increase
- The greatest percent decrease
- The greatest trading volume
