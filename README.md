# stock-analysis

## Project Overview
In this project, we built a [data analysis tool](VBA_Challenge.xlsm) to quickly analyze the performance of several different stocks across multiple years. The first version of the tool allowed the user to see a summary report on 12 different stocks in either the year 2017 or 2018. The analysis showed the total daily volume and the year over year return for each stock by year selected. In the second version, I refactored the code so that the report ran appoximately 5x faster. This report was built in Excel using VBA.

## Results
The Excel file has stock data divided up by year into two worksheets, one for 2017 and one for 2018. First, I prompted the user to specify which year they wanted to run the full analysis on by using an `InputBox` in VBA.

```yearValue = InputBox("What year would you like to run the analysis on?")```

Arrays were added for several types of data, including for the specific stock tickers in the dataset, for ticker volumes, ticker starting prices, and ticker ending prices. There were a total of 12 tickers, so the size of each array created was 12. 

``` Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single```
   
    
    
    
    
