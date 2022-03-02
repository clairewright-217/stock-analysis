# stock-analysis

## Project Overview
In this project, we built a [data analysis tool](VBA_Challenge.xlsm) to quickly analyze the performance of several different stocks across multiple years. The first version of the tool allowed the user to see a summary report on 12 different stocks in either the year 2017 or 2018. The analysis showed the total daily volume and the year over year return for each stock by year selected. In the second version, I refactored the code so that the report ran appoximately 5x faster. This report was built in Excel using VBA.

## Results
The Excel file has stock data divided up by year into two worksheets, one for 2017 and one for 2018. First, I prompted the user to specify which year they wanted to run the full analysis on by using an `InputBox` in VBA.

```yearValue = InputBox("What year would you like to run the analysis on?")```

Arrays were added for several types of data, including for the specific stock tickers in the dataset, for ticker volumes, ticker starting prices, and ticker ending prices. There were a total of 12 tickers, so the size of each array created was 12. 

``` 
Dim tickerVolumes(12) As Long
Dim tickerStartingPrices(12) As Single
Dim tickerEndingPrices(12) As Single
```
Next I needed to set up a variable called `tickerIndex` so I could access the correct index within each of the arrays, and I set this initially to 0. 

The macro needed to calculate and print the following information in a separate worksheet dedicated to the analysis results:
- the name of each stock ticker, which was specified as an array in the macro already
- the total daily volume of each stock over the course of one year (as selected by the user)
- the return of each stock as a percentage, calculated as the ticker's ending price divided by its starting price over the course of the selected year

To start, I created a simple `For` loop to first set each element in the `tickerVolumes`  array to zero.

```
For i = 0 To 11
    tickerVolumes(i) = 0
 Next i
 ```
Then I created a more complex `For` loop that looped through all of the rows of the selected worksheet ("2017" or "2018"). First, based on the `tickerIndex` value, it added up the total ticker volumes for the selected ticker. Then, it set the ticker starting price and ticker ending price based on a conditional statement. Finally, it bumped up the `tickerIndex` value based on the same conditional statement. 

```
For j = 2 To RowCount
tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value

  If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
     tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
  End If
       
  If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
        
     tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
     tickerIndex = tickerIndex + 1
            
  End If
    
Next j
```
After this for loop, I created another to drop the results into the summary worksheet. The loop used the index values of the ticker array to match the results with the correct ticker.     

```For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
             
Next i
```
I also added formatting to the beginning and end of the macro.

Refactoring the code this way improved the total time it took for the macro to run. Here are the results for running the full stocks analysis on 2017 and 2018. 

<img width="259" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/14280739/156291448-c9d7246a-77df-4807-b182-3c28911168e3.png">

<img width="261" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/14280739/156291450-54d85ee2-362a-4ab8-bd73-d8206d70ace0.png">

## Summary

### Advantages and Disadvantages of Refactoring Code
The big **advantages** to refactoring code were around code performance, reflecting on the problem deeper, and learning new coding techniques and new ways to problem solve. Because of refactoring, I was able to drastically cut the time it took to perform the analysis, and so if thousands of more rows of stock data were added to this file, the refactored code would significantly reduce the time it takes to run this analysis. Refactoring also forced me to review the initial problem again, and that made me understand it better. I could see refactoring code as a great way to learn more about a project that I am just starting on. Finally, I had to wrap my head around thinking of the same problem in a new way. I really wanted to use a nested for loop with the `tickerIndex` variable, because it was easier for me to think about it that way, but that resulted in slower code.

A couple of **disadvantages** to refactoring code is that it takes more time, and it might result in a solution that is more elegant yet harder to understand upfront. On some projets, you might not have the luxury to take extra time to revisit a coding problem that basically works already; it might make sense to move on to the next problem or feature rather than spend more time on a solved problem, even if the solution is not perfect. Also, I could see an issue arise if someone knows the code inside out, and refactors it so that it feels "perfect", but it also is now so hard to understand that teammates or future coders have to spend a huge amouunt of time understanding the code. That is, the code is so elegant that it's hard to follow without reading through it many times. 

### Advantages and Disadvantages of the original and refactored VBA script
The advantage of the original VBA script was that it was much easier to read through once and understand just through the code alone what was going on. The code read just like how one would solve the problem using pseudocode the first time. The disadvantage of the original VBA script was that it took more computing time to finish the analysis, so it would not scale as well as the refactored code if the dataset grew.

