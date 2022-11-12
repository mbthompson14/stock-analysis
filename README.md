# stock-analysis

## Overview

The purpose of this analysis was to decrease the execution time for a VBA script that analyzes stock prices. This will allow the script to be used with larger amounts of data. I will accomplish this by refactoring the script to loop through the data one time instead of once for every ticker name.

The Excel workbook [vba_challenge.xslm](/vba_challenge.xlsm) contains prices and volumes for the stocks of interest for both 2017 and 2018. The orginial script is "AllStocksAnalysis" and the refactored script is "AllStocksAnalysisRefactored".

## Results

The original script looped through the data for every ticker name--a total of 12 times:
```
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
        Worksheets(yearValue).Activate
        
        For j = rowStart To rowEnd
            
            'increase totalVolume
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                'set starting price
                startingPrice = Cells(j, 6).Value
            End If
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                'set ending price
                endingPrice = Cells(j, 6).Value
            End If
        
        Next j
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = ticker
        Cells(i + 4, 2).Value = totalVolume
        Cells(i + 4, 3).Value = endingPrice / startingPrice - 1
        
    Next i
```

The refactored script loops through the data only once and stores all of the calculated values in arrays:

```
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            '3d Increase the tickerIndex
            tickerIndex = tickerIndex + 1
        End If
    
    Next i
```

I was successful in reducing the execution time of the script. The original execution time for the script was 1.058594 seconds (on my old laptop):

![](/resources/vba_challenge_2018_original.png)

The execution time of the refactored script was 0.2421875 seconds:

![](/resources/vba_challenge_2018_refactored.png)


## Summary

### What are the advantages or disadvantages of refactoring code?

There are several reasons to refactor code. Rewriting code to make it more efficient can decrease both the execution time and the memory used. It can also make the code easier to read and interpret by peers, and potentially make debugging easier. However, refactoring can be time-intensive and might not be a good idea if there is an upcoming deadline. In some cases the time it would take to refactor might be higher than the time it would take to write new code from scratch.

### How do these pros and cons apply to refactoring the original VBA script?

In regards to this analysis, refactoring the original VBA script significantly decreased the time it took for the script to execute. The refactoring process was fairly straightforward and was not time-intensive.