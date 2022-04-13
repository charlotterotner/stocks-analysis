# stocks-analysis

## Overview of Project

The purpose of this project was to refacator the VBA code that we had originally written for Steve so it can be potentially used at a larger scale to examine more than 12 stocks and still run quickly. Our original code that we wrote for Steve worked to output total daily volume and return so he could make a recommendation to his parents on stocks to inveset in. We want to ensure he can continue to use the VBA script we wrote in the future if his sample of stocks increases, or if anyone else wants to easily take a look at our code in the future and easily understand it.
        
## Results

### Refactor Overview

To refactor the code we started by creating a tickerindex and setting it equal to 0. Before this we intialized an array of tickers, activated the worksheet for the year we wanted run the macro on, and got the number of rows to loop over. 
```
    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerindex = 0
    
```

Then we created our three output arrays, defined the variables and started a loop to initialize the tickerVolumes to zero

```
    '1b) Create three output arrays
 
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12)  As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    tickerVolumes(i) = 0


    Next i
    
```

We then activated our year worksheet and and created a loop to loop overall rows in the spreadsheet
```
    ''2b) Loop over all the rows in the spreadsheet.
    
        Worksheets(yearValue).Activate
         For i = 2 To RowCount
```
Then we got into our if statements for the output amounts
We wrote an if statement to increase the volume for the current ticker
```
      If Cells(i, 1).Value = tickers(tickerindex) Then
        tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        End If
```
Then we wrote an if statement to find the starting price. This statement looked to see if the tickers(tickerindex) was the first row in the selected tickerindex
```
        '3b) Check if the current row is the first row with the selected tickerIndex.
        
               If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
               
        'If it is Then set the starting price

        tickerStartingPrices(tickerindex) = Cells(i, 6).Value
               
        End If
 ```
 
We then did a similar if statement to find the ending price by seeing if the tickers(tickerindex) was the last row for that tickerindex we set an ending price and we also increased the ticker by 1 to repeat the loop with the next tickerindex in the series.
```
        '3c) check if the current row is the last row with the selected ticker
        
          If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
          
            'If the current row is the last row Then set the ending price price
            
                 tickerEndingPrices(tickerindex) = Cells(i, 6).Value
                 
            ' 3d If the next row’s ticker doesn’t match, increase the tickerIndex.
            
                  tickerindex = tickerindex + 1
        
            End If
            
            Next i

```

We then looped through our orrayrs to output the Ticker, Total Daily Volume, and Return. 
code
```
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
```

After doing the refactors described above we re-ran our code. From the timer message box we could see that the code did indeed run faster. Below are the screenshots of the timer before the refactor: 

**2017 Timer**

 ![image description or alt text](https://raw.githubusercontent.com/charlotterotner/stocks-analysis/main/Resources/2017_first%20run.png)
 
**2018 timer**

 ![image description or alt text](https://raw.githubusercontent.com/charlotterotner/stocks-analysis/main/Resources/2018_first%20run.png)
 
And after the refactor:

**2017 timer**

![image description or alt text](https://raw.githubusercontent.com/charlotterotner/stocks-analysis/main/Resources/2017_post%20refactor.png)

**2018 timer**

![image description or alt text](https://raw.githubusercontent.com/charlotterotner/stocks-analysis/main/Resources/2018_post%20refactor.png)

And that the refactored code is accurate since it gave the same outputs for each year that we originally saw.

**2017 Outputs**

![image description or alt text](https://raw.githubusercontent.com/charlotterotner/stocks-analysis/main/Resources/2017_outputs.png)

**2018 Outputs**

![image description or alt text](https://raw.githubusercontent.com/charlotterotner/stocks-analysis/main/Resources/2018_outputs.png)

### Analysis 
We can see that most of the stocks had an overall increase in return in 2017. The one exception to this was ticker TERP - Steves parents should likely avoid that one. In 2018, we saw a majority of the tickers actually decrease in returns. I was surprised by this at first, but it is not too surprising after reading that 2018 was one of the worst years for stocks in 10 years. All but two stocks decreased, RUN and ENPH. After reviewing the two years of stocks I’d probably recommend ENPH as a good next choice for Steve’s parents to invest in. It was strong in 2017 and withstood a stock crash in 2018. 

## Summary

Advantages and disadvantages or refactoring code:

Some advantage of refactoring code is having it run faster for large data sets, use less memory, and present code that is easier to understand. A disadvantage is that it is time consuming and could result in errors if not done correctly

Advantages and disadvantages of the original and refactored VBA script:

When we originally wrote the VBA code for Steve we had the For loop going off the tickers variable which is defined as a string. Strings take up more memory, so I’m assuming by refactoring our code to use the tickerindex variable which is set as an integer, the code will run faster, which it did. Our refactored code will definitely be better if Steve decides he wants to analyze more tickers than just the 12 we analyzed. One advantage of the original script is that there was one less For loop just making the code easier to read for a beginner in VBA. A disadvantage is just the extra time it took to refactor, and some of the errors made along the way. For the amount of time saved in running code (less than a second) it may not have been worth the hours it took to complete the refactor. In the future it would be good to understand how the codes future intended use before spending time refactoring. 
