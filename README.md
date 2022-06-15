# Stock Analysis With Excel VBA
Click here to view the Excel file: [VBA Challenge – Stock Analysis]https://github.com/caitlinbighem/stock-analysis/blob/main/VBA_Challenge.xlsm

## Overview of Project
### Purpose
The following data and analysis are to determine whether or not specific stocks are worth investing in from stock data in the years 2017 and 2018. These conclusions are drawn through a code editing process known as refactoring, making the original code more efficient and effective to make advised stock decisions.

## Results
### Analysis
The original code provided included the details for creating the table layout, ticker array and general formatting. Each step for the refractor edit was shown to provide an easy to follow structure. Please see below the code and directions in the order provided.

'1a) Create a ticker Index
    tickerIndex = 0
       
    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
    
       
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'If the next row’s ticker doesn’t match, increase the tickerIndex.
        For i = 0 To 11
            tickerVolume = 0
        Next i

       ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
            tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(i, 8).Value
              
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then

               tickerStartingPrices = Cells(j, 6).Value
               
          End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row's ticker doesn't match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerEndingPrices = Cells(i, 6).Value
            
        End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
                tickerIndex = tickerIndex + 1

        End If

    Next i
       
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
        Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerStartingPrices(i) / tickerEndingPrices(i) - 1
        
        Next i

## SUMMARY: Our Statement:

### Advantages and Disadvantages of refactoring code

You need to perform code refactoring in small steps. Make tiny changes in your program, each of the small changes makes your code slightly better and leaves the application in a working state.

**Advantages:**
> - Errors are more easily identified in a more formatted and structured code with loops as well as nested conditionals.  
> - A restructure of the original code, if done correctly, will improve function and efficiency. 
> - Removes redundancies and duplications also contributing to effectiveness of the code.
> - Updates are self contained and prevent changes from having an impact on other parts of the code.

**Disadvantages:**

> - Incorrect edits will present new errors. 
> - Since the functionality is not impacted or obvious, improved code is difficult for users to catch. 
> - Testing and results may be effected as a result of refactor. 


**2. How do these pros and cons apply to refactoring the original VBA script?**

> Refactoring is defined as a restructure of source code or a piece of software in order to improve it without altering or compromising its functionality. This process can also be described as cleaning up a kitchen as you cook. If you continue to clean your surroundings as you cook, there is no chaotic mess of dishes, pots and pans left at the end of the night. Refactoring allows us to keep our code neat, orderly and easy to maintain. 

![VBA_Challenge_2017]https://github.com/caitlinbighem/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG

![VBA_Challenge_2018]https://github.com/caitlinbighem/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG
