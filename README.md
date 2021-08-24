# Stock Analysis

## Overview of Project
### Purpose 
The purpose of the following project is to analyze stock data using VBA. We are conducting this research to help out a recent graduate, Steve, who has been tasked by his parents to see which stocks are the best to invest in based on two factors, their total volumes and percent return. Initially this analysis was created to comb through 12 stocks but now we are refactoring the code to loop through all data at once so that we are able to run through our data faster. This will allow the code to run more efficiently as well as allow for the possibility to run an analysis on hundreds or thousands of stocks if we so choose. 

## Results
Insert a text file of the code, explain how it worked

'''

'1a) Create a ticker Index
    'Ticker Index will be what we are using for volumes, starting, and ending prices in order to talk about which ticker we are on.
    'It will generalize it so that we could have more than 12 stocks if we wanted to, thousands if even.
    tickerIndex = 0

    '1b) Create three output arrays
    'This is so that we can scan through the data all at once as opposed to starting over each time.
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    Sheets(yearValue).Activate

     '2a) Create a for loop to initialize the tickerVolumes to zero.
     'This tells us that we are going to start at 0 and then increase by 1 each time for the index and that the ticker volumes will start at zero.
     For i = 0 To 11
     ticker = tickers(i)
     tickerVolumes(i) = 0

     Next i

    '2b) Loop over all the rows in the spreadsheet.
    'This is so that we know to run through all of the rows starting with the second and then going to the end of the spreadsheet.
     Sheets(yearValue).Activate
     For i = 2 To RowCount
    
     '3a) Increase volume for current ticker
     'We can get rid of the if statement that we had here earlier, for each volume of the ticker index type we add it together.
     tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
      
     '3b) Check if the current row is the first row with the selected tickerIndex.
     'Starting prices and ending prices if statements are the same as before but change variables and add in the tickerIndex for referencing which index we are on.
     If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
     End If

     '3c) Check if the current row is the last row with the selected ticker.
     'If the next row's ticker doesn't match, increase the tickerIndex.
     If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
     '3d) Increase the tickerIndex.
     'We do not need another if statement, this can just be tacked onto the end the if statement from above since they need the same requirements.
         tickerIndex = tickerIndex + 1
     End If
     
     Next i
    
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    'Create a separate loop for this, need to access our output sheet to output our data.
    For i = 0 To 11
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
    
'''

### Stock Performance Between 2017 and 2018
-use images and examples from the code to explain differences in the two years
  total volume refers to how often the stock is traded throughout the year- if its taded often the price will accurately reflect the price of the stock 
        with the excepttion of a couple of stocks, the total volume was much higher in 2018
  percent retun is the icrease of decrease in price, how much your investment grew or shrank.
        2017 had higher yearly returns, more growth- 2018 has was more shrinkage.

### Original Script vs. Refactored Script
-compare execution times of the two different scripts using the images
In the original script we were looping through each index separately and then testing for the next one, in our refactored script we are looping through all at once, This improves efficiency. In old script for each year I was getting around 0.8 now I'm getting around 0.2

## Summary
### What are the advantages or disadvantages of refactoring code?
-advantage- it runs faster, it can run for more stocks because it's more generalized, it makes it cleaner and more efficient 
-disadvantages, it can add in bugs that were not there before, you're taking a stable code and messing with it.
### How do these pros and cons apply to refactoring the original VBA script?
- refactoring it made it run quite a bit faster which means we could analyze thousands of things if we wanted to. It made it more efficient and the tickerIndex generalized many items. However it took quite a bit of time to refactor. In a company this may take time that that is not there to refactor working code. Also it was important to save a copy of the old code because since we were messing with it we didn't want to introduce bugs or problems that made it then unsuable.
