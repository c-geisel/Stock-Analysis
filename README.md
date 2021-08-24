# Stock Analysis

## Overview of Project
### Purpose 
The purpose of the following project is to analyze stock data using VBA. We are conducting this research to help out a recent graduate, Steve, who has been tasked by his parents to see which stocks are the best to invest in based on two factors, their total volumes and percent return. Initially this analysis was created to comb through 12 stocks but now we are refactoring the code to loop through all data at once so that we are able to run through our data faster. This will allow the code to run more efficiently as well as allow for the possibility to run an analysis on hundreds or thousands of stocks if we so choose. 

## Results
Insert a text file of the code, explain how it worked

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
