# Stock-Analysis

## Overview of Project

performing data analysis on several thousand crowdfunding projects to uncover any hidden trends. 
How many columns and rows are there? 14 COLUMNS AND 4115 ROWS 
What types of data are present?
Is the data readable, or does it need to be converted in some way?
some of the columns contain text while other columns contain numbers. And not all numbers are in the same format; there are monetary units present as well as dates.
view columns I and J. You may notice that the data in these columns look like dates, but they're not in a readable format.



## Results

#### The tickerIndex is set equal to zero before looping over the rows. 
![1](https://user-images.githubusercontent.com/38533045/126405518-d9eca1e7-f799-4201-980e-25147990ed4b.JPG)

#### Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
![2](https://user-images.githubusercontent.com/38533045/126405762-3a8aa223-5716-48ff-b184-a5fa94648fd0.JPG)

#### The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.



#### The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.


#### Code for formatting the cells in the spreadsheet is working.


#### There are comments to explain the purpose of the code.


#### The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.


#### The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png 


## Summary

### What are the advantages or disadvantages of refactoring code?
#### Advantages
Refactoring improves the Design of Existing Code - Martin Fowler.
In Traditional Engineering, design puts 15% of total effort but design puts 90% effort into Software Engineering.
Design helps to make the code good and maintainable.
Design makes code easily understandable because code is loosely coupled, restructured, and duplicate code is eliminated.
#### Disadvantages
It is expensive and risky in the view of management.
It may introduce bugs.
Delivery schedule is very tight.
Management doesn't care about maintainability and extension of code base.

### How do these pros and cons apply to refactoring the original VBA script?
#### When Code Should not be Refactored?
Code should not be refactored on below cases:
Delivery Deadline is near and new development is planned.
The cost of refactoring is higher than rewriting the code from scratch.
Don't refactor the code if you don't have the time to test the refactored code before release. It can introduce bugs. 
Don't do delayed refactoring because it contains a big mess and makes it very difficult to refactor. Do early refactoring.
Stable code should not be refactored.
#### What are the Goals of Code Refactoring?
Code Refactoring has the below goals:
Maintainability - The main goal of code refactoring is to make it easy to enhance and maintain in the future. It should not violate Open Close Principle.
Removing Bad Smell - Bad code smell motivates the refactoring of code. It prevents many future defects. Code Size is reduced. Confused coding is properly restructured.
