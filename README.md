# Stock-Analysis

## Overview of Project

In this project and analyisis, we’ll refactor the Stock Market Dataset with VBA solution code to loop through all the data one time. We’ll determine whether refactoring your code successfully made the VBA script run faster.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

## Results

#### The tickerIndex is set equal to zero before looping over the rows. 
![1](https://user-images.githubusercontent.com/38533045/126405518-d9eca1e7-f799-4201-980e-25147990ed4b.JPG)

#### Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
![2](https://user-images.githubusercontent.com/38533045/126405762-3a8aa223-5716-48ff-b184-a5fa94648fd0.JPG)

#### The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.
![3](https://user-images.githubusercontent.com/38533045/126406702-fd64dc2a-0c74-4888-8b42-71b4aa6f38aa.JPG)

#### The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.
![4](https://user-images.githubusercontent.com/38533045/126406302-2e7b3faf-7c18-4456-8f73-0390299021be.JPG)


#### The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module.
![VBA_Challenge_2017](https://user-images.githubusercontent.com/38533045/126406792-3d1d8444-a30c-4d77-acd3-0654c0c10936.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/38533045/126406853-4951a6e7-febc-4f01-86df-3ad5e47571e4.png)


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
