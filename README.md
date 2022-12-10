Overview of Project: Explain the purpose of this analysis.
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?


# Automated Stocks-Analysis using VBA

## VBA Project Overview
This was a very great project for two reasons:
1. I was asked by my good friend Steve for some help. His parents have asked him to be their financial advisor now that he recently graduated with his finance degree. Their first request is for Steve to let them know the performance of a particular stock. This request quickly turned into Steve needing to create a tool that would enable him to analyze the performance of hundreds of stocks with the click of a button. This would enable him the ability to keep his parents up-to-date on their best choices when investing.
3. It also provided me with a terrific opportunity to test my newly learned skills using Microscoft VBA tools. 

### Analysis

This was my first time using VBA editor so it was a **LOT** of trial and error as I worked through the creation of the script. The first request was to analyze one specific stock ticker "DAQO". This was the ticker that Steve's parents were most interested in knowing the performance. The first step of the macro I built was something as simple as creating the output headers. That required me to list the header names. I used the following to accomplish this:

 ![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/DQ%20Analysis%20Output%20Headers.png)

Next I had to analyze all stocks in order to pull the data for only "DAQO" This was my first use of a **``FOR LOOP``** that I used to loop though all the rows looking for the "DAQO" ticker. Nested inside the **``FOR LOOP``** I leveraged **``IF``** statements. Inside the ``**FOR LOOP``** I wrote the commands to analyze the yearly stock reports for the "DQ" ticker, "Starting Prices", and "Ending Prices". Next it summed up the daily volumes to increase the total volume. 

![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/DQ%20Analysis%20FOR%20LOOP.png)


