Overview of Project: Explain the purpose of this analysis.
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?


# Automated Stocks-Analysis using VBA

## VBA Project Overview
This was a very great project for two reasons:
1. I was asked by my good friend Steve for some help. His parents have asked him to be their financial advisor now that he recently graduated with his finance degree. Their first request is for Steve to let them know the performance of a particular stock. This request quickly turned into Steve needing to create a tool that would enable him to analyze the performance of hundreds of stocks with the click of a button. This would enable him the ability to keep his parents up-to-date on their best choices when investing. 
3. It also provided me with a terrific opportunity to test my newly learned skills using Microscoft VBA tools. I added a little extra in order to help myself understand the implications of the different ways of creating this code. In order to measure the performance of my first script versus refractoring, I added run timers for comparing. 

---

## Analysis



This was my first time using VBA editor so it was a **LOT** of trial and error as I worked through the creation of the script. The first request was to analyze one specific stock ticker "DQ". This was the ticker that Steve's parents were most interested in knowing the performance. The first step of the macro I built was something as simple as creating the output headers. That required me to list the header names. I used the following to accomplish this:

 ![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/DQ%20Analysis%20Output%20Headers.png)

Next I had to analyze all stocks in order to pull the data for only "DAQO" This was my first use of a **``FOR``** LOOP that I used to loop though all the rows looking for the "DQ" ticker. Nested inside the **``FOR``** Loop I leveraged **``IF``** statements. Inside the ``**FOR``** LOOP I wrote the commands to analyze the yearly stock reports for the "DQ" ticker, "Starting Prices", and "Ending Prices". Next it summed up the daily volumes to increase the total volume. 

![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/DQ%20Analysis%20FOR%20LOOP.png)

The results from this analysis quickly shows that while Steve's parents are pretty adamant about purchasing "DQ" stock, it's annual performance dmonstrates it had a negative return of -.06%. 

![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/DQ%20Analysis%202018.png)



The next request from Steve was to offer the ability to run the same annual performace analysis on all 11 stock tickers that are present in his reports. This affords Steve to quickly comnpare all tickers in one view. Sorting by the return percentage provides Steve with a picture of each ticker's trend. 

To acomplish this I used the core script from above to create the output headers which were the same as for the "DQ" ticker anaylisis. Next, I had to create an array of all available tickers in the reports (**This was one of the more manual tasks in this anaylisis**).  Below is a list of the highlevel steps I took to create the anaylisis for all stocks. 

-- I used some of my beginning code form the DQ analysis

1. Created a "start time" request
2. Create an array with all 11 tickers listed
3. Create a ``FOR`` LOOP to review all rows
4. Create ``IF`` statement to sum all volumes by ticker
5. Create ``IF`` statement to get starting and ending prices by ticker
6. Create outputs for each ticker that displayed "Ticker", "Total Volume", and the "Return" (same as the DQ analysis)

Here are a few of the blocks of code I used to accomplish the above steps.

**Create the ticker array:**

![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/Ticker%20array.png)

**Create the ``FOR`` LOOP:**

![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/All%20Stocks%20Analysis%20FOR%20Loop.png)

The end result of this code compared to the first macro (DQ Analysis) shows the totals for all 11 tickers in the same column. This makes it easy to see the best performing stocks over the past year. In addition to creating this powerful analysis tool, I added some featurtes that alow the user to interact with the worksheet. These useful additions reduce the manual button clicks required by Steve which ultimately increases his productivity and allows him to accomplish more work during his day. 

Adding some simple code can make a big difference. As a data analylist I am expected to deliver a product that streamlines the workflow, is easy to use, and visually pleasing and insighful. Steps I took to acomplish this for Steve were:

1. Formatted the output with the use of **bold** column headers
2. Color code the annual return column based on value using the ``conditional formatting`` tool in Excel
3. Created bottons on the sheet that allow the user to "Run All Stocks Analysis" and "Clear Sheet"
4. Created a ``MsgBox`` that prompts the user to select the year they would like to run the analysis on after pressing the "Run All Stock Analysis" button 


![screen-gif](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/Demo.gif)

---

## Results

1. Based on reviewing the results from both 2017 and 2018, Steve can quickly see that his parents desired stock "DQ" performed considerably better in 2017 than 2018. Since the results are color coded Steve can focus his efforts of additional research on both "ENPH" and "RUN" instead of "DQ".

![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/Results_Comparison.png)

2. This was an awesome challenge as it was both fun and difficult at the same time. We learned that while there are several different ways we can write the syntax to successful build a macro, they are not all created equally. Some may be very basic and require more steps on the worksheet from the user while some can very complicated and complex. Neither are wrong but one may be much more effiecient. Refactoring our first script demonstrated how by organizing, reducing redundancy, and streamlining the requests, we can end up with a final product that is easy to read, easy to debug, and runs more effieciently. 

> *"The main purpose of code refactoring is pretty clear — to make the code clean, neat, more efficient and maintainable."*
-> https://lvivity.com/what-is-code-refactoring

Run Times **Before Refactoring**

![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/Run%20times.png)

Runt Times **After Refactoring**

![](https://github.com/TONY-H83/Stock-Analysis/blob/main/Resources/Refactored%20Run%20Times.png)

The reults are crystal clear that through the refactoring process the run times vastly improved:

The 2017 run time was reduced by **76%**

The 2018 run time was reduced by **77%**

##Summary

*What are the advantages or disadvantages of refactoring code?*

Refactoring code can improve effieciency, make it easier for other users to read and understand your code,and make code more universal for repurposing.

*How do these pros and cons apply to refactoring the original VBA script?*

