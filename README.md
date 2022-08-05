# VBA of Wallstreet - Module 2 Challenge
## Overview of Project: 

Visual Basic for applications is a powerful tool. It can help make applications easier to use by automating certain processes and procedures, as well as creating visualizations to better present data and make it more intuitive. This module allows the student to begin developing VBA skills and understand key concepts. These include:

-	Writing code
-	Understanding syntax
-	Becoming familiar with types of variables and how to declare them
-	Understanding lists and arrays
-	Understanding subroutines, If then statements, loops, nested loops and conditionals
-	How to use comments and pseudocode to help structure code
-	How to repurpose existing code
-	Understanding static and conditional formatting
-	How to create and use the button feature
-	And, how to make code interactive

The project also helps students understand how script can be more efficient, and how run times are affected by code structure. Concepts such as using indices and avoiding hard coding are presented, with the goal of improving run times and better use of computing resources.

## Presentation of Results:

### Stock Performance 
In 2017,  stocks in this cohort had positive returns, with some of them generating gains of over 100% (more than doubling in value). In 2018, however, the market for these stocks took a turn for the worse with only two stocks posting positive returns. Three of 12 stocks lost half their value in 2018. Overall, 2017 was a bullish year for these green stocks, followed by a year of sharp declines.

While total trading volume of all of the stocks combined grew modestly from $3.16 to $3.31 Billion year-over year, within individual stocks, trading volume varied greatly. For example, DQ’s trading volume increased by 206% and ENPH increased by 174%, while SPWR, FSLR, CSIQ and AY all saw trading volumes fall by 30% or more. 

### Runtime Performance
Original Runtimes were .2695 seconds for the ‘2017 ‘ input and .2852 seconds for ‘2018’. Using the refactored code, runtimes improved to .0664 seconds for ‘2017’ and 0.625 seconds for ‘2018’.

![Image of 2017](https://github.com/vjtrom/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

### Reasons for the improvement:
Run time improves from more efficient and better-structured code. These include use of indices in arguments versus string data, avoiding hardcoding, leveraging the use of arrays, and effective use of loops and nested loops. Making better use of arithmetic operations within the If Then statements is more efficient than evaluating string variables each time the loop runs.  

## Advantages and disadvantages of refactoring code:
There are both advantages and disadvantages to refactoring script. 
Some of the ***advantages*** include:
-	Leverages code that might already be working, therefore reducing development times and improving design
-	Allows the coder to build off of previous code incrementally in order to better facilitate debugging
-	Allows developer to look for inefficiencies in the original code and build upon improving run times and efficiency
 
Conversely, some of the ***disadvantages*** to rafactoring code include:
-	If original code is not running well, subsequent code might inefficient as well
-	If original code is poorly documented, it can cause confusion and slow down development
-	Structure of original code is inherently inefficient and poorly designed, thereby causing ‘GIGO’ (Garbage In, Garbage Out)

## How these pros and cons apply to refactoring the original VBA script:
The original code was well written in some areas, including setting the timer, declaring the string variables, and managing the formatting. These parts of the original code were for the most part left in intact in the refactored script. Documentation in the original code  made it easy to understand, and the flow and design were generally good. 

On the other hand, passing sting variables into the If Then statements were less efficient than using the TickerIndex variable, so a restructuring of the code to use the index over the three arrays vastly improved run time and efficiency. In addition, the original code did not efficiently use computer resources by leveraging mathematical calculations. 

Overall, the script was improved such that runtimes were reduced by 75%. 

Here is the link to the excel file:
[VBA_Chalenge.xlsm](https://github.com/vjtrom/stock-analysis/blob/main/VBA_Challenge.xlsm)

