# Stock-analysis
## 1. Overview of the Project

### Purpose of the Analysis: 

The purpose of this project is to analyse the performance of 12 selected stocks over a two year-period (2017 and 2018) in order to aid the investment decision of some  potential investors (Steve's parents). 

Using VBA to create a script, the individual performance of these 12 selected stocks was analysed using 2 important investment parameters;

1. __Total Daily Volumes of Traded Shares__: This parameter analyses the number of shares traded throughout the day for a particular stock. It's a measure of how actively a stock is traded. An actively traded stock may indicate a potentially good investment.

2. __Yearly Return of Each Stock__: This parameter analyses the percentage difference in price of a stock from the beginning of the year to the end of the year. A positive percentage difference is an indication of growth in stockholder's investment whilst a negative percentage difference is an indication of a decline in stockholder's investment.

## 2. Result
__Year-On-Year Comparative Performance__: The following images show a detailed view of the performance of each of the 12 selected stocks in 2017 & 2018.


__2017 Performance Result__

<img src= "Resources\VBA_Challenge_2017.png">



__2018 Performance Result__

<img src="Resources\VBA_Challenge_2018.png">


A comparative review of the performances of the stocks over the 2 analysed years show that 11 of the 12 stocks reviewed (representing about 92% of the total analysed portfolio of stocks) generated positive returns in 2017, with _DQ, SEDG, ENPH & FSLR_ generating tripple digit returns on investment during the year. In contrast, 2018 appears to be a difficult year for most of the stocks with only 2 of the 12 stocks(representing about 17% of the analysed portfolio of stocks) generating positive returns on investment. _RUN & ENPH_ were the stocks that generated these positive returns in 2018.

In conclusion and on the premise of the performance analysis carried out over these 2 years, we can position that the stock - _ENPH_ will be a more suitable portfolio for potential investors as it is the only stock that generated positve returns in the 2 years that were reviewed.  

__Original & Refactored Script Execution Times:__ On the average, the execution time of the refactored script for both years (2017 & 2018) was approximately 0.3 seconds (as shown in the above screen shot). On the contrary, the average execution time of the original script for both years was approximately 1.4 seconds. This shows that the refactored script was more efficient than the original script 

## 3. Summary

__1. Advantages or Disadvantages of Refactoring Codes:__

_Some of the general advantages of refactored code include_; 

A) Improved efficiency and speed of execution of the code which leads to faster generation of output / result. This improved efficiency could be as a result improved logic, less memory utilization or taking fewer steps to achieve the output. 

B) Refactoring may also help us better understand our code and identify bugs or redundant lines in the code that do not add extra benefit to the output / result. 

C) Another advantage of refactoring may include the ability to easily / efficiently execute a refactored code against a larger dataset that the original / initial code may not have been able to execute.

_Some of the general disadvantages of refactored code include_; 

A) The fact that code refactoring usually consumes more time to redesign and often diffcult to achieve. 

B) If not properly carried out, code refactoring can be a very risky activity especially if the application that uses the underlying code is already in production environment and being used by customers / users. Improper implementation may lead to distortion in service to customers / users. 



__2. Application of the Advantages and Disadvantages of the Original and Refactored VBA Script:__

The points highlighted below explain how the advantages and disadvantages of code refactoring apply to the original and refactored VBA script; 

_Advantages_ - When compared to the original script, the refactored VBA script loops through our stock dataset just once (_i.e. improved logic & lesser steps_) and in the process gathers the relevant information required to generate the expected output / result (_Ticker, Total Daily Volumes and Returns for each stock_). This logic led to a significantly improved efficiency and speed of execution of the refactored code, which led to faster generation of output / result. In contrast, the original VBA script looped through the dataset thrice to fetch thesame set of information which led to slower generation of output / result. The fact that refactored VBA script uses a logic that ensure only one loop cycle to generate result means that it will also be better suited or adaptable for the analysis of larger stock datasets (more than 12 stocks & more than 2 years dataset) when compared to the original VBA script whose logic is designed to loop through dataset thrice 

_Disadvantages_ - When compared to the original script, the logic behind the refactored VBA script was more difficult and time consuming to design and implement when compared to the original VBA script.

