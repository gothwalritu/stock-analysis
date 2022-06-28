## Module 2

## Challenge Deliverable 2

# 1.	Overview of Project: 

In this weeks’ challenge we are helping Steve with the analysis of stocks by using VBA.
Steve’s parents want to invest in one of the green energy companies. They want to invest 100% of their money in a company called DAQO New Energy Corp. But Steve wants to diversify the money in other company. So, we are helping him to make decision based on the analysis of previous years’ stocks. We are writing VBA code to perform calculations and automating the analysis for Steve. So, that he can reuse it for any year, and as many times he wanted to perform the analysis with reduced chances to accidents and errors.
We created the macro which enabled the worksheets to be interactive.We are also refactored the code and checked it for the run elapsed time. 

# 2.	Results: 

We developed the code to perform the stocks analysis. We made an interactive automated worksheet, which also produced results for elapsed runtime. The results of the performed analysis are shown in figures 2.1 and 2.2 for 2017 and 2018 respectively. The analysis of stocks for 2017 shows that the highest return was given by ENPH followed by SEDG ( i.e. 199.4% and 184.5% respectively), and the lowest were from TERP (-7.2%). However, the overall returns for year 2017 are better than the year 2018. In 2018, only two company stocks’ ENPH and Run could produce positive returns while the rest of the company’s stocks were negative, which is clearly shown with the color indicators (red for negative returns and green for positive returns). 

#### Figure 2.1: Results of the analysis for year 2017   
![/Resources/Figure_2.1.png](https://github.com/gothwalritu/stock-analysis/blob/main/Resources/Figure_2.1.png)            
                          


#### Figure 2.2: Results of the analysis for year 2018
![/Resources/Figure_2.2.png](https://github.com/gothwalritu/stock-analysis/blob/main/Resources/Figure_2.2.png)


After performing the stocks analysis for the two consecutive years, we refactored the code to improve its quality, readability and performance. The functionality of the code was not altered; hence, it gave the similar results to the original code. However, the elapsed runtime for the code was reduced to some extent. The comparison of the runtime for years 2017 and 2018 with original code and refactored code is shown in table 2.1. 

#### Table 2.1: The runtime values for original and refactored code for year 2017 and 2018.

![/Resources/Table_2.1.png](https://github.com/gothwalritu/stock-analysis/blob/main/Resources/Table_2.1.png)



We were able to refactor the code using the following changes:
#### •	We introduced a variable “tickerIndex” to iterate over all the rows to access the index for four different arrays. 
#### •	We created three output arrays called “tickerVolumes” to replace it with “totalVolume”, “tickerStartingPrices” to replace it with “startingPrices”, and “tickerEndingPrices” to “replace it with endingPrices”.

 #### Screenshot: 1 
 ![/Resources/Screenshot_1.png](https://github.com/gothwalritu/stock-analysis/blob/main/Resources/refactored_code_1.png)
 
 

#### •	Now these variables are storing all the values in arrays and not changing its value over the internal “For” loop for rows. 

#### Screenshot: 2
![/Resources/Screenshot_2.png](https://github.com/gothwalritu/stock-analysis/blob/main/Resources/refactored_code_2.png)
 
#### •	In the end, we added one more “For” loop which produce the output values of “Ticker”, “Total Daily Volume” and “Return” in the final spreadsheet.

#### Screenshot: 3
![/Resources/Screenshot_1.png](https://github.com/gothwalritu/stock-analysis/blob/main/Resources/refactored_code_3.png)



# 3.	Summary:

## Advantages or disadvantages of refactoring code
Refactoring is the process of restructuring the code to improve the operations without altering its functionality and behavior, in simple words we can say that it is a way to improve the quality of the code.  Refactoring makes coding free from bugs and complexity, which helps to run the code faster and makes it more efficient. It helps to reduce the confusion for the first-time users of the code, hence, the main goal of refactoring any code is to enhance and for maintain it for future use. However, refactoring the code requires extra efforts, time and money. Therefore, the code should be refactored when it is going to be used and enhanced further in the future, there is “code smell” detected, there are lots of bugs which arises with different case test run, there is enough time and money for the project. 

## The application of refactoring over the original VBA script
In this module challenge, the code is refactored by introducing the variables, called “tickerIndex”, “tickerVolumes”, “tickerStartingPrices” and “tickerEndingPrices”. After refactoring, the “totalVolume” variable which is carrying only one value is replaced by the “tickerVolumes” which is an array and could hold all the tickers total volume altogether. With this exercise the quality of the code is improved in terms of time and clarity. However, it took me more than a couple of hours to do so. I noticed that every time I was running the code it was showing a different elapsed run time. So, I am not sure why it was giving the different elapsed time, but it was oscillating between 0.75 seconds and 1.2 seconds. It could be due to the performance of my system and how much free space I have on this system.

