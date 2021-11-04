# Overview of Project - Refactoring VBA Script  

## Purpose
The purpose of this project was to edit Module 2 solution code to loop through year 2017 and 2018 data in order to determine if refactoring of the code successfully made the VBA script run faster as well as analyses the performance of 12 different stock performances for 2017 and 2018 to assist Steve research in determining which stock is best for his parents to invest in.

### Background
Steve was very interested in the module 2 exercise and wanted to be able to analyze an entire data-set at the click of a button. Based on that, I obtained performed an additional analysis using module 2 green stocks excel workbook for 12 different stock performances 2017 and 2018. 

### Result of Analysis
This analysis was performed using the codes provided and steps for Deliverable 1 requirements.
 * Downloaded  the `challenge_starter_code.vbs` file and renamed it `VBA_Challenge.vbs`.
  
 * Renamed the green_stocks.xlsm file as VBA_Challenge.xlsm.
 * Added the `VBA_Challenge.vbs` script to the Microsoft Visual Basic editor.
 * `tickerIndex` was set equal to zero before looping over the rows. 
 * Created arrays for `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` 
 * `tickerIndex` was used to access the stock ticker index for the `tickers`, `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` arrays 
 * This script was used to loop through stock data, reading and storing all of the following values from each row:  `tickers`,  `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices`. 
 * Copied and reused the Code for formatting the cells in the spreadsheet from module 2 exercise
 * Added comments to explain the purpose of the code used
 * The outputs for the 2017 and 2018 stock analyses in the `VBA_Challenge.xlsm` workbook was compared to the outputs from the “AllStockAnalysis” in the module and results are the same
 * The pop-up messages showing the elapsed run time for the script were saved as 
 
 `VBA_Challenge_2017.png`
 
  ![VBA_Challenge_2017.png](https://github.com/charleside2001/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)     
 `VBA_Challenge_2018.png`
 
 ![VBA_Challenge_2018.png](https://github.com/charleside2001/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png.png).                
 * Comparing the original and refactored analyses, and it shows tremendous decrease from the original analysis time. Screenshots are provided below for this comparison.
 

## Summary
### Advantages and Disadvantages of Refactoring Code 
 1. Code refactoring creates simplicity and clarity for others to understand 
 2. Per background information from this challenge, refactoring uses fewer steps, less memory and improve the logic of the code to make it easier for future users to read
 3. The disadvantage is that new code learners may not have the knowledge of code refactoring 

### The Advantages and Disadvantages of The Original and Refactored VBA Script

- Refactoring the code saw a tremendous decrease from the original analysis time
  * Below are the original and refactored VBA script  screenshots from the analysis.
  
  `2017 AllStocksAnalysis.png`
  
  ![2017 AllStocksAnalysis.png](https://github.com/charleside2001/Stock-analysis/blob/main/Resources/2017%20AllStocksAnalysis.png)
  
  `VBA_Challenge_2017.png`
  
  ![VBA_Challenge_2017.png](https://github.com/charleside2001/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
  
  `2018 AllStocksAnalysis.png`
  
  ![2018 AllStocksAnalysis.png](https://github.com/charleside2001/Stock-analysis/blob/main/Resources/2018%20AllStocksAnalysis.png)
  
  `VBA_Challenge_2018.png`
  
  ![VBA_Challenge_2018.png](https://github.com/charleside2001/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png.png)
   
  

 [Stock-analysis link on Github](https://github.com/charleside2001/Stock-analysis.git)

