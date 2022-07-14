# stock-analysis

## Overview of Project
The goal of this project is to demonstrate proficiency in Microsoft Excel VBA code by completing a macro designed to collect & analyze stock information over the years 2017 & 2018. The data provided for the exercise covers 12 different stocks daily performance for every day in a given year. Each day provided the stock's opening, high, low, & closing price as well as the investor's return. The macro is built and run to capture and record the Ticker, Accumulated Daily Volume, and Percentage of Return on Investment.

**ALSO** An original macro was provided accomplishing this in a previous exercise, however an additional goal of this module is to edit the code & improve its efficiency.

## Results

### Stock Performance
The resulting data is presented as grouped by years 2017 and 2018. The stocks yielding a positive return are highlighted in green while those yielding a negative return are highlighted in red. 
- **2017**
Nearly all stocks returned a positive investment save for stock ticker "TERP" which ended the year 7.2% lower than it began.
The Excel File VBA_Challenge.xlsx displays Yearly breadown of each stock, providing Total Daily Volume and the YoY% return.

- **2018** 
Nearly all stocks returned a negative investment save for stocks "ENPH" and "RUN", which returned a positive yield of 82.9% and 84.0% respectively.

![image](https://user-images.githubusercontent.com/47199557/178871547-72f8f8f1-5bd3-4f40-8038-fc8332d4ebe8.png)

### Code Performance
The original macro provided completed its analysis of 2017 and 2018 in ~.59 seconds and ~.48 seconds respectively.
HOWEVER, the refactored code completed its analysis of 2017 and 2018 MUCH more quickly, in ~.09 seconds each.

**Inefficient Analysis Code**

![VBA_Challenge_2017_INEFFICIENT_CODE](https://user-images.githubusercontent.com/47199557/178882895-d521f43e-e3dc-4b2e-a203-2d35268f5e6c.png)
![VBA_Challenge_2018_INEFFICIENT_CODE](https://user-images.githubusercontent.com/47199557/178882905-b79f8c88-79ad-4427-9437-bd47834ffc71.png)



**Refactored Analysis Code**

![VBA_Challenge_2017](https://user-images.githubusercontent.com/47199557/178883015-afefcc90-bbd0-455d-b4fd-c5917a038273.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/47199557/178883045-983c8ba5-87d0-4ddd-a7cb-0d1b7d21db0b.png)


## Summary
