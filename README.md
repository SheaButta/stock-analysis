# Stock Analysis using Excel VBA

Dataset: [VBA Challenge - Stock Analysis](https://github.com/SheaButta/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Overview of Project

### Purpose
The purpose of this project is to analyze existing stock data and refactor legacy VBA code to increase 
processing performance. The legacy VBA code processed the data in just over one (1) second; moreover, this effort 
will visualize the performance gain. This data is separated by two (2) worksheets in the "VBA Challenge - Stock Analysis" file.  
The two worksheets are;
- 2017
- 2018


## Results

### Analysis of 2017 Refactoring
Using the 2017 dataset, the refactoring of the VBA code visualized all stocks, except one (1), had returns of 5% or 
greater and rendered a performance gain just under 1 second. The orignal VBA code completed processing in over 1 second.  


### Analysis of 2018 Refactoring
Using the 2018 dataset, the refactoring of the VBA code visualized two (2) stocks with returns over 80%.  The other stocks had no return and
may suggest to sell.  The orignal VBA code completed processing in over 1 second while the refactored code show improvements under 1 second.


## Summary

### Refactoring Stock Analysis
In summary, refactoring the VBA code proved to be beneficial and there was some performance gain for the 2017 and 2018 stock data.  
Documenting the edits, using "arrays" and "for loops" in the VBA code were critical additions to this performance gain.  Although the code
may look a bit more complex, having clear documenation will only help with future development.

![2017 Performance Gain](https://github.com/SheaButta/stock-analysis/blob/main/resources/VBA_Challenge_2017.png)

![2018 Performance Gain](https://github.com/SheaButta/stock-analysis/blob/main/resources/VBA_Challenge_2018.png)
