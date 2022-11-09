# vba_challenge

## Overview of Project
Analysis of certain stocks was carried out for years 2017 and 2018 in an Excel workbook, with macros written in VBA, and output to a new workbook.

## Results
### 2017 Stock Performance

![](/images/2017_results.png)

Stocks are clearly up across the board, with the main benefactors being DQ (199.4%) and SEDG (184.5%). TERP is the only negative stock at -7.2%.

This macro runs in 0.082 seconds.

![](/images/2017_runtime.png)

### 2018 Stock Performance

![](/images/2018_results.png)

As opposed to 2017, 2018 was a very bad year for our tickers. While ENPH (81.9%) and RUN (84.0%) had impressive performances, every single other stock was down, with DQ suffering the most at -62.6%.

This macro runs in 0.082 seconds.

![](/images/2018_runtime.png)

## Summary

Refactoring code is important when new functionality needs to be added or the scope of the project needs to be expanded. This is particularly important when the magnitude of the data being processed is increased, leading to exponential gains in runtime.

The refactored code runs incredibly quickly over 12 stocks for a year, and it is all localized in a single button press. However, since 12 stocks run in ust under 0.1 seconds, expanding this code to the entire stock market would not be feasible, since the runtime would scale linearly with stocks added. It is recommended to refactor the code once again for a large increase in stocks, or entirely ditching VBA as a platform is also recommended.