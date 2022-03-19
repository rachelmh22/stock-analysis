# Analyzing Stock for Steve

## Overview of Project

### Purpose

The initial purpose of this project was to help Steve analyze stocks so he can present more options to his parents who want to invest everything into one stock. After presenting Steve with the analysis, we want to refactor the code to make it more efficient. We want to understand if we can refactor the code to make the script run more quickly so that we can use it to analyze multiple stocks, rather than the 12 that Steve asked to be analyzed. 

## Results

### Analysis of the Stocks in 2017

The results of the analysis illustrated that all the stocks analyzed except for one had a positive return. Therefore, 2017 was a good year for this who invested. It seems everyone had some kind of return, and some more than others. It was an especially good year for those who invested in DQ, ENPH, FSLR, and SEDG; as all 4 stocks had over 100% return. The only stock that failed to generate a return on its investment is TERP but the negative return is fairly low. 

### Analysis of Refactored Code for 2017 Data 
As for the code, after refactoring it, the script run time had decreased by half. Originally, it took 0.6 seconds to run but after the refactored code, it took 0.3 seconds. 

### Analysis of the Stocks in 2018

Stock investment in 2018 for the stocks analyzed was not as positive as in 2017. Almost all stocks generated a negative return, except for ENPH and RUN. Both those stocks had a fairly high return rate. This means those who invested in ENPH in 2017 and continued to invest in it in 2018 had very good investments since both years generated high returns. On the other hand, DQ and JKS had very bad years as they had the highest negative return. This would be very devastating for anyone who invested in DQ in 2017 and continued to invest in the following year as they had a very high return in 2017. 

### Analysis of Refactored Code for 2018 Data 
The code ran much quicker for the 2018 data. After running the refactored code, the script run time was 0.1 seconds, which is much quicker than for the 2017 data. The original run time for the script was 0.6 seconds but the refactored code was much more efficient.

### Refactored Code

‘’’

    tickerIndex = 0


    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
 
    For i = 0 To 11
    tickerVolumes(i) = 0
    Next i

        
    For i = 2 To RowCount

    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
        
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
   
      
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
      
    tickerIndex = tickerIndex + 1
    End If

    Next i
   

    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
‘’’

## Summary

### Advantages of Refactoring Code

From this project, the biggest advantage of refactoring code is making it so the VBA script runs more efficiently. While both codes were able to run in less than a second, a refactored code would be much more effective if there more a very high number of stocks and stock data. Additionally, refactored code is much clean and easier to understand. Since the code is made simpler, it would be more understandable and in turn, easier to maintain and adjust. 

### Disadvantages of Refactoring Code
A disadvantage to refactoring code is the time consumption. It takes time to go through the code and find the areas that can be made better and when you find what can be fixed, it will take time to look over the code, understand it, then refactor it. This would be especially true and time consuming if the code was given by another person so it would be unfamiliar. Additionally, refactoring code may mean running into errors. Since the original code worked, rewriting areas and refactoring might introduce errors before the code can be made simpler. 

### Advantages and Disadvantages of Refactoring the Original VBA Script for this Project

The biggest advantage mentioned above applied to this project’s script as after the code was refactored, the script run time decreased significantly.  This makes the code more efficient and since it is made simpler as well, the code would be easier to adjust and can be applied to larger data sets and will run with less time. Regarding the disadvantage, it also applied for this project. The time it took to refactor the code was about as long, if not, longer as writing the original code. This was because of the unfamiliarity of the code. Along with that, refactoring the code gave me some errors and took me time to fix before it ran successfully. However, it was fortunate that with the time spent, the code was able to run successfully and much more efficiently after refactoring.
