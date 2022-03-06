# Stock Analysis Using Excel VBA

## You asked for me to prepare VBA code to expand our analysis from just selected renewable energy stocks to the entire stock market. For this analysis, I refactored my original code to make it run faster, so that it can efficiently analyze any number of stocks and provide you with the Total Daily Volume and Annual Return for each stock. It also allows you to quickly analyze data for different years. With this information, you can make informed decisions to diversity your parents stock portfolio. 


# Results 

## Stock Performance in 2017 and 2018

Overall, the selected renewable energy stocks performed much better in 2017 than 2018, as depicted in the graph below.

In 2017, most of the stocks had a positive annual return, with several stocks nearing 200% - DAQO New Energy Corp (DQ) at 199.4% and SolarEdge Technologies (SEDG) at 184.5%. Only Terraform Power (TERP) had a negative annual return, at -7.2%. 

In 2018, most of the stocks had a negative annual return, with only two stocks performing positively - Enphase Energy (ENPH) at 81.9% and Sunrun (RUN) at 84.0%.


The total daily volume for each stock fluctuated between the two years with no discernable pattern, as depicted in the graph below. The overall total daily volume for all the renewable energy stocks is about the same in 2017 and 2018 (approximately 3 trillion shares), indicating that interest in trading this stocks is stable.

## Script Execution Time

The stock analysis code was refactored to produce the same results in a simpler and more efficient way. The refactored code performs about 10 times faster than the original code. The code also times itself and reports the time to the user. Below are the message boxes for the original code for each year: 

And the message boxes for the refactored code for each year:

The original code needs to examine each line of the data worksheet 12 times, once for each of the 12 tickers.

'''
(5) Loop through rows in the data
    Sheets(yearValue).Activate
        For j = rowStart To rowEnd
        
        '(5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
            End If
        
    
        '(5b) get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
            End If
            
        '(5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
            End If
            
        Next j
'''

The refactored code only analyzes each line once, detecting when a new ticker has reached and storing the extracted information in an array. It is likely this difference will be exacerbated even more as we expand to the entire stock market, where the refactored code will be much more efficient. See below for a table summary of the script execution time. 

# Summary

The most obvious advantage of refactoring code is that it can result in faster, more efficient, and more generalizable code. 
