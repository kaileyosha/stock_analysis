# Stock Analysis (VBA Challenge)

The purpose of this challenge was to improve stock analysis code through refactoring. The hope was not to make new discoveries, but to make the code more efficient. 

## Results 

The refactoring of this code started with the creation of three more arrays, tickerVolume, tickerStartingPrices, and tickerEndingPrices to track stock data over the two year period. 

```  
Dim tickerVolumes(12) As Long
     
Dim tickerStartingPrices(12) As Single
     
Dim tickerEndingPrices(12) As Single
```

Then a ``for loop `` was used to iterate through the rows of the spreadsheet for the year the user selected, and `` if then `` statements were used to pull the relevant material. Then another `` for loop `` was used to activate the stock analysis, insert the appropriate data, and format it into an easily readable table. 

## Conclusion

The re-factored code featured a distinct advantage, which is that it ran much quicker than the original code. It also did much more than the original macro that only put the data on a spreadsheet and used a separate macro to format the data into a table. Personally, while I can see the benefits of refactoring, I prefer the original method of coding. This most likely boils down to my growing experience with VBA, but at this early stage the refactoring required a lot more detail and debugging than the standard code did. While the original code takes longer to run, it was simpler for me as a new analyst to follow and easier for me to assess and fix. 
