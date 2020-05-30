# stock_analysis
module 2 stock analysis

#Challenge

For the updated version I didn't really 'switch the nesting order' of the code... I just made the logic so that using the pre-sorted data as the instructions suggested it goes low by row and changes to the next ticker when I got to a new ticker; the if function increases the TickerIndex & logs the endPriceArry(TickerIndex).  As you move down the rows it logs the cumulative volume by the current ticker using my totalVolumeArray(tickerIndex), removing the necessity for the if statement in the original code.  The last array StartingPriceArray is just the same simple if statement from the original code with the TickerIndex logic applied.

Lastly, made a simple loop to display the output from the 3 stored data sets in the array.

