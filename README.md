# Stock Market Analysis  
Stock information from 2014-2016 analyzed with a Macro VB Script. Yearly change, percent change, total stock volume shown for each ticker. Conditional formatting used to show increase
or decrease. Stock identified for greatest increase, decrease, and total volume summed for each year. 

## Key Items

-```Summary_Table_Row = 2``` <-Row number for second table, which has 1 added to it to go to next row

-Looping needed to be done across all worksheets to derive results for rows Cells (i,1). The last row was created by ```ws.Cells(Rows.Count, 1).End(xlUp).Row```.

-Ticker symbol in cell was was checked to see if in same stock with if statement ```If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value``` and then calculations performed  

-Counters created to calculate yearly change and yearly change percent to dynamically move back to the first row of each ticker:  

```counter = WorksheetFunction.CountIf(ws.Range("A:A"), ticker)``` <-counts number of times ticker values appear  
 ```counter2 = ws.Cells(i - (counter - 1), 3)``` <-takes ticker count to back to first row for opening price. Takes row number and subtracts number of count of ticker minus 1 to go back to the first line.

-conditional formatting done by recording Macro and copying VB Script
 
-Min, Max, and Match functions used to get values and get corresponding ticker 

