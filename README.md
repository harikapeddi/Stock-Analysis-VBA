# Stock-Analysis_VBA

### Problem Statement

The excel files in this repository have large amounts of data containing different companies' name, stock opening price, stock close price on a specific data of the year. The problem statement has two parts:

1. To find the open price at the beginning of the year, close price at the ending of the year and the total stock volume summarized by each ticker
2. The second part is to find 3 data points
   - The company with greatest increase in the percentage change from the opening price to closing price
   - The company with greatest decrease in the percentage change from the opening price to closing price
   - The company with greatest amount of stock volume



### Solution 

The solution is written in a VBA code to loop through all the rows of each worksheet in the workbook. 

I have used two loops to come up find the required data for the above problem statement

the first loop runs through the data available in each row and checks if the ticker is repeating in the next row. I have used three if conditions to set my criteria

- ​	If the current row ticker is not same as the previous row ticker, then take the current row ticker for open price 
- ​	If the current row ticker is same as the next row ticker, then add the stock volume of the ticker