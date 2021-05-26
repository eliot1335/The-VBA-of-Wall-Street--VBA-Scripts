# The VBA of Wall Street - VBA Scripts


These two VBScript files serve the same purpose: to loop through the 
raw data of <02-VBA-Scripting_Homework_Instructions_Resources_Multiple_year_stock_data> and 
populate the following info: Yearly Change/Percent Change/Total Stock Trading Volume/
for each ticker symbol each year in a summary table.

On top of the above info, these scripts would also make a conditional 
format to highlight the positively and negatively performed stock for easier reading. Also, 
the scripts will populate another summary table to display the ticker symbols with the 
"Greatest % increase," "Greatest % decrease," and "Greatest total volume."


Both scripts took the same approach in solving the calculation since:
My scripts will loop through the raw data from top to bottom and collect all the elements needed:
the opening price at the beginning of the year, closing price at the end of the year, and each stock trading volume. 
Record them into variables, perform its calculation base on them, and populate the assignment info.

The only difference between these two is their approach to finding out the highest and lowest percent change. 
Module 1 utilized Application.Worksheetfunction.Max/Min to populate the "Great% increase," "Greatest % decrease," values. 
And then use these values to backtrack the ticker symbol. Whereas Module 2 will be using a For loop to compare the current 
row and next row to determine the highest and lowest values.
