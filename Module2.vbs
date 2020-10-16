Attribute VB_Name = "Module2"
Sub Multiple_year_stock_summary()

    ' Build headers required for each worksheet
For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    

    ' Declaring variables types and store initial value
    Dim i As Long
    
    ' Row variable for the new table
    Dim TSBRow As Integer
    TSBRow = 2
    
    ' Set initial open_yr
    Dim open_yr As Double
    open_yr = ws.Cells(2, 3).Value
    
    
    Dim close_yr As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    
    Dim total_volume As LongLong
    
    total_volume = 0
    
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    ' Adding loop and if statement
    ' if ticker = next ticker during the loop, then accumulate vol to total_volume
    ' if tickre <> next ticker, then grab close_yr, grab new open_yr
        ' and output "yearly change"/"percent change"/"total volume"
    
    
    For i = 2 To lastrow
 
    
        If ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            total_volume = total_volume + ws.Cells(i, 7).Value
       
       
        Else
        ' If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value then
        ' Grab ticker symbol and output
            ws.Cells(TSBRow, 9).Value = ws.Cells(i, 1).Value
       
        ' Grab close_yr
        ' Calculate yearly change by close_yr - open_yr and output
            close_yr = ws.Cells(i, 6).Value
            yearly_change = close_yr - open_yr
            ws.Cells(TSBRow, 10).Value = yearly_change
       
        ' Calculate percent change by yearly_change/open_yr and then output
            
            ' To counter divided by 0 bug
            If open_yr = 0 Then
               percent_change = 0
               ws.Cells(TSBRow, 11).Value = percent_change
               ws.Cells(TSBRow, 11).Style = "percent"
            
            Else
               percent_change = (yearly_change / open_yr)
               ws.Cells(TSBRow, 11).Value = percent_change
               ws.Cells(TSBRow, 11).Style = "percent"
            
            End If
            
        ' Add up total_volume and output
            total_volume = total_volume + ws.Cells(i, 7).Value
            ws.Cells(TSBRow, 12).Value = total_volume
       
        ' Reset variable values for new ticker symbol
            ' new open_yr
            open_yr = ws.Cells(i + 1, 3).Value
            ' new total_volume
            total_volume = 0
            ' new TSBRow
            TSBRow = TSBRow + 1
        
        End If
        
    Next i

'----------------------------------------------------------------------

    ' Conditional format "Yearly Change" cell
    ' Declaring variables
    Dim summary_table_lastrow As Long
    summary_table_lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    ' Loop through that column
    For i = 2 To summary_table_lastrow

    ' If statement positive green, negative red
        If ws.Cells(i, 10).Value >= 0 Then
           ws.Cells(i, 10).Interior.ColorIndex = 4
       
        Else
    
           ws.Cells(i, 10).Interior.ColorIndex = 3
    
        End If
    
    Next i
    
'----------------------------------------------------------------------

    ' Build challenge summary table headers
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'----------------------------------------------------------------------

    ' loop through percent changed column to find highest and lowest value and output
    ' Declaring variables for highest and lowest percent change and the lastrow of percent change
    ' Also need to initialize highest_per and lowest_per variable to the first box of percent change
    
     
    Dim highest_per As Double
    highest_per = ws.Cells(i, 11)
    Dim lowest_per As Double
    lowest_per = ws.Cells(i, 11)
     
     
    ' Start the For loop
    For i = 2 To summary_table_lastrow
    
    ' If statement for to find highest value within the column
    ' Compare current row with next row
         If highest_per < ws.Cells(i + 1, 11).Value Then
         
            ' Record the higher one to variable and populate the value to designated boxes
            highest_per = ws.Cells(i + 1, 11).Value
            ws.Cells(2, 17).Value = highest_per
            ws.Cells(2, 17).Style = "Percent"
            ' Output ticker symbol
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
        
        
        ElseIf lowest_per > ws.Cells(i + 1, 11).Value Then
            ' Record the lower one to variable, and populate value to boxes
            lowest_per = ws.Cells(i + 1, 11).Value
            ws.Cells(3, 17).Value = lowest_per
            ws.Cells(3, 17).Style = "Percent"
            ' Output ticker symbol
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
        End If
    
    Next i
    
'----------------------------------------------------------------------
    ' To find the greatest total volume
    ' Declaring variable for highest total volume
    
    Dim highest_total_vol As LongLong
    highest_total_vol = 0
    
    
    For i = 2 To summary_table_lastrow
    
        If highest_total_vol < ws.Cells(i, 12).Value Then
        highest_total_vol = ws.Cells(i, 12).Value
        ws.Cells(4, 17).Value = highest_total_vol
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    
        End If
    
    Next i
    

Next ws


End Sub
