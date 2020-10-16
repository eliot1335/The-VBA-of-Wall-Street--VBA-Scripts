Attribute VB_Name = "Module1"
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
    Dim yearly_change_lastrow As Long
    yearly_change_lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row

    ' Loop through that column
    For i = 2 To yearly_change_lastrow

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

    ' Find highest and lowest percent changed and record them
    
    Dim challenge_lastrow As Long
    challenge_lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    
    Dim highest_per As Double
    highest_per = Application.WorksheetFunction.Max(ws.Range("K2:K" & challenge_lastrow))
    ws.Cells(2, 17).Value = highest_per
    ws.Cells(2, 17).Style = "Percent"
    
    
    Dim lowest_per As Double
    lowest_per = Application.WorksheetFunction.Min(ws.Range("K2:K" & challenge_lastrow))
    ws.Cells(3, 17).Value = lowest_per
    ws.Cells(3, 17).Style = "Percent"
    
    For i = 2 To challenge_lastrow
        If ws.Cells(i, 11).Value = highest_per Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            
        ElseIf ws.Cells(i, 11).Value = lowest_per Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
        End If
        
    Next i
    
    
    ' Failed attempts
    '----------------------------------------------------------------------
    ' loop through percent changed column to find highest and lowest value and output
    ' For i = 2 To challenge_lastrow
    
    ' Declaring variables for highest and lowest percent change
    ' Dim highest_change As Double
    ' highest_change = 0
    ' Dim lowest_change As Double
    ' lowest_change = 0
    ' Dim highest_per As Double
    ' highest_per =  ws.Cells(i, 11)
    ' Dim lowest_per As Double
    
    
    ' If statement for to find highest and lowest value within the column
        ' If highest_per < ws.Cells(i + 1, 11).Value Then
           ' Record the highest at the moment to variable
           ' highest_per = ws.Cells(i, 11).Value
           
        ' ws.Cells(2, 17).Value = highest_per
        ' ws.Cells(2, 17).Style = "Percent"
           
           ' Output ticker symbol
           ' ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
        
        
        
        'ElseIf lowest_change > ws.Cells(i, 10).Value Then
            'Record the lowest at the moment to variable
            'lowest_per = ws.Cells(i, 11).Value
            
            'ws.Cells(3, 17).Value = lowest_per
            'ws.Cells(3, 17).Style = "Percent"
            
            ' Output ticker symbol
            'ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            
        ' End If
    
    ' Next i
    ' Still getting bugs here, should I be using 2 For loops?
    
'----------------------------------------------------------------------
    Dim highest_total_vol As LongLong
    highest_total_vol = 0
    
    
    For i = 2 To challenge_lastrow
    
        If highest_total_vol < ws.Cells(i, 12).Value Then
        highest_total_vol = ws.Cells(i, 12).Value
        ws.Cells(4, 17).Value = highest_total_vol
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    
        End If
    
    Next i
    

Next ws


End Sub
