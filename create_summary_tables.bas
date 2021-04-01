Attribute VB_Name = "Module1"
Sub WallStreet()
    
    Dim ws As Worksheet
    
    'Loops through all sheets
    For Each ws In Worksheets
        
        'Find the last row of each worksheet
        last_row = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        'Declare variables to store information
        Dim start_price As Double
        Dim end_price As Double
        Dim price_change As Double
        Dim percent_change As Double
        Dim stock_volume As Double
        Dim row As Integer
        Dim column As Integer
        
        'Assign values to some variables
        stock_volume = 0
        row = 2
        start_price = ws.Range("C2")
        
        'Create headings for Summary Table
        ws.Range("J1") = "Ticker"
        ws.Range("K1") = "Yearly Change"
        ws.Range("L1") = "Percentage Change"
        ws.Range("M1") = "Total Stock Volume"
        
        'Loops through all rows of sheet
        For i = 2 To last_row
            
            'Add stock volume until cells don't equal
            stock_volume = stock_volume + ws.Cells(i, 7)
            
            'Stops once tickers don't match
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Put ticker into Summary Table
                ws.Cells(row, 10).Value = ws.Cells(i, 1).Value
                
                'Store ending price in variable and calculate price change
                end_price = ws.Cells(i, 6).Value
                price_change = end_price - start_price
                ws.Cells(row, 11).Value = price_change
                
                'Calculate and display percentage change
                If price_change = 0 Or start_price = 0 Then
                    ws.Cells(row, 12).Value = 0
                Else
                    percent_change = price_change / start_price
                    ws.Cells(row, 12).Value = percent_change
                End If
                
                'Format percentage change
                If ws.Cells(row, 12).Value > 0 Then
                    ws.Cells(row, 12).Interior.ColorIndex = 4
                ElseIf ws.Cells(row, 12).Value < 0 Then
                    ws.Cells(row, 12).Interior.ColorIndex = 3
                End If
                ws.Columns(12).NumberFormat = "0.00%"
                
                'Display stock_volume and then set to 0 for next iteration
                ws.Cells(row, 13).Value = stock_volume
                stock_volume = 0
                
                'Store starting price for next iteration
                start_price = ws.Cells(i + 1, 3).Value
                
                row = row + 1
    
            End If
                
        Next i
        
    Next ws
    
    
    'Create headings for Second Summary Table
    Range("P2") = "Greatest % Increase"
    Range("P3") = "Greatest % Decrease"
    Range("P4") = "Greatest Total Volume"
    Range("Q1") = "Ticker"
    Range("R1") = "Value"
        
    'Declare variables for Second Summary Table
    Dim largest_pct_ticker As String
    Dim largest_pct_value As Double
        
    Dim smallest_pct_ticker As String
    Dim smallest_pct_value As Double
        
    Dim largest_vol_ticker As String
    Dim largest_vol_value As Double
        
    'Set variables to starting point
    largest_pct_value = 0
    smallest_pct_value = 0
    largest_vol_value = 0
    
    
    Dim wh As Worksheet
    
    'Loops through all sheets
    For Each wh In Worksheets
        
        'Find last row of First Summary Table
        last_row_two = wh.Cells(Rows.Count, "J").End(xlUp).row
        
        'Loops through First Summary Table
        For j = 2 To last_row_two
            
            'If percentage change is larger then store that number and ticker
            If wh.Cells(j, 12).Value > largest_pct_value Then
                largest_pct_value = wh.Cells(j, 12).Value
                largest_pct_ticker = wh.Cells(j, 10).Value
            End If
            
            'If percentage change is smaller then store that number and ticker
            If wh.Cells(j, 12).Value < smallest_pct_value Then
                smallest_pct_value = wh.Cells(j, 12).Value
                smallest_pct_ticker = wh.Cells(j, 10).Value
            End If
            
            'If stock volume is larger then store that number and ticker
            If wh.Cells(j, 13).Value > largest_vol_value Then
                largest_vol_value = wh.Cells(j, 13).Value
                largest_vol_ticker = wh.Cells(j, 10).Value
            End If
            
        Next j
        
    Next wh
    
    'Display Values
    Range("Q2").Value = largest_pct_ticker
    Range("R2").Value = largest_pct_value
        
    Range("Q3").Value = smallest_pct_ticker
    Range("R3").Value = smallest_pct_value
        
    Range("Q4").Value = largest_vol_ticker
    Range("R4").Value = largest_vol_value
    
    'Format Percentages
    Range("R2").NumberFormat = "0.00%"
    Range("R3").NumberFormat = "0.00%"

End Sub
