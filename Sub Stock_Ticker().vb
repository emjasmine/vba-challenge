Sub Stock_Ticker()

'loop through the worksheets
For Each ws In ActiveWorkbook.Worksheets

          ' Set an initial variable for holding the stock ticker
          Dim Stock_Ticker As String
        
          ' Set an initial variable for holding the total per stock
          Dim Ticker_Total As Double
          Ticker_Total = 0
          
          'set the initial variable for yearly change
         Dim yearly_change As Double
         
         'set the initial variable for percent change
         Dim percent_change As Variant
          
          'set the last cell and last stock
          Dim last_cell As Long
        last_cell = 1 + ws.Cells(Rows.Count, "A").End(xlUp).Row
        last_stock = 1 + ws.Cells(Rows.Count, "G").End(xlUp).Row
        
        
          ' Keep track of the location for each stock ticker in the summary table
          Dim Summary_Table_Row As Integer
          Summary_Table_Row = 2
          last_cell = 1 + ws.Cells(Rows.Count, "A").End(xlUp).Row
          
          'Calculate the yearly change (difference between open and close), need to pull opening price on first day (datemin) and closing price on last day (datemax) and  calculate open/close
        'this will be the first stock's open
        Dim annual_open As Double
        annual_open = ws.Cells(2, 3).Value
        
        'set annual close price
        Dim annual_close As Double
        
         'insert the summary table headers
         ws.Cells(1, 10).Value = "Ticker"
         ws.Cells(1, 11).Value = "Yearly Change"
         ws.Cells(1, 12).Value = "Percent Change"
         ws.Cells(1, 13).Value = "Total Stock Volume"
         
         'insert headers for the bonus summary
         ws.Cells(1, 16).Value = "Ticker"
         ws.Cells(1, 17).Value = "Value"
         ws.Cells(2, 15).Value = "Greatest % increase"
         ws.Cells(3, 15).Value = "Greatest % decrease"
         ws.Cells(4, 15).Value = "Greatest Total Volume"
        
         
         'set the intial value for the greatest % increase
         Max = ws.Cells(2, 12).Value
         
         'set the initial value for the greatest % decrease
         Min = ws.Cells(2, 12).Value
         
         'set the initial value for the greatest total volume
         Max_volume = ws.Cells(2, 13).Value
         
          ' Loop through the stocks, columns A thru G
          For i = 2 To last_cell
        
                ' Check if we are still within the same stock , if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                  ' Set the Stock Ticker
                  Stock_Ticker = ws.Cells(i, 1).Value
            
                  ' Add to the Stock TickerTotal
                  Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                  
                   ' Here, i is the last row of the current stock, so assign it
                    annual_close = ws.Cells(i, 6).Value
            
                    
                    'calculate the yearly change
                    yearly_change = annual_open - annual_close
                    
                     'addresses if the opening value is 0, to look at the next row
                    If annual_open = 0 Then
                        annual_open = ws.Cells(i + 2, 3).Value
                        'check that this new opening price is not 0
                        If annual_open <> 0 Then
                            annual_open = ws.Cells(i + 2, 3).Value
                        End If
                        
                    End If
                    
                    'calculate the percent change
                    percent_change = yearly_change / annual_open
                    
                    ' Here, i + 1 is the first row of next stock
                    ' so assign it
                    annual_open = ws.Cells(i + 1, 3).Value

        
            
                  ' Print the Stock Ticker in the Summary Table
                  ws.Range("J" & Summary_Table_Row).Value = Stock_Ticker
            
                  ' Print the Stock volumen total to the Summary Table
                  ws.Range("M" & Summary_Table_Row).Value = Ticker_Total
                  
                  'Print the yearly change to summary table
                  ws.Range("K" & Summary_Table_Row).Value = yearly_change
                  
                  'Print the percent change to summary table
                  ws.Range("L" & Summary_Table_Row).Value = percent_change
            
                  ' Add one to the summary table row
                  Summary_Table_Row = Summary_Table_Row + 1
                  
                  ' Reset the Ticker Total
                  Ticker_Total = 0
            
                ' If the cell immediately following a row is the same brand...
                Else
            
                  ' Add to the Ticker Total
                  Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            
            End If
    
        Next i
        
        'loops through summary table, columns J thru M, for bonus exercise
        For i = 3 To Summary_Table_Row
            'check to see if the next percent change in row i is larger then current Max(greatest % increase)
            If ws.Cells(i, 12).Value > Max Then
                'if yes, then resets the Max value to the row 1 value
                Max = ws.Cells(i, 12).Value
                'Get the ticker for the corresponding new Max value
                Max_ticker = ws.Cells(i, 10).Value
            'checks to see if the next percent change in row i is smaller than the current Min(greatest % decrease)
            ElseIf ws.Cells(i, 12).Value < Min Then
                Min = ws.Cells(i, 12).Value
                'get the ticker for the corresponding new Min value
                Min_ticker = ws.Cells(i, 10).Value
            End If
            
            'checks to see if the next total volume is greater than the current
            If ws.Cells(i, 13).Value > Max_volume Then
                Max_volume = ws.Cells(i, 13).Value
                'get the ticker for the corresponding new Max_volume
                Max_volume_ticker = ws.Cells(i, 10).Value
            End If
        
        Next i
        
        'print the ticker for the Max value to the summary table
        ws.Cells(2, 16).Value = Max_ticker
        'print ticker for the Min value to the summary table
        ws.Cells(3, 16).Value = Min_ticker
        'print ticker for the Max Volume to the summary table
        ws.Cells(4, 16).Value = Max_volume_ticker
        'print Max value to summary table
        ws.Cells(2, 17).Value = Max
        'print Min value to summary table
        ws.Cells(3, 17).Value = Min
        'print Max volume total to summary table
        ws.Cells(4, 17).Value = Max_volume
        
    'loop through for column K for the conditional formating
    For i = 2 To Summary_Table_Row
     ' conditional formatting cell color
                      If ws.Cells(i, 11) < 0 Then
                          ws.Cells(i, 11).Interior.ColorIndex = 3
                      Else
                          ws.Cells(i, 11).Interior.ColorIndex = 4
                      
                      End If
    Next i
    
Next ws

End Sub


