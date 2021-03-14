Sub VBA_Challenge()


    'Define variables
    Dim num_rows As Integer
    Dim skip_header As Integer
    
    'Skip header flag (0 = False, 1...n = Number of lines to skip)
    skip_header = 1
        
    'new_stock flag
    new_stock = True
        
    'Set report parameters
    'Write headers of report table
    'Start report @Range "I1" = Cells(1,11)
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    'Set report row for first stock
    report_row = 2
    
    'Count number of rows
    'Reference: Excel Documentation "Loop through a list of data by using macros"
    'https://docs.microsoft.com/en-us/office/troubleshoot/excel/loop-through-data-using-macro
    
    'Number of rows in column A with data (assuming first empty cell as end of data)
    num_rows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    'Loop rows
    For i = 1 + skip_header To num_rows
      
        If new_stock = True Then
            'Write ticker symbol of new stock in report table
            Cells(report_row, 9) = Cells(i, 1).Value
            'Store opening price (beginning of year)
            open_year_price = Cells(i, 3).Value
            'Store trading volume (beginning of year)
            stock_volume = Cells(i, 7).Value
            'Set new_stock flag as False
            new_stock = False
        
        'If new_stock is False then accumulate volume of current stock
        Else
            stock_volume = stock_volume + Cells(i, 7).Value
        
            'Check if next row belongs to a new stock
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                'Store closing price of current stock (end of year)
                close_year_price = Cells(i, 6).Value
                'Write yearly price change of current stock in report table
                Cells(report_row, 10).Value = close_year_price - open_year_price
                'Write yearly percent change of price of current stock in report table
                Cells(report_row, 11).Value = (close_year_price - open_year_price) / open_year_price
                'Write accumulated volume of current stock in report table
                Cells(report_row, 12).Value = stock_volume
                
                'Set report row for next stock
                report_row = report_row + 1
                'Set new_stock flag as True
                new_stock = True
            End If
        
        End If

    Next i


End Sub
