Sub VBA_Challenge()


    'VARIABLES
    'Declare variable type
    Dim num_rows As Double
    Dim skip_header As Integer
    Dim new_stock As Boolean
    Dim report_row As Double
    
    Dim stock_volume As Double
    Dim open_year_price As Double
    Dim close_year_price As Double
    
    'FLAGS
    'Skip Header flag:
        '0 = No header lines to skip
        '1...n = Number of header lines to skip
    skip_header = 1
    
    'New Stock flag:
        'True = Next row is data of new stock (to be analysed)
        'False = Next row is data of current stock (being analysed)
    new_stock = True
        
    'REPORT TABLE
    'Write headers of report table (start @Cells(1, 9)
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    'Set first report row
    report_row = 2
            
    'COUNT ROWS
    '---------------------------
    'Reference: Excel Documentation "Loop through a list of data by using macros"
    'https://docs.microsoft.com/en-us/office/troubleshoot/excel/loop-through-data-using-macro
    '---------------------------
    'Number of rows in column A with data (assuming first empty cell as end of data)
    num_rows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    'STOCK ANALYSIS
    
    'Loop rows with data to obtain information & report in table
    For i = 1 + skip_header To num_rows
        
        'Start analysis for new stock
        If new_stock = True Then
            
            'Report ticker symbol of new stock in table
            Cells(report_row, 9) = Cells(i, 1).Value
            
            'Store opening price (beginning of year)
            open_year_price = Cells(i, 3).Value
            
            'Store trading volume (beginning of year)
            stock_volume = Cells(i, 7).Value
            
            'Set new_stock flag as False
            new_stock = False
        
        'Acumulate volume of current stock
        Else
            stock_volume = stock_volume + Cells(i, 7).Value
        
            'Check if next row belongs to a new stock
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                
                'Store closing price of current stock (end of year)
                close_year_price = Cells(i, 6).Value
                
                'Report yearly price change of current stock in table
                Cells(report_row, 10).Value = close_year_price - open_year_price
                
                'Format yearly place change
                If Cells(report_row, 10).Value > 0 Then
                    Cells(report_row, 10).Interior.ColorIndex = 4
                ElseIf Cells(report_row, 10).Value < 0 Then
                    Cells(report_row, 10).Interior.ColorIndex = 3
                End If
                
                'Report & Format yearly percent change of price of current stock in table
                Cells(report_row, 11).Value = (close_year_price - open_year_price) / open_year_price
                Cells(report_row, 11).NumberFormat = "0.00%"
                                
                'Report accumulated volume of current stock in table
                Cells(report_row, 12).Value = stock_volume
                
                'Set report row for next stock
                report_row = report_row + 1
                
                'Set new_stock flag as True
                new_stock = True
                
            End If
        
        End If

    Next i


End Sub