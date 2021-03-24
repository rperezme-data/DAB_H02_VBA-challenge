Sub VBA_Challenge()

    'VARIABLES
    'Declare variables and set type
    Dim ws As Worksheet
    Dim last_row As Double
    Dim skip_header As Integer
    Dim report_row As Double
    
    Dim ticker As String
    Dim open_year_price As Double
    Dim close_year_price As Double
    Dim stock_volume As Double
    
    Dim year_price_change As Double
    Dim year_pct_change As Double
    
    'FLAGS
    'Skip header flag (0 = False, 1...n = Number of lines to skip)
    skip_header = 1
     
    'WORKSHEET LOOP
    'Loop worksheets to analyse stock information
    For Each ws In Worksheets
      
        'REPORT TABLE
        'Write headers of report table (start @ws.Cells(1, 9)
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        'Set column width (autofit)
        ws.Columns("L").AutoFit
        'Set first report row
        report_row = 2
                
        'FIND LAST ROW
        'Get row number of last row in column A with data
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'STOCK ANALYSIS
        'Get open price for first stock
        open_year_price = ws.Cells(1 + skip_header, 3).Value
        
        'ROW LOOP
        'Loop rows with data to obtain information & report in table
        For i = 1 + skip_header To last_row
            
            'Check if next row belongs to a new stock
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
                'Get ticker of current stock
                ticker = ws.Cells(i, 1).Value
                
                'Store closing price of current stock (end of year)
                close_year_price = ws.Cells(i, 6).Value
                
                'Accumulate volume of current stock
                stock_volume = stock_volume + ws.Cells(i, 7).Value
    
                'COMPUTE & REPORT RESULTS
                
                'Report ticker in table
                ws.Cells(report_row, 9).Value = ticker
                
                'Compute & Report Yearly price change of current stock in table
                year_price_change = close_year_price - open_year_price
                ws.Cells(report_row, 10).Value = year_price_change
                
                'Format Yearly price change of current stock in table
                If ws.Cells(report_row, 10).Value > 0 Then
                    ws.Cells(report_row, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(report_row, 10).Value < 0 Then
                    ws.Cells(report_row, 10).Interior.ColorIndex = 3
                End If
                           
                'Compute, Report & Format yearly percent change of price of current stock in table
                If open_year_price = 0 Then
                    ws.Cells(report_row, 11).Value = "DIV/0 Error"
                    ws.Cells(report_row, 11).HorizontalAlignment = xlRight
                Else
                    year_pct_change = year_price_change / open_year_price
                    ws.Cells(report_row, 11).Value = year_pct_change
                    ws.Cells(report_row, 11).NumberFormat = "0.00%"
                End If
                
                'Report accumulated volume of current stock in table
                ws.Cells(report_row, 12).Value = stock_volume
                
                'Set report row for next stock
                report_row = report_row + 1
                                
                'SET VALUES FOR NEXT STOCK
                
                'Get open price for next stock
                open_year_price = ws.Cells(i + 1, 3).Value
                
                'Reset stock volume variable
                stock_volume = 0
            
            Else
            
            'If next row belongs to current stock, only accumulate volume of current stock
            stock_volume = stock_volume + ws.Cells(i, 7).Value
                    
            End If
    
        Next i
       
    Next ws

End Sub