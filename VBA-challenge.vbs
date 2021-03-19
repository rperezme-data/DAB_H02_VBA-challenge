Sub VBA_Challenge()


    'TIMEIT
    'Time script execution
    Dim start_time As Double
    Dim end_time As Double
    start_time = Hour(Now) * 3600 + Minute(Now) * 60 + Second(Now)

    'VARIABLES
    'Declare variable type
    'Dim ws As Worksheet
    Dim last_row As Double
    Dim skip_header As Integer
    Dim report_row As Double
            
    Dim open_year_price As Double
    Dim stock_volume As Double
    Dim close_year_price As Double
    
    'FLAGS
    'Skip Header flag:
        '0 = No header lines to skip
        '1...n = Number of header lines to skip
    skip_header = 1
     
    'WORKSHEET LOOP
    'Loop worksheets to analyse stock information
    For Each ws In Worksheets
    'MsgBox (ws.Name)
    
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
        'MsgBox (last_row)
        
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
                
                'SET VALUES FOR NEW STOCK
                'Get open price for first stock
                open_year_price = ws.Cells(i + 1, 3).Value
                
                'Reset stock volume variable
                stock_volume = 0
            
            Else
            
            'If next row belongs to the current stock
            stock_volume = stock_volume + ws.Cells(i, 7).Value
                    
            End If
    
        Next i
    
    
        'BONUS REPORT
        'Write headers of bonus report (start @ws.Cells(1, 15)
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        'Set column width (autofit)
        ws.Columns("O").AutoFit
    
    
    Next ws

    end_time = Hour(Now) * 3600 + Minute(Now) * 60 + Second(Now)
    MsgBox ("Execution time: " & end_time - start_time & " seconds")

End Sub
