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
    report_row = 2
    
    'Count number of rows
    'Reference: Excel Documentation "Loop through a list of data by using macros"
    'https://docs.microsoft.com/en-us/office/troubleshoot/excel/loop-through-data-using-macro
    
    'Number of rows in column A with data (assuming first empty cell as end of data)
    num_rows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    
    'Loop rows
    For i = 1 + skip_header To num_rows
            
        'Write ticker of new_stock in report table
        If new_stock = True Then
            Cells(report_row, 9) = Cells(i, 1).Value
            new_stock = False
        End If

    Next i


End Sub