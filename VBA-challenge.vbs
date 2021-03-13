Sub VBA_Challenge()


    'Define variables
    Dim num_rows As Integer
    
    
    'Count number of rows (with data) in Column A
    'Reference: Excel Documentation "Loop through a list of data by using macros"
    'https://docs.microsoft.com/en-us/office/troubleshoot/excel/loop-through-data-using-macro
    
    'Number of rows with data (assuming first empty cell as end of data in column A)
    num_rows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    MsgBox "NumRows =" + Str(num_rows)



End Sub

