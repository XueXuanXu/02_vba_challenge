Sub ticker():


    'create a variable that store the ticker name
    Dim ticker As String
    ticker = Cells(2, 1).Value
    
    
    'create a variable that store the row number that ticker should be store in
    Dim ticker_row As Integer
    ticker_row = 2
    
    'determine the last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'add a header for ticker symbol
    Cells(1, 9).Value = "Ticker"

    'loop through the tickers
    For i = 2 To lastrow
    
        
        'create a condition that compare the tickers of next row is the same as the last row one or not
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            
            'store the ticker symbol
            Cells(ticker_row, 9).Value = ticker
            
            'switch to the next ticker symbol
            ticker = Cells(i + 1, 1).Value
            
            'let the next ticker symbol can be store in the next row
            ticker_row = ticker_row + 1
         
        End If
        
        
    Next i
    
End Sub

