Sub yearly_change():

    'create variables that store the opening price, closing price and yearly price change
    Dim open_p As Double
    open_p = Cells(2, 3).Value
    Dim close_p As Double
    Dim yearly_change As Double
    
    
    'create a variable that store the row number that price change should be store in
    Dim price_change_row As Integer
    price_change_row = 2
    
    'determine the last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'add a header for price change
    Cells(1, 10).Value = "Yearly Change"
    
    
    'loop through the rows
    For i = 2 To lastrow
    
        
        'create a condition that compare the tickers of next row is the same as the last row one or not
        If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
            
            'store the closing price
            close_p = Cells(i, 6).Value
            
            'store the yearly price change
            yearly_change = close_p - open_p
            
            'insert yearly price change to the table
            Cells(price_change_row, 10).Value = yearly_change

            'conditional formatting that will highlight positive change in green and negative change in red
            If (yearly_change >0) Then
                Cells(price_change_row,10).interior.colorindex =4
            Elseif (yearly_change < 0) Then
                Cells(price_change_row,10).interior.colorindex =3
            
            End If
            
            
            'let the next price change can be store in the next row
            price_change_row = price_change_row + 1

            'store the opening price for the next ticker
            open_p = Cells(i+1, 3).Value


         
         
        End If
        
        
    Next i
    



End Sub
