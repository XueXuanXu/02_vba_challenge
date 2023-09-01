Sub percent_change():

    'create variables that store the opening price, closing price and yearly price change
    Dim open_p As Double
    open_p = Cells(2, 3).Value
    Dim close_p As Double
    Dim yearly_change As Double
    
    'create a variable that store the percent change
    dim percent_change as Double

    
    'create a variable that store the row number that percent change should be store in
    Dim percent_change_row As Integer
    percent_change_row = 2
    
    'determine the last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'add a header for percent change
    Cells(1, 11).Value = "Percent Change"
    
    
    'loop through the rows
    For i = 2 To lastrow
    
        
        'create a condition that compare the tickers of next row is the same as the last row one or not
        If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
            
            'store the closing price
            close_p = Cells(i, 6).Value
            
            'store the yearly price change
            yearly_change = close_p - open_p

            'calculate the percent change of price through the year
            percent_change = yearly_change / open_p
            
            'insert percent change to the table
            Cells(percent_change_row, 11).Value = formatpercent(percent_change, 2)

            'conditional formatting that will highlight positive change in green and negative change in red
            If (percent_change >0) Then
                Cells(percent_change_row,11).interior.colorindex =4
            Elseif (percent_change < 0) Then
                Cells(percent_change_row,11).interior.colorindex =3
            
            End If
            
            'let the next price change can be store in the next row
            percent_change_row = percent_change_row + 1

            'store the opening price for the next ticker
            open_p = Cells(i+1, 3).Value
         
         
        End If
               
    Next i


End Sub