Sub four_in_one()

    'create a variable that store the ticker name
    Dim ticker As String
    ticker = Cells(2, 1).Value
    
    'create variables that store the opening price, closing price and yearly price change
    Dim open_p As Double
    open_p = Cells(2, 3).Value
    Dim close_p As Double
    Dim yearly_change As Double

    'create a variable that store the percent change
    dim percent_change as Double

    'create a variable that store the total stock volume
    Dim total_stock_volume as Double
    total_stock_volume = 0

    'create a variable that store the row number to store in
    Dim store_row As Integer
    store_row = 2
    
    'determine the last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'add headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    'loop through the tickers
    For i = 2 To lastrow
    
        
        'create a condition that compare the tickers of next row is the same as the last row one or not
        If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
            
            'store the ticker symbol
            Cells(store_row, 9).Value = ticker
            
            'switch to the next ticker symbol
            ticker = Cells(i + 1, 1).Value

            'store the closing price
            close_p = Cells(i, 6).Value
            
            'store the yearly price change
            yearly_change = close_p - open_p
            
            'insert yearly price change to the table
            Cells(store_row, 10).Value = yearly_change
  
            'calculate the percent change of price through the year
            percent_change = yearly_change / open_p
            
            'insert percent change to the table
            Cells(store_row, 11).Value = formatpercent(percent_change, 2)

            'add up the total stock volume
            total_stock_volume = total_stock_volume + Cells(i,7).Value

            'store the total stock volume
            Cells(store_row, 12).Value = total_stock_volume
            Cells(store_row, 12).numberformat ="General"

            'restart the total stock volume
            total_stock_volume =0

            'store the opening price for the next ticker
            open_p = Cells(i+1, 3).Value

            'conditional formatting that will highlight positive change in green and negative change in red
            If (yearly_change >0) Then
                Cells(store_row,10).interior.colorindex =4
                Cells(store_row,11).interior.colorindex =4
            Elseif (yearly_change < 0) Then
                Cells(store_row,10).interior.colorindex =3
                Cells(store_row,11).interior.colorindex =3
            
            End If
            
            'let the next ticker symbol can be store in the next row
            store_row = store_row + 1
         
        Elseif (Cells(i, 1).Value = Cells(i + 1, 1).Value) Then

            'add up the total stock volume
            total_stock_volume = total_stock_volume + Cells(i,7).Value

        End If
        
        
    Next i
    

End Sub