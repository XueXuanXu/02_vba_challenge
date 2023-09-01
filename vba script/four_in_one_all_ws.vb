Sub four_in_one_all_ws()

    'loop through all the sheets
    For Each ws In Worksheets
        
        
        'create a variable that store the ticker name
        Dim ticker As String
        ticker = ws.Cells(2, 1).Value
    
        'create variables that store the opening price, closing price and yearly price change
        Dim open_p As Double
        open_p = ws.Cells(2, 3).Value
        Dim close_p As Double
        Dim yearly_change As Double

        'create a variable that store the percent change
        Dim percent_change As Double

        'create a variable that store the total stock volume
        Dim total_stock_volume As Double
        total_stock_volume = 0

        'create a variable that store the row number to store in
        Dim store_row As Integer
        store_row = 2
    
        'determine the last row
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'add headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        'loop through the tickers
        For i = 2 To lastrow
    
        
            'create a condition that compare the tickers of next row is the same as the last row one or not
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            
                'store the ticker symbol
                ws.Cells(store_row, 9).Value = ticker
            
                'switch to the next ticker symbol
                ticker = ws.Cells(i + 1, 1).Value

                'store the closing price
                close_p = ws.Cells(i, 6).Value
            
                'store the yearly price change
                yearly_change = close_p - open_p
            
                'insert yearly price change to the table
                ws.Cells(store_row, 10).Value = yearly_change
  
                'calculate the percent change of price through the year
                percent_change = yearly_change / open_p
            
                'insert percent change to the table
                ws.Cells(store_row, 11).Value = FormatPercent(percent_change, 2)

                'add up the total stock volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

                'store the total stock volume
                ws.Cells(store_row, 12).Value = total_stock_volume
                ws.Cells(store_row, 12).NumberFormat = "General"

                'restart the total stock volume
                total_stock_volume = 0

                'store the opening price for the next ticker
                open_p = ws.Cells(i + 1, 3).Value

                'conditional formatting that will highlight positive change in green and negative change in red
                If (yearly_change > 0) Then
                    ws.Cells(store_row, 10).Interior.ColorIndex = 4
                    ws.Cells(store_row, 11).Interior.ColorIndex = 4
                ElseIf (yearly_change < 0) Then
                    ws.Cells(store_row, 10).Interior.ColorIndex = 3
                    ws.Cells(store_row, 11).Interior.ColorIndex = 3
            
                End If
            
                'let the next ticker symbol can be store in the next row
                store_row = store_row + 1
         
            ElseIf (ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value) Then

                'add up the total stock volume
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

            End If
        
        
        Next i
    
        
    
    Next ws
    
End Sub
