Sub total_stock_volume():


    'create a variable that store the total stock volume
    Dim total_stock_volume as Double
    total_stock_volume = 0
    
    
    'create a variable that store the row number that total stock volume should be store in
    Dim volume_row As Integer
    volume_row = 2
    
    'determine the last row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'add a header for total stock volume
    Cells(1, 12).Value = "Total Stock Volume"

    'loop through the rows
    For i = 2 To lastrow
    
        
        'create a condition that compare the tickers of next row is the same as the last row one or not
        If (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
            
            'add up the total stock volume
            total_stock_volume = total_stock_volume + Cells(i,7).Value

            'store the total stock volume
            Cells(volume_row, 12).Value = total_stock_volume
            Cells(volume_row, 12).numberformat ="General"
            
            'let the next total stock volume can be store in the next row
            volume_row = volume_row + 1

            'restart the total stock volume
            total_stock_volume =0
         
        Elseif (Cells(i, 1).Value = Cells(i + 1, 1).Value) Then

            'add up the total stock volume
            total_stock_volume = total_stock_volume + Cells(i,7).Value

        End If
        
        
    Next i
    
End Sub

