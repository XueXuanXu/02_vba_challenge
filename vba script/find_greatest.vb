Sub find_greatest():

    'create variables to store the data
    dim greatest_increase_ticker as string
    dim greatest_decrease_ticker as string
    dim greatest_total_volume_ticker as string
    dim greatest_increase_value as double
    dim greatest_decrease_value as double
    dim greatest_total_volume_value as double

    'store the initial data
    greatest_increase_ticker = Cells(2,9).Value
    greatest_decrease_ticker = Cells(2,9).Value
    greatest_total_volume_ticker = Cells(2,9).Value
    greatest_increase_value = Cells(2,11).Value
    greatest_decrease_value = Cells(2,11).Value
    greatest_total_volume_value = Cells(2,12).Value

    'determine the last row
    lastrow = Cells(Rows.Count, 9).End(xlUp).Row

    'loop through the rows
    For i = 2 to lastrow

        'create condition for greatest increase
        If (Cells(i, 11).value > greatest_increase_value) then
            greatest_increase_ticker = Cells(i,9).Value
            greatest_increase_value = Cells(i,11).Value
        End If

        'create condition for greatest decrease
        If (Cells(i, 11).value < greatest_decrease_value) then
            greatest_decrease_ticker = Cells(i,9).Value
            greatest_decrease_value = Cells(i,11).Value
        End If

        'create condition for greatest total volume
        If (Cells(i, 12).value > greatest_total_volume_value) then
            greatest_total_volume_ticker = Cells(i,9).Value
            greatest_total_volume_value = Cells(i,12).Value
        End If


    Next i   

    'add headers
    Cells(1,16).Value ="Ticker"
    Cells(1,17).Value ="Value"
    Cells(2,15).Value ="Greatest % Increase"
    Cells(3,15).Value ="Greatest % Decrease"
    Cells(4,15).Value ="Greatest Total Volume"

    'insert value to the table
    Cells(2,16).Value = greatest_increase_ticker
    Cells(2,17).Value = formatpercent(greatest_increase_value,2)
    Cells(3,16).Value = greatest_decrease_ticker
    Cells(3,17).Value = formatpercent(greatest_decrease_value,2)
    Cells(4,16).Value = greatest_total_volume_ticker
    Cells(4,17).Value = greatest_total_volume_value

End Sub