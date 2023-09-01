Sub find_greatest_all_ws():

    'loop through all the sheets
    For Each ws In Worksheets

        'create variables to store the data
        Dim greatest_increase_ticker As String
        Dim greatest_decrease_ticker As String
        Dim greatest_total_volume_ticker As String
        Dim greatest_increase_value As Double
        Dim greatest_decrease_value As Double
        Dim greatest_total_volume_value As Double

        'store the initial data
        greatest_increase_ticker = ws.Cells(2, 9).Value
        greatest_decrease_ticker = ws.Cells(2, 9).Value
        greatest_total_volume_ticker = ws.Cells(2, 9).Value
        greatest_increase_value = ws.Cells(2, 11).Value
        greatest_decrease_value = Cells(2, 11).Value
        greatest_total_volume_value = ws.Cells(2, 12).Value

        'determine the last row
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

        'loop through the rows
        For i = 2 To lastrow

            'create condition for greatest increase
            If (ws.Cells(i, 11).Value > greatest_increase_value) Then
                greatest_increase_ticker = ws.Cells(i, 9).Value
                greatest_increase_value = ws.Cells(i, 11).Value
            End If

            'create condition for greatest decrease
            If (ws.Cells(i, 11).Value < greatest_decrease_value) Then
                greatest_decrease_ticker = ws.Cells(i, 9).Value
                greatest_decrease_value = ws.Cells(i, 11).Value
            End If

            'create condition for greatest total volume
            If (ws.Cells(i, 12).Value > greatest_total_volume_value) Then
                greatest_total_volume_ticker = ws.Cells(i, 9).Value
                greatest_total_volume_value = ws.Cells(i, 12).Value
            End If


        Next i

        'add headers
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

        'insert value to the table
        ws.Cells(2, 16).Value = greatest_increase_ticker
        ws.Cells(2, 17).Value = FormatPercent(greatest_increase_value, 2)
        ws.Cells(3, 16).Value = greatest_decrease_ticker
        ws.Cells(3, 17).Value = FormatPercent(greatest_decrease_value, 2)
        ws.Cells(4, 16).Value = greatest_total_volume_ticker
        ws.Cells(4, 17).Value = greatest_total_volume_value

    Next ws

End Sub
