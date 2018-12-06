Sub ticker_Total()

volume_subtotal = 0
current_unique_row_number = 2
Cells(current_unique_row_number, 8).Value = Cells(2, 1).Value
Cells(1, 10).Value = "Ticker"
Cells(1, 11).Value = "Total Volume"
last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
For Row = 2 To last_row

current_ticker = Cells(Row, 1).Value
next_ticker = Cells(Row + 1, 1).Value
current_volume = Cells(Row, 7).Value




If current_ticker = next_ticker Then 'add up the charge
    volume_subtotal = volume_subtotal + current_volume
    
Else 'start a new total
    volume_subtotal = volume_subtotal + current_volume
    current_unique_row_number = current_unique_row_number + 1
    Cells(current_unique_row_number - 1, 10).Value = current_ticker
    Cells(current_unique_row_number - 1, 11).Value = volume_subtotal
        volume_subtotal = 0


End If

Next Row

End Sub