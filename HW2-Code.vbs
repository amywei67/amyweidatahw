Sub stock()

Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim ticker_counter As Long
ticker_counter = 0
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Total Volume"
Cells(1, 11).Value = "Opening Price"
Cells(1, 12).Value = "Closing Price"
Cells(1, 13).Value = "Percent Change"
Cells(1,14).value = "Yearly Change"

Dim yearly_change as double

Dim total_volume As Double
total_volume = 0

Dim opening_value, closing_value As Double
opening_value = Cells(2, 3).Value

Dim i As Double

For i = 2 To lastrow

total_volume = total_volume + Cells(i, 7).Value

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ticker_counter = ticker_counter + 1
        
        closing_value = Cells(i, 6).Value
        
        Cells(ticker_counter + 1, 9).Value = Cells(i, 1).Value
        Cells(ticker_counter + 1, 10).Value = total_volume
        Cells(ticker_counter + 1, 11).Value = opening_value
        Cells(ticker_counter + 1, 12).Value = closing_value
        

        If opening_value = 0 Then
            Cells(ticker_counter + 1, 13).Value = 0
        Else
            Cells(ticker_counter + 1, 13).Value = (closing_value - opening_value) / opening_value
        End If
        
        yearly_change = closing_value-opening_value
        Cells(ticker_counter+1,14).value = yearly_change
        
        Cells(ticker_counter + 1, 13).NumberFormat = "0.00%"
        
    
        If cells(ticker_counter+1,14).value >0 Then
        cells(ticker_counter+1,14).interior.colorindex=4
        Else
        cells(ticker_counter+1,14).interior.colorindex=3
        End if
        
        total_volume = 0
        opening_value = Cells(i + 1, 3).Value

    End If

Next i

'Hard solution tasks
Dim lastrow_percentage As Long
lastrow_percentage = Cells(Rows.Count, 9).End(xlUp).Row

Dim maximum_percentage, minimum_percentage, current_percentage, current_volume, maximum_volume As Double

maximum_percentage = 0
minimum_percentage = 0
maximum_volume = 0

Dim maximum_cell, minimum_cell, greatest_volume_cell As String


For i = 2 To lastrow_percentage

    current_percentage = Cells(i, 13).Value
    current_volume = Cells(i, 10).Value

    If current_percentage > maximum_percentage Then
        maximum_percentage = current_percentage
        maximum_cell = Cells(i, 9).Value
    ElseIf current_percentage < minimum_percentage Then
        minimum_percentage = current_percentage
        minimum_cell = Cells(i, 9).Value
    End If
    
    If current_volume > maximum_volume Then
        maximum_volume = current_volume
        greatest_volume_cell = Cells(i, 9).Value
    End If
    
Next i

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Cells(2, 16).Value = maximum_cell
Cells(3, 16).Value = minimum_cell
Cells(4, 16).Value = greatest_volume_cell

Cells(2, 17).Value = maximum_percentage
Cells(3, 17).Value = minimum_percentage
Cells(4, 17).Value = maximum_volume

Cells(2, 17).NumberFormat = "0.00%"
Cells(3, 17).NumberFormat = "0.00%"
End Sub