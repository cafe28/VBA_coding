Sub TickerCalculator()

    ' Loop over each worksheet
    For Each Worksheet In ActiveWorkbook.Sheets

        ' Set column headers
        Worksheet.Cells(1, 9).Value = "Ticker"
        Worksheet.Cells(1, 10).Value = "Yearly Change"
        Worksheet.Cells(1, 11).Value = "Percent Change"
        Worksheet.Cells(1, 12).Value = "Total Stock Volume"

        ' Define variables
        Dim ticker_name As String
        Dim first_ticker_row As Long
        Dim total_ticker_volume As LongLong

        ' Row index of where we should write the calculated data
        Dim current_ticker_row As Long
    
        ' Set default values
        current_ticker_row = 2
        first_ticker_row = 2
        total_ticker_volume = 0
        
        ' Loop over all the rows that have values
        For i = 2 To Worksheet.Cells(Rows.Count, 1).End(xlUp).Row
    
            ' Update current ticker name
            ticker_name = Worksheet.Cells(i, 1).Value

            ' Add the volume to the total ticker volume value
            total_ticker_volume = total_ticker_volume + Worksheet.Cells(i, 7).Value
        
            ' Check if the new row has a different ticker name
            ' If so, this means that we are on the last row of the current ticker
            If ticker_name <> Worksheet.Cells(i + 1, 1).Value Then

                ' Define variables
                Dim open_price As Double
                Dim close_price As Double
            
                ' Get the open price from the first row of the current ticker
                open_price = Worksheet.Cells(first_ticker_row, 3).Value

                ' Get the close price from the last row of the current ticker
                close_price = Worksheet.Cells(i, 6).Value
            
                ' Calculate the values and write them to appropriate cells
                Worksheet.Cells(current_ticker_row, 9).Value = ticker_name
                Worksheet.Cells(current_ticker_row, 10).Value = close_price - open_price
                Worksheet.Cells(current_ticker_row, 11).Value = (close_price - open_price) / open_price
                Worksheet.Cells(current_ticker_row, 12).Value = total_ticker_volume
            
                ' Move the row index one cell down
                current_ticker_row = current_ticker_row + 1

                ' Set the first row as the new cell
                first_ticker_row = i + 1

                ' Reset the total volume value
                total_ticker_volume = 0

            End If

        Next i
        
        ' Set column and row headers
        Worksheet.Cells(1, 16).Value = "Ticker"
        Worksheet.Cells(1, 17).Value = "Value"
        Worksheet.Cells(2, 15).Value = "Greatest % Increase"
        Worksheet.Cells(3, 15).Value = "Greatest % Decrease"
        Worksheet.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Define variables
        Dim greatest_increase_ticker_name As String
        Dim greatest_decrease_ticker_name As String
        Dim greatest_volume_ticker_name As String
        
        Dim greatest_increase_value As Double
        Dim greatest_decrease_value As Double
        Dim greatest_volume_value As LongLong
        
        ' Set default values
        greatest_increase_value = Worksheet.Cells(2, 11).Value
        greatest_decrease_value = Worksheet.Cells(2, 11).Value
        greatest_volume_value = Worksheet.Cells(2, 12).Value
        
        ' Loop over all rows of the data generated in the previous loop
        For j = 2 To Worksheet.Cells(Rows.Count, 9).End(xlUp).Row
        
            ' Define variables
            Dim current_ticker_name As String

            ' Remember the current ticker name
            current_ticker_name = Worksheet.Cells(j, 9).Value
        
            ' Compare the stored greatest increase value to the value in the current row
            ' If it's larger, store it instead
            If Worksheet.Cells(j, 11).Value > greatest_increase_value Then
                
                greatest_increase_value = Worksheet.Cells(j, 11).Value
                greatest_increase_ticker_name = current_ticker_name
                
            End If
            
            ' Compare the stored greatest decrease value to the value in the current row
            ' If it's smaller, store it instead
            If Worksheet.Cells(j, 11).Value < greatest_decrease_value Then
                
                greatest_decrease_value = Worksheet.Cells(j, 11).Value
                greatest_decrease_ticker_name = current_ticker_name
                
            End If
            
            ' Compare the stored greatest volume value to the value in the current row
            ' If it's larger, store it instead
            If Worksheet.Cells(j, 12).Value > greatest_volume_value Then
                
                greatest_volume_value = Worksheet.Cells(j, 12).Value
                greatest_volume_ticker_name = current_ticker_name
                
            End If
        
        Next j
        
        ' Write the values to appropriate cells
        Worksheet.Cells(2, 16).Value = greatest_increase_ticker_name
        Worksheet.Cells(3, 16).Value = greatest_decrease_ticker_name
        Worksheet.Cells(4, 16).Value = greatest_volume_ticker_name
        
        Worksheet.Cells(2, 17).Value = greatest_increase_value
        Worksheet.Cells(3, 17).Value = greatest_decrease_value
        Worksheet.Cells(4, 17).Value = greatest_volume_value
        
    Next Worksheet

End Sub
