Sub ticker_info():

For Each ws In Worksheets

Dim ticker As String
Dim volume_total As Double
volume_total = 0

Dim open_total As Double
Dim open_final As Double
Dim close_final As Double
open_total = 0

Dim summary_row As Integer
summary_row = 2

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To last_row
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ticker = Cells(i, 1).Value
        volume_total = volume_total + Cells(i, 7).Value
        Cells(summary_row, 9).Value = ticker
        Cells(summary_row, 12).Value = volume_total
        
        close_final = Cells(i, 6).Value
        
        Cells(summary_row, 10).Value = close_final - open_final
        Cells(summary_row, 11).Value = (close_final - open_final) / open_final
        Cells(summary_row, 11).NumberFormat = "0.00%"
        
        summary_row = summary_row + 1
        volume_total = 0
        open_total = 0
        open_final = 0
        close_final = 0
        
        Else
        
        volume_total = volume_total + Cells(i, 7).Value
        
        If open_total = 0 Then
        
            open_total = open_total + Cells(i, 3).Value
            open_final = Cells(i, 3).Value
            
            Else
        
            open_total = open_total + Cells(i, 3).Value
            
            End If
            
        open_total = open_total + Cells(i, 3).Value
        
        
        
        End If
        
        If Cells(i, 10).Value > 0 Then
        
        Cells(i, 10).Interior.ColorIndex = 4
        
        ElseIf Cells(i, 10).Value < 0 Then
        
        Cells(i, 10).Interior.ColorIndex = 3
        
        End If
        
    Next i
    
Dim greatest_increase As Double
Dim increase_ticker As String
Dim greateset_decrease As Double
Dim decrease_ticker As String
Dim greatest_total_volume As Double
Dim greatest_volume_ticker As String

greatest_increase = 0
greatest_decrease = 0
greatest_total_volume = 0

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

    For i = 2 To last_row
    
        If Cells(i, 11).Value > greatest_increase Then
        
        greatest_increase = Cells(i, 11).Value
        increase_ticker = Cells(i, 9).Value
        Cells(2, 17).Value = greatest_increase
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(2, 16).Value = increase_ticker
        
        End If
        
        If Cells(i, 11).Value < greatest_decrease Then
        
        greatest_decrease = Cells(i, 11).Value
        decrease_ticker = Cells(i, 9).Value
        Cells(3, 17).Value = greatest_decrease
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(3, 16).Value = decrease_ticker
        
        End If
        
        If Cells(i, 12).Value > greatest_total_volume Then
        
        greatest_total_volume = Cells(i, 12).Value
        greatest_volume_ticker = Cells(i, 9).Value
        Cells(4, 17).Value = greatest_total_volume
        Cells(4, 16).Value = greatest_volume_ticker
        
        End If
        
    Next i
    
Next ws

End Sub


