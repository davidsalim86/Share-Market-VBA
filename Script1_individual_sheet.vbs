Sub VBA_script1()

'cells' heading
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

'Variable setting
Dim stock_name As String
Dim last_row As Long
Dim last_row2 As Long
Dim total_volume As Variant
Dim summary_row As Long
Dim yearly_open As Double
Dim yearly_close As Double
Dim yearly_change As Double
Dim open_amount As Long
Dim percent_change As Double
Dim max_increase As Double
Dim max_decrease As Double
Dim max_vol As Variant

last_row = Cells(Rows.Count, 1).End(xlUp).Row

'initial value
total_volume = 0
summary_row = 2
open_amount = 2

'looping through data
    For i = 2 To last_row
       
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            stock_name = Cells(i, 1).Value
            yearly_open = Cells(open_amount, 3).Value
            yearly_close = Cells(i, 6).Value
            total_volume = total_volume + Cells(i, 7).Value
            
            
            If yearly_open <> 0 Then
                percent_change = yearly_close / yearly_open - 1
                yearly_change = yearly_close - yearly_open
            End If
            
            If total_volume = 0 Then
                percent_change = 0
                yearly_change = 0
            End If
                              
            Range("i" & summary_row).Value = stock_name
            Range("J" & summary_row).Value = yearly_change
            Range("K" & summary_row).Value = percent_change
            Range("l" & summary_row).Value = total_volume
            
            If yearly_change >= 0 Then
                Range("J" & summary_row).Interior.ColorIndex = 4
            Else
                Range("J" & summary_row).Interior.ColorIndex = 3
            End If

            Range("K" & summary_row) = Format(Range("K" & summary_row), "percent")
            
            summary_row = summary_row + 1
            open_amount = i + 1
            total_volume = 0
        
        Else
            total_volume = total_volume + Cells(i, 7).Value
        End If
        
    Next i

last_row2 = Cells(Rows.Count, 9).End(xlUp).Row

'Greatest % increase
max_increase = WorksheetFunction.Max(Range("k2:K" & last_row2))
max_increase_tick = WorksheetFunction.Match(max_increase, Range("k2:k" & last_row2), 0)
Cells(2, 16).Value = Cells(max_increase_tick + 1, 9)
Cells(2, 17).Value = max_increase

'Greatest % decrease
max_decrease = WorksheetFunction.Min(Range("k2:K" & last_row2))
max_decrease_tick = WorksheetFunction.Match(max_decrease, Range("k2:k" & last_row2), 0)
Cells(3, 16).Value = Cells(max_decrease_tick + 1, 9)
Cells(3, 17).Value = max_decrease

'Greatest volume
max_vol = WorksheetFunction.Max(Range("l2:l" & last_row2))
max_vol_tick = WorksheetFunction.Match(max_vol, Range("l2:l" & last_row2), 0)
Cells(4, 16).Value = Cells(max_vol_tick + 1, 9)
Cells(4, 17).Value = max_vol

'format
Cells(2, 17) = Format(Cells(2, 17), "percent")
Cells(3, 17) = Format(Cells(3, 17), "percent")
Columns("I:q").AutoFit

End Sub



