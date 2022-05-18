Sub VBA_script2()

For Each ws In Worksheets

'cells' heading
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

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

last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

'initial value
total_volume = 0
summary_row = 2
open_amount = 2

'looping through data
    For i = 2 To last_row
       
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            stock_name = ws.Cells(i, 1).Value
            yearly_open = ws.Cells(open_amount, 3).Value
            yearly_close = ws.Cells(i, 6).Value
            total_volume = total_volume + ws.Cells(i, 7).Value
            
            
            If yearly_open <> 0 Then
                percent_change = yearly_close / yearly_open - 1
                yearly_change = yearly_close - yearly_open
            End If
            
            If total_volume = 0 Then
                percent_change = 0
                yearly_change = 0
            End If
                 
            ws.Range("i" & summary_row).Value = stock_name
            ws.Range("J" & summary_row).Value = yearly_change
            ws.Range("K" & summary_row).Value = percent_change
            ws.Range("l" & summary_row).Value = total_volume
            
            If yearly_change >= 0 Then
                ws.Range("J" & summary_row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & summary_row).Interior.ColorIndex = 3
            End If

            ws.Range("K" & summary_row) = Format(ws.Range("K" & summary_row), "percent")
            
            summary_row = summary_row + 1
            open_amount = i + 1
            total_volume = 0
        
        Else
            total_volume = total_volume + ws.Cells(i, 7).Value
        End If
        
    Next i

last_row2 = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Greatest % increase
max_increase = WorksheetFunction.Max(ws.Range("k2:K" & last_row2))
max_increase_tick = WorksheetFunction.Match(max_increase, ws.Range("k2:k" & last_row2), 0)
ws.Cells(2, 16).Value = ws.Cells(max_increase_tick + 1, 9)
ws.Cells(2, 17).Value = max_increase

'Greatest % decrease
max_decrease = WorksheetFunction.Min(ws.Range("k2:K" & last_row2))
max_decrease_tick = WorksheetFunction.Match(max_decrease, ws.Range("k2:k" & last_row2), 0)
ws.Cells(3, 16).Value = ws.Cells(max_decrease_tick + 1, 9)
ws.Cells(3, 17).Value = max_decrease

'Greatest volume
max_vol = WorksheetFunction.Max(ws.Range("l2:l" & last_row2))
max_vol_tick = WorksheetFunction.Match(max_vol, ws.Range("l2:l" & last_row2), 0)
ws.Cells(4, 16).Value = ws.Cells(max_vol_tick + 1, 9)
ws.Cells(4, 17).Value = max_vol

'format
ws.Cells(2, 17) = Format(ws.Cells(2, 17), "percent")
ws.Cells(3, 17) = Format(ws.Cells(3, 17), "percent")
ws.Columns("I:q").AutoFit

Next ws

End Sub
