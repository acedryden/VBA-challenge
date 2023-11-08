Attribute VB_Name = "Module1"
Sub LoopTest():
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets


'add headings to new columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

'set initial variables
        Dim ticker As String
        Dim yearly_change As Double
        Dim total_stock As LongLong
        total_stock = 0
        Dim last_closing_price As Double
        Dim first_opening_price As Double
        Dim Summary_table_Row As Double
        Summary_table_Row = 2
        Counter = 0

'Update Ticker & Total Stock Volume Column
        first_opening_price = ws.Cells(2, 3).Value
        For i = 2 To 760000
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                total_stock = total_stock + Cells(i, 7).Value
                last_closing_price = ws.Cells(i, 6).Value
                yearly_change = (last_closing_price - first_opening_price)
                percent_change = ((last_closing_price - first_opening_price) / first_opening_price)
                ws.Range("I" & Summary_table_Row).Value = ticker
                ws.Range("L" & Summary_table_Row).Value = total_stock
                ws.Range("J" & Summary_table_Row).Value = yearly_change
                ws.Range("K" & Summary_table_Row).Value = FormatPercent(percent_change)
                Summary_table_Row = Summary_table_Row + 1
                total_stock = 0
                first_opening_price = ws.Cells(i + 1, 3).Value
             Else
                total_stock = total_stock + ws.Cells(i, 7).Value
        
        End If
        Next i

'Conditional Formatting:
        For i = 2 To 760000
            If ws.Cells(i, 10) >= 0 Then
               ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i

        For i = 2 To 760000
            If ws.Cells(i, 11) >= 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 11).Interior.ColorIndex = 3
            End If
        Next i

'Summary Table
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volumne"

'Greatest Total Volume
        Dim greatest_vol As LongLong
        Dim vol_ticker As String
        greatest_vol = 0
        For i = 2 To 760000
            If ws.Cells(i + 1, 12).Value > greatest_vol Then
                greatest_vol = ws.Cells(i + 1, 12).Value
                vol_ticker = ws.Cells(i, 9).Value
                ws.Range("Q4").Value = greatest_vol
                ws.Range("P4").Value = vol_ticker
            Else
        
            End If
        Next i

'Greatest % Increase
        Dim greatest_inc As Double
        Dim inc_ticker As String
        greatest_inc = 0
        For i = 2 To 760000
            If ws.Cells(i + 1, 11).Value > greatest_inc Then
                greatest_inc = ws.Cells(i + 1, 11).Value
                inc_ticker = ws.Cells(i + 1, 9).Value
                ws.Range("Q2").Value = FormatPercent(greatest_inc)
                ws.Range("P2").Value = inc_ticker
            Else
            End If
        Next i
'Greatest % Decrease
        Dim greatest_dec As Double
        Dim dec_ticker As String
        greatest_dec = 0
        For i = 2 To 760000
            If ws.Cells(i + 1, 11).Value < greatest_dec Then
             greatest_dec = ws.Cells(i + 1, 11).Value
             dec_ticker = ws.Cells(i + 1, 9).Value
             ws.Range("Q3").Value = FormatPercent(greatest_dec)
             ws.Range("P3").Value = dec_ticker
            Else
            
         End If
         Next i
    Next ws

End Sub
