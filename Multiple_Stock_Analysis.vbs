Sub Stock_Analysis()
    
    'Variable declarations
    Dim ticker As String
    Dim year_open As Double
    Dim year_close As Double
    Dim yearly_change As Double
    Dim total_stock_volume As Double
    Dim percent_change As Double
    Dim start_data As Integer
        
    Dim ws As Worksheet
    
    'Loop through all worksheets to execute the code all at once
    For Each ws In Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        Dim previous_row As Long
        Dim end_row As Long
        
        start_data = 2
        previous_row = 1
        total_stock_volume = 0
    
        end_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
        For i = 2 To end_row
    
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
                ticker = ws.Cells(i, 1).Value
                previous_row = previous_row + 1
                year_open = ws.Cells(previous_row, 3).Value
                year_close = ws.Cells(i, 6).Value
    
                For j = previous_row To i
                    total_stock_volume = total_stock_volume + ws.Cells(j, 7).Value
                Next j
    
                If year_open = 0 Then
                    percent_change = year_close
                Else
                    yearly_change = year_close - year_open
                    percent_change = yearly_change / year_open
                End If
    
                ws.Cells(start_data, 9).Value = ticker
                ws.Cells(start_data, 10).Value = yearly_change
                ws.Cells(start_data, 11).Value = percent_change
    
                ws.Cells(start_data, 11).NumberFormat = "0.00%"
                ws.Cells(start_data, 12).Value = total_stock_volume
    
                start_data = start_data + 1
    
                total_stock_volume = 0
                yearly_change = 0
                percent_change = 0
    
                previous_row = i
    
            End If
        Next i
    
        'Greatest summary table for row K
        kend_row = ws.Cells(Rows.Count, "K").End(xlUp).Row
        
        Increase = 0
        Decrease = 0
        Greatest = 0
    
        For k = 3 To kend_row
            
            last_k = k - 1
            current_k = ws.Cells(k, 11).Value
            previous_k = ws.Cells(last_k, 11).Value
            volume = ws.Cells(k, 12).Value
            previous_vol = ws.Cells(last_k, 12).Value
            
            'Greatest Increase
            If Increase > current_k And Increase > previous_k Then
                Increase = Increase
            ElseIf current_k > Increase And current_k > previous_k Then
                Increase = current_k
                increase_name = ws.Cells(k, 9).Value
            ElseIf previous_k > Increase And previous_k > current_k Then
                Increase = previous_k
                increase_name = ws.Cells(last_k, 9).Value
            End If

            'Greatest Decrease
            If Decrease < current_k And Decrease < previous_k Then
                Decrease = Decrease
            ElseIf current_k < Increase And current_k < previous_k Then
                Decrease = current_k
                decrease_name = ws.Cells(k, 9).Value
            ElseIf previous_k < Increase And previous_k < current_k Then
                Decrease = previous_k
                decrease_name = ws.Cells(last_k, 9).Value
            End If
            
            'Greatest Volume
            If Greatest > volume And Greatest > previous_vol Then
                Greatest = Greatest
            ElseIf volume > Greatest And volume > previous_vol Then
                Greatest = volume
                greatest_name = ws.Cells(k, 9).Value
            ElseIf previous_vol > Greatest And previous_vol > volume Then
                Greatest = previous_vol
                greatest_name = ws.Cells(last_k, 9).Value
            End If
        Next k
        
        'Fill Column names for greatest summary table
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
    
        'Fill Column values for greatest summary table
        ws.Range("O2").Value = increase_name
        ws.Range("O3").Value = decrease_name
        ws.Range("O4").Value = greatest_name
        ws.Range("P2").Value = Increase
        ws.Range("P3").Value = Decrease
        ws.Range("P4").Value = Greatest
    
        'Greatest increase and decrease in percentage format
        ws.Range("P2").NumberFormat = "0.00%"
        ws.Range("P3").NumberFormat = "0.00%"
    
        'Conditional formatting columns colors in row J
        jend_row = ws.Cells(Rows.Count, "J").End(xlUp).Row
    
        For j = 2 To jend_row
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
    Next ws
End Sub