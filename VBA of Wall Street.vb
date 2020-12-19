Sub stocks()

For Each ws In Worksheets

'Declaration of variables'
    Dim lastrow As Long
    Dim ticker As String
    Dim total_vol As Double
    Dim open_value As Double
    Dim close_value As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim output_row As Integer

    total_vol = 0
    output_row = 2

'Column header formatting'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"


    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

    For i = 2 To lastrow

    'comparison of ticker symbols'
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
        ticker = ws.Cells(i, 1).Value
    'total volume counter
        total_vol = total_vol + ws.Cells(i, 7).Value

        ws.Range("I" & output_row).Value = ticker
        ws.Range("L" & output_row).Value = total_vol

'reset volume counter
    total_vol = 0

    close_value = ws.Cells(i, 6)

        If open_value = 0 Then
            yearly_change = 0
            percent_change = 0
        Else
            yearly_change = close_value - open_value
            percent_change = (close_value - open_value) / open_value
    End If
    
    ws.Range("J" & output_row).Value = yearly_change
    ws.Range("K" & output_row).Value = percent_change
    ws.Range("K" & output_row).Style = "Percent"
    ws.Range("K" & output_row).NumberFormat = "0.00%"
        
    'conditional formatting'
        If ws.Range("J" & output_row).Value >= 0 Then
            ws.Range("J" & output_row).Interior.ColorIndex = 4
        Else
            ws.Range("J" & output_row).Interior.ColorIndex = 3
        End If

output_row = output_row + 1

    ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
        open_value = ws.Cells(i, 3)

    Else
        total_vol = total_vol + ws.Cells(i, 7).Value


    End If


    Next i


'Attempting bonus section'
    For b = 2 To lastrow
                If ws.Range("K" & b).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & b).Value
                    ws.Range("P2").Value = ws.Range("I" & b).Value
                End If

                If ws.Range("K" & b).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & b).Value
                    ws.Range("P3").Value = ws.Range("I" & b).Value
                End If

                If ws.Range("L" & b).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & b).Value
                    ws.Range("P4").Value = ws.Range("I" & b).Value
                End If

            Next b
        ' Format Double To Include % Symbol And Two Decimal Places
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
        ' Format Table Columns To Auto Fit
        ws.Columns("I:Q").AutoFit

    Next ws

     
End Sub
