Attribute VB_Name = "Module1"
Sub Stocks()

    Dim ticker As String
    Dim opening As Double, closing As Double
    Dim quarterly As Double, Percent_Change As Double, total_stock As Double
    Dim previous_price As Long, ticker_row As Long
    Dim ws As Worksheet
    Dim end_ticker As Long
    Dim greatest_increase As Double, greatest_decrease As Double, greatest_volume As Double
    Dim percent_end As Long

    For Each ws In Worksheets
        ' Initialize headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        total_stock = 0
        previous_price = 2
        ticker_row = 2

        end_ticker = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Calculate stock metrics
        For a = 2 To end_ticker
            total_stock = total_stock + ws.Cells(a, 7).Value
            
            If ws.Cells(a + 1, 1).Value <> ws.Cells(a, 1).Value Then
                ticker = ws.Cells(a, 1).Value
                ws.Range("I" & ticker_row).Value = ticker
                ws.Range("L" & ticker_row).Value = total_stock
                total_stock = 0 ' Reset for next ticker
                
                opening = ws.Cells(previous_price, 3).Value
                closing = ws.Cells(a, 6).Value
                quarterly = closing - opening
                ws.Range("J" & ticker_row).Value = quarterly
                
                If opening <> 0 Then
                    Percent_Change = quarterly / opening
                Else
                    Percent_Change = 0
                End If
                
                ' Color the Quarterly Change cell based on the value
                If quarterly >= 0 Then
                    ws.Range("J" & ticker_row).Interior.ColorIndex = 4 ' Green for positive
                Else
                    ws.Range("J" & ticker_row).Interior.ColorIndex = 3 ' Red for negative
                End If
                
                ws.Range("K" & ticker_row).Value = Percent_Change
                ws.Range("K" & ticker_row).NumberFormat = "0.00%" ' Format as percentage
                
                previous_price = a + 1
                ticker_row = ticker_row + 1
            End If
        Next a

        ' Initialize summary section
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        greatest_increase = 0
        greatest_decrease = 0
        greatest_volume = 0
        
        percent_end = ws.Cells(Rows.Count, 11).End(xlUp).Row

        ' Find greatest metrics
        For a = 2 To percent_end
            If ws.Range("K" & a).Value > greatest_increase Then
                greatest_increase = ws.Range("K" & a).Value
                ws.Range("R2").Value = greatest_increase
                ws.Range("Q2").Value = ws.Range("I" & a).Value
            End If
            
            If ws.Range("K" & a).Value < greatest_decrease Then
                greatest_decrease = ws.Range("K" & a).Value
                ws.Range("R3").Value = greatest_decrease
                ws.Range("Q3").Value = ws.Range("I" & a).Value
            End If
            
            If ws.Range("L" & a).Value > greatest_volume Then
                greatest_volume = ws.Range("L" & a).Value
                ws.Range("R4").Value = greatest_volume
                ws.Range("Q4").Value = ws.Range("I" & a).Value
            End If
        Next a
        
        ' Format the summary results
        ws.Range("R2:R3").NumberFormat = "0.00%" ' Format percentage values
    Next ws

End Sub
