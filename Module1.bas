Attribute VB_Name = "Module1"
Sub Calculate_Stock_Stats():
Dim ticker As String
Dim number_tickers As Integer
Dim lastRowState As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease As Double
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume As Double
Dim greatest_stock_volume_ticker As String


For Each ws In Worksheets

    
    ws.Activate

    
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    
    For i = 2 To lastRowState

        
        ticker = Cells(i, 1).Value
        
        
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        
        If Cells(i + 1, 1).Value <> ticker Then
            
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            
            closing_price = Cells(i, 6)
            
            
            yearly_change = closing_price - opening_price
            
            
            Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' If yearly change value is greater than 0, shade cell green.
            If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ' If yearly change value is less than 0, shade cell red.
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            ' If yearly change value is 0, shade cell yellow.
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            
            
            opening_price = 0
            
            
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            
            total_stock_volume = 0
        End If
        
    Next i
    
    ' Add section to display greatest percent increase, greatest percent decrease, and greatest total volume for each year.
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    '
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    
    For i = 2 To lastRowState
    
        ' Find the ticker with the greatest percent increase.
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        ' Find the ticker with the greatest percent decrease.
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        ' Find the ticker with the greatest stock volume.
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Add the values for greatest percent increase, decrease, and stock volume to each worksheet.
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws


End Sub
