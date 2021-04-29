Sub Stock_Data()

Dim ws As Worksheet

    For Each ws In Worksheets

    Dim sumtabrow As Long
    Dim totvol As Double
    Dim open_price As Double
    Dim close_price As Double
    Dim yr_change As Double
    Dim tick As String
    Dim i As Long
    Dim percent_change As Double
    Dim lrow As Long
    
    lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    sumtabrow = 2
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change "
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    
        For i = 2 To lrow


            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1) Then
    
                open_price = ws.Cells(i, 3).Value
    
            End If
    
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    
                tick = ws.Cells(i, 1).Value
        
                totvol = totvol + ws.Cells(i, 7).Value
        
                close_price = ws.Cells(i, 6).Value
        
                yearly_change = close_price - open_price
    
                If open_price <> 0 Then
    
                    percent_change = yearly_change / open_price
    
                Else
    
                    percent_change = 0
    
                End If
    
                If yearly_change > 0 Then
    
                    ws.Cells(sumtabrow, 10).Interior.ColorIndex = 4
    
                Else
    
                    ws.Cells(sumtabrow, 10).Interior.ColorIndex = 3
    
                End If
    
    
                ws.Cells(sumtabrow, 9).Value = tick
        
                ws.Cells(sumtabrow, 12).Value = totvol
        
                ws.Cells(sumtabrow, 10).Value = yearly_change
        
                ws.Cells(sumtabrow, 11).Value = percent_change
        
                ws.Cells(sumtabrow, 11).NumberFormat = "0.00%"
        
                sumtabrow = sumtabrow + 1
        
                totvol = 0
        
            Else
    
                totvol = totvol + ws.Cells(i, 7).Value
    
            End If


        Next i
    
    
        Dim high_change, low_change As Double
    
        Dim high_vol, lastrow As Long
    
        Dim low_change_val, high_change_val, high_vol_val As String
    
        lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
        high_change = ws.Cells(2, 11).Value
    
        low_change = ws.Cells(2, 11).Value
    
        high_vol = ws.Cells(2, 12).Value
    
        For n = 2 To lastrow
    
            If ws.Cells(n, 11).Value > high_change Then
        
                high_change = ws.Cells(n, 11).Value
        
                high_change_val = ws.Cells(n, 9).Value
 
            End If
    
        
            If ws.Cells(n, 11).Value < low_change Then
    
                low_change = ws.Cells(n, 11).Value
        
                low_change_val = ws.Cells(n, 9).Value
   
            End If
    
    
            If ws.Cells(n, 12).Value > high_vol Then
    
                high_vol = ws.Cells(n, 12).Value
        
                high_vol_val = ws.Cells(n, 9).Value
 
            End If
            
            
            ws.Range("P2").Value = high_change
        
            ws.Range("Q2").Value = high_change_val
        
            ws.Range("P2").NumberFormat = "0.00%"
            
            ws.Range("P3").Value = low_change
        
            ws.Range("Q3").Value = low_change_val
        
            ws.Range("P3").NumberFormat = "0.00%"
            
            ws.Range("P4").Value = high_vol
        
            ws.Range("Q4").Value = high_vol_val
            

        Next n
        
        ws.Columns("I:L").EntireColumn.AutoFit
        
        ws.Columns("O:Q").EntireColumn.AutoFit
        
    Next ws
    
     
End Sub