Attribute VB_Name = "Module2"
Sub Quarterlystocks()



Dim open_price As Double
Dim close_price As Double
Dim quaterly_change As Double
Dim ticker_name As String
Dim percent_change As Double
Dim volume As Double
Dim row As Double
Dim column As Integer
Dim current_worksheet As Worksheet

    For Each current_worksheet In ActiveWorkbook.Worksheets
    
    current_worksheet.Activate
    
    last_row = current_worksheet.Cells(rows.Count, 1).End(xlUp).row
    
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
      
        
        
        volume = 0
        row = 2
        column = 1
       
    
        open_price = Cells(2, column + 2).Value
        
         
        For i = 2 To last_row
        
         
            If Cells(i + 1, column).Value <> Cells(i, column).Value Then
            
                
                ticker_name = Cells(i, column).Value
                    Cells(row, column + 8).Value = ticker_name
                
               
                close_price = Cells(i, column + 5).Value
                
              
                quaterly_change = close_price - open_price
                    Cells(row, column + 9).Value = quaterly_change
               
               
                percent_change = quaterly_change / open_price
                    Cells(row, column + 10).Value = percent_change
                    Cells(row, column + 10).NumberFormat = "0.00%"
               
               
              
                volume = volume + Cells(i, column + 6).Value
                    Cells(row, column + 11).Value = volume
                
                
                row = row + 1
                
                
                open_price = Cells(i + 1, column + 2)
                
                
                volume = 0
                
            Else
                volume = volume + Cells(i, column + 6).Value
            End If
        Next i
        
        
        
        quaterly_change_last_row = current_worksheet.Cells(rows.Count, 9).End(xlUp).row
        
        
        For j = 2 To quaterly_change_last_row
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
        
        For k = 2 To quaterly_change_last_row
            If Cells(k, 11).Value = Application.WorksheetFunction.Max(current_worksheet.Range("K2:K" & quaterly_change_last_row)) Then
                Cells(2, 16).Value = Cells(k, 9).Value
                Cells(2, 17).Value = Cells(k, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(current_worksheet.Range("K2:K" & quaterly_change_last_row)) Then
                Cells(3, 16).Value = Cells(k, 9).Value
                Cells(3, 17).Value = Cells(k, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(k, column + 11).Value = Application.WorksheetFunction.Max(current_worksheet.Range("L2:L" & quaterly_change_last_row)) Then
                Cells(4, 16).Value = Cells(k, 9).Value
                Cells(4, 17).Value = Cells(k, 12).Value
            End If
        Next k
        
        
    Next current_worksheet
    
End Sub
