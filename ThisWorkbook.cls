VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_analysis()
    Dim row_count As Long
    Dim company As String
    Dim volume_sum As Double
    Dim quarter_start As Double
    Dim quarter_end As Double
    Dim company_count As Integer
    Dim biggest_inc As Double
    Dim biggest_dec As Double
    Dim biggest_vol As Double
    
    company_count = 2
    
    row_count = Cells(Rows.Count, 1).End(xlUp).Row
    
    company = Cells(2, 1).Value
    quarter_start = Cells(2, 3).Value
    volume_sum = 0
    
    For i = 2 To row_count
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            quarter_end = Cells(i, 6).Value
            Cells(company_count, 9) = company
            Cells(company_count, 10) = quarter_end - quarter_start
            Cells(company_count, 11) = (quarter_end - quarter_start) / quarter_start
            volume_sum = volume_sum + Cells(i, 7).Value
            Cells(company_count, 12) = volume_sum
            
            company = Cells(i + 1, 1).Value
            quarter_start = Cells(i + 1, 3).Value
            
            volume_sum = 0
            company_count = company_count + 1
            
        Else
            
            volume_sum = volume_sum + Cells(i, 7).Value
            
        End If

    Next i
    
    biggest_inc = 0
    biggest_dec = 0
    biggest_vol = 0
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    For i = 2 To company_count
    
        If Cells(i, 11).Value > biggest_inc Then
            
            biggest_inc = Cells(i, 11).Value
            Cells(2, 17).Value = biggest_inc
            Cells(2, 16).Value = Cells(i, 9)
        
        End If
        
        If Cells(i, 11).Value < biggest_dec Then
            
            biggest_dec = Cells(i, 11).Value
            Cells(3, 17).Value = biggest_dec
            Cells(3, 16).Value = Cells(i, 9)
            
        End If
            
        If Cells(i, 12).Value > biggest_vol Then
        
            biggest_vol = Cells(i, 12).Value
            Cells(4, 17).Value = biggest_vol
            Cells(4, 16).Value = Cells(i, 9).Value
            
        End If
    
    Next i
    

End Sub
