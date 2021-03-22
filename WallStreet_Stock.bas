Attribute VB_Name = "WallStreet_Stock"
Sub Stock_AllYears()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Stock_OneYear
    Next
    Application.ScreenUpdating = True
End Sub
Sub Stock_OneYear()

    Dim i, j, k, m, n As Long
    Dim Total_Stock_Volume As LongLong
    Dim opening_price As Variant
    Dim closing_price As Variant
    Dim YearlyChange As Variant
    Dim Last_Row As Long
    Dim Greatest_Increase As Variant
    Dim Greatest_Decrease As Variant
    Dim Greatest_Total_Volume As Variant
 
         
    
 
    Total_Stock_Volume = 0
    YearlyChange = 0
    j = 2
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
 
    
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
   

    For i = 2 To Last_Row
        
         Total_Stock_Volume = Total_Stock_Volume + CLng(Cells(i, 7).Value)
    
        
        If CStr(Cells(i, 1).Value) <> CStr(Cells(i - 1, 1).Value) Then
            opening_price = CDec(Cells(i, 3).Value)
        End If
        
        If CStr(Cells(i, 1).Value) <> CStr(Cells(i + 1, 1).Value) Then
            
            closing_price = CDec(Cells(i, 6).Value)
            YearlyChange = closing_price - opening_price
            Cells(j, 9).Value = CStr(Cells(i, 1).Value)
            Cells(j, 10).Value = YearlyChange
            
            If YearlyChange >= 0 Then
                Cells(j, 10).Interior.Color = RGB(0, 255, 0)
            Else
                Cells(j, 10).Interior.Color = RGB(255, 0, 0)
            End If
            If opening_price Then
                Cells(j, 11).Value = (YearlyChange / opening_price)
            Else
                Cells(j, 11).Value = 0
            End If
            Cells(j, 12).Value = Total_Stock_Volume
            j = j + 1
            Total_Stock_Volume = 0
            opening_price = 0
            closing_price = 0
         
        End If
       
    Next i
    
    Columns("K").NumberFormat = "0.00%"
    
      
   
    
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total_Volume = 0
    k = 0
    m = 0
    n = 0
    
    Last_Row = Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To Last_Row - 1
        
        If Greatest_Increase < Cells(i, 11).Value Then
            Greatest_Increase = Cells(i, 11).Value
            k = i
        End If
        
        If Greatest_Decrease > Cells(i, 11).Value Then
            Greatest_Decrease = Cells(i, 11).Value
            m = i
        End If

        If Greatest_Total_Volume < Cells(i, 12).Value Then
            Greatest_Total_Volume = Cells(i, 12).Value
            n = i
        End If
    Next i
     
    Cells(2, 15).Value = "Greatest % increase"
    Cells(3, 15).Value = "Greatest % decrease"
    Cells(4, 15).Value = "Greatest total volume"
    
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    
    Cells(2, 16).Value = Cells(k, 9).Value
    Cells(2, 17).Value = Cells(k, 11).Value
    
    Cells(3, 16).Value = Cells(m, 9).Value
    Cells(3, 17).Value = Cells(m, 11).Value

    Cells(4, 16).Value = Cells(n, 9).Value
    Cells(4, 17).Value = Cells(n, 12).Value
    
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    Columns("I:L").Columns.AutoFit
    Columns("O:Q").Columns.AutoFit
    
End Sub



