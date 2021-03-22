Attribute VB_Name = "WallStreet_Stock"
Sub Stock_AllYears()
    
    'defines a worksheet variable
    Dim xSh As Worksheet
    
    'turns screen updating off to speed up the code. it prevents us to see what the macro is doing until it finishes
    Application.ScreenUpdating = False
    
    ' goes through each sheet in the workbook and runs "Stock_OneYear" macro on it
    For Each xSh In Worksheets
    
        xSh.Select
        Call Stock_OneYear
        
    Next
    
    
    'turns screen updating back on to show the latest change after macro run
    Application.ScreenUpdating = True

End Sub
Sub Stock_OneYear()

    'define variables
    Dim i, j, k, m, n As Long
    Dim Total_Stock_Volume As LongLong
    Dim opening_price As Variant
    Dim closing_price As Variant
    Dim YearlyChange As Variant
    Dim Last_Row As Long
    Dim Greatest_Increase As Variant
    Dim Greatest_Decrease As Variant
    Dim Greatest_Total_Volume As Variant
 
         
    
    ' initialize some of variables
    Total_Stock_Volume = 0
    YearlyChange = 0
    j = 2   'will be used for ticker data_set
    
    'Creat headers for Ticker data_set
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
 
    'find the last row of main data_set
    Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
   
    'going through each record in main data_set to find out Total_Stock_Volume and Yearly_change and Change_Percentage for each ticker
    For i = 2 To Last_Row
        
        ' accumulate Total_Stock_Volume for each Ticker
         Total_Stock_Volume = Total_Stock_Volume + CLng(Cells(i, 7).Value)
    
        'finds the first record of each ticker to get the openning_price of the year
        If CStr(Cells(i, 1).Value) <> CStr(Cells(i - 1, 1).Value) Then
            opening_price = CDec(Cells(i, 3).Value)
        End If
        
        'finds the last record of each ticker in the year
        If CStr(Cells(i, 1).Value) <> CStr(Cells(i + 1, 1).Value) Then
            
            'get the closing_price of the year
            closing_price = CDec(Cells(i, 6).Value)
            'calculate yearly_change
            YearlyChange = closing_price - opening_price
            
            'fill the Ticker data_set section with Ticker icon
            Cells(j, 9).Value = CStr(Cells(i, 1).Value)
            'fill the Ticker data_set section with YearlyChange
            Cells(j, 10).Value = YearlyChange
            'Applies conditional formatting that will highlight positive change in green and negative change in red for YearlyChange Column
            If YearlyChange >= 0 Then
                Cells(j, 10).Interior.Color = RGB(0, 255, 0)
            Else
                Cells(j, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            'This if statement is addressing "devide_by_zero error" when the opening_price is zero
            If opening_price Then
                Cells(j, 11).Value = (YearlyChange / opening_price)
            Else
                Cells(j, 11).Value = 0
            End If
            
            'fills the Ticker data_set section by Total_Stock_Volume
            Cells(j, 12).Value = Total_Stock_Volume
           'gets ready for next Ticker in Ticker data_set
            j = j + 1
            
            're-initialize the requried varibales in for next Ticker
            Total_Stock_Volume = 0
            opening_price = 0
            closing_price = 0
         
        End If
       
    Next i
    
    'changes Percent_change column format to percentage
    Columns("K").NumberFormat = "0.00%"
    
    'initialize the varibales required in finding the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
    Greatest_Increase = 0
    Greatest_Decrease = 0
    Greatest_Total_Volume = 0
    k = 0
    m = 0
    n = 0
    
    'find the last row of Ticker data_set
    Last_Row = Cells(Rows.Count, 11).End(xlUp).Row
    
    'going through each record in Ticker data_set to find the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
    For i = 2 To Last_Row - 1
        
        'finds the ticker with Greatest_Increase in the year
        If Greatest_Increase < Cells(i, 11).Value Then
            Greatest_Increase = Cells(i, 11).Value
            k = i
        End If
        
        'finds the ticker with Greatest_decrease in the year
        If Greatest_Decrease > Cells(i, 11).Value Then
            Greatest_Decrease = Cells(i, 11).Value
            m = i
        End If
        
        'finds the ticker with Greatest_Total_Volume in the year
        If Greatest_Total_Volume < Cells(i, 12).Value Then
            Greatest_Total_Volume = Cells(i, 12).Value
            n = i
        End If
    Next i
     
    'Create lable of each row in third data_set
    Cells(2, 15).Value = "Greatest % increase"
    Cells(3, 15).Value = "Greatest % decrease"
    Cells(4, 15).Value = "Greatest total volume"
    
    'create headers of third data_set
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'fills the ticker with greatest increase showing the increase in front of it
    Cells(2, 16).Value = Cells(k, 9).Value
    Cells(2, 17).Value = Cells(k, 11).Value
    
    'fills the ticker with greatest decrease showing the decrease in front of it
    Cells(3, 16).Value = Cells(m, 9).Value
    Cells(3, 17).Value = Cells(m, 11).Value

    'fills the ticker with greatest total volume showing the total volume in front of it
    Cells(4, 16).Value = Cells(n, 9).Value
    Cells(4, 17).Value = Cells(n, 12).Value
    
    'adjust Greatet increase and greaset decrease value cells format with percentage format
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    'autofit both Ticker and third data_sets
    Columns("I:L").Columns.AutoFit
    Columns("O:Q").Columns.AutoFit
    
End Sub



