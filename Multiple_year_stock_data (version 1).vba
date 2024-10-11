

Sub Multiple_Year_Stock_Data()

    Dim Ticker As String
    Dim lastrow As Long
    Dim i As Long
    Dim Ticker_Table_row As Integer
    Ticker_Table_row = 2
    Total_Stock_Volume = 0
    i_initial = 2
    Dim qtr_change As Double
    Dim Percent_C As Double
    j = 0

    
    ' Declare column names
    Range("I1").Value = "Ticker"
    Range("j1").Value = "Quarterly Change"
    Range("k1").Value = "Percent Change"
    Range("l1").Value = "Total_Stock_Volume"
    Range("N2").Value = "Greatest % increase"
    Range("N3").Value = "Greatest % decrease"
    Range("N4").Value = "Greatest total Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    
    
   
    
    
    
' Stock_Total = Stock_Total + Cells(i, 7).Value
   Range("I" & Ticker_Table_row).Value = Ticker
   Range("L" & Ticker_Table_row).Value = Total_Stock_Volume
  
   
 ' Count the number of rows
   lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
  ' Check till ticker value changes
  
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            Range("I" & Ticker_Table_row).Value = Ticker
    
            ' Calculate total stock volume
            
            Range("L" & Ticker_Table_row).Value = Total_Stock_Volume
            
        
            ' Calculate quarterly change
            
            
            qtr_change = (Cells(i, 6) - Cells(i_initial, 3))
            Range("J" & Ticker_Table_row).Value = qtr_change
            Range("J" & Ticker_Table_row).NumberFormat = "0.00"
            
                      
            
            'Calculate percent change
            Percent_C = qtr_change / Cells(i_initial, 3)
            Range("K" & Ticker_Table_row).Value = Percent_C
            Range("K" & Ticker_Table_row).NumberFormat = "0.00%"
            
           
          ' Conditonal Formatting for Quarterly Change column
          
            If qtr_change > 0 Then
                Range("J" & Ticker_Table_row).Interior.ColorIndex = 4
                Else
                If qtr_change < 0 Then
                     Range("J" & Ticker_Table_row).Interior.ColorIndex = 3
                 Else
                     Range("J" & Ticker_Table_row).Interior.ColorIndex = 0
                 End If
                 
            End If
            
         i_initial = i + 1
            
            Ticker_Table_row = Ticker_Table_row + 1
            Total_Stock_Volume = 0
            qtr_change = 0
            Percent_C = 0
        
        
        Else
        
            'When Ticker value is same
            
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
           
        End If
  
    Next i
    
    
    'Dim qtr_change_row As Integer
    
    'qtr_change_row = Cells(Rows.Count, "I").End(xlUp).Row - 1
 
 Range("P2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
 Range("P3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
 Range("P4") = WorksheetFunction.Max(Range("L2:L" & lastrow))
    
   Dim max_num As Double
   Dim min_num As Double
   Dim max_value As Double
    
   max_num = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
   min_num = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
   max_value = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
   
   Range("O2") = Cells(max_num + 1, 9)
   Range("O3") = Cells(min_num + 1, 9)
   Range("O4") = Cells(max_value + 1, 9)
   
   ActiveSheet.Range("I:Q").Font.Bold = True
   ActiveSheet.Range("I:Q").EntireColumn.AutoFit

End Sub




 
