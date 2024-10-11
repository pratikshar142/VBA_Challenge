

Sub Multiple_Year_Stock_Data_All()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

    ' Declare variables
    Dim Ticker As String
    Dim lastrow As Long
    Dim i As Long
    Dim Ticker_Table_row As Integer
    Dim qtr_change As Double
    Dim Percent_C As Double
    
    'Initilaize values
    j = 0
    Ticker_Table_row = 2
    Total_Stock_Volume = 0
    i_initial = 2
    
    ' Declare column names
    ws.Range("I1").Value = "Ticker"
    ws.Range("j1").Value = "Quarterly Change($)"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total_Stock_Volume"
    ws.Range("N2").Value = "Greatest % increase"
    ws.Range("N3").Value = "Greatest % decrease"
    ws.Range("N4").Value = "Greatest total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
       
          
' Stock_Total = Stock_Total + Cells(i, 7).Value
   ws.Range("I" & Ticker_Table_row).Value = Ticker
   ws.Range("L" & Ticker_Table_row).Value = Total_Stock_Volume
  
   
 ' Count the number of rows
   lastrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
  ' Check till ticker value changes
  
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            ws.Range("I" & Ticker_Table_row).Value = Ticker
    
            ' Calculate total stock volume
            
            ws.Range("L" & Ticker_Table_row).Value = Total_Stock_Volume
            
        
            ' Calculate quarterly change
            
            
            qtr_change = (ws.Cells(i, 6) - ws.Cells(i_initial, 3))
            ws.Range("J" & Ticker_Table_row).Value = qtr_change
            ws.Range("J" & Ticker_Table_row).NumberFormat = "0.00"
            
                      
            
            'Calculate percent change
            Percent_C = qtr_change / ws.Cells(i_initial, 3)
            ws.Range("K" & Ticker_Table_row).Value = Percent_C
            ws.Range("K" & Ticker_Table_row).NumberFormat = "0.00%"
            
           
          ' Conditonal Formatting for Quarterly Change column
          
            If qtr_change > 0 Then
                ws.Range("J" & Ticker_Table_row).Interior.ColorIndex = 4
                Else
                If qtr_change < 0 Then
                     ws.Range("J" & Ticker_Table_row).Interior.ColorIndex = 3
                 Else
                     ws.Range("J" & Ticker_Table_row).Interior.ColorIndex = 0
                 End If
                 
            End If
            
         i_initial = i + 1
            
            Ticker_Table_row = Ticker_Table_row + 1
            Total_Stock_Volume = 0
            qtr_change = 0
            Percent_C = 0
        
        
        Else
        
            'When Ticker value is same
            
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
           
        End If
  
    Next i
    
    
    'Dim qtr_change_row As Integer
     
 ws.Range("P2") = "%" & WorksheetFunction.Max(Range("K2:K" & lastrow)) * 100
 ws.Range("P3") = "%" & WorksheetFunction.Min(Range("K2:K" & lastrow)) * 100
 ws.Range("P4") = WorksheetFunction.Max(Range("L2:L" & lastrow))
    
   Dim max_num As Double
   Dim min_num As Double
   Dim max_value As Double
    
   max_num = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
   min_num = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & lastrow)), Range("K2:K" & lastrow), 0)
   max_value = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & lastrow)), Range("L2:L" & lastrow), 0)
   
   Range("O2") = ws.Cells(max_num + 1, 9)
   Range("O3") = ws.Cells(min_num + 1, 9)
   Range("O4") = ws.Cells(max_value + 1, 9)
   
   ActiveSheet.Range("I:Q").Font.Bold = True
   ActiveSheet.Range("I:Q").EntireColumn.AutoFit

Next ws
End Sub




 
