
Sub homework()
'loop through all worksheets
For Each ws In Worksheets

    'set variable ticker
    Dim Ticker As String
    
    'set total to zero
    Dim Total_Volume As Double
    Total_Volume = 0
    
    'set row count to zero
    Dim Row_Count As Double
    Row_Count = 0
    
    
    'create headers for summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Yearly % Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    'set location of each ticker in summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    'create last row for loop
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set first ticker openprice before loop
    OpenPrice = ws.Cells(2, 3).Value
    
    'loop through stocks
    For i = 2 To Lastrow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'set ticker
        Ticker = ws.Cells(i, 1).Value
    
    'Calculate yearly change
    ClosePrice = ws.Cells(i, 6).Value
    Price_Change = ClosePrice - OpenPrice

    'Calculate Price Change Percent, adjust for zeros throwing error
    If (OpenPrice <> 0) Then
    Percent_Price_Change = Price_Change / OpenPrice
    End If
    
    'Add Total Volume
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
    'print ticker to summary table
    ws.Range("I" & Summary_Table_Row).Value = Ticker
    
    'print price change
    ws.Range("J" & Summary_Table_Row).Value = Price_Change
    
        
    'format color for price change
        If (Price_Change > 0) Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf (Price_Change <= 0) Then
        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
    
    'print percent price change
    ws.Range("K" & Summary_Table_Row).Value = (Percent_Price_Change)
    ws.Range("K" & Summary_Table_Row).Style = "Percent"
    ws.Range("K" & Summary_Table_Row).NumberFormat = "#.##%"
    
    'print total volume
    ws.Range("L" & Summary_Table_Row).Value = Total_Volume
    
    'move to next summary table row
    Summary_Table_Row = Summary_Table_Row + 1
    
    'reset next ticker open price
    OpenPrice = ws.Cells(i + 1, 3).Value
    
    
    'reset the total volume back to zero
    Total_Volume = 0
    Price_Change = 0
    Percent_Price_Change = 0
    
    
    Else
    'add to total volume
    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
    
    End If
    
    Next i
    
    

Next ws

End Sub


