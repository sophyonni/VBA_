Sub multiple_year_stock_data()

' Set an initial variable for holding the ticker
Dim ticker As String

' Set an initial variable for holding the total per ticker
Dim ticker_volume As Double

' Keep track of the location for each ticker in the summary table
Dim summary_table_row As Integer
summary_table_row = 2

' Loop through all ticker
For i = 2 To 760192

' check if we are still within the same ticker, if it is not...
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    ' Set the ticker
    ticker = Cells(i, 1).Value
    
    ' Add to the ticker volume
    ticker_volume = ticker_volume + Cells(i, 7).Value
    
    ' Print the ticker in the Summary Table
    Range("I" & summary_table_row).Value = ticker
    
     ' Print the ticker in the Summary Table
    Range("J" & summary_table_row).Value = ticker_volume
    
    ' Add one to the summary table row
    summary_table_row = summary_table_row + 1
    
    ' Reset the ticker volume
    ticker_volume = 0
    
    ' if the cell immediately following a row is the same ticker...
    Else
    
    ' Add to the ticker volume
    ticker_volume = ticker_volume + Cells(i, 7).Value
    
      End If
    
    Next i
    
    

End Sub
