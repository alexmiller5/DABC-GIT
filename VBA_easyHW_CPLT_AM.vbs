Sub VBAstockVolume():
 ' set columns for Ticker_id and Total_Volume
  
    Range("i1").Value = "Ticker ID"
    Range("j1").Value = "Total Volume"
    
' define parameters
    Dim Ticker_id As String
    
    Dim Total_Volume As Double
    Dim CurrentRow As Double
    Dim LastRow As Double
    
    Total_Volume = 0
    CurrentRow = 2
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set Loop
        For i = 2 To LastRow
        Total_Volume = Total_Volume + Cells(i, 7)
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        Ticker_id = Cells(i, 1).Value
        
        Cells(CurrentRow, 9).Value = Ticker_id
        Cells(CurrentRow, 10).Value = Total_Volume
        
        CurrentRow = CurrentRow + 1
        
    Total_Volume = 0

    
    End If
    
    Next i
        
End Sub