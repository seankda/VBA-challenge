Sub stockTable()

    ' Declare variables
    Dim symbol As String
    Dim yearlyChange As Double
    Dim totalStockVolume As Double
    Dim summaryTable As Integer
    
    ' Set values to variables
    yearlyChange = 0
    totalStockVolume = 0
    summaryTable = 2
    
    
    ' Name the summary table headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Loop through rows in the column
    ' Start at row 2 to skip headers, and go to the last row 70926
    For i = 2 To 70926
    
        ' Searches for when the value of the next cell is different than that of the current cell
        ' If the value is different, write to the summary table
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            totalStockVolume = totalStockVolume + Cells(i, 7).Value
          
            ' Set the symbol
            symbol = Cells(i, 1).Value
            
            ' Write to the summaryTable
            ' Print the symbol to the summaryTable
            Range("I" & summaryTable).Value = symbol
            ' Print the totalStockVolume
            Cells(summaryTable, 12).Value = totalStockVolume
            
            ' Reset the totalStockVolume amount
            totalStockVolume = 0
            
            ' Add one to the summary table row
            summaryTable = summaryTable + 1
          
        ' If this last value is the same, add it to the count
        Else
            totalStockVolume = totalStockVolume + Cells(i, 7).Value
        
        End If
        
    Next i
    
    'Message Box of the current cell and value of the next cell
    'MsgBox (Cells(i,column).Value & " and then " & Cells(i + 1, column).Value)
End Sub

