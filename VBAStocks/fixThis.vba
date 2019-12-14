Sub stockTable()

    ' Declare variables
    Dim yearOpen As Double
    Dim yearClose As Double
    
    Dim symbol As String
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim volume As Double
    
    Dim lastRow As Long
    Dim i As Long

    Dim summaryTable As Integer
    
    ' Set values to variables
    yearOpen = 0
    yearClose = 0
    yearlyChange = 0
    percentChange = 0
    volume = 0
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    summaryTable = 2
    
    ' Name the summary table headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    ' Set K2:K (Percent Change) to a Percent Number Format
    Range(Range("K2"), Range("K2").End(xlDown)).NumberFormat = "0.00%"
    
    ' Loop through rows in the column
    ' Start at row 2 to skip headers, and go to the last row 70926
    For i = 2 To lastRow
    
        ' Searches for when the value of the next cell is different than that of the current cell
        ' If the value is different, write to the summary table
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            volume = volume + Cells(i, 7).Value
          
            ' Set the symbol
            symbol = Cells(i, 1).Value
            
            ' Year Open
            yearOpen = Cells(i + 1, 3).Value
            
            ' Year Close
            yearClose = Cells(i, 6).Value
            
            ' Yearly Change
            yearlyChange = yearClose - yearOpen
            
            ' Percent Change
            percentChange = (yearClose - yearOpen) / yearClose
            
            'Message Box of the current cell and value of the next cell
            'MsgBox (Cells(i,column).Value & " and then " & Cells(i + 1, column).Value)
            
            ' Write to the summaryTable
            ' Print the symbol to the summaryTable
            Range("I" & summaryTable).Value = symbol
            ' Print the yearlyChange to the summaryTable
            Cells(summaryTable, 10).Value = yearlyChange
            ' Print the percentChange to the summaryTable
            Cells(summaryTable, 11).Value = percentChange
            ' Print the volume
            Cells(summaryTable, 12).Value = volume
            
            
            ' Reset the volume amount
            yearOpen = 0
            yearClose = 0
            yearlyChange = 0
            percentChange = 0
            volume = 0
            
            ' Add one to the summary table row
            summaryTable = summaryTable + 1
          
        ' If this last value is the same, add it to the count
        Else
            volume = volume + Cells(i, 7).Value
        
        End If
        
    Next i
    
End Sub




