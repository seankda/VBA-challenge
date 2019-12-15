Sub stockTable()

    ' Make this run on every worksheet
    Dim sheet As Worksheet
    
    ' Loop through all the sheets
    For Each sheet In Worksheets

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
        
        ' Dim variables for Challenge
        Dim topSymbol As String
        Dim topSymbolVal As Double
        Dim botSymbol As String
        Dim botSymbolVal As Double
        Dim maxSymbol As String
        Dim maxSymbolVal As Double
    
        
        ' Set values to variables
        yearOpen = 0
        yearClose = 0
        yearlyChange = 0
        percentChange = 0
        volume = 0
        
        ' Reference variables
        lastRow = sheet.Cells(Rows.Count, 1).End(xlUp).Row
        summaryTable = 2
        
        ' Set values to Challenge variables
        topSymbolVal = 0
        botSymbolVal = 0
    
        
        ' Name the summary table headers
        sheet.Range("I1").Value = "Ticker"
        sheet.Range("J1").Value = "Yearly Change"
        sheet.Range("K1").Value = "Percent Change"
        sheet.Range("L1").Value = "Total Stock Volume"
        
        ' Challenge Table
        sheet.Range("N2").Value = "Greatest % Increase"
        sheet.Range("N3").Value = "Greatest % Decrease"
        sheet.Range("N4").Value = "Greatest Total Volume"
        sheet.Range("O1").Value = "Ticker"
        sheet.Range("P1").Value = "Value"
        
        
        ' Format Cells and Columns to Percent
        sheet.Range("K2:K" & lastRow).NumberFormat = "0.00%"
        sheet.Range("P2").NumberFormat = "0.00%"
        sheet.Range("P3").NumberFormat = "0.00%"
        
         ' Year Open
         ' Note to self: I had to move this out of the loop otherwise it would get the last row's
         '               open value and subtract it from the last close each time
        yearOpen = sheet.Cells(2, 3).Value
    
        ' Loop through rows in the column
        ' Start at row 2 to skip headers, and go to the last row
        For i = 2 To lastRow
        
            ' Searches for when the value of the next cell is different than that of the current cell
            ' If the value is different, write to the summary table
            If sheet.Cells(i + 1, 1).Value <> sheet.Cells(i, 1).Value Then
            
            ' Repeat inside the loop to prevent value exclusion
                volume = volume + sheet.Cells(i, 7).Value
              
                ' Set the symbol
                symbol = sheet.Cells(i, 1).Value
                
                ' Year Close
                yearClose = sheet.Cells(i, 6).Value
                
                ' Yearly Change
                yearlyChange = yearClose - yearOpen
                
                ' Convert to percent
                ' <> 0 needs to be used in order to skip any 0's in the data that would cause an overflow due to division by 0
                If yearOpen <> 0 Then
                    percentChange = (yearClose - yearOpen) / yearOpen
                End If
                
                'Early check/test
                'Message Box of the current cell and value of the next cell
                'MsgBox (Cells(i,column).Value & " and then " & Cells(i + 1, column).Value)
                
                ' Write to the summaryTable
                ' Print the symbol to the summaryTable
                ' Range(summaryTable, 9).Value = symbol is another way of doing it but requires you to count index.
                sheet.Range("I" & summaryTable).Value = symbol
                ' Print the yearlyChange to the summaryTable
                sheet.Range("J" & summaryTable).Value = yearlyChange
                ' Print the percentChange to the summaryTable
                sheet.Range("K" & summaryTable).Value = percentChange
                ' Print the volume
                sheet.Range("L" & summaryTable).Value = volume
                
                ' Conditional formatting based on values
                If (yearlyChange > 0) Then
                ' Fill cell with Green if yearlyChange is greater than 0
                    sheet.Cells(summaryTable, 10).Interior.ColorIndex = 4
                ElseIf (yearlyChange <= 0) Then
                ' Fill cell with Red if yearlyChange is less than or equal to 0
                    sheet.Cells(summaryTable, 10).Interior.ColorIndex = 3
                ' Note to self: Stop forgetting to end if statements
                End If
                
                ' Add one to the summary table row
                summaryTable = summaryTable + 1
                ' Note to self: It's important not to set the yearOpen back to 0, otherwise it only works for first instance
                yearOpen = sheet.Cells(i + 1, 3).Value
                yearClose = 0
                yearlyChange = 0
                
                ' Challenge
                ' Set the topSymbolValue to percentChange and topSymbol to symbol if percentChange is greater than topSymbolVal
                If (percentChange > topSymbolVal) Then
                    topSymbolVal = percentChange
                    topSymbol = symbol
                ' If not true, then...
                ' Set the botSymbolVal to percentChange and the botSymbol to symbol if percentChange is less than botSymbolVal
                ElseIf (percentChange < botSymbolVal) Then
                    botSymbolVal = percentChange
                    botSymbol = symbol
                End If
                
                ' If volume is greater than maxSymbolVal then make maxSymbolVal equal to volume and maxSymbol equal to symbol
                If (volume > maxSymbolVal) Then
                    maxSymbolVal = volume
                    maxSymbol = symbol
                End If
                
                ' Assign variables to their cell/range positions
                sheet.Range("O2").Value = topSymbol
                sheet.Range("O3").Value = botSymbol
                sheet.Range("O4").Value = maxSymbol
                sheet.Range("P2").Value = topSymbolVal
                sheet.Range("P3").Value = botSymbolVal
                sheet.Range("P4").Value = maxSymbolVal
                
                ' Reset the percentChange and volume amounts
                percentChange = 0
                volume = 0
              
            ' If this last value is the same, add it to the count
            Else
                volume = volume + sheet.Cells(i, 7).Value
            End If
        Next i
        ' Autofit the data to make it look nice :)
        ' Probably could have created a last column variable for this just like we did for the last row variable
        sheet.Columns("A:P").AutoFit
    ' Repeat everything in the loop for the next worksheet
    Next sheet
' THIS IS THE END OF VBA!
End Sub







