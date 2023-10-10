Sub sheet_function()
    For Each ws In Worksheets
        'Add titles to the columns for ticker summary table
        ws.Range("I1:L1").Value = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
        
        Dim rowCount As Long
        rowCount = ws.Range("A1").End(xlDown).Row
        
        ' Initialize variables 
        Dim firstRow As Long
        Dim tickerCount As Long
        
        tickerRow = 2
        firstRow = 2
    
        For i = 2 To rowCount
            'Check if the next ticker is different
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'Store the ticker
                ws.Range("I" & tickerRow).Value = ws.Cells(i, 1).Value
                'Store the Value Change in $
                ws.Range("J" & tickerRow).Value = (ws.Cells(i, 6).Value - ws.Cells(firstRow, 3).Value)
                'Store the Percentage Change
                ws.Range("K" & tickerRow).Value = ((ws.Cells(i, 6).Value - ws.Cells(firstRow, 3).Value) / ws.Cells(firstRow, 3).Value)
                'Store the sum of column G for this Ticker symbol
                ws.Range("L" & tickerRow).Value = WorksheetFunction.Sum(ws.Range("G" & firstRow & ":G" & i))
                'Increment the counter/pointer for the summay table
                tickerRow = tickerRow + 1
                'Capture the next index as firstRow for the next Ticker
                firstRow = i + 1
            End If
        Next i
        
        'Add titles to the columns for ticker summary table
        rowCount = ws.Range("I1").End(xlDown).Row
        
        Dim profit As FormatCondition
        Dim loss As FormatCondition
        
        'Clear conditiona formatting on column J
        ws.Range("J1:K" & rowCount).FormatConditions.Delete
        
        'Show cells red for <0 and Green for >0
        ws.Columns("J").NumberFormat = "0.00"
        Set profit = ws.Range("J2:K" & rowCount).FormatConditions.Add(xlCellValue, xlGreater, "0")
        profit.Interior.Color = RGB(0, 255, 0)
        Set loss = ws.Range("J2:K" & rowCount).FormatConditions.Add(xlCellValue, xlLess, "0")
        loss.Interior.Color = RGB(255, 0, 0)
        
        'Format column K as percentage and fit the summary table
        ws.Columns("K").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit

        'Create another summary table        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value "
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Fetch greatest %
        ws.Range("I1:L" & rowCount).Sort Key1:=ws.Range("K1"), Order1:=xlDescending, Header:=xlYes
        ws.Range("P2").Value = ws.Range("I2").Value
        ws.Range("Q2").Value = ws.Range("K2").Value
        ws.Range("Q2").NumberFormat = "0.00%"
        
        'Fetch the lowest %
        ws.Range("I1:L" & rowCount).Sort Key1:=ws.Range("K1"), Order1:=xlAscending, Header:=xlYes
        ws.Range("P3").Value = ws.Range("I2").Value
        ws.Range("Q3").Value = ws.Range("K2").Value
        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Fetch the largest Total
        ws.Range("I1:L" & rowCount).Sort Key1:=ws.Range("L1"), Order1:=xlDescending, Header:=xlYes
        ws.Range("P4").Value = ws.Range("I2").Value
        ws.Range("Q4").Value = ws.Range("L2").Value
        
        'Reset the sorting for the range and autofit the new summary table
        ws.Range("I1:L" & rowCount).Sort Key1:=ws.Range("I1"), Order1:=xlAscending, Header:=xlYes
        ws.Columns("O:Q").AutoFit
    Next ws
End Sub

