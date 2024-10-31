Sub Stock_Analysis():

        'Create Variable
        Dim total As Double
        Dim row As Long
        Dim rowcount As Double
        Dim quarterlyChange As Double
        Dim percentChange As Double
        Dim summaryTableRow As Long
        Dim stockStartRow As Long
        Dim startValue As Long
        Dim lastTicker As String
        
        'loop through all worksheet
        For Each ws In Worksheets
        
        'set the title row of the summary section
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'set the title row of the aggregate section
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'initialize values
        summaryTableRow = 0
        total = 0
        quarterlyChange = 0
        stockStartRow = 2
        startValue = 2
        
        'get the value of the last row in the current sheet
        rowcount = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        'find the last ticker so that we can break out of the loop
        lastTicker = ws.Cells(rowcount, 1).Value
        
        'loop until it get to the end of the sheet
        For row = 2 To rowcount
        
            'check to see if the ticker changed
            If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
                
                
                'add to the total stock volume one last time
                total = total + ws.Cells(row, 7).Value     'gets the value from 7th column (G)
            
                'check to see if the value of the total stock volume is 0
                If total = 0 Then
                    'print the results in the summary table section (Column I -L)
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value       
                    ws.Range("J" & 2 + summaryTableRow).Value = 0                                  
                    ws.Range("K" & 2 + summaryTableRow).Value = 0 & "%"                         
                    ws.Range("L" & 2 + summaryTableRow).Value = 0                                  
                    
                Else
                    'Find the first non-zero first open value for the stock
                    If ws.Cells(startValue, 3).Value = 0 Then
                        'if the first open is 0, search for the first non-zero stock open value by moving to the next row
                        For findValue = startValue To row
                        
                            'check to see if the next (or rows afterwards) open value does not equa 0
                            If ws.Cells(findValue, 3).Value <> 0 Then
                                
                                startValue = findValue
                                'break out of the loop
                                Exit For
                            End If
                            
                        Next findValue
                    End If
                    
                    'Calculate the quarterly change (difference in the last close - first-open)
                    quarterlyChange = ws.Cells(row, 6).Value - ws.Cells(startValue, 3).Value
                    
                    'Calculate the percent change (quarterly change / first open)
                    percentChange = quarterlyChange / ws.Cells(startValue, 3).Value
                    
                    'print the results in the summary table section (Column I -L)
                    ws.Range("I" & 2 + summaryTableRow).Value = ws.Cells(row, 1).Value
                    ws.Range("J" & 2 + summaryTableRow).Value = quarterlyChange
                    ws.Range("K" & 2 + summaryTableRow).Value = percentChange
                    ws.Range("L" & 2 + summaryTableRow).Value = total
                    
                    'color the quarterly change column in the summary section based on the value of the quarterly change
                    If quarterlyChange > 0 Then
                        'color the cell green
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 4
                    ElseIf quarterlyChange < 0 Then
                        'color the cell red
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 3
                    Else
                        'color the cell clear or no change
                        ws.Range("J" & 2 + summaryTableRow).Interior.ColorIndex = 0
                    End If
                
                    'reset / update the values for the next ticker
                    total = 0
                    averageChange = 0
                    quarterlyChange = 0
                    startValue = row + 1
                    'move to the next row in the summary table
                    summaryTableRow = summaryTableRow + 1
                
                End If
                
            Else
                'if its in the same ticker, keep adding to the total stock volum
                total = total + ws.Cells(row, 7).Value     
                
            End If

    
    Next row
    
    'update the summary table row
    summaryTableRow = ws.Cells(Rows.Count, "I").End(xlUp).row
    
    'find the last data in the extra rows from column J-L
    Dim lastExtraRow As Long
    lastExtraRow = ws.Cells(Rows.Count, "J").End(xlUp).row
    
    'loop that clears the extra data from columns I-L
    For e = summaryTableRow To lastExtraRow
            'for loop that goes through column I-L (9-12)
            For Column = 9 To 12
                ws.Cells(e, Column).Value = ""
                ws.Cells(e, Column).Interior.ColorIndex = 0
            Next Column
    Next e
    
    'after generating the info in the summarysection, find the greates % increase and decrease then find the greates stock volume
    ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2))
    ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2))
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2))
    
    'use Match() to find the row numbers of the ticker names associated with the greates increase and decrease, then find the greates total volume
    Dim greatestIncreaseRow As Double
    Dim greatestDecreaseRow As Double
    Dim greatestTotVolRow As Double
    greatestIncreaseRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
    greatestDecreaseRow = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & summaryTableRow + 2)), ws.Range("K2:K" & summaryTableRow + 2), 0)
    greatestTotVolRow = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & summaryTableRow + 2)), ws.Range("L2:L" & summaryTableRow + 2), 0)
    
    'display the ticker symbol for the greates increase, decrease, and total stock volume
    ws.Range("P2").Value = ws.Cells(greatestIncreaseRow + 1, 9).Value
    ws.Range("P3").Value = ws.Cells(greatestDecreaseRow + 1, 9).Value
    ws.Range("P4").Value = ws.Cells(greatestTotVolRow + 1, 9).Value
    
    'format the summary table columns
    For s = 0 To summaryTableRow
            ws.Range("J" & 2 + s).NumberFormat = "0.00"
            ws.Range("K" & 2 + s).NumberFormat = "0.00%"
            ws.Range("L" & 2 + s).NumberFormat = "#,###"
    Next s
    
    'format the summary aggregates
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "#,###"
    
    'Autofit the info across all columns
    Columns("A:Q").AutoFit
    
    Next ws
    
    End Sub


