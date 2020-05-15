Sub MultipleYearStockData()

'looping through all sheets in the workbook
For Each ws In Worksheets


'Declaring variables to be used for this project
 
        Dim r As Long ' first ticker session
        Dim c As Long
        Dim TickerSymbolCount As Long
        Dim Summary_Table_Row As Long
        Dim Ticker_total As Long
        Dim Greatperincr As Double
        Dim GreatperDecr As Double
        Dim WorksheetName As String
        Dim percentchange As Double
        Dim GreatTotVol As Double
        
        'Get the WorksheetName
        WorksheetName = ws.Name
 
 
       'Create column headers for each sheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker counter
         TickerSymbolCount = 2
         c = 2
         
         'Count the non blank cell in ticker column
         Summary_Table_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
         
         
         ' Loop through all Ticker
           For r = 2 To Summary_Table_Row
           
                 'Check if we are still within the same ticker, if it is not...
                    If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
           
                'Write ticker in column I
                ws.Cells(TickerSymbolCount, 9).Value = ws.Cells(r, 1).Value
                
                
                    'Calculate and write Yearly Change
                    ws.Cells(TickerSymbolCount, 10).Value = ws.Cells(r, 6).Value - ws.Cells(c, 3).Value
                
                    'Conditional formating
                    If ws.Cells(TickerSymbolCount, 10).Value < 0 Then
                
                    'Set cell background color to red
                    ws.Cells(TickerSymbolCount, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green
                    ws.Cells(TickerSymbolCount, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change
                    If ws.Cells(c, 3).Value <> 0 Then
                    percentchange = ((ws.Cells(r, 6).Value - ws.Cells(c, 3).Value) / ws.Cells(c, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TickerSymbolCount, 11).Value = Format(percentchange, "Percent")
                    
                    Else
                    
                    ws.Cells(TickerSymbolCount, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write total volume.
                ws.Cells(TickerSymbolCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(c, 7), ws.Cells(r, 7)))
                
                'Increase TickerSymbolCount by 1
                TickerSymbolCount = TickerSymbolCount + 1
                
                'Set new start row of the TickerSymbolCount
                c = r + 1
                
                End If
        
                
           
           Next r

'HARD PART

        Ticker_total = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            'Prepare for summary for Great Total Volume
        GreatTotVol = ws.Cells(2, 12).Value
        Greatperincr = ws.Cells(2, 11).Value
        GreatperDecr = ws.Cells(2, 11).Value
        
        
        For r = 2 To Ticker_total
        
                    
                If ws.Cells(r, 12).Value > GreatTotVol Then
                GreatTotVol = ws.Cells(r, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(r, 9).Value
                
                Else
                
                GreatTotVol = GreatTotVol
                
                End If
                
                
                If ws.Cells(r, 11).Value > Greatperincr Then
                Greatperincr = ws.Cells(r, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(r, 9).Value
                
                Else
                
                Greatperincr = Greatperincr
                
                End If
                
                
                If ws.Cells(r, 11).Value < GreatperDecr Then
                GreatperDecr = ws.Cells(r, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(r, 9).Value
                
                Else
                
                GreatperDecr = GreatperDecr
                
                End If
                
          
            ws.Cells(2, 17).Value = Format(Greatperincr, "Percent")
            ws.Cells(3, 17).Value = Format(GreatperDecr, "Percent")
            ws.Cells(4, 17).Value = Format(GreatTotVol, "Scientific")
            
            
            
        Next r
        
    'Format the results table
     Worksheets(WorksheetName).Columns("A:R").AutoFit
     ws.Cells(2, 17).NumberFormat = "0.00%"
     ws.Cells(3, 17).NumberFormat = "0.00%"

  Next ws

End Sub

