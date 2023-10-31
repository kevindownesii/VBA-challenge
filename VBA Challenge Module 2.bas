Attribute VB_Name = "Module1"
Sub MultipleYearStockData()

'Definitions
    For Each ws In Worksheets
    
        Dim WorksheetName As String
        'Open Stock Price
        Dim OpenPrice As Double
        'Closing Stock Price
        Dim ClosePrice As Double
        'First row of data
        Dim i As Long
        'Start row of ticker
        Dim j As Long
        'Stock Symbol to popualte Ticker column
        Dim Ticker As Long
        'Value for yearly change calculation
        Dim YearlyChange As Double
        'Value for percent change calculation
        Dim PercentChange As Double
        'Value for greatest increase calculation
        Dim GreatestIncrease As Double
        'Value for greatest decrease calculation
        Dim GreatestDecrease As Double
        'Value for greatest total volume
        Dim GreatestVolume As Double
        'Last row column A
        Dim LastRowColumnA As Long
        'Last row column I
        Dim LastRowColumnI As Long
        
                
        'Create column headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Get the WorksheetName
        WorksheetName = ws.Name
        
        'Set Ticker to first row
        Ticker = 2
        
        'Set beginning row to 2
         j = 2
        
        'Set beginning row to 2
        i = 2
        
       
        'Find the last cell in Column A
        LastRowColumnA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
            'Loop through all rows
            For i = 2 To LastRowColumnA
            
                'Find unique ticker symbol
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Populate ticker symbol in column I "Ticker"
                ws.Cells(Ticker, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate Yearly Change in column J "Yearly Change"
                ws.Cells(Ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(i, 3).Value
                
                'Format value to $
                ws.Cells(Ticker, 10).NumberFormat = "$#,##0.00"
                                
                    'Format color
                    If ws.Cells(Ticker, 10).Value >= 0 Then
                
                    'Color Green for positive
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 4
                
                    Else
                
                    'Color Red for negative
                    ws.Cells(Ticker, 10).Interior.ColorIndex = 3
                
                    End If
                    
                    'Calculate percent change in column K "Percent Change"
                    If ws.Cells(Ticker, 3).Value <> 0 Then
                    PercentChange = ((ws.Cells(i, 6).Value - ws.Cells(i, 3).Value) / ws.Cells(i, 3).Value)
                    
                    'Format value to a %
                    ws.Cells(Ticker, 11).Value = Format(PercentChange, "Percent")
                    
                    Else
                    
                    ws.Cells(Ticker, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                    'Format color
                    If ws.Cells(Ticker, 11).Value >= 0 Then
                
                    'Color Green for positive
                    ws.Cells(Ticker, 11).Interior.ColorIndex = 4
                
                    Else
                
                    'Color Red for negative
                    ws.Cells(Ticker, 11).Interior.ColorIndex = 3
                
                    End If
                    
                'Calculate total stock volume in column L "Total Stock Volume"
                ws.Cells(Ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase Ticker by 1
                Ticker = Ticker + 1
                
                'Set new beginning row of the ticker
                j = i + 1
                
                End If
            
            Next i
            
        'Find last cell in column I
        LastRowColumnI = ws.Cells(Rows.Count, 9).End(xlUp).Row
       
        
        'Insert values for summary table calculations
        GreatestIncrease = ws.Cells(2, 11).Value
        GreatestDecrease = ws.Cells(2, 11).Value
        GreatestVolume = ws.Cells(2, 12).Value
            
            'Loop for summary calculations
            For i = 2 To LastRowColumnI
            
                                
                'For greatest increase: if following value is bigger take that value if not take first value
                If ws.Cells(i, 11).Value >= GreatestIncrease Then
                GreatestIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestIncrease = GreatestIncrease
                
                End If
                
                'For greatest decrease: if following value is smaller take that value if not take first value
                If ws.Cells(i, 11).Value <= GreatestDecrease Then
                GreatestDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestDecrease = GreatestDecrease
                
                End If
                
                'For greatest total volume: if following value is bigger take that value if not take first value
                If ws.Cells(i, 12).Value >= GreatestVolume Then
                GreatestVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GreatestVolume = GreatestVolume
                
                End If
                
            'Format summary results
            ws.Range("Q2").Value = Format(GreatestIncrease, "Percent")
            ws.Range("Q3").Value = Format(GreatestDecrease, "Percent")
            ws.Range("Q4").Value = Format(GreatestVolume, "General Number")
            
            Next i
            
        'Bold the title row
        ws.Range("A1:Q1").Font.Bold = True
        ws.Range("O1:O4").Font.Bold = True
        
        'Autofit Columns to Fit
        ws.Columns("A:Q").AutoFit
        
            
    Next ws
        
End Sub


