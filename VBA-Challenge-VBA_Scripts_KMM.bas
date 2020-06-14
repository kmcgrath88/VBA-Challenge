Attribute VB_Name = "Module1"
Sub VBA_HW():

'Loop for worksheets
For Each ws In Worksheets

'CHALLENGE - label cells
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

'Create labels for summary table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


    'Set variable for holding the ticker name
    Dim ticker_name As String

    'Set variable for holding the total volume per ticker
    Dim total_vol As Double
    total_vol = 0

    'Keep track of the location for each ticker in the summary table
    Dim table_row As Integer
    table_row = 2
    
    'Set Yearly Open and Yearly Close
    Dim yearOpen, yearClose, yearlyChange, percentChange As Double
    yearOpen = ws.Cells(2, 3).Value
              
    'Last row
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Loop through all tickers ----- change this
    For i = 2 To lastRow
    
        'Check if the ticker matches the one below it, if not then...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        total_vol = total_vol + ws.Cells(i, 7).Value
                                   
        'Output ticker and volume to table
        ws.Range("I" & table_row).Value = ticker
        ws.Range("L" & table_row).Value = total_vol
            
        'yearClose, YearlyChange, Print in table
        yearClose = ws.Cells(i, 6).Value
        yearlyChange = yearClose - yearOpen
        ws.Range("J" & table_row).Value = yearlyChange
                
            'Conditional formating - nested if for yearlyChange
            If yearlyChange < 0 Then
            ws.Range("J" & table_row).Interior.ColorIndex = 3
                
            Else
            ws.Range("J" & table_row).Interior.ColorIndex = 4
                
            End If
                
            'Percent Change
            If yearOpen = 0 Then
            percentChange = yearClose - yearOpen
            Else
            percentChange = (yearlyChange / yearOpen)
            End If
            ws.Range("K" & table_row).Value = percentChange
            
                     
            'Move down a row & reset total volume
            table_row = table_row + 1
            total_vol = 0
            
            'Reset yearOpen
            yearOpen = ws.Cells(i + 1, 3).Value
            
                            
        'If the ticker below is the same..
        Else
            'Add the total volume of like tickers
            total_vol = total_vol + ws.Cells(i, 7).Value
        
        End If
    
        
    Next i
    
    'Change Percent Change column to percent
    ws.Range("K2:K" & lastRow).NumberFormat = "0.00%"
    
    'Challenge - Greatest % Increase and Match Ticker
    ws.Range("P2").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & table_row))
    ws.Range("P2").NumberFormat = "0.00%"
    Dim increaseNumber As Integer
    increaseNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & table_row)), ws.Range("K2:K" & table_row), 0)
    
    'Challenge - Greatest % Decrease and Match Ticker
    ws.Range("P3").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & table_row))
    ws.Range("P3").NumberFormat = "0.00%"
    Dim decreaseNumber As Integer
    decreaseNumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & table_row)), ws.Range("K2:K" & table_row), 0)
    
    'Challenge - Greatest Volume and Match Ticker
    ws.Range("P4") = WorksheetFunction.Max(ws.Range("L2:L" & table_row))
    Dim great_volNumber As Integer
    great_volNumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & table_row)), ws.Range("L2:L" & table_row), 0)
    
    'Challenge - Print Tickers
    ws.Range("O2") = ws.Cells(increaseNumber + 1, 9).Value
    ws.Range("O3") = ws.Cells(decreaseNumber + 1, 9).Value
    ws.Range("O4") = ws.Cells(great_volNumber + 1, 9).Value
    
        
Next ws

    
    
End Sub
