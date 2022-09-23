Attribute VB_Name = "Module1"
Sub VBA_Challenge()
'Instructions: Create a script that loops through all the stocks for one year and outputs the following information:
'1)The ticker symbol.
'2)Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'3)The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'4)The total stock volume of the stock.
'Note: Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
'BONUS: Return the stock with the Greatest % increase, Greatest % decrease, and Greatest total volume

'Declare current worksheet object variable
Dim ws As Worksheet

'Loop through all sheets
For Each ws In Worksheets
    
    'Set an initial variable for holding the ticker name
    Dim Ticker As String
    
    'Set an initial variable for holding the ticker open price
    Dim Open_Price As Double
    
    'Set an initial variable for holding the ticker close price
    Dim Close_Price As Double

    'Set an initial variable for holding the total stock volume per ticker
    Dim Ticker_Volume As LongLong
    Ticker_Volume = 0

    'Keep track of the location for each ticker symbol in the summary table
    Dim Ticker_Table_Row As Integer
    Ticker_Table_Row = 2
    
    'Determine the Last Row on worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    '''''Determine the Row that has the close price
    Dim ClosePriceRow As Integer
    
    'Add a Column for Ticker
    ws.Range("I1").Value = "Ticker"
    
    'Add a Column for Yearly Change
    ws.Range("J1").Value = "Yearly Change"
    
    'Add a Column for Percent Change
    ws.Range("K1").Value = "Percent Change"
    
    'Add a Column for Total Stock Volume
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Set the ticker Open price
    Open_Price = ws.Cells(2, 3).Value
    
    'Loop through all tickers on worksheet
    For i = 2 To LastRow
           
        'Check if we are still within the same ticker, if it is not. . .
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Set the Ticker name
            Ticker = ws.Cells(i, 1).Value
                        
            'Set the Ticker Close price
            Close_Price = ws.Cells(i, 6).Value
            
            'Print the Yearly change from opening price to closing price
            ws.Range("J" & Ticker_Table_Row).Value = Close_Price - Open_Price
            
                'Conditionally format Yearly Change column to show red for negatives and green for positives
                If ws.Range("J" & Ticker_Table_Row).Value >= 0 Then
                ws.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 4
                Else
                ws.Range("J" & Ticker_Table_Row).Interior.ColorIndex = 3
                End If
            
            'Print the Percent change from opening price to closing price
            ws.Range("K" & Ticker_Table_Row).Value = (Close_Price - Open_Price) / Open_Price
            
            'Format Percent change cell as 0.00%
            ws.Range("K" & Ticker_Table_Row).NumberFormat = "0.00%"
            
            'Add to the Ticker Volume total
            Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
            
            'Print the Ticker symbol in the summary table
            ws.Range("I" & Ticker_Table_Row).Value = Ticker
            
            'Print the ticker volume total in the summary table
            ws.Range("L" & Ticker_Table_Row).Value = Ticker_Volume
            
            'Add one to the summary table row
            Ticker_Table_Row = Ticker_Table_Row + 1
            
            'Reset the ticker open price
            Open_Price = ws.Cells(i + 1, 3)
            
            'Reset the ticker volume total
            Ticker_Volume = 0
            
        'If the cell immediately following a row is the same brand. . .
        Else
        
            'Add to the Ticker volume total
            Ticker_Volume = Ticker_Volume + ws.Cells(i, 7).Value
        
        End If
        
    Next i
    
Dim GreatestIncrease As Double, GreatestDecrease As Double, GreatestVolume As LongLong
GreatestIncrease = 0
GreatestDecrease = 0
GreatestVolume = 0

    'Add a table for the bonus question
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Determine the Last Row on summary table
    LastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Loop through all tickers on summary table
    For i = 2 To LastSummaryRow
    
        'Check and print the greatest % increase
        If ws.Cells(i, 11).Value > GreatestIncrease Then
        GreatestIncrease = ws.Cells(i, 11).Value
        ws.Range("Q2").Value = GreatestIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
        End If
    Next i
    
    'Loop through all tickers on summary table
    For i = 2 To LastSummaryRow
    
        'Check and print the greatest % decrease
        If ws.Cells(i, 11).Value < GreatestDecrease Then
        GreatestDecrease = ws.Cells(i, 11).Value
        ws.Range("Q3").Value = GreatestDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
        End If
     Next i
     
    'Loop through Ticker Summary Table and print the tickers with the greatest % increase
    For i = 2 To LastSummaryRow
        If ws.Cells(i, 11).Value = GreatestIncrease Then
        ws.Range("P2").Value = ws.Cells(i, 9).Value
        End If
    Next i
        
    'Loop through Ticker Summary Table and print the ticker with the greatest % decrease
    For i = 2 To LastSummaryRow
        If ws.Cells(i, 11).Value = GreatestDecrease Then
        ws.Range("P3").Value = ws.Cells(i, 9).Value
        End If
    Next i
    
    'Loop through Total Stock Volume and print the stock with the greatest total volume
    For i = 2 To LastSummaryRow
        
        If ws.Cells(i, 12).Value > GreatestVolume Then
        GreatestVolume = ws.Cells(i, 12).Value
        End If
        ws.Range("Q4").Value = GreatestVolume
    Next i
    
    'Loop through Ticker Summary Table and print the ticker with the greatest total volume
    For i = 2 To LastSummaryRow
        If ws.Cells(i, 12).Value = GreatestVolume Then
        ws.Range("P4").Value = ws.Cells(i, 9).Value
        End If
    Next i
    
    
Next ws

End Sub

