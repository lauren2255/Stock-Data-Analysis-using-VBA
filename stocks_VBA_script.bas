Attribute VB_Name = "Module1"
Sub stocks()

'declare variables

'declare worksheet
Dim ws As Worksheet

'loop across worksheets
For Each ws In Worksheets

'make variable for last row for raw data
Dim LastRowA As Long
LastRowA = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    
'make variable for ticker name
Dim ticker_name As String

'make variable for opening price & closing price
Dim open_price As Double
open_price = 0
Dim close_price As Double
close_price = 0

'make variable for total volume of stock
Dim stock_total_volume As Double
stock_total_volume = 0

'make variable for yearly change & percent change
Dim Yearly_Change As Double
Yearly_Change = 0
Dim Percent_Change As Double
Percent_Change = 0

'make variable for greatest increase & greatest decrease
Dim greatest_increase As Double
greatest_increase = 0
Dim greatest_decrease As Double
greatest_decrease = 0

'make variable for greatest total volume
Dim greatest_total_volume As Double
greatest_total_volume = 0

'make variables for tickers for greatest increase, decrease, and total volume
Dim greatest_increase_ticker As String
Dim greatest_decrease_ticker As String
Dim greatest_volume_ticker As String

'make variable for Table Row location
Dim SummaryTableRow As Integer
SummaryTableRow = 2

'make column names

'column names for summary table
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'column names for calculations
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'row names for calculations
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

'retrieve opening value for first ticker of worksheet
open_price = ws.Cells(2, 3).Value

    'create loop to cycle through rows for data and summary table for all rows in Column A
    For i = 2 To LastRowA

        'conditional statement for when ticker name changes
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
            'set value for ticker_name variable
            ticker_name = ws.Cells(i, 1).Value
            
            'set value for close_price variable
            close_price = ws.Cells(i, 6).Value
            
            'making yearly change variable
            Yearly_Change = (close_price - open_price)
            
            'making percent change variable
            Percent_Change = Round(((Yearly_Change / open_price) * 100), 2)
            
            'making the total stock volume
            stock_total_volume = stock_total_volume + ws.Cells(i, 7).Value
            
            'set locations for values in summary table
            ws.Range("I" & SummaryTableRow).Value = ticker_name
            ws.Range("J" & SummaryTableRow).Value = Yearly_Change
            'concatenate and add percent sign
            ws.Range("K" & SummaryTableRow).Value = Percent_Change & "%"
            ws.Range("L" & SummaryTableRow).Value = stock_total_volume
    
        'conditional formatting green if >0 and red if <= 0 for yearly change
        If Yearly_Change > 0 Then
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        
        ElseIf Yearly_Change <= 0 Then
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
        
        End If
            
            'reset variables for next loop i.e. next ticker
            stock_total_volume = 0
            Yearly_Change = 0
            Percent_Change = 0
            close_price = 0
            
            'retrieve opening price for first row of new ticker (one more row after the last row of the previous ticker)
            open_price = ws.Cells(i + 1, 3).Value
            
            'reset summary table row for next conditional that meets the if statement
            SummaryTableRow = SummaryTableRow + 1
    
        Else
            
            'keep adding to stock_total volume when conditional is not met i.e. adding up all the stock total for one ticker
            stock_total_volume = stock_total_volume + Cells(i, 7).Value
        
        End If
        
    'next counter for calculations loop
    Next i

'define variable for last row for calculations
Dim LastRowI As Long
LastRowI = ws.Range("I" & ws.Rows.Count).End(xlUp).Row

    'for all rows in Column I
    For i = 2 To LastRowI

        'conditional for greatest increase, setting value to variable, and setting location for greatest increase value and ticker
        If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K:K")) Then
            greatest_increase = ws.Cells(i, 11).Value
            greatest_increase_ticker = ws.Cells(i, 9).Value
            ws.Range("P2").Value = greatest_increase_ticker
            ws.Range("Q2").Value = (greatest_increase * 100) & "%"

        'conditional for greatest decrease, setting value to variable, and setting location for greatest decrease value and ticker
        ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K:K")) Then
            greatest_decrease = ws.Cells(i, 11).Value
            greatest_decrease_ticker = ws.Cells(i, 9).Value
            ws.Range("P3").Value = greatest_decrease_ticker
            ws.Range("Q3").Value = (greatest_decrease * 100) & "%"
        
        End If
        
        'conditional for greatest total, setting value to variable, and setting location for greatest total and ticker
        If ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L:L")) Then
            greatest_total_volume = ws.Cells(i, 12).Value
            greatest_volume_ticker = ws.Cells(i, 9).Value
            ws.Range("P4").Value = greatest_volume_ticker
            ws.Range("Q4").Value = greatest_total_volume
        
        End If
        
    'move to next row
    Next i
    
    'autofit column width
    ws.Columns.AutoFit

'move to next worksheet
Next ws

End Sub
