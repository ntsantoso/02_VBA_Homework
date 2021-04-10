Option Explicit

Sub HomeworkCode()

'Declaring all variables that will be useful for our code'

Dim r As Long
Dim c As Long

Dim StockVolumeCounter As Long
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim Ticker As Long
Dim last_row As Long
Dim ResultsRow As Long
Dim sheet_number As Long

For sheet_number = 1 To Worksheets.Count

Dim ws As Worksheet
Set ws = Worksheets(sheet_number)

'Insert Column headers for ResultsTable'
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume in Thousands"


'Create last row to navigate much easier'
last_row = ws.Range("A2").End(xlDown).Row

'Indicate the index row for the output table'
ResultsRow = 2

'Set OpeningPrice to the first price for each stock Ticker'
OpeningPrice = ws.Range("C2").Value

'Set Volume to 0 to begin counting for each Ticker'
StockVolumeCounter = 0


'Start loop at A2'
For r = 2 To last_row
    
    'Create a loop with a condition to check if the row below contains the same ticker as the previous one'
    If ws.Cells(r, 1).Value <> ws.Cells(r + 1, 1).Value Then
        
        'Add the volume counter, dividing by 1000 to store values easily'
        StockVolumeCounter = StockVolumeCounter + ws.Cells(r, 7).Value / 1000
        
        'Set the closing price to Column F'
        ClosingPrice = ws.Cells(r, 6).Value
        
        'Percent change condition, prevent an error from dividing by 0'
        If OpeningPrice = 0 Then
            ws.Cells(ResultsRow, 11).Value = "NaN"
            
        Else
        'Calculate the percentage change'
           ws.Cells(ResultsRow, 11).Value = (ClosingPrice - OpeningPrice) / OpeningPrice
           ws.Cells(ResultsRow, 11).NumberFormat = "0.00%"
        
        End If
        
        'Yearly Change calculation + conditional formatting to identify positive/negative'
        ws.Cells(ResultsRow, 10).Value = ClosingPrice - OpeningPrice
            
            If ws.Cells(ResultsRow, 10).Value < 0 Then
                ws.Cells(ResultsRow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(ResultsRow, 10).Interior.ColorIndex = 4
            
            End If
        
        'Total Stock Volume'
        
        ws.Cells(ResultsRow, 12).Value = StockVolumeCounter
        
        'Ticker'
        ws.Cells(ResultsRow, 9).Value = ws.Cells(r, 1).Value
        
        StockVolumeCounter = 0
        
        ResultsRow = ResultsRow + 1
        
        OpeningPrice = ws.Cells(r + 1, 3).Value
        
        
    Else
    
    StockVolumeCounter = StockVolumeCounter + ws.Cells(r, 7).Value / 1000
    

    'if it's not the edge, then dont change the open value, add whatever is in colG to it'


    End If


Next r

Next sheet_number


End Sub
