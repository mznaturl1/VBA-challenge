Option Explicit

Sub stock_summary()

'Declare WS as worksheet
Dim ws As Worksheet
For Each ws In Worksheets

'Establishing headings for results table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
'Set the last row of data
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
 'Designate header row
    Dim summary_row, start_row As Long
    summary_row = 2
    start_row = 2
  
  'Creating loop for the daily records
    Dim i As Long
    For i = 2 To last_row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(summary_row, 9).Value = ws.Cells(i, 1).Value
    
            Dim year_closing, year_opening, year_change, percent_change As Double
            year_closing = ws.Cells(i, 6).Value
            year_opening = ws.Cells(start_row, 3).Value
  
  'Calculating yr opening and closing values as percent
            year_change = year_closing - year_opening
            
            If (year_opening > 0) Then
                percent_change = year_change / year_opening
            Else
                percent_change = 0
            End If
    
            Dim sum_range As String
            sum_range = "G" & start_row & ":G" & i
    
            Dim total_volume As LongLong
            total_volume = WorksheetFunction.Sum(ws.Range(sum_range))
    
    'conditional formatting
            ws.Cells(summary_row, 10).Value = year_change
            If ws.Cells(summary_row, 10).Value >= 0 Then
                ws.Cells(summary_row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(summary_row, 10).Interior.ColorIndex = 3
            End If
            
            ws.Cells(summary_row, 11).Value = percent_change
            ws.Cells(summary_row, 11).NumberFormat = "0.00%"
            ws.Cells(summary_row, 12).Value = total_volume
    
            start_row = i + 1
            summary_row = summary_row + 1
        End If
    Next i
    
'Add formatting to columns i-l
    ws.Columns("I:L").AutoFit
    
  'Set next loop for max & min change, and max volume
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    Dim total_summary_rows As Long
    total_summary_rows = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    Dim max_change, min_change As Double
    Dim max_position, min_position, max_stock_position As Long
    Dim max_volume As LongLong
    
    max_change = ws.Cells(2, 11).Value
    max_position = 2
    
    min_change = ws.Cells(2, 11).Value
    min_position = 2
    
    max_volume = ws.Cells(2, 12).Value
    max_stock_position = 2
    
'loop for max & min change, and max volume
    Dim j As Long
    For j = 3 To total_summary_rows
        If (ws.Cells(j, 11).Value > max_change) Then
            max_change = ws.Cells(j, 11).Value
            max_position = j
        End If
        
        If (ws.Cells(j, 11).Value < min_change) Then
            min_change = ws.Cells(j, 11).Value
            min_position = j
        End If
        
        If (ws.Cells(j, 12).Value > max_volume) Then
            max_volume = ws.Cells(j, 12).Value
            max_stock_position = j
        End If
        
    Next j
        
 'Functionality to provide stock with greastest % +/-, and volume
 
        ws.Range("P2").Value = ws.Cells(max_position, 9).Value
        ws.Range("P3").Value = ws.Cells(min_position, 9).Value
        ws.Range("P4").Value = ws.Cells(max_stock_position, 9).Value
        
        ws.Range("Q2").Value = max_change
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = min_change
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = max_volume
 
 'Add formatting to columns o-q
        ws.Columns("O:Q").AutoFit
        
    Next ws
    
End Sub