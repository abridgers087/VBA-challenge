Attribute VB_Name = "Module2"
Sub ChallengePart2()

'Now do it for all the worksheets
For Each ws In Worksheets

    'Name all the things
    Dim ticker As String
    Dim stock_volume As Double
        stock_volume = 0
        
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim stock_open As Double
    Dim stock_close As Double
    
    Dim last_row As Long
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
    'set up summary table output area
    Dim summary_table As Double
        summary_table = 2
        
    'name more things for additional funtionality
    Dim greatest_increase As Double
        greatest_increase = 0
    
    Dim greatest_decrease As Double
        greatest_decrease = 0
        
    'define ranges for 3rd calc
    Dim greatest_volume As Single
        greatest_volume = 0
        
    'data for summary table (across worksheets)
    ws.Range("i1") = "Ticker"
    ws.Range("j1") = "Yearly Change"
    ws.Range("k1") = "Percent Change"
    ws.Range("l1") = "Total Stock Volume"
    
    'add greater functionality (across worksheets)
    ws.Range("o2") = "Greatest % Increase"
    ws.Range("o3") = "Greatest % Decrease"
    ws.Range("o4") = "Greatest Total Volume"
    ws.Range("p1") = "Ticker"
    ws.Range("q1") = "Volume"
    
    
    'let us start the loop(s) (across worksheets)
        For i = 2 To last_row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                ticker = ws.Cells(i, 1).Value
                stock_volume = stock_volume + ws.Cells(i, 7).Value
        
                ws.Range("I" & summary_table).Value = ticker
                ws.Range("L" & summary_table).Value = stock_volume
                                    
                stock_volume = 0
                
                stock_close = ws.Cells(i, 6)
                
                    'next if
                    If stock_open = 0 Then
                        yearly_change = 0
                        percent_change = 0
                        
                    Else:
                        yearly_change = stock_close - stock_open
                        percent_change = (stock_close - stock_open) / stock_open
                    
                    End If
                               
                'output data (across sheets)
                ws.Range("J" & summary_table).Value = yearly_change
                ws.Range("K" & summary_table).Value = percent_change
                
                'set to dollars (across sheets)
                ws.Range("J" & summary_table).Style = "Currency"
                ws.Range("J" & summary_table).NumberFormat = "$#,##0.00"
                           
                'set to percent (across sheets)
                ws.Range("K" & summary_table).Style = "Percent"
                ws.Range("K" & summary_table).NumberFormat = "0.00%"
                        
                'Next summary table row
                summary_table = summary_table + 1
                        
                '(across sheets)
                ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                       stock_open = ws.Cells(i, 3)
                        
                Else:
                     stock_volume = stock_volume + ws.Cells(i, 7).Value

            End If
                             
        Next i
        
        'formatting for results of yearly change (across sheets)
        For j = 2 To last_row
        
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
                
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
                
            End If
        
        Next j
        
        'additional functionality loop
        'for greatest increase
        For k = 2 To last_row
            
            If ws.Cells(k, 11).Value > greatest_increase Then
                greatest_increase = ws.Cells(k, 11).Value
                ws.Range("p2").Value = ws.Cells(k, 9).Value
                ws.Range("q2").Value = ws.Cells(k, 12).Value
            
            End If
        
        Next k
        
        'for greatest decrease
        For m = 2 To last_row
        
            If ws.Cells(m, 11).Value < greatest_decrease Then
                greatest_decrease = ws.Cells(m, 11).Value
                ws.Range("p3").Value = ws.Cells(m, 9).Value
                ws.Range("q3").Value = ws.Cells(m, 12).Value
                
            End If
            
        Next m
        
        'for greatest volume
        For n = 2 To last_row
            If ws.Cells(n, 12).Value > greatest_volume Then
                greatest_volume = ws.Cells(n, 12)
                ws.Range("q4").Value = greatest_volume
                ws.Range("p4").Value = Cells(n, 9).Value
            
        End If
        
        Next n
        
        'make things pretty
        ws.Columns.AutoFit
                
'go on to the next worksheet
Next ws

End Sub
