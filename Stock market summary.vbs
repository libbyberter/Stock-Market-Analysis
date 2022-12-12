Sub stocks()

Dim ws As Worksheet     'worksheet name
Dim r As Double         'counter for rows to loop thru in raw data
Dim s As Double         'start price
Dim e As Double         'end price
Dim x As Integer        'counter for rows in summary table
Dim t As Double         'total volume
Dim year As Range       'range of cells for conditional formatting
Dim last As Double      'last row for loops
Dim gi As Double        'greatest increase
Dim gd As Double        'greatest decrease
Dim ticker_i As String  'ticker name of greatest increase
Dim ticker_d As String  'ticker name of greatest decrease
Dim ticker_v As String  'ticker name of greatest volume
Dim vol As Double       'greatest volume



For Each ws In Worksheets

'MsgBox ws.Name
    last = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Sort data by ticker & date to ensure tickers are all grouped together
'and the dates are in ascending order to ensure the looping works correctly

    ws.Range("A1:G" & last).Sort _
    Key1:=ws.Range("A1:A" & last), Order1:=xlAscending, _
    Key2:=ws.Range("B1:B" & last), Order2:=xlAscending, _
    Header:=xlYes
   
'Create new columns for ticker, yearly change, % change, and volume
    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    
'Create new columns for greatest increase & decrease and greatest volume
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
'Format new columns to make them easy to read & in correct number formats
    ws.Columns("J:M").HorizontalAlignment = xlCenter
    ws.Columns("J:M").AutoFit
    ws.Range("J1:M1").Borders(xlEdgeBottom).Weight = xlThick
    ws.Columns("K").NumberFormat = "#,##0.00"
    ws.Columns("L").NumberFormat = "##0.00%"
    ws.Columns("M").NumberFormat = "#,##0"
    
    ws.Range("P1:Q1").Borders(xlEdgeBottom).Weight = xlThick
    ws.Range("Q2:Q3").NumberFormat = "##0.00%"
    ws.Cells(4, 17).NumberFormat = "#,##0"
    
'Set initial values for counters (x=row in table; s=start price; t=total of volume)
    x = 2
    s = Cells(2, 3).Value
    t = 0


'loop through rows & create summary table
    last = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To last
        
        
        If ws.Cells(r + 1, 1).Value = ws.Cells(r, 1) Then
            
            t = t + ws.Cells(r, 7).Value
            
        
        ElseIf ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1) Then
            
            t = t + ws.Cells(r, 7).Value
            e = ws.Cells(r, 6).Value
            ws.Cells(x, 10).Value = ws.Cells(r, 1).Value
            
            
            ws.Cells(x, 11).Value = e - s
            ws.Cells(x, 12).Value = (e - s) / s
            ws.Cells(x, 13).Value = t
            
            x = x + 1
            s = ws.Cells(r + 1, 3).Value
            t = 0
        
        End If
    
    Next r
    
    
'conditional formating for yearly change values ($ & %)
    last = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To last
    
        For y = 11 To 12
            Set year = ws.Cells(i, y)
            'make cells red when <0
            year.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
            year.FormatConditions(1).Interior.Color = vbRed
            
            'make cells green when >0
            year.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
            year.FormatConditions(2).Interior.Color = vbGreen
        Next y
        
    Next i
    
        
'gets largest increase, decrease, volume & creates table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ws.Range("P1:Q1").Borders(xlEdgeBottom).Weight = xlThick
    ws.Range("Q2:Q3").NumberFormat = "##0.00%"
    ws.Cells(4, 17).NumberFormat = "0.00E+00"

    gi = 0
    
    gd = 0
    
    vol = 0
    
    last = ws.Cells(Rows.Count, 10).End(xlUp).Row

    'greatest increase & decrease loop
    For i = 2 To last
    
        If ws.Cells(i, 12) >= 0 Then
            
            If ws.Cells(i, 12) > gi Then
                
                gi = ws.Cells(i, 12).Value
                ticker_i = ws.Cells(i, 10).Value
            
            End If
            
        ElseIf ws.Cells(i, 12) < 0 Then
            
            If Abs(ws.Cells(i, 12).Value) > gd Then
            
                gd = Abs(ws.Cells(i, 12).Value)
                ticker_d = ws.Cells(i, 10).Value
            
            End If
                
            
        End If
        
    Next i
    
    ' get greatest volume
    For i = 2 To last
    
        If ws.Cells(i, 13) >= vol Then
                
                vol = ws.Cells(i, 13).Value
                ticker_v = ws.Cells(i, 10).Value
            
        End If
        
    Next i
    
    
    'data in table
    ws.Cells(2, 16).Value = ticker_i
    ws.Cells(2, 17).Value = gi
    ws.Cells(3, 16).Value = ticker_d
    ws.Cells(3, 17).Value = gd * (-1)
    ws.Cells(4, 16).Value = ticker_v
    ws.Cells(4, 17).Value = vol
    
    'format table
    ws.Columns("P:Q").HorizontalAlignment = xlCenter
    ws.Columns("O:Q").AutoFit
    ws.Cells(4, 17).NumberFormat = "0.00E+00"

        
        
Next

End Sub
