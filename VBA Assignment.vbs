Sub Testing()
    'Headers
    'ws.Range("I1").Value = "Ticker"'
    'ws.Range("J1").Value = "Yearly Change"
    'ws.Range("K1").Value = "Percent Change"
    'ws.Range("L1").Value = "Total Stock Volume
    Dim ws As Worksheet
For Each ws In Worksheets
            Dim i As Long
            Dim ticker As String
            Dim close_price As Double
            Dim open_price As Double
            Dim Change As Double
            Dim volume As Double
            Dim Percent As Double
            Dim Endrow As Long
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            volume = 0
            Dim table As Long
            table = 2
            Dim yearly As Long
            Dim open_row As Long
            open_row = 2
            
            RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
            
            'looping ticker name
            For i = 2 To RowCount
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    volume = volume + ws.Cells(i, 7).Value
                    'yearly change
                     close_price = ws.Cells(i, 6).Value
                     open_price = ws.Cells(open_row, 3).Value
                     Change = close_price - open_price
                     Percent = Change / open_price
                     open_row = i + 1
                 'print ticker
                 ws.Range("I" & table).Value = ticker
                 ws.Range("J" & table).Value = Change
                 ws.Range("J" & table).NumberFormat = "0.00"
                 ws.Range("K" & table).Value = Percent
                 'ws.Range("K" & table).Style = "Percent"
                 ws.Range("K" & table).NumberFormat = "0.00%"
                 ws.Range("L" & table).Value = volume
                 'print volume
                 
                If ws.Range("J" & table).Value >= 0 Then
                    ws.Range("J" & table).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & table).Interior.ColorIndex = 3
                End If
                table = table + 1
                volume = 0
                Change = 0
                Else
                volume = volume + ws.Cells(i, 7).Value
                End If
              Next i
             
        '% increase/decrease
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Dim GreatestIncrease As Double
        'Dim GreatestDecrease As Double
        'Dim GreatestVolume As Double
        
        Endrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & Endrow))
        ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & Endrow))
        ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & Endrow))
        
        GreatestIncrease = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Endrow)), ws.Range("K2:K" & Endrow), 0)
        GreatestDecrease = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Endrow)), ws.Range("K2:K" & Endrow), 0)
        GreatestVolume = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Endrow)), ws.Range("L2:L" & Endrow), 0)
        
        
        ws.Range("P2").Value = ws.Cells(GreatestIncrease + 1, 9).Value
        ws.Range("P3").Value = ws.Cells(GreatestDecrease + 1, 9).Value
        ws.Range("P4").Value = ws.Cells(GreatestVolume + 1, 9).Value
          
        
    
   
      
Next ws
End Sub