Attribute VB_Name = "Challenge2"
'Challenge2:
' - Includes calculations of Greatest % increase, decrease and total volume
' - Includes adjustments to run on every wroksheet


Sub StockExtended2()

'------------------------------------
'CHALLENGE 2 - LOOP through all worksheets
'UPDATE Cells & Range WITH "ws" SUFFIX, SO IT WOULD WORK IN ALL WORKSHEETS
'------------------------------------

For Each ws In Worksheets

'APPLY a TAB to the CODE
    Dim Ticker As String
    Dim OpenPrice, ClosePrice, PriceChange As Double
    Dim PctChange As Double
    Dim StockVolume As Double
    StockVolume = 0
    
    
    Dim Column, TickerRow, PriceChangeRow, PctChangeRow, StockVolumeRow As Integer
    Column = 10
    TickerRow = 2
    PriceChangeRow = 2
    PctChangeRow = 2
    StockVolumeRow = 2
    
    
    'Headers
    ws.Cells(TickerRow - 1, Column).Value = "Ticker"
    ws.Cells(PriceChangeRow - 1, Column + 1).Value = "Price Change"
    ws.Cells(PctChangeRow - 1, Column + 2).Value = "% Change"
    ws.Cells(StockVolumeRow - 1, Column + 3).Value = "Stock Volume"
    
    
    ' Determine the Last Row
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
    'Initial Open Price
    OpenPrice = ws.Cells(2, 3).Value
    
    For i = 2 To last_row
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Ticker = ws.Cells(i, 1).Value
            ws.Cells(TickerRow, Column).Value = Ticker
            
            ClosePrice = ws.Cells(i, 6).Value
            PriceChange = ClosePrice - OpenPrice
            ws.Cells(PriceChangeRow, Column + 1).Value = PriceChange
            
                If PriceChange > 0 Then
                    ws.Cells(PriceChangeRow, Column + 1).Interior.ColorIndex = 4
                ElseIf PriceChange < 0 Then
                    ws.Cells(PriceChangeRow, Column + 1).Interior.ColorIndex = 3
                End If
            
        
            If OpenPrice = 0 Then
                PctChange = 0
                ws.Columns(Column + 2).NumberFormat = "#.##%"
            Else
                PctChange = ClosePrice / OpenPrice
                ws.Cells(PctChangeRow, Column + 2).Value = PctChange
                ws.Columns(Column + 2).NumberFormat = "#.##%"
            End If
            OpenPrice = ws.Cells(i + 1, 3).Value
            
            StockVolume = StockVolume + ws.Cells(i, 7).Value
            ws.Cells(StockVolumeRow, Column + 3).Value = StockVolume
    
    
    'New lines for the summary table
            TickerRow = TickerRow + 1
            PriceChangeRow = PriceChangeRow + 1
            PctChangeRow = PctChangeRow + 1
            StockVolumeRow = StockVolumeRow + 1
            
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
            StockVolume = StockVolume + ws.Cells(i, 7).Value
        
        End If
    Next i
    
    '------------------------------------
    'CHALLENGE 1 - Find Greatest Values
    '------------------------------------
       
    Dim GreatPctInc, GreatPctDec, GreatVol As Double
    
    'ChallengeTableHeaders
    ws.Cells(1, Column + 6).Value = "Ticker"
    ws.Cells(1, Column + 7).Value = "Value"
    
    'ChallengeTableRows
    ws.Cells(2, Column + 5).Value = "Greatest % Increase"
    ws.Cells(3, Column + 5).Value = "Greatest % Decrease"
    ws.Cells(4, Column + 5).Value = "Greatest Volume"
    
    
    ' Find greatest values
    GreatPctInc = WorksheetFunction.Max(ws.Range("L2:L" & last_row))
    GreatPctDec = WorksheetFunction.Min(ws.Range("L2:L" & last_row))
    GreatVol = WorksheetFunction.Max(ws.Range("M2:M" & last_row))
    
    
    ' Check if the Ticker matches the max/min value
    For k = 2 To last_row
        
        If ws.Cells(k, Column + 2).Value = GreatPctInc Then
            ws.Cells(2, Column + 6).Value = ws.Cells(k, Column).Value
            ws.Cells(2, Column + 7).Value = GreatPctInc
            ws.Cells(2, Column + 7).NumberFormat = "#.##%"
                
        ElseIf ws.Cells(k, Column + 2).Value = GreatPctDec Then
            ws.Cells(3, Column + 6).Value = ws.Cells(k, Column).Value
            ws.Cells(3, Column + 7).Value = GreatPctDec
            ws.Cells(3, Column + 7).NumberFormat = "#.##%"
                
        ElseIf ws.Cells(k, Column + 3).Value = GreatVol Then
            ws.Cells(4, Column + 6).Value = ws.Cells(k, Column).Value
            ws.Cells(4, Column + 7).Value = GreatVol
                
        End If
            
    Next k

'CLOSE the FOR EACH function
Next ws

End Sub


