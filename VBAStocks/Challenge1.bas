Attribute VB_Name = "Challenge1"
'Challenge1 - Includes calculations of Greatest % increase, decrease and total volume


Sub StockExtended1()

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
Cells(TickerRow - 1, Column).Value = "Ticker"
Cells(PriceChangeRow - 1, Column + 1).Value = "Price Change"
Cells(PctChangeRow - 1, Column + 2).Value = "% Change"
Cells(StockVolumeRow - 1, Column + 3).Value = "Stock Volume"


' Determine the Last Row
last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
'Initial Open Price
OpenPrice = Cells(2, 3).Value

For i = 2 To last_row

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        Ticker = Cells(i, 1).Value
        Cells(TickerRow, Column).Value = Ticker
        
        ClosePrice = Cells(i, 6).Value
        PriceChange = ClosePrice - OpenPrice
        Cells(PriceChangeRow, Column + 1).Value = PriceChange
        
            If PriceChange > 0 Then
                Cells(PriceChangeRow, Column + 1).Interior.ColorIndex = 4
            ElseIf PriceChange < 0 Then
                Cells(PriceChangeRow, Column + 1).Interior.ColorIndex = 3
            End If
        
        If OpenPrice = 0 Then
            PctChange = 0
            Columns("L").NumberFormat = "#.##%"
        Else
            PctChange = ClosePrice / OpenPrice
            Cells(PctChangeRow, Column + 2).Value = PctChange
            Columns("L").NumberFormat = "#.##%"
        End If
        OpenPrice = Cells(i + 1, 3).Value
        
        StockVolume = StockVolume + Cells(i, 7).Value
        Cells(StockVolumeRow, Column + 3).Value = StockVolume


'New lines for the summary table
        TickerRow = TickerRow + 1
        PriceChangeRow = PriceChangeRow + 1
        PctChangeRow = PctChangeRow + 1
        StockVolumeRow = StockVolumeRow + 1
        
    ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
        StockVolume = StockVolume + Cells(i, 7).Value
    
    End If
Next i

'------------------------------------
'CHALLENGE 1 - Find Greatest Values
'------------------------------------

Dim GreatPctInc, GreatPctDec, GreatVol As Double


'ChallengeTableHeaders
Cells(1, Column + 6).Value = "Ticker"
Cells(1, Column + 7).Value = "Value"

'ChallengeTableRows
Cells(2, Column + 5).Value = "Greatest % Increase"
Cells(3, Column + 5).Value = "Greatest % Decrease"
Cells(4, Column + 5).Value = "Greatest Volume"


' Find greatest values
GreatPctInc = WorksheetFunction.Max(Range("L2:L" & last_row))
GreatPctDec = WorksheetFunction.Min(Range("L2:L" & last_row))
GreatVol = WorksheetFunction.Max(Range("M2:M" & last_row))


' Check if the Ticker matches the max/min value
For k = 2 To last_row
    
    If Cells(k, Column + 2).Value = GreatPctInc Then
        Cells(2, Column + 6).Value = Cells(k, Column).Value
        Cells(2, Column + 7).Value = GreatPctInc
        Cells(2, Column + 7).NumberFormat = "#.##%"
            
    ElseIf Cells(k, Column + 2).Value = GreatPctDec Then
        Cells(3, Column + 6).Value = Cells(k, Column).Value
        Cells(3, Column + 7).Value = GreatPctDec
        Cells(3, Column + 7).NumberFormat = "#.##%"
            
    ElseIf Cells(k, Column + 3).Value = GreatVol Then
        Cells(4, Column + 6).Value = Cells(k, Column).Value
        Cells(4, Column + 7).Value = GreatVol
            
    End If
        
Next k

End Sub

