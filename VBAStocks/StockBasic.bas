Attribute VB_Name = "StockBasic"
'Ex1 - Applies for just 1 worksheet
'


Sub Stock()

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

End Sub

