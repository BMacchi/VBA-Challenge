Sub StockChecker()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

'Variables
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentageChange As Double
    Dim TotalVolume As Double
    Dim j As Long
    Dim GreatestTicker As String
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestTotalVolume As Double
    j = 2
    OpenPrice = Range("C2").Value
    
    GreatestIncrease = 0
    GreatestDecrease = 9999
    GreatestTotalVolume = 0
        
'Loop thru all data
        For I = 2 To 753001
        TotalVolume = TotalVolume + Cells(I, 7).Value
    
'Check if we are still within the same ticker, if not...
        If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
        ClosePrice = Cells(I, 6).Value
        Ticker = Cells(I, 1).Value

'List the Ticker name in column I
        Range("I" & j).Value = Ticker
    
'List YearlyChange in Yearly Change Column
        YearlyChange = (ClosePrice - OpenPrice)
        Range("J" & j).Value = YearlyChange
    
'Color code negative and positive yearly change
        If YearlyChange >= 0 Then
            Cells(j, 10).Interior.Color = vbGreen
        Else
            Cells(j, 10).Interior.Color = vbRed
        End If

'List Percentage Change
    PercentageChange = (ClosePrice - OpenPrice) / OpenPrice * 100
    Range("K" & j).Value = PercentageChange
    OpenPrice = Cells(I + 1, 3).Value

'Color code negative and positive Percentage change
    If PercentageChange >= 0 Then
        Cells(j, 11).Interior.Color = vbGreen
    Else
        Cells(j, 11).Interior.Color = vbRed
    End If

'List Greatest Percentage Increase
    If PercentageChange > GreatestIncrease Then
        GreatestIncrease = PercentageChange
        GreatestTicker = Cells(j, 9).Value
        Range("Q" & 2).Value = GreatestIncrease
        Range("P" & 2).Value = GreatestTicker
    End If
        
'List Greatest Percentage Decrease
    If PercentageChange < GreatestDecrease Then
        GreatestDecrease = PercentageChange
        GreatestTicker = Cells(j, 9).Value
        Range("Q" & 3).Value = GreatestDecrease
        Range("P" & 3).Value = GreatestTicker
    End If
        
' List Greatest Total Volume
    If TotalVolume > GreatestTotalVolume Then
        GreatestTotalVolume = TotalVolume
        GreatestTicker = Ticker
        Range("Q" & 4).Value = GreatestTotalVolume
        Range("P" & 4).Value = GreatestTicker
    End If
            
'List Total Stock Volume
    Range("L" & j).Value = TotalVolume
    TotalVolume = 0
    
    j = j + 1
    End If

    Next I
    
Next ws

End Sub
