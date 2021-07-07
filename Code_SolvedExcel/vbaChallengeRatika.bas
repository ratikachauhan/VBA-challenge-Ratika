Attribute VB_Name = "Module1"
Sub VBAChallenge()

For Each ws In Worksheets

'Define Varibales

Dim TickerName As String
Dim TotalStockVolume As Double
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim YearlyChange As Double
Dim PercentChange As Double
'Hard Solution
Dim MaxTickerName As String
Dim MinTickerName As String
Dim MaxPercentChange As Double
Dim MinPercentChange As Double
Dim MaxTotalVolumnTicker As String
Dim MaxTotalVolume As Double


'Set Initial Values

TickerName = " "
TotalStockVolume = 0
OpenPrice = 0
ClosePrice = 0
YearlyChange = 0
PercentChange = 0
'Hard Solution
MaxTickerName = " "
MinTickerName = " "
MaxPercentChange = 0
MinPercentChange = 0
MaxTotalVolumnTicker = " "
MaxTotalVolumn = 0


'Summary Table

Dim SummaryTableRow As Long
SummaryTableRow = 2

'Print Summary Column Names
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
     ' Hard
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Last Row

Dim LastRow As Long
Dim i As Long



LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
OpenPrice = ws.Cells(2, 3).Value

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        TickerName = ws.Cells(i, 1).Value
        ClosePrice = ws.Cells(i, 6).Value
        YearlyChange = ClosePrice - OpenPrice
    
        'Condition for Percent Change Div by zero
        'PercentChange = (YearlyChange / OpenPrice) * 100
            If OpenPrice <> 0 Then
                PercentChange = (YearlyChange / OpenPrice) * 100
            Else
                PercentChange = 0
            End If
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    
        'Print Summary Table
    
        ws.Range("I" & SummaryTableRow).Value = TickerName
        ws.Range("J" & SummaryTableRow).Value = YearlyChange
        'Color YearlyChange
        If (PercentChange > 0) Then
            'Fill column with GREEN
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        ElseIf (PercentChange <= 0) Then
            'Fill column with RED
            ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
        End If
    
        ws.Range("K" & SummaryTableRow).Value = (CStr(PercentChange) & "%")
        ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
    
        SummaryTableRow = SummaryTableRow + 1
        YearlyChange = 0
        ClosePrice = 0
        OpenPrice = ws.Cells(i + 1, 3).Value
        
        'Hard
        
        If (PercentChange > MaxPercentChange) Then
        MaxPercentChange = PercentChange
        MaxTickerName = TickerName
        ElseIf (PercentChange < MinPercentChange) Then
        MinPercentChange = PercentChange
        MinTickerName = TickerName
        End If
        
        If (TotalStockVolume >= MaxTotalVolumn) Then
        MaxTotalVolumnTicker = TickerName
        MaxTotalVolumn = TotalStockVolume
        End If
        
        PercentChange = 0
        TotalStockVolume = 0
    
    Else
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    End If

Next i
    
    ws.Range("P2").Value = MaxTickerName
    ws.Range("P3").Value = MinTickerName
    ws.Range("P4").Value = MaxTotalVolumnTicker
    ws.Range("Q2").Value = (CStr(MaxPercentChange) & "%")
    ws.Range("Q3").Value = (CStr(MinPercentChange) & "%")
    ws.Range("Q4").Value = MaxTotalVolumn
    
    


Next ws

End Sub



