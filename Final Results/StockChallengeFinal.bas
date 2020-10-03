
Sub StockSummary():

Dim Ticker As String
Dim StartPrice As Double
Dim EndPrice As Double
Dim PriceChange As Double
Dim PercentChange As Double
Dim StockVolume As Double
Dim SummaryTableRow As Long
Dim StartRow As Long
Dim LastRow As Long
Dim SummaryLastRow As Long
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
ws.Activate

ws.Range("J1").Value = "Stock Ticker"
ws.Range("K1").Value = "Price Change"
ws.Range("L1").Value = "% Change"
ws.Range("M1").Value = "Stock Volume"

ws.Range("J1:M1").Font.Bold = True

StockVolume = 0
SummaryTableRow = 2
StartRow = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
StartPrice = 0
EndPrice = 0
PriceChange = 0
PercentChange = 0

For i = 2 To LastRow
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        StartPrice = ws.Cells(StartRow, 3).Value
        EndPrice = ws.Cells(i, 6).Value
        PriceChange = EndPrice - StartPrice
        
      
        
        If StartPrice = 0 And EndPrice = 0 Then
            PercentChange = PriceChange
        
        ElseIf StartPrice = 0 And EndPrice > 0 Then
            PercentChange = 0
        Else
        PercentChange = PriceChange / StartPrice
        End If
            
        StockVolume = StockVolume + ws.Cells(i, 7).Value
        
        ws.Range("J" & SummaryTableRow).Value = Ticker
        SummaryLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        ws.Range("K" & SummaryTableRow).Value = PriceChange
        
         For j = 2 To SummaryLastRow
         
            If ws.Cells(j, 11).Value > 0 Then
            ws.Cells(j, 11).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 11).Value < 0 Then
            ws.Cells(j, 11).Interior.ColorIndex = 3
            Else
            ws.Cells(j, 11).Interior.ColorIndex = 0
            End If
          Next j
        
        ws.Range("L" & SummaryTableRow).Value = PercentChange
        ws.Range("L" & SummaryTableRow).Value = FormatPercent(PercentChange, 2)
        ws.Range("M" & SummaryTableRow).Value = StockVolume
        
        SummaryTableRow = SummaryTableRow + 1
        StartRow = i + 1
        
        StartPrice = 0
        EndPrice = 0
        PriceChange = 0
        PercentChange = 0
        StockVolume = 0
        
    Else
    StockVolume = StockVolume + ws.Cells(i, 7).Value
    
    End If
    
    Next i
    
    Next ws
    
    
End Sub

Sub StockChallange():

Dim MaxPercentTicker As String
Dim MinPercentTicker As String
Dim MaxVolumeTicker As String
Dim MaxPercentValue As Double
Dim MinPercentValue As Double
Dim MaxStockVolume As Double
Dim SummaryLastRow As Long
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
ws.Activate

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Maximum % Change"
ws.Range("O3").Value = "Minimum % Change"
ws.Range("O4").Value = "Maximum Stock Volume"


ws.Range("P1:Q1").Font.Bold = True
ws.Range("O2:O4").Font.Bold = True

MaxPercentTicker = ws.Cells(2, 10).Value
MinPercentTicker = ws.Cells(2, 10).Value
MaxVolumeTicker = ws.Cells(2, 10).Value
MaxPercentValue = ws.Cells(2, 12).Value
MinPercentValue = ws.Cells(2, 12).Value
MaxStockVolume = ws.Cells(2, 13).Value
SummaryLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row


  For k = 2 To SummaryLastRow
  
        
        If ws.Cells(k, 12) > MaxPercentValue Then
        MaxPercentValue = ws.Cells(k, 12).Value
        MaxPercentTicker = ws.Cells(k, 10).Value

        ElseIf ws.Cells(k, 12) < MinPercentValue Then
        MinPercentValue = ws.Cells(k, 12).Value
        MinPercentTicker = ws.Cells(k, 10).Value
        End If

        If ws.Cells(k, 13) > MaxStockVolume Then
        MaxStockVolume = ws.Cells(k, 13).Value
        MaxVolumeTicker = ws.Cells(k, 10).Value
        

    End If
    
    
    ws.Range("P2").Value = MaxPercentTicker
    ws.Range("Q2").Value = MaxPercentValue
    ws.Range("Q2").Value = FormatPercent(MaxPercentValue, 2)
    ws.Range("P3").Value = MinPercentTicker
    ws.Range("Q3") = MinPercentValue
    ws.Range("Q3").Value = FormatPercent(MinPercentValue, 2)
    ws.Range("P4") = MaxVolumeTicker
    ws.Range("Q4") = MaxStockVolume
    
Next k
Next ws
    
End Sub

