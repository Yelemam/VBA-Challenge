Sub MultipleYearStock()

'Do the following to all worksheets in the workbook
For Each ws In Worksheets
Dim i As Long
Dim StartDate As Date
Dim OpenValue As Double
Dim CloseValue As Double
Dim Volume As Double
Dim Ticker As String
Ticker = " "
Dim TickerCount As Long
TickerCount = 1
Dim QuarterlyChange As Double
Dim PercentageChange As Double
Dim LastRow As Long

LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

'loop through i=2 to lastRow

For i = 2 To LastRow
 If ws.Cells(i, 1).Value = Ticker Then
  Volume = Volume + ws.Cells(i, 7).Value
 
 Else
  Ticker = ws.Cells(i, 1).Value
  TickerCount = TickerCount + 1
  ws.Cells(TickerCount, 9).Value = Ticker
  Volume = 0 + ws.Cells(i, 7).Value
  OpenValue = ws.Cells(i, 3).Value
  StartDate = ws.Cells(i, 2).Value
  
 
 End If
 
 If ws.Cells(i + 1, 1).Value <> Ticker Or i = LastRow Then
  CloseValue = ws.Cells(i, 6).Value
  QuarterlyChange = CloseValue - OpenValue
  If OpenValue <> 0 Then
  PercentageChange = QuarterlyChange / OpenValue
 
 Else
 PercentageChange = 0
 End If
 PercentageChange = Application.WorksheetFunction.Round(PercentageChange, 4)
 
 ws.Cells(TickerCount, 10) = QuarterlyChange
 ws.Cells(TickerCount, 11) = PercentageChange
 ws.Cells(TickerCount, 12) = Volume
  
  'To fill background Color depending on the QuarterlyChange values
  If QuarterlyChange < 0 Then
    ws.Cells(TickerCount, 10).Interior.ColorIndex = 3
ElseIf QuarterlyChange > 0 Then
    ws.Cells(TickerCount, 10).Interior.ColorIndex = 4
ElseIf QuarterlyChange = 0 Then
    ws.Cells(TickerCount, 10).Interior.ColorIndex = xlColorIndexNone

  End If
End If
Next i

ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("L:L").NumberFormat = "0"
With ws.Range("I1:L1")
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    
End With
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "QuarterlyChange"
ws.Range("K1").Value = "PercentageChange"
ws.Range("L1").Value = "TotalVolume"

Dim DataEnd As Long
Dim Increase As Double
Dim Decrease As Double
Dim MaxVolume As LongLong
Dim IncreaseTicker As String
Dim DecTicker As String
Dim MaxVolTicker As String

' Find the last row with data in column 9
DataEnd = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row

' Set variables= 0
Increase = 0
Decrease = 0
MaxVolume = 0

' Loop through the data range from row 2 to DataEnd
For i = 2 To DataEnd
    ' Check for increase
    If Increase < ws.Cells(i, 11).Value Then
        Increase = ws.Cells(i, 11).Value
        IncreaseTicker = ws.Cells(i, 9).Value
    End If
    
    ' Check for decrease
    If Decrease > ws.Cells(i, 11).Value Then
        Decrease = ws.Cells(i, 11).Value
        DecTicker = ws.Cells(i, 9).Value
    End If
    
    ' Check for maximum volume
    If MaxVolume < ws.Cells(i, 12).Value Then
        MaxVolume = ws.Cells(i, 12).Value
        MaxVolTicker = ws.Cells(i, 9).Value
    End If
Next i
 
 ws.Range("P2").Value = IncreaseTicker
 ws.Range("Q2").Value = Increase
 ws.Range("P3").Value = DecTicker
 ws.Range("Q3").Value = Decrease
 ws.Range("P4").Value = MaxVolTicker
 ws.Range("Q4").Value = MaxVolume
 
 ws.Range("Q2:Q3").NumberFormat = "0.00%"
 ws.Range("Q4").NumberFormat = "0,000"
 
 ws.Range("O2").Value = "Greatest Percentage Increase"
 ws.Range("O3").Value = "Greatest Percentage Decrease"
 ws.Range("O4").Value = "Greatest Total Volume"
 ws.Range("P1").Value = "Ticker"
 ws.Range("Q1").Value = "Value"
 
 
 Next
 
End Sub