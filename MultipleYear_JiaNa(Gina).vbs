Sub alphabetical_testing()

For Each ws In Worksheets

'Setting variables

Dim TickerSymbol As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim TotalStockVolume As Double
Dim i As Double


Dim LastRow As Double
TotalStockVolume = 0
SumRow = 2

'Set the SumTable Column name
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volumn"
    
'Iteration Variables
'TotalStockVolume=TotalStockVolume+Cells(i,7).value
'SumRow=SumRow +1


'Find the last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Define the initial OpenPrice
OpenPrice = ws.Range("C2").Value

For i = 2 To LastRow

'Condition1 when i= last row of a Ticker symbol,  assign the value to TickerSymbol and TotalStockVolume

If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    TickerSymbol = ws.Cells(i, 1).Value
    ws.Range("I" & SumRow).Value = TickerSymbol
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
    ws.Range("L" & SumRow).Value = TotalStockVolume
    
  
    'Find the ClosePrice, do the calculation. Assign the value to Yearly Change and Percent Change.
    
    ClosePrice = ws.Cells(i, 6).Value
    YearlyChange = ClosePrice - OpenPrice
    ws.Range("J" & SumRow).Value = YearlyChange
    
    'Calculate PercentChange data
    If OpenPrice = 0 Then
    PercentChange = 0
    Else
    PercentChange = Round(YearlyChange / OpenPrice, 4)
    ws.Range("K" & SumRow) = PercentChange
    End If

    
    'Find the OpenPrice for next TickerSymbol
    OpenPrice = ws.Cells(i + 1, 3).Value
    
    'Set the SumRow and TotalStockVolume for the next SumRow calculation
    SumRow = SumRow + 1
    TotalStockVolume = 0
    
Else
    'Set the iteration for TotalStockVolume
    TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
   

End If

Next i


'Challenges


ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Find the last row in Sum Table
LastRowK = ws.Cells(Rows.Count, 11).End(xlUp).Row

'Find the Max and Min in Sum Talbe
RangeK = ws.Range("K2:K" & LastRowK).Value
MaxK = Application.WorksheetFunction.Max(RangeK)
ws.Range("Q2") = MaxK
MinK = Application.WorksheetFunction.Min(RangeK)
ws.Range("Q3") = MinK

RangeL = ws.Range("L2:L" & LastRowK).Value
MaxL = Application.WorksheetFunction.Max(RangeL)
ws.Range("Q4") = MaxL

For i = 2 To LastRowK

    If ws.Cells(i, 11).Value = MaxK Then
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("P2").Value = ws.Range("I" & i).Value
    
    ElseIf ws.Cells(i, 11).Value = MinK Then
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("P3").Value = ws.Range("I" & i).Value
    
    ElseIf ws.Cells(i, 12).Value = MaxL Then
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = ws.Range("I" & i).Value
    
    End If
    
Next i
    
'Formatting the SumTable
'Change the fill color
For i = 2 To LastRowK

    If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
    End If
    
Next i
    


    For j = 2 To LastRowK
        ws.Cells(j, 11).NumberFormat = "0.00%"
    Next j


'Set data format
ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next ws

End Sub
 



