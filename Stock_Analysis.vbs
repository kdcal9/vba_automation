Attribute VB_Name = "Stock_Analysis"
Sub Stock_Analysis():

 For Each ws In Worksheets
 
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Volume"
 
 lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
 Dim ticker_symbol As String
 Dim stock_volume As Double
 stock_volume = 0
 Dim summaryTable_row As Integer
 summaryTable_row = 2
 
 Dim open_price As Double
 Dim close_price As Double
 Dim yearly_change As Double
 Dim percent_change As Double
 open_price = ws.Cells(2, 3).Value
 
 For i = 2 To lastRow
 
 If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
 ticker_symbol = ws.Cells(i, 1).Value
 stock_volume = stock_volume + ws.Cells(i, 7).Value
 close_price = ws.Cells(i, 6).Value
 yearly_change = close_price - open_price
 
 If open_price = 0 Then
 percent_change = 0
 Else: percent_change = yearly_change / open_price
 End If
 
 ws.Range("I" & summaryTable_row).Value = ticker_symbol
 ws.Range("J" & summaryTable_row).Value = yearly_change
 ws.Range("K" & summaryTable_row).Value = percent_change
 ws.Range("L" & summaryTable_row).Value = stock_volume
 
 ws.Range("K" & summaryTable_row).Style = "Percent"
 
 summaryTable_row = summaryTable_row + 1
 
 stock_volume = 0
 
 open_price = ws.Cells(i + 1, 3).Value
 Else
 
 stock_volume = stock_volume + ws.Cells(i, 7).Value

 
 End If
 
 Next i
 
 lastRow_summaryTable = ws.Cells(Rows.Count, 10).End(xlUp).Row
 

 For i = 2 To lastRow_summaryTable
 
 If ws.Cells(i, 10).Value > 0 Then
 ws.Cells(i, 10).Interior.ColorIndex = 4
 Else
 ws.Cells(i, 10).Interior.ColorIndex = 3
 End If
 
 Next i
 
 Next ws
 
End Sub
