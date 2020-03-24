Sub stock()

For Each ws In Worksheets
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim ticker As String

Dim Total As Double

Dim open_price As Double

Dim close_price As Double


Total = 0
open_price = 0
close_price = 0


Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

Range("J1").Value = "Ticker"
Range("K1").Value = "Yearly Change"
Range("L1").Value = "Percent Change"
Range("M1").Value = "Total Stock Volume"

For i = 2 To LastRow

 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 ticker = Cells(i, 1).Value
 open_price = open_price + Cells(i, 3).Value
 close_price = close_price + Cells(i, 6).Value
 Total = Total + Cells(i, 7).Value

 Range("J" & Summary_Table_Row).Value = ticker
 Range("K" & Summary_Table_Row).Value = close_price - open_price
 Range("L" & Summary_Table_Row).Value = (close_price - open_price) / open_price
 Range("M" & Summary_Table_Row).Value = Total
Summary_Table_Row = Summary_Table_Row + 1

Total = 0
open_price = 0
close_price = 0

Else
Total = Total + Cells(i, 7).Value
open_price = open_price + Cells(i, 3).Value
close_price = close_price + Cells(i, 6).Value
    End If

  Next i
  Range("L2:L" & LastRow).NumberFormat = "0.00%"
  
  LastRow_summary = ws.Cells(Rows.Count, 12).End(xlUp).Row
  For j = 2 To LastRow_summary

    If Cells(j, 12) > 0 Then
    Cells(j, 12).Interior.ColorIndex = 4
    
    Else
    Cells(j, 12).Interior.ColorIndex = 3

    End If
 Next j
Next ws
End Sub
