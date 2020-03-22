Sub stock()

For Each ws in Worksheets
 LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim ticker As String

Dim Total As Double
Total=0

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

For i = 2 to LastRow

 If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
 ticker = Cells(i, 1).Value
 Total = Total + Cells(i, 7).Value
 Range("J" & Summary_Table_Row).Value = ticker
 Range("K" & Summary_Table_Row).Value = Total
Summary_Table_Row = Summary_Table_Row + 1
Total = 0
Else
Total = Total + Cells(i, 7).Value

    End If

  Next i
Next ws
End Sub