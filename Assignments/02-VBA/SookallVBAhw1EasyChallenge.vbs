Sub vbahweasychallenge()

For Each ws In ActiveWorkbook.Worksheets
         
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total Stock Volume"

Dim I As Double
' value to hold ticker number
Dim ticker As String
Dim stockvolume As Double
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' loop
For I = 2 To lastrow
' compare current ticker with next to see if there's a difference
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
' if different, set ticker name value to the current
      ticker = ws.Cells(I, 1).Value
' Add to the stock volume
      stockvolume = stockvolume + ws.Cells(I, 7).Value
' Print ticker number in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker
' Print stock volume to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = stockvolume
' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
' Reset the Brand Total
      stockvolume = 0
      
' If the cell immediately following a row is the same brand...
    Else
' Add to the stock volume
      stockvolume = stockvolume + ws.Cells(I, 7).Value
      
    End If

Next I
ws.Columns("A:Q").AutoFit
Next ws
End Sub


