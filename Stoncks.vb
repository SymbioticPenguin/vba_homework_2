Sub stoncks()

Dim lastrow, starting(6000), ending(6000) As Single
Dim count As Integer
Dim ws As Worksheet

'loop through all 3 sheets

For Each ws In ThisWorkbook.Sheets


starting(0) = ws.Range("C2").Value

count = 0


'assign headers to specific columns

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Stock Volume"



lastrow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row

For i = 2 To lastrow

If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then

'Pull ticker name

ws.Cells(count + 2, 9).Value = ws.Cells(i, 1).Value

'Add first open/last close stock price to array, then print yearly change

starting(count + 1) = ws.Cells(i + 1, 3).Value
ending(count) = ws.Cells(i, 6).Value

ws.Range("j" & (count + 2)).Value = ending(count) - starting(count)

'Calc the percentage change...

If (starting(count) = 0 And ending(count) < 0) Then
ws.Range("k" & (count + 2)).Value = -1
ElseIf (starting(count) = 0 And ending(count) > 0) Then
ws.Range("k" & (count + 2)).Value = 1
ElseIf (starting(count) = 0 And ending(count) = 0) Then
ws.Range("k" & (count + 2)).Value = 0
Else
ws.Range("k" & (count + 2)).Value = (ending(count) - starting(count)) / starting(count)
End If

'then color coat
If (ws.Range("j" & (count + 2)).Value > 0) Then
ws.Range("j" & (count + 2)).Interior.Color = RGB(0, 255, 0)
ElseIf (ws.Range("j" & (count + 2)).Value < 0) Then
ws.Range("j" & (count + 2)).Interior.Color = RGB(255, 0, 0)
Else
ws.Range("j" & (count + 2)).Interior.Color = RGB(255, 255, 0)
End If

'Do the final volume calc, then increment the count

ws.Range("l" & (count + 2)).Value = ws.Range("l" & (count + 2)).Value + ws.Cells(i, 7).Value
count = count + 1

Else

'Aggregate the stock volume until the ticker changes

ws.Range("l" & (count + 2)).Value = ws.Range("l" & (count + 2)).Value + ws.Cells(i, 7).Value

End If

Next i

'Formatting and autofitting

ws.Range("j2:j" & ws.Cells(ws.Rows.count, 10).End(xlUp).Row).NumberFormat = "0.00"
ws.Range("k2:k" & ws.Cells(ws.Rows.count, 11).End(xlUp).Row).NumberFormat = "0.00%"
ws.Columns("I:L").AutoFit



Next ws


End Sub