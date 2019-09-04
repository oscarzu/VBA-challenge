
Sub Stocks()
Dim Lr As Long
Dim V As LongLong
Dim Ticker As String
Dim Summ_table As Long
Dim Headers(3) As String


Headers(0) = "Ticker"
Headers(1) = "Yearly Change"
Headers(2) = "Percentage Change"
Headers(3) = "Total Stock Volume"
V = 0
Summ_table = 2

Range("I1:L1").Value = Headers()

Lr = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To Lr

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker = Cells(i, 1).Value
V = V + Cells(i, 7)
mindate = Application.WorksheetFunction.Min(ActiveSheet.Range("B:B"))
maxdate = Application.WorksheetFunction.Max(ActiveSheet.Range("B:B"))
Range("I" & Summ_table).Value = Ticker
Range("L" & Summ_table).Value = V
Range("L" & Summ_table).NumberFormat = "000,000"
Summ_table = Summ_table + 1
V = 0
Else

V = V + Cells(i, 7)

End If
Next i
End Sub

