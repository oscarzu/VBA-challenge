Sub StocksH()

    Dim V As LongLong
    Dim change As Double
    Dim startd As Long
    Dim Lr As Long
    Dim percentchange As Double
    Dim greatest As Double
    Dim h As Long
    Dim lowest As Double
    
    
    'Placing the headers

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"

    'Initial values
    j = 0
    V = 0
    change = 0
    startd = 2
    
    Lr = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To Lr
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        V = V + Cells(i, 7).Value
        
        If V = 0 Then
            
            Range("I" & j + 2).Value = Cells(i, 1).Value
            Range("J" & j + 2).Value = 0
            Range("K" & j + 2).Value = 0
            Range("L" & j + 2).Value = 0
            
            Else
            If Cells(startd, 3) = 0 Then
                For Findv = startd To i
                If Cells(Findv, 3).Value <> 0 Then
                startd = Findv
                Exit For
            End If
            Next Findv
        End If
    
    'Calculating the absolut variation
    change = (Cells(i, 6) - Cells(startd, 3))
    percentchange = change / Cells(startd, 3)
    
    startd = i + 1
    
        Range("I" & j + 2).Value = Cells(i, 1).Value
        Range("J" & j + 2).Value = change
        Range("K" & j + 2).Value = percentchange
        Range("K" & j + 2).NumberFormat = "0.00%"
        Range("L" & j + 2).Value = V
        Range("L" & j + 2).NumberFormat = "000,000"
        
            Select Case change
            Case Is > 0
             Range("J" & j + 2).Interior.Color = RGB(102, 255, 51)
            Case Is < 0
            Range("J" & j + 2).Interior.Color = RGB(255, 51, 0)
            Case Else
            Range("J" & j + 2).Interior.Color = RGB(255, 255, 255)
            End Select
    End If
   V = 0
   change = 0
   j = j + 1
   
    Else
    V = V + Cells(i, 7).Value
    End If
Next i
    
End Sub
