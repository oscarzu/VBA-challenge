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

    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
    'Finding Highest and lowest values
    
     greatest = Application.WorksheetFunction.Max(Range("K:K"))
     Vh = Application.WorksheetFunction.Max(Range("L:L"))
     lowest = Application.WorksheetFunction.Min(Range("K:K"))
     
    Cells(3, 17).Value = lowest
    Cells(3, 17).NumberFormat = "0.00%"
    
    Cells(4, 17).Value = Vh
    Cells(4, 17).NumberFormat = "000,000"
    
    Cells(2, 17).Value = greatest
    Cells(2, 17).NumberFormat = "0.00%"
    
    LrSummary = Cells(Rows.Count, 9).End(xlUp).Row
    For h = 2 To LrSummary
        If Cells(h, 11).Value = greatest Then
            Cells(2, 16) = Cells(h, 9).Value
           Exit For
        End If
       Next h
    
    For k = 2 To LrSummary
         If Cells(k, 12) = Vh Then
                Cells(4, 16) = Cells(k, 9).Value
                Exit For
            End If
        Next k
        
      For l = 2 To LrSummary
         If Cells(l, 11) = lowest Then
                Cells(3, 16) = Cells(l, 9).Value
                Exit For
            End If
        Next l
          
End Sub
