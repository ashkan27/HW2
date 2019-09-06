Sub thewolfofwallstreet():

'Dim ws As Worksheet

For Each ws In Worksheets


Dim T As String
Dim a As Integer
Dim temp As String
Dim op As Double
Dim cl As Double
Dim clo As Double
Dim LR As Long
Dim k As Long
Dim test As Long
Dim max As Double
Dim min As Double
Dim maxn As Integer
Dim minn As Integer
Dim TotalV As Variant
Dim Totaln As Integer

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"
ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"


LR = ws.Cells(Rows.Count, 1).End(xlUp).Row

T = ws.Cells(2, 1)
ws.Range("I2") = T
a = 3
op = ws.Cells(2, 3)
clo = ws.Cells(LR, 6)
k = 2


For i = 2 To LR
    
    temp = ws.Cells(i, 1)
    
    If T <> temp Then
    
    ws.Cells(a, 9) = temp
    cl = ws.Cells(i - 1, 6)
    ws.Cells(a - 1, 10) = cl - op
    
    If op = 0 Then
    ws.Cells(a - 1, 11) = Format(0, "Percent")
    
    Else
    
    ws.Cells(a - 1, 11) = Format(((cl - op) / op), "Percent")
    End If
    If ws.Cells(a - 1, 10) > 0 Then
    
    ws.Cells(a - 1, 10).Interior.ColorIndex = 4
    
    Else
    
    ws.Cells(a - 1, 10).Interior.ColorIndex = 3
    
    End If
    
    If ws.Cells(a - 1, 11) > 0 Then
    
    ws.Cells(a - 1, 11).Interior.ColorIndex = 4
    
    Else
    
    ws.Cells(a - 1, 11).Interior.ColorIndex = 3
    
    End If
    
    

    
    ws.Cells(a - 1, 12) = Application.Sum(Range(ws.Cells(k, 7), ws.Cells(i - 1, 7)))
    
    k = i
    op = ws.Cells(i, 3)
    T = temp
    a = a + 1
    
    End If
    
Next i

ws.Cells(a - 1, 10) = clo - op

If ws.Cells(a - 1, 10) > 0 Then
    
    ws.Cells(a - 1, 10).Interior.ColorIndex = 4
    
    Else
    
    ws.Cells(a - 1, 10).Interior.ColorIndex = 3
    
    End If
    
ws.Cells(a - 1, 11) = Format(((clo - op) / op), "Percent")

If ws.Cells(a - 1, 11) > 0 Then
    
    ws.Cells(a - 1, 11).Interior.ColorIndex = 4
    
    Else
    
    ws.Cells(a - 1, 11).Interior.ColorIndex = 3
    
    End If


ws.Cells(a - 1, 12) = Application.Sum(Range(ws.Cells(k, 7), ws.Cells(LR, 7)))

max = 0
For j = 2 To a - 1
If max < ws.Cells(j, 11) Then

max = ws.Cells(j, 11)
maxn = j
End If

Next j

ws.Range("Q2") = Format(max, "Percent")
ws.Range("P2") = ws.Cells(maxn, 9)

min = 0
For m = 2 To a - 1
If min > ws.Cells(m, 11) Then

min = ws.Cells(m, 11)
minn = m

End If

Next m

ws.Range("Q3") = Format(min, "Percent")
ws.Range("P3") = ws.Cells(minn, 9)

TotalV = ws.Cells(2, 12)

For n = 2 To a - 1
If TotalV < ws.Cells(n, 12) Then

TotalV = ws.Cells(n, 12)
Totaln = n

End If
Next n

ws.Range("Q4") = TotalV
ws.Range("P4") = ws.Cells(Totaln, 9)


Next ws
End Sub

