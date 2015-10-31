Sub tset()

Dim i As Integer
Dim j As Integer
Dim n As String
Dim k As String

For i = 2 To 75

    For j = 1 To ActiveSheet.UsedRange.Rows.Count

    If Cells(i, 1) = Cells(j, 2) Then
    Cells(i, 3) = "A" & j
    k = Cells(i, 1).Text
    n = "Header01!" & "A" & j
    
    Cells(i, 1).Select
    ActiveSheet.Hyperlinks.add Anchor:=Selection, Address:="", SubAddress:=n, TextToDisplay:=k
    
    End If
    
    Next j

Next i

End Sub
