Sub tset()

Dim i As Integer
Dim j As Integer
Dim n As String
Dim k As String

For i = 1 To ActiveSheet.UsedRange.Rows.Count

    For j = 1 To 75

    If Cells(i, 2) = Cells(j, 1) And Cells(i, 2) <> "" Then
    k = Cells(i, 2).Text
    n = "TOC!" & "A" & j
    
    Cells(i, 2).Select
    ActiveSheet.Hyperlinks.add Anchor:=Selection, Address:="", SubAddress:=n, TextToDisplay:=k
    
    End If
    
    Next j

Next i

End Sub
