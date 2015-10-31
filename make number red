Attribute VB_Name = "Ä£¿é1"
Sub test()

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim t As String
Dim s As Integer

s = InputBox("Êý¾Ý´ÓexcelµÚ¼¸ÁÐ¿ªÊ¼£¿")

For i = 1 To ActiveSheet.UsedRange.Rows.count
 For j = s To ActiveSheet.UsedRange.Columns.count 'input¾ö¶¨
  For k = 2 To Len(Cells(i, j))
   t = Mid(Cells(i, j), k, 1)
   If InStr(Cells(i, j), "1") <> 0 Or InStr(Cells(i, j), "2") <> 0 Or InStr(Cells(i, j), "3") <> 0 Or InStr(Cells(i, j), "4") <> 0 Or InStr(Cells(i, j), "5") <> 0 Or InStr(Cells(i, j), "6") <> 0 Or InStr(Cells(i, j), "7") <> 0 Or InStr(Cells(i, j), "8") <> 0 Or InStr(Cells(i, j), "9") <> 0 Or InStr(Cells(i, j), "0") <> 0 Then
    If t <> "1" And t <> "2" And t <> "3" And t <> "4" And t <> "5" And t <> "6" And t <> "7" And t <> "8" And t <> "9" And t <> "0" And t <> "." And t <> "%" Then
     Cells(i, j).Select
     With Selection.Characters(Start:=k, Length:=1).Font
     .Color = -16776961
     End With
    End If
   End If
  Next k
 Next j
Next i
End Sub
