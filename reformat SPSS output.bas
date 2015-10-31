Attribute VB_Name = "Ä£¿é111"
Sub ÃÀ»¯SPSS½»²æ±í()
Attribute ÃÀ»¯SPSS½»²æ±í.VB_ProcData.VB_Invoke_Func = " \n14"

'ÊÊÓÃÇ°Ìá£ºSPSSµÄOUTPUTÇ°ÈýÁÐ¸ñÊ½ÓÐÒªÇó£º
'1£©Ô¤ÉèºÃ±íÍ·µÄ×ÖÄ¸
'2£©Êý¾ÝÏÈround³É3Î»Ð¡Êý£¬ºóÆÚ°Ù·Ö±ÈÊý¾Ý±ãÖ±½Ó±£ÁôÒ»Î»
'3£©µÚÒ»ÁÐÌâ¸É
'4£©µÚ¶þÁÐÑ¡Ïî£«Total±êÊ¶
'5£©µÚÈýÁÐColumn%»òCount±êÊ¶
'6£©µÚËÄÁÐÆðÎªÊý¾Ý
'7£©TotalÐè·ÅÔÚÃ¿ÕÅ±íÏÂ·½



'££££µÚÒ»²½£¬Ôö´óÌâÖ®¼ä¼ä¸ô££££

Dim row1 As Integer
Dim row2 As Integer
Dim count As Integer
Dim i As Integer
Dim Myselect As Range

count = WorksheetFunction.CountA([A:A])
Range("A1").Select
For i = 1 To count 'Êµ¼ÊÇé¿öÓÐ¿ÉÄÜÊÇcount-1»ò+1
 Selection.End(xlDown).Select
 Selection.EntireRow.Insert
 Selection.EntireRow.Insert
 Selection.EntireRow.Insert
 row1 = ActiveCell.Row
 row2 = row1 + 3
 Cells(row1, 2) = Cells(row2, 1)
 Cells(row2, 1) = ""
 Selection.EntireRow.Insert
 Selection.EntireRow.Insert
 Selection.EntireRow.Insert
Next i


'££££µÚ¶þ²½Total/BaseÍùÉÏÒÆ££££

For i = 1 To ActiveSheet.UsedRange.Rows.count
 If Cells(i, 2) = "Total" Then
    Cells(i, 2).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Cut
    Selection.End(xlUp).Select
    Selection.Offset(-1, 0).Select
    Selection.Insert Shift:=xlDown
 End If
Next i

'££££µÚÈý²½£¬¼Ó°Ù·Ö±È·ûºÅ££££
For i = 1 To ActiveSheet.UsedRange.Rows.count
 If Cells(i, 2) = "Total" Then
  Range(Cells(i + 1, 4), Cells(i + 1, ActiveSheet.UsedRange.Columns.count + 1)).Select
  Selection.FormulaR1C1 = "%"
 End If
Next i


'££££µÚËÄ²½ÉèÖÃ¸ñÊ½££££
For i = 1 To ActiveSheet.UsedRange.Rows.count
 If Cells(i, 4) <> "" Then
   Cells(i, 4).Select
   Range(Selection, Selection.End(xlToRight)).Select
   If InStr(Cells(i, 3), "%") <> 0 Then
     Selection.NumberFormatLocal = "0.00%"
     Selection.Replace What:="%", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
     'Selection.NumberFormatLocal = "General"
       If Cells(i, 2) = "" Then Cells(i, 2) = "Response"
   'Else
    'Selection.NumberFormatLocal = "General"
   End If
 End If
Next i


'££££µÚÎå²½¼ÓÊÂÏÈµ÷ºÃµÄ±íÍ·££££
Set Myselect = Application.InputBox("select header region", Type:=8)
 If Myselect Is Nothing Then Exit Sub
 
 For i = 1 To ActiveSheet.UsedRange.Rows.count + count * Myselect.Rows.count
  If Cells(i, 2) = "Total" Then
    Myselect.Copy
    Cells(i, 2).Select
    Selection.Insert Shift:=xlDown
    i = i + Myselect.Rows.count
  End If
 Next i

'££££µÚÁù²½ResponseÇ°²åÈë---££££
For i = 1 To ActiveSheet.UsedRange.Rows.count + count
 If Cells(i, 2) = "Response" Then
    Cells(i, 2).Select
    Selection.EntireRow.Insert
    Range(Cells(i, 4), Cells(i, ActiveSheet.UsedRange.Columns.count + 1)).Select
    Selection.FormulaR1C1 = "---"
    i = i + 2
 End If
Next i

'Ö®ºó»¹ÐèÊÖ¶¯¾ÓÖÐ£¬É¾³ýµÚÒ»µÚÈýÁÐ£¬Total¸ÄBase£¬Resonse¸ÄTotal£¬0¸Ä£­
'¼ìÑé×ÖÄ¸ÊÖ¶¯Ìí¼Ó£¬²¢ÔÚ±í¸ñÎ²²¿Ìí¼ÓËµÃ÷

End Sub
