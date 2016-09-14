Sub CleanUp()
'
' CleanUp Macro
' Reformats the checkrun export (My first ever stab at a VBA Script for excel)
'
' Keyboard Shortcut: Ctrl+r
'
Application.ScreenUpdating = False

    Cells.Select
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A2:L90")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("B5").Select
    ActiveWindow.SmallScroll Down:=-21
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("C:C").Select
    Selection.EntireColumn.Hidden = True
    Columns("E:G").Select
    Selection.EntireColumn.Hidden = True
    Columns("I:L").Select
    Selection.EntireColumn.Hidden = True



For i = 1 To 117
  If ActiveSheet.Cells(i, 1) = "" Then
    ActiveSheet.Cells(i, 1).EntireRow.Hidden = True
End If
Next i

For i = 1 To 117
Dim j As Integer
j = i + 1
Dim r As Integer
r = 4
If ActiveSheet.Cells(i, 1) = ActiveSheet.Cells(j, 1) Then
    ActiveSheet.Cells(i, 1).Select
    ActiveSheet.Cells(i, r).Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveSheet.Paste
    ActiveCell.Offset(-1, 0).Rows("1:1").EntireRow.Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlUp
End If

Next i

Application.ScreenUpdating = True

End Sub
