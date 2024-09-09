Attribute VB_Name = "m4_get_depts_number"
Function GetDeptsNumber(диапазон1 As Range, диапазон2 As Range, диапазон3 As Range, диапазон4 As Range) As Integer

    Dim i As Long
    Dim count1 As Long, count2 As Long, count3 As Long, count4 As Long
    Dim continuousCount As Long

    ' Проверка количества ячеек в диапазонах
    'MsgBox диапазон1.Cells.Count & ", " & диапазон2.Cells.Count & ", " & диапазон3.Cells.Count & ", " & диапазон4.Cells.Count
    If диапазон1.Cells.Count <> диапазон2.Cells.Count Or диапазон1.Cells.Count <> диапазон3.Cells.Count Or диапазон1.Cells.Count <> диапазон4.Cells.Count Then
        GetDeptsNumber = -1
        Exit Function
        'GoTo finish
    End If

    ' Подсчет ячеек, удовлетворяющих условию в каждом диапазоне
    For i = 1 To диапазон1.Cells.Count
        If диапазон1.Cells(i).Value > 0 Then count1 = count1 + 1
        If диапазон2.Cells(i).Value > 0 Then count2 = count2 + 1
        If диапазон3.Cells(i).Value > 0 Then count3 = count3 + 1
        If диапазон4.Cells(i).Value > 0 Then count4 = count4 + 1

        ' Подсчет непрерывных ячеек с самого начала
        If диапазон1.Cells(i).Value > 0 And диапазон2.Cells(i).Value > 0 And диапазон3.Cells(i).Value > 0 And диапазон4.Cells(i).Value > 0 And _
           диапазон1.Cells(i).Value <> "" And диапазон2.Cells(i).Value <> "" And диапазон3.Cells(i).Value <> "" And диапазон4.Cells(i).Value <> "" Then
            continuousCount = continuousCount + 1
        Else
            GetDeptsNumber = continuousCount
            Exit Function
            'continuousCount = 0
        End If
        'MsgBox count1 & ", " & count2 & ", " & count3 & ", " & count4
        'MsgBox continuousCount
    Next i

    ' Проверка равенства количества ячеек, удовлетворяющих условию
    If count1 = count2 And count1 = count3 And count1 = count4 Then
        GetDeptsNumber = continuousCount
    Else
        GetDeptsNumber = -1
    End If
'finish:
End Function
