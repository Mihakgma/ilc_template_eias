Attribute VB_Name = "m4_get_depts_number"
Function GetDeptsNumber(��������1 As Range, ��������2 As Range, ��������3 As Range, ��������4 As Range) As Integer

    Dim i As Long
    Dim count1 As Long, count2 As Long, count3 As Long, count4 As Long
    Dim continuousCount As Long

    ' �������� ���������� ����� � ����������
    'MsgBox ��������1.Cells.Count & ", " & ��������2.Cells.Count & ", " & ��������3.Cells.Count & ", " & ��������4.Cells.Count
    If ��������1.Cells.Count <> ��������2.Cells.Count Or ��������1.Cells.Count <> ��������3.Cells.Count Or ��������1.Cells.Count <> ��������4.Cells.Count Then
        GetDeptsNumber = -1
        Exit Function
        'GoTo finish
    End If

    ' ������� �����, ��������������� ������� � ������ ���������
    For i = 1 To ��������1.Cells.Count
        If ��������1.Cells(i).Value > 0 Then count1 = count1 + 1
        If ��������2.Cells(i).Value > 0 Then count2 = count2 + 1
        If ��������3.Cells(i).Value > 0 Then count3 = count3 + 1
        If ��������4.Cells(i).Value > 0 Then count4 = count4 + 1

        ' ������� ����������� ����� � ������ ������
        If ��������1.Cells(i).Value > 0 And ��������2.Cells(i).Value > 0 And ��������3.Cells(i).Value > 0 And ��������4.Cells(i).Value > 0 And _
           ��������1.Cells(i).Value <> "" And ��������2.Cells(i).Value <> "" And ��������3.Cells(i).Value <> "" And ��������4.Cells(i).Value <> "" Then
            continuousCount = continuousCount + 1
        Else
            GetDeptsNumber = continuousCount
            Exit Function
            'continuousCount = 0
        End If
        'MsgBox count1 & ", " & count2 & ", " & count3 & ", " & count4
        'MsgBox continuousCount
    Next i

    ' �������� ��������� ���������� �����, ��������������� �������
    If count1 = count2 And count1 = count3 And count1 = count4 Then
        GetDeptsNumber = continuousCount
    Else
        GetDeptsNumber = -1
    End If
'finish:
End Function
