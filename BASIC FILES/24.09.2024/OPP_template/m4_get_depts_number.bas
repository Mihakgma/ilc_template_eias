Attribute VB_Name = "m4_get_depts_number"
Option Explicit

Function GetDeptsNumber(diap1 As Range, diap2 As Range, diap3 As Range, diap4 As Range) As Integer

    Dim i As Long
    Dim count1 As Long, count2 As Long, count3 As Long, count4 As Long
    Dim continuousCount As Long

    ' check not empty cells amount in range
    'MsgBox diap1.Cells.Count & ", " & diap2.Cells.Count & ", " & diap3.Cells.Count & ", " & diap4.Cells.Count
    If diap1.Cells.Count <> diap2.Cells.Count Or diap1.Cells.Count <> diap3.Cells.Count Or diap1.Cells.Count <> diap4.Cells.Count Then
        GetDeptsNumber = -1
        Exit Function
        'GoTo finish
    End If

    ' calculate cells amount which are ok according to predefined condition
    For i = 1 To diap1.Cells.Count
        If diap1.Cells(i).Value > 0 Then count1 = count1 + 1
        If diap2.Cells(i).Value > 0 Then count2 = count2 + 1
        If diap3.Cells(i).Value > 0 Then count3 = count3 + 1
        If diap4.Cells(i).Value > 0 Then count4 = count4 + 1

        ' calculate cells in a row (without blank cells inside range) from the beginning
        If diap1.Cells(i).Value > 0 And diap2.Cells(i).Value > 0 And diap3.Cells(i).Value > 0 And diap4.Cells(i).Value > 0 And _
           diap1.Cells(i).Value <> "" And diap2.Cells(i).Value <> "" And diap3.Cells(i).Value <> "" And diap4.Cells(i).Value <> "" Then
            continuousCount = continuousCount + 1
        Else
            GetDeptsNumber = continuousCount
            Exit Function
        End If
        'MsgBox count1 & ", " & count2 & ", " & count3 & ", " & count4
        'MsgBox continuousCount
    Next i

    ' the last cell ranges lengths condition check
    If count1 = count2 And count1 = count3 And count1 = count4 Then
        GetDeptsNumber = continuousCount
    Else
        GetDeptsNumber = -1
    End If
End Function
