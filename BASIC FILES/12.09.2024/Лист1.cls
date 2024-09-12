VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)

  ' Диапазон, для которого будет срабатывать макрос
  Dim targetRange As Range
  Dim ITK As String
  Dim dataSheetName As String
  Dim rowChangedColNum As Integer
  Dim deptsStartLetter As String
  Dim deptsFinishLetter As String
  Dim deptsRowAdress As String
  Dim deptsRowRange As Range
  Dim rng As Range
  Dim rw As Range
  
  
  
  ITK = ThisWorkbook.Name
  dataSheetName = "info"
  checkCellRange = Workbooks(ITK).Worksheets(dataSheetName).Range("J7").Value
  rowChangedColNum = Workbooks(ITK).Worksheets(dataSheetName).Range("J8").Value
  deptsStartLetter = Workbooks(ITK).Worksheets(dataSheetName).Range("J9").Value
  deptsFinishLetter = Workbooks(ITK).Worksheets(dataSheetName).Range("J10").Value
  
  Set targetRange = Range(checkCellRange)
  
  Set rng = Application.Intersect(Target, targetRange)

  ' Проверяем, является ли изменяемая ячейка в заданном диапазоне
  If Not rng Is Nothing Then
    Set rng = Application.Intersect(rng.EntireRow, targetRange)
    For Each rw In rng.Rows 'loop over affected rows
        ' получаем номер строки изменяемой ячейки
        Dim rowNum As Long
        'rowNum = Target.row
        rowNum = rw.Row
        ' Получаем диапазон ячеек для проверки распределения проб по отделам
        deptsRowAdress = deptsStartLetter & rowNum & ":" & deptsFinishLetter & rowNum
        'MsgBox deptsRowAdress
        Set deptsRowRange = Range(deptsRowAdress)
        ' Проходим по каждому столбцу в диапазоне
        For i = 1 To deptsRowRange.Columns.Count
          ' Проверяем, заполнен ли столбец цифрами
          If Application.WorksheetFunction.CountIf(deptsRowRange.Columns(i), ">0") Then
            Cells(rowNum, rowChangedColNum + i - 1).Value = 1
          ' This functionality is temporarly off cause its necesarry to keep in df previously marked depts
          ' Notwithstanding it has been marked wright or wrongly
          'Else
          '  Cells(rowNum, rowChangedColNum + i - 1).Value = ""
          End If
        Next i
        ' Вносим изменения в выбранную строку определенного столбца
        'Cells(rowNum, rowChangedColNum).Value = 1
        ' Пример: столбец 5  = E
    Next rw
  End If

End Sub
