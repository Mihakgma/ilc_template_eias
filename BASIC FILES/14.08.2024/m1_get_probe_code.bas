Attribute VB_Name = "m1_get_probe_code"
Function GetCode(ByVal rangeToCheck As Range, _
ByVal codes As Range, ByVal dateValue As Date, _
ByVal sep As String) As String

  ' Проверяем, что диапазон и коды переданы корректно
  If rangeToCheck Is Nothing Then
    GetCode = ""
    Exit Function
  End If

  Dim cell As Range
  Dim codeString As String
  Dim i As Integer
  Dim year As String
  Dim month As String
  Dim monthYear As String
  Dim dateMinus As Integer
  Dim datePlus As Integer
  Dim ITK As String
  Dim dataSheetName As String
  Dim maxProbesNumber As Integer
  Dim errorResult As String
  
  
  ' Значение по умолчанию для строки кодов (сепаратором)
  codeString = sep
  ' Присваиваем дате пустую строку
  monthYear = ""
  ' Берем значения для проверки валидности дат поступления (чего?)
  ITK = ThisWorkbook.Name
  dataSheetName = "info"
  dateMinus = Workbooks(ITK).Worksheets(dataSheetName).Range("J3").Value
  datePlus = Workbooks(ITK).Worksheets(dataSheetName).Range("J4").Value
  maxProbesNumber = Workbooks(ITK).Worksheets(dataSheetName).Range("J5").Value
  errorResult = Workbooks(ITK).Worksheets(dataSheetName).Range("J6").Value

  ' Проходим по каждому столбцу в диапазоне
  For i = 1 To rangeToCheck.Columns.Count
    ' Проверяем, заполнен ли столбец цифрами и не равен ли 0
    If Application.WorksheetFunction.CountIf(rangeToCheck.Columns(i), ">0") > 0 And rangeToCheck.Columns(i) < maxProbesNumber Then
      ' Добавляем код, соответствующий столбцу
      codeString = codeString & codes.Columns(i) & sep
    ElseIf Application.WorksheetFunction.CountIf(rangeToCheck.Columns(i), ">0") > 0 And rangeToCheck.Columns(i) >= maxProbesNumber Then
      MsgBox "Probes number cannot be too big (" & rangeToCheck.Columns(i) & "), i.e. more than <" & maxProbesNumber & "> !!!"
      GetCode = errorResult
      Exit Function
      'GoTo ErrorHandler
    ElseIf rangeToCheck.Columns(i) < 0 Then
      MsgBox "Probes number (" & rangeToCheck.Columns(i) & ") cannot be negative, i.e. less than <0> !!!"
      GetCode = errorResult
      Exit Function
      'GoTo ErrorHandler
    'Else
      'GetCode = errorResult
      'Exit Function
      'GoTo ErrorHandler
    End If
  Next i

  On Error GoTo ErrorHandler
    If dateValue > (Now() - dateMinus) And dateValue < (Now() + datePlus) Then
      ' Получаем дату в виде строки
      ' ДОБАВИТЬ ПРОВЕРКУ НА ДАТУ!!!
      year = Format(dateValue, "yy")
      month = Format(dateValue, "mm")
      monthYear = year & sep & month
    ElseIf dateValue > 100 Then
      MsgBox "Inputted date (" & dateValue & ") is Wrong! It has to be between <" & (Now() - dateMinus) & "> AND <" & (Now() + datePlus) & ">"
      GetCode = errorResult
      Exit Function
    Else
      GetCode = errorResult
      Exit Function
    End If
    'Debug.Print monthYear
    ' Объединяем с датой
    codeString = codeString & monthYear
    'Debug.Print codeString
ErrorHandler:
  ' Возвращаем строку кодов
  GetCode = codeString
  

End Function


Function ExtractMonthYear(ByVal dateValue As Date, ByVal sep As String) As String

  ' Извлечение года из значения даты
  Dim year As String
  year = Format(dateValue, "yy")

  ' Извлечение месяца из значения даты
  Dim month As String
  month = Format(dateValue, "mm")

  ' Создание итоговой строки в формате ГГ-ММ
  ExtractMonthYear = year & sep & month

End Function


Function CheckIfNeedUpdate(ByVal rangeToCheck As Range) As Integer

  ' Проверяем, что диапазон передан корректно
  If rangeToCheck Is Nothing Then
    CheckIfNeedUpdate = 0
    Exit Function
  End If
  
  ' Проходим по каждому столбцу в диапазоне
  For i = 1 To rangeToCheck.Columns.Count
    ' Проверяем, заполнен ли столбец цифрами
    If Application.WorksheetFunction.CountIf(rangeToCheck.Columns(i), ">0") Then
      CheckIfNeedUpdate = 1
      Exit Function
    End If
  Next i
  
  CheckIfNeedUpdate = 0

End Function
