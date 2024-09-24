Attribute VB_Name = "m1_get_probe_code"
Option Explicit

Function GetCode(ByVal rangeToCheck As Range, _
ByVal codes As Range, ByVal dateValue As Date, _
ByVal sep As String) As String

  ' Check if cell diapason is correct
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
  
  
  ' Separator is default output string value
  codeString = sep
  ' Empty string to data mark
  monthYear = ""
  ' Get constants values from work sheet
  ITK = ThisWorkbook.Name
  dataSheetName = "info"
  dateMinus = Workbooks(ITK).Worksheets(dataSheetName).Range("J3").Value
  datePlus = Workbooks(ITK).Worksheets(dataSheetName).Range("J4").Value
  maxProbesNumber = Workbooks(ITK).Worksheets(dataSheetName).Range("J5").Value
  errorResult = Workbooks(ITK).Worksheets(dataSheetName).Range("J6").Value

  ' Loop through each column in diap
  For i = 1 To rangeToCheck.Columns.Count
    ' Check if column values contain digits not and if their values are not equal to 0
    If Application.WorksheetFunction.CountIf(rangeToCheck.Columns(i), ">0") > 0 And rangeToCheck.Columns(i) < maxProbesNumber Then
      ' Append code according to column
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

    End If
  Next i

  On Error GoTo ErrorHandler
    If dateValue > (Now() - dateMinus) And dateValue < (Now() + datePlus) Then
      ' Date in string format
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
    ' string concat with date
    codeString = codeString & monthYear
    'Debug.Print codeString
ErrorHandler:
  ' return string with result code
  GetCode = codeString
  

End Function


Function ExtractMonthYear(ByVal dateValue As Date, ByVal sep As String) As String

  ' extract year from date
  Dim year As String
  year = Format(dateValue, "yy")

  ' extract month from date
  Dim month As String
  month = Format(dateValue, "mm")

  ' Output string in format YY-SEPARATOR-MM
  ExtractMonthYear = year & sep & month

End Function


Function CheckIfNeedUpdate(ByVal rangeToCheck As Range) As Integer

  ' Check is diap correct
  If rangeToCheck Is Nothing Then
    CheckIfNeedUpdate = 0
    Exit Function
  End If
  
  ' through each column inside diap
  For i = 1 To rangeToCheck.Columns.Count
    ' check if column contains digits
    If Application.WorksheetFunction.CountIf(rangeToCheck.Columns(i), ">0") Then
      CheckIfNeedUpdate = 1
      Exit Function
    End If
  Next i
  
  CheckIfNeedUpdate = 0

End Function
