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

  ' Diap for macros work
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
  Dim password As String
  Dim lockUnlockOk As Integer
  
  
  
  ITK = ThisWorkbook.Name
  dataSheetName = "info"
  checkCellRange = Workbooks(ITK).Worksheets(dataSheetName).Range("J7").Value
  rowChangedColNum = Workbooks(ITK).Worksheets(dataSheetName).Range("J8").Value
  deptsStartLetter = Workbooks(ITK).Worksheets(dataSheetName).Range("J9").Value
  deptsFinishLetter = Workbooks(ITK).Worksheets(dataSheetName).Range("J10").Value
  
  Set targetRange = Range(checkCellRange)
  
  Set rng = Application.Intersect(Target, targetRange)

  ' Checked cell is within predefined cell range?
  If Not rng Is Nothing Then
    password = Workbooks(ITK).Worksheets(dataSheetName).Range("J11").Value
    lockUnlockOk = UnlockSheet(ThisWorkbook.ActiveSheet, password)
    Set rng = Application.Intersect(rng.EntireRow, targetRange)
    For Each rw In rng.Rows 'loop over affected rows
        ' get row number of changed cell
        Dim rowNum As Long
        rowNum = rw.Row
        ' get cell range with probes number upon each dept
        deptsRowAdress = deptsStartLetter & rowNum & ":" & deptsFinishLetter & rowNum
        'MsgBox deptsRowAdress
        Set deptsRowRange = Range(deptsRowAdress)
        ' loop over each couln in range
        For i = 1 To deptsRowRange.Columns.Count
          ' check if column contains digit value bigger than 0
          If Application.WorksheetFunction.CountIf(deptsRowRange.Columns(i), ">0") Then
            Cells(rowNum, rowChangedColNum + i - 1).Value = 1
          End If
        Next i
    Next rw
    lockUnlockOk = LockSheet(ThisWorkbook.ActiveSheet, password)
  End If

End Sub
