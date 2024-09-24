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
  ' Diap for macros processing
  Dim targetRange As Range
  Dim ITK As String
  Dim dataSheetName As String
  Dim rowChangedColNum As Integer
  Dim checkCellRange As String
  Dim passWord As String
  Dim rng As Range
  Dim rw As Range
  
  
  ITK = ThisWorkbook.Name
  'MsgBox WorkBookName
  dataSheetName = "info"
  checkCellRange = Workbooks(ITK).Worksheets(dataSheetName).Range("B5").Value
  rowChangedColNum = Workbooks(ITK).Worksheets(dataSheetName).Range("B6").Value
  passWord = Workbooks(ITK).Worksheets(dataSheetName).Range("B7").Value
  Set targetRange = Range(checkCellRange)
  Set rng = Application.Intersect(Target, targetRange)

  ' Checked cell is within predefined cell range?
  If Not rng Is Nothing Then
    UnlockSheet (passWord)
    Set rng = Application.Intersect(rng.EntireRow, targetRange)
    For Each rw In rng.Rows 'loop over affected rows
        ' get row number of changed cell
        Dim rowNum As Long
        rowNum = rw.Row
        ' set value in each edited row
        Cells(rowNum, rowChangedColNum).Value = 1
    Next rw
    ' lock active work sheet
    LockSheet (passWord)
  End If
End Sub
