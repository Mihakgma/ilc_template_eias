Attribute VB_Name = "Module1"
Function UnlockSheet(ByVal UnlockSheetPassword As String) As Integer
    ' ...
    'wb.Worksheets(WorkSheetName).Activate
    ActiveSheet.Unprotect passWord:=UnlockSheetPassword
    UnlockSheet = 1
End Function

Function LockSheet(ByVal UnlockSheetPassword As String) As Integer
    ' ...
    'wb.Worksheets(WorkSheetName).Activate
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
                False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                AllowFormattingRows:=True, AllowFiltering:=True, passWord:=UnlockSheetPassword
    LockSheet = 1
End Function
