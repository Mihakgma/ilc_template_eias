Attribute VB_Name = "m12_lock_unlock_sheet"
Option Explicit

Function UnlockSheet(ByVal ws As Worksheet, _
                     ByVal UnlockSheetPassword As String) As Integer
    On Error GoTo ErrorHandler
    ws.Unprotect password:=UnlockSheetPassword
    UnlockSheet = 1
    Exit Function
ErrorHandler:
    UnlockSheet = 0
End Function


Function LockSheet(ByVal ws As Worksheet, _
                   ByVal UnlockSheetPassword As String) As Integer
    On Error GoTo ErrorHandler
    ws.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
       False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
       AllowFormattingRows:=True, AllowFiltering:=True, password:=UnlockSheetPassword
    LockSheet = 1
    Exit Function
ErrorHandler:
    LockSheet = 0
End Function

