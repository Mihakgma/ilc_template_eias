Attribute VB_Name = "m6_file_exists"
Option Explicit

Public Function CheckFileExists(ByVal filePath As String) As Boolean
On Error GoTo ErrorHandler
    CheckFileExists = Dir(filePath, vbNormal) <> ""
    Exit Function
ErrorHandler:
    CheckFileExists = False
End Function
