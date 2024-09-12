Attribute VB_Name = "m6_file_exists"
Option Explicit

Public Function CheckFileExists(ByVal filePath As String) As Boolean

    CheckFileExists = Dir(filePath, vbNormal) <> ""

End Function
