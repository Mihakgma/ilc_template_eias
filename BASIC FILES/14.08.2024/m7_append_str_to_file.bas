Attribute VB_Name = "m7_append_str_to_file"
Public Function SaveStringToFile(ByVal filePath As String, ByVal textLine As String) As Integer

    Dim fullPath As String
    On Error GoTo ErrorHandler

    ' is absolute path?
    If InStr(filePath, ":\") > 0 Then
        ' does dir exist?
        If Dir(Left(filePath, InStr(filePath, "\") - 1), vbDirectory) = "" Then
            ' dir doesn't exist
            MsgBox "The next directory doesn't exist: " & Left(filePath, InStr(filePath, "\") - 1)
            SaveStringToFile = 0
            Exit Function
        End If
        fullPath = filePath ' using full path
    Else
        ' Not absolute path -> use current file dir
        fullPath = ThisWorkbook.Path & "\" & filePath
    End If

    ' Open txt file for appending new info from string
    Open fullPath For Append As #1
    Print #1, textLine
    Close #1

    SaveStringToFile = 1
    Exit Function

ErrorHandler:
    SaveStringToFile = 0
    MsgBox "An error occured while trying to save data in txt-file: " & Err.Description

End Function
