Attribute VB_Name = "m9_wb_current_owner"
Option Explicit

Public Function WorkbookOpenedBy(workbookName As String) As String

    Dim wb As Workbook
    Dim wbOpened As Boolean
    'Dim users As Variant
    'Dim row As Integer
    Dim User1 As String
    Dim Date1 As String
    Dim status1 As String
    Dim f
    Dim i
    Dim x
    Dim inUseBy
    Dim tempfile
    Dim filename As String
    
    
     tempfile = Environ("TEMP") + "\tempfile" + CStr(Int(Rnd * 1000))
    
    f = FreeFile
    i = InStrRev(workbookName, "\")
    If (i > 0) Then
        filename = Mid(workbookName, 1, i) + "~$" + Mid(workbookName, 1 + i)
    Else
        filename = "~$" + workbookName
    End If
    
    On Error GoTo ErrorHandler
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    fso.CopyFile filename, tempfile
    
    Open tempfile For Binary Access Read As #f
    Input #f, x
    Close (f)
    inUseBy = Mid(x, 2, Asc(x))
    fso.Deletefile tempfile
    Set fso = Nothing

    WorkbookOpenedBy = " нига: <" & workbookName & "> открыта пользователем: " & inUseBy
    Exit Function
    
ErrorHandler:
    WorkbookOpenedBy = ""
End Function
