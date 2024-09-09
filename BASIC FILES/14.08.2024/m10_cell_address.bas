Attribute VB_Name = "m10_cell_address"
Public Function CellAddress(ByVal rowNumber As Long, _
                            ByVal colNumber As Long, _
                            ByVal removeDollar As Long) As String
    Dim address As String
    Dim errrorOut As String
    
    
    ' If removeDollar != 1 returns cell address WITH $ symbols (locked cell address)
    ' ElseIf removeDollar = 1 returns cell address WITHOUT $ symbols (UNlocked cell address)
    
    ' checks if numbers less than 1
    If rowNumber < 1 Then
        MsgBox "row number cannot be less than 1. it's current value is: <" & rowNumber & ">"
        CellAddress = errrorOut
        Exit Function
    End If
    If colNumber < 1 Then
        MsgBox "column number cannot be less than 1. it's current value is: <" & colNumber & ">"
        CellAddress = errrorOut
        Exit Function
    End If
    
    address = Cells(rowNumber, colNumber).address
    
    On Error GoTo ErrorHandler
    If removeDollar = 1 Then
        CellAddress = Replace(address, "$", "")
    Else
        CellAddress = address
    End If
    Exit Function
    
ErrorHandler:
    CellAddress = address
End Function
