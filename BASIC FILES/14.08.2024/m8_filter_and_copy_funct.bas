Attribute VB_Name = "m8_filter_and_copy_funct"
Public Function FilterAndCopySpecificColumns( _
    ByVal wbTargetPath As String, _
    ByVal wsTargetName As String, _
    ByVal filterColumn As Integer, _
    ByVal idColumnFROM As Integer, _
    ByVal idColumnTO As Integer, _
    ByVal copyColumns As Variant, _
    ByVal pasteColumns As Variant, _
    ByVal codeOPPcolNumber As Integer, _
    ByVal codeDeptColNumber As Integer, _
    ByVal codeErrorText As String, _
    ByVal passWord As String) As String
    
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
    Dim canOpenWB As Boolean
    Dim targetWBIsOpened As Boolean
    Dim targetWSExists As Boolean
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim outResult As String
    Dim idValue As Variant
    Dim author As String
    Dim lastRowSource As Long, lastRowTarget As Long
    Dim filterRow As Long, targetRow As Long
    Dim lastColumnNumSource As Integer
    Dim codeFROM As String
    
    
    
    OptimizeVBA (True)
    'function to set all the things you want to set, but hate keying in

    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    'this should stop those pesky enable prompts
    
    ' Variables default values
    canOpenWB = False
    targetWBIsOpened = False
    outResult = ""
    targetWSExists = False
    
    
    ' Function body
    ' Basic and aim workbooks
    Set wbSource = ThisWorkbook ' Current (basic) workbook
    ' Check if target wb is already opened
    author = WorkbookOpenedBy(wbTargetPath)
    'MsgBox author
    If author <> "" Then
        canOpenWB = True
        targetWBIsOpened = True
        outResult = author
        GoTo ErrorHandler
    End If
    On Error GoTo ErrorHandler
    Set wbTarget = Workbooks.Open(wbTargetPath)
    canOpenWB = True
    
    Set wsSource = wbSource.ActiveSheet
    Set wsTarget = wbTarget.Sheets(wsTargetName)
    ' Check if target sheet not exists
    If wsTarget Is Nothing Then
        wbTarget.Close
        GoTo ErrorHandler
    End If
    targetWSExists = True
    ' Close target workbook
    ' ##### FILTER AND SENDING DATA TO TARGET WB & WS
    
    ' last row after filter apply in source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, filterColumn).End(xlUp).row
    MsgBox "source WS after filter rows number is: <" & lastRowSource & ">"
    ' skip header
    filterRow = 2
    'ЗАМЕНИТЬ lastColumnNumSource ЕСЛИ УВЕЛИЧИТСЯ КОЛИЧЕСТВО СТОЛБЦОВ!!!
    lastColumnNumSource = 56
    ' Filter data in source...
    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRowSource, lastColumnNumSource)) _
    .AutoFilter Field:=filterColumn, Criteria1:="1", Operator:=xlFilterValues
    
    'MsgBox "Filter while processing: <" & wbTargetPath & ">"
    
    
    If UnlockSheet(wsTarget, passWord) = 1 Then
        MsgBox "WS Target has been successfully unlocked"
    End If
    
    Do While filterRow <> lastRowSource
        codeFROM = wsSource.Cells(filterRow, codeOPPcolNumber).Value
        If wsSource.Cells(filterRow, filterColumn).Value <> "" And _
        codeFROM <> "" And codeFROM <> codeErrorText Then
            idValue = wsSource.Cells(filterRow, idColumnFROM).Value ' Получаем значение id
            ' Проверяем, существует ли строка с таким id в целевой книге
            lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, idColumnTO).End(xlUp).row
            targetRow = 2
            
            MsgBox "current id = <" & idValue & ">..." & ", lab code = <" & codeFROM & ">" & ", IDs in target WS: <" & lastRowTarget & ">"
            
            Do While targetRow <= lastRowTarget
                If wsTarget.Cells(targetRow, idColumnTO).Value = idValue Then
                MsgBox "id = <" & idValue & "> has been found in target WS!"
                    ' row with equal ID is in source WS!!!
                    If targetRow <= wsSource.Rows.Count Then
                        For i = LBound(copyColumns) To UBound(copyColumns)
                            ' Copy values to target sheet cells
                            MsgBox "trying to send data, row N = <" & targetRow & ">, array index N = <" & i & ">, from column N = <" _
                                   & copyColumns(i) & ">, to column N = <" & pasteColumns(i) & ">, to cell address: <" & CellAddress(targetRow, pasteColumns(i), 1) _
                                   & ">, from cell address: <" & CellAddress(targetRow, copyColumns(i), 1) & ">..."
                            '& wsSource.Cells(targetRow, copyColumns(i)).Value
                            'wsTarget.Cells(targetRow, pasteColumns(i)) = wsSource.Cells(targetRow, copyColumns(i)).Value
                            'Range attr. will help (?)
                            wsTarget.Range(CellAddress(targetRow, pasteColumns(i), 1)) = wsSource.Range(CellAddress(targetRow, copyColumns(i), 1)).Value
                            'MsgBox "data successfully has been sent"
                        Next i
                    End If
                    Exit Do ' ID has been successfully found -> exit from while loop
                End If
            targetRow = targetRow + 1 ' next row iteration starting...
            Loop
        MsgBox "data with pre-existed ID in target WS has been successfully sent"
        ' ID hasn't found in target WS!!!
        If targetRow > lastRowTarget Then
            ' HERE I PROCEED...
            ' ID value need to copy into target WS!!!
            wsTarget.Range(CellAddress(lastRowTarget + 1, idColumnTO, 1)) = wsSource.Range(CellAddress(filterRow, idColumnFROM, 1)).Value
            For i = LBound(copyColumns) To UBound(copyColumns)
                'wsTarget.Cells(lastRowTarget + 1, pasteColumns(i)).Value = wsSource.Cells(filterRow, copyColumns(i)).Value
                wsTarget.Range(CellAddress(lastRowTarget + 1, pasteColumns(i), 1)) = wsSource.Range(CellAddress(filterRow, copyColumns(i), 1)).Value
            Next i
        End If
        
        End If
        
        filterRow = filterRow + 1
    Loop
    
    
ErrorHandler:

    ' ##### END OF SENDING DATA!!!
    ' Filter Off
    'wsSource.AutoFilterMode = False
    ' All data is refreshed!!!
    'wbTarget.RefreshAll
    'Application.AutomationSecurity = msoAutomationSecurityLow
    'make sure you set this up when done
    'OptimizeVBA (False)
    

    If canOpenWB <> True Then
        'outResult = "Cannot open <" & wbTargetPath
        outResult = outResult & Err.Description & " with <" & wbTargetPath & ">"
        Application.AutomationSecurity = msoAutomationSecurityLow
        OptimizeVBA (False)
        wsSource.AutoFilterMode = False
        Exit Function
    ElseIf targetWSExists <> True And targetWBIsOpened <> True Then
        wbTarget.Close SaveChanges:=False
        outResult = outResult & "Sheet <" & wsTargetName & "> doesn't exist in workbook <" & wbTargetPath & ">"
        Application.AutomationSecurity = msoAutomationSecurityLow
        OptimizeVBA (False)
        wsSource.AutoFilterMode = False
        Exit Function
    ElseIf Err.Description <> "" Then
        MsgBox Err.Description
        ' GoTo Doesn't work!!!
        ' On Error GoTo wbTargetClosed
        If WorkbookOpenedBy(wbTargetPath) <> "" Then
            ' WB Target has been already closed!
            wbTarget.Close SaveChanges:=False
        End If
' wbTargetClosed:
        outResult = Err.Description & " with <" & wbTargetPath & ">"
        Application.AutomationSecurity = msoAutomationSecurityLow
        OptimizeVBA (False)
        wsSource.AutoFilterMode = False
        Exit Function
    End If
    
    FilterAndCopySpecificColumns = outResult
    
    Application.AutomationSecurity = msoAutomationSecurityLow
    OptimizeVBA (False)
    wbTarget.RefreshAll
    wbTarget.Close SaveChanges:=True
    ' Need to think about filter in source WS!!!
    ' wsSource.AutoFilterMode = True

End Function

