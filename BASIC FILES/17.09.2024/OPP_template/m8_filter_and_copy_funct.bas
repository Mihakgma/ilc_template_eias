Attribute VB_Name = "m8_filter_and_copy_funct"
Option Explicit

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
    ByVal password As String, _
    ByVal redRGB As Integer, _
    ByVal greenRGB As Integer, _
    ByVal blueRGB As Integer) As String
    
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
    Dim i As Integer
    Dim rowsClearContents() As Long
    Dim filteredRowsNumber As Long
    Dim j As Integer
    Dim k As Integer
    Dim nullFiltered As Boolean
    
    
    
    OptimizeVBA (True)
    'function to set all the things you want to set, but hate keying in

    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    'this should stop those pesky enable prompts
    
    ' Variables default values
    canOpenWB = False
    targetWBIsOpened = False
    outResult = ""
    targetWSExists = False
    nullFiltered = False
    
    
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
    'RIGHT HERE NOW!!!
    On Error GoTo ErrorHandler
    ' TARGET WB CAN BE OPENED IN SAFE MODE BY ANY USER!!! SO THE LAST AUTHOR CHECK UP MIGHT BE FAIL!
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
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, filterColumn).End(xlUp).Row + 1
    'MsgBox "source WS max rows to copy: <" & lastRowSource & ">"
    ' skip header
    filterRow = 2
    j = 0
    'REPLACE lastColumnNumSource IF COLUMNS COUNT (TOTAL NUMBER) GROWS BIGGER!!!
    lastColumnNumSource = 56
    ' Filter data in source...
    wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRowSource, lastColumnNumSource)) _
    .AutoFilter Field:=filterColumn, Criteria1:="1", Operator:=xlFilterValues
    'filteredRowsNumber = wsSource.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Count
    'filteredRowsNumber = wsSource.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Areas(1).Rows.Count
    filteredRowsNumber = wsSource.AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    'MsgBox "Filtered rows to copy: <" & filteredRowsNumber & ">"
    If filteredRowsNumber < 1 Then
        nullFiltered = True
        GoTo ErrorHandler
    End If
    ReDim rowsClearContents(filteredRowsNumber)
    'MsgBox "Filter while processing: <" & wbTargetPath & ">"
    
    If UnlockSheet(wsTarget, password) = 1 Then
        Dim gg As Boolean
        'MsgBox "WS Target has been successfully unlocked"
    End If
    
    Do While filterRow <> lastRowSource
        codeFROM = wsSource.Cells(filterRow, codeOPPcolNumber).Value
        If wsSource.Cells(filterRow, filterColumn).Value <> "" And _
        codeFROM <> "" And codeFROM <> codeErrorText Then
            idValue = wsSource.Cells(filterRow, idColumnFROM).Value ' ѕолучаем значение id
            ' ѕровер€ем, существует ли строка с таким id в целевой книге
            lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, idColumnTO).End(xlUp).Row
            targetRow = 2
            'Stop
            'MsgBox "current id = <" & idValue & ">..." & ", lab code = <" & codeFROM & ">" & ", IDs in target WS: <" & lastRowTarget & ">"
            
            Do While targetRow <= lastRowTarget
                If wsTarget.Cells(targetRow, idColumnTO).Value = idValue Then
                    
                    'MsgBox "id = <" & idValue & "> has been found in target WS!"
                    ' row with equal ID is in source WS!!!
                    If targetRow <= wsSource.Rows.Count Then
                        For i = LBound(copyColumns) To UBound(copyColumns)
                            ' Copy values to target sheet cells
                            'MsgBox "trying to send data, row N = <" & targetRow & ">, array index N = <" & i & ">, from column N = <" _
                            '       & copyColumns(i) & ">, to column N = <" & pasteColumns(i) & ">, to cell address: <" & CellAddress(targetRow, pasteColumns(i), 1) _
                            '       & ">, from cell address: <" & CellAddress(targetRow, copyColumns(i), 1) & ">..."
                            
                            ' Debug.Print
                            ' differs lab codes then fill it with RGB color
                            If pasteColumns(i) = codeDeptColNumber And _
                            wsTarget.Range(CellAddress(targetRow, pasteColumns(i), 1)).Value <> codeFROM Then
                                wsTarget.Range(CellAddress(targetRow, pasteColumns(i), 1)) = wsSource.Range(CellAddress(filterRow, copyColumns(i), 1)).Value
                                wsTarget.Range(CellAddress(targetRow, pasteColumns(i), 1)).Font.Color = RGB(redRGB, greenRGB, blueRGB)
                                wsTarget.Range(CellAddress(targetRow, pasteColumns(i), 1)).Font.Bold = True
                            'ElseIf pasteColumns(i) <> codeDeptColNumber Then
                            Else
                                wsTarget.Range(CellAddress(targetRow, pasteColumns(i), 1)) = wsSource.Range(CellAddress(filterRow, copyColumns(i), 1)).Value
                            End If
                            'MsgBox "data successfully has been sent"
                        Next i
                    End If
                    Exit Do ' ID has been successfully found -> exit from while loop
                End If
            targetRow = targetRow + 1 ' next row iteration starting...
            Loop
        'MsgBox "data with pre-existed ID in target WS has been successfully sent"
        If targetRow > lastRowTarget Then
            ' HERE I PROCEED...
            ' ID value need to copy into target WS!!!
            wsTarget.Range(CellAddress(lastRowTarget + 1, idColumnTO, 1)) = wsSource.Range(CellAddress(filterRow, idColumnFROM, 1)).Value
            For i = LBound(copyColumns) To UBound(copyColumns)
                'wsTarget.Cells(lastRowTarget + 1, pasteColumns(i)).Value = wsSource.Cells(filterRow, copyColumns(i)).Value
                wsTarget.Range(CellAddress(lastRowTarget + 1, pasteColumns(i), 1)) = wsSource.Range(CellAddress(filterRow, copyColumns(i), 1)).Value
            Next i
        End If
        rowsClearContents(j) = filterRow
        j = j + 1
        End If
        filterRow = filterRow + 1
    Loop
    
    
ErrorHandler:

    ' ##### END OF SENDING DATA!!!
    

    If nullFiltered = True Then
        wbTarget.Close SaveChanges:=False
        Application.AutomationSecurity = msoAutomationSecurityLow
        OptimizeVBA (False)
        wsSource.AutoFilterMode = False
        FilterAndCopySpecificColumns = outResult
        Exit Function
    ElseIf canOpenWB <> True Then
        outResult = outResult & Err.Description & " with <" & wbTargetPath & ">"
        Application.AutomationSecurity = msoAutomationSecurityLow
        OptimizeVBA (False)
        'wsSource.AutoFilterMode = False
        FilterAndCopySpecificColumns = outResult
        Exit Function
    ElseIf targetWSExists <> True And targetWBIsOpened <> True Then
        wbTarget.Close SaveChanges:=False
        outResult = outResult & "Sheet <" & wsTargetName & "> doesn't exist in workbook <" & wbTargetPath & ">"
        Application.AutomationSecurity = msoAutomationSecurityLow
        OptimizeVBA (False)
        wsSource.AutoFilterMode = False
        FilterAndCopySpecificColumns = outResult
        Exit Function
    ' somebody opened target wb
    ElseIf targetWBIsOpened <> False Then
        Application.AutomationSecurity = msoAutomationSecurityLow
        OptimizeVBA (False)
        'wsSource.AutoFilterMode = False
        FilterAndCopySpecificColumns = outResult
        Exit Function
    ElseIf Err.Description <> "" Then
        MsgBox Err.Description
        outResult = outResult & Err.Description & " with <" & wbTargetPath & ">"
        Application.AutomationSecurity = msoAutomationSecurityLow
        OptimizeVBA (False)
        wsSource.AutoFilterMode = False
        FilterAndCopySpecificColumns = outResult
        Exit Function
    End If
    
    If LockSheet(wsTarget, password) = 0 Then
        outResult = outResult & " ws target lock problem, "
    End If
    
    FilterAndCopySpecificColumns = outResult
    Application.AutomationSecurity = msoAutomationSecurityLow
    OptimizeVBA (False)
    wbTarget.RefreshAll
    wbTarget.Close SaveChanges:=True
    
    ' Clear marks from successfully copied rows
    For k = 2 To lastRowSource
        On Error GoTo arrayError
        If k <> 1 Then
            wsSource.Range(CellAddress(k, filterColumn, 1)).ClearContents
        End If
arrayError:
    Next k
    
    wsSource.AutoFilterMode = False

End Function

