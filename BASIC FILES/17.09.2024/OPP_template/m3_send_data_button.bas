Attribute VB_Name = "m3_send_data_button"
Option Explicit

Sub Кнопка1_Щелчок()

  Dim deptsNumberDetected As Integer
  Dim idOPPcolNumber As Integer
  Dim idDeptColNumber As Integer
  Dim codeOPPcolNumber As Integer
  Dim codeDeptColNumber As Integer
  
  Dim RGBredWarning As Integer
  Dim RGBgreenWarning As Integer
  Dim RGBblueWarning As Integer
  
  Dim ITK As String
  Dim dataSheetName As String
  Dim password As String
  Dim deptsMainSheetName As String
  Dim i As Integer
  Dim ColNumbersToCopy As Variant
  Dim ColNumbersToPaste As Variant
  Dim s As String
  Dim arraysFROMTOLengthsEqual As Boolean
  Dim arraysDEPTSLengthsEqual As Boolean
  Dim ArrayFromLen As Integer
  Dim ArrayToLen As Integer
  Dim fileNames As Variant
  Dim needRefreshDataColumnNumbers As Variant
  Dim probesAmountColumnNumbers As Variant
  Dim deptsNames As Variant
  Dim fileNamesLen As Integer
  Dim needRefreshDataLen As Integer
  Dim probesAmountLen As Integer
  Dim deptsNamesLen As Integer
  
  Dim currentFileName As String
  Dim currentNeedRefreshDataColumnNumber As Integer
  Dim currentProbesAmountColumnNumber As Integer
  Dim currentDeptName As String
  
  Dim filesSuccessfullyProcessed As Integer
  Dim errorText As String
  Dim errorLog As String
  Dim txtFilePath As String
  Dim delimeter As String
  Dim timeStart As String
  Dim codeErrorText As String
  Dim lockUnlockOk As Integer
  
  
  
  arraysFROMTOLengthsEqual = True
  arraysDEPTSLengthsEqual = True
  errorLog = ""
  timeStart = Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss")
  delimeter = "-||-"
  
  
  ' MAIN INFO
  ITK = ThisWorkbook.Name
  dataSheetName = "info"
  deptsNumberDetected = Workbooks(ITK).Worksheets(dataSheetName).Range("J22").Value
  password = Workbooks(ITK).Worksheets(dataSheetName).Range("J11").Value
  deptsMainSheetName = Workbooks(ITK).Worksheets(dataSheetName).Range("J20").Value
  txtFilePath = Workbooks(ITK).Worksheets(dataSheetName).Range("J23").Value
  codeErrorText = Workbooks(ITK).Worksheets(dataSheetName).Range("J6").Value
  ' column numbers
  idOPPcolNumber = Workbooks(ITK).Worksheets(dataSheetName).Range("K15").Value
  idDeptColNumber = Workbooks(ITK).Worksheets(dataSheetName).Range("K16").Value
  codeOPPcolNumber = Workbooks(ITK).Worksheets(dataSheetName).Range("K17").Value
  codeDeptColNumber = Workbooks(ITK).Worksheets(dataSheetName).Range("K18").Value
  ' Deprecated function RangeToArray() plz, don't use it!!!
  'ColNumbersToCopy = RangeToArray(Workbooks(ITK).Worksheets(dataSheetName).Range("K12:U12"))
  ColNumbersToCopy = Application.Transpose(Workbooks(ITK).Worksheets(dataSheetName).Range("K12:U12"))
  ColNumbersToCopy = FilterArray(ColNumbersToCopy)
  ColNumbersToPaste = Application.Transpose(Workbooks(ITK).Worksheets(dataSheetName).Range("K13:U13"))
  ColNumbersToPaste = FilterArray(ColNumbersToPaste)
  ' Check whither arrays got lengths are equal?
  ArrayFromLen = ArrayLen(ColNumbersToCopy)
  ArrayToLen = ArrayLen(ColNumbersToPaste)
  If ArrayFromLen <> ArrayToLen Then
    arraysFROMTOLengthsEqual = False
    GoTo ErrorHandler
  End If
  ' RGB colors for warning in depts files if code is not equal for row
  RGBredWarning = Workbooks(ITK).Worksheets(dataSheetName).Range("J19").Value
  RGBgreenWarning = Workbooks(ITK).Worksheets(dataSheetName).Range("K19").Value
  RGBblueWarning = Workbooks(ITK).Worksheets(dataSheetName).Range("L19").Value
  ' Other arrays with info about departments
  fileNames = Application.Transpose(Workbooks(ITK).Worksheets(dataSheetName).Range("K4:U4"))
  fileNames = FilterArray(fileNames)
  needRefreshDataColumnNumbers = Application.Transpose(Workbooks(ITK).Worksheets(dataSheetName).Range("K5:U5"))
  needRefreshDataColumnNumbers = FilterArray(needRefreshDataColumnNumbers)
  probesAmountColumnNumbers = Application.Transpose(Workbooks(ITK).Worksheets(dataSheetName).Range("K14:U14"))
  probesAmountColumnNumbers = FilterArray(probesAmountColumnNumbers)
  deptsNames = Application.Transpose(Workbooks(ITK).Worksheets(dataSheetName).Range("K2:U2"))
  deptsNames = FilterArray(deptsNames)
  ' Check whither arrays got lengths are equal?
  fileNamesLen = ArrayLen(fileNames)
  needRefreshDataLen = ArrayLen(needRefreshDataColumnNumbers)
  probesAmountLen = ArrayLen(probesAmountColumnNumbers)
  deptsNamesLen = ArrayLen(deptsNames)
  If fileNamesLen <> needRefreshDataLen Or _
     fileNamesLen <> probesAmountLen Or _
     fileNamesLen <> deptsNamesLen Or _
     needRefreshDataLen <> probesAmountLen Or _
     needRefreshDataLen <> deptsNamesLen Or _
     probesAmountLen <> deptsNamesLen Then
    arraysDEPTSLengthsEqual = False
    GoTo ErrorHandler
  End If
  
  ' unlocking current WS
  lockUnlockOk = UnlockSheet(ThisWorkbook.ActiveSheet, password)
  
  MsgBox deptsNumberDetected & ", " & password & ", " & idOPPcolNumber & ", " & idDeptColNumber & ", " & codeOPPcolNumber & ", " & codeDeptColNumber & ", " & deptsMainSheetName
  'MsgBox "RGB: (" & RGBredWarning & ", " & RGBgreenWarning & ", " & RGBblueWarning & ")"
  
  For i = 0 To deptsNumberDetected - 1
    Dim resultStr As String
    ' loop through arrays inside dept
    ' check if file exists -> send data to dept file else skip current iteration
    currentFileName = fileNames(i)
    currentNeedRefreshDataColumnNumber = needRefreshDataColumnNumbers(i)
    currentProbesAmountColumnNumber = probesAmountColumnNumbers(i)
    currentDeptName = deptsNames(i)
    ' Check for file exists or not? If not GoTo IterErrorHandler (UPDATE FUNCTION DOESN'T WORK ON CURRENT ITERATION)!
    ' USE IF-ELSE STATEMENT INSTEAD OF GOTO KEY WORD!!!
    If CheckFileExists(currentFileName) Then
        ' UPDATE (REFRESH) FUNCTION WORKS HERE!
        resultStr = FilterAndCopySpecificColumns(currentFileName, _
                                        deptsMainSheetName, _
                                        currentNeedRefreshDataColumnNumber, _
                                        idOPPcolNumber, _
                                        idDeptColNumber, _
                                        ColNumbersToCopy, _
                                        ColNumbersToPaste, _
                                        codeOPPcolNumber, _
                                        codeDeptColNumber, _
                                        codeErrorText, _
                                        password, _
                                        RGBredWarning, _
                                        RGBgreenWarning, _
                                        RGBblueWarning)
        
        MsgBox "Department name: <" & currentDeptName & ">, iteration number = <" & i & ">, filename = <" _
               & currentFileName & ">, need refresh data columns number = <" _
               & currentNeedRefreshDataColumnNumber & ">, probes amount for current department columns number = <" _
               & currentProbesAmountColumnNumber & ">"
        
        If resultStr <> "" Then
            errorLog = errorLog + resultStr + delimeter
        Else
            filesSuccessfullyProcessed = filesSuccessfullyProcessed + 1
        End If
        
    Else
        errorText = "File <" & currentFileName & "> doesn't exist."
                    
        MsgBox errorText
        errorLog = errorLog + errorText + delimeter
    
    End If
  Next i
  
  s = "<" & filesSuccessfullyProcessed & "> of overall <" & deptsNumberDetected & "> files have been successfully processed." & delimeter
  
  MsgBox s
  
  
    
ErrorHandler:
    If arraysFROMTOLengthsEqual <> True Then
        errorText = "Arrays Len Error: Col Numbers From (CopyFrom) Has <" _
                    & ArrayFromLen & "> elements. While one To (CopyTo) has <" & ArrayToLen _
                    & "> elements. Need to fix it on the sheet <" & dataSheetName & ">!"
        MsgBox errorText
        errorLog = errorLog + errorText + delimeter
    End If
    
    If arraysDEPTSLengthsEqual <> True Then
        errorText = "Arrays Len Error: " _
                    & "file names number provided = <" & fileNamesLen & ">. " _
                    & "departments names number provided = <" & deptsNamesLen & ">. " _
                    & "need refresh data columns number provided = <" & needRefreshDataLen & ">. " _
                    & "probes amount data columns number provided = <" & probesAmountLen & ">. "
                    
        MsgBox errorText
        errorLog = errorLog + errorText + delimeter
    End If
    
    
  errorLog = timeStart + delimeter + errorLog + s + delimeter + Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss")
  SaveStringToFile txtFilePath, errorLog
  
  ' SET BACK COLUMN FILTER IN WS SOURCE
  ThisWorkbook.ActiveSheet.Range("A1").AutoFilter
  
  ' unlocking current WS
  lockUnlockOk = LockSheet(ThisWorkbook.ActiveSheet, password)
    
End Sub
