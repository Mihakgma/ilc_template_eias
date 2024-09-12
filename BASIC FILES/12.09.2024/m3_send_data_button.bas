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
  Dim passWord As String
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
  
  
  
  
  ' ДОБАВИТЬ ЕЩЕ ТРИ АРРЭЯ ДЛЯ ИТЕРАЦИИ ПО НИМ (В СООТВЕТСВИИ С ОТДЕЛОМ) ВНУТРИ ТЕЛА ЦИКЛА FOR!!!
  ' fileNames, needRefreshDataColumnNumbers, probesAmountColumnNumbers
  ' Три переменных для текущего значения из соответствующего эррея
  ' currentFileName (String), currentNeedRefreshDataColumnNumber (Integer), currentProbesAmountColumnNumber (Integer)
  ' Счетчик успешных итераций по файлам (удалось открыть и обновить / отправить в нем данные)
  ' filesSuccessfullyProcessed (Integer)
  ' errorLog is a full report string of the errors occerred during macros processing
  ' it's planned to record TimeStamp + errorLog in a log txt-file in the same directory as main file has been set...
  ' it's necessary to create function doesFileExits(filePath String) -> Boolean
  
  
  arraysFROMTOLengthsEqual = True
  arraysDEPTSLengthsEqual = True
  errorLog = ""
  timeStart = Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss")
  delimeter = "-||-"
  
  ' MAIN INFO
  ITK = ThisWorkbook.Name
  dataSheetName = "info"
  deptsNumberDetected = Workbooks(ITK).Worksheets(dataSheetName).Range("J22").Value
  passWord = Workbooks(ITK).Worksheets(dataSheetName).Range("J11").Value
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
  
  MsgBox deptsNumberDetected & ", " & passWord & ", " & idOPPcolNumber & ", " & idDeptColNumber & ", " & codeOPPcolNumber & ", " & codeDeptColNumber & ", " & deptsMainSheetName
  MsgBox "RGB: (" & RGBredWarning & ", " & RGBgreenWarning & ", " & RGBblueWarning & ")"
  'MsgBox ColNumbersToCopy
  'MsgBox ColNumbersToPaste
  's = ""
  'For Each element In ColNumbersToCopy
  '      s = s & element & ", "
  '  Next element
  'MsgBox s
  
  For i = 0 To deptsNumberDetected - 1
    Dim resultStr As String
    ' Проходимся по всем множествам, относящимся к каждому отделу (...)
    ' Проверяем существует ли указанный файл, если не существует, то скипаем эту итерацию...
    ' Если файл существует (в противном случае), то вызываем функцию по отправке данных в отдел...
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
                                        passWord, _
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
        errorText = "Impossible to send data to departments cause of Col Numbers From (CopyFrom) Has <" _
                    & ArrayFromLen & "> elements. While one To (CopyTo) has <" & ArrayToLen _
                    & "> elements. Need to fix it on the sheet <" & dataSheetName & ">!"
        MsgBox errorText
        errorLog = errorLog + errorText + delimeter
    End If
    
    If arraysDEPTSLengthsEqual <> True Then
        errorText = "Impossible to send data to departments cause of Data Errors: " _
                    & "file names number provided = <" & fileNamesLen & ">. " _
                    & "departments names number provided = <" & deptsNamesLen & ">. " _
                    & "need refresh data columns number provided = <" & needRefreshDataLen & ">. " _
                    & "probes amount data columns number provided = <" & probesAmountLen & ">. "
                    
        MsgBox errorText
        errorLog = errorLog + errorText + delimeter
    End If
    
    
  errorLog = timeStart + delimeter + errorLog + s + delimeter + Format(DateTime.Now, "yyyy-MM-dd hh:mm:ss")
  SaveStringToFile txtFilePath, errorLog
    
End Sub
