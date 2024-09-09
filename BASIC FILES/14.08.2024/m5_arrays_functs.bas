Attribute VB_Name = "m5_arrays_functs"
Public Function RangeToArray(inputRange As Range) As Variant()
Dim size As Integer
Dim inputValue As Variant, outputArray() As Variant

    ' inputValue will either be an variant array for ranges with more than 1 cell
    ' or a single variant value for range will only 1 cell
    inputValue = inputRange

    On Error Resume Next
    size = UBound(inputValue)

    If Err.Number = 0 Then
        RangeToArray = inputValue
    Else
        On Error GoTo 0
        ReDim outputArray(1 To 1, 1 To 1)
        outputArray(1, 1) = inputValue
        RangeToArray = outputArray
    End If

    On Error GoTo 0

End Function

Public Function FilterArray(inputArray As Variant) As Variant()
Dim outputArray() As Variant
Dim i As Integer, j As Integer
    
    ReDim outputArray(UBound(inputArray))
    j = 0

    For Each elem In inputArray
        If elem <> "" Then
            outputArray(j) = elem
            j = j + 1
        End If
    Next
    
    ReDim Preserve outputArray(j - 1)
    
    FilterArray = outputArray
    
End Function


Public Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

