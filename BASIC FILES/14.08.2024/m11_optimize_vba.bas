Attribute VB_Name = "m11_optimize_vba"
Function OptimizeVBA(ByRef Status As Boolean)

    If Status = True Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        'DoCmd.SetWarnings = False
    Else
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        'DoCmd.SetWarnings = True
    End If

    End Function
