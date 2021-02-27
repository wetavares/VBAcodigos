Public Sub OtimizaInicio()
'Sub rotina que otimiza o funcionamento para planilha ficar mais rapida
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
End Sub