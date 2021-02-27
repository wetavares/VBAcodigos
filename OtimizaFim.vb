Public Sub OtimizaFim()
'Sub rotina que finaliza a otimização do funcionamento da planilha ficar mais rapida
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
End Sub
