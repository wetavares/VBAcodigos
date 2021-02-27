Public Function SoNumeros(l As IReturnInteger)
'Rotina para garantir que seja digitado apenas numeros no campo
    Select Case l
        Case Asc("0") To Asc("9")
            SoNumeros = l
        Case Else
            SoNumeros = 0
            MsgBox "Favor inserir apenas números!", vbExclamation, "CAMPO TIPO NÚMERO"
    End Select
End Function