Public Sub formataMoeda(valor)
'Esta sub sever para ser usada nos formularios do vba no excel
'Formatar campo de valor com ponto e virgula ao digitar
    If IsNumeric(valor) Then
        If InStr(1, valor, "-") >= 1 Then valor = Replace(valor, "-", "") 'retira sinal negativo caso digitado
        If InStr(1, valor, ",") >= 1 Then valor = CDbl(Replace(valor, ",", "")) 'retirar a virgula
        If InStr(1, valor, ".") >= 1 Then valor = Replace(valor, ".", "") 'para trabalhar melhor retiramos ponto
        'verifica casas para inserção de ponto
		Select Case Len(valor) 
            Case 1
            numPonto = "00" & valor
            Case 2
            numPonto = "0" & valor
            Case 6 To 8
            numPonto = Left(valor, Len(valor) - 5) & "." & Right(valor, 5)
            Case 9 To 11
            numPonto = inseriPonto(8, valor)
            Case 12 To 14
            numPonto = inseriPonto(11, valor)
            Case Else
            numPonto = valor
        End Select
        numVirgula = Left(numPonto, Len(numPonto) - 2) & "," & Right(numPonto, 2)
		'indica qual formulario e controle recebe a formatação
        UserForm_Entrada.TextBox_Valor.Value = numVirgula
    Else
        If valor = "" Then Exit Sub
        'Mensagem de numero ou caracter invalido
        MsgBox "Número invalido", vbCritical, "Caracter Invalido"
        Exit Sub
    End If
End Sub
