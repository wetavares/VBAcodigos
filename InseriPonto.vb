Function InseriPonto(inicio, valor)
'Função que insere o ponto no valor digitado
    i = Left(valor, Len(valor) - inicio)
    M1 = Left(Right(valor, inicio), 3)
    M2 = Left(Right(valor, 8), 3)
    f = Right(valor, 5)
    If (M2 = M1) And (Len(valor) < 12) Then
        inseriPonto = i & "." & M1 & "." & f
    Else
        inseriPonto = i & "." & M1 & "." & M2 & "." & f
    End If
End Function
