Public Sub CabecListaSaida(ByRef LstBox As Object)
'rotina para preencher o listbox lista de saida
Dim linha, coluna As Integer
linha = 2
coluna = 1
LstBox.ColumnCount = 8
LstBox.ColumnWidths = "50;140;50;180;40;40;50;50"
With LstBox
    LstBox.AddItem
    .List(0, 0) = "RELATORIO"
    .List(0, 1) = "IGREJA"
    .List(0, 2) = "CODIGO"
    .List(0, 3) = "PRODUTO"
    .List(0, 4) = "QTDE"
    .List(0, 5) = "UND"
    .List(0, 6) = "VALOR"
    .List(0, 7) = "DATA"
End With
End Sub
