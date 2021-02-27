Public Sub CabecListaEnt(ByRef LstBox As Object)
'rotina para preencher o listbox lista de saida
Dim linha, coluna As Integer
linha = 2
coluna = 1
LstBox.ColumnCount = 9
LstBox.ColumnWidths = "40;50;50;30;40;160;30;50;180"
With LstBox
    LstBox.AddItem
    .List(0, 0) = "CODIGO"
    .List(0, 1) = "DATA MOV."
    .List(0, 2) = "NR NOTA"
    .List(0, 3) = "QTDE"
    .List(0, 4) = "VALOR"
    .List(0, 5) = "PRODUTO"
    .List(0, 6) = "UND"
    .List(0, 7) = "COD.FORN"
    .List(0, 8) = "FORNECEDOR"
End With
End Sub
