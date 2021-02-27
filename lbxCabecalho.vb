Public Sub lbxCabecalho(c As Object, p As Worksheet)
'rotina para preencher o cabeçario do listbox de pesquisa
Dim linha, coluna As Integer
linha = 2
coluna = 1
c.ColumnCount = 10
c.ColumnWidths = "30;70;30;40;50;140;50;40;100;80"
With p
    Do Until IsEmpty(p.Cells(linha, coluna + 1))
        c.AddItem
        c.List(0, coluna - 1) = p.Cells(linha, coluna).Text
        coluna = coluna + 1
    Loop
End With
End Sub
