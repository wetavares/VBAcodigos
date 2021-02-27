Public Sub CarregaListBox_PodValor(controle As Object, p As Worksheet)
'rotina que preenche o listbox produtos com valor
Dim rng As Range
Dim linha As Integer
Dim linhalistbox As Integer
linha = 4
controle.ColumnCount = 4
controle.ColumnHeads = True
controle.ColumnWidths = "40 pt;160 pt;25 pt;30 pt"
With p 'Sheets("Controle Financeiro")
    Set rng = p.Range(.Cells(5, "A"), .Cells(.Cells(.Rows.Count, "A").End(xlUp).Row, "H"))
End With
'Preenche os dados da planilha fornecedor no combobox
 Do Until p.Cells(linha, 1) = ""
    controle.AddItem 'p.Cells(linha, "A").Value
    controle.List(controle.ListCount - 1, 0) = p.Cells(linha, "A").Value
    controle.List(controle.ListCount - 1, 1) = p.Cells(linha, "B").Value
    controle.List(controle.ListCount - 1, 2) = p.Cells(linha, "C").Value
    controle.List(controle.ListCount - 1, 3) = p.Cells(linha, "H").Value
    linha = linha + 1
Loop
End Sub
