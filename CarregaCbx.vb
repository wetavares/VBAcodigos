Public Sub CarregaCbx(cmBox As Object, p As Worksheet)
'Rotina para carregar um combobox
'Carrega o Me.ComboBox_Fornecedor.ListIndex
Dim lin As Integer
Dim a(0, 2)
lin = 1
cmBox.ColumnCount = 3
cmBox.ColumnHeads = True
cmBox.ColumnWidths = "100 pt;100 pt;40 pt"
    
''    a(0, 0) = "Nome"
''    a(0, 1) = "CPF/CNPJ"
''    a(0, 2) = "CÓDIGO"
''  cmBox.List() = a
With p
    Set rng = p.Range(.Cells(1, "A"), .Cells(.Cells(.Rows.Count, "A").End(xlUp).Row, "E"))
End With
'Preenche os dados da planilha fornecedor no combobox
 Do Until p.Cells(lin, 1) = ""
    cmBox.AddItem
    cmBox.List(lin - 1, 0) = rng.Cells(lin, "C").Value    'p.Cells(lin, "C").Value
    cmBox.List(lin - 1, 1) = p.Cells(lin, "B").Value
    cmBox.List(lin - 1, 2) = p.Cells(lin, "A").Value
    lin = lin + 1
Loop

 
End Sub
