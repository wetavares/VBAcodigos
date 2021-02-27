Public Sub cbxCriterio(c As Object, p As Worksheet)
'rotina para preencher os combobox de criterios
Dim linha, coluna As Integer
linha = 2
coluna = 1
c.ColumnCount = 1
c.ListRows = 11
c.ColumnWidths = 100
With p
    Do Until IsEmpty(p.Cells(linha, coluna))
        c.AddItem p.Cells(linha, coluna).Text
        coluna = coluna + 1
    Loop
End With
End Sub
