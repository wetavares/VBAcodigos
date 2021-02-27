Sub GravaEntrada(ByRef LstBox As Object, p As Worksheet)
    'Grava os dados da Listbox1(5 colunas) da Plan1 na Plan2
    'Na Plan2 - Grava nas colunas A ate E a partir da Linha 2
    'antes limpando todos dados gravados anteriormente
    Dim nItem As Integer 'contador
    Dim Ctd1 As Integer
    Dim nlin_LBox As Integer 'Num linhas da ListBox
    Dim nlin_PL As Long 'Num linha para Limpar/gravar
    Dim V_Plan2 As Worksheet
    Dim V_LBox As OLEObject
    'ultima linha
    ulinha = p.Cells(Rows.Count, 1).End(xlUp).Row + 1
'    Set V_Plan2 = Worksheets("Entradas")
    nlin_LBox = LstBox.ListCount
    If nlin_LBox > 1 Then
        'Verifica ultima linha com dados na Plan2
        nlin_PL = p.Cells(Rows.Count, 1).End(xlUp).Row + 1
        'Loop grava listbox na planilha
        For nItem = 1 To nlin_LBox - 1 'ctd numero de linhas da listbox
            'For Ctd1 = 0 To 4  'ctd1 numero de colunas da listbox e Plan2
                p.Cells(nlin_PL, 1).Value = LstBox.List(nItem, 0) 'cod produto
                p.Cells(nlin_PL, 2).Value = Format(LstBox.List(nItem, 1), "dd/mm/yyyy") 'data
                p.Cells(nlin_PL, 3).Value = LstBox.List(nItem, 2) 'nota fiscal
                p.Cells(nlin_PL, 4).Value = LstBox.List(nItem, 3) 'qtde
                p.Cells(nlin_PL, 5).Value = CDbl(LstBox.List(nItem, 4)) 'valor
                p.Cells(nlin_PL, 7).Value = LstBox.List(nItem, 5) 'produto
                p.Cells(nlin_PL, 8).Value = LstBox.List(nItem, 6) 'unidade
                p.Cells(nlin_PL, 9).Value = LstBox.List(nItem, 7) 'cod.For
            'Next
            nlin_PL = nlin_PL + 1
        Next
        MsgBox "Conteudo da Lista gravados com sucesso"
    Else
        MsgBox "Lista vazia", vbOKOnly, "ATENÇÂO"
    End If
End Sub
