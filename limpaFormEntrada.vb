Public Sub limpaFormEntrada(frm As UserForm)
'limpar o formulario de entrada
    frm.Label_Codigo.Caption = Empty
    frm.TextBox_NrNFE.Value = Empty
    frm.TextBox_NfeEmissao.Value = Empty
    frm.TextBox_ValorTotal.Value = Format(0, "R$ #,###0.00")
    frm.TextBox_Unidade.Value = Empty
    frm.TextBox_Produto_Entrada.Value = Empty
    frm.TextBox_Quantidade.Value = Empty
    frm.TextBox_Valor.Value = Empty
    frm.ComboBox_Fornecedor.Value = Empty
    frm.TextBox_CodFornecedor.Value = Empty
    frm.OptionButton_Cento.Value = False
    frm.OptionButton_Duzia.Value = False
    frm.ListBox_ListaEnt.Clear
    
End Sub
