Private Sub LimpaTelaCadastro()
'LIMPA O FORMULARIO DE CADASTRO
'Desbloqueia os campos do produto
    Me.TextBox_und.Enabled = True
    Me.TextBox_EstoqueMinimo.Enabled = True
    Me.TextBox_SaldoInicial.Enabled = True
    Me.TextBox_und.BackColor = &H80000005
    Me.TextBox_EstoqueMinimo.BackColor = &H80000005
    Me.TextBox_SaldoInicial.BackColor = &H80000005
    'TextBox_codigo.Value = ""
    TextBox_Produto.Value = ""
    TextBox_und.Value = ""
    ComboBox_Fornecedor.Value = ""
    TextBox_email.Value = ""
    TextBox_CnpjCpf.Value = ""
    TextBox_celular.Value = ""
    TextBox_EstoqueMinimo.Value = ""
    TextBox_SaldoInicial.Value = ""
    TextBox_EstoqueMinimo.Enabled = True
    TextBox_EstoqueMinimo.BackColor = &H80000005
    TextBox_SaldoInicial.Enabled = True
    TextBox_SaldoInicial.BackColor = &H80000005
' Ativa a auto numeração no Form
    cod = Range("A5000").End(xlUp).Offset(0, 0).Value
    ' Aqui está nosso contador atuando sempre e atualizando a cada gravação
    Me.TextBox_codigo = cod + 1
    With TextBox_Produto.Value = ""
        TextBox_und.SetFocus
    End With
End Sub
