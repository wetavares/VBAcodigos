Public Sub LimpaForm(Formulario As UserForm)
'Limpar todos controles do formulario
Dim Control As Object
'Percorre todo o formulario verificando buscando os controles
For Each Control In Formulario.Controls
    'Verifica qual tipo de controle e limpa
    If TypeOf Control Is MSForms.TextBox Or TypeOf Control Is MSForms.ComboBox Then
        Control.Value = Empty
        ElseIf TypeOf Control Is MSForms.ListBox Then
            Control.Clear
            ElseIf TypeOf Control Is MSForms.OptionButton Then
                Control.Value = False
                ElseIf TypeOf Control Is MSForms.Label And Control.Tag <> "#" Then
                    Control.Caption = Empty                   
    End If
    
End Sub