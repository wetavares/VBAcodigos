Public Sub LimpaForm(Formulario As UserForm)
'Limpar todos controles do formulario
Dim Control As Object

For Each Control In Formulario.Controls
    If TypeOf Control Is MSForms.TextBox Or TypeOf Control Is MSForms.ComboBox Then
        Control.Value = Empty
        ElseIf TypeOf Control Is MSForms.ListBox Then
            Control.Clear
            ElseIf TypeOf Control Is MSForms.OptionButton Then
                Control.Value = False
    End If
Next
End Sub