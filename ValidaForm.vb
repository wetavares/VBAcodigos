Public Function ValidaForm(f As UserForm)
'Validar se os campos obrigatorios estão preenchidos
    Dim c As Control
    Dim i As Integer
    ValidaForm = False
    For i = 0 To f.Controls.Count - 1
        Set c = f.Controls(i)
'    MsgBox "controle - " & c.TabIndex & vbCrLf & "= " & c.Name & vbCrLf & "= " & c.Tag, vbOKOnly
        If c.Tag = "*" Then
            If TypeOf c Is TextBox Then
                If c.Text = "" Then
                     MsgBox "Preencha todos os campos"
                     c.SetFocus
                     Exit Function
                End If
            ElseIf TypeOf c Is ComboBox Then
    '           If c.Tag = "*" Then
                If c.Text = "" Then
                      MsgBox "Preencha todos os campos"
                      'c.SetFocus
                      Exit Function
                End If
    '           End If
            ElseIf TypeOf c Is ListBox Then
    '           If c.Tag = "*" Then
                If c.Text = "" Then
                     MsgBox "Preencha todos os campos"
                     c.SetFocus
                     Exit Function
                End If
    '           End If
            ElseIf TypeOf c Is Label Then
    '           If c.Tag = "*" Then
                If c.Text = "" Then
                     MsgBox "Preencha todos os campos"
                     c.SetFocus
                     Exit Function
                End If
    '           End If
            ElseIf TypeOf c Is OptionButton Then
    '           If c.Tag = "*" Then
                If c.Text = "" Then
                     MsgBox "Preencha todos os campos"
                     c.SetFocus
                     Exit Function
                End If
    '           End If
            Else
                ValidaForm = True
            End If
        End If
    Next
''    ValidaForm = True
End Function
