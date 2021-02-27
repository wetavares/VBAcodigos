Sub LimparForm(frm As UserForm)
'Rotina que limpa todos controles do formulario que tenham a propriedade TAG vazia 
       For Each c In frm.Controls
           If c.Tag = Empty Then
              c.Text = Empty
           End If
      Next
End Sub
