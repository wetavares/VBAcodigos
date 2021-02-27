Public Sub RemoverItemLbx(frm As UserForm)
'Excluir um item selecionado do listbox
Dim i As Long
'Verifica se a lista está vazia
If frm.ListBox_Listagem.ListCount >= 1 Then
 'Caso não tenha item selecionado sai
    If frm.ListBox_Listagem.ListIndex = -1 Then
        frm.CommandButton_excluirItem.Locked = False
        frm.CommandButton_excluirItem.BackColor = &H80000000
        Exit Sub
    End If
    For i = frm.ListBox_Listagem.ListCount - 1 To 0 Step -1
        If frm.ListBox_Listagem.Selected(0) Then
        frm.CommandButton_excluirItem.Locked = False
        frm.CommandButton_excluirItem.BackColor = &H80000000
        Exit For
        ElseIf frm.ListBox_Listagem.Selected(i) Then
            If MsgBox("Deseja excluir este item ?", vbExclamation + vbYesNo) = vbYes Then
                frm.ListBox_Listagem.RemoveItem (frm.ListBox_Listagem.ListIndex)
                MsgBox "Item deletado com sucesso", vbInformation
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    Next i
'    frm.ListBox_Listagem.RemoveItem (frm.ListBox_Listagem.ListIndex)
End If
End Sub
