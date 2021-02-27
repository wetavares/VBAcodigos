Public Sub CarregaLbx(ByRef LstBox As Object, ByRef txtbox As Object, ByRef p As Worksheet)
'Carrega o Me.ListBox_Produto com 3 colunas: codigo, nome, unidade
Dim lin, col, linlistbox As Double
Dim rng As Range
Dim txtCelula, texto, pesquisa As String
lin = 4
Select Case LstBox.Name
    'Preenche o listbox no formulario que precisar do codigo das casas de oração
    Case "ListBox_Igreja"
    'novo para filtrar
        lin = 2
        col = 1
        linlistbox = 1
        pesquisa = txtbox.Text
        LstBox.ColumnCount = 4
        LstBox.ColumnWidths = "0;200;0;70"
        LstBox.Clear        
        With LstBox
            .AddItem
            .List(0, 0) = p.Cells(1, 1).Value
            .List(0, 1) = p.Cells(1, 2).Value
            .List(0, 2) = p.Cells(1, 3).Value
            .List(0, 3) = p.Cells(1, 4).Value
        End With        
        While p.Cells(lin, 1).Value <> Empty            
            texto = p.Cells(lin, 3).Text
                    If UCase(texto) Like UCase("*" & (pesquisa) & "*") Or pesquisa = Empty Then
                       txtCelula = p.Cells(lin, 3).Value
                            If p.Cells(lin, 1).Value <> Empty Then
                                With LstBox
                                    LstBox.AddItem
                                    .List(linlistbox, 0) = p.Cells(lin, 1)
                                    .List(linlistbox, 1) = p.Cells(lin, 2)
                                    .List(linlistbox, 2) = p.Cells(lin, 3)
                                    .List(linlistbox, 3) = p.Cells(lin, 4)
                                End With
                                linlistbox = linlistbox + 1
                            End If
                    End If
                    lin = lin + 1
        Wend
   ' Preenche o listbox no formulario que apresenta estoque
    Case "ListBox_Estoque"
    'novo com filtro no listbox estoque
        lin = 5
        col = 1
        linlistbox = 1
        pesquisa = txtbox.Text
        LstBox.ColumnCount = 8
        LstBox.ColumnWidths = "40;160;25;0;0;40;30;0"
        LstBox.Clear       
        With LstBox
            .AddItem
            .List(0, 0) = p.Cells(4, 1).Value
            .List(0, 1) = p.Cells(4, 2).Value
            .List(0, 2) = p.Cells(4, 3).Value
            .List(0, 5) = p.Cells(4, 6).Value
            .List(0, 6) = p.Cells(4, 7).Value
        End With
        While p.Cells(lin, 1).Value <> Empty
            texto = p.Cells(lin, 2).Text
                    If UCase(texto) Like UCase("*" & (pesquisa) & "*") Or pesquisa = Empty Then
                        If p.Cells(lin, 1).Value <> Empty Then
                            With LstBox
                                LstBox.AddItem
                                .List(linlistbox, 0) = p.Cells(lin, 1)
                                .List(linlistbox, 1) = p.Cells(lin, 2)
                                .List(linlistbox, 2) = p.Cells(lin, 3)
                                .List(linlistbox, 5) = p.Cells(lin, 6)
                                .List(linlistbox, 6) = p.Cells(lin, 7)
                            End With
                            linlistbox = linlistbox + 1
                        End If
                    End If
                    lin = lin + 1
        Wend
    'carrega o listbox de produtos no formulario entrada
    Case "ListBox_Produto_Entradas"
    'novo com filtro no listbox estoque
        lin = 5
        col = 1
        linlistbox = 1
        pesquisa = txtbox.Text
        LstBox.ColumnCount = 8
        LstBox.ColumnWidths = "40;160;25;0;0;40;30;0"
        LstBox.Clear
        
        With LstBox
            .AddItem
            .List(0, 0) = p.Cells(4, 1).Value
            .List(0, 1) = p.Cells(4, 2).Value
            .List(0, 2) = p.Cells(4, 3).Value
            .List(0, 5) = p.Cells(4, 6).Value
            .List(0, 6) = p.Cells(4, 7).Value
        End With
        While p.Cells(lin, 1).Value <> Empty
            texto = p.Cells(lin, 2).Text
                    If UCase(texto) Like UCase("*" & (pesquisa) & "*") Or pesquisa = Empty Then
                        If p.Cells(lin, 1).Value <> Empty Then
                            With LstBox
                                LstBox.AddItem
                                .List(linlistbox, 0) = p.Cells(lin, 1)
                                .List(linlistbox, 1) = p.Cells(lin, 2)
                                .List(linlistbox, 2) = p.Cells(lin, 3)
                                .List(linlistbox, 5) = p.Cells(lin, 6)
                                .List(linlistbox, 6) = p.Cells(lin, 7)
                            End With
                            linlistbox = linlistbox + 1
                        End If
                    End If
                    lin = lin + 1
        Wend         
    'carregar o listbos da pesquisa de saida
    Case "ListBoxPesqSaida"
        With p
            Set rng = p.Range(.Cells(2, "A"), .Cells(.Cells(.Rows.Count, "A").End(xlUp).Row, "K"))
        End With
        With LstBox
            LstBox.ColumnCount = 10
            LstBox.ColumnWidths = "30;70;30;40;50;140;50;40;100;80"
            LstBox.RowSource = rng.Address
        End With
    'Preenche o listbox dos produtos de entrada
    Case "ListBox_Produto_Entradas"
        With p
            Set rng = p.Range(.Cells(5, "A"), .Cells(.Cells(.Rows.Count, "A").End(xlUp).Row, "H"))
        End With        
        With LstBox
            LstBox.ColumnCount = 3
            LstBox.ColumnHeads = True
            LstBox.ColumnWidths = "28,5 pt;110 pt;15 pt"
            LstBox.RowSource = rng.Address
            End With
    End Select
End Sub
