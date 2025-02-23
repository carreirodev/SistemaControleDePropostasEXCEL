 Código para o formulário frmCliente

Private Sub btnBuscar_Click()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim searchText As String
    
    ' Limpa a lista de resultados
    lstResultados.Clear
    
    ' Define a planilha "BancoDeCliente"
    Set ws = ThisWorkbook.Worksheets("BancoDeCliente")
    
    ' Encontra a última linha com dados
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Obtém o texto de busca
    searchText = LCase(txtBuscaCliente.Value)
    
    ' Percorre todos os clientes e adiciona os que correspondem à busca
    For i = 2 To lastRow ' Assumindo que a primeira linha é cabeçalho
        If InStr(1, LCase(ws.Cells(i, 1).Value), searchText) > 0 Then
            lstResultados.AddItem
            lstResultados.List(lstResultados.ListCount - 1, 0) = ws.Cells(i, 1).Value ' Nome
            lstResultados.List(lstResultados.ListCount - 1, 1) = ws.Cells(i, 2).Value ' Cidade
            lstResultados.List(lstResultados.ListCount - 1, 2) = ws.Cells(i, 3).Value ' Estado
        End If
    Next i
End Sub

Private Sub btnSelecionar_Click()
    Dim selectedIndex As Long
    
    ' Verifica se um cliente foi selecionado
    If lstResultados.ListIndex = -1 Then
        MsgBox "Por favor, selecione um cliente da lista.", vbExclamation
        Exit Sub
    End If
    
    ' Obtém o índice do cliente selecionado
    selectedIndex = lstResultados.ListIndex
    
    ' Chama a sub do frmProposta para preencher os dados do cliente
    frmProposta.PreencherDadosCliente lstResultados.List(selectedIndex, 0), _
                                      lstResultados.List(selectedIndex, 1), _
                                      lstResultados.List(selectedIndex, 2)
    
    ' Fecha o formulário frmCliente
    Unload Me
End Sub

Private Sub btnLimpar_Click()
    ' Limpa todos os campos de texto
    txtBuscaCliente.Value = ""
    txtCidade.Value = ""
    txtEstado.Value = ""
    
    ' Limpa a lista de resultados
    lstResultados.Clear
End Sub

Private Sub btnFechar_Click()
    ' Fecha o formulário frmCliente sem nenhuma ação
    Unload Me
End Sub

Private Sub lstResultados_Click()
    ' Verifica se um item foi selecionado
    If lstResultados.ListIndex > -1 Then
        ' Preenche os campos com as informações do cliente selecionado
        txtBuscaCliente.Value = lstResultados.List(lstResultados.ListIndex, 0)
        txtCidade.Value = lstResultados.List(lstResultados.ListIndex, 1)
        txtEstado.Value = lstResultados.List(lstResultados.ListIndex, 2)
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Centraliza o formulário na tela
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    ' Configura os cabeçalhos do ListBox
    With lstResultados
        .AddItem
        .List(0, 0) = "Nome"
        .List(0, 1) = "Cidade"
        .List(0, 2) = "Estado"
    End With
End Sub