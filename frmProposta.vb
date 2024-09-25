Private Sub UserForm_Initialize()
    ' Configurando a ListBox para ter 5 colunas
    With Me.lstCliente
        .ColumnCount = 5
        ' Definindo as larguras das colunas: ID, Nome, Contato, Cidade, Estado
        .ColumnWidths = "48;166;152;108;24"
    End With
    
    ' Desabilitar o botão Selecionar Cliente por padrão
    Me.btnSelecionarCliente.Enabled = False
End Sub

Private Sub lstCliente_Click()
    ' Verifica se algum item está selecionado
    If Me.lstCliente.ListIndex <> -1 Then
        ' Preenche os campos com as informações do cliente selecionado
        Me.txtID.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 0) ' ID
        Me.txtNomeCliente.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 1) ' Nome
        Me.txtPessoaContato.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 2) ' Contato
        Me.txtCidade.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 3) ' Cidade
        Me.txtEstado.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 4) ' Estado
        
        ' Habilitar o botão Selecionar Cliente
        Me.btnSelecionarCliente.Enabled = True
    Else
        ' Desabilitar o botão Selecionar Cliente se nenhum item estiver selecionado
        Me.btnSelecionarCliente.Enabled = False
    End If
End Sub

Private Sub btnBuscaCliente_Click()
    Dim wsClientes As Worksheet
    Dim rngClientes As Range
    Dim cel As Range
    Dim idBusca As String
    Dim nomeBusca As String
    Dim encontrado As Boolean
    
    ' Definindo a planilha de clientes
    Set wsClientes = ThisWorkbook.Sheets("CLIENTES")
    ' Definindo o intervalo de dados dos clientes (ajuste conforme necessário)
    Set rngClientes = wsClientes.Range("A2", wsClientes.Cells(wsClientes.Rows.Count, "H").End(xlUp))
    
    ' Obtendo os valores de busca
    idBusca = Trim(Me.txtID.Value)
    nomeBusca = Trim(Me.txtNomeCliente.Value)
    
    ' Limpando a ListBox antes de adicionar novos itens
    Me.lstCliente.Clear
    
    ' Iterando sobre cada cliente
    For Each cel In rngClientes.Columns(1).Cells ' Coluna A para ID
        ' Verificando se o ID ou Nome contém o texto buscado
        If (idBusca <> "" And InStr(1, cel.Value, idBusca, vbTextCompare) > 0) Or _
           (nomeBusca <> "" And InStr(1, cel.Offset(0, 1).Value, nomeBusca, vbTextCompare) > 0) Then
           
            ' Adicionando o cliente à ListBox
            Me.lstCliente.AddItem cel.Value
            Me.lstCliente.List(Me.lstCliente.ListCount - 1, 1) = cel.Offset(0, 1).Value ' Nome
            Me.lstCliente.List(Me.lstCliente.ListCount - 1, 2) = cel.Offset(0, 2).Value ' Contato
            Me.lstCliente.List(Me.lstCliente.ListCount - 1, 3) = cel.Offset(0, 4).Value ' Cidade
            Me.lstCliente.List(Me.lstCliente.ListCount - 1, 4) = cel.Offset(0, 5).Value ' Estado
            
            encontrado = True
        End If
    Next cel
    
    ' Mensagem caso nenhum cliente seja encontrado
    If Not encontrado Then
        MsgBox "Nenhum cliente encontrado com os critérios fornecidos.", vbInformation
    End If
End Sub

Private Sub btnSelecionarCliente_Click()
    ' Desabilitar os campos para edição
    Me.txtID.Enabled = False
    Me.txtNomeCliente.Enabled = False
    Me.txtPessoaContato.Enabled = False
    Me.txtCidade.Enabled = False
    Me.txtEstado.Enabled = False
    
    ' Desabilitar a ListBox para impedir novas seleções
    Me.lstCliente.Enabled = False
End Sub

Private Sub btnLimparCliente_Click()
    ' Limpar os campos
    Me.txtID.Value = ""
    Me.txtNomeCliente.Value = ""
    Me.txtPessoaContato.Value = ""
    Me.txtCidade.Value = ""
    Me.txtEstado.Value = ""
    lstCliente.Clear
    
    ' Reabilitar os campos txtID e txtNomeCliente para edição
    Me.txtID.Enabled = True
    Me.txtNomeCliente.Enabled = True
    
    ' Desabilitar os outros campos
    Me.txtPessoaContato.Enabled = False
    Me.txtCidade.Enabled = False
    Me.txtEstado.Enabled = False
    
    ' Desabilitar o botão Selecionar Cliente
    Me.btnSelecionarCliente.Enabled = False
    
    ' Reabilitar a ListBox para permitir novas seleções
    Me.lstCliente.Enabled = True
End Sub


