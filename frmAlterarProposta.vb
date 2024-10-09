Private Sub UserForm_Initialize()
    ' Configurar a ListBox lstClientesListados
    With Me.lstClientesListados
        .ColumnCount = 3
        .ColumnWidths = "45;120;120"
    End With
    
    ' Desabilitar o botão Selecionar Cliente por padrão (se existir)
    Me.btnSelecionarCliente.Enabled = False
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
    Set rngClientes = wsClientes.Range("A2", wsClientes.Cells(wsClientes.Rows.Count, "C").End(xlUp))
    
    ' Obtendo os valores de busca
    idBusca = Trim(Me.txtID.Value)
    nomeBusca = Trim(Me.txtNomeCliente.Value)
    
    ' Limpando a ListBox antes de adicionar novos itens
    Me.lstClientesListados.Clear
    
    ' Iterando sobre cada cliente
    For Each cel In rngClientes.Columns(1).Cells ' Coluna A para ID
        ' Verificando se o ID ou Nome contém o texto buscado
        If (idBusca <> "" And InStr(1, cel.Value, idBusca, vbTextCompare) > 0) Or _
           (nomeBusca <> "" And InStr(1, cel.Offset(0, 1).Value, nomeBusca, vbTextCompare) > 0) Then
            
            ' Adicionando o cliente à ListBox
            Me.lstClientesListados.AddItem cel.Value
            Me.lstClientesListados.List(Me.lstClientesListados.ListCount - 1, 1) = cel.Offset(0, 1).Value ' Nome
            Me.lstClientesListados.List(Me.lstClientesListados.ListCount - 1, 2) = cel.Offset(0, 2).Value ' Contato
            
            encontrado = True
        End If
        Me.btnSelecionarCliente.Enabled = True
    Next cel
    
    ' Mensagem caso nenhum cliente seja encontrado
    If Not encontrado Then
        MsgBox "Nenhum cliente encontrado com os critérios fornecidos.", vbInformation
            Me.btnSelecionarCliente.Enabled = False
    End If


End Sub




Private Sub lstClientesListados_Click()
    If Me.lstClientesListados.ListIndex <> -1 Then
        ' Preencher os campos com as informações do cliente selecionado
        Me.txtID.Value = Me.lstClientesListados.List(Me.lstClientesListados.ListIndex, 0)
        Me.txtNomeCliente.Value = Me.lstClientesListados.List(Me.lstClientesListados.ListIndex, 1)
        
        ' Desabilitar o botão Selecionar Cliente após a seleção
        Me.btnSelecionarCliente.Enabled = True
        
        ' Desabilitar os campos para edição
        Me.txtID.Enabled = False
        Me.txtNomeCliente.Enabled = False
        
        ' Desabilitar a ListBox para impedir novas seleções
        Me.lstClientesListados.Enabled = False

        ' Desabilitar o botão Limpar e Buscar Cliente
        Me.btnBuscaCliente.Enabled = False
    End If
End Sub

' Ajuste no btnLimparCliente_Click para reabilitar todos os controles
Private Sub btnLimparCliente_Click()
    ' Limpar os campos
    Me.txtID.Value = ""
    Me.txtNomeCliente.Value = ""
    lstClientesListados.Clear
    
    ' Reabilitar os campos txtID e txtNomeCliente para edição
    Me.txtID.Enabled = True
    Me.txtNomeCliente.Enabled = True
    
    ' Reabilitar a ListBox para permitir novas seleções
    Me.lstClientesListados.Enabled = True
    Me.lstPropostasCliente.Clear


    ' Reabilitar os botões
    Me.btnSelecionarCliente.Enabled = False
    Me.btnLimparCliente.Enabled = True
    Me.btnBuscaCliente.Enabled = True

    ' Foco no nome
    txtNomeCliente.SetFocus
End Sub


Private Sub ListarPropostasCliente(clienteID As String)
    Dim wsPropostas As Worksheet
    Dim rngPropostas As Range
    Dim cel As Range
    Dim ultimaLinha As Long
    Dim numeroAtual As String
    Dim valorTotal As Double
    Dim referencia As String
    
    ' Definindo a planilha de propostas
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    
    ' Encontrando a última linha com dados
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).row
    
    ' Definindo o intervalo de dados das propostas
    Set rngPropostas = wsPropostas.Range("A2:H" & ultimaLinha)
    
    ' Limpando a ListBox antes de adicionar novos itens
    Me.lstPropostasCliente.Clear
    
    ' Configurando a ListBox para 3 colunas
    With Me.lstPropostasCliente
        .ColumnCount = 3
        .ColumnWidths = "80;100;120" ' Ajuste conforme necessário
    End With
    
    ' Iterando sobre cada linha da planilha de propostas
    numeroAtual = ""
    valorTotal = 0
    
    For Each cel In rngPropostas.Columns(2).Cells ' Coluna B para ID do cliente
        If cel.Value = clienteID Then
            If cel.Offset(0, -1).Value <> numeroAtual Then
                ' Nova proposta encontrada
                If numeroAtual <> "" Then
                    ' Adicionar proposta anterior à ListBox
                    Me.lstPropostasCliente.AddItem numeroAtual
                    Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 1) = Format(valorTotal, "#,##0.00")
                    Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 2) = referencia
                End If
                
                ' Iniciar nova proposta
                numeroAtual = cel.Offset(0, -1).Value
                valorTotal = cel.Offset(0, 5).Value ' Coluna G (SUBTOTAL)
                referencia = cel.Offset(0, 6).Value ' Coluna H (REFERENCIA)
            Else
                ' Continuar somando para a proposta atual
                valorTotal = valorTotal + cel.Offset(0, 5).Value
            End If
        End If
    Next cel
    
    ' Adicionar a última proposta à ListBox
    If numeroAtual <> "" Then
        Me.lstPropostasCliente.AddItem numeroAtual
        Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 1) = Format(valorTotal, "#,##0.00")
        Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 2) = referencia
    End If
    
    ' Habilitar a ListBox de propostas
    Me.lstPropostasCliente.Enabled = True
End Sub


Private Sub btnSelecionarCliente_Click()
    If Me.lstClientesListados.ListIndex <> -1 Then
        Dim clienteID As String
        clienteID = Me.lstClientesListados.List(Me.lstClientesListados.ListIndex, 0)
        
        ' Chamar a função para listar as propostas do cliente
        ListarPropostasCliente clienteID
        
        ' Desabilitar o botão após a seleção
        Me.btnSelecionarCliente.Enabled = False
    Else
        MsgBox "Por favor, selecione um cliente primeiro.", vbExclamation
    End If
End Sub




'#####################################

'Analise o codigo acima pois irei precisar de ajuda em uma implementacao




' txtID
' txtNomeCliente
' btnBuscaCliente
' btnLimparCliente
' lstClientesListados
' btnSelecionarCliente
' lstPropostasCliente
' btnSelecionarProposta

' txtNrProposta
' txtReferencia
' btnAtualizarRef

' btnBuscarProduto
' txtCodProduto
' txtDescricao
' txtQTD
' txtPreco
' txtItem
' btnAdicionarProduto

