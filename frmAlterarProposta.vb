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
    Dim qtdItens As Integer
    
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).row
    Set rngPropostas = wsPropostas.Range("A2:H" & ultimaLinha)
    
    Me.lstPropostasCliente.Clear
    
    With Me.lstPropostasCliente
        .ColumnCount = 4
        .ColumnWidths = "45;58;40;120"
    End With
    
    numeroAtual = ""
    valorTotal = 0
    qtdItens = 0
    
    For Each cel In rngPropostas.Columns(2).Cells
        If cel.Value = clienteID Then
            If cel.Offset(0, -1).Value <> numeroAtual Then
                If numeroAtual <> "" Then
                    Me.lstPropostasCliente.AddItem numeroAtual
                    Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 1) = Format(valorTotal, "#,##0.00")
                    Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 2) = qtdItens
                    Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 3) = referencia
                End If
                
                numeroAtual = cel.Offset(0, -1).Value
                valorTotal = cel.Offset(0, 5).Value
                referencia = cel.Offset(0, 6).Value
                qtdItens = 1
            Else
                valorTotal = valorTotal + cel.Offset(0, 5).Value
                qtdItens = qtdItens + 1
            End If
        End If
    Next cel
    
    If numeroAtual <> "" Then
        Me.lstPropostasCliente.AddItem numeroAtual
        Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 1) = Format(valorTotal, "#,##0.00")
        Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 2) = qtdItens
        Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListCount - 1, 3) = referencia
    End If
    
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


Private Sub lstPropostasCliente_Click()
    If Me.lstPropostasCliente.ListIndex <> -1 Then
        Dim numeroPropostaSelecionada As String
        numeroPropostaSelecionada = Me.lstPropostasCliente.List(Me.lstPropostasCliente.ListIndex, 0)
        ExibirDetalhesProposta numeroPropostaSelecionada
    End If
End Sub


Private Sub ExibirDetalhesProposta(numeroPropostaSelecionada As String)
    Dim wsPropostas As Worksheet
    Dim rngPropostas As Range
    Dim cel As Range
    Dim ultimaLinha As Long
    Dim item As listItem
    Dim referenciaProposta As String
    
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).row
    Set rngPropostas = wsPropostas.Range("A2:H" & ultimaLinha)
    
    lvwProdutosDaProposta.ListItems.Clear
    
    With lvwProdutosDaProposta
        .View = lvwReport
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Item", 40
        .ColumnHeaders.Add , , "Código", 60
        .ColumnHeaders.Add , , "Descrição", 150
        .ColumnHeaders.Add , , "Preço Unitário", 80
        .ColumnHeaders.Add , , "Qtd", 40
        .ColumnHeaders.Add , , "Subtotal", 80
    End With
    
    For Each cel In rngPropostas.Columns(1).Cells
        If cel.Value = numeroPropostaSelecionada Then
            Set item = lvwProdutosDaProposta.ListItems.Add(, , cel.Offset(0, 2).Value) ' ITEM
            item.SubItems(1) = cel.Offset(0, 3).Value ' CODIGO
            item.SubItems(2) = BuscarDetalheProduto(cel.Offset(0, 3).Value) ' Descrição do produto
            item.SubItems(3) = Format(cel.Offset(0, 4).Value, "#,##0.00") ' PRECO UNITARIO
            item.SubItems(4) = cel.Offset(0, 5).Value ' QUANTIDADE
            item.SubItems(5) = Format(cel.Offset(0, 6).Value, "#,##0.00") ' SUBTOTAL
            
            ' Capturar a referência da proposta (assumindo que é a mesma para todos os itens da proposta)
            referenciaProposta = cel.Offset(0, 7).Value ' REFERENCIA (Coluna H)
        End If
    Next cel
    
    ' Preencher o txtReferencia com a referência da proposta
    Me.txtReferencia.Value = referenciaProposta
    
    ' Ajustar o tamanho das colunas automaticamente
    AjustarColunasListView lvwProdutosDaProposta
End Sub


Private Function BuscarDetalheProduto(codigoProduto As String) As String
    Dim wsPrecos As Worksheet
    Dim rngPrecos As Range
    Dim cel As Range
    
    Set wsPrecos = ThisWorkbook.Sheets("ListaDePrecos")
    Set rngPrecos = wsPrecos.UsedRange
    
    For Each cel In rngPrecos.Columns(1).Cells ' Assumindo que o código está na primeira coluna
        If cel.Value = codigoProduto Then
            ' Retorna a descrição do produto (assumindo que está na segunda coluna)
            BuscarDetalheProduto = cel.Offset(0, 1).Value
            Exit Function
        End If
    Next cel
    
    BuscarDetalheProduto = "Produto não encontrado"
End Function



Private Sub AjustarColunasListView(lv As ListView)
    Dim col As ColumnHeader
    Dim maxWidth As Long
    Dim item As listItem
    Dim subItem As ListSubItem
    Dim i As Long
    Dim totalWidth As Long
    Dim avgCharWidth As Long
    Const FIXED_COL_WIDTH As Long = 45  ' Largura fixa para ITEM, CODIGO e QTD
    Const MAX_COL_WIDTH As Long = 230   ' Largura máxima para outras colunas
    Const MIN_COL_WIDTH As Long = 45    ' Largura mínima para outras colunas
    
    ' Estimar a largura média de um caractere
    avgCharWidth = lv.Font.Size * 0.7
    
    totalWidth = 0
    For Each col In lv.ColumnHeaders
        Select Case col.Index
            Case 1, 2, 5 ' ITEM, CODIGO, QTD
                col.Width = FIXED_COL_WIDTH
                totalWidth = totalWidth + FIXED_COL_WIDTH
            Case Else ' Outras colunas (Descrição, Preço Unitário, Subtotal)
                maxWidth = col.Width
                For Each item In lv.ListItems
                    maxWidth = WorksheetFunction.Max(maxWidth, avgCharWidth * Len(item.SubItems(col.Index - 1)))
                Next item
                
                ' Limitar a largura entre o mínimo e o máximo
                maxWidth = WorksheetFunction.Max(MIN_COL_WIDTH, WorksheetFunction.Min(maxWidth + 15, MAX_COL_WIDTH))
                
                col.Width = maxWidth
                totalWidth = totalWidth + maxWidth
        End Select
    Next col
    
    ' Se a largura total for menor que a largura do ListView, ajustar proporcionalmente as colunas não fixas
    If totalWidth < lv.Width Then
        Dim remainingWidth As Long
        Dim nonFixedWidth As Long
        Dim ratio As Double
        
        remainingWidth = lv.Width - (FIXED_COL_WIDTH * 3) ' Subtrair a largura das colunas fixas
        nonFixedWidth = totalWidth - (FIXED_COL_WIDTH * 3)
        
        If nonFixedWidth > 0 Then
            ratio = remainingWidth / nonFixedWidth
            For Each col In lv.ColumnHeaders
                If col.Index <> 1 And col.Index <> 2 And col.Index <> 5 Then
                    col.Width = col.Width * ratio
                End If
            Next col
        End If
    End If
End Sub


Private Sub btnSelecionarProposta_Click()
    If lstPropostasCliente.ListIndex = -1 Then
        MsgBox "Por favor, selecione uma proposta antes de continuar.", vbExclamation
        Exit Sub
    End If
    
    ' Obter o número da proposta selecionada
    Dim numeroPropostaSelecionada As String
    numeroPropostaSelecionada = lstPropostasCliente.List(lstPropostasCliente.ListIndex, 0)
    
    ' Exibir detalhes da proposta
    ExibirDetalhesProposta numeroPropostaSelecionada
    
    ' Desabilitar a lista de propostas
    lstPropostasCliente.Enabled = False
    
    ' Desabilitar o botão de selecionar proposta
    btnSelecionarProposta.Enabled = False

    ' Atualizar a interface para refletir que uma proposta foi selecionada para edição
    AtualizarInterfacePropostaSelecionada
End Sub


Private Sub AtualizarInterfacePropostaSelecionada()
    
    ' Desabilitar outros controles que não devem ser usados durante a edição
    btnSelecionarCliente.Enabled = False
    btnBuscaCliente.Enabled = False
    
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



