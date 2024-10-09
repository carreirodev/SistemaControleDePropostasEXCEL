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
        .ColumnWidths = "45;58;120" ' Ajuste conforme necessário
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
    
    ' Definindo a planilha de propostas
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    
    ' Encontrando a última linha com dados
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).Row
    
    ' Definindo o intervalo de dados das propostas
    Set rngPropostas = wsPropostas.Range("A2:H" & ultimaLinha)
    
    ' Limpar o ListView antes de adicionar novos itens
    Me.lvwProdutosDaProposta.ListItems.Clear
    
    ' Configurar as colunas do ListView (se ainda não estiver configurado)
    With Me.lvwProdutosDaProposta
        .View = lvwReport
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Item", 50
        .ColumnHeaders.Add , , "Código", 80
        .ColumnHeaders.Add , , "Preço Unitário", 100
        .ColumnHeaders.Add , , "Quantidade", 80
        .ColumnHeaders.Add , , "Subtotal", 100
    End With
    
    ' Iterando sobre cada linha da planilha de propostas
    For Each cel In rngPropostas.Columns(1).Cells ' Coluna A para NUMERO
        If cel.Value = numeroPropostaSelecionada Then
            Dim lstItem As ListItem
            Set lstItem = Me.lvwProdutosDaProposta.ListItems.Add(, , cel.Offset(0, 2).Value) ' ITEM
            lstItem.SubItems(1) = cel.Offset(0, 3).Value ' CODIGO
            lstItem.SubItems(2) = Format(cel.Offset(0, 4).Value, "#,##0.00") ' PRECO UNITARIO
            lstItem.SubItems(3) = cel.Offset(0, 5).Value ' QUANTIDADE
            lstItem.SubItems(4) = Format(cel.Offset(0, 6).Value, "#,##0.00") ' SUBTOTAL
        End If
    Next cel
    
    ' Ajustar o tamanho das colunas para se ajustarem ao conteúdo
    ' AjustarColunasListView Me.lvwProdutosDaProposta
End Sub



' Private Sub AjustarColunasListView(lv As ListView)
'     Dim col As ColumnHeader
'     Dim i As Long
'     Dim larguraMaxima As Long
'     Dim item As ListItem
'     Dim textoItem As String
    
'     lv.Visible = False ' Esconde temporariamente o ListView para melhorar o desempenho
    
'     For i = 1 To lv.ColumnHeaders.Count
'         larguraMaxima = 0
        
'         ' Verifica a largura do cabeçalho
'         larguraMaxima = Max(larguraMaxima, TextWidth(lv.ColumnHeaders(i).Text))
        
'         ' Verifica a largura de cada item na coluna
'         For Each item In lv.ListItems
'             If i = 1 Then
'                 textoItem = item.Text
'             Else
'                 textoItem = item.SubItems(i - 1)
'             End If
'             larguraMaxima = Max(larguraMaxima, TextWidth(textoItem))
'         Next item
        
'         ' Adiciona um pequeno espaço extra e define a largura da coluna
'         lv.ColumnHeaders(i).Width = larguraMaxima + 20
'     Next i
    
'     lv.Visible = True ' Torna o ListView visível novamente
' End Sub

' ' Função auxiliar para obter o máximo entre dois valores
' Private Function Max(a As Long, b As Long) As Long
'     If a > b Then
'         Max = a
'     Else
'         Max = b
'     End If
' End Function

' ' Função auxiliar para calcular a largura do texto
' Private Function TextWidth(texto As String) As Long
'     On Error Resume Next
'     Dim tmpControl As Control
'     Set tmpControl = Me.Controls.Add("Forms.Label.1", "tmpLabel")
'     If Err.Number = 0 Then
'         With tmpControl
'             .AutoSize = True
'             .Caption = texto
'             TextWidth = .Width
'         End With
'         Me.Controls.Remove "tmpLabel"
'     Else
'         ' Alternativa se não puder adicionar controle
'         TextWidth = Len(texto) * 7 ' Estimativa aproximada
'     End If
'     On Error GoTo 0
' End Function







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

