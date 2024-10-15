Private Sub UserForm_Initialize()
    

    ' Configurando o ListBox para ter colunas
    With Me.lstProdutosDaProposta
        .ColumnCount = 6
        .ColumnWidths = "35;55;320;35;85;85"
        .AddItem "Item"
        .List(0, 1) = "Código"
        .List(0, 2) = "Descrição"
        .List(0, 3) = "Qtd"
        .List(0, 4) = "$ Unitário"
        .List(0, 5) = "$ Total"
    End With
    
    ' Configurar a ListBox lstCliente (mantida como está)
    With Me.lstCliente
        .ColumnCount = 5
        .ColumnWidths = "50;200;140;70;24"
    End With
    
    ' Carregar os vendedores na ComboBox cmbVendedor
    Dim wsVendedores As Worksheet
    Dim rngVendedores As Range
    Dim cel As Range
    
    ' Definindo a planilha de vendedores
    Set wsVendedores = ThisWorkbook.Sheets("VENDEDORES")
    ' Definindo o intervalo de dados dos vendedores (ajuste conforme necessário)
    Set rngVendedores = wsVendedores.ListObjects("Vendedor").DataBodyRange
    
    ' Limpando a ComboBox antes de adicionar novos itens
    Me.cmbVendedor.Clear
    
    ' Iterando sobre cada vendedor e adicionando à ComboBox
    For Each cel In rngVendedores.Columns(1).Cells
        Me.cmbVendedor.AddItem cel.Value
    Next cel

    ' Carregar as condições de pagamento na ComboBox cmbCondPagamento
    Dim wsCondPagto As Worksheet
    Dim rngCondPagto As Range
    
    ' Definindo a planilha de condições de pagamento
    Set wsCondPagto = ThisWorkbook.Sheets("CondPagto")
    ' Definindo o intervalo de dados das condições de pagamento
    Set rngCondPagto = wsCondPagto.ListObjects("CondPagto").DataBodyRange
    
    ' Limpando a ComboBox antes de adicionar novos itens
    Me.cmbCondPagamento.Clear
    
    ' Iterando sobre cada condição de pagamento e adicionando à ComboBox
    For Each cel In rngCondPagto.Columns(1).Cells
        Me.cmbCondPagamento.AddItem cel.Value
    Next cel

    ' Desabilitar o botão Salvar Proposta por padrão
    Me.btnSalvarProposta.Enabled = False    

    ' Inicializar o valor total
    Me.txtValorTotal.Value = "0.00"

End Sub


Private Sub lstCliente_Click()
    If Me.lstCliente.ListIndex <> -1 Then
        ' Preenche os campos com as informações do cliente selecionado
        Me.txtID.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 0) ' ID
        Me.txtNomeCliente.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 1) ' Nome
        Me.txtPessoaContato.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 2) ' Contato
        Me.txtCidade.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 3) ' Cidade
        Me.txtEstado.Value = Me.lstCliente.List(Me.lstCliente.ListIndex, 4) ' Estado
        
        ' Desabilitar os campos para edição
        Me.txtID.Enabled = False
        Me.txtNomeCliente.Enabled = False
        Me.txtPessoaContato.Enabled = False
        Me.txtCidade.Enabled = False
        Me.txtEstado.Enabled = False
        
        ' Desabilitar a ListBox para impedir novas seleções
        Me.lstCliente.Enabled = False

        ' Desabilitar o botão Limpar e Buscar Cliente
        ' Me.btnLimparCliente.Enabled = False
        Me.btnBuscaCliente.Enabled = False

        ' Gerar novo número de proposta e registrar na planilha
        CriarNovaProposta
        
        ' Carregar as propostas do cliente selecionado
        CarregarPropostasDoCliente Me.txtID.Value
    Else
        ' Se nenhum item estiver selecionado, manter os campos editáveis
        Me.txtID.Enabled = True
        Me.txtNomeCliente.Enabled = True
        Me.txtPessoaContato.Enabled = True
        Me.txtCidade.Enabled = True
        Me.txtEstado.Enabled = True
    End If

    ' Verificar se pode habilitar o botão Salvar Proposta
    VerificarSalvarProposta
End Sub

Private Sub CarregarPropostasDoCliente(clienteID As String)
    Dim wsPropostas As Worksheet
    Dim rngPropostas As Range
    Dim cel As Range
    Dim ultimaLinha As Long
    Dim propostaJaAdicionada As Collection
    Dim valorTotal As Double
    
    ' Limpar o ListBox de propostas do cliente
    Me.lstPropostasDoCliente.Clear
    
    ' Configurar as colunas do ListBox
    With Me.lstPropostasDoCliente
        .ColumnCount = 4
        .ColumnWidths = "70;100;100;100"
        .AddItem "Número"
        .List(0, 1) = "Referência"
        .List(0, 2) = "Vendedor"
        .List(0, 3) = "Valor Total"
    End With
    
    ' Definindo a planilha de propostas
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    
    ' Encontrar a última linha com dados
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).Row
    
    ' Definindo o intervalo de dados das propostas
    Set rngPropostas = wsPropostas.Range("A2:K" & ultimaLinha)
    
    ' Criar uma coleção para rastrear propostas já adicionadas
    Set propostaJaAdicionada = New Collection
    
    ' Iterando sobre cada proposta
    For Each cel In rngPropostas.Columns(2).Cells ' Coluna B para CLIENTE
        If cel.Value = clienteID Then
            Dim numeroProposta As String
            numeroProposta = cel.Offset(0, -1).Value ' Coluna A para NUMERO
            
            ' Verificar se esta proposta já foi adicionada
            On Error Resume Next
            propostaJaAdicionada.Add numeroProposta, CStr(numeroProposta)
            If Err.Number = 0 Then ' Se não houve erro, a proposta é nova
                ' Calcular o valor total da proposta
                valorTotal = CalcularValorTotalProposta(numeroProposta)
                
                ' Adicionar a proposta ao ListBox
                Me.lstPropostasDoCliente.AddItem numeroProposta
                Me.lstPropostasDoCliente.List(Me.lstPropostasDoCliente.ListCount - 1, 1) = cel.Offset(0, 6).Value  ' Coluna H para REFERENCIA
                Me.lstPropostasDoCliente.List(Me.lstPropostasDoCliente.ListCount - 1, 2) = cel.Offset(0, 7).Value  ' Coluna I para VENDEDOR
                Me.lstPropostasDoCliente.List(Me.lstPropostasDoCliente.ListCount - 1, 3) = Format(valorTotal, "#,##0.00")  ' Valor Total calculado
            End If
            On Error GoTo 0 ' Restaurar o tratamento de erro padrão
        End If
    Next cel
End Sub

Private Function CalcularValorTotalProposta(numeroProposta As String) As Double
    Dim wsPropostas As Worksheet
    Dim rngPropostas As Range
    Dim cel As Range
    Dim valorTotal As Double
    Dim ultimaLinha As Long
    
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).Row
    Set rngPropostas = wsPropostas.Range("A2:G" & ultimaLinha)
    
    valorTotal = 0
    
    For Each cel In rngPropostas.Columns(1).Cells ' Coluna A para NUMERO
        If cel.Value = numeroProposta Then
            valorTotal = valorTotal + cel.Offset(0, 6).Value ' Coluna G para SUBTOTAL
        End If
    Next cel
    
    CalcularValorTotalProposta = valorTotal
End Function








Private Sub VerificarSalvarProposta()
    If Me.txtID.Value <> "" And Me.lstProdutosDaProposta.ListCount > 1 And Not propostaExistenteCarregada Then
        Me.btnSalvarProposta.Enabled = True
    Else
        Me.btnSalvarProposta.Enabled = False
    End If
End Sub



Private Sub btnFechar_Click()
    Unload Me
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

Private Sub btnLimparCliente_Click()
    ' Limpar os campos
    Me.txtID.Value = ""
    Me.txtNomeCliente.Value = ""
    Me.txtPessoaContato.Value = ""
    Me.txtCidade.Value = ""
    Me.txtEstado.Value = ""
    Me.txtNrProposta.Value = ""
    lstCliente.Clear
    lstPropostasDoCliente.Clear
    
    ' Reabilitar os campos txtID e txtNomeCliente para edição
    Me.txtID.Enabled = True
    Me.txtNomeCliente.Enabled = True
    
    ' Desabilitar os outros campos
    Me.txtPessoaContato.Enabled = False
    Me.txtCidade.Enabled = False
    Me.txtEstado.Enabled = False
    
    
    ' Reabilitar a ListBox para permitir novas seleções
    Me.lstCliente.Enabled = True
    Me.btnBuscaCliente.Enabled = True

    lstPropostasDoCliente.Clear

    ' Foco no nome
    txtNomeCliente.SetFocus

        Me.txtNrProposta.Value = ""
    Me.txtReferencia.Value = ""
    Me.cmbVendedor.Value = ""
    Me.txtPrazoEntrega.Value = ""
    Me.cmbCondPagamento.Value = ""
    Me.txtValorTotal.Value = ""
    Me.lstProdutosDaProposta.Clear
    
    ' Reabilitar o botão Salvar Nova Proposta e desabilitar o botão Alterar Proposta
    Me.btnSalvarProposta.Enabled = True
    Me.btnAlterarProposta.Enabled = False

        
    VerificarSalvarProposta
    
End Sub


Private Sub CriarNovaProposta()
    Dim wsPropostas As Worksheet
    Dim numeroBase As Long
    Dim novoNumero As String
    Dim estadoCliente As String
    
    ' Definindo a planilha de propostas
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    
    ' Obter o último número base da proposta da célula N1
    numeroBase = wsPropostas.Range("N1").Value
    
    ' Incrementar o número base
    numeroBase = numeroBase + 1
    
    ' Atualizar a célula N1 com o novo número base
    wsPropostas.Range("N1").Value = numeroBase
    
    ' Formatar o número da proposta com quatro dígitos
    novoNumero = Format(numeroBase, "0000")
    
    ' Obter o estado do cliente
    estadoCliente = Me.txtEstado.Value
    
    ' Concatenar o número formatado com o estado do cliente
    novoNumero = novoNumero & "-" & estadoCliente
    
    ' Preencher o número da proposta no campo txtNrProposta
    Me.txtNrProposta.Value = novoNumero
End Sub

Private Sub btnBuscarProduto_Click()
    Dim wsPrecos As Worksheet
    Dim rngPrecos As Range
    Dim cel As Range
    Dim codigoBusca As String
    Dim encontrado As Boolean
    
    ' Definindo a planilha de preços
    Set wsPrecos = ThisWorkbook.Sheets("ListaDePrecos")
    ' Definindo o intervalo de dados dos produtos
    Set rngPrecos = wsPrecos.ListObjects("TabPrecos").DataBodyRange
    
    ' Obtendo o código do produto
    codigoBusca = Trim(Me.txtCodProduto.Value)
    
    ' Iterando sobre cada produto
    For Each cel In rngPrecos.Columns(1).Cells ' Coluna Produto
        ' Verificando se o código corresponde
        If cel.Value = codigoBusca Then
            ' Preenchendo os campos com as informações do produto
            Me.txtDescricao.Value = cel.Offset(0, 1).Value ' Descrição
            Me.txtPreco.Value = Format(cel.Offset(0, 2).Value, "#,##0.00") ' Preço
            Me.txtQTD.Value = 1 ' Preencher quantidade com 1
            
            ' Verificar se o campo txtItem está vazio e definir como 1 apenas se estiver
            If Me.txtItem.Value = "" Then
                Me.txtItem.Value = 1
            End If
            
            encontrado = True
            Exit For
        End If
    Next cel
    
    ' Mensagem caso o produto não seja encontrado
    If Not encontrado Then
        MsgBox "Produto não encontrado.", vbInformation
    End If
End Sub

Private Sub btnAdicionarProduto_Click()
    ' Verificar se o número da proposta está preenchido
    If Me.txtNrProposta.Value = "" Then
        MsgBox "Selecione um cliente antes de adicionar produtos à proposta.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se o cliente foi selecionado
    If Me.txtID.Value = "" Or Me.txtID.Enabled = True Then
        MsgBox "Selecione um cliente antes de adicionar produtos à proposta.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se o código do produto está preenchido
    If Me.txtCodProduto.Value = "" Then
        MsgBox "Insira o código do produto antes de adicionar à proposta.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se a descrição do produto está preenchida
    If Me.txtDescricao.Value = "" Then
        MsgBox "Insira a descrição do produto antes de adicionar à proposta.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se o preço do produto está preenchido
    If Me.txtPreco.Value = "" Then
        MsgBox "Insira o preço do produto antes de adicionar à proposta.", vbExclamation
        Exit Sub
    End If
    
    ' Verificar se a quantidade está preenchida
    If Me.txtQTD.Value = "" Then
        MsgBox "Insira a quantidade antes de adicionar à proposta.", vbExclamation
        Exit Sub
    End If

    ' Continuar com a adição do produto
    Dim Item As Long
    Dim codigo As String
    Dim descricao As String
    Dim precoUnitario As Double
    Dim quantidade As Long
    Dim subtotal As Double
    
    ' Obtendo os valores dos campos
    Item = CLng(Me.txtItem.Value)
    codigo = Me.txtCodProduto.Value
    descricao = Me.txtDescricao.Value
    precoUnitario = CDbl(Me.txtPreco.Value)
    quantidade = CLng(Me.txtQTD.Value)
    subtotal = precoUnitario * quantidade
    
    ' Adicionar o item ao ListBox
    Me.lstProdutosDaProposta.AddItem Item
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 1) = codigo
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 2) = descricao
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 3) = quantidade
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 4) = Format(precoUnitario, "#,##0.00")
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 5) = Format(subtotal, "#,##0.00")
    
    ' Limpar os campos de entrada
    Me.txtCodProduto.Value = ""
    Me.txtDescricao.Value = ""
    Me.txtPreco.Value = ""
    Me.txtQTD.Value = ""
    Me.txtItem.Value = ""
    
    ' Reposicionar o cursor para o campo txtCodProduto
    Me.txtCodProduto.SetFocus
    
    ' Incrementar o número do item para o próximo produto
    Me.txtItem.Value = Item + 1

    ' Verificar se pode habilitar o botão Salvar Proposta
    VerificarSalvarProposta

    ' Atualizar o valor total
    AtualizarValorTotal
End Sub


Private Sub btnRemoverProduto_Click()
    If Me.lstProdutosDaProposta.ListIndex > 0 Then ' Não remover o cabeçalho
        Me.lstProdutosDaProposta.RemoveItem Me.lstProdutosDaProposta.ListIndex
        AtualizarValorTotal
        VerificarSalvarProposta
    Else
        MsgBox "Selecione um produto para remover.", vbExclamation
    End If
End Sub




Private Sub ValidarProduto()
    Dim ws As Worksheet
    Dim rng As Range
    Dim produtoEncontrado As Boolean
    Dim descricaoCorreta As Boolean
    
    ' Define a planilha "ListaDePrecos"
    Set ws = ThisWorkbook.Sheets("ListaDePrecos")
    
    ' Procura o código do produto na tabela "TabPrecos"
    Set rng = ws.Range("TabPrecos").Columns("A").Find(What:=Trim(txtCodProduto.Value), LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Verifica se o produto foi encontrado
    produtoEncontrado = Not rng Is Nothing
    
    ' Verifica se a descrição corresponde, somente se o produto foi encontrado
    If produtoEncontrado Then
        descricaoCorreta = (rng.Offset(0, 1).Value = Trim(txtDescricao.Value))
    Else
        descricaoCorreta = False
    End If
    
    ' Habilita o botão apenas se o produto for encontrado e a descrição estiver correta
    btnAdicionarProduto.Enabled = produtoEncontrado And descricaoCorreta
End Sub

Private Sub txtCodProduto_Change()
    ValidarProduto
End Sub

Private Sub txtDescricao_Change()
    ValidarProduto
End Sub

Private Sub btnSalvarProposta_Click()
    Dim wsPropostas As Worksheet
    Dim numeroProposta As String
    Dim novaReferencia As String
    Dim ultimaLinha As Long
    Dim i As Long
    Dim vendedor As String
    Dim condicaoPagamento As String
    Dim prazoEntrega As String
    
    ' Definindo a planilha de propostas
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    
    ' Obtendo o número da proposta e a nova referência
    numeroProposta = Me.txtNrProposta.Value
    novaReferencia = Me.txtReferencia.Value
    vendedor = Me.cmbVendedor.Value
    condicaoPagamento = Me.cmbCondPagamento.Value
    prazoEntrega = Me.txtPrazoEntrega.Value
    
    ' Limpar itens antigos da proposta na planilha
    For i = wsPropostas.Cells(wsPropostas.Rows.Count, 1).End(xlUp).Row To 2 Step -1
        If wsPropostas.Cells(i, 1).Value = numeroProposta Then
            wsPropostas.Rows(i).Delete
        End If
    Next i
    
    ' Encontrar a próxima linha vazia para registrar a nova proposta
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Iterar sobre os itens do ListBox e adicionar à planilha
    For i = 1 To Me.lstProdutosDaProposta.ListCount - 1 ' Começando de 1 para pular o cabeçalho
        wsPropostas.Cells(ultimaLinha, 1).Value = numeroProposta ' Coluna NUMERO
        wsPropostas.Cells(ultimaLinha, 2).Value = Me.txtID.Value ' Coluna CLIENTE
        wsPropostas.Cells(ultimaLinha, 3).Value = Me.lstProdutosDaProposta.List(i, 0) ' Coluna ITEM
        wsPropostas.Cells(ultimaLinha, 4).Value = Me.lstProdutosDaProposta.List(i, 1) ' Coluna CODIGO
        wsPropostas.Cells(ultimaLinha, 5).Value = CDbl(Me.lstProdutosDaProposta.List(i, 4)) ' Coluna PRECO UNITARIO
        wsPropostas.Cells(ultimaLinha, 6).Value = CLng(Me.lstProdutosDaProposta.List(i, 3)) ' Coluna QUANTIDADE
        wsPropostas.Cells(ultimaLinha, 7).Value = CDbl(Me.lstProdutosDaProposta.List(i, 5)) ' Coluna SUBTOTAL
        wsPropostas.Cells(ultimaLinha, 8).Value = novaReferencia ' Coluna REFERENCIA
        wsPropostas.Cells(ultimaLinha, 9).Value = vendedor ' Coluna VENDEDOR
        wsPropostas.Cells(ultimaLinha, 10).Value = condicaoPagamento ' Coluna CONDICAO DE PAGAMENTO
        wsPropostas.Cells(ultimaLinha, 11).Value = prazoEntrega ' Coluna PRAZO DE ENTREGA
        ultimaLinha = ultimaLinha + 1
    Next i
    
    MsgBox "Proposta salva com sucesso!", vbInformation

    Unload Me
End Sub

Private Sub AtualizarValorTotal()
    Dim i As Long
    Dim valorTotal As Double
    
    valorTotal = 0
    
    ' Iterar sobre todos os itens na lstProdutosDaProposta, exceto o cabeçalho
    For i = 1 To Me.lstProdutosDaProposta.ListCount - 1
        ' Somar o valor total de cada item (coluna 5)
        valorTotal = valorTotal + CDbl(Me.lstProdutosDaProposta.List(i, 5))
    Next i
    
    ' Atualizar o txtValorTotal com o valor calculado
    Me.txtValorTotal.Value = Format(valorTotal, "#,##0.00")
End Sub


Private Sub lstPropostasDoCliente_Click()
    If Me.lstPropostasDoCliente.ListIndex > 0 Then ' Ignorar o cabeçalho
        Dim numeroProposta As String
        numeroProposta = Me.lstPropostasDoCliente.List(Me.lstPropostasDoCliente.ListIndex, 0)
        
        CarregarDetalhesPropostaExistente numeroProposta
        
        ' Desabilitar o botão Salvar Nova Proposta e habilitar o botão Alterar Proposta
        Me.btnSalvarProposta.Enabled = False
        Me.btnAlterarProposta.Enabled = True
    End If
End Sub

Private Sub CarregarDetalhesPropostaExistente(numeroProposta As String)
    Dim wsPropostas As Worksheet
    Dim rngPropostas As Range
    Dim cel As Range
    Dim ultimaLinha As Long
    
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).Row
    Set rngPropostas = wsPropostas.Range("A2:K" & ultimaLinha)
    
    ' Limpar a lista de produtos da proposta
    Me.lstProdutosDaProposta.Clear
    
    ' Adicionar cabeçalho à lista de produtos
    With Me.lstProdutosDaProposta
        .AddItem "Item"
        .List(0, 1) = "Código"
        .List(0, 2) = "Descrição"
        .List(0, 3) = "Qtd"
        .List(0, 4) = "$ Unitário"
        .List(0, 5) = "$ Total"
    End With
    
    Dim valorTotal As Double
    valorTotal = 0
    
    For Each cel In rngPropostas.Columns(1).Cells ' Coluna A para NUMERO
        If cel.Value = numeroProposta Then
            ' Preencher informações gerais da proposta
            Me.txtNrProposta.Value = numeroProposta
            Me.txtReferencia.Value = cel.Offset(0, 7).Value ' Coluna H para REFERENCIA
            Me.cmbVendedor.Value = cel.Offset(0, 8).Value ' Coluna I para VENDEDOR
            Me.txtPrazoEntrega.Value = cel.Offset(0, 10).Value ' Coluna K para PRAZO DE ENTREGA
            Me.cmbCondPagamento.Value = cel.Offset(0, 9).Value ' Coluna J para CONDICAO DE PAGAMENTO
            
            ' Adicionar item à lista de produtos da proposta
            Me.lstProdutosDaProposta.AddItem cel.Offset(0, 2).Value ' Coluna C para ITEM
            Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 1) = cel.Offset(0, 3).Value ' Coluna D para CODIGO
            Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 2) = ObterDescricaoProduto(cel.Offset(0, 3).Value) ' Obter descrição do produto
            Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 3) = cel.Offset(0, 5).Value ' Coluna F para QUANTIDADE
            Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 4) = Format(cel.Offset(0, 4).Value, "#,##0.00") ' Coluna E para PRECO UNITARIO
            Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 5) = Format(cel.Offset(0, 6).Value, "#,##0.00") ' Coluna G para SUBTOTAL
            
            valorTotal = valorTotal + cel.Offset(0, 6).Value
        End If
    Next cel
    
    ' Atualizar o valor total
    Me.txtValorTotal.Value = Format(valorTotal, "#,##0.00")

    Me.btnSalvarProposta.Enabled = False
    Me.btnAlterarProposta.Enabled = True

    propostaExistenteCarregada = True
    Me.btnSalvarProposta.Enabled = False
    Me.btnAlterarProposta.Enabled = True

End Sub

Private Function ObterDescricaoProduto(codigoProduto As String) As String
    Dim wsPrecos As Worksheet
    Dim rngPrecos As Range
    Dim cel As Range
    
    Set wsPrecos = ThisWorkbook.Sheets("ListaDePrecos")
    Set rngPrecos = wsPrecos.ListObjects("TabPrecos").DataBodyRange
    
    For Each cel In rngPrecos.Columns(1).Cells ' Coluna Produto
        If cel.Value = codigoProduto Then
            ObterDescricaoProduto = cel.Offset(0, 1).Value ' Coluna Descrição
            Exit Function
        End If
    Next cel
    
    ObterDescricaoProduto = "Descrição não encontrada"
End Function


Private Sub btnAlterarProposta_Click()
    Dim wsPropostas As Worksheet
    Dim rngPropostas As Range
    Dim cel As Range
    Dim numeroProposta As String
    Dim ultimaLinha As Long
    Dim i As Long
    
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).Row
    Set rngPropostas = wsPropostas.Range("A2:K" & ultimaLinha)
    
    numeroProposta = Me.txtNrProposta.Value
    
    ' Remover todos os registros antigos da proposta
    Application.ScreenUpdating = False
    For i = rngPropostas.Rows.Count To 1 Step -1
        If rngPropostas.Cells(i, 1).Value = numeroProposta Then
            rngPropostas.Rows(i).Delete
        End If
    Next i
    
    ' Adicionar novos registros da proposta atualizada no final da planilha
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).Row + 1
    
    For i = 1 To Me.lstProdutosDaProposta.ListCount - 1 ' Começando de 1 para pular o cabeçalho
        wsPropostas.Cells(ultimaLinha, 1).Value = numeroProposta
        wsPropostas.Cells(ultimaLinha, 2).Value = Me.txtID.Value
        wsPropostas.Cells(ultimaLinha, 3).Value = Me.lstProdutosDaProposta.List(i, 0)
        wsPropostas.Cells(ultimaLinha, 4).Value = Me.lstProdutosDaProposta.List(i, 1)
        wsPropostas.Cells(ultimaLinha, 5).Value = Format(CDbl(Me.lstProdutosDaProposta.List(i, 4)), "#,##0.00")
        wsPropostas.Cells(ultimaLinha, 6).Value = CLng(Me.lstProdutosDaProposta.List(i, 3))
        wsPropostas.Cells(ultimaLinha, 7).Value = Format(CDbl(Me.lstProdutosDaProposta.List(i, 5)), "#,##0.00")
        wsPropostas.Cells(ultimaLinha, 8).Value = Me.txtReferencia.Value
        wsPropostas.Cells(ultimaLinha, 9).Value = Me.cmbVendedor.Value
        wsPropostas.Cells(ultimaLinha, 10).Value = Me.cmbCondPagamento.Value
        wsPropostas.Cells(ultimaLinha, 11).Value = Me.txtPrazoEntrega.Value
        ultimaLinha = ultimaLinha + 1
    Next i
    
    Application.ScreenUpdating = True
    
    MsgBox "Proposta alterada com sucesso!", vbInformation
    
    ' Atualizar a lista de propostas do cliente
    CarregarPropostasDoCliente Me.txtID.Value
    
    ' Resetar o formulário para um estado inicial
    LimparFormulario
End Sub



Private Sub LimparFormulario()
    ' Limpar todos os campos e listas relevantes
    Me.txtNrProposta.Value = ""
    Me.txtReferencia.Value = ""
    Me.cmbVendedor.Value = ""
    Me.txtPrazoEntrega.Value = ""
    Me.cmbCondPagamento.Value = ""
    Me.txtValorTotal.Value = ""
    Me.lstProdutosDaProposta.Clear
    
    ' Reabilitar o botão Salvar Nova Proposta e desabilitar o botão Alterar Proposta
    Me.btnSalvarProposta.Enabled = True
    Me.btnAlterarProposta.Enabled = False
    
End Sub



'_______________________

'Analise o codigo acima