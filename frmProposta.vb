Option Explicit ' É uma boa prática sempre incluir isso

Private propostaExistenteCarregada  As Boolean


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
        ' Limpar informações da proposta anterior
        LimparInformacoesProposta
        
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

    VerificarSalvarProposta
End Sub

Private Sub AplicarFormatoMoeda(ByRef celula As Range)

    On Error Resume Next
    ' Tenta limpar a formatação, se falhar (devido a células mescladas), continua
    celula.ClearFormats
    On Error GoTo 0 ' Restaura o tratamento de erro normal
    With celula
        .NumberFormat = "General" ' Reset para formato geral
        .NumberFormat = """R$"" #,##0.00" ' Aplica novo formato
    End With
End Sub



Private Sub LimparInformacoesProposta()
    ' Limpar campos relacionados à proposta
    Me.txtNrProposta.Value = ""
    Me.txtReferencia.Value = ""
    Me.cmbVendedor.Value = ""
    Me.txtPrazoEntrega.Value = ""
    Me.cmbCondPagamento.Value = ""
    Me.txtValorTotal.Value = "0.00"
    
    ' Limpar a lista de produtos da proposta
    Me.lstProdutosDaProposta.Clear
    With Me.lstProdutosDaProposta
        .AddItem "Item"
        .List(0, 1) = "Código"
        .List(0, 2) = "Descrição"
        .List(0, 3) = "Qtd"
        .List(0, 4) = "$ Unitário"
        .List(0, 5) = "$ Total"
    End With
    
    ' Resetar variáveis de controle
    propostaExistenteCarregada = False
    Me.btnSalvarProposta.Enabled = False
    Me.btnAlterarProposta.Enabled = False
    
    ' Limpar campos de produto
    LimparCamposProduto
    Me.txtItem.Value = "1"
End Sub



Private Sub CarregarPropostasDoCliente(clienteID As String)
    Dim wsPropostas As Worksheet
    Dim rngPropostas As Range
    Dim cel As Range
    Dim ultimaLinha As Long
    Dim propostasOrdenadas As New Collection
    Dim valorTotal As Double
    Dim numeroProposta As String
    Dim i As Long
    
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
    
    ' Iterando sobre cada proposta e armazenando em uma coleção ordenada
    For Each cel In rngPropostas.Columns(2).Cells ' Coluna B para CLIENTE
        If cel.Value = clienteID Then
            numeroProposta = cel.Offset(0, -1).Value ' Coluna A para NUMERO
            
            ' Verificar se esta proposta já foi adicionada
            If Not ExisteNaColecao(propostasOrdenadas, numeroProposta) Then
                ' Calcular o valor total da proposta
                valorTotal = CalcularValorTotalProposta(numeroProposta)
                
                ' Criar um dicionário para armazenar os detalhes da proposta
                Dim proposta As Object
                Set proposta = CreateObject("Scripting.Dictionary")
                proposta("Numero") = numeroProposta
                proposta("Referencia") = cel.Offset(0, 6).Value  ' Coluna H para REFERENCIA
                proposta("Vendedor") = cel.Offset(0, 7).Value  ' Coluna I para VENDEDOR
                proposta("ValorTotal") = valorTotal
                
                ' Adicionar a proposta à coleção ordenada
                AdicionarOrdenado propostasOrdenadas, proposta
            End If
        End If
    Next cel
    
    ' Adicionar as propostas ordenadas ao ListBox
    For i = 1 To propostasOrdenadas.Count
        Dim prop As Object
        Set prop = propostasOrdenadas(i)
        Me.lstPropostasDoCliente.AddItem prop("Numero")
        Me.lstPropostasDoCliente.List(Me.lstPropostasDoCliente.ListCount - 1, 1) = prop("Referencia")
        Me.lstPropostasDoCliente.List(Me.lstPropostasDoCliente.ListCount - 1, 2) = prop("Vendedor")
        Me.lstPropostasDoCliente.List(Me.lstPropostasDoCliente.ListCount - 1, 3) = Format(prop("ValorTotal"), "#,##0.00")
    Next i
End Sub

Private Function ExisteNaColecao(col As Collection, chave As String) As Boolean
    Dim Item As Variant
    For Each Item In col
        If Item("Numero") = chave Then
            ExisteNaColecao = True
            Exit Function
        End If
    Next Item
    ExisteNaColecao = False
End Function

Private Sub AdicionarOrdenado(col As Collection, novaProposta As Object)
    Dim i As Long
    Dim inserido As Boolean
    
    inserido = False
    
    For i = 1 To col.Count
        If col(i)("Numero") > novaProposta("Numero") Then
            col.Add novaProposta, , i
            inserido = True
            Exit For
        End If
    Next i
    
    If Not inserido Then
        col.Add novaProposta
    End If
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
    propostaExistenteCarregada = False
    Me.btnSalvarProposta.Enabled = False
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

    ' Verificar se é uma atualização ou adição de novo item
    If btnAdicionarProduto.Caption = "ATUALIZAR PRODUTO" Then
        ' Atualizar o item existente
        AtualizarItemExistente
    Else
        ' Adicionar novo item
        AdicionarNovoItem
        
        ' Incrementar o número do item apenas para novos produtos
        If IsNumeric(Me.txtItem.Value) Then
            Me.txtItem.Value = CLng(Me.txtItem.Value) + 1
        Else
            ' Se o campo estiver vazio ou não for numérico, começar do 1
            Me.txtItem.Value = 1
        End If
    End If

    ' Limpar os campos de entrada
    LimparCamposProduto
    
    ' Resetar o botão e habilitar o campo de código do produto
    btnAdicionarProduto.Caption = "ADICIONAR PRODUTO"
    txtCodProduto.Enabled = True
    
    ' Esconder o botão de cancelar edição, se existir
    If Not btnCancelarEdicao Is Nothing Then
        btnCancelarEdicao.Visible = True
    End If
    
    ' Reposicionar o cursor para o campo txtCodProduto
    Me.txtCodProduto.SetFocus

    ' Verificar se pode habilitar o botão Salvar Proposta
    VerificarSalvarProposta

    ' Atualizar o valor total
    AtualizarValorTotal
End Sub


Private Sub btnRemoverProduto_Click()
    If Me.lstProdutosDaProposta.ListIndex > 0 Then ' Não remover o cabeçalho
        Me.lstProdutosDaProposta.RemoveItem Me.lstProdutosDaProposta.ListIndex
        AtualizarValorTotal
        
        ' Verificar se deve habilitar/desabilitar botões
        If propostaExistenteCarregada Then
            Me.btnAlterarProposta.Enabled = True
            Me.btnSalvarProposta.Enabled = False
        Else
            VerificarSalvarProposta
            Me.btnAlterarProposta.Enabled = False
        End If
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


Private Sub AtualizarValorTotal()
    Dim i As Long
    Dim valorTotal As Double
    Dim valorItem As String
    
    valorTotal = 0
    
    ' Iterar sobre todos os itens na lstProdutosDaProposta, exceto o cabeçalho
    For i = 1 To Me.lstProdutosDaProposta.ListCount - 1
        valorItem = Me.lstProdutosDaProposta.List(i, 5)
        If Len(Trim(valorItem)) > 0 Then
            valorTotal = valorTotal + ConvertToNumber(valorItem)
        End If
    Next i
    
    ' Atualizar o txtValorTotal com o valor calculado
    Me.txtValorTotal.Value = FormatarNumero(valorTotal)
End Sub



Private Sub lstPropostasDoCliente_Click()
    If Me.lstPropostasDoCliente.ListIndex > 0 Then ' Ignorar o cabeçalho
        Dim numeroProposta As String
        numeroProposta = Me.lstPropostasDoCliente.List(Me.lstPropostasDoCliente.ListIndex, 0)
        
        CarregarDetalhesPropostaExistente numeroProposta
        
        ' Desabilitar o botão Salvar Nova Proposta e habilitar o botão Alterar Proposta
        Me.btnSalvarProposta.Enabled = False
        Me.btnAlterarProposta.Enabled = True
        
        ' Marcar que uma proposta existente foi carregada
        propostaExistenteCarregada = True
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
    propostaExistenteCarregada = True
    

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



Private Sub btnSalvarProposta_Click()
    Dim wsPropostas As Worksheet
    Dim numeroProposta As String
    Dim novaReferencia As String
    Dim ultimaLinha As Long
    Dim i As Long
    Dim vendedor As String
    Dim condicaoPagamento As String
    Dim prazoEntrega As String
    Dim itensOrdenados As Collection
    Dim Item As Variant
    
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
    
    ' Criar uma coleção para armazenar os itens ordenados
    Set itensOrdenados = New Collection
    
    ' Adicionar itens à coleção
    For i = 1 To Me.lstProdutosDaProposta.ListCount - 1 ' Começando de 1 para pular o cabeçalho
        Set Item = CreateObject("Scripting.Dictionary")
        Item("Item") = CLng(Me.lstProdutosDaProposta.List(i, 0))
        Item("Codigo") = Me.lstProdutosDaProposta.List(i, 1)
        Item("PrecoUnitario") = CDbl(Me.lstProdutosDaProposta.List(i, 4))
        Item("Quantidade") = CLng(Me.lstProdutosDaProposta.List(i, 3))
        Item("Subtotal") = CDbl(Me.lstProdutosDaProposta.List(i, 5))
        
        ' Adicionar o item na posição correta
        If itensOrdenados.Count = 0 Then
            itensOrdenados.Add Item
        Else
            Dim j As Long
            For j = 1 To itensOrdenados.Count
                If Item("Item") < itensOrdenados(j)("Item") Then
                    itensOrdenados.Add Item, , j
                    Exit For
                ElseIf j = itensOrdenados.Count Then
                    itensOrdenados.Add Item
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' Iterar sobre os itens ordenados e adicionar à planilha
    For Each Item In itensOrdenados
        wsPropostas.Cells(ultimaLinha, 1).Value = numeroProposta ' Coluna NUMERO
        wsPropostas.Cells(ultimaLinha, 2).Value = Me.txtID.Value ' Coluna CLIENTE
        wsPropostas.Cells(ultimaLinha, 3).Value = Item("Item") ' Coluna ITEM
        wsPropostas.Cells(ultimaLinha, 4).Value = Item("Codigo") ' Coluna CODIGO
        wsPropostas.Cells(ultimaLinha, 5).Value = Item("PrecoUnitario") ' Coluna PRECO UNITARIO
        wsPropostas.Cells(ultimaLinha, 6).Value = Item("Quantidade") ' Coluna QUANTIDADE
        wsPropostas.Cells(ultimaLinha, 7).Value = Item("Subtotal") ' Coluna SUBTOTAL
        wsPropostas.Cells(ultimaLinha, 8).Value = novaReferencia ' Coluna REFERENCIA
        wsPropostas.Cells(ultimaLinha, 9).Value = vendedor ' Coluna VENDEDOR
        wsPropostas.Cells(ultimaLinha, 10).Value = condicaoPagamento ' Coluna CONDICAO DE PAGAMENTO
        wsPropostas.Cells(ultimaLinha, 11).Value = prazoEntrega ' Coluna PRAZO DE ENTREGA
        ultimaLinha = ultimaLinha + 1
    Next Item
    
    MsgBox "Proposta salva com sucesso!", vbInformation

    Unload Me
End Sub



Private Sub btnAlterarProposta_Click()
    Dim wsPropostas As Worksheet
    Dim rngPropostas As Range
    Dim cel As Range
    Dim numeroProposta As String
    Dim ultimaLinha As Long
    Dim i As Long
    Dim itensOrdenados As Collection
    Dim Item As Variant
    
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
    
    ' Criar uma coleção para armazenar os itens ordenados
    Set itensOrdenados = New Collection
    
    ' Adicionar itens à coleção
    For i = 1 To Me.lstProdutosDaProposta.ListCount - 1 ' Começando de 1 para pular o cabeçalho
        Set Item = CreateObject("Scripting.Dictionary")
        Item("Item") = CLng(Me.lstProdutosDaProposta.List(i, 0))
        Item("Codigo") = Me.lstProdutosDaProposta.List(i, 1)
        Item("PrecoUnitario") = CDbl(Me.lstProdutosDaProposta.List(i, 4))
        Item("Quantidade") = CLng(Me.lstProdutosDaProposta.List(i, 3))
        Item("Subtotal") = CDbl(Me.lstProdutosDaProposta.List(i, 5))
        
        ' Adicionar o item na posição correta
        If itensOrdenados.Count = 0 Then
            itensOrdenados.Add Item
        Else
            Dim j As Long
            For j = 1 To itensOrdenados.Count
                If Item("Item") < itensOrdenados(j)("Item") Then
                    itensOrdenados.Add Item, , j
                    Exit For
                ElseIf j = itensOrdenados.Count Then
                    itensOrdenados.Add Item
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' Adicionar novos registros da proposta atualizada no final da planilha
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).Row + 1
    
    For Each Item In itensOrdenados
        wsPropostas.Cells(ultimaLinha, 1).Value = numeroProposta
        wsPropostas.Cells(ultimaLinha, 2).Value = Me.txtID.Value
        wsPropostas.Cells(ultimaLinha, 3).Value = Item("Item")
        wsPropostas.Cells(ultimaLinha, 4).Value = Item("Codigo")
        wsPropostas.Cells(ultimaLinha, 5).Value = Format(Item("PrecoUnitario"), "#,##0.00")
        wsPropostas.Cells(ultimaLinha, 6).Value = Item("Quantidade")
        wsPropostas.Cells(ultimaLinha, 7).Value = Format(Item("Subtotal"), "#,##0.00")
        wsPropostas.Cells(ultimaLinha, 8).Value = Me.txtReferencia.Value
        wsPropostas.Cells(ultimaLinha, 9).Value = Me.cmbVendedor.Value
        wsPropostas.Cells(ultimaLinha, 10).Value = Me.cmbCondPagamento.Value
        wsPropostas.Cells(ultimaLinha, 11).Value = Me.txtPrazoEntrega.Value
        ultimaLinha = ultimaLinha + 1
    Next Item
    
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

   
    Unload Me
    
End Sub




Private Sub lstProdutosDaProposta_Click()
    If lstProdutosDaProposta.ListIndex > 0 Then ' Ignorar o cabeçalho
        txtItem.Value = lstProdutosDaProposta.List(lstProdutosDaProposta.ListIndex, 0)
        txtCodProduto.Value = lstProdutosDaProposta.List(lstProdutosDaProposta.ListIndex, 1)
        txtDescricao.Value = lstProdutosDaProposta.List(lstProdutosDaProposta.ListIndex, 2)
        txtQTD.Value = lstProdutosDaProposta.List(lstProdutosDaProposta.ListIndex, 3)
        txtPreco.Value = lstProdutosDaProposta.List(lstProdutosDaProposta.ListIndex, 4)
        
        ' Desabilitar a edição do código do produto
        txtCodProduto.Enabled = False
        
        ' Mudar o texto do botão para indicar que está editando
        btnAdicionarProduto.Caption = "ATUALIZAR PRODUTO"
    End If
End Sub

Private Sub AtualizarItemExistente()
    Dim index As Integer
    index = lstProdutosDaProposta.ListIndex
    
    If index > 0 Then ' Ignorar o cabeçalho
        Dim precoUnitario As Double
        Dim quantidade As Long
        Dim subtotal As Double
        
        precoUnitario = CDbl(Replace(Replace(txtPreco.Value, ".", ","), ",", "."))
        quantidade = CLng(txtQTD.Value)
        subtotal = precoUnitario * quantidade
        
        lstProdutosDaProposta.List(index, 0) = txtItem.Value
        lstProdutosDaProposta.List(index, 2) = txtDescricao.Value
        lstProdutosDaProposta.List(index, 3) = quantidade
        lstProdutosDaProposta.List(index, 4) = FormatarNumero(precoUnitario)
        lstProdutosDaProposta.List(index, 5) = FormatarNumero(subtotal)
    End If
End Sub


Private Sub AdicionarNovoItem()
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
    precoUnitario = ConvertToNumber(Me.txtPreco.Value) ' Converte considerando separador decimal
    quantidade = CLng(Me.txtQTD.Value)
    subtotal = precoUnitario * quantidade
    
    ' Adicionar o item ao ListBox
    Me.lstProdutosDaProposta.AddItem Item
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 1) = codigo
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 2) = descricao
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 3) = quantidade
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 4) = FormatarNumero(precoUnitario)
    Me.lstProdutosDaProposta.List(Me.lstProdutosDaProposta.ListCount - 1, 5) = FormatarNumero(subtotal)
End Sub



Private Sub LimparCamposProduto()
    Me.txtCodProduto.Value = ""
    Me.txtDescricao.Value = ""
    Me.txtPreco.Value = ""
    Me.txtQTD.Value = ""
    ' Não limpar o campo txtItem aqui, pois ele é incrementado para novos itens
End Sub



Private Sub btnCancelarEdicao_Click()
    LimparCamposProduto
    btnAdicionarProduto.Caption = "ADICIONAR PRODUTO"
    txtCodProduto.Enabled = True
    ' btnCancelarEdicao.Visible = False
End Sub








Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function



Private Function FormatarNumero(ByVal valor As Double, Optional casasDecimais As Integer = 2) As String
    Dim formatString As String
    Dim decimalSep As String
    Dim thousandsSep As String
    
    decimalSep = GetDecimalSeparator()
    thousandsSep = GetThousandsSeparator()
    
    formatString = "#" & thousandsSep & "##0"
    If casasDecimais > 0 Then
        formatString = formatString & decimalSep & String(casasDecimais, "0")
    End If
    
    FormatarNumero = Format(valor, formatString)
End Function



Private Function ConvertToNumber(ByVal strValue As String) As Double
    On Error GoTo TratarErro
    
    Dim cleanValue As String
    
    ' Remove espaços em branco
    cleanValue = Trim(strValue)
    
    ' Remove qualquer caractere que não seja número, vírgula ou ponto
    Dim i As Long
    Dim char As String
    Dim newValue As String
    
    For i = 1 To Len(cleanValue)
        char = Mid(cleanValue, i, 1)
        If char Like "[0-9]" Or char = "," Or char = "." Then
            newValue = newValue & char
        End If
    Next i
    
    ' Se não houver valor numérico, retorna 0
    If Len(newValue) = 0 Then
        ConvertToNumber = 0
        Exit Function
    End If
    
    ' Determina qual é o separador decimal baseado no último separador encontrado
    Dim ultimoPonto As Long
    Dim ultimaVirgula As Long
    
    ultimoPonto = InStrRev(newValue, ".")
    ultimaVirgula = InStrRev(newValue, ",")
    
    ' Remove todos os separadores exceto o último
    If ultimoPonto > ultimaVirgula Then
        ' Último separador é ponto, então usa ponto como decimal
        newValue = Replace(newValue, ",", "")
        ConvertToNumber = Val(newValue)
    Else
        ' Último separador é vírgula ou não há separadores
        newValue = Replace(newValue, ".", "")
        newValue = Replace(newValue, ",", ".")
        ConvertToNumber = Val(newValue)
    End If
    
    Exit Function

    TratarErro:
        ConvertToNumber = 0
End Function




Private Sub PreencherItensProposta(wsNovaProposta As Worksheet, wsPropostas As Worksheet, wsPrecos As Worksheet, numeroProposta As String)
    Dim rngPropostas As Range
    Dim rngProposta As Range
    Dim ultimaLinha As Long
    Dim i As Long
    Dim countItens As Long
    
    ' Definir o intervalo de dados das propostas
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, "A").End(xlUp).Row
    Set rngPropostas = wsPropostas.Range("A2:K" & ultimaLinha)
    
    ' Contar o número de itens na proposta
    countItens = Application.WorksheetFunction.CountIf(wsPropostas.Range("A2:A" & ultimaLinha), numeroProposta)
    
    ' Replicar a formatação da linha 15 para as linhas subsequentes
    If countItens > 1 Then
        wsNovaProposta.Rows("15:15").Copy
        wsNovaProposta.Rows("16:" & 15 + countItens - 1).Insert Shift:=xlDown
    End If
    
    i = 15 ' Linha inicial para os itens (após o cabeçalho)
    
    ' Iterar sobre cada linha na planilha de propostas
    For Each rngProposta In rngPropostas.Rows
        If rngProposta.Cells(1, 1).Value = numeroProposta Then
            wsNovaProposta.Cells(i, 1).Value = rngProposta.Cells(1, 3).Value ' Item
            wsNovaProposta.Cells(i, 2).Value = rngProposta.Cells(1, 6).Value ' Quantidade
            wsNovaProposta.Cells(i, 3).Value = rngProposta.Cells(1, 4).Value ' Código do Produto
            
            ' Buscar descrição e outras informações do produto
            Dim rngProduto As Range
            Set rngProduto = wsPrecos.Range("A:I").Find(What:=rngProposta.Cells(1, 4).Value, LookIn:=xlValues, LookAt:=xlWhole)
            
            If Not rngProduto Is Nothing Then
                wsNovaProposta.Cells(i, 4).Value = rngProduto.Offset(0, 1).Value & vbNewLine & _
                                    vbNewLine & _
                                    "NCM: " & rngProduto.Offset(0, 5).Value & vbNewLine & _
                                    "ANVISA: " & rngProduto.Offset(0, 3).Value & vbNewLine & _
                                    "SIMPRO: " & rngProduto.Offset(0, 7).Value
                
                ' Ajustar a altura da linha para acomodar o texto adicional
                wsNovaProposta.Rows(i).RowHeight = 75.00
            End If
            
            ' Configurar a célula para quebra de texto
            wsNovaProposta.Cells(i, 4).WrapText = True

            ' Formatação dos valores numéricos
            Dim precoUnitario As Double
            Dim subtotal As Double
            
            precoUnitario = CDbl(rngProposta.Cells(1, 5).Value)
            subtotal = CDbl(rngProposta.Cells(1, 7).Value)
            
            wsNovaProposta.Cells(i, 9).Value = precoUnitario
            AplicarFormatoMoeda wsNovaProposta.Cells(i, 9)
            
            wsNovaProposta.Cells(i, 11).Value = subtotal
            AplicarFormatoMoeda wsNovaProposta.Cells(i, 11)
            
            i = i + 1
        End If
    Next rngProposta
    
    ' Preencher informações finais
    Dim valorTotal As Double
    valorTotal = Application.Sum(wsNovaProposta.Range("K15:K" & i - 1))
    
    wsNovaProposta.Range("K" & i + 1).Value = valorTotal
    AplicarFormatoMoeda wsNovaProposta.Range("K" & (i + 1))
    
    wsNovaProposta.Range("J" & i + 1).Value = valorTotal
    AplicarFormatoMoeda wsNovaProposta.Range("J" & (i + 1))
    
    ' Preencher Condição de Pagamento e Prazo de Entrega
    wsNovaProposta.Range("E" & i + 2).Value = Me.cmbCondPagamento.Value
    wsNovaProposta.Range("E" & i + 3).Value = Me.txtPrazoEntrega.Value
    
    ' Preencher informações do vendedor
    Dim vendedorNome As String
    Dim vendedorCargo As String
    Dim vendedorEmail As String
    Dim vendedorFone As String
    Dim wsVendedores As Worksheet
    Dim rngVendedor As Range
    Dim linhaVendedor As Long
    
    vendedorNome = Me.cmbVendedor.Value
    Set wsVendedores = ThisWorkbook.Sheets("VENDEDORES")
    Set rngVendedor = wsVendedores.Range("A:D").Find(What:=vendedorNome, LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not rngVendedor Is Nothing Then
        vendedorEmail = rngVendedor.Offset(0, 1).Value
        vendedorFone = rngVendedor.Offset(0, 2).Value
        vendedorCargo = rngVendedor.Offset(0, 3).Value
    Else
        vendedorEmail = "Email não encontrado"
        vendedorFone = "Fone não encontrado"
        vendedorCargo = ""
    End If
    
    linhaVendedor = i + 6
    wsNovaProposta.Range("A" & linhaVendedor).Value = vendedorNome
    
    If vendedorCargo <> "" Then
        linhaVendedor = linhaVendedor + 1
        wsNovaProposta.Range("A" & linhaVendedor).Value = vendedorCargo
    End If
    
    linhaVendedor = linhaVendedor + 1
    wsNovaProposta.Range("A" & linhaVendedor).Value = vendedorEmail
    
    linhaVendedor = linhaVendedor + 1
    wsNovaProposta.Range("A" & linhaVendedor).Value = vendedorFone
End Sub




Private Sub btnImprimir_Click()
    Dim wsModelo As Worksheet
    Dim wsNovaProposta As Worksheet
    Dim wsPropostas As Worksheet
    Dim wsClientes As Worksheet
    Dim wsPrecos As Worksheet
    Dim numeroProposta As String
    Dim ultimaLinha As Long
    Dim i As Long, j As Long
    
    ' Definir as planilhas
    Set wsModelo = ThisWorkbook.Sheets("IMPRESSAO")
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    Set wsClientes = ThisWorkbook.Sheets("CLIENTES")
    Set wsPrecos = ThisWorkbook.Sheets("ListaDePrecos")
    
    ' Obter o número da proposta atual
    numeroProposta = Me.txtNrProposta.Value
    
    ' Criar nome da nova planilha
    Dim novoNomePlanilha As String
    novoNomePlanilha = "Proposta_" & numeroProposta
    
    ' Verificar se já existe uma planilha com esse nome e adicionar letra se necessário
    Dim letra As String
    letra = ""
    Do While SheetExists(novoNomePlanilha & IIf(letra = "", "", "-" & letra))
        If letra = "" Then
            letra = "A"
        Else
            letra = Chr(Asc(letra) + 1)
        End If
    Loop
    novoNomePlanilha = novoNomePlanilha & IIf(letra = "", "", "-" & letra)
    
    ' Criar nova planilha para a proposta
    Set wsNovaProposta = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsNovaProposta.Name = novoNomePlanilha
    
    ' Copiar o modelo para a nova planilha
    wsModelo.UsedRange.Copy wsNovaProposta.Range("A1")
    
    ' Preencher informações da proposta
    With wsNovaProposta
        ' Colocar a data como texto na célula A6
        .Range("L6").Value = Format(Date, "dd ""de"" mmmm ""de"" yyyy")
        .Range("L6").NumberFormat = "@"
        .Range("L6").HorizontalAlignment = xlRight
        .Range("C7").Value = numeroProposta
        
        ' Tratamento para a Referência
        Dim referenciaValor As String
        referenciaValor = Trim(Me.txtReferencia.Value)
        If referenciaValor <> "" Then
            .Range("F7").Value = "Referência:"
            .Range("G7").Value = referenciaValor
        Else
            .Range("F7").Value = ""
            .Range("G7").Value = ""
        End If
        
        ' Preencher informações do cliente
        Dim clienteID As String
        clienteID = Me.txtID.Value
        Dim rngCliente As Range
        Set rngCliente = wsClientes.Range("A:H").Find(What:=clienteID, LookIn:=xlValues, LookAt:=xlWhole)
        If Not rngCliente Is Nothing Then
            .Range("A9").Value = rngCliente.Offset(0, 1).Value ' Nome do cliente
            .Range("B10").Value = rngCliente.Offset(0, 2).Value ' Contato
            ' Tratamento para Endereço, Cidade e Estado
            Dim enderecoCliente As String
            Dim cidadeCliente As String
            Dim estadoCliente As String
            Dim cidadeEstado As String
            enderecoCliente = Trim(rngCliente.Offset(0, 3).Value)
            cidadeCliente = Trim(rngCliente.Offset(0, 4).Value)
            estadoCliente = Trim(rngCliente.Offset(0, 5).Value)
            ' Montar string de cidade/estado
            If cidadeCliente <> "" And estadoCliente <> "" Then
                cidadeEstado = cidadeCliente & " / " & estadoCliente
            ElseIf cidadeCliente <> "" Then
                cidadeEstado = cidadeCliente
            ElseIf estadoCliente <> "" Then
                cidadeEstado = estadoCliente
            Else
                cidadeEstado = ""
            End If
            ' Verificar se há pelo menos uma informação de endereço
            If enderecoCliente <> "" Or cidadeCliente <> "" Or estadoCliente <> "" Then
                .Range("A11").Value = "End.:"
                ' Preencher o endereço
                If enderecoCliente <> "" Then
                    .Range("B11").Value = enderecoCliente
                    .Range("B12").Value = cidadeEstado
                Else
                    .Range("B11").Value = cidadeEstado
                    .Range("B12").Value = ""
                End If
            Else
                ' Se não houver informações de endereço, limpar as células
                .Range("A11").Value = ""
                .Range("B11").Value = ""
                .Range("B12").Value = ""
            End If
            ' Tratamento específico para o telefone
            Dim telefoneCliente As String
            telefoneCliente = Trim(rngCliente.Offset(0, 6).Value)
            If telefoneCliente <> "" Then
                .Range("G10").Value = "Telefone:"
                .Range("H10").Value = "'" & telefoneCliente
                .Range("H10").NumberFormat = "@" ' Manter formato de texto
            Else
                .Range("G10").Value = ""
                .Range("H10").Value = ""
            End If
            ' Tratamento específico para o email
            Dim emailCliente As String
            emailCliente = Trim(rngCliente.Offset(0, 7).Value)
            If emailCliente <> "" Then
                .Range("G11").Value = "Email:"
                .Range("H11").Value = emailCliente
            Else
                .Range("G11").Value = ""
                .Range("H11").Value = ""
            End If
        End If
        
        ' Preencher itens da proposta
        PreencherItensProposta wsNovaProposta, wsPropostas, wsPrecos, numeroProposta
        
        ' Ajustar larguras das colunas conforme especificado
        .Columns("A:B").ColumnWidth = 4.82
        .Columns("C").ColumnWidth = 7.91
        .Columns("D:H").ColumnWidth = 9.36
        .Columns("I:L").ColumnWidth = 5.27
    End With
    
    ' Configurações básicas de página
    With wsNovaProposta.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(2)
        .RightMargin = Application.CentimetersToPoints(2)
        .TopMargin = Application.CentimetersToPoints(1)
        .BottomMargin = Application.CentimetersToPoints(1)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.5)
        .CenterHorizontally = True
        .CenterVertically = False
    End With
    
    MsgBox "Proposta criada com sucesso na planilha: " & wsNovaProposta.Name, vbInformation
    Unload Me
End Sub




Private Function GetDecimalSeparator() As String
    GetDecimalSeparator = Application.International(xlDecimalSeparator)
End Function



Private Function GetThousandsSeparator() As String
    GetThousandsSeparator = Application.International(xlThousandsSeparator)
End Function



Private Function FormatarMoeda(ByVal valor As Double) As String
    FormatarMoeda = FormatCurrency(valor, 2, vbTrue, vbTrue, vbTrue)
End Function


Private Function ValidarValorNumerico(ByVal strValue As String) As Boolean
    Dim decimalSep As String
    Dim thousandsSep As String
    
    decimalSep = GetDecimalSeparator()
    thousandsSep = GetThousandsSeparator()
    
    ' Remove espaços
    strValue = Trim(strValue)
    
    ' Verifica se está vazio
    If strValue = "" Then
        ValidarValorNumerico = False
        Exit Function
    End If
    
    ' Tenta converter para número
    On Error Resume Next
    ConvertToNumber strValue
    ValidarValorNumerico = (Err.Number = 0)
    On Error GoTo 0
End Function



' _______________________

' Analise o codigo acima pois preciso resolver um problema no arquivo criado para impressao