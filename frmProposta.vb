Private Sub UserForm_Initialize()
    ' Configurando a ListBox para ter 5 colunas
    With Me.lstCliente
        .ColumnCount = 5
        ' Definindo as larguras das colunas: ID, Nome, Contato, Cidade, Estado
        .ColumnWidths = "48;140;140;108;24"
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

    ' Foco no nome
        txtNomeCliente.SetFocus
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

    ' Desabilitar o botão Limpar e Buscar Cliente
    Me.btnLimparCliente.Enabled = False
    Me.btnBuscaCliente.Enabled = False

    ' Gerar novo número de proposta e registrar na planilha
    CriarNovaProposta
End Sub

Private Sub CriarNovaProposta()
    Dim wsPropostas As Worksheet
    Dim numeroBase As Long
    Dim novoNumero As String
    Dim estadoCliente As String
    Dim ultimaLinha As Long

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
    
    ' Encontrar a próxima linha vazia para registrar a nova proposta
    ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, 1).End(xlUp).Row + 1
    
    ' Preencher a nova linha na planilha de propostas
    wsPropostas.Cells(ultimaLinha, 1).Value = novoNumero ' Coluna NUMERO
    wsPropostas.Cells(ultimaLinha, 2).Value = Me.txtID.Value ' Coluna CLIENTE
    
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
            Me.txtItem.Value = 1 ' Iniciar item com 1 se for o primeiro produto
            
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
    Dim wsPropostas As Worksheet
    Dim ultimaLinha As Long
    Dim numeroProposta As String
    Dim cliente As String
    Dim item As Long
    Dim codigo As String
    Dim descricao As String
    Dim precoUnitario As Double
    Dim quantidade As Long
    Dim subtotal As Double
    Dim linhaProposta As Long

    ' Definindo a planilha de propostas
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    
    ' Obtendo o número da proposta e cliente selecionado
    numeroProposta = Me.txtNrProposta.Value
    cliente = Me.txtID.Value
    
    ' Obtendo os valores dos campos
    item = CLng(Me.txtItem.Value)
    codigo = Me.txtCodProduto.Value
    descricao = Me.txtDescricao.Value
    precoUnitario = CDbl(Me.txtPreco.Value)
    quantidade = CLng(Me.txtQTD.Value)
    subtotal = precoUnitario * quantidade
    
    ' Encontrar a linha da proposta atual
    linhaProposta = wsPropostas.Columns(1).Find(What:=numeroProposta, LookIn:=xlValues, LookAt:=xlWhole).Row
    
    ' Verificar se a linha da proposta já tem um item
    If wsPropostas.Cells(linhaProposta, 3).Value = "" Then
        ' Preencher a linha existente na planilha de propostas
        wsPropostas.Cells(linhaProposta, 3).Value = item ' Coluna ITEM
        wsPropostas.Cells(linhaProposta, 4).Value = codigo ' Coluna CODIGO
        wsPropostas.Cells(linhaProposta, 5).Value = precoUnitario ' Coluna PRECO UNITARIO
        wsPropostas.Cells(linhaProposta, 6).Value = quantidade ' Coluna QUANTIDADE
        wsPropostas.Cells(linhaProposta, 7).Value = subtotal ' Coluna SUBTOTAL
    Else
        ' Encontrar a próxima linha vazia para registrar o novo item da proposta
        ultimaLinha = wsPropostas.Cells(wsPropostas.Rows.Count, 1).End(xlUp).Row + 1
        
        ' Preencher a nova linha na planilha de propostas
        wsPropostas.Cells(ultimaLinha, 1).Value = numeroProposta ' Coluna NUMERO
        wsPropostas.Cells(ultimaLinha, 2).Value = cliente ' Coluna CLIENTE
        wsPropostas.Cells(ultimaLinha, 3).Value = item ' Coluna ITEM
        wsPropostas.Cells(ultimaLinha, 4).Value = codigo ' Coluna CODIGO
        wsPropostas.Cells(ultimaLinha, 5).Value = precoUnitario ' Coluna PRECO UNITARIO
        wsPropostas.Cells(ultimaLinha, 6).Value = quantidade ' Coluna QUANTIDADE
        wsPropostas.Cells(ultimaLinha, 7).Value = subtotal ' Coluna SUBTOTAL
    End If
    
    ' Limpar os campos de entrada
    Me.txtCodProduto.Value = ""
    Me.txtDescricao.Value = ""
    Me.txtPreco.Value = ""
    Me.txtQTD.Value = ""
    Me.txtItem.Value = ""
    
    ' Reposicionar o cursor para o campo txtCodProduto
    Me.txtCodProduto.SetFocus
    
    ' Incrementar o número do item para o próximo produto
    Me.txtItem.Value = item + 1
End Sub
