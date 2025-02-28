' Variáveis de módulo para controle do estado do formulário
Private modoEdicao As Boolean
Private propostaAlterada As Boolean
Private nrPropostaOriginal As String

Private Sub UserForm_Initialize()
    ' Desabilitar botões inicialmente
    btnBuscarProduto.Enabled = False
    btnAdicionarProduto.Enabled = False
    btnSalvarNovaProposta.Enabled = False
    btnAlterarProposta.Enabled = False
    
    ' Inicializa o ListBox
    With lstProdutosDaProposta
        .Clear
        .ColumnCount = 6
        .ColumnWidths = "40;60;320;90;90;120"
    End With
    
    ' Adiciona o cabeçalho
    lstProdutosDaProposta.AddItem ""
    lstProdutosDaProposta.List(0, 0) = "Item"
    lstProdutosDaProposta.List(0, 1) = "Qtd"
    lstProdutosDaProposta.List(0, 2) = "Descrição"
    lstProdutosDaProposta.List(0, 3) = "Código"
    lstProdutosDaProposta.List(0, 4) = "Preço"
    lstProdutosDaProposta.List(0, 5) = "Sub Total"

    ' Preencher ComboBoxes
    PreencherComboBoxes
    
    ' Inicializar em modo de nova proposta
    modoEdicao = False
    propostaAlterada = False
    nrPropostaOriginal = ""
    
    ' Valor inicial para o próximo item
    txtItem.Value = "1"
End Sub

Private Sub PreencherComboBoxes()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ListasDeEscolha")
    
    ' Preencher Vendedores usando a tabela nomeada "Vendedor"
    With cmbVendedor
        .Clear
        .List = ws.ListObjects("Vendedor").DataBodyRange.Value
    End With
    
    ' Preencher Condições de Pagamento usando a tabela nomeada "CondPagto"
    With cmbCondPagamento
        .Clear
        .List = ws.ListObjects("CondPagto").DataBodyRange.Value
    End With
    
    ' Preencher Tipos de Frete usando a tabela nomeada "Frete"
    With cmbFrete
        .Clear
        .List = ws.ListObjects("Frete").DataBodyRange.Value
    End With
End Sub

' Esta sub gerencia o estado dos botões Salvar Nova Proposta e Alterar Proposta
Private Sub CheckEnableSalvarProposta()
    Dim camposPreenchidos As Boolean
    Dim temItens As Boolean
    
    ' Verifica se todos os campos obrigatórios estão preenchidos
    camposPreenchidos = (Trim(txtNomeCliente.Value) <> "" And _
                         Trim(txtCidade.Value) <> "" And _
                         Trim(txtEstado.Value) <> "" And _
                         cmbVendedor.Value <> "" And _
                         cmbCondPagamento.Value <> "" And _
                         Trim(txtPrazoEntrega.Value) <> "" And _
                         cmbFrete.Value <> "")
    
    ' Verifica se há pelo menos um item na lista (além do cabeçalho)
    temItens = (lstProdutosDaProposta.ListCount > 1)
    
    ' Lógica para habilitar/desabilitar os botões baseada no modo
    If modoEdicao Then
        ' Modo de edição de proposta existente
        btnSalvarNovaProposta.Enabled = False ' Sempre desabilitado em modo edição
        
        ' Botão Alterar só fica habilitado se houver itens, campos preenchidos e alterações
        btnAlterarProposta.Enabled = camposPreenchidos And temItens And propostaAlterada
    Else
        ' Modo de nova proposta
        btnSalvarNovaProposta.Enabled = camposPreenchidos And temItens
        btnAlterarProposta.Enabled = False ' Sempre desabilitado em modo criação
    End If
End Sub

' Função para rastrear mudanças quando em modo de edição
Private Sub MarcarComoAlterado()
    If modoEdicao Then
        propostaAlterada = True
        CheckEnableSalvarProposta
    End If
End Sub

' Eventos para verificar a habilitação do botão Buscar Produto
Private Sub CheckEnableBuscarProduto()
    btnBuscarProduto.Enabled = (Trim(txtNomeCliente.Value) <> "" And _
                                Trim(txtCidade.Value) <> "" And _
                                Trim(txtEstado.Value) <> "")
End Sub

Private Sub txtNomeCliente_Change()
    MarcarComoAlterado
    CheckEnableBuscarProduto
    CheckEnableSalvarProposta
End Sub

Private Sub txtCidade_Change()
    MarcarComoAlterado
    CheckEnableBuscarProduto
    CheckEnableSalvarProposta
End Sub

Private Sub txtEstado_Change()
    MarcarComoAlterado
    CheckEnableBuscarProduto
    CheckEnableSalvarProposta
End Sub

' Eventos para verificar a habilitação do botão Adicionar Produto
Private Sub CheckEnableAdicionarProduto()
    btnAdicionarProduto.Enabled = (Trim(txtNomeCliente.Value) <> "" And _
                                   Trim(txtCidade.Value) <> "" And _
                                   Trim(txtEstado.Value) <> "" And _
                                   Trim(txtCodProduto.Value) <> "" And _
                                   Trim(txtDescricao.Value) <> "" And _
                                   Trim(txtPreco.Value) <> "" And _
                                   Trim(txtQTD.Value) <> "" And _
                                   Trim(txtItem.Value) <> "")
End Sub

Private Sub txtCodProduto_Change()
    CheckEnableAdicionarProduto
End Sub

Private Sub txtDescricao_Change()
    CheckEnableAdicionarProduto
End Sub

Private Sub txtPreco_Change()
    CheckEnableAdicionarProduto
End Sub

Private Sub txtQTD_Change()
    CheckEnableAdicionarProduto
End Sub

Private Sub txtItem_Change()
    CheckEnableAdicionarProduto
End Sub

' ======================
' ROTINAS PARA BUSCA DE CLIENTE
' ======================

Private Sub btnBuscaCliente_Click()
    ' Abre o formulário frmCliente
    frmCliente.Show
End Sub

' Esta sub será chamada pelo frmCliente quando um cliente for selecionado
Public Sub PreencherDadosCliente(nome As String, cidade As String, estado As String)
    txtNomeCliente.Value = nome
    txtCidade.Value = cidade
    txtEstado.Value = estado
End Sub

' ======================
' ROTINAS PARA BUSCA DE PRODUTO
' ======================

Private Sub btnBuscarProduto_Click()
    Dim ws As Worksheet
    Dim codigo As String
    Dim ultimaLinha As Long
    Dim i As Long
    Dim encontrado As Boolean
    Dim preco As Double
    Set ws = ThisWorkbook.Worksheets("ListaDePrecos")
    codigo = Trim(txtCodProduto.Value)
    If codigo = "" Then
        MsgBox "Por favor, digite um código de produto.", vbExclamation
        Exit Sub
    End If
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    encontrado = False
    For i = 2 To ultimaLinha
        If ws.Cells(i, "A").Value = codigo Then
            txtDescricao.Value = ws.Cells(i, "B").Value
            preco = CDbl(ws.Cells(i, "C").Value)
            txtPreco.Value = Format(preco, "#,##0.00")
            encontrado = True
            Exit For
        End If
    Next i
    If Not encontrado Then
        MsgBox "Produto não encontrado.", vbExclamation
        txtDescricao.Value = ""
        txtPreco.Value = ""
    End If
    CheckEnableAdicionarProduto
End Sub

' ======================
' ROTINAS PARA ADICIONAR ITENS NA PROPOSTA
' ======================

Private Sub btnAdicionarProduto_Click()
    Dim subTotal As Double
    Dim preco As Double
    Dim quantidade As Double
    
    ' Verifica se é o primeiro item
    If txtNovaProposta.Value = "" And Not modoEdicao Then
        GerarNumeroProposta
    End If
    
    ' Converte os valores de texto para números
    preco = ConverterParaNumero(txtPreco.Value)
    quantidade = CDbl(txtQTD.Value)
    
    ' Calcula o Sub Total
    subTotal = quantidade * preco
    
    ' Adiciona o item à lista
    With lstProdutosDaProposta
        .AddItem ""
        .List(.ListCount - 1, 0) = txtItem.Value
        .List(.ListCount - 1, 1) = txtQTD.Value
        .List(.ListCount - 1, 2) = txtDescricao.Value
        .List(.ListCount - 1, 3) = txtCodProduto.Value
        .List(.ListCount - 1, 4) = Format(preco, "#,##0.00")
        .List(.ListCount - 1, 5) = Format(subTotal, "#,##0.00") ' Sub Total
    End With
    
    ' Atualiza o valor total da proposta
    AtualizarValorTotal
    
    ' Limpa os campos do produto
    txtItem.Value = ""
    txtQTD.Value = ""
    txtCodProduto.Value = ""
    txtDescricao.Value = ""
    txtPreco.Value = ""
    
    ' Desabilita o botão Adicionar Produto
    btnAdicionarProduto.Enabled = False
    
    ' Incrementa automaticamente o número do item
    txtItem.Value = lstProdutosDaProposta.ListCount
    
    ' Se estiver em modo de edição, marcar como alterado
    If modoEdicao Then
        MarcarComoAlterado
    End If
    
    ' Verificar estado dos botões Salvar/Alterar
    CheckEnableSalvarProposta
End Sub

Private Sub GerarNumeroProposta()
    Dim ws As Worksheet
    Dim ultimoNumero As Long
    Dim novoNumero As String
    Dim iniciais As String
    
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    
    ' Obter o último número da proposta
    ultimoNumero = ws.Range("U1").Value
    
    ' Incrementar o número
    ultimoNumero = ultimoNumero + 1
    
    ' Atualizar o número na planilha
    ws.Range("U1").Value = ultimoNumero
    
    ' Gerar as iniciais do cliente
    iniciais = Left(txtNomeCliente.Value, 1) & Mid(txtNomeCliente.Value, InStr(txtNomeCliente.Value, " ") + 1, 1)
    
    ' Formatar o novo número da proposta
    novoNumero = Format(Date, "yyyy-mm-dd") & "_" & UCase(iniciais) & "_" & Format(ultimoNumero, "00000")
    
    ' Atualizar o campo txtNovaProposta
    txtNovaProposta.Value = novoNumero
End Sub

Private Sub btnRemoverProduto_Click()
    Dim index As Long
    
    ' Verificar se existe algum item selecionado
    index = lstProdutosDaProposta.ListIndex
    
    ' Verificar se não é o cabeçalho (índice 0) e se tem algum item selecionado
    If index > 0 Then
        ' Remover o item selecionado
        lstProdutosDaProposta.RemoveItem index
        
        ' Renumerar os itens restantes
        RenumerarItens
        
        ' Atualizar o valor total
        AtualizarValorTotal
        
        ' Se estiver em modo de edição, marcar como alterado
        If modoEdicao Then
            MarcarComoAlterado
        End If
        
        ' Verificar estado dos botões após remoção
        CheckEnableSalvarProposta
    Else
        MsgBox "Selecione um item da proposta para remover.", vbExclamation
    End If
End Sub

Private Sub RenumerarItens()
    Dim i As Long
    
    ' Renumerar todos os itens na listbox (começando do item 1, já que 0 é o cabeçalho)
    For i = 1 To lstProdutosDaProposta.ListCount - 1
        lstProdutosDaProposta.List(i, 0) = i
    Next i
    
    ' Atualizar o próximo número de item para adicionar
    If lstProdutosDaProposta.ListCount > 1 Then
        txtItem.Value = lstProdutosDaProposta.ListCount
    Else
        txtItem.Value = "1" ' Começar do 1 se não houver itens
    End If
End Sub

' ======================
' ROTINAS PARA BUSCA DE PROPOSTA
' ======================

Private Sub btnBuscaProposta_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim criterio As String
    Dim propostasEncontradas As Collection
    
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    criterio = LCase(txtNrProposta.Value)
    
    Set propostasEncontradas = New Collection
    
    ' Limpar o ListBox
    lstBuscaProposta.Clear
    
    ' Adicionar cabeçalho
    lstBuscaProposta.AddItem
    lstBuscaProposta.List(0, 0) = "Nr da Proposta"
    lstBuscaProposta.List(0, 1) = "Nome do Cliente"
    
    ' Buscar propostas únicas
    For i = 2 To ultimaLinha
        If InStr(1, LCase(ws.Cells(i, "A").Value), criterio) > 0 Then
            On Error Resume Next
            propostasEncontradas.Add ws.Cells(i, "A").Value, CStr(ws.Cells(i, "A").Value)
            On Error GoTo 0
        End If
    Next i
    
    ' Adicionar propostas únicas ao ListBox
    For i = 1 To propostasEncontradas.Count
        lstBuscaProposta.AddItem
        lstBuscaProposta.List(lstBuscaProposta.ListCount - 1, 0) = propostasEncontradas(i)
        ' Encontrar o nome do cliente correspondente
        Dim clienteLinha As Long
        On Error Resume Next
        clienteLinha = Application.Match(propostasEncontradas(i), ws.Range("A:A"), 0)
        On Error GoTo 0
        If clienteLinha > 0 Then
            lstBuscaProposta.List(lstBuscaProposta.ListCount - 1, 1) = ws.Cells(clienteLinha, "B").Value
        End If
    Next i
    
    If lstBuscaProposta.ListCount = 1 Then ' Só tem o cabeçalho
        MsgBox "Nenhuma proposta encontrada.", vbInformation
    End If
End Sub

Private Sub lstBuscaProposta_Click()
    Dim ws As Worksheet
    Dim linha As Long
    Dim nrProposta As String
    
    ' Verificação completa para evitar erros de índice
    If lstBuscaProposta.ListCount <= 1 Then Exit Sub    ' Lista vazia ou só com cabeçalho
    If lstBuscaProposta.ListIndex < 0 Then Exit Sub     ' Nenhum item selecionado
    If lstBuscaProposta.ListIndex = 0 Then Exit Sub     ' Cabeçalho selecionado
    
    On Error Resume Next
    nrProposta = lstBuscaProposta.List(lstBuscaProposta.ListIndex, 0)
    If Err.Number <> 0 Then
        MsgBox "Erro ao selecionar a proposta. Por favor, tente novamente.", vbExclamation
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    If Trim(nrProposta) = "" Then
        MsgBox "Proposta inválida selecionada.", vbExclamation
        Exit Sub
    End If
    
    ' Limpar formulário antes de carregar nova proposta
    LimparFormulario
    
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    
    ' Armazenar o número da proposta original
    nrPropostaOriginal = nrProposta
    
    ' Encontrar a linha da proposta
    On Error Resume Next
    linha = Application.Match(nrProposta, ws.Range("A:A"), 0)
    If Err.Number <> 0 Then
        MsgBox "Erro ao localizar a proposta no banco de dados.", vbExclamation
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    
    If linha <= 0 Then
        MsgBox "Proposta não encontrada.", vbExclamation
        Exit Sub
    End If
    
    ' Preencher os campos
    txtNrProposta.Value = nrProposta
    txtNomeCliente.Value = ws.Cells(linha, "B").Value
    txtCidade.Value = ws.Cells(linha, "C").Value
    txtEstado.Value = ws.Cells(linha, "D").Value
    txtPessoaContato.Value = ws.Cells(linha, "E").Value
    txtFone.Value = ws.Cells(linha, "G").Value
    txtEmail.Value = ws.Cells(linha, "F").Value
    txtRefProposta.Value = ws.Cells(linha, "L").Value
    cmbVendedor.Value = ws.Cells(linha, "M").Value
    cmbCondPagamento.Value = ws.Cells(linha, "N").Value
    txtPrazoEntrega.Value = ws.Cells(linha, "O").Value
    cmbFrete.Value = ws.Cells(linha, "P").Value
    
    ' Definir como modo de edição
    modoEdicao = True
    propostaAlterada = False ' Inicialmente não alterada
    
    ' Preencher o ListBox com os itens da proposta
    PreencherListBoxItens nrProposta
    
    ' Calcular e preencher o valor total
    AtualizarValorTotal
    
    ' Atualizar estado dos botões
    CheckEnableSalvarProposta
End Sub

Private Sub PreencherListBoxItens(nrProposta As String)
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim wsPrecos As Worksheet
    Dim descricao As String
    
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    Set wsPrecos = ThisWorkbook.Worksheets("ListaDePrecos")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Limpar o ListBox de itens
    lstProdutosDaProposta.Clear
    
    ' Adicionar cabeçalho
    With lstProdutosDaProposta
        .AddItem
        .List(0, 0) = "Item"
        .List(0, 1) = "Qtd"
        .List(0, 2) = "Descrição"
        .List(0, 3) = "Código"
        .List(0, 4) = "Preço"
        .List(0, 5) = "Sub Total"
    End With
    
    ' Preencher com os itens da proposta
    For i = 2 To ultimaLinha
        If ws.Cells(i, "A").Value = nrProposta Then
            ' Buscar descrição na planilha ListaDePrecos
            On Error Resume Next
            descricao = Application.WorksheetFunction.VLookup(ws.Cells(i, "I").Value, wsPrecos.Range("A:B"), 2, False)
            If Err.Number <> 0 Then
                descricao = "Descrição não encontrada"
                Err.Clear
            End If
            On Error GoTo 0
            
            With lstProdutosDaProposta
                .AddItem
                .List(.ListCount - 1, 0) = ws.Cells(i, "H").Value ' Item
                .List(.ListCount - 1, 1) = ws.Cells(i, "K").Value ' Quantidade
                .List(.ListCount - 1, 2) = descricao ' Descrição
                .List(.ListCount - 1, 3) = ws.Cells(i, "I").Value ' Código
                .List(.ListCount - 1, 4) = Format(ws.Cells(i, "J").Value, "#,##0.00") ' Preço
                .List(.ListCount - 1, 5) = Format(CDbl(ws.Cells(i, "J").Value) * CDbl(ws.Cells(i, "K").Value), "#,##0.00") ' Sub Total
            End With
        End If
    Next i
End Sub

' ======================
' ROTINAS PARA SALVAR E ALTERAR PROPOSTA
' ======================

Private Sub btnSalvarNovaProposta_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    ' Verificações de segurança
    If modoEdicao Then
        MsgBox "Esta operação não é válida em modo de edição.", vbExclamation
        Exit Sub
    End If
    
    If lstProdutosDaProposta.ListCount <= 1 Then
        MsgBox "Adicione pelo menos um item à proposta antes de salvar.", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Salvar cada item da proposta
    For i = 1 To lstProdutosDaProposta.ListCount - 1
        ws.Cells(ultimaLinha, "A").Value = txtNovaProposta.Value
        ws.Cells(ultimaLinha, "B").Value = txtNomeCliente.Value
        ws.Cells(ultimaLinha, "C").Value = txtCidade.Value
        ws.Cells(ultimaLinha, "D").Value = txtEstado.Value
        ws.Cells(ultimaLinha, "E").Value = txtPessoaContato.Value
        ws.Cells(ultimaLinha, "F").Value = txtEmail.Value
        ws.Cells(ultimaLinha, "G").Value = txtFone.Value
        ws.Cells(ultimaLinha, "H").Value = lstProdutosDaProposta.List(i, 0) ' Item
        ws.Cells(ultimaLinha, "I").Value = lstProdutosDaProposta.List(i, 3) ' Código
        ws.Cells(ultimaLinha, "J").Value = ConverterParaNumero(lstProdutosDaProposta.List(i, 4)) ' Preço
        ws.Cells(ultimaLinha, "K").Value = lstProdutosDaProposta.List(i, 1) ' Quantidade
        ws.Cells(ultimaLinha, "L").Value = txtRefProposta.Value ' Referência da Proposta
        ws.Cells(ultimaLinha, "M").Value = cmbVendedor.Value
        ws.Cells(ultimaLinha, "N").Value = cmbCondPagamento.Value
        ws.Cells(ultimaLinha, "O").Value = txtPrazoEntrega.Value
        ws.Cells(ultimaLinha, "P").Value = cmbFrete.Value
        
        ultimaLinha = ultimaLinha + 1
    Next i
    
    MsgBox "Proposta salva com sucesso!", vbInformation
    LimparFormulario
End Sub

Private Sub btnAlterarProposta_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long, j As Long
    Dim nrProposta As String
    Dim linhasParaExcluir As Collection
    
    ' Verificações de segurança
    If Not modoEdicao Then
        MsgBox "Não há proposta carregada para alteração.", vbExclamation
        Exit Sub
    End If
    
    If Not propostaAlterada Then
        MsgBox "Nenhuma alteração foi feita na proposta.", vbInformation
        Exit Sub
    End If
    
    If lstProdutosDaProposta.ListCount <= 1 Then
        MsgBox "Adicione pelo menos um item à proposta antes de alterá-la.", vbExclamation
        Exit Sub
    End If
    
    ' Confirmar a alteração
    If MsgBox("Deseja realmente alterar esta proposta?", vbQuestion + vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    nrProposta = nrPropostaOriginal ' Usar o número original da proposta
    
    ' Encontrar todas as linhas com a proposta atual
    Set linhasParaExcluir = New Collection
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Primeiro, identifica todas as linhas que contêm o número da proposta
    For i = ultimaLinha To 2 Step -1 ' Começa de baixo para cima para não afetar os índices
        If ws.Cells(i, "A").Value = nrProposta Then
            On Error Resume Next
            linhasParaExcluir.Add i
            On Error GoTo 0
        End If
    Next i
    
    ' Excluir as linhas identificadas (em ordem decrescente para não afetar índices)
    For i = 1 To linhasParaExcluir.Count
        ws.Rows(linhasParaExcluir(i)).Delete
    Next i
    
    ' Encontrar a nova última linha após as exclusões
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Salvar os itens atualizados da proposta
    For i = 1 To lstProdutosDaProposta.ListCount - 1
        ws.Cells(ultimaLinha, "A").Value = nrProposta
        ws.Cells(ultimaLinha, "B").Value = txtNomeCliente.Value
        ws.Cells(ultimaLinha, "C").Value = txtCidade.Value
        ws.Cells(ultimaLinha, "D").Value = txtEstado.Value
        ws.Cells(ultimaLinha, "E").Value = txtPessoaContato.Value
        ws.Cells(ultimaLinha, "F").Value = txtEmail.Value
        ws.Cells(ultimaLinha, "G").Value = txtFone.Value
        ws.Cells(ultimaLinha, "H").Value = lstProdutosDaProposta.List(i, 0) ' Item
        ws.Cells(ultimaLinha, "I").Value = lstProdutosDaProposta.List(i, 3) ' Código
        ws.Cells(ultimaLinha, "J").Value = ConverterParaNumero(lstProdutosDaProposta.List(i, 4)) ' Preço
        ws.Cells(ultimaLinha, "K").Value = lstProdutosDaProposta.List(i, 1) ' Quantidade
        ws.Cells(ultimaLinha, "L").Value = txtRefProposta.Value ' Referência da Proposta
        ws.Cells(ultimaLinha, "M").Value = cmbVendedor.Value
        ws.Cells(ultimaLinha, "N").Value = cmbCondPagamento.Value
        ws.Cells(ultimaLinha, "O").Value = txtPrazoEntrega.Value
        ws.Cells(ultimaLinha, "P").Value = cmbFrete.Value
        
        ultimaLinha = ultimaLinha + 1
    Next i
    
    MsgBox "Proposta alterada com sucesso!", vbInformation
    LimparFormulario
End Sub

' ======================
' FUNÇÕES DE SUPORTE
' ======================

Private Sub LimparFormulario()
    ' Limpa e reinicializa o ListBox
    With lstProdutosDaProposta
        .Clear
        .AddItem ""
        .List(0, 0) = "Item"
        .List(0, 1) = "Qtd"
        .List(0, 2) = "Descrição"
        .List(0, 3) = "Código"
        .List(0, 4) = "Preço"
        .List(0, 5) = "Sub Total"
    End With
    
    ' Limpar campos de informações da proposta
    cmbVendedor.Value = ""
    cmbCondPagamento.Value = ""
    txtPrazoEntrega.Value = ""
    cmbFrete.Value = ""
    txtRefProposta.Value = ""
    txtValorTotal.Value = "0,00"
    
    ' Limpar campos de busca de proposta e identificação da proposta atual
    txtNrProposta.Value = ""
    txtNovaProposta.Value = ""
    
    ' Limpar campos do cliente
    txtNomeCliente.Value = ""
    txtCidade.Value = ""
    txtEstado.Value = ""
    txtPessoaContato.Value = ""
    txtFone.Value = ""
    txtEmail.Value = ""
    
    ' Limpar campos do produto atual
    txtItem.Value = "1"
    txtQTD.Value = ""
    txtCodProduto.Value = ""
    txtDescricao.Value = ""
    txtPreco.Value = ""
    
    ' Limpar lista de resultados da busca
    lstBuscaProposta.Clear
    
    ' Resetar o estado do formulário
    modoEdicao = False
    propostaAlterada = False
    nrPropostaOriginal = ""
    
    ' Gerenciar estado dos botões
    btnSalvarNovaProposta.Enabled = False
    btnAlterarProposta.Enabled = False
    btnAdicionarProduto.Enabled = False
    btnBuscarProduto.Enabled = False
    
    ' Verificar habilitação do botão de busca de produto
    CheckEnableBuscarProduto
End Sub

Private Sub AtualizarValorTotal()
    Dim total As Double
    Dim i As Long
    
    total = 0
    ' Começa do 1 pois 0 é o cabeçalho
    For i = 1 To lstProdutosDaProposta.ListCount - 1
        ' Pega o valor do Sub Total (coluna 5) e soma o total
        total = total + CDbl(ConverterParaNumero(lstProdutosDaProposta.List(i, 5)))
    Next i
    
    ' Atualiza o campo txtValorTotal com formatação de moeda
    txtValorTotal.Value = Format(total, "#,##0.00")
End Sub

Private Function ConverterParaNumero(valor As String) As Double
    Dim temp As String
    
    ' Validar entrada
    If Trim(valor) = "" Then
        ConverterParaNumero = 0
        Exit Function
    End If
    
    ' Remover os separadores de milhar (pontos)
    temp = Replace(valor, ".", "")
    ' Substituir a vírgula decimal por um ponto
    temp = Replace(temp, ",", ".")
    
    ' Tentar converter para número
    On Error Resume Next
    ConverterParaNumero = Val(temp)
    If Err.Number <> 0 Then
        ConverterParaNumero = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

' ======================
' DETECTAR ALTERAÇÕES PARA MODO DE EDIÇÃO
' ======================

' Eventos para os campos de texto que afetam a proposta
Private Sub txtRefProposta_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub

Private Sub txtPessoaContato_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub

Private Sub txtEmail_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub

Private Sub txtFone_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub

Private Sub cmbVendedor_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub

Private Sub cmbCondPagamento_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub

Private Sub txtPrazoEntrega_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub

Private Sub cmbFrete_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub

