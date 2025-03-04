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
    btnApagarProposta.Enabled = False  ' Botão de apagar inicialmente desabilitado
    btnImprimir.Enabled = False  ' Botão de imprimir inicialmente desabilitado
    
    ' Inicializa o ListBox
    With lstProdutosDaProposta
        .Clear
        .ColumnCount = 6
        .ColumnWidths = "30;35;290;70;70;100"
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

'==========================================
' SUBROTINA 1: Preencher ComboBoxes
'==========================================
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
    
    ' Preencher Prazos de Entrega usando a tabela nomeada "PrazoEntrega"
    With cmbPrazoEntrega
        .Clear
        .List = ws.ListObjects("PrazoEntrega").DataBodyRange.Value
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
                         cmbPrazoEntrega.Value <> "" And _
                         cmbFrete.Value <> "")
    
    ' Verifica se há pelo menos um item na lista (além do cabeçalho)
    temItens = (lstProdutosDaProposta.ListCount > 1)
    
    ' Lógica para habilitar/desabilitar os botões baseada no modo
    If modoEdicao Then
        ' Modo de edição de proposta existente
        btnSalvarNovaProposta.Enabled = False ' Sempre desabilitado em modo edição
        
        ' Botão Alterar só fica habilitado se houver itens, campos preenchidos e alterações
        btnAlterarProposta.Enabled = camposPreenchidos And temItens And propostaAlterada
        
        ' Botão Apagar fica habilitado em modo de edição
        btnApagarProposta.Enabled = True
        
        ' Botão Imprimir fica habilitado em modo de edição
        btnImprimir.Enabled = True
    Else
        ' Modo de nova proposta
        btnSalvarNovaProposta.Enabled = camposPreenchidos And temItens
        btnAlterarProposta.Enabled = False ' Sempre desabilitado em modo criação
        btnApagarProposta.Enabled = False ' Sempre desabilitado em modo criação
        
        ' Botão Imprimir é habilitado se tem proposta salva (txtNovaProposta não estiver vazio)
        btnImprimir.Enabled = (Trim(txtNovaProposta.Value) <> "")
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
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
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
    Dim buscarPorCliente As Boolean
    Dim mensagemNaoEncontrado As String
    
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Determinar o critério de busca
    buscarPorCliente = (Trim(txtNrProposta.Value) = "" And Trim(txtNomeCliente.Value) <> "")
    
    If buscarPorCliente Then
        ' Busca pelo nome do cliente
        criterio = LCase(Trim(txtNomeCliente.Value))
        mensagemNaoEncontrado = "Nenhuma proposta encontrada para o cliente " & txtNomeCliente.Value & "."
    Else
        ' Busca pelo número da proposta (comportamento original)
        criterio = LCase(Trim(txtNrProposta.Value))
        mensagemNaoEncontrado = "Nenhuma proposta encontrada."
    End If
    
    Set propostasEncontradas = New Collection
    
    ' Limpar o ListBox
    lstBuscaProposta.Clear
    
    ' Adicionar cabeçalho
    lstBuscaProposta.AddItem
    lstBuscaProposta.List(0, 0) = "Nr da Proposta"
    lstBuscaProposta.List(0, 1) = "Nome do Cliente"
    
    ' Buscar propostas
    For i = 2 To ultimaLinha
        ' Verifica se a linha atual corresponde ao critério de busca
        If (buscarPorCliente And InStr(1, LCase(ws.Cells(i, "B").Value), criterio) > 0) Or _
           (Not buscarPorCliente And InStr(1, LCase(ws.Cells(i, "A").Value), criterio) > 0) Then
            ' Tenta adicionar à coleção (ignorando duplicatas)
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
        MsgBox mensagemNaoEncontrado, vbInformation
    End If
End Sub

' NOVA SUB-ROTINA: Limpar o formulário preservando a lista de busca
Private Sub LimparFormularioPreservandoLista()
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
    cmbPrazoEntrega.Value = ""
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
    
    ' Resetar o estado do formulário
    modoEdicao = False
    propostaAlterada = False
    nrPropostaOriginal = ""
    
    ' Gerenciar estado dos botões
    btnSalvarNovaProposta.Enabled = False
    btnAlterarProposta.Enabled = False
    btnAdicionarProduto.Enabled = False
    btnBuscarProduto.Enabled = False
    btnApagarProposta.Enabled = False  ' Desabilitar o botão Apagar
    btnImprimir.Enabled = False  ' Desabilitar o botão Imprimir
    
    ' Verificar habilitação do botão de busca de produto
    CheckEnableBuscarProduto
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
    
    ' Limpar formulário antes de carregar nova proposta, PRESERVANDO a lista de busca
    LimparFormularioPreservandoLista
    
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
    txtNovaProposta.Value = nrProposta  ' Preencher também txtNovaProposta
    txtNomeCliente.Value = ws.Cells(linha, "B").Value
    txtCidade.Value = ws.Cells(linha, "C").Value
    txtEstado.Value = ws.Cells(linha, "D").Value
    txtPessoaContato.Value = ws.Cells(linha, "E").Value
    txtFone.Value = ws.Cells(linha, "G").Value
    txtEmail.Value = ws.Cells(linha, "F").Value
    txtRefProposta.Value = ws.Cells(linha, "L").Value
    cmbVendedor.Value = ws.Cells(linha, "M").Value
    cmbCondPagamento.Value = ws.Cells(linha, "N").Value
    cmbPrazoEntrega.Value = ws.Cells(linha, "O").Value
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
    
    ' Habilitar explicitamente o botão de impressão
    btnImprimir.Enabled = True
End Sub

Private Sub PreencherListBoxItens(nrProposta As String)
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim wsPrecos As Worksheet
    Dim descricao As String
    
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    Set wsPrecos = ThisWorkbook.Worksheets("ListaDePrecos")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
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
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    
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
        ws.Cells(ultimaLinha, "O").Value = cmbPrazoEntrega.Value
        ws.Cells(ultimaLinha, "P").Value = cmbFrete.Value
        
        ultimaLinha = ultimaLinha + 1
    Next i
    
    ' Habilitar o botão de impressão antes de limpar o formulário
    btnImprimir.Enabled = True
    
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
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
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
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    
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
        ws.Cells(ultimaLinha, "O").Value = cmbPrazoEntrega.Value
        ws.Cells(ultimaLinha, "P").Value = cmbFrete.Value
        
        ultimaLinha = ultimaLinha + 1
    Next i
    
    MsgBox "Proposta alterada com sucesso!", vbInformation
    LimparFormulario
End Sub

' ======================
' NOVA ROTINA PARA APAGAR PROPOSTA
' ======================

Private Sub btnApagarProposta_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim nrProposta As String
    Dim linhasParaExcluir As Collection
    
    ' Verificar se há uma proposta carregada
    If Not modoEdicao Then
        MsgBox "Não há proposta carregada para exclusão. Por favor, selecione uma proposta primeiro.", vbExclamation
        Exit Sub
    End If
    
    ' Confirmar a exclusão
    If MsgBox("Deseja realmente APAGAR esta proposta? Esta ação não pode ser desfeita.", vbQuestion + vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
        Exit Sub
    End If
    
    ' Segunda confirmação para evitar exclusões acidentais
    If MsgBox("CONFIRMAÇÃO: Esta proposta será PERMANENTEMENTE EXCLUÍDA. Deseja continuar?", vbQuestion + vbYesNo + vbDefaultButton2 + vbCritical) = vbNo Then
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    nrProposta = nrPropostaOriginal ' Usar o número original da proposta
    
    ' Encontrar todas as linhas com a proposta atual
    Set linhasParaExcluir = New Collection
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    ' Primeiro, identifica todas as linhas que contêm o número da proposta
    For i = ultimaLinha To 2 Step -1 ' Começa de baixo para cima para não afetar os índices
        If ws.Cells(i, "A").Value = nrProposta Then
            On Error Resume Next
            linhasParaExcluir.Add i
            On Error GoTo 0
        End If
    Next i
    
    ' Verificar se foram encontradas linhas para excluir
    If linhasParaExcluir.Count = 0 Then
        MsgBox "Não foram encontrados registros para a proposta " & nrProposta & ".", vbExclamation
        Exit Sub
    End If
    
    ' Excluir as linhas identificadas (em ordem decrescente para não afetar índices)
    For i = 1 To linhasParaExcluir.Count
        ws.Rows(linhasParaExcluir(i)).Delete
    Next i
    
    MsgBox "Proposta " & nrProposta & " excluída com sucesso!", vbInformation
    
    ' Limpar o formulário
    LimparFormulario
End Sub

' ======================
' NOVA ROTINA PARA IMPRIMIR PROPOSTA
' ======================
Private Sub btnImprimir_Click()
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim wsPrecos As Worksheet
    Dim nomePlanilha As String
    Dim dataFormatada As String
    Dim mes As String
    Dim i As Long
    Dim linhaAtual As Long
    Dim codigoProduto As String
    Dim descricaoCompleta As String
    Dim descricaoBase As String
    Dim marca As String
    Dim anvisa As String
    Dim simpro As String
    Dim linhaProduto As Long
    Dim ultimaLinha As Long
    Dim linhaTotal As Long
    Dim formulaSoma As String
    
    ' Verificar se existe um número de proposta válido
    If Trim(txtNovaProposta.Value) = "" Then
        MsgBox "Não há proposta selecionada para impressão.", vbExclamation
        Exit Sub
    End If
    
    ' Obter o nome para a nova planilha (usando o número da proposta)
    nomePlanilha = txtNovaProposta.Value
    
    ' Verificar se já existe uma planilha com esse nome
    On Error Resume Next
    Set wsDestino = ThisWorkbook.Worksheets(nomePlanilha)
    On Error GoTo 0
    
    ' Se a planilha já existir, perguntar se deseja substituí-la
    If Not wsDestino Is Nothing Then
        If MsgBox("Já existe uma planilha com este nome. Deseja substituí-la?", vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        Else
            Application.DisplayAlerts = False
            ThisWorkbook.Worksheets(nomePlanilha).Delete
            Application.DisplayAlerts = True
        End If
    End If
    
    ' Referenciar a planilha de modelo e a planilha de preços
    Set wsOrigem = ThisWorkbook.Worksheets("IMPRESSAO")
    Set wsPrecos = ThisWorkbook.Worksheets("ListaDePrecos")
    
    ' Criar uma cópia da planilha de modelo com o nome da proposta
    wsOrigem.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set wsDestino = ThisWorkbook.Worksheets(ThisWorkbook.Sheets.Count)
    wsDestino.Name = nomePlanilha
    
    ' Formatar a data atual no formato "São Paulo, DD de MMMM de YYYY"
    ' Obter o mês por extenso em português
    Select Case Month(Date)
        Case 1: mes = "janeiro"
        Case 2: mes = "fevereiro"
        Case 3: mes = "março"
        Case 4: mes = "abril"
        Case 5: mes = "maio"
        Case 6: mes = "junho"
        Case 7: mes = "julho"
        Case 8: mes = "agosto"
        Case 9: mes = "setembro"
        Case 10: mes = "outubro"
        Case 11: mes = "novembro"
        Case 12: mes = "dezembro"
    End Select
    
    dataFormatada = "São Paulo, " & Day(Date) & " de " & mes & " de " & Year(Date)
    
    ' Inserir a data formatada na célula L5 e alinhar à direita
    With wsDestino.Range("L5")
        .Value = dataFormatada
        .HorizontalAlignment = xlRight
    End With
    
    ' Preencher as informações adicionais na planilha
    wsDestino.Range("B6").Value = txtNovaProposta.Value        ' NÚMERO DA PROPOSTA
    wsDestino.Range("A8").Value = txtNomeCliente.Value         ' NOME DO CLIENTE
    wsDestino.Range("J9").Value = txtRefProposta.Value         ' REFERÊNCIA DA PROPOSTA
    wsDestino.Range("B9").Value = txtPessoaContato.Value       ' NOME DE CONTATO DO CLIENTE
    wsDestino.Range("B10").Value = "'" & txtFone.Value         ' TELEFONE DO CLIENTE (com apóstrofo na frente para forçar formato texto)
    wsDestino.Range("D10").Value = txtEmail.Value              ' EMAIL DO CLIENTE
    
    ' Encontrar a última linha na planilha ListaDePrecos
    ultimaLinha = wsPrecos.Cells(wsPrecos.Rows.Count, "A").End(xlUp).row
    
    ' Agora vamos adicionar os itens da proposta a partir da linha 14
    linhaAtual = 14
    
    ' Começamos do índice 1 porque o índice 0 é o cabeçalho
    For i = 1 To lstProdutosDaProposta.ListCount - 1
        ' Se não for o primeiro item, precisamos inserir uma nova linha copiando a formatação da linha 14
        If i > 1 Then
            ' Copia a linha 14 (que tem a formatação correta) e insere abaixo da linha atual
            wsDestino.Rows(14).Copy
            wsDestino.Rows(linhaAtual + 1).Insert Shift:=xlDown
            linhaAtual = linhaAtual + 1
        End If
        
        ' Definir altura da linha para 94 pixels (aproximadamente 70,5 pontos)
        wsDestino.Rows(linhaAtual).RowHeight = 70.5
        
        ' Obter o código do produto para este item
        codigoProduto = lstProdutosDaProposta.List(i, 3)
        
        ' Buscar na planilha ListaDePrecos as informações adicionais
        descricaoBase = ""
        marca = ""
        anvisa = ""
        simpro = ""
        
        ' Procurar o código na planilha ListaDePrecos
        For linhaProduto = 2 To ultimaLinha
            If wsPrecos.Cells(linhaProduto, "A").Value = codigoProduto Then
                ' Encontrou o produto, obter as informações
                descricaoBase = wsPrecos.Cells(linhaProduto, "B").Value
                anvisa = wsPrecos.Cells(linhaProduto, "D").Value
                simpro = wsPrecos.Cells(linhaProduto, "H").Value
                marca = wsPrecos.Cells(linhaProduto, "J").Value
                Exit For
            End If
        Next linhaProduto
        
        ' Montar a descrição completa no formato solicitado
        descricaoCompleta = descricaoBase & vbCrLf & vbCrLf & _
                           "Marca: " & marca & vbCrLf & _
                           "ANVISA: " & anvisa & vbCrLf & _
                           "SIMPRO: " & simpro
        
        ' Preencher os dados do item na linha atual
        wsDestino.Cells(linhaAtual, "A").Value = lstProdutosDaProposta.List(i, 0)  ' ITEM
        wsDestino.Cells(linhaAtual, "B").Value = lstProdutosDaProposta.List(i, 1)  ' QUANTIDADE
        wsDestino.Cells(linhaAtual, "C").Value = codigoProduto                     ' CÓDIGO DO PRODUTO
        
        ' Ajustar a célula para quebra de texto e preencher com a descrição completa
        With wsDestino.Cells(linhaAtual, "D")
            .Value = descricaoCompleta
            .WrapText = True
        End With
        
        ' Formatamos o valor unitário como moeda
        wsDestino.Cells(linhaAtual, "K").Value = ConverterParaNumero(lstProdutosDaProposta.List(i, 4))  ' VALOR UNITÁRIO
        
        ' Se necessário, atualize a fórmula do valor total na coluna L para refletir a linha atual
        wsDestino.Cells(linhaAtual, "L").Formula = "=B" & linhaAtual & "*K" & linhaAtual
    Next i
    
    ' Adicionar Condição de Pagamento, Prazo de Entrega e Frete
    ' 3 linhas abaixo do último item (linhaAtual)
    wsDestino.Cells(linhaAtual + 4, "E").Value = cmbCondPagamento.Value  ' Condição de Pagamento
    wsDestino.Cells(linhaAtual + 5, "E").Value = cmbPrazoEntrega.Value   ' Prazo de Entrega
    wsDestino.Cells(linhaAtual + 6, "E").Value = cmbFrete.Value          ' Frete
    
    ' Calcular a linha do Total (duas linhas abaixo do último item)
    linhaTotal = linhaAtual + 2
    
    ' Criar a fórmula de soma para somar todos os valores da coluna L dos itens
    formulaSoma = "=SUM(L14:L" & linhaAtual & ")"
    
    ' Aplicar a fórmula de soma na coluna J da linha total
    With wsDestino.Cells(linhaTotal, "J")
        .Formula = formulaSoma
        .Font.Bold = True
        .NumberFormat = "#,##0.00"  ' Formato moeda
    End With
    
    ' Adicionar informações do vendedor 9 linhas abaixo do último item
    Dim wsVendedores As Worksheet
    Dim vendedorNome As String
    Dim vendedorCargo As String
    Dim vendedorFone As String
    Dim vendedorEmail As String
    Dim tblVendedor As ListObject
    Dim vLinhaVendedor As Long
    
    ' Referenciar a planilha com as informações dos vendedores
    Set wsVendedores = ThisWorkbook.Worksheets("ListasDeEscolha")
    Set tblVendedor = wsVendedores.ListObjects("Vendedor")
    
    ' Obter o nome do vendedor (da proposta atual)
    vendedorNome = cmbVendedor.Value
    
    ' Informações do vendedor - definidas manualmente para garantir a ordem correta
    vendedorCargo = ""
    vendedorFone = ""
    vendedorEmail = ""
    
    ' Procurar o vendedor na tabela "Vendedor"
    On Error Resume Next
    For vLinhaVendedor = 1 To tblVendedor.ListRows.Count
        If tblVendedor.ListRows(vLinhaVendedor).Range.Cells(1, 1).Value = vendedorNome Then
            ' Verificar cada coluna da tabela para identificar corretamente os dados
            ' Independente da ordem das colunas na tabela
            
            ' Iterar pelas colunas da tabela para identificar os campos
            For i = 1 To tblVendedor.HeaderRowRange.Columns.Count
                Dim colHeader As String
                colHeader = tblVendedor.HeaderRowRange.Cells(1, i).Value
                
                ' Identificar colunas por nome de cabeçalho
                Select Case LCase(colHeader)
                    Case "cargo"
                        vendedorCargo = tblVendedor.ListRows(vLinhaVendedor).Range.Cells(1, i).Value
                    Case "fone", "telefone"
                        vendedorFone = tblVendedor.ListRows(vLinhaVendedor).Range.Cells(1, i).Value
                    Case "email", "e-mail"
                        vendedorEmail = tblVendedor.ListRows(vLinhaVendedor).Range.Cells(1, i).Value
                End Select
            Next i
            
            Exit For
        End If
    Next vLinhaVendedor
    On Error GoTo 0
    
    ' Adicionar as informações do vendedor 9 linhas abaixo do último item na COLUNA A
    ' Na ordem especificada: Nome, Cargo, Fone, Email
    With wsDestino
        .Cells(linhaAtual + 9, "A").Value = vendedorNome
        .Cells(linhaAtual + 9, "A").Font.Bold = True
        
        .Cells(linhaAtual + 10, "A").Value = vendedorCargo
        
        .Cells(linhaAtual + 11, "A").Value = vendedorFone
        
        .Cells(linhaAtual + 12, "A").Value = vendedorEmail
    End With
    
    ' Configurar parâmetros de impressão da planilha
    With wsDestino.PageSetup
        ' Orientação e tamanho do papel
        .Orientation = xlPortrait         ' Retrato
        .PaperSize = xlPaperA4            ' Tamanho A4
        
        ' Margens em centímetros convertidos para polegadas (1 polegada = 2,54 cm)
        .TopMargin = Application.CentimetersToPoints(0.5)      ' Margem superior 0,5 cm
        .BottomMargin = Application.CentimetersToPoints(1#)    ' Margem inferior 1,0 cm
        .LeftMargin = Application.CentimetersToPoints(1.3)     ' Margem esquerda 1,3 cm
        .RightMargin = Application.CentimetersToPoints(1.3)    ' Margem direita 1,3 cm
        
        ' Centralizar na página
        .CenterHorizontally = True        ' Centralizar horizontalmente
        
        ' Ajustar para 1 página de largura
        .FitToPagesWide = 1
        .FitToPagesTall = False           ' Altura automática baseada no conteúdo
        
        ' Repetir linhas no topo (1 a 13)
        .PrintTitleRows = "$1:$13"        ' Repetir linhas 1 a 13 em cada página
    End With
    
    ' Ativar a planilha recém-criada
    wsDestino.Activate
    
    ' Confirmar para o usuário
    MsgBox "Planilha de impressão criada com sucesso!", vbInformation
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
    cmbPrazoEntrega.Value = ""
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
    btnApagarProposta.Enabled = False  ' Desabilitar o botão Apagar
    btnImprimir.Enabled = False  ' Desabilitar o botão Imprimir
    
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

Private Sub cmbPrazoEntrega_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub

Private Sub cmbCondPagamento_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub


Private Sub cmbFrete_Change()
    MarcarComoAlterado
    CheckEnableSalvarProposta
End Sub


Private Sub btnLimparCliente_Click()
    ' Limpa apenas os campos relacionados ao cliente
    txtNomeCliente.Value = ""
    txtCidade.Value = ""
    txtEstado.Value = ""
    txtPessoaContato.Value = ""
    txtFone.Value = ""
    txtEmail.Value = ""
    
    ' Atualizar estado dos botões que dependem desses campos
    CheckEnableBuscarProduto
    CheckEnableAdicionarProduto
    CheckEnableSalvarProposta
    
    ' Marcar como alterado se estiver em modo edição
    MarcarComoAlterado
End Sub

Private Sub btnLimpaTudo_Click()
    ' Limpar todos os campos de texto e seleção
    ' Campos de cliente
    txtNomeCliente.Value = ""
    txtCidade.Value = ""
    txtEstado.Value = ""
    txtPessoaContato.Value = ""
    txtFone.Value = ""
    txtEmail.Value = ""
    
    ' Campos de proposta
    txtRefProposta.Value = ""
    cmbPrazoEntrega.Value = ""
    txtNrProposta.Value = ""
    txtNovaProposta.Value = ""
    txtValorTotal.Value = "0,00"
    
    ' Campos de produto
    txtItem.Value = "1"  ' Reiniciar para 1, como no UserForm_Initialize
    txtQTD.Value = ""
    txtCodProduto.Value = ""
    txtDescricao.Value = ""
    txtPreco.Value = ""
    
    ' Limpar ComboBoxes
    cmbVendedor.Value = ""
    cmbCondPagamento.Value = ""
    cmbFrete.Value = ""
    
    ' Limpar e reinicializar ListBox de produtos com cabeçalho
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
    
    ' Limpar lista de resultados da busca
    lstBuscaProposta.Clear
    
    ' Resetar variáveis de controle do formulário
    modoEdicao = False
    propostaAlterada = False
    nrPropostaOriginal = ""
    
    ' Atualizar estado dos botões
    btnBuscarProduto.Enabled = False
    btnAdicionarProduto.Enabled = False
    btnSalvarNovaProposta.Enabled = False
    btnAlterarProposta.Enabled = False
    btnApagarProposta.Enabled = False
    btnImprimir.Enabled = False
End Sub


Private Sub btnFechar_Click()
    ' Fecha o formulário sem realizar nenhuma ação adicional
    Unload Me
End Sub