Private Sub UserForm_Initialize()
    btnBuscarProduto.Enabled = False
    btnAdicionarProduto.Enabled = False
    
    ' Inicializa o ListBox
    With lstProdutosDaProposta
        .Clear
        .ColumnCount = 6  ' Aumentado para 6 colunas
        .ColumnWidths = "40;60;320;90;90;120"  ' Adicionada largura para a nova coluna
    End With
    
    ' Adiciona o cabeçalho
    lstProdutosDaProposta.AddItem ""
    lstProdutosDaProposta.List(0, 0) = "Item"
    lstProdutosDaProposta.List(0, 1) = "Qtd"
    lstProdutosDaProposta.List(0, 2) = "Descrição"
    lstProdutosDaProposta.List(0, 3) = "Código"
    lstProdutosDaProposta.List(0, 4) = "Preço"
    lstProdutosDaProposta.List(0, 5) = "Sub Total"  ' Nova coluna
End Sub



Private Sub txtNomeCliente_Change()
    CheckEnableBuscarProduto
End Sub

Private Sub txtCidade_Change()
    CheckEnableBuscarProduto
End Sub

Private Sub txtEstado_Change()
    CheckEnableBuscarProduto
End Sub

Private Sub CheckEnableBuscarProduto()
    btnBuscarProduto.Enabled = (Trim(txtNomeCliente.Value) <> "" And _
                                Trim(txtCidade.Value) <> "" And _
                                Trim(txtEstado.Value) <> "")
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






Private Sub btnAdicionarProduto_Click()
    Dim subTotal As Double
    Dim preco As Double
    Dim quantidade As Double
    
    ' Verifica se é o primeiro item
    If txtNovaProposta.Value = "" Then
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
        .List(.ListCount - 1, 5) = Format(subTotal, "#,##0.00") ' Novo Sub Total
    End With
    
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



Private Sub btnSalvarNovaProposta_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    
    ' Encontrar a última linha preenchida
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Salvar cada item da proposta
    For i = 1 To lstProdutosDaProposta.ListCount - 1 ' -1 para ignorar o cabeçalho
        ws.Cells(ultimaLinha, "A").Value = txtNovaProposta.Value
        ws.Cells(ultimaLinha, "B").Value = txtNomeCliente.Value
        ws.Cells(ultimaLinha, "C").Value = txtCidade.Value
        ws.Cells(ultimaLinha, "D").Value = txtEstado.Value
        ws.Cells(ultimaLinha, "E").Value = txtPessoaContato.Value
        ws.Cells(ultimaLinha, "F").Value = txtFone.Value
        ws.Cells(ultimaLinha, "G").Value = txtEmail.Value
        ws.Cells(ultimaLinha, "H").Value = lstProdutosDaProposta.List(i, 0) ' Item
        ws.Cells(ultimaLinha, "I").Value = lstProdutosDaProposta.List(i, 3) ' Código
        ws.Cells(ultimaLinha, "J").Value = lstProdutosDaProposta.List(i, 4) ' Preço
        ws.Cells(ultimaLinha, "K").Value = lstProdutosDaProposta.List(i, 1) ' Quantidade
        ultimaLinha = ultimaLinha + 1
    Next i
    
    MsgBox "Proposta salva com sucesso!", vbInformation
    
    ' Limpar o formulário
    LimparFormulario
End Sub





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
        .List(0, 5) = "Sub Total"  ' Nova coluna
    End With
    
End Sub

Private Function ConverterParaNumero(valor As String) As Double
    Dim temp As String
    ' Primeiro, remover os separadores de milhar (pontos)
    temp = Replace(valor, ".", "")
    ' Depois, substituir a vírgula decimal por um ponto
    temp = Replace(temp, ",", ".")
    ' Converter para número
    ConverterParaNumero = Val(temp)
End Function
