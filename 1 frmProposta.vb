
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



Private Sub btnNovaProposta_Click()
    ' Verifica se os campos obrigatórios estão preenchidos
    If Trim(txtNomeCliente.Value) = "" Or Trim(txtCidade.Value) = "" Or Trim(txtEstado.Value) = "" Then
        MsgBox "Os campos Nome do Cliente, Cidade e Estado são obrigatórios!", vbExclamation
        Exit Sub
    End If

    ' Obtém a data atual no formato yyyy-mm-dd
    Dim dataAtual As String
    dataAtual = Format(Date, "yyyy-mm-dd")

    ' Obtém as iniciais do cliente
    Dim iniciais As String
    iniciais = ObterIniciaisCliente(txtNomeCliente.Value)

    ' Obtém e incrementa o número sequencial
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("BancoDePropostas")
    Dim numeroSequencial As Long
    numeroSequencial = ws.Range("U1").Value

    ' Formata o número sequencial com 5 dígitos (00001)
    Dim numeroFormatado As String
    numeroFormatado = Format(numeroSequencial, "00000")

    ' Monta o número da proposta
    Dim numeroProposta As String
    numeroProposta = dataAtual & "_" & iniciais & "_" & numeroFormatado

    ' Incrementa o número sequencial
    ws.Range("U1").Value = numeroSequencial + 1

    ' Preenche o campo txtNovaProposta
    txtNovaProposta.Value = numeroProposta
    
    ' Encontra a próxima linha vazia na coluna A
    Dim ultimaLinha As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Adiciona as informações na planilha
    With ws
        .Cells(ultimaLinha, "A").Value = numeroProposta        ' Número da Proposta
        .Cells(ultimaLinha, "B").Value = txtNomeCliente.Value  ' Nome do Cliente
        .Cells(ultimaLinha, "C").Value = txtCidade.Value       ' Cidade
        .Cells(ultimaLinha, "D").Value = txtEstado.Value       ' Estado
        .Cells(ultimaLinha, "E").Value = txtPessoaContato.Value ' Pessoa de Contato
        .Cells(ultimaLinha, "F").Value = txtFone.Value         ' Telefone
        .Cells(ultimaLinha, "G").Value = txtEmail.Value        ' Email
    End With
    
    ' Desabilita os botões
    btnNovaProposta.Enabled = False
    btnBuscaCliente.Enabled = False
    
    ' Opcional: Mensagem informando que a proposta foi criada
    MsgBox "Proposta " & numeroProposta & " criada com sucesso!", vbInformation
End Sub



Private Function ObterIniciaisCliente(nomeCompleto As String) As String
    Dim palavras() As String
    Dim iniciais As String
    
    ' Remove espaços extras e divide o nome em palavras
    palavras = Split(Trim(nomeCompleto))
    
    ' Se tiver pelo menos duas palavras
    If UBound(palavras) >= 1 Then
        ' Pega a primeira letra de cada uma das duas primeiras palavras
        iniciais = UCase(Left(palavras(0), 1) & Left(palavras(1), 1))
    Else
        ' Se tiver apenas uma palavra, usa as duas primeiras letras
        iniciais = UCase(Left(palavras(0), 2))
    End If
    
    ObterIniciaisCliente = iniciais
End Function


Private Sub btnBuscarProduto_Click()
    Dim ws As Worksheet
    Dim codigo As String
    Dim ultimaLinha As Long
    Dim i As Long
    Dim encontrado As Boolean
    Dim preco As Double
    
    ' Define a planilha "TabelaPrecos"
    Set ws = ThisWorkbook.Worksheets("TabelaPrecos")
    
    ' Obtém o código do produto digitado
    codigo = Trim(txtCodProduto.Value)
    
    ' Verifica se o código foi digitado
    If codigo = "" Then
        MsgBox "Por favor, digite um código de produto.", vbExclamation
        Exit Sub
    End If
    
    ' Encontra a última linha com dados
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Procura o código na coluna A
    encontrado = False
    For i = 2 To ultimaLinha ' Assumindo que a primeira linha é cabeçalho
        If ws.Cells(i, "A").Value = codigo Then
            ' Preenche os campos com as informações encontradas
            txtDescricao.Value = ws.Cells(i, "B").Value
            
            ' Obtém o preço e formata corretamente
            preco = CDbl(ws.Cells(i, "C").Value)
            txtPreco.Value = Format(preco, "#,##0.00") ' <-- Ajuste aqui
            
            encontrado = True
            Exit For
        End If
    Next i
    
    ' Se não encontrou o produto, exibe uma mensagem
    If Not encontrado Then
        MsgBox "Produto não encontrado.", vbExclamation
        ' Limpa os campos
        txtDescricao.Value = ""
        txtPreco.Value = ""
    End If
End Sub

