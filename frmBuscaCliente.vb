' Código para o formulário frmBuscaCliente

' Variáveis globais para armazenar os valores originais
Dim originalNomeCliente As String
Dim originalPessoaContato As String
Dim originalEndereco As String
Dim originalCidade As String
Dim originalEstado As String
Dim originalTelefone As String
Dim originalEmail As String

' Variável para rastrear se um registro foi selecionado
Dim registroSelecionado As Boolean

Private Sub UserForm_Initialize()
    ' Inicializa o formulário
    ConfigurarListBox
    btnAlterar.Enabled = False ' Desabilita o botão ALTERAR inicialmente
    btnApagar.Enabled = False ' Desabilita o botão APAGAR inicialmente
    registroSelecionado = False
End Sub

Private Sub ConfigurarListBox()
    ' Configura as colunas da ListBox sem adicionar cabeçalho
    With lstResultados
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "49;170;155;110;24"
    End With
End Sub

Private Sub txtNomeCliente_Change()
    txtNomeCliente.Value = UCase(txtNomeCliente.Value)
    VerificarAlteracoes
End Sub

Private Sub txtEstado_Change()
    txtEstado.Value = UCase(txtEstado.Value)
    VerificarAlteracoes
End Sub

Private Sub txtPessoaContato_Change()
    txtPessoaContato.Value = FormatarPrimeiraLetraMaiuscula(txtPessoaContato)
    VerificarAlteracoes
End Sub

Private Sub txtEndereco_Change()
    txtEndereco.Value = FormatarPrimeiraLetraMaiuscula(txtEndereco)
    VerificarAlteracoes
End Sub

Private Sub txtCidade_Change()
    txtCidade.Value = FormatarPrimeiraLetraMaiuscula(txtCidade)
    VerificarAlteracoes
End Sub

Private Sub txtEmail_Change()
    txtEmail.Value = LCase(txtEmail.Value)
    VerificarAlteracoes
End Sub

Private Sub txtTelefone_Change()
    VerificarAlteracoes
End Sub

Private Sub VerificarAlteracoes()
    ' Habilita o botão ALTERAR somente se um registro foi selecionado e algum campo (exceto ID) for alterado
    If registroSelecionado And _
       (txtNomeCliente.Value <> originalNomeCliente Or _
       txtPessoaContato.Value <> originalPessoaContato Or _
       txtEndereco.Value <> originalEndereco Or _
       txtCidade.Value <> originalCidade Or _
       txtEstado.Value <> originalEstado Or _
       txtTelefone.Value <> originalTelefone Or _
       txtEmail.Value <> originalEmail) Then
        btnAlterar.Enabled = True
    Else
        btnAlterar.Enabled = False
    End If
End Sub

Private Sub btnBuscar_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim encontrou As Boolean
    
    Set ws = ThisWorkbook.Sheets("CLIENTES")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Limpa a ListBox antes de uma nova busca
    lstResultados.Clear
    
    encontrou = False
    
    For i = 2 To ultimaLinha ' Assumindo que a linha 1 é o cabeçalho na planilha
        If (InStr(1, ws.Cells(i, 1).Value, txtID.Value, vbTextCompare) > 0 And Len(txtID.Value) > 0) Or _
           (InStr(1, ws.Cells(i, 2).Value, txtNomeCliente.Value, vbTextCompare) > 0 And Len(txtNomeCliente.Value) > 0) Then
            
            lstResultados.AddItem
            lstResultados.List(lstResultados.ListCount - 1, 0) = ws.Cells(i, 1).Value ' ID
            lstResultados.List(lstResultados.ListCount - 1, 1) = ws.Cells(i, 2).Value ' Nome do Cliente
            lstResultados.List(lstResultados.ListCount - 1, 2) = ws.Cells(i, 3).Value ' Pessoa de Contato
            lstResultados.List(lstResultados.ListCount - 1, 3) = ws.Cells(i, 5).Value ' Cidade
            lstResultados.List(lstResultados.ListCount - 1, 4) = ws.Cells(i, 6).Value ' Estado
            
            encontrou = True
        End If
    Next i
    
    If Not encontrou Then
        MsgBox "Nenhum cliente encontrado com os critérios fornecidos.", vbInformation
    Else
        btnApagar.Enabled = False ' Desabilita o botão APAGAR após a busca, até que um registro seja selecionado
    End If

End Sub

Private Sub lstResultados_Click()
    If lstResultados.ListIndex >= 0 Then
        PreencherCamposCliente lstResultados.List(lstResultados.ListIndex, 0)
        registroSelecionado = True
        btnApagar.Enabled = True ' Habilita o botão APAGAR
    End If
End Sub

Private Sub PreencherCamposCliente(clienteID As String)
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("CLIENTES")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If ws.Cells(i, 1).Value = clienteID Then
            txtID.Value = ws.Cells(i, 1).Value
            txtNomeCliente.Value = ws.Cells(i, 2).Value
            txtPessoaContato.Value = ws.Cells(i, 3).Value
            txtEndereco.Value = ws.Cells(i, 4).Value
            txtCidade.Value = ws.Cells(i, 5).Value
            txtEstado.Value = ws.Cells(i, 6).Value
            txtTelefone.Value = ws.Cells(i, 7).Value
            txtEmail.Value = ws.Cells(i, 8).Value
            
            ' Armazena os valores originais
            originalNomeCliente = txtNomeCliente.Value
            originalPessoaContato = txtPessoaContato.Value
            originalEndereco = txtEndereco.Value
            originalCidade = txtCidade.Value
            originalEstado = txtEstado.Value
            originalTelefone = txtTelefone.Value
            originalEmail = txtEmail.Value
            
            Exit For
        End If
    Next i
End Sub

Private Sub btnAlterar_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim clienteID As String
    Dim telefoneFormatado As String
    
    ' Verifica se um cliente foi selecionado
    If lstResultados.ListIndex < 0 Then
        MsgBox "Selecione um cliente para alterar.", vbExclamation
        Exit Sub
    End If
    
    ' Obtém o ID do cliente selecionado
    clienteID = lstResultados.List(lstResultados.ListIndex, 0)
    
    Set ws = ThisWorkbook.Sheets("CLIENTES")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Formata o telefone com apóstrofo no início
    telefoneFormatado = "'" & txtTelefone.Value
    
    ' Procura o cliente na planilha e atualiza os dados
    For i = 2 To ultimaLinha
        If ws.Cells(i, 1).Value = clienteID Then
            ' Atualiza apenas os campos permitidos
            ws.Cells(i, 2).Value = UCase(txtNomeCliente.Value) ' Nome do Cliente em maiúsculas
            ws.Cells(i, 3).Value = FormatarPrimeiraLetraMaiuscula(txtPessoaContato) ' Pessoa de Contato
            ws.Cells(i, 4).Value = FormatarPrimeiraLetraMaiuscula(txtEndereco) ' Endereço
            ws.Cells(i, 5).Value = FormatarPrimeiraLetraMaiuscula(txtCidade) ' Cidade
            ws.Cells(i, 6).Value = UCase(txtEstado.Value) ' Estado em maiúsculas
            ws.Cells(i, 7).Value = telefoneFormatado ' Telefone formatado
            ws.Cells(i, 8).Value = LCase(txtEmail.Value) ' Email em minúsculas
            btnAlterar.Enabled = False 
            MsgBox "Informações do cliente alteradas com sucesso.", vbInformation
            LimparFormulario
            Exit Sub
            
        End If

    Next i

    MsgBox "Erro ao alterar o cliente. Cliente não encontrado.", vbCritical
End Sub

Private Function FormatarPrimeiraLetraMaiuscula(txt As MSForms.TextBox) As String
    Dim texto As String
    Dim palavras() As String
    Dim i As Integer
    Dim novoTexto As String
    
    texto = txt.Text
    palavras = Split(texto, " ")
    
    For i = 0 To UBound(palavras)
        If Len(palavras(i)) > 0 Then
            palavras(i) = UCase(Left(palavras(i), 1)) & LCase(Mid(palavras(i), 2))
        End If
    Next i
    
    novoTexto = Join(palavras, " ")
    FormatarPrimeiraLetraMaiuscula = novoTexto
End Function

Private Sub btnFechar_Click()
    ' Fecha o formulário sem nenhuma ação
    Unload Me
End Sub

Private Sub btnLimpar_Click()
    ' Limpa todos os campos do formulário
    LimparFormulario
    btnAlterar.Enabled = False ' Desabilita o botão ALTERAR após limpar
    registroSelecionado = False ' Reseta a seleção de registro
End Sub

Private Sub LimparFormulario()
    ' Limpa todos os campos do formulário
    txtID.Value = ""
    txtNomeCliente.Value = ""
    txtPessoaContato.Value = ""
    txtEndereco.Value = ""
    txtCidade.Value = ""
    txtEstado.Value = ""
    txtTelefone.Value = ""
    txtEmail.Value = ""
    lstResultados.Clear
    
    ' Coloca o foco no primeiro campo
    btnAlterar.Enabled = False 

    txtNomeCliente.SetFocus
End Sub



Private Sub btnApagar_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim clienteID As String
    Dim resposta As VbMsgBoxResult
    
    ' Verifica se um cliente foi selecionado
    If lstResultados.ListIndex < 0 Then
        MsgBox "Selecione um cliente para apagar.", vbExclamation
        Exit Sub
    End If
    
    ' Obtém o ID do cliente selecionado
    clienteID = lstResultados.List(lstResultados.ListIndex, 0)
    
    ' Pede confirmação ao usuário
    resposta = MsgBox("Tem certeza de que deseja apagar o cliente selecionado?", vbYesNo + vbQuestion, "Confirmação")
    If resposta = vbNo Then Exit Sub
    
    Set ws = ThisWorkbook.Sheets("CLIENTES")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Procura o cliente na planilha e apaga a linha correspondente
    For i = 2 To ultimaLinha
        If ws.Cells(i, 1).Value = clienteID Then
            ws.Rows(i).Delete
            MsgBox "Cliente apagado com sucesso.", vbInformation
            LimparFormulario
            btnAlterar.Enabled = False  
            Exit Sub
        End If
    Next i
    
    MsgBox "Erro ao apagar o cliente. Cliente não encontrado.", vbCritical
End Sub



'VERSAO 3.5


Private Sub btnCriarNovaProposta_Click()
    Dim wsPropostas As Worksheet
    Dim ultimaLinhaProposta As Long
    Dim numeroProposta As Long
    Dim clienteID As String
    Dim nomeCliente As String
    
    ' Verifica se um cliente foi selecionado
    If lstResultados.ListIndex < 0 Then
        MsgBox "Selecione um cliente para criar uma nova proposta.", vbExclamation
        Exit Sub
    End If
    
    ' Obtém o ID do cliente selecionado
    clienteID = lstResultados.List(lstResultados.ListIndex, 0)
    
    ' Obtém o nome do cliente selecionado
    nomeCliente = lstResultados.List(lstResultados.ListIndex, 1)
    
    ' Define a planilha de propostas
    Set wsPropostas = ThisWorkbook.Sheets("ListaDePropostas")
    
    ' Gera um número de proposta único
    numeroProposta = GerarNumeroPropostaUnico(wsPropostas)
    
    ' Adiciona um novo registro na planilha "ListaDePropostas"
    With wsPropostas
        ultimaLinhaProposta = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
        .Cells(ultimaLinhaProposta, 1).Value = Format(numeroProposta, "0000") ' Número da proposta com 4 dígitos
        .Cells(ultimaLinhaProposta, 2).Value = clienteID ' ID do cliente
    End With
    
    ' Exibe uma mensagem ao usuário com o número da proposta e o nome do cliente
    MsgBox "Nova proposta criada para o cliente: " & vbCrLf & _
            nomeCliente & vbCrLf & _
           "Número da Proposta: " & Format(numeroProposta, "0000"), vbInformation
    Unload Me
End Sub





Private Function GerarNumeroPropostaUnico(ws As Worksheet) As Long
    Dim ultimaLinha As Long
    Dim i As Long
    Dim maiorNumero As Long
    Dim numeroAtual As Long
    
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    maiorNumero = 0
    
    For i = 2 To ultimaLinha ' Assume que a linha 1 é o cabeçalho
        numeroAtual = Val(ws.Cells(i, 1).Value)
        If numeroAtual > maiorNumero Then
            maiorNumero = numeroAtual
        End If
    Next i
    
    GerarNumeroPropostaUnico = maiorNumero + 1
End Function
