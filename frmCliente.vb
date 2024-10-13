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




Private Sub btnSalvar_Click()
    If Len(Trim(txtNomeCliente.Value)) = 0 Or Len(Trim(txtPessoaContato.Value)) = 0 Then
        MsgBox "Os campos 'Nome do Cliente' e 'Pessoa de Contato' são obrigatórios.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim novoID As String
    Dim nomeClienteMaiusculo As String
    Dim telefoneFormatado As String

    Set ws = ThisWorkbook.Sheets("CLIENTES")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    nomeClienteMaiusculo = UCase(txtNomeCliente.Value)
    txtID = GerarNovoID(nomeClienteMaiusculo, ultimaLinha)
    telefoneFormatado = "'" & txtTelefone.Value

    ws.Cells(ultimaLinha + 1, 1).Value = txtID
    ws.Cells(ultimaLinha + 1, 2).Value = nomeClienteMaiusculo
    ws.Cells(ultimaLinha + 1, 3).Value = txtPessoaContato.Value
    ws.Cells(ultimaLinha + 1, 4).Value = txtEndereco.Value
    ws.Cells(ultimaLinha + 1, 5).Value = txtCidade.Value
    ws.Cells(ultimaLinha + 1, 6).Value = UCase(txtEstado.Value)
    ws.Cells(ultimaLinha + 1, 7).Value = telefoneFormatado
    ws.Cells(ultimaLinha + 1, 8).Value = LCase(txtEmail.Value)

    MsgBox "Cliente cadastrado com sucesso!" & vbNewLine & "ID: " & txtID, vbInformation
    LimparFormulario
End Sub



Private Function GerarNovoID(nomeCliente As String, ultimaLinha As Long) As String
    Dim prefixo As String
    Dim numero As Long
    Dim idExistente As Boolean
    Dim novoID As String
    Dim rng As Range
    Dim caracteres As String
    Dim i As Integer

    caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    numero = ultimaLinha

    Do
        If Len(nomeCliente) >= 2 Then
            prefixo = UCase(Left(nomeCliente, 2))
        Else
            prefixo = ""
        End If

        If prefixo = "" Or idExistente Then
            prefixo = ""
            For i = 1 To 2
                prefixo = prefixo & Mid(caracteres, Int((Len(caracteres) * Rnd) + 1), 1)
            Next i
        End If

        novoID = prefixo & Format(numero, "00000")
        Set rng = ThisWorkbook.Sheets("CLIENTES").Columns("A").Find(What:=novoID, LookIn:=xlValues, LookAt:=xlWhole)
        idExistente = Not rng Is Nothing
    Loop While idExistente

    GerarNovoID = novoID
End Function


Private Sub ValidarCamposObrigatorios()
    btnSalvar.Enabled = (Len(Trim(txtNomeCliente.Value)) > 0 And Len(Trim(txtPessoaContato.Value)) > 0)
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

Private Sub txtNomeCliente_Change()
    txtNomeCliente.Text = UCase(txtNomeCliente.Text)
    txtNomeCliente.SelStart = Len(txtNomeCliente.Text)
    ' Remova a chamada para ValidarCamposObrigatorios
    VerificarAlteracoes
End Sub

Private Sub txtPessoaContato_Change()
    txtPessoaContato.Text = FormatarPrimeiraLetraMaiuscula(txtPessoaContato)
    ' Remova a chamada para ValidarCamposObrigatorios
    VerificarAlteracoes
End Sub

Private Sub txtEndereco_Change()
    txtEndereco.Text = FormatarPrimeiraLetraMaiuscula(txtEndereco)
    VerificarAlteracoes
End Sub

Private Sub txtCidade_Change()
    txtCidade.Text = FormatarPrimeiraLetraMaiuscula(txtCidade)
    VerificarAlteracoes
End Sub

Private Sub txtEstado_Change()
    txtEstado.Text = UCase(txtEstado.Text)
    txtEstado.SelStart = Len(txtEstado.Text)
    VerificarAlteracoes
End Sub

Private Sub txtEmail_Change()
    Dim cursorPos As Long
    cursorPos = txtEmail.SelStart
    txtEmail.Text = LCase(txtEmail.Text)
    txtEmail.SelStart = cursorPos
    VerificarAlteracoes
End Sub

Private Sub txtTelefone_Change()
    VerificarAlteracoes
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


Private Sub btnFechar_Click()
    ' Fecha o formulário sem nenhuma ação
    Unload Me
End Sub