' Código para o formulário frmBuscaCliente

Private Sub UserForm_Initialize()
    ' Inicializa o formulário
    ConfigurarListBox
End Sub

Private Sub ConfigurarListBox()
    ' Configura as colunas da ListBox sem adicionar cabeçalho
    With lstResultados
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "44;150;130;100;22"
    End With
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
    End If
End Sub

Private Sub lstResultados_Click()
    If lstResultados.ListIndex >= 0 Then ' Alterado para >= 0 já que não há mais linha de cabeçalho
        PreencherCamposCliente lstResultados.List(lstResultados.ListIndex, 0) ' Passa o ID do cliente selecionado
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
            MsgBox "Informações do cliente alteradas com sucesso.", vbInformation
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
    txtNomeCliente.SetFocus

End Sub

