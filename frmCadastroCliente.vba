' Código para o formulário frmCadastroCliente

Private Sub UserForm_Initialize()
    ' Inicializa o formulário (se necessário, você pode adicionar código aqui)
End Sub

Private Sub btnSalvar_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim novoID As String
    Dim nomeClienteMaiusculo As String
    
    ' Define a planilha "CLIENTES"
    Set ws = ThisWorkbook.Sheets("CLIENTES")
    
    ' Encontra a última linha usada na coluna A
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Converte o nome do cliente para maiúsculas
    nomeClienteMaiusculo = UCase(txtNomeCliente.Value)
    
    ' Gera o novo ID usando o nome em maiúsculas
    novoID = GerarNovoID(nomeClienteMaiusculo, ultimaLinha)
    
    ' Adiciona os dados do novo cliente
    ws.Cells(ultimaLinha + 1, 1).Value = novoID
    ws.Cells(ultimaLinha + 1, 2).Value = nomeClienteMaiusculo  ' Nome em maiúsculas
    ws.Cells(ultimaLinha + 1, 3).Value = txtPessoaContato.Value
    ws.Cells(ultimaLinha + 1, 4).Value = txtEndereco.Value
    ws.Cells(ultimaLinha + 1, 5).Value = txtCidade.Value
    ws.Cells(ultimaLinha + 1, 6).Value = txtEstado.Value
    ws.Cells(ultimaLinha + 1, 7).Value = txtTelefone.Value
    ws.Cells(ultimaLinha + 1, 8).Value = txtEmail.Value
    
    MsgBox "Cliente cadastrado com sucesso!" & vbNewLine & "ID: " & novoID, vbInformation
    
    ' Fecha o formulário
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    ' Fecha o formulário sem salvar
    Unload Me
End Sub

Private Function GerarNovoID(nomeCliente As String, ultimaLinha As Long) As String
    Dim prefixo As String
    Dim numero As Long
    
    ' Pega os dois primeiros caracteres do nome do cliente e transforma em maiúsculas
    prefixo = UCase(Left(nomeCliente, 2))
    
    ' Gera o número sequencial
    numero = ultimaLinha
    
    ' Formata o ID (prefixo + número de 5 dígitos)
    GerarNovoID = prefixo & Format(numero, "00000")
End Function

Private Sub txtNomeCliente_Change()
    txtNomeCliente.Text = UCase(txtNomeCliente.Text)
    txtNomeCliente.SelStart = Len(txtNomeCliente.Text)
End Sub

Private Sub txtPessoaContato_Change()
    FormatarPrimeiraLetraMaiuscula txtPessoaContato
End Sub

Private Sub txtEndereco_Change()
    FormatarPrimeiraLetraMaiuscula txtEndereco
End Sub

Private Sub txtCidade_Change()
    FormatarPrimeiraLetraMaiuscula txtCidade
End Sub

Private Sub txtEmail_Change()
    Dim cursorPos As Long
    cursorPos = txtEmail.SelStart
    txtEmail.Text = LCase(txtEmail.Text)
    txtEmail.SelStart = cursorPos
End Sub

Private Sub FormatarPrimeiraLetraMaiuscula(txt As TextBox)
    Dim texto As String
    Dim palavras() As String
    Dim i As Integer
    Dim novoTexto As String
    Dim cursorPos As Long
    
    cursorPos = txt.SelStart
    texto = txt.Text
    palavras = Split(texto, " ")
    
    For i = 0 To UBound(palavras)
        If Len(palavras(i)) > 0 Then
            palavras(i) = UCase(Left(palavras(i), 1)) & LCase(Mid(palavras(i), 2))
        End If
    Next i
    
    novoTexto = Join(palavras, " ")
    txt.Text = novoTexto
    txt.SelStart = cursorPos
End Sub
