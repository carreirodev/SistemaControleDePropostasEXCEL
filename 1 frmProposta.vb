' Código para o formulário frmProposta
Private Sub btnBuscaCliente_Click()
    ' Verifica se há texto para busca
    If Trim(txtNomeCliente.Value) = "" Then
        MsgBox "Por favor, digite um nome para busca.", vbExclamation
        Exit Sub
    End If
    
    ' Abre o formulário frmCliente passando o texto de busca
    frmCliente.IniciarBusca txtNomeCliente.Value
End Sub

' Esta sub será chamada pelo frmCliente quando um cliente for selecionado
Public Sub PreencherDadosCliente(nome As String, cidade As String, estado As String)
    txtNomeCliente.Value = nome
    txtCidade.Value = cidade
    txtEstado.Value = estado
End Sub
